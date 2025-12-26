using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ZontSpecExtractor
{
    public class TelegramLogger : ILogHandler
    {
        private readonly string _botToken;
        private readonly string _chatId;
        private readonly HttpClient _httpClient;
        private readonly LogLevel _minLevel;
        private readonly Queue<LogEntry> _pendingLogs = new Queue<LogEntry>();
        private readonly object _queueLock = new object();
        private bool _isSending = false;

        public TelegramLogger(string botToken, string chatId, LogLevel minLevel = LogLevel.WARNING)
        {
            _botToken = botToken ?? throw new ArgumentNullException(nameof(botToken));
            _chatId = chatId ?? throw new ArgumentNullException(nameof(chatId));
            _minLevel = minLevel;
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        }

        public void Handle(LogEntry entry)
        {
            if (entry.Level < _minLevel) return;

            lock (_queueLock)
            {
                _pendingLogs.Enqueue(entry);
            }

            // Асинхронная отправка без блокировки
            _ = Task.Run(SendPendingLogsAsync);
        }

        private async Task SendPendingLogsAsync()
        {
            if (_isSending) return;
            _isSending = true;

            try
            {
                while (true)
                {
                    LogEntry? entry = null;
                    lock (_queueLock)
                    {
                        if (_pendingLogs.Count == 0)
                        {
                            _isSending = false;
                            return;
                        }
                        entry = _pendingLogs.Dequeue();
                    }

                    if (entry != null)
                    {
                        await SendLogToTelegramAsync(entry);
                        await Task.Delay(100); // Небольшая задержка между сообщениями
                    }
                }
            }
            catch (Exception ex)
            {
                // Критично: не падаем при ошибке отправки
                System.Diagnostics.Debug.WriteLine($"Telegram send error: {ex.Message}");
            }
            finally
            {
                _isSending = false;
            }
        }

        private async Task SendLogToTelegramAsync(LogEntry entry)
        {
            try
            {
                var message = FormatTelegramMessage(entry);
                var url = $"https://api.telegram.org/bot{_botToken}/sendMessage";

                var payload = new
                {
                    chat_id = _chatId,
                    text = message,
                    parse_mode = "HTML"
                };

                var json = System.Text.Json.JsonSerializer.Serialize(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(url, content);
                
                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"Telegram API error: {errorContent}");
                }
            }
            catch (Exception ex)
            {
                // Молча игнорируем ошибки сети
                System.Diagnostics.Debug.WriteLine($"Telegram send exception: {ex.Message}");
            }
        }

        private string FormatTelegramMessage(LogEntry entry)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"<b>🔔 {entry.Level} Log</b>");
            sb.AppendLine($"<b>Time:</b> {entry.Timestamp:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"<b>Class:</b> <code>{entry.ClassName}</code>");
            sb.AppendLine($"<b>Method:</b> <code>{entry.MethodName}</code>");
            sb.AppendLine($"<b>Thread:</b> {entry.ThreadId}");
            sb.AppendLine($"<b>Memory:</b> {entry.MemoryUsage / 1024 / 1024} MB");
            
            if (entry.Parameters != null && entry.Parameters.Count > 0)
            {
                sb.AppendLine($"<b>Parameters:</b>");
                foreach (var kvp in entry.Parameters)
                {
                    sb.AppendLine($"  • {kvp.Key} = {kvp.Value}");
                }
            }
            
            sb.AppendLine($"<b>Message:</b> {HtmlEncode(entry.Message)}");
            
            if (entry.ReturnValue != null)
            {
                sb.AppendLine($"<b>Return:</b> {HtmlEncode(entry.ReturnValue.ToString())}");
            }
            
            if (entry.Exception != null)
            {
                sb.AppendLine($"<b>❌ Exception:</b>");
                sb.AppendLine($"<code>{HtmlEncode(entry.Exception.GetType().Name)}</code>");
                sb.AppendLine($"<code>{HtmlEncode(entry.Exception.Message)}</code>");
                
                if (!string.IsNullOrEmpty(entry.Exception.StackTrace))
                {
                    var stackTrace = entry.Exception.StackTrace;
                    if (stackTrace.Length > 1000)
                        stackTrace = stackTrace.Substring(0, 1000) + "...";
                    sb.AppendLine($"<pre>{HtmlEncode(stackTrace)}</pre>");
                }
            }

            return sb.ToString();
        }

        public async Task SendFullLogAsync(string logContent)
        {
            try
            {
                // Telegram ограничивает длину сообщения, разбиваем на части
                const int maxLength = 4000;
                var parts = SplitString(logContent, maxLength);

                foreach (var part in parts)
                {
                    var url = $"https://api.telegram.org/bot{_botToken}/sendMessage";
                    var payload = new
                    {
                        chat_id = _chatId,
                        text = $"<pre>{HtmlEncode(part)}</pre>",
                        parse_mode = "HTML"
                    };

                    var json = System.Text.Json.JsonSerializer.Serialize(payload);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    await _httpClient.PostAsync(url, content);
                    await Task.Delay(500); // Задержка между частями
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to send full log: {ex.Message}");
            }
        }

        private List<string> SplitString(string text, int maxLength)
        {
            var parts = new List<string>();
            for (int i = 0; i < text.Length; i += maxLength)
            {
                var length = Math.Min(maxLength, text.Length - i);
                parts.Add(text.Substring(i, length));
            }
            return parts;
        }

        private string HtmlEncode(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            
            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&#39;");
        }

        /// <summary>
        /// Отправляет файл в Telegram
        /// </summary>
        public async Task SendFileAsync(string filePath, string? caption = null)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    System.Diagnostics.Debug.WriteLine($"File not found: {filePath}");
                    return;
                }

                var url = $"https://api.telegram.org/bot{_botToken}/sendDocument";
                
                using (var formData = new MultipartFormDataContent())
                {
                    formData.Add(new StringContent(_chatId), "chat_id");
                    
                    if (!string.IsNullOrEmpty(caption))
                    {
                        formData.Add(new StringContent(caption), "caption");
                    }

                    var fileContent = new ByteArrayContent(File.ReadAllBytes(filePath));
                    fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                    var fileName = Path.GetFileName(filePath) ?? "file";
                    formData.Add(fileContent, "document", fileName);

                    var response = await _httpClient.PostAsync(url, formData);
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"Telegram API error when sending file: {errorContent}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to send file to Telegram: {ex.Message}");
            }
        }

        /// <summary>
        /// Отправляет несколько файлов в Telegram
        /// </summary>
        public async Task SendFilesAsync(List<string> filePaths, string? message = null)
        {
            try
            {
                // Сначала отправляем текстовое сообщение, если есть
                if (!string.IsNullOrEmpty(message))
                {
                    await SendFullLogAsync(message);
                    await Task.Delay(500);
                }

                // Отправляем каждый файл
                foreach (var filePath in filePaths)
                {
                    if (File.Exists(filePath))
                    {
                        var fileName = Path.GetFileName(filePath);
                        await SendFileAsync(filePath, $"📎 {fileName}");
                        await Task.Delay(500); // Задержка между файлами
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to send files to Telegram: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}

