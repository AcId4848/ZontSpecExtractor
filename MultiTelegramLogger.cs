using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ZontSpecExtractor
{
    /// <summary>
    /// Логгер для отправки сообщений нескольким получателям Telegram
    /// </summary>
    public class MultiTelegramLogger : ILogHandler
    {
        private readonly List<TelegramLogger> _loggers = new List<TelegramLogger>();
        private readonly LogLevel _minLevel;

        public MultiTelegramLogger(List<TelegramRecipient> recipients, LogLevel minLevel = LogLevel.WARNING)
        {
            _minLevel = minLevel;
            
            if (recipients != null && recipients.Count > 0)
            {
                foreach (var recipient in recipients)
                {
                    if (!string.IsNullOrEmpty(recipient.BotToken) && !string.IsNullOrEmpty(recipient.ChatId))
                    {
                        try
                        {
                            var logger = new TelegramLogger(recipient.BotToken, recipient.ChatId, minLevel);
                            _loggers.Add(logger);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to create Telegram logger for {recipient.ChatId}: {ex.Message}");
                        }
                    }
                }
            }
        }

        public void Handle(LogEntry entry)
        {
            if (entry.Level < _minLevel) return;

            // Отправляем всем получателям (каждому свой токен и chat ID)
            // Каждый TelegramLogger уже обрабатывает отправку асинхронно внутри
            foreach (var logger in _loggers)
            {
                try
                {
                    // Вызываем Handle для каждого логгера - каждый отправит сообщение своему получателю
                    logger.Handle(entry);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to send log to Telegram: {ex.Message}");
                }
            }
        }

        public async Task SendFullLogAsync(string logContent)
        {
            var tasks = _loggers.Select(logger => 
            {
                try
                {
                    return logger.SendFullLogAsync(logContent);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to send full log: {ex.Message}");
                    return Task.CompletedTask;
                }
            });

            await Task.WhenAll(tasks);
        }

        /// <summary>
        /// Отправляет файл всем получателям
        /// </summary>
        public async Task SendFileAsync(string filePath, string? caption = null)
        {
            var tasks = _loggers.Select(logger => 
            {
                try
                {
                    return logger.SendFileAsync(filePath, caption);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to send file: {ex.Message}");
                    return Task.CompletedTask;
                }
            });

            await Task.WhenAll(tasks);
        }

        /// <summary>
        /// Отправляет несколько файлов всем получателям
        /// </summary>
        public async Task SendFilesAsync(List<string> filePaths, string? message = null)
        {
            var tasks = _loggers.Select(logger => 
            {
                try
                {
                    return logger.SendFilesAsync(filePaths, message);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to send files: {ex.Message}");
                    return Task.CompletedTask;
                }
            });

            await Task.WhenAll(tasks);
        }

        public void Dispose()
        {
            foreach (var logger in _loggers)
            {
                try
                {
                    logger.Dispose();
                }
                catch { }
            }
            _loggers.Clear();
        }
    }
}

