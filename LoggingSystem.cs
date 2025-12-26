using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ZontSpecExtractor
{
    public enum LogLevel
    {
        DEBUG = 0,
        INFO = 1,
        WARNING = 2,
        ERROR = 3,
        CRITICAL = 4
    }

    public class LogEntry
    {
        public DateTime Timestamp { get; set; }
        public LogLevel Level { get; set; }
        public string Message { get; set; }
        public string MethodName { get; set; }
        public string ClassName { get; set; }
        public Exception Exception { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
        public object ReturnValue { get; set; }
        public long MemoryUsage { get; set; }
        public int ThreadId { get; set; }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append($"[{Timestamp:yyyy-MM-dd HH:mm:ss.fff}] ");
            sb.Append($"[{Level}] ");
            sb.Append($"[{ClassName}.{MethodName}] ");
            sb.Append($"[Thread:{ThreadId}] ");
            sb.Append($"[Mem:{MemoryUsage / 1024 / 1024}MB] ");
            
            if (Parameters != null && Parameters.Count > 0)
            {
                sb.Append($"Params: {string.Join(", ", Parameters.Select(kv => $"{kv.Key}={kv.Value}"))} ");
            }
            
            sb.Append($"Message: {Message}");
            
            if (ReturnValue != null)
            {
                sb.Append($" | Return: {ReturnValue}");
            }
            
            if (Exception != null)
            {
                sb.Append($" | Exception: {Exception.GetType().Name}: {Exception.Message}");
                sb.Append($" | StackTrace: {Exception.StackTrace}");
            }
            
            return sb.ToString();
        }
    }

    public static class LoggingSystem
    {
        private static readonly ConcurrentQueue<LogEntry> _logQueue = new ConcurrentQueue<LogEntry>();
        private static readonly List<ILogHandler> _handlers = new List<ILogHandler>();
        private static readonly object _lockObject = new object();
        private static bool _isInitialized = false;
        private static LogLevel _minLevel = LogLevel.DEBUG;
        private static bool _isPaused = false;
        private static readonly StringBuilder _logBuffer = new StringBuilder();
        private static System.Threading.Timer _autoSaveTimer;
        private static bool _autoSaveEnabled = false;
        private static int _autoSaveIntervalSeconds = 60;

        public static event EventHandler<LogEntry> LogAdded;

        public static void Initialize(LogLevel minLevel = LogLevel.DEBUG)
        {
            if (_isInitialized) return;
            
            _minLevel = minLevel;
            _isInitialized = true;
            
            // Запускаем обработку логов в фоне
            Task.Run(ProcessLogQueue);
            
            Log(LogLevel.INFO, "LoggingSystem", "Initialize", "Logging system initialized", null);
        }

        public static void AddHandler(ILogHandler handler)
        {
            lock (_lockObject)
            {
                _handlers.Add(handler);
            }
        }

        public static void RemoveHandler(ILogHandler handler)
        {
            lock (_lockObject)
            {
                _handlers.Remove(handler);
            }
        }

        public static void Log(LogLevel level, string className, string methodName, string message, 
            Exception exception = null, Dictionary<string, object> parameters = null, object returnValue = null)
        {
            if (!_isInitialized || level < _minLevel || _isPaused) return;

            var entry = new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = level,
                Message = message,
                MethodName = methodName ?? "Unknown",
                ClassName = className ?? "Unknown",
                Exception = exception,
                Parameters = parameters,
                ReturnValue = returnValue,
                MemoryUsage = GC.GetTotalMemory(false),
                ThreadId = Thread.CurrentThread.ManagedThreadId
            };

            _logQueue.Enqueue(entry);
            _logBuffer.AppendLine(entry.ToString());
            
            LogAdded?.Invoke(null, entry);
        }

        public static void LogMethodEntry(string className, string methodName, Dictionary<string, object> parameters = null)
        {
            Log(LogLevel.DEBUG, className, methodName, $"→ Entering method", null, parameters);
        }

        public static void LogMethodExit(string className, string methodName, object returnValue = null)
        {
            Log(LogLevel.DEBUG, className, methodName, $"← Exiting method", null, null, returnValue);
        }

        public static void LogException(string className, string methodName, Exception ex, Dictionary<string, object> parameters = null)
        {
            Log(LogLevel.ERROR, className, methodName, $"Exception occurred: {ex.Message}", ex, parameters);
        }

        private static async Task ProcessLogQueue()
        {
            while (_isInitialized)
            {
                try
                {
                    while (_logQueue.TryDequeue(out var entry))
                    {
                        lock (_lockObject)
                        {
                            foreach (var handler in _handlers)
                            {
                                try
                                {
                                    handler.Handle(entry);
                                }
                                catch (Exception ex)
                                {
                                    // Критично: обработчик логов не должен падать
                                    Debug.WriteLine($"Log handler error: {ex.Message}");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Log processing error: {ex.Message}");
                }

                await Task.Delay(100); // Проверяем каждые 100мс
            }
        }

        public static string GetLogBuffer()
        {
            return _logBuffer.ToString();
        }

        public static void ClearLogBuffer()
        {
            _logBuffer.Clear();
        }

        public static void Pause()
        {
            _isPaused = true;
            Log(LogLevel.INFO, "LoggingSystem", "Pause", "Logging paused");
        }

        public static void Resume()
        {
            _isPaused = false;
            Log(LogLevel.INFO, "LoggingSystem", "Resume", "Logging resumed");
        }

        public static void EnableAutoSave(int intervalSeconds, string filePath = null)
        {
            _autoSaveEnabled = true;
            _autoSaveIntervalSeconds = intervalSeconds;
            
            _autoSaveTimer?.Dispose();
            _autoSaveTimer = new System.Threading.Timer(_ =>
            {
                if (_autoSaveEnabled)
                {
                    SaveToFile(filePath ?? GetDefaultLogPath());
                }
            }, null, TimeSpan.Zero, TimeSpan.FromSeconds(intervalSeconds));
            
            Log(LogLevel.INFO, "LoggingSystem", "EnableAutoSave", $"Auto-save enabled: {intervalSeconds}s interval");
        }

        public static void DisableAutoSave()
        {
            _autoSaveEnabled = false;
            _autoSaveTimer?.Dispose();
            _autoSaveTimer = null;
            Log(LogLevel.INFO, "LoggingSystem", "DisableAutoSave", "Auto-save disabled");
        }

        public static void SaveToFile(string filePath = null)
        {
            try
            {
                filePath = filePath ?? GetDefaultLogPath();
                var logContent = GetLogBuffer();
                File.AppendAllText(filePath, logContent);
                ClearLogBuffer();
                Log(LogLevel.INFO, "LoggingSystem", "SaveToFile", $"Log saved to {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to save log: {ex.Message}");
            }
        }

        private static string GetDefaultLogPath()
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ZontSpecExtractor_Logs");
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, $"log_{timestamp}.log");
        }

        public static List<LogEntry> FilterLogs(string searchTerm, LogLevel? minLevel = null)
        {
            // Это упрощенная версия - в реальности нужно хранить все логи в памяти или БД
            var allLogs = new List<LogEntry>();
            // В реальной реализации здесь был бы доступ к хранилищу логов
            return allLogs.Where(log =>
                (minLevel == null || log.Level >= minLevel) &&
                (string.IsNullOrEmpty(searchTerm) || 
                 log.Message.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                 log.ClassName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                 log.MethodName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase))
            ).ToList();
        }

        public static (long MemoryBytes, double CpuPercent) GetSystemMetrics()
        {
            var process = Process.GetCurrentProcess();
            long memory = process.WorkingSet64;
            
            // CPU процент - упрощенная версия
            double cpu = 0;
            try
            {
                var startTime = DateTime.UtcNow;
                var startCpuUsage = process.TotalProcessorTime;
                Thread.Sleep(100);
                var endTime = DateTime.UtcNow;
                var endCpuUsage = process.TotalProcessorTime;
                var cpuUsedMs = (endCpuUsage - startCpuUsage).TotalMilliseconds;
                var totalMsPassed = (endTime - startTime).TotalMilliseconds;
                cpu = (cpuUsedMs / (Environment.ProcessorCount * totalMsPassed)) * 100;
            }
            catch { }
            
            return (memory, cpu);
        }
    }

    public interface ILogHandler
    {
        void Handle(LogEntry entry);
    }
}

