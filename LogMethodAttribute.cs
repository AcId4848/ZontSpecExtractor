using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace ZontSpecExtractor
{
    /// <summary>
    /// Атрибут для автоматического логирования вызовов методов
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class LogMethodAttribute : Attribute
    {
        public LogLevel MinLevel { get; set; } = LogLevel.DEBUG;
        public bool LogParameters { get; set; } = true;
        public bool LogReturnValue { get; set; } = true;
        public bool LogExecutionTime { get; set; } = false;
    }

    /// <summary>
    /// Статический класс для перехвата вызовов методов (упрощенная версия через AOP или ручное логирование)
    /// </summary>
    public static class MethodLogger
    {
        /// <summary>
        /// Обертка для логирования метода (вызывается вручную в начале метода)
        /// </summary>
        public static void LogEntry([CallerMemberName] string methodName = "", [CallerFilePath] string filePath = "", [CallerLineNumber] int lineNumber = 0, params object[] parameters)
        {
            var className = System.IO.Path.GetFileNameWithoutExtension(filePath);
            var paramsDict = new Dictionary<string, object>();
            
            if (parameters != null && parameters.Length > 0)
            {
                for (int i = 0; i < parameters.Length; i++)
                {
                    paramsDict[$"param{i}"] = parameters[i] ?? "null";
                }
            }
            
            LoggingSystem.LogMethodEntry(className, methodName, paramsDict);
        }

        /// <summary>
        /// Обертка для логирования выхода из метода
        /// </summary>
        public static void LogExit(object returnValue = null, [CallerMemberName] string methodName = "", [CallerFilePath] string filePath = "")
        {
            var className = System.IO.Path.GetFileNameWithoutExtension(filePath);
            LoggingSystem.LogMethodExit(className, methodName, returnValue);
        }

        /// <summary>
        /// Обертка для безопасного выполнения метода с логированием
        /// </summary>
        public static T ExecuteWithLogging<T>(Func<T> method, [CallerMemberName] string methodName = "", [CallerFilePath] string filePath = "")
        {
            var className = System.IO.Path.GetFileNameWithoutExtension(filePath);
            var stopwatch = Stopwatch.StartNew();
            
            try
            {
                LoggingSystem.LogMethodEntry(className, methodName);
                var result = method();
                stopwatch.Stop();
                
                LoggingSystem.Log(LogLevel.DEBUG, className, methodName, 
                    $"Method completed in {stopwatch.ElapsedMilliseconds}ms", null, null, result);
                
                return result;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                LoggingSystem.LogException(className, methodName, ex);
                throw;
            }
        }

        /// <summary>
        /// Обертка для void методов
        /// </summary>
        public static void ExecuteWithLogging(Action method, [CallerMemberName] string methodName = "", [CallerFilePath] string filePath = "")
        {
            var className = System.IO.Path.GetFileNameWithoutExtension(filePath);
            var stopwatch = Stopwatch.StartNew();
            
            try
            {
                LoggingSystem.LogMethodEntry(className, methodName);
                method();
                stopwatch.Stop();
                
                LoggingSystem.Log(LogLevel.DEBUG, className, methodName, 
                    $"Method completed in {stopwatch.ElapsedMilliseconds}ms");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                LoggingSystem.LogException(className, methodName, ex);
                throw;
            }
        }
    }

    /// <summary>
    /// Расширения для упрощения логирования
    /// </summary>
    public static class LoggingExtensions
    {
        /// <summary>
        /// Безопасное выполнение с логированием ошибок
        /// </summary>
        public static T SafeExecute<T>(this object obj, Func<T> action, T defaultValue = default(T), [CallerMemberName] string methodName = "")
        {
            try
            {
                return action();
            }
            catch (Exception ex)
            {
                var className = obj?.GetType().Name ?? "Unknown";
                LoggingSystem.LogException(className, methodName, ex);
                return defaultValue;
            }
        }

        /// <summary>
        /// Безопасное выполнение void методов
        /// </summary>
        public static void SafeExecute(this object obj, Action action, [CallerMemberName] string methodName = "")
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                var className = obj?.GetType().Name ?? "Unknown";
                LoggingSystem.LogException(className, methodName, ex);
            }
        }
    }
}


