using System;
using System.IO;

namespace ExcelClone.Services
{
    public static class LoggingService
    {
        private static readonly string LogDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelClone", "Logs");
        public static string LogDirectory => LogDir;
        private static readonly string LogFile = Path.Combine(LogDir, "app.log");
        private static readonly object LockObj = new object();

        static LoggingService()
        {
            try
            {
                Directory.CreateDirectory(LogDir);
            }
            catch { }
        }

        public static void LogInfo(string message) => WriteEntry("INFO", message);
        public static void LogWarning(string message) => WriteEntry("WARN", message);
        public static void LogError(string message, Exception? ex = null)
        {
            string detail = ex != null ? $"{message} | Exception: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}" : message;
            WriteEntry("ERROR", detail);
        }

        public static string GetLogFilePath() => LogFile;

        public static string ReadLogContents()
        {
            lock (LockObj)
            {
                if (!File.Exists(LogFile)) return "No log entries recorded.";
                try
                {
                    return File.ReadAllText(LogFile);
                }
                catch (Exception ex)
                {
                    return $"Failed to read log: {ex.Message}";
                }
            }
        }

        public static void ClearLogs()
        {
            lock (LockObj)
            {
                try
                {
                    if (File.Exists(LogFile)) File.Delete(LogFile);
                }
                catch { }
            }
        }

        private static void WriteEntry(string level, string message)
        {
            lock (LockObj)
            {
                try
                {
                    string entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";
                    File.AppendAllText(LogFile, entry);
                }
                catch { }
            }
        }
    }
}

