using System;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.EditorInput;

namespace XingManager.Services
{
    internal static class CommandLogger
    {
        private static readonly object _sync = new object();
        private static readonly string _logFilePath = InitializeLogFilePath();

        private static string InitializeLogFilePath()
        {
            string rootDirectory = null;

            try
            {
                rootDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (string.IsNullOrWhiteSpace(rootDirectory))
                {
                    rootDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                }
            }
            catch
            {
                rootDirectory = null;
            }

            if (string.IsNullOrWhiteSpace(rootDirectory))
            {
                try
                {
                    rootDirectory = Path.GetTempPath();
                }
                catch
                {
                    rootDirectory = ".";
                }
            }

            string folder;
            try
            {
                folder = Path.Combine(rootDirectory, "CrossingManager", "Logs");
            }
            catch
            {
                folder = rootDirectory;
            }

            try
            {
                Directory.CreateDirectory(folder);
            }
            catch
            {
                folder = rootDirectory;
            }

            var fileName = $"CrossingManager_{DateTime.Now:yyyyMMdd}.log";
            try
            {
                return Path.Combine(folder, fileName);
            }
            catch
            {
                return Path.Combine(rootDirectory, "CrossingManager.log");
            }
        }

        private static void AppendLine(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            var timestamped = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [CrossingManager] {message}";

            lock (_sync)
            {
                try
                {
                    File.AppendAllText(_logFilePath, timestamped + Environment.NewLine, Encoding.UTF8);
                }
                catch
                {
                    // Swallow logging failures to avoid impacting command execution.
                }
            }
        }

        public static void Log(string message)
        {
            AppendLine(message);
        }

        public static void Log(Editor editor, string message, bool alsoToCommandBar = false)
        {
            AppendLine(message);

            if (!alsoToCommandBar || editor == null)
                return;

            Logger.Info(editor, message);
        }

        public static string LogFilePath => _logFilePath;
    }
}

