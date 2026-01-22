using System;
using System.Diagnostics;
using Autodesk.AutoCAD.EditorInput;

namespace XingManager.Services
{
    public static class Logger
    {
        public enum Level { Debug = 0, Info = 1, Warn = 2, Error = 3 }

        static Logger()
        {
            var env = (Environment.GetEnvironmentVariable("XING_LOG_LEVEL") ?? string.Empty)
                .Trim()
                .ToUpperInvariant();

            if (env == "DEBUG")
            {
                CurrentLevel = Level.Debug;
            }
            else if (env == "WARN")
            {
                CurrentLevel = Level.Warn;
            }
            else if (env == "ERROR")
            {
                CurrentLevel = Level.Error;
            }
            else
            {
                CurrentLevel = Level.Info;
            }
        }

        public static Level CurrentLevel { get; set; }

        public static void Debug(Editor ed, string msg)
        {
            if (CurrentLevel <= Level.Debug) Write(ed, "DEBUG", msg);
        }

        public static void Info(Editor ed, string msg)
        {
            if (CurrentLevel <= Level.Info) Write(ed, "INFO", msg);
        }

        public static void Warn(Editor ed, string msg)
        {
            if (CurrentLevel <= Level.Warn) Write(ed, "WARN", msg);
        }

        public static void Error(Editor ed, string msg)
        {
            Write(ed, "ERROR", msg);
        }

        public static IDisposable Scope(Editor ed, string name, string kv = null)
        {
            Write(ed, "INFO", $"start {name}{(string.IsNullOrEmpty(kv) ? string.Empty : " " + kv)}");
            var sw = Stopwatch.StartNew();
            return new ScopeImpl(() =>
            {
                sw.Stop();
                Write(ed, "INFO", $"end   {name} elapsed_ms={sw.ElapsedMilliseconds}{(string.IsNullOrEmpty(kv) ? string.Empty : " " + kv)}");
            });
        }

        private sealed class ScopeImpl : IDisposable
        {
            private readonly Action _onDispose;

            public ScopeImpl(Action onDispose)
            {
                _onDispose = onDispose;
            }

            public void Dispose()
            {
                try
                {
                    _onDispose?.Invoke();
                }
                catch
                {
                }
            }
        }

        private static void Write(Editor ed, string level, string msg)
        {
            if (string.IsNullOrWhiteSpace(msg))
            {
                return;
            }

            try
            {
                ed?.WriteMessage($"\n[CrossingManager][{level}] {msg}");
            }
            catch
            {
            }
        }
    }
}

/////////////////////////////////////////////////////////////////////

