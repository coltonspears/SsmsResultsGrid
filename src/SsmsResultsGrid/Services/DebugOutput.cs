using System.Diagnostics;
using System;
using System.IO;

namespace SsmsResultsGrid.Services
{
    internal static class DebugOutput
    {
        private const string Prefix = "[SsmsResultsGrid] ";
        private const string LogPath = @"C:\logs\ssms_resultsview.txt";
        private static readonly object SyncRoot = new object();

        public static void Write(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            var line = Prefix + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + message;
            Debug.WriteLine(line);
            Trace.WriteLine(line);

            try
            {
                lock (SyncRoot)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                    File.AppendAllText(LogPath, line + Environment.NewLine);
                }
            }
            catch
            {
                // Logging must never interrupt SSMS query execution or grid capture.
            }
        }
    }
}
