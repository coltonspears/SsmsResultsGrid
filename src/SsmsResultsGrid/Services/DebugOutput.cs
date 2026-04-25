using System.Diagnostics;

namespace SsmsResultsGrid.Services
{
    internal static class DebugOutput
    {
        private const string Prefix = "[SsmsResultsGrid] ";

        public static void Write(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            Debug.WriteLine(Prefix + message);
            Trace.WriteLine(Prefix + message);
        }
    }
}
