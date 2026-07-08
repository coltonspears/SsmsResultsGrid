using System;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SsmsResultsGrid.Services.Diagnostics
{
    /// <summary>
    /// Lazily-created "Results View" pane in the Output window. Failures are always
    /// logged; success paths stay silent so the pane is empty when everything works.
    /// Safe to call from any thread.
    /// </summary>
    internal sealed class DiagnosticsPane
    {
        private static readonly Guid PaneGuid = new Guid("8f1c9c2c-4f0a-4e9e-9d7b-1e2a4f3b0007");

        private readonly AsyncPackage _package;
        private IVsOutputWindowPane _pane;

        public DiagnosticsPane(AsyncPackage package)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
        }

        public void WriteFailure(string step, Exception ex) =>
            Write($"FAILED {step}: {ex}");

        public void WriteInfo(string message) =>
            Write(message);

        private void Write(string message)
        {
            var line = string.Format(
                CultureInfo.InvariantCulture,
                "[{0:HH:mm:ss.fff}] {1}{2}",
                DateTime.Now,
                message,
                Environment.NewLine);

            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await _package.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);
                    var pane = EnsurePane();
                    pane?.OutputStringThreadSafe(line);
                }
                catch
                {
                    // Logging must never take the host down.
                }
            });
        }

        private IVsOutputWindowPane EnsurePane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_pane != null) return _pane;

            if (!(_package.GetService<SVsOutputWindow, IVsOutputWindow>() is IVsOutputWindow output)) return null;

            var guid = PaneGuid;
            output.GetPane(ref guid, out _pane);
            if (_pane == null)
            {
                output.CreatePane(ref guid, Resources.Strings.OutputPaneTitle, fInitVisible: 1, fClearWithSolution: 0);
                output.GetPane(ref guid, out _pane);
            }
            return _pane;
        }
    }
}
