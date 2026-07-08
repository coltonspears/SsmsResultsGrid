using System;
using System.Windows.Threading;
using SsmsResultsGrid.Core.Mvvm;
using SsmsResultsGrid.Services.Diagnostics;

namespace SsmsResultsGrid.Views
{
    /// <summary>
    /// IUiDispatcher over the WPF dispatcher of the VS main thread. BeginInvoke
    /// preserves posting order, which the progressive-load pipeline relies on
    /// (BeginLoad → AppendRows → CompleteLoad must apply in sequence).
    /// </summary>
    internal sealed class WpfUiDispatcher : IUiDispatcher
    {
        private readonly Dispatcher _dispatcher;
        private readonly DiagnosticsPane _pane;

        /// <summary>Create on the main thread.</summary>
        public WpfUiDispatcher(DiagnosticsPane pane)
        {
            _dispatcher = Dispatcher.CurrentDispatcher;
            _pane = pane;
        }

        public bool CheckAccess() => _dispatcher.CheckAccess();

        public void Post(Action action)
        {
            if (action == null) return;
            // VSTHRD001: Dispatcher.BeginInvoke is deliberate here — the progressive-load
            // pipeline requires strict FIFO ordering of posted actions, which
            // SwitchToMainThreadAsync does not guarantee across JTF continuations.
#pragma warning disable VSTHRD001
            _ = _dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    _pane?.WriteFailure("UiDispatcher.Post", ex);
                }
            }), DispatcherPriority.Normal);
#pragma warning restore VSTHRD001
        }
    }
}
