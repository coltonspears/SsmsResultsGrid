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
            _dispatcher.BeginInvoke(new Action(() =>
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
        }
    }
}
