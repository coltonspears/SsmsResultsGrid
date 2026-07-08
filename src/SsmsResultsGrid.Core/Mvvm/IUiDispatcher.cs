using System;

namespace SsmsResultsGrid.Core.Mvvm
{
    /// <summary>
    /// Abstraction over the UI thread so view models stay unit-testable outside WPF.
    /// The VSIX supplies a JoinableTaskFactory-backed implementation; tests supply a
    /// synchronous fake.
    /// </summary>
    public interface IUiDispatcher
    {
        bool CheckAccess();

        /// <summary>Queue work onto the UI thread (fire-and-forget, exceptions observed by the host).</summary>
        void Post(Action action);
    }
}
