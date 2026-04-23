using System;
using System.Threading;
using System.Windows.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SsmsResultsGrid.ToolWindows;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid.Services
{
    /// <summary>
    /// Watches for SSMS query execution so the filterable grid can auto-refresh.
    ///
    /// SSMS does not raise a public "query completed" event. The reliable signal is
    /// the T-SQL Execute command being invoked (guid
    /// {52692960-56BC-4989-B5D3-94C47B513AE0}, id 0x0100). After the command fires,
    /// results render asynchronously — we poll the capture service a handful of
    /// times on the UI thread until data appears or we give up.
    /// </summary>
    internal sealed class QueryExecutionListener : IVsRunningDocTableEvents, IDisposable
    {
        private static readonly Guid SqlEditorCmdSet = new Guid("52692960-56BC-4989-B5D3-94C47B513AE0");
        private const uint ExecuteCmdId = 0x0100;

        private readonly FilterableGridPackage _package;
        private IVsRunningDocumentTable _rdt;
        private uint _rdtCookie;
        private DispatcherTimer _pollTimer;
        private int _pollAttemptsLeft;

        public QueryExecutionListener(FilterableGridPackage package)
        {
            _package = package;
        }

        public async Task StartAsync(CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Hook the command dispatcher so we can observe the T-SQL Execute command.
            if (await _package.GetServiceAsync(typeof(SVsRegisterPriorityCommandTarget)) is IVsRegisterPriorityCommandTarget reg)
            {
                reg.RegisterPriorityCommandTarget(0, new ExecuteCommandObserver(this), out _);
            }

            // RDT events are a cheap way to know when new docs open; we don't strictly
            // need them, but they let us wire up future per-document state if needed.
            _rdt = await _package.GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            _rdt?.AdviseRunningDocTableEvents(this, out _rdtCookie);
        }

        public void Stop() => Dispose();

        public void Dispose()
        {
            _pollTimer?.Stop();
            _pollTimer = null;
            if (_rdt != null && _rdtCookie != 0)
            {
                _rdt.UnadviseRunningDocTableEvents(_rdtCookie);
                _rdtCookie = 0;
            }
        }

        internal void OnQueryExecuted()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // SSMS renders results asynchronously after the Execute command returns.
            // Poll every 500ms for up to 30s so streaming queries still populate.
            _pollAttemptsLeft = 60;
            if (_pollTimer == null)
            {
                _pollTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                _pollTimer.Tick += PollForResults;
            }
            _pollTimer.Stop();
            _pollTimer.Start();
        }

        private void PollForResults(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _pollAttemptsLeft--;
            var table = _package.CaptureService.TryCaptureActive();

            if (table != null && table.Rows.Count > 0)
            {
                _pollTimer.Stop();
                PushToWindow(table);
                return;
            }

            if (_pollAttemptsLeft <= 0)
            {
                _pollTimer.Stop();
                // If we captured an empty grid, still push it so column headers show.
                if (table != null) PushToWindow(table);
            }
        }

        private void PushToWindow(System.Data.DataTable table)
        {
            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                await _package.JoinableTaskFactory.SwitchToMainThreadAsync();
                var window = await _package.ShowToolWindowAsync(
                    typeof(FilterableGridToolWindow),
                    0,
                    create: true,
                    cancellationToken: _package.DisposalToken) as FilterableGridToolWindow;
                window?.LoadData(table);
            });
        }

        #region IVsRunningDocTableEvents (stubs)
        public int OnAfterFirstDocumentLock(uint _, uint __, uint ___, uint ____) => VSConstants.S_OK;
        public int OnBeforeLastDocumentUnlock(uint _, uint __, uint ___, uint ____) => VSConstants.S_OK;
        public int OnAfterSave(uint _) => VSConstants.S_OK;
        public int OnAfterAttributeChange(uint _, uint __) => VSConstants.S_OK;
        public int OnBeforeDocumentWindowShow(uint _, int __, IVsWindowFrame ___) => VSConstants.S_OK;
        public int OnAfterDocumentWindowHide(uint _, IVsWindowFrame __) => VSConstants.S_OK;
        #endregion

        /// <summary>
        /// Priority command target that sniffs Exec() calls without consuming them.
        /// Returns OLECMDERR_E_NOTSUPPORTED so SSMS's normal execute path still runs.
        /// </summary>
        private sealed class ExecuteCommandObserver : Microsoft.VisualStudio.OLE.Interop.IOleCommandTarget
        {
            private readonly QueryExecutionListener _owner;

            public ExecuteCommandObserver(QueryExecutionListener owner) { _owner = owner; }

            public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, Microsoft.VisualStudio.OLE.Interop.OLECMD[] prgCmds, IntPtr pCmdText)
            {
                return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
            }

            public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
            {
                if (pguidCmdGroup == SqlEditorCmdSet && nCmdID == ExecuteCmdId)
                {
                    try { _owner.OnQueryExecuted(); } catch { }
                }
                return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
            }
        }
    }
}
