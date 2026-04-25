using System;
using System.Collections.Generic;
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
        private static readonly Guid VsStd97CmdSet = VSConstants.GUID_VSStandardCommandSet97;
        private const uint ExecuteCmdId = 0x0100;
        private static readonly HashSet<uint> AdditionalExecuteCommandIds = new HashSet<uint>
        {
            0x0101, // SSMS variants sometimes route "execute with options" through adjacent IDs.
            0x0102
        };

        private readonly FilterableGridPackage _package;
        private IVsRunningDocumentTable _rdt;
        private uint _rdtCookie;
        private DispatcherTimer _pollTimer;
        private int _pollAttemptsLeft;
        private string _activeDocumentKey;
        private string _lastExecutionSource;

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
                DebugOutput.Write("Registered priority command target for query execution.");
            }
            else
            {
                DebugOutput.Write("Could not acquire SVsRegisterPriorityCommandTarget; auto-refresh will not observe Execute commands.");
            }

            // RDT events are a cheap way to know when new docs open; we don't strictly
            // need them, but they let us wire up future per-document state if needed.
            _rdt = await _package.GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            _rdt?.AdviseRunningDocTableEvents(this, out _rdtCookie);
            DebugOutput.Write("QueryExecutionListener started.");
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

        internal void OnQueryExecuted(string source)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // SSMS renders results asynchronously after the Execute command returns.
            // Poll every 500ms for up to 30s so streaming queries still populate.
            _lastExecutionSource = source;
            _activeDocumentKey = GetActiveDocumentKey();
            _pollAttemptsLeft = 60;
            DebugOutput.Write($"Query execution observed: source={source}, document={_activeDocumentKey ?? "<unknown>"}.");
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
            var table = _package.CaptureService.TryCaptureActiveDetailed(out var diagnostics);
            var attempt = 60 - _pollAttemptsLeft;
            if (attempt == 1 || attempt % 10 == 0 || table != null)
            {
                DebugOutput.Write($"Capture poll {attempt}: table={(table == null ? "null" : $"{table.Rows.Count} rows/{table.Columns.Count} cols")}, candidates={diagnostics?.CandidateCount ?? 0}, visible={diagnostics?.VisibleCandidateCount ?? 0}.");
            }

            if (table != null && table.Rows.Count > 0)
            {
                _pollTimer.Stop();
                DebugOutput.Write("Capture succeeded; pushing result table to filterable surface.");
                PushToSurface(table, null, _activeDocumentKey, activateInlineTab: true);
                return;
            }

            if (_pollAttemptsLeft <= 0)
            {
                _pollTimer.Stop();
                // If we captured an empty grid, still push it so column headers show.
                if (table != null)
                {
                    DebugOutput.Write("Capture timed out with empty table; pushing headers/empty result to filterable surface.");
                    PushToSurface(table, null, _activeDocumentKey, activateInlineTab: true);
                }
                else
                {
                    var details = _package.CaptureService.LastFailureReason ?? "Unable to capture an active SSMS results grid.";
                    if (!string.IsNullOrEmpty(_lastExecutionSource))
                    {
                        details += $" [source={_lastExecutionSource}]";
                    }
                    if (diagnostics != null && diagnostics.CandidateCount > 0)
                    {
                        details += $" [poll-timeout after 60 attempts; candidates={diagnostics.CandidateCount}]";
                    }
                    DebugOutput.Write("Capture timed out without table: " + details);
                    PushToSurface(null, details, _activeDocumentKey, activateInlineTab: true);
                }
            }
        }

        private void PushToSurface(System.Data.DataTable table, string failureReason, string documentKey, bool activateInlineTab)
        {
            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                await _package.JoinableTaskFactory.SwitchToMainThreadAsync();
                var sourceGrid = _package.CaptureService.LastCapturedGridControl;
                if (_package.InlineTabService != null &&
                    _package.InlineTabService.TryShowOrUpdate(sourceGrid, table, failureReason, activateInlineTab, out var inlineReason))
                {
                    DebugOutput.Write("Pushed capture to inline surface: " + inlineReason);
                    return;
                }

                DebugOutput.Write("Inline surface unavailable; falling back to tool window.");
                var window = await _package.ShowToolWindowAsync(
                    typeof(FilterableGridToolWindow),
                    0,
                    create: true,
                    cancellationToken: _package.DisposalToken) as FilterableGridToolWindow;
                window?.LoadCaptureResult(table, failureReason, documentKey);
            });
        }

        private string GetActiveDocumentKey()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (Package.GetGlobalService(typeof(SVsShellMonitorSelection)) is IVsMonitorSelection monitorSelection &&
                monitorSelection.GetCurrentElementValue((uint)VSConstants.VSSELELEMID.SEID_WindowFrame, out var frameObj) == VSConstants.S_OK)
            {
                if (frameObj is IVsWindowFrame frame)
                {
                    return TryGetDocumentKeyFromFrame(frame);
                }
            }
            return null;
        }

        private string TryGetDocumentKeyFromFrame(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (frame == null) return null;
            if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out var mkObj) == VSConstants.S_OK)
            {
                var moniker = mkObj as string;
                if (!string.IsNullOrWhiteSpace(moniker)) return moniker;
            }

            if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out var captionObj) == VSConstants.S_OK)
            {
                var caption = captionObj as string;
                if (!string.IsNullOrWhiteSpace(caption)) return caption;
            }

            return null;
        }

        #region IVsRunningDocTableEvents (stubs)
        public int OnAfterFirstDocumentLock(uint _, uint __, uint ___, uint ____) => VSConstants.S_OK;
        public int OnBeforeLastDocumentUnlock(uint _, uint __, uint ___, uint ____) => VSConstants.S_OK;
        public int OnAfterSave(uint _) => VSConstants.S_OK;
        public int OnAfterAttributeChange(uint _, uint __) => VSConstants.S_OK;
        public int OnBeforeDocumentWindowShow(uint _, int __, IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return VSConstants.S_OK;
        }
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
                if (pguidCmdGroup == SqlEditorCmdSet &&
                    (nCmdID == ExecuteCmdId || AdditionalExecuteCommandIds.Contains(nCmdID)))
                {
                    DebugOutput.Write($"Exec observed for SQL command: {nCmdID:X4}.");
                    try { _owner.OnQueryExecuted($"sql-cmd:{nCmdID:X4}"); } catch (Exception ex) { DebugOutput.Write("OnQueryExecuted failed: " + ex); }
                }
                else if (pguidCmdGroup == VsStd97CmdSet && nCmdID == (uint)VSConstants.VSStd97CmdID.Start)
                {
                    DebugOutput.Write("Exec observed for VS Start command.");
                    try { _owner.OnQueryExecuted("vsstd97-start"); } catch (Exception ex) { DebugOutput.Write("OnQueryExecuted failed: " + ex); }
                }
                return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
            }
        }
    }
}
