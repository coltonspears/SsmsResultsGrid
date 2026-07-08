using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SsmsResultsGrid.Services.Diagnostics;
using SsmsResultsGrid.Services.InPaneTab;
using SsmsResultsGrid.Services.Settings;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid.Services.Execution
{
    /// <summary>
    /// Slim priority command target that observes (never consumes) the T-SQL Execute
    /// command. Its only job is to attach a <see cref="ResultsTabSupervisor"/> to the
    /// active query window and arm it — actual capture is driven by the supervisor's
    /// query-completion event hooks.
    /// </summary>
    internal sealed class QueryExecutionTrigger :
        Microsoft.VisualStudio.OLE.Interop.IOleCommandTarget, IDisposable
    {
        // SQL editor command set; 0x0100 is Execute (F5 / toolbar / Ctrl+E).
        private static readonly Guid SqlEditorCmdSet = new Guid("52692960-56BC-4989-B5D3-94C47B513AE0");
        private const uint ExecuteCmdId = 0x0100;

        private readonly AsyncPackage _package;
        private readonly ExtensionSettings _settings;
        private readonly DiagnosticsPane _pane;
        private IVsRegisterPriorityCommandTarget _registrar;
        private uint _cookie;

        public QueryExecutionTrigger(AsyncPackage package, ExtensionSettings settings, DiagnosticsPane pane)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _pane = pane;
        }

        public async Task StartAsync(System.Threading.CancellationToken ct)
        {
            await _package.JoinableTaskFactory.SwitchToMainThreadAsync(ct);
            if (await _package.GetServiceAsync(typeof(SVsRegisterPriorityCommandTarget))
                is IVsRegisterPriorityCommandTarget registrar)
            {
                _registrar = registrar;
                registrar.RegisterPriorityCommandTarget(0, this, out _cookie);
            }
            else
            {
                _pane?.WriteFailure(
                    nameof(QueryExecutionTrigger),
                    new InvalidOperationException("SVsRegisterPriorityCommandTarget unavailable; auto-capture disabled."));
            }
        }

        public void Dispose()
        {
            if (_registrar != null && _cookie != 0)
            {
                try { _registrar.UnregisterPriorityCommandTarget(_cookie); }
                catch { /* shutdown */ }
                _cookie = 0;
            }
            _registrar = null;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds,
            Microsoft.VisualStudio.OLE.Interop.OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (pguidCmdGroup == SqlEditorCmdSet && nCmdID == ExecuteCmdId && _settings.ResultsToFilterGrid)
            {
                try
                {
                    OnExecuteObserved();
                }
                catch (Exception ex)
                {
                    _pane?.WriteFailure(nameof(OnExecuteObserved), ex);
                }
            }
            return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
        }

        private void OnExecuteObserved()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var docView = ActiveDocViewResolver.GetActiveSqlEditorDocView(_package);
            if (docView == null) return;

            var supervisor = ResultsTabSupervisor.GetOrCreate(docView, _package, _pane, _settings);
            supervisor.OnExecuteObserved();
        }
    }

    /// <summary>Resolves the active document's docView when it is an SSMS SQL query editor.</summary>
    internal static class ActiveDocViewResolver
    {
        public static object GetActiveSqlEditorDocView(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(serviceProvider.GetService(typeof(SVsShellMonitorSelection)) is IVsMonitorSelection monitor))
                return null;

            if (monitor.GetCurrentElementValue(
                    (uint)VSConstants.VSSELELEMID.SEID_DocumentFrame, out object frameObj) != VSConstants.S_OK)
                return null;

            if (!(frameObj is IVsWindowFrame frame)) return null;
            if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) != VSConstants.S_OK)
                return null;

            var typeName = docView?.GetType().FullName ?? string.Empty;
            return typeName.IndexOf("SqlScriptEditorControl", StringComparison.Ordinal) >= 0 ? docView : null;
        }
    }
}
