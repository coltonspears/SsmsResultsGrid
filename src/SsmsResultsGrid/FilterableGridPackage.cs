using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using SsmsResultsGrid.Commands;
using SsmsResultsGrid.Services.Diagnostics;
using SsmsResultsGrid.Services.Execution;
using SsmsResultsGrid.Services.Settings;
using SsmsResultsGrid.ToolWindows;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid
{
    /// <summary>
    /// Package entry point. Owns the three long-lived services (settings, diagnostics
    /// pane, execute trigger) and registers the two menu commands. Everything per
    /// query window lives in <see cref="Services.InPaneTab.ResultsTabSupervisor"/>.
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("Results View for SSMS", "Filterable results view for SSMS 22", "1.0.0")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(FilterableGridToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    [Guid(PackageGuids.PackageGuidString)]
    [ProvideAutoLoad(Microsoft.VisualStudio.VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(Microsoft.VisualStudio.VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class FilterableGridPackage : AsyncPackage
    {
        internal ExtensionSettings Settings { get; private set; }
        internal DiagnosticsPane Diagnostics { get; private set; }

        private QueryExecutionTrigger _executionTrigger;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            Diagnostics = new DiagnosticsPane(this);

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            Settings = new ExtensionSettings(this);

            _executionTrigger = new QueryExecutionTrigger(this, Settings, Diagnostics);
            await _executionTrigger.StartAsync(cancellationToken);

            await ShowFilterableGridCommand.InitializeAsync(this);
            await ToggleResultsToGridCommand.InitializeAsync(this);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _executionTrigger?.Dispose();
                _executionTrigger = null;
            }
            base.Dispose(disposing);
        }
    }
}
