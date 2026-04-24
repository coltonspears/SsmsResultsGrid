using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using SsmsResultsGrid.Commands;
using SsmsResultsGrid.Services;
using SsmsResultsGrid.ToolWindows;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("Filterable Results Grid", "Filterable results grid for SSMS 22", "0.1.0")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(FilterableGridToolWindow), Style = Microsoft.VisualStudio.Shell.VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    [Guid(PackageGuids.PackageGuidString)]
    [ProvideAutoLoad(Microsoft.VisualStudio.VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(Microsoft.VisualStudio.VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class FilterableGridPackage : AsyncPackage
    {
        public static FilterableGridPackage Instance { get; private set; }

        internal SsmsGridCaptureService CaptureService { get; private set; }
        internal InlineFilterTabService InlineTabService { get; private set; }
        internal QueryExecutionListener ExecutionListener { get; private set; }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            Instance = this;
            CaptureService = new SsmsGridCaptureService(this);
            InlineTabService = new InlineFilterTabService();
            ExecutionListener = new QueryExecutionListener(this);

            await ShowFilterableGridCommand.InitializeAsync(this);
            await ExecutionListener.StartAsync(cancellationToken);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                ExecutionListener?.Stop();
            }
            base.Dispose(disposing);
        }
    }
}
