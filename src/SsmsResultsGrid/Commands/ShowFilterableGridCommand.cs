using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SsmsResultsGrid.ToolWindows;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid.Commands
{
    internal sealed class ShowFilterableGridCommand
    {
        private readonly AsyncPackage _package;

        private ShowFilterableGridCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package;
            var cmdId = new CommandID(PackageGuids.CommandSet, PackageGuids.ShowFilterableGridCmdId);
            var cmd = new MenuCommand(Execute, cmdId);
            commandService.AddCommand(cmd);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService
                ?? throw new InvalidOperationException("Unable to acquire IMenuCommandService.");
            _ = new ShowFilterableGridCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                await _package.JoinableTaskFactory.SwitchToMainThreadAsync();
                // Trigger an on-demand capture so the user sees current data immediately.
                if (FilterableGridPackage.Instance?.CaptureService != null)
                {
                    var table = FilterableGridPackage.Instance.CaptureService.TryCaptureActiveDetailed(out _);
                    var failure = table == null ? FilterableGridPackage.Instance.CaptureService.LastFailureReason : null;

                    var sourceGrid = FilterableGridPackage.Instance.CaptureService.LastCapturedGridControl;
                    if (FilterableGridPackage.Instance.InlineTabService != null &&
                        FilterableGridPackage.Instance.InlineTabService.TryShowOrUpdate(sourceGrid, table, failure, activateTab: true, out _))
                    {
                        return;
                    }
                }

                var window = await _package.ShowToolWindowAsync(
                    typeof(FilterableGridToolWindow),
                    0,
                    create: true,
                    cancellationToken: _package.DisposalToken) as FilterableGridToolWindow;
                if (window == null) return;

                if (window.Frame is IVsWindowFrame frame)
                {
                    ErrorHandler.ThrowOnFailure(frame.Show());
                }

                if (FilterableGridPackage.Instance?.CaptureService != null)
                {
                    var table = FilterableGridPackage.Instance.CaptureService.TryCaptureActiveDetailed(out _);
                    var failure = table == null ? FilterableGridPackage.Instance.CaptureService.LastFailureReason : null;
                    window.LoadCaptureResult(table, failure, contextKey: null);
                }
            });
        }
    }
}
