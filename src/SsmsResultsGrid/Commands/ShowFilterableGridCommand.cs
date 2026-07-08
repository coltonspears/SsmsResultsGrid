using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using SsmsResultsGrid.Services.Execution;
using SsmsResultsGrid.Services.InPaneTab;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid.Commands
{
    /// <summary>
    /// Tools-menu command that captures the active query window's grid results on
    /// demand and shows/activates the injected "Results View" tab.
    /// </summary>
    internal sealed class ShowFilterableGridCommand
    {
        private readonly FilterableGridPackage _package;

        private ShowFilterableGridCommand(FilterableGridPackage package, OleMenuCommandService commandService)
        {
            _package = package;
            var cmdId = new CommandID(PackageGuids.CommandSet, PackageGuids.ShowFilterableGridCmdId);
            commandService.AddCommand(new MenuCommand(Execute, cmdId));
        }

        public static async Task InitializeAsync(FilterableGridPackage package)
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
                try
                {
                    await _package.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);

                    var docView = ActiveDocViewResolver.GetActiveSqlEditorDocView(_package);
                    if (docView == null)
                    {
                        _package.Diagnostics?.WriteInfo(
                            "Show Results View invoked without an active SQL query window.");
                        return;
                    }

                    var supervisor = ResultsTabSupervisor.GetOrCreate(
                        docView, _package, _package.Diagnostics, _package.Settings);
                    await supervisor.ShowAsync();
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    _package.Diagnostics?.WriteFailure(nameof(ShowFilterableGridCommand), ex);
                }
            });
        }
    }
}
