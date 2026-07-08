using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid.Commands
{
    /// <summary>
    /// Checkable Tools-menu command mirroring SSMS's native "Results To" options:
    /// when checked, every query completion routes results into the filterable
    /// Results View tab automatically; when unchecked, the tab is only shown via
    /// the explicit Show command.
    /// </summary>
    internal sealed class ToggleResultsToGridCommand
    {
        private readonly FilterableGridPackage _package;
        private readonly OleMenuCommand _command;

        private ToggleResultsToGridCommand(FilterableGridPackage package, OleMenuCommandService commandService)
        {
            _package = package;
            var cmdId = new CommandID(PackageGuids.CommandSet, PackageGuids.ToggleResultsToGridCmdId);
            _command = new OleMenuCommand(Execute, cmdId);
            _command.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(_command);
        }

        public static async Task InitializeAsync(FilterableGridPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService
                ?? throw new InvalidOperationException("Unable to acquire IMenuCommandService.");
            _ = new ToggleResultsToGridCommand(package, commandService);
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            _command.Checked = _package.Settings?.ResultsToFilterGrid == true;
        }

        private void Execute(object sender, EventArgs e)
        {
            var settings = _package.Settings;
            if (settings == null) return;

            settings.ResultsToFilterGrid = !settings.ResultsToFilterGrid;
            _command.Checked = settings.ResultsToFilterGrid;
        }
    }
}
