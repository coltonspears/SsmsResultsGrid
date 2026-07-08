using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using SsmsResultsGrid.Core.ViewModels;
using SsmsResultsGrid.Views;

namespace SsmsResultsGrid.ToolWindows
{
    /// <summary>
    /// Dockable fallback host for <see cref="ResultsViewControl"/>, used only when
    /// in-pane tab injection fails on an unexpected SSMS build. Content mirrors the
    /// same view model the tab would have shown.
    /// </summary>
    [Guid(PackageGuids.ToolWindowGuidString)]
    public sealed class FilterableGridToolWindow : ToolWindowPane
    {
        private readonly ResultsViewControl _control;

        public FilterableGridToolWindow() : base(null)
        {
            Caption = Resources.Strings.ToolWindowCaption;
            _control = new ResultsViewControl();
            _control.ApplySettings(Services.Settings.ExtensionSettings.Instance);
            Content = _control;
        }

        public void Bind(ResultsViewModel viewModel)
        {
            _control.DataContext = viewModel;
        }
    }
}
