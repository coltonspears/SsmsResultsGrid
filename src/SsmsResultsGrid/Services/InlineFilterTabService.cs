using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows;
using Microsoft.VisualStudio.Shell;
using SsmsResultsGrid.ToolWindows;

namespace SsmsResultsGrid.Services
{
    /// <summary>
    /// Hosts the filterable grid as a sibling tab inside SSMS's results/messages tab strip.
    /// Falls back to the tool window when the host cannot be discovered safely.
    /// </summary>
    internal sealed class InlineFilterTabService
    {
        private const string FilterTabName = "__SsmsResultsGrid_FilterTab";
        private const string FilterTabTitle = "Filter";

        public bool TryShowOrUpdate(Control sourceGrid, DataTable table, string failureReason, bool activateTab, out string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            reason = "inline-host-unavailable";

            if (sourceGrid == null)
            {
                reason = "source-grid-null";
                return false;
            }

            if (!TryFindTabHost(sourceGrid, out var tabControl, out _, out reason))
            {
                return false;
            }

            var filterPage = tabControl.TabPages.Cast<TabPage>()
                .FirstOrDefault(page => string.Equals(page.Name, FilterTabName, StringComparison.Ordinal));

            var firstLoad = false;
            if (filterPage == null)
            {
                filterPage = new TabPage(FilterTabTitle)
                {
                    Name = FilterTabName,
                    Padding = Padding.Empty
                };
                filterPage.AutoScroll = false;
                tabControl.TabPages.Add(filterPage);
                firstLoad = true;
            }

            var host = filterPage.Controls.OfType<ElementHost>().FirstOrDefault();
            var control = host?.Child as FilterableGridControl;
            if (control == null)
            {
                filterPage.Controls.Clear();
                host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Margin = Padding.Empty,
                    Padding = Padding.Empty
                };
                control = new FilterableGridControl
                {
                    Margin = new Thickness(0),
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
                    VerticalAlignment = System.Windows.VerticalAlignment.Stretch
                };
                host.Child = control;
                filterPage.Controls.Add(host);
                firstLoad = true;
            }

            control.LoadCaptureResult(table, failureReason);

            if (activateTab || (firstLoad && tabControl.SelectedTab == null))
            {
                tabControl.SelectedTab = filterPage;
            }

            reason = "ok";
            return true;
        }

        private static bool TryFindTabHost(Control grid, out TabControl tabControl, out TabPage currentPage, out string reason)
        {
            tabControl = null;
            currentPage = null;
            reason = "tab-host-not-found";

            for (var node = grid; node != null; node = node.Parent)
            {
                if (!(node is TabPage page)) continue;
                currentPage = page;
                tabControl = page.Parent as TabControl;
                if (tabControl != null)
                {
                    reason = "ok";
                    return true;
                }
            }

            reason = "tab-control-ancestor-not-found";
            return false;
        }
    }
}
