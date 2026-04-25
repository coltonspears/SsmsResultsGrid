using System;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Microsoft.VisualStudio.Shell;
using SsmsResultsGrid.ToolWindows;

namespace SsmsResultsGrid.Services
{
    /// <summary>
    /// Hosts the filterable grid as a sibling tab inside SSMS's results/messages tab strip.
    /// This intentionally avoids re-parenting SSMS's native results controls.
    /// </summary>
    internal sealed class InlineFilterTabService
    {
        private const string FilterTabName = "__SsmsResultsGrid_FilterTab";
        private const string FilterTabTitle = "Results View";
        private const string InlineHostName = "__SsmsResultsGrid_FilterHost";
        private string _lastFilterText = string.Empty;

        public bool TryShowOrUpdate(Control sourceGrid, DataTable table, string failureReason, bool activateTab, out string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            reason = "inline-host-unavailable";

            if (sourceGrid == null)
            {
                reason = "source-grid-null";
                DebugOutput.Write("Inline update skipped: source grid is null.");
                return false;
            }

            if (!TryFindTabHost(sourceGrid, out var tabControl, out var currentPage, out reason))
            {
                DebugOutput.Write("Inline update skipped: " + reason);
                return false;
            }

            var filterPage = tabControl.TabPages.Cast<TabPage>()
                .FirstOrDefault(page => string.Equals(page.Name, FilterTabName, StringComparison.Ordinal));
            if (filterPage == null)
            {
                filterPage = new TabPage(FilterTabTitle)
                {
                    Name = FilterTabName,
                    Padding = Padding.Empty,
                    AutoScroll = false
                };
                CopyTabIcon(currentPage, filterPage);
                tabControl.TabPages.Insert(0, filterPage);
                DebugOutput.Write("Created inline Results View tab at index 0.");
            }
            else
            {
                filterPage.Text = FilterTabTitle;
                CopyTabIcon(currentPage, filterPage);
                MoveTabToFirstPosition(tabControl, filterPage);
            }

            var control = GetOrCreateFilterControl(filterPage.Controls);
            TrackFilterText(control);
            if (!string.IsNullOrEmpty(control.FilterText) || string.IsNullOrEmpty(_lastFilterText))
            {
                _lastFilterText = control.FilterText;
            }

            control.FilterText = _lastFilterText;
            control.LoadCaptureResult(table, failureReason);
            _lastFilterText = control.FilterText;

            if (activateTab)
            {
                tabControl.SelectedTab = filterPage;
                DebugOutput.Write("Activated inline Results View tab.");
            }

            reason = "ok-inline-tab";
            DebugOutput.Write($"Inline Results View tab updated: rows={(table == null ? "null" : table.Rows.Count.ToString())}, cols={(table == null ? "null" : table.Columns.Count.ToString())}.");
            return true;
        }

        private static void MoveTabToFirstPosition(TabControl tabControl, TabPage filterPage)
        {
            if (tabControl == null || filterPage == null)
            {
                return;
            }

            var currentIndex = tabControl.TabPages.IndexOf(filterPage);
            if (currentIndex <= 0)
            {
                return;
            }

            var selected = tabControl.SelectedTab;
            tabControl.TabPages.Remove(filterPage);
            tabControl.TabPages.Insert(0, filterPage);
            tabControl.SelectedTab = selected == filterPage ? filterPage : selected;
            DebugOutput.Write("Moved inline Results View tab to index 0.");
        }

        private static void CopyTabIcon(TabPage sourcePage, TabPage targetPage)
        {
            if (sourcePage == null || targetPage == null)
            {
                return;
            }

            targetPage.ImageKey = sourcePage.ImageKey;
            targetPage.ImageIndex = sourcePage.ImageIndex;
        }

        private void TrackFilterText(FilterableGridControl control)
        {
            control.FilterTextChanged -= FilterControl_FilterTextChanged;
            control.FilterTextChanged += FilterControl_FilterTextChanged;
        }

        private void FilterControl_FilterTextChanged(object sender, EventArgs e)
        {
            if (sender is FilterableGridControl control)
            {
                _lastFilterText = control.FilterText;
            }
        }

        private static FilterableGridControl GetOrCreateFilterControl(Control.ControlCollection controls)
        {
            var host = controls.OfType<ElementHost>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name, InlineHostName, StringComparison.Ordinal));
            var control = host?.Child as FilterableGridControl;
            if (control != null)
            {
                return control;
            }

            controls.Clear();
            host = new ElementHost
            {
                Name = InlineHostName,
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
            controls.Add(host);
            return control;
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
