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
    /// Hosts the filterable grid beside SSMS's native results grid when the
    /// grid host can be safely wrapped. Falls back to the older sibling tab
    /// approach when SSMS changes the immediate host shape.
    /// </summary>
    internal sealed class InlineFilterTabService
    {
        private const string InlineSplitName = "__SsmsResultsGrid_InlineSplit";
        private const string InlineHostName = "__SsmsResultsGrid_FilterHost";
        private const string FilterTabName = "__SsmsResultsGrid_FilterTab";
        private const string FilterTabTitle = "Filter";
        private string _lastFilterText = string.Empty;

        public bool TryShowOrUpdate(Control sourceGrid, DataTable table, string failureReason, bool activateTab, out string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            reason = "inline-host-unavailable";

            if (sourceGrid == null)
            {
                reason = "source-grid-null";
                return false;
            }

            if (TryShowBesideSourceGrid(sourceGrid, table, failureReason, out reason))
            {
                return true;
            }

            return TryShowInSiblingTab(sourceGrid, table, failureReason, activateTab, out reason);
        }

        private bool TryShowBesideSourceGrid(Control sourceGrid, DataTable table, string failureReason, out string reason)
        {
            reason = "inline-split-unavailable";

            if (!TryGetExistingSplit(sourceGrid, out var split) &&
                !TryCreateSplitAroundSourceGrid(sourceGrid, out split, out reason))
            {
                return false;
            }

            var control = GetOrCreateFilterControl(split.Panel2.Controls);
            TrackFilterText(control);
            if (!string.IsNullOrEmpty(control.FilterText) || string.IsNullOrEmpty(_lastFilterText))
            {
                _lastFilterText = control.FilterText;
            }

            control.FilterText = _lastFilterText;
            control.LoadCaptureResult(table, failureReason);
            _lastFilterText = control.FilterText;

            reason = "ok-inline-split";
            return true;
        }

        private static bool TryGetExistingSplit(Control sourceGrid, out SplitContainer split)
        {
            split = null;
            for (var node = sourceGrid; node != null; node = node.Parent)
            {
                if (node is SplitContainer candidate &&
                    string.Equals(candidate.Name, InlineSplitName, StringComparison.Ordinal))
                {
                    split = candidate;
                    return true;
                }
            }
            return false;
        }

        private static bool TryCreateSplitAroundSourceGrid(Control sourceGrid, out SplitContainer split, out string reason)
        {
            split = null;
            reason = "source-grid-parent-null";

            var parent = sourceGrid.Parent;
            if (parent == null)
            {
                return false;
            }

            if (parent is SplitterPanel && parent.Parent is SplitContainer existing &&
                string.Equals(existing.Name, InlineSplitName, StringComparison.Ordinal))
            {
                split = existing;
                reason = "ok-existing-inline-split";
                return true;
            }

            var childIndex = parent.Controls.GetChildIndex(sourceGrid);
            var originalDock = sourceGrid.Dock;
            var originalAnchor = sourceGrid.Anchor;
            var originalBounds = sourceGrid.Bounds;
            var originalMargin = sourceGrid.Margin;

            split = new SplitContainer
            {
                Name = InlineSplitName,
                Orientation = Orientation.Vertical,
                Dock = originalDock,
                Anchor = originalAnchor,
                Bounds = originalBounds,
                Margin = originalMargin,
                Panel1MinSize = 220,
                Panel2MinSize = 280,
                SplitterWidth = 5,
                TabStop = false
            };

            parent.SuspendLayout();
            split.Panel1.SuspendLayout();
            try
            {
                parent.Controls.Remove(sourceGrid);
                sourceGrid.Dock = DockStyle.Fill;
                sourceGrid.Margin = Padding.Empty;
                split.Panel1.Controls.Add(sourceGrid);
                parent.Controls.Add(split);
                parent.Controls.SetChildIndex(split, childIndex);
            }
            finally
            {
                split.Panel1.ResumeLayout();
                parent.ResumeLayout();
            }

            ConfigureSplitterDistance(split);
            EventHandler firstSize = null;
            firstSize = (sender, args) =>
            {
                split.SizeChanged -= firstSize;
                ConfigureSplitterDistance(split);
            };
            split.SizeChanged += firstSize;

            reason = "ok-created-inline-split";
            return true;
        }

        private static void ConfigureSplitterDistance(SplitContainer split)
        {
            if (split == null || split.Width <= 0)
            {
                return;
            }

            var maxDistance = split.Width - split.Panel2MinSize - split.SplitterWidth;
            if (maxDistance < split.Panel1MinSize)
            {
                return;
            }

            var preferredDistance = Math.Max(split.Panel1MinSize, (int)(split.Width * 0.58));
            split.SplitterDistance = Math.Min(preferredDistance, maxDistance);
        }

        private bool TryShowInSiblingTab(Control sourceGrid, DataTable table, string failureReason, bool activateTab, out string reason)
        {
            reason = "tab-host-unavailable";

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

            var control = GetOrCreateFilterControl(filterPage.Controls);
            TrackFilterText(control);
            firstLoad = firstLoad || filterPage.Controls.Count == 1;

            if (!string.IsNullOrEmpty(control.FilterText) || string.IsNullOrEmpty(_lastFilterText))
            {
                _lastFilterText = control.FilterText;
            }
            control.FilterText = _lastFilterText;
            control.LoadCaptureResult(table, failureReason);
            _lastFilterText = control.FilterText;

            if (activateTab || (firstLoad && tabControl.SelectedTab == null))
            {
                tabControl.SelectedTab = filterPage;
            }

            reason = "ok";
            return true;
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
