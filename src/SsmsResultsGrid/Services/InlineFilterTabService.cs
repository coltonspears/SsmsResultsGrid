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
    /// Hosts the filterable grid beside SSMS's native results grid by wrapping
    /// the existing results tab contents in a split view.
    /// </summary>
    internal sealed class InlineFilterTabService
    {
        private const string InlineSplitName = "__SsmsResultsGrid_InlineSplit";
        private const string InlineHostName = "__SsmsResultsGrid_FilterHost";
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

            if (!TryFindTabHost(sourceGrid, out _, out var currentPage, out reason))
            {
                return false;
            }

            if (!TryGetExistingSplit(currentPage, sourceGrid, out var split) &&
                !TryCreateSplitAroundResultsPage(currentPage, out split, out reason))
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

            if (activateTab && currentPage.Parent is TabControl tabControl)
            {
                tabControl.SelectedTab = currentPage;
            }

            reason = "ok-inline-split";
            return true;
        }

        private static bool TryGetExistingSplit(TabPage currentPage, Control sourceGrid, out SplitContainer split)
        {
            split = null;
            split = currentPage.Controls.OfType<SplitContainer>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name, InlineSplitName, StringComparison.Ordinal));
            if (split != null)
            {
                return true;
            }

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

        private static bool TryCreateSplitAroundResultsPage(TabPage currentPage, out SplitContainer split, out string reason)
        {
            split = null;
            reason = "results-page-empty";

            if (currentPage == null || currentPage.Controls.Count == 0)
            {
                return false;
            }

            var existingControls = currentPage.Controls.Cast<Control>().ToList();

            split = new SplitContainer
            {
                Name = InlineSplitName,
                Orientation = Orientation.Vertical,
                Dock = DockStyle.Fill,
                Panel1MinSize = 220,
                Panel2MinSize = 280,
                SplitterWidth = 5,
                TabStop = false
            };
            split.Panel1MinSize = 220;
            split.Panel2MinSize = 280;
            split.Panel1.Padding = Padding.Empty;
            split.Panel2.Padding = Padding.Empty;

            currentPage.SuspendLayout();
            split.Panel1.SuspendLayout();
            try
            {
                currentPage.Controls.Clear();
                currentPage.Controls.Add(split);
                foreach (var control in existingControls)
                {
                    split.Panel1.Controls.Add(control);
                }
            }
            finally
            {
                split.Panel1.ResumeLayout();
                currentPage.ResumeLayout();
            }

            ConfigureSplitterDistance(split);
            var createdSplit = split;
            EventHandler firstSize = null;
            firstSize = (sender, args) =>
            {
                createdSplit.SizeChanged -= firstSize;
                ConfigureSplitterDistance(createdSplit);
            };
            createdSplit.SizeChanged += firstSize;

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
