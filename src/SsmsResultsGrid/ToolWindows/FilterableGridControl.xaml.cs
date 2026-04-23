using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;

namespace SsmsResultsGrid.ToolWindows
{
    public partial class FilterableGridControl : UserControl
    {
        private DataView _view;
        private DispatcherTimer _debounce;

        public FilterableGridControl()
        {
            InitializeComponent();
        }

        public void LoadData(DataTable table)
        {
            if (table == null)
            {
                ResultsGrid.ItemsSource = null;
                StatusText.Text = "No results captured.";
                _view = null;
                return;
            }

            _view = table.DefaultView;
            ResultsGrid.ItemsSource = _view;
            UpdateStatus();
            ApplyFilter(FilterBox.Text);
        }

        private void UpdateStatus()
        {
            if (_view == null)
            {
                StatusText.Text = "No results captured.";
                return;
            }
            var total = _view.Table.Rows.Count;
            var shown = _view.Count;
            StatusText.Text = shown == total
                ? $"{total:N0} row{(total == 1 ? "" : "s")}"
                : $"{shown:N0} of {total:N0} row{(total == 1 ? "" : "s")}";
        }

        private void FilterBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Debounce so filtering stays responsive on large result sets.
            if (_debounce == null)
            {
                _debounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
                _debounce.Tick += (_, __) =>
                {
                    _debounce.Stop();
                    ApplyFilter(FilterBox.Text);
                };
            }
            _debounce.Stop();
            _debounce.Start();
        }

        private void ApplyFilter(string query)
        {
            if (_view == null) return;

            if (string.IsNullOrEmpty(query))
            {
                _view.RowFilter = string.Empty;
                UpdateStatus();
                return;
            }

            var escaped = query.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]");
            var clauses = new System.Text.StringBuilder();
            bool first = true;
            foreach (DataColumn col in _view.Table.Columns)
            {
                if (!first) clauses.Append(" OR ");
                first = false;
                clauses.Append("CONVERT([").Append(col.ColumnName.Replace("]", "]]")).Append("], 'System.String') LIKE '%").Append(escaped).Append("%'");
            }

            try
            {
                _view.RowFilter = clauses.ToString();
            }
            catch (Exception)
            {
                // Malformed filter expression — fall back to clearing so the grid stays usable.
                _view.RowFilter = string.Empty;
            }
            UpdateStatus();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            var service = FilterableGridPackage.Instance?.CaptureService;
            if (service == null) return;
            var table = service.TryCaptureActive();
            if (table != null)
            {
                LoadData(table);
            }
            else
            {
                StatusText.Text = service.LastFailureReason ?? "Unable to capture an active SSMS results grid.";
            }
        }
    }
}
