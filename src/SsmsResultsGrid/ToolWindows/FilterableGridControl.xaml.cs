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
        private string _lastDiagnostics;

        public FilterableGridControl()
        {
            InitializeComponent();
        }

        public event EventHandler FilterTextChanged;

        public string FilterText
        {
            get => FilterBox.Text ?? string.Empty;
            set
            {
                var next = value ?? string.Empty;
                if (!string.Equals(FilterBox.Text, next, StringComparison.Ordinal))
                {
                    FilterBox.Text = next;
                }
            }
        }

        public void LoadData(DataTable table)
        {
            if (table == null)
            {
                ResultsGrid.ItemsSource = null;
                StatusText.Text = "No results captured.";
                _view = null;
                SetDiagnostics(null);
                return;
            }

            _view = table.DefaultView;
            ResultsGrid.ItemsSource = _view;
            UpdateStatus();
            ApplyFilter(FilterBox.Text);
            SetDiagnostics(null);
        }

        public void LoadCaptureResult(DataTable table, string failureReason)
        {
            if (table == null)
            {
                ResultsGrid.ItemsSource = null;
                _view = null;
                var message = string.IsNullOrWhiteSpace(failureReason)
                    ? "Unable to capture an active SSMS results grid."
                    : failureReason;
                StatusText.Text = FirstLine(message);
                SetDiagnostics(message);
                return;
            }

            LoadData(table);
        }

        private static string FirstLine(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return string.Empty;
            var firstNewline = message.IndexOf('\n');
            if (firstNewline >= 0) return message.Substring(0, firstNewline).Trim();
            return message.Trim();
        }

        private void SetDiagnostics(string rawMessage)
        {
            _lastDiagnostics = string.IsNullOrWhiteSpace(rawMessage) ? null : FormatDiagnostics(rawMessage);
            DiagnosticsText.Text = _lastDiagnostics ?? string.Empty;
            DiagnosticsBorder.Visibility = _lastDiagnostics == null ? Visibility.Collapsed : Visibility.Visible;
        }

        private static string FormatDiagnostics(string rawMessage)
        {
            var text = rawMessage.Replace(" [", "\n[");
            text = text.Replace("; ", ";\n");
            text = text.Replace(", ", ",\n");
            text = text.Replace("sample=", "sample=\n");
            text = text.Replace("managedTypes=", "managedTypes=\n");
            text = text.Replace("win32Classes=", "win32Classes=\n");
            return text;
        }

        private void UpdateStatus()
        {
            if (_view == null)
            {
                StatusText.Text = "No captured rows";
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
            ClearFilterButton.Visibility = string.IsNullOrEmpty(FilterBox.Text)
                ? Visibility.Collapsed
                : Visibility.Visible;
            FilterTextChanged?.Invoke(this, EventArgs.Empty);

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
            var table = service.TryCaptureActiveDetailed(out _);
            if (table != null)
            {
                LoadData(table);
            }
            else
            {
                var message = service.LastFailureReason ?? "Unable to capture an active SSMS results grid.";
                StatusText.Text = FirstLine(message);
                SetDiagnostics(message);
            }
        }

        private void ClearFilterButton_Click(object sender, RoutedEventArgs e)
        {
            FilterBox.Clear();
            FilterBox.Focus();
        }

        private void CopyDiagnosticsButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_lastDiagnostics)) return;
            Clipboard.SetText(_lastDiagnostics);
            StatusText.Text = "Capture diagnostics copied.";
        }
    }
}
