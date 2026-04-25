using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using Microsoft.Win32;

namespace SsmsResultsGrid.ToolWindows
{
    public partial class FilterableGridControl : UserControl
    {
        private DataView _view;
        private DispatcherTimer _debounce;
        private string _lastDiagnostics;
        private bool _defaultCaseSensitive;

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
                SummaryText.Text = "Rows: 0 | Columns: 0";
                FilterLatencyText.Text = "Filter: ready";
                LastUpdatedText.Text = "Updated: --";
                SetDiagnostics(null);
                return;
            }

            _view = table.DefaultView;
            _defaultCaseSensitive = table.CaseSensitive;
            ResultsGrid.ItemsSource = _view;
            UpdateStatus();
            UpdateSummary();
            ApplyFilter(FilterBox.Text);
            LastUpdatedText.Text = "Updated: " + DateTime.Now.ToString("HH:mm:ss");
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
                SummaryText.Text = "Rows: 0 | Columns: 0";
                FilterLatencyText.Text = "Filter: ready";
                LastUpdatedText.Text = "Updated: --";
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
                SummaryText.Text = "Rows: 0 | Columns: 0";
                return;
            }
            var total = _view.Table.Rows.Count;
            var shown = _view.Count;
            StatusText.Text = shown == total
                ? $"{total:N0} row{(total == 1 ? "" : "s")}"
                : $"{shown:N0} of {total:N0} row{(total == 1 ? "" : "s")}";
        }

        private void UpdateSummary()
        {
            if (_view == null)
            {
                SummaryText.Text = "Rows: 0 | Columns: 0";
                return;
            }

            var total = _view.Table.Rows.Count;
            var cols = _view.Table.Columns.Count;
            SummaryText.Text = $"Rows: {total:N0} | Columns: {cols:N0}";
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
            var sw = Stopwatch.StartNew();
            _view.Table.CaseSensitive = MatchCaseCheckBox.IsChecked == true || _defaultCaseSensitive;
            var mode = GetFilterMode();

            if (string.IsNullOrEmpty(query))
            {
                _view.RowFilter = string.Empty;
                UpdateStatus();
                FilterLatencyText.Text = $"Filter: {sw.ElapsedMilliseconds} ms";
                return;
            }

            var escaped = query.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]");
            var clauses = new System.Text.StringBuilder();
            bool first = true;
            foreach (DataColumn col in _view.Table.Columns)
            {
                if (!first) clauses.Append(" OR ");
                first = false;
                var colExpr = "CONVERT([" + col.ColumnName.Replace("]", "]]") + "], 'System.String')";
                if (mode == 1)
                {
                    clauses.Append(colExpr).Append(" LIKE '").Append(escaped).Append("%'");
                }
                else if (mode == 2)
                {
                    clauses.Append(colExpr).Append(" = '").Append(escaped).Append("'");
                }
                else
                {
                    clauses.Append(colExpr).Append(" LIKE '%").Append(escaped).Append("%'");
                }
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
            FilterLatencyText.Text = $"Filter: {sw.ElapsedMilliseconds} ms";
        }

        private int GetFilterMode()
        {
            if (FilterModeBox?.SelectedIndex is int index && index >= 0) return index;
            return 0;
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

        private void MatchCaseCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            ApplyFilter(FilterBox.Text);
        }

        private void FilterModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;
            ApplyFilter(FilterBox.Text);
        }

        private void ResultsGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString("N0");
        }

        private void CopySelectedButton_Click(object sender, RoutedEventArgs e)
        {
            if (ResultsGrid.Items.Count == 0) return;
            ApplicationCommands.Copy.Execute(null, ResultsGrid);
            StatusText.Text = "Selection copied to clipboard.";
        }

        private void ExportCsvButton_Click(object sender, RoutedEventArgs e)
        {
            if (_view == null || _view.Count == 0) return;

            var dialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                FileName = "results-view.csv",
                AddExtension = true,
                OverwritePrompt = true
            };

            if (dialog.ShowDialog() != true) return;

            var sb = new StringBuilder();
            var headers = _view.Table.Columns.Cast<DataColumn>().Select(c => EscapeCsv(c.ColumnName));
            sb.AppendLine(string.Join(",", headers));

            foreach (DataRowView rowView in _view)
            {
                var cells = rowView.Row.ItemArray.Select(value => EscapeCsv(value?.ToString() ?? string.Empty));
                sb.AppendLine(string.Join(",", cells));
            }

            File.WriteAllText(dialog.FileName, sb.ToString(), Encoding.UTF8);
            StatusText.Text = $"Exported {_view.Count:N0} row{(_view.Count == 1 ? string.Empty : "s")} to CSV.";
        }

        private static string EscapeCsv(string value)
        {
            if (value == null) return string.Empty;
            if (value.Contains("\"")) value = value.Replace("\"", "\"\"");
            if (value.Contains(",") || value.Contains("\n") || value.Contains("\r") || value.Contains("\""))
            {
                return "\"" + value + "\"";
            }
            return value;
        }

        private void CopyDiagnosticsButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_lastDiagnostics)) return;
            Clipboard.SetText(_lastDiagnostics);
            StatusText.Text = "Capture diagnostics copied.";
        }
    }
}
