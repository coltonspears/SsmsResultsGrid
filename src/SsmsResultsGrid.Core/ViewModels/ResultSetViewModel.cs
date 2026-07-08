using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using SsmsResultsGrid.Core.Export;
using SsmsResultsGrid.Core.Filtering;
using SsmsResultsGrid.Core.Models;
using SsmsResultsGrid.Core.Mvvm;
using SsmsResultsGrid.Core.Resources;
using SsmsResultsGrid.Core.Sorting;

namespace SsmsResultsGrid.Core.ViewModels
{
    /// <summary>
    /// One captured result set. Owns the row store, the debounced/cancellable
    /// background filter pipeline, and sort state. All public members must be
    /// touched on the UI thread; the filter scan itself runs on the threadpool
    /// over an immutable snapshot and publishes back through the dispatcher.
    /// </summary>
    public sealed class ResultSetViewModel : ObservableObject
    {
        public static readonly TimeSpan DefaultFilterDebounce = TimeSpan.FromMilliseconds(200);

        private readonly IUiDispatcher _dispatcher;
        private readonly Action<Exception> _onError;
        private readonly TimeSpan _filterDebounce;
        private readonly List<ResultRow> _rows = new List<ResultRow>();

        private CancellationTokenSource _filterCts;
        private string _filterText = string.Empty;
        private int _filterModeIndex;
        private bool _filterCaseSensitive;
        private string _appliedFilterText = string.Empty;
        private int _appliedFilterModeIndex;
        private bool _appliedFilterCaseSensitive;
        private IReadOnlyList<ResultRow> _visibleRows = Array.Empty<ResultRow>();
        private IReadOnlyList<string> _columnNames = Array.Empty<string>();
        private long _totalRowCount;
        private bool _isTruncated;
        private ResultSetLoadState _loadState = ResultSetLoadState.Pending;
        private int? _sortColumnIndex;
        private bool _sortDescending;
        private string _filterLatencyText = Strings.FilterReady;

        public ResultSetViewModel(
            int gridIndex,
            IUiDispatcher dispatcher,
            Action<Exception> onError,
            TimeSpan? filterDebounce = null)
        {
            GridIndex = gridIndex;
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            _onError = onError ?? (_ => { });
            _filterDebounce = filterDebounce ?? DefaultFilterDebounce;
        }

        public int GridIndex { get; }

        /// <summary>
        /// User-adjusted column widths keyed by "index|name", kept on the view model so
        /// they survive SSMS recreating the results tab (and thus the view) on every run.
        /// </summary>
        public Dictionary<string, double> ColumnWidthMemory { get; } = new Dictionary<string, double>();

        /// <summary>Raised whenever <see cref="ColumnNames"/> is replaced; the view rebuilds grid columns.</summary>
        public event EventHandler ColumnsChanged;

        public IReadOnlyList<string> ColumnNames
        {
            get => _columnNames;
            private set
            {
                _columnNames = value ?? Array.Empty<string>();
                OnPropertyChanged();
                OnPropertyChanged(nameof(SummaryText));
                ColumnsChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public IReadOnlyList<ResultRow> VisibleRows
        {
            get => _visibleRows;
            private set => SetProperty(ref _visibleRows, value);
        }

        public long TotalRowCount
        {
            get => _totalRowCount;
            private set
            {
                if (SetProperty(ref _totalRowCount, value))
                {
                    OnPropertyChanged(nameof(Title));
                }
            }
        }

        public bool IsTruncated
        {
            get => _isTruncated;
            private set => SetProperty(ref _isTruncated, value);
        }

        public ResultSetLoadState LoadState
        {
            get => _loadState;
            private set => SetProperty(ref _loadState, value);
        }

        public long LoadedRowCount => _rows.Count;

        /// <summary>Selector caption, e.g. "Result 2 (12,345 rows)".</summary>
        public string Title =>
            LoadState == ResultSetLoadState.Pending && TotalRowCount == 0
                ? Strings.Format(Strings.ResultSetTitlePending, GridIndex + 1)
                : Strings.Format(Strings.ResultSetTitle, GridIndex + 1, TotalRowCount);

        public string FilterText
        {
            get => _filterText;
            set
            {
                if (SetProperty(ref _filterText, value ?? string.Empty))
                {
                    ScheduleFilter(debounce: true);
                }
            }
        }

        /// <summary>Bound to the mode ComboBox SelectedIndex; values map to <see cref="FilterMode"/>.</summary>
        public int FilterModeIndex
        {
            get => _filterModeIndex;
            set
            {
                if (SetProperty(ref _filterModeIndex, value))
                {
                    ScheduleFilter(debounce: false);
                }
            }
        }

        public bool FilterCaseSensitive
        {
            get => _filterCaseSensitive;
            set
            {
                if (SetProperty(ref _filterCaseSensitive, value))
                {
                    ScheduleFilter(debounce: false);
                }
            }
        }

        /// <summary>Filter text of the last completed scan — drives cell match highlighting so
        /// highlights always agree with the visible row set.</summary>
        public string AppliedFilterText
        {
            get => _appliedFilterText;
            private set => SetProperty(ref _appliedFilterText, value);
        }

        public int AppliedFilterModeIndex
        {
            get => _appliedFilterModeIndex;
            private set => SetProperty(ref _appliedFilterModeIndex, value);
        }

        public bool AppliedFilterCaseSensitive
        {
            get => _appliedFilterCaseSensitive;
            private set => SetProperty(ref _appliedFilterCaseSensitive, value);
        }

        public string FilterLatencyText
        {
            get => _filterLatencyText;
            private set => SetProperty(ref _filterLatencyText, value);
        }

        public string RowSummaryText
        {
            get
            {
                if (IsTruncated && string.IsNullOrEmpty(AppliedFilterText))
                {
                    return Strings.Format(Strings.RowSummaryTruncated, _rows.Count, TotalRowCount);
                }
                if (!string.IsNullOrEmpty(AppliedFilterText))
                {
                    return Strings.Format(Strings.RowSummaryFiltered, VisibleRows.Count, _rows.Count);
                }
                return _rows.Count == 1
                    ? Strings.RowSummaryOne
                    : Strings.Format(Strings.RowSummaryAll, _rows.Count);
            }
        }

        public string SummaryText => Strings.Format(Strings.SummaryFormat, _rows.Count, ColumnNames.Count);

        // ---- load pipeline (called by the capture supervisor through the dispatcher) ----

        /// <summary>Starts (or restarts) a load: resets rows, publishes columns and the expected total.</summary>
        public void BeginLoad(IReadOnlyList<string> columnNames, long totalRowCount)
        {
            CancelActiveFilter();
            _rows.Clear();
            TotalRowCount = totalRowCount;
            IsTruncated = false;
            LoadState = ResultSetLoadState.Loading;
            ColumnNames = columnNames ?? Array.Empty<string>();
            PublishVisibleRows();
        }

        /// <summary>Registers metadata for a lazily-loaded set without fetching rows.</summary>
        public void SetPendingMetadata(IReadOnlyList<string> columnNames, long totalRowCount)
        {
            _rows.Clear();
            TotalRowCount = totalRowCount;
            IsTruncated = false;
            LoadState = ResultSetLoadState.Pending;
            ColumnNames = columnNames ?? Array.Empty<string>();
            PublishVisibleRows();
        }

        public void AppendRows(IReadOnlyList<string[]> rows, long startRow)
        {
            if (rows == null || rows.Count == 0) return;
            _rows.Capacity = Math.Max(_rows.Capacity, _rows.Count + rows.Count);
            for (int i = 0; i < rows.Count; i++)
            {
                _rows.Add(new ResultRow(rows[i], startRow + i + 1));
            }
            PublishVisibleRows();
        }

        public void CompleteLoad(bool truncated)
        {
            IsTruncated = truncated;
            LoadState = ResultSetLoadState.Loaded;
            PublishVisibleRows();
        }

        // ---- sorting (invoked from the view's Sorting event) ----

        public int? SortColumnIndex => _sortColumnIndex;
        public bool SortDescending => _sortDescending;

        public void SetSort(int columnIndex, bool descending)
        {
            _sortColumnIndex = columnIndex;
            _sortDescending = descending;
            ScheduleFilter(debounce: false);
        }

        public void ClearSort()
        {
            _sortColumnIndex = null;
            _sortDescending = false;
            ScheduleFilter(debounce: false);
        }

        // ---- export ----

        public async Task<int> ExportVisibleRowsAsync(string path, CancellationToken ct)
        {
            var columns = ColumnNames;
            var rows = VisibleRows;
            await Task.Run(() => CsvWriter.WriteFile(path, columns, rows, ct), ct).ConfigureAwait(false);
            return rows.Count;
        }

        // ---- filter pipeline ----

        /// <summary>Snapshots state on the UI thread, then filters/sorts on the threadpool.</summary>
        private void ScheduleFilter(bool debounce) => PublishVisibleRows(debounce);

        private void PublishVisibleRows(bool debounce = false)
        {
            CancelActiveFilter();

            var request = new FilterRequest(_filterText, (FilterMode)_filterModeIndex, _filterCaseSensitive);
            var sortColumn = _sortColumnIndex;
            var sortDescending = _sortDescending;

            // Fast path: nothing to scan or sort — publish the snapshot synchronously.
            if (request.IsEmpty && sortColumn == null)
            {
                VisibleRows = _rows.ToArray();
                ApplyCompletedFilter(request, 0);
                return;
            }

            var cts = new CancellationTokenSource();
            _filterCts = cts;
            var snapshot = _rows.ToArray();
            _ = RunFilterAsync(snapshot, request, sortColumn, sortDescending, _columnNames.Count, debounce, cts.Token);
        }

        private async Task RunFilterAsync(
            ResultRow[] snapshot,
            FilterRequest request,
            int? sortColumn,
            bool sortDescending,
            int columnCount,
            bool debounce,
            CancellationToken ct)
        {
            try
            {
                if (debounce)
                {
                    await Task.Delay(_filterDebounce, ct).ConfigureAwait(false);
                }

                var stopwatch = Stopwatch.StartNew();
                var visible = await Task.Run(() =>
                {
                    var filtered = RowFilterEngine.Filter(snapshot, request, ct);
                    if (sortColumn.HasValue && sortColumn.Value < columnCount)
                    {
                        filtered = SnapshotSorter.Sort(filtered, sortColumn.Value, sortDescending, ct);
                    }
                    return filtered;
                }, ct).ConfigureAwait(false);

                ct.ThrowIfCancellationRequested();
                stopwatch.Stop();

                _dispatcher.Post(() =>
                {
                    if (ct.IsCancellationRequested) return;
                    VisibleRows = visible;
                    ApplyCompletedFilter(request, stopwatch.ElapsedMilliseconds);
                });
            }
            catch (OperationCanceledException)
            {
                // Superseded by a newer filter run — expected.
            }
            catch (Exception ex)
            {
                _onError(ex);
            }
        }

        private void ApplyCompletedFilter(FilterRequest request, long elapsedMs)
        {
            AppliedFilterText = request.Text;
            AppliedFilterModeIndex = (int)request.Mode;
            AppliedFilterCaseSensitive = request.CaseSensitive;
            FilterLatencyText = Strings.Format(Strings.FilterLatency, elapsedMs);
            OnPropertyChanged(nameof(RowSummaryText));
            OnPropertyChanged(nameof(SummaryText));
        }

        private void CancelActiveFilter()
        {
            var cts = _filterCts;
            _filterCts = null;
            if (cts != null)
            {
                cts.Cancel();
                cts.Dispose();
            }
        }
    }
}
