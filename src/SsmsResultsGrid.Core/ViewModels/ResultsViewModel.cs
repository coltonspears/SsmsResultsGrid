using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Threading.Tasks;
using SsmsResultsGrid.Core.Mvvm;
using SsmsResultsGrid.Core.Resources;

namespace SsmsResultsGrid.Core.ViewModels
{
    /// <summary>
    /// Per-query-window root view model. Owns the result-set collection and
    /// capture-lifecycle status. It survives SSMS recreating the results tab strip
    /// on every execution, which is how filter state persists across F5 runs.
    /// All members are UI-thread affine; the capture supervisor marshals in via
    /// the dispatcher.
    /// </summary>
    public sealed class ResultsViewModel : ObservableObject
    {
        private readonly IUiDispatcher _dispatcher;
        private readonly Action<Exception> _onError;
        private readonly TimeSpan? _filterDebounce;

        private ResultSetViewModel _selectedResultSet;
        private string _statusText = Strings.StatusNoResults;
        private string _lastUpdatedText = Strings.UpdatedNever;
        private bool _isCapturing;
        private Func<Task> _refreshHandler;

        public ResultsViewModel(IUiDispatcher dispatcher, Action<Exception> onError, TimeSpan? filterDebounce = null)
        {
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            _onError = onError ?? (_ => { });
            _filterDebounce = filterDebounce;
            RefreshCommand = new AsyncRelayCommand(
                () => _refreshHandler?.Invoke() ?? Task.CompletedTask,
                canExecute: () => _refreshHandler != null && !IsCapturing,
                onError: _onError);
        }

        public ObservableCollection<ResultSetViewModel> ResultSets { get; } =
            new ObservableCollection<ResultSetViewModel>();

        /// <summary>Raised when a pending (lazily-loaded) result set is selected and needs its rows fetched.</summary>
        public event EventHandler<ResultSetViewModel> ResultSetLoadRequested;

        public ResultSetViewModel SelectedResultSet
        {
            get => _selectedResultSet;
            set
            {
                if (SetProperty(ref _selectedResultSet, value))
                {
                    OnPropertyChanged(nameof(HasSelectedResultSet));
                    if (value != null && value.LoadState == ResultSetLoadState.Pending)
                    {
                        ResultSetLoadRequested?.Invoke(this, value);
                    }
                }
            }
        }

        public bool HasSelectedResultSet => _selectedResultSet != null;

        public bool HasMultipleResultSets => ResultSets.Count > 1;

        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        public string LastUpdatedText
        {
            get => _lastUpdatedText;
            private set => SetProperty(ref _lastUpdatedText, value);
        }

        public bool IsCapturing
        {
            get => _isCapturing;
            private set
            {
                if (SetProperty(ref _isCapturing, value))
                {
                    RefreshCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public AsyncRelayCommand RefreshCommand { get; }

        /// <summary>The supervisor installs the actual re-capture routine here.</summary>
        public void SetRefreshHandler(Func<Task> handler)
        {
            _refreshHandler = handler;
            RefreshCommand.RaiseCanExecuteChanged();
        }

        // ---- capture lifecycle (called by the supervisor on the UI thread) ----

        /// <summary>Aligns the result-set collection with a new execution's grid count and resets selection.</summary>
        public void BeginCapture(int gridCount)
        {
            IsCapturing = true;
            StatusText = Strings.StatusCapturing;

            while (ResultSets.Count > gridCount)
            {
                ResultSets.RemoveAt(ResultSets.Count - 1);
            }
            while (ResultSets.Count < gridCount)
            {
                ResultSets.Add(new ResultSetViewModel(ResultSets.Count, _dispatcher, _onError, _filterDebounce));
            }
            OnPropertyChanged(nameof(HasMultipleResultSets));

            if (SelectedResultSet == null || !ResultSets.Contains(SelectedResultSet))
            {
                _selectedResultSet = ResultSets.Count > 0 ? ResultSets[0] : null;
                OnPropertyChanged(nameof(SelectedResultSet));
                OnPropertyChanged(nameof(HasSelectedResultSet));
            }
        }

        public ResultSetViewModel GetResultSet(int gridIndex) =>
            gridIndex >= 0 && gridIndex < ResultSets.Count ? ResultSets[gridIndex] : null;

        public void ReportLoadProgress(long loadedRows, long totalRows)
        {
            StatusText = Strings.Format(Strings.StatusLoadingRows, loadedRows, totalRows);
        }

        public void CompleteCapture(DateTime timestamp)
        {
            IsCapturing = false;
            StatusText = SelectedResultSet?.RowSummaryText ?? Strings.StatusNoResults;
            LastUpdatedText = Strings.Format(Strings.UpdatedAt, timestamp.ToString("T", CultureInfo.CurrentCulture));
        }

        public void SetNoGridResults()
        {
            IsCapturing = false;
            ResultSets.Clear();
            _selectedResultSet = null;
            OnPropertyChanged(nameof(SelectedResultSet));
            OnPropertyChanged(nameof(HasSelectedResultSet));
            OnPropertyChanged(nameof(HasMultipleResultSets));
            StatusText = Strings.StatusNoGridResults;
        }

        public void SetCaptureFailed()
        {
            IsCapturing = false;
            StatusText = Strings.StatusCaptureFailed;
        }
    }
}
