using System;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;

namespace SsmsResultsGrid.Services.Settings
{
    /// <summary>
    /// Persisted user settings backed by the VS WritableSettingsStore.
    /// Values load once at package init (main thread) and write through immediately.
    /// </summary>
    internal sealed class ExtensionSettings
    {
        private const string CollectionPath = "SsmsResultsGrid";
        private const string ResultsToFilterGridName = "ResultsToFilterGrid";
        private const string MaxRowsName = "MaxRows";
        private const string MaxCellCharsName = "MaxCellChars";

        public const int DefaultMaxRows = 100_000;
        public const int DefaultMaxCellChars = 65_535;

        private readonly WritableSettingsStore _store;
        private bool _resultsToFilterGrid = true;
        private int _maxRows = DefaultMaxRows;
        private int _maxCellChars = DefaultMaxCellChars;

        public ExtensionSettings(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var manager = new ShellSettingsManager(serviceProvider);
            _store = manager.GetWritableSettingsStore(SettingsScope.UserSettings);
            Load();
        }

        public event EventHandler Changed;

        /// <summary>
        /// The "Results To" mode: when true, every query completion injects and
        /// activates the filterable Results View tab; when false nothing is injected
        /// automatically (the Tools-menu command still works on demand).
        /// </summary>
        public bool ResultsToFilterGrid
        {
            get => _resultsToFilterGrid;
            set
            {
                if (_resultsToFilterGrid == value) return;
                _resultsToFilterGrid = value;
                SaveBool(ResultsToFilterGridName, value);
                Changed?.Invoke(this, EventArgs.Empty);
            }
        }

        /// <summary>Maximum rows captured per result set before truncating.</summary>
        public int MaxRows => _maxRows;

        /// <summary>Maximum characters fetched per cell.</summary>
        public int MaxCellChars => _maxCellChars;

        private void Load()
        {
            try
            {
                if (_store == null || !_store.CollectionExists(CollectionPath)) return;
                _resultsToFilterGrid = _store.GetBoolean(CollectionPath, ResultsToFilterGridName, true);
                _maxRows = Math.Max(1, _store.GetInt32(CollectionPath, MaxRowsName, DefaultMaxRows));
                _maxCellChars = Math.Max(256, _store.GetInt32(CollectionPath, MaxCellCharsName, DefaultMaxCellChars));
            }
            catch
            {
                // Corrupt store — keep defaults.
            }
        }

        private void SaveBool(string name, bool value)
        {
            try
            {
                if (_store == null) return;
                if (!_store.CollectionExists(CollectionPath)) _store.CreateCollection(CollectionPath);
                _store.SetBoolean(CollectionPath, name, value);
            }
            catch
            {
                // Persisting is best-effort; the in-memory value still applies this session.
            }
        }
    }
}
