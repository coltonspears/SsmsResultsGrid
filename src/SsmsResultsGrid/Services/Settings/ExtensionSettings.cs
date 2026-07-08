using System;
using System.Globalization;
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
        private const string GridFontSizeName = "GridFontSize";
        private const string AlternateRowColorsName = "AlternateRowColors";

        public const int DefaultMaxRows = 100_000;
        public const int DefaultMaxCellChars = 65_535;
        public const double DefaultGridFontSize = 12.0;
        public const double MinGridFontSize = 8.0;
        public const double MaxGridFontSize = 24.0;

        /// <summary>Set once at package init so the Tools &gt; Options page can reach the store.</summary>
        public static ExtensionSettings Instance { get; private set; }

        private readonly WritableSettingsStore _store;
        private bool _resultsToFilterGrid = true;
        private int _maxRows = DefaultMaxRows;
        private int _maxCellChars = DefaultMaxCellChars;
        private double _gridFontSize = DefaultGridFontSize;
        private bool _alternateRowColors = true;

        public ExtensionSettings(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var manager = new ShellSettingsManager(serviceProvider);
            _store = manager.GetWritableSettingsStore(SettingsScope.UserSettings);
            Load();
            Instance = this;
        }

        public event EventHandler Changed;

        /// <summary>
        /// The "Results To" mode: when true, every query completion injects and
        /// activates the filterable Results View tab; when false nothing is injected
        /// automatically (the menu command still works on demand).
        /// </summary>
        public bool ResultsToFilterGrid
        {
            get => _resultsToFilterGrid;
            set
            {
                if (_resultsToFilterGrid == value) return;
                _resultsToFilterGrid = value;
                Save(ResultsToFilterGridName, value);
                RaiseChanged();
            }
        }

        /// <summary>Maximum rows captured per result set before truncating.</summary>
        public int MaxRows
        {
            get => _maxRows;
            set
            {
                value = Math.Max(1, value);
                if (_maxRows == value) return;
                _maxRows = value;
                Save(MaxRowsName, value);
                RaiseChanged();
            }
        }

        /// <summary>Maximum characters fetched per cell (bounds data size for wide tables).</summary>
        public int MaxCellChars
        {
            get => _maxCellChars;
            set
            {
                value = Math.Max(256, value);
                if (_maxCellChars == value) return;
                _maxCellChars = value;
                Save(MaxCellCharsName, value);
                RaiseChanged();
            }
        }

        /// <summary>Font size for the results grid text.</summary>
        public double GridFontSize
        {
            get => _gridFontSize;
            set
            {
                value = Math.Min(MaxGridFontSize, Math.Max(MinGridFontSize, value));
                if (Math.Abs(_gridFontSize - value) < 0.01) return;
                _gridFontSize = value;
                Save(GridFontSizeName, value.ToString(CultureInfo.InvariantCulture));
                RaiseChanged();
            }
        }

        /// <summary>Zebra-stripe alternating rows (theme-derived tint).</summary>
        public bool AlternateRowColors
        {
            get => _alternateRowColors;
            set
            {
                if (_alternateRowColors == value) return;
                _alternateRowColors = value;
                Save(AlternateRowColorsName, value);
                RaiseChanged();
            }
        }

        private void RaiseChanged() => Changed?.Invoke(this, EventArgs.Empty);

        private void Load()
        {
            try
            {
                if (_store == null || !_store.CollectionExists(CollectionPath)) return;
                _resultsToFilterGrid = _store.GetBoolean(CollectionPath, ResultsToFilterGridName, true);
                _maxRows = Math.Max(1, _store.GetInt32(CollectionPath, MaxRowsName, DefaultMaxRows));
                _maxCellChars = Math.Max(256, _store.GetInt32(CollectionPath, MaxCellCharsName, DefaultMaxCellChars));
                _alternateRowColors = _store.GetBoolean(CollectionPath, AlternateRowColorsName, true);

                var fontRaw = _store.GetString(CollectionPath, GridFontSizeName, string.Empty);
                if (double.TryParse(fontRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out var size))
                {
                    _gridFontSize = Math.Min(MaxGridFontSize, Math.Max(MinGridFontSize, size));
                }
            }
            catch
            {
                // Corrupt store — keep defaults.
            }
        }

        private void Save(string name, bool value)
        {
            if (!EnsureCollection()) return;
            try { _store.SetBoolean(CollectionPath, name, value); } catch { /* best-effort */ }
        }

        private void Save(string name, int value)
        {
            if (!EnsureCollection()) return;
            try { _store.SetInt32(CollectionPath, name, value); } catch { /* best-effort */ }
        }

        private void Save(string name, string value)
        {
            if (!EnsureCollection()) return;
            try { _store.SetString(CollectionPath, name, value); } catch { /* best-effort */ }
        }

        private bool EnsureCollection()
        {
            try
            {
                if (_store == null) return false;
                if (!_store.CollectionExists(CollectionPath)) _store.CreateCollection(CollectionPath);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
