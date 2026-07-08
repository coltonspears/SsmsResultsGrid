using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace SsmsResultsGrid.Services.Settings
{
    /// <summary>
    /// Tools &gt; Options &gt; Results View page. The grid-of-properties UI comes from
    /// DialogPage; persistence delegates to <see cref="ExtensionSettings"/> so the
    /// options page, the menu toggle, and the capture pipeline all share one store
    /// and one change event.
    /// </summary>
    [Guid("8f1c9c2c-4f0a-4e9e-9d7b-1e2a4f3b0008")]
    [ComVisible(true)]
    public sealed class ResultsViewOptionsPage : DialogPage
    {
        private double _gridFontSize = ExtensionSettings.DefaultGridFontSize;
        private bool _alternateRowColors = true;
        private int _maxRows = ExtensionSettings.DefaultMaxRows;
        private int _maxCellChars = ExtensionSettings.DefaultMaxCellChars;
        private bool _resultsToFilterGrid = true;

        [Category("Appearance")]
        [DisplayName("Grid font size")]
        [Description("Font size (in points) for text in the results grid. Allowed range: 8-24.")]
        public double GridFontSize
        {
            get => _gridFontSize;
            set => _gridFontSize = Math.Min(ExtensionSettings.MaxGridFontSize,
                Math.Max(ExtensionSettings.MinGridFontSize, value));
        }

        [Category("Appearance")]
        [DisplayName("Alternate row colors")]
        [Description("Shade every other row with a subtle tint derived from the current theme.")]
        public bool AlternateRowColors
        {
            get => _alternateRowColors;
            set => _alternateRowColors = value;
        }

        [Category("Capture")]
        [DisplayName("Maximum rows")]
        [Description("Maximum number of rows captured per result set. Larger result sets are truncated (the status bar indicates truncation).")]
        public int MaxRows
        {
            get => _maxRows;
            set => _maxRows = Math.Max(1, value);
        }

        [Category("Capture")]
        [DisplayName("Maximum characters per cell")]
        [Description("Upper bound on characters fetched per cell. Lower this to reduce memory usage for very wide tables or large text columns.")]
        public int MaxCellChars
        {
            get => _maxCellChars;
            set => _maxCellChars = Math.Max(256, value);
        }

        [Category("Behavior")]
        [DisplayName("Results to filterable grid")]
        [Description("Automatically show and activate the Results View tab after every query execution.")]
        public bool ResultsToFilterGrid
        {
            get => _resultsToFilterGrid;
            set => _resultsToFilterGrid = value;
        }

        public override void LoadSettingsFromStorage()
        {
            var settings = ExtensionSettings.Instance;
            if (settings == null) return;

            _gridFontSize = settings.GridFontSize;
            _alternateRowColors = settings.AlternateRowColors;
            _maxRows = settings.MaxRows;
            _maxCellChars = settings.MaxCellChars;
            _resultsToFilterGrid = settings.ResultsToFilterGrid;
        }

        public override void SaveSettingsToStorage()
        {
            var settings = ExtensionSettings.Instance;
            if (settings == null) return;

            settings.GridFontSize = _gridFontSize;
            settings.AlternateRowColors = _alternateRowColors;
            settings.MaxRows = _maxRows;
            settings.MaxCellChars = _maxCellChars;
            settings.ResultsToFilterGrid = _resultsToFilterGrid;
        }
    }
}
