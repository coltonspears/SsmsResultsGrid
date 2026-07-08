using System.Globalization;
using System.Resources;

namespace SsmsResultsGrid.Resources
{
    /// <summary>
    /// Hand-written resource accessor (msbuild does not run the resx code generator
    /// outside Visual Studio). Public so XAML can reference members via x:Static.
    /// </summary>
    public static class Strings
    {
        private static readonly ResourceManager Rm =
            new ResourceManager("SsmsResultsGrid.Resources.Strings", typeof(Strings).Assembly);

        private static string Get(string name) => Rm.GetString(name, CultureInfo.CurrentUICulture) ?? name;

        public static string TabTitle => Get(nameof(TabTitle));
        public static string OutputPaneTitle => Get(nameof(OutputPaneTitle));
        public static string ToolWindowCaption => Get(nameof(ToolWindowCaption));
        public static string ViewTitle => Get(nameof(ViewTitle));
        public static string FilterPlaceholder => Get(nameof(FilterPlaceholder));
        public static string FilterBoxTooltip => Get(nameof(FilterBoxTooltip));
        public static string FilterModeTooltip => Get(nameof(FilterModeTooltip));
        public static string FilterModeContains => Get(nameof(FilterModeContains));
        public static string FilterModeStartsWith => Get(nameof(FilterModeStartsWith));
        public static string FilterModeExact => Get(nameof(FilterModeExact));
        public static string MatchCase => Get(nameof(MatchCase));
        public static string MatchCaseTooltip => Get(nameof(MatchCaseTooltip));
        public static string ClearFilterTooltip => Get(nameof(ClearFilterTooltip));
        public static string Refresh => Get(nameof(Refresh));
        public static string RefreshTooltip => Get(nameof(RefreshTooltip));
        public static string Copy => Get(nameof(Copy));
        public static string CopyTooltip => Get(nameof(CopyTooltip));
        public static string ExportCsv => Get(nameof(ExportCsv));
        public static string ExportCsvTooltip => Get(nameof(ExportCsvTooltip));
        public static string ResultSetSelectorTooltip => Get(nameof(ResultSetSelectorTooltip));
        public static string CsvDialogFilter => Get(nameof(CsvDialogFilter));
        public static string CsvDefaultFileName => Get(nameof(CsvDefaultFileName));
        public static string CmdToggleResultsToGrid => Get(nameof(CmdToggleResultsToGrid));
        public static string CmdShowResultsView => Get(nameof(CmdShowResultsView));
    }
}
