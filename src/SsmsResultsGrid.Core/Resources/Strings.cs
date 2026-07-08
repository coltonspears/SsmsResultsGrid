using System.Globalization;
using System.Resources;

namespace SsmsResultsGrid.Core.Resources
{
    /// <summary>
    /// Hand-written resource accessor (msbuild does not run the resx code generator
    /// outside Visual Studio). Satellite assemblies localize by convention.
    /// </summary>
    public static class Strings
    {
        private static readonly ResourceManager Rm =
            new ResourceManager("SsmsResultsGrid.Core.Resources.Strings", typeof(Strings).Assembly);

        private static string Get(string name) => Rm.GetString(name, CultureInfo.CurrentUICulture) ?? name;

        public static string StatusNoResults => Get(nameof(StatusNoResults));
        public static string StatusCapturing => Get(nameof(StatusCapturing));
        public static string StatusLoadingRows => Get(nameof(StatusLoadingRows));
        public static string StatusNoGridResults => Get(nameof(StatusNoGridResults));
        public static string StatusCaptureFailed => Get(nameof(StatusCaptureFailed));
        public static string RowSummaryAll => Get(nameof(RowSummaryAll));
        public static string RowSummaryOne => Get(nameof(RowSummaryOne));
        public static string RowSummaryFiltered => Get(nameof(RowSummaryFiltered));
        public static string RowSummaryTruncated => Get(nameof(RowSummaryTruncated));
        public static string SummaryFormat => Get(nameof(SummaryFormat));
        public static string FilterReady => Get(nameof(FilterReady));
        public static string FilterLatency => Get(nameof(FilterLatency));
        public static string UpdatedAt => Get(nameof(UpdatedAt));
        public static string UpdatedNever => Get(nameof(UpdatedNever));
        public static string ResultSetTitle => Get(nameof(ResultSetTitle));
        public static string ResultSetTitlePending => Get(nameof(ResultSetTitlePending));
        public static string ExportComplete => Get(nameof(ExportComplete));
        public static string CopyComplete => Get(nameof(CopyComplete));

        public static string Format(string format, params object[] args) =>
            string.Format(CultureInfo.CurrentCulture, format, args);
    }
}
