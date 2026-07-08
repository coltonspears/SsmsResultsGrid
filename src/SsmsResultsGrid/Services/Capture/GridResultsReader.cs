using System;
using System.Threading;
using System.Threading.Tasks;
using SsmsResultsGrid.Core.Models;

namespace SsmsResultsGrid.Services.Capture
{
    /// <summary>
    /// Pages a result grid out of the brokered service in fixed-size chunks,
    /// reporting each page through <see cref="IProgress{T}"/> so the UI fills
    /// progressively. Runs entirely off the UI thread.
    /// </summary>
    internal static class GridResultsReader
    {
        public const int RowsPerPage = 5_000;

        /// <summary>
        /// Display/capture cap on columns. SQL Server allows up to 4,096 select-list
        /// columns but a WPF DataGrid stops being useful long before that.
        /// </summary>
        public const int MaxColumns = 512;

        /// <summary>
        /// Reads grid metadata only (columns + row/grid counts) via a single 1-row request.
        /// Returns null when the grid does not exist.
        /// </summary>
        public static Task<GridSegment> ProbeGridAsync(
            GridBrokeredClient client, int gridIndex, int maxCellChars, CancellationToken ct)
        {
            return client.GetGridSegmentAsync(
                gridIndex, startColumn: 0, columnCount: MaxColumns, startRow: 0,
                maxRows: 1, maxCellChars: maxCellChars, ct: ct);
        }

        /// <summary>
        /// Reads a whole grid up to <paramref name="maxRows"/>. The first batch carries
        /// column names; the final batch is flagged, with truncation state.
        /// </summary>
        public static async Task ReadGridAsync(
            GridBrokeredClient client,
            int gridIndex,
            int maxRows,
            int maxCellChars,
            IProgress<CapturedBatch> progress,
            CancellationToken ct)
        {
            if (client == null) throw new ArgumentNullException(nameof(client));
            if (progress == null) throw new ArgumentNullException(nameof(progress));

            long startRow = 0;
            bool first = true;
            long totalRowCount;
            int totalGridCount;
            int columnCount;

            // Metadata probe: learn the real column/row counts before paging.
            var probe = await ProbeGridAsync(client, gridIndex, maxCellChars, ct).ConfigureAwait(false)
                ?? throw new InvalidOperationException(
                    $"GetGridResultsSegmentAsync returned null for grid {gridIndex}.");

            totalRowCount = probe.TotalRowCount;
            totalGridCount = probe.TotalGridCount;
            columnCount = Math.Min(Math.Max(probe.TotalColumnCount, probe.ColumnNames.Count), MaxColumns);
            long captureTarget = Math.Min(totalRowCount, maxRows);

            if (captureTarget <= 0)
            {
                progress.Report(new CapturedBatch(
                    gridIndex, totalGridCount, totalRowCount, probe.ColumnNames,
                    rows: Array.Empty<string[]>(), startRow: 0,
                    isFinal: true, isTruncated: totalRowCount > 0));
                return;
            }

            while (true)
            {
                ct.ThrowIfCancellationRequested();

                int want = (int)Math.Min(RowsPerPage, captureTarget - startRow);
                var segment = await client.GetGridSegmentAsync(
                    gridIndex, startColumn: 0, columnCount: columnCount, startRow: startRow,
                    maxRows: want, maxCellChars: maxCellChars, ct: ct)
                    .ConfigureAwait(false);

                var rows = segment?.Rows ?? new System.Collections.Generic.List<string[]>();
                long nextRow = startRow + rows.Count;
                bool exhausted = rows.Count == 0 || nextRow >= captureTarget;
                bool truncated = exhausted && nextRow < totalRowCount;

                progress.Report(new CapturedBatch(
                    gridIndex,
                    totalGridCount,
                    totalRowCount,
                    columnNames: first ? (segment?.ColumnNames ?? probe.ColumnNames) : null,
                    rows: rows,
                    startRow: startRow,
                    isFinal: exhausted,
                    isTruncated: truncated));

                if (exhausted) return;

                startRow = nextRow;
                first = false;
            }
        }
    }
}
