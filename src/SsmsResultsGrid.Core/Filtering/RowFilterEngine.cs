using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using SsmsResultsGrid.Core.Models;

namespace SsmsResultsGrid.Core.Filtering
{
    /// <summary>
    /// Cancellable full-scan filter over an immutable row snapshot. A row matches
    /// when any cell matches the request. Designed to run on a threadpool thread;
    /// never touches UI state.
    /// </summary>
    public static class RowFilterEngine
    {
        /// <summary>Row count above which the scan is parallelized.</summary>
        internal const int ParallelThreshold = 50_000;

        /// <summary>Rows scanned between cancellation checks on the sequential path.</summary>
        private const int CancellationStride = 4_096;

        public static IReadOnlyList<ResultRow> Filter(ResultRow[] rows, FilterRequest request, CancellationToken ct)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (request == null) throw new ArgumentNullException(nameof(request));

            if (request.IsEmpty || rows.Length == 0)
            {
                return rows;
            }

            return rows.Length >= ParallelThreshold
                ? FilterParallel(rows, request, ct)
                : FilterSequential(rows, request, ct);
        }

        private static IReadOnlyList<ResultRow> FilterSequential(ResultRow[] rows, FilterRequest request, CancellationToken ct)
        {
            var result = new List<ResultRow>();
            for (int i = 0; i < rows.Length; i++)
            {
                if ((i & (CancellationStride - 1)) == 0) ct.ThrowIfCancellationRequested();
                if (RowMatches(rows[i], request)) result.Add(rows[i]);
            }
            return result;
        }

        private static IReadOnlyList<ResultRow> FilterParallel(ResultRow[] rows, FilterRequest request, CancellationToken ct)
        {
            var matches = new bool[rows.Length];
            try
            {
                Parallel.For(
                    0,
                    rows.Length,
                    new ParallelOptions { CancellationToken = ct },
                    i => matches[i] = RowMatches(rows[i], request));
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException oce)
            {
                throw oce;
            }

            ct.ThrowIfCancellationRequested();

            var result = new List<ResultRow>();
            for (int i = 0; i < rows.Length; i++)
            {
                if (matches[i]) result.Add(rows[i]);
            }
            return result;
        }

        internal static bool RowMatches(ResultRow row, FilterRequest request)
        {
            var cells = row.Cells;
            var text = request.Text;
            var comparison = request.Comparison;

            for (int c = 0; c < cells.Length; c++)
            {
                var cell = cells[c];
                if (string.IsNullOrEmpty(cell)) continue;

                switch (request.Mode)
                {
                    case FilterMode.StartsWith:
                        if (cell.StartsWith(text, comparison)) return true;
                        break;
                    case FilterMode.Exact:
                        if (cell.Equals(text, comparison)) return true;
                        break;
                    default:
                        if (cell.IndexOf(text, comparison) >= 0) return true;
                        break;
                }
            }
            return false;
        }
    }
}
