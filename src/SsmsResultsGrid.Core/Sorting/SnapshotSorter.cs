using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using SsmsResultsGrid.Core.Models;

namespace SsmsResultsGrid.Core.Sorting
{
    /// <summary>
    /// Sorts a row list by one column on a background thread. When every non-empty
    /// cell in the column parses as a decimal the sort is numeric; otherwise ordinal
    /// string. Empty cells always sort first ascending / last descending.
    /// </summary>
    public static class SnapshotSorter
    {
        public static IReadOnlyList<ResultRow> Sort(
            IReadOnlyList<ResultRow> rows,
            int columnIndex,
            bool descending,
            CancellationToken ct)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (rows.Count <= 1) return rows;

            var copy = new ResultRow[rows.Count];
            for (int i = 0; i < copy.Length; i++) copy[i] = rows[i];

            ct.ThrowIfCancellationRequested();

            var numericKeys = TryBuildNumericKeys(copy, columnIndex, ct);
            ct.ThrowIfCancellationRequested();

            Comparison<ResultRow> comparison;
            if (numericKeys != null)
            {
                comparison = (a, b) => numericKeys[a].CompareTo(numericKeys[b]);
            }
            else
            {
                comparison = (a, b) => string.CompareOrdinal(a[columnIndex], b[columnIndex]);
            }

            if (descending)
            {
                var inner = comparison;
                comparison = (a, b) => inner(b, a);
            }

            Array.Sort(copy, comparison);
            ct.ThrowIfCancellationRequested();
            return copy;
        }

        /// <summary>
        /// Returns per-row numeric keys when the entire column is numeric, else null.
        /// Empty cells are keyed as decimal.MinValue so they group at one end.
        /// </summary>
        private static Dictionary<ResultRow, decimal> TryBuildNumericKeys(
            ResultRow[] rows, int columnIndex, CancellationToken ct)
        {
            var keys = new Dictionary<ResultRow, decimal>(rows.Length);
            for (int i = 0; i < rows.Length; i++)
            {
                if ((i & 0x0FFF) == 0) ct.ThrowIfCancellationRequested();
                var cell = rows[i][columnIndex];
                if (string.IsNullOrEmpty(cell) || cell == "NULL")
                {
                    keys[rows[i]] = decimal.MinValue;
                    continue;
                }
                if (!decimal.TryParse(cell, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
                {
                    return null;
                }
                keys[rows[i]] = value;
            }
            return keys;
        }
    }
}
