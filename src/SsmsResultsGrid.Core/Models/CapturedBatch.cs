using System.Collections.Generic;

namespace SsmsResultsGrid.Core.Models
{
    /// <summary>
    /// One page of rows captured from a result grid, reported progressively while
    /// a capture is in flight. The first batch of a grid carries the column names.
    /// </summary>
    public sealed class CapturedBatch
    {
        public CapturedBatch(
            int gridIndex,
            int totalGridCount,
            long totalRowCount,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<string[]> rows,
            long startRow,
            bool isFinal,
            bool isTruncated)
        {
            GridIndex = gridIndex;
            TotalGridCount = totalGridCount;
            TotalRowCount = totalRowCount;
            ColumnNames = columnNames;
            Rows = rows;
            StartRow = startRow;
            IsFinal = isFinal;
            IsTruncated = isTruncated;
        }

        public int GridIndex { get; }
        public int TotalGridCount { get; }

        /// <summary>Total rows the source grid holds (may exceed what will be captured).</summary>
        public long TotalRowCount { get; }

        /// <summary>Column headers; only populated on the first batch of a grid, null afterwards.</summary>
        public IReadOnlyList<string> ColumnNames { get; }

        public IReadOnlyList<string[]> Rows { get; }

        /// <summary>0-based index of the first row in <see cref="Rows"/> within the source grid.</summary>
        public long StartRow { get; }

        /// <summary>True when this is the last batch for the grid.</summary>
        public bool IsFinal { get; }

        /// <summary>True when capture stopped at the row cap before reaching TotalRowCount.</summary>
        public bool IsTruncated { get; }
    }
}
