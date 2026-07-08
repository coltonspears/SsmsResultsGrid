using System;

namespace SsmsResultsGrid.Core.Models
{
    /// <summary>
    /// A single captured result row. The integer indexer is what WPF template
    /// columns bind to (binding path "[0]", "[1]", ...), so it must never throw
    /// for out-of-range indexes — ragged rows simply render empty cells.
    /// </summary>
    public sealed class ResultRow
    {
        private readonly string[] _cells;

        public ResultRow(string[] cells, long sourceRowNumber)
        {
            _cells = cells ?? Array.Empty<string>();
            SourceRowNumber = sourceRowNumber;
        }

        /// <summary>1-based row number in the original result set (stable under filtering/sorting).</summary>
        public long SourceRowNumber { get; }

        public int CellCount => _cells.Length;

        public string this[int index] =>
            index >= 0 && index < _cells.Length ? _cells[index] ?? string.Empty : string.Empty;

        /// <summary>Direct cell access for the filter/sort/export hot paths (no bounds cushioning).</summary>
        public string[] Cells => _cells;
    }
}
