using System;

namespace SsmsResultsGrid.Core.Filtering
{
    /// <summary>Immutable description of one filter run.</summary>
    public sealed class FilterRequest
    {
        public FilterRequest(string text, FilterMode mode, bool caseSensitive)
        {
            Text = text ?? string.Empty;
            Mode = mode;
            CaseSensitive = caseSensitive;
        }

        public string Text { get; }
        public FilterMode Mode { get; }
        public bool CaseSensitive { get; }

        public bool IsEmpty => Text.Length == 0;

        public StringComparison Comparison =>
            CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
    }
}
