using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace SsmsResultsGrid.ToolWindows
{
    /// <summary>
    /// Attached helpers that turn a TextBlock into a match-highlighting cell renderer.
    /// Bind <see cref="SourceTextProperty"/> to the cell value and <see cref="HighlightTextProperty"/>
    /// to the active filter query; runs matching the query are visually highlighted.
    /// </summary>
    public static class TextHighlight
    {
        public static readonly DependencyProperty SourceTextProperty =
            DependencyProperty.RegisterAttached(
                "SourceText",
                typeof(string),
                typeof(TextHighlight),
                new PropertyMetadata(string.Empty, OnAnyChanged));

        public static readonly DependencyProperty HighlightTextProperty =
            DependencyProperty.RegisterAttached(
                "HighlightText",
                typeof(string),
                typeof(TextHighlight),
                new PropertyMetadata(string.Empty, OnAnyChanged));

        public static readonly DependencyProperty MatchModeProperty =
            DependencyProperty.RegisterAttached(
                "MatchMode",
                typeof(int),
                typeof(TextHighlight),
                new PropertyMetadata(0, OnAnyChanged));

        public static readonly DependencyProperty CaseSensitiveProperty =
            DependencyProperty.RegisterAttached(
                "CaseSensitive",
                typeof(bool),
                typeof(TextHighlight),
                new PropertyMetadata(false, OnAnyChanged));

        public static string GetSourceText(DependencyObject d) => (string)d.GetValue(SourceTextProperty);
        public static void SetSourceText(DependencyObject d, string v) => d.SetValue(SourceTextProperty, v);

        public static string GetHighlightText(DependencyObject d) => (string)d.GetValue(HighlightTextProperty);
        public static void SetHighlightText(DependencyObject d, string v) => d.SetValue(HighlightTextProperty, v);

        public static int GetMatchMode(DependencyObject d) => (int)d.GetValue(MatchModeProperty);
        public static void SetMatchMode(DependencyObject d, int v) => d.SetValue(MatchModeProperty, v);

        public static bool GetCaseSensitive(DependencyObject d) => (bool)d.GetValue(CaseSensitiveProperty);
        public static void SetCaseSensitive(DependencyObject d, bool v) => d.SetValue(CaseSensitiveProperty, v);

        private static void OnAnyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TextBlock tb) Refresh(tb);
        }

        private static void Refresh(TextBlock tb)
        {
            var source = GetSourceText(tb) ?? string.Empty;
            var query = GetHighlightText(tb) ?? string.Empty;
            var mode = GetMatchMode(tb);
            var caseSensitive = GetCaseSensitive(tb);

            tb.Inlines.Clear();

            if (string.IsNullOrEmpty(source))
            {
                return;
            }

            if (string.IsNullOrEmpty(query))
            {
                tb.Inlines.Add(new Run(source));
                return;
            }

            var cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            switch (mode)
            {
                case 2: // Exact
                    if (source.Equals(query, cmp))
                    {
                        tb.Inlines.Add(BuildHighlightRun(tb, source));
                    }
                    else
                    {
                        tb.Inlines.Add(new Run(source));
                    }
                    return;

                case 1: // Starts With
                    if (source.StartsWith(query, cmp))
                    {
                        tb.Inlines.Add(BuildHighlightRun(tb, source.Substring(0, query.Length)));
                        if (source.Length > query.Length)
                        {
                            tb.Inlines.Add(new Run(source.Substring(query.Length)));
                        }
                    }
                    else
                    {
                        tb.Inlines.Add(new Run(source));
                    }
                    return;

                default: // Contains
                    AppendContainsRuns(tb, source, query, cmp);
                    return;
            }
        }

        private static void AppendContainsRuns(TextBlock tb, string source, string query, StringComparison cmp)
        {
            if (query.Length == 0)
            {
                tb.Inlines.Add(new Run(source));
                return;
            }

            int idx = 0;
            while (idx <= source.Length)
            {
                int found = source.IndexOf(query, idx, cmp);
                if (found < 0)
                {
                    if (idx < source.Length)
                    {
                        tb.Inlines.Add(new Run(source.Substring(idx)));
                    }
                    break;
                }
                if (found > idx)
                {
                    tb.Inlines.Add(new Run(source.Substring(idx, found - idx)));
                }
                tb.Inlines.Add(BuildHighlightRun(tb, source.Substring(found, query.Length)));
                idx = found + query.Length;
            }
        }

        private static Run BuildHighlightRun(TextBlock host, string text)
        {
            var run = new Run(text) { FontWeight = FontWeights.SemiBold };
            if (host.TryFindResource("Match.Background") is Brush bg)
            {
                run.Background = bg;
            }
            if (host.TryFindResource("Match.Foreground") is Brush fg)
            {
                run.Foreground = fg;
            }
            return run;
        }
    }
}
