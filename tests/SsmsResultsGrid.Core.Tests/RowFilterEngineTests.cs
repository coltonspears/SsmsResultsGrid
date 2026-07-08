using System;
using System.Linq;
using System.Threading;
using SsmsResultsGrid.Core.Filtering;
using Xunit;

namespace SsmsResultsGrid.Core.Tests
{
    public class RowFilterEngineTests
    {
        [Fact]
        public void EmptyFilter_ReturnsAllRows()
        {
            var rows = Rows.Make(new[] { "a", "b" }, new[] { "c", "d" });
            var result = RowFilterEngine.Filter(rows, new FilterRequest("", FilterMode.Contains, false), CancellationToken.None);
            Assert.Equal(2, result.Count);
        }

        [Fact]
        public void Contains_MatchesAnyColumn_CaseInsensitiveByDefault()
        {
            var rows = Rows.Make(
                new[] { "Alpha", "one" },
                new[] { "beta", "TWO" },
                new[] { "gamma", "three" });

            var result = RowFilterEngine.Filter(rows, new FilterRequest("two", FilterMode.Contains, false), CancellationToken.None);

            Assert.Single(result);
            Assert.Equal("beta", result[0][0]);
        }

        [Fact]
        public void Contains_CaseSensitive_RespectsCase()
        {
            var rows = Rows.Make(new[] { "Alpha" }, new[] { "alpha" });

            var sensitive = RowFilterEngine.Filter(rows, new FilterRequest("Alpha", FilterMode.Contains, true), CancellationToken.None);
            var insensitive = RowFilterEngine.Filter(rows, new FilterRequest("Alpha", FilterMode.Contains, false), CancellationToken.None);

            Assert.Single(sensitive);
            Assert.Equal(2, insensitive.Count);
        }

        [Fact]
        public void StartsWith_OnlyMatchesPrefix()
        {
            var rows = Rows.Make(new[] { "prefix-value" }, new[] { "value-prefix" });
            var result = RowFilterEngine.Filter(rows, new FilterRequest("prefix", FilterMode.StartsWith, false), CancellationToken.None);
            Assert.Single(result);
            Assert.Equal("prefix-value", result[0][0]);
        }

        [Fact]
        public void Exact_RequiresFullCellMatch()
        {
            var rows = Rows.Make(new[] { "abc" }, new[] { "abcd" });
            var result = RowFilterEngine.Filter(rows, new FilterRequest("abc", FilterMode.Exact, false), CancellationToken.None);
            Assert.Single(result);
            Assert.Equal("abc", result[0][0]);
        }

        [Fact]
        public void NullAndEmptyCells_NeverMatch_AndNeverThrow()
        {
            var rows = Rows.Make(new[] { (string)null, "" }, new[] { "match", null });
            var result = RowFilterEngine.Filter(rows, new FilterRequest("match", FilterMode.Contains, false), CancellationToken.None);
            Assert.Single(result);
        }

        [Fact]
        public void UnicodeText_Matches()
        {
            var rows = Rows.Make(new[] { "héllo wörld" }, new[] { "日本語テキスト" });
            var result = RowFilterEngine.Filter(rows, new FilterRequest("日本語", FilterMode.Contains, false), CancellationToken.None);
            Assert.Single(result);
        }

        [Fact]
        public void ParallelPath_ProducesSameResultsInOrder()
        {
            var rows = Rows.Sequence(RowFilterEngine.ParallelThreshold + 100,
                i => new[] { "row" + i, i % 7 == 0 ? "lucky" : "plain" });
            var request = new FilterRequest("lucky", FilterMode.Contains, false);

            var result = RowFilterEngine.Filter(rows, request, CancellationToken.None);

            Assert.Equal(rows.Count(r => r[1] == "lucky"), result.Count);
            // Original order preserved.
            for (int i = 1; i < result.Count; i++)
            {
                Assert.True(result[i].SourceRowNumber > result[i - 1].SourceRowNumber);
            }
        }

        [Fact]
        public void Cancellation_ThrowsOperationCanceled()
        {
            var rows = Rows.Sequence(200_000, i => new[] { "value" + i });
            using var cts = new CancellationTokenSource();
            cts.Cancel();

            Assert.ThrowsAny<OperationCanceledException>(() =>
                RowFilterEngine.Filter(rows, new FilterRequest("needle", FilterMode.Contains, false), cts.Token));
        }
    }
}
