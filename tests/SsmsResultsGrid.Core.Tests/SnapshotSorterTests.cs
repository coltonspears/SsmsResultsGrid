using System.Linq;
using System.Threading;
using SsmsResultsGrid.Core.Sorting;
using Xunit;

namespace SsmsResultsGrid.Core.Tests
{
    public class SnapshotSorterTests
    {
        [Fact]
        public void NumericColumn_SortsNumerically_NotLexically()
        {
            var rows = Rows.Make(new[] { "10" }, new[] { "9" }, new[] { "100" });
            var sorted = SnapshotSorter.Sort(rows, 0, descending: false, CancellationToken.None);
            Assert.Equal(new[] { "9", "10", "100" }, sorted.Select(r => r[0]));
        }

        [Fact]
        public void MixedColumn_FallsBackToOrdinalString()
        {
            var rows = Rows.Make(new[] { "b" }, new[] { "10" }, new[] { "a" });
            var sorted = SnapshotSorter.Sort(rows, 0, descending: false, CancellationToken.None);
            Assert.Equal(new[] { "10", "a", "b" }, sorted.Select(r => r[0]));
        }

        [Fact]
        public void Descending_ReversesOrder()
        {
            var rows = Rows.Make(new[] { "1" }, new[] { "3" }, new[] { "2" });
            var sorted = SnapshotSorter.Sort(rows, 0, descending: true, CancellationToken.None);
            Assert.Equal(new[] { "3", "2", "1" }, sorted.Select(r => r[0]));
        }

        [Fact]
        public void EmptyAndNullCells_GroupTogetherInNumericSort()
        {
            var rows = Rows.Make(new[] { "5" }, new[] { "" }, new[] { "NULL" }, new[] { "1" });
            var sorted = SnapshotSorter.Sort(rows, 0, descending: false, CancellationToken.None);
            Assert.Equal(new[] { "", "NULL", "1", "5" }, sorted.Select(r => r[0]));
        }

        [Fact]
        public void DecimalValues_SortByMagnitude()
        {
            var rows = Rows.Make(new[] { "2.5" }, new[] { "-1.25" }, new[] { "0.75" });
            var sorted = SnapshotSorter.Sort(rows, 0, descending: false, CancellationToken.None);
            Assert.Equal(new[] { "-1.25", "0.75", "2.5" }, sorted.Select(r => r[0]));
        }

        [Fact]
        public void SingleRow_ReturnsSameInstance()
        {
            var rows = Rows.Make(new[] { "only" });
            var sorted = SnapshotSorter.Sort(rows, 0, descending: false, CancellationToken.None);
            Assert.Same(rows, sorted);
        }
    }
}
