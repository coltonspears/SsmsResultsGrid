using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SsmsResultsGrid.Core.ViewModels;
using Xunit;

namespace SsmsResultsGrid.Core.Tests
{
    public class ResultSetViewModelTests
    {
        private static readonly TimeSpan ShortDebounce = TimeSpan.FromMilliseconds(50);

        private static ResultSetViewModel CreateLoadedVm(int rowCount = 10)
        {
            var vm = new ResultSetViewModel(0, new ImmediateDispatcher(), null, ShortDebounce);
            vm.BeginLoad(new[] { "id", "name" }, rowCount);
            vm.AppendRows(Rows.CellBatch(0, rowCount, i => new[] { i.ToString(), "name" + i }), startRow: 0);
            vm.CompleteLoad(truncated: false);
            return vm;
        }

        private static async Task WaitForAsync(Func<bool> condition, int timeoutMs = 5000)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            while (!condition())
            {
                Assert.True(DateTime.UtcNow < deadline, "Timed out waiting for condition.");
                await Task.Delay(10);
            }
        }

        [Fact]
        public void Load_PublishesAllRowsAndColumns()
        {
            var vm = CreateLoadedVm(5);
            Assert.Equal(5, vm.VisibleRows.Count);
            Assert.Equal(new[] { "id", "name" }, vm.ColumnNames);
            Assert.Equal(ResultSetLoadState.Loaded, vm.LoadState);
        }

        [Fact]
        public void ProgressiveAppend_GrowsVisibleRows_WithStableRowNumbers()
        {
            var vm = new ResultSetViewModel(0, new ImmediateDispatcher(), null, ShortDebounce);
            vm.BeginLoad(new[] { "c" }, 6);
            vm.AppendRows(Rows.CellBatch(0, 3, i => new[] { "v" + i }), startRow: 0);
            Assert.Equal(3, vm.VisibleRows.Count);

            vm.AppendRows(Rows.CellBatch(3, 3, i => new[] { "v" + i }), startRow: 3);
            Assert.Equal(6, vm.VisibleRows.Count);
            Assert.Equal(6, vm.VisibleRows[5].SourceRowNumber);
        }

        [Fact]
        public async Task FilterText_IsDebounced_AndApplied()
        {
            var vm = CreateLoadedVm(20);
            vm.FilterText = "name1";

            // Immediately after typing, the applied filter has not caught up yet.
            Assert.Equal(string.Empty, vm.AppliedFilterText);

            await WaitForAsync(() => vm.AppliedFilterText == "name1");
            // name1, name10..name19
            Assert.Equal(11, vm.VisibleRows.Count);
        }

        [Fact]
        public async Task RapidFilterChanges_OnlyLastOneWins()
        {
            var vm = CreateLoadedVm(50);
            vm.FilterText = "name1";
            vm.FilterText = "name2";
            vm.FilterText = "name33";

            await WaitForAsync(() => vm.AppliedFilterText == "name33");
            Assert.Single(vm.VisibleRows);
            Assert.Equal("name33", vm.VisibleRows[0][1]);
        }

        [Fact]
        public async Task ClearingFilter_RestoresAllRowsSynchronously()
        {
            var vm = CreateLoadedVm(10);
            vm.FilterText = "name3";
            await WaitForAsync(() => vm.AppliedFilterText == "name3");

            vm.FilterText = string.Empty;
            // Empty filter takes the synchronous fast path.
            Assert.Equal(10, vm.VisibleRows.Count);
            Assert.Equal(string.Empty, vm.AppliedFilterText);
        }

        [Fact]
        public async Task Sort_AppliesToFilteredRows()
        {
            var vm = CreateLoadedVm(10);
            vm.SetSort(0, descending: true);

            await WaitForAsync(() => vm.VisibleRows.Count == 10 && vm.VisibleRows[0][0] == "9");
            Assert.Equal("0", vm.VisibleRows[9][0]);

            vm.ClearSort();
            Assert.Equal("0", vm.VisibleRows[0][0]);
        }

        [Fact]
        public async Task CaseSensitiveToggle_Refilters()
        {
            var vm = new ResultSetViewModel(0, new ImmediateDispatcher(), null, ShortDebounce);
            vm.BeginLoad(new[] { "c" }, 2);
            vm.AppendRows(Rows.CellBatch(0, 2, i => new[] { i == 0 ? "Value" : "value" }), startRow: 0);
            vm.CompleteLoad(false);

            vm.FilterText = "Value";
            await WaitForAsync(() => vm.AppliedFilterText == "Value");
            Assert.Equal(2, vm.VisibleRows.Count);

            vm.FilterCaseSensitive = true;
            await WaitForAsync(() => vm.AppliedFilterCaseSensitive);
            Assert.Single(vm.VisibleRows);
        }

        [Fact]
        public void Truncation_ReflectedInSummary()
        {
            var vm = new ResultSetViewModel(0, new ImmediateDispatcher(), null, ShortDebounce);
            vm.BeginLoad(new[] { "c" }, 1000);
            vm.AppendRows(Rows.CellBatch(0, 100, i => new[] { "v" + i }), startRow: 0);
            vm.CompleteLoad(truncated: true);

            Assert.True(vm.IsTruncated);
            Assert.Contains("100", vm.RowSummaryText);
            Assert.Contains("1,000", vm.RowSummaryText);
        }

        [Fact]
        public async Task ReLoad_ResetsRowsAndKeepsFilterText()
        {
            var vm = CreateLoadedVm(10);
            vm.FilterText = "name2";
            await WaitForAsync(() => vm.AppliedFilterText == "name2");

            // Simulate an F5 re-execution: same schema, new rows.
            vm.BeginLoad(new[] { "id", "name" }, 5);
            vm.AppendRows(Rows.CellBatch(0, 5, i => new[] { i.ToString(), "name" + i }), startRow: 0);
            vm.CompleteLoad(false);

            Assert.Equal("name2", vm.FilterText);
            await WaitForAsync(() => vm.AppliedFilterText == "name2" && vm.VisibleRows.Count == 1);
        }
    }
}
