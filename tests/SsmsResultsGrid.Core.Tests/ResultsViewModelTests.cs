using System;
using System.Collections.Generic;
using SsmsResultsGrid.Core.ViewModels;
using Xunit;

namespace SsmsResultsGrid.Core.Tests
{
    public class ResultsViewModelTests
    {
        private static ResultsViewModel Create() =>
            new ResultsViewModel(new ImmediateDispatcher(), null, TimeSpan.FromMilliseconds(10));

        [Fact]
        public void BeginCapture_CreatesResultSets_AndSelectsFirst()
        {
            var vm = Create();
            vm.BeginCapture(3);

            Assert.Equal(3, vm.ResultSets.Count);
            Assert.Same(vm.ResultSets[0], vm.SelectedResultSet);
            Assert.True(vm.HasMultipleResultSets);
            Assert.True(vm.IsCapturing);
        }

        [Fact]
        public void BeginCapture_ShrinksExcessResultSets_AndKeepsValidSelection()
        {
            var vm = Create();
            vm.BeginCapture(3);
            vm.SelectedResultSet = vm.ResultSets[1];

            vm.BeginCapture(2);

            Assert.Equal(2, vm.ResultSets.Count);
            Assert.Same(vm.ResultSets[1], vm.SelectedResultSet);
        }

        [Fact]
        public void BeginCapture_SelectionOutOfRange_FallsBackToFirst()
        {
            var vm = Create();
            vm.BeginCapture(3);
            vm.SelectedResultSet = vm.ResultSets[2];

            vm.BeginCapture(1);

            Assert.Single(vm.ResultSets);
            Assert.Same(vm.ResultSets[0], vm.SelectedResultSet);
        }

        [Fact]
        public void SelectingPendingResultSet_RaisesLoadRequested()
        {
            var vm = Create();
            vm.BeginCapture(2);
            var requested = new List<ResultSetViewModel>();
            vm.ResultSetLoadRequested += (_, set) => requested.Add(set);

            vm.ResultSets[1].SetPendingMetadata(new[] { "c" }, 42);
            vm.SelectedResultSet = vm.ResultSets[1];

            Assert.Single(requested);
            Assert.Same(vm.ResultSets[1], requested[0]);
        }

        [Fact]
        public void SelectingLoadedResultSet_DoesNotRaiseLoadRequested()
        {
            var vm = Create();
            vm.BeginCapture(2);
            var requests = 0;
            vm.ResultSetLoadRequested += (_, __) => requests++;

            vm.ResultSets[1].BeginLoad(new[] { "c" }, 0);
            vm.ResultSets[1].CompleteLoad(false);
            vm.SelectedResultSet = vm.ResultSets[1];

            Assert.Equal(0, requests);
        }

        [Fact]
        public void SetNoGridResults_ClearsEverything()
        {
            var vm = Create();
            vm.BeginCapture(2);

            vm.SetNoGridResults();

            Assert.Empty(vm.ResultSets);
            Assert.Null(vm.SelectedResultSet);
            Assert.False(vm.IsCapturing);
            Assert.False(vm.HasMultipleResultSets);
        }

        [Fact]
        public void CompleteCapture_UpdatesStatusAndTimestamp()
        {
            var vm = Create();
            vm.BeginCapture(1);
            vm.ResultSets[0].BeginLoad(new[] { "c" }, 2);
            vm.ResultSets[0].AppendRows(Rows.CellBatch(0, 2, i => new[] { "v" + i }), 0);
            vm.ResultSets[0].CompleteLoad(false);

            vm.CompleteCapture(new DateTime(2026, 7, 8, 12, 30, 0));

            Assert.False(vm.IsCapturing);
            Assert.NotEqual(Core.Resources.Strings.UpdatedNever, vm.LastUpdatedText);
            Assert.Contains("2", vm.StatusText);
        }

        [Fact]
        public void RefreshCommand_DisabledUntilHandlerInstalled_AndWhileCapturing()
        {
            var vm = Create();
            Assert.False(vm.RefreshCommand.CanExecute(null));

            vm.SetRefreshHandler(() => System.Threading.Tasks.Task.CompletedTask);
            Assert.True(vm.RefreshCommand.CanExecute(null));

            vm.BeginCapture(1);
            Assert.False(vm.RefreshCommand.CanExecute(null));

            vm.CompleteCapture(DateTime.Now);
            Assert.True(vm.RefreshCommand.CanExecute(null));
        }
    }
}
