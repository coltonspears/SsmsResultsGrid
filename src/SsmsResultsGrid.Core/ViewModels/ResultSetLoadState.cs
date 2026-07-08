namespace SsmsResultsGrid.Core.ViewModels
{
    public enum ResultSetLoadState
    {
        /// <summary>Metadata known (columns, row count) but rows not fetched yet — loads on first selection.</summary>
        Pending,
        Loading,
        Loaded
    }
}
