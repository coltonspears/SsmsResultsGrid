using System;
using System.Collections.Generic;
using SsmsResultsGrid.Core.Models;
using SsmsResultsGrid.Core.Mvvm;

namespace SsmsResultsGrid.Core.Tests
{
    /// <summary>Runs posted actions immediately on the calling thread.</summary>
    internal sealed class ImmediateDispatcher : IUiDispatcher
    {
        public bool CheckAccess() => true;
        public void Post(Action action) => action();
    }

    internal static class Rows
    {
        public static ResultRow[] Make(params string[][] cellRows)
        {
            var result = new ResultRow[cellRows.Length];
            for (int i = 0; i < cellRows.Length; i++)
            {
                result[i] = new ResultRow(cellRows[i], i + 1);
            }
            return result;
        }

        public static ResultRow[] Sequence(int count, Func<int, string[]> factory)
        {
            var result = new ResultRow[count];
            for (int i = 0; i < count; i++)
            {
                result[i] = new ResultRow(factory(i), i + 1);
            }
            return result;
        }

        public static List<string[]> CellBatch(int start, int count, Func<int, string[]> factory)
        {
            var list = new List<string[]>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(factory(start + i));
            }
            return list;
        }
    }
}
