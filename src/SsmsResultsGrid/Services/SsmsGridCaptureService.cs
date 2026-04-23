using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SsmsResultsGrid.Services
{
    /// <summary>
    /// Reflection-based capture of SSMS's native results grid.
    ///
    /// SSMS hosts a WinForms control from Microsoft.SqlServer.GridControl.dll,
    /// type Microsoft.SqlServer.Management.UI.Grid.GridControl, inside the SQL
    /// editor's results pane. No public API exposes it, so we reach in via
    /// reflection against:
    ///   GridControl.GridStorage    (IGridStorage)
    ///   GridControl.NumRowsInt     (Int64)
    ///   GridControl.ColumnsNumber  (Int32)
    ///   GridControl.GetHeaderInfo(int col, out string text, out Bitmap, out GridCheckBoxState)
    ///   IGridStorage.GetCellDataAsString(Int64 row, Int32 col) -> string
    ///
    /// The capture walks ALL open document frames, not just the active one,
    /// because when the user clicks Refresh on our own tool window, that tool
    /// window is the active frame. Any failure returns null and a reason.
    /// </summary>
    internal sealed class SsmsGridCaptureService
    {
        private const string GridControlTypeName = "Microsoft.SqlServer.Management.UI.Grid.GridControl";
        private const int MaxRowsToCapture = 100_000;

        public SsmsGridCaptureService(AsyncPackage _ = null) { }

        public string LastFailureReason { get; private set; }

        public DataTable TryCaptureActive()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            LastFailureReason = null;

            try
            {
                var grid = FindGridViaWin32();
                if (grid == null)
                {
                    LastFailureReason = "No SSMS GridControl found in any open document. Ensure a query has run in Results-to-Grid mode.";
                    return null;
                }
                var table = ExtractFromSsmsGrid(grid);
                if (table == null)
                {
                    LastFailureReason = "Found the grid but could not read its data (SSMS internals may have changed).";
                }
                return table;
            }
            catch (Exception ex)
            {
                LastFailureReason = "Capture threw: " + ex.GetType().Name + " — " + ex.Message;
                return null;
            }
        }

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        /// <summary>
        /// Scan every child HWND of the SSMS main window. Any managed Control
        /// whose runtime type matches Microsoft.SqlServer.Management.UI.Grid.GridControl
        /// qualifies. Visible matches win over hidden ones (hidden = in a
        /// backgrounded result tab). This ignores the VS frame model entirely,
        /// which is necessary because when the user clicks Refresh, the active
        /// frame IS our own tool window — not the SQL editor.
        /// </summary>
        private static Control FindGridViaWin32()
        {
            var main = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (main == IntPtr.Zero) return null;

            Control visibleHit = null;
            Control anyHit = null;
            EnumChildWindows(main, (hwnd, _) =>
            {
                var ctl = Control.FromHandle(hwnd);
                if (ctl != null && ctl.GetType().FullName == GridControlTypeName)
                {
                    if (anyHit == null) anyHit = ctl;
                    if (IsWindowVisible(hwnd))
                    {
                        visibleHit = ctl;
                        return false; // stop enumeration on first visible hit
                    }
                }
                return true;
            }, IntPtr.Zero);
            return visibleHit ?? anyHit;
        }

        private static DataTable ExtractFromSsmsGrid(Control grid)
        {
            var t = grid.GetType();

            var storage = GetProp(grid, "GridStorage");
            if (storage == null) return null;

            int colCount = Convert.ToInt32(GetProp(grid, "ColumnsNumber") ?? 0);
            if (colCount <= 0) return null;

            long rowCount = ReadRowCount(grid, storage);

            var getCell = storage.GetType().GetMethod(
                "GetCellDataAsString",
                BindingFlags.Public | BindingFlags.Instance,
                binder: null,
                types: new[] { typeof(long), typeof(int) },
                modifiers: null);
            if (getCell == null) return null;

            var getHeader = t.GetMethod(
                "GetHeaderInfo",
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);

            var table = new DataTable();
            for (int c = 0; c < colCount; c++)
            {
                string header = ReadHeader(grid, getHeader, c) ?? $"Col {c + 1}";
                var unique = header;
                int dedup = 2;
                while (table.Columns.Contains(unique))
                    unique = header + " (" + dedup++ + ")";
                table.Columns.Add(unique, typeof(string));
            }

            long cap = Math.Min(rowCount, MaxRowsToCapture);
            for (long r = 0; r < cap; r++)
            {
                var row = table.NewRow();
                for (int c = 0; c < colCount; c++)
                {
                    try
                    {
                        row[c] = (string)getCell.Invoke(storage, new object[] { r, c }) ?? string.Empty;
                    }
                    catch
                    {
                        row[c] = string.Empty;
                    }
                }
                table.Rows.Add(row);
            }
            return table;
        }

        private static object GetProp(object target, string name)
        {
            var p = target.GetType().GetProperty(name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
            if (p == null) return null;
            try { return p.GetValue(target); } catch { return null; }
        }

        private static long ReadRowCount(Control grid, object storage)
        {
            // Preferred: GridControl.NumRowsInt (Int64).
            var p = grid.GetType().GetProperty("NumRowsInt",
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
            if (p != null)
            {
                try { return Convert.ToInt64(p.GetValue(grid)); } catch { }
            }
            // Fallback: IGridStorage.NumRows (method, Int64).
            var m = storage.GetType().GetMethod("NumRows",
                BindingFlags.Public | BindingFlags.Instance,
                null, Type.EmptyTypes, null);
            if (m != null)
            {
                try { return Convert.ToInt64(m.Invoke(storage, null)); } catch { }
            }
            return 0;
        }

        private static string ReadHeader(Control grid, MethodInfo getHeader, int columnIndex)
        {
            if (getHeader == null) return null;
            try
            {
                var parameters = getHeader.GetParameters();
                var args = new object[parameters.Length];
                args[0] = columnIndex;
                for (int i = 1; i < args.Length; i++) args[i] = null;
                getHeader.Invoke(grid, args);
                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i] is string s && !string.IsNullOrEmpty(s)) return s;
                }
            }
            catch { }
            return null;
        }
    }
}
