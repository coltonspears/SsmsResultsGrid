using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
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
        internal Control LastCapturedGridControl { get; private set; }

        public DataTable TryCaptureActive()
        {
            return TryCaptureActiveDetailed(out _);
        }

        public DataTable TryCaptureActiveDetailed(out CaptureDiagnostics diagnostics)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            LastFailureReason = null;
            LastCapturedGridControl = null;
            diagnostics = new CaptureDiagnostics();

            try
            {
                var candidates = FindGridCandidatesViaWin32(diagnostics);
                diagnostics.CandidateCount = candidates.Count;
                diagnostics.VisibleCandidateCount = candidates.Count(c => c.IsVisible);

                var grid = SelectBestGridCandidate(candidates, diagnostics);
                if (grid == null)
                {
                    LastFailureReason = BuildFailureReason(diagnostics, "grid-not-found",
                        "No SSMS GridControl found in any open document. Ensure a query has run in Results-to-Grid mode.");
                    return null;
                }

                LastCapturedGridControl = grid;
                var table = ExtractFromSsmsGrid(grid);
                if (table == null)
                {
                    LastFailureReason = BuildFailureReason(diagnostics, "extract-failed",
                        "Found the grid but could not read its data (SSMS internals may have changed).");
                }
                return table;
            }
            catch (Exception ex)
            {
                LastFailureReason = BuildFailureReason(diagnostics, "capture-threw",
                    "Capture threw: " + ex.GetType().Name + " - " + ex.Message);
                return null;
            }
        }

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumChildProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        /// <summary>
        /// Scan every child HWND of the SSMS main window. Any managed Control
        /// whose runtime type matches Microsoft.SqlServer.Management.UI.Grid.GridControl
        /// qualifies. Visible matches win over hidden ones (hidden = in a
        /// backgrounded result tab). This ignores the VS frame model entirely,
        /// which is necessary because when the user clicks Refresh, the active
        /// frame IS our own tool window — not the SQL editor.
        /// </summary>
        private static List<GridCandidate> FindGridCandidatesViaWin32(CaptureDiagnostics diagnostics)
        {
            var roots = new List<IntPtr>();
            var main = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (main != IntPtr.Zero) roots.Add(main);

            var currentPid = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            EnumWindows((hwnd, _) =>
            {
                GetWindowThreadProcessId(hwnd, out var pid);
                if (pid == currentPid && hwnd != main)
                {
                    roots.Add(hwnd);
                }
                return true;
            }, IntPtr.Zero);
            diagnostics.RootWindowCount = roots.Distinct().Count();

            var candidates = new List<GridCandidate>();
            var seenHandles = new HashSet<IntPtr>();
            var seenControls = new HashSet<int>();

            foreach (var root in roots.Distinct())
            {
                EnumChildWindows(root, (hwnd, _) =>
                {
                    if (!seenHandles.Add(hwnd))
                    {
                        return true;
                    }
                    diagnostics.EnumeratedHandleCount++;

                    var ctl = Control.FromHandle(hwnd);
                    if (ctl == null)
                    {
                        var cls = GetWindowClassName(hwnd);
                        if (!string.IsNullOrWhiteSpace(cls) && diagnostics.UnmanagedClassSamples.Count < 8)
                        {
                            diagnostics.UnmanagedClassSamples.Add(cls);
                        }
                        return true;
                    }

                    diagnostics.ManagedHandleCount++;
                    var typeName = ctl.GetType().FullName ?? ctl.GetType().Name;
                    if (diagnostics.ManagedTypeSamples.Count < 10)
                    {
                        diagnostics.ManagedTypeSamples.Add(typeName);
                    }

                    try
                    {
                        CollectCandidatesFromControlTree(ctl, hwnd, candidates, diagnostics, seenControls);
                    }
                    catch (Exception ex)
                    {
                        if (diagnostics.DiscoveryErrors.Count < 5)
                        {
                            diagnostics.DiscoveryErrors.Add($"{typeName}: {ex.GetType().Name} - {ex.Message}");
                        }
                    }
                    return true;
                }, IntPtr.Zero);
            }

            return candidates;
        }

        private static void CollectCandidatesFromControlTree(
            Control root,
            IntPtr hwnd,
            List<GridCandidate> candidates,
            CaptureDiagnostics diagnostics,
            HashSet<int> seenControls)
        {
            if (root == null) return;
            var stack = new Stack<Control>();
            stack.Push(root);

            while (stack.Count > 0)
            {
                var ctl = stack.Pop();
                if (ctl == null) continue;

                var key = RuntimeHelpers.GetHashCode(ctl);
                if (!seenControls.Add(key)) continue;

                var typeName = ctl.GetType().FullName ?? ctl.GetType().Name;
                if (diagnostics.ManagedTypeSamples.Count < 10)
                {
                    diagnostics.ManagedTypeSamples.Add(typeName);
                }

                if (TryCreateCandidate(ctl, hwnd, "control-tree", out var candidate))
                {
                    candidates.Add(candidate);
                    diagnostics.HeuristicHitCount++;
                }

                try
                {
                    foreach (Control child in ctl.Controls)
                    {
                        if (child != null) stack.Push(child);
                    }
                }
                catch (Exception ex)
                {
                    if (diagnostics.DiscoveryErrors.Count < 5)
                    {
                        diagnostics.DiscoveryErrors.Add($"{typeName}.Controls: {ex.GetType().Name} - {ex.Message}");
                    }
                }
            }
        }

        private static bool TryCreateCandidate(Control ctl, IntPtr hwnd, string source, out GridCandidate candidate)
        {
            candidate = null;
            var typeName = ctl.GetType().FullName ?? ctl.GetType().Name;
            var storage = GetProp(ctl, "GridStorage");

            string matchReason = null;
            if (string.Equals(typeName, GridControlTypeName, StringComparison.Ordinal))
            {
                matchReason = "exact-type";
            }
            else if (LooksLikeSsmsGridType(typeName) && storage != null)
            {
                matchReason = "grid-namespace+storage";
            }
            else if (HasGridShape(ctl, storage))
            {
                matchReason = "grid-shape";
            }

            if (matchReason == null) return false;

            candidate = new GridCandidate
            {
                Control = ctl,
                Handle = hwnd,
                IsVisible = IsWindowVisible(hwnd),
                RowCount = ReadRowCount(ctl, storage),
                TypeName = typeName,
                MatchReason = matchReason,
                Source = source
            };
            return true;
        }

        private static bool LooksLikeSsmsGridType(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName)) return false;
            return typeName.IndexOf(".Management.UI.Grid.", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   typeName.EndsWith(".GridControl", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasGridShape(Control ctl, object storage)
        {
            if (storage == null) return false;
            var type = ctl.GetType();
            var hasColumns = type.GetProperty("ColumnsNumber", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance) != null;
            var hasHeader = FindMethod(type, "GetHeaderInfo") != null;
            var hasGetCell = FindMethod(storage.GetType(), "GetCellDataAsString") != null;
            return hasColumns && hasHeader && hasGetCell;
        }

        private static string GetWindowClassName(IntPtr hwnd)
        {
            var buffer = new System.Text.StringBuilder(256);
            var len = GetClassName(hwnd, buffer, buffer.Capacity);
            return len > 0 ? buffer.ToString() : null;
        }

        private static Control SelectBestGridCandidate(IEnumerable<GridCandidate> candidates, CaptureDiagnostics diagnostics)
        {
            GridCandidate best = null;
            var bestScore = int.MinValue;

            foreach (var candidate in candidates)
            {
                var score = 0;
                if (candidate.IsVisible) score += 10_000;
                if (candidate.RowCount > 0) score += 1_000;
                score += (int)Math.Min(candidate.RowCount, 500);

                diagnostics.CandidateSummaries.Add(
                    $"{candidate.Handle}: type={candidate.TypeName}, source={candidate.Source}, match={candidate.MatchReason}, visible={candidate.IsVisible}, rows={candidate.RowCount}, score={score}");

                if (score <= bestScore) continue;
                bestScore = score;
                best = candidate;
            }

            diagnostics.SelectedCandidate = best?.Handle.ToString() ?? "none";
            var probe = "n/a";
            if (best?.Control != null)
            {
                InlineReplacementProbe.TryPrepareInlineReplacement(best.Control, out probe);
            }
            diagnostics.InlineReplacementProbe = probe;
            return best?.Control;
        }

        private static string BuildFailureReason(CaptureDiagnostics diagnostics, string step, string baseMessage)
        {
            diagnostics.FailureStep = step;
            var summary = diagnostics.CandidateSummaries.Count == 0
                ? "none"
                : string.Join("; ", diagnostics.CandidateSummaries.Take(3));
            var managedTypes = diagnostics.ManagedTypeSamples.Count == 0
                ? "none"
                : string.Join(", ", diagnostics.ManagedTypeSamples.Distinct().Take(4));
            var unmanagedClasses = diagnostics.UnmanagedClassSamples.Count == 0
                ? "none"
                : string.Join(", ", diagnostics.UnmanagedClassSamples.Distinct().Take(4));
            var discoveryErrors = diagnostics.DiscoveryErrors.Count == 0
                ? "none"
                : string.Join(" | ", diagnostics.DiscoveryErrors.Take(3));
            return $"{baseMessage} [step={step}; candidates={diagnostics.CandidateCount}; visible={diagnostics.VisibleCandidateCount}; selected={diagnostics.SelectedCandidate}; roots={diagnostics.RootWindowCount}; handles={diagnostics.EnumeratedHandleCount}; managed={diagnostics.ManagedHandleCount}; heuristics={diagnostics.HeuristicHitCount}; inlineProbe={diagnostics.InlineReplacementProbe}; managedTypes={managedTypes}; win32Classes={unmanagedClasses}; discoveryErrors={discoveryErrors}; sample={summary}]";
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

            var getHeader = FindMethod(t, "GetHeaderInfo");

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
            if (target == null || string.IsNullOrWhiteSpace(name)) return null;
            const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic;
            PropertyInfo p = null;

            try
            {
                p = target.GetType().GetProperty(name, flags);
            }
            catch (AmbiguousMatchException)
            {
                // Fall through to explicit property selection.
            }

            if (p == null)
            {
                p = target.GetType()
                    .GetProperties(flags)
                    .Where(prop => string.Equals(prop.Name, name, StringComparison.Ordinal))
                    .Where(prop => prop.GetIndexParameters().Length == 0)
                    .OrderByDescending(prop => prop.DeclaringType == target.GetType())
                    .FirstOrDefault();
            }

            if (p == null || !p.CanRead) return null;
            try { return p.GetValue(target); } catch { return null; }
        }

        private static MethodInfo FindMethod(Type type, string name)
        {
            if (type == null || string.IsNullOrWhiteSpace(name)) return null;
            const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic;
            try
            {
                return type.GetMethod(name, flags);
            }
            catch (AmbiguousMatchException)
            {
                return type.GetMethods(flags)
                    .Where(m => string.Equals(m.Name, name, StringComparison.Ordinal))
                    .OrderBy(m => m.GetParameters().Length)
                    .FirstOrDefault();
            }
        }

        private static long ReadRowCount(Control grid, object storage)
        {
            if (storage == null) return 0;
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

        internal sealed class CaptureDiagnostics
        {
            public int CandidateCount { get; set; }
            public int VisibleCandidateCount { get; set; }
            public string FailureStep { get; set; }
            public string SelectedCandidate { get; set; } = "none";
            public string InlineReplacementProbe { get; set; } = "n/a";
            public int RootWindowCount { get; set; }
            public int EnumeratedHandleCount { get; set; }
            public int ManagedHandleCount { get; set; }
            public int HeuristicHitCount { get; set; }
            public List<string> CandidateSummaries { get; } = new List<string>();
            public List<string> ManagedTypeSamples { get; } = new List<string>();
            public List<string> UnmanagedClassSamples { get; } = new List<string>();
            public List<string> DiscoveryErrors { get; } = new List<string>();
        }

        private sealed class GridCandidate
        {
            public Control Control { get; set; }
            public IntPtr Handle { get; set; }
            public bool IsVisible { get; set; }
            public long RowCount { get; set; }
            public string TypeName { get; set; }
            public string MatchReason { get; set; }
            public string Source { get; set; }
        }
    }
}
