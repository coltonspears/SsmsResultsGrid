using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;

namespace SsmsResultsGrid.Services.Capture
{
    /// <summary>
    /// Encapsulates all reflection against the SSMS brokered-service surface so the
    /// rest of the codebase deals with plain CLR types only. One MethodInfo.Invoke
    /// per page of rows — never per cell.
    ///
    /// The proxy reads from the ACTIVE SQL editor (SSMS 22.5) or the editor named by
    /// the moniker captured at creation time (22.6+). Create per capture, dispose after.
    /// </summary>
    internal sealed class GridBrokeredClient : IDisposable
    {
        private readonly ContractTypes _types;
        private readonly object _proxy;
        private readonly string _editorMoniker;

        private GridBrokeredClient(ContractTypes types, object proxy, string editorMoniker)
        {
            _types = types;
            _proxy = proxy;
            _editorMoniker = editorMoniker;
        }

        public static async Task<GridBrokeredClient> CreateAsync(AsyncPackage package, CancellationToken ct)
        {
            if (package == null) throw new ArgumentNullException(nameof(package));

            var types = await ContractTypes.GetAsync(ct).ConfigureAwait(true);

            await package.JoinableTaskFactory.SwitchToMainThreadAsync(ct);

            var sbcType = ResolveType("Microsoft.VisualStudio.Shell.ServiceBroker.SVsBrokeredServiceContainer");
            if (sbcType == null)
                throw new InvalidOperationException("SVsBrokeredServiceContainer type not loaded.");

            var container = await package.GetServiceAsync(sbcType).ConfigureAwait(true);
            if (container == null)
                throw new InvalidOperationException("SVsBrokeredServiceContainer service unavailable.");

            await package.JoinableTaskFactory.SwitchToMainThreadAsync(ct);

            var getBroker = container.GetType().GetMethod("GetFullAccessServiceBroker");
            if (getBroker == null)
                throw new InvalidOperationException("GetFullAccessServiceBroker not found on container.");

            var broker = getBroker.Invoke(container, null);
            if (broker == null)
                throw new InvalidOperationException("Service broker is null.");

            // The moniker hop is only needed on SSMS builds whose tab-data methods take
            // a leading editorMoniker parameter (22.6+). On 22.5 the service is implicitly
            // bound to the active editor.
            string moniker = null;
            if (types.GridSegmentRequiresEditorMoniker || types.AvailablePanesRequiresEditorMoniker)
            {
                var editorServiceProxy = await GetProxyAsync(
                    broker, types.ISqlEditorServiceBrokered, types.SqlEditorServiceMoniker, ct).ConfigureAwait(true);
                if (editorServiceProxy == null)
                    throw new NoActiveEditorException(
                        "IServiceBroker.GetProxyAsync<ISqlEditorServiceBrokered>() returned null.");
                try
                {
                    moniker = await GetCurrentEditorMonikerAsync(editorServiceProxy, types, ct).ConfigureAwait(true);
                }
                finally
                {
                    (editorServiceProxy as IDisposable)?.Dispose();
                }

                if (string.IsNullOrEmpty(moniker))
                    throw new NoActiveEditorException("No active SQL editor window (EditorMoniker is empty).");
            }

            var tabDataProxy = await GetProxyAsync(
                broker, types.IQueryEditorTabDataServiceBrokered, types.QueryEditorTabDataServiceMoniker, ct).ConfigureAwait(true);
            if (tabDataProxy == null)
                throw new InvalidOperationException(
                    "IServiceBroker.GetProxyAsync<IQueryEditorTabDataServiceBrokered>() returned null. " +
                    "Ensure a SQL query window is the active document.");

            return new GridBrokeredClient(types, tabDataProxy, moniker);
        }

        /// <summary>True when the active editor currently exposes a GridResults pane.</summary>
        public async Task<bool> IsGridResultsPaneAvailableAsync(CancellationToken ct)
        {
            var positional = _types.AvailablePanesRequiresEditorMoniker
                ? new object[] { _editorMoniker }
                : Array.Empty<object>();
            var args = BuildArgs(_types.GetAvailablePanesAsyncMethod, positional, ct);
            var task = _types.GetAvailablePanesAsyncMethod.Invoke(_proxy, args);
            var panes = await UnwrapAsync(task).ConfigureAwait(true);
            if (!(panes is IEnumerable enumerable)) return false;

            foreach (var info in enumerable)
            {
                if (info == null) continue;
                var paneType = _types.QueryResultsPaneInfo_PaneType.GetValue(info);
                if (Equals(paneType, _types.QueryResultsPane_GridResults))
                    return true;
            }
            return false;
        }

        public async Task<GridSegment> GetGridSegmentAsync(
            int gridIndex, int startColumn, int columnCount, long startRow, int maxRows, int maxCellChars,
            CancellationToken ct)
        {
            var positional = _types.GridSegmentRequiresEditorMoniker
                ? new object[] { _editorMoniker, gridIndex, startColumn, columnCount, startRow, maxRows, maxCellChars }
                : new object[] { gridIndex, startColumn, columnCount, startRow, maxRows, maxCellChars };
            var args = BuildArgs(_types.GetGridResultsSegmentAsyncMethod, positional, ct);
            var task = _types.GetGridResultsSegmentAsyncMethod.Invoke(_proxy, args);
            var segment = await UnwrapAsync(task).ConfigureAwait(true);
            if (segment == null) return null;

            return new GridSegment(
                gridIndex: Convert.ToInt32(_types.GridResultsSegment_GridIndex.GetValue(segment)),
                totalGridCount: Convert.ToInt32(_types.GridResultsSegment_TotalGridCount.GetValue(segment)),
                totalRowCount: Convert.ToInt64(_types.GridResultsSegment_TotalRowCount.GetValue(segment)),
                totalColumnCount: Convert.ToInt32(_types.GridResultsSegment_TotalColumnCount.GetValue(segment)),
                startRow: Convert.ToInt64(_types.GridResultsSegment_StartRow.GetValue(segment)),
                columnNames: ToStringList(_types.GridResultsSegment_ColumnNames.GetValue(segment)),
                rows: ToRowList(_types.GridResultsSegment_Rows.GetValue(segment)));
        }

        public void Dispose()
        {
            try { (_proxy as IDisposable)?.Dispose(); }
            catch { /* best-effort */ }
        }

        private static List<string> ToStringList(object value)
        {
            var result = new List<string>();
            if (value is IEnumerable enumerable)
            {
                foreach (var item in enumerable)
                {
                    result.Add(item?.ToString() ?? string.Empty);
                }
            }
            return result;
        }

        private static List<string[]> ToRowList(object value)
        {
            var result = new List<string[]>();
            if (!(value is IEnumerable rows)) return result;

            foreach (var row in rows)
            {
                switch (row)
                {
                    case null:
                        result.Add(Array.Empty<string>());
                        break;
                    case string[] direct:
                        result.Add(direct);
                        break;
                    case IEnumerable cells:
                        var list = new List<string>();
                        foreach (var cell in cells)
                        {
                            list.Add(cell as string ?? cell?.ToString() ?? string.Empty);
                        }
                        result.Add(list.ToArray());
                        break;
                    default:
                        result.Add(new[] { row.ToString() });
                        break;
                }
            }
            return result;
        }

        private static async Task<string> GetCurrentEditorMonikerAsync(
            object editorServiceProxy, ContractTypes types, CancellationToken ct)
        {
            var args = BuildArgs(types.GetCurrentConnectionAsyncMethod, Array.Empty<object>(), ct);
            object task;
            try
            {
                task = types.GetCurrentConnectionAsyncMethod.Invoke(editorServiceProxy, args);
            }
            catch (TargetInvocationException tie) when (tie.InnerException != null)
            {
                throw new NoActiveEditorException(
                    "GetCurrentConnectionAsync threw " + tie.InnerException.GetType().Name + ": " +
                    tie.InnerException.Message, tie.InnerException);
            }

            object details;
            try
            {
                details = await UnwrapAsync(task).ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                throw new NoActiveEditorException(
                    "GetCurrentConnectionAsync RPC failed: " + ex.GetType().Name + ": " + ex.Message, ex);
            }

            if (details == null) return null;
            return (string)types.SqlEditorConnectionDetails_EditorMoniker.GetValue(details);
        }

        // Build an args array sized to the resolved brokered method's parameter list: copy our
        // known positional args into the leading slots, slot the CancellationToken by type, and
        // fill any trailing parameters with defaults.
        private static object[] BuildArgs(MethodInfo method, object[] positional, CancellationToken ct)
        {
            var ps = method.GetParameters();
            var args = new object[ps.Length];
            int n = Math.Min(positional.Length, ps.Length);
            for (int i = 0; i < n; i++) args[i] = positional[i];
            for (int i = n; i < ps.Length; i++)
            {
                if (ps[i].ParameterType == typeof(CancellationToken)) args[i] = ct;
                else if (ps[i].HasDefaultValue) args[i] = ps[i].DefaultValue;
                else args[i] = ps[i].ParameterType.IsValueType
                    ? Activator.CreateInstance(ps[i].ParameterType)
                    : null;
            }
            return args;
        }

        // Brokered proxy methods return ValueTask<T> (or Task<T>); unwrap either via reflection.
        private static async Task<object> UnwrapAsync(object taskOrValueTask)
        {
            if (taskOrValueTask == null) return null;
            if (taskOrValueTask is Task t)
            {
                await t.ConfigureAwait(true);
                return t.GetType().GetProperty("Result")?.GetValue(t);
            }
            var asTask = taskOrValueTask.GetType().GetMethod("AsTask", Type.EmptyTypes);
            if (asTask == null)
                throw new InvalidOperationException(
                    "Cannot await result of type " + taskOrValueTask.GetType().FullName);
            var task = (Task)asTask.Invoke(taskOrValueTask, null);
            await task.ConfigureAwait(true);
            return task.GetType().GetProperty("Result")?.GetValue(task);
        }

        private static async Task<object> GetProxyAsync(
            object broker, Type contractType, object monikerObj, CancellationToken ct)
        {
            // The static moniker may be a ServiceMoniker or a ServiceRpcDescriptor wrapping one.
            if (monikerObj.GetType().Name.IndexOf("ServiceRpcDescriptor", StringComparison.Ordinal) >= 0)
            {
                var monikerProp = monikerObj.GetType().GetProperty("Moniker");
                var inner = monikerProp?.GetValue(monikerObj);
                if (inner != null) monikerObj = inner;
            }

            var generics = broker.GetType().GetMethods()
                .Where(m => m.Name == "GetProxyAsync" && m.IsGenericMethodDefinition)
                .ToList();

            MethodInfo chosen = generics.FirstOrDefault(m =>
            {
                var ps = m.GetParameters();
                return ps.Length >= 1 && ps[0].ParameterType.IsAssignableFrom(monikerObj.GetType());
            }) ?? generics.FirstOrDefault();

            if (chosen == null)
                throw new InvalidOperationException("IServiceBroker.GetProxyAsync<T> not found.");

            var generic = chosen.MakeGenericMethod(contractType);
            var ps2 = generic.GetParameters();
            var args = new object[ps2.Length];
            args[0] = monikerObj;
            for (int i = 1; i < ps2.Length; i++)
            {
                if (ps2[i].ParameterType == typeof(CancellationToken))
                    args[i] = ct;
                else if (ps2[i].HasDefaultValue)
                    args[i] = ps2[i].DefaultValue;
                else
                    args[i] = ps2[i].ParameterType.IsValueType
                        ? Activator.CreateInstance(ps2[i].ParameterType)
                        : null;
            }

            var result = generic.Invoke(broker, args);
            return await UnwrapAsync(result).ConfigureAwait(true);
        }

        private static Type ResolveType(string fullName)
        {
            var t = Type.GetType(fullName);
            if (t != null) return t;
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                t = asm.GetType(fullName);
                if (t != null) return t;
            }
            return null;
        }
    }

    /// <summary>Raised when no SQL editor is active or the moniker lookup fails.</summary>
    internal sealed class NoActiveEditorException : Exception
    {
        public NoActiveEditorException(string message) : base(message) { }
        public NoActiveEditorException(string message, Exception inner) : base(message, inner) { }
    }

    /// <summary>Plain-CLR copy of one brokered GridResultsSegment.</summary>
    internal sealed class GridSegment
    {
        public GridSegment(
            int gridIndex, int totalGridCount, long totalRowCount, int totalColumnCount,
            long startRow, List<string> columnNames, List<string[]> rows)
        {
            GridIndex = gridIndex;
            TotalGridCount = totalGridCount;
            TotalRowCount = totalRowCount;
            TotalColumnCount = totalColumnCount;
            StartRow = startRow;
            ColumnNames = columnNames ?? new List<string>();
            Rows = rows ?? new List<string[]>();
        }

        public int GridIndex { get; }
        public int TotalGridCount { get; }
        public long TotalRowCount { get; }
        public int TotalColumnCount { get; }
        public long StartRow { get; }
        public List<string> ColumnNames { get; }
        public List<string[]> Rows { get; }
    }
}
