using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Threading;

namespace SsmsResultsGrid.Services.Capture
{
    /// <summary>
    /// Late-bound reflection cache for the SSMS-shipped BrokeredContracts assembly.
    /// The DLL ships with SSMS and cannot be redistributed, so it is loaded from the
    /// SSMS install directory at runtime and every type/method is looked up by name
    /// with a descriptive exception on drift.
    ///
    /// Signature drift note: SSMS 22.5 exposes the tab-data methods WITHOUT a leading
    /// editor-moniker parameter; SSMS 22.6+ added a leading `string editorMoniker`.
    /// Each method is resolved against both shapes and the moniker requirement is
    /// recorded so callers only perform the extra moniker hop when needed.
    /// </summary>
    internal sealed class ContractTypes
    {
        public const string ContractsDllName =
            "Microsoft.SqlServer.Management.UI.VSIntegration.SqlEditor.BrokeredContracts.dll";

        private const string Ns = "Microsoft.SqlServer.Management.UI.VSIntegration.BrokeredServices";

        private static readonly AsyncLazy<ContractTypes> _instance =
            new AsyncLazy<ContractTypes>(LoadAsync, joinableTaskFactory: null);

        public static Task<ContractTypes> GetAsync(CancellationToken ct) => _instance.GetValueAsync(ct);

        public Assembly ContractsAssembly { get; private set; }
        public Type IQueryEditorTabDataServiceBrokered { get; private set; }
        public Type ISqlEditorServiceBrokered { get; private set; }
        public object QueryEditorTabDataServiceMoniker { get; private set; }
        public object SqlEditorServiceMoniker { get; private set; }

        public MethodInfo GetGridResultsSegmentAsyncMethod { get; private set; }
        public bool GridSegmentRequiresEditorMoniker { get; private set; }
        public MethodInfo GetAvailablePanesAsyncMethod { get; private set; }
        public bool AvailablePanesRequiresEditorMoniker { get; private set; }
        public MethodInfo GetCurrentConnectionAsyncMethod { get; private set; }

        public PropertyInfo GridResultsSegment_GridIndex { get; private set; }
        public PropertyInfo GridResultsSegment_TotalGridCount { get; private set; }
        public PropertyInfo GridResultsSegment_TotalRowCount { get; private set; }
        public PropertyInfo GridResultsSegment_TotalColumnCount { get; private set; }
        public PropertyInfo GridResultsSegment_StartRow { get; private set; }
        public PropertyInfo GridResultsSegment_ColumnNames { get; private set; }
        public PropertyInfo GridResultsSegment_Rows { get; private set; }
        public PropertyInfo QueryResultsPaneInfo_PaneType { get; private set; }
        public PropertyInfo SqlEditorConnectionDetails_EditorMoniker { get; private set; }
        public object QueryResultsPane_GridResults { get; private set; }

        private static Task<ContractTypes> LoadAsync()
        {
            // Purely CPU-bound (Assembly.LoadFrom + reflection); no main-thread or RPC work.
            var ct = new ContractTypes();
            ct.Initialize();
            return Task.FromResult(ct);
        }

        private void Initialize()
        {
            string ideDir = AppDomain.CurrentDomain.BaseDirectory;
            string dllPath = Path.Combine(ideDir, ContractsDllName);
            if (!File.Exists(dllPath))
                throw new FileNotFoundException(
                    "SSMS BrokeredContracts assembly not found at " + dllPath +
                    " — this extension requires SQL Server Management Studio 22 or later.", dllPath);

            ContractsAssembly = Assembly.LoadFrom(dllPath);

            IQueryEditorTabDataServiceBrokered = RequireType("IQueryEditorTabDataServiceBrokered");
            ISqlEditorServiceBrokered = RequireType("ISqlEditorServiceBrokered");
            var gridSegmentType = RequireType("GridResultsSegment");
            var paneEnumType = RequireType("QueryResultsPane");
            var paneInfoType = RequireType("QueryResultsPaneInfo");
            var connectionDetailsType = RequireType("SqlEditorConnectionDetails");
            var tabDataDescriptorsType = RequireType("QueryEditorTabDataServiceDescriptors");
            var editorServiceDescriptorsType = RequireType("SqlEditorBrokeredServiceDescriptors");

            // gridIndex, startColumn, columnCount, startRow, maxRows, maxCellTextLength
            var gridSegmentPrefix = new[]
                { typeof(int), typeof(int), typeof(int), typeof(long), typeof(int), typeof(int) };
            GetGridResultsSegmentAsyncMethod = ResolveDualSignature(
                IQueryEditorTabDataServiceBrokered,
                "GetGridResultsSegmentAsync",
                gridSegmentPrefix,
                out var gridRequiresMoniker);
            GridSegmentRequiresEditorMoniker = gridRequiresMoniker;

            GetAvailablePanesAsyncMethod = ResolveDualSignature(
                IQueryEditorTabDataServiceBrokered,
                "GetAvailablePanesAsync",
                Type.EmptyTypes,
                out var panesRequiresMoniker);
            AvailablePanesRequiresEditorMoniker = panesRequiresMoniker;

            GetCurrentConnectionAsyncMethod = ResolveMethod(
                ISqlEditorServiceBrokered,
                "GetCurrentConnectionAsync",
                requiredPositionalTypes: Type.EmptyTypes,
                requiresCancellationToken: true);

            GridResultsSegment_GridIndex = RequireProperty(gridSegmentType, "GridIndex");
            GridResultsSegment_TotalGridCount = RequireProperty(gridSegmentType, "TotalGridCount");
            GridResultsSegment_TotalRowCount = RequireProperty(gridSegmentType, "TotalRowCount");
            GridResultsSegment_TotalColumnCount = RequireProperty(gridSegmentType, "TotalColumnCount");
            GridResultsSegment_StartRow = RequireProperty(gridSegmentType, "StartRow");
            GridResultsSegment_ColumnNames = RequireProperty(gridSegmentType, "ColumnNames");
            GridResultsSegment_Rows = RequireProperty(gridSegmentType, "Rows");
            QueryResultsPaneInfo_PaneType = RequireProperty(paneInfoType, "PaneType");
            SqlEditorConnectionDetails_EditorMoniker = RequireProperty(connectionDetailsType, "EditorMoniker");

            QueryResultsPane_GridResults = Enum.Parse(paneEnumType, "GridResults");
            QueryEditorTabDataServiceMoniker = ResolveMonikerFrom(tabDataDescriptorsType, "QueryEditorTabDataService");
            SqlEditorServiceMoniker = ResolveMonikerFrom(editorServiceDescriptorsType, "SqlEditorService");
        }

        /// <summary>
        /// Resolves a method that may (22.6+) or may not (22.5) take a leading
        /// `string editorMoniker` before the given positional parameters.
        /// </summary>
        private static MethodInfo ResolveDualSignature(
            Type type, string name, Type[] positionalTypes, out bool requiresEditorMoniker)
        {
            try
            {
                requiresEditorMoniker = false;
                return ResolveMethod(type, name, positionalTypes, requiresCancellationToken: true);
            }
            catch (InvalidOperationException)
            {
                var withMoniker = new Type[positionalTypes.Length + 1];
                withMoniker[0] = typeof(string);
                Array.Copy(positionalTypes, 0, withMoniker, 1, positionalTypes.Length);
                requiresEditorMoniker = true;
                return ResolveMethod(type, name, withMoniker, requiresCancellationToken: true);
            }
        }

        private Type RequireType(string simpleName)
        {
            var full = Ns + "." + simpleName;
            var t = ContractsAssembly.GetType(full, throwOnError: false);
            if (t == null)
                throw new InvalidOperationException(
                    "BrokeredContracts type not found: " + full +
                    " (SSMS minor version may have moved it)");
            return t;
        }

        // Among all overloads with the given name, pick the one whose parameter list begins with
        // the required positional types (in order), contains a CancellationToken, and whose extra
        // parameters are all defaultable. This survives SSMS minor versions adding new optional
        // parameters without breaking older versions.
        private static MethodInfo ResolveMethod(
            Type type, string name, Type[] requiredPositionalTypes, bool requiresCancellationToken)
        {
            var candidates = type.GetMethods().Where(m => m.Name == name).ToList();
            if (candidates.Count == 0)
                throw new InvalidOperationException(
                    "BrokeredContracts method not found: " + type.FullName + "." + name);

            MethodInfo best = null;
            int bestExtras = int.MaxValue;
            foreach (var m in candidates)
            {
                var ps = m.GetParameters();
                if (ps.Length < requiredPositionalTypes.Length) continue;

                bool prefixOk = true;
                for (int i = 0; i < requiredPositionalTypes.Length; i++)
                    if (ps[i].ParameterType != requiredPositionalTypes[i]) { prefixOk = false; break; }
                if (!prefixOk) continue;

                if (requiresCancellationToken &&
                    !ps.Skip(requiredPositionalTypes.Length).Any(p => p.ParameterType == typeof(CancellationToken)))
                    continue;

                bool extrasOk = true;
                foreach (var p in ps.Skip(requiredPositionalTypes.Length))
                {
                    if (p.ParameterType == typeof(CancellationToken)) continue;
                    if (!p.HasDefaultValue && !p.ParameterType.IsValueType) { extrasOk = false; break; }
                }
                if (!extrasOk) continue;

                int extras = ps.Length - requiredPositionalTypes.Length - (requiresCancellationToken ? 1 : 0);
                if (extras < bestExtras) { best = m; bestExtras = extras; }
            }

            if (best == null)
                throw new InvalidOperationException(
                    "BrokeredContracts method not found: " + type.FullName + "." + name +
                    " with required prefix [" + string.Join(", ", requiredPositionalTypes.Select(t => t.Name)) + "]" +
                    (requiresCancellationToken ? " + CancellationToken" : ""));
            return best;
        }

        private static PropertyInfo RequireProperty(Type type, string name)
        {
            var p = type.GetProperty(name);
            if (p == null)
                throw new InvalidOperationException(
                    "BrokeredContracts property not found: " + type.FullName + "." + name);
            return p;
        }

        private static object ResolveMonikerFrom(Type descriptorsType, string memberName)
        {
            const BindingFlags bf =
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance;

            var prop = descriptorsType.GetProperty(memberName, bf);
            if (prop != null && prop.GetGetMethod(true)?.IsStatic == true)
            {
                var v = prop.GetValue(null);
                if (v != null) return v;
            }
            var field = descriptorsType.GetField(memberName, bf);
            if (field != null && field.IsStatic)
            {
                var v = field.GetValue(null);
                if (v != null) return v;
            }

            // Fallback: any moniker- or descriptor-shaped static member on the type.
            foreach (var p in descriptorsType.GetProperties(bf))
            {
                if (p.GetGetMethod(true)?.IsStatic != true) continue;
                if (!IsMonikerShape(p.PropertyType)) continue;
                var v = p.GetValue(null);
                if (v != null) return v;
            }
            foreach (var f in descriptorsType.GetFields(bf))
            {
                if (!f.IsStatic) continue;
                if (!IsMonikerShape(f.FieldType)) continue;
                var v = f.GetValue(null);
                if (v != null) return v;
            }

            throw new InvalidOperationException(
                "Could not locate " + descriptorsType.Name + "." + memberName + " moniker");
        }

        private static bool IsMonikerShape(Type t) =>
            t != null && (t.Name.IndexOf("ServiceMoniker", StringComparison.Ordinal) >= 0
                          || (t.FullName ?? string.Empty).IndexOf("ServiceRpcDescriptor", StringComparison.Ordinal) >= 0);
    }
}
