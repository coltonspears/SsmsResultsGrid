using System;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.VisualStudio.Shell;

namespace SsmsResultsGrid.Services.InPaneTab
{
    /// <summary>
    /// Resolves the WinForms TabControl (Results / Messages tab strip) inside an SSMS
    /// SqlScriptEditorControl docView. `TabPageHost` is a public property on SSMS 22;
    /// private-field fallbacks cover potential renames in minor versions.
    /// </summary>
    internal static class TabPageHostResolver
    {
        public static TabControl Resolve(object docView)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (docView == null) throw new ArgumentNullException(nameof(docView));

            var hostObj = GetTabPageHost(docView)
                ?? throw new InvalidOperationException("SqlScriptEditorControl.TabPageHost is null.");

            return AsTabControl(hostObj)
                ?? throw new InvalidOperationException(
                    "TabPageHost is neither a TabControl nor a container of one: " + hostObj.GetType().FullName);
        }

        private static object GetTabPageHost(object docView)
        {
            var t = docView.GetType();

            var prop = t.GetProperty("TabPageHost",
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            if (prop != null && prop.CanRead) return prop.GetValue(docView);

            var field = FindField(t, "tabPagesHost") ?? FindField(t, "m_tabPagesHost");
            if (field != null) return field.GetValue(docView);

            throw new InvalidOperationException(
                "Could not find TabPageHost property or tabPagesHost field on " + t.FullName);
        }

        private static TabControl AsTabControl(object obj)
        {
            if (obj is TabControl direct) return direct;
            if (obj is Control container) return FindFirstTabControl(container);
            return null;
        }

        private static TabControl FindFirstTabControl(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                if (child is TabControl tc) return tc;
                var deep = FindFirstTabControl(child);
                if (deep != null) return deep;
            }
            return null;
        }

        internal static FieldInfo FindField(Type t, string name)
        {
            for (var cursor = t; cursor != null; cursor = cursor.BaseType)
            {
                var f = cursor.GetField(name,
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                if (f != null) return f;
            }
            return null;
        }
    }
}
