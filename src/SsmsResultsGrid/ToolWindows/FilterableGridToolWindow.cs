using System;
using System.Data;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace SsmsResultsGrid.ToolWindows
{
    [Guid(PackageGuids.ToolWindowGuidString)]
    public sealed class FilterableGridToolWindow : ToolWindowPane
    {
        private readonly FilterableGridControl _control;

        public FilterableGridToolWindow() : base(null)
        {
            Caption = "Filterable Results";
            _control = new FilterableGridControl();
            Content = _control;
        }

        public void LoadData(DataTable table)
        {
            if (table == null) return;
            _control.LoadData(table);
        }

        public void LoadCaptureResult(DataTable table, string failureReason, string contextKey)
        {
            if (!string.IsNullOrWhiteSpace(contextKey))
            {
                Caption = "Filterable Results - " + System.IO.Path.GetFileName(contextKey);
            }
            else
            {
                Caption = "Filterable Results";
            }

            _control.LoadCaptureResult(table, failureReason);
        }
    }
}
