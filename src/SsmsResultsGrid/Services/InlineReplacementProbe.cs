using System.Windows.Forms;

namespace SsmsResultsGrid.Services
{
    /// <summary>
    /// Safe, non-invasive probe for future replace-in-place experiments.
    /// We inspect the host chain but never mutate SSMS controls.
    /// </summary>
    internal static class InlineReplacementProbe
    {
        public static bool TryPrepareInlineReplacement(Control grid, out string reason)
        {
            reason = "probe-unavailable";
            if (grid == null)
            {
                reason = "grid-null";
                return false;
            }

            var parent = grid.Parent;
            if (parent == null)
            {
                reason = "grid-parent-null";
                return false;
            }

            var hostType = parent.GetType().FullName ?? parent.GetType().Name;
            reason = "host=" + hostType + "; childCount=" + parent.Controls.Count;
            return false;
        }
    }
}
