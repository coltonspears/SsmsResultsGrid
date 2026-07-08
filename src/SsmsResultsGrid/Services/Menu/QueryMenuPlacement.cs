using System;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using SsmsResultsGrid.Services.Diagnostics;

namespace SsmsResultsGrid.Services.Menu
{
    /// <summary>
    /// Moves the extension's commands from the fallback Tools-menu spot into
    /// SSMS's Query &gt; Results To submenu at runtime via DTE command bars.
    ///
    /// The Query menu is owned by the SQL editor package and its vsct group IDs are
    /// not public, so static CommandPlacement is not an option. Placement is retried
    /// lazily (the Query command bar may not exist until a SQL editor has loaded)
    /// and is idempotent across sessions: command-bar customizations persist in the
    /// user's CmdUI cache, so existing controls are detected before adding.
    /// </summary>
    internal sealed class QueryMenuPlacement
    {
        private static readonly string[] CommandNames =
        {
            "SsmsResultsGrid.ToggleResultsToGrid",
            "SsmsResultsGrid.ShowResultsView",
        };

        private readonly AsyncPackage _package;
        private readonly DiagnosticsPane _pane;
        private bool _placed;
        private int _attempts;

        public QueryMenuPlacement(AsyncPackage package, DiagnosticsPane pane)
        {
            _package = package;
            _pane = pane;
        }

        /// <summary>True once the commands live under Query &gt; Results To (Tools copies hidden).</summary>
        public bool Placed => _placed;

        /// <summary>Attempts placement; safe to call repeatedly from the main thread.</summary>
        public void TryPlace()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_placed || _attempts >= 10) return;
            _attempts++;

            try
            {
                if (!(_package.GetService<Microsoft.VisualStudio.Shell.Interop.SDTE, DTE2>() is DTE2 dte)) return;
                if (!(dte.CommandBars is CommandBars bars)) return;

                var resultsTo = FindResultsToPopup(bars);
                if (resultsTo == null) return; // Query menu not built yet — retry later

                foreach (var name in CommandNames)
                {
                    PlaceCommand(dte, resultsTo, name);
                }

                RemoveFromToolsMenu(bars, dte);
                _placed = true;
            }
            catch (Exception ex)
            {
                // Localized captions or a reshaped menu tree land here; Tools fallback stays.
                _pane?.WriteFailure(nameof(QueryMenuPlacement), ex);
                _attempts = int.MaxValue;
            }
        }

        private static CommandBarPopup FindResultsToPopup(CommandBars bars)
        {
            CommandBar queryBar;
            try
            {
                queryBar = bars["Query"];
            }
            catch
            {
                return null;
            }
            if (queryBar == null) return null;

            foreach (CommandBarControl control in queryBar.Controls)
            {
                if (control is CommandBarPopup popup &&
                    Normalize(popup.Caption).IndexOf("results to", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return popup;
                }
            }
            return null;
        }

        private void PlaceCommand(DTE2 dte, CommandBarPopup resultsTo, string commandName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var command = dte.Commands.Item(commandName);
            if (command == null) return;

            // Idempotency: skip when a control for this command already exists
            // (persisted from an earlier session's AddControl).
            string caption = Normalize(command.Name == CommandNames[0]
                ? Resources.Strings.CmdToggleResultsToGrid
                : Resources.Strings.CmdShowResultsView);
            foreach (CommandBarControl existing in resultsTo.Controls)
            {
                if (string.Equals(Normalize(existing.Caption), caption, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            var added = (CommandBarControl)command.AddControl(
                resultsTo.CommandBar, resultsTo.Controls.Count + 1);
            if (added != null && command.Name == CommandNames[0])
            {
                added.BeginGroup = true; // separator between native Results To modes and ours
            }
        }

        private void RemoveFromToolsMenu(CommandBars bars, DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandBar toolsBar;
            try
            {
                toolsBar = bars["Tools"];
            }
            catch
            {
                return;
            }
            if (toolsBar == null) return;

            var ourCaptions = new[]
            {
                Normalize(Resources.Strings.CmdToggleResultsToGrid),
                Normalize(Resources.Strings.CmdShowResultsView),
            };

            for (int i = toolsBar.Controls.Count; i >= 1; i--)
            {
                var control = toolsBar.Controls[i];
                var caption = Normalize(control.Caption);
                foreach (var ours in ourCaptions)
                {
                    if (string.Equals(caption, ours, StringComparison.OrdinalIgnoreCase))
                    {
                        try { control.Delete(false); } catch { /* leave it visible */ }
                        break;
                    }
                }
            }
        }

        private static string Normalize(string caption) =>
            (caption ?? string.Empty).Replace("&", string.Empty).Trim();
    }
}
