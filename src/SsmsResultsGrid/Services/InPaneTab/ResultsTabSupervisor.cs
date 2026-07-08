using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Threading;
using SsmsResultsGrid.Core.Models;
using SsmsResultsGrid.Core.ViewModels;
using SsmsResultsGrid.Services.Capture;
using SsmsResultsGrid.Services.Diagnostics;
using SsmsResultsGrid.Services.Execution;
using SsmsResultsGrid.Services.Settings;
using SsmsResultsGrid.Views;
using Task = System.Threading.Tasks.Task;

namespace SsmsResultsGrid.Services.InPaneTab
{
    /// <summary>
    /// One supervisor per SSMS query window (SqlScriptEditorControl docView). Owns the
    /// persistent <see cref="ResultsViewModel"/>, the injected "Results View" TabPage
    /// lifecycle, and the query-completion event hooks that drive auto-capture.
    ///
    /// Stored in a ConditionalWeakTable keyed on the docView so the supervisor (and its
    /// event subscriptions) are collected together with the query window — no explicit
    /// unhook needed.
    /// </summary>
    internal sealed class ResultsTabSupervisor
    {
        private const string TabPageName = "__SsmsResultsGrid_ResultsView";

        /// <summary>Trailing debounce for bursts of completion events from one execution.</summary>
        private static readonly TimeSpan CompletionCoalesceDelay = TimeSpan.FromMilliseconds(300);

        // Heuristic completion-event name patterns (StatisticsParserExtension-proven).
        private static readonly string[] EventNamePatterns =
            { "Completed", "Executed", "Finished", "Stopped", "Done" };

        private static readonly ConditionalWeakTable<object, ResultsTabSupervisor> Supervisors =
            new ConditionalWeakTable<object, ResultsTabSupervisor>();

        private readonly object _docView;
        private readonly AsyncPackage _package;
        private readonly DiagnosticsPane _pane;
        private readonly ExtensionSettings _settings;
        private readonly ResultsViewModel _viewModel;
        private readonly WpfUiDispatcher _dispatcher;

        private TabPage _tabPage;
        private ResultsViewControl _view;
        private int _hookedEventCount;
        private bool _loggedHooks;
        private CancellationTokenSource _completionDebounceCts;
        private CancellationTokenSource _captureCts;

        public static ResultsTabSupervisor GetOrCreate(
            object docView, AsyncPackage package, DiagnosticsPane pane, ExtensionSettings settings)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return Supervisors.GetValue(docView, dv => new ResultsTabSupervisor(dv, package, pane, settings));
        }

        private ResultsTabSupervisor(
            object docView, AsyncPackage package, DiagnosticsPane pane, ExtensionSettings settings)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _docView = docView;
            _package = package;
            _pane = pane;
            _settings = settings;
            _dispatcher = new WpfUiDispatcher(pane);
            _viewModel = new ResultsViewModel(_dispatcher, ex => _pane?.WriteFailure("ViewModel", ex));
            _viewModel.SetRefreshHandler(() => CaptureAsync(activateTab: true, manual: true));
            _viewModel.ResultSetLoadRequested += OnResultSetLoadRequested;

            HookQueryCompletionEvents();
        }

        /// <summary>Called by the execute trigger just before SSMS runs a query in this window.</summary>
        public void OnExecuteObserved()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CancelInFlightCapture();

            if (_hookedEventCount == 0)
            {
                // Event-name heuristic found nothing on this SSMS build: fall back to a
                // bounded, off-UI-thread poll of the brokered pane list.
                StartPaneAvailabilityFallback();
            }
        }

        /// <summary>Manual entry point (Tools menu / Refresh button).</summary>
        public Task ShowAsync() => CaptureAsync(activateTab: true, manual: true);

        // ---- completion detection ----

        private void HookQueryCompletionEvents()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                int hooked = HookEventsOn(_docView);

                var resultsControlField = TabPageHostResolver.FindField(_docView.GetType(), "m_sqlResultsControl");
                var resultsControl = resultsControlField?.GetValue(_docView);
                if (resultsControl != null)
                {
                    hooked += HookEventsOn(resultsControl);
                }

                _hookedEventCount = hooked;
                if (hooked == 0)
                {
                    _pane?.WriteInfo(
                        "No query-completion events matched the name heuristic on this SSMS build; " +
                        "auto-refresh will use a bounded brokered-service poll instead.");
                }
            }
            catch (Exception ex)
            {
                _pane?.WriteFailure(nameof(HookQueryCompletionEvents), ex);
            }
        }

        private int HookEventsOn(object instance)
        {
            int hooked = 0;
            var hookedNames = new System.Text.StringBuilder();
            for (var t = instance.GetType(); t != null && t != typeof(object); t = t.BaseType)
            {
                EventInfo[] events;
                try
                {
                    events = t.GetEvents(BindingFlags.Public | BindingFlags.NonPublic |
                                         BindingFlags.Instance | BindingFlags.DeclaredOnly);
                }
                catch { continue; }

                foreach (var evt in events)
                {
                    if (evt.EventHandlerType == null) continue;
                    bool match = EventNamePatterns.Any(p =>
                        evt.Name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!match) continue;

                    try
                    {
                        var handler = BuildHandler(evt.EventHandlerType);
                        evt.AddEventHandler(instance, handler);
                        hooked++;
                        hookedNames.Append(t.Name).Append('.').Append(evt.Name).Append("; ");
                    }
                    catch (Exception ex)
                    {
                        _pane?.WriteFailure("AddEventHandler " + t.Name + "." + evt.Name, ex);
                    }
                }
            }

            if (hooked > 0 && !_loggedHooks)
            {
                _loggedHooks = true;
                _pane?.WriteInfo("Hooked completion events: " + hookedNames);
            }
            return hooked;
        }

        // Builds a delegate of the event's actual delegate type whose body forwards to
        // OnQueryCompletionEvent, adapting arbitrary 2-parameter signatures.
        private Delegate BuildHandler(Type delegateType)
        {
            var invokeMethod = delegateType.GetMethod("Invoke")
                ?? throw new InvalidOperationException("Delegate type has no Invoke method.");
            var parameters = invokeMethod.GetParameters();
            if (parameters.Length != 2)
                throw new InvalidOperationException("Unsupported event signature (need 2 parameters).");

            var senderParam = Expression.Parameter(parameters[0].ParameterType, "sender");
            var argsParam = Expression.Parameter(parameters[1].ParameterType, "args");

            var thisConst = Expression.Constant(this);
            var handlerMethod = typeof(ResultsTabSupervisor).GetMethod(
                nameof(OnQueryCompletionEvent), BindingFlags.NonPublic | BindingFlags.Instance);

            var callHandler = Expression.Call(thisConst, handlerMethod);
            return Expression.Lambda(delegateType, callHandler, senderParam, argsParam).Compile();
        }

        // May fire on any thread, several times per execution — coalesce with a trailing delay.
        private void OnQueryCompletionEvent()
        {
            var cts = new CancellationTokenSource();
            var previous = Interlocked.Exchange(ref _completionDebounceCts, cts);
            previous?.Cancel();

            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await Task.Delay(CompletionCoalesceDelay, cts.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return; // superseded by a later event in the same burst
                }

                if (!_settings.ResultsToFilterGrid) return;

                try
                {
                    await CaptureAsync(activateTab: _settings.ResultsToFilterGrid, manual: false);
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    _pane?.WriteFailure(nameof(OnQueryCompletionEvent), ex);
                }
            });
        }

        private void StartPaneAvailabilityFallback()
        {
            var captureCts = ResetCaptureCts();
            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    // SSMS tears the results pane down at query start and recreates it on
                    // completion, so "GridResults pane exists" approximates completion.
                    await Task.Delay(500, captureCts.Token).ConfigureAwait(false);
                    for (int attempt = 0; attempt < 40; attempt++)
                    {
                        captureCts.Token.ThrowIfCancellationRequested();
                        bool available;
                        using (var client = await GridBrokeredClient.CreateAsync(_package, captureCts.Token))
                        {
                            await TaskScheduler.Default;
                            available = await client.IsGridResultsPaneAvailableAsync(captureCts.Token);
                        }
                        if (available)
                        {
                            await CaptureAsync(activateTab: _settings.ResultsToFilterGrid, manual: false);
                            return;
                        }
                        await Task.Delay(500, captureCts.Token).ConfigureAwait(false);
                    }
                }
                catch (OperationCanceledException) { }
                catch (NoActiveEditorException) { }
                catch (Exception ex)
                {
                    _pane?.WriteFailure(nameof(StartPaneAvailabilityFallback), ex);
                }
            });
        }

        // ---- capture pipeline ----

        private async Task CaptureAsync(bool activateTab, bool manual)
        {
            var captureCts = ResetCaptureCts();
            var ct = captureCts.Token;

            await _package.JoinableTaskFactory.SwitchToMainThreadAsync(ct);

            // SSMS 22.5's tab-data service is implicitly bound to the ACTIVE editor.
            // Capturing while another window is focused would read the wrong grids.
            var activeDocView = ActiveDocViewResolver.GetActiveSqlEditorDocView(_package);
            if (!ReferenceEquals(activeDocView, _docView))
            {
                return; // this window's next execute re-triggers capture
            }

            GridBrokeredClient client = null;
            try
            {
                client = await GridBrokeredClient.CreateAsync(_package, ct);

                await TaskScheduler.Default;
                bool hasGridResults = await client.IsGridResultsPaneAvailableAsync(ct);
                if (!hasGridResults)
                {
                    _dispatcher.Post(() =>
                    {
                        if (ct.IsCancellationRequested) return;
                        _viewModel.SetNoGridResults();
                    });
                    return;
                }

                bool tabEnsured = false;
                long loadedRows = 0;
                var progress = new DelegateProgress<CapturedBatch>(batch => _dispatcher.Post(() =>
                {
                    if (ct.IsCancellationRequested) return;

                    if (batch.ColumnNames != null)
                    {
                        _viewModel.BeginCapture(batch.TotalGridCount);
                        _viewModel.GetResultSet(0)?.BeginLoad(batch.ColumnNames, batch.TotalRowCount);
                    }

                    var set = _viewModel.GetResultSet(0);
                    if (set == null) return;

                    if (batch.Rows.Count > 0)
                    {
                        set.AppendRows(batch.Rows, batch.StartRow);
                        loadedRows += batch.Rows.Count;
                        _viewModel.ReportLoadProgress(loadedRows, Math.Min(batch.TotalRowCount, _settings.MaxRows));
                    }

                    if (batch.ColumnNames != null && !tabEnsured)
                    {
                        tabEnsured = true;
                        EnsureTabAndBind(activateTab);
                    }

                    if (batch.IsFinal)
                    {
                        set.CompleteLoad(batch.IsTruncated);
                    }
                }));

                await GridResultsReader.ReadGridAsync(
                    client, gridIndex: 0, maxRows: _settings.MaxRows, maxCellChars: _settings.MaxCellChars,
                    progress, ct);

                // Secondary result sets: metadata now, rows on first selection.
                var firstProbe = await GridResultsReader.ProbeGridAsync(client, 0, _settings.MaxCellChars, ct);
                int gridCount = firstProbe?.TotalGridCount ?? 1;
                for (int i = 1; i < gridCount; i++)
                {
                    var probe = await GridResultsReader.ProbeGridAsync(client, i, _settings.MaxCellChars, ct);
                    if (probe == null) continue;
                    int gridIndex = i;
                    _dispatcher.Post(() =>
                    {
                        if (ct.IsCancellationRequested) return;
                        _viewModel.GetResultSet(gridIndex)?.SetPendingMetadata(probe.ColumnNames, probe.TotalRowCount);
                    });
                }

                _dispatcher.Post(() =>
                {
                    if (ct.IsCancellationRequested) return;
                    _viewModel.CompleteCapture(DateTime.Now);
                });
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (NoActiveEditorException)
            {
                // Focus moved away mid-capture — quiet; next execute re-arms.
            }
            catch (Exception ex)
            {
                _pane?.WriteFailure(manual ? "ManualCapture" : "AutoCapture", ex);
                _dispatcher.Post(() => _viewModel.SetCaptureFailed());
            }
            finally
            {
                client?.Dispose();
            }
        }

        private void OnResultSetLoadRequested(object sender, ResultSetViewModel set)
        {
            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                // Deliberately NOT tied to the capture CTS: loading a secondary set
                // must not cancel a grid-0 capture that may still be streaming.
                var ct = _package.DisposalToken;
                GridBrokeredClient client = null;
                try
                {
                    await _package.JoinableTaskFactory.SwitchToMainThreadAsync(ct);
                    client = await GridBrokeredClient.CreateAsync(_package, ct);
                    await TaskScheduler.Default;

                    var progress = new DelegateProgress<CapturedBatch>(batch => _dispatcher.Post(() =>
                    {
                        if (ct.IsCancellationRequested) return;
                        if (batch.ColumnNames != null)
                        {
                            set.BeginLoad(batch.ColumnNames, batch.TotalRowCount);
                        }
                        if (batch.Rows.Count > 0)
                        {
                            set.AppendRows(batch.Rows, batch.StartRow);
                        }
                        if (batch.IsFinal)
                        {
                            set.CompleteLoad(batch.IsTruncated);
                        }
                    }));

                    await GridResultsReader.ReadGridAsync(
                        client, set.GridIndex, _settings.MaxRows, _settings.MaxCellChars, progress, ct);
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    _pane?.WriteFailure("LoadResultSet " + set.GridIndex, ex);
                }
                finally
                {
                    client?.Dispose();
                }
            });
        }

        // ---- tab injection ----

        /// <summary>
        /// (Re-)creates our TabPage inside the query window's Results/Messages strip.
        /// SSMS disposes the strip's pages on query start, so a fresh TabPage,
        /// ElementHost, and view are built each time; state lives in the persistent
        /// view model that gets re-bound as DataContext.
        /// </summary>
        private void EnsureTabAndBind(bool activate)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var tabControl = TabPageHostResolver.Resolve(_docView);

                bool missing = _tabPage == null || _tabPage.IsDisposed || !tabControl.TabPages.Contains(_tabPage);
                if (missing)
                {
                    if (_tabPage != null && !_tabPage.IsDisposed)
                    {
                        _tabPage.Dispose();
                    }

                    _view = new ResultsViewControl { DataContext = _viewModel };
                    var host = new ElementHost { Dock = DockStyle.Fill, Child = _view };
                    _tabPage = new TabPage(Resources.Strings.TabTitle) { Name = TabPageName };
                    _tabPage.Controls.Add(host);
                    tabControl.TabPages.Insert(0, _tabPage);
                }

                if (activate)
                {
                    tabControl.SelectedTab = _tabPage;
                }
            }
            catch (Exception ex)
            {
                _pane?.WriteFailure(nameof(EnsureTabAndBind), ex);
                FallBackToToolWindow();
            }
        }

        private void FallBackToToolWindow()
        {
            _ = _package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await _package.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);
                    var window = await _package.ShowToolWindowAsync(
                        typeof(ToolWindows.FilterableGridToolWindow), 0, create: true,
                        cancellationToken: _package.DisposalToken) as ToolWindows.FilterableGridToolWindow;
                    window?.Bind(_viewModel);
                }
                catch (Exception ex)
                {
                    _pane?.WriteFailure(nameof(FallBackToToolWindow), ex);
                }
            });
        }

        // ---- plumbing ----

        private CancellationTokenSource ResetCaptureCts()
        {
            var next = CancellationTokenSource.CreateLinkedTokenSource(_package.DisposalToken);
            var previous = Interlocked.Exchange(ref _captureCts, next);
            previous?.Cancel();
            return next;
        }

        private void CancelInFlightCapture()
        {
            var previous = Interlocked.Exchange(ref _captureCts, null);
            previous?.Cancel();
        }

        private sealed class DelegateProgress<T> : IProgress<T>
        {
            private readonly Action<T> _handler;
            public DelegateProgress(Action<T> handler) => _handler = handler;
            public void Report(T value) => _handler(value);
        }
    }
}
