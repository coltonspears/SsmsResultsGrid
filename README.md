# Results View for SSMS 22

A VSIX extension for SQL Server Management Studio 22 that adds a filterable
**Results View** tab next to the native Results/Messages tabs. After each
query execution it captures the grid results through SSMS's own brokered
services (no per-cell reflection, no UI-thread scraping) and renders them in
a virtualized WPF `DataGrid` with live filtering, background sorting, match
highlighting, and CSV export.

![Results View](image.png)

## Features

- **Filterable grid in the results pane** — a "Results View" tab is injected
  directly into the query window's tab strip; filter state survives re-runs.
- **"Results to Filterable Grid" toggle** (*Tools* menu) — like the native
  *Results To* options: when checked, every completed query activates the
  Results View automatically; when unchecked, use *Tools → Show Results View*
  on demand.
- **Fast on large result sets** — rows are captured asynchronously in 5,000-row
  pages off the UI thread and stream into the grid progressively; filtering
  and sorting run on background threads with cancellation, so typing in the
  filter box never blocks the UI.
- **Multiple result sets** — a selector lists every grid; secondary sets load
  lazily on first selection.
- **Filter modes** — Contains / Starts With / Exact, optional case sensitivity,
  match highlighting, and a live row counter (`42 of 1,337 rows`).
- **Copy & CSV export** of the currently visible (filtered) rows.
- **Theming & localization** — colors resolve from the running VS theme
  (Light/Dark/Blue/High-Contrast) and all UI strings live in `.resx` resources.
- **Diagnostics** — failures are logged to a "Results View" pane in the Output
  window; the success path stays silent.

## How it works

SSMS 22 exposes query grid data through brokered services
(`IQueryEditorTabDataServiceBrokered`). The extension:

1. Observes (never consumes) the T-SQL *Execute* command with a priority
   command target, then hooks query-completion events on the editor control
   (with a bounded brokered-service poll as fallback).
2. Reads grid segments in pages via the brokered contract — one RPC per page
   of rows, resolved against both the SSMS 22.5 and 22.6+ method signatures.
3. Injects a `TabPage` hosting the WPF view into the query window's native
   tab strip (`TabPageHost`), falling back to a dockable tool window if the
   host layout is unrecognized.

The contracts DLL ships with SSMS and is loaded from the install directory at
runtime, so the VSIX is publicly distributable with zero configuration.

## Prerequisites

- **Visual Studio 2022** (17.14 or newer) with the
  *Visual Studio extension development* workload.
- **.NET Framework 4.8** developer pack.
- **SQL Server Management Studio 22** (for running the installed VSIX).

## Build & test

```powershell
# From the repo root:
msbuild SsmsResultsGrid.sln /t:Restore /p:Configuration=Release
msbuild SsmsResultsGrid.sln /t:Build   /p:Configuration=Release

# Core engine unit tests (filtering, sorting, CSV, view-model behavior):
dotnet test tests\SsmsResultsGrid.Core.Tests
```

The packaged extension lands at:

```
src\SsmsResultsGrid\bin\Release\SsmsResultsGrid.vsix
```

Or use the one-shot script that builds, uninstalls any previous version, and
installs into SSMS 22:

```powershell
.\build-reinstall.ps1 -Configuration Release -CloseSsms -RelaunchSsms
```

## Install

Double-click the `.vsix` and let VSIXInstaller target **SSMS 22**. If SSMS
is not offered in the installer, drop the file here and restart SSMS:

```
%LocalAppData%\Microsoft\SQL Server Management Studio\22.0\Extensions\
```

## Use

1. Open a query window and execute any `SELECT` (try `test-queries.sql`).
2. The **Results View** tab appears next to Results/Messages and activates
   automatically (toggle this via *Tools → Results to Filterable Grid*).
3. Type in the filter box — the grid narrows in real time with matches
   highlighted. Switch filter mode or case sensitivity as needed.
4. Click a column header to sort (asc → desc → unsorted) on a background thread.
5. Use **Copy** / **Export CSV** for the visible rows, **Refresh** to
   re-capture on demand.

## Project layout

```
src/SsmsResultsGrid/                    # VSIX (net48, VS SDK 17.14)
├── FilterableGridPackage.cs            # AsyncPackage entry point
├── FilterableGridPackage.vsct          # Tools-menu commands (show + toggle)
├── Commands/                           # ShowFilterableGridCommand, ToggleResultsToGridCommand
├── Services/
│   ├── Capture/                        # brokered-service client + paged grid reader
│   ├── Execution/                      # priority command target observing Execute
│   ├── InPaneTab/                      # per-query-window supervisor + tab injection
│   ├── Diagnostics/                    # Output-window pane
│   └── Settings/                       # persisted user settings
├── Views/                              # ResultsViewControl (MVVM view) + helpers
└── ToolWindows/                        # dockable fallback host

src/SsmsResultsGrid.Core/               # portable engine (netstandard2.0)
├── Models/                             # ResultRow, CapturedBatch
├── Filtering/ Sorting/ Export/         # cancellable filter/sort engines, CSV writer
├── Mvvm/                               # ObservableObject, RelayCommand, AsyncRelayCommand
└── ViewModels/                         # ResultsViewModel, ResultSetViewModel

tests/SsmsResultsGrid.Core.Tests/       # xUnit tests for the Core engine
```

## Known limitations

- **Row cap**: capture is capped at 100,000 rows per result set (configurable
  via the `SsmsResultsGrid` collection in the VS settings store) to keep the
  grid responsive; the status bar indicates truncation.
- **BLOB / image columns**: captured as their string representation.
- **Results-to-text / Results-to-file modes**: there is no grid to capture.
- **SSMS 20 and earlier** used the VS Isolated Shell and a different VSIX
  target schema. This project targets SSMS 22 only.

## License

MIT.
