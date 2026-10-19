# Memory budget (desktop)

The calc engine (IronCalc) runs in the Rust process and holds a 2.3M-cell
model in roughly 1–1.5 GB with no practical ceiling. What runs out of memory
on huge workbooks is WebKit's content process: Univer's model of the grid plus
everything JavaScript keeps around it. macOS terminates that process under
pressure; the app then loses the tab (the unsaved-work journal keeps the
edits — see unsaved-work-journal.md).

## Estimated budget on a 2.27M-cell workbook

| Resident in the content process | Estimate |
|---|---|
| Univer's model of the cells, render and matrix structures | 600 MB – 1 GB |
| Load transient: 136 MB snapshot string + parsed JSON + mapped objects | ~700 MB |
| Save baseline (`buildCellBaseline`): 2.3M-entry Map with formula text | 300–400 MB |
| Univer calc worker holding a mirrored copy of the sheets (pre-2026-10-06) | ~300 MB |
| Agent turns: snapshots, context strings, validation scans, tool results | spiky |
| ExcelJS full re-parse on the fallback (non-surgical) save path | +1.2 GB |

## Instrumentation (2026-10-06)

`memory_report` (Rust): resident size of this process and of the WebKit
content process — WebKit spawns that over XPC, so it is not our child; the
pid comes from WKWebView's `_webProcessIdentifier`. `memoryMonitor.ts` polls
it every 10 s, logs `[memory] webview X GB (pid N) · engine/app Y GB`, logs a
labelled reading after every open, and raises one warning toast when the
content process passes 2.5 GB (re-armed under 2 GB).

## Reductions

- **No Univer calc worker in engine mode** (2026-10-06). The worker's engine
  ran with `notExecuteFormula` yet Univer still mirrored the workbook into it.
  The preset only disables the main-thread engine when a worker is supplied,
  so UniverGrid builds the preset without one and flips `notExecuteFormula`
  on the formula plugins itself. Kill switch (restores the worker):
  `localStorage.setItem("gridpath.univerWorker", "1")`.

## Next levers, in order of payoff

1. Save baseline into the engine: the engine knows the file's values and
   which cells changed since load; the in-place patch can come from the
   engine's edit record plus the save mirror, removing the 2.3M-entry Map and
   the save-time full diff.
2. Refuse the ExcelJS fallback writer above a size threshold (it is a 1.2 GB
   allocation in a process already near its limit); Save As through the
   surgical path is the safe route.
3. Lazy sheets: only the active sheet's cells in Univer, others on tab
   switch; agent reads served from the engine. The two supply sheets are
   1.6M of the 2.27M cells.
