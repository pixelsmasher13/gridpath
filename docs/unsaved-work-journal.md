# Unsaved-work journal (desktop)

**Status:** shipped 2026-10-06. File-backed workbooks only; untitled drafts
keep their 30-second `.gpsnap` auto-snapshot.

## Why

Until Cmd-S, edits to an opened `.xlsx` lived only in memory. A crash — the
app running out of memory on a 2M-cell model, a kernel panic — lost them, and
the file most likely to exhaust memory was exactly the one with no protection.
Snapshotting a 2M-cell workbook every 30 s is itself a memory and time problem,
so file-backed tabs journal *what changed* instead.

## What is journaled

`src/screens/SpreadsheetScreen/journal.ts`, one JSON file per open workbook
under `<app_data>/journals/wb-<hash of path>.json`, rewritten atomically at
most every 2 s when something changed, deleted on save:

- **edits** — in order: every cell write, row/column insert/delete and sheet
  add/delete/rename that reaches the calc-engine mirror
  (`ironcalcEngine.ts`: `engineEditBatch`, `engineStructural`,
  `engineSheetOp`). The mirror sees Univer's command stream, so user typing,
  pastes, undo/redo and agent tool calls are all captured, as raw input text
  with sheet names. A command the mirror cannot decode (sort, move range)
  records a `resync` marker.
- **batches** — the agent's review batches not yet saved (`persisted: false`,
  status pending/accepted), so the review UI and the in-place saver
  (formats, merges, sheet ops) get them back.
- **manual** — the toolbar mirrors the saver keeps (formats, borders, merges,
  row/col ops, freeze, visibility, sheet ops).
- **stamp** — `size:mtime` of the file the journal belongs to.

## Recovery

On opening a path (menu open or session restore), if a journal exists for it
and the file's stamp still matches, a native dialog offers *Restore* /
*Discard* with a summary (`describeJournal`). Restore replays `edits` through
the grid handle in order, then the style-level mutations of the batches (their
cell writes and structure already came through the edit stream), then the
toolbar mirrors; the batches rejoin tab state as unsaved and the tab is marked
dirty. The replay runs through Univer's command stream, so the engine mirror
and a fresh journal pick the recovered state up as new unsaved work. A journal
containing a `resync` marker is replayed as far as it goes and reported as
partial.

A stale journal (file changed since) is deleted silently. Closing a dirty tab
keeps its journal (the next open offers it); closing a clean tab or choosing
*Discard* in the quit guard deletes it.

## Limitations

- Numeric-looking text typed into a cell (`0123`) replays as a number.
- Replayed batch formatting uses the coordinates recorded at the time; if
  later journaled structure edits moved those cells, the formatting lands on
  the old coordinates.
- Column widths, row heights and auto-filters set from the toolbar are not
  replayed (they are still carried for the save mirror).
- When the IronCalc engine is off (kill switch) only batches and toolbar
  mirrors are journaled — there is no edit stream.
