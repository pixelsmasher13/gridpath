# GridPath MCP v0 — plan

Branch: `mcp-v0` (off `main` at `a403a6a`). Written 2026-09-29.

## What v0 is

`npx gridpath mcp` — a Node MCP server that lets any agent host (Claude Code,
Cursor, Codex) edit `.xlsx` files through GridPath's tools, with a local review
tab before anything is saved. The file on disk is the source of truth; saves go
through the surgical patcher, so untouched parts come back byte-identical.

Pitch: **let your agent edit Excel files without butchering them.**

Not in v0: `edit_workbook` (our own agent loop), the in-app embed, the bulk CLI.

## Architecture

```
host agent (Claude Code …)
   │ MCP (stdio)
   ▼
packages/gridpath  (Node, TypeScript)
   ├─ tools/        describe_workbook · find_rows · read_range · run_script
   │                set_cell/range/format · structure ops · stage_data
   │                fetch_web · edgar_lookup · keep_pages · read_source
   ├─ workbook/     open(bytes) → IronCalc model (WASM); mutations applied to
   │                the model; evaluate; readback; pending batch on disk
   ├─ save/         buildWorkbookPatch(live, baseline, mirror) → engine.patch()
   └─ review/       localhost:<port>/batch/<id> — existing React+Univer review
                    screen, HTTP instead of Tauri invoke (phase 3)
crates/engine-wasm (Rust → wasm32, wasm-bindgen --target nodejs)
   ├─ ironcalc (vendored fork)   import · evaluate · edit · structure ops · snapshot
   └─ xlsx_patch                 apply_patch_json
```

**One reader.** v0 reads workbooks through IronCalc only (no ExcelJS). The
WASM exposes `snapshot()` producing the Univer-shaped `IWorkbookData` the TS
modules already consume (`cellData[r][c] = {v, f, s}` with resolved style
objects, merges, col widths, row heights, freeze, hidden, defined names).
Known gaps vs `xlsxImport.ts` to verify against the corpus tests: rich text
flattening, hyperlink text, shared-formula translation, date serials,
`_xlfn.` stripping, `_xll.`/external-ref pinning, CF/DV (not needed by tools).

**Mutations apply to the model, not a snapshot copy.** `set_user_input`,
`insert_rows`, `delete_columns`, `rename_sheet`, `new_sheet`, widths, freeze,
defined names — IronCalc shifts references itself. After each write tool:
`evaluate()`, readback from the model (same JSON the app returns today,
including `row_map`).

**Save reuses tested code.** `buildSaveMirror` (moved out of the React file)
+ `buildWorkbookPatch(liveSnapshot, baselineSnapshot, mirror)` from
`surgicalPatch.ts`, unchanged, then `engine.patch(originalBytes, patchJson)`.
Baseline = the snapshot taken at open.

**Pending batch on disk.** `<dir>/.gridpath/<file>.batches/<id>.json` holding
the `ChangeBatch` (prompt, justification, mutations, status). `save_workbook`
accepts pending batches and writes the file; the review tab can reject
individual batches first. Nothing is written to the `.xlsx` until save.

## Tools exposed over MCP (v0)

Same names and JSON schemas as `tools.rs` (emitted to `tools.json` at build
by a small Rust bin so the two never drift). `done` is dropped (no loop);
`read_reference` deferred; `web_search` is host-provided.

Host-facing additions: `list_batches`, `reject_batch`, `save_workbook`,
`review_url`.

## Phases

### Phase 1 — engine package (≈3 days)
1. Extract `src-tauri/src/engine/workbook/xlsx_patch` into `crates/xlsx-patch`
   (std-only today); `src-tauri` depends on it. No behaviour change.
2. `crates/engine-wasm`: wasm-bindgen exports — `open(bytes) → handle`,
   `snapshot`, `evaluate`, `setInput/setFormat/insertRows/…`, `cellValue`,
   `patch(bytes, json)`. Build script → `packages/gridpath/engine/`.
3. Port the scratchpad probe into a vitest that opens `eval/fixtures/*.xlsx`
   and `project-159` (local-only), evaluates, patches, asserts part identity.

### Phase 2 — tools + MCP server (≈4 days)
4. `packages/gridpath` (ESM, `@modelcontextprotocol/sdk` 1.31, zod). Move the
   pure modules under `src/screens/SpreadsheetScreen/agent/` that the package
   needs into `packages/gridpath/src/` and re-import them from the app (types
   `CellFormatShape`/`SaveMirror` move out of `UniverGrid.tsx` first).
5. Workbook session: open → model → snapshot → preview (`captureContext` on
   snapshot; style resolver instead of `getCellFormat`) → index.
6. Tool handlers on the model. `run_script` in `worker_threads` with the same
   5 s / 20k-cell limits (the current Node fallback has no isolation — do not
   ship it).
7. Port `fetch_web` (fetch + `html-to-text`, `unpdf` for PDFs, keep column
   alignment — don't collapse whitespace inside PDF text), `edgar_lookup`
   (one SEC User-Agent), pagination/`source_id`/`keep_pages` from
   `commands.rs`. Gating becomes advisory text in tool results (no loop to
   enforce it).
8. `gridpath mcp` bin, stdio transport; config snippets for Claude Code,
   Cursor, Codex in the README.

### Phase 3 — review tab (≈4 days)
9. Serve the existing review screen from the MCP process at
   `localhost:<port>/batch/<id>`; replace the ~25 `invoke()` calls
   (`sessionDb`, `workbookIo`, `agentClient`, `settingsApi`) with HTTP to the
   local server. Accept / reject / save from the tab.
10. Tool results carry `review_url` after every write.

### Phase 4 — test before launch (≈2 days)
11. Edit eval: Claude Code + GridPath MCP vs Claude Code alone. Pass = clean
    parts every run; wall-clock in the 10–30 s range.
12. Univer CLI as a third lane on the same fixtures + project-159.
13. Five real files from people who aren't us.

## Decisions taken
- Off `main`, not off `spike/mcp-server`: the spike is Rust-in-app over HTTP;
  v0 is a Node package over stdio. Its batch shape and tool list carry over.
- No MCP sampling: Claude Code doesn't support it as a client (issue #1785).
  Irrelevant to v0; matters only if `edit_workbook` ships later.
- Univer stays confined to the review tab. Packages 1–2 have no Univer dep.
- The `parts_preserved` check from the eval harness becomes a unit test in
  the package.

## Open risks
- IronCalc reader vs ExcelJS feature gap (phase 1.3 measures it).
- IronCalc API coverage for structure ops the tools need (merges, notes,
  hidden rows/cols, tab color) — anything missing goes through the patch
  only, invisible to readback until save.
- Review tab decoupling touches the React file; keep the desktop app green
  throughout (vitest + `cargo check` on every step).
