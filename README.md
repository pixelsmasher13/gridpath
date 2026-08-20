# GridPath

A desktop AI agent for Excel. You prompt it, it builds and edits real `.xlsx` files — with every change presented as a reviewable diff you can accept or reject before anything touches your workbook.

Works with your existing Claude Pro/Max or ChatGPT Plus/Pro subscription — no API key required. API keys are also supported for users who prefer that path.

## What it does

You open a workbook, type a request, and watch the agent work. Each tool call (set a range of values, format cells, fetch a web page for data, merge cells, freeze panes) applies live so you can preview the result, but every edit is captured as a pending mutation. A single "Accept" commits the batch; "Reject" rewinds it cell-by-cell back to the exact prior values, formulas, number formats, and fills.

Under the hood the agent has read/write access to a Univer.js-backed grid (real formula engine, real A1 references, real number formats) and can call:

- `set_cell` / `set_range` — write values and formulas in A1 notation
- `set_format` — bulk format ops (bold, currency, percent, colors, alignment, etc.) in one call
- `merge_cells` / `freeze_panes` / `set_column_width` / `set_row_height`
- `create_sheet` / `rename_sheet`
- `clear_range`
- `read_range` — let the agent sanity-check its own work before composing dependent formulas
- `fetch_web` — pull source data (SEC filings, Macrotrends, etc.) directly

Saving is preservation-first. The default path is a **surgical save**: a Rust patcher rewrites only the cells, formats, and structure you actually changed inside the original `.xlsx` package, so everything GridPath doesn't model — charts, pivot tables, conditional formatting, data validation, comments, external links, VBA, custom XML — comes out byte-identical. The patcher handles cell/format edits, sheet create/rename/delete, and row/column insert/delete (including chart series references, comment anchors, and cross-sheet formula shifts). When an edit falls outside what it can patch (e.g. restructuring around pivot tables or slicers), GridPath falls back to a full ExcelJS re-export — silently for plain files with nothing at risk, and for at-risk files it refuses to overwrite the original and instead offers a reduced-fidelity copy that names exactly what the copy will be missing ("2 charts, 14 comments…").

## Why it's different

Most "AI for Excel" tools either run in a web canvas (your data leaves your machine), require an API key from day one (everyone already has a Claude or ChatGPT sub), or apply edits irreversibly (one wrong prompt nukes the model). GridPath does:

- **Real local desktop, real .xlsx** — Tauri 2 + Rust shell, no upload, no SaaS lock-in. Open the file, edit it, save it, close. The file on disk is a normal `.xlsx` everyone else can open.
- **Subscription auth via the official CLIs** — Anthropic's `claude setup-token` and OpenAI's Codex CLI both mint OAuth tokens that GridPath replays against `/v1/messages` and `/responses`. No separate API budget required.
- **Reversible edit batches** — every agent action captures a pre-state snapshot of the cells it touches; "Reject" walks the stack in reverse. The "save a copy of the original" dance is unnecessary.
- **Live preview while pending review** — Univer's facade is fast enough that the user sees the agent's edits as they happen, but they're still in "pending" state until the user commits.
- **Parallel sessions per workbook** — open multiple tabs, run different agent sessions on different files concurrently. Each session has its own persisted history, token usage, and accept/reject batches.

## Features

| Area | What's there |
|---|---|
| Agent loop | Multi-turn tool use with Claude (`/v1/messages`) or Codex (`/responses`); cancellation, per-tool readback timeouts |
| Tool surface | Cell/range writes, formulas, formatting, dimensions, merges, freeze, sheets, web fetch |
| Grid | Univer.js — full Excel formula engine, A1 references, number formats |
| Round-trip | Surgical in-package save: untouched parts (charts / pivots / CF / validation / comments / VBA) stay byte-identical; structure edits (sheets, rows/columns) patched in place with reference shifting; gated ExcelJS fallback that never silently overwrites at-risk files |
| Diff & review | Per-batch accept/reject with exact pre-state restoration |
| Sessions | Per-workbook chat history, batch log, token usage, persisted to SQLite |
| Auth | Claude OAuth (subscription) or API key; ChatGPT subscription via Codex OAuth |
| Prompt caching | System prompt + tool schema cached via Anthropic's `cache_control` ephemeral breakpoint |
| Untitled drafts | New blank workbooks auto-snapshot to disk so they survive a restart |
| Updater | Tauri auto-updater wired to S3-hosted release feed |
| Distribution | Signed + notarized `.dmg` and `.msi` artifacts |

## Tech stack

- **Desktop shell:** Tauri 2 (Rust), SQLite via Diesel, `tauri-plugin-updater` for auto-update
- **Frontend:** React 18, Vite 5, TypeScript, Chakra UI, Univer.js (grid + formula engine), ExcelJS (load + fallback export; saves normally go through the Rust surgical patcher)
- **LLM providers:** Anthropic `/v1/messages` (streaming, tools) and OpenAI Codex `/responses` (streaming, tools)
- **Marketing site:** Static HTML/CSS served from `marketing/` (Vercel)

## Repository layout

```
gridpath/
├── src/                                  React frontend
│   ├── screens/SpreadsheetScreen/        Main product UI (grid, chat, diff review)
│   │   ├── components/                   UniverGrid, ChatPanel, TabBar, SessionSidebar, …
│   │   ├── agent/                        Tool-call interpretation, context capture, agent client
│   │   ├── state/                        Workspace reducer (tabs, batches, sessions)
│   │   └── …
│   └── packages/                         Shared design system + hooks
├── src-tauri/
│   ├── src/
│   │   ├── main.rs                       Binary entry, command registration, tray, DB wiring
│   │   └── engine/
│   │       ├── llm_providers/            claude.rs · openai_codex.rs (streaming + tools)
│   │       ├── spreadsheet_agent/        Agent loop, tool schemas, system prompt
│   │       └── workbook/                 .xlsx + untitled session I/O
│   ├── migrations/                       Diesel SQLite migrations
│   ├── icons/                            App icon set
│   └── tauri.conf.json                   Tauri / bundle / updater config
├── marketing/                            gridpath.dev (HTML + CSS, no framework)
└── scripts/
    ├── build-mac.sh                      Build + sign + notarize + S3 upload
    └── build-platform.sh                 Cross-platform thin wrapper
```

## Quick start

### Requirements

- **Node 20+** (Tauri 2 / Vite 5 baseline)
- **Rust** (stable) — install via [rustup](https://www.rust-lang.org/tools/install)
- Platform toolchains for Tauri — see [Tauri prerequisites](https://v2.tauri.app/start/prerequisites/)

### Run in dev

```bash
npm install
npm run tauri dev
```

### Build (local, unsigned)

```bash
npm install
npm run tauri build
```

### Build for macOS distribution (signed + notarized + uploaded to S3)

```bash
./scripts/build-mac.sh                # build + sign + notarize + upload
./scripts/build-mac.sh --no-upload    # local signed build only
```

Credentials are auto-loaded from `.env.build` (gitignored). See the script header for the full list of required env vars.

### Configure

On first launch you'll be asked which provider to use:

1. **Claude subscription** — paste an OAuth token from `claude setup-token`, OR
2. **Claude API key** — paste an `sk-ant-api03-*` key, OR
3. **ChatGPT subscription** — sign in via the Codex OAuth flow (browser tab opens automatically).

Credentials are stored locally in the SQLite app database — they don't leave your machine.

## Project status

GridPath is under active development. The agent's core loop is solid for income-statement-class models; complex multi-sheet builds (full 3-statement models, sensitivity tables across sheets) still occasionally trip on layout-tracking and are an active area of work.

If you're reading the source for the first time, the most useful entry points are:

- `src/screens/SpreadsheetScreen/SpreadsheetScreen.tsx` — main product UI, event handlers, accept/reject
- `src/screens/SpreadsheetScreen/components/UniverGrid.tsx` — grid + ExcelJS round-trip
- `src-tauri/src/engine/spreadsheet_agent/commands.rs` — agent loop, provider dispatch, tool result handling
- `src-tauri/src/engine/spreadsheet_agent/tools.rs` — tool schemas + the system prompt that drives behavior
- `src-tauri/src/engine/llm_providers/` — Claude + Codex provider implementations

## License

GridPath is source-available under the [Functional Source License, Version 1.1, with Apache 2.0 Future License](LICENSE) (FSL-1.1-Apache-2.0). In short:

- **You can** read, fork, modify, run, and redistribute the source for any non-competing use — including your own commercial use of the software.
- **You can't** ship the software (or a substantially similar fork) as a paid product or service that competes with GridPath.
- **In two years**, each release automatically re-licenses to Apache 2.0 — no restrictions.

If you'd like to use GridPath in a way the FSL doesn't permit, get in touch.
