# GridPath

A desktop AI agent for Excel that edits the workbooks you already have — and gives them back intact.

Open a `.xlsx`, type what you want, review the diff, accept it. Your file never leaves your machine, and everything the agent didn't touch comes back byte-for-byte identical. Works with the Claude or ChatGPT subscription you're already paying for.

## What it does

- **Edits real workbooks in place** — formulas, formatting, whole sections, on files someone else built. Updating a 663-formula quarterly model with a new quarter's actuals takes about 4 minutes ([benchmark](https://github.com/pixelsmasher13/gridpath-evals)).
- **Every change is reversible** — each batch of edits lands live but stays pending until you accept it. Reject rewinds cell-by-cell to the exact prior values, formulas, number formats and fills. ⌘Z after that, and a `.bak` after saving.
- **Your file comes back whole** — a Rust patcher rewrites only the cells and structure that changed inside the original package, so charts, pivot tables, conditional formatting, data validation, comments, external links, VBA and plugin data stay untouched. Not "usually" — by construction.
- **Several workbooks at once** — each tab runs its own agent session with its own history and pending batches, in parallel. Attach up to five more files as read-only references the agent can pull ranges from.
- **Bring your own model** — Claude Pro/Max or ChatGPT Plus/Pro via the official OAuth flows, or API keys if you prefer. Switch provider and effort level per session.
- **Grounded in real data** — the agent can look up SEC filings and fetch source documents directly, so figures come from the filing rather than from memory.
- **Runs on your machine** — Tauri + Rust desktop app, SQLite on disk. Only your prompt and the cells the agent reads or writes go to the model; the workbook itself never uploads.

## How it compares

| | GridPath | Excel Copilot | Claude for Excel | ChatGPT upload | CLI coding agent |
|---|---|---|---|---|---|
| Edits your existing `.xlsx` in place | ✅ | ✅ | ✅ | ❌ | ✅ |
| Workbook stays on your machine | ✅ | ❌ | ❌ | ❌ | ✅ |
| Preserves charts, pivots, plugin data | ✅ | ✅ | ✅ | ❌ | ❌ |
| Accept/reject diff before changes land | ✅ | partial | partial | ❌ | ❌ |
| Choose and switch the model | ✅ | ❌ | ❌ | ❌ | ✅ |
| Parallel sessions across workbooks | ✅ | ❌ | ❌ | ❌ | partial |
| Source available | ✅ | ❌ | ❌ | ❌ | ❌ |
| Included with a subscription you have | ✅ | ❌ | ✅ | ✅ | ✅ |

*Best-effort comparison as of August 2026. The "preserves" and "diff" rows for coding agents are measured, not inferred — see the [benchmark repo](https://github.com/pixelsmasher13/gridpath-evals), where a headless coding agent silently dropped parts of a workbook in 2 of 5 edit runs.*

## Quick start

```bash
npm install
npm run tauri dev
```

Requires **Node 20+** and **Rust** (stable, via [rustup](https://www.rust-lang.org/tools/install)) plus your platform's [Tauri prerequisites](https://v2.tauri.app/start/prerequisites/). `npm run tauri build` produces a local unsigned build.

On first launch, pick a provider: a Claude OAuth token from `claude setup-token`, a `sk-ant-api03-*` API key, or the ChatGPT sign-in flow. Credentials live in the local SQLite database and don't leave your machine.

## Built with

Tauri 2 and a Rust core, React + TypeScript frontend, a vendored [IronCalc](vendor/ironcalc) fork (MIT/Apache-2.0) for the formula engine, and SQLite for local session storage. The surgical `.xlsx` patcher is in `src-tauri/src/engine/workbook/xlsx_patch/`, the agent loop and tool schemas in `src-tauri/src/engine/spreadsheet_agent/`, and the benchmark harness in `eval/`.

## Project status

Under active development. Editing existing workbooks — the case above — is the strongest path and the one under continuous benchmark. Building large multi-sheet models from scratch works but is improving; the [benchmark repo](https://github.com/pixelsmasher13/gridpath-evals) publishes those results too, including the ones we lose.

## License

Source-available under the [Functional Source License, Version 1.1, with Apache 2.0 Future License](LICENSE) (FSL-1.1-Apache-2.0):

- **You can** read, fork, modify, run and redistribute the source for any non-competing use, including commercially.
- **You can't** ship it, or a substantially similar fork, as a paid product competing with GridPath.
- **In two years** every release re-licenses to Apache 2.0, no restrictions.

If you want to use GridPath in a way the FSL doesn't permit, get in touch.
