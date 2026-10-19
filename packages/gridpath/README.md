# gridpath

[![npm](https://img.shields.io/npm/v/gridpath?color=2f6feb&label=npm)](https://www.npmjs.com/package/gridpath)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/pixelsmasher13/gridpath/blob/main/packages/gridpath/LICENSE)
[![node](https://img.shields.io/badge/node-%3E%3D20-3c873a)](https://nodejs.org)

**Real Excel files, in your web app and in your agent.** Open the `.xlsx`
your users already have, keep every formula live at full size, let AI edit
it with review, and save back the same file.

| | In your web app | In your coding agent |
|---|---|---|
| Install | `npm install gridpath` | `claude mcp add gridpath -- npx -y gridpath mcp` |
| What you get | An embeddable spreadsheet: a canvas grid over a Rust/WebAssembly engine in a Web Worker, in-place `.xlsx` saves, and agent tools for your own model | An MCP server for Claude Code, Claude Desktop, Cursor, Codex: the agent edits files on disk and you review every change before it is saved |
| Docs | [In your web app](#in-your-web-app) | [In your coding agent](#in-your-coding-agent-mcp) |

Both run the same engine (a fork of [IronCalc](https://github.com/ironcalc/IronCalc)):
real formulas evaluated in process, and saves that patch the original file so
charts, pivots, VBA macros, add-in data and custom XML come back byte-identical.

## In your web app

### Install

```sh
npm install gridpath react react-dom
```

React and React DOM are peer dependencies. Your app does not have to use
React: `createSpreadsheet` mounts the grid into any element.

### Use

```ts
import { createSpreadsheet } from "gridpath";

const el = document.getElementById("sheet")!; // give it a height
const sheet = await createSpreadsheet(el, {
  file: await file.arrayBuffer(), // omit for a blank workbook
  onEdit: () => console.log("edited"),
});

sheet.setCell("Sheet1", "B2", "=SUM(A1:A10)");
console.log(await sheet.getCell("Sheet1", "B2")); // { value: 55, formula: "=SUM(A1:A10)" }

const xlsx = await sheet.save(); // Uint8Array; or saveBlob() for a download
```

React:

```tsx
import { Spreadsheet } from "gridpath";

<div style={{ height: 600 }}>
  <Spreadsheet file={bytes} onReady={(sheet) => (sheetRef.current = sheet)} />
</div>;
```

The workbook loads once per mount. To open another file, remount with a new
`key`.

#### Options

| Option | |
|---|---|
| `file` | `.xlsx` bytes (`ArrayBuffer` or `Uint8Array`). Without it the workbook starts blank. |
| `sheets` | Sheet names for a blank workbook (default `["Sheet1"]`). |
| `fileName` | Name used in messages. |
| `onEdit` | Called after each user edit. |
| `onProgress` | Loading stages. |
| `engine.worker` | `false` runs the engine on the page instead of a worker. |
| `engine.workerUrl`, `engine.wasmUrl` | Where to load the worker script and `.wasm` from, when your bundler does not serve them next to the package. |

#### Methods

| Method | |
|---|---|
| `getSheetNames()` | Sheet names in order. |
| `getCell(sheet, "A1")` | `{ value, display?, formula }` or `null`; waits for pending recalculation. |
| `setCell(sheet, "A1", value)` | Strings starting with `=` are formulas. |
| `setCells(sheet, [{ row, col, value }])` | Many cells in one step (0-based). |
| `whenCalculated()` | Resolves when pending writes are recalculated. |
| `select(sheet, "A1")`, `focus()` | Move the selection; give the grid keyboard focus. |
| `save()`, `saveBlob()` | The workbook as `.xlsx`. |
| `destroy()` | Unmount and stop the engine (`createSpreadsheet` only). |

### AI: your model, your key

`sheet.ai` is GridPath's agent tool set (the same tools and result shapes as
the [MCP server](#in-your-coding-agent-mcp)) acting on the grid the user is
looking at. Give the tools to your model, run its tool calls with
`sheet.ai.run()`, and let the user review: every write lands at once,
highlighted, as a pending batch they accept or reject.

GridPath never sees a key. The model call happens on your server; the tool
calls run in the browser, where the workbook is.

**Server** (any Node route; this one is the demo's `server/agent.ts`). The
system prompt and tools are fixed here; taking them from the request would
make the route a general proxy for your key. Add your own auth and rate
limits.

```ts
import Anthropic from "@anthropic-ai/sdk";
import { AGENT_INSTRUCTIONS, agentTools } from "gridpath/agent";

const client = new Anthropic(); // ANTHROPIC_API_KEY
const tools = agentTools() as Anthropic.Beta.BetaTool[];

export async function agentTurn(messages: Anthropic.Beta.BetaMessageParam[]) {
  const response = await client.beta.messages.create({
    model: "claude-opus-5-5",
    max_tokens: 16000,
    output_config: { effort: "high" },
    system: AGENT_INSTRUCTIONS,
    tools,
    cache_control: { type: "ephemeral" }, // system + tools are identical every call
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default", // a declined request re-runs on a fallback model
    messages,
  });
  return { content: response.content, stop_reason: response.stop_reason };
}
```

**Browser** (the demo's `src/agentLoop.ts`, shortened):

```ts
const messages = [{ role: "user", content: [
  { type: "text", text: `The open workbook:\n${await sheet.ai.context()}` },
  { type: "text", text: prompt },
] }];
for (;;) {
  const out = await (await fetch("/api/agent", { method: "POST", body: JSON.stringify({ messages }) })).json();
  messages.push({ role: "assistant", content: out.content }); // unchanged, thinking blocks included
  const calls = out.content.filter((b) => b.type === "tool_use");
  if (out.stop_reason !== "tool_use" || calls.length === 0) break;
  const results = [];
  for (const c of calls) {
    const r = await sheet.ai.run(c.name, c.input);
    results.push({ type: "tool_result", tool_use_id: c.id, content: JSON.stringify(r), ...(r?.ok === false ? { is_error: true } : {}) });
  }
  messages.push({ role: "user", content: results }); // all results in one message
}
```

Already have an agent framework? Use `sheet.ai.tools()` (Anthropic format)
or `sheet.ai.tools("openai")` and route tool calls to `sheet.ai.run()`.

| `sheet.ai` | |
|---|---|
| `tools(format?)` | Tool definitions: `"anthropic"` (default) or `"openai"`. Also exported from `gridpath/agent` as `agentTools()`. |
| `instructions()` | How to use them, for the system prompt (`AGENT_INSTRUCTIONS`). |
| `context()` | The workbook overview for the first turn. |
| `run(name, input)` | Run one tool call; resolves with JSON for the model. Never throws. |
| `list()`, `onChange(fn)` | The batches, pending or accepted. |
| `accept(id?)` | Keep one batch, or every pending one. |
| `reject(id?)` | Undo a batch **and every batch after it**, newest first, so rows and columns moved by later batches stay right. Resolves with the ids undone. |

Included: describe_workbook, find_rows, query_rows, read_range, set_cell,
set_range, copy_range, set_format, clear_range, run_script (sandboxed in a
worker), row/column/sheet structure, freeze, hide, merge, notes, sizes,
define_name. Not included: tools that need the desktop app or the network
(fetch_web, edgar_lookup, reference workbooks).

`save()` writes every applied batch, pending or accepted; reject what you
don't want first.

### How saving works

Every save replays the session's changes (cells, formats, sizes, rows and
columns, sheets, filters) onto the bytes the workbook was opened from. Parts
the session did not touch are copied through unchanged. New formulas are
written without a cached result and the file is marked for recalculation on
open, which Excel honours; a reader that never calculates (pandas, openpyxl
with `data_only=True`) sees those cells empty until the file has been opened
and saved in a spreadsheet app.

### Compared with other embeddable spreadsheets

Against the spreadsheet components a web app would otherwise embed, on real
models in Chromium ([workbooks, method and full results](https://github.com/pixelsmasher13/gridpath/tree/main/eval/embed-bench#results-2026-10-19)):

| | GridPath | SpreadJS 19.2 | Handsontable 18.1 + HyperFormula 3.5 + ExcelJS 4.4 |
|---|---|---|---|
| 1.4 MB analyst model: open, fully recalculated | 0.6 s, 0.8 s | 1.5 s, 3.6 s | 1.9 s |
| … agrees with Excel's results | 95.5% | 94.0% | 56.9% |
| … lost on save | nothing (75 of 76 parts byte-identical) | custom XML, printer settings | custom XML, printer settings, comments, 3,335 of 3,418 names |
| Data-intensive model over 10 MB (2.3M cells): open, fully recalculated | 4.2 s, 10.9 s | tab crashes (lazy mode: 6.5 s, 3.7 min) | 10.8 min |
| … edit → dependent updated | 2.7 s | — (lazy mode crashes on the first edit) | 64 s |

### Bundlers

The package loads its worker and engine with
`new URL("…", import.meta.url)`, which Next.js, Vite, webpack 5 and Rollup
bundle on their own.

- **Next.js** (tested on 16 with React 19, Turbopack and webpack, `next dev`
  and production builds): no configuration. Render the spreadsheet from a
  client component; the page around it can stay a server component, and
  `gridpath/agent` works in a route handler:

  ```tsx
  // app/Sheet.tsx
  "use client";
  import { Spreadsheet } from "gridpath";

  export function Sheet() {
    return <Spreadsheet onReady={(sheet) => { /* keep `sheet` */ }} />;
  }
  ```

  ```ts
  // app/api/agent/route.ts
  import { AGENT_INSTRUCTIONS, agentTools } from "gridpath/agent";
  ```

  A complete app is in [`examples/nextjs-demo`](https://github.com/pixelsmasher13/gridpath/tree/main/examples/nextjs-demo).
- **Vite**: production builds work as is. For the dev server, exclude the
  package from dependency pre-bundling:

  ```ts
  // vite.config.ts
  export default defineConfig({ optimizeDeps: { exclude: ["gridpath"] } });
  ```

- **webpack 5** (outside Next.js): no configuration; `new URL` workers and
  assets are built in.
- **Anything else**: copy `engine.worker.js`, its chunks and
  `gridpath_engine_bg.wasm` from `dist/` to your static files and pass
  `engine: { workerUrl, wasmUrl }`.

Size in the browser: about 84 KB of JavaScript and a 1.2 MB engine, gzipped
(0.87 MB with brotli). The engine downloads once and is cached by the browser.

### Build from source

```sh
crates/engine-wasm/build.sh   # needs rustup target wasm32-unknown-unknown and wasm-bindgen-cli
cd packages/gridpath && npm install && npm pack
```

`examples/embed-demo` installs the packed tarball the way an app would.

## In your coding agent (MCP)

**Let your agent edit Excel files. Review before anything is saved.**

An MCP server for Claude Code, Claude Desktop, Cursor, Codex or any MCP host
that edits the `.xlsx` / `.xlsm` files you already have, with a review step:

- **Review before save.** Every write returns a `review_url`: a local page with
  the workbook, changed cells highlighted, before/after per cell, Accept /
  Reject per batch or per cell, Save or Save as… (a copy). With `--review-required`, only
  that page can write the file; the agent just hands you the link.
- **Leaves your Excel file intact.** No re-export through a library, no
  conversion into another format. Saves patch the original bytes: only the
  parts an edit touched are rewritten; charts, pivots, **VBA macros**, add-in
  data and custom XML come back byte-identical.
- **Real formulas, in process.** A calculation engine (a fork of
  [IronCalc](https://github.com/ironcalc/IronCalc), compiled to WebAssembly)
  runs inside the server. The agent reads evaluated values and writes formulas
  that work when the file opens in Excel; inserts and deletes shift references.
- **Runs anywhere Node runs.** No Python, no openpyxl, no LibreOffice, no Excel
  install. macOS, Windows, Linux.
- **Sandboxed scripts.** `run_script` runs the agent's JavaScript against the
  workbook in a worker with a hard timeout and lands the result as one
  reviewable batch.

![GridPath review page: workbook with a changed cell highlighted and a panel of pending change batches](https://gridpath.dev/mcp-review.png)

### Install

#### Claude Desktop (no terminal)

Download [`gridpath-0.2.0.mcpb`](https://github.com/pixelsmasher13/gridpath/releases/download/v0.2.0/gridpath-0.2.0.mcpb)
and open it. Claude Desktop shows an install dialog where you pick the folders
Claude may edit and whether saving requires review. Node is bundled; nothing else
to install. Claude Desktop shows its standard "not verified" warning for any
extension installed from a file; the source is [public](https://github.com/pixelsmasher13/gridpath/tree/main/packages/gridpath)
and nothing is written until you accept and save on the review page.

#### Claude Code

```sh
claude mcp add gridpath -- npx -y gridpath mcp --review-required
```

#### Cursor

[Add to Cursor](cursor://anysphere.cursor-deeplink/mcp/install?name=gridpath&config=eyJjb21tYW5kIjoibnB4IiwiYXJncyI6WyIteSIsImdyaWRwYXRoIiwibWNwIiwiLS1yZXZpZXctcmVxdWlyZWQiXX0=)
(one click), or add to `mcp.json`:

```json
{ "mcpServers": { "gridpath": { "command": "npx", "args": ["-y", "gridpath", "mcp", "--review-required"] } } }
```

#### Codex (`~/.codex/config.toml`)

```toml
[mcp_servers.gridpath]
command = "npx"
args = ["-y", "gridpath", "mcp", "--review-required"]
```

#### Gemini CLI, Windsurf, VS Code, anything else

Same `command`/`args` as the Cursor JSON in that host's MCP settings. The server
speaks stdio.

#### Options

| Flag | Meaning |
|---|---|
| `--allow <dir> [<dir>...]` | Limit the server to these folders. Default: your home directory. |
| `--review-required` | The agent cannot save. `save_workbook` returns the review link; you accept and save there. Also `GRIDPATH_REVIEW_REQUIRED=true`. |
| `--port <n>` | Port for the local review server (default: random). |
| `gridpath review <file>` | Reopen the review page for a workbook's pending batches later. |

Pending changes live next to the workbook (`.gridpath/<file>.batches.json`) until
saved or rejected, so nothing is lost if the agent's session ends.

### Tools

Every tool takes `path`. Cells and ranges are A1 notation.

| Tool | What it does |
|---|---|
| `describe_workbook` `(sheet?)` | Structural map: sheets, used ranges, header rows, sections, named ranges. Call first. |
| `find_rows` `(query, sheet?, max_results?)` | Find rows by label across sheets, with an evaluated sample. |
| `read_range` `(sheet, range)` | Evaluated values, formulas and display text for a range (≤ 500 cells). |
| `set_cell` `(sheet, cell, value \| formula)` | Write one cell. |
| `set_range` `(sheet, top_left, values[][])` | Write a block; strings starting with `=` are formulas. |
| `copy_range` `(sheet, source, dest_sheet, dest, mode)` | Excel-style copy with reference shifting; `all` / `values` / `formats`. |
| `run_script` `(script)` | Sandboxed JS with `sheet(name).get/set/values/setValues/format/clear`; one batch. |
| `set_format` `(sheet, range \| operations[], format)` | Number formats, bold/italic, colors, alignment, wrap. |
| `clear_range` `(sheet, range)` | Clear values and formulas, keep formatting. |
| `insert_rows` / `delete_rows` / `insert_columns` / `delete_columns` | Structural edits; references, CF and validation ranges follow. |
| `create_sheet` / `rename_sheet` / `delete_sheet` | Sheet ops. |
| `set_column_width` / `set_row_height` / `freeze_panes` / `unfreeze_panes` / `hide_rows` / `show_rows` / `hide_columns` / `show_columns` | Layout. |
| `merge_cells` / `unmerge_cells` / `set_note` / `delete_note` / `define_name` | Merges, notes, named ranges. |
| `list_batches` | Pending batches and the review link. |
| `reject_batch` `(batch_id)` | Drop a pending batch; the rest replay onto a fresh model. |
| `save_workbook` `(as?)` | Patch pending changes into the file, or with `as` write a patched copy and leave the original untouched. The only tool that writes. In-place saves are gated by `--review-required`. |

Writes never touch the file. They accumulate as pending batches and every write
returns a readback of what landed (values, formulas, a `row_map` of labels) plus
the `review_url`.

### What the agent said

Claude, after editing a 1,400-row sell-side model (2,500+ named ranges, hidden
sheets) through GridPath in Cursor, asked how it compared with its built-in
approach:

> **It recalculates live.** After I changed row 58, I could read back the new
> growth rates right away (28.0%, 28.0%, 27.9%) and confirm the change worked.
> My usual approach, a Python library called openpyxl, doesn't calculate
> formulas. I'd have to run the file through LibreOffice to get values, and
> that can differ from Excel on some functions.
>
> **It finds things fast.** Its row search and workbook map located "Product
> revenue" and the Assumptions structure in a few calls. Otherwise I'd be
> scanning a 1,400-row sheet cell by cell.
>
> **It leaves the rest of the file alone.** It rewrites only the cells that
> changed. When openpyxl opens and re-saves a heavy model like this one, it can
> quietly drop or damage charts, data validation, pivot caches or some
> formatting. That's the biggest risk with my built-in approach.

### How it compares

| | gridpath | File-library MCPs (openpyxl etc.) | Live-Excel MCPs (COM) | Browser spreadsheet SDKs |
|---|---|---|---|---|
| Edits the file you have | yes, patched in place | rewrites the whole file | yes, via Excel | converts to its own format |
| Formulas evaluated | yes, in process | no | yes, needs Excel open | yes, in a hosted app |
| Review before write | yes | no | no | in the vendor's UI |
| Needs installed | Node | Python | Windows + Excel | a web app + account |
| Macros, pivots, add-in data survive | byte-identical | often dropped | yes | not guaranteed |

Benchmarks and methodology, including runs where a plain coding agent does as
well as we do, are in the [gridpath-evals](https://github.com/pixelsmasher13/gridpath-evals) repo.

## Supported files

`.xlsx` and `.xlsm`. Verified on real analyst models: a 1.4 MB, 12-sheet,
55,000-formula workbook (76 package parts, 75 byte-identical after an edit) and a
13-sheet `.xlsm` bank model whose `vbaProject.bin` came back byte-identical.

## How it works

The engine parses the `.xlsx` into an in-process model; edits mutate it and
read back evaluated results, recalculating only what an edit affects. Saving
diffs the session's changes against the file as loaded and patches the
original bytes at the zip-part level: untouched parts are copied through,
nothing regenerates the package. In the browser the engine runs in a Web
Worker and the grid draws only the visible window, so large models stay light.

## Headless (Node)

`gridpath/headless` is the engine without the grid: open a workbook, apply
the same tool calls, read results, and save, from a server or a script.

```ts
import { WorkbookSession } from "gridpath/headless";
```

## Status

v0.2. The browser spreadsheet is new in this release (previously the package
was the MCP server only; its Node API moved from `gridpath` to
`gridpath/headless`). Not yet: charts drawn in the grid (they are preserved,
not rendered), data-validation dropdowns, real-time collaboration; for agents,
`fetch_web` / `edgar_lookup` sourcing tools and charts or pivots as edits.

GridPath is also a [desktop app](https://gridpath.dev) with the same engine
and review and its own agent built in.

## License

Apache-2.0 ([LICENSE](LICENSE), [NOTICE](NOTICE)); third-party software in the
package is listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md). Releases
before 0.2.0 were FSL-1.1-Apache-2.0. "GridPath" is a trademark of Heelix
Technologies Inc.
