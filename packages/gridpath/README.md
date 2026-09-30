# gridpath

[![npm](https://img.shields.io/npm/v/gridpath?color=2f6feb&label=npm)](https://www.npmjs.com/package/gridpath)
[![license](https://img.shields.io/badge/license-FSL--1.1--Apache--2.0-blue)](https://github.com/pixelsmasher13/gridpath/blob/main/LICENSE)
[![node](https://img.shields.io/badge/node-%3E%3D20-3c873a)](https://nodejs.org)

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

## Install

### Claude Desktop (no terminal)

Download [`gridpath-0.1.7.mcpb`](https://github.com/pixelsmasher13/gridpath/releases/download/v0.1.7/gridpath-0.1.7.mcpb)
and open it. Claude Desktop shows an install dialog where you pick the folders
Claude may edit and whether saving requires review. Node is bundled; nothing else
to install. Claude Desktop shows its standard "not verified" warning for any
extension installed from a file; the source is [public](https://github.com/pixelsmasher13/gridpath/tree/main/packages/gridpath)
and nothing is written until you accept and save on the review page.

### Claude Code

```sh
claude mcp add gridpath -- npx -y gridpath mcp --review-required
```

### Cursor

[Add to Cursor](cursor://anysphere.cursor-deeplink/mcp/install?name=gridpath&config=eyJjb21tYW5kIjoibnB4IiwiYXJncyI6WyIteSIsImdyaWRwYXRoIiwibWNwIiwiLS1yZXZpZXctcmVxdWlyZWQiXX0=)
(one click), or add to `mcp.json`:

```json
{ "mcpServers": { "gridpath": { "command": "npx", "args": ["-y", "gridpath", "mcp", "--review-required"] } } }
```

### Codex (`~/.codex/config.toml`)

```toml
[mcp_servers.gridpath]
command = "npx"
args = ["-y", "gridpath", "mcp", "--review-required"]
```

### Gemini CLI, Windsurf, VS Code, anything else

Same `command`/`args` as the Cursor JSON in that host's MCP settings. The server
speaks stdio.

### Options

| Flag | Meaning |
|---|---|
| `--allow <dir> [<dir>...]` | Limit the server to these folders. Default: your home directory. |
| `--review-required` | The agent cannot save. `save_workbook` returns the review link; you accept and save there. Also `GRIDPATH_REVIEW_REQUIRED=true`. |
| `--port <n>` | Port for the local review server (default: random). |
| `gridpath review <file>` | Reopen the review page for a workbook's pending batches later. |

Pending changes live next to the workbook (`.gridpath/<file>.batches.json`) until
saved or rejected, so nothing is lost if the agent's session ends.

## Supported files

`.xlsx` and `.xlsm`. Verified on real analyst models: a 1.4 MB, 12-sheet,
55,000-formula workbook (76 package parts, 75 byte-identical after an edit) and a
13-sheet `.xlsm` bank model whose `vbaProject.bin` came back byte-identical.

## Tools

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

## What the agent said

Claude, after editing a 1,400-row Canalyst model (2,500+ named ranges, hidden
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

## How it compares

| | gridpath | File-library MCPs (openpyxl etc.) | Live-Excel MCPs (COM) | Browser spreadsheet SDKs |
|---|---|---|---|---|
| Edits the file you have | yes, patched in place | rewrites the whole file | yes, via Excel | converts to its own format |
| Formulas evaluated | yes, in process | no | yes, needs Excel open | yes, in a hosted app |
| Review before write | yes | no | no | in the vendor's UI |
| Needs installed | Node | Python | Windows + Excel | a web app + account |
| Macros, pivots, add-in data survive | byte-identical | often dropped | yes | not guaranteed |

Benchmarks and methodology, including runs where a plain coding agent does as
well as we do, are in the [gridpath-evals](https://github.com/pixelsmasher13/gridpath-evals) repo.

## How it works

`open` parses the `.xlsx` into the in-process calculation engine. Tools mutate
that model and read back evaluated results. `save_workbook` diffs the model
against the file as loaded and patches the original bytes at the zip-part
level. Untouched parts are copied through; nothing regenerates the package.

## Status

v0.1. Not yet: `fetch_web` / `edgar_lookup` sourcing tools, `stage_data`,
charts and pivots as *edits* (they are
preserved, not authored), and the desktop app's own agent loop as a tool.

GridPath is also a [desktop app](https://gridpath.dev) with the same engine
and review and its own agent built in.

## License

FSL-1.1-Apache-2.0. Each release becomes Apache 2.0 after two years.
