# gridpath

**Let your agent edit Excel files without butchering them.**

An MCP server that gives Claude Code, Cursor, Codex or any MCP host a safe way to
edit the `.xlsx` files you already have:

- **Edits the real file.** No conversion into another format. Saves rewrite only
  the parts an edit touched; every other part of the workbook (charts, pivots,
  VBA, add-in data, custom XML) comes back byte-identical.
- **Real formulas.** A calculation engine runs in-process, so the agent reads
  evaluated values and writes formulas that work when the file opens in Excel.
- **Nothing lands until you say so.** Writes accumulate as pending batches with
  a readback of what changed; `save_workbook` is the only tool that touches disk.
- **Review before save.** Every write returns a `review_url`: a local page that
  shows the workbook with changed cells highlighted, before/after per cell, and
  Accept / Reject per batch and Save. Run with `--review-required` and only that
  page can save — the agent just hands you the link.
- **Sandboxed scripts.** `run_script` runs the agent's JavaScript against the
  workbook in a worker with a hard timeout, and lands the result as one batch.

## Install

```sh
npx gridpath mcp --allow ~/models
```

Files outside the `--allow` directories are refused. Default: the working directory.
Add `--review-required` if only a human should be able to write the file.

Pending changes live next to the workbook (`.gridpath/<file>.batches.json`) until
saved or rejected, so `gridpath review model.xlsx` can reopen the review page later.

### Claude Code

```sh
claude mcp add gridpath -- npx -y gridpath mcp --allow ~/models
```

### Cursor / Claude Desktop (`mcp.json`)

```json
{
  "mcpServers": {
    "gridpath": { "command": "npx", "args": ["-y", "gridpath", "mcp", "--allow", "/Users/you/models"] }
  }
}
```

### Codex (`~/.codex/config.toml`)

```toml
[mcp_servers.gridpath]
command = "npx"
args = ["-y", "gridpath", "mcp", "--allow", "/Users/you/models"]
```

Then: *"In model.xlsx, add an FY2029E column to the income statement."*

## Tools

Reads: `describe_workbook`, `find_rows`, `read_range`.
Writes (pending until saved): `set_cell`, `set_range`, `set_format`, `run_script`,
`clear_range`, `insert_rows`/`delete_rows`, `insert_columns`/`delete_columns`,
`create_sheet`/`rename_sheet`/`delete_sheet`, widths, heights, freeze panes,
hide/show, `define_name`, notes, merges.
Batches: `list_batches`, `reject_batch`, `save_workbook` (gated by `--review-required`).

Every tool takes `path`. Coordinates are A1 notation.

## How it works

`open` parses the `.xlsx` into an in-process calculation engine (a fork of
[IronCalc](https://github.com/ironcalc/IronCalc), compiled to WebAssembly).
Tools mutate that model and read back evaluated results. `save_workbook` diffs
the model against the file as loaded and patches the original bytes at the
zip-part level. Untouched parts are copied through; the engine never
regenerates the package.

## Status

v0. Not yet in v0: `fetch_web`/`edgar_lookup` sourcing tools, `copy_range`,
`stage_data`, per-cell reject in the review page, and the desktop app's own
agent loop as a tool. Benchmarks against other approaches are in the
[gridpath-evals](https://github.com/pixelsmasher13/gridpath-evals) repo.

## License

FSL-1.1-Apache-2.0. Each release becomes Apache 2.0 after two years.
