use serde_json::{json, Value};

/// Tool schemas advertised to Claude on every spreadsheet-agent turn.
///
/// Coordinates are **Excel A1 notation only** (cell="A1", top_left="B17",
/// "A1:C10" etc.). Earlier versions of these tools used 0-indexed numeric
/// (row, col) args, which caused off-by-one mistakes in formula references
/// because Claude wrote `=B17` referencing what it *placed at* B17 in its
/// head, while the data actually landed at B18 due to 0-indexing. A1
/// notation removes the dual coordinate system entirely — the agent has
/// seen Excel formulas in training, this is the format it's most reliable in.
pub fn agent_tools() -> Vec<Value> {
    vec![
        json!({
            "name": "set_cell",
            "description": "Set the value or formula of a single cell. `cell` is an Excel A1-style address (case-insensitive, e.g. \"A1\", \"B17\", \"AA42\"). For formulas, pass them in `formula` starting with '=' (e.g. \"=SUM(A1:A10)\") — do NOT put formulas in `value`.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string", "description": "Sheet name (case-sensitive)." },
                    "cell":  { "type": "string", "description": "A1 address. Examples: \"A1\", \"B17\", \"AA42\"." },
                    "value": { "description": "Literal value (string, number, boolean, or null). Leave null when using `formula`." },
                    "formula": {
                        "type": ["string", "null"],
                        "description": "Excel-style formula starting with '='. References inside use A1 notation. Example: \"=SUM(A1:A10)\".",
                    },
                },
                "required": ["sheet", "cell"]
            }
        }),
        json!({
            "name": "set_range",
            "description": "Set a rectangular block of values starting at `top_left` (A1 address). `values` is a 2D array — outer rows, inner cells. Strings starting with '=' inside `values` are treated as formulas; formula references must use A1 notation pointing at the FINAL placement of cells (e.g. if `top_left` is \"A9\" and the value at row index 2 / col 0 should reference the data row just above your block at A8, write \"=A8\" — not \"=A1\" or \"=A11\"). **Off-by-one guard: if formulas will reference rows created by THIS call, split the work — write labels + hardcoded values first, read the `row_map` in this call's result, then write formulas in a second set_range against those verified row numbers.** Composing formulas in the same call that lays out their target rows is the top source of off-by-one references. Use this for bulk writes; prefer it over many set_cell calls. **`null` and empty string `\"\"` inside `values` are treated as PRESERVE — the cell is left untouched, NOT cleared.** This means you can ship `[[\"Gross Profit\", \"\", \"\", \"\"]]` to set only the label in column A without wiping the data in B–D. To actually clear cells, use `clear_range`. **Writing to a sheet name that doesn't exist CREATES that sheet automatically** (near-miss spellings of an existing sheet are rejected as probable typos) — so organize multi-section builds across purposefully named sheets (e.g. Data / Assumptions / Model / Notes) instead of piling everything onto one long sheet; a new tab costs nothing.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "top_left": { "type": "string", "description": "A1 address of the top-left cell, e.g. \"A9\"." },
                    "values": {
                        "type": "array",
                        "description": "2D array — outer = rows, inner = cells. Strings beginning with '=' are formulas.",
                        "items": { "type": "array", "items": {} }
                    }
                },
                "required": ["sheet", "top_left", "values"]
            }
        }),
        json!({
            "name": "copy_range",
            "description": "Copy a rectangular range to another location with Excel copy/paste semantics, executed against the LIVE grid: relative formula references shift by the paste offset exactly like Excel ($ anchors are respected — `$B$5` stays put, `B5` shifts), and cell formatting (number format, font, fill, alignment) travels with the cells. STRONGLY PREFER this over reading a range and re-typing shifted formulas yourself — it cannot make off-by-one mistakes and costs a fraction of the tokens. Ideal for: extending a projection by one more period (copy the last full year column block to the next column), duplicating a section as a template, or seeding a new sheet from an existing block. Blank source cells leave the destination untouched (they do NOT clear it). Result reports the evaluated destination values so you can verify. Limited to 5000 cells per call.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string", "description": "Source sheet name (case-sensitive)." },
                    "source": { "type": "string", "description": "A1 range to copy, e.g. \"BT5:BT120\" or a single cell \"C7\"." },
                    "dest": { "type": "string", "description": "A1 address of the destination TOP-LEFT cell, e.g. \"BU5\". The copy lands in a rectangle of the same size starting here." },
                    "dest_sheet": { "type": ["string", "null"], "description": "Destination sheet if different from `sheet`. Formula references shift the same way; sheet-qualified refs keep pointing at their original sheets." },
                    "mode": { "type": "string", "enum": ["all", "values", "formats"], "description": "\"all\" (default): shifted formulas + values + formats. \"values\": paste evaluated values only — use to freeze a calculation as hardcodes. \"formats\": formatting only, contents untouched." }
                },
                "required": ["sheet", "source", "dest"]
            }
        }),
        json!({
            "name": "run_script",
            "description": "Execute a short JavaScript program against the workbook and land ALL its writes as one reviewable batch. STRONGLY PREFER this over long set_range payloads or many tool turns whenever the edit is loop-shaped: per-row conditional transforms (\"for every row where margin < 0 …\"), repeating formula patterns across many rows/columns, address arithmetic, or edits that depend on reading existing cells first. One run_script call replaces dozens of tool turns and cannot drift out of alignment mid-payload, because it computes addresses programmatically.\n\nAPI (plain synchronous JavaScript, strict mode — no async/await, no imports, no network, no DOM):\n  sheets() -> array of sheet names\n  sheet(name) -> handle for one sheet (name must match exactly; the name arg may be omitted ONLY if the workbook has a single sheet). Handle methods:\n    .get(\"B7\") -> {value, formula} — evaluated value as of BEFORE this script, overlaid with your own literal writes\n    .values(\"A2:C50\") -> 2D array of values (rows × cols)\n    .set(\"B7\", x) — write a value; strings starting with '=' are formulas; null clears the cell's content\n    .setValues(\"A9\", [[...],[...]]) — bulk write with set_range semantics: null / \"\" preserve the existing cell\n    .format(\"A1:F1\", {bold: true, number_format: \"0.0%\", background_color: \"#1F4E79\", font_color: \"#FFFFFF\"}) — same keys as set_format\n    .clear(\"B2:D10\") — empty values + formulas, formatting preserved\n    .usedRange() -> \"A1:Q48\" or null\n  log(...) — up to 100 entries, returned to you in the result as `script_logs` for debugging\n\nRules: formulas you write are NOT evaluated during the script — .get on a cell you just wrote a formula into returns {value: null, formula}; the automatic readback after the script (cells + row_map) carries the evaluated results, so verify there. Follow the same conventions as set_range: write raw numbers (never pre-formatted strings), write formulas for derived values, compose against row_map. Limits: 5-second CPU cap, 20,000 touched-cell cap. `sheet(\"NewName\")` on a name that doesn't exist AUTO-CREATES that sheet when the script's writes land — use it to organize builds across purposeful tabs. Other structural ops (insert/delete rows or columns, sheet delete/rename, merges, freeze) are NOT in the script API — use the dedicated tools for those. The script's writes flow into the same pending batch as every other tool: the user reviews and can reject them.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "script": {
                        "type": "string",
                        "description": "JavaScript source, executed as a function body. Example: const s = sheet(\"Model\"); const rows = s.values(\"A2:C50\"); rows.forEach((r, i) => { if (r[2] !== null && r[2] < 0) s.set(\"D\" + (i + 2), \"=C\" + (i + 2) + \"*-1\"); });"
                    }
                },
                "required": ["script"]
            }
        }),
        json!({
            "name": "set_format",
            "description": "Apply formatting to one OR MANY cell ranges in a single call. STRONGLY PREFER the bulk `operations` form when you have more than one format to apply — it's one tool call instead of N, dramatically faster. Each operation's `format` is a partial object; only specified properties are applied, others are left as-is. Use this for: bold headers, currency / percent / number formatting on data columns, italic notes, alignment, font color/size/family, **text wrap via `wrap_text`** (long methodology notes / descriptions in a fixed-width column — wrap beats restructuring the layout around overflow), and **cell fill via `background_color`** (e.g. dark-blue header bars, yellow assumption highlights — see the financial-model color conventions). v1 limitation: cell borders are not yet supported.\n\nTwo equivalent shapes:\n  • Single range: { sheet, range, format }\n  • Bulk:         { sheet, operations: [ { range, format }, ... ] }",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "range": { "type": "string", "description": "A1 range (only when using single-range form). Examples: \"A1\", \"B2:B10\", \"A1:F1\"." },
                    "format": {
                        "type": "object",
                        "description": "Partial format (only when using single-range form).",
                        "properties": {
                            "bold":          { "type": "boolean" },
                            "italic":        { "type": "boolean" },
                            "underline":     { "type": "boolean" },
                            "strike":        { "type": "boolean", "description": "Strikethrough." },
                            "font_color":    { "type": "string", "description": "CSS color, e.g. \"#000000\" or \"red\"." },
                            "background_color": { "type": ["string", "null"], "description": "Cell fill color, e.g. \"#1F4E79\" for a dark-blue header bar or \"#FFFF00\" for a yellow assumption highlight. Pass null to clear. **Always pair light/white `font_color` with a contrasting `background_color` — white-on-white reads as invisible text.**" },
                            "font_size":     { "type": "number" },
                            "font_family":   { "type": "string", "description": "Font family name, e.g. \"Calibri\", \"Arial\", \"Aptos Narrow\"." },
                            "horizontal_align": { "type": "string", "enum": ["left", "center", "right"] },
                            "vertical_align":   { "type": "string", "enum": ["top", "middle", "bottom"] },
                            "wrap_text":     { "type": "boolean", "description": "Wrap long text within the cell instead of overflowing. Pair with a set_column_width + vertical_align \"top\" for description/notes columns." },
                            "number_format": {
                                "type": "string",
                                "description": "Excel-style number format. Common: \"$#,##0.00\" currency, \"0.0%\" or \"0.00%\" percent, \"#,##0\" integers with commas, \"#,##0.00\" decimals, \"@\" plain text."
                            }
                        }
                    },
                    "operations": {
                        "type": "array",
                        "description": "Bulk form — array of {range, format} pairs. Use this whenever you have ≥2 formats to apply. Example: [{\"range\":\"A1\",\"format\":{\"bold\":true,\"font_size\":16}},{\"range\":\"B8:F8\",\"format\":{\"number_format\":\"#,##0\"}},{\"range\":\"B9:F9\",\"format\":{\"number_format\":\"0.0%\"}}].",
                        "items": {
                            "type": "object",
                            "properties": {
                                "range":  { "type": "string" },
                                "format": { "type": "object" }
                            },
                            "required": ["range", "format"]
                        }
                    }
                },
                "required": ["sheet"]
            }
        }),
        json!({
            "name": "set_column_width",
            "description": "Set the pixel width of one or more columns. **Always prefer the bulk `operations` form when you need multiple widths** (e.g. label column A wide, data columns narrower) — one call instead of N. The flat single-width form is kept for trivial cases. Default Excel column is ~64px; a wide column for labels is typically 180–220.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "operations": {
                        "type": "array",
                        "description": "Bulk form. Each entry sets a width on a group of columns. Use this when you need different widths for different column groups.",
                        "items": {
                            "type": "object",
                            "properties": {
                                "columns": { "type": "string", "description": "Comma-separated column letters, e.g. \"A\" or \"B,C,D,E,F\"." },
                                "width": { "type": "number", "minimum": 1, "description": "Pixel width." }
                            },
                            "required": ["columns", "width"]
                        }
                    },
                    "columns": { "type": "string", "description": "Single-op form (omit `operations`). Comma-separated column letters." },
                    "width": { "type": "number", "minimum": 1, "description": "Single-op form pixel width." }
                },
                "required": ["sheet"]
            }
        }),
        json!({
            "name": "set_row_height",
            "description": "Set the pixel height of one or more rows. **Always prefer the bulk `operations` form when you need multiple heights** (e.g. title row tall, body rows short) — one call instead of N. The flat single-height form is kept for trivial cases. Default row is ~24.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "operations": {
                        "type": "array",
                        "description": "Bulk form. Each entry sets a height on a group of rows.",
                        "items": {
                            "type": "object",
                            "properties": {
                                "rows": { "type": "string", "description": "Comma-separated 1-indexed row numbers, e.g. \"1\" or \"4,5,6,7,8\"." },
                                "height": { "type": "number", "minimum": 1, "description": "Pixel height." }
                            },
                            "required": ["rows", "height"]
                        }
                    },
                    "rows": { "type": "string", "description": "Single-op form (omit `operations`). Comma-separated 1-indexed row numbers." },
                    "height": { "type": "number", "minimum": 1, "description": "Single-op form pixel height." }
                },
                "required": ["sheet"]
            }
        }),
        json!({
            "name": "merge_cells",
            "description": "Merge a rectangular range into a single visible cell (e.g. for section headers). `range` is an A1 range like \"A1:F1\".",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "range": { "type": "string" }
                },
                "required": ["sheet", "range"]
            }
        }),
        json!({
            "name": "unmerge_cells",
            "description": "Unmerge a previously merged range.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "range": { "type": "string" }
                },
                "required": ["sheet", "range"]
            }
        }),
        json!({
            "name": "set_note",
            "description": "Attach a note (Excel-style cell comment) to a single cell — a small indicator that shows the given text on hover. Use this for methodology callouts, source citations, or caveats that would clutter the cell's own value. Overwrites any existing note on that cell.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "cell": { "type": "string", "description": "A1 address, e.g. \"B7\"." },
                    "text": { "type": "string" }
                },
                "required": ["sheet", "cell", "text"]
            }
        }),
        json!({
            "name": "delete_note",
            "description": "Remove a cell's note, if it has one. No-op if the cell has no note.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "cell": { "type": "string", "description": "A1 address, e.g. \"B7\"." }
                },
                "required": ["sheet", "cell"]
            }
        }),
        json!({
            "name": "stage_data",
            "description": "Stage REPORTED figures into a data sheet the moment you extract them from a source — then build the model by FORMULA REFERENCE to the staged cells (e.g. =Data!C5) instead of retyping numbers or deriving them in your head. The harness picks the placement (blocks stack down the staging sheet) and the result returns the EXACT address of every row and column keyed by the labels you passed, so referencing costs nothing. Staged blocks render with their source line; presentation sheets that reference them are auditable by construction — a reader can trace every model figure back to its filing. Use one call per source table (an income statement, a segment table, a KPI set). Values must be RAW numbers exactly as reported (rule 18) — no derived or adjusted figures here; derive those by formula in the model. Figures you stage do NOT need to be duplicated in keep_pages `extracted_notes` — the staged cells are the durable record.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string", "description": "Staging sheet name (default \"Data\"). Auto-created if new." },
                    "title": { "type": "string", "description": "Block caption, e.g. \"AAPL 10-Q Q3 FY2026 — income statement (reported)\"." },
                    "source": { "type": "string", "description": "REQUIRED provenance: the URL or filing name these figures come from." },
                    "units": { "type": "string", "description": "e.g. \"$M\", \"€m\". Rendered with the source line." },
                    "columns": { "type": "array", "items": { "type": "string" }, "description": "Period/header labels for the value columns, e.g. [\"FY2024A\", \"FY2025A\"]. 1–24 entries." },
                    "rows": {
                        "type": "array",
                        "description": "1–80 entries. Each row: {\"label\": \"Revenue — Products\", \"values\": [307001, 296100]} (values align with `columns`; omit/null a slot to leave it blank). Shorthand [label, v1, v2, ...] also accepted.",
                        "items": {}
                    }
                },
                "required": ["title", "source", "columns", "rows"]
            }
        }),
        json!({
            "name": "create_sheet",
            "description": "Create a new sheet (worksheet tab) in the workbook. Use a clear human-readable name like \"Assumptions\", \"Q4 2024\", \"Inputs\". The new sheet starts empty — write cells with set_cell / set_range afterwards. NOTE: writes to a not-yet-existing sheet name auto-create it, so an explicit create_sheet is only needed when you want to set a tab_color or create an empty tab ahead of time.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "name": { "type": "string", "description": "Sheet name (must be unique in the workbook)." },
                    "tab_color": { "type": ["string", "null"], "description": "Optional CSS color for the tab, e.g. \"#22c55e\"." }
                },
                "required": ["name"]
            }
        }),
        json!({
            "name": "delete_sheet",
            "description": "Remove a sheet from the workbook. Permanent — every cell on this sheet is lost. Cross-sheet references in other sheets will become #REF!.",
            "input_schema": {
                "type": "object",
                "properties": { "name": { "type": "string" } },
                "required": ["name"]
            }
        }),
        json!({
            "name": "rename_sheet",
            "description": "Rename a sheet. Cross-sheet formula references in other sheets will be updated automatically by the workbook engine.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "old_name": { "type": "string" },
                    "new_name": { "type": "string" }
                },
                "required": ["old_name", "new_name"]
            }
        }),
        json!({
            "name": "clear_range",
            "description": "Empty the values AND formulas of cells in `range`. Use this when you want to delete content (not just overwrite with new values). `range` is an A1 range like \"A1\", \"A1:F1\", or \"B2:D10\". Cell formatting is preserved.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "range": { "type": "string" }
                },
                "required": ["sheet", "range"]
            }
        }),
        json!({
            "name": "insert_rows",
            "description": "Insert `count` empty rows starting at (and shifting down from) row `before` (1-indexed Excel row number). Existing content at row `before` and below shifts down. Use this to make space in the middle of a model. Charts and conditional-formatting references in unrelated areas of the workbook may not auto-adjust to the new layout.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "before": { "type": "integer", "minimum": 1, "description": "1-indexed row number to insert before." },
                    "count": { "type": "integer", "minimum": 1, "default": 1 }
                },
                "required": ["sheet", "before"]
            }
        }),
        json!({
            "name": "delete_rows",
            "description": "Delete `count` rows starting at row `start` (1-indexed). Content below shifts up. Permanent — the removed rows are gone.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "start": { "type": "integer", "minimum": 1 },
                    "count": { "type": "integer", "minimum": 1, "default": 1 }
                },
                "required": ["sheet", "start"]
            }
        }),
        json!({
            "name": "insert_columns",
            "description": "Insert `count` empty columns starting at (and shifting right from) column `before` (column letter like \"C\"). Existing content shifts right.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "before": { "type": "string", "description": "Column letter to insert before, e.g. \"C\"." },
                    "count": { "type": "integer", "minimum": 1, "default": 1 }
                },
                "required": ["sheet", "before"]
            }
        }),
        json!({
            "name": "delete_columns",
            "description": "Delete `count` columns starting at column `start` (column letter like \"C\"). Content to the right shifts left. Permanent.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "start": { "type": "string" },
                    "count": { "type": "integer", "minimum": 1, "default": 1 }
                },
                "required": ["sheet", "start"]
            }
        }),
        json!({
            "name": "freeze_panes",
            "description": "Freeze the top `rows` rows and left `cols` columns of a sheet so they stay visible while scrolling. Classic financial-model setup is `freeze_rows: 1, freeze_cols: 1` (lock the header row and label column). Pass 0 to disable that axis.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "freeze_rows": { "type": "integer", "minimum": 0, "description": "Number of top rows to freeze." },
                    "freeze_cols": { "type": "integer", "minimum": 0, "description": "Number of left columns to freeze." }
                },
                "required": ["sheet", "freeze_rows", "freeze_cols"]
            }
        }),
        json!({
            "name": "unfreeze_panes",
            "description": "Remove any freeze panes from a sheet.",
            "input_schema": {
                "type": "object",
                "properties": { "sheet": { "type": "string" } },
                "required": ["sheet"]
            }
        }),
        json!({
            "name": "hide_rows",
            "description": "Hide one or more rows. `rows` is a comma-separated list of 1-indexed row numbers (e.g. \"3\" or \"3,5,12\"). Hidden rows still hold data but don't display in the workbook.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "rows": { "type": "string", "description": "Comma-separated 1-indexed row numbers." }
                },
                "required": ["sheet", "rows"]
            }
        }),
        json!({
            "name": "show_rows",
            "description": "Unhide one or more rows. Same `rows` format as hide_rows.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "rows": { "type": "string" }
                },
                "required": ["sheet", "rows"]
            }
        }),
        json!({
            "name": "hide_columns",
            "description": "Hide one or more columns. `columns` is a comma-separated list of column letters (e.g. \"D\" or \"D,F,H\").",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "columns": { "type": "string", "description": "Comma-separated column letters." }
                },
                "required": ["sheet", "columns"]
            }
        }),
        json!({
            "name": "show_columns",
            "description": "Unhide one or more columns. Same `columns` format as hide_columns.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string" },
                    "columns": { "type": "string" }
                },
                "required": ["sheet", "columns"]
            }
        }),
        json!({
            "name": "read_range",
            "description": "Read evaluated cell values from a specific A1 range WITHOUT modifying anything. Use this to verify your own work, confirm where the assumption block actually landed before composing dependent formulas, or pull current values you need to reference. Returns each non-empty cell's value AND formula (if any) in `A1 = value` form. Costs one tool turn but is the cheapest way to prevent the formula-pointing-at-wrong-row class of bug on long sheets. Range limited to 500 cells per call. To LOCATE a row/section/header by NAME, don't probe with this — use `find_rows` or `describe_workbook` first, then read the exact range.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string", "description": "Sheet name (case-sensitive)." },
                    "range": { "type": "string", "description": "A1 range, e.g. \"G45:K45\", \"A1:Z10\", or a single cell \"B7\"." }
                },
                "required": ["sheet", "range"]
            }
        }),
        json!({
            "name": "describe_workbook",
            "description": "Get a structural map of the ACTIVE workbook without reading cell data. Per sheet: the used range, the detected header row with its period/column headers mapped to column letters (e.g. {\"BU\": \"FY2026E\"}), and a table of contents of sections — contiguous row bands with their titles, like {\"rows\": \"5-42\", \"title\": \"Revenue build\"}. Also lists workbook defined names. Served from an in-memory index, so it's near-instant and returns in ONE call what would take many read_range probes. Call it once to orient on any sheet whose preview is truncated. Optional `sheet` filters to a single sheet.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "sheet": { "type": "string", "description": "Optional: limit the map to this sheet (case-sensitive)." }
                },
                "required": []
            }
        }),
        json!({
            "name": "find_rows",
            "description": "Search the ACTIVE workbook's row labels and column headers by name and get exact positions back — e.g. query \"total revenue\" returns each matching row's sheet, 1-indexed row number, containing section, and one evaluated sample cell (\"BT209 = =BT194 → 39001.1\") read live from the grid. Matching is case-insensitive substring plus word-prefix (\"tot rev\" matches \"Total revenue\"). Near-instant. Use this — NOT exploratory read_range probes — whenever you need to LOCATE a row or header (\"where is Subscriber ARPU?\", \"which column is FY2027E?\"); then read_range the exact range you need. Searches only the active workbook, not reference workbooks.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "query": { "type": "string", "description": "Label text to find, e.g. \"total revenue\"." },
                    "sheet": { "type": "string", "description": "Optional: search only this sheet (case-sensitive)." },
                    "max_results": { "type": "integer", "minimum": 1, "maximum": 50, "description": "Max matches to return (default 20)." }
                },
                "required": ["query"]
            }
        }),
        json!({
            "name": "read_reference",
            "description": "Read cell values from a REFERENCE workbook — one of the read-only xlsx files the user attached to this session (listed under \"Reference workbooks\" in the context, addressed by their label, e.g. \"AAPL-GS.xlsx\"). Use this to pull figures beyond the truncated previews: full estimate rows, assumption blocks, segment detail. Returns each non-empty cell's value AND formula in the same shape as read_range. References are read-only — you can never write to them; write tools always target the active workbook. Range limited to 500 cells per call.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "workbook": { "type": "string", "description": "The reference workbook's label exactly as listed in the context, e.g. \"AAPL-GS.xlsx\"." },
                    "sheet": { "type": "string", "description": "Sheet name within the reference workbook." },
                    "range": { "type": "string", "description": "A1 range, e.g. \"B5:N40\", or a single cell \"B7\"." }
                },
                "required": ["workbook", "sheet", "range"]
            }
        }),
        json!({
            "name": "fetch_web",
            "description": "Fetch one or more web pages in parallel and return their readable text. Use this to ground edits in fresh data: a company's latest revenue from an SEC filing, a current price, a docs page. Up to 5 URLs per call. **PDF links work too** — press releases, IR decks and filings are text-extracted automatically (tables come through as aligned text lines), so fetch the IR PDF directly instead of hunting for an HTML mirror; only scanned/image-only PDFs fail. **The result comes back PAGINATED** — numbered `PAGE 1 … PAGE N` under a short `source_id` (e.g. \"src1\"), because a filing is large. **Read every page NOW and extract every figure you'll need** — on your VERY NEXT turn you MUST call `keep_pages` first; every other tool call is refused with an error until you do. In that call you (a) name the pages worth keeping AND (b) write your extracted figures into the required `extracted_notes` field. `extracted_notes` is your DURABLE scratchpad — it stays in the assistant message forever even after pages are evicted, so it's how the figures survive the rest of the build. Pages you don't keep are evicted from context (only `read_source` can claw one back, a last resort). So review thoroughly here and plan exactly what you'll write into `extracted_notes` next turn. Content is NOT auto-inserted into the spreadsheet — you decide what to write.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "urls": {
                        "type": "array",
                        "items": { "type": "string" },
                        "minItems": 1,
                        "maxItems": 5,
                        "description": "List of URLs to fetch. Each must be a full URL including scheme."
                    }
                },
                "required": ["urls"]
            }
        }),
        json!({
            "name": "edgar_lookup",
            "description": "Look up a U.S. SEC registrant's recent filings and get EXACT primary-document URLs (10-K, 10-Q, 8-K, ...) — no guessing, no aggregators. Pass a ticker (\"AAPL\") or company name (\"Delivery Hero\" won't match — non-US companies aren't SEC registrants; use their IR page via web_search instead). Returns [{form, filed, period, url}] — then `fetch_web` the URL(s) you need: the filing IS the authoritative source for reported line items (rule 22), where aggregators round, restate, and mix bases. Prefer this over web_search for any US-listed company's financials.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "company": { "type": "string", "description": "Ticker symbol (best) or company name." },
                    "forms": { "type": "array", "items": { "type": "string" }, "description": "Form types to include. Default [\"10-K\", \"10-Q\"]. Exact match (add \"10-K/A\" explicitly if you want amendments)." },
                    "count": { "type": "integer", "description": "Max filings to return, newest first. Default 6, max 12." }
                },
                "required": ["company"]
            }
        }),
        json!({
            "name": "keep_pages",
            "description": "Prune a fetched source. **On the turn right after a `fetch_web` this call is MANDATORY and must come first — every other tool call is refused with an error until each pending source is pruned.** Pass the short `source_id` from the fetch_web result (e.g. \"src1\"), the 1-indexed page numbers worth keeping in context for the rest of the build, AND `extracted_notes` — your own bullet-style summary of every figure you'll need from this source. **`extracted_notes` is your durable scratchpad**: page text gets evicted from context after this turn, but your `extracted_notes` stays in the assistant message forever (the tool_use input is never compacted). Write enough that you can rebuild the model from notes alone if every page were dropped — actual numbers with units and years (e.g. \"FY2024A revenue $97.69B; COGS $79.11B; OpEx $10.04B; net income $7.13B\"), not vague reminders (\"revenue and costs\"). For `pages`: choose the few that hold the figures relevant to the user's request and to a high-quality output; pass `[]` only if your `extracted_notes` already captures everything you'll need. ONLY the kept pages stay in (cached) context; everything else is dropped (recoverable via `read_source`, a last resort). If you fetched several sources, call keep_pages once per `source_id`. Copy the source_id EXACTLY as given.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "source_id": { "type": "string", "description": "The short source_id returned by fetch_web, e.g. \"src1\". Copy it exactly." },
                    "pages": {
                        "type": "array",
                        "items": { "type": "integer", "minimum": 1 },
                        "description": "1-indexed page numbers to retain. Empty array = drop the entire source (only safe if `extracted_notes` already has every figure you'll need)."
                    },
                    "extracted_notes": {
                        "type": "string",
                        "description": "Bullet-style scratchpad of every figure / fact you'll need from this source going forward. Persists in the assistant message even after pages are evicted, so write actual numbers with units and years — not vague reminders. Example: \"Tesla 10-K FY2024A: Revenue $97.69B (-1% YoY); Automotive rev $77.07B; Auto gross margin 18.4%; OpEx $10.04B; Net income $7.13B; Diluted shares 3,194M; CapEx $11.34B.\" Cap ~2 KB; if there are too many figures, prefer keeping the source page that holds them over over-stuffing this field."
                    }
                },
                "required": ["source_id", "pages", "extracted_notes"]
            }
        }),
        json!({
            "name": "read_source",
            "description": "Re-surface specific pages of an already-fetched source whose pages were evicted from context. Pass the short `source_id` (e.g. \"src1\") + 1-indexed `pages`. Returns the raw page text. Two uses: (1) LAST-RESORT recovery within the current run if you genuinely under-kept during review — the normal path is to extract everything when you first see the source and `keep_pages` the rest; (2) the NORMAL path for sources fetched on EARLIER turns — fetched sources persist for the whole session and are listed under \"# Previously fetched sources\" in the user message. ALWAYS prefer read_source over re-fetching one of those URLs.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "source_id": { "type": "string", "description": "The short source_id returned by fetch_web, e.g. \"src1\". Copy it exactly." },
                    "pages": {
                        "type": "array",
                        "items": { "type": "integer", "minimum": 1 },
                        "minItems": 1,
                        "description": "1-indexed page numbers to recover."
                    }
                },
                "required": ["source_id", "pages"]
            }
        }),
        json!({
            "name": "define_name",
            "description": "Create (or replace) a workbook-scoped named range so formulas can reference a cell/range by a readable name instead of a raw address — e.g. name `GrowthRate` for the assumption cell, then write `=Revenue*(1+GrowthRate)`. The name persists in the saved .xlsx. `name` must be a valid Excel name (letters, digits, underscores; can't look like a cell reference such as \"A1\"). `ref` must be sheet-qualified, e.g. \"Model!B22\" or \"Model!B5:B12\" (absolute `$` signs optional — they're added automatically). Calling again with the same `name` replaces its target.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "name": { "type": "string", "description": "Defined name, e.g. \"GrowthRate\" or \"Revenue\". Must not look like a cell address." },
                    "ref": { "type": "string", "description": "Sheet-qualified target, e.g. \"Model!B22\" or \"Assumptions!B5:B12\"." },
                    "sheet": { "type": "string", "description": "Optional. Only needed if `ref` is a bare range without a sheet prefix; supplies the sheet for it." }
                },
                "required": ["name", "ref"]
            }
        }),
        json!({
            "name": "done",
            "description": "Call this LAST when all edits are written and formulas evaluate cleanly. The runtime SCANS every cell you touched for #REF! / #DIV/0! / #VALUE! / #N/A / #NAME? / self-references BEFORE accepting done — if any are found, done is REJECTED and you get the error list back; fix them (compose against row_map) and call done again. Provide a short justification of what you changed and why so the user can review, plus a `turn_summary` — your durable memory for FUTURE turns in this session.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "justification": { "type": "string", "description": "1–3 sentence summary of the edits made and the reasoning. Shown to the user in the review panel." },
                    "turn_summary": { "type": "string", "description": "Structured record for YOUR OWN future turns — your tool history and reasoning are NOT replayed on the next user message; only this summary (plus a digest of the conversation) survives. Cover: (1) WHAT you built/changed and WHERE (sheets, ranges, key row numbers), (2) key DECISIONS and interpretations you made and why, (3) key FACTS/FIGURES learned from fetched sources or references, with units and years (e.g. \"FY2024A revenue $97.69B; gross margin 19.0%; source: 10-K src1\"). Write actual numbers, not vague reminders. ~2 KB max." }
                },
                "required": ["justification", "turn_summary"]
            }
        }),
    ]
}

/// Build the agent's system prompt. `has_web_search` controls rule 10's
/// wording: Claude has access to a hosted `web_search` server tool; Codex
/// (ChatGPT-subscription) doesn't, so we describe `fetch_web` as the only
/// option and warn against guessing deep URLs. Avoids advertising a tool
/// the model can't actually call on the Codex path.
pub fn system_prompt(has_web_search: bool) -> String {
    let rule_10 = if has_web_search {
        "10. If the user asks for something that depends on external data (current prices, financial filings, company facts you're not sure about), pull the data BEFORE writing cells. Two tools available:\n    - **`web_search`** — provider-hosted. Use this when you DON'T know the exact URL: \"Tesla 10-K 2024 SEC filing\", \"Apple Q3 2025 earnings\", \"Airbnb listings Aspen Dec 20-22\". Returns ranked search results with snippets — read them, pick the right source.\n    - **`fetch_web`** — pulls a specific URL's text. Use this once you have a URL (either from the user, from web_search results, or a well-known homepage like sec.gov or yahoo.com/finance).\n    Typical pattern for \"build me a [company] DCF\": `web_search(\"[company] 10-K latest\")` → pick the SEC filing URL from results → `fetch_web(that URL)` → extract revenue/segments/margins → build the model. Don't invent numbers — search and fetch them. For US-listed companies, call `edgar_lookup` FIRST — it returns exact primary filing URLs (10-K/10-Q) to fetch, beating both guessed URLs and aggregator restatements."
    } else {
        "10. If the user asks for something that depends on external data (current prices, financial filings, company facts you're not sure about), use `fetch_web` to pull the page(s) first. Read the returned text, extract the values you need, THEN call your write tools. Don't invent numbers — fetch them. For URLs you don't know, prefer well-known landing pages (sec.gov/edgar/browse, yahoo.com/finance, the company's investor relations page) over guessing deep URLs — guessed accession numbers and article slugs frequently 404 because they're unique and time-sensitive. For US-listed companies, call `edgar_lookup` FIRST — it returns exact primary filing URLs (10-K/10-Q), no guessing needed."
    };
    let base = r##"You are a spreadsheet editing agent. The user describes a change in natural language and you execute it by calling tools that write cells.

Rules:
1. Don't fragment work that belongs together. Prefer `set_range` over many `set_cell` calls when filling contiguous areas. When the new content is a shifted copy of existing cells — extending a projection by another period, duplicating a section — use `copy_range` instead of re-typing formulas: it shifts relative references server-side exactly like Excel copy/paste and carries formatting, so it can't make off-by-one mistakes. **For formatting, ALWAYS use `set_format` with the bulk `operations` array** — collapse every format you need (headers, %, $, alignment, bold) into ONE set_format call, not many. Independent tool calls fire together in one turn (rule 12a). For small EDITS, keep the whole change tight — a two-cell fix shouldn't stretch across five turns. For multi-section BUILDS the economics are different: turns are cheap (your prompt prefix is cached between turns) and in-head simulation is not — write one section, read it back, then continue (rule 12b). Never plan an entire model mentally and serialize it as one giant write: that path routinely exhausts the per-turn output budget mid-write, and lands mentally-derived numbers as unverifiable hardcodes.
1a. **When an edit is loop-shaped, write a script instead of a payload.** `run_script` executes a short JavaScript program against the workbook (see its description for the API) and lands everything as one batch: per-row conditional edits, repeating formula patterns across many rows or columns, transformations that read existing cells to decide what to write. One run_script turn replaces dozens of set_cell/set_range turns and computes its addresses programmatically instead of drifting off-by-one mid-payload. Division of labor: `set_range` for one contiguous literal block; `copy_range` for a shifted copy of existing cells; `run_script` as soon as there's a loop, a condition, or address arithmetic. Inside a script, the same conventions apply — write formulas (not precomputed results) for derived values, raw numbers not formatted strings.
2. All coordinates are **Excel A1 notation**: column letters + row number, where row 1 is the first row. Never use 0-indexed numeric coordinates.
3. **Formula references must point at the FINAL placement of cells.** If you `set_range` with `top_left="A9"` and want the cell at row offset 2 to reference the data one row above your block, that data is at row 8, so write `=A8`. Mentally walk through where each value LANDS before composing its formula. **Before submitting**, scan each formula for: (a) self-references (cell A7 containing `=A7` or anything that depends on A7), (b) off-by-one errors in ranges (`=SUM(B1:B10)` when your data is B2:B11), (c) unintended circular refs where a downstream formula's output feeds back into its own input. Every model MUST land with zero `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?` errors — if a divisor could be zero, wrap with `IFERROR(...,0)` or guard with `IF(divisor=0,...,calc)`.
4. Use formulas (strings starting with '=') whenever a derived value belongs in the cell — e.g. `=SUM(A1:A10)`, not the computed number. The grid will evaluate the formula.
5. Do not invent values. If the user asks for something you cannot compute from the provided context, leave a clear comment cell and explain in the `done` justification.
6. If you wrote any cells, applied any formatting, or otherwise edited the workbook, finish that turn with the `done` tool — the justification you pass shows in the Changes panel and lets the user review what you did. The `turn_summary` you pass is your memory across user messages: your tool calls and reasoning are NOT replayed on the next message, so record what you did (with sheet/row locations), key decisions, and key sourced figures there — a vague summary means re-deriving all of it next turn. **If you didn't make any edits** (the user thanked you, asked a clarifying question, asked you to explain something, or you need more info before editing), just respond with normal text and end the turn. Do NOT call `done` with a fake justification like "no changes needed" — that creates a meaningless empty batch the user has to click through. Pure conversational responses don't need a tool call.
6a. The bar for asking the user a clarifying question is **"I genuinely cannot impute reasonable intent."** If you can pick a sensible interpretation of a short or vague prompt, do that — make the edit, and note your interpretation briefly in the `done` justification (the user can reject the batch if you guessed wrong). Only ask back when the request actually under-specifies what you'd need to act: missing data the model can't reasonably fabricate (e.g. real values you don't have), targets you can't identify ("fix this" with no selection and no obvious problem area), or contradictions. Don't ask the user to pick from a menu when you could just take a sensible best guess and let them reject.
    One genuine exception: when the deliverable you're about to build materially diverges from the user's words in METRIC or COVERAGE — they asked for X and the available data only supports a component or proxy of X (asked for "bookings by game" but only full-game unit-sale revenue is buildable, with live services as a residual; asked for ten years, only three are sourceable) — ask back before building. Catch this at PLANNING time — if you notice yourself re-scoping the deliverable more than twice while "inferring intent", that's the signal. If you proceed anyway, the divergence must be the FIRST sentence of your `done` justification ("You asked for bookings by game; this shows full-game revenue by game — live services aren't split by title in any available source"), never a footnote or a sheet subtitle.
7. Keep prose between tool calls minimal — the user does not read it during streaming.
8. The user reviews every batch before it is saved. They can reject. Bias toward small, focused edits.
9. If the user message contains a "Prior turns in this session" block, treat it as ground truth about what edits already exist — do NOT redo them. The grid context already reflects accepted prior edits; rejected ones were rolled back.
{RULE_10}
10b. **Fetched sources are paginated, and pruning is FORCED — and that prune turn is when you commit figures to durable memory.** A `fetch_web` result comes back as numbered pages (`PAGE 1 … PAGE N`) under a short `source_id` like "src1". Read ALL of it the moment it arrives. Then, on your very next turn, **`keep_pages(source_id, [pages], extracted_notes)` must come FIRST — until every pending source is pruned, any other tool call is refused with an error and NOT executed** (you'd just have to re-issue it). Three things, all required:\n    - `pages`: the 1-indexed pages worth keeping in (cached) context — choose the few that hold the figures relevant to the user's request.\n    - `extracted_notes`: **your durable scratchpad — write actual numbers with units and years**, e.g. `"FY2024A revenue $97.69B; COGS $79.11B; gross margin 19.0%; OpEx $10.04B; CapEx $11.34B; diluted shares 3,194M."` Page text gets evicted after this turn, but `extracted_notes` lives on the assistant message forever. **If you don't write a figure here, assume it's gone** — do not rely on remembering numbers between turns, especially if you don't see your prior reasoning replayed.\n    - `source_id`: copy EXACTLY as given.\n    Pass `pages: []` only if `extracted_notes` already captures every figure you'll need. ONLY the kept pages stay in (cached) context — `read_source(source_id, [...])` can claw back an evicted page, but it's a LAST RESORT (extract thoroughly here so you rarely need it). **`fetch_web` returns the ENTIRE page regardless of any `#fragment`/anchor** — once you've fetched a URL you have every section; NEVER re-fetch the same URL to "jump to" a section or "get more detail," it just wastes a turn. **Fetched sources persist across user messages**: if the user message lists sources under "# Previously fetched sources", their pages are still stored from earlier turns — use `read_source(source_id, [pages])` to read them instead of re-fetching the URL.
10c. **A failed fetch yields NOTHING.** If `fetch_web` returns 0 pages, you have no data from that source: never narrate as if you retrieved it, and never present figures recalled from memory as if they were fetched — find an alternative source (a different URL, or web_search first). If you do proceed on model knowledge for some figure, say so plainly rather than implying it was sourced.
10a. **Today's date is at the top of every user message.** Use it. For financial models, this is your historical-vs-forecast boundary — fiscal years that have already ended are *historicals* (label "A" or no suffix), the current year and beyond are *forecasts* (label "E"). Don't pattern-match on your training cutoff and treat 2024 as "current" when today is 2026. **A forecast model MUST include at least the most recent completed fiscal year as an actuals anchor column — even when the user asks only for future years ("forecast 2026–2028").** Forecasts hang off reported actuals, not thin air: source that anchor column, tie it to reported figures (rule 21), and derive the first forecast year's growth off it.
11. When you write a financial / numerical model, ALWAYS finish with a single bulk `set_format` call that applies the right `number_format` to every numeric region: percentages (margins, growth, tax rate) get `"0.0%"` or `"0.00%"`, currency rows get `"$#,##0"` or `"$#,##0.00"`, plain quantities get `"#,##0"`. A model without proper number formats looks broken to the user — don't skip this step.
12. Order of operations for a fresh **compact, single-block** model (one section that fits in one coherent `set_range` — a simple DCF, a comps table, a one-segment projection): (a) `fetch_web` if needed, (b) ONE big `set_range` for all data, (c) ONE bulk `set_format` covering every numeric region, (d) **ONE** `set_column_width` with the bulk `operations` array covering every width you need, **ONE** `set_row_height` with the bulk `operations` array covering every height, plus `merge_cells` if needed, (e) `done`. That's ~6 tool calls total for a complete model — do NOT call `set_column_width` once per column group or `set_row_height` once per row group, that wastes turns. For large multi-section builds, this recipe does NOT apply — use the section-by-section cadence in rule 12b instead.
12b. **Build large models SECTION BY SECTION — the sheet is your working memory, not your context window.** When a build has multiple sections whose numbers feed each other (a multi-segment P&L rolling into a consolidation, a 3-statement model, per-division schedules feeding a summary), do NOT simulate the entire model in your reasoning and serialize it as one giant write at the end. Two failure modes make that approach unreliable: (a) long chains of mental arithmetic drift — you cannot reliably carry a hundred derived figures across sections in your head, and every downstream number inherits the error; (b) a single row-mapping mistake in a monolithic write corrupts every section at once, with no checkpoint to recover from. Instead, work in a build–verify loop, one section per iteration:
    - **Write one section** — labels, hardcodes, and formulas for that section only. Within the section, still use bulk writes (`set_range` / `run_script`); this rule changes the *cadence*, not the write granularity.
    - **Read it back** — `read_range` the section's computed subtotals, margins, and check cells in the same turn-boundary before touching the next section.
    - **Verify against ground truth** — compare the EVALUATED values to reported figures and to your expectations. The recalc engine's output is authoritative; your in-head arithmetic is not. If a margin looks aggressive or a check row isn't 0, fix it NOW, while the error is one section deep.
    - **Then start the next section**, referencing upstream numbers via cell formulas and reasoning from values you actually read back — never from figures you derived mentally several sections ago.
    The extra turns are the point, not waste: each read-back converts "I believe this ties" into "the engine confirmed this ties", and an error caught after one section costs one section instead of the whole build. Rough sizing: if the build is big enough to warrant per-section check rows (rule 21), it is big enough for this cadence.
12a. **Emit independent tool calls TOGETHER in one turn — don't serialize them across turns.** A turn is one full model round-trip and is the expensive unit; the number of *tool calls* barely matters, the number of *turns* does. The entire layout pass — `set_format`, `set_column_width`, `set_row_height`, `merge_cells`, `freeze_panes` — applies instantly and needs no read-back, so issue ALL of them as parallel tool calls in a **single** turn (they're applied in order on your behalf). The only reason to take a fresh turn is a genuine data dependency: you must SEE evaluated values (via read-back or `read_range`) before you can write the next thing correctly. Concretely: after the data `set_range` lands, do the whole layout pass + `done` in ONE turn rather than one tool per turn. Don't wait to "see" formatting or widths land — they can't fail in a way you'd react to.
13. If the user message contains a "User focus" block — selection or @-mentioned cells — treat those cells as the primary target for the edit. The user picked them deliberately. Don't edit cells outside this focus unless the prompt explicitly asks you to.
14. **Preserve existing template conventions.** When editing a workbook that already has established formatting, color coding, fonts, header style, section structure, or number-format patterns, mirror those conventions for any new cells you add. Don't impose your default styling on a sheet whose author has already chosen one. Existing template conventions ALWAYS override the financial-model defaults in rule 15. **The sheet preview includes a "Style conventions" block and per-cell `[fmt …]` annotations** (bold, fill, font color, number_format, etc.) — treat those as ground truth for how this workbook is styled. If you're filling cells into a row that uses Calibri 11 with a specific blue header bar, your additions should match — not introduce a new font or color scheme.
15. **Financial model conventions** (apply when building a fresh model — DCF, 3-statement, projections, comps, LBO, sensitivities — unless rule 14 says to mirror an existing template instead):
    - **Color coding** (font colors): blue (`#0000FF`) for hardcoded inputs and scenario-driver numbers the user will tweak; black (`#000000`) for all formulas and calculations; green (`#008000`) for links that pull from other worksheets in the same workbook; red (`#FF0000`) for external-file links. Yellow background (`#FFFF00`) for key assumption cells that need user attention.
    - **Assumptions placement**: every growth rate, margin, multiple, and driver lives in its own cell (typically in a dedicated assumptions block). Formulas reference those cells — never hardcode `=B5*1.05`, write `=B5*(1+$B$22)` where `$B$22` is the growth assumption. For headline drivers the user will tweak often, consider `define_name` to give the cell a readable name (e.g. `GrowthRate`) so formulas read `=B5*(1+GrowthRate)` — but don't over-name; reserve it for a handful of key drivers, not every cell.
    - **Number formats**: currency uses `"#,##0;(#,##0);-"` (parens for negatives, dash for zero) — include units in the header (e.g. "Revenue ($mm)"); percentages default to `"0.0%;(0.0%);-"`; valuation multiples use `"0.0x"`; years are text strings (`"2024"` not formatted as `2,024`); negatives always render in parentheses, not with a minus sign.
    - **Font**: stick to a single professional font across the model (Calibri or Arial). Don't mix.
16. **Compose formulas against the `row_map`, never against your remembered layout.** Every `set_cell`/`set_range` result includes a `row_map` — the ACTUAL row each label landed at, e.g. `[{"row":58,"label":"Revenue"},{"row":59,"label":"Cost of Revenue"}]`. This is ground truth; your mental layout is not. The #1 failure mode is formula references drifting by N rows — an extra header row pushes everything down, a row gets duplicated, an A1 anchor is off by one — and then formulas point at the wrong row or at their own cell (circular reference). To avoid it: **write the labels + hardcoded data FIRST (one `set_range`), read the `row_map` that comes back, THEN write the formulas in the next turn referencing only rows that appear in `row_map`.** Before calling a model done, scan: does every formula's row reference match a `row_map` entry for the line it's supposed to use? If your plan said "Gross Profit is row 19" but `row_map` says row 20, the `row_map` wins — fix the formula. Don't "rebuild the whole block to realign it" (that just re-rolls the same dice) — read the actual positions and patch.
16a. **`done` is gated on layout validation.** When you call `done`, the runtime scans every cell you touched this batch for Excel formula errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, …) and self-references. If any are found, `done` is REJECTED — you get back `{ok:false, errors:[…]}` listing the broken cells. Patch those cells (using `row_map` / `read_range`), then call `done` again. Do NOT ignore the error list or rebuild from scratch.
17. After writing a large `set_range` with formulas, **read back a sample of the formula cells** via `read_range` to confirm they evaluated to sane magnitudes. If revenue should be in the billions and you see `1.2M`, your formula is wrong — fix it before calling `done`.
17a. **Locate before you read.** On a sheet whose preview is truncated, do NOT page through `read_range` probes to find where something lives — call `describe_workbook` once for the section map (and header-row → column letters), or `find_rows("total revenue")` for exact row numbers with an evaluated sample. One call replaces several exploratory `read_range` probes and a full turn each. Then use `read_range` only on the exact rows you've located, and reserve it for what it's for: verifying values and pulling ranges you already know you need.
17b. **Live evaluation can collapse on imported models — attribute it correctly.** GridPath re-evaluates every formula with its own engine, which cannot fully compute some constructs Excel can (INDIRECT chains, external-workbook references, add-in functions). On such cells the LIVE value can read 0, empty, or an Excel error even though the user's file is perfectly sound. `read_range` detects this on cells you haven't edited and returns Excel's last-saved value alongside: `{"cell":"BO67","value":0,"file_saved":596.6}`; the workbook context may also carry a calc-health warning naming the affected sheets. `file_saved` is attached ONLY when live is unusable (0 / empty / error) — two finite numbers that merely disagree are not flagged; the live grid value is what the user sees, report that. When `file_saved` is present: (a) treat `file_saved` as the authoritative figure for reporting, comparisons, and hardcoding; (b) do NOT build live formula links onto flagged cells — they would evaluate wrong in the grid; hardcode the `file_saved` figure instead and label its source; (c) NEVER describe the divergence to the user as an error, staleness, or "data integrity issue" in THEIR model — it is a GridPath evaluation limitation, say so plainly if it's worth mentioning at all. Cells you edited this session are never flagged; their live values are the only truth for them.
18. **Write raw numbers, not pre-formatted strings.** When putting numeric data into cells, use the bare number — `7839`, `0.215`, `-1837` — never a display-formatted string like `"7,839"`, `"21.5%"`, or `"(1,837)"`. Pre-formatted strings get stored as text, downstream formulas can't do arithmetic on them, and you waste turns re-fixing the row. Number formatting (commas, percent signs, parentheses for negatives, currency symbols) is the job of `set_format` via `number_format`, never the cell value.
19. **Light font colors REQUIRE a contrasting `background_color`.** If you set `font_color` to white (`#FFFFFF`), a near-white gray, or any pale tone, you MUST also set `background_color` to a dark color in the same `set_format` op — otherwise the text renders invisibly on the default white cell fill. The classic "dark blue header bar with white text" pattern is `{"background_color": "#1F4E79", "font_color": "#FFFFFF", "bold": true}`. Never set one without the other.
20. **Reference workbooks are READ-ONLY context.** When the user message contains a "Reference workbooks" block, those are other xlsx files the user attached for comparison — analyst models, comps, prior versions. Use their previews and the `read_reference` tool to pull figures, but NEVER try to write to them: every write tool (set_cell, set_range, set_format, …) targets the ACTIVE workbook only, and sheet names in write calls always mean active-workbook sheets. When you build a comparison from several references, pull the actual ranges via `read_reference` first (don't rely on a truncated preview for numbers you'll hardcode into cells), and label which reference each figure came from (e.g. a column per source model) so the user can trace them.
21. **Tie out to reported figures EXACTLY — compute plugs, don't accept drift.** When modeling from a company's reported financials, the anchor lines (revenue, EBIT, pretax income, net income) must match the reported values exactly, not "close": if your derived P&L is off by even €5m, do not ship it with a caveat — make the reconciling line a COMPUTED plug instead. Structure the model so this is mechanical: hardcode the sourced actuals (revenue, COGS, opex, interest, tax, net income), derive subtotals by formula (gross profit, EBIT), and let one explicitly-labeled reconciling line (e.g. "Other non-operating (derived)") absorb the residual: plug = reported_pretax − derived_EBIT − interest. Then add a **check row** (`=modeled_NI − reported_NI`, formatted as a "Check (should be 0)" line) for each anchor, and read it back: every check must evaluate to 0 before `done`. A model whose checks tie to zero is trustworthy; a model that is "approximately right" is not a model, it's a sketch.
    Run the tie-out at EXTRACTION time too, not only before `done`: the moment you pull figures you intend to build a breakdown or reconciliation from (reference-workbook rows, web sources, filings), sum the named parts against the total they must reconcile to. A negative or implausible residual means the source mixes bases or definitions (pro-forma standalone vs consolidated, pre- vs post-acquisition, calendar vs fiscal) — disqualify or re-scope that sourcing THEN, before designing the sheet around it, instead of discovering the mismatch after several build iterations.
22. **Triangulate financial data — one aggregator is not a source.** Aggregators (Yahoo Finance, stockanalysis, etc.) round, normalize, and sometimes mix consolidation scopes. For any company model: confirm the key figures against a primary source when reachable (the company's IR press release or filing via web_search/fetch_web), and watch for DEFINITION mismatches that aggregators hide — IFRS/GAAP revenue vs. "total segment revenue" (intersegment eliminations), reported vs. adjusted EBITDA, margins on revenue vs. on GMV/volume, continuing vs. discontinued operations. When two sources disagree, say which one each number came from — put a short Sources & Notes section (or tab) in the model recording source + definition per hardcoded block, and mark derived figures as derived. If guidance exists for a forecast year, sanity-check your projection against the guidance range and say so.
"##;
    base.replace("{RULE_10}", rule_10)
}
