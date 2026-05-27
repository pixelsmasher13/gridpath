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
            "description": "Set a rectangular block of values starting at `top_left` (A1 address). `values` is a 2D array — outer rows, inner cells. Strings starting with '=' inside `values` are treated as formulas; formula references must use A1 notation pointing at the FINAL placement of cells (e.g. if `top_left` is \"A9\" and the value at row index 2 / col 0 should reference the data row just above your block at A8, write \"=A8\" — not \"=A1\" or \"=A11\"). Use this for bulk writes; prefer it over many set_cell calls. **`null` and empty string `\"\"` inside `values` are treated as PRESERVE — the cell is left untouched, NOT cleared.** This means you can ship `[[\"Gross Profit\", \"\", \"\", \"\"]]` to set only the label in column A without wiping the data in B–D. To actually clear cells, use `clear_range`.",
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
            "name": "set_format",
            "description": "Apply formatting to one OR MANY cell ranges in a single call. STRONGLY PREFER the bulk `operations` form when you have more than one format to apply — it's one tool call instead of N, dramatically faster. Each operation's `format` is a partial object; only specified properties are applied, others are left as-is. Use this for: bold headers, currency / percent / number formatting on data columns, italic notes, alignment. v1 limitations: background colors and borders are not yet supported.\n\nTwo equivalent shapes:\n  • Single range: { sheet, range, format }\n  • Bulk:         { sheet, operations: [ { range, format }, ... ] }",
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
            "name": "create_sheet",
            "description": "Create a new sheet (worksheet tab) in the workbook. Use a clear human-readable name like \"Assumptions\", \"Q4 2024\", \"Inputs\". The new sheet starts empty — write cells with set_cell / set_range afterwards.",
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
            "description": "Read evaluated cell values from a specific A1 range WITHOUT modifying anything. Use this to verify your own work, confirm where the assumption block actually landed before composing dependent formulas, or pull current values you need to reference. Returns each non-empty cell's value AND formula (if any) in `A1 = value` form. Costs one tool turn but is the cheapest way to prevent the formula-pointing-at-wrong-row class of bug on long sheets. Range limited to 500 cells per call.",
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
            "name": "fetch_web",
            "description": "Fetch one or more web pages in parallel and return their readable text content. Use this to ground edits in fresh data: pull a company's latest revenue from an SEC filing, look up a current price, read a docs page, etc. Returns the extracted text for each URL — content is NOT auto-inserted into the spreadsheet, you decide what to write after reading. Up to 5 URLs per call.",
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
            "name": "done",
            "description": "Call this LAST when all edits are written. Provide a short justification of what you changed and why so the user can review.",
            "input_schema": {
                "type": "object",
                "properties": {
                    "justification": { "type": "string", "description": "1–3 sentence summary of the edits made and the reasoning." }
                },
                "required": ["justification"]
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
        "10. If the user asks for something that depends on external data (current prices, financial filings, company facts you're not sure about), pull the data BEFORE writing cells. Two tools available:\n    - **`web_search`** — provider-hosted. Use this when you DON'T know the exact URL: \"Tesla 10-K 2024 SEC filing\", \"Apple Q3 2025 earnings\", \"Airbnb listings Aspen Dec 20-22\". Returns ranked search results with snippets — read them, pick the right source.\n    - **`fetch_web`** — pulls a specific URL's text. Use this once you have a URL (either from the user, from web_search results, or a well-known homepage like sec.gov or yahoo.com/finance).\n    Typical pattern for \"build me a [company] DCF\": `web_search(\"[company] 10-K latest\")` → pick the SEC filing URL from results → `fetch_web(that URL)` → extract revenue/segments/margins → build the model. Don't invent numbers — search and fetch them."
    } else {
        "10. If the user asks for something that depends on external data (current prices, financial filings, company facts you're not sure about), use `fetch_web` to pull the page(s) first. Read the returned text, extract the values you need, THEN call your write tools. Don't invent numbers — fetch them. For URLs you don't know, prefer well-known landing pages (sec.gov/edgar/browse, yahoo.com/finance, the company's investor relations page) over guessing deep URLs — guessed accession numbers and article slugs frequently 404 because they're unique and time-sensitive."
    };
    let base = r##"You are a spreadsheet editing agent. The user describes a change in natural language and you execute it by calling tools that write cells.

Rules:
1. Output as few tool calls as possible. Prefer `set_range` over many `set_cell` calls when filling contiguous areas. **For formatting, ALWAYS use `set_format` with the bulk `operations` array** — collapse every format you need (headers, %, $, alignment, bold) into ONE set_format call, not many. Each separate tool call costs a full LLM turn.
2. All coordinates are **Excel A1 notation**: column letters + row number, where row 1 is the first row. Never use 0-indexed numeric coordinates.
3. **Formula references must point at the FINAL placement of cells.** If you `set_range` with `top_left="A9"` and want the cell at row offset 2 to reference the data one row above your block, that data is at row 8, so write `=A8`. Mentally walk through where each value LANDS before composing its formula. **Before submitting**, scan each formula for: (a) self-references (cell A7 containing `=A7` or anything that depends on A7), (b) off-by-one errors in ranges (`=SUM(B1:B10)` when your data is B2:B11), (c) unintended circular refs where a downstream formula's output feeds back into its own input. Every model MUST land with zero `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?` errors — if a divisor could be zero, wrap with `IFERROR(...,0)` or guard with `IF(divisor=0,...,calc)`.
4. Use formulas (strings starting with '=') whenever a derived value belongs in the cell — e.g. `=SUM(A1:A10)`, not the computed number. The grid will evaluate the formula.
5. Do not invent values. If the user asks for something you cannot compute from the provided context, leave a clear comment cell and explain in the `done` justification.
6. If you wrote any cells, applied any formatting, or otherwise edited the workbook, finish that turn with the `done` tool — the justification you pass shows in the Changes panel and lets the user review what you did. **If you didn't make any edits** (the user thanked you, asked a clarifying question, asked you to explain something, or you need more info before editing), just respond with normal text and end the turn. Do NOT call `done` with a fake justification like "no changes needed" — that creates a meaningless empty batch the user has to click through. Pure conversational responses don't need a tool call.
6a. The bar for asking the user a clarifying question is **"I genuinely cannot impute reasonable intent."** If you can pick a sensible interpretation of a short or vague prompt, do that — make the edit, and note your interpretation briefly in the `done` justification (the user can reject the batch if you guessed wrong). Only ask back when the request actually under-specifies what you'd need to act: missing data the model can't reasonably fabricate (e.g. real values you don't have), targets you can't identify ("fix this" with no selection and no obvious problem area), or contradictions. Don't ask the user to pick from a menu when you could just take a sensible best guess and let them reject.
7. Keep prose between tool calls minimal — the user does not read it during streaming.
8. The user reviews every batch before it is saved. They can reject. Bias toward small, focused edits.
9. If the user message contains a "Prior turns in this session" block, treat it as ground truth about what edits already exist — do NOT redo them. The grid context already reflects accepted prior edits; rejected ones were rolled back.
{RULE_10}
10a. **Today's date is at the top of every user message.** Use it. For financial models, this is your historical-vs-forecast boundary — fiscal years that have already ended are *historicals* (label "A" or no suffix), the current year and beyond are *forecasts* (label "E"). Don't pattern-match on your training cutoff and treat 2024 as "current" when today is 2026.
11. When you write a financial / numerical model, ALWAYS finish with a single bulk `set_format` call that applies the right `number_format` to every numeric region: percentages (margins, growth, tax rate) get `"0.0%"` or `"0.00%"`, currency rows get `"$#,##0"` or `"$#,##0.00"`, plain quantities get `"#,##0"`. A model without proper number formats looks broken to the user — don't skip this step.
12. Order of operations for a fresh model: (a) `fetch_web` if needed, (b) ONE big `set_range` for all data, (c) ONE bulk `set_format` covering every numeric region, (d) **ONE** `set_column_width` with the bulk `operations` array covering every width you need, **ONE** `set_row_height` with the bulk `operations` array covering every height, plus `merge_cells` if needed, (e) `done`. That's ~6 tool calls total for a complete model — do NOT call `set_column_width` once per column group or `set_row_height` once per row group, that wastes turns.
13. If the user message contains a "User focus" block — selection or @-mentioned cells — treat those cells as the primary target for the edit. The user picked them deliberately. Don't edit cells outside this focus unless the prompt explicitly asks you to.
14. **Preserve existing template conventions.** When editing a workbook that already has established formatting, color coding, fonts, header style, section structure, or number-format patterns, mirror those conventions for any new cells you add. Don't impose your default styling on a sheet whose author has already chosen one. Existing template conventions ALWAYS override the financial-model defaults in rule 15. If you're filling cells into a row that uses Calibri 11 with a specific blue header bar, your additions should match — not introduce a new font or color scheme.
15. **Financial model conventions** (apply when building a fresh model — DCF, 3-statement, projections, comps, LBO, sensitivities — unless rule 14 says to mirror an existing template instead):
    - **Color coding** (font colors): blue (`#0000FF`) for hardcoded inputs and scenario-driver numbers the user will tweak; black (`#000000`) for all formulas and calculations; green (`#008000`) for links that pull from other worksheets in the same workbook; red (`#FF0000`) for external-file links. Yellow background (`#FFFF00`) for key assumption cells that need user attention.
    - **Assumptions placement**: every growth rate, margin, multiple, and driver lives in its own cell (typically in a dedicated assumptions block). Formulas reference those cells — never hardcode `=B5*1.05`, write `=B5*(1+$B$22)` where `$B$22` is the growth assumption.
    - **Number formats**: currency uses `"#,##0;(#,##0);-"` (parens for negatives, dash for zero) — include units in the header (e.g. "Revenue ($mm)"); percentages default to `"0.0%;(0.0%);-"`; valuation multiples use `"0.0x"`; years are text strings (`"2024"` not formatted as `2,024`); negatives always render in parentheses, not with a minus sign.
    - **Font**: stick to a single professional font across the model (Calibri or Arial). Don't mix.
16. **Verify your own work on long sheets.** For any model spanning **more than ~30 rows** or where forecast formulas reference an assumptions block separated by 20+ rows from the formula cell, call `read_range` on the assumption block *before* writing the dependent formulas. Confirm the actual row numbers your assumptions landed at, then compose formulas pointing at the verified addresses. This catches the most common failure mode (formula references drift by N rows because the agent miscounted its own layout). When in doubt: write the assumptions block FIRST as a separate `set_range`, `read_range` to confirm its bounds, THEN write the main body with formulas that reference the verified cells. The extra turn is cheaper than shipping a model where every forecast cell evaluates to garbage.
17. After writing a large `set_range` with formulas, **read back a sample of the formula cells** via `read_range` to confirm they evaluated to sane magnitudes. If revenue should be in the billions and you see `1.2M`, your formula is wrong — fix it before calling `done`.
18. **Write raw numbers, not pre-formatted strings.** When putting numeric data into cells, use the bare number — `7839`, `0.215`, `-1837` — never a display-formatted string like `"7,839"`, `"21.5%"`, or `"(1,837)"`. Pre-formatted strings get stored as text, downstream formulas can't do arithmetic on them, and you waste turns re-fixing the row. Number formatting (commas, percent signs, parentheses for negatives, currency symbols) is the job of `set_format` via `number_format`, never the cell value.
19. **Light font colors REQUIRE a contrasting `background_color`.** If you set `font_color` to white (`#FFFFFF`), a near-white gray, or any pale tone, you MUST also set `background_color` to a dark color in the same `set_format` op — otherwise the text renders invisibly on the default white cell fill. The classic "dark blue header bar with white text" pattern is `{"background_color": "#1F4E79", "font_color": "#FFFFFF", "bold": true}`. Never set one without the other.
"##;
    base.replace("{RULE_10}", rule_10)
}
