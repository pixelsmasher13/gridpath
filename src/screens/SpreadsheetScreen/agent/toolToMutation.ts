import type { UniverMutation } from "../types";

/**
 * Convert one Claude tool_use call into one or more UniverMutations.
 *
 * Tool args now use **A1 notation only** (`cell="A1"`, `top_left="B17"`).
 * We parse them here into the (row, col) 0-indexed pair Univer's facade
 * actually expects — Univer's API surface IS 0-indexed under the hood,
 * but the agent never sees that. Single source of truth at this seam.
 */
export type ToolResult =
  | { kind: "mutations"; mutations: UniverMutation[] }
  | { kind: "done"; justification: string; turn_summary?: string }
  | { kind: "fetch"; urls: string[] }
  | { kind: "read"; sheet: string; range: string }
  | {
      kind: "copy";
      sheet: string;
      source: string;
      dest_sheet: string;
      dest: string;
      mode: "all" | "values" | "formats";
    }
  | { kind: "read_reference"; workbook: string; sheet: string; range: string }
  | { kind: "describe_workbook"; sheet: string | null }
  | { kind: "find_rows"; query: string; sheet: string | null; max_results: number }
  | { kind: "script"; code: string }
  | {
      kind: "stage";
      sheet: string;
      title: string;
      source: string;
      units: string | null;
      columns: string[];
      rows: Array<{ label: string; values: unknown[] }>;
    }
  | { kind: "ignored"; reason: string };

/**
 * The model occasionally emits a nested-array tool arg as a JSON-encoded
 * STRING — e.g. `"values": "[[...]]"` instead of `"values": [[...]]` — a
 * common failure on large `set_range` payloads. When that happens the arg
 * looks like a plain string and our `Array.isArray` checks reject it as "bad
 * args", which the model then misdiagnoses (blaming em-dashes, parentheses…)
 * and burns turns flailing. Accept both shapes: if it's a string that parses
 * to an array, use the parsed array; otherwise return the value untouched so
 * the caller's own validation still rejects genuinely-bad input.
 */
function coerceArray(v: any): any {
  if (Array.isArray(v)) return v;
  if (typeof v === "string" && v.trim().startsWith("[")) {
    try {
      const parsed = JSON.parse(v);
      if (Array.isArray(parsed)) return parsed;
    } catch {
      /* not valid JSON — fall through, caller validates */
    }
  }
  return v;
}

/**
 * Parse a column list that may mix single letters and runs: "A", "B,AO",
 * "C:N", "A,C:F,H". Returns 0-indexed column numbers, deduped, in order.
 * Models emit the colon-run form constantly (it's Excel's own notation),
 * and the old comma-only parser silently mangled it.
 */
export function parseColumnList(raw: string): number[] {
  const lettersToCol = (letters: string): number => {
    let c = 0;
    for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
    return c - 1;
  };
  const out: number[] = [];
  for (const part of raw.split(",")) {
    const p = part.trim().toUpperCase();
    if (!p) continue;
    const m = /^([A-Z]+)(?::([A-Z]+))?$/.exec(p);
    if (!m) continue;
    const a = lettersToCol(m[1]);
    const b = m[2] ? lettersToCol(m[2]) : a;
    for (let c = Math.min(a, b); c <= Math.max(a, b); c++) out.push(c);
  }
  return [...new Set(out)].filter((n) => n >= 0);
}

/**
 * Parse a 1-indexed row list that may mix single rows and runs:
 * "1", "4,5,6", "4:8", "1,3:5". Returns 0-indexed rows, deduped.
 */
export function parseRowList(raw: string): number[] {
  const out: number[] = [];
  for (const part of raw.split(",")) {
    const p = part.trim();
    if (!p) continue;
    const m = /^(\d+)(?::(\d+))?$/.exec(p);
    if (!m) continue;
    const a = parseInt(m[1], 10);
    const b = m[2] ? parseInt(m[2], 10) : a;
    if (!Number.isFinite(a) || !Number.isFinite(b) || a < 1 || b < 1) continue;
    for (let r = Math.min(a, b); r <= Math.max(a, b); r++) out.push(r - 1);
  }
  return [...new Set(out)];
}

export function interpretToolCall(name: string, input: any): ToolResult {
  switch (name) {
    case "set_cell": {
      const sheet = String(input?.sheet ?? "");
      const cell = String(input?.cell ?? "");
      const parsed = parseA1(cell);
      if (!sheet || !parsed) {
        return { kind: "ignored", reason: `bad set_cell args: ${JSON.stringify(input)}` };
      }
      const formula =
        typeof input?.formula === "string" && input.formula.length > 0
          ? (input.formula.startsWith("=") ? input.formula : `=${input.formula}`)
          : null;
      const value = "value" in (input ?? {}) ? input.value : null;
      return {
        kind: "mutations",
        mutations: [
          {
            type: "set_cell",
            address: { sheet, row: parsed.row, col: parsed.col },
            old_value: null,
            new_value: formula ?? value ?? null,
            new_formula: formula,
          },
        ],
      };
    }
    case "set_range": {
      const sheet = String(input?.sheet ?? "");
      const topLeft = String(input?.top_left ?? "");
      const parsed = parseA1(topLeft);
      const values = coerceArray(input?.values);
      if (!sheet || !parsed || !Array.isArray(values)) {
        return {
          kind: "ignored",
          reason:
            `set_range: 'values' must be a 2D JSON array of rows (e.g. [["A",1],["B",2]]), ` +
            `got ${typeof input?.values}. Send it as an actual array, not a stringified one, ` +
            `and ensure 'sheet' and a valid A1 'top_left' are set.`,
        };
      }
      // Expand into individual set_cell mutations so the diff overlay can
      // tag each cell independently. set_range as a single mutation makes
      // partial-rejection awkward; per-cell mutations are the right grain.
      //
      // Treat `undefined`, `null`, AND empty string `""` as "preserve" —
      // skip them and leave whatever's already in the cell alone. This
      // matches the agent's intuitive use of these slots as "padding to
      // keep row alignment" rather than "explicitly clear the cell."
      // Previously empty strings landed as writes, which silently wiped
      // existing data whenever the agent shipped a row like
      // `["Gross Profit","","","","","","","","",""]` intending to set
      // only the label in column A. Agents that actually want to clear a
      // cell can use `clear_range` (explicit) or `set_cell` with value=null.
      const mutations: UniverMutation[] = [];
      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        if (!Array.isArray(row)) continue;
        for (let c = 0; c < row.length; c++) {
          const raw = row[c];
          if (raw === undefined || raw === null || raw === "") continue;
          const isFormula = typeof raw === "string" && raw.startsWith("=");
          mutations.push({
            type: "set_cell",
            address: { sheet, row: parsed.row + r, col: parsed.col + c },
            old_value: null,
            new_value: isFormula ? raw : raw,
            new_formula: isFormula ? raw : null,
          });
        }
      }
      return { kind: "mutations", mutations };
    }
    case "copy_range": {
      // Excel-like copy/paste. Only ARG validation happens here — the
      // expansion into per-cell mutations needs the live grid (source
      // values, formulas, formats), so it runs in the tool-call handler
      // via expandCopy().
      const sheet = String(input?.sheet ?? "");
      const source = String(input?.source ?? "").trim();
      const dest = String(input?.dest ?? "").trim();
      const destSheet = typeof input?.dest_sheet === "string" && input.dest_sheet.trim()
        ? input.dest_sheet.trim()
        : sheet;
      const modeRaw = typeof input?.mode === "string" ? input.mode : "all";
      const mode = modeRaw === "values" || modeRaw === "formats" ? modeRaw : "all";
      const corners = source.split(":");
      const srcOk =
        corners.length <= 2 && corners.every((p) => parseA1(p) !== null);
      if (!sheet || !srcOk || !parseA1(dest)) {
        return {
          kind: "ignored",
          reason:
            `copy_range: bad args — need sheet, an A1 'source' range (e.g. "BT5:BT120") ` +
            `and an A1 'dest' top-left cell (e.g. "BU5"). Got ${JSON.stringify(input)}`,
        };
      }
      return { kind: "copy", sheet, source, dest_sheet: destSheet, dest, mode };
    }
    case "set_format": {
      const sheet = String(input?.sheet ?? "");
      if (!sheet) {
        return { kind: "ignored", reason: `set_format: missing sheet` };
      }

      // Two input shapes: bulk { operations: [{range, format}] } or single
      // { range, format }. Both produce one FormatMutation per range so the
      // diff list shows each range as its own row.
      const ops: Array<{ range: string; format: any }> = [];
      const formatOps = coerceArray(input?.operations);
      if (Array.isArray(formatOps)) {
        for (const op of formatOps) {
          if (op && typeof op.range === "string" && typeof op.format === "object" && op.format) {
            ops.push({ range: String(op.range), format: op.format });
          }
        }
      } else if (typeof input?.range === "string" && typeof input?.format === "object" && input.format) {
        ops.push({ range: String(input.range), format: input.format });
      }

      if (ops.length === 0) {
        return { kind: "ignored", reason: `bad set_format args: ${JSON.stringify(input)}` };
      }

      const mutations: any[] = [];
      for (const { range, format } of ops) {
        const cells = expandA1Range(range);
        if (cells.length === 0) continue;
        mutations.push({
          type: "set_format",
          sheet,
          range,
          cells,
          old_format: [],
          new_format: format,
        });
      }
      if (mutations.length === 0) {
        return { kind: "ignored", reason: `set_format: no valid ranges` };
      }
      return { kind: "mutations", mutations };
    }
    case "set_column_width": {
      const sheet = String(input?.sheet ?? "");
      if (!sheet) {
        return { kind: "ignored", reason: `bad set_column_width args: ${JSON.stringify(input)}` };
      }
      // Accept either bulk `operations: [{columns, width}]` or the flat
      // single-op `{columns, width}` shape. Normalize to an array of ops.
      const widthOps = coerceArray(input?.operations);
      const ops: Array<{ columns: string; width: number }> =
        Array.isArray(widthOps)
          ? widthOps.map((o: any) => ({ columns: String(o?.columns ?? ""), width: Number(o?.width) }))
          : [{ columns: String(input?.columns ?? ""), width: Number(input?.width) }];

      const muts: any[] = [];
      for (const op of ops) {
        if (!op.columns || !Number.isFinite(op.width) || op.width <= 0) continue;
        const columns = parseColumnList(op.columns);
        if (columns.length === 0) continue;
        muts.push({
          type: "set_column_width",
          sheet,
          columns,
          old_widths: [],
          new_width: op.width,
        });
      }
      if (muts.length === 0) {
        return { kind: "ignored", reason: `set_column_width: no valid ops in ${JSON.stringify(input)}` };
      }
      return { kind: "mutations", mutations: muts };
    }
    case "set_row_height": {
      const sheet = String(input?.sheet ?? "");
      if (!sheet) {
        return { kind: "ignored", reason: `bad set_row_height args: ${JSON.stringify(input)}` };
      }
      const heightOps = coerceArray(input?.operations);
      const ops: Array<{ rows: string; height: number }> =
        Array.isArray(heightOps)
          ? heightOps.map((o: any) => ({ rows: String(o?.rows ?? ""), height: Number(o?.height) }))
          : [{ rows: String(input?.rows ?? ""), height: Number(input?.height) }];

      const muts: any[] = [];
      for (const op of ops) {
        if (!op.rows || !Number.isFinite(op.height) || op.height <= 0) continue;
        const rows = parseRowList(op.rows);
        if (rows.length === 0) continue;
        muts.push({
          type: "set_row_height",
          sheet,
          rows,
          old_heights: [],
          new_height: op.height,
        });
      }
      if (muts.length === 0) {
        return { kind: "ignored", reason: `set_row_height: no valid ops in ${JSON.stringify(input)}` };
      }
      return { kind: "mutations", mutations: muts };
    }
    case "merge_cells":
    case "unmerge_cells": {
      const sheet = String(input?.sheet ?? "");
      const range = String(input?.range ?? "");
      const cells = expandA1Range(range);
      if (!sheet || cells.length === 0) {
        return { kind: "ignored", reason: `bad ${name} args: ${JSON.stringify(input)}` };
      }
      const rows = cells.map((c) => c.row);
      const cols = cells.map((c) => c.col);
      return {
        kind: "mutations",
        mutations: [{
          type: name,
          sheet,
          range,
          start_row: Math.min(...rows),
          start_col: Math.min(...cols),
          end_row: Math.max(...rows),
          end_col: Math.max(...cols),
        } as any],
      };
    }
    case "set_note": {
      const sheet = String(input?.sheet ?? "");
      const parsed = parseA1(String(input?.cell ?? ""));
      const text = typeof input?.text === "string" ? input.text : "";
      if (!sheet || !parsed || !text) {
        return { kind: "ignored", reason: `bad set_note args: ${JSON.stringify(input)}` };
      }
      return {
        kind: "mutations",
        mutations: [{ type: "set_note", sheet, row: parsed.row, col: parsed.col, text } as any],
      };
    }
    case "delete_note": {
      const sheet = String(input?.sheet ?? "");
      const parsed = parseA1(String(input?.cell ?? ""));
      if (!sheet || !parsed) {
        return { kind: "ignored", reason: `bad delete_note args: ${JSON.stringify(input)}` };
      }
      return {
        kind: "mutations",
        mutations: [{ type: "delete_note", sheet, row: parsed.row, col: parsed.col } as any],
      };
    }
    case "stage_data": {
      const sheet = String(input?.sheet ?? "").trim() || "Data";
      const title = String(input?.title ?? "").trim();
      const source = String(input?.source ?? "").trim();
      const units = typeof input?.units === "string" && input.units.trim() ? input.units.trim() : null;
      const columnsRaw = coerceArray(input?.columns);
      const columns: string[] = Array.isArray(columnsRaw) ? columnsRaw.map((c: unknown) => String(c)) : [];
      const rawRows = coerceArray(input?.rows);
      if (!title) return { kind: "ignored", reason: "stage_data: 'title' is required (what is this block?)" };
      if (!source) {
        return { kind: "ignored", reason: "stage_data: 'source' is required — cite where these figures come from (URL or filing name)." };
      }
      if (columns.length === 0 || columns.length > 24) {
        return { kind: "ignored", reason: `stage_data: 'columns' must have 1–24 period/header labels, got ${columns.length}` };
      }
      if (!Array.isArray(rawRows) || rawRows.length === 0 || rawRows.length > 80) {
        return { kind: "ignored", reason: `stage_data: 'rows' must have 1–80 entries, got ${Array.isArray(rawRows) ? rawRows.length : typeof input?.rows}` };
      }
      const rows: Array<{ label: string; values: unknown[] }> = [];
      for (const r of rawRows) {
        // Accept {label, values:[...]} or the shorthand [label, v1, v2, ...].
        if (Array.isArray(r) && r.length > 0) {
          rows.push({ label: String(r[0]), values: r.slice(1) });
        } else if (r && typeof r === "object" && typeof (r as any).label === "string") {
          const values = coerceArray((r as any).values) ?? [];
          rows.push({ label: (r as any).label, values: Array.isArray(values) ? values : [] });
        } else {
          return { kind: "ignored", reason: `stage_data: each row must be {label, values:[...]} or [label, v1, ...], got ${JSON.stringify(r).slice(0, 80)}` };
        }
        const last = rows[rows.length - 1];
        if (last.values.length > columns.length) {
          return { kind: "ignored", reason: `stage_data: row ${JSON.stringify(last.label)} has ${last.values.length} values but only ${columns.length} columns` };
        }
      }
      return { kind: "stage", sheet, title, source, units, columns, rows };
    }
    case "create_sheet": {
      const sheetName = String(input?.name ?? "");
      if (!sheetName) return { kind: "ignored", reason: "create_sheet: missing name" };
      const tab_color = typeof input?.tab_color === "string" ? input.tab_color : null;
      return {
        kind: "mutations",
        mutations: [{ type: "create_sheet", name: sheetName, tab_color } as any],
      };
    }
    case "delete_sheet": {
      const sheetName = String(input?.name ?? "");
      if (!sheetName) return { kind: "ignored", reason: "delete_sheet: missing name" };
      return { kind: "mutations", mutations: [{ type: "delete_sheet", name: sheetName } as any] };
    }
    case "rename_sheet": {
      const oldName = String(input?.old_name ?? "");
      const newName = String(input?.new_name ?? "");
      if (!oldName || !newName) return { kind: "ignored", reason: "rename_sheet: missing names" };
      return {
        kind: "mutations",
        mutations: [{ type: "rename_sheet", old_name: oldName, new_name: newName } as any],
      };
    }
    case "clear_range": {
      const sheetName = String(input?.sheet ?? "");
      const rangeStr = String(input?.range ?? "");
      const cells = expandA1Range(rangeStr);
      if (!sheetName || cells.length === 0) {
        return { kind: "ignored", reason: `clear_range: bad args ${JSON.stringify(input)}` };
      }
      return {
        kind: "mutations",
        mutations: [{
          type: "clear_range",
          sheet: sheetName,
          range: rangeStr,
          cells: cells.map((c) => ({ row: c.row, col: c.col, old_value: null, old_formula: null })),
        } as any],
      };
    }
    case "insert_rows":
    case "delete_rows": {
      const sheetName = String(input?.sheet ?? "");
      const count = Math.max(1, Number(input?.count ?? 1));
      if (name === "insert_rows") {
        const before = Number(input?.before);
        if (!sheetName || !Number.isFinite(before) || before < 1) {
          return { kind: "ignored", reason: `insert_rows: bad args ${JSON.stringify(input)}` };
        }
        return {
          kind: "mutations",
          mutations: [{ type: "insert_rows", sheet: sheetName, before: before - 1, count } as any],
        };
      } else {
        const start = Number(input?.start);
        if (!sheetName || !Number.isFinite(start) || start < 1) {
          return { kind: "ignored", reason: `delete_rows: bad args ${JSON.stringify(input)}` };
        }
        return {
          kind: "mutations",
          mutations: [{ type: "delete_rows", sheet: sheetName, start: start - 1, count } as any],
        };
      }
    }
    case "insert_columns":
    case "delete_columns": {
      const sheetName = String(input?.sheet ?? "");
      const count = Math.max(1, Number(input?.count ?? 1));
      const lettersToCol = (letters: string): number => {
        let n = 0;
        for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n - 1;
      };
      if (name === "insert_columns") {
        const beforeLetters = String(input?.before ?? "");
        if (!sheetName || !/^[A-Za-z]+$/.test(beforeLetters)) {
          return { kind: "ignored", reason: `insert_columns: bad args ${JSON.stringify(input)}` };
        }
        return {
          kind: "mutations",
          mutations: [{
            type: "insert_columns",
            sheet: sheetName,
            before: lettersToCol(beforeLetters),
            count,
          } as any],
        };
      } else {
        const startLetters = String(input?.start ?? "");
        if (!sheetName || !/^[A-Za-z]+$/.test(startLetters)) {
          return { kind: "ignored", reason: `delete_columns: bad args ${JSON.stringify(input)}` };
        }
        return {
          kind: "mutations",
          mutations: [{
            type: "delete_columns",
            sheet: sheetName,
            start: lettersToCol(startLetters),
            count,
          } as any],
        };
      }
    }
    case "freeze_panes": {
      const sheetName = String(input?.sheet ?? "");
      const freeze_rows = Math.max(0, Number(input?.freeze_rows ?? 0));
      const freeze_cols = Math.max(0, Number(input?.freeze_cols ?? 0));
      if (!sheetName) return { kind: "ignored", reason: "freeze_panes: missing sheet" };
      return {
        kind: "mutations",
        mutations: [{ type: "freeze_panes", sheet: sheetName, freeze_rows, freeze_cols } as any],
      };
    }
    case "unfreeze_panes": {
      const sheetName = String(input?.sheet ?? "");
      if (!sheetName) return { kind: "ignored", reason: "unfreeze_panes: missing sheet" };
      return {
        kind: "mutations",
        mutations: [{ type: "unfreeze_panes", sheet: sheetName } as any],
      };
    }
    case "hide_rows":
    case "show_rows": {
      const sheetName = String(input?.sheet ?? "");
      const rowsRaw = String(input?.rows ?? "");
      const rows = parseRowList(rowsRaw);
      if (!sheetName || rows.length === 0) {
        return { kind: "ignored", reason: `${name}: no valid rows in "${rowsRaw}"` };
      }
      return {
        kind: "mutations",
        mutations: [{ type: name, sheet: sheetName, rows } as any],
      };
    }
    case "hide_columns":
    case "show_columns": {
      const sheetName = String(input?.sheet ?? "");
      const colsRaw = String(input?.columns ?? "");
      const columns = parseColumnList(colsRaw);
      if (!sheetName || columns.length === 0) {
        return { kind: "ignored", reason: `${name}: no valid columns in "${colsRaw}"` };
      }
      return {
        kind: "mutations",
        mutations: [{ type: name, sheet: sheetName, columns } as any],
      };
    }
    case "define_name": {
      const nm = String(input?.name ?? "").trim();
      let ref = String(input?.ref ?? "").trim();
      if (!nm || !ref) {
        return { kind: "ignored", reason: `define_name: missing name or ref ${JSON.stringify(input)}` };
      }
      // A defined name must be anchored to a sheet. Accept either a
      // sheet-qualified ref ("Model!B5:B12") or a separate `sheet` field
      // plus a bare range ("B5:B12").
      if (!ref.includes("!")) {
        const sheet = String(input?.sheet ?? "").trim();
        if (!sheet) {
          return {
            kind: "ignored",
            reason: `define_name: ref must be sheet-qualified, e.g. "Model!B5:B12", or pass a separate 'sheet' field. Got "${ref}".`,
          };
        }
        ref = `${sheet}!${ref}`;
      }
      const normalized = normalizeDefinedNameRef(ref);
      if (!normalized) {
        return { kind: "ignored", reason: `define_name: could not parse ref "${ref}"` };
      }
      return {
        kind: "mutations",
        mutations: [{ type: "define_name", name: nm, ref: normalized, old_ref: null } as any],
      };
    }
    case "done": {
      const justification = String(input?.justification ?? "");
      const turn_summary =
        typeof input?.turn_summary === "string" && input.turn_summary.trim()
          ? input.turn_summary.trim()
          : undefined;
      return { kind: "done", justification, turn_summary };
    }
    case "fetch_web": {
      // Rust handles the actual fetch — frontend just renders an info chip
      // in the chat so the user sees what the agent looked up.
      const urlsIn = input?.urls;
      const urls = Array.isArray(urlsIn) ? urlsIn.filter((u) => typeof u === "string") : [];
      return { kind: "fetch", urls };
    }
    case "read_range": {
      const sheet = String(input?.sheet ?? "");
      const range = String(input?.range ?? "");
      if (!sheet || !range) {
        return { kind: "ignored", reason: `bad read_range args: ${JSON.stringify(input)}` };
      }
      return { kind: "read", sheet, range };
    }
    case "describe_workbook": {
      const sheet =
        typeof input?.sheet === "string" && input.sheet.trim() ? input.sheet.trim() : null;
      return { kind: "describe_workbook", sheet };
    }
    case "find_rows": {
      const query = String(input?.query ?? "").trim();
      if (!query) {
        return { kind: "ignored", reason: `find_rows: missing query` };
      }
      const sheet =
        typeof input?.sheet === "string" && input.sheet.trim() ? input.sheet.trim() : null;
      const rawMax = Number(input?.max_results);
      const max_results = Number.isFinite(rawMax) ? Math.min(50, Math.max(1, Math.trunc(rawMax))) : 20;
      return { kind: "find_rows", query, sheet, max_results };
    }
    case "run_script": {
      // Execution needs the live grid (snapshot for the read model), so it
      // happens in the tool-call handler via executeSheetScript() — same
      // split as copy_range. Only arg validation lives here.
      const code = String(input?.script ?? "");
      if (!code.trim()) {
        return { kind: "ignored", reason: "run_script: missing or empty 'script'" };
      }
      return { kind: "script", code };
    }
    case "read_reference": {
      const workbook = String(input?.workbook ?? "");
      const sheet = String(input?.sheet ?? "");
      const range = String(input?.range ?? "");
      if (!workbook || !sheet || !range) {
        return { kind: "ignored", reason: `bad read_reference args: ${JSON.stringify(input)}` };
      }
      return { kind: "read_reference", workbook, sheet, range };
    }
    default:
      return { kind: "ignored", reason: `unknown tool: ${name}` };
  }
}

/**
 * Expand an A1 range string into a list of 0-indexed cells. Accepts a
 * single-cell address ("A1") OR a rectangular range ("A1:C3"). Returns
 * an empty array on malformed input.
 */
export function expandA1Range(rangeStr: string): Array<{ row: number; col: number }> {
  const parts = rangeStr.split(":");
  if (parts.length === 1) {
    const p = parseA1(parts[0]);
    return p ? [p] : [];
  }
  if (parts.length === 2) {
    const a = parseA1(parts[0]);
    const b = parseA1(parts[1]);
    if (!a || !b) return [];
    const r0 = Math.min(a.row, b.row);
    const r1 = Math.max(a.row, b.row);
    const c0 = Math.min(a.col, b.col);
    const c1 = Math.max(a.col, b.col);
    const out: Array<{ row: number; col: number }> = [];
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) out.push({ row: r, col: c });
    }
    return out;
  }
  return [];
}

/**
 * Normalize a sheet-qualified defined-name target into the canonical absolute
 * form both Univer (`insertDefinedName`) and ExcelJS (`definedNames.add`)
 * expect — e.g. `Model!B5:B12` → `Model!$B$5:$B$12`, and quote sheet names
 * that contain spaces or other non-identifier characters. Returns null if the
 * range portion can't be parsed as A1.
 */
export function normalizeDefinedNameRef(ref: string): string | null {
  const bang = ref.lastIndexOf("!");
  if (bang < 0) return null;
  let sheet = ref.slice(0, bang).trim();
  const rangePart = ref.slice(bang + 1).trim();
  // Strip any quoting the caller already added so we re-quote consistently.
  if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  if (!sheet) return null;

  const abs = (cell: string): string | null => {
    const m = /^\s*\$?([A-Za-z]+)\$?(\d+)\s*$/.exec(cell);
    if (!m) return null;
    return `$${m[1].toUpperCase()}$${m[2]}`;
  };

  const parts = rangePart.split(":");
  let absRange: string;
  if (parts.length === 1) {
    const a = abs(parts[0]);
    if (!a) return null;
    absRange = a;
  } else if (parts.length === 2) {
    const a = abs(parts[0]);
    const b = abs(parts[1]);
    if (!a || !b) return null;
    absRange = `${a}:${b}`;
  } else {
    return null;
  }

  const needsQuote = !/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheet);
  const sheetRef = needsQuote ? `'${sheet.replace(/'/g, "''")}'` : sheet;
  return `${sheetRef}!${absRange}`;
}

/**
 * Parse an Excel A1 address like "A1", "B17", "AA42" (case-insensitive)
 * into 0-indexed { row, col }. Returns null on malformed input.
 */
export function parseA1(addr: string): { row: number; col: number } | null {
  const m = /^\s*([A-Za-z]+)\s*(\d+)\s*$/.exec(addr);
  if (!m) return null;
  const colLetters = m[1].toUpperCase();
  const row1 = parseInt(m[2], 10);
  if (!Number.isFinite(row1) || row1 < 1) return null;
  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { row: row1 - 1, col: col - 1 };
}
