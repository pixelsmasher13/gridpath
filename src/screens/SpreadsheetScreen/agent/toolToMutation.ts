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
  | { kind: "done"; justification: string }
  | { kind: "fetch"; urls: string[] }
  | { kind: "read"; sheet: string; range: string }
  | { kind: "ignored"; reason: string };

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
      const values = input?.values;
      if (!sheet || !parsed || !Array.isArray(values)) {
        return { kind: "ignored", reason: `bad set_range args: ${JSON.stringify(input)}` };
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
    case "set_format": {
      const sheet = String(input?.sheet ?? "");
      if (!sheet) {
        return { kind: "ignored", reason: `set_format: missing sheet` };
      }

      // Two input shapes: bulk { operations: [{range, format}] } or single
      // { range, format }. Both produce one FormatMutation per range so the
      // diff list shows each range as its own row.
      const ops: Array<{ range: string; format: any }> = [];
      if (Array.isArray(input?.operations)) {
        for (const op of input.operations) {
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
      const ops: Array<{ columns: string; width: number }> =
        Array.isArray(input?.operations)
          ? input.operations.map((o: any) => ({ columns: String(o?.columns ?? ""), width: Number(o?.width) }))
          : [{ columns: String(input?.columns ?? ""), width: Number(input?.width) }];

      const muts: any[] = [];
      for (const op of ops) {
        if (!op.columns || !Number.isFinite(op.width) || op.width <= 0) continue;
        const columns = op.columns
          .split(",")
          .map((s) => s.trim().toUpperCase())
          .filter(Boolean)
          .map((letters) => {
            let c = 0;
            for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
            return c - 1;
          })
          .filter((n) => n >= 0);
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
      const ops: Array<{ rows: string; height: number }> =
        Array.isArray(input?.operations)
          ? input.operations.map((o: any) => ({ rows: String(o?.rows ?? ""), height: Number(o?.height) }))
          : [{ rows: String(input?.rows ?? ""), height: Number(input?.height) }];

      const muts: any[] = [];
      for (const op of ops) {
        if (!op.rows || !Number.isFinite(op.height) || op.height <= 0) continue;
        const rows = op.rows
          .split(",")
          .map((s) => parseInt(s.trim(), 10))
          .filter((n) => Number.isFinite(n) && n >= 1)
          .map((n) => n - 1);
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
      const rows = rowsRaw
        .split(",")
        .map((s) => parseInt(s.trim(), 10))
        .filter((n) => Number.isFinite(n) && n >= 1)
        .map((n) => n - 1);
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
      const columns = colsRaw
        .split(",")
        .map((s) => s.trim().toUpperCase())
        .filter(Boolean)
        .map((letters) => {
          let c = 0;
          for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
          return c - 1;
        })
        .filter((n) => n >= 0);
      if (!sheetName || columns.length === 0) {
        return { kind: "ignored", reason: `${name}: no valid columns in "${colsRaw}"` };
      }
      return {
        kind: "mutations",
        mutations: [{ type: name, sheet: sheetName, columns } as any],
      };
    }
    case "done": {
      const justification = String(input?.justification ?? "");
      return { kind: "done", justification };
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
