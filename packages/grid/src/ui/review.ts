/**
 * Plain-language descriptions of an agent's changes, for the review panel:
 * a title per batch ("Write cells · A1:B3") and a few detail lines
 * ("B4: empty → =SUM(B2:B3)").
 */
import type { AgentBatch } from "../ai/toolkit";
import type { CellFormat, UniverMutation } from "../../../gridpath/src/core/types";

export function colName(c: number): string {
  let s = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export const a1 = (row: number, col: number) => `${colName(col)}${row + 1}`;

const TOOL_LABELS: Record<string, string> = {
  set_cell: "Edit cell",
  set_range: "Write cells",
  set_format: "Format",
  clear_range: "Clear",
  copy_range: "Copy",
  run_script: "Script",
  insert_rows: "Insert rows",
  delete_rows: "Delete rows",
  insert_columns: "Insert columns",
  delete_columns: "Delete columns",
  create_sheet: "New sheet",
  delete_sheet: "Delete sheet",
  rename_sheet: "Rename sheet",
  freeze_panes: "Freeze panes",
  merge_cells: "Merge",
  set_column_width: "Column width",
  set_row_height: "Row height",
  define_name: "Define name",
  set_note: "Note",
};

export function toolLabel(tool: string): string {
  return TOOL_LABELS[tool] ?? tool.replace(/_/g, " ").replace(/^./, (c) => c.toUpperCase());
}

/** The range the batch's cells span on their first sheet ("Sheet1!A1:B3"), and how many sheets. */
export function cellSpan(cells: AgentBatch["cells"]): { label: string; first: { sheet: string; row: number; col: number } | null } {
  if (!cells.length) return { label: "", first: null };
  const sheet = cells[0].sheet;
  const own = cells.filter((c) => c.sheet === sheet);
  const r0 = Math.min(...own.map((c) => c.row)), r1 = Math.max(...own.map((c) => c.row));
  const c0 = Math.min(...own.map((c) => c.col)), c1 = Math.max(...own.map((c) => c.col));
  const range = r0 === r1 && c0 === c1 ? a1(r0, c0) : `${a1(r0, c0)}:${a1(r1, c1)}`;
  const others = new Set(cells.map((c) => c.sheet)).size - 1;
  return { label: `${sheet}!${range}${others ? ` + ${others} more sheet${others === 1 ? "" : "s"}` : ""}`, first: { sheet, row: own[0].row, col: own[0].col } };
}

function shown(v: unknown, formula?: string | null): string {
  if (formula) return formula.startsWith("=") ? formula : `=${formula}`;
  if (v === null || v === undefined || v === "") return "empty";
  const s = typeof v === "string" ? v : String(v);
  return s.length > 40 ? `${s.slice(0, 39)}…` : s;
}

function formatWords(f: CellFormat): string {
  const words: string[] = [];
  if (f.bold !== undefined) words.push(f.bold ? "bold" : "not bold");
  if (f.italic !== undefined) words.push(f.italic ? "italic" : "not italic");
  if (f.underline !== undefined) words.push(f.underline ? "underline" : "no underline");
  if (f.strike !== undefined) words.push(f.strike ? "strikethrough" : "no strikethrough");
  if (f.font_color) words.push(`text ${f.font_color}`);
  if (f.background_color !== undefined) words.push(f.background_color ? `fill ${f.background_color}` : "no fill");
  if (f.font_size) words.push(`${f.font_size} pt`);
  if (f.font_family) words.push(f.font_family);
  if (f.horizontal_align) words.push(`align ${f.horizontal_align}`);
  if (f.vertical_align) words.push(`align ${f.vertical_align}`);
  if (f.wrap_text !== undefined) words.push(f.wrap_text ? "wrap" : "no wrap");
  if (f.number_format) words.push(`format ${f.number_format}`);
  return words.join(", ") || "format";
}

const plural = (n: number, w: string) => `${n} ${w}${n === 1 ? "" : "s"}`;

/** Detail lines for one mutation (cell edits list each cell). */
function mutationLines(m: UniverMutation): string[] {
  switch (m.type) {
    case "set_cell":
      return [`${a1(m.address.row, m.address.col)}: ${shown(m.old_value, m.old_formula)} → ${shown(m.new_value, m.new_formula)}`];
    case "set_range": {
      const out: string[] = [];
      m.values.forEach((row, i) =>
        row.forEach((v, j) => out.push(`${a1(m.start_row + i, m.start_col + j)}: ${shown(m.old_values?.[i]?.[j])} → ${shown(v)}`)),
      );
      return out;
    }
    case "set_format":
      return [`${m.range}: ${formatWords(m.new_format)}`];
    case "clear_range":
      return [`${m.range}: cleared (${plural(m.cells.length, "cell")} had content)`];
    case "insert_rows":
      return [`${plural(m.count, "row")} inserted before row ${m.before + 1}`];
    case "delete_rows":
      return [`row${m.count === 1 ? "" : "s"} ${m.start + 1}${m.count > 1 ? `–${m.start + m.count}` : ""} deleted`];
    case "insert_columns":
      return [`${plural(m.count, "column")} inserted before ${colName(m.before)}`];
    case "delete_columns":
      return [`column${m.count === 1 ? "" : "s"} ${colName(m.start)}${m.count > 1 ? `–${colName(m.start + m.count - 1)}` : ""} deleted`];
    case "create_sheet":
      return [`sheet "${m.name}" added`];
    case "delete_sheet":
      return [`sheet "${m.name}" deleted`];
    case "rename_sheet":
      return [`sheet "${m.old_name}" renamed "${m.new_name}"`];
    case "freeze_panes":
      return [`freeze ${plural(m.freeze_rows, "row")}, ${plural(m.freeze_cols, "column")}`];
    case "unfreeze_panes":
      return ["unfreeze panes"];
    case "merge_cells":
      return [`${m.range}: merged`];
    case "unmerge_cells":
      return [`${m.range}: unmerged`];
    case "set_column_width":
      return [`column${m.columns.length === 1 ? "" : "s"} ${m.columns.map(colName).join(", ")}: width ${m.new_width}`];
    case "set_row_height":
      return [`row${m.rows.length === 1 ? "" : "s"} ${m.rows.map((r) => r + 1).join(", ")}: height ${m.new_height}`];
    case "hide_rows":
    case "show_rows":
      return [`row${m.rows.length === 1 ? "" : "s"} ${m.rows.map((r) => r + 1).join(", ")} ${m.type === "hide_rows" ? "hidden" : "shown"}`];
    case "hide_columns":
    case "show_columns":
      return [`column${m.columns.length === 1 ? "" : "s"} ${m.columns.map(colName).join(", ")} ${m.type === "hide_columns" ? "hidden" : "shown"}`];
    case "define_name":
      return [`name ${m.name} → ${m.ref}`];
    case "set_note":
      return [`${a1(m.row, m.col)}: note "${shown(m.text)}"`];
    case "delete_note":
      return [`${a1(m.row, m.col)}: note removed`];
    default:
      return [];
  }
}

/** A batch for the panel: title, where, and up to `max` detail lines. */
export function describeBatch(batch: AgentBatch, max = 4): { title: string; where: string; lines: string[]; more: number } {
  const all = batch.mutations.flatMap(mutationLines);
  const span = cellSpan(batch.cells);
  const count = batch.cells.length;
  const title = `${toolLabel(batch.tool)}${count > 1 ? ` · ${plural(count, "cell")}` : ""}`;
  return { title, where: span.label, lines: all.slice(0, max), more: Math.max(0, all.length - max) };
}
