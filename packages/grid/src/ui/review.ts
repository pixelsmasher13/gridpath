/**
 * Plain-language descriptions of an agent's changes, for the review panel.
 * Tool calls that touch the same cells one after another (a write, then the
 * formatting of what it wrote) are one change; each change says where it
 * is, what it did ("8 cells added · italic, percent") and shows a few of
 * the cells it wrote.
 */
import type { AgentBatch } from "../ai/toolkit";
import type { CellFormat, UniverMutation } from "../../../gridpath/src/core/types";

export function colName(c: number): string {
  let s = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export const a1 = (row: number, col: number) => `${colName(col)}${row + 1}`;

/** Where some cells are: their first sheet, the range they span there,
 * how many other sheets they touch, and the first cell (to jump to). */
export function cellSpan(cells: AgentBatch["cells"]): { sheet: string; range: string; others: number; first: { sheet: string; row: number; col: number } | null } {
  if (!cells.length) return { sheet: "", range: "", others: 0, first: null };
  const sheet = cells[0].sheet;
  const own = cells.filter((c) => c.sheet === sheet);
  const r0 = Math.min(...own.map((c) => c.row)), r1 = Math.max(...own.map((c) => c.row));
  const c0 = Math.min(...own.map((c) => c.col)), c1 = Math.max(...own.map((c) => c.col));
  const range = r0 === r1 && c0 === c1 ? a1(r0, c0) : `${a1(r0, c0)}:${a1(r1, c1)}`;
  const first = own.reduce((a, c) => (c.row < a.row || (c.row === a.row && c.col < a.col) ? c : a), own[0]);
  return { sheet, range, others: new Set(cells.map((c) => c.sheet)).size - 1, first: { sheet, row: first.row, col: first.col } };
}

/** A number format as a word: percent, currency, date or number. */
function formatKind(code: string): string {
  const bare = code.replace(/"[^"]*"/g, "").replace(/\\./g, "");
  if (bare.includes("%")) return "percent";
  if (/[$€£¥]/.test(code)) return "currency";
  if (/[dmyhs]/i.test(bare) && !/[0#]/.test(bare)) return "date";
  return "number format";
}

/** What a format change does, in words ("italic, text color, percent"). */
function formatWords(f: CellFormat): string[] {
  const words: string[] = [];
  if (f.bold !== undefined) words.push(f.bold ? "bold" : "not bold");
  if (f.italic !== undefined) words.push(f.italic ? "italic" : "not italic");
  if (f.underline !== undefined) words.push(f.underline ? "underlined" : "not underlined");
  if (f.strike !== undefined) words.push(f.strike ? "strikethrough" : "no strikethrough");
  if (f.font_color) words.push("text color");
  if (f.background_color !== undefined) words.push(f.background_color ? "fill" : "no fill");
  if (f.font_size || f.font_family) words.push("font");
  if (f.horizontal_align) words.push(`aligned ${f.horizontal_align}`);
  if (f.wrap_text !== undefined) words.push(f.wrap_text ? "wrapped" : "not wrapped");
  if (f.number_format) words.push(formatKind(f.number_format));
  return words;
}

const plural = (n: number, w: string) => `${n} ${w}${n === 1 ? "" : "s"}`;

/** Lines for the changes that aren't cell writes or formatting (rows, sheets, names…). */
function otherLines(m: UniverMutation): string[] {
  switch (m.type) {
    case "clear_range":
      return [`${m.range} cleared`];
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
      return [`froze ${plural(m.freeze_rows, "row")} and ${plural(m.freeze_cols, "column")}`];
    case "unfreeze_panes":
      return ["unfroze panes"];
    case "merge_cells":
      return [`${m.range} merged`];
    case "unmerge_cells":
      return [`${m.range} unmerged`];
    case "set_column_width":
      return [`column${m.columns.length === 1 ? "" : "s"} ${m.columns.map(colName).join(", ")} resized`];
    case "set_row_height":
      return [`row${m.rows.length === 1 ? "" : "s"} ${m.rows.map((r) => r + 1).join(", ")} resized`];
    case "hide_rows":
    case "show_rows":
      return [`row${m.rows.length === 1 ? "" : "s"} ${m.rows.map((r) => r + 1).join(", ")} ${m.type === "hide_rows" ? "hidden" : "shown"}`];
    case "hide_columns":
    case "show_columns":
      return [`column${m.columns.length === 1 ? "" : "s"} ${m.columns.map(colName).join(", ")} ${m.type === "hide_columns" ? "hidden" : "shown"}`];
    case "define_name":
      return [`name ${m.name} defined as ${m.ref}`];
    case "set_note":
      return [`note added to ${a1(m.row, m.col)}`];
    case "delete_note":
      return [`note removed from ${a1(m.row, m.col)}`];
    default:
      return [];
  }
}

/** Per-sheet bounding boxes of some cells. */
function boxes(cells: AgentBatch["cells"]): Map<string, [number, number, number, number]> {
  const out = new Map<string, [number, number, number, number]>();
  for (const c of cells) {
    const b = out.get(c.sheet);
    out.set(c.sheet, b ? [Math.min(b[0], c.row), Math.max(b[1], c.row), Math.min(b[2], c.col), Math.max(b[3], c.col)] : [c.row, c.row, c.col, c.col]);
  }
  return out;
}

function overlaps(a: AgentBatch["cells"], b: AgentBatch["cells"]): boolean {
  const bb = boxes(b);
  for (const [sheet, [r0, r1, c0, c1]] of boxes(a)) {
    const o = bb.get(sheet);
    if (o && r0 <= o[1] && o[0] <= r1 && c0 <= o[3] && o[2] <= c1) return true;
  }
  return false;
}

/** One change to review: consecutive pending batches over the same cells. */
export type ChangeGroup = { batches: AgentBatch[]; cells: AgentBatch["cells"] };

/** The pending batches as changes to review, in order. A batch joins the
 * change before it when it touches the same cells and nothing was accepted
 * in between. */
export function groupChanges(batches: AgentBatch[]): ChangeGroup[] {
  const groups: ChangeGroup[] = [];
  let open: ChangeGroup | null = null;
  for (const b of batches) {
    if (b.status !== "pending") {
      open = null;
      continue;
    }
    if (open && b.cells.length && overlaps(open.cells, b.cells)) {
      open.batches.push(b);
      open.cells = [...open.cells, ...b.cells];
    } else {
      open = { batches: [b], cells: [...b.cells] };
      groups.push(open);
    }
  }
  return groups;
}

export type CellLine = { ref: string; value: string; old: string | null; formula: boolean };

function written(v: unknown, formula?: string | null): { text: string; formula: boolean } {
  if (formula) return { text: formula.startsWith("=") ? formula : `=${formula}`, formula: true };
  if (v === null || v === undefined || v === "") return { text: "", formula: false };
  const s = typeof v === "string" ? v : String(v);
  return { text: s, formula: s.startsWith("=") };
}

/** A change for the panel: what it did in a few words, the cells it wrote
 * (up to `max`), and its other effects (rows, sheets, names…). */
export function describeChange(group: ChangeGroup, max = 4): { summary: string; cells: CellLine[]; more: number; other: string[] } {
  const writes = new Map<string, { row: number; col: number; value: string; old: string; formula: boolean }>();
  const formats: string[] = [];
  const other: string[] = [];
  const put = (row: number, col: number, now: { text: string; formula: boolean }, before: { text: string }) => {
    const key = `${row},${col}`;
    const prev = writes.get(key);
    writes.set(key, { row, col, value: now.text, old: prev ? prev.old : before.text, formula: now.formula });
  };
  for (const b of group.batches) {
    for (const m of b.mutations) {
      if (m.type === "set_cell") put(m.address.row, m.address.col, written(m.new_value, m.new_formula), written(m.old_value, m.old_formula));
      else if (m.type === "set_range")
        m.values.forEach((row, i) => row.forEach((v, j) => put(m.start_row + i, m.start_col + j, written(v), written(m.old_values?.[i]?.[j]))));
      else if (m.type === "set_format") for (const w of formatWords(m.new_format)) if (!formats.includes(w)) formats.push(w);
      else other.push(...otherLines(m));
    }
  }
  const all = [...writes.values()].sort((a, b) => a.row - b.row || a.col - b.col);
  const added = all.filter((w) => !w.old && w.value).length;
  const changed = all.filter((w) => w.old && w.value).length;
  const cleared = all.filter((w) => w.old && !w.value).length;
  const parts: string[] = [];
  if (added) parts.push(`${plural(added, "cell")} added`);
  if (changed) parts.push(`${added ? changed : plural(changed, "cell")} changed`);
  if (cleared) parts.push(`${added || changed ? cleared : plural(cleared, "cell")} cleared`);
  const summary = [parts.join(", "), formats.length ? (parts.length ? formats.join(", ") : `formatted: ${formats.join(", ")}`) : ""].filter(Boolean).join(" · ");
  const cells = all.slice(0, max).map((w) => ({ ref: a1(w.row, w.col), value: w.value, old: w.old || null, formula: w.formula }));
  return { summary: summary.replace(/^./, (c) => c.toUpperCase()), cells, more: Math.max(0, all.length - max), other };
}
