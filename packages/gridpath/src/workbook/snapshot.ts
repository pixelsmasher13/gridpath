/**
 * Engine snapshot → the Univer-shaped workbook snapshot (`IWorkbookData`-like)
 * that the core modules (`workbookIndex`, `scriptRunner`, `captureContext`,
 * `surgicalPatch`) already consume. The engine is the only reader of the
 * .xlsx; this file is the one place that knows both shapes.
 *
 * Engine addressing is 1-based (IronCalc); the snapshot is 0-based.
 */
import type { CellFormatShape } from "../core/types";

// ---- engine JSON (see crates/engine-wasm/src/lib.rs `snapshot`) ----

export type EngineColor = string | [number, number] | null | undefined;
export type EngineStyle = {
  alignment?: { horizontal?: string; vertical?: string; wrap_text?: boolean };
  num_fmt: string;
  fill: { color?: EngineColor };
  font: { strike?: boolean; u?: boolean; b?: boolean; i?: boolean; sz: number; color?: EngineColor; name: string };
  border: Record<string, unknown>;
  quote_prefix: boolean;
};
export type EngineCell = { r: number; c: number; f?: string; v: string | number | boolean | null; d?: string; s?: number };
export type EngineSheet = {
  name: string;
  hidden: boolean;
  color: string | null;
  frozenRows: number;
  frozenColumns: number;
  showGridLines: boolean;
  merges: string[];
  cols: Array<{ min: number; max: number; width: number; hidden: boolean }>;
  rows: Array<{ r: number; height: number; hidden: boolean }>;
  comments: Array<{ ref: string; text: string; author: string }>;
  cells: EngineCell[];
};
export type EngineSnapshot = {
  sheets: EngineSheet[];
  styles: Record<string, EngineStyle>;
  definedNames: Array<{ name: string; formula: string; sheetId: number | null }>;
};

// ---- output ----

export type SheetSnapshot = {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cellData: Record<number, Record<number, { v?: string | number | boolean; f?: string; s?: Record<string, unknown> }>>;
  mergeData: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
  columnData: Record<number, { w?: number; hd?: 0 | 1 }>;
  rowData: Record<number, { h?: number; hd?: 0 | 1 }>;
  freeze?: { xSplit: number; ySplit: number; startRow: number; startColumn: number };
  hidden?: 1;
  tabColor?: string;
  showGridlines?: 0;
};
export type WorkbookSnapshot = {
  id: string;
  sheetOrder: string[];
  sheets: Record<string, SheetSnapshot>;
  resources: Array<{ name: string; data: string }>;
};

export type MappedSnapshot = {
  workbook: WorkbookSnapshot;
  /** `${sheet}\u0000${row}\u0000${col}` → number-formatted display text. */
  display: Map<string, string>;
  /** Same key → the agent-facing format summary of the cell's style. */
  formats: Map<string, CellFormatShape>;
  /** Cell notes keyed the same way. */
  notes: Map<string, string>;
};

export const cellKey = (sheet: string, row: number, col: number) => `${sheet}\u0000${row}\u0000${col}`;

/** Excel column width (chars) → px, matching xlsxImport.ts. */
const colWidthPx = (chars: number) => Math.round(chars * 7 + 5);
/** Row height (pt) → px, matching xlsxImport.ts. */
const rowHeightPx = (pts: number) => Math.round(pts * 1.333);

export function sheetIdOf(name: string): string {
  return "sheet_" + name.replace(/[^a-zA-Z0-9]/g, "_");
}

function rgbOf(c: EngineColor): string | undefined {
  return typeof c === "string" ? c : undefined;
}

/** IronCalc style → the agent-facing `CellFormatShape`. */
export function styleToFormat(s: EngineStyle | undefined): CellFormatShape | null {
  if (!s) return null;
  const f: CellFormatShape = {};
  if (s.font?.b) f.bold = true;
  if (s.font?.i) f.italic = true;
  if (s.font?.u) f.underline = true;
  if (s.font?.strike) f.strike = true;
  const fc = rgbOf(s.font?.color);
  if (fc) f.font_color = fc;
  const bg = rgbOf(s.fill?.color);
  if (bg) f.background_color = bg;
  if (s.font?.sz) f.font_size = s.font.sz;
  if (s.font?.name) f.font_family = s.font.name;
  const h = s.alignment?.horizontal;
  if (h === "left" || h === "center" || h === "right") f.horizontal_align = h;
  const v = s.alignment?.vertical;
  if (v === "top" || v === "bottom") f.vertical_align = v;
  else if (v === "center") f.vertical_align = "middle";
  if (s.num_fmt && s.num_fmt.toLowerCase() !== "general") f.number_format = s.num_fmt;
  if (s.alignment?.wrap_text) f.wrap_text = true;
  return Object.keys(f).length ? f : null;
}

/** IronCalc style → Univer inline `IStyleData` (mirrors buildUniverStyle in xlsxImport.ts). */
export function styleToUniver(s: EngineStyle | undefined): Record<string, unknown> | undefined {
  const f = styleToFormat(s);
  if (!f) return undefined;
  const u: Record<string, unknown> = {};
  if (f.background_color) u.bg = { rgb: f.background_color };
  if (f.font_color) u.cl = { rgb: f.font_color };
  if (f.bold) u.bl = 1;
  if (f.italic) u.it = 1;
  if (f.underline) u.ul = { s: 1 };
  if (f.strike) u.st = { s: 1 };
  if (f.font_size) u.fs = f.font_size;
  if (f.font_family) u.ff = f.font_family;
  if (f.horizontal_align) u.ht = f.horizontal_align === "left" ? 1 : f.horizontal_align === "center" ? 2 : 3;
  if (f.vertical_align) u.vt = f.vertical_align === "top" ? 1 : f.vertical_align === "middle" ? 2 : 3;
  if (f.number_format) u.n = { pattern: f.number_format };
  if (f.wrap_text) u.tb = 2;
  return u;
}

/** "B2:C4" → 0-based range; "B2" → single cell. Null on malformed input. */
export function parseA1Range(ref: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref.trim());
  if (!m) return null;
  const col = (s: string) => s.toUpperCase().split("").reduce((a, ch) => a * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  const startColumn = col(m[1]);
  const startRow = Number(m[2]) - 1;
  const endColumn = m[3] ? col(m[3]) : startColumn;
  const endRow = m[4] ? Number(m[4]) - 1 : startRow;
  return { startRow, endRow, startColumn, endColumn };
}

export function mapSnapshot(engine: EngineSnapshot, id = "workbook"): MappedSnapshot {
  const display = new Map<string, string>();
  const formats = new Map<string, CellFormatShape>();
  const notes = new Map<string, string>();
  const univerStyles = new Map<number, Record<string, unknown> | undefined>();
  const formatCache = new Map<number, CellFormatShape | null>();
  const sheets: Record<string, SheetSnapshot> = {};
  const sheetOrder: string[] = [];

  for (const es of engine.sheets) {
    const sid = sheetIdOf(es.name);
    sheetOrder.push(sid);
    const cellData: SheetSnapshot["cellData"] = {};
    let maxRow = 0;
    let maxCol = 0;
    for (const c of es.cells) {
      const r0 = c.r - 1;
      const c0 = c.c - 1;
      const cell: { v?: string | number | boolean; f?: string; s?: Record<string, unknown> } = {};
      if (c.v !== null && c.v !== undefined) cell.v = c.v;
      if (c.f) cell.f = c.f;
      if (c.s !== undefined) {
        if (!univerStyles.has(c.s)) {
          univerStyles.set(c.s, styleToUniver(engine.styles[String(c.s)]));
          formatCache.set(c.s, styleToFormat(engine.styles[String(c.s)]));
        }
        const us = univerStyles.get(c.s);
        if (us) cell.s = us;
        const fmt = formatCache.get(c.s);
        if (fmt) formats.set(cellKey(es.name, r0, c0), fmt);
      }
      if (cell.v === undefined && !cell.f && !cell.s) continue;
      (cellData[r0] ??= {})[c0] = cell;
      if (c.d !== undefined && c.d !== "" && String(c.v) !== c.d) display.set(cellKey(es.name, r0, c0), c.d);
      if (r0 > maxRow) maxRow = r0;
      if (c0 > maxCol) maxCol = c0;
    }
    const columnData: SheetSnapshot["columnData"] = {};
    for (const col of es.cols) {
      for (let c = col.min; c <= col.max; c++) {
        const entry: { w?: number; hd?: 0 | 1 } = {};
        if (col.width) entry.w = colWidthPx(col.width);
        if (col.hidden) entry.hd = 1;
        if (entry.w !== undefined || entry.hd) columnData[c - 1] = entry;
      }
    }
    const rowData: SheetSnapshot["rowData"] = {};
    for (const row of es.rows) {
      const entry: { h?: number; hd?: 0 | 1 } = {};
      if (row.height) entry.h = rowHeightPx(row.height);
      if (row.hidden) entry.hd = 1;
      if (entry.h !== undefined || entry.hd) rowData[row.r - 1] = entry;
    }
    const mergeData: SheetSnapshot["mergeData"] = [];
    for (const m of es.merges) {
      const r = parseA1Range(m);
      if (r) mergeData.push(r);
    }
    for (const n of es.comments) {
      const r = parseA1Range(n.ref);
      if (r) notes.set(cellKey(es.name, r.startRow, r.startColumn), n.text);
    }
    const sheet: SheetSnapshot = {
      id: sid,
      name: es.name,
      rowCount: Math.max(10000, maxRow + 100),
      columnCount: Math.max(200, maxCol + 20),
      cellData,
      mergeData,
      columnData,
      rowData,
    };
    if (es.frozenRows || es.frozenColumns) {
      sheet.freeze = { xSplit: es.frozenColumns, ySplit: es.frozenRows, startRow: es.frozenRows, startColumn: es.frozenColumns };
    }
    if (es.hidden) sheet.hidden = 1;
    if (es.color) sheet.tabColor = es.color;
    if (!es.showGridLines) sheet.showGridlines = 0;
    sheets[sid] = sheet;
  }

  const resources: WorkbookSnapshot["resources"] = [];
  if (engine.definedNames.length) {
    // Same resource shape workbookIndex reads for defined names.
    const names: Record<string, { name: string; formulaOrRefString: string; localSheetId?: string }> = {};
    engine.definedNames.forEach((d, i) => {
      names[`dn_${i}`] = { name: d.name, formulaOrRefString: d.formula };
    });
    resources.push({ name: "SHEET_DEFINED_NAME_PLUGIN", data: JSON.stringify(names) });
  }

  return { workbook: { id, sheetOrder, sheets, resources }, display, formats, notes };
}
