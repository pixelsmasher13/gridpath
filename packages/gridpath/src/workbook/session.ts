/**
 * Headless workbook session: the .xlsx bytes, one engine model built from
 * them, and the mutation → engine → readback → patch loop the tools drive.
 *
 * The engine (IronCalc) is the single source of truth for cell state.
 * Mutations are applied to it directly, so reference shifting on structural
 * edits is the engine's job, not ours. Save = surgical patch of the ORIGINAL
 * bytes from a diff of the load-time baseline against the live snapshot.
 */
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { CellFormatShape, UniverMutation, ChangeBatch } from "../core/types";
import { buildCellBaseline, buildWorkbookPatch, type BaselineCell } from "../core/surgicalPatch";
import { buildSaveMirror } from "../core/saveMirror";
import { cellKey, mapSnapshot, parseA1Range, type EngineSnapshot, type MappedSnapshot, type WorkbookSnapshot } from "./snapshot";

// ---- engine loading ----

type EngineWorkbook = {
  evaluate(): void;
  sheetNames(): string[];
  snapshot(): string;
  cell(sheet: number, row: number, col: number): string;
  setInput(sheet: number, row: number, col: number, value: string): void;
  clearContents(sheet: number, row: number, col: number, height: number, width: number): void;
  setRangeStyle(sheet: number, row: number, col: number, height: number, width: number, path: string, value: string): void;
  insertRows(sheet: number, row: number, count: number): void;
  deleteRows(sheet: number, row: number, count: number): void;
  insertColumns(sheet: number, col: number, count: number): void;
  deleteColumns(sheet: number, col: number, count: number): void;
  newSheet(): number;
  renameSheet(sheet: number, name: string): void;
  deleteSheet(sheet: number): void;
  setSheetColor(sheet: number, rgb: string): void;
  setFrozen(sheet: number, rows: number, cols: number): void;
  setColumnsWidth(sheet: number, start: number, end: number, width: number): void;
  setRowsHeight(sheet: number, start: number, end: number, height: number): void;
  setColumnsHidden(sheet: number, start: number, end: number, hidden: boolean): void;
  setRowsHidden(sheet: number, start: number, end: number, hidden: boolean): void;
  newDefinedName(name: string, formula: string): void;
};
type Engine = {
  Workbook: { open(bytes: Uint8Array): EngineWorkbook };
  patch(base: Uint8Array, patchJson: string): Uint8Array;
};

let engineCache: Engine | null = null;
export function loadEngine(): Engine {
  if (engineCache) return engineCache;
  const require = createRequire(import.meta.url);
  const here = path.dirname(fileURLToPath(import.meta.url));
  // src/workbook → ../../engine (dev) or dist → ../engine (published).
  const candidates = [path.resolve(here, "../../engine/gridpath_engine.js"), path.resolve(here, "../engine/gridpath_engine.js")];
  for (const c of candidates) {
    try {
      engineCache = require(c) as Engine;
      return engineCache;
    } catch (e: any) {
      if (e?.code !== "MODULE_NOT_FOUND") throw e;
    }
  }
  throw new Error("gridpath engine not found — run `npm run build:engine` (looked in: " + candidates.join(", ") + ")");
}

// ---- helpers ----

/** Excel px → chars, inverse of snapshot.ts colWidthPx. */
const pxToChars = (px: number) => Math.max(0, (px - 5) / 7);
/** px → points, inverse of snapshot.ts rowHeightPx. */
const pxToPts = (px: number) => px / 1.333;

/** What the engine should be told for a mutation's literal value. */
function inputOf(value: string | number | boolean | null, formula?: string | null): string {
  if (formula) return formula.startsWith("=") ? formula : `=${formula}`;
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

/** Per-format-key engine style paths. Returns [path, value] pairs. */
function stylePaths(f: CellFormatShape): Array<[string, string]> {
  const out: Array<[string, string]> = [];
  if (f.bold !== undefined) out.push(["font.b", String(!!f.bold)]);
  if (f.italic !== undefined) out.push(["font.i", String(!!f.italic)]);
  if (f.underline !== undefined) out.push(["font.u", String(!!f.underline)]);
  if (f.strike !== undefined) out.push(["font.strike", String(!!f.strike)]);
  if (f.font_color !== undefined) out.push(["font.color", f.font_color ?? ""]);
  if (f.background_color !== undefined) out.push(["fill.color", f.background_color ?? ""]);
  if (f.font_size !== undefined) out.push(["font.size", String(f.font_size)]);
  if (f.horizontal_align !== undefined) out.push(["alignment.horizontal", f.horizontal_align]);
  if (f.vertical_align !== undefined) out.push(["alignment.vertical", f.vertical_align === "middle" ? "center" : f.vertical_align]);
  if (f.number_format !== undefined) out.push(["num_fmt", f.number_format]);
  if (f.wrap_text !== undefined) out.push(["alignment.wrap_text", String(!!f.wrap_text)]);
  // font_family and indent have no engine path; they still reach the file
  // through the save mirror.
  return out;
}

export type ReadCell = { cell: string; value: string | number | boolean | null; formula: string | null; display?: string };

/** The slice of the grid handle `captureContext` / `copyRange` read from. */
export type CellReader = {
  getWorkbookSnapshot(): WorkbookSnapshot;
  getSheetNames(): string[];
  getCell(sheet: string, row: number, col: number): { value: any; display?: string; formula: string | null } | null;
  getCellFormat(sheet: string, row: number, col: number): CellFormatShape | null;
  getExternalPin(sheet: string, row: number, col: number): { v: any; f: string } | null;
  getNote(sheet: string, row: number, col: number): string | null;
};

const colLetters = (c: number) => {
  let s = "";
  let n = c + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};
export const a1Of = (row: number, col: number) => `${colLetters(col)}${row + 1}`;

// ---- session ----

export class WorkbookSession {
  private mapped: MappedSnapshot | null = null;
  private readonly baseline: Map<string, Map<string, BaselineCell>>;
  /** Batches applied this session, in order. Save consumes accepted + pending. */
  readonly batches: ChangeBatch[] = [];

  private constructor(
    readonly path: string,
    readonly originalBytes: Uint8Array,
    private readonly wb: EngineWorkbook,
  ) {
    this.baseline = buildCellBaseline(this.snapshot());
  }

  static open(filePath: string, bytes: Uint8Array): WorkbookSession {
    const engine = loadEngine();
    return new WorkbookSession(filePath, bytes, engine.Workbook.open(bytes));
  }

  // -- reads --

  /** Univer-shaped snapshot of the live model (cached until the next write). */
  snapshot(): WorkbookSnapshot {
    return this.mappedSnapshot().workbook;
  }

  private mappedSnapshot(): MappedSnapshot {
    if (!this.mapped) {
      const raw = JSON.parse(this.wb.snapshot()) as EngineSnapshot;
      this.mapped = mapSnapshot(raw, this.path);
    }
    return this.mapped;
  }

  sheetNames(): string[] {
    return this.wb.sheetNames();
  }

  sheetIndex(name: string): number {
    const i = this.wb.sheetNames().indexOf(name);
    if (i < 0) throw new Error(`Sheet not found: ${name}`);
    return i;
  }

  /** Read one cell (0-based) from the live model. */
  cell(sheet: string, row: number, col: number): ReadCell {
    const idx = this.sheetIndex(sheet);
    const c = JSON.parse(this.wb.cell(idx, row + 1, col + 1)) as { content: string; value: any; display: string };
    const formula = typeof c.content === "string" && c.content.startsWith("=") ? c.content : null;
    const out: ReadCell = { cell: a1Of(row, col), value: c.value ?? null, formula };
    if (c.display && String(c.value) !== c.display) out.display = c.display;
    return out;
  }

  /** Every non-empty cell in an A1 range, from the live model. */
  readRange(sheet: string, range: string, limit = 500): { cells: ReadCell[]; truncated: boolean } {
    const r = parseA1Range(range);
    if (!r) throw new Error(`Bad range: ${range}`);
    const cells: ReadCell[] = [];
    let truncated = false;
    outer: for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startColumn; col <= r.endColumn; col++) {
        const c = this.cell(sheet, row, col);
        if (c.value === null && !c.formula) continue;
        if (cells.length >= limit) {
          truncated = true;
          break outer;
        }
        cells.push(c);
      }
    }
    return { cells, truncated };
  }

  /** Structural reader for the pure modules (captureContext, copyRange, …). */
  reader(): CellReader {
    const self = this;
    return {
      getWorkbookSnapshot: () => self.snapshot(),
      getSheetNames: () => self.sheetNames(),
      getCell(sheet, row, col) {
        const c = self.cell(sheet, row, col);
        if (c.value === null && !c.formula) return null;
        return { value: c.value, display: c.display, formula: c.formula };
      },
      getCellFormat: (sheet, row, col) => self.mappedSnapshot().formats.get(cellKey(sheet, row, col)) ?? null,
      getExternalPin: () => null,
      getNote: (sheet, row, col) => self.mappedSnapshot().notes.get(cellKey(sheet, row, col)) ?? null,
    };
  }

  // -- writes --

  /**
   * Apply a batch of mutations to the engine, evaluate once, invalidate the
   * snapshot. Throws on the first mutation the engine rejects; the batch is
   * recorded only if every mutation applied.
   */
  apply(batch: Omit<ChangeBatch, "status" | "created_at"> & Partial<Pick<ChangeBatch, "status" | "created_at">>): ChangeBatch {
    // Capture the pre-edit state on each cell mutation so the review UI can
    // show before/after and a rejected batch is a real record of what it
    // undid. Tool calls arrive with old_* unset (no grid to read from).
    const mutations = batch.mutations.map((m) => this.withOldState(m));
    for (const m of mutations) this.applyOne(m);
    this.wb.evaluate();
    this.mapped = null;
    const rec: ChangeBatch = { status: "pending", created_at: new Date().toISOString(), ...batch, mutations };
    this.batches.push(rec);
    return rec;
  }

  private withOldState(m: UniverMutation): UniverMutation {
    const known = (name: string) => this.wb.sheetNames().includes(name);
    if (m.type === "set_cell" && known(m.address.sheet) && m.old_value == null && m.old_formula == null) {
      const c = this.cell(m.address.sheet, m.address.row, m.address.col);
      return { ...m, old_value: c.value as string | number | null, old_formula: c.formula };
    }
    if (m.type === "set_range" && known(m.sheet) && !m.old_values) {
      const old_values = m.values.map((rowVals, dr) =>
        rowVals.map((_, dc) => this.cell(m.sheet, m.start_row + dr, m.start_col + dc).value as string | number | null),
      );
      return { ...m, old_values };
    }
    if (m.type === "clear_range" && known(m.sheet)) {
      const cells = m.cells.map((c) => {
        if (c.old_value != null || c.old_formula) return c;
        const cur = this.cell(m.sheet, c.row, c.col);
        return { ...c, old_value: cur.value, old_formula: cur.formula };
      });
      return { ...m, cells };
    }
    return m;
  }

  private applyOne(m: UniverMutation): void {
    const wb = this.wb;
    const idx = (name: string) => this.sheetIndex(name);
    switch (m.type) {
      case "set_cell":
        wb.setInput(idx(m.address.sheet), m.address.row + 1, m.address.col + 1, inputOf(m.new_value, m.new_formula));
        return;
      case "set_range": {
        const s = idx(m.sheet);
        m.values.forEach((rowVals, dr) =>
          rowVals.forEach((v, dc) => {
            if (v === null || v === undefined || v === "") return;
            wb.setInput(s, m.start_row + dr + 1, m.start_col + dc + 1, inputOf(v));
          }),
        );
        return;
      }
      case "set_format": {
        const s = idx(m.sheet);
        const paths = stylePaths(m.new_format);
        for (const c of m.cells) for (const [p, v] of paths) wb.setRangeStyle(s, c.row + 1, c.col + 1, 1, 1, p, v);
        return;
      }
      case "clear_range": {
        const s = idx(m.sheet);
        for (const c of m.cells) wb.clearContents(s, c.row + 1, c.col + 1, 1, 1);
        return;
      }
      case "insert_rows":
        wb.insertRows(idx(m.sheet), m.before + 1, m.count);
        return;
      case "delete_rows":
        wb.deleteRows(idx(m.sheet), m.start + 1, m.count);
        return;
      case "insert_columns":
        wb.insertColumns(idx(m.sheet), m.before + 1, m.count);
        return;
      case "delete_columns":
        wb.deleteColumns(idx(m.sheet), m.start + 1, m.count);
        return;
      case "create_sheet": {
        const s = wb.newSheet();
        wb.renameSheet(s, m.name);
        if (m.tab_color) wb.setSheetColor(s, m.tab_color);
        return;
      }
      case "rename_sheet":
        wb.renameSheet(idx(m.old_name), m.new_name);
        return;
      case "delete_sheet":
        wb.deleteSheet(idx(m.name));
        return;
      case "set_column_width": {
        const s = idx(m.sheet);
        for (const c of m.columns) wb.setColumnsWidth(s, c + 1, c + 1, pxToChars(m.new_width));
        return;
      }
      case "set_row_height": {
        const s = idx(m.sheet);
        for (const r of m.rows) wb.setRowsHeight(s, r + 1, r + 1, pxToPts(m.new_height));
        return;
      }
      case "freeze_panes":
        wb.setFrozen(idx(m.sheet), m.freeze_rows, m.freeze_cols);
        return;
      case "unfreeze_panes":
        wb.setFrozen(idx(m.sheet), 0, 0);
        return;
      case "hide_rows":
      case "show_rows": {
        const s = idx(m.sheet);
        for (const r of m.rows) wb.setRowsHidden(s, r + 1, r + 1, m.type === "hide_rows");
        return;
      }
      case "hide_columns":
      case "show_columns": {
        const s = idx(m.sheet);
        for (const c of m.columns) wb.setColumnsHidden(s, c + 1, c + 1, m.type === "hide_columns");
        return;
      }
      case "define_name":
        wb.newDefinedName(m.name, m.ref.startsWith("=") ? m.ref : `=${m.ref}`);
        return;
      case "merge_cells":
      case "unmerge_cells":
      case "set_note":
      case "delete_note":
        // No engine representation; they reach the file through the save
        // mirror / patch only. Nothing to evaluate.
        return;
    }
  }

  // -- save --

  /**
   * Patch the ORIGINAL bytes with everything in accepted + pending batches.
   * Returns the new file bytes, or the reason the surgical path can't be used.
   */
  buildSave(): { ok: true; bytes: Uint8Array; patch: unknown } | { ok: false; reason: string } {
    const mirror = buildSaveMirror(this.batches);
    const built = buildWorkbookPatch(this.snapshot(), this.baseline, mirror);
    if (!built.ok) return { ok: false, reason: built.reason };
    const bytes = loadEngine().patch(this.originalBytes, JSON.stringify(built.patch));
    return { ok: true, bytes, patch: built.patch };
  }
}
