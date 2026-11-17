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
import {
  cellKey,
  mapCell,
  mapSnapshot,
  newStyleCache,
  parseA1Range,
  sheetIdOf,
  stripImplicitIntersection,
  type EngineCell,
  type EngineSnapshot,
  type EngineStyle,
  type MappedSnapshot,
  type WorkbookSnapshot,
} from "./snapshot";
import { buildWorkbookIndex, type WorkbookIndex } from "../core/workbookIndex";
import { isLinkSheetName, linkSheetName, numberedPrefix, rewriteLinks, splitLinkSheetName } from "./links";

// ---- engine loading ----

/** What `evaluateEdited` returns: refresh everything, or just these cells. */
export type EngineDelta =
  | { full: true }
  | { full: false; cells: Array<EngineCell & { sheet: number }>; styles: Record<string, EngineStyle> };

type EngineWorkbook = {
  evaluate(): void;
  evaluateEdited(): string;
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
  /** Structural index (engine-built; same shape as buildWorkbookIndex). */
  workbookIndex?(): string;
  /** query_rows reply JSON for a query whose `sheet` is an index. */
  queryRows?(query: string): string;
  /**
   * Workspace: keep a hidden copy of `source`'s sheet in this workbook under
   * the name formulas use for it (`[NOW.xlsx]Model`). Returns
   * `{"sheet","created","changed"}`.
   */
  attachLink?(name: string, source: EngineWorkbook, sourceSheet: number): string;
  /** Workspace: carry changed source cells (flat 1-based row, column pairs) into the copy. */
  pushLinkCells?(linkSheet: number, source: EngineWorkbook, sourceSheet: number, cells: Int32Array): number;
  /** Workspace: the copied sheet was renamed in its own workbook. */
  renameLinkSheet?(sheet: number, name: string): void;
  /** Save: for each hidden copy formulas read, where they read it and the values there (JSON, `LinkCacheSheet[]`). */
  linkCache?(): string;
  /** The sheets of other workbooks that formulas name and no copy stands behind (JSON list, `["[1]Model"]`). */
  unresolvedLinks?(): string;
};
type Engine = {
  Workbook: { open(bytes: Uint8Array): EngineWorkbook };
  patch(base: Uint8Array, patchJson: string): Uint8Array;
  /** Bytes of a new, empty .xlsx with these sheets. */
  blankXlsx?(sheetNames: string[]): Uint8Array;
  /** The links to other workbooks an .xlsx was saved with (JSON, `SavedLink[]`). */
  externalLinks?(bytes: Uint8Array): string;
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

/** One other workbook an .xlsx was saved reading: `[index]Sheet!A1` in its formulas. */
export type SavedLink = {
  index: number;
  /** The other file as the .xlsx names it (`NOW.xlsx`, `file:///…`). Null for a link that is not to a workbook. */
  target: string | null;
  /** The other file's sheet names, as the link lists them. */
  sheets: string[];
};

/** What the engine reports of one hidden copy for a save: 0-based areas read, and the values in them. */
type LinkCacheSheet = { sheet: string; areas: Array<[number, number, number, number]>; cells: Array<{ r: number; c: number; v: unknown }> };

/**
 * What a save needs to know about the workbook's hidden copies, which only
 * the workspace can tell: the file each one copies.
 */
export type LinkSave = {
  /**
   * Each copy, by upper-cased name (`[NOW.XLSX]MODEL`): the workbook it
   * copies, and that workbook's number in the file when the file already
   * links to it.
   */
  copies: Map<string, { source: string; index?: number }>;
  /** How the file is to name a workbook it does not link to yet (`NOW.xlsx`, `file:///…`). */
  target: (source: string) => string;
};

/** A batch to apply: `status` and `created_at` are filled in when left out. */
export type NewBatch = Omit<ChangeBatch, "status" | "created_at"> & Partial<Pick<ChangeBatch, "status" | "created_at">>;

export type ApplyOptions = {
  /** Apply only the mutations this accepts; the recorded batch keeps them all. */
  only?: (m: UniverMutation) => boolean;
  /** False: leave the evaluation to a later `settle()`. */
  evaluate?: boolean;
};

export class WorkbookSession {
  private mapped: MappedSnapshot | null = null;
  private saved: SavedLink[] | null | undefined;
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

  /** The workbook's own sheets. The hidden copies of other workbooks' sheets are not among them. */
  sheetNames(): string[] {
    return this.wb.sheetNames().filter((n) => !isLinkSheetName(n));
  }

  /** Every sheet the engine holds, hidden copies included, in the engine's order (what a delta's sheet number counts). */
  allSheetNames(): string[] {
    return this.wb.sheetNames();
  }

  /** Names of the hidden copies this workbook holds (`[NOW.xlsx]Model`). */
  linkSheetNames(): string[] {
    return this.wb.sheetNames().filter(isLinkSheetName);
  }

  /** Whether the engine build can link workbooks. */
  supportsLinks(): boolean {
    return typeof this.wb.attachLink === "function" && typeof this.wb.pushLinkCells === "function";
  }

  /**
   * Structural index of the workbook, built by the engine from its own
   * model (`gridpath_engine_core::index`). Falls back to the TypeScript
   * builder over the snapshot when the engine binary predates the export.
   */
  workbookIndex(): WorkbookIndex {
    if (typeof this.wb.workbookIndex === "function") {
      return JSON.parse(this.wb.workbookIndex()) as WorkbookIndex;
    }
    return buildWorkbookIndex(this.snapshot());
  }

  /** `query_rows` served by the engine; `sheet` is a sheet NAME here. */
  queryRows(sheet: string, query: Record<string, unknown>): unknown {
    if (typeof this.wb.queryRows !== "function") {
      throw new Error("query_rows needs a newer engine build — run `npm run build:engine`");
    }
    const idx = this.sheetIndex(sheet);
    return JSON.parse(this.wb.queryRows(JSON.stringify({ ...query, sheet: idx })));
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
    const formula = typeof c.content === "string" && c.content.startsWith("=") ? stripImplicitIntersection(c.content) : null;
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
   * Apply a batch of mutations to the engine and evaluate once. Cell edits
   * recalculate only their dependency closure and the cached snapshot is
   * patched in place; structural edits fall back to a full recalculation and
   * a fresh snapshot. Throws on the first mutation the engine rejects; the
   * batch is recorded only if every mutation applied.
   */
  apply(batch: NewBatch, opts: ApplyOptions = {}): ChangeBatch {
    return this.applyTracked(batch, opts).batch;
  }

  /**
   * `apply`, also returning what the evaluation changed (null when
   * `opts.evaluate` is false: the caller evaluates with `settle`).
   */
  applyTracked(batch: NewBatch, opts: ApplyOptions = {}): { batch: ChangeBatch; delta: EngineDelta | null } {
    // Capture the pre-edit state on each cell mutation so the review UI can
    // show before/after and a rejected batch is a real record of what it
    // undid. Tool calls arrive with old_* unset (no grid to read from).
    const mutations = batch.mutations.map((m) => this.withOldState(m));
    for (const m of mutations) if (!opts.only || opts.only(m)) this.applyOne(m);
    const delta = opts.evaluate === false ? null : this.settle();
    const rec: ChangeBatch = { status: "pending", created_at: new Date().toISOString(), ...batch, mutations };
    this.batches.push(rec);
    return { batch: rec, delta };
  }

  /** Evaluate what was edited since the last evaluation and refresh the cached snapshot. */
  settle(): EngineDelta {
    const delta = JSON.parse(this.wb.evaluateEdited()) as EngineDelta;
    if (delta.full || !this.mapped) this.mapped = null;
    else this.patchMapped(delta);
    return delta;
  }

  // -- links to other workbooks (driven by the workspace) --

  /**
   * Make this workbook's hidden copy `name` (`[NOW.xlsx]Model`) equal to
   * `sourceSheet` of `source`, creating it when there is none. The changed
   * cells recalculate on the next `apply` / `settle`.
   */
  syncLink(name: string, source: WorkbookSession, sourceSheet: string): { created: boolean; changed: number } {
    if (source === this) throw new Error("a workbook does not link to itself");
    if (!this.wb.attachLink) throw new Error("links between workbooks need a newer engine build — run `npm run build:engine`");
    const r = JSON.parse(this.wb.attachLink(name, source.wb, source.sheetIndex(sourceSheet))) as { created: boolean; changed: number };
    return { created: r.created, changed: r.changed };
  }

  /** Carry the cells `source` just changed (flat 1-based row, column pairs) into the copy `name`. */
  pushLink(name: string, source: WorkbookSession, sourceSheetIndex: number, cells: Int32Array): number {
    if (!this.wb.pushLinkCells) throw new Error("links between workbooks need a newer engine build — run `npm run build:engine`");
    return this.wb.pushLinkCells(this.sheetIndex(name), source.wb, sourceSheetIndex, cells);
  }

  /** Drop the copy `name`: the formulas that read it show #REF! from the next evaluation. */
  detachLink(name: string): void {
    this.wb.deleteSheet(this.sheetIndex(name));
  }

  /**
   * The copy `from` gets the name `to`, and the formulas that read it are
   * respelled. Values do not change.
   */
  renameLink(from: string, to: string): void {
    if (!this.wb.renameLinkSheet) throw new Error("links between workbooks need a newer engine build — run `npm run build:engine`");
    this.wb.renameLinkSheet(this.sheetIndex(from), to);
    this.mapped = null;
  }

  /**
   * The links to other workbooks the file was saved with, in the order of
   * their numbers. Null when this engine build cannot tell.
   */
  savedLinks(): SavedLink[] | null {
    if (this.saved === undefined) {
      const read = loadEngine().externalLinks;
      if (typeof read !== "function") this.saved = null;
      else {
        try {
          this.saved = JSON.parse(read(this.originalBytes)) as SavedLink[];
        } catch {
          // Link parts nobody could read: the formulas keep their saved values.
          this.saved = [];
        }
      }
    }
    return this.saved;
  }

  /**
   * The sheets of other workbooks that formulas here name and no copy
   * stands behind (`[1]Model` in a file as it was saved). Such a formula
   * shows the value the file was saved with.
   */
  unresolvedLinks(): string[] {
    if (typeof this.wb.unresolvedLinks !== "function") return [];
    return JSON.parse(this.wb.unresolvedLinks()) as string[];
  }

  /**
   * The copies something here reads (a formula, a defined name), by
   * upper-cased name. Null when this engine build cannot tell.
   */
  copiesRead(): Set<string> | null {
    if (typeof this.wb.linkCache !== "function") return null;
    return new Set((JSON.parse(this.wb.linkCache()) as LinkCacheSheet[]).map((c) => c.sheet.toUpperCase()));
  }

  /**
   * Fold an incremental delta into the cached snapshot: the same per-cell
   * mapping `mapSnapshot` uses, applied to just the cells the engine says
   * changed. A 2M-cell workbook no longer re-serialises per edit.
   */
  private patchMapped(delta: Extract<EngineDelta, { full: false }>): void {
    const mapped = this.mapped;
    if (!mapped) return;
    const names = this.wb.sheetNames();
    const cache = newStyleCache();
    for (const ec of delta.cells) {
      const name = names[ec.sheet];
      if (name !== undefined && isLinkSheetName(name)) continue;
      const sheet = name !== undefined ? mapped.workbook.sheets[sheetIdOf(name)] : undefined;
      if (!sheet) {
        // A sheet the snapshot doesn't know: safer to rebuild than to guess.
        this.mapped = null;
        return;
      }
      const r0 = ec.r - 1;
      const c0 = ec.c - 1;
      const key = cellKey(name!, r0, c0);
      const { cell, format, display } = mapCell(ec, delta.styles, cache);
      if (cell) {
        (sheet.cellData[r0] ??= {})[c0] = cell;
        if (r0 + 100 > sheet.rowCount) sheet.rowCount = r0 + 100;
        if (c0 + 20 > sheet.columnCount) sheet.columnCount = c0 + 20;
      } else if (sheet.cellData[r0]) {
        delete sheet.cellData[r0][c0];
        if (Object.keys(sheet.cellData[r0]).length === 0) delete sheet.cellData[r0];
      }
      if (format) mapped.formats.set(key, format);
      else mapped.formats.delete(key);
      if (display !== null) mapped.display.set(key, display);
      else mapped.display.delete(key);
    }
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
        // The copy of another workbook's sheet follows that sheet's new name.
        if (isLinkSheetName(m.old_name) && wb.renameLinkSheet && splitLinkSheetName(m.new_name)) wb.renameLinkSheet(idx(m.old_name), m.new_name);
        else wb.renameSheet(idx(m.old_name), m.new_name);
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
   *
   * A workbook that reads others needs `links` (from the workspace): each
   * other workbook is written into the file the way Excel does it, as a
   * numbered link with the file's path and the last value of every cell
   * read there, and the formulas are spelled `[1]Model!B23`.
   */
  buildSave(links?: LinkSave | null): { ok: true; bytes: Uint8Array; patch: unknown } | { ok: false; reason: string; links?: true } {
    const copies = new Set(this.linkSheetNames().map((n) => n.toUpperCase()));
    if (copies.size > 0 && typeof this.wb.linkCache !== "function") {
      return { ok: false, links: true, reason: "Writing links between workbooks into an .xlsx needs a newer engine build (run `npm run build:engine`). Nothing was saved." };
    }

    // Other workbook → its number in the file: the one the file has, or the next free one.
    const savedCount = this.savedLinks()?.length ?? 0;
    const fresh = new Map<string, number>();
    const targets = new Map<number, string>();
    const unknown = new Set<string>();
    const numberOf = (name: string): number | null => {
      const parts = splitLinkSheetName(name);
      if (parts && /^\d+$/.test(parts.book)) return Number(parts.book);
      const copy = links?.copies.get(name.toUpperCase());
      if (!copy) {
        unknown.add(name);
        return null;
      }
      if (copy.index !== undefined) return copy.index;
      let n = fresh.get(copy.source);
      if (n === undefined) {
        n = savedCount + fresh.size + 1;
        fresh.set(copy.source, n);
        targets.set(n, links!.target(copy.source));
      }
      return n;
    };
    // Number → sheet (upper-cased) → what the file keeps of it.
    type KeptSheet = { name: string; areas: LinkCacheSheet["areas"]; cells: LinkCacheSheet["cells"] };
    const kept = new Map<number, Map<string, KeptSheet>>();
    const keep = (n: number, sheet: string): KeptSheet => {
      let sheets = kept.get(n);
      if (!sheets) kept.set(n, (sheets = new Map()));
      let k = sheets.get(sheet.toUpperCase());
      if (!k) sheets.set(sheet.toUpperCase(), (k = { name: sheet, areas: [], cells: [] }));
      return k;
    };
    // A link that has no copy behind it is left as it is written.
    const formulaText = (f: string): string =>
      f.includes("[")
        ? rewriteLinks(f, (l) => {
            const name = linkSheetName(l.book, l.sheet);
            if (l.dir || !copies.has(name.toUpperCase())) return null;
            const n = numberOf(name);
            if (n === null) return null;
            keep(n, l.sheet);
            return /^\d+$/.test(l.book) ? null : numberedPrefix(n, l.sheet);
          })
        : f;

    const mirror = buildSaveMirror(this.batches);
    const built = buildWorkbookPatch(this.snapshot(), this.baseline, mirror, copies.size > 0 ? { formulaText } : undefined);
    if (!built.ok) return { ok: false, reason: built.reason };

    if (copies.size > 0) {
      for (const c of JSON.parse(this.wb.linkCache!()) as LinkCacheSheet[]) {
        const parts = splitLinkSheetName(c.sheet);
        const n = parts ? numberOf(c.sheet) : null;
        if (!parts || n === null) continue;
        const k = keep(n, parts.sheet);
        k.areas = c.areas;
        k.cells = c.cells;
      }
      if (unknown.size > 0) {
        return {
          ok: false,
          links: true,
          reason:
            `Formulas here read ${[...unknown].join(", ")}, and the save was not told which ${unknown.size === 1 ? "file that is" : "files those are"}, so nothing was saved. ` +
            "Save through the workspace (save_workbook), which knows the workbooks that are open together.",
        };
      }
      const externalLinks = [...kept]
        .sort((a, b) => a[0] - b[0])
        .map(([index, sheets]) => ({ index, ...(targets.has(index) ? { target: targets.get(index) } : {}), sheets: [...sheets.values()] }));
      if (externalLinks.length > 0) built.patch.externalLinks = externalLinks;
    }

    const bytes = loadEngine().patch(this.originalBytes, JSON.stringify(built.patch));
    return { ok: true, bytes, patch: built.patch };
  }
}
