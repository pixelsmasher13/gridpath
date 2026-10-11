/**
 * Pure model behind the viewport grid (big-model mode): sheet geometry
 * from the engine's layout, a tile cache for cells fetched on demand, and
 * style resolution. No DOM, no Tauri — unit-testable, and shared with the
 * headless harness that renders against a mocked engine.
 *
 * Coordinates: the engine is 1-based (row, column); the grid handle API the
 * rest of the app uses is 0-based (row0, col0). This module is 0-based
 * throughout and converts at the fetch boundary.
 */

// ---------------------------------------------------------------------------
// Engine payloads
// ---------------------------------------------------------------------------

export interface EngineAutoFilter {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  columns: Array<{
    /** 0-based offset within the range. */
    col: number;
    values: string[] | null;
    blank: boolean;
    custom: { and: boolean; items: Array<{ op: string; val: string }> } | null;
    /** Criteria this model cannot evaluate (top10, dynamic, colour, icon). */
    other: boolean;
  }>;
}

export interface EngineLayoutSheet {
  name: string;
  hidden: boolean;
  /** The sheet Excel had selected when the file was saved. */
  selected?: boolean;
  color: string | null;
  frozenRows: number;
  frozenColumns: number;
  showGridLines: boolean;
  merges: string[];
  cols: Array<{ min: number; max: number; width: number; hidden: boolean }>;
  /** `custom` = the file set the height explicitly (otherwise it is the default). */
  rows: Array<{ r: number; height: number; hidden: boolean; custom?: boolean }>;
  cellCount: number;
  maxRow: number;
  maxCol: number;
  /** Last row/column holding a value or formula (1-based; 0 when empty). */
  usedRow?: number;
  usedCol?: number;
  /** Outline (group) levels as [1-based index, level] pairs, sorted. */
  rowOutline?: Array<[number, number]>;
  colOutline?: Array<[number, number]>;
  /** Where the +/- button sits: after the group (Excel's default) or before. */
  summaryBelow?: boolean;
  summaryRight?: boolean;
  /** The sheet's AutoFilter (1-based bounds, header row = r0), if any. */
  autoFilter?: EngineAutoFilter | null;
  /** Saved view: first row/column of the scrolling pane and the active cell, 1-based. */
  topRow?: number;
  leftColumn?: number;
  activeRow?: number;
  activeCol?: number;
  /** The sheet's saved zoom in percent (Excel remembers it per sheet). */
  zoom?: number;
}

export interface EngineStyle {
  fill?: { color?: string | null };
  font?: { b?: boolean; i?: boolean; u?: boolean; strike?: boolean; color?: string | null; sz?: number; name?: string };
  alignment?: { horizontal?: string; vertical?: string; wrap_text?: boolean; indent?: number };
  num_fmt?: string;
  border?: Record<string, { style?: string; color?: string | null } | undefined>;
}

export interface EngineLayout {
  sheets: EngineLayoutSheet[];
  styles: Record<string, EngineStyle>;
  definedNames: Array<{ name: string; formula: string; sheetId: number | null }>;
}

/** One cell as the engine's window / snapshot serialiser writes it (1-based r/c). */
export interface EngineWindowCell {
  r: number;
  c: number;
  v?: unknown;
  d?: string;
  f?: string;
  s?: number;
  live?: unknown;
  fileSaved?: boolean;
}

export interface EngineWindow {
  sheet: number;
  r0: number;
  r1: number;
  c0: number;
  c1: number;
  cells: EngineWindowCell[];
  styles: Record<string, EngineStyle>;
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

// Univer's defaults for rows and columns without an explicit size, so the
// two grids draw the same workbook the same way.
export const DEFAULT_COL_PX = 88;
export const DEFAULT_ROW_PX = 24;
export const HEADER_W = 46;
export const HEADER_H = 22;
/** Rows/columns kept navigable past the used range. */
const ROW_SLACK = 40;
const COL_SLACK = 8;
const MIN_ROWS = 100;
const MIN_COLS = 26;

export const colWidthPx = (chars: number) => Math.round(chars * 7 + 5);
export const rowHeightPx = (pts: number) => Math.round(pts * 1.333);

export function colLetters(c0: number): string {
  let n = c0;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

export function parseA1(ref: string): { row0: number; col0: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row0: Number(m[2]) - 1, col0: col - 1 };
}

export interface MergeRect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/**
 * Pixel geometry of one sheet: per-axis sizes with hidden rows/cols at 0,
 * prefix sums for O(log n) hit-testing, merges as rectangles.
 */
export interface SizeOverrides {
  rows: Map<number, number>;
  cols: Map<number, number>;
  /** Session replacement for the sheet's merge list (A1 ranges). */
  merges?: string[];
  /** Session replacement for the frozen pane counts. */
  frozen?: { rows: number; cols: number };
}

/** Width of one outline level column in the group gutter (sheet px). */
export const OUTLINE_STEP = 13;
/** Gutter size for `maxLevel` outline levels: one column per level plus one
 * for the level buttons (Excel shows levels 1..n+1). Zero without groups. */
export function outlineGutter(maxLevel: number): number {
  return maxLevel > 0 ? (maxLevel + 1) * OUTLINE_STEP + 3 : 0;
}
/** Runs of consecutive indices whose level is at least `level`. */
export function outlineGroups(levels: Uint8Array, level: number): Array<{ start: number; end: number }> {
  const out: Array<{ start: number; end: number }> = [];
  let start = -1;
  for (let i = 0; i < levels.length; i++) {
    if (levels[i] >= level) {
      if (start < 0) start = i;
    } else if (start >= 0) {
      out.push({ start, end: i - 1 });
      start = -1;
    }
  }
  if (start >= 0) out.push({ start, end: levels.length - 1 });
  return out;
}

export class SheetGeometry {
  readonly rowCount: number;
  readonly colCount: number;
  /** Outline level per row/column (0 = not grouped). */
  readonly rowLevel: Uint8Array;
  readonly colLevel: Uint8Array;
  readonly maxRowLevel: number;
  readonly maxColLevel: number;
  readonly summaryBelow: boolean;
  readonly summaryRight: boolean;
  readonly rowTop: Float64Array; // rowTop[i] = y of row i; length rowCount+1
  readonly colLeft: Float64Array;
  readonly merges: MergeRect[];
  /** (row0,col0) → merge rect for every cell covered by a merge. */
  readonly mergeAt: Map<number, MergeRect>;
  readonly frozenRows: number;
  readonly frozenCols: number;

  /**
   * `overrides` are session-local sizes in px (0 = hidden) layered over
   * the file's: header drags and the agent's width/height/visibility ops.
   */
  constructor(readonly layout: EngineLayoutSheet, overrides?: SizeOverrides) {
    this.rowCount = Math.max(MIN_ROWS, layout.maxRow + ROW_SLACK);
    this.colCount = Math.max(MIN_COLS, layout.maxCol + COL_SLACK);
    const rowH = new Float64Array(this.rowCount).fill(DEFAULT_ROW_PX);
    for (const r of layout.rows) {
      const i = r.r - 1;
      if (i < 0 || i >= this.rowCount) continue;
      rowH[i] = r.hidden ? 0 : rowHeightPx(r.height);
    }
    const colW = new Float64Array(this.colCount).fill(DEFAULT_COL_PX);
    for (const c of layout.cols) {
      for (let i = c.min - 1; i <= c.max - 1 && i < this.colCount; i++) {
        if (i < 0) continue;
        colW[i] = c.hidden ? 0 : colWidthPx(c.width);
      }
    }
    if (overrides) {
      for (const [i, px] of overrides.rows) if (i >= 0 && i < this.rowCount) rowH[i] = px;
      for (const [i, px] of overrides.cols) if (i >= 0 && i < this.colCount) colW[i] = px;
    }
    this.rowTop = prefix(rowH);
    this.colLeft = prefix(colW);
    this.rowLevel = new Uint8Array(this.rowCount);
    this.colLevel = new Uint8Array(this.colCount);
    let maxRL = 0, maxCL = 0;
    for (const [r, l] of layout.rowOutline ?? []) { if (r >= 1 && r <= this.rowCount) { this.rowLevel[r - 1] = l; if (l > maxRL) maxRL = l; } }
    for (const [c, l] of layout.colOutline ?? []) { if (c >= 1 && c <= this.colCount) { this.colLevel[c - 1] = l; if (l > maxCL) maxCL = l; } }
    this.maxRowLevel = Math.min(7, maxRL);
    this.maxColLevel = Math.min(7, maxCL);
    this.summaryBelow = layout.summaryBelow !== false;
    this.summaryRight = layout.summaryRight !== false;
    this.frozenRows = Math.max(0, (overrides?.frozen ? overrides.frozen.rows : layout.frozenRows) | 0);
    this.frozenCols = Math.max(0, (overrides?.frozen ? overrides.frozen.cols : layout.frozenColumns) | 0);
    this.merges = [];
    this.mergeAt = new Map();
    for (const m of overrides?.merges ?? layout.merges) {
      const [a, b] = m.split(":");
      const p = parseA1(a);
      const q = b ? parseA1(b) : p;
      if (!p || !q) continue;
      const rect = { r0: Math.min(p.row0, q.row0), c0: Math.min(p.col0, q.col0), r1: Math.max(p.row0, q.row0), c1: Math.max(p.col0, q.col0) };
      this.merges.push(rect);
      for (let r = rect.r0; r <= rect.r1; r++) for (let c = rect.c0; c <= rect.c1; c++) this.mergeAt.set(r * 1_000_000 + c, rect);
    }
  }

  get totalHeight(): number {
    return this.rowTop[this.rowCount];
  }
  get totalWidth(): number {
    return this.colLeft[this.colCount];
  }
  rowHeight(r: number): number {
    return this.rowTop[r + 1] - this.rowTop[r];
  }
  /** The file's own size for a row/column (ignoring hidden), for unhide. */
  naturalRowPx(r: number): number {
    const e = this.layout.rows.find((x) => x.r === r + 1);
    return e && e.custom && e.height > 0 ? rowHeightPx(e.height) : DEFAULT_ROW_PX;
  }
  naturalColPx(c: number): number {
    const e = this.layout.cols.find((x) => c + 1 >= x.min && c + 1 <= x.max);
    return e && e.width > 0 ? colWidthPx(e.width) : DEFAULT_COL_PX;
  }
  colWidth(c: number): number {
    return this.colLeft[c + 1] - this.colLeft[c];
  }
  /** Row containing y (sheet pixel space, no header), clamped. */
  rowAt(y: number): number {
    return upperBound(this.rowTop, y, this.rowCount);
  }
  colAt(x: number): number {
    return upperBound(this.colLeft, x, this.colCount);
  }
  mergeOf(r: number, c: number): MergeRect | undefined {
    return this.mergeAt.get(r * 1_000_000 + c);
  }
}

function prefix(sizes: Float64Array): Float64Array {
  const out = new Float64Array(sizes.length + 1);
  let acc = 0;
  for (let i = 0; i < sizes.length; i++) {
    out[i] = acc;
    acc += sizes[i];
  }
  out[sizes.length] = acc;
  return out;
}

/** Largest index i in [0, n) with starts[i] <= v (binary search). */
function upperBound(starts: Float64Array, v: number, n: number): number {
  if (v <= 0) return 0;
  let lo = 0;
  let hi = n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (starts[mid] <= v) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

// ---------------------------------------------------------------------------
// Cells and tiles
// ---------------------------------------------------------------------------

export interface CfOverlay {
  fill?: string | null;
  color?: string | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  dataBar?: { value: number; axis: number; positive: string | null; negative: string | null; gradient: boolean; showValue: boolean };
}

export interface GridCell {
  /** Conditional formatting that matched (fill/font deltas, data bar). */
  cf?: CfOverlay;
  v: unknown; // raw value (number | string | boolean | null)
  d?: string; // display text for numbers/booleans
  f?: string; // formula with leading '='
  s?: number; // style index
  fileSaved?: boolean;
}

export const TILE_ROWS = 64;
/** 64 wide: sideways scrolling crosses a tile boundary half as often. */
export const TILE_COLS = 64;

export const tileKey = (sheet: number, tr: number, tc: number) => `${sheet}:${tr}:${tc}`;

/**
 * Tiles of cells keyed by (sheet, tileRow, tileCol). A tile is either
 * loaded (a Map of cells by `row0*1e6+col0`, possibly empty) or pending.
 * Least-recently-used tiles beyond `capacity` are evicted on insert, which
 * is what keeps memory flat in workbook size.
 */
export class TileCache {
  private tiles = new Map<string, Map<number, GridCell>>();
  private pending = new Set<string>();
  constructor(readonly capacity = 400) {}

  get size(): number {
    return this.tiles.size;
  }

  has(key: string): boolean {
    return this.tiles.has(key);
  }
  isPending(key: string): boolean {
    return this.pending.has(key);
  }
  markPending(key: string): void {
    this.pending.add(key);
  }
  clearPending(key: string): void {
    this.pending.delete(key);
  }

  cell(sheet: number, row0: number, col0: number): GridCell | undefined {
    const key = tileKey(sheet, Math.floor(row0 / TILE_ROWS), Math.floor(col0 / TILE_COLS));
    const t = this.tiles.get(key);
    if (!t) return undefined;
    // Touch for LRU.
    this.tiles.delete(key);
    this.tiles.set(key, t);
    return t.get(row0 * 1_000_000 + col0);
  }

  /** Tile keys covering a 0-based inclusive cell rectangle that are neither loaded nor pending. */
  missing(sheet: number, r0: number, r1: number, c0: number, c1: number): Array<{ key: string; tr: number; tc: number }> {
    const out: Array<{ key: string; tr: number; tc: number }> = [];
    for (let tr = Math.floor(r0 / TILE_ROWS); tr <= Math.floor(r1 / TILE_ROWS); tr++) {
      for (let tc = Math.floor(c0 / TILE_COLS); tc <= Math.floor(c1 / TILE_COLS); tc++) {
        const key = tileKey(sheet, tr, tc);
        if (!this.tiles.has(key) && !this.pending.has(key)) out.push({ key, tr, tc });
      }
    }
    return out;
  }

  /** Pending tile keys inside a 0-based inclusive cell rectangle. */
  pendingIn(sheet: number, r0: number, r1: number, c0: number, c1: number): string[] {
    const out: string[] = [];
    for (let tr = Math.floor(r0 / TILE_ROWS); tr <= Math.floor(r1 / TILE_ROWS); tr++)
      for (let tc = Math.floor(c0 / TILE_COLS); tc <= Math.floor(c1 / TILE_COLS); tc++) {
        const key = tileKey(sheet, tr, tc);
        if (this.pending.has(key)) out.push(key);
      }
    return out;
  }

  /** Store a fetched window as one tile (the window must be exactly the tile's rectangle). */
  putTile(key: string, cells: EngineWindowCell[]): void {
    const m = new Map<number, GridCell>();
    for (const c of cells) {
      const row0 = c.r - 1;
      const col0 = c.c - 1;
      const gc: GridCell = { v: c.v ?? null };
      if (c.d !== undefined) gc.d = c.d;
      if (c.f !== undefined) gc.f = c.f;
      if (c.s !== undefined) gc.s = c.s;
      if (c.fileSaved) gc.fileSaved = true;
      if ((c as any).cf) gc.cf = (c as any).cf;
      m.set(row0 * 1_000_000 + col0, gc);
    }
    this.pending.delete(key);
    this.tiles.set(key, m);
    while (this.tiles.size > this.capacity) {
      const oldest = this.tiles.keys().next().value as string | undefined;
      if (oldest === undefined) break;
      this.tiles.delete(oldest);
    }
  }

  /** Engine deltas: update loaded tiles in place; untouched tiles stay. */
  applyValues(sheet: number, cells: Array<{ row0: number; col0: number; v: unknown; d?: string }>): void {
    for (const c of cells) {
      const key = tileKey(sheet, Math.floor(c.row0 / TILE_ROWS), Math.floor(c.col0 / TILE_COLS));
      const t = this.tiles.get(key);
      if (!t) continue;
      const k = c.row0 * 1_000_000 + c.col0;
      const prev = t.get(k);
      if (prev) {
        prev.v = c.v;
        if (c.d !== undefined) prev.d = c.d;
        else delete prev.d;
      } else {
        t.set(k, { v: c.v, ...(c.d !== undefined ? { d: c.d } : {}) });
      }
    }
  }

  /** A write the grid made itself: set content immediately (optimistic). */
  setLocal(sheet: number, row0: number, col0: number, input: string): void {
    const key = tileKey(sheet, Math.floor(row0 / TILE_ROWS), Math.floor(col0 / TILE_COLS));
    const t = this.tiles.get(key);
    if (!t) return;
    const k = row0 * 1_000_000 + col0;
    const prev = t.get(k) ?? { v: null };
    if (input.startsWith("=")) {
      prev.f = input;
      delete prev.d;
    } else {
      delete prev.f;
      delete prev.d;
      prev.v = coerce(input);
    }
    t.set(k, prev);
  }

  /** Forget one tile so the next draw refetches it (engine values arrived without display strings). */
  dropTile(key: string): void {
    this.tiles.delete(key);
    this.pending.delete(key);
  }

  dropSheet(sheet: number): void {
    for (const key of [...this.tiles.keys()]) if (key.startsWith(`${sheet}:`)) this.tiles.delete(key);
  }
  clear(): void {
    this.tiles.clear();
    this.pending.clear();
  }
}

export function coerce(input: string): unknown {
  if (input === "") return null;
  if (input === "TRUE") return true;
  if (input === "FALSE") return false;
  if (/^-?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(input)) return Number(input);
  return input;
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

export interface ResolvedStyle {
  fill: string | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: string;
  /** Font size in points; the canvas font string uses pt, as Univer's does. */
  sizePt: number;
  family: string;
  vertical: "top" | "middle" | "bottom";
  align: "left" | "center" | "right" | null;
  wrap: boolean;
  borderBottom: boolean;
  borderTop: boolean;
  borderLeft: boolean;
  borderRight: boolean;
  /** Per-side style name (thin, medium, thick, double, dotted, dashed…) and colour. */
  borders: { top?: BorderLine; bottom?: BorderLine; left?: BorderLine; right?: BorderLine };
}

export interface BorderLine {
  style: string;
  color: string;
}

/** Canvas line width and dash pattern for an Excel border style name. */
export function borderStroke(style: string): { width: number; dash: number[]; double: boolean } {
  const s = style.toLowerCase();
  const width = s.includes("thick") ? 3 : s.includes("medium") ? 2 : 1;
  const dash = s.includes("dashdot") ? [6, 2, 2, 2] : s.includes("dash") ? [6, 3] : s.includes("dot") || s === "hair" ? [1, 2] : [];
  return { width, dash, double: s === "double" };
}

// Univer's defaults (DEFAULT_STYLES): Arial 11pt. Calibri is not a system
// font on macOS, so naming it first only changes the fallback order.
const DEFAULT_STYLE: ResolvedStyle = {
  fill: null, bold: false, italic: false, underline: false, strike: false, color: "#000000", sizePt: 11, family: "Arial",
  vertical: "bottom", align: null, wrap: false, borderBottom: false, borderTop: false, borderLeft: false, borderRight: false, borders: {},
};

/** Canvas font string for a resolved style, in the same form Univer builds. */
export function fontString(st: ResolvedStyle): string {
  return `${st.italic ? "italic " : ""}${st.bold ? "bold " : ""}${Math.ceil(st.sizePt)}pt ${st.family}, Arial, sans-serif`;
}

export class StyleBook {
  private resolved = new Map<number, ResolvedStyle>();
  constructor(private raw: Record<string, EngineStyle> = {}) {}

  merge(more: Record<string, EngineStyle>): void {
    for (const k of Object.keys(more)) {
      if (!(k in this.raw)) {
        this.raw[k] = more[k];
        this.resolved.delete(Number(k));
      }
    }
  }

  /** The engine's style record, for number formats and other raw fields. */
  rawOf(id: number | undefined): EngineStyle | undefined {
    return id === undefined ? undefined : this.raw[String(id)];
  }

  get(id: number | undefined): ResolvedStyle {
    if (id === undefined || id === 0 && !this.raw["0"]) return DEFAULT_STYLE;
    const cached = this.resolved.get(id);
    if (cached) return cached;
    const s = this.raw[String(id)];
    if (!s) return DEFAULT_STYLE;
    // A solid white fill is a real fill: Excel paints it over the gridlines,
    // which is how models get their gridline-free look without turning
    // gridlines off.
    const fill = s.fill?.color ? s.fill.color : null;
    const h = s.alignment?.horizontal;
    const r: ResolvedStyle = {
      fill,
      bold: !!s.font?.b,
      italic: !!s.font?.i,
      underline: !!s.font?.u,
      strike: !!s.font?.strike,
      color: s.font?.color && s.font.color !== "" ? s.font.color : DEFAULT_STYLE.color,
      sizePt: s.font?.sz ? s.font.sz : DEFAULT_STYLE.sizePt,
      family: s.font?.name ? `"${s.font.name}"` : DEFAULT_STYLE.family,
      vertical: s.alignment?.vertical === "top" ? "top" : s.alignment?.vertical === "center" ? "middle" : "bottom",
      align: h === "center" || h === "right" || h === "left" ? h : null,
      wrap: !!s.alignment?.wrap_text,
      borderBottom: !!s.border?.bottom?.style,
      borderTop: !!s.border?.top?.style,
      borderLeft: !!s.border?.left?.style,
      borderRight: !!s.border?.right?.style,
      borders: {
        ...(s.border?.top?.style ? { top: { style: s.border.top.style, color: s.border.top.color || "#000000" } } : {}),
        ...(s.border?.bottom?.style ? { bottom: { style: s.border.bottom.style, color: s.border.bottom.color || "#000000" } } : {}),
        ...(s.border?.left?.style ? { left: { style: s.border.left.style, color: s.border.left.color || "#000000" } } : {}),
        ...(s.border?.right?.style ? { right: { style: s.border.right.style, color: s.border.right.color || "#000000" } } : {}),
      },
    };
    this.resolved.set(id, r);
    return r;
  }
}

/** Text a cell shows: display string for numbers/booleans, else the value. */
export function cellText(cell: GridCell | undefined): string {
  if (!cell) return "";
  if (cell.d !== undefined && cell.d !== null) return String(cell.d);
  const v = cell.v;
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return formatNumberFallback(v);
  return String(v);
}

function formatNumberFallback(n: number): string {
  if (Number.isInteger(n)) return String(n);
  const s = n.toPrecision(10);
  return String(Number(s));
}

/** Numbers and booleans right-align by default; text left; explicit alignment wins. */
export function cellAlign(cell: GridCell | undefined, style: ResolvedStyle): "left" | "center" | "right" {
  if (style.align) return style.align;
  if (!cell) return "left";
  const v = cell.v;
  if (typeof v === "number" || typeof v === "boolean") return "right";
  if (typeof v === "string" && v.startsWith("#")) return "center";
  return "left";
}
