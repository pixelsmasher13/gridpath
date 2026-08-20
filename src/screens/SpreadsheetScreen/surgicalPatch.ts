/**
 * Surgical-save patch building (see docs/surgical-save-plan.md).
 *
 * Pure functions only — no React, no ExcelJS, no Univer imports — so this
 * module is trivially testable outside the app. UniverGrid captures the
 * baseline at load and calls buildWorkbookPatch at save.
 */

import type { SaveMirror } from "./components/UniverGrid";
import { addXlfnPrefixes } from "./xlfnCompat";

export type BaselineCell = { f?: string; v?: any };

/**
 * Why a patch could not be built. Each reason maps to specific user-facing
 * copy in the save flow — "structure changed" must never be shown for a
 * value-encoding problem.
 */
export type PatchFallbackReason =
  | "sheet_structure" // sheet set changed in a way ops don't explain
  | "row_col_structure" // row/col op couldn't be encoded
  | "shared_formula" // a Univer shared formula couldn't be resolved
  | "unsupported_value" // a cell value shape the patch format can't carry
  | "missing_baseline" // no load-time baseline (fresh/failed load)
  | "feature_edits"; // CF / data-validation / note state changed in-session

export type PatchBuildResult =
  | { ok: true; patch: any }
  | { ok: false; reason: PatchFallbackReason; detail?: string };

/** Reason-specific user copy — never a generic "structure changed". */
export function describePatchFallback(result: { reason: PatchFallbackReason }): string {
  switch (result.reason) {
    case "sheet_structure":
      return "The sheet list changed in a way the in-place saver couldn't verify";
    case "row_col_structure":
      return "A row/column change couldn't be encoded for in-place saving";
    case "shared_formula":
      return "A shared formula couldn't be resolved for in-place saving";
    case "unsupported_value":
      return "A cell holds a value type the in-place saver doesn't support yet";
    case "missing_baseline":
      return "The workbook baseline from load time is unavailable";
    case "feature_edits":
      return "Conditional formatting, data validation or cell-note changes can't be saved in place yet";
  }
}

export function buildCellBaseline(univerData: any): Map<string, Map<string, BaselineCell>> {
  const out = new Map<string, Map<string, BaselineCell>>();
  const order: string[] = univerData?.sheetOrder ?? Object.keys(univerData?.sheets ?? {});
  for (const id of order) {
    const sheet = univerData.sheets[id];
    if (!sheet) continue;
    const cells = new Map<string, BaselineCell>();
    const cellData = sheet.cellData ?? {};
    for (const rk of Object.keys(cellData)) {
      const row = cellData[rk];
      for (const ck of Object.keys(row)) {
        const norm = normalizeCellForDiff(row[ck]);
        if (norm) cells.set(`${rk},${ck}`, norm);
      }
    }
    out.set(sheet.name, cells);
  }
  return out;
}

/**
 * Collapse a cell to the two facts the file cares about: its formula (sans
 * leading '=') and its value. For formula cells the value is the evaluated
 * result, which the patch caches into the file as `<v>` — reopening in-app
 * then renders instantly instead of recomputing the whole dependency graph
 * (Excel recalculates on open regardless, via fullCalcOnLoad).
 */
function normalizeCellForDiff(cell: any): BaselineCell | null {
  if (!cell || typeof cell !== "object") return null;
  const rawF = cell.f;
  const f =
    typeof rawF === "string" && rawF.trim().replace(/^=/, "").length > 0
      ? rawF.trim().replace(/^=/, "")
      : undefined;
  let v = cell.v;
  if (cell.t === 3 && typeof v === "number") v = v !== 0; // Univer boolean cells store 0/1
  if (v === null || v === undefined || v === "") v = undefined;
  if (f === undefined && v === undefined) return null;
  return { f, v };
}

function cellChanged(a: BaselineCell | undefined, b: BaselineCell | null): boolean {
  const af = a?.f;
  const bf = b?.f;
  if (af !== undefined || bf !== undefined) {
    if (af !== bf) return true;
    // Same formula text but a drifted evaluated value is still a change:
    // the file caches formula results, and a stale cache next to a fresh
    // input would display wrong numbers on reopen. This also backfills
    // caches into files saved before values were cached (their baseline
    // v is undefined; the live snapshot has the computed value).
    return a?.v !== b?.v;
  }
  const av = a?.v;
  const bv = b?.v;
  if (av === undefined && bv === undefined) return false;
  return av !== bv;
}

export const ERROR_VALUES = new Set([
  "#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!",
]);

function encodeCellValue(v: any): any | null {
  if (typeof v === "number" && Number.isFinite(v)) return { t: "n", n: v };
  if (typeof v === "boolean") return { t: "b", b: v };
  if (typeof v === "string") {
    if (ERROR_VALUES.has(v)) return { t: "e", e: v };
    return { t: "s", s: v };
  }
  return null;
}

// ---------------------------------------------------------------------------
// A1 formula shifting (for resolving Univer shared formulas)
// ---------------------------------------------------------------------------

const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

function colLettersToIndex(letters: string): number | null {
  if (!letters || letters.length > 3) return null;
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const d = ch.charCodeAt(0) - 64;
    if (d < 1 || d > 26) return null;
    n = n * 26 + d;
  }
  return n > MAX_COL ? null : n - 1;
}

function colIndexToLetters(c: number): string {
  let out = "";
  for (;;) {
    out = String.fromCharCode(65 + (c % 26)) + out;
    if (c < 26) break;
    c = Math.floor(c / 26) - 1;
  }
  return out;
}

/**
 * Shift every relative A1 reference in `formula` by (dr, dc) — the same
 * translation Excel applies to shared-formula members. Quoted strings and
 * quoted sheet names are copied verbatim; identifiers (LOG10, MY_A1) are
 * never mistaken for refs. Returns null when a shifted ref leaves the sheet.
 */
export function shiftFormulaA1(formula: string, dr: number, dc: number): string | null {
  let out = "";
  let i = 0;
  const n = formula.length;
  const isIdent = (ch: string) => /[A-Za-z0-9_.]/.test(ch);
  while (i < n) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      // String literal / quoted sheet name: copy through doubled-quote escapes.
      const q = ch;
      let j = i + 1;
      while (j < n) {
        if (formula[j] === q) {
          if (formula[j + 1] === q) {
            j += 2;
            continue;
          }
          j += 1;
          break;
        }
        j += 1;
      }
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    const prevIsIdent = i > 0 && isIdent(formula[i - 1]);
    if (!prevIsIdent && /[$A-Za-z0-9]/.test(ch)) {
      const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/.exec(formula.slice(i));
      if (m) {
        const after = formula[i + m[0].length];
        const isRef = !(after !== undefined && (isIdent(after) || after === "("));
        const col = colLettersToIndex(m[2]);
        const row = Number(m[4]);
        if (isRef && col !== null && row >= 1 && row <= MAX_ROW) {
          const nc = m[1] ? col : col + dc;
          const nr = m[3] ? row : row + dr;
          if (nc < 0 || nc >= MAX_COL || nr < 1 || nr > MAX_ROW) return null;
          out += `${m[1]}${colIndexToLetters(nc)}${m[3]}${nr}`;
          i += m[0].length;
          continue;
        }
      }
      // Not a ref: consume the identifier/number as a unit.
      let j = i;
      while (j < n && (isIdent(formula[j]) || formula[j] === "$")) j += 1;
      if (j === i) j = i + 1;
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

// ---------------------------------------------------------------------------
// baseline transformation for structural ops
// ---------------------------------------------------------------------------

type RowColOpMirror = NonNullable<SaveMirror["rowColOps"]>[number];

/**
 * Move baseline cell keys the way a row/col op moves cells. Baseline
 * FORMULA TEXT is intentionally left alone: cells whose formulas Univer
 * adjusted across the shift show up as diffs and get rewritten with
 * Univer's (correct) adjusted text, which is exactly the end state.
 */
function shiftBaselineCells(
  cells: Map<string, BaselineCell>,
  op: RowColOpMirror,
): Map<string, BaselineCell> {
  const out = new Map<string, BaselineCell>();
  for (const [key, cell] of cells) {
    const [r, c] = key.split(",").map(Number);
    let nr = r;
    let nc = c;
    if (op.kind === "insertRows") {
      if (r >= op.before) nr = r + op.count;
    } else if (op.kind === "deleteRows") {
      if (r >= op.start && r < op.start + op.count) continue;
      if (r >= op.start + op.count) nr = r - op.count;
    } else if (op.kind === "insertColumns") {
      if (c >= op.before) nc = c + op.count;
    } else if (op.kind === "deleteColumns") {
      if (c >= op.start && c < op.start + op.count) continue;
      if (c >= op.start + op.count) nc = c - op.count;
    }
    out.set(`${nr},${nc}`, cell);
  }
  return out;
}

/**
 * Apply the mirror's structural ops to the load-time baseline so the cell
 * diff runs in live (post-op) coordinates, and produce the wire-shape op
 * lists for the Rust patcher.
 */
function transformBaseline(
  baseline: Map<string, Map<string, BaselineCell>>,
  mirror: SaveMirror | undefined,
): {
  base: Map<string, Map<string, BaselineCell>>;
  sheetOps: any[];
  rowColOps: any[];
  deletedSheets: Set<string>;
} {
  const base = new Map(
    [...baseline.entries()].map(([name, cells]) => [name, cells] as const),
  );
  const sheetOps: any[] = [];
  const createdSheets = new Set<string>();
  const deletedSheets = new Set<string>();
  const renameMap = new Map<string, string>(); // recorded-at name → live name

  for (const op of mirror?.sheetOps ?? []) {
    if (op.kind === "create") {
      sheetOps.push({
        op: "create",
        name: op.name,
        ...(typeof op.tabColor === "string" ? { tabColor: op.tabColor } : {}),
      });
      createdSheets.add(op.name);
      if (!base.has(op.name)) base.set(op.name, new Map());
    } else if (op.kind === "rename") {
      sheetOps.push({ op: "rename", oldName: op.oldName, newName: op.newName });
      const cells = base.get(op.oldName);
      if (cells) {
        base.delete(op.oldName);
        base.set(op.newName, cells);
      }
      if (createdSheets.delete(op.oldName)) createdSheets.add(op.newName);
      for (const [from, to] of renameMap) {
        if (to === op.oldName) renameMap.set(from, op.newName);
      }
      renameMap.set(op.oldName, op.newName);
    } else if (op.kind === "delete") {
      sheetOps.push({ op: "delete", name: op.name });
      base.delete(op.name);
      deletedSheets.add(op.name);
      createdSheets.delete(op.name);
    }
  }

  const rowColOps: any[] = [];
  for (const op of mirror?.rowColOps ?? []) {
    // Ops recorded before a later rename carry the old name; the Rust
    // patcher applies renames first, so translate to the live name.
    const sheet = renameMap.get(op.sheet) ?? op.sheet;
    if (op.kind === "insertRows") {
      rowColOps.push({ op: "insertRows", sheet, before: op.before, count: op.count });
    } else if (op.kind === "deleteRows") {
      rowColOps.push({ op: "deleteRows", sheet, start: op.start, count: op.count });
    } else if (op.kind === "insertColumns") {
      rowColOps.push({ op: "insertColumns", sheet, before: op.before, count: op.count });
    } else if (op.kind === "deleteColumns") {
      rowColOps.push({ op: "deleteColumns", sheet, start: op.start, count: op.count });
    }
    const cells = base.get(sheet);
    if (cells) base.set(sheet, shiftBaselineCells(cells, op));
  }

  return { base, sheetOps, rowColOps, deletedSheets };
}

/**
 * Univer shared formula: a cell carries `si` but no `f`; the text lives on
 * the group's master. Resolve it by translating the master's formula by the
 * cell's offset — the same derivation Excel and Univer use.
 */
function buildSharedFormulaIndex(cellData: any): Map<any, { r: number; c: number; f: string }> {
  const masters = new Map<any, { r: number; c: number; f: string }>();
  for (const rk of Object.keys(cellData ?? {})) {
    const row = cellData[rk];
    for (const ck of Object.keys(row)) {
      const raw = row[ck];
      if (raw && raw.si != null && typeof raw.f === "string" && raw.f && !masters.has(raw.si)) {
        masters.set(raw.si, { r: Number(rk), c: Number(ck), f: raw.f });
      }
    }
  }
  return masters;
}

/**
 * Diff the live snapshot against the load-time baseline and fold in the
 * SaveMirror's format/structure ops. Returns the wire-shape patch object,
 * or a typed reason when these edits need the full-export path.
 */
export function buildWorkbookPatch(
  snapshot: any,
  baseline: Map<string, Map<string, BaselineCell>>,
  mirror?: SaveMirror,
): PatchBuildResult {
  const { base: shiftedBaseline, sheetOps, rowColOps, deletedSheets } =
    transformBaseline(baseline, mirror);

  const sheetPatches = new Map<string, any>();
  const forSheet = (name: string) => {
    let sp = sheetPatches.get(name);
    if (!sp) {
      sp = { name };
      sheetPatches.set(name, sp);
    }
    return sp;
  };

  // --- cell diff (live coordinates; baseline already shifted) ---
  const order: string[] = snapshot?.sheetOrder ?? Object.keys(snapshot?.sheets ?? {});
  const liveNames = new Set<string>();
  for (const id of order) {
    const sheet = snapshot.sheets?.[id];
    if (!sheet) continue;
    liveNames.add(sheet.name);
    const base = shiftedBaseline.get(sheet.name);
    // A live sheet neither the file nor a create op accounts for: some
    // structural change slipped past the mirror — fall back rather than
    // guess.
    if (!base) {
      return {
        ok: false,
        reason: "sheet_structure",
        detail: `live sheet "${sheet.name}" has no baseline and no create op`,
      };
    }

    const cellData = sheet.cellData ?? {};
    const sharedMasters = buildSharedFormulaIndex(cellData);

    const cells: any[] = [];
    const seen = new Set<string>();
    for (const rk of Object.keys(cellData)) {
      const row = cellData[rk];
      for (const ck of Object.keys(row)) {
        let raw = row[ck];
        // Univer-side shared formula (si without f): derive the text from
        // the group master so these cells diff like any other formula cell.
        if (raw && raw.si != null && !(typeof raw.f === "string" && raw.f)) {
          const master = sharedMasters.get(raw.si);
          const derived = master
            ? shiftFormulaA1(master.f, Number(rk) - master.r, Number(ck) - master.c)
            : null;
          if (derived === null || derived === undefined) {
            return {
              ok: false,
              reason: "shared_formula",
              detail: `cell r${rk} c${ck} si=${raw.si} has no resolvable master`,
            };
          }
          raw = { ...raw, f: derived };
        }
        const cur = normalizeCellForDiff(raw);
        const key = `${rk},${ck}`;
        if (cur) seen.add(key);
        if (!cellChanged(base.get(key), cur)) continue;
        const out: any = { r: Number(rk), c: Number(ck) };
        if (cur?.f !== undefined) {
          // Restore the `_xlfn.` markers Excel requires for post-2007
          // functions (stripped at import so Univer's engine resolves them).
          out.f = addXlfnPrefixes(cur.f);
          // Cache the evaluated result alongside the formula. Best-effort:
          // an unencodable value costs a recalc on reopen, nothing more.
          if (cur.v !== undefined) {
            const enc = encodeCellValue(cur.v);
            if (enc) out.v = enc;
          }
        } else if (cur?.v !== undefined) {
          const enc = encodeCellValue(cur.v);
          if (!enc) {
            return {
              ok: false,
              reason: "unsupported_value",
              detail: `cell r${rk} c${ck} on "${sheet.name}" holds a ${typeof cur.v}`,
            };
          }
          out.v = enc;
        } else {
          out.clear = true;
        }
        cells.push(out);
      }
    }
    // Cells that existed in the file but are gone from the live model.
    for (const key of base.keys()) {
      if (!seen.has(key)) {
        const [r, c] = key.split(",").map(Number);
        cells.push({ r, c, clear: true });
      }
    }
    if (cells.length) forSheet(sheet.name).cells = cells;
  }
  // A baseline sheet missing from the live model without a delete op.
  for (const name of shiftedBaseline.keys()) {
    if (!liveNames.has(name) && !deletedSheets.has(name)) {
      return {
        ok: false,
        reason: "sheet_structure",
        detail: `baseline sheet "${name}" missing from the live model with no delete op`,
      };
    }
  }

  // --- styles (merged per cell so later ops win field-by-field, matching
  // the sequential ExcelJS mirror application) ---
  const styleByCell = new Map<string, any>();
  const styleFor = (sheet: string, row: number, col: number) => {
    const key = `${sheet} ${row},${col}`;
    let sp = styleByCell.get(key);
    if (!sp) {
      sp = { sheet, r: row, c: col };
      styleByCell.set(key, sp);
    }
    return sp;
  };
  for (const cf of mirror?.cellFormats ?? []) {
    const sp = styleFor(cf.sheet, cf.row, cf.col);
    const f = cf.format ?? {};
    if (f.bold !== undefined) sp.bold = f.bold;
    if (f.italic !== undefined) sp.italic = f.italic;
    if (f.underline !== undefined) sp.underline = f.underline;
    if (f.strike !== undefined) sp.strike = f.strike;
    if (f.font_color !== undefined) sp.fontColor = f.font_color;
    const bg = cf.background !== undefined ? cf.background : f.background_color;
    if (bg !== undefined) sp.backgroundColor = bg; // null clears the fill
    if (f.font_size !== undefined) sp.fontSize = f.font_size;
    if (f.font_family !== undefined) sp.fontFamily = f.font_family;
    if (f.horizontal_align !== undefined) sp.horizontalAlign = f.horizontal_align;
    if (f.vertical_align !== undefined) sp.verticalAlign = f.vertical_align;
    if (f.number_format !== undefined) sp.numberFormat = f.number_format;
    if (f.wrap_text !== undefined) sp.wrapText = f.wrap_text;
    if (f.indent !== undefined) sp.indent = f.indent;
  }
  for (const cb of mirror?.cellBorders ?? []) {
    const sp = styleFor(cb.sheet, cb.row, cb.col);
    sp.borders = sp.borders ?? {};
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const s = cb.borders?.[side];
      if (s === undefined) continue;
      sp.borders[side] = s === null ? null : { style: s.style, ...(s.color ? { color: s.color } : {}) };
    }
  }
  for (const sp of styleByCell.values()) {
    const { sheet, ...rest } = sp;
    (forSheet(sheet).styles = forSheet(sheet).styles ?? []).push(rest);
  }

  // --- sheet furniture from the mirror ---
  for (const cw of mirror?.columnWidths ?? []) {
    // Inverse of the import conversion (w = chars * 7 + 5).
    const chars = Math.max(0, (cw.widthPx - 5) / 7);
    (forSheet(cw.sheet).colWidths = forSheet(cw.sheet).colWidths ?? []).push({
      c: cw.col,
      chars: Math.round(chars * 100) / 100,
    });
  }
  for (const rh of mirror?.rowHeights ?? []) {
    // Inverse of the import conversion (px = pts * 1.333).
    (forSheet(rh.sheet).rowHeights = forSheet(rh.sheet).rowHeights ?? []).push({
      r: rh.row,
      pts: Math.round((rh.heightPx / 1.333) * 100) / 100,
    });
  }
  for (const vis of mirror?.visibility ?? []) {
    const sp = forSheet(vis.sheet);
    if (vis.kind === "hideRows" || vis.kind === "showRows") {
      sp.hiddenRows = sp.hiddenRows ?? [];
      for (const r of vis.rows) sp.hiddenRows.push({ r, hidden: vis.kind === "hideRows" });
    } else {
      sp.hiddenCols = sp.hiddenCols ?? [];
      for (const c of vis.columns) sp.hiddenCols.push({ c, hidden: vis.kind === "hideColumns" });
    }
  }
  for (const m of mirror?.merges ?? []) {
    (forSheet(m.sheet).merges = forSheet(m.sheet).merges ?? []).push({
      range: m.range,
      merge: m.merge,
    });
  }
  for (const fp of mirror?.freezePanes ?? []) {
    // Last write per sheet wins, matching the ExcelJS mirror.
    forSheet(fp.sheet).freeze = { rows: fp.freezeRows, cols: fp.freezeCols };
  }
  for (const af of mirror?.autoFilters ?? []) {
    forSheet(af.sheet).autoFilter = af.range; // string sets, null clears
  }

  const sheets = [...sheetPatches.values()];
  const definedNames = (mirror?.definedNames ?? []).map((d) => ({ name: d.name, ref: d.ref }));
  return {
    ok: true,
    patch: {
      version: 1,
      ...(sheets.length ? { sheets } : {}),
      ...(definedNames.length ? { definedNames } : {}),
      ...(sheetOps.length ? { sheetOps } : {}),
      ...(rowColOps.length ? { rowColOps } : {}),
    },
  };
}