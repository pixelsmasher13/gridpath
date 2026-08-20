import type { CellMutation, UniverMutation } from "../types";

/**
 * Pure client-side grouping of a batch's mutations for the review modal.
 *
 * No model involvement: regions are contiguous row bands computed from the
 * mutation list (microseconds for hundreds of edits), and the optional
 * human label for a region comes from a local grid read supplied by the
 * caller (same leftmost-text-label trick as the agent's row_map readback).
 */

export type ReviewCell = {
  /** Position of this mutation inside batch.mutations (for reject bookkeeping). */
  sheet: string;
  row: number;
  col: number;
  addr: string;
  oldValue: string;
  newValue: string;
  isFormula: boolean;
  /** This cell's row label (leftmost text cell of the row), or null. */
  rowLabel: string | null;
  /** This cell's column header (topmost text/date cell of the column), or null. */
  header: string | null;
};

export type ReviewRegion = {
  sheet: string;
  /** Bounding-box A1 range of the region, e.g. "BU73:BU148". */
  range: string;
  /** Row-label inferred from the sheet (e.g. "Revenue build"), or null. */
  label: string | null;
  cells: ReviewCell[];
};

export type ReviewSummaryRow = { desc: string };

export type ReviewGroups = {
  regions: ReviewRegion[];
  /** Non-cell mutations (structure, formats, dimensions) as one-line summaries. */
  others: ReviewSummaryRow[];
  countsBySheet: Record<string, number>;
  totalCells: number;
};

/** Split regions when consecutive edited rows are further apart than this. */
const ROW_GAP = 10;

export function colLetters(col: number): string {
  let n = col;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

export function a1Of(row: number, col: number): string {
  return `${colLetters(col)}${row + 1}`;
}

function fmtVal(v: any): string {
  if (v === null || v === undefined || v === "") return "";
  return String(v);
}

/**
 * Group a batch's mutations into review regions.
 *
 * `labelOf(sheet, row)` and `headerOf(sheet, col)` may read the grid to
 * produce human context (row label / column header) per edited cell. Both
 * are memoized here per row/column, so grid reads stay proportional to the
 * batch's distinct rows + columns, not its cell count. Pass undefined for
 * pure range labels.
 */
export function groupBatchForReview(
  mutations: UniverMutation[],
  labelOf?: (sheet: string, row: number) => string | null,
  headerOf?: (sheet: string, col: number) => string | null,
): ReviewGroups {
  const rowLabelCache = new Map<string, string | null>();
  const rowLabelAt = (sheet: string, row: number): string | null => {
    if (!labelOf) return null;
    const key = `${sheet}!${row}`;
    let v = rowLabelCache.get(key);
    if (v === undefined) {
      v = labelOf(sheet, row);
      rowLabelCache.set(key, v);
    }
    return v;
  };
  const headerCache = new Map<string, string | null>();
  const headerAt = (sheet: string, col: number): string | null => {
    if (!headerOf) return null;
    const key = `${sheet}!${col}`;
    let v = headerCache.get(key);
    if (v === undefined) {
      v = headerOf(sheet, col);
      headerCache.set(key, v);
    }
    return v;
  };
  const cellMuts: CellMutation[] = [];
  const others: ReviewSummaryRow[] = [];

  for (const m of mutations) {
    if (m.type === "set_cell") cellMuts.push(m);
    else {
      const desc = describeMutation(m);
      if (desc) others.push({ desc });
    }
  }

  const countsBySheet: Record<string, number> = {};
  for (const m of cellMuts) {
    countsBySheet[m.address.sheet] = (countsBySheet[m.address.sheet] ?? 0) + 1;
  }

  // Cluster per sheet into row bands: sort by row, split on gaps > ROW_GAP.
  // Row bands (rather than column runs) match both write shapes we see:
  // a column extension (one col, many rows) stays one region per section,
  // and a set_range block (many cols, adjacent rows) collapses to one region.
  const bySheet = new Map<string, CellMutation[]>();
  for (const m of cellMuts) {
    const arr = bySheet.get(m.address.sheet) ?? [];
    arr.push(m);
    bySheet.set(m.address.sheet, arr);
  }

  const regions: ReviewRegion[] = [];
  for (const [sheet, muts] of bySheet) {
    muts.sort((x, y) => x.address.row - y.address.row || x.address.col - y.address.col);
    let band: CellMutation[] = [];
    const flush = () => {
      if (band.length === 0) return;
      let minRow = Infinity;
      let maxRow = -1;
      let minCol = Infinity;
      let maxCol = -1;
      for (const m of band) {
        minRow = Math.min(minRow, m.address.row);
        maxRow = Math.max(maxRow, m.address.row);
        minCol = Math.min(minCol, m.address.col);
        maxCol = Math.max(maxCol, m.address.col);
      }
      const range =
        minRow === maxRow && minCol === maxCol
          ? a1Of(minRow, minCol)
          : `${a1Of(minRow, minCol)}:${a1Of(maxRow, maxCol)}`;
      regions.push({
        sheet,
        range,
        label: rowLabelAt(sheet, minRow),
        cells: band.map((m) => ({
          sheet,
          row: m.address.row,
          col: m.address.col,
          addr: a1Of(m.address.row, m.address.col),
          oldValue: fmtVal(m.old_formula ?? m.old_value),
          newValue: fmtVal(m.new_formula ?? m.new_value),
          isFormula: !!m.new_formula,
          rowLabel: rowLabelAt(sheet, m.address.row),
          header: headerAt(sheet, m.address.col),
        })),
      });
      band = [];
    };
    for (const m of muts) {
      const prev = band[band.length - 1];
      if (prev && m.address.row - prev.address.row > ROW_GAP) flush();
      band.push(m);
    }
    flush();
  }

  return { regions, others, countsBySheet, totalCells: cellMuts.length };
}

/** One-line human description of a non-cell mutation for the review's structural list. */
export function describeMutation(m: UniverMutation): string | null {
  switch (m.type) {
    case "set_format": {
      const n = m.cells.length;
      return `Format ${m.sheet}!${m.range} (${n} cell${n === 1 ? "" : "s"})`;
    }
    case "set_column_width":
      return `Column width ${m.columns.map(colLetters).join(", ")} → ${m.new_width}px (${m.sheet})`;
    case "set_row_height":
      return `Row height ${m.rows.map((r) => r + 1).join(", ")} → ${m.new_height}px (${m.sheet})`;
    case "merge_cells":
      return `Merge ${m.sheet}!${m.range}`;
    case "unmerge_cells":
      return `Unmerge ${m.sheet}!${m.range}`;
    case "create_sheet":
      return `New sheet "${m.name}"`;
    case "delete_sheet":
      return `Delete sheet "${m.name}"`;
    case "rename_sheet":
      return `Rename sheet "${m.old_name}" → "${m.new_name}"`;
    case "clear_range":
      return `Clear ${m.sheet}!${m.range}`;
    case "insert_rows":
      return `Insert ${m.count} row${m.count === 1 ? "" : "s"} before row ${m.before + 1} (${m.sheet})`;
    case "delete_rows":
      return `Delete ${m.count} row${m.count === 1 ? "" : "s"} from row ${m.start + 1} (${m.sheet})`;
    case "insert_columns":
      return `Insert ${m.count} column${m.count === 1 ? "" : "s"} before ${colLetters(m.before)} (${m.sheet})`;
    case "delete_columns":
      return `Delete ${m.count} column${m.count === 1 ? "" : "s"} from ${colLetters(m.start)} (${m.sheet})`;
    case "freeze_panes":
      return `Freeze ${m.freeze_rows} row${m.freeze_rows === 1 ? "" : "s"} × ${m.freeze_cols} col${m.freeze_cols === 1 ? "" : "s"} (${m.sheet})`;
    case "unfreeze_panes":
      return `Unfreeze panes (${m.sheet})`;
    case "hide_rows":
      return `Hide rows ${m.rows.map((r) => r + 1).join(", ")} (${m.sheet})`;
    case "show_rows":
      return `Show rows ${m.rows.map((r) => r + 1).join(", ")} (${m.sheet})`;
    case "hide_columns":
      return `Hide columns ${m.columns.map(colLetters).join(", ")} (${m.sheet})`;
    case "show_columns":
      return `Show columns ${m.columns.map(colLetters).join(", ")} (${m.sheet})`;
    case "define_name":
      return `Define name ${m.name} = ${m.ref}`;
    case "set_range":
      return null; // legacy type, never emitted
    default:
      return null;
  }
}
