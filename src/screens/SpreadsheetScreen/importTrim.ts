/**
 * Import-time trimming of styled-but-empty sheet tails.
 *
 * Real-world files routinely carry huge bands of cells that have a style but
 * no content — the classic case is a column formatted all the way down to
 * Excel's last row (1,048,576), which writes a <row> element with one styled
 * blank <c> for every single row. Excel absorbs this because its parser is
 * native and its model stores styling sparsely; for us every one of those
 * cells became a JS object in the Univer snapshot, a styleOp, a baseline
 * entry and a rowData height record. One production model ("Share Tracker",
 * project-279.xlsx) carried ~1,048,400 such rows — 90 MB of XML, ~88% of all
 * cells in the workbook — and dominated open time and memory.
 *
 * The fix: compute the sheet's CONTENT extent — the last row/column holding
 * a value, a formula, a note, or a merged range — and drop style-only cells
 * (plus row heights / hidden flags / outline levels) beyond that extent plus
 * a small margin. The margin keeps banding visible immediately below/right
 * of the content so the sheet still looks right where users actually work.
 *
 * Save safety: trimming only affects the in-memory Univer model.
 *  - Surgical save diffs formula/value pairs against the load baseline;
 *    style-only cells were never part of either side (normalizeCellForDiff
 *    returns null for them), so no spurious patches appear.
 *  - Full save writes Univer cells INTO the retained ExcelJS workbook and
 *    never clears cells absent from the Univer snapshot, so the trimmed
 *    styling rides through to the file untouched.
 *
 * Pure module — no ExcelJS/Univer imports — so it's trivially testable.
 * The Worksheet/Row/Cell shapes below are the structural subset of ExcelJS
 * this module needs.
 */

/** Styled blanks within this many rows below the last content row are kept. */
export const TRIM_ROW_MARGIN = 50;
/** Styled blanks within this many columns right of the last content column are kept. */
export const TRIM_COL_MARGIN = 10;
/**
 * A sheet with styling but NO content anywhere (a blank formatted template)
 * has no extent to anchor on; keep its styling up to these caps instead of
 * nuking it entirely — or letting a million styled blank rows through.
 */
export const STYLE_ONLY_ROW_CAP = 500;
export const STYLE_ONLY_COL_CAP = 100;

export type CellLike = { value: unknown; note?: unknown };
export type RowLike = {
  eachCell(opts: { includeEmpty: boolean }, cb: (cell: CellLike, colNumber: number) => void): void;
};
export type WorksheetLike = {
  eachRow(opts: { includeEmpty: boolean }, cb: (row: RowLike, rowNumber: number) => void): void;
};

export type ContentExtent = {
  /** Last 0-indexed row/col with real content (value/formula/note/merge); -1 when none. */
  contentMaxRow: number;
  contentMaxCol: number;
  /** Last 0-indexed row/col with ANY cell at all, styled blanks included; -1 when none. */
  anyMaxRow: number;
  anyMaxCol: number;
};

export type TrimBounds = {
  /** Keep cells with row <= trimRow and col <= trimCol; drop the rest (they are style-only). */
  trimRow: number;
  trimCol: number;
};

/**
 * Walk the worksheet once with cheap property checks (no style extraction,
 * no formula translation) and record where real content ends. A cell counts
 * as content when it has a value (which covers formulas — their `value` is
 * a formula object) or a note; merged ranges count via `merges` because a
 * merge's structure is content even when every underlying cell is blank.
 */
export function computeContentExtent(
  ws: WorksheetLike,
  merges: Array<{ endRow: number; endColumn: number }> = [],
): ContentExtent {
  let contentMaxRow = -1;
  let contentMaxCol = -1;
  let anyMaxRow = -1;
  let anyMaxCol = -1;
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const r = rowNumber - 1;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const c = colNumber - 1;
      if (r > anyMaxRow) anyMaxRow = r;
      if (c > anyMaxCol) anyMaxCol = c;
      const v = cell.value;
      if ((v !== null && v !== undefined) || cell.note) {
        if (r > contentMaxRow) contentMaxRow = r;
        if (c > contentMaxCol) contentMaxCol = c;
      }
    });
  });
  for (const m of merges) {
    if (m.endRow > contentMaxRow) contentMaxRow = m.endRow;
    if (m.endColumn > contentMaxCol) contentMaxCol = m.endColumn;
  }
  return { contentMaxRow, contentMaxCol, anyMaxRow, anyMaxCol };
}

export function trimBoundsFor(extent: ContentExtent): TrimBounds {
  if (extent.contentMaxRow < 0 && extent.contentMaxCol < 0) {
    return {
      trimRow: Math.min(extent.anyMaxRow, STYLE_ONLY_ROW_CAP - 1),
      trimCol: Math.min(extent.anyMaxCol, STYLE_ONLY_COL_CAP - 1),
    };
  }
  return {
    trimRow: extent.contentMaxRow + TRIM_ROW_MARGIN,
    trimCol: extent.contentMaxCol + TRIM_COL_MARGIN,
  };
}
