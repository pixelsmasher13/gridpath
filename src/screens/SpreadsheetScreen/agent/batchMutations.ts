/**
 * Apply / reverse agent mutations against a UniverGrid handle.
 *
 * Reject, Undo, and Redo all share this so structural ops (insert/delete
 * rows & columns, sheet create/delete/rename, clear, freeze, hide) rewind
 * the same way cell values and formats do. Mutations must carry the
 * pre-state fields captured at apply time (old_*, deleted_*, sheet_snapshot,
 * old_freeze_*).
 */
import type { UniverMutation, CellFormat } from "../types";
import type { UniverGridHandle } from "../components/UniverGrid";

export type CapturedCell = {
  row: number;
  col: number;
  value: any;
  formula: string | null;
  format: CellFormat | null;
};

/** Apply one mutation's NEW state (forward). Used by live agent apply + Redo. */
export function applyMutationForward(
  grid: UniverGridHandle | null | undefined,
  m: UniverMutation,
): void {
  if (!grid) return;
  switch (m.type) {
    case "set_cell":
      grid.setCell(m.address.sheet, m.address.row, m.address.col, m.new_formula ?? m.new_value);
      break;
    case "set_format": {
      // One range-level command instead of every facade setter per cell —
      // new_format is uniform across the mutation, and m.cells comes from
      // expandA1Range so its bounding box IS the range. (Reverse stays
      // per-cell: old formats differ cell to cell.)
      if (m.cells.length === 0) break;
      let r0 = m.cells[0].row, c0 = m.cells[0].col, r1 = r0, c1 = c0;
      for (const c of m.cells) {
        if (c.row < r0) r0 = c.row;
        if (c.col < c0) c0 = c.col;
        if (c.row > r1) r1 = c.row;
        if (c.col > c1) c1 = c.col;
      }
      grid.setRangeFormat(m.sheet, r0, c0, r1, c1, m.new_format);
      break;
    }
    case "set_column_width":
      for (const col of m.columns) grid.setColumnWidth(m.sheet, col, m.new_width);
      break;
    case "set_row_height":
      for (const row of m.rows) grid.setRowHeight(m.sheet, row, m.new_height);
      break;
    case "merge_cells":
      grid.mergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
      break;
    case "unmerge_cells":
      grid.unmergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
      break;
    case "set_note":
      grid.setNote(m.sheet, m.row, m.col, m.text);
      break;
    case "delete_note":
      grid.deleteNote(m.sheet, m.row, m.col);
      break;
    case "create_sheet":
      grid.createSheet(m.name, m.tab_color ?? null);
      break;
    case "delete_sheet":
      grid.deleteSheet(m.name);
      break;
    case "rename_sheet":
      grid.renameSheet(m.old_name, m.new_name);
      break;
    case "clear_range":
      for (const c of m.cells) grid.setCell(m.sheet, c.row, c.col, null);
      break;
    case "insert_rows":
      grid.insertRows(m.sheet, m.before, m.count);
      break;
    case "delete_rows":
      grid.deleteRows(m.sheet, m.start, m.count);
      break;
    case "insert_columns":
      grid.insertColumns(m.sheet, m.before, m.count);
      break;
    case "delete_columns":
      grid.deleteColumns(m.sheet, m.start, m.count);
      break;
    case "freeze_panes":
      grid.freezePanes(m.sheet, m.freeze_rows, m.freeze_cols);
      break;
    case "unfreeze_panes":
      grid.unfreezePanes(m.sheet);
      break;
    case "hide_rows":
      grid.hideRows(m.sheet, m.rows);
      break;
    case "show_rows":
      grid.showRows(m.sheet, m.rows);
      break;
    case "hide_columns":
      grid.hideColumns(m.sheet, m.columns);
      break;
    case "show_columns":
      grid.showColumns(m.sheet, m.columns);
      break;
    case "define_name":
      grid.defineName(m.name, m.ref);
      break;
  }
}

/**
 * Reverse one mutation using its captured pre-state. Callers MUST iterate
 * mutations in reverse order so nested structural shifts unwind correctly.
 */
export function reverseMutation(
  grid: UniverGridHandle | null | undefined,
  m: UniverMutation,
): void {
  if (!grid) return;
  switch (m.type) {
    case "set_cell":
      grid.setCell(m.address.sheet, m.address.row, m.address.col, m.old_formula ?? m.old_value);
      break;
    case "set_format":
      for (const o of m.old_format) {
        if (o.format) grid.setCellFormat(m.sheet, o.row, o.col, o.format);
      }
      break;
    case "set_column_width":
      for (const o of m.old_widths) {
        if (o.width != null) grid.setColumnWidth(m.sheet, o.col, o.width);
      }
      break;
    case "set_row_height":
      for (const o of m.old_heights) {
        if (o.height != null) grid.setRowHeight(m.sheet, o.row, o.height);
      }
      break;
    case "merge_cells":
      grid.unmergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
      break;
    case "unmerge_cells":
      grid.mergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
      break;
    case "set_note":
      if (m.old_text != null) grid.setNote(m.sheet, m.row, m.col, m.old_text);
      else grid.deleteNote(m.sheet, m.row, m.col);
      break;
    case "delete_note":
      if (m.old_text != null) grid.setNote(m.sheet, m.row, m.col, m.old_text);
      break;
    case "create_sheet":
      grid.deleteSheet(m.name);
      break;
    case "delete_sheet":
      if (m.sheet_snapshot) grid.restoreSheetSnapshot(m.sheet_snapshot);
      else grid.createSheet(m.name);
      break;
    case "rename_sheet":
      grid.renameSheet(m.new_name, m.old_name);
      break;
    case "clear_range":
      for (const c of m.cells) {
        grid.setCell(m.sheet, c.row, c.col, c.old_formula ?? c.old_value);
      }
      break;
    case "insert_rows":
      // Inverse of insert at `before` is delete the same band.
      grid.deleteRows(m.sheet, m.before, m.count);
      break;
    case "delete_rows": {
      grid.insertRows(m.sheet, m.start, m.count);
      restoreCapturedCells(grid, m.sheet, m.deleted_cells ?? []);
      break;
    }
    case "insert_columns":
      grid.deleteColumns(m.sheet, m.before, m.count);
      break;
    case "delete_columns": {
      grid.insertColumns(m.sheet, m.start, m.count);
      restoreCapturedCells(grid, m.sheet, m.deleted_cells ?? []);
      break;
    }
    case "freeze_panes": {
      const rows = m.old_freeze_rows ?? 0;
      const cols = m.old_freeze_cols ?? 0;
      if (rows === 0 && cols === 0) grid.unfreezePanes(m.sheet);
      else grid.freezePanes(m.sheet, rows, cols);
      break;
    }
    case "unfreeze_panes": {
      const rows = m.old_freeze_rows ?? 0;
      const cols = m.old_freeze_cols ?? 0;
      if (rows > 0 || cols > 0) grid.freezePanes(m.sheet, rows, cols);
      break;
    }
    case "hide_rows":
      grid.showRows(m.sheet, m.rows);
      break;
    case "show_rows":
      grid.hideRows(m.sheet, m.rows);
      break;
    case "hide_columns":
      grid.showColumns(m.sheet, m.columns);
      break;
    case "show_columns":
      grid.hideColumns(m.sheet, m.columns);
      break;
    case "define_name":
      if (m.old_ref) grid.defineName(m.name, m.old_ref);
      else grid.deleteName(m.name);
      break;
  }
}

function restoreCapturedCells(
  grid: UniverGridHandle,
  sheet: string,
  cells: CapturedCell[],
): void {
  for (const c of cells) {
    grid.setCell(sheet, c.row, c.col, c.formula ?? c.value);
    if (c.format) grid.setCellFormat(sheet, c.row, c.col, c.format);
  }
}

/**
 * A batch is "literal-only" when every mutation is a set_cell writing a
 * plain value (no formula). Such a batch cannot change any touched cell
 * through recalculation — the values read back are exactly the values just
 * written synchronously — so the post-apply engine-idle wait (two worker
 * RPC round-trips + 180ms of grace sleeps inside whenCalculated) is pure
 * overhead and is skipped. Formulas, formats, clears, or structure ops all
 * disqualify: they can leave touched cells with values that only settle
 * after the formula worker runs.
 */
export function isLiteralOnlyBatch(mutations: UniverMutation[]): boolean {
  return (
    mutations.length > 0 &&
    mutations.every((m) => m.type === "set_cell" && !m.new_formula)
  );
}

/** Excel formula-error tokens Univer surfaces as cell values. */
const FORMULA_ERROR_RE =
  /^(#REF!|#DIV\/0!|#VALUE!|#N\/A|#NAME\?|#NULL!|#NUM!|#GETTING_DATA|#SPILL!|#CALC!|#CIRCULAR!|#ERROR!)/i;

export type LayoutValidationError = {
  sheet: string;
  cell: string;
  formula: string | null;
  value: string;
  kind: "formula_error" | "self_ref";
};

/**
 * Scan cells touched by a batch for formula errors and obvious self-refs.
 * Used to gate `done` so the agent must fix layout/formula drift before the
 * batch is presented for Accept/Reject.
 */
export function validateBatchLayout(
  grid: UniverGridHandle | null | undefined,
  mutations: UniverMutation[],
): LayoutValidationError[] {
  if (!grid) return [];
  const errors: LayoutValidationError[] = [];
  const seen = new Set<string>();

  const check = (sheet: string, row: number, col: number) => {
    const key = `${sheet}!${row},${col}`;
    if (seen.has(key)) return;
    seen.add(key);
    const c = grid.getCell(sheet, row, col);
    if (!c) return;
    const addr = a1Of(row, col);
    const valStr = c.value == null ? "" : String(c.value);
    if (FORMULA_ERROR_RE.test(valStr)) {
      errors.push({
        sheet,
        cell: addr,
        formula: c.formula,
        value: valStr,
        kind: "formula_error",
      });
      return;
    }
    // Self-reference: formula mentions its own A1 address as a token.
    // The `!` in the leading class is load-bearing: a SHEET-QUALIFIED
    // reference to the same address on another sheet ('Income Statement'!H8
    // written in Assumptions!H8) is an ordinary cross-sheet link, not a
    // cycle. Without excluding `!` the gate rejects every such formula —
    // which multi-sheet models produce constantly, since parallel sheets
    // naturally line up on the same rows/columns. Trade-off: a same-sheet
    // qualified self-ref (Sheet1!H8 inside Sheet1!H8) is no longer caught
    // here, but the engine resolves it to a formula error, which the
    // FORMULA_ERROR_RE branch above already rejects.
    if (c.formula) {
      const selfRe = new RegExp(`(^|[^A-Z0-9_!])${addr}(?![0-9A-Z])`, "i");
      if (selfRe.test(c.formula)) {
        errors.push({
          sheet,
          cell: addr,
          formula: c.formula,
          value: valStr || "(self-ref)",
          kind: "self_ref",
        });
      }
    }
  };

  for (const m of mutations) {
    if (m.type === "set_cell") {
      check(m.address.sheet, m.address.row, m.address.col);
    } else if (m.type === "set_format") {
      for (const c of m.cells) check(m.sheet, c.row, c.col);
    } else if (m.type === "clear_range") {
      for (const c of m.cells) check(m.sheet, c.row, c.col);
    }
  }
  return errors;
}

function a1Of(row: number, col: number): string {
  let n = col;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return `${letters}${row + 1}`;
}
