/**
 * Apply an agent tool's mutations to the grid, recording each one's
 * pre-state (old values, formats, widths, deleted bands, sheet snapshots…)
 * so `reverseMutation` can undo it on reject. The order and buffering follow
 * the desktop's apply loop (SpreadsheetScreen): cell writes go out as one
 * setCells per sheet run, formats after them as exact rectangles, and every
 * structural op flushes both first (it may shift coordinates).
 *
 * The viewport grid only holds the tiles near the viewport, so the cells a
 * pre-state read needs are prefetched from the engine first.
 */
import type { CellFormat, FormatMutation, UniverMutation } from "../../../gridpath/src/core/types";
import type { UniverGridHandle } from "../../../../src/screens/SpreadsheetScreen/components/UniverGrid";
import { cellRuns } from "../../../../src/screens/SpreadsheetScreen/formatRuns";

type Box = { r0: number; c0: number; r1: number; c1: number };

/** Bounding boxes per sheet of the cells the mutations read or write (0-based). */
function boxesOf(mutations: UniverMutation[]): Map<string, Box> {
  const boxes = new Map<string, Box>();
  const add = (sheet: string, row: number, col: number) => {
    const b = boxes.get(sheet);
    if (!b) boxes.set(sheet, { r0: row, c0: col, r1: row, c1: col });
    else {
      b.r0 = Math.min(b.r0, row);
      b.c0 = Math.min(b.c0, col);
      b.r1 = Math.max(b.r1, row);
      b.c1 = Math.max(b.c1, col);
    }
  };
  for (const m of mutations) {
    if (m.type === "set_cell") add(m.address.sheet, m.address.row, m.address.col);
    else if (m.type === "set_format" || m.type === "clear_range") for (const c of m.cells) add(m.sheet, c.row, c.col);
    else if (m.type === "set_note" || m.type === "delete_note") add(m.sheet, m.row, m.col);
  }
  return boxes;
}

async function prefetch(grid: UniverGridHandle, mutations: UniverMutation[]): Promise<void> {
  const jobs: Promise<void>[] = [];
  for (const [sheet, b] of boxesOf(mutations)) {
    // A pathologically sparse batch (A1 + ZZ9999) would fetch a huge
    // rectangle for a handful of cells: fetch per cell instead.
    if ((b.r1 - b.r0 + 1) * (b.c1 - b.c0 + 1) <= 200_000) jobs.push(grid.prefetchRange!(sheet, b.r0, b.c0, b.r1, b.c1));
    else
      for (const m of mutations)
        if (m.type === "set_cell" && m.address.sheet === sheet) jobs.push(grid.prefetchRange!(sheet, m.address.row, m.address.col, m.address.row, m.address.col));
  }
  // Deleted bands are captured for reject (same extents as the desktop).
  for (const m of mutations) {
    if (m.type === "delete_rows") jobs.push(grid.prefetchRange!(m.sheet, m.start, 0, m.start + m.count - 1, 80));
    if (m.type === "delete_columns") jobs.push(grid.prefetchRange!(m.sheet, 0, m.start, 2000, m.start + m.count - 1));
  }
  await Promise.all(jobs);
}

export type Applied = {
  mutations: UniverMutation[];
  /** Cells written, cleared or reformatted (for the pending highlight). */
  touched: Array<{ sheet: string; row: number; col: number }>;
};

export async function applyMutations(grid: UniverGridHandle, raw: UniverMutation[]): Promise<Applied> {
  await prefetch(grid, raw);
  const out: UniverMutation[] = [];
  const touched: Applied["touched"] = [];

  let cellBuf: { sheet: string; cells: Array<{ row: number; col: number; value: string | number | boolean | null }> } | null = null;
  const flushCells = () => {
    if (cellBuf?.cells.length) grid.setCells(cellBuf.sheet, cellBuf.cells);
    cellBuf = null;
  };
  // Formats wait behind the cells (copy_range "all" interleaves them) and
  // land as exact rectangles, never a bounding box.
  let formatBuf: FormatMutation[] = [];
  const flushFormats = () => {
    for (const fm of formatBuf) {
      const old_format = fm.cells.map(({ row, col }) => ({ row, col, format: grid.getCellFormat(fm.sheet, row, col) ?? null }));
      for (const r of cellRuns(fm.cells)) grid.setRangeFormat(fm.sheet, r.r0, r.c0, r.r1, r.c1, fm.new_format as CellFormat);
      for (const { row, col } of fm.cells) touched.push({ sheet: fm.sheet, row, col });
      out.push({ ...fm, old_format } as UniverMutation);
    }
    formatBuf = [];
  };

  for (const m of raw) {
    if (m.type === "set_cell") {
      const { sheet, row, col } = m.address;
      const before = grid.getCell(sheet, row, col);
      if (cellBuf && cellBuf.sheet !== sheet) flushCells();
      (cellBuf ??= { sheet, cells: [] }).cells.push({ row, col, value: m.new_formula ?? m.new_value });
      touched.push({ sheet, row, col });
      out.push({ ...m, old_value: before?.value ?? null, old_formula: before?.formula ?? null });
      continue;
    }
    if (m.type === "set_format") {
      formatBuf.push(m);
      continue;
    }
    flushCells();
    flushFormats();
    switch (m.type) {
      case "set_column_width": {
        const old_widths = m.columns.map((col) => ({ col, width: grid.getColumnWidth(m.sheet, col) ?? null }));
        for (const col of m.columns) grid.setColumnWidth(m.sheet, col, m.new_width);
        out.push({ ...m, old_widths });
        break;
      }
      case "set_row_height": {
        const old_heights = m.rows.map((row) => ({ row, height: grid.getRowHeight(m.sheet, row) ?? null }));
        for (const row of m.rows) grid.setRowHeight(m.sheet, row, m.new_height);
        out.push({ ...m, old_heights });
        break;
      }
      case "merge_cells":
        grid.mergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
        out.push(m);
        break;
      case "unmerge_cells":
        grid.unmergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
        out.push(m);
        break;
      case "create_sheet":
        if (grid.createSheet(m.name, m.tab_color ?? null)) out.push(m);
        break;
      case "delete_sheet": {
        const sheet_snapshot = grid.getSheetSnapshot(m.name) ?? null;
        if (grid.deleteSheet(m.name)) out.push({ ...m, sheet_snapshot });
        break;
      }
      case "rename_sheet":
        if (grid.renameSheet(m.old_name, m.new_name)) out.push(m);
        break;
      case "clear_range": {
        const cells = m.cells.map((c) => {
          const before = grid.getCell(m.sheet, c.row, c.col);
          return { row: c.row, col: c.col, old_value: before?.value ?? null, old_formula: before?.formula ?? null };
        });
        grid.setCells(m.sheet, cells.map((c) => ({ row: c.row, col: c.col, value: null })));
        for (const c of cells) touched.push({ sheet: m.sheet, row: c.row, col: c.col });
        out.push({ ...m, cells });
        break;
      }
      case "insert_rows":
        grid.insertRows(m.sheet, m.before, m.count);
        out.push(m);
        break;
      case "delete_rows": {
        const deleted_cells = grid.captureCellBand(m.sheet, m.start, 0, m.start + m.count - 1, 80) ?? [];
        grid.deleteRows(m.sheet, m.start, m.count);
        out.push({ ...m, deleted_cells });
        break;
      }
      case "insert_columns":
        grid.insertColumns(m.sheet, m.before, m.count);
        out.push(m);
        break;
      case "delete_columns": {
        const deleted_cells = grid.captureCellBand(m.sheet, 0, m.start, 2000, m.start + m.count - 1) ?? [];
        grid.deleteColumns(m.sheet, m.start, m.count);
        out.push({ ...m, deleted_cells });
        break;
      }
      case "freeze_panes": {
        const prior = grid.getFreezePanes(m.sheet) ?? { freezeRows: 0, freezeCols: 0 };
        grid.freezePanes(m.sheet, m.freeze_rows, m.freeze_cols);
        out.push({ ...m, old_freeze_rows: prior.freezeRows, old_freeze_cols: prior.freezeCols });
        break;
      }
      case "unfreeze_panes": {
        const prior = grid.getFreezePanes(m.sheet) ?? { freezeRows: 0, freezeCols: 0 };
        grid.unfreezePanes(m.sheet);
        out.push({ ...m, old_freeze_rows: prior.freezeRows, old_freeze_cols: prior.freezeCols });
        break;
      }
      case "set_note": {
        const old_text = grid.getNote(m.sheet, m.row, m.col) ?? null;
        grid.setNote(m.sheet, m.row, m.col, m.text);
        out.push({ ...m, old_text });
        break;
      }
      case "delete_note": {
        const old_text = grid.getNote(m.sheet, m.row, m.col) ?? null;
        grid.deleteNote(m.sheet, m.row, m.col);
        out.push({ ...m, old_text });
        break;
      }
      case "hide_rows":
        grid.hideRows(m.sheet, m.rows);
        out.push(m);
        break;
      case "show_rows":
        grid.showRows(m.sheet, m.rows);
        out.push(m);
        break;
      case "hide_columns":
        grid.hideColumns(m.sheet, m.columns);
        out.push(m);
        break;
      case "show_columns":
        grid.showColumns(m.sheet, m.columns);
        out.push(m);
        break;
      case "define_name": {
        const old_ref = grid.getDefinedNameRef(m.name) ?? null;
        grid.defineName(m.name, m.ref);
        out.push({ ...m, old_ref });
        break;
      }
    }
  }
  flushCells();
  flushFormats();
  return { mutations: out, touched };
}
