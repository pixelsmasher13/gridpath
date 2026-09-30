/**
 * Collapse the mutations of the accepted + pending batches into the
 * `SaveMirror` the surgical patcher consumes. Pure; moved out of
 * SpreadsheetScreen.tsx so the headless package can reuse it.
 */
import type { SaveMirror } from "./types";

export function buildSaveMirror(batches: any[]): SaveMirror {
  const cellFormats: SaveMirror["cellFormats"] = [];
  const columnWidths: SaveMirror["columnWidths"] = [];
  const rowHeights: SaveMirror["rowHeights"] = [];
  const merges: SaveMirror["merges"] = [];
  const sheetOps: SaveMirror["sheetOps"] = [];
  const clears: SaveMirror["clears"] = [];
  const rowColOps: SaveMirror["rowColOps"] = [];
  const freezePanes: SaveMirror["freezePanes"] = [];
  const visibility: SaveMirror["visibility"] = [];
  // Defined names: last (accepted/pending) write per name wins, so collapse by name.
  const definedNameMap = new Map<string, { name: string; ref: string }>();
  for (const b of batches) {
    if (b.status !== "accepted" && b.status !== "pending") continue;
    // Already accounted for by the file + save baseline (replayed session
    // history, or batches consumed by a full-export save). Replaying them
    // would apply structural ops a second time.
    if (b.persisted) continue;
    for (const m of b.mutations ?? []) {
      if (m.type === "set_format") {
        // Save mirror keeps `background` as a sibling field on each cell
        // entry (the ExcelJS side reads it from there in applyStyleMirror).
        // Hoist background_color out of the format object so the saved
        // xlsx actually carries the fill, not just the in-app display.
        const bg = m.new_format?.background_color ?? null;
        for (const c of m.cells ?? []) {
          cellFormats!.push({
            sheet: m.sheet,
            row: c.row,
            col: c.col,
            format: m.new_format,
            background: bg,
          });
        }
      } else if (m.type === "set_column_width") {
        for (const col of m.columns ?? []) {
          columnWidths!.push({ sheet: m.sheet, col, widthPx: m.new_width });
        }
      } else if (m.type === "set_row_height") {
        for (const row of m.rows ?? []) {
          rowHeights!.push({ sheet: m.sheet, row, heightPx: m.new_height });
        }
      } else if (m.type === "merge_cells") {
        merges!.push({ sheet: m.sheet, range: m.range, merge: true });
      } else if (m.type === "unmerge_cells") {
        merges!.push({ sheet: m.sheet, range: m.range, merge: false });
      } else if (m.type === "create_sheet") {
        sheetOps!.push({ kind: "create", name: m.name, tabColor: m.tab_color });
      } else if (m.type === "delete_sheet") {
        sheetOps!.push({ kind: "delete", name: m.name });
      } else if (m.type === "rename_sheet") {
        sheetOps!.push({ kind: "rename", oldName: m.old_name, newName: m.new_name });
      } else if (m.type === "clear_range") {
        for (const c of m.cells ?? []) clears!.push({ sheet: m.sheet, row: c.row, col: c.col });
      } else if (m.type === "insert_rows") {
        rowColOps!.push({ kind: "insertRows", sheet: m.sheet, before: m.before, count: m.count });
      } else if (m.type === "delete_rows") {
        rowColOps!.push({ kind: "deleteRows", sheet: m.sheet, start: m.start, count: m.count });
      } else if (m.type === "insert_columns") {
        rowColOps!.push({ kind: "insertColumns", sheet: m.sheet, before: m.before, count: m.count });
      } else if (m.type === "delete_columns") {
        rowColOps!.push({ kind: "deleteColumns", sheet: m.sheet, start: m.start, count: m.count });
      } else if (m.type === "freeze_panes") {
        freezePanes!.push({ sheet: m.sheet, freezeRows: m.freeze_rows, freezeCols: m.freeze_cols });
      } else if (m.type === "unfreeze_panes") {
        freezePanes!.push({ sheet: m.sheet, freezeRows: 0, freezeCols: 0 });
      } else if (m.type === "hide_rows") {
        visibility!.push({ kind: "hideRows", sheet: m.sheet, rows: m.rows });
      } else if (m.type === "show_rows") {
        visibility!.push({ kind: "showRows", sheet: m.sheet, rows: m.rows });
      } else if (m.type === "hide_columns") {
        visibility!.push({ kind: "hideColumns", sheet: m.sheet, columns: m.columns });
      } else if (m.type === "show_columns") {
        visibility!.push({ kind: "showColumns", sheet: m.sheet, columns: m.columns });
      } else if (m.type === "define_name") {
        definedNameMap.set(m.name, { name: m.name, ref: m.ref });
      }
    }
  }
  return {
    cellFormats, columnWidths, rowHeights, merges, sheetOps, clears, rowColOps, freezePanes, visibility,
    definedNames: Array.from(definedNameMap.values()),
  };
}
