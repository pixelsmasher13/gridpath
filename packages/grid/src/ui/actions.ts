/**
 * What the toolbar's buttons do, over the grid handle: the desktop app's
 * Home-ribbon actions for the controls the package shows. Every change goes
 * through the grid's own undo history and save log.
 */
import type { CellFormat } from "../../../gridpath/src/core/types";
import type { UniverGridHandle } from "../../../../src/screens/SpreadsheetScreen/components/UniverGrid";
import type { FormatToolbarProps, InsertAction } from "../../../../src/screens/SpreadsheetScreen/components/FormatToolbar";

type Handlers = Omit<FormatToolbarProps, "disabled" | "hidden" | "canUndo" | "canRedo" | "trailing" | "painterActive" | "outline">;

/** Excel's number-format pattern with `delta` more (or fewer) decimals. */
export function adjustDecimalPlaces(pattern: string | undefined, delta: number): string {
  const main = (!pattern || pattern === "General" ? "0" : pattern).split(";")[0];
  const dot = main.indexOf(".");
  let intPart: string;
  let suffix: string;
  let current: number;
  if (dot >= 0) {
    intPart = main.slice(0, dot);
    const m = main.slice(dot + 1).match(/^(0*)(.*)$/);
    current = m ? m[1].length : 0;
    suffix = m ? m[2] : "";
  } else {
    const m = main.match(/^([#0,]*)(.*)$/);
    intPart = m ? m[1] : main;
    suffix = m ? m[2] : "";
    current = 0;
  }
  if (!intPart) intPart = "0";
  const next = Math.max(0, Math.min(10, current + delta));
  return `${intPart}${next > 0 ? "." + "0".repeat(next) : ""}${suffix}` || "0";
}

const FONT_SIZE_LADDER = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

export function stepFontSize(current: number, delta: 1 | -1): number {
  if (delta > 0) return FONT_SIZE_LADDER.find((s) => s > current) ?? Math.min(409, Math.round(current) + 1);
  return [...FONT_SIZE_LADDER].reverse().find((s) => s < current) ?? Math.max(1, Math.round(current) - 1);
}

/**
 * The toolbar's handlers for one grid. `notify` shows a short message
 * ("Select cells first"); `changed` tells the chrome to refresh (undo/redo
 * availability).
 */
export function toolbarActions(grid: () => UniverGridHandle | null, notify: (message: string) => void, changed: () => void): Handlers {
  const selection = (message = "Select cells first") => {
    const g = grid();
    const sel = g?.getActiveSelection?.() ?? null;
    if (!g || !sel) {
      notify(message);
      return null;
    }
    return { g, sel };
  };
  const format = (patch: Partial<CellFormat>) => {
    const s = selection();
    if (!s) return;
    s.g.setRangeFormat(s.sel.sheet, s.sel.startRow, s.sel.startCol, s.sel.endRow, s.sel.endCol, patch);
  };
  const anchorFormat = () => {
    const s = selection();
    return s ? s.g.getCellFormat(s.sel.sheet, s.sel.startRow, s.sel.startCol) : null;
  };
  const insert = (action: InsertAction) => {
    const s = selection();
    if (!s) return;
    const { g, sel } = s;
    const rows = sel.endRow - sel.startRow + 1;
    const cols = sel.endCol - sel.startCol + 1;
    const log = { log: true };
    if (action === "rowAbove") g.insertRows(sel.sheet, sel.startRow, rows, log);
    else if (action === "rowBelow") g.insertRows(sel.sheet, sel.endRow + 1, rows, log);
    else if (action === "colLeft") g.insertColumns(sel.sheet, sel.startCol, cols, log);
    else if (action === "colRight") g.insertColumns(sel.sheet, sel.endCol + 1, cols, log);
    else if (action === "deleteRows") g.deleteRows(sel.sheet, sel.startRow, rows, log);
    else if (action === "deleteColumns") g.deleteColumns(sel.sheet, sel.startCol, cols, log);
  };
  const none = () => {};
  return {
    onUndo: () => {
      grid()?.undo?.();
      changed();
    },
    onRedo: () => {
      grid()?.redo?.();
      changed();
    },
    onFormatPainter: none,
    onConditionalFormatting: none,
    onDataValidation: none,
    onFindReplace: none,
    onOutlineLevel: none,
    onMerge: none,
    onBorders: none,
    onClear: (kind) => {
      const s = selection();
      if (!s) return;
      const { g, sel } = s;
      if (kind === "formats" || kind === "all") {
        g.setRangeFormat(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol, {
          bold: false, italic: false, underline: false, strike: false, font_color: "#000000", background_color: null,
          number_format: "General", wrap_text: false, horizontal_align: "left", vertical_align: "bottom",
        });
      }
      if (kind === "contents" || kind === "all") g.clearRange(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol);
    },
    onFreeze: (kind) => {
      const g = grid();
      const sel = g?.getActiveSelection?.() ?? null;
      if (!g) return;
      const sheet = sel?.sheet ?? g.getActiveSheetName?.() ?? g.getSheetNames()[0];
      if (kind === "none") return g.unfreezePanes(sheet);
      if (kind === "atSelection" && !sel) return notify("Select a cell first");
      const rows = kind === "topRow" ? 1 : kind === "atSelection" ? sel!.startRow : 0;
      const cols = kind === "firstCol" ? 1 : kind === "atSelection" ? sel!.startCol : 0;
      g.freezePanes(sheet, rows, cols);
    },
    onToggleFilter: () => {
      const s = selection("Select the data range to filter");
      if (s) s.g.toggleFilter(s.sel.sheet, s.sel.startRow, s.sel.startCol, s.sel.endRow, s.sel.endCol);
    },
    onToggle: (key) => {
      const cur = anchorFormat();
      if (cur) format({ [key]: !cur[key] });
    },
    onFontColor: (color) => format({ font_color: color }),
    onFillColor: (color) => format({ background_color: color }),
    onAlign: (a) => format({ horizontal_align: a }),
    onVerticalAlign: (a) => format({ vertical_align: a }),
    onWrap: () => {
      const cur = anchorFormat();
      if (cur) format({ wrap_text: !cur.wrap_text });
    },
    onFontFamily: (family) => format({ font_family: family }),
    onFontSize: (size) => format({ font_size: size }),
    onAdjustFontSize: (delta) => {
      const cur = anchorFormat();
      if (cur) format({ font_size: stepFontSize(cur.font_size ?? 11, delta) });
    },
    onSort: (direction) => {
      const s = selection();
      if (!s) return;
      if (s.sel.startRow === s.sel.endRow) return notify("Select several rows to sort");
      s.g.sortRange(s.sel.sheet, s.sel.startRow, s.sel.startCol, s.sel.endRow, s.sel.endCol, s.sel.startCol, direction);
    },
    onInsert: insert,
    onNumberFormat: (pattern) => format({ number_format: pattern }),
    onAdjustDecimals: (delta) => {
      const cur = anchorFormat();
      if (cur) format({ number_format: adjustDecimalPlaces(cur.number_format, delta) });
    },
  };
}
