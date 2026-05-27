import React, { useEffect, useRef } from "react";
import ExcelJS from "exceljs";

// Univer ships a Tailwind-prefixed CSS bundle that must be loaded for the
// grid to render. Without these classes the canvas paints but the toolbar,
// formula bar, scrollbars, dropdowns and selection chrome are all unstyled
// and the UI looks broken. CSS imports are side-effect-only — Vite bundles
// them automatically.
import "@univerjs/preset-sheets-core/lib/index.css";

// Univer 0.23+ uses the preset-based facade API. We lazy-import so the rest
// of the screen renders even before `npm install` resolves the packages, and
// so we never pull Univer's bundle on routes that don't need it.
//
// v1 uses SheetJS to translate xlsx <-> Univer's IWorkbookData. Once we move
// to Univer's commercial xlsx exchange we just swap the two translators.

export interface CellFormatShape {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  font_color?: string;
  /** CSS color for cell fill (e.g. "#1F4E79"). Null/empty clears it. */
  background_color?: string | null;
  font_size?: number;
  font_family?: string;
  horizontal_align?: "left" | "center" | "right";
  vertical_align?: "top" | "middle" | "bottom";
  number_format?: string;
  wrap_text?: boolean;
  indent?: number;
}

export type SaveMirror = {
  /** Format changes the agent applied — pushed onto ExcelJS cells. */
  cellFormats?: Array<{ sheet: string; row: number; col: number; format: CellFormatShape; background?: string | null }>;
  /** Column width changes — pushed onto ExcelJS columns. */
  columnWidths?: Array<{ sheet: string; col: number; widthPx: number }>;
  /** Row height changes — pushed onto ExcelJS rows. */
  rowHeights?: Array<{ sheet: string; row: number; heightPx: number }>;
  /** Merge / unmerge changes — pushed onto the worksheet. */
  merges?: Array<{ sheet: string; range: string; merge: boolean }>;
  /** Sheet-level ops — applied in order. */
  sheetOps?: Array<
    | { kind: "create"; name: string; tabColor?: string | null }
    | { kind: "delete"; name: string }
    | { kind: "rename"; oldName: string; newName: string }
  >;
  /** Cells the agent cleared. ExcelJS cell.value set to null but format preserved. */
  clears?: Array<{ sheet: string; row: number; col: number }>;
  /** Row/column structure ops applied in order; cells will shift accordingly. */
  rowColOps?: Array<
    | { kind: "insertRows"; sheet: string; before: number; count: number }
    | { kind: "deleteRows"; sheet: string; start: number; count: number }
    | { kind: "insertColumns"; sheet: string; before: number; count: number }
    | { kind: "deleteColumns"; sheet: string; start: number; count: number }
  >;
  /** Per-sheet freeze pane state. */
  freezePanes?: Array<{ sheet: string; freezeRows: number; freezeCols: number }>;
  /** Hide / show row/col ops. */
  visibility?: Array<
    | { kind: "hideRows"; sheet: string; rows: number[] }
    | { kind: "showRows"; sheet: string; rows: number[] }
    | { kind: "hideColumns"; sheet: string; columns: number[] }
    | { kind: "showColumns"; sheet: string; columns: number[] }
  >;
};

export interface UniverGridHandle {
  loadBytes: (bytes: Uint8Array) => Promise<void>;
  /**
   * Write the workbook back to xlsx bytes. Optional `mirror` carries the
   * agent's format/width/height/merge mutations so the saved file actually
   * contains them — without it, only cell values/formulas are written.
   */
  exportBytes: (mirror?: SaveMirror) => Promise<Uint8Array>;
  setCell: (sheet: string, row: number, col: number, value: string | number | null) => void;
  /** Returns the current cell { value, formula } or null when the cell is empty. */
  getCell: (sheet: string, row: number, col: number) => { value: any; formula: string | null } | null;
  /** Paint or clear a cell background (pass `null` to clear). Used for the diff overlay. */
  setCellBackground: (sheet: string, row: number, col: number, color: string | null) => void;
  /** Apply a partial format object to a cell. Properties set to undefined are left alone. */
  setCellFormat: (sheet: string, row: number, col: number, format: CellFormatShape) => void;
  /** Read the current cell format. Returns only the properties we manage. */
  getCellFormat: (sheet: string, row: number, col: number) => CellFormatShape;
  /** Read-only snapshot of the workbook model. Used to ship context to the agent. */
  getWorkbookSnapshot: () => any | null;
  /**
   * Restore the workbook from a Univer-native JSON snapshot (the shape
   * returned by `getWorkbookSnapshot`). Lossless round-trip — preserves
   * all formatting that the xlsx exporter would otherwise drop. Used by
   * the auto-snapshot path for untitled drafts so reopening doesn't lose
   * fills, fonts, number formats, etc.
   */
  loadSnapshot: (snapshot: any) => Promise<void>;
  /**
   * Returns the user's current cell selection (sheet + rectangular bounds),
   * or null when there is no live selection. Used to ship selection-as-context
   * to the agent. Selection is read on demand at prompt-submit time.
   */
  getActiveSelection: () => {
    sheet: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  } | null;
  /** Returns the current column width in pixels (or null if unknown). */
  getColumnWidth: (sheet: string, col: number) => number | null;
  /** Set a column width in pixels. */
  setColumnWidth: (sheet: string, col: number, width: number) => void;
  /** Returns the current row height in pixels (or null if unknown). */
  getRowHeight: (sheet: string, row: number) => number | null;
  /** Set a row height in pixels. */
  setRowHeight: (sheet: string, row: number, height: number) => void;
  /** Merge a rectangular range (start/end inclusive, 0-indexed). */
  mergeCells: (sheet: string, startRow: number, startCol: number, endRow: number, endCol: number) => void;
  /** Unmerge a rectangular range. */
  unmergeCells: (sheet: string, startRow: number, startCol: number, endRow: number, endCol: number) => void;

  /** Create a new (empty) sheet. Returns true on success. */
  createSheet: (name: string, tabColor?: string | null) => boolean;
  /** Delete a sheet by name. */
  deleteSheet: (name: string) => boolean;
  /** Rename a sheet. */
  renameSheet: (oldName: string, newName: string) => boolean;

  /** Clear values/formulas of every cell in the rectangular range. Formatting preserved. */
  clearRange: (sheet: string, startRow: number, startCol: number, endRow: number, endCol: number) => void;

  /** Insert blank rows. `before` is 0-indexed. */
  insertRows: (sheet: string, before: number, count: number) => void;
  /** Delete `count` rows starting at `start` (0-indexed). */
  deleteRows: (sheet: string, start: number, count: number) => void;
  /** Insert blank columns. `before` is 0-indexed. */
  insertColumns: (sheet: string, before: number, count: number) => void;
  /** Delete `count` columns starting at `start` (0-indexed). */
  deleteColumns: (sheet: string, start: number, count: number) => void;

  /** Freeze top `freezeRows` and left `freezeCols`. Pass 0 to disable an axis. */
  freezePanes: (sheet: string, freezeRows: number, freezeCols: number) => void;
  /** Unfreeze all panes on a sheet. */
  unfreezePanes: (sheet: string) => void;

  /** Hide/show row(s) (0-indexed). */
  hideRows: (sheet: string, rows: number[]) => void;
  showRows: (sheet: string, rows: number[]) => void;
  /** Hide/show column(s) (0-indexed). */
  hideColumns: (sheet: string, columns: number[]) => void;
  showColumns: (sheet: string, columns: number[]) => void;
}

export const UniverGrid = React.forwardRef<UniverGridHandle, { workbookPath: string | null }>(
  function UniverGrid({ workbookPath }, ref) {
    const containerRef = useRef<HTMLDivElement>(null);
    const univerAPIRef = useRef<any>(null);
    const univerRef = useRef<any>(null);
    /**
     * Latest user selection captured from Univer's command stream. Univer 0.23's
     * facade `getActiveRange()` only returns the active cell (1×1), so polling
     * it can never see a multi-cell drag. We instead intercept the internal
     * `sheet.operation.set-selections` command — that fires every time the
     * user drags, shift-clicks, or arrow-extends — and cache the rectangle here.
     * `getActiveSelection` prefers this over the facade probe.
     */
    const liveSelectionRef = useRef<{
      sheet: string;
      startRow: number;
      startCol: number;
      endRow: number;
      endCol: number;
    } | null>(null);
    // The ExcelJS workbook we loaded from disk lives here. On save we patch
    // ONLY the cells the agent changed and write this object back out —
    // charts, conditional formatting, named ranges, data validation,
    // comments, frozen panes, drawings, themes, etc. all flow through
    // untouched because we never tried to recreate them. SheetJS used to
    // throw all of that away on the round-trip.
    const exceljsWorkbookRef = useRef<ExcelJS.Workbook | null>(null);
    // Resolved when the Univer instance finishes its async bootstrap so we
    // can `await` it from `loadBytes` / `setCell` without a race.
    const readyResolve = useRef<((api: any) => void) | null>(null);
    const readyPromise = useRef<Promise<any>>(
      new Promise((res) => {
        readyResolve.current = res;
      }),
    );

    const waitForApi = async (): Promise<any> => {
      if (univerAPIRef.current) return univerAPIRef.current;
      return readyPromise.current;
    };

    useEffect(() => {
      let cancelled = false;

      (async () => {
        if (!containerRef.current) return;

        const [{ createUniver, LocaleType, defaultTheme, merge }, { UniverSheetsCorePreset }] =
          await Promise.all([
            import("@univerjs/presets"),
            import("@univerjs/preset-sheets-core"),
          ]);

        // Locales — preset-sheets-core ships its own en-US bundle.
        const enUS = (
          await import("@univerjs/preset-sheets-core/locales/en-US")
        ).default;

        if (cancelled || !containerRef.current) return;

        const { univer, univerAPI } = createUniver({
          locale: LocaleType.EN_US,
          locales: { [LocaleType.EN_US]: merge({}, enUS) },
          theme: defaultTheme,
          presets: [
            UniverSheetsCorePreset({
              container: containerRef.current,
            }),
          ],
        });

        univerRef.current = univer;
        univerAPIRef.current = univerAPI;

        univerAPI.createWorkbook(blankWorkbook());

        // Subscribe to Univer's command stream and watch for selection ops.
        // The facade has no stable `onSelectionChange` in 0.23, but every
        // user drag/click/keyboard-extend dispatches a command whose id
        // contains "set-selections". We pull the rectangle out of the
        // params and stash it on liveSelectionRef. We log the first hit
        // so it's easy to confirm in DevTools.
        try {
          if (typeof univerAPI.onCommandExecuted === "function") {
            univerAPI.onCommandExecuted((cmd: any) => {
              if (!cmd || typeof cmd.id !== "string") return;
              if (!cmd.id.includes("set-selections") && !cmd.id.endsWith("SetSelectionsOperation")) return;
              const selections = cmd.params?.selections;
              if (!Array.isArray(selections) || selections.length === 0) return;
              // Primary selection is the *largest* by area — Univer puts
              // the active cell as a 1×1 selection in the same list when
              // a multi-cell drag exists, so picking the biggest is the
              // safest heuristic.
              let bestRange: any = null;
              let bestArea = 0;
              for (const s of selections) {
                const r = s?.range;
                if (!r || typeof r.startRow !== "number") continue;
                const sr = r.startRow;
                const sc = r.startColumn ?? r.startCol;
                const er = r.endRow ?? sr;
                const ec = r.endColumn ?? r.endCol ?? sc;
                const area = (er - sr + 1) * (ec - sc + 1);
                if (area > bestArea) {
                  bestArea = area;
                  bestRange = { startRow: sr, startCol: sc, endRow: er, endCol: ec };
                }
              }
              if (!bestRange) return;
              const sheetName = univerAPI.getActiveWorkbook?.()?.getActiveSheet?.()?.getName?.() ?? "";

              // Safety net for the boundary-wrap problem. If the previous
              // selection was a single cell at column 0 and the new one
              // jumped to a far-right column on the same row, that's the
              // ArrowLeft wrap we want to undo. Same idea for row-0 → far-bottom.
              const prev = liveSelectionRef.current;
              if (prev && prev.sheet === sheetName) {
                const wasSingle = prev.startRow === prev.endRow && prev.startCol === prev.endCol;
                const newSingle = bestRange.startRow === bestRange.endRow && bestRange.startCol === bestRange.endCol;
                if (wasSingle && newSingle) {
                  const wrapLeft =
                    prev.startCol === 0 &&
                    bestRange.startCol > 10 &&
                    bestRange.startRow === prev.startRow;
                  const wrapUp =
                    prev.startRow === 0 &&
                    bestRange.startRow > 10 &&
                    bestRange.startCol === prev.startCol;
                  if (wrapLeft || wrapUp) {
                    try {
                      const wb = univerAPI.getActiveWorkbook?.();
                      const sheet = wb?.getSheetByName?.(sheetName) ?? wb?.getActiveSheet?.();
                      sheet?.getRange?.(prev.startRow, prev.startCol)?.activate?.();
                    } catch (e) {
                      console.warn("[univer-sel] wrap snap-back failed:", e);
                    }
                    return;
                  }
                }
              }
              liveSelectionRef.current = { sheet: sheetName, ...bestRange };
              if (!(window as any).__univer_cmd_logged__) {
                (window as any).__univer_cmd_logged__ = true;
                console.log(`[univer-sel] command intercepted cmd.id=${cmd.id} → ${sheetName}!${bestRange.startRow},${bestRange.startCol}→${bestRange.endRow},${bestRange.endCol} area=${bestArea}`);
              }
            });
          } else {
            console.warn("[univer-sel] univerAPI.onCommandExecuted not available — falling back to facade polling");
          }
        } catch (e) {
          console.warn("[univer-sel] command subscription failed:", e);
        }

        console.log("[univer] grid ready");
        readyResolve.current?.(univerAPI);
      })().catch((e) => {
        console.error("[univer] init failed:", e);
      });

      return () => {
        cancelled = true;
        // Defer Univer's dispose() out of React's commit phase. Disposing
        // synchronously triggers React's "Attempted to synchronously unmount
        // a root while React was already rendering" warning because Univer
        // unmounts its internal React roots, which collides with our parent
        // tree still finishing its own unmount.
        const u = univerRef.current;
        univerRef.current = null;
        univerAPIRef.current = null;
        exceljsWorkbookRef.current = null;
        setTimeout(() => {
          try { u?.dispose?.(); } catch (e) { console.warn("[univer] dispose error:", e); }
        }, 0);
      };
    }, []);

    // Boundary-clamp arrow keys. Univer's default keyboard nav wraps when
    // the user presses Left at column A or Up at row 1 — the selection
    // jumps to the far column / far row, which is jarring. Window-capture
    // listener so we beat Univer's own handler. Filter by container
    // membership rather than tag name, because Univer captures keystrokes
    // via a hidden <input> inside the canvas — filtering by tag would
    // skip exactly the events we want to clamp.
    useEffect(() => {
      const onKeyDown = (e: KeyboardEvent) => {
        const target = e.target as Node | null;
        const container = containerRef.current;
        if (!container || !target || !container.contains(target)) return;
        if (e.metaKey || e.ctrlKey || e.altKey) return;
        const sel = liveSelectionRef.current;
        if (!sel) return;
        if (e.key === "ArrowLeft" && sel.startCol === 0 && sel.endCol === 0) {
          e.preventDefault();
          e.stopPropagation();
          (e as any).stopImmediatePropagation?.();
        } else if (e.key === "ArrowUp" && sel.startRow === 0 && sel.endRow === 0) {
          e.preventDefault();
          e.stopPropagation();
          (e as any).stopImmediatePropagation?.();
        }
      };
      window.addEventListener("keydown", onKeyDown, true);
      return () => window.removeEventListener("keydown", onKeyDown, true);
    }, []);

    React.useImperativeHandle(ref, () => ({
      loadBytes: async (bytes: Uint8Array) => {
        const api = await waitForApi();
        if (!api) {
          console.warn("[univer] loadBytes: API never became ready");
          return;
        }
        const { excelJs, univerData, styleOps } = await xlsxBytesToWorkbook(bytes);
        exceljsWorkbookRef.current = excelJs;
        const active = api.getActiveWorkbook?.();
        if (active?.dispose) {
          try { active.dispose(); } catch {}
        }
        api.createWorkbook(univerData);

        // Push every cell-level style we extracted from ExcelJS into Univer
        // via the facade so the user actually SEES the original colors,
        // fonts, number formats, alignment etc. — not just preserves them
        // on save. Run after createWorkbook so the sheets exist.
        for (const op of styleOps) {
          try {
            const wb = api.getActiveWorkbook?.();
            const sheet = wb?.getSheetByName?.(op.sheet) ?? wb?.getActiveSheet?.();
            const range = sheet?.getRange?.(op.row, op.col);
            if (!range) continue;
            if (op.format.bold !== undefined) range.setFontWeight?.(op.format.bold ? "bold" : "normal");
            if (op.format.italic !== undefined) range.setFontStyle?.(op.format.italic ? "italic" : "normal");
            if (op.format.underline !== undefined) range.setFontLine?.(op.format.underline ? "underline" : "none");
            if (op.format.strike !== undefined) range.setFontLine?.(op.format.strike ? "line-through" : "none");
            if (op.format.font_color) range.setFontColor?.(op.format.font_color);
            if (op.format.font_size) range.setFontSize?.(op.format.font_size);
            if (op.format.font_family) range.setFontFamily?.(op.format.font_family);
            if (op.format.horizontal_align) range.setHorizontalAlignment?.(op.format.horizontal_align);
            if (op.format.vertical_align) range.setVerticalAlignment?.(op.format.vertical_align);
            if (op.format.number_format) range.setNumberFormat?.(op.format.number_format);
            if (op.background) range.setBackgroundColor?.(op.background);
            if (op.borders) applyBordersToRange(range, op.borders);
          } catch (e) {
            // Per-cell style failures shouldn't abort the whole load.
            console.warn("[univer] style apply failed at", op.sheet, op.row, op.col, e);
          }
        }

        console.log("[univer] loaded workbook via ExcelJS:", {
          sheets: Object.keys(univerData.sheets ?? {}).length,
          sheetNames: (univerData.sheetOrder ?? []).map((id: string) => univerData.sheets[id]?.name),
          stylesApplied: styleOps.length,
          preserved: "charts / cf / named ranges / validation / comments live in ExcelJS workbook",
        });
      },
      exportBytes: async (mirror?: SaveMirror) => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        const data = wb?.getSnapshot?.();
        return workbookToXlsxBytes(data, exceljsWorkbookRef.current, mirror);
      },
      setCell: (sheetName, row, col, value) => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return;
        // Univer's setValue accepts either a primitive or { v, f } — passing
        // a formula string starting with '=' as the primitive does NOT make
        // it a formula. We unwrap that here for convenience.
        if (typeof value === "string" && value.startsWith("=")) {
          range.setValue({ f: value });
        } else {
          range.setValue(value);
        }
      },
      getCell: (sheetName, row, col) => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return null;
        try {
          const v = range.getValue?.();
          const f = range.getFormula?.();
          if (v === undefined && !f) return null;
          return { value: v ?? null, formula: f ?? null };
        } catch {
          return null;
        }
      },
      setCellBackground: (sheetName, row, col, color) => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return;
        try {
          if (color === null) range.setBackgroundColor?.(null);
          else range.setBackgroundColor?.(color);
        } catch (e) {
          console.warn("setBackgroundColor failed:", e);
        }
      },
      setCellFormat: (sheetName, row, col, format) => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return;
        try {
          if (format.bold !== undefined) range.setFontWeight?.(format.bold ? "bold" : "normal");
          if (format.italic !== undefined) range.setFontStyle?.(format.italic ? "italic" : "normal");
          if (format.underline !== undefined) range.setFontLine?.(format.underline ? "underline" : "none");
          if (format.strike !== undefined) range.setFontLine?.(format.strike ? "line-through" : "none");
          if (format.font_color) range.setFontColor?.(format.font_color);
          if (format.font_size) range.setFontSize?.(format.font_size);
          if (format.font_family) range.setFontFamily?.(format.font_family);
          if (format.horizontal_align) range.setHorizontalAlignment?.(format.horizontal_align);
          if (format.vertical_align) range.setVerticalAlignment?.(format.vertical_align);
          if (format.number_format) range.setNumberFormat?.(format.number_format);
          // background_color is stored in CellFormatShape but applied via
          // setBackgroundColor since Univer's facade exposes background as
          // a distinct API. Null/empty clears the fill.
          if (format.background_color !== undefined) {
            range.setBackgroundColor?.(format.background_color || null);
          }
        } catch (e) {
          console.warn("setCellFormat failed:", e);
        }
      },
      getCellFormat: (sheetName, row, col): CellFormatShape => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        const f: CellFormatShape = {};
        if (!range) return f;
        try {
          const fw = range.getFontWeight?.();
          if (fw === "bold") f.bold = true;
          const fs = range.getFontStyle?.();
          if (fs === "italic") f.italic = true;
          const fc = range.getFontColor?.();
          if (fc) f.font_color = fc;
          const sz = range.getFontSize?.();
          if (sz) f.font_size = sz;
          const ff = range.getFontFamily?.();
          if (ff) f.font_family = ff;
          const ha = range.getHorizontalAlignment?.();
          if (ha) f.horizontal_align = ha;
          const va = range.getVerticalAlignment?.();
          if (va) f.vertical_align = va;
          const nf = range.getNumberFormat?.();
          if (nf) f.number_format = nf;
          const bg = range.getBackgroundColor?.();
          // Univer sometimes returns a transparent default ("#ffffff" or
          // an empty string). Only surface a real fill so Reject can tell
          // "user set white" from "no fill explicitly set".
          if (bg && bg !== "rgba(0,0,0,0)" && bg !== "transparent") f.background_color = bg;
        } catch {}
        return f;
      },
      getWorkbookSnapshot: () => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        return wb?.getSnapshot?.() ?? null;
      },
      loadSnapshot: async (snapshot: any) => {
        const api = await waitForApi();
        if (!api) {
          console.warn("[univer] loadSnapshot: API never became ready");
          return;
        }
        const active = api.getActiveWorkbook?.();
        if (active?.dispose) {
          try { active.dispose(); } catch {}
        }
        // Univer's dispose() doesn't free the unit id immediately, so
        // re-using the snapshot's original id (typically "blank" for
        // freshly-created workbooks) triggers
        //   "cannot create a unit with the same unit id: blank"
        // Stamp a fresh, time-keyed id before createWorkbook — Univer
        // doesn't care what the id is as long as it's unique per session.
        const fresh = { ...snapshot, id: `gp-${Date.now()}-${Math.random().toString(36).slice(2, 8)}` };
        // The snapshot shape is what getWorkbookSnapshot produced. No
        // intermediate xlsx serialization → zero format loss on the
        // round-trip.
        api.createWorkbook(fresh);
      },
      getActiveSelection: () => {
        // Prefer the live ref captured by the command-stream subscription —
        // that's the only path that reliably sees multi-cell drags in
        // Univer 0.23. Fall back to the facade probe (active cell only)
        // if the subscription never fired (older / different Univer build).
        if (liveSelectionRef.current) return liveSelectionRef.current;

        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        const sheet = wb?.getActiveSheet?.();
        if (!sheet) return null;

        const tryReadRange = (range: any): { startRow: number; startCol: number; endRow: number; endCol: number } | null => {
          if (!range) return null;
          // Shape A: FRange with getRow/getColumn/getNumRows/getNumColumns
          if (typeof range.getRow === "function" && typeof range.getColumn === "function") {
            const startRow = range.getRow();
            const startCol = range.getColumn();
            const numRows = typeof range.getNumRows === "function" ? range.getNumRows() : 1;
            const numCols = typeof range.getNumColumns === "function" ? range.getNumColumns() : 1;
            if (typeof startRow === "number" && typeof startCol === "number") {
              return { startRow, startCol, endRow: startRow + numRows - 1, endCol: startCol + numCols - 1 };
            }
          }
          // Shape B: FRange with getRange() returning an IRange { startRow, startColumn, endRow, endColumn }
          if (typeof range.getRange === "function") {
            const ir = range.getRange();
            if (ir && typeof ir.startRow === "number") {
              return {
                startRow: ir.startRow,
                startCol: ir.startColumn,
                endRow: ir.endRow,
                endCol: ir.endColumn,
              };
            }
          }
          // Shape C: raw IRange { startRow, startColumn, endRow, endColumn }
          if (typeof range.startRow === "number" && typeof range.startColumn === "number") {
            return {
              startRow: range.startRow,
              startCol: range.startColumn,
              endRow: range.endRow ?? range.startRow,
              endCol: range.endColumn ?? range.startColumn,
            };
          }
          // Shape D: {row, column, numRows, numColumns}
          if (typeof range.row === "number" && typeof range.column === "number") {
            return {
              startRow: range.row,
              startCol: range.column,
              endRow: range.row + (range.numRows ?? 1) - 1,
              endCol: range.column + (range.numColumns ?? 1) - 1,
            };
          }
          return null;
        };

        try {
          // Probe every known facade entry point. In Univer 0.23,
          // `sheet.getActiveRange()` often returns the *cursor cell* (1×1)
          // even when the user has dragged out a multi-cell selection —
          // the real range lives in `getSelections()` or
          // `getActiveSelection().getActiveRangeList()`. So instead of
          // returning the first non-null candidate (which gave us a
          // permanently stuck "A1" for any drag), we collect all valid
          // ranges and pick the largest by area. A genuine multi-cell
          // drag always wins over a 1×1 cursor.
          // Each entry is [label, value] so we can log which one wins.
          const labelled: Array<[string, any]> = [
            ["sheet.getActiveRange", sheet.getActiveRange?.()],
            ["sheet.getSelection.getActiveRange", sheet.getSelection?.()?.getActiveRange?.()],
            ["sheet.getSelection.getCurrentCell", sheet.getSelection?.()?.getCurrentCell?.()],
            ["wb.getActiveSelection.getActiveRange", wb?.getActiveSelection?.()?.getActiveRange?.()],
            ["wb.getActiveRange", wb?.getActiveRange?.()],
          ];
          const sheetSelections = Array.isArray(sheet.getSelections?.()) ? sheet.getSelections() : [];
          sheetSelections.forEach((s: any, i: number) => labelled.push([`sheet.getSelections[${i}]`, s]));
          const wbRanges = Array.isArray(wb?.getActiveSelection?.()?.getActiveRangeList?.()) ? wb.getActiveSelection().getActiveRangeList() : [];
          wbRanges.forEach((s: any, i: number) => labelled.push([`wb.getActiveSelection.getActiveRangeList[${i}]`, s]));

          if (!(window as any).__univer_sel_logged__) {
            (window as any).__univer_sel_logged__ = true;
            console.log("[univer] selection probe — sheet methods:",
              sheet ? Object.getOwnPropertyNames(Object.getPrototypeOf(sheet) || {}).slice(0, 40) : null,
              "wb methods:", wb ? Object.getOwnPropertyNames(Object.getPrototypeOf(wb) || {}).slice(0, 40) : null);
          }

          let best: { startRow: number; startCol: number; endRow: number; endCol: number } | null = null;
          let bestArea = 0;
          let bestLabel = "(none)";
          const parsed: Array<[string, any]> = [];
          for (const [label, c] of labelled) {
            const r = tryReadRange(c);
            if (!r) continue;
            const area = (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
            parsed.push([label, { ...r, area }]);
            if (area > bestArea) {
              best = r;
              bestArea = area;
              bestLabel = label;
            }
          }
          // Throttled per-tick log: only emit when the picked range changes,
          // so we don't flood the console while the user is just clicking.
          const sigParts = best ? [bestLabel, best.startRow, best.startCol, best.endRow, best.endCol] : ["null"];
          const sig = sigParts.join(":");
          if ((window as any).__univer_sel_last_sig__ !== sig) {
            (window as any).__univer_sel_last_sig__ = sig;
            const bestStr = best
              ? `${best.startRow},${best.startCol} → ${best.endRow},${best.endCol} (area=${bestArea})`
              : "null";
            const parsedStr = parsed
              .map(([l, r]) => `${l}=${r.startRow},${r.startCol}→${r.endRow},${r.endCol} area=${r.area}`)
              .join(" | ");
            console.log(`[univer-sel] picked=${bestLabel} ${bestStr}  ::  ${parsedStr}`);
          }
          if (best) {
            return {
              sheet: sheet.getName?.() ?? "",
              ...best,
            };
          }
        } catch (e) {
          console.warn("[univer] getActiveSelection failed:", e);
        }
        return null;
      },
      getColumnWidth: (sheetName, col) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const w = sheet?.getColumnWidth?.(col);
          return typeof w === "number" ? w : null;
        } catch { return null; }
      },
      setColumnWidth: (sheetName, col, width) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try { sheet?.setColumnWidth?.(col, width); }
        catch (e) { console.warn("setColumnWidth failed:", e); }
      },
      getRowHeight: (sheetName, row) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const h = sheet?.getRowHeight?.(row);
          return typeof h === "number" ? h : null;
        } catch { return null; }
      },
      setRowHeight: (sheetName, row, height) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try { sheet?.setRowHeight?.(row, height); }
        catch (e) { console.warn("setRowHeight failed:", e); }
      },
      mergeCells: (sheetName, startRow, startCol, endRow, endCol) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const range = sheet?.getRange?.(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
          range?.merge?.();
        } catch (e) { console.warn("merge failed:", e); }
      },
      unmergeCells: (sheetName, startRow, startCol, endRow, endCol) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const range = sheet?.getRange?.(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
          range?.unmerge?.() ?? range?.breakApart?.();
        } catch (e) { console.warn("unmerge failed:", e); }
      },

      // ----- sheet ops -----

      createSheet: (name, tabColor): boolean => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        if (!wb) return false;
        try {
          // Univer facade methods vary by minor version — try the common ones.
          const created =
            wb.create?.(name) ??
            wb.insertSheet?.(name) ??
            wb.addSheet?.(name) ??
            null;
          if (created && tabColor) {
            try { created.setTabColor?.(tabColor); } catch {}
          }
          return !!created;
        } catch (e) {
          console.warn("[univer] createSheet failed:", e);
          return false;
        }
      },
      deleteSheet: (name): boolean => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        if (!wb) return false;
        try {
          const sheet = wb.getSheetByName?.(name);
          if (!sheet) return false;
          // Try a few API variants.
          const removed =
            wb.deleteSheet?.(sheet) ??
            wb.removeSheet?.(name) ??
            sheet.delete?.() ??
            null;
          return removed !== false;
        } catch (e) {
          console.warn("[univer] deleteSheet failed:", e);
          return false;
        }
      },
      renameSheet: (oldName, newName): boolean => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        const sheet = wb?.getSheetByName?.(oldName);
        if (!sheet) return false;
        try {
          sheet.setName?.(newName);
          return true;
        } catch (e) {
          console.warn("[univer] renameSheet failed:", e);
          return false;
        }
      },

      // ----- cell content ops -----

      clearRange: (sheetName, startRow, startCol, endRow, endCol) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const range = sheet?.getRange?.(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
          // Prefer clearContent — keeps formatting. clear() wipes both.
          range?.clearContent?.() ?? range?.clear?.();
        } catch (e) { console.warn("clearRange failed:", e); }
      },

      // ----- row/column insert/delete -----

      insertRows: (sheetName, before, count) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          sheet?.insertRows?.(before, count) ??
            sheet?.insertRowBefore?.(before, count) ??
            sheet?.insertRowsBefore?.(before, count);
        } catch (e) { console.warn("insertRows failed:", e); }
      },
      deleteRows: (sheetName, start, count) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          sheet?.deleteRows?.(start, count);
        } catch (e) { console.warn("deleteRows failed:", e); }
      },
      insertColumns: (sheetName, before, count) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          sheet?.insertColumns?.(before, count) ??
            sheet?.insertColumnBefore?.(before, count) ??
            sheet?.insertColumnsBefore?.(before, count);
        } catch (e) { console.warn("insertColumns failed:", e); }
      },
      deleteColumns: (sheetName, start, count) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          sheet?.deleteColumns?.(start, count);
        } catch (e) { console.warn("deleteColumns failed:", e); }
      },

      // ----- freeze / hide -----

      freezePanes: (sheetName, freezeRows, freezeCols) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          // Univer 0.23 facade — try the known shapes in order.
          sheet?.setFreeze?.({
            xSplit: freezeCols,
            ySplit: freezeRows,
            startRow: freezeRows,
            startColumn: freezeCols,
          }) ??
          sheet?.setFrozenRows?.(freezeRows) ??
          sheet?.setFrozenColumns?.(freezeCols);
        } catch (e) { console.warn("freezePanes failed:", e); }
      },
      unfreezePanes: (sheetName) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          sheet?.cancelFreeze?.() ??
          sheet?.setFreeze?.({ xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 });
        } catch (e) { console.warn("unfreezePanes failed:", e); }
      },

      hideRows: (sheetName, rows) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          for (const r of rows) sheet?.hideRow?.(r) ?? sheet?.hideRows?.(r, 1);
        } catch (e) { console.warn("hideRows failed:", e); }
      },
      showRows: (sheetName, rows) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          for (const r of rows) sheet?.showRow?.(r) ?? sheet?.showRows?.(r, 1) ?? sheet?.unhideRow?.(r);
        } catch (e) { console.warn("showRows failed:", e); }
      },
      hideColumns: (sheetName, columns) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          for (const c of columns) sheet?.hideColumn?.(c) ?? sheet?.hideColumns?.(c, 1);
        } catch (e) { console.warn("hideColumns failed:", e); }
      },
      showColumns: (sheetName, columns) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          for (const c of columns) sheet?.showColumn?.(c) ?? sheet?.showColumns?.(c, 1) ?? sheet?.unhideColumn?.(c);
        } catch (e) { console.warn("showColumns failed:", e); }
      },
    }));

    return (
      <div
        ref={containerRef}
        style={{ width: "100%", height: "100%", minHeight: 400, background: "#fff" }}
        data-workbook-path={workbookPath ?? ""}
      />
    );
  },
);

function resolveRange(api: any, sheetName: string, row: number, col: number): any {
  const wb = api?.getActiveWorkbook?.();
  if (!wb) return null;
  const sheet = wb.getSheetByName?.(sheetName) ?? wb.getActiveSheet?.();
  return sheet?.getRange?.(row, col) ?? null;
}

function resolveSheet(api: any, sheetName: string): any {
  const wb = api?.getActiveWorkbook?.();
  if (!wb) return null;
  return wb.getSheetByName?.(sheetName) ?? wb.getActiveSheet?.();
}

function blankWorkbook(): any {
  return {
    id: "blank",
    sheets: {
      sheet1: {
        id: "sheet1",
        name: "Sheet1",
        rowCount: 10000,
        columnCount: 200,
        cellData: {},
      },
    },
    sheetOrder: ["sheet1"],
  };
}

type BorderSide = { style: string; color: string | null } | null;
type BordersShape = {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
};

type StyleOp = {
  sheet: string;
  row: number;
  col: number;
  format: CellFormatShape;
  background: string | null;
  borders: BordersShape | null;
};

/**
 * Parse xlsx bytes with ExcelJS, returning:
 *   - the ExcelJS Workbook (kept alive in memory for round-trip save)
 *   - the IWorkbookData Univer needs to render the grid
 *   - a list of cell-level style ops to push into Univer's facade after
 *     createWorkbook so the user actually SEES the workbook's original
 *     colors / fonts / number formats / alignment (not just preserves them
 *     on save).
 *
 * Everything ExcelJS knows about that we DON'T touch — charts, conditional
 * formatting, named ranges, data validation, comments, drawings, themes,
 * sheet protection, frozen panes — stays inside the ExcelJS object. SheetJS
 * used to strip all of that on the round trip.
 */
async function xlsxBytesToWorkbook(bytes: Uint8Array): Promise<{
  excelJs: ExcelJS.Workbook;
  univerData: any;
  styleOps: StyleOp[];
}> {
  const excelJs = new ExcelJS.Workbook();
  // ExcelJS wants an ArrayBuffer; copy out of the Uint8Array's backing buffer
  // in case it's a slice (which is the case for Tauri-delivered bytes).
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await excelJs.xlsx.load(ab as ArrayBuffer);

  const sheets: Record<string, any> = {};
  const sheetOrder: string[] = [];
  const styleOps: StyleOp[] = [];

  excelJs.worksheets.forEach((ws) => {
    const name = ws.name;
    const id = `sheet_${name.replace(/[^a-zA-Z0-9]/g, "_")}`;
    const cellData: Record<number, Record<number, any>> = {};
    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const r = rowNumber - 1; // ExcelJS is 1-indexed
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const c = colNumber - 1;
        if (!cellData[r]) cellData[r] = {};
        const cellDescriptor = cellFromExcelJs(cell);
        const { format, background, borders } = extractStyleFromExcelJs(cell);
        // Embed style INLINE in cellData. Univer reads this on createWorkbook
        // and renders accordingly — number_format, fonts, colors, borders,
        // alignment all land on the right cells without a second facade pass.
        // The post-load styleOps loop is kept as a defensive fallback only.
        const s = buildUniverStyle(format, background, borders);
        if (s) cellDescriptor.s = s;
        cellData[r][c] = cellDescriptor;
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;

        if (Object.keys(format).length > 0 || background || borders) {
          styleOps.push({ sheet: name, row: r, col: c, format, background, borders });
        }
      });
    });

    // Merged cell ranges live in worksheet.model.merges as A1 strings like
    // "A1:A5" or "B2:D2". Translate to Univer's IRange shape so Univer
    // renders the merge natively on createWorkbook.
    const mergeData: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }> = [];
    const merges: any = (ws as any).model?.merges ?? [];
    if (Array.isArray(merges)) {
      for (const m of merges) {
        const ir = parseA1RangeString(typeof m === "string" ? m : String(m?.range ?? ""));
        if (ir) mergeData.push(ir);
      }
    }

    // Column widths — Excel stores them in character units. Convert to px
    // (rough approximation: width * 7 + 5 for default font). Univer wants
    // pixels in columnData[N] = { w: <px> }.
    const columnData: Record<number, { w?: number; hd?: number }> = {};
    if ((ws as any).columns) {
      ((ws as any).columns as any[]).forEach((col: any, idx: number) => {
        if (col && typeof col.width === "number" && col.width > 0) {
          columnData[idx] = { w: Math.round(col.width * 7 + 5) };
        }
        if (col?.hidden) (columnData[idx] = columnData[idx] ?? {}).hd = 1;
      });
    }

    // Row heights — Excel stores in points. 1 point ≈ 1.333 pixels at 96 DPI.
    const rowData: Record<number, { h?: number; hd?: number }> = {};
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const idx = rowNumber - 1;
      if (typeof row.height === "number" && row.height > 0) {
        rowData[idx] = { h: Math.round(row.height * 1.333) };
      }
      if ((row as any).hidden) (rowData[idx] = rowData[idx] ?? {}).hd = 1;
    });

    // Sheet tab color (the colored stripe on the tab at the bottom of Excel).
    // ExcelJS exposes it via worksheet.properties.tabColor as { argb }.
    let tabColor: string | null = null;
    const tabColorArgb = (ws as any).properties?.tabColor?.argb;
    if (tabColorArgb) tabColor = argbToCss(tabColorArgb);

    // Freeze panes — ExcelJS stores them in `worksheet.views` as
    // [{ state: 'frozen', xSplit, ySplit }]. xSplit = frozen-column count,
    // ySplit = frozen-row count.
    let frozenRows = 0;
    let frozenCols = 0;
    const views = (ws as any).views;
    if (Array.isArray(views)) {
      const frozenView = views.find((v: any) => v?.state === "frozen" || v?.state === "frozenSplit");
      if (frozenView) {
        frozenRows = Math.max(0, Number(frozenView.ySplit ?? 0));
        frozenCols = Math.max(0, Number(frozenView.xSplit ?? 0));
      }
    }

    sheets[id] = {
      id,
      name,
      // Excel's max is 1,048,576 × 16,384. We use 10000 × 200 — big enough
      // that arrow-key navigation never hits an edge for normal use, small
      // enough that Univer doesn't spend memory pre-allocating massive grids.
      // Was 100 × 26 — that was why ← at column A wrapped to column Z.
      rowCount: Math.max(10000, maxRow + 100),
      columnCount: Math.max(200, maxCol + 10),
      cellData,
      mergeData,
      columnData,
      rowData,
      // Univer's IWorksheetData supports tabColor (string). When null Univer
      // uses its default neutral.
      ...(tabColor ? { tabColor } : {}),
      // Freeze panes — Univer's IFreeze shape: { xSplit, ySplit, startRow, startColumn }
      // xSplit = frozen-col count, ySplit = frozen-row count, startRow/Column
      // = where scrolling begins. Match ExcelJS's freeze on load.
      ...(frozenRows > 0 || frozenCols > 0
        ? {
            freeze: {
              xSplit: frozenCols,
              ySplit: frozenRows,
              startRow: frozenRows,
              startColumn: frozenCols,
            },
          }
        : {}),
    };
    sheetOrder.push(id);
  });

  if (sheetOrder.length === 0) {
    sheets["sheet1"] = { id: "sheet1", name: "Sheet1", rowCount: 10000, columnCount: 200, cellData: {} };
    sheetOrder.push("sheet1");
  }

  return {
    excelJs,
    univerData: { id: `wb_${Date.now()}`, sheets, sheetOrder },
    styleOps,
  };
}

function cellFromExcelJs(cell: ExcelJS.Cell): any {
  const out: any = {};
  // ExcelJS cell.value shapes:
  //   primitive (number | string | boolean) | null | Date
  //   { formula, result } — formulas
  //   { sharedFormula, result } — shared formula
  //   { richText: [{ text, font }] } — rich text
  //   { text, hyperlink } — hyperlinks
  //   { error } — error cells (#REF!, #VALUE! etc.)
  const v: any = cell.value;
  if (v && typeof v === "object" && "formula" in v) {
    out.f = `=${(v as any).formula}`;
    if ("result" in v && (v as any).result !== undefined && (v as any).result !== null) {
      const r = (v as any).result;
      out.v = typeof r === "object" && r !== null && "error" in r ? (r as any).error : r;
    }
  } else if (v && typeof v === "object" && "sharedFormula" in v) {
    out.f = `=${(v as any).sharedFormula}`;
    if ("result" in v && (v as any).result !== undefined && (v as any).result !== null) {
      out.v = (v as any).result;
    }
  } else if (v && typeof v === "object" && "richText" in v) {
    out.v = (v as any).richText.map((r: any) => r.text).join("");
  } else if (v && typeof v === "object" && "hyperlink" in v) {
    out.v = (v as any).text ?? String((v as any).hyperlink);
  } else if (v && typeof v === "object" && "error" in v) {
    out.v = (v as any).error;
  } else if (v instanceof Date) {
    out.v = v.toISOString();
  } else if (v === null || v === undefined) {
    // empty
  } else if (typeof v === "object") {
    out.v = JSON.stringify(v);
  } else {
    out.v = v;
  }
  return out;
}

/**
 * Apply per-side borders to a Univer range. Univer's facade exposes border
 * APIs that vary slightly between minor versions — we try the most common
 * shapes in order. The worst case is "border doesn't render in our app" —
 * the saved file still has it because ExcelJS preserves the cell border
 * object untouched on the round-trip.
 */
function applyBordersToRange(range: any, borders: BordersShape): void {
  try {
    // Shape A: setBorder(top, left, bottom, right, vertical, horizontal, style, color)
    // documented in some Univer facade builds. Booleans for which sides + a
    // single style + a single color.
    if (typeof range.setBorder === "function") {
      const t = !!borders.top;
      const b = !!borders.bottom;
      const l = !!borders.left;
      const r = !!borders.right;
      if (t || b || l || r) {
        // Pick any side's style/color as the representative — Univer may
        // not support per-side colors via setBorder, so we'll degrade.
        const sample = borders.top ?? borders.bottom ?? borders.left ?? borders.right;
        const style = sample?.style ?? "thin";
        const color = sample?.color ?? "#000000";
        try {
          range.setBorder(t, l, b, r, false, false, style, color);
          return;
        } catch {
          // fall through to other shapes
        }
      }
    }
    // Shape B: setBorders({ top: { style, color }, ... })
    if (typeof (range as any).setBorders === "function") {
      try {
        (range as any).setBorders({
          top: borders.top ? { style: borders.top.style, color: borders.top.color } : undefined,
          bottom: borders.bottom ? { style: borders.bottom.style, color: borders.bottom.color } : undefined,
          left: borders.left ? { style: borders.left.style, color: borders.left.color } : undefined,
          right: borders.right ? { style: borders.right.style, color: borders.right.color } : undefined,
        });
        return;
      } catch {
        // fall through
      }
    }
    // Shape C: setBorderByPosition('top', style, color) per side
    if (typeof (range as any).setBorderByPosition === "function") {
      for (const [pos, side] of Object.entries(borders) as Array<[string, BorderSide]>) {
        if (!side) continue;
        try {
          (range as any).setBorderByPosition(pos, side.style, side.color ?? "#000000");
        } catch {}
      }
    }
  } catch (e) {
    // Don't let one bad border block the whole load. The file still has its
    // borders preserved through ExcelJS regardless of in-app rendering.
    console.warn("[univer] applyBordersToRange failed:", e);
  }
}

/**
 * Convert our CellFormatShape + background + borders into Univer's inline
 * cell style object (the IStyleData shape Univer reads from cellData[r][c].s).
 * Returning a style here means it's applied at createWorkbook time, which is
 * the only way certain attributes (notably number_format) actually take
 * effect — the post-load facade pass via setNumberFormat was a no-op.
 *
 * Univer style codes (from preset-sheets-core IStyleData):
 *   bg = background fill { rgb }
 *   cl = font color { rgb }
 *   bl = bold (0/1)
 *   it = italic (0/1)
 *   ul = underline { s: 0|1 }
 *   st = strikethrough { s: 0|1 }
 *   fs = font size
 *   ff = font family
 *   ht = horizontal align (1=left, 2=center, 3=right)
 *   vt = vertical align (1=top, 2=middle, 3=bottom)
 *   n  = number format { pattern }
 *   bd = borders { t, b, l, r each { s: <style code>, cl: { rgb } } }
 */
function buildUniverStyle(
  format: CellFormatShape,
  background: string | null,
  borders: BordersShape | null,
): any | null {
  const s: any = {};
  if (background) s.bg = { rgb: background };
  if (format.font_color) s.cl = { rgb: format.font_color };
  if (format.bold) s.bl = 1;
  if (format.italic) s.it = 1;
  if (format.underline) s.ul = { s: 1 };
  if (format.strike) s.st = { s: 1 };
  if (format.font_size) s.fs = format.font_size;
  if (format.font_family) s.ff = format.font_family;
  if (format.horizontal_align) {
    s.ht = format.horizontal_align === "left" ? 1 : format.horizontal_align === "center" ? 2 : 3;
  }
  if (format.vertical_align) {
    s.vt = format.vertical_align === "top" ? 1 : format.vertical_align === "middle" ? 2 : 3;
  }
  if (format.number_format) s.n = { pattern: format.number_format };
  // Wrap text — Univer uses tb (text-break): 1=clip, 2=wrap, 3=overflow.
  if (format.wrap_text) s.tb = 2;
  // Indent — Univer's pd (padding) or td (text-indent); the property name
  // varies across versions, so we set both as a defensive measure.
  if (format.indent && format.indent > 0) {
    s.pd = { l: format.indent * 8 };
    (s as any).td = format.indent;
  }

  if (borders) {
    const bd: any = {};
    const sideStyleCode = (style: string): number => {
      // Univer border style codes (approx — actual values may vary by build):
      // 1=thin, 2=hair, 3=dotted, 4=dashed, 5=dashDot, 6=dashDotDot,
      // 7=double, 8=medium, 9=mediumDashed, 10=mediumDashDot,
      // 11=mediumDashDotDot, 12=slantDashDot, 13=thick. Default to thin.
      switch (style) {
        case "thick": return 13;
        case "double": return 7;
        case "medium": return 8;
        case "dashed": return 4;
        case "dotted": return 3;
        case "hair": return 2;
        default: return 1;
      }
    };
    const sideOf = (side: BorderSide | undefined) =>
      side
        ? { s: sideStyleCode(side.style), cl: { rgb: side.color ?? "#000000" } }
        : undefined;
    const t = sideOf(borders.top);
    const b = sideOf(borders.bottom);
    const l = sideOf(borders.left);
    const r = sideOf(borders.right);
    if (t) bd.t = t;
    if (b) bd.b = b;
    if (l) bd.l = l;
    if (r) bd.r = r;
    if (Object.keys(bd).length > 0) s.bd = bd;
  }

  return Object.keys(s).length > 0 ? s : null;
}

/**
 * Parse "A1:B2" (or just "A1") into Univer's IRange shape. Returns null for
 * malformed input so the caller can skip silently.
 */
function parseA1RangeString(s: string): {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
} | null {
  const parts = s.split(":");
  const decode = (ref: string): { row: number; col: number } | null => {
    const m = /^\s*([A-Za-z]+)(\d+)\s*$/.exec(ref);
    if (!m) return null;
    const letters = m[1].toUpperCase();
    let col = 0;
    for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
    return { row: parseInt(m[2], 10) - 1, col: col - 1 };
  };
  if (parts.length === 1) {
    const a = decode(parts[0]);
    if (!a) return null;
    return { startRow: a.row, endRow: a.row, startColumn: a.col, endColumn: a.col };
  }
  if (parts.length === 2) {
    const a = decode(parts[0]);
    const b = decode(parts[1]);
    if (!a || !b) return null;
    return {
      startRow: Math.min(a.row, b.row),
      endRow: Math.max(a.row, b.row),
      startColumn: Math.min(a.col, b.col),
      endColumn: Math.max(a.col, b.col),
    };
  }
  return null;
}

function argbToCss(argb: string | undefined | null): string | null {
  if (!argb) return null;
  // ExcelJS colors are 8-char ARGB ("FF00FF00"). Strip the alpha prefix and
  // prepend '#'. Some colors come as theme refs ({ theme: N, tint: ... }) —
  // those we skip; Univer's default render will use the cell's plain values.
  if (typeof argb !== "string" || argb.length < 6) return null;
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  return `#${hex.toUpperCase()}`;
}

function extractStyleFromExcelJs(cell: ExcelJS.Cell): {
  format: CellFormatShape;
  background: string | null;
  borders: BordersShape | null;
} {
  const format: CellFormatShape = {};
  let background: string | null = null;
  let borders: BordersShape | null = null;

  const border = cell.border as any;
  if (border) {
    const sideOf = (s: any): BorderSide => {
      if (!s || !s.style) return null;
      const color = s.color ? argbToCss(s.color.argb) : null;
      return { style: String(s.style), color };
    };
    const t = sideOf(border.top);
    const b = sideOf(border.bottom);
    const l = sideOf(border.left);
    const r = sideOf(border.right);
    if (t || b || l || r) {
      borders = {};
      if (t) borders.top = t;
      if (b) borders.bottom = b;
      if (l) borders.left = l;
      if (r) borders.right = r;
    }
  }

  const font = cell.font as any;
  if (font) {
    if (font.bold) format.bold = true;
    if (font.italic) format.italic = true;
    if (font.underline) format.underline = true;
    if (font.strike) format.strike = true;
    if (font.size && typeof font.size === "number") format.font_size = font.size;
    if (font.name && typeof font.name === "string") format.font_family = font.name;
    if (font.color) {
      const c = argbToCss(font.color.argb);
      if (c) format.font_color = c;
    }
  }

  const alignment = cell.alignment as any;
  if (alignment) {
    if (alignment.horizontal === "left" || alignment.horizontal === "center" || alignment.horizontal === "right") {
      format.horizontal_align = alignment.horizontal;
    }
    if (alignment.vertical === "top" || alignment.vertical === "middle" || alignment.vertical === "bottom") {
      format.vertical_align = alignment.vertical;
    }
    if (alignment.wrapText) format.wrap_text = true;
    if (typeof alignment.indent === "number" && alignment.indent > 0) {
      format.indent = alignment.indent;
    }
  }

  // numFmt can be a string like "$#,##0.00" or "0.00%" — pass through verbatim.
  // ExcelJS exposes the resolved string for built-in formats too.
  const nf = cell.numFmt;
  if (nf && typeof nf === "string" && nf !== "General") {
    format.number_format = nf;
  }

  const fill = cell.fill as any;
  if (fill && fill.type === "pattern" && fill.pattern === "solid") {
    const fgColor = fill.fgColor as any;
    if (fgColor) {
      const c = argbToCss(fgColor.argb);
      if (c) background = c;
    }
  }

  return { format, background, borders };
}

/**
 * Write the workbook back to xlsx bytes. For each cell with content in
 * Univer's current model, push the value/formula back into the ExcelJS
 * workbook — which still carries the original style and surrounding
 * structure — then writeBuffer. Cells we never touched keep their original
 * ExcelJS state (including font/fill/numFmt). Charts, conditional
 * formatting, named ranges, validation, etc. flow through untouched.
 *
 * For an untitled (brand-new) workbook with no ExcelJS source, build a
 * fresh ExcelJS workbook from Univer's snapshot.
 */
async function workbookToXlsxBytes(
  data: any,
  excelJsWorkbook: ExcelJS.Workbook | null,
  mirror?: SaveMirror,
): Promise<Uint8Array> {
  if (!excelJsWorkbook) {
    const fresh = new ExcelJS.Workbook();
    const order: string[] = data?.sheetOrder ?? Object.keys(data?.sheets ?? {});
    for (const sheetId of order) {
      const sheet = data.sheets[sheetId];
      if (!sheet) continue;
      const ws = fresh.addWorksheet(sheet.name);
      writeUniverSheetIntoExcelJs(sheet, ws);
    }
    if (fresh.worksheets.length === 0) fresh.addWorksheet("Sheet1");
    // For brand-new workbooks the agent's formatting also needs to be mirrored.
    // Style-only here — structure ops on a brand-new workbook are uncommon
    // and don't need the pre-cells split.
    if (mirror) applyStyleMirror(fresh, mirror);
    const buf = await fresh.xlsx.writeBuffer();
    return new Uint8Array(buf as ArrayBuffer);
  }

  // Apply STRUCTURE ops first (sheet create/delete/rename, row/col splices)
  // so that when we then write Univer's cells to ExcelJS, the worksheets
  // exist with the correct names and the existing-row layout is already
  // shifted. Otherwise ExcelJS still has old-position cells lingering when
  // we write to new positions, producing duplicate content.
  if (mirror) applyStructureMirror(excelJsWorkbook, mirror);

  const order: string[] = data?.sheetOrder ?? Object.keys(data?.sheets ?? {});
  for (const sheetId of order) {
    const sheet = data.sheets[sheetId];
    if (!sheet) continue;
    let ws = excelJsWorkbook.getWorksheet(sheet.name);
    if (!ws) {
      ws = excelJsWorkbook.addWorksheet(sheet.name);
    }
    writeUniverSheetIntoExcelJs(sheet, ws);
  }

  // Style ops (cell formats / widths / heights / merges) come AFTER cells
  // since they reference final positions and depend on the cell existing.
  if (mirror) applyStyleMirror(excelJsWorkbook, mirror);

  const buf = await excelJsWorkbook.xlsx.writeBuffer();
  return new Uint8Array(buf as ArrayBuffer);
}

/**
 * Structure ops — sheet create/delete/rename, row/column splices, explicit
 * clears. Applied BEFORE we write Univer's cell values into ExcelJS so the
 * sheets exist and the row/column layout is already shifted by the time
 * cell writes happen. Mixing the order produces duplicate / phantom content.
 */
function applyStructureMirror(wb: ExcelJS.Workbook, mirror: SaveMirror): void {
  // --- sheet ops ---
  for (const op of mirror.sheetOps ?? []) {
    try {
      if (op.kind === "create") {
        if (!wb.getWorksheet(op.name)) {
          const ws = wb.addWorksheet(op.name);
          if (op.tabColor) {
            const argb = cssToArgb(op.tabColor);
            (ws.properties as any).tabColor = { argb };
          }
        }
      } else if (op.kind === "delete") {
        const ws = wb.getWorksheet(op.name);
        if (ws) wb.removeWorksheet(ws.id);
      } else if (op.kind === "rename") {
        const ws = wb.getWorksheet(op.oldName);
        if (ws) ws.name = op.newName;
      }
    } catch (e) {
      console.warn("[save] sheet op failed:", op, e);
    }
  }

  // --- row/col insert/delete ---
  for (const op of mirror.rowColOps ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    try {
      if (op.kind === "insertRows") {
        // ExcelJS spliceRows(start, deleteCount, ...rows). To insert empty
        // rows at index `before` (0-indexed) we pass a count of zero
        // deletions and N empty arrays.
        const blanks = Array(op.count).fill([]);
        ws.spliceRows(op.before + 1, 0, ...blanks);
      } else if (op.kind === "deleteRows") {
        ws.spliceRows(op.start + 1, op.count);
      } else if (op.kind === "insertColumns") {
        const blanks = Array(op.count).fill([]);
        (ws as any).spliceColumns?.(op.before + 1, 0, ...blanks);
      } else if (op.kind === "deleteColumns") {
        (ws as any).spliceColumns?.(op.start + 1, op.count);
      }
    } catch (e) {
      console.warn("[save] row/col op failed:", op, e);
    }
  }

  // --- explicit clears (set value to null, keep formatting) ---
  for (const op of mirror.clears ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    try {
      ws.getCell(op.row + 1, op.col + 1).value = null;
    } catch {}
  }
}

/**
 * Style ops — cell formats, column widths, row heights, merges. Applied
 * AFTER cells have been written into ExcelJS so we never restyle a cell
 * that doesn't exist yet, and AFTER structure ops so the row/col indices
 * are stable.
 */
function applyStyleMirror(wb: ExcelJS.Workbook, mirror: SaveMirror): void {
  // --- cell formats ---
  for (const op of mirror.cellFormats ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    const cell = ws.getCell(op.row + 1, op.col + 1);
    const f = op.format;

    // Merge into existing font/fill/alignment rather than replacing — so
    // an agent setting `bold: true` doesn't wipe the cell's existing
    // font color or size from the original file.
    const prevFont = (cell.font as any) ?? {};
    const nextFont: any = { ...prevFont };
    if (f.bold !== undefined) nextFont.bold = f.bold;
    if (f.italic !== undefined) nextFont.italic = f.italic;
    if (f.underline !== undefined) nextFont.underline = f.underline;
    if (f.strike !== undefined) nextFont.strike = f.strike;
    if (f.font_size !== undefined) nextFont.size = f.font_size;
    if (f.font_family !== undefined) nextFont.name = f.font_family;
    if (f.font_color) nextFont.color = { argb: cssToArgb(f.font_color) };
    if (Object.keys(nextFont).length > 0) cell.font = nextFont;

    if (op.background) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: cssToArgb(op.background) },
      };
    }

    const prevAlign = (cell.alignment as any) ?? {};
    const nextAlign: any = { ...prevAlign };
    if (f.horizontal_align) nextAlign.horizontal = f.horizontal_align;
    if (f.vertical_align) nextAlign.vertical = f.vertical_align;
    if (f.wrap_text !== undefined) nextAlign.wrapText = f.wrap_text;
    if (f.indent !== undefined) nextAlign.indent = f.indent;
    if (Object.keys(nextAlign).length > 0) cell.alignment = nextAlign;

    if (f.number_format) cell.numFmt = f.number_format;
  }

  // --- column widths --- (px → Excel char-units, inverse of load-time conversion)
  for (const op of mirror.columnWidths ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    const col = ws.getColumn(op.col + 1);
    col.width = Math.max(1, (op.widthPx - 5) / 7);
  }

  // --- row heights --- (px → Excel points, inverse of load-time conversion)
  for (const op of mirror.rowHeights ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    const row = ws.getRow(op.row + 1);
    row.height = op.heightPx / 1.333;
  }

  // --- merges / unmerges ---
  for (const op of mirror.merges ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    try {
      if (op.merge) ws.mergeCells(op.range);
      else ws.unMergeCells(op.range);
    } catch (e) {
      console.warn("[save] merge/unmerge failed:", op, e);
    }
  }

  // --- freeze panes ---
  for (const op of mirror.freezePanes ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    try {
      if (op.freezeRows === 0 && op.freezeCols === 0) {
        (ws as any).views = [{ state: "normal" }];
      } else {
        (ws as any).views = [
          {
            state: "frozen",
            xSplit: op.freezeCols,
            ySplit: op.freezeRows,
            topLeftCell: undefined,
            activeCell: undefined,
          },
        ];
      }
    } catch (e) {
      console.warn("[save] freeze failed:", op, e);
    }
  }

  // --- hide / show ---
  for (const op of mirror.visibility ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    try {
      if (op.kind === "hideRows" || op.kind === "showRows") {
        for (const r of op.rows) {
          const row = ws.getRow(r + 1);
          row.hidden = op.kind === "hideRows";
        }
      } else {
        for (const c of op.columns) {
          const col = ws.getColumn(c + 1);
          col.hidden = op.kind === "hideColumns";
        }
      }
    } catch (e) {
      console.warn("[save] hide/show failed:", op, e);
    }
  }
}

function cssToArgb(css: string): string {
  // Strip leading '#'; assume opaque if no alpha provided.
  const hex = css.startsWith("#") ? css.slice(1) : css;
  if (hex.length === 6) return `FF${hex.toUpperCase()}`;
  if (hex.length === 8) return hex.toUpperCase();
  // Fall back to opaque black for unrecognized formats.
  return "FF000000";
}

function writeUniverSheetIntoExcelJs(univerSheet: any, ws: ExcelJS.Worksheet): void {
  const cellData = univerSheet.cellData ?? {};
  for (const rowKey of Object.keys(cellData)) {
    const r = Number(rowKey);
    const row = cellData[rowKey];
    for (const colKey of Object.keys(row)) {
      const c = Number(colKey);
      const cell = row[colKey];
      const eCell = ws.getCell(r + 1, c + 1);
      if (cell.f) {
        // Setting .value to { formula } preserves the cell's existing
        // style object (font/fill/numFmt) since we don't reassign it.
        eCell.value = { formula: String(cell.f).replace(/^=/, "") } as any;
      } else if (cell.v !== undefined && cell.v !== null) {
        eCell.value = cell.v;
      } else {
        eCell.value = null;
      }
    }
  }
}
