import React, { useEffect, useRef } from "react";
import ExcelJS from "exceljs";
import { applyExcelJsPatches } from "../exceljsPatches";
import { addXlfnPrefixes } from "../xlfnCompat";
import { buildCellBaseline, buildWorkbookPatch, ERROR_VALUES, type BaselineCell, type PatchBuildResult } from "../surgicalPatch";
import { computeCalcHealth, type CalcHealth } from "../agent/calcHealth";
import { scanFidelityRisks, type FidelityRisks } from "../fidelityScan";
import {
  detectIterativeCalc,
  externalPinKey,
  externalPinsByWorkbook,
  parseXlsxWorkbook,
  type ExternalPinMap,
  type SheetImage,
  type SheetOutline,
  type StyleOp,
} from "../xlsxImport";
import {
  ironcalcShadowOnCommand,
  registerUniverIdleWaiter,
  registerUniverValueGetter,
} from "../ironcalcShadow";
import {
  ironcalcEngineEnabled,
  ironcalcEngineOnCommand,
  ironcalcEngineSettled,
  registerIroncalcApplier,
  registerIroncalcSnapshotter,
  unregisterIroncalcApplier,
  unregisterIroncalcSnapshotter,
  type GridResyncSheet,
} from "../ironcalcEngine";
import {
  applyFeatureMirror,
  featureDriftDetail,
  perSheetFeatureState,
  type SheetFeatureState,
} from "../featureIo";

// Must run before the first xlsx.load — see exceljsPatches.ts for the bugs
// these work around (unparseable chart drawings, defined-name explosions).
applyExcelJsPatches();

// Univer ships a Tailwind-prefixed CSS bundle that must be loaded for the
// grid to render. Without these classes the canvas paints but the toolbar,
// formula bar, scrollbars, dropdowns and selection chrome are all unstyled
// and the UI looks broken. CSS imports are side-effect-only — Vite bundles
// them automatically.
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-find-replace/lib/index.css";
import "@univerjs/preset-sheets-filter/lib/index.css";
import "@univerjs/preset-sheets-drawing/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
import "@univerjs/preset-sheets-data-validation/lib/index.css";
import "@univerjs/preset-sheets-sort/lib/index.css";
import "@univerjs/preset-sheets-note/lib/index.css";

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
  /**
   * Per-cell border edits (currently from the manual format toolbar). Each
   * side: an object to set it, `null` to clear it, `undefined` to leave it.
   */
  cellBorders?: Array<{ sheet: string; row: number; col: number; borders: BordersShape }>;
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
  /** Workbook-scoped defined names (named ranges) to (re)create on save. */
  definedNames?: Array<{ name: string; ref: string }>;
  /**
   * Per-sheet AutoFilter range (A1 notation), or `null` to clear it. Written to
   * ExcelJS `worksheet.autoFilter`. Last entry per sheet wins.
   */
  autoFilters?: Array<{ sheet: string; range: string | null }>;
};

export interface UniverGridHandle {
  /** `onProgress` receives coarse stage labels ("Parsing workbook…") for a
   *  loading overlay — big models take seconds and used to look frozen. */
  loadBytes: (bytes: Uint8Array, onProgress?: (stage: string) => void) => Promise<void>;
  /**
   * Write the workbook back to xlsx bytes. Optional `mirror` carries the
   * agent's format/width/height/merge mutations so the saved file actually
   * contains them — without it, only cell values/formulas are written.
   */
  exportBytes: (mirror?: SaveMirror) => Promise<Uint8Array>;
  /**
   * Build a surgical-save patch: the cell diff (live Univer model vs the
   * load-time baseline) plus the mirror's format/structure ops (including
   * sheet create/rename/delete and row/col insert-delete), in the wire
   * shape `save_workbook_patched` expects. When the surgical path can't
   * represent the edits, returns a TYPED reason — callers decide between
   * the full export and blocking, and show reason-specific copy.
   */
  exportPatch: (mirror?: SaveMirror) => PatchBuildResult;
  /**
   * At-risk package content found at load time (charts, comments, external
   * links…) — what the ExcelJS full-export fallback would drop. null when
   * no file was loaded or the scan failed.
   */
  getFidelityRisks: () => FidelityRisks | null;
  /** The exact bytes this grid was loaded from (for post-export audits). */
  getSourceBytes: () => Uint8Array | null;
  /**
   * Reset the save baseline to "the file now holds exactly `bytes`, which
   * reflect the live model". Called after a successful FULL-EXPORT save:
   * unlike surgical saves (which replay the whole session against the
   * load-time bytes and are therefore idempotent), a full export bakes
   * every batch and structural op into the file — replaying them against
   * it on the next save would apply row inserts / sheet deletes twice.
   * The caller must stop replaying consumed batches/ops in tandem.
   *
   * The in-memory ExcelJS workbook is NOT reset: exportBytes already
   * mutated it (structure + style mirrors) to match `bytes`.
   */
  commitSavedBaseline: (bytes: Uint8Array) => Promise<void>;
  /**
   * Resolves true once Univer AND an active workbook exist (bounded wait).
   * Agent tool calls MUST await this before mutating: on a fresh tab the
   * first tool call can arrive while Univer is still loading its dynamic
   * imports, and every grid write is optional-chained — without this gate a
   * whole set_range silently no-ops and the agent builds on a phantom layout.
   */
  whenReady: () => Promise<boolean>;
  /** Sheet names currently in the workbook (empty when not ready). */
  getSheetNames: () => string[];
  /**
   * Max outline (grouping) depth of a sheet's columns/rows, from the file's
   * outlineLevel attributes. {cols: 0, rows: 0} when the sheet has no groups.
   */
  getOutlineSummary: (sheetName: string) => { cols: number; rows: number } | null;
  /**
   * Excel's "1 2 3" outline buttons: show grouped columns/rows with
   * outlineLevel < level, hide those >= level (level = maxDepth+1 shows
   * everything). Applies to the live grid and returns the affected indices
   * so the caller can record visibility ops for the save mirror.
   */
  applyOutlineLevel: (
    sheetName: string,
    axis: "cols" | "rows",
    level: number,
  ) => { hide: number[]; show: number[] } | null;
  setCell: (sheet: string, row: number, col: number, value: string | number | null) => void;
  /**
   * Write many cells on one sheet in a SINGLE Univer command (sparse object
   * matrix through set-range-values). One command pipeline pass, one undo
   * entry, one formula dirty-graph pass — instead of one full cycle per cell.
   * Cells not in the list are untouched (no rectangle overwrite).
   */
  setCells: (
    sheet: string,
    cells: Array<{ row: number; col: number; value: string | number | boolean | null }>,
  ) => boolean;
  /**
   * Resolves when the (worker-side) formula engine finishes recalculating,
   * or after `timeoutMs` if nothing is computing / it takes too long.
   */
  whenCalculated: (timeoutMs?: number) => Promise<void>;
  /**
   * Returns the current cell { value, display?, formula } or null when the
   * cell is empty. `value` is the RAW stored value (0.7235, never "72.35%");
   * `display` carries the number-formatted rendering only when it differs
   * from the raw value, so consumers can show both without re-deriving it.
   */
  getCell: (
    sheet: string,
    row: number,
    col: number,
  ) => { value: any; display?: string; formula: string | null } | null;
  /**
   * The cell as the FILE last saved it ({f: formula sans '=', v: cached
   * value}), from the load-time baseline. null when the workbook wasn't
   * loaded from disk or the cell was empty at load. Read-only provenance
   * for live-vs-file divergence flagging (rule 17b).
   */
  getFileSavedCell: (sheet: string, row: number, col: number) => { f?: string; v?: any } | null;
  /**
   * One-shot live-vs-file-saved divergence scan, computed after the formula
   * engine settles and cached for the lifetime of the loaded workbook (the
   * result feeds the cached context block, so its bytes must never jitter).
   * null when there's no disk baseline (untitled tabs).
   */
  getCalcHealth: () => Promise<CalcHealth | null>;
  /**
   * Bulk-read an inclusive rectangle in a few facade calls instead of one
   * resolveRange per cell. Row-major 2D arrays, [r - sr][c - sc] addressing;
   * formulas are null when absent. `values` are RAW stored values (0.7235,
   * never "72.35%" — display strings here corrupted old_value captures);
   * `displays` carries the number-formatted renderings when available.
   * Returns null when the sheet/range can't be resolved — callers fall back
   * to per-cell getCell.
   */
  getRangeData: (
    sheet: string,
    sr: number,
    sc: number,
    er: number,
    ec: number,
  ) => { values: any[][]; formulas: (string | null)[][]; displays?: (string | null)[][] } | null;
  /** Paint or clear a cell background (pass `null` to clear). Used for the diff overlay. */
  setCellBackground: (sheet: string, row: number, col: number, color: string | null) => void;
  /** Apply a partial format object to a cell. Properties set to undefined are left alone. */
  setCellFormat: (sheet: string, row: number, col: number, format: CellFormatShape) => void;
  /**
   * Apply a partial format to an entire rectangular range in ONE Univer command,
   * so a single ⌘Z reverts the whole multi-cell change (per-cell setCellFormat
   * would create one undo step per cell). Used by the manual formatting toolbar.
   */
  setRangeFormat: (
    sheet: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
    format: CellFormatShape,
  ) => void;
  /** Read the current cell format. Returns only the properties we manage. */
  getCellFormat: (sheet: string, row: number, col: number) => CellFormatShape;
  /** Undo the last Univer command on the active workbook (programmatic). */
  undo: () => void;
  /** Redo the last undone Univer command on the active workbook. */
  redo: () => void;
  /** True when DOM focus is currently inside the grid container. */
  containsFocus: () => boolean;
  /** Open Univer's Find (or Find & Replace) dialog. */
  openFindReplace: (replace?: boolean) => void;
  /** Open Univer's conditional-formatting side panel (rules for the active sheet). */
  openConditionalFormatting: () => void;
  /** Open Univer's data-validation side panel. */
  openDataValidation: () => void;
  /**
   * Toggle an AutoFilter over the given rectangle. If the sheet already has a
   * filter it is removed; otherwise one is created over the range. Returns the
   * resulting state (true = filter now active).
   */
  toggleFilter: (sheet: string, sr: number, sc: number, er: number, ec: number) => boolean;
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
   * Freeze (false) or resume (true) this grid's canvas render loop.
   * Background (hidden) tabs keep their full command pipeline, IronCalc
   * mirror and data model — only the per-frame scene painting stops, so
   * agent writes to a non-visible tab skip all layout/paint work. Dirty
   * flags accumulate while frozen and the first frame after resuming
   * repaints everything that changed. Safe to call before the grid is
   * ready: the desired state is re-applied whenever a workbook is created.
   */
  setRenderActive: (active: boolean) => void;
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
  /**
   * Name of the sheet the user is currently viewing, or null when Univer
   * isn't ready. Shipped to the agent with every prompt so it knows where
   * the user is even without a selection.
   */
  getActiveSheetName: () => string | null;
  /**
   * Bring a cell into view: activate its sheet, scroll it near the top-left
   * of the viewport, and select it. Used by the changes-review modal's
   * jump-to-cell. Returns false if the sheet doesn't exist.
   */
  jumpToCell: (sheet: string, row: number, col: number) => boolean;
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

  /** Set (or overwrite) a cell's note/comment. */
  setNote: (sheet: string, row: number, col: number, text: string) => void;
  /** Remove a cell's note. No-op if it has none. */
  deleteNote: (sheet: string, row: number, col: number) => void;
  /** Current note text on a cell, or null if it has none. */
  getNote: (sheet: string, row: number, col: number) => string | null;

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

  /**
   * Apply per-side borders to a single cell. Sides set to `null` are cleared;
   * sides left `undefined` are untouched. In-app rendering is best-effort
   * (Univer's facade border API varies by build) — the saved xlsx is the
   * source of truth via the SaveMirror.borders path.
   */
  setCellBorders: (sheet: string, row: number, col: number, borders: BordersShape) => void;
  /**
   * Apply borders across a whole rectangle in a single Univer command so one
   * ⌘Z undoes the entire border action (Excel-style). `kind` mirrors the
   * toolbar menu: "all" boxes every cell, "outer" draws only the perimeter,
   * "none" clears.
   */
  setRangeBorders: (
    sheet: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
    kind: "all" | "outer" | "none",
  ) => void;
  /**
   * Sort the rows of a rectangle by `keyCol` (an absolute column index inside
   * the rectangle). Carries each row's full cell data (value + style) so
   * formatting travels with the row, Excel-style. Single command → single undo.
   */
  sortRange: (
    sheet: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
    keyCol: number,
    direction: "asc" | "desc",
  ) => void;

  /**
   * Create or replace a workbook-scoped defined name (named range). `ref` is
   * a sheet-qualified absolute A1 string, e.g. "Model!$B$5:$B$12".
   */
  defineName: (name: string, ref: string) => void;
  /** Delete a defined name. No-op if it doesn't exist. */
  deleteName: (name: string) => void;
  /** Current ref string of a defined name, or null if it doesn't exist. */
  getDefinedNameRef: (name: string) => string | null;

  /**
   * Capture every non-empty cell in a rectangular band (inclusive, 0-indexed)
   * with value + formula + format. Used before delete_rows / delete_columns so
   * Reject can restore the deleted content after re-inserting the band.
   */
  captureCellBand: (
    sheet: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
  ) => Array<{
    row: number;
    col: number;
    value: any;
    formula: string | null;
    format: CellFormatShape | null;
  }>;

  /** Current freeze pane counts for a sheet, or {0,0} if none / unreadable. */
  getFreezePanes: (sheet: string) => { freezeRows: number; freezeCols: number };

  /**
   * Extract one sheet's Univer snapshot (from getWorkbookSnapshot) by name.
   * Used before delete_sheet so Reject can restoreSheetSnapshot.
   */
  getSheetSnapshot: (name: string) => any | null;

  /**
   * Re-insert a previously captured sheet snapshot into the live workbook.
   * Used by Reject after delete_sheet.
   */
  restoreSheetSnapshot: (sheetSnapshot: any) => boolean;
}

// Wrapped in React.memo — without it, the parent re-rendering for ANY reason
// (e.g. a keystroke in a sibling chat/input, or any other unrelated state
// update in SpreadsheetScreen) re-renders this whole 2000+ line component
// for every mounted tab and, combined with an unmemoized useImperativeHandle,
// rebuilds every handle method from scratch each time. Verified against a
// real crash: HeelixNotes hit exactly this (ProjectModel/components/UniverGrid.tsx,
// ported from this file) — typing a longer chat message with the agent
// completely idle was enough to freeze/crash the tab, purely from this
// re-render cascade. The memo comparison is only useful if callers also keep
// `onUserEdit` (and the ref callback) referentially stable — see
// SpreadsheetScreen's per-tab callback caches.
/**
 * A sheet-level structure change the USER made through Univer's own UI
 * (sheet-tab "+", right-click delete/rename, duplicate) — as opposed to
 * agent mutations or our toolbar, which record their ops themselves. The
 * save mirror needs these too, or the surgical patcher meets a live sheet
 * it can't explain and falls back to the lossy full export.
 */
export type ManualSheetOp =
  | { kind: "create"; name: string }
  | { kind: "delete"; name: string }
  | { kind: "rename"; oldName: string; newName: string };

export const UniverGrid = React.memo(React.forwardRef<
  UniverGridHandle,
  {
    /** Workspace tab id — engine-mode plumbing (appliers, command routing,
     * settle) is keyed per tab since every tab's grid stays mounted. */
    tabId: string;
    workbookPath: string | null;
    /**
     * Fired when the user edits the grid DIRECTLY (typing in a cell, paste,
     * native context-menu format) — i.e. an edit not driven by our own handle
     * methods. Programmatic edits (agent, load, diff overlay, undo/redo) are
     * suppressed via a depth guard, so this only signals genuine user actions.
     */
    onUserEdit?: () => void;
    /**
     * Fired when the user adds/deletes/renames a sheet through Univer's own
     * UI. Same depth-guard rules as onUserEdit. The parent records these
     * into the save mirror's sheetOps.
     */
    onManualSheetOp?: (op: ManualSheetOp) => void;
  }
>(
  function UniverGrid({ tabId, workbookPath, onUserEdit, onManualSheetOp }, ref) {
    const containerRef = useRef<HTMLDivElement>(null);
    const univerAPIRef = useRef<any>(null);
    const univerRef = useRef<any>(null);
    /** Formula-calculation worker; terminated alongside Univer's dispose. */
    const calcWorkerRef = useRef<Worker | null>(null);
    /** @univerjs/core's covertCellValue, captured at lazy-init time. */
    const covertCellValueRef = useRef<((v: any) => any) | null>(null);
    // Bumped >0 while any handle mutator runs so the command-stream listener can
    // tell programmatic edits (agent/load/overlay/undo) from raw user edits.
    const programmaticDepthRef = useRef(0);

    /**
     * Desired state of this grid's canvas render loop. Background tabs are
     * deactivated: the render unit freezes (SheetRenderController stops the
     * rAF loop on activated$=false, so scene.render() never runs) while the
     * data model, command pipeline and IronCalc mirror keep working.
     * Mutations still mark the scene/viewports dirty, so the first frame
     * after reactivation repaints everything that changed while hidden.
     */
    const renderActiveRef = useRef(true);
    /**
     * Push renderActiveRef onto the CURRENT workbook's render unit. Must be
     * re-run after every createWorkbook — each one makes a fresh render
     * unit, and new units start activated regardless of our desired state.
     */
    const applyRenderActive = async () => {
      const univer = univerRef.current;
      const unitId = univerAPIRef.current?.getActiveWorkbook?.()?.getId?.();
      if (!univer || !unitId) return;
      try {
        const { IRenderManagerService } = await import("@univerjs/preset-sheets-core");
        const renderManager = (univer as any).__getInjector().get(IRenderManagerService);
        // The render unit is created on the unit-added stream, which can
        // lag createWorkbook by a beat — poll briefly instead of missing it.
        for (let i = 0; i < 20; i++) {
          const renderer = renderManager.getRenderById(unitId);
          if (renderer) {
            if (renderActiveRef.current) renderer.activate();
            else renderer.deactivate();
            console.log(`[univer] render loop ${renderActiveRef.current ? "resumed" : "frozen"} (${tabId})`);
            return;
          }
          await new Promise((r) => setTimeout(r, 50));
        }
        console.warn("[univer] render unit never appeared for", unitId);
      } catch (e) {
        console.warn("[univer] render loop (de)activation failed:", e);
      }
    };

    // Perf diagnostics: performance.now() of the most recent editing command
    // (user or agent). calculationStart consumes it to report the edit →
    // recalc-start latency (Univer's dirty-marking debounce + worker RPC hop).
    const lastEditPerfTsRef = useRef<number | null>(null);
    /** Engine calc-cycle state from calculationStart/End events — the only
     * reliable idle signal with the engine in a worker (the main-thread
     * facade's whenComputingCompleteAsync resolves immediately). */
    const calcActiveRef = useRef(false);
    const lastCalcStartTsRef = useRef(0);
    const lastCalcEndTsRef = useRef(0);
    // Latest onUserEdit, read via ref so the long-lived command subscription
    // never closes over a stale prop.
    const onUserEditRef = useRef(onUserEdit);
    useEffect(() => { onUserEditRef.current = onUserEdit; }, [onUserEdit]);
    const onManualSheetOpRef = useRef(onManualSheetOp);
    useEffect(() => { onManualSheetOpRef.current = onManualSheetOp; }, [onManualSheetOp]);
    /**
     * Sheet id → name map maintained from the command stream, so remove /
     * rename mutations (which carry only the subUnitId) can be reported
     * with the OLD sheet name. (Re)seeded from the live workbook whenever
     * we see an id we don't know yet.
     */
    const sheetNamesByIdRef = useRef<Map<string, string>>(new Map());
    /** Rebuild the sheet id→name map from the live workbook. */
    const syncSheetNamesById = () => {
      try {
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        const sheets = wb?.getSheets?.() ?? [];
        const map = sheetNamesByIdRef.current;
        map.clear();
        for (const s of sheets) {
          const id = s?.getSheetId?.();
          const name = s?.getSheetName?.() ?? s?.getName?.();
          if (typeof id === "string" && typeof name === "string") map.set(id, name);
        }
      } catch (e) {
        console.warn("[univer] syncSheetNamesById failed:", e);
      }
    };
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
    //
    // Since the parse moved into the import worker, this starts out NULL for
    // worker-parsed loads (the ExcelJS object graph can't cross the worker
    // boundary and holding it cost hundreds of MB on big models). It's
    // populated lazily from sourceBytesRef the first time a full-export save
    // actually needs it — see ensureExcelJsWorkbook.
    const exceljsWorkbookRef = useRef<ExcelJS.Workbook | null>(null);
    /** External-workbook / add-in formula pins from the load-time parse. */
    const externalPinsRef = useRef<ExternalPinMap | null>(null);
    /**
     * Frozen copy of the cell model as loaded from disk, keyed by sheet name
     * then "row,col". exportPatch diffs the live Univer model against this;
     * everything not in the diff stays byte-identical in the file. NOT
     * refreshed after saves — patches are idempotent, so re-sending an
     * already-saved change is harmless, while refreshing risks losing one.
     */
    const baselineCellsRef = useRef<Map<string, Map<string, BaselineCell>> | null>(null);
    /**
     * Memoized live-vs-file divergence scan (see agent/calcHealth.ts).
     * undefined = not yet computed for this workbook; null = no baseline.
     * Cached forever per load: the summary line rides in the CACHED context
     * block, so its bytes must be identical every turn.
     */
    const calcHealthRef = useRef<CalcHealth | null | undefined>(undefined);
    /** At-risk package content (charts, comments…) scanned at load time. */
    const fidelityRisksRef = useRef<FidelityRisks | null>(null);
    /**
     * Per-sheet CF / data-validation / note state as Univer serialized it
     * right after load. exportPatch compares the live state against this:
     * unchanged → surgical save (original XML parts survive untouched);
     * changed → typed fallback so the gated full export carries the edits.
     * Captured post-hydration (not from our own import data) so both sides
     * of the comparison went through the same Univer normalization.
     */
    const featureBaselineRef = useRef<Map<string, SheetFeatureState> | null>(null);
    /** The exact bytes we loaded, for post-export loss audits. */
    const sourceBytesRef = useRef<Uint8Array | null>(null);
    /** Per-sheet column/row outline (grouping) levels from the loaded file. */
    const outlinesRef = useRef<Record<string, SheetOutline> | null>(null);
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

        const [
          { createUniver, LocaleType, defaultTheme, merge },
          { UniverSheetsCorePreset },
          { UniverSheetsFindReplacePreset },
          { UniverSheetsFilterPreset },
          { UniverSheetsDrawingPreset },
          { UniverSheetsConditionalFormattingPreset },
          { UniverSheetsDataValidationPreset },
          { UniverSheetsSortPreset },
          { UniverSheetsNotePreset },
          univerCore,
        ] = await Promise.all([
          import("@univerjs/presets"),
          import("@univerjs/preset-sheets-core"),
          import("@univerjs/preset-sheets-find-replace"),
          import("@univerjs/preset-sheets-filter"),
          import("@univerjs/preset-sheets-drawing"),
          import("@univerjs/preset-sheets-conditional-formatting"),
          import("@univerjs/preset-sheets-data-validation"),
          import("@univerjs/preset-sheets-sort"),
          import("@univerjs/preset-sheets-note"),
          import("@univerjs/core"),
        ]);
        // Same primitive→ICellData conversion the facade's setValue uses
        // (formula detection, numfmt-pattern string parsing) so batched
        // setCells writes behave identically to per-cell setCell.
        covertCellValueRef.current = (univerCore as any).covertCellValue ?? null;

        // Locales — each preset ships its own en-US bundle; merge them so the
        // Find & Replace dialog and Filter UI have their strings too.
        const [enUS, frEnUS, filterEnUS, drawingEnUS, cfEnUS, dvEnUS, sortEnUS, noteEnUS] = await Promise.all([
          import("@univerjs/preset-sheets-core/locales/en-US").then((m) => m.default),
          import("@univerjs/preset-sheets-find-replace/locales/en-US").then((m) => m.default),
          import("@univerjs/preset-sheets-filter/locales/en-US").then((m) => m.default),
          import("@univerjs/preset-sheets-drawing/locales/en-US").then((m) => m.default),
          import("@univerjs/preset-sheets-conditional-formatting/locales/en-US").then((m) => m.default),
          import("@univerjs/preset-sheets-data-validation/locales/en-US").then((m) => m.default),
          import("@univerjs/preset-sheets-sort/locales/en-US").then((m) => m.default),
          import("@univerjs/preset-sheets-note/locales/en-US").then((m) => m.default),
        ]);

        if (cancelled || !containerRef.current) return;

        // Formula engine runs in a dedicated Web Worker: passing workerURL
        // flips the main-thread engine to notExecuteFormula and proxies
        // recalculation over RPC, so bulk agent edits and big dependency
        // chains no longer block painting (the "is it stuck?" freeze).
        // Engine mode: the worker is named so it registers the formula engine
        // with notExecuteFormula — Univer stops calculating and IronCalc's
        // values (applied as mutations) are what the grid shows. Two literal
        // constructor calls keep vite's static worker detection happy.
        {
          // Loud, unambiguous status line — the mode is decided HERE, at
          // worker creation, and a stale kill switch has already cost a
          // debugging round.
          const on = ironcalcEngineEnabled();
          let reason = "default";
          try {
            if (localStorage.getItem("gridpath.ironcalcEngine") === "0") {
              reason = 'kill switch: localStorage gridpath.ironcalcEngine = "0" — run localStorage.removeItem("gridpath.ironcalcEngine") + reload to re-enable';
            }
          } catch { /* no localStorage */ }
          console.log(
            `%c[engine] ${on ? "IRONCALC" : "UNIVER"} calculates this session (${reason})`,
            `font-weight:bold;color:${on ? "#0a7" : "#c60"}`,
          );
        }
        const calcWorker = ironcalcEngineEnabled()
          ? new Worker(new URL("./univer.worker.ts", import.meta.url), {
              type: "module",
              name: "univer-calc-noformula",
            })
          : new Worker(new URL("./univer.worker.ts", import.meta.url), {
              type: "module",
              name: "univer-calc",
            });
        calcWorkerRef.current = calcWorker;

        const { univer, univerAPI } = createUniver({
          locale: LocaleType.EN_US,
          locales: {
            [LocaleType.EN_US]: merge({}, enUS, frEnUS, filterEnUS, drawingEnUS, cfEnUS, dvEnUS, sortEnUS, noteEnUS),
          },
          theme: defaultTheme,
          presets: [
            UniverSheetsCorePreset({
              container: containerRef.current,
              workerURL: calcWorker,
              // GridPath ships its own dark "Home" formatting ribbon above the
              // grid (FormatToolbar), so Univer's built-in toolbar (the
              // Start / Formulas / Data tabs and their buttons) just duplicated
              // those controls. Hide it. `toolbar: false` keeps the formula bar
              // (it lives in the header), sheet tabs (footer) and the
              // right-click context menu — none of which we duplicate.
              toolbar: false,
            }),
            // Native Excel-style Find & Replace (⌘F / ⌘H + right-click menu).
            // Pure in-grid feature; nothing to persist on save.
            UniverSheetsFindReplacePreset(),
            // AutoFilter dropdowns. The filter RANGE is persisted to xlsx via
            // the save mirror (ExcelJS autoFilter); active filter criteria are
            // in-session only.
            UniverSheetsFilterPreset(),
            // Floating images. Display-only: imported pictures are rendered
            // from the ExcelJS media store; the file's drawing parts are
            // preserved untouched by the surgical save, so nothing here is
            // ever written back.
            UniverSheetsDrawingPreset(),
            // Conditional formatting, data validation and cell notes render
            // from snapshot resources seeded at load (see featureIo.ts).
            // In-session EDITS to them can't ride the surgical patch — the
            // save path detects the drift and routes to the gated full
            // export, whose mirror writes them back (applyFeatureMirror).
            UniverSheetsConditionalFormattingPreset(),
            UniverSheetsDataValidationPreset(),
            // Right-click + filter-dropdown sort. Sort is an action, not
            // state — nothing to persist beyond the cells it moves.
            UniverSheetsSortPreset(),
            UniverSheetsNotePreset(),
          ],
        });

        univerRef.current = univer;
        univerAPIRef.current = univerAPI;

        // Raise the depth guard: createWorkbook runs outside any wrapped
        // handle method, and any sheet mutations it fires internally must
        // not be mistaken for manual user sheet ops.
        programmaticDepthRef.current++;
        try {
          univerAPI.createWorkbook(blankWorkbook());
        } finally {
          programmaticDepthRef.current--;
        }
        syncSheetNamesById();
        // A tab can be opened in the background (session restore) — apply
        // the desired render-loop state the moment the first workbook exists.
        void applyRenderActive();

        // Univer's paste pipeline ends with an auto-row-height pass over the
        // ENTIRE pasted range, text-layout-measuring every row that has no
        // rowData entry (~2.5ms each). A whole-column selection spans the
        // sheet's 10k-row capacity, not its used range, so pasting one column
        // of CVNA.xlsx stalled ~21s to compute "no heights changed". Excel
        // doesn't auto-fit on paste at all; keep the nicety only for pastes
        // small enough that the pass is imperceptible.
        try {
          const { ISheetClipboardService } = await import("@univerjs/sheets-ui");
          const clip = (univer as any).__getInjector().get(ISheetClipboardService) as any;
          const origAutoHeight = clip._getPastedRangeAutoHeightMutation?.bind(clip);
          if (origAutoHeight) {
            const MAX_AUTO_HEIGHT_ROWS = 1000;
            clip._getPastedRangeAutoHeightMutation = (unitId: string, subUnitId: string, pastedRange: any) => {
              const rowCount = pastedRange?.rows?.length ?? 0;
              if (rowCount > MAX_AUTO_HEIGHT_ROWS) {
                console.log(`[univer] paste auto-height skipped (${rowCount}-row range)`);
                return null;
              }
              return origAutoHeight(unitId, subUnitId, pastedRange);
            };
          } else {
            console.warn("[univer] paste auto-height guard: method not found (Univer upgrade?) — pastes of whole columns will be slow");
          }
        } catch (e) {
          console.warn("[univer] paste auto-height guard failed to install:", e);
        }

        // Engine mode: computed values arrive as one raw set-range-values
        // MUTATION per sheet — merges into existing cell data (keeps f/s),
        // no undo entry, and invisible to the command hook above (which only
        // listens to sheet.command.*), so no feedback loop.
        registerIroncalcApplier(tabId, (sheetName, cells) => {
          try {
            const wb = univerAPI.getActiveWorkbook?.();
            const sheet = wb?.getSheetByName?.(sheetName);
            if (!wb || !sheet) return false;
            const unitId = wb.getId?.();
            const subUnitId = sheet.getSheetId?.();
            if (!unitId || !subUnitId) return false;
            const cellValue: Record<number, Record<number, unknown>> = {};
            for (const c of cells) {
              (cellValue[c.row0] ??= {})[c.col0] = { v: c.v };
            }
            univerAPI.executeCommand(
              "sheet.mutation.set-range-values",
              {
                unitId,
                subUnitId,
                cellValue,
                // Marks this write as ours: the engine bridge buffers
                // set-range-values MUTATIONS (that's how undo/redo replays
                // are caught) and must skip its own write-backs.
                __ironcalc: true,
              },
              // CRITICAL: onlyLocal makes the sheets-formula plugin's
              // beforeCommandExecuted interceptor skip this mutation. Without
              // it the plugin treats "cell received a value" as "formula
              // replaced by a literal" and dispatches a follow-up mutation
              // writing f:null — physically DELETING the formulas we are
              // computing values for. Same options Univer's own engine uses
              // for its result write-backs.
              { onlyLocal: true, fromFormula: true } as any,
            );
            return true;
          } catch (e) {
            console.warn("[ironcalc-engine] grid apply failed:", e);
            return false;
          }
        });

        // Engine mode RESYNC source: a full content dump of every sheet
        // (formula text + current raw value per cell). The bridge calls this
        // when it hits an operation it cannot mirror (sort, range move,
        // unrecognized command) and reseeds the IronCalc model from it — the
        // grid is the source of truth for content. MUST be getRawValues():
        // getValues() runs the numfmt interceptor and returns the DISPLAY
        // string ("72.35%", "$8.31") for any number-formatted cell, which
        // resync would then seed into IronCalc as text literals — every
        // formula over them collapses (SUM=0, AVERAGE=#DIV/0!) and the
        // write-backs mark grid cells "number stored as text".
        registerIroncalcSnapshotter(tabId, () => {
          try {
            const wb = univerAPIRef.current?.getActiveWorkbook?.();
            if (!wb) return null;
            const out: GridResyncSheet[] = [];
            for (const sheet of wb.getSheets?.() ?? []) {
              const name = sheet?.getSheetName?.();
              if (!name) return null;
              const cells: GridResyncSheet["cells"] = [];
              const lastRow = sheet.getLastRow?.() ?? -1;
              const lastCol = sheet.getLastColumn?.() ?? -1;
              if (lastRow >= 0 && lastCol >= 0) {
                const range = sheet.getRange?.(0, 0, lastRow + 1, lastCol + 1);
                const values: any[][] = range?.getRawValues?.() ?? range?.getValues?.() ?? [];
                const formulas: any[][] = range?.getFormulas?.() ?? [];
                for (let r = 0; r <= lastRow; r++) {
                  const vRow = values[r] ?? [];
                  const fRow = formulas[r] ?? [];
                  for (let c = 0; c <= lastCol; c++) {
                    const f =
                      typeof fRow[c] === "string" && fRow[c].length > 0 ? (fRow[c] as string) : null;
                    const v = vRow[c];
                    if (f === null && (v === null || v === undefined || v === "")) continue;
                    cells.push({ row: r + 1, column: c + 1, formula: f, value: v ?? null });
                  }
                }
              }
              out.push({ name, cells });
            }
            return out;
          } catch (e) {
            console.warn("[ironcalc-engine] grid snapshot for resync failed:", e);
            return null;
          }
        });

        // IronCalc shadow cross-check reads Univer values through the facade.
        registerUniverValueGetter((sheetName, row0, col0) => {
          try {
            const range = univerAPI
              .getActiveWorkbook?.()
              ?.getSheetByName?.(sheetName)
              ?.getRange?.(row0, col0);
            // getCellData().v is the raw stored value; getValue() can return
            // the number-formatted display string ("$70.00", "(65.2%)").
            const cd = range?.getCellData?.();
            if (cd && cd.v !== undefined && cd.v !== null) return cd.v;
            return range?.getValue?.() ?? null;
          } catch {
            return null;
          }
        });
        // The cross-check must wait for Univer's own recalc (5-15s on heavy
        // files) or it compares against stale values. whenComputingCompleteAsync
        // resolves immediately with the engine in a worker, so wait on the
        // calculationStart/End EVENTS instead: resolve once a cycle ends after
        // the waiter started, or once it's clear no recalc is coming.
        registerUniverIdleWaiter(async (timeoutMs) => {
          const since = performance.now();
          const deadline = since + timeoutMs;
          const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));
          while (performance.now() < deadline) {
            if (lastCalcEndTsRef.current > since) return;
            // No cycle started within 2.5s and none is running: the edit
            // didn't trigger a recalc (or it finished before we subscribed).
            if (
              performance.now() - since > 2500 &&
              !calcActiveRef.current &&
              lastCalcStartTsRef.current < since
            ) {
              return;
            }
            await sleep(150);
          }
        });

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

              // Mirror committed edits into IronCalc: engine mode drives the
              // grid from its results; shadow mode only measures. Both no-op
              // unless their localStorage flag is set.
              try {
                if (ironcalcEngineEnabled()) {
                  ironcalcEngineOnCommand(tabId, cmd, sheetNamesByIdRef.current);
                } else {
                  ironcalcShadowOnCommand(cmd, sheetNamesByIdRef.current);
                }
              } catch (e) {
                console.warn("[ironcalc] command hook failed:", e);
              }

              // Raw user-edit detection. A genuine user edit dispatches a
              // top-level `sheet.command.*` editing command while no handle
              // mutator is running (depth 0). Recalculations fire
              // `sheet.mutation.*` (excluded by the prefix), and agent / load /
              // overlay / undo edits all run through guarded handle methods
              // (depth > 0), so neither trips this.
              if (cmd.id.startsWith("sheet.command.") && isEditingCommand(cmd.id)) {
                // Any editing command (user or agent) may kick off a recalc;
                // stamp it so the calc-perf log can attribute the latency.
                lastEditPerfTsRef.current = performance.now();
                if (programmaticDepthRef.current === 0) {
                  onUserEditRef.current?.();
                }
              }

              // Manual sheet structure ops (Univer's sheet-tab UI: "+",
              // right-click delete/rename/duplicate). We watch the MUTATIONS,
              // not the commands, because the mutation params carry the final
              // generated sheet name ("Sheet1") and, for removes, the old
              // name. Agent / load / undoBatch paths run through guarded
              // handle methods (depth > 0) and are excluded — batches record
              // their own sheet ops.
              if (cmd.id.startsWith("sheet.mutation.")) {
                // The id→name map is maintained at ANY depth (agent renames
                // must not leave it stale); ops are only REPORTED as manual
                // at depth 0 — batches record their own sheet ops.
                const manual = programmaticDepthRef.current === 0;
                const map = sheetNamesByIdRef.current;
                if (cmd.id === "sheet.mutation.insert-sheet") {
                  const name = cmd.params?.sheet?.name;
                  const id = cmd.params?.sheet?.id;
                  if (typeof name === "string" && name) {
                    if (typeof id === "string") map.set(id, name);
                    if (manual) onManualSheetOpRef.current?.({ kind: "create", name });
                  }
                } else if (cmd.id === "sheet.mutation.remove-sheet") {
                  const id = cmd.params?.subUnitId;
                  const name = cmd.params?.subUnitName ?? (typeof id === "string" ? map.get(id) : undefined);
                  if (typeof id === "string") map.delete(id);
                  if (manual && typeof name === "string" && name) {
                    onManualSheetOpRef.current?.({ kind: "delete", name });
                  }
                } else if (cmd.id === "sheet.mutation.set-worksheet-name") {
                  const id = cmd.params?.subUnitId;
                  const newName = cmd.params?.name;
                  const oldName = typeof id === "string" ? map.get(id) : undefined;
                  if (typeof id === "string" && typeof newName === "string") map.set(id, newName);
                  if (manual && typeof newName === "string" && newName && typeof oldName === "string" && oldName !== newName) {
                    onManualSheetOpRef.current?.({ kind: "rename", oldName, newName });
                  }
                  // Unknown old name (map never seeded): record nothing
                  // rather than a bogus rename; the map is now warm.
                }
              }

              if (!cmd.id.includes("set-selections") && !cmd.id.endsWith("SetSelectionsOperation")) return;
              // Keep the id→name map warm off the selection stream: it fires
              // on every click/drag, well before any first manual sheet op.
              if (sheetNamesByIdRef.current.size === 0) syncSheetNamesById();
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
              // Sheet resolution order: the event's own subUnitId (authoritative
              // — it names the sheet the selection actually happened on, via the
              // id→name map kept warm above), then the active-sheet probe, then
              // the previous selection's sheet. The old probe-only version
              // returned "" during sheet-switch races, which shipped sheetless
              // "!A8" labels to the agent and broke its cell-value lookups.
              const subUnitId = cmd.params?.subUnitId;
              const sheetName =
                (typeof subUnitId === "string" ? sheetNamesByIdRef.current.get(subUnitId) : undefined) ||
                univerAPI.getActiveWorkbook?.()?.getActiveSheet?.()?.getName?.() ||
                liveSelectionRef.current?.sheet ||
                "";
              if (!sheetName) return; // a sheetless selection is worse than a stale one

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

        // Calc-cycle timing: one console line per recalculation with the
        // edit→start gap (dirty-marking debounce + worker RPC), the engine
        // time, and the formula count — so "the sheet feels slow" can be
        // attributed to a specific stage instead of guessed at.
        try {
          const formula = univerAPI.getFormula?.();
          if (formula?.calculationStart && formula?.calculationEnd) {
            let calcStartTs = 0;
            let editToStartMs: number | null = null;
            let formulaCount = 0;
            let cycleIndex = 0;
            formula.calculationStart((forced: boolean) => {
              calcStartTs = performance.now();
              calcActiveRef.current = true;
              lastCalcStartTsRef.current = calcStartTs;
              const editTs = lastEditPerfTsRef.current;
              editToStartMs = editTs != null ? calcStartTs - editTs : null;
              lastEditPerfTsRef.current = null;
              formulaCount = 0;
              cycleIndex = 0;
              if (forced) console.log("[calc-perf] forced full recalculation started");
            });
            formula.calculationProcessing?.((stage: any) => {
              const n = stage?.totalFormulasToCalculate;
              if (typeof n === "number" && n > formulaCount) formulaCount = n;
              const c = stage?.formulaCycleIndex;
              if (typeof c === "number" && c > cycleIndex) cycleIndex = c;
            });
            formula.calculationEnd(() => {
              calcActiveRef.current = false;
              lastCalcEndTsRef.current = performance.now();
              if (!calcStartTs) return;
              const calcMs = performance.now() - calcStartTs;
              calcStartTs = 0;
              const trigger = editToStartMs != null
                ? `edit→start ${editToStartMs.toFixed(0)}ms | `
                : "";
              const iter = cycleIndex > 0 ? ` | iterations ${cycleIndex + 1}` : "";
              console.log(
                `[calc-perf] ${trigger}calc ${calcMs.toFixed(0)}ms | formulas ${formulaCount}${iter}`,
              );
            });
          }
        } catch (e) {
          console.warn("[calc-perf] instrumentation setup failed:", e);
        }

        console.log("[univer] grid ready");
        readyResolve.current?.(univerAPI);
      })().catch((e) => {
        console.error("[univer] init failed:", e);
      });

      return () => {
        cancelled = true;
        unregisterIroncalcApplier(tabId);
        unregisterIroncalcSnapshotter(tabId);
        // Defer Univer's dispose() out of React's commit phase. Disposing
        // synchronously triggers React's "Attempted to synchronously unmount
        // a root while React was already rendering" warning because Univer
        // unmounts its internal React roots, which collides with our parent
        // tree still finishing its own unmount.
        const u = univerRef.current;
        univerRef.current = null;
        univerAPIRef.current = null;
        exceljsWorkbookRef.current = null;
        externalPinsRef.current = null;
        const cw = calcWorkerRef.current;
        calcWorkerRef.current = null;
        setTimeout(() => {
          try { u?.dispose?.(); } catch (e) { console.warn("[univer] dispose error:", e); }
          try { cw?.terminate(); } catch {}
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

    // macOS Ctrl+click → right-click, ⌘+click → add-to-selection (the
    // Excel-for-Mac conventions). Univer hard-codes the Windows mapping:
    // its context menus (grid, row/col headers) open ONLY on pointer
    // button 2 (SheetContextMenuRenderController checks `event.button ===
    // 2`), and its multi-range selection modifier is ctrlKey — so on Mac a
    // Ctrl+click grows the selection instead of opening a menu, and
    // ⌘+click does nothing special. Window-capture listeners (same trick
    // as the arrow-key clamp above) rewrite the events before Univer sees
    // them. Scoped to CANVAS targets only: the grid and headers are
    // canvas-rendered and consume pointer events exclusively, while DOM
    // parts (sheet-tab bar, panels) open their menus from the native
    // `contextmenu` event, which Ctrl+click already fires — those keep
    // their default behavior.
    useEffect(() => {
      if (!/mac/i.test(navigator.platform || navigator.userAgent)) return;
      const SYNTH = "__gpSyntheticPointer";
      // pointerId of a converted Ctrl+click press, so its release converts too.
      let convertedPointerId: number | null = null;

      const isGridCanvas = (target: EventTarget | null): target is HTMLElement => {
        const container = containerRef.current;
        return (
          !!container &&
          target instanceof HTMLElement &&
          target.tagName === "CANVAS" &&
          container.contains(target)
        );
      };

      const redispatch = (e: PointerEvent, overrides: PointerEventInit) => {
        const clone = new PointerEvent(e.type, {
          bubbles: true,
          cancelable: true,
          composed: true,
          view: window,
          detail: e.detail,
          screenX: e.screenX,
          screenY: e.screenY,
          clientX: e.clientX,
          clientY: e.clientY,
          pointerId: e.pointerId,
          pointerType: e.pointerType || "mouse",
          isPrimary: e.isPrimary,
          ...overrides,
        });
        (clone as any)[SYNTH] = true;
        e.target?.dispatchEvent(clone);
      };

      const onPointerDown = (e: PointerEvent) => {
        if ((e as any)[SYNTH] || !isGridCanvas(e.target) || e.button !== 0) return;
        if (e.ctrlKey && !e.metaKey) {
          // Ctrl+click → right-press. preventDefault also cancels the
          // gesture's compatibility mouse events, so Univer's engine sees
          // exactly one (synthetic) button-2 press.
          e.preventDefault();
          e.stopImmediatePropagation();
          convertedPointerId = e.pointerId;
          redispatch(e, { button: 2, buttons: 2, ctrlKey: false });
        } else if (e.metaKey && !e.ctrlKey) {
          // ⌘+click → Univer's ctrl-flavored add-to-selection. No
          // preventDefault: focus handling stays native.
          e.stopImmediatePropagation();
          redispatch(e, { button: 0, buttons: e.buttons, ctrlKey: true, metaKey: false });
        }
      };

      const onPointerUp = (e: PointerEvent) => {
        if ((e as any)[SYNTH] || convertedPointerId === null || e.pointerId !== convertedPointerId) return;
        convertedPointerId = null;
        // Swallow the release unconditionally once we own this pointerId —
        // a trusted button-0 up with no matching down must never leak to
        // Univer, even when the target below can't take the redispatch.
        e.stopImmediatePropagation();
        if (e.target instanceof HTMLElement) {
          redispatch(e, { button: 2, buttons: 0, ctrlKey: false });
        }
      };

      // The trusted Ctrl+click gesture still fires a native `contextmenu`
      // on WebKit — suppress it over the canvas so the WebView's default
      // menu can't stack on top of Univer's. Genuine right-clicks
      // (ctrlKey false) are left exactly as they behave today.
      const onContextMenu = (e: MouseEvent) => {
        if (!isGridCanvas(e.target) || !e.ctrlKey) return;
        e.preventDefault();
        e.stopImmediatePropagation();
      };

      window.addEventListener("pointerdown", onPointerDown, true);
      window.addEventListener("pointerup", onPointerUp, true);
      window.addEventListener("contextmenu", onContextMenu, true);
      return () => {
        window.removeEventListener("pointerdown", onPointerDown, true);
        window.removeEventListener("pointerup", onPointerUp, true);
        window.removeEventListener("contextmenu", onContextMenu, true);
      };
    }, []);

    React.useImperativeHandle(ref, () => {
      /**
       * The ExcelJS workbook for full-export saves. Worker-parsed loads
       * don't retain one (see exceljsWorkbookRef), so re-parse the retained
       * source bytes on first need and re-attach the load-time formula pins.
       * Costs a few seconds on very large files, but only on the full-export
       * fallback path — surgical saves never call this. Untitled workbooks
       * have no source bytes and correctly return null (fresh-build path).
       */
      const ensureExcelJsWorkbook = async (): Promise<ExcelJS.Workbook | null> => {
        if (exceljsWorkbookRef.current) return exceljsWorkbookRef.current;
        const bytes = sourceBytesRef.current;
        if (!bytes) return null;
        const wb = new ExcelJS.Workbook();
        const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
        await wb.xlsx.load(ab as ArrayBuffer);
        externalPinsByWorkbook.set(wb, externalPinsRef.current ?? new Map());
        exceljsWorkbookRef.current = wb;
        return wb;
      };

      const handle: UniverGridHandle = {
      loadBytes: async (bytes: Uint8Array, onProgress?: (stage: string) => void) => {
        const api = await waitForApi();
        if (!api) {
          console.warn("[univer] loadBytes: API never became ready");
          return;
        }
        // Yield a frame after each stage report so the overlay text actually
        // paints before the next long synchronous stretch begins.
        const stage = async (label: string) => {
          onProgress?.(label);
          await new Promise<void>((r) => requestAnimationFrame(() => r()));
        };
        await stage("Parsing workbook…");
        // Heavy parse runs in a one-shot worker (see xlsxImport.ts) so the
        // UI stays responsive; `excelJs` is null in that case and the save
        // path re-parses lazily via ensureExcelJsWorkbook.
        const parsed = await parseXlsxWorkbook(bytes);
        const { univerData, styleOps, outlines, featureCounts } = parsed;
        exceljsWorkbookRef.current = parsed.excelJs;
        externalPinsRef.current = new Map(parsed.pins);
        baselineCellsRef.current = buildCellBaseline(univerData);
        calcHealthRef.current = undefined;
        sourceBytesRef.current = bytes;
        fidelityRisksRef.current = await scanFidelityRisks(bytes);
        outlinesRef.current = outlines;
        // Circular references: mirror the file's iterative-calc opt-in
        // (<calcPr iterate="1" iterateCount="N">) onto Univer's engine
        // BEFORE the workbook exists, so the initial calculation pass
        // already converges LBO-style debt/interest loops the way Excel
        // does. Files without the opt-in reset to 1 pass — Excel doesn't
        // converge their circular refs either.
        try {
          const iter = await detectIterativeCalc(bytes);
          api.getFormula?.()?.setMaxIteration?.(iter ? iter.iterateCount : 1);
          console.log(
            iter
              ? `[univer] iterative calculation enabled (${iter.iterateCount} rounds)`
              : "[univer] iterative calculation off (1 pass)",
          );
        } catch (e) {
          console.warn("[univer] iterative-calc detection failed:", e);
        }
        await stage("Building grid…");
        const active = api.getActiveWorkbook?.();
        if (active?.dispose) {
          try { active.dispose(); } catch {}
        }
        api.createWorkbook(univerData);
        syncSheetNamesById();
        // New workbook = new render unit (starts activated); if this tab is
        // hidden, freeze it now so the style/image passes below skip paint.
        void applyRenderActive();
        // Resource hydration is synchronous with createWorkbook (the core's
        // resource manager loads on the unit-added stream), so the snapshot
        // read here is the normalized post-load state.
        try {
          featureBaselineRef.current = perSheetFeatureState(
            api.getActiveWorkbook?.()?.getSnapshot?.()?.resources,
          );
        } catch (e) {
          console.warn("[univer] feature baseline capture failed:", e);
          featureBaselineRef.current = null;
        }

        // Univer activates the first sheet in order — if the file starts
        // with hidden helper sheets (common in vendor models), that would
        // open the workbook staring at a hidden sheet. Activate the first
        // visible one instead.
        try {
          const wb = api.getActiveWorkbook?.();
          const order: string[] = univerData.sheetOrder ?? [];
          const firstVisibleId = order.find((sid: string) => univerData.sheets[sid]?.hidden !== 1);
          if (firstVisibleId && univerData.sheets[order[0]]?.hidden === 1) {
            const target = wb?.getSheetByName?.(univerData.sheets[firstVisibleId].name);
            target?.activate?.();
          }
        } catch (e) {
          console.warn("[univer] first-visible-sheet activation failed:", e);
        }
        await stage("Applying formats…");

        // Push every cell-level style we extracted from ExcelJS into Univer
        // via the facade so the user actually SEES the original colors,
        // fonts, number formats, alignment etc. — not just preserves them
        // on save. Run after createWorkbook so the sheets exist.
        //
        // This loop is a defensive fallback — the same styles ride inline on
        // cellData (`s`) and render via createWorkbook. Facade calls cost
        // real time each, so on style-heavy files (bank models carry 50k+
        // styled cells) the fallback is skipped rather than hanging the load.
        const styleOpsToApply = styleOps.length <= 5000 ? styleOps : [];
        if (styleOps.length > 5000) {
          console.log(`[univer] ${styleOps.length} style ops — inline styles only, facade fallback skipped`);
        }
        for (const op of styleOpsToApply) {
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
            const opHa = toFacadeHAlign(op.format.horizontal_align);
            if (opHa) range.setHorizontalAlignment?.(opHa);
            if (op.format.vertical_align) range.setVerticalAlignment?.(op.format.vertical_align);
            if (op.format.number_format) range.setNumberFormat?.(op.format.number_format);
            if (op.background) range.setBackgroundColor?.(op.background);
            if (op.borders) applyBordersToRange(range, op.borders);
          } catch (e) {
            // Per-cell style failures shouldn't abort the whole load.
            console.warn("[univer] style apply failed at", op.sheet, op.row, op.col, e);
          }
        }

        // Anchored pictures (logos, banners). Rendered best-effort after the
        // sheets exist; never blocks the load.
        let imagesInserted = 0;
        try {
          imagesInserted = await insertWorksheetImages(api, parsed.images);
        } catch (e) {
          console.warn("[univer] image rendering failed:", e);
        }

        console.log("[univer] loaded workbook via ExcelJS:", {
          sheets: Object.keys(univerData.sheets ?? {}).length,
          sheetNames: (univerData.sheetOrder ?? []).map((id: string) => univerData.sheets[id]?.name),
          stylesApplied: styleOps.length,
          imagesInserted,
          rendered: featureCounts,
          preserved: "charts / named ranges live in ExcelJS workbook; untouched CF/validation/notes parts survive surgical saves byte-identical",
        });
        if (featureCounts.cfDropped.length) {
          console.info("[univer] CF rules without a renderable mapping (preserved in file):", featureCounts.cfDropped);
        }
      },
      exportBytes: async (mirror?: SaveMirror) => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        const data = wb?.getSnapshot?.();
        return workbookToXlsxBytes(data, await ensureExcelJsWorkbook(), mirror, featureBaselineRef.current);
      },
      exportPatch: (mirror?: SaveMirror): PatchBuildResult => {
        // sourceBytesRef (not exceljsWorkbookRef) is the "did a file load"
        // signal now: worker-parsed loads legitimately have no ExcelJS
        // workbook in memory, but the patch diff only needs the baseline.
        if (!sourceBytesRef.current || !baselineCellsRef.current) {
          return { ok: false, reason: "missing_baseline", detail: "no workbook baseline captured at load" };
        }
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        const data = wb?.getSnapshot?.();
        if (!data) {
          return { ok: false, reason: "missing_baseline", detail: "no live Univer snapshot" };
        }
        // CF / data-validation / note edits live in plugin resources the
        // Rust patcher can't rewrite. Any drift from the load-time state
        // must take the gated full-export path, which mirrors them.
        try {
          const liveNames = new Map<string, string>();
          for (const id of Object.keys(data.sheets ?? {})) {
            const nm = data.sheets[id]?.name;
            if (typeof nm === "string") liveNames.set(id, nm);
          }
          const drift = featureDriftDetail(
            featureBaselineRef.current,
            perSheetFeatureState(data.resources),
            liveNames,
          );
          if (drift) {
            return { ok: false, reason: "feature_edits", detail: drift };
          }
        } catch (e) {
          console.warn("[univer] feature drift check failed:", e);
        }
        return buildWorkbookPatch(data, baselineCellsRef.current, mirror);
      },
      getFidelityRisks: () => fidelityRisksRef.current,
      getSourceBytes: () => sourceBytesRef.current,
      commitSavedBaseline: async (bytes: Uint8Array) => {
        const api = univerAPIRef.current;
        const data = api?.getActiveWorkbook?.()?.getSnapshot?.();
        if (!data) return;
        baselineCellsRef.current = buildCellBaseline(data);
        sourceBytesRef.current = bytes;
        fidelityRisksRef.current = await scanFidelityRisks(bytes);
        // The export just baked the live CF/validation/note state into the
        // file — it IS the new baseline, or every later save would fall
        // back again for edits that are already on disk.
        featureBaselineRef.current = perSheetFeatureState(data.resources);
      },
      whenReady: async () => {
        const api = await waitForApi();
        if (!api) return false;
        // The workbook is created moments after the API resolves; poll
        // briefly (≤3s) rather than failing a race by milliseconds.
        for (let i = 0; i < 60; i++) {
          if (api.getActiveWorkbook?.()) return true;
          await new Promise((r) => setTimeout(r, 50));
        }
        return !!api.getActiveWorkbook?.();
      },
      getSheetNames: () => {
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        const sheets = wb?.getSheets?.() ?? [];
        return sheets.map((s: any) => s?.getSheetName?.() ?? s?.getName?.()).filter(Boolean);
      },
      getOutlineSummary: (sheetName) => {
        const o = outlinesRef.current?.[sheetName];
        return o ? { cols: o.maxColLevel, rows: o.maxRowLevel } : { cols: 0, rows: 0 };
      },
      applyOutlineLevel: (sheetName, axis, level) => {
        const outline = outlinesRef.current?.[sheetName];
        if (!outline) return null;
        const entries = axis === "cols" ? outline.cols : outline.rows;
        if (!entries.length) return null;
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        const sheet = wb?.getSheetByName?.(sheetName);
        if (!sheet) return null;
        const hide = entries.filter(([, l]) => l >= level).map(([i]) => i);
        const show = entries.filter(([, l]) => l < level).map(([i]) => i);
        try {
          // Show BEFORE hide so a boundary index shared by runs ends hidden,
          // matching the returned op order used for the save mirror.
          for (const [start, count] of contiguousRuns(show)) {
            if (axis === "cols") sheet.showColumns?.(start, count);
            else sheet.showRows?.(start, count);
          }
          for (const [start, count] of contiguousRuns(hide)) {
            if (axis === "cols") sheet.hideColumns?.(start, count);
            else sheet.hideRows?.(start, count);
          }
        } catch (e) {
          console.warn("[univer] outline toggle failed:", e);
          return null;
        }
        return { hide, show };
      },
      setCell: (sheetName, row, col, value) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        ensureSheetCapacity(sheet, row, col);
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return;
        try {
          // Univer's setValue accepts either a primitive or { v, f } — passing
          // a formula string starting with '=' as the primitive does NOT make
          // it a formula. We unwrap that here for convenience.
          if (value === null || value === undefined) {
            // Clear contents, keep style. Bare setValue(null) fails
            // covertCellValue's guard and throws "Invalid value" — which the
            // catch below swallows, turning clears into silent no-ops (the
            // clear_range bug). Same cell shape as setCells' clear branch.
            range.setValue({ v: null, f: null, si: null, p: null });
          } else if (typeof value === "string" && value.startsWith("=")) {
            range.setValue({ f: value });
          } else {
            range.setValue(value);
          }
        } catch (e) {
          // A throwing write must not kill the caller's whole apply loop —
          // that turns one bad cell into a silent all-or-nothing failure.
          console.warn("[univer] setCell failed:", sheetName, row, col, e);
        }
      },
      setCells: (sheetName, cells) => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        const sheet = wb?.getSheetByName?.(sheetName) ?? wb?.getActiveSheet?.();
        if (!api || !wb || !sheet || cells.length === 0) return false;
        // Grow the sheet before writing: a batch landing past rowCount /
        // columnCount (fresh agent-created sheet, wide copy) must not
        // depend on the sheet's incidental capacity.
        let maxR = 0;
        let maxC = 0;
        for (const c of cells) {
          if (c.row > maxR) maxR = c.row;
          if (c.col > maxC) maxC = c.col;
        }
        ensureSheetCapacity(sheet, maxR, maxC);
        const covert = covertCellValueRef.current;
        // Sparse object matrix — SetRangeValuesCommand applies exactly these
        // cells (its realCellValue branch); holes are never touched, so this
        // is NOT a rectangle overwrite.
        const matrix: Record<number, Record<number, any>> = {};
        for (const { row, col, value } of cells) {
          let data: any;
          if (value === null || value === undefined) {
            // Clear contents, keep style (facade setValue(null) would throw).
            data = { v: null, f: null, si: null, p: null };
          } else if (covert) {
            data = covert(value);
          } else if (typeof value === "string" && value.startsWith("=")) {
            data = { f: value, v: null, p: null };
          } else {
            data = { v: value, f: null, p: null };
          }
          (matrix[row] ??= {})[col] = data;
        }
        try {
          return api.syncExecuteCommand?.("sheet.command.set-range-values", {
            unitId: wb.getId(),
            subUnitId: sheet.getSheetId(),
            value: matrix,
          }) !== false;
        } catch (e) {
          console.warn("[univer] setCells failed:", e);
          return false;
        }
      },
      getRangeData: (sheetName, sr, sc, er, ec) => {
        const range = resolveRangeRect(univerAPIRef.current, sheetName, sr, sc, er, ec);
        if (!range) return null;
        try {
          // Raw values, NOT getValues(): the numfmt interceptor turns
          // getValues() into display strings ("72.35%") for formatted cells.
          const values = range.getRawValues?.() ?? range.getValues?.();
          const rawFormulas = range.getFormulas?.();
          if (!Array.isArray(values) || !Array.isArray(rawFormulas)) return null;
          // getFormulas reports "" for formula-less cells — normalize to null
          // so callers can share getCell's { value, formula } contract.
          const formulas = rawFormulas.map((row: any[]) =>
            row.map((f: any) => (typeof f === "string" && f.length > 0 ? f : null)),
          );
          const out: { values: any[][]; formulas: (string | null)[][]; displays?: (string | null)[][] } = {
            values,
            formulas,
          };
          const displays = range.getDisplayValues?.();
          if (Array.isArray(displays)) out.displays = displays;
          return out;
        } catch (e) {
          console.warn("[univer] getRangeData failed:", sheetName, e);
          return null;
        }
      },
      whenCalculated: async (timeoutMs = 15000) => {
        if (ironcalcEngineEnabled()) {
          // Univer never calculates in engine mode — its calculationEnd
          // would never fire and this helper used to burn the full timeout
          // per agent operation. Wait for the IronCalc queue instead: once
          // drained, every mirrored edit is recalculated and its deltas are
          // applied to the grid, so readbacks see computed values.
          await ironcalcEngineSettled(tabId, timeoutMs);
          await new Promise<void>((r) => setTimeout(r, 30));
          return;
        }
        const api = univerAPIRef.current;
        const formula = api?.getFormula?.();
        // whenComputingCompleteAsync resolves immediately when the engine is
        // idle (onCalculationEnd would wait for a NEXT end event that may
        // never come). Deprecated upstream but semantically what we need.
        if (typeof formula?.whenComputingCompleteAsync !== "function") return;
        const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));
        try {
          // Right after a write the engine hasn't flipped to "computing"
          // yet (worker RPC hop), so an immediate check would falsely
          // report idle. Grace beat, wait, then re-check once to catch
          // chained dirty passes.
          await sleep(120);
          await formula.whenComputingCompleteAsync(timeoutMs);
          await sleep(60);
          await formula.whenComputingCompleteAsync(timeoutMs);
        } catch { /* timeout — proceed with whatever is computed */ }
      },
      getFileSavedCell: (sheetName, row, col) =>
        baselineCellsRef.current?.get(sheetName)?.get(`${row},${col}`) ?? null,
      getCalcHealth: async () => {
        if (calcHealthRef.current !== undefined) return calcHealthRef.current;
        const baseline = baselineCellsRef.current;
        if (!baseline) return null; // untitled tab — no disk baseline, nothing to compare
        try {
          // Let the worker engine settle so an early first prompt doesn't
          // cache not-yet-computed cells as "divergent".
          await handle.whenCalculated(20000);
          const snapshot = univerAPIRef.current?.getActiveWorkbook?.()?.getSnapshot?.();
          if (!snapshot) return null; // transient not-ready — retry next call
          calcHealthRef.current = computeCalcHealth(snapshot, baseline);
        } catch (e) {
          console.warn("[univer] calc-health scan failed:", e);
          calcHealthRef.current = null;
        }
        return calcHealthRef.current ?? null;
      },
      getCell: (sheetName, row, col) => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return null;
        try {
          // Raw stored value FIRST: getValue() runs the numfmt interceptor
          // and returns the display string ("72.35%") for number-formatted
          // cells, which sent the agent down "the data is text" rabbit
          // holes and leaked display strings into old_value captures (lossy
          // undo/reject) and copy_range. The formatted rendering is still
          // surfaced, as a separate `display` field.
          const cd = range.getCellData?.();
          const raw = cd && cd.v !== undefined && cd.v !== null ? cd.v : range.getValue?.();
          const f = range.getFormula?.();
          if ((raw === undefined || raw === null) && !f) return null;
          const value = raw ?? null;
          const out: { value: any; display?: string; formula: string | null } = {
            value,
            formula: f ?? null,
          };
          const d = range.getDisplayValue?.();
          if (typeof d === "string" && d !== "" && value !== null && d !== String(value)) {
            out.display = d;
          }
          return out;
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
        ensureSheetCapacity(resolveSheet(univerAPIRef.current, sheetName), row, col);
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return;
        applyFormatToFRange(range, format);
      },
      setRangeFormat: (sheetName, sr, sc, er, ec, format) => {
        ensureSheetCapacity(resolveSheet(univerAPIRef.current, sheetName), er, ec);
        const range = resolveRangeRect(univerAPIRef.current, sheetName, sr, sc, er, ec);
        if (!range) return;
        applyFormatToFRange(range, format);
      },
      setCellBorders: (sheetName, row, col, borders) => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        if (!range) return;
        // applyBordersToRange handles the facade-shape fallbacks. Clearing
        // (all sides null) is best-effort in-app; the saved xlsx is exact.
        applyBordersToRange(range, borders);
      },
      setRangeBorders: (sheetName, sr, sc, er, ec, kind) => {
        const range = resolveRangeRect(univerAPIRef.current, sheetName, sr, sc, er, ec);
        if (!range || typeof (range as any).setBorder !== "function") return;
        try {
          if (kind === "none") {
            (range as any).setBorder("none", 0);
          } else if (kind === "all") {
            (range as any).setBorder("all", borderStyleToUniver("thin"), "#000000");
          } else {
            (range as any).setBorder("outside", borderStyleToUniver("thin"), "#000000");
          }
        } catch (e) {
          console.warn("[univer] setRangeBorders failed:", e);
        }
      },
      undo: () => {
        try { void univerAPIRef.current?.undo?.(); } catch (e) { console.warn("[univer] undo failed:", e); }
      },
      redo: () => {
        try { void univerAPIRef.current?.redo?.(); } catch (e) { console.warn("[univer] redo failed:", e); }
      },
      containsFocus: () => {
        const el = containerRef.current;
        const active = document.activeElement;
        return !!(el && active && el.contains(active));
      },
      openConditionalFormatting: () => {
        try {
          univerAPIRef.current?.executeCommand?.("sheet.operation.open.conditional.formatting.panel");
        } catch (e) {
          console.warn("[univer] openConditionalFormatting failed:", e);
        }
      },
      openDataValidation: () => {
        try {
          univerAPIRef.current?.executeCommand?.("data-validation.operation.open-validation-panel");
        } catch (e) {
          console.warn("[univer] openDataValidation failed:", e);
        }
      },
      openFindReplace: (replace) => {
        try {
          univerAPIRef.current?.executeCommand?.(
            replace ? "ui.operation.open-replace-dialog" : "ui.operation.open-find-dialog",
          );
        } catch (e) {
          console.warn("[univer] openFindReplace failed:", e);
        }
      },
      toggleFilter: (sheetName, sr, sc, er, ec): boolean => {
        try {
          const sheet = resolveSheet(univerAPIRef.current, sheetName);
          const existing = sheet?.getFilter?.();
          if (existing) {
            existing.remove?.();
            return false;
          }
          const range = resolveRangeRect(univerAPIRef.current, sheetName, sr, sc, er, ec);
          range?.createFilter?.();
          return true;
        } catch (e) {
          console.warn("[univer] toggleFilter failed:", e);
          return false;
        }
      },
      sortRange: (sheetName, sr, sc, er, ec, keyCol, direction) => {
        const range = resolveRangeRect(univerAPIRef.current, sheetName, sr, sc, er, ec);
        if (!range) return;
        try {
          // getCellDatas carries value + style per cell so formatting moves
          // with the row (and reads raw v via getCellRaw). The fallback must
          // be getRawValues — getValues returns numfmt display strings,
          // which would sort percent columns as text and rewrite them lossily.
          const grid: any[][] =
            typeof (range as any).getCellDatas === "function"
              ? (range as any).getCellDatas()
              : ((range as any).getRawValues?.() ?? (range as any).getValues?.());
          if (!Array.isArray(grid) || grid.length < 2) return; // nothing to reorder

          const keyIdx = Math.max(0, Math.min(ec, keyCol) - sc);
          const dir = direction === "desc" ? -1 : 1;
          const rawOf = (cell: any): unknown => {
            if (cell == null) return null;
            return typeof cell === "object" && "v" in cell ? (cell as any).v : cell;
          };
          const isEmpty = (v: unknown) => v === null || v === undefined || v === "";

          const sorted = [...grid].sort((rowA, rowB) => {
            const av = rawOf(rowA?.[keyIdx]);
            const bv = rawOf(rowB?.[keyIdx]);
            // Blanks always sink to the bottom regardless of direction.
            if (isEmpty(av) && isEmpty(bv)) return 0;
            if (isEmpty(av)) return 1;
            if (isEmpty(bv)) return -1;
            const an = typeof av === "number" ? av : Number(av);
            const bn = typeof bv === "number" ? bv : Number(bv);
            const bothNumeric = !Number.isNaN(an) && !Number.isNaN(bn);
            if (bothNumeric) return (an - bn) * dir;
            return String(av).localeCompare(String(bv), undefined, { numeric: true, sensitivity: "base" }) * dir;
          });

          (range as any).setValues(sorted);
        } catch (e) {
          console.warn("[univer] sortRange failed:", e);
        }
      },
      defineName: (name, ref) => {
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        if (!wb) return;
        try {
          // insertDefinedName throws / no-ops on a duplicate name, so drop
          // any existing entry first to get replace semantics.
          if (wb.getDefinedName?.(name)) wb.deleteDefinedName?.(name);
          wb.insertDefinedName?.(name, ref);
        } catch (e) {
          console.warn("[univer] defineName failed:", name, ref, e);
        }
      },
      deleteName: (name) => {
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        if (!wb) return;
        try {
          if (wb.getDefinedName?.(name)) wb.deleteDefinedName?.(name);
        } catch (e) {
          console.warn("[univer] deleteName failed:", name, e);
        }
      },
      getDefinedNameRef: (name) => {
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        if (!wb) return null;
        try {
          const dn = wb.getDefinedName?.(name);
          return dn?.getFormulaOrRefString?.() ?? null;
        } catch {
          return null;
        }
      },
      captureCellBand: (sheetName, startRow, startCol, endRow, endCol) => {
        const out: Array<{
          row: number;
          col: number;
          value: any;
          formula: string | null;
          format: CellFormatShape | null;
        }> = [];
        // Cap the band so a pathological "delete 5000 rows" doesn't freeze
        // the UI capturing every empty cell. We only keep non-empty cells.
        const maxRows = Math.min(endRow, startRow + 2000);
        const maxCols = Math.min(endCol, startCol + 200);
        for (let r = startRow; r <= maxRows; r++) {
          for (let c = startCol; c <= maxCols; c++) {
            const range = resolveRange(univerAPIRef.current, sheetName, r, c);
            if (!range) continue;
            let value: any = null;
            let formula: string | null = null;
            try {
              value = range.getValue?.() ?? null;
              formula = range.getFormula?.() ?? null;
            } catch {
              continue;
            }
            if ((value === null || value === undefined || value === "") && !formula) continue;
            // Reuse getCellFormat logic inline via the handle once built —
            // call the style reader directly here to avoid circular refs.
            const format = readCellFormat(range);
            out.push({
              row: r,
              col: c,
              value: value ?? null,
              formula,
              format: Object.keys(format).length > 0 ? format : null,
            });
          }
        }
        return out;
      },
      getFreezePanes: (sheetName) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        if (!sheet) return { freezeRows: 0, freezeCols: 0 };
        try {
          const freeze = sheet.getFreeze?.() ?? sheet.getFrozen?.();
          if (freeze && typeof freeze === "object") {
            return {
              freezeRows: Number(freeze.ySplit ?? freeze.startRow ?? freeze.rows ?? 0) || 0,
              freezeCols: Number(freeze.xSplit ?? freeze.startColumn ?? freeze.cols ?? 0) || 0,
            };
          }
          const rows = Number(sheet.getFrozenRows?.() ?? sheet.getFreezeRows?.() ?? 0) || 0;
          const cols = Number(sheet.getFrozenColumns?.() ?? sheet.getFreezeColumns?.() ?? 0) || 0;
          return { freezeRows: rows, freezeCols: cols };
        } catch {
          return { freezeRows: 0, freezeCols: 0 };
        }
      },
      getSheetSnapshot: (name) => {
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        const snap = wb?.getSnapshot?.();
        if (!snap?.sheets) return null;
        try {
          for (const id of Object.keys(snap.sheets)) {
            const sheet = snap.sheets[id];
            if (sheet?.name === name) {
              // Deep-ish clone so later mutations don't mutate the captured copy.
              return JSON.parse(JSON.stringify(sheet));
            }
          }
        } catch (e) {
          console.warn("[univer] getSheetSnapshot failed:", name, e);
        }
        return null;
      },
      restoreSheetSnapshot: (sheetSnapshot) => {
        const wb = univerAPIRef.current?.getActiveWorkbook?.();
        if (!wb || !sheetSnapshot) return false;
        try {
          // Prefer Univer's create/insert sheet APIs. Fall back to createSheet
          // + cell-by-cell restore if the snapshot API isn't available.
          const name = sheetSnapshot.name ?? "Sheet";
          if (typeof wb.createSheet === "function") {
            // Some builds accept (name, snapshot); others only (name).
            try {
              wb.createSheet(name, sheetSnapshot);
            } catch {
              wb.createSheet(name);
            }
          } else if (typeof wb.insertSheet === "function") {
            wb.insertSheet(sheetSnapshot);
          } else {
            // Last resort: create empty + write cells from cellData.
            const created = resolveSheet(univerAPIRef.current, name);
            if (!created) {
              // Try facade create via handle path — createSheet is defined above
              // but we're inside the same object literal; call through API.
              try {
                wb.addSheet?.(name) ?? wb.create?.(name);
              } catch {}
            }
          }
          // Always overlay cellData so content comes back even if createSheet
          // ignored the snapshot payload.
          const cellData = sheetSnapshot.cellData ?? {};
          const sheet = resolveSheet(univerAPIRef.current, name);
          if (sheet && cellData && typeof cellData === "object") {
            for (const rStr of Object.keys(cellData)) {
              const row = Number(rStr);
              const cols = cellData[rStr];
              if (!cols || typeof cols !== "object") continue;
              for (const cStr of Object.keys(cols)) {
                const col = Number(cStr);
                const cell = cols[cStr];
                const range = resolveRange(univerAPIRef.current, name, row, col);
                if (!range || !cell) continue;
                if (cell.f) range.setValue({ f: cell.f });
                else if (cell.v !== undefined) range.setValue(cell.v);
              }
            }
          }
          return true;
        } catch (e) {
          console.warn("[univer] restoreSheetSnapshot failed:", e);
          return false;
        }
      },
      getCellFormat: (sheetName, row, col): CellFormatShape => {
        const range = resolveRange(univerAPIRef.current, sheetName, row, col);
        return readCellFormat(range);
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
        syncSheetNamesById();
        void applyRenderActive();
        try {
          featureBaselineRef.current = perSheetFeatureState(
            api.getActiveWorkbook?.()?.getSnapshot?.()?.resources,
          );
        } catch {
          featureBaselineRef.current = null;
        }
      },
      setRenderActive: (active: boolean) => {
        if (renderActiveRef.current === active) return;
        renderActiveRef.current = active;
        void applyRenderActive();
      },
      jumpToCell: (sheetName, row, col) => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        const sheet = wb?.getSheetByName?.(sheetName);
        if (!sheet) return false;
        try { sheet.activate?.(); } catch {}
        try { sheet.getRange?.(row, col, 1, 1)?.activate?.(); } catch {}
        try {
          // Leave a little context above/left of the target cell.
          sheet.scrollToCell?.(Math.max(0, row - 3), Math.max(0, col - 2));
        } catch (e) {
          console.warn("[univer] scrollToCell failed:", e);
        }
        return true;
      },
      getActiveSheetName: () => {
        try {
          const name = univerAPIRef.current?.getActiveWorkbook?.()?.getActiveSheet?.()?.getName?.();
          if (typeof name === "string" && name) return name;
        } catch { /* fall through to selection-stream fallback */ }
        return liveSelectionRef.current?.sheet || null;
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
      setNote: (sheetName, row, col, text) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const range = sheet?.getRange?.(row, col, 1, 1);
          range?.createOrUpdateNote?.({ note: text, width: 160, height: 100, show: false });
        } catch (e) { console.warn("setNote failed:", e); }
      },
      deleteNote: (sheetName, row, col) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const range = sheet?.getRange?.(row, col, 1, 1);
          range?.deleteNote?.();
        } catch (e) { console.warn("deleteNote failed:", e); }
      },
      getNote: (sheetName, row, col) => {
        const sheet = resolveSheet(univerAPIRef.current, sheetName);
        try {
          const range = sheet?.getRange?.(row, col, 1, 1);
          const note = range?.getNote?.();
          return typeof note?.note === "string" ? note.note : null;
        } catch { return null; }
      },

      // ----- sheet ops -----

      createSheet: (name, tabColor): boolean => {
        const api = univerAPIRef.current;
        const wb = api?.getActiveWorkbook?.();
        if (!wb) return false;
        try {
          // Univer facade methods vary by minor version — try the common ones.
          // Dimensions MUST be passed explicitly: facade create() defaults to
          // 1000×20, and a write past column T on a fresh sheet then throws —
          // which is how an agent copy into A1:AO10 of a new sheet silently
          // produced an empty sheet. Match the 10000×200 we use on load.
          const created =
            wb.create?.(name, 10000, 200) ??
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
      };

      // Wrap every mutating method so the command stream can tell programmatic
      // edits (these) from raw user typing. Read-only methods are left alone.
      const MUTATORS: (keyof UniverGridHandle)[] = [
        "loadBytes", "loadSnapshot", "setCell", "setCells", "setCellBackground", "setCellFormat",
        "setRangeFormat", "setColumnWidth", "setRowHeight", "mergeCells", "unmergeCells",
        "setNote", "deleteNote",
        "createSheet", "deleteSheet", "renameSheet", "clearRange", "insertRows", "deleteRows",
        "insertColumns", "deleteColumns", "freezePanes", "unfreezePanes", "hideRows",
        "showRows", "hideColumns", "showColumns", "setCellBorders", "setRangeBorders",
        "sortRange", "defineName", "deleteName", "restoreSheetSnapshot",
      ];
      for (const key of MUTATORS) {
        const orig = handle[key] as (...args: any[]) => any;
        (handle[key] as any) = (...args: any[]) => {
          programmaticDepthRef.current++;
          let out: any;
          try {
            out = orig(...args);
          } catch (e) {
            programmaticDepthRef.current--;
            throw e;
          }
          // Keep the guard raised until async mutators (load*) fully settle,
          // otherwise their later commands would look like user edits.
          if (out && typeof out.then === "function") {
            return out.finally(() => { programmaticDepthRef.current--; });
          }
          programmaticDepthRef.current--;
          return out;
        };
      }

      return handle;
      // Every method above closes over refs (read via `.current` at call
      // time) or module-level helper functions — never `workbookPath` /
      // `onUserEdit` directly (that prop is mirrored into `onUserEditRef`
      // specifically so long-lived closures never go stale). So the handle
      // itself never needs to change across renders — empty deps makes
      // React build it ONCE instead of recreating every method closure on
      // every re-render.
    }, []);

    return (
      <div
        ref={containerRef}
        style={{ width: "100%", height: "100%", minHeight: 400, background: "#fff" }}
        data-workbook-path={workbookPath ?? ""}
      />
    );
  },
));

function resolveRange(api: any, sheetName: string, row: number, col: number): any {
  const wb = api?.getActiveWorkbook?.();
  if (!wb) return null;
  const sheet = wb.getSheetByName?.(sheetName) ?? wb.getActiveSheet?.();
  return sheet?.getRange?.(row, col) ?? null;
}

/**
 * Tokens that mark a `sheet.command.*` id as an actual content/format/structure
 * edit (as opposed to selection, scroll, zoom, copy, etc.). Used to recognize
 * raw user edits in the command stream. Substring match keeps it resilient to
 * Univer's id variants (insert-row / insert-row-before / insert-multi-rows-...).
 */
const EDIT_COMMAND_HINTS = [
  "set-range-values", "set-range-format", "set-bold", "set-italic", "set-underline",
  "set-stroke", "set-strike", "set-font", "set-background-color", "reset-background-color",
  "set-text-color", "reset-text-color", "set-number-format", "set-text-wrap",
  "set-horizontal-text-align", "set-vertical-text-align", "set-text-rotation",
  "set-border", "insert-row", "insert-col", "remove-row", "remove-col", "append-row",
  "move-rows", "move-cols", "insert-range-move", "delete-range-move", "worksheet-merge",
  "clear-selection", "delta-column-width", "delta-row-height", "set-col-data", "set-row-data",
  "set-worksheet-row-height", "set-worksheet-col-width", "set-col-width", "set-row-height",
  "set-col-hidden", "set-row-hidden", "paste",
  // Sort moves cell content; CF / data-validation / note edits change plugin
  // state the save path must persist (via the feature-drift fallback), so
  // all of them must mark the tab dirty. "toggle-note-popup" is deliberately
  // absent — hovering a note is not an edit.
  "sort-range", "conditional-rule", "DataValidation", "data-validation",
  "update-note", "delete-note",
];

function isEditingCommand(id: string): boolean {
  return EDIT_COMMAND_HINTS.some((h) => id.includes(h));
}

/** Resolve a rectangular FRange (inclusive bounds) for whole-range ops. */
function resolveRangeRect(
  api: any,
  sheetName: string,
  sr: number,
  sc: number,
  er: number,
  ec: number,
): any {
  const wb = api?.getActiveWorkbook?.();
  if (!wb) return null;
  const sheet = wb.getSheetByName?.(sheetName) ?? wb.getActiveSheet?.();
  return sheet?.getRange?.(sr, sc, er - sr + 1, ec - sc + 1) ?? null;
}

/**
 * Apply a partial CellFormat to an FRange (single cell or rectangle). Shared by
 * setCellFormat / setRangeFormat so both stay in sync. Each property is applied
 * only when present; the whole call is a single Univer command (one undo step).
 */
/**
 * Univer's facade has a quirky horizontal-alignment vocabulary:
 * `'left' | 'center' | 'normal'` where `'normal'` means RIGHT. Passing the
 * literal `'right'` throws ("Invalid horizontal alignment"), which is why the
 * align-right button silently did nothing. Translate our value here.
 */
function toFacadeHAlign(a: string | undefined): "left" | "center" | "normal" | undefined {
  if (a === "left" || a === "center") return a;
  if (a === "right") return "normal";
  return undefined;
}

function applyFormatToFRange(range: any, format: CellFormatShape): void {
  try {
    if (format.bold !== undefined) range.setFontWeight?.(format.bold ? "bold" : "normal");
    if (format.italic !== undefined) range.setFontStyle?.(format.italic ? "italic" : "normal");
    if (format.underline !== undefined) range.setFontLine?.(format.underline ? "underline" : "none");
    if (format.strike !== undefined) range.setFontLine?.(format.strike ? "line-through" : "none");
    if (format.font_color) range.setFontColor?.(format.font_color);
    if (format.font_size) range.setFontSize?.(format.font_size);
    if (format.font_family) range.setFontFamily?.(format.font_family);
    const ha = toFacadeHAlign(format.horizontal_align);
    if (ha) range.setHorizontalAlignment?.(ha);
    if (format.vertical_align) range.setVerticalAlignment?.(format.vertical_align);
    if (format.wrap_text !== undefined) range.setWrap?.(format.wrap_text);
    if (format.number_format) range.setNumberFormat?.(format.number_format);
    // background_color is applied via setBackgroundColor since Univer's facade
    // exposes background as a distinct API. Null/empty clears the fill.
    if (format.background_color !== undefined) {
      range.setBackgroundColor?.(format.background_color || null);
    }
  } catch (e) {
    console.warn("applyFormatToFRange failed:", e);
  }
}

/** Read managed format props from an FRange. Shared by getCellFormat + captureCellBand. */
function readCellFormat(range: any): CellFormatShape {
  const f: CellFormatShape = {};
  if (!range) return f;
  try {
    // Font/fill/number-format are read via getCellStyle() — the Univer
    // 0.23 FRange facade has NO getFontWeight/getFontColor/etc getters,
    // so the older reads silently returned undefined and broke toggles
    // (pressing Bold again couldn't tell the cell was already bold).
    const st = range.getCellStyle?.();
    if (st) {
      if (st.bold) f.bold = true;
      if (st.italic) f.italic = true;
      if (st.underline?.show) f.underline = true;
      if (st.strikethrough?.show) f.strike = true;
      if (st.color?.rgb) f.font_color = st.color.rgb;
      if (st.fontSize) f.font_size = st.fontSize;
      if (st.fontFamily) f.font_family = st.fontFamily;
      if (st.numberFormat?.pattern) f.number_format = st.numberFormat.pattern;
      // Univer sometimes reports a transparent default fill. Only surface
      // a real color so Reject can distinguish "user set white" from
      // "no fill explicitly set".
      const bg = st.background?.rgb;
      if (bg && bg !== "rgba(0,0,0,0)" && bg !== "transparent") f.background_color = bg;
    }
    // Alignment + wrap DO have dedicated facade getters. Univer reports
    // horizontal align as 'left' | 'center' | 'normal' (='right') |
    // 'general' (=unset) — normalize back to our vocabulary.
    const ha = range.getHorizontalAlignment?.();
    if (ha === "left" || ha === "center") f.horizontal_align = ha;
    else if (ha === "normal") f.horizontal_align = "right";
    const va = range.getVerticalAlignment?.();
    if (va) f.vertical_align = va;
    const wrap = range.getWrap?.();
    if (typeof wrap === "boolean") f.wrap_text = wrap;
  } catch {}
  return f;
}

function resolveSheet(api: any, sheetName: string): any {
  const wb = api?.getActiveWorkbook?.();
  if (!wb) return null;
  return wb.getSheetByName?.(sheetName) ?? wb.getActiveSheet?.();
}

/**
 * Grow a sheet's row/column capacity so (row, col) is writable. Univer sheets
 * have a fixed rowCount/columnCount (facade default: 1000×20) and writes past
 * the edge are dropped or throw — the failure mode behind "the agent copied
 * data onto my new sheet and nothing appeared". Grows in comfortable slabs so
 * repeated writes near the edge don't re-trigger the command every time.
 */
function ensureSheetCapacity(sheet: any, row: number, col: number): void {
  if (!sheet) return;
  try {
    const rows = sheet.getMaxRows?.();
    const cols = sheet.getMaxColumns?.();
    if (typeof rows === "number" && row >= rows) {
      sheet.setRowCount?.(Math.max(row + 100, rows));
    }
    if (typeof cols === "number" && col >= cols) {
      sheet.setColumnCount?.(Math.max(col + 20, cols));
    }
  } catch (e) {
    console.warn("[univer] ensureSheetCapacity failed:", e);
  }
}

/** [3,4,5,9] → [[3,3],[9,1]] as (start, count) runs of consecutive indices. */
function contiguousRuns(sorted: number[]): Array<[start: number, count: number]> {
  const runs: Array<[number, number]> = [];
  for (const i of sorted) {
    const last = runs[runs.length - 1];
    if (last && i === last[0] + last[1]) last[1] += 1;
    else runs.push([i, 1]);
  }
  return runs;
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

export type BorderSide = { style: string; color: string | null } | null;
export type BordersShape = {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
};

/**
 * Render the workbook's anchored pictures via the drawing preset. Read-only
 * fidelity: images come from the parse result (extracted from the ExcelJS
 * media store, possibly inside the import worker) and are placed at their
 * anchor's top-left cell. The file's own drawing parts are never rewritten
 * (surgical save copies them raw), so this cannot corrupt anything — worst
 * case an image simply doesn't render.
 */
async function insertWorksheetImages(api: any, images: SheetImage[]): Promise<number> {
  let inserted = 0;
  if (!images.length) return 0;
  const wb = api?.getActiveWorkbook?.();
  if (!wb) return 0;
  for (const img of images) {
    const sheet = wb.getSheetByName?.(img.sheetName);
    if (!sheet) continue;
    if (!sheet.insertImage) {
      console.warn("[univer] drawing facade unavailable; images not rendered");
      return inserted;
    }
    try {
      const blob = new Blob([img.buffer], { type: `image/${img.extension}` });
      const url = URL.createObjectURL(blob);
      await sheet.insertImage(url, img.col, img.row, 0, 0);
      inserted++;
    } catch (e) {
      console.warn("[univer] image insert failed on", img.sheetName, e);
    }
  }
  return inserted;
}

/**
 * Apply per-side borders to a Univer range. Univer's facade exposes border
 * APIs that vary slightly between minor versions — we try the most common
 * shapes in order. The worst case is "border doesn't render in our app" —
 * the saved file still has it because ExcelJS preserves the cell border
 * object untouched on the round-trip.
 */
/**
 * Map our string border-style names to Univer's numeric `BorderStyleTypes`
 * enum. We hardcode the values (rather than importing the enum) because the
 * facade is consumed as `any` and the numeric values are stable across builds:
 * NONE=0, THIN=1, HAIR=2, DOTTED=3, DASHED=4, DOUBLE=7, MEDIUM=8, THICK=13.
 */
function borderStyleToUniver(style?: string): number {
  switch ((style ?? "thin").toLowerCase()) {
    case "none": return 0;
    case "hair": return 2;
    case "dotted": return 3;
    case "dashed": return 4;
    case "double": return 7;
    case "medium": return 8;
    case "thick": return 13;
    default: return 1; // thin
  }
}

/**
 * Apply per-side borders to a range via the Univer facade.
 *
 * The real facade signature is `setBorder(type: BorderType, style:
 * BorderStyleTypes, color?: string)` where `type` is a string enum
 * ("all" | "top" | "bottom" | "left" | "right" | "outside" | "none" | ...)
 * and `style` is the numeric enum above. An earlier version called this with
 * eight boolean args, which threw and silently rendered nothing.
 *
 * Used on load (per-cell borders extracted from the xlsx) and as the building
 * block for clearing. Range-level "all/outer/none" toolbar clicks go through
 * `setRangeBorders` instead so they collapse to a single undo step.
 */
function applyBordersToRange(range: any, borders: BordersShape): void {
  if (typeof range?.setBorder !== "function") return;
  const sides: Array<["top" | "bottom" | "left" | "right", BorderSide | null | undefined]> = [
    ["top", borders.top],
    ["bottom", borders.bottom],
    ["left", borders.left],
    ["right", borders.right],
  ];
  try {
    // Full box → single ALL call so Univer draws shared edges cleanly.
    if (borders.top && borders.bottom && borders.left && borders.right) {
      range.setBorder("all", borderStyleToUniver(borders.top.style), borders.top.color ?? "#000000");
      return;
    }
    for (const [pos, side] of sides) {
      if (side === undefined) continue; // leave this side untouched
      if (side === null) {
        // Explicit clear of a single side.
        range.setBorder(pos, 0);
      } else {
        range.setBorder(pos, borderStyleToUniver(side.style), side.color ?? "#000000");
      }
    }
  } catch (e) {
    // Don't let one bad border block the whole load. The file still has its
    // borders preserved through ExcelJS regardless of in-app rendering.
    console.warn("[univer] applyBordersToRange failed:", e);
  }
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
  featureBaseline?: Map<string, SheetFeatureState> | null,
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
    if (mirror) applyDefinedNamesMirror(fresh, mirror);
    // Null baseline = write ALL live CF/validation/note state — a fresh
    // workbook has no original parts to protect.
    const freshDropped = applyFeatureMirror(fresh, data, null).dropped;
    if (freshDropped.length) console.info("[save] feature mirror dropped:", freshDropped);
    const buf = await fresh.xlsx.writeBuffer();
    return new Uint8Array(buf as ArrayBuffer);
  }

  // Apply STRUCTURE ops first (sheet create/delete/rename, row/col splices)
  // so that when we then write Univer's cells to ExcelJS, the worksheets
  // exist with the correct names and the existing-row layout is already
  // shifted. Otherwise ExcelJS still has old-position cells lingering when
  // we write to new positions, producing duplicate content.
  if (mirror) applyStructureMirror(excelJsWorkbook, mirror);

  const externalPins = externalPinsByWorkbook.get(excelJsWorkbook);

  const order: string[] = data?.sheetOrder ?? Object.keys(data?.sheets ?? {});
  for (const sheetId of order) {
    const sheet = data.sheets[sheetId];
    if (!sheet) continue;
    let ws = excelJsWorkbook.getWorksheet(sheet.name);
    if (!ws) {
      ws = excelJsWorkbook.addWorksheet(sheet.name);
    }
    // Sync visibility from Univer (the user can hide/unhide from the tab
    // bar mid-session). Hidden keeps the file's existing hidden flavor —
    // a veryHidden sheet the user never touched stays veryHidden.
    if (sheet.hidden === 1) {
      if (ws.state === "visible") ws.state = "hidden";
    } else if (ws.state !== "visible") {
      ws.state = "visible";
    }
    writeUniverSheetIntoExcelJs(sheet, ws, externalPins);
  }

  // Style ops (cell formats / widths / heights / merges) come AFTER cells
  // since they reference final positions and depend on the cell existing.
  if (mirror) applyStyleMirror(excelJsWorkbook, mirror);

  // Defined names are workbook-level — apply after structure ops so any
  // renamed/created sheets they reference already exist.
  if (mirror) applyDefinedNamesMirror(excelJsWorkbook, mirror);

  // CF / data-validation / notes: replace-on-drift only. Sheets whose state
  // still matches the load-time baseline keep ExcelJS's own parsed models
  // (which may hold rule types our converters can't express); drifted
  // sheets are rewritten from the live Univer state so in-session edits
  // actually reach the file.
  const featDropped = applyFeatureMirror(excelJsWorkbook, data, featureBaseline ?? null).dropped;
  if (featDropped.length) console.info("[save] feature mirror dropped:", featDropped);

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
 * Workbook-scoped defined names (named ranges) the agent created. ExcelJS
 * accumulates ranges under a name across `add` calls, so we first strip any
 * existing ranges for the name (replace semantics) and then add the new ref.
 */
function applyDefinedNamesMirror(wb: ExcelJS.Workbook, mirror: SaveMirror): void {
  for (const dn of mirror.definedNames ?? []) {
    if (!dn?.name || !dn?.ref) continue;
    try {
      const existing = wb.definedNames.getRanges(dn.name);
      for (const r of existing?.ranges ?? []) {
        try { wb.definedNames.remove(r, dn.name); } catch {}
      }
    } catch {}
    try {
      wb.definedNames.add(dn.ref, dn.name);
    } catch (e) {
      console.warn("[save] defined name failed:", dn, e);
    }
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

  // --- cell borders --- (manual format toolbar). Merge into the cell's
  // existing border object so touching one side doesn't wipe the others.
  // A `null` side clears it; an `undefined` side is left untouched.
  for (const op of mirror.cellBorders ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    const cell = ws.getCell(op.row + 1, op.col + 1);
    const next: any = { ...((cell.border as any) ?? {}) };
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const s = (op.borders as any)[side];
      if (s === undefined) continue;
      if (s === null) {
        delete next[side];
      } else {
        next[side] = {
          style: excelBorderStyle(s.style),
          color: { argb: cssToArgb(s.color ?? "#000000") },
        };
      }
    }
    cell.border = next;
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

  // --- autofilter ---
  for (const op of mirror.autoFilters ?? []) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    try {
      // ExcelJS expects e.g. "A1:C10"; assigning undefined clears it.
      (ws as any).autoFilter = op.range ?? undefined;
    } catch (e) {
      console.warn("[save] autofilter failed:", op, e);
    }
  }
}

/**
 * Map our border style string to an ExcelJS `BorderStyle`. Our toolbar only
 * emits "thin" today, but the agent/load paths use the wider vocabulary, so we
 * translate the full set and default unknowns to "thin".
 */
function excelBorderStyle(style: string): string {
  switch (style) {
    case "hair": return "hair";
    case "dotted": return "dotted";
    case "dashed": return "dashed";
    case "dashDot": return "dashDot";
    case "medium": return "medium";
    case "thick": return "thick";
    case "double": return "double";
    default: return "thin";
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

function writeUniverSheetIntoExcelJs(
  univerSheet: any,
  ws: ExcelJS.Worksheet,
  externalPins?: ExternalPinMap,
): void {
  const cellData = univerSheet.cellData ?? {};
  for (const rowKey of Object.keys(cellData)) {
    const r = Number(rowKey);
    const row = cellData[rowKey];
    for (const colKey of Object.keys(row)) {
      const c = Number(colKey);
      const cell = row[colKey];
      // A pinned external-workbook cell whose value is untouched must NOT be
      // written: the original ExcelJS cell still holds { formula, result }
      // and skipping the write is what preserves the formula in the file.
      if (externalPins && !cell.f) {
        const pin = externalPins.get(externalPinKey(ws.name, r, c));
        if (pin && cell.v === pin.v) continue;
      }
      const eCell = ws.getCell(r + 1, c + 1);
      if (cell.f) {
        // Setting .value to { formula } preserves the cell's existing
        // style object (font/fill/numFmt) since we don't reassign it.
        // addXlfnPrefixes restores the `_xlfn.` markers Excel requires for
        // post-2007 functions (stripped at import for Univer's engine).
        const fv: any = { formula: addXlfnPrefixes(String(cell.f).replace(/^=/, "")) };
        // Cache the evaluated result so reopening in-app renders without a
        // full recalculation (Excel recalcs on open regardless, via
        // fullCalcOnLoad). Best-effort: unencodable results stay uncached.
        let v = cell.v;
        if (cell.t === 3 && typeof v === "number") v = v !== 0; // Univer bools are 0/1
        if ((typeof v === "number" && Number.isFinite(v)) || typeof v === "boolean") {
          fv.result = v;
        } else if (typeof v === "string" && v !== "") {
          fv.result = ERROR_VALUES.has(v) ? { error: v } : v;
        }
        eCell.value = fv;
      } else if (cell.v !== undefined && cell.v !== null) {
        eCell.value = cell.v;
      } else {
        eCell.value = null;
      }
    }
  }
}
