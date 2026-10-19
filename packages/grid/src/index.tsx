/**
 * gridpath: a spreadsheet for web apps. Opens real .xlsx files (or
 * starts blank), calculates with GridPath's Rust engine compiled to
 * WebAssembly in a Web Worker, draws only the visible window on a canvas,
 * and saves by patching the original file so every part it did not touch
 * (charts, pivots, VBA, add-in data) comes back byte-identical.
 *
 *   const sheet = await createSpreadsheet(el, { file: await f.arrayBuffer() });
 *   const xlsx = await sheet.save();
 */
import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { ViewportGrid } from "../../../src/screens/SpreadsheetScreen/components/ViewportGrid";
import type { ManualSheetOp, UniverGridHandle } from "../../../src/screens/SpreadsheetScreen/components/UniverGrid";
import { appendManualSheetOp, type SheetOp } from "../../../src/screens/SpreadsheetScreen/manualSheetOps";
import { buildSaveMirror } from "../../gridpath/src/core/saveMirror";
import { BrowserEngine, type EngineOptions } from "./browserEngine";
import { AgentToolkit, type SheetAgent } from "./ai/toolkit";

export { BrowserEngine, type EngineOptions } from "./browserEngine";
export type { GridEngine, DeltaApplier } from "./engine";
export { INSTRUCTIONS as AGENT_INSTRUCTIONS, type SheetAgent, type AgentBatch, type ToolDefinition, type OpenAIToolDefinition, type ReadCell } from "./ai/toolkit";

export interface SpreadsheetOptions {
  /** An .xlsx to open. Without it the workbook starts blank. */
  file?: ArrayBuffer | Uint8Array;
  /** Sheet names for a blank workbook (default ["Sheet1"]). */
  sheets?: string[];
  /** Name shown for the workbook in messages; default "workbook.xlsx". */
  fileName?: string;
  /** Called after each user edit in the grid. */
  onEdit?: () => void;
  /** Loading stages ("Loading workbook…", "Building grid…"). */
  onProgress?: (stage: string) => void;
  /** Where and how the engine runs. */
  engine?: EngineOptions;
}

export interface CellValue {
  /** The raw value: a number, string, boolean, or null when empty. */
  value: unknown;
  /** The formatted text shown in the cell, when it differs from `value`. */
  display?: string;
  /** The formula without "=", or null for a constant. */
  formula: string | null;
}

export interface SaveOptions {
  /** Save agent changes the user has not accepted yet. Default false:
   * `save()` throws a `PendingChangesError` while any are pending. */
  includePending?: boolean;
}

/** `save()` was called while agent changes were pending review. */
export class PendingChangesError extends Error {
  /** How many batches are pending. */
  readonly pending: number;
  constructor(pending: number) {
    super(`${pending} agent change${pending === 1 ? " is" : "s are"} pending review: accept or reject ${pending === 1 ? "it" : "them"} before saving, or save with { includePending: true }`);
    this.name = "PendingChangesError";
    this.pending = pending;
  }
}

export interface Spreadsheet {
  /** Sheet names in workbook order. */
  getSheetNames(): string[];
  /** One cell by A1 reference ("B7"). */
  getCell(sheet: string, ref: string): Promise<CellValue | null>;
  /** Write one cell by A1 reference. Strings starting with "=" are formulas. */
  setCell(sheet: string, ref: string, value: string | number | boolean | null): void;
  /** Write many cells on one sheet in one step (0-based row and col). */
  setCells(sheet: string, cells: Array<{ row: number; col: number; value: string | number | boolean | null }>): void;
  /** Resolves when pending writes are recalculated. */
  whenCalculated(timeoutMs?: number): Promise<void>;
  /** Scroll to and select a cell. */
  select(sheet: string, ref: string): void;
  /** Give the grid keyboard focus (typing then edits the selected cell). */
  focus(): void;
  /** The workbook as .xlsx bytes: the original file with this session's
   * changes patched in (for a blank workbook, a new file). While agent
   * changes are pending review it throws a `PendingChangesError`: accept or
   * reject them first (`ai.accept()`, `ai.reject()`), or pass
   * `{ includePending: true }` to save them unreviewed. */
  save(options?: SaveOptions): Promise<Uint8Array>;
  /** `save()` as a Blob with the .xlsx media type, ready to download or upload. */
  saveBlob(options?: SaveOptions): Promise<Blob>;
  /** Unmount the grid and stop the engine. */
  destroy(): void;
  /** GridPath's agent tools over this workbook: give `ai.tools()` to your
   * model, run its tool calls with `ai.run()`, and let the user `accept` or
   * `reject` the highlighted changes. Your model, your key. */
  readonly ai: SheetAgent;
  /** The grid's internal handle (unstable; for GridPath's own tooling). */
  readonly internal: unknown;
}

/** "B7" → { row: 6, col: 1 } (0-based). */
function parseRef(ref: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  if (!m) throw new Error(`not an A1 cell reference: ${ref}`);
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

let instances = 0;

export interface SpreadsheetProps extends SpreadsheetOptions {
  /** Called once the workbook is open and drawn. */
  onReady?: (sheet: Spreadsheet) => void;
  /** Called when the engine or the file fails to load. */
  onError?: (error: Error) => void;
  className?: string;
  style?: React.CSSProperties;
}

/**
 * The React component. It fills its parent; give the parent a height.
 * The workbook loads once: to open a different file, remount it with a
 * new `key`.
 */
export function Spreadsheet(props: SpreadsheetProps): React.ReactElement {
  const tabId = useMemo(() => `gridpath-${++instances}`, []);
  const [engine, setEngine] = useState<BrowserEngine | null>(null);
  const gridRef = useRef<UniverGridHandle | null>(null);
  const hostRef = useRef<HTMLDivElement | null>(null);
  const sheetOps = useRef<SheetOp[]>([]);
  const propsRef = useRef(props);
  propsRef.current = props;
  const engineOpts = props.engine;

  useEffect(() => {
    let live = true;
    let created: BrowserEngine | null = null;
    BrowserEngine.create(engineOpts).then(
      (e) => {
        created = e;
        if (live) setEngine(e);
        else e.close();
      },
      (err) => propsRef.current.onError?.(err instanceof Error ? err : new Error(String(err))),
    );
    return () => {
      live = false;
      created?.close();
    };
    // The engine is created once per mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    const grid = gridRef.current;
    if (!engine || !grid) return;
    let live = true;
    (async () => {
      const p = propsRef.current;
      const base = p.file ? new Uint8Array(p.file instanceof Uint8Array ? p.file : new Uint8Array(p.file)) : await engine.blankXlsx(p.sheets ?? ["Sheet1"]);
      const res = await grid.loadBytes(base, (s) => propsRef.current.onProgress?.(s), { tabId, path: p.fileName ?? "workbook.xlsx" });
      if (!live) return;
      if (!res.engineLoaded) throw new Error(engine.degradedReason() ?? "the workbook could not be opened");
      const ai = new AgentToolkit(grid, engine);
      const api: Spreadsheet = {
        getSheetNames: () => grid.getSheetNames(),
        getCell: async (sheet, ref) => {
          const { row, col } = parseRef(ref);
          await grid.whenCalculated(30_000);
          await grid.prefetchRange?.(sheet, row, col, row, col);
          const c = grid.getCell(sheet, row, col);
          return c ? { value: c.value, ...(c.display !== undefined ? { display: c.display } : {}), formula: c.formula } : null;
        },
        setCell: (sheet, ref, value) => {
          const { row, col } = parseRef(ref);
          grid.setCells(sheet, [{ row, col, value }]);
        },
        setCells: (sheet, cells) => void grid.setCells(sheet, cells),
        whenCalculated: (timeoutMs) => grid.whenCalculated(timeoutMs),
        select: (sheet, ref) => {
          const { row, col } = parseRef(ref);
          grid.jumpToCell(sheet, row, col);
        },
        focus: () => hostRef.current?.querySelector<HTMLElement>("[tabindex='0']")?.focus(),
        save: async (options) => {
          // Review comes first: unaccepted agent edits are not saved by accident.
          const pending = ai.list().filter((b) => b.status === "pending").length;
          if (pending && !options?.includePending) throw new PendingChangesError(pending);
          // Formula results are cached into the file: let recalculation
          // finish first, or the file would reopen with stale numbers.
          await grid.whenCalculated(30_000);
          await grid.syncEditLog?.();
          // The agent's applied batches carry their structure and format
          // ops (as on the desktop); the user's own sheet-tab ops follow.
          const mirror = buildSaveMirror(ai.appliedBatches());
          if (sheetOps.current.length) mirror.sheetOps = [...(mirror.sheetOps ?? []), ...sheetOps.current];
          // Current results of every formula this session recalculated (new
          // ones included), cached into the file for readers that don't calculate.
          mirror.cacheRefresh = await engine.recalculatedCells();
          const built = grid.exportPatch(mirror as any);
          if (!built.ok) throw new Error(`cannot save: ${built.reason}${(built as any).detail ? ` (${(built as any).detail})` : ""}`);
          // Every save replays the whole session onto the ORIGINAL bytes.
          return engine.patch(base, JSON.stringify(built.patch));
        },
        saveBlob: async (options) => new Blob([(await api.save(options)) as BlobPart], { type: XLSX_TYPE }),
        destroy: () => {
          throw new Error("destroy() is only available from createSpreadsheet(); unmount the component instead");
        },
        ai,
        internal: grid,
      };
      propsRef.current.onReady?.(api);
    })().catch((err) => {
      if (live) propsRef.current.onError?.(err instanceof Error ? err : new Error(String(err)));
    });
    return () => {
      live = false;
    };
  }, [engine, tabId]);

  return (
    <div ref={hostRef} className={props.className} style={{ position: "relative", width: "100%", height: "100%", overflow: "hidden", ...props.style }}>
      {engine && (
        <ViewportGrid
          ref={gridRef}
          tabId={tabId}
          workbookPath={props.fileName ?? "workbook.xlsx"}
          engine={engine}
          onUserEdit={() => propsRef.current.onEdit?.()}
          onManualSheetOp={(op: ManualSheetOp) => appendManualSheetOp(sheetOps.current, op)}
        />
      )}
    </div>
  );
}

/**
 * Mount a spreadsheet into `el` (which should have a height) and resolve
 * once the workbook is open. Works without React in your app; React and
 * React DOM are peer dependencies of the package.
 */
export function createSpreadsheet(el: HTMLElement, options: SpreadsheetOptions = {}): Promise<Spreadsheet> {
  const root = createRoot(el);
  return new Promise((resolve, reject) => {
    const destroy = () => root.unmount();
    root.render(
      <Spreadsheet
        {...options}
        onReady={(api) => resolve({ ...api, destroy })}
        onError={(err) => {
          destroy();
          reject(err);
        }}
      />,
    );
  });
}
