/**
 * gridpath: a spreadsheet for web apps. Opens real .xlsx files (or
 * starts blank), calculates with GridPath's Rust engine compiled to
 * WebAssembly in a Web Worker, draws only the visible window on a canvas,
 * and saves by patching the original file so every part it did not touch
 * (charts, pivots, VBA, add-in data) comes back byte-identical.
 *
 * Out of the box it has a formatting toolbar, saves as you work (in this
 * browser, or to your backend with `store` / `onSave`), reopens where you
 * left off, keeps a version history, and downloads .xlsx.
 *
 *   const sheet = await createSpreadsheet(el, { file: await f.arrayBuffer() });
 *   const xlsx = await sheet.save();
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { ViewportGrid } from "../../../src/screens/SpreadsheetScreen/components/ViewportGrid";
import type { ManualSheetOp, UniverGridHandle } from "../../../src/screens/SpreadsheetScreen/components/UniverGrid";
import { appendManualSheetOp, type SheetOp } from "../../../src/screens/SpreadsheetScreen/manualSheetOps";
import { buildSaveMirror } from "../../gridpath/src/core/saveMirror";
import { BrowserEngine, type EngineOptions } from "./browserEngine";
import { AgentToolkit, type AgentBatch, type SheetAgent } from "./ai/toolkit";
import { Chrome, HistoryPanel, ReviewPanel, type SaveStatus } from "./ui/Chrome";
import { toolbarActions } from "./ui/actions";
import { browserStore, fingerprint, type WorkbookStore, type WorkbookVersion } from "./ui/store";

export { BrowserEngine, type EngineOptions } from "./browserEngine";
export type { GridEngine, DeltaApplier } from "./engine";
export { INSTRUCTIONS as AGENT_INSTRUCTIONS, type SheetAgent, type AgentBatch, type ToolDefinition, type OpenAIToolDefinition, type ReadCell } from "./ai/toolkit";
export { browserStore, type BrowserStoreOptions, type WorkbookStore, type StoredWorkbook, type WorkbookVersion } from "./ui/store";
export type { SaveStatus } from "./ui/Chrome";

export interface SpreadsheetOptions {
  /** An .xlsx to open. Without it the workbook starts blank (or reopens the
   * stored one, see `documentId`). */
  file?: ArrayBuffer | Uint8Array;
  /** Sheet names for a blank workbook (default ["Sheet1"]). */
  sheets?: string[];
  /** The workbook's file name (downloads, the recent list); default "Untitled.xlsx". */
  fileName?: string;
  /**
   * Identifies the workbook in the store, so it reopens where the user left
   * off: the stored state is opened instead of `file` when it was saved from
   * the same file (or when no `file` is given). Default: `fileName`, else
   * "Untitled.xlsx".
   */
  documentId?: string;
  /** The toolbar: formatting, File menu, save status, Download. Default true. */
  toolbar?: boolean;
  /** The File menu's Open, New and Recent entries. Default true. */
  fileMenu?: boolean;
  /** The "AI changes" panel: the agent's pending changes, each with Accept
   * and Reject; it opens when changes arrive for review. Default true; turn
   * it off when your app has its own review UI (`sheet.ai.list()`). */
  reviewPanel?: boolean;
  /**
   * Where saves, recent workbooks and version history are kept. Default:
   * this browser (`browserStore()`, IndexedDB). Pass your own to save to
   * your backend, or `false` for none (then `save()` / Download only).
   */
  store?: WorkbookStore | false;
  /** Save automatically once edits pause: true (2 s), a delay in ms, or
   * false. Needs a store or `onSave`. Default true. */
  autosave?: boolean | number;
  /** Called with the workbook's bytes on every save (yours to upload). */
  onSave?: (workbook: { id: string; name: string; bytes: Uint8Array }) => void | Promise<void>;
  /** The save status changed ("Saving…", "Saved", "Unsaved changes"…). */
  onStatusChange?: (status: SaveStatus) => void;
  /** Another workbook was opened (File menu, recent list, a restored version, `open()`). */
  onDocumentChange?: (doc: { id: string; name: string }) => void;
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
  /** Undo / redo the last change (the user's or an agent's). */
  undo(): void;
  redo(): void;
  /** The workbook as .xlsx bytes: the original file with this session's
   * changes patched in (for a blank workbook, a new file). While agent
   * changes are pending review it throws a `PendingChangesError`: accept or
   * reject them first (`ai.accept()`, `ai.reject()`), or pass
   * `{ includePending: true }` to save them unreviewed. */
  save(options?: SaveOptions): Promise<Uint8Array>;
  /** `save()` as a Blob with the .xlsx media type, ready to download or upload. */
  saveBlob(options?: SaveOptions): Promise<Blob>;
  /** Save to the store and `onSave` now (what autosave does); resolves with the new status. */
  saveNow(): Promise<SaveStatus>;
  /** Download the workbook as an .xlsx file. */
  download(fileName?: string): Promise<void>;
  /** Where the workbook stands relative to its last save. */
  getStatus(): SaveStatus;
  /** The open workbook. */
  getDocument(): { id: string; name: string };
  /** Open another workbook in this spreadsheet (saving the current one first). */
  open(file: ArrayBuffer | Uint8Array, options?: { name?: string; id?: string }): Promise<void>;
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
const AUTOSAVE_MS = 2000;

let instances = 0;

/** The workbook on screen. `key` remounts the grid when it changes. */
type Doc = { key: number; id: string; name: string; bytes: Uint8Array | null; source: string; savedAt: string | null };

/** What belongs to the open workbook (replaced when another opens). */
type Current = {
  grid: UniverGridHandle | null;
  ai: AgentToolkit | null;
  base: Uint8Array | null;
  sheetOps: SheetOp[];
  /** Freeze panes per sheet as the workbook opened. */
  initialFreeze: Map<string, { rows: number; cols: number }>;
};

function download(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xlsm") ? name : `${name}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function readView(id: string): { sheet: string; row: number; col: number } | null {
  try {
    const raw = localStorage.getItem(`gridpath:view:${id}`);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function writeView(id: string, grid: UniverGridHandle | null) {
  try {
    const sel = grid?.getActiveSelection?.();
    const sheet = sel?.sheet ?? grid?.getActiveSheetName?.();
    if (sheet) localStorage.setItem(`gridpath:view:${id}`, JSON.stringify({ sheet, row: sel?.startRow ?? 0, col: sel?.startCol ?? 0 }));
  } catch {
    // storage unavailable: nothing to remember
  }
}

/** `sheet.ai` stays the same object across workbooks; listeners carry over. */
function forwardingAgent(current: () => AgentToolkit | null, listeners: Set<(b: AgentBatch[]) => void>): SheetAgent {
  const need = () => {
    const ai = current();
    if (!ai) throw new Error("the workbook is still opening");
    return ai;
  };
  return {
    tools: ((format?: "anthropic" | "openai") => need().tools(format as any)) as SheetAgent["tools"],
    instructions: () => need().instructions(),
    context: () => need().context(),
    run: (name, input) => need().run(name, input),
    list: () => current()?.list() ?? [],
    onChange: (fn) => {
      listeners.add(fn);
      return () => listeners.delete(fn);
    },
    accept: (id) => need().accept(id),
    reject: (id) => need().reject(id),
  };
}

export interface SpreadsheetProps extends SpreadsheetOptions {
  /** Called once the first workbook is open and drawn. */
  onReady?: (sheet: Spreadsheet) => void;
  /** Called when the engine or the file fails to load. */
  onError?: (error: Error) => void;
  className?: string;
  style?: React.CSSProperties;
}

/**
 * The React component. It fills its parent; give the parent a height.
 * `file` is read when it mounts; to open another file, call `open()` (or
 * use the File menu), or remount with a new `key`.
 */
export function Spreadsheet(props: SpreadsheetProps): React.ReactElement {
  const instance = useMemo(() => ++instances, []);
  const [engine, setEngine] = useState<BrowserEngine | null>(null);
  const [doc, setDoc] = useState<Doc | null>(null);
  const [status, setStatusState] = useState<SaveStatus>({ state: "idle" });
  const [history, setHistory] = useState({ canUndo: false, canRedo: false });
  const [notice, setNotice] = useState<string | null>(null);
  const [versions, setVersions] = useState<WorkbookVersion[] | null | false>(false);
  const [batches, setBatches] = useState<AgentBatch[]>([]);
  const [reviewOpen, setReviewOpen] = useState(false);
  const pendingBefore = useRef(0);
  // Server rendering (Next.js) draws an empty box: the toolbar depends on
  // browser storage, the grid on the engine, both only in the browser.
  const [mounted, setMounted] = useState(false);
  useEffect(() => setMounted(true), []);
  const gridRef = useRef<UniverGridHandle | null>(null);
  const hostRef = useRef<HTMLDivElement | null>(null);
  const propsRef = useRef(props);
  propsRef.current = props;
  const docRef = useRef<Doc | null>(null);
  docRef.current = doc;
  const cur = useRef<Current>({ grid: null, ai: null, base: null, sheetOps: [], initialFreeze: new Map() });
  const statusRef = useRef<SaveStatus>(status);
  const listeners = useRef(new Set<(b: AgentBatch[]) => void>());
  const apiRef = useRef<Spreadsheet | null>(null);
  const readyCalled = useRef(false);
  // Change bookkeeping: `changes` counts edits, `saved` is its value at the
  // last good save; they differ while there are unsaved changes.
  const changes = useRef(0);
  const saved = useRef(0);
  const saving = useRef<Promise<SaveStatus> | null>(null);
  const lastSaveMs = useRef(0);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const noticeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  // `store` and friends are read once: the default store is this browser's.
  const store = useMemo<WorkbookStore | null>(() => {
    if (props.store === false) return null;
    if (props.store) return props.store;
    return typeof indexedDB === "undefined" ? null : browserStore();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const canSave = !!store || !!props.onSave;
  const autosaveMs = props.autosave === false ? 0 : typeof props.autosave === "number" ? Math.max(250, props.autosave) : AUTOSAVE_MS;

  const setStatus = useCallback((s: SaveStatus) => {
    statusRef.current = s;
    setStatusState(s);
    propsRef.current.onStatusChange?.(s);
  }, []);

  const notify = useCallback((message: string) => {
    setNotice(message);
    if (noticeTimer.current) clearTimeout(noticeTimer.current);
    noticeTimer.current = setTimeout(() => setNotice(null), 2500);
  }, []);

  const pendingCount = () => cur.current.ai?.list().filter((b) => b.status === "pending").length ?? 0;
  const dirty = () => changes.current !== saved.current;

  // --- engine ----------------------------------------------------------------
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

  // --- the first workbook: the host's file, or the stored one ----------------
  useEffect(() => {
    let live = true;
    (async () => {
      const p = propsRef.current;
      const id = p.documentId ?? p.fileName ?? "Untitled.xlsx";
      const file = p.file ? new Uint8Array(p.file instanceof Uint8Array ? p.file : new Uint8Array(p.file)) : null;
      const source = file ? fingerprint(file) : "blank";
      const stored = store ? await store.load(id).catch(() => null) : null;
      if (!live) return;
      // Reopen where the user left off, unless the host passed a different file.
      if (stored && (!file || stored.source === source)) {
        setDoc({ key: 1, id, name: stored.name, bytes: stored.bytes, source: stored.source ?? source, savedAt: stored.savedAt });
      } else {
        setDoc({ key: 1, id, name: p.fileName ?? "Untitled.xlsx", bytes: file, source, savedAt: null });
      }
    })().catch((err) => propsRef.current.onError?.(err instanceof Error ? err : new Error(String(err))));
    return () => {
      live = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // --- saving ----------------------------------------------------------------

  /** The workbook's bytes now (the patch of this session onto the file it opened from). */
  const buildBytes = useCallback(async (options?: SaveOptions): Promise<Uint8Array> => {
    const { grid, ai, base } = cur.current;
    if (!grid || !ai || !base || !engine) throw new Error("the workbook is still opening");
    // Review comes first: unaccepted agent edits are not saved by accident.
    const pending = pendingCount();
    if (pending && !options?.includePending) throw new PendingChangesError(pending);
    // Formula results are cached into the file: let recalculation finish.
    await grid.whenCalculated(30_000);
    await grid.syncEditLog?.();
    // The agent's applied batches carry their structure and format ops (as
    // on the desktop); the user's own sheet-tab ops follow.
    const mirror = buildSaveMirror(ai.appliedBatches());
    if (cur.current.sheetOps.length) mirror.sheetOps = [...(mirror.sheetOps ?? []), ...cur.current.sheetOps];
    // Freeze panes as they are now, for every sheet where they changed (the
    // agent's, the toolbar's, undone or not).
    const renamed = new Map<string, string>(); // current name → name at open
    for (const op of mirror.sheetOps ?? []) {
      if (op.kind === "rename") renamed.set(op.newName, renamed.get(op.oldName) ?? op.oldName);
    }
    mirror.freezePanes = [];
    for (const name of grid.getSheetNames()) {
      const now = grid.getFreezePanes(name);
      const was = cur.current.initialFreeze.get(renamed.get(name) ?? name) ?? { rows: 0, cols: 0 };
      if (now.freezeRows !== was.rows || now.freezeCols !== was.cols) mirror.freezePanes.push({ sheet: name, freezeRows: now.freezeRows, freezeCols: now.freezeCols });
    }
    // Current results of every formula this session recalculated (new ones
    // included), cached into the file for readers that don't calculate.
    mirror.cacheRefresh = await engine.recalculatedCells();
    const built = grid.exportPatch(mirror as any);
    if (!built.ok) throw new Error(`cannot save: ${built.reason}${(built as any).detail ? ` (${(built as any).detail})` : ""}`);
    // Every save replays the whole session onto the ORIGINAL bytes.
    return engine.patch(base, JSON.stringify(built.patch));
  }, [engine]);

  const saveNow = useCallback(async (opts?: { checkpoint?: boolean }): Promise<SaveStatus> => {
    const d = docRef.current;
    const p = propsRef.current;
    if (!d || (!store && !p.onSave)) return statusRef.current;
    if (timer.current) {
      clearTimeout(timer.current);
      timer.current = null;
    }
    const pending = pendingCount();
    if (pending) {
      setStatus({ state: "review", pending });
      return statusRef.current;
    }
    if (saving.current) return saving.current;
    const run = (async (): Promise<SaveStatus> => {
      const at = changes.current;
      const t0 = performance.now();
      setStatus({ state: "saving" });
      try {
        const bytes = await buildBytes();
        if (store) await store.save({ id: d.id, name: d.name, bytes, source: d.source, checkpoint: opts?.checkpoint });
        await p.onSave?.({ id: d.id, name: d.name, bytes });
        saved.current = at;
        lastSaveMs.current = performance.now() - t0;
        writeView(d.id, cur.current.grid);
        setStatus(dirty() ? { state: "unsaved" } : { state: "saved", savedAt: new Date().toISOString() });
      } catch (err) {
        if (err instanceof PendingChangesError) setStatus({ state: "review", pending: err.pending });
        else setStatus({ state: "error", message: err instanceof Error ? err.message : String(err) });
      }
      return statusRef.current;
    })();
    saving.current = run;
    try {
      return await run;
    } finally {
      saving.current = null;
      // Edits made while saving get their own save.
      if (dirty() && statusRef.current.state === "unsaved") schedule();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [buildBytes, store, setStatus]);

  const schedule = useCallback(() => {
    if (!autosaveMs || !canSave) return;
    if (timer.current) clearTimeout(timer.current);
    // A large workbook takes a while to save: leave room between saves.
    timer.current = setTimeout(() => void saveNow(), Math.max(autosaveMs, lastSaveMs.current * 3));
  }, [autosaveMs, canSave, saveNow]);

  /** Something changed: the edit counter, undo/redo state, status, autosave. */
  const changed = useCallback(() => {
    changes.current++;
    const g = cur.current.grid;
    setHistory({ canUndo: g?.canUndo?.() ?? false, canRedo: g?.canRedo?.() ?? false });
    if (statusRef.current.state === "saving") return;
    const pending = pendingCount();
    if (pending) setStatus({ state: "review", pending });
    else if (dirty()) {
      setStatus({ state: "unsaved" });
      schedule();
    }
  }, [schedule, setStatus]);

  /** The agent's batches changed (added, accepted, rejected). */
  const batchesChanged = useCallback(
    (batches: AgentBatch[]) => {
      for (const fn of listeners.current) fn(batches);
      setBatches(batches);
      // A new round of changes to review opens the panel.
      const waiting = batches.filter((b) => b.status === "pending").length;
      if (waiting > 0 && pendingBefore.current === 0 && propsRef.current.reviewPanel !== false) {
        setVersions(false);
        setReviewOpen(true);
      }
      pendingBefore.current = waiting;
      if (statusRef.current.state === "saving") return;
      const pending = batches.filter((b) => b.status === "pending").length;
      if (pending) setStatus({ state: "review", pending });
      else if (dirty()) {
        setStatus({ state: "unsaved" });
        schedule();
      } else if (statusRef.current.state === "review") setStatus({ state: "idle" });
    },
    [schedule, setStatus],
  );

  const downloadNow = useCallback(async (fileName?: string) => {
    try {
      const bytes = await buildBytes();
      download(new Blob([bytes as BlobPart], { type: XLSX_TYPE }), fileName ?? docRef.current?.name ?? "workbook.xlsx");
    } catch (err) {
      if (err instanceof PendingChangesError) notify(`Accept or reject the ${err.pending} pending AI change${err.pending === 1 ? "" : "s"} first`);
      else notify(err instanceof Error ? err.message : String(err));
    }
  }, [buildBytes, notify]);

  // --- switching workbooks ---------------------------------------------------

  const openDoc = useCallback(async (next: Omit<Doc, "key">, opts?: { thenCheckpoint?: boolean }): Promise<boolean> => {
    const pending = pendingCount();
    if (pending) {
      notify(`Accept or reject the ${pending} pending AI change${pending === 1 ? "" : "s"} first`);
      return false;
    }
    if (dirty() && cur.current.base) {
      if (store || propsRef.current.onSave) {
        const s = await saveNow({ checkpoint: true });
        if (s.state === "error" && !window.confirm(`The current workbook could not be saved (${s.message}). Open the other one anyway?`)) return false;
      } else if (!window.confirm("Discard unsaved changes?")) return false;
    }
    if (timer.current) clearTimeout(timer.current);
    setVersions(false);
    setReviewOpen(false);
    pendingCheckpoint.current = !!opts?.thenCheckpoint;
    const key = (docRef.current?.key ?? 0) + 1;
    const loaded = new Promise<void>((resolve) => loadWaiters.current.set(key, resolve));
    loading.current = loaded;
    setDoc({ ...next, key });
    await loaded;
    if (loading.current === loaded) loading.current = null;
    return true;
  }, [notify, saveNow, store]);
  const pendingCheckpoint = useRef(false);
  /** Resolvers for `open()`, by the key of the workbook being loaded. */
  const loadWaiters = useRef(new Map<number, () => void>());
  /** A workbook being opened (its load, and a restored version's first save). */
  const loading = useRef<Promise<void> | null>(null);

  const openBytes = useCallback(async (bytes: Uint8Array, name: string, id?: string) => {
    const source = fingerprint(bytes);
    let key = id ?? name;
    const stored = store ? await store.load(key).catch(() => null) : null;
    if (stored && stored.source === source && !id) {
      // The same file again: carry on from its saved state.
      return openDoc({ id: key, name: stored.name, bytes: stored.bytes, source, savedAt: stored.savedAt });
    }
    if (stored && !id) key = `${name}~${source}`; // a different file with the same name
    return openDoc({ id: key, name, bytes, source, savedAt: null });
  }, [openDoc, store]);

  const showHistory = useCallback(async () => {
    const d = docRef.current;
    if (!store || !d) return;
    setReviewOpen(false);
    setVersions(null);
    // Let an opening workbook and a save in flight finish, then save what's
    // left, so the newest entry is what's on screen.
    if (loading.current) await loading.current;
    if (saving.current) await saving.current;
    if (dirty() && !pendingCount()) await saveNow();
    setVersions(await store.versions(d.id).catch(() => []));
  }, [saveNow, store]);

  const restoreVersion = useCallback(async (versionId: string) => {
    const d = docRef.current;
    if (!store || !d) return;
    const bytes = await store.loadVersion(d.id, versionId);
    if (!bytes) return notify("That version is no longer available");
    // The state on screen is kept as a version of its own (openDoc saves it),
    // and the restored one becomes current with a checkpoint after it opens.
    await openDoc({ id: d.id, name: d.name, bytes, source: d.source, savedAt: null }, { thenCheckpoint: true });
  }, [notify, openDoc, store]);

  const downloadVersion = useCallback(async (versionId: string) => {
    const d = docRef.current;
    if (!store || !d) return;
    const bytes = await store.loadVersion(d.id, versionId);
    if (bytes) download(new Blob([bytes as BlobPart], { type: XLSX_TYPE }), d.name.replace(/\.(xlsx|xlsm)$/i, "") + " (earlier version).xlsx");
  }, [store]);

  // --- the API (one object for the component's lifetime) ---------------------
  if (!apiRef.current) {
    const grid = () => {
      const g = cur.current.grid;
      if (!g) throw new Error("the workbook is still opening");
      return g;
    };
    const api: Spreadsheet = {
      getSheetNames: () => grid().getSheetNames(),
      getCell: async (sheet, ref) => {
        const g = grid();
        const { row, col } = parseRef(ref);
        await g.whenCalculated(30_000);
        await g.prefetchRange?.(sheet, row, col, row, col);
        const c = g.getCell(sheet, row, col);
        return c ? { value: c.value, ...(c.display !== undefined ? { display: c.display } : {}), formula: c.formula } : null;
      },
      setCell: (sheet, ref, value) => {
        const { row, col } = parseRef(ref);
        grid().setCells(sheet, [{ row, col, value }]);
      },
      setCells: (sheet, cells) => void grid().setCells(sheet, cells),
      whenCalculated: (timeoutMs) => grid().whenCalculated(timeoutMs),
      select: (sheet, ref) => {
        const { row, col } = parseRef(ref);
        grid().jumpToCell(sheet, row, col);
      },
      focus: () => hostRef.current?.querySelector<HTMLElement>("[data-gridpath-grid] [tabindex='0']")?.focus(),
      undo: () => grid().undo?.(),
      redo: () => grid().redo?.(),
      save: (options) => buildBytesRef.current(options),
      saveBlob: async (options) => new Blob([(await buildBytesRef.current(options)) as BlobPart], { type: XLSX_TYPE }),
      saveNow: () => saveNowRef.current(),
      download: (fileName) => downloadRef.current(fileName),
      getStatus: () => statusRef.current,
      getDocument: () => ({ id: docRef.current?.id ?? "", name: docRef.current?.name ?? "" }),
      open: async (file, options) => {
        const bytes = new Uint8Array(file instanceof Uint8Array ? file : new Uint8Array(file));
        await openBytesRef.current(bytes, options?.name ?? "workbook.xlsx", options?.id);
      },
      destroy: () => {
        throw new Error("destroy() is only available from createSpreadsheet(); unmount the component instead");
      },
      ai: forwardingAgent(() => cur.current.ai, listeners.current),
      get internal() {
        return cur.current.grid;
      },
    };
    apiRef.current = api;
  }
  // The API's methods reach the latest callbacks through these.
  const buildBytesRef = useRef(buildBytes);
  buildBytesRef.current = buildBytes;
  const saveNowRef = useRef(saveNow);
  saveNowRef.current = saveNow;
  const downloadRef = useRef(downloadNow);
  downloadRef.current = downloadNow;
  const openBytesRef = useRef(openBytes);
  openBytesRef.current = openBytes;

  // --- loading the current workbook into the grid ----------------------------
  const tabId = doc ? `gridpath-${instance}-${doc.key}` : null;
  useEffect(() => {
    const grid = gridRef.current;
    const d = docRef.current;
    if (!engine || !grid || !d || !tabId) return;
    let live = true;
    (async () => {
      const p = propsRef.current;
      const base = d.bytes ?? (await engine.blankXlsx(p.sheets ?? ["Sheet1"]));
      const res = await grid.loadBytes(base, (s) => propsRef.current.onProgress?.(s), { tabId, path: d.name });
      if (!live) return;
      if (!res.engineLoaded) throw new Error(engine.degradedReason() ?? "the workbook could not be opened");
      const ai = new AgentToolkit(grid, engine);
      ai.onChange(batchesChanged);
      const initialFreeze = new Map<string, { rows: number; cols: number }>();
      for (const name of grid.getSheetNames()) {
        const f = grid.getFreezePanes(name);
        initialFreeze.set(name, { rows: f.freezeRows, cols: f.freezeCols });
      }
      cur.current = { grid, ai, base, sheetOps: [], initialFreeze };
      setBatches([]);
      pendingBefore.current = 0;
      changes.current = 0;
      saved.current = 0;
      setHistory({ canUndo: false, canRedo: false });
      // Back where the user was.
      const view = d.savedAt ? readView(d.id) : null;
      if (view && grid.getSheetNames().includes(view.sheet)) grid.jumpToCell(view.sheet, view.row, view.col);
      if (pendingCheckpoint.current) {
        // A restored version: it becomes the current state, as a version of its own.
        pendingCheckpoint.current = false;
        changes.current = 1;
        await saveNow({ checkpoint: true });
      } else {
        setStatus(d.savedAt ? { state: "saved", savedAt: d.savedAt } : { state: "idle" });
      }
      if (!readyCalled.current) {
        readyCalled.current = true;
        propsRef.current.onReady?.(apiRef.current!);
      } else {
        propsRef.current.onDocumentChange?.({ id: d.id, name: d.name });
      }
    })()
      .catch((err) => {
        if (live) propsRef.current.onError?.(err instanceof Error ? err : new Error(String(err)));
      })
      .finally(() => {
        loadWaiters.current.get(d.key)?.();
        loadWaiters.current.delete(d.key);
      });
    return () => {
      live = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [engine, tabId]);

  // --- leaving the page: save what's left, warn about what can't be ---------
  useEffect(() => {
    const flush = () => {
      if (document.visibilityState !== "hidden") return;
      writeView(docRef.current?.id ?? "", cur.current.grid);
      if (dirty() && !pendingCount()) void saveNowRef.current();
    };
    const warn = (e: BeforeUnloadEvent) => {
      if (dirty() || saving.current) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    document.addEventListener("visibilitychange", flush);
    window.addEventListener("beforeunload", warn);
    return () => {
      document.removeEventListener("visibilitychange", flush);
      window.removeEventListener("beforeunload", warn);
      if (timer.current) clearTimeout(timer.current);
      if (noticeTimer.current) clearTimeout(noticeTimer.current);
    };
  }, []);

  // ⌘S / Ctrl+S: save (or download when nothing is configured to save to).
  const onKeyDown = (e: React.KeyboardEvent) => {
    if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === "s") {
      e.preventDefault();
      if (canSave) void saveNow();
      else void downloadNow();
    }
  };

  const actions = useMemo(() => toolbarActions(() => cur.current.grid, notify, () => changed()), [notify, changed]);
  const showToolbar = props.toolbar !== false;

  return (
    <div
      ref={hostRef}
      className={props.className}
      onKeyDown={onKeyDown}
      // Not clipped: the toolbar's dropdowns may reach past the component's bottom edge.
      style={{ position: "relative", width: "100%", height: "100%", display: "flex", flexDirection: "column", ...props.style }}
    >
      {showToolbar && mounted && (
        <Chrome
          actions={actions}
          disabled={!doc}
          canUndo={history.canUndo}
          canRedo={history.canRedo}
          status={canSave ? status : status.state === "review" ? status : dirty() ? { state: "unsaved" } : { state: "idle" }}
          canSave={canSave}
          onSave={() => void saveNow()}
          onDownload={() => void downloadNow()}
          fileMenu={props.fileMenu !== false}
          onOpenFile={(f) => void f.arrayBuffer().then((b) => openBytes(new Uint8Array(b), f.name))}
          onNew={() => void openDoc({ id: `untitled-${Date.now().toString(36)}`, name: "Untitled.xlsx", bytes: null, source: "blank", savedAt: null })}
          listRecent={store ? () => store.list() : null}
          onOpenRecent={(id) =>
            void store?.load(id).then((s) => s && openDoc({ id, name: s.name, bytes: s.bytes, source: s.source ?? "blank", savedAt: s.savedAt }))
          }
          currentId={doc?.id ?? ""}
          onShowHistory={store ? () => void showHistory() : null}
          onReview={props.reviewPanel === false ? null : () => {
            setVersions(false);
            setReviewOpen(true);
          }}
          notice={notice}
        />
      )}
      <div data-gridpath-grid="" style={{ position: "relative", flex: 1, minHeight: 0, overflow: "hidden" }}>
        {engine && doc && tabId && (
          <ViewportGrid
            key={tabId}
            ref={gridRef}
            tabId={tabId}
            workbookPath={doc.name}
            engine={engine}
            onUserEdit={() => propsRef.current.onEdit?.()}
            onChange={changed}
            onManualSheetOp={(op: ManualSheetOp) => {
              appendManualSheetOp(cur.current.sheetOps, op);
              changed();
            }}
          />
        )}
        {reviewOpen && doc && (
          <ReviewPanel
            batches={batches}
            onAccept={(id) => cur.current.ai?.accept(id)}
            onReject={(id) => void cur.current.ai?.reject(id)}
            onJump={(c) => cur.current.grid?.jumpToCell(c.sheet, c.row, c.col)}
            onClose={() => setReviewOpen(false)}
          />
        )}
        {versions !== false && doc && (
          <HistoryPanel
            name={doc.name}
            versions={versions}
            onRestore={(id) => void restoreVersion(id)}
            onDownload={(id) => void downloadVersion(id)}
            onClose={() => setVersions(false)}
          />
        )}
      </div>
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
        onReady={(api) => {
          api.destroy = destroy;
          resolve(api);
        }}
        onError={(err) => {
          destroy();
          reject(err);
        }}
      />,
    );
  });
}
