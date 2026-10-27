/**
 * gridpath: a spreadsheet for web apps. Opens real .xlsx files (or CSV
 * and other delimited text, or starts blank), calculates with GridPath's Rust engine compiled to
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
import { BrowserEngine, type EngineOptions, type FileSavedCell } from "./browserEngine";
import { AgentToolkit, type AgentBatch, type SheetAgent } from "./ai/toolkit";
import { Chrome, ConfirmDialog, EngineBanner, HistoryPanel, IssuesPanel, LoadError, ReviewPanel, type SaveStatus } from "./ui/Chrome";
import { toolbarActions } from "./ui/actions";
import { a1 } from "./ui/review";
import { browserStore, fingerprint, localDrafts, type WorkbookStore, type WorkbookVersion } from "./ui/store";
import {
  GridPathError,
  PendingChangesError,
  UnsavedChangesError,
  calcTimeout,
  engineError,
  engineStartError,
  fileProblem,
  isDelimitedText,
  notReady,
  openError,
  saveFailed,
  storageError,
  textTableError,
  workbookName,
} from "./errors";

export { BrowserEngine, type EngineOptions } from "./browserEngine";
export type { GridEngine, DeltaApplier } from "./engine";
export { INSTRUCTIONS as AGENT_INSTRUCTIONS, type SheetAgent, type AgentBatch, type ToolDefinition, type OpenAIToolDefinition, type ReadCell } from "./ai/toolkit";
export { browserStore, type BrowserStoreOptions, type WorkbookStore, type StoredWorkbook, type WorkbookVersion } from "./ui/store";
export { httpStore, HttpStoreError, type HttpStoreOptions } from "./ui/httpStore";
export { GridPathError, PendingChangesError, UnsavedChangesError, type GridPathErrorCode } from "./errors";
export type { SaveStatus } from "./ui/Chrome";

export interface SpreadsheetOptions {
  /** An .xlsx to open, or CSV or other delimited text (it opens as a new
   * .xlsx). Without it the workbook starts blank (or reopens the stored one,
   * see `documentId`). */
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
   * this browser (`browserStore()`, IndexedDB). Pass `httpStore({ url })`
   * (or your own) to save to your backend, or `false` for none (then
   * `save()` / Download only).
   */
  store?: WorkbookStore | false;
  /** Save automatically once edits pause: true (2 s), a delay in ms, or
   * false. Needs a store or `onSave`. Default true. */
  autosave?: boolean | number;
  /** Called with the workbook's bytes on every save (yours to upload). A
   * rejection counts as a failed save, retried like a store's. */
  onSave?: (workbook: { id: string; name: string; bytes: Uint8Array }) => void | Promise<void>;
  /** The save status changed ("Saving…", "Saved", "Unsaved changes", a
   * failed save with its `code` and `hint`…). */
  onStatusChange?: (status: SaveStatus) => void;
  /** Another workbook was opened (File menu, recent list, a restored version, `open()`). */
  onDocumentChange?: (doc: { id: string; name: string }) => void;
  /** Called after each user edit in the grid. */
  onEdit?: () => void;
  /** Loading stages ("Loading workbook…", "Building grid…"). */
  onProgress?: (stage: string) => void;
  /** The engine failed, or a workbook could not be opened (the first one,
   * or one from the File menu). Failed saves arrive through
   * `onStatusChange` instead. */
  onError?: (error: GridPathError) => void;
  /** Where and how the engine runs. */
  engine?: EngineOptions;
}

export interface CellValue {
  /** The raw value: a number, string, boolean, or null when empty. */
  value: unknown;
  /** The formatted text shown in the cell, when it differs from `value`. */
  display?: string;
  /** The formula, starting with "=", or null for a constant. */
  formula: string | null;
  /** True when the value is the one Excel saved in the file, because
   * GridPath couldn't calculate this formula. */
  fromFile?: boolean;
  /** Why, in a sentence to show ("Showing Excel's saved value because…"). */
  note?: string;
}

/** A formula showing Excel's saved value because GridPath couldn't calculate it. */
export interface CalcIssue {
  sheet: string;
  /** A1 reference ("B7"), and the same cell 0-based. */
  ref: string;
  row: number;
  col: number;
  /** The formula, starting with "=". */
  formula: string;
  /** The value Excel saved, which the cell shows. */
  value: unknown;
  /** GridPath's own result when it calculated one (an error such as
   * "#VALUE!"), null when the formula is kept as Excel saved it (an
   * unknown or add-in function, a link to another workbook). */
  calculated: unknown;
  /** Why, in a sentence to show. */
  note: string;
}

export interface SaveOptions {
  /** Save agent changes the user has not accepted yet. Default false:
   * `save()` throws a `PendingChangesError` while any are pending. */
  includePending?: boolean;
}

export interface OpenOptions {
  /** The file name (downloads, the recent list); default "workbook.xlsx". */
  name?: string;
  /** The id to save it under in the store; default the name (or, when the
   * store has a different file by that name, the name plus a fingerprint). */
  id?: string;
  /** Open even when the current workbook's changes could not be saved: they are dropped. */
  discardUnsaved?: boolean;
}

export interface Spreadsheet {
  /** Native chart descriptors and current source values, after calculation. */
  getCharts(): Promise<import("./charts/types").NativeChart[]>;
  /** Breakdown of the last engine edit; milliseconds, excluding caller/UI waits. */
  getLastEditTimings(): import("./charts/types").EditTimings | null;
  /** Sheet names in workbook order. */
  getSheetNames(): string[];
  /**
   * One cell by A1 reference ("B7"), read after pending writes are
   * recalculated (waiting up to 30 s; then it rejects with a
   * `calculation-timeout` error). A formula GridPath couldn't calculate
   * shows the value Excel saved: `fromFile` is true and `note` says why.
   */
  getCell(sheet: string, ref: string): Promise<CellValue | null>;
  /** Every formula showing Excel's saved value because GridPath couldn't
   * calculate it (unsupported functions, links to other workbooks…). */
  getCalcIssues(limit?: number): Promise<CalcIssue[]>;
  /** Write one cell by A1 reference. Strings starting with "=" are formulas.
   * Returns at once: writes are queued and recalculated in order, so a
   * `getCell` after it sees the result. */
  setCell(sheet: string, ref: string, value: string | number | boolean | null): void;
  /** Write many cells on one sheet in one step (0-based row and col). */
  setCells(sheet: string, cells: Array<{ row: number; col: number; value: string | number | boolean | null }>): void;
  /** Resolves once pending writes are recalculated; rejects with a
   * `calculation-timeout` error after `timeoutMs` (default 30 000). */
  whenCalculated(timeoutMs?: number): Promise<void>;
  /** Scroll to and select a cell. */
  select(sheet: string, ref: string): void;
  /** Give the grid keyboard focus (typing then edits the selected cell). */
  focus(): void;
  /** Undo / redo the last change (the user's or an agent's). */
  undo(): void;
  redo(): void;
  /**
   * The workbook as .xlsx bytes: the original file with this session's
   * changes patched in (for a blank workbook, a new file). It waits for
   * recalculation (up to 60 s, else `calculation-timeout`), since results
   * are cached into the file. While agent changes are pending review it
   * rejects with a `PendingChangesError`: accept or reject them first
   * (`ai.accept()`, `ai.reject()`), or pass `{ includePending: true }`.
   * `save-failed` when the changes can't be written into the file.
   * It does not touch the store: see `saveNow()`.
   */
  save(options?: SaveOptions): Promise<Uint8Array>;
  /** `save()` as a Blob with the .xlsx media type, ready to download or upload. */
  saveBlob(options?: SaveOptions): Promise<Blob>;
  /**
   * Save to the store and `onSave` now (what autosave does). Resolves with
   * the new status and never rejects: a failure is `{ state: "error", code,
   * message, hint }`. Storage failures are retried (after 2, 5, 15 and 30
   * seconds, then every minute) and, with a remote store, the changes are
   * kept in this browser meanwhile and reopened from there if the page is
   * closed before they reach the store.
   */
  saveNow(): Promise<SaveStatus>;
  /** Download the workbook as an .xlsx file. */
  download(fileName?: string): Promise<void>;
  /** Where the workbook stands relative to its last save. */
  getStatus(): SaveStatus;
  /** True while there are edits the store doesn't have yet (AI changes
   * awaiting review included). */
  hasUnsavedChanges(): boolean;
  /** The open workbook. */
  getDocument(): { id: string; name: string };
  /**
   * Open another workbook in this spreadsheet.
   * - AI changes awaiting review: rejects with `PendingChangesError`.
   * - Unsaved changes are saved first (store / `onSave`). When they can't
   *   be (or nothing is set up to save to) it rejects with
   *   `UnsavedChangesError`, unless `discardUnsaved`. A remote store's
   *   failed save still counts as saved when the changes are kept in this
   *   browser.
   * - CSV and other delimited text open as a new .xlsx (sales.csv saves
   *   as sales.xlsx).
   * - A file that can't be opened rejects with `unsupported-file` or
   *   `malformed-file`, and the previous workbook stays open.
   */
  open(file: ArrayBuffer | Uint8Array, options?: OpenOptions): Promise<void>;
  /** Open a workbook from the store by id, as the File menu's Recent list
   * does (`not-found` when there is none); the rules of `open()` apply. */
  openStored(id: string, options?: { discardUnsaved?: boolean }): Promise<void>;
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
/** Waits after failed saves; the last one repeats. */
const RETRY_MS = [2_000, 5_000, 15_000, 30_000, 60_000];
const CALC_WAIT_MS = 30_000;
const SAVE_CALC_WAIT_MS = 60_000;
/** The calculation-issues panel lists at most this many. */
const MAX_ISSUES = 200;

let instances = 0;

/** The workbook on screen. `key` remounts the grid when it changes.
 * `draft`: reopened from this browser's copy of changes the store doesn't
 * have yet, so it saves again once open. */
type Doc = { key: number; id: string; name: string; bytes: Uint8Array | null; source: string; savedAt: string | null; draft?: boolean };

/** What belongs to the open workbook (replaced when another opens). */
type Current = {
  /** The doc these belong to: until the next one has opened, the previous one's. */
  key: number;
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

const asGridPathError = (err: unknown, fallback: (e: unknown) => GridPathError): GridPathError =>
  err instanceof GridPathError ? err : fallback(err);

/** Why a formula shows Excel's saved value, in a sentence. */
function calcNote(r: FileSavedCell): string {
  const fn = (r.name ?? "").replace(/^_xl(fn|ws)\./i, "").toUpperCase();
  const err = typeof r.live === "string" && r.live.startsWith("#") ? r.live : null;
  const why =
    r.why === "function"
      ? /^_xll\./i.test(r.name ?? "")
        ? `it calls ${fn.replace(/^_XLL\./, "")}, an Excel add-in function`
        : `it uses ${fn}, which GridPath doesn't calculate yet`
      : r.why === "link"
        ? "it refers to another workbook"
        : r.why === "syntax"
          ? "GridPath couldn't read the formula"
          : err === "#NAME?" || err === "#N/IMPL!"
            ? "it uses a function GridPath doesn't calculate yet"
            : err
              ? `GridPath's result was ${err}`
              : "GridPath's result was empty";
  return `Showing Excel's saved value because GridPath couldn't calculate this formula: ${why}.`;
}

function toIssue(r: FileSavedCell): CalcIssue {
  return {
    sheet: r.sheet,
    ref: a1(r.row, r.col),
    row: r.row,
    col: r.col,
    formula: r.formula ?? "",
    value: r.shown,
    calculated: r.why === "error" ? r.live : null,
    note: calcNote(r),
  };
}

/** `sheet.ai` stays the same object across workbooks; listeners carry over. */
function forwardingAgent(current: () => AgentToolkit | null, listeners: Set<(b: AgentBatch[]) => void>): SheetAgent {
  const need = () => {
    const ai = current();
    if (!ai) throw notReady();
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
  className?: string;
  style?: React.CSSProperties;
}

type Panel = "review" | "history" | "issues" | null;
type Confirm = { title: string; body: string; action: string; run: () => void };

/**
 * The React component. It fills its parent; give the parent a height.
 * `file` is read when it mounts; to open another file, call `open()` (or
 * use the File menu), or remount with a new `key`.
 */
export function Spreadsheet(props: SpreadsheetProps): React.ReactElement {
  const instance = useMemo(() => ++instances, []);
  const [engine, setEngine] = useState<BrowserEngine | null>(null);
  const [doc, setDoc] = useState<Doc | null>(null);
  /** The doc on screen finished opening. */
  const [loaded, setLoaded] = useState(false);
  /** Why there is no workbook on screen (the engine, or the file). */
  const [loadError, setLoadError] = useState<GridPathError | null>(null);
  /** The engine stopped after the workbook opened. */
  const [engineDown, setEngineDown] = useState<GridPathError | null>(null);
  const [status, setStatusState] = useState<SaveStatus>({ state: "idle" });
  const [history, setHistory] = useState({ canUndo: false, canRedo: false });
  const [notice, setNotice] = useState<string | null>(null);
  const [panel, setPanel] = useState<Panel>(null);
  const [versions, setVersions] = useState<WorkbookVersion[] | null>(null);
  const [batches, setBatches] = useState<AgentBatch[]>([]);
  const [issues, setIssues] = useState<CalcIssue[]>([]);
  const [confirm, setConfirm] = useState<Confirm | null>(null);
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
  const engineRef = useRef<BrowserEngine | null>(null);
  engineRef.current = engine;
  const loadedRef = useRef(false);
  loadedRef.current = loaded;
  /** Unmounted: the API answers not-ready from then on. */
  const gone = useRef(false);
  const issuesRef = useRef(issues);
  issuesRef.current = issues;
  const keyRef = useRef(0);
  const cur = useRef<Current>({ key: 0, grid: null, ai: null, base: null, sheetOps: [], initialFreeze: new Map() });
  const statusRef = useRef<SaveStatus>(status);
  const listeners = useRef(new Set<(b: AgentBatch[]) => void>());
  const apiRef = useRef<Spreadsheet | null>(null);
  const readyCalled = useRef(false);
  // Change bookkeeping: `changes` counts edits, `saved` is its value at the
  // last good save; they differ while there are unsaved changes.
  const changes = useRef(0);
  const saved = useRef(0);
  /** The workbook as of its last save (or as it opened): what comes back
   * when opening another one fails. */
  const lastBytes = useRef<Uint8Array | null>(null);
  const saving = useRef<Promise<SaveStatus> | null>(null);
  const lastSaveMs = useRef(0);
  /** Failed saves in a row (the retry delay grows with them). */
  const attempts = useRef(0);
  /** This browser holds a draft of the open workbook the store lacks. */
  const hasDraft = useRef(false);
  /** The previous workbook is being reopened after a failed open. */
  const restoring = useRef(false);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const retryTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const noticeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const issuesTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  // `store` and friends are read once: the default store is this browser's.
  const store = useMemo<WorkbookStore | null>(() => {
    if (props.store === false) return null;
    if (props.store) return props.store;
    return typeof indexedDB === "undefined" ? null : browserStore();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  // A remote store's unsaved changes wait in this browser.
  const drafts = useMemo(() => (store && !store.local && typeof indexedDB !== "undefined" ? localDrafts() : null), [store]);
  const canSave = !!store || !!props.onSave;
  const autosaveMs = props.autosave === false ? 0 : typeof props.autosave === "number" ? Math.max(250, props.autosave) : AUTOSAVE_MS;

  const setStatus = useCallback((s: SaveStatus) => {
    statusRef.current = s;
    setStatusState(s);
    propsRef.current.onStatusChange?.(s);
  }, []);

  const notify = useCallback((message: string, ms = 2500) => {
    setNotice(message);
    if (noticeTimer.current) clearTimeout(noticeTimer.current);
    noticeTimer.current = setTimeout(() => setNotice(null), ms);
  }, []);

  const report = useCallback((err: GridPathError) => propsRef.current.onError?.(err), []);
  /** The save to run on unmount, if there is anything to save. */
  const unmountSave = useRef<() => Promise<unknown> | null>(() => null);

  const pendingCount = () => cur.current.ai?.list().filter((b) => b.status === "pending").length ?? 0;
  const dirty = () => changes.current !== saved.current;
  const clearSaveTimers = () => {
    if (timer.current) clearTimeout(timer.current);
    if (retryTimer.current) clearTimeout(retryTimer.current);
    timer.current = retryTimer.current = null;
  };

  // --- engine ----------------------------------------------------------------
  const engineOpts = props.engine;
  useEffect(() => {
    let live = true;
    let created: BrowserEngine | null = null;
    BrowserEngine.create(engineOpts).then(
      (e) => {
        created = e;
        if (!live) return e.close();
        e.onFailure((reason) => {
          const err = engineError(reason);
          setEngineDown(err);
          report(err);
        });
        setEngine(e);
      },
      (cause) => {
        if (!live) return;
        const err = engineStartError(cause);
        setLoadError(err);
        report(err);
      },
    );
    gone.current = false;
    return () => {
      live = false;
      gone.current = true;
      const e = created;
      if (!e) return;
      // Unmounted with edits not saved yet (a client-side navigation, a
      // remount for another workbook): save them, then stop the engine.
      const flush = unmountSave.current();
      if (flush) void flush.finally(() => e.close());
      else e.close();
    };
    // The engine is created once per mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /** A stored workbook, or this browser's newer draft of it; null when the store has neither. */
  const fromStore = useCallback(
    async (id: string): Promise<{ name: string; bytes: Uint8Array; savedAt: string | null; source?: string; draft?: boolean } | null> => {
      if (!store) return null;
      let stored: Awaited<ReturnType<WorkbookStore["load"]>> = null;
      let failure: unknown = null;
      try {
        stored = await store.load(id);
      } catch (err) {
        failure = err;
      }
      const draft = drafts ? await drafts.get(id).catch(() => null) : null;
      if (draft && (!stored || Date.parse(draft.savedAt) >= (Date.parse(stored.savedAt) || 0))) {
        return { name: draft.name, bytes: draft.bytes, savedAt: null, source: draft.source, draft: true };
      }
      if (failure) throw storageError(`load "${id}"`, failure);
      return stored && { name: stored.name, bytes: stored.bytes, savedAt: stored.savedAt, source: stored.source };
    },
    [store, drafts],
  );

  // --- the first workbook: the host's file, or the stored one ----------------
  const openFirst = useCallback(async () => {
    const p = propsRef.current;
    const id = p.documentId ?? p.fileName ?? "Untitled.xlsx";
    const file = p.file ? new Uint8Array(p.file instanceof Uint8Array ? p.file : new Uint8Array(p.file)) : null;
    const source = file ? fingerprint(file) : "blank";
    let found: Awaited<ReturnType<typeof fromStore>> = null;
    try {
      found = await fromStore(id);
    } catch (err) {
      // This browser's storage is unavailable: open the file without it. A
      // remote store that can't be reached stops here instead, so an
      // autosave never replaces the user's saved workbook with the file.
      if (!store?.local) throw err;
    }
    // Reopen where the user left off, unless the host passed a different file.
    if (found && (!file || found.source === source)) {
      setDoc({ key: ++keyRef.current, id, name: found.name, bytes: found.bytes, source: found.source ?? source, savedAt: found.savedAt, draft: found.draft });
    } else {
      setDoc({ key: ++keyRef.current, id, name: workbookName(p.fileName ?? "Untitled.xlsx", file), bytes: file, source, savedAt: null });
    }
  }, [fromStore, store]);
  const retryFirst = useCallback(() => {
    setLoadError(null);
    openFirst().catch((err) => {
      const e = asGridPathError(err, (x) => storageError("load the workbook", x));
      setLoadError(e);
      report(e);
    });
  }, [openFirst, report]);
  useEffect(() => {
    retryFirst();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // --- saving ----------------------------------------------------------------

  /** The workbook's bytes now (the patch of this session onto the file it opened from). */
  const buildBytes = useCallback(async (options?: SaveOptions): Promise<Uint8Array> => {
    const { grid, ai, base } = cur.current;
    // Only the workbook on screen, once it has opened (never the previous one's grid under the next one's id).
    if (!grid || !ai || !base || !engine || !loadedRef.current || cur.current.key !== docRef.current?.key) throw notReady();
    // Review comes first: unaccepted agent edits are not saved by accident.
    const pending = pendingCount();
    if (pending && !options?.includePending) throw new PendingChangesError(pending);
    // Formula results are cached into the file: let recalculation finish.
    if (!(await engine.settledWithin(SAVE_CALC_WAIT_MS))) throw calcTimeout(SAVE_CALC_WAIT_MS);
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
    if (!built.ok) throw saveFailed(`${built.reason}${(built as any).detail ? `: ${(built as any).detail}` : ""}`);
    // Every save replays the whole session onto the ORIGINAL bytes.
    try {
      return await engine.patch(base, JSON.stringify(built.patch));
    } catch (err) {
      const reason = err instanceof Error ? err.message : String(err);
      throw engine.alive() ? saveFailed(reason, err) : engineError(reason, err);
    }
  }, [engine]);

  const saveNow = useCallback(async (opts?: { checkpoint?: boolean }): Promise<SaveStatus> => {
    const d = docRef.current;
    const p = propsRef.current;
    if (!d || !cur.current.base || !loadedRef.current || cur.current.key !== d.key || (!store && !p.onSave)) return statusRef.current;
    clearSaveTimers();
    const pending = pendingCount();
    if (pending) {
      setStatus({ state: "review", pending });
      return statusRef.current;
    }
    if (saving.current) return saving.current;
    /** A failed save: the status says what and what to do; transient ones retry. */
    const failed = (err: GridPathError, retry: boolean, keptLocally = false) => {
      const delay = retry && autosaveMs ? RETRY_MS[Math.min(attempts.current, RETRY_MS.length - 1)] : 0;
      attempts.current++;
      setStatus({
        state: "error",
        code: err.code,
        message: err.message,
        hint: err.hint,
        ...(keptLocally ? { keptLocally } : {}),
        ...(delay ? { retryAt: new Date(Date.now() + delay).toISOString() } : {}),
      });
      if (delay) retryTimer.current = setTimeout(() => void saveNowRef.current(), delay);
    };
    const run = (async (): Promise<SaveStatus> => {
      const at = changes.current;
      const t0 = performance.now();
      setStatus({ state: "saving" });
      let bytes: Uint8Array;
      try {
        bytes = await buildBytes();
      } catch (err) {
        if (err instanceof PendingChangesError) setStatus({ state: "review", pending: err.pending });
        else {
          const e = asGridPathError(err, (x) => saveFailed(x instanceof Error ? x.message : String(x), x));
          failed(e, e.code === "calculation-timeout");
        }
        return statusRef.current;
      }
      try {
        if (store) await store.save({ id: d.id, name: d.name, bytes, source: d.source, checkpoint: opts?.checkpoint });
        await p.onSave?.({ id: d.id, name: d.name, bytes });
      } catch (err) {
        // Kept in this browser until the store has it.
        let kept = false;
        if (drafts) {
          kept = await drafts
            .put({ id: d.id, name: d.name, bytes, source: d.source, savedAt: new Date().toISOString() })
            .then(() => true, () => false);
          if (kept) hasDraft.current = true;
        }
        failed(storageError(`save "${d.name}"`, err, { kept, retrying: !!autosaveMs }), true, kept);
        return statusRef.current;
      }
      saved.current = at;
      lastBytes.current = bytes;
      lastSaveMs.current = performance.now() - t0;
      attempts.current = 0;
      writeView(d.id, cur.current.grid);
      if (hasDraft.current && drafts) {
        hasDraft.current = false;
        void drafts.remove(d.id).catch(() => {});
      }
      setStatus(dirty() ? { state: "unsaved" } : { state: "saved", savedAt: new Date().toISOString(), ...(store?.label ? { where: store.label } : {}) });
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
  }, [buildBytes, store, drafts, autosaveMs, setStatus]);

  const schedule = useCallback(() => {
    if (!autosaveMs || !canSave) return;
    if (timer.current) clearTimeout(timer.current);
    // A large workbook takes a while to save: leave room between saves.
    timer.current = setTimeout(() => void saveNow(), Math.max(autosaveMs, lastSaveMs.current * 3));
  }, [autosaveMs, canSave, saveNow]);

  /** Re-read which formulas show Excel's saved value (after the workbook
   * settles; edits can only fix them). */
  const refreshIssues = useCallback((delay = 300) => {
    const e = engine;
    if (!e) return;
    if (issuesTimer.current) clearTimeout(issuesTimer.current);
    issuesTimer.current = setTimeout(async () => {
      const key = docRef.current?.key;
      await e.settled(600_000);
      const rows = await e.fileSavedCells(MAX_ISSUES).catch((): FileSavedCell[] => []);
      if (docRef.current?.key === key) setIssues(rows.map(toIssue));
    }, delay);
  }, [engine]);

  /** Something changed: the edit counter, undo/redo state, status, autosave. */
  const changed = useCallback(() => {
    changes.current++;
    const g = cur.current.grid;
    setHistory({ canUndo: g?.canUndo?.() ?? false, canRedo: g?.canRedo?.() ?? false });
    if (issuesRef.current.length) refreshIssues(1500);
    if (statusRef.current.state === "saving") return;
    // A failed save waiting to retry keeps its status (the retry saves these edits too).
    if (statusRef.current.state === "error" && retryTimer.current) return;
    const pending = pendingCount();
    if (pending) setStatus({ state: "review", pending });
    else if (dirty()) {
      setStatus({ state: "unsaved" });
      schedule();
    }
  }, [schedule, setStatus, refreshIssues]);

  /** The agent's batches changed (added, accepted, rejected). */
  const batchesChanged = useCallback(
    (batches: AgentBatch[]) => {
      for (const fn of listeners.current) fn(batches);
      setBatches(batches);
      // A new round of changes to review opens the panel.
      const waiting = batches.filter((b) => b.status === "pending").length;
      if (waiting > 0 && pendingBefore.current === 0 && propsRef.current.reviewPanel !== false) setPanel("review");
      pendingBefore.current = waiting;
      if (statusRef.current.state === "saving") return;
      if (waiting) setStatus({ state: "review", pending: waiting });
      else if (dirty()) {
        setStatus({ state: "unsaved" });
        schedule();
      } else if (statusRef.current.state === "review") setStatus({ state: "idle" });
    },
    [schedule, setStatus],
  );

  const showError = useCallback((err: GridPathError) => notify(`${err.message}. ${err.hint}`, 6000), [notify]);

  const downloadNow = useCallback(async (fileName?: string) => {
    try {
      const bytes = await buildBytes();
      download(new Blob([bytes as BlobPart], { type: XLSX_TYPE }), fileName ?? docRef.current?.name ?? "workbook.xlsx");
    } catch (err) {
      if (err instanceof PendingChangesError) notify(`Accept or reject the ${err.pending} pending AI change${err.pending === 1 ? "" : "s"} first`);
      else showError(asGridPathError(err, (x) => saveFailed(String(x), x)));
    }
  }, [buildBytes, notify, showError]);

  // --- switching workbooks ---------------------------------------------------

  /** Resolvers for the load of each doc, by key. */
  const loadWaiters = useRef(new Map<number, { resolve: () => void; reject: (e: GridPathError) => void }>());
  /** A workbook being opened (its load, and a restored version's first save). */
  const loading = useRef<Promise<void> | null>(null);
  const pendingCheckpoint = useRef(false);

  /** Put `next` on screen; settles when it has opened (or failed to). */
  const loadDoc = useCallback((next: Omit<Doc, "key">): Promise<void> => {
    const key = ++keyRef.current;
    const done = new Promise<void>((resolve, reject) => loadWaiters.current.set(key, { resolve, reject }));
    const settled = done.then(() => {}, () => {});
    loading.current = settled;
    void settled.then(() => {
      if (loading.current === settled) loading.current = null;
    });
    setLoaded(false);
    setDoc({ ...next, key });
    return done;
  }, []);

  const openDoc = useCallback(async (next: Omit<Doc, "key">, opts: { thenCheckpoint?: boolean; discardUnsaved?: boolean } = {}): Promise<void> => {
    if (!engine) throw notReady();
    const pending = pendingCount();
    if (pending) throw new PendingChangesError(pending);
    const prev = docRef.current;
    const prevOpen = loadedRef.current && !!cur.current.base && prev ? prev : null;
    if (prevOpen && dirty() && !opts.discardUnsaved) {
      if (!canSave) throw new UnsavedChangesError(prevOpen.name);
      const s = await saveNow({ checkpoint: true });
      if (dirty() && !(s.state === "error" && s.keptLocally)) throw new UnsavedChangesError(prevOpen.name);
    }
    clearSaveTimers();
    setPanel(null);
    // Dropped (discardUnsaved): nothing may save them on the way out.
    saved.current = changes.current;
    // What comes back if `next` can't be opened: the last saved state.
    const st = statusRef.current;
    const restore: Omit<Doc, "key"> | null =
      prevOpen && lastBytes.current
        ? { ...prevOpen, bytes: lastBytes.current, savedAt: st.state === "saved" ? st.savedAt : prevOpen.savedAt, draft: hasDraft.current }
        : null;
    pendingCheckpoint.current = !!opts.thenCheckpoint;
    try {
      await loadDoc(next);
    } catch (err) {
      pendingCheckpoint.current = false;
      const e = asGridPathError(err, (x) => engineError(String(x), x));
      if (restore) {
        restoring.current = true;
        try {
          await loadDoc(restore);
        } catch (again) {
          setLoadError(asGridPathError(again, (x) => engineError(String(x), x)));
        } finally {
          restoring.current = false;
        }
      } else setLoadError(e);
      throw e;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [engine, canSave, saveNow, loadDoc]);

  const openBytes = useCallback(async (bytes: Uint8Array, name: string, opts: { id?: string; discardUnsaved?: boolean } = {}) => {
    // Plainly not an .xlsx: say so without leaving the current workbook.
    const problem = fileProblem(bytes, name);
    if (problem) throw problem;
    const source = fingerprint(bytes);
    let key = opts.id ?? name;
    if (!opts.id && store) {
      let found: Awaited<ReturnType<typeof fromStore>> = null;
      let unknown = false;
      try {
        found = await fromStore(key);
      } catch {
        unknown = true;
      }
      // The same file again: carry on from its saved state.
      if (found && found.source === source) {
        return openDoc({ id: key, name: found.name, bytes: found.bytes, source, savedAt: found.savedAt, draft: found.draft }, opts);
      }
      // A different file by that name (or the store can't say): its own id.
      if (found || unknown) key = `${name}~${source}`;
    }
    return openDoc({ id: key, name: workbookName(name, bytes), bytes, source, savedAt: null }, opts);
  }, [openDoc, store, fromStore]);

  const openStored = useCallback(async (id: string, opts: { discardUnsaved?: boolean } = {}) => {
    const found = await fromStore(id);
    if (!found) {
      throw new GridPathError(
        "not-found",
        store ? `No saved workbook "${id}"` : "There is no store to open from",
        store ? "It may have been deleted; open a file instead." : "Pass a store to the spreadsheet (the default is this browser's).",
      );
    }
    return openDoc({ id, name: found.name, bytes: found.bytes, source: found.source ?? "blank", savedAt: found.savedAt, draft: found.draft }, opts);
  }, [fromStore, openDoc, store]);

  /** An open from the toolbar: failures become a message, unsaved changes a question. */
  const uiOpen = useCallback((run: (discardUnsaved: boolean) => Promise<void>) => {
    const attempt = (discardUnsaved: boolean): Promise<void> =>
      run(discardUnsaved).catch((err) => {
        if (err instanceof PendingChangesError) {
          notify(`Accept or reject the ${err.pending} pending AI change${err.pending === 1 ? "" : "s"} first`);
          if (propsRef.current.reviewPanel !== false) setPanel("review");
        } else if (err instanceof UnsavedChangesError) {
          setConfirm({
            title: "Discard unsaved changes?",
            body: `${err.message}. Opening another workbook drops them; download it first to keep a copy.`,
            action: "Discard and open",
            run: () => {
              setConfirm(null);
              void attempt(true);
            },
          });
        } else {
          const e = asGridPathError(err, (x) => engineError(String(x), x));
          showError(e);
          report(e);
        }
      });
    void attempt(false);
  }, [notify, report, showError]);

  const showHistory = useCallback(async () => {
    const d = docRef.current;
    if (!store || !d) return;
    setPanel("history");
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
    let bytes: Uint8Array | null;
    try {
      bytes = await store.loadVersion(d.id, versionId);
    } catch (err) {
      return showError(storageError("load that version", err));
    }
    if (!bytes) return notify("That version is no longer available");
    const restored = bytes;
    // The state on screen is kept as a version of its own (openDoc saves it),
    // and the restored one becomes current with a checkpoint after it opens.
    uiOpen((discardUnsaved) => openDoc({ id: d.id, name: d.name, bytes: restored, source: d.source, savedAt: null }, { thenCheckpoint: true, discardUnsaved }));
  }, [notify, openDoc, showError, store, uiOpen]);

  const downloadVersion = useCallback(async (versionId: string) => {
    const d = docRef.current;
    if (!store || !d) return;
    try {
      const bytes = await store.loadVersion(d.id, versionId);
      if (bytes) download(new Blob([bytes as BlobPart], { type: XLSX_TYPE }), d.name.replace(/\.(xlsx|xlsm)$/i, "") + " (earlier version).xlsx");
    } catch (err) {
      showError(storageError("load that version", err));
    }
  }, [showError, store]);

  // --- the API (one object for the component's lifetime) ---------------------
  if (!apiRef.current) {
    const grid = () => {
      const g = cur.current.grid;
      if (gone.current) throw notReady(true);
      if (!g || !loadedRef.current) throw notReady();
      return g;
    };
    const calculated = async (ms = CALC_WAIT_MS) => {
      const e = engineRef.current;
      if (gone.current) throw notReady(true);
      if (!e) throw notReady();
      if (!(await e.settledWithin(ms))) throw calcTimeout(ms);
    };
    const api: Spreadsheet = {
      getCharts: async () => { grid(); await calculated(); return structuredClone(engineRef.current!.charts()); },
      getLastEditTimings: () => engineRef.current?.lastEditTimings() ?? null,
      getSheetNames: () => grid().getSheetNames(),
      getCell: async (sheet, ref) => {
        const g = grid();
        const { row, col } = parseRef(ref);
        await calculated();
        await g.prefetchRange?.(sheet, row, col, row, col);
        const c = g.getCell(sheet, row, col);
        if (!c) return null;
        const out: CellValue = { value: c.value, ...(c.display !== undefined ? { display: c.display } : {}), formula: c.formula };
        if (c.formula) {
          const index = g.getSheetNames().indexOf(sheet);
          const kept = index >= 0 ? await engineRef.current?.fileSavedAt(index, row + 1, col + 1).catch(() => null) : null;
          if (kept) {
            out.fromFile = true;
            out.note = calcNote(kept);
          }
        }
        return out;
      },
      getCalcIssues: async (limit = 1000) => {
        grid();
        await calculated();
        return (await engineRef.current!.fileSavedCells(limit)).map(toIssue);
      },
      setCell: (sheet, ref, value) => {
        const { row, col } = parseRef(ref);
        grid().setCells(sheet, [{ row, col, value }]);
      },
      setCells: (sheet, cells) => void grid().setCells(sheet, cells),
      whenCalculated: async (timeoutMs) => {
        grid();
        await calculated(timeoutMs);
      },
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
      hasUnsavedChanges: () => dirty() || pendingCount() > 0,
      getDocument: () => ({ id: docRef.current?.id ?? "", name: docRef.current?.name ?? "" }),
      open: async (file, options) => {
        const bytes = new Uint8Array(file instanceof Uint8Array ? file : new Uint8Array(file));
        await openBytesRef.current(bytes, options?.name ?? "workbook.xlsx", { id: options?.id, discardUnsaved: options?.discardUnsaved });
      },
      openStored: (id, options) => openStoredRef.current(id, options),
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
  const openStoredRef = useRef(openStored);
  openStoredRef.current = openStored;
  unmountSave.current = () => (canSave && loadedRef.current && dirty() && !pendingCount() ? saveNow() : null);

  // --- loading the current workbook into the grid ----------------------------
  const tabId = doc ? `gridpath-${instance}-${doc.key}` : null;
  useEffect(() => {
    const grid = gridRef.current;
    const d = docRef.current;
    if (!engine || !grid || !d || !tabId) return;
    let live = true;
    /** The load is over: tell whoever opened it (an `open()` call), or show it. */
    const finish = (err?: GridPathError) => {
      const waiter = loadWaiters.current.get(d.key);
      loadWaiters.current.delete(d.key);
      if (!err) return waiter?.resolve();
      if (waiter) return waiter.reject(err);
      setLoadError(err);
      report(err);
    };
    (async () => {
      const p = propsRef.current;
      // CSV and other delimited text open as a new .xlsx made from it.
      const fromText = !!d.bytes && isDelimitedText(d.bytes, d.name);
      let base: Uint8Array;
      try {
        base = !d.bytes ? await engine.blankXlsx(p.sheets ?? ["Sheet1"]) : fromText ? await engine.csvToXlsx(d.bytes, d.name) : d.bytes;
      } catch (err) {
        throw fromText && engine.alive() ? textTableError(d.name, err) : engineError(err instanceof Error ? err.message : String(err), err);
      }
      const res = await grid.loadBytes(base, (s) => propsRef.current.onProgress?.(s), { tabId, path: d.name });
      if (!live) return finish();
      if (!res.engineLoaded) {
        const reason = engine.openFailureReason();
        throw !engine.alive() || reason === null ? engineError(engine.degradedReason() ?? "the workbook could not be opened") : openError(base, d.name, reason);
      }
      const ai = new AgentToolkit(grid, engine);
      ai.onChange(batchesChanged);
      const initialFreeze = new Map<string, { rows: number; cols: number }>();
      for (const name of grid.getSheetNames()) {
        const f = grid.getFreezePanes(name);
        initialFreeze.set(name, { rows: f.freezeRows, cols: f.freezeCols });
      }
      cur.current = { key: d.key, grid, ai, base, sheetOps: [], initialFreeze };
      lastBytes.current = base;
      hasDraft.current = !!d.draft;
      attempts.current = 0;
      setBatches([]);
      setIssues([]);
      pendingBefore.current = 0;
      changes.current = 0;
      saved.current = 0;
      setHistory({ canUndo: false, canRedo: false });
      setLoaded(true);
      loadedRef.current = true;
      setLoadError(null);
      setEngineDown(null);
      // Back where the user was.
      const view = d.savedAt || d.draft ? readView(d.id) : null;
      if (view && grid.getSheetNames().includes(view.sheet)) grid.jumpToCell(view.sheet, view.row, view.col);
      if (pendingCheckpoint.current) {
        // A restored version: it becomes the current state, as a version of its own.
        pendingCheckpoint.current = false;
        changes.current = 1;
        await saveNow({ checkpoint: true });
      } else if (d.draft) {
        // Changes the store never got: they go there now.
        changes.current = 1;
        setStatus({ state: "unsaved" });
        notify(`Restored changes that weren't saved${store?.label ? ` ${store.label}` : ""}`, 4000);
        schedule();
      } else {
        setStatus(d.savedAt ? { state: "saved", savedAt: d.savedAt, ...(store?.label ? { where: store.label } : {}) } : { state: "idle" });
        if (fromText) notify(`Opened as a new workbook · saves as ${d.name}`, 4000);
      }
      refreshIssues();
      if (!readyCalled.current) {
        readyCalled.current = true;
        propsRef.current.onReady?.(apiRef.current!);
      } else if (!restoring.current) {
        propsRef.current.onDocumentChange?.({ id: d.id, name: d.name });
      }
      finish();
    })().catch((err) => {
      if (!live) return finish();
      finish(asGridPathError(err, (x) => engineError(x instanceof Error ? x.message : String(x), x)));
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
      for (const t of [timer, retryTimer, noticeTimer, issuesTimer]) if (t.current) clearTimeout(t.current);
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
  const ready = loaded && !loadError;

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
          disabled={!ready}
          fileDisabled={!engine}
          canUndo={history.canUndo}
          canRedo={history.canRedo}
          status={canSave ? status : status.state === "review" ? status : dirty() ? { state: "unsaved" } : { state: "idle" }}
          canSave={canSave}
          onSave={() => void saveNow()}
          onDownload={() => void downloadNow()}
          fileMenu={props.fileMenu !== false}
          onOpenFile={(f) => uiOpen((discardUnsaved) => f.arrayBuffer().then((b) => openBytes(new Uint8Array(b), f.name, { discardUnsaved })))}
          onNew={() =>
            uiOpen((discardUnsaved) =>
              openDoc({ id: `untitled-${Date.now().toString(36)}`, name: "Untitled.xlsx", bytes: null, source: "blank", savedAt: null }, { discardUnsaved }),
            )
          }
          listRecent={store ? () => store.list() : null}
          onOpenRecent={(id) => uiOpen((discardUnsaved) => openStored(id, { discardUnsaved }))}
          currentId={doc?.id ?? ""}
          onShowHistory={store ? () => void showHistory() : null}
          onReview={props.reviewPanel === false ? null : () => setPanel("review")}
          issueCount={issues.length}
          onShowIssues={() => setPanel("issues")}
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
        {engineDown && !loadError && <EngineBanner error={engineDown} onDownload={() => void downloadNow()} />}
        {loadError && <LoadError error={loadError} onRetry={loadError.code === "storage-failed" && !doc ? retryFirst : null} />}
        {panel === "review" && doc && (
          <ReviewPanel
            batches={batches}
            onAccept={(id) => cur.current.ai?.accept(id)}
            onReject={(id) => void cur.current.ai?.reject(id)}
            onJump={(c) => cur.current.grid?.jumpToCell(c.sheet, c.row, c.col)}
            onClose={() => setPanel(null)}
          />
        )}
        {panel === "history" && doc && (
          <HistoryPanel
            name={doc.name}
            where={store?.label}
            versions={versions}
            onRestore={(id) => void restoreVersion(id)}
            onDownload={(id) => void downloadVersion(id)}
            onClose={() => setPanel(null)}
          />
        )}
        {panel === "issues" && doc && (
          <IssuesPanel
            issues={issues}
            max={MAX_ISSUES}
            onJump={(c) => cur.current.grid?.jumpToCell(c.sheet, c.row, c.col)}
            onClose={() => setPanel(null)}
          />
        )}
        {confirm && <ConfirmDialog {...confirm} onCancel={() => setConfirm(null)} />}
      </div>
    </div>
  );
}

/**
 * Mount a spreadsheet into `el` (which should have a height) and resolve
 * once the workbook is open. If it can't open (the engine didn't start, the
 * file isn't a readable .xlsx, the store can't be reached), the promise
 * rejects with a `GridPathError` and `el` is left empty. Works without
 * React in your app; React and React DOM are peer dependencies of the package.
 */
export function createSpreadsheet(el: HTMLElement, options: SpreadsheetOptions = {}): Promise<Spreadsheet> {
  const root = createRoot(el);
  let ready = false;
  return new Promise((resolve, reject) => {
    const destroy = () => root.unmount();
    root.render(
      <Spreadsheet
        {...options}
        onReady={(api) => {
          ready = true;
          api.destroy = destroy;
          resolve(api);
        }}
        onError={(err) => {
          options.onError?.(err);
          if (ready) return;
          destroy();
          reject(err);
        }}
      />,
    );
  });
}
