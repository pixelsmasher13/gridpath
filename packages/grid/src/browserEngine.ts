/**
 * GridEngine over the WASM engine, for the browser. The engine runs in a
 * Web Worker by default (recalculation never freezes the page) or on the
 * page itself (`worker: false`). Ordering follows the desktop bridge
 * (ironcalcEngine.ts): one serial queue, consecutive cell writes coalesced
 * into one engine batch, recomputed values pushed to the grid's applier.
 */
import type { DeltaApplier, FindHit, GridEngine } from "./engine";
import type { CellDelta, HostRequest, HostResponse } from "./engineHost";
import type { NativeChart, EditTimings } from "./charts/types";

export interface EngineOptions {
  /** Run the engine in a Web Worker (default true). */
  worker?: boolean;
  /** Where to load the engine's .wasm from, when the bundler does not
   * serve it next to the package's script (CDNs, strict asset pipelines). */
  wasmUrl?: string | URL;
  /** Where to load the worker script from, likewise. */
  workerUrl?: string | URL;
}

type Call = (cmd: string, args?: any, transfer?: Transferable[]) => Promise<any>;

/** A formula showing the file's saved value. `why`: "function" calls an
 * unknown function (`name`), "link" refers to another workbook, "syntax"
 * couldn't be parsed (these three keep the value as a constant), or
 * "error": it calculated to an error or nothing (`live`) where the file
 * had a value. 0-based. */
export type FileSavedCell = {
  sheet: string;
  row: number;
  col: number;
  formula: string;
  shown: unknown;
  live: unknown;
  why: "function" | "link" | "syntax" | "error";
  name?: string;
};
type Channel = { call: Call; close: () => void; onFatal: (fn: (err: Error) => void) => void };

/** A request/response channel to the engine, in a worker or in-page. */
async function connect(opts: EngineOptions): Promise<Channel> {
  const wasmUrl = opts.wasmUrl ? String(opts.wasmUrl) : undefined;
  if (opts.worker === false) {
    const { EngineHost } = await import("./engineHost");
    const host = new EngineHost();
    // Same one-at-a-time ordering the worker gives.
    let chain: Promise<unknown> = Promise.resolve();
    const call: Call = (cmd, args) => {
      const run = chain.then(() => host.handle(cmd, args));
      chain = run.catch(() => {});
      return run;
    };
    await call("init", { wasmUrl });
    return { call, close: () => void call("unload").catch(() => {}), onFatal: () => {} };
  }
  const worker = opts.workerUrl
    ? new Worker(opts.workerUrl, { type: "module" })
    : new Worker(new URL("./engine.worker.js", import.meta.url), { type: "module" });
  let next = 1;
  // Once the worker dies every call fails at once instead of waiting forever.
  let dead: Error | null = null;
  let fatal: (err: Error) => void = () => {};
  const waiting = new Map<number, { resolve: (v: unknown) => void; reject: (e: Error) => void }>();
  worker.onmessage = (ev: MessageEvent<HostResponse>) => {
    const w = waiting.get(ev.data.id);
    if (!w) return;
    waiting.delete(ev.data.id);
    if (ev.data.ok) w.resolve(ev.data.result);
    else w.reject(new Error(ev.data.error));
  };
  worker.onerror = (ev) => {
    const err = new Error(
      `engine worker failed: ${ev.message || "could not load the worker script"}. ` +
        `With Vite's dev server, add optimizeDeps: { exclude: ["gridpath"] } to vite.config; ` +
        `elsewhere, serve the package's engine.worker.js and pass engine: { workerUrl }.`,
    );
    dead = err;
    for (const w of waiting.values()) w.reject(err);
    waiting.clear();
    fatal(err);
  };
  const call: Call = (cmd, args, transfer) =>
    new Promise((resolve, reject) => {
      if (dead) return reject(dead);
      const id = next++;
      waiting.set(id, { resolve, reject });
      const msg: HostRequest = { id, cmd, args };
      worker.postMessage(msg, transfer ?? []);
    });
  await call("init", { wasmUrl });
  return { call, close: () => worker.terminate(), onFatal: (fn) => (fatal = fn) };
}

type Edit = { row: number; column: number; value: string };

/** Edits waiting for their queued task; open until that task starts or
 * another engine task queues behind them (which seals it, so ordering
 * against sheet and structural ops holds). */
interface EditSink {
  bySheet: Map<string, Edit[]>;
  cells: number;
  sealed: boolean;
}

const MAX_EDIT_CELLS = 50_000;

export class BrowserEngine implements GridEngine {
  private chartData: NativeChart[] = [];
  private chartListeners = new Set<() => void>();
  private editTiming: EditTimings | null = null;
  charts = (): NativeChart[] => this.chartData;
  lastEditTimings = (): EditTimings | null => this.editTiming && { ...this.editTiming };
  subscribeCharts = (listener: () => void): (() => void) => {
    this.chartListeners.add(listener);
    return () => { this.chartListeners.delete(listener); };
  };
  private acceptResult(res: { cells?: CellDelta[]; charts?: NativeChart[] }): void {
    if (res.cells) this.applyDeltas(res.cells);
    if (res.charts) {
      this.chartData = res.charts;
      for (const listener of this.chartListeners) listener();
    }
  }
  private sheets: string[] = [];
  private ready = false;
  private degraded: string | null = null;
  private queue: Promise<void> = Promise.resolve();
  private sink: EditSink | null = null;
  private applier: DeltaApplier | null = null;
  private pending: CellDelta[] = [];
  private settleStarted = false;
  private settleTimer: ReturnType<typeof setTimeout> | null = null;
  private openFailure: string | null = null;
  private dead = false;
  private failureListeners = new Set<(reason: string) => void>();

  private constructor(private readonly rpc: Channel) {
    rpc.onFatal((err) => {
      this.dead = true;
      this.degrade(err.message);
    });
  }

  static async create(opts: EngineOptions = {}): Promise<BrowserEngine> {
    return new BrowserEngine(await connect(opts));
  }

  /** Called when the engine stops mirroring edits (values on screen freeze).
   * Not called for a file that fails to open: see `openFailureReason`. */
  onFailure(fn: (reason: string) => void): () => void {
    this.failureListeners.add(fn);
    return () => this.failureListeners.delete(fn);
  }

  /** False once the worker has died: every call fails from then on. */
  alive(): boolean {
    return !this.dead;
  }

  /** Why the last `openLayout` failed (the engine's words), or null. */
  openFailureReason(): string | null {
    return this.openFailure;
  }

  /** An empty .xlsx with these sheets: the base a new workbook opens from. */
  blankXlsx(sheets: string[]): Promise<Uint8Array> {
    return this.rpc.call("blank", { sheets });
  }

  /** Delimited text (CSV, TSV…) as a new .xlsx with one sheet named after
   * `name`: the base a workbook opened from text opens from. */
  csvToXlsx(bytes: Uint8Array, name: string, delimiter?: string): Promise<Uint8Array> {
    return this.rpc.call("csv", { bytes, name, delimiter });
  }

  /** The original file with every session change patched in; untouched
   * package parts come back byte-identical. */
  patch(base: Uint8Array, patchJson: string): Promise<Uint8Array> {
    return this.rpc.call("patch", { base, patch: patchJson });
  }

  // ---- reads for the agent tools: queued behind every pending write ----

  /** One cell, 1-based: `{content, value, display, file_saved?}`. */
  cellAt(sheet: number, row: number, col: number): Promise<{ content: string; value: unknown; display: string }> {
    return this.read(async () => JSON.parse(await this.rpc.call("cell", { sheet, row, col })));
  }

  /** Non-empty cells of a 1-based inclusive rectangle. */
  readRect(sheet: number, r0: number, r1: number, c0: number, c1: number): Promise<any> {
    return this.read(async () => JSON.parse(await this.rpc.call("window", { sheet, r0, r1, c0, c1 })));
  }

  /** The engine-built structural index (same shape as buildWorkbookIndex). */
  workbookIndex(): Promise<any> {
    return this.read(async () => JSON.parse(await this.rpc.call("workbookIndex")));
  }

  /** query_rows against a sheet index. */
  queryRows(query: Record<string, unknown>): Promise<any> {
    return this.read(async () => JSON.parse(await this.rpc.call("queryRows", { query })));
  }

  /** Formulas showing the value Excel saved instead of a calculated one,
   * and why (see `FileSavedCell`), 0-based. */
  fileSavedCells(limit = 1000): Promise<FileSavedCell[]> {
    return this.read(() => this.rpc.call("fileSaved", { limit }));
  }

  /** One cell's `fileSavedCells` entry (1-based), or null when it calculates. */
  fileSavedAt(sheet: number, row: number, col: number): Promise<FileSavedCell | null> {
    return this.read(() => this.rpc.call("fileSavedAt", { sheet, row, col }));
  }

  /** True when every queued write and recalculation finished within `ms`. */
  async settledWithin(ms: number): Promise<boolean> {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timedOut = new Promise<false>((resolve) => (timer = setTimeout(() => resolve(false), ms)));
    try {
      return await Promise.race([this.queue.then(() => true as const), timedOut]);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  /** Formula cells this session's edits changed, with their current values,
   * for the save patch's cached results (0-based, live sheet names). */
  recalculatedCells(): Promise<Array<{ sheet: string; r: number; c: number; value: any }>> {
    return this.read(() => this.rpc.call("recalculated"));
  }

  /** One sheet's cells as engine JSON bytes (run_script's read model). */
  sheetCellsBuffer(sheet: number): Promise<ArrayBuffer> {
    return this.read(async () => {
      const bytes: Uint8Array = await this.rpc.call("sheetCells", { sheet });
      return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    });
  }

  private read<T>(fn: () => Promise<T>): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      this.enqueue(async () => {
        try {
          resolve(await fn());
        } catch (e) {
          reject(e);
        }
      });
    });
  }

  /** Why the engine stopped mirroring edits, or null while healthy. */
  degradedReason(): string | null {
    return this.degraded;
  }

  close(): void {
    if (this.settleTimer) clearTimeout(this.settleTimer);
    this.applier = null;
    this.rpc.close();
  }

  // ---- GridEngine ----

  openLayout(_path: string, bytes: Uint8Array): Promise<any | null> {
    // A new workbook (the engine outlives the documents opened in it): its
    // deferred first evaluation hasn't run.
    this.ready = false;
    this.settleStarted = false;
    this.openFailure = null;
    // A failure belonged to the previous document, unless the worker died.
    if (!this.dead) this.degraded = null;
    if (this.settleTimer) clearTimeout(this.settleTimer);
    return new Promise((resolve) => {
      this.enqueue(async () => {
        try {
          // A copy crosses to the engine; the grid keeps its bytes for saving.
          const res = await this.rpc.call("load", { bytes: bytes.slice() });
          this.sheets = res.sheets;
          this.acceptResult(res);
          const layout = JSON.parse(await this.rpc.call("layout"));
          this.ready = true;
          resolve(layout);
          // The grid starts the deferred evaluate once its first tiles are
          // on screen; this is the backstop.
          this.settleTimer = setTimeout(() => this.startDeferredSettle(), 10_000);
        } catch (e) {
          this.openFailure = e instanceof Error ? e.message : String(e);
          this.degraded = `open failed: ${this.openFailure}`;
          resolve(null);
        }
      });
    });
  }

  startDeferredSettle(): void {
    if (this.settleStarted || !this.ready) return;
    this.settleStarted = true;
    if (this.settleTimer) clearTimeout(this.settleTimer);
    this.enqueue(async () => {
      if (this.degraded) return;
      try {
        const res = await this.rpc.call("evaluatePending");
        this.acceptResult(res);
      } catch (e) {
        this.degrade(`first evaluate failed: ${e}`);
      }
    });
  }

  fetchLayout(): Promise<any | null> {
    if (this.degraded) return Promise.resolve(null);
    return new Promise((resolve) => {
      this.enqueue(async () => {
        try {
          resolve(JSON.parse(await this.rpc.call("layout")));
        } catch (e) {
          console.error("[gridpath] layout fetch failed:", e);
          resolve(null);
        }
      });
    });
  }

  async fetchWindow(sheet: number, r0: number, r1: number, c0: number, c1: number): Promise<any | null> {
    if (!this.ready || this.degraded) return null;
    try {
      return JSON.parse(await this.rpc.call("window", { sheet, r0, r1, c0, c1 }));
    } catch (e) {
      console.error("[gridpath] window fetch failed:", e);
      return null;
    }
  }

  async calcHealth(): Promise<any | null> {
    return null;
  }

  setStyle(sheetName: string, r0: number, c0: number, r1: number, c1: number, patch: Record<string, unknown>): Promise<boolean> {
    if (this.degraded) return Promise.resolve(false);
    return new Promise((resolve) => {
      this.enqueue(async () => {
        const sheet = this.sheets.indexOf(sheetName);
        if (sheet < 0) return resolve(false);
        try {
          await this.rpc.call("setStyle", { sheet, r0: r0 + 1, r1: r1 + 1, c0: c0 + 1, c1: c1 + 1, patch });
          resolve(true);
        } catch (e) {
          console.warn("[gridpath] set style failed:", e);
          resolve(false);
        }
      });
    });
  }

  defineName(name: string, formula: string): Promise<boolean> {
    if (this.degraded) return Promise.resolve(false);
    this.startDeferredSettle();
    return new Promise((resolve) => {
      this.enqueue(async () => {
        try {
          this.acceptResult(await this.rpc.call("defineName", { name, formula }));
          resolve(true);
        } catch (e) {
          console.warn(`[gridpath] define name ${name} failed:`, e);
          resolve(false);
        }
      });
    });
  }

  async find(query: string, sheet?: number, limit = 500): Promise<FindHit[]> {
    if (!this.ready || this.degraded) return [];
    try {
      return await this.rpc.call("find", { query, sheet: sheet ?? null, limit });
    } catch (e) {
      console.warn("[gridpath] find failed:", e);
      return [];
    }
  }

  sheetNames(): string[] {
    return this.sheets.slice();
  }

  writeCells(sheetName: string, _label: string, cells: Edit[]): void {
    if (cells.length === 0) return;
    // Writes apply to an evaluated workbook: the deferred first evaluation
    // goes ahead of them in the queue (or edits made right after opening are
    // recalculated against the file's cached values and lost on save).
    this.startDeferredSettle();
    const open = this.sink;
    if (open && !open.sealed && open.cells + cells.length <= MAX_EDIT_CELLS) {
      let list = open.bySheet.get(sheetName);
      if (!list) open.bySheet.set(sheetName, (list = []));
      list.push(...cells);
      open.cells += cells.length;
      return;
    }
    const sink: EditSink = { bySheet: new Map([[sheetName, [...cells]]]), cells: cells.length, sealed: false };
    this.sink = sink;
    this.enqueueRaw(async () => {
      sink.sealed = true;
      if (this.sink === sink) this.sink = null;
      for (const [name, cs] of sink.bySheet) await this.editBatch(name, cs);
    });
  }

  sheetOp(op: "add" | "rename" | "delete", name: string, newName?: string): void {
    this.startDeferredSettle();
    this.enqueue(() => this.runSheetOp(op, name, newName));
  }

  structural(op: string, sheetName: string, index1: number, count: number): void {
    this.startDeferredSettle();
    this.enqueue(async () => {
      if (this.degraded) return;
      const sheet = this.sheets.indexOf(sheetName);
      if (sheet < 0) return this.degrade(`${op}: unknown sheet "${sheetName}"`);
      try {
        this.acceptResult(await this.rpc.call("structural", { op, sheet, index: index1, count }));
      } catch (e) {
        this.degrade(`${op} failed: ${e}`);
      }
    });
  }

  settled(timeoutMs = 30_000): Promise<void> {
    return Promise.race([this.queue, new Promise<void>((resolve) => setTimeout(resolve, timeoutMs))]);
  }

  registerApplier(fn: DeltaApplier): void {
    this.applier = fn;
    if (this.pending.length > 0) {
      const pending = this.pending;
      this.pending = [];
      this.applyDeltas(pending);
    }
  }

  unregisterApplier(): void {
    this.applier = null;
  }

  // ---- internals ----

  private enqueue(work: () => Promise<void>): void {
    if (this.sink) this.sink.sealed = true;
    this.enqueueRaw(work);
  }

  private enqueueRaw(work: () => Promise<void>): void {
    this.queue = this.queue.then(work).catch((e) => console.error("[gridpath] engine task failed:", e));
  }

  private async editBatch(sheetName: string, cells: Edit[]): Promise<void> {
    if (this.degraded) return;
    let sheet = this.sheets.indexOf(sheetName);
    if (sheet < 0) {
      // A write names a sheet the grid has, so the sheet exists and only
      // its add has not reached the engine yet.
      await this.runSheetOp("add", sheetName);
      sheet = this.sheets.indexOf(sheetName);
      if (sheet < 0) return this.degrade(`edit for unknown sheet "${sheetName}"`);
    }
    try {
      const edits = cells.map((c) => ({ sheet, row: c.row, column: c.column, value: c.value }));
      const start = performance.now();
      const res = await this.rpc.call("editBatch", { edits });
      const returned = performance.now();
      this.acceptResult(res);
      this.editTiming = { ...res.timing, roundTripMs: returned-start, applyMs: performance.now()-returned };
    } catch (e) {
      this.degrade(`edit failed: ${e}`);
    }
  }

  private async runSheetOp(op: string, name: string, newName?: string): Promise<void> {
    if (this.degraded) return;
    // Idempotent add: the unknown-sheet fallback and an explicit add can
    // both ask for the same sheet.
    if (op === "add" && this.sheets.includes(name)) return;
    try {
      const res = await this.rpc.call("sheetOp", { op, name, newName: newName ?? null });
      this.sheets = res.sheets;
      this.acceptResult(res);
    } catch (e) {
      this.degrade(`sheet ${op} failed: ${e}`);
    }
  }

  private applyDeltas(deltas: CellDelta[]): void {
    if (deltas.length === 0) return;
    if (!this.applier) {
      this.pending.push(...deltas);
      return;
    }
    const bySheet = new Map<number, { row0: number; col0: number; v: unknown }[]>();
    for (const d of deltas) {
      let list = bySheet.get(d.sheet);
      if (!list) bySheet.set(d.sheet, (list = []));
      list.push({ row0: d.row - 1, col0: d.column - 1, v: d.value });
    }
    for (const [sheetIndex, cells] of bySheet) {
      const name = this.sheets[sheetIndex];
      if (!name) continue;
      if (!this.applier(name, cells)) {
        this.pending.push(...cells.map((c) => ({ sheet: sheetIndex, row: c.row0 + 1, column: c.col0 + 1, value: c.v })));
      }
    }
  }

  /** There is no grid dump to rebuild the model from (the desktop's
   * resync); values freeze and the reason is kept for `degradedReason`. */
  private degrade(reason: string): void {
    if (this.degraded) return;
    this.degraded = reason;
    console.error(`[gridpath] engine stopped: ${reason}`);
    for (const fn of this.failureListeners) fn(reason);
  }
}
