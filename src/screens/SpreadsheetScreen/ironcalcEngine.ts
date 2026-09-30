/**
 * IronCalc ENGINE mode (enable with
 * `localStorage.setItem("gridpath.ironcalcEngine", "1")` + reload).
 *
 * Unlike shadow mode, IronCalc is the calculator here: Univer's formula
 * execution is disabled (the calc worker starts with `notExecuteFormula`, see
 * UniverGrid), one IronCalc model is held per tab in the Tauri process, and
 * the values it computes are written into the grid.
 *
 * Mirroring has two paths:
 *
 * 1. COMMANDS (`sheet.command.set-range-values`, insert/remove row/col,
 *    clear) — normal user typing, agent writes, structural edits.
 * 2. UNDO/REDO replays dispatch no edit commands — they re-execute inverse
 *    MUTATIONS directly. Those mutations arrive in the stream BEFORE their
 *    `univer.command.undo|redo` parent (children complete first), so
 *    unattributed mutations are buffered for one tick: if an undo/redo
 *    parent claims them they are flushed into the engine; if a normal edit
 *    command claims them they are dropped (the command path already mirrors
 *    it); if nothing claims them (the app's own load/style population — a
 *    56k-cell write on open!) they evaporate. Our own value write-backs are
 *    tagged `__ironcalc: true` and never buffered.
 *
 * Engine work per tab runs through a serial queue — including the initial
 * load, so edits made while the model is still importing queue behind it
 * instead of failing. Values land as one raw `set-range-values` mutation per
 * sheet (merges `v`, keeps formulas/styles, no undo entries).
 *
 * Operations the bridge cannot mirror (sort, range moves, merges,
 * shared-formula `si` fills, unrecognized commands that touch content)
 * trigger a RESYNC: the model is reseeded in place from a full grid dump
 * (formulas keep their grid value as the cached value, so frozen add-in /
 * external-ref data survives; defined names live in the model and are kept).
 * Unrecognized commands whose child mutations are plain decodable writes
 * (drag-fill, simple paste) are mirrored directly from those mutations —
 * resync is the fallback, not the first resort. Only unrecoverable failures
 * (load errors, resync itself failing) still DEGRADE: values freeze +
 * console error; reopening the file recovers. All output is on the
 * `[ironcalc-engine]` console prefix.
 */
import { invoke } from "@tauri-apps/api/core";

import { extractEditedCells } from "./ironcalcShadow";

interface CellDelta {
  sheet: number;
  row: number; // 1-based (IronCalc)
  column: number; // 1-based
  value: unknown;
}

interface CalcEditResult {
  recalc_ms: number;
  changed: CellDelta[];
}

interface CalcSnapshot {
  ms: number;
  cells: CellDelta[];
}

interface CalcLoadResult {
  sheets: string[];
  load_ms: number;
  eval_ms: number;
}

/** Writes computed values into the visible grid; registered by UniverGrid. */
export type DeltaApplier = (
  sheetName: string,
  cells: { row0: number; col0: number; v: unknown }[],
) => boolean;

/** One cell of a grid content dump (resync seed). Row/column are 1-based. */
export interface GridResyncCell {
  row: number;
  column: number;
  /** Formula text ("=SUM(..)"); when set, `value` is the grid's current
   * computed value — it becomes the engine's cached value, which is what
   * keeps unevaluable add-in/external-ref formulas frozen at real data. */
  formula: string | null;
  value: unknown;
}
export interface GridResyncSheet {
  name: string;
  cells: GridResyncCell[];
}
/** Reads the grid's full content for a resync; registered by UniverGrid
 * alongside the applier. Returns null when the grid cannot be read. */
export type GridSnapshotter = () => GridResyncSheet[] | null;

interface TabState {
  path: string;
  sheets: string[];
  ready: boolean;
  degraded: string | null;
  /** Serializes engine work for this tab (load first, then edits in order):
   * async invokes must apply in mutation order or a fast second edit could
   * be overwritten by a slow first one. */
  queue: Promise<void>;
  /** Deltas that arrived before this tab's grid registered its applier;
   * flushed on registration. */
  pending: CellDelta[];
  /** Unattributed mutations from THIS tab's command stream, held one tick
   * for parent attribution (undo/redo vs normal command vs load noise).
   * Per-tab: all grids are mounted simultaneously and their streams can
   * interleave in time. */
  buffer: BufferedMutation[];
  bufferDropScheduled: boolean;
  /** A resync task is queued but has not run yet. Commands arriving in this
   * window are NOT mirrored — the resync snapshot (read when the task runs)
   * already contains their grid effects. */
  resyncQueued: boolean;
}

const appliers = new Map<string, DeltaApplier>();
const snapshotters = new Map<string, GridSnapshotter>();
let activeTabId: string | null = null;
const tabs = new Map<string, TabState>();

export function ironcalcEngineEnabled(): boolean {
  // IronCalc IS the calculation engine — on by default. What remains is an
  // internal kill switch for debugging and Univer-comparison during the
  // rollout period: localStorage.setItem("gridpath.ironcalcEngine", "0")
  // (or VITE_IRONCALC_ENGINE=0) reverts to the Univer engine. Delete the
  // switch entirely once the engine has fully soaked.
  try {
    if (localStorage.getItem("gridpath.ironcalcEngine") === "0") return false;
  } catch {
    /* no localStorage (tests) — default applies */
  }
  try {
    if ((import.meta as any).env?.VITE_IRONCALC_ENGINE === "0") return false;
  } catch {
    /* default applies */
  }
  return true;
}

/** Each mounted grid registers its own applier (closing over its own
 * Univer instance): every tab's grid stays live in the background, so
 * deltas apply immediately regardless of which tab is visible. */
export function registerIroncalcApplier(tabId: string, fn: DeltaApplier): void {
  appliers.set(tabId, fn);
  const st = tabs.get(tabId);
  if (st && st.pending.length > 0) {
    const pending = st.pending;
    st.pending = [];
    applyDeltas(tabId, pending);
  }
}

export function unregisterIroncalcApplier(tabId: string): void {
  appliers.delete(tabId);
}

export function registerIroncalcSnapshotter(tabId: string, fn: GridSnapshotter): void {
  snapshotters.set(tabId, fn);
}

export function unregisterIroncalcSnapshotter(tabId: string): void {
  snapshotters.delete(tabId);
}

export function ironcalcEngineSetActiveTab(tabId: string): void {
  if (!ironcalcEngineEnabled()) return;
  activeTabId = tabId;
  const st = tabs.get(tabId);
  if (st && st.pending.length > 0) {
    const pending = st.pending;
    st.pending = [];
    applyDeltas(tabId, pending);
  }
}

function newTabState(path: string): TabState {
  return {
    path,
    sheets: [],
    ready: false,
    degraded: null,
    queue: Promise.resolve(),
    pending: [],
    buffer: [],
    bufferDropScheduled: false,
    resyncQueued: false,
  };
}

export async function ironcalcEngineOnOpen(tabId: string, path: string): Promise<void> {
  if (!ironcalcEngineEnabled()) return;
  if (!path || path.startsWith("untitled-") || !/\.xls[xm]$/i.test(path)) return;
  activeTabId = tabId;
  tabs.set(tabId, newTabState(path));
  // The load is the queue's first item: edits arriving during the multi-
  // second import run after it instead of hitting "no workbook loaded".
  enqueue(tabId, async () => {
    try {
      const t0 = performance.now();
      const res = await invoke<CalcLoadResult>("calc_engine_load", { tab: tabId, path });
      const st = tabs.get(tabId);
      if (!st) return; // tab closed while loading
      st.sheets = res.sheets;
      const snap = await invoke<CalcSnapshot>("calc_engine_snapshot", { tab: tabId });
      applyDeltas(tabId, snap.cells);
      st.ready = true;
      console.log(
        `[ironcalc-engine] ready: ${path} — import ${res.load_ms}ms, calc ${res.eval_ms}ms, ` +
          `${snap.cells.length} formula values applied (total ${(performance.now() - t0).toFixed(0)}ms)`,
      );
    } catch (e) {
      console.error(
        "[ironcalc-engine] load failed — grid shows only the file's cached values:",
        e,
      );
      const st = tabs.get(tabId);
      if (st) st.degraded = String(e);
    }
  });
}

/** Registers an EMPTY model for a fresh untitled workbook (the LLM-builds-
 * a-model flow): there is no file to import, but edits must still mirror
 * into an engine or nothing ever calculates. */
export function ironcalcEngineOnBlank(tabId: string, sheetNames: string[]): void {
  if (!ironcalcEngineEnabled()) return;
  activeTabId = tabId;
  tabs.set(tabId, newTabState("(untitled)"));
  enqueue(tabId, async () => {
    try {
      const res = await invoke<CalcLoadResult>("calc_engine_new", {
        tab: tabId,
        sheets: sheetNames,
      });
      const st = tabs.get(tabId);
      if (!st) return;
      st.sheets = res.sheets;
      st.ready = true;
      console.log(`[ironcalc-engine] ready: blank workbook (${res.sheets.join(", ")})`);
    } catch (e) {
      console.error("[ironcalc-engine] blank-model init failed:", e);
      const st = tabs.get(tabId);
      if (st) st.degraded = String(e);
    }
  });
}

/**
 * Registers a model for a RESTORED untitled draft, seeded from the Univer
 * snapshot the session was resumed from (there is no file to import).
 * Reuses calc_engine_new + one edit batch: the batch's returned deltas also
 * give the restored formulas their computed values in the grid.
 */
export function ironcalcEngineOnRestore(tabId: string, snapshot: unknown): void {
  if (!ironcalcEngineEnabled()) return;
  const snap = snapshot as {
    sheetOrder?: string[];
    sheets?: Record<string, { name?: string; cellData?: Record<string, Record<string, unknown>> }>;
  };
  const order = snap?.sheetOrder ?? Object.keys(snap?.sheets ?? {});
  const sheetNames: string[] = [];
  const edits: EditCellWire[] = [];
  let failed: string | null = null;
  for (const sheetId of order) {
    const sheet = snap?.sheets?.[sheetId];
    if (!sheet || typeof sheet.name !== "string") continue;
    const sheetIndex = sheetNames.length;
    sheetNames.push(sheet.name);
    for (const [r, row] of Object.entries(sheet.cellData ?? {})) {
      if (!row || typeof row !== "object") continue;
      for (const [c, cell] of Object.entries(row as Record<string, unknown>)) {
        const input = mutationCellToInput(cell);
        if (input === "si") {
          failed = `restored draft has shared-formula refs (${sheet.name})`;
          break;
        }
        if (input === null || input === "") continue; // style-only / empty
        edits.push({ sheet: sheetIndex, row: Number(r) + 1, column: Number(c) + 1, value: input });
      }
    }
  }
  if (sheetNames.length === 0) sheetNames.push("Sheet1");
  activeTabId = tabId;
  tabs.set(tabId, newTabState("(restored draft)"));
  if (failed) {
    degrade(tabId, failed);
    return;
  }
  enqueue(tabId, async () => {
    try {
      const t0 = performance.now();
      const res = await invoke<CalcLoadResult>("calc_engine_new", {
        tab: tabId,
        sheets: sheetNames,
      });
      const st = tabs.get(tabId);
      if (!st) return;
      st.sheets = res.sheets;
      st.ready = true;
      if (edits.length > 0) {
        await engineEditBatch(tabId, "restored draft seed", edits);
      }
      console.log(
        `[ironcalc-engine] ready: restored draft — ${sheetNames.length} sheet(s), ` +
          `${edits.length} cells seeded (${(performance.now() - t0).toFixed(0)}ms)`,
      );
    } catch (e) {
      console.error("[ironcalc-engine] restored-draft init failed:", e);
      const st = tabs.get(tabId);
      if (st) st.degraded = String(e);
    }
  });
}

/**
 * Resolves when the active tab's engine queue is drained (all mirrored
 * edits recalculated and their deltas applied). The engine-mode replacement
 * for waiting on Univer's calculationEnd, which never fires.
 */
export function ironcalcEngineSettled(tabId?: string, timeoutMs = 30_000): Promise<void> {
  const target = tabId ?? activeTabId;
  if (!target) return Promise.resolve();
  const st = tabs.get(target);
  if (!st) return Promise.resolve();
  return Promise.race([
    st.queue,
    new Promise<void>((resolve) => setTimeout(resolve, timeoutMs)),
  ]).then(() => undefined);
}

export function ironcalcEngineOnClose(tabId: string): void {
  appliers.delete(tabId);
  snapshotters.delete(tabId);
  if (!tabs.has(tabId)) return;
  tabs.delete(tabId);
  if (activeTabId === tabId) activeTabId = null;
  void invoke("calc_engine_unload", { tab: tabId }).catch(() => {});
}

function degrade(tabId: string, reason: string): void {
  const st = tabs.get(tabId);
  if (!st || st.degraded) return;
  st.degraded = reason;
  console.error(
    `[ironcalc-engine] DEGRADED (${reason}) — values frozen for this tab; ` +
      `reopen the workbook to resync the engine`,
  );
}

/** Safety valve for resync itself: a grid too big to dump falls back to the
 * old degrade behavior instead of shipping a 100MB IPC payload. */
const RESYNC_MAX_CELLS = 2_000_000;

/**
 * Rebuilds the engine model from the grid — the recovery path for anything
 * the command bridge cannot mirror. The grid is the source of truth for
 * cell content; the model keeps its defined names and styles (reseeded in
 * place, never rebuilt), and formula cells are seeded WITH their current
 * grid value so frozen add-in/external-ref data survives.
 *
 * Queued on the tab's serial queue: everything already mirrored applies
 * first, and the snapshot is read at task run time so it includes any grid
 * changes that raced the scheduling. Coalesced — one queued resync absorbs
 * all triggers until it runs (commands arriving meanwhile are not mirrored;
 * the snapshot carries their effects).
 */
function resync(tabId: string, reason: string): void {
  const st = tabs.get(tabId);
  if (!st || st.degraded || st.resyncQueued) return;
  st.resyncQueued = true;
  console.warn(`[ironcalc-engine] RESYNC (${reason}) — rebuilding the model from the grid`);
  enqueue(tabId, async () => {
    const stNow = tabs.get(tabId);
    if (!stNow) return;
    stNow.resyncQueued = false;
    const snap = snapshotters.get(tabId)?.() ?? null;
    if (!snap || snap.length === 0) {
      degrade(tabId, `${reason} — grid snapshot unavailable`);
      return;
    }
    const total = snap.reduce((n, s) => n + s.cells.length, 0);
    if (total > RESYNC_MAX_CELLS) {
      degrade(tabId, `${reason} — grid too large to resync (${total} cells)`);
      return;
    }
    try {
      const t0 = performance.now();
      const res = await invoke<{ sheets: string[]; cells: CellDelta[]; seeded: number; ms: number }>(
        "calc_engine_resync",
        { tab: tabId, sheets: snap },
      );
      stNow.sheets = res.sheets;
      stNow.ready = true;
      applyDeltas(tabId, res.cells);
      console.log(
        `[ironcalc-engine] resynced (${reason}): ${res.seeded} cells seeded → ` +
          `${res.cells.length} formula values re-applied ` +
          `(engine ${res.ms}ms, total ${(performance.now() - t0).toFixed(0)}ms)`,
      );
    } catch (e) {
      degrade(tabId, `resync failed: ${e}`);
    }
  });
}

/** Chains engine work for a tab so operations apply strictly in order. */
function enqueue(tabId: string, work: () => Promise<void>): void {
  const st = tabs.get(tabId);
  if (!st) return;
  st.queue = st.queue.then(work).catch(() => {});
}

function applyDeltas(tabId: string, deltas: CellDelta[]): void {
  const st = tabs.get(tabId);
  if (!st || deltas.length === 0) return;
  const applier = appliers.get(tabId);
  if (!applier) {
    // Grid not mounted yet (load raced the mount) — flushed on register.
    st.pending.push(...deltas);
    return;
  }
  const bySheet = new Map<number, { row0: number; col0: number; v: unknown }[]>();
  for (const d of deltas) {
    let list = bySheet.get(d.sheet);
    if (!list) {
      list = [];
      bySheet.set(d.sheet, list);
    }
    list.push({ row0: d.row - 1, col0: d.column - 1, v: d.value });
  }
  for (const [sheetIndex, cells] of bySheet) {
    const sheetName = st.sheets[sheetIndex];
    if (!sheetName) continue;
    const ok = applier(sheetName, cells);
    if (!ok) {
      // Grid write failed (sheet renamed/missing?) — queue for retry on next
      // activation rather than silently dropping computed values.
      st.pending.push(
        ...cells.map((c) => ({
          sheet: sheetIndex,
          row: c.row0 + 1,
          column: c.col0 + 1,
          value: c.v,
        })),
      );
    }
  }
}

interface EditCellWire {
  sheet: number;
  row: number; // 1-based
  column: number; // 1-based
  value: string;
}

const MAX_EDIT_CELLS = 50_000;

async function engineEditBatch(tabId: string, label: string, edits: EditCellWire[]): Promise<void> {
  const st = tabs.get(tabId);
  if (!st || st.degraded || st.resyncQueued || edits.length === 0) return;
  try {
    const res = await invoke<CalcEditResult>("calc_engine_edit_batch", { tab: tabId, edits });
    console.log(
      `[ironcalc-engine] ${label}: ${edits.length} cell${edits.length === 1 ? "" : "s"} → ` +
        `recalc ${res.recalc_ms}ms, ${res.changed.length} cells updated`,
    );
    applyDeltas(tabId, res.changed);
  } catch (e) {
    // The model may be mid-mutation (a failed multi-cell batch) — rebuild it
    // from the grid rather than freezing the session.
    resync(tabId, `edit failed: ${e}`);
  }
}

async function engineStructural(
  tabId: string,
  op: string,
  sheetName: string,
  index1: number,
  count: number,
): Promise<void> {
  const st = tabs.get(tabId);
  if (!st || st.degraded || st.resyncQueued) return;
  const sheet = st.sheets.indexOf(sheetName);
  if (sheet < 0) {
    resync(tabId, `${op}: unknown sheet "${sheetName}"`);
    return;
  }
  try {
    const snap = await invoke<CalcSnapshot>("calc_engine_structural", {
      tab: tabId,
      op,
      sheet,
      index: index1,
      count,
    });
    console.log(
      `[ironcalc-engine] ${op}(${index1}, ${count}) on ${sheetName} → full recalc ${snap.ms}ms, ` +
        `${snap.cells.length} values re-applied`,
    );
    applyDeltas(tabId, snap.cells);
  } catch (e) {
    resync(tabId, `${op} failed: ${e}`);
  }
}

/** Content-moving commands we cannot mirror — grid truth via resync. Paste
 * is deliberately absent: plain pastes carry decodable set-range-values
 * children and take the fast mirror path; anything opaque in a paste (si
 * fills) is caught by the child classification instead. */
const RESYNC_HINTS = [
  "move-rows",
  "move-cols",
  "insert-range-move",
  "delete-range-move",
  "sort-range",
  "worksheet-merge",
];

/** Mutations that MOVE content (sort reorders, range moves). Buffered so an
 * unrecognized parent command (or an undo/redo replay) can be recognized as
 * content-changing; never decodable into edits — always a resync. */
const CONTENT_MOVER_MUTATIONS = new Set([
  "sheet.mutation.move-range",
  "sheet.mutation.move-rows",
  "sheet.mutation.move-cols",
  "sheet.mutation.reorder-range",
]);

// ---------------------------------------------------------------------------
// Undo/redo replay buffer.
//
// A command's child mutations complete (and appear in the stream) BEFORE the
// parent command does. Unattributed edit mutations are held here for one
// tick; the parent decides their fate: undo/redo → flush into the engine,
// normal edit command → drop (the command path mirrors it), nobody (load /
// style population) → evaporate on the timeout.
// ---------------------------------------------------------------------------

interface BufferedMutation {
  id: string;
  params: Record<string, unknown>;
  /** Sheet-name map values snapshotted when the mutation was seen (rename
   * old-name detection needs the pre/post views). */
  namesSnapshot?: Set<string>;
}

function bufferMutation(
  st: TabState,
  id: string,
  params: Record<string, unknown>,
  namesSnapshot?: Set<string>,
): void {
  st.buffer.push({ id, params, namesSnapshot });
  if (!st.bufferDropScheduled) {
    st.bufferDropScheduled = true;
    setTimeout(() => {
      st.bufferDropScheduled = false;
      st.buffer = [];
    }, 0);
  }
}

function flushReplayBuffer(tabId: string, sheetNamesById: Map<string, string>): void {
  const st = tabs.get(tabId);
  if (!st) return;
  const buffered = st.buffer;
  st.buffer = [];
  if (st.degraded || st.resyncQueued || buffered.length === 0) return;
  for (const m of buffered) {
    if (CONTENT_MOVER_MUTATIONS.has(m.id)) {
      // Undoing a sort/move replays reorder/move mutations we can't decode.
      resync(tabId, `undo/redo replay ${m.id}`);
      return;
    }
    if (m.id === "sheet.mutation.set-range-values") {
      const sheetName = resolveSheet(m.params, sheetNamesById);
      if (!sheetName) {
        resync(tabId, "undo replay without resolvable sheet");
        return;
      }
      const sheet = st.sheets.indexOf(sheetName);
      if (sheet < 0) {
        resync(tabId, `undo replay: unknown sheet "${sheetName}"`);
        return;
      }
      const decoded = decodeMutationCells(m.params);
      if (decoded.fail) {
        resync(tabId, `undo replay: ${decoded.fail}`);
        return;
      }
      if (decoded.cells.length === 0) continue; // style-only
      const edits: EditCellWire[] = decoded.cells.map((c) => ({
        sheet,
        row: c.row0 + 1,
        column: c.col0 + 1,
        value: c.input,
      }));
      enqueue(tabId, () => engineEditBatch(tabId, "undo/redo replay", edits));
      continue;
    }
    const structural = decodeStructural(m.id, m.params);
    if (structural) {
      const sheetName = resolveSheet(m.params, sheetNamesById);
      if (!sheetName) {
        resync(tabId, `${m.id}: unknown sheet`);
        return;
      }
      enqueue(tabId, () =>
        engineStructural(tabId, structural.op, sheetName, structural.index1, structural.count),
      );
    }
  }
}

const SHEET_OP_MUTATIONS = new Set([
  "sheet.mutation.insert-sheet",
  "sheet.mutation.remove-sheet",
  "sheet.mutation.set-worksheet-name",
]);

/** Mirrors buffered sheet-structure mutations (called when a sheet command
 * or an undo/redo claims them); drops everything else in the buffer. */
function flushSheetOps(tabId: string): void {
  const st = tabs.get(tabId);
  if (!st) return;
  const buffered = st.buffer;
  st.buffer = [];
  for (const m of buffered) {
    if (SHEET_OP_MUTATIONS.has(m.id)) mirrorSheetMutation(tabId, m);
  }
}

function mirrorSheetMutation(tabId: string, m: BufferedMutation): void {
  const st = tabs.get(tabId);
  if (!st || st.degraded) return;
  const p = m.params as Record<string, unknown> & { sheet?: { name?: unknown }; name?: unknown };
  if (m.id === "sheet.mutation.insert-sheet") {
    const name = typeof p.sheet?.name === "string" ? p.sheet.name : undefined;
    if (!name) {
      resync(tabId, "insert-sheet without a name");
      return;
    }
    enqueue(tabId, () => engineSheetOp(tabId, "add", name, undefined));
    return;
  }
  if (m.id === "sheet.mutation.remove-sheet") {
    const name =
      typeof p.sheet?.name === "string"
        ? p.sheet.name
        : typeof (p as any).sheetName === "string"
          ? ((p as any).sheetName as string)
          : undefined;
    if (!name) {
      resync(tabId, "remove-sheet without a name");
      return;
    }
    enqueue(tabId, () => engineSheetOp(tabId, "delete", name, undefined));
    return;
  }
  if (m.id === "sheet.mutation.set-worksheet-name") {
    const newName = typeof p.name === "string" ? p.name : undefined;
    if (!newName) {
      resync(tabId, "rename without a name");
      return;
    }
    enqueue(tabId, async () => {
      const stNow = tabs.get(tabId);
      if (!stNow || stNow.degraded || stNow.resyncQueued) return;
      // Old name = the model sheet no longer among the grid's names (the
      // snapshot was taken after the rename mutation applied, so exactly one
      // model sheet should be missing from it).
      const gridNames = m.namesSnapshot;
      const missing = gridNames ? stNow.sheets.filter((n) => !gridNames.has(n)) : [];
      if (missing.length !== 1) {
        // Resync reconciles the sheet set by name, so an ambiguous rename
        // heals itself.
        resync(tabId, `rename: cannot identify old name for "${newName}"`);
        return;
      }
      await engineSheetOp(tabId, "rename", missing[0], newName);
    });
    return;
  }
}

async function engineSheetOp(
  tabId: string,
  op: string,
  name: string,
  newName: string | undefined,
): Promise<void> {
  const st = tabs.get(tabId);
  if (!st || st.degraded || st.resyncQueued) return;
  try {
    const res = await invoke<{ sheets: string[]; cells: CellDelta[] }>("calc_engine_sheet_op", {
      tab: tabId,
      op,
      name,
      newName: newName ?? null,
    });
    st.sheets = res.sheets;
    console.log(
      `[ironcalc-engine] sheet ${op} "${name}"${newName ? ` -> "${newName}"` : ""} → ` +
        `sheets [${res.sheets.join(", ")}], ${res.cells.length} values re-applied`,
    );
    applyDeltas(tabId, res.cells);
  } catch (e) {
    resync(tabId, `sheet ${op} failed: ${e}`);
  }
}

function resolveSheet(
  p: Record<string, unknown>,
  sheetNamesById: Map<string, string>,
): string | null {
  const subUnitId = p.subUnitId;
  return (typeof subUnitId === "string" ? sheetNamesById.get(subUnitId) : undefined) ?? null;
}

/**
 * Feed the Univer command/mutation stream into the engine. Called from
 * UniverGrid's `onCommandExecuted` when engine mode is on.
 */
export function ironcalcEngineOnCommand(
  tabId: string,
  cmd: { id?: unknown; params?: unknown },
  sheetNamesById: Map<string, string>,
): void {
  const st = tabs.get(tabId);
  // resyncQueued: the pending resync's snapshot will carry this command's
  // grid effects — mirroring it into the drifted model would only add noise.
  if (!st || st.degraded || st.resyncQueued) return;
  if (!cmd || typeof cmd.id !== "string") return;
  const p = (cmd.params ?? {}) as Record<string, unknown>;

  // Our own value write-back (tagged by the applier in UniverGrid).
  if (p.__ironcalc === true) return;

  // Mutations: buffer for parent-command attribution.
  if (cmd.id.startsWith("sheet.mutation.")) {
    if (
      cmd.id === "sheet.mutation.set-range-values" ||
      decodeStructural(cmd.id, p) !== null ||
      SHEET_OP_MUTATIONS.has(cmd.id) ||
      CONTENT_MOVER_MUTATIONS.has(cmd.id)
    ) {
      // Sheet-op mutations need the *pre-mutation* name map for rename
      // old-name detection; snapshot it now.
      bufferMutation(st, cmd.id, p, new Set(sheetNamesById.values()));
    }
    return;
  }

  // Undo/redo: the buffered mutations are the replayed content.
  if (cmd.id === "univer.command.undo" || cmd.id === "univer.command.redo") {
    const buffered = st.buffer;
    for (const m of buffered) {
      if (SHEET_OP_MUTATIONS.has(m.id)) mirrorSheetMutation(tabId, m);
    }
    st.buffer = buffered.filter((m) => !SHEET_OP_MUTATIONS.has(m.id));
    flushReplayBuffer(tabId, sheetNamesById);
    return;
  }

  if (!cmd.id.startsWith("sheet.command.")) return;
  // Sheet-structure commands claim their buffered mutations: the MUTATION
  // carries the authoritative sheet name (an insert command may not name
  // the generated "Sheet2"), so mirroring happens from the buffer here.
  if (
    cmd.id.includes("insert-sheet") ||
    cmd.id.includes("remove-sheet") ||
    cmd.id.includes("worksheet-name")
  ) {
    flushSheetOps(tabId);
    return;
  }
  // A RECOGNIZED edit command claims (and thus invalidates) any buffered
  // child mutations — the command path below mirrors the change itself. For
  // an UNRECOGNIZED command the claimed mutations are the evidence of what
  // it did to cell content: plain writes are mirrored from them directly at
  // the bottom, anything opaque triggers a resync.
  const claimed = st.buffer;
  st.buffer = [];

  const sheetName = resolveSheet(p, sheetNamesById);

  if (cmd.id === "sheet.command.set-range-values") {
    const cells = extractEditedCells(p, MAX_EDIT_CELLS);
    if (!sheetName || cells === null) {
      resync(tabId, "unparsed or oversized set-range-values params");
      return;
    }
    const sheet = st.sheets.indexOf(sheetName);
    if (sheet < 0 && st.ready) {
      resync(tabId, `unknown sheet "${sheetName}"`);
      return;
    }
    const edits: EditCellWire[] = cells.map((c) => ({
      sheet,
      row: c.row0 + 1,
      column: c.col0 + 1,
      value: c.input,
    }));
    if (edits.length > MAX_EDIT_CELLS) {
      resync(tabId, `edit over ${MAX_EDIT_CELLS} cells`);
      return;
    }
    if (!st.ready) {
      // Sheets unknown until the load finishes; re-resolve inside the queue.
      enqueue(tabId, async () => {
        const stNow = tabs.get(tabId);
        if (!stNow) return;
        const sheetNow = stNow.sheets.indexOf(sheetName);
        if (sheetNow < 0) {
          resync(tabId, `unknown sheet "${sheetName}"`);
          return;
        }
        await engineEditBatch(
          tabId,
          sheetName,
          edits.map((e) => ({ ...e, sheet: sheetNow })),
        );
      });
    } else {
      enqueue(tabId, () => engineEditBatch(tabId, sheetName, edits));
    }
    return;
  }

  // Clearing a selection = writing empty over the range.
  if (cmd.id.includes("clear-selection")) {
    const range = p.range as Record<string, unknown> | undefined;
    if (!sheetName || !range || typeof range.startRow !== "number") {
      resync(tabId, cmd.id);
      return;
    }
    const sheet = st.sheets.indexOf(sheetName);
    if (sheet < 0) {
      resync(tabId, `${cmd.id}: unknown sheet "${sheetName}"`);
      return;
    }
    const edits: EditCellWire[] = [];
    for (let r = range.startRow as number; r <= (range.endRow as number); r++) {
      for (let c = range.startColumn as number; c <= (range.endColumn as number); c++) {
        edits.push({ sheet, row: r + 1, column: c + 1, value: "" });
        if (edits.length > MAX_EDIT_CELLS) {
          resync(tabId, `${cmd.id} over ${MAX_EDIT_CELLS} cells`);
          return;
        }
      }
    }
    enqueue(tabId, () => engineEditBatch(tabId, `${sheetName} clear`, edits));
    return;
  }

  // Row/column structure via the user-facing commands.
  const structural = decodeStructural(cmd.id, p);
  if (structural) {
    if (!sheetName) {
      resync(tabId, `${cmd.id}: unknown sheet`);
      return;
    }
    enqueue(tabId, () =>
      engineStructural(tabId, structural.op, sheetName, structural.index1, structural.count),
    );
    return;
  }

  if (RESYNC_HINTS.some((h) => (cmd.id as string).includes(h))) {
    resync(tabId, cmd.id);
    return;
  }

  // Unrecognized sheet.command.*: judge it by its claimed child mutations.
  // Plain decodable writes (drag-fill, simple paste) mirror directly — the
  // mutations are exactly what Univer applied. Anything content-moving or
  // undecodable means the grid changed in a way we can't reproduce: resync.
  const verdict = classifyClaimedMutations(st.sheets, claimed, sheetNamesById);
  if (verdict.kind === "edits") {
    enqueue(tabId, () => engineEditBatch(tabId, `mirror ${cmd.id}`, verdict.edits));
  } else if (verdict.kind === "opaque") {
    resync(tabId, `${cmd.id} (${verdict.why})`);
  }
}

/**
 * Classifies the child mutations an unrecognized command claimed:
 * `none` — nothing content-relevant happened (style/selection commands);
 * `edits` — plain cell writes, decodable and safe to mirror as one batch;
 * `opaque` — content changed in a way we cannot decode (movers, si fills,
 * unknown sheets) — the caller must resync.
 */
function classifyClaimedMutations(
  sheets: string[],
  claimed: BufferedMutation[],
  sheetNamesById: Map<string, string>,
):
  | { kind: "none" }
  | { kind: "edits"; edits: EditCellWire[] }
  | { kind: "opaque"; why: string } {
  const edits: EditCellWire[] = [];
  for (const m of claimed) {
    if (CONTENT_MOVER_MUTATIONS.has(m.id) || SHEET_OP_MUTATIONS.has(m.id)) {
      return { kind: "opaque", why: m.id };
    }
    if (decodeStructural(m.id, m.params) !== null) {
      return { kind: "opaque", why: m.id };
    }
    if (m.id !== "sheet.mutation.set-range-values") continue;
    const sheetName = resolveSheet(m.params, sheetNamesById);
    if (!sheetName) return { kind: "opaque", why: "unresolvable sheet" };
    const sheet = sheets.indexOf(sheetName);
    if (sheet < 0) return { kind: "opaque", why: `unknown sheet "${sheetName}"` };
    const decoded = decodeMutationCells(m.params);
    if (decoded.fail) return { kind: "opaque", why: decoded.fail };
    for (const c of decoded.cells) {
      edits.push({ sheet, row: c.row0 + 1, column: c.col0 + 1, value: c.input });
      if (edits.length > MAX_EDIT_CELLS) {
        return { kind: "opaque", why: `over ${MAX_EDIT_CELLS} cells` };
      }
    }
  }
  return edits.length > 0 ? { kind: "edits", edits } : { kind: "none" };
}

interface EditedCell {
  row0: number;
  col0: number;
  input: string;
}

/**
 * Decode a SetRangeValuesMutation `cellValue` matrix into engine inputs.
 * Style-only patches are skipped; shared-formula `si` refs without formula
 * text cannot be reconstructed and fail the whole mutation (resync beats
 * drift).
 */
function decodeMutationCells(p: Record<string, unknown>): {
  cells: EditedCell[];
  fail: string | null;
} {
  const value = p.cellValue ?? p.value;
  if (value === null || value === undefined) return { cells: [], fail: null };
  if (typeof value !== "object") return { cells: [], fail: "unrecognized payload" };

  const cells: EditedCell[] = [];
  for (const [r, row] of Object.entries(value as Record<string, unknown>)) {
    if (!/^\d+$/.test(r)) return { cells: [], fail: "unrecognized matrix key" };
    if (row === null || row === undefined) continue;
    if (typeof row !== "object") return { cells: [], fail: "unrecognized row payload" };
    for (const [c, cell] of Object.entries(row as Record<string, unknown>)) {
      if (!/^\d+$/.test(c)) return { cells: [], fail: "unrecognized matrix key" };
      const input = mutationCellToInput(cell);
      if (input === "si") return { cells: [], fail: "shared-formula si without text" };
      if (input === null) continue; // style/rich-text-only patch
      cells.push({ row0: Number(r), col0: Number(c), input });
      if (cells.length > MAX_EDIT_CELLS) {
        return { cells: [], fail: `mutation over ${MAX_EDIT_CELLS} cells` };
      }
    }
  }
  return { cells, fail: null };
}

/**
 * ICellData → engine input string; `null` = not a content change (skip);
 * `"si"` sentinel = unreconstructable shared-formula ref.
 */
function mutationCellToInput(cell: unknown): string | null {
  if (cell === null || cell === undefined) return ""; // cleared cell
  if (typeof cell !== "object") return null;
  const c = cell as Record<string, unknown>;
  if (typeof c.f === "string" && c.f.length > 0) return c.f;
  if (typeof c.si === "string" && c.si.length > 0) return "si";
  const hasContentKey = "v" in c || "f" in c;
  if (!hasContentKey) return null; // style-only ({s}), rich-text-only ({p})
  if (c.v === null || c.v === undefined) return ""; // explicit clear
  if (typeof c.v === "boolean") return c.v ? "TRUE" : "FALSE";
  return String(c.v);
}

function decodeStructural(
  id: string,
  p: Record<string, unknown>,
): { op: string; index1: number; count: number } | null {
  const m = id.match(/^sheet\.(command|mutation)\.(.+)$/);
  if (!m) return null;
  const kind = m[1];
  const op = m[2];
  const range = (p.range ?? (p as any).ranges?.[0]) as Record<string, unknown> | undefined;
  if (!range) return null;
  const rows =
    typeof range.startRow === "number" && typeof range.endRow === "number"
      ? {
          index1: (range.startRow as number) + 1,
          count: (range.endRow as number) - (range.startRow as number) + 1,
        }
      : null;
  const cols =
    typeof range.startColumn === "number" && typeof range.endColumn === "number"
      ? {
          index1: (range.startColumn as number) + 1,
          count: (range.endColumn as number) - (range.startColumn as number) + 1,
        }
      : null;
  if (kind === "command") {
    // Every insert/delete funnels through exactly one -by-range command: the
    // facade helpers (the agent path) dispatch it directly, and the menu's
    // bare / -before / -after commands reach it via delegation. Matching ONLY
    // the funnel mirrors each operation once, with the handler-resolved range
    // (the bare remove-row command may carry no range at all — it falls back
    // to the current selection inside its handler).
    if (op === "insert-row-by-range") return rows && { op: "insert_rows", ...rows };
    if (op === "remove-row-by-range") return rows && { op: "delete_rows", ...rows };
    if (op === "insert-col-by-range") return cols && { op: "insert_columns", ...cols };
    if (op === "remove-col-by-range") return cols && { op: "delete_columns", ...cols };
    return null;
  }
  // Mutations (undo/redo replays execute these directly, no parent command).
  // Exact ids as Univer names them: remove-rows is plural, the others are not.
  if (op === "insert-row") return rows && { op: "insert_rows", ...rows };
  if (op === "remove-row" || op === "remove-rows") return rows && { op: "delete_rows", ...rows };
  if (op === "insert-col") return cols && { op: "insert_columns", ...cols };
  if (op === "remove-col" || op === "remove-cols") return cols && { op: "delete_columns", ...cols };
  return null;
}

/** For tests. */
export const __internals = {
  decodeStructural,
  decodeMutationCells,
  mutationCellToInput,
  classifyClaimedMutations,
  tabs,
};
