/**
 * Unsaved-work journal for FILE-BACKED workbooks.
 *
 * Until Cmd-S, every edit to an opened .xlsx lives only in memory: a crash
 * (the app running out of memory on a 2M-cell model, a kernel panic) loses
 * them, while untitled drafts have long had a 30-second auto-snapshot.
 * Snapshotting a 2M-cell workbook every half minute is itself a memory and
 * time problem, so file-backed tabs get a journal instead: a small JSON
 * record of what changed since the last save, written a couple of seconds
 * after each change to <app_data>/journals, deleted on save. On reopening
 * the same, unchanged file the user is offered a replay.
 *
 * What is journaled, in order:
 * - `edits`: every cell write, row/column insert/delete and sheet
 *   add/delete/rename that reaches the calc-engine mirror (ironcalcEngine.ts)
 *   — user typing, pastes, undo/redo and agent tool calls alike, because the
 *   mirror sees Univer's command stream. Values are the raw input text
 *   ("=SUM(A1:A3)", "42", "hello").
 * - `batches`: the agent's review batches not yet saved, so the review UI
 *   and the in-place saver (formats, merges, sheet ops) get them back.
 * - `manual`: the toolbar mirrors (formats, borders, merges, row/col ops,
 *   freeze, visibility, sheet ops) the saver already keeps.
 *
 * Replay writes `edits` back through the grid handle in order, then re-applies
 * the style-level parts of batches and the toolbar mirrors. Content-moving
 * commands the engine mirror cannot decode (sort, move range) are recorded
 * as a `resync` marker: a journal containing one can only be replayed
 * partially, and the user is told so.
 */
import { invoke } from "@tauri-apps/api/core";
import type { ChangeBatch, SaveMirror, UniverMutation } from "./types";
import type { UniverGridHandle } from "./components/UniverGrid";
import { applyMutationForward } from "./agent/batchMutations";

export type JournalEdit =
  | { k: "cells"; sheet: string; cells: Array<{ r: number; c: number; v: string }> }
  | { k: "rows" | "cols"; op: "insert" | "delete"; sheet: string; at: number; count: number }
  | { k: "sheet"; op: "add" | "delete" | "rename"; name: string; newName?: string }
  | { k: "resync"; reason: string };

export type ManualMirrors = {
  formats: NonNullable<SaveMirror["cellFormats"]>;
  borders: NonNullable<SaveMirror["cellBorders"]>;
  merges: NonNullable<SaveMirror["merges"]>;
  rowCol: NonNullable<SaveMirror["rowColOps"]>;
  freeze: NonNullable<SaveMirror["freezePanes"]>[number] | null;
  visibility: NonNullable<SaveMirror["visibility"]>;
  sheetOps: NonNullable<SaveMirror["sheetOps"]>;
};

export type WorkbookJournal = {
  version: 1;
  path: string;
  /** `${size}:${mtime_ms}` of the file the edits apply to. */
  stamp: string;
  updatedAt: string;
  edits: JournalEdit[];
  batches: ChangeBatch[];
  manual: ManualMirrors;
};

type TabJournal = {
  path: string;
  stamp: string;
  edits: JournalEdit[];
  /** Serialised form last written to disk, to skip no-op flushes. */
  lastWritten: string | null;
  /** Edits were recorded since the last flush. */
  dirty: boolean;
  /**
   * Wall-clock ms of the last save of this tab's file (0 = none this
   * session). Every batch that existed at that moment is IN the file, so
   * the journal must not carry it — `persisted` cannot be used for this:
   * after a surgical save the batches stay unpersisted on purpose (the
   * next save re-applies them against the original bytes), and filtering
   * on it made the 2-second flush write every saved batch straight back
   * into the journal, so the next open offered to "recover" saved work.
   */
  savedAt: number;
};

/** ms of a batch's creation; unparsable → 0 (journal it rather than lose it). */
function createdMs(b: ChangeBatch): number {
  const t = Date.parse(b.created_at);
  return Number.isFinite(t) ? t : 0;
}

const tabs = new Map<string, TabJournal>();

/** Journal recording is suspended for a tab while its own journal replays. */
const replaying = new Set<string>();

/** Cap on recorded edits per tab: beyond this the journal keeps the first
 * entries and a resync marker — a bounded file beats an unbounded one. */
const MAX_EDITS = 200_000;

function fnv1a(s: string): string {
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  for (let i = 0; i < s.length; i++) {
    const ch = s.charCodeAt(i);
    h1 ^= ch;
    h1 = Math.imul(h1, 0x01000193) >>> 0;
    h2 ^= ch;
    h2 = Math.imul(h2, 0x811c9dc5) >>> 0;
  }
  return h1.toString(16).padStart(8, "0") + h2.toString(16).padStart(8, "0");
}

export function journalName(path: string): string {
  return `wb-${fnv1a(path)}.json`;
}

async function fileStamp(path: string): Promise<string | null> {
  try {
    const s = await invoke<{ size: number; mtime_ms: number }>("workbook_file_stamp", { path });
    return `${s.size}:${s.mtime_ms}`;
  } catch {
    return null;
  }
}

export function emptyManual(): ManualMirrors {
  return { formats: [], borders: [], merges: [], rowCol: [], freeze: null, visibility: [], sheetOps: [] };
}

function manualIsEmpty(m: ManualMirrors): boolean {
  return (
    !m.formats.length && !m.borders.length && !m.merges.length && !m.rowCol.length && !m.freeze && !m.visibility.length && !m.sheetOps.length
  );
}

export function journalHasContent(j: WorkbookJournal): boolean {
  return j.edits.length > 0 || j.batches.length > 0 || !manualIsEmpty(j.manual);
}

/** Human summary for the recovery prompt. */
export function describeJournal(j: WorkbookJournal): string {
  let cells = 0;
  let structural = 0;
  let sheets = 0;
  let partial = false;
  for (const e of j.edits) {
    if (e.k === "cells") cells += e.cells.length;
    else if (e.k === "rows" || e.k === "cols") structural++;
    else if (e.k === "sheet") sheets++;
    else partial = true;
  }
  const parts: string[] = [];
  if (cells) parts.push(`${cells} cell edit${cells === 1 ? "" : "s"}`);
  if (structural) parts.push(`${structural} row/column change${structural === 1 ? "" : "s"}`);
  if (sheets) parts.push(`${sheets} sheet change${sheets === 1 ? "" : "s"}`);
  if (j.batches.length) parts.push(`${j.batches.length} agent batch${j.batches.length === 1 ? "" : "es"}`);
  if (!manualIsEmpty(j.manual)) parts.push("toolbar formatting");
  let s = parts.join(", ") || "changes";
  if (partial) s += " (a sort or move could not be journaled — recovery may be partial)";
  return s;
}

/** Start journaling a file-backed tab. Call once the file is open. */
export async function journalOpen(tabId: string, path: string): Promise<void> {
  const stamp = (await fileStamp(path)) ?? "unknown";
  tabs.set(tabId, { path, stamp, edits: [], lastWritten: null, dirty: false,
    savedAt: 0 });
}

/**
 * The journal left for `path` by a previous session, if it belongs to the
 * file as it is now. A stale journal (the file changed since) is deleted.
 */
export async function journalLoad(path: string): Promise<WorkbookJournal | null> {
  const name = journalName(path);
  let raw: string | null = null;
  try {
    raw = await invoke<string | null>("journal_read", { name });
  } catch (e) {
    console.warn("[journal] read failed:", e);
    return null;
  }
  if (!raw) return null;
  let j: WorkbookJournal;
  try {
    j = JSON.parse(raw) as WorkbookJournal;
  } catch {
    void invoke("journal_delete", { name }).catch(() => {});
    return null;
  }
  if (j.version !== 1 || !journalHasContent(j)) {
    void invoke("journal_delete", { name }).catch(() => {});
    return null;
  }
  const stamp = await fileStamp(path);
  if (stamp && j.stamp !== "unknown" && j.stamp !== stamp) {
    console.info(`[journal] ${path}: file changed since the journal was written (${j.stamp} → ${stamp}); discarding it`);
    void invoke("journal_delete", { name }).catch(() => {});
    return null;
  }
  return j;
}

/** Record one edit that reached the calc-engine mirror. */
export function journalRecordEdit(tabId: string, edit: JournalEdit): void {
  const t = tabs.get(tabId);
  if (!t || replaying.has(tabId)) return;
  if (t.edits.length >= MAX_EDITS) {
    const last = t.edits[t.edits.length - 1];
    if (last.k !== "resync") t.edits.push({ k: "resync", reason: "journal edit cap reached" });
    t.dirty = true;
    return;
  }
  // Coalesce consecutive cell writes on the same sheet.
  const last = t.edits[t.edits.length - 1];
  if (edit.k === "cells" && last?.k === "cells" && last.sheet === edit.sheet && last.cells.length + edit.cells.length <= 5000) {
    last.cells.push(...edit.cells);
  } else {
    t.edits.push(edit);
  }
  t.dirty = true;
}

/**
 * Write the tab's journal if anything changed since the last write. Called
 * on a timer by the screen with the tab's current batches and toolbar
 * mirrors (those live in React state / refs, not here).
 */
export async function journalFlush(tabId: string, batches: ChangeBatch[], manual: ManualMirrors): Promise<void> {
  const t = tabs.get(tabId);
  if (!t) return;
  const payload: WorkbookJournal = {
    version: 1,
    path: t.path,
    stamp: t.stamp,
    updatedAt: new Date().toISOString(),
    edits: t.edits,
    // Only batches created after the last save: everything older is in the file.
    batches: t.savedAt > 0 ? batches.filter((b) => createdMs(b) === 0 || createdMs(b) > t.savedAt) : batches,
    manual,
  };
  const name = journalName(t.path);
  if (!journalHasContent(payload)) {
    if (t.lastWritten !== null) {
      t.lastWritten = null;
      t.dirty = false;
      await invoke("journal_delete", { name }).catch(() => {});
    }
    return;
  }
  // Compare without the timestamp so an unchanged journal is not rewritten.
  const body = JSON.stringify({ ...payload, updatedAt: undefined });
  if (!t.dirty && body === t.lastWritten) return;
  t.dirty = false;
  t.lastWritten = body;
  try {
    await invoke("journal_write", { name, json: JSON.stringify(payload) });
  } catch (e) {
    console.warn("[journal] write failed:", e);
  }
}

/** The file was saved: the journal is history. Re-stamps for the saved file. */
export async function journalClear(tabId: string, pathAfterSave?: string): Promise<void> {
  const t = tabs.get(tabId);
  if (!t) return;
  await invoke("journal_delete", { name: journalName(t.path) }).catch(() => {});
  if (pathAfterSave && pathAfterSave !== t.path) {
    await invoke("journal_delete", { name: journalName(pathAfterSave) }).catch(() => {});
    t.path = pathAfterSave;
  }
  t.edits = [];
  t.lastWritten = null;
  t.dirty = false;
  t.savedAt = Date.now();
  t.stamp = (await fileStamp(t.path)) ?? "unknown";
}

/**
 * Stop journaling a tab. The on-disk journal is kept when the tab still had
 * unsaved changes (the user may have closed by accident; the next open of
 * the file offers it back) and removed otherwise.
 */
export async function journalClose(tabId: string, keepFile: boolean): Promise<void> {
  const t = tabs.get(tabId);
  tabs.delete(tabId);
  replaying.delete(tabId);
  if (t && !keepFile) await invoke("journal_delete", { name: journalName(t.path) }).catch(() => {});
}

/** Every sheet a journal touches — lazy sheets must be loaded before replay. */
export function journalSheets(j: WorkbookJournal): Set<string> {
  const out = new Set<string>();
  for (const e of j.edits) {
    if (e.k === "cells" || e.k === "rows" || e.k === "cols") out.add(e.sheet);
    else if (e.k === "sheet") {
      out.add(e.name);
      if (e.newName) out.add(e.newName);
    }
  }
  for (const b of j.batches) for (const m of b.mutations as any[]) {
    if (typeof m.sheet === "string") out.add(m.sheet);
    if (typeof m.address?.sheet === "string") out.add(m.address.sheet);
  }
  const m = j.manual;
  for (const x of [...m.formats, ...m.borders, ...m.merges, ...m.rowCol, ...m.visibility]) out.add((x as any).sheet);
  if (m.freeze) out.add(m.freeze.sheet);
  return out;
}

/** Everything a tab currently has journaled — for tests and diagnostics. */
export function journalPeek(tabId: string): JournalEdit[] {
  return tabs.get(tabId)?.edits ?? [];
}

/** Text typed into a cell → the value `setCells` stores for it. */
export function coerceInput(v: string): string | number | boolean | null {
  if (v === "") return null;
  if (v.startsWith("=")) return v;
  if (v === "TRUE") return true;
  if (v === "FALSE") return false;
  if (/^-?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(v)) {
    const n = Number(v);
    if (Number.isFinite(n)) return n;
  }
  return v;
}

const STYLE_LEVEL: ReadonlySet<UniverMutation["type"]> = new Set<UniverMutation["type"]>([
  "set_format",
  "merge_cells",
  "unmerge_cells",
  "freeze_panes",
  "unfreeze_panes",
  "hide_rows",
  "show_rows",
  "hide_columns",
  "show_columns",
  "set_column_width",
  "set_row_height",
  "set_note",
  "delete_note",
  "define_name",
]);

export type ReplayResult = { cells: number; structural: number; sheets: number; styles: number; partial: boolean };

/**
 * Replay a journal onto a freshly loaded grid. Edits go first, in order;
 * then the style-level mutations of the batches (their cell writes and
 * structure already came through the edit stream) and the toolbar mirrors.
 * Recording for `tabId` is suspended meanwhile so the replay does not
 * journal itself twice (the grid's own command stream re-records the
 * resulting state as fresh unsaved work, which is what it is).
 */
export function replayJournal(tabId: string, grid: UniverGridHandle, j: WorkbookJournal): ReplayResult {
  const res: ReplayResult = { cells: 0, structural: 0, sheets: 0, styles: 0, partial: false };
  replaying.add(tabId);
  try {
    for (const e of j.edits) {
      try {
        if (e.k === "cells") {
          const ok = grid.setCells(
            e.sheet,
            e.cells.map((c) => ({ row: c.r, col: c.c, value: coerceInput(c.v) })),
          );
          if (ok) res.cells += e.cells.length;
          else res.partial = true;
        } else if (e.k === "rows") {
          if (e.op === "insert") grid.insertRows(e.sheet, e.at, e.count);
          else grid.deleteRows(e.sheet, e.at, e.count);
          res.structural++;
        } else if (e.k === "cols") {
          if (e.op === "insert") grid.insertColumns(e.sheet, e.at, e.count);
          else grid.deleteColumns(e.sheet, e.at, e.count);
          res.structural++;
        } else if (e.k === "sheet") {
          const ok =
            e.op === "add" ? grid.createSheet(e.name, null) : e.op === "delete" ? grid.deleteSheet(e.name) : grid.renameSheet(e.name, e.newName ?? e.name);
          if (ok) res.sheets++;
          else res.partial = true;
        } else {
          res.partial = true;
        }
      } catch (err) {
        console.warn("[journal] replay step failed:", e, err);
        res.partial = true;
      }
    }
    for (const b of j.batches) {
      if (b.status !== "accepted" && b.status !== "pending") continue;
      for (const m of b.mutations) {
        if (!STYLE_LEVEL.has(m.type)) continue;
        try {
          applyMutationForward(grid, m);
          res.styles++;
        } catch (err) {
          console.warn("[journal] batch style replay failed:", m.type, err);
          res.partial = true;
        }
      }
    }
    const m = j.manual;
    for (const f of m.formats) {
      try {
        grid.setCellFormat(f.sheet, f.row, f.col, f.format);
        if (f.background !== undefined) grid.setCellBackground(f.sheet, f.row, f.col, f.background ?? null);
        res.styles++;
      } catch {
        res.partial = true;
      }
    }
    for (const b of m.borders) {
      try {
        grid.setCellBorders(b.sheet, b.row, b.col, b.borders);
        res.styles++;
      } catch {
        res.partial = true;
      }
    }
    for (const mg of m.merges) {
      const r = parseA1(mg.range);
      if (!r) continue;
      try {
        if (mg.merge) grid.mergeCells(mg.sheet, r.r1, r.c1, r.r2, r.c2);
        else grid.unmergeCells(mg.sheet, r.r1, r.c1, r.r2, r.c2);
        res.styles++;
      } catch {
        res.partial = true;
      }
    }
    if (m.freeze) {
      try {
        grid.freezePanes(m.freeze.sheet, m.freeze.freezeRows, m.freeze.freezeCols);
        res.styles++;
      } catch {
        res.partial = true;
      }
    }
    for (const v of m.visibility) {
      try {
        if (v.kind === "hideRows") grid.hideRows(v.sheet, v.rows);
        else if (v.kind === "showRows") grid.showRows(v.sheet, v.rows);
        else if (v.kind === "hideColumns") grid.hideColumns(v.sheet, v.columns);
        else grid.showColumns(v.sheet, v.columns);
        res.styles++;
      } catch {
        res.partial = true;
      }
    }
  } finally {
    replaying.delete(tabId);
  }
  return res;
}

function parseA1(ref: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref.trim());
  if (!m) return null;
  const col = (s: string) => s.toUpperCase().split("").reduce((a, ch) => a * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  const c1 = col(m[1]);
  const r1 = Number(m[2]) - 1;
  return { r1, c1, r2: m[4] ? Number(m[4]) - 1 : r1, c2: m[3] ? col(m[3]) : c1 };
}
