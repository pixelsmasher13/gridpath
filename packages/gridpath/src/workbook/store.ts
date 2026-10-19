/**
 * Pending batches on disk, next to the workbook:
 *
 *   <dir>/.gridpath/<file>.batches.json
 *
 * So the review link outlives the agent's session, a later `gridpath review`
 * can pick the batches up, and nothing about pending state lives only in
 * one process's memory. The .xlsx itself is untouched until save.
 */
import fs from "node:fs";
import path from "node:path";
import type { ChangeBatch } from "../core/types";

export type StoredState = {
  /** Absolute workbook path the batches apply to. */
  path: string;
  /** Size + mtime of the .xlsx the batches were built against. */
  base: { size: number; mtimeMs: number };
  batches: ChangeBatch[];
};

export function storePathFor(workbookPath: string): string {
  const dir = path.join(path.dirname(workbookPath), ".gridpath");
  return path.join(dir, `${path.basename(workbookPath)}.batches.json`);
}

export function baseStamp(workbookPath: string): { size: number; mtimeMs: number } {
  const st = fs.statSync(workbookPath);
  return { size: st.size, mtimeMs: st.mtimeMs };
}

export function loadStored(workbookPath: string): StoredState | null {
  const p = storePathFor(workbookPath);
  if (!fs.existsSync(p)) return null;
  try {
    const s = JSON.parse(fs.readFileSync(p, "utf8")) as StoredState;
    const now = baseStamp(workbookPath);
    // The file changed underneath (saved by us, or edited elsewhere): the
    // stored batches no longer apply. Drop them rather than replay blindly.
    if (s.base.size !== now.size || Math.abs(s.base.mtimeMs - now.mtimeMs) > 1) {
      fs.rmSync(p, { force: true });
      return null;
    }
    return s;
  } catch {
    return null;
  }
}

export function saveStored(workbookPath: string, batches: ChangeBatch[]): void {
  const p = storePathFor(workbookPath);
  if (batches.length === 0) {
    fs.rmSync(p, { force: true });
    return;
  }
  fs.mkdirSync(path.dirname(p), { recursive: true });
  const state: StoredState = { path: workbookPath, base: baseStamp(workbookPath), batches };
  const tmp = `${p}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(state));
  fs.renameSync(tmp, p);
}

export function clearStored(workbookPath: string): void {
  fs.rmSync(storePathFor(workbookPath), { force: true });
}
