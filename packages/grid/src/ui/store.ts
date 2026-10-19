/**
 * Where saved workbooks live: the current state of each workbook, its
 * version history, and the recent list. `browserStore()` keeps them in this
 * browser (IndexedDB); a host with its own backend implements the interface.
 */

/** A saved workbook, as the recent list shows it. */
export interface StoredWorkbook {
  id: string;
  name: string;
  /** ISO time of the last save. */
  savedAt: string;
  size: number;
}

/** One entry of a workbook's version history. */
export interface WorkbookVersion {
  id: string;
  savedAt: string;
  size: number;
}

export interface WorkbookStore {
  /** Save a workbook's current bytes; the store also keeps a version.
   * `source` fingerprints the file it was opened from (see `load`);
   * `checkpoint` asks for a version of its own (before and after a restore),
   * where a store may otherwise fold frequent saves into one. */
  save(workbook: { id: string; name: string; bytes: Uint8Array; source?: string; checkpoint?: boolean }): Promise<void>;
  /** The workbook's latest saved state, or null. */
  load(id: string): Promise<{ id: string; name: string; bytes: Uint8Array; savedAt: string; source?: string } | null>;
  /** Saved workbooks, most recent first. */
  list(): Promise<StoredWorkbook[]>;
  /** A workbook's versions, newest first. */
  versions(id: string): Promise<WorkbookVersion[]>;
  loadVersion(id: string, versionId: string): Promise<Uint8Array | null>;
  /** Forget a workbook and its versions. */
  remove?(id: string): Promise<void>;
}

const DB = "gridpath";

export interface BrowserStoreOptions {
  /** A save within this many minutes of the latest version replaces it
   * instead of adding one (default 5; 0 keeps every save). */
  checkpointMinutes?: number;
  /** Versions kept as they are (default 20); older ones thin out to the
   * last of each day for `keepDays` days (default 30). */
  keepRecent?: number;
  keepDays?: number;
}

type DocRow = { id: string; name: string; bytes: Uint8Array; savedAt: string; source?: string; size: number };
type VersionRow = { key?: number; workbookId: string; savedAt: string; bytes: Uint8Array; size: number };

function request<T>(r: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}

function done(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error ?? new Error("transaction aborted"));
  });
}

/** Workbooks saved in this browser (IndexedDB database "gridpath"). */
export function browserStore(options: BrowserStoreOptions = {}): WorkbookStore {
  const CHECKPOINT_MS = Math.max(0, options.checkpointMinutes ?? 5) * 60_000;
  const KEEP_RECENT = Math.max(1, options.keepRecent ?? 20);
  const KEEP_DAYS = Math.max(0, options.keepDays ?? 30);
  let opening: Promise<IDBDatabase> | null = null;
  const db = () =>
    (opening ??= new Promise<IDBDatabase>((resolve, reject) => {
      if (typeof indexedDB === "undefined") return reject(new Error("this browser has no IndexedDB"));
      const r = indexedDB.open(DB, 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore("workbooks", { keyPath: "id" });
        d.createObjectStore("versions", { keyPath: "key", autoIncrement: true }).createIndex("byWorkbook", "workbookId");
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    }));

  const versionRows = async (id: string): Promise<VersionRow[]> => {
    const tx = (await db()).transaction("versions", "readonly");
    const rows = (await request(tx.objectStore("versions").index("byWorkbook").getAll(id))) as VersionRow[];
    return rows.sort((a, b) => (a.savedAt < b.savedAt ? 1 : -1));
  };

  return {
    async save({ id, name, bytes, source, checkpoint }) {
      const now = new Date().toISOString();
      const rows = await versionRows(id);
      const tx = (await db()).transaction(["workbooks", "versions"], "readwrite");
      const copy = bytes.slice();
      tx.objectStore("workbooks").put({ id, name, bytes: copy, savedAt: now, source, size: copy.byteLength } satisfies DocRow);
      const versions = tx.objectStore("versions");
      const latest = rows[0];
      const fold = !checkpoint && !!latest && Date.now() - Date.parse(latest.savedAt) < CHECKPOINT_MS;
      if (fold) {
        versions.put({ ...latest, savedAt: now, bytes: copy, size: copy.byteLength });
      } else {
        versions.add({ workbookId: id, savedAt: now, bytes: copy, size: copy.byteLength } satisfies VersionRow);
      }
      // Thin out: the newest KEEP_RECENT as they are, then the last of each
      // day for KEEP_DAYS days.
      const older = fold ? rows.slice(1) : rows;
      const keepDays = new Set<string>();
      older.forEach((v, i) => {
        if (i < KEEP_RECENT - 1) return;
        const day = v.savedAt.slice(0, 10);
        const ageDays = (Date.now() - Date.parse(v.savedAt)) / 86_400_000;
        if (ageDays <= KEEP_DAYS && !keepDays.has(day)) keepDays.add(day);
        else if (v.key !== undefined) versions.delete(v.key);
      });
      await done(tx);
    },
    async load(id) {
      const tx = (await db()).transaction("workbooks", "readonly");
      const row = (await request(tx.objectStore("workbooks").get(id))) as DocRow | undefined;
      return row ? { id: row.id, name: row.name, bytes: row.bytes, savedAt: row.savedAt, source: row.source } : null;
    },
    async list() {
      const tx = (await db()).transaction("workbooks", "readonly");
      const rows = (await request(tx.objectStore("workbooks").getAll())) as DocRow[];
      return rows
        .map(({ id, name, savedAt, size }) => ({ id, name, savedAt, size }))
        .sort((a, b) => (a.savedAt < b.savedAt ? 1 : -1));
    },
    async versions(id) {
      return (await versionRows(id)).map((v) => ({ id: String(v.key), savedAt: v.savedAt, size: v.size }));
    },
    async loadVersion(id, versionId) {
      const tx = (await db()).transaction("versions", "readonly");
      const row = (await request(tx.objectStore("versions").get(Number(versionId)))) as VersionRow | undefined;
      return row && row.workbookId === id ? row.bytes : null;
    },
    async remove(id) {
      const rows = await versionRows(id);
      const tx = (await db()).transaction(["workbooks", "versions"], "readwrite");
      tx.objectStore("workbooks").delete(id);
      for (const v of rows) if (v.key !== undefined) tx.objectStore("versions").delete(v.key);
      await done(tx);
    },
  };
}

/** A cheap fingerprint of a file's bytes (FNV-1a, two lanes, plus length):
 * tells "the same file the stored workbook came from" from a new one. */
export function fingerprint(bytes: Uint8Array): string {
  let a = 0x811c9dc5;
  let b = 0x01000193 ^ bytes.length;
  for (let i = 0; i < bytes.length; i++) {
    a = Math.imul(a ^ bytes[i], 0x01000193);
    b = Math.imul(b ^ bytes[bytes.length - 1 - i], 0x01000193);
  }
  return `${bytes.length.toString(36)}-${(a >>> 0).toString(36)}${(b >>> 0).toString(36)}`;
}
