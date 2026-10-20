// Workbooks on disk, a folder per user: the storage behind the example's
// API. Swap these bodies for your database or object storage (S3, R2,
// GCS…); the route handler stays the same.
import { promises as fs } from "node:fs";
import path from "node:path";
import { randomUUID } from "node:crypto";

const ROOT = path.join(process.cwd(), ".data", "books");
/** Versions kept per workbook, and the window in which saves fold into the latest. */
const KEEP_VERSIONS = 20;
const FOLD_MS = 5 * 60_000;

export type Meta = { id: string; name: string; savedAt: string; source?: string; size: number };

const enc = (s: string) => Buffer.from(s).toString("base64url");
const folder = (user: string, id: string) => path.join(ROOT, enc(user), enc(id));

/** Written beside, then renamed: a reader never sees half a file. */
async function writeAtomic(file: string, data: Uint8Array | string) {
  const tmp = `${file}.${randomUUID()}.tmp`;
  await fs.writeFile(tmp, data);
  await fs.rename(tmp, file);
}

const missing = (e: unknown) => (e as NodeJS.ErrnoException).code === "ENOENT";

async function versionFiles(dir: string): Promise<Array<{ ms: number; file: string; size: number }>> {
  const names = await fs.readdir(path.join(dir, "versions")).catch((e) => (missing(e) ? [] : Promise.reject(e)));
  const out = await Promise.all(
    names
      .filter((n) => /^\d+\.xlsx$/.test(n))
      .map(async (file) => ({ ms: Number(file.slice(0, -5)), file, size: (await fs.stat(path.join(dir, "versions", file))).size })),
  );
  return out.sort((a, b) => b.ms - a.ms);
}

export async function save(user: string, id: string, w: { name: string; bytes: Uint8Array; source?: string; checkpoint?: boolean }) {
  const dir = folder(user, id);
  await fs.mkdir(path.join(dir, "versions"), { recursive: true });
  const meta: Meta = { id, name: w.name, savedAt: new Date().toISOString(), source: w.source, size: w.bytes.length };
  await writeAtomic(path.join(dir, "current.xlsx"), w.bytes);
  await writeAtomic(path.join(dir, "meta.json"), JSON.stringify(meta));
  // A version per save; saves within a few minutes of the latest replace it
  // (a checkpoint, before and after a restore, always gets its own).
  const [latest] = await versionFiles(dir);
  if (!w.checkpoint && latest && Date.now() - latest.ms < FOLD_MS) await fs.rm(path.join(dir, "versions", latest.file), { force: true });
  await fs.writeFile(path.join(dir, "versions", `${Date.now()}.xlsx`), w.bytes);
  for (const old of (await versionFiles(dir)).slice(KEEP_VERSIONS)) await fs.rm(path.join(dir, "versions", old.file), { force: true });
}

export async function load(user: string, id: string): Promise<{ meta: Meta; bytes: Uint8Array } | null> {
  const dir = folder(user, id);
  try {
    const meta = JSON.parse(await fs.readFile(path.join(dir, "meta.json"), "utf8")) as Meta;
    return { meta, bytes: await fs.readFile(path.join(dir, "current.xlsx")) };
  } catch (e) {
    if (missing(e)) return null;
    throw e;
  }
}

export async function list(user: string): Promise<Meta[]> {
  const root = path.join(ROOT, enc(user));
  const dirs = await fs.readdir(root).catch((e) => (missing(e) ? [] : Promise.reject(e)));
  const metas = await Promise.all(
    dirs.map((d) => fs.readFile(path.join(root, d, "meta.json"), "utf8").then((t) => JSON.parse(t) as Meta, () => null)),
  );
  return metas.filter((m): m is Meta => !!m).sort((a, b) => (a.savedAt < b.savedAt ? 1 : -1));
}

export async function versions(user: string, id: string) {
  return (await versionFiles(folder(user, id))).map((v) => ({ id: String(v.ms), savedAt: new Date(v.ms).toISOString(), size: v.size }));
}

export async function loadVersion(user: string, id: string, versionId: string): Promise<Uint8Array | null> {
  if (!/^\d+$/.test(versionId)) return null;
  return fs.readFile(path.join(folder(user, id), "versions", `${versionId}.xlsx`)).catch((e) => (missing(e) ? null : Promise.reject(e)));
}

export async function remove(user: string, id: string) {
  await fs.rm(folder(user, id), { recursive: true, force: true });
}
