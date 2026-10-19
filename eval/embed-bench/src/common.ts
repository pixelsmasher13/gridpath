// The page side of the benchmark protocol (see run.mjs).
export type Prep = {
  id: string;
  input: { sheet: string; ref: string; original: number; changed: number };
  dep: { sheet: string; ref: string };
  sheets: Array<{ name: string; cells: number[] }>;
};

export interface Bench {
  phase: "loading" | "interactive" | "calculated" | "error";
  error?: string;
  t: { start?: number; interactive?: number; calculated?: number };
  /** Set the input, resolve with ms until the watched dependent changes. */
  edit?: (value: number) => Promise<number>;
  /** Values at the prep's formula positions, per sheet (null for GridPath: scored in Node). */
  dump?: () => Promise<unknown[][] | null>;
  /** The saved .xlsx, base64. */
  save?: () => Promise<string>;
}

export const B: Bench = ((window as any).__bench = { phase: "loading", t: {} });
window.addEventListener("error", (e) => fail(e.error ?? e.message));
window.addEventListener("unhandledrejection", (e) => fail(e.reason));

export function fail(e: unknown) {
  B.phase = "error";
  B.error = e instanceof Error ? `${e.message}\n${e.stack}` : String(e);
}

export async function inputs(): Promise<{ id: string; bytes: Uint8Array; prep: Prep }> {
  const id = new URLSearchParams(location.search).get("file")!;
  const [bytes, prep] = await Promise.all([
    fetch(`/files/${id}.xlsx`).then((r) => r.arrayBuffer()).then((b) => new Uint8Array(b)),
    fetch(`/prep/${id}.json`).then((r) => r.json()),
  ]);
  return { id, bytes, prep };
}

/** "B7" → 1-based { row: 7, col: 2 }. */
export function ref(a1: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]), col };
}

export const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Poll `read` until it differs from `before`; ms since `t0`. */
export async function until(read: () => unknown | Promise<unknown>, before: unknown, t0: number, timeoutMs = 60_000): Promise<number> {
  for (;;) {
    const v = await read();
    if (v !== before && !(typeof v === "number" && typeof before === "number" && Math.abs(v - before) < 1e-12)) return performance.now() - t0;
    if (performance.now() - t0 > timeoutMs) throw new Error(`dependent did not change within ${timeoutMs} ms (stays ${JSON.stringify(v)})`);
    await sleep(0);
  }
}

export function toBase64(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** Excel serial for a Date produced from a UTC-based reader (ExcelJS). */
export const serialUtc = (d: Date) => (d.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
