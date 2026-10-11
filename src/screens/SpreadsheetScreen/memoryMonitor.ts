/**
 * Memory monitor. The calc engine runs in the Rust process and is not what
 * runs out of memory on huge workbooks — WebKit's content process is, and
 * macOS terminates it under pressure. This polls the Rust side for both
 * resident sizes (`memory_report`), logs them on the `[memory]` prefix so
 * every open and every agent turn can be read against real numbers, and
 * raises one warning when the content process passes a threshold.
 */
import { invoke } from "@tauri-apps/api/core";

export type MemoryReport = {
  app_rss: number;
  webview_pid: number | null;
  webview_rss: number | null;
};

export const GB = 1024 ** 3;

/** Above this the content process is in the zone where macOS kills it. */
// 3.5 GB: a 13 MB / 2.3M-cell workbook sits at ~2.8 GB with the Univer
// calc worker on and works fine, so the toast must not fire on a normal
// open. We have not yet measured the level at which WebKit's content
// process actually gets killed; revisit once the monitor logs catch one.
export const WEBVIEW_WARN_BYTES = 3.5 * GB;

export async function readMemory(): Promise<MemoryReport | null> {
  try {
    return await invoke<MemoryReport>("memory_report");
  } catch (e) {
    console.warn("[memory] report failed:", e);
    return null;
  }
}

export function formatBytes(n: number | null | undefined): string {
  if (n === null || n === undefined) return "n/a";
  return n >= GB ? `${(n / GB).toFixed(2)} GB` : `${Math.round(n / 1024 ** 2)} MB`;
}

export function describeMemory(r: MemoryReport): string {
  return `webview ${formatBytes(r.webview_rss)}${r.webview_pid ? ` (pid ${r.webview_pid})` : ""} · engine/app ${formatBytes(r.app_rss)}`;
}

/** One-off reading, labelled, for before/after comparisons in the console. */
export async function logMemory(label: string): Promise<MemoryReport | null> {
  const r = await readMemory();
  if (r) console.log(`[memory] ${label}: ${describeMemory(r)}`);
  return r;
}

export type MemoryMonitorOptions = {
  intervalMs?: number;
  warnBytes?: number;
  /** Called once when the webview crosses `warnBytes`; re-armed after it
   * drops back under 80% of the threshold. */
  onWarn?: (report: MemoryReport) => void;
  onReport?: (report: MemoryReport) => void;
};

/** Start polling; returns a stop function. */
export function startMemoryMonitor(opts: MemoryMonitorOptions = {}): () => void {
  const intervalMs = opts.intervalMs ?? 10_000;
  const warnBytes = opts.warnBytes ?? WEBVIEW_WARN_BYTES;
  let armed = true;
  let stopped = false;
  const tick = async () => {
    if (stopped) return;
    const r = await readMemory();
    if (!r || stopped) return;
    opts.onReport?.(r);
    console.log(`[memory] ${describeMemory(r)}`);
    const wv = r.webview_rss ?? 0;
    if (armed && wv > warnBytes) {
      armed = false;
      opts.onWarn?.(r);
    } else if (!armed && wv < warnBytes * 0.8) {
      armed = true;
    }
  };
  const id = setInterval(() => void tick(), intervalMs);
  void tick();
  return () => {
    stopped = true;
    clearInterval(id);
  };
}
