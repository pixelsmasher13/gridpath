/**
 * What the viewport grid needs from a calculation engine. One instance per
 * workbook. The desktop app implements it over its Tauri engine commands
 * (src/screens/SpreadsheetScreen/desktopGridEngine.ts); the package
 * implements it over the WASM engine in a Web Worker (browserEngine.ts).
 *
 * Rows and columns are 1-based unless a parameter says otherwise, as in
 * IronCalc. Writes are fire-and-forget and applied in call order; `settled`
 * resolves once everything queued before it has run and its recomputed
 * values have reached the grid through the registered applier.
 */

/** Writes recomputed values into the grid (0-based cells). Returns false
 * when the sheet is unknown to the grid; the engine keeps those values and
 * offers them again on the next registration. */
export type DeltaApplier = (sheetName: string, cells: { row0: number; col0: number; v: unknown }[]) => boolean;

export type FindHit = { sheet: number; row: number; column: number; text: string };
import type { NativeChart } from "./charts/types";

export interface GridEngine {
  /** Native charts, with values from the latest completed calculation. */
  charts?(): NativeChart[];
  subscribeCharts?(listener: () => void): () => void;
  /** Load the workbook and return its layout (geometry, merges, frozen
   * panes, styles); null when the load fails. The first evaluate is
   * deferred until `startDeferredSettle`. `bytes` are the file's bytes; an
   * engine that reads `path` itself may ignore them. */
  openLayout(path: string, bytes: Uint8Array): Promise<any | null>;
  /** Start the deferred first evaluate (idempotent). */
  startDeferredSettle(): void;
  /** The layout again, after everything queued before it. */
  fetchLayout(): Promise<any | null>;
  /** One rectangle of cells, inclusive bounds; null on failure. */
  fetchWindow(sheet: number, r0: number, r1: number, c0: number, c1: number): Promise<any | null>;
  /** Live-vs-file divergence report, or null when unavailable. */
  calcHealth(): Promise<any | null>;
  /** A partial style on a 0-based inclusive rectangle. */
  setStyle(sheetName: string, r0: number, c0: number, r1: number, c1: number, patch: Record<string, unknown>): Promise<boolean>;
  defineName(name: string, formula: string): Promise<boolean>;
  find(query: string, sheet?: number, limit?: number): Promise<FindHit[]>;
  /** Sheet names in engine order. */
  sheetNames(): string[];
  /** Cell inputs ("=A1*2", "42", "hello"; "" clears). */
  writeCells(sheetName: string, label: string, cells: Array<{ row: number; column: number; value: string }>): void;
  sheetOp(op: "add" | "rename" | "delete", name: string, newName?: string): void;
  /** insert_rows / delete_rows / insert_columns / delete_columns at a 1-based index. */
  structural(op: string, sheetName: string, index1: number, count: number): void;
  settled(timeoutMs?: number): Promise<void>;
  registerApplier(fn: DeltaApplier): void;
  unregisterApplier(): void;
}
