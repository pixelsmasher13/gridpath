/**
 * IronCalc shadow engine (spike; enable with
 * `localStorage.setItem("gridpath.ironcalcShadow", "1")` + reload).
 *
 * Runs the vendored IronCalc model in the Tauri process ALONGSIDE Univer:
 * the workbook file is loaded into it on open, every committed value edit is
 * mirrored via `calc_engine_edit`, and the returned dirty closure is used to
 * (a) log incremental-recalc latency next to Univer's own `[calc-perf]`
 * lines and (b) cross-check computed values against Univer's engine a couple
 * of seconds later. Nothing the user sees is driven by IronCalc yet — this
 * exists to validate the engine on real workbooks and real edits before any
 * swap. All output is on the `[ironcalc-shadow]` console prefix.
 *
 * Only plain value/formula cell edits are mirrored. Structural or bulk ops
 * (insert/remove rows, paste, sort, clear…) put the shadow into a "desynced"
 * state so its numbers stop being reported as trustworthy.
 */
import { invoke } from "@tauri-apps/api/core";

interface ChangedCell {
  sheet: number;
  row: number; // 1-based (IronCalc)
  column: number; // 1-based
  value: unknown;
  formatted: string;
}

interface CalcEditResult {
  recalc_ms: number;
  changed: ChangedCell[];
}

interface CalcLoadResult {
  sheets: string[];
  load_ms: number;
  eval_ms: number;
}

type UniverValueGetter = (sheetName: string, row0: number, col0: number) => unknown;

let sheets: string[] = [];
let loadedPath: string | null = null;
let shadowTabId: string | null = null;
let loading = false;
let desynced: string | null = null;
let univerValueGetter: UniverValueGetter | null = null;
let univerIdleWaiter: ((timeoutMs: number) => Promise<void>) | null = null;

export function ironcalcShadowEnabled(): boolean {
  // Engine mode supersedes the shadow: IronCalc drives the grid directly and
  // there is no independent Univer calculation left to compare against. The
  // engine is ON by default now — the shadow only applies when the engine
  // kill switch ("0") is set AND the shadow flag is on.
  try {
    if (localStorage.getItem("gridpath.ironcalcEngine") !== "0") return false;
  } catch {
    return false;
  }
  try {
    if (localStorage.getItem("gridpath.ironcalcShadow") === "1") return true;
  } catch {
    /* no localStorage (tests) */
  }
  // Env override so eval / CI runs can turn the shadow on without a browser.
  try {
    return (import.meta as any).env?.VITE_IRONCALC_SHADOW === "1";
  } catch {
    return false;
  }
}

/** UniverGrid registers a facade-backed value reader for the cross-check. */
export function registerUniverValueGetter(fn: UniverValueGetter): void {
  univerValueGetter = fn;
}

/** UniverGrid registers "resolve when the Univer engine is idle". */
export function registerUniverIdleWaiter(fn: (timeoutMs: number) => Promise<void>): void {
  univerIdleWaiter = fn;
}

export async function ironcalcShadowLoad(path: string, tabId: string): Promise<void> {
  if (!ironcalcShadowEnabled()) return;
  if (!path || path.startsWith("untitled-") || !/\.xls[xm]$/i.test(path)) return;
  loading = true;
  desynced = null;
  shadowTabId = tabId;
  try {
    const t0 = performance.now();
    const res = await invoke<CalcLoadResult>("calc_engine_load", { tab: tabId, path });
    sheets = res.sheets;
    loadedPath = path;
    console.log(
      `[ironcalc-shadow] loaded ${path}: ${res.sheets.length} sheets — ` +
        `import ${res.load_ms}ms, full calc ${res.eval_ms}ms ` +
        `(invoke round-trip ${(performance.now() - t0).toFixed(0)}ms)`,
    );
  } catch (e) {
    loadedPath = null;
    console.warn("[ironcalc-shadow] load failed:", e);
  } finally {
    loading = false;
  }
}

function markDesynced(reason: string): void {
  if (!loadedPath || desynced) return;
  desynced = reason;
  console.warn(
    `[ironcalc-shadow] desynced (${reason}) — shadow values untrustworthy until next load`,
  );
}

/**
 * Feed one committed Univer command into the shadow. Call from the
 * `onCommandExecuted` stream; cheap no-op when the shadow is disabled/idle.
 */
export function ironcalcShadowOnCommand(
  cmd: { id?: unknown; params?: unknown },
  sheetNamesById: Map<string, string>,
): void {
  if (!loadedPath && !loading) return;
  if (!cmd || typeof cmd.id !== "string" || !cmd.id.startsWith("sheet.command.")) return;

  if (cmd.id === "sheet.command.set-range-values") {
    const p = (cmd.params ?? {}) as Record<string, unknown>;
    const subUnitId = p.subUnitId;
    const sheetName =
      (typeof subUnitId === "string" ? sheetNamesById.get(subUnitId) : undefined) ?? null;
    const cells = extractEditedCells(p);
    if (!sheetName || cells === null) {
      markDesynced("unparsed set-range-values params");
      return;
    }
    for (const c of cells) {
      void shadowEdit(sheetName, c.row0, c.col0, c.input);
    }
    return;
  }

  // Structural / bulk operations we do not mirror (formatting-only commands
  // are irrelevant to calculation and simply ignored).
  if (DESYNC_HINTS.some((h) => (cmd.id as string).includes(h))) {
    markDesynced(cmd.id);
  }
}

const DESYNC_HINTS = [
  "insert-row",
  "insert-col",
  "remove-row",
  "remove-col",
  "append-row",
  "move-rows",
  "move-cols",
  "insert-range-move",
  "delete-range-move",
  "clear-selection",
  "paste",
  "sort-range",
  "worksheet-merge",
];

export interface EditedCell {
  row0: number;
  col0: number;
  /** Spreadsheet input syntax: "=SUM(A1:B2)", "42", "text", "" clears. */
  input: string;
}

/** Univer ICellData → what a user would have typed into the cell. */
export function cellDataToInput(cell: unknown): string | null {
  if (cell === null || cell === undefined) return "";
  if (typeof cell !== "object") return null;
  const c = cell as Record<string, unknown>;
  if (typeof c.f === "string" && c.f.length > 0) return c.f;
  if (typeof c.si === "string") return null; // shared-formula ref without text
  if (c.v === null || c.v === undefined) {
    // Style-only patch ({s: ...}): not a value edit.
    return Object.keys(c).some((k) => k === "v" || k === "f") ? "" : null;
  }
  return String(c.v);
}

/**
 * Decode the `value` shapes SetRangeValuesCommand uses: a row/col matrix
 * (`{5: {2: cell}}`) or a single ICellData applied to `range`. Returns null
 * when the shape is unrecognized or too large to mirror cell-by-cell.
 *
 * `maxCells` defaults to shadow mode's sampling budget; ENGINE mode passes
 * its own (much higher) mirroring limit — a shadow-sized cap there rejected
 * any 400+-cell agent write and froze the engine for the session.
 */
export function extractEditedCells(
  p: Record<string, unknown>,
  maxCells = 400,
): EditedCell[] | null {
  const MAX_CELLS = maxCells;
  const value = p.value ?? p.cellValue;
  if (value === null || value === undefined) return null;

  const isIndexMatrix = (o: unknown): o is Record<string, Record<string, unknown>> =>
    typeof o === "object" &&
    o !== null &&
    Object.keys(o as object).length > 0 &&
    Object.keys(o as object).every((k) => /^\d+$/.test(k));

  if (isIndexMatrix(value) && Object.values(value).every((row) => isIndexMatrix(row) || row === null)) {
    const out: EditedCell[] = [];
    for (const [r, row] of Object.entries(value)) {
      if (!row) continue;
      for (const [c, cell] of Object.entries(row)) {
        const input = cellDataToInput(cell);
        if (input === null) return null;
        out.push({ row0: Number(r), col0: Number(c), input });
        if (out.length > MAX_CELLS) return null;
      }
    }
    return out;
  }

  // Single ICellData over a range (or over the current selection cell).
  const range = p.range as Record<string, unknown> | undefined;
  const input = cellDataToInput(value);
  if (input === null) return null;
  if (
    range &&
    typeof range.startRow === "number" &&
    typeof range.startColumn === "number" &&
    typeof range.endRow === "number" &&
    typeof range.endColumn === "number"
  ) {
    const area = (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1);
    if (area > MAX_CELLS) return null;
    const out: EditedCell[] = [];
    for (let r = range.startRow as number; r <= (range.endRow as number); r++) {
      for (let c = range.startColumn as number; c <= (range.endColumn as number); c++) {
        out.push({ row0: r, col0: c, input });
      }
    }
    return out;
  }
  return null;
}

async function shadowEdit(sheetName: string, row0: number, col0: number, input: string) {
  const sheet = sheets.indexOf(sheetName);
  if (sheet < 0) {
    markDesynced(`unknown sheet "${sheetName}"`);
    return;
  }
  try {
    const t0 = performance.now();
    if (!shadowTabId) return;
    const res = await invoke<CalcEditResult>("calc_engine_edit", {
      tab: shadowTabId,
      sheet,
      row: row0 + 1,
      column: col0 + 1,
      value: input,
    });
    console.log(
      `[ironcalc-shadow] edit ${sheetName}!r${row0}c${col0} → ` +
        `recalc ${res.recalc_ms}ms (round-trip ${(performance.now() - t0).toFixed(0)}ms), ` +
        `${res.changed.length} cells changed${desynced ? " [DESYNCED]" : ""}`,
    );
    if (!desynced) scheduleCrossCheck(res.changed);
  } catch (e) {
    console.warn("[ironcalc-shadow] edit failed:", e);
  }
}

/**
 * After Univer's own recalc has finished (5-15s on heavy files — waiting is
 * essential, or every "mismatch" is just a stale Univer value), compare a
 * sample of the changed cells against Univer's values.
 */
function scheduleCrossCheck(changed: ChangedCell[]): void {
  if (!univerValueGetter) return;
  const sample = changed.slice(0, 40);
  void (async () => {
    const t0 = performance.now();
    if (univerIdleWaiter) {
      await univerIdleWaiter(30_000);
    } else {
      await new Promise((r) => setTimeout(r, 2500));
    }
    await new Promise((r) => setTimeout(r, 150));
    const waitedMs = (performance.now() - t0).toFixed(0);
    if (desynced || !univerValueGetter) return;
    let agree = 0;
    const mismatches: string[] = [];
    for (const c of sample) {
      const sheetName = sheets[c.sheet];
      if (!sheetName) continue;
      const uv = univerValueGetter(sheetName, c.row - 1, c.column - 1);
      if (valuesAgree(uv, c.value)) {
        agree++;
      } else if (mismatches.length < 5) {
        mismatches.push(`${sheetName}!r${c.row - 1}c${c.column - 1}: univer=${JSON.stringify(uv)} ironcalc=${JSON.stringify(c.value)}`);
      }
    }
    console.log(
      `[ironcalc-shadow] cross-check (univer idle after ${waitedMs}ms): ${agree}/${sample.length} cells agree`,
    );
    if (mismatches.length > 0) console.warn("[ironcalc-shadow] mismatches:", mismatches);
  })();
}

/** "4,624.1" | "$70.00" | "(65.2%)" | "29.3%" → number; null if not numeric. */
export function parseFormattedNumber(x: unknown): number | null {
  if (typeof x === "number") return x;
  if (typeof x !== "string") return null;
  let s = x.trim();
  let negative = false;
  const paren = /^\((.*)\)$/.exec(s);
  if (paren) {
    negative = true;
    s = paren[1];
  }
  let percent = false;
  if (s.endsWith("%")) {
    percent = true;
    s = s.slice(0, -1);
  }
  s = s.replace(/[$\u20ac\u00a3,\s]/g, "");
  if (!/^-?\d*\.?\d+(e[+-]?\d+)?$/i.test(s)) return null;
  let n = Number(s);
  if (negative) n = -n;
  if (percent) n /= 100;
  return n;
}

export function valuesAgree(univer: unknown, ironcalc: unknown): boolean {
  if (univer === null || univer === undefined || univer === "") {
    return ironcalc === null || ironcalc === "" || ironcalc === 0;
  }
  if (typeof univer === "number" && typeof ironcalc === "number") {
    const denom = Math.max(Math.abs(univer), Math.abs(ironcalc), 1e-12);
    return Math.abs(univer - ironcalc) / denom < 1e-9;
  }
  if (typeof univer === "boolean" || typeof ironcalc === "boolean") {
    return univer === ironcalc;
  }
  // One side is a number-formatted display string ("4,624.1", "$70.00",
  // "(65.2%)"): compare loosely, since display rounding hides digits.
  const nu = parseFormattedNumber(univer);
  const ni = parseFormattedNumber(ironcalc);
  if (nu !== null && ni !== null) {
    const denom = Math.max(Math.abs(nu), Math.abs(ni), 1e-9);
    return Math.abs(nu - ni) / denom < 5e-3;
  }
  return String(univer) === String(ironcalc);
}
