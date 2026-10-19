/**
 * run_script — sandboxed JavaScript execution against the workbook.
 *
 * The agent's JSON tools (set_cell / set_range / …) are the right grain for
 * small precise edits, but loop-shaped work — per-row conditional rewrites,
 * repeating formula patterns, address arithmetic across many rows — costs a
 * full model turn per tool call and drifts out of alignment mid-payload.
 * `run_script` lets the model send ONE short program instead; the program's
 * writes are recorded as ordinary UniverMutations and fall through the same
 * apply pipeline as every other tool (diff tinting, per-cell reject, undo,
 * save mirror), so nothing downstream knows a script was involved.
 *
 * Execution model:
 *  - The script runs against a plain-data READ MODEL built from the Univer
 *    snapshot (evaluated values + formulas), never against the live grid.
 *    Reads see the pre-script workbook overlaid with the script's own
 *    literal writes; formulas the script writes are NOT evaluated until the
 *    mutations land in Univer afterwards.
 *  - It runs in a throwaway Blob Worker so a runaway loop can be terminated
 *    with a hard timeout instead of wedging the UI thread. `scriptCore` is
 *    a fully self-contained function whose SOURCE is serialized into the
 *    worker — it must not reference anything outside its own body.
 *  - No network, no imports: the worker preamble neuters fetch/XHR/sockets.
 *    `fetch_web` is the sanctioned way for the agent to pull external data.
 */
import type { UniverMutation } from "./types";
import { expandA1Range } from "./toolToMutation";

/** One write recorded by the script, in tool coordinates (0-indexed). */
export type ScriptOp =
  | { kind: "set"; sheet: string; row: number; col: number; value: string | number | boolean | null }
  | { kind: "format"; sheet: string; range: string; format: Record<string, unknown> }
  | { kind: "clear"; sheet: string; range: string };

/** Sparse per-sheet cell data keyed "row,col" (0-indexed). */
export type ScriptSheetData = Record<string, { v: string | number | boolean | null; f: string | null }>;

export type ScriptModel = {
  /** Sheet names in tab order. */
  sheetNames: string[];
  sheets: Record<string, ScriptSheetData>;
  /** Sheets whose cells were NOT loaded into this model (engine-sourced
   * runs preload only the sheets the script names). Reading one throws an
   * agent-readable error instead of returning blanks; writing is fine. */
  unloaded?: string[];
};

/** How much engine JSON a single script run may pull into its sandbox. */
export const SCRIPT_MAX_SOURCE_BYTES = 400 * 1024 * 1024;

export type ScriptLimits = {
  /** Cap on cells touched by set/setValues/format/clear combined. */
  maxTouched: number;
  /** Cap on log() entries retained. */
  maxLogs: number;
  maxLogChars: number;
  maxLogBytes: number;
};

export type ScriptCoreResult =
  | { ok: true; ops: ScriptOp[]; logs: string[] }
  | { ok: false; error: string; logs: string[] };

export const SCRIPT_MAX_TOUCHED = 20_000;
export const SCRIPT_MAX_LOGS = 100;
/** Per-entry cap; a cut entry ends with "… [+N chars]" so the model knows. */
export const SCRIPT_MAX_LOG_CHARS = 4_000;
/** Total log budget across entries; one final entry says when it was hit. */
export const SCRIPT_MAX_LOG_BYTES = 64 * 1024;
export const SCRIPT_TIMEOUT_MS = 5_000;
/** Backstop against pathological snapshots — structured-cloning a multi-
 * million-cell model into the worker would stall the UI thread. */
const MAX_MODEL_CELLS = 1_000_000;

/**
 * Execute agent script `code` against `model`. SELF-CONTAINED ON PURPOSE:
 * this function's source is stringified into a Blob Worker, so it must not
 * close over imports or module state — every helper lives inside. The A1
 * helpers intentionally duplicate toolToMutation's; they can't be imported.
 */
export function scriptCore(payload: {
  code: string;
  model: ScriptModel;
  limits: ScriptLimits;
}): ScriptCoreResult {
  "use strict";
  const { code, model, limits } = payload;
  const logs: string[] = [];
  const ops: ScriptOp[] = [];
  let touched = 0;

  const parseAddr = (addr: unknown): { row: number; col: number } => {
    const m = /^\s*([A-Za-z]+)\s*(\d+)\s*$/.exec(String(addr ?? ""));
    if (!m) throw new Error(`invalid A1 address: ${JSON.stringify(addr)} (expected e.g. "B7")`);
    let col = 0;
    const letters = m[1].toUpperCase();
    for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
    const row = parseInt(m[2], 10);
    if (row < 1) throw new Error(`invalid A1 address: ${JSON.stringify(addr)}`);
    return { row: row - 1, col: col - 1 };
  };
  const parseRange = (range: unknown): { r0: number; c0: number; r1: number; c1: number } => {
    const parts = String(range ?? "").split(":");
    if (parts.length > 2) throw new Error(`invalid A1 range: ${JSON.stringify(range)}`);
    const a = parseAddr(parts[0]);
    const b = parts.length === 2 ? parseAddr(parts[1]) : a;
    return {
      r0: Math.min(a.row, b.row),
      c0: Math.min(a.col, b.col),
      r1: Math.max(a.row, b.row),
      c1: Math.max(a.col, b.col),
    };
  };
  const colLetters = (col: number): string => {
    let n = col;
    let s = "";
    do {
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return s;
  };
  const budget = (cells: number, what: string) => {
    touched += cells;
    if (touched > limits.maxTouched) {
      throw new Error(
        `write cap exceeded: this script touched more than ${limits.maxTouched} cells (at ${what}). ` +
          `Split the work into smaller scripts.`,
      );
    }
  };
  const FORMAT_KEYS = [
    "bold", "italic", "underline", "strike", "font_color", "background_color",
    "font_size", "font_family", "horizontal_align", "vertical_align",
    "wrap_text", "number_format",
  ];

  const handles: Record<string, unknown> = {};
  const makeHandle = (name: string) => {
    const data = model.sheets[name];
    const unloaded = Array.isArray(model.unloaded) && model.unloaded.indexOf(name) !== -1;
    const assertReadable = (what: string) => {
      if (unloaded) {
        throw new Error(
          `${what}: sheet "${name}" was not loaded for this script. Only sheets named as string literals in the ` +
            `script (e.g. sheet("${name}")) are loaded before it runs — name every sheet you read, or call sheets() ` +
            `in the script to load them all.`,
        );
      }
    };
    /** Script's own writes, overlaid on reads. */
    const written: Record<string, { v: string | number | boolean | null; f: string | null }> = {};
    const getCell = (addr: unknown): { value: string | number | boolean | null; formula: string | null } => {
      assertReadable(`get(${String(addr)})`);
      const { row, col } = parseAddr(addr);
      const w = written[`${row},${col}`];
      // A formula written by this script has no evaluated value yet — the
      // grid evaluates only after the mutations land.
      if (w) return { value: w.f ? null : w.v, formula: w.f };
      const c = data[`${row},${col}`];
      return c ? { value: c.v, formula: c.f } : { value: null, formula: null };
    };
    const setCell = (addr: unknown, value: unknown) => {
      const { row, col } = parseAddr(addr);
      const t = typeof value;
      if (value !== null && t !== "string" && t !== "number" && t !== "boolean") {
        throw new Error(
          `set(${JSON.stringify(String(addr))}): value must be a string, number, boolean, or null — got ${t}`,
        );
      }
      if (t === "number" && !Number.isFinite(value as number)) {
        throw new Error(`set(${JSON.stringify(String(addr))}): value is ${String(value)} — check your arithmetic`);
      }
      budget(1, `set(${colLetters(col)}${row + 1})`);
      const v = value as string | number | boolean | null;
      const f = typeof v === "string" && v.startsWith("=") ? v : null;
      written[`${row},${col}`] = { v: f ? null : v, f };
      ops.push({ kind: "set", sheet: name, row, col, value: v });
    };
    return {
      name,
      get: getCell,
      set: setCell,
      values: (range: unknown): (string | number | boolean | null)[][] => {
        assertReadable(`values(${String(range)})`);
        const { r0, c0, r1, c1 } = parseRange(range);
        const out: (string | number | boolean | null)[][] = [];
        for (let r = r0; r <= r1; r++) {
          const rowOut: (string | number | boolean | null)[] = [];
          for (let c = c0; c <= c1; c++) rowOut.push(getCell(`${colLetters(c)}${r + 1}`).value);
          out.push(rowOut);
        }
        return out;
      },
      setValues: (topLeft: unknown, rows: unknown) => {
        const origin = parseAddr(topLeft);
        if (!Array.isArray(rows) || rows.some((r) => !Array.isArray(r))) {
          throw new Error(`setValues: second argument must be a 2D array of rows`);
        }
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < rows[r].length; c++) {
            const v = rows[r][c];
            // Mirror set_range semantics: null / "" / undefined = preserve.
            if (v === null || v === undefined || v === "") continue;
            setCell(`${colLetters(origin.col + c)}${origin.row + r + 1}`, v);
          }
        }
      },
      format: (range: unknown, format: unknown) => {
        const { r0, c0, r1, c1 } = parseRange(range);
        if (typeof format !== "object" || format === null || Array.isArray(format)) {
          throw new Error(`format: second argument must be a format object, e.g. {bold: true}`);
        }
        for (const k of Object.keys(format)) {
          if (FORMAT_KEYS.indexOf(k) === -1) {
            throw new Error(`format: unknown key "${k}". Allowed: ${FORMAT_KEYS.join(", ")}`);
          }
        }
        budget((r1 - r0 + 1) * (c1 - c0 + 1), `format(${String(range)})`);
        ops.push({ kind: "format", sheet: name, range: String(range), format: { ...(format as object) } });
      },
      clear: (range: unknown) => {
        const { r0, c0, r1, c1 } = parseRange(range);
        budget((r1 - r0 + 1) * (c1 - c0 + 1), `clear(${String(range)})`);
        for (let r = r0; r <= r1; r++) {
          for (let c = c0; c <= c1; c++) written[`${r},${c}`] = { v: null, f: null };
        }
        ops.push({ kind: "clear", sheet: name, range: String(range) });
      },
      usedRange: (): string | null => {
        assertReadable("usedRange()");
        // Pre-script extent — the script's own writes don't move it.
        let maxR = -1;
        let maxC = -1;
        for (const key of Object.keys(data)) {
          const comma = key.indexOf(",");
          const r = +key.slice(0, comma);
          const c = +key.slice(comma + 1);
          if (r > maxR) maxR = r;
          if (c > maxC) maxC = c;
        }
        return maxR < 0 ? null : `A1:${colLetters(maxC)}${maxR + 1}`;
      },
    };
  };

  const sheetFn = (name?: unknown) => {
    let resolved: string;
    if (name === undefined || name === null) {
      if (model.sheetNames.length !== 1) {
        throw new Error(
          `sheet() requires a name when the workbook has multiple sheets. Available: ${model.sheetNames.join(", ")}`,
        );
      }
      resolved = model.sheetNames[0];
    } else {
      resolved = String(name);
      if (!model.sheets[resolved]) {
        // Auto-register: targeting a new sheet name creates that sheet when
        // the script's writes land (near-miss spellings of existing names
        // are rejected as typos at apply time). Reads see an empty sheet.
        model.sheets[resolved] = {};
        model.sheetNames.push(resolved);
      }
    }
    return (handles[resolved] ??= makeHandle(resolved));
  };
  const sheetsFn = () => model.sheetNames.slice();
  let logBytes = 0;
  let logBudgetNoted = false;
  const logFn = (...args: unknown[]) => {
    if (logs.length >= limits.maxLogs) return;
    const full = args
      .map((a) => {
        if (typeof a === "string") return a;
        try {
          return JSON.stringify(a);
        } catch {
          return String(a);
        }
      })
      .join(" ");
    // Cuts are visible: a row dump that lost its tail used to look complete.
    const line = full.length > limits.maxLogChars ? `${full.slice(0, limits.maxLogChars)}… [+${full.length - limits.maxLogChars} chars]` : full;
    if (logBytes + line.length > limits.maxLogBytes) {
      if (!logBudgetNoted) {
        logBudgetNoted = true;
        logs.push(`(log budget of ${Math.round(limits.maxLogBytes / 1024)} KB reached — later log() calls were dropped; use query_rows / read_range for bulk reads)`);
      }
      return;
    }
    logBytes += line.length;
    logs.push(line);
  };

  try {
    // eslint-disable-next-line no-new-func
    const fn = new Function("sheet", "sheets", "log", `"use strict";\n${code}`);
    fn(sheetFn, sheetsFn, logFn);
    return { ok: true, ops, logs };
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return { ok: false, error: msg, logs };
  }
}

/**
 * Build the script's read model from the ENGINE's per-sheet JSON
 * (`calc_engine_sheet_cells` / wasm `sheetCells`: `{sheet, cells:[{r,c,f,v}]}`,
 * 1-based) for the sheets that were fetched; every other sheet is listed but
 * marked unloaded. Engine-sourced: the model is the truth regardless of what
 * the grid holds (lazy sheets), and nothing is copied out of the grid.
 * SELF-CONTAINED like scriptCore: its source is stringified into the worker.
 */
export function buildScriptModelFromEngineSheets(sheetNames: string[], loaded: Record<string, string>): ScriptModel {
  const sheets: Record<string, ScriptSheetData> = {};
  const unloaded: string[] = [];
  for (const name of sheetNames) {
    const raw = loaded[name];
    if (raw === undefined) {
      sheets[name] = {};
      unloaded.push(name);
      continue;
    }
    const parsed = JSON.parse(raw) as { cells?: Array<{ r: number; c: number; f?: string; v?: unknown }> };
    const data: ScriptSheetData = {};
    for (const cell of parsed.cells ?? []) {
      const f = typeof cell.f === "string" && cell.f ? cell.f : null;
      const v = cell.v === undefined ? null : (cell.v as string | number | boolean | null);
      if (v === null && f === null) continue;
      data[`${cell.r - 1},${cell.c - 1}`] = { v, f };
    }
    sheets[name] = data;
  }
  return { sheetNames: sheetNames.slice(), sheets, unloaded };
}

/**
 * Which sheets a script reads: those named as string literals in its source.
 * `sheets()` (enumeration) or a non-literal `sheet(expr)` means "could be
 * any" and loads them all.
 */
export function sheetsReferencedByScript(code: string, sheetNames: string[]): { names: string[]; all: boolean } {
  const all = /\bsheets\s*\(/.test(code) || /\bsheet\s*\(\s*[^"'`)\s]/.test(code);
  if (all) return { names: sheetNames.slice(), all: true };
  return { names: sheetNames.filter((n) => code.includes(n)), all: false };
}

/**
 * Build the script's read model from a Univer workbook snapshot
 * (`grid.getWorkbookSnapshot()` — same source the structural index walks).
 * Throws (with an agent-readable message) if the snapshot is unusable.
 */
export function buildScriptModelFromSnapshot(snapshot: any): ScriptModel {
  if (!snapshot || !snapshot.sheets) throw new Error("no workbook is loaded");
  const order: string[] = snapshot.sheetOrder ?? Object.keys(snapshot.sheets);
  const sheetNames: string[] = [];
  const sheets: Record<string, ScriptSheetData> = {};
  let total = 0;
  for (const id of order) {
    const sh = snapshot.sheets[id];
    if (!sh?.name) continue;
    sheetNames.push(sh.name);
    const data: ScriptSheetData = {};
    const cellData = sh.cellData ?? {};
    for (const r of Object.keys(cellData)) {
      const rowCells = cellData[r] ?? {};
      for (const c of Object.keys(rowCells)) {
        const cell = rowCells[c];
        if (!cell) continue;
        const v = cell.v ?? null;
        const f = typeof cell.f === "string" && cell.f ? cell.f : null;
        if (v === null && f === null) continue;
        data[`${r},${c}`] = { v, f };
        total++;
      }
    }
    sheets[sh.name] = data;
  }
  if (total > MAX_MODEL_CELLS) {
    throw new Error(`workbook too large for run_script (${total} non-empty cells) — use read_range + set_range`);
  }
  return { sheetNames, sheets };
}

/**
 * Convert recorded script ops into ordinary UniverMutations. Repeated `set`
 * writes to one cell are coalesced (last write wins) so the review diff
 * shows each cell once — EXCEPT across an intervening clear() on the same
 * sheet, where in-place replacement would reorder the write before the
 * clear and flip the final state.
 */
export function scriptOpsToMutations(ops: ScriptOp[]): UniverMutation[] {
  const out: UniverMutation[] = [];
  const setIndexByCell = new Map<string, number>();
  const clearEpochBySheet = new Map<string, number>();
  for (const op of ops) {
    if (op.kind === "set") {
      const isFormula = typeof op.value === "string" && op.value.startsWith("=");
      const m: UniverMutation = {
        type: "set_cell",
        address: { sheet: op.sheet, row: op.row, col: op.col },
        old_value: null,
        new_value: op.value as string | number | null,
        new_formula: isFormula ? (op.value as string) : null,
      };
      const epoch = clearEpochBySheet.get(op.sheet) ?? 0;
      const key = `${op.sheet}!${op.row},${op.col}#${epoch}`;
      const prev = setIndexByCell.get(key);
      if (prev !== undefined) {
        out[prev] = m;
      } else {
        setIndexByCell.set(key, out.length);
        out.push(m);
      }
      continue;
    }
    if (op.kind === "format") {
      const cells = expandA1Range(op.range);
      if (cells.length === 0) continue; // validated in-core; belt and braces
      out.push({
        type: "set_format",
        sheet: op.sheet,
        range: op.range,
        cells,
        old_format: [],
        new_format: op.format as any,
      });
      continue;
    }
    // clear — old values are captured downstream at apply time.
    const cells = expandA1Range(op.range);
    if (cells.length === 0) continue;
    clearEpochBySheet.set(op.sheet, (clearEpochBySheet.get(op.sheet) ?? 0) + 1);
    out.push({
      type: "clear_range",
      sheet: op.sheet,
      range: op.range,
      cells: cells.map((c: { row: number; col: number }) => ({ row: c.row, col: c.col, old_value: null, old_formula: null })),
    });
  }
  return out;
}

/** Neuter I/O the script has no business using — fetch_web is the
 * sanctioned path for external data, and scripts must stay deterministic
 * local transforms. Assignment (not delete) works on worker globals. */
const WORKER_PREAMBLE = `
try { self.fetch = undefined; } catch (e) {}
try { self.XMLHttpRequest = undefined; } catch (e) {}
try { self.WebSocket = undefined; } catch (e) {}
try { self.EventSource = undefined; } catch (e) {}
try { self.importScripts = undefined; } catch (e) {}
try { self.indexedDB = undefined; } catch (e) {}
try { self.caches = undefined; } catch (e) {}
`;

export type ScriptExecResult =
  | { ok: true; mutations: UniverMutation[]; writes: number; logs: string[] }
  | { ok: false; error: string; logs: string[] };

/**
 * Full worker source: preamble + serialized scriptCore + message pump.
 * Exported so tests can execute the EXACT code the worker runs — the
 * serialization silently breaks if a build transform ever makes scriptCore
 * reference module-scope helpers, and only running the string catches that.
 */
export function buildWorkerSource(): string {
  return (
    WORKER_PREAMBLE +
    `const __core = ${scriptCore.toString()};\n` +
    `const __fromEngine = ${buildScriptModelFromEngineSheets.toString()};\n` +
    `self.onmessage = (e) => {\n` +
    `  let r;\n` +
    `  try {\n` +
    `    const d = e.data;\n` +
    `    if (d.sheetBuffers) {\n` +
    `      const dec = new TextDecoder();\n` +
    `      const loaded = {};\n` +
    `      for (const name of Object.keys(d.sheetBuffers)) loaded[name] = dec.decode(new Uint8Array(d.sheetBuffers[name]));\n` +
    `      d.model = __fromEngine(d.sheetNames, loaded);\n` +
    `      d.sheetBuffers = null;\n` +
    `    }\n` +
    `    r = __core(d);\n` +
    `  } catch (err) { r = { ok: false, error: String((err && err.message) || err), logs: [] }; }\n` +
    `  self.postMessage(r);\n` +
    `};\n`
  );
}

/**
 * Run an agent script in a throwaway sandboxed worker and convert its
 * recorded ops to mutations. Never rejects — every failure mode returns
 * `{ok: false}` with an agent-readable error so the model can fix and retry.
 */
/** Engine-backed source for a script's read model: the engine's per-sheet
 * JSON as bytes, fetched only for the sheets the script names. */
export type ScriptEngineSource = {
  sheetNames: string[];
  fetchSheet: (name: string) => Promise<ArrayBuffer | null>;
};

export async function executeSheetScript(
  grid: { getWorkbookSnapshot: () => any | null },
  code: string,
  opts?: { timeoutMs?: number; maxTouched?: number; source?: ScriptEngineSource },
): Promise<ScriptExecResult> {
  const timeoutMs = opts?.timeoutMs ?? SCRIPT_TIMEOUT_MS;
  const limits: ScriptLimits = {
    maxTouched: opts?.maxTouched ?? SCRIPT_MAX_TOUCHED,
    maxLogs: SCRIPT_MAX_LOGS,
    maxLogChars: SCRIPT_MAX_LOG_CHARS,
    maxLogBytes: SCRIPT_MAX_LOG_BYTES,
  };
  // Engine-sourced run: bytes per named sheet, transferred into the worker,
  // parsed there. The grid is never copied, so workbook size stops being a
  // reason to refuse (the whole-snapshot clone below is capped at 1M cells).
  let engineBuffers: Record<string, ArrayBuffer> | null = null;
  if (opts?.source) {
    const { names, all } = sheetsReferencedByScript(code, opts.source.sheetNames);
    engineBuffers = {};
    let bytes = 0;
    for (const name of names) {
      const buf = await opts.source.fetchSheet(name);
      if (!buf) continue;
      bytes += buf.byteLength;
      if (bytes > SCRIPT_MAX_SOURCE_BYTES) {
        return {
          ok: false,
          error:
            `this script would load ${Math.round(bytes / 1e6)} MB of sheet data (${all ? "it enumerates every sheet" : `${names.length} sheets`}); ` +
            `name only the sheets you actually read, as string literals, and split the work across scripts if needed`,
          logs: [],
        };
      }
      engineBuffers[name] = buf;
    }
  }

  let model: ScriptModel | null = null;
  if (!engineBuffers) {
    try {
      model = buildScriptModelFromSnapshot(grid.getWorkbookSnapshot());
    } catch (e: unknown) {
      return { ok: false, error: e instanceof Error ? e.message : String(e), logs: [] };
    }
  }

  let core: ScriptCoreResult;
  if (typeof Worker === "undefined") {
    // Test / headless environments only — the shipped webview always has
    // Worker. Direct execution has no kill switch, so it never runs in prod.
    if (engineBuffers) {
      const dec = new TextDecoder();
      const loaded: Record<string, string> = {};
      for (const name of Object.keys(engineBuffers)) loaded[name] = dec.decode(new Uint8Array(engineBuffers[name]));
      model = buildScriptModelFromEngineSheets(opts!.source!.sheetNames, loaded);
    }
    core = scriptCore({ code, model: model!, limits });
  } else {
    const payload: any = engineBuffers
      ? { code, limits, sheetNames: opts!.source!.sheetNames, sheetBuffers: engineBuffers }
      : { code, model, limits };
    const transfer = engineBuffers ? Object.values(engineBuffers) : [];
    const url = URL.createObjectURL(new Blob([buildWorkerSource()], { type: "application/javascript" }));
    const worker = new Worker(url);
    core = await new Promise<ScriptCoreResult>((resolve) => {
      const timer = setTimeout(() => {
        worker.terminate();
        resolve({
          ok: false,
          error: `script timed out after ${timeoutMs}ms (infinite loop?) — simplify it or split the work`,
          logs: [],
        });
      }, timeoutMs);
      worker.onmessage = (e: MessageEvent) => {
        clearTimeout(timer);
        resolve(e.data as ScriptCoreResult);
      };
      worker.onerror = (e: ErrorEvent) => {
        clearTimeout(timer);
        resolve({ ok: false, error: `script worker error: ${e.message || "unknown"}`, logs: [] });
      };
      worker.postMessage(payload, transfer);
    }).finally(() => {
      worker.terminate();
      URL.revokeObjectURL(url);
    });
  }

  if (!core.ok) return { ok: false, error: core.error, logs: core.logs };
  const mutations = scriptOpsToMutations(core.ops);
  const writes = core.ops.length;
  return { ok: true, mutations, writes, logs: core.logs };
}
