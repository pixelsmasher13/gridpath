import type { UniverGridHandle, CellFormatShape } from "../components/UniverGrid";
import type { WorkbookContext } from "./agentClient";
import { colIndexToLetters, previewNumber, shiftFormulaCols } from "../formulaRuns";
import { firstLabelIn, isDateFormat, serialToIsoDate, snapshotContentKey } from "./workbookIndex";

/**
 * Capture a compact snapshot of the active workbook to ship to the agent.
 *
 * For every sheet we emit:
 *   1. A short **Style conventions** palette — unique format patterns sampled
 *      from the sheet, so the agent can mirror template styling (rule 14)
 *      without guessing.
 *   2. Cell values/formulas in "A1 = value" form, with a compact `[fmt …]`
 *      suffix on cells that carry non-default formatting.
 *
 * FILL-RUN COMPRESSION: a quarterly model row is one formula filled across
 * 50–100 columns; spelling out every translated copy made a 13-sheet model
 * cost ~430K tokens per session. Consecutive cells whose formulas are exact
 * column-translations of each other collapse to one line. The SECOND
 * formula is spelled out too — A/B testing showed a model reading
 * "(filled right)" as "the same literal formula repeated", concluding every
 * cell referenced the first column. Interior value samples guard the
 * "what's the value mid-run?" case:
 *
 *   D28:CX28 = =C28*(1+D$4), then =D28*(1+E$4), …  (filled right ×99)
 *     → D28=1.23, AB28=7.9, BC28=21.4, CX28=45.6
 *
 * Numbers are shown at ~6 significant digits — this is a preview, not a
 * source of truth. The output is deterministic for identical sheet content;
 * the prompt-cache split depends on that byte stability.
 *
 * Caps keep this under the context budget: 1500 lines/sheet (a run counts
 * as one line, so coverage IMPROVES with compression), 80 format
 * annotations/sheet, 24 unique style samples in the palette.
 *
 * Edits do NOT re-capture immediately: the preview is pinned at a base
 * snapshot and changed cells ship separately in `delta` — see the
 * BASE+DELTA CAPTURE comment on the memo below.
 */
const MAX_CELLS_PER_SHEET = 1500;
/** Max cells that get an inline `[fmt …]` annotation per sheet. */
const MAX_FORMAT_ANNOTATIONS = 80;
/** Max unique style samples listed in the Style conventions block. */
const MAX_STYLE_SAMPLES = 24;
/** Minimum consecutive column-translated formulas to collapse into a run. */
const MIN_RUN = 4;
/**
 * Hard byte budget per sheet preview. With run compression, ~32KB of lines
 * covers tens of thousands of cells; without a byte cap, text-heavy sheets
 * (disclosure boilerplate) and fully-enumerated sheets dominate the prompt.
 * When a sheet IS truncated, a row-label index of the cut rows follows so
 * the agent knows what exists below the fold and can ask for ranges —
 * A/B testing showed a capped sheet silently steering the answer to
 * whatever rows happened to fit.
 */
const MAX_SHEET_PREVIEW_BYTES = 32 * 1024;
/** Max entries in the truncated-rows label index. */
const MAX_TRUNCATION_INDEX = 48;
/** Long text cells (paragraph boilerplate) are clipped in the preview. */
const MAX_TEXT_CELL_CHARS = 160;

/**
 * BASE+DELTA CAPTURE. The context block is the dominant token cost per
 * request and sits under a cache_control breakpoint, so any byte change
 * re-bills the whole block at the cache-WRITE rate (measured: a 12-turn
 * editing session on a ~100K-token model preview costs ~2.4x more than the
 * same session with a stable block). Instead of re-capturing on every
 * content change, we PIN the preview at a base snapshot and keep returning
 * those exact bytes (cache hit); the cells that changed since the base ride
 * along in `delta`, which Rust renders into the UNCACHED turn tail. Base +
 * delta always equals current state — nothing is hidden, it's presentation.
 *
 * Re-snapshot (fresh capture, new base, cache re-write) happens when:
 *   - the sheet set/names/order changed (structural edits: renames, adds,
 *     row/col inserts show up as huge deltas and trip the byte cap), or
 *   - the rendered delta exceeds its byte budget — at that point one
 *     re-write is cheaper than re-shipping a fat delta every turn.
 *
 * The read tools (describe_workbook / find_rows / read_range) are served
 * from the CURRENT snapshot/grid, never the base — only the preview is
 * pinned, and the delta header says so.
 *
 * The memo keys on a hash of sheet CONTENT (formulas + stored values), not
 * evaluated values — those jitter across captures (recalc settling, facade
 * timing) and live in the uncached delta only.
 */
/**
 * The delta budget is what makes or breaks this on big workbooks. When the
 * delta fits, it rides UNCACHED in the turn tail (paid at ~1x every turn);
 * over budget, we re-snapshot (one cache WRITE at 1.25x of the whole base,
 * then reads at 0.1x). Which is cheaper depends on delta size vs base size
 * and how many turns remain:
 *
 *   carry:      delta * T
 *   re-write:   base * 1.25 + base * 0.1 * T
 *
 * They cross around delta ≈ 0.1*base for long sessions and much higher for
 * short ones — i.e. re-writing a 180K-token base to avoid carrying a
 * 25K-token delta is a bad trade until ~16 turns out. So the budget must
 * SCALE WITH THE BASE, not be a flat cap: a fixed 8KB re-snapshots ~20x too
 * eagerly on a 290KB preview (measured on a real Amazon model — one edit
 * turn re-billed 181K cache-write tokens on the FOLLOWING "thank you"). We
 * carry a delta up to ~a third of the base, capped so a single turn's
 * uncached tail can't balloon latency, floored so tiny books still
 * re-snapshot promptly (a re-write there is cheap).
 */
const DELTA_BUDGET_FRACTION = 0.34;
/** Upper cap on the uncached delta — bounds per-turn prefill latency. */
const DELTA_MAX_BYTES = 64 * 1024;
/** Lower floor — below this, re-snapshot rather than carry a delta. */
const DELTA_MIN_BUDGET_BYTES = 2 * 1024;

type BaseCell = { f: string | null; v: unknown };
type CaptureMemoEntry = {
  key: string;
  sheets: WorkbookContext["sheets"];
  /** Sheet names in order at base time — any drift forces a re-snapshot. */
  sheetNames: string[];
  /** Deep-copied {f, v} per non-empty cell at base time, keyed "r,c". */
  baseCells: Map<string, Map<string, BaseCell>>;
  /** Total preview bytes of the base — sizes the delta budget. */
  previewBytes: number;
};
const captureMemo = new Map<string, CaptureMemoEntry>();

export function captureWorkbookContext(
  path: string,
  grid: UniverGridHandle | null,
): WorkbookContext {
  const snapshot = grid?.getWorkbookSnapshot?.();
  if (!snapshot || !snapshot.sheets) {
    return { path, sheets: [] };
  }

  const contentKey = snapshotContentKey(snapshot);
  const memo = captureMemo.get(path);
  if (memo && memo.key === contentKey) {
    return { path, sheets: memo.sheets };
  }
  // Content changed since the base: try to keep the (cached) base preview
  // and ship only the changed cells. Falls through to a fresh capture when
  // the delta is structural or over budget.
  if (memo) {
    const delta = buildSnapshotDelta(memo, snapshot, grid);
    if (delta !== null) {
      return { path, sheets: memo.sheets, delta };
    }
  }
  const order: string[] = snapshot.sheetOrder ?? Object.keys(snapshot.sheets);
  const sheets = order
    .map((sheetId) => {
      const sheet = snapshot.sheets[sheetId];
      if (!sheet) return null;
      const sheetName = sheet.name as string;
      const cellData = sheet.cellData ?? {};
      const rowKeys = Object.keys(cellData)
        .map(Number)
        .sort((a, b) => a - b);

      // Hidden sheets (Excel state "hidden"/"veryHidden") are usually the
      // scratch/cache/calc plumbing behind the visible tabs. A full cell
      // preview earns its context cost poorly there, so ship only the used
      // range plus a row-label index — enough to know the sheet exists and
      // trace references into it; describe_workbook / find_rows / read_range
      // serve hidden sheets in full on demand. Edits/recalcs on hidden cells
      // still ride in `delta` like any other change.
      if (sheet.hidden === 1) {
        return {
          name: sheetName,
          row_count: Number(sheet.rowCount ?? 100),
          column_count: Number(sheet.columnCount ?? 26),
          used_range: usedRangeOf(cellData, rowKeys),
          cells_preview: hiddenSheetPreview(cellData, rowKeys),
        };
      }

      const lines: string[] = [];
      // signature → first example cell address (for the palette)
      const styleSamples = new Map<string, string>();
      let formatAnnotations = 0;

      // Evaluated value of a cell for display, preferring the live grid
      // (sees recalc results) over the snapshot's cached value. Values are
      // RAW (0.7235, never the "72.35%" display string — that misled agents
      // into "the data is text" rabbit holes); the formatted rendering is
      // appended as an annotation instead. Date-formatted numbers are Excel
      // serials — annotate with the ISO date so the agent doesn't reason
      // about "38183".
      const evaluatedOf = (r: number, c: number, cell: any, numberFormat?: string | null): string | null => {
        const liveCell = grid?.getCell?.(sheetName, r, c);
        const live = liveCell?.value;
        const v = live !== undefined && live !== null ? live : cell.v;
        if (v === undefined || v === null) return null;
        if (typeof v === "number") {
          const iso = numberFormat && isDateFormat(numberFormat) ? serialToIsoDate(v) : null;
          const shown = iso ?? liveCell?.display ?? null;
          return shown ? `${previewNumber(v)} (${shown})` : previewNumber(v);
        }
        return JSON.stringify(clipText(v));
      };

      // Normalized formula text ('=' stripped) for run comparison.
      const normF = (f: unknown): string | null =>
        typeof f === "string" && f.trim() ? f.trim().replace(/^=/, "") : null;

      let previewBytes = 0;
      let truncatedAtRow: number | null = null;
      outer: for (const r of rowKeys) {
        const row = cellData[r];
        if (!row) continue;
        const cols = Object.keys(row)
          .map(Number)
          .sort((a, b) => a - b);
        let i = 0;
        while (i < cols.length) {
          if (lines.length >= MAX_CELLS_PER_SHEET || previewBytes >= MAX_SHEET_PREVIEW_BYTES) {
            truncatedAtRow = r;
            break outer;
          }
          const c = cols[i];
          const cell = row[c];
          const addr = colIndexToLetters(c) + (r + 1);

          // --- fill-run detection: consecutive columns whose formulas are
          // exact column-translations of their left neighbour ---
          const f0 = normF(cell?.f);
          if (f0) {
            let j = i;
            let prevF = f0;
            while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) {
              const nextF = normF(row[cols[j + 1]]?.f);
              if (!nextF) break;
              const expected = shiftFormulaCols(prevF, 1);
              if (expected === null || expected !== nextF) break;
              prevF = nextF;
              j++;
            }
            const runLen = j - i + 1;
            if (runLen >= MIN_RUN) {
              const endCol = cols[j];
              const endAddr = colIndexToLetters(endCol) + (r + 1);
              const fmt = grid?.getCellFormat?.(sheetName, r, c);
              // Second formula spelled out so "translated per column" can't
              // be misread as "repeated verbatim".
              const second = normF(row[cols[i + 1]]?.f);
              let runLine = `${addr}:${endAddr} = ${String(cell.f).trim()}, then =${second}, …  (filled right ×${runLen})`;
              // Sample values: endpoints always, plus up to 3 interior
              // points on longer runs (evenly spaced) so mid-run questions
              // don't force a guess.
              const sampleIdx = new Set<number>([i, j]);
              if (runLen >= 8) {
                for (const frac of [0.25, 0.5, 0.75]) {
                  sampleIdx.add(i + Math.round((j - i) * frac));
                }
              }
              const samples: string[] = [];
              for (const k of [...sampleIdx].sort((a, b) => a - b)) {
                const kc = cols[k];
                const val = evaluatedOf(r, kc, row[kc], fmt?.number_format);
                if (val !== null) samples.push(`${colIndexToLetters(kc)}${r + 1}=${val}`);
              }
              if (samples.length) {
                runLine += `  → ${samples.join(", ")}`;
              }
              const fmtSummary = summarizeFormat(fmt);
              if (fmtSummary) {
                if (styleSamples.size < MAX_STYLE_SAMPLES && !styleSamples.has(fmtSummary)) {
                  styleSamples.set(fmtSummary, addr);
                }
                if (formatAnnotations < MAX_FORMAT_ANNOTATIONS) {
                  runLine = `${runLine}  [${fmtSummary}]`;
                  formatAnnotations++;
                }
              }
              lines.push(runLine);
              previewBytes += runLine.length + 1;
              i = j + 1;
              continue;
            }
          }

          const fmt = grid?.getCellFormat?.(sheetName, r, c);
          let valueLine: string | null = null;
          // For formula cells, ship BOTH the formula AND its evaluated
          // value (when available). Without this the agent can never see
          // #VALUE! errors or what its own formulas computed to.
          if (cell.f) {
            const evaluated = evaluatedOf(r, c, cell, fmt?.number_format);
            valueLine = evaluated !== null ? `${addr} = ${cell.f}  → ${evaluated}` : `${addr} = ${cell.f}`;
          } else if (cell.v !== undefined) {
            const rendered = evaluatedOf(r, c, cell, fmt?.number_format);
            if (rendered !== null) valueLine = `${addr} = ${rendered}`;
          }
          if (!valueLine) {
            i++;
            continue;
          }

          // Format sample — only when the cell has managed styling. Cap
          // annotations so a fully-styled sheet doesn't double the preview.
          const fmtSummary = summarizeFormat(fmt);
          if (fmtSummary) {
            if (styleSamples.size < MAX_STYLE_SAMPLES && !styleSamples.has(fmtSummary)) {
              styleSamples.set(fmtSummary, addr);
            }
            if (formatAnnotations < MAX_FORMAT_ANNOTATIONS) {
              valueLine = `${valueLine}  [${fmtSummary}]`;
              formatAnnotations++;
            }
          }

          lines.push(valueLine);
          previewBytes += valueLine.length + 1;
          i++;
        }
      }

      // Truncated: append an index of the rows that were cut (row number +
      // first text label) so the agent can still SEE what exists below the
      // fold and pull specific ranges, instead of concluding the sheet ends
      // where the budget did.
      if (truncatedAtRow !== null) {
        lines.push(
          `... (preview truncated at row ${truncatedAtRow + 1} — remaining rows listed below; use ranges to read them)`,
        );
        let indexed = 0;
        for (const r of rowKeys) {
          if (r < truncatedAtRow) continue;
          if (indexed >= MAX_TRUNCATION_INDEX) {
            lines.push("    … (more rows)");
            break;
          }
          const row = cellData[r];
          if (!row) continue;
          // Row label = first cell value that looks like a label (flag
          // columns, years, and numbers are rejected by the shared heuristic).
          const label = firstLabelIn(
            Object.keys(row)
              .map(Number)
              .sort((a, b) => a - b)
              .map((c) => row[c]?.v),
          );
          if (label !== null) {
            lines.push(`    row ${r + 1}: ${JSON.stringify(clipText(label))}`);
            indexed++;
          }
        }
      }

      const previewParts: string[] = [];
      if (styleSamples.size > 0) {
        previewParts.push("Style conventions (mirror these when editing this sheet):");
        for (const [summary, example] of styleSamples) {
          previewParts.push(`  ${example}: ${summary}`);
        }
        previewParts.push("");
      }
      previewParts.push(lines.length > 0 ? lines.join("\n") : "(empty)");

      const usedRange = usedRangeOf(cellData, rowKeys);

      return {
        name: sheetName,
        row_count: Number(sheet.rowCount ?? 100),
        column_count: Number(sheet.columnCount ?? 26),
        used_range: usedRange,
        cells_preview: previewParts.join("\n"),
      };
    })
    .filter((s): s is NonNullable<typeof s> => s !== null);
  const { sheetNames, baseCells } = snapshotBaseCells(snapshot);
  captureMemo.set(path, {
    key: contentKey,
    sheets,
    sheetNames,
    baseCells,
    previewBytes: sheets.reduce((n, s) => n + s.cells_preview.length, 0),
  });
  return { path, sheets };
}

/**
 * Used data extent in A1 ("A1:Q48") — what the model should anchor to.
 * sheet.rowCount/columnCount are grid CAPACITY (10000×200), which reads as
 * "this sheet has 10000 rows of data" — misinformation.
 */
function usedRangeOf(cellData: any, rowKeys: number[]): string | null {
  let maxUsedRow = -1;
  let maxUsedCol = -1;
  for (const rk of rowKeys) {
    const row = cellData[rk];
    if (!row) continue;
    for (const ck of Object.keys(row)) {
      if (row[ck]?.v !== undefined || row[ck]?.f !== undefined) {
        if (rk > maxUsedRow) maxUsedRow = rk;
        const cn = Number(ck);
        if (cn > maxUsedCol) maxUsedCol = cn;
      }
    }
  }
  return maxUsedRow >= 0 ? `A1:${colIndexToLetters(maxUsedCol)}${maxUsedRow + 1}` : null;
}

/**
 * Stand-in preview for hidden sheets: a marker line plus a row-label index
 * (same shape as the truncation index) so the agent can see WHAT the sheet
 * holds without paying for its cells.
 */
function hiddenSheetPreview(cellData: any, rowKeys: number[]): string {
  const lines: string[] = [
    "(This sheet is HIDDEN in the workbook — typically scratch/cache/calc " +
      "data feeding the visible sheets. Cell preview omitted; the row labels " +
      "below show what it contains. describe_workbook / find_rows / " +
      "read_range serve its cells in full — read before referencing or " +
      "editing it.)",
  ];
  let indexed = 0;
  for (const r of rowKeys) {
    if (indexed >= MAX_TRUNCATION_INDEX) {
      lines.push("    … (more rows)");
      break;
    }
    const row = cellData[r];
    if (!row) continue;
    const label = firstLabelIn(
      Object.keys(row)
        .map(Number)
        .sort((a, b) => a - b)
        .map((c) => row[c]?.v),
    );
    if (label !== null) {
      lines.push(`    row ${r + 1}: ${JSON.stringify(clipText(label))}`);
      indexed++;
    }
  }
  return lines.join("\n");
}

/**
 * Deep-copy the content cells of a snapshot into plain maps for later
 * diffing. Copied (not referenced) because the grid may hand back live
 * structures — a base that drifts with edits would diff to an empty delta.
 */
function snapshotBaseCells(snapshot: any): {
  sheetNames: string[];
  baseCells: CaptureMemoEntry["baseCells"];
} {
  const order: string[] = snapshot.sheetOrder ?? Object.keys(snapshot.sheets);
  const sheetNames: string[] = [];
  const baseCells: CaptureMemoEntry["baseCells"] = new Map();
  for (const id of order) {
    const sheet = snapshot.sheets[id];
    if (!sheet) continue;
    const name = String(sheet.name);
    sheetNames.push(name);
    const cells = new Map<string, BaseCell>();
    const cellData = sheet.cellData ?? {};
    for (const r of Object.keys(cellData)) {
      const row = cellData[r];
      if (!row) continue;
      for (const c of Object.keys(row)) {
        const cell = row[c];
        if (!cell) continue;
        const f = typeof cell.f === "string" && cell.f.trim() ? cell.f.trim() : null;
        const v = cell.v === undefined ? null : cell.v;
        if (f === null && v === null) continue;
        cells.set(`${r},${c}`, { f, v });
      }
    }
    baseCells.set(name, cells);
  }
  return { sheetNames, baseCells };
}

/** Values compare: primitives by identity, rare object values structurally. */
function sameValue(a: unknown, b: unknown): boolean {
  if (Object.is(a, b)) return true;
  if (typeof a === "object" && a !== null && typeof b === "object" && b !== null) {
    try {
      return JSON.stringify(a) === JSON.stringify(b);
    } catch {
      return false;
    }
  }
  return false;
}

/**
 * Render the cells that changed between the memoized base and the current
 * snapshot, in the same `A1 = value` / fill-run notation as the preview.
 * Returns null when a fresh capture is the better deal: sheet set changed,
 * the delta is over budget, or (defensively) nothing actually differs.
 */
function buildSnapshotDelta(
  memo: CaptureMemoEntry,
  snapshot: any,
  grid: UniverGridHandle | null,
): string | null {
  const order: string[] = snapshot.sheetOrder ?? Object.keys(snapshot.sheets);
  const current = order
    .map((id) => snapshot.sheets[id])
    .filter((s: any) => s != null);
  const currentNames = current.map((s: any) => String(s.name));
  if (
    currentNames.length !== memo.sheetNames.length ||
    currentNames.some((n, i) => n !== memo.sheetNames[i])
  ) {
    return null; // structural change — re-snapshot
  }

  const budget = Math.min(
    DELTA_MAX_BYTES,
    Math.max(DELTA_MIN_BUDGET_BYTES, Math.floor(memo.previewBytes * DELTA_BUDGET_FRACTION)),
  );
  const header =
    "# Workbook changes since the preview snapshot\n" +
    "The workbook preview reflects an earlier snapshot. The cells below have " +
    "changed since then — the CURRENT values/formulas shown here SUPERSEDE the " +
    "preview for these cells. `A1 = …` lines are edits (new content); " +
    "`A1 recalc → v` lines are formula cells whose formula is UNCHANGED from " +
    "the preview but whose computed value moved (upstream edit rippled " +
    "through). Cells not listed are unchanged from the preview. " +
    "describe_workbook, find_rows and read_range always serve current data.\n";
  let bytes = header.length;
  const parts: string[] = [header];

  const normF = (f: unknown): string | null =>
    typeof f === "string" && f.trim() ? f.trim().replace(/^=/, "") : null;

  for (const sheet of current) {
    const sheetName = String(sheet.name);
    const base = memo.baseCells.get(sheetName) ?? new Map<string, BaseCell>();
    const cellData = sheet.cellData ?? {};

    // Changed/added cells, grouped by row with sorted columns so run
    // compression can collapse a filled or recalculated row to one line.
    // `recalc` marks cells whose FORMULA is unchanged and only the cached
    // value moved — an upstream edit rippling through. One assumption edit
    // recalcs hundreds of dependents on a dense model; spelling each out
    // would trip the budget and force the re-snapshot we're avoiding, so
    // recalc cells render as compact per-row segments instead.
    const changedByRow = new Map<number, Array<{ c: number; cell: any; recalc: boolean }>>();
    const seen = new Set<string>();
    for (const rk of Object.keys(cellData)) {
      const row = cellData[rk];
      if (!row) continue;
      for (const ck of Object.keys(row)) {
        const cell = row[ck];
        if (!cell) continue;
        const f = typeof cell.f === "string" && cell.f.trim() ? cell.f.trim() : null;
        const v = cell.v === undefined ? null : cell.v;
        if (f === null && v === null) continue;
        seen.add(`${rk},${ck}`);
        const b = base.get(`${rk},${ck}`);
        if (b && b.f === f && sameValue(b.v, v)) continue;
        const r = Number(rk);
        if (!changedByRow.has(r)) changedByRow.set(r, []);
        changedByRow
          .get(r)!
          .push({ c: Number(ck), cell, recalc: b !== undefined && f !== null && b.f === f });
      }
    }
    const cleared: string[] = [];
    for (const key of base.keys()) {
      if (!seen.has(key)) {
        const [r, c] = key.split(",").map(Number);
        cleared.push(colIndexToLetters(c) + (r + 1));
      }
    }
    if (changedByRow.size === 0 && cleared.length === 0) continue;

    const evaluated = (r: number, c: number, cell: any): string | null => {
      const live = grid?.getCell?.(sheetName, r, c)?.value;
      const v = live !== undefined && live !== null ? live : cell.v;
      if (v === undefined || v === null) return null;
      if (typeof v === "number") {
        const nf = grid?.getCellFormat?.(sheetName, r, c)?.number_format;
        const iso = nf && isDateFormat(nf) ? serialToIsoDate(v) : null;
        return iso ? `${previewNumber(v)} (${iso})` : previewNumber(v);
      }
      return JSON.stringify(clipText(v));
    };

    const lines: string[] = [`## Sheet: ${sheetName}`];
    for (const r of [...changedByRow.keys()].sort((a, b) => a - b)) {
      const cols = changedByRow.get(r)!.sort((a, b) => a.c - b.c);
      let i = 0;
      while (i < cols.length) {
        const { c, cell } = cols[i];
        const addr = colIndexToLetters(c) + (r + 1);

        // Recalc segment: consecutive value-only changes collapse to one
        // line with endpoint samples — the formulas are already in the
        // preview, only the computed values moved.
        if (cols[i].recalc) {
          let j = i;
          while (j + 1 < cols.length && cols[j + 1].c === cols[j].c + 1 && cols[j + 1].recalc) {
            j++;
          }
          const first = evaluated(r, c, cell);
          if (j > i) {
            const endAddr = colIndexToLetters(cols[j].c) + (r + 1);
            const last = evaluated(r, cols[j].c, cols[j].cell);
            const samples = [
              first !== null ? `${addr}=${first}` : null,
              last !== null ? `${endAddr}=${last}` : null,
            ].filter((s): s is string => s !== null);
            // Cells without a renderable value (e.g. error cells) still get
            // the segment line, just without the sample tail.
            lines.push(
              samples.length
                ? `${addr}:${endAddr} recalc → ${samples.join(", ")}`
                : `${addr}:${endAddr} recalc`,
            );
          } else {
            lines.push(first !== null ? `${addr} recalc → ${first}` : `${addr} recalc`);
          }
          i = j + 1;
          continue;
        }

        // Fill-run compression over the EDITED cells (an agent fill-right
        // or a dragged formula shouldn't cost one line per column).
        const f0 = normF(cell?.f);
        if (f0) {
          let j = i;
          let prevF = f0;
          while (
            j + 1 < cols.length &&
            cols[j + 1].c === cols[j].c + 1 &&
            !cols[j + 1].recalc
          ) {
            const nextF = normF(cols[j + 1].cell?.f);
            if (!nextF) break;
            const expected = shiftFormulaCols(prevF, 1);
            if (expected === null || expected !== nextF) break;
            prevF = nextF;
            j++;
          }
          if (j - i + 1 >= MIN_RUN) {
            const endAddr = colIndexToLetters(cols[j].c) + (r + 1);
            const second = normF(cols[i + 1].cell?.f);
            const first = evaluated(r, c, cell);
            const last = evaluated(r, cols[j].c, cols[j].cell);
            let line = `${addr}:${endAddr} = ${String(cell.f).trim()}, then =${second}, …  (filled right ×${j - i + 1})`;
            const samples = [
              first !== null ? `${addr}=${first}` : null,
              last !== null ? `${colIndexToLetters(cols[j].c)}${r + 1}=${last}` : null,
            ].filter((s): s is string => s !== null);
            if (samples.length) line += `  → ${samples.join(", ")}`;
            lines.push(line);
            i = j + 1;
            continue;
          }
        }
        const val = evaluated(r, c, cell);
        if (cell.f) {
          lines.push(val !== null ? `${addr} = ${cell.f}  → ${val}` : `${addr} = ${cell.f}`);
        } else if (val !== null) {
          lines.push(`${addr} = ${val}`);
        } else {
          lines.push(`${addr} = (cleared)`);
        }
        i++;
      }
    }
    for (const addr of cleared.sort()) {
      lines.push(`${addr} = (cleared)`);
    }
    const block = lines.join("\n") + "\n";
    bytes += block.length;
    if (bytes > budget) return null; // over budget — re-snapshot instead
    parts.push(block);
  }

  // Content key differed but no cell-level diff (e.g. only a sheet-level
  // attribute we don't render) — let the fresh capture path handle it.
  if (parts.length === 1) return null;
  return parts.join("\n");
}

/**
 * Compact, model-readable format summary. Returns null when the cell has
 * no managed styling worth shipping (keeps the preview lean on plain sheets).
 */
function summarizeFormat(f: CellFormatShape | null | undefined): string | null {
  if (!f) return null;
  const parts: string[] = [];
  if (f.bold) parts.push("bold");
  if (f.italic) parts.push("italic");
  if (f.underline) parts.push("underline");
  if (f.strike) parts.push("strike");
  if (f.font_family) parts.push(f.font_family);
  if (f.font_size) parts.push(`${f.font_size}pt`);
  if (f.font_color) parts.push(`color ${f.font_color}`);
  if (f.background_color) parts.push(`fill ${f.background_color}`);
  if (f.horizontal_align) parts.push(`align ${f.horizontal_align}`);
  if (f.vertical_align && f.vertical_align !== "bottom") parts.push(`valign ${f.vertical_align}`);
  if (f.wrap_text) parts.push("wrap");
  if (f.number_format && f.number_format !== "General") {
    parts.push(`fmt ${f.number_format}`);
  }
  return parts.length > 0 ? parts.join(", ") : null;
}

/** Clip long text cells (disclosure paragraphs) for preview display. */
function clipText(v: unknown): unknown {
  if (typeof v === "string" && v.length > MAX_TEXT_CELL_CHARS) {
    return `${v.slice(0, MAX_TEXT_CELL_CHARS)}… (+${v.length - MAX_TEXT_CELL_CHARS} chars)`;
  }
  return v;
}
