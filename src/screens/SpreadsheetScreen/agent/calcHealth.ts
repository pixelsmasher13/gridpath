import { ERROR_VALUES, type BaselineCell } from "../surgicalPatch";

/**
 * Live-evaluation vs file-saved divergence detection.
 *
 * GridPath re-evaluates every formula with its own engine, which cannot
 * fully compute some constructs Excel can (INDIRECT chains, external-workbook
 * references, add-in functions). On such cells the live value reads 0, an
 * error, or empty even though the user's FILE is sound — and the agent,
 * seeing authoritative-looking wrong numbers, historically concluded the
 * USER's model had "data integrity issues" and degraded its output to
 * hardcodes. These helpers compare live values against the load-time
 * baseline (`buildCellBaseline` — Excel's last-saved value per cell) so
 * read_range can attach `file_saved` to affected cells and the workbook
 * context can carry a calc-health warning (see system-prompt rule 17b).
 *
 * Only UNMODIFIED formula cells are ever flagged: if the live formula
 * differs from the file's, the cell was edited this session and the live
 * value is the only truth for it.
 *
 * Dual *plausible* numbers (live 15% vs Excel-cached 17%) are NOT flagged.
 * The user is looking at the live grid; surfacing a second figure makes the
 * agent debate which is "real" and invent staleness stories. file_saved is
 * reserved for live results that are unusable for reporting.
 */

/** Relative tolerance for numeric comparison — engine rounding is not divergence. */
const REL_TOL = 1e-6;
const ABS_TOL = 1e-9;

/** Formula equality up to cosmetic differences ('=' prefix, spacing, case). */
export function sameFormula(liveFormula: string | null | undefined, savedFormula: string | undefined): boolean {
  const norm = (f: string) => f.trim().replace(/^=/, "").replace(/\s+/g, "").toUpperCase();
  if (typeof liveFormula !== "string" || typeof savedFormula !== "string") return false;
  const a = norm(liveFormula);
  const b = norm(savedFormula);
  return a.length > 0 && a === b;
}

function isExcelError(v: unknown): boolean {
  return typeof v === "string" && ERROR_VALUES.has(v.trim());
}

function isEmpty(v: unknown): boolean {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

/**
 * True when the live result is unusable for reporting — empty, an Excel
 * error, or collapsed to 0 while Excel cached a real number. Two finite
 * non-zero numbers that merely disagree (15% vs 17%) are not broken: the
 * grid the user is looking at is the live value, and a sidecar figure
 * would only confuse the agent.
 */
export function liveValueBroken(liveValue: unknown, savedValue: unknown): boolean {
  if (isEmpty(liveValue)) return true;
  if (isExcelError(liveValue)) return true;
  // A computed boolean is a real result even when it disagrees with Excel.
  if (typeof liveValue === "boolean") return false;
  if (
    typeof liveValue === "number" &&
    Number.isFinite(liveValue) &&
    typeof savedValue === "number" &&
    Number.isFinite(savedValue)
  ) {
    return Math.abs(liveValue) <= ABS_TOL && Math.abs(savedValue) > ABS_TOL;
  }
  return false;
}

/**
 * True when the live value materially differs from the file's saved value.
 * Kept for rounding-tolerance tests; reporting uses `liveValueBroken`.
 */
export function valuesDiverge(liveValue: unknown, savedValue: unknown): boolean {
  const live = typeof liveValue === "boolean" ? Number(liveValue) : liveValue;
  const saved = typeof savedValue === "boolean" ? Number(savedValue) : savedValue;

  if (typeof live === "string" && typeof saved === "string") {
    return live.trim() !== saved.trim();
  }
  if (typeof live === "number" && typeof saved === "number") {
    const diff = Math.abs(live - saved);
    return diff > ABS_TOL && diff > REL_TOL * Math.max(Math.abs(live), Math.abs(saved));
  }
  if (live === null || live === undefined) return true; // engine produced nothing
  // Type mismatch (live error string vs saved number, number vs string…).
  return true;
}

/**
 * The file's saved value when — and only when — it should be surfaced next
 * to a live read: unmodified formula cell, live result unusable, and the
 * file cached a usable replacement. Returns undefined otherwise.
 */
export function fileSavedIfDivergent(
  liveValue: unknown,
  liveFormula: string | null,
  saved: BaselineCell | null | undefined,
): unknown {
  if (!saved || saved.f === undefined || saved.v === undefined) return undefined;
  if (!sameFormula(liveFormula, saved.f)) return undefined;
  if (isEmpty(saved.v) || isExcelError(saved.v)) return undefined;
  return liveValueBroken(liveValue, saved.v) ? saved.v : undefined;
}

export type CalcHealth = {
  /** Formula cells with a file-saved value whose formulas are unmodified. */
  checked: number;
  divergent: number;
  /** Sheets with divergent cells, worst first. */
  bySheet: Array<{ sheet: string; divergent: number }>;
  /** True when the scan stopped at the cell cap. */
  truncated: boolean;
};

/** Comparison cap — bounds the one-time scan on pathological workbooks. */
const MAX_HEALTH_CELLS = 200_000;

/**
 * One-shot workbook scan: compare every unmodified formula cell's live value
 * (from the Univer snapshot, where computed results live in `cellData[r][c].v`)
 * against the load-time baseline. Pure function of its inputs for testability.
 */
export function computeCalcHealth(
  snapshot: any,
  baseline: Map<string, Map<string, BaselineCell>>,
  cap: number = MAX_HEALTH_CELLS,
): CalcHealth {
  const liveByName = new Map<string, any>();
  for (const id of Object.keys(snapshot?.sheets ?? {})) {
    const s = snapshot.sheets[id];
    if (s?.name) liveByName.set(s.name, s.cellData ?? {});
  }

  let checked = 0;
  let divergent = 0;
  let truncated = false;
  const perSheet = new Map<string, number>();

  outer: for (const [sheetName, cells] of baseline) {
    const cellData = liveByName.get(sheetName);
    if (!cellData) continue; // sheet renamed/deleted in-session — nothing comparable
    for (const [key, saved] of cells) {
      if (saved.f === undefined || saved.v === undefined) continue;
      if (checked >= cap) {
        truncated = true;
        break outer;
      }
      const comma = key.indexOf(",");
      const r = key.slice(0, comma);
      const c = key.slice(comma + 1);
      const live = cellData[r]?.[c];
      if (!live) continue; // cleared in-session — a modification, not divergence
      const liveF = typeof live.f === "string" ? live.f : null;
      if (!sameFormula(liveF, saved.f)) continue; // edited in-session — not comparable
      let liveV = live.v;
      if (live.t === 3 && typeof liveV === "number") liveV = liveV !== 0;
      if (liveV === "") liveV = null;
      checked++;
      if (!isEmpty(saved.v) && !isExcelError(saved.v) && liveValueBroken(liveV, saved.v)) {
        divergent++;
        perSheet.set(sheetName, (perSheet.get(sheetName) ?? 0) + 1);
      }
    }
  }

  const bySheet = Array.from(perSheet.entries())
    .map(([sheet, n]) => ({ sheet, divergent: n }))
    .sort((a, b) => b.divergent - a.divergent);
  return { checked, divergent, bySheet, truncated };
}

/**
 * The "Calc health" context line shipped to the agent (stable bytes — the
 * health is computed once per loaded workbook and cached, so this never
 * jitters the cached context block). Empty string when there is nothing
 * worth warning about.
 */
export function formatCalcHealthLine(health: CalcHealth | null): string {
  if (!health || health.divergent === 0) return "";
  const top = health.bySheet
    .slice(0, 5)
    .map((s) => `${s.sheet}: ${s.divergent}`)
    .join(", ");
  return (
    `Live evaluation is unusable (0, empty, or error) vs the file's last-saved values on ` +
    `${health.divergent} of ${health.checked} checked formula cells (${top}). GridPath's engine ` +
    `cannot fully evaluate some constructs this workbook uses (e.g. INDIRECT chains, external ` +
    `references, add-in functions) — the FILE is sound; this is an evaluation limitation, not ` +
    `a defect in the user's model. read_range returns \`file_saved\` alongside the live value ` +
    `on those cells: prefer file_saved for reporting and comparisons, and do not build formula ` +
    `links onto affected cells (see rule 17b).`
  );
}
