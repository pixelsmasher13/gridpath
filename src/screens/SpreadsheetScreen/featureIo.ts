/**
 * Conditional formatting, data validation and cell notes — conversion between
 * ExcelJS's parsed worksheet models and Univer's snapshot plugin resources.
 *
 * Load: UniverGrid extracts the ExcelJS models and calls the excel→univer
 * converters here to seed `IWorkbookData.resources`, so imported workbooks
 * actually RENDER their red-negative rules, dropdowns and notes instead of
 * only preserving them in the file.
 *
 * Save: the surgical patcher can't rewrite these XML parts, so UniverGrid
 * compares the live resource state against the load-time baseline
 * (`perSheetFeatureState` / `featureDriftDetail`). Untouched state keeps the
 * surgical path (original parts survive byte-identical); edited state routes
 * to the gated ExcelJS full export, where `applyFeatureMirror` writes the
 * live rules back via the univer→excel converters.
 *
 * Pure functions only — no React, no Univer imports, and ExcelJS objects are
 * duck-typed (never imported) — so this module is trivially testable, same
 * contract as surgicalPatch.ts.
 */

import { resolveXlsxColor } from "./xlsxColors";

/** Univer snapshot resource keys (the `name` of entries in `resources: []`). */
export const CF_RESOURCE_NAME = "SHEET_CONDITIONAL_FORMATTING_PLUGIN";
export const DV_RESOURCE_NAME = "SHEET_DATA_VALIDATION_PLUGIN";
export const NOTE_RESOURCE_NAME = "SHEET_NOTE_PLUGIN";

export interface UniverRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

// ---------------------------------------------------------------------------
// A1 helpers (local copies — surgicalPatch/UniverGrid keep theirs private)
// ---------------------------------------------------------------------------

const MAX_ROW = 1_048_576;

function colLettersToIndex(letters: string): number | null {
  if (!letters || letters.length > 3) return null;
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const d = ch.charCodeAt(0) - 64;
    if (d < 1 || d > 26) return null;
    n = n * 26 + d;
  }
  return n - 1;
}

function colIndexToLetters(c: number): string {
  let out = "";
  for (;;) {
    out = String.fromCharCode(65 + (c % 26)) + out;
    if (c < 26) break;
    c = Math.floor(c / 26) - 1;
  }
  return out;
}

/**
 * Parse one A1 ref ("B2", "$A$1:$C$10", full-column "A:C") into a Univer
 * IRange. `$` anchors are irrelevant to a rectangle and stripped. Returns
 * null on anything malformed so callers can skip that ref.
 */
export function parseA1Range(s: string): UniverRange | null {
  const clean = s.replace(/\$/g, "").trim();
  if (!clean) return null;
  const parts = clean.split(":");
  if (parts.length > 2) return null;
  const decode = (ref: string): { row: number | null; col: number | null } | null => {
    const m = /^([A-Za-z]{0,3})(\d{0,7})$/.exec(ref.trim());
    if (!m || (!m[1] && !m[2])) return null;
    const col = m[1] ? colLettersToIndex(m[1]) : null;
    if (m[1] && col === null) return null;
    const row = m[2] ? parseInt(m[2], 10) - 1 : null;
    if (m[2] && (row === null || row < 0 || row >= MAX_ROW)) return null;
    return { row, col };
  };
  const a = decode(parts[0]);
  const b = parts.length === 2 ? decode(parts[1]) : a;
  if (!a || !b) return null;
  // Full-column ("A:C") / full-row ("1:3") refs default the open axis.
  const ar = a.row ?? 0;
  const br = b.row ?? MAX_ROW - 1;
  const ac = a.col ?? 0;
  const bc = b.col ?? 16_383;
  if (a.col === null && b.col === null && a.row === null) return null;
  return {
    startRow: Math.min(ar, br),
    endRow: Math.max(ar, br),
    startColumn: Math.min(ac, bc),
    endColumn: Math.max(ac, bc),
  };
}

/**
 * Ranges spanning more rows/cols than this are treated as "applies to the
 * whole column/row" (Excel files routinely carry CF/validation on A:A =
 * 1,048,576 rows). Univer evaluates whole-range rules (duplicates, rank,
 * color scales) lazily as cells scroll into view — a million-row range
 * makes that first evaluation scan a million rows. Clamping to the used
 * area plus a margin keeps the rendered behavior identical for real data
 * while killing the pathological scan. Hand-authored ranges below the
 * span threshold are never touched.
 */
const CLAMP_ROW_SPAN = 50_000;
const CLAMP_COL_SPAN = 1_000;
const CLAMP_ROW_MARGIN = 500;
const CLAMP_COL_MARGIN = 50;

export interface UsedBounds {
  maxRow: number;
  maxCol: number;
}

function clampRangeToUsed(r: UniverRange, bounds: UsedBounds): UniverRange {
  let endRow = r.endRow;
  let endColumn = r.endColumn;
  if (endRow - r.startRow + 1 > CLAMP_ROW_SPAN) {
    endRow = Math.max(r.startRow, Math.min(endRow, bounds.maxRow + CLAMP_ROW_MARGIN));
  }
  if (endColumn - r.startColumn + 1 > CLAMP_COL_SPAN) {
    endColumn = Math.max(r.startColumn, Math.min(endColumn, bounds.maxCol + CLAMP_COL_MARGIN));
  }
  if (endRow === r.endRow && endColumn === r.endColumn) return r;
  return { startRow: r.startRow, endRow, startColumn: r.startColumn, endColumn };
}

export function rangeToA1(r: UniverRange): string {
  const tl = `${colIndexToLetters(r.startColumn)}${r.startRow + 1}`;
  if (r.startRow === r.endRow && r.startColumn === r.endColumn) return tl;
  return `${tl}:${colIndexToLetters(r.endColumn)}${r.endRow + 1}`;
}

function cellA1(row: number, col: number): string {
  return `${colIndexToLetters(col)}${row + 1}`;
}

function cssToArgb(css: string): string {
  const hex = css.replace(/^#/, "").toUpperCase();
  return hex.length === 6 ? `FF${hex}` : hex;
}

// ---------------------------------------------------------------------------
// Conditional formatting: ExcelJS → Univer
// ---------------------------------------------------------------------------

/** One entry of ExcelJS's `worksheet.conditionalFormattings`. */
export interface ExcelCfEntry {
  ref: string;
  rules: any[];
}

/**
 * ExcelJS dxf style (font/fill/numFmt) → Univer IStyleBase. dxf pattern
 * fills carry the highlight color in bgColor (unlike ordinary cell fills,
 * which use fgColor), so both are checked.
 */
function dxfToUniverStyle(style: any, palette: string[] | null): any {
  const out: any = {};
  const font = style?.font;
  if (font?.bold) out.bl = 1;
  if (font?.italic) out.it = 1;
  if (font?.underline) out.ul = { s: 1 };
  if (font?.strike) out.st = { s: 1 };
  const fontColor = font?.color ? resolveXlsxColor(font.color, palette) : null;
  if (fontColor) out.cl = { rgb: fontColor };
  const fill = style?.fill;
  const fillColor = fill
    ? resolveXlsxColor(fill.fgColor, palette) ?? resolveXlsxColor(fill.bgColor, palette)
    : null;
  if (fillColor) out.bg = { rgb: fillColor };
  const numFmt = typeof style?.numFmt === "string" ? style.numFmt : style?.numFmt?.formatCode;
  if (typeof numFmt === "string" && numFmt) out.n = { pattern: numFmt };
  return out;
}

/** Univer IStyleBase → ExcelJS dxf style (inverse of dxfToUniverStyle). */
function univerStyleToDxf(s: any): any {
  const out: any = {};
  const font: any = {};
  if (s?.bl) font.bold = true;
  if (s?.it) font.italic = true;
  if (s?.ul?.s) font.underline = true;
  if (s?.st?.s) font.strike = true;
  if (s?.cl?.rgb) font.color = { argb: cssToArgb(s.cl.rgb) };
  if (Object.keys(font).length) out.font = font;
  if (s?.bg?.rgb) {
    const argb = cssToArgb(s.bg.rgb);
    // dxf convention: highlight color rides in bgColor. fgColor is set too
    // so readers that check only one side still see the fill.
    out.fill = { type: "pattern", pattern: "solid", bgColor: { argb }, fgColor: { argb } };
  }
  if (s?.n?.pattern) out.numFmt = s.n.pattern;
  return out;
}

const CF_NUMBER_OPERATORS = new Set([
  "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual",
  "between", "notBetween", "equal", "notEqual",
]);

const CF_TIME_PERIODS = new Set([
  "today", "yesterday", "tomorrow", "last7Days", "thisMonth", "lastMonth",
  "nextMonth", "thisWeek", "lastWeek", "nextWeek",
]);

function cfvoToValueConfig(cfvo: any): any {
  const type = String(cfvo?.type ?? "");
  const map: Record<string, string> = {
    min: "min", autoMin: "min", max: "max", autoMax: "max",
    num: "num", percent: "percent", percentile: "percentile", formula: "formula",
  };
  const t = map[type] ?? "num";
  const out: any = { type: t };
  if (cfvo?.value !== undefined && cfvo?.value !== null) {
    const n = Number(cfvo.value);
    out.value = Number.isFinite(n) && t !== "formula" ? n : cfvo.value;
  }
  return out;
}

/** Pull the literal text back out of Excel's generated text-rule formulas. */
function textFromCfFormula(operator: string, formula: string): string | null {
  if (typeof formula !== "string") return null;
  if (operator === "containsText" || operator === "notContainsText") {
    const m = /SEARCH\("((?:[^"]|"")*)"/.exec(formula);
    return m ? m[1].replace(/""/g, '"') : null;
  }
  // LEFT(A1,3)="foo" / RIGHT(A1,3)="foo"
  const m = /=\s*"((?:[^"]|"")*)"\s*$/.exec(formula);
  return m ? m[1].replace(/""/g, '"') : null;
}

const prefixEq = (f: string) => (f.startsWith("=") ? f : `=${f}`);

/**
 * Convert one ExcelJS CF rule to a Univer rule config, or null when the
 * shape has no Univer mapping (the caller records it as dropped). Style is
 * omitted for the non-highlight families (colorScale/dataBar), which carry
 * their colors in their own config.
 */
function convertExcelCfRule(rule: any, ranges: UniverRange[], palette: string[] | null): any | null {
  const style = dxfToUniverStyle(rule?.style, palette);
  const hl = (subType: string, extra: any) => ({ type: "highlightCell", subType, style, ...extra });
  const tl = ranges.length ? cellA1(ranges[0].startRow, ranges[0].startColumn) : "A1";
  const formulae: any[] = Array.isArray(rule?.formulae) ? rule.formulae : [];

  switch (rule?.type) {
    case "expression": {
      const f = formulae[0];
      if (typeof f !== "string" || !f) return null;
      return hl("formula", { value: prefixEq(f) });
    }
    case "cellIs": {
      const op = String(rule.operator ?? "");
      if (!CF_NUMBER_OPERATORS.has(op)) return null;
      const nums = formulae.map((f: any) => Number(f));
      const isNumeric = formulae.length > 0 && nums.every((n: number) => Number.isFinite(n));
      if (isNumeric) {
        if (op === "between" || op === "notBetween") {
          if (nums.length < 2) return null;
          return hl("number", { operator: op, value: [nums[0], nums[1]] });
        }
        return hl("number", { operator: op, value: nums[0] });
      }
      // Non-numeric comparand (a ref or text) — express as the equivalent
      // formula rule, which is what Excel evaluates anyway.
      const f0 = String(formulae[0] ?? "");
      if (!f0) return null;
      const f1 = String(formulae[1] ?? "");
      const cmp: Record<string, string> = {
        greaterThan: ">", greaterThanOrEqual: ">=", lessThan: "<",
        lessThanOrEqual: "<=", equal: "=", notEqual: "<>",
      };
      if (op === "between") return hl("formula", { value: `=AND(${tl}>=${f0},${tl}<=${f1})` });
      if (op === "notBetween") return hl("formula", { value: `=OR(${tl}<${f0},${tl}>${f1})` });
      return hl("formula", { value: `=${tl}${cmp[op]}${f0}` });
    }
    case "containsText": {
      // ExcelJS folds containsText/containsBlanks/notContainsBlanks/
      // containsErrors/notContainsErrors into type=containsText + operator.
      const op = String(rule.operator ?? "containsText");
      if (op === "containsBlanks" || op === "notContainsBlanks" ||
          op === "containsErrors" || op === "notContainsErrors") {
        return hl("text", { operator: op });
      }
      const text = rule.text ?? textFromCfFormula(op, formulae[0]);
      if (typeof text === "string" && text.length) {
        return hl("text", { operator: op, value: text });
      }
      const f = formulae[0];
      return typeof f === "string" && f ? hl("formula", { value: prefixEq(f) }) : null;
    }
    case "beginsWith":
    case "endsWith":
    case "notContainsText": {
      const text = rule.text ?? textFromCfFormula(rule.type, formulae[0]);
      if (typeof text === "string" && text.length) {
        return hl("text", { operator: rule.type, value: text });
      }
      const f = formulae[0];
      return typeof f === "string" && f ? hl("formula", { value: prefixEq(f) }) : null;
    }
    case "timePeriod": {
      const period = String(rule.timePeriod ?? "");
      if (!CF_TIME_PERIODS.has(period)) return null;
      return hl("timePeriod", { operator: period });
    }
    case "top10":
      return hl("rank", {
        isBottom: !!rule.bottom,
        isPercent: !!rule.percent,
        value: Number.isFinite(Number(rule.rank)) ? Number(rule.rank) : 10,
      });
    case "aboveAverage":
      return hl("average", { operator: rule.aboveAverage === false ? "lessThan" : "greaterThan" });
    case "duplicateValues":
      return hl("duplicateValues", {});
    case "uniqueValues":
      return hl("uniqueValues", {});
    case "colorScale": {
      const cfvo: any[] = Array.isArray(rule.cfvo) ? rule.cfvo : [];
      const color: any[] = Array.isArray(rule.color) ? rule.color : [];
      if (cfvo.length < 2) return null;
      return {
        type: "colorScale",
        config: cfvo.map((v, i) => ({
          index: i,
          color: resolveXlsxColor(color[i], palette) ?? "#FFFFFF",
          value: cfvoToValueConfig(v),
        })),
      };
    }
    case "dataBar": {
      const cfvo: any[] = Array.isArray(rule.cfvo) ? rule.cfvo : [];
      return {
        type: "dataBar",
        isShowValue: rule.showValue !== false,
        config: {
          min: cfvoToValueConfig(cfvo[0] ?? { type: "min" }),
          max: cfvoToValueConfig(cfvo[1] ?? { type: "max" }),
          isGradient: rule.gradient !== false,
          positiveColor: resolveXlsxColor(rule.color, palette) ?? "#638EC6",
          // Univer's field name for the negative bar color (sic).
          nativeColor: "#FF0000",
        },
      };
    }
    default:
      // iconSet and anything unknown: no faithful Univer mapping. The rule
      // stays in the file via the surgical path; it just doesn't render.
      return null;
  }
}

/**
 * Convert a worksheet's ExcelJS conditionalFormattings into Univer's rule
 * array (the per-sheet value inside the CF resource). Rules are ordered by
 * Excel priority (1 = strongest) because Univer applies by array order.
 */
export function excelCfToUniverRules(
  entries: ExcelCfEntry[] | undefined | null,
  palette: string[] | null,
  bounds?: UsedBounds,
): { rules: any[]; dropped: string[] } {
  const collected: Array<{ priority: number; ranges: UniverRange[]; rule: any; stopIfTrue: boolean }> = [];
  const dropped: string[] = [];
  for (const entry of entries ?? []) {
    const ranges = String(entry?.ref ?? "")
      .split(/\s+/)
      .map(parseA1Range)
      .filter((r): r is UniverRange => !!r)
      .map((r) => (bounds ? clampRangeToUsed(r, bounds) : r));
    if (!ranges.length) continue;
    for (const rule of entry.rules ?? []) {
      const converted = convertExcelCfRule(rule, ranges, palette);
      if (!converted) {
        dropped.push(String(rule?.type ?? "unknown"));
        continue;
      }
      collected.push({
        priority: Number.isFinite(Number(rule?.priority)) ? Number(rule.priority) : Number.MAX_SAFE_INTEGER,
        ranges,
        rule: converted,
        stopIfTrue: !!rule?.stopIfTrue,
      });
    }
  }
  collected.sort((a, b) => a.priority - b.priority);
  return {
    rules: collected.map((c, i) => ({
      cfId: `cf-${i + 1}`,
      ranges: c.ranges,
      stopIfTrue: c.stopIfTrue,
      rule: c.rule,
    })),
    dropped,
  };
}

// ---------------------------------------------------------------------------
// Conditional formatting: Univer → ExcelJS (full-export mirror)
// ---------------------------------------------------------------------------

function escapeFormulaText(s: string): string {
  return s.replace(/"/g, '""');
}

/**
 * Convert Univer CF rules back into ExcelJS conditionalFormattings entries.
 * Only reached for sheets whose CF state changed in-session, on the gated
 * full-export path. Families ExcelJS cannot write (duplicate/unique values,
 * icon sets) are reported in `dropped` for the caller to log.
 */
export function univerCfToExcelCf(rules: any[] | undefined | null): {
  entries: ExcelCfEntry[];
  dropped: string[];
} {
  const entries: ExcelCfEntry[] = [];
  const dropped: string[] = [];
  let priority = 1;
  for (const wrapper of rules ?? []) {
    const ranges: UniverRange[] = Array.isArray(wrapper?.ranges) ? wrapper.ranges : [];
    if (!ranges.length) continue;
    const ref = ranges.map(rangeToA1).join(" ");
    const tl = cellA1(ranges[0].startRow, ranges[0].startColumn);
    const r = wrapper?.rule ?? {};
    const style = univerStyleToDxf(r.style);
    const base = { priority: priority, ref, style };
    let excelRule: any | null = null;

    if (r.type === "highlightCell") {
      switch (r.subType) {
        case "formula":
          excelRule = { ...base, type: "expression", formulae: [String(r.value ?? "").replace(/^=/, "")] };
          break;
        case "number": {
          const op = String(r.operator ?? "");
          if (!CF_NUMBER_OPERATORS.has(op)) break;
          const formulae = Array.isArray(r.value) ? [r.value[0], r.value[1]] : [r.value];
          excelRule = { ...base, type: "cellIs", operator: op, formulae };
          break;
        }
        case "text": {
          const op = String(r.operator ?? "");
          const v = String(r.value ?? "");
          // ExcelJS's writer natively handles the containsText family; the
          // rest are emitted as equivalent expression rules.
          if (op === "containsText" || op === "containsBlanks" || op === "notContainsBlanks" ||
              op === "containsErrors" || op === "notContainsErrors") {
            excelRule = { ...base, type: "containsText", operator: op, text: v || undefined };
          } else {
            const esc = escapeFormulaText(v);
            const byOp: Record<string, string> = {
              beginsWith: `LEFT(${tl},LEN("${esc}"))="${esc}"`,
              endsWith: `RIGHT(${tl},LEN("${esc}"))="${esc}"`,
              notContainsText: `ISERROR(SEARCH("${esc}",${tl}))`,
              equal: `${tl}="${esc}"`,
              notEqual: `${tl}<>"${esc}"`,
            };
            const f = byOp[op];
            if (f) excelRule = { ...base, type: "expression", formulae: [f] };
          }
          break;
        }
        case "timePeriod":
          if (CF_TIME_PERIODS.has(String(r.operator))) {
            excelRule = { ...base, type: "timePeriod", timePeriod: r.operator };
          }
          break;
        case "rank":
          excelRule = {
            ...base,
            type: "top10",
            rank: Number.isFinite(Number(r.value)) ? Number(r.value) : 10,
            percent: !!r.isPercent,
            bottom: !!r.isBottom,
          };
          break;
        case "average": {
          const op = String(r.operator ?? "greaterThan");
          excelRule = { ...base, type: "aboveAverage", aboveAverage: !op.startsWith("lessThan") };
          break;
        }
        default:
          break;
      }
    } else if (r.type === "colorScale") {
      const config: any[] = Array.isArray(r.config) ? r.config : [];
      if (config.length >= 2) {
        excelRule = {
          ...base,
          type: "colorScale",
          cfvo: config.map((c) => ({
            type: c?.value?.type ?? "num",
            ...(c?.value?.value !== undefined ? { value: c.value.value } : {}),
          })),
          color: config.map((c) => ({ argb: cssToArgb(String(c?.color ?? "#FFFFFF")) })),
        };
      }
    } else if (r.type === "dataBar") {
      const cfg = r.config ?? {};
      excelRule = {
        ...base,
        type: "dataBar",
        cfvo: [cfg.min ?? { type: "min" }, cfg.max ?? { type: "max" }].map((v: any) => ({
          type: v?.type ?? "num",
          ...(v?.value !== undefined ? { value: v.value } : {}),
        })),
        color: { argb: cssToArgb(String(cfg.positiveColor ?? "#638EC6")) },
        gradient: cfg.isGradient !== false,
        showValue: r.isShowValue !== false,
      };
    }

    if (!excelRule) {
      dropped.push(String(r.subType ?? r.type ?? "unknown"));
      continue;
    }
    entries.push({ ref, rules: [excelRule] });
    priority += 1;
  }
  return { entries, dropped };
}

// ---------------------------------------------------------------------------
// Data validation: ExcelJS → Univer
// ---------------------------------------------------------------------------

const DV_TYPES = new Set(["list", "whole", "decimal", "date", "time", "textLength", "custom"]);
const DV_OPERATORS = new Set([
  "between", "notBetween", "equal", "notEqual",
  "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual",
]);

/**
 * Merge a set of single-cell addresses back into rectangles. ExcelJS expands
 * every dv sqref into per-cell model keys; feeding Univer one range per cell
 * would explode rule counts (a 10k-row dropdown column = 10k ranges).
 * Greedy: horizontal runs per row, then vertically merge identical runs.
 */
export function coalesceAddressesToRanges(addresses: string[]): UniverRange[] {
  const byRow = new Map<number, number[]>();
  for (const a of addresses) {
    const r = parseA1Range(a);
    if (!r || r.startRow !== r.endRow || r.startColumn !== r.endColumn) continue;
    let cols = byRow.get(r.startRow);
    if (!cols) byRow.set(r.startRow, (cols = []));
    cols.push(r.startColumn);
  }
  // Row → sorted list of [startCol, endCol] runs.
  const runsByRow = new Map<number, Array<[number, number]>>();
  for (const [row, cols] of byRow) {
    cols.sort((a, b) => a - b);
    const runs: Array<[number, number]> = [];
    for (const c of cols) {
      const last = runs[runs.length - 1];
      if (last && c === last[1] + 1) last[1] = c;
      else if (!last || c > last[1]) runs.push([c, c]);
    }
    runsByRow.set(row, runs);
  }
  const rows = [...runsByRow.keys()].sort((a, b) => a - b);
  const out: UniverRange[] = [];
  // Open rectangles keyed by "startCol,endCol" awaiting vertical extension.
  let open = new Map<string, UniverRange>();
  for (const row of rows) {
    const next = new Map<string, UniverRange>();
    for (const [c0, c1] of runsByRow.get(row)!) {
      const key = `${c0},${c1}`;
      const prev = open.get(key);
      if (prev && prev.endRow === row - 1) {
        prev.endRow = row;
        next.set(key, prev);
      } else {
        const rect = { startRow: row, endRow: row, startColumn: c0, endColumn: c1 };
        out.push(rect);
        next.set(key, rect);
      }
    }
    open = next;
  }
  return out;
}

/** Excel dv formula → Univer formula1/formula2 string. */
function dvFormulaToUniver(f: any, isList: boolean): string | undefined {
  if (f === undefined || f === null) return undefined;
  if (typeof f === "number") return String(f);
  const s = String(f).trim();
  if (!s) return undefined;
  if (isList) {
    // `"a,b,c"` is Excel's literal-list encoding; Univer wants it unquoted.
    if (/^".*"$/s.test(s)) return s.slice(1, -1);
    return s.startsWith("=") ? s : `=${s}`;
  }
  if (/^[-+]?\d+(\.\d+)?$/.test(s)) return s;
  return s.startsWith("=") ? s : `=${s}`;
}

/**
 * Rules that coalesce into more rectangles than this are dropped on import.
 * Univer's data-validation RuleMatrix runs rectangle algebra (mergeRanges /
 * subtractMulti, both quadratic-ish in range count) for every hydrated rule
 * against every earlier rule — a vendor artifact with ~1,000 cells scattered
 * every 256 columns (Lotus-compat copy residue, seen in the wild) freezes
 * createWorkbook for minutes. Hand-authored validations are a handful of
 * rectangles; the dropped rules stay in the file because the surgical save
 * path never rewrites untouched dataValidation XML.
 */
const DV_IMPORT_RANGE_CAP = 200;

/**
 * ExcelJS `worksheet.dataValidations.model` (per-cell keys, shared rule
 * objects) → Univer's per-sheet rule array. Cells sharing one parsed rule
 * object are re-grouped into rectangles.
 */
export function excelDvToUniverRules(
  dvModel: Record<string, any> | undefined | null,
  bounds?: UsedBounds,
): any[] {
  if (!dvModel) return [];
  const groups = new Map<any, string[]>();
  for (const [address, rule] of Object.entries(dvModel)) {
    if (!rule || typeof rule !== "object") continue;
    let list = groups.get(rule);
    if (!list) groups.set(rule, (list = []));
    list.push(address);
  }
  const out: any[] = [];
  let n = 0;
  for (const [rule, addresses] of groups) {
    const type = String(rule.type ?? "");
    if (!DV_TYPES.has(type)) continue; // 'any'/'none' impose nothing
    const ranges = coalesceAddressesToRanges(addresses).map((r) =>
      bounds ? clampRangeToUsed(r, bounds) : r,
    );
    if (!ranges.length) continue;
    if (ranges.length > DV_IMPORT_RANGE_CAP) {
      console.warn(
        `[dv] dropping ${type} validation on import: ${ranges.length} fragmented ranges ` +
          `(cap ${DV_IMPORT_RANGE_CAP}) — rendering it would freeze the grid; the file keeps it`,
      );
      continue;
    }
    n += 1;
    const isList = type === "list";
    const univerRule: any = {
      uid: `dv-${n}`,
      type,
      ranges,
      allowBlank: rule.allowBlank !== false,
    };
    const f1 = dvFormulaToUniver(rule.formulae?.[0], isList);
    const f2 = dvFormulaToUniver(rule.formulae?.[1], isList);
    if (f1 !== undefined) univerRule.formula1 = f1;
    if (f2 !== undefined) univerRule.formula2 = f2;
    if (!isList && DV_OPERATORS.has(String(rule.operator))) univerRule.operator = rule.operator;
    if (isList) {
      // OOXML's showDropDown attribute is inverted: truthy means SUPPRESS
      // the in-cell dropdown. Univer's showDropDown means "show it".
      univerRule.showDropDown = !rule.showDropDown;
    }
    if (rule.showErrorMessage) {
      univerRule.showErrorMessage = true;
      if (rule.error) univerRule.error = String(rule.error);
      if (rule.errorTitle) univerRule.errorTitle = String(rule.errorTitle);
      const styleMap: Record<string, number> = { information: 0, stop: 1, warning: 2 };
      const es = styleMap[String(rule.errorStyle ?? "")];
      if (es !== undefined) univerRule.errorStyle = es;
    }
    if (rule.showInputMessage) {
      univerRule.showInputMessage = true;
      if (rule.prompt) univerRule.prompt = String(rule.prompt);
      if (rule.promptTitle) univerRule.promptTitle = String(rule.promptTitle);
    }
    out.push(univerRule);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Data validation: Univer → ExcelJS (full-export mirror)
// ---------------------------------------------------------------------------

/** Beyond this many cells a rule is dropped (with a log) instead of exploding
 *  the per-cell model — e.g. a validation applied to a whole 1M-row column. */
const DV_EXPORT_CELL_CAP = 50_000;

export function univerDvToExcelDvModel(rules: any[] | undefined | null): {
  model: Record<string, any>;
  dropped: string[];
} {
  const model: Record<string, any> = {};
  const dropped: string[] = [];
  for (const rule of rules ?? []) {
    const type = String(rule?.type ?? "");
    // listMultiple degrades to a plain list; checkbox/any/none have no
    // Excel dv equivalent.
    const excelType = type === "listMultiple" ? "list" : type;
    if (!DV_TYPES.has(excelType)) {
      dropped.push(type || "unknown");
      continue;
    }
    const isList = excelType === "list";
    const backFormula = (f: any): any => {
      if (f === undefined || f === null || f === "") return undefined;
      const s = String(f);
      if (isList) {
        if (s.startsWith("=")) return s.slice(1);
        return `"${s.replace(/"/g, '""')}"`;
      }
      return s.startsWith("=") ? s.slice(1) : s;
    };
    const excelRule: any = {
      type: excelType,
      allowBlank: rule.allowBlank !== false,
      formulae: [backFormula(rule.formula1), backFormula(rule.formula2)].filter(
        (f) => f !== undefined,
      ),
    };
    if (!isList && DV_OPERATORS.has(String(rule.operator))) excelRule.operator = rule.operator;
    if (isList && rule.showDropDown === false) excelRule.showDropDown = true; // inverted, see import
    if (rule.showErrorMessage) {
      excelRule.showErrorMessage = true;
      if (rule.error) excelRule.error = String(rule.error);
      if (rule.errorTitle) excelRule.errorTitle = String(rule.errorTitle);
      const styleMap: Record<number, string> = { 0: "information", 1: "stop", 2: "warning" };
      const es = styleMap[Number(rule.errorStyle)];
      if (es) excelRule.errorStyle = es;
    }
    if (rule.showInputMessage) {
      excelRule.showInputMessage = true;
      if (rule.prompt) excelRule.prompt = String(rule.prompt);
      if (rule.promptTitle) excelRule.promptTitle = String(rule.promptTitle);
    }

    const ranges: UniverRange[] = Array.isArray(rule?.ranges) ? rule.ranges : [];
    const totalCells = ranges.reduce(
      (sum, r) => sum + (r.endRow - r.startRow + 1) * (r.endColumn - r.startColumn + 1),
      0,
    );
    if (totalCells === 0) continue;
    if (totalCells > DV_EXPORT_CELL_CAP) {
      dropped.push(`${excelType} (${totalCells} cells)`);
      continue;
    }
    for (const r of ranges) {
      for (let row = r.startRow; row <= r.endRow; row++) {
        for (let col = r.startColumn; col <= r.endColumn; col++) {
          model[cellA1(row, col)] = excelRule;
        }
      }
    }
  }
  return { model, dropped };
}

// ---------------------------------------------------------------------------
// Notes
// ---------------------------------------------------------------------------

export interface ImportedNote {
  sheetId: string;
  row: number;
  col: number;
  text: string;
}

/** Flatten an ExcelJS note (plain string or rich-text model) to its text. */
export function noteTextFromExcelJs(note: any): string | null {
  if (typeof note === "string") return note.length ? note : null;
  const texts = note?.texts;
  if (Array.isArray(texts)) {
    const joined = texts.map((t: any) => String(t?.text ?? "")).join("");
    return joined.length ? joined : null;
  }
  return null;
}

/** Excel's default comment box is roughly this size at 96 DPI. */
const NOTE_DEFAULT_W = 216;
const NOTE_DEFAULT_H = 92;

export function notesToUniverResource(
  notes: ImportedNote[],
): Record<string, Record<number, Record<number, any>>> {
  const out: Record<string, Record<number, Record<number, any>>> = {};
  for (const n of notes) {
    const sheet = (out[n.sheetId] ??= {});
    const row = (sheet[n.row] ??= {});
    row[n.col] = {
      id: `note-${n.row}-${n.col}`,
      row: n.row,
      col: n.col,
      width: NOTE_DEFAULT_W,
      height: NOTE_DEFAULT_H,
      note: n.text,
      show: false,
    };
  }
  return out;
}

/** Univer note resource (one sheet's slice) → flat {row, col, text} list. */
export function univerNotesToList(
  sheetNotes: Record<string, Record<string, any>> | undefined | null,
): Array<{ row: number; col: number; text: string }> {
  const out: Array<{ row: number; col: number; text: string }> = [];
  for (const [rowKey, cols] of Object.entries(sheetNotes ?? {})) {
    if (!cols || typeof cols !== "object") continue;
    for (const [colKey, note] of Object.entries(cols)) {
      const text = typeof note?.note === "string" ? note.note : null;
      if (text === null) continue;
      out.push({ row: Number(rowKey), col: Number(colKey), text });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Iterative calculation (circular references)
// ---------------------------------------------------------------------------

/**
 * Excel's iterative-calculation setting from xl/workbook.xml's <calcPr>:
 * `iterate="1"` opts the workbook into resolving circular references by
 * iteration, `iterateCount` bounds the rounds (Excel default 100). ExcelJS
 * parses <calcPr> but DISCARDS its attributes (its xform reads none), so
 * the caller feeds us the raw workbook.xml text instead. Returns null when
 * iterative calculation is off — Univer should then stay at its 1-pass
 * default, which mirrors Excel's own refusal to converge circular refs
 * without the opt-in.
 */
export function parseIterativeCalc(workbookXml: string): { iterateCount: number } | null {
  const calcPr = /<calcPr\b[^>]*>/.exec(workbookXml)?.[0];
  if (!calcPr) return null;
  const iterate = /\biterate\s*=\s*"(1|true)"/.exec(calcPr);
  if (!iterate) return null;
  const count = Number(/\biterateCount\s*=\s*"(\d+)"/.exec(calcPr)?.[1] ?? 100);
  // Excel allows up to 32767 but stops early once the change per round drops
  // below iterateDelta — Univer's engine has NO such convergence exit and
  // burns every round on every recalc of the cyclic subgraph. Cap at Excel's
  // default of 100: models that genuinely need more rounds to converge are
  // pathological, and each extra round is a full recalc pass per keystroke.
  return { iterateCount: Math.min(Math.max(1, count), 100) };
}

// ---------------------------------------------------------------------------
// Feature state (drift detection for the save path)
// ---------------------------------------------------------------------------

export type SheetFeatureState = { cf?: string; dv?: string; note?: string };

function stableStringify(value: any): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`).join(",")}}`;
}

function isEmptyState(v: any): boolean {
  if (v == null) return true;
  if (Array.isArray(v)) return v.length === 0;
  if (typeof v === "object") return Object.keys(v).length === 0;
  return false;
}

/**
 * Per-sheet normalized snapshot of the three feature resources. Key-order
 * differences are normalized away; empty per-sheet entries are treated the
 * same as absent ones (Univer serializes `{}`/`[]` where the import may have
 * had no resource at all).
 */
export function perSheetFeatureState(
  resources: Array<{ name: string; data: string }> | undefined | null,
): Map<string, SheetFeatureState> {
  const out = new Map<string, SheetFeatureState>();
  const fold = (resourceName: string, field: keyof SheetFeatureState) => {
    const entry = (resources ?? []).find((r) => r?.name === resourceName);
    if (!entry || typeof entry.data !== "string" || !entry.data) return;
    let parsed: any;
    try {
      parsed = JSON.parse(entry.data);
    } catch {
      return;
    }
    if (!parsed || typeof parsed !== "object") return;
    for (const [sheetId, value] of Object.entries(parsed)) {
      if (isEmptyState(value)) continue;
      const state = out.get(sheetId) ?? {};
      state[field] = stableStringify(value);
      out.set(sheetId, state);
    }
  };
  fold(CF_RESOURCE_NAME, "cf");
  fold(DV_RESOURCE_NAME, "dv");
  fold(NOTE_RESOURCE_NAME, "note");
  return out;
}

const FEATURE_LABEL: Record<keyof SheetFeatureState, string> = {
  cf: "conditional formatting",
  dv: "data validation",
  note: "cell notes",
};

/**
 * Human-readable description of what changed since the baseline, or null
 * when nothing did. Only sheets in `liveSheetNamesById` are compared — a
 * deleted sheet's feature state is gone WITH the sheet, which the patcher
 * already handles.
 */
export function featureDriftDetail(
  baseline: Map<string, SheetFeatureState> | null | undefined,
  current: Map<string, SheetFeatureState>,
  liveSheetNamesById: Map<string, string>,
): string | null {
  const changed = new Map<keyof SheetFeatureState, Set<string>>();
  for (const sheetId of liveSheetNamesById.keys()) {
    const b = baseline?.get(sheetId) ?? {};
    const c = current.get(sheetId) ?? {};
    for (const field of ["cf", "dv", "note"] as const) {
      if ((b[field] ?? null) !== (c[field] ?? null)) {
        let set = changed.get(field);
        if (!set) changed.set(field, (set = new Set()));
        set.add(liveSheetNamesById.get(sheetId) ?? sheetId);
      }
    }
  }
  if (!changed.size) return null;
  return [...changed.entries()]
    .map(([field, sheets]) => `${FEATURE_LABEL[field]} (${[...sheets].join(", ")})`)
    .join(", ");
}

// ---------------------------------------------------------------------------
// Full-export mirror
// ---------------------------------------------------------------------------

/**
 * Write CF / data-validation / note state from the live Univer snapshot into
 * the ExcelJS workbook before a full export. Replace-on-drift: a sheet whose
 * feature state still matches the load-time baseline is left alone, so
 * ExcelJS's own parsed models (which may hold rule types we can't express,
 * e.g. icon sets) round-trip untouched. Pass `baseline = null` to write
 * everything (fresh/untitled workbooks).
 *
 * `workbook` is a duck-typed ExcelJS workbook; failures on one sheet or one
 * family must never abort the export, so each block is best-effort.
 */
export function applyFeatureMirror(
  workbook: any,
  snapshot: any,
  baseline: Map<string, SheetFeatureState> | null,
): { dropped: string[] } {
  const dropped: string[] = [];
  const resources: Array<{ name: string; data: string }> = snapshot?.resources ?? [];
  const current = perSheetFeatureState(resources);
  const parseResource = (name: string): any => {
    const entry = resources.find((r) => r?.name === name);
    if (!entry?.data) return {};
    try {
      return JSON.parse(entry.data) ?? {};
    } catch {
      return {};
    }
  };
  const cfBySheet = parseResource(CF_RESOURCE_NAME);
  const dvBySheet = parseResource(DV_RESOURCE_NAME);
  const noteBySheet = parseResource(NOTE_RESOURCE_NAME);

  const sheets = snapshot?.sheets ?? {};
  for (const sheetId of Object.keys(sheets)) {
    const name = sheets[sheetId]?.name;
    if (!name) continue;
    const ws = workbook?.getWorksheet?.(name);
    if (!ws) continue;
    const cur = current.get(sheetId) ?? {};
    const base = baseline?.get(sheetId) ?? {};
    const dirty = (field: keyof SheetFeatureState) =>
      baseline === null || (base[field] ?? null) !== (cur[field] ?? null);

    if (dirty("cf")) {
      try {
        const { entries, dropped: d } = univerCfToExcelCf(cfBySheet[sheetId]);
        if (typeof ws.removeConditionalFormatting === "function") {
          ws.removeConditionalFormatting(() => false);
        } else {
          ws.conditionalFormattings = [];
        }
        for (const e of entries) ws.addConditionalFormatting?.(e);
        dropped.push(...d.map((t) => `cf:${t} (${name})`));
      } catch (e) {
        console.warn("[featureIo] cf mirror failed on", name, e);
      }
    }

    if (dirty("dv")) {
      try {
        const { model, dropped: d } = univerDvToExcelDvModel(dvBySheet[sheetId]);
        if (ws.dataValidations) ws.dataValidations.model = model;
        dropped.push(...d.map((t) => `dv:${t} (${name})`));
      } catch (e) {
        console.warn("[featureIo] dv mirror failed on", name, e);
      }
    }

    if (dirty("note")) {
      try {
        // Clear every existing note first (removed notes must not survive),
        // then re-apply from the live state. `_comment` is ExcelJS's private
        // storage, but there is no public "delete note" API and the model
        // getter reads exactly this field.
        ws.eachRow?.({ includeEmpty: true }, (row: any) => {
          row.eachCell?.({ includeEmpty: true }, (cell: any) => {
            if (cell._comment) cell._comment = undefined;
          });
        });
        for (const n of univerNotesToList(noteBySheet[sheetId])) {
          const cell = ws.getCell?.(n.row + 1, n.col + 1);
          if (cell) cell.note = n.text;
        }
      } catch (e) {
        console.warn("[featureIo] note mirror failed on", name, e);
      }
    }
  }
  return { dropped };
}
