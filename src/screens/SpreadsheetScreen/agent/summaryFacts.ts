/**
 * What create_summary_table needs to know about the data before the table
 * can be laid out: where each field is, which rows hold data, the distinct
 * keys as the cells hold them, and the number formats to carry over. All
 * of it comes from the calculation engine's row query (any sheet size), so
 * nothing is read cell by cell.
 */
import type { SummaryFacts, SummaryField, SummaryFn, SummaryKey, SummaryRequest } from "./summaryLayout";

export const MAX_SUMMARY_ROWS = 500;
export const MAX_SUMMARY_COLUMNS = 60;

type QueryAnswer = { ok: true; result: any } | { ok: false; error: string; message: string };

export interface SummaryIo {
  /** The engine's row query on a sheet of the workbook. */
  query: (sheet: string, query: Record<string, unknown>) => Promise<QueryAnswer>;
  /** The number format of a cell (0-based), once `prefetch` has loaded it. */
  numberFormat: (sheet: string, row: number, col: number) => string | null | undefined;
  prefetch: (sheet: string, row: number, col: number) => Promise<void>;
}

export type SummaryBounds = { firstRow?: number | null; lastRow?: number | null };

export type GatheredSummary =
  | { ok: true; facts: SummaryFacts; skippedBlank: number }
  | { ok: false; error: string; message: string };

const ENGINE_FN: Record<SummaryFn, string> = { sum: "sum", count: "count", average: "avg", min: "min", max: "max" };

function lettersToCol(letters: string): number {
  let c = 0;
  for (const ch of letters.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
}

/** Numbers first, in order; then text, ignoring case; then TRUE/FALSE. */
function byLabel(a: SummaryKey, b: SummaryKey): number {
  const rank = (v: SummaryKey) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === "number" && typeof b === "number") return a - b;
  return String(a).localeCompare(String(b), undefined, { sensitivity: "base", numeric: true });
}

export async function gatherSummaryFacts(req: SummaryRequest & SummaryBounds, io: SummaryIo): Promise<GatheredSummary> {
  const bounds = {
    ...(req.headerRow ? { header_row: req.headerRow } : {}),
    ...(req.firstRow ? { start_row: req.firstRow } : {}),
    ...(req.lastRow ? { end_row: req.lastRow } : {}),
  };
  const fail = (error: string, message: string): GatheredSummary => ({ ok: false, error, message });

  // One grouped query per field: its column, its distinct keys, and (for
  // the row field) the first value per group, to order the rows by.
  const first = req.values[0];
  const group = async (field: string, limit: number, withValue: boolean) => {
    const answer = await io.query(req.sourceSheet, {
      ...bounds,
      group_by: field,
      ...(withValue ? { aggregate: [{ column: first.field, fn: ENGINE_FN[first.fn] }] } : {}),
      limit: limit + 1,
    });
    if (!answer.ok) return answer;
    const r = answer.result;
    const letter = r?.group_by?.column;
    if (typeof letter !== "string") return { ok: false as const, error: "bad_query", message: `The engine could not group ${req.sourceSheet} by "${field}".` };
    const groups: any[] = Array.isArray(r.groups) ? r.groups : [];
    const keyed = groups.filter((g) => g.value !== undefined && g.value !== null && g.value !== "");
    const blank = groups.filter((g) => !keyed.includes(g)).reduce((n, g) => n + (Number(g.count) || 0), 0);
    return { ok: true as const, r, letter, header: typeof r.group_by.header === "string" && r.group_by.header ? r.group_by.header : `Column ${letter}`, keyed, blank, distinct: Number(r.group_count) || groups.length };
  };

  const rowsAnswer = await group(req.rows, MAX_SUMMARY_ROWS, true);
  if (!rowsAnswer.ok) return fail(rowsAnswer.error, rowsAnswer.message);
  if (rowsAnswer.distinct > MAX_SUMMARY_ROWS + 1 || rowsAnswer.keyed.length > MAX_SUMMARY_ROWS) {
    return fail(
      "too_many_groups",
      `"${rowsAnswer.header}" has ${rowsAnswer.distinct} distinct values; a summary table takes up to ${MAX_SUMMARY_ROWS}. Group by a field with fewer values (a category, a region, a year).`,
    );
  }
  if (rowsAnswer.keyed.length === 0) return fail("no_groups", `"${rowsAnswer.header}" holds no values in the data rows of ${req.sourceSheet}.`);
  const range = /^(\d+)-(\d+)$/.exec(String(rowsAnswer.r.row_range ?? ""));
  const firstRow = range ? Number(range[1]) : 0;
  const lastRow = range ? Number(range[2]) : 0;
  if (!range || lastRow < firstRow) return fail("no_data", `${req.sourceSheet} has no data rows below its header row.`);

  let colsAnswer: Awaited<ReturnType<typeof group>> | null = null;
  if (req.columns) {
    colsAnswer = await group(req.columns, MAX_SUMMARY_COLUMNS, false);
    if (!colsAnswer.ok) return fail(colsAnswer.error, colsAnswer.message);
    if (colsAnswer.distinct > MAX_SUMMARY_COLUMNS + 1 || colsAnswer.keyed.length > MAX_SUMMARY_COLUMNS) {
      return fail(
        "too_many_groups",
        `"${colsAnswer.header}" has ${colsAnswer.distinct} distinct values; a field across the top takes up to ${MAX_SUMMARY_COLUMNS}. Put it in 'rows' instead, or leave it out.`,
      );
    }
    if (colsAnswer.keyed.length === 0) return fail("no_groups", `"${colsAnswer.header}" holds no values in the data rows of ${req.sourceSheet}.`);
  }

  // Where each value field is.
  const measures: SummaryFacts["measures"] = [];
  for (const v of req.values) {
    const answer = await io.query(req.sourceSheet, { ...bounds, columns: [v.field], limit: 1 });
    if (!answer.ok) return fail(answer.error, answer.message);
    const found = answer.result?.columns?.[0];
    if (typeof found?.column !== "string") return fail("bad_query", `The engine could not find the field "${v.field}" on ${req.sourceSheet}.`);
    measures.push({ col: lettersToCol(found.column), header: typeof found.header === "string" && found.header ? found.header : `Column ${found.column}`, fn: v.fn, format: null });
  }

  // Number formats: a date key must read as a date, an amount as an amount.
  const sheet = typeof rowsAnswer.r.sheet === "string" ? rowsAnswer.r.sheet : req.sourceSheet;
  const formatOf = async (col: number): Promise<string | null> => {
    await io.prefetch(sheet, firstRow - 1, col);
    const f = io.numberFormat(sheet, firstRow - 1, col);
    return typeof f === "string" && f && !/^general$/i.test(f) && f !== "@" ? f : null;
  };
  const field = async (answer: Extract<Awaited<ReturnType<typeof group>>, { ok: true }>, sort: "label" | "value"): Promise<SummaryField> => {
    const keys: SummaryKey[] = answer.keyed.map((g) => g.value as SummaryKey);
    // The engine gives groups largest first by the value; by label is ours.
    if (sort === "label") keys.sort(byLabel);
    const col = lettersToCol(answer.letter);
    return { col, header: answer.header, keys, format: keys.some((k) => typeof k === "number") ? await formatOf(col) : null };
  };
  const rowField = await field(rowsAnswer, req.sort);
  const colField = colsAnswer && colsAnswer.ok ? await field(colsAnswer, "label") : null;
  for (const m of measures) m.format = await formatOf(m.col);

  return { ok: true, facts: { sheet, firstRow, lastRow, rowField, colField, measures }, skippedBlank: rowsAnswer.blank };
}
