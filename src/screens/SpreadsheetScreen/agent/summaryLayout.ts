/**
 * create_summary_table: a pivot-style summary written as ordinary formulas.
 *
 * The agent names a data sheet, the field to group by, optionally a second
 * field across the top, and what to total. The app (not the model) reads
 * the distinct keys from the engine and lays the table out: one label per
 * group, one SUMIFS / COUNTIFS / AVERAGEIFS / MINIFS / MAXIFS per cell,
 * totals taken straight from the data, and a check that the table still
 * accounts for every row. Every cell is a formula a person can click into,
 * and the table follows edits to the data it reads.
 *
 * Pure function: (request, facts) → mutations + report, so the layout is
 * testable without a grid. The mutations go through the standard apply
 * path: review, undo, save and readback are the same as for any write.
 */
import type { UniverMutation } from "../types";
import { expandA1Range } from "./toolToMutation";
import { colLetters } from "./reviewGroups";

export type SummaryFn = "sum" | "count" | "average" | "min" | "max";

/** What the agent asked for. */
export interface SummaryRequest {
  sourceSheet: string;
  /** 1-based; found by the engine when left out. */
  headerRow: number | null;
  /** The field to group by: a header text or a column letter. */
  rows: string;
  /** A second field across the top (a crosstab), or null. */
  columns: string | null;
  values: Array<{ field: string; fn: SummaryFn }>;
  destSheet: string;
  /** Top-left cell of the table, 0-based. */
  dest: { row: number; col: number };
  title: string | null;
  /** Rows by label, or largest first by the first value. */
  sort: "label" | "value";
}

export type SummaryKey = string | number | boolean;

/** One field of the data as the engine found it. */
export interface SummaryField {
  /** 0-based column on the data sheet. */
  col: number;
  header: string;
  /** Distinct keys as the cells hold them (a date is its serial number), in output order. */
  keys: SummaryKey[];
  /** The number format of the field's cells, when it is not General. */
  format: string | null;
}

export interface SummaryFacts {
  /** The data sheet's name as the workbook spells it. */
  sheet: string;
  /** First and last data row, 1-based. */
  firstRow: number;
  lastRow: number;
  rowField: SummaryField;
  colField: SummaryField | null;
  measures: Array<{ col: number; header: string; fn: SummaryFn; format: string | null }>;
}

export interface SummaryReport {
  sheet: string;
  /** The whole table, e.g. "A1:F9". */
  block: string;
  header_row: number;
  first_data_row: number;
  last_data_row: number;
  total_row: number;
  /** The cell that must read 0, when the table has one. */
  check_cell: string | null;
  /** Label → 1-based row (the first 60; `row_count` says how many there are). */
  rows: Record<string, number>;
  row_count: number;
  /** Column heading → column letters. */
  cols: Record<string, string>;
  source: string;
  note: string;
}

const FN_LABEL: Record<SummaryFn, string> = { sum: "Sum", count: "Count", average: "Average", min: "Min", max: "Max" };

/** `'P&L 2026'!` or `Data!`: a sheet name in front of a range. */
export function sheetQualifier(name: string): string {
  const plain = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name);
  return `${plain ? name : `'${name.replace(/'/g, "''")}'`}!`;
}

/** A text key that SUMIFS would read as a pattern or a comparison ("A*",
 * ">10", "=x") is written into the formula as an exact-match criterion;
 * every other key is read from its label cell. */
function criterion(key: SummaryKey, cell: string): string {
  if (typeof key !== "string" || !/[*?~]|^[=<>]/.test(key)) return cell;
  return `"=${key.replace(/([*?~])/g, "~$1").replace(/"/g, '""')}"`;
}

/** One aggregate over the rows that meet `criteria` (pairs of range and criterion). */
function aggregate(fn: SummaryFn, values: string, criteria: Array<[string, string]>): string {
  const where = criteria.map(([range, crit]) => `${range},${crit}`).join(",");
  if (criteria.length === 0) {
    switch (fn) {
      case "sum": return `=SUM(${values})`;
      case "count": return `=COUNTA(${values})`;
      case "average": return `=IFERROR(AVERAGE(${values}),"")`;
      case "min": return `=MIN(${values})`;
      case "max": return `=MAX(${values})`;
    }
  }
  switch (fn) {
    case "sum": return `=SUMIFS(${values},${where})`;
    case "count": return `=COUNTIFS(${where},${values},"<>")`;
    case "average": return `=IFERROR(AVERAGEIFS(${values},${where}),"")`;
    // No rows for the group: empty, not a 0 that reads as a minimum.
    case "min": return `=IF(COUNTIFS(${where},${values},"<>")=0,"",MINIFS(${values},${where}))`;
    case "max": return `=IF(COUNTIFS(${where},${values},"<>")=0,"",MAXIFS(${values},${where}))`;
  }
}

export function layoutSummary(req: SummaryRequest, facts: SummaryFacts): { mutations: UniverMutation[]; report: SummaryReport } {
  const muts: any[] = [];
  const sheet = req.destSheet;
  const setCell = (row: number, col: number, value: SummaryKey) => {
    const isFormula = typeof value === "string" && value.startsWith("=");
    muts.push({
      type: "set_cell",
      address: { sheet, row, col },
      old_value: null,
      new_value: isFormula ? null : value,
      new_formula: isFormula ? value : null,
    });
  };
  const fmt = (range: string, format: Record<string, unknown>) => {
    muts.push({ type: "set_format", sheet, range, cells: expandA1Range(range), old_format: [], new_format: format });
  };
  const a1 = (row: number, col: number) => `${colLetters(col)}${row + 1}`;

  const q = sheetQualifier(facts.sheet);
  const column = (col: number) => `${q}$${colLetters(col)}$${facts.firstRow}:$${colLetters(col)}$${facts.lastRow}`;
  const rowRange = column(facts.rowField.col);
  const colRange = facts.colField ? column(facts.colField.col) : null;

  const c0 = req.dest.col;
  let r = req.dest.row;
  if (req.title) {
    setCell(r, c0, req.title);
    fmt(a1(r, c0), { bold: true });
    r += 1;
  }
  const headerRow = r;
  const firstData = headerRow + 1;
  const rowKeys = facts.rowField.keys;
  const lastData = firstData + rowKeys.length - 1;
  const totalRow = lastData + 1;
  const checkRow = totalRow + 1;
  const rowsMap: Record<string, number> = {};
  const colsMap: Record<string, string> = {};
  const rowCrit = (i: number) => criterion(rowKeys[i], `$${colLetters(c0)}${firstData + i + 1}`);

  setCell(headerRow, c0, facts.rowField.header);
  rowKeys.forEach((key, i) => {
    setCell(firstData + i, c0, key);
    if (i < 60) rowsMap[String(key)] = firstData + i + 1;
  });
  setCell(totalRow, c0, "Total");

  let lastCol = c0;
  let checkCell: string | null = null;
  // Sum and count add up: the groups must come to the total of the data.
  const addsUp = (fn: SummaryFn) => fn === "sum" || fn === "count";

  if (facts.colField && colRange) {
    // Crosstab: one measure, the second field across the top, totals at
    // the right and at the bottom.
    const m = facts.measures[0];
    const values = column(m.col);
    const colKeys = facts.colField.keys;
    const totalCol = c0 + 1 + colKeys.length;
    lastCol = totalCol;
    const colCrit = (j: number) => criterion(colKeys[j], `${colLetters(c0 + 1 + j)}$${headerRow + 1}`);
    colKeys.forEach((key, j) => {
      setCell(headerRow, c0 + 1 + j, key);
      colsMap[String(key)] = colLetters(c0 + 1 + j);
      rowKeys.forEach((_, i) => setCell(firstData + i, c0 + 1 + j, aggregate(m.fn, values, [[rowRange, rowCrit(i)], [colRange, colCrit(j)]])));
      setCell(totalRow, c0 + 1 + j, aggregate(m.fn, values, [[colRange, colCrit(j)]]));
    });
    setCell(headerRow, totalCol, "Total");
    colsMap["Total"] = colLetters(totalCol);
    rowKeys.forEach((_, i) => setCell(firstData + i, totalCol, aggregate(m.fn, values, [[rowRange, rowCrit(i)]])));
    setCell(totalRow, totalCol, aggregate(m.fn, values, []));
    if (addsUp(m.fn)) {
      setCell(checkRow, c0, "Check (0 = the table holds every row of the data)");
      setCell(checkRow, totalCol, `=SUM(${a1(firstData, c0 + 1)}:${a1(lastData, totalCol - 1)})-${a1(totalRow, totalCol)}`);
      checkCell = a1(checkRow, totalCol);
    }
    if (facts.colField.format) fmt(`${a1(headerRow, c0 + 1)}:${a1(headerRow, totalCol - 1)}`, { number_format: facts.colField.format });
    if (m.format && m.fn !== "count") fmt(`${a1(firstData, c0 + 1)}:${a1(totalRow, totalCol)}`, { number_format: m.format });
  } else {
    // One column per measure.
    facts.measures.forEach((m, k) => {
      const col = c0 + 1 + k;
      const values = column(m.col);
      const heading = `${FN_LABEL[m.fn]} of ${m.header}`;
      setCell(headerRow, col, heading);
      colsMap[heading] = colLetters(col);
      rowKeys.forEach((_, i) => setCell(firstData + i, col, aggregate(m.fn, values, [[rowRange, rowCrit(i)]])));
      setCell(totalRow, col, aggregate(m.fn, values, []));
      if (addsUp(m.fn)) {
        if (checkCell === null) setCell(checkRow, c0, "Check (0 = the table holds every row of the data)");
        setCell(checkRow, col, `=SUM(${a1(firstData, col)}:${a1(lastData, col)})-${a1(totalRow, col)}`);
        checkCell ??= a1(checkRow, col);
      }
      if (m.format && m.fn !== "count") fmt(`${a1(firstData, col)}:${a1(totalRow, col)}`, { number_format: m.format });
    });
    lastCol = c0 + facts.measures.length;
  }

  if (facts.rowField.format) fmt(`${a1(firstData, c0)}:${a1(lastData, c0)}`, { number_format: facts.rowField.format });
  fmt(`${a1(headerRow, c0)}:${a1(headerRow, lastCol)}`, { bold: true });
  fmt(`${a1(totalRow, c0)}:${a1(totalRow, lastCol)}`, { bold: true });
  if (checkCell) fmt(`${a1(checkRow, c0)}:${a1(checkRow, lastCol)}`, { italic: true, font_color: "#808080" });

  const lastRow = checkCell ? checkRow : totalRow;
  const source = `${q}${colLetters(Math.min(facts.rowField.col, ...facts.measures.map((m) => m.col)))}${facts.firstRow}:…${facts.lastRow}`;
  return {
    mutations: muts as UniverMutation[],
    report: {
      sheet,
      block: `${a1(req.dest.row, c0)}:${a1(lastRow, lastCol)}`,
      header_row: headerRow + 1,
      first_data_row: firstData + 1,
      last_data_row: lastData + 1,
      total_row: totalRow + 1,
      check_cell: checkCell,
      rows: rowsMap,
      row_count: rowKeys.length,
      cols: colsMap,
      source,
      note:
        `Every cell is a formula over rows ${facts.firstRow}-${facts.lastRow} of ${facts.sheet}: the table follows edits to that data. ` +
        `Rows added below row ${facts.lastRow}, or a new ${facts.rowField.header}${facts.colField ? ` or ${facts.colField.header}` : ""} that is not listed, are NOT picked up` +
        (checkCell ? `: ${checkCell} then stops reading 0, and the table has to be made again.` : ": make the table again then."),
    },
  };
}

/**
 * What the agent is told about the table once it is on the sheet: where it
 * is, and whether it accounts for every row of the data. `valueAt` reads a
 * computed cell (0-based).
 */
export function summaryResult(
  report: SummaryReport & { rows_without_a_key?: number },
  valueAt: (sheet: string, row: number, col: number) => unknown,
): Record<string, unknown> {
  if (!report.check_cell) return { ...report };
  const at = expandA1Range(report.check_cell)[0];
  const value = at ? valueAt(report.sheet, at.row, at.col) : null;
  const off = typeof value === "number" && Math.abs(value) > 1e-6;
  return {
    ...report,
    check_value: value ?? null,
    ...(off
      ? {
          check_warning:
            `The check cell ${report.check_cell} reads ${value}, not 0: the Total covers data the rows do not. ` +
            (report.rows_without_a_key
              ? `${report.rows_without_a_key} data row${report.rows_without_a_key === 1 ? " has" : "s have"} an empty key and are in no row of the table. `
              : "") +
            "If something other than records sits inside the data range (a totals row at the bottom is the usual cause), clear this table and make it again with `last_row`. Tell the user what the difference is; do not hide the check.",
        }
      : {}),
  };
}
