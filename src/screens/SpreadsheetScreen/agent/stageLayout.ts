import type { UniverMutation } from "../types";
import { expandA1Range } from "./toolToMutation";

/**
 * stage_data layout — the sheet-as-calculator staging affordance.
 *
 * The model hands over REPORTED figures the moment it extracts them from a
 * source; the harness (not the model) picks the placement on the staging
 * sheet, writes title / provenance / header / data rows through the normal
 * mutation pipeline, and returns an address map keyed by the labels the
 * model just supplied — so `=Data!C15` costs zero mental bookkeeping and
 * "derive by formula, don't retype" becomes the lazy path.
 *
 * Pure function: (payload, startRow) → mutations + report. All placement
 * logic lives here so it is unit-testable without a grid.
 */

export interface StagePayload {
  sheet: string;
  title: string;
  source: string;
  units: string | null;
  columns: string[];
  rows: Array<{ label: string; values: unknown[] }>;
}

export interface StageReport {
  sheet: string;
  block: string;
  header_row: number;
  cols: Record<string, string>;
  rows: Record<string, number>;
  hint: string;
}

/** 0-indexed column → A1 letters. */
export function colLetter(n: number): string {
  let s = "";
  let x = n + 1;
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

export function layoutStageBlock(
  p: StagePayload,
  startRow: number,
): { mutations: UniverMutation[]; report: StageReport; endRow: number } {
  const muts: any[] = [];
  const setCell = (row: number, col: number, value: unknown) => {
    muts.push({
      type: "set_cell",
      address: { sheet: p.sheet, row, col },
      old_value: null,
      new_value: value,
      new_formula: null,
    });
  };
  const fmt = (range: string, format: Record<string, unknown>) => {
    muts.push({
      type: "set_format",
      sheet: p.sheet,
      range,
      cells: expandA1Range(range),
      old_format: [],
      new_format: format,
    });
  };

  const r0 = startRow;
  setCell(r0, 0, p.title);
  setCell(r0 + 1, 0, `Source: ${p.source}${p.units ? `  ·  units: ${p.units}` : ""}`);

  const headerRow = r0 + 2;
  for (let c = 0; c < p.columns.length; c++) setCell(headerRow, c + 1, p.columns[c]);

  const firstDataRow = headerRow + 1;
  const rowsMap: Record<string, number> = {};
  p.rows.forEach((row, i) => {
    const r = firstDataRow + i;
    setCell(r, 0, row.label);
    // Last write wins on duplicate labels — the map points at the final one.
    rowsMap[row.label] = r + 1; // 1-indexed A1 row
    row.values.forEach((v, c) => {
      // Same preserve semantics as set_range: blanks are gaps, not clears.
      if (v === undefined || v === null || v === "") return;
      setCell(r, c + 1, v);
    });
  });

  const endRow = firstDataRow + p.rows.length - 1;
  const lastCol = p.columns.length; // 0-indexed index of the last value column

  fmt(`A${r0 + 1}`, { bold: true });
  fmt(`A${r0 + 2}`, { font_color: "#808080" });
  fmt(`A${headerRow + 1}:${colLetter(lastCol)}${headerRow + 1}`, { bold: true });

  const colsMap: Record<string, string> = {};
  p.columns.forEach((name, c) => {
    colsMap[name] = colLetter(c + 1);
  });

  const report: StageReport = {
    sheet: p.sheet,
    block: `A${r0 + 1}:${colLetter(lastCol)}${endRow + 1}`,
    header_row: headerRow + 1,
    cols: colsMap,
    rows: rowsMap,
    hint:
      `Staged. Reference these cells by FORMULA (e.g. ='${p.sheet}'!${colsMap[p.columns[0]] ?? "B"}${
        Object.values(rowsMap)[0] ?? firstDataRow + 1
      }) — derive, don't retype.`,
  };
  return { mutations: muts as UniverMutation[], report, endRow };
}
