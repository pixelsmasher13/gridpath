import type { CellFormat, FormatMutation, UniverMutation } from "../types";
import { parseA1 } from "./toolToMutation";
import { shiftFormulaA1 } from "../surgicalPatch";

/**
 * Expansion of the agent's `copy_range` tool into ordinary per-cell
 * mutations (set_cell + set_format), with Excel copy/paste semantics:
 *
 *   - relative formula references shift by the paste offset, `$` anchors
 *     don't (reuses shiftFormulaA1 — the same translation the surgical
 *     save applies to shared-formula members);
 *   - cell formats travel with the cells;
 *   - blank source cells leave the destination untouched (consistent with
 *     set_range's null-means-preserve semantics — NOT Excel's clear).
 *
 * Emitting plain mutations (rather than a dedicated mutation type) means
 * everything downstream works unchanged: diff tinting, Accept/Reject,
 * undo, layout validation on `done`, and the save mirror's format carry.
 *
 * All source cells are read BEFORE any mutation is applied, so an
 * overlapping paste (copy A1:A10 to A5) behaves like Excel: it copies the
 * pre-paste snapshot, never its own output.
 */

export const COPY_CELL_LIMIT = 5000;

/** Minimal slice of UniverGridHandle the expansion needs — keeps tests free of Univer. */
export type CopyGridReader = {
  getCell: (
    sheet: string,
    row: number,
    col: number,
  ) => { value: any; formula: string | null } | null;
  getCellFormat: (sheet: string, row: number, col: number) => CellFormat;
};

export type CopySpec = {
  sheet: string;
  source: string;
  dest_sheet: string;
  dest: string;
  mode: "all" | "values" | "formats";
};

export type CopyExpansion =
  | {
      ok: true;
      mutations: UniverMutation[];
      dest_range: string;
      copied_cells: number;
    }
  | { ok: false; error: string; message: string; ref_errors?: string[] };

function colLetters(col: number): string {
  let out = "";
  let c = col + 1;
  while (c > 0) {
    out = String.fromCharCode(65 + ((c - 1) % 26)) + out;
    c = Math.floor((c - 1) / 26);
  }
  return out;
}

function a1(row: number, col: number): string {
  return `${colLetters(col)}${row + 1}`;
}

const MAX_ROW = 1048576;
const MAX_COL = 16384;

export function expandCopy(grid: CopyGridReader, spec: CopySpec): CopyExpansion {
  const corners = spec.source.split(":");
  const a = parseA1(corners[0] ?? "");
  const b = corners.length === 2 ? parseA1(corners[1]) : a;
  const d = parseA1(spec.dest);
  if (!a || !b || !d) {
    return { ok: false, error: "bad_args", message: `copy_range: could not parse "${spec.source}" → "${spec.dest}"` };
  }
  const r0 = Math.min(a.row, b.row);
  const r1 = Math.max(a.row, b.row);
  const c0 = Math.min(a.col, b.col);
  const c1 = Math.max(a.col, b.col);
  const rows = r1 - r0 + 1;
  const cols = c1 - c0 + 1;
  if (rows * cols > COPY_CELL_LIMIT) {
    return {
      ok: false,
      error: "range_too_large",
      message: `copy_range: source is ${rows}×${cols} = ${rows * cols} cells; the limit is ${COPY_CELL_LIMIT}. Split into smaller copies.`,
    };
  }
  const dr = d.row - r0;
  const dc = d.col - c0;
  if (dr === 0 && dc === 0 && spec.dest_sheet === spec.sheet) {
    return { ok: false, error: "noop", message: "copy_range: destination equals source — nothing to do." };
  }
  if (d.row + rows - 1 >= MAX_ROW || d.col + cols - 1 >= MAX_COL) {
    return { ok: false, error: "out_of_bounds", message: "copy_range: destination rectangle extends past the sheet's limits." };
  }

  const cellMutations: UniverMutation[] = [];
  // Distinct format → list of destination cells, so a block sharing one
  // style lands as ONE FormatMutation (one diff row, one Reject snapshot
  // group) instead of hundreds.
  const formatGroups = new Map<string, { format: CellFormat; cells: Array<{ row: number; col: number }> }>();
  const refErrors: string[] = [];

  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const destRow = r + dr;
      const destCol = c + dc;
      const cell = grid.getCell(spec.sheet, r, c);

      if (spec.mode !== "formats" && cell && (cell.value !== null || cell.formula !== null)) {
        let newValue: any = null;
        let newFormula: string | null = null;
        if (spec.mode === "all" && cell.formula) {
          const shifted = shiftFormulaA1(cell.formula, dr, dc);
          if (shifted === null) {
            refErrors.push(`${a1(r, c)} → ${a1(destRow, destCol)}: "${cell.formula}"`);
            continue;
          }
          newFormula = shifted.startsWith("=") ? shifted : `=${shifted}`;
          newValue = newFormula;
        } else {
          // mode "values", or a plain literal in mode "all" — paste the
          // evaluated value as-is.
          newValue = cell.value;
        }
        if (newValue !== null || newFormula !== null) {
          cellMutations.push({
            type: "set_cell",
            address: { sheet: spec.dest_sheet, row: destRow, col: destCol },
            old_value: null,
            new_value: newValue,
            new_formula: newFormula,
          });
        }
      }

      if (spec.mode !== "values") {
        const fmt = grid.getCellFormat(spec.sheet, r, c);
        const keys = Object.keys(fmt ?? {});
        if (keys.length > 0) {
          const key = JSON.stringify(
            keys.sort().map((k) => [k, (fmt as any)[k]]),
          );
          const group = formatGroups.get(key) ?? { format: fmt, cells: [] };
          group.cells.push({ row: destRow, col: destCol });
          formatGroups.set(key, group);
        }
      }
    }
  }

  if (refErrors.length > 0) {
    return {
      ok: false,
      error: "ref_out_of_bounds",
      message:
        `copy_range: ${refErrors.length} formula(s) would shift references past the sheet edge ` +
        `(Excel would show #REF!). Nothing was copied. Fix the offset or copy without those cells.`,
      ref_errors: refErrors.slice(0, 20),
    };
  }

  const destRange =
    rows === 1 && cols === 1
      ? a1(d.row, d.col)
      : `${a1(d.row, d.col)}:${a1(d.row + rows - 1, d.col + cols - 1)}`;

  const mutations: UniverMutation[] = [...cellMutations];
  for (const { format, cells } of formatGroups.values()) {
    const fm: FormatMutation = {
      type: "set_format",
      sheet: spec.dest_sheet,
      range: destRange,
      cells,
      old_format: [],
      new_format: format,
    };
    mutations.push(fm);
  }

  if (mutations.length === 0) {
    return {
      ok: false,
      error: "empty_source",
      message: "copy_range: every source cell is empty and unformatted — nothing to copy.",
    };
  }

  return { ok: true, mutations, dest_range: destRange, copied_cells: rows * cols };
}
