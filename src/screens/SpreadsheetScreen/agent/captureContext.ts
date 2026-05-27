import type { UniverGridHandle } from "../components/UniverGrid";
import type { WorkbookContext } from "./agentClient";

/**
 * Capture a compact snapshot of the active workbook to ship to the agent.
 *
 * v1 strategy: for every sheet, emit at most `MAX_CELLS_PER_SHEET` cells in
 * "A1 = value" form. We send the used-range only, not the whole grid — most
 * xlsx files have a tiny used range vs. their nominal row/col counts.
 *
 * This is deliberately simple. When v2 lands a `read_range` tool the agent
 * can pull more cells on demand and we can shrink the up-front snapshot.
 */
/**
 * Cells per sheet sent in the workbook context block. We iterate the used
 * range (Univer's cellData only contains non-empty cells), so this is the
 * upper bound on cells shipped to the agent per sheet. 400 was too low for
 * typical financial models (50 rows × 12 cols already > 400). 1500 fits a
 * substantial multi-section model comfortably while staying well under
 * Claude's context budget across multiple sheets.
 */
const MAX_CELLS_PER_SHEET = 1500;

export function captureWorkbookContext(
  path: string,
  grid: UniverGridHandle | null,
): WorkbookContext {
  const snapshot = grid?.getWorkbookSnapshot?.();
  if (!snapshot || !snapshot.sheets) {
    return { path, sheets: [] };
  }
  const order: string[] = snapshot.sheetOrder ?? Object.keys(snapshot.sheets);
  const sheets = order
    .map((sheetId) => {
      const sheet = snapshot.sheets[sheetId];
      if (!sheet) return null;
      const cellData = sheet.cellData ?? {};
      const rowKeys = Object.keys(cellData)
        .map(Number)
        .sort((a, b) => a - b);
      const lines: string[] = [];
      outer: for (const r of rowKeys) {
        const row = cellData[r];
        if (!row) continue;
        const cols = Object.keys(row)
          .map(Number)
          .sort((a, b) => a - b);
        for (const c of cols) {
          if (lines.length >= MAX_CELLS_PER_SHEET) {
            lines.push("... (snapshot truncated)");
            break outer;
          }
          const cell = row[c];
          const addr = colLetters(c) + (r + 1);
          // For formula cells, ship BOTH the formula AND its evaluated
          // value (when available). Without this the agent can never see
          // #VALUE! errors or what its own formulas computed to. The
          // evaluated value comes from the live grid since the snapshot's
          // .v field is the most recent computed value Univer cached.
          if (cell.f) {
            const evaluated = grid?.getCell?.(sheet.name, r, c)?.value;
            if (evaluated !== undefined && evaluated !== null) {
              lines.push(`${addr} = ${cell.f}  → ${JSON.stringify(evaluated)}`);
            } else if (cell.v !== undefined) {
              lines.push(`${addr} = ${cell.f}  → ${JSON.stringify(cell.v)}`);
            } else {
              lines.push(`${addr} = ${cell.f}`);
            }
          } else if (cell.v !== undefined) {
            lines.push(`${addr} = ${JSON.stringify(cell.v)}`);
          }
        }
      }
      return {
        name: sheet.name as string,
        row_count: Number(sheet.rowCount ?? 100),
        column_count: Number(sheet.columnCount ?? 26),
        cells_preview: lines.length > 0 ? lines.join("\n") : "(empty)",
      };
    })
    .filter((s): s is NonNullable<typeof s> => s !== null);
  return { path, sheets };
}

function colLetters(col: number): string {
  let n = col;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}
