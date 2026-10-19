import type { UniverGridHandle } from "../components/UniverGrid";
import { parseA1, expandA1Range } from "./toolToMutation";

/**
 * Build a "User focus" block to ship to the agent — combines the live grid
 * selection (if any) with any @-mentioned ranges from the prompt. Both
 * resolve to actual cell values via the live UniverGrid, so the agent sees
 * what the user actually has in those cells (formulas + evaluated values),
 * not a stale snapshot.
 *
 * Returns null when there's nothing to focus on (no selection, no mentions).
 */
export function buildFocusContext(
  prompt: string,
  grid: UniverGridHandle | null,
): { text: string; selectionLabel: string | null; mentionLabels: string[] } | null {
  if (!grid) return null;

  const selection = grid.getActiveSelection?.() ?? null;
  const mentions = extractMentions(prompt);

  // Skip "selection" only when it's a single cell — single-cell focus is
  // usually just where the cursor parked, not a deliberate target. Any
  // 1×N or N×N range counts as user focus. We DO still inline its values
  // in mention form if the user explicitly @-mentions a single cell.
  const meaningfulSelection =
    selection && (selection.endRow !== selection.startRow || selection.endCol !== selection.startCol)
      ? selection
      : null;
  if (selection && !meaningfulSelection) {
    console.log("[focus] single-cell selection — not shipped as focus block:", selection);
  } else if (meaningfulSelection) {
    console.log("[focus] shipping selection to agent:", meaningfulSelection);
  } else {
    console.log("[focus] no selection captured");
  }

  if (!meaningfulSelection && mentions.length === 0) return null;

  const blocks: string[] = [];
  const mentionLabels: string[] = [];
  let selectionLabel: string | null = null;

  if (meaningfulSelection) {
    const rangeLabel = `${a1(meaningfulSelection.startRow, meaningfulSelection.startCol)}:${a1(
      meaningfulSelection.endRow,
      meaningfulSelection.endCol,
    )}`;
    const lines = readCells(grid, meaningfulSelection.sheet, [
      {
        startRow: meaningfulSelection.startRow,
        startCol: meaningfulSelection.startCol,
        endRow: meaningfulSelection.endRow,
        endCol: meaningfulSelection.endCol,
      },
    ]);
    const cellCount =
      (meaningfulSelection.endRow - meaningfulSelection.startRow + 1) *
      (meaningfulSelection.endCol - meaningfulSelection.startCol + 1);
    selectionLabel = `${meaningfulSelection.sheet}!${rangeLabel} (${cellCount} cell${cellCount === 1 ? "" : "s"})`;
    blocks.push(
      `## Selection\nThe user has selected ${selectionLabel}. Treat this as the primary target for edits unless the prompt clearly says otherwise.\n${lines || "(empty cells)"}`,
    );
  }

  for (const m of mentions) {
    const cells = expandA1Range(m.range);
    if (cells.length === 0) continue;
    const lines = readCellsFromList(grid, m.sheet, cells);
    const label = `${m.sheet}!${m.range}`;
    mentionLabels.push(label);
    blocks.push(`## @-mention ${m.raw}\n${label}\n${lines || "(empty cells)"}`);
  }

  if (blocks.length === 0) return null;

  return {
    text: `# User focus (cells the user is directing the agent at)\n${blocks.join("\n\n")}\n`,
    selectionLabel,
    mentionLabels,
  };
}

/**
 * Parse `@A1`, `@A1:C10`, `@Sheet2!A1:C10` patterns out of the prompt text.
 * Returns the resolved sheet + range for each. When sheet is omitted, the
 * caller will resolve it against the active sheet later.
 */
export function extractMentions(prompt: string): Array<{
  raw: string;
  sheet: string;
  range: string;
}> {
  const out: Array<{ raw: string; sheet: string; range: string }> = [];
  // Allow @Sheet!A1:B2 or @A1:B2 — sheet name is up to first ! delimiter
  // or absent. Range must be alpha+digit, optionally with :alpha+digit.
  const re = /@(?:([A-Za-z0-9_ ]+)!)?([A-Za-z]+\d+(?::[A-Za-z]+\d+)?)/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(prompt)) !== null) {
    const sheet = (match[1] ?? "").trim();
    const range = match[2];
    out.push({ raw: match[0], sheet, range });
  }
  return out;
}

function readCells(
  grid: UniverGridHandle,
  sheet: string,
  rects: Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>,
): string {
  const lines: string[] = [];
  const cap = 200; // hard cap on cells inlined per focus block
  outer: for (const rect of rects) {
    for (let r = rect.startRow; r <= rect.endRow; r++) {
      for (let c = rect.startCol; c <= rect.endCol; c++) {
        if (lines.length >= cap) {
          lines.push("... (truncated)");
          break outer;
        }
        const cell = grid.getCell?.(sheet, r, c);
        if (!cell) continue;
        const addr = a1(r, c);
        if (cell.formula) {
          lines.push(`${addr} = ${cell.formula}  → ${JSON.stringify(cell.value)}`);
        } else if (cell.value !== null && cell.value !== undefined) {
          lines.push(`${addr} = ${JSON.stringify(cell.value)}`);
        }
      }
    }
  }
  return lines.join("\n");
}

function readCellsFromList(
  grid: UniverGridHandle,
  sheet: string,
  cells: Array<{ row: number; col: number }>,
): string {
  const lines: string[] = [];
  const cap = 200;
  for (const { row, col } of cells) {
    if (lines.length >= cap) {
      lines.push("... (truncated)");
      break;
    }
    const cell = grid.getCell?.(sheet, row, col);
    if (!cell) continue;
    const addr = a1(row, col);
    if (cell.formula) {
      lines.push(`${addr} = ${cell.formula}  → ${JSON.stringify(cell.value)}`);
    } else if (cell.value !== null && cell.value !== undefined) {
      lines.push(`${addr} = ${JSON.stringify(cell.value)}`);
    }
  }
  return lines.join("\n");
}

function a1(row: number, col: number): string {
  let n = col;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return `${s}${row + 1}`;
}

// Resolve mentions to the active sheet name when @-mention omits a sheet.
export function resolveMentionSheets<T extends { sheet: string }>(
  mentions: T[],
  defaultSheet: string,
): T[] {
  return mentions.map((m) => (m.sheet ? m : { ...m, sheet: defaultSheet }));
}
