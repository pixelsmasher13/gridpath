/**
 * Turn a sparse cell list into the smallest set of rectangles that covers
 * exactly those cells: per-row runs of consecutive columns, merged with the
 * run directly below when it spans the same columns. A full rectangle comes
 * back as itself, so range-level formats still cost one call.
 */
export function cellRuns(cells: Array<{ row: number; col: number }>): Array<{ r0: number; c0: number; r1: number; c1: number }> {
  const byRow = new Map<number, number[]>();
  for (const { row, col } of cells) {
    let cols = byRow.get(row);
    if (!cols) byRow.set(row, (cols = []));
    cols.push(col);
  }
  const rows = [...byRow.keys()].sort((a, b) => a - b);
  const out: Array<{ r0: number; c0: number; r1: number; c1: number }> = [];
  // Open rectangles from the previous row, keyed by "c0,c1".
  let open = new Map<string, { r0: number; c0: number; r1: number; c1: number }>();
  for (const row of rows) {
    const cols = [...new Set(byRow.get(row)!)].sort((a, b) => a - b);
    const runs: Array<[number, number]> = [];
    for (const c of cols) {
      const last = runs[runs.length - 1];
      if (last && c === last[1] + 1) last[1] = c;
      else runs.push([c, c]);
    }
    const next = new Map<string, { r0: number; c0: number; r1: number; c1: number }>();
    for (const [c0, c1] of runs) {
      const key = `${c0},${c1}`;
      const prev = open.get(key);
      if (prev && prev.r1 === row - 1) {
        prev.r1 = row;
        next.set(key, prev);
        open.delete(key);
      } else {
        const rect = { r0: row, c0, r1: row, c1 };
        out.push(rect);
        next.set(key, rect);
      }
    }
    open = next;
  }
  return out;
}
