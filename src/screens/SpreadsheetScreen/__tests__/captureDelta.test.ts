import { describe, expect, it } from "vitest";
import { captureWorkbookContext } from "../agent/captureContext";

/**
 * Base+delta capture: the preview must stay byte-identical (prompt-cache
 * hit) across content edits, with changed cells shipped in `delta` instead,
 * until the delta is structural or over budget — then a fresh capture.
 *
 * The memo is module-global state keyed by path, so every test uses a
 * fresh path.
 */

let pathCounter = 0;
const freshPath = () => `/tmp/capture-delta-test-${pathCounter++}.xlsx`;

/** Minimal grid handle: capture only needs the snapshot for these tests. */
const gridOf = (snapshot: any) => ({ getWorkbookSnapshot: () => snapshot }) as any;

/** Build a snapshot from { SheetName: { r: { c: {v?, f?} } } }. */
function snap(cells: Record<string, Record<number, Record<number, any>>>): any {
  const sheetOrder: string[] = [];
  const sheets: any = {};
  let i = 0;
  for (const [name, cellData] of Object.entries(cells)) {
    const id = `s${i++}`;
    sheetOrder.push(id);
    sheets[id] = { name, cellData, rowCount: 100, columnCount: 26 };
  }
  return { sheetOrder, sheets };
}

describe("base+delta capture", () => {
  it("returns the memoized preview object with no delta while content is unchanged", () => {
    const path = freshPath();
    const cells = { Model: { 0: { 0: { v: "Revenue" }, 1: { v: 100 } } } };
    const first = captureWorkbookContext(path, gridOf(snap(cells)));
    // Fresh snapshot object, identical content — must hit the memo.
    const second = captureWorkbookContext(path, gridOf(snap(cells)));
    expect(second.sheets).toBe(first.sheets);
    expect(second.delta).toBeUndefined();
  });

  it("keeps the base preview bytes and ships an edit in delta", () => {
    const path = freshPath();
    const first = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "Revenue" }, 1: { v: 100 } } } })),
    );
    const second = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "Revenue" }, 1: { v: 42 } } } })),
    );
    // The cached block must be the exact same object (byte-identical).
    expect(second.sheets).toBe(first.sheets);
    expect(second.delta).toContain("# Workbook changes since the preview snapshot");
    expect(second.delta).toContain("## Sheet: Model");
    expect(second.delta).toContain("B1 = 42");
    // Deterministic while content stays put: a third capture ships the same delta.
    const third = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "Revenue" }, 1: { v: 42 } } } })),
    );
    expect(third.sheets).toBe(first.sheets);
    expect(third.delta).toBe(second.delta);
  });

  it("collapses a recalculated row (formulas unchanged, values moved) to one segment", () => {
    const path = freshPath();
    const row = (vs: number[]) => {
      const r: Record<number, any> = {};
      // B2:K2 = translated formulas =B1*2, =C1*2, … with cached values
      const letters = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];
      vs.forEach((v, i) => {
        r[i + 1] = { f: `=${letters[i]}1*2`, v };
      });
      return r;
    };
    captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 1: { v: 1 } }, 1: row([2, 4, 6, 8, 10, 12, 14, 16, 18, 20]) } })),
    );
    const after = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 1: { v: 2 } }, 1: row([3, 5, 7, 9, 11, 13, 15, 17, 19, 21]) } })),
    );
    expect(after.delta).toBeDefined();
    const recalcLines = after.delta!.split("\n").filter((l) => l.includes("recalc"));
    // One collapsed segment for B2:K2 — not ten lines. (The header mentions
    // "recalc" once as notation documentation.)
    const segment = recalcLines.filter((l) => l.startsWith("B2:K2"));
    expect(segment).toHaveLength(1);
    expect(segment[0]).toContain("B2=3");
    expect(segment[0]).toContain("K2=21");
    // The direct edit itself renders as an edit line, not a recalc.
    expect(after.delta).toContain("B1 = 2");
  });

  it("compresses an agent fill-right of translated formulas into one run line", () => {
    const path = freshPath();
    captureWorkbookContext(path, gridOf(snap({ Model: { 0: { 0: { v: "seed" } } } })));
    const filled: Record<number, any> = {};
    const letters = ["B", "C", "D", "E", "F", "G", "H", "I"];
    letters.forEach((col, i) => {
      filled[i + 1] = { f: `=${col}1*2`, v: (i + 1) * 2 };
    });
    const after = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "seed" } }, 2: filled } })),
    );
    expect(after.delta).toContain("(filled right ×8)");
    expect(after.delta).not.toContain("C3 = =C1*2");
  });

  it("lists cleared cells", () => {
    const path = freshPath();
    captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "Revenue" }, 1: { v: 100 } } } })),
    );
    const after = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "Revenue" } } } })),
    );
    expect(after.delta).toContain("B1 = (cleared)");
  });

  it("serves the memoized base with no delta when only formula values drift (recalc jitter)", () => {
    const path = freshPath();
    // Base captured while the engine was still settling: formula cells with
    // stale/absent computed values.
    const first = captureWorkbookContext(
      path,
      gridOf(
        snap({ Model: { 0: { 0: { v: 10 }, 1: { f: "=A1*2" }, 2: { f: "=B1+1" } } } }),
      ),
    );
    // Recalc finished: same formulas, computed values now present. No
    // authored content changed, so this must be a pure memo hit — same
    // preview object, no delta, no re-snapshot (the 181K "thank you" bug).
    const second = captureWorkbookContext(
      path,
      gridOf(
        snap({ Model: { 0: { 0: { v: 10 }, 1: { f: "=A1*2", v: 20 }, 2: { f: "=B1+1", v: 21 } } } }),
      ),
    );
    expect(second.sheets).toBe(first.sheets);
    expect(second.delta).toBeUndefined();
  });

  it("re-snapshots on sheet-set changes instead of shipping a delta", () => {
    const path = freshPath();
    const first = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "Revenue" } } } })),
    );
    const after = captureWorkbookContext(
      path,
      gridOf(
        snap({
          Model: { 0: { 0: { v: "Revenue" } } },
          Assumptions: { 0: { 0: { v: "Growth" } } },
        }),
      ),
    );
    expect(after.delta).toBeUndefined();
    expect(after.sheets).not.toBe(first.sheets);
    expect(after.sheets.map((s) => s.name)).toEqual(["Model", "Assumptions"]);
  });

  it("re-snapshots when the delta exceeds its budget", () => {
    const path = freshPath();
    const first = captureWorkbookContext(
      path,
      gridOf(snap({ Model: { 0: { 0: { v: "tiny" } } } })),
    );
    // Tiny base preview → budget floors at DELTA_MIN_BUDGET_BYTES (2KB).
    // 80 long distinct strings on non-adjacent rows (nothing compresses)
    // blow well past it, forcing a re-snapshot.
    const big: Record<number, Record<number, any>> = { 0: { 0: { v: "tiny" } } };
    for (let i = 0; i < 80; i++) {
      big[2 * i + 2] = { 0: { v: `a long edited text value number ${i} with some padding here` } };
    }
    const after = captureWorkbookContext(path, gridOf(snap({ Model: big })));
    expect(after.delta).toBeUndefined();
    expect(after.sheets).not.toBe(first.sheets);
    expect(after.sheets[0].cells_preview).toContain("a long edited text value number 0");
    // The re-snapshot is the new base: same content memoizes again.
    const again = captureWorkbookContext(path, gridOf(snap({ Model: big })));
    expect(again.sheets).toBe(after.sheets);
    expect(again.delta).toBeUndefined();
  });
});

describe("hidden sheets", () => {
  const withHidden = (s: any, names: string[]) => {
    for (const id of s.sheetOrder) {
      if (names.includes(s.sheets[id].name)) s.sheets[id].hidden = 1;
    }
    return s;
  };

  it("collapses hidden sheets to a marker + used range + row labels", () => {
    const path = freshPath();
    const ctx = captureWorkbookContext(
      path,
      gridOf(
        withHidden(
          snap({
            Model: { 0: { 0: { v: "Revenue" }, 1: { v: 100 } } },
            Calc: {
              0: { 0: { v: "WACC" }, 1: { v: 0.09 } },
              4: { 0: { v: "Terminal growth" }, 3: { v: 0.02 } },
            },
          }),
          ["Calc"],
        ),
      ),
    );
    const calc = ctx.sheets.find((s) => s.name === "Calc")!;
    expect(calc.cells_preview).toContain("HIDDEN");
    expect(calc.cells_preview).toContain('row 1: "WACC"');
    expect(calc.cells_preview).toContain('row 5: "Terminal growth"');
    // Cell values stay out of the collapsed preview.
    expect(calc.cells_preview).not.toContain("0.09");
    expect(calc.used_range).toBe("A1:D5");
    // The visible sheet still gets its full preview.
    expect(ctx.sheets.find((s) => s.name === "Model")!.cells_preview).toContain("B1 = 100");
  });

  it("re-snapshots when a sheet's visibility flips with content unchanged", () => {
    const path = freshPath();
    const cells = {
      Model: { 0: { 0: { v: "Revenue" } } },
      Calc: { 0: { 0: { v: "WACC" }, 1: { v: 0.09 } } },
    };
    const first = captureWorkbookContext(path, gridOf(withHidden(snap(cells), ["Calc"])));
    expect(first.sheets.find((s) => s.name === "Calc")!.cells_preview).toContain("HIDDEN");
    const second = captureWorkbookContext(path, gridOf(snap(cells)));
    expect(second.delta).toBeUndefined();
    expect(second.sheets).not.toBe(first.sheets);
    expect(second.sheets.find((s) => s.name === "Calc")!.cells_preview).toContain("B1 = 0.09");
  });

  it("still ships hidden-sheet edits in the delta", () => {
    const path = freshPath();
    const base = {
      Model: { 0: { 0: { v: "Revenue" } } },
      Calc: { 0: { 0: { v: "WACC" }, 1: { v: 0.09 } } },
    };
    const first = captureWorkbookContext(path, gridOf(withHidden(snap(base), ["Calc"])));
    const edited = {
      Model: { 0: { 0: { v: "Revenue" } } },
      Calc: { 0: { 0: { v: "WACC" }, 1: { v: 0.095 } } },
    };
    const after = captureWorkbookContext(path, gridOf(withHidden(snap(edited), ["Calc"])));
    expect(after.sheets).toBe(first.sheets);
    expect(after.delta).toContain("## Sheet: Calc");
    expect(after.delta).toContain("B1 = 0.095");
  });
});
