import { describe, expect, it } from "vitest";

import { __internals } from "../ironcalcEngine";

const { decodeStructural, decodeMutationCells, mutationCellToInput, classifyClaimedMutations } =
  __internals;

describe("decodeStructural", () => {
  it("maps insert-row mutation with a range to insert_rows (1-based)", () => {
    expect(
      decodeStructural("sheet.mutation.insert-row", {
        range: { startRow: 4, endRow: 6, startColumn: 0, endColumn: 16383 },
      }),
    ).toEqual({ op: "insert_rows", index1: 5, count: 3 });
  });

  it("maps remove-row and column mutations", () => {
    expect(
      decodeStructural("sheet.mutation.remove-row", {
        range: { startRow: 9, endRow: 9, startColumn: 0, endColumn: 16383 },
      }),
    ).toEqual({ op: "delete_rows", index1: 10, count: 1 });
    expect(
      decodeStructural("sheet.mutation.insert-col", {
        range: { startRow: 0, endRow: 1048575, startColumn: 2, endColumn: 3 },
      }),
    ).toEqual({ op: "insert_columns", index1: 3, count: 2 });
    expect(
      decodeStructural("sheet.mutation.remove-col", {
        range: { startRow: 0, endRow: 1048575, startColumn: 5, endColumn: 5 },
      }),
    ).toEqual({ op: "delete_columns", index1: 6, count: 1 });
  });

  it("maps the -by-range commands — the funnel every insert/delete goes through", () => {
    // The facade helpers (agent path) dispatch these DIRECTLY; the menu's
    // bare/-before/-after commands reach them via delegation. Matching only
    // the funnel mirrors each operation exactly once.
    expect(
      decodeStructural("sheet.command.insert-row-by-range", {
        range: { startRow: 4, endRow: 6, startColumn: 0, endColumn: 16383 },
      }),
    ).toEqual({ op: "insert_rows", index1: 5, count: 3 });
    expect(
      decodeStructural("sheet.command.remove-row-by-range", {
        range: { startRow: 9, endRow: 9, startColumn: 0, endColumn: 16383 },
      }),
    ).toEqual({ op: "delete_rows", index1: 10, count: 1 });
    expect(
      decodeStructural("sheet.command.insert-col-by-range", {
        range: { startRow: 0, endRow: 1048575, startColumn: 2, endColumn: 3 },
      }),
    ).toEqual({ op: "insert_columns", index1: 3, count: 2 });
    expect(
      decodeStructural("sheet.command.remove-col-by-range", {
        range: { startRow: 0, endRow: 1048575, startColumn: 5, endColumn: 5 },
      }),
    ).toEqual({ op: "delete_columns", index1: 6, count: 1 });
  });

  it("ignores the bare/-before/-after commands that fire AROUND the funnel", () => {
    // On the manual-menu path insert-row-by-range completes first, then the
    // bare insert-row, then insert-row-before — decoding any of the outer
    // commands too would apply the operation twice.
    expect(
      decodeStructural("sheet.command.insert-row", {
        range: { startRow: 4, endRow: 6 },
      }),
    ).toBeNull();
    expect(
      decodeStructural("sheet.command.insert-row-before", {
        range: { startRow: 4, endRow: 6 },
      }),
    ).toBeNull();
    expect(
      decodeStructural("univer.command.undo", {
        range: { startRow: 4, endRow: 6 },
      }),
    ).toBeNull();
  });

  it("returns null without a usable range", () => {
    expect(decodeStructural("sheet.mutation.insert-row", {})).toBeNull();
  });
});

describe("classifyClaimedMutations", () => {
  const namesById = new Map([["sub1", "Sheet1"]]);
  const sheets = ["Sheet1"];

  it("classifies style/empty children as none (set-style must not resync)", () => {
    expect(classifyClaimedMutations(sheets, [], namesById)).toEqual({ kind: "none" });
    expect(
      classifyClaimedMutations(
        sheets,
        [
          {
            id: "sheet.mutation.set-range-values",
            params: { subUnitId: "sub1", cellValue: { "0": { "0": { s: "styleId" } } } },
          },
        ],
        namesById,
      ),
    ).toEqual({ kind: "none" });
  });

  it("decodes plain writes into a mirrorable edit batch (drag-fill, simple paste)", () => {
    const verdict = classifyClaimedMutations(
      sheets,
      [
        {
          id: "sheet.mutation.set-range-values",
          params: {
            subUnitId: "sub1",
            cellValue: { "4": { "2": { v: 7 }, "3": { f: "=A1+1" } } },
          },
        },
      ],
      namesById,
    );
    expect(verdict).toEqual({
      kind: "edits",
      edits: [
        { sheet: 0, row: 5, column: 3, value: "7" },
        { sheet: 0, row: 5, column: 4, value: "=A1+1" },
      ],
    });
  });

  it("goes opaque on movers, structural children and si fills", () => {
    expect(
      classifyClaimedMutations(
        sheets,
        [{ id: "sheet.mutation.reorder-range", params: {} }],
        namesById,
      ),
    ).toEqual({ kind: "opaque", why: "sheet.mutation.reorder-range" });
    expect(
      classifyClaimedMutations(
        sheets,
        [
          {
            id: "sheet.mutation.insert-row",
            params: { range: { startRow: 1, endRow: 1 } },
          },
        ],
        namesById,
      ),
    ).toEqual({ kind: "opaque", why: "sheet.mutation.insert-row" });
    expect(
      classifyClaimedMutations(
        sheets,
        [
          {
            id: "sheet.mutation.set-range-values",
            params: { subUnitId: "sub1", cellValue: { "0": { "0": { si: "0" } } } },
          },
        ],
        namesById,
      ),
    ).toEqual({ kind: "opaque", why: "shared-formula si without text" });
  });

  it("goes opaque when the sheet cannot be resolved", () => {
    expect(
      classifyClaimedMutations(
        sheets,
        [
          {
            id: "sheet.mutation.set-range-values",
            params: { subUnitId: "unknown-sub", cellValue: { "0": { "0": { v: 1 } } } },
          },
        ],
        namesById,
      ),
    ).toEqual({ kind: "opaque", why: "unresolvable sheet" });
  });
});

describe("mutationCellToInput", () => {
  it("prefers formula text", () => {
    expect(mutationCellToInput({ f: "=SUM(A1:B2)", v: 42 })).toBe("=SUM(A1:B2)");
  });
  it("restores plain values (undo payloads carry v)", () => {
    expect(mutationCellToInput({ v: 42 })).toBe("42");
    expect(mutationCellToInput({ v: "hello" })).toBe("hello");
    expect(mutationCellToInput({ v: true })).toBe("TRUE");
  });
  it("clears on null cell or explicit null v", () => {
    expect(mutationCellToInput(null)).toBe("");
    expect(mutationCellToInput({ v: null, f: null })).toBe("");
  });
  it("skips style-only and rich-text-only patches", () => {
    expect(mutationCellToInput({ s: "xyz" })).toBeNull();
    expect(mutationCellToInput({ p: { body: {} } })).toBeNull();
  });
  it("flags unreconstructable shared-formula refs", () => {
    expect(mutationCellToInput({ si: "0" })).toBe("si");
  });
});

describe("decodeMutationCells", () => {
  it("decodes a cellValue matrix into 0-based edits", () => {
    const res = decodeMutationCells({
      cellValue: { "4": { "2": { v: 123 }, "3": { f: "=A1" } } },
    });
    expect(res.fail).toBeNull();
    expect(res.cells).toEqual([
      { row0: 4, col0: 2, input: "123" },
      { row0: 4, col0: 3, input: "=A1" },
    ]);
  });

  it("skips style-only cells but keeps content cells", () => {
    const res = decodeMutationCells({
      cellValue: { "0": { "0": { s: "abc" }, "1": { v: 7 } } },
    });
    expect(res.fail).toBeNull();
    expect(res.cells).toEqual([{ row0: 0, col0: 1, input: "7" }]);
  });

  it("fails the whole mutation on si refs (degrade beats drift)", () => {
    const res = decodeMutationCells({
      cellValue: { "0": { "0": { si: "3" } } },
    });
    expect(res.fail).toMatch(/si/);
    expect(res.cells).toEqual([]);
  });

  it("treats an empty payload as a no-op", () => {
    expect(decodeMutationCells({})).toEqual({ cells: [], fail: null });
  });
});
