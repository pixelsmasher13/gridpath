import { describe, expect, it } from "vitest";
import { appendManualSheetOp, type SheetOp } from "../manualSheetOps";

function run(ops: SheetOp[]): SheetOp[] {
  const store: SheetOp[] = [];
  for (const op of ops) appendManualSheetOp(store, op);
  return store;
}

describe("appendManualSheetOp", () => {
  it("records create, delete, and rename as-is", () => {
    expect(run([{ kind: "create", name: "Sheet1" }])).toEqual([{ kind: "create", name: "Sheet1" }]);
    expect(run([{ kind: "delete", name: "Model" }])).toEqual([{ kind: "delete", name: "Model" }]);
    expect(run([{ kind: "rename", oldName: "Model", newName: "Model v2" }])).toEqual([
      { kind: "rename", oldName: "Model", newName: "Model v2" },
    ]);
  });

  it("cancels create+delete of the same session sheet (add then undo)", () => {
    expect(
      run([
        { kind: "create", name: "Sheet1" },
        { kind: "delete", name: "Sheet1" },
      ]),
    ).toEqual([]);
  });

  it("keeps unrelated ops when cancelling a create", () => {
    expect(
      run([
        { kind: "create", name: "A" },
        { kind: "create", name: "B" },
        { kind: "delete", name: "A" },
      ]),
    ).toEqual([{ kind: "create", name: "B" }]);
  });

  it("rewrites a session-created sheet's create op on rename", () => {
    expect(
      run([
        { kind: "create", name: "Sheet1" },
        { kind: "rename", oldName: "Sheet1", newName: "Scratch" },
      ]),
    ).toEqual([{ kind: "create", name: "Scratch" }]);
  });

  it("collapses consecutive renames a→b, b→c to a→c", () => {
    expect(
      run([
        { kind: "rename", oldName: "A", newName: "B" },
        { kind: "rename", oldName: "B", newName: "C" },
      ]),
    ).toEqual([{ kind: "rename", oldName: "A", newName: "C" }]);
  });

  it("create → rename → delete nets to nothing", () => {
    expect(
      run([
        { kind: "create", name: "Sheet1" },
        { kind: "rename", oldName: "Sheet1", newName: "Scratch" },
        { kind: "delete", name: "Scratch" },
      ]),
    ).toEqual([]);
  });

  it("does not compact a delete across a rename of a pre-existing sheet", () => {
    // "Model" existed in the file; rename then delete must both replay.
    expect(
      run([
        { kind: "rename", oldName: "Model", newName: "Model v2" },
        { kind: "delete", name: "Model v2" },
      ]),
    ).toEqual([
      { kind: "rename", oldName: "Model", newName: "Model v2" },
      { kind: "delete", name: "Model v2" },
    ]);
  });

  it("handles recreate after delete of the same name", () => {
    expect(
      run([
        { kind: "delete", name: "Notes" },
        { kind: "create", name: "Notes" },
      ]),
    ).toEqual([
      { kind: "delete", name: "Notes" },
      { kind: "create", name: "Notes" },
    ]);
  });
});
