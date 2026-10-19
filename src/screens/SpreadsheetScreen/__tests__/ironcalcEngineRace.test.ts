import { describe, it, expect, vi, beforeEach } from "vitest";

// Drive the engine mirror through its public entry points with a fake Tauri
// bridge. The sheet-op call is deliberately slow so the write that follows a
// create_sheet arrives before the engine knows the sheet — the sequence that
// used to fire a resync (and, on a >2M-cell grid, degrade the tab).
const calls: Array<{ cmd: string; args: any }> = [];
let sheets: string[] = [];
vi.mock("@tauri-apps/api/core", () => ({
  invoke: async (cmd: string, args: any) => {
    calls.push({ cmd, args });
    if (cmd === "calc_engine_new") {
      sheets = [...args.sheets];
      return { sheets, load_ms: 0, eval_ms: 0 };
    }
    if (cmd === "calc_engine_sheet_op") {
      await new Promise((r) => setTimeout(r, 15));
      if (args.op === "add") sheets = [...sheets, args.name];
      return { sheets, cells: [] };
    }
    if (cmd === "calc_engine_edit_batch") return { recalc_ms: 1, changed: [] };
    if (cmd === "calc_engine_resync") return { sheets, cells: [], seeded: 0, ms: 0 };
    if (cmd === "workbook_file_stamp") return { size: 1, mtime_ms: 1 };
    return null;
  },
}));

const engine = await import("../ironcalcEngine");

beforeEach(() => {
  calls.length = 0;
  sheets = [];
  (globalThis as any).localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {} };
});

describe("engine mirror: write right after create_sheet", () => {
  it("resolves the new sheet's index when the queued edit runs, without a resync", async () => {
    const tab = "t-race";
    engine.ironcalcEngineOnBlank(tab, ["Sheet1"]);
    await engine.ironcalcEngineSettled(tab);
    const names = new Map([["s1", "Sheet1"], ["s2", "Recon"]]);

    // Univer: insert-sheet mutation, then its parent command…
    engine.ironcalcEngineOnCommand(tab, { id: "sheet.mutation.insert-sheet", params: { sheet: { id: "s2", name: "Recon" } } }, names);
    engine.ironcalcEngineOnCommand(tab, { id: "sheet.command.insert-sheet", params: {} }, names);
    // …and the agent's write to the new sheet in the same tick.
    engine.ironcalcEngineOnCommand(
      tab,
      { id: "sheet.command.set-range-values", params: { subUnitId: "s2", value: { 9: { 0: { f: "=1+1" } } } } },
      names,
    );
    await engine.ironcalcEngineSettled(tab);

    expect(calls.filter((c) => c.cmd === "calc_engine_resync")).toHaveLength(0);
    expect(engine.ironcalcEngineDegradedReason(tab)).toBeNull();
    const edit = calls.find((c) => c.cmd === "calc_engine_edit_batch");
    expect(edit?.args.edits).toEqual([{ sheet: 1, row: 10, column: 1, value: "=1+1" }]);
    const order = calls.map((c) => c.cmd).filter((c) => c.startsWith("calc_engine_"));
    expect(order).toEqual(["calc_engine_new", "calc_engine_sheet_op", "calc_engine_edit_batch"]);
  });

  it("mirrors the clear_range tool (clearContent → `ranges: [...]`) as empty writes, no resync", async () => {
    const tab = "t-clear";
    engine.ironcalcEngineOnBlank(tab, ["Sheet1"]);
    await engine.ironcalcEngineSettled(tab);
    const names = new Map([["s1", "Sheet1"]]);
    engine.ironcalcEngineOnCommand(
      tab,
      { id: "sheet.command.clear-selection-content", params: { unitId: "u", subUnitId: "s1", ranges: [{ startRow: 2, endRow: 3, startColumn: 1, endColumn: 2 }] } },
      names,
    );
    await engine.ironcalcEngineSettled(tab);
    expect(calls.filter((c) => c.cmd === "calc_engine_resync")).toHaveLength(0);
    expect(engine.ironcalcEngineDegradedReason(tab)).toBeNull();
    const edit = calls.find((c) => c.cmd === "calc_engine_edit_batch");
    expect(edit?.args.edits).toEqual([
      { sheet: 0, row: 3, column: 2, value: "" },
      { sheet: 0, row: 3, column: 3, value: "" },
      { sheet: 0, row: 4, column: 2, value: "" },
      { sheet: 0, row: 4, column: 3, value: "" },
    ]);
  });

  it("adds a sheet the engine has never seen instead of resyncing (and degrading a big grid)", async () => {
    // A write names a sheet the grid has — the sheet exists, only its add
    // never reached the engine (insert-sheet mutation never claimed, or a
    // script auto-created its destination in the same apply pass). Mirror
    // the add, then the write; never resync for this.
    const tab = "t-unknown";
    engine.ironcalcEngineOnBlank(tab, ["Sheet1"]);
    await engine.ironcalcEngineSettled(tab);
    const names = new Map([["s1", "Sheet1"], ["s9", "Ghost"]]);
    engine.ironcalcEngineOnCommand(
      tab,
      { id: "sheet.command.set-range-values", params: { subUnitId: "s9", value: { 0: { 0: { v: 1 } } } } },
      names,
    );
    await engine.ironcalcEngineSettled(tab);
    expect(calls.filter((c) => c.cmd === "calc_engine_resync")).toHaveLength(0);
    expect(engine.ironcalcEngineDegradedReason(tab)).toBeNull();
    const add = calls.find((c) => c.cmd === "calc_engine_sheet_op");
    expect(add?.args).toMatchObject({ op: "add", name: "Ghost" });
    const edit = calls.find((c) => c.cmd === "calc_engine_edit_batch");
    expect(edit?.args.edits).toEqual([{ sheet: 1, row: 1, column: 1, value: "1" }]);
  });
});
