import { describe, it, expect, vi, beforeEach } from "vitest";

// Links between open workbooks, on the webview's side: the engine bridge
// carries what a command changed in OTHER tabs to those tabs' grids, and a
// save folds the file's spelling of the links into its patch. The Tauri
// side is a fake that answers the way the Rust workspace layer does.
const calls: Array<{ cmd: string; args: any }> = [];
let answers: Record<string, (args: any) => any> = {};
vi.mock("@tauri-apps/api/core", () => ({
  invoke: async (cmd: string, args: any) => {
    calls.push({ cmd, args });
    const answer = answers[cmd];
    return answer ? answer(args) : null;
  },
}));

const engine = await import("../ironcalcEngine");
const { mergeLinkPlan } = await import("../workspaceLinks");

const layoutBytes = () => new TextEncoder().encode(JSON.stringify({ sheets: [] })).buffer;

beforeEach(() => {
  calls.length = 0;
  answers = {
    calc_engine_workbook_layout: () => layoutBytes(),
  };
  (globalThis as any).localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {} };
});

/** Opens `tab` on `path` the way the viewport grid does, with a grid applier that records what it is given. */
async function openTab(tab: string, path: string, sheets: string[], load: Record<string, unknown> = {}) {
  answers.calc_engine_load = () => ({ sheets, load_ms: 0, eval_ms: 0, others: [], needs_saving: false, ...load });
  const applied: Array<{ sheet: string; cells: Array<{ row0: number; col0: number; v: unknown }> }> = [];
  await engine.ironcalcEngineOpenLayout(tab, path);
  engine.registerIroncalcApplier(tab, (sheet, cells) => {
    applied.push({ sheet, cells });
    return true;
  });
  return applied;
}

describe("engine bridge: workbooks that read each other", () => {
  it("asks for links on a viewport open and not when they are switched off", async () => {
    await openTab("t-on", "/m/NOW.xlsx", ["Model"]);
    expect(calls.find((c) => c.cmd === "calc_engine_load")?.args).toMatchObject({ tab: "t-on", path: "/m/NOW.xlsx", links: true });

    (globalThis as any).localStorage = { getItem: (k: string) => (k === "gridpath.workspaceLinks" ? "0" : null), setItem: () => {}, removeItem: () => {} };
    calls.length = 0;
    await openTab("t-off", "/m/ANT.xlsx", ["Model"]);
    expect(calls.find((c) => c.cmd === "calc_engine_load")?.args.links).toBe(false);
    engine.ironcalcEngineOnClose("t-on");
    engine.ironcalcEngineOnClose("t-off");
  });

  it("gives an edit's effect on another workbook to that workbook's grid", async () => {
    const now = await openTab("t-now", "/m/NOW.xlsx", ["Model"]);
    const comps = await openTab("t-comps", "/m/Comps.xlsx", ["Comps", "[NOW.xlsx]Model"]);
    answers.calc_engine_edit_batch = (args) => {
      expect(args.tab).toBe("t-now");
      return {
        recalc_ms: 1,
        changed: [{ sheet: 0, row: 32, column: 74, value: 5 }],
        others: [{ tab: "t-comps", changed: [{ sheet: 0, row: 2, column: 2, value: 5 }] }],
        sheets: null,
        links: null,
      };
    };
    engine.ironcalcEngineWriteCells("t-now", "Model", "typed", [{ row: 32, column: 74, value: "5" }]);
    await engine.ironcalcEngineSettled("t-now");

    expect(now).toEqual([{ sheet: "Model", cells: [{ row0: 31, col0: 73, v: 5 }] }]);
    expect(comps).toEqual([{ sheet: "Comps", cells: [{ row0: 1, col0: 1, v: 5 }] }]);
    engine.ironcalcEngineOnClose("t-now");
    engine.ironcalcEngineOnClose("t-comps");
  });

  it("takes the new sheet list when a typed link adds a hidden copy, and reports links left as typed", async () => {
    const comps = await openTab("t-comps", "/m/Comps.xlsx", ["Comps"]);
    const events: any[] = [];
    const off = engine.onWorkspaceEvent((e) => events.push(e));
    answers.calc_engine_edit_batch = () => ({
      recalc_ms: 1,
      changed: [{ sheet: 0, row: 1, column: 1, value: 7 }],
      others: [],
      sheets: ["Comps", "[NOW.xlsx]Model"],
      links: { changed: true, notes: ["Missing.xlsx could not be opened"] },
    });
    engine.ironcalcEngineWriteCells("t-comps", "Comps", "typed", [{ row: 1, column: 1, value: "='[NOW.xlsx]Model'!BV32" }]);
    await engine.ironcalcEngineSettled("t-comps");

    expect(engine.ironcalcEngineSheetNames("t-comps")).toEqual(["Comps", "[NOW.xlsx]Model"]);
    expect(comps).toEqual([{ sheet: "Comps", cells: [{ row0: 0, col0: 0, v: 7 }] }]);
    expect(events).toContainEqual({ kind: "notes", tab: "t-comps", notes: ["Missing.xlsx could not be opened"] });
    expect(events).toContainEqual({ kind: "changed" });
    off();
    engine.ironcalcEngineOnClose("t-comps");
  });

  it("tells the grid of a workbook whose formulas followed inserted rows to read them again", async () => {
    await openTab("t-now", "/m/NOW.xlsx", ["Model"]);
    const comps = await openTab("t-comps", "/m/Comps.xlsx", ["Comps", "[NOW.xlsx]Model"]);
    let reloads = 0;
    engine.registerIroncalcReloader("t-comps", () => reloads++);
    const events: any[] = [];
    const off = engine.onWorkspaceEvent((e) => events.push(e));
    answers.calc_engine_structural = () => ({
      ms: 1,
      cells: [],
      others: [{ tab: "t-comps", changed: [{ sheet: 0, row: 2, column: 2, value: 9 }] }],
      moved: ["t-comps"],
    });
    engine.ironcalcEngineStructuralDirect("t-now", "insert_rows", "Model", 2, 2);
    await engine.ironcalcEngineSettled("t-now");

    expect(reloads).toBe(1);
    expect(comps).toEqual([{ sheet: "Comps", cells: [{ row0: 1, col0: 1, v: 9 }] }]);
    expect(events).toContainEqual({ kind: "moved", tabs: ["t-comps"], by: "t-now" });
    off();
    engine.ironcalcEngineOnClose("t-now");
    engine.ironcalcEngineOnClose("t-comps");
  });

  it("reports a workbook opened with changes it was given while it had no tab", async () => {
    const events: any[] = [];
    const off = engine.onWorkspaceEvent((e) => events.push(e));
    await openTab("t-comps", "/m/Comps.xlsx", ["Comps", "[NOW.xlsx]Model"], { needs_saving: true });
    expect(events).toContainEqual({ kind: "needs_saving", tab: "t-comps" });
    off();
    engine.ironcalcEngineOnClose("t-comps");
  });

  it("closes a tab saying whether it had unsaved changes, and applies what that takes back", async () => {
    await openTab("t-now", "/m/NOW.xlsx", ["Model"]);
    const comps = await openTab("t-comps", "/m/Comps.xlsx", ["Comps", "[NOW.xlsx]Model"]);
    answers.calc_engine_unload = (args) => (args.tab === "t-now" ? [{ tab: "t-comps", changed: [{ sheet: 0, row: 2, column: 2, value: 1 }] }] : []);
    engine.ironcalcEngineOnClose("t-now", true);
    await new Promise((r) => setTimeout(r, 0));

    expect(calls.find((c) => c.cmd === "calc_engine_unload")?.args).toEqual({ tab: "t-now", dirty: true });
    expect(comps).toEqual([{ sheet: "Comps", cells: [{ row0: 1, col0: 1, v: 1 }] }]);
    engine.ironcalcEngineOnClose("t-comps");
  });

  it("asks for the save plan after queued edits, and marks the file saved", async () => {
    await openTab("t-comps", "/m/Comps.xlsx", ["Comps"]);
    answers.calc_engine_edit_batch = async () => {
      await new Promise((r) => setTimeout(r, 10));
      return { recalc_ms: 1, changed: [], others: [], sheets: null, links: null };
    };
    answers.calc_engine_link_save_plan = () => ({ cells: [], rewriteAll: false, externalLinks: [], reads: [] });
    answers.calc_engine_mark_saved = () => [];
    engine.ironcalcEngineWriteCells("t-comps", "Comps", "typed", [{ row: 1, column: 1, value: "1" }]);
    const plan = await engine.ironcalcEngineLinkSavePlan("t-comps", "/m/Copy.xlsx");
    expect(plan).toEqual({ cells: [], rewriteAll: false, externalLinks: [], reads: [] });
    const order = calls.map((c) => c.cmd).filter((c) => c === "calc_engine_edit_batch" || c === "calc_engine_link_save_plan");
    expect(order).toEqual(["calc_engine_edit_batch", "calc_engine_link_save_plan"]);
    expect(calls.find((c) => c.cmd === "calc_engine_link_save_plan")?.args).toEqual({ tab: "t-comps", targetPath: "/m/Copy.xlsx" });

    await engine.ironcalcEngineMarkSaved("t-comps", "/m/Copy.xlsx");
    expect(calls.find((c) => c.cmd === "calc_engine_mark_saved")?.args).toEqual({ tab: "t-comps", path: "/m/Copy.xlsx" });
    engine.ironcalcEngineOnClose("t-comps");
  });
});

describe("save patch: the file's spelling of links", () => {
  const link = { index: 1, target: "NOW.xlsx", sheets: [{ name: "Model", areas: [], cells: [{ r: 31, c: 73, v: { t: "n", n: 5 } }] }] };

  it("respells a formula the session wrote and gives it its value", () => {
    const patch: any = { version: 1, sheets: [{ name: "Comps", cells: [{ r: 1, c: 1, f: "'[NOW.xlsx]Model'!BV32" }, { r: 0, c: 0, v: { t: "s", s: "Company" } }] }] };
    const res = mergeLinkPlan(patch, {
      cells: [{ sheet: "Comps", r: 1, c: 1, f: "[1]Model!BV32", v: { t: "n", n: 5 } }],
      rewriteAll: false,
      externalLinks: [link],
      reads: ["/m/NOW.xlsx"],
    });
    expect(res).toEqual({ respelled: 1, added: 0, parts: 1 });
    expect(patch.sheets[0].cells).toEqual([
      { r: 1, c: 1, f: "[1]Model!BV32", v: { t: "n", n: 5 } },
      { r: 0, c: 0, v: { t: "s", s: "Company" } },
    ]);
    expect(patch.externalLinks).toEqual([link]);
  });

  it("leaves formulas the session did not write alone unless the links moved", () => {
    const plan = {
      cells: [
        { sheet: "Comps", r: 1, c: 1, f: "[1]Model!BV34", v: { t: "n", n: 5 } },
        { sheet: "Other", r: 4, c: 0, f: "SUM([1]Model!B3:D3)" },
      ],
      rewriteAll: false,
      externalLinks: [{ index: 1, sheets: link.sheets }],
      reads: ["/m/NOW.xlsx"],
    };
    const untouched: any = { version: 1 };
    expect(mergeLinkPlan(untouched, plan)).toEqual({ respelled: 0, added: 0, parts: 1 });
    expect(untouched.sheets).toBeUndefined();

    const moved: any = { version: 1, sheets: [{ name: "Comps", cells: [{ r: 0, c: 0, v: { t: "s", s: "x" } }] }] };
    expect(mergeLinkPlan(moved, { ...plan, rewriteAll: true })).toEqual({ respelled: 0, added: 2, parts: 1 });
    expect(moved.sheets).toEqual([
      { name: "Comps", cells: [{ r: 0, c: 0, v: { t: "s", s: "x" } }, { r: 1, c: 1, f: "[1]Model!BV34", v: { t: "n", n: 5 } }] },
      { name: "Other", cells: [{ r: 4, c: 0, f: "SUM([1]Model!B3:D3)" }] },
    ]);
  });

  it("keeps what the session did to a cell that is no longer a link", () => {
    const patch: any = { version: 1, sheets: [{ name: "Comps", cells: [{ r: 1, c: 1, clear: true }, { r: 2, c: 1, v: { t: "n", n: 3 } }] }] };
    mergeLinkPlan(patch, {
      cells: [
        { sheet: "Comps", r: 1, c: 1, f: "[1]Model!B2" },
        { sheet: "Comps", r: 2, c: 1, f: "[1]Model!B3" },
      ],
      rewriteAll: true,
      externalLinks: [],
      reads: [],
    });
    expect(patch.sheets[0].cells).toEqual([{ r: 1, c: 1, clear: true }, { r: 2, c: 1, v: { t: "n", n: 3 } }]);
    expect(patch.externalLinks).toBeUndefined();
  });

  it("keeps an @ the user wrote, drops one the engine added, and marks newer functions", () => {
    const patch: any = {
      version: 1,
      sheets: [{ name: "S", cells: [{ r: 0, c: 0, f: "@'[NOW.xlsx]Model'!B:B" }, { r: 1, c: 0, f: "XLOOKUP(1,'[NOW.xlsx]Model'!A:A,'[NOW.xlsx]Model'!B:B)" }] }],
    };
    mergeLinkPlan(patch, {
      cells: [
        { sheet: "S", r: 0, c: 0, f: "@[1]Model!B:B" },
        { sheet: "S", r: 1, c: 0, f: "XLOOKUP(1,@[1]Model!A:A,[1]Model!B:B)" },
      ],
      rewriteAll: false,
      externalLinks: [],
      reads: [],
    });
    expect(patch.sheets[0].cells[0].f).toBe("@[1]Model!B:B");
    expect(patch.sheets[0].cells[1].f).toBe("_xlfn.XLOOKUP(1,[1]Model!A:A,[1]Model!B:B)");
  });

  it("does nothing without a plan", () => {
    const patch: any = { version: 1 };
    expect(mergeLinkPlan(patch, null)).toEqual({ respelled: 0, added: 0, parts: 0 });
    expect(patch).toEqual({ version: 1 });
  });
});
