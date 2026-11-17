import { describe, it, expect, vi, beforeEach } from "vitest";
import { buildWorkbookIndex } from "../agent/workbookIndex";

// The agent's reads of another workbook of the workspace. The Tauri side is
// a fake holding one model, "SNOW.xlsx": a `Model` sheet, and before it the
// hidden copy of a sheet of a third workbook (so sheet indices are not the
// positions in the list the agent sees).
const calls: Array<{ cmd: string; args: any }> = [];
const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).buffer;

// Model sheet: labels in column A, three periods in B:D (0-based rows).
const MODEL: Record<number, Record<number, { v?: unknown; f?: string }>> = {
  0: { 0: { v: "$ millions" }, 1: { v: "FY2025" }, 2: { v: "FY2026" }, 3: { v: "FY2027E" } },
  1: { 0: { v: "Product revenue" }, 1: { v: 2600 }, 2: { v: 3400 }, 3: { f: "=C2*1.25", v: 4250 } },
  2: { 0: { v: "Services revenue" }, 1: { v: 180 }, 2: { v: 200 }, 3: { f: "=C3*1.1", v: 220 } },
  3: { 0: { v: "Total revenue" }, 1: { f: "=B2+B3", v: 2780 }, 2: { f: "=C2+C3", v: 3600 }, 3: { f: "=@D2+D3", v: 4470 } },
};
const index = buildWorkbookIndex({ sheetOrder: ["s1"], sheets: { s1: { name: "Model", cellData: MODEL } }, styles: {} });

let modelLoads = 0;
let dropOnce = false;
vi.mock("@tauri-apps/api/core", () => ({
  invoke: async (cmd: string, args: any) => {
    calls.push({ cmd, args });
    if (cmd === "workspace_model") {
      if (!String(args.path).endsWith("/SNOW.xlsx")) throw new Error("the file was not found");
      modelLoads++;
      return { key: "bg:/m/SNOW.xlsx", sheets: [{ name: "[ANT.xlsx]P&L", link: true }, { name: "Model", link: false }] };
    }
    if (dropOnce && cmd.startsWith("calc_engine_")) {
      dropOnce = false;
      throw new Error("no workbook loaded for tab");
    }
    if (cmd === "calc_engine_sheet_window") {
      const cells: any[] = [];
      for (let r = args.r0; r <= args.r1; r++) {
        for (let c = args.c0; c <= args.c1; c++) {
          const cell = args.sheet === 1 ? MODEL[r - 1]?.[c - 1] : undefined;
          if (cell) cells.push({ r, c, v: cell.v, ...(cell.f ? { f: cell.f } : {}), ...(r === 4 && c === 4 ? { d: "4,470" } : {}) });
        }
      }
      return bytes({ sheet: args.sheet, cells, styles: {} });
    }
    if (cmd === "calc_engine_workbook_index") return bytes(index);
    if (cmd === "calc_engine_query_rows") return { matched: 2, rows: [], sheet_index: args.query.sheet };
    return null;
  },
}));

const { readWorkspaceWorkbook, workspaceBooks, linkPrefix, rangeBounds } = await import("../agent/workspaceRead");

const member = (path: string, over: Record<string, unknown> = {}) => ({
  path,
  name: path.split("/").pop()!,
  state: "idle" as const,
  member: true,
  reads: [],
  readBy: [],
  broken: [],
  needsSaving: false,
  error: null,
  ...over,
});
const BOOKS = [{ label: "SNOW.xlsx", path: "/m/SNOW.xlsx" }];
const read = async (name: string, input: any, books = BOOKS) => JSON.parse(await readWorkspaceWorkbook(name, input, books));

beforeEach(() => {
  calls.length = 0;
  modelLoads = 0;
  dropOnce = false;
});

describe("which workbooks the agent can read", () => {
  it("offers the workspace's other models, by file name", () => {
    const members = [
      member("/m/SNPS.xlsx", { state: "open" }),
      member("/m/SNOW.xlsx"),
      member("/m/Gone.xlsx", { state: "missing" }),
      member("/m/Bad.xlsx", { state: "failed" }),
      member("/other/SNOW.xlsx"),
      member("/old/SNPS.xlsx"),
      member("/m/NOW.xlsx", { member: false, state: "loaded" }),
    ];
    expect(workspaceBooks(members, "/m/SNPS.xlsx")).toEqual([
      { label: "SNOW.xlsx", path: "/m/SNOW.xlsx" },
      { label: "NOW.xlsx", path: "/m/NOW.xlsx" },
    ]);
    expect(workspaceBooks(members, null).map((b) => b.label)).toEqual(["SNPS.xlsx", "SNOW.xlsx", "NOW.xlsx"]);
  });

  it("spells the link prefix quoted", () => {
    expect(linkPrefix("SNOW.xlsx", "Model")).toBe("'[SNOW.xlsx]Model'!");
    expect(linkPrefix("Bob's model.xlsx", "P&L '26")).toBe("'[Bob''s model.xlsx]P&L ''26'!");
  });

  it("reads A1 ranges", () => {
    expect(rangeBounds("B5:D20")).toEqual({ r0: 5, c0: 2, r1: 20, c1: 4 });
    expect(rangeBounds("$d$20:b5")).toEqual({ r0: 5, c0: 2, r1: 20, c1: 4 });
    expect(rangeBounds("C7")).toEqual({ r0: 7, c0: 3, r1: 7, c1: 3 });
    expect(rangeBounds("A:A")).toBeNull();
    expect(rangeBounds("Model!A1")).toBeNull();
  });
});

describe("reading another workbook", () => {
  it("reads a range from that workbook's model, with what a link to it looks like", async () => {
    const out = await read("read_range", { workbook: "snow.xlsx", sheet: "Model", range: "A4:D4" });
    expect(out).toMatchObject({ workbook: "SNOW.xlsx", sheet: "Model", range: "A4:D4", truncated: false, link_prefix: "'[SNOW.xlsx]Model'!" });
    expect(out.cells).toEqual([
      { cell: "A4", value: "Total revenue", formula: null },
      { cell: "B4", value: 2780, formula: "=B2+B3" },
      { cell: "C4", value: 3600, formula: "=C2+C3" },
      { cell: "D4", value: 4470, formula: "=D2+D3", display: "4,470" },
    ]);
    expect(out.link_note).toContain("='[SNOW.xlsx]Model'!A4");
    // The engine was asked under the model's key, for the sheet's engine index (the hidden copy is index 0).
    expect(calls.find((c) => c.cmd === "workspace_model")?.args).toEqual({ path: "/m/SNOW.xlsx" });
    expect(calls.find((c) => c.cmd === "calc_engine_sheet_window")?.args).toEqual({ tab: "bg:/m/SNOW.xlsx", sheet: 1, r0: 4, r1: 4, c0: 1, c1: 4 });
  });

  it("gives at most 500 cells and says so", async () => {
    const out = await read("read_range", { workbook: "SNOW.xlsx", sheet: "Model", range: "A1:Z100" });
    expect(out.truncated).toBe(true);
    const asked = calls.find((c) => c.cmd === "calc_engine_sheet_window")!.args;
    expect((asked.r1 - asked.r0 + 1) * (asked.c1 - asked.c0 + 1)).toBeLessThanOrEqual(500);
    expect(asked).toMatchObject({ r0: 1, c0: 1, c1: 26, r1: 19 });
  });

  it("finds rows by label, with a sample and the link to it", async () => {
    const out = await read("find_rows", { workbook: "SNOW.xlsx", query: "total revenue" });
    expect(out.workbook).toBe("SNOW.xlsx");
    expect(out.matches).toHaveLength(1);
    expect(out.matches[0]).toMatchObject({ sheet: "Model", row: 4, label: "Total revenue", sample: "D4 = =D2+D3 → 4470", sample_link: "='[SNOW.xlsx]Model'!D4" });
  });

  it("maps the workbook", async () => {
    const out = await read("describe_workbook", { workbook: "SNOW.xlsx" });
    expect(out.workbook).toBe("SNOW.xlsx");
    expect(out.link_form).toContain("='[SNOW.xlsx]Model'!B2");
    expect(JSON.stringify(out)).toContain("Model");
    expect(JSON.stringify(out)).not.toContain("[ANT.xlsx]");
  });

  it("queries rows on the sheet's engine index", async () => {
    const out = await read("query_rows", { workbook: "SNOW.xlsx", sheet: "Model", where: [{ column: "A", op: "contains", value: "revenue" }] });
    expect(out).toMatchObject({ workbook: "SNOW.xlsx", matched: 2, sheet_index: 1, link_prefix: "'[SNOW.xlsx]Model'!" });
  });

  it("says what is there when the name is wrong", async () => {
    expect(await read("read_range", { workbook: "Snowflake.xlsx", sheet: "Model", range: "A1" })).toMatchObject({ ok: false, error: "unknown_workbook", available: ["SNOW.xlsx"] });
    expect(await read("read_range", { workbook: "SNOW.xlsx", sheet: "model", range: "A1" })).toMatchObject({ ok: false, error: "unknown_sheet", available: ["Model"] });
    // The hidden copy of another workbook's sheet is not a sheet of this one.
    expect(await read("read_range", { workbook: "SNOW.xlsx", sheet: "[ANT.xlsx]P&L", range: "A1" })).toMatchObject({ ok: false, error: "unknown_sheet" });
    expect(await read("read_range", { workbook: "SNOW.xlsx", sheet: "Model", range: "A:A" })).toMatchObject({ ok: false, error: "bad_range" });
    expect(await read("find_rows", { workbook: "SNOW.xlsx" })).toMatchObject({ ok: false, error: "bad_arguments" });
    // A wrong workbook name or missing arguments never load a model.
    expect(calls.filter((c) => c.cmd === "workspace_model")).toHaveLength(3);
  });

  it("reports a workbook that cannot be loaded instead of failing", async () => {
    const out = await read("describe_workbook", { workbook: "Lost.xlsx" }, [{ label: "Lost.xlsx", path: "/m/Lost.xlsx" }]);
    expect(out).toMatchObject({ ok: false, error: "workbook_unavailable", workbook: "Lost.xlsx" });
    expect(out.message).toContain("the file was not found");
  });

  it("loads the model again when it was let go between the two calls", async () => {
    dropOnce = true;
    const out = await read("read_range", { workbook: "SNOW.xlsx", sheet: "Model", range: "B2" });
    expect(out.cells).toEqual([{ cell: "B2", value: 2600, formula: null }]);
    expect(modelLoads).toBe(2);
  });
});
