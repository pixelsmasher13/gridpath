/**
 * Several workbooks worked on together through the tool handlers: a formula
 * in one reads another, and edits to the one that is read (values, inserted
 * rows, renamed sheets, rejected batches) reach the ones that read it, in
 * this process and after a restart. And saving: the links go into the file
 * the way Excel writes them, and are live again when the file is opened.
 *
 * Needs an engine build that can link workbooks; skipped on an older one.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, it, expect, afterEach } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { ToolHandlers } from "../tools/handlers";
import { createServer } from "../mcp/server";
import { diffForReview } from "../review/api";
import { loadEngine } from "./session";
import { readersPathFor, storePathFor } from "./store";
import { linkTarget } from "./workspace";

let linking = false;
try {
  const engine = loadEngine() as any;
  linking =
    typeof engine.Workbook?.prototype?.attachLink === "function" &&
    typeof engine.Workbook?.prototype?.linkCache === "function" &&
    typeof engine.externalLinks === "function" &&
    typeof engine.blankXlsx === "function";
} catch {
  linking = false;
}

const dirs: string[] = [];
afterEach(() => {
  for (const d of dirs.splice(0)) fs.rmSync(d, { recursive: true, force: true });
});

/** A folder of its own, and tool handlers allowed to work in it. */
function workspace() {
  // realpath: on macOS the temp folder is reached through a symlink.
  const dir = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "gridpath-ws-")));
  dirs.push(dir);
  const handlers = () => new ToolHandlers({ allowRoots: [dir] });
  return { dir, handlers, at: (name: string) => path.join(dir, name) };
}

type Call = (name: string, args: Record<string, unknown>) => Promise<any>;
const caller = (h: ToolHandlers): Call => (name, args) => h.call(name, args) as Promise<any>;

/**
 * A new process starts later than the one before it stopped; batches are
 * ordered by the time they were made, to the millisecond.
 */
const later = () => new Promise((r) => setTimeout(r, 3));
async function restart(handlers: () => ToolHandlers): Promise<Call> {
  await later();
  return caller(handlers());
}

/** The value and formula of one cell. */
async function cell(call: Call, file: string, sheet: string, a1: string): Promise<{ value: unknown; formula: string | null }> {
  const r = await call("read_range", { path: file, sheet, range: a1 });
  expect(r.ok, JSON.stringify(r)).not.toBe(false);
  const c = r.cells[0];
  return { value: c?.value ?? null, formula: c?.formula ?? null };
}

/** `src.xlsx` (sheet Data: a label column, an input and a formula) and an empty `sum.xlsx`. */
async function pair(call: Call, at: (n: string) => string, srcName = "src.xlsx", sheet = "Data") {
  const src = at(srcName);
  const sum = at("sum.xlsx");
  expect((await call("create_workbook", { path: src, sheets: [sheet, "Other"] })).ok).toBe(true);
  expect((await call("create_workbook", { path: sum })).ok).toBe(true);
  const w = await call("set_range", {
    path: src,
    sheet,
    top_left: "A1",
    values: [
      ["Units", 100],
      ["Revenue", "=B1*2"],
      ["Costs", 30],
    ],
  });
  expect(w.ok, JSON.stringify(w)).toBe(true);
  return { src, sum };
}

const stored = (file: string) => JSON.parse(fs.readFileSync(storePathFor(file), "utf8")) as { batches: any[] };

/** The links to other workbooks a file on disk holds. */
const linksIn = (file: string) => JSON.parse((loadEngine() as any).externalLinks(fs.readFileSync(file))) as Array<{ index: number; target: string | null; sheets: string[] }>;

describe.skipIf(!linking)("workbooks worked on together", () => {
  it("a formula reads another workbook and follows its edits", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);

    const w = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2+1" });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.links).toEqual([{ link: "[src.xlsx]Data", status: "linked", workbook: src, sheet: "Data" }]);
    expect(w.cells[0].value).toBe(201);
    expect(w.cells[0].formula).toContain("[src.xlsx]Data");

    const e = await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 150 });
    expect(e.ok).toBe(true);
    expect(e.recalculated_workbooks).toEqual([sum]);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(301);

    // An edit that changes nothing the reader reads does not touch it.
    const quiet = await call("set_cell", { path: src, sheet: "Other", cell: "A1", value: 5 });
    expect(quiet.recalculated_workbooks).toBeUndefined();
  });

  it("finds the file and the sheet whatever the capitals, and spells them as they are", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    const w = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[SRC]data!B1" });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.links[0]).toMatchObject({ status: "linked", workbook: src, sheet: "Data" });
    expect(w.cells[0].value).toBe(100);
    expect(w.cells[0].formula).toContain("[src.xlsx]Data");
  });

  it("links through names with spaces, an apostrophe and an ampersand", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at, "Bob's Q3&Q4 model.xlsx", "Income Statement");
    const w = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "='[Bob''s Q3&Q4 model.xlsx]Income Statement'!B2" });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.cells[0].value).toBe(200);
    await call("set_cell", { path: src, sheet: "Income Statement", cell: "B1", value: 7 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(14);
    // The formula the engine shows is one it reads back.
    const shown = (await cell(call, sum, "Sheet1", "A1")).formula!;
    const again = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A2", formula: shown });
    expect(again.ok, JSON.stringify(again)).toBe(true);
    expect(again.cells[0].value).toBe(14);
  });

  it("a row or column inserted in the workbook that is read moves the reader's links", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_range", {
      path: sum,
      sheet: "Sheet1",
      top_left: "A1",
      values: [["=[src.xlsx]Data!B2", "=SUM([src.xlsx]Data!B1:B3)", '=INDEX([src.xlsx]Data!B1:B9, MATCH("Costs", [src.xlsx]Data!A1:A9, 0))']],
    });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    expect((await cell(call, sum, "Sheet1", "B1")).value).toBe(330);
    expect((await cell(call, sum, "Sheet1", "C1")).value).toBe(30);

    const ins = await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 2 });
    expect(ins.ok, JSON.stringify(ins)).toBe(true);
    expect(ins.recalculated_workbooks).toEqual([sum]);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B4");
    expect(a1.value).toBe(200);
    expect((await cell(call, sum, "Sheet1", "B1")).formula).toContain("B3:B5");
    expect((await cell(call, sum, "Sheet1", "B1")).value).toBe(330);
    expect((await cell(call, sum, "Sheet1", "C1")).value).toBe(30);

    const col = await call("insert_columns", { path: src, sheet: "Data", before: "A", count: 1 });
    expect(col.ok, JSON.stringify(col)).toBe(true);
    expect((await cell(call, sum, "Sheet1", "A1")).formula).toContain("C4");
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);

    // The source's cells moved: an edit at the new place still reaches the reader.
    await call("set_cell", { path: src, sheet: "Data", cell: "C3", value: 1000 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(2000);
    expect((await cell(call, sum, "Sheet1", "B1")).value).toBe(3030);

    // The repeat of the row edit is the reader's own record, and is not a batch to review.
    const listed = await call("list_batches", { path: sum });
    expect(listed.batches).toHaveLength(1);
    expect(listed.batches[0].reads).toEqual([src]);
    expect(stored(sum).batches.filter((b) => b.link_origin)).toHaveLength(2);
  });

  it("reaches a workbook that reads a workbook that reads the one edited", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    const top = at("top.xlsx");
    await call("create_workbook", { path: top });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2*10" });
    await call("set_cell", { path: top, sheet: "Sheet1", cell: "A1", formula: "=[sum.xlsx]Sheet1!A1+1" });
    expect((await cell(call, top, "Sheet1", "A1")).value).toBe(2001);
    const e = await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 1 });
    expect(e.recalculated_workbooks.sort()).toEqual([sum, top].sort());
    expect((await cell(call, top, "Sheet1", "A1")).value).toBe(21);

    // A link back down the chain would go round in a circle.
    const back = await call("set_cell", { path: src, sheet: "Data", cell: "D1", formula: "=[top.xlsx]Sheet1!A1" });
    expect(back.ok).toBe(false);
    expect(back.links[0].status).toBe("circular");
    expect((await cell(call, src, "Data", "D1")).value).toBeNull();
  });

  it("refuses a write whose link leads nowhere, and writes nothing", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { sum } = await pair(call, at);

    const noFile = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[missing.xlsx]Data!B2" });
    expect(noFile.ok).toBe(false);
    expect(noFile.links[0].status).toBe("workbook_not_found");

    const noSheet = await call("set_range", { path: sum, sheet: "Sheet1", top_left: "A1", values: [[1, "=[src.xlsx]Dta!B2"]] });
    expect(noSheet.ok).toBe(false);
    expect(noSheet.links[0]).toMatchObject({ status: "sheet_not_found", sheets: ["Data", "Other"] });
    expect(noSheet.error).toContain("Data, Other");

    expect((await call("list_batches", { path: sum })).batches).toHaveLength(0);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBeNull();
    // Nothing was attached either: the workbook still saves as an ordinary one.
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", value: 1 });
    expect((await call("save_workbook", { path: sum })).saved).toBe(true);

    const name = await call("define_name", { path: sum, name: "Rev", ref: "[src.xlsx]Data!B2" });
    expect(name.ok).toBe(false);
    expect(name.error).toContain("defined name");
  });

  it("writes a link to the workbook itself as an ordinary reference", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src } = await pair(call, at);
    const w = await call("set_cell", { path: src, sheet: "Other", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.links[0].status).toBe("same_workbook");
    expect(w.cells[0].value).toBe(200);
    expect(w.cells[0].formula).not.toContain("[");
    expect((await call("save_workbook", { path: src })).saved).toBe(true);
  });

  it("refuses a numbered link in a file that has no link of that number", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { sum } = await pair(call, at);
    const w = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[1]Data!B2" });
    expect(w.ok).toBe(false);
    expect(w.links[0].status).toBe("workbook_not_found");
    expect(w.links[0].detail).toContain("no link [1]");
    expect((await call("list_batches", { path: sum })).batches).toHaveLength(0);
  });

  it("the copies of other workbooks' sheets are not sheets of the reader", async () => {
    const { handlers, at } = workspace();
    const h = handlers();
    const call = caller(h);
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });

    const d = await call("describe_workbook", { path: sum });
    expect(d.sheets.map((s: any) => s.name)).toEqual(["Sheet1"]);
    expect(d.linked_workbooks).toEqual([{ workbook: src, sheets: ["Data"] }]);
    expect((await call("describe_workbook", { path: src })).read_by).toEqual([sum]);
    expect(h.session(sum).sheetNames()).toEqual(["Sheet1"]);
    expect(Object.values(h.session(sum).snapshot().sheets).map((s: any) => s.name)).toEqual(["Sheet1"]);
    expect(diffForReview(h, sum).sheetNames).toEqual(["Sheet1"]);

    for (const [tool, args] of [
      ["read_range", { sheet: "[src.xlsx]Data", range: "A1:B3" }],
      ["describe_workbook", { sheet: "[src.xlsx]Data" }],
      ["find_rows", { query: "Revenue", sheet: "[src.xlsx]Data" }],
      ["query_rows", { sheet: "[src.xlsx]Data" }],
    ] as const) {
      const r = await call(tool, { path: sum, ...args });
      expect(r.ok, tool).toBe(false);
      expect(r.error, tool).toContain(src);
    }
    const found = await call("find_rows", { path: sum, query: "Revenue" });
    expect(found.matches).toEqual([]);

    for (const [tool, args] of [
      ["set_cell", { sheet: "[src.xlsx]Data", cell: "B1", value: 1 }],
      ["insert_rows", { sheet: "[src.xlsx]Data", before: 1 }],
      ["delete_sheet", { name: "[src.xlsx]Data" }],
      ["rename_sheet", { old_name: "[src.xlsx]Data", new_name: "Mine" }],
      ["create_sheet", { name: "[x.xlsx]Data" }],
    ] as const) {
      const r = await call(tool, { path: sum, ...args });
      expect(r.ok, tool).toBe(false);
    }
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    expect((await call("list_batches", { path: sum })).batches).toHaveLength(1);
  });

  it("saves a workbook that reads another: the link is in the file, and live again after a restart", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    const linked = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2+1" });
    expect(linked.note).toContain("Call save_workbook");

    const saved = await call("save_workbook", { path: sum });
    expect(saved, JSON.stringify(saved)).toMatchObject({ ok: true, saved: true, reads: [src] });
    expect(saved.note).toContain("src.xlsx");
    // The figures of src are not in its own file yet, and the save says so.
    expect(saved.note).toContain("not saved yet");
    expect(fs.existsSync(storePathFor(sum))).toBe(false);
    expect(linksIn(sum)).toEqual([{ index: 1, target: "src.xlsx", sheets: ["Data"] }]);
    expect((await call("save_workbook", { path: src })).saved).toBe(true);

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.value).toBe(201);
    expect(a1.formula).toContain("[src.xlsx]Data");
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toEqual([{ workbook: src, sheets: ["Data"] }]);
    expect((await call("describe_workbook", { path: src })).read_by).toEqual([sum]);
    expect((await call("list_batches", { path: sum })).batches).toHaveLength(0);
    expect(fs.existsSync(storePathFor(sum))).toBe(false);
    // Reopened from its file, the source is still followed.
    const e = await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 4 });
    expect(e.recalculated_workbooks).toEqual([sum]);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(9);
  });

  it("rejecting a batch of the workbook that is read takes the reader back with it", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    const ins = await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 3 });
    const edit = await call("set_cell", { path: src, sheet: "Data", cell: "B4", value: 9 });
    expect((await cell(call, sum, "Sheet1", "A1")).formula).toContain("B5");
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(18);

    // The reader's record of the row edit cannot be rejected by itself.
    const mirror = stored(sum).batches.find((b) => b.link_origin);
    expect(mirror.link_origin.batch_id).toBe(ins.batch_id);
    const direct = await call("reject_batch", { path: sum, batch_id: mirror.id });
    expect(direct.ok).toBe(false);
    expect(direct.error).toContain(ins.batch_id);

    const rej = await call("reject_batch", { path: src, batch_id: edit.batch_id });
    expect(rej.recalculated_workbooks).toEqual([sum]);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    expect((await cell(call, sum, "Sheet1", "A1")).formula).toContain("B5");

    await call("reject_batch", { path: src, batch_id: ins.batch_id });
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B2");
    expect(a1.value).toBe(200);
    expect(stored(sum).batches.filter((b) => b.link_origin)).toHaveLength(0);

    // Rejecting the reader's own link drops the copy with it: it saves again.
    const link = (await call("list_batches", { path: sum })).batches[0];
    await call("reject_batch", { path: sum, batch_id: link.id });
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toBeUndefined();
    expect((await call("describe_workbook", { path: src })).read_by).toBeUndefined();
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", value: 1 });
    expect((await call("save_workbook", { path: sum })).saved).toBe(true);
  });

  it("comes back the same in a new process", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    await call("insert_rows", { path: src, sheet: "Data", before: 2, count: 1 });
    await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 11 });

    // The reader first: it opens the workbook it reads by itself.
    await later();
    const h = handlers();
    call = caller(h);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B3");
    expect(a1.value).toBe(22);
    expect(h.openWorkbooks().map((w) => [path.basename(w.path), w.pending]).sort()).toEqual([
      ["src.xlsx", 3],
      ["sum.xlsx", 1],
    ]);
    await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 12 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(24);

    // The source first, edited before the reader is opened.
    call = await restart(handlers);
    await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 13 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(26);
    expect((await cell(call, sum, "Sheet1", "A1")).formula).toContain("B3");
  });

  it("catches up with rows inserted while the reader was not open", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });

    call = await restart(handlers);
    const ins = await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 4 });
    expect(ins.recalculated_workbooks).toBeUndefined();

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B6");
    expect(a1.value).toBe(200);
    expect(stored(sum).batches.filter((b) => b.link_origin)).toHaveLength(1);

    // And lets go of them when the insert is rejected while it is closed again.
    call = await restart(handlers);
    await call("reject_batch", { path: src, batch_id: ins.batch_id });
    call = await restart(handlers);
    const back = await cell(call, sum, "Sheet1", "A1");
    expect(back.formula).toContain("B2");
    expect(back.value).toBe(200);
    expect(stored(sum).batches.filter((b) => b.link_origin)).toHaveLength(0);
  });

  it("keeps following after the source is saved with its rows moved", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("save_workbook", { path: src });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 1 });
    expect((await call("save_workbook", { path: src })).saved).toBe(true);
    expect((await cell(call, sum, "Sheet1", "A1")).formula).toContain("B3");

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B3");
    expect(a1.value).toBe(200);
  });

  it("a reader that is not open takes up moved rows before the workbook it reads is saved", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("save_workbook", { path: src });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    expect(JSON.parse(fs.readFileSync(readersPathFor(src), "utf8")).readers).toEqual([sum]);

    // Only the source is opened here; the rows move and are saved into its file.
    call = await restart(handlers);
    await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 2 });
    await call("rename_sheet", { path: src, old_name: "Data", new_name: "Figures" });
    expect((await call("save_workbook", { path: src })).saved).toBe(true);
    expect(stored(sum).batches.filter((b) => b.link_origin)).toHaveLength(2);

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("[src.xlsx]Figures");
    expect(a1.formula).toContain("B4");
    expect(a1.value).toBe(200);
    await call("set_cell", { path: src, sheet: "Figures", cell: "B3", value: 9 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(18);

    // A workbook that stopped reading it is taken off the list at the next such save.
    const link = (await call("list_batches", { path: sum })).batches[0];
    await call("reject_batch", { path: sum, batch_id: link.id });
    call = await restart(handlers);
    await call("insert_rows", { path: src, sheet: "Figures", before: 1, count: 1 });
    await call("save_workbook", { path: src });
    expect(fs.existsSync(readersPathFor(src))).toBe(false);
  });

  it("follows a renamed sheet, and shows #REF! for a deleted one", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });

    const ren = await call("rename_sheet", { path: src, old_name: "Data", new_name: "Numbers 2026" });
    expect(ren.ok, JSON.stringify(ren)).toBe(true);
    let a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("[src.xlsx]Numbers 2026");
    expect(a1.value).toBe(200);
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toEqual([{ workbook: src, sheets: ["Numbers 2026"] }]);
    await call("set_cell", { path: src, sheet: "Numbers 2026", cell: "B1", value: 3 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(6);

    call = await restart(handlers);
    a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("[src.xlsx]Numbers 2026");
    expect(a1.value).toBe(6);
    await call("set_cell", { path: src, sheet: "Numbers 2026", cell: "B1", value: 5 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(10);

    const del = await call("delete_sheet", { path: src, name: "Numbers 2026" });
    expect(del.ok, JSON.stringify(del)).toBe(true);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe("#REF!");
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toBeUndefined();

    // Bringing the sheet back (rejecting the delete) brings the value back.
    await call("reject_batch", { path: src, batch_id: del.batch_id });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(10);
  });

  it("tells two files with the same name apart by folder, and links only one of them into a workbook", async () => {
    const { dir, handlers, at } = workspace();
    const call = caller(handlers());
    fs.mkdirSync(at("a"));
    fs.mkdirSync(at("b"));
    const sum = at("a/sum.xlsx");
    for (const [file, n] of [
      [at("a/src.xlsx"), 1],
      [at("b/src.xlsx"), 2],
      [at("b/only.xlsx"), 3],
    ] as const) {
      await call("create_workbook", { path: file, sheets: ["Data"] });
      await call("set_cell", { path: file, sheet: "Data", cell: "A1", value: n });
    }
    await call("create_workbook", { path: sum });

    // Next to the reader wins; a folder can be written, relative to the reader or absolute.
    const near = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!A1" });
    expect(near.cells[0].value).toBe(1);
    const far = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A2", formula: "='../b/[only.xlsx]Data'!A1" });
    expect(far.ok, JSON.stringify(far)).toBe(true);
    expect(far.cells[0].value).toBe(3);
    const abs = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A3", formula: `='${dir}/b/[only.xlsx]Data'!A1*2` });
    expect(abs.cells[0].value).toBe(6);
    // Open elsewhere and nowhere near: found among the open workbooks.
    const other = at("b/other.xlsx");
    await call("create_workbook", { path: other });
    const open = await call("set_cell", { path: other, sheet: "Sheet1", cell: "A1", formula: "=[sum.xlsx]Sheet1!A2" });
    expect(open.ok, JSON.stringify(open)).toBe(true);
    expect(open.cells[0].value).toBe(3);

    const second = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A4", formula: "='../b/[src.xlsx]Data'!A1" });
    expect(second.ok).toBe(false);
    expect(second.links[0].status).toBe("several_workbooks");

    const outside = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A4", formula: "='/etc/[src.xlsx]Data'!A1" });
    expect(outside.ok).toBe(false);
  });

  it("does not open a linked file outside the allowed folders", async () => {
    const { dir, at } = workspace();
    fs.mkdirSync(at("in"));
    fs.mkdirSync(at("out"));
    const all = caller(new ToolHandlers({ allowRoots: [dir] }));
    await all("create_workbook", { path: at("out/secret.xlsx"), sheets: ["Data"] });
    await all("set_cell", { path: at("out/secret.xlsx"), sheet: "Data", cell: "A1", value: 42 });
    await all("save_workbook", { path: at("out/secret.xlsx") });

    const boxed = caller(new ToolHandlers({ allowRoots: [at("in")] }));
    await boxed("create_workbook", { path: at("in/sum.xlsx") });
    const w = await boxed("set_cell", { path: at("in/sum.xlsx"), sheet: "Sheet1", cell: "A1", formula: "='../out/[secret.xlsx]Data'!A1" });
    expect(w.ok).toBe(false);
    expect(w.links[0].status).toBe("not_allowed");
    expect((await cell(boxed, at("in/sum.xlsx"), "Sheet1", "A1")).value).toBeNull();
  });

  it("lists the workbooks of a folder and creates new ones without replacing any", async () => {
    const { dir, handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    fs.writeFileSync(at("~$src.xlsx"), "lock");
    fs.writeFileSync(at("notes.txt"), "x");
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });

    const l = await call("list_workbooks", { path: dir });
    expect(l.workbooks.map((w: any) => w.name)).toEqual(["src.xlsx", "sum.xlsx"]);
    expect(l.workbooks[0]).toMatchObject({ path: src, open: true, pending_batches: 1, read_by: [sum] });
    expect(l.workbooks[1]).toMatchObject({ path: sum, open: true, pending_batches: 1, reads: [{ workbook: src, sheets: ["Data"] }] });
    const fresh = await caller(handlers())("list_workbooks", { path: dir });
    expect(fresh.workbooks.map((w: any) => w.open)).toEqual([undefined, undefined]);

    const again = await call("create_workbook", { path: src });
    expect(again.ok).toBe(false);
    expect(again.error).toContain("already exists");
    expect((await call("create_workbook", { path: at("new") })).path).toBe(at("new.xlsx"));
    expect((await call("create_workbook", { path: at("bad.xlsx"), sheets: ["a/b"] })).ok).toBe(false);
    expect((await call("create_workbook", { path: at("dup.xlsx"), sheets: ["A", "a"] })).ok).toBe(false);
    expect((await call("create_workbook", { path: "/etc/nope.xlsx" })).ok).toBe(false);
    expect((await call("describe_workbook", { path: at("new.xlsx") })).sheets.map((s: any) => s.name)).toEqual(["Sheet1"]);
  });

  it("carries text, booleans, errors and blanks, and clears what is cleared", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_range", { path: src, sheet: "Data", top_left: "D1", values: [["text", true, "=1/0"]] });
    const w = await call("set_range", {
      path: sum,
      sheet: "Sheet1",
      top_left: "A1",
      values: [
        [
          "=[src.xlsx]Data!D1",
          "=[src.xlsx]Data!E1",
          "=[src.xlsx]Data!F1",
          "=SUM([src.xlsx]Data!B1:B20)",
          "=COUNTA([src.xlsx]Data!A1:A20)",
          '=IF([src.xlsx]Data!Z99="","empty","full")',
          '=VLOOKUP("Costs",[src.xlsx]Data!A1:B20,2,FALSE)',
          '=SUMIFS([src.xlsx]Data!B1:B20,[src.xlsx]Data!A1:A20,"Units")',
        ],
      ],
    });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    const row = async () => (await call("read_range", { path: sum, sheet: "Sheet1", range: "A1:H1" })).cells.map((c: any) => c.value);
    expect(await row()).toEqual(["text", true, "#DIV/0!", 330, 3, "empty", 30, 100]);

    const cleared = await call("clear_range", { path: src, sheet: "Data", range: "A3:B3" });
    expect(cleared.ok, JSON.stringify(cleared)).toBe(true);
    await call("set_cell", { path: src, sheet: "Data", cell: "Z99", value: "x" });
    await call("set_cell", { path: src, sheet: "Data", cell: "F1", formula: "=1/4" });
    expect(await row()).toEqual(["text", true, 0.25, 300, 2, "full", "#N/A", 100]);

    // The same after a restart, where the copy is taken whole.
    const again = await restart(handlers);
    expect((await again("read_range", { path: sum, sheet: "Sheet1", range: "A1:H1" })).cells.map((c: any) => c.value)).toEqual([
      "text",
      true,
      0.25,
      300,
      2,
      "full",
      "#N/A",
      100,
    ]);
  });

  it("an error made in a saved workbook reaches its reader as that error, not as the number last saved", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    // Saved: the file now holds 200 for Revenue (=B1*2).
    expect((await call("save_workbook", { path: src })).saved).toBe(true);
    call = await restart(handlers);
    const w = await call("set_range", { path: sum, sheet: "Sheet1", top_left: "A1", values: [["=[src.xlsx]Data!B2", "=SUM([src.xlsx]Data!B1:B3)"]] });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.cells.map((c: any) => c.value)).toEqual([200, 330]);

    // The input becomes text: Revenue is #VALUE! in the source and in the reader.
    const text = await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: "n/a" });
    expect(text.ok, JSON.stringify(text)).toBe(true);
    expect((await cell(call, src, "Data", "B2")).value).toBe("#VALUE!");
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe("#VALUE!");
    expect((await cell(call, sum, "Sheet1", "B1")).value).toBe("#VALUE!");
    await call("reject_batch", { path: src, batch_id: text.batch_id });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);

    // The row the formula reads is deleted: #REF! in both.
    const del = await call("delete_rows", { path: src, sheet: "Data", start: 1, count: 1 });
    expect(del.ok, JSON.stringify(del)).toBe(true);
    expect((await cell(call, src, "Data", "B1")).value).toBe("#REF!");
    const check = async () => {
      const a1 = await cell(call, sum, "Sheet1", "A1");
      expect(a1.formula).toContain("B1");
      expect(a1.value).toBe("#REF!");
      expect((await cell(call, sum, "Sheet1", "B1")).value).toBe("#REF!");
    };
    await check();
    call = await restart(handlers);
    await check();
  });

  it("rows deleted in the workbook that is read: links below move up, links into them show #REF!", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_range", { path: sum, sheet: "Sheet1", top_left: "A1", values: [["=[src.xlsx]Data!B1", "=[src.xlsx]Data!B3", "=SUM([src.xlsx]Data!B1:B3)"]] });
    const del = await call("delete_rows", { path: src, sheet: "Data", start: 1, count: 1 });
    expect(del.ok, JSON.stringify(del)).toBe(true);
    const check = async () => {
      expect((await cell(call, sum, "Sheet1", "A1")).value).toBe("#REF!");
      const b1 = await cell(call, sum, "Sheet1", "B1");
      expect(b1.formula).toContain("B2");
      expect(b1.value).toBe(30);
      // The source's own B1 (=B1*2 moved up, its input gone) is what the copy shows.
      expect((await cell(call, sum, "Sheet1", "C1")).formula).toContain("B1:B2");
    };
    await check();
    call = await restart(handlers);
    await check();

    await call("reject_batch", { path: src, batch_id: del.batch_id });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(100);
    expect((await cell(call, sum, "Sheet1", "B1")).formula).toContain("B3");
    expect((await cell(call, sum, "Sheet1", "C1")).value).toBe(330);
  });

  it("one summary reads several workbooks by label, and keeps reading them as they change", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const summary = at("summary.xlsx");
    await call("create_workbook", { path: summary, sheets: ["Comps"] });
    const models: Array<[string, number, number]> = [
      ["ANT model.xlsx", 500, 120],
      ["NOW.xlsx", 900, 310],
      ["third.xlsx", 70, -5],
    ];
    for (const [file, revenue, income] of models) {
      await call("create_workbook", { path: at(file), sheets: ["Income Statement"] });
      await call("set_range", {
        path: at(file),
        sheet: "Income Statement",
        top_left: "A1",
        values: [
          ["", "FY2024", "FY2025"],
          ["Revenue", revenue / 2, revenue],
          ["Operating income", income / 2, income],
        ],
      });
    }
    const pull = (file: string, label: string) =>
      `=INDEX('[${file}]Income Statement'!$A$1:$H$40, MATCH("${label}", '[${file}]Income Statement'!$A$1:$A$40, 0), MATCH("FY2025", '[${file}]Income Statement'!$A$1:$H$1, 0))`;
    const w = await call("set_range", {
      path: summary,
      sheet: "Comps",
      top_left: "A1",
      values: [
        ["Company", "Revenue", "Operating income", "Margin"],
        ...models.map(([file], i) => [file.replace(/\..*$/, ""), pull(file, "Revenue"), pull(file, "Operating income"), `=C${i + 2}/B${i + 2}`]),
        ["Total", "=SUM(B2:B4)", "=SUM(C2:C4)", "=C5/B5"],
      ],
    });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.links.map((l: any) => l.status)).toEqual(["linked", "linked", "linked"]);
    const totals = async () => (await call("read_range", { path: summary, sheet: "Comps", range: "B5:C5" })).cells.map((c: any) => c.value);
    expect(await totals()).toEqual([1470, 425]);
    expect((await call("describe_workbook", { path: summary })).linked_workbooks).toHaveLength(3);

    // Each model changes its own way; the summary follows all of them.
    const e = await call("set_cell", { path: at("NOW.xlsx"), sheet: "Income Statement", cell: "C2", value: 1000 });
    expect(e.recalculated_workbooks).toEqual([summary]);
    await call("insert_rows", { path: at("ANT model.xlsx"), sheet: "Income Statement", before: 2, count: 3 });
    await call("insert_columns", { path: at("third.xlsx"), sheet: "Income Statement", before: "B", count: 2 });
    await call("set_cell", { path: at("ANT model.xlsx"), sheet: "Income Statement", cell: "C6", value: 150 });
    expect(await totals()).toEqual([1570, 455]);

    call = await restart(handlers);
    expect(await totals()).toEqual([1570, 455]);
    await call("set_cell", { path: at("third.xlsx"), sheet: "Income Statement", cell: "E2", value: 100 });
    expect(await totals()).toEqual([1600, 455]);
  });

  it("two workbooks read one; a workbook reads another both directly and through a third", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    const top = at("top.xlsx");
    await call("create_workbook", { path: top });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B1*10" });
    await call("set_range", { path: top, sheet: "Sheet1", top_left: "A1", values: [["=[src.xlsx]Data!B1", "=[sum.xlsx]Sheet1!A1", "=A1+B1"]] });
    expect((await cell(call, top, "Sheet1", "C1")).value).toBe(1100);
    const e = await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 2 });
    expect(e.recalculated_workbooks.sort()).toEqual([sum, top].sort());
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(20);
    expect((await cell(call, top, "Sheet1", "C1")).value).toBe(22);

    // A row inserted in the middle workbook moves the top one's link to it.
    await call("insert_rows", { path: sum, sheet: "Sheet1", before: 1, count: 2 });
    expect((await cell(call, top, "Sheet1", "B1")).formula).toContain("A3");
    await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 3 });
    expect((await cell(call, top, "Sheet1", "C1")).value).toBe(33);

    // Rejecting the bottom edit goes all the way up.
    const batches = (await call("list_batches", { path: src })).batches;
    const rej = await call("reject_batch", { path: src, batch_id: batches[batches.length - 1].id });
    expect(rej.recalculated_workbooks.sort()).toEqual([sum, top].sort());
    expect((await cell(call, top, "Sheet1", "C1")).value).toBe(22);
    expect((await cell(call, top, "Sheet1", "B1")).formula).toContain("A3");
  });

  it("links written by a script, and copied with their cell", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    const script = await call("run_script", {
      path: sum,
      script: 'const s = sheet("Sheet1"); for (let i = 1; i <= 3; i++) s.set("A" + i, "=[src.xlsx]data!B" + i);',
    });
    expect(script.ok, JSON.stringify(script)).toBe(true);
    expect(script.links).toEqual([{ link: "[src.xlsx]data", status: "linked", workbook: src, sheet: "Data" }]);
    expect((await call("read_range", { path: sum, sheet: "Sheet1", range: "A1:A3" })).cells.map((c: any) => c.value)).toEqual([100, 200, 30]);

    const copy = await call("copy_range", { path: sum, sheet: "Sheet1", source: "A1:A2", dest: "C2" });
    expect(copy.ok, JSON.stringify(copy)).toBe(true);
    const c3 = await cell(call, sum, "Sheet1", "C3");
    expect(c3.formula).toContain("[src.xlsx]Data");
    expect(c3.formula).toContain("D3");

    // A script cannot make a sheet named like a link, nor copy out of a copy.
    const bad = await call("run_script", { path: sum, script: 'sheet("[src.xlsx]Other").set("A1", 1);' });
    expect(bad.ok).toBe(false);
    const from = await call("copy_range", { path: sum, sheet: "[src.xlsx]Data", source: "A1:B3", dest_sheet: "Sheet1", dest: "F1" });
    expect(from.ok).toBe(false);
    expect(from.error).toContain(src);
  });

  it("rejecting single cells: of the workbook that is read, and of the link itself", async () => {
    const { handlers, at } = workspace();
    const h = handlers();
    const call = caller(h);
    const { src, sum } = await pair(call, at);
    await call("set_range", { path: sum, sheet: "Sheet1", top_left: "A1", values: [["=[src.xlsx]Data!B2", "=[src.xlsx]Other!A1+1"]] });
    const edit = await call("set_range", { path: src, sheet: "Data", top_left: "B1", values: [[7], [null], [8]] });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(14);
    const r = h.rejectCells(src, edit.batch_id, [{ sheet: "Data", row: 0, col: 1 }]) as any;
    expect(r.ok, JSON.stringify(r)).toBe(true);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    expect((await cell(call, src, "Data", "B3")).value).toBe(8);

    // Reject one of the two linking cells: that link is no longer read.
    const link = (await call("list_batches", { path: sum })).batches[0];
    h.rejectCells(sum, link.id, [{ sheet: "Sheet1", row: 0, col: 1 }]);
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toEqual([{ workbook: src, sheets: ["Data"] }]);
    expect(stored(sum).batches[0].links).toEqual([{ name: "[src.xlsx]Data", source: src, sheet: "Data" }]);
    // And the other: nothing is read, and the workbook saves.
    h.rejectCells(sum, link.id, [{ sheet: "Sheet1", row: 0, col: 0 }]);
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toBeUndefined();
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", value: 5 });
    expect((await call("save_workbook", { path: sum })).saved).toBe(true);
  });

  it("a link replaced by a plain value saves, without the other workbook's rows", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_range", { path: sum, sheet: "Sheet1", top_left: "A1", values: [["=[src.xlsx]Data!B2", "kept"]] });
    await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 2 });
    await call("rename_sheet", { path: src, old_name: "Data", new_name: "Renamed" });
    expect((await call("save_workbook", { path: sum })).error).toBe("save_source_first");

    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", value: 200 });
    const saved = await call("save_workbook", { path: sum });
    expect(saved.saved, JSON.stringify(saved)).toBe(true);
    expect(fs.existsSync(storePathFor(sum))).toBe(false);

    call = await restart(handlers);
    const d = await call("describe_workbook", { path: sum });
    expect(d.sheets.map((s: any) => s.name)).toEqual(["Sheet1"]);
    expect(d.linked_workbooks).toBeUndefined();
    expect((await call("read_range", { path: sum, sheet: "Sheet1", range: "A1:B1" })).cells.map((c: any) => c.value)).toEqual([200, "kept"]);
    expect((await call("list_batches", { path: sum })).batches).toHaveLength(0);
  });

  it("a saved copy of the workbook that is read leaves the link on the original", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    const copy = await call("save_workbook", { path: src, as: at("src copy.xlsx") });
    expect(copy.saved, JSON.stringify(copy)).toBe(true);
    await call("set_cell", { path: at("src copy.xlsx"), sheet: "Data", cell: "B1", value: 1 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 2 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(4);
  });

  it("shows #REF! while the file it reads is away, and reads it again when it is back", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("save_workbook", { path: src });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });

    fs.renameSync(src, at("moved.xlsx"));
    call = await restart(handlers);
    const away = await cell(call, sum, "Sheet1", "A1");
    expect(away.value).toBe("#REF!");
    expect(away.formula).toContain("[src.xlsx]Data");
    expect(stored(sum).batches[0].links).toHaveLength(1);

    fs.renameSync(at("moved.xlsx"), src);
    call = await restart(handlers);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 6 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(12);
  });

  it("the review of a workbook shows its own changes, with the values its links read", async () => {
    const { handlers, at } = workspace();
    const h = handlers();
    const call = caller(h);
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 1 });
    const diff = diffForReview(h, sum) as any;
    expect(diff.batches).toHaveLength(1);
    expect(JSON.stringify(diff)).not.toContain("linked workbook");
    expect(h.openWorkbooks().find((w) => w.path === sum)!.pending).toBe(1);
    expect(h.pendingBatches(sum).map((b) => b.prompt)).toEqual(["set_cell"]);
  });

  it("works through the MCP server: two workbooks made, linked, and edited over the protocol", async () => {
    const { dir, at } = workspace();
    const { server } = createServer({ allowRoots: [dir] });
    const [a, b] = InMemoryTransport.createLinkedPair();
    await server.connect(a);
    const client = new Client({ name: "test", version: "0" });
    await client.connect(b);
    const call: Call = async (name, args) => JSON.parse(((await client.callTool({ name, arguments: args })) as any).content[0].text);

    const names = (await client.listTools()).tools.map((t) => t.name);
    expect(names).toEqual(expect.arrayContaining(["list_workbooks", "create_workbook"]));
    expect(await call("list_workbooks", { path: dir })).toMatchObject({ ok: true, workbooks: [] });

    const { src, sum } = await pair(call, at);
    const listed = await call("list_workbooks", { path: dir });
    expect(listed.workbooks.map((w: any) => path.basename(w.path)).sort()).toEqual(["src.xlsx", "sum.xlsx"]);

    const w = await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.links).toMatchObject([{ link: "[src.xlsx]Data", status: "linked" }]);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);

    const e = await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 7 });
    expect(e.recalculated_workbooks).toEqual([sum]);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(14);

    const saved = await call("save_workbook", { path: sum });
    expect(saved).toMatchObject({ ok: true, saved: true, reads: [src] });
    expect(linksIn(sum)).toEqual([{ index: 1, target: "src.xlsx", sheets: ["Data"] }]);
    await client.close();
  });
});

describe.skipIf(!linking)("links written into the file", () => {
  /** `src.xlsx` saved, `sum.xlsx` saved with A1 reading it, and a new process. */
  async function savedPair(handlers: () => ToolHandlers, at: (n: string) => string, formula = "=[src.xlsx]Data!B2") {
    const first = caller(handlers());
    const { src, sum } = await pair(first, at);
    expect((await first("save_workbook", { path: src })).saved).toBe(true);
    expect((await first("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula })).ok).toBe(true);
    const saved = await first("save_workbook", { path: sum });
    expect(saved.saved, JSON.stringify(saved)).toBe(true);
    return { src, sum, call: await restart(handlers) };
  }

  it("is not saved while the workbook it reads has unsaved row edits; that one is saved first", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 2 });
    const before = fs.readFileSync(sum);

    const refused = await call("save_workbook", { path: sum });
    expect(refused).toMatchObject({ ok: false, error: "save_source_first", save_first: [src] });
    expect(refused.reason).toContain("src.xlsx");
    expect(fs.readFileSync(sum).equals(before)).toBe(true);
    // A copy would point at the same rows.
    expect((await call("save_workbook", { path: sum, as: at("copy.xlsx") })).error).toBe("save_source_first");
    expect(fs.existsSync(at("copy.xlsx"))).toBe(false);
    expect((await call("list_batches", { path: sum })).batches).toHaveLength(1);

    expect((await call("save_workbook", { path: src })).saved).toBe(true);
    const saved = await call("save_workbook", { path: sum });
    expect(saved.saved, JSON.stringify(saved)).toBe(true);

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B4");
    expect(a1.value).toBe(200);
  });

  it("a new sheet of the workbook that is read has to be in its file first", async () => {
    const { handlers, at } = workspace();
    const call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("save_workbook", { path: src });
    await call("create_sheet", { path: src, name: "Fresh" });
    await call("set_cell", { path: src, sheet: "Fresh", cell: "A1", value: 3 });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Fresh!A1" });
    expect(await call("save_workbook", { path: sum })).toMatchObject({ ok: false, error: "save_source_first", save_first: [src] });
    await call("save_workbook", { path: src });
    expect((await call("save_workbook", { path: sum })).saved).toBe(true);
    expect(linksIn(sum)).toEqual([{ index: 1, target: "src.xlsx", sheets: ["Fresh"] }]);
  });

  it("a second save keeps the file's link number, adds a sheet to it, and numbers a new workbook after it", async () => {
    const { handlers, at } = workspace();
    let { src, sum, call } = await savedPair(handlers, at);
    const third = at("third.xlsx");
    await call("create_workbook", { path: third });
    await call("set_cell", { path: third, sheet: "Sheet1", cell: "A1", value: 7 });
    await call("set_cell", { path: src, sheet: "Other", cell: "A1", value: 5 });
    // `[1]` is the file's own name for src.xlsx, and is read as that.
    const w = await call("set_range", { path: sum, sheet: "Sheet1", top_left: "B1", values: [["=[src.xlsx]Other!A1", "=[third.xlsx]Sheet1!A1", "=[1]Data!B3"]] });
    expect(w.ok, JSON.stringify(w)).toBe(true);
    expect(w.cells.map((c: any) => c.value)).toEqual([5, 7, 30]);
    expect(w.links.map((l: any) => [l.link, l.status, l.workbook])).toEqual([
      ["[src.xlsx]Other", "linked", src],
      ["[third.xlsx]Sheet1", "linked", third],
      ["[1]Data", "linked", src],
    ]);
    const saved = await call("save_workbook", { path: sum });
    expect(saved.saved, JSON.stringify(saved)).toBe(true);
    expect(linksIn(sum)).toEqual([
      { index: 1, target: "src.xlsx", sheets: ["Data", "Other"] },
      { index: 2, target: "third.xlsx", sheets: ["Sheet1"] },
    ]);

    call = await restart(handlers);
    expect((await call("read_range", { path: sum, sheet: "Sheet1", range: "A1:D1" })).cells.map((c: any) => c.value)).toEqual([200, 5, 7, 30]);
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toEqual([
      { workbook: src, sheets: ["Data", "Other"] },
      { workbook: third, sheets: ["Sheet1"] },
    ]);
  });

  it("a save after reopening writes the cells that changed and no others", async () => {
    const { handlers, at } = workspace();
    const first = caller(handlers());
    const { src, sum } = await pair(first, at);
    await first("save_workbook", { path: src });
    await first("set_range", { path: sum, sheet: "Sheet1", top_left: "A1", values: [["=[src.xlsx]Data!B2"], ["=SUM('[src.xlsx]Data'!B1:B3)"]] });
    expect((await first("save_workbook", { path: sum })).saved).toBe(true);
    await later();

    const h = handlers();
    const call = caller(h);
    expect((await call("read_range", { path: sum, sheet: "Sheet1", range: "A1:A2" })).cells.map((c: any) => c.value)).toEqual([200, 330]);
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "C1", value: 1 });
    const plan = (h as any).workspace.savePlan(sum);
    expect(plan.ok).toBe(true);
    const built = h.session(sum).buildSave(plan.links) as any;
    expect(built.ok, JSON.stringify(built.reason)).toBe(true);
    expect(built.patch.sheets.flatMap((s: any) => s.cells.map((c: any) => [s.name, c.r, c.c]))).toEqual([["Sheet1", 0, 2]]);
    // The values read are kept up to date under the number the file already has.
    expect(built.patch.externalLinks).toHaveLength(1);
    expect(built.patch.externalLinks[0]).toMatchObject({ index: 1, sheets: [{ name: "Data" }] });
    expect(built.patch.externalLinks[0].target).toBeUndefined();
    const kept = built.patch.externalLinks[0].sheets[0].cells.map((c: any) => [c.r, c.c]).sort();
    expect(kept).toEqual([
      [0, 1],
      [1, 1],
      [2, 1],
    ]);
  });

  it("rows inserted in the workbook that is read: the reader's file follows once both are saved", async () => {
    const { handlers, at } = workspace();
    let { src, sum, call } = await savedPair(handlers, at);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 2 });
    expect((await cell(call, sum, "Sheet1", "A1")).formula).toContain("B4");
    // Nothing was written in the reader, and its file is behind all the same.
    expect((await call("list_batches", { path: sum })).batches).toHaveLength(0);
    expect((await call("save_workbook", { path: sum })).error).toBe("save_source_first");
    const saved = await call("save_workbook", { path: src });
    expect(saved, JSON.stringify(saved)).toMatchObject({ saved: true, also_save: [sum] });

    // Left unsaved, the reader still knows where its rows went.
    call = await restart(handlers);
    const kept = await cell(call, sum, "Sheet1", "A1");
    expect(kept.formula).toContain("B4");
    expect(kept.value).toBe(200);
    const again = await call("save_workbook", { path: sum });
    expect(again.saved, JSON.stringify(again)).toBe(true);
    expect(fs.existsSync(storePathFor(sum))).toBe(false);

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B4");
    expect(a1.value).toBe(200);
    expect((await call("save_workbook", { path: sum })).note).toBe("nothing pending");
  });

  it("a reader that is not open takes up the moved rows when the workbook it reads is saved", async () => {
    const { handlers, at } = workspace();
    let { src, sum, call } = await savedPair(handlers, at);
    // Only src is touched in this process.
    await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 1 });
    const saved = await call("save_workbook", { path: src });
    expect(saved, JSON.stringify(saved)).toMatchObject({ saved: true, also_save: [sum] });

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B3");
    expect(a1.value).toBe(200);
  });

  it("a sheet renamed in the workbook that is read: the reader's file takes the new name", async () => {
    const { handlers, at } = workspace();
    let { src, sum, call } = await savedPair(handlers, at);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    await call("rename_sheet", { path: src, old_name: "Data", new_name: "Figures 2026" });
    expect((await call("save_workbook", { path: sum })).error).toBe("save_source_first");
    expect((await call("save_workbook", { path: src })).also_save).toEqual([sum]);
    expect((await call("save_workbook", { path: sum })).saved).toBe(true);

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("[src.xlsx]Figures 2026");
    expect(a1.value).toBe(200);
    await call("set_cell", { path: src, sheet: "Figures 2026", cell: "B1", value: 1 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(2);
  });

  it("rejecting rows inserted in the workbook that is read takes a reader opened from its file back too", async () => {
    const { handlers, at } = workspace();
    const { src, sum, call } = await savedPair(handlers, at);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    const ins = await call("insert_rows", { path: src, sheet: "Data", before: 1, count: 3 });
    expect((await cell(call, sum, "Sheet1", "A1")).formula).toContain("B5");
    expect((await call("reject_batch", { path: src, batch_id: ins.batch_id })).ok).toBe(true);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("B2");
    expect(a1.value).toBe(200);
    // Nothing is left to save in either.
    expect((await call("save_workbook", { path: sum })).note).toBe("nothing pending");
  });

  it("names with spaces: the file and the sheet", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at, "my src.xlsx", "Income Statement");
    await call("save_workbook", { path: src });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "='[my src.xlsx]Income Statement'!B2" });
    expect((await call("save_workbook", { path: sum })).saved).toBe(true);
    expect(linksIn(sum)).toEqual([{ index: 1, target: "my%20src.xlsx", sheets: ["Income Statement"] }]);

    call = await restart(handlers);
    const a1 = await cell(call, sum, "Sheet1", "A1");
    expect(a1.formula).toContain("[my src.xlsx]Income Statement");
    expect(a1.value).toBe(200);
    await call("set_cell", { path: src, sheet: "Income Statement", cell: "B1", value: 3 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(6);
  });

  it("a copy saved next to it or elsewhere finds the workbook it reads", async () => {
    const { handlers, at } = workspace();
    let call = caller(handlers());
    const { src, sum } = await pair(call, at);
    await call("save_workbook", { path: src });
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    const near = at("near.xlsx");
    const far = at(path.join("sub", "far.xlsx"));
    expect((await call("save_workbook", { path: sum, as: near })).saved).toBe(true);
    expect((await call("save_workbook", { path: sum, as: far })).saved).toBe(true);
    expect(linksIn(near)[0].target).toBe("src.xlsx");
    expect(linksIn(far)[0].target).toBe(`file://${src.split(path.sep).join("/")}`);
    // The original is as it was: its link is still pending.
    expect((await call("list_batches", { path: sum })).batches).toHaveLength(1);

    call = await restart(handlers);
    expect((await cell(call, near, "Sheet1", "A1")).value).toBe(200);
    expect((await cell(call, far, "Sheet1", "A1")).value).toBe(200);
    const e = await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 1 });
    expect([...e.recalculated_workbooks].sort()).toEqual([far, near].sort());
    expect((await cell(call, far, "Sheet1", "A1")).value).toBe(2);
  });

  it("how a file names the workbook it reads", () => {
    const root = path.resolve("/models");
    const p = (...parts: string[]) => path.join(root, ...parts);
    expect(linkTarget(p("NOW.xlsx"), root)).toBe("NOW.xlsx");
    expect(linkTarget(p("peers", "NOW 2026 #1.xlsx"), root)).toBe("peers/NOW%202026%20%231.xlsx");
    expect(linkTarget(p("NOW.xlsx"), p("mine"))).toMatch(/^file:\/\/\/.*models\/NOW\.xlsx$/);
  });

  it("keeps the saved value while the file it reads is away, and reads it again when it is back", async () => {
    const { handlers, at } = workspace();
    let { src, sum, call } = await savedPair(handlers, at);
    fs.renameSync(src, at("moved.xlsx"));
    call = await restart(handlers);
    const away = await cell(call, sum, "Sheet1", "A1");
    expect(away.value).toBe(200);
    expect(away.formula).toContain("[1]Data");
    expect((await call("describe_workbook", { path: sum })).linked_workbooks).toBeUndefined();
    // It is edited and saved all the same; the link stays in the file as it was.
    await call("set_cell", { path: sum, sheet: "Sheet1", cell: "C1", value: 1 });
    const saved = await call("save_workbook", { path: sum });
    expect(saved.saved, JSON.stringify(saved)).toBe(true);
    expect(linksIn(sum)).toEqual([{ index: 1, target: "src.xlsx", sheets: ["Data"] }]);

    fs.renameSync(at("moved.xlsx"), src);
    call = await restart(handlers);
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(200);
    await call("set_cell", { path: src, sheet: "Data", cell: "B1", value: 6 });
    expect((await cell(call, sum, "Sheet1", "A1")).value).toBe(12);
    expect((await cell(call, sum, "Sheet1", "C1")).value).toBe(1);
  });

  it("does not open a workbook outside the allowed folders because a file names it", async () => {
    const { handlers, at, dir } = workspace();
    const first = caller(handlers());
    const { src, sum } = await pair(first, at);
    await first("save_workbook", { path: src });
    await first("set_cell", { path: sum, sheet: "Sheet1", cell: "A1", formula: "=[src.xlsx]Data!B2" });
    // A copy one folder down: its link names src.xlsx by its whole path.
    const inner = path.join(dir, "inner");
    const copy = path.join(inner, "sum.xlsx");
    expect((await first("save_workbook", { path: sum, as: copy })).saved).toBe(true);
    expect(linksIn(copy)[0].target).toContain("/src.xlsx");

    // Allowed to open both, the link is live.
    const both = await restart(handlers);
    expect((await cell(both, copy, "Sheet1", "A1")).formula).toContain("[src.xlsx]Data");

    // Allowed only the copy's folder, it shows what the file was saved with and src.xlsx stays closed.
    await later();
    const h = new ToolHandlers({ allowRoots: [inner] });
    const a1 = await cell(caller(h), copy, "Sheet1", "A1");
    expect(a1.value).toBe(200);
    expect(a1.formula).toContain("[1]Data");
    expect(h.openWorkbooks().map((w) => w.path)).toEqual([copy]);
  });
});
