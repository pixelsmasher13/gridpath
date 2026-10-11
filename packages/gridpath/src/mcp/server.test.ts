/**
 * End-to-end through the MCP protocol with an in-memory transport: list
 * tools, describe, write, run a script, save — and the saved file keeps
 * every untouched part byte-identical.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, it, expect, beforeAll, afterAll } from "vitest";
import JSZip from "jszip";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "./server";
import { loadEngine } from "../workbook/session";

let built = true;
try {
  loadEngine();
} catch {
  built = false;
}

const FIXTURE = path.resolve(import.meta.dirname, "../../../../eval/fixtures/rich-model.xlsx");

async function parts(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, Uint8Array>();
  for (const [n, f] of Object.entries(zip.files)) if (!f.dir) out.set(n, await f.async("uint8array"));
  return out;
}
const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((x, i) => x === b[i]);

describe.skipIf(!built)("mcp server", () => {
  let dir: string;
  let file: string;
  let client: Client;
  const text = (r: any) => JSON.parse(r.content[0].text);

  beforeAll(async () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "gridpath-mcp-"));
    file = path.join(dir, "model.xlsx");
    fs.copyFileSync(FIXTURE, file);
    const { server } = createServer({ allowRoots: [dir] });
    const [a, b] = InMemoryTransport.createLinkedPair();
    await server.connect(a);
    client = new Client({ name: "test", version: "0" });
    await client.connect(b);
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  it("lists tools with path added to every schema", async () => {
    const { tools } = await client.listTools();
    const names = tools.map((t) => t.name);
    expect(names).toEqual(expect.arrayContaining(["describe_workbook", "read_range", "set_cell", "run_script", "save_workbook"]));
    expect(names).not.toContain("done");
    for (const t of tools) expect((t.inputSchema as any).required).toContain("path");
  });

  it("refuses paths outside the allowed roots", async () => {
    const r = text(await client.callTool({ name: "describe_workbook", arguments: { path: FIXTURE } }));
    expect(r.ok).toBe(false);
    expect(r.error).toMatch(/outside the allowed folders/);
    expect(r.error).toMatch(/Do NOT copy/);
  });

  it("query_rows filters, projects and groups rows through the engine", async () => {
    const described = text(await client.callTool({ name: "describe_workbook", arguments: { path: file } }));
    const sheetEntry = described.sheets.find((s: any) => s.header_row !== null) ?? described.sheets[0];
    const sheet: string = sheetEntry.name;
    // Every row, two columns, small limit: rows carry 1-indexed numbers and
    // cells keyed by column letter; `matched` counts beyond the limit.
    const all = text(await client.callTool({ name: "query_rows", arguments: { path: file, sheet, columns: ["A", "B"], limit: 3 } }));
    expect(all.sheet).toBe(sheet);
    expect(all.rows.length).toBeLessThanOrEqual(3);
    expect(all.rows.length).toBeGreaterThan(0);
    expect(all.matched).toBeGreaterThanOrEqual(all.rows.length);
    expect(all.columns.map((c: any) => c.column)).toEqual(["A", "B"]);
    for (const r of all.rows) {
      expect(typeof r.row).toBe("number");
      expect(Object.keys(r.cells).sort()).toEqual(["A", "B"]);
    }
    // A predicate on column A that nothing satisfies.
    const none = text(await client.callTool({ name: "query_rows", arguments: { path: file, sheet, where: [{ column: "A", op: "eq", value: "no such label 1234" }] } }));
    expect(none.matched).toBe(0);
    expect(none.rows).toEqual([]);
    // not_empty on A, grouped by A: every group has count ≥ 1.
    const grouped = text(await client.callTool({ name: "query_rows", arguments: { path: file, sheet, where: [{ column: "A", op: "not_empty" }], group_by: "A", limit: 5 } }));
    expect(grouped.group_by.column).toBe("A");
    expect(grouped.groups.length).toBeGreaterThan(0);
    for (const g of grouped.groups) expect(g.count).toBeGreaterThanOrEqual(1);
    // Errors are readable, not thrown.
    const bad = text(await client.callTool({ name: "query_rows", arguments: { path: file, sheet: "Nope" } }));
    expect(bad.ok).toBe(false);
    expect(bad.error).toMatch(/Sheet not found/);
    const badOp = text(await client.callTool({ name: "query_rows", arguments: { path: file, sheet, where: [{ column: "A", op: "like", value: "x" }] } }));
    expect(badOp.ok).toBe(false);
    expect(badOp.error).toMatch(/unknown op/);
  });

  it("describe → write → readback → script → save keeps untouched parts identical", async () => {
    const described = text(await client.callTool({ name: "describe_workbook", arguments: { path: file } }));
    const sheet: string = described.sheets?.[0]?.name ?? described.sheet_names?.[0];
    expect(sheet).toBeTruthy();

    const w = text(await client.callTool({ name: "set_cell", arguments: { path: file, sheet, cell: "AO1", value: 41 } }));
    expect(w).toMatchObject({ ok: true });
    expect(w.write_check).toBe("ok");
    expect(w.cells[0]).toMatchObject({ cell: "AO1", value: 41 });

    const sc = text(
      await client.callTool({
        name: "run_script",
        arguments: { path: file, script: `const s = sheet("${sheet}"); s.set("AO2", "=AO1+1"); log("done");` },
      }),
    );
    expect(sc.ok).toBe(true);
    expect(sc.script_logs).toEqual(["done"]);
    expect(sc.cells[0]).toMatchObject({ cell: "AO2", value: 42, formula: "=AO1+1" });

    const lb = text(await client.callTool({ name: "list_batches", arguments: { path: file } }));
    expect(lb.batches.length).toBe(2);

    const before = await parts(fs.readFileSync(FIXTURE));
    const saved = text(await client.callTool({ name: "save_workbook", arguments: { path: file } }));
    expect(saved).toMatchObject({ ok: true, saved: true });
    const after = await parts(fs.readFileSync(file));
    // Formulas are saved as dynamic-array formulas, which adds the one
    // workbook-wide metadata part (and registers it in the content types
    // and workbook relationships); everything else keeps its identity.
    expect([...after.keys()].sort()).toEqual([...before.keys(), "xl/metadata.xml"].sort());
    const changed = [...before.keys()].filter((k) => !same(before.get(k)!, after.get(k)!));
    const allowed = new Set(["xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"]);
    expect(changed.every((k) => allowed.has(k) || /^xl\/worksheets\/sheet\d+\.xml$/.test(k))).toBe(true);

    // The saved file is the new baseline: reading it back shows the edit.
    const rr = text(await client.callTool({ name: "read_range", arguments: { path: file, sheet, range: "AO1:AO2" } }));
    expect(rr.cells.map((c: any) => c.value)).toEqual([41, 42]);
  });

  it("script timeouts are reported, not hung", async () => {
    const r = text(await client.callTool({ name: "run_script", arguments: { path: file, script: "while (true) {}" } }));
    expect(r.ok).toBe(false);
    expect(r.error).toMatch(/timed out/);
  }, 15_000);
});
