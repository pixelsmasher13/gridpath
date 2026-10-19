/**
 * Review server + on-disk batch store: a write via the handlers shows up in
 * the API with before/after, persists across a fresh handler instance,
 * rejects via the API, and saves from the review path even when the MCP
 * side is --review-required.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { ToolHandlers } from "../tools/handlers";
import { startReviewServer, type ReviewServer } from "./api";
import { loadEngine } from "../workbook/session";
import { storePathFor } from "../workbook/store";

let built = true;
try {
  loadEngine();
} catch {
  built = false;
}
const FIXTURE = path.resolve(import.meta.dirname, "../../../../eval/fixtures/rich-model.xlsx");

describe.skipIf(!built)("review api + store", () => {
  let dir: string;
  let file: string;
  let handlers: ToolHandlers;
  let srv: ReviewServer;
  const api = (p: string, init?: RequestInit) => fetch(`${srv.url}${p}${p.includes("?") ? "&" : "?"}token=${srv.token}`, init).then((r) => r.json() as Promise<any>);

  beforeAll(async () => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), "gridpath-review-"));
    file = path.join(dir, "model.xlsx");
    fs.copyFileSync(FIXTURE, file);
    let base: { url: string; token: string } | null = null;
    handlers = new ToolHandlers({ allowRoots: [dir], reviewRequired: true, reviewUrl: () => base });
    srv = await startReviewServer(handlers);
    base = { url: srv.url, token: srv.token };
  });
  afterAll(async () => {
    await srv.close();
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it("rejects requests without the token", async () => {
    const r = await fetch(`${srv.url}/api/workbooks`);
    expect(r.status).toBe(401);
  });

  it("a write is pending, persisted, and visible with before/after", async () => {
    const w: any = await handlers.call("set_cell", { path: file, sheet: "Model", cell: "B7", value: 5000 });
    expect(w.ok).toBe(true);
    expect(w.review_url).toMatch(/^http:\/\/127\.0\.0\.1:\d+\/review\?path=.+&token=[0-9a-f]+$/);
    expect(fs.existsSync(storePathFor(file))).toBe(true);

    const wb = await api(`/api/workbook?path=${encodeURIComponent(file)}`);
    expect(wb.batches.length).toBe(1);
    expect(wb.batches[0].cells[0]).toMatchObject({ sheet: "Model", row: 6, col: 1, before: 5155, after: 5000 });
    expect(wb.snapshot.sheetOrder.length).toBe(2);
  });

  it("save_workbook is gated; the store survives a new handler instance", async () => {
    const s: any = await handlers.call("save_workbook", { path: file });
    expect(s).toMatchObject({ ok: true, saved: false, pending_batches: 1 });
    expect(fs.readFileSync(file).equals(fs.readFileSync(FIXTURE))).toBe(true);

    const again = new ToolHandlers({ allowRoots: [dir] });
    const lb: any = await again.call("list_batches", { path: file });
    expect(lb.batches.length).toBe(1);
    expect(again.session(file).cell("Model", 6, 1).value).toBe(5000);
  });

  it("reject via the API drops the batch and clears the store", async () => {
    const r = await api("/api/batch/reject", { method: "POST", body: JSON.stringify({ path: file, batch_id: "b1" }) });
    expect(r).toMatchObject({ ok: true, remaining: 0 });
    expect(fs.existsSync(storePathFor(file))).toBe(false);
    expect(handlers.session(file).cell("Model", 6, 1).value).toBe(5155);
  });

  it("save from the review path writes the file and clears pending", async () => {
    await handlers.call("set_cell", { path: file, sheet: "Model", cell: "B7", value: 5001 });
    const r = await api("/api/save", { method: "POST", body: JSON.stringify({ path: file }) });
    expect(r).toMatchObject({ ok: true, saved: true });
    expect(fs.existsSync(storePathFor(file))).toBe(false);
    const fresh = new ToolHandlers({ allowRoots: [dir] });
    expect(fresh.session(file).cell("Model", 6, 1).value).toBe(5001);
    expect(fresh.session(file).batches.length).toBe(0);
  });
});

describe.skipIf(!built)("save as", () => {
  it("writes a patched copy and leaves the original and its batches alone", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "gridpath-saveas-"));
    const file = path.join(dir, "model.xlsx"); fs.copyFileSync(FIXTURE, file);
    const h = new ToolHandlers({ allowRoots: [dir], reviewRequired: true });
    await h.call("set_cell", { path: file, sheet: "Model", cell: "B7", value: 4321 });
    const r: any = await h.call("save_workbook", { path: file, as: path.join(dir, "model (edited).xlsx") });
    expect(r).toMatchObject({ ok: true, saved: true, copy: true });
    expect(fs.readFileSync(file).equals(fs.readFileSync(FIXTURE))).toBe(true);
    expect(h.session(file).batches.length).toBe(1);
    const copy = new ToolHandlers({ allowRoots: [dir] });
    expect(copy.session(path.join(dir, "model (edited).xlsx")).cell("Model", 6, 1).value).toBe(4321);
    fs.rmSync(dir, { recursive: true, force: true });
  });
});

describe.skipIf(!built)("reject cells", () => {
  it("drops one cell from a batch and replays the rest", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "gridpath-cells-"));
    const file = path.join(dir, "model.xlsx"); fs.copyFileSync(FIXTURE, file);
    const h = new ToolHandlers({ allowRoots: [dir] });
    await h.call("set_range", { path: file, sheet: "Model", top_left: "A20", values: [["x", 1, 2]] });
    const r: any = h.rejectCells(file, "b1", [{ sheet: "Model", row: 19, col: 1 }]);
    expect(r).toMatchObject({ ok: true, rejected_cells: 1, batches: 1 });
    const s = h.session(file);
    expect(s.cell("Model", 19, 0).value).toBe("x");
    expect(s.cell("Model", 19, 1).value).toBe(null);
    expect(s.cell("Model", 19, 2).value).toBe(2);
    fs.rmSync(dir, { recursive: true, force: true });
  });
});
