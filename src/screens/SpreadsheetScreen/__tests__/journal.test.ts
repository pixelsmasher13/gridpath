import { describe, it, expect, vi } from "vitest";

// The journal module talks to Tauri only through `invoke`; stub it so the
// pure parts (recording, replay planning, coercion) run in Node.
const invokeCalls: Array<{ cmd: string; args: any }> = [];
vi.mock("@tauri-apps/api/core", () => ({
  invoke: async (cmd: string, args: any) => {
    invokeCalls.push({ cmd, args });
    if (cmd === "workbook_file_stamp") return { size: 123, mtime_ms: 456 };
    if (cmd === "journal_read") return null;
    return null;
  },
}));

const {
  coerceInput,
  describeJournal,
  emptyManual,
  journalClear,
  journalFlush,
  journalName,
  journalOpen,
  journalPeek,
  journalRecordEdit,
  replayJournal,
} = await import("../journal");

function fakeGrid() {
  const calls: string[] = [];
  const rec = (name: string) => (...a: unknown[]) => { calls.push(`${name}(${a.map((x) => JSON.stringify(x)).join(",")})`); return true; };
  const grid: any = {
    setCells: rec("setCells"),
    insertRows: rec("insertRows"),
    deleteRows: rec("deleteRows"),
    insertColumns: rec("insertColumns"),
    deleteColumns: rec("deleteColumns"),
    createSheet: rec("createSheet"),
    deleteSheet: rec("deleteSheet"),
    renameSheet: rec("renameSheet"),
    setCellFormat: rec("setCellFormat"),
    setCellBackground: rec("setCellBackground"),
    setCellBorders: rec("setCellBorders"),
    mergeCells: rec("mergeCells"),
    unmergeCells: rec("unmergeCells"),
    freezePanes: rec("freezePanes"),
    hideRows: rec("hideRows"),
    showRows: rec("showRows"),
    hideColumns: rec("hideColumns"),
    showColumns: rec("showColumns"),
    setRangeFormat: rec("setRangeFormat"),
  };
  return { grid, calls };
}

describe("journal", () => {
  it("coerces typed input the way the grid stores it", () => {
    expect(coerceInput("42")).toBe(42);
    expect(coerceInput("-3.5e2")).toBe(-350);
    expect(coerceInput("=SUM(A1:A3)")).toBe("=SUM(A1:A3)");
    expect(coerceInput("TRUE")).toBe(true);
    expect(coerceInput("hello")).toBe("hello");
    expect(coerceInput("")).toBeNull();
    expect(coerceInput("0123")).toBe(123); // numeric text loses its zero — documented limitation
  });

  it("records edits in order, coalescing consecutive cell writes per sheet", async () => {
    await journalOpen("t1", "/tmp/a.xlsx");
    journalRecordEdit("t1", { k: "cells", sheet: "S", cells: [{ r: 0, c: 0, v: "1" }] });
    journalRecordEdit("t1", { k: "cells", sheet: "S", cells: [{ r: 1, c: 0, v: "2" }] });
    journalRecordEdit("t1", { k: "rows", op: "insert", sheet: "S", at: 0, count: 1 });
    journalRecordEdit("t1", { k: "cells", sheet: "S", cells: [{ r: 2, c: 0, v: "=A1" }] });
    journalRecordEdit("t1", { k: "sheet", op: "add", name: "New" });
    const edits = journalPeek("t1");
    expect(edits.map((e) => e.k)).toEqual(["cells", "rows", "cells", "sheet"]);
    expect((edits[0] as any).cells).toHaveLength(2);
  });

  it("flushes a journal with content and deletes it when empty", async () => {
    invokeCalls.length = 0;
    await journalOpen("t2", "/tmp/b.xlsx");
    await journalFlush("t2", [], emptyManual());
    expect(invokeCalls.filter((c) => c.cmd === "journal_write")).toHaveLength(0);
    journalRecordEdit("t2", { k: "cells", sheet: "S", cells: [{ r: 0, c: 0, v: "x" }] });
    await journalFlush("t2", [], emptyManual());
    const writes = invokeCalls.filter((c) => c.cmd === "journal_write");
    expect(writes).toHaveLength(1);
    expect(writes[0].args.name).toBe(journalName("/tmp/b.xlsx"));
    const payload = JSON.parse(writes[0].args.json);
    expect(payload.stamp).toBe("123:456");
    expect(payload.edits[0].cells[0]).toEqual({ r: 0, c: 0, v: "x" });
    // Unchanged → no second write.
    await journalFlush("t2", [], emptyManual());
    expect(invokeCalls.filter((c) => c.cmd === "journal_write")).toHaveLength(1);
  });

  it("after a save, batches that existed at save time are never journaled again", async () => {
    invokeCalls.length = 0;
    await journalOpen("t-save", "/tmp/s.xlsx");
    const before = { id: "b-old", status: "accepted", created_at: new Date(Date.now() - 60_000).toISOString(), mutations: [{ type: "set_cell", address: { sheet: "S", row: 0, col: 0 }, new_value: 1 }] } as any;
    await journalFlush("t-save", [before], emptyManual());
    expect(invokeCalls.filter((c) => c.cmd === "journal_write")).toHaveLength(1);
    // Save: the journal is cleared; the SAME unpersisted batch (a surgical
    // save leaves batches unpersisted on purpose) must not come back on
    // the next flush.
    await journalClear("t-save");
    invokeCalls.length = 0;
    await journalFlush("t-save", [before], emptyManual());
    expect(invokeCalls.filter((c) => c.cmd === "journal_write")).toHaveLength(0);
    // A batch created after the save is journaled.
    await new Promise((r) => setTimeout(r, 5));
    const after = { ...before, id: "b-new", created_at: new Date().toISOString() };
    await journalFlush("t-save", [before, after], emptyManual());
    const writes = invokeCalls.filter((c) => c.cmd === "journal_write");
    expect(writes).toHaveLength(1);
    expect(JSON.parse(writes[0].args.json).batches.map((b: any) => b.id)).toEqual(["b-new"]);
  });

  it("replays edits in order, then batch styles and toolbar mirrors", () => {
    const { grid, calls } = fakeGrid();
    const j: any = {
      version: 1, path: "/tmp/c.xlsx", stamp: "1:1", updatedAt: new Date().toISOString(),
      edits: [
        { k: "sheet", op: "add", name: "Recon" },
        { k: "cells", sheet: "Recon", cells: [{ r: 0, c: 0, v: "Total" }, { r: 0, c: 1, v: "=SUM(Data!A:A)" }] },
        { k: "rows", op: "insert", sheet: "Data", at: 4, count: 2 },
        { k: "cells", sheet: "Data", cells: [{ r: 4, c: 0, v: "100" }] },
        { k: "sheet", op: "rename", name: "Recon", newName: "Reconciliation" },
      ],
      batches: [{ id: "b1", status: "pending", mutations: [
        { type: "set_cell", address: { sheet: "Data", row: 4, col: 0 }, new_value: 100 },
        { type: "set_format", sheet: "Reconciliation", cells: [{ row: 0, col: 0 }], new_format: { bold: true } },
      ] }],
      manual: { ...emptyManual(), freeze: { sheet: "Data", freezeRows: 1, freezeCols: 0 }, merges: [{ sheet: "Data", range: "A1:B1", merge: true }] },
    };
    const res = replayJournal("t3", grid, j);
    expect(res).toMatchObject({ cells: 3, structural: 1, sheets: 2, partial: false });
    expect(calls[0]).toBe('createSheet("Recon",null)');
    expect(calls[1]).toContain('setCells("Recon"');
    expect(calls[1]).toContain('"value":"=SUM(Data!A:A)"');
    expect(calls[2]).toBe('insertRows("Data",4,2)');
    expect(calls[3]).toContain('"value":100');
    expect(calls[4]).toBe('renameSheet("Recon","Reconciliation")');
    // set_cell from the batch is NOT replayed (it came through the edit stream); set_format is.
    expect(calls.filter((c) => c.startsWith("setCells")).length).toBe(2);
    expect(calls.some((c) => c.startsWith("setCellFormat") || c.startsWith("setRangeFormat"))).toBe(true);
    expect(calls).toContain('mergeCells("Data",0,0,0,1)');
    expect(calls).toContain('freezePanes("Data",1,0)');
    expect(describeJournal(j)).toContain("3 cell edits");
  });

  it("flags a resync marker as partial", () => {
    const { grid } = fakeGrid();
    const j: any = { version: 1, path: "x", stamp: "1:1", updatedAt: "", edits: [{ k: "resync", reason: "sort" }], batches: [], manual: emptyManual() };
    expect(replayJournal("t4", grid, j).partial).toBe(true);
    expect(describeJournal(j)).toContain("partial");
  });
});
