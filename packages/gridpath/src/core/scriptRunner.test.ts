import { describe, it, expect } from "vitest";
import { buildScriptModelFromEngineSheets, sheetsReferencedByScript, executeSheetScript } from "./scriptRunner";

describe("engine-sourced read model", () => {
  const enc = (sheet: number, cells: Array<{ r: number; c: number; v?: unknown; f?: string }>) =>
    new TextEncoder().encode(JSON.stringify({ sheet, cells, styles: {} })).buffer;

  it("builds the model from engine per-sheet JSON (1-based → 0-based) and marks the rest unloaded", () => {
    const m = buildScriptModelFromEngineSheets(["Data", "Other"], {
      Data: JSON.stringify({ sheet: 0, cells: [{ r: 1, c: 1, v: "Summary" }, { r: 2, c: 1, v: 5, f: "=1+4" }, { r: 3, c: 1, v: null }] }),
    });
    expect(m.sheets.Data["0,0"]).toEqual({ v: "Summary", f: null });
    expect(m.sheets.Data["1,0"]).toEqual({ v: 5, f: "=1+4" });
    expect(m.sheets.Data["2,0"]).toBeUndefined();
    expect(m.unloaded).toEqual(["Other"]);
  });

  it("picks the sheets a script names, or all when it enumerates", () => {
    const names = ["Store Sales", "Regions", "Summary"];
    expect(sheetsReferencedByScript('const s = sheet("Store Sales"); sheet("Summary").set("A1", 1)', names).names).toEqual(["Store Sales", "Summary"]);
    expect(sheetsReferencedByScript("for (const n of sheets()) log(n)", names).all).toBe(true);
    expect(sheetsReferencedByScript("const n = pick(); sheet(n).values('A1:B2')", names).all).toBe(true);
  });

  it("runs a filter-and-write script against engine bytes without touching the grid", async () => {
    const fetched: string[] = [];
    const source = {
      sheetNames: ["Sales", "Report"],
      fetchSheet: async (name: string) => {
        fetched.push(name);
        if (name !== "Sales") return null;
        return enc(0, [
          { r: 1, c: 1, v: "store-1" }, { r: 1, c: 2, v: "Acme" }, { r: 1, c: 3, v: 10 },
          { r: 2, c: 1, v: "store-2" }, { r: 2, c: 2, v: "Initech" }, { r: 2, c: 3, v: 7 },
          { r: 3, c: 1, v: "store-3" }, { r: 3, c: 2, v: "Acme" }, { r: 3, c: 3, v: 2.5 },
        ]);
      },
    };
    const code = `
      const rows = sheet("Sales").values("A1:C3");
      const acme = rows.filter((r) => r[1] === "Acme");
      sheet("Report").setValues("A1", acme.map((r) => [r[0], r[2]]));
      log(acme.length);
    `;
    const res = await executeSheetScript({ getWorkbookSnapshot: () => { throw new Error("grid must not be read"); } }, code, { source });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    // Every sheet named as a literal is fetched (the engine has no "Report" yet → null → unloaded; writes still work).
    expect(fetched).toEqual(["Sales", "Report"]);
    expect(res.logs).toEqual(["2"]);
    expect(res.mutations.map((m: any) => [m.address?.sheet, m.address?.row, m.address?.col, m.new_value])).toEqual([
      ["Report", 0, 0, "store-1"], ["Report", 0, 1, 10], ["Report", 1, 0, "store-3"], ["Report", 1, 1, 2.5],
    ]);
  });

  it("reading a sheet the script did not name fails loudly instead of returning blanks", async () => {
    // The engine cannot supply "B" (fetch returns null) → B is in the model but unloaded.
    const source = { sheetNames: ["A", "B"], fetchSheet: async (name: string) => (name === "A" ? enc(0, [{ r: 1, c: 1, v: 1 }]) : null) };
    const code = 'sheet("A").get("A1"); log(sheet("B").get("A1").value)';
    const res = await executeSheetScript({ getWorkbookSnapshot: () => null }, code, { source });
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error).toMatch(/sheet "B" was not loaded/);
  });
});
