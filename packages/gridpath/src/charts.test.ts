import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { describe, it, expect } from "vitest";
import JSZip from "jszip";

const enginePath = path.resolve(import.meta.dirname, "../engine/gridpath_engine.js");
const built = fs.existsSync(enginePath);
const engine = built ? createRequire(import.meta.url)(enginePath) : null;
const fixture = path.resolve(import.meta.dirname, "../../../examples/embed-demo/public/the-ai-bill.xlsx");
const chartPart = "xl/drawings/charts/chart1.xml";

describe.skipIf(!built)("native line chart integration", () => {
  it("reads live values, restores them, and reopens a native save with current caches", async () => {
    const bytes = fs.readFileSync(fixture);
    const wb = engine.Workbook.openDeferred(bytes);
    const chart = () => JSON.parse(wb.charts())[0];
    try {
      expect(chart().unsupported).toBeNull();
      wb.evaluatePending();
      const base = chart().series[0].values;
      expect(base).toHaveLength(31);
      expect(base.reduce((sum: number, v: number) => sum + v, 0)).toBeCloseTo(2019.65438772, 6);
      wb.setInput(0, 5, 2, "0.8");
      const delta = JSON.parse(wb.evaluateEdited());
      expect(delta.timing.evaluateMs).toBeGreaterThan(0);
      const edited = chart().series[0].values;
      for (let i = 0; i < 31; i++) expect(edited[i]).toBe(JSON.parse(wb.cell(0, 23+i, 2)).value);
      expect(edited.reduce((sum: number, v: number) => sum + v, 0)).toBeCloseTo(877.01962248, 6);

      const refreshed = JSON.parse(wb.recalculatedCells());
      const sheets = wb.sheetNames().map((name: string) => ({ name, cells: refreshed
        .filter((c: any) => c.sheet === name).map((c: any) => ({ r: c.r, c: c.c, cache: c.value })) }));
      sheets[0].cells.push({ r: 4, c: 1, v: { t: "n", n: .8 } });
      const saved = engine.patch(bytes, JSON.stringify({ version: 1, sheets }));
      const originalZip = await JSZip.loadAsync(bytes), savedZip = await JSZip.loadAsync(saved);
      const originalXml = await originalZip.file(chartPart)!.async("string");
      const savedXml = await savedZip.file(chartPart)!.async("string");
      const withoutCache = (s: string) => s.replace(/<c:(numCache|strCache)[\s\S]*?<\/c:\1>/g, "");
      expect(withoutCache(savedXml)).toBe(withoutCache(originalXml));
      expect(savedXml).not.toBe(originalXml);
      const reopened = engine.Workbook.openDeferred(saved);
      try {
        const cached = JSON.parse(reopened.charts())[0].series[0].values;
        expect(cached).toHaveLength(31);
        cached.forEach((v: number, i: number) => expect(v).toBeCloseTo(edited[i], 10));
      }
      finally { reopened.free(); }

      wb.setInput(0, 5, 2, "0.2"); wb.evaluateEdited();
      expect(chart().series[0].values).toEqual(base);
      wb.renameSheet(0, "AI Bill"); wb.evaluateEdited();
      expect(chart().sheet).toBe("AI Bill");
      expect(chart().series[0].valuesRef).toBe("'AI Bill'!$B$23:$B$53");
      expect(chart().series[0].values).toEqual(base);
    } finally { wb.free(); }
  }, 60_000);
});
