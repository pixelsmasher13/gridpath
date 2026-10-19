import * as GC from "@mescius/spread-sheets";
import "@mescius/spread-sheets-io";
// The feature plugins a developer working with Excel files would load.
import "@mescius/spread-sheets-shapes";
import "@mescius/spread-sheets-charts";
import "@mescius/spread-sheets-pivot-addon";
import "@mescius/spread-sheets-slicers";
import "@mescius/spread-sheets/styles/gc.spread.sheets.excel2013white.css";
import { B, fail, inputs, ref, toBase64, until } from "./common";

// No license key: evaluation mode.
function plain(v: unknown): unknown {
  if (v instanceof Date) return (v.getTime() - v.getTimezoneOffset() * 60_000 - Date.UTC(1899, 11, 30)) / 86_400_000;
  if (v instanceof GC.Spread.CalcEngine.CalcError) return { e: v.toString() };
  return v;
}

(async () => {
  const { id, bytes, prep } = await inputs();
  const spread = new GC.Spread.Sheets.Workbook(document.getElementById("app")!);
  B.t.start = performance.now();
  await new Promise<void>((resolve, reject) =>
    spread.import(new File([bytes as BlobPart], `${id}.xlsx`), () => resolve(), (e: any) => reject(new Error(e?.errorMessage ?? JSON.stringify(e))), {
      fileType: GC.Spread.Sheets.FileType.excel,
      // ?mode=lazy: SpreadJS's large-file mode (only the active sheet loads up front).
      ...(new URLSearchParams(location.search).get("mode") === "lazy" ? { openMode: GC.Spread.Sheets.OpenMode.lazy } : {}),
    } as any),
  );
  B.t.interactive = performance.now();
  B.phase = "interactive";
  // Import keeps the file's cached results; recompute everything, as the
  // other engines do on open.
  spread.calculate(GC.Spread.Sheets.CalculationType.rebuild);
  B.t.calculated = performance.now();
  const inSheet = spread.getSheetFromName(prep.input.sheet);
  const depSheet = spread.getSheetFromName(prep.dep.sheet);
  const i = ref(prep.input.ref);
  const d = ref(prep.dep.ref);
  const read = () => plain(depSheet.getValue(d.row - 1, d.col - 1));
  B.edit = async (value) => {
    const before = read();
    const t0 = performance.now();
    inSheet.setValue(i.row - 1, i.col - 1, value);
    return until(read, before, t0);
  };
  B.dump = async () =>
    prep.sheets.map((s) => {
      const sh = spread.getSheetFromName(s.name);
      const out: unknown[] = [];
      for (let k = 0; k < s.cells.length; k += 2) out.push(sh ? plain(sh.getValue(s.cells[k] - 1, s.cells[k + 1] - 1)) : undefined);
      return out;
    });
  B.save = () =>
    new Promise((resolve, reject) =>
      spread.export(
        async (blob: Blob) => resolve(toBase64(new Uint8Array(await blob.arrayBuffer()))),
        (e: any) => reject(new Error(e?.errorMessage ?? JSON.stringify(e))),
        { fileType: GC.Spread.Sheets.FileType.excel } as any,
      ),
    );
  B.phase = "calculated";
})().catch(fail);
