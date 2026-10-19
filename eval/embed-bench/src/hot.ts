// Handsontable has no xlsx support and HyperFormula no file I/O: the stack a
// developer assembles is ExcelJS (read/write) + HyperFormula (calculate) +
// Handsontable (display). This page is that glue, kept straightforward.
import ExcelJS from "exceljs";
import { DetailedCellError, HyperFormula } from "hyperformula";
import Handsontable from "handsontable";
import "handsontable/styles/handsontable.min.css";
import "handsontable/styles/ht-theme-main.min.css";
import { B, fail, inputs, ref, serialUtc, toBase64, until } from "./common";

const stripPrefixes = (f: string) => f.replace(/_xlfn\.|_xlws\.|_xludf\./g, "");

function cellInput(cell: ExcelJS.Cell): unknown {
  if (cell.formula) return `=${stripPrefixes(cell.formula)}`;
  const v = cell.value as any;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return serialUtc(v);
  if (typeof v === "object") {
    if (v.richText) return v.richText.map((r: any) => r.text).join("");
    if (v.text !== undefined) return v.text;
    if (v.error) return v.error;
    if (v.result !== undefined) return v.result;
    return null;
  }
  return v;
}

function plain(v: unknown): unknown {
  if (v instanceof DetailedCellError) return { e: v.value };
  return v;
}

(async () => {
  const { bytes, prep } = await inputs();
  B.t.start = performance.now();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer as ArrayBuffer);
  const sheets: Record<string, unknown[][]> = {};
  for (const ws of wb.worksheets) {
    const rows: unknown[][] = [];
    ws.eachRow({ includeFront: false } as any, (row, r) => {
      const out: unknown[] = [];
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        out[c - 1] = cellInput(cell);
      });
      rows[r - 1] = out;
    });
    for (let r = 0; r < rows.length; r++) rows[r] ??= [];
    sheets[ws.name] = rows;
  }
  // Workbook names ExcelJS kept, as named expressions (one range each).
  const named = (wb.definedNames as any).model
    .filter((n: any) => n.ranges?.length === 1 && /^[A-Za-z_\\][\w.]*$/.test(n.name) && !n.name.startsWith("_xlnm"))
    .map((n: any) => ({ name: n.name, expression: `=${n.ranges[0]}` }));
  const config = {
    licenseKey: "gpl-v3",
    maxRows: 1_048_576,
    maxColumns: 16_384,
    // HyperFormula's documented Excel-compatibility settings.
    leapYear1900: true,
    evaluateNullToZero: true,
    useWildcards: true,
    useRegularExpressions: false,
    caseSensitive: false,
    matchWholeCell: true,
  };
  let hf: HyperFormula;
  try {
    hf = HyperFormula.buildFromSheets(sheets as any, config, named);
  } catch (e) {
    console.warn("named expressions rejected, building without them:", e);
    hf = HyperFormula.buildFromSheets(sheets as any, config);
  }
  const shown = wb.worksheets[0].name;
  const shownId = hf.getSheetId(shown)!;
  const hot = new Handsontable(document.getElementById("app")!, {
    data: hf.getSheetValues(shownId).map((row) => row.map((v) => (v instanceof DetailedCellError ? v.value : v))),
    rowHeaders: true,
    colHeaders: true,
    width: "100%",
    height: "100%",
    readOnly: true,
    themeName: "ht-theme-main",
    licenseKey: "non-commercial-and-evaluation",
  });
  // HyperFormula calculates everything while building.
  B.t.interactive = B.t.calculated = performance.now();
  B.phase = "interactive";

  const i = ref(prep.input.ref);
  const d = ref(prep.dep.ref);
  const inId = hf.getSheetId(prep.input.sheet)!;
  const depId = hf.getSheetId(prep.dep.sheet)!;
  const read = () => plain(hf.getCellValue({ sheet: depId, row: d.row - 1, col: d.col - 1 }));
  B.edit = async (value) => {
    const before = read();
    const t0 = performance.now();
    const changes = hf.setCellContents({ sheet: inId, row: i.row - 1, col: i.col - 1 }, [[value]]);
    // Keep the visible sheet in step with the engine.
    const visible = changes.filter((c: any) => c.address?.sheet === shownId).map((c: any) => [c.address.row, c.address.col, plain(c.newValue)] as [number, number, unknown]);
    if (visible.length) hot.setDataAtCell(visible as any, "engine");
    wb.getWorksheet(prep.input.sheet)!.getCell(prep.input.ref).value = value;
    return until(read, before, t0);
  };
  B.dump = async () =>
    prep.sheets.map((s) => {
      const id = hf.getSheetId(s.name);
      const out: unknown[] = [];
      for (let k = 0; k < s.cells.length; k += 2) out.push(id === undefined ? undefined : plain(hf.getCellValue({ sheet: id, row: s.cells[k] - 1, col: s.cells[k + 1] - 1 })));
      return out;
    });
  B.save = async () => toBase64(new Uint8Array(await wb.xlsx.writeBuffer()));
  B.phase = "calculated";
})().catch(fail);
