/**
 * xlsx → Univer import conversion, worker-hostable.
 *
 * Everything here is UI-free (no React, no Univer runtime) so the whole
 * conversion can run inside a Web Worker: parsing a large model with ExcelJS
 * takes multiple seconds of pure CPU (project-279.xlsx: ~8s / 1.5 GB), and on
 * the webview main thread that reads as a frozen window. UniverGrid calls
 * `parseXlsxWorkbook`, which prefers the worker and transparently falls back
 * to an in-process parse when workers are unavailable or fail.
 *
 * The worker hands back only structured-cloneable data (Univer snapshot,
 * style ops, outlines, pins, extracted images). The ExcelJS Workbook object
 * itself — needed only for the full-export save path — stays behind in the
 * worker and dies with it; UniverGrid lazily re-parses the retained source
 * bytes if and when a full export actually happens. Surgical saves (the
 * default) never need it. This also means the multi-hundred-MB ExcelJS
 * object graph no longer lives for the whole session.
 */

import ExcelJS from "exceljs";
import { applyExcelJsPatches } from "./exceljsPatches";
import { parseThemePalette, resolveXlsxColor } from "./xlsxColors";
import { stripXlfnPrefixes } from "./xlfnCompat";
import { computeContentExtent, trimBoundsFor } from "./importTrim";
import {
  CF_RESOURCE_NAME,
  DV_RESOURCE_NAME,
  NOTE_RESOURCE_NAME,
  excelCfToUniverRules,
  excelDvToUniverRules,
  notesToUniverResource,
  noteTextFromExcelJs,
  parseIterativeCalc,
  type ImportedNote,
} from "./featureIo";
import type { CellFormatShape, BordersShape } from "./components/UniverGrid";

// Must run before the first xlsx.load — see exceljsPatches.ts for the bugs
// these work around (unparseable chart drawings, defined-name explosions).
// Idempotent, so the main thread applying them too is harmless; in a worker
// context this is the only call site.
applyExcelJsPatches();

export type StyleOp = {
  sheet: string;
  row: number;
  col: number;
  format: CellFormatShape;
  background: string | null;
  borders: BordersShape | null;
};

/** Column/row grouping depth per index (only grouped indices listed). */
export type SheetOutline = {
  cols: Array<[index: number, level: number]>;
  rows: Array<[index: number, level: number]>;
  maxColLevel: number;
  maxRowLevel: number;
};

/**
 * Unresolvable-formula pinning. Two kinds of formulas can never be computed
 * by our engine:
 *
 *  - External-workbook refs like
 *    `HLOOKUP(C4,'[2]Income Statement'!$A$1:CL$268,...)` — the other .xlsx
 *    is often on a network share we can't reach. Excel displays the cached
 *    results stored in xl/externalLinks/*.xml.
 *  - Add-in (XLL) function calls like `_xll.FDS(...)` / `_xll.TR(...)` —
 *    FactSet/Refinitiv terminal functions that only exist with the vendor
 *    add-in installed. Excel displays the cached result stored in the cell.
 *
 * Univer recalculates every formula cell on load, can't resolve either
 * shape, and replaces the cached value with an error that cascades to
 * downstream cells (financial models wrap these in thousands of IFERRORs).
 *
 * So on import we "pin" such cells: strip the formula from the Univer model
 * (leaving the cached value as a static) and record the pin here. On save,
 * untouched pinned cells are skipped so the original ExcelJS cell — which
 * still carries { formula, result } — rounds through the file unchanged.
 * If the user overwrites a pinned cell, the new content wins and the
 * external formula is dropped from the file, same as overwriting it in
 * Excel. Structural edits (row/col splices, sheet renames) can shift a cell
 * away from its recorded pin key; those cells degrade to a static-value
 * write on save, never to a wrong value.
 */
export type ExternalPinMap = Map<string, { f: string; v: any }>;
export const externalPinsByWorkbook = new WeakMap<ExcelJS.Workbook, ExternalPinMap>();
export const externalPinKey = (sheetName: string, r: number, c: number) => `${sheetName} ${r},${c}`;

/**
 * True when a formula references another workbook (`[2]Sheet1!A1` or
 * `'[2]Income Statement'!$A70`). String literals are blanked first so a
 * bracket inside quoted text can't false-positive. The leading-character
 * guard keeps structured table refs like `Table1[2]` (identifier before the
 * bracket) from matching.
 */
function formulaHasExternalWorkbookRef(formula: string): boolean {
  const withoutStrings = formula.replace(/"[^"]*"/g, '""');
  return /(^|[^A-Za-z0-9_\]])\[\d+\]/.test(withoutStrings);
}

/**
 * True when a formula calls an add-in (XLL) function — `_xll.FDS(...)`,
 * `_xll.TR(...)` etc. The `_xlfn.` prefix (future built-ins like XLOOKUP)
 * is deliberately NOT matched: those are real functions an engine may
 * support. String literals are blanked first so "_xll." inside quoted text
 * can't false-positive.
 */
function formulaHasAddinFunctionRef(formula: string): boolean {
  const withoutStrings = formula.replace(/"[^"]*"/g, '""');
  return /(^|[^A-Za-z0-9_.])_xll\./i.test(withoutStrings);
}

/**
 * Parse xlsx bytes with ExcelJS, returning:
 *   - the ExcelJS Workbook (kept alive in memory for round-trip save)
 *   - the IWorkbookData Univer needs to render the grid
 *   - a list of cell-level style ops to push into Univer's facade after
 *     createWorkbook so the user actually SEES the workbook's original
 *     colors / fonts / number formats / alignment (not just preserves them
 *     on save).
 *
 * Everything ExcelJS knows about that we DON'T touch — charts, conditional
 * formatting, named ranges, data validation, comments, drawings, themes,
 * sheet protection, frozen panes — stays inside the ExcelJS object. SheetJS
 * used to strip all of that on the round trip.
 */
export async function xlsxBytesToWorkbook(bytes: Uint8Array): Promise<{
  excelJs: ExcelJS.Workbook;
  univerData: any;
  styleOps: StyleOp[];
  outlines: Record<string, SheetOutline>;
  featureCounts: { cf: number; dv: number; notes: number; cfDropped: string[] };
}> {
  const excelJs = new ExcelJS.Workbook();
  // ExcelJS wants an ArrayBuffer; copy out of the Uint8Array's backing buffer
  // in case it's a slice (which is the case for Tauri-delivered bytes).
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  await excelJs.xlsx.load(ab as ArrayBuffer);

  // Sheets that hide their gridlines with a background picture (the classic
  // bank-model trick: a white PNG tiled behind the sheet). ExcelJS doesn't
  // model <picture>, so this reads the raw package.
  const bgPictureSheets = await detectBackgroundPictureSheets(bytes);

  // Theme palette for color resolution — the vast majority of fills/fonts in
  // real files are theme-indexed, not literal argb.
  const themePalette = parseThemePalette((excelJs.model as any)?.themes?.theme1);

  const sheets: Record<string, any> = {};
  const sheetOrder: string[] = [];
  const styleOps: StyleOp[] = [];
  const outlines: Record<string, SheetOutline> = {};
  const externalPins: ExternalPinMap = new Map();
  externalPinsByWorkbook.set(excelJs, externalPins);
  // CF / data-validation / notes, keyed by Univer sheet id — seeded into the
  // snapshot's plugin resources so they RENDER, not just survive in the file.
  const cfResource: Record<string, any[]> = {};
  const dvResource: Record<string, any[]> = {};
  const importedNotes: ImportedNote[] = [];
  const cfDropped: string[] = [];

  excelJs.worksheets.forEach((ws) => {
    const name = ws.name;
    const id = `sheet_${name.replace(/[^a-zA-Z0-9]/g, "_")}`;
    const cellData: Record<number, Record<number, any>> = {};
    let maxRow = 0;
    let maxCol = 0;

    // Merged cell ranges live in worksheet.model.merges as A1 strings like
    // "A1:A5" or "B2:D2". Translate to Univer's IRange shape so Univer
    // renders the merge natively on createWorkbook. Parsed BEFORE the cell
    // walk because the trim extent below counts merges as content.
    const mergeData: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }> = [];
    const merges: any = (ws as any).model?.merges ?? [];
    if (Array.isArray(merges)) {
      for (const m of merges) {
        const ir = parseA1RangeString(typeof m === "string" ? m : String(m?.range ?? ""));
        if (ir) mergeData.push(ir);
      }
    }

    // Styled-blank tail trim (see importTrim.ts). Style-only cells beyond
    // the content extent + margin are dropped: a column formatted to the
    // bottom of Excel's grid otherwise floods the snapshot with ~1M junk
    // cell objects (project-279.xlsx's "Share Tracker" — 90 MB of XML, 88%
    // of the workbook's cells) and dominates open time and memory.
    const extent = computeContentExtent(ws, mergeData);
    const { trimRow, trimCol } = trimBoundsFor(extent);
    if (extent.anyMaxRow > trimRow + 1000 || extent.anyMaxCol > trimCol + 100) {
      console.info(
        `[import] ${name}: trimming styled-blank tail (content ends r${extent.contentMaxRow} c${extent.contentMaxCol}, styling extends r${extent.anyMaxRow} c${extent.anyMaxCol})`,
      );
    }

    // includeEmpty at the ROW level too: eachRow's default skips rows where
    // no cell has a value — which is exactly what a fully styled-but-blank
    // band row looks like. Rows with no cells at all iterate zero cells
    // below, so the cost of including them is nil.
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const r = rowNumber - 1; // ExcelJS is 1-indexed
      // Past the trim boundary only style-only cells exist (the extent walk
      // counted values, notes and merges), so the whole row is droppable.
      if (r > trimRow) return;
      // includeEmpty: styled-but-valueless cells (<c r="B2" s="5"/>) are how
      // colored bands and header backgrounds are painted — skipping them
      // loses most of a sheet's fills.
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const c = colNumber - 1;
        if (c > trimCol) return; // style-only by construction — see above
        const cellDescriptor = cellFromExcelJs(cell);
        // Cell note (legacy Excel comment). Collected BEFORE the empty-slot
        // bail-out below — a note can sit on an otherwise empty cell.
        const noteText = (cell as any).note ? noteTextFromExcelJs((cell as any).note) : null;
        if (noteText) importedNotes.push({ sheetId: id, row: r, col: c, text: noteText });
        // Pin external-workbook and add-in (_xll.*) formulas to their cached
        // result — Univer can't resolve either and would turn the value into
        // an error. Only pin when a cached value exists; otherwise keep the
        // formula so the cell shows an honest error instead of silently
        // going blank.
        if (
          cellDescriptor.f &&
          cellDescriptor.v !== undefined &&
          (formulaHasExternalWorkbookRef(String(cellDescriptor.f)) ||
            formulaHasAddinFunctionRef(String(cellDescriptor.f)))
        ) {
          externalPins.set(externalPinKey(name, r, c), {
            f: cellDescriptor.f,
            v: cellDescriptor.v,
          });
          delete cellDescriptor.f;
        }
        const { format, background, borders } = extractStyleFromExcelJs(cell, themePalette);
        // Embed style INLINE in cellData. Univer reads this on createWorkbook
        // and renders accordingly — number_format, fonts, colors, borders,
        // alignment all land on the right cells without a second facade pass.
        // The post-load styleOps loop is kept as a defensive fallback only.
        const s = buildUniverStyle(format, background, borders);
        if (s) cellDescriptor.s = s;
        const hasContent = cellDescriptor.f !== undefined || cellDescriptor.v !== undefined;
        if (!hasContent && !s) return; // truly empty slot
        if (!cellData[r]) cellData[r] = {};
        cellData[r][c] = cellDescriptor;
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;

        if (Object.keys(format).length > 0 || background || borders) {
          styleOps.push({ sheet: name, row: r, col: c, format, background, borders });
        }
      });
    });

    // Column widths — Excel stores them in character units. Convert to px
    // (rough approximation: width * 7 + 5 for default font). Univer wants
    // pixels in columnData[N] = { w: <px> }.
    const columnData: Record<number, { w?: number; hd?: number }> = {};
    const outlineCols: Array<[number, number]> = [];
    if ((ws as any).columns) {
      ((ws as any).columns as any[]).forEach((col: any, idx: number) => {
        if (col && typeof col.width === "number" && col.width > 0) {
          columnData[idx] = { w: Math.round(col.width * 7 + 5) };
        }
        if (col?.hidden) (columnData[idx] = columnData[idx] ?? {}).hd = 1;
        const ol = Number(col?.outlineLevel ?? 0);
        if (ol > 0) outlineCols.push([idx, ol]);
      });
    }

    // Row heights — Excel stores in points. 1 point ≈ 1.333 pixels at 96 DPI.
    // includeEmpty: hidden/height/outline are ROW attributes that exist on
    // valueless rows too — a hidden spacer row skipped here stayed VISIBLE
    // in the grid, leaving ragged half-collapsed bands (NFLX's Model sheet
    // alone has 63 hidden rows with no cell values).
    const rowData: Record<number, { h?: number; hd?: number }> = {};
    const outlineRows: Array<[number, number]> = [];
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const idx = rowNumber - 1;
      // Same trim as the cell walk: a custom height / hidden flag / outline
      // level on a contentless tail row (Excel writes one per row when a
      // column is formatted to the grid bottom) would otherwise materialise
      // ~1M rowData entries here.
      if (idx > trimRow) return;
      if (typeof row.height === "number" && row.height > 0) {
        rowData[idx] = { h: Math.round(row.height * 1.333) };
      }
      if ((row as any).hidden) (rowData[idx] = rowData[idx] ?? {}).hd = 1;
      const ol = Number((row as any).outlineLevel ?? 0);
      if (ol > 0) outlineRows.push([idx, ol]);
    });

    if (outlineCols.length || outlineRows.length) {
      outlines[name] = {
        cols: outlineCols,
        rows: outlineRows,
        maxColLevel: outlineCols.reduce((m, [, l]) => Math.max(m, l), 0),
        maxRowLevel: outlineRows.reduce((m, [, l]) => Math.max(m, l), 0),
      };
    }

    // Conditional formatting + data validation → Univer plugin resources.
    // Unmappable CF rule families (icon sets…) are skipped here but stay in
    // the file: surgical saves never touch those parts.
    try {
      const { rules, dropped } = excelCfToUniverRules(
        (ws as any).conditionalFormattings,
        themePalette,
        { maxRow, maxCol },
      );
      if (rules.length) cfResource[id] = rules;
      cfDropped.push(...dropped);
    } catch (e) {
      console.warn("[univer] CF import failed on", name, e);
    }
    try {
      const dvRules = excelDvToUniverRules((ws as any).dataValidations?.model, { maxRow, maxCol });
      if (dvRules.length) dvResource[id] = dvRules;
    } catch (e) {
      console.warn("[univer] data-validation import failed on", name, e);
    }

    // Sheet tab color (the colored stripe on the tab at the bottom of Excel).
    // ExcelJS exposes it via worksheet.properties.tabColor ({argb}/{theme}).
    const tabColor = resolveXlsxColor((ws as any).properties?.tabColor, themePalette);

    // Sheet visibility. Excel has three states: visible, hidden (user can
    // unhide from the tab bar) and veryHidden (only VBA/editors). Univer
    // only models hidden/shown, so both map to hidden — the important part
    // is that helper/data sheets the model author hid don't show up as
    // ordinary tabs. veryHidden survives round-trips untouched because both
    // save paths keep the original workbook.xml state attribute.
    const sheetHidden = ws.state === "hidden" || ws.state === "veryHidden";

    // Freeze panes — ExcelJS stores them in `worksheet.views` as
    // [{ state: 'frozen', xSplit, ySplit }]. xSplit = frozen-column count,
    // ySplit = frozen-row count.
    let frozenRows = 0;
    let frozenCols = 0;
    const views = (ws as any).views;
    if (Array.isArray(views)) {
      const frozenView = views.find((v: any) => v?.state === "frozen" || v?.state === "frozenSplit");
      if (frozenView) {
        frozenRows = Math.max(0, Number(frozenView.ySplit ?? 0));
        frozenCols = Math.max(0, Number(frozenView.xSplit ?? 0));
      }
    }

    // Gridline visibility. Two ways a sheet hides its grid: the sheetView
    // showGridLines flag (ExcelJS exposes it), or an opaque background
    // picture covering the canvas — in that case Excel technically still
    // "shows" gridlines but the user never sees them, so hiding ours is the
    // faithful rendering. Display-only: neither save path touches it.
    const gridHiddenByFlag =
      Array.isArray(views) && views.some((v: any) => v && v.showGridLines === false);
    const hideGridlines = gridHiddenByFlag || bgPictureSheets.has(name);

    sheets[id] = {
      id,
      name,
      // Excel's max is 1,048,576 × 16,384. We use 10000 × 200 — big enough
      // that arrow-key navigation never hits an edge for normal use, small
      // enough that Univer doesn't spend memory pre-allocating massive grids.
      // Was 100 × 26 — that was why ← at column A wrapped to column Z.
      rowCount: Math.max(10000, maxRow + 100),
      columnCount: Math.max(200, maxCol + 10),
      cellData,
      mergeData,
      columnData,
      rowData,
      // Univer's IWorksheetData supports tabColor (string). When null Univer
      // uses its default neutral.
      ...(tabColor ? { tabColor } : {}),
      ...(hideGridlines ? { showGridlines: 0 } : {}),
      // BooleanNumber.TRUE — Univer hides the tab; unhide is available in
      // the tab bar's context menu just like Excel.
      ...(sheetHidden ? { hidden: 1 } : {}),
      // Freeze panes — Univer's IFreeze shape: { xSplit, ySplit, startRow, startColumn }
      // xSplit = frozen-col count, ySplit = frozen-row count, startRow/Column
      // = where scrolling begins. Match ExcelJS's freeze on load.
      ...(frozenRows > 0 || frozenCols > 0
        ? {
            freeze: {
              xSplit: frozenCols,
              ySplit: frozenRows,
              startRow: frozenRows,
              startColumn: frozenCols,
            },
          }
        : {}),
    };
    sheetOrder.push(id);
  });

  if (sheetOrder.length === 0) {
    sheets["sheet1"] = { id: "sheet1", name: "Sheet1", rowCount: 10000, columnCount: 200, cellData: {} };
    sheetOrder.push("sheet1");
  }

  const noteResource = notesToUniverResource(importedNotes);
  const resources: Array<{ name: string; data: string }> = [];
  if (Object.keys(cfResource).length) resources.push({ name: CF_RESOURCE_NAME, data: JSON.stringify(cfResource) });
  if (Object.keys(dvResource).length) resources.push({ name: DV_RESOURCE_NAME, data: JSON.stringify(dvResource) });
  if (Object.keys(noteResource).length) resources.push({ name: NOTE_RESOURCE_NAME, data: JSON.stringify(noteResource) });

  return {
    excelJs,
    univerData: { id: `wb_${Date.now()}`, sheets, sheetOrder, ...(resources.length ? { resources } : {}) },
    styleOps,
    outlines,
    featureCounts: {
      cf: Object.values(cfResource).reduce((a, r) => a + r.length, 0),
      dv: Object.values(dvResource).reduce((a, r) => a + r.length, 0),
      notes: importedNotes.length,
      cfDropped,
    },
  };
}

/**
 * The workbook's iterative-calculation opt-in from xl/workbook.xml. ExcelJS
 * parses <calcPr> but drops its attributes, so this peeks at the raw
 * package (same pattern as detectBackgroundPictureSheets below). null =
 * iterative calc off or unreadable — either way Univer stays at 1 pass.
 */
export async function detectIterativeCalc(bytes: Uint8Array): Promise<{ iterateCount: number } | null> {
  try {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("xl/workbook.xml")?.async("string");
    return xml ? parseIterativeCalc(xml) : null;
  } catch (e) {
    console.warn("[univer] calcPr read failed:", e);
    return null;
  }
}

/**
 * Sheet names whose worksheet XML carries a `<picture>` element — a tiled
 * background image. ExcelJS doesn't parse it, so we peek at the raw package
 * with jszip (already a dependency of ExcelJS). Failures just mean "no
 * background detected" — rendering falls back to showing gridlines.
 */
async function detectBackgroundPictureSheets(bytes: Uint8Array): Promise<Set<string>> {
  const out = new Set<string>();
  try {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(bytes);
    const wbXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    if (!wbXml || !relsXml) return out;

    const ridToTarget = new Map<string, string>();
    for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) ?? []) {
      const id = /\bId="([^"]*)"/.exec(tag)?.[1];
      const target = /\bTarget="([^"]*)"/.exec(tag)?.[1];
      if (id && target) ridToTarget.set(id, target);
    }
    for (const tag of wbXml.match(/<sheet\b[^>]*>/g) ?? []) {
      const name = /\bname="([^"]*)"/.exec(tag)?.[1];
      const rid = /\br:id="([^"]*)"/.exec(tag)?.[1];
      const target = rid && ridToTarget.get(rid);
      if (!name || !target) continue;
      const part = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
      const sheetXml = await zip.file(part)?.async("string");
      if (sheetXml && /<picture[\s/>]/.test(sheetXml)) {
        out.add(decodeXmlEntities(name));
      }
    }
  } catch (e) {
    console.warn("[univer] background picture detection failed:", e);
  }
  return out;
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/**
 * Normalize a value ExcelJS may hand us for cell content. Dates become
 * Excel serial numbers (UTC math — the cell's number format renders them,
 * same as Excel). INVALID dates become undefined: ExcelJS coerces text like
 * "2028E" through a date number format into `Invalid Date`, destroying the
 * original text — omitting the cached value lets Univer recalculate the
 * formula and pull the real text from the referenced cell.
 */
function normalizeExcelValue(v: any): any {
  if (v instanceof Date) {
    const t = v.getTime();
    if (!Number.isFinite(t)) return undefined;
    const serial = (t - Date.UTC(1899, 11, 30)) / 86400000;
    return Math.round(serial * 1e7) / 1e7;
  }
  return v;
}

export function cellFromExcelJs(cell: ExcelJS.Cell): any {
  const out: any = {};
  // ExcelJS cell.value shapes:
  //   primitive (number | string | boolean) | null | Date
  //   { formula, result } — formulas
  //   { sharedFormula, result } — shared formula
  //   { richText: [{ text, font }] } — rich text
  //   { text, hyperlink } — hyperlinks
  //   { error } — error cells (#REF!, #VALUE! etc.)
  const v: any = cell.value;
  if (v && typeof v === "object" && "formula" in v) {
    // stripXlfnPrefixes: `_xlfn.IFS(...)` → `IFS(...)` so Univer evaluates
    // it instead of erroring on an unknown name. Save paths restore the
    // prefix (see xlfnCompat.ts).
    out.f = `=${stripXlfnPrefixes(String((v as any).formula))}`;
    if ("result" in v && (v as any).result !== undefined && (v as any).result !== null) {
      const r = (v as any).result;
      const norm = normalizeExcelValue(
        typeof r === "object" && r !== null && "error" in r ? (r as any).error : r,
      );
      if (norm !== undefined) out.v = norm;
    }
  } else if (v && typeof v === "object" && "sharedFormula" in v) {
    // Shared-formula DEPENDENT: cell.value only carries the master's address
    // ("CT27"), which is useless as a formula — importing it as `=CT27` gave
    // the agent (and the grid) wrong formulas for every filled row. The
    // `formula` getter returns the properly TRANSLATED formula (CU27 →
    // "CU22-CT22"); fall back to the cached value alone when translation
    // isn't available — an honest static beats a wrong reference.
    const translated = (cell as any).formula;
    if (typeof translated === "string" && translated && translated !== (v as any).sharedFormula) {
      out.f = `=${stripXlfnPrefixes(translated)}`;
    }
    if ("result" in v && (v as any).result !== undefined && (v as any).result !== null) {
      const norm = normalizeExcelValue((v as any).result);
      if (norm !== undefined) out.v = norm;
    }
  } else if (v && typeof v === "object" && "richText" in v) {
    out.v = (v as any).richText.map((r: any) => r.text).join("");
  } else if (v && typeof v === "object" && "hyperlink" in v) {
    out.v = (v as any).text ?? String((v as any).hyperlink);
  } else if (v && typeof v === "object" && "error" in v) {
    out.v = (v as any).error;
  } else if (v instanceof Date) {
    // Date cells: Excel stores a SERIAL NUMBER rendered through the cell's
    // number format ("FY "yyyy → "FY 2004"). Storing an ISO string here made
    // Univer display raw text instead of applying the format. Convert to
    // the serial and let the number format render it, same as Excel.
    const norm = normalizeExcelValue(v);
    if (norm !== undefined) out.v = norm;
  } else if (v === null || v === undefined) {
    // empty
  } else if (typeof v === "object") {
    out.v = JSON.stringify(v);
  } else {
    out.v = v;
  }
  return out;
}

/**
 * Convert our CellFormatShape + background + borders into Univer's inline
 * cell style object (the IStyleData shape Univer reads from cellData[r][c].s).
 * Returning a style here means it's applied at createWorkbook time, which is
 * the only way certain attributes (notably number_format) actually take
 * effect — the post-load facade pass via setNumberFormat was a no-op.
 *
 * Univer style codes (from preset-sheets-core IStyleData):
 *   bg = background fill { rgb }
 *   cl = font color { rgb }
 *   bl = bold (0/1)
 *   it = italic (0/1)
 *   ul = underline { s: 0|1 }
 *   st = strikethrough { s: 0|1 }
 *   fs = font size
 *   ff = font family
 *   ht = horizontal align (1=left, 2=center, 3=right)
 *   vt = vertical align (1=top, 2=middle, 3=bottom)
 *   n  = number format { pattern }
 *   bd = borders { t, b, l, r each { s: <style code>, cl: { rgb } } }
 */
function buildUniverStyle(
  format: CellFormatShape,
  background: string | null,
  borders: BordersShape | null,
): any | null {
  const s: any = {};
  if (background) s.bg = { rgb: background };
  if (format.font_color) s.cl = { rgb: format.font_color };
  if (format.bold) s.bl = 1;
  if (format.italic) s.it = 1;
  if (format.underline) s.ul = { s: 1 };
  if (format.strike) s.st = { s: 1 };
  if (format.font_size) s.fs = format.font_size;
  if (format.font_family) s.ff = format.font_family;
  if (format.horizontal_align) {
    s.ht = format.horizontal_align === "left" ? 1 : format.horizontal_align === "center" ? 2 : 3;
  }
  if (format.vertical_align) {
    s.vt = format.vertical_align === "top" ? 1 : format.vertical_align === "middle" ? 2 : 3;
  }
  if (format.number_format) s.n = { pattern: format.number_format };
  // Wrap text — Univer uses tb (text-break): 1=clip, 2=wrap, 3=overflow.
  if (format.wrap_text) s.tb = 2;
  // Indent — Univer's pd (padding) or td (text-indent); the property name
  // varies across versions, so we set both as a defensive measure.
  if (format.indent && format.indent > 0) {
    s.pd = { l: format.indent * 8 };
    (s as any).td = format.indent;
  }

  if (borders) {
    const bd: any = {};
    const sideStyleCode = (style: string): number => {
      // Univer border style codes (approx — actual values may vary by build):
      // 1=thin, 2=hair, 3=dotted, 4=dashed, 5=dashDot, 6=dashDotDot,
      // 7=double, 8=medium, 9=mediumDashed, 10=mediumDashDot,
      // 11=mediumDashDotDot, 12=slantDashDot, 13=thick. Default to thin.
      switch (style) {
        case "thick": return 13;
        case "double": return 7;
        case "medium": return 8;
        case "dashed": return 4;
        case "dotted": return 3;
        case "hair": return 2;
        default: return 1;
      }
    };
    const sideOf = (side: NonNullable<BordersShape["top"]> | undefined) =>
      side
        ? { s: sideStyleCode(side.style), cl: { rgb: side.color ?? "#000000" } }
        : undefined;
    const t = sideOf(borders.top ?? undefined);
    const b = sideOf(borders.bottom ?? undefined);
    const l = sideOf(borders.left ?? undefined);
    const r = sideOf(borders.right ?? undefined);
    if (t) bd.t = t;
    if (b) bd.b = b;
    if (l) bd.l = l;
    if (r) bd.r = r;
    if (Object.keys(bd).length > 0) s.bd = bd;
  }

  return Object.keys(s).length > 0 ? s : null;
}

/**
 * Parse "A1:B2" (or just "A1") into Univer's IRange shape. Returns null for
 * malformed input so the caller can skip silently.
 */
function parseA1RangeString(s: string): {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
} | null {
  const parts = s.split(":");
  const decode = (ref: string): { row: number; col: number } | null => {
    const m = /^\s*([A-Za-z]+)(\d+)\s*$/.exec(ref);
    if (!m) return null;
    const letters = m[1].toUpperCase();
    let col = 0;
    for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
    return { row: parseInt(m[2], 10) - 1, col: col - 1 };
  };
  if (parts.length === 1) {
    const a = decode(parts[0]);
    if (!a) return null;
    return { startRow: a.row, endRow: a.row, startColumn: a.col, endColumn: a.col };
  }
  if (parts.length === 2) {
    const a = decode(parts[0]);
    const b = decode(parts[1]);
    if (!a || !b) return null;
    return {
      startRow: Math.min(a.row, b.row),
      endRow: Math.max(a.row, b.row),
      startColumn: Math.min(a.col, b.col),
      endColumn: Math.max(a.col, b.col),
    };
  }
  return null;
}

function extractStyleFromExcelJs(
  cell: ExcelJS.Cell,
  themePalette: string[] | null,
): {
  format: CellFormatShape;
  background: string | null;
  borders: BordersShape | null;
} {
  const format: CellFormatShape = {};
  let background: string | null = null;
  let borders: BordersShape | null = null;

  const border = cell.border as any;
  if (border) {
    const sideOf = (s: any): BordersShape["top"] => {
      if (!s || !s.style) return null;
      const color = resolveXlsxColor(s.color, themePalette);
      return { style: String(s.style), color };
    };
    const t = sideOf(border.top);
    const b = sideOf(border.bottom);
    const l = sideOf(border.left);
    const r = sideOf(border.right);
    if (t || b || l || r) {
      borders = {};
      if (t) borders.top = t;
      if (b) borders.bottom = b;
      if (l) borders.left = l;
      if (r) borders.right = r;
    }
  }

  const font = cell.font as any;
  if (font) {
    if (font.bold) format.bold = true;
    if (font.italic) format.italic = true;
    if (font.underline) format.underline = true;
    if (font.strike) format.strike = true;
    if (font.size && typeof font.size === "number") format.font_size = font.size;
    if (font.name && typeof font.name === "string") format.font_family = font.name;
    if (font.color) {
      const c = resolveXlsxColor(font.color, themePalette);
      if (c) format.font_color = c;
    }
  }

  const alignment = cell.alignment as any;
  if (alignment) {
    if (alignment.horizontal === "left" || alignment.horizontal === "center" || alignment.horizontal === "right") {
      format.horizontal_align = alignment.horizontal;
    }
    if (alignment.vertical === "top" || alignment.vertical === "middle" || alignment.vertical === "bottom") {
      format.vertical_align = alignment.vertical;
    }
    if (alignment.wrapText) format.wrap_text = true;
    if (typeof alignment.indent === "number" && alignment.indent > 0) {
      format.indent = alignment.indent;
    }
  }

  // numFmt can be a string like "$#,##0.00" or "0.00%" — pass through verbatim.
  // ExcelJS exposes the resolved string for built-in formats too.
  const nf = cell.numFmt;
  if (nf && typeof nf === "string" && nf !== "General") {
    format.number_format = nf;
  }

  const fill = cell.fill as any;
  if (fill && fill.type === "pattern" && fill.pattern === "solid") {
    const c = resolveXlsxColor(fill.fgColor, themePalette);
    if (c) background = c;
  }

  return { format, background, borders };
}

/**
 * An anchored picture lifted out of the ExcelJS media store into a plain
 * cloneable shape, so image rendering doesn't need the ExcelJS Workbook on
 * the main thread.
 */
export type SheetImage = {
  sheetName: string;
  extension: string;
  buffer: ArrayBuffer;
  col: number;
  row: number;
};

export function extractWorksheetImages(excelJs: ExcelJS.Workbook): SheetImage[] {
  const out: SheetImage[] = [];
  for (const ws of excelJs.worksheets) {
    let images: any[] = [];
    try {
      images = (ws as any).getImages?.() ?? [];
    } catch {
      continue;
    }
    for (const img of images) {
      const media = (excelJs as any).getImage?.(Number(img.imageId));
      if (!media?.buffer) continue;
      const src: Uint8Array = media.buffer;
      // Copy out of any shared backing store so the result is independently
      // transferable/cloneable.
      const buffer = src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength) as ArrayBuffer;
      out.push({
        sheetName: ws.name,
        extension: String(media.extension || "png"),
        buffer,
        col: Math.max(0, Math.floor(Number(img.range?.tl?.col ?? 0))),
        row: Math.max(0, Math.floor(Number(img.range?.tl?.row ?? 0))),
      });
    }
  }
  return out;
}

/**
 * Everything the grid needs from a parse, in structured-cloneable form.
 * `excelJs` is populated only when the parse ran in-process (worker
 * unavailable or failed); when null, callers that need the ExcelJS workbook
 * (full-export save) re-parse lazily from the retained source bytes.
 */
export type ParsedWorkbook = {
  excelJs: ExcelJS.Workbook | null;
  univerData: any;
  styleOps: StyleOp[];
  outlines: Record<string, SheetOutline>;
  featureCounts: { cf: number; dv: number; notes: number; cfDropped: string[] };
  pins: Array<[string, { f: string; v: any }]>;
  images: SheetImage[];
};

/**
 * Parse xlsx bytes off the main thread when possible. The heavy ExcelJS
 * parse + conversion runs in a dedicated one-shot worker so the UI stays
 * responsive (large models take multiple seconds of pure CPU); the worker is
 * terminated as soon as it answers, releasing the ExcelJS object graph.
 * Any worker failure falls back to the identical in-process code path.
 */
export async function parseXlsxWorkbook(bytes: Uint8Array): Promise<ParsedWorkbook> {
  if (typeof Worker !== "undefined") {
    try {
      return await parseInWorker(bytes);
    } catch (e) {
      console.warn("[import] worker parse failed; falling back to main-thread parse:", e);
    }
  }
  const { excelJs, univerData, styleOps, outlines, featureCounts } = await xlsxBytesToWorkbook(bytes);
  return {
    excelJs,
    univerData,
    styleOps,
    outlines,
    featureCounts,
    pins: Array.from(externalPinsByWorkbook.get(excelJs) ?? []),
    images: extractWorksheetImages(excelJs),
  };
}

function parseInWorker(bytes: Uint8Array): Promise<ParsedWorkbook> {
  return new Promise((resolve, reject) => {
    let worker: Worker;
    try {
      worker = new Worker(new URL("./xlsxImport.worker.ts", import.meta.url), {
        type: "module",
        name: "xlsx-import",
      });
    } catch (e) {
      reject(e);
      return;
    }
    const settle = (fn: () => void) => {
      worker.terminate();
      fn();
    };
    worker.onerror = (ev) => settle(() => reject(ev.error ?? new Error(ev.message || "xlsx import worker error")));
    worker.onmessage = (ev) => {
      const msg = ev.data;
      if (msg?.ok) {
        settle(() => resolve({ excelJs: null, ...msg.result }));
      } else {
        settle(() => reject(new Error(String(msg?.error ?? "xlsx import worker: unknown failure"))));
      }
    };
    // Clone rather than transfer: the caller keeps using `bytes` for the
    // fidelity scan and as the save-time re-parse source.
    worker.postMessage({ bytes });
  });
}
