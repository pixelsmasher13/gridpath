import ExcelJS from "exceljs";
import JSZip from "jszip";
import { readWorkbookBytes } from "../workbookIo";
import { cellFromExcelJs } from "../xlsxImport";
import { expandA1Range } from "./toolToMutation";
import type { ReferenceWorkbookContext } from "./agentClient";

/**
 * Reference workbooks — other xlsx files the user attaches to a session as
 * READ-ONLY context (e.g. analyst models to compare against while building a
 * new one). They are never opened in a Univer tab: we parse the file with
 * ExcelJS once, keep only values + formulas in a sparse in-memory map, and
 * serve both the compact upfront preview and on-demand `read_reference`
 * tool calls from that cache.
 */

export type ReferenceCell = { value: unknown; formula: string | null };

export type ParsedReferenceSheet = {
  name: string;
  rowCount: number;
  columnCount: number;
  /** Sparse cell map keyed by "row:col" (0-indexed). Non-empty cells only. */
  cells: Map<string, ReferenceCell>;
};

export type ParsedReference = {
  path: string;
  /** Display label the agent addresses this workbook by — the filename. */
  label: string;
  sheets: ParsedReferenceSheet[];
};

/** Max references attachable per session — keeps prompt size bounded. */
export const MAX_REFERENCES = 5;

/**
 * Cells per sheet in the upfront preview. Deliberately smaller than the
 * active workbook's 1500 — several attached analyst models would otherwise
 * dominate the prompt. The agent pulls detail on demand via read_reference.
 */
const MAX_CELLS_PER_REF_SHEET = 300;

/** Same per-call cap as the active workbook's read_range tool. */
const MAX_READ_CELLS = 500;

export function referenceLabelFromPath(path: string): string {
  return path.split("/").pop() ?? path;
}

// Module-level cache: parsing a large xlsx is expensive and reference files
// are read-only, so one parse per path per app run is enough. Removing the
// chip evicts the entry; re-attaching re-reads from disk.
const cache = new Map<string, ParsedReference>();

export function evictReference(path: string): void {
  cache.delete(path);
}

export function getCachedReference(path: string): ParsedReference | null {
  return cache.get(path) ?? null;
}

/**
 * Load + parse a reference workbook, serving from cache when possible.
 * Throws (with the underlying I/O or parse error) if the file is missing or
 * unreadable — callers surface that as an error chip and skip the file.
 */
export async function loadReferenceWorkbook(path: string): Promise<ParsedReference> {
  const hit = cache.get(path);
  if (hit) return hit;

  const bytes = await readWorkbookBytes(path);
  const excelJs = new ExcelJS.Workbook();
  // ExcelJS 4.4.0 crashes with "undefined is not an object (evaluating
  // 'drawing.anchors')" on workbooks whose drawing parts it can't fully
  // parse — chart-only drawings are the usual trigger, and analyst models
  // are full of charts. Since references only need values + formulas, we
  // remove chart/drawing parts from the zip entirely before parsing (the
  // workbook-level reconcile iterates model.drawings unconditionally, so an
  // ignoreNodes filter alone doesn't reach it). Costs ~100ms on a 1MB file,
  // paid once per attach thanks to the cache. The file on disk is untouched.
  const stripped = await stripHeavyParts(bytes);
  // Also skip worksheet XML nodes we don't need for a values+formulas read —
  // faster, and avoids the same anchors crash at the worksheet level for
  // sheets whose <drawing> rel points at a part ExcelJS didn't register.
  await excelJs.xlsx.load(stripped, {
    ignoreNodes: [
      "drawing",
      "picture",
      "dataValidations",
      "conditionalFormatting",
      "autoFilter",
      "sheetProtection",
      "tableParts",
      "headerFooter",
      "pageSetup",
      "pageMargins",
      "printOptions",
      "rowBreaks",
      "hyperlinks",
      "mergeCells",
      "extLst",
    ],
  });

  const sheets: ParsedReferenceSheet[] = excelJs.worksheets.map((ws) => {
    const cells = new Map<string, ReferenceCell>();
    let maxRow = 0;
    let maxCol = 0;
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const r = rowNumber - 1;
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const c = colNumber - 1;
        const d = cellFromExcelJs(cell);
        if (d.v === undefined && !d.f) return;
        cells.set(`${r}:${c}`, { value: d.v ?? null, formula: d.f ?? null });
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      });
    });
    return {
      name: ws.name,
      rowCount: maxRow + 1,
      columnCount: maxCol + 1,
      cells,
    };
  });

  const parsed: ParsedReference = {
    path,
    label: referenceLabelFromPath(path),
    sheets,
  };
  // The ExcelJS workbook (styles, drawings, themes) goes out of scope here —
  // only the sparse value/formula maps stay resident.
  cache.set(path, parsed);
  return parsed;
}

/**
 * Build the compact per-reference context block shipped upfront each turn.
 * Same "A1 = value" line format as captureWorkbookContext, tighter cap.
 */
export function captureReferenceContext(refs: ParsedReference[]): ReferenceWorkbookContext[] {
  return refs.map((ref) => ({
    path: ref.path,
    label: ref.label,
    sheets: ref.sheets.map((sheet) => {
      const entries = Array.from(sheet.cells.entries())
        .map(([key, cell]) => {
          const [r, c] = key.split(":").map(Number);
          return { r, c, cell };
        })
        .sort((a, b) => a.r - b.r || a.c - b.c);
      const lines: string[] = [];
      for (const { r, c, cell } of entries) {
        if (lines.length >= MAX_CELLS_PER_REF_SHEET) {
          lines.push("... (preview truncated — use read_reference for more)");
          break;
        }
        const addr = colLetters(c) + (r + 1);
        if (cell.formula) {
          if (cell.value !== undefined && cell.value !== null) {
            lines.push(`${addr} = ${cell.formula}  → ${JSON.stringify(cell.value)}`);
          } else {
            lines.push(`${addr} = ${cell.formula}`);
          }
        } else {
          lines.push(`${addr} = ${JSON.stringify(cell.value)}`);
        }
      }
      return {
        name: sheet.name,
        row_count: sheet.rowCount,
        column_count: sheet.columnCount,
        cells_preview: lines.length > 0 ? lines.join("\n") : "(empty)",
      };
    }),
  }));
}

/**
 * Serve a `read_reference` tool call from the cache. Matches the workbook by
 * label (filename, case-insensitive) or full path. Response mirrors
 * read_range's shape so the agent handles both identically.
 */
export function readReferenceRange(
  referencePaths: string[],
  workbook: string,
  sheetName: string,
  rangeA1: string,
): string {
  const wanted = workbook.trim().toLowerCase();
  const path = referencePaths.find(
    (p) =>
      p.toLowerCase() === wanted ||
      referenceLabelFromPath(p).toLowerCase() === wanted,
  );
  const ref = path ? cache.get(path) : null;
  if (!ref) {
    const available = referencePaths.map(referenceLabelFromPath);
    return JSON.stringify({
      error: `unknown reference workbook "${workbook}"`,
      available_references: available,
    });
  }
  const sheet = ref.sheets.find((s) => s.name === sheetName)
    ?? ref.sheets.find((s) => s.name.toLowerCase() === sheetName.trim().toLowerCase());
  if (!sheet) {
    return JSON.stringify({
      error: `unknown sheet "${sheetName}" in reference "${ref.label}"`,
      available_sheets: ref.sheets.map((s) => s.name),
    });
  }
  const rangeCells = expandA1Range(rangeA1);
  if (rangeCells.length === 0) {
    return JSON.stringify({ error: `bad A1 range "${rangeA1}"` });
  }
  const cells = rangeCells
    .slice(0, MAX_READ_CELLS)
    .map(({ row, col }) => {
      const c = sheet.cells.get(`${row}:${col}`);
      return {
        cell: colLetters(col) + (row + 1),
        value: c?.value ?? null,
        formula: c?.formula ?? null,
      };
    })
    .filter((c) => c.value !== null || c.formula !== null);
  return JSON.stringify({
    workbook: ref.label,
    sheet: sheet.name,
    range: rangeA1,
    cells,
    truncated: rangeCells.length > MAX_READ_CELLS,
  });
}

/**
 * Remove parts a values+formulas read never needs, so ExcelJS neither crashes
 * on them nor blows up memory. vmlDrawings (legacy comment shapes) are left
 * alone — ExcelJS handles those fine and comments can carry useful analyst
 * notes.
 *
 * - charts/drawings: ExcelJS 4.4.0 crashes reconciling chart-only drawings.
 * - externalLinks / pivot caches / calcChain: dead weight for a read-only
 *   reference; big sell-side models carry 100+ externalLink parts (MBs each).
 * - definedNames + externalReferences in workbook.xml: the real memory bomb.
 *   ExcelJS expands every defined name into a per-cell matrix, and external
 *   links mint thousands of hidden names — often full-column ranges that
 *   explode into millions of entries and OOM the webview (observed: a 4.3MB
 *   analyst model exhausting a 4GB heap; gutted, it parses in ~0.4s/150MB).
 */
async function stripHeavyParts(bytes: Uint8Array): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(bytes);
  const toRemove = Object.keys(zip.files).filter(
    (name) =>
      name.startsWith("xl/charts/") ||
      name.startsWith("xl/externalLinks/") ||
      name.startsWith("xl/pivotCache/") ||
      name.startsWith("xl/pivotTables/") ||
      name === "xl/calcChain.xml" ||
      /^xl\/drawings\/drawing\d+\.xml$/.test(name) ||
      /^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/.test(name),
  );
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const guttedXml = workbookXml
    ?.replace(/<definedNames>[\s\S]*?<\/definedNames>/, "")
    .replace(/<externalReferences>[\s\S]*?<\/externalReferences>/, "");
  const workbookChanged = guttedXml !== undefined && guttedXml !== workbookXml;
  if (toRemove.length === 0 && !workbookChanged) {
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  }
  for (const name of toRemove) zip.remove(name);
  if (workbookChanged && guttedXml !== undefined) {
    zip.file("xl/workbook.xml", guttedXml);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

function colLetters(col: number): string {
  let n = col;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}
