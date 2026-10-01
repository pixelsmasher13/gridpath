/**
 * Schema sanitation for ExcelJS full-export output.
 *
 * ExcelJS models the cfRule types containsText / containsBlanks /
 * notContainsBlanks / containsErrors / notContainsErrors as
 * `{type:'containsText', operator:<real type>}` and, on write, restores the
 * type from the operator but ALSO emits the operator attribute verbatim.
 * `operator="containsErrors"` (and friends) are not legal values of
 * ST_ConditionalFormattingOperator, so the written workbook is
 * schema-invalid — Excel opens it only through the repair flow, which
 * strips the rules. The xform classes aren't reachable through the bundled
 * public API, so strip the bogus attribute from the package instead.
 * These four values are invalid as operators in every context, so a global
 * strip cannot break a valid rule.
 */

import JSZip from "jszip";

const BAD_OPERATOR_RE =
  / operator="(containsBlanks|notContainsBlanks|containsErrors|notContainsErrors)"/g;

/**
 * CT_Worksheet child sequence (ECMA-376) — mirrors WORKSHEET_CHILD_ORDER in
 * src-tauri/.../xlsx_patch/sheet_xml.rs. ExcelJS appends `<legacyDrawing>`
 * (cell comments' VML anchor) AFTER `<tableParts>` on sheets that carry an
 * Excel table; the schema puts it before, and the Rust write boundary
 * rejects the export. Stable-sort the top-level children into schema order.
 */
const WORKSHEET_CHILD_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
  "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
  "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
  "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
  "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
  "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
  "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
  "controls", "webPublishItems", "tableParts", "extLst",
];

const TAG_RE = /<(\/?)([A-Za-z0-9_:]+)(?:\s[^>]*?)?(\/?)>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g;

/**
 * Reorders the direct children of `<worksheet>` into CT_Worksheet sequence.
 * Returns the input unchanged when already ordered or when the XML can't be
 * segmented confidently (unbalanced tags, no worksheet root).
 */
export function reorderWorksheetChildren(xml: string): string {
  const rootOpen = /<worksheet(?:\s[^>]*)?>/.exec(xml);
  if (!rootOpen) return xml;
  const bodyStart = rootOpen.index + rootOpen[0].length;
  const rootClose = xml.lastIndexOf("</worksheet>");
  if (rootClose < bodyStart) return xml;
  const body = xml.slice(bodyStart, rootClose);

  type Chunk = { name: string; text: string; order: number };
  const chunks: Chunk[] = [];
  let depth = 0;
  let chunkStart = -1;
  let chunkName = "";
  let lastEnd = 0;
  let lastKnown = -1;
  TAG_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = TAG_RE.exec(body))) {
    if (m[0].startsWith("<!")) continue; // comment / CDATA: opaque
    const isClose = m[1] === "/";
    const selfClose = m[3] === "/";
    const name = m[2].includes(":") ? m[2].slice(m[2].indexOf(":") + 1) : m[2];
    if (!isClose) {
      if (depth === 0) {
        if (body.slice(lastEnd, m.index).trim() !== "") return xml; // stray text at top level
        chunkStart = m.index;
        chunkName = name;
      }
      if (!selfClose) depth++;
    } else {
      depth--;
      if (depth < 0) return xml;
    }
    if (depth === 0 && (isClose || selfClose)) {
      if (isClose && name !== chunkName) return xml;
      const idx = WORKSHEET_CHILD_ORDER.indexOf(chunkName);
      if (idx >= 0) lastKnown = idx;
      chunks.push({ name: chunkName, text: body.slice(chunkStart, m.index + m[0].length), order: idx >= 0 ? idx : lastKnown });
      lastEnd = m.index + m[0].length;
    }
  }
  if (depth !== 0 || body.slice(lastEnd).trim() !== "") return xml;
  const sorted = chunks.slice().sort((a, b) => a.order - b.order); // Array.sort is stable
  if (sorted.every((c, i) => c === chunks[i])) return xml;
  return xml.slice(0, bodyStart) + sorted.map((c) => c.text).join("") + xml.slice(rootClose);
}

/**
 * Returns sanitized bytes, or the input unchanged when nothing needed
 * fixing (or the package couldn't be read — the caller keeps the export).
 */
export async function sanitizeExportedPackage(bytes: Uint8Array): Promise<Uint8Array> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return bytes;
  }
  let dirty = false;
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(name)) continue;
    const xml = await zip.file(name)!.async("string");
    let fixed = xml;
    BAD_OPERATOR_RE.lastIndex = 0;
    if (BAD_OPERATOR_RE.test(fixed)) {
      BAD_OPERATOR_RE.lastIndex = 0;
      fixed = fixed.replace(BAD_OPERATOR_RE, "");
    }
    fixed = reorderWorksheetChildren(fixed);
    if (fixed === xml) continue;
    zip.file(name, fixed);
    dirty = true;
  }
  if (!dirty) return bytes;
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
