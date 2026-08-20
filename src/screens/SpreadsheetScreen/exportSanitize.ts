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
    if (!BAD_OPERATOR_RE.test(xml)) continue;
    BAD_OPERATOR_RE.lastIndex = 0;
    zip.file(name, xml.replace(BAD_OPERATOR_RE, ""));
    dirty = true;
  }
  if (!dirty) return bytes;
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
