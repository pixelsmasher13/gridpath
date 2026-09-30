/**
 * Defined-name preservation for the full-export (ExcelJS) fallback path.
 *
 * ExcelJS destroys defined names whose text is not a cell range at PARSE
 * time: its extractRanges filters each comma-separated segment through
 * colCache.decodeEx and silently drops anything that throws. Real models
 * (Canalyst et al.) carry thousands of names, many of which are string
 * constants ("TRUE", version tags…) used as vendor metadata flags. After
 * the round-trip those names either disappear or — worse — get written as
 * empty <definedName/> elements, which is schema-invalid and makes Excel
 * run its repair flow on the saved file.
 *
 * The xform classes aren't reachable through ExcelJS's bundled public API,
 * so the fix operates on the exported package instead: compare the
 * exported xl/workbook.xml against the pristine source bytes (which we
 * keep for the loss audit anyway) and
 *   1. restore the original text of entries the parser emptied or
 *      partially stripped (exported segments ⊂ original segments),
 *   2. graft back dropped constants — names whose text carries no sheet
 *      reference can't be affected by any edit we support, so re-adding
 *      them verbatim is always safe,
 *   3. drop any still-empty elements so the output is schema-valid.
 * Names the user intentionally redefined this session are exempt.
 */

import JSZip from "jszip";

const WORKBOOK_PART = "xl/workbook.xml";

type NameEntry = {
  /** Identity: name + scope. localSheetId is kept as the literal attribute. */
  key: string;
  name: string;
  text: string; // raw inner XML (still entity-escaped)
  raw: string; // the full element, verbatim
  start: number;
  end: number;
};

export type DefinedNameRepair = {
  bytes: Uint8Array;
  /** Entries whose original text was restored over a gutted export. */
  restored: string[];
  /** Constant entries the export dropped entirely and we re-added. */
  grafted: string[];
  /** Empty entries with no original counterpart, removed for validity. */
  dropped: string[];
};

const ENTRY_RE = /<definedName\b([^>]*?)(?:\/>|>([\s\S]*?)<\/definedName>)/g;

function attrOf(attrs: string, name: string): string | null {
  const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attrs);
  return m ? m[1] : null;
}

function xmlDecode(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function parseEntries(workbookXml: string): NameEntry[] {
  const out: NameEntry[] = [];
  ENTRY_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = ENTRY_RE.exec(workbookXml))) {
    const attrs = m[1] ?? "";
    const name = xmlDecode(attrOf(attrs, "name") ?? "");
    if (!name) continue;
    const localSheetId = attrOf(attrs, "localSheetId");
    out.push({
      key: `${name}\u0000${localSheetId ?? ""}`,
      name,
      text: m[2] ?? "",
      raw: m[0],
      start: m.index,
      end: m.index + m[0].length,
    });
  }
  return out;
}

/**
 * True when `exported`'s comma-separated segments are a strict subset of
 * `original`'s — the signature of the parser dropping segments it couldn't
 * decode (any intentional edit produces segments the original didn't have).
 * Quoted sheet names containing commas mis-split and simply fail the
 * check, which errs toward leaving the export untouched.
 */
function isStrictSubset(exported: string, original: string): boolean {
  const seg = (s: string) =>
    xmlDecode(s)
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean);
  const exp = seg(exported);
  const orig = new Set(seg(original));
  return exp.length > 0 && exp.length < orig.size && exp.every((s) => orig.has(s));
}

async function readWorkbookXml(zip: JSZip): Promise<string | null> {
  const f = zip.file(WORKBOOK_PART);
  if (!f) return null;
  try {
    return await f.async("string");
  } catch {
    return null;
  }
}

/**
 * Repair the defined names of `exportedBytes` against `sourceBytes`.
 * Returns null when either package can't be read (caller keeps the export
 * as-is). `editedNames` — names the user redefined this session — are
 * never touched.
 */
export async function repairDefinedNames(
  sourceBytes: Uint8Array,
  exportedBytes: Uint8Array,
  editedNames: ReadonlySet<string> = new Set(),
): Promise<DefinedNameRepair | null> {
  let sourceZip: JSZip;
  let exportZip: JSZip;
  try {
    [sourceZip, exportZip] = await Promise.all([
      JSZip.loadAsync(sourceBytes),
      JSZip.loadAsync(exportedBytes),
    ]);
  } catch {
    return null;
  }
  const sourceXml = await readWorkbookXml(sourceZip);
  const exportXml = await readWorkbookXml(exportZip);
  if (!sourceXml || !exportXml) return null;

  const sourceEntries = parseEntries(sourceXml);
  const exportEntries = parseEntries(exportXml);
  if (!sourceEntries.length) return { bytes: exportedBytes, restored: [], grafted: [], dropped: [] };

  const byKey = new Map(sourceEntries.map((e) => [e.key, e]));
  const byName = new Map<string, NameEntry[]>();
  for (const e of sourceEntries) {
    const list = byName.get(e.name) ?? [];
    list.push(e);
    byName.set(e.name, list);
  }
  const findOriginal = (e: NameEntry): NameEntry | undefined => {
    const exact = byKey.get(e.key);
    if (exact) return exact;
    // Sheet indexes can shift across a structural fallback export; fall
    // back to the name alone when it's unambiguous.
    const candidates = byName.get(e.name);
    return candidates?.length === 1 ? candidates[0] : undefined;
  };

  const restored: string[] = [];
  const dropped: string[] = [];
  const grafted: string[] = [];

  // Rewrite existing entries in place (iterate backwards so offsets hold).
  let xml = exportXml;
  for (let i = exportEntries.length - 1; i >= 0; i--) {
    const e = exportEntries[i];
    if (editedNames.has(e.name)) continue;
    const orig = findOriginal(e);
    let replacement: string | null = null;
    // NB: replacer functions, not strings — the ranges contain `$1` etc.,
    // which a string replacement would treat as capture-group references.
    const withText = (text: string) =>
      e.raw.endsWith("/>")
        ? `${e.raw.slice(0, -2)}>${text}</definedName>`
        : e.raw.replace(/>[\s\S]*?<\/definedName>$/, () => `>${text}</definedName>`);
    if (!e.text.trim()) {
      if (orig?.text.trim()) {
        replacement = withText(orig.text);
        restored.push(e.name);
      } else {
        replacement = ""; // empty with no original counterpart: invalid, drop
        dropped.push(e.name);
      }
    } else if (orig && orig.text !== e.text && isStrictSubset(e.text, orig.text)) {
      replacement = withText(orig.text);
      restored.push(e.name);
    }
    if (replacement !== null) {
      xml = xml.slice(0, e.start) + replacement + xml.slice(e.end);
    }
  }

  // Graft back dropped constants: no sheet reference in the text means no
  // structural edit can invalidate it, so verbatim re-insertion is safe.
  const exportedNames = new Set(exportEntries.map((e) => e.name));
  const graftedRaw: string[] = [];
  for (const orig of sourceEntries) {
    if (exportedNames.has(orig.name) || editedNames.has(orig.name)) continue;
    const text = xmlDecode(orig.text).trim();
    if (!text || text.includes("!")) continue;
    graftedRaw.push(orig.raw);
    grafted.push(orig.name);
  }
  if (graftedRaw.length) {
    const block = graftedRaw.join("");
    if (xml.includes("</definedNames>")) {
      xml = xml.replace("</definedNames>", () => `${block}</definedNames>`);
    } else if (xml.includes("</sheets>")) {
      xml = xml.replace("</sheets>", () => `</sheets><definedNames>${block}</definedNames>`);
    } else {
      grafted.length = 0; // nowhere safe to put them
    }
  }

  if (!restored.length && !dropped.length && !grafted.length) {
    return { bytes: exportedBytes, restored, grafted, dropped };
  }

  exportZip.file(WORKBOOK_PART, xml);
  const bytes = await exportZip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
  });
  return { bytes, restored, grafted, dropped };
}

/**
 * Names present in the source but absent from the exported package — the
 * loss that remains after repair. Reported alongside the missing-parts
 * audit so defined-name loss is never silent.
 */
export async function auditDefinedNameLoss(
  sourceBytes: Uint8Array,
  exportedBytes: Uint8Array,
): Promise<string[] | null> {
  let sourceZip: JSZip;
  let exportZip: JSZip;
  try {
    [sourceZip, exportZip] = await Promise.all([
      JSZip.loadAsync(sourceBytes),
      JSZip.loadAsync(exportedBytes),
    ]);
  } catch {
    return null;
  }
  const sourceXml = await readWorkbookXml(sourceZip);
  const exportXml = await readWorkbookXml(exportZip);
  if (!sourceXml || !exportXml) return null;
  const exported = new Set(parseEntries(exportXml).map((e) => e.name));
  const missing: string[] = [];
  for (const e of parseEntries(sourceXml)) {
    if (!exported.has(e.name) && !missing.includes(e.name)) missing.push(e.name);
  }
  return missing;
}
