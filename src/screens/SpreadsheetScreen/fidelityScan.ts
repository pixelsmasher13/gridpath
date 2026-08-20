/**
 * Content-aware fidelity risk scan.
 *
 * The ExcelJS full-export fallback only writes what ExcelJS models, so any
 * package part outside that model is silently dropped. At load time we scan
 * the raw package for the part families known to be at risk; the save flow
 * uses the result to decide between a silent safe fallback (plain files)
 * and a preservation-first block that names exactly what would be lost.
 */

export interface FidelityRisks {
  charts: number;
  /** Individual comments (not parts) so the warning can say "14 comments". */
  comments: number;
  externalLinks: number;
  pivotTables: number;
  webExtensions: number;
  customXml: number;
  printerSettings: number;
  tables: number;
  slicers: number;
  drawings: number;
  total: number;
}

export async function scanFidelityRisks(bytes: Uint8Array): Promise<FidelityRisks | null> {
  try {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(bytes);
    const names = Object.keys(zip.files);
    const count = (re: RegExp) => names.filter((n) => re.test(n)).length;

    let comments = 0;
    for (const n of names) {
      if (!/^xl\/comments[^/]*\.xml$/i.test(n)) continue;
      const xml = await zip.file(n)?.async("string");
      comments += xml ? (xml.match(/<comment[\s>]/g)?.length ?? 0) : 0;
    }

    const risks = {
      charts: count(/^xl\/charts\/chart[^/]*\.xml$/i),
      comments,
      externalLinks: count(/^xl\/externalLinks\/externalLink[^/]*\.xml$/i),
      pivotTables: count(/^xl\/pivotTables\/[^/]+\.xml$/i),
      webExtensions: count(/^xl\/webextensions\/[^/]+/i),
      customXml: count(/^customXml\/item[^/]*\.xml$/i),
      printerSettings: count(/^xl\/printerSettings\/[^/]+/i),
      tables: count(/^xl\/tables\/[^/]+\.xml$/i),
      slicers: count(/^xl\/slicers\/[^/]+/i),
      drawings: count(/^xl\/drawings\/drawing[^/]*\.xml$/i),
    };
    const total = Object.values(risks).reduce((a, b) => a + b, 0);
    return { ...risks, total };
  } catch (e) {
    console.warn("[fidelity] risk scan failed:", e);
    return null;
  }
}

/**
 * The gate decision for a save that could not go through the surgical path.
 * - Untitled or content-plain workbooks: full export, silently — there is
 *   nothing meaningful to lose.
 * - At-risk workbook, in place: BLOCK. The lossy writer must never
 *   overwrite the only copy of content it cannot preserve.
 * - At-risk workbook, Save As: allowed, with a warning naming the losses —
 *   the original file stays intact.
 */
export type FallbackGate = "export" | "export_with_warning" | "block";

export function gateFallbackSave(opts: {
  isUntitled: boolean;
  isSaveAs: boolean;
  risks: FidelityRisks | null;
}): FallbackGate {
  const atRisk = !opts.isUntitled && !!opts.risks && opts.risks.total > 0;
  if (!atRisk) return "export";
  return opts.isSaveAs ? "export_with_warning" : "block";
}

/** "2 charts, 14 comments, 1 external link" — for warning copy. */
export function describeFidelityRisks(r: FidelityRisks): string {
  const parts: string[] = [];
  const add = (n: number, singular: string, plural = `${singular}s`) => {
    if (n > 0) parts.push(`${n} ${n === 1 ? singular : plural}`);
  };
  add(r.charts, "chart");
  add(r.comments, "comment");
  add(r.externalLinks, "external workbook link");
  add(r.pivotTables, "pivot table");
  add(r.tables, "Excel table");
  add(r.slicers, "slicer part");
  add(r.webExtensions, "add-in part");
  add(r.customXml, "custom XML part");
  add(r.printerSettings, "printer-settings part", "printer-settings parts");
  add(r.drawings, "drawing");
  return parts.join(", ");
}

/**
 * Post-export audit: which package parts of the source are missing from the
 * exported copy? Reported to the user so a reduced-fidelity copy names its
 * losses precisely. Parts are NOT grafted back into the export — a part
 * whose relationships were rewritten by ExcelJS can't be proven independent,
 * and a wrong graft corrupts the copy.
 */
/** Group raw missing part names into readable categories for warning copy. */
export function describeMissingParts(missing: string[]): string {
  const buckets: Array<[RegExp, string, string]> = [
    [/^xl\/charts\//i, "chart part", "chart parts"],
    [/^xl\/comments/i, "comment part", "comment parts"],
    [/^xl\/externalLinks\//i, "external-link part", "external-link parts"],
    [/^xl\/pivot/i, "pivot part", "pivot parts"],
    [/^xl\/webextensions\//i, "add-in part", "add-in parts"],
    [/^customXml\//i, "custom XML part", "custom XML parts"],
    [/^xl\/printerSettings\//i, "printer-settings part", "printer-settings parts"],
    [/^xl\/tables\//i, "table part", "table parts"],
    [/^xl\/slicer/i, "slicer part", "slicer parts"],
    [/^xl\/drawings\//i, "drawing part", "drawing parts"],
    [/^xl\/media\//i, "media file", "media files"],
  ];
  const counts = new Map<string, { n: number; plural: string }>();
  let other = 0;
  for (const name of missing) {
    const hit = buckets.find(([re]) => re.test(name));
    if (hit) {
      const cur = counts.get(hit[1]) ?? { n: 0, plural: hit[2] };
      cur.n += 1;
      counts.set(hit[1], cur);
    } else {
      other += 1;
    }
  }
  const parts = [...counts.entries()].map(
    ([singular, { n, plural }]) => `${n} ${n === 1 ? singular : plural}`,
  );
  if (other > 0) parts.push(`${other} other ${other === 1 ? "part" : "parts"}`);
  return parts.join(", ");
}

export async function auditExportLoss(
  source: Uint8Array,
  exported: Uint8Array,
): Promise<string[] | null> {
  try {
    const JSZip = (await import("jszip")).default;
    const src = await JSZip.loadAsync(source);
    const out = await JSZip.loadAsync(exported);
    const missing = Object.keys(src.files).filter(
      (n) => !src.files[n].dir && !out.files[n] && !/\/_rels\/|^_rels\//.test(n),
    );
    return missing;
  } catch (e) {
    console.warn("[fidelity] export audit failed:", e);
    return null;
  }
}
