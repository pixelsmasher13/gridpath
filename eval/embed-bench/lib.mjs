// Shared helpers for the embed benchmark: reading an .xlsx's own cached
// formula results (what Excel computed when it last saved), comparing
// values across engines, and diffing what a product's saved file kept.
import fs from "node:fs";
import JSZip from "jszip";

export const ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#CIRC!", "#ERROR!", "#N/IMPL!"]);

/** One comparable form for a cell value from any engine. */
export function norm(v) {
  if (v === undefined || v === null || v === "") return null;
  if (typeof v === "object" && v && "e" in v) return `#ERR`;
  if (typeof v === "string" && ERRORS.has(v.trim())) return `#ERR`;
  if (typeof v === "string" && /^#[A-Z/0!?_]+$/.test(v.trim())) return `#ERR`;
  return v;
}

export function same(a, b) {
  a = norm(a);
  b = norm(b);
  if (typeof a === "number" && typeof b === "number") {
    return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  }
  // A cached boolean and an engine's 1/0 differ; keep it strict.
  return a === b;
}

const unxml = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n))).replace(/&amp;/g, "&");

export function colIndex(letters) {
  let c = 0;
  for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c;
}
export function a1(r, c) {
  let s = "";
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return `${s}${r}`;
}

/** Sheet name → worksheet part path, in workbook order. */
async function sheetParts(zip) {
  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const target = new Map();
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1];
    const t = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && t) target.set(id, t.startsWith("/") ? t.slice(1) : `xl/${t}`);
  }
  const out = [];
  // Attribute values may contain ">" (a sheet named "Inputs ---->").
  for (const m of wbXml.matchAll(/<sheet\b(?:\s+[\w:]+="[^"]*")*\s*\/?>/g)) {
    const name = unxml(/name="([^"]*)"/.exec(m[0])[1]);
    const rid = /r:id="([^"]+)"/.exec(m[0])?.[1];
    out.push({ name, part: target.get(rid) });
  }
  return out;
}

async function sharedStrings(zip) {
  const f = zip.file("xl/sharedStrings.xml");
  if (!f) return [];
  const xml = await f.async("string");
  const out = [];
  for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    out.push(unxml([...m[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join("")));
  }
  return out;
}

/** Every formula cell that carries a cached result: [{ name, cells: [r, c, ...], values: [...] }]. */
export async function cachedFormulaValues(path) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path));
  const ss = await sharedStrings(zip);
  const out = [];
  for (const { name, part } of await sheetParts(zip)) {
    const xml = await zip.file(part).async("string");
    const cells = [];
    const values = [];
    for (const m of xml.matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const inner = m[4];
      if (!inner || !inner.includes("<f")) continue;
      const vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
      // No result, or an empty one written by a generator that never calculated.
      if (!vm || vm[1] === "") continue;
      const t = /\bt="([^"]+)"/.exec(m[3])?.[1] ?? "n";
      const raw = unxml(vm[1]);
      let v;
      if (t === "s") v = ss[Number(raw)];
      else if (t === "str" || t === "inlineStr") v = raw;
      else if (t === "b") v = raw === "1";
      else if (t === "e") v = { e: raw };
      else v = Number(raw);
      cells.push(Number(m[2]), colIndex(m[1]));
      values.push(v);
    }
    out.push({ name, cells, values });
  }
  return out;
}

/** What a product's saved file kept of the original package. */
export async function fidelity(origPath, outBytes) {
  const a = await JSZip.loadAsync(fs.readFileSync(origPath));
  const b = await JSZip.loadAsync(outBytes);
  const names = (z) => Object.keys(z.files).filter((n) => !z.files[n].dir);
  const an = names(a);
  const bn = new Set(names(b));
  let identical = 0;
  const changed = [];
  const dropped = [];
  for (const n of an) {
    if (!bn.has(n)) {
      dropped.push(n);
      continue;
    }
    const [x, y] = await Promise.all([a.file(n).async("uint8array"), b.file(n).async("uint8array")]);
    if (x.length === y.length && x.every((v, i) => v === y[i])) identical++;
    else changed.push(n);
  }
  const added = [...bn].filter((n) => !an.includes(n));
  // Feature parts by kind, so a renamed part is not counted as lost.
  const kinds = { charts: /^xl\/charts\/chart\d*\.xml$/, drawings: /^xl\/drawings\/drawing\d*\.xml$/, vml: /\.vml$/, comments: /^xl\/comments\d*\.xml$|threadedComments/, customXml: /^customXml\/item\d*\.xml$/, externalLinks: /^xl\/externalLinks\//, pivots: /^xl\/pivot/, vba: /vbaProject\.bin$/, printerSettings: /printerSettings/, tables: /^xl\/tables\//, media: /^xl\/media\// };
  const features = {};
  for (const [k, re] of Object.entries(kinds)) {
    const before = an.filter((n) => re.test(n)).length;
    if (before) features[k] = { before, after: [...bn].filter((n) => re.test(n)).length };
  }
  const definedNames = async (z) => ((await z.file("xl/workbook.xml")?.async("string")) ?? "").match(/<definedName\b/g)?.length ?? 0;
  return { parts: an.length, identical, changed, dropped, added, features, definedNames: { before: await definedNames(a), after: await definedNames(b) } };
}
