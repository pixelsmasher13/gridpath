/**
 * What's in an .xlsx, and what a save changed, read straight from the zip:
 * the central directory lists every part with the CRC-32 of its contents,
 * so comparing two files part by part needs no decompression at all.
 * Formula and name counts read the sheet and workbook XML.
 */

export interface ZipPart {
  name: string;
  crc: number;
  size: number;
  compressedSize: number;
  method: number;
  offset: number;
}

/** The parts of a zip archive, from its central directory. */
export function zipParts(bytes: Uint8Array): ZipPart[] {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("not a zip archive");
  let count = dv.getUint16(eocd + 10, true);
  let at = dv.getUint32(eocd + 16, true);
  // Zip64: the real count and offset are in the zip64 end record.
  if ((at === 0xffffffff || count === 0xffff) && eocd >= 20 && dv.getUint32(eocd - 20, true) === 0x07064b50) {
    const z = Number(dv.getBigUint64(eocd - 12, true));
    count = Number(dv.getBigUint64(z + 32, true));
    at = Number(dv.getBigUint64(z + 48, true));
  }
  const parts: ZipPart[] = [];
  const decoder = new TextDecoder();
  for (let n = 0; n < count && dv.getUint32(at, true) === 0x02014b50; n++) {
    const nameLen = dv.getUint16(at + 28, true);
    const extraLen = dv.getUint16(at + 30, true);
    const commentLen = dv.getUint16(at + 32, true);
    const name = decoder.decode(bytes.subarray(at + 46, at + 46 + nameLen));
    // Folder entries (some zip writers add them) aren't parts.
    if (!name.endsWith("/")) parts.push({
      name,
      method: dv.getUint16(at + 10, true),
      crc: dv.getUint32(at + 16, true),
      compressedSize: dv.getUint32(at + 20, true),
      size: dv.getUint32(at + 24, true),
      offset: dv.getUint32(at + 42, true),
    });
    at += 46 + nameLen + extraLen + commentLen;
  }
  return parts;
}

/** One part's contents, decompressed as a stream of chunks. */
async function* partChunks(bytes: Uint8Array, part: ZipPart): AsyncGenerator<Uint8Array> {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const start = part.offset + 30 + dv.getUint16(part.offset + 26, true) + dv.getUint16(part.offset + 28, true);
  const data = bytes.subarray(start, start + part.compressedSize);
  if (part.method === 0) {
    yield data;
    return;
  }
  if (part.method !== 8 || typeof DecompressionStream === "undefined") return;
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const reader = stream.getReader();
  for (;;) {
    const { value, done } = await reader.read();
    if (done) return;
    yield value;
  }
}

/** Occurrences of a pattern in a part's text, read chunk by chunk. */
async function countIn(bytes: Uint8Array, part: ZipPart, pattern: RegExp, overlap: number): Promise<number> {
  const decoder = new TextDecoder();
  let tail = "";
  let n = 0;
  for await (const chunk of partChunks(bytes, part)) {
    const text = tail + decoder.decode(chunk, { stream: true });
    // Matches that end inside the carried-over tail were counted already.
    for (const m of text.matchAll(pattern)) if (m.index! + m[0].length > tail.length) n++;
    tail = text.slice(-overlap);
  }
  return n;
}

export interface FileContents {
  parts: number;
  sheets: number;
  formulas: number | null;
  names: number | null;
  charts: number;
  pivotTables: number;
  tables: number;
  images: number;
  comments: number;
  macros: boolean;
  externalLinks: number;
  customXml: number;
}

const count = (parts: ZipPart[], re: RegExp) => parts.filter((p) => re.test(p.name)).length;

/** What the file carries: the cheap counts at once, formulas and names after reading the XML. */
export function fileContents(bytes: Uint8Array): FileContents {
  const parts = zipParts(bytes);
  return {
    parts: parts.length,
    sheets: count(parts, /^xl\/worksheets\/[^/]+\.xml$/),
    formulas: null,
    names: null,
    charts: count(parts, /^xl\/(?:drawings\/)?charts\/chart\d+\.xml$/),
    pivotTables: count(parts, /^xl\/pivotTables\/[^/]+\.xml$/),
    tables: count(parts, /^xl\/tables\/[^/]+\.xml$/),
    images: count(parts, /^xl\/media\//),
    comments: count(parts, /^xl\/(comments\d*\.xml|threadedComments\/)/),
    macros: parts.some((p) => p.name === "xl/vbaProject.bin"),
    externalLinks: count(parts, /^xl\/externalLinks\/[^/]+\.xml$/),
    customXml: count(parts, /^customXml\/item\d+\.xml$/),
  };
}

/** Formula cells (every `<f>` in the sheets) and defined names. */
export async function countFormulasAndNames(bytes: Uint8Array): Promise<{ formulas: number; names: number } | null> {
  if (typeof DecompressionStream === "undefined") return null;
  const parts = zipParts(bytes);
  let formulas = 0;
  for (const p of parts.filter((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p.name))) {
    formulas += await countIn(bytes, p, /<(?:\w+:)?f[\s>/]/g, 8);
  }
  const wb = parts.find((p) => p.name === "xl/workbook.xml");
  const names = wb ? await countIn(bytes, wb, /<(?:\w+:)?definedName[\s>]/g, 16) : 0;
  return { formulas, names };
}

export interface SaveComparison {
  total: number;
  unchanged: number;
  rewritten: string[];
  added: string[];
  removed: string[];
}

/** Which parts of `original` came back with the same contents in `saved`. */
export function compareSaves(original: Uint8Array, saved: Uint8Array): SaveComparison {
  const before = zipParts(original);
  const after = new Map(zipParts(saved).map((p) => [p.name, p]));
  const out: SaveComparison = { total: before.length, unchanged: 0, rewritten: [], added: [], removed: [] };
  for (const p of before) {
    const q = after.get(p.name);
    if (!q) out.removed.push(p.name);
    else if (q.crc === p.crc && q.size === p.size) out.unchanged++;
    else out.rewritten.push(p.name);
    after.delete(p.name);
  }
  out.added = [...after.keys()];
  return out;
}

/** Why a part usually changes on save, for the report. */
export function partNote(name: string): string | null {
  if (/^xl\/worksheets\//.test(name)) return "a sheet you edited";
  if (name === "xl/calcChain.xml") return "Excel rebuilds it on open";
  if (name === "xl/workbook.xml") return "sheets, names or the recalculation flag";
  if (name === "xl/styles.xml") return "formats you added";
  if (name === "xl/sharedStrings.xml") return "text you typed";
  if (name === "[Content_Types].xml" || /_rels\//.test(name)) return "the package's index of parts";
  if (/^docProps\//.test(name)) return "document properties";
  return null;
}
