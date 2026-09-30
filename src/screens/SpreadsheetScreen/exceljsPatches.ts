/**
 * Targeted runtime patches for ExcelJS 4.4 bugs that break real workbooks.
 * Applied once, at module scope, before any workbook is loaded.
 *
 * 1. Unparseable drawing parts crash the whole load.
 *    Charts can carry a `userShapes` part (`<c:userShapes>`, chartml
 *    namespace) — annotations drawn on top of a chart. ExcelJS's drawing
 *    xform only understands `<xdr:wsDr>` and parses anything else to
 *    `undefined`, then its reconcile step dereferences `.anchors` on it:
 *    "TypeError: undefined is not an object (evaluating 'drawing.anchors')".
 *    Fix: drop undefined drawing entries before reconcile. The part itself
 *    is untouched on disk (surgical save copies it raw).
 *
 * 2. Thousands of defined names OOM the renderer.
 *    ExcelJS explodes every named range into per-cell matrix entries.
 *    Measured on a real research model: 2,752 names turned a 154 ms / 44 MB
 *    load into 13 s / 5.7 GB — instant OOM in a webview. Fix: past a
 *    threshold, store the parsed name list raw and skip the matrix. The
 *    model getter merges matrix-derived names (runtime edits — they win)
 *    with the raw list, so serialization round-trips all names. Bonus: raw
 *    names with formula refs (e.g. `INDIRECT(...)`), which stock ExcelJS
 *    silently drops, survive too.
 *
 *    Trade-off: for such workbooks, ExcelJS's row/column-splice name
 *    shifting no longer applies — acceptable, since structure edits go
 *    through the (warned) fallback path and the alternative was a crash.
 *
 * 3. Zero-byte media parts abort the whole save.
 *    Some generators write empty image parts (seen in the wild: a 0-byte
 *    image3.emf / image4.png). Excel tolerates them; ExcelJS loads them
 *    with no buffer and its addMedia throws "Unsupported media", failing
 *    the entire export. Fix: give unwritable entries an empty buffer so
 *    they round-trip as the 0-byte parts they were. Indices are untouched,
 *    so drawing→media references stay valid.
 */

import ExcelJS from "exceljs";

/** Explode names into the per-cell matrix only below this count. */
const DEFINED_NAMES_MATRIX_LIMIT = 200;

let applied = false;

export function applyExcelJsPatches(): void {
  if (applied) return;
  applied = true;

  const probe = new ExcelJS.Workbook();

  // --- 1. reconcile guard ---
  const xlsxProto = Object.getPrototypeOf(probe.xlsx) as any;
  const origReconcile = xlsxProto.reconcile;
  xlsxProto.reconcile = function (model: any, options: any) {
    if (model?.drawings) {
      for (const key of Object.keys(model.drawings)) {
        if (!model.drawings[key]) {
          // STUB, don't delete: a worksheet that references this drawing
          // dereferences `options.drawings[name].anchors` during its own
          // reconcile — deleting the entry just moves the crash there
          // (seen with chart-only drawings, e.g. the eval fixture). Empty
          // anchors parse as "drawing with nothing in it".
          console.warn(`[exceljs] stubbing unparseable drawing part: ${key}`);
          model.drawings[key] = { anchors: [] };
        }
      }
    }
    return origReconcile.call(this, model, options);
  };

  // --- 3. media write guard ---
  const origAddMedia = xlsxProto.addMedia;
  xlsxProto.addMedia = async function (zip: any, model: any) {
    const media: any[] = model?.media ?? [];
    const writable = (m: any) => m && m.type === "image" && (m.buffer || m.filename || m.base64);
    if (media.some((m) => !writable(m))) {
      const patched = media.map((m, i) => {
        if (writable(m)) return m;
        console.warn(
          `[exceljs] media entry ${i} (${m?.name ?? "?"}.${m?.extension ?? "?"}) has no content; writing it as an empty part`,
        );
        return {
          ...m,
          type: "image",
          name: m?.name ?? `image_placeholder_${i}`,
          extension: m?.extension ?? "png",
          buffer: new Uint8Array(0),
        };
      });
      return origAddMedia.call(this, zip, { ...model, media: patched });
    }
    return origAddMedia.call(this, zip, model);
  };

  // --- 2. defined-names explosion guard ---
  const dnProto = Object.getPrototypeOf(probe.definedNames) as any;
  const desc = Object.getOwnPropertyDescriptor(dnProto, "model");
  if (!desc?.get || !desc?.set) {
    console.warn("[exceljs] definedNames.model descriptor not found; skipping names patch");
    return;
  }
  Object.defineProperty(dnProto, "model", {
    configurable: true,
    get(this: any) {
      const derived: Array<{ name: string; ranges: string[] }> = desc.get!.call(this) ?? [];
      const raw: Array<{ name: string; ranges: string[] }> = this._rawNames ?? [];
      if (!raw.length) return derived;
      const overridden = new Set(derived.map((d) => d.name));
      // Zero-range entries serialize as EMPTY <definedName/> elements —
      // schema-invalid, and Excel's repair flow strips them (with the scary
      // dialog). They exist because ExcelJS's parser guts constant-valued
      // names ("TRUE") at load. Dropping them here keeps the output valid;
      // the save path separately restores the original text from the
      // source package (see definedNamePreserve.ts).
      return [
        ...derived,
        ...raw.filter((r) => !overridden.has(r.name) && (r.ranges?.length ?? 0) > 0),
      ];
    },
    set(this: any, value: Array<{ name: string; ranges: string[] }>) {
      if (Array.isArray(value) && value.length > DEFINED_NAMES_MATRIX_LIMIT) {
        console.warn(
          `[exceljs] ${value.length} defined names — storing raw (matrix explosion guard)`,
        );
        this._rawNames = value;
        this.matrixMap = {};
        return;
      }
      this._rawNames = [];
      desc.set!.call(this, value);
    },
  });
}
