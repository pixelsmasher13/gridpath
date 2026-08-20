/// <reference lib="webworker" />
/**
 * One-shot xlsx parse worker. Receives the file bytes, runs the full
 * ExcelJS → Univer conversion (see xlsxImport.ts), and posts back only
 * structured-cloneable results. The ExcelJS Workbook object never leaves
 * this worker — the main thread re-parses lazily if a full-export save
 * needs it. The parent terminates this worker after the single reply.
 */
import {
  extractWorksheetImages,
  externalPinsByWorkbook,
  xlsxBytesToWorkbook,
} from "./xlsxImport";

self.onmessage = async (e: MessageEvent<{ bytes: Uint8Array }>) => {
  try {
    const { excelJs, univerData, styleOps, outlines, featureCounts } = await xlsxBytesToWorkbook(e.data.bytes);
    const result = {
      univerData,
      styleOps,
      outlines,
      featureCounts,
      pins: Array.from(externalPinsByWorkbook.get(excelJs) ?? []),
      images: extractWorksheetImages(excelJs),
    };
    (self as unknown as Worker).postMessage(
      { ok: true, result },
      // Image buffers can be large (logos, banner PNGs) — transfer, don't copy.
      result.images.map((img) => img.buffer),
    );
  } catch (err: any) {
    (self as unknown as Worker).postMessage({ ok: false, error: String(err?.stack ?? err) });
  }
};
