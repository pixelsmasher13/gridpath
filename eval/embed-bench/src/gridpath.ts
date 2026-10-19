import { createSpreadsheet } from "gridpath";
import { B, fail, inputs, toBase64, until } from "./common";

(async () => {
  const { id, bytes, prep } = await inputs();
  B.t.start = performance.now();
  const sheet = await createSpreadsheet(document.getElementById("app")!, { file: bytes, fileName: `${id}.xlsx` });
  B.t.interactive = performance.now();
  B.phase = "interactive";
  await sheet.whenCalculated(60 * 60_000);
  B.t.calculated = performance.now();
  const read = async () => (await sheet.getCell(prep.dep.sheet, prep.dep.ref))?.value;
  B.edit = async (value) => {
    const before = await read();
    const t0 = performance.now();
    sheet.setCell(prep.input.sheet, prep.input.ref, String(value));
    return until(read, before, t0);
  };
  // Same engine build as the Node scoring in prep.mjs.
  B.dump = async () => null;
  B.save = async () => toBase64(await sheet.save());
  B.phase = "calculated";
})().catch(fail);
