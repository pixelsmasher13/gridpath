/**
 * Headless calc-engine benchmark. Boots Univer (core + sheets + formula
 * engine, no UI, no worker) in Node, loads a real xlsx via the app's own
 * ExcelJS import path, and measures:
 *
 *   1. initial full calculation (what the user waits for at load when the
 *      file has no cached values),
 *   2. recalc latency after editing one literal input cell,
 *   3. recalc latency after editing one formula cell.
 *
 * Set `pro: true` to swap in @univerjs-pro/engine-formula (license-free
 * evaluation) and compare. Run via scripts/calcBench.test.ts.
 */
import fs from "node:fs";
import { performance } from "node:perf_hooks";
import ExcelJS from "exceljs";

import { cellFromExcelJs } from "../src/screens/SpreadsheetScreen/components/UniverGrid";

type CellData = Record<number, Record<number, any>>;

// Vitest intercepts console output; mirror every line to a file so results
// survive regardless of reporter settings.
const OUT_FILE = "/tmp/calc-bench-out.txt";
function log(line: string): void {
  console.log(line);
  fs.appendFileSync(OUT_FILE, line + "\n");
}

interface LoadedSheet {
  name: string;
  cellData: CellData;
  rowCount: number;
  columnCount: number;
  formulaCount: number;
}

async function loadXlsx(path: string): Promise<LoadedSheet[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fs.readFileSync(path).buffer as ArrayBuffer);
  const sheets: LoadedSheet[] = [];
  wb.eachSheet((ws) => {
    const cellData: CellData = {};
    let maxRow = 0;
    let maxCol = 0;
    let formulaCount = 0;
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const r = rowNumber - 1;
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const c = colNumber - 1;
        const desc = cellFromExcelJs(cell);
        if (desc.f === undefined && desc.v === undefined) return;
        if (desc.f !== undefined) formulaCount++;
        (cellData[r] ??= {})[c] = desc;
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      });
    });
    sheets.push({
      name: ws.name,
      cellData,
      rowCount: maxRow + 20,
      columnCount: maxCol + 5,
      formulaCount,
    });
  });
  return sheets;
}

/** Most-referenced literal cell on the sheet with the most formulas — a
 * realistic "assumption" input whose edit fans out through the model. */
function pickInputCell(sheet: LoadedSheet): { r: number; c: number } | null {
  const refCounts = new Map<string, number>();
  for (const rk of Object.keys(sheet.cellData)) {
    for (const ck of Object.keys(sheet.cellData[Number(rk)])) {
      const f = sheet.cellData[Number(rk)][Number(ck)].f;
      if (typeof f !== "string") continue;
      for (const m of f.matchAll(/(?<![A-Za-z0-9_.!$])\$?([A-Z]{1,3})\$?(\d{1,7})(?![A-Za-z0-9_(])/g)) {
        refCounts.set(`${m[1]}${m[2]}`, (refCounts.get(`${m[1]}${m[2]}`) ?? 0) + 1);
      }
    }
  }
  const colToIndex = (letters: string) => {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  };
  let best: { r: number; c: number } | null = null;
  let bestCount = 0;
  for (const [addr, count] of refCounts) {
    if (count <= bestCount) continue;
    const m = /^([A-Z]+)(\d+)$/.exec(addr)!;
    const r = Number(m[2]) - 1;
    const c = colToIndex(m[1]);
    const cell = sheet.cellData[r]?.[c];
    if (!cell || cell.f !== undefined || typeof cell.v !== "number") continue;
    best = { r, c };
    bestCount = count;
  }
  return best;
}

function pickFormulaCell(sheet: LoadedSheet): { r: number; c: number; f: string } | null {
  for (const rk of Object.keys(sheet.cellData)) {
    for (const ck of Object.keys(sheet.cellData[Number(rk)])) {
      const cell = sheet.cellData[Number(rk)][Number(ck)];
      if (typeof cell.f === "string" && typeof cell.v === "number") {
        return { r: Number(rk), c: Number(ck), f: cell.f };
      }
    }
  }
  return null;
}

// --- optional CPU profiling of one edit (CALC_BENCH_PROF=1) ---------------
async function startCpuProfile(): Promise<any> {
  const inspector = await import("node:inspector");
  const session = new inspector.Session();
  session.connect();
  const post = (method: string) =>
    new Promise<any>((resolve, reject) =>
      session.post(method, (err, res) => (err ? reject(err) : resolve(res))),
    );
  await post("Profiler.enable");
  await post("Profiler.start");
  return { session, post };
}

async function stopCpuProfile(prof: any): Promise<void> {
  const { profile } = await prof.post("Profiler.stop");
  fs.writeFileSync("/tmp/calc-edit.cpuprofile", JSON.stringify(profile));
  // Rough self-time attribution: sample counts per function.
  const total = profile.nodes.reduce((n: number, node: any) => n + (node.hitCount ?? 0), 0);
  const byFn = new Map<string, number>();
  for (const node of profile.nodes) {
    if (!node.hitCount) continue;
    const cf = node.callFrame;
    const file = String(cf.url).split("/").slice(-2).join("/");
    const key = `${cf.functionName || "(anonymous)"} @ ${file}`;
    byFn.set(key, (byFn.get(key) ?? 0) + node.hitCount);
  }
  const top = [...byFn.entries()].sort((a, b) => b[1] - a[1]).slice(0, 20);
  log(`--- cpu profile of edit #2 (${total} samples) → /tmp/calc-edit.cpuprofile ---`);
  for (const [key, hits] of top) {
    log(`  ${((hits / total) * 100).toFixed(1).padStart(5)}%  ${key}`);
  }
  prof.session.disconnect();
}

export async function runBench(filePath: string, pro: boolean): Promise<void> {
  const label = pro ? "PRO" : "OSS";
  log(`\n=== calc bench [${label}] ${filePath} ===`);

  const t0 = performance.now();
  const sheets = await loadXlsx(filePath);
  const totalFormulas = sheets.reduce((n, s) => n + s.formulaCount, 0);
  log(
    `parsed in ${(performance.now() - t0).toFixed(0)}ms — ${sheets.length} sheets, ${totalFormulas} formulas`,
  );

  const { Univer, LocaleType } = await import("@univerjs/core");
  const { FUniver } = await import("@univerjs/core/facade");
  const { UniverSheetsPlugin } = await import("@univerjs/sheets");
  const { UniverSheetsFormulaPlugin } = await import("@univerjs/sheets-formula");
  await import("@univerjs/sheets/facade");
  await import("@univerjs/engine-formula/facade");
  await import("@univerjs/sheets-formula/facade");

  const univer = new Univer({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: {} as any },
  });
  univer.registerPlugin(UniverSheetsPlugin);
  if (pro) {
    const { UniverProFormulaEnginePlugin } = await import("@univerjs-pro/engine-formula");
    univer.registerPlugin(UniverProFormulaEnginePlugin as any);
  } else {
    const { UniverFormulaEnginePlugin } = await import("@univerjs/engine-formula");
    univer.registerPlugin(UniverFormulaEnginePlugin);
  }
  univer.registerPlugin(UniverSheetsFormulaPlugin);

  const univerAPI = FUniver.newAPI(univer);
  const formula: any = univerAPI.getFormula();

  // One line per calculation cycle, mirroring the app's [calc-perf] log.
  let calcStart = 0;
  let lastCalcMs = -1;
  let cycles = 0;
  formula.calculationStart(() => {
    calcStart = performance.now();
  });
  formula.calculationEnd(() => {
    lastCalcMs = performance.now() - calcStart;
    cycles++;
  });

  const workbookData: any = {
    id: "bench",
    name: "bench",
    sheetOrder: sheets.map((_, i) => `s${i}`),
    sheets: Object.fromEntries(
      sheets.map((s, i) => [
        `s${i}`,
        {
          id: `s${i}`,
          name: s.name,
          rowCount: Math.max(100, s.rowCount),
          columnCount: Math.max(20, s.columnCount),
          cellData: s.cellData,
        },
      ]),
    ),
    styles: {},
  };

  const waitForCalc = async (timeoutMs = 300_000): Promise<number> => {
    const before = cycles;
    const deadline = performance.now() + timeoutMs;
    // Give the engine a beat to flip to "computing", then wait it out.
    await new Promise((r) => setTimeout(r, 150));
    while (cycles === before && performance.now() < deadline) {
      const done = await formula.whenComputingCompleteAsync?.(5_000);
      if (done && cycles > before) break;
      if (cycles > before) break;
      await new Promise((r) => setTimeout(r, 100));
    }
    return cycles > before ? lastCalcMs : -1;
  };

  // --- 1. initial calculation --------------------------------------------
  const tCreate = performance.now();
  univerAPI.createWorkbook(workbookData);
  const initialMs = await waitForCalc();
  log(
    `initial calc: ${initialMs < 0 ? "did not run (values cached?)" : `${initialMs.toFixed(0)}ms`} ` +
      `(createWorkbook→idle ${(performance.now() - tCreate).toFixed(0)}ms)`,
  );

  // --- 2. edit one literal input cell -------------------------------------
  const target = sheets.reduce((a, b) => (b.formulaCount > a.formulaCount ? b : a));
  const wb = univerAPI.getActiveWorkbook()!;
  const sheet = wb.getSheetByName(target.name)!;

  const input = pickInputCell(target);
  if (input) {
    const orig = target.cellData[input.r][input.c].v as number;
    for (let i = 1; i <= 3; i++) {
      const profiling = i === 2 && process.env.CALC_BENCH_PROF === "1";
      const prof = profiling ? await startCpuProfile() : null;
      const tEdit = performance.now();
      sheet.getRange(input.r, input.c).setValue(orig * (1 + i * 0.01));
      const calcMs = await waitForCalc(120_000);
      const total = performance.now() - tEdit;
      if (prof) await stopCpuProfile(prof);
      log(
        `edit value ${target.name}!r${input.r}c${input.c} #${i}: ` +
          `${calcMs < 0 ? "no recalc observed" : `calc ${calcMs.toFixed(0)}ms`}, edit→idle ${total.toFixed(0)}ms`,
      );
    }
    sheet.getRange(input.r, input.c).setValue(orig);
    await waitForCalc(120_000);
  } else {
    log("no literal input cell found to edit");
  }

  // --- 3. rewrite one formula cell ----------------------------------------
  const fc = pickFormulaCell(target);
  if (fc) {
    const tEdit = performance.now();
    sheet.getRange(fc.r, fc.c).setValue({ f: fc.f } as any);
    const calcMs = await waitForCalc(120_000);
    const total = performance.now() - tEdit;
    log(
      `edit formula ${target.name}!r${fc.r}c${fc.c}: ` +
        `${calcMs < 0 ? "no recalc observed" : `calc ${calcMs.toFixed(0)}ms`}, edit→idle ${total.toFixed(0)}ms`,
    );
  }

  // --- 4. fixed-overhead probe: formula with zero refs in an empty cell ----
  {
    const tEdit = performance.now();
    sheet.getRange(target.rowCount - 2, 1).setValue({ f: "=1+1" } as any);
    const calcMs = await waitForCalc(120_000);
    const total = performance.now() - tEdit;
    log(
      `edit isolated =1+1: ${calcMs < 0 ? "no recalc observed" : `calc ${calcMs.toFixed(0)}ms`}, edit→idle ${total.toFixed(0)}ms`,
    );
  }

  log(`total calc cycles: ${cycles}`);
  univer.dispose();
}
