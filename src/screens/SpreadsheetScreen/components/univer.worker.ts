/**
 * Univer calculation worker. The main thread passes this to
 * UniverSheetsCorePreset as `workerURL`, which flips the main-thread formula
 * engine to `notExecuteFormula` and proxies all recalculation here over RPC —
 * bulk agent edits no longer freeze painting while dependency chains recalc.
 *
 * UniverSheetsCoreWorkerPreset is inlined below (it is a 4-plugin list, see
 * @univerjs/preset-sheets-core/lib/es/worker.js) because the preset only
 * forwards `function` to the engine plugin and hides `intervalCount` — the
 * number of formulas calculated between checks for a stop message from the
 * main thread. The default of 500 was tuned for main-thread execution; in a
 * dedicated worker those pauses buy nothing, so we raise it and cut the
 * per-chunk overhead on large dependency chains.
 *
 * Pro engine eval: start the app with VITE_UNIVER_PRO_FORMULA=1 to swap in
 * @univerjs-pro/engine-formula (license-free evaluation mode — watermark +
 * import-size limits, fine for benchmarking). Compare the [calc-perf]
 * console lines against the OSS engine on the same workbook before any
 * licensing decision — Univer's own numbers range from 1.1x to 5.2x
 * depending on the file.
 */
import { createUniver, LocaleType, merge } from "@univerjs/presets";
import { UniverFormulaEnginePlugin } from "@univerjs/engine-formula";
import { UniverRPCWorkerThreadPlugin } from "@univerjs/rpc";
import { UniverSheetsPlugin } from "@univerjs/sheets";
import { UniverRemoteSheetsFormulaPlugin } from "@univerjs/sheets-formula";
import { UniverSheetsFilterWorkerPreset } from "@univerjs/preset-sheets-filter/worker";
import coreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import filterEnUS from "@univerjs/preset-sheets-filter/locales/en-US";

async function init() {
  // IronCalc engine mode: UniverGrid names the worker "univer-calc-noformula"
  // and the engine plugin is registered with notExecuteFormula — the RPC
  // scaffolding stays intact but no formula is ever computed here (IronCalc
  // computes in the Tauri process and writes values into the grid).
  const noFormula = typeof self !== "undefined" && (self as unknown as { name?: string }).name === "univer-calc-noformula";
  const usePro = !noFormula && import.meta.env.VITE_UNIVER_PRO_FORMULA === "1";
  // Dynamic import so the pro package stays out of the bundle when the
  // flag is off.
  const enginePlugin: [any, object] | any = usePro
    ? (await import("@univerjs-pro/engine-formula")).UniverProFormulaEnginePlugin
    : [UniverFormulaEnginePlugin, { intervalCount: 5000, notExecuteFormula: noFormula }];
  if (usePro) console.log("[univer.worker] PRO formula engine active (evaluation mode)");
  if (noFormula) console.log("[univer.worker] formula execution disabled — ironcalc engine mode");

  createUniver({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: merge({}, coreEnUS, filterEnUS) },
    presets: [
      {
        plugins: [
          [UniverSheetsPlugin, { onlyRegisterFormulaRelatedMutations: true }],
          enginePlugin,
          UniverRPCWorkerThreadPlugin,
          UniverRemoteSheetsFormulaPlugin,
        ],
      },
      UniverSheetsFilterWorkerPreset(),
    ],
  });
}

void init();
