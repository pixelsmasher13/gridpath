/**
 * Review-page calc worker: Univer's RPC scaffolding with formula execution
 * DISABLED. Values shown in the review grid are the ones the Node engine
 * computed; the browser must not recompute (and disagree). Lean on purpose —
 * the desktop app's worker also carries the filter preset and the Pro engine.
 */
import { createUniver, LocaleType } from "@univerjs/presets";
import { UniverFormulaEnginePlugin } from "@univerjs/engine-formula";
import { UniverRPCWorkerThreadPlugin } from "@univerjs/rpc";
import { UniverSheetsPlugin } from "@univerjs/sheets";
import { UniverRemoteSheetsFormulaPlugin } from "@univerjs/sheets-formula";

createUniver({
  locale: LocaleType.EN_US,
  locales: {},
  presets: [
    {
      plugins: [
        [UniverSheetsPlugin, { onlyRegisterFormulaRelatedMutations: true }],
        [UniverFormulaEnginePlugin, { notExecuteFormula: true }],
        UniverRPCWorkerThreadPlugin,
        UniverRemoteSheetsFormulaPlugin,
      ],
    },
  ],
});
