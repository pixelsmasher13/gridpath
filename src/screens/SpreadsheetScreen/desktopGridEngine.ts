/**
 * The viewport grid's engine on the desktop: the per-tab Tauri engine
 * bridge in ironcalcEngine.ts, bound to one tab. The grid itself no longer
 * imports the bridge, so the same component runs in the browser package
 * over the WASM engine (packages/grid).
 */
import type { GridEngine } from "../../../packages/grid/src/engine";
import {
  ironcalcEngineCalcHealth,
  ironcalcEngineDefineName,
  ironcalcEngineFetchLayout,
  ironcalcEngineFetchWindow,
  ironcalcEngineFind,
  ironcalcEngineOpenLayout,
  ironcalcEngineSetStyle,
  ironcalcEngineSettled,
  ironcalcEngineSheetNames,
  ironcalcEngineSheetOpDirect,
  ironcalcEngineStartDeferredSettle,
  ironcalcEngineStructuralDirect,
  ironcalcEngineWriteCells,
  registerIroncalcApplier,
  unregisterIroncalcApplier,
} from "./ironcalcEngine";

const engines = new Map<string, GridEngine>();

/** One stable engine object per tab (the grid's handle closes over it). */
export function desktopGridEngine(tabId: string): GridEngine {
  let e = engines.get(tabId);
  if (e) return e;
  e = {
    // The Tauri side reads the file from `path`; the bytes stay in the webview.
    openLayout: (path) => ironcalcEngineOpenLayout(tabId, path),
    startDeferredSettle: () => ironcalcEngineStartDeferredSettle(tabId),
    fetchLayout: () => ironcalcEngineFetchLayout(tabId),
    fetchWindow: (sheet, r0, r1, c0, c1) => ironcalcEngineFetchWindow(tabId, sheet, r0, r1, c0, c1),
    calcHealth: () => ironcalcEngineCalcHealth(tabId),
    setStyle: (sheetName, r0, c0, r1, c1, patch) => ironcalcEngineSetStyle(tabId, sheetName, r0, c0, r1, c1, patch),
    defineName: (name, formula) => ironcalcEngineDefineName(tabId, name, formula),
    find: (query, sheet, limit) => ironcalcEngineFind(tabId, query, sheet, limit),
    sheetNames: () => ironcalcEngineSheetNames(tabId),
    writeCells: (sheetName, label, cells) => ironcalcEngineWriteCells(tabId, sheetName, label, cells),
    sheetOp: (op, name, newName) => ironcalcEngineSheetOpDirect(tabId, op, name, newName),
    structural: (op, sheetName, index1, count) => ironcalcEngineStructuralDirect(tabId, op, sheetName, index1, count),
    settled: (timeoutMs) => ironcalcEngineSettled(tabId, timeoutMs),
    registerApplier: (fn) => registerIroncalcApplier(tabId, fn),
    unregisterApplier: () => unregisterIroncalcApplier(tabId),
  };
  engines.set(tabId, e);
  return e;
}

/** Drop a closed tab's engine object. */
export function forgetDesktopGridEngine(tabId: string): void {
  engines.delete(tabId);
}
