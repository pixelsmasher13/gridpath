/// <reference types="vite/client" />

// @univerjs/preset-sheets-drawing ships its types only under the `exports`
// map, which tsc's classic "Node" moduleResolution can't see. Minimal ambient
// surface for what we use (vite resolves the real modules at build time).
declare module "@univerjs/preset-sheets-drawing" {
  export function UniverSheetsDrawingPreset(config?: Record<string, unknown>): unknown;
}
declare module "@univerjs/preset-sheets-drawing/locales/en-US" {
  const locale: Record<string, unknown>;
  export default locale;
}
