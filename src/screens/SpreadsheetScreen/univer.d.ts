// Univer ships ESM types via `exports`, which the project's `moduleResolution: "Node"`
// can't resolve. We treat them as `any` at the type-check boundary; runtime is fine.
declare module "@univerjs/presets";
declare module "@univerjs/presets/*";
declare module "@univerjs/preset-sheets-core";
declare module "@univerjs/preset-sheets-core/*";
declare module "@univerjs/preset-sheets-find-replace";
declare module "@univerjs/preset-sheets-find-replace/*";
declare module "@univerjs/preset-sheets-filter";
declare module "@univerjs/preset-sheets-filter/*";
declare module "@univerjs/preset-sheets-conditional-formatting";
declare module "@univerjs/preset-sheets-conditional-formatting/*";
declare module "@univerjs/preset-sheets-data-validation";
declare module "@univerjs/preset-sheets-data-validation/*";
declare module "@univerjs/preset-sheets-sort";
declare module "@univerjs/preset-sheets-sort/*";
declare module "@univerjs/preset-sheets-note";
declare module "@univerjs/preset-sheets-note/*";
