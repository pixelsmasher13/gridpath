// Univer ships ESM types via `exports`, which the project's `moduleResolution: "Node"`
// can't resolve. We treat them as `any` at the type-check boundary; runtime is fine.
declare module "@univerjs/presets";
declare module "@univerjs/presets/*";
declare module "@univerjs/preset-sheets-core";
declare module "@univerjs/preset-sheets-core/*";
