import { defineConfig } from "tsup";

export default defineConfig({
  entry: { index: "src/index.ts", "bin/gridpath": "src/bin/gridpath.ts" },
  format: ["esm"],
  target: "node20",
  platform: "node",
  dts: { entry: { index: "src/index.ts" } },
  clean: true,
  sourcemap: true,
  // The WASM engine is loaded at runtime from ./engine, never bundled.
  external: [/\.\/engine\//],
});
