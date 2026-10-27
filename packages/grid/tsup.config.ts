import { copyFileSync } from "node:fs";

// The browser half of the `gridpath` package: built into its dist/web.
const OUT = "../gridpath/dist/web";
import { defineConfig } from "tsup";

export default defineConfig({
  entry: { index: "src/index.tsx", agent: "src/agent.ts", "engine.worker": "src/engine.worker.ts" },
  format: ["esm"],
  platform: "browser",
  target: "es2020",
  // The page-thread engine (`worker: false`) is a separate chunk, loaded
  // only when asked for.
  splitting: true,
  dts: { entry: { index: "src/index.tsx", agent: "src/agent.ts" } },
  outDir: OUT,
  clean: true,
  sourcemap: true,
  external: ["react", "react-dom", "react/jsx-runtime", "styled-components"],
  noExternal: [/^echarts(?:\/|$)/, /^zrender(?:\/|$)/],
  esbuildOptions(o) {
    o.jsx = "automatic";
  },
  // The wasm-bindgen glue loads the engine from
  // `new URL("gridpath_engine_bg.wasm", import.meta.url)`: it must sit next
  // to the chunk that holds the glue, in dist/web.
  onSuccess: async () => {
    copyFileSync("engine/gridpath_engine_bg.wasm", `${OUT}/gridpath_engine_bg.wasm`);
  },
});
