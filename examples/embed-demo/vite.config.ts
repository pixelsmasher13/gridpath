import { defineConfig, type Connect, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";
import { agentTurn } from "./server/agent";

/** POST /api/agent on the dev and preview servers: the demo's backend. */
function agentRoute(): Plugin {
  const handle: Connect.NextHandleFunction = async (req, res, next) => {
    if (req.url !== "/api/agent" || req.method !== "POST") return next();
    let raw = "";
    for await (const chunk of req) raw += chunk;
    res.setHeader("content-type", "application/json");
    try {
      res.end(JSON.stringify(await agentTurn(JSON.parse(raw))));
    } catch (e: any) {
      res.statusCode = typeof e?.status === "number" ? e.status : 500;
      res.end(JSON.stringify({ error: e?.message ?? String(e) }));
    }
  };
  return {
    name: "gridpath-agent-route",
    configureServer: (server) => void server.middlewares.use(handle),
    configurePreviewServer: (server) => void server.middlewares.use(handle),
  };
}

export default defineConfig({
  // An already-open recording page can start another worker after a rebuild.
  // Keep its hashed assets available until that page is reloaded.
  build: process.env.VITE_RECORDING === "1" ? { emptyOutDir: false } : undefined,
  // Recording uses the locally rebuilt engine, including the namespace-save
  // regression fix. The ordinary demo retains its packaged dependency.
  resolve: process.env.VITE_RECORDING === "1" ? {
    alias: [{ find: /^gridpath$/, replacement: fileURLToPath(new URL("../../packages/gridpath/dist/web/index.js", import.meta.url)) }],
  } : undefined,
  plugins: [react(), agentRoute()],
  // The package loads its worker and .wasm with `new URL(..., import.meta.url)`;
  // dev-server pre-bundling would move the script away from those files.
  optimizeDeps: { exclude: ["gridpath"] },
});
