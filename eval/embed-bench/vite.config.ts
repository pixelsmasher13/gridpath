import { defineConfig } from "vite";

export default defineConfig({
  build: {
    target: "esnext",
    rollupOptions: { input: { gridpath: "gridpath.html", spreadjs: "spreadjs.html", hot: "hot.html" } },
    chunkSizeWarningLimit: 20_000,
  },
  optimizeDeps: { exclude: ["gridpath"] },
});
