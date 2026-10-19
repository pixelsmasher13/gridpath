// Builds the gridpath review page (packages/gridpath/review-ui) into
// packages/gridpath/ui, which the review server serves. Lives in the root
// project so it resolves Univer from the app's node_modules.
import { defineConfig, type Plugin } from "vite";
import path from "node:path";

// @univerjs/engine-render dynamically imports ~80 hyphenation dictionaries
// (document typography). A read-only sheet review never hyphenates; without
// this they become 80 chunks and 3 MB of tarball.
function stubHyphenationDictionaries(): Plugin {
  const dir = path.sep + path.join("@univerjs", "engine-render", "lib", "es") + path.sep;
  return {
    name: "gridpath-stub-hyphenation",
    enforce: "pre",
    load(id) {
      if (id.includes(dir) && !/[\\/](index|facade)\.js$/.test(id)) return "export default {};";
      return null;
    },
  };
}

export default defineConfig({
  root: path.resolve(__dirname, "packages/gridpath/review-ui"),
  base: "./",
  build: {
    outDir: path.resolve(__dirname, "packages/gridpath/ui"),
    emptyOutDir: true,
    sourcemap: false,
  },
  worker: { format: "es", plugins: () => [stubHyphenationDictionaries()] },
  plugins: [stubHyphenationDictionaries()],
});
