// Writes packages/gridpath/THIRD_PARTY_NOTICES.md: the third-party software
// compiled or bundled into the gridpath package, with its license texts.
//
//   node scripts/third-party-notices.mjs
//
// Covers the Rust crates compiled into the engine (engine/, dist/web/*.wasm),
// the npm packages bundled into the web build (dist/web/, built first) and
// into the review page (ui/). React, styled-components and the MCP SDK are
// installed as dependencies and carry their own licenses.
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const OUT = path.join(ROOT, "packages/gridpath/THIRD_PARTY_NOTICES.md");
const LICENSE_FILE = /^(licen[cs]e|copying|unlicense|notice)([-._].*)?$/i;

function licenseTexts(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs
    .readdirSync(dir)
    .filter((f) => LICENSE_FILE.test(f) && fs.statSync(path.join(dir, f)).isFile())
    .sort()
    .map((f) => fs.readFileSync(path.join(dir, f), "utf8").trim());
}

// Rust: normal (non-dev, non-build) dependencies of the engine for wasm32.
function rustCrates() {
  const meta = JSON.parse(
    execFileSync("cargo", ["metadata", "--format-version", "1", "--filter-platform", "wasm32-unknown-unknown"], {
      cwd: path.join(ROOT, "crates/engine-wasm"),
      maxBuffer: 256 << 20,
    }),
  );
  const nodes = new Map(meta.resolve.nodes.map((n) => [n.id, n]));
  const pkgs = new Map(meta.packages.map((p) => [p.id, p]));
  const seen = new Set();
  const stack = [meta.resolve.root];
  while (stack.length) {
    const id = stack.pop();
    if (seen.has(id)) continue;
    seen.add(id);
    for (const dep of nodes.get(id).deps) if (dep.dep_kinds.some((k) => k.kind === null)) stack.push(dep.pkg);
  }
  return [...seen]
    .map((id) => pkgs.get(id))
    // Ours, and proc macros (they run at compile time; nothing of them is in the binary).
    .filter((p) => !p.name.startsWith("gridpath_") && p.name !== "xlsx_patch" && !p.targets.every((t) => t.kind.includes("proc-macro")))
    .map((p) => {
      let dir = path.dirname(p.manifest_path);
      // The vendored IronCalc crates keep their license files at the fork root.
      if (dir.includes(`${path.sep}vendor${path.sep}ironcalc`)) dir = path.join(ROOT, "vendor/ironcalc");
      return { name: p.name, version: p.version, license: p.license ?? "see text", texts: licenseTexts(dir) };
    })
    .sort((a, b) => a.name.localeCompare(b.name));
}

// npm: the packages whose code a bundle's source maps list.
function bundledPackages(dir) {
  const names = new Set();
  const walk = (d) => {
    for (const f of fs.readdirSync(d)) {
      const p = path.join(d, f);
      if (fs.statSync(p).isDirectory()) walk(p);
      else if (p.endsWith(".map"))
        for (const s of JSON.parse(fs.readFileSync(p, "utf8")).sources ?? []) {
          if (!s.includes("node_modules/")) continue;
          const parts = s.split("node_modules/").pop().split("/");
          names.add(parts[0].startsWith("@") ? `${parts[0]}/${parts[1]}` : parts[0]);
        }
    }
  };
  walk(dir);
  // A package's own source map can point into its source repository
  // (lucide-react's "shared/src/…"): only installed packages are listed;
  // that code is part of the package that ships it.
  return [...names].filter((name) => fs.existsSync(path.join(ROOT, "node_modules", name, "package.json"))).sort().map((name) => {
    const dir = path.join(ROOT, "node_modules", name);
    const pj = JSON.parse(fs.readFileSync(path.join(dir, "package.json"), "utf8"));
    const license = typeof pj.license === "string" ? pj.license : "see text";
    const author = typeof pj.author === "string" ? pj.author : pj.author?.name;
    const texts = licenseTexts(dir);
    return { name, version: pj.version, license, texts: texts.length ? texts : standardText(license, author ?? `the ${name} authors`) };
  });
}

// For packages published without a license file: the standard text of the
// license they declare, with their author as the copyright holder.
function standardText(license, holder) {
  if (license === "MIT")
    return [`MIT License\n\nCopyright (c) ${holder}\n\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\nTHE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.`];
  if (license === "ISC")
    return [`ISC License\n\nCopyright (c) ${holder}\n\nPermission to use, copy, modify, and/or distribute this software for any purpose with or without fee is hereby granted, provided that the above copyright notice and this permission notice appear in all copies.\n\nTHE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.`];
  if (license === "Apache-2.0") return [fs.readFileSync(path.join(ROOT, "vendor/ironcalc/LICENSE-Apache-2.0"), "utf8").split("   APPENDIX")[0].trim() + `\n\nCopyright ${holder}`];
  return [];
}

function reviewPagePackages() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "gridpath-notices-"));
  execFileSync("npx", ["vite", "build", "-c", "vite.review.config.ts", "--sourcemap", "--outDir", tmp, "--emptyOutDir"], { cwd: ROOT, stdio: "ignore" });
  try {
    return bundledPackages(tmp);
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

const WEB = path.join(ROOT, "packages/gridpath/dist/web");
if (!fs.existsSync(WEB)) throw new Error("build the package first (npm run build in packages/gridpath)");
const rust = rustCrates();
const web = bundledPackages(WEB);
const npm = reviewPagePackages();
const missing = [...rust, ...web, ...npm].filter((p) => p.texts.length === 0);

// Identical texts are printed once and referenced.
const textIds = new Map();
const textOf = (t) => {
  if (!textIds.has(t)) textIds.set(t, textIds.size + 1);
  return textIds.get(t);
};
const row = (p) => `| ${p.name} | ${p.version} | ${p.license} | ${p.texts.map((t) => `[${textOf(t)}](#license-text-${textOf(t)})`).join(", ") || "—"} |`;

let md = `# Third-party notices

The \`gridpath\` package (Apache-2.0, see LICENSE and NOTICE) includes the
third-party software below. Generated by \`scripts/third-party-notices.mjs\`.

React, React DOM, styled-components, the MCP SDK and zod are not bundled:
they are installed as dependencies and carry their own licenses.

## Compiled into the engine (\`engine/\`, \`dist/web/gridpath_engine_bg.wasm\`)

| Crate | Version | License | Text |
|---|---|---|---|
${rust.map(row).join("\n")}

## Bundled into the web build (\`dist/web/\`)

| Package | Version | License | Text |
|---|---|---|---|
${web.map(row).join("\n") || "| none | | | |"}

## Bundled into the review page (\`ui/\`)

| Package | Version | License | Text |
|---|---|---|---|
${npm.map(row).join("\n")}

## License texts
`;
for (const [text, id] of textIds) md += `\n### License text ${id}\n\n\`\`\`\n${text}\n\`\`\`\n`;
fs.writeFileSync(OUT, md);
console.log(`${rust.length} crates, ${web.length} + ${npm.length} npm packages, ${textIds.size} distinct license texts -> ${path.relative(ROOT, OUT)}`);
if (missing.length) console.log("no license file found for:", missing.map((p) => `${p.name}@${p.version} (${p.license})`).join(", "));
