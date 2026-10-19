/**
 * Builds the /docs section of gridpath.dev.
 *
 *   npm run docs            (from marketing/)
 *
 * Inputs
 *   docs-src/<slug>.html    hand-written page bodies with two leading
 *                           comment lines: <!-- title: … --> <!-- description: … -->
 *   toolSchemas()           the MCP server's real tool list, imported from the
 *                           package so the reference page cannot drift from the code
 * Output
 *   docs/<slug>.html        complete static pages (header, sidebar, footer)
 *   docs/tools.html         the generated reference
 *
 * The site itself has no build step: the output is committed and Vercel serves it.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { toolSchemas, type ToolSchema } from "../../packages/gridpath/src/tools/handlers";

const here = path.dirname(fileURLToPath(import.meta.url));
const SRC = path.join(here, "..", "docs-src");
const OUT = path.join(here, "..", "docs");
const PKG = JSON.parse(fs.readFileSync(path.join(here, "../../packages/gridpath/package.json"), "utf8"));
const VERSION: string = PKG.version;

/* ------------------------------------------------------------------ */
/* Sidebar                                                             */
/* ------------------------------------------------------------------ */
const NAV: Array<{ group: string; pages: Array<{ slug: string; label: string }> }> = [
  {
    group: "Getting started",
    pages: [
      { slug: "index", label: "Overview" },
      { slug: "claude-code", label: "Claude Code" },
      { slug: "claude-desktop", label: "Claude Desktop" },
      { slug: "cursor", label: "Cursor" },
      { slug: "codex", label: "Codex" },
    ],
  },
  {
    group: "Guides",
    pages: [
      { slug: "review", label: "Review before save" },
      { slug: "reading-a-workbook", label: "Reading a workbook" },
      { slug: "writing-cells", label: "Writing cells and formulas" },
      { slug: "run-script", label: "run_script" },
      { slug: "formatting-and-layout", label: "Formatting and layout" },
      { slug: "structure", label: "Rows, columns and sheets" },
    ],
  },
  {
    group: "Reference",
    pages: [
      { slug: "tools", label: "All tools" },
      { slug: "options", label: "CLI options and limits" },
    ],
  },
];

const urlOf = (slug: string) => (slug === "index" ? "/docs" : `/docs/${slug}`);

/* ------------------------------------------------------------------ */
/* HTML helpers                                                        */
/* ------------------------------------------------------------------ */
const esc = (s: string) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Escape, then turn `code` spans into <code> and blank lines into paragraphs. */
function prose(s: string): string {
  const paras = s.trim().split(/\n\s*\n/);
  return paras
    .map((p) => {
      const lines = p.split("\n").map((l) => esc(l).replace(/`([^`]+)`/g, "<code>$1</code>"));
      return `<p>${lines.join("<br>")}</p>`;
    })
    .join("\n");
}

function sidebar(active: string): string {
  return `<nav class="docs-nav" aria-label="Docs">
${NAV.map(
  (g) => `  <p class="docs-nav-group">${esc(g.group)}</p>
  <ul>
${g.pages.map((p) => `    <li><a href="${urlOf(p.slug)}"${p.slug === active ? ' aria-current="page"' : ""}>${esc(p.label)}</a></li>`).join("\n")}
  </ul>`,
).join("\n")}
  <p class="docs-nav-version">gridpath ${esc(VERSION)}</p>
</nav>`;
}

function page(slug: string, title: string, description: string, body: string): string {
  const url = `https://gridpath.dev${urlOf(slug)}`;
  const fullTitle = slug === "index" ? `${title} — GridPath Docs` : `${title} — GridPath Docs`;
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${esc(fullTitle)}</title>
  <meta name="description" content="${esc(description)}" />
  <link rel="canonical" href="${url}" />
  <meta property="og:type" content="article" />
  <meta property="og:site_name" content="GridPath" />
  <meta property="og:title" content="${esc(fullTitle)}" />
  <meta property="og:description" content="${esc(description)}" />
  <meta property="og:url" content="${url}" />
  <meta property="og:image" content="https://gridpath.dev/og-image.png" />
  <meta name="twitter:card" content="summary_large_image" />
  <meta name="twitter:title" content="${esc(fullTitle)}" />
  <meta name="twitter:description" content="${esc(description)}" />
  <meta name="twitter:image" content="https://gridpath.dev/og-image.png" />
  <link rel="icon" type="image/svg+xml" href="/favicon.svg" />
  <link rel="icon" type="image/png" sizes="32x32" href="/favicon-32.png" />
  <link rel="icon" type="image/png" sizes="64x64" href="/favicon-64.png" />
  <link rel="apple-touch-icon" sizes="128x128" href="/apple-touch-icon.png" />
  <link rel="stylesheet" href="/style.css?v=r5" />
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@type": "TechArticle",
    "headline": ${JSON.stringify(title)},
    "description": ${JSON.stringify(description)},
    "url": ${JSON.stringify(url)},
    "author": { "@type": "Organization", "name": "GridPath" },
    "publisher": { "@type": "Organization", "name": "GridPath", "url": "https://gridpath.dev/" },
    "isPartOf": { "@type": "WebSite", "name": "GridPath", "url": "https://gridpath.dev/" }
  }
  </script>
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@type": "BreadcrumbList",
    "itemListElement": [
      { "@type": "ListItem", "position": 1, "name": "Docs", "item": "https://gridpath.dev/docs" }${
        slug === "index" ? "" : `,
      { "@type": "ListItem", "position": 2, "name": ${JSON.stringify(title)}, "item": ${JSON.stringify(url)} }`
      }
    ]
  }
  </script>
</head>
<body>

  <header class="nav">
    <div class="container nav-inner">
      <a class="brand" href="/">
        <svg width="22" height="22" viewBox="0 0 22 22" fill="none" aria-hidden="true">
          <rect x="1" y="1" width="20" height="20" rx="4" stroke="currentColor" stroke-width="1.6" />
          <path d="M5 8.5 H17 M5 13.5 H17 M8.5 5 V17 M13.5 5 V17" stroke="currentColor" stroke-width="1.4" />
        </svg>
        GridPath
      </a>
      <nav class="nav-links">
        <a href="/#capabilities">Capabilities</a>
        <a href="/embed">Embed</a>
        <a href="/mcp">MCP</a>
        <a href="/docs">Docs</a>
        <a href="/fidelity">Fidelity</a>
        <a href="/#faq">FAQ</a>
        <a href="/blog">Blog</a>
        <a href="https://github.com/pixelsmasher13/gridpath" rel="noopener" target="_blank">GitHub</a>
        <a href="/#download" class="nav-cta">Download</a>
      </nav>
    </div>
  </header>

  <main class="docs">
    <div class="container docs-layout">
      <aside class="docs-side">
${sidebar(slug)}
      </aside>
      <article class="docs-body legal post">
        <p class="docs-crumb"><a href="/docs">Docs</a>${slug === "index" ? "" : ` / ${esc(title)}`}</p>
${body.trim()}
        <p class="docs-foot">Found a mistake? <a href="https://github.com/pixelsmasher13/gridpath/issues" rel="noopener" target="_blank">Open an issue</a> or email <a href="mailto:support@gridpath.dev">support@gridpath.dev</a>.</p>
      </article>
    </div>
  </main>

  <footer class="footer">
    <div class="container footer-grid">
      <div class="footer-brand">
        <a class="brand" href="/">
          <svg width="20" height="20" viewBox="0 0 22 22" fill="none" aria-hidden="true">
            <rect x="1" y="1" width="20" height="20" rx="4" stroke="currentColor" stroke-width="1.6" />
            <path d="M5 8.5 H17 M5 13.5 H17 M8.5 5 V17 M13.5 5 V17" stroke="currentColor" stroke-width="1.4" />
          </svg>
          GridPath
        </a>
        <p class="footer-tag">Best way to build spreadsheets with AI.</p>
      </div>
      <div class="footer-col">
        <h4>Product</h4>
        <a href="/#capabilities">Capabilities</a>
        <a href="/embed">Embeddable spreadsheet</a>
        <a href="/mcp">MCP server</a>
        <a href="/fidelity">Fidelity</a>
        <a href="/#providers">Models</a>
        <a href="/#download">Download</a>
      </div>
      <div class="footer-col">
        <h4>Resources</h4>
        <a href="/docs">Docs</a>
        <a href="/#faq">FAQ</a>
        <a href="/blog">Blog</a>
        <a href="https://github.com/pixelsmasher13/gridpath" rel="noopener" target="_blank">GitHub</a>
      </div>
      <div class="footer-col">
        <h4>Company</h4>
        <a href="/contact">Contact</a>
        <a href="/privacy">Privacy</a>
        <a href="/terms">Terms</a>
      </div>
    </div>
    <div class="container footer-bottom">
      <span>© 2026 GridPath</span>
      <span>Made for people who actually live in spreadsheets.</span>
    </div>
  </footer>

</body>
</html>
`;
}

/* ------------------------------------------------------------------ */
/* Tool reference                                                      */
/* ------------------------------------------------------------------ */
const GROUPS: Array<{ title: string; blurb: string; tools: string[] }> = [
  { title: "Reading", blurb: "Read-only. Hosts can run these without an approval prompt.", tools: ["describe_workbook", "find_rows", "query_rows", "read_range"] },
  { title: "Writing", blurb: "Each call lands as one pending batch and returns a readback of what was written.", tools: ["set_cell", "set_range", "copy_range", "clear_range", "run_script"] },
  { title: "Formatting and layout", blurb: "", tools: ["set_format", "set_column_width", "set_row_height", "merge_cells", "unmerge_cells", "freeze_panes", "unfreeze_panes", "hide_rows", "show_rows", "hide_columns", "show_columns", "set_note", "delete_note", "define_name"] },
  { title: "Rows, columns and sheets", blurb: "Structural edits. Formula references shift the way they do in Excel.", tools: ["insert_rows", "delete_rows", "insert_columns", "delete_columns", "create_sheet", "rename_sheet", "delete_sheet"] },
  { title: "Batches and saving", blurb: "Nothing touches the file until save_workbook.", tools: ["list_batches", "reject_batch", "save_workbook"] },
];
const READ_ONLY = new Set(["describe_workbook", "find_rows", "query_rows", "read_range", "list_batches"]);

/** First sentence of a tool description, as an HTML snippet: trailing "— e.g." / "(e.g." fragments dropped, `code` rendered. */
const firstSentence = (s: string) => {
  const flat = s.replace(/\s+/g, " ");
  const m = flat.match(/^(.+?[.!?])(\s|$)/);
  let one = (m ? m[1] : flat).replace(/\*\*/g, "");
  one = one.replace(/\s*[—–-]\s*e\.g\.$/i, ".").replace(/\s*\(e\.g\.$/i, ".").replace(/\s*[—–:,]\s*$/, ".");
  return esc(one).replace(/`([^`]+)`/g, "<code>$1</code>");
};

const typeOf = (p: any): string => {
  if (!p) return "";
  const t = Array.isArray(p.type) ? p.type.join(" | ") : p.type ?? (p.enum ? "string" : "any");
  if (t === "array" && p.items) {
    const it = p.items.type ?? (p.items.properties ? "object" : "any");
    return `array of ${Array.isArray(it) ? it.join(" | ") : it}`;
  }
  if (p.enum) return `${t}: ${p.enum.map((e: string) => `"${e}"`).join(" | ")}`;
  return t === undefined ? "any" : t;
};

function paramRows(props: Record<string, any>, required: Set<string>, prefix = ""): string {
  const out: string[] = [];
  const names = Object.keys(props).filter((k) => k !== "path");
  // required first, then alphabetical
  names.sort((a, b) => Number(required.has(b)) - Number(required.has(a)) || a.localeCompare(b));
  for (const name of names) {
    const p = props[name] ?? {};
    const req = required.has(name);
    out.push(
      `<tr><td><code>${esc(prefix + name)}</code>${req ? ' <span class="docs-req">required</span>' : ""}</td><td>${esc(typeOf(p))}</td><td>${p.description ? prose(String(p.description)).replace(/^<p>|<\/p>$/g, "").replace(/<\/p>\n<p>/g, "<br>") : ""}</td></tr>`,
    );
    const nested = p.properties ?? p.items?.properties;
    if (nested) {
      const nreq = new Set<string>((p.required ?? p.items?.required ?? []) as string[]);
      out.push(paramRows(nested, nreq, `${prefix}${name}${p.items ? "[]" : ""}.`));
    }
  }
  return out.join("\n");
}

function toolSection(t: ToolSchema): string {
  const props = t.input_schema?.properties ?? {};
  const required = new Set<string>((t.input_schema?.required ?? []) as string[]);
  const badge = READ_ONLY.has(t.name)
    ? '<span class="docs-badge docs-badge-read">read-only</span>'
    : t.name === "save_workbook"
      ? '<span class="docs-badge docs-badge-save">writes the file</span>'
      : '<span class="docs-badge docs-badge-write">pending batch</span>';
  const rows = paramRows(props, required);
  return `<section class="docs-tool" id="${esc(t.name)}">
  <h3><a href="#${esc(t.name)}"><code>${esc(t.name)}</code></a> ${badge}</h3>
  <div class="docs-tool-desc">
${prose(t.description)}
  </div>
${
  rows
    ? `  <div class="table-wrap"><table>
    <thead><tr><th>Parameter</th><th>Type</th><th>Description</th></tr></thead>
    <tbody>
${rows}
    </tbody>
  </table></div>`
    : `  <p class="docs-muted">No parameters beyond <code>path</code>.</p>`
}
</section>`;
}

function toolsPage(tools: ToolSchema[]): { title: string; description: string; body: string } {
  const byName = new Map(tools.map((t) => [t.name, t]));
  const placed = new Set<string>();
  const groups = GROUPS.map((g) => ({ ...g, list: g.tools.filter((n) => byName.has(n)).map((n) => (placed.add(n), byName.get(n)!)) }));
  const rest = tools.filter((t) => !placed.has(t.name));
  if (rest.length) groups.push({ title: "Other", blurb: "", tools: rest.map((t) => t.name), list: rest });

  const index = groups
    .map(
      (g) => `<h2 id="${esc(g.title.toLowerCase().replace(/[^a-z]+/g, "-"))}">${esc(g.title)}</h2>
${g.blurb ? `<p>${esc(g.blurb)}</p>` : ""}
<div class="table-wrap"><table class="docs-index">
<tbody>
${g.list.map((t) => `<tr><td><a href="#${esc(t.name)}"><code>${esc(t.name)}</code></a></td><td>${firstSentence(t.description)}</td></tr>`).join("\n")}
</tbody>
</table></div>`,
    )
    .join("\n");

  const sections = groups.map((g) => g.list.map(toolSection).join("\n")).join("\n");

  const body = `<h1 class="legal-title">All tools.</h1>
<p class="legal-lede">Every tool the GridPath MCP server exposes, generated from the server's own tool definitions for gridpath ${esc(VERSION)}. The descriptions are the ones the agent sees. ${tools.length} tools.</p>

<p>Every tool takes <code>path</code>: the <code>.xlsx</code> or <code>.xlsm</code> to operate on, absolute or <code>~/…</code>. It must sit inside an allowed folder (see <a href="/docs/options">CLI options</a>). Cells and ranges are A1 notation; sheet names are case-sensitive.</p>
<p>Write tools never touch the file. They accumulate as pending batches, each write returns a readback of what landed plus a <code>review_url</code>, and <code>save_workbook</code> is the only call that writes to disk. See <a href="/docs/review">Review before save</a>.</p>

${index}

<h2 id="reference">Reference</h2>
${sections}`;
  return { title: "All tools", description: `Reference for every tool in the GridPath Excel MCP server (gridpath ${VERSION}): parameters, types and the descriptions the agent sees. Generated from the server's tool definitions.`, body };
}

/* ------------------------------------------------------------------ */
/* Build                                                               */
/* ------------------------------------------------------------------ */
function readFragment(file: string): { title: string; description: string; body: string } {
  const raw = fs.readFileSync(file, "utf8");
  const title = raw.match(/<!--\s*title:\s*(.*?)\s*-->/)?.[1];
  const description = raw.match(/<!--\s*description:\s*(.*?)\s*-->/)?.[1];
  if (!title || !description) throw new Error(`${path.basename(file)}: needs <!-- title: --> and <!-- description: --> comments`);
  const body = raw.replace(/<!--\s*(title|description):.*?-->\s*/g, "");
  return { title, description, body };
}

fs.mkdirSync(OUT, { recursive: true });
const written: string[] = [];

for (const f of fs.readdirSync(SRC).filter((f) => f.endsWith(".html"))) {
  const slug = f.replace(/\.html$/, "");
  const { title, description, body } = readFragment(path.join(SRC, f));
  fs.writeFileSync(path.join(OUT, f), page(slug, title, description, body));
  written.push(slug);
}

const tools = toolSchemas();
const tp = toolsPage(tools);
fs.writeFileSync(path.join(OUT, "tools.html"), page("tools", tp.title, tp.description, tp.body));
written.push("tools");

// Every sidebar entry must exist, and every page must be in the sidebar.
const inNav = new Set(NAV.flatMap((g) => g.pages.map((p) => p.slug)));
for (const s of inNav) if (!written.includes(s)) throw new Error(`sidebar links to /docs/${s} but no page was built`);
for (const s of written) if (!inNav.has(s)) throw new Error(`docs/${s}.html is not in the sidebar`);

console.log(`built ${written.length} docs pages (${tools.length} tools) → marketing/docs/`);
