# GridPath — marketing site

Static landing page for the GridPath desktop app. Plain HTML + CSS, no build step,
zero-config Vercel deployment.

## Files

- `index.html` — hero, features, comparison table, how-it-works, download, footer
- `style.css` — light theme, mirrors the desktop app's `#3363AD` blue accent
- `favicon.svg` — inline SVG logo (grid mark)
- `vercel.json` — security headers + caching rules
- `package.json` — convenience scripts (no build deps)

## Local preview

```sh
npm run dev
# or:
python3 -m http.server 5173
```

Then open `http://localhost:5173`.

## Deploy to Vercel

### One-time setup

```sh
npm install -g vercel    # if you don't have the CLI
cd marketing
vercel login
```

### Deploy

```sh
# preview deployment (gets a unique URL)
vercel

# production deployment (your custom domain or *.vercel.app)
vercel --prod
```

First `vercel` run will ask:

- **Set up and deploy?** Yes
- **Scope?** pick your team / account
- **Link to existing project?** No (or yes if linking)
- **Project name?** GridPath-marketing
- **Directory?** `./` — accept the default
- **Override settings?** No — `vercel.json` handles it

### What `vercel.json` configures

- `buildCommand: null` — no compile step
- `outputDirectory: "."` — serve from this folder
- `cleanUrls: true` — `/about` instead of `/about.html`
- Security headers: `X-Content-Type-Options`, `X-Frame-Options`,
  `Referrer-Policy`, `Permissions-Policy`
- Aggressive caching for static assets (1y, immutable); `index.html`
  revalidated on every request so copy edits ship instantly

### Alt deploys (also work)

- **Cloudflare Pages** — set build command empty, output dir `.`
- **Netlify** — drag-and-drop the folder
- **GitHub Pages** — push to `gh-pages` branch, enable in repo settings

## Editing

- Hero copy & CTAs: top of `index.html`
- Comparison rows: search for `<table class="compare-table">`
- Footer columns: search for `<footer class="footer">`
- Colors / type: CSS variables at the top of `style.css`
