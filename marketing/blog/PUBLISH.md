# Blog staging plan (created July 18, 2026)

> **July 20:** `same-model-benchmark.html` is LIVE — listed, indexed, dated
> July 20, main-page card + FAQ link to it. Distribution notes at the bottom
> still apply.

Three posts, released one per week. Staged posts are deployed but **unlisted**
(not in `blog/index.html` or `sitemap.xml`) and carry a `noindex` meta tag so
search engines ignore them until launch.

| Week | Post | File | Status |
|------|------|------|--------|
| 1 (now) | Dead numbers | `dead-numbers.html` | LIVE — listed, indexed, dated July 18 |
| 2 | Spreadsheet disasters | `spreadsheet-disasters.html` | LIVE — listed, indexed, dated Aug 4 (+ new "never make the news" section from r/excel material) |
| 3 | AI in Excel guide | `ai-in-excel.html` | staged (noindex, unlisted) |

> **Aug 4:** `claude-vs-chatgpt-for-excel.html` published same day (not part of
> this staging plan — research notes in `RESEARCH-claude-vs-chatgpt-excel.md`).
> NOTE: deploys now run `vercel --prod` from the REPO ROOT (project root
> directory is set to `marketing`; a root `.vercelignore` limits the upload).

Distribution notes per post are at the bottom. Delete this file when all three
are out.

---

## Week 2 — publish `spreadsheet-disasters.html`

1. **Remove the noindex tag** in `spreadsheet-disasters.html` (the line marked
   `STAGED: remove at publish`).
2. **Update the date** to the actual publish day in three places in that file —
   `article:published_time`, JSON-LD `datePublished` (both `2026-07-17`), and
   the `legal-meta` line (`JULY 17, 2026`).
3. **Add the index card** to `blog/index.html`, above the dead-numbers card
   (update the date to match step 2):

```html
        <a class="blog-card" href="/blog/spreadsheet-disasters">
          <span class="blog-card-date">July 24, 2026</span>
          <h2>The London Whale and other great spreadsheet disasters</h2>
          <p>
            A $6.2 billion trading loss, a global austerity policy, a
            pandemic response, and the human genome — four disasters that
            ran through a spreadsheet, none caused by people who were bad at
            spreadsheets, and the one failure mode they all share.
          </p>
          <span class="blog-read">Read the post →</span>
        </a>
```

4. **Add the sitemap entry** to `sitemap.xml`, next to the other blog posts:

```xml
  <url>
    <loc>https://gridpath.dev/blog/spreadsheet-disasters</loc>
    <changefreq>monthly</changefreq>
    <priority>0.6</priority>
  </url>
```

5. Deploy (`vercel --prod` from `marketing/`).

---

## Week 3 — publish `ai-in-excel.html`

1. **Remove the noindex tag** in `ai-in-excel.html`.
2. **Update the date** in three places — `article:published_time`, JSON-LD
   `datePublished` (both `2026-07-18`), and the `legal-meta` line
   (`JULY 18, 2026`).
3. **Add the index card** to `blog/index.html`, above the disasters card
   (update the date):

```html
        <a class="blog-card" href="/blog/ai-in-excel">
          <span class="blog-card-date">July 31, 2026</span>
          <h2>AI in Excel: the ultimate guide to what actually works</h2>
          <p>
            Copilot in Excel, ChatGPT uploads, =AI() formulas, and agents on
            the file — the four ways to put AI in a spreadsheet, what each is
            genuinely good at, and where each quietly fails. Cross-checked
            against what practitioners on r/excel and CFO forums actually
            report.
          </p>
          <span class="blog-read">Read the post →</span>
        </a>
```

4. **Add the sitemap entry**:

```xml
  <url>
    <loc>https://gridpath.dev/blog/ai-in-excel</loc>
    <changefreq>monthly</changefreq>
    <priority>0.6</priority>
  </url>
```

5. **Restore the link in `dead-numbers.html`** (poster-test section). Replace:

```html
        takes ten seconds, and it's the single highest-value habit for
        working with AI in spreadsheets. (For the wider tooling landscape,
        see <a href="/blog/free-excel-alternatives">our honest look at free
        Excel alternatives</a>.)
```

   with:

```html
        takes ten seconds, and it's the single highest-value habit for
        working with AI in spreadsheets. (For where each kind of AI tool
        fits overall, see
        <a href="/blog/ai-in-excel">our guide to AI in Excel</a>.)
```

6. Deploy. Then delete this file.

---

## Distribution cheat sheet

- **Dead numbers (week 1):** LinkedIn post built around the "poster test"
  (change one assumption by 10x, watch nothing move). Pitch to an FP&A
  newsletter or two. Not an HN post.
- **Spreadsheet disasters (week 2):** Submit to Hacker News (plain title,
  morning US time). Share to r/excel / r/Accounting as content, disclose
  you're the author in the comments, engage on the stories not the product.
- **AI in Excel guide (week 3):** SEO asset — don't push hard. Optional
  X/LinkedIn thread of the four-category comparison table. Let search and
  internal links do the work.

---

## Benchmark post — publish `same-model-benchmark.html`

1. **Remove the noindex tag** (line marked `STAGED: remove at publish`).
2. **Update the date** if publishing later than July 19 — `article:published_time`,
   JSON-LD `datePublished`, and the `legal-meta` line.
3. **Add the index card** to `blog/index.html` (top position):

```html
        <a class="blog-card" href="/blog/same-model-benchmark">
          <span class="blog-card-date">July 19, 2026</span>
          <h2>Same model, same edit: benchmarking GridPath against a CLI coding agent</h2>
          <p>
            Five runs each, same Claude model. One harness was byte-clean
            every time; the other silently dropped file content every time —
            and our own audit deleted our best speed number along the way.
            Methodology, caveats, and the two commands to reproduce it.
          </p>
          <span class="blog-read">Read the post →</span>
        </a>
```

4. **Add the sitemap entry**:

```xml
  <url>
    <loc>https://gridpath.dev/blog/same-model-benchmark</loc>
    <changefreq>monthly</changefreq>
    <priority>0.6</priority>
  </url>
```

Distribution: r/excel + r/Accounting angle is the silent-loss finding, not
the speed; HN angle is the "our audit deleted our own best number" section.
Link the eval/ directory in any comments — reproducibility is the pitch.
