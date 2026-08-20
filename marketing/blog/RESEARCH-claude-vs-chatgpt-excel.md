# Research: Claude for Excel vs ChatGPT for Excel (compiled Aug 3, 2026)

Working title candidates:
- "Claude vs ChatGPT for Excel: what practitioners actually report"
- "Claude for Excel vs ChatGPT for Excel: the honest comparison"
- (with original data) "Claude vs ChatGPT on the same spreadsheet task: N runs each"

Target queries: `claude for excel vs chatgpt`, `chatgpt for excel`,
`best ai for excel`, `claude excel add-in review`. FindSkill.ai reports
"claude for excel" is the **#1 rising related query** for "chatgpt for
excel" on Google Trends (May 2026) — demand is live and comparison-shaped.

---

## 1. Product facts (verified against official sources, Aug 3 2026)

### Claude for Excel (Anthropic)
Sources: claude.com/docs/office-agents/excel · support.claude.com article
12650343 · claude.com/claude-for-microsoft-365 · claude.com/pricing

- Sidebar add-in, part of the Claude-for-Microsoft-365 suite (Word,
  PowerPoint GA; Outlook beta). Released Oct 2025, **GA March 11, 2026**.
- **Paid plans only**: Pro ($20/mo, $17 annual), Max, Team, Enterprise.
  Not on Free. No separate fee.
- Model switcher in sidebar: Opus 4.7, Opus 4.6, Sonnet 4.6.
- Capabilities: multi-tab navigation with **cell-level citations**;
  update assumptions preserving formula dependencies; debug REF/VALUE/
  circular refs; build models from scratch; native Excel ops (sort,
  filter, edit pivot tables and charts, conditional formatting, data
  validation, print prep). Supports .xlsx and .xlsm.
- **MCP connectors**: S&P Global, LSEG, Daloopa, PitchBook, Moody's,
  FactSet (per Sonnet 4.6 launch post).
- Overwrite protection (warns before data loss); auto-compaction of long
  conversations.
- **Limitations (Anthropic's own docs)**: no data tables; **no macros/
  VBA**; not recommended for audit-critical calculations or final client
  deliverables without human review.
- **Version wall**: requires M365 Excel — web, Windows build
  16.0.13127.20296+, Mac 16.46+. Does NOT run on Excel 2016/2019
  perpetual/volume licenses or Android. (iPad: help center says
  supported at 2.51+, docs page says unsupported due to SharedRuntime —
  **verify at draft time**.)
- Data handling: inputs/outputs auto-deleted from Anthropic backend
  within 30 days; chat history stored locally in browser, doesn't sync
  across devices; NOT yet in Enterprise audit logs or Compliance API and
  doesn't inherit custom retention settings (material enterprise gap).

### ChatGPT for Excel (OpenAI)
Sources: openai.com/index/chatgpt-for-excel · chatgpt.com/apps/spreadsheets
· help.openai.com article 20001063 · Microsoft Marketplace WA200010215

- Sidebar add-in for **Excel AND Google Sheets**. Beta late 2025 (GPT-5.4
  era) alongside financial data integrations (FactSet, Dow Jones Factiva,
  LSEG, Daloopa, S&P Global); **GA May 5, 2026 across all plans, powered
  by GPT-5.5**.
- **Free tier included** (Free, Go, Plus, Pro, Business, Enterprise, Edu,
  K-12) — big distribution advantage over Claude's paid-only add-in.
  (Initial beta excluded EU for Plus/Pro; GA page says "globally" —
  **verify EU status at draft time**.)
- Claims: builds full spreadsheets from a prompt, insights across tabs,
  links answers to referenced cells, preserves formulas/formatting, asks
  permission before making changes.
- **Limitations (OpenAI's own docs)**: complex formulas/edge cases need
  manual refinement; add-in chats do NOT sync with chatgpt.com history;
  **no memory**; "can make mistakes, including making unintended edits or
  deleting content if a request is unclear"; VBA "may not be fully
  supported"; some outputs need formatting cleanup.
- Data handling: Business/Enterprise/Edu/Teachers excluded from training
  by default; **consumer Free/Plus may be used for training unless the
  user opts out** — the privacy asymmetry vs Claude's 30-day deletion is
  a real point of comparison.

### Copilot in Excel / Agent Mode (Microsoft) — the third wheel
Sources: techcommunity.microsoft.com Excel blog (Agent Mode GA post) ·
learn.microsoft.com wordexcelppt-agents · Glide blog analysis

- Agent Mode GA: Excel web Dec 2025; **Windows/Mac Jan 27, 2026**. Model
  switcher: GPT 5.2, Claude Opus 4.5, or Auto — yes, Copilot can now run
  Claude.
- Requires M365 Copilot license (~$30/user/mo commercial) or M365
  Premium; Personal/Family via AI credits. **Not available in EU or UK.**
- **Files must live in OneDrive/SharePoint** (docs: all agent-created
  documents saved to OneDrive for governance). Local files → Copilot
  greys out. (This is its own blog-post idea, already on the list.)
- **No preview/review mode** — Agent Mode applies changes to the workbook
  directly, in real time. You can stop it, but edits land immediately.
- Only works on the currently open workbook; no cross-file access.

---

## 2. Practitioner sentiment (Reddit, gathered Aug 3 2026)

### r/Accounting — "Am i crazy or is Claude's own plug-in in Excel just
better than Copilot's Claude model" (3mo ago, 189 upvotes, 61 comments)
- OP: Claude's add-in beats Copilot every time — formulas actually work,
  handles multi-step tasks and messy client files.
- Consensus: **Copilot widely considered the worst of the three**;
  refuses complex macros, produces partial snippets, enterprise
  guardrails, drifts mid-chain on multi-step tasks.
- Claude strengths: holds context across multi-step tasks, financial
  statement analysis, formula generation, finding errors in GL data.
- ChatGPT strengths: long-form reasoning, pattern-heavy VBA generation,
  trial-and-error problem solving, explaining code.
- Quote (FP&A): "Claude excel is jaw-dropping good… Finding errors in
  thousands of lines of GL data, setting analytics or underwriting
  models is a good prompt away."

### r/dataanalysis — "What is the best AI tool for working with
spreadsheets" (June 2026, 50 comments)
- Top comment: "I don't think you can get any better than Claude for
  excel" — recommends providing context/MD files describing sheet
  structure.
- Claude limitation: struggles with highly complex / poorly constructed
  files; file organization + provided context is the key variable.
- Useful frame from thread: **two different jobs** — (1) manipulate the
  workbook (Claude good) vs (2) analyze the data (code-sandbox tools:
  Julius, Querri, ChatGPT Data Analyst better).
- Gemini in Sheets rated "utter trash."
- Best-practice tip: export raw data to CSV, let AI work on that with
  code, keep formatting as a separate layer.

### r/FPandA — "OpenAI ChatGPT vs Anthropic Claude vs other AI Chatbots
for FP&A" (9mo ago, ~20 comments)
- Split camp, no clear winner; personal preference + use case.
- Claude: better macros, formulas and M code "less superfluous," less
  sycophantic. ChatGPT: complex formulas, VBA, Power Query, executive
  comms; "feels stronger for analytical and structured work."
- Copilot: works for some VBA/Power Query but inconsistent.
- Universal tip: be very specific/descriptive in prompts.

### r/excel — ChatGPT for M Code thread (2023, 3 comments)
- Claude "never wrote anything that works" for M code — 2023-era, stale;
  useful only as a "how far this has moved" contrast if at all.

---

## 3. Published head-to-head tests (competition to beat, and citable)

- **F9 Finance** (f9finance.com, tested all 3 on same FP&A dataset,
  3 prompts): Claude first "and it wasn't close" — ~90% usable on first
  pass; Copilot 60–70%; **ChatGPT last with structurally broken output**,
  missing a core line item; author "would not use the ChatGPT for Excel
  plugin for core FP&A work right now."
- **FindSkill.ai** (4 tasks side-by-side, May 2026): Claude edge for
  serious financial modeling; ChatGPT edge for speed, polish, Google
  Sheets crossover.
- **aitoolbriefing.com** error-detection test (planted 3 errors in a
  model): Claude 3/3 caught with explanations; ChatGPT 2/3; Copilot 1/3
  (only the circular ref Excel itself flags).
- **matvelloso tweet, May 16 2026** (widely shared): Claude single-shot
  a valid .xlsx; ChatGPT produced a corrupt file after four tries.
  Fidelity angle — connects to our benchmark post.
- **Jea Hun Shin, CPA (LinkedIn)** — honest Claude-limitations material:
  vibe-coded forecast model came out surface-level (drivers too coarse),
  formulas static rather than dynamic, and missed open-pipeline logic —
  "a fundamental error that would materially misstate results." Getting
  deliverable quality required "an extreme level of prompt detail."
- Rework.com "13 tools ranked" and NomadLab/sagnikbhattacharya
  comparisons: feature-table posts, no original testing — this is most
  of the SERP. Only F9 and FindSkill ran real tasks.

Consensus across tests + Reddit: **Claude for model construction,
error-finding, and financial work; ChatGPT for speed, iteration volume,
explanation, and Sheets crossover; Copilot last on capability but wins
on IT-department default.**

---

## 4. What our post has that the SERP doesn't

1. **Same-harness original data.** Every published test compares two
   different products (different harnesses, prompts, UIs) — model and
   harness are confounded. GridPath runs Claude AND GPT models through an
   identical tool surface. `eval/run-gridpath.mjs` drives the real
   product path with whichever model is set in the app; tasks
   `dh-income-statement`, `aapl-forecast`, `fixture-model-edit` are
   ready; `fixture-model-edit` is committed and reproducible by anyone.
   N runs per model → the only apples-to-apples Claude-vs-GPT spreadsheet
   comparison on the internet. (Caveats per eval/README: model parity set
   in app DB; window must stay visible; cached-value warns.)
2. **Practitioner synthesis.** Nobody has aggregated the Reddit threads;
   we can quote real accountants/FP&A folks (with sub + thread cited)
   instead of a one-off test.
3. **Honest-limitations credibility** (house style): Anthropic's own
   "not for audit-critical work" line, the CPA's failed vibe-coded
   forecast, ChatGPT's own "may delete content" warning. We report both
   vendors' self-disclosed limits — nobody else does.

## 5. Where GridPath fits (differentiators vs BOTH add-ins)

- Both add-ins send workbook contents to vendor cloud; GridPath is
  local-first — the file never leaves the machine.
- Neither add-in has reviewable diffs. ChatGPT "asks permission" per
  change; Copilot Agent Mode has **no preview at all**; Claude has
  overwrite warnings. GridPath: every batch is a reviewable diff with
  exact cell-level rewind.
- **Both add-ins require M365 subscription builds of Excel.** Excel
  2016/2019 perpetual-license users — and people without Excel at all —
  are locked out of both. GridPath needs no Excel installed. Underserved
  audience the SERP ignores.
- Same subscriptions: GridPath runs on the Claude Pro/Max or ChatGPT
  Plus/Pro sub the reader already bought for the add-in. The pitch isn't
  "instead of," it's "your $20 already covers both paths."
- No history sync in either add-in (Claude: local browser storage;
  ChatGPT: no sync, no memory). GridPath persists per-workbook sessions
  in SQLite.

## 6. Suggested structure (house style: honest, specific, soft CTA)

1. Lede: both products went GA within weeks of each other in 2026; the
   comparison is the top rising query; here's what the tests, the
   practitioners, and our own runs actually say.
2. What each product actually is (facts, pricing, version walls).
3. What the published tests found (F9, FindSkill, error-detection).
4. What practitioners report (Reddit synthesis, quoted + linked).
5. Where each genuinely fails (both vendors' own disclosed limits +
   the CPA forecast story).
6. [Optional, strongest] Our same-harness runs: same task, same tool
   surface, Claude vs GPT, N runs each — table + reproduce commands.
7. The Copilot question (one section, not a third pillar — link to
   future "Copilot greyed out" post).
8. So which should you use? (by job, mirroring free-alternatives post)
9. Where GridPath fits (legal-note box).

## 7. Distribution

- r/Accounting and r/FPandA both have proven appetite (189-upvote
  thread). Share as content, disclose authorship, engage on findings.
- LinkedIn: FP&A angle via the F9-style "usability on first pass" frame.
- HN: only if the same-harness eval section exists ("model vs harness
  confounding" is the HN-shaped finding); link eval/ for reproducibility.
- Internal links: ai-in-excel (category guide → this is the head-to-head
  depth piece), same-model-benchmark (harness methodology), dead-numbers
  (values-vs-formulas), fidelity page (corrupt-file tweet).

## 8. Open items to verify at draft time

- [ ] Claude add-in iPad support (docs contradict help center).
- [ ] ChatGPT add-in EU availability at GA (beta excluded EU Plus/Pro).
- [ ] Current model versions in both sidebars (moves monthly).
- [ ] Usage limits in practice (Claude auto-compaction; one blogger
      claims ChatGPT has "no token wall" — unverified).
- [ ] Overlap check vs ai-in-excel post: that post compares CATEGORIES
      (copilot / upload / formula / agent); this one compares the two
      named products within the add-in category. Keep the lens distinct;
      cross-link rather than repeat.
