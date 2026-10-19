<!-- Research note, 2026-09-30. Produced by a research pass over primary pages (GitHub READMEs, vendor docs/pricing) and third-party reviews; facts marked (3rd-party) are not from the vendor. Kept as-is for the record; conclusions are in the summary. -->

# Competitive landscape — agents editing .xlsx (2026-09-30)

## Summary (ours)

- Nobody found combines GridPath's three pillars: in-place surgical patch + in-process calc + a visual review page before write. But three entrants converged on "byte-identical untouched parts" in **September 2026 alone**: BetterOffice (lossless + calc + accept/reject in a browser editor, MCP just landed), xlsplice (byte-identical CLI, started Sep 10), GenOffice (8.2k stars, byte-preserving, free/BYOK, one-click skill). "Byte-identical" is becoming a category claim; the window to own it is months.
- The default Excel MCP is haris-musa/excel-mcp-server (4.2k stars, v1.0 on Sep 28). Its own README: formulas "read as empty until the file is opened… in Excel", row/col insert "doesn't update dependent formulas", "shapes, slicers, embedded objects may be lost". That is the honest comparison to make, not Claude Code.
- For agents working on files on disk (Claude Code, Cursor, Codex), the built-in path is the openpyxl skill with documented breakage (claude-code #22044: ~50 investment models corrupted; charts, external links, VBA caveats in Anthropic's own skill). Inside Excel, the giants own it: Copilot Agent Mode (no preview at all), Claude for Excel and ChatGPT for Excel (ask before changes). Our lane is files-on-disk agents plus governance.
- Pricing anchors: Shortcut $100/mo Pro, $320+ teams; GPT for Work $25–45; Copilot $20–30/user; Univer Pro gates xlsx import/export; xlsx-for-ai is hosted (the file leaves the machine) with a free 10k-calls tier.

# Competitive landscape: "AI agent edits .xlsx files" (as of 2026-09-30)

Method: WebSearch + WebFetch of primary pages (GitHub READMEs, vendor docs/pricing) plus third-party reviews where primary pages were blocked (openai.com, help.openai.com, MS TechCommunity, Endex AppSource returned 403). Star counts are as rendered at fetch time. Facts marked (3rd-party) come from blogs/reviews, not the vendor.

---

## 1. MCP servers for Excel files (file-library, no Excel required unless noted)

**haris-musa/excel-mcp-server** — https://github.com/haris-musa/excel-mcp-server
- Python/openpyxl MCP server; create/read/edit workbooks, formatting, charts, "pivot tables", tables, macros. Category: file-library MCP.
- 4.2k stars / 471 forks, MIT. **v1.0 released 2026-09-28** (SSE dropped; stdio + streamable-HTTP).
- Edits in place via openpyxl load/save (re-serialization). Formulas **not calculated**: "Formulas written by this server read as empty until the file is opened and saved in Excel or LibreOffice."
- No review/diff step. Safety features instead: "optional folder confinement, a formula safety check, read-only mode, localhost-only HTTP by default, and atomic saves"; WEBSERVICE/HYPERLINK rejected; 100k-cell cap per call.
- Install: `uvx excel-mcp-server stdio`, pip, **.mcpb bundle for Claude Desktop**, manual config for Cursor/VS Code.
- Stated limits: no real PivotTable creation; row/col insert "doesn't update dependent formulas"; "shapes, slicers, embedded objects may be lost during editing"; no .xls/.csv.
- Free/OSS. The de-facto default Excel MCP by traction.

**negokaz/excel-mcp-server** — https://github.com/negokaz/excel-mcp-server
- Go + Node wrapper; read/write values & formulas, create tables, copy sheet, format range; **Windows-only "Live editing" and screen capture** of a running Excel. ~1.0k stars / 130 forks, MIT.
- Preserves formulas when writing; does not evaluate. No review step. Default 4000-cell paging limit.
- Install: `npx --yes @negokaz/excel-mcp-server`; Smithery one-liner. Node 20+.

**nikhilwoodruff/xlsx (mcp-xlsx)** — https://github.com/nikhilwoodruff/xlsx
- Tiny Python MCP (7 tools: read/write/update_cell/format_cell/list/create/delete sheet). 0 stars, 3 commits. No formula eval, no review. Install: clone + `uv pip install -e .` + `claude mcp add`. Hobby-grade.

**OfficeMCP/ExcelMCP** — https://github.com/OfficeMCP/ExcelMCP
- Live-Excel COM MCP, "Not working on Linux/MacOS"; `RunPython` tool executes openpyxl + Excel app object. 4 stars, GPL-3.0, `uvx excelmcp`. Sibling suite OfficeMCP/OfficeMCP covers Word/Outlook/etc.

**dosev-ai/mcp-office** — https://github.com/dosev-ai/mcp-office/
- "Local-first, governed MCP servers for Microsoft Office" — Excel (65 tools), PPT, Word; Windows 10/11, COM for styling/PDF; some read-only ops without Office. 9 stars, MIT, `pip install mcp-office`. "This is not a Copilot replacement. It is a developer-first execution layer." No documented dry-run/approval.

**excel-com-mcp (benvdbergh/excel-mcp-server)** — https://github.com/benvdbergh/excel-mcp-server
- Fork of haris-musa; "COM-first routing targets the live Excel session when the workbook is open in Microsoft Excel; otherwise… tools use openpyxl on disk." v0.6.0; explicit `save_workbook` (auto-save removed in 0.3.0). 1 star. `uvx excel-com-mcp==0.6.0 stdio`, `pip install excel-com-mcp[com]`. Windows COM / file-only on mac/linux. No review step.

**sbroenne/mcp-server-excel (ExcelMcp)** — https://github.com/sbroenne/mcp-server-excel
- Live-Excel COM automation: "31 tools with 326 operations" incl. Power Query, DAX, VBA, Python-in-Excel; **real recalculation via Excel**. 794 stars, MIT. Requires "Windows, Microsoft Excel 2016 or later, and an interactive desktop." Install: VS Code extension, MCP config, `excelcli` dotnet tool ("substantially lower token usage" CLI vs MCP). Claims to preserve "PivotTables, charts, macros, the Data Model, and workbook formatting" (because Excel itself saves). No approval step. Open bugs show Excel-side re-save side effects, e.g. issue #881 "Saving through the server rewrites Excel Table column default number formats… from ja-JP built-in definitions to en-US" — https://github.com/sbroenne/mcp-server-excel/issues/881

**Other live-COM MCPs (brief):** mroshdy91/Excellm (dual COM/openpyxl, 34 tools, `preview_only=True` on find_replace, v1.1.0 2026-01-05, 1 star) https://github.com/mroshdy91/Excellm ; kousunh/Excel-mcp-server (live or closed files) https://github.com/kousunh/Excel-mcp-server ; ThepExcel/ThepExcelMCP (Windows COM) https://github.com/ThepExcel/ThepExcelMCP ; gvantage/mcp-server-excel (COM, 23 tools/214 ops) https://github.com/gvantage/mcp-server-excel ; dsbissett/office-addin-mcp (Excel via WebView2 CDP) https://github.com/dsbissett/office-addin-mcp

**xlsx-for-ai (senoff)** — https://xlsx-for-ai.dev/ ; announced on Codex discussions 2026-05-10: https://github.com/openai/codex/discussions/22069
- **Hosted** MCP: "thin npm client (~200 lines) over a hosted API at https://api.xlsx-for-ai.dev". 37–50 tools depending on page: read/write, `xlsx_diff` ("deterministic cell- and formula-level diff between two workbooks"), `xlsx_doctor` (health report ranked HIGH/MEDIUM/LOW), macros, pivots, charts, CF, protection, "stateful supervisor sessions".
- **Calculates formulas**: "382 Excel functions… portable engine—not stale cached values"; excludes INDIRECT/WEBSERVICE/RTD/DDE.
- Fidelity claim: named ranges, merged cells, DV, CF "treated as first-class elements rather than collapsed"; "cross-engine validation" to detect corruption. Per-fix approval only inside `xlsx_data_clean`.
- Pricing: free tier 10,000 calls/month, anonymous UUID auto-registration. MIT client. **Your file leaves the machine** (hosted) — the main contrast with GridPath.

**PSU3D0/agent-spreadsheet + spreadsheet-mcp (spreadsheet-kit)** — https://github.com/PSU3D0/agent-spreadsheet , https://github.com/psu3d0/spreadsheet-mcp , https://lib.rs/crates/spreadsheet-mcp
- Rust; CLI `asp`, MCP server, TS SDK, WASM engine. Formula engine **Formualizer** (same author; "400+ functions, dynamic arrays") with LibreOffice as opt-in alternative. https://github.com/PSU3D0/formualizer
- **Closest OSS analogue to GridPath's review model**: "fork-based editing with staged changes", checkpoints, "changeset review", `asp verify diff` ("grouped workbook diff"), `--dry-run`, "safe mutation, not blind mutation". Writes require `--recalc-enabled`. Review is CLI/JSON, not a visual page.
- 58 stars; spreadsheet-mcp v0.11.1 (2026-07-20). Apache-2.0. Install: shell installer, npm, cargo, Homebrew, Docker. Write-fidelity of untouched parts not documented.

**Smaller file MCPs:** theluckystrike/mcp-spreadsheet (TS; never overwrites unless `overwrite`; "Writing an xlsx replaces one sheet" and formulas/CF/charts on that sheet become "plain values"; .mcpb + hosted token; 0 stars) https://github.com/theluckystrike/mcp-spreadsheet ; jwadow/mcp-excel (Python, **read-only** "SQL for Excel", 46 stars) https://github.com/jwadow/mcp-excel ; guillehr2/Excel-MCP-Server-Master (Node, `npx @guillehr2/excel-mcp-server`, 34 stars) https://github.com/guillehr2/Excel-MCP-Server-Master ; mort-lab/excel-mcp (openpyxl, on Smithery) https://github.com/mort-lab/excel-mcp ; ArchimedesCrypto/excel-reader-mcp (SheetJS, read-only w/ chunking) https://github.com/ArchimedesCrypto/excel-reader-mcp ; office_oxide (Rust multi-lang; `office-oxide-mcp` with extract/replace_text/info; 142 stars) https://github.com/yfedoseev/office_oxide ; Spire.XLS MCP (commercial lib) https://glama.ai/mcp/servers/eiceblue/spire-xls-mcp-server

**Cloud/Graph-API Excel MCPs (edit files in OneDrive/SharePoint, not local):** Softeria ms-365-mcp-server https://mcpservers.org/servers/softeria/ms-365-mcp-server ; Arcade.dev Office 365 MCP https://www.arcade.dev/blog/microsoft-office-365-mcp-servers-launch/ ; StackOne https://www.stackone.com/connectors/microsoftexcel/mcp/ ; Aanerud MCP-Microsoft-Office (117 tools) https://mcpservers.org/servers/Aanerud/MCP-Microsoft-Office

**Google Sheets MCPs (brief):** Google now ships an **official remote Sheets MCP server** (Workspace Developer Preview) https://developers.google.com/workspace/sheets/api/guides/configure-mcp-server and an official Workspace CLI + MCP (per Wes Bos) https://x.com/wesbos/status/2029373589770703148 ; community: mkummer225/google-sheets-mcp https://github.com/mkummer225/google-sheets-mcp , henilcalagiya, FinnaA. All operate on cloud Sheets via API — no local .xlsx.

**Generic review layer:** vscode-diff-mcp — agent calls `edit_with_diff`, VS Code shows native diff, "click Accept or Reject… nothing is written on reject" (text files, not xlsx) https://marketplace.visualstudio.com/items?itemName=iamseb4s.vscode-diff-mcp

---

## 2. Anthropic's official `xlsx` skill / claude.ai file creation

- Skill: https://github.com/anthropics/skills/blob/main/skills/xlsx/SKILL.md — openpyxl for formulas/formatting, pandas for bulk, markitdown preview, **LibreOffice via `scripts/recalc.py`** for recalculation (opens+saves through soffice, so the file is re-serialized by LibreOffice).
- Key rules quoted: "Editing an existing file: match its conventions exactly… write only there, and leave every existing formula untouched"; "openpyxl writes formulas as strings with no cached values"; "`errors_found` exits 0, so never treat a clean exit as a clean workbook."
- Stated limitations: `data_only=True`+save "all formulas replaced with literals (permanent)"; `.xlsm` loses macros unless `keep_vba=True`; "External file links are stripped when re-saved with openpyxl; recalc.py refuses to run unless --force"; "Never use: XLOOKUP, XMATCH, SORT, FILTER, UNIQUE, SEQUENCE (LibreOffice cannot evaluate)"; merged cells only via anchor.
- openpyxl doc itself: "images and charts will be lost from existing files if they are opened and saved with the same name" https://openpyxl.readthedocs.io/en/stable/tutorial.html
- Field evidence of breakage: claude-code issue #22044 (opened 2026-01-30, "high-priority", closed not planned): "for more realistic financial models (with .xlsm macros, complex conditional formatting, named ranges, etc.) openpyxl silently strips or breaks parts of the file structure"; user had ~50 investment models corrupted https://github.com/anthropics/claude-code/issues/22044 . Also #69106 (openpyxl chart XML bug) https://github.com/anthropics/claude-code/issues/69106
- Lumetric on Claude Cowork: "The last 10-20% is where things break"; form controls/ActiveX "frequently stripped or broken"; CF ranges "may not adjust" https://www.lumetric.ai/resources/how-claude-cowork-makes-excel-files-and-why-it-is-not-a-perfect-solution
- claude.ai "Create files": runs Python in a sandbox, 30MB/file, all plans; outputs a new download rather than an in-place edit (per support article summary) https://support.claude.com/en/articles/12111783-create-and-edit-files-with-claude
- OpenAI equivalent: openai/skills `spreadsheet` (openpyxl+pandas; "Modify existing workbooks without breaking formulas, references, or formatting"; recalculation only "if an internal spreadsheet recalculation/rendering tool is available") https://skills.sh/openai/skills/spreadsheet (repo now deprecated in favor of OpenAI Plugins) https://github.com/openai/skills/blob/main/README.md
- Community zip-surgery skill (gist, 2026-09-01): "Change only the parts required for the requested edit, and copy every other ZIP member through unchanged"; checklist requires styles.xml/theme1.xml/unchanged sharedStrings "match the original byte-for-byte" https://gist.github.com/telotortium/844386f762c4b3bab49999ba99f72f5b

---

## 3. Univer (dream-num) — "The Office Harness for AI Agents"

- Core SDK: 22.2k stars, Apache-2.0; xlsx import/export is a **Univer Pro / closed-source** feature (free to use, paid upgrade) https://github.com/dream-num/univer , https://docs.univer.ai/guides/sheets/features/import-export
- Site claims: "Multi-Agent Workflows on Worktrees", "Parallel by default, human-reviewed by design", "Humans review and merge every contribution", "Git-Style Version History", "Rust-based formula runtime… one million formulas in seconds", "500+ built-in functions", "High-fidelity import and export for XLSX…" https://univer.ai/
- **univer-cli** (13 stars, Apache-2.0, Node 24+, Chrome for screenshots): imports .xls/.xlsx/.xlsm/.csv, exports .xlsx/.csv/.tsv — i.e. **convert-to-.univer and re-export, not in-place**. "Changes stay in an isolated Worktree… returns a local Viewer URL so you can review the result and decide whether to merge, revise, or discard." Install `npx skills add dream-num/skills -s univer-cli -g` https://github.com/dream-num/univer-cli
- **dream-num/skills** (75 stars): "brings Git-style diffs, reviews, approvals, and rollbacks to spreadsheet workflows" https://github.com/dream-num/skills
- **univer-workspace** (2.2k stars, self-host Docker, Node 24+/pnpm): "Intermediate changes remain isolated from shared content until a person accepts them." https://github.com/dream-num/univer-workspace
- **univer-mcp** (38 stars): proxies tool calls to a running Univer instance; needs API key from console.univer.ai; Cursor one-click, Claude Code, Gemini CLI https://github.com/dream-num/univer-mcp
- Also a DeepSeek Harness plugin https://github.com/dream-num/dsh-univer-office

---

## 4. First-party

**Microsoft Copilot / Agent Mode in Excel** (add-in native)
- GA: Excel web Dec 2025; **Windows/Mac 2026-01-27**; Word/Excel/PPT all GA and "now the default experience" 2026-04-22; OpenAI + Anthropic model choice; "not available in the EU or UK" https://office-watch.com/2026/copilot-agent-mode-word-excel-powerpoint/ , https://techcommunity.microsoft.com/blog/excelblog/agent-mode-in-excel-is-now-generally-available-on-desktop/4457408 (3rd-party summaries: https://pasqualepillitteri.it/en/news/1401/microsoft-copilot-agent-mode-word-excel-powerpoint-april-2026 )
- Renamed "Edit with Copilot" 2026-03-09; AutoSave/OneDrive no longer required so **local files work**; "Advanced Analysis with Python currently missing" (3rd-party) https://windowsforum.com/threads/excel-copilot-consolidation-copilot-chat-and-agent-mode-replace-app-skills.403686/ ; Python now usable inside Edit with Copilot (April 2026 What's New) https://techcommunity.microsoft.com/blog/excelblog/whats-new-in-excel-april-2026/4502696
- **No review step**: "There's no preview mode where you review changes before they're applied" (Glide, 2026-01-27) https://www.glideapps.com/blog/excel-agent-mode ; sidebar "shows each step"; unsaved Ctrl+N workbooks unsupported.
- Edits the live workbook in place (Excel is the engine, so full calc + fidelity). Price: M365 Copilot $30/user/mo enterprise, Copilot Business $21, M365 Premium $19.99 https://justinmckelvey.com/blog/microsoft-365-copilot-pricing

**Claude for Excel (Anthropic)** — https://claude.com/docs/office-agents/excel
- GA on Pro/Max/Team/Enterprise; Excel web / Windows M365 / Mac 16.46+; via AppSource "Claude for Microsoft 365"; beta Oct 2025 → Pro rollout 2026-01-24 https://the-decoder.com/anthropic-opens-claudes-improved-excel-integration-to-all-pro-subscribers-after-limited-beta/
- Official: "cell-level citations", "Adjust assumptions while keeping formula relationships intact", pivot/CF/DV/sort/filter native ops, connectors (S&P, LSEG, Daloopa), Skills. **"Overwrite protection: Claude warns you before overwriting existing data"**; "When Claude proposes a risky operation, you are asked to confirm before it runs." Best practice: "Always review changes before finalizing." Unsupported: data tables, macros/VBA. Explicit prompt-injection warning. Pro $20/mo (3rd-party) https://screenapp.io/blog/claude-ai-pricing
- Third-party reviews claim highlighted diffs/accept-reject (e.g. https://agentscamp.com/guides/analytics/claude-for-excel-guide ) — the official doc only promises overwrite warnings + risky-op confirmation; "no native rollback" (3rd-party). Live add-in: Excel calculates; edits in place.

**ChatGPT for Excel (OpenAI)** — official post (403 to fetch): https://openai.com/index/chatgpt-for-excel/
- Beta 2026-03-05/06 (US/CA/AU; Plus/Pro/Team/Business/Enterprise/Edu), Sheets 2026-04-22, **GA all plans 2026-05-05 (GPT-5.5)** (3rd-party) https://www.eweek.com/news/openai-chatgpt-excel-gpt-5-4-launch/ , https://pasqualepillitteri.it/en/news/1963/chatgpt-for-excel-google-sheets-complete-guide-2026
- "asks permission before making any changes so users can review or undo"; cell links; data partners Moody's, FactSet, MSCI, Factiva. **Codex** in ChatGPT desktop can drive the open workbook through the add-in (help.openai.com, 403) https://help.openai.com/en/articles/20001063-chatgpt-for-excel-and-google-sheets
- ChatGPT Work also creates/edits xlsx as files (sandbox) https://help.openai.com/en/articles/20001278-creating-and-editing-documents-spreadsheets-and-presentations-with-chatgpt-work

**Gemini in Google Sheets** — https://workspace.google.com/resources/spreadsheet-ai/
- Side panel builds tables/formulas/charts/pivots/CF/dropdowns and "multi-step spreadsheet changes"; =AI()/=Gemini() functions; "Fill with Gemini" (Apr 2026) https://workspaceupdates.googleblog.com/2026/04/effortlessly-automate-data-entry-in-Google-Sheets-using-Fill-with-Gemini.html ; included in Workspace Business Standard+. Cloud Sheets only; no local xlsx; no documented preview gate.

---

## 5. Startups / add-ins

| Product | What / category | In-place on user's .xlsx? | Formula eval | Review gate | Price | Notes / URL |
|---|---|---|---|---|---|---|
| **Shortcut** | Finance agent; Excel add-in + web + Windows desktop + Sheets plugin + **CLI** | Yes (add-in, Excel calcs) | Excel | "Approve all edits" | Free 20 credits/wk; Pro $100/mo; Teams $320+$100/seat | Claims 76% vs Claude-for-Excel 66% / Copilot 59%; "3 of the 5 largest multi-strat hedge funds" (3rd-party). https://shortcut.ai/ , https://shortcut.ai/pricing |
| **Endex** | Excel-native finance agent (sidebar add-in), OpenAI Startup Fund | Yes | Excel | not documented | Undisclosed, sales/waitlist (3rd-party) | SOC2/ISO27001, ZDR. https://www.fixedlabs.ai/tools/endex , AppSource wa200008783 |
| **Deckary** | "agentic AI Excel add-in" for consultants/finance (2026 entrant) | Yes | Excel | "pauses before destructive edits"; "ask before destructive edits loop" | $180/yr Premium | Blog dated 2026-05-10 frames "Excel agents vs AI spreadsheets". https://deckary.com/blog/spreadsheet-ai |
| **GPT for Work (Talarian)** | Excel/Sheets agent add-in | Yes | Excel/Sheets | n/d | Standard $25/mo, Business $45/seat; credit packs from $29 (subs introduced Jun 2026) | https://gptforwork.com/pricing |
| **Numerous.ai** | =AI() formulas add-in (Excel/Sheets) | cell functions only | host app | n/a | Personal $8/mo, Pro $24/mo (yearly) | https://www.buildfastwithai.com/ai-tools/numerous-ai |
| **Formula Bot** | Formula generator + data analyzer + add-on | mostly generates text/formulas | host app | n/a | ~$9–30/mo, free tier | https://www.toolsforhumans.ai/ai-tools/excel-formula-bot |
| **Julius AI** | Chat-with-data analysis; uploads xlsx, outputs charts/new files | No (analysis/export) | Python | n/a | Free 15 msgs; Plus $35; Pro $45 | "designed for data analysis rather than direct Excel file editing". https://dupple.com/tools/julius-ai |
| **Quadratic** | AI-native spreadsheet (Python/SQL cells), SOC2/HIPAA | No — import into its own grid | own engine | n/a | Free; Pro $18/user/mo (+$20 AI credits); Business $36 | https://www.quadratichq.com/pricing |
| **Rows** | AI spreadsheet SaaS | No (import/export) | own | n/a | — | **Acquired by Superhuman Feb 2026, wound down 2026-05-31** (3rd-party) https://www.buildfastwithai.com/ai-tools/rows |
| **Equals** | Finance spreadsheet w/ warehouse connectors + "Analyst" AI | No | own | n/a | from $1,250/mo (3rd-party) | https://equals.com/ |
| **Sourcetable** | "self-driving spreadsheet"; imports .xls/.xlsx up to 10GB | No (import) | own | "real-time code evaluation loop… tests and verifies AI-generated actions before they are implemented" (PR claim) | Free 50 credits; Pro $29; Max $100 | https://sourcetable.com/pricing , https://www.unite.ai/sourcetable-raises-4-3m-... |
| **Paradigm** | Agent-per-cell research spreadsheet | No | own | n/a | free tier + paid; 10k paid users, $7M raised (3rd-party) | https://paradigmai.com/ |
| **Bricks** | AI-native spreadsheet/dashboards | No | own | n/a | Free 20 msgs; $25/seat | https://www.saasworthy.com/product/bricks |
| **GenOffice (Genspark)** | Free OSS AI office suite + `genoffice` CLI + agent skill for Claude Code/Codex/Cursor | **Yes, byte-preserving** ("Only what you edit is rewritten. Everything else… survives byte-for-byte") | "live formulas" (engine n/d) | tracked changes + one-click rollback in Docs; diffs shown | Free, BYOK | 8.2k stars, Apache-2.0 (ee/ reserved); mac/win/linux; one-click skill install. https://github.com/genspark-ai/genoffice |
| **HermesOffice** | OSS Electron office suite; Rust sidecar **calamine + IronCalc** | partial (sidecar import/export) | IronCalc | "block-level edits with snapshots and diffs" | Free | 593 stars. https://github.com/criptogus/HermesOffice |
| **BetterOffice (openooxml)** | Native OOXML Rust engines → WASM; xlsx/docx/pptx packages (npm/crates/PyPI); MCP server + agent tools added Sep 2026 | **Yes**: "preserves untouched file parts losslessly when round-tripping"; "version-checked atomic edit batches" | own; "99.60%" recalc accuracy, 95ms avg (README benchmark) | **"Review attributed agent edits through tracked changes, inline diffs, and before-and-after previews. Accept or reject changes directly in the editor."** | Apache-2.0 | 258 stars. https://github.com/openooxml/betteroffice , https://betteroffice.dev/ , https://github.com/openooxml/betteroffice/pull/871 |
| **xlsplice** | Rust CLI "surgical edits to Excel packages"; ships SKILL.md | **Yes**: "Every part it does not target stays byte-identical; there is no formula engine and no whole-file re-serialisation." | none | none | MIT | 0 stars; **started 2026-09-10**. https://github.com/niko86/xlsplice |
| **omnidoc (Go lib)** | "preservation-first" xlsx editor | Yes: "Edit followed by Save of an untouched workbook is part-for-part byte-identical" | "Formulas are not evaluated: the cached <v> is the value" | none (library) | MIT | https://pkg.go.dev/github.com/nathanstitt/omnidoc/pkg/xlsx |

"Cursor for Excel": one builder killed the project Oct 2025 citing saturation (Endex, Shortcut, Copilot, Claude, Sheets AI) https://techforest.substack.com/p/the-journey-building-a-cursor-for ; HN thread https://news.ycombinator.com/item?id=46495502 . "Tally" is not a real spreadsheet-agent product in results.

---

## 6. Libraries under the hood (preserve untouched parts? / calculates?)

- **openpyxl** — re-serializes whole file; "images and charts will be lost from existing files"; no calc; strips external links, VBA unless keep_vba. https://openpyxl.readthedocs.io/en/stable/tutorial.html
- **XlsxWriter** — write-only: "cannot read or modify an existing Excel XLSX file"; no calc. https://xlsxwriter.readthedocs.io/introduction.html
- **SheetJS CE** — re-serializes; basic styles only (Pro for styles/CF); reads+writes formulas; calc only in **SheetJS Pro** formula calculator. https://docs.sheetjs.com/docs/csf/features/formulae/
- **ExcelJS** — re-serializes; known to drop/corrupt conditional formatting on round-trip (issues #1024, #2178, #1583); no calc. https://github.com/exceljs/exceljs/issues/1024
- **IronCalc** — Rust engine, 4.2k stars, MIT/Apache; "300+ functions, LET and LAMBDA"; xlsx reader/writer (own model, not part-preserving; v1.0 "mid 2026" roadmap w/ dynamic arrays). https://github.com/ironcalc/IronCalc , https://www.ironcalc.com/
- **Formualizer** — Rust/Py/JS/WASM engine, "400+ functions, dynamic arrays", MIT/Apache; powers agent-spreadsheet. https://github.com/PSU3D0/formualizer
- **formulas (vinci1it2000)** — Python interpreter; compiles workbooks, calculates without Excel; EUPL. https://github.com/vinci1it2000/formulas
- **pycel** — compiles Excel→Python graph; 632 stars; calc only (reads via openpyxl). https://github.com/dgorissen/pycel
- **LibreOffice headless** — full recalculation but re-saves through Calc's model (fidelity "very good for typical… complex layouts can shift"); cannot evaluate XLOOKUP/FILTER/etc per Anthropic skill. https://www.converterer.com/blog/libreoffice-headless/
- **Aspose.Cells** — commercial; embedded calc engine (`CalculateFormula`) + high-fidelity load/save (own model). https://docs.aspose.com/cells/net/formula-calculation-engine-in-aspose-cells/
- **calamine / umya-spreadsheet (Rust)** — calamine read-only; umya read/write (re-serialize). https://github.com/MathNya/umya-spreadsheet

---

## (a) Comparison table

| Product | In-place edit of user's .xlsx | Formula eval | Review before write | Needs installed | Fidelity claim | Price |
|---|---|---|---|---|---|---|
| haris-musa excel-mcp-server | openpyxl re-save | No ("read as empty until opened in Excel") | No (atomic saves, formula safety check) | uvx/pip/.mcpb; no Excel | Warns shapes/slicers/objects "may be lost" | Free |
| negokaz excel-mcp-server | Go re-save; live COM on Windows | No | No | npx; Excel optional (Win) | none | Free |
| nikhilwoodruff mcp-xlsx | openpyxl | No | No | uv + clone | none | Free |
| sbroenne ExcelMcp | Live Excel (COM) | Excel | No | Windows + Excel 2016+ | "preserves PivotTables, charts, macros, Data Model" | Free |
| OfficeMCP/ExcelMCP, dosev mcp-office, excel-com-mcp, Excellm | Live Excel COM (openpyxl fallback) | Excel when live | No (Excellm preview_only on find/replace) | Windows + Excel | via Excel | Free |
| xlsx-for-ai | Hosted API rewrites file | Yes, 382 fns | Per-fix approve in data_clean only | npm client + cloud | "preserving structure", cross-engine validation | Free 10k calls/mo |
| PSU3D0 agent-spreadsheet / spreadsheet-mcp | Fork → materialize (Rust) | Yes, Formualizer 400+ / LibreOffice | Yes: dry-run, staged changeset, `verify diff` (CLI/JSON) | binary/cargo/npm/brew | "safe mutation"; part-fidelity n/d | Free |
| theluckystrike mcp-spreadsheet | New file by default; sheet replaced → values only | expression lang | No | .mcpb / hosted | none (destroys formulas on written sheet) | Free |
| Anthropic xlsx skill / claude.ai files | openpyxl re-save; new download | LibreOffice recalc.py | No | claude.ai / Claude Code sandbox | Explicit caveats (charts/links/VBA) | Claude plans $0–200 |
| OpenAI spreadsheet skill / ChatGPT Work | openpyxl | "if available" | No | Codex/ChatGPT | "without breaking formulas" | ChatGPT plans |
| Univer CLI / Workspace / MCP | Import→.univer→export .xlsx | Univer engine "500+" | **Yes**: worktree + Viewer/merge/discard, "Git-style diffs" | Node 24+, Chrome; Pro key for import/export | "High-fidelity import and export" | OSS + Univer Pro |
| Copilot Agent Mode / Edit with Copilot | Live Excel | Excel | **No** ("no preview mode") | Excel + Copilot license | Excel-native | $19.99–30/user/mo |
| Claude for Excel | Live Excel add-in | Excel | Overwrite warning + risky-op confirm | Excel (web/Win/Mac) | "keeping formula relationships intact" | Pro $20+ |
| ChatGPT for Excel | Live Excel add-in | Excel | "asks permission before making any changes"; undo | Excel | n/d | all ChatGPT plans (GA May 2026) |
| Gemini in Sheets | Cloud Sheets | Sheets | No documented gate | Workspace Business Std+ | n/a (.xlsx not native) | bundled |
| Shortcut | Live Excel add-in / web / CLI | Excel | "Approve all edits" | Excel or browser | 76% benchmark claim | $0 / $100 / $320+ |
| Endex / Deckary / GPT for Work | Live Excel add-in | Excel | Deckary "pauses before destructive edits"; GPTfW n/d; Endex n/d | Excel | n/d | Endex undisclosed; Deckary $180/yr; GPTfW $25+ |
| Quadratic / Sourcetable / Paradigm / Equals / Bricks | Import into SaaS grid | own | Sourcetable "verifies… before implemented" | browser | n/a | $0–$1,250/mo |
| GenOffice | **Yes, byte-preserving patches** | "live formulas" | tracked changes + rollback (docs); diffs | desktop app (mac/win/linux) + skill | "Everything else… survives byte-for-byte" | Free, BYOK |
| BetterOffice | **Yes, lossless untouched parts** | own, 99.6% bench | **Yes**: tracked changes, inline diffs, before/after, accept/reject | npm/crates/pip; browser | "preserves untouched file parts losslessly" | Apache-2.0 |
| xlsplice | **Yes, byte-identical untouched parts** | None | None | Rust binary | "no whole-file re-serialisation" | MIT |

## (b) Features competitors sizzle that gridpath.dev/mcp (per your description) doesn't obviously surface

1. **Benchmarks/accuracy numbers** — Shortcut leads with "76% vs 66% vs 59%" and $/task; BetterOffice publishes recalc accuracy (99.60%, 95ms). GridPath has eval data (SpreadsheetBench 87.5% shard) it could publish.
2. **Formula-function count** — xlsx-for-ai "382 functions", Formualizer "400+", Univer "500+", IronCalc "300+". State your IronCalc-fork coverage (and dynamic-array status) explicitly.
3. **Diff/doctor tools as named MCP tools** — xlsx-for-ai's `xlsx_diff` (two arbitrary workbooks) and `xlsx_doctor` health report; agent-spreadsheet's `verify proof`. Your review page diffs a session; a general workbook-vs-workbook diff tool is a gap.
4. **Git-style vocabulary** — Univer/BetterOffice sell "worktrees", "merge/discard", "tracked changes", "rollback", CRDT co-editing. GridPath has accept/reject/save but no named rollback/version history.
5. **Safety knobs** — haris-musa: folder confinement, read-only mode, formula safety check (rejects WEBSERVICE/HYPERLINK), cell caps; Claude for Excel: prompt-injection warning. Worth listing GridPath's `--review-required`, allowed-folders and any read-only mode as first-class bullets.
6. **Multi-client one-liners** — competitors show Cursor one-click, Codex `config.toml`, Gemini CLI, VS Code, Smithery, `npx skills add`. GridPath lists Claude Desktop .mcpb + Node; add Codex/Cursor/Gemini snippets and a SKILL.md (xlsplice, univer-cli, GenOffice all ship skills).
7. **Live-Excel / VBA / PivotTable / Power Query coverage** — sbroenne (326 ops), Copilot, Claude for Excel do things a file patcher can't; be explicit about what GridPath deliberately doesn't do (VBA execution, Power Query refresh) and what it preserves untouched (macros, pivots, charts, external links) — the exact things openpyxl-based tools break.
8. **Charts/pivot creation** — haris-musa, xlsx-for-ai, Univer advertise creating charts/pivots; GridPath's page should say whether creation is supported or only preservation.
9. **Hosted/free-tier framing** — xlsx-for-ai "10,000 calls/month free, no email"; GenOffice "free, BYOK". GridPath's local-only/no-upload stance is a differentiator to state plainly against xlsx-for-ai.
10. **Screenshots/render verification** — negokaz screen capture, univer-cli Viewer + screenshots, OpenAI skill "render for visual review". A rendered preview of changed regions is a sizzle item.
11. **Desktop app + CLI both** — Shortcut and Univer sell add-in + web + desktop + CLI; GridPath has MCP + desktop app; a CLI mode (agent-spreadsheet `asp`, xlsplice, `excelcli`) is a common ask.

## (c) Competitors making claims similar to GridPath's

**"Review before save / accept-reject":**
- Univer: "human-reviewed by design", worktree → Viewer URL → "merge, revise, or discard"; skills repo: "Git-style diffs, reviews, approvals, and rollbacks" https://github.com/dream-num/univer-cli , https://github.com/dream-num/skills
- BetterOffice: "before-and-after previews. Accept or reject changes directly in the editor" https://betteroffice.dev/
- PSU3D0 spreadsheet-mcp: "changeset review", staged forks, `asp verify diff` https://lib.rs/crates/spreadsheet-mcp
- Claude for Excel: "Overwrite protection… warns you before overwriting"; "asked to confirm before it runs" https://claude.com/docs/office-agents/excel
- ChatGPT for Excel: "asks permission before making any changes so users can review or undo" https://www.eweek.com/news/openai-chatgpt-excel-gpt-5-4-launch/
- Shortcut: "Approve all edits" https://shortcut.ai/ ; Deckary: "pauses before destructive edits" https://deckary.com/blog/spreadsheet-ai
- xlsx-for-ai: per-fix approval in `xlsx_data_clean` https://xlsx-for-ai.dev/
- vscode-diff-mcp (generic files): accept/reject in native diff, "nothing is written on reject" https://marketplace.visualstudio.com/items?itemName=iamseb4s.vscode-diff-mcp
- Counter-example to exploit: Copilot Agent Mode "no preview mode where you review changes before they're applied" https://www.glideapps.com/blog/excel-agent-mode

**"Byte-identical / untouched parts preserved":**
- xlsplice (Sep 2026): "Every part it does not target stays byte-identical… no whole-file re-serialisation" https://github.com/niko86/xlsplice
- GenOffice: "Only what you edit is rewritten. Everything else in the file survives byte-for-byte" https://github.com/genspark-ai/genoffice
- omnidoc: "part-for-part byte-identical… re-serializes only the XML parts it dirtied" https://pkg.go.dev/github.com/nathanstitt/omnidoc/pkg/xlsx
- BetterOffice: "preserves untouched file parts losslessly when round-tripping" https://github.com/openooxml/betteroffice
- telotortium gist skill: "copy every other ZIP member through unchanged… byte-for-byte" https://gist.github.com/telotortium/844386f762c4b3bab49999ba99f72f5b
- HermesOffice (docx side): "only dirty paragraphs are regenerated… byte-for-byte" https://github.com/criptogus/HermesOffice
- sbroenne (via Excel): "Preserves PivotTables, charts, macros, the Data Model, and workbook formatting" https://github.com/sbroenne/mcp-server-excel

**Nobody found combines all three of GridPath's pillars** (surgical in-place patch + in-process WASM calc + visual local review page before write). Closest: BetterOffice (lossless + calc + accept/reject, but browser-editor centric, MCP just landed Sep 2026), agent-spreadsheet (calc + staged diff, no visual page, fidelity unstated), xlsplice (byte-identical, no calc, no review), Univer CLI (review, calc, but convert-and-export rather than in-place).
