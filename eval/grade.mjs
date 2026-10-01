#!/usr/bin/env node
/**
 * GridPath eval grader — layout-independent assertions against an output xlsx.
 *
 *   node eval/grade.mjs <task-id|task.json> <output.xlsx> [--original <start.xlsx>]
 *
 * Assertions are label-driven (find the row whose leading text matches a
 * regex, then check its cells), because different harnesses produce
 * different layouts for the same task. Three outcomes per assertion:
 *   pass | fail | warn (warn = can't be evaluated fairly, e.g. a formula
 *   cell with no cached value in the file — needs a recalc pass).
 * Exit code is 1 if any assertion hard-fails.
 */

import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const ExcelJS = require("exceljs");
const JSZip = require("jszip");

const evalDir = path.dirname(fileURLToPath(import.meta.url));

// ExcelJS crashes on drawing parts it can't parse (chart-only drawings,
// userShapes — parsed to undefined, then dereferenced). Deleting the entry
// (the app's older guard) just moves the crash: the worksheet still holds
// the drawing rel and dereferences `options.drawings[name].anchors`. Stub
// the entry with empty anchors instead so reconcile walks nothing.
{
  const proto = Object.getPrototypeOf(new ExcelJS.Workbook().xlsx);
  const orig = proto.reconcile;
  proto.reconcile = function (model, options) {
    if (model?.drawings) {
      for (const k of Object.keys(model.drawings)) {
        if (!model.drawings[k]) model.drawings[k] = { anchors: [] };
      }
    }
    return orig.call(this, model, options);
  };
}

function loadTask(spec) {
  const p = fs.existsSync(spec) ? spec : path.join(evalDir, "tasks", `${spec}.json`);
  return { task: JSON.parse(fs.readFileSync(p, "utf8")), path: p };
}

// --- recalc normalization (--recalc) ---------------------------------------
// Value assertions can only read CACHED formula results — neither harness is
// guaranteed to save them (GridPath's surgical save strips them; openpyxl
// can't compute them, or worse leaves stale ones). Normalize through
// LibreOffice headless with a profile seeded to "always recalculate on
// load", so every formula cell carries a fresh, truthful cached value. Only
// VALUE-domain assertions read the recalced copy; structure-domain
// assertions (formulas, zip parts) keep reading the raw output — the LO
// round-trip must never touch the fidelity axis.

function findSoffice() {
  const candidates = [
    process.env.SOFFICE_PATH,
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
  ].filter(Boolean);
  for (const c of candidates) if (fs.existsSync(c)) return c;
  const which = spawnSync("which", ["soffice"], { encoding: "utf8" });
  return which.status === 0 ? which.stdout.trim() : null;
}

const RECALC_XCU = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
`;

/** Returns the recalced copy's path (written next to the output as
 *  *.recalc.xlsx), or null with a printed warning. */
function recalcFile(outputPath) {
  const soffice = findSoffice();
  if (!soffice) {
    console.warn("  ! --recalc: LibreOffice (soffice) not found — set SOFFICE_PATH. Grading raw file.");
    return null;
  }
  // Throwaway profile per invocation: avoids the soffice singleton lock when
  // grades ever run concurrently, at the cost of ~2s profile init.
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "gp-recalc-"));
  const profile = path.join(tmp, "profile");
  const outDir = path.join(tmp, "out");
  fs.mkdirSync(path.join(profile, "user"), { recursive: true });
  fs.mkdirSync(outDir);
  fs.writeFileSync(path.join(profile, "user", "registrymodifications.xcu"), RECALC_XCU);
  const res = spawnSync(
    soffice,
    [
      "--headless",
      "--norestore",
      `-env:UserInstallation=file://${profile}`,
      "--convert-to",
      "xlsx",
      "--outdir",
      outDir,
      outputPath,
    ],
    { encoding: "utf8", timeout: 180_000 },
  );
  const converted = path.join(outDir, path.basename(outputPath));
  if (res.status !== 0 || !fs.existsSync(converted)) {
    console.warn(`  ! --recalc: soffice conversion failed — grading raw file. ${(res.stderr || "").trim().slice(0, 200)}`);
    fs.rmSync(tmp, { recursive: true, force: true });
    return null;
  }
  const dest = outputPath.replace(/\.xlsx$/, "") + ".recalc.xlsx";
  fs.copyFileSync(converted, dest);
  fs.rmSync(tmp, { recursive: true, force: true });
  return dest;
}

/** Assertion types that read computed VALUES (and so should see the recalced
 *  workbook). Everything else reads structure and stays on the raw output. */
const VALUE_ASSERTIONS = new Set(["label_value", "check_rows_zero", "no_error_cells"]);

/** Extract every row across sheets: {sheet, row, label, cells:[{value, formula, evaluated}]} */
function extractRows(wb, sheetFilter) {
  const rows = [];
  for (const ws of wb.worksheets) {
    if (sheetFilter && ws.name !== sheetFilter) continue;
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      const cells = [];
      let label = null;
      row.eachCell({ includeEmpty: false }, (cell, cn) => {
        const v = cell.value;
        let value = null;
        let formula = null;
        if (v && typeof v === "object" && ("formula" in v || "sharedFormula" in v)) {
          formula = v.formula ?? cell.formula ?? "(shared)";
          value = v.result ?? null;
          if (value && typeof value === "object" && "error" in value) value = value.error;
        } else if (v && typeof v === "object" && "richText" in v) {
          value = v.richText.map((t) => t.text).join("");
        } else if (v instanceof Date) {
          value = v;
        } else {
          value = v;
        }
        if (label === null && cn <= 6 && typeof value === "string" && value.trim().length >= 2) {
          label = value.trim();
        }
        cells.push({ col: cn, value, formula });
      });
      rows.push({ sheet: ws.name, row: rn, label, cells });
    });
  }
  return rows;
}

const ERR_RE = /^#(REF!|DIV\/0!|VALUE!|N\/A|NAME\?|NULL!|NUM!)$/;

function numbersInRow(r) {
  return r.cells
    .map((c) => ({ ...c, num: typeof c.value === "number" ? c.value : null }))
    .filter((c) => c.num !== null);
}

function matchRows(rows, labelRegex) {
  const re = new RegExp(labelRegex, "i");
  return rows.filter((r) => r.label && re.test(r.label));
}

function gradeAssertion(a, rows, ctx) {
  const scoped = a.sheet ? rows.filter((r) => r.sheet === a.sheet) : rows;
  switch (a.type) {
    case "label_value": {
      const hits = matchRows(scoped, a.label);
      if (hits.length === 0) return { status: "fail", detail: `no row labeled /${a.label}/i` };
      const tol = (a.tolerance_pct ?? 1) / 100;
      const targets = [a.value, a.value * 1000, a.value / 1000]; // $mm vs $bn etc.
      let sawUnevaluated = false;
      for (const r of hits) {
        for (const c of numbersInRow(r)) {
          for (const t of targets) {
            const ok = t === 0 ? Math.abs(c.num) <= (a.tolerance_abs ?? 0.5) : Math.abs(c.num - t) <= Math.abs(t) * tol;
            if (ok) return { status: "pass", detail: `${r.sheet}!row ${r.row} "${r.label}" → ${c.num}` };
          }
        }
        if (r.cells.some((c) => c.formula && c.value == null)) sawUnevaluated = true;
      }
      if (sawUnevaluated)
        return { status: "warn", detail: `label found but formula cells have no cached values — recalc the file to grade` };
      return { status: "fail", detail: `label found, no cell within ${a.tolerance_pct ?? 1}% of ${a.value}` };
    }
    case "label_formula": {
      const hits = matchRows(scoped, a.label);
      if (hits.length === 0) return { status: "fail", detail: `no row labeled /${a.label}/i` };
      const withFormula = hits.find((r) => r.cells.some((c) => c.formula));
      return withFormula
        ? { status: "pass", detail: `${withFormula.sheet}!row ${withFormula.row} has formulas` }
        : { status: "fail", detail: "label found but every cell is hardcoded" };
    }
    case "label_absent": {
      // Inverse of label_exists — the roll-forward case: a stale marker
      // (e.g. a "Q2'26E" estimate header) MUST be gone after the update.
      // Catches the agent that appends new data without retiring the old
      // labels, leaving a workbook that lies about what's actual.
      const re = new RegExp(a.label, "i");
      const hit = scoped.find(
        (r) => (r.label && re.test(r.label)) || r.cells.some((c) => typeof c.value === "string" && re.test(c.value)),
      );
      return hit
        ? { status: "fail", detail: `stale /${a.label}/i still present at ${hit.sheet}!row ${hit.row}` }
        : { status: "pass", detail: `no /${a.label}/i anywhere` };
    }
    case "label_exists": {
      const re = new RegExp(a.label, "i");
      const hit = scoped.find(
        (r) => (r.label && re.test(r.label)) || r.cells.some((c) => typeof c.value === "string" && re.test(c.value)),
      );
      return hit
        ? { status: "pass", detail: `${hit.sheet}!row ${hit.row}` }
        : { status: "fail", detail: `nothing matches /${a.label}/i` };
    }
    case "check_rows_zero": {
      const hits = matchRows(scoped, a.label ?? "check");
      if (hits.length === 0) return { status: "warn", detail: "no check rows in the model (rule 21 wants them)" };
      const tol = a.tolerance_abs ?? 1;
      for (const r of hits) {
        for (const c of numbersInRow(r)) {
          if (Math.abs(c.num) > tol) {
            return { status: "fail", detail: `${r.sheet}!row ${r.row} "${r.label}" has |${c.num}| > ${tol}` };
          }
        }
      }
      return { status: "pass", detail: `${hits.length} check row(s) tie to ~0` };
    }
    case "min_formula_cells": {
      const n = scoped.reduce((acc, r) => acc + r.cells.filter((c) => c.formula).length, 0);
      return n >= a.count
        ? { status: "pass", detail: `${n} formula cells` }
        : { status: "fail", detail: `${n} formula cells < ${a.count}` };
    }
    case "no_error_cells": {
      for (const r of scoped) {
        for (const c of r.cells) {
          if (typeof c.value === "string" && ERR_RE.test(c.value)) {
            return { status: "fail", detail: `${r.sheet}!row ${r.row} contains ${c.value}` };
          }
        }
      }
      return { status: "pass", detail: "no error values" };
    }
    case "file_opens": {
      return { status: "pass", detail: `parsed: ${ctx.wb.worksheets.length} sheets` };
    }
    case "part_contains": {
      // In-sheet features (conditional formatting, data validation, autoFilter)
      // aren't separate zip parts, so parts_preserved can't see them — assert
      // on the sheet XML content instead.
      if (!ctx.outputZip) return { status: "warn", detail: "no output zip" };
      if (!ctx.outputZip.files[a.part]) return { status: "fail", detail: `part missing: ${a.part}` };
      const text = ctx.partText?.get(a.part) ?? "";
      const re = new RegExp(a.needle, "i");
      return re.test(text)
        ? { status: "pass", detail: `${a.part} contains /${a.needle}/i` }
        : { status: "fail", detail: `${a.part} lacks /${a.needle}/i` };
    }
    case "parts_preserved": {
      if (!ctx.originalZip || !ctx.outputZip) return { status: "warn", detail: "needs --original <start.xlsx>" };
      const allowRemoved = new Set(a.allow_removed ?? []);
      const origNames = Object.keys(ctx.originalZip.files).filter((n) => !ctx.originalZip.files[n].dir);
      const outNames = new Set(Object.keys(ctx.outputZip.files).filter((n) => !ctx.outputZip.files[n].dir));
      const missing = origNames.filter((n) => !outNames.has(n) && !allowRemoved.has(n));
      const mustKeepMissing = (a.must_keep ?? []).filter((n) => !outNames.has(n));
      if (mustKeepMissing.length) {
        return { status: "fail", detail: `dropped critical parts: ${mustKeepMissing.join(", ")}` };
      }
      if (missing.length) {
        return {
          status: "fail",
          detail: `dropped ${missing.length} part(s): ${missing.slice(0, 8).join(", ")}${missing.length > 8 ? "…" : ""}`,
        };
      }
      return { status: "pass", detail: `all ${origNames.length} original parts survive` };
    }
    default:
      return { status: "warn", detail: `unknown assertion type ${a.type}` };
  }
}

async function main() {
  const [specArg, outputArg, ...rest] = process.argv.slice(2);
  if (!specArg || !outputArg) {
    console.error("usage: node eval/grade.mjs <task-id|task.json> <output.xlsx> [--original <start.xlsx>] [--recalc]");
    process.exit(2);
  }
  const origIdx = rest.indexOf("--original");
  const originalArg = origIdx >= 0 ? rest[origIdx + 1] : null;
  const doRecalc = rest.includes("--recalc");
  const { task } = loadTask(specArg);

  const outBytes = fs.readFileSync(outputArg);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(outBytes);
  const rows = extractRows(wb, null);

  // Value assertions read the LibreOffice-recalced copy when --recalc is on
  // (and it succeeded); structure assertions always read the raw output.
  let valueRows = rows;
  let recalcPath = null;
  if (doRecalc) {
    recalcPath = recalcFile(outputArg);
    if (recalcPath) {
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(fs.readFileSync(recalcPath));
      valueRows = extractRows(wb2, null);
    }
  }

  const ctx = {
    wb,
    outputZip: await JSZip.loadAsync(outBytes),
    originalZip: originalArg ? await JSZip.loadAsync(fs.readFileSync(originalArg)) : null,
  };

  // Pre-extract part text for part_contains assertions — JSZip reads are
  // async and gradeAssertion is deliberately sync.
  ctx.partText = new Map();
  for (const a of task.assertions) {
    if (a.type === "part_contains" && ctx.outputZip.files[a.part]) {
      ctx.partText.set(a.part, await ctx.outputZip.files[a.part].async("string"));
    }
  }

  const results = task.assertions.map((a) => {
    let r;
    try {
      r = gradeAssertion(a, VALUE_ASSERTIONS.has(a.type) ? valueRows : rows, ctx);
    } catch (e) {
      r = { status: "fail", detail: `grader error: ${e.message}` };
    }
    return { ...r, type: a.type, label: a.label ?? "", why: a.why ?? "" };
  });

  const counts = { pass: 0, fail: 0, warn: 0 };
  for (const r of results) counts[r.status]++;
  const icon = { pass: "✓", fail: "✗", warn: "…" };
  console.log(`\n${task.id} — ${outputArg}${recalcPath ? " (values recalced via LibreOffice)" : ""}`);
  for (const r of results) {
    console.log(`  ${icon[r.status]} [${r.type}${r.label ? ` /${r.label}/` : ""}] ${r.detail}`);
  }
  console.log(`\n  score: ${counts.pass}/${results.length} pass, ${counts.fail} fail, ${counts.warn} warn\n`);

  const reportPath = outputArg.replace(/\.xlsx$/, "") + ".grade.json";
  fs.writeFileSync(
    reportPath,
    JSON.stringify({ task: task.id, output: outputArg, recalc: Boolean(recalcPath), counts, results }, null, 2),
  );
  process.exit(counts.fail > 0 ? 1 : 0);
}

main().catch((e) => {
  console.error("grader crashed:", e);
  process.exit(2);
});
