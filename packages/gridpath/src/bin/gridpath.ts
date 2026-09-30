#!/usr/bin/env node
/**
 * gridpath — let your agent edit Excel files without butchering them.
 *
 *   gridpath mcp [--allow <dir>]... [--review-required] [--port <n>]
 *       start the MCP server on stdio (+ the local review server)
 *   gridpath review <file.xlsx> [--port <n>]
 *       open the review tab for a workbook's pending batches
 */
import { spawn } from "node:child_process";
import path from "node:path";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createServer } from "../mcp/server";
import { startReviewServer } from "../review/api";
import { ToolHandlers } from "../tools/handlers";

const [, , command = "help", ...rest] = process.argv;

function flag(args: string[], name: string): boolean {
  return args.includes(name);
}
function values(args: string[], name: string): string[] {
  const out: string[] = [];
  for (let i = 0; i < args.length; i++) if (args[i] === name && args[i + 1]) out.push(args[++i]);
  return out;
}
function positional(args: string[]): string[] {
  const out: string[] = [];
  for (let i = 0; i < args.length; i++) {
    if (args[i].startsWith("--")) {
      if (args[i] !== "--review-required") i++;
      continue;
    }
    out.push(args[i]);
  }
  return out;
}

function openBrowser(url: string): void {
  const cmd = process.platform === "darwin" ? "open" : process.platform === "win32" ? "start" : "xdg-open";
  try {
    spawn(cmd, [url], { detached: true, stdio: "ignore" }).unref();
  } catch {
    /* printing the URL is enough */
  }
}

if (command === "mcp") {
  const allowRoots = values(rest, "--allow");
  const port = Number(values(rest, "--port")[0] ?? 0);
  let reviewBase: { url: string; token: string } | null = null;
  const { server, handlers } = createServer({ allowRoots, reviewRequired: flag(rest, "--review-required"), reviewUrl: () => reviewBase });
  const review = await startReviewServer(handlers, { port });
  reviewBase = { url: review.url, token: review.token };
  await server.connect(new StdioServerTransport());
  // stdout is the protocol channel; say hello on stderr only.
  process.stderr.write(`gridpath mcp ready — roots: ${allowRoots.join(", ") || process.cwd()} — review server: ${review.url}${handlers.reviewRequired ? " (review required to save)" : ""}\n`);
} else if (command === "review") {
  const file = positional(rest)[0];
  if (!file) {
    process.stderr.write("usage: gridpath review <file.xlsx>\n");
    process.exit(2);
  }
  const abs = path.resolve(file);
  const handlers = new ToolHandlers({ allowRoots: [path.dirname(abs)], reviewUrl: () => base });
  let base: { url: string; token: string } | null = null;
  const review = await startReviewServer(handlers, { port: Number(values(rest, "--port")[0] ?? 0) });
  base = { url: review.url, token: review.token };
  const s = handlers.session(abs);
  const url = handlers.reviewUrlFor(abs)!;
  process.stdout.write(`${s.batches.length} pending batch(es) for ${abs}\n${url}\n`);
  openBrowser(url);
  // Keep serving until the tab is closed and the user hits Ctrl-C.
} else {
  process.stdout.write(`gridpath — let your agent edit Excel files without butchering them.

usage:
  gridpath mcp [--allow <dir>]... [--review-required] [--port <n>]
      start the MCP server (stdio). Files outside the allowed dirs are refused.
      --review-required: only the review tab can save; agents get the link.
  gridpath review <file.xlsx>
      open the review tab for a workbook's pending (unsaved) batches.
`);
}
