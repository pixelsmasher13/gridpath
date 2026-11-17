/** Links to other workbooks as formula text: finding them and respelling them. */
import { describe, it, expect } from "vitest";
import { findLinks, isLinkSheetName, linkSheetName, linkText, quotedPrefix, rewriteLinks, splitLinkSheetName } from "./links";

const found = (f: string) => findLinks(f).map((l) => [l.dir, l.book, l.sheet]);

describe("links in formula text", () => {
  it("finds the bare form", () => {
    expect(found("=[NOW.xlsx]Model!B23+1")).toEqual([["", "NOW.xlsx", "Model"]]);
    expect(found("=SUM([NOW.xlsx]Model!B2:B9)/[1]Sheet_2!$A$1")).toEqual([
      ["", "NOW.xlsx", "Model"],
      ["", "1", "Sheet_2"],
    ]);
    // A file name with a space is still one name between the brackets.
    expect(found("=[NOW 2026.xlsx]Model!B2")).toEqual([["", "NOW 2026.xlsx", "Model"]]);
  });

  it("finds the quoted form, with spaces, a folder and doubled apostrophes", () => {
    expect(found("='[NOW 2026.xlsx]Income Statement'!B23")).toEqual([["", "NOW 2026.xlsx", "Income Statement"]]);
    expect(found("='/Users/me/Models/[NOW.xlsx]Model'!B2")).toEqual([["/Users/me/Models/", "NOW.xlsx", "Model"]]);
    expect(found("='C:\\Models\\[NOW.xlsx]Model'!B2")).toEqual([["C:\\Models\\", "NOW.xlsx", "Model"]]);
    expect(found("='[Bob''s.xlsx]P&L'!A1*2")).toEqual([["", "Bob's.xlsx", "P&L"]]);
  });

  it("reports where each prefix starts and ends", () => {
    const f = "=1+'[A.xlsx]S 1'!B2";
    const [l] = findLinks(f);
    expect(f.slice(l.start, l.end)).toBe("'[A.xlsx]S 1'!");
    expect(linkText(l)).toBe("[A.xlsx]S 1");
  });

  it("leaves text, ordinary sheets and table references alone", () => {
    expect(found('="[NOW.xlsx]Model!A1"&A1')).toEqual([]);
    expect(found('=IF(A1="say ""[x]y!"" twice",1,2)')).toEqual([]);
    expect(found("='My Sheet'!A1+Other!B2")).toEqual([]);
    expect(found("=SUM(Sales[Amount])")).toEqual([]);
    expect(found("=[@Amount]*2")).toEqual([]);
    expect(found("=Sales[[#Headers],[Amount]]")).toEqual([]);
    expect(found("=SUM(Sales[[#This Row],[Q1]:[Q4]])")).toEqual([]);
    expect(found("=[Amount]")).toEqual([]);
    // Unfinished quotes do not hang or invent a link.
    expect(found("='[A.xlsx]S")).toEqual([]);
  });

  it("rewrites every link and nothing else", () => {
    const f = '=[now.xlsx]model!B2+\'[NOW.xlsx]Income Statement\'!C3&"[x]y!"+Sheet2!A1';
    const out = rewriteLinks(f, (l) => quotedPrefix(linkSheetName("NOW.xlsx", l.sheet === "model" ? "Model" : l.sheet)));
    expect(out).toBe("='[NOW.xlsx]Model'!B2+'[NOW.xlsx]Income Statement'!C3&\"[x]y!\"+Sheet2!A1");
    // null keeps a link as it was written.
    expect(rewriteLinks(f, () => null)).toBe(f);
  });

  it("names the hidden copy like the link", () => {
    expect(linkSheetName("NOW.xlsx", "Model")).toBe("[NOW.xlsx]Model");
    expect(isLinkSheetName("[NOW.xlsx]Model")).toBe(true);
    expect(isLinkSheetName("Model")).toBe(false);
    expect(splitLinkSheetName("[NOW.xlsx]Income Statement")).toEqual({ book: "NOW.xlsx", sheet: "Income Statement" });
    expect(splitLinkSheetName("Model")).toBeNull();
    expect(splitLinkSheetName("[NOW.xlsx]")).toBeNull();
    expect(quotedPrefix("[Bob's.xlsx]P&L")).toBe("'[Bob''s.xlsx]P&L'!");
    // What the quoted prefix says is read back as the same link.
    expect(found(`=${quotedPrefix("[Bob's.xlsx]P&L")}A1`)).toEqual([["", "Bob's.xlsx", "P&L"]]);
  });
});
