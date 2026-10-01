/**
 * `_xlfn.` prefix compatibility.
 *
 * The xlsx format stores every function Excel added after 2007 with an
 * `_xlfn.` prefix in the formula text (`_xlfn.IFS(...)`), and the
 * dynamic-array pair FILTER/SORT with `_xlfn._xlws.`. Excel's UI hides the
 * prefixes and evaluates the functions natively; any other engine (Univer
 * included) sees the prefixed token as an unknown defined name and the cell
 * renders #NAME? / "Invalid Name" even though the file is perfectly healthy.
 *
 * Strip the prefixes on import so Univer evaluates what it can, and restore
 * them on both save paths (surgical patch + ExcelJS full export) so written
 * formulas stay schema-canonical for Excel.
 *
 * `_xlpm.` (LAMBDA parameter prefix) is deliberately left alone: stripping it
 * without re-deriving parameter scopes on save would corrupt LAMBDA formulas,
 * and Univer doesn't evaluate them either way.
 *
 * Pure module — no Univer/ExcelJS imports — so surgicalPatch.ts can use it
 * without breaking its dependency-free contract.
 */

/**
 * Function names Excel stores `_xlfn.`-prefixed. Vendored from Univer's
 * NEW_EXCEL_FUNCTIONS (@univerjs/engine-formula 0.23) — which module can't
 * be imported here without dragging the whole engine into pure code paths —
 * plus post-2007 names that list omits (CONCAT, LET, LAMBDA...). The set is
 * defined by Excel's file format, so it only grows when Excel ships new
 * functions; a missing entry just means that one function keeps today's
 * behavior (#NAME? in the grid, prefix untouched in the file).
 */
const XLFN_FUNCTIONS = new Set([
  "ACOT", "ACOTH", "ARABIC", "BASE", "CEILING.MATH", "CEILING.PRECISE",
  "COMBINA", "COT", "COTH", "CSC", "CSCH", "DECIMAL", "FLOOR.MATH",
  "FLOOR.PRECISE", "MUNIT", "RANDARRAY", "SEC", "SECH", "SEQUENCE",
  "CHOOSECOLS", "CHOOSEROWS", "DROP", "EXPAND", "FILTER", "FORMULATEXT",
  "HSTACK", "SORT", "SORTBY", "TAKE", "TOCOL", "TOROW", "UNIQUE", "VSTACK",
  "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "BITAND", "BITLSHIFT",
  "BITOR", "BITRSHIFT", "BITXOR", "ERF.PRECISE", "ERFC.PRECISE", "IMCOSH",
  "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN",
  "ISFORMULA", "SHEET", "SHEETS", "IFNA", "IFS", "SWITCH", "XOR",
  "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV",
  "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST",
  "CONFIDENCE.NORM", "CONFIDENCE.T", "COVARIANCE.P", "COVARIANCE.S",
  "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
  "FORECAST.LINEAR", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
  "GAUSS", "HYPGEOM.DIST", "LOGNORM.DIST", "LOGNORM.INV", "MAXIFS",
  "MINIFS", "MODE.MULT", "MODE.SNGL", "NEGBINOM.DIST", "NORM.DIST",
  "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "PERCENTILE.EXC",
  "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUTATIONA",
  "PHI", "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC", "RANK.AVG",
  "RANK.EQ", "SKEW.P", "STDEV.P", "STDEV.S", "T.DIST", "T.DIST.2T",
  "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "VAR.P", "VAR.S",
  "WEIBULL.DIST", "Z.TEST", "ARRAYTOTEXT", "ENCODEURL", "NUMBERVALUE",
  "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "UNICHAR", "UNICODE",
  "VALUETOTEXT", "DAYS", "ISOWEEKNUM", "PDURATION", "RRI", "BYCOL", "BYROW",
  "MAKEARRAY", "MAP", "REDUCE", "SCAN",
  // Not in Univer's list but _xlfn.-prefixed by Excel:
  "CONCAT", "LET", "LAMBDA", "ISOMITTED", "STOCKHISTORY", "IMAGE",
]);

/** Functions Excel stores with the longer `_xlfn._xlws.` prefix. */
const XLWS_FUNCTIONS = new Set(["FILTER", "SORT"]);

/** Split so odd segments are double-quoted string literals ("" = escaped quote). */
const STRING_LITERAL_SPLIT = /("(?:""|[^"])*")/;

export function stripXlfnPrefixes(formula: string): string {
  if (!formula.includes("_xl")) return formula;
  return formula
    .split(STRING_LITERAL_SPLIT)
    .map((seg, i) => (i % 2 ? seg : seg.replace(/_xlfn\.(?:_xlws\.)?/gi, "")))
    .join("");
}

/**
 * Re-add `_xlfn.` / `_xlfn._xlws.` before any known post-2007 function call.
 * Already-prefixed calls are untouched (the char before the name is `.`),
 * as are defined names that merely start with a function name (no `(`) and
 * anything inside string literals.
 */
export function addXlfnPrefixes(formula: string): string {
  return formula
    .split(STRING_LITERAL_SPLIT)
    .map((seg, i) => {
      if (i % 2) return seg;
      // Lookahead for the '(' (rather than consuming it) so a nested call
      // like SORT(FILTER(...)) still matches the inner function.
      return seg.replace(
        /(^|[^A-Za-z0-9_.$'])([A-Za-z][A-Za-z0-9._]*)(?=\()/g,
        (match, pre: string, name: string) => {
          const upper = name.toUpperCase();
          if (!XLFN_FUNCTIONS.has(upper)) return match;
          const prefix = XLWS_FUNCTIONS.has(upper) ? "_xlfn._xlws." : "_xlfn.";
          return `${pre}${prefix}${name}`;
        },
      );
    })
    .join("");
}
