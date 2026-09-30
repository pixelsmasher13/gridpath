/**
 * xlsx color resolution: theme + tint and legacy indexed colors → CSS hex.
 *
 * Excel stores most colors as THEME references ({theme: 4, tint: 0.4}) into
 * xl/theme/theme1.xml, not literal argb — measured on a real bank model, 98%
 * of fills are theme colors. ExcelJS passes these through unresolved, so
 * anything that only reads `.argb` drops nearly every color in the file.
 */

/** ExcelJS color object: { argb } | { theme, tint? } | { indexed } */
export type XlsxColor = {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
} | null | undefined;

/**
 * Extract the 12 scheme colors from theme1.xml, ordered by the THEME INDEX
 * used in styles ("theme" attribute). Note the spreadsheet quirk: indices
 * 0/1 and 2/3 are lt1/dk1 and lt2/dk2 — swapped relative to the XML order.
 */
export function parseThemePalette(themeXml: string | null | undefined): string[] | null {
  if (!themeXml) return null;
  const scheme = /<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(themeXml)?.[1];
  if (!scheme) return null;
  const byName: Record<string, string> = {};
  for (const name of [
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
  ]) {
    const el = new RegExp(`<a:${name}>([\\s\\S]*?)</a:${name}>`).exec(scheme)?.[1];
    if (!el) continue;
    const v =
      /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(el)?.[1] ??
      /<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/.exec(el)?.[1];
    if (v) byName[name] = v.toUpperCase();
  }
  const indexOrder = [
    "lt1", "dk1", "lt2", "dk2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
  ];
  const palette = indexOrder.map((n) => byName[n] ?? "");
  return palette.some(Boolean) ? palette : null;
}

/**
 * Apply an Excel tint to a hex color per ECMA-376: scale HSL luminance
 * toward black (tint < 0) or white (tint > 0).
 */
export function applyTint(hex6: string, tint: number): string {
  if (!tint) return hex6;
  const r = parseInt(hex6.slice(0, 2), 16) / 255;
  const g = parseInt(hex6.slice(2, 4), 16) / 255;
  const b = parseInt(hex6.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let l = (max + min) / 2;
  let h = 0;
  let s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  l = Math.min(1, Math.max(0, l));

  const hue = (p: number, q: number, t: number): number => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r2: number;
  let g2: number;
  let b2: number;
  if (s === 0) {
    r2 = g2 = b2 = l;
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r2 = hue(p, q, h + 1 / 3);
    g2 = hue(p, q, h);
    b2 = hue(p, q, h - 1 / 3);
  }
  const to2 = (v: number) => Math.round(v * 255).toString(16).padStart(2, "0").toUpperCase();
  return to2(r2) + to2(g2) + to2(b2);
}

/** Legacy indexed palette (0–63). 64/65 are system colors → unresolved. */
const INDEXED_PALETTE: string[] = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

/** Resolve any ExcelJS color object to "#RRGGBB", or null when it can't be. */
export function resolveXlsxColor(color: XlsxColor, themePalette: string[] | null): string | null {
  if (!color) return null;
  if (color.argb && typeof color.argb === "string") {
    const hex = color.argb.length === 8 ? color.argb.slice(2) : color.argb;
    if (/^[0-9A-Fa-f]{6}$/.test(hex)) return `#${hex.toUpperCase()}`;
    return null;
  }
  if (typeof color.theme === "number") {
    const base = themePalette?.[color.theme];
    if (!base) return null;
    return `#${applyTint(base, color.tint ?? 0)}`;
  }
  if (typeof color.indexed === "number") {
    const hex = INDEXED_PALETTE[color.indexed];
    return hex ? `#${hex}` : null;
  }
  return null;
}
