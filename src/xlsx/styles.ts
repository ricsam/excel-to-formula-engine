/**
 * Parses xl/styles.xml into the subset of formatting the engine can express.
 *
 * Excel's style model is indirect: a cell carries an index into `cellXfs`, and
 * each `xf` points at a font, fill, border and number format. This flattens
 * that into one record per `xf`.
 */

import type { CellStyle } from "@ricsam/formula-engine";
import { asArray, attr, boolAttr, numAttr, type XlsxPackage, type XmlNode } from "./package";

export interface ExcelFormat {
  style: CellStyle;
  /** Number format code, e.g. `"0.00%"` or `"yyyy-mm-dd"`. */
  numberFormat?: string;
  /** True when the number format renders a date or a time. */
  isDateFormat: boolean;
}

export interface StyleTable {
  formats: ExcelFormat[];
  /** Differential formats, referenced by conditional formatting rules. */
  differentialFormats: CellStyle[];
}

/**
 * Number format ids Excel defines implicitly. Only the date/time ones matter
 * here: everything else is carried through as a metadata string.
 */
const BUILTIN_FORMATS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

const BUILTIN_DATE_FORMAT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
]);

export function parseStyles(pkg: XlsxPackage): StyleTable {
  const root = pkg.xml("xl/styles.xml");
  const styleSheet = root?.["styleSheet"] as XmlNode | undefined;
  if (!styleSheet) {
    return { formats: [], differentialFormats: [] };
  }

  const theme = parseThemeColors(pkg);
  const indexedColors = parseIndexedColors(styleSheet);

  const numberFormats = new Map<number, string>();
  for (const numFmt of asArray(
    (styleSheet["numFmts"] as XmlNode | undefined)?.["numFmt"]
  )) {
    const id = numAttr(numFmt, "numFmtId");
    const code = attr(numFmt, "formatCode");
    if (id !== undefined && code !== undefined) {
      numberFormats.set(id, code);
    }
  }

  const fonts = asArray((styleSheet["fonts"] as XmlNode | undefined)?.["font"]);
  const fills = asArray((styleSheet["fills"] as XmlNode | undefined)?.["fill"]);
  const borders = asArray(
    (styleSheet["borders"] as XmlNode | undefined)?.["border"]
  );

  const cellXfs = asArray(
    (styleSheet["cellXfs"] as XmlNode | undefined)?.["xf"]
  );

  const formats: ExcelFormat[] = cellXfs.map((rawXf) => {
    const xf = rawXf as XmlNode;
    const style: CellStyle = {};

    if (boolAttr(xf, "applyFont", true)) {
      const font = fonts[numAttr(xf, "fontId") ?? -1];
      if (font) {
        applyFont(style, font as XmlNode, theme, indexedColors);
      }
    }

    if (boolAttr(xf, "applyFill", true)) {
      const fill = fills[numAttr(xf, "fillId") ?? -1];
      if (fill) {
        applyFill(style, fill as XmlNode, theme, indexedColors);
      }
    }

    if (boolAttr(xf, "applyBorder", true)) {
      const border = borders[numAttr(xf, "borderId") ?? -1];
      if (border) {
        applyBorder(style, border as XmlNode, theme, indexedColors);
      }
    }

    if (boolAttr(xf, "applyAlignment", true)) {
      const alignment = xf["alignment"];
      if (alignment && boolAttr(alignment, "wrapText")) {
        style.wrapText = true;
      }
    }

    const numFmtId = numAttr(xf, "numFmtId") ?? 0;
    const numberFormat = numberFormats.get(numFmtId) ?? BUILTIN_FORMATS[numFmtId];

    return {
      style,
      numberFormat: numberFormat === "General" ? undefined : numberFormat,
      isDateFormat: isDateFormat(numFmtId, numberFormat),
    };
  });

  const differentialFormats = asArray(
    (styleSheet["dxfs"] as XmlNode | undefined)?.["dxf"]
  ).map((dxf) => {
    const style: CellStyle = {};
    const node = dxf as XmlNode;
    if (node["font"]) {
      applyFont(style, node["font"] as XmlNode, theme, indexedColors);
    }
    if (node["fill"]) {
      applyFill(style, node["fill"] as XmlNode, theme, indexedColors);
    }
    if (node["border"]) {
      applyBorder(style, node["border"] as XmlNode, theme, indexedColors);
    }
    return style;
  });

  return { formats, differentialFormats };
}

function applyFont(
  style: CellStyle,
  font: XmlNode,
  theme: string[],
  indexedColors: string[]
): void {
  if (font["b"] !== undefined && isOn(font["b"])) {
    style.bold = true;
  }
  if (font["i"] !== undefined && isOn(font["i"])) {
    style.italic = true;
  }
  if (font["u"] !== undefined && isOn(font["u"])) {
    style.underline = true;
  }
  const size = numAttr(font["sz"], "val");
  if (size !== undefined) {
    // Excel stores font size in points; the engine's fontSize is in pixels.
    style.fontSize = Math.round(size * (96 / 72));
  }
  const color = resolveColor(font["color"], theme, indexedColors);
  if (color) {
    style.color = color;
  }
}

function applyFill(
  style: CellStyle,
  fill: XmlNode,
  theme: string[],
  indexedColors: string[]
): void {
  const pattern = fill["patternFill"] as XmlNode | undefined;
  if (!pattern) {
    return;
  }
  const patternType = attr(pattern, "patternType");
  if (patternType === "none" || patternType === undefined) {
    return;
  }
  // For a solid fill Excel puts the visible colour in fgColor, not bgColor.
  const color =
    resolveColor(pattern["fgColor"], theme, indexedColors) ??
    resolveColor(pattern["bgColor"], theme, indexedColors);
  if (color) {
    style.backgroundColor = color;
  }
}

const BORDER_SIDES = ["top", "right", "bottom", "left"] as const;

function applyBorder(
  style: CellStyle,
  border: XmlNode,
  theme: string[],
  indexedColors: string[]
): void {
  const sides: NonNullable<CellStyle["borderSides"]> = {};
  let color: string | undefined;
  let any = false;

  for (const side of BORDER_SIDES) {
    const node = border[side] as XmlNode | undefined;
    const borderStyle = node ? attr(node, "style") : undefined;
    if (!node || borderStyle === undefined || borderStyle === "none") {
      continue;
    }
    sides[side] = true;
    any = true;
    color ??= resolveColor(node["color"], theme, indexedColors);
  }

  if (any) {
    style.borderSides = sides;
    // The engine has one border colour per cell, so the first side's colour
    // stands in for all of them.
    style.borderColor = color ?? "#000000";
  }
}

function isOn(node: unknown): boolean {
  // <b/> means bold; <b val="0"/> means not bold.
  const val = attr(node, "val");
  return val === undefined || val === "1" || val.toLowerCase() === "true";
}

function isDateFormat(numFmtId: number, code: string | undefined): boolean {
  if (BUILTIN_DATE_FORMAT_IDS.has(numFmtId)) {
    return true;
  }
  if (!code) {
    return false;
  }
  // Strip literal text and colour/condition sections before looking for date
  // tokens, so that "0.00" with a [Red] section is not read as a date.
  const stripped = code
    .replace(/\[[^\]]*\]/g, "")
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "");
  return /[dmyhs]/i.test(stripped) && !/^[^dmyhs]*$/i.test(stripped);
}

/** Convert an Excel colour node to a `#rrggbb` string. */
function resolveColor(
  node: unknown,
  theme: string[],
  indexedColors: string[]
): string | undefined {
  if (!node || typeof node !== "object") {
    return undefined;
  }

  const rgb = attr(node, "rgb");
  if (rgb) {
    return applyTint(argbToHex(rgb), numAttr(node, "tint"));
  }

  const themeIndex = numAttr(node, "theme");
  if (themeIndex !== undefined) {
    const base = theme[themeIndex];
    if (base) {
      return applyTint(base, numAttr(node, "tint"));
    }
  }

  const indexed = numAttr(node, "indexed");
  if (indexed !== undefined) {
    const base = indexedColors[indexed];
    if (base) {
      return applyTint(base, numAttr(node, "tint"));
    }
  }

  return undefined;
}

/** Excel writes colours as AARRGGBB; the alpha channel is dropped. */
function argbToHex(argb: string): string {
  const value = argb.replace(/^#/, "");
  const rgb = value.length === 8 ? value.slice(2) : value;
  return `#${rgb.toLowerCase()}`;
}

/**
 * Excel lightens or darkens a base colour by a tint in [-1, 1]. Negative tints
 * darken, positive tints lighten toward white.
 */
function applyTint(hex: string, tint: number | undefined): string {
  if (!tint) {
    return hex;
  }
  const value = hex.replace("#", "");
  const channels = [0, 2, 4].map((offset) =>
    Number.parseInt(value.slice(offset, offset + 2), 16)
  );
  const tinted = channels.map((channel) => {
    const next =
      tint < 0 ? channel * (1 + tint) : channel * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(next)));
  });
  return `#${tinted.map((c) => c.toString(16).padStart(2, "0")).join("")}`;
}

/** Theme colours, in the order conditional formats and styles index them. */
function parseThemeColors(pkg: XlsxPackage): string[] {
  const themePath = pkg
    .list("xl/theme/")
    .find((path) => path.endsWith(".xml"));
  if (!themePath) {
    return [];
  }
  const text = pkg.text(themePath);
  if (!text) {
    return [];
  }

  const scheme = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(text)?.[0];
  if (!scheme) {
    return [];
  }

  const colors: string[] = [];
  const entry =
    /<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>([\s\S]*?)<\/a:\1>/g;
  let match: RegExpExecArray | null;
  while ((match = entry.exec(scheme)) !== null) {
    const body = match[2] ?? "";
    const srgb = /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(body)?.[1];
    const sys = /<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/.exec(body)?.[1];
    colors.push(`#${(srgb ?? sys ?? "000000").toLowerCase()}`);
  }

  // Excel's theme index order swaps the first two pairs relative to the file.
  if (colors.length >= 4) {
    const [dk1, lt1, dk2, lt2, ...rest] = colors;
    return [lt1!, dk1!, lt2!, dk2!, ...rest];
  }
  return colors;
}

function parseIndexedColors(styleSheet: XmlNode): string[] {
  const colors = asArray(
    (
      (styleSheet["colors"] as XmlNode | undefined)?.["indexedColors"] as
        | XmlNode
        | undefined
    )?.["rgbColor"]
  );
  if (colors.length > 0) {
    return colors.map((color) => argbToHex(attr(color, "rgb") ?? "FF000000"));
  }
  return DEFAULT_INDEXED_COLORS;
}

/** Excel's legacy 56-colour palette, used when a file omits its own. */
const DEFAULT_INDEXED_COLORS: string[] = [
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff",
  "#00ffff", "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00",
  "#ff00ff", "#00ffff", "#800000", "#008000", "#000080", "#808000", "#800080",
  "#008080", "#c0c0c0", "#808080", "#9999ff", "#993366", "#ffffcc", "#ccffff",
  "#660066", "#ff8080", "#0066cc", "#ccccff", "#000080", "#ff00ff", "#ffff00",
  "#00ffff", "#800080", "#800000", "#008080", "#0000ff", "#00ccff", "#ccffff",
  "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99", "#3366ff",
  "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696",
  "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399",
  "#333333",
];
