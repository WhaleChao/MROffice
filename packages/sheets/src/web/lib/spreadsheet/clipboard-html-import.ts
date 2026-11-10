/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded, fail-closed import of the `text/html` clipboard flavour emitted by
 * Excel, WPS and browser spreadsheet grids.
 *
 * The browser only exposes interoperable text/HTML clipboard flavours, not the
 * proprietary native workbook payload. We therefore preserve the table shape,
 * displayed text, explicit formula attributes and a conservative formatting
 * subset. Unsupported CSS is ignored rather than executed or flattened into an
 * unsafe style string.
 */

import { PX_PER_PT, ptToPx } from "@mrnote/core/web/lib/office-units.ts";
import { cssNamedColorHex } from "@mrnote/core/web/lib/css-named-colors.ts";
import type { BorderLineStyle, CellBorderEdges, RichBorderEdge } from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import { sanitizeHyperlink } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { PASTE_MAX_COLS, PASTE_MAX_ROWS } from "@mrnote/core/web/lib/spreadsheet/paste-limits.ts";
import {
  normalizeCellRuns,
  type TextRun,
} from "@mrnote/core/web/lib/spreadsheet/cell-rich-text.ts";
import type { RichClipboard, RichClipboardCell } from "@mrnote/core/web/lib/spreadsheet/paste-special.ts";
import { formulaR1C1ToA1 } from "@mrnote/core/web/lib/spreadsheet/r1c1.ts";

const MAX_HTML_CHARS = 4 * 1024 * 1024;
const MAX_ROWS = PASTE_MAX_ROWS;
const MAX_COLS = PASTE_MAX_COLS;
const MAX_CELL_CHARS = 32_767;
const MAX_FORMULA_CHARS = 8_192;
const MAX_CSS_RULES = 512;
const MAX_RICH_RUNS = 200;

interface CssRule {
  selector: string;
  declarations: string;
}

export interface ParsedSpreadsheetClipboardHtml {
  rows: string[][];
  cells: RichClipboard;
  /** Relative merge spans authored by HTML rowspan/colspan cells. */
  merges: SpreadsheetClipboardMerge[];
  hasRichMetadata: boolean;
  /** Rows or columns beyond the paste bound were left out. */
  truncated: boolean;
}

export interface SpreadsheetClipboardMerge {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

export interface SpreadsheetClipboardHtmlParseOptions {
  /** Zero-based destination anchor used to resolve Excel/WPS R1C1 formulas. */
  baseRow?: number;
  baseCol?: number;
}

const NAMED_COLORS: Record<string, string> = {
  aqua: "#00ffff",
  black: "#000000",
  blue: "#0000ff",
  fuchsia: "#ff00ff",
  gray: "#808080",
  grey: "#808080",
  green: "#008000",
  lime: "#00ff00",
  maroon: "#800000",
  navy: "#000080",
  purple: "#800080",
  red: "#ff0000",
  silver: "#c0c0c0",
  teal: "#008080",
  white: "#ffffff",
  windowtext: "#000000",
  yellow: "#ffff00",
};

function byteHex(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value)))
    .toString(16)
    .padStart(2, "0");
}

function normalizeColor(raw: string | undefined): string | undefined {
  const value = raw?.trim().toLowerCase();
  if (!value || value === "none" || value === "transparent") return undefined;
  const short = /^#([0-9a-f]{3})$/i.exec(value);
  if (short) {
    return `#${short[1]
      .split("")
      .map((digit) => `${digit}${digit}`)
      .join("")}`;
  }
  if (/^#[0-9a-f]{6}$/i.test(value)) return value;
  const rgb =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*(\d*\.?\d+))?\s*\)$/i.exec(
      value,
    );
  if (rgb) {
    const alpha = rgb[4] === undefined ? 1 : Number(rgb[4]);
    if (!Number.isFinite(alpha) || alpha <= 0) return undefined;
    // A cell colour is opaque: a translucent one is laid over the white
    // paper it was shown on (rgba(255,0,0,0.2) is the light red #ffcccc,
    // not pure red - live audit B439).
    const over = (channel: string) => 255 - (255 - Number(channel)) * Math.min(1, alpha);
    return `#${byteHex(over(rgb[1]))}${byteHex(over(rgb[2]))}${byteHex(over(rgb[3]))}`;
  }
  // Every CSS named colour (olive, orange, darkcyan …), not a short list of
  // the basic ones: the same table the rest of the product reads (B489).
  return NAMED_COLORS[value] ?? cssNamedColorHex(value)?.toLowerCase() ?? undefined;
}

/**
 * The declaration text without its CSS comments. A comment inside a string
 * ("Fira /* x *\/ Code") is that string's text and stays (B486: a comment's
 * semicolon split a declaration, and a commented-out bold was applied).
 */
function stripCssComments(input: string): string {
  let out = "";
  let quote = "";
  let escaped = false;
  for (let index = 0; index < input.length; index += 1) {
    const char = input[index];
    if (escaped) {
      out += char;
      escaped = false;
      continue;
    }
    if (char === "\\") {
      out += char;
      escaped = true;
      continue;
    }
    if (quote) {
      if (char === quote) quote = "";
      out += char;
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      out += char;
      continue;
    }
    if (char === "/" && input[index + 1] === "*") {
      const end = input.indexOf("*/", index + 2);
      if (end < 0) break;
      index = end + 1;
      out += " ";
      continue;
    }
    out += char;
  }
  return out;
}

/**
 * Whether a declared value is one the importer can read for its property.
 * An unreadable value is no declaration at all: it never wins the cascade
 * over an earlier readable one (B485: `color: mrai-invalid` after
 * `color: red` turned the red off).
 */
function cssValueIsSupported(property: string, value: string): boolean {
  const v = value.trim().toLowerCase();
  if (!v) return false;
  switch (property) {
    case "color":
    case "background-color":
    case "background":
      return normalizeColor(v) !== undefined;
    case "font-weight":
      return (
        /^(normal|bold|bolder|lighter)$/.test(v) ||
        (/^\d{1,4}$/.test(v) && Number(v) >= 1 && Number(v) <= 1000)
      );
    case "font-style":
      return /^(normal|italic|oblique)$/.test(v);
    case "font-size":
      return (
        /^\d*\.?\d+(px|pt|em|rem|%|in|cm|mm|pc)$/.test(v) ||
        /^(xx-small|x-small|small|medium|large|x-large|xx-large|xxx-large|smaller|larger)$/.test(v)
      );
    case "text-align":
      return /^(left|center|right|justify|start|end)$/.test(v);
    case "vertical-align":
      return /^(top|middle|bottom|baseline|sub|super|text-top|text-bottom)$/.test(v);
    case "text-decoration":
    case "text-decoration-line":
      return /\b(none|underline|line-through|overline)\b/.test(v);
    case "white-space":
      return /^(normal|nowrap|pre|pre-wrap|pre-line|break-spaces)$/.test(v);
    default:
      return true;
  }
}

function splitCssDeclarations(rawInput: string): string[] {
  const input = stripCssComments(rawInput);
  const parts: string[] = [];
  let start = 0;
  let quote = "";
  let escaped = false;
  let parenDepth = 0;
  for (let index = 0; index < input.length; index += 1) {
    const char = input[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (char === "\\") {
      escaped = true;
      continue;
    }
    if (quote) {
      if (char === quote) quote = "";
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === "(") parenDepth += 1;
    else if (char === ")" && parenDepth > 0) parenDepth -= 1;
    else if (char === ";" && parenDepth === 0) {
      parts.push(input.slice(start, index));
      start = index + 1;
    }
  }
  parts.push(input.slice(start));
  return parts;
}

/** One declaration block and the specificity of what it came from. */
type DeclarationSource = string | { readonly text: string; readonly specificity: number };

/**
 * The cascaded value of each property across declaration blocks: an
 * `!important` declaration beats a normal one, then the higher specificity
 * wins, then the later block (live audit B438). `!important` is a priority,
 * not part of the value. A plain string is an inline style.
 */
function declarationsMap(inputs: readonly DeclarationSource[]): Map<string, string> {
  const winners = new Map<string, { value: string; rank: [number, number, number] }>();
  inputs.forEach((input, order) => {
    const text = typeof input === "string" ? input : input.text;
    const specificity = typeof input === "string" ? 1000 : input.specificity;
    for (const declaration of splitCssDeclarations(text)) {
      const colon = declaration.indexOf(":");
      if (colon < 1) continue;
      const property = declaration.slice(0, colon).trim().toLowerCase();
      let value = declaration.slice(colon + 1).trim();
      const important = /!\s*important\s*$/i.test(value);
      if (important) value = value.replace(/\s*!\s*important\s*$/i, "").trim();
      if (!property || !value) continue;
      const rank: [number, number, number] = [important ? 1 : 0, specificity, order];
      // The `font` shorthand is its longhands, in its place in the block:
      // what it omits is reset to the initial value (live audit B478).
      const longhands: ReadonlyArray<readonly [string, string]> =
        property === "font" ? expandFontShorthand(value) ?? [] : [[property, value]];
      for (const [name, longhandValue] of longhands) {
        if (!cssValueIsSupported(name, longhandValue)) continue;
        const current = winners.get(name);
        const beats =
          !current ||
          rank[0] > current.rank[0] ||
          (rank[0] === current.rank[0] &&
            (rank[1] > current.rank[1] || (rank[1] === current.rank[1] && rank[2] >= current.rank[2])));
        if (beats) winners.set(name, { value: longhandValue, rank });
      }
    }
  });
  return new Map(Array.from(winners, ([property, { value }]) => [property, value]));
}

const FONT_STRETCH_KEYWORDS = new Set([
  "ultra-condensed", "extra-condensed", "condensed", "semi-condensed",
  "semi-expanded", "expanded", "extra-expanded", "ultra-expanded",
]);

/**
 * The longhands of one CSS `font` shorthand (MDN "font"): `[style ||
 * variant || weight || stretch]? size [/ line-height]? family`. Every
 * longhand it can set is returned, the omitted ones as `normal`, so a later
 * `font: 20px Arial` takes back an earlier bold or italic. null when the
 * value is not a shorthand this reader can expand (a system font keyword, a
 * CSS-wide keyword, or invalid CSS, which a browser ignores whole).
 */
function expandFontShorthand(value: string): Array<[string, string]> | null {
  let rest = value.trim();
  if (rest.length > 1024) return null;
  let fontStyle: string | undefined;
  let fontVariant: string | undefined;
  let fontWeight: string | undefined;
  let fontStretch: string | undefined;
  for (let index = 0; index < 4; index += 1) {
    const match = /^(\S+)\s+/.exec(rest);
    if (!match) break;
    const token = match[1].toLowerCase();
    if (token === "normal") {
      // `normal` may stand for any of the four.
    } else if (token === "italic" || token === "oblique") {
      if (fontStyle) return null;
      fontStyle = token;
    } else if (token === "small-caps") {
      if (fontVariant) return null;
      fontVariant = token;
    } else if (
      token === "bold" ||
      token === "bolder" ||
      token === "lighter" ||
      (/^\d+(?:\.\d+)?$/.test(token) && Number(token) >= 1 && Number(token) <= 1000)
    ) {
      if (fontWeight) return null;
      fontWeight = token;
    } else if (FONT_STRETCH_KEYWORDS.has(token)) {
      if (fontStretch) return null;
      fontStretch = token;
    } else {
      break;
    }
    rest = rest.slice(match[0].length);
  }
  const sized = /^([^\s/]+)(?:\s*\/\s*([^\s/]+))?\s+(\S[\s\S]*)$/.exec(rest);
  if (!sized || cssFontSizePx(sized[1], DEFAULT_FONT_PX) === undefined) return null;
  return [
    ["font-style", fontStyle ?? "normal"],
    ["font-variant", fontVariant ?? "normal"],
    ["font-weight", fontWeight ?? "normal"],
    ["font-stretch", fontStretch ?? "normal"],
    ["font-size", sized[1]],
    ["line-height", sized[2] ?? "normal"],
    ["font-family", sized[3].trim()],
  ];
}

/** Specificity of the selectors collectCssRules keeps: a tag 1, a class 10. */
function selectorSpecificity(selector: string): number {
  const classes = (selector.match(/\.[a-z_][a-z0-9_-]*/gi) || []).length;
  return classes * 10 + (/^(td|th)/i.test(selector) ? 1 : 0);
}

function collectCssRules(document: Document): CssRule[] {
  const rules: CssRule[] = [];
  for (const style of Array.from(document.querySelectorAll("style"))) {
    const text = (style.textContent || "").slice(0, 512 * 1024);
    const withoutComments = text.replace(/\/\*[\s\S]*?\*\//g, "");
    const matcher = /([^{}]+)\{([^{}]*)\}/g;
    let match: RegExpExecArray | null;
    while (
      (match = matcher.exec(withoutComments)) &&
      rules.length < MAX_CSS_RULES
    ) {
      for (const selector of match[1].split(",")) {
        const normalized = selector.trim();
        if (
          /^(?:td|th)?(?:\.[a-z_][a-z0-9_-]*)+$/i.test(normalized) ||
          normalized.toLowerCase() === "td" ||
          normalized.toLowerCase() === "th"
        ) {
          rules.push({ selector: normalized, declarations: match[2] });
        }
        if (rules.length >= MAX_CSS_RULES) break;
      }
    }
    if (rules.length >= MAX_CSS_RULES) break;
  }
  return rules;
}

function ruleMatchesCell(rule: CssRule, cell: Element): boolean {
  const selector = rule.selector;
  const tag = cell.tagName.toLowerCase();
  const tagMatch = /^(td|th)/i.exec(selector);
  if (tagMatch && tagMatch[1].toLowerCase() !== tag) return false;
  const classes = selector.match(/\.([a-z_][a-z0-9_-]*)/gi) || [];
  return classes.every((name) => cell.classList.contains(name.slice(1)));
}

/** The cascaded declarations of one element: its matching rules and inline style. */
function elementDeclarations(element: Element, rules: CssRule[]): Map<string, string> {
  // HTML's own presentation attributes (align, valign, bgcolor) are the
  // lowest-priority declarations of the cascade, under every stylesheet rule
  // and the inline style (B488: a <td bgcolor align valign> lost all three).
  const presentational: string[] = [];
  const align = element.getAttribute("align");
  if (align) presentational.push(`text-align:${align}`);
  const valign = element.getAttribute("valign");
  if (valign) presentational.push(`vertical-align:${valign}`);
  const bgcolor = element.getAttribute("bgcolor");
  if (bgcolor) presentational.push(`background-color:${bgcolor}`);
  const inputs: DeclarationSource[] = presentational.length
    ? [{ text: presentational.join(";"), specificity: -1 }]
    : [];
  inputs.push(
    ...rules
      .filter((rule) => ruleMatchesCell(rule, element))
      .map((rule) => ({ text: rule.declarations, specificity: selectorSpecificity(rule.selector) })),
  );
  const inline = element.getAttribute("style");
  if (inline) inputs.push(inline);
  return declarationsMap(inputs);
}

const INHERITED_FONT_PROPERTIES = ["color", "font-size", "font-weight", "font-style", "font-family"] as const;

/** The inherited font properties a cell receives from its ancestors up to its table. */
function inheritedFontDeclarations(cell: Element, rules: CssRule[]): Map<string, string> {
  const chain: Element[] = [];
  const table = cell.closest("table");
  for (let element = cell.parentElement; element; element = element.parentElement) {
    chain.unshift(element);
    if (element === table) break;
  }
  const inherited = new Map<string, string>();
  for (const element of chain) {
    const declarations = elementDeclarations(element, rules);
    for (const property of INHERITED_FONT_PROPERTIES) {
      const value = declarations.get(property);
      if (value && value.toLowerCase() !== "inherit") inherited.set(property, value);
    }
  }
  return inherited;
}

/** An element the source page does not show: `hidden`, or display:none. */
function isHiddenElement(element: Element, rules: CssRule[]): boolean {
  if (element.hasAttribute("hidden")) return true;
  return elementDeclarations(element, rules).get("display")?.toLowerCase() === "none";
}

function unquoteCssValue(value: string): string {
  const trimmed = value.trim();
  const body =
    (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
      ? trimmed.slice(1, -1)
      : trimmed;
  return body.replace(/\\(["'@\\])/g, "$1").trim();
}

/** CSS `medium`, the initial font size a page without one starts from. */
const DEFAULT_FONT_PX = 16;

/** The CSS absolute-size keywords at the 16px `medium` (CSS Fonts 4). */
const ABSOLUTE_FONT_SIZE_PX: Readonly<Record<string, number>> = {
  "xx-small": 9,
  "x-small": 10,
  small: 13,
  medium: 16,
  large: 18,
  "x-large": 24,
  "xx-large": 32,
  "xxx-large": 48,
};

const FONT_SIZE_UNIT_PX: Readonly<Record<string, number>> = {
  px: 1,
  pt: PX_PER_PT,
  pc: 16,
  in: 96,
  cm: 96 / 2.54,
  mm: 96 / 25.4,
  q: 96 / 101.6,
  // The root is the page's own `medium`: the clipboard has no <html> size.
  rem: DEFAULT_FONT_PX,
};

/**
 * A `font-size` value as computed CSS pixels against the parent's computed
 * size (MDN "font-size": `em` and `%` are relative to the parent's font size,
 * `larger`/`smaller` step from it). undefined when the value is not one.
 */
function cssFontSizePx(value: string, parentPx: number): number | undefined {
  const text = value.trim().toLowerCase();
  const keyword = ABSOLUTE_FONT_SIZE_PX[text];
  if (keyword !== undefined) return keyword;
  if (text === "larger") return parentPx * 1.2;
  if (text === "smaller") return parentPx / 1.2;
  const match = /^(\d*\.?\d+)\s*(px|pt|pc|in|cm|mm|q|rem|em|%)$/.exec(text);
  if (!match) return undefined;
  const numeric = Number(match[1]);
  const unit = match[2];
  const pixels =
    unit === "em" ? numeric * parentPx : unit === "%" ? (numeric * parentPx) / 100 : numeric * FONT_SIZE_UNIT_PX[unit]!;
  return Number.isFinite(pixels) ? pixels : undefined;
}

/** A computed size in the editor's pixels when Excel can hold it. */
function fontSizeInRange(pixels: number | undefined): number | undefined {
  // Excel's own range, 1-409 pt (live audit B430: 73 pt was dropped).
  if (pixels === undefined || !Number.isFinite(pixels) || pixels < ptToPx(1) || pixels > ptToPx(409)) {
    return undefined;
  }
  return Math.round(pixels * 100) / 100;
}

function parseFontSize(value: string | undefined, parentPx: number): number | undefined {
  return value === undefined ? undefined : fontSizeInRange(cssFontSizePx(value, parentPx));
}

/**
 * The computed font size of an element from its table down (each relative
 * size resolved against its parent's computed size, B477), or undefined when
 * nothing on that chain states a size.
 */
function statedFontPx(element: Element, rules: CssRule[]): number | undefined {
  // One paste asks for every row and table once per cell: remember them.
  let cache = statedFontPxCache.get(rules);
  if (!cache) {
    cache = new Map();
    statedFontPxCache.set(rules, cache);
  }
  const known = cache.get(element);
  if (known !== undefined) return known ?? undefined;
  const parent = element.tagName.toLowerCase() === "table" ? null : element.parentElement;
  const parentPx = parent ? statedFontPx(parent, rules) : undefined;
  const value = elementDeclarations(element, rules).get("font-size");
  // An invalid declaration is ignored; `inherit` keeps the parent's size.
  const own = value === undefined ? undefined : cssFontSizePx(value, parentPx ?? DEFAULT_FONT_PX);
  const pixels = own ?? parentPx;
  cache.set(element, pixels ?? null);
  return pixels;
}

const statedFontPxCache = new WeakMap<CssRule[], Map<Element, number | null>>();

function parseFontFamily(value: string | undefined): string | undefined {
  const first = value
    ?.split(",")[0]
    ?.trim()
    .replace(/^['"]|['"]$/g, "");
  if (!first || first.length > 64) return undefined;
  return /^[\p{L}\p{N} _-]+$/u.test(first) ? first : undefined;
}

const HTML_BORDER_FLAG = {
  top: "borderTop",
  right: "borderRight",
  bottom: "borderBottom",
  left: "borderLeft",
} as const;

/** A CSS length in CSS pixels (px, pt, the keywords); undefined when unread. */
function cssBorderWidthPx(token: string): number | undefined {
  const keyword = { thin: 1, medium: 3, thick: 5 }[token];
  if (keyword !== undefined) return keyword;
  const match = /^(\d*\.?\d+)(px|pt|em|rem|in|cm|mm)?$/.exec(token);
  if (!match) return undefined;
  const value = Number(match[1]);
  const scale = { px: 1, pt: PX_PER_PT, em: 16, rem: 16, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 }[match[2] ?? "px"];
  return Number.isFinite(value) ? value * (scale ?? 1) : undefined;
}

const CSS_BORDER_LINES = new Set([
  "solid", "dashed", "dotted", "double", "groove", "ridge", "inset", "outset", "none", "hidden",
]);

/**
 * One side's border as the cell's declarations state it (the longhands over
 * `border-<side>` over `border`): a rich edge when a width, line or colour
 * is read, null when the side is only known to have a border, undefined
 * when it has none.
 */
function htmlBorderEdge(
  declarations: Map<string, string>,
  side: "top" | "right" | "bottom" | "left",
): RichBorderEdge | null | undefined {
  let width: number | undefined;
  let line: string | undefined;
  let color: string | undefined;
  let present = false;
  const read = (value: string | undefined) => {
    if (value === undefined) return;
    present = true;
    // rgb(1, 2, 3) is one token.
    for (const token of value.trim().toLowerCase().match(/[a-z-]+\([^)]*\)|[^\s]+/g) ?? []) {
      if (CSS_BORDER_LINES.has(token)) line = token;
      else if (cssBorderWidthPx(token) !== undefined) width = cssBorderWidthPx(token);
      else if (token === "windowtext" || token === "currentcolor" || token === "auto") color = undefined;
      else color = normalizeColor(token) ?? color;
    }
  };
  read(declarations.get("border"));
  read(declarations.get(`border-${side}`));
  const longhand = (property: string, apply: (value: string) => void) => {
    const value = declarations.get(`border-${side}-${property}`) ?? declarations.get(`border-${property}`);
    if (value !== undefined && !/\s/.test(value.trim())) {
      present = true;
      apply(value.trim().toLowerCase());
    }
  };
  longhand("style", (value) => {
    if (CSS_BORDER_LINES.has(value)) line = value;
  });
  longhand("width", (value) => {
    width = cssBorderWidthPx(value) ?? width;
  });
  longhand("color", (value) => {
    color = value === "windowtext" || value === "currentcolor" ? undefined : normalizeColor(value) ?? color;
  });
  if (!present || line === "none" || line === "hidden" || width === 0) return undefined;
  if (line === undefined && width === undefined && color === undefined) return null;
  // Excel's three weights of a solid line, by the width a browser draws:
  // up to 1.5px thin, up to 2.5px medium, thicker thick. A hairline
  // (Excel's own .5pt) is thin; dashes and dots have a medium form too.
  const px = width ?? 1;
  const weight = px > 2.5 ? "thick" : px > 1.5 ? "medium" : "thin";
  const style: BorderLineStyle =
    line === "double"
      ? "double"
      : line === "dashed"
        ? weight === "thin" ? "dashed" : "mediumDashed"
        : line === "dotted"
          ? "dotted"
          : weight;
  // A thin line in the automatic colour (Excel's `.5pt solid windowtext`)
  // or in this grid's own default border grey is the plain border.
  if (style === "thin" && (color === undefined || color === "#475569")) return null;
  return color ? { style, color } : { style };
}

function styleFromCell(
  cell: HTMLTableCellElement,
  rules: CssRule[],
): Record<string, unknown> | undefined {
  const declarations = elementDeclarations(cell, rules);
  // A header cell is bold by HTML's own stylesheet (th { font-weight: bold }),
  // an element rule that an author's declaration overrides and inheritance
  // does not (B487: a bare <th> pasted regular).
  if (cell.tagName.toLowerCase() === "th" && !declarations.has("font-weight")) {
    declarations.set("font-weight", "bold");
  }
  // The font properties CSS inherits come down from the row, row group and
  // table when the cell does not set them (live audit B435).
  const inherited = inheritedFontDeclarations(cell, rules);
  for (const [property, value] of inherited) {
    if (!declarations.has(property)) declarations.set(property, value);
  }
  const style: Record<string, unknown> = {};

  const weight = declarations.get("font-weight")?.toLowerCase();
  if (weight === "bold" || weight === "bolder" || Number(weight) >= 600) {
    style.bold = true;
  }
  if (/^(italic|oblique)$/i.test(declarations.get("font-style") || "")) {
    style.italic = true;
  }
  const decoration = declarations.get("text-decoration")?.toLowerCase() || "";
  if (decoration.includes("underline")) style.underline = true;
  if (decoration.includes("line-through")) style.strikethrough = true;

  const color = normalizeColor(declarations.get("color"));
  const backgroundColor = normalizeColor(
    declarations.get("background-color") || declarations.get("background"),
  );
  if (color) style.color = color;
  if (backgroundColor) style.backgroundColor = backgroundColor;

  const horizontal = declarations.get("text-align")?.toLowerCase();
  if (
    horizontal === "left" ||
    horizontal === "center" ||
    horizontal === "right"
  ) {
    style.textAlign = horizontal;
  }
  const vertical = declarations.get("vertical-align")?.toLowerCase();
  if (vertical === "top" || vertical === "middle" || vertical === "bottom") {
    style.verticalAlign = vertical;
  }
  if (declarations.get("white-space")?.toLowerCase() === "normal") {
    style.wrapText = true;
  }

  // The size the cell is shown at: an `em` or `%` on the cell or a row is
  // relative to its parent's size, not a size of its own (live audit B477).
  const fontSize = fontSizeInRange(statedFontPx(cell, rules));
  const fontFamily = parseFontFamily(declarations.get("font-family"));
  if (fontSize !== undefined) style.fontSize = fontSize;
  if (fontFamily) style.fontFamily = fontFamily;

  const numberFormat = unquoteCssValue(
    declarations.get("mso-number-format") || "",
  );
  if (
    numberFormat &&
    numberFormat.length <= 128 &&
    !/[\u0000-\u001f\u007f]/.test(numberFormat)
  ) {
    style.numberFormat = numberFormat;
  }

  // A border the HTML states is that border: its width, line and colour
  // (`border-top: 3px solid #ff0000` is a thick red top edge, Excel's own
  // clipboard HTML writes `.5pt solid windowtext`). Two edges used to paste
  // as the same default grey line whatever they said (live audit B449). An
  // edge that states nothing but its presence keeps the plain flag.
  const edges: CellBorderEdges = {};
  let bordered = false;
  for (const property of declarations.keys()) {
    if (property.startsWith("border")) {
      bordered = true;
      break;
    }
  }
  for (const side of bordered ? (["top", "right", "bottom", "left"] as const) : []) {
    const stated = htmlBorderEdge(declarations, side);
    if (stated === undefined) continue;
    if (stated === null) {
      style[HTML_BORDER_FLAG[side]] = true;
    } else {
      edges[side] = stated;
    }
  }
  if (Object.keys(edges).length > 0) style.borderEdges = edges;
  if (style.borderTop && style.borderRight && style.borderBottom && style.borderLeft && declarations.has("border")) {
    for (const flag of Object.values(HTML_BORDER_FLAG)) delete style[flag];
    style.border = true;
  }

  return Object.keys(style).length > 0 ? style : undefined;
}

/**
 * Elements that start and end a line of their own inside a cell, as a browser
 * lays them out: <div>Top</div><div>Next</div> is two lines, like Top<br>Next,
 * where an inline <b> or <span> joins its text to its neighbours.
 */
const BLOCK_TAGS = new Set([
  "address", "article", "aside", "blockquote", "dd", "div", "dl", "dt",
  "figcaption", "figure", "footer", "h1", "h2", "h3", "h4", "h5", "h6",
  "header", "hr", "li", "main", "nav", "ol", "p", "pre", "section", "ul",
]);

/**
 * Walk a cell's content once for both its plain text and its runs, so the two
 * cannot disagree: script/style are skipped, <br> is a line break, a block
 * element sits on its own line, and a non-breaking space stays U+00A0 (A&nbsp;B
 * pasted from HTML equals the same text pasted as plain text, and differs
 * from "A B").
 */
function walkCellContent<S>(
  cell: Element,
  root: S,
  styleOf: (element: Element, inherited: S) => S,
  emit: (text: string, style: S) => boolean,
  hidden: (element: Element) => boolean = () => false,
): void {
  let hasContent = false;
  let lastWasBreak = false;
  let pendingBreak = false;
  let stopped = false;
  const put = (text: string, style: S) => {
    if (stopped || !text) return;
    if (pendingBreak && hasContent && !lastWasBreak) {
      if (!emit("\n", style)) {
        stopped = true;
        return;
      }
    }
    pendingBreak = false;
    if (!emit(text, style)) {
      stopped = true;
      return;
    }
    hasContent = true;
    lastWasBreak = text.endsWith("\n");
  };
  const visit = (node: Node, inherited: S): void => {
    if (stopped) return;
    if (node.nodeType === 3) {
      put((node.nodeValue || "").replace(/\r\n?/g, "\n"), inherited);
      return;
    }
    if (!(node instanceof Element)) return;
    const tag = node.tagName.toLowerCase();
    if (tag === "script" || tag === "style" || tag === "noscript") return;
    // Words the source page does not show are not pasted (live audit B437).
    if (hidden(node)) return;
    if (tag === "br") {
      pendingBreak = false;
      put("\n", inherited);
      return;
    }
    const block = BLOCK_TAGS.has(tag);
    if (block && hasContent) pendingBreak = true;
    const style = styleOf(node, inherited);
    for (const child of Array.from(node.childNodes)) {
      visit(child, style);
      if (stopped) return;
    }
    if (block && hasContent) pendingBreak = true;
  };
  for (const child of Array.from(cell.childNodes)) {
    visit(child, root);
    if (stopped) return;
  }
}

/**
 * The value a table cell says it holds, beyond its display text, in Excel's
 * own HTML vocabulary: `x:num="123.456789"` is the exact number behind a
 * displayed 123.46, and `x:str` or a Text format (mso-number-format "\@")
 * says the text is text (0007 stays four characters, not the number 7).
 */
function cellValueIdentity(
  cell: Element,
  text: string,
  style: Record<string, unknown> | undefined,
): Pick<RichClipboardCell, "raw" | "kind" | "scalar"> | undefined {
  const excelNumber = cell.getAttribute("x:num");
  if (excelNumber !== null && excelNumber.trim() !== "") {
    const value = Number(excelNumber.trim());
    if (Number.isFinite(value) && excelNumber.length <= 64) {
      const raw = Object.is(value, -0) ? "0" : String(value);
      return { raw, kind: "number", scalar: { type: "number", value } };
    }
  }
  const textFormat = style?.numberFormat === "@" || style?.numberFormat === "\\@";
  if (text !== "" && (cell.hasAttribute("x:str") || textFormat)) {
    return { raw: text, kind: "text", scalar: { type: "text", value: text } };
  }
  return undefined;
}

/** A cell whose whole text is non-breaking spaces is an empty cell: that is
 *  how an HTML table writer fills a cell it has nothing to put in. */
function withoutPlaceholderSpace(text: string): string {
  return /^[\u00a0\s]+$/.test(text) && text.includes("\u00a0") ? "" : text;
}

function cellText(cell: HTMLTableCellElement, rules: CssRule[]): string {
  let text = "";
  walkCellContent(cell, null, () => null, (part) => {
    text += part;
    return text.length < MAX_CELL_CHARS;
  }, (element) => isHiddenElement(element, rules));
  return withoutPlaceholderSpace(text).slice(0, MAX_CELL_CHARS);
}

function runStyleFromElement(
  element: Element,
  inherited: NonNullable<TextRun["style"]>,
  cellFontPx: number,
  rules: CssRule[] = [],
): NonNullable<TextRun["style"]> {
  const style = { ...inherited };
  const tag = element.tagName.toLowerCase();
  if (tag === "b" || tag === "strong") style.bold = true;
  if (tag === "i" || tag === "em") style.italic = true;
  if (tag === "u") style.underline = true;
  if (tag === "s" || tag === "strike" || tag === "del") style.strike = true;

  // The element's cascade: the stylesheet's rules that match it (a
  // `<span class="runHot">` carries its class's colour, weight and size as
  // the cell does) under its inline style (B479).
  const declarations = elementDeclarations(element, rules);
  const weight = declarations.get("font-weight")?.toLowerCase();
  if (weight === "normal" || Number(weight) < 600) delete style.bold;
  else if (
    weight === "bold" ||
    weight === "bolder" ||
    Number(weight) >= 600
  ) {
    style.bold = true;
  }
  const fontStyle = declarations.get("font-style")?.toLowerCase();
  if (fontStyle === "normal") delete style.italic;
  else if (fontStyle === "italic" || fontStyle === "oblique") {
    style.italic = true;
  }
  const decoration = (
    declarations.get("text-decoration") ||
    declarations.get("text-decoration-line") ||
    ""
  ).toLowerCase();
  if (decoration.includes("none")) {
    delete style.underline;
    delete style.strike;
  } else {
    if (decoration.includes("underline")) style.underline = true;
    if (decoration.includes("line-through")) style.strike = true;
  }

  const color = normalizeColor(
    declarations.get("color") || element.getAttribute("color") || undefined,
  );
  if (color) style.color = color;
  // Relative to the run's parent: the enclosing sized run, else the cell.
  const fontSize = parseFontSize(declarations.get("font-size"), inherited.fontSize ?? cellFontPx);
  if (fontSize !== undefined) style.fontSize = fontSize;
  const fontFamily = parseFontFamily(
    declarations.get("font-family") ||
      element.getAttribute("face") ||
      undefined,
  );
  if (fontFamily) style.fontFamily = fontFamily;
  return style;
}

function richRunsFromCell(
  cell: HTMLTableCellElement,
  flatText: string,
  rules: CssRule[],
  cellStyle: Record<string, unknown> | undefined,
): TextRun[] | undefined {
  const runs: TextRun[] = [];
  let remaining = MAX_CELL_CHARS;
  let overflowed = false;

  const append = (
    raw: string,
    style: NonNullable<TextRun["style"]>,
  ): boolean => {
    if (!raw || remaining <= 0 || overflowed) return !overflowed && remaining > 0;
    const text = raw.slice(0, remaining);
    remaining -= text.length;
    if (!text) return remaining > 0;
    const nextStyle = Object.keys(style).length > 0 ? { ...style } : undefined;
    const previous = runs.at(-1);
    if (
      previous &&
      JSON.stringify(previous.style || {}) === JSON.stringify(nextStyle || {})
    ) {
      previous.text += text;
      return remaining > 0;
    }
    if (runs.length >= MAX_RICH_RUNS) {
      overflowed = true;
      return false;
    }
    runs.push({ text, ...(nextStyle ? { style: nextStyle } : {}) });
    return remaining > 0;
  };

  // The runs start from the cell's own weight and slant, so a span that
  // explicitly sets them back to normal is a run that is NOT bold (live audit
  // B436) - a run model that only records "bold" could not say so otherwise.
  const root: NonNullable<TextRun["style"]> = {};
  if (cellStyle?.bold === true) root.bold = true;
  if (cellStyle?.italic === true) root.italic = true;
  const cellFontPx = statedFontPx(cell, rules) ?? DEFAULT_FONT_PX;
  walkCellContent<NonNullable<TextRun["style"]>>(
    cell,
    root,
    (element, inherited) => runStyleFromElement(element, inherited, cellFontPx, rules),
    append,
    (element) => isHiddenElement(element, rules),
  );
  if (overflowed) return undefined;
  if (flatText === "") return undefined;
  // Text all in the cell's own weight and slant is not rich text: Excel's
  // bold header cell is a bold cell, not a cell with one bold run (which made
  // every pasted bold cell a rich-text cell).
  const rootStyle = JSON.stringify(root);
  if (runs.every((run) => JSON.stringify(run.style ?? {}) === rootStyle)) return undefined;
  // A run that turned the cell's bold or italic off: the runs carry it, not
  // the whole cell.
  if (cellStyle) {
    if (cellStyle.bold === true && runs.some((run) => !run.style?.bold)) delete cellStyle.bold;
    if (cellStyle.italic === true && runs.some((run) => !run.style?.italic)) delete cellStyle.italic;
  }
  return normalizeCellRuns(runs, flatText);
}

function formulaFromJson(value: unknown, depth = 0): string | undefined {
  if (depth > 4) return undefined;
  if (
    typeof value === "string" &&
    value.startsWith("=") &&
    value.length <= MAX_FORMULA_CHARS
  ) {
    return value;
  }
  if (Array.isArray(value)) {
    for (const item of value.slice(0, 20)) {
      const formula = formulaFromJson(item, depth + 1);
      if (formula) return formula;
    }
  } else if (value && typeof value === "object") {
    for (const item of Object.values(value).slice(0, 20)) {
      const formula = formulaFromJson(item, depth + 1);
      if (formula) return formula;
    }
  }
  return undefined;
}

function formulaFromCell(
  cell: HTMLTableCellElement,
  anchor: { row: number; col: number },
): string | undefined {
  for (const attribute of [
    "x:fmla",
    "data-formula",
    "data-cell-formula",
    "data-sheets-formula",
  ]) {
    const raw = cell.getAttribute(attribute)?.trim();
    if (!raw) continue;
    if (raw.startsWith("=") && raw.length <= MAX_FORMULA_CHARS) {
      // Excel/WPS HTML uses x:fmla for its relocatable R1C1 expression. The
      // converter is intentionally token-aware: an A1 expression in the same
      // attribute remains unchanged, while RC[-1] resolves against the actual
      // destination cell instead of being stored as an unsupported formula.
      const formula =
        attribute === "x:fmla" ? formulaR1C1ToA1(raw, anchor) : raw;
      return formula.length <= MAX_FORMULA_CHARS ? formula : undefined;
    }
    if (
      attribute === "data-sheets-formula" &&
      raw.length <= MAX_FORMULA_CHARS * 2
    ) {
      try {
        const formula = formulaFromJson(JSON.parse(raw));
        if (formula) return formula;
      } catch {
        // Malformed application metadata is ignored; the displayed text remains.
      }
    }
  }
  return undefined;
}

function positiveSpan(raw: string | null, maximum: number): number {
  const numeric = Number(raw);
  return Number.isInteger(numeric) && numeric > 0
    ? Math.min(maximum, numeric)
    : 1;
}

/**
 * Parse one interoperable HTML table clipboard payload.
 *
 * Returns null for non-table, oversized or empty payloads so the caller can
 * fall back to its existing quote-aware TSV path.
 */
export function parseSpreadsheetClipboardHtml(
  html: string,
  options: SpreadsheetClipboardHtmlParseOptions = {},
): ParsedSpreadsheetClipboardHtml | null {
  if (
    typeof html !== "string" ||
    html.length === 0 ||
    html.length > MAX_HTML_CHARS ||
    typeof DOMParser === "undefined"
  ) {
    return null;
  }
  const document = new DOMParser().parseFromString(html, "text/html");
  const table = document.querySelector("table");
  if (!(table instanceof HTMLTableElement)) return null;
  const rules = collectCssRules(document);
  const baseRow = Math.max(0, Math.trunc(options.baseRow ?? 0));
  const baseCol = Math.max(0, Math.trunc(options.baseCol ?? 0));
  // Rows in the order a table shows them: the header group, the bodies, then
  // the footer group wherever its markup sits (live audit B433). A row the
  // source does not show is not pasted (B437).
  const groupRank = (row: Element) => {
    const group = row.parentElement?.tagName.toLowerCase();
    return group === "thead" ? 0 : group === "tfoot" ? 2 : 1;
  };
  const allRows = Array.from(table.querySelectorAll("tr"))
    .filter((row) => row.closest("table") === table)
    .filter(
      (row) =>
        !isHiddenElement(row, rules) &&
        !(row.parentElement && row.parentElement !== table && isHiddenElement(row.parentElement, rules)),
    )
    .map((row, index) => ({ row, index, rank: groupRank(row) }))
    .sort((a, b) => a.rank - b.rank || a.index - b.index)
    .map(({ row }) => row);
  const sourceRows = allRows.slice(0, MAX_ROWS);
  if (sourceRows.length === 0) return null;
  let truncated = allRows.length > MAX_ROWS;

  const cells: Array<Array<RichClipboardCell | undefined>> = [];
  const merges: SpreadsheetClipboardMerge[] = [];
  let hasRichMetadata = false;
  for (let rowIndex = 0; rowIndex < sourceRows.length; rowIndex += 1) {
    cells[rowIndex] ??= [];
    let columnIndex = 0;
    const allCells = Array.from(sourceRows[rowIndex].children).filter(
      (node): node is HTMLTableCellElement =>
        node instanceof HTMLTableCellElement,
    );
    if (allCells.length > MAX_COLS) truncated = true;
    const sourceCells = allCells.slice(0, MAX_COLS);
    for (const sourceCell of sourceCells) {
      while (cells[rowIndex][columnIndex] && columnIndex < MAX_COLS) {
        columnIndex += 1;
      }
      if (columnIndex >= MAX_COLS) break;
      const text = cellText(sourceCell, rules);
      const formula = formulaFromCell(sourceCell, {
        row: baseRow + rowIndex,
        col: baseCol + columnIndex,
      });
      const style = styleFromCell(sourceCell, rules);
      const identity = formula ? undefined : cellValueIdentity(sourceCell, text, style);
      const runs =
        formula || identity?.kind === "number"
          ? undefined
          : richRunsFromCell(sourceCell, text, rules, style);
      const anchor: RichClipboardCell = {
        raw: formula || identity?.raw || text,
        computed: text,
        ...(identity ? { kind: identity.kind, scalar: identity.scalar } : {}),
        ...(style ? { style } : {}),
        ...(runs ? { runs } : {}),
      };
      const link = sourceCell.querySelector("a[href]");
      const hyperlink = sanitizeHyperlink(link?.getAttribute("href"));
      if (hyperlink) {
        anchor.hyperlink = hyperlink;
        anchor.hyperlinkDisplay = text;
      }
      // Excel's exact number behind a rounded display is metadata worth the
      // rich paste too: x:num="123.456789" shown as 123.46 (live audit B420).
      if (formula || style || runs || hyperlink || identity) hasRichMetadata = true;

      // rowspan="0" spans the rest of the cell's row group (HTML), here the
      // rows of that group still ahead (live audit B434).
      const group = sourceRows[rowIndex].parentElement;
      let rowsLeftInGroup = 0;
      for (let next = rowIndex; next < sourceRows.length && sourceRows[next].parentElement === group; next += 1) {
        rowsLeftInGroup += 1;
      }
      const rowSpan =
        sourceCell.getAttribute("rowspan")?.trim() === "0"
          ? Math.max(1, Math.min(rowsLeftInGroup, MAX_ROWS - rowIndex))
          : positiveSpan(sourceCell.getAttribute("rowspan"), MAX_ROWS - rowIndex);
      const colSpan = positiveSpan(
        sourceCell.getAttribute("colspan"),
        MAX_COLS - columnIndex,
      );
      if (rowSpan > 1 || colSpan > 1) {
        merges.push({
          row: rowIndex,
          col: columnIndex,
          rowSpan,
          colSpan,
        });
        hasRichMetadata = true;
      }
      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        const targetRow = rowIndex + rowOffset;
        if (targetRow >= MAX_ROWS) break;
        cells[targetRow] ??= [];
        for (let colOffset = 0; colOffset < colSpan; colOffset += 1) {
          const targetCol = columnIndex + colOffset;
          if (targetCol >= MAX_COLS) break;
          cells[targetRow][targetCol] =
            rowOffset === 0 && colOffset === 0
              ? anchor
              : { raw: "", computed: "" };
        }
      }
      columnIndex += colSpan;
    }
  }

  const rowCount = Math.min(MAX_ROWS, cells.length);
  const columnCount = Math.min(
    MAX_COLS,
    Math.max(0, ...cells.slice(0, rowCount).map((row) => row.length)),
  );
  if (rowCount === 0 || columnCount === 0) return null;
  const normalizedCells: RichClipboard = Array.from(
    { length: rowCount },
    (_, row) =>
      Array.from(
        { length: columnCount },
        (_, col) => cells[row]?.[col] || { raw: "", computed: "" },
      ),
  );
  return {
    rows: normalizedCells.map((row) => row.map((cell) => cell.raw)),
    cells: normalizedCells,
    merges: merges.filter(
      (merge) =>
        merge.row + merge.rowSpan <= rowCount &&
        merge.col + merge.colSpan <= columnCount,
    ),
    hasRichMetadata,
    truncated,
  };
}
