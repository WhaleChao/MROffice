/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeTextOutlineCss } from "../office-text-outline/office-text-outline-css";
import { OFFICE_CHART_TEXT_CAPS, OFFICE_CHART_TEXT_SPACING } from "./chart-letter-spacing";
import { ptToPx } from "../office-units";
import { normalizeOfficeShapeFontFamily } from "../office-shape-text-font";
import { primaryOfficeFontFace } from "../../components/office-ribbon/officeFontFace";
import { windowsMissingFaceAdapter, windowsOfficeFace } from "../office-fonts/windows-office-faces";

/**
 * The Font dialog's Underline style list, in its order (ECMA-376
 * `ST_TextUnderlineType` less `none`): Home's Underline writes `single` and
 * `double`; the dialog offers every one.
 */
export const OFFICE_CHART_TEXT_UNDERLINES = [
  "words",
  "single",
  "double",
  "heavy",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dotDashHeavy",
  "dotDotDash",
  "dotDotDashHeavy",
  "wavy",
  "wavyHeavy",
  "wavyDouble",
] as const;
export type OfficeChartTextUnderline =
  (typeof OFFICE_CHART_TEXT_UNDERLINES)[number];

/** `a:defRPr@sz` is hundredths of a point in 100..400000 (§21.1.2.3.9). */
export const OFFICE_CHART_TEXT_MIN_SIZE_PT = 1;
export const OFFICE_CHART_TEXT_MAX_SIZE_PT = 400;
/** Font ▸ Effects ▸ Offset: a superscript is raised, a subscript lowered,
 * by this share of the font's height (`a:defRPr@baseline`, thousandths of
 * a percent). Office's dialog keeps it within -100%..100%. */
export const OFFICE_CHART_TEXT_MAX_BASELINE = 100;
/** A raised (superscript) or lowered (subscript) run draws at two thirds of
 * its size, as PowerPoint draws a chart's `baseline` run (the CDC TB deck's
 * § and * marks). */
export const OFFICE_CHART_SUPERSCRIPT_SCALE = 2 / 3;
/** Character Spacing ▸ Expanded / Condensed By, in points (`a:defRPr@spc`,
 * hundredths of a point); the dialog's spinner stops at 1584 pt. */
export const OFFICE_CHART_TEXT_MAX_SPACING_PT = 1584;
/** Character Spacing ▸ Kerning for fonts N Points and above (`a:defRPr@kern`). */
export const OFFICE_CHART_TEXT_MAX_KERNING_PT = 1584;

export interface OfficeChartText {
  /** `a:latin@typeface` (written for `a:cs`, and for `a:ea` unless the
   * text names an Asian font of its own). */
  readonly fontFamily?: string;
  /** Font ▸ Asian text font: `a:ea@typeface`, stated only beside a Latin
   * font and only when it is another face. */
  readonly eastAsianFontFamily?: string;
  /** `a:defRPr@sz` / 100, to a hundredth of a point. */
  readonly fontSizePt?: number;
  /** `a:defRPr@b`. */
  readonly bold?: boolean;
  /** `a:defRPr@i`. */
  readonly italic?: boolean;
  /** `a:defRPr@u`. */
  readonly underline?: OfficeChartTextUnderline;
  /** Font ▸ Underline color: `a:uFill/a:solidFill/a:srgbClr`. Absent is
   * Automatic (the text's own colour). */
  readonly underlineColor?: `#${string}`;
  /** `a:solidFill/a:srgbClr`, lower-case `#rrggbb`. */
  readonly color?: `#${string}`;
  /** Font ▸ Effects ▸ Strikethrough / Double Strikethrough
   * (`a:defRPr@strike` `sngStrike` / `dblStrike`). */
  readonly strike?: "single" | "double";
  /** Font ▸ Effects ▸ Superscript (a positive Offset) / Subscript (a
   * negative one), in percent of the font's height (`a:defRPr@baseline`). */
  readonly baseline?: number;
  /** Font ▸ Effects ▸ Small Caps / All Caps (`a:defRPr@cap`). */
  readonly caps?: "small" | "all";
  /** Font ▸ Effects ▸ Equalize Character Height (`a:defRPr@normalizeH`). */
  readonly equalizeHeight?: true;
  /** Character Spacing ▸ Spacing: Expanded (positive) or Condensed
   * (negative) By, in points (`a:defRPr@spc`). Absent is Normal. */
  readonly spacingPt?: number;
  /** Character Spacing ▸ Kerning for fonts N Points and above
   * (`a:defRPr@kern`). Absent is no kerning. */
  readonly kerningPt?: number;
  /** Text Effects ▸ Glow (`a:effectLst/a:glow`): its colour, size (`rad`)
   * in points and opacity (`a:alpha`; absent is opaque). */
  readonly glow?: OfficeChartTextGlow;
  



  readonly outline?: OfficeChartTextOutline;
}

export interface OfficeChartTextOutline {
  readonly color: `#${string}`;
  readonly widthPt: number;
}

/** Office's Text Outline Width spinner's bound, points. */
export const OFFICE_CHART_TEXT_MAX_OUTLINE_PT = 1584;

export interface OfficeChartTextGlow {
  readonly color: `#${string}`;
  readonly radiusPt: number;
  readonly opacity?: number;
}

/** Office's Glow size spinner's bound, points. */
export const OFFICE_CHART_TEXT_MAX_GLOW_RADIUS_PT = 150;

export const OFFICE_CHART_TEXT_KEYS = [
  "fontFamily",
  "eastAsianFontFamily",
  "fontSizePt",
  "bold",
  "italic",
  "underline",
  "underlineColor",
  "color",
  "strike",
  "baseline",
  "caps",
  "equalizeHeight",
  "spacingPt",
  "kerningPt",
  "glow",
  "outline",
] as const satisfies readonly (keyof OfficeChartText)[];

const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;
/** The model's underline names in DrawingML (`ST_TextUnderlineType`). */
export const OFFICE_CHART_TEXT_UNDERLINE_TO_DRAWING: Record<OfficeChartTextUnderline, string> = {
  words: "words",
  single: "sng",
  double: "dbl",
  heavy: "heavy",
  dotted: "dotted",
  dottedHeavy: "dottedHeavy",
  dash: "dash",
  dashHeavy: "dashHeavy",
  dashLong: "dashLong",
  dashLongHeavy: "dashLongHeavy",
  dotDash: "dotDash",
  dotDashHeavy: "dotDashHeavy",
  dotDotDash: "dotDotDash",
  dotDotDashHeavy: "dotDotDashHeavy",
  wavy: "wavy",
  wavyHeavy: "wavyHeavy",
  wavyDouble: "wavyDbl",
};

/** A number within `[low, high]` to `places` decimals, else null. */
function bounded(value: unknown, low: number, high: number, places = 2): number | null {
  if (typeof value !== "number" || !Number.isFinite(value) || value < low || value > high) return null;
  const scale = 10 ** places;
  return Math.round(value * scale) / scale;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * The chart text as the model keeps it, or `undefined` when the value is
 * not one: an unknown key, a value outside its range or nothing stated at
 * all. Strict on purpose — a host commits what it is given and a file
 * writer never guesses.
 */
export function normalizeOfficeChartText(
  value: unknown,
): OfficeChartText | undefined {
  if (!isRecord(value)) return undefined;
  if (
    Object.keys(value).some(
      (key) => !(OFFICE_CHART_TEXT_KEYS as readonly string[]).includes(key),
    )
  ) {
    return undefined;
  }
  const next: {
    -readonly [K in keyof OfficeChartText]: OfficeChartText[K];
  } = {};
  if (value.fontFamily !== undefined) {
    const family = normalizeOfficeShapeFontFamily(value.fontFamily);
    if (!family || family !== value.fontFamily) return undefined;
    next.fontFamily = family;
  }
  if (value.fontSizePt !== undefined) {
    const size = value.fontSizePt;
    if (
      typeof size !== "number" ||
      !Number.isFinite(size) ||
      size < OFFICE_CHART_TEXT_MIN_SIZE_PT ||
      size > OFFICE_CHART_TEXT_MAX_SIZE_PT
    ) {
      return undefined;
    }
    next.fontSizePt = Math.round(size * 100) / 100;
  }
  for (const key of ["bold", "italic"] as const) {
    if (value[key] === undefined) continue;
    if (typeof value[key] !== "boolean") return undefined;
    next[key] = value[key] as boolean;
  }
  if (value.underline !== undefined) {
    if (
      !(OFFICE_CHART_TEXT_UNDERLINES as readonly unknown[]).includes(
        value.underline,
      )
    ) {
      return undefined;
    }
    next.underline = value.underline as OfficeChartTextUnderline;
  }
  for (const key of ["color", "underlineColor"] as const) {
    if (value[key] === undefined) continue;
    const color = value[key];
    if (typeof color !== "string" || !HEX_COLOR_RE.test(color)) return undefined;
    next[key] = color.toLowerCase() as `#${string}`;
  }
  if (value.eastAsianFontFamily !== undefined) {
    // An Asian face of its own, beside the Latin one: a text that names
    // only one face names it for both (`fontFamily`).
    const family = normalizeOfficeShapeFontFamily(value.eastAsianFontFamily);
    if (!family || family !== value.eastAsianFontFamily || !next.fontFamily || family === next.fontFamily) {
      return undefined;
    }
    next.eastAsianFontFamily = family;
  }
  if (value.strike !== undefined) {
    if (value.strike !== "single" && value.strike !== "double") return undefined;
    next.strike = value.strike;
  }
  if (value.caps !== undefined) {
    if (value.caps !== "small" && value.caps !== "all") return undefined;
    next.caps = value.caps;
  }
  if (value.equalizeHeight !== undefined) {
    if (value.equalizeHeight !== true) return undefined;
    next.equalizeHeight = true;
  }
  if (value.baseline !== undefined) {
    const offset = bounded(value.baseline, -OFFICE_CHART_TEXT_MAX_BASELINE, OFFICE_CHART_TEXT_MAX_BASELINE, 3);
    if (offset === null || offset === 0) return undefined;
    next.baseline = offset;
  }
  if (value.spacingPt !== undefined) {
    const spacing = bounded(value.spacingPt, -OFFICE_CHART_TEXT_MAX_SPACING_PT, OFFICE_CHART_TEXT_MAX_SPACING_PT);
    if (spacing === null || spacing === 0) return undefined;
    next.spacingPt = spacing;
  }
  if (value.kerningPt !== undefined) {
    const kerning = bounded(value.kerningPt, 0, OFFICE_CHART_TEXT_MAX_KERNING_PT);
    if (kerning === null || kerning === 0) return undefined;
    next.kerningPt = kerning;
  }
  if (value.glow !== undefined) {
    const glow = value.glow;
    if (!isRecord(glow) || Object.keys(glow).some((key) => !["color", "radiusPt", "opacity"].includes(key))) return undefined;
    const radius = bounded(glow.radiusPt, 0, OFFICE_CHART_TEXT_MAX_GLOW_RADIUS_PT);
    const opacity = glow.opacity === undefined ? 1 : bounded(glow.opacity, 0, 1, 4);
    if (typeof glow.color !== "string" || !HEX_COLOR_RE.test(glow.color) || !radius || !opacity) return undefined;
    next.glow = {
      color: glow.color.toLowerCase() as `#${string}`,
      radiusPt: radius,
      ...(opacity < 1 ? { opacity } : {}),
    };
  }
  if (value.outline !== undefined) {
    const outline = value.outline;
    if (!isRecord(outline) || Object.keys(outline).some((key) => !["color", "widthPt"].includes(key))) return undefined;
    const width = bounded(outline.widthPt, 0, OFFICE_CHART_TEXT_MAX_OUTLINE_PT);
    if (typeof outline.color !== "string" || !HEX_COLOR_RE.test(outline.color) || !width) return undefined;
    next.outline = { color: outline.color.toLowerCase() as `#${string}`, widthPt: width };
  }
  return Object.keys(next).length ? next : undefined;
}








function glowColor(glow: OfficeChartTextGlow): string {
  const channel = (offset: number) => Number.parseInt(glow.color.slice(offset, offset + 2), 16);
  return `rgba(${channel(1)}, ${channel(3)}, ${channel(5)}, ${Math.min(1, (glow.opacity ?? 1) * 2.5)})`;
}

/** A glow's blur, px. */
const glowBlurPx = (glow: OfficeChartTextGlow) => ptToPx(glow.radiusPt) / 3;

/**
 * One Home ▸ Font command over the chart's text: the next chart text, or
 * `null` when the command changes nothing. A toggle reads the chart as a
 * whole, as Excel does with a shape: bold turns off only when it is on.
 */
export type OfficeChartTextCommand =
  | { kind: "family"; family: string | null }
  | { kind: "size"; pt: number }
  | { kind: "toggle"; field: "bold" | "italic" }
  | { kind: "underline"; variant: OfficeChartTextUnderline }
  /** `null` is Automatic: the chart's text goes back to the theme's colour. */
  | { kind: "color"; color: string | null }
  /** The Font dialog's OK: the whole text it shows (`undefined`: nothing stated). */
  | { kind: "replace"; text: OfficeChartText | undefined };

export function applyOfficeChartTextCommand(
  current: OfficeChartText | undefined,
  command: OfficeChartTextCommand,
): OfficeChartText | null | undefined {
  const draft: Record<string, unknown> = { ...(current ?? {}) };
  switch (command.kind) {
    case "family": {
      // The chart's text names ONE typeface (`a:latin@typeface`). A font menu
      // hands over a CSS stack (`Georgia, "Times New Roman", serif`), whose
      // quotes broke the drawing's style attribute: the text lost its font
      // and weight and a WordArt title vanished (PPT-B121). Home's font box
      // sets the face for all the chart's text, its Asian text too.
      const face = primaryOfficeFontFace(command.family);
      if (face) draft.fontFamily = face;
      else delete draft.fontFamily;
      delete draft.eastAsianFontFamily;
      break;
    }
    case "replace":
      for (const key of Object.keys(draft)) delete draft[key];
      Object.assign(draft, command.text ?? {});
      break;
    case "size":
      draft.fontSizePt = command.pt;
      break;
    case "toggle":
      if (current?.[command.field] === true) delete draft[command.field];
      else draft[command.field] = true;
      break;
    case "underline":
      if (current?.underline === command.variant) delete draft.underline;
      else draft.underline = command.variant;
      break;
    case "color":
      if (command.color) draft.color = command.color.toLowerCase();
      else delete draft.color;
      break;
  }
  const next = Object.keys(draft).length
    ? normalizeOfficeChartText(draft)
    : undefined;
  if (Object.keys(draft).length && !next) return null;
  return JSON.stringify(next ?? null) === JSON.stringify(current ?? null)
    ? null
    : next;
}

/**
 * The face Office's chart engine draws chart text in when the face the text
 * names is not on the machine (Word for Mac, STRESS2 page 5 and its
 * variants: a txPr naming an unknown family, one the document only embeds
 * in word/fonts, or one whose font-table entry gives `altName` Times New
 * Roman and a roman/modern panose all print in Calibri; a theme minor font
 * of Georgia does not change it; an installed Courier New is used). The
 * chart engine never reads the document's embedded fonts nor its font
 * table's substitution hints. The stack ends in it, so a face the reader's
 * browser cannot show lands where Word's does; MRAI serves Calibri (or its
 * metric twin) on every page.
 */
export const OFFICE_CHART_MISSING_FACE = "Calibri";








const CHART_WINDOWS_FACES = windowsMissingFaceAdapter("excel", () => OFFICE_CHART_MISSING_FACE);

/** The ECharts text style the chart text draws with (px, as ECharts sizes). */
/**
 * The key an ECharts text style carries a text's underline and
 * strikethrough under: the engine draws neither, so the picture renderer
 * draws the lines itself (axis tick labels, `chart-render-server`, X-157).
 */
export const OFFICE_CHART_TEXT_DECORATION = "officeTextDecoration";

export type OfficeChartTextDecoration = Readonly<{
  underline?: OfficeChartTextUnderline;
  underlineColor?: `#${string}`;
  strike?: "single" | "double";
  /** Text Outline, stroked over the letters (the engine strokes under them). */
  outline?: Readonly<{ color: `#${string}`; widthPx: number }>;
}>;

export function officeChartTextEChartsStyle(
  text: OfficeChartText | undefined,
  scale = 1,
): Record<string, unknown> {
  if (!text) return {};
  const named = text.fontFamily ? primaryOfficeFontFace(text.fontFamily) : null;
  const latin = named ? windowsOfficeFace(named, CHART_WINDOWS_FACES)?.face ?? named : null;
  // An Asian face of its own draws the text the Latin face has no glyphs
  // for, as Office draws a run in its `a:ea` face.
  const asian = text.eastAsianFontFamily ? primaryOfficeFontFace(text.eastAsianFontFamily) : null;
  // Single quotes: the server's SVG writes this stack into a double-quoted
  // `style` attribute unescaped, and a double quote ended the attribute
  // there -- the picture lost its font-family (every chart whose text names
  // an Asian face of its own drew in the page's face).
  const quoted = (face: string) => (/^[\w-]+$/.test(face) ? face : `'${face.replace(/["']/g, "")}'`);
  return {
    // A page stored before the menu's stack was reduced still draws its face.
    ...(latin
      ? {
          fontFamily: [latin, ...(asian && asian !== latin ? [asian] : []), OFFICE_CHART_MISSING_FACE]
            .filter((face, index, faces) => faces.findIndex((other) => other.toLowerCase() === face.toLowerCase()) === index)
            .map(quoted)
            .join(", "),
        }
      : {}),
    ...(text.fontSizePt === undefined
      ? {}
      : { fontSize: (ptToPx(text.fontSizePt)) * scale }),
    ...(text.bold === undefined
      ? {}
      : { fontWeight: text.bold ? "bold" : "normal" }),
    ...(text.italic === undefined
      ? {}
      : { fontStyle: text.italic ? "italic" : "normal" }),
    ...(text.color ? { color: text.color } : {}),
    // Character Spacing and All / Small Caps: measured and drawn by the
    // picture renderer (`chart-letter-spacing`).
    ...(text.spacingPt ? { [OFFICE_CHART_TEXT_SPACING]: ptToPx(text.spacingPt) * scale } : {}),
    ...(text.caps ? { [OFFICE_CHART_TEXT_CAPS]: text.caps } : {}),
    // Text Effects ▸ Glow: a soft halo round the glyphs, as wide as its size.
    ...(text.glow
      ? {
          textShadowColor: glowColor(text.glow),
          textShadowBlur: glowBlurPx(text.glow) * scale,
          textShadowOffsetX: 0,
          textShadowOffsetY: 0,
        }
      : {}),
    ...(text.underline !== undefined || text.strike !== undefined || text.outline
      ? {
          [OFFICE_CHART_TEXT_DECORATION]: {
            ...(text.underline !== undefined ? { underline: text.underline } : {}),
            ...(text.underline !== undefined && text.underlineColor ? { underlineColor: text.underlineColor } : {}),
            ...(text.strike !== undefined ? { strike: text.strike } : {}),
            ...(text.outline
              ? { outline: { color: text.outline.color, widthPx: ptToPx(text.outline.widthPt) * scale } }
              : {}),
          } satisfies OfficeChartTextDecoration,
        }
      : {}),
  };
}

/** The CSS the chart text draws with (slide px: 4/3 px per point). */
export function officeChartTextCss(
  text: OfficeChartText | undefined,
): Record<string, string | number> {
  if (!text) return {};
  const style = officeChartTextEChartsStyle(text);
  return {
    ...(typeof style.fontFamily === "string" ? { fontFamily: style.fontFamily } : {}),
    ...(typeof style.fontSize === "number" ? { fontSize: style.fontSize } : {}),
    ...(typeof style.fontWeight === "string" ? { fontWeight: style.fontWeight } : {}),
    ...(typeof style.fontStyle === "string" ? { fontStyle: style.fontStyle } : {}),
    ...(text.underline !== undefined ? { textDecoration: "underline" } : {}),
    ...(text.spacingPt ? { letterSpacing: `${ptToPx(text.spacingPt)}px` } : {}),
    ...(text.caps === "all" ? { textTransform: "uppercase" } : text.caps === "small" ? { fontVariant: "small-caps" } : {}),
    ...(text.glow ? { textShadow: `0 0 ${glowBlurPx(text.glow)}px ${glowColor(text.glow)}` } : {}),
    // Office's outline (one round stroke of the merged glyphs, over the
    // fill) is the overlay's; the CSS stroke stands in until it draws.
    ...(text.outline
      ? officeTextOutlineCss(text.outline.color, text.outline.widthPt, `${ptToPx(text.outline.widthPt)}px ${text.outline.color}`)
      : {}),
  };
}

/**
 * Home's Font group state for a selected chart, in the "mixed"-aware shape
 * the editors' summaries use: a chart has one text, so nothing is mixed.
 */
export function officeChartTextSummary(text: OfficeChartText | undefined): {
  readonly fontFamily: string | undefined;
  readonly fontSizePt: number | undefined;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: OfficeChartTextUnderline | "none";
  readonly color: `#${string}` | undefined;
} {
  return {
    fontFamily: text?.fontFamily,
    fontSizePt: text?.fontSizePt,
    bold: text?.bold === true,
    italic: text?.italic === true,
    underline: text?.underline ?? "none",
    color: text?.color,
  };
}
