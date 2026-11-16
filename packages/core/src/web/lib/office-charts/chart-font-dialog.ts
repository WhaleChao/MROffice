/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The chart text's Font dialog (Excel and PowerPoint: right-click a chart ▸
 * Font..., with the chart area selected), as a draft over the one chart text
 * (`chart-text.ts`, `c:chartSpace/c:txPr`).
 *
 * Office's dialog has two pages. Font: Latin text font, Asian text font,
 * Font style, Size, Font color, Underline style and Underline color, and the
 * Effects -- Strikethrough, Double Strikethrough, Superscript / Subscript
 * with their Offset, Small Caps, All Caps, Equalize Character Height.
 * Character Spacing: Spacing (Normal, Expanded, Condensed) By points, and
 * Kerning for fonts N Points and above. Strikethrough and Double
 * Strikethrough, Superscript and Subscript, Small Caps and All Caps are pairs
 * that exclude each other. The draft holds every field as the dialog shows
 * it; OK writes only the fields the author changed, over the chart's text,
 * so a field left alone keeps what the chart states (or does not state).
 */

import type { ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";
import { primaryOfficeFontFace } from "../../components/office-ribbon/officeFontFace";
import {
  officeChartElementOwnText,
  officeChartElementText,
  officeChartWithElementText,
  type OfficeChartTextElement,
} from "./chart-element-text";
import {
  applyOfficeChartTextCommand,
  normalizeOfficeChartText,
  OFFICE_CHART_TEXT_MAX_BASELINE,
  OFFICE_CHART_TEXT_MAX_KERNING_PT,
  OFFICE_CHART_TEXT_MAX_SIZE_PT,
  OFFICE_CHART_TEXT_MAX_SPACING_PT,
  OFFICE_CHART_TEXT_MIN_SIZE_PT,
  type OfficeChartText,
  type OfficeChartTextUnderline,
} from "./chart-text";
import {
  PRESENTATION_UNDERLINE_FONT_PAGE,
  PRESENTATION_UNDERLINE_FROM_FONT_PAGE,
} from "../presentation-text-decoration";

/** The Word-spelled underline names the shared Font page lists. */
export type OfficeChartFontDialogUnderline =
  | "none"
  | "single"
  | "words"
  | "double"
  | "thick"
  | "dotted"
  | "dottedHeavy"
  | "dash"
  | "dashedHeavy"
  | "dashLong"
  | "dashLongHeavy"
  | "dotDash"
  | "dashDotHeavy"
  | "dotDotDash"
  | "dashDotDotHeavy"
  | "wavy"
  | "wavyHeavy"
  | "wavyDouble";

const UNDERLINE_TO_PAGE: Record<OfficeChartTextUnderline, Exclude<OfficeChartFontDialogUnderline, "none">> =
  PRESENTATION_UNDERLINE_FONT_PAGE;
const PAGE_TO_UNDERLINE: Record<Exclude<OfficeChartFontDialogUnderline, "none">, OfficeChartTextUnderline> =
  PRESENTATION_UNDERLINE_FROM_FONT_PAGE;

/** Office's Offset for a newly checked Superscript / Subscript. */
export const OFFICE_CHART_SUPERSCRIPT_OFFSET = 30;
export const OFFICE_CHART_SUBSCRIPT_OFFSET = 25;
/** Office's size for a newly checked Kerning for fonts. */
export const OFFICE_CHART_DEFAULT_KERNING_PT = 12;

export type OfficeChartFontSpacing = "normal" | "expanded" | "condensed";

export interface OfficeChartFontDraft {
  /** Latin text font: a face (or a catalogue font stack); null is the chart's default. */
  family: string | null;
  /** Asian text font: a face (or stack); null follows the Latin font. */
  asianFamily: string | null;
  bold: boolean;
  italic: boolean;
  /** Points as typed; "" is the chart's default size. */
  sizePt: string;
  /** #rrggbb; null is Automatic. */
  color: string | null;
  underline: OfficeChartFontDialogUnderline;
  /** #rrggbb; null is Automatic (the text's colour). */
  underlineColor: string | null;
  strike: boolean;
  doubleStrike: boolean;
  superscript: boolean;
  subscript: boolean;
  /** The Offset in percent, as typed (always positive in the box). */
  offset: string;
  smallCaps: boolean;
  allCaps: boolean;
  equalizeHeight: boolean;
  spacing: OfficeChartFontSpacing;
  /** Spacing By, points as typed. */
  spacingBy: string;
  kerning: boolean;
  /** Kerning for fonts N Points and above, as typed. */
  kerningPt: string;
}

export type OfficeChartFontDraftField = keyof OfficeChartFontDraft;

function formatNumber(value: number): string {
  return String(Math.round(value * 100) / 100);
}

/** The dialog as it opens on a chart's text. */
export function officeChartFontDraft(text: OfficeChartText | undefined): OfficeChartFontDraft {
  const baseline = text?.baseline ?? 0;
  const spacing = text?.spacingPt ?? 0;
  return {
    family: text?.fontFamily ?? null,
    asianFamily: text?.eastAsianFontFamily ?? null,
    bold: text?.bold === true,
    italic: text?.italic === true,
    sizePt: text?.fontSizePt === undefined ? "" : formatNumber(text.fontSizePt),
    color: text?.color ?? null,
    underline: text?.underline ? UNDERLINE_TO_PAGE[text.underline] : "none",
    underlineColor: text?.underlineColor ?? null,
    strike: text?.strike === "single",
    doubleStrike: text?.strike === "double",
    superscript: baseline > 0,
    subscript: baseline < 0,
    offset: baseline === 0 ? "" : formatNumber(Math.abs(baseline)),
    smallCaps: text?.caps === "small",
    allCaps: text?.caps === "all",
    equalizeHeight: text?.equalizeHeight === true,
    spacing: spacing > 0 ? "expanded" : spacing < 0 ? "condensed" : "normal",
    spacingBy: spacing === 0 ? "" : formatNumber(Math.abs(spacing)),
    kerning: text?.kerningPt !== undefined,
    kerningPt: text?.kerningPt === undefined ? "" : formatNumber(text.kerningPt),
  };
}

/**
 * One change in the dialog, with Office's exclusive pairs kept: checking
 * Strikethrough clears Double Strikethrough, Superscript clears Subscript,
 * Small Caps clears All Caps (and back), and a newly checked Superscript,
 * Subscript or Kerning takes Office's own value when its box is empty.
 * Returns the next draft and the fields it touched.
 */
export function changeOfficeChartFontDraft<K extends OfficeChartFontDraftField>(
  draft: OfficeChartFontDraft,
  field: K,
  value: OfficeChartFontDraft[K],
): { draft: OfficeChartFontDraft; touched: OfficeChartFontDraftField[] } {
  const next: OfficeChartFontDraft = { ...draft, [field]: value };
  const touched: OfficeChartFontDraftField[] = [field];
  const pairs: [OfficeChartFontDraftField, OfficeChartFontDraftField][] = [
    ["strike", "doubleStrike"],
    ["superscript", "subscript"],
    ["smallCaps", "allCaps"],
  ];
  for (const [left, right] of pairs) {
    const other = field === left ? right : field === right ? left : null;
    if (other && value === true && next[other] === true) {
      (next as unknown as Record<string, unknown>)[other] = false;
      touched.push(other);
    }
  }
  if ((field === "superscript" || field === "subscript") && value === true && !next.offset.trim()) {
    next.offset = String(field === "superscript" ? OFFICE_CHART_SUPERSCRIPT_OFFSET : OFFICE_CHART_SUBSCRIPT_OFFSET);
    touched.push("offset");
  }
  if (field === "kerning" && value === true && !next.kerningPt.trim()) {
    next.kerningPt = String(OFFICE_CHART_DEFAULT_KERNING_PT);
    touched.push("kerningPt");
  }
  if (field === "spacing" && value !== "normal" && !next.spacingBy.trim()) {
    next.spacingBy = "1";
    touched.push("spacingBy");
  }
  return { draft: next, touched };
}

/** Why the draft cannot be written, as the field at fault, or null. */
export type OfficeChartFontDraftProblem = "size" | "offset" | "spacing" | "kerning";

function parsed(raw: string): number | null {
  const value = Number(raw.trim().replace(/(pt|%|磅)$/i, "").trim());
  return raw.trim() && Number.isFinite(value) ? value : null;
}

/**
 * The chart text OK writes: the chart's `current` text with the `touched`
 * fields of the draft written over it. `defaultFace` is the face the chart
 * draws with when it names none (an Asian face is stated beside a Latin one
 * only). `{ problem }` when a touched field is out of Office's range.
 */
export function officeChartTextFromFontDraft(
  current: OfficeChartText | undefined,
  draft: OfficeChartFontDraft,
  touched: ReadonlySet<OfficeChartFontDraftField>,
  defaultFace: string,
): { text: OfficeChartText | undefined } | { problem: OfficeChartFontDraftProblem } {
  const next: Record<string, unknown> = { ...(current ?? {}) };
  const face = (value: string | null) => (value ? primaryOfficeFontFace(value) || null : null);
  if (touched.has("family")) {
    const latin = face(draft.family);
    if (latin) next.fontFamily = latin;
    else delete next.fontFamily;
  }
  if (touched.has("family") || touched.has("asianFamily")) {
    const asian = face(draft.asianFamily);
    if (!asian) delete next.eastAsianFontFamily;
    else {
      if (!next.fontFamily) next.fontFamily = face(defaultFace) || "Calibri";
      if (asian === next.fontFamily) delete next.eastAsianFontFamily;
      else next.eastAsianFontFamily = asian;
    }
    if (!next.fontFamily) delete next.eastAsianFontFamily;
  }
  if (touched.has("bold")) next.bold = draft.bold;
  if (touched.has("italic")) next.italic = draft.italic;
  if (touched.has("sizePt")) {
    if (!draft.sizePt.trim()) delete next.fontSizePt;
    else {
      const size = parsed(draft.sizePt);
      if (size === null || size < OFFICE_CHART_TEXT_MIN_SIZE_PT || size > OFFICE_CHART_TEXT_MAX_SIZE_PT) {
        return { problem: "size" };
      }
      next.fontSizePt = Math.round(size * 2) / 2;
    }
  }
  if (touched.has("color")) {
    if (draft.color) next.color = draft.color.toLowerCase();
    else delete next.color;
  }
  if (touched.has("underline")) {
    if (draft.underline === "none") {
      delete next.underline;
      delete next.underlineColor;
    } else next.underline = PAGE_TO_UNDERLINE[draft.underline];
  }
  if (touched.has("underlineColor")) {
    if (draft.underlineColor && next.underline) next.underlineColor = draft.underlineColor.toLowerCase();
    else delete next.underlineColor;
  }
  if (touched.has("strike") || touched.has("doubleStrike")) {
    if (draft.doubleStrike) next.strike = "double";
    else if (draft.strike) next.strike = "single";
    else delete next.strike;
  }
  if (touched.has("superscript") || touched.has("subscript") || touched.has("offset")) {
    if (!draft.superscript && !draft.subscript) delete next.baseline;
    else {
      const offset = parsed(draft.offset);
      if (offset === null || offset <= 0 || offset > OFFICE_CHART_TEXT_MAX_BASELINE) return { problem: "offset" };
      next.baseline = draft.superscript ? offset : -offset;
    }
  }
  if (touched.has("smallCaps") || touched.has("allCaps")) {
    if (draft.allCaps) next.caps = "all";
    else if (draft.smallCaps) next.caps = "small";
    else delete next.caps;
  }
  if (touched.has("equalizeHeight")) {
    if (draft.equalizeHeight) next.equalizeHeight = true;
    else delete next.equalizeHeight;
  }
  if (touched.has("spacing") || touched.has("spacingBy")) {
    if (draft.spacing === "normal") delete next.spacingPt;
    else {
      const by = parsed(draft.spacingBy);
      if (by === null || by < 0 || by > OFFICE_CHART_TEXT_MAX_SPACING_PT) return { problem: "spacing" };
      if (by === 0) delete next.spacingPt;
      else next.spacingPt = draft.spacing === "expanded" ? by : -by;
    }
  }
  if (touched.has("kerning") || touched.has("kerningPt")) {
    if (!draft.kerning) delete next.kerningPt;
    else {
      const kerning = parsed(draft.kerningPt);
      if (kerning === null || kerning <= 0 || kerning > OFFICE_CHART_TEXT_MAX_KERNING_PT) return { problem: "kerning" };
      next.kerningPt = kerning;
    }
  }
  if (Object.keys(next).length === 0) return { text: undefined };
  const text = normalizeOfficeChartText(next);
  return text ? { text } : { problem: "size" };
}

/**
 * What the Font dialog opens with on one chart element (the chart area: all
 * of the chart's text): `text` is the element's own text, the one OK writes
 * the changed fields over; `shown` is what the element draws with, its own
 * over what it inherits from the chart, which the dialog's fields show.
 */
export function officeChartFontDialogText(
  attrs: Pick<ChartBlockAttrsV2, "format">,
  element: OfficeChartTextElement,
): { text: OfficeChartText | undefined; shown: OfficeChartText | undefined } {
  return {
    text: officeChartElementOwnText(attrs, element),
    shown: officeChartElementText(attrs, element),
  };
}

/**
 * The chart with the Font dialog's OK written on the element: its own text
 * replaced by the one OK shows (`undefined`: nothing stated). Null when that
 * changes nothing. Every editor's chart host commits this the way its Home
 * Font group commits the chart's text.
 */
export function officeChartWithFontDialogText(
  attrs: ChartBlockAttrsV2,
  element: OfficeChartTextElement,
  text: OfficeChartText | undefined,
): ChartBlockAttrsV2 | null {
  const own = officeChartElementOwnText(attrs, element);
  const next = applyOfficeChartTextCommand(own, { kind: "replace", text });
  if (next === null || JSON.stringify(next ?? null) === JSON.stringify(own ?? null)) return null;
  return officeChartWithElementText(attrs, element, next);
}
