/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Font dialog (Home > Font launcher, Character Spacing > More
 * Spacing...) over a slide's text, as pure draft arithmetic.
 *
 * The dialog itself is the shared Office Font page (`DocumentFontDialogFontPage`,
 * the page Docs and a chart's Font dialog already draw) plus PowerPoint's
 * Character Spacing page. This module decides what OK writes: a run-format
 * patch of exactly the fields the person changed (the same patch the Home
 * buttons commit) and, when the spacing page was changed, the text box's
 * character spacing in points.
 *
 * Only what the slide model keeps end to end is offered (Latin font, style,
 * size, colour, every DrawingML underline line, strikethrough and double
 * strikethrough, superscript / subscript, small / all caps, spacing); the
 * rest of the page is shown disabled by the host rather than taking a choice
 * the export would drop.
 */

import type { RunStyle } from "@mrnote/core/web/lib/presentation-runs.ts";
import {
  PRESENTATION_UNDERLINE_FONT_PAGE,
  PRESENTATION_UNDERLINE_FROM_FONT_PAGE,
  PRESENTATION_UNDERLINE_STYLES,
  normalizePresentationUnderlineStyle,
  type PresentationUnderlineStyle,
} from "@mrnote/core/web/lib/presentation-text-decoration.ts";

/** Home > Font > Character Spacing presets, in points (PowerPoint's values). */
export const PRESENTATION_CHARACTER_SPACING_PRESETS = [
  { id: "veryTight", points: -3 },
  { id: "tight", points: -1.5 },
  { id: "normal", points: 0 },
  { id: "loose", points: 3 },
  { id: "veryLoose", points: 6 },
] as const;

export type PresentationCharacterSpacingPresetId =
  (typeof PRESENTATION_CHARACTER_SPACING_PRESETS)[number]["id"];

/**
 * The largest tracking the deck's writer puts back as `a:rPr@spc` (the
 * server's `signed_tracking_points` bound). The dialog refuses more instead
 * of storing a value the export would silently drop.
 */
export const PRESENTATION_MAX_CHARACTER_SPACING_PT = 40;

/** The preset a stored spacing equals, if any (absent spacing is Normal). */
export function presentationCharacterSpacingPreset(
  points: number | undefined,
): PresentationCharacterSpacingPresetId | null {
  const value = points ?? 0;
  return (
    PRESENTATION_CHARACTER_SPACING_PRESETS.find((preset) => preset.points === value)?.id ?? null
  );
}

export type PresentationFontDialogSpacingKind = "normal" | "expanded" | "condensed";
/**
 * The Font page's underline names a slide's words keep: (none), Single and
 * every DrawingML line by the shared table (Word-spelled, as the page lists
 * them). Words only is not kept by the slide's runs.
 */
export type PresentationFontDialogUnderline =
  | "none"
  | "single"
  | (typeof PRESENTATION_UNDERLINE_FONT_PAGE)[PresentationUnderlineStyle];

/** The Font page's underline list for a slide's words, in PowerPoint's order. */
export const PRESENTATION_FONT_DIALOG_UNDERLINES: readonly PresentationFontDialogUnderline[] = [
  "none",
  // (none), Words only, Single, Double ... as PowerPoint lists them.
  PRESENTATION_UNDERLINE_FONT_PAGE.words,
  "single",
  ...PRESENTATION_UNDERLINE_STYLES.filter((style) => style !== "words").map((style) => PRESENTATION_UNDERLINE_FONT_PAGE[style]),
];

/** The run's line for a Font page name ("single" and "none" carry no style). */
export function presentationFontDialogUnderlineStyle(
  underline: PresentationFontDialogUnderline,
): PresentationUnderlineStyle | undefined {
  if (underline === "none" || underline === "single") return undefined;
  return normalizePresentationUnderlineStyle(PRESENTATION_UNDERLINE_FROM_FONT_PAGE[underline]);
}
export type PresentationFontDialogCaps = "none" | "small" | "all";

/** One field's drafted value; `null` where the selection is mixed and untouched. */
export interface PresentationFontDialogDraft {
  /** A font stack, "" for the inherited (theme) font, null mixed. */
  fontFamily: string | null;
  bold: boolean | null;
  italic: boolean | null;
  /** Points as typed; "" unset or mixed. */
  sizePt: string;
  /** #rrggbb, "" for Automatic, null mixed. */
  color: string | null;
  underline: PresentationFontDialogUnderline | null;
  strike: boolean;
  /** Double Strikethrough (`strike="dblStrike"`); excludes the single line. */
  doubleStrike: boolean;
  superscript: boolean;
  subscript: boolean;
  caps: PresentationFontDialogCaps;
  spacing: PresentationFontDialogSpacingKind;
  /** Points as typed, always non-negative; the kind gives the sign. */
  spacingBy: string;
}

export type PresentationFontDialogField = keyof PresentationFontDialogDraft;

/** What the selection carries now, as the ribbon reads it (absent = mixed/unset). */
export interface PresentationFontDialogSource {
  readonly style: Partial<RunStyle> | undefined;
  readonly fontFamily: string | null | undefined;
  readonly fontMixed: boolean;
  readonly sizePt: string;
  readonly letterSpacing: number | undefined;
}

function trimNumber(value: number): string {
  return String(Math.round(value * 100) / 100);
}

export function presentationFontDialogDraft(
  source: PresentationFontDialogSource,
): PresentationFontDialogDraft {
  const style = source.style ?? {};
  const spacing = source.letterSpacing ?? 0;
  return {
    fontFamily: source.fontMixed ? null : (source.fontFamily ?? ""),
    bold: style.bold === true,
    italic: style.italic === true,
    sizePt: source.sizePt,
    color: style.color ?? "",
    underline: style.underline === true
      ? (() => {
          const kind = normalizePresentationUnderlineStyle(style.underlineStyle);
          return kind ? PRESENTATION_UNDERLINE_FONT_PAGE[kind] : "single";
        })()
      : "none",
    strike: style.strike === true && style.strikeStyle !== "double",
    doubleStrike: style.strike === true && style.strikeStyle === "double",
    superscript: style.superscript === true,
    subscript: style.subscript === true,
    caps: style.caps ?? "none",
    spacing: spacing > 0 ? "expanded" : spacing < 0 ? "condensed" : "normal",
    spacingBy: spacing === 0 ? "" : trimNumber(Math.abs(spacing)),
  };
}

/**
 * One control's change, with the fields it moves: superscript and subscript
 * exclude each other, and so do the two caps and the two strikethroughs.
 */
export function changePresentationFontDialogDraft<K extends PresentationFontDialogField>(
  draft: PresentationFontDialogDraft,
  field: K,
  value: PresentationFontDialogDraft[K],
): { draft: PresentationFontDialogDraft; touched: PresentationFontDialogField[] } {
  const next = { ...draft, [field]: value } as PresentationFontDialogDraft;
  const touched: PresentationFontDialogField[] = [field];
  if (field === "superscript" && value === true) {
    next.subscript = false;
    touched.push("subscript");
  }
  if (field === "subscript" && value === true) {
    next.superscript = false;
    touched.push("superscript");
  }
  // Strikethrough and Double Strikethrough are one line or the other, as
  // PowerPoint's Font dialog checks them.
  if (field === "strike" && value === true) {
    next.doubleStrike = false;
    touched.push("doubleStrike");
  }
  if (field === "doubleStrike" && value === true) {
    next.strike = false;
    touched.push("strike");
  }
  if (field === "spacing" && value === "normal") {
    next.spacingBy = "";
    touched.push("spacingBy");
  }
  if (field === "spacing" && value !== "normal" && draft.spacingBy === "") {
    next.spacingBy = "1";
    touched.push("spacingBy");
  }
  if (field === "spacingBy") touched.push("spacing");
  return { draft: next, touched };
}

export type PresentationFontDialogProblem = "size" | "spacing";

export interface PresentationFontDialogResult {
  /** The run format OK writes over the selection; empty when no font field changed. */
  readonly runPatch: Partial<RunStyle>;
  /** Typed points for the size, when the size changed. */
  readonly sizePt?: number;
  /** The box's new character spacing in points (undefined = Normal), when it changed. */
  readonly letterSpacing?: { readonly points: number | undefined };
}

function parsePoints(raw: string): number | null {
  const trimmed = raw.trim().replace(/\s*(pt|磅)$/i, "");
  if (!trimmed) return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

/** What OK commits, or the first field whose value cannot be stored. */
export function presentationFontDialogResult(
  draft: PresentationFontDialogDraft,
  touched: ReadonlySet<PresentationFontDialogField>,
): PresentationFontDialogResult | { problem: PresentationFontDialogProblem } {
  const runPatch: Partial<RunStyle> = {};
  if (touched.has("fontFamily") && draft.fontFamily !== null) {
    runPatch.fontFamily = draft.fontFamily || undefined;
  }
  if (touched.has("bold") && draft.bold !== null) runPatch.bold = draft.bold || undefined;
  if (touched.has("italic") && draft.italic !== null) runPatch.italic = draft.italic || undefined;
  if (touched.has("color") && draft.color !== null) runPatch.color = draft.color || undefined;
  if (touched.has("underline") && draft.underline !== null) {
    runPatch.underline = draft.underline === "none" ? undefined : true;
    runPatch.underlineStyle = presentationFontDialogUnderlineStyle(draft.underline);
  }
  if (touched.has("strike") || touched.has("doubleStrike")) {
    runPatch.strike = draft.strike || draft.doubleStrike || undefined;
    runPatch.strikeStyle = draft.doubleStrike ? "double" : undefined;
  }
  if (touched.has("superscript") || touched.has("subscript")) {
    runPatch.superscript = draft.superscript || undefined;
    runPatch.subscript = draft.subscript || undefined;
  }
  if (touched.has("caps")) runPatch.caps = draft.caps === "none" ? undefined : draft.caps;

  let sizePt: number | undefined;
  if (touched.has("sizePt")) {
    const value = parsePoints(draft.sizePt);
    if (value === null || value < 1 || value > 400) return { problem: "size" };
    sizePt = Math.round(value * 10) / 10;
  }

  let letterSpacing: PresentationFontDialogResult["letterSpacing"];
  if (touched.has("spacing") || touched.has("spacingBy")) {
    if (draft.spacing === "normal") {
      letterSpacing = { points: undefined };
    } else {
      const by = parsePoints(draft.spacingBy);
      if (by === null || by < 0 || by > PRESENTATION_MAX_CHARACTER_SPACING_PT) {
        return { problem: "spacing" };
      }
      const rounded = Math.round(by * 100) / 100;
      letterSpacing = {
        points: rounded === 0 ? undefined : draft.spacing === "condensed" ? -rounded : rounded,
      };
    }
  }
  return {
    runPatch,
    ...(sizePt === undefined ? {} : { sizePt }),
    ...(letterSpacing === undefined ? {} : { letterSpacing }),
  };
}
