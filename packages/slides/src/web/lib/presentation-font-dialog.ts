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
 * size, colour, single/double underline, strikethrough, superscript /
 * subscript, small / all caps, spacing); the rest of the page is shown
 * disabled by the host rather than taking a choice the export would drop.
 */

import type { RunStyle } from "@mrnote/core/web/lib/presentation-runs.ts";

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
export type PresentationFontDialogUnderline = "none" | "single" | "double";
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
      ? style.underlineStyle === "double" ? "double" : "single"
      : "none",
    strike: style.strike === true,
    superscript: style.superscript === true,
    subscript: style.subscript === true,
    caps: style.caps ?? "none",
    spacing: spacing > 0 ? "expanded" : spacing < 0 ? "condensed" : "normal",
    spacingBy: spacing === 0 ? "" : trimNumber(Math.abs(spacing)),
  };
}

/**
 * One control's change, with the fields it moves: superscript and subscript
 * exclude each other, and so do the two caps.
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
    runPatch.underlineStyle = draft.underline === "double" ? "double" : undefined;
  }
  if (touched.has("strike")) runPatch.strike = draft.strike || undefined;
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
