/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The document's own definition of Word's built-in `Heading 1`..`Heading 6`.
 *
 * Word resolves a paragraph through `w:docDefaults`, then the `w:basedOn`
 * style chain, then direct formatting (ECMA-376 §17.7.5.1: document defaults
 * "are applied first in the style hierarchy; therefore they are superseded by
 * any further conflicting formatting"). A generated type scale stamped onto
 * every heading RUN therefore looked right and was inert — editing "Heading 2"
 * in the Styles gallery changed nothing, because the direct formatting won.
 *
 * Stating the scale as the style DEFINITION buys that linkage while keeping the
 * heading semantics a custom style would lose: `w:pStyle w:val="Heading2"` is
 * what resolves `w:outlineLvl` through the cascade, and that is what the
 * navigation pane and an automatic table of contents read.
 *
 * This module is the single contract shared by the editor (which renders the
 * ladder), the offline DOCX serializer and the server writer, so preview and
 * both exports can never drift.
 */

import { ptToPx } from "./office-units";
import { normalizeLineHeight } from "../components/console/editor/extensions/DocumentFormatting";
import {
  normalizeFontFamily,
  normalizeTextColor,
} from "../components/console/editor/extensions/TextStyling";

/**
 * One redefined heading level.
 *
 * Units are Word's own: points at half-point granularity for the run size
 * (OOXML stores `w:sz` in half-points, so a quarter-point value cannot reopen)
 * and a unitless ratio for the line height (`w:spacing w:line` in 240ths with
 * `w:lineRule="auto"`).
 */
export interface HeadingStyleDefinition {
  fontFamily: string;
  fontSizePt: number;
  textColor: string;
  spaceBeforePt: number;
  spaceAfterPt: number;
  lineHeight: string;
  bold?: boolean;
  italic?: boolean;
}

export type HeadingStyleMap = Record<string, HeadingStyleDefinition>;

/** The six outline levels a document may redefine. */
export const HEADING_STYLE_LEVELS = ["1", "2", "3", "4", "5", "6"] as const;

const MAX_HEADING_SPACING_PT = 1584;

/** Bounded, two-decimal paragraph spacing in points; `null` when off-range. */
function headingSpacingPt(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  if (value < 0 || value > MAX_HEADING_SPACING_PT) return null;
  // Word stores the value in twips; anything finer is not representable.
  return Math.round(value * 20) === value * 20 ? value : null;
}

/**
 * Whitelist-validate a COMPLETE `Heading 1`..`Heading 6` redefinition.
 *
 * All six levels must be present and each must carry the full modeled
 * typography. A half-specified ladder would silently leave Word's own
 * theme-derived colour, font scheme and size standing for the rest — neither
 * the document's design nor readable back as a literal definition — so it is
 * rejected outright rather than partially honoured. Returns `null` for a
 * document that carries no ladder of its own, which is the normal case for an
 * imported file whose headings are defined against its theme.
 */
export function normalizeHeadingStyles(value: unknown): HeadingStyleMap | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const source = value as Record<string, unknown>;
  if (Object.keys(source).length !== HEADING_STYLE_LEVELS.length) return null;
  const result: HeadingStyleMap = {};
  for (const level of HEADING_STYLE_LEVELS) {
    const raw = source[level];
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const entry = raw as Record<string, unknown>;
    const fontFamily = normalizeFontFamily(entry.fontFamily);
    const textColor = normalizeTextColor(entry.textColor);
    const lineHeight = normalizeLineHeight(entry.lineHeight);
    const spaceBeforePt = headingSpacingPt(entry.spaceBeforePt);
    const spaceAfterPt = headingSpacingPt(entry.spaceAfterPt);
    const size = entry.fontSizePt;
    if (
      !fontFamily ||
      !textColor ||
      !lineHeight ||
      spaceBeforePt === null ||
      spaceAfterPt === null ||
      typeof size !== "number" ||
      !Number.isFinite(size) ||
      size < 4.5 ||
      size > 72 ||
      size * 2 !== Math.round(size * 2)
    ) {
      return null;
    }
    const definition: HeadingStyleDefinition = {
      fontFamily,
      fontSizePt: size,
      textColor,
      spaceBeforePt,
      spaceAfterPt,
      lineHeight,
    };
    if (entry.bold === true) definition.bold = true;
    if (entry.italic === true) definition.italic = true;
    result[level] = definition;
  }
  return result;
}

/** A point size as the canonical px token the document model persists. */
function pointsToCssPx(points: number): string {
  return (ptToPx(points)).toFixed(4).replace(/\.?0+$/, "");
}

/**
 * The CSS a heading paragraph of this level inherits from the definition.
 *
 * The editor paints exactly the properties the two DOCX writers emit onto the
 * `Heading N` style, so what the author sees is what Word resolves. Nothing is
 * emitted for an absent definition: the stylesheet's own heading rules stand,
 * which is how every document written before this contract keeps rendering.
 */
export function headingStyleCss(
  styles: HeadingStyleMap | null,
  level: unknown,
): Record<string, string> {
  const definition =
    styles && Number.isInteger(level) ? styles[String(level)] : undefined;
  if (!definition) return {};
  return {
    "font-family": definition.fontFamily,
    // The same px spelling the document model itself persists for a point
    // size (four decimals, trailing zeros stripped), so the rendered value and
    // an authored `fontSize` mark are directly comparable.
    "font-size": `${pointsToCssPx(definition.fontSizePt)}px`,
    "font-weight": definition.bold ? "700" : "400",
    "font-style": definition.italic ? "italic" : "normal",
    color: definition.textColor,
    "margin-top": `${definition.spaceBeforePt}pt`,
    "margin-bottom": `${definition.spaceAfterPt}pt`,
    "line-height": definition.lineHeight,
  };
}

/** The same declarations as one inline `style` attribute value. */
export function headingStyleInlineCss(
  styles: HeadingStyleMap | null,
  level: unknown,
): string {
  return Object.entries(headingStyleCss(styles, level))
    .map(([property, value]) => `${property}: ${value}`)
    .join("; ");
}
