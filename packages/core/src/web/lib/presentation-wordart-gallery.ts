/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Insert > WordArt gallery and Shape Format > WordArt Styles
 * (PPT-A059): the twenty styles PowerPoint for Mac offers, in its 4 x 5
 * order (MsoPresetTextEffect 31-50), each as the paint PowerPoint itself
 * wrote for it (`presentation-wordart-gallery.json`, generated from the
 * server catalogue; data only, no markup).
 *
 * A style is a statement in theme colours (`office-text-paint-stated`): under
 * a deck's theme it draws in that theme's colours, as PowerPoint's gallery
 * does, and the server writes the statement back so PowerPoint reads the
 * same style. Its name is PowerPoint's own template ("Fill: ^0, Accent color
 * 1; Shadow") with the theme colours' names in place of ^0/^1/^2.
 */
import catalog from "./presentation-wordart-gallery.json";
import {
  normalizeOfficeTextPaint,
  type OfficeTextPaint,
} from "./office-text-paint";
import {
  normalizeOfficeTextPaintStated,
  officeTextPaintDisplayFromStated,
  officeTextPaintStatedUnderTheme,
  sameOfficeTextPaintStatement,
  type OfficeTextPaintStated,
} from "./office-text-paint-stated";

export interface PresentationWordArtStyle {
  /** `ppt-wordart-<item>`: the gallery tile's id. */
  readonly id: string;
  /** 1-20, the gallery position (row-major, five to a row). */
  readonly item: number;
  readonly paint: OfficeTextPaint;
  /** The style's own character format: bold, character spacing (points). */
  readonly bold?: true;
  readonly letterSpacingPt?: number;
}

interface CatalogItem {
  item: number;
  preset: number;
  run?: { bold?: boolean; letterSpacingPt?: number };
  stated: unknown;
}

const ITEMS = (catalog as unknown as { items: CatalogItem[] }).items.map((entry) => ({
  ...entry,
  stated: normalizeOfficeTextPaintStated(entry.stated)!,
}));

/** The gallery's number of styles. */
export const PRESENTATION_WORDART_STYLE_COUNT = ITEMS.length;

/**
 * The theme colours each style's name speaks of, in its template's order
 * (^0, ^1, ^2). Read off PowerPoint's own names (Shape Format > WordArt
 * Styles, accessibility dump): item 16's ^0 is Background 1, item 20's Text 2.
 */
const NAME_COLORS: Readonly<Record<number, readonly string[]>> = {
  1: ["tx1"],
  2: ["accent1"],
  3: ["accent2"],
  4: ["", "accent5"],
  5: ["accent4"],
  6: [],
  7: ["accent5"],
  8: ["accent4"],
  9: ["", "accent1"],
  10: ["accent3"],
  11: ["tx1", "bg1"],
  12: ["tx1", "bg1", "accent5"],
  13: ["accent5", "bg1"],
  14: ["", "accent2"],
  15: ["bg2"],
  16: ["bg1"],
  17: ["accent3"],
  18: ["accent1"],
  19: ["accent5"],
  20: ["tx2"],
};

/**
 * The twenty styles under a deck's theme (`theme`: the ten-colour grid,
 * `bg1`..`accent6`). `fallbackFill` is the colour text keeps under a style
 * that names no fill of its own (item 1: Text 1).
 */
export function presentationWordArtStyles(
  theme: Readonly<Record<string, string>>,
  fallbackFill?: string,
): readonly PresentationWordArtStyle[] {
  const ink = (fallbackFill ?? theme.tx1 ?? "#000000").toLowerCase();
  return ITEMS.map((entry) => {
    const bound = officeTextPaintStatedUnderTheme(entry.stated, theme);
    const stated: OfficeTextPaintStated = bound.fill ? bound : { ...bound, inheritedFill: ink };
    const paint = normalizeOfficeTextPaint({
      ...officeTextPaintDisplayFromStated(stated, ink),
      stated,
    })!;
    return {
      id: `ppt-wordart-${entry.item}`,
      item: entry.item,
      paint,
      ...(entry.run?.bold ? { bold: true as const } : {}),
      ...(entry.run?.letterSpacingPt ? { letterSpacingPt: entry.run.letterSpacingPt } : {}),
    };
  });
}

/** The gallery style a paint IS (its statement is the style's), or null. */
export function presentationWordArtStyleOf(paint: OfficeTextPaint | null | undefined): number | null {
  const stated = paint?.stated;
  if (!stated) return null;
  const entry = ITEMS.find((candidate) => sameOfficeTextPaintStatement(candidate.stated, stated));
  return entry ? entry.item : null;
}

/** The i18n key of a style's name template. */
export function presentationWordArtStyleNameKey(item: number): string {
  return `presentation.wordArtGallery.style.${item}`;
}

/**
 * A style's name: PowerPoint's template with the theme colours' names.
 * `t` is the console-notebooks translator.
 */
export function presentationWordArtStyleName(
  item: number,
  t: (key: string, values?: Record<string, string>) => string,
): string {
  const values: Record<string, string> = {};
  (NAME_COLORS[item] ?? []).forEach((slot, index) => {
    if (slot) values[`c${index}`] = t(`presentation.wordArtGallery.color.${slot}`);
  });
  return t(presentationWordArtStyleNameKey(item), values);
}

/**
 * A WordArt paint under a new deck theme: its statement's theme colours take
 * the theme's colours and it draws them, as PowerPoint recolours a theme
 * coloured style when the theme changes. A paint without a statement (plain
 * colours) stays as it is.
 */
export function presentationTextPaintUnderTheme(
  paint: OfficeTextPaint,
  theme: Readonly<Record<string, string>>,
): OfficeTextPaint {
  if (!paint.stated) return paint;
  const stated = officeTextPaintStatedUnderTheme(paint.stated, theme);
  try {
    return normalizeOfficeTextPaint({
      ...officeTextPaintDisplayFromStated(stated, paint.fill === "none" ? "#000000" : paint.fill),
      stated,
    })!;
  } catch {
    return paint;
  }
}
