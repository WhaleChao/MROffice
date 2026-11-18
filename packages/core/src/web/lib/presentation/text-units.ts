/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The one place a presentation's text sizes change unit.
//
// A slide's words are sized in points (DrawingML `a:rPr@sz`, the Font Size
// box, every download). The page states a block's sizes in the block's own
// unit: points for an imported box and any box marked
// `source.textFontSizeUnit: "point"`, design px (96 dpi, 4/3 of a point) for
// a box the editor wrote before point-unit boxes existed. Every reader asks
// this module for points (`storedFontSizeToPoints`) and every writer hands it
// points (`pointsToStoredFontSize`); only the canvas turns points into CSS px

// 42.67, 18 pt grown to 18.67) were each a reader or writer converting by
// hand, once too often or not at all.
//
// `Points` and `DesignPx` are branded: a size read in one unit cannot be
// passed where the other is meant without going through a conversion here,
// and tests/unit/presentation-text-units-guard.test.ts refuses a hand-written
// conversion (`* 0.75`, `/ 0.75`, `4 / 3`, `96 / 72`) anywhere else in the
// presentation sources.
//
// This module imports only office-units.ts (no imports of its own), so a
// module-level constant anywhere may be derived from it; which unit a BLOCK
// states is text-units-block.ts.
import { PX_PER_PT } from "../office-units";

declare const unit: unique symbol;
/** A length in typographic points (1/72 in). */
export type Points = number & { readonly [unit]: "pt" };
/** A length in the slide's design px (1/96 in; 960 px across a 10 in slide). */
export type DesignPx = number & { readonly [unit]: "design-px" };

/** The unit a block states its text sizes in. */
export type PresentationFontSizeUnit = "design-px" | "point";

/** CSS px per point (`office-units.ts`, every editor's one source). */
export const CSS_PX_PER_POINT = PX_PER_PT;

/**
 * A free text box's default size: PowerPoint's 18 pt (the presentation's
 * default text style), in the editor's design px (pt x 4/3). A new text box
 * states it (`NEW_TEXT_BOX_FONT_SIZE_PX`, PPT-G011); a box that states no
 * size is drawn at the canvas's 17 px fallback while the PDF, the PPTX and
 * PowerPoint draw 18 pt. Stated here, where a module-level constant anywhere
 * may be derived from it (presentation-shape-buttons.ts does), not in
 * presentation-text-buttons.ts, inside whose import cycle that module runs.
 */
export const PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PX = 24;
/** The same 18 pt, as a box marked `textFontSizeUnit: "point"` states it. */
export const PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PT = 18;

/** A number known to be points. */
export const asPoints = (value: number): Points => value as Points;
/** A number known to be design px. */
export const asDesignPx = (value: number): DesignPx => value as DesignPx;

const hundredths = (value: number) => Math.round(value * 100) / 100;

/** Points as design px (hundredths: 10.5 pt is 14 px, 32 pt 42.67). */
export function pointsToDesignPx(points: Points | number): DesignPx {
  return asDesignPx(hundredths(points * CSS_PX_PER_POINT));
}

/** Design px as points (hundredths). */
export function designPxToPoints(px: DesignPx | number): Points {
  return asPoints(hundredths(px / CSS_PX_PER_POINT));
}

/** A line's weight in points as design px, exact: a quarter-point edge is a
 * third of a pixel, not 0.33 (Office's nine weights round-trip). */
export function lineWeightPointsToDesignPx(points: Points | number): DesignPx {
  return asDesignPx(points * CSS_PX_PER_POINT);
}

/** A line's weight in design px as points, exact. */
export function lineWeightDesignPxToPoints(px: DesignPx | number): Points {
  return asPoints(px / CSS_PX_PER_POINT);
}

/** Points as the CSS px the canvas draws them at (the slide is laid out in design px). */
export function pointsToCssPx(points: Points | number): number {
  return pointsToDesignPx(points);
}

/** A size as stored in a block of `unit`, read as points. */
export function storedFontSizeToPoints(stored: number, unit: PresentationFontSizeUnit): Points {
  return unit === "point" ? asPoints(stored) : designPxToPoints(stored);
}

/** Points, as a block of `unit` stores them. */
export function pointsToStoredFontSize(points: Points, unit: PresentationFontSizeUnit): number {
  return unit === "point" ? points : pointsToDesignPx(points);
}

/** A size as stored in a block of `unit`, as the CSS px the canvas draws. */
export function storedFontSizeToCssPx(stored: number, unit: PresentationFontSizeUnit): number {
  return unit === "point" ? pointsToCssPx(stored) : stored;
}
