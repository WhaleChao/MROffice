/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Line WEIGHT — one unit, one step, one way to print it, for all four hosts.
 *
 * ## What Office says
 *
 * "Go to Format > Shape Outline, point to Weight, and then choose a
 * thickness… If you don't see the thickness you want, select More Lines, and
 * set the thickness you want in the Width box" (Microsoft Support, *Add or
 * remove a border on a text box, shape, or SmartArt graphic*). The submenu and
 * that Width box are in POINTS in Word, Excel and PowerPoint alike, and the
 * submenu's own entries are the nine below — the quarter-point ladder
 * ¼ / ½ / ¾ / 1 / 1½ / 2¼ / 3 / 4½ / 6 pt.
 *
 * The file format is not in points: `a:ln@w` is EMU (ECMA-376 §20.1.10.35,
 * "If this attribute is omitted, then a value of 0 is assumed"). 12 700 EMU
 * is one point.
 *
 * ## What MRAI had
 *
 * Three controls, three shapes, two units (SHAPE-IMPORT-AUDIT.md §57.7, §64):
 * Docs a point dropdown, the deck a PIXEL number field and a PIXEL menu
 * labelled `{value} pt`, the sheet a command whose values were pixels. A deck's
 * "6 pt" was 6 px = 4.5 pt, so the menu named a weight it did not write.
 *
 * This module is the one answer. A host that stores pixels keeps storing
 * pixels — the page contract is not a unit question — and converts AT THE
 * CONTROL, which is the only place a user reads a number.
 */

/**
 * PowerPoint's / Word's / Excel's own Shape Outline > Weight submenu, in
 * points and in its own order.
 */
export const OFFICE_LINE_WEIGHT_PT: readonly number[] = Object.freeze([
  0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6,
]);

/** The increment the Width box steps by: a quarter of a point. */
export const OFFICE_LINE_WEIGHT_STEP_PT = 0.25;

/** One point, in EMU (ECMA-376). The file's unit, for the writers. */
import { EMU_PER_PT as EMU_PER_POINT, EMU_PER_PX as EMU_PER_CSS_PIXEL } from "./office-units";
export { EMU_PER_POINT };

/** One CSS pixel, in EMU. The unit the deck and the sheet store weights in. */
export { EMU_PER_CSS_PIXEL };

/** A stored pixel weight as the points a control shows. 1 px = 0.75 pt. */
export function officeLineWeightPxToPt(px: number): number {
  return (px * EMU_PER_CSS_PIXEL) / EMU_PER_POINT;
}

/** A control's points as the pixels a host stores. 1 pt = 4/3 px. */
export function officeLineWeightPtToPx(pt: number): number {
  return (pt * EMU_PER_POINT) / EMU_PER_CSS_PIXEL;
}








export const OFFICE_OUTLINE_SHOWN_WEIGHT_PT = 0.75;

/**
 * The weight a colour or a dash picked for an outline leaves the line at, in
 * the host's own unit: its own weight, or `shown` when it has none. A colour
 * on a zero weight drew nothing, so the pick looked ignored.
 */
export function officeOutlineWeightWhenShown(current: number, shown: number): number {
  return Number.isFinite(current) && current > 0 ? current : shown;
}

/**
 * How a weight is PRINTED: two decimals, as the Width box shows it
 * ("0.25", "2.25", "3.75"). One formatter so a menu entry, a spinner and a
 * tooltip cannot disagree about the same number.
 */
export function formatOfficeLineWeightPt(pt: number): string {
  return (Number.isFinite(pt) ? pt : 0).toFixed(2);
}

/** The nearest quarter-point, for a value a user typed or a package stated. */
export function snapOfficeLineWeightPt(pt: number): number {
  if (!Number.isFinite(pt)) return 0;
  return (
    Math.round(Math.max(0, pt) / OFFICE_LINE_WEIGHT_STEP_PT) *
    OFFICE_LINE_WEIGHT_STEP_PT
  );
}

/**
 * A stored PIXEL weight as the text a Width box shows — the points it is, to
 * two decimals. The deck and the sheet store pixels; only the control speaks
 * points, so this is the one place the two meet.
 */
export function officeLineWeightPxAsPointText(px: number): string {
  return formatOfficeLineWeightPt(officeLineWeightPxToPt(px));
}

/**
 * The pixels to store for what a user typed into a POINT Width box: snapped
 * to Office's own quarter-point increment and bounded by the host's maximum.
 *
 * `null` means "that is not a weight" — blank, or not a number. The caller
 * then leaves the stored value exactly as it was instead of writing a number
 * the user never asked for.
 */
export function officeLineWeightPointTextToPx(
  text: string,
  maxPx: number,
): number | null {
  const trimmed = text.trim();
  if (trimmed === "") return null;
  const pt = Number(trimmed);
  if (!Number.isFinite(pt)) return null;
  return Math.min(maxPx, officeLineWeightPtToPx(snapOfficeLineWeightPt(pt)));
}

/**
 * One quarter-point up or down from a stored pixel weight, in pixels — what a
 * Width box's spinner does. The current weight is snapped to the ladder first,
 * so a file's odd 1.7 pt steps to 1.75 / 1.5 rather than 1.95 / 1.45.
 */
export function officeLineWeightSteppedPx(
  px: number,
  direction: 1 | -1,
  maxPx: number,
): number {
  const pt =
    snapOfficeLineWeightPt(officeLineWeightPxToPt(px)) +
    direction * OFFICE_LINE_WEIGHT_STEP_PT;
  return Math.min(maxPx, Math.max(0, officeLineWeightPtToPx(pt)));
}

/**
 * The same ladder in the pixels a host stores, bounded by that host's own
 * maximum. A weight the host cannot store is dropped rather than clamped: a
 * menu entry that silently writes a different number is the label lie this
 * module exists to end.
 */
export function officeLineWeightPxLadder(maxPx: number): readonly number[] {
  return Object.freeze(
    OFFICE_LINE_WEIGHT_PT.map(officeLineWeightPtToPx).filter(
      (px) => px <= maxPx + 1e-9,
    ),
  );
}
