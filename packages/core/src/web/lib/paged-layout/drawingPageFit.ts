/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Unit-agnostic geometry for an inline drawing that must fit one Word page's
 * content box. It is unit-agnostic so an exporter can deliberately apply the
 * same rule in EMUs once its fidelity policy supplies the correct section box.
 * The current editor caller uses CSS pixels; merely importing this helper does
 * not authorize an exporter to rewrite an imported Office drawing extent.
 *
 * This is a projection rule: callers keep the authored/original geometry and
 * apply the returned size only to the page projection or OOXML extent.
 */

export interface DrawingSize {
  width: number;
  height: number;
}

export interface DrawingPageFitResult extends DrawingSize {
  scale: number;
  fitted: boolean;
}

/**
 * Only a non-anchored inline image may use this projection. Unknown non-empty
 * values fail closed: an imported future Word wrap mode must never be mistaken
 * for an ordinary inline picture merely because today's normalizer ignores it.
 */
export function isInlineImagePageFitEligible(
  dataFloat: unknown,
  officeDrawingKind?: unknown,
): boolean {
  const kindIsAbsent =
    officeDrawingKind === null ||
    officeDrawingKind === undefined ||
    officeDrawingKind === "";
  return (
    (dataFloat === null || dataFloat === undefined || dataFloat === "") &&
    (kindIsAbsent || officeDrawingKind === "inline")
  );
}

function positiveFinite(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

/**
 * Fit one drawing inside both axes of a page content box without cropping or
 * enlargement. Returns null for incomplete/invalid geometry so a caller can
 * wait for intrinsic image metadata instead of inventing a ratio.
 */
export function fitDrawingWithinPageBox(
  drawing: DrawingSize,
  pageBox: DrawingSize,
): DrawingPageFitResult | null {
  if (
    !positiveFinite(drawing.width) ||
    !positiveFinite(drawing.height) ||
    !positiveFinite(pageBox.width) ||
    !positiveFinite(pageBox.height)
  ) {
    return null;
  }

  const scale = Math.min(
    1,
    pageBox.width / drawing.width,
    pageBox.height / drawing.height,
  );
  return {
    width: drawing.width * scale,
    height: drawing.height * scale,
    scale,
    fitted: scale < 1,
  };
}
