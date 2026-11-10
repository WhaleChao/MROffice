/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pxToPt } from "./office-units";
import type { OfficeShapeStyleRecipe } from "./office-shape-style-catalog";
import type { DocumentShapeAttrs } from "./document-shapes";
export function documentShapeStylePatch(r: OfficeShapeStyleRecipe) {
  return {
    fill: r.fill,
    outline: r.stroke,
    outlineWidthPt: pxToPt(r.strokeWidth),
    outlineStyle: "solid" as const,
    fillGradient: r.fillGradient ?? null,
    // A style's fill replaces a picture fill, as Office's Shape Styles do.
    fillPicture: null,
    shapeShadow: r.shapeShadow ?? null,
    effect: "none" as const,
    textFill: r.textColor,
  };
}
export function documentShapeMatchesStyle(
  a: DocumentShapeAttrs,
  r: OfficeShapeStyleRecipe,
) {
  const patch = documentShapeStylePatch(r);
  return Object.entries(patch).every(
    ([key, value]) =>
      JSON.stringify(a[key as keyof DocumentShapeAttrs] ?? null) ===
      JSON.stringify(value),
  );
}
