/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { cssFontFamilyList } from "@mrnote/core/web/lib/css-font-family.ts";
import { pageTextWidth } from "@mrnote/core/web/lib/text-engine/page-adapter.ts";
let freeTextMeasureContext: CanvasRenderingContext2D | null | undefined;

/** One line's width in points, in the family the overlay draws the box in. */
export function measureFreeTextLineWidth(
  text: string,
  font: { fontFamily: string; bold: boolean; italic: boolean },
  fontSizePoints: number,
): number {
  if (freeTextMeasureContext === undefined) {
    freeTextMeasureContext =
      typeof document === "undefined"
        ? null
        : document.createElement("canvas").getContext("2d");
  }
  const css = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${fontSizePoints}px ${cssFontFamilyList(font.fontFamily)}`;
  // The served face's own advances (lib/text-engine), the same in every
  // browser, where the engine knows the faces; else the canvas.
  const engine = pageTextWidth("pdf", text, css);
  if (engine !== undefined) return engine;
  const context = freeTextMeasureContext;
  if (!context) return Number.NaN;
  context.font = css;
  return context.measureText(text).width;
}
