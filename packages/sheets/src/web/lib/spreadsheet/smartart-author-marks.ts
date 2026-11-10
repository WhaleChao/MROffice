/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A worksheet SmartArt shape formatted on its own keeps that formatting as the
 * AUTHOR's -- the sheet twin of the deck's `withSmartArtNodeAuthority`
 * (`lib/presentation-shape-buttons.ts`): Shape Fill / Outline / a Shape Styles
 * pick, or the Format Painter, on ONE shape of a graphic is that shape's own
 * customization (Office keeps it on the item's point), so a relayout (Add
 * Shape, Change Colors' re-spread) carries it, the export writes it, and Reset
 * Shape / Reset Graphic discard it (Microsoft Support, "Restore the default
 * layout and color of your SmartArt graphic").
 *
 * Markers, as the sheet already spells them:
 * - `smartArtAuthorPaint` -- the node's fill or line is the author's (only a
 *   `node` has paint of its own to claim);
 * - `smartArtAuthorText` -- the words' colour is the author's;
 * - `smartArtAuthorTextSize` -- the words' size is the author's (Office's
 *   `custT`), the same rule `lib/spreadsheet/shape-text-font.ts` applies.
 */

import { readSmartArtBlockMarker } from "@mrnote/core/web/lib/office-smartart/model.ts";

const PAINT_KEYS = [
  "fill",
  "fillGradient",
  "fillPattern",
  "fillTransparency",
  "fillPicture",
  "stroke",
  "strokeWidth",
  "outlineGradient",
  "outlineTransparency",
] as const;

function shapeOf(object: unknown): Record<string, unknown> {
  const shape = (object as { shape?: unknown } | null)?.shape;
  return shape && typeof shape === "object" && !Array.isArray(shape) ? (shape as Record<string, unknown>) : {};
}

const differs = (left: unknown, right: unknown) => JSON.stringify(left ?? null) !== JSON.stringify(right ?? null);

/**
 * `next` (the object after one format write) with the author markers its
 * change earns, judged against `previous`. Anything not a member of a
 * SmartArt graphic, or a change that touches none of the marked fields, comes
 * back as it was.
 */
export function withSheetSmartArtAuthorMarks<T extends object>(previous: object, next: T): T {
  const marker = readSmartArtBlockMarker(next);
  if (!marker || (marker.role !== "node" && marker.role !== "label")) return next;
  const before = shapeOf(previous);
  const after = shapeOf(next);
  const marks: Record<string, true> = {};
  if (marker.role === "node" && PAINT_KEYS.some((key) => differs(before[key], after[key]))) {
    marks.smartArtAuthorPaint = true;
  }
  if (differs(before.textColor, after.textColor)) marks.smartArtAuthorText = true;
  if (differs(before.textSize, after.textSize)) marks.smartArtAuthorTextSize = true;
  return Object.keys(marks).length ? ({ ...next, ...marks } as T) : next;
}
