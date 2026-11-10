/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PDF page's live ink in the room's presence (lib/collab/ink-presence), so
 * the others see a stroke as it is drawn and until it lands, not only once it
 * is saved.
 *
 * Points travel in tenths of a page percent (the presence format keeps a
 * quarter of a unit: a 0.025% step, a fifth of a pixel on an 800 px page),
 * the same on every member's screen whatever their zoom. Each ink tool has
 * its own surfaces, as each keeps its own strokes and widths: the reader's
 * Draw tab (`reader`: widths as `inkWidthPagePercent` reads them) and the
 * Edit PDF ink layer (`edit`: its own stroke widths).
 */
import type { InkPresenceStroke } from "@mrnote/core/web/lib/collab/ink-presence.ts";
import type { InkPoint } from "@mrnote/core/web/lib/pdf/ink.ts";

/** Presence units per page percent. */
const UNITS_PER_PERCENT = 10;

export type PdfInkPresenceTool = "reader" | "edit";

/** The presence surface of a page (1-based for the reader, as its strokes). */
export function pdfInkPresenceSurface(tool: PdfInkPresenceTool, page: number): string {
  return `pdf-${tool}:${page}`;
}

/** A stroke in page percent (0..100) as presence carries it; `alpha` 1 is a pen. */
export function pdfInkPresenceStroke(
  points: readonly InkPoint[],
  color: string,
  width: number,
  alpha = 1,
): InkPresenceStroke {
  return {
    color,
    // In the same tenths, so a 0.25 reader width keeps its hundredths.
    width: width * UNITS_PER_PERCENT,
    alpha,
    points: points.map((point) => ({
      x: point.x * UNITS_PER_PERCENT,
      y: point.y * UNITS_PER_PERCENT,
      ...(point.p !== undefined ? { p: point.p } : {}),
    })),
  };
}

/** A peer's stroke width back in the tool's own units. */
export function pdfInkWidthFromPresence(stroke: InkPresenceStroke): number {
  return stroke.width / UNITS_PER_PERCENT;
}

/** A peer's live stroke back in page percent. */
export function pdfInkPointsFromPresence(stroke: InkPresenceStroke): InkPoint[] {
  return stroke.points.map((point) => ({
    x: Math.min(100, Math.max(0, point.x / UNITS_PER_PERCENT)),
    y: Math.min(100, Math.max(0, point.y / UNITS_PER_PERCENT)),
    ...(point.p !== undefined ? { p: point.p } : {}),
  }));
}
