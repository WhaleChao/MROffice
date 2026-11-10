/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Freeze Panes and the worksheet's drawing layer.
 *
 * A frozen worksheet is four panes onto one sheet: the corner (frozen rows
 * and columns), the top band (frozen rows), the left band (frozen columns)
 * and the body, which alone scrolls on both axes. An object belongs to the
 * pane it lies in: one inside the frozen rows stays put while the body
 * scrolls under it, and one in the body passes UNDER the frozen cells, never
 * over them.
 *
 * An object that straddles a freeze line is cut in two there, one piece per
 * pane, as Excel shows it: it is drawn once in each pane it reaches, and each
 * pane shows only its own part of the sheet.
 */

export type SheetDrawingPane = "corner" | "top" | "left" | "body";

export const SHEET_DRAWING_PANES: readonly SheetDrawingPane[] = ["body", "top", "left", "corner"];

/**
 * The frozen bands' far edges in content pixels, headings included; 0 on an
 * axis that is not frozen.
 */
export interface SheetFreezeBand {
  readonly bottom: number;
  readonly right: number;
}

export interface SheetDrawingBox {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

const EDGE_TOLERANCE_PX = 0.5;

/** The sides of a frozen edge a span reaches (both when it crosses it). */
function sidesOfEdge(start: number, size: number, edge: number): ("frozen" | "free")[] {
  if (edge <= 0) return ["free"];
  if (start + size <= edge + EDGE_TOLERANCE_PX) return ["frozen"];
  if (start >= edge - EDGE_TOLERANCE_PX) return ["free"];
  return ["frozen", "free"];
}

/**
 * The panes an object is drawn in: one for an object inside a pane, up to
 * four for one over the corner of the freeze lines; none without frozen panes.
 */
export function sheetDrawingPanes(
  box: SheetDrawingBox,
  band: SheetFreezeBand,
): SheetDrawingPane[] {
  if (band.bottom <= 0 && band.right <= 0) return [];
  const panes: SheetDrawingPane[] = [];
  for (const rows of sidesOfEdge(box.top, box.height, band.bottom)) {
    for (const cols of sidesOfEdge(box.left, box.width, band.right)) {
      panes.push(
        rows === "frozen" ? (cols === "frozen" ? "corner" : "top") : cols === "frozen" ? "left" : "body",
      );
    }
  }
  return panes;
}

/**
 * Where a pointer's pane begins on screen, per axis: a frozen axis is
 * measured from the scrollport (its cells do not scroll), a free one from the
 * scrolled stage. A point is in the frozen rows while it is above the band's
 * bottom edge on screen, whatever the body has scrolled to.
 */
export function sheetDrawingPaneOrigin(
  point: { readonly x: number; readonly y: number },
  stage: { readonly left: number; readonly top: number },
  scrollport: { readonly left: number; readonly top: number },
  band: SheetFreezeBand,
  screenPerContent: { readonly x: number; readonly y: number },
): { left: number; top: number } {
  const frozenRows =
    band.bottom > 0 && point.y < scrollport.top + band.bottom * screenPerContent.y;
  const frozenCols =
    band.right > 0 && point.x < scrollport.left + band.right * screenPerContent.x;
  return {
    left: frozenCols ? scrollport.left : stage.left,
    top: frozenRows ? scrollport.top : stage.top,
  };
}
