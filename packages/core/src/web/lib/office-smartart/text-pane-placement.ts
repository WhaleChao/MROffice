/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where the SmartArt Text Pane stands: beside the diagram it edits, as
 * Office's Text Pane does (Microsoft Support, "Use the Text pane to edit text
 * in a SmartArt graphic": the pane opens to the left of the graphic's frame).
 *
 * Left of the frame when the canvas has room there, else right of it, else
 * against the edge of the roomier side; level with the frame's top, never
 * past the canvas's top or bottom. Every host (Docs, Slides, Sheets) asks
 * this one function, so the pane stands in the same place in each.
 */

export interface PaneRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface SmartArtTextPanePlacement {
  left: number;
  top: number;
  maxHeight: number;
  side: "left" | "right" | "inside";
}

/** Between the pane and the frame: clear of the frame's Text Pane control. */
export const SMARTART_TEXT_PANE_GAP = 24;
/** Between the pane and the canvas's edge. */
export const SMARTART_TEXT_PANE_MARGIN = 8;

export function smartArtTextPanePlacement({
  anchor,
  bounds,
  pane,
}: {
  anchor: PaneRect;
  bounds: PaneRect;
  pane: { width: number; height: number };
}): SmartArtTextPanePlacement {
  const gap = SMARTART_TEXT_PANE_GAP;
  const margin = SMARTART_TEXT_PANE_MARGIN;
  const maxHeight = Math.max(120, bounds.bottom - bounds.top - margin * 2);
  const height = Math.min(pane.height, maxHeight);
  let left = anchor.left - gap - pane.width;
  let side: SmartArtTextPanePlacement["side"] = "left";
  if (left < bounds.left + margin) {
    left = anchor.right + gap;
    side = "right";
    if (left + pane.width > bounds.right - margin) {
      // Neither side has room: the side with more, against the canvas's
      // edge, so the pane covers as little of the graphic as it can.
      side = "inside";
      left =
        anchor.left - bounds.left >= bounds.right - anchor.right
          ? bounds.left + margin
          : bounds.right - margin - pane.width;
    }
  }
  const top = Math.max(
    bounds.top + margin,
    Math.min(anchor.top, bounds.bottom - margin - height),
  );
  return { left, top, maxHeight, side };
}

/** The part of `rect` inside `clip` (the window, for a canvas taller than it). */
export function intersectPaneRect(rect: PaneRect, clip: PaneRect): PaneRect {
  return {
    left: Math.max(rect.left, clip.left),
    top: Math.max(rect.top, clip.top),
    right: Math.min(rect.right, clip.right),
    bottom: Math.min(rect.bottom, clip.bottom),
  };
}

/** The nearest ancestor that scrolls: the canvas a diagram is drawn on. */
export function scrollingAncestor(element: Element): Element | null {
  for (let node = element.parentElement; node; node = node.parentElement) {
    const style = getComputedStyle(node);
    if (/(auto|scroll)/.test(`${style.overflowY} ${style.overflowX}`)) return node;
  }
  return null;
}

/**
 * The first element matching `selector` in the nearest ancestor of `from`
 * that holds one: the frame in the pane's own editor window, not another
 * window's.
 */
export function nearestMatch(from: Element, selector: string): Element | null {
  for (let node = from.parentElement; node; node = node.parentElement) {
    const hit = node.querySelector(selector);
    if (hit) return hit;
  }
  return null;
}
