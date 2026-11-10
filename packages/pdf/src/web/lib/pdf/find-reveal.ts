/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface FindRevealRect {
  /** Percent of the page wrapper. */
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface FindRevealBox {
  top: number;
  left: number;
  width: number;
  height: number;
}

export type FindRevealPlan =
  | { kind: "visible" }
  | { kind: "scroll"; offset: number; offsetX: number };

export function planFindMatchReveal(input: {
  rects: readonly FindRevealRect[];
  /** The page wrapper's on-screen box. */
  pageRect: FindRevealBox;
  /** The page wrapper's layout size in reader CSS pixels. */
  pageSize: { width: number; height: number };
  /** The scroll container's on-screen box. */
  viewRect: FindRevealBox;
  /** The scroll container's client size in reader CSS pixels. */
  viewSize: { width: number; height: number };
}): FindRevealPlan | null {
  const { rects, pageRect, pageSize, viewRect, viewSize } = input;
  if (rects.length === 0 || !(pageRect.height > 0) || !(pageRect.width > 0)) {
    return null;
  }
  const top = Math.min(...rects.map((r) => r.y));
  const bottom = Math.max(...rects.map((r) => r.y + r.height));
  const left = Math.min(...rects.map((r) => r.x));
  const right = Math.max(...rects.map((r) => r.x + r.width));
  const screenTop = pageRect.top + (top / 100) * pageRect.height;
  const screenBottom = pageRect.top + (bottom / 100) * pageRect.height;
  const screenLeft = pageRect.left + (left / 100) * pageRect.width;
  const screenRight = pageRect.left + (right / 100) * pageRect.width;
  if (
    screenTop >= viewRect.top &&
    screenBottom <= viewRect.top + viewRect.height &&
    screenLeft >= viewRect.left &&
    screenRight <= viewRect.left + viewRect.width
  ) {
    return { kind: "visible" };
  }
  return {
    kind: "scroll",
    offset: Math.max(0, (top / 100) * pageSize.height - viewSize.height / 3),
    offsetX: Math.max(0, (left / 100) * pageSize.width - viewSize.width / 3),
  };
}
