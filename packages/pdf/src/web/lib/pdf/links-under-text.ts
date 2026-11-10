/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PageFractionBox {
  /** Page fractions (0..1) of the shown page, top-left origin. */
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface PagePercentLink {
  readonly id: string;
  /** 1-based page number. */
  readonly page: number;
  /** Page percentages (0..100) of the shown page, top-left origin. */
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly deleted?: boolean;
}

/** Ids of the page's links whose area is mostly on `box` (at least `share`). */
export function linkIdsOnText(
  links: readonly PagePercentLink[],
  page: number,
  box: PageFractionBox,
  share = 0.6,
): string[] {
  const left = box.x * 100;
  const top = box.y * 100;
  const right = (box.x + box.width) * 100;
  const bottom = (box.y + box.height) * 100;
  if (!(right > left) || !(bottom > top)) return [];
  const ids: string[] = [];
  for (const link of links) {
    if (link.deleted || link.page !== page) continue;
    const area = link.width * link.height;
    if (!(area > 0)) continue;
    const overlapWidth = Math.min(link.x + link.width, right) - Math.max(link.x, left);
    const overlapHeight = Math.min(link.y + link.height, bottom) - Math.max(link.y, top);
    if (overlapWidth <= 0 || overlapHeight <= 0) continue;
    if ((overlapWidth * overlapHeight) / area >= share) ids.push(link.id);
  }
  return ids;
}
