/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfGhostRect = Readonly<{ left: number; top: number; right: number; bottom: number }>;

const area = (rect: PdfGhostRect) => Math.max(0, rect.right - rect.left) * Math.max(0, rect.bottom - rect.top);

const union = (rects: readonly PdfGhostRect[]): PdfGhostRect => rects.reduce((all, rect) => ({
  left: Math.min(all.left, rect.left), top: Math.min(all.top, rect.top),
  right: Math.max(all.right, rect.right), bottom: Math.max(all.bottom, rect.bottom),
}), rects[0]);

/** Whether `rect` reaches into any of `regions`. */
export function pdfGhostRectInRegions(rect: PdfGhostRect, regions: readonly PdfGhostRect[]): boolean {
  return area(rect) > 0 && regions.some((region) => rect.right > region.left && rect.left < region.right &&
    rect.bottom > region.top && rect.top < region.bottom);
}

/**
 * The regions of the page the ghost copies, and the boxes among `boxes`
 * that are the line's own (kept as copied, their dashed frames stood down);
 * every other box a region reaches is another line's and is left to the
 * page. `own` is the index of the box the line was opened from, or null.
 *
 * A box is under the frame when the frame covers half of it or more.
 */
export function pdfFinishGhostRegions(input: Readonly<{
  frame: PdfGhostRect;
  boxes: readonly PdfGhostRect[];
  own: number | null;
}>): Readonly<{ regions: readonly PdfGhostRect[]; ownBoxes: readonly number[]; movedOff: boolean }> {
  const { frame, boxes } = input;
  const under: number[] = [];
  boxes.forEach((box, index) => {
    const width = Math.min(box.right, frame.right) - Math.max(box.left, frame.left);
    const height = Math.min(box.bottom, frame.bottom) - Math.max(box.top, frame.top);
    if (width > 0 && height > 0 && width * height >= 0.5 * area(box)) under.push(index);
  });
  const own = input.own !== null && input.own >= 0 && input.own < boxes.length && area(boxes[input.own]) > 0
    ? input.own : null;
  if (own !== null && !under.includes(own)) {
    return Object.freeze({ regions: Object.freeze([frame, boxes[own]]), ownBoxes: Object.freeze([own]), movedOff: true });
  }
  return Object.freeze({
    regions: Object.freeze([union([frame, ...under.map((index) => boxes[index])])]),
    ownBoxes: Object.freeze(under),
    movedOff: false,
  });
}





















export function pdfFinishGhostDraftShows(input: Readonly<{
  candidateCopied: boolean;
  surfaceCopied: boolean;
  liftShown: boolean;
  surfaceCarriesLift: boolean;
  /** The line's old ink is hidden on the page itself (the typing cover,
   * carried with the ghost): the draft is the line's words there too. */
  coverCarried?: boolean;
}>): boolean {
  return (input.surfaceCopied || input.coverCarried === true) && !input.candidateCopied && !input.liftShown &&
    !input.surfaceCarriesLift;
}
