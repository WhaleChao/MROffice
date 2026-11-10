/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { BoxRect } from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";
import { coverRectForEdit } from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";

export interface TextBoxAnchorRect extends BoxRect {
  originX?: number;
  originY?: number;
  originWidth?: number;
  originHeight?: number;
}

function rectArea(rect: BoxRect): number {
  return Math.max(0, rect.width) * Math.max(0, rect.height);
}

function intersectionArea(a: BoxRect, b: BoxRect): number {
  const left = Math.max(a.x, b.x);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const top = Math.max(a.y, b.y);
  const bottom = Math.min(a.y + a.height, b.y + b.height);
  return Math.max(0, right - left) * Math.max(0, bottom - top);
}

function center(rect: BoxRect): { x: number; y: number } {
  return {
    x: rect.x + rect.width / 2,
    y: rect.y + rect.height / 2,
  };
}

function anchorScore(edit: TextBoxAnchorRect, box: BoxRect): number {
  const anchor = coverRectForEdit(edit);
  const anchorArea = rectArea(anchor);
  const boxArea = rectArea(box);
  if (anchorArea <= 0 || boxArea <= 0) return -Infinity;
  const overlap =
    intersectionArea(anchor, box) / Math.max(0.000001, Math.min(anchorArea, boxArea));
  const anchorCenter = center(anchor);
  const boxCenter = center(box);
  const dx = Math.abs(anchorCenter.x - boxCenter.x);
  const dy = Math.abs(anchorCenter.y - boxCenter.y);
  return overlap - dx * 0.25 - dy * 0.75;
}

/**
 * Persisted PDF text-box edits are keyed by extraction ids such as `p18-1`.
 * Those ids can shift when the grouping algorithm improves. A stale edit should
 * never override a newly extracted box unless its recorded origin/current rect
 * still overlaps that same extracted box.
 */
export function textBoxEditAnchorsExtractedBox(
  edit: TextBoxAnchorRect,
  box: BoxRect,
): boolean {
  const anchor = coverRectForEdit(edit);
  const anchorArea = rectArea(anchor);
  const boxArea = rectArea(box);
  if (anchorArea <= 0 || boxArea <= 0) return false;

  const overlapRatio =
    intersectionArea(anchor, box) /
    Math.max(0.000001, Math.min(anchorArea, boxArea));
  if (overlapRatio < 0.35) return false;

  const anchorCenter = center(anchor);
  const boxCenter = center(box);
  const maxDx = Math.max(anchor.width, box.width, 0.01) * 0.75;
  const maxDy = Math.max(anchor.height, box.height, 0.01) * 1.1;
  return (
    Math.abs(anchorCenter.x - boxCenter.x) <= maxDx &&
    Math.abs(anchorCenter.y - boxCenter.y) <= maxDy
  );
}

export function findTextBoxEditAnchor<T extends BoxRect>(
  edit: TextBoxAnchorRect,
  boxes: readonly T[],
): T | null {
  let best: T | null = null;
  let bestScore = -Infinity;
  for (const box of boxes) {
    if (!textBoxEditAnchorsExtractedBox(edit, box)) continue;
    const score = anchorScore(edit, box);
    if (score > bestScore) {
      best = box;
      bestScore = score;
    }
  }
  return best;
}
