/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Convert a viewport-coord client rect into a wrapper-relative percentage
 * rect that is guaranteed to satisfy the backend's PDF highlight bounds:
 *
 *   x, y, width, height in [0, 100]
 *   width > 0 and height > 0
 *   x + width <= 100
 *   y + height <= 100
 *
 * See `apps/api/app/routers/datasets.py` `_validate_annotation_payload`.
 *
 * Returns `null` when the rect is fully outside the wrapper or the wrapper
 * has zero size (avoids `Infinity` / `NaN` percentages).
 */

export interface BoundingRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface NormalizedRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

const MIN_DIMENSION = 0.01;

export function normalizeHighlightRect(
  client: BoundingRect,
  wrapper: BoundingRect,
): NormalizedRect | null {
  if (wrapper.width <= 0 || wrapper.height <= 0) return null;
  if (client.width <= 0 || client.height <= 0) return null;

  const rawX = ((client.left - wrapper.left) / wrapper.width) * 100;
  const rawY = ((client.top - wrapper.top) / wrapper.height) * 100;
  const rawWidth = (client.width / wrapper.width) * 100;
  const rawHeight = (client.height / wrapper.height) * 100;

  if (rawX >= 100 || rawY >= 100) return null;
  if (rawX + rawWidth <= 0 || rawY + rawHeight <= 0) return null;

  const x = Math.max(0, Math.min(100 - MIN_DIMENSION, rawX));
  const y = Math.max(0, Math.min(100 - MIN_DIMENSION, rawY));

  const widthCap = 100 - x;
  const heightCap = 100 - y;

  let width =
    rawX < 0 ? Math.max(0, rawWidth + rawX) : rawWidth;
  let height =
    rawY < 0 ? Math.max(0, rawHeight + rawY) : rawHeight;

  width = Math.max(MIN_DIMENSION, Math.min(widthCap, width));
  height = Math.max(MIN_DIMENSION, Math.min(heightCap, height));

  return { x, y, width, height };
}

export function normalizeHighlightRects(
  clientRects: ReadonlyArray<BoundingRect>,
  wrapper: BoundingRect,
): NormalizedRect[] {
  const out: NormalizedRect[] = [];
  for (const client of clientRects) {
    const rect = normalizeHighlightRect(client, wrapper);
    if (rect) out.push(rect);
  }
  return out;
}

export interface HighlightPageBox {
  pageNumber: number;
  rect: BoundingRect;
}

export interface HighlightPageRects {
  pageNumber: number;
  rects: NormalizedRect[];
}










export function splitHighlightRectsByPage(
  clientRects: ReadonlyArray<BoundingRect>,
  pages: ReadonlyArray<HighlightPageBox>,
): HighlightPageRects[] {
  const byPage = new Map<number, NormalizedRect[]>();
  for (const client of clientRects) {
    if (client.width <= 0 || client.height <= 0) continue;
    const cx = client.left + client.width / 2;
    const cy = client.top + client.height / 2;
    let owner = pages.find(
      ({ rect }) =>
        cx >= rect.left &&
        cx <= rect.left + rect.width &&
        cy >= rect.top &&
        cy <= rect.top + rect.height,
    );
    if (!owner) {
      let best = 0;
      for (const page of pages) {
        const overlap =
          Math.max(
            0,
            Math.min(client.left + client.width, page.rect.left + page.rect.width) -
              Math.max(client.left, page.rect.left),
          ) *
          Math.max(
            0,
            Math.min(client.top + client.height, page.rect.top + page.rect.height) -
              Math.max(client.top, page.rect.top),
          );
        if (overlap > best) {
          best = overlap;
          owner = page;
        }
      }
    }
    if (!owner) continue;
    const rect = normalizeHighlightRect(client, owner.rect);
    if (!rect) continue;
    const list = byPage.get(owner.pageNumber) ?? [];
    list.push(rect);
    byPage.set(owner.pageNumber, list);
  }
  return [...byPage.entries()]
    .sort(([a], [b]) => a - b)
    .map(([pageNumber, rects]) => ({ pageNumber, rects }));
}
