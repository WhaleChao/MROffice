/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function pdfRangeGlyphRect(range: Range): DOMRect {
  const rects = typeof range.getClientRects === "function" ? Array.from(range.getClientRects()) : [];
  const drawn = rects.filter((rect) => rect.width > 0 && rect.height > 0);
  if (drawn.length === 1) return drawn[0];
  if (drawn.length > 1) {
    const left = Math.min(...drawn.map((rect) => rect.left));
    const top = Math.min(...drawn.map((rect) => rect.top));
    const right = Math.max(...drawn.map((rect) => rect.right));
    const bottom = Math.max(...drawn.map((rect) => rect.bottom));
    return new DOMRect(left, top, right - left, bottom - top);
  }
  return range.getBoundingClientRect();
}
