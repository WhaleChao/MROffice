/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface NotePopoverPlacement {
  readonly left: number;
  readonly top: number;
  /** Set when the popover is taller than the page; it scrolls inside. */
  readonly maxHeight: number | null;
}

/** Room kept between the popover and the paper's edge. */
const EDGE_MARGIN = 4;

function placeOnAxis(anchor: number, size: number, extent: number): number {
  if (size >= extent - 2 * EDGE_MARGIN) return Math.max(0, (extent - size) / 2);
  if (anchor + size <= extent - EDGE_MARGIN) return Math.max(EDGE_MARGIN, anchor);
  if (anchor - size >= EDGE_MARGIN) return anchor - size;
  return Math.max(EDGE_MARGIN, extent - EDGE_MARGIN - size);
}

export function planNotePopoverPlacement(input: {
  /** Anchor in page-relative percent [0..100]. */
  readonly x: number;
  readonly y: number;
  /** The popover's own size, in page pixels. */
  readonly width: number;
  readonly height: number;
  /** The page's size, in pixels. */
  readonly pageWidth: number;
  readonly pageHeight: number;
}): NotePopoverPlacement | null {
  const { x, y, width, height, pageWidth, pageHeight } = input;
  if (
    ![x, y, width, height, pageWidth, pageHeight].every(Number.isFinite) ||
    pageWidth <= 0 ||
    pageHeight <= 0 ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }
  const anchorX = (Math.min(Math.max(x, 0), 100) / 100) * pageWidth;
  const anchorY = (Math.min(Math.max(y, 0), 100) / 100) * pageHeight;
  const maxHeight = height > pageHeight ? Math.floor(pageHeight) : null;
  const effectiveHeight = maxHeight ?? height;
  return {
    left: Math.round(placeOnAxis(anchorX, width, pageWidth)),
    top: Math.round(
      maxHeight !== null ? 0 : placeOnAxis(anchorY, effectiveHeight, pageHeight),
    ),
    maxHeight,
  };
}
