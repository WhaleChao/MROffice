/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry helpers for moving/resizing PDF text boxes (WPS-style handles).
 * All coordinates are page-relative fractions (0..1). Framework-free for unit
 * testing; the overlay applies these to pointer deltas.
 */

export interface BoxRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Small epsilon for comparing page-fraction rectangles. */
export const BOX_RECT_EPSILON = 0.0001;

/** Return the bounding rectangle that covers both input rectangles. */
export function unionBoxRects(a: BoxRect, b: BoxRect): BoxRect {
  const left = clamp(Math.min(a.x, b.x), 0, 1);
  const top = clamp(Math.min(a.y, b.y), 0, 1);
  const right = clamp(Math.max(a.x + a.width, b.x + b.width), 0, 1);
  const bottom = clamp(Math.max(a.y + a.height, b.y + b.height), 0, 1);
  return {
    x: left,
    y: top,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  };
}

/** True when `outer` already covers `inner`, allowing tiny float drift. */
export function boxRectContains(
  outer: BoxRect,
  inner: BoxRect,
  epsilon = BOX_RECT_EPSILON,
): boolean {
  return (
    outer.x <= inner.x + epsilon &&
    outer.y <= inner.y + epsilon &&
    outer.x + outer.width >= inner.x + inner.width - epsilon &&
    outer.y + outer.height >= inner.y + inner.height - epsilon
  );
}

/** Resize handle directions (corners + edges), like WPS selection handles. */
export type ResizeHandle = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";

export const MIN_BOX_FRACTION = 0.01;

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

/** Move a box by a fractional delta, clamped so it stays within the page.
 * A box that already reaches past the page (a rotated arXiv stamp, a frame
 * drawn to the edge) is not pulled onto it: it moves by the delta, and only
 * further out than it stands is refused. */
export function moveBox<T extends BoxRect>(box: T, dx: number, dy: number): T {
  const width = clamp(box.width, MIN_BOX_FRACTION, 1);
  const height = clamp(box.height, MIN_BOX_FRACTION, 1);
  return {
    ...box,
    x: clamp(box.x + dx, Math.min(0, box.x), Math.max(1 - width, box.x)),
    y: clamp(box.y + dy, Math.min(0, box.y), Math.max(1 - height, box.y)),
  };
}

/**
 * Resize a box by dragging a handle. Edge handles move one side; corner handles
 * move two. The box stays within [0,1] and never shrinks below MIN_BOX_FRACTION.
 */
export function resizeBox<T extends BoxRect>(
  box: T,
  handle: ResizeHandle,
  dx: number,
  dy: number,
): T {
  let { x, y, width, height } = box;
  const right = x + width;
  const bottom = y + height;

  if (handle.includes("e")) {
    width = clamp(width + dx, MIN_BOX_FRACTION, 1 - x);
  }
  if (handle.includes("w")) {
    const nextX = clamp(x + dx, 0, right - MIN_BOX_FRACTION);
    width = right - nextX;
    x = nextX;
  }
  if (handle.includes("s")) {
    height = clamp(height + dy, MIN_BOX_FRACTION, 1 - y);
  }
  if (handle.includes("n")) {
    const nextY = clamp(y + dy, 0, bottom - MIN_BOX_FRACTION);
    height = bottom - nextY;
    y = nextY;
  }

  return { ...box, x, y, width, height };
}

/**
 * The ORIGINAL glyph region an edit must keep covered, independent of where the
 * (movable) edit box currently sits. An edit records the source box rect once
 * (`originX/originY/originWidth/originHeight`); dragging/resizing only changes
 * the live rect, never the origin. This is the rect the opaque cover (overlay)
 * and the bake mask (export) paint over so the original PDF glyphs never show
 * through after the box is moved away.
 *
 * Backward-compatible: an edit persisted before origin tracking omits these
 * fields, so the cover falls back to the edit's CURRENT rect — exactly the old
 * (move-the-cover-with-the-box) behaviour, no worse than before.
 */
export interface OriginRectLike {
  x: number;
  y: number;
  width: number;
  height: number;
  originX?: number;
  originY?: number;
  originWidth?: number;
  originHeight?: number;
}

export interface GlyphCoverOptions {
  /**
   * Width/height of the rendered PDF page in the same pixel space. Font size is
   * stored as a fraction of page height, while x/width are fractions of page
   * width; portrait pages therefore need a wider horizontal bleed in fraction
   * space to cover the same number of painted pixels.
   */
  pageWidth?: number;
  pageHeight?: number;
  /** Alias for callers that already carry a `{ width, height }` page record. */
  width?: number;
  height?: number;
  /** Display rotation in degrees; 90/270 swaps rendered width and height. */
  rotation?: number;
}

function normalizedQuarterTurn(value: unknown): 0 | 90 | 180 | 270 {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  const normalized = ((Math.trunc(numeric) % 360) + 360) % 360;
  return normalized === 90 || normalized === 180 || normalized === 270
    ? normalized
    : 0;
}

function renderedGlyphCoverPageSize(
  options: GlyphCoverOptions,
): { pageWidth: number | null; pageHeight: number | null } {
  const rawWidth =
    Number.isFinite(options.pageWidth) && (options.pageWidth as number) > 0
      ? (options.pageWidth as number)
      : Number.isFinite(options.width) && (options.width as number) > 0
        ? (options.width as number)
        : null;
  const rawHeight =
    Number.isFinite(options.pageHeight) && (options.pageHeight as number) > 0
      ? (options.pageHeight as number)
      : Number.isFinite(options.height) && (options.height as number) > 0
        ? (options.height as number)
        : null;
  const rotation = normalizedQuarterTurn(options.rotation);
  if (rawWidth && rawHeight && (rotation === 90 || rotation === 270)) {
    return { pageWidth: rawHeight, pageHeight: rawWidth };
  }
  return { pageWidth: rawWidth, pageHeight: rawHeight };
}

/**
 * The stationary cover rect for an edit: its recorded origin region when all
 * four origin fields are finite, else the edit's current rect. Pure — the
 * overlay paints this (fixed) and the bake masks this (fixed) while the editable
 * box itself moves freely on top.
 */
export function coverRectForEdit(edit: OriginRectLike): BoxRect {
  const { originX, originY, originWidth, originHeight } = edit;
  if (
    Number.isFinite(originX) &&
    Number.isFinite(originY) &&
    Number.isFinite(originWidth) &&
    Number.isFinite(originHeight) &&
    (originWidth as number) > 0 &&
    (originHeight as number) > 0
  ) {
    return {
      x: originX as number,
      y: originY as number,
      width: originWidth as number,
      height: originHeight as number,
    };
  }
  return { x: edit.x, y: edit.y, width: edit.width, height: edit.height };
}

/**
 * PDF text item boxes describe an advance width, not the exact painted glyph
 * bounds. Embedded fonts, horizontal scaling and glyph overhang can draw a few
 * pixels outside that box; if the cover/mask is exact, the original tail of a
 * line can remain visible beside the edited text. Expand only the cover rect,
 * never the editable box itself.
 */
export function expandRectForGlyphCover(
  rect: BoxRect,
  fontSizeFraction?: number,
  options: GlyphCoverOptions = {},
): BoxRect {
  const size =
    Number.isFinite(fontSizeFraction) && (fontSizeFraction as number) > 0
      ? (fontSizeFraction as number)
      : Math.max(MIN_BOX_FRACTION, rect.height * 0.8);
  const { pageWidth, pageHeight } = renderedGlyphCoverPageSize(options);
  const widthScale =
    pageWidth && pageHeight ? clamp(pageHeight / pageWidth, 0.5, 2.5) : 1.35;
  const xSize = size * widthScale;
  const leftBleed = Math.max(0.001, Math.min(0.016, xSize * 0.24));
  const isTitleSized = size >= 0.028;
  const rightBleed = Math.max(
    0.004,
    Math.min(isTitleSized ? 0.075 : 0.045, xSize * (isTitleSized ? 1.8 : 1.15)),
  );
  const yBleed = Math.max(0.001, Math.min(0.012, size * 0.22));

  const left = clamp(rect.x - leftBleed, 0, 1);
  const top = clamp(rect.y - yBleed, 0, 1);
  const right = clamp(rect.x + rect.width + rightBleed, 0, 1);
  const bottom = clamp(rect.y + rect.height + yBleed, 0, 1);
  return {
    x: left,
    y: top,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  };
}

/**
 * The axis-aligned bounds, in page fractions, of `rect` turned about its own
 * centre by `angle` degrees (clockwise on screen), as a box with an angle is
 * drawn. A line on a page turned a quarter is recorded in its own axes with
 * the page's turn as its angle, so the cover of the scanned words it
 * replaced (PDF 0107) lies where the turned rect does, not where the
 * recorded one does: unturned, it lay across the lines beside it. Exact for
 * a quarter turn, the bounds for any other angle. `options` names the page
 * as shown so the turn keeps the pixel aspect.
 */
export function turnedRectBounds(
  rect: BoxRect,
  angle: number | undefined,
  options: GlyphCoverOptions = {},
): BoxRect {
  const degrees = Number(angle);
  const turn = Number.isFinite(degrees) ? ((degrees % 360) + 360) % 360 : 0;
  if (turn < 1e-6 || 360 - turn < 1e-6) return rect;
  const { pageWidth, pageHeight } = renderedGlyphCoverPageSize(options);
  const w = pageWidth && pageHeight ? pageWidth : 1;
  const h = pageWidth && pageHeight ? pageHeight : 1;
  const cx = (rect.x + rect.width / 2) * w;
  const cy = (rect.y + rect.height / 2) * h;
  const halfW = (rect.width * w) / 2;
  const halfH = (rect.height * h) / 2;
  const radians = (turn * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const boundsHalfW = halfW * cos + halfH * sin;
  const boundsHalfH = halfW * sin + halfH * cos;
  const left = clamp((cx - boundsHalfW) / w, 0, 1);
  const top = clamp((cy - boundsHalfH) / h, 0, 1);
  const right = clamp((cx + boundsHalfW) / w, 0, 1);
  const bottom = clamp((cy + boundsHalfH) / h, 0, 1);
  return { x: left, y: top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
}

/**
 * Shrink a glyph-cover rect so it does not intrude into neighbouring text.
 *
 * The cover overshoots its origin on purpose (anti-aliased glyph edges), but
 * on a dense page the overshoot lands on the NEXT line, the bullet before the
 * text, or the label beside it — and paints an opaque block over words the
 * user never touched. Each side retreats just far enough to clear any
 * neighbour it would overlap, and never past the origin itself: hiding the
 * original is still the cover's job.
 *
 * A neighbour that already overlaps the ORIGIN is ignored — it was under the
 * text to begin with (overlapping extraction artefacts), and retreating for
 * it would expose the very glyphs the cover exists to hide.
 */
export function clampCoverToNeighbors(
  cover: BoxRect,
  origin: BoxRect,
  neighbors: readonly BoxRect[],
): BoxRect {
  const EPS = 0.0005;
  const originLeft = clamp(origin.x, 0, 1);
  const originTop = clamp(origin.y, 0, 1);
  const originRight = clamp(origin.x + origin.width, 0, 1);
  const originBottom = clamp(origin.y + origin.height, 0, 1);

  let left = cover.x;
  let top = cover.y;
  let right = cover.x + cover.width;
  let bottom = cover.y + cover.height;
  let retreated = false;

  const overlaps = (
    aLeft: number,
    aTop: number,
    aRight: number,
    aBottom: number,
    b: BoxRect,
  ) =>
    aLeft < b.x + b.width - EPS &&
    aRight > b.x + EPS &&
    aTop < b.y + b.height - EPS &&
    aBottom > b.y + EPS;

  for (const neighbor of neighbors) {
    if (!(neighbor.width > 0) || !(neighbor.height > 0)) continue;
    // Already under the original text: not the cover's problem.
    if (overlaps(originLeft, originTop, originRight, originBottom, neighbor)) {
      continue;
    }
    if (!overlaps(left, top, right, bottom, neighbor)) continue;
    const nLeft = neighbor.x;
    const nTop = neighbor.y;
    const nRight = neighbor.x + neighbor.width;
    const nBottom = neighbor.y + neighbor.height;
    // Retreat the one side whose retreat is smallest, and only into the
    // bleed — never past the origin edge.
    const candidates: Array<{ cost: number; apply: () => void }> = [];
    if (nRight <= originLeft + EPS) {
      candidates.push({
        cost: nRight - left,
        apply: () => {
          left = Math.max(left, Math.min(nRight, originLeft));
        },
      });
    }
    if (nLeft >= originRight - EPS) {
      candidates.push({
        cost: right - nLeft,
        apply: () => {
          right = Math.min(right, Math.max(nLeft, originRight));
        },
      });
    }
    if (nBottom <= originTop + EPS) {
      candidates.push({
        cost: nBottom - top,
        apply: () => {
          top = Math.max(top, Math.min(nBottom, originTop));
        },
      });
    }
    if (nTop >= originBottom - EPS) {
      candidates.push({
        cost: bottom - nTop,
        apply: () => {
          bottom = Math.min(bottom, Math.max(nTop, originBottom));
        },
      });
    }
    if (candidates.length === 0) continue;
    candidates.sort((a, b) => a.cost - b.cost);
    candidates[0].apply();
    retreated = true;
  }

  // Rebuilding width/height from edges picks up float noise; an untouched
  // cover comes back exactly as it went in.
  if (!retreated) return cover;
  return {
    x: left,
    y: top,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  };
}

/** The 8 handles with their CSS anchor (fractions of the box) and cursor. */
export const RESIZE_HANDLES: {
  handle: ResizeHandle;
  cx: number;
  cy: number;
  cursor: string;
}[] = [
  { handle: "nw", cx: 0, cy: 0, cursor: "nwse-resize" },
  { handle: "n", cx: 0.5, cy: 0, cursor: "ns-resize" },
  { handle: "ne", cx: 1, cy: 0, cursor: "nesw-resize" },
  { handle: "e", cx: 1, cy: 0.5, cursor: "ew-resize" },
  { handle: "se", cx: 1, cy: 1, cursor: "nwse-resize" },
  { handle: "s", cx: 0.5, cy: 1, cursor: "ns-resize" },
  { handle: "sw", cx: 0, cy: 1, cursor: "nesw-resize" },
  { handle: "w", cx: 0, cy: 0.5, cursor: "ew-resize" },
];
