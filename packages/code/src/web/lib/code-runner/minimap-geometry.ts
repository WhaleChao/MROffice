/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry math for the code editor minimap (VSCode-style overview column).
 *
 * Framework-free so it is fully unit-tested; the editor renders the column and
 * the draggable viewport rectangle, and wires these helpers to the textarea's
 * scrollTop / clientHeight / scrollHeight.
 *
 * Coordinate spaces:
 *  - "doc px"     : the textarea's own scroll metrics (scrollTop, scrollHeight,
 *                   clientHeight) in CSS pixels.
 *  - "minimap px" : vertical position inside the minimap column, 0 at the top of
 *                   the column down to `minimapHeight` at the bottom.
 *
 * All helpers are total (never throw) and defensive against zero / NaN inputs so
 * a freshly-mounted, unmeasured editor renders something sane instead of NaN.
 */

/** A rectangle inside the minimap column, in minimap px. */
export interface ViewportRect {
  /** Distance from the top of the minimap column to the top of the rect. */
  top: number;
  /** Height of the rect (the slice of the file currently visible). */
  height: number;
}

/** Caps for per-line density blocks, mirroring VSCode's minimap proportions. */
export interface BlockOptions {
  /** Widest a block can get (px), reached at or above `maxChars`. */
  maxWidth?: number;
  /** Line length (chars) that maps to the full `maxWidth`. */
  maxChars?: number;
  /** Minimum width for a non-empty line so it stays visible. */
  minWidth?: number;
}

const DEFAULT_BLOCK: Required<BlockOptions> = {
  maxWidth: 56,
  maxChars: 80,
  minWidth: 1,
};

/**
 * Clamp a number into [lo, hi]. NaN coerces to `lo`; ±Infinity clamp to the
 * nearer bound (so a runaway scrollTop pins to the edge instead of breaking).
 */
export function clamp(value: number, lo: number, hi: number): number {
  if (Number.isNaN(value)) return lo;
  if (value < lo) return lo;
  if (value > hi) return hi;
  return value;
}

/**
 * Whether the minimap should be visible by default for a file of `totalLines`.
 * VSCode shows it always; here we keep it simple and only auto-enable past a
 * line threshold so tiny snippets stay uncluttered (the user can still toggle).
 */
export function shouldShowMinimapByDefault(
  totalLines: number,
  threshold = 40,
): boolean {
  if (!Number.isFinite(totalLines)) return false;
  return totalLines > threshold;
}

/**
 * Height (px) of a single line's density block in the minimap. The whole file
 * is scaled to fit `minimapHeight`, but each block is capped at `maxBlockHeight`
 * so a short file doesn't render fat bars; for long files lines compress (and
 * may be sub-pixel, which is fine for a density overview).
 */
export function lineBlockHeight(
  totalLines: number,
  minimapHeight: number,
  maxBlockHeight = 4,
): number {
  const lines = Math.max(1, Math.floor(totalLines) || 1);
  const h = Math.max(0, minimapHeight) / lines;
  return Math.min(h, Math.max(0, maxBlockHeight));
}

/**
 * Total px height the stacked line blocks occupy in the minimap. This equals
 * the column height for long files (lines compressed to fill it) and is shorter
 * than the column for small files (blocks capped, leaving slack at the bottom).
 */
export function contentHeight(
  totalLines: number,
  minimapHeight: number,
  maxBlockHeight = 4,
): number {
  const lines = Math.max(1, Math.floor(totalLines) || 1);
  return lineBlockHeight(lines, minimapHeight, maxBlockHeight) * lines;
}

/**
 * Width (px) of one line's density block: proportional to line length, capped.
 * Empty lines render zero width; any non-empty line is at least `minWidth`.
 */
export function lineBlockWidth(
  lineLength: number,
  options: BlockOptions = {},
): number {
  const { maxWidth, maxChars, minWidth } = { ...DEFAULT_BLOCK, ...options };
  const len = Math.max(0, Math.floor(lineLength) || 0);
  if (len <= 0) return 0;
  const chars = maxChars > 0 ? maxChars : 1;
  const w = (len / chars) * maxWidth;
  return clamp(w, minWidth, Math.max(minWidth, maxWidth));
}

/** Per-line block widths for an array of line lengths (caps applied). */
export function lineBlockWidths(
  lineLengths: readonly number[],
  options: BlockOptions = {},
): number[] {
  return lineLengths.map((len) => lineBlockWidth(len, options));
}

/**
 * The viewport rectangle (in minimap px) for the currently-visible slice of the
 * document. Mapped against the minimap's CONTENT height (see {@link contentHeight})
 * so the rect lines up with the rendered density blocks, not the column slack.
 *
 * Degenerate cases (everything fits, no scroll) yield a rect that spans the full
 * content, so the editor always draws a sensible box.
 */
export function computeViewportRect(
  scrollTop: number,
  clientHeight: number,
  scrollHeight: number,
  minimapHeight: number,
  totalLines = 1,
  maxBlockHeight = 4,
): ViewportRect {
  const content = contentHeight(totalLines, minimapHeight, maxBlockHeight);
  if (content <= 0) return { top: 0, height: 0 };
  // Nothing to scroll: the whole document is visible.
  if (!(scrollHeight > clientHeight) || !(scrollHeight > 0)) {
    return { top: 0, height: content };
  }
  const ratio = content / scrollHeight;
  const rawTop = scrollTop * ratio;
  const rawHeight = clientHeight * ratio;
  // Keep a 1px minimum visible height, but never exceed the content itself: for
  // sub-pixel content (content < 1) forcing a 1px floor would overflow the box.
  const minHeight = Math.min(1, content);
  const height = clamp(rawHeight, minHeight, content);
  // content - height is always >= 0 here, so top stays within [0, content].
  const top = clamp(rawTop, 0, Math.max(0, content - height));
  return { top, height };
}

/**
 * Map a click/drag position (minimap px, measured from the top of the column) to
 * a target textarea scrollTop, CENTERING the viewport on the clicked point the
 * way VSCode does. Clamped to the legal scroll range.
 */
export function minimapYToScrollTop(
  y: number,
  clientHeight: number,
  scrollHeight: number,
  minimapHeight: number,
  totalLines = 1,
  maxBlockHeight = 4,
): number {
  const maxScroll = Math.max(0, scrollHeight - clientHeight);
  if (maxScroll <= 0) return 0;
  const content = contentHeight(totalLines, minimapHeight, maxBlockHeight);
  if (content <= 0) return 0;
  const yy = clamp(y, 0, content);
  // Position in the document the click points at, then center the viewport.
  const docCenter = (yy / content) * scrollHeight;
  const target = docCenter - clientHeight / 2;
  return clamp(target, 0, maxScroll);
}

/**
 * Map a click position (minimap px) to a 1-based line number. Useful for
 * "scroll to this line" behaviour and for tests independent of pixel scroll.
 */
export function minimapYToLine(
  y: number,
  minimapHeight: number,
  totalLines: number,
  maxBlockHeight = 4,
): number {
  const lines = Math.max(1, Math.floor(totalLines) || 1);
  const content = contentHeight(lines, minimapHeight, maxBlockHeight);
  if (content <= 0) return 1;
  const yy = clamp(y, 0, content);
  const line = Math.floor((yy / content) * lines) + 1;
  return clamp(line, 1, lines);
}
