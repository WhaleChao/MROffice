/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry for the code editor's minimap (a scaled-down line-density strip
 * shown beside the editor). Framework-free + unit-tested; `CodeMirrorEditor`
 * renders the strip and wires the scroll/click handlers to these functions.
 *
 * There is intentionally NO new dependency here: the minimap is driven entirely
 * from CM6's `scrollDOM` metrics (`scrollTop` / `scrollHeight` / `clientHeight`)
 * plus the document line count, all of which the component already has.
 */

/** Scroll metrics read off CM6's `EditorView.scrollDOM`. */
export interface ScrollMetrics {
  scrollTop: number;
  scrollHeight: number;
  clientHeight: number;
}

/** Geometry of the minimap viewport box (the highlighted "you are here" rect). */
export interface MinimapViewport {
  /** Top of the viewport box, in minimap pixels. */
  top: number;
  /** Height of the viewport box, in minimap pixels. */
  height: number;
}

/** One density bar rendered in the minimap. Lines are zero-based. */
export interface MinimapDensityMark {
  startLine: number;
  lineSpan: number;
  fill: number;
}

/** Bounded density model consumed by the React minimap renderer. */
export interface MinimapDensity {
  lineCount: number;
  marks: MinimapDensityMark[];
}

/**
 * Maximum minimap density marks. The silhouette is painted on one canvas, but
 * a mark per source line would still make painting grow with the file; 400
 * retains a useful document silhouette at a bounded cost.
 */
export const MAX_MINIMAP_DENSITY_MARKS = 400;

/**
 * How many lines a document has, without building an array of them.
 *
 * {@link minimapDensity} already returns this, but only while the silhouette is
 * being rendered — with the minimap off (by preference, or automatically on a
 * large file) its model collapses to `lineCount: 1`. Anything that needs the
 * REAL count, like the Go-to-Line input telling the user the range it accepts,
 * has to ask separately or it will advertise a one-line document. Counting
 * newlines allocates nothing, which matters because the caller reaches for this
 * exactly when the document is big enough that `split` was avoided.
 */
export function countDocumentLines(value: string): number {
  let lines = 1;
  for (let index = 0; index < value.length; index += 1) {
    if (value.charCodeAt(index) === 10) lines += 1;
  }
  return lines;
}

/**
 * Build the minimap silhouette in one linear scan. Small documents keep one
 * mark per line. Larger documents are divided into bounded contiguous buckets,
 * whose average non-whitespace density becomes one proportionally tall mark.
 */
export function minimapDensity(
  value: string,
  maxMarks = MAX_MINIMAP_DENSITY_MARKS,
): MinimapDensity {
  const lines = value.split("\n");
  const lineCount = Math.max(1, lines.length);
  const boundedMax = Number.isFinite(maxMarks)
    ? Math.max(1, Math.floor(maxMarks))
    : MAX_MINIMAP_DENSITY_MARKS;
  const bucketCount = Math.min(lineCount, boundedMax);
  const marks: MinimapDensityMark[] = [];

  for (let bucket = 0; bucket < bucketCount; bucket += 1) {
    const startLine = Math.floor((bucket * lineCount) / bucketCount);
    const endLine = Math.max(
      startLine + 1,
      Math.floor(((bucket + 1) * lineCount) / bucketCount),
    );
    let densityTotal = 0;
    for (let line = startLine; line < endLine; line += 1) {
      densityTotal += Math.min((lines[line]?.trim().length ?? 0) / 40, 1);
    }
    const fill = densityTotal / (endLine - startLine);
    if (fill <= 0) continue;
    marks.push({
      startLine,
      lineSpan: endLine - startLine,
      fill,
    });
  }

  return { lineCount, marks };
}

/**
 * Fraction (0..1) of the document that is scrolled past the top. Guards the
 * degenerate case where content fits without scrolling (scrollHeight <=
 * clientHeight) by returning 0.
 */
export function scrollFraction(m: ScrollMetrics): number {
  const scrollable = m.scrollHeight - m.clientHeight;
  if (scrollable <= 0) return 0;
  return clamp(m.scrollTop / scrollable, 0, 1);
}

/**
 * The viewport box to draw on a minimap of height `minimapHeight`, given the
 * editor's scroll metrics. The box height is proportional to the visible
 * fraction of the document; its top tracks the scroll position so the box
 * travels from the strip's top (at scrollTop 0) to its bottom (fully scrolled).
 */
export function minimapViewport(
  m: ScrollMetrics,
  minimapHeight: number,
): MinimapViewport {
  if (minimapHeight <= 0 || m.scrollHeight <= 0) {
    return { top: 0, height: 0 };
  }
  const visibleFraction = clamp(m.clientHeight / m.scrollHeight, 0, 1);
  const height = Math.max(minimapHeight * visibleFraction, MIN_BOX_HEIGHT);
  // Available travel for the box top so it never overflows the strip.
  const travel = Math.max(minimapHeight - height, 0);
  const top = travel * scrollFraction(m);
  return { top, height };
}

/**
 * Map a click/drag at `pointerY` (pixels from the minimap's top) to the
 * `scrollTop` that centres the editor viewport on that point. Used so clicking
 * the minimap scrubs the editor.
 */
export function scrollTopForPointer(
  pointerY: number,
  minimapHeight: number,
  m: ScrollMetrics,
): number {
  const scrollable = Math.max(m.scrollHeight - m.clientHeight, 0);
  if (minimapHeight <= 0 || scrollable <= 0) return 0;
  const fraction = clamp(pointerY / minimapHeight, 0, 1);
  return Math.round(fraction * scrollable);
}

/**
 * Per-line layout for rendering the minimap as stacked line marks. Returns the
 * top offset (px, within the minimap) and height of each line's mark so a
 * caller can draw `lineCount` evenly-distributed bars. Distinct from the
 * viewport box (which is the scroll indicator).
 */
export function minimapLineHeight(lineCount: number, minimapHeight: number): number {
  if (lineCount <= 0 || minimapHeight <= 0) return 0;
  return minimapHeight / lineCount;
}

/** Horizontal inset (px) of every density mark inside the minimap strip. */
export const MINIMAP_MARK_INSET = 4;

/** Where one density mark is painted, in minimap CSS pixels. */
export interface MinimapMarkRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Geometry of one density mark on a strip `width` px wide whose lines are
 * `lineHeight` px tall. This is the exact box the per-line DOM marks used to
 * occupy (left inset, proportional width with a 2px floor, 1px gap between
 * consecutive lines), so the canvas silhouette is the same picture.
 */
export function minimapMarkRect(
  mark: MinimapDensityMark,
  lineHeight: number,
  width: number,
): MinimapMarkRect {
  return {
    x: MINIMAP_MARK_INSET,
    y: mark.startLine * lineHeight,
    width: Math.max(mark.fill * (width - MINIMAP_MARK_INSET * 2), 2),
    height: Math.max(mark.lineSpan * lineHeight - 1, 1),
  };
}

/** The slice of a 2D canvas context the minimap painter uses. */
export interface MinimapPaintContext {
  fillStyle: string | CanvasGradient | CanvasPattern;
  clearRect(x: number, y: number, width: number, height: number): void;
  fillRect(x: number, y: number, width: number, height: number): void;
}

/**
 * Paint the density silhouette onto a canvas context whose transform already
 * maps CSS pixels (the caller applies devicePixelRatio). One `fillRect` per
 * mark and no allocation per mark: the strip is a single element however many
 * lines the document has. Returns the number of marks painted.
 */
export function paintMinimapDensity(
  context: MinimapPaintContext,
  model: MinimapDensity,
  width: number,
  height: number,
  color: string,
): number {
  context.clearRect(0, 0, width, height);
  const lineHeight = minimapLineHeight(model.lineCount, height);
  if (lineHeight <= 0 || width <= 0) return 0;
  context.fillStyle = color;
  for (const mark of model.marks) {
    const rect = minimapMarkRect(mark, lineHeight, width);
    context.fillRect(rect.x, rect.y, rect.width, rect.height);
  }
  return model.marks.length;
}

const MIN_BOX_HEIGHT = 12;

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}
