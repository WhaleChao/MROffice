/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PdfTextBoxRectLike {
  x?: unknown;
  y?: unknown;
  width?: unknown;
  height?: unknown;
}

export interface PdfTextBoxRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PdfTextBoxGeometryLike extends PdfTextBoxRectLike {
  originX?: unknown;
  originY?: unknown;
  originWidth?: unknown;
  originHeight?: unknown;
}

export interface PdfTextBoxGeometry extends PdfTextBoxRect {
  originX?: number;
  originY?: number;
  originWidth?: number;
  originHeight?: number;
}

/** One run of a line as the reader reads it (a pdf.js text item): its words,
 * its font as the PDF names it, and that face's own weight and slant. */
export type PdfNativeTextReaderRun = Readonly<{ text: string; sourceFontName: string; bold: boolean; italic: boolean }>;

export interface PdfNativeTextSourceLayout extends PdfTextBoxRectLike {
  id?: string;
  page?: number;
  fontSize?: number;
  angle?: number;
  sourceFontName?: string;
  displayFontFamily?: string;
  /** The face's own weight and slant (its pdf.js font, or its name). */
  faceBold?: boolean;
  faceItalic?: boolean;
}

function finiteNumber(value: unknown): number | null {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}

function rightEdge(x: number, width: number): number {
  const right = x + width;
  if (Number.isFinite(right)) return right;
  return width > 0 ? Number.POSITIVE_INFINITY : x;
}

function stableFraction(value: number): number {
  return Math.round(value * 1_000_000_000_000) / 1_000_000_000_000;
}

/**
 * Persisted PDF text-box geometry is stored in page-relative fractions. Older
 * documents and AI commands may contain slightly off-page values, so intersect
 * the box with the visible page instead of trusting it or dropping salvageable
 * content. Completely off-page / empty boxes are rejected.
 */
export function sanitizePdfTextBoxRect(
  rect: PdfTextBoxRectLike,
): PdfTextBoxRect | null {
  const x = finiteNumber(rect.x);
  const y = finiteNumber(rect.y);
  const width = finiteNumber(rect.width);
  const height = finiteNumber(rect.height);
  if (
    x === null ||
    y === null ||
    width === null ||
    height === null ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }

  const left = Math.max(0, x);
  const top = Math.max(0, y);
  const right = Math.min(1, rightEdge(x, width));
  const bottom = Math.min(1, rightEdge(y, height));
  const nextWidth = right - left;
  const nextHeight = bottom - top;
  if (nextWidth <= 0 || nextHeight <= 0) return null;

  return {
    x: stableFraction(left),
    y: stableFraction(top),
    width: stableFraction(nextWidth),
    height: stableFraction(nextHeight),
  };
}

export function sanitizePdfTextBoxEditGeometry(
  edit: PdfTextBoxGeometryLike,
): PdfTextBoxGeometry | null {
  const rect = sanitizePdfTextBoxRect(edit);
  if (!rect) return null;

  const origin = sanitizePdfTextBoxRect({
    x: edit.originX,
    y: edit.originY,
    width: edit.originWidth,
    height: edit.originHeight,
  });
  if (!origin) return rect;

  return {
    ...rect,
    originX: origin.x,
    originY: origin.y,
    originWidth: origin.width,
    originHeight: origin.height,
  };
}

/** Repair old native-hit rectangles using an exact item-id source, not proximity.
 * Keep user movement/resizing relative to the original; never edit stored data.
 */
export function alignNativeTextBoxEditToSource<
  T extends PdfTextBoxGeometryLike & {
    id?: string;
    page?: number;
    created?: boolean;
    angle?: number;
    writingMode?: string;
  },
>(edit: T, sources: readonly PdfNativeTextSourceLayout[]): T {
  if (
    edit.created ||
    typeof edit.id !== "string" ||
    !/^native-text:\d+:\d+:\d+$/.test(edit.id) ||
    edit.angle ||
    edit.writingMode
  )
    return edit;
  const matches = sources.filter(
    (source) =>
      source.id === edit.id && source.page === edit.page && !source.angle,
  );
  if (matches.length !== 1) return edit;
  const source = sanitizePdfTextBoxRect(matches[0]);
  if (!source) return edit;
  const geometry = sanitizePdfTextBoxEditGeometry(edit);
  if (
    !geometry ||
    geometry.originX === undefined ||
    geometry.originY === undefined ||
    geometry.originWidth === undefined ||
    geometry.originHeight === undefined
  )
    return edit;
  const sourceSize = matches[0].fontSize;
  const tolerance =
    (typeof sourceSize === "number" &&
    Number.isFinite(sourceSize) &&
    sourceSize > 0
      ? sourceSize
      : source.height) * 0.5;
  if (
    Math.abs(source.y - geometry.originY) > tolerance ||
    Math.abs(source.x - geometry.originX) > tolerance ||
    Math.abs(source.width - geometry.originWidth) > tolerance
  )
    return edit;
  return {
    ...edit,
    x: geometry.x + source.x - geometry.originX,
    y: geometry.y + source.y - geometry.originY,
    width:
      Math.abs(geometry.width - geometry.originWidth) < 1e-8
        ? source.width
        : geometry.width,
    height:
      Math.abs(geometry.height - geometry.originHeight) < 1e-8
        ? source.height
        : geometry.height,
    originX: source.x,
    originY: source.y,
    originWidth: source.width,
    originHeight: source.height,
  };
}
