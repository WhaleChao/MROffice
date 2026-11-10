/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { HighlightMarkupStyle } from "./highlight-markup";

export type ManagedPdfTextMarkupSubtype =
  "Highlight" | "Underline" | "StrikeOut" | "Squiggly";

export interface ManagedPdfTextMarkup {
  subtype: ManagedPdfTextMarkupSubtype;
  rect: [number, number, number, number];
  quadPoints: number[];
  /**
   * Reader text is retained in the page model for excerpts and reopen
   * matching, but is deliberately not written into PDF `/Contents`.
   */
  quote?: string;
  /** True when links/replies exist and native download must fail closed. */
  hasReaderOnlyFeatures?: boolean;
}

export interface ManagedPdfTextMarkupBinding {
  id: string;
  managedMarkup: ManagedPdfTextMarkup;
  /** The page row's own page and colour: what the reader's row of the mark
   * is written again from when it is missing (self-healing on load). */
  page?: number;
  color?: string;
  createdAt?: string;
}

export interface ManagedPdfTextMarkupSync {
  id: string;
  page: number;
  quote: string;
  note?: string;
  color: string;
  createdAt?: string;
  markupStyle: HighlightMarkupStyle;
  author?: string;
  replies?: unknown[];
  hasReaderOnlyFeatures?: boolean;
  managedMarkup: ManagedPdfTextMarkup;
}

export type ManagedPdfProofingKind = "replace" | "insert";

/**
 * Source geometry for one native PDF proofing annotation. Replace marks own a
 * standard `/Caret` parent plus a grouped `/StrikeOut` child; insert marks own
 * only the `/Caret`. `quote` is reader context and is never written into
 * `/Contents`, which remains the replacement/inserted text.
 */
export interface ManagedPdfProofingAnnotation {
  kind: ManagedPdfProofingKind;
  rect: [number, number, number, number];
  quadPoints?: number[];
  quote?: string;
}

export interface ManagedPdfProofingSync {
  id: string;
  page: number;
  replacementText: string;
  color: string;
  author?: string;
  managedProofing: ManagedPdfProofingAnnotation;
}

export interface PdfPointViewport {
  width: number;
  height: number;
  convertToPdfPoint(x: number, y: number): [number, number];
}

export interface PercentTextRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

const MAX_QUAD_POINT_VALUES = 64 * 8;
const MAX_QUOTE_LENGTH = 4_000;
const GEOMETRY_TOLERANCE = 2;

export function isManagedPdfProofingKind(
  value: unknown,
): value is ManagedPdfProofingKind {
  return value === "replace" || value === "insert";
}

export function isManagedPdfTextMarkupSubtype(
  value: unknown,
): value is ManagedPdfTextMarkupSubtype {
  return (
    value === "Highlight" ||
    value === "Underline" ||
    value === "StrikeOut" ||
    value === "Squiggly"
  );
}

export function highlightMarkupStyleToPdfSubtype(
  style: HighlightMarkupStyle,
): ManagedPdfTextMarkupSubtype {
  if (style === "underline") return "Underline";
  if (style === "strikethrough") return "StrikeOut";
  if (style === "squiggly") return "Squiggly";
  return "Highlight";
}

export function pdfSubtypeToHighlightMarkupStyle(
  subtype: ManagedPdfTextMarkupSubtype,
): HighlightMarkupStyle {
  if (subtype === "Underline") return "underline";
  if (subtype === "StrikeOut") return "strikethrough";
  if (subtype === "Squiggly") return "squiggly";
  return "highlight";
}

function roundPdfCoordinate(value: number): number {
  return Number(value.toFixed(6));
}

function normalizeRect(
  value: unknown,
): [number, number, number, number] | null {
  if (
    !Array.isArray(value) ||
    value.length !== 4 ||
    value.some(
      (coordinate) =>
        typeof coordinate !== "number" || !Number.isFinite(coordinate),
    )
  ) {
    return null;
  }
  const x0 = Math.min(value[0], value[2]);
  const y0 = Math.min(value[1], value[3]);
  const x1 = Math.max(value[0], value[2]);
  const y1 = Math.max(value[1], value[3]);
  if (x1 <= x0 || y1 <= y0) return null;
  return [
    roundPdfCoordinate(x0),
    roundPdfCoordinate(y0),
    roundPdfCoordinate(x1),
    roundPdfCoordinate(y1),
  ];
}

export function sanitizeManagedPdfTextMarkup(
  value: unknown,
): ManagedPdfTextMarkup | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  if (!isManagedPdfTextMarkupSubtype(source.subtype)) return undefined;
  const rect = normalizeRect(source.rect);
  if (!rect) return undefined;
  if (
    !Array.isArray(source.quadPoints) ||
    source.quadPoints.length < 8 ||
    source.quadPoints.length > MAX_QUAD_POINT_VALUES ||
    source.quadPoints.length % 8 !== 0 ||
    source.quadPoints.some(
      (coordinate) =>
        typeof coordinate !== "number" || !Number.isFinite(coordinate),
    )
  ) {
    return undefined;
  }
  const [x0, y0, x1, y1] = rect;
  const quadPoints = source.quadPoints.map(roundPdfCoordinate);
  for (let index = 0; index < quadPoints.length; index += 1) {
    const coordinate = quadPoints[index];
    const lower = index % 2 === 0 ? x0 : y0;
    const upper = index % 2 === 0 ? x1 : y1;
    if (
      coordinate < lower - GEOMETRY_TOLERANCE ||
      coordinate > upper + GEOMETRY_TOLERANCE
    ) {
      return undefined;
    }
  }
  const quote =
    typeof source.quote === "string"
      ? source.quote.slice(0, MAX_QUOTE_LENGTH)
      : undefined;
  return {
    subtype: source.subtype,
    rect,
    quadPoints,
    ...(quote ? { quote } : {}),
    ...(source.hasReaderOnlyFeatures === true
      ? { hasReaderOnlyFeatures: true }
      : {}),
  };
}

export function sanitizeManagedPdfProofingAnnotation(
  value: unknown,
): ManagedPdfProofingAnnotation | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  if (!isManagedPdfProofingKind(source.kind)) return undefined;
  const rect = normalizeRect(source.rect);
  if (!rect) return undefined;
  let quadPoints: number[] | undefined;
  if (source.kind === "replace") {
    if (
      !Array.isArray(source.quadPoints) ||
      source.quadPoints.length < 8 ||
      source.quadPoints.length > MAX_QUAD_POINT_VALUES ||
      source.quadPoints.length % 8 !== 0 ||
      source.quadPoints.some(
        (coordinate) =>
          typeof coordinate !== "number" ||
          !Number.isFinite(coordinate) ||
          Math.abs(coordinate) > 10_000_000,
      )
    ) {
      return undefined;
    }
    quadPoints = source.quadPoints.map(roundPdfCoordinate);
  } else if (source.quadPoints !== undefined) {
    return undefined;
  }
  const quote =
    typeof source.quote === "string"
      ? source.quote.slice(0, MAX_QUOTE_LENGTH)
      : undefined;
  return {
    kind: source.kind,
    rect,
    ...(quadPoints ? { quadPoints } : {}),
    ...(quote ? { quote } : {}),
  };
}

function rgbChannel(value: string): number | null {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > 255) return null;
  return Math.round(parsed);
}

function channelHex(value: number): string {
  return value.toString(16).toUpperCase().padStart(2, "0");
}

/**
 * Convert the reader's strict hex/rgb(a) color subset to native PDF RGB.
 * Alpha is intentionally ignored: native highlight opacity is emitted in
 * `/CA` and the appearance stream, while line markups remain opaque.
 */
export function pdfMarkupColorToHex(
  value: unknown,
  fallback = "#FCD34D",
): string {
  if (typeof value === "string") {
    const hex = value.match(/^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/);
    if (hex) {
      const body =
        hex[1].length === 3
          ? hex[1]
              .split("")
              .map((character) => `${character}${character}`)
              .join("")
          : hex[1];
      return `#${body.toUpperCase()}`;
    }
    const rgb = value.match(
      /^rgba?\(\s*(\d{1,3}(?:\.\d+)?)\s*,\s*(\d{1,3}(?:\.\d+)?)\s*,\s*(\d{1,3}(?:\.\d+)?)\s*(?:,\s*(?:0|1|0?\.\d+)\s*)?\)$/i,
    );
    if (rgb) {
      const channels = rgb.slice(1, 4).map(rgbChannel);
      if (channels.every((channel): channel is number => channel !== null)) {
        return `#${channels.map(channelHex).join("")}`;
      }
    }
  }
  return /^#[0-9A-Fa-f]{6}$/.test(fallback)
    ? fallback.toUpperCase()
    : "#FCD34D";
}

/**
 * Convert the selection's displayed page percentages through pdf.js' active
 * (rotation-aware) viewport into default PDF user-space QuadPoints.
 */
export function textMarkupGeometryFromPercentRects(
  rects: readonly PercentTextRect[],
  viewport: PdfPointViewport,
  subtype: ManagedPdfTextMarkupSubtype,
  quote?: string,
): ManagedPdfTextMarkup | null {
  if (
    rects.length < 1 ||
    rects.length > 64 ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height) ||
    viewport.width <= 0 ||
    viewport.height <= 0
  ) {
    return null;
  }

  const quadPoints: number[] = [];
  for (const rect of rects) {
    if (
      !Number.isFinite(rect.x) ||
      !Number.isFinite(rect.y) ||
      !Number.isFinite(rect.width) ||
      !Number.isFinite(rect.height) ||
      rect.x < 0 ||
      rect.y < 0 ||
      rect.width <= 0 ||
      rect.height <= 0 ||
      rect.x + rect.width > 100 ||
      rect.y + rect.height > 100
    ) {
      return null;
    }
    const left = (rect.x / 100) * viewport.width;
    const top = (rect.y / 100) * viewport.height;
    const right = ((rect.x + rect.width) / 100) * viewport.width;
    const bottom = ((rect.y + rect.height) / 100) * viewport.height;
    const points = [
      viewport.convertToPdfPoint(left, top),
      viewport.convertToPdfPoint(right, top),
      viewport.convertToPdfPoint(left, bottom),
      viewport.convertToPdfPoint(right, bottom),
    ];
    if (
      points.some(
        (point) =>
          !Array.isArray(point) ||
          point.length !== 2 ||
          point.some((coordinate) => !Number.isFinite(coordinate)),
      )
    ) {
      return null;
    }
    for (const [x, y] of points) {
      quadPoints.push(roundPdfCoordinate(x), roundPdfCoordinate(y));
    }
  }

  const xs = quadPoints.filter((_value, index) => index % 2 === 0);
  const ys = quadPoints.filter((_value, index) => index % 2 === 1);
  const padding = subtype === "Highlight" ? 0 : 1.5;
  const rect: [number, number, number, number] = [
    roundPdfCoordinate(Math.min(...xs) - padding),
    roundPdfCoordinate(Math.min(...ys) - padding),
    roundPdfCoordinate(Math.max(...xs) + padding),
    roundPdfCoordinate(Math.max(...ys) + padding),
  ];
  return (
    sanitizeManagedPdfTextMarkup({
      subtype,
      rect,
      quadPoints,
      ...(quote ? { quote } : {}),
    }) ?? null
  );
}

/** A viewport that maps PDF user space to its own (rotation-aware) pixels. */
export interface PdfViewportPointMap {
  width: number;
  height: number;
  convertToViewportPoint(x: number, y: number): [number, number];
}

/**
 * The page percentages a managed text markup covers: each QuadPoints
 * quadrilateral through the (rotation-aware) viewport, as the box it spans,
 * clamped to the page. The inverse of textMarkupGeometryFromPercentRects, for
 * a reader's row written again from the page's record of the mark.
 */
export function percentRectsFromTextMarkupGeometry(
  markup: ManagedPdfTextMarkup,
  viewport: PdfViewportPointMap,
): PercentTextRect[] {
  if (!Number.isFinite(viewport.width) || !Number.isFinite(viewport.height) ||
    viewport.width <= 0 || viewport.height <= 0) return [];
  const rects: PercentTextRect[] = [];
  const clamp = (value: number) => Math.min(100, Math.max(0, value));
  for (let index = 0; index + 7 < markup.quadPoints.length && rects.length < 64; index += 8) {
    const points = [0, 2, 4, 6].map((offset) =>
      viewport.convertToViewportPoint(markup.quadPoints[index + offset], markup.quadPoints[index + offset + 1]));
    if (points.some((point) => !Array.isArray(point) || point.some((value) => !Number.isFinite(value)))) continue;
    const xs = points.map(([x]) => (x / viewport.width) * 100);
    const ys = points.map(([, y]) => (y / viewport.height) * 100);
    const x = clamp(Math.min(...xs));
    const y = clamp(Math.min(...ys));
    const width = clamp(Math.max(...xs)) - x;
    const height = clamp(Math.max(...ys)) - y;
    if (width > 0 && height > 0) rects.push({ x, y, width, height });
  }
  return rects;
}

export function proofingReplaceGeometryFromPercentRects(
  rects: readonly PercentTextRect[],
  viewport: PdfPointViewport,
  quote?: string,
): ManagedPdfProofingAnnotation | null {
  const strikeout = textMarkupGeometryFromPercentRects(
    rects,
    viewport,
    "StrikeOut",
    quote,
  );
  if (!strikeout) return null;
  const [x0, y0, x1, y1] = strikeout.rect;
  const width = Math.min(Math.max((y1 - y0) * 0.55, 4), x1 - x0);
  return (
    sanitizeManagedPdfProofingAnnotation({
      kind: "replace",
      rect: [x1 - width, y0, x1, y1],
      quadPoints: strikeout.quadPoints,
      quote,
    }) ?? null
  );
}

export function proofingInsertGeometryFromPercentPoint(
  input: { x: number; top: number; bottom: number },
  viewport: PdfPointViewport,
): ManagedPdfProofingAnnotation | null {
  if (
    !Number.isFinite(input.x) ||
    !Number.isFinite(input.top) ||
    !Number.isFinite(input.bottom) ||
    input.x < 0 ||
    input.x > 100 ||
    input.top < 0 ||
    input.bottom > 100 ||
    input.bottom <= input.top ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height) ||
    viewport.width <= 0 ||
    viewport.height <= 0
  ) {
    return null;
  }
  const x = (input.x / 100) * viewport.width;
  const top = (input.top / 100) * viewport.height;
  const bottom = (input.bottom / 100) * viewport.height;
  const [topX, topY] = viewport.convertToPdfPoint(x, top);
  const [bottomX, bottomY] = viewport.convertToPdfPoint(x, bottom);
  if (![topX, topY, bottomX, bottomY].every(Number.isFinite)) return null;
  const height = Math.max(Math.abs(bottomY - topY), Math.abs(bottomX - topX));
  const halfWidth = Math.max(2, Math.min(6, height * 0.28));
  const left = Math.min(topX, bottomX) - halfWidth;
  const right = Math.max(topX, bottomX) + halfWidth;
  const low = Math.min(topY, bottomY);
  const high = Math.max(topY, bottomY);
  return (
    sanitizeManagedPdfProofingAnnotation({
      kind: "insert",
      rect: [left, low, right, high],
    }) ?? null
  );
}

/**
 * The proofing tools' geometry from a selection's QuadPoints (PDF user
 * space: top-left, top-right, bottom-left, bottom-right per line), as the
 * Copilot binds a selection: Replace Text strikes the lines out with the
 * caret at their end (proofingReplaceGeometryFromPercentRects); Insert Text
 * puts the caret after the last line's end, or before the first line's
 * start (proofingInsertGeometryFromPercentPoint at that edge).
 */
export function proofingGeometryFromQuadPoints(
  kind: ManagedPdfProofingKind,
  quadPoints: readonly number[],
  insertAt: "after" | "before" = "after",
  quote?: string,
): ManagedPdfProofingAnnotation | null {
  if (
    quadPoints.length < 8 ||
    quadPoints.length % 8 !== 0 ||
    quadPoints.length > MAX_QUAD_POINT_VALUES ||
    !quadPoints.every((value) => Number.isFinite(value))
  ) {
    return null;
  }
  if (kind === "replace") {
    const xs = quadPoints.filter((_value, index) => index % 2 === 0);
    const ys = quadPoints.filter((_value, index) => index % 2 === 1);
    const [x0, y0, x1, y1] = [
      Math.min(...xs) - 1.5,
      Math.min(...ys) - 1.5,
      Math.max(...xs) + 1.5,
      Math.max(...ys) + 1.5,
    ];
    const width = Math.min(Math.max((y1 - y0) * 0.55, 4), x1 - x0);
    return (
      sanitizeManagedPdfProofingAnnotation({
        kind: "replace",
        rect: [x1 - width, y0, x1, y1],
        quadPoints: quadPoints.map(roundPdfCoordinate),
        quote,
      }) ?? null
    );
  }
  const at = insertAt === "before" ? 0 : quadPoints.length - 8;
  // The line's top and bottom at that edge: TL / BL before, TR / BR after.
  const top = insertAt === "before" ? at : at + 2;
  const bottom = insertAt === "before" ? at + 4 : at + 6;
  const [topX, topY] = [quadPoints[top], quadPoints[top + 1]];
  const [bottomX, bottomY] = [quadPoints[bottom], quadPoints[bottom + 1]];
  const height = Math.max(Math.abs(bottomY - topY), Math.abs(bottomX - topX));
  const halfWidth = Math.max(2, Math.min(6, height * 0.28));
  return (
    sanitizeManagedPdfProofingAnnotation({
      kind: "insert",
      rect: [
        Math.min(topX, bottomX) - halfWidth,
        Math.min(topY, bottomY),
        Math.max(topX, bottomX) + halfWidth,
        Math.max(topY, bottomY),
      ],
      quote,
    }) ?? null
  );
}

export function managedTextMarkupMatchesNative(
  managed: ManagedPdfTextMarkup,
  native: {
    subtype: ManagedPdfTextMarkupSubtype;
    rect: [number, number, number, number];
    text?: string;
    color?: string;
  },
  expected: { text?: string; color?: string },
): boolean {
  if (managed.subtype !== native.subtype) return false;
  if (
    managed.rect.some(
      (coordinate, index) => Math.abs(coordinate - native.rect[index]) > 0.05,
    )
  ) {
    return false;
  }
  if ((expected.text ?? "").trim() !== (native.text ?? "").trim()) return false;
  return (
    pdfMarkupColorToHex(expected.color) === pdfMarkupColorToHex(native.color)
  );
}
