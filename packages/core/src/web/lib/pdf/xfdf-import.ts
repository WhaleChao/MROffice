/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Turn parsed XFDF annotations back into the reader's own annotation model.
 *
 * Export is the easy direction: the viewer's percent geometry maps forward into
 * user space. Import has to come back, which needs the exact inverse of
 * `percentPointToUserSpace` — including its page-rotation cases, where a 90 or
 * 270 rotation swaps the axes. The inverse is proven the only way an inverse
 * can honestly be proven: by round-tripping against the forward function on
 * every rotation, rather than by re-deriving the algebra and hoping.
 *
 * What comes back is deliberately narrow. Text markup, sticky notes, Ink,
 * closed outline-only Polygon objects and open PolyLine objects map to exact
 * reader models. PolyLine stays open throughout import, editing and native PDF
 * writeback; it is never changed into Ink or a closed Polygon.
 */

import { pageOrigin, renderedPageBox, type FlattenPageSize } from "./flatten";
import {
  XFDF_SOURCE_ID_MAX_LENGTH,
  normalizeXfdfSourceId,
  type ParsedXfdfAnnotation,
} from "./xfdf";
import type { EditableAnnotationReply } from "./editable-annotations";
import {
  buildShapePayload,
  normalizePolygonPoints,
  normalizePolylinePoints,
  type ShapeAnnotation,
  type ShapeKind,
  type ShapeLineEnding,
  type ShapeLineStyle,
  type ShapePoint,
} from "./shapes";

import { isAxisAlignedHighlightQuad } from "./highlight-markup";

export { normalizeXfdfSourceId } from "./xfdf";

/** Round-off at a page edge must not fail the stored model's 0..100 range. */
function clampPercent(value: number): number {
  return Math.min(100, Math.max(0, value));
}

/** Inverse of `percentPointToUserSpace`: user space back to page percent. */
export function userSpacePointToPercent(
  userX: number,
  userY: number,
  page: FlattenPageSize,
): { x: number; y: number } {
  const { width: rw, height: rh, rotation } = renderedPageBox(page);
  const mediaW = Number.isFinite(page.width) ? page.width : 0;
  const mediaH = Number.isFinite(page.height) ? page.height : 0;
  // A user-space point is measured from the crop box's own origin; the
  // percent model starts at the visible box's corner. Ignoring the origin
  // put a note on a page cropped to [100 150 ...] 100 pt right and 150 pt up
  
  const origin = pageOrigin(page);
  const x = userX - origin.x;
  const y = userY - origin.y;
  // Undo the rotation to land back in the rendered box (bottom-left origin).
  let rx: number;
  let ry: number;
  switch (rotation) {
    case 90:
      // Forward (pdf.js): (rx, ryTop) = (my, mx); ry is from the bottom.
      rx = y;
      ry = rh - x;
      break;
    case 180:
      rx = mediaW - x;
      ry = mediaH - y;
      break;
    case 270:
      // Forward: (rx, ryTop) = (rw - my, rh - mx), so ry = mx.
      rx = rw - y;
      ry = x;
      break;
    default:
      rx = x;
      ry = y;
  }
  const yTop = rh - ry;
  return {
    x: rw > 0 ? (rx / rw) * 100 : 0,
    y: rh > 0 ? (yTop / rh) * 100 : 0,
  };
}

/** Inverse of `percentRectToUserSpace` for an axis-aligned user-space rect. */
export function userSpaceRectToPercent(
  rect: readonly [number, number, number, number],
  page: FlattenPageSize,
): { x: number; y: number; width: number; height: number } | null {
  const [x1, y1, x2, y2] = rect;
  if (![x1, y1, x2, y2].every((n) => Number.isFinite(n))) return null;
  // Map the two opposite corners back and take the bounding box: a 90/270
  // rotation swaps which corner is which, so "top-left" is only knowable after
  // the inverse has been applied.
  const a = userSpacePointToPercent(Math.min(x1, x2), Math.max(y1, y2), page);
  const b = userSpacePointToPercent(Math.max(x1, x2), Math.min(y1, y2), page);
  const left = Math.min(a.x, b.x);
  const top = Math.min(a.y, b.y);
  const width = Math.abs(b.x - a.x);
  const height = Math.abs(b.y - a.y);
  if (!(width > 0) || !(height > 0)) return null;
  return { x: left, y: top, width, height };
}

/** A highlight recovered from XFDF, in the reader's percent model. */
export interface ImportedHighlight {
  kind: "highlight";
  id?: string;
  page: number;
  color?: string;
  markupStyle: "highlight" | "underline" | "strikethrough" | "squiggly";
  rects: Array<{ x: number; y: number; width: number; height: number }>;
  




  quads?: number[][];
  note?: string;
  authorName?: string;
  replies: EditableAnnotationReply[];
}

/** A sticky note recovered from XFDF, in the reader's percent model. */
export interface ImportedNote {
  kind: "note";
  id?: string;
  page: number;
  x: number;
  y: number;
  text: string;
  color?: string;
  authorName?: string;
  replies: EditableAnnotationReply[];
}

/**
 * An ink gesture recovered from XFDF, flattened to the reader's stroke model.
 *
 * XFDF carries plain polylines; the reader's own strokes carry pressure and
 * smoothing an XFDF file cannot supply, so an imported stroke is a uniform-
 * width polyline — honest about what the file actually said, rather than
 * inventing pressure it never recorded.
 */
export interface ImportedInk {
  kind: "ink";
  id?: string;
  page: number;
  color?: string;
  /** Stroke width in PDF points (the XFDF `width` attribute), default 2. */
  width: number;
  /** One flattened [x0,y0,x1,y1,…] per gesture, page fractions (0..1). */
  strokes: number[][];
  authorName?: string;
}

type ImportedShapeKind = Extract<
  ShapeKind,
  "rect" | "ellipse" | "line" | "polygon" | "polyline"
>;

/** A drawing markup recovered without inventing a different shape type. */
export interface ImportedShape {
  kind: "shape";
  shapeKind: ImportedShapeKind;
  id?: string;
  page: number;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  /** Stroke width in page-percent units, matching the managed shape model. */
  width: number;
  points?: ShapePoint[];
  fillColor?: string;
  fillOpacity?: number;
  opacity?: number;
  cornerRadius?: number;
  lineStyle?: ShapeLineStyle;
  lineStart?: ShapeLineEnding;
  lineEnd?: ShapeLineEnding;
}

/** True only when XFDF can retain every persisted managed-shape property. */
export function canRoundTripShapeThroughXfdf(
  shape: Pick<
    ShapeAnnotation,
    | "kind"
    | "purpose"
    | "fillColor"
    | "fillOpacity"
    | "opacity"
    | "cornerRadius"
  >,
): boolean {
  return (
    shape.purpose === undefined &&
    ["rect", "ellipse", "line", "arrow", "polygon", "polyline"].includes(
      shape.kind,
    )
  );
}

export type ImportedAnnotation =
  | ImportedHighlight
  | ImportedNote
  | ImportedInk
  | ImportedShape;

function stableXfdfHash(value: string): string {
  let hash = 0xcbf29ce484222325n;
  const prime = 0x100000001b3n;
  const mask = 0xffffffffffffffffn;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= BigInt(value.charCodeAt(index));
    hash = (hash * prime) & mask;
  }
  return hash.toString(16).padStart(16, "0");
}

/**
 * Stable import authority carried through server-generated annotation ids.
 * Named XFDF annotations keep their native `/NM` identity. For an id-less
 * annotation, the semantic fingerprint is paired with its zero-based
 * occurrence among identical siblings: two deliberately identical comments
 * both survive, while importing the same file again resolves to the same pair.
 */
export function xfdfSourceIdentity(
  item: ImportedAnnotation,
  duplicateOccurrence = 0,
): string {
  const named = normalizeXfdfSourceId(item.id);
  if (named) return named;
  const canonical =
    item.kind === "highlight"
      ? JSON.stringify({
          kind: item.kind,
          page: item.page,
          color: item.color ?? null,
          markupStyle: item.markupStyle,
          rects: item.rects,
          ...(item.quads ? { quads: item.quads } : {}),
          note: item.note ?? null,
          authorName: item.authorName ?? null,
          replies: item.replies,
        })
      : item.kind === "note"
        ? JSON.stringify({
            kind: item.kind,
            page: item.page,
            x: item.x,
            y: item.y,
            text: item.text,
            color: item.color ?? null,
            authorName: item.authorName ?? null,
            replies: item.replies,
          })
        : item.kind === "ink"
          ? JSON.stringify({
              kind: item.kind,
              page: item.page,
              color: item.color ?? null,
              width: item.width,
              strokes: item.strokes,
              authorName: item.authorName ?? null,
            })
          : JSON.stringify({
              kind: item.kind,
              shapeKind: item.shapeKind,
              page: item.page,
              color: item.color ?? null,
              width: item.width,
              x1: item.x1,
              y1: item.y1,
              x2: item.x2,
              y2: item.y2,
              points: item.points,
              fillColor: item.fillColor ?? null,
              fillOpacity: item.fillOpacity ?? null,
              opacity: item.opacity ?? null,
              cornerRadius: item.cornerRadius ?? null,
              lineStyle: item.lineStyle ?? "solid",
              lineStart: item.lineStart ?? "none",
              lineEnd: item.lineEnd ?? "none",
            });
  const base = `mrai-xfdf-v1-${stableXfdfHash(canonical)}`;
  return duplicateOccurrence > 0 ? `${base}:${duplicateOccurrence}` : base;
}

/** Stable child identity for each gesture inside one XFDF Ink annotation. */
export function xfdfInkGestureSourceIdentity(
  annotationIdentity: string,
  gestureIndex: number,
): string {
  const parent = normalizeXfdfSourceId(annotationIdentity);
  if (!parent) throw new Error("Invalid XFDF annotation identity");
  if (!Number.isSafeInteger(gestureIndex) || gestureIndex < 0) {
    throw new Error("Invalid XFDF ink gesture index");
  }
  if (gestureIndex === 0) return parent;
  const suffix = `:ink:${gestureIndex}`;
  if (parent.length + suffix.length <= XFDF_SOURCE_ID_MAX_LENGTH) {
    return `${parent}${suffix}`;
  }
  return `mrai-xfdf-ink-v1-${stableXfdfHash(parent)}:${gestureIndex}`;
}

export interface XfdfImportResult {
  annotations: ImportedAnnotation[];
  /** Entries that parsed but could not be converted — reported, never silent. */
  skipped: number;
}

function rgbaOf(
  color: readonly [number, number, number] | undefined,
  opacity: number | undefined,
): string | undefined {
  if (!color || typeof opacity !== "number" || !Number.isFinite(opacity) || opacity >= 1) {
    return undefined;
  }
  const channel = (value: number) =>
    Math.max(0, Math.min(255, Math.round(value * 255)));
  const alpha = Math.max(0, Math.min(1, opacity));
  return `rgba(${channel(color[0])}, ${channel(color[1])}, ${channel(color[2])}, ${Number(alpha.toFixed(3))})`;
}

function hexOf(color: readonly [number, number, number] | undefined): string | undefined {
  if (!color) return undefined;
  const part = (v: number) =>
    Math.round(Math.max(0, Math.min(1, v)) * 255)
      .toString(16)
      .padStart(2, "0");
  return `#${part(color[0])}${part(color[1])}${part(color[2])}`;
}

const XFDF_LINE_ENDINGS: Record<string, ShapeLineEnding> = {
  none: "none",
  openarrow: "openArrow",
  closedarrow: "closedArrow",
  circle: "circle",
  square: "square",
};

function xfdfLineEnding(value: string | undefined): ShapeLineEnding | null {
  if (value === undefined) return "none";
  return XFDF_LINE_ENDINGS[value.toLowerCase()] ?? null;
}

function xfdfLineStyle(
  dash: readonly number[] | undefined,
  width: number,
): ShapeLineStyle | null {
  if (dash === undefined) return "solid";
  if (
    dash.length < 1 ||
    dash.length > 16 ||
    !dash.every(
      (value) => Number.isFinite(value) && value > 0 && value <= 1_000,
    )
  ) {
    return null;
  }
  return dash[0] <= width * 1.5 ? "dotted" : "dashed";
}

const STYLE_FOR = {
  Highlight: "highlight",
  Underline: "underline",
  StrikeOut: "strikethrough",
  Squiggly: "squiggly",
} as const;

/**
 * Convert parsed XFDF into the reader's annotations.
 *
 * `pageSizes` is keyed by 1-based page number. An annotation on a page the
 * document does not have is skipped and counted rather than clamped onto page
 * one — a comment silently moved to the wrong page is worse than one reported
 * as not imported.
 */
export function importXfdfAnnotations(
  parsed: ReadonlyArray<ParsedXfdfAnnotation & { replies: EditableAnnotationReply[] }>,
  pageSizes: ReadonlyMap<number, FlattenPageSize>,
): XfdfImportResult {
  const annotations: ImportedAnnotation[] = [];
  let skipped = 0;
  for (const item of parsed) {
    // XFDF numbers pages from 0; the reader's model is 1-based.
    const pageNumber = item.pageIndex + 1;
    const size = pageSizes.get(pageNumber);
    if (!size) {
      skipped += 1;
      continue;
    }
    if (item.subtype === "Text") {
      const point = userSpacePointToPercent(item.rect[0], item.rect[3], size);
      annotations.push({
        kind: "note",
        id: item.name,
        page: pageNumber,
        x: point.x,
        y: point.y,
        text: item.contents ?? "",
        color: hexOf(item.color),
        authorName: item.author,
        replies: item.replies,
      });
      continue;
    }
    if (item.subtype === "Ink") {
      const gestures = (item.inkList ?? [])
        .map((flat) => {
          const points: number[] = [];
          for (let at = 0; at + 1 < flat.length; at += 2) {
            const point = userSpacePointToPercent(flat[at], flat[at + 1], size);
            // The stroke model stores fractions; XFDF import goes through the
            // same percent conversion as everything else, then scales.
            points.push(point.x / 100, point.y / 100);
          }
          return points;
        })
        .filter((points) => points.length >= 4);
      if (gestures.length === 0) {
        skipped += 1;
        continue;
      }
      annotations.push({
        kind: "ink",
        id: item.name,
        page: pageNumber,
        color: hexOf(item.color),
        width: item.borderWidth && item.borderWidth > 0 ? item.borderWidth : 2,
        strokes: gestures,
        authorName: item.author,
      });
      continue;
    }
    if (
      item.subtype === "Square" ||
      item.subtype === "Circle" ||
      item.subtype === "Line" ||
      item.subtype === "Polygon" ||
      item.subtype === "PolyLine"
    ) {
      // An explicit width of 0 is "no border" (PDF 32000-1 table 166); only
      // an absent width takes the 1 pt default. Treating 0 as unset put a red
      
      const widthPoints =
        item.borderWidth === 0
          ? 0
          : item.borderWidth && item.borderWidth > 0
            ? item.borderWidth
            : 1;
      const lineStyle = xfdfLineStyle(item.borderDashes, widthPoints);
      const lineStart = xfdfLineEnding(item.lineStartStyle);
      const lineEnd = xfdfLineEnding(item.lineEndStyle);
      const supportsEndings =
        item.subtype === "Line" || item.subtype === "PolyLine";
      // Keep only properties the managed shape model can reproduce exactly.
      // Cloudy borders and shape comment threads stay
      // reported rather than being flattened into a different-looking object.
      if (
        lineStyle === null ||
        lineStart === null ||
        lineEnd === null ||
        (!supportsEndings &&
          (item.lineStartStyle !== undefined ||
            item.lineEndStyle !== undefined)) ||
        item.contents ||
        item.author ||
        item.replies.length > 0 ||
        item.borderEffectStyle ||
        item.borderEffectIntensity !== undefined ||
        ((item.subtype === "Line" ||
          item.subtype === "Polygon" ||
          item.subtype === "PolyLine") &&
          item.interiorColor)
      ) {
        skipped += 1;
        continue;
      }
      const rendered = renderedPageBox(size);
      const shorter = Math.min(rendered.width, rendered.height);
      if (!(shorter > 0)) {
        skipped += 1;
        continue;
      }
      const width = (widthPoints / shorter) * 100;
      const color = hexOf(item.color) ?? "#000000";
      const opacity = item.opacity;
      const fillOpacity = item.interiorColor
        ? (item.shapeFillOpacity ?? 1)
        : undefined;
      const cornerRadius = item.shapeCornerRadius;
      let payload: ReturnType<typeof buildShapePayload> = null;
      if (item.subtype === "Polygon" || item.subtype === "PolyLine") {
        const rawPoints: ShapePoint[] = [];
        const vertices = item.vertices ?? [];
        for (let at = 0; at + 1 < vertices.length; at += 2) {
          rawPoints.push(
            userSpacePointToPercent(vertices[at], vertices[at + 1], size),
          );
        }
        const points =
          item.subtype === "Polygon"
            ? normalizePolygonPoints(rawPoints)
            : normalizePolylinePoints(rawPoints);
        if (points) {
          payload = buildShapePayload({
            page: pageNumber,
            kind: item.subtype === "Polygon" ? "polygon" : "polyline",
            x1: 0,
            y1: 0,
            x2: 0,
            y2: 0,
            points,
            color,
            width,
            lineStyle,
            lineStart,
            lineEnd,
            opacity,
            cornerRadius,
          });
        }
      } else if (item.subtype === "Line") {
        const line = item.line;
        if (line) {
          const start = userSpacePointToPercent(line[0], line[1], size);
          const end = userSpacePointToPercent(line[2], line[3], size);
          payload = buildShapePayload({
            page: pageNumber,
            kind: "line",
            x1: start.x,
            y1: start.y,
            x2: end.x,
            y2: end.y,
            color,
            width,
            lineStyle,
            lineStart,
            lineEnd,
            opacity,
            cornerRadius,
          });
        }
      } else {
        // Our editable writer expands /Rect by half the stroke plus one point
        // so a viewer never clips the outline. Undo that standard appearance
        // pad before rebuilding the managed centerline geometry.
        const pad = widthPoints / 2 + 1;
        const rect = userSpaceRectToPercent(
          [
            item.rect[0] + pad,
            item.rect[1] + pad,
            item.rect[2] - pad,
            item.rect[3] - pad,
          ],
          size,
        );
        if (rect) {
          payload = buildShapePayload({
            page: pageNumber,
            kind: item.subtype === "Square" ? "rect" : "ellipse",
            x1: rect.x,
            y1: rect.y,
            x2: rect.x + rect.width,
            y2: rect.y + rect.height,
            color,
            width,
            fillColor: hexOf(item.interiorColor),
            fillOpacity,
            opacity,
            cornerRadius,
            lineStyle,
          });
        }
      }
      if (!payload || !canRoundTripShapeThroughXfdf(payload)) {
        skipped += 1;
        continue;
      }
      annotations.push({
        kind: "shape",
        shapeKind: payload.kind as ImportedShapeKind,
        id: item.name,
        page: pageNumber,
        x1: payload.x1,
        y1: payload.y1,
        x2: payload.x2,
        y2: payload.y2,
        color: payload.color,
        width: payload.width,
        points: payload.points,
        fillColor: payload.fillColor,
        fillOpacity: payload.fillOpacity,
        opacity: payload.opacity,
        cornerRadius: payload.cornerRadius,
        lineStyle: payload.lineStyle,
        lineStart: payload.lineStart,
        lineEnd: payload.lineEnd,
      });
      continue;
    }

    const style = STYLE_FOR[item.subtype as keyof typeof STYLE_FOR];
    if (!style) {
      // Ink, shapes and anything else: parsed, but not something the reader's
      // own model can hold faithfully. Counted so the UI can say so.
      skipped += 1;
      continue;
    }
    // Prefer the quads — they carry each covered line — and fall back to the
    // annotation rect when a producer omitted them. A quad on slanted text is
    // kept as its four vertices; its bounding box alone would mark the
    
    const rects: Array<{ x: number; y: number; width: number; height: number }> = [];
    const percentQuads: number[][] = [];
    let slanted = false;
    const quads = item.quadPoints ?? [];
    for (let i = 0; i + 7 < quads.length; i += 8) {
      const xs = [quads[i], quads[i + 2], quads[i + 4], quads[i + 6]];
      const ys = [quads[i + 1], quads[i + 3], quads[i + 5], quads[i + 7]];
      const mapped = userSpaceRectToPercent(
        [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)],
        size,
      );
      if (!mapped) continue;
      const percentQuad: number[] = [];
      for (let vertex = 0; vertex < 4; vertex += 1) {
        const point = userSpacePointToPercent(xs[vertex], ys[vertex], size);
        percentQuad.push(clampPercent(point.x), clampPercent(point.y));
      }
      if (!isAxisAlignedHighlightQuad(quads.slice(i, i + 8))) slanted = true;
      rects.push(mapped);
      percentQuads.push(percentQuad);
    }
    if (rects.length === 0) {
      const mapped = userSpaceRectToPercent(item.rect, size);
      if (mapped) rects.push(mapped);
    }
    if (rects.length === 0) {
      skipped += 1;
      continue;
    }
    annotations.push({
      kind: "highlight",
      id: item.name,
      page: pageNumber,
      // A markup's own opacity travels in its colour, the way the editor's
      // presets do (rgba); dropping it painted a 25% highlight as a solid
      
      color: rgbaOf(item.color, item.opacity) ?? hexOf(item.color),
      markupStyle: style,
      rects,
      ...(slanted && percentQuads.length === rects.length
        ? { quads: percentQuads }
        : {}),
      note: item.contents,
      authorName: item.author,
      replies: item.replies,
    });
  }
  return { annotations, skipped };
}
