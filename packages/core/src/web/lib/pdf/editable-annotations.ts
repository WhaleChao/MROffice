/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Export annotations as REAL PDF annotation objects instead of baked pixels.
 *
 * Everything except links currently flattens into the page content stream. That
 * is the right default for "here is the marked-up document", but it is a dead
 * end for anyone who opens the file in Acrobat, WPS or Preview: a highlight is
 * no longer a highlight, its comment is gone, a sticky note has no popup, and
 * nothing can be moved, recoloured or deleted. Round-tripping with a supervisor
 * who marks up in Acrobat is impossible.
 *
 * This module turns the same inputs the flatten planner reads into annotation
 * SPECS — plain data describing /Annots dictionaries, including the appearance
 * stream each one needs. It is deliberately free of pdf-lib so the geometry, the
 * QuadPoints order and the appearance operators are unit-testable on their own;
 * `writeEditableAnnotations` (in flatten-apply) turns a spec into objects.
 *
 * Two details that are easy to get wrong and are pinned by tests:
 *
 *  - QuadPoints are NOT in reading order. The spec's own diagram is
 *    (x1,y1) (x2,y2) (x3,y3) (x4,y4) = upper-left, upper-RIGHT, lower-left,
 *    lower-right — the top edge first, then the bottom edge — which is the
 *    opposite of the Z the name suggests. A viewer given the intuitive order
 *    draws a bow-tie.
 *  - A highlight's appearance has to multiply, or it paints over the text it is
 *    meant to highlight. The form's ExtGState carries /BM /Multiply, and the
 *    alpha lives there too rather than in the annotation's /CA, so a viewer that
 *    regenerates appearances and one that trusts ours agree.
 */

import {
  highlightQuadMarkupGeometry,
  normalizeHighlightQuads,
  type HighlightMarkupStyle,
} from "./highlight-markup";
import {
  normalizeColor,
  parseHexColor,
  percentHighlightQuadToUserSpace,
  percentPointToUserSpace,
  percentRectToUserSpace,
  renderedPageBox,
  type FlattenAnnotations,
  type FlattenPageSize,
  type FlattenRect,
} from "./flatten";
import {
  cloudOutlinePoints,
  resolveShapeCornerRadius,
  resolveShapeFill,
  resolveShapeLineEndings,
  resolveShapeLineStyle,
  resolveShapeOpacity,
  shapeLineDashPattern,
  shapeLineMarkerGeometry,
  shapeLineMarkerSize,
  type ShapeLineDashPattern,
  type ShapeLineEnding,
  type ShapePoint,
} from "./shapes";

/** PDF annotation subtypes this exporter can emit. */
export type EditableAnnotationSubtype =
  | "Highlight"
  | "Underline"
  | "StrikeOut"
  | "Squiggly"
  | "Text"
  | "Ink"
  | "Square"
  | "Circle"
  | "Line"
  | "Polygon"
  | "PolyLine";

/** The appearance stream (/AP /N form) for one annotation. */
export interface EditableAppearance {
  /** Content-stream operators, in the form's own coordinate space. */
  stream: string;
  /** The form's /BBox, matching the operators' space. */
  bbox: [number, number, number, number];
  /** Emit an ExtGState with /BM /Multiply (markup fills sit UNDER the glyphs). */
  multiply?: boolean;
  /** Constant alpha for that ExtGState, when below 1. */
  alpha?: number;
  /** Shape-only final stroking/nonstroking alpha for `/GSStroke`. */
  strokeAlpha?: number;
  /** Shape-only final nonstroking alpha for `/GSFill`. */
  fillAlpha?: number;
}

/** One annotation to write, plus everything its dictionary needs. */
export interface EditableAnnotationSpec {
  /** 0-based index of the page in the document being written. */
  pageIndex: number;
  subtype: EditableAnnotationSubtype;
  /** /Rect as [x1, y1, x2, y2] in unrotated user space. */
  rect: [number, number, number, number];
  /** /C colour, components in 0..1. */
  color: [number, number, number];
  /** /Contents — the comment a reader typed, not the highlighted text. */
  contents?: string;
  /** /T — the annotation's author. */
  author?: string;
  /**
   * /NM — a name that stays stable across exports. Ours are the annotation ids,
   * which is what makes an XFDF merge idempotent and lets a re-export replace
   * rather than duplicate.
   */
  name?: string;
  /** /QuadPoints, in the spec's top-edge-then-bottom-edge order. */
  quadPoints?: number[];
  /** /InkList — one flat [x y x y …] array per stroke. */
  inkList?: number[][];
  /** /L — the two endpoints of a Line annotation. */
  line?: [number, number, number, number];
  /** /Vertices — exact polygon points in PDF user space. */
  vertices?: number[];
  /** /IC — interior fill for Square / Circle, when the shape has one. */
  interiorColor?: [number, number, number];
  /** /BS /W — border width in points. */
  borderWidth?: number;
  /** /BS /D — canonical non-solid border pattern. */
  borderDashArray?: [number, number];
  /** Rounded Square radius in PDF user-space points. */
  borderCornerRadius?: number;
  /** Exact physical shape box used to validate the radius (not padded /Rect). */
  borderCornerBox?: [number, number, number, number];
  /** Canonical managed-shape fill opacity for MRAI XFDF round-trip. */
  shapeFillOpacity?: number;
  /** Canonical managed-shape rounded-corner percentage for MRAI XFDF. */
  shapeCornerRadius?: number;
  /** /LE — line endpoint styles; absent means /None. */
  lineStart?: Exclude<ShapeLineEnding, "none">;
  lineEnd?: Exclude<ShapeLineEnding, "none">;
  /** /BE /Cloudy — standard cloudy border effect for a cloud shape. */
  borderEffect?: "cloudy";
  /** /CA — constant opacity applied to the whole annotation. */
  opacity?: number;
  /** Replies, each written as its own annotation with /IRT back to this one. */
  replies?: EditableAnnotationReply[];
  appearance: EditableAppearance;
}

export interface EditableAnnotationReply {
  name?: string;
  author?: string;
  contents: string;
}

/** Sticky-note marker size in points — Acrobat's own note icon is 20x20. */
export const NOTE_ICON_SIZE = 20;

/** Opacity a highlight fill is painted with, matching the on-screen overlay. */
export const HIGHLIGHT_FILL_ALPHA = 0.4;

/** Thickness of an underline / strikeout band, as a fraction of line height. */
const BAND_THICKNESS = 0.07;

/** Amplitude and period of the squiggly wave, as fractions of line height. */
const SQUIGGLE_AMPLITUDE = 0.12;
const SQUIGGLE_PERIOD = 0.5;

type Rgb = [number, number, number];

function rgbOf(value: string | undefined, fallback: Rgb): Rgb {
  const parsed = typeof value === "string" ? parseHexColor(value) : null;
  if (!parsed) return fallback;
  return [parsed.r, parsed.g, parsed.b];
}

function fmt(n: number): string {
  // Three decimals is finer than a point is visible and keeps streams short.
  return Number.isFinite(n) ? String(Math.round(n * 1000) / 1000) : "0";
}

function colorOps(color: Rgb, op: "rg" | "RG"): string {
  return `${fmt(color[0])} ${fmt(color[1])} ${fmt(color[2])} ${op}`;
}

/**
 * QuadPoints for one rectangle, in the order the spec's figure gives:
 * upper-left, upper-right, lower-left, lower-right. Getting this wrong is the
 * classic bow-tie highlight.
 */
export function quadPointsForRect(rect: {
  x: number;
  y: number;
  width: number;
  height: number;
}): number[] {
  const x1 = rect.x;
  const x2 = rect.x + rect.width;
  const y1 = rect.y + rect.height; // top
  const y2 = rect.y; // bottom
  return [x1, y1, x2, y1, x1, y2, x2, y2];
}

/** The bounding rect of several user-space rects, as [x1, y1, x2, y2]. */
function boundsOf(
  rects: ReadonlyArray<{ x: number; y: number; width: number; height: number }>,
): [number, number, number, number] | null {
  if (rects.length === 0) return null;
  let x1 = Infinity;
  let y1 = Infinity;
  let x2 = -Infinity;
  let y2 = -Infinity;
  for (const r of rects) {
    x1 = Math.min(x1, r.x);
    y1 = Math.min(y1, r.y);
    x2 = Math.max(x2, r.x + r.width);
    y2 = Math.max(y2, r.y + r.height);
  }
  if (!Number.isFinite(x1) || x2 <= x1 || y2 <= y1) return null;
  return [x1, y1, x2, y2];
}

/** Appearance operators for a text-markup annotation over its quads. */
function markupAppearance(
  style: "highlight" | "underline" | "strikethrough" | "squiggly",
  rects: ReadonlyArray<{ x: number; y: number; width: number; height: number }>,
  color: Rgb,
  bbox: [number, number, number, number],
): EditableAppearance {
  const parts: string[] = [];
  if (style === "highlight") {
    parts.push(colorOps(color, "rg"));
    for (const r of rects) {
      parts.push(
        `${fmt(r.x)} ${fmt(r.y)} ${fmt(r.width)} ${fmt(r.height)} re`,
      );
    }
    parts.push("f");
    return {
      stream: parts.join("\n"),
      bbox,
      multiply: true,
      alpha: HIGHLIGHT_FILL_ALPHA,
    };
  }

  parts.push(colorOps(color, "RG"));
  for (const r of rects) {
    const thickness = Math.max(0.5, r.height * BAND_THICKNESS);
    parts.push(`${fmt(thickness)} w`);
    if (style === "underline" || style === "strikethrough") {
      // Underline sits just under the baseline; strikeout crosses the middle.
      const y =
        style === "underline"
          ? r.y + thickness
          : r.y + r.height / 2;
      parts.push(`${fmt(r.x)} ${fmt(y)} m`);
      parts.push(`${fmt(r.x + r.width)} ${fmt(y)} l`);
      parts.push("S");
      continue;
    }
    // Squiggly: a zig-zag along the baseline, its period and amplitude scaled
    // to the line so it reads the same on a footnote and a heading.
    const amplitude = Math.max(0.6, r.height * SQUIGGLE_AMPLITUDE);
    const period = Math.max(2, r.height * SQUIGGLE_PERIOD);
    const baseY = r.y + amplitude;
    parts.push(`${fmt(r.x)} ${fmt(baseY)} m`);
    let up = true;
    for (let x = r.x + period / 2; x < r.x + r.width; x += period / 2) {
      parts.push(`${fmt(x)} ${fmt(up ? baseY + amplitude : baseY)} l`);
      up = !up;
    }
    parts.push("S");
  }
  return { stream: parts.join("\n"), bbox };
}

/**
 * Appearance operators for a text markup over slanted quads (user space): the
 * same outline / band / wave the reader and the flatten bake draw, so a
 * viewer that trusts our appearance paints the slant, not its bounds.
 */
function quadMarkupAppearance(
  style: HighlightMarkupStyle,
  quads: ReadonlyArray<ReadonlyArray<number>>,
  color: Rgb,
  bbox: [number, number, number, number],
): EditableAppearance {
  const parts: string[] = [];
  const stroke = style === "squiggly";
  parts.push(colorOps(color, stroke ? "RG" : "rg"));
  for (const quad of quads) {
    const geometry = highlightQuadMarkupGeometry(quad, style);
    const [first, ...rest] = geometry.points;
    if (!first) continue;
    if (geometry.kind === "stroke") {
      // Line height = the quad's left side, which stays true on a slant.
      const height = Math.hypot(quad[0] - quad[4], quad[1] - quad[5]);
      parts.push(`${fmt(Math.max(0.5, height * BAND_THICKNESS))} w`);
    }
    parts.push(`${fmt(first.x)} ${fmt(first.y)} m`);
    for (const point of rest) parts.push(`${fmt(point.x)} ${fmt(point.y)} l`);
    parts.push(geometry.kind === "stroke" ? "S" : "h");
  }
  if (!stroke) parts.push("f");
  return style === "highlight"
    ? { stream: parts.join("\n"), bbox, multiply: true, alpha: HIGHLIGHT_FILL_ALPHA }
    : { stream: parts.join("\n"), bbox };
}

/** Appearance operators for a sticky note's marker. */
function noteAppearance(
  color: Rgb,
  bbox: [number, number, number, number],
): EditableAppearance {
  const w = bbox[2] - bbox[0];
  const h = bbox[3] - bbox[1];
  const inset = Math.max(1, Math.min(w, h) * 0.12);
  const x1 = bbox[0] + inset;
  const y1 = bbox[1] + inset;
  const bw = w - inset * 2;
  const bh = h - inset * 2;
  // A filled page with a folded corner and two rule lines: recognisable at 20pt
  // and cheap enough to keep the stream inline.
  const lines = [
    colorOps(color, "rg"),
    "0.15 0.15 0.15 RG",
    "0.6 w",
    `${fmt(x1)} ${fmt(y1)} ${fmt(bw)} ${fmt(bh)} re`,
    "B",
    "1 1 1 RG",
    `${fmt(x1 + bw * 0.18)} ${fmt(y1 + bh * 0.62)} m`,
    `${fmt(x1 + bw * 0.82)} ${fmt(y1 + bh * 0.62)} l`,
    "S",
    `${fmt(x1 + bw * 0.18)} ${fmt(y1 + bh * 0.38)} m`,
    `${fmt(x1 + bw * 0.62)} ${fmt(y1 + bh * 0.38)} l`,
    "S",
  ];
  return { stream: lines.join("\n"), bbox };
}

/** Appearance operators for an ink annotation: one polyline per stroke. */
function inkAppearance(
  strokes: ReadonlyArray<ReadonlyArray<{ x: number; y: number }>>,
  color: Rgb,
  width: number,
  bbox: [number, number, number, number],
  alpha: number | undefined,
): EditableAppearance {
  const parts = [colorOps(color, "RG"), `${fmt(width)} w`, "1 J", "1 j"];
  for (const stroke of strokes) {
    if (stroke.length === 0) continue;
    parts.push(`${fmt(stroke[0].x)} ${fmt(stroke[0].y)} m`);
    for (const point of stroke.slice(1)) {
      parts.push(`${fmt(point.x)} ${fmt(point.y)} l`);
    }
    parts.push("S");
  }
  return {
    stream: parts.join("\n"),
    bbox,
    // A highlighter pen is translucent; the pen is not. Carrying the alpha in
    // the appearance's ExtGState keeps the exported stroke the same weight as
    // the one on screen.
    alpha: alpha !== undefined && alpha < 1 ? alpha : undefined,
  };
}

function markerAppearanceParts(
  ending: ShapeLineEnding,
  tip: ShapePoint,
  toward: ShapePoint,
  size: number | null,
  color: Rgb,
): string[] {
  if (size === null) return [];
  const geometry = shapeLineMarkerGeometry(ending, { tip, toward, size });
  if (!geometry) return [];
  if (geometry.kind === "openArrow") {
    return [
      `${fmt(geometry.left.x)} ${fmt(geometry.left.y)} m`,
      `${fmt(geometry.tip.x)} ${fmt(geometry.tip.y)} l`,
      `${fmt(geometry.right.x)} ${fmt(geometry.right.y)} l`,
      "S",
    ];
  }
  const fill = colorOps(color, "rg");
  if (geometry.kind === "closedArrow") {
    return [
      fill,
      `${fmt(geometry.tip.x)} ${fmt(geometry.tip.y)} m`,
      `${fmt(geometry.left.x)} ${fmt(geometry.left.y)} l`,
      `${fmt(geometry.right.x)} ${fmt(geometry.right.y)} l`,
      "h",
      "B",
    ];
  }
  if (geometry.kind === "square") {
    return [
      fill,
      `${fmt(geometry.x)} ${fmt(geometry.y)} ${fmt(geometry.width)} ${fmt(geometry.height)} re`,
      "B",
    ];
  }
  const kappa = 0.5522847498307936;
  const { cx, cy, radius } = geometry;
  return [
    fill,
    `${fmt(cx + radius)} ${fmt(cy)} m`,
    `${fmt(cx + radius)} ${fmt(cy + kappa * radius)} ${fmt(cx + kappa * radius)} ${fmt(cy + radius)} ${fmt(cx)} ${fmt(cy + radius)} c`,
    `${fmt(cx - kappa * radius)} ${fmt(cy + radius)} ${fmt(cx - radius)} ${fmt(cy + kappa * radius)} ${fmt(cx - radius)} ${fmt(cy)} c`,
    `${fmt(cx - radius)} ${fmt(cy - kappa * radius)} ${fmt(cx - kappa * radius)} ${fmt(cy - radius)} ${fmt(cx)} ${fmt(cy - radius)} c`,
    `${fmt(cx + kappa * radius)} ${fmt(cy - radius)} ${fmt(cx + radius)} ${fmt(cy - kappa * radius)} ${fmt(cx + radius)} ${fmt(cy)} c`,
    "h",
    "B",
  ];
}

/** Appearance operators for a Square / Circle / Line annotation. */
function shapeAppearance(
  kind:
    | "rect"
    | "ellipse"
    | "line"
    | "arrow"
    | "cloud"
    | "polygon"
    | "polyline",
  rect: { x: number; y: number; width: number; height: number },
  color: Rgb,
  width: number,
  fill: Rgb | null,
  strokeAlpha: number,
  fillAlpha: number | undefined,
  cornerRadius: number,
  polygonPoints: readonly { x: number; y: number }[] | undefined,
  bbox: [number, number, number, number],
  linePoints: readonly ShapePoint[] | undefined,
  linePattern: ShapeLineDashPattern | null,
  lineStart: ShapeLineEnding,
  lineEnd: ShapeLineEnding,
  markerSize: number | null,
): EditableAppearance {
  const parts = [colorOps(color, "RG"), `${fmt(width)} w`];
  if (linePattern) {
    parts.push(linePattern.lineCap === "round" ? "1 J" : "0 J");
    parts.push(
      `[${linePattern.dashArray.map((value) => fmt(value)).join(" ")}] 0 d`,
    );
  }
  if (fill) parts.push(colorOps(fill, "rg"));
  const path: string[] = [];
  if (kind === "rect") {
    if (cornerRadius > 0) {
      const left = rect.x;
      const right = rect.x + rect.width;
      const bottom = rect.y;
      const top = rect.y + rect.height;
      const radius = Math.min(cornerRadius, rect.width / 2, rect.height / 2);
      const handle = radius * 0.5522847498307936;
      path.push(
        `${fmt(left + radius)} ${fmt(bottom)} m`,
        `${fmt(right - radius)} ${fmt(bottom)} l`,
        `${fmt(right - radius + handle)} ${fmt(bottom)} ${fmt(right)} ${fmt(bottom + radius - handle)} ${fmt(right)} ${fmt(bottom + radius)} c`,
        `${fmt(right)} ${fmt(top - radius)} l`,
        `${fmt(right)} ${fmt(top - radius + handle)} ${fmt(right - radius + handle)} ${fmt(top)} ${fmt(right - radius)} ${fmt(top)} c`,
        `${fmt(left + radius)} ${fmt(top)} l`,
        `${fmt(left + radius - handle)} ${fmt(top)} ${fmt(left)} ${fmt(top - radius + handle)} ${fmt(left)} ${fmt(top - radius)} c`,
        `${fmt(left)} ${fmt(bottom + radius)} l`,
        `${fmt(left)} ${fmt(bottom + radius - handle)} ${fmt(left + radius - handle)} ${fmt(bottom)} ${fmt(left + radius)} ${fmt(bottom)} c`,
        "h",
      );
    } else {
      path.push(
        `${fmt(rect.x)} ${fmt(rect.y)} ${fmt(rect.width)} ${fmt(rect.height)} re`,
      );
    }
  } else if (kind === "ellipse") {
    // Four Béziers; 0.5523 is the standard circle-to-cubic magic ratio.
    const k = 0.5523;
    const cx = rect.x + rect.width / 2;
    const cy = rect.y + rect.height / 2;
    const rx = rect.width / 2;
    const ry = rect.height / 2;
    path.push(`${fmt(cx + rx)} ${fmt(cy)} m`);
    path.push(
      `${fmt(cx + rx)} ${fmt(cy + ry * k)} ${fmt(cx + rx * k)} ${fmt(cy + ry)} ${fmt(cx)} ${fmt(cy + ry)} c`,
    );
    path.push(
      `${fmt(cx - rx * k)} ${fmt(cy + ry)} ${fmt(cx - rx)} ${fmt(cy + ry * k)} ${fmt(cx - rx)} ${fmt(cy)} c`,
    );
    path.push(
      `${fmt(cx - rx)} ${fmt(cy - ry * k)} ${fmt(cx - rx * k)} ${fmt(cy - ry)} ${fmt(cx)} ${fmt(cy - ry)} c`,
    );
    path.push(
      `${fmt(cx + rx * k)} ${fmt(cy - ry)} ${fmt(cx + rx)} ${fmt(cy - ry * k)} ${fmt(cx + rx)} ${fmt(cy)} c`,
    );
  } else if (kind === "cloud") {
    const cloud = cloudOutlinePoints({
      x1: rect.x,
      y1: rect.y,
      x2: rect.x + rect.width,
      y2: rect.y + rect.height,
    });
    if (cloud.length > 0) {
      path.push(`${fmt(cloud[0].x)} ${fmt(cloud[0].y)} m`);
      for (const point of cloud.slice(1)) {
        path.push(`${fmt(point.x)} ${fmt(point.y)} l`);
      }
      path.push("h");
    }
  } else if (kind === "polygon" && polygonPoints?.length) {
    path.push(`${fmt(polygonPoints[0].x)} ${fmt(polygonPoints[0].y)} m`);
    for (const point of polygonPoints.slice(1)) {
      path.push(`${fmt(point.x)} ${fmt(point.y)} l`);
    }
    path.push("h");
  } else if (kind === "polyline" && polygonPoints?.length) {
    path.push(`${fmt(polygonPoints[0].x)} ${fmt(polygonPoints[0].y)} m`);
    for (const point of polygonPoints.slice(1)) {
      path.push(`${fmt(point.x)} ${fmt(point.y)} l`);
    }
  } else if (linePoints?.length === 2) {
    path.push(`${fmt(linePoints[0].x)} ${fmt(linePoints[0].y)} m`);
    path.push(`${fmt(linePoints[1].x)} ${fmt(linePoints[1].y)} l`);
  } else {
    path.push(`${fmt(rect.x)} ${fmt(rect.y)} m`);
    path.push(`${fmt(rect.x + rect.width)} ${fmt(rect.y + rect.height)} l`);
  }
  const splitFill = fill !== null && fillAlpha !== strokeAlpha;
  const useStrokeState = strokeAlpha < 1 || splitFill;
  if (fill && splitFill) {
    parts.push("/GSFill gs", ...path, "f", "/GSStroke gs", ...path, "S");
  } else {
    if (useStrokeState) parts.push("/GSStroke gs");
    parts.push(...path, fill ? "B" : "S");
  }
  const markerPoints = kind === "polyline" ? polygonPoints : linePoints;
  if (markerPoints && markerPoints.length >= 2) {
    parts.push(
      ...markerAppearanceParts(
        lineStart,
        markerPoints[0],
        markerPoints[1],
        markerSize,
        color,
      ),
      ...markerAppearanceParts(
        lineEnd,
        markerPoints[markerPoints.length - 1],
        markerPoints[markerPoints.length - 2],
        markerSize,
        color,
      ),
    );
  }
  return {
    stream: parts.join("\n"),
    bbox,
    ...(useStrokeState ? { strokeAlpha } : null),
    ...(splitFill && fillAlpha !== undefined ? { fillAlpha } : null),
  };
}

/** Extra per-annotation fields the editable export needs but flattening does
 *  not — carried alongside so the flatten inputs keep their shape. */
export interface EditableAnnotationMeta {
  /** Stable id, written as /NM. */
  id?: string;
  /** The reader's comment, written as /Contents. */
  contents?: string;
  /** Display name of whoever made it, written as /T. */
  authorName?: string;
  replies?: ReadonlyArray<EditableAnnotationReply>;
}

export interface BuildEditableAnnotationsInput {
  annotations: FlattenAnnotations;
  /** Page sizes by 1-based page number, as the flatten planner receives them. */
  pageSizes: ReadonlyMap<number, FlattenPageSize>;
  /**
   * Extra fields keyed by `${kind}:${index}` where index is the annotation's
   * position in its source array — the flatten inputs carry no ids, and pairing
   * by position keeps this module from having to know the payload shapes.
   */
  meta?: ReadonlyMap<string, EditableAnnotationMeta>;
  /**
   * Maps a 1-based SOURCE page number to the 0-based index it occupies in the
   * exported document. Omitted means "unchanged order", which is what an export
   * with no page reordering wants.
   */
  pageIndexFor?: (sourcePage: number) => number | null;
}

/** Build the annotation specs for an editable-annotation export. */
export function buildEditableAnnotations(
  input: BuildEditableAnnotationsInput,
): EditableAnnotationSpec[] {
  const { annotations, pageSizes, meta, pageIndexFor } = input;
  const specs: EditableAnnotationSpec[] = [];
  const indexOf = (sourcePage: number): number | null => {
    if (pageIndexFor) return pageIndexFor(sourcePage);
    return Number.isInteger(sourcePage) && sourcePage >= 1
      ? sourcePage - 1
      : null;
  };
  const metaFor = (kind: string, index: number): EditableAnnotationMeta =>
    meta?.get(`${kind}:${index}`) ?? {};

  (annotations.highlights ?? []).forEach((highlight, index) => {
    const pageIndex = indexOf(highlight.page);
    const size = pageSizes.get(highlight.page);
    if (pageIndex === null || !size) return;
    const mapped = (highlight.rects ?? [])
      .map((rect: FlattenRect) => percentRectToUserSpace(rect, size))
      .filter(
        (r): r is { x: number; y: number; width: number; height: number } =>
          r !== null,
      );
    const bbox = boundsOf(mapped);
    if (!bbox) return;
    const style = highlight.markupStyle ?? "highlight";
    const subtype: EditableAnnotationSubtype =
      style === "underline"
        ? "Underline"
        : style === "strikethrough"
          ? "StrikeOut"
          : style === "squiggly"
            ? "Squiggly"
            : "Highlight";
    const color = rgbOf(highlight.color, [1, 0.85, 0.2]);
    const extra = metaFor("highlight", index);
    // Slanted source quads export as themselves; their bounding rects would
    
    const percentQuads =
      normalizeHighlightQuads(
        highlight.quads ?? undefined,
        (highlight.rects ?? []).length,
      ) ?? undefined;
    const userQuads =
      percentQuads && mapped.length === percentQuads.length
        ? percentQuads.map((quad) => percentHighlightQuadToUserSpace(quad, size))
        : null;
    const quadPoints: number[] = [];
    if (userQuads) for (const quad of userQuads) quadPoints.push(...quad);
    else for (const rect of mapped) quadPoints.push(...quadPointsForRect(rect));
    specs.push({
      pageIndex,
      subtype,
      rect: bbox,
      color,
      contents: extra.contents,
      author: extra.authorName,
      name: extra.id,
      quadPoints,
      replies: extra.replies ? [...extra.replies] : undefined,
      appearance: userQuads
        ? quadMarkupAppearance(style, userQuads, color, bbox)
        : markupAppearance(style, mapped, color, bbox),
    });
  });

  (annotations.notes ?? []).forEach((note, index) => {
    const pageIndex = indexOf(note.page);
    const size = pageSizes.get(note.page);
    if (pageIndex === null || !size) return;
    const point = percentPointToUserSpace(note.x, note.y, size);
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return;
    // The marker hangs BELOW its anchor point, the way the on-screen pin does.
    const bbox: [number, number, number, number] = [
      point.x,
      point.y - NOTE_ICON_SIZE,
      point.x + NOTE_ICON_SIZE,
      point.y,
    ];
    const color = rgbOf(note.color, [1, 0.82, 0.25]);
    const extra = metaFor("note", index);
    specs.push({
      pageIndex,
      subtype: "Text",
      rect: bbox,
      color,
      // A note's own text IS its comment — that is what the popup shows.
      contents: extra.contents ?? note.text,
      author: extra.authorName,
      name: extra.id,
      replies: extra.replies ? [...extra.replies] : undefined,
      appearance: noteAppearance(color, bbox),
    });
  });

  (annotations.inks ?? []).forEach((ink, index) => {
    const pageIndex = indexOf(ink.page);
    const size = pageSizes.get(ink.page);
    if (pageIndex === null || !size) return;
    const points = (ink.points ?? [])
      .filter((p) => p && Number.isFinite(p.x) && Number.isFinite(p.y))
      .map((p) => percentPointToUserSpace(p.x, p.y, size));
    if (points.length < 2) return;
    // Stroke width is stored as a percentage of the page; convert against the
    // shorter RENDERED dimension exactly as the baking planner does, so the two
    // export modes put down the same thickness.
    const box = renderedPageBox(size);
    const shorter = Math.min(box.width, box.height);
    const widthPct = Number.isFinite(ink.width) ? ink.width : 0;
    const thickness = Math.max(0.25, (widthPct / 100) * shorter);
    const { color, opacity } = normalizeColor(ink.color, 1);
    const rgb: Rgb = [color.r, color.g, color.b];
    const alpha = Number.isFinite(ink.opacity) ? ink.opacity : opacity;
    // The /Rect has to contain the stroke INCLUDING its width, or a viewer that
    // clips to the rect shaves the outer half of every stroke.
    const pad = thickness / 2 + 1;
    const xs = points.map((p) => p.x);
    const ys = points.map((p) => p.y);
    const bbox: [number, number, number, number] = [
      Math.min(...xs) - pad,
      Math.min(...ys) - pad,
      Math.max(...xs) + pad,
      Math.max(...ys) + pad,
    ];
    const extra = metaFor("ink", index);
    specs.push({
      pageIndex,
      subtype: "Ink",
      rect: bbox,
      color: rgb,
      borderWidth: thickness,
      opacity: alpha !== undefined && alpha < 1 ? alpha : undefined,
      contents: extra.contents,
      author: extra.authorName,
      name: extra.id,
      inkList: [points.flatMap((p) => [p.x, p.y])],
      replies: extra.replies ? [...extra.replies] : undefined,
      appearance: inkAppearance([points], rgb, thickness, bbox, alpha),
    });
  });

  (annotations.shapes ?? []).forEach((shape, index) => {
    const pageIndex = indexOf(shape.page);
    const size = pageSizes.get(shape.page);
    if (pageIndex === null || !size) return;
    const a = percentPointToUserSpace(shape.x1, shape.y1, size);
    const b = percentPointToUserSpace(shape.x2, shape.y2, size);
    if (![a.x, a.y, b.x, b.y].every(Number.isFinite)) return;
    const box = renderedPageBox(size);
    const shorter = Math.min(box.width, box.height);
    const widthPct = Number.isFinite(shape.width) ? shape.width : 0;
    const thickness = Math.max(0.25, (widthPct / 100) * shorter);
    const linePattern = shapeLineDashPattern(
      resolveShapeLineStyle(shape),
      thickness,
    );
    const lineEndings = resolveShapeLineEndings(shape);
    const markerSize = shapeLineMarkerSize(thickness);
    const { color } = normalizeColor(shape.color, 1);
    const rgb: Rgb = [color.r, color.g, color.b];
    const polygonPoints =
      (shape.kind === "polygon" || shape.kind === "polyline") && shape.points
        ? shape.points.map((point) =>
            percentPointToUserSpace(point.x, point.y, size),
          )
        : undefined;
    if (
      (shape.kind === "polygon" || shape.kind === "polyline") &&
      (!polygonPoints ||
        polygonPoints.length < (shape.kind === "polygon" ? 3 : 2) ||
        polygonPoints.some((point) => !Number.isFinite(point.x + point.y)))
    ) {
      return;
    }
    const rect = {
      x: Math.min(a.x, b.x),
      y: Math.min(a.y, b.y),
      width: Math.abs(b.x - a.x),
      height: Math.abs(b.y - a.y),
    };
    const hasMarker =
      lineEndings.start !== "none" || lineEndings.end !== "none";
    const pad = Math.max(
      thickness / 2 + 1,
      hasMarker && markerSize !== null
        ? markerSize * 0.55 + thickness / 2 + 1
        : 0,
    );
    const bbox: [number, number, number, number] = [
      rect.x - pad,
      rect.y - pad,
      rect.x + rect.width + pad,
      rect.y + rect.height + pad,
    ];
    // An arrow has no PDF subtype of its own; it exports as a Line, which is
    // what carries the endpoints a viewer needs to keep it editable. Its head is
    // in the appearance only — better an editable line than a baked arrow.
    const subtype: EditableAnnotationSubtype =
      shape.kind === "rect" || shape.kind === "cloud"
        ? "Square"
        : shape.kind === "polygon"
          ? "Polygon"
        : shape.kind === "polyline"
          ? "PolyLine"
        : shape.kind === "ellipse"
          ? "Circle"
          : "Line";
    const wholeOpacity = resolveShapeOpacity(shape.opacity);
    const resolvedFill = resolveShapeFill(shape);
    const fill = resolvedFill ? rgbOf(resolvedFill.color, rgb) : null;
    const finalFillAlpha = resolvedFill
      ? wholeOpacity * resolvedFill.opacity
      : undefined;
    const cornerPercent = resolveShapeCornerRadius(
      shape.kind,
      shape.cornerRadius,
    );
    const cornerRadius = Math.min(
      rect.width / 2,
      rect.height / 2,
      (Math.min(rect.width, rect.height) * cornerPercent) / 100,
    );
    const extra = metaFor("shape", index);
    specs.push({
      pageIndex,
      subtype,
      rect: bbox,
      color: rgb,
      interiorColor: fill ?? undefined,
      borderWidth: thickness,
      borderDashArray: linePattern?.dashArray,
      ...(cornerRadius > 0
        ? {
            borderCornerRadius: cornerRadius,
            borderCornerBox: [
              rect.x,
              rect.y,
              rect.x + rect.width,
              rect.y + rect.height,
            ] as [number, number, number, number],
            shapeCornerRadius: cornerPercent,
          }
        : null),
      ...(resolvedFill ? { shapeFillOpacity: resolvedFill.opacity } : null),
      lineStart: lineEndings.start === "none" ? undefined : lineEndings.start,
      lineEnd: lineEndings.end === "none" ? undefined : lineEndings.end,
      borderEffect: shape.kind === "cloud" ? "cloudy" : undefined,
      opacity: wholeOpacity < 1 ? wholeOpacity : undefined,
      line:
        subtype === "Line" ? [a.x, a.y, b.x, b.y] : undefined,
      vertices: polygonPoints?.flatMap((point) => [point.x, point.y]),
      contents: extra.contents,
      author: extra.authorName,
      name: extra.id,
      replies: extra.replies ? [...extra.replies] : undefined,
      appearance: shapeAppearance(
        shape.kind,
        rect,
        rgb,
        thickness,
        fill,
        wholeOpacity,
        finalFillAlpha,
        cornerRadius,
        polygonPoints,
        bbox,
        (shape.kind === "line" || shape.kind === "arrow") &&
          (linePattern || hasMarker)
          ? [a, b]
          : undefined,
        linePattern,
        lineEndings.start,
        lineEndings.end,
        markerSize,
      ),
    });
  });

  return specs;
}
