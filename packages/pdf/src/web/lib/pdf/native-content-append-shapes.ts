/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Turn a drawn shape (page-percent geometry, top-left origin as displayed)
 * into one native path block: real page content instead of a Square / Circle
 * / Line annotation. Geometry is mapped through the page's crop box and
 * /Rotate so the shape lands where the reader drew it.
 */

import { type PdfNativeContentAppendColor, type PdfNativeContentAppendPageGeometry, type PdfNativeContentAppendPathCommand, type PdfNativeContentAppendPathContent } from "./engine/native/content-append-transaction";
import { pdfNativeContentAppendColorFromHex, pdfNativeContentAppendDisplayPoint, pdfNativeContentAppendDisplaySize } from "./engine/native/content-append-commands";
import {
  DEFAULT_SHAPE_WIDTH,
  PDF_SHAPE_MAX_WIDTH,
  PDF_SHAPE_MIN_WIDTH,
  type ShapeAnnotation,
  type ShapeLineEnding,
} from "@mrnote/core/web/lib/pdf/shapes.ts";

/** Bézier circle constant. */
const KAPPA = 0.5522847498;

interface DisplayPoint {
  readonly x: number;
  readonly y: number;
}

interface Mapper {
  /** Display-space point (page units, top-left origin) to PDF user space. */
  readonly toPdf: (point: DisplayPoint) => PdfNativeContentAppendPathCommand & { op: "m" | "l" };
  readonly width: number;
  readonly height: number;
}

function mapper(geometry: PdfNativeContentAppendPageGeometry): Mapper {
  const size = pdfNativeContentAppendDisplaySize(geometry);
  return {
    width: size.width,
    height: size.height,
    toPdf: (point) => {
      const mapped = pdfNativeContentAppendDisplayPoint(
        geometry,
        point.x / size.width,
        point.y / size.height,
      );
      return { op: "l", x: mapped.x, y: mapped.y };
    },
  };
}

function moveTo(map: Mapper, point: DisplayPoint): PdfNativeContentAppendPathCommand {
  return { ...map.toPdf(point), op: "m" };
}

function lineTo(map: Mapper, point: DisplayPoint): PdfNativeContentAppendPathCommand {
  return { ...map.toPdf(point), op: "l" };
}

function curveTo(
  map: Mapper,
  control1: DisplayPoint,
  control2: DisplayPoint,
  end: DisplayPoint,
): PdfNativeContentAppendPathCommand {
  const c1 = map.toPdf(control1);
  const c2 = map.toPdf(control2);
  const target = map.toPdf(end);
  return { op: "c", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: target.x, y: target.y };
}

function percentPoint(map: Mapper, xPercent: number, yPercent: number): DisplayPoint {
  return { x: (xPercent / 100) * map.width, y: (yPercent / 100) * map.height };
}

function rectangleCommands(
  map: Mapper,
  left: number,
  top: number,
  right: number,
  bottom: number,
  cornerRadius: number,
): PdfNativeContentAppendPathCommand[] {
  const radius = Math.max(0, Math.min(cornerRadius, (right - left) / 2, (bottom - top) / 2));
  if (radius <= 0.01) {
    return [
      moveTo(map, { x: left, y: top }),
      lineTo(map, { x: right, y: top }),
      lineTo(map, { x: right, y: bottom }),
      lineTo(map, { x: left, y: bottom }),
      { op: "h" },
    ];
  }
  const k = KAPPA * radius;
  return [
    moveTo(map, { x: left + radius, y: top }),
    lineTo(map, { x: right - radius, y: top }),
    curveTo(map, { x: right - radius + k, y: top }, { x: right, y: top + radius - k }, { x: right, y: top + radius }),
    lineTo(map, { x: right, y: bottom - radius }),
    curveTo(map, { x: right, y: bottom - radius + k }, { x: right - radius + k, y: bottom }, { x: right - radius, y: bottom }),
    lineTo(map, { x: left + radius, y: bottom }),
    curveTo(map, { x: left + radius - k, y: bottom }, { x: left, y: bottom - radius + k }, { x: left, y: bottom - radius }),
    lineTo(map, { x: left, y: top + radius }),
    curveTo(map, { x: left, y: top + radius - k }, { x: left + radius - k, y: top }, { x: left + radius, y: top }),
    { op: "h" },
  ];
}

function ellipseCommands(
  map: Mapper,
  left: number,
  top: number,
  right: number,
  bottom: number,
): PdfNativeContentAppendPathCommand[] {
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2;
  const rx = (right - left) / 2;
  const ry = (bottom - top) / 2;
  const kx = KAPPA * rx;
  const ky = KAPPA * ry;
  return [
    moveTo(map, { x: cx + rx, y: cy }),
    curveTo(map, { x: cx + rx, y: cy + ky }, { x: cx + kx, y: cy + ry }, { x: cx, y: cy + ry }),
    curveTo(map, { x: cx - kx, y: cy + ry }, { x: cx - rx, y: cy + ky }, { x: cx - rx, y: cy }),
    curveTo(map, { x: cx - rx, y: cy - ky }, { x: cx - kx, y: cy - ry }, { x: cx, y: cy - ry }),
    curveTo(map, { x: cx + kx, y: cy - ry }, { x: cx + rx, y: cy - ky }, { x: cx + rx, y: cy }),
    { op: "h" },
  ];
}

/** Arrow head / marker at `tip`, pointing along `direction` (unit vector). */
function endingCommands(
  map: Mapper,
  ending: ShapeLineEnding,
  tip: DisplayPoint,
  direction: DisplayPoint,
  strokeWidth: number,
): PdfNativeContentAppendPathCommand[] {
  const size = Math.max(6, strokeWidth * 4);
  const normal = { x: -direction.y, y: direction.x };
  const base = { x: tip.x - direction.x * size, y: tip.y - direction.y * size };
  const wingA = { x: base.x + normal.x * size * 0.5, y: base.y + normal.y * size * 0.5 };
  const wingB = { x: base.x - normal.x * size * 0.5, y: base.y - normal.y * size * 0.5 };
  switch (ending) {
    case "openArrow":
      return [moveTo(map, wingA), lineTo(map, tip), lineTo(map, wingB)];
    case "closedArrow":
      return [moveTo(map, tip), lineTo(map, wingA), lineTo(map, wingB), { op: "h" }];
    case "square": {
      const half = size / 2;
      return rectangleCommands(map, tip.x - half, tip.y - half, tip.x + half, tip.y + half, 0);
    }
    case "circle": {
      const half = size / 2;
      return ellipseCommands(map, tip.x - half, tip.y - half, tip.x + half, tip.y + half);
    }
    default:
      return [];
  }
}

function unit(from: DisplayPoint, to: DisplayPoint): DisplayPoint | null {
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const length = Math.hypot(dx, dy);
  if (!Number.isFinite(length) || length < 0.001) return null;
  return { x: dx / length, y: dy / length };
}

function isClosedEnding(ending: ShapeLineEnding | undefined): boolean {
  return ending === "closedArrow" || ending === "square" || ending === "circle";
}

/**
 * Build the native path content for a drawn shape, or null when the shape
 * cannot be expressed as one painted path (cloud outlines, visual covers,
 * degenerate geometry).
 */
function composedFillOpacity(
  fillOpacity: number | null,
  opacity: number | null,
): number | null {
  const composed = (fillOpacity ?? 1) * (opacity ?? 1);
  return composed < 1 ? Math.max(0, composed) : null;
}

export function pdfNativeContentAppendPathFromShape(
  shape: ShapeAnnotation,
  geometry: PdfNativeContentAppendPageGeometry,
): PdfNativeContentAppendPathContent | null {
  if (shape.purpose || shape.kind === "cloud") return null;
  const map = mapper(geometry);
  const fillColor: PdfNativeContentAppendColor | null =
    shape.fillColor && (shape.kind === "rect" || shape.kind === "ellipse" || shape.kind === "polygon")
      ? pdfNativeContentAppendColorFromHex(shape.fillColor)
      : null;
  // A filled shape with an explicit width of 0 has no outline: it is
  
  const strokeless = shape.width === 0 && fillColor !== null;
  const strokeColor = pdfNativeContentAppendColorFromHex(shape.color);
  if (!strokeColor && !strokeless) return null;
  const widthPercent = Number.isFinite(shape.width)
    ? Math.min(PDF_SHAPE_MAX_WIDTH, Math.max(PDF_SHAPE_MIN_WIDTH, shape.width))
    : DEFAULT_SHAPE_WIDTH;
  const strokeWidth = (widthPercent / 100) * map.width;
  const fillOpacity =
    fillColor && typeof shape.fillOpacity === "number" && Number.isFinite(shape.fillOpacity)
      ? Math.min(1, Math.max(0, shape.fillOpacity))
      : null;
  if (fillColor && fillOpacity === 0) {
    // Zero fill opacity means "no fill" in the annotation model.
    return pdfNativeContentAppendPathFromShape({ ...shape, fillColor: null }, geometry);
  }
  const opacity =
    typeof shape.opacity === "number" && Number.isFinite(shape.opacity) && shape.opacity < 1
      ? Math.max(0, shape.opacity)
      : null;
  const start = percentPoint(map, shape.x1, shape.y1);
  const end = percentPoint(map, shape.x2, shape.y2);
  const left = Math.min(start.x, end.x);
  const right = Math.max(start.x, end.x);
  const top = Math.min(start.y, end.y);
  const bottom = Math.max(start.y, end.y);
  let commands: PdfNativeContentAppendPathCommand[] = [];
  let fill = fillColor;
  switch (shape.kind) {
    case "rect": {
      if (right - left < 0.5 || bottom - top < 0.5) return null;
      const shorter = Math.min(right - left, bottom - top);
      const radius =
        typeof shape.cornerRadius === "number" && shape.cornerRadius > 0
          ? (Math.min(50, shape.cornerRadius) / 100) * shorter
          : 0;
      commands = rectangleCommands(map, left, top, right, bottom, radius);
      break;
    }
    case "ellipse":
      if (right - left < 0.5 || bottom - top < 0.5) return null;
      commands = ellipseCommands(map, left, top, right, bottom);
      break;
    case "polygon":
    case "polyline": {
      const points = (shape.points ?? []).map((point) => percentPoint(map, point.x, point.y));
      if (points.length < 2) return null;
      commands = [moveTo(map, points[0]), ...points.slice(1).map((point) => lineTo(map, point))];
      if (shape.kind === "polygon") commands.push({ op: "h" });
      if (shape.kind === "polyline") {
        const startDirection = unit(points[1], points[0]);
        const endDirection = unit(points[points.length - 2], points[points.length - 1]);
        if (startDirection && shape.lineStart && shape.lineStart !== "none") {
          commands.push(...endingCommands(map, shape.lineStart, points[0], startDirection, strokeWidth));
        }
        if (endDirection && shape.lineEnd && shape.lineEnd !== "none") {
          commands.push(...endingCommands(map, shape.lineEnd, points[points.length - 1], endDirection, strokeWidth));
        }
      }
      break;
    }
    case "line":
    case "arrow": {
      const direction = unit(start, end);
      if (!direction) return null;
      commands = [moveTo(map, start), lineTo(map, end)];
      const lineStart = shape.lineStart ?? "none";
      const lineEnd = shape.lineEnd ?? (shape.kind === "arrow" ? "openArrow" : "none");
      if (lineStart !== "none") {
        commands.push(...endingCommands(map, lineStart, start, { x: -direction.x, y: -direction.y }, strokeWidth));
      }
      if (lineEnd !== "none") {
        commands.push(...endingCommands(map, lineEnd, end, direction, strokeWidth));
      }
      break;
    }
    default:
      return null;
  }
  // Closed markers on a line are painted with the stroke colour; a line has
  // no interior of its own, so filling the whole path only fills the markers.
  const closedMarkers =
    (shape.kind === "line" || shape.kind === "arrow" || shape.kind === "polyline") &&
    (isClosedEnding(shape.lineStart) ||
      isClosedEnding(shape.lineEnd ?? (shape.kind === "arrow" ? "openArrow" : "none")));
  if (closedMarkers && !fill) fill = strokeColor;
  const lineStyle = shape.lineStyle ?? "solid";
  const dash =
    lineStyle === "dashed"
      ? [strokeWidth * 3, strokeWidth * 2]
      : lineStyle === "dotted"
        ? [Math.max(0.01, strokeWidth * 0.01), strokeWidth * 2]
        : null;
  return Object.freeze({
    kind: "path" as const,
    stroke:
      strokeless || !strokeColor
        ? null
        : Object.freeze({
            color: strokeColor,
            width: strokeWidth,
            dash,
            cap: (lineStyle === "dotted" ? 1 : 0) as 0 | 1,
            join: 1 as const,
          }),
    fill: fill ? Object.freeze({ color: fill }) : null,
    opacity,
    // The shape's own opacity applies to the fill as well as the stroke; a
    // fill-specific opacity multiplies it. Handing the fill opacity through
    // on its own let the default fill opacity of 1 cancel a 50% shape's
    
    fillOpacity: fill ? composedFillOpacity(fillOpacity, opacity) : null,
    commands,
  });
}
