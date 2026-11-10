/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  buildRedactionPayload,
  type RedactionAnnotation,
} from "@mrnote/core/web/lib/pdf/redaction.ts";
import { nativeShapeSourceLineGeometry } from "./native-shape-annotations";
import { buildShapePayload, type ShapeAnnotation } from "@mrnote/core/web/lib/pdf/shapes.ts";
import {
  buildSignaturePayload,
  type SignatureAnnotation,
} from "@mrnote/core/web/lib/pdf/signatures.ts";
import {
  buildStampPayload,
  type StampAnnotation,
} from "./stamps-annotation";

const PASTE_OFFSET_PERCENT = 2;
const MAX_CASCADE_STEP = 8;

type ShapePastePayload = NonNullable<ReturnType<typeof buildShapePayload>>;
type SignaturePastePayload = NonNullable<
  ReturnType<typeof buildSignaturePayload>
>;
type StampPastePayload = NonNullable<ReturnType<typeof buildStampPayload>>;
type RedactionPastePayload = NonNullable<
  ReturnType<typeof buildRedactionPayload>
>;

export type PdfPlacedAnnotationClipboardSource =
  | { kind: "shape"; annotation: ShapeAnnotation }
  | { kind: "signature"; annotation: SignatureAnnotation }
  | { kind: "stamp"; annotation: StampAnnotation }
  | { kind: "redaction"; annotation: RedactionAnnotation };

/** A page as it is displayed (rotation applied), in points. */
export interface PdfClipboardPageSize {
  width: number;
  height: number;
}

interface ClipboardItemBase {
  sourcePage: number;
  




  sourcePageSize?: PdfClipboardPageSize;
}

export type PdfPlacedAnnotationClipboardItem =
  | (ClipboardItemBase & {
      kind: "shape";
      annotation: Omit<ShapeAnnotation, "id" | "page">;
    })
  | (ClipboardItemBase & {
      kind: "signature";
      annotation: Omit<SignatureAnnotation, "id" | "page">;
    })
  | (ClipboardItemBase & {
      kind: "stamp";
      annotation: Omit<StampAnnotation, "id" | "page">;
    })
  | (ClipboardItemBase & {
      kind: "redaction";
      annotation: Omit<RedactionAnnotation, "id" | "page">;
    });

function validPageSize(
  size: PdfClipboardPageSize | null | undefined,
): size is PdfClipboardPageSize {
  return Boolean(
    size &&
      Number.isFinite(size.width) &&
      Number.isFinite(size.height) &&
      size.width > 0 &&
      size.height > 0,
  );
}

export type PdfPlacedAnnotationPaste =
  | {
      kind: "shape";
      annotationType: "pdf_shapes";
      payload: ShapePastePayload;
    }
  | {
      kind: "signature";
      annotationType: "pdf_signatures";
      payload: SignaturePastePayload;
    }
  | {
      kind: "stamp";
      annotationType: "pdf_stamps";
      payload: StampPastePayload;
    }
  | {
      kind: "redaction";
      annotationType: "pdf_redactions";
      payload: RedactionPastePayload;
    };

/**
 * Capture a validated, record-identity-free snapshot of a placed annotation.
 * Safety-relevant semantic identity (such as `purpose: "visual_mask"`) remains
 * attached so paste cannot turn a non-secure cover into an unlabeled shape.
 *
 * Native third-party shapes use the same `ShapeAnnotation` geometry contract,
 * so callers intentionally pass them as `kind: "shape"`: a paste creates a new
 * managed shape instead of duplicating the opaque source object reference.
 */
export function copyPlacedAnnotation(
  source: PdfPlacedAnnotationClipboardSource,
  sourcePageSize?: PdfClipboardPageSize | null,
): PdfPlacedAnnotationClipboardItem {
  const item = copyPlacedAnnotationContent(source);
  return validPageSize(sourcePageSize)
    ? {
        ...item,
        sourcePageSize: {
          width: sourcePageSize.width,
          height: sourcePageSize.height,
        },
      }
    : item;
}

function copyPlacedAnnotationContent(
  source: PdfPlacedAnnotationClipboardSource,
): PdfPlacedAnnotationClipboardItem {
  if (source.kind === "shape") {
    const {
      page,
      kind,
      x1,
      y1,
      x2,
      y2,
      points,
      color,
      width,
      purpose,
      fillColor,
      fillOpacity,
      lineStyle,
      lineStart,
      lineEnd,
      opacity,
      cornerRadius,
    } = source.annotation;
    // A native source shape's record box is its /Rect; a Line draws its own
    
    
    const native = source.annotation as ShapeAnnotation & {
      sourceLineInBox?: Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2">;
      sourceFillColor?: string;
      sourceFigure?: {
        kind: "polygon" | "polyline";
        points: { x: number; y: number }[];
      };
      text?: string;
      comment?: string;
    };
    // The drawing's comment travels with it: a copy of a reviewed rectangle
    
    const comment = native.comment ?? native.text;
    // A source Polygon/PolyLine copies as the same figure, vertices and all,
    
    const figure = native.sourceFigure;
    const segment =
      kind === "line" || kind === "arrow"
        ? nativeShapeSourceLineGeometry(native)
        : null;
    const nativeFill =
      fillColor == null && typeof native.sourceFillColor === "string"
        ? native.sourceFillColor
        : null;
    return {
      kind: "shape",
      sourcePage: page,
      annotation: {
        kind: figure ? figure.kind : kind,
        ...(segment ?? { x1, y1, x2, y2 }),
        ...(figure
          ? {
              points: figure.points.map((point) => ({
                x: Math.min(x1, x2) + point.x * Math.abs(x2 - x1),
                y: Math.min(y1, y2) + point.y * Math.abs(y2 - y1),
              })),
            }
          : points
            ? { points: points.map((point) => ({ ...point })) }
            : null),
        color,
        width,
        ...(purpose ? { purpose } : null),
        fillColor: nativeFill ?? fillColor,
        fillOpacity: nativeFill ? 1 : fillOpacity,
        ...(lineStyle !== undefined ? { lineStyle } : null),
        ...(lineStart !== undefined ? { lineStart } : null),
        ...(lineEnd !== undefined ? { lineEnd } : null),
        ...(opacity !== undefined ? { opacity } : null),
        ...(cornerRadius !== undefined ? { cornerRadius } : null),
        ...(typeof comment === "string" && comment.trim() ? { comment } : null),
      },
    };
  }
  if (source.kind === "signature") {
    const {
      page,
      x,
      y,
      width,
      height,
      kind,
      image,
      text,
      font,
      nativeAppearanceImage,
      objectRotation,
    } = source.annotation;
    return {
      kind: "signature",
      sourcePage: page,
      annotation: {
        x,
        y,
        width,
        height,
        kind,
        image,
        text,
        font,
        ...(nativeAppearanceImage ? { nativeAppearanceImage } : {}),
        
        ...(objectRotation ? { objectRotation } : {}),
      },
    };
  }
  if (source.kind === "stamp") {
    const { page, x, y, width, height, kind, preset, image } =
      source.annotation;
    return {
      kind: "stamp",
      sourcePage: page,
      annotation: { x, y, width, height, kind, preset, image },
    };
  }
  const { page, x, y, width, height } = source.annotation;
  return {
    kind: "redaction",
    sourcePage: page,
    annotation: { x, y, width, height },
  };
}

function normalizedCascadeOffset(sequence: number): number {
  if (!Number.isFinite(sequence)) return PASTE_OFFSET_PERCENT;
  const step = Math.min(MAX_CASCADE_STEP, Math.max(1, Math.trunc(sequence)));
  return step * PASTE_OFFSET_PERCENT;
}

/**
 * Prefer a down/right cascade. At a page edge, flip that axis so the copy stays
 * fully visible and distinct instead of silently shrinking or leaving the page.
 */
function offsetAxis(min: number, max: number, offset: number): number {
  if (max + offset <= 100) return offset;
  if (min - offset >= 0) return -offset;
  return 0;
}

function offsetBox(
  box: { x: number; y: number; width: number; height: number },
  offset: number,
): { x: number; y: number; width: number; height: number } {
  const dx = offsetAxis(box.x, box.x + box.width, offset);
  const dy = offsetAxis(box.y, box.y + box.height, offset);
  return {
    x: box.x + dx,
    y: box.y + dy,
    width: box.width,
    height: box.height,
  };
}

function offsetShape(
  shape: Omit<ShapeAnnotation, "id" | "page">,
  offset: number,
): Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2" | "points"> {
  const dx = offsetAxis(
    Math.min(shape.x1, shape.x2),
    Math.max(shape.x1, shape.x2),
    offset,
  );
  const dy = offsetAxis(
    Math.min(shape.y1, shape.y2),
    Math.max(shape.y1, shape.y2),
    offset,
  );
  return {
    x1: shape.x1 + dx,
    y1: shape.y1 + dy,
    x2: shape.x2 + dx,
    y2: shape.y2 + dy,
    ...(shape.points
      ? {
          points: shape.points.map((point) => ({
            x: point.x + dx,
            y: point.y + dy,
          })),
        }
      : null),
  };
}








export function rescaleClipboardItemForPage(
  item: PdfPlacedAnnotationClipboardItem,
  targetPageSize?: PdfClipboardPageSize | null,
): PdfPlacedAnnotationClipboardItem {
  const source = item.sourcePageSize;
  if (!validPageSize(source) || !validPageSize(targetPageSize)) return item;
  const sx = source.width / targetPageSize.width;
  const sy = source.height / targetPageSize.height;
  if (Math.abs(sx - 1) < 1e-9 && Math.abs(sy - 1) < 1e-9) return item;
  const fit = (box: { left: number; top: number; right: number; bottom: number }) => {
    const width = box.right - box.left;
    const height = box.bottom - box.top;
    const shrink = Math.min(
      1,
      width > 0 ? 100 / width : 1,
      height > 0 ? 100 / height : 1,
    );
    const fittedWidth = width * shrink;
    const fittedHeight = height * shrink;
    const left = Math.min(100 - fittedWidth, Math.max(0, box.left));
    const top = Math.min(100 - fittedHeight, Math.max(0, box.top));
    return {
      x: (value: number) => left + (value - box.left) * shrink,
      y: (value: number) => top + (value - box.top) * shrink,
      shrink,
    };
  };
  if (item.kind === "shape") {
    const shape = item.annotation;
    const xs = [shape.x1, shape.x2, ...(shape.points?.map((p) => p.x) ?? [])].map(
      (value) => value * sx,
    );
    const ys = [shape.y1, shape.y2, ...(shape.points?.map((p) => p.y) ?? [])].map(
      (value) => value * sy,
    );
    const place = fit({
      left: Math.min(...xs),
      top: Math.min(...ys),
      right: Math.max(...xs),
      bottom: Math.max(...ys),
    });
    // The stroke is a share of the page's shorter side.
    const strokeScale =
      Math.min(source.width, source.height) /
      Math.min(targetPageSize.width, targetPageSize.height);
    return {
      ...item,
      annotation: {
        ...shape,
        x1: place.x(shape.x1 * sx),
        y1: place.y(shape.y1 * sy),
        x2: place.x(shape.x2 * sx),
        y2: place.y(shape.y2 * sy),
        ...(shape.points
          ? {
              points: shape.points.map((point) => ({
                x: place.x(point.x * sx),
                y: place.y(point.y * sy),
              })),
            }
          : null),
        width: shape.width * strokeScale * place.shrink,
      },
    };
  }
  const box = item.annotation;
  const place = fit({
    left: box.x * sx,
    top: box.y * sy,
    right: (box.x + box.width) * sx,
    bottom: (box.y + box.height) * sy,
  });
  const x = place.x(box.x * sx);
  const y = place.y(box.y * sy);
  const resized = {
    x,
    y,
    width: place.x((box.x + box.width) * sx) - x,
    height: place.y((box.y + box.height) * sy) - y,
  };
  if (item.kind === "signature") {
    return { ...item, annotation: { ...item.annotation, ...resized } };
  }
  if (item.kind === "stamp") {
    return { ...item, annotation: { ...item.annotation, ...resized } };
  }
  return { ...item, annotation: { ...item.annotation, ...resized } };
}

/**
 * Rebuild a paste through the same strict payload builders used for a fresh
 * annotation. This keeps clipboard content bounded, sanitized and byte-aligned
 * with the existing API/native-writeback path.
 */
export function buildPlacedAnnotationPaste(
  sourceItem: PdfPlacedAnnotationClipboardItem,
  targetPage: number,
  sequence = 1,
  targetPageSize?: PdfClipboardPageSize | null,
): PdfPlacedAnnotationPaste | null {
  if (!Number.isInteger(targetPage) || targetPage < 1 || targetPage > 10_000) {
    return null;
  }
  const item = rescaleClipboardItemForPage(sourceItem, targetPageSize);
  const offset = normalizedCascadeOffset(sequence);
  if (item.kind === "shape") {
    const geometry = offsetShape(item.annotation, offset);
    const payload = buildShapePayload({
      page: targetPage,
      kind: item.annotation.kind,
      ...geometry,
      points: geometry.points,
      color: item.annotation.color,
      width: item.annotation.width,
      purpose: item.annotation.purpose,
      fillColor: item.annotation.fillColor,
      fillOpacity: item.annotation.fillOpacity,
      lineStyle: item.annotation.lineStyle,
      lineStart: item.annotation.lineStart,
      lineEnd: item.annotation.lineEnd,
      opacity: item.annotation.opacity,
      cornerRadius: item.annotation.cornerRadius,
    });
    return payload
      ? { kind: "shape", annotationType: "pdf_shapes", payload }
      : null;
  }
  if (item.kind === "signature") {
    const box = offsetBox(item.annotation, offset);
    const payload = buildSignaturePayload({
      page: targetPage,
      ...box,
      kind: item.annotation.kind,
      image: item.annotation.image,
      text: item.annotation.text,
      font: item.annotation.font,
      nativeAppearanceImage: item.annotation.nativeAppearanceImage,
      objectRotation: item.annotation.objectRotation,
    });
    return payload
      ? { kind: "signature", annotationType: "pdf_signatures", payload }
      : null;
  }
  if (item.kind === "stamp") {
    const box = offsetBox(item.annotation, offset);
    const payload = buildStampPayload({
      page: targetPage,
      ...box,
      kind: item.annotation.kind,
      preset: item.annotation.preset,
      image: item.annotation.image,
    });
    return payload
      ? { kind: "stamp", annotationType: "pdf_stamps", payload }
      : null;
  }
  const box = offsetBox(item.annotation, offset);
  const payload = buildRedactionPayload({ page: targetPage, ...box });
  return payload
    ? { kind: "redaction", annotationType: "pdf_redactions", payload }
    : null;
}
