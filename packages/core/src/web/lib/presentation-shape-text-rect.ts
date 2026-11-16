/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A native shape's TEXT RECTANGLE (DrawingML `a:rect`), as fractions of its
 * frame.
 *
 * PowerPoint lays a shape's label out inside the geometry's text rectangle,
 * not inside the frame: an ellipse's inscribed rectangle, the lower part of a
 * triangle, the box above a down-arrow callout's arrow. The import states it
 * (`source.nativeShapeDrawing.textRect`, in the drawing's design-pixel
 * viewBox, beside the paths it belongs to) only when it is not the frame
 * itself, and it scales with the drawing exactly as the paths do.
 *
 * Pure and dependency-free: the canvas label, and any measurement of it, read
 * the same answer.
 */

import { officePresetTextRect } from "./office-shape-adjust-handles";

export type ShapeTextRectFractions = {
  left: number;
  top: number;
  right: number;
  bottom: number;
};

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/**
 * The label's rectangle as `[0..1]`-ish fractions of the frame, or null. Only
 * while the drawing is still the one of this shape kind: a shape re-kinded in
 * the editor no longer has the imported geometry's text rectangle.
 */
export function shapeTextRectFractions(
  source: unknown,
  shapeKind: string,
): ShapeTextRectFractions | null {
  const drawing = record(record(source)?.nativeShapeDrawing);
  if (drawing?.role !== "shape" || drawing.shapeKind !== shapeKind) return null;
  return rectFractions(drawing.textRect, record(drawing?.geometry)?.viewBox);
}

/**
 * The same rectangle for a shape with NO imported drawing (authored here, or
 * re-kinded): the catalog preset's own `a:rect`, evaluated with the shape's
 * adjust values at its size by the ONE evaluator Docs and the worksheet use
 * (`officePresetTextRect`). Null when it is the whole frame, so a rectangle
 * keeps exactly the layout it always had.
 */
export function authoredShapeTextRectFractions(
  drawingPreset: string | null | undefined,
  widthEmu: number,
  heightEmu: number,
  adjustments?: Readonly<Record<string, number>> | null,
): ShapeTextRectFractions | null {
  const rect = officePresetTextRect(drawingPreset, widthEmu / 9525, heightEmu / 9525, adjustments);
  if (!rect) return null;
  const full = Math.abs(rect.left) < 1e-6 && Math.abs(rect.top) < 1e-6 &&
    Math.abs(rect.right - 1) < 1e-6 && Math.abs(rect.bottom - 1) < 1e-6;
  return full ? null : rect;
}









export function textBlockTextRectFractions(
  source: unknown,
): ShapeTextRectFractions | null {
  const stated = record(record(source)?.nativeTextRect);
  return stated ? rectFractions(stated.rect, stated.viewBox) : null;
}

function rectFractions(
  rect: unknown,
  viewBox: unknown,
): ShapeTextRectFractions | null {
  if (
    !Array.isArray(rect) ||
    rect.length !== 4 ||
    !Array.isArray(viewBox) ||
    viewBox.length !== 2
  ) {
    return null;
  }
  const [width, height] = viewBox as unknown[];
  if (
    typeof width !== "number" ||
    typeof height !== "number" ||
    !(width > 0) ||
    !(height > 0) ||
    rect.some((value) => typeof value !== "number" || !Number.isFinite(value))
  ) {
    return null;
  }
  const [left, top, right, bottom] = rect as number[];
  // A rectangle may reach past the frame (DrawingML allows it), but one that
  // is inverted or absurdly far out is not a rectangle this renderer draws in.
  const fractions = {
    left: left / width,
    top: top / height,
    right: right / width,
    bottom: bottom / height,
  };
  if (
    fractions.right <= fractions.left ||
    fractions.bottom <= fractions.top ||
    Object.values(fractions).some((value) => value < -1 || value > 2)
  ) {
    return null;
  }
  return fractions;
}












export function textBodyRotationCss(
  source: unknown,
  box: "frame" | "content",
  block?: unknown,
): { transform?: string; transformOrigin?: string } {
  const degrees = record(source)?.nativeTextBodyRotation;
  const turn =
    typeof degrees === "number" && Number.isFinite(degrees)
      ? ((degrees % 360) + 360) % 360
      : 0;
  // A flipped block's text is turned back first (`flippedTextTransform`), in
  // the same rectangle the body turns in.
  const counter = flippedTextTransform(block);
  if (turn === 0 && !counter) return {};
  const rect = box === "frame" ? textBlockTextRectFractions(source) : null;
  const percent = (fraction: number) => Math.round(fraction * 1_000_000) / 10_000;
  const origin = rect
    ? `${percent((rect.left + rect.right) / 2)}% ${percent((rect.top + rect.bottom) / 2)}%`
    : "50% 50%";
  const transform = [counter, turn ? `rotate(${turn}deg)` : undefined]
    .filter(Boolean)
    .join(" ");
  return { transform, transformOrigin: origin };
}




















export function flippedTextTransform(block: unknown): string | undefined {
  const flips = record(block);
  return (flips?.flipH === true) !== (flips?.flipV === true)
    ? "scaleX(-1)"
    : undefined;
}
