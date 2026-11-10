/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The pixels a source vector carries while its frame is dragged.
 *
 * A native vector write lands seconds after the pointer is released, so
 * during the gesture only the frame used to move. The shape's own pixels are
 * on the page canvas already; this cuts them out at the frame's start and,
 * when the renderer proved a simple closed outline for the object, keeps
 * only what lies inside that outline so a circle drawn over other shapes
 * does not drag a square of its neighbours along.
 */

import { cropCanvasRegion } from "./canvas-crop";
import type { FractionRect } from "./embedded-image-edits";
import type { PdfVectorContentItem } from "./vector-content-inventory";

/** The outline runs down the middle of a stroke; grow it so the edge comes along whole. */
const OUTLINE_GROW_PX = 3;

export function captureNativeVectorDragGhost(
  source: HTMLCanvasElement | null | undefined,
  bounds: FractionRect,
  item: Pick<PdfVectorContentItem, "simpleSubpath" | "paintMode"> | null | undefined,
  options: { createCanvas?: () => HTMLCanvasElement } = {},
): HTMLCanvasElement | null {
  const crop = cropCanvasRegion(source, bounds, {
    createCanvas: options.createCanvas,
  });
  if (!crop) return null;
  const { canvas } = crop;
  const outline = item?.simpleSubpath;
  // A fill closes its path on its own (`m c c c c f*` carries no `h`); only a
  // stroke-only open path has no inside to keep.
  const encloses = outline?.closed || item?.paintMode !== "stroke";
  if (
    !outline ||
    !encloses ||
    outline.commands.length < 3 ||
    bounds.width <= 0 ||
    bounds.height <= 0 ||
    canvas.width <= 0 ||
    canvas.height <= 0
  ) {
    return canvas;
  }
  const ctx = canvas.getContext("2d");
  if (!ctx) return canvas;
  try {
    ctx.save();
    // Keep the crop only where the (slightly grown) outline is filled.
    ctx.globalCompositeOperation = "destination-in";
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.scale(
      (canvas.width + 2 * OUTLINE_GROW_PX) / canvas.width,
      (canvas.height + 2 * OUTLINE_GROW_PX) / canvas.height,
    );
    ctx.translate(-canvas.width / 2, -canvas.height / 2);
    traceNativeVectorOutline(ctx, outline, bounds, canvas.width, canvas.height);
    ctx.fillStyle = "#000";
    ctx.fill();
    ctx.restore();
  } catch {
    // A canvas that cannot be masked keeps its plain crop.
  }
  return canvas;
}

/** Lay the renderer's proved outline, in page fractions, onto a crop of `bounds`. */
export function traceNativeVectorOutline(
  ctx: CanvasRenderingContext2D,
  outline: NonNullable<PdfVectorContentItem["simpleSubpath"]>,
  bounds: FractionRect,
  width: number,
  height: number,
): void {
  const scaleX = width / bounds.width;
  const scaleY = height / bounds.height;
  const px = (point: Readonly<{ x: number; y: number }>) => ({
    x: (point.x - bounds.x) * scaleX,
    y: (point.y - bounds.y) * scaleY,
  });
  ctx.beginPath();
  for (const command of outline.commands) {
    if (command.kind === "moveTo") {
      const p = px(command.point);
      ctx.moveTo(p.x, p.y);
    } else if (command.kind === "lineTo") {
      const p = px(command.point);
      ctx.lineTo(p.x, p.y);
    } else if (command.kind === "curveTo") {
      const c1 = px(command.control1);
      const c2 = px(command.control2);
      const end = px(command.end);
      ctx.bezierCurveTo(c1.x, c1.y, c2.x, c2.y, end.x, end.y);
    } else if (command.kind === "quadraticCurveTo") {
      const c = px(command.control);
      const end = px(command.end);
      ctx.quadraticCurveTo(c.x, c.y, end.x, end.y);
    } else {
      ctx.closePath();
    }
  }
  ctx.closePath();
}
