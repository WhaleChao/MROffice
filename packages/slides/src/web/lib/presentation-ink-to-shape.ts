/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink to Shape on a slide (PPT-A069): the Office shape a pen stroke
 * plainly is, as a slide shape block.
 *
 * Microsoft Support, "Convert ink to shapes in Office": with Ink to Shape on,
 * a drawn rectangle, ellipse, triangle, diamond, other polygon, line or arrow
 * becomes that shape, outlined in the pen's colour and width; anything else
 * stays ink. The fit is the shared one Excel's Draw tab uses
 * (`lib/office-ink-shape-fit.ts`); this only states the result the way a
 * slide stores a shape: a preset kind in the stroke's box, a line or arrow
 * whose direction is its flips, and any other polygon as custom geometry.
 */

import { buildFreeformCustomGeometry } from "@mrnote/core/web/lib/presentation-freeform.ts";
import { inkShapeOpacityForTool } from "@mrnote/core/web/lib/presentation-annotation.ts";
import { inkHoldShapeFit } from "@mrnote/core/web/lib/office-ink-hold-shape.ts";
import { fitInkStrokeToShape, inkSamplesForFitting, type InkFitPoint } from "@mrnote/core/web/lib/office-ink-shape-fit.ts";

/** Slide EMU per CSS pixel (96 dpi): the fit thinks in pixels. */
const EMU_PER_PX = 9_525;

export interface PresentationInkShapePaint {
  readonly color: string;
  /** The pen's stroke in the slide's CSS pixels. */
  readonly strokePx: number;
  readonly tool: "pen" | "highlighter";
}

/** The shape block (without its id) a stroke in slide EMU fits, or null. */
export function presentationShapeFromInkStroke(
  pointsEmu: readonly InkFitPoint[],
  paint: PresentationInkShapePaint,
  /**
   * `stored`: a landed stroke's stored samples (read back at a pen's density
   * first). `held`: drawn and held, a shape the writer asked for -- where the
   * recognizer is unsure, the held stroke's own shape (a line, a polygon).
   */
  options: { readonly stored?: boolean; readonly held?: boolean } = {},
): Record<string, unknown> | null {
  const px = pointsEmu.map((point) => ({ x: point.x / EMU_PER_PX, y: point.y / EMU_PER_PX }));
  const fit =
    fitInkStrokeToShape(options.stored ? inkSamplesForFitting(px) : px) ?? (options.held ? inkHoldShapeFit(px) : null);
  if (!fit) return null;
  const outline = {
    stroke: paint.color,
    strokeWidth: Math.max(1, paint.strokePx),
    ...(paint.tool === "highlighter"
      ? { opacity: inkShapeOpacityForTool("highlighter") }
      : {}),
  };
  const emu = (value: number) => Math.round(value * EMU_PER_PX);
  if (fit.kind === "line" || fit.kind === "arrow") {
    const left = Math.min(fit.start.x, fit.end.x);
    const top = Math.min(fit.start.y, fit.end.y);
    return {
      type: "shape",
      shape: fit.kind,
      ...outline,
      left: emu(left),
      top: emu(top),
      width: Math.max(1, emu(Math.abs(fit.end.x - fit.start.x))),
      height: Math.max(1, emu(Math.abs(fit.end.y - fit.start.y))),
      ...(fit.start.x > fit.end.x ? { flipH: true } : {}),
      ...(fit.start.y > fit.end.y ? { flipV: true } : {}),
    };
  }
  if (fit.kind === "polygon") {
    const result = buildFreeformCustomGeometry(
      fit.points.map((point) => ({ x: emu(point.x), y: emu(point.y) })),
      true,
    );
    if (!result) return null;
    return {
      type: "shape",
      shape: "custom-geometry",
      fill: "none",
      ...outline,
      customPath: result.customPath,
      left: result.left,
      top: result.top,
      width: result.width,
      height: result.height,
    };
  }
  // What is left states a box (a rectangle, ellipse, diamond or triangle).
  if (!("box" in fit)) return null;
  return {
    type: "shape",
    shape: fit.kind,
    fill: "none",
    ...outline,
    left: emu(fit.box.left),
    top: emu(fit.box.top),
    width: Math.max(1, emu(fit.box.width)),
    height: Math.max(1, emu(fit.box.height)),
    ...(fit.flipH ? { flipH: true } : {}),
  };
}

interface InkBlockLike {
  readonly id: string;
  readonly type?: unknown;
  readonly shape?: unknown;
  readonly points?: unknown;
  readonly stroke?: unknown;
  readonly strokeWidth?: unknown;
  readonly opacity?: unknown;
  readonly left?: unknown;
  readonly top?: unknown;
  readonly width?: unknown;
  readonly height?: unknown;
  readonly [key: string]: unknown;
}

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/**
 * Ink to Shape over strokes already on the slide (lasso-selected): each ink
 * stroke in `ids` that plainly is a shape becomes it, keeping its id and its
 * place in the stacking order; the rest stay ink. Null when none converted.
 */
export function convertPresentationInkBlocksToShapes<T extends InkBlockLike>(
  blocks: readonly T[],
  ids: ReadonlySet<string>,
): T[] | null {
  let converted = 0;
  const next = blocks.map((block) => {
    if (!ids.has(block.id) || block.type !== "shape" || block.shape !== "ink") return block;
    if (typeof block.points !== "string") return block;
    const left = finite(block.left);
    const top = finite(block.top);
    const width = finite(block.width);
    const height = finite(block.height);
    if (left === null || top === null || width === null || height === null) return block;
    const pointsEmu = block.points
      .trim()
      .split(/\s+/)
      .map((pair) => pair.split(",").map(Number))
      .filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y))
      .map(([x, y]) => ({ x: left + (x! / 100) * width, y: top + (y! / 100) * height }));
    const opacity = finite(block.opacity);
    const shape = presentationShapeFromInkStroke(
      pointsEmu,
      {
        color: typeof block.stroke === "string" ? block.stroke : "#000000",
        strokePx: finite(block.strokeWidth) ?? 1,
        tool: opacity !== null && opacity < 100 ? "highlighter" : "pen",
      },
      { stored: true },
    );
    if (!shape) return block;
    converted += 1;
    return { id: block.id, ...shape } as unknown as T;
  });
  return converted > 0 ? next : null;
}
