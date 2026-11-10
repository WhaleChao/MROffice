/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { InkShapeFit } from "./office-ink-shape-fit";

/** Ink to Shape's shapes: a fitted preset, a line or arrow, or a polygon. */
export type DrawnShapeKind = InkShapeFit["kind"];

export interface DrawnShapeOutlinePoint {
  readonly x: number;
  readonly y: number;
}

/** Samples around an ellipse's outline: fine enough for an eraser's reach on a page-sized ellipse. */
const ELLIPSE_SAMPLES = 72;

/**
 * The outline of a drawn shape of `kind` (a polygon's corners in `polygon`,
 * fractions of the box) as fractions of its box, mirrored as it is drawn;
 * null for a polygon with no corners.
 */
export function drawnShapeOutlineFractions(
  kind: DrawnShapeKind,
  options: {
    readonly polygon?: readonly DrawnShapeOutlinePoint[];
    readonly closed?: boolean;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
  } = {},
): DrawnShapeOutlinePoint[] | null {
  let outline: DrawnShapeOutlinePoint[];
  switch (kind) {
    case "rectangle":
      outline = [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }, { x: 0, y: 0 }];
      break;
    case "diamond":
      outline = [{ x: 0.5, y: 0 }, { x: 1, y: 0.5 }, { x: 0.5, y: 1 }, { x: 0, y: 0.5 }, { x: 0.5, y: 0 }];
      break;
    case "triangle":
      outline = [{ x: 0.5, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }, { x: 0.5, y: 0 }];
      break;
    case "right-triangle":
      // DrawingML's rtTriangle: the right angle at the bottom left.
      outline = [{ x: 0, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }, { x: 0, y: 0 }];
      break;
    case "ellipse":
      outline = Array.from({ length: ELLIPSE_SAMPLES + 1 }, (_, index) => {
        const angle = (index / ELLIPSE_SAMPLES) * Math.PI * 2;
        return { x: 0.5 + 0.5 * Math.cos(angle), y: 0.5 + 0.5 * Math.sin(angle) };
      });
      break;
    case "polygon": {
      const corners = options.polygon ?? [];
      if (corners.length === 0) return null;
      outline = [...corners];
      if (options.closed !== false) outline.push({ ...corners[0]! });
      break;
    }
    default:
      // A line or an arrow, corner to corner of its box.
      outline = [{ x: 0, y: 0 }, { x: 1, y: 1 }];
  }
  const flipH = options.flipH === true;
  const flipV = options.flipV === true;
  return outline.map((point) => ({ x: flipH ? 1 - point.x : point.x, y: flipV ? 1 - point.y : point.y }));
}
