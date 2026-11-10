/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Draw tab's Small and Medium Erasers on a worksheet: they rub out the
 * ink under the pointer, and what is left of a stroke stays as one stroke per
 * unbroken run, in the stroke's own pen (Office's point erasers; the Stroke
 * Eraser removes whole strokes and is the editor's own hit test).
 *
 * Coordinates are the worksheet's content pixels, the space a Draw gesture
 * and an absolute ink anchor share.
 */

import {
  documentInkEraserCut,
  documentInkPolylineBetween,
} from "@mrnote/core/web/lib/document-ink.ts";
import { drawnShapeOutlineFractions, type DrawnShapeKind } from "@mrnote/core/web/lib/office-drawn-shape-outline.ts";
import {
  createSpreadsheetInkShapeObject,
  isSpreadsheetInkShapeStyle,
  SPREADSHEET_CUSTOM_SHAPE_KIND,
  parseSpreadsheetInkPoints,
  parseSpreadsheetInkPressure,
  type SpreadsheetInkPoint,
  type SpreadsheetShapeObject,
  type SpreadsheetShapeStyle,
} from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import { officeFramePlace } from "@mrnote/core/web/lib/office-object-frame.ts";

const EMU_PER_CSS_PIXEL = 9525;

/** The worksheet objects an eraser looks at: anything, an ink stroke or not. */
export interface SpreadsheetInkEraseCandidate {
  id: string;
  type?: string;
  shape?: unknown;
  anchor?: unknown;
}

/**
 * The width a stroke is painted with, in content pixels: its pen's width
 * (the canvas paints it in the box's own pixels, so a resized box keeps it),
 * or the canvas's 2px fallback.
 */
export function spreadsheetInkStrokeWidthPx(
  shape: Pick<SpreadsheetShapeStyle, "strokeWidth"> | undefined,
): number {
  const width = Number(shape?.strokeWidth);
  return Number.isFinite(width) && width > 0 ? width : 2;
}

/** A stroke's samples in content pixels, or null when it is not placed ink. */
/** An ink object's own box in sheet pixels (before its rotation), or null. */
export function spreadsheetInkObjectBoxPx(
  object: SpreadsheetInkEraseCandidate,
): { left: number; top: number; width: number; height: number } | null {
  const anchor = object.anchor as
    | { kind?: string; xEmu?: number; yEmu?: number; widthEmu?: number; heightEmu?: number }
    | undefined;
  if (anchor?.kind !== "absolute") return null;
  const values = [anchor.xEmu, anchor.yEmu, anchor.widthEmu, anchor.heightEmu].map(Number);
  if (!values.every(Number.isFinite)) return null;
  const [x, y, w, h] = values;
  return {
    left: Math.max(0, x / EMU_PER_CSS_PIXEL),
    top: Math.max(0, y / EMU_PER_CSS_PIXEL),
    width: Math.max(1, w / EMU_PER_CSS_PIXEL),
    height: Math.max(1, h / EMU_PER_CSS_PIXEL),
  };
}

export function spreadsheetInkStrokePoints(
  object: SpreadsheetInkEraseCandidate,
): SpreadsheetInkPoint[] | null {
  if (object.type !== "shape") return null;
  const shape = object.shape as SpreadsheetShapeStyle | undefined;
  if (!isSpreadsheetInkShapeStyle(shape)) return null;
  const anchor = object.anchor as
    | {
        kind?: string;
        xEmu?: number;
        yEmu?: number;
        widthEmu?: number;
        heightEmu?: number;
        rotation?: number;
        flipH?: boolean;
        flipV?: boolean;
      }
    | undefined;
  if (anchor?.kind !== "absolute") return null;
  const values = [anchor.xEmu, anchor.yEmu, anchor.widthEmu, anchor.heightEmu].map(Number);
  if (!values.every(Number.isFinite)) return null;
  const [x, y, w, h] = values;
  const left = Math.max(0, x / EMU_PER_CSS_PIXEL);
  const top = Math.max(0, y / EMU_PER_CSS_PIXEL);
  const width = Math.max(1, w / EMU_PER_CSS_PIXEL);
  const height = Math.max(1, h / EMU_PER_CSS_PIXEL);
  // A stroke resized, rotated or flipped with its selection frame draws its
  // percentage points inside that box, turned about the box's centre; the
  // eraser and the lasso read the same place the stroke shows.
  const frame = { left, top, width, height, rotation: anchor.rotation, flipH: anchor.flipH, flipV: anchor.flipV };
  const stored = parseSpreadsheetInkPoints(shape.inkPoints);
  // A pressure pen's reading rides with each point, so a piece the eraser
  // leaves keeps the width the pen drew it with.
  const pressures = parseSpreadsheetInkPressure(shape.inkPressure, stored.length);
  const points = stored.map((point, index) => {
    const placed: SpreadsheetInkPoint = { ...officeFramePlace(frame, point.x / 100, point.y / 100) };
    if (pressures) placed.p = pressures[index];
    return placed;
  });
  return points.length >= 2 ? points : null;
}

/** The sheet kinds Ink to Shape draws, by the fit's own kind (a polygon is custom geometry). */
const DRAWN_SHAPE_KINDS: Readonly<Record<string, DrawnShapeKind>> = {
  rectangle: "rectangle",
  ellipse: "ellipse",
  triangle: "triangle",
  "right-triangle": "right-triangle",
  diamond: "diamond",
  line: "line",
  arrow: "arrow",
  [SPREADSHEET_CUSTOM_SHAPE_KIND]: "polygon",
};








export function spreadsheetDrawnShapeOutlinePoints(
  object: SpreadsheetInkEraseCandidate,
): SpreadsheetInkPoint[] | null {
  if (object.type !== "shape") return null;
  const shape = object.shape as SpreadsheetShapeStyle | undefined;
  if (!shape || isSpreadsheetInkShapeStyle(shape)) return null;
  const kind = DRAWN_SHAPE_KINDS[String(shape.kind)];
  if (!kind || shape.fill !== "none" || shape.stroke === "none") return null;
  if ((typeof shape.text === "string" && shape.text.trim() !== "") || shape.customPath) return null;
  const anchor = object.anchor as { rotation?: number; flipH?: boolean; flipV?: boolean } | undefined;
  const box = spreadsheetInkObjectBoxPx(object);
  if (!box || !anchor) return null;
  const polygon =
    kind === "polygon"
      ? parseSpreadsheetInkPoints(shape.customPoints).map((point) => ({ x: point.x / 100, y: point.y / 100 }))
      : undefined;
  if (kind === "polygon" && (!polygon || polygon.length < 3)) return null;
  const outline = drawnShapeOutlineFractions(kind, {
    polygon,
    flipH: anchor.flipH === true,
    flipV: anchor.flipV === true,
  });
  if (!outline) return null;
  // The outline already holds the mirror.
  const frame = { ...box, rotation: anchor.rotation };
  return outline.map((point) => ({ ...officeFramePlace(frame, point.x, point.y, { mirrored: false }) }));
}

/**
 * The objects once a point eraser has passed along `eraser` with reach
 * `radiusPx`, or null when it touched no ink. `canErase` says which strokes
 * the gesture may change (the ones unchanged since it began); the rest, and
 * every object that is not ink, keep their places. Each stroke touched is
 * replaced by the pieces the eraser left of it, cut flat where the eraser's
 * edge was (the shared `documentInkEraserCut`).
 */
export function eraseSpreadsheetInkUnder<T extends SpreadsheetInkEraseCandidate>(
  objects: readonly T[],
  eraser: readonly SpreadsheetInkPoint[],
  radiusPx: number,
  canErase: (object: T) => boolean = () => true,
): Array<T | SpreadsheetShapeObject> | null {
  let changed = false;
  const next: Array<T | SpreadsheetShapeObject> = [];
  for (const object of objects) {
    const points = canErase(object) ? spreadsheetInkStrokePoints(object) : null;
    if (!points) {
      next.push(object);
      continue;
    }
    const shape = object.shape as SpreadsheetShapeStyle & {
      inkTool: NonNullable<SpreadsheetShapeStyle["inkTool"]>;
    };
    // Tested against the painted line, as the stroke shows it.
    const pieces = documentInkEraserCut(
      points,
      eraser,
      radiusPx,
      spreadsheetInkStrokeWidthPx(shape),
      shape.inkCutEnds,
    );
    if (!pieces) {
      next.push(object);
      continue;
    }
    changed = true;
    for (const { from, to, cutEnds } of pieces) {
      const piece = createSpreadsheetInkShapeObject({
        points: documentInkPolylineBetween(points, from, to),
        tool: shape.inkTool,
        color: shape.stroke,
        width: shape.strokeWidth,
        effect: shape.inkEffect,
        cutEnds,
      });
      if (piece) next.push(piece);
    }
  }
  return changed ? next : null;
}
