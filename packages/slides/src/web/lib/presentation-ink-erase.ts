/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Draw tab's Small and Medium Erasers on a slide (PowerPoint: Eraser ▸
 * Small Eraser / Medium Eraser): they rub out the ink under the pointer, and
 * what is left of each stroke stays as one stroke per piece, in the stroke's
 * own pen, each cut flat where the eraser's edge was. The geometry is the
 * shared `documentInkEraserCut` the page and the worksheet use; this only
 * knows how a slide stores a stroke (`shape: "ink"`, `points` in its box).
 *
 * Coordinates are slide EMU, the space `persistedInkPaths` places a stroke
 * in (turned and flipped as it is drawn).
 */

import { documentInkEraserCut, documentInkEraserRadiusPx, documentInkPolylineBetween, type DocumentInkEraser } from "@mrnote/core/web/lib/document-ink.ts";
import type { PersistedInkPath } from "@mrnote/core/web/lib/presentation-annotation.ts";
import { inkStrokeToShape } from "@mrnote/core/web/lib/presentation-freeform.ts";

import { documentInkPressureOf } from "@mrnote/core/web/lib/document-ink.ts";
interface Point {
  x: number;
  y: number;
}

/** The Stroke Eraser's reach on a slide (screen px): a stroke this near the
 * eraser's path, or crossed by it, goes whole. */
export const PRESENTATION_STROKE_ERASER_REACH_PX = 10;

/** How far an eraser reaches around the pointer on a slide (screen px): what
 * it rubs out, and how wide its trail is drawn. */
export function presentationInkEraserReachPx(eraser: DocumentInkEraser): number {
  return eraser === "stroke" ? PRESENTATION_STROKE_ERASER_REACH_PX : documentInkEraserRadiusPx(eraser);
}

/** What a piece does not take over from the stroke it was cut from: its
 * place (rebuilt, already turned and flipped), its cut ends (its own), and
 * the source package's identity of the whole stroke, which a piece is not:
 * the shape it was, its place in that package's order and the name the
 * package gave it. A piece is a new stroke, which is what lets an imported
 * deck write it back as one (an appended shape) rather than be rewritten. */
const PLACEMENT_KEYS = [
  "points",
  "customPath",
  "source",
  "sourceOrder",
  "zIndex",
  "name",
  "left",
  "top",
  "width",
  "height",
  "rotation",
  "flipH",
  "flipV",
  "inkCutEnds",
  // A piece keeps the stroke's pen (`inkTool`, `inkEffect`): the highlighter
  
  // pressure is its own stretch of the stroke's, taken again below.
  "inkPressure",
] as const;

/**
 * The slide's blocks once a point eraser has passed along `eraser` (slide
 * EMU) with reach `radiusEmu`, and the strokes it changed, or null when it
 * touched none. `strokes` are the ink the gesture may change, by index into
 * `blocks`; each stroke touched is replaced in place by its pieces (new ids,
 * `newId`), so the z-order holds. `emuPerPx` is the slide's scale on screen:
 * the cut is placed to a hundredth of a screen pixel.
 */
export function erasePresentationInkUnder<B extends { id: string }>(
  blocks: readonly B[],
  strokes: ReadonlyArray<{ readonly index: number; readonly path: PersistedInkPath; readonly whole?: boolean }>,
  eraser: readonly Point[],
  radiusEmu: number,
  emuPerPx: number,
  newId: () => string,
): { blocks: B[]; erasedIds: Set<string> } | null {
  if (!(emuPerPx > 0) || eraser.length === 0) return null;
  const px = (point: Point) => ({ x: point.x / emuPerPx, y: point.y / emuPerPx });
  const eraserPx = eraser.map(px);
  const replaced = new Map<number, B[]>();
  const erasedIds = new Set<string>();
  for (const { index, path, whole } of strokes) {
    const block = blocks[index];
    if (!block || block.id !== path.id) continue;
    const pieces = documentInkEraserCut(
      path.points.map(px),
      eraserPx,
      radiusEmu / emuPerPx,
      (path.width ?? 0) / emuPerPx,
      path.cutEnds,
    );
    if (!pieces) continue;
    if (whole) {
      // A drawn shape is one outline, not ink to cut: the eraser that
      // touches it takes all of it, as the Stroke Eraser does.
      replaced.set(index, []);
      erasedIds.add(block.id);
      continue;
    }
    const kept: Record<string, unknown> = { ...block };
    for (const key of PLACEMENT_KEYS) delete kept[key];
    replaced.set(
      index,
      pieces.flatMap(({ from, to, cutEnds }) => {
        const piece = documentInkPolylineBetween(path.points, from, to);
        const shape = inkStrokeToShape(piece);
        if (!shape) return [];
        // The pen's pressure along the piece, one reading per stored point.
        const inkPressure =
          shape.points.split(" ").length === piece.length ? documentInkPressureOf(piece) : undefined;
        return [
          {
            ...kept,
            id: newId(),
            points: shape.points,
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
            ...(cutEnds ? { inkCutEnds: cutEnds } : {}),
            ...(inkPressure ? { inkPressure } : {}),
          } as unknown as B,
        ];
      }),
    );
    erasedIds.add(block.id);
  }
  if (replaced.size === 0) return null;
  return {
    blocks: blocks.flatMap((block, index) => replaced.get(index) ?? [block]),
    erasedIds,
  };
}
