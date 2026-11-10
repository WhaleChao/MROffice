/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shape Format > Insert Shapes > Merge Shapes for the editors whose shapes are
 * not slide blocks (Docs, Spreadsheet): the same server boolean the deck asks
 * (engine `tools`, `shapes.merge`, `lib/presentation-shape-merge.ts`) and the
 * same two-shape rule, with each result turned into the custom outline every
 * Office surface stores (`lib/presentation-freeform`).
 *
 * PowerPoint is the only Office application that lists Merge Shapes
 * (Microsoft Support, "Draw a picture by combining and merging shapes" applies
 * to PowerPoint only); MRAI offers it in all three editors so the Shape Format
 * tab is one tab. The host maps its shape into a deck-like `MergeableShape`
 * (a preset kind, or a polygon, and its drawn frame in one unit) and writes
 * the results back in that same unit.
 */

import { callEngine } from "./engine-host/client";
import type { ToolsMergeShapesOperation } from "./engine-host/tools-types";
import {
  normalizePresentationCustomPath,
  PRESENTATION_CUSTOM_PATH_GRID,
  type PresentationCustomPath,
} from "./presentation-freeform";
import { parsePolygonString } from "./presentation-polygon-points";
import {
  isMergeableShapeBlock,
  mergeableShapeFields,
  type MergeableShape,
  type MergedShapeResult,
} from "./presentation-shapes";

/** PowerPoint's five, in its order: Union, Combine, Fragment, Intersect,
 * Subtract (the deck's `MERGE_SHAPES_OPS`). */
export const OFFICE_MERGE_SHAPES_OPERATIONS = [
  "union",
  "combine",
  "fragment",
  "intersect",
  "subtract",
] as const satisfies readonly ToolsMergeShapesOperation[];
export type OfficeMergeShapesOperation = ToolsMergeShapesOperation;

/**
 * A custom outline as the polygon Merge Shapes reads (0..100 in its own box):
 * only one closed figure of straight sides is a polygon; a curve, an arc or
 * several figures answer null (the deck refuses those too).
 */
export function officeCustomPathMergePoints(path: PresentationCustomPath): string | null {
  const normalized = normalizePresentationCustomPath(path);
  if (!normalized?.closed) return null;
  const scale = 100 / PRESENTATION_CUSTOM_PATH_GRID;
  const points: string[] = [];
  for (const [index, command] of normalized.commands.entries()) {
    if (command.type === "close" && index === normalized.commands.length - 1) continue;
    if (command.type !== "moveTo" && command.type !== "lineTo") return null;
    if ((index === 0) !== (command.type === "moveTo")) return null;
    points.push(`${+(command.x * scale).toFixed(3)},${+(command.y * scale).toFixed(3)}`);
  }
  return points.length >= 3 ? points.join(" ") : null;
}

/** Whether two shapes can be merged: both have a fillable outline and a frame. */
export function officeShapesMergeable(first: MergeableShape | null, second: MergeableShape | null): boolean {
  return Boolean(first && second && isMergeableShapeBlock(first) && isMergeableShapeBlock(second));
}

/** The results of one Merge Shapes on two shapes, in the inputs' unit. */
export async function requestOfficeShapeMerge(
  op: OfficeMergeShapesOperation,
  first: MergeableShape,
  second: MergeableShape,
  signal?: AbortSignal,
): Promise<MergedShapeResult[]> {
  const answer = await callEngine(
    "tools",
    "shapes.merge",
    { op, first: mergeableShapeFields(first), second: mergeableShapeFields(second) },
    { signal, cache: true },
  );
  return answer.results;
}

/**
 * One result as a custom outline in its own box: the multi-figure path the
 * server answered, or its single polygon (0..100 in the result's box) on the
 * shared 0..100000 grid. Null when neither can be stored.
 */
export function officeMergedShapeCustomPath(result: MergedShapeResult): PresentationCustomPath | null {
  if (result.customPath) return normalizePresentationCustomPath(result.customPath);
  const points = typeof result.points === "string" ? parsePolygonString(result.points) : [];
  if (points.length < 3) return null;
  const scale = PRESENTATION_CUSTOM_PATH_GRID / 100;
  const grid = (value: number) =>
    Math.max(0, Math.min(PRESENTATION_CUSTOM_PATH_GRID, Math.round(value * scale)));
  return normalizePresentationCustomPath({
    version: 1,
    closed: true,
    commands: points.map((point, index) => ({
      type: index === 0 ? "moveTo" : "lineTo",
      x: grid(point.x),
      y: grid(point.y),
    })),
  });
}
