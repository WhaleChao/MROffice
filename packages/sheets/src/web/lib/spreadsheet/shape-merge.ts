/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shape Format > Insert Shapes > Merge Shapes on a worksheet: a sheet shape
 * read as the deck-like shape the server's boolean takes
 * (`lib/office-shape-merge-client`), and the merged results written back as
 * the sheet's own custom-outline shapes. PowerPoint is the Office application
 * that lists Merge Shapes; MRAI offers it on a sheet too so the Shape Format
 * tab is one tab.
 */

import { officeShapeDefinition } from "@mrnote/core/web/lib/office-shape-catalog.ts";
import {
  officeCustomPathMergePoints,
  officeMergedShapeCustomPath,
} from "@mrnote/core/web/lib/office-shape-merge-client.ts";
import type { MergeableShape, MergedShapeResult } from "@mrnote/core/web/lib/presentation-shapes.ts";
import {
  isSpreadsheetLineLikeShapeKind,
  isSpreadsheetPresetShapeKind,
  normalizeShapeStyle,
  officeShapePresetForSpreadsheetKind,
} from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

interface MergeObjectLike {
  id: string;
  type: string;
  shape?: unknown;
  anchor?: { rotation?: unknown; flipH?: unknown; flipV?: unknown } | undefined;
}

/**
 * A sheet shape as Merge Shapes reads it, its frame in CSS pixels: a preset
 * as the deck's kind, a custom outline of straight sides as its polygon.
 * Null for a picture, a text box, a line or connector, ink, or a curved
 * outline (none has an area the boolean can take).
 */
export function spreadsheetMergeableShape(
  object: MergeObjectLike,
  frame: { left: number; top: number; width: number; height: number },
): MergeableShape | null {
  if (object.type !== "shape" || !object.shape) return null;
  const style = normalizeShapeStyle(object.shape as never);
  const kind = style.kind as string;
  if (kind === "text-box" || kind === "ink") return null;
  let shape: string;
  let points: string | null = null;
  if (kind === "custom-geometry") {
    points = style.customPath ? officeCustomPathMergePoints(style.customPath) : null;
    if (!points) return null;
    shape = "rectangle";
  } else {
    if (!isSpreadsheetPresetShapeKind(kind) || isSpreadsheetLineLikeShapeKind(kind)) return null;
    shape = officeShapeDefinition(officeShapePresetForSpreadsheetKind(kind)).previewKind;
  }
  const rotation = Number(object.anchor?.rotation);
  return {
    shape,
    ...(points ? { points } : {}),
    left: frame.left,
    top: frame.top,
    width: frame.width,
    height: frame.height,
    rotation: Number.isFinite(rotation) ? rotation : 0,
    flipH: object.anchor?.flipH === true,
    flipV: object.anchor?.flipV === true,
  };
}

/**
 * The sheet's objects after a merge: the results take the first shape's place
 * in the collection and its paint (PowerPoint keeps the first selected
 * shape's formatting), without its words, its group or the file identity it
 * came with; the second shape goes. `anchorFor` places a result's pixel frame.
 */
export function spreadsheetObjectsAfterMerge<O extends MergeObjectLike>(
  objects: readonly O[],
  firstId: string,
  secondId: string,
  results: readonly MergedShapeResult[],
  newId: (index: number) => string,
  anchorFor: (first: O, frame: { left: number; top: number; width: number; height: number }) => O["anchor"],
): O[] | null {
  const first = objects.find((object) => object.id === firstId);
  if (!first || !objects.some((object) => object.id === secondId)) return null;
  const shape = normalizeShapeStyle(first.shape as never) as unknown as Record<string, unknown>;
  const paint: Record<string, unknown> = { ...shape };
  for (const key of Object.keys(paint)) {
    if (key.startsWith("text") || key.startsWith("ink") || key === "equation" || key === "adjustments" ||
        key === "customPath" || key === "customPoints" || key.startsWith("headEnd") || key.startsWith("tailEnd")) {
      delete paint[key];
    }
  }
  const merged = results.flatMap((result, index) => {
    const customPath = officeMergedShapeCustomPath(result);
    if (!customPath) return [];
    const {
      source: _source,
      groupId: _groupId,
      name: _name,
      altText: _altText,
      ...rest
    } = first as O & { source?: unknown; name?: unknown; altText?: unknown; groupId?: unknown };
    const plain: Record<string, unknown> = { ...rest };
    for (const key of Object.keys(plain)) {
      if (key.startsWith("smartArt") || key.startsWith("native")) delete plain[key];
    }
    return [
      {
        ...plain,
        id: newId(index),
        shape: { ...paint, kind: "custom-geometry", customPath },
        anchor: anchorFor(first, { left: result.left, top: result.top, width: result.width, height: result.height }),
      } as unknown as O,
    ];
  });
  if (merged.length === 0) return null;
  return objects.flatMap((object) =>
    object.id === firstId ? merged : object.id === secondId ? [] : [object],
  );
}
