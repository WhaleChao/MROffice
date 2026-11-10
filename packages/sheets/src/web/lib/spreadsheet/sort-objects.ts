/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX,
  DEFAULT_SPREADSHEET_HEADER_WIDTH_PX,
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
  offsetOfLine,
  type SizeMap,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import {
  EMU_PER_CSS_PIXEL,
  objectBoundsPixels,
  type DrawingObjectLike,
} from "./object-buttons";

/** The sheet fields that size its rows and columns (CSS pixels). */
export interface SortObjectsGrid {
  rowHeights?: SizeMap;
  columnWidths?: SizeMap;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
}

const positive = (value: unknown, fallback: number) =>
  typeof value === "number" && Number.isFinite(value) && value >= 0 ? value : fallback;

/**
 * `objects` after the rows of `span` are permuted (`destinationOf(source)`
 * is where source row `source` lands, or undefined when it does not move).
 * The same array when nothing travels.
 */
export function objectsAfterRowSort<O extends DrawingObjectLike>(
  objects: O[] | undefined,
  grid: SortObjectsGrid,
  span: { r1: number; r2: number; c1: number; c2: number },
  destinationOf: (sourceRow: number) => number | undefined,
): O[] | undefined {
  if (!objects?.length) return objects;
  const rowHeight = (row: number) =>
    positive(grid.rowHeights?.[row], positive(grid.defaultRowHeight, DEFAULT_SPREADSHEET_ROW_HEIGHT_PX));
  const colWidth = (column: number) =>
    positive(grid.columnWidths?.[column], positive(grid.defaultColumnWidth, DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX));
  const rowTop = (row: number) => offsetOfLine(row, row, rowHeight);
  const left = offsetOfLine(span.c1, span.c1, colWidth);
  const right = offsetOfLine(span.c2 + 1, span.c2 + 1, colWidth);
  const epsilon = 0.01;
  let changed = false;
  const next = objects.map((object) => {
    // A grouped object travels only as its whole group; ours are kept apart.
    if ((object as { groupId?: unknown }).groupId || !object.anchor) return object;
    const bounds = objectBoundsPixels(object, colWidth, rowHeight);
    const top = bounds.top - DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX;
    const x = bounds.left - DEFAULT_SPREADSHEET_HEADER_WIDTH_PX;
    if (x < left - epsilon || x + bounds.width > right + epsilon) return object;
    let row = span.r1;
    while (row <= span.r2 && rowTop(row + 1) <= top + epsilon) row += 1;
    if (row > span.r2 || top < rowTop(row) - epsilon) return object;
    if (top + bounds.height > rowTop(row + 1) + epsilon) return object;
    const destination = destinationOf(row);
    if (destination === undefined || destination === row) return object;
    changed = true;
    const anchor = object.anchor;
    if (anchor.kind === "absolute") {
      const yEmu = Math.max(0, Number(anchor.yEmu ?? 0) + Math.round((rowTop(destination) - rowTop(row)) * EMU_PER_CSS_PIXEL));
      return { ...object, anchor: { ...anchor, yEmu } };
    }
    const delta = destination - row;
    return {
      ...object,
      anchor: {
        ...anchor,
        ...(anchor.from ? { from: { ...anchor.from, row: Math.max(0, Number(anchor.from.row ?? 0) + delta) } } : {}),
        ...(anchor.to ? { to: { ...anchor.to, row: Math.max(0, Number(anchor.to.row ?? 0) + delta) } } : {}),
      },
    };
  });
  return changed ? next : objects;
}

/** The sheet's own charts after the same sort: by their cell anchors. */
export function chartsAfterRowSort<C extends { id: string; drawingAnchor?: unknown }>(
  charts: C[] | undefined,
  grid: SortObjectsGrid,
  span: { r1: number; r2: number; c1: number; c2: number },
  destinationOf: (sourceRow: number) => number | undefined,
): C[] | undefined {
  if (!charts?.some((chart) => chart.drawingAnchor)) return charts;
  const anchored = charts.map((chart) => ({
    id: chart.id,
    anchor: chart.drawingAnchor as DrawingObjectLike["anchor"],
  }));
  const sorted = objectsAfterRowSort(anchored, grid, span, destinationOf);
  if (sorted === anchored || !sorted) return charts;
  return charts.map((chart, index) =>
    sorted[index] === anchored[index] ? chart : { ...chart, drawingAnchor: sorted[index].anchor },
  );
}
