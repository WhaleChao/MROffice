/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  anchorForObjectBounds,
  objectBoundsPixels,
  type DrawingGrid,
  type SpreadsheetObjectAnchor,
  type SpreadsheetObjectBounds,
} from "./object-buttons";
import { DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX, DEFAULT_SPREADSHEET_HEADER_WIDTH_PX } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { clampChartSize, type SpreadsheetChart, type SpreadsheetChartDrawingAnchor } from "@mrnote/core/web/lib/spreadsheet/chart.ts";

type ChartFrame = Pick<SpreadsheetChart, "x" | "y" | "width" | "height" | "drawingAnchor">;

/** The frame a chart with no stored position is drawn at (the window's stack). */
export type ChartDefaultFrame = (chart: ChartFrame) => { x: number; y: number };

function frameOf(chart: ChartFrame, fallback?: ChartDefaultFrame): SpreadsheetObjectBounds | null {
  const x = chart.x ?? fallback?.(chart).x;
  const y = chart.y ?? fallback?.(chart).y;
  if (x === undefined || y === undefined) return null;
  return { left: x, top: y, width: chart.width, height: chart.height };
}

/** The chart's frame (drawing-layer px) on `grid`, from its cell anchor. */
export function chartFrameFromAnchor(
  anchor: SpreadsheetChartDrawingAnchor,
  grid: Pick<DrawingGrid, "colWidth" | "rowHeight">,
): SpreadsheetObjectBounds {
  return objectBoundsPixels({ id: "", anchor: anchor as SpreadsheetObjectAnchor }, grid.colWidth, grid.rowHeight);
}

/**
 * The cell anchor for the chart at `frame` (drawing-layer px) on `grid`,
 * keeping the placement it has (a new one: Move and size with cells).
 */
export function chartAnchorForFrame(
  current: SpreadsheetChartDrawingAnchor | undefined,
  frame: SpreadsheetObjectBounds,
  grid: DrawingGrid,
): SpreadsheetChartDrawingAnchor {
  const base: SpreadsheetObjectAnchor = current
    ? (current as SpreadsheetObjectAnchor)
    : { kind: "twoCell", from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
  // A drawing starts at cell A1's corner at the earliest, as in Excel: a
  // frame reaching over the headings is moved onto the sheet whole (its
  // size kept), never cut down to the part that is on it.
  const placed = base.kind === "absolute"
    ? frame
    : {
        ...frame,
        left: Math.max(frame.left, DEFAULT_SPREADSHEET_HEADER_WIDTH_PX),
        top: Math.max(frame.top, DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX),
      };
  const next = anchorForObjectBounds(base, placed, grid.rowCount, grid.columnCount, grid.colWidth, grid.rowHeight);
  return next as unknown as SpreadsheetChartDrawingAnchor;
}

/** The chart as drawn on `grid`: its frame read off its cell anchor. */
export function chartDrawnOnGrid<C extends ChartFrame>(chart: C, grid: Pick<DrawingGrid, "colWidth" | "rowHeight">): C {
  if (!chart.drawingAnchor) return chart;
  const frame = chartFrameFromAnchor(chart.drawingAnchor, grid);
  if (
    chart.x === frame.left &&
    chart.y === frame.top &&
    chart.width === frame.width &&
    chart.height === frame.height
  ) {
    return chart;
  }
  return { ...chart, x: frame.left, y: frame.top, width: frame.width, height: frame.height };
}

/**
 * `next`'s charts with their cell anchors reconciled against `previous`'s, on
 * the grid the change was made on: a chart whose frame the change moved or
 * sized is anchored where it now stands (its placement kept); a chart with no
 * anchor (new, or saved before charts had anchors) is anchored on the frame
 * it is drawn at, Move and size with cells. The same array when nothing
 * changed.
 */
export function chartsWithAnchors<C extends ChartFrame & { id: string }>(
  next: readonly C[] | undefined,
  previous: readonly C[] | undefined,
  grid: DrawingGrid,
  fallback?: ChartDefaultFrame,
): C[] | undefined {
  if (!next?.length) return next as C[] | undefined;
  const before = new Map((previous ?? []).map((chart) => [chart.id, chart]));
  let changed = false;
  const out = next.map((chart) => {
    const prior = before.get(chart.id);
    const moved =
      prior !== undefined &&
      prior.drawingAnchor !== undefined &&
      chart.drawingAnchor === prior.drawingAnchor &&
      (chart.x !== prior.x || chart.y !== prior.y || chart.width !== prior.width || chart.height !== prior.height);
    if (chart.drawingAnchor && !moved) return chart;
    const frame = frameOf(chart, fallback);
    if (!frame) return chart;
    changed = true;
    // The stored frame stays inside the chart model's size bounds (a chart
    // drawn smaller, its rows partly hidden, is still that chart).
    const size = clampChartSize(chart.width, chart.height);
    return { ...chart, ...size, drawingAnchor: chartAnchorForFrame(chart.drawingAnchor, frame, grid) };
  });
  return changed ? out : (next as C[]);
}

/**
 * Whether a commit landed the charts a writer asked for: the very array, or
 * the same charts given their cell anchors by the commit
 * ({@link chartsWithAnchors}) and nothing else.
 */
export function chartsLanded(
  landed: readonly object[] | undefined,
  asked: readonly object[] | undefined,
): boolean {
  if (landed === asked) return true;
  if (!landed || !asked || landed.length !== asked.length) return false;
  const anchorKeys = new Set(["drawingAnchor", "width", "height"]);
  return landed.every((chart, index) => {
    const wanted = asked[index] as Record<string, unknown>;
    const got = chart as Record<string, unknown>;
    if (got === wanted) return true;
    return [...new Set([...Object.keys(got), ...Object.keys(wanted)])].every(
      (key) => anchorKeys.has(key) || got[key] === wanted[key],
    );
  });
}
