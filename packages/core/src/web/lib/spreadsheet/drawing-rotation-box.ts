/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { resolveNativeRowHeights, type NativeGrid } from "./native-typography";

const EMU_PER_PX = 9525;
const DEFAULT_COLUMN_PX = 80;
const DEFAULT_ROW_PX = 20;

/** Whether Excel stores the anchor of an object at `rotation` degrees as the
 * quarter-turned rectangle: [45, 135) or [225, 315). */
export function anchorBoxTurned(rotation: unknown): boolean {
  if (typeof rotation !== "number" || !Number.isFinite(rotation)) return false;
  const degrees = ((rotation % 360) + 360) % 360;
  return (degrees >= 45 && degrees < 135) || (degrees >= 225 && degrees < 315);
}

/** One axis of the grid in EMU from cell A1's corner. */
export interface DrawingBoxAxis {
  /** EMU from the grid's start to the start of `line`. */
  start(line: number): number;
  /** The (line, offset EMU) a position falls in; before the grid's start,
   * line 0 at a negative offset (a frame turned about a box near the edge). */
  locate(position: number): [number, number];
  /** EMU of a line with no size of its own. */
  readonly fallback: number;
}

function finitePositive(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value) && value > 0;
}

function sparseAxis(defaultPx: number, overrides: unknown): DrawingBoxAxis {
  const fallback = Math.max(0, defaultPx) * EMU_PER_PX;
  const entries: Array<[number, number]> = [];
  if (overrides && typeof overrides === "object" && !Array.isArray(overrides)) {
    for (const [key, value] of Object.entries(overrides as Record<string, unknown>)) {
      const index = Number(key);
      if (!Number.isInteger(index) || index < 0) continue;
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0) continue;
      entries.push([index, value * EMU_PER_PX]);
    }
  }
  entries.sort((left, right) => left[0] - right[0]);
  const keys = entries.map(([index]) => index);
  const prefix = [0];
  for (const [, size] of entries) prefix.push(prefix[prefix.length - 1] + size - fallback);
  const below = (line: number) => {
    let low = 0;
    let high = keys.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (keys[middle] < line) low = middle + 1;
      else high = middle;
    }
    return low;
  };
  const start = (line: number) => line * fallback + prefix[below(line)];
  const anySize = fallback > 0 || entries.some(([, size]) => size > 0);
  return {
    start,
    fallback,
    locate(position) {
      if (position < 0 || !anySize) return [0, Math.round(position)];
      let low = 0;
      let high = 1;
      while (start(high) <= position && high < 1 << 24) high *= 2;
      while (high - low > 1) {
        const middle = Math.floor((low + high) / 2);
        if (start(middle) <= position) low = middle;
        else high = middle;
      }
      return [low, Math.round(position - start(low))];
    },
  };
}

/** The grid fields the axes are read from. */
export interface DrawingBoxSheet {
  columnWidths?: Record<string | number, number>;
  rowHeights?: Record<string | number, number>;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
  nativeGrid?: NativeGrid;
}

/** The sheet's column and row axes as the page draws them (and the save
 * measures them): a row paired with its point height at that height. */
export function sheetDrawingBoxAxes(sheet: DrawingBoxSheet): { columns: DrawingBoxAxis; rows: DrawingBoxAxis } {
  const exact = resolveNativeRowHeights(
    sheet.nativeGrid?.rowMetrics,
    sheet.rowHeights as Record<string, number> | undefined,
    sheet.defaultRowHeight,
  );
  const rowDefault = exact ? exact.defaultHeight : sheet.defaultRowHeight;
  return {
    columns: sparseAxis(finitePositive(sheet.defaultColumnWidth) ? sheet.defaultColumnWidth : DEFAULT_COLUMN_PX, sheet.columnWidths),
    rows: sparseAxis(finitePositive(rowDefault) ? rowDefault : DEFAULT_ROW_PX, exact ? exact.heights : sheet.rowHeights),
  };
}

/**
 * `axis` after `delta` lines are inserted (positive) or deleted (negative)
 * at `at`. An inserted line is as large as the one before it, as Insert's
 * default (Format Same As Above / Left) sizes it.
 */
export function drawingBoxAxisAfterLineEdit(axis: DrawingBoxAxis, at: number, delta: number): DrawingBoxAxis {
  const first = axis.start(at);
  if (delta > 0) {
    // Nothing before the first line: an insert there is the default size.
    const each = at > 0 ? axis.start(at) - axis.start(at - 1) : axis.fallback;
    const added = each * delta;
    return {
      fallback: axis.fallback,
      start: (line) => (line <= at ? axis.start(line) : line < at + delta ? first + (line - at) * each : axis.start(line - delta) + added),
      locate(position) {
        if (position < first) return axis.locate(position);
        if (position < first + added && each > 0) {
          const line = Math.floor((position - first) / each);
          return [at + line, Math.round(position - first - line * each)];
        }
        const [line, offset] = axis.locate(position - added);
        return [line + delta, offset];
      },
    };
  }
  const count = -delta;
  const removed = axis.start(at + count) - first;
  return {
    fallback: axis.fallback,
    start: (line) => (line <= at ? axis.start(line) : axis.start(line + count) - removed),
    locate(position) {
      if (position < first) return axis.locate(position);
      const [line, offset] = axis.locate(position + removed);
      return [Math.max(at, line - count), offset];
    },
  };
}

type Point = {
  row?: number;
  column?: number;
  col?: number;
  rowOffset?: number;
  columnOffset?: number;
  colOffset?: number;
};

type TurnableAnchor = {
  kind?: string;
  from?: Point;
  to?: Point;
  widthEmu?: number;
  heightEmu?: number;
};

function pointPosition(point: Point | undefined, columns: DrawingBoxAxis, rows: DrawingBoxAxis): [number, number] | null {
  if (!point) return null;
  const row = point.row;
  const column = point.column ?? point.col;
  if (typeof row !== "number" || typeof column !== "number" || row < 0 || column < 0) return null;
  const rowOffset = Number(point.rowOffset ?? 0) || 0;
  const columnOffset = Number(point.columnOffset ?? point.colOffset ?? 0) || 0;
  return [columns.start(Math.trunc(column)) + columnOffset, rows.start(Math.trunc(row)) + rowOffset];
}

function pointAt<P extends Point>(point: P, x: number, y: number, columns: DrawingBoxAxis, rows: DrawingBoxAxis): P {
  const [column, columnOffset] = columns.locate(x);
  const [row, rowOffset] = rows.locate(y);
  const next: Point = { ...point, row, rowOffset };
  if (typeof point.column !== "number" && typeof point.col === "number") {
    next.col = column;
    next.colOffset = columnOffset;
    delete next.columnOffset;
  } else {
    next.column = column;
    next.columnOffset = columnOffset;
    delete next.colOffset;
  }
  return next as P;
}

/**
 * A one-cell or two-cell anchor turned a quarter about its centre on the
 * given axes: Excel's box from the page's frame, or the frame from the box
 * (the turn is its own inverse). Null when it cannot be measured.
 */
export function turnDrawingAnchorBox<A extends TurnableAnchor>(anchor: A, columns: DrawingBoxAxis, rows: DrawingBoxAxis): A | null {
  const start = pointPosition(anchor.from, columns, rows);
  if (!start || !anchor.from) return null;
  if (anchor.kind === "twoCell") {
    const end = pointPosition(anchor.to, columns, rows);
    if (!end || !anchor.to) return null;
    const width = end[0] - start[0];
    const height = end[1] - start[1];
    if (width <= 0 || height <= 0) return null;
    const centerX = start[0] + width / 2;
    const centerY = start[1] + height / 2;
    return {
      ...anchor,
      from: pointAt(anchor.from, centerX - height / 2, centerY - width / 2, columns, rows),
      to: pointAt(anchor.to, centerX + height / 2, centerY + width / 2, columns, rows),
    };
  }
  if (anchor.kind !== "oneCell") return null;
  const width = anchor.widthEmu;
  const height = anchor.heightEmu;
  if (!finitePositive(width) || !finitePositive(height)) return null;
  return {
    ...anchor,
    from: pointAt(anchor.from, start[0] + (width - height) / 2, start[1] + (height - width) / 2, columns, rows),
    widthEmu: Math.round(height),
    heightEmu: Math.round(width),
  };
}
