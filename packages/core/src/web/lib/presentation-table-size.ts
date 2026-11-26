/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Optional per-axis SIZE arrays for presentation table blocks (PowerPoint /
 * WPS "column width" / "row height" parity).
 *
 * A table block historically carries only `rows` (a 2-D cell grid). This module
 * adds two OPTIONAL, migration-safe arrays:
 *   - `colWidths[c]`  — the width  (px) of column `c`
 *   - `rowHeights[r]` — the height (px) of row `r`
 * Each entry is a bounded positive integer; a missing / out-of-range entry means
 * "auto" (the column/row sizes to content, exactly as before). When a table
 * carries no sizing it serializes byte-for-byte as it always did.
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit-tested
 * in isolation and shared by the editor, the present overlay, and the schema.
 */

import {
  MAX_TABLE_SPLIT_COLUMNS,
  MAX_TABLE_SPLIT_ROWS,
  logicalCellPositions,
  tableLogicalWidth,
  type MergeableCell,
} from "./presentation-table-merge";
import { PRESENTATION_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";

/** Min / max for a single column width or row height, in CSS px. */
export const MIN_TABLE_AXIS_PX = 16;
export const MAX_TABLE_AXIS_PX = 4_000;
/** DrawingML geometry units used by PowerPoint/WPS table size controls. */
export const TABLE_FRAME_EMU_PER_CM = 360_000;
export const TABLE_FRAME_EMU_PER_CSS_PX = 9_525;
export const MIN_TABLE_FRAME_CM = 0.1;
export const MAX_TABLE_FRAME_EMU = PRESENTATION_GEOMETRY_MAX_EMU;
export const MAX_TABLE_FRAME_CM =
  Math.floor((MAX_TABLE_FRAME_EMU / TABLE_FRAME_EMU_PER_CM) * 10) / 10;
const MIN_TABLE_FRAME_EMU = Math.round(
  MIN_TABLE_FRAME_CM * TABLE_FRAME_EMU_PER_CM,
);








export const MAX_TABLE_AXIS_ENTRIES = Math.max(MAX_TABLE_SPLIT_ROWS, MAX_TABLE_SPLIT_COLUMNS);
/**
 * Fallback equal size (px) used by {@link equalizedAxisSizes} for the
 * "distribute rows / columns" action when NO slot currently carries an explicit
 * size (so an all-auto table still gets a concrete, even size to distribute).
 */
export const DEFAULT_DISTRIBUTE_AXIS_PX = 96;
/** PowerPoint/WPS table-layout column sizing commands. */
export const TABLE_COLUMN_AUTOFIT_MODES = [
  "contents",
  "window",
  "fixed",
] as const;
export type TableColumnAutoFitMode =
  (typeof TABLE_COLUMN_AUTOFIT_MODES)[number];
export type TableFrameAxis = "width" | "height";
export interface ResizeTableFrameOptions {
  /**
   * PowerPoint/WPS Layout → Lock aspect ratio.
   *
   * This is intentionally an EDITOR-SESSION option, not a table-model field:
   * Office does not serialize the checkbox into PPTX table OOXML. When enabled,
   * one resize command scales both native frame axes and both grid arrays.
   */
  lockAspectRatio?: boolean;
}

export interface ResizableTableFrame {
  rows: MergeableCell[][];
  width?: number;
  height?: number;
  colWidths?: unknown;
  rowHeights?: unknown;
}

const MIN_CONTENT_COLUMN_PX = 48;
const MAX_CONTENT_COLUMN_PX = 640;
const DEFAULT_CELL_HORIZONTAL_PADDING_PX = 24;

function boundedAxisCount(value: number): number {
  return Math.min(Math.max(0, Math.floor(value)), MAX_TABLE_AXIS_ENTRIES);
}

function roundToTenths(value: number): number {
  return Math.round(value * 10) / 10;
}

function roundToHundredths(value: number): number {
  return Math.round(value * 100) / 100;
}

/** Read a DrawingML table-frame extent as the centimetres shown by Office. */
export function tableFrameSizeCm(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    return null;
  }
  return roundToHundredths(value / TABLE_FRAME_EMU_PER_CM);
}

/**
 * PowerPoint shows every size in the system measurement unit (zh-CN: cm),
 * "1.35 厘米", not CSS pixels. Row heights / column widths stay stored in px;
 * these convert at the field boundary only.
 */
export function tableAxisPxToCm(px: number): number {
  return roundToHundredths((px * TABLE_FRAME_EMU_PER_CSS_PX) / TABLE_FRAME_EMU_PER_CM);
}

export function tableAxisCmToPx(centimetres: number): number {
  return roundToTenths((centimetres * TABLE_FRAME_EMU_PER_CM) / TABLE_FRAME_EMU_PER_CSS_PX);
}

function tableFrameEmuFromCm(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    return null;
  }
  const centimetres = Math.min(
    MAX_TABLE_FRAME_CM,
    Math.max(MIN_TABLE_FRAME_CM, roundToHundredths(value)),
  );
  return Math.min(
    MAX_TABLE_FRAME_EMU,
    Math.max(1, Math.round(centimetres * TABLE_FRAME_EMU_PER_CM)),
  );
}

function materializedAxisSizes(
  current: unknown,
  count: number,
  fallbackTotalPx: number,
): number[] {
  const slots = Array.from({ length: count }, (_, index) =>
    axisSizeAt(current, index),
  );
  const explicit = slots.filter((size): size is number => size !== null);
  const explicitTotal = explicit.reduce((sum, size) => sum + size, 0);
  const missingCount = count - explicit.length;
  const fallback =
    missingCount > 0
      ? Math.max(
          MIN_TABLE_AXIS_PX,
          (fallbackTotalPx - explicitTotal) / missingCount,
        )
      : 0;
  return slots.map((size) => size ?? fallback);
}

/** `materializedAxisSizes` over the sizes the page holds: a stated size below
 * the 16 px floor counts at its own size; only auto slots take the share. */
function storedMaterializedAxisSizes(current: unknown, count: number, fallbackTotalPx: number): number[] {
  const slots = Array.from({ length: count }, (_, index) =>
    Array.isArray(current) && index < current.length ? storedAxisSize(current[index]) : null,
  );
  const explicitTotal = slots.reduce<number>((sum, size) => sum + (size ?? 0), 0);
  const missingCount = slots.filter((size) => size === null).length;
  const fallback =
    missingCount > 0 ? Math.max(MIN_TABLE_AXIS_PX, (fallbackTotalPx - explicitTotal) / missingCount) : 0;
  return slots.map((size) => size ?? fallback);
}

function resizeTableFrameAxis<T extends ResizableTableFrame>(
  table: T,
  axis: TableFrameAxis,
  targetEmu: number,
): T {
  const count =
    axis === "width"
      ? boundedAxisCount(tableLogicalWidth(table.rows))
      : boundedAxisCount(table.rows.length);
  if (count <= 0) return table;

  const axisKey = axis === "width" ? "colWidths" : "rowHeights";
  const currentEmu =
    typeof table[axis] === "number" &&
    Number.isFinite(table[axis]) &&
    table[axis] > 0
      ? table[axis]
      : null;
  const targetTotalPx = targetEmu / TABLE_FRAME_EMU_PER_CSS_PX;
  const currentTotalPx =
    currentEmu === null
      ? targetTotalPx
      : currentEmu / TABLE_FRAME_EMU_PER_CSS_PX;
  const base = materializedAxisSizes(table[axisKey], count, currentTotalPx);
  const baseTotal = base.reduce((sum, size) => sum + size, 0);
  const scale = baseTotal > 0 ? targetTotalPx / baseTotal : 1;
  const scaled = base.map((size) => {
    const next = clampAxisSize(roundToTenths(size * scale));
    return next ?? MIN_TABLE_AXIS_PX;
  });
  // Keep equal Office columns/rows equal in the web model.  The small aggregate
  // difference introduced by 0.1px display precision is closed by integer-EMU
  // apportionment during PPTX patchback, where the slots sum exactly to the
  // native shape extent.

  return {
    ...table,
    [axis]: targetEmu,
    [axisKey]: scaled,
  };
}

/**
 * The table narrowed to at most `maxWidthEmu`, every column scaled by the
 * same share (Table Size > Width); a table no wider than that is returned as
 * it is.
 */
export function tableNarrowedToWidth<T extends ResizableTableFrame>(table: T, maxWidthEmu: number): T {
  const current = nativeFrameExtent(table, "width");
  if (current === null || !(maxWidthEmu > 0) || current <= maxWidthEmu) return table;
  return resizeTableFrameAxis(table, "width", Math.max(MIN_TABLE_FRAME_EMU, Math.round(maxWidthEmu)));
}

/**
 * Table Layout > Cell Size > Width / Height over a column or row range, as
 * PowerPoint sets it: those columns take the stated size and the others keep
 * the size they draw at, so the table grows or shrinks by the difference,
 * its top-left fixed. The sizes had been weights spread over an unchanged
 * frame, so setting every column from 80 to 160 drew the same table
 * (PPT-B117). A table without a stated frame keeps the weights alone.
 */
export function tableWithAxisSizeRange<T extends ResizableTableFrame>(
  table: T,
  axis: "col" | "row",
  start: number,
  end: number,
  size: number,
): T {
  const frame: TableFrameAxis = axis === "col" ? "width" : "height";
  const axisKey = axis === "col" ? "colWidths" : "rowHeights";
  const count =
    axis === "col"
      ? boundedAxisCount(tableLogicalWidth(table.rows))
      : boundedAxisCount(table.rows.length);
  const sizes = setAxisSizeRange(table[axisKey], count, start, end, size);
  const next = { ...table } as T & Record<string, unknown>;
  if (sizes) next[axisKey] = sizes;
  else delete next[axisKey];
  const stated = clampAxisSize(size);
  const currentEmu = nativeFrameExtent(table, frame);
  if (count <= 0 || stated === null || currentEmu === null) return next;
  // The frame scales with the axis total, so the other columns draw exactly as
  // they did and the stated ones grow or shrink by what changed. Both totals
  // are of the sizes the page holds (no 16 px floor, as `setAxisSizeRange`
  // keeps them): with a floor under the untouched ones the frame grew by less
  // than the grid it holds, and the writer, which scales the grid to the
  
  // 12.25 px columns drew 3.89 cm).
  const currentPx = currentEmu / TABLE_FRAME_EMU_PER_CSS_PX;
  const base = storedMaterializedAxisSizes(table[axisKey], count, currentPx);
  const first = Math.max(0, Math.min(start, end));
  const last = Math.min(count - 1, Math.max(start, end));
  const before = base.reduce((sum, value) => sum + value, 0);
  const after = storedMaterializedAxisSizes(sizes, count, currentPx).reduce(
    (sum, value, index) => sum + (index >= first && index <= last ? stated : value),
    0,
  );
  if (!(before > 0) || !(after > 0)) return next;
  // A grid that sums to its frame but for the 0.1 px the page rounds each
  // size to draws at its sizes: the frame becomes their total, so every
  // column draws as the page states it. A grid stated as weights over a frame
  // of another size scales with it.
  const extent =
    Math.abs(before - currentPx) <= 0.05 * count + 1e-9
      ? after * TABLE_FRAME_EMU_PER_CSS_PX
      : (currentEmu * after) / before;
  next[frame] = Math.min(MAX_TABLE_FRAME_EMU, Math.max(MIN_TABLE_FRAME_EMU, Math.round(extent)));
  return next;
}









export function tableFrameAfterAxisChange(
  before: ResizableTableFrame,
  after: { rows: MergeableCell[][]; colWidths?: unknown; rowHeights?: unknown },
  axis: "col" | "row",
): number | null {
  const frame: TableFrameAxis = axis === "col" ? "width" : "height";
  const axisKey = axis === "col" ? "colWidths" : "rowHeights";
  const currentEmu = nativeFrameExtent(before, frame);
  if (currentEmu === null) return null;
  const countOf = (rows: MergeableCell[][]) =>
    axis === "col" ? boundedAxisCount(tableLogicalWidth(rows)) : boundedAxisCount(rows.length);
  const beforeCount = countOf(before.rows);
  const afterCount = countOf(after.rows);
  if (beforeCount <= 0 || afterCount <= 0) return null;
  const fallbackPx = currentEmu / TABLE_FRAME_EMU_PER_CSS_PX;
  const total = (values: number[]) => values.reduce((sum, value) => sum + value, 0);
  
  // deleted beside ten 12.25 px ones shrank the frame by less than the grid,
  // and every column drew wider than it states).
  const beforeTotal = total(storedMaterializedAxisSizes(before[axisKey], beforeCount, fallbackPx));
  // A slot with no stated size draws its share of the old frame.
  const autoShare = fallbackPx / beforeCount;
  const afterSizes = Array.from({ length: afterCount }, (_, index) => {
    const value = after[axisKey];
    return (Array.isArray(value) && index < value.length ? storedAxisSize(value[index]) : null) ?? autoShare;
  });
  const afterTotal = total(afterSizes);
  if (!(beforeTotal > 0) || !(afterTotal > 0)) return null;
  // As `tableWithAxisSizeRange`: a grid that sums to its frame but for the
  // page's 0.1 px rounding gets a frame of its total.
  const extent =
    Math.abs(beforeTotal - fallbackPx) <= 0.05 * beforeCount + 1e-9
      ? afterTotal * TABLE_FRAME_EMU_PER_CSS_PX
      : (currentEmu * afterTotal) / beforeTotal;
  return Math.min(MAX_TABLE_FRAME_EMU, Math.max(MIN_TABLE_FRAME_EMU, Math.round(extent)));
}

function nativeFrameExtent(
  table: ResizableTableFrame,
  axis: TableFrameAxis,
): number | null {
  const value = table[axis];
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : null;
}

/**
 * Resize a complete native table frame like PowerPoint/WPS Layout → Table Size.
 *
 * Office keeps the frame's top-left fixed and scales the complete matching
 * grid axis (`a:gridCol` for width, `a:tr` for height). With the session-only
 * aspect lock enabled, it scales BOTH axes in the same immutable edit. Exact
 * frame extents stay in EMU while web axis arrays retain 0.1 CSS-pixel
 * precision.
 */
export function resizeTableFrame<T extends ResizableTableFrame>(
  table: T,
  axis: TableFrameAxis,
  centimetres: number,
  options: ResizeTableFrameOptions = {},
): T {
  const requestedTargetEmu = tableFrameEmuFromCm(centimetres);
  if (requestedTargetEmu === null) return table;
  if (!options.lockAspectRatio) {
    return resizeTableFrameAxis(table, axis, requestedTargetEmu);
  }

  const otherAxis: TableFrameAxis = axis === "width" ? "height" : "width";
  const currentPrimary = nativeFrameExtent(table, axis);
  const currentSecondary = nativeFrameExtent(table, otherAxis);
  if (currentPrimary === null || currentSecondary === null) {
    return resizeTableFrameAxis(table, axis, requestedTargetEmu);
  }

  // Keep the native ratio exact until the final integer-EMU rounding. Constrain
  // the shared scale (rather than clamping each axis independently) so an
  // extreme input cannot silently distort the locked aspect ratio.
  const requestedScale = requestedTargetEmu / currentPrimary;
  const minimumScale = Math.max(
    MIN_TABLE_FRAME_EMU / currentPrimary,
    MIN_TABLE_FRAME_EMU / currentSecondary,
  );
  const maximumScale = Math.min(
    MAX_TABLE_FRAME_EMU / currentPrimary,
    MAX_TABLE_FRAME_EMU / currentSecondary,
  );
  const scale = Math.min(maximumScale, Math.max(minimumScale, requestedScale));
  const primaryTarget = Math.max(1, Math.round(currentPrimary * scale));
  const secondaryTarget = Math.max(1, Math.round(currentSecondary * scale));

  const primaryResized = resizeTableFrameAxis(table, axis, primaryTarget);
  return resizeTableFrameAxis(primaryResized, otherAxis, secondaryTarget);
}

/** Approximate one line's visual width without relying on DOM/canvas APIs. */
function estimatedLineWidth(text: string): number {
  let width = 0;
  for (const character of Array.from(text)) {
    const codePoint = character.codePointAt(0) ?? 0;
    // CJK/full-width glyphs are close to one em; ASCII is roughly half an em.
    width += codePoint > 0xff ? 14 : 8;
  }
  return width;
}

function cellContentWidth(cell: MergeableCell): number {
  const longestLine = String(cell.text ?? "")
    .split(/\r?\n/)
    .reduce((largest, line) => Math.max(largest, estimatedLineWidth(line)), 0);
  const left =
    typeof cell.marginLeft === "number" && Number.isFinite(cell.marginLeft)
      ? Math.max(0, cell.marginLeft)
      : DEFAULT_CELL_HORIZONTAL_PADDING_PX / 2;
  const right =
    typeof cell.marginRight === "number" && Number.isFinite(cell.marginRight)
      ? Math.max(0, cell.marginRight)
      : DEFAULT_CELL_HORIZONTAL_PADDING_PX / 2;
  return Math.min(
    MAX_CONTENT_COLUMN_PX,
    Math.max(MIN_CONTENT_COLUMN_PX, Math.ceil(longestLine + left + right)),
  );
}

/**
 * Produce explicit column widths for PowerPoint/WPS table AutoFit commands.
 *
 * - `contents`: estimate the widest native cell in each logical column. Merged
 *   cell demand is shared across every column in its span.
 * - `window`: divide the table's current visual width across all columns.
 * - `fixed`: preserve every explicit width and materialize any remaining auto
 *   slots from the unused table width, so future text edits no longer resize
 *   the column.
 *
 * The result is a full, bounded array suitable for native OOXML `a:gridCol`
 * widths. Invalid/empty tables remain byte-clean (`undefined`).
 */
export function autoFitTableColumnWidths(
  rows: MergeableCell[][],
  mode: TableColumnAutoFitMode,
  tableWidthPx: number,
  current?: unknown,
): Array<number | null> | undefined {
  const count = boundedAxisCount(tableLogicalWidth(rows));
  if (count <= 0 || !TABLE_COLUMN_AUTOFIT_MODES.includes(mode)) {
    return undefined;
  }
  const safeTableWidth = Math.max(
    count * MIN_TABLE_AXIS_PX,
    Math.round(
      Number.isFinite(tableWidthPx)
        ? tableWidthPx
        : count * DEFAULT_DISTRIBUTE_AXIS_PX,
    ),
  );

  if (mode === "contents") {
    const widths = Array.from({ length: count }, () => MIN_CONTENT_COLUMN_PX);
    const positions = logicalCellPositions(rows);
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      for (
        let cellIndex = 0;
        cellIndex < rows[rowIndex].length;
        cellIndex += 1
      ) {
        const position = positions[rowIndex]?.[cellIndex];
        if (!position || position.col >= count) continue;
        const cell = rows[rowIndex][cellIndex];
        const span = Math.min(
          count - position.col,
          Math.max(1, Math.floor(Number(cell.colspan) || 1)),
        );
        const widthPerColumn = Math.ceil(cellContentWidth(cell) / span);
        for (let offset = 0; offset < span; offset += 1) {
          widths[position.col + offset] = Math.max(
            widths[position.col + offset],
            widthPerColumn,
          );
        }
      }
    }
    return widths.map((width) => clampAxisSize(width));
  }

  const distributeAcrossWindow = (): Array<number | null> => {
    const base = Math.floor(safeTableWidth / count);
    const remainder = safeTableWidth - base * count;
    return Array.from({ length: count }, (_unused, index) =>
      clampAxisSize(base + (index < remainder ? 1 : 0)),
    );
  };
  if (mode === "window") return distributeAcrossWindow();

  const explicit = Array.from({ length: count }, (_unused, index) =>
    axisSizeAt(current, index),
  );
  const missing = explicit.filter((width) => width === null).length;
  if (missing === count) return distributeAcrossWindow();
  const used = explicit.reduce<number>(
    (sum, width) => sum + (width === null ? 0 : width),
    0,
  );
  const fallback = clampAxisSize(
    missing > 0
      ? Math.max(
          MIN_TABLE_AXIS_PX,
          Math.floor((safeTableWidth - used) / missing),
        )
      : DEFAULT_DISTRIBUTE_AXIS_PX,
  );
  return explicit.map((width) => width ?? fallback);
}

/**
 * Clamp an arbitrary value to a valid integer axis size, or `null` when it is
 * not a usable size (so the slot falls back to "auto"). A non-finite value, a
 * non-number, or a value <= 0 all collapse to `null`; a real value is rounded
 * and clamped to `MIN_TABLE_AXIS_PX..MAX_TABLE_AXIS_PX`.
 */






function storedAxisSize(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) return null;
  // As the page holds it, unrounded: an imported size is its EMU / 9525
  
  // and a 19-row table ran 0.8 px short of PowerPoint by its last row
  
  // the person did not touch also goes back to the save exactly as the page
  // holds it, so neither this release's API nor the one before reads an
  // edit. Tenths stay the step a person drags or types (`clampAxisSize`).
  return Math.min(MAX_TABLE_AXIS_PX, value);
}

function storedAxisSizeAt(value: unknown, index: number): number | null {
  if (!Array.isArray(value) || index < 0 || index >= value.length) return null;
  return storedAxisSize(value[index]);
}

export function clampAxisSize(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  const rounded = roundToTenths(value);
  if (rounded <= 0) return null;
  return Math.min(MAX_TABLE_AXIS_PX, Math.max(MIN_TABLE_AXIS_PX, rounded));
}

/**
 * Row heights as the browser's table lays them without losing a fraction per
 * row: Chromium sets a row's height in LayoutUnits (1/64 px) and drops what
 * is below one, so a 19.2889 px row drew 19.2813 and a table of 30 such rows
 * ended 0.23 px short of PowerPoint's, which draws the stated EMU exactly
 * (tp-rowquant, TRIAGE 143). Each row here is the difference of its rounded
 * running total, so every row line stands within 1/128 px of where the
 * stated heights put it and no row is cut more than the one before gave it.
 * Non-number slots (auto rows) pass through and restart the running total.
 */
export function layoutUnitAxisSizes<T>(sizes: readonly (number | T)[]): (number | T)[] {
  const unit = 64;
  let total = 0;
  let placed = 0;
  return sizes.map((size) => {
    if (typeof size !== "number" || !Number.isFinite(size) || size <= 0) {
      total = 0;
      placed = 0;
      return size;
    }
    total += size;
    const end = Math.round(total * unit) / unit;
    const height = end - placed;
    placed = end;
    return height;
  });
}

/**
 * Normalize an arbitrary persisted value into a bounded axis-size array, or
 * `undefined` when it carries no usable size at all. The array is truncated to
 * `length` (the real row/column count) and `MAX_TABLE_AXIS_ENTRIES`; each slot
 * is a clamped size or `null` (auto). TRAILING auto (`null`) slots are dropped so
 * a table that sizes only its first column doesn't persist a tail of nulls, and
 * an all-auto array collapses to `undefined` (byte-stable / no phantom field).
 */
export function normalizeAxisSizes(
  value: unknown,
  length: number,
): Array<number | null> | undefined {
  if (!Array.isArray(value)) return undefined;
  const bound = Math.min(
    Math.max(0, Math.floor(length)),
    MAX_TABLE_AXIS_ENTRIES,
  );
  const out: Array<number | null> = [];
  for (let index = 0; index < bound; index += 1) {
    out.push(storedAxisSize(value[index]));
  }
  // Drop trailing autos so the persisted array is as short as possible.
  while (out.length > 0 && out[out.length - 1] === null) out.pop();
  return out.length > 0 ? out : undefined;
}

/**
 * Set ONE axis slot (`index`) to `size` (or clear it to auto when `size` is not a
 * usable value), returning a NEW bounded array. The array is grown with `null`
 * (auto) padding up to `index` when needed, then re-normalized so trailing autos
 * are trimmed. An out-of-range index (negative or beyond `MAX_TABLE_AXIS_ENTRIES`)
 * returns the input normalized unchanged. The result is `undefined` when the
 * table ends up fully auto again (so clearing the only sized slot drops the
 * field entirely — byte-stable).
 */
export function setAxisSize(
  current: unknown,
  index: number,
  size: number | null,
): Array<number | null> | undefined {
  const base = Array.isArray(current) ? [...current] : [];
  if (index < 0 || index >= MAX_TABLE_AXIS_ENTRIES) {
    return normalizeAxisSizes(base, base.length);
  }
  while (base.length <= index) base.push(null);
  base[index] = clampAxisSize(size);
  // Re-normalize over the FULL (possibly grown) length so the clamp + trailing
  // trim run uniformly.
  return normalizeAxisSizes(base, base.length);
}

/**
 * Resolve one row/column slot to the concrete size used when a pointer drag
 * begins. Explicit sizes win; an auto slot receives an equal share of the
 * unused native table-frame extent. This mirrors the materialization used by
 * frame resizing, so grabbing an auto boundary never makes the column/row jump
 * before the pointer has moved.
 */
export function materializedAxisSizeAt(
  current: unknown,
  count: number,
  index: number,
  totalSizePx: number,
): number | null {
  const bound = boundedAxisCount(count);
  if (bound <= 0 || !Number.isInteger(index) || index < 0 || index >= bound) {
    return null;
  }
  const safeTotal =
    Number.isFinite(totalSizePx) && totalSizePx > 0
      ? totalSizePx
      : bound * DEFAULT_DISTRIBUTE_AXIS_PX;
  return clampAxisSize(materializedAxisSizes(current, bound, safeTotal)[index]);
}

/**
 * Convert a client-pixel pointer delta into a native table-axis size. The
 * caller supplies the table's layout-px/client-px scale, keeping the gesture
 * correct when the editor surface is transformed or zoomed. Invalid
 * deltas/scales are a no-op; valid values retain the same 0.1px precision and
 * bounds as numeric ribbon edits.
 */
export function axisSizeFromPointerDrag(
  startSize: unknown,
  clientDeltaPx: number,
  modelPxPerClientPx: number,
): number | null {
  const start = clampAxisSize(startSize);
  if (start === null) return null;
  if (
    !Number.isFinite(clientDeltaPx) ||
    !Number.isFinite(modelPxPerClientPx) ||
    modelPxPerClientPx <= 0
  ) {
    return start;
  }
  return clampAxisSize(
    Math.max(MIN_TABLE_AXIS_PX, start + clientDeltaPx * modelPxPerClientPx),
  );
}

function boundedAxisRange(
  count: number,
  start: number,
  end: number,
): { count: number; start: number; end: number } | null {
  const bound = boundedAxisCount(count);
  if (bound <= 0 || !Number.isFinite(start) || !Number.isFinite(end)) {
    return null;
  }
  const first = Math.min(Math.floor(start), Math.floor(end));
  const last = Math.max(Math.floor(start), Math.floor(end));
  const rangeStart = Math.min(bound - 1, Math.max(0, first));
  const rangeEnd = Math.min(bound - 1, Math.max(rangeStart, last));
  return { count: bound, start: rangeStart, end: rangeEnd };
}

/**
 * Set every slot in one contiguous selected row/column range to the same size.
 * Slots outside the range stay untouched; trailing auto entries remain trimmed.
 */
export function setAxisSizeRange(
  current: unknown,
  count: number,
  start: number,
  end: number,
  size: number | null,
): Array<number | null> | undefined {
  const range = boundedAxisRange(count, start, end);
  if (!range) return normalizeAxisSizes(current, boundedAxisCount(count));
  // The slots outside the range keep the size the page holds (`storedAxisSize`,
  // no 16 px floor): Cell Size on one row of an imported table raised its
  
  
  const slots = Array.from({ length: range.count }, (_, index) =>
    Array.isArray(current) && index < current.length ? storedAxisSize(current[index]) : null,
  );
  const nextSize = clampAxisSize(size);
  for (let index = range.start; index <= range.end; index += 1) {
    slots[index] = nextSize;
  }
  return normalizeAxisSizes(slots, range.count);
}

/**
 * Read the size at `index` from a (possibly absent) axis-size array, or `null`
 * (auto) when the slot is absent / out of range / not a usable size. A pure
 * read used by the renderer + inspector. `value` is re-clamped so a hand-edited
 * persisted file can never inject an out-of-bounds size into the render.
 */
export function axisSizeAt(value: unknown, index: number): number | null {
  if (!Array.isArray(value) || index < 0 || index >= value.length) return null;
  const entry = value[index];
  if (typeof entry !== "number" || !Number.isFinite(entry) || entry <= 0) return null;
  // The size the page holds, unrounded (`storedAxisSize`), within the bounds a
  // person can set.
  return Math.min(MAX_TABLE_AXIS_PX, Math.max(MIN_TABLE_AXIS_PX, entry));
}

export interface TableAxisSizeSelectionState {
  mixed: boolean;
  value: number | null;
}

/**
 * Resolve the honest Office size-box state for one selected axis range.
 * Uniform auto remains `{ value: null, mixed: false }`; any explicit-size or
 * auto/explicit disagreement is mixed, so the UI can distinguish a real Auto
 * value from a blank mixed selection without inspecting persisted arrays.
 */
export function axisSizeSelectionState(
  current: unknown,
  count: number,
  start: number,
  end: number,
): TableAxisSizeSelectionState | null {
  const range = boundedAxisRange(count, start, end);
  if (!range) return null;
  const value = axisSizeAt(current, range.start);
  for (let index = range.start + 1; index <= range.end; index += 1) {
    if (!Object.is(axisSizeAt(current, index), value)) {
      return { value, mixed: true };
    }
  }
  return { value, mixed: false };
}

/**
 * Read the shared explicit size for a selected range. A mixed or partly-auto
 * selection returns `null`, matching Office's blank size box for mixed values.
 */
export function commonAxisSizeInRange(
  current: unknown,
  count: number,
  start: number,
  end: number,
): number | null {
  const state = axisSizeSelectionState(current, count, start, end);
  return state && !state.mixed ? state.value : null;
}

/**
 * Insert one row/column size slot. When the axis carries explicit sizing, the
 * new slot inherits the focused neighbour's size (PowerPoint/WPS behaviour);
 * a fully-auto axis remains absent and byte-clean.
 */
export function insertAxisSize(
  current: unknown,
  count: number,
  index: number,
  sourceIndex: number,
): Array<number | null> | undefined {
  const bound = Math.min(
    Math.max(0, Math.floor(count)),
    MAX_TABLE_AXIS_ENTRIES,
  );
  if (bound <= 0 || !Number.isInteger(index) || index < 0 || index > bound) {
    return normalizeAxisSizes(current, bound);
  }
  // As `deleteAxisSize`: the sizes the page holds, a new slot its source's.
  const slots = Array.from({ length: bound }, (_, slot) => storedAxisSizeAt(current, slot));
  if (!slots.some((size) => size !== null)) return undefined;
  const inherited = storedAxisSizeAt(current, sourceIndex);
  slots.splice(index, 0, inherited);
  return normalizeAxisSizes(slots, bound + 1);
}

/**
 * Subdivide one explicit row/column size into `parts` equal native grid slots.
 *
 * PowerPoint/WPS "Split Cells" inserts deck-wide grid lines without resizing
 * the table frame, so the focused slot's total is distributed across the new
 * slots. A fully-auto axis remains absent; an auto target in a partially-sized
 * axis produces auto children while preserving every other explicit slot.
 */
export function splitAxisSize(
  current: unknown,
  count: number,
  index: number,
  parts: number,
): Array<number | null> | undefined {
  const bound = Math.min(
    Math.max(0, Math.floor(count)),
    MAX_TABLE_AXIS_ENTRIES,
  );
  const splitParts = Math.floor(parts);
  if (
    bound <= 0 ||
    !Number.isInteger(index) ||
    index < 0 ||
    index >= bound ||
    !Number.isFinite(parts) ||
    splitParts < 1 ||
    bound + splitParts - 1 > MAX_TABLE_AXIS_ENTRIES
  ) {
    return normalizeAxisSizes(current, bound);
  }
  const slots = Array.from({ length: bound }, (_, slot) =>
    axisSizeAt(current, slot),
  );
  if (!slots.some((size) => size !== null)) return undefined;
  const source = slots[index];
  let pieces: Array<number | null>;
  if (source === null || splitParts === 1) {
    pieces = Array.from({ length: splitParts }, () => source);
  } else {
    const shared = Math.round((source / splitParts) * 10) / 10;
    pieces = Array.from({ length: splitParts }, () => shared);
    pieces[pieces.length - 1] =
      Math.round((source - shared * (splitParts - 1)) * 10) / 10;
  }
  slots.splice(index, 1, ...pieces);
  return normalizeAxisSizes(slots, bound + splitParts - 1);
}

/** Delete one row/column size slot and shift every later explicit size left. */
export function deleteAxisSize(
  current: unknown,
  count: number,
  index: number,
): Array<number | null> | undefined {
  const bound = Math.min(
    Math.max(0, Math.floor(count)),
    MAX_TABLE_AXIS_ENTRIES,
  );
  if (bound <= 0 || !Number.isInteger(index) || index < 0 || index >= bound) {
    return normalizeAxisSizes(current, bound);
  }
  // The slots left keep the size the page holds, no 16 px floor, as
  
  // the ten 12.25 px ones beside it to 16 px).
  const slots = Array.from({ length: bound }, (_, slot) => storedAxisSizeAt(current, slot));
  slots.splice(index, 1);
  return normalizeAxisSizes(slots, bound - 1);
}

/**
 * Equalize an axis so EVERY slot (column width / row height) is the same size —
 * PowerPoint / WPS "Distribute Columns" / "Distribute Rows" (equal width /
 * equal height). Returns a NEW bounded array of `count` identical sizes, or
 * `undefined` when `count` is non-positive (so a 0-cell table stays auto /
 * byte-clean).
 *
 * The shared size is the LARGEST currently-explicit slot (so distributing a
 * table whose tallest row is already sized keeps the content visible rather
 * than shrinking it); when NO slot is sized yet, every slot falls back to
 * {@link DEFAULT_DISTRIBUTE_AXIS_PX}. The chosen value is clamped to the axis
 * bounds. `count` is bounded by {@link MAX_TABLE_AXIS_ENTRIES}.
 *
 * Unlike {@link normalizeAxisSizes}, the returned array is NOT trailing-trimmed:
 * distribute means "all slots equal", so an explicit, full-length array is the
 * point (a trimmed tail would let those rows/cols fall back to auto and break
 * the even distribution).
 */
export function equalizedAxisSizes(
  current: unknown,
  count: number,
): Array<number | null> | undefined {
  const bound = Math.min(
    Math.max(0, Math.floor(count)),
    MAX_TABLE_AXIS_ENTRIES,
  );
  if (bound <= 0) return undefined;
  // The largest currently-sized slot drives the equal size; fall back to the
  // default when the table carries no explicit size at all.
  let largest = 0;
  if (Array.isArray(current)) {
    for (const entry of current) {
      const size = clampAxisSize(entry);
      if (size !== null && size > largest) largest = size;
    }
  }
  const size = clampAxisSize(
    largest > 0 ? largest : DEFAULT_DISTRIBUTE_AXIS_PX,
  );
  // `size` is non-null here (both inputs are positive + clamped), but guard so
  // the array is well-typed.
  return Array.from({ length: bound }, () => size);
}

/**
 * Distribute only a contiguous selected range while preserving every slot
 * outside it. The selected total is preserved when sizes are explicit by using
 * their average; auto slots materialize from the same Office fallback used by
 * full-axis distribution.
 */
export function equalizedAxisSizeRange(
  current: unknown,
  count: number,
  start: number,
  end: number,
): Array<number | null> | undefined {
  const range = boundedAxisRange(count, start, end);
  if (!range) return normalizeAxisSizes(current, boundedAxisCount(count));
  const slots = Array.from({ length: range.count }, (_, index) =>
    axisSizeAt(current, index),
  );
  let total = 0;
  let selectedCount = 0;
  for (let index = range.start; index <= range.end; index += 1) {
    total += slots[index] ?? DEFAULT_DISTRIBUTE_AXIS_PX;
    selectedCount += 1;
  }
  const sharedSize =
    clampAxisSize(total / Math.max(1, selectedCount)) ??
    DEFAULT_DISTRIBUTE_AXIS_PX;
  for (let index = range.start; index <= range.end; index += 1) {
    slots[index] = sharedSize;
  }
  return normalizeAxisSizes(slots, range.count);
}
