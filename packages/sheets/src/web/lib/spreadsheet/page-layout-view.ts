/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pxToPt } from "@mrnote/core/web/lib/office-units.ts";
import { aspectLockedResizeSize } from "./drawing-aspect";
import {
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import {
  computePageBreakLines,
  normalizePrintSettings,
  paginateSheet,
  parseA1Range,
  printAreaRanges,
  resolveMargins,
  resolvePaperSizePoints,
  substituteHeaderFooterTokens,
  HEADER_FOOTER_PICTURE_MARK,
  type HeaderFooterContext,
  type PaginateInput,
  type PageBreakLines,
  type PageBreakResult,
  type PrintHeaderFooter,
  type PrintHeaderFooterPicture,
  type PrintHeaderFooterRun,
  type PrintMargins,
  type PrintSettings,
} from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";

/** The grid/PDF bridge already used by Page Break Preview: 96 CSS px = 72 pt. */
export const SPREADSHEET_PAGE_LAYOUT_PX_TO_PT = pxToPt(1);
export const SPREADSHEET_PAGE_LAYOUT_DEFAULT_ROW_HEIGHT_PT =
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX * SPREADSHEET_PAGE_LAYOUT_PX_TO_PT;
export const SPREADSHEET_PAGE_LAYOUT_DEFAULT_COLUMN_WIDTH_PT =
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX * SPREADSHEET_PAGE_LAYOUT_PX_TO_PT;

/** Keep a pathological 10k × 1k sheet from mounting millions of paper nodes. */
export const MAX_RENDERED_PAGE_LAYOUT_PAGES = 120;
export const MAX_RENDERED_PAGE_LAYOUT_CELLS_PER_PAGE = 4_096;
export const MAX_RENDERED_PAGE_LAYOUT_CELLS_TOTAL = 24_000;

export function pageLayoutEdgeScrollStep(point: number, start: number, end: number) {
  return point < start + 32 ? -Math.min(16, (start + 32 - point) / 2)
    : point > end - 32 ? Math.min(16, (point - end + 32) / 2) : 0;
}

/** The identity order of the last row count asked for, shared: a sheet
 * without a Sheet View's order reads its rows in place. */
let identityOrder: readonly number[] = [];
function identityRows(count: number): readonly number[] {
  if (identityOrder.length !== count) identityOrder = Object.freeze(Array.from({ length: count }, (_, row) => row));
  return identityOrder;
}

/** A Sheet View is a permutation, never a rewrite of the worksheet rows.
 * Built lazily: a sheet in its own order maps every row and offset to
 * itself, and only a Sheet View's order (or a read of the key) walks the
 * rows. The grid rebuilds this map whenever a row height or a hidden row
 * changes, which on a 50k-row sheet was two 50k arrays, a 50k Map and a
 * JSON string of every row height each time. */
export function createPageLayoutRowMap(
  count: number,
  requestedOrder: readonly number[] | null | undefined,
  size: (row: number) => number,
) {
  const candidate = requestedOrder?.slice(0, count);
  const permuted = candidate?.length === count && new Set(candidate).size === count &&
    candidate.every((row) => Number.isInteger(row) && row >= 0 && row < count) &&
    candidate.some((row, visual) => row !== visual);
  let geometry: {
    readonly positions: ReadonlyMap<number, number>;
    readonly sourceOffsets: readonly number[];
    readonly visualOffsets: readonly number[];
  } | null = null;
  let key: string | null = null;
  const order = (): readonly number[] => (permuted ? candidate! : identityRows(count));
  const sourceSize = (row: number) => Math.max(0, size(row));
  const geometryOf = () => {
    if (geometry) return geometry;
    const rows = order();
    const positions = new Map(rows.map((row, visual) => [row, visual]));
    const sourceSizes = Array.from({ length: count }, (_, row) => sourceSize(row));
    const prefix = (sizes: readonly number[]) => {
      const result = [0];
      for (const value of sizes) result.push(result[result.length - 1] + value);
      return result;
    };
    geometry = { positions, sourceOffsets: prefix(sourceSizes), visualOffsets: prefix(rows.map((row) => sourceSizes[row])) };
    return geometry;
  };
  const convert = (value: number, from: readonly number[], to: readonly number[], index: (row: number) => number) => {
    const point = Math.max(0, value);
    if (count === 0 || point >= from[count]) return point;
    let low = 0, high = count;
    while (low + 1 < high) {
      const mid = (low + high) >>> 1;
      if (from[mid] <= point) low = mid;
      else high = mid;
    }
    return to[index(low)] + point - from[low];
  };
  return {
    get order(): readonly number[] { return order(); },
    /** Equal for two maps exactly when their orders and row sizes are: the
     * sizes as runs ([size, length]...), the order only when permuted. */
    get key(): string {
      if (key !== null) return key;
      const runs: number[] = [];
      for (let row = 0; row < count; row += 1) {
        const value = sourceSize(row);
        if (runs.length && Object.is(runs[runs.length - 2], value)) runs[runs.length - 1] += 1;
        else runs.push(value, 1);
      }
      key = JSON.stringify([count, permuted ? candidate : null, runs]);
      return key;
    },
    sourceRow: (visual: number) => (permuted ? candidate![visual] ?? visual : visual),
    visualRow: (source: number) => (permuted ? geometryOf().positions.get(source) ?? source : source),
    // In the sheet's own order every offset is where it stands.
    toVisual: (point: number) => {
      if (!permuted) return Math.max(0, point);
      const { sourceOffsets, visualOffsets, positions } = geometryOf();
      return convert(point, sourceOffsets, visualOffsets, (row) => positions.get(row) ?? row);
    },
    toSource: (point: number) => {
      if (!permuted) return Math.max(0, point);
      const { sourceOffsets, visualOffsets } = geometryOf();
      return convert(point, visualOffsets, sourceOffsets, (row) => candidate![row] ?? row);
    },
  };
}

export interface PageLayoutDrawingFrame {
  left: number;
  top: number;
  width: number;
  height: number;
}
export type PageLayoutDrawingHandle = "north-west" | "north" | "north-east" | "east" | "south-east" | "south" | "south-west" | "west";
export const PAGE_LAYOUT_DRAWING_HANDLES: readonly PageLayoutDrawingHandle[] = ["north-west", "north", "north-east", "east", "south-east", "south", "south-west", "west"];

/** Source bands keep repeated print titles separate from the page's data. */
export function pageLayoutAxisBands(sizes: readonly number[], indices: readonly number[]) {
  const bands: Array<{ start: number; end: number; offset: number }> = [];
  const prefix = [0];
  for (const size of sizes) prefix.push(prefix[prefix.length - 1] + size);
  let offset = 0;
  for (const index of indices) {
    const start = prefix[index] ?? 0, end = start + (sizes[index] ?? 0);
    const previous = bands.at(-1);
    if (previous && Math.abs(previous.end - start) < 0.00001) previous.end = end;
    else bands.push({ start, end, offset });
    offset += end - start;
  }
  return bands;
}

/** Without an explicit print area, the final data line is not a drawing
 * clipping boundary. Use the remaining printable paper for its last band;
 * repeated-title bands and intermediate page boundaries stay unchanged. */
export function pageLayoutDrawingBands(sizes: readonly number[], indices: readonly number[], availablePoints: number, finalUnboundedPage: boolean) {
  const bands = pageLayoutAxisBands(sizes, indices);
  const last = bands.at(-1);
  if (last && finalUnboundedPage && Number.isFinite(availablePoints)) {
    last.end = Math.max(last.end, last.start + Math.max(0, availablePoints - last.offset));
  }
  return bands;
}

/** Inverse of drawing clipping. Paper margins/gaps never become sheet offsets. */
export function pageLayoutAxisSourceAt(position: number, bands: ReturnType<typeof pageLayoutAxisBands>) {
  if (!bands.length) return 0;
  const last = bands[bands.length - 1];
  const clamped = Math.max(0, Math.min(position, last.offset + last.end - last.start));
  const band = bands.find((entry) => clamped < entry.offset + entry.end - entry.start) ?? last;
  return band.start + clamped - band.offset;
}

export function pageLayoutDrawingHandlePoint(frame: PageLayoutDrawingFrame, handle: PageLayoutDrawingHandle, rotation: number) {
  const x = handle.includes("west") ? -frame.width / 2 : handle.includes("east") ? frame.width / 2 : 0;
  const y = handle.includes("north") ? -frame.height / 2 : handle.includes("south") ? frame.height / 2 : 0;
  const angle = rotation * Math.PI / 180;
  return { x: frame.left + frame.width / 2 + x * Math.cos(angle) - y * Math.sin(angle), y: frame.top + frame.height / 2 + x * Math.sin(angle) + y * Math.cos(angle) };
}

/**
 * Resize in the object's axes, keeping the opposite rotated edge fixed. With
 * `aspect` (width over height: Lock aspect ratio, or Shift) the proportions
 * hold by the Normal view's rule; a side handle grows the other axis about
 * the centre.
 */
export function transformPageLayoutDrawing(frame: PageLayoutDrawingFrame, rotation: number, handle: PageLayoutDrawingHandle | undefined, dx: number, dy: number, constrainSize?: (width: number, height: number) => { width: number; height: number }, aspect?: number | null): PageLayoutDrawingFrame {
  if (!handle) return { ...frame, left: Math.max(0, frame.left + dx), top: Math.max(0, frame.top + dy) };
  const angle = rotation * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle);
  const localX = dx * cos + dy * sin, localY = -dx * sin + dy * cos;
  const xSign = handle.includes("west") ? -1 : handle.includes("east") ? 1 : 0;
  const ySign = handle.includes("north") ? -1 : handle.includes("south") ? 1 : 0;
  const at = (fraction: number) => {
    let proposedWidth = xSign ? Math.max(32, frame.width + xSign * localX * fraction) : frame.width;
    let proposedHeight = ySign ? Math.max(24, frame.height + ySign * localY * fraction) : frame.height;
    if (aspect && aspect > 0) {
      ({ width: proposedWidth, height: proposedHeight } = aspectLockedResizeSize(
        frame,
        { width: proposedWidth, height: proposedHeight },
        { horizontal: xSign !== 0, vertical: ySign !== 0 },
        aspect,
        { width: 32, height: 24 },
      ));
    }
    const { width, height } = constrainSize?.(proposedWidth, proposedHeight) ?? { width: proposedWidth, height: proposedHeight };
    const centerX = xSign * (width - frame.width) / 2, centerY = ySign * (height - frame.height) / 2;
    return { left: frame.left + (frame.width - width) / 2 + centerX * cos - centerY * sin, top: frame.top + (frame.height - height) / 2 + centerX * sin + centerY * cos, width, height };
  };
  const proposed = at(1);
  if (proposed.left >= 0 && proposed.top >= 0) return proposed;
  // Stop at the worksheet origin instead of allowing anchor serialization to
  // silently clamp a negative offset and shift the opposite resize edge.
  let low = 0, high = 1, result = frame;
  for (let step = 0; step < 24; step += 1) {
    const fraction = (low + high) / 2, candidate = at(fraction);
    if (candidate.left >= 0 && candidate.top >= 0) { low = fraction; result = candidate; }
    else high = fraction;
  }
  return result;
}

interface PageLayoutAxisProjection {
  visibleIndices?: readonly number[];
  repeatStart: number;
  repeatCount: number;
  dataStart: number;
  dataCount: number;
  paginateCount: number;
}

export interface SpreadsheetPageLayoutGeometry {
  settings: PrintSettings;
  paperWidth: number;
  paperHeight: number;
  margins: PrintMargins;
  sourceMargins?: PrintMargins;
  headerMargin: number;
  footerMargin: number;
  dataAreaWidth: number;
  dataAreaHeight: number;
  rowHeight: number;
  colWidth: number;
  rowSizes: readonly number[];
  colSizes: readonly number[];
  contentScale: number;
  /** Printed row and column headings (Page Setup > Sheet), in points at the
   * content scale: the column-letter band's height and the row-number
   * column's width. Zero when the sheet does not print them. */
  headingHeight: number;
  headingWidth: number;
  paginateInput: PaginateInput;
  pagination: PageBreakResult;
  totalPages: number;
  rowProjection: PageLayoutAxisProjection;
  colProjection: PageLayoutAxisProjection;
  /** Set by the browsing extension: the page rows/columns that hold content.
   * Pages beyond them are blank paper the user may scroll into; they are
   * planned after every content page so they never crowd one out of the cell
   * budget. Absent when every page is content. */
  contentPages?: { rowPages: number; colPages: number };
  /** Excel prints each area of a multi-area print area as its own pages, in
   * the order the areas were added ("Set or clear a print area on a
   * worksheet"). A composite geometry keeps one single-area geometry per
   * block here; its own pagination/projection fields mirror the first block
   * and `totalPages` spans them all. Page numbering runs on across blocks. */
  segments?: readonly SpreadsheetPageLayoutGeometry[];
  /** Cumulative first page index of each segment, plus the total at the end. */
  segmentPageOffsets?: readonly number[];
}

export interface SpreadsheetPageLayoutPage {
  index: number;
  pageNumber: number;
  rowRange: [number, number];
  colRange: [number, number];
  rows: number[];
  cols: number[];
  totalCellCount: number;
  renderedCellCount: number;
  cellsTruncated: boolean;
}

export interface SpreadsheetPageLayoutRenderPlan {
  pages: SpreadsheetPageLayoutPage[];
  renderedCellCount: number;
  pageLimitReached: boolean;
  cellLimitReached: boolean;
}

export interface ResolvedHeaderFooterSegment {
  text: string;
  runs?: PrintHeaderFooterRun[];
  /**
   * The section's picture, drawn where the text (or a run) holds
   * {@link HEADER_FOOTER_PICTURE_MARK}: the first `&[Picture]` code only, as
   * a section holds one picture.
   */
  picture?: PrintHeaderFooterPicture;
}

function finitePositive(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : fallback;
}

function finiteNonNegative(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? value
    : fallback;
}

function boundedDimension(value: unknown): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value)) return 1;
  return Math.max(1, value);
}

/** Drawings enlarge only the disposable print domain. Bound that expansion
 * independently of authored dimensions so a bad coordinate cannot allocate
 * millions of phantom cells before the render-plan limits apply. */
function drawingAxisCount(total: number, extent: number, maximum: number, sizeAt: (index: number) => number): number {
  if (!(extent > 0)) return total;
  const limit = Math.max(total, Math.min(maximum, MAX_RENDERED_PAGE_LAYOUT_CELLS_TOTAL));
  let count = 0, position = 0;
  while (count < limit && position + 1e-7 < extent) {
    position += sizeAt(count);
    count += 1;
  }
  return Math.max(total, count);
}

function clampIndex(value: number, maximum: number): number {
  return Math.max(0, Math.min(maximum, value));
}

function columnLabel(index: number): string {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function resolveAxisProjection(input: {
  total: number;
  areaStart: number;
  areaEnd: number;
  repeatStart?: number;
  repeatEnd?: number;
  hidden?: ReadonlySet<number>;
  order?: readonly number[];
}): PageLayoutAxisProjection {
  const maximum = input.total - 1;
  const areaStart = clampIndex(input.areaStart, maximum);
  const areaEnd = Math.max(areaStart, clampIndex(input.areaEnd, maximum));
  const hasRepeat =
    typeof input.repeatStart === "number" &&
    typeof input.repeatEnd === "number";
  const repeatStart = hasRepeat
    ? clampIndex(input.repeatStart!, maximum)
    : 0;
  const repeatEnd = hasRepeat
    ? Math.max(repeatStart, clampIndex(input.repeatEnd!, maximum))
    : -1;
  const repeatCount = hasRepeat ? repeatEnd - repeatStart + 1 : 0;
  // When the print area starts inside the title band, those rows/columns are
  // represented by the repeat prefix and must not be emitted a second time.
  const dataStart =
    repeatCount > 0 && repeatStart <= areaStart && repeatEnd >= areaStart
      ? Math.min(areaEnd + 1, repeatEnd + 1)
      : areaStart;
  const dataCount = Math.max(0, areaEnd - dataStart + 1);
  if (input.order) {
    const repeatIndices: number[] = [], dataIndices: number[] = [];
    input.order.forEach((source, visual) => {
      if (input.hidden?.has(visual)) return;
      if (hasRepeat && source >= repeatStart && source <= repeatEnd) repeatIndices.push(visual);
      else if (source >= areaStart && source <= areaEnd) dataIndices.push(visual);
    });
    return { repeatStart: repeatIndices[0] ?? 0, dataStart: dataIndices[0] ?? 0,
      repeatCount: repeatIndices.length, dataCount: dataIndices.length,
      paginateCount: repeatIndices.length + Math.max(1, dataIndices.length), visibleIndices: [...repeatIndices, ...dataIndices] };
  }
  if (input.hidden?.size) {
    const repeatIndices = Array.from({ length: repeatCount }, (_, i) => repeatStart + i).filter(i => !input.hidden!.has(i));
    const dataIndices = Array.from({ length: dataCount }, (_, i) => dataStart + i).filter(i => !input.hidden!.has(i));
    return {
      repeatStart, dataStart, repeatCount: repeatIndices.length, dataCount: dataIndices.length,
      paginateCount: repeatIndices.length + Math.max(1, dataIndices.length),
      visibleIndices: [...repeatIndices, ...dataIndices],
    };
  }
  return {
    repeatStart,
    repeatCount,
    dataStart,
    dataCount,
    // paginateSheet requires a positive domain. The one phantom position is
    // filtered by mapLocalAxisIndex when the area is entirely a title band.
    paginateCount: repeatCount + Math.max(1, dataCount),
  };
}

function mapLocalAxisIndex(
  projection: PageLayoutAxisProjection,
  localIndex: number,
): number | null {
  if (localIndex < 0) return null;
  if (projection.visibleIndices) return projection.visibleIndices[localIndex] ?? null;
  if (localIndex < projection.repeatCount) {
    return projection.repeatStart + localIndex;
  }
  const dataOffset = localIndex - projection.repeatCount;
  return dataOffset >= 0 && dataOffset < projection.dataCount
    ? projection.dataStart + dataOffset
    : null;
}

function mapActualAxisIndex(
  projection: PageLayoutAxisProjection,
  actualIndex: number,
): number | null {
  if (projection.visibleIndices) {
    const index = projection.visibleIndices.indexOf(actualIndex);
    return index >= 0 ? index : null;
  }
  if (
    actualIndex >= projection.repeatStart &&
    actualIndex < projection.repeatStart + projection.repeatCount
  ) {
    return actualIndex - projection.repeatStart;
  }
  if (
    actualIndex >= projection.dataStart &&
    actualIndex < projection.dataStart + projection.dataCount
  ) {
    return projection.repeatCount + actualIndex - projection.dataStart;
  }
  return null;
}

function pageCountAtScale(input: PaginateInput, scale: number): {
  pagination: PageBreakResult;
  pageCount: number;
} {
  const pagination = paginateSheet({
    ...input,
    rowHeight: input.rowHeight * scale,
    colWidth: input.colWidth * scale,
    rowSizes: input.rowSizes?.map((size) => size * scale),
    colSizes: input.colSizes?.map((size) => size * scale),
    headingHeight: (input.headingHeight ?? 0) * scale,
    headingWidth: (input.headingWidth ?? 0) * scale,
  });
  return {
    pagination,
    pageCount: Math.max(
      1,
      pagination.rowRanges.length * pagination.colRanges.length,
    ),
  };
}

/** Resolve fit policies by asking paginateSheet itself, never a second pager. */
function resolveContentScale(input: PaginateInput, override?: number): number {
  if (typeof override === "number" && Number.isFinite(override) && override > 0) {
    return Math.max(0.1, Math.min(4, override));
  }
  const scaling = input.settings.scaling;
  if (!scaling) return 1;
  if (scaling.kind === "percent") {
    return Math.max(0.1, Math.min(4, scaling.percent / 100));
  }
  const fits = (scale: number) => {
    const result = pageCountAtScale(input, scale);
    if (scaling.kind === "fitToWidth") {
      return result.pagination.colRanges.length <= Math.max(1, scaling.pages);
    }
    if ("pages" in scaling) {
      return result.pageCount <= Math.max(1, scaling.pages);
    }
    return (
      (scaling.widthPages === 0 ||
        result.pagination.colRanges.length <= scaling.widthPages) &&
      (scaling.heightPages === 0 ||
        result.pagination.rowRanges.length <= scaling.heightPages)
    );
  };
  // Excel Fit To shrinks only, and its print scale remains within 10–400%.
  // Manual breaks are preserved in settings but ignored by paginateSheet here.
  let low = 0.1;
  let high = 1;
  if (!fits(low)) return low;
  if (fits(high)) return high;
  // Page counts are monotonic in scale. A fixed iteration bound keeps fit
  // resolution deterministic and cannot hang on malformed workbook geometry.
  for (let iteration = 0; iteration < 28; iteration += 1) {
    const middle = (low + high) / 2;
    if (fits(middle)) low = middle;
    else high = middle;
  }
  // Excel prints a Fit To sheet at a whole percentage (the Adjust to N %
  // the Page Setup then shows; PageSetup.Zoom is a whole 10-400): the largest
  // one that fits, so the content stops short of the margins by the
  // fraction (P01: 94 %, 797 of 806 pt; a 14-column probe at 62.0 %).
  return Math.max(0.1, Math.floor(low * 100 + 1e-9) / 100);
}

/**
 * Resolve the one pagination authority shared by Page Layout and Page Break
 * Preview. It deliberately feeds the existing paginateSheet helper with the
 * same 0.75 pt/px grid dimensions, so the two views cannot drift by inventing
 * separate paper math.
 */
export function resolveSpreadsheetPageLayoutGeometry(input: {
  totalRows: number;
  rowOrder?: readonly number[];
  totalCols: number;
  hiddenRows?: ReadonlySet<number>;
  hiddenCols?: ReadonlySet<number>;
  printSettings?: PrintSettings;
  rowHeight?: number;
  colWidth?: number;
  rowHeights?: Readonly<Record<string, number>>;
  columnWidths?: Readonly<Record<string, number>>;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  nativeGridDpi?: number;
  /** The grid font's maximum digit width in device pixels at nativeGridDpi
   * (officeDigitWidth; 7 for Calibri 11 at 96 dpi): what the printed
   * row-number column is counted in. */
  nativeDigitWidth?: number;
  /** Visible drawing AABBs in unscaled worksheet CSS pixels. */
  drawingBounds?: readonly (readonly [left: number, top: number, width: number, height: number])[];
  /** Internal: a multi-area composite resolves every block at one scale. */
  contentScaleOverride?: number;
}): SpreadsheetPageLayoutGeometry {
  const settings = normalizePrintSettings(input.printSettings) ?? {};
  const printAreaBlocks = printAreaRanges(settings.printArea);
  if (printAreaBlocks.length > 1) {
    return resolveMultiAreaPageLayoutGeometry(input, settings, printAreaBlocks);
  }
  const baseRowHeight = finitePositive(
    finitePositive(input.defaultRowHeight, input.rowHeight ?? DEFAULT_SPREADSHEET_ROW_HEIGHT_PX) *
      SPREADSHEET_PAGE_LAYOUT_PX_TO_PT,
    SPREADSHEET_PAGE_LAYOUT_DEFAULT_ROW_HEIGHT_PT,
  );
  const baseColWidth = finitePositive(
    finitePositive(input.defaultColumnWidth, input.colWidth ?? DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX) *
      SPREADSHEET_PAGE_LAYOUT_PX_TO_PT,
    SPREADSHEET_PAGE_LAYOUT_DEFAULT_COLUMN_WIDTH_PT,
  );
  const sourceTotalRows = boundedDimension(input.totalRows);
  const sourceTotalCols = boundedDimension(input.totalCols);
  const printArea = settings.printArea
    ? parseA1Range(settings.printArea)
    : null;
  let drawingRight = 0, drawingBottom = 0;
  if (!printArea) {
    for (const bounds of input.drawingBounds ?? []) {
      if (!Array.isArray(bounds) || bounds.length !== 4 || !bounds.every(Number.isFinite)) continue;
      const [left, top, width, height] = bounds;
      const right = left + width, bottom = top + height;
      if (width <= 0 || height <= 0 || !Number.isFinite(right) || !Number.isFinite(bottom) || right <= 0 || bottom <= 0) continue;
      drawingRight = Math.max(drawingRight, right);
      drawingBottom = Math.max(drawingBottom, bottom);
    }
  }
  // A print area prints whole, however few rows and columns the sheet has
  // cells in: a sheet holding only a chart (its data on another sheet) has a
  // one-cell extent, and its print area A1:M24 printed one cell, the chart
  // cut to a corner of its frame. Excel prints the area.
  // (A whole-row or whole-column area, $A:$C, keeps to the sheet's extent.)
  const boundedArea = printArea && printArea.r2 < MAX_GRID_ROWS - 1 && printArea.c2 < MAX_GRID_COLUMNS - 1 ? printArea : null;
  const areaRows = boundedArea ? boundedArea.r2 + 1 : 0;
  const areaCols = boundedArea ? boundedArea.c2 + 1 : 0;
  const totalRows = drawingAxisCount(Math.max(sourceTotalRows, areaRows), drawingBottom, MAX_GRID_ROWS, index =>
    input.hiddenRows?.has(index) ? 0 : finitePositive(input.rowHeights?.[String(index)], baseRowHeight / SPREADSHEET_PAGE_LAYOUT_PX_TO_PT));
  const totalCols = drawingAxisCount(Math.max(sourceTotalCols, areaCols), drawingRight, MAX_GRID_COLUMNS, index =>
    input.hiddenCols?.has(index) ? 0 : finitePositive(input.columnWidths?.[String(index)], baseColWidth / SPREADSHEET_PAGE_LAYOUT_PX_TO_PT));
  // A sorted view's existing source permutation remains intact. The blank
  // rows needed only to reach a drawing follow it in ordinary source order.
  // The order may cover more rows than are printed (the window passes the
  // whole sheet's order beside the printed extent): it is cut to the rows
  // laid out, and only rows past its end are added -- appending past the
  // printed extent instead repeated rows it already holds, and a drawing was
  // printed twice, the second copy a page band lower (Sheets print parity).
  const orderedRows = input.rowOrder?.slice(0, totalRows);
  const rowOrder = orderedRows && totalRows > orderedRows.length
    ? [...orderedRows, ...Array.from({ length: totalRows - orderedRows.length }, (_, index) => orderedRows.length + index)]
    : orderedRows;
  const repeatRows = settings.repeatTitleRows
    ? parseA1Range(settings.repeatTitleRows)
    : null;
  const repeatCols = settings.repeatTitleCols
    ? parseA1Range(settings.repeatTitleCols)
    : null;
  const rowProjection = resolveAxisProjection({
    total: totalRows,
    order: rowOrder,
    hidden: input.hiddenRows,
    areaStart: printArea?.r1 ?? 0,
    areaEnd: printArea?.r2 ?? totalRows - 1,
    ...(repeatRows
      ? { repeatStart: repeatRows.r1, repeatEnd: repeatRows.r2 }
      : {}),
  });
  const colProjection = resolveAxisProjection({
    total: totalCols,
    hidden: input.hiddenCols,
    areaStart: printArea?.c1 ?? 0,
    areaEnd: printArea?.c2 ?? totalCols - 1,
    ...(repeatCols
      ? { repeatStart: repeatCols.c1, repeatEnd: repeatCols.c2 }
      : {}),
  });
  const stock = resolvePaperSizePoints(settings.paperSize);
  const landscape = settings.orientation === "landscape";
  const paperWidth = landscape ? stock.height : stock.width;
  const paperHeight = landscape ? stock.width : stock.height;
  const sourceMargins = resolveMargins(settings.margins);
  // Native Mac worksheet origins use whole point pixels. Keep authored
  // precision separately so merely viewing or grabbing a ruler cannot edit it.
  const margins = input.nativeGridDpi === 72
    ? { left: Math.floor(sourceMargins.left), right: Math.floor(sourceMargins.right),
        top: Math.floor(sourceMargins.top), bottom: Math.floor(sourceMargins.bottom) }
    : sourceMargins;
  const dataAreaWidth = Math.max(1, paperWidth - margins.left - margins.right);
  const dataAreaHeight = Math.max(
    1,
    paperHeight - margins.top - margins.bottom,
  );
  const sourceRowSizes = Array.from({ length: totalRows }, (_, index) =>
    input.hiddenRows?.has(index) ? 0 : finitePositive(input.rowHeights?.[String(index)], baseRowHeight / SPREADSHEET_PAGE_LAYOUT_PX_TO_PT) *
    SPREADSHEET_PAGE_LAYOUT_PX_TO_PT,
  );
  const sourceColSizes = Array.from({ length: totalCols }, (_, index) =>
    input.hiddenCols?.has(index) ? 0 : finitePositive(input.columnWidths?.[String(index)], baseColWidth / SPREADSHEET_PAGE_LAYOUT_PX_TO_PT) *
    SPREADSHEET_PAGE_LAYOUT_PX_TO_PT,
  );
  const projectSizes = (
    projection: PageLayoutAxisProjection,
    source: readonly number[],
    fallback: number,
  ) =>
    Array.from({ length: projection.paginateCount }, (_, index) => {
      const actual = mapLocalAxisIndex(projection, index);
      return actual === null ? fallback : (source[actual] ?? fallback);
    });
  const paginationSettings: PrintSettings = {
    ...settings,
    repeatTitleRows:
      rowProjection.repeatCount > 0
        ? `1:${rowProjection.repeatCount}`
        : undefined,
    repeatTitleCols:
      colProjection.repeatCount > 0
        ? `A:${columnLabel(colProjection.repeatCount - 1)}`
        : undefined,
    manualRowBreaks: settings.manualRowBreaks
      ?.map((value) => mapActualAxisIndex(rowProjection, rowOrder ? rowOrder.indexOf(value) : value))
      .filter((value): value is number => value !== null),
    manualColBreaks: settings.manualColBreaks
      ?.map((value) => mapActualAxisIndex(colProjection, value))
      .filter((value): value is number => value !== null),
  };
  // Printed headings, measured in Excel for Mac (Page Setup > Sheet > Row and
  // column headings, its own PDF): the letter band is one default row tall;
  // the number column holds the widest row number the print shows, never
  // fewer than three digits, plus padding -- Calibri 11 (6 pt digits at
  // 72 dpi) is 23 pt through row 999 and 29 pt at row 1200; Arial 16 (9 pt
  // digits) is 34 pt. One width for every page of the sheet.
  const headingDpi = input.nativeGridDpi === 72 ? 72 : 96;
  const headingDigit = finitePositive(input.nativeDigitWidth, headingDpi === 72 ? 6 : 7);
  const lastPrintedRow = Math.max(printArea?.r2 ?? totalRows - 1, repeatRows?.r2 ?? 0);
  const headingDigits = Math.max(3, String(lastPrintedRow + 1).length);
  const unscaledHeadingHeight = settings.showHeadings === true ? baseRowHeight : 0;
  const unscaledHeadingWidth = settings.showHeadings === true
    ? ((headingDigits * headingDigit + Math.round(headingDigit * 0.8)) * 72) / headingDpi
    : 0;
  const unscaledPaginateInput: PaginateInput = {
    headingHeight: unscaledHeadingHeight,
    headingWidth: unscaledHeadingWidth,
    totalRows: rowProjection.paginateCount,
    totalCols: colProjection.paginateCount,
    dataAreaHeight,
    dataAreaWidth,
    rowHeight: baseRowHeight,
    colWidth: baseColWidth,
    rowSizes: projectSizes(rowProjection, sourceRowSizes, baseRowHeight),
    colSizes: projectSizes(colProjection, sourceColSizes, baseColWidth),
    settings: paginationSettings,
  };
  const contentScale = resolveContentScale(unscaledPaginateInput, input.contentScaleOverride);
  const rowHeight = baseRowHeight * contentScale;
  const colWidth = baseColWidth * contentScale;
  const paginateInput: PaginateInput = {
    ...unscaledPaginateInput,
    rowHeight,
    colWidth,
    rowSizes: unscaledPaginateInput.rowSizes?.map((size) => size * contentScale),
    colSizes: unscaledPaginateInput.colSizes?.map((size) => size * contentScale),
    headingHeight: unscaledHeadingHeight * contentScale,
    headingWidth: unscaledHeadingWidth * contentScale,
  };
  const pagination = paginateSheet(paginateInput);
  const totalPages = Math.max(
    1,
    pagination.rowRanges.length * pagination.colRanges.length,
  );
  return {
    settings,
    paperWidth,
    paperHeight,
    margins,
    ...(input.nativeGridDpi === 72 ? { sourceMargins } : {}),
    headerMargin: finiteNonNegative(settings.headerMargin, 21.6),
    footerMargin: finiteNonNegative(settings.footerMargin, 21.6),
    dataAreaWidth,
    dataAreaHeight,
    rowHeight,
    colWidth,
    rowSizes: sourceRowSizes.map((size) => size * contentScale),
    colSizes: sourceColSizes.map((size) => size * contentScale),
    contentScale,
    headingHeight: unscaledHeadingHeight * contentScale,
    headingWidth: unscaledHeadingWidth * contentScale,
    paginateInput,
    pagination,
    totalPages,
    rowProjection,
    colProjection,
  };
}

/**
 * A multi-area print area: every block is paged as a print area of its own
 * (Excel: "Each print area prints on its own page"), all at one content
 * scale so a Fit To policy holds for the whole sheet, and the pages are
 * numbered straight through in block order.
 */
function resolveMultiAreaPageLayoutGeometry(
  input: Parameters<typeof resolveSpreadsheetPageLayoutGeometry>[0],
  settings: PrintSettings,
  blocks: readonly string[],
): SpreadsheetPageLayoutGeometry {
  const resolveBlock = (block: string, contentScaleOverride?: number) =>
    resolveSpreadsheetPageLayoutGeometry({
      ...input,
      printSettings: { ...settings, printArea: block },
      ...(contentScaleOverride !== undefined ? { contentScaleOverride } : {}),
    });
  let segments = blocks.map((block) => resolveBlock(block));
  const scale = Math.min(...segments.map((segment) => segment.contentScale));
  if (segments.some((segment) => segment.contentScale !== scale)) {
    segments = blocks.map((block) => resolveBlock(block, scale));
  }
  const segmentPageOffsets = [0];
  for (const segment of segments) {
    segmentPageOffsets.push(segmentPageOffsets[segmentPageOffsets.length - 1] + segment.totalPages);
  }
  return {
    ...segments[0],
    settings,
    totalPages: segmentPageOffsets[segmentPageOffsets.length - 1],
    segments,
    segmentPageOffsets,
  };
}

/** The segment holding a composite page index, and the index inside it. */
function resolveSegmentPage(
  geometry: SpreadsheetPageLayoutGeometry,
  index: number,
): { segment: SpreadsheetPageLayoutGeometry; segmentIndex: number; localIndex: number } | null {
  const segments = geometry.segments;
  const offsets = geometry.segmentPageOffsets;
  if (!segments || !offsets) return null;
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    if (index >= offsets[segmentIndex] && index < offsets[segmentIndex + 1]) {
      return { segment: segments[segmentIndex], segmentIndex, localIndex: index - offsets[segmentIndex] };
    }
  }
  return null;
}

function axisPageIndices(
  geometry: SpreadsheetPageLayoutGeometry,
  index: number,
): { rowPageIndex: number; colPageIndex: number } {
  const rowPageCount = geometry.pagination.rowRanges.length;
  const colPageCount = geometry.pagination.colRanges.length;
  const overThenDown = geometry.settings.pageOrder === "overThenDown";
  return {
    rowPageIndex: overThenDown ? Math.floor(index / colPageCount) : index % rowPageCount,
    colPageIndex: overThenDown ? index % colPageCount : Math.floor(index / rowPageCount),
  };
}

/**
 * Page Layout view lays paper tiles on a grid. One print area (or none) is
 * the pager's own row-pages × column-pages grid; a multi-area print area
 * stacks each area's grid below the previous one, so the tiles read in the
 * same order the pages print.
 */
export function resolveSpreadsheetPageLayoutTileGrid(
  geometry: SpreadsheetPageLayoutGeometry,
): { rows: number; cols: number } {
  if (!geometry.segments) {
    return { rows: geometry.pagination.rowRanges.length, cols: geometry.pagination.colRanges.length };
  }
  return geometry.segments.reduce(
    (grid, segment) => ({
      rows: grid.rows + segment.pagination.rowRanges.length,
      cols: Math.max(grid.cols, segment.pagination.colRanges.length),
    }),
    { rows: 0, cols: 0 },
  );
}

/** The tile a page occupies, or null for an index outside the geometry. */
export function resolveSpreadsheetPageLayoutTilePosition(
  geometry: SpreadsheetPageLayoutGeometry,
  index: number,
): { row: number; col: number } | null {
  if (!Number.isSafeInteger(index) || index < 0 || index >= geometry.totalPages) return null;
  if (!geometry.segments) {
    const { rowPageIndex, colPageIndex } = axisPageIndices(geometry, index);
    return { row: rowPageIndex, col: colPageIndex };
  }
  const located = resolveSegmentPage(geometry, index);
  if (!located) return null;
  let rowOffset = 0;
  for (let segmentIndex = 0; segmentIndex < located.segmentIndex; segmentIndex += 1) {
    rowOffset += geometry.segments[segmentIndex].pagination.rowRanges.length;
  }
  const { rowPageIndex, colPageIndex } = axisPageIndices(located.segment, located.localIndex);
  return { row: rowOffset + rowPageIndex, col: colPageIndex };
}

/** The page index at a tile, or null when the tile holds no page. */
export function resolveSpreadsheetPageLayoutPageAtTile(
  geometry: SpreadsheetPageLayoutGeometry,
  row: number,
  col: number,
): number | null {
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col) || row < 0 || col < 0) return null;
  const indexIn = (target: SpreadsheetPageLayoutGeometry, localRow: number) => {
    const rowCount = target.pagination.rowRanges.length;
    const colCount = target.pagination.colRanges.length;
    if (localRow >= rowCount || col >= colCount) return null;
    return target.settings.pageOrder === "overThenDown" ? localRow * colCount + col : col * rowCount + localRow;
  };
  if (!geometry.segments || !geometry.segmentPageOffsets) return indexIn(geometry, row);
  let rowOffset = 0;
  for (let segmentIndex = 0; segmentIndex < geometry.segments.length; segmentIndex += 1) {
    const segment = geometry.segments[segmentIndex];
    const rowCount = segment.pagination.rowRanges.length;
    if (row < rowOffset + rowCount) {
      const local = indexIn(segment, row - rowOffset);
      return local === null ? null : geometry.segmentPageOffsets[segmentIndex] + local;
    }
    rowOffset += rowCount;
  }
  return null;
}

function countIndicesForPageRange(
  range: [number, number],
  projection: PageLayoutAxisProjection,
): number {
  const dataStart = Math.max(range[0], projection.repeatCount);
  const dataEnd = Math.min(
    range[1],
    projection.repeatCount + projection.dataCount - 1,
  );
  return projection.repeatCount + Math.max(0, dataEnd - dataStart + 1);
}

function indicesForPageRange(
  range: [number, number],
  projection: PageLayoutAxisProjection,
  limit: number,
): number[] {
  const result: number[] = [];
  for (
    let index = 0;
    index < projection.repeatCount && result.length < limit;
    index += 1
  ) {
    const mapped = mapLocalAxisIndex(projection, index);
    if (mapped !== null) result.push(mapped);
  }
  const dataStart = Math.max(range[0], projection.repeatCount);
  const dataEnd = Math.min(
    range[1],
    projection.repeatCount + projection.dataCount - 1,
  );
  for (
    let index = dataStart;
    index <= dataEnd && result.length < limit;
    index += 1
  ) {
    const mapped = mapLocalAxisIndex(projection, index);
    if (mapped !== null) result.push(mapped);
  }
  return result;
}

function boundedAxisLimits(
  rowCount: number,
  colCount: number,
  cellBudget: number,
): { rows: number; cols: number } {
  const budget = Math.max(1, Math.floor(cellBudget));
  let rows = Math.min(
    rowCount,
    Math.max(1, Math.floor(Math.sqrt((budget * rowCount) / colCount))),
  );
  const cols = Math.min(colCount, Math.max(1, Math.floor(budget / rows)));
  rows = Math.min(rows, Math.max(1, Math.floor(budget / cols)));
  return { rows, cols };
}

/** Resolve a positional page without materialising the full row×column product. */
export function resolveSpreadsheetPageLayoutPage(
  geometry: SpreadsheetPageLayoutGeometry,
  index: number,
  cellBudget = MAX_RENDERED_PAGE_LAYOUT_CELLS_PER_PAGE,
): SpreadsheetPageLayoutPage | null {
  if (
    !Number.isSafeInteger(index) ||
    index < 0 ||
    index >= geometry.totalPages
  ) {
    return null;
  }
  if (geometry.segments) {
    const located = resolveSegmentPage(geometry, index);
    const page = located
      ? resolveSpreadsheetPageLayoutPage(located.segment, located.localIndex, cellBudget)
      : null;
    return page
      ? { ...page, index, pageNumber: (geometry.settings.firstPageNumber ?? 1) + index }
      : null;
  }
  const { rowPageIndex, colPageIndex } = axisPageIndices(geometry, index);
  const rowRange = geometry.pagination.rowRanges[rowPageIndex];
  const colRange = geometry.pagination.colRanges[colPageIndex];
  if (!rowRange || !colRange) return null;
  const totalRows = countIndicesForPageRange(
    rowRange,
    geometry.rowProjection,
  );
  const totalCols = countIndicesForPageRange(
    colRange,
    geometry.colProjection,
  );
  const limits = boundedAxisLimits(totalRows, totalCols, cellBudget);
  const rows = indicesForPageRange(
    rowRange,
    geometry.rowProjection,
    limits.rows,
  );
  const cols = indicesForPageRange(
    colRange,
    geometry.colProjection,
    limits.cols,
  );
  const mappedRowStart = mapLocalAxisIndex(
    geometry.rowProjection,
    rowRange[0],
  );
  const mappedRowEnd = mapLocalAxisIndex(
    geometry.rowProjection,
    rowRange[1],
  );
  const mappedColStart = mapLocalAxisIndex(
    geometry.colProjection,
    colRange[0],
  );
  const mappedColEnd = mapLocalAxisIndex(
    geometry.colProjection,
    colRange[1],
  );
  const totalCellCount = totalRows * totalCols;
  const renderedCellCount = rows.length * cols.length;
  return {
    index,
    pageNumber: (geometry.settings.firstPageNumber ?? 1) + index,
    rowRange: [
      mappedRowStart ?? geometry.rowProjection.repeatStart,
      mappedRowEnd ??
        geometry.rowProjection.dataStart + geometry.rowProjection.dataCount - 1,
    ],
    colRange: [
      mappedColStart ?? geometry.colProjection.repeatStart,
      mappedColEnd ??
        geometry.colProjection.dataStart + geometry.colProjection.dataCount - 1,
    ],
    rows,
    cols,
    totalCellCount,
    renderedCellCount,
    cellsTruncated: renderedCellCount < totalCellCount,
  };
}

export function resolveSpreadsheetPageLayoutRenderPlan(
  geometry: SpreadsheetPageLayoutGeometry,
  limits: {
    maxPages?: number;
    maxCellsPerPage?: number;
    maxCellsTotal?: number;
  } = {},
): SpreadsheetPageLayoutRenderPlan {
  const maxPages = Math.max(
    1,
    Math.floor(limits.maxPages ?? MAX_RENDERED_PAGE_LAYOUT_PAGES),
  );
  const maxCellsPerPage = Math.max(
    1,
    Math.floor(
      limits.maxCellsPerPage ?? MAX_RENDERED_PAGE_LAYOUT_CELLS_PER_PAGE,
    ),
  );
  const maxCellsTotal = Math.max(
    1,
    Math.floor(limits.maxCellsTotal ?? MAX_RENDERED_PAGE_LAYOUT_CELLS_TOTAL),
  );
  const pages: SpreadsheetPageLayoutPage[] = [];
  let renderedCellCount = 0;
  let cellLimitReached = false;
  const candidatePages = Math.min(geometry.totalPages, maxPages);
  // Content pages take the budget first; browsing blanks follow in index order.
  const content = geometry.contentPages;
  const isContentPage = (index: number) => {
    if (!content) return true;
    const rowPageCount = geometry.pagination.rowRanges.length;
    const colPageCount = geometry.pagination.colRanges.length;
    const overThenDown = geometry.settings.pageOrder === "overThenDown";
    const rowPageIndex = overThenDown ? Math.floor(index / colPageCount) : index % rowPageCount;
    const colPageIndex = overThenDown ? index % colPageCount : Math.floor(index / rowPageCount);
    return rowPageIndex < content.rowPages && colPageIndex < content.colPages;
  };
  const order = Array.from({ length: candidatePages }, (_, index) => index);
  if (content) order.sort((a, b) => Number(isContentPage(b)) - Number(isContentPage(a)) || a - b);
  for (const index of order) {
    const remaining = maxCellsTotal - renderedCellCount;
    if (remaining <= 0) {
      cellLimitReached = true;
      break;
    }
    const page = resolveSpreadsheetPageLayoutPage(
      geometry,
      index,
      Math.min(maxCellsPerPage, remaining),
    );
    if (!page) break;
    pages.push(page);
    renderedCellCount += page.renderedCellCount;
    if (page.cellsTruncated) cellLimitReached = true;
  }
  if (pages.length < candidatePages) cellLimitReached = true;
  return {
    pages,
    renderedCellCount,
    pageLimitReached: geometry.totalPages > candidatePages,
    cellLimitReached,
  };
}

/** Cells one batch of a print mounts at once, and one page at most. */
export const PRINT_BATCH_CELLS = 20_000;
export const PRINT_PAGE_CELLS = 50_000;

/**
 * The pages a print mounts from `start` on (Page Layout's printMode): every
 * page, a batch of them at a time within PRINT_BATCH_CELLS, so a sheet of
 * any length prints whole; `nextStart` is where the next batch begins (null
 * after the last). A page past PRINT_PAGE_CELLS is truncated
 * (`cellLimitReached`: the print is refused rather than short).
 */
export function resolveSpreadsheetPrintBatchPlan(
  geometry: SpreadsheetPageLayoutGeometry,
  start: number,
): SpreadsheetPageLayoutRenderPlan & { nextStart: number | null } {
  const pages: SpreadsheetPageLayoutPage[] = [];
  let renderedCellCount = 0;
  let cellLimitReached = false;
  let index = Math.max(0, Math.floor(start));
  for (; index < geometry.totalPages; index += 1) {
    const page = resolveSpreadsheetPageLayoutPage(geometry, index, PRINT_PAGE_CELLS);
    if (!page) break;
    if (pages.length && renderedCellCount + page.renderedCellCount > PRINT_BATCH_CELLS) break;
    pages.push(page);
    renderedCellCount += page.renderedCellCount;
    if (page.cellsTruncated) cellLimitReached = true;
  }
  return {
    pages,
    renderedCellCount,
    pageLimitReached: false,
    cellLimitReached,
    nextStart: index < geometry.totalPages ? index : null,
  };
}

/** Map shared local pagination lines back to live worksheet coordinates. */
export function resolveSpreadsheetPageLayoutBreakLines(
  geometry: SpreadsheetPageLayoutGeometry,
): PageBreakLines {
  if (geometry.segments) {
    const union: PageBreakLines = { rowBreaks: [], colBreaks: [], manualRows: new Set(), manualCols: new Set() };
    for (const segment of geometry.segments) {
      const lines = resolveSpreadsheetPageLayoutBreakLines(segment);
      union.rowBreaks.push(...lines.rowBreaks);
      union.colBreaks.push(...lines.colBreaks);
      lines.manualRows.forEach((value) => union.manualRows.add(value));
      lines.manualCols.forEach((value) => union.manualCols.add(value));
    }
    union.rowBreaks = Array.from(new Set(union.rowBreaks)).sort((a, b) => a - b);
    union.colBreaks = Array.from(new Set(union.colBreaks)).sort((a, b) => a - b);
    return union;
  }
  const local = computePageBreakLines(geometry.paginateInput);
  const mapValues = (
    values: Iterable<number>,
    projection: PageLayoutAxisProjection,
  ) =>
    Array.from(values)
      .map((value) => mapLocalAxisIndex(projection, value))
      .filter((value): value is number => value !== null)
      .sort((a, b) => a - b);
  const rowBreaks = Array.from(
    new Set(mapValues(local.rowBreaks, geometry.rowProjection)),
  );
  const colBreaks = Array.from(
    new Set(mapValues(local.colBreaks, geometry.colProjection)),
  );
  return {
    rowBreaks,
    colBreaks,
    manualRows: new Set(mapValues(local.manualRows, geometry.rowProjection)),
    manualCols: new Set(mapValues(local.manualCols, geometry.colProjection)),
  };
}

/**
 * Page Break Preview's printed region and its pages, in live worksheet lines.
 * Each axis is an ascending list of page boundaries: the first entry is the
 * printed region's leading line, every later entry but the last is the first
 * line of a new page, and the last is one past the region's trailing line.
 * Rows are visual (Sheet View) positions, the space the grid is drawn in.
 */
export interface SpreadsheetPageBreakPreviewPages {
  rows: number[];
  cols: number[];
  /** Excel numbers pages down, then over, unless Page Setup says otherwise. */
  overThenDown: boolean;
  /** Page number of this region's first page; a later print area continues
   * the count where the previous one stopped. */
  firstPage: number;
}

function pageBreakPreviewAxis(
  ranges: readonly (readonly [number, number])[],
  projection: PageLayoutAxisProjection,
  hullStart: number,
  hullEnd: number,
): number[] {
  const boundaries = [hullStart];
  for (let index = 1; index < ranges.length; index += 1) {
    const line = mapLocalAxisIndex(projection, ranges[index][0]);
    if (line !== null && line > boundaries[boundaries.length - 1] && line < hullEnd) {
      boundaries.push(line);
    }
  }
  boundaries.push(Math.max(hullEnd, boundaries[boundaries.length - 1]));
  return boundaries;
}

/**
 * Without a print area the whole sheet prints, exactly as the shared pager
 * paginates it. With one, the region is the print area, clamped to the grid;
 * under a sorted Sheet View it is the span of the area's visual rows.
 */
export function resolveSpreadsheetPageBreakPreviewPages(
  geometry: SpreadsheetPageLayoutGeometry,
  grid: { rowCount: number; columnCount: number; rowOrder?: readonly number[] },
  firstPage = 1,
): SpreadsheetPageBreakPreviewPages {
  const rowCount = Math.max(0, Math.floor(grid.rowCount));
  const columnCount = Math.max(0, Math.floor(grid.columnCount));
  const area = geometry.settings.printArea
    ? parseA1Range(geometry.settings.printArea)
    : null;
  let rowStart = area ? Math.min(area.r1, rowCount) : 0;
  let rowEnd = area ? Math.min(area.r2 + 1, rowCount) : rowCount;
  if (area && grid.rowOrder?.some((source, visual) => source !== visual)) {
    let first = rowCount, last = -1;
    grid.rowOrder.forEach((source, visual) => {
      if (visual >= rowCount || source < area.r1 || source > area.r2) return;
      first = Math.min(first, visual);
      last = Math.max(last, visual);
    });
    rowStart = last < 0 ? rowCount : first;
    rowEnd = last < 0 ? rowCount : last + 1;
  }
  const colStart = area ? Math.min(area.c1, columnCount) : 0;
  const colEnd = area ? Math.min(area.c2 + 1, columnCount) : columnCount;
  return {
    rows: pageBreakPreviewAxis(geometry.pagination.rowRanges, geometry.rowProjection, rowStart, Math.max(rowStart, rowEnd)),
    cols: pageBreakPreviewAxis(geometry.pagination.colRanges, geometry.colProjection, colStart, Math.max(colStart, colEnd)),
    overThenDown: geometry.settings.pageOrder === "overThenDown",
    firstPage,
  };
}

/**
 * Every printed region of the sheet: one for the whole sheet or its single
 * print area, one per block of a multi-area print area. Excel's Page Break
 * Preview shows each area as its own white region with its pages numbered
 * on from the previous area; everything outside every region is shaded.
 */
export function resolveSpreadsheetPageBreakPreviewRegions(
  geometry: SpreadsheetPageLayoutGeometry,
  grid: { rowCount: number; columnCount: number; rowOrder?: readonly number[] },
): SpreadsheetPageBreakPreviewPages[] {
  if (!geometry.segments) return [resolveSpreadsheetPageBreakPreviewPages(geometry, grid)];
  let firstPage = 1;
  return geometry.segments.map((segment) => {
    const region = resolveSpreadsheetPageBreakPreviewPages(segment, grid, firstPage);
    firstPage += segment.totalPages;
    return region;
  });
}

/** Offsets of many lines on one axis in a single walk, instead of one
 * prefix sum per line. `lines` need not be sorted; the result maps each line
 * (clamped to [0, count]) to the summed size of the lines before it. */
export function pageBreakPreviewLineOffsets(
  lines: Iterable<number>,
  count: number,
  size: (line: number) => number,
): Map<number, number> {
  const safeCount = Math.max(0, Math.floor(count));
  const wanted = Array.from(new Set(lines))
    .filter((line) => Number.isFinite(line))
    .sort((a, b) => a - b);
  const offsets = new Map<number, number>();
  let line = 0, offset = 0;
  for (const target of wanted) {
    const bound = Math.min(Math.max(0, Math.floor(target)), safeCount);
    for (; line < bound; line += 1) offset += size(line);
    offsets.set(target, offset);
  }
  return offsets;
}

/**
 * Which header/footer pair a printed page wears: the first-page variant wins,
 * then the even-page variant, then odd/default. Both follow the page's
 * PHYSICAL position in print order, never its printed number: Excel's own PDF
 * of a sheet that starts at page 2 prints the odd header on its first page.
 * `pageNumber` is the printed number, so the position is its distance from
 * the sheet's first page number.
 */
export function resolvePageHeaderFooterVariant(
  settings: PrintSettings,
  page: Pick<SpreadsheetPageLayoutPage, "pageNumber">,
): "first" | "even" | "default" {
  const position = page.pageNumber - (settings.firstPageNumber ?? 1);
  if (position === 0 && settings.differentFirstPage === true) return "first";
  if (settings.differentOddEven === true && position % 2 === 1) return "even";
  return "default";
}

/** The header/footer pair {@link resolvePageHeaderFooterVariant} selects. */
export function resolvePageHeaderFooter(
  settings: PrintSettings,
  page: Pick<SpreadsheetPageLayoutPage, "pageNumber">,
): { header?: PrintHeaderFooter; footer?: PrintHeaderFooter } {
  const variant = resolvePageHeaderFooterVariant(settings, page);
  if (variant === "first") {
    return {
      header: settings.firstHeader,
      footer: settings.firstFooter,
    };
  }
  if (variant === "even") {
    return {
      header: settings.evenHeader,
      footer: settings.evenFooter,
    };
  }
  return { header: settings.header, footer: settings.footer };
}

/**
 * Substitute fields without losing the imported run boundaries. A malformed
 * hand-edited run overlay fails safely to the authoritative plain segment.
 */
export function resolveHeaderFooterSegment(
  text: string | undefined,
  runs: PrintHeaderFooterRun[] | undefined,
  context: HeaderFooterContext,
  picture?: PrintHeaderFooterPicture,
): ResolvedHeaderFooterSegment | null {
  if (typeof text !== "string" || text === "") return null;
  const options = picture ? { pictureMark: HEADER_FOOTER_PICTURE_MARK } : {};
  // Only the first &[Picture] code draws the section's one picture.
  const firstMarkOnly = () => {
    let marked = false;
    return (value: string) =>
      value.replace(/\uFFFC/g, () => {
        if (marked) return "";
        marked = true;
        return HEADER_FOOTER_PICTURE_MARK;
      });
  };
  const resolvedText = firstMarkOnly()(substituteHeaderFooterTokens(text, context, options));
  const withPicture = (resolved: ResolvedHeaderFooterSegment) =>
    picture && resolvedText.includes(HEADER_FOOTER_PICTURE_MARK)
      ? { ...resolved, picture }
      : resolved;
  if (!runs?.length || runs.map((run) => run.text).join("") !== text) {
    return withPicture({ text: resolvedText });
  }
  const markRuns = firstMarkOnly();
  return withPicture({
    text: resolvedText,
    runs: runs.map((run) => ({
      ...run,
      text: markRuns(substituteHeaderFooterTokens(run.text, context, options)),
    })),
  });
}

/** Only explicit percent scaling can map directly to header/footer font size. */
export function resolveHeaderFooterDocumentScale(
  settings: PrintSettings,
  resolvedContentScale?: number,
): number {
  if (settings.headerFooterScaleWithDoc === false) return 1;
  if (
    typeof resolvedContentScale === "number" &&
    Number.isFinite(resolvedContentScale) &&
    resolvedContentScale > 0
  ) {
    return resolvedContentScale;
  }
  if (settings.scaling?.kind !== "percent") return 1;
  return Math.max(0.1, Math.min(4, settings.scaling.percent / 100));
}
