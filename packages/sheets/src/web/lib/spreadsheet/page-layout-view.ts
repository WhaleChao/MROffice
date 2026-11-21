/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pxToPt } from "@mrnote/core/web/lib/office-units.ts";
import { aspectLockedResizeSize } from "./drawing-aspect";
import { OFFICE_PRINT_IDENTITY, type OfficePrintScale } from "./office-print-metrics";
import { styledColumnEdge } from "@mrnote/core/web/lib/spreadsheet/line-styles.ts";
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
  printPageSparePoints,
  printedColumnPoints,
  printedRowPoints,
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

/** The columns wider than the page's printable width, as [start, end] sheet
 * spans in the printed sizes given, and that width. */
export interface PageLayoutWideColumns {
  spans: readonly (readonly [start: number, end: number])[];
  limit: number;
}

/** Windows Excel prints a column wider than the printable width cut at that
 * width, and lays out what lies right of it -- columns, pictures, charts --
 * as if the column were exactly that wide (lookwin probe p48-lw: a picture
 * 600 pt into a 150-character B prints at C's left margin; R59-X13: a 758 pt
 * picture from A reaches past B onto C's page). Null when no column is. */
export function pageLayoutWideColumns(sizes: readonly number[], limit: number): PageLayoutWideColumns | null {
  if (!(limit > 0) || !Number.isFinite(limit)) return null;
  const spans: Array<readonly [number, number]> = [];
  let position = 0;
  for (const size of sizes) {
    if (size > limit + 1e-6) spans.push([position, position + size]);
    position += size;
  }
  return spans.length ? { spans, limit } : null;
}

/** A sheet x (printed points) as the printed page places it: past a wide
 * column's printable width it stands at that column's cut end. */
export function pageLayoutClampColumnPoint(value: number, wide: PageLayoutWideColumns | null | undefined): number {
  if (!wide) return value;
  let cut = 0;
  for (const [start, end] of wide.spans) {
    if (value <= start + wide.limit) break;
    cut += Math.min(value, end) - (start + wide.limit);
  }
  return value - cut;
}

/** Inverse of pageLayoutClampColumnPoint: a printed x back to the sheet's
 * (a point at a wide column's cut end is that column's end). */
export function pageLayoutUnclampColumnPoint(value: number, wide: PageLayoutWideColumns | null | undefined): number {
  if (!wide) return value;
  let cut = 0;
  for (const [start, end] of wide.spans) {
    if (value + cut < start + wide.limit) break;
    cut += end - start - wide.limit;
  }
  return value + cut;
}

/** Without an explicit print area, the final data line is not a drawing
 * clipping boundary. Use the remaining printable paper for its last band;
 * repeated-title bands and intermediate page boundaries stay unchanged.
 * Wide columns (pageLayoutWideColumns) clip at the printable width, and the
 * bands after them start where the printed page puts them. */
export function pageLayoutDrawingBands(sizes: readonly number[], indices: readonly number[], availablePoints: number, finalUnboundedPage: boolean, wide?: PageLayoutWideColumns | null) {
  const bands = pageLayoutAxisBands(sizes, indices);
  if (wide) {
    let offset = 0;
    for (const band of bands) {
      band.start = pageLayoutClampColumnPoint(band.start, wide);
      band.end = pageLayoutClampColumnPoint(band.end, wide);
      band.offset = offset;
      offset += band.end - band.start;
    }
  }
  const last = bands.at(-1);
  if (last && finalUnboundedPage && Number.isFinite(availablePoints)) {
    last.end = Math.max(last.end, last.start + Math.max(0, availablePoints - last.offset));
  }
  return bands;
}

/**
 * The cells that keep a trailing page for the blanks whose border or fill
 * record widens Excel's range (the import's printedFormatExtent, reading
 * 956): the two it ends at where the import names them (`ends`, reading
 * 968: R454-X14's Assumptions ends at L38 and M5, and Windows Excel leaves
 * out the M38 page nothing is on), else the range's far corner (lookwin
 * p35-lw: Z5 alone, its blank band printed).
 */
export function printedFormatExtentKeys(extent: unknown): string[] {
  const { rows, cols, ends } = (extent ?? {}) as { rows?: unknown; cols?: unknown; ends?: unknown };
  if (typeof rows !== "number" || typeof cols !== "number" || !Number.isInteger(rows) || !Number.isInteger(cols)
    || rows <= 0 || cols <= 0) return [];
  const named = Array.isArray(ends) && ends.length > 0 && ends.every((end) => Array.isArray(end) && end.length === 2
    && end.every((index) => typeof index === "number" && Number.isInteger(index) && index >= 0))
    ? (ends as [number, number][]).map(([row, col]) => `${row}:${col}`) : null;
  return named ?? [`${rows - 1}:${cols - 1}`];
}

/**
 * Whether a cell of its own (a line format painted onto the grid's blanks
 * set aside) widens the range Excel prints without a print area: a value, or
 * a blank whose format shows on paper -- a border, a pattern, a gradient or
 * a fill, white too (lookwin p46 whitefar: an own white blank far past the
 * data widens Excel's print; the server's writer counts it, sheet_extent). A
 * blank carrying only a font, alignment or number format does not (MX060).
 */
export function printedCellWidensRange(cell: { value?: unknown; formula?: unknown; style?: Record<string, unknown> }): boolean {
  const style = cell.style ?? {};
  const fill = typeof style.backgroundColor === "string" && style.backgroundColor !== "" && !/^transparent$/i.test(style.backgroundColor);
  const border = Boolean(style.border) || Boolean(style.borderEdges) || Boolean(style.borderTop || style.borderBottom || style.borderLeft || style.borderRight);
  const valued = (typeof cell.value === "string" && cell.value !== "") || (typeof cell.formula === "string" && cell.formula !== "");
  return valued || fill || border || Boolean(style.pattern) || Boolean(style.gradientFill);
}

const VALUED_CELLS_EXTENT = new WeakMap<object, { rows: number; cols: number }>();

/**
 * The range a sheet's real cells occupy, for its print: the import's
 * occupiedExtent (`occupied`) widened to the cells holding a value or a
 * formula now; null when the import read none. A tiled workbook's page body
 * is projected from its stub, which leaves every valued cell to the tiles
 * (they own them: a tiled save, sort and recalc write them there), so its
 * extent stopped at the stub's last valued row (R13-X27: 50 rows for 51,
 * R379-X18: 5 for 140); the cells the window holds carry the tiles' too. A
 * whole read's extent already holds every valued cell, and a value entered
 * past it is in Excel's used range as well. The API's writer reads the same
 * (`_occupied_extent_with_values`).
 */
export function occupiedExtentWithValues(
  cells: Readonly<Record<string, unknown>>,
  occupied: unknown,
): { rows: number; cols: number } | null {
  const extent = occupied && typeof occupied === "object" ? occupied as { rows?: unknown; cols?: unknown } : null;
  const importedRows = extent?.rows;
  const importedCols = extent?.cols;
  if (typeof importedRows !== "number" || typeof importedCols !== "number") return null;
  let valued = VALUED_CELLS_EXTENT.get(cells);
  if (!valued) {
    let rows = 0;
    let cols = 0;
    for (const key in cells) {
      const cell = cells[key] as { value?: unknown; formula?: unknown } | undefined;
      if (!cell || typeof cell !== "object") continue;
      if ((cell.value === undefined || cell.value === null || cell.value === "")
        && !(typeof cell.formula === "string" && cell.formula !== "")) continue;
      const separator = key.indexOf(":");
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      if (separator <= 0 || !Number.isInteger(row) || !Number.isInteger(col)) continue;
      if (row + 1 > rows) rows = row + 1;
      if (col + 1 > cols) cols = col + 1;
    }
    valued = { rows, cols };
    VALUED_CELLS_EXTENT.set(cells, valued);
  }
  return { rows: Math.max(importedRows, valued.rows), cols: Math.max(importedCols, valued.cols) };
}

/**
 * The blanks that widen the printed range though they only show their row's
 * or column's format: those painted with it (a fill or a border,
 * printedCellWidensRange) inside the range the file's own cells occupy
 * (`occupied`, the import's occupiedExtent). Excel's used range is that
 * range, and it prints the line's paint across it; the same format past it
 * is the line's alone and prints nothing (RO046, R105-X05). The API's writer
 * counts the same (sheet_extent). R453-X52: every column filled white and
 * K5 a cell of its own carrying a font; Excel prints A:G, then the white
 * H:K on a second page. `own` is `cells` less the blanks a line format
 * fully describes (compactLineStyleCells).
 */
export function printedOccupiedLineBlanks(
  cells: Readonly<Record<string, unknown>>,
  own: Readonly<Record<string, unknown>>,
  occupied: unknown,
): string[] {
  // Down to the last valued row (a tiled body's valued cells stay the
  // tiles', occupiedExtentWithValues), across the import's own columns only:
  // past them a column format's painted blanks kept a trailing page Windows
  // Excel leaves out (R379-X18's survey, values in O:Q beside an extent of
  // N: 24 pages for 23).
  const extent = occupied && typeof occupied === "object" ? occupied as { rows?: unknown; cols?: unknown } : null;
  const rows = occupiedExtentWithValues(cells, occupied)?.rows ?? 0;
  const cols = typeof extent?.cols === "number" ? extent.cols : 0;
  if (!(rows > 0) || !(cols > 0) || cells === own) return [];
  const keys: string[] = [];
  for (const key in cells) {
    if (key in own) continue;
    const separator = key.indexOf(":");
    if (separator <= 0 || !(Number(key.slice(0, separator)) < rows) || !(Number(key.slice(separator + 1)) < cols)) continue;
    const cell = cells[key];
    if (cell && typeof cell === "object" && printedCellWidensRange(cell as { value?: unknown; formula?: unknown; style?: Record<string, unknown> })) keys.push(key);
  }
  return keys;
}












export function printedTailUsedColumnEdge(sheet: object): number {
  const { cells, occupiedExtent } = sheet as { cells?: unknown; occupiedExtent?: unknown };
  const cols = cells && typeof cells === "object"
    ? occupiedExtentWithValues(cells as Record<string, unknown>, occupiedExtent)?.cols
    : occupiedExtent && typeof occupiedExtent === "object" ? (occupiedExtent as { cols?: unknown }).cols : undefined;
  const occupiedCols = typeof cols === "number" && Number.isInteger(cols) && cols > 0 && cols <= MAX_GRID_COLUMNS ? cols : 0;
  return Math.max(styledColumnEdge(sheet), occupiedCols);
}

/**
 * Whether a merged block whose value lies in its anchor covers the cell: a
 * carried tail counts its columns as holding a value in the rows it spans
 * (lookwin p222-lw n5: B10:C12 holding "merged", the note's tail prints on
 * C's page; the server's column_holds).
 */
export function printedMergeHoldsAt(
  spans: readonly { row: number; col: number; rowSpan: number; colSpan: number }[],
  cells: Readonly<Record<string, unknown>>,
  row: number,
  col: number,
): boolean {
  for (const span of spans) {
    if (row < span.row || row >= span.row + span.rowSpan || col < span.col || col >= span.col + span.colSpan) continue;
    const anchor = cells[`${span.row}:${span.col}`] as { value?: unknown } | undefined;
    if (typeof anchor?.value === "string" && anchor.value !== "") return true;
  }
  return false;
}

/**
 * Whether a row's `spans` (the file's ``spans``) keep a page that holds no
 * cell of its own: only where the page paints a row's or column's format (a
 * fill or border on its blanks, `paintsLineFormat`) -- R51-X06's bordered
 * rows 424-432 state 1:35 and keep the last band's page (lookwin p67-lw). A
 * page that paints nothing is blank whatever its rows state: R305-X04's rows
 * span 1:13 to M, whose cells carry a font only, and Excel prints none of
 * the five pages below the first in L:M (7 pages; 12 here). The API's
 * page_shows_line_formats_only is the same rule.
 */
export function pageLayoutSpansKeepPage(
  spans: unknown,
  rows: readonly [number, number],
  cols: readonly [number, number],
  paintsLineFormat: () => boolean,
): boolean {
  if (!Array.isArray(spans)) return false;
  const reaches = spans.some((run) => Array.isArray(run) && run.length === 4
    && run[0] <= rows[1] && run[1] >= rows[0] && run[2] <= cols[1] && run[3] >= cols[0]);
  return reaches && paintsLineFormat();
}

/**
 * Whether Windows Excel prints the tail of a text run carried onto a later
 * page: the run reaches a column holding a value in the page's rows (lookwin
 * p69-lw: a note from D ending in G, G5 holding a value, prints; p68-lw: the
 * band's only value in J, the note ending in G, prints nothing; R327-X12's
 * third page J13 / K18), or -- running out past the used range -- the page's
 * last column (p60-lw: a Title note three times as long prints over every
 * band it crosses; p64-lw t119_21: a tail ending inside the band a longer
 * note set does not). The server's writer follows the same rule
 * (draw_spill_from_left). Sizes are printed points, `textEnd` measured from
 * the source cell's left edge to the end of its text; `runEndCol` the last
 * free column of the run (a value stops it), `usedColumns` the used range's
 * column count before spills widen it, a column with a style of its own
 * counted in (lookwin p86b-lw: a tail ending in a styled column prints
 * nothing on its page; R162-X08 Metadata).
 */
export function pageLayoutCarriedRunPrints(input: {
  pageCols: readonly number[];
  colSizes: readonly number[];
  colWidth: number;
  source: number;
  runEndCol: number;
  textEnd: number;
  padding: number;
  usedColumns?: number;
  columnHolds: (col: number) => boolean;
  /** "print": Windows Excel's File > Print, which prints every tail that
   * reaches the page (lookwin p64-lw-winprint, CRITERIA XC-155); else its
   * Save as PDF's rule below. */
  route?: "export" | "print";
}): boolean {
  const { pageCols, source, runEndCol, textEnd, padding } = input;
  const first = pageCols[0];
  const last = pageCols[pageCols.length - 1];
  if (first === undefined || last === undefined || !Number.isFinite(textEnd) || first <= source) return false;
  const size = (col: number) => {
    const value = input.colSizes[col];
    return typeof value === "number" && Number.isFinite(value) ? value : input.colWidth;
  };
  // A column's left edge from the source cell's, points.
  const lefts = new Map<number, number>();
  let position = 0;
  // As far as the text and the page's columns reach: the used range may run
  // to the sheet's end (a styled column to XFD), past which nothing matters.
  for (let col = source; col <= Math.max(last, input.usedColumns ?? 0); col += 1) {
    lefts.set(col, position);
    if (col > last && position > textEnd + padding) break;
    position += size(col);
  }
  const leftOf = (col: number) => lefts.get(col) ?? position;
  // Half a pixel: a run ending right at an edge does not reach past it.
  const slack = 0.375;
  if (textEnd <= leftOf(first) + slack) return false;
  // File > Print: p64-lw t119_44_91's third page prints END2, END3 and END4;
  // p81-lw n127far and p86b-lw socolB (a styled B) print theirs.
  if (input.route === "print") return runEndCol >= first;
  for (const col of pageCols) {
    if (col > runEndCol || leftOf(col) >= textEnd - slack) break;
    if (input.columnHolds(col)) return true;
  }
  // Running on past the page, it prints here too when it reaches a column
  // holding a value further on before a value in its own row stops it
  // (lookwin p222-lw n4: C5's "x" on the third page, the note prints on the
  // second as well). The server's writer: reaches_later_value.
  if (runEndCol >= last && textEnd + padding > leftOf(last) + slack) {
    for (let col = last + 1; col <= runEndCol && lefts.has(col) && leftOf(col) < textEnd - slack; col += 1) {
      if (input.columnHolds(col)) return true;
    }
  }
  // With the cell's right padding (p63-lw's filled notes end 3 pt short of
  // the unfilled one that set the band, and print).
  const reachesLast = runEndCol >= last && textEnd + padding > leftOf(last) + slack;
  const pastUsed = input.usedColumns === undefined || input.usedColumns <= source ||
    textEnd + padding > leftOf(input.usedColumns);
  return reachesLast && pastUsed;
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
  /** The printer's scale of the screen grid (officePrintScale): rowSizes,
   * colSizes and the headings already carry it; a drawing's frame is mapped
   * by its anchor (officePrintedDrawingFrame). Absent: the identity. */
  printScale?: OfficePrintScale;
  /** The rows' and columns' sizes at the content scale before the printer's
   * whole dots (rowSizes / colSizes round them) and quarter pixels: what a
   * drawing's frame is measured in, mapped onto the printed lines by
   * pageLayoutPrintedAxisPoint. */
  sourceRowPoints?: readonly number[];
  sourceColPoints?: readonly number[];
}

const axisPrefixes = new WeakMap<readonly number[], Float64Array>();
function axisPrefix(sizes: readonly number[]): Float64Array {
  let prefix = axisPrefixes.get(sizes);
  if (!prefix) {
    prefix = new Float64Array(sizes.length + 1);
    for (let index = 0; index < sizes.length; index += 1) prefix[index + 1] = prefix[index]! + (sizes[index] ?? 0);
    axisPrefixes.set(sizes, prefix);
  }
  return prefix;
}

/**
 * A point of a drawing's frame on the printed lines: Windows Excel prints a
 * drawing by its markers in the printed grid, the line it starts in and its
 * offset there (never past the line's printed size), so the frame stays on
 * its cells however the lines above round to the printer's whole dots (the
 * API's spreadsheet_drawing_geometry; R361-X03's page 6 text boxes, anchored
 * 120 rows down, printed 7.6 pt low where the frame was scaled whole).
 * `source` are the lines' sizes the frame is measured in, `printed` the
 * lines' printed sizes, both indexed by line.
 */
export function pageLayoutPrintedAxisPoint(point: number, source: readonly number[], printed: readonly number[]): number {
  if (!Number.isFinite(point) || source.length === 0 || source.length !== printed.length) return point;
  const from = axisPrefix(source);
  const to = axisPrefix(printed);
  // The last line starting at or before the point.
  let low = 0, high = source.length - 1;
  while (low < high) {
    const middle = (low + high + 1) >> 1;
    if (from[middle]! <= point) low = middle;
    else high = middle - 1;
  }
  const offset = Math.max(0, point - from[low]!);
  return to[low]! + (low === source.length - 1 ? offset : Math.min(offset, printed[low] ?? offset));
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
  // Print titles that are rows (columns) inside the range print once, as
  // titles: the range's own share leaves them out (Windows Excel, lookwin
  // R59-X18: titles 3:22 of a range from row 1; the API's project_axis takes
  // the titles out of the body too). They print in their own place on the
  // page that runs past them (pageTitledIndices).
  const titlesInside = repeatCount > 0 && repeatStart > dataStart && repeatStart <= dataStart + dataCount - 1;
  if (input.hidden?.size || titlesInside) {
    const isTitle = (index: number) => repeatCount > 0 && index >= repeatStart && index <= repeatEnd;
    const repeatIndices = Array.from({ length: repeatCount }, (_, i) => repeatStart + i).filter(i => !input.hidden?.has(i));
    const dataIndices = Array.from({ length: dataCount }, (_, i) => dataStart + i).filter(i => !input.hidden?.has(i) && !isTitle(i));
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
    rowSizes: input.rowSizes?.map((size) => printedRowPoints(size, scale)),
    colSizes: input.colSizes?.map((size) => printedColumnPoints(size, scale)),
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

/** The print titles that are rows or columns of the printed range itself
 * but laid out only as titles (A:C of a range A:AA): Fit To counts them a
 * second time, as part of the range (fitTitlesTwice). */
function titlesInRangeSizes(
  projection: PageLayoutAxisProjection,
  sizes: readonly number[] | undefined,
  areaStart: number,
  areaEnd: number,
  sourceOf: (actual: number) => number = (actual) => actual,
): number[] {
  const result: number[] = [];
  for (let local = 0; local < projection.repeatCount; local += 1) {
    const actual = mapLocalAxisIndex(projection, local);
    if (actual === null) continue;
    const source = sourceOf(actual);
    if (source < areaStart || source > areaEnd) continue;
    // Titles the range's own run holds already (F:H of A:AA) are in it.
    if (!projection.visibleIndices && actual >= projection.dataStart && actual < projection.dataStart + projection.dataCount) continue;
    const size = sizes?.[local];
    if (typeof size === "number" && Number.isFinite(size) && size > 0) result.push(size);
  }
  return result;
}

/** Fit To counts the print titles on top of the whole range, those that are
 * part of it too, as if every page, the first included, printed them before
 * its share of the range (lookwin p74-lw/p74b-lw: 40 columns Fit 1 wide print
 * 37%, with titles A or A:C 36% / 35%, F:H 35%; 200 rows 1 tall 26%, titles
 * 50:59 25%, 1:20 24%; R87-X14 51% -> 48%, as Excel). The pages then print
 * each title once. The API's fit (titles_in_range) is the same rule. */
function fitTitlesTwice(
  input: PaginateInput,
  twice: { rows: readonly number[]; cols: readonly number[] },
): PaginateInput {
  const scaling = input.settings.scaling;
  if (!scaling || scaling.kind === "percent") return input;
  const fitted = (axis: "row" | "col") =>
    scaling.kind === "fitToWidth" ? axis === "col"
      : "widthPages" in scaling ? (axis === "row" ? scaling.heightPages : scaling.widthPages) > 0
        : true;
  const rows = fitted("row") ? twice.rows : [];
  const cols = fitted("col") ? twice.cols : [];
  if (!rows.length && !cols.length) return input;
  const titleCount = (range: string | undefined, rowsAxis: boolean) => {
    const parsed = range ? parseA1Range(range) : null;
    return parsed ? (rowsAxis ? parsed.r2 - parsed.r1 + 1 : parsed.c2 - parsed.c1 + 1) : 0;
  };
  const insert = (sizes: readonly number[] | undefined, total: number, fallback: number, at: number, extra: readonly number[]) => {
    const all = Array.from({ length: total }, (_, index) => {
      const size = sizes?.[index];
      return typeof size === "number" && Number.isFinite(size) && size > 0 ? size : fallback;
    });
    return [...all.slice(0, at), ...extra, ...all.slice(at)];
  };
  const rowTitles = Math.min(titleCount(input.settings.repeatTitleRows, true), input.totalRows);
  const colTitles = Math.min(titleCount(input.settings.repeatTitleCols, false), input.totalCols);
  return {
    ...input,
    totalRows: input.totalRows + rows.length,
    totalCols: input.totalCols + cols.length,
    rowSizes: rows.length ? insert(input.rowSizes, input.totalRows, input.rowHeight, rowTitles, rows) : input.rowSizes,
    colSizes: cols.length ? insert(input.colSizes, input.totalCols, input.colWidth, colTitles, cols) : input.colSizes,
  };
}

/** The screen width (px) past which a column prints cut at the page's
 * printable width (`dataAreaWidth`, points) at a percent scale: Windows Excel
 * prints such a column at the page's width (lookwin p48-lw, p61-lw).
 * Infinity under Fit To, whose scale shrinks the columns anyway. */
export function pageLayoutWideColumnLimitPx(
  dataAreaWidth: number,
  scaling: PrintSettings["scaling"],
  printScale?: OfficePrintScale,
): number {
  if (scaling && scaling.kind !== "percent") return Infinity;
  const colPx = SPREADSHEET_PAGE_LAYOUT_PX_TO_PT * validPrintScale(printScale).x;
  return dataAreaWidth / (colPx * Math.max(0.1, Math.min(4, scaling ? scaling.percent / 100 : 1)));
}

/** The page's printable width in points for a sheet's print settings. */
export function pageLayoutDataAreaWidth(printSettings: PrintSettings | undefined): number {
  const settings = normalizePrintSettings(printSettings) ?? {};
  const stock = resolvePaperSizePoints(settings.paperSize);
  const paperWidth = settings.orientation === "landscape" ? stock.height : stock.width;
  const margins = resolveMargins(settings.margins);
  return Math.max(1, paperWidth - margins.left - margins.right);
}

/** Resolve fit policies by asking paginateSheet itself, never a second pager. */
function resolveContentScale(
  pageInput: PaginateInput,
  override?: number,
  titlesTwice: { rows: readonly number[]; cols: readonly number[] } = { rows: [], cols: [] },
): number {
  if (typeof override === "number" && Number.isFinite(override) && override > 0) {
    return Math.max(0.1, Math.min(4, override));
  }
  const scaling = pageInput.settings.scaling;
  if (!scaling) return 1;
  if (scaling.kind === "percent") {
    return Math.max(0.1, Math.min(4, scaling.percent / 100));
  }
  const input = fitTitlesTwice(pageInput, titlesTwice);
  // A column (row) wider (taller) than the page is a page of its own at any
  // scale, so the page count alone never asks Fit To to shrink it: Windows
  // Excel shrinks until it fits the page (R450-X05's Contents, a 188.6-wide
  // column A fitted 1 x 1, prints at 46 %; the API's writer fits the whole
  // columns too). Its widest one, judged at its printed width.
  const widest = (sizes: readonly number[] | undefined, fallback: number) => {
    let most = fallback;
    for (const size of sizes ?? []) if (Number.isFinite(size) && size > most) most = size;
    return most;
  };
  const widestCol = widest(input.colSizes, input.colWidth);
  const tallestRow = widest(input.rowSizes, input.rowHeight);
  const spare = printPageSparePoints(scaling);
  const colFits = (scale: number) =>
    printedColumnPoints(widestCol, scale) + (input.headingWidth ?? 0) * scale <= input.dataAreaWidth - spare + 1e-6;
  const rowFits = (scale: number) =>
    printedRowPoints(tallestRow, scale) + (input.headingHeight ?? 0) * scale <= input.dataAreaHeight - spare + 1e-6;
  const fits = (scale: number) => {
    const result = pageCountAtScale(input, scale);
    if (scaling.kind === "fitToWidth") {
      return result.pagination.colRanges.length <= Math.max(1, scaling.pages) && colFits(scale);
    }
    if ("pages" in scaling) {
      return result.pageCount <= Math.max(1, scaling.pages) && colFits(scale) && rowFits(scale);
    }
    return (
      (scaling.widthPages === 0 ||
        (result.pagination.colRanges.length <= scaling.widthPages && colFits(scale))) &&
      (scaling.heightPages === 0 ||
        (result.pagination.rowRanges.length <= scaling.heightPages && rowFits(scale)))
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
  /** Windows Excel's printer scale of this workbook's grid (officePrintScale
   * of its Normal font): Excel pages and prints rows and columns at their
   * printed sizes, so the pages break where its own do. */
  printScale?: OfficePrintScale;
}): SpreadsheetPageLayoutGeometry {
  const settings = normalizePrintSettings(input.printSettings) ?? {};
  const printAreaBlocks = printAreaRanges(settings.printArea);
  if (printAreaBlocks.length > 1) {
    return resolveMultiAreaPageLayoutGeometry(input, settings, printAreaBlocks);
  }
  const printScale = validPrintScale(input.printScale);
  // Sizes in points as the printer lays them out; the px maps below stay in
  // screen pixels (rowPx / colPx), which is what the sheet's maps hold.
  const rowPx = SPREADSHEET_PAGE_LAYOUT_PX_TO_PT * printScale.y;
  const colPx = SPREADSHEET_PAGE_LAYOUT_PX_TO_PT * printScale.x;
  const baseRowHeight = finitePositive(
    finitePositive(input.defaultRowHeight, input.rowHeight ?? DEFAULT_SPREADSHEET_ROW_HEIGHT_PX) * rowPx,
    SPREADSHEET_PAGE_LAYOUT_DEFAULT_ROW_HEIGHT_PT,
  );
  const baseColWidth = finitePositive(
    finitePositive(input.defaultColumnWidth, input.colWidth ?? DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX) * colPx,
    SPREADSHEET_PAGE_LAYOUT_DEFAULT_COLUMN_WIDTH_PT,
  );
  const sourceTotalRows = boundedDimension(input.totalRows);
  const sourceTotalCols = boundedDimension(input.totalCols);
  const printArea = settings.printArea
    ? parseA1Range(settings.printArea)
    : null;
  const stock = resolvePaperSizePoints(settings.paperSize);
  const landscape = settings.orientation === "landscape";
  const paperWidth = landscape ? stock.height : stock.width;
  const paperHeight = landscape ? stock.width : stock.height;
  const sourceMargins = resolveMargins(settings.margins);
  // Native Mac worksheet origins use whole point pixels. Keep authored
  // precision separately so merely viewing or grabbing a ruler cannot edit it.
  // An imported Windows grid's are whole printer dots (600 dpi), truncated,
  // for where the page prints and the room Fit To fits in (lookwin probes
  // p162-lw/p162b-lw: 259.82 + 260.82 dots fit as 259 + 260; R13-X27's
  // 448.82-dot margins made Table 4 fit 57% where Excel prints 58%).
  const dots = (value: number) => Math.floor((value * 600) / 72 + 1e-6) * 72 / 600;
  const margins = input.nativeGridDpi === 72
    ? { left: Math.floor(sourceMargins.left), right: Math.floor(sourceMargins.right),
        top: Math.floor(sourceMargins.top), bottom: Math.floor(sourceMargins.bottom) }
    : input.nativeGridDpi === 96
      ? { left: dots(sourceMargins.left), right: dots(sourceMargins.right), top: dots(sourceMargins.top), bottom: dots(sourceMargins.bottom) }
      : sourceMargins;
  const dataAreaWidth = Math.max(1, paperWidth - margins.left - margins.right);
  const dataAreaHeight = Math.max(
    1,
    paperHeight - margins.top - margins.bottom,
  );
  const columnPx = (index: number) =>
    input.hiddenCols?.has(index) ? 0 : finitePositive(input.columnWidths?.[String(index)], baseColWidth / colPx);
  // A column wider than the page prints at the page's width, and a drawing
  // right of it moves left with it (pageLayoutWideColumns): the columns its
  // reach needs are counted at those printed widths. Taken at the stated
  // percent; a Fit To scale is not known yet and shrinks the columns anyway.
  const scaling = settings.scaling;
  const wideLimitPx = pageLayoutWideColumnLimitPx(dataAreaWidth, scaling, printScale);
  let drawingWide: PageLayoutWideColumns | null = null;
  if (!printArea && Number.isFinite(wideLimitPx) && input.drawingBounds?.length) {
    const reach = Math.max(0, ...input.drawingBounds.map((bounds) =>
      Array.isArray(bounds) && bounds.length === 4 && bounds.every(Number.isFinite) ? bounds[0] : 0));
    const sizes: number[] = [];
    for (let index = 0, position = 0; position <= reach && index < Math.min(MAX_GRID_COLUMNS, MAX_RENDERED_PAGE_LAYOUT_CELLS_TOTAL); index += 1) {
      sizes.push(columnPx(index));
      position += sizes[index];
    }
    drawingWide = pageLayoutWideColumns(sizes, wideLimitPx);
  }
  let drawingRight = 0, drawingBottom = 0;
  if (!printArea) {
    for (const bounds of input.drawingBounds ?? []) {
      if (!Array.isArray(bounds) || bounds.length !== 4 || !bounds.every(Number.isFinite)) continue;
      const [left, top, width, height] = bounds;
      const right = pageLayoutClampColumnPoint(left, drawingWide) + width, bottom = top + height;
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
    input.hiddenRows?.has(index) ? 0 : finitePositive(input.rowHeights?.[String(index)], baseRowHeight / rowPx));
  const totalCols = drawingAxisCount(Math.max(sourceTotalCols, areaCols), drawingRight, MAX_GRID_COLUMNS, index =>
    Math.min(columnPx(index), wideLimitPx));
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
  // Windows Excel prints a row of a fractional pixel height at its nearest
  // quarter pixel (lookwin p139-lw: 17.1 pt rows, 22.8 px, 137 dots apart,
  // not 138; p173-lw: 31.05 pt, 41.4 px, 251 dots as 41.5 px; the API's
  // axis_sizes).
  const sourceRowSizes = Array.from({ length: totalRows }, (_, index) =>
    input.hiddenRows?.has(index) ? 0 : Math.floor(finitePositive(input.rowHeights?.[String(index)], baseRowHeight / rowPx) * 4 + 0.5) / 4 * rowPx,
  );
  const sourceColSizes = Array.from({ length: totalCols }, (_, index) =>
    input.hiddenCols?.has(index) ? 0 : finitePositive(input.columnWidths?.[String(index)], baseColWidth / colPx) * colPx,
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
  // Windows Excel's printer (lookwin probe p18: Calibri 11, Arial 10, ＭＳ Ｐゴシック
  // 11): the digits at the printed digit width plus 32 printer dots, 20.76 pt
  // for three Calibri 11 digits, 26.4 pt for four.
  const unscaledHeadingWidth = settings.showHeadings !== true
    ? 0
    : printScale !== OFFICE_PRINT_IDENTITY
      ? ((headingDigits * headingDigit * 72) / headingDpi) * printScale.x + (32 * 72) / 600
      : ((headingDigits * headingDigit + Math.round(headingDigit * 0.8)) * 72) / headingDpi;
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
  const contentScale = resolveContentScale(unscaledPaginateInput, input.contentScaleOverride, {
    rows: titlesInRangeSizes(rowProjection, unscaledPaginateInput.rowSizes, printArea?.r1 ?? 0, printArea?.r2 ?? totalRows - 1,
      (actual) => rowOrder?.[actual] ?? actual),
    cols: titlesInRangeSizes(colProjection, unscaledPaginateInput.colSizes, printArea?.c1 ?? 0, printArea?.c2 ?? totalCols - 1),
  });
  const rowHeight = baseRowHeight * contentScale;
  const colWidth = baseColWidth * contentScale;
  const paginateInput: PaginateInput = {
    ...unscaledPaginateInput,
    rowHeight,
    colWidth,
    rowSizes: unscaledPaginateInput.rowSizes?.map((size) => printedRowPoints(size, contentScale)),
    colSizes: unscaledPaginateInput.colSizes?.map((size) => printedColumnPoints(size, contentScale)),
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
    ...(input.nativeGridDpi === 72 || input.nativeGridDpi === 96 ? { sourceMargins } : {}),
    headerMargin: finiteNonNegative(settings.headerMargin, 21.6),
    footerMargin: finiteNonNegative(settings.footerMargin, 21.6),
    dataAreaWidth,
    dataAreaHeight,
    rowHeight,
    colWidth,
    rowSizes: sourceRowSizes.map((size) => printedRowPoints(size, contentScale)),
    colSizes: sourceColSizes.map((size) => printedColumnPoints(size, contentScale)),
    // A drawing's frame is measured on the screen's own heights, before the
    // printer's quarter pixels (sourceRowSizes floors them).
    sourceRowPoints: Array.from({ length: totalRows }, (_, index) => input.hiddenRows?.has(index) ? 0
      : finitePositive(input.rowHeights?.[String(index)], baseRowHeight / rowPx) * rowPx * contentScale),
    sourceColPoints: sourceColSizes.map((size) => size * contentScale),
    contentScale,
    headingHeight: unscaledHeadingHeight * contentScale,
    headingWidth: unscaledHeadingWidth * contentScale,
    paginateInput,
    pagination,
    totalPages,
    rowProjection,
    colProjection,
    ...(printScale !== OFFICE_PRINT_IDENTITY ? { printScale } : {}),
  };
}

function validPrintScale(scale: OfficePrintScale | undefined): OfficePrintScale {
  // A factor past these bounds is not a printer's metrics (the measured
  // faces lie between 0.85 and 1.17): it would only distort the page.
  const ok = (value: unknown) => typeof value === "number" && Number.isFinite(value) && value >= 0.5 && value <= 2;
  return scale && ok(scale.x) && ok(scale.y) && (scale.x !== 1 || scale.y !== 1)
    ? { x: scale.x, y: scale.y }
    : OFFICE_PRINT_IDENTITY;
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
 * How many of the sheet's pages print: Excel leaves out the trailing blank
 * pages ("Trailing blank page suppression is the default behavior of Excel"),
 * those after the last page with anything on it, in page order; the first
 * page always prints, a blank page before a printing one prints, and a print
 * area prints every page. `pagePrints` says whether a page (its own rows and
 * columns, as live worksheet indices, and the print titles printed with it)
 * holds anything: a value, a fill or border of a cell of its own, a drawing,
 * a spill, a row's `spans` (lookwin p67-lw). The API's page_meta blank /
 * dropped is the same rule.
 */
export function pageLayoutPrintedPageCount(
  geometry: SpreadsheetPageLayoutGeometry,
  pagePrints: (rows: readonly number[], cols: readonly number[],
    titleRows: readonly number[], titleCols: readonly number[]) => boolean,
): number {
  if (geometry.segments || geometry.settings.printArea || geometry.totalPages <= 1) return geometry.totalPages;
  const indices = (projection: PageLayoutAxisProjection, first: number, last: number) => {
    const result: number[] = [];
    for (let local = first; local <= last; local += 1) {
      const actual = mapLocalAxisIndex(projection, local);
      if (actual !== null) result.push(actual);
    }
    return result;
  };
  const titleRows = indices(geometry.rowProjection, 0, geometry.rowProjection.repeatCount - 1);
  const titleCols = indices(geometry.colProjection, 0, geometry.colProjection.repeatCount - 1);
  for (let index = geometry.totalPages - 1; index > 0; index -= 1) {
    const { rowPageIndex, colPageIndex } = axisPageIndices(geometry, index);
    const rowRange = geometry.pagination.rowRanges[rowPageIndex];
    const colRange = geometry.pagination.colRanges[colPageIndex];
    if (!rowRange || !colRange) return index + 1;
    if (pagePrints(indices(geometry.rowProjection, rowRange[0], rowRange[1]),
      indices(geometry.colProjection, colRange[0], colRange[1]), titleRows, titleCols)) return index + 1;
  }
  return 1;
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
  const titles: number[] = [];
  for (let index = 0; index < projection.repeatCount; index += 1) {
    const mapped = mapLocalAxisIndex(projection, index);
    if (mapped !== null) titles.push(mapped);
  }
  const dataStart = Math.max(range[0], projection.repeatCount);
  const dataEnd = Math.min(
    range[1],
    projection.repeatCount + projection.dataCount - 1,
  );
  const body: number[] = [];
  for (let index = dataStart; index <= dataEnd && titles.length + body.length < limit; index += 1) {
    const mapped = mapLocalAxisIndex(projection, index);
    if (mapped !== null) body.push(mapped);
  }
  return pageTitledIndices(titles, body).slice(0, limit);
}

/**
 * A page's rows (or columns) in print order: its print titles, then its
 * share of the range -- except on the page whose share runs from before the
 * titles to past them, where the titles print in their own place, once
 * (Windows Excel, lookwin R59-X18: titles 3:22 of a range from row 1, the
 * first page prints rows 1-33 in order; later pages print the titles above
 * their share). The API's titled_lines is the same rule.
 */
export function pageTitledIndices(titles: readonly number[], body: readonly number[]): number[] {
  const firstTitle = titles[0];
  const lastTitle = titles[titles.length - 1];
  if (firstTitle === undefined || lastTitle === undefined || !body.length) return [...titles, ...body];
  const before = body.filter((index) => index < firstTitle);
  const after = body.filter((index) => index > lastTitle);
  if (!before.length || !after.length || before.length + after.length !== body.length) return [...titles, ...body];
  return [...before, ...titles, ...after];
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
