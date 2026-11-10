/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { hasHiddenAxisLines, lineIndexAtOffset, nextVisibleAxisLine, totalAxisSize } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

export interface GridScrollViewport { left: number; top: number; width: number; height: number }
export interface GridScrollGeometry {
  rowCount: number; columnCount: number;
  rowSize: (index: number) => number; columnSize: (index: number) => number;
  zoom: number; headerHeight: number; headerWidth: number;
  rowBudget: number; columnBudget: number; rowOverscan: number; columnOverscan: number;
  /** Directional prefetch reallocates the existing buffer, never enlarges it. */
  rowDirection?: -1 | 0 | 1; columnDirection?: -1 | 0 | 1;
  /** A fling: the last frame moved at least half a viewport. The buffer is
   *  then laid almost entirely ahead of the visible band, so the compositor,
   *  which keeps scrolling while the next band mounts, meets rendered cells
   *  for as long as possible. */
  rowFling?: boolean; columnFling?: boolean;
}

function visibleBand(viewport: GridScrollViewport, geometry: GridScrollGeometry) {
  const { rowCount, columnCount, rowSize, columnSize, headerHeight, headerWidth } = geometry;
  const zoom = Number.isFinite(geometry.zoom) && geometry.zoom > 0 ? geometry.zoom : 1;
  const rowAt = (position: number) => lineIndexAtOffset(Math.max(0, position / zoom - headerHeight), rowCount, rowSize);
  const columnAt = (position: number) => lineIndexAtOffset(Math.max(0, position / zoom - headerWidth), columnCount, columnSize);
  return { firstRow: rowAt(viewport.top), lastRow: rowAt(viewport.top + Math.max(0, viewport.height)),
    firstColumn: columnAt(Math.abs(viewport.left)), lastColumn: columnAt(Math.abs(viewport.left) + Math.max(0, viewport.width)) };
}

function visibleAxisWindow(first:number,last:number,count:number,size:(index:number)=>number,budget:number,buffer:number,direction=0) {
  let start=nextVisibleAxisLine(first,count,size);
  if(start>=count)start=nextVisibleAxisLine(count-1,count,size,-1);
  if(start<0)return {start:0,end:count,indices:[] as number[]};
  const endVisible=Math.max(start,nextVisibleAxisLine(last,count,size,-1));
  const indices:number[]=[];
  for(let index=start;index<count&&index<=endVisible;index=nextVisibleAxisLine(index+1,count,size))indices.push(index);
  const target=Math.max(budget,indices.length+2*buffer);
  const before=direction>0?Math.ceil(buffer/2):direction<0?Math.floor(buffer*1.5):buffer;
  for(let n=0;n<before;n++){
    const index=nextVisibleAxisLine(indices[0]-1,count,size,-1);
    if(index<0)break;indices.unshift(index);
  }
  while(indices.length<target){
    const index=nextVisibleAxisLine(indices[indices.length-1]+1,count,size);
    if(index>=count)break;indices.push(index);
  }
  while(indices.length<target){
    const index=nextVisibleAxisLine(indices[0]-1,count,size,-1);
    if(index<0)break;indices.unshift(index);
  }
  return {start:indices[0],end:indices[indices.length-1]+1,indices};
}

/**
 * The viewport a scroll is about to show as well as the one on screen: the
 * span from the current offset to where the motion is heading (`reach`, in
 * scroll pixels, negative towards the start). A window resolved for the span
 * holds every band the compositor passes on the way there, so it can be
 * mounted before the scroll moves rather than after the band came into view.
 */
export function spreadsheetGridReachViewport(viewport: GridScrollViewport, reach: { top: number; left: number }): GridScrollViewport {
  const top = Number.isFinite(reach.top) ? reach.top : 0;
  const left = Number.isFinite(reach.left) ? reach.left : 0;
  return {
    top: Math.max(0, viewport.top + Math.min(0, top)),
    left: Math.max(0, viewport.left + Math.min(0, left)),
    width: viewport.width + Math.abs(left),
    height: viewport.height + Math.abs(top),
  };
}

/** The rows and columns on screen plus the overscan on each side. The fixed
 * budgets only stand in for a viewport that has not been measured yet (the
 * first render, a test DOM): a measured one mounts what it shows, so a sheet
 * switch or a new band costs what the screen holds, not a fixed 64 x 24. The
 * worksheet's full address space never becomes a DOM budget. */
export function spreadsheetGridWindow(viewport: GridScrollViewport, geometry: GridScrollGeometry) {
  const visible = visibleBand(viewport, geometry);
  const rowFloor = viewport.height > 0 ? 0 : geometry.rowBudget;
  const columnFloor = viewport.width > 0 ? 0 : geometry.columnBudget;
  const rowBudget = Math.max(rowFloor, visible.lastRow - visible.firstRow + 1 + 2 * geometry.rowOverscan);
  const columnBudget = Math.max(columnFloor, visible.lastColumn - visible.firstColumn + 1 + 2 * geometry.columnOverscan);
  const before = (buffer: number, direction = 0, fling = false, budget = 0, visibleCount = 0) => {
    // The trailing side keeps the coverage guard's few lines (see
    // spreadsheetGridWindowCovers), so a fling frame that only inched does
    // not force another full band.
    const trailing = Math.min(4, Math.ceil(buffer / 2));
    if (fling && direction > 0) return trailing;
    if (fling && direction < 0) return Math.max(buffer, budget - visibleCount - trailing);
    return direction > 0 ? Math.ceil(buffer / 2) : direction < 0 ? Math.floor(buffer * 1.5) : buffer;
  };
  const rowStart = Math.max(0, Math.min(visible.firstRow - before(geometry.rowOverscan, geometry.rowDirection, geometry.rowFling, rowBudget, visible.lastRow - visible.firstRow + 1), geometry.rowCount - rowBudget));
  const colStart = Math.max(0, Math.min(visible.firstColumn - before(geometry.columnOverscan, geometry.columnDirection, geometry.columnFling, columnBudget, visible.lastColumn - visible.firstColumn + 1), geometry.columnCount - columnBudget));
  const rowEnd = Math.min(geometry.rowCount, rowStart + rowBudget);
  const colEnd = Math.min(geometry.columnCount, colStart + columnBudget);
  const maskedRows=hasHiddenAxisLines(geometry.rowSize)?visibleAxisWindow(visible.firstRow,visible.lastRow,geometry.rowCount,geometry.rowSize,rowFloor,geometry.rowOverscan,geometry.rowDirection):null;
  const maskedCols=hasHiddenAxisLines(geometry.columnSize)?visibleAxisWindow(visible.firstColumn,visible.lastColumn,geometry.columnCount,geometry.columnSize,columnFloor,geometry.columnOverscan,geometry.columnDirection):null;
  return { rowStart:maskedRows?.start??rowStart, colStart:maskedCols?.start??colStart, rowEnd:maskedRows?.end??rowEnd, colEnd:maskedCols?.end??colEnd,
    rows: maskedRows?.indices??Array.from({ length: rowEnd - rowStart }, (_, index) => rowStart + index),
    cols: maskedCols?.indices??Array.from({ length: colEnd - colStart }, (_, index) => colStart + index),
    ...spreadsheetGridStageSize(geometry) };
}

/** The whole sheet's stage, headings included, in content pixels: it does not move with the scroll. */
export function spreadsheetGridStageSize(geometry: Pick<GridScrollGeometry, "headerWidth" | "headerHeight" | "rowCount" | "columnCount" | "rowSize" | "columnSize">) {
  return {
    totalWidth: geometry.headerWidth + totalAxisSize(geometry.columnCount, geometry.columnSize),
    totalHeight: geometry.headerHeight + totalAxisSize(geometry.rowCount, geometry.rowSize),
  };
}

/**
 * A merged cell is one table cell at its anchor, as wide and as tall as the
 * lines of it that are mounted. A window that cuts through a merge would draw
 * its text centred in the mounted part (a title merged across the sheet sits
 * off-centre) or, with the anchor outside, not at all. The window therefore
 * takes in the whole of every merge it touches, up to `limit` lines beyond
 * itself on each side, so a pathological merge cannot become a DOM budget.
 */
export function spreadsheetGridWindowWithMerges<W extends ReturnType<typeof spreadsheetGridWindow>>(
  window: W,
  merges: ReadonlyArray<{ row: number; col: number; rowSpan: number; colSpan: number }>,
  geometry: Pick<GridScrollGeometry, "rowCount" | "columnCount" | "rowSize" | "columnSize">,
  limit = 48,
): W {
  if (merges.length === 0 || window.rows.length === 0 || window.cols.length === 0) return window;
  let { rowStart, rowEnd, colStart, colEnd } = window;
  for (const merge of merges) {
    if (merge.rowSpan <= 1 && merge.colSpan <= 1) continue;
    const lastRow = merge.row + merge.rowSpan, lastCol = merge.col + merge.colSpan;
    if (merge.row >= window.rowEnd || lastRow <= window.rowStart || merge.col >= window.colEnd || lastCol <= window.colStart) continue;
    rowStart = Math.min(rowStart, merge.row); rowEnd = Math.max(rowEnd, lastRow);
    colStart = Math.min(colStart, merge.col); colEnd = Math.max(colEnd, lastCol);
  }
  rowStart = Math.max(0, rowStart, window.rowStart - limit); rowEnd = Math.min(geometry.rowCount, rowEnd, window.rowEnd + limit);
  colStart = Math.max(0, colStart, window.colStart - limit); colEnd = Math.min(geometry.columnCount, colEnd, window.colEnd + limit);
  if (rowStart === window.rowStart && rowEnd === window.rowEnd && colStart === window.colStart && colEnd === window.colEnd) return window;
  const lines = (start: number, end: number, size: (index: number) => number) => {
    const masked = hasHiddenAxisLines(size), indices: number[] = [];
    for (let index = start; index < end; index++) if (!masked || size(index) > 0) indices.push(index);
    return indices;
  };
  return {
    ...window, rowStart, rowEnd, colStart, colEnd,
    rows: rowStart === window.rowStart && rowEnd === window.rowEnd ? window.rows : lines(rowStart, rowEnd, geometry.rowSize),
    cols: colStart === window.colStart && colEnd === window.colEnd ? window.cols : lines(colStart, colEnd, geometry.columnSize),
  };
}

/** The rows a viewport shows, by its scroll offset and height. */
export function spreadsheetGridVisibleRows(viewport: GridScrollViewport, geometry: GridScrollGeometry): { firstRow: number; lastRow: number } {
  const { firstRow, lastRow } = visibleBand(viewport, geometry);
  return { firstRow, lastRow };
}

/**
 * A band that comes in a few rows a frame. Mounting a whole new band of rows
 * in one frame is the grid's one remaining long task while scrolling (40–50
 * ms for a 22-column band, almost all of it React mounting the cells); the
 * compositor shows the gridline placeholder where cells are not yet mounted.
 * So a band that overlaps the rows already in the DOM keeps those and adds at
 * most `budget` rows beside them, the ones the scroll is heading for first;
 * the host asks again next frame until the band is whole. A band that shares
 * no row with the DOM (a jump: Go To, the name box, a scroll bar drag far
 * away) is mounted whole, as before — there is nothing on screen to keep.
 */
export function revealGridWindowRows<W extends { rows: number[]; rowStart: number; rowEnd: number }>(
  window: W,
  mounted: readonly number[],
  budget: number,
  direction: -1 | 0 | 1 = 0,
  /** The rows on screen (or about to be, for a wheel step): mounted in this
   * frame whatever the budget, with the rows between them and the DOM's. */
  required: { firstRow: number; lastRow: number } | null = null,
  /** A page key's band: drawn from `required` even with none of it mounted,
   * so a whole screen is not one long frame. */
  sliceFromNothing = false,
): W {
  if (window.rows.length === 0 || budget <= 0) return window;
  const have = new Set(mounted);
  const mountedRows = window.rows.filter((row) => have.has(row));
  if (window.rows.length - mountedRows.length <= budget) return window;
  if (mountedRows.length === 0 && !(sliceFromNothing && required)) return window;
  let lo = mountedRows[0] ?? Number.POSITIVE_INFINITY, hi = mountedRows[mountedRows.length - 1] ?? Number.NEGATIVE_INFINITY;
  if (required) {
    const onScreen = window.rows.filter((row) => row >= required.firstRow && row <= required.lastRow);
    if (onScreen.length > 0) { lo = Math.min(lo, onScreen[0]); hi = Math.max(hi, onScreen[onScreen.length - 1]); }
  }
  const kept = window.rows.filter((row) => row >= lo && row <= hi);
  if (window.rows.length - kept.length <= budget) return window;
  const first = kept[0], last = kept[kept.length - 1];
  const after = window.rows.filter((row) => row > last);
  const before = window.rows.filter((row) => row < first).reverse();
  // From nothing, the rows asked for are the frame: a heavy sheet's 20 rows
  // held a single page key 700 ms.
  const taken = mountedRows.length === 0 ? [] : (direction < 0 ? [...before, ...after] : [...after, ...before]).slice(0, budget);
  const rows = [...kept, ...taken].sort((a, b) => a - b);
  return { ...window, rows, rowStart: rows[0], rowEnd: rows[rows.length - 1] + 1 };
}

/** Keep the unchanged axis stable across scroll updates. Only the latest two
 * numeric bands are retained; no geometry callback or workbook is cached. */
export function createSpreadsheetGridWindowResolver() {
  let rows: number[] | undefined;
  let cols: number[] | undefined;
  const equal = (left: readonly number[] | undefined, right: readonly number[]) =>
    left?.length === right.length && right.every((value, index) => value === left[index]);
  return (viewport: GridScrollViewport, geometry: GridScrollGeometry) => {
    const next = spreadsheetGridWindow(viewport, geometry);
    if (rows && equal(rows, next.rows)) next.rows = rows;
    else rows = next.rows;
    if (cols && equal(cols, next.cols)) next.cols = cols;
    else cols = next.cols;
    return next;
  };
}

/** Hysteresis: keep existing DOM until the visible rectangle approaches an
 * overscan edge, rather than rebuilding all cells for each scrolled pixel. */
export function spreadsheetGridWindowCovers(viewport: GridScrollViewport, window: ReturnType<typeof spreadsheetGridWindow>, geometry: GridScrollGeometry): boolean {
  if (viewport.width <= 0 || viewport.height <= 0) return false;
  const zoom = Number.isFinite(geometry.zoom) && geometry.zoom > 0 ? geometry.zoom : 1;
  // Safari can report a trailing-edge elastic scroll past the old extent.
  // Index lookup clamps to its final line, which otherwise falsely declares
  // coverage while capacity growth recomputes the band from stale scroll state.
  if (Math.abs(viewport.left) + viewport.width > window.totalWidth * zoom + 1 ||
    Math.max(0, viewport.top) + viewport.height > window.totalHeight * zoom + 1) return false;
  const visible = visibleBand(viewport, geometry);
  const rowGuard = Math.min(3, Math.floor(geometry.rowOverscan / 2));
  const colGuard = Math.min(1, Math.floor(geometry.columnOverscan / 2));
  const starts=(indices:readonly number[],start:number,first:number,guard:number,size:(index:number)=>number,count:number)=>
    nextVisibleAxisLine(start-1,count,size,-1)<0||first>=(indices[Math.min(guard,indices.length-1)]??start);
  const ends=(indices:readonly number[],end:number,last:number,guard:number,size:(index:number)=>number,count:number)=>
    nextVisibleAxisLine(end,count,size)>=count||last<(indices[Math.max(0,indices.length-guard)]??end);
  return starts(window.rows,window.rowStart,visible.firstRow,rowGuard,geometry.rowSize,geometry.rowCount) &&
    ends(window.rows,window.rowEnd,visible.lastRow,rowGuard,geometry.rowSize,geometry.rowCount) &&
    starts(window.cols,window.colStart,visible.firstColumn,colGuard,geometry.columnSize,geometry.columnCount) &&
    ends(window.cols,window.colEnd,visible.lastColumn,colGuard,geometry.columnSize,geometry.columnCount);
}

/** One cancellable animation frame owns the latest event in a burst. */
export function latestScrollFrame<T>(request: (callback: () => void) => number, cancel: (id: number) => void, publish: (value: T) => void) {
  let frame: number | null = null;
  let generation = 0;
  let latest: T;
  return {
    schedule(value: T) {
      latest = value;
      if (frame !== null) return;
      frame = -1;
      const owner = ++generation;
      const id = request(() => { if (owner !== generation) return; frame = null; publish(latest); });
      // Also support hosts that deliver a frame synchronously.
      if (frame !== null) frame = id;
    },
    cancel() { generation += 1; if (frame !== null && frame >= 0) cancel(frame); frame = null; },
  };
}
