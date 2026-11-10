/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@mrnote/core/web/lib/spreadsheet/grid-sizing.ts';
import { paginateSheet } from '@mrnote/core/web/lib/spreadsheet/print-settings.ts';
import {
  MAX_RENDERED_PAGE_LAYOUT_CELLS_PER_PAGE, MAX_RENDERED_PAGE_LAYOUT_CELLS_TOTAL,
  resolveSpreadsheetPageLayoutPage, resolveSpreadsheetPageLayoutPageAtTile, resolveSpreadsheetPageLayoutTileGrid,
  type SpreadsheetPageLayoutGeometry, type SpreadsheetPageLayoutRenderPlan,
} from './page-layout-view';

/** Browsing is a disposable extension of already scaled content geometry.
 * Print/export retain the original geometry and its totalPages. */
export function extendPageLayoutForBrowsing(
  base: SpreadsheetPageLayoutGeometry,
  minimum: { rowPages: number; colPages: number },
): SpreadsheetPageLayoutGeometry {
  if (base.settings.printArea) return base;
  const requested = (value: number) => Number.isFinite(value) ? Math.max(1, Math.ceil(value)) : 1;
  const rowTarget = requested(minimum.rowPages), colTarget = requested(minimum.colPages);
  let result = base;
  while ((result.pagination.rowRanges.length < rowTarget && result.rowSizes.length < MAX_GRID_ROWS) ||
    (result.pagination.colRanges.length < colTarget && result.colSizes.length < MAX_GRID_COLUMNS)) {
    const extend = (axis: 'row' | 'col', target: number, limit: number) => {
      const source = axis === 'row' ? result.rowSizes : result.colSizes;
      const projection = axis === 'row' ? result.rowProjection : result.colProjection;
      const local = axis === 'row' ? result.paginateInput.rowSizes : result.paginateInput.colSizes;
      const pages = axis === 'row' ? result.pagination.rowRanges.length : result.pagination.colRanges.length;
      const size = axis === 'row' ? result.rowHeight : result.colWidth;
      const capacity = axis === 'row' ? result.dataAreaHeight : result.dataAreaWidth;
      const add = Math.min(limit - source.length, Math.max(0, target - pages) * Math.max(1, Math.ceil(capacity / size)));
      if (!add) return { source, projection, local };
      const appended = Array.from({ length: add }, () => size);
      const dataCount = projection.dataCount + add;
      // Drop only the empty-domain pagination sentinel; it is not a source row.
      const localCount = projection.repeatCount + projection.dataCount;
      const indices = projection.visibleIndices ?? Array.from({ length: localCount }, (_, index) =>
        index < projection.repeatCount ? projection.repeatStart + index : projection.dataStart + index - projection.repeatCount);
      return {
        source: [...source, ...appended],
        local: [...(local ?? []).slice(0, localCount), ...appended],
        projection: { ...projection, dataCount, paginateCount: projection.repeatCount + dataCount,
          visibleIndices: [...indices, ...Array.from({ length: add }, (_, index) => source.length + index)] },
      };
    };
    const rows = extend('row', rowTarget, MAX_GRID_ROWS), cols = extend('col', colTarget, MAX_GRID_COLUMNS);
    const paginateInput = { ...result.paginateInput,
      totalRows: rows.projection.paginateCount, totalCols: cols.projection.paginateCount,
      rowSizes: rows.local, colSizes: cols.local };
    const pagination = paginateSheet(paginateInput);
    result = { ...result, rowSizes: rows.source, colSizes: cols.source,
      rowProjection: rows.projection, colProjection: cols.projection, paginateInput, pagination,
      totalPages: pagination.rowRanges.length * pagination.colRanges.length,
      contentPages: base.contentPages ?? { rowPages: base.pagination.rowRanges.length, colPages: base.pagination.colRanges.length } };
  }
  return result;
}

export interface PageLayoutViewport {
  left: number; top: number; width: number; height: number; paperScale: number;
  /** CSS tile geometry is configurable so rulers/chrome remain UI-owned. */
  tileWidth?: number; tileHeight?: number; offsetLeft?: number; offsetTop?: number; gap?: number;
}

/** Resolve nearby physical paper tiles, including indices beyond the old
 * first-120 preview limit. Cell budgets apply to mounted pages, not navigation. */
export function resolvePageLayoutViewportPlan(
  geometry: SpreadsheetPageLayoutGeometry, viewport: PageLayoutViewport, overscan = 1,
): SpreadsheetPageLayoutRenderPlan {
  const positive = (value: number | undefined, fallback: number) => value !== undefined && Number.isFinite(value) && value > 0 ? value : fallback;
  const finite = (value: number | undefined, fallback: number) => value !== undefined && Number.isFinite(value) ? value : fallback;
  const scale = positive(viewport.paperScale, 1), gap = Math.max(0, finite(viewport.gap, 24));
  const tileWidth = positive(viewport.tileWidth, geometry.paperWidth * scale);
  const tileHeight = positive(viewport.tileHeight, geometry.paperHeight * scale + 32);
  const left = finite(viewport.left, 0) - finite(viewport.offsetLeft, 80);
  const top = finite(viewport.top, 0) - finite(viewport.offsetTop, 40);
  const width = positive(viewport.width, tileWidth), height = positive(viewport.height, tileHeight);
  const extra = Math.max(0, Math.floor(finite(overscan, 1)));
  // Tiles, not pager cells: a multi-area print area stacks one grid per area.
  const { rows: rowCount, cols: colCount } = resolveSpreadsheetPageLayoutTileGrid(geometry);
  const bounds = (start: number, extent: number, tile: number, count: number) => ({
    first: Math.max(0, Math.min(count - 1, Math.floor(start / (tile + gap)))),
    last: Math.max(0, Math.min(count - 1, Math.floor((start + extent) / (tile + gap)))),
  });
  const rb = bounds(top, height, tileHeight, rowCount), cb = bounds(left, width, tileWidth, colCount);
  const candidates: { row: number; col: number; visible: boolean }[] = [];
  // Even an extreme viewport/overscan must not enumerate the whole workbook.
  // Every resolved page mounts at least one cell, so this candidate ceiling
  // cannot discard a page that fits the existing total cell budget.
  let candidateLimitReached = false;
  function* nearest(first: number, last: number) {
    const middle = Math.floor((first + last) / 2);
    for (let distance = 0; middle - distance >= first || middle + distance <= last; distance++) {
      if (middle - distance >= first) yield middle - distance;
      if (distance && middle + distance <= last) yield middle + distance;
    }
  }
  const collect = (firstRow: number, lastRow: number, firstCol: number, lastCol: number, visible: boolean) => {
    for (const row of nearest(firstRow, lastRow)) for (const col of nearest(firstCol, lastCol)) {
      if (!visible && row >= rb.first && row <= rb.last && col >= cb.first && col <= cb.last) continue;
      if (candidates.length >= MAX_RENDERED_PAGE_LAYOUT_CELLS_TOTAL) { candidateLimitReached = true; return; }
      candidates.push({ row, col, visible });
    }
  };
  collect(rb.first, rb.last, cb.first, cb.last, true);
  if (!candidateLimitReached) collect(Math.max(0, rb.first - extra), Math.min(rowCount - 1, rb.last + extra),
    Math.max(0, cb.first - extra), Math.min(colCount - 1, cb.last + extra), false);
  candidates.sort((a, b) => Number(b.visible) - Number(a.visible) ||
    Math.abs(a.row - (rb.first + rb.last) / 2) + Math.abs(a.col - (cb.first + cb.last) / 2) -
    Math.abs(b.row - (rb.first + rb.last) / 2) - Math.abs(b.col - (cb.first + cb.last) / 2));
  const pages: SpreadsheetPageLayoutRenderPlan['pages'] = [];
  let renderedCellCount = 0, cellLimitReached = candidateLimitReached;
  for (const { row, col } of candidates) {
    const remaining = MAX_RENDERED_PAGE_LAYOUT_CELLS_TOTAL - renderedCellCount;
    if (remaining <= 0) { cellLimitReached = true; break; }
    const index = resolveSpreadsheetPageLayoutPageAtTile(geometry, row, col);
    if (index === null) continue;
    const page = resolveSpreadsheetPageLayoutPage(geometry, index, Math.min(remaining, MAX_RENDERED_PAGE_LAYOUT_CELLS_PER_PAGE));
    if (!page) continue;
    pages.push(page); renderedCellCount += page.renderedCellCount;
    cellLimitReached ||= page.cellsTruncated;
  }
  return { pages, renderedCellCount, pageLimitReached: false, cellLimitReached };
}
