/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { parseBlankCellRangeRef } from "@mrnote/core/web/lib/spreadsheet/blank-cell-formats.ts";
import {
  clampColumnWidth,
  clampRowHeight,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
  MIN_COLUMN_WIDTH_PX,
  MIN_ROW_HEIGHT_PX,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { nativeRowPixelPairMatches, normalizeNativeGrid } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";

type UnknownRecord = Record<string, unknown>;

function isRecord(value: unknown): value is UnknownRecord {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isGridCount(value: unknown, ceiling: number): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= ceiling
  );
}















function gridGrownToStoredCells(
  sheet: UnknownRecord,
): { rowCount: number; columnCount: number } | null {
  const { rowCount, columnCount } = sheet;
  if (
    !isGridCount(rowCount, MAX_GRID_ROWS) ||
    !isGridCount(columnCount, MAX_GRID_COLUMNS)
  ) {
    return null;
  }
  let rows = rowCount;
  let columns = columnCount;
  const cover = (lastRow: number, lastColumn: number) => {
    if (lastRow >= MAX_GRID_ROWS || lastColumn >= MAX_GRID_COLUMNS) return;
    rows = Math.max(rows, lastRow + 1);
    columns = Math.max(columns, lastColumn + 1);
  };
  if (isRecord(sheet.cells)) {
    for (const key of Object.keys(sheet.cells)) {
      const match = /^(\d+):(\d+)$/.exec(key);
      if (match) cover(Number(match[1]), Number(match[2]));
    }
  }
  if (Array.isArray(sheet.blankCellFormats)) {
    for (const group of sheet.blankCellFormats) {
      const ranges = isRecord(group) && Array.isArray(group.ranges) ? group.ranges : [];
      for (const ref of ranges) {
        const rect = parseBlankCellRangeRef(ref);
        if (rect) cover(rect.r2, rect.c2);
      }
    }
  }
  return rows === rowCount && columns === columnCount
    ? null
    : { rowCount: rows, columnCount: columns };
}

function normalizeFiniteSizeMap(
  value: unknown,
  clamp: (value: unknown, fallback: number) => number,
  fallback: number,
  preserve?: (index: string, pixels: number) => boolean,
): { value: unknown; changed: boolean } {
  if (!isRecord(value)) return { value, changed: false };
  let changed = false;
  const normalized: UnknownRecord = {};
  for (const [key, rawSize] of Object.entries(value)) {
    if (typeof rawSize !== "number" || !Number.isFinite(rawSize)) {
      normalized[key] = rawSize;
      continue;
    }
    const nextSize = preserve?.(key, rawSize) ? rawSize : clamp(rawSize, fallback);
    normalized[key] = nextSize;
    if (nextSize !== rawSize) changed = true;
  }
  return { value: changed ? normalized : value, changed };
}

/** A row height the page stores that {@link normalizeSpreadsheetImportDimensions} shows otherwise. */
export interface ImportRowHeightShownAs {
  sheetIndex: number;
  row: string;
  stored: number;
  shown: number;
}








export function importRowHeightsShownAs(
  content: Record<string, unknown> | undefined,
): ImportRowHeightShownAs[] {
  if (!content || !Array.isArray(content.sheets)) return [];
  const out: ImportRowHeightShownAs[] = [];
  content.sheets.forEach((rawSheet, sheetIndex) => {
    if (!isRecord(rawSheet) || !isRecord(rawSheet.rowHeights)) return;
    const stored = rawSheet.rowHeights;
    const nativeRows = normalizeNativeGrid(rawSheet.nativeGrid)?.rowMetrics;
    const rows = normalizeFiniteSizeMap(
      stored,
      clampRowHeight,
      MIN_ROW_HEIGHT_PX,
      (index, pixels) => nativeRowPixelPairMatches(nativeRows, index, pixels),
    );
    if (!rows.changed || !isRecord(rows.value)) return;
    for (const [row, size] of Object.entries(stored)) {
      const shown = rows.value[row];
      if (typeof size === "number" && typeof shown === "number" && shown !== size) {
        out.push({ sheetIndex, row, stored: size, shown });
      }
    }
  });
  return out;
}

/**
 * Normalize imported Excel/WPS line sizes before the strict page schema gate.
 *
 * Legacy sizes use the interactive bounds. Paired native row dimensions retain
 * their source projection so the renderer can resolve exact point heights.
 * Normalizing the remaining sizes before schema validation avoids
 * an entire workbook being replaced by an empty sheet when Office emits a
 * legitimate compact row (for example a 9pt/12px header) below the web grid's
 * 18px interactive minimum. Unknown or non-numeric values remain untouched so
 * malformed payloads still fail the schema rather than being silently hidden.
 *
 * The declared rowCount/columnCount also grow to hold every stored cell (see
 * gridGrownToStoredCells), so a merged workbook opens editable and the next
 * save writes the grown grid.
 */
export function normalizeSpreadsheetImportDimensions(
  content: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  if (!content || !Array.isArray(content.sheets)) return content;

  let changed = false;
  const sheets = content.sheets.map((rawSheet) => {
    if (!isRecord(rawSheet)) return rawSheet;
    const nativeRows = normalizeNativeGrid(rawSheet.nativeGrid)?.rowMetrics;
    const columns = normalizeFiniteSizeMap(
      rawSheet.columnWidths,
      clampColumnWidth,
      MIN_COLUMN_WIDTH_PX,
    );
    const rows = normalizeFiniteSizeMap(
      rawSheet.rowHeights,
      clampRowHeight,
      MIN_ROW_HEIGHT_PX,
      (index, pixels) => nativeRowPixelPairMatches(nativeRows, index, pixels),
    );
    const grid = gridGrownToStoredCells(rawSheet);
    if (!columns.changed && !rows.changed && !grid) return rawSheet;
    changed = true;
    return {
      ...rawSheet,
      ...(columns.changed ? { columnWidths: columns.value } : {}),
      ...(rows.changed ? { rowHeights: rows.value } : {}),
      ...(grid ?? {}),
    };
  });

  return changed ? { ...content, sheets } : content;
}
