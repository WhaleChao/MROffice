/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { nearestKnownCellMap } from "@mrnote/core/web/lib/spreadsheet/cell-map-lineage.ts";
import type { ColumnWidthReading } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import {
  columnPixelsToCharacters,
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  roundSizeReading,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { normalizeManualHidden } from "@mrnote/core/web/lib/spreadsheet/manual-hidden.ts";
import { hiddenColsFromGroups, normalizeOutlineGroups } from "@mrnote/core/web/lib/spreadsheet/outline-groups.ts";
import { normalizeNativeGrid } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { resolveSpreadsheetNativeGrid } from "./row-column-buttons";

/** The sheet fields a column width reading comes from. */
export interface ColumnWidthSheet {
  rowCount?: unknown;
  columnCount?: unknown;
  columnWidths?: unknown;
  defaultColumnWidth?: unknown;
  nativeGrid?: unknown;
  manuallyHiddenCols?: unknown;
  outlineGroups?: unknown;
}

/**
 * A sheet's column widths in characters, as its Column Width dialog reads
 * them (SpreadsheetWindow's `columnWidthReading`): the pixels the grid draws,
 * through the imported workbook's own digit width and DPI, or Excel's 7-pixel
 * digit for a sheet made here; a hidden column is 0. CELL("width") rounds
 * them (X-207). The workbook theme's replacement of the grid font is not
 * applied: a sheet re-measured for a new theme font reads its file widths.
 */
export function sheetColumnWidthReading(sheet: ColumnWidthSheet): ColumnWidthReading {
  const widths =
    sheet.columnWidths && typeof sheet.columnWidths === "object"
      ? (sheet.columnWidths as Record<string, number>)
      : undefined;
  const defaultWidth = typeof sheet.defaultColumnWidth === "number" ? sheet.defaultColumnWidth : undefined;
  const grid = resolveSpreadsheetNativeGrid(normalizeNativeGrid(sheet.nativeGrid), widths, defaultWidth);
  const characters = (pixels: number) =>
    roundSizeReading(grid ? columnPixelsToCharacters(pixels, grid.digitWidth, grid.dpi) : columnPixelsToCharacters(pixels));
  const columns = new Map<number, number>();
  for (const [key, pixels] of Object.entries(grid?.widths ?? widths ?? {})) {
    const col = Number(key);
    if (Number.isInteger(col) && col >= 0 && typeof pixels === "number" && Number.isFinite(pixels)) {
      columns.set(col, characters(pixels));
    }
  }
  const columnCount = Number(sheet.columnCount) || 0;
  const hidden = [
    ...normalizeManualHidden(sheet.manuallyHiddenCols, columnCount),
    ...hiddenColsFromGroups(
      normalizeOutlineGroups(sheet.outlineGroups, { rowCount: Number(sheet.rowCount) || 0, columnCount }),
    ),
  ];
  for (const col of hidden) columns.set(col, 0);
  return {
    defaultWidth: characters(grid?.defaultWidth ?? defaultWidth ?? DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX),
    columns: [...columns],
  };
}

const callsCell = new WeakMap<object, boolean>();

/**
 * Whether any formula of these sheets calls CELL( -- the one reader of
 * column widths, which a live workbook only carries when asked: a width
 * otherwise changes nothing a formula reads, and carried along, every
 * column drag would recalculate the workbook. Remembered per cell map.
 */
export function workbookCallsCell(sheets: ReadonlyArray<{ cells: Readonly<Record<string, { value?: unknown } | undefined>> }>): boolean {
  return sheets.some((sheet) => {
    const cached = callsCell.get(sheet.cells);
    if (cached !== undefined) return cached;
    const calls = (key: string) => {
      const value = sheet.cells[key]?.value;
      return typeof value === "string" && value.charCodeAt(0) === 61 && /CELL\s*\(/i.test(value);
    };
    // A map that did not call it, with these keys changed, calls it only there.
    const prior = nearestKnownCellMap(sheet.cells, (before) => callsCell.get(before));
    let found = false;
    if (prior?.value === false) found = prior.keys.some(calls);
    else for (const key in sheet.cells) if (calls(key)) { found = true; break; }
    callsCell.set(sheet.cells, found);
    return found;
  });
}
