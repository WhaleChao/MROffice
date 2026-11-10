/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  NATIVE_TABLE_STYLE_IDS,
  normalizeCellStyle,
  type TableCellBorder,
} from "@mrnote/core/web/lib/presentation-table-style.ts";

export const MIN_INSERT_TABLE_ROWS = 1;
/**
 * Insert > Table > Insert Table...: PowerPoint opens the dialog at 5 columns
 * by 2 rows every time, whatever the grid last inserted (PPT-B001).
 */
export const INSERT_TABLE_DIALOG_DEFAULT_COLUMNS = 5;
export const INSERT_TABLE_DIALOG_DEFAULT_ROWS = 2;
/**
 * Insert Table takes 1 to 75 rows and 1 to 75 columns, PowerPoint's own bounds
 * (PowerPoint 2007 on; it alerts on an entry past 75 in either box). Microsoft
 * Learn's `Shapes.AddTable` and the Support "Add a table to a slide" pages state
 * no number; the 75 x 75 is what PowerPoint's dialog takes. A table grows past
 * it by adding rows and columns, and an imported one of any size reaches the
 * page whole (`MAX_TABLE_SPLIT_ROWS` / `_COLUMNS`).
 */
export const MAX_INSERT_TABLE_ROWS = 75;
export const MIN_INSERT_TABLE_COLUMNS = 1;


export const MAX_INSERT_TABLE_COLUMNS = 75;
/** Every cell of the largest table Insert Table makes (75 x 75). */
export const MAX_INSERT_TABLE_CELLS = MAX_INSERT_TABLE_ROWS * MAX_INSERT_TABLE_COLUMNS;
export const QUICK_INSERT_TABLE_ROWS = 8;
export const QUICK_INSERT_TABLE_COLUMNS = 10;

export type TableInsertDimensionError = "required" | "integer" | "range";

export interface TableInsertDimensionsValidation {
  readonly ok: boolean;
  readonly rows: number | null;
  readonly columns: number | null;
  readonly rowError: TableInsertDimensionError | null;
  readonly columnError: TableInsertDimensionError | null;
  readonly capacityExceeded: boolean;
}

export interface TableInsertGridPosition {
  readonly rows: number;
  readonly columns: number;
}

export interface PresentationTableInsertGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface NativePresentationTableInsert {
  [key: string]: unknown;
  id: string;
  type: "table";
  rows: Array<Array<{ text: string }>>;
  left: number;
  top: number;
  width: number;
  height: number;
  borders: true;
  headerRow: true;
  bandedRows: true;
  tableStyleId: typeof NATIVE_TABLE_STYLE_IDS.mediumAccent1;
}

function finitePositive(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : fallback;
}

function boundedInteger(
  value: unknown,
  minimum: number,
  maximum: number,
): number {
  const parsed =
    typeof value === "number"
      ? value
      : typeof value === "string"
        ? Number.parseInt(value, 10)
        : Number.NaN;
  if (!Number.isFinite(parsed)) return minimum;
  return Math.min(maximum, Math.max(minimum, Math.trunc(parsed)));
}

function strictDimension(
  value: unknown,
  minimum: number,
  maximum: number,
): { value: number | null; error: TableInsertDimensionError | null } {
  if (typeof value === "string" && value.trim() === "") {
    return { value: null, error: "required" };
  }
  const parsed =
    typeof value === "number"
      ? value
      : typeof value === "string" && /^[0-9]+$/.test(value.trim())
        ? Number(value.trim())
        : Number.NaN;
  if (!Number.isSafeInteger(parsed)) {
    return { value: null, error: "integer" };
  }
  if (parsed < minimum || parsed > maximum) {
    return { value: null, error: "range" };
  }
  return { value: parsed, error: null };
}

/**
 * Strict authoring-edge validation for the Insert Table dialog and live writer.
 * Unlike the legacy normalization helper, this never turns malformed user text
 * into a different valid table and enforces the schema's dense-cell budget.
 */
export function validateTableInsertDimensions(
  rows: unknown,
  columns: unknown,
): TableInsertDimensionsValidation {
  const row = strictDimension(
    rows,
    MIN_INSERT_TABLE_ROWS,
    MAX_INSERT_TABLE_ROWS,
  );
  const column = strictDimension(
    columns,
    MIN_INSERT_TABLE_COLUMNS,
    MAX_INSERT_TABLE_COLUMNS,
  );
  const capacityExceeded =
    row.value !== null &&
    column.value !== null &&
    row.value * column.value > MAX_INSERT_TABLE_CELLS;
  return {
    ok: row.value !== null && column.value !== null && !capacityExceeded,
    rows: row.value,
    columns: column.value,
    rowError: row.error,
    columnError: column.error,
    capacityExceeded,
  };
}

function clampQuickTableGridPosition(
  position: TableInsertGridPosition,
): TableInsertGridPosition {
  return {
    rows: Math.max(
      1,
      Math.min(QUICK_INSERT_TABLE_ROWS, Math.trunc(position.rows) || 1),
    ),
    columns: Math.max(
      1,
      Math.min(QUICK_INSERT_TABLE_COLUMNS, Math.trunc(position.columns) || 1),
    ),
  };
}

/** Arrow/Home/End navigation for the 10×8 quick-insert gallery. */
export function tableInsertGridNavigationTarget(
  current: TableInsertGridPosition,
  key: string,
  wholeGrid = false,
): TableInsertGridPosition | null {
  const position = clampQuickTableGridPosition(current);
  switch (key) {
    case "ArrowUp":
      return { ...position, rows: Math.max(1, position.rows - 1) };
    case "ArrowDown":
      return {
        ...position,
        rows: Math.min(QUICK_INSERT_TABLE_ROWS, position.rows + 1),
      };
    case "ArrowLeft":
      return { ...position, columns: Math.max(1, position.columns - 1) };
    case "ArrowRight":
      return {
        ...position,
        columns: Math.min(QUICK_INSERT_TABLE_COLUMNS, position.columns + 1),
      };
    case "Home":
      return wholeGrid ? { rows: 1, columns: 1 } : { ...position, columns: 1 };
    case "End":
      return wholeGrid
        ? {
            rows: QUICK_INSERT_TABLE_ROWS,
            columns: QUICK_INSERT_TABLE_COLUMNS,
          }
        : { ...position, columns: QUICK_INSERT_TABLE_COLUMNS };
    default:
      return null;
  }
}

export function normalizeTableInsertDimensions(
  rows: unknown,
  columns: unknown,
): { rows: number; columns: number } {
  return {
    rows: boundedInteger(rows, MIN_INSERT_TABLE_ROWS, MAX_INSERT_TABLE_ROWS),
    columns: boundedInteger(
      columns,
      MIN_INSERT_TABLE_COLUMNS,
      MAX_INSERT_TABLE_COLUMNS,
    ),
  };
}

export function blankTableRows(
  rowCount: unknown,
  columnCount: unknown,
): Array<Array<{ text: string }>> {
  const dimensions = normalizeTableInsertDimensions(rowCount, columnCount);
  return Array.from({ length: dimensions.rows }, () =>
    Array.from({ length: dimensions.columns }, () => ({ text: "" })),
  );
}

/**
 * PowerPoint/WPS-like default frame for a newly inserted native table.
 *
 * The table is centred within the current slide and scales from a readable
 * two-column minimum up to a bounded 82% × 74% working area. Calculating in
 * slide-space EMU keeps the result independent of browser zoom and preserves
 * the document's native aspect ratio.
 */
export function tableInsertGeometry(
  slideWidth: unknown,
  slideHeight: unknown,
  rowCount: unknown,
  columnCount: unknown,
): PresentationTableInsertGeometry {
  const widthEmu = finitePositive(slideWidth, 9_144_000);
  const heightEmu = finitePositive(slideHeight, 5_143_500);
  const dimensions = normalizeTableInsertDimensions(rowCount, columnCount);

  const width = Math.round(
    Math.min(
      widthEmu * 0.82,
      Math.max(widthEmu * 0.56, dimensions.columns * widthEmu * 0.13),
    ),
  );
  const height = Math.round(
    Math.min(
      heightEmu * 0.74,
      Math.max(heightEmu * 0.18, dimensions.rows * heightEmu * 0.085),
    ),
  );

  return {
    left: Math.round((widthEmu - width) / 2),
    top: Math.round((heightEmu - height) / 2),
    width,
    height,
  };
}

/**
 * Build the editable table block shared by the ribbon picker, legacy insert
 * button, and content-placeholder shortcut. The default mirrors a new WPS
 * table: native Accent 1 style with header and banded-row options enabled.
 */
export function createNativePresentationTableInsert(options: {
  id: string;
  rows: unknown;
  columns: unknown;
  slideWidth: unknown;
  slideHeight: unknown;
  geometry?: PresentationTableInsertGeometry;
}): NativePresentationTableInsert {
  const dimensions = normalizeTableInsertDimensions(
    options.rows,
    options.columns,
  );
  const geometry =
    options.geometry ??
    tableInsertGeometry(
      options.slideWidth,
      options.slideHeight,
      dimensions.rows,
      dimensions.columns,
    );

  return {
    id: options.id,
    type: "table",
    rows: blankTableRows(dimensions.rows, dimensions.columns),
    ...geometry,
    borders: true,
    headerRow: true,
    bandedRows: true,
    tableStyleId: NATIVE_TABLE_STYLE_IDS.mediumAccent1,
  };
}

/** A dragged frame is a real unstyled one-cell table, not a rectangle shape. */
export function createDrawnPresentationTable(options: {
  id: string;
  geometry: PresentationTableInsertGeometry;
  slideWidth: number;
  slideHeight: number;
  pen: TableCellBorder;
}) {
  const { geometry, slideWidth, slideHeight } = options;
  if (
    ![
      geometry.left,
      geometry.top,
      geometry.width,
      geometry.height,
      slideWidth,
      slideHeight,
    ].every(Number.isFinite) ||
    geometry.left < 0 ||
    geometry.top < 0 ||
    geometry.width < 16 * 9525 ||
    geometry.height < 16 * 9525 ||
    geometry.left + geometry.width > slideWidth ||
    geometry.top + geometry.height > slideHeight
  )
    return null;
  const border = normalizeCellStyle({ borderTop: options.pen }).borderTop;
  if (!border || border.style === "none") return null;
  return {
    id: options.id,
    type: "table" as const,
    left: geometry.left,
    top: geometry.top,
    width: geometry.width,
    height: geometry.height,
    tableStyleId: NATIVE_TABLE_STYLE_IDS.noStyleNoGrid,
    borders: false,
    headerRow: false,
    bandedRows: false,
    rows: [
      [
        {
          text: "",
          borderTop: { ...border },
          borderBottom: { ...border },
          borderLeft: { ...border },
          borderRight: { ...border },
        },
      ],
    ],
  };
}
