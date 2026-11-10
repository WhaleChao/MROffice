/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The legacy chart block's data grid — the ONE reading the preview and the
 * PPTX writer share.
 *
 * A legacy `chart` block (no `chartV2`) carries `data.headers` + `data.rows`
 * of strings. The preview (`seriesFromRows` in MiniChart for scatter/bubble,
 * `buildPresentationChartDisplayPlan` for the category kinds) reads that grid
 * by the rules below; `apps/api/app/services/file_adapters/pptx/legacy_chart_grid.py`
 * is the writer's copy of the same rules, and the tests on both sides run the
 * same fixture rows (`mini-chart.test.tsx`, `test_legacy_chart_grid_contract.py`).
 *
 *  - column 0 is the point's label (a category, a scatter/bubble point's
 *    name), never a number;
 *  - for the category kinds every later column is one series;
 *  - for `scatter` column 1 is X and column 2 is Y; for `bubble` column 3 is
 *    the size. There is one series per grid;
 *  - a cell is a number only when `Number(cell)` is finite — a blank, a
 *    non-numeric text or an overflowing/NaN cell is MISSING (`null`), never 0;
 *  - a bubble grid without a size column plots as a scatter, and a scatter
 *    grid without a Y column plots as a column chart.
 *
 * `seriesFromRows` keeps the grid's value columns in `series[0..]`, so the
 * X/Y/size roles are `LEGACY_XY_SERIES` there and `LEGACY_XY_COLUMNS` in the
 * raw row.
 */

/** Column index of each role in a scatter/bubble row: [label, x, y, size]. */
export const LEGACY_XY_COLUMNS = { label: 0, x: 1, y: 2, size: 3 } as const;

/** The same roles as indexes into `seriesFromRows(rows).series` (value columns only). */
export const LEGACY_XY_SERIES = {
  x: LEGACY_XY_COLUMNS.x - 1,
  y: LEGACY_XY_COLUMNS.y - 1,
  size: LEGACY_XY_COLUMNS.size - 1,
} as const;

/** A grid cell as a finite number, or null when it is missing (blank, text, NaN, overflow). */
export function finiteChartNumber(raw: unknown): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "string" && raw.trim() === "") return null;
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

export interface LegacyXYPoint {
  label: string;
  x: number;
  y: number;
  size: number | null;
}

export interface LegacyXYSeries {
  /** What the grid can plot: a bubble grid with no size column is a scatter. */
  kind: "scatter" | "bubble";
  points: LegacyXYPoint[];
  /** Rows with a missing coordinate; the preview plots them at 0, the file leaves them out. */
  droppedRows: number;
}

/**
 * Read a scatter/bubble grid (data rows only, no header row) as
 * [label, x, y(, size)] points — the reading `renderScatter` / `renderBubble`
 * draw and the PPTX writer exports. Null when there is no Y column at all.
 */
export function legacyXYSeriesFromRows(
  rows: readonly (readonly string[])[] | undefined,
  kind: "scatter" | "bubble",
): LegacyXYSeries | null {
  if (!Array.isArray(rows)) return null;
  const usable = rows.filter((row): row is string[] => Array.isArray(row));
  const width = usable.reduce((max, row) => Math.max(max, row.length), 0);
  if (width <= LEGACY_XY_COLUMNS.y) return null;
  const actual: "scatter" | "bubble" =
    kind === "bubble" && width > LEGACY_XY_COLUMNS.size ? "bubble" : "scatter";
  const points: LegacyXYPoint[] = [];
  let droppedRows = 0;
  usable.forEach((row) => {
    const x = finiteChartNumber(row[LEGACY_XY_COLUMNS.x]);
    const y = finiteChartNumber(row[LEGACY_XY_COLUMNS.y]);
    const size =
      actual === "bubble" ? finiteChartNumber(row[LEGACY_XY_COLUMNS.size]) : null;
    if (x === null || y === null || (actual === "bubble" && size === null)) {
      droppedRows += 1;
      return;
    }
    points.push({ label: String(row[LEGACY_XY_COLUMNS.label] ?? ""), x, y, size });
  });
  return { kind: actual, points, droppedRows };
}
