/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The worksheet's A1 range helpers for a chart's data source, and the reading
 * of an older imported chart object's display toggles (`normalizeChartConfig`,
 * what the object layer draws an unmodelled imported chart with). Chart edits
 * are the chart model's (lib/spreadsheet/chart-v2-model); the former object
 * editor these helpers served is gone.
 *
 * Everything here is defensive: a persisted value can come from an imported
 * .xlsx, an AI write, or a hand-edited JSON, so each normalizer clamps / drops
 * out-of-range input rather than trusting it.
 */

import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import { columnToLetters, lettersToColumn } from "./formula-engine";

// MiniChart only renders these three kinds; every other label normalizes to one
// of them (so an imported "column"/"area"/"donut" still draws something).
export type ChartType = "bar" | "line" | "pie";

// Bounds mirror the schema (spreadsheetChartObjectSchema / the grid clamps in
// SpreadsheetWindow) so a normalized range can never address outside the grid.
// The last cell of Excel's grid: a chart may read any rows of a sheet (a
// large workbook's reach far past 10,000).
const MAX_RANGE_ROW = MAX_GRID_ROWS - 1;
const MAX_RANGE_COL = MAX_GRID_COLUMNS - 1;
const MAX_TITLE_LENGTH = 2_000;

/** A 0-based inclusive cell rectangle, normalized so r1<=r2 and c1<=c2. */
export interface ChartRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * The additive config persisted on a chart object. Every field is optional so
 * existing chart objects (which never had these) still validate; the editor and
 * MiniChart fall back to sensible defaults when a field is absent.
 */
export interface ChartConfig {
  chartType: ChartType;
  title: string;
  range: string | null;
  showLegend: boolean;
  showAxisLabels: boolean;
  /** Where the legend sits (Excel "Legend position"). Default "top". */
  legendPosition: "top" | "bottom" | "left" | "right" | "topRight";
}

/**
 * Normalize a free-form chart-type label to one of MiniChart's three kinds.
 * Matches MiniChart's own mapping: line/area -> "line", pie/donut/doughnut ->
 * "pie", everything else (bar/column/unknown/empty) -> "bar".
 */
export function normalizeChartType(type: unknown): ChartType {
  const t = typeof type === "string" ? type.trim().toLowerCase() : "";
  if (t.includes("line") || t.includes("area")) return "line";
  if (t.includes("pie") || t.includes("donut") || t.includes("doughnut")) {
    return "pie";
  }
  return "bar";
}

/** Clamp an integer index into [0, max]; non-finite -> 0. */
function clampIndex(value: number, max: number): number {
  if (!Number.isFinite(value)) return 0;
  const n = Math.trunc(value);
  if (n < 0) return 0;
  if (n > max) return max;
  return n;
}

// A single A1 cell, e.g. "A1" / "$B$3". Letters 1..3, digits 1..7 — matches the
// formula engine's CELL_REF_RE so the two stay consistent.
const A1_CELL_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;

function parseA1Cell(token: string): { row: number; col: number } | null {
  const m = A1_CELL_RE.exec(token.trim());
  if (!m) return null;
  const col = lettersToColumn(m[1]);
  const row = Number(m[2]) - 1;
  if (col < 0 || row < 0 || !Number.isFinite(row)) return null;
  return { row, col };
}

/**
 * Parse + validate an A1-style range ("A1:B10", a single "A1", or already a
 * {r1,c1,r2,c2} object) into a normalized, in-grid ChartRange. Indices are
 * clamped to the grid and the corners are sorted so r1<=r2 / c1<=c2. Returns
 * null for anything that is not a recognizable range.
 */
export function normalizeChartRange(input: unknown): ChartRange | null {
  // Already a rectangle object (e.g. round-tripped from a previous normalize).
  if (input && typeof input === "object" && !Array.isArray(input)) {
    const o = input as Record<string, unknown>;
    if (
      typeof o.r1 === "number" &&
      typeof o.c1 === "number" &&
      typeof o.r2 === "number" &&
      typeof o.c2 === "number"
    ) {
      const r1 = clampIndex(o.r1, MAX_RANGE_ROW);
      const r2 = clampIndex(o.r2, MAX_RANGE_ROW);
      const c1 = clampIndex(o.c1, MAX_RANGE_COL);
      const c2 = clampIndex(o.c2, MAX_RANGE_COL);
      return {
        r1: Math.min(r1, r2),
        c1: Math.min(c1, c2),
        r2: Math.max(r1, r2),
        c2: Math.max(c1, c2),
      };
    }
    return null;
  }

  if (typeof input !== "string") return null;
  const trimmed = input.trim().replace(/^=/, "");
  if (!trimmed) return null;
  // The cells after a sheet qualifier ("Sheet1!A1:B2"); the sheet itself is
  // chartRangeSheetName's, for the caller to resolve.
  const bang = trimmed.lastIndexOf("!");
  const body = bang >= 0 ? trimmed.slice(bang + 1) : trimmed;

  const parts = body.split(":");
  if (parts.length === 1) {
    const cell = parseA1Cell(parts[0]);
    if (!cell) return null;
    const row = clampIndex(cell.row, MAX_RANGE_ROW);
    const col = clampIndex(cell.col, MAX_RANGE_COL);
    return { r1: row, c1: col, r2: row, c2: col };
  }
  if (parts.length !== 2) return null;
  const a = parseA1Cell(parts[0]);
  const b = parseA1Cell(parts[1]);
  if (!a || !b) return null;
  const r1 = clampIndex(a.row, MAX_RANGE_ROW);
  const r2 = clampIndex(b.row, MAX_RANGE_ROW);
  const c1 = clampIndex(a.col, MAX_RANGE_COL);
  const c2 = clampIndex(b.col, MAX_RANGE_COL);
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

/**
 * The worksheet a typed data range names ("'Sales 2024'!$A$1:$D$4" ->
 * "Sales 2024", quotes and doubled apostrophes undone), or undefined when it
 * names none. normalizeChartRange reads only the cells; whoever applies the
 * range must read them on this sheet, or refuse a sheet the workbook does not
 * have, never read them on its own (user audit S-012).
 */
export function chartRangeSheetName(input: unknown): string | undefined {
  if (typeof input !== "string") return undefined;
  const trimmed = input.trim().replace(/^=/, "");
  const bang = trimmed.lastIndexOf("!");
  if (bang < 0) return undefined;
  const raw = trimmed.slice(0, bang).trim();
  const quoted = /^'(.*)'$/s.exec(raw);
  const name = quoted ? quoted[1].replace(/''/g, "'") : raw;
  return name;
}

/** Render a normalized ChartRange back to its A1 string ("A1:B10" or "A1"). */
export function chartRangeToA1(range: ChartRange | null): string {
  if (!range) return "";
  const from = `${columnToLetters(range.c1)}${range.r1 + 1}`;
  if (range.r1 === range.r2 && range.c1 === range.c2) return from;
  const to = `${columnToLetters(range.c2)}${range.r2 + 1}`;
  return `${from}:${to}`;
}

/** A chart's data source as the user sees it: "#REF!" once its source lines
 * were deleted (Excel's broken series reference), otherwise its A1 range. */
export function chartSourceRangeToA1(chart: {
  range: ChartRange | null;
  sourceRefError?: boolean;
}): string {
  return chart.sourceRefError ? "#REF!" : chartRangeToA1(chart.range);
}

/**
 * Read the chart config off a persisted chart object, applying defaults for any
 * additive field that is absent (so an old chart with only chartType/title still
 * yields a complete config). `range` is normalized to its canonical A1 string
 * (or null when there is no valid range). `ranges` is the legacy plural array;
 * the first usable entry seeds the range when the singular field is missing.
 */
export function normalizeChartConfig(input: unknown): ChartConfig {
  const o =
    input && typeof input === "object" ? (input as Record<string, unknown>) : {};

  const chartType = normalizeChartType(o.chartType);

  const rawTitle =
    typeof o.title === "string"
      ? o.title
      : typeof o.caption === "string"
        ? o.caption
        : "";
  const title = rawTitle.slice(0, MAX_TITLE_LENGTH);

  // Prefer the singular `range`; fall back to the first entry of the legacy
  // `ranges` array. Either is run through normalize -> canonical A1 string.
  let range: string | null = null;
  const fromSingular = normalizeChartRange(o.range);
  if (fromSingular) {
    range = chartRangeToA1(fromSingular);
  } else if (Array.isArray(o.ranges)) {
    for (const candidate of o.ranges) {
      const parsed = normalizeChartRange(candidate);
      if (parsed) {
        range = chartRangeToA1(parsed);
        break;
      }
    }
  }

  // Default both toggles ON (Excel shows legend + axis labels by default); only
  // an explicit `false` turns them off.
  const showLegend = o.showLegend === undefined ? true : o.showLegend !== false;
  const showAxisLabels =
    o.showAxisLabels === undefined ? true : o.showAxisLabels !== false;
  // Legend position: "bottom" only when explicitly set, else "top" (the prior
  // fixed placement), so an existing chart config round-trips unchanged.
  const legendPosition =
    o.legendPosition === "bottom" ||
    o.legendPosition === "left" ||
    o.legendPosition === "right" ||
    o.legendPosition === "topRight"
      ? o.legendPosition
      : "top";

  return { chartType, title, range, showLegend, showAxisLabels, legendPosition };
}
