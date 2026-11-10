/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeChartSeriesRefs, parseChartRefAreas, type ParsedChartRefRange as ParsedArea } from "./chart";

export type ChartVisibility = { state: "known"; plotVisOnly: boolean } | { state: "absent" | "invalid" };
export function normalizeChartVisibility(raw: unknown): ChartVisibility | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (raw && typeof raw === "object" && !Array.isArray(raw)) {
    const value = raw as Record<string, unknown>;
    if (Object.keys(value).length === 2 && value.state === "known" && typeof value.plotVisOnly === "boolean") return { state: "known", plotVisOnly: value.plotVisOnly };
    if (Object.keys(value).length === 1 && (value.state === "absent" || value.state === "invalid")) return { state: value.state };
  }
  return { state: "invalid" };
}

type Sheet = { name: string; manuallyHiddenRows?: readonly number[]; manuallyHiddenCols?: readonly number[]; outlineGroups?: readonly { axis: string; start: number; end: number; collapsed?: boolean }[]; advancedFilterState?: { hiddenRows: readonly number[] }; filters?: unknown; colorFilters?: unknown; conditionFilters?: unknown; filterState?: unknown; dataTables?: readonly { columnFilters?: unknown }[] };
export type ChartVisibilityProjection = { state: "resolved"; sourcePointIndices: number[]; sourceSeriesIndices: number[]; implicitCategories?: true } | { state: "unavailable" | "unsupported"; reason: string };
const nonempty = (value: unknown) => value != null && (typeof value !== "object" || Object.keys(value).length > 0);
function hidden(sheet: Sheet, row: number, col: number): boolean {
  return Boolean(sheet.manuallyHiddenRows?.includes(row) || sheet.manuallyHiddenCols?.includes(col) || sheet.advancedFilterState?.hiddenRows.includes(row) || sheet.outlineGroups?.some(group => group.collapsed === true && (group.axis === "row" ? row : col) >= group.start && (group.axis === "row" ? row : col) <= group.end));
}

const sheetFiltered = (sheet: Sheet) =>
  [sheet.filters, sheet.colorFilters, sheet.conditionFilters, sheet.filterState].some(nonempty) ||
  Boolean(sheet.dataTables?.some(table => nonempty(table.columnFilters)));

/** Only proven same-sheet, same-direction aligned category masks. The caller
 * resolves values once; this function never reparses formulas or edits cells.
 * `filterHiddenRows` is the rows a sheet's filters hide as the grid computed
 * them, when the caller has them: a filter that hides no row of the chart's
 * data then changes nothing (user audit S-018 blanked a chart over a table
 * filtered far below it). A filtered sheet without them stays unresolved. */
export function chartVisiblePointIndices(chart: { chartVisibility?: unknown; chartType?: unknown; type?: unknown; seriesRefs?: unknown; secondaryAxisSeries?: unknown; sliceStyles?: unknown; chartFormat?: unknown }, sheets: readonly Sheet[], count: number, filterHiddenRows?: (sheetName: string) => ReadonlySet<number> | undefined): ChartVisibilityProjection {
  const policy = normalizeChartVisibility(chart.chartVisibility);
  const refs = normalizeChartSeriesRefs(chart.seriesRefs);
  const resolved: ChartVisibilityProjection = { state: "resolved", sourcePointIndices: Array.from({ length: count }, (_, index) => index), sourceSeriesIndices: (refs ?? []).map((_, index) => index) };
  if (policy?.state === "known" && !policy.plotVisOnly) return resolved;
  if (!refs?.length) return { state: "unavailable", reason: "chart_visibility_reference_unavailable" };
  const vectors: { sheet: Sheet; r1: number; r2: number; c1: number; c2: number; filtered?: ReadonlySet<number> }[] = [];
  for (const ref of refs) {
    for (const formula of [ref.categories, ref.values, ref.xValues, ref.bubbleSizes]) {
      if (!formula) continue;
      const areas = parseChartRefAreas(formula);
      const ownerOf = (area: { sheetName?: string }) => sheets.filter(sheet => sheet.name.toLowerCase() === area.sheetName?.toLowerCase());
      if (!areas || areas.some(area => ownerOf(area).length !== 1)) return { state: "unavailable", reason: "chart_visibility_reference_unavailable" };
      // A union of areas (a chart over a non-contiguous selection) plots each
      // area's cells in turn. Like a block, it only has to prove that no cell
      // of it hides; one that does stays unresolved.
      const blockHides = (area: ParsedArea) => {
        const sheet = ownerOf(area)[0];
        const filtered = sheetFiltered(sheet) ? filterHiddenRows?.(sheet.name) : undefined;
        if (sheetFiltered(sheet) && !filtered) return true;
        const { r1, r2, c1, c2 } = area.range;
        const hiddenAt = (row: number, col: number) => hidden(sheet, row, col) || Boolean(filtered?.has(row));
        for (let row = r1; row <= r2; row += 1) if (hiddenAt(row, c1)) return true;
        for (let col = c1; col <= c2; col += 1) if (hiddenAt(r1, col)) return true;
        return false;
      };
      if (areas.length > 1) {
        if (areas.some(blockHides)) return { state: "unavailable", reason: "chart_visibility_reference_unavailable" };
        continue;
      }
      const parsed = areas[0];
      const { r1, r2, c1, c2 } = parsed.range, sheet = ownerOf(parsed)[0];
      const filtered = sheetFiltered(sheet) ? filterHiddenRows?.(sheet.name) : undefined;
      const hiddenAt = (row: number, col: number) => hidden(sheet, row, col) || Boolean(filtered?.has(row));
      if (r1 !== r2 && c1 !== c2) {
        // A treemap's or sunburst's category levels span rows and columns at
        // once. A block with every cell visible cannot change which points are
        // plotted, so it only has to prove that; a filtered sheet or a hidden
        // row or column inside it stays unresolved. A cell hides only through
        // its row or its column, so one pass down the first column and one
        // across the first row reach every hidden cell of the block.
        let blockHidden = sheetFiltered(sheet) && !filtered;
        for (let row = r1; row <= r2 && !blockHidden; row += 1) blockHidden = hiddenAt(row, c1);
        for (let col = c1; col <= c2 && !blockHidden; col += 1) blockHidden = hiddenAt(r1, col);
        if (blockHidden) return { state: "unavailable", reason: "chart_visibility_reference_unavailable" };
        continue;
      }
      if (sheetFiltered(sheet) && !filtered) return { state: "unsupported", reason: "chart_visibility_filter_unsupported" };
      vectors.push({ sheet, r1, r2, c1, c2, filtered });
    }
  }
  const masks = vectors.map(vector => Array.from({ length: Math.max(vector.r2-vector.r1, vector.c2-vector.c1)+1 }, (_, index) => {
    const row = vector.r1 + (vector.c1 === vector.c2 ? index : 0);
    return hidden(vector.sheet, row, vector.c1 + (vector.c1 === vector.c2 ? 0 : index)) || Boolean(vector.filtered?.has(row));
  }));
  // With all referenced source cells visible, either policy gives exactly the
  // existing projection. No unknown flag is converted into a stored boolean.
  if (!masks.some(mask => mask.some(Boolean))) return resolved;
  if (!policy || policy.state !== "known") return { state: "unavailable", reason: "chart_visibility_unknown" };
  const format = chart.chartFormat && typeof chart.chartFormat === "object" ? chart.chartFormat as Record<string, unknown> : undefined;
  if (!["column", "bar"].includes(String(chart.chartType ?? chart.type)) || nonempty(chart.secondaryAxisSeries) || nonempty(chart.sliceStyles) || nonempty(format?.sliceStyles) || refs.some(ref => ref.axis === "secondary" || ref.xValues || ref.implicitX || ref.bubbleSizes)) return { state: "unsupported", reason: "chart_visibility_geometry_unsupported" };
  // A chart whose series name no categories plots its points as 1, 2, ...:
  // Excel numbers the visible ones afresh (C09_plotVisOnly: rows 6-7 hidden,
  // the categories read 1, 2; X-103). Some series with categories and some
  // without stay unresolved.
  const implicitCategories = refs.every(ref => !ref.categories);
  if (!implicitCategories && refs.some(ref => !ref.categories)) return { state: "unsupported", reason: "chart_visibility_category_reference_missing" };
  const first = vectors[0], vertical = first.c1 === first.c2;
  if (vectors.some(vector => vector.sheet !== first.sheet || (vector.c1 === vector.c2) !== vertical || Math.max(vector.r2-vector.r1, vector.c2-vector.c1)+1 !== count || (vertical ? vector.r1 !== first.r1 || vector.r2 !== first.r2 : vector.c1 !== first.c1 || vector.c2 !== first.c2))) return { state: "unsupported", reason: "chart_visibility_alignment_unsupported" };
  if (masks.some(mask => mask.some((value, index) => value !== masks[0][index]))) return { state: "unsupported", reason: "chart_visibility_masks_disagree" };
  resolved.sourcePointIndices = resolved.sourcePointIndices.filter(index => !masks[0][index]);
  if (implicitCategories) resolved.implicitCategories = true;
  return resolved;
}

export function projectChartVisibilityRows(rows: string[][], projection: ChartVisibilityProjection): string[][] {
  if (projection.state !== "resolved" || !rows.length) return [];
  return [
    rows[0],
    ...projection.sourcePointIndices.map((index, place) =>
      projection.implicitCategories && rows[index + 1] ? [String(place + 1), ...rows[index + 1].slice(1)] : rows[index + 1],
    ),
  ];
}
