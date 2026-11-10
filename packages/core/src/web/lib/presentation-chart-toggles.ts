/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure resolution of a presentation chart block's display toggles — the legend
 * and category axis-labels switches (PowerPoint / WPS parity, mirroring the
 * spreadsheet chart-object's `normalizeChartConfig`).
 *
 * The presentation chart block historically carried only `type` + `rows`, so its
 * `<MiniChart>` rendered with the legend AND axis labels OFF (MiniChart defaults
 * both to `false`) — strictly worse than the spreadsheet chart object, which
 * persists `showLegend`/`showAxisLabels` and defaults them ON (Excel shows both
 * by default). These helpers bring the block to parity:
 *
 *  - {@link resolveChartToggles} reads a (possibly partial / hand-edited) block
 *    and yields the two booleans the call sites pass to `<MiniChart>`. Like the
 *    spreadsheet normalizer, BOTH default ON; only an explicit `false` turns one
 *    off (so an old block with neither key shows legend + axis labels, matching
 *    PowerPoint / the spreadsheet object).
 *  - {@link applyChartToggle} produces the patch the inspector checkbox writes:
 *    it persists ONLY a value that DIFFERS from the ON default. Setting a toggle
 *    back to its default (`true`) DROPS the key, so a chart that has never been
 *    customized serializes byte-for-byte as before (additive + migration-safe,
 *    the same discipline `presentation-text-columns` / the chart-config writer
 *    use elsewhere). The `false` (off) value is the only one persisted.
 *
 * Dependency-free (no React / no DOM) so the logic is unit-testable in isolation
 * and shared by both `<MiniChart>` call sites + the inspector.
 */
import {
  chartSupportsDataTable as sharedChartSupportsDataTable,
  chartSupportsGridlines as sharedChartSupportsGridlines,
  chartSupportsTrendline,
} from "../components/console/editor/extensions/chartBlockShared";


/** The two chart display toggles a presentation chart block can carry. */
export const CHART_TOGGLE_KEYS = ["showLegend", "showAxisLabels"] as const;
export type ChartToggleKey = (typeof CHART_TOGGLE_KEYS)[number];

export const PRESENTATION_CHART_LEGEND_POSITIONS = [
  "right",
  "left",
  "top",
  "bottom",
] as const;
export type PresentationChartLegendPosition =
  (typeof PRESENTATION_CHART_LEGEND_POSITIONS)[number];

export interface PresentationChartLinearTrendline {
  type: "linear";
}

/** Resolved booleans for the two toggles (both default ON). */
export interface ResolvedChartToggles {
  showLegend: boolean;
  showAxisLabels: boolean;
}

/**
 * The bounded Quick Layout subset Presentation can round-trip today. Each
 * preset is an exact combination of the two native chart-element toggles; no
 * title, data-label, gridline, data-table or style authority is implied.
 */
export const PRESENTATION_CHART_QUICK_LAYOUT_IDS = [
  "both",
  "legend-only",
  "axis-only",
  "minimal",
] as const;

export type PresentationChartQuickLayoutId =
  (typeof PRESENTATION_CHART_QUICK_LAYOUT_IDS)[number];

const PRESENTATION_CHART_QUICK_LAYOUTS = Object.freeze({
  both: Object.freeze({ showLegend: true, showAxisLabels: true }),
  "legend-only": Object.freeze({
    showLegend: true,
    showAxisLabels: false,
  }),
  "axis-only": Object.freeze({
    showLegend: false,
    showAxisLabels: true,
  }),
  minimal: Object.freeze({ showLegend: false, showAxisLabels: false }),
} satisfies Record<PresentationChartQuickLayoutId, ResolvedChartToggles>);

export function presentationChartQuickLayout(
  id: PresentationChartQuickLayoutId,
): ResolvedChartToggles {
  return PRESENTATION_CHART_QUICK_LAYOUTS[id];
}

/**
 * Read one toggle off a loose record, defaulting ON: only a strict `false`
 * (the only value the writer persists) turns it off; `undefined`, a missing key,
 * or any non-`false` value resolves to `true`. Mirrors `normalizeChartConfig`'s
 * `o.showLegend === undefined ? true : o.showLegend !== false` rule exactly.
 */
function readToggle(value: unknown): boolean {
  return value !== false;
}

/**
 * Resolve BOTH chart display toggles off a (possibly partial / hand-edited) chart
 * block. Both default ON (Excel / spreadsheet-object parity); only an explicit
 * `false` turns one off. A non-object input yields both ON.
 */
export function resolveChartToggles(block: unknown): ResolvedChartToggles {
  const o =
    block && typeof block === "object"
      ? (block as Record<string, unknown>)
      : {};
  return {
    showLegend: readToggle(o.showLegend),
    showAxisLabels: readToggle(o.showAxisLabels),
  };
}

/** Resolve every valid toggle combination back to its stable gallery id. */
export function presentationChartQuickLayoutId(
  block: unknown,
): PresentationChartQuickLayoutId {
  const toggles = resolveChartToggles(block);
  return PRESENTATION_CHART_QUICK_LAYOUT_IDS.find((id) => {
    const layout = presentationChartQuickLayout(id);
    return (
      layout.showLegend === toggles.showLegend &&
      layout.showAxisLabels === toggles.showAxisLabels
    );
  })!;
}

/**
 * Build the minimal patch for setting `key` to `next` on a chart block, applied
 * over the EXISTING block so the result is a complete block (the caller spreads
 * it through its single mutation entry point). Byte-clean: the ON default
 * (`true`) DROPS the key (so an un-customized chart stays byte-for-byte as
 * before); only `false` (off) is persisted. The input block is NOT mutated.
 */
export function applyChartToggle<T extends Record<string, unknown>>(
  block: T,
  key: ChartToggleKey,
  next: boolean,
): T {
  const updated: Record<string, unknown> = { ...block };
  if (next) {
    // ON is the default — persist nothing (drop any stored value) so the chart
    // serializes byte-for-byte as before it was ever toggled.
    delete updated[key];
  } else {
    updated[key] = false;
    // Native OOXML has no legend node (and therefore no placement) while the
    // legend is hidden. Drop the dependent atom rather than preserving a UI-
    // only position that cannot survive a PowerPoint round trip.
    if (key === "showLegend") delete updated.legendPosition;
  }
  return updated as T;
}

/** PowerPoint data labels default OFF; only a strict true opts the chart in. */
export function resolveChartDataLabels(block: unknown): boolean {
  return Boolean(
    block &&
      typeof block === "object" &&
      (block as Record<string, unknown>).showDataLabels === true,
  );
}

/** Persist the non-default ON atom and omit the default OFF atom. */
export function applyChartDataLabels<T extends Record<string, unknown>>(
  block: T,
  visible: boolean,
): T {
  const updated: Record<string, unknown> = { ...block };
  if (visible) updated.showDataLabels = true;
  else {
    delete updated.showDataLabels;
    delete updated.dataLabelStyle;
  }
  return updated as T;
}

/**
 * The family and subtype a deck chart draws as, when the block says so: a
 * Chart V2 block states them, and an imported chart carries the reading the
 * shared Office reader took from its own part. The three gates below used to
 * pattern-match the legacy `chartType` word alone — "bar", "line", "area" —
 * which cannot tell a stacked or a 3-D chart from a plain one, a bubble from a
 * scatter, or a radar from a line, so the ribbon offered a trendline, a data
 * table and gridlines the save then gave up (measured on the 60-chart
 * PowerPoint corpus: 41 trendline picks and 27 data-table picks declined).
 * With a leaf in hand the shared contract answers, the same answer Docs, the
 * sheet and the part writers give.
 */
function presentationChartLeaf(
  block: unknown,
): { family: string; subtype: string } | null {
  if (!block || typeof block !== "object") return null;
  const record = block as Record<string, unknown>;
  const source =
    record.source && typeof record.source === "object"
      ? (record.source as Record<string, unknown>)
      : null;
  const view =
    source?.nativeChartView && typeof source.nativeChartView === "object"
      ? (source.nativeChartView as Record<string, unknown>).attrs
      : null;
  for (const candidate of [record.chartV2, view]) {
    if (!candidate || typeof candidate !== "object") continue;
    const attrs = candidate as Record<string, unknown>;
    if (
      typeof attrs.chartType === "string" &&
      typeof attrs.chartSubtype === "string"
    )
      return { family: attrs.chartType, subtype: attrs.chartSubtype };
  }
  return null;
}

/** Gridlines exist only on chart families with a value axis. A filled map is
 * drawn on geography and Office gives it none, and a funnel's one axis is its
 * categories — Excel's Chart Elements menu offers it no gridlines — as the
 * shared contract's `CHART_GRIDLINE_FAMILIES` says. */
export function chartSupportsGridlines(block: unknown): boolean {
  const record =
    block && typeof block === "object"
      ? (block as Record<string, unknown>)
      : {};
  const leaf = presentationChartLeaf(block);
  if (leaf) return sharedChartSupportsGridlines(leaf.family);
  const kind = String(record.chartType ?? record.chart_type ?? "").toLowerCase();
  return !/(?:pie|doughnut|donut|treemap|sunburst|map|funnel)/.test(kind);
}

/** Primary major gridlines default ON on axis-bearing PowerPoint charts. */
export function resolveChartGridlines(block: unknown): boolean {
  if (!chartSupportsGridlines(block)) return false;
  return !(
    block &&
    typeof block === "object" &&
    (block as Record<string, unknown>).showGridlines === false
  );
}

/** Persist only the non-default OFF atom for primary major gridlines. */
export function applyChartGridlines<T extends Record<string, unknown>>(
  block: T,
  visible: boolean,
): T {
  const updated: Record<string, unknown> = { ...block };
  if (visible) delete updated.showGridlines;
  else {
    updated.showGridlines = false;
    delete updated.gridlineStyle;
  }
  return updated as T;
}

/** Data tables are bounded to category-based Cartesian chart families. */
export function chartSupportsDataTable(block: unknown): boolean {
  const record =
    block && typeof block === "object"
      ? (block as Record<string, unknown>)
      : {};
  const kind = String(record.chartType ?? record.chart_type ?? "column")
    .trim()
    .toLowerCase();
  const leaf = presentationChartLeaf(block);
  if (leaf) return sharedChartSupportsDataTable(leaf.family);
  return /(?:bar|column|line|area|stack)/.test(kind);
}

/** PowerPoint chart data tables default OFF. */
export function resolveChartDataTable(block: unknown): boolean {
  return Boolean(
    chartSupportsDataTable(block) &&
      block &&
      typeof block === "object" &&
      (block as Record<string, unknown>).showDataTable === true,
  );
}

/** Persist the non-default ON atom and omit the default OFF atom. */
export function applyChartDataTable<T extends Record<string, unknown>>(
  block: T,
  visible: boolean,
): T {
  const updated: Record<string, unknown> = { ...block };
  if (visible && chartSupportsDataTable(block)) updated.showDataTable = true;
  else {
    delete updated.showDataTable;
    delete updated.dataTableStyle;
  }
  return updated as T;
}

/** Native PowerPoint charts place the legend on the right when unspecified. */
export function resolveChartLegendPosition(
  block: unknown,
): PresentationChartLegendPosition {
  const value =
    block && typeof block === "object"
      ? (block as Record<string, unknown>).legendPosition
      : undefined;
  return (PRESENTATION_CHART_LEGEND_POSITIONS as readonly unknown[]).includes(
    value,
  )
    ? (value as PresentationChartLegendPosition)
    : "right";
}

/** Persist only a non-default placement; right is PowerPoint's native default. */
export function applyChartLegendPosition<T extends Record<string, unknown>>(
  block: T,
  position: PresentationChartLegendPosition,
): T {
  const updated: Record<string, unknown> = { ...block };
  if (position === "right") delete updated.legendPosition;
  else updated.legendPosition = position;
  return updated as T;
}

/** Linear trendlines are bounded to ordered cartesian chart families. */
export function chartSupportsLinearTrendline(block: unknown): boolean {
  const record =
    block && typeof block === "object"
      ? (block as Record<string, unknown>)
      : {};
  const kind = String(record.chartType ?? record.chart_type ?? "column")
    .trim()
    .toLowerCase();
  const leaf = presentationChartLeaf(block);
  if (leaf) return chartSupportsTrendline(leaf.family, leaf.subtype);
  return /(?:bar|column|line|area|scatter)/.test(kind);
}

/** Read only the presentation-owned linear trendline subset. */
export function resolveChartLinearTrendline(block: unknown): boolean {
  if (
    !chartSupportsLinearTrendline(block) ||
    !block ||
    typeof block !== "object"
  ) {
    return false;
  }
  const trendline = (block as Record<string, unknown>).trendline;
  return Boolean(
    trendline &&
      typeof trendline === "object" &&
      !Array.isArray(trendline) &&
      (trendline as Record<string, unknown>).type === "linear",
  );
}

/** Persist one native-compatible atom, or remove it entirely when disabled. */
export function applyChartLinearTrendline<T extends Record<string, unknown>>(
  block: T,
  visible: boolean,
): T {
  const updated: Record<string, unknown> = { ...block };
  if (visible && chartSupportsLinearTrendline(block)) {
    updated.trendline = {
      type: "linear",
    } satisfies PresentationChartLinearTrendline;
  } else {
    delete updated.trendline;
    delete updated.trendlineStyle;
  }
  return updated as T;
}

/**
 * Apply one Quick Layout through the same default-on / explicit-false atoms as
 * the individual Add Chart Element commands. An already-matching layout keeps
 * object identity so callers cannot create a no-op history transaction.
 */
export function applyPresentationChartQuickLayout<
  T extends Record<string, unknown>,
>(block: T, id: PresentationChartQuickLayoutId): T {
  if (presentationChartQuickLayoutId(block) === id) return block;
  const layout = presentationChartQuickLayout(id);
  const laidOut = applyChartToggle(
    applyChartToggle(block, "showLegend", layout.showLegend),
    "showAxisLabels",
    layout.showAxisLabels,
  ) as Record<string, unknown>;
  // A layout decides the axes as a whole. A per-axis pick made earlier
  // (`chartElements`, applied after the block's own switch) would outvote it,
  // so the layout takes those two keys back.
  const elements = laidOut.chartElements;
  if (elements && typeof elements === "object" && !Array.isArray(elements)) {
    const { axisX: _x, axisY: _y, ...rest } = elements as Record<string, unknown>;
    if (Object.keys(rest).length > 0) laidOut.chartElements = rest;
    else delete laidOut.chartElements;
  }
  return laidOut as T;
}
