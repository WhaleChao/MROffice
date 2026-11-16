/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The "Add Chart Element" picks, as pure changes to a Chart V2 block.
 *
 * Docs, the deck and the sheet each grew their own menu, and each menu its own
 * idea of which chart takes which element: the deck offered a trendline on any
 * column, bar, line or scatter — a stacked or a 3-D one included — and error
 * bars on the same four, so the pick drew on the slide and was given up at the
 * save, where the shared contract refuses it. One element has one rule, the
 * one `chartBlockShared` states and the server's `document_charts` repeats;
 * every host's menu asks it here, and every pick is made here, so a control a
 * host shows is a control the file can carry.
 *
 * Each function returns the changed block, or `null` when the chart cannot
 * take the element, the pick changes nothing, or the result is not a valid
 * chart. A host commits what it is given and never edits `format` by hand.
 */
import { fitChartTitleRuns } from "./chart-title-runs";
import { chartLegendEntryOfPosition } from "./chart-element-placement";
import {
  chartDataLabelOffsetSigns,
  chartTakesElementLayout,
  normalizeChartElementLayout,
  withChartElementPlace,
  type ChartElementLayoutKey,
} from "./chart-element-layout";
import { officeChartAxisRole, officeChartPointArea, type OfficeChartPicturePart } from "./chart-hit";
import {
  CHART_DISPLAY_BLANKS_AS,
  chartDefaultView3D,
  chartTakesView3D,
  chartTakesBarSpacing,
  chartTakesBarShape,
  chartTakesGapDepth,
  isChartBarShape,
  type ChartBarShape,
  chartBarSpacingDefaults,
  chartPieOptionsRefusal,
  effectiveChartView3D,
  normalizeChartView3D,
  sameChartView3D,
  chartDataLabelPositions,
  chartDataLabelStatement,
  chartFormatReconcileDataLabelOverrides,
  effectiveChartDisplayBlanksAs,
  chartHasAxis,
  chartHasSecondaryValueAxis,
  chartHasDepthAxis,
  chartMapLabelLayouts,
  chartSupportsDataLabels,
  chartSupportsDataLabelCallout,
  chartSupportsDataTable,
  chartSupportsDropLines,
  chartDataTakesErrorBars,
  chartFormatFollowingDataEdit,
  chartSupportsGridlines,
  chartSupportsMinorGridlines,
  chartSupportsHighLowLines,
  chartSupportsSeriesLines,
  chartSupportsTrendline,
  chartSupportsUpDownBars,
  chartSupportsAxisTitle,
  chartSeriesRenderAsOf,
  chartSeriesTakesMarker,
  chartSeriesTakesSmooth,
  clampChartText,
  isChartDisplayBlanksAs,
  normalizeChartDataLabelOverrides,
  normalizeChartErrorBars,
  normalizeChartSeriesMarker,
  normalizeChartSeriesTrendlines,
  normalizeChartTrendline,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ChartBlockFormat,
  type ChartBoxWhisker,
  type ChartParentLabelLayout,
  type ChartDataLabelOverride,
  type ChartDataLabelPosition,
  type ChartDataLabelStatement,
  type ChartDisplayBlanksAs,
  type ChartErrorBars,
  type ChartView3D,
  type ChartMapLabelLayout,
  type ChartSeriesMarker,
  type ChartTrendline,
} from "../../components/console/editor/extensions/chartBlockShared";
import { chartPreset } from "../../components/console/editor/extensions/chartPresets";
import {
  chartColorPaletteColor,
  chartColorTargetCount,
  DEFAULT_CHART_THEME_ACCENTS,
} from "./chart-color-palettes";
import {
  MAX_CHART_OUTLINE_WIDTH_PT,
  normalizeOfficeChartShapePaint,
} from "../office-chart-shape-paint";
import type { PresentationChartSeriesStyle } from "../presentation-chart-area-style";
import {
  normalizeOfficeChartLayout,
  officeChartFormatWithOneLayoutRecord,
  officeChartLayoutOf,
} from "./chart-layout";
import { officeTrendlineKindFits } from "./trendline";
import { chartSeriesPlottedOn, chartWithSeriesAxis } from "./chart-type";
import { officeChartDateAxisTimes, withOfficeChartAxisOptions } from "./axis-options";

export type ChartElementLine = "drop" | "highLow" | "series";
export type ChartElementErrorBars = ChartErrorBars;

/** The three built-in error amounts Office's menu lists, as the contract spells them. */
export const CHART_ELEMENT_ERROR_BAR_RECIPES = [
  { type: "standardError" },
  { type: "percentage", value: 5 },
  { type: "standardDeviation", value: 1 },
] as const satisfies readonly ChartElementErrorBars[];

/** Office's own names for the label positions, as catalogue keys. */
export const CHART_DATA_LABEL_POSITION_NAMES: Record<
  ChartDataLabelPosition,
  string
> = {
  bestFit: "office.chart.labelPosition.bestFit",
  bottom: "office.chart.labelPosition.bottom",
  center: "office.chart.labelPosition.center",
  insideBase: "office.chart.labelPosition.insideBase",
  insideEnd: "office.chart.labelPosition.insideEnd",
  left: "office.chart.labelPosition.left",
  outsideEnd: "office.chart.labelPosition.outsideEnd",
  right: "office.chart.labelPosition.right",
  top: "office.chart.labelPosition.top",
};

/**
 * Office's own names for a filled map's Map labels values.
 *
 * The English is Microsoft's ("Best fit only" / "Show all", Format a Map
 * Chart); the third member of the tri-state, "None", is the labels switch
 * itself. A map's labels are geographic NAMES, not values, so this list
 * replaces the position list on that one family rather than joining it.
 */
export const CHART_MAP_LABEL_LAYOUT_NAMES: Record<ChartMapLabelLayout, string> =
  {
    bestFitOnly: "office.chart.mapLabelLayout.bestFitOnly",
    showAll: "office.chart.mapLabelLayout.showAll",
  };

/** The one placement token a chart's labels take: a position, or a map layout. */
export type ChartDataLabelPlacement =
  | ChartDataLabelPosition
  | ChartMapLabelLayout;

function seriesCount(attrs: ChartBlockAttrsV2): number {
  return attrs.data.kind === "category" ? attrs.data.series.length : 0;
}

/** Which elements one chart takes. The menu disables exactly the rest. */
export interface ChartElementAvailability {
  dataLabels: boolean;
  dataLabelPositions: readonly ChartDataLabelPosition[];
  /**
   * A filled map's Map labels values. One family has these and no positions,
   * and every other family the reverse, so a host renders ONE dropdown from
   * whichever list is non-empty (`chartDataLabelPlacements`) and sends the pick
   * through the same `chartWithDataLabels`.
   */
  dataLabelRegionLayouts: readonly ChartMapLabelLayout[];
  /** Data Labels > Data Callout (a 2-D classic chart). */
  dataLabelCallout: boolean;
  axisX: boolean;
  axisY: boolean;
  /** Axes > Secondary Vertical: a chart that draws a secondary value axis. */
  axisSecondaryY: boolean;
  /** Axes > Depth Axis: a 3-D chart with a depth (series) axis. */
  axisDepth: boolean;
  gridlines: boolean;
  dataTable: boolean;
  trendline: boolean;
  errorBars: boolean;
  upDownBars: boolean;
  dropLines: boolean;
  highLowLines: boolean;
  seriesLines: boolean;
}

export function chartElementAvailability(
  attrs: ChartBlockAttrsV2,
): ChartElementAvailability {
  const family = attrs.chartType,
    subtype = attrs.chartSubtype;
  // Up-down bars, high-low lines and series lines are all drawn BETWEEN
  // series, so a chart with one series has nothing to draw them between.
  const several = seriesCount(attrs) >= 2;
  const dataLabels = chartSupportsDataLabels(family);
  return {
    dataLabels,
    dataLabelPositions: dataLabels
      ? chartDataLabelPositions(family, subtype)
      : [],
    dataLabelRegionLayouts: dataLabels ? chartMapLabelLayouts(family) : [],
    dataLabelCallout: dataLabels && chartSupportsDataLabelCallout(family, subtype),
    axisX: chartHasAxis(family, "x"),
    axisY: chartHasAxis(family, "y"),
    axisSecondaryY: chartHasSecondaryValueAxis(family, subtype, attrs.data),
    axisDepth: chartHasDepthAxis(family, subtype),
    gridlines: chartSupportsGridlines(family),
    dataTable: chartSupportsDataTable(family),
    trendline: chartSupportsTrendline(family, subtype),
    errorBars: chartDataTakesErrorBars(family, subtype, attrs.data),
    upDownBars: chartSupportsUpDownBars(family, subtype) && several,
    dropLines: chartSupportsDropLines(family, subtype),
    highLowLines: chartSupportsHighLowLines(family, subtype) && several,
    seriesLines: chartSupportsSeriesLines(family, subtype) && several,
  };
}

function committed(
  attrs: ChartBlockAttrsV2,
  format: ChartBlockFormat,
): ChartBlockAttrsV2 | null {
  // A chart-wide label command reaches every series' and every point's own
  // labels too, as Office's does.
  const labelled = chartFormatReconcileDataLabelOverrides(attrs.format, format) ?? format;
  // A layout's atoms are one record. A command that wrote the layout wrote it
  // from what the chart states (`officeChartLayoutOf`): the chart's own
  // spacing keys go. Any other command keeps the layout as the chart states
  // it (`officeChartFormatWithOneLayoutRecord`).
  const reconciled =
    labelled.layout !== undefined && JSON.stringify(labelled.layout) !== JSON.stringify(attrs.format?.layout)
      ? (({ barGapWidth: _gap, barOverlap: _overlap, ...rest }) => rest)(labelled)
      : officeChartFormatWithOneLayoutRecord(attrs.chartType, labelled);
  const next: ChartBlockAttrsV2 = {
    ...attrs,
    format: Object.keys(reconciled).length ? reconciled : undefined,
  };
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

export function chartDataLabelsShown(attrs: ChartBlockAttrsV2): boolean {
  return attrs.format?.dataLabels ?? chartPreset(attrs.stylePreset).dataLabels;
}

/** The dropdown a host draws beside the labels switch, for either family kind. */
export function chartDataLabelPlacements(
  attrs: ChartBlockAttrsV2,
): readonly ChartDataLabelPlacement[] {
  const available = chartElementAvailability(attrs);
  return available.dataLabelRegionLayouts.length
    ? available.dataLabelRegionLayouts
    : available.dataLabelPositions;
}

/** That dropdown's current pick, for either family kind. */
export function chartDataLabelPlacement(
  attrs: ChartBlockAttrsV2,
): ChartDataLabelPlacement | null {
  return chartElementAvailability(attrs).dataLabelRegionLayouts.length
    ? (attrs.format?.dataLabelRegionLayout ?? null)
    : (attrs.format?.dataLabelPosition ?? null);
}

/** Its label's catalogue key, for either family kind. */
export function chartDataLabelPlacementName(
  placement: ChartDataLabelPlacement,
): string {
  return (
    CHART_MAP_LABEL_LAYOUT_NAMES[placement as ChartMapLabelLayout] ??
    CHART_DATA_LABEL_POSITION_NAMES[placement as ChartDataLabelPosition]
  );
}

/**
 * Labels on or off, at `placement` — `null` is where Office puts them itself.
 *
 * `placement` is a data-label POSITION on the seven families that write
 * `cx:dataLabels`, and a filled map's Map labels layout on the one that does
 * not: a map states its labels as `cx:layoutPr/cx:regionLabelLayout`
 * (None / Best fit only / Show all), shows region NAMES rather than values, and
 * takes no position at all. One switch and one dropdown serve both; which key
 * the pick lands on is decided here, once, for the three editors.
 */
export function chartWithDataLabels(
  attrs: ChartBlockAttrsV2,
  visible: boolean,
  placement: ChartDataLabelPlacement | null = null,
  callout = false,
): ChartBlockAttrsV2 | null {
  const available = chartElementAvailability(attrs);
  if (!available.dataLabels) return null;
  if (callout) {
    // Data Callout: the labels in a callout bubble at Office's own place,
    // showing the category name and the value (as Excel's does on a column
    // chart); any other pick of the submenu takes the bubble off again.
    if (!visible || !available.dataLabelCallout) return null;
    if (attrs.format?.dataLabelCallout === true && chartDataLabelsShown(attrs)) return null;
    const format: ChartBlockFormat = {
      ...attrs.format,
      dataLabels: true,
      dataLabelCallout: true,
      dataLabelCategoryName: true,
    };
    delete format.dataLabelPosition;
    delete format.dataLabelValue;
    delete format.dataLabelLeaderLines;
    delete format.dataLabelLeaderLineStyle;
    return committed(attrs, format);
  }
  const mapLabels = available.dataLabelRegionLayouts.length > 0;
  const offered: readonly string[] = mapLabels
    ? available.dataLabelRegionLayouts
    : available.dataLabelPositions;
  if (visible && placement && !offered.includes(placement)) return null;
  const currentPosition = mapLabels
    ? (attrs.format?.dataLabelRegionLayout ?? null)
    : (attrs.format?.dataLabelPosition ?? null);
  // A calibrated Quick Layout's single labels (Layout 6's last category, 10's
  // last series) are data labels too: None removes them with the rest, as
  // Office's None removes every label of the chart (PPT-B090).
  const layout = normalizeOfficeChartLayout(attrs.format?.layout);
  const layoutLabels = (layout?.labels.length ?? 0) > 0;
  if (
    chartDataLabelsShown(attrs) === visible &&
    (visible ? currentPosition === placement : !layoutLabels) &&
    attrs.format?.dataLabelCallout !== true
  )
    return null;
  const format: ChartBlockFormat = { ...attrs.format, dataLabels: visible };
  if (!visible && layout && layoutLabels) format.layout = { ...layout, labels: [] };
  delete format.dataLabelCallout;
  delete format.dataLabelPosition;
  delete format.dataLabelRegionLayout;
  if (visible && placement) {
    if (mapLabels) format.dataLabelRegionLayout = placement as ChartMapLabelLayout;
    else format.dataLabelPosition = placement as ChartDataLabelPosition;
  }
  if (visible && mapLabels) {
    // A map's labels are its regions' names. Excel writes no value on one, so
    // the switch must not send `dataLabelValue` for this family — it would be
    // content the file cannot carry, on the family Microsoft's own answer says
    // does not fully support data labels.
    format.dataLabelCategoryName = true;
    delete format.dataLabelValue;
    delete format.dataLabelSeriesName;
    delete format.dataLabelPercentage;
    delete format.dataLabelBubbleSize;
    delete format.dataLabelLegendKey;
    delete format.dataLabelSeparator;
  }
  if (!visible) {
    // A position and every content field require visible labels. An imported
    // chart states its own, so leaving them behind makes the edited chart
    // invalid and the switch reads as dead.
    delete format.dataLabelValue;
    delete format.dataLabelSeriesName;
    delete format.dataLabelCategoryName;
    delete format.dataLabelPercentage;
    delete format.dataLabelBubbleSize;
    delete format.dataLabelLegendKey;
    delete format.dataLabelSeparator;
  }
  return committed(attrs, format);
}

/** What a data label shows: Format Data Labels > Label Contains, Separator and
 * Show Leader Lines, as a host's pane states them. */
export interface ChartDataLabelContents {
  readonly dataLabelValue: boolean;
  readonly dataLabelLegendKey: boolean;
  readonly dataLabelCategoryName: boolean;
  readonly dataLabelSeriesName: boolean;
  readonly dataLabelPercentage: boolean;
  readonly dataLabelBubbleSize: boolean;
  readonly dataLabelSeparator: NonNullable<ChartBlockFormat["dataLabelSeparator"]>;
  readonly dataLabelLeaderLines: boolean;
}

/**
 * The chart format with these label contents on, stated as the chart's model
 * states them: each field only when it differs from its default, and the
 * separator only between two or more texts. A pane that kept the separator
 * when a second text went off wrote a chart the validator refuses, so the
 * check box silently stayed on (PPT-B098). One reading for every host.
 */
export function chartFormatWithDataLabelContents(
  format: ChartBlockFormat | undefined,
  contents: ChartDataLabelContents,
): ChartBlockFormat {
  const next: ChartBlockFormat = { ...(format ?? {}), dataLabels: true };
  delete next.dataLabelValue;
  delete next.dataLabelLegendKey;
  delete next.dataLabelCategoryName;
  delete next.dataLabelSeriesName;
  delete next.dataLabelPercentage;
  delete next.dataLabelBubbleSize;
  delete next.dataLabelSeparator;
  if (!contents.dataLabelValue) next.dataLabelValue = false;
  if (contents.dataLabelLegendKey) next.dataLabelLegendKey = true;
  if (contents.dataLabelCategoryName) next.dataLabelCategoryName = true;
  if (contents.dataLabelSeriesName) next.dataLabelSeriesName = true;
  if (contents.dataLabelPercentage) next.dataLabelPercentage = true;
  if (contents.dataLabelBubbleSize) next.dataLabelBubbleSize = true;
  const texts = [
    contents.dataLabelValue,
    contents.dataLabelCategoryName,
    contents.dataLabelSeriesName,
    contents.dataLabelPercentage,
    contents.dataLabelBubbleSize,
  ].filter(Boolean).length;
  if (texts > 1) next.dataLabelSeparator = contents.dataLabelSeparator;
  if (contents.dataLabelLeaderLines) next.dataLabelLeaderLines = true;
  else {
    delete next.dataLabelLeaderLines;
    delete next.dataLabelLeaderLineStyle;
  }
  return next;
}

/** One series' labels, or one point's label, addressed on the one model. */
export interface ChartDataLabelTarget {
  readonly seriesIndex: number;
  /** Absent: the whole series' labels. */
  readonly pointIndex?: number;
}

/** What Format Data Labels on a series or on one label changes: a key set
 * states it for the target, an undefined key hands it back to the chart's
 * (or the series', for a point). */
export type ChartDataLabelOverridePatch = Partial<
  Omit<ChartDataLabelOverride, "seriesIndex" | "pointIndex">
>;

/** What the target's labels show: the chart's, then the series', then the
 * point's own statement (the resolution the renderer draws). */
export function chartDataLabelStatementOf(
  attrs: ChartBlockAttrsV2,
  target: ChartDataLabelTarget,
): ChartDataLabelStatement {
  return chartDataLabelStatement(
    attrs.format,
    chartDataLabelsShown(attrs),
    target.seriesIndex,
    target.pointIndex,
  );
}

/** The target's own label statement in `format.dataLabelOverrides`, if any. */
export function chartDataLabelOverrideOf(
  attrs: ChartBlockAttrsV2,
  target: ChartDataLabelTarget,
): ChartDataLabelOverride | undefined {
  return attrs.format?.dataLabelOverrides?.find(
    (entry) =>
      entry.seriesIndex === target.seriesIndex &&
      (entry.pointIndex ?? null) === (target.pointIndex ?? null),
  );
}

/**
 * Format Data Labels on one series or on one label: Office writes the
 * change on that series' `c:dLbls` or that point's `c:dLbl`, and the rest of
 * the chart keeps its labels. `patch` merges into the target's own statement
 * (`format.dataLabelOverrides`); `null` drops that statement, so the target
 * follows the chart again. Hidden labels state nothing else, as the model
 * has it. Null when nothing changes or the chart cannot carry it (a filled
 * map's region labels are the chart's alone; a position the family does not
 * offer; a percentage off a pie).
 */
export function chartWithDataLabelOverride(
  attrs: ChartBlockAttrsV2,
  target: ChartDataLabelTarget,
  patch: ChartDataLabelOverridePatch | null,
): ChartBlockAttrsV2 | null {
  const available = chartElementAvailability(attrs);
  if (!available.dataLabels || available.dataLabelRegionLayouts.length) return null;
  const overrides = attrs.format?.dataLabelOverrides ?? [];
  const prior = chartDataLabelOverrideOf(attrs, target);
  const rest = overrides.filter((entry) => entry !== prior);
  let entry: Record<string, unknown> | null = null;
  if (patch) {
    entry = {
      seriesIndex: target.seriesIndex,
      ...(target.pointIndex !== undefined ? { pointIndex: target.pointIndex } : {}),
      ...(prior ?? {}),
    };
    for (const [key, value] of Object.entries(patch)) {
      if (value === undefined) delete entry[key];
      else entry[key] = value;
    }
    // A position stated or handed back is the label's place: no longer the
    // chart type's own (`dataLabelPositionAutomatic`).
    if ("dataLabelPosition" in patch && !("dataLabelPositionAutomatic" in patch)) delete entry.dataLabelPositionAutomatic;
    if (entry.dataLabels === false) {
      entry = {
        seriesIndex: target.seriesIndex,
        ...(target.pointIndex !== undefined ? { pointIndex: target.pointIndex } : {}),
        dataLabels: false,
      };
    }
    const stated = Object.keys(entry).filter(
      (key) => key !== "seriesIndex" && key !== "pointIndex",
    );
    if (!stated.length) entry = null;
  }
  const list = entry ? [...rest, entry] : rest;
  const normalized = list.length ? normalizeChartDataLabelOverrides(list) : undefined;
  if (list.length && !normalized) return null;
  if (
    JSON.stringify(normalized ?? null) ===
    JSON.stringify(attrs.format?.dataLabelOverrides ?? null)
  )
    return null;
  const format: ChartBlockFormat = { ...attrs.format };
  if (normalized) format.dataLabelOverrides = normalized;
  else delete format.dataLabelOverrides;
  return committed(attrs, format);
}

/**
 * Hidden and Empty Cells ▸ Show empty cells as: the methods this chart's
 * family takes (a method the family draws as another is not offered; a
 * family that draws every blank as zero -- stacked lines and areas, 3-D
 * areas, every 3-D chart, a funnel -- offers none).
 */
export function chartDisplayBlanksOptions(
  attrs: ChartBlockAttrsV2,
): readonly ChartDisplayBlanksAs[] {
  if (attrs.chartSubtype.endsWith("3d") || attrs.chartType === "funnel") return [];
  const [family, subtype] =
    attrs.chartType === "combo" ? ["line", "standard"] : [attrs.chartType, attrs.chartSubtype];
  const options = CHART_DISPLAY_BLANKS_AS.filter(
    (method) => effectiveChartDisplayBlanksAs(family, subtype, method) === method,
  );
  return options.length > 1 ? options : [];
}

/** What the chart draws for a blank point: its own setting through its family's rule. */
export function chartDisplayBlanksShown(attrs: ChartBlockAttrsV2): ChartDisplayBlanksAs {
  if (attrs.chartSubtype.endsWith("3d") || attrs.chartType === "funnel") return "zero";
  const [family, subtype] =
    attrs.chartType === "combo" ? ["line", "standard"] : [attrs.chartType, attrs.chartSubtype];
  return effectiveChartDisplayBlanksAs(family, subtype, attrs.format?.displayBlanksAs);
}

/** Show empty cells as `value` (`c:dispBlanksAs`), or null when the family
 * does not take it or the chart already says so. */
export function chartWithDisplayBlanksAs(
  attrs: ChartBlockAttrsV2,
  value: ChartDisplayBlanksAs,
): ChartBlockAttrsV2 | null {
  if (!chartDisplayBlanksOptions(attrs).includes(value)) return null;
  if (attrs.format?.displayBlanksAs === value) return null;
  return committed(attrs, { ...attrs.format, displayBlanksAs: value });
}

/**
 * A Copilot's `format.displayBlanksAs` leaf, written as the Hidden & Empty
 * Cells menu writes it -- through the one element intent
 * (`chartWithElementIntent`): the chart itself when it already says so (a
 * null leaf, like an absent key, is Gaps), null when its family does not
 * take the method (the patch is refused, as the menu would be greyed).
 */
export function chartWithDisplayBlanksLeaf(
  attrs: ChartBlockAttrsV2,
  value: unknown,
): ChartBlockAttrsV2 | null {
  const method = value === null || value === undefined ? "gap" : value;
  if (!isChartDisplayBlanksAs(method)) return null;
  if ((attrs.format?.displayBlanksAs ?? "gap") === method) return attrs;
  return chartWithElementIntent(attrs, { element: "displayBlanksAs", value: method });
}

/**
 * Format Chart Area ▸ 3-D Rotation of a 3-D chart: X Rotation (`rotY`, the
 * turn, 0..359 -- a 3-D pie's Angle of First Slice), Y Rotation (`rotX`, the
 * tilt, -90..90, 0..90 for a pie), Perspective (the field of view, stored in
 * half degrees as the file states it), Right Angle Axes, Depth and Height
 * (% of base; a Height of `undefined` is Autoscale). `null` is Default
 * Rotation: the view Office gives the leaf, stated, so a part turned
 * elsewhere is turned back. Null when the chart takes no view or already
 * draws that one as stated.
 */
export function chartWithView3D(
  attrs: ChartBlockAttrsV2,
  view: Partial<ChartView3D> | null,
): ChartBlockAttrsV2 | null {
  if (!chartTakesView3D(attrs.chartType, attrs.chartSubtype)) return null;
  const defaults = chartDefaultView3D(attrs.chartType, attrs.chartSubtype);
  const next = view === null
    ? { ...defaults }
    : normalizeChartView3D({ ...effectiveChartView3D(attrs), ...view }, defaults);
  if (!next) return null;
  // A pie is not tilted below its plane (Excel offers it 0..90).
  if (attrs.chartType === "pie" && view !== null && view.rotX !== undefined && view.rotX < 0) return null;
  const stated = normalizeChartView3D(attrs.format?.view3D, defaults);
  if (stated && sameChartView3D(stated, next)) return null;
  if (!stated && view === null) return null;
  return committed(attrs, { ...attrs.format, view3D: next });
}

export function chartWithDataTable(
  attrs: ChartBlockAttrsV2,
  visible: boolean,
  legendKeys = true,
): ChartBlockAttrsV2 | null {
  if (!chartElementAvailability(attrs).dataTable) return null;
  const shown = attrs.format?.dataTable ?? false;
  const keys = attrs.format?.dataTableLegendKeys !== false;
  if (shown === visible && (!visible || keys === legendKeys)) return null;
  const format: ChartBlockFormat = { ...attrs.format, dataTable: visible };
  delete format.dataTableLegendKeys;
  if (visible && !legendKeys) format.dataTableLegendKeys = false;
  return committed(attrs, format);
}

/**
 * Format Data Table ▸ Table Borders: the table's Horizontal, Vertical and
 * Outline borders (`c:showHorzBorder`, `c:showVertBorder`, `c:showOutline`).
 * Each is shown unless stated off, as Office draws a new data table; a
 * border turned back on states nothing. Null on a chart showing no data
 * table, or when nothing changes.
 */
export function chartWithDataTableBorders(
  attrs: ChartBlockAttrsV2,
  borders: { readonly horizontal?: boolean; readonly vertical?: boolean; readonly outline?: boolean },
): ChartBlockAttrsV2 | null {
  if (!chartElementAvailability(attrs).dataTable || attrs.format?.dataTable !== true) return null;
  const format: ChartBlockFormat = { ...attrs.format };
  for (const [side, key] of [
    ["horizontal", "dataTableHorizontalBorder"],
    ["vertical", "dataTableVerticalBorder"],
    ["outline", "dataTableOutline"],
  ] as const) {
    const shown = borders[side];
    if (shown === undefined) continue;
    if (shown) delete format[key];
    else format[key] = false;
  }
  return JSON.stringify(format) === JSON.stringify(attrs.format) ? null : committed(attrs, format);
}

/**
 * Format Data Series ▸ Series Options of a pie or doughnut: Angle of first
 * slice, Doughnut Hole Size and Pie (Doughnut) Explosion. A value at the
 * writer's default (0°, 50 %, 0 %) states nothing. Null on a chart that
 * cannot take the change, or when nothing changes.
 */
export function chartWithPieOptions(
  attrs: ChartBlockAttrsV2,
  options: { readonly firstSliceAngle?: number; readonly doughnutHoleSize?: number; readonly pieExplosion?: number },
): ChartBlockAttrsV2 | null {
  if (attrs.chartType !== "pie") return null;
  const format: ChartBlockFormat = { ...attrs.format };
  for (const [key, fallback] of [
    ["firstSliceAngle", 0],
    ["doughnutHoleSize", 50],
    ["pieExplosion", 0],
  ] as const) {
    const value = options[key];
    if (value === undefined) continue;
    if (value === fallback) delete format[key];
    else format[key] = value;
  }
  if (
    JSON.stringify(format) === JSON.stringify(attrs.format ?? {}) ||
    chartPieOptionsRefusal(attrs.chartType, attrs.chartSubtype, format as Readonly<Record<string, unknown>>)
  )
    return null;
  return committed(attrs, format);
}

/** Format Data Point ▸ Point Explosion (0..400 %) of one slice of a pie. */
export function chartWithPointExplosion(
  attrs: ChartBlockAttrsV2,
  target: { readonly seriesIndex: number; readonly pointIndex: number; readonly explosion: number },
): ChartBlockAttrsV2 | null {
  if (attrs.chartType !== "pie" || !Number.isInteger(target.explosion) || target.explosion < 0 || target.explosion > 400)
    return null;
  const styles = [...(attrs.format?.sliceStyles ?? [])];
  const at = styles.findIndex(
    (style) => style.pointIndex === target.pointIndex && (style.seriesIndex ?? 0) === target.seriesIndex,
  );
  const current = at >= 0 ? styles[at] : undefined;
  if ((current?.explosion ?? 0) === target.explosion) return null;
  const next: Record<string, unknown> = {
    ...(current ?? { pointIndex: target.pointIndex, ...(target.seriesIndex ? { seriesIndex: target.seriesIndex } : {}) }),
  };
  if (target.explosion) next.explosion = target.explosion;
  else delete next.explosion;
  const keeps = Object.keys(next).some((key) => key !== "pointIndex" && key !== "seriesIndex");
  if (at >= 0) {
    if (keeps) styles[at] = next as (typeof styles)[number];
    else styles.splice(at, 1);
  } else if (keeps) styles.push(next as (typeof styles)[number]);
  const format: ChartBlockFormat = { ...attrs.format };
  if (styles.length) format.sliceStyles = styles;
  else delete format.sliceStyles;
  return committed(attrs, format);
}

export function chartWithAxis(
  attrs: ChartBlockAttrsV2,
  axis: "x" | "y" | "y2" | "z",
  visible: boolean,
): ChartBlockAttrsV2 | null {
  if (axis === "z") {
    // Axes > Depth Axis: a 3-D chart's depth (series) axis; shown is the
    // default, so only a hidden one is stated (`axisOptions.series.hidden`).
    if (!chartHasDepthAxis(attrs.chartType, attrs.chartSubtype)) return null;
    const record = attrs.format?.axisOptions?.series;
    if ((record?.hidden !== true) === visible) return null;
    const next = withOfficeChartAxisOptions(attrs, "series", { ...record, hidden: visible ? undefined : true });
    return committed(attrs, next.format ?? {});
  }
  if (axis === "y2") {
    // Axes > Secondary Vertical: the secondary value axis of a chart that
    // draws one; shown is the default, so only a hidden one is stated.
    if (!chartHasSecondaryValueAxis(attrs.chartType, attrs.chartSubtype, attrs.data)) return null;
    if ((attrs.format?.secondaryValueAxis ?? true) === visible) return null;
    const format: ChartBlockFormat = { ...attrs.format };
    if (visible) delete format.secondaryValueAxis;
    else format.secondaryValueAxis = false;
    return committed(attrs, format);
  }
  if (!chartHasAxis(attrs.chartType, axis)) return null;
  const key = axis === "x" ? "axisX" : "axisY";
  if ((attrs.format?.[key] ?? true) === visible) return null;
  return committed(attrs, { ...attrs.format, [key]: visible });
}

export function chartWithGridlines(
  attrs: ChartBlockAttrsV2,
  axis: "x" | "y",
  visible: boolean,
): ChartBlockAttrsV2 | null {
  if (!chartElementAvailability(attrs).gridlines) return null;
  const key = axis === "x" ? "gridlinesX" : "gridlinesY";
  if (attrs.format?.[key] === visible) return null;
  return committed(attrs, { ...attrs.format, [key]: visible });
}

/**
 * The Trendline menu, in one order for the three editors: Office's six kinds,
 * with the defaults its own menu uses for the two that take a number (a
 * second-order polynomial, a two-period moving average).
 */
export const CHART_ELEMENT_TRENDLINES = [
  { type: "linear" },
  { type: "exponential" },
  { type: "logarithmic" },
  { type: "polynomial", order: 2 },
  { type: "power" },
  { type: "movingAverage", period: 2 },
] as const satisfies readonly ChartTrendline[];

/**
 * Each category's x on a date axis (Format Axis > Axis Type: Date axis) as
 * Excel fits a trendline there: its date serial, so the fit is a straight
 * line over the dates' real spacing. Null for any other axis, or when a
 * category is not a date (the axis is then drawn as text, PPT-B309).
 */
export function chartTrendlineDateXs(attrs: ChartBlockAttrsV2): number[] | null {
  if (attrs.data.kind !== "category" && attrs.data.kind !== "stock") return null;
  const record = attrs.format?.axisOptions?.category;
  if (record?.axisType !== "date") return null;
  const times = officeChartDateAxisTimes(attrs.data.categories, record.baseTimeUnit);
  return times ? times.map((time) => time / 86_400_000 + 25_569) : null;
}

/**
 * The points a chart's trendline is fitted to: its first series, by category
 * place (1, 2, 3, ...) or by its own x on a scatter.
 */
export function chartTrendlineFitPoints(
  attrs: ChartBlockAttrsV2,
  seriesIndex = 0,
): Array<{ x: number; y: number }> {
  // A blank point (an empty cell) is no point of the fit; the others keep
  // their places.
  const plotted = (x: number | null, y: number | null) =>
    x === null || y === null || Number.isNaN(x) || Number.isNaN(y) ? [] : [{ x, y }];
  if (attrs.data.kind === "category" || attrs.data.kind === "stock") {
    const values = attrs.data.series[seriesIndex]?.values ?? [];
    if (values.length < 2 || attrs.data.categories.length !== values.length) return [];
    // On a date axis a point's x is its date; on a text axis its ordinal.
    const dates = chartTrendlineDateXs(attrs);
    return values.flatMap((y, index) => plotted(dates ? dates[index] : index + 1, y));
  }
  if (attrs.data.kind === "xy") {
    const series = attrs.data.series[seriesIndex];
    if (!series || series.xValues.length < 2) return [];
    return series.xValues.flatMap((x, index) => plotted(x, series.yValues[index] ?? null));
  }
  return [];
}

/** The trendline kinds this chart's data cannot take (greyed with a reason). */
export function chartUnfittableTrendlines(
  attrs: ChartBlockAttrsV2,
): ChartTrendline["type"][] {
  const points = chartTrendlineFitPoints(attrs);
  return CHART_ELEMENT_TRENDLINES.map((entry) => entry.type).filter(
    (type) => !officeTrendlineKindFits(type, points),
  );
}







export const CHART_TRENDLINE_NAMES: Record<ChartTrendline["type"], string> = {
  linear: "office.chart.trendline.linear",
  exponential: "office.chart.trendline.exponential",
  logarithmic: "office.chart.trendline.logarithmic",
  polynomial: "office.chart.trendline.polynomial",
  power: "office.chart.trendline.power",
  movingAverage: "spreadsheet.uiCopy.movingAverage",
};

/**
 * Add Chart Element ▸ Trendline, the rows Word, Excel and PowerPoint list:
 * Linear, Exponential, Linear Forecast and Two Period Moving Average (then
 * More Trendline Options…, whose Format Trendline pane offers all six kinds
 * and the forecast). Linear Forecast is a linear trendline forecast forward
 * two periods (Excel's own preset: Format Trendline > Forecast > Forward 2.0);
 * each row states its forecast, so picking Linear on a forecast trendline
 * takes the forecast off.
 */
export const CHART_ELEMENT_TRENDLINE_MENU = [
  { id: "linear", trendline: { type: "linear" }, forward: 0 },
  { id: "exponential", trendline: { type: "exponential" }, forward: 0 },
  { id: "linearForecast", trendline: { type: "linear" }, forward: 2 },
  { id: "movingAverage", trendline: { type: "movingAverage", period: 2 }, forward: 0 },
] as const satisfies ReadonlyArray<{
  id: string;
  trendline: ChartTrendline;
  forward: number;
}>;

export type ChartTrendlineMenuRowId =
  (typeof CHART_ELEMENT_TRENDLINE_MENU)[number]["id"];

/** Office's names for the menu rows, as catalogue keys. */
export const CHART_TRENDLINE_MENU_NAMES: Record<
  ChartTrendlineMenuRowId,
  string
> = {
  linear: CHART_TRENDLINE_NAMES.linear,
  exponential: CHART_TRENDLINE_NAMES.exponential,
  linearForecast: "office.chart.trendlineMenu.linearForecast",
  // The menu row is the two-period preset (Office's own row name).
  movingAverage: "office.chart.trendline.movingAverage",
};

/** The menu row a trendline is, or null (a kind only the pane offers). */
export function chartTrendlineMenuRow(
  trendline: ChartTrendline | null | undefined,
): ChartTrendlineMenuRowId | null {
  const current = normalizeChartTrendline(trendline);
  if (!current) return null;
  const forward = (current as { forward?: number }).forward ?? 0;
  const backward = (current as { backward?: number }).backward ?? 0;
  if (backward !== 0) return null;
  return (
    CHART_ELEMENT_TRENDLINE_MENU.find(
      (row) =>
        row.trendline.type === current.type &&
        row.forward === forward &&
        (row.trendline.type !== "movingAverage" ||
          (current as { period?: number }).period === row.trendline.period),
    )?.id ?? null
  );
}

/**
 * The Format pane's options the chart model states as one field each and
 * every writer already carries: Format Data Labels ▸ Alignment (Custom angle,
 * Wrap text in shape), a bubble chart's Scale bubble size to, a pie of pie's
 * Values in second plot, a box plot's Series Options, a treemap's Label
 * Options and a histogram's, pareto's or waterfall's Gap Width. `null` (or Wrap on) takes the field off:
 * Office's default.
 */
export type ChartFormatOption =
  | { key: "dataLabelRotation"; value: number | null }
  | { key: "dataLabelWrap"; value: boolean }
  | { key: "bubbleScale"; value: number | null }
  | { key: "ofPieSplitPosition"; value: number | null }
  | { key: "boxWhisker"; value: ChartBoxWhisker | null }
  | { key: "parentLabelLayout"; value: ChartParentLabelLayout | null }
  | { key: "gapWidth"; value: number | null };

/** The chart with one such option set, or null when nothing changes or the
 * model does not hold the value on this chart. */
export function chartWithFormatOption(
  attrs: ChartBlockAttrsV2,
  option: ChartFormatOption,
): ChartBlockAttrsV2 | null {
  const format: Record<string, unknown> = { ...attrs.format };
  if (option.key === "dataLabelWrap") {
    if (option.value) delete format.dataLabelWrap;
    else format.dataLabelWrap = false;
  } else if (option.value === null) delete format[option.key];
  else format[option.key] = option.value;
  if (JSON.stringify(format) === JSON.stringify({ ...attrs.format })) return null;
  return committed(attrs, format as ChartBlockFormat);
}

export function chartWithTrendline(
  attrs: ChartBlockAttrsV2,
  trendline: ChartTrendline | null,
): ChartBlockAttrsV2 | null {
  if (!chartElementAvailability(attrs).trendline) return null;
  if (
    JSON.stringify(normalizeChartTrendline(attrs.format?.trendline) ?? null) ===
    JSON.stringify(trendline)
  )
    return null;
  const format: ChartBlockFormat = { ...attrs.format };
  if (trendline) format.trendline = trendline;
  else {
    // Chart Elements ▸ Trendline off takes every trendline off, the other
    // series' too, as in Office.
    delete format.trendline;
    delete format.trendlineStyle;
    delete format.seriesTrendlines;
  }
  return committed(attrs, format);
}

/**
 * The trendline Format Data Series shows for one series: the first one that
 * series carries (the first series' is `format.trendline`, every other one is
 * among `format.seriesTrendlines`), or null.
 */
export function chartSeriesTrendline(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
): ChartTrendline | null {
  if (seriesIndex === 0) {
    const first = normalizeChartTrendline(attrs.format?.trendline);
    if (first) return first;
  }
  return (
    normalizeChartSeriesTrendlines(attrs.format?.seriesTrendlines)?.find(
      (entry) => entry.seriesIndex === seriesIndex,
    )?.trendline ?? null
  );
}

/**
 * Format Data Series ▸ Trendline on one series, as Office adds a trendline to
 * the series it formats (Microsoft, "Add a trend or moving average line to a
 * chart": the trendline is based on the series picked). The series' first
 * trendline takes `trendline`, keeping its own line paint; `null` takes every
 * trendline of that series off and leaves the other series' alone. Null when
 * the chart takes no trendline, the series is not the chart's, or nothing
 * changes.
 */
export function chartWithSeriesTrendline(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  trendline: ChartTrendline | null,
): ChartBlockAttrsV2 | null {
  if (!chartElementAvailability(attrs).trendline) return null;
  const seriesCount =
    attrs.data.kind === "category" || attrs.data.kind === "stock" || attrs.data.kind === "xy"
      ? attrs.data.series.length
      : 0;
  if (!Number.isInteger(seriesIndex) || seriesIndex < 0 || seriesIndex >= seriesCount)
    return null;
  const stated = trendline === null ? null : normalizeChartTrendline(trendline);
  if (stated === undefined) return null;
  const format: ChartBlockFormat = { ...attrs.format };
  const others = [...(normalizeChartSeriesTrendlines(format.seriesTrendlines) ?? [])];
  if (seriesIndex === 0) {
    if (stated) format.trendline = stated;
    else {
      delete format.trendline;
      delete format.trendlineStyle;
    }
  }
  const own = others.findIndex((entry) => entry.seriesIndex === seriesIndex);
  let entries = others;
  if (!stated) entries = others.filter((entry) => entry.seriesIndex !== seriesIndex);
  else if (seriesIndex !== 0) {
    if (own >= 0) entries[own] = { ...entries[own], trendline: stated };
    else {
      const at = entries.findIndex((entry) => entry.seriesIndex > seriesIndex);
      entries.splice(at < 0 ? entries.length : at, 0, { seriesIndex, trendline: stated });
    }
  }
  if (entries.length) format.seriesTrendlines = entries;
  else delete format.seriesTrendlines;
  if (JSON.stringify(format) === JSON.stringify({ ...attrs.format })) return null;
  return committed(attrs, format);
}

/**
 * Office's Format Trendline options this menu offers, and for which kinds
 * (Microsoft, "Add a trend or moving average line to a chart" / the Format
 * Trendline pane): Display Equation and Display R-squared value for every kind
 * but a moving average, Forecast forward likewise. They ride on the trendline,
 * so picking another KIND keeps the ones that kind still takes -- as Office
 * does -- instead of resetting the line.
 */
export type ChartTrendlineOptionKey = "displayEquation" | "displayRSquared";

export const CHART_TRENDLINE_OPTION_NAMES: Record<
  ChartTrendlineOptionKey,
  string
> = {
  displayEquation: "office.chart.trendlineOption.displayEquation",
  displayRSquared: "office.chart.trendlineOption.displayRSquared",
};

/** Forecast forward, in periods: Office's pane takes any number; the menu offers these. */
export const CHART_TRENDLINE_FORECASTS = [0, 1, 2] as const;

export function chartTrendlineTakesOptions(
  trendline: ChartTrendline | null | undefined,
): boolean {
  return Boolean(trendline) && trendline?.type !== "movingAverage";
}

/** `trendline` re-typed as `kind`, keeping every option the new kind takes. */
export function chartTrendlineOfKind(
  current: ChartTrendline | null | undefined,
  kind: ChartTrendline,
): ChartTrendline {
  if (!current) return kind;
  const { type: _type, ...rest } = current as Record<string, unknown> & {
    type: string;
  };
  delete rest.order;
  delete rest.period;
  // Try everything, then without the intercept (not offered for log / power,
  // and an exponential one must be above zero), then the name alone.
  for (const keep of [
    rest,
    Object.fromEntries(Object.entries(rest).filter(([key]) => key !== "intercept")),
    Object.fromEntries(Object.entries(rest).filter(([key]) => key === "name")),
  ]) {
    const candidate = normalizeChartTrendline({ ...kind, ...keep });
    if (candidate) return candidate;
  }
  return kind;
}

/**
 * One Format Trendline option: Display Equation / R-squared, Forecast
 * Forward / Backward (periods; 0 states none), Set Intercept (null clears
 * it), the Trendline Name (Custom; null is Automatic), a polynomial's Order
 * and a moving average's Period.
 */
export type ChartTrendlineOptionChange =
  | { key: ChartTrendlineOptionKey; on: boolean }
  | { key: "forward"; periods: number }
  | { key: "backward"; periods: number }
  | { key: "intercept"; value: number | null }
  | { key: "name"; name: string | null }
  | { key: "order"; order: number }
  | { key: "period"; period: number };

export function chartWithTrendlineOption(
  attrs: ChartBlockAttrsV2,
  change: Readonly<ChartTrendlineOptionChange>,
  /** The series whose trendline changes (Format Data Series); absent: the
   * chart's own, its first series'. */
  seriesIndex?: number,
): ChartBlockAttrsV2 | null {
  const current =
    seriesIndex === undefined
      ? normalizeChartTrendline(attrs.format?.trendline)
      : chartSeriesTrendline(attrs, seriesIndex);
  if (!current) return null;
  // A moving average takes its period and a name, nothing else.
  const kindOwn = change.key === "name" || change.key === "order" || change.key === "period";
  if (!kindOwn && !chartTrendlineTakesOptions(current)) return null;
  const next: Record<string, unknown> = { ...current };
  if (change.key === "forward" || change.key === "backward") {
    if (change.periods > 0) next[change.key] = change.periods;
    else delete next[change.key];
  } else if (change.key === "intercept") {
    if (change.value === null) delete next.intercept;
    else next.intercept = change.value;
  } else if (change.key === "name") {
    if (change.name === null || !change.name.trim()) delete next.name;
    else next.name = change.name;
  } else if (change.key === "order") {
    if (current.type !== "polynomial") return null;
    next.order = change.order;
  } else if (change.key === "period") {
    if (current.type !== "movingAverage") return null;
    next.period = change.period;
  } else if ("on" in change && change.on) next[change.key] = true;
  else delete next[change.key];
  const trendline = normalizeChartTrendline(next);
  if (!trendline) return null;
  return seriesIndex === undefined
    ? chartWithTrendline(attrs, trendline)
    : chartWithSeriesTrendline(attrs, seriesIndex, trendline);
}

/**
 * Error Bars at any of Office's Error Amounts (Format Error Bars ▸ Error
 * Amount: Fixed value, Percentage, Standard deviation(s), Standard error),
 * the menu's three picks among them; `null` removes the bars and their
 * format. Null when the chart takes none or the amount is not one Office
 * states (a negative, a zero percentage or deviation, Custom).
 */
export function chartWithErrorBars(
  attrs: ChartBlockAttrsV2,
  recipe: ChartElementErrorBars | null,
  /** Format Error Bars on another series than the first: that series'
   * own bars (`seriesErrorBars`), the others' left as they are. */
  seriesIndex = 0,
): ChartBlockAttrsV2 | null {
  if (!chartElementAvailability(attrs).errorBars) return null;
  const stated = recipe === null ? null : normalizeChartErrorBars(recipe);
  if (stated === undefined) return null;
  if (seriesIndex > 0) {
    const listed = "series" in attrs.data ? (attrs.data.series as readonly unknown[]).length : 0;
    if (!Number.isInteger(seriesIndex) || seriesIndex >= listed) return null;
    const entries = attrs.format?.seriesErrorBars ?? [];
    const own = entries.find((entry) => entry.seriesIndex === seriesIndex);
    if (JSON.stringify(own?.errorBars ?? null) === JSON.stringify(stated)) return null;
    const rest = entries.filter((entry) => entry.seriesIndex !== seriesIndex);
    const next = stated
      ? [...rest, { ...own, seriesIndex, errorBars: stated }].sort((left, right) => left.seriesIndex - right.seriesIndex)
      : rest;
    const format: ChartBlockFormat = { ...attrs.format };
    if (next.length) format.seriesErrorBars = next;
    else delete format.seriesErrorBars;
    return committed(attrs, format);
  }
  if (
    JSON.stringify(attrs.format?.errorBars ?? null) === JSON.stringify(stated) &&
    (stated !== null || (attrs.format?.errorBarsX === undefined && attrs.format?.seriesErrorBars === undefined))
  )
    return null;
  const format: ChartBlockFormat = { ...attrs.format };
  if (stated) format.errorBars = stated;
  else {
    // Off takes every error bar away, an X Y series' X bars and every other
    // series' bars too (Excel).
    delete format.errorBars;
    delete format.errorBarStyle;
    delete format.errorBarsX;
    delete format.errorBarXStyle;
    delete format.seriesErrorBars;
  }
  return committed(attrs, format);
}

type MutableSeriesStyle = {
  -readonly [K in keyof PresentationChartSeriesStyle]: PresentationChartSeriesStyle[K];
};

/** A line-drawing series: a 2-D line, a scatter, a radar, a combination's line. */
function chartSeriesDrawsLine(attrs: ChartBlockAttrsV2, seriesIndex: number): boolean {
  return (
    chartSeriesTakesMarker(
      attrs.chartType,
      attrs.chartSubtype,
      chartSeriesRenderAsOf(attrs, seriesIndex),
    ) ||
    (attrs.chartType === "radar" && attrs.chartSubtype === "filled")
  );
}

function chartWithSeriesStyle(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  edit: (style: MutableSeriesStyle) => void,
): ChartBlockAttrsV2 | null {
  const count =
    attrs.data.kind === "category" || attrs.data.kind === "xy"
      ? attrs.data.series.length
      : 0;
  if (!Number.isInteger(seriesIndex) || seriesIndex < 0 || seriesIndex >= count)
    return null;
  const styles = [...(attrs.format?.seriesStyles ?? [])];
  const at = styles.findIndex((style) => style.seriesIndex === seriesIndex);
  const before: PresentationChartSeriesStyle =
    at >= 0 ? styles[at] : { seriesIndex };
  const next: MutableSeriesStyle = { ...before };
  edit(next);
  if (JSON.stringify(next) === JSON.stringify(before)) return null;
  if (at >= 0) styles.splice(at, 1);
  if (Object.keys(next).length > 1) styles.push(next);
  styles.sort((left, right) => left.seriesIndex - right.seriesIndex);
  const format: ChartBlockFormat = { ...attrs.format };
  if (styles.length) format.seriesStyles = styles;
  else delete format.seriesStyles;
  return committed(attrs, format);
}

/**
 * Format Data Series ▸ Marker ▸ Marker Options on one series: a built-in
 * symbol (and size in points), `{ symbol: "none" }` for None, `null` for
 * Automatic (the marker its subtype draws). Null when that series draws no
 * marker (a column, an area, a bubble, a 3-D line, a filled radar), the
 * marker is not one Office states, or it changes nothing.
 */
export function chartWithSeriesMarker(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  marker: ChartSeriesMarker | null,
): ChartBlockAttrsV2 | null {
  if (
    !chartSeriesTakesMarker(
      attrs.chartType,
      attrs.chartSubtype,
      chartSeriesRenderAsOf(attrs, seriesIndex),
    )
  )
    return null;
  const stated = marker === null ? null : normalizeChartSeriesMarker(marker);
  if (stated === undefined) return null;
  return chartWithSeriesStyle(attrs, seriesIndex, (style) => {
    if (stated) style.marker = stated;
    else delete style.marker;
  });
}

/**
 * ▸ Line ▸ Smoothed line on one series: `true` / `false` state its own
 * `c:smooth`, which overrides the chart-wide `format.curve` for it; `null`
 * lets it follow the chart again. Null when that series has no curve (only a
 * line, a scatter and a combination's line do) or nothing changes.
 */
export function chartWithSeriesSmooth(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  smooth: boolean | null,
): ChartBlockAttrsV2 | null {
  if (
    !chartSeriesTakesSmooth(
      attrs.chartType,
      attrs.chartSubtype,
      chartSeriesRenderAsOf(attrs, seriesIndex),
    )
  )
    return null;
  return chartWithSeriesStyle(attrs, seriesIndex, (style) => {
    if (smooth === null) delete style.smooth;
    else style.smooth = smooth;
  });
}

/**
 * ▸ Line ▸ Width on one line-drawing series, in points: its `c:spPr/a:ln/@w`,
 * stated as `shapePaint.outlineWidthPt`. The model states a width with its
 * line's colour, so a series that states none takes the colour it is drawn
 * in now (its fill, else its Change Colors or style colour). `null` drops the
 * width (and a shadow or gradient riding with it keeps nothing to ride on,
 * so `shapePaint` goes with it only when it held the width alone).
 */
export function chartWithSeriesLineWidth(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  widthPt: number | null,
): ChartBlockAttrsV2 | null {
  if (!chartSeriesDrawsLine(attrs, seriesIndex)) return null;
  if (
    widthPt !== null &&
    (!Number.isFinite(widthPt) || widthPt < 0 || widthPt > MAX_CHART_OUTLINE_WIDTH_PT)
  )
    return null;
  return chartWithSeriesStyle(attrs, seriesIndex, (style) => {
    if (widthPt === null) {
      if (!style.shapePaint) return;
      const { outlineWidthPt: _width, ...rest } = style.shapePaint;
      if (Object.keys(rest).length) {
        // Office's default width under what else the paint states.
        style.shapePaint = normalizeOfficeChartShapePaint({ ...rest, outlineWidthPt: 2.25 });
      } else delete style.shapePaint;
      return;
    }
    if (style.outline === undefined) {
      const palette = attrs.format?.colorPalette;
      const preset = chartPreset(attrs.stylePreset).palette;
      style.outline = (
        style.fill ??
        (palette
          ? chartColorPaletteColor(
              palette,
              DEFAULT_CHART_THEME_ACCENTS,
              seriesIndex,
              chartColorTargetCount(attrs),
            )
          : preset[seriesIndex % preset.length])
      ).toLowerCase() as MutableSeriesStyle["outline"];
    }
    style.shapePaint = normalizeOfficeChartShapePaint({
      ...style.shapePaint,
      outlineWidthPt: widthPt,
    });
  });
}

export function chartWithUpDownBars(
  attrs: ChartBlockAttrsV2,
  visible: boolean,
): ChartBlockAttrsV2 | null {
  if (!chartElementAvailability(attrs).upDownBars) return null;
  if ((attrs.format?.upDownBars === true) === visible) return null;
  const format: ChartBlockFormat = { ...attrs.format };
  if (visible) format.upDownBars = true;
  else {
    delete format.upDownBars;
    delete format.upBarStyle;
    delete format.downBarStyle;
  }
  return committed(attrs, format);
}

const LINE_KEYS = {
  drop: ["dropLines", "dropLineStyle"],
  highLow: ["highLowLines", "highLowLineStyle"],
  series: ["seriesLines", "seriesLineStyle"],
} as const;

export function chartLineAvailable(
  attrs: ChartBlockAttrsV2,
  line: ChartElementLine,
): boolean {
  const available = chartElementAvailability(attrs);
  return line === "drop"
    ? available.dropLines
    : line === "highLow"
      ? available.highLowLines
      : available.seriesLines;
}

export function chartWithLine(
  attrs: ChartBlockAttrsV2,
  line: ChartElementLine,
  visible: boolean,
): ChartBlockAttrsV2 | null {
  if (!chartLineAvailable(attrs, line)) return null;
  const [key, styleKey] = LINE_KEYS[line];
  if ((attrs.format?.[key] === true) === visible) return null;
  const format: ChartBlockFormat = { ...attrs.format };
  if (visible) format[key] = true;
  else {
    delete format[key];
    delete format[styleKey];
  }
  return committed(attrs, format);
}

/**
 * One Add Chart Element pick, before any editor's model is touched.
 *
 * Word, Excel and PowerPoint draw one Add Chart Element menu (Microsoft
 * Support, "Change the layout or style of a chart", "Add or remove titles in a
 * chart", "Display or hide chart gridlines": Chart Design ▸ Add Chart Element ▸
 * element ▸ option). The shared menu (`OfficeChartElementMenuItems`) emits one
 * of these; a host whose chart IS a Chart V2 block (Docs, the deck's owned
 * chart) applies it with {@link chartWithElementIntent}, and the worksheet maps
 * it onto its own command ids so its command authority and receipts stay the
 * one gate a sheet edit passes.
 */
export type ChartElementIntent =
  | { element: "chartTitle"; on: boolean; text: string }
  | { element: "axisTitle"; axis: "x" | "y" | "z"; on: boolean; text: string }
  | { element: "axis"; axis: "x" | "y" | "y2" | "z"; on: boolean }
  | { element: "gridlines"; axis: "x" | "y"; on: boolean; minor?: boolean }
  | {
      element: "dataLabels";
      on: boolean;
      placement: ChartDataLabelPlacement | null;
      /** Data Labels > Data Callout. */
      callout?: boolean;
    }
  | {
      /** Format Data Labels on one series or one label (chartWithDataLabelOverride). */
      element: "dataLabelOverride";
      target: ChartDataLabelTarget;
      patch: ChartDataLabelOverridePatch | null;
    }
  | { element: "dataTable"; on: boolean; legendKeys: boolean }
  /** Format Data Table ▸ Table Borders: Horizontal, Vertical, Outline. */
  | { element: "dataTableBorders"; horizontal?: boolean; vertical?: boolean; outline?: boolean }
  /** Format Legend ▸ Show the legend without overlapping the chart. */
  | { element: "legendOverlay"; overlay: boolean }
  /** Format Chart Area ▸ Border ▸ Rounded corners. */
  | { element: "roundedCorners"; on: boolean }
  /** Format Data Series ▸ Series Options of a pie or doughnut. */
  | { element: "pieOptions"; firstSliceAngle?: number; doughnutHoleSize?: number; pieExplosion?: number }
  /** Format Data Point ▸ Point Explosion of one slice. */
  | { element: "pointExplosion"; seriesIndex: number; pointIndex: number; explosion: number }
  | { element: "displayBlanksAs"; value: ChartDisplayBlanksAs }
  | { element: "view3D"; view: Partial<ChartView3D> | null }
  | { element: "errorBars"; recipe: ChartElementErrorBars | null; seriesIndex?: number }
  | { element: "legend"; position: NonNullable<ChartBlockFormat["legend"]> }
  | { element: "line"; line: ChartElementLine; on: boolean }
  | { element: "linesNone" }
  | {
      element: "trendline";
      trendline: ChartTrendline | null;
      /** A menu row's own forecast (Linear Forecast: 2; the others: none);
       * absent, the trendline keeps the forecast it has (the pane's type). */
      forward?: number;
      /** Format Data Series' trendline: that series' alone (Office bases a
       * trendline on the series picked); absent, the chart's. */
      seriesIndex?: number;
    }
  | {
      element: "trendlineOption";
      change: ChartTrendlineOptionChange;
      seriesIndex?: number;
    }
  | { element: "upDownBars"; on: boolean }
  | { element: "formatOption"; option: ChartFormatOption }
  | {
      /** Format Data Series ▸ Plot Series On (`chartWithSeriesAxis`). */
      element: "seriesAxis";
      seriesIndex: number;
      axis: "primary" | "secondary";
    }
  | {
      /** Format Data Series > Series Options of a column or bar chart:
       * Series Overlap (-100..100 %) and Gap Width (0..500 %), Office's
       * `c:overlap` / `c:gapWidth`, kept on the chart's layout. */
      element: "seriesOptions";
      overlap?: number;
      gapWidth?: number;
      /** A 3-D chart's Gap Depth (0..500 %, `c:gapDepth`). */
      gapDepth?: number;
      /** A 3-D column or bar chart's Column shape: of the series
       * `seriesIndex` names (its `c:shape`), else of every series. */
      barShape?: ChartBarShape;
      seriesIndex?: number;
    };

/** The legend position a chart shows (its own, else its style's default). */
export function chartLegendShown(
  attrs: ChartBlockAttrsV2,
): NonNullable<ChartBlockFormat["legend"]> {
  return attrs.format?.legend ?? chartPreset(attrs.stylePreset).legendPosition;
}

/** Lines ▸ None: every predefined line off, each with its own format. */
export function chartWithoutLines(
  attrs: ChartBlockAttrsV2,
): ChartBlockAttrsV2 | null {
  const format: ChartBlockFormat = { ...attrs.format };
  let changed = false;
  for (const [key, styleKey] of Object.values(LINE_KEYS)) {
    if (format[key] !== undefined || format[styleKey] !== undefined) {
      changed = true;
      delete format[key];
      delete format[styleKey];
    }
  }
  return changed ? committed(attrs, format) : null;
}

/** The pick applied to a Chart V2 block, or `null` when it changes nothing
 * or the chart cannot carry it. */
/** Whether the Add Chart Element menu offers this pick on this chart
 * (the rest it greys out). */
export function chartElementIntentAvailable(
  attrs: ChartBlockAttrsV2,
  intent: ChartElementIntent,
): boolean {
  const available = chartElementAvailability(attrs);
  switch (intent.element) {
    case "chartTitle":
    case "legend":
      return true;
    case "axisTitle":
      // Axis Titles > Depth: a 3-D chart with a depth (series) axis.
      return intent.axis === "z"
        ? chartHasDepthAxis(attrs.chartType, attrs.chartSubtype)
        : chartSupportsAxisTitle(attrs.chartType, intent.axis);
    case "axis":
      return intent.axis === "y2"
        ? chartHasSecondaryValueAxis(attrs.chartType, attrs.chartSubtype, attrs.data)
        : intent.axis === "z"
          ? chartHasDepthAxis(attrs.chartType, attrs.chartSubtype)
          : chartHasAxis(attrs.chartType, intent.axis);
    case "gridlines":
      if (!intent.minor) return available.gridlines;
      return (
        chartSupportsMinorGridlines(attrs.chartType, attrs.chartSubtype) ||
        officeChartLayoutOf({ ...attrs, layout: attrs.format?.layout }) !== null
      );
    case "dataLabels":
      if (!available.dataLabels) return false;
      if (intent.callout) return intent.on && available.dataLabelCallout;
      return !intent.on || !intent.placement || chartDataLabelPlacements(attrs).includes(intent.placement);
    case "dataLabelOverride":
      return available.dataLabels && available.dataLabelRegionLayouts.length === 0;
    case "dataTable":
    case "dataTableBorders":
      return available.dataTable;
    case "pieOptions":
    case "pointExplosion":
      return attrs.chartType === "pie";
    case "legendOverlay":
      return chartLegendShown(attrs) !== "none";
    case "roundedCorners":
      return true;
    case "displayBlanksAs":
      return chartDisplayBlanksOptions(attrs).includes(intent.value);
    case "view3D":
      return chartTakesView3D(attrs.chartType, attrs.chartSubtype);
    case "errorBars":
      return available.errorBars;
    case "line":
      return chartLineAvailable(attrs, intent.line);
    case "linesNone":
      return available.dropLines || available.highLowLines || available.seriesLines;
    case "trendline":
    case "trendlineOption":
      return available.trendline;
    case "upDownBars":
      return available.upDownBars;
    case "seriesOptions":
      return chartSeriesOptions(attrs) !== null || chart3DSeriesOptions(attrs) !== null;
    case "seriesAxis":
      return chartSeriesPlottedOn(attrs, intent.seriesIndex) !== null;
    case "formatOption":
      return chartWithFormatOption(attrs, intent.option) !== null;
  }
}

/**
 * The pick applied, `"unchanged"` when the chart already shows it, or null
 * when the menu would grey it out on this chart: what an AI button reports
 * (applied / unchanged / not applicable), which {@link chartWithElementIntent}
 * alone folds into one null.
 */
export function chartElementIntentOutcome(
  attrs: ChartBlockAttrsV2,
  intent: ChartElementIntent,
): ChartBlockAttrsV2 | "unchanged" | null {
  if (!chartElementIntentAvailable(attrs, intent)) return null;
  return chartWithElementIntent(attrs, intent) ?? "unchanged";
}

/**
 * The chart with the hand placements of elements it no longer draws let go
 * (Office drops a deleted title's or legend's layout with it), and the
 * legend's placement given back to the chart when a ribbon place is chosen
 * for it ("Placement presets also discard any custom sizing").
 */
export function chartWithElementLayoutKept(
  attrs: ChartBlockAttrsV2,
  previous?: ChartBlockAttrsV2,
): ChartBlockAttrsV2 {
  const layout = attrs.format?.elementLayout;
  const hidden = attrs.format?.legendHiddenEntries;
  const legendLook = attrs.format?.legendStyle !== undefined || attrs.format?.legendOverlay !== undefined;
  if (!layout && !hidden && !(legendLook && chartLegendShown(attrs) === "none")) return attrs;
  const next: Record<string, unknown> = { ...(layout ?? {}) };
  if (!attrs.title && attrs.format?.automaticTitle !== true) delete next.title;
  const xRole = officeChartAxisRole(attrs, "X");
  const axisTitle = (role: "category" | "value") => (xRole === role ? attrs.axisTitleX : attrs.axisTitleY);
  if (!axisTitle("category")) delete next.axisTitleCategory;
  if (!axisTitle("value")) delete next.axisTitleValue;
  if (!attrs.format?.seriesAxisTitle) delete next.axisTitleSeries;
  const legend = chartLegendShown(attrs);
  const legendMoved = previous !== undefined && chartLegendShown(previous) !== legend;
  if (legend === "none" || legendMoved) delete next.legend;
  const format: ChartBlockFormat = { ...attrs.format };
  const kept = normalizeChartElementLayout(next);
  if (kept) format.elementLayout = kept;
  else delete format.elementLayout;
  if (legend === "none") {
    // Legend > None deletes the legend whole, its look with it: shown again,
    // PowerPoint draws the chart style's own legend, at its place, every
    
    
    delete format.legendHiddenEntries;
    delete format.legendStyle;
    delete format.legendOverlay;
  }
  if (JSON.stringify(format) === JSON.stringify(attrs.format ?? {})) return attrs;
  return { ...attrs, format: Object.keys(format).length ? format : undefined };
}

/**
 * The chart with an element dragged (or sized) by hand to `box`, fractions of
 * the chart: the title and the axis titles keep their place, the legend and
 * the plot area their place and size (Office's `c:manualLayout`). Null when
 * nothing changes or the element is not one a hand places.
 */
export function chartWithPlacedElement(
  attrs: ChartBlockAttrsV2,
  part: OfficeChartPicturePart,
  box: { x: number; y: number; w: number; h: number },
  /** One data label dragged: whose it is and where it was (fractions). */
  label?: { label?: { series: number; point: number }; from?: { x: number; y: number } },
): ChartBlockAttrsV2 | null {
  if (part === "dataLabels") {
    // Office keeps a dragged label's offset from where the chart puts it,
    // a fraction of the chart (`c:dLbl/c:layout/c:manualLayout`).
    if (!label?.label || !label.from) return null;
    const target = { seriesIndex: label.label.series, pointIndex: label.label.point };
    const prior = chartDataLabelOverrideOf(attrs, target)?.offset ?? { x: 0, y: 0 };
    const clamp = (value: number) => Math.round(Math.min(1, Math.max(-1, value)) * 1e6) / 1e6;
    // Stated along the value axis' own direction, as it is drawn.
    const signs = chartDataLabelOffsetSigns(attrs);
    const offset = {
      x: clamp(prior.x + (box.x - label.from.x) * signs.x),
      y: clamp(prior.y + (box.y - label.from.y) * signs.y),
    };
    return chartWithDataLabelOverride(attrs, target, { offset: offset.x === 0 && offset.y === 0 ? undefined : offset });
  }
  if (!chartTakesElementLayout(attrs.chartType)) return null;
  let key: ChartElementLayoutKey;
  switch (part) {
    case "title":
      if (!attrs.title && attrs.format?.automaticTitle !== true) return null;
      key = "title";
      break;
    case "legend":
      if (chartLegendShown(attrs) === "none") return null;
      key = "legend";
      break;
    case "plot":
      key = "plotArea";
      break;
    case "axisX":
    case "axisY":
      key = officeChartAxisRole(attrs, part === "axisX" ? "X" : "Y") === "category" ? "axisTitleCategory" : "axisTitleValue";
      break;
    default:
      return null;
  }
  const sized = key === "legend" || key === "plotArea";
  const layout = withChartElementPlace(attrs.format?.elementLayout, key, sized ? box : { x: box.x, y: box.y });
  if (JSON.stringify(layout ?? null) === JSON.stringify(attrs.format?.elementLayout ?? null)) return null;
  const format: ChartBlockFormat = { ...attrs.format };
  if (layout) format.elementLayout = layout;
  else delete format.elementLayout;
  return committed(attrs, format);
}

/**
 * The chart with its selected element deleted, as Office's Delete key does
 * with a chart element selected: the title, an axis title, the legend, the
 * data labels, an axis (Excel hides the axis Delete is pressed on, as Chart
 * Elements > Axes does, and the chart keeps its gridlines) -- and inside the
 * legend one entry (`legendHiddenEntries`), of the labels one series' or one
 * label; one point (a slice) selected, its own formatting and label
 * (Office's Delete on one point clears what was stated for it, the point
 * stays). "keep" for an element Delete leaves alone here (the plot area): the
 * chart itself stays. A whole series selected is the host's to remove
 * (`officeChartDeletedSeries`, `chartWithoutSeries`: its data binding is the
 * host's); here it is "keep". Null when the chart area is what is selected
 * (Delete takes the whole chart).
 */
export function chartWithoutSelectedElement(
  attrs: ChartBlockAttrsV2,
  area: string | null,
  detail: { legendEntry?: number; labels?: { series: number; point?: number } } | null,
): ChartBlockAttrsV2 | "keep" | null {
  if (!area || area === "chartAreaStyle") return null;
  if (detail?.legendEntry !== undefined && area === "legendStyle") {
    if (!chartTakesElementLayout(attrs.chartType)) return "keep";
    const entry = chartLegendEntryOfPosition(attrs)(detail.legendEntry);
    const hidden = [...new Set([...(attrs.format?.legendHiddenEntries ?? []), entry])].sort((a, b) => a - b);
    return committed(attrs, { ...attrs.format, legendHiddenEntries: hidden }) ?? "keep";
  }
  if (detail?.labels) {
    const target = detail.labels.point === undefined
      ? { seriesIndex: detail.labels.series }
      : { seriesIndex: detail.labels.series, pointIndex: detail.labels.point };
    return chartWithDataLabelOverride(attrs, target, { dataLabels: false }) ?? "keep";
  }
  const slice = /^slice:(\d+)$/.exec(area);
  const point = officeChartPointArea(area) ?? (slice ? { seriesIndex: 0, pointIndex: Number(slice[1]) } : null);
  if (point) return chartWithoutPointFormat(attrs, point.seriesIndex, point.pointIndex) ?? "keep";
  const axisTitle = /^axisTitle:(category|value)$/.exec(area);
  const axis = /^axis:(category|value)$/.exec(area);
  const intent: ChartElementIntent | null =
    area === "chartTitleStyle"
      ? { element: "chartTitle", on: false, text: "" }
      : area === "legendStyle"
        ? { element: "legend", position: "none" }
        : area === "dataLabelStyle"
          ? { element: "dataLabels", on: false, placement: null }
          : axisTitle
            ? {
                element: "axisTitle",
                axis: officeChartAxisRole(attrs, "X") === axisTitle[1] ? "x" : "y",
                on: false,
                text: "",
              }
            : axis
              ? { element: "axis", axis: officeChartAxisRole(attrs, "X") === axis[1] ? "x" : "y", on: false }
              : null;
  if (!intent) return "keep";
  return chartWithElementIntent(attrs, intent) ?? "keep";
}

/**
 * One point's own formatting and label gone (Delete on one selected point or
 * slice): its paint (`sliceStyles`) and its label statement
 * (`dataLabelOverrides`); null when it states neither.
 */
export function chartWithoutPointFormat(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  pointIndex: number,
): ChartBlockAttrsV2 | null {
  const format = attrs.format;
  if (!format) return null;
  const own = (entry: { seriesIndex?: number; pointIndex?: number }) =>
    (entry.seriesIndex ?? 0) === seriesIndex && entry.pointIndex === pointIndex;
  const slices = format.sliceStyles?.filter((entry) => !own(entry));
  const labels = format.dataLabelOverrides?.filter((entry) => !own(entry));
  if (slices?.length === format.sliceStyles?.length && labels?.length === format.dataLabelOverrides?.length) return null;
  const next: ChartBlockFormat = { ...format };
  if (slices?.length) next.sliceStyles = slices;
  else delete next.sliceStyles;
  if (labels?.length) next.dataLabelOverrides = labels;
  else delete next.dataLabelOverrides;
  return committed(attrs, next);
}

/**
 * The series a Delete takes: a whole series selected (one click on a column,
 * a line, a bar), not its data labels (`labels`) nor one of its points.
 */
export function officeChartDeletedSeries(
  area: string | null,
  detail: { labels?: unknown } | null,
): number | null {
  if (!area || detail?.labels) return null;
  const match = /^series:(\d+)$/.exec(area);
  return match ? Number(match[1]) : null;
}

/**
 * The chart without one series, as Office's Delete on a selected series and
 * Select Data > Remove make it (the series leaves the chart; its cells stay
 * in the sheet): the format follows the other series to their new places
 * (`chartFormatFollowingDataEdit`), the removed series' paint, labels,
 * trendlines and error bars go with it, and a legend entry hidden by hand
 * follows its series. Null where Select Data's Remove is off: the last series
 * (two for a combination), a stock chart's fixed roles, a kind with no series.
 */
export function chartWithoutSeries(attrs: ChartBlockAttrsV2, seriesIndex: number): ChartBlockAttrsV2 | null {
  const data = attrs.data;
  if (data.kind !== "category" && data.kind !== "xy" && data.kind !== "sample") return null;
  const count = data.series.length;
  const minimum = attrs.chartType === "combo" ? 2 : 1;
  if (!Number.isInteger(seriesIndex) || seriesIndex < 0 || seriesIndex >= count || count <= minimum) return null;
  const origins = Array.from({ length: count }, (_, index) => index).filter((index) => index !== seriesIndex);
  const nextData = { ...data, series: data.series.filter((_, index) => index !== seriesIndex) } as typeof data;
  let format = chartFormatFollowingDataEdit(attrs.format, origins, null);
  // A legend entry is a series on every chart but a pie (whose entries are
  // its categories).
  if (format?.legendHiddenEntries && attrs.chartType !== "pie") {
    const hidden = format.legendHiddenEntries
      .filter((entry) => entry !== seriesIndex)
      .map((entry) => (entry > seriesIndex ? entry - 1 : entry));
    format = { ...format };
    if (hidden.length) format.legendHiddenEntries = hidden;
    else delete format.legendHiddenEntries;
  }
  const next: ChartBlockAttrsV2 = { ...attrs, data: nextData };
  if (format && Object.keys(format).length) next.format = format;
  else delete next.format;
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

export function chartWithElementIntent(
  attrs: ChartBlockAttrsV2,
  intent: ChartElementIntent,
): ChartBlockAttrsV2 | null {
  const next = chartWithElementIntentAlone(attrs, intent);
  return next ? chartWithElementLayoutKept(next, attrs) : next;
}

function chartWithElementIntentAlone(
  attrs: ChartBlockAttrsV2,
  intent: ChartElementIntent,
): ChartBlockAttrsV2 | null {
  switch (intent.element) {
    case "chartTitle": {
      const title = intent.on ? clampChartText(intent.text.trim()) : "";
      // Office's automatic title (`format.automaticTitle`) is a title shown:
      // turning the title on keeps it, turning it off (or typing words of the
      // chart's own) ends it.
      const automatic = !attrs.title && attrs.format?.automaticTitle === true;
      if (automatic && intent.on && !title) return null;
      if ((attrs.title ?? "") === title && !automatic) return null;
      // The runs' format follows the words typed over them.
      const { titleRuns: _previous, ...rest } = attrs;
      const titleRuns = fitChartTitleRuns(attrs.titleRuns, title);
      const next: ChartBlockAttrsV2 = { ...rest, title, ...(titleRuns ? { titleRuns } : {}) };
      if (attrs.format?.automaticTitle) {
        const { automaticTitle: _automatic, ...format } = attrs.format;
        next.format = Object.keys(format).length ? format : undefined;
        if (!next.format) delete next.format;
      }
      return validateChartBlockAttrsV2(next).valid ? next : null;
    }
    case "axisTitle": {
      if (intent.axis === "z") {
        // Axis Titles > Depth (`format.seriesAxisTitle`, `c:serAx/c:title`).
        if (!chartHasDepthAxis(attrs.chartType, attrs.chartSubtype)) return null;
        if (Boolean(attrs.format?.seriesAxisTitle) === intent.on) return null;
        const format: ChartBlockFormat = { ...attrs.format };
        if (intent.on) format.seriesAxisTitle = clampChartText(intent.text.trim()) || "Axis Title";
        else delete format.seriesAxisTitle;
        // A title taken off lets its hand placement go with it.
        const next = chartWithElementLayoutKept({ ...attrs, format: Object.keys(format).length ? format : undefined });
        return validateChartBlockAttrsV2(next).valid ? next : null;
      }
      if (!chartSupportsAxisTitle(attrs.chartType, intent.axis)) return null;
      const key = intent.axis === "x" ? "axisTitleX" : "axisTitleY";
      if (Boolean(attrs[key]) === intent.on) return null;
      const next: ChartBlockAttrsV2 = { ...attrs };
      if (intent.on) next[key] = clampChartText(intent.text.trim());
      else delete next[key];
      return validateChartBlockAttrsV2(next).valid ? next : null;
    }
    case "axis":
      return chartWithAxis(attrs, intent.axis, intent.on);
    case "gridlines": {
      if (!intent.minor) return chartWithGridlines(attrs, intent.axis, intent.on);
      if (chartSupportsMinorGridlines(attrs.chartType, attrs.chartSubtype)) {
        // Any other classic chart states them as its own format keys, which
        // every classic writer turns into `c:minorGridlines` on that axis.
        const key = intent.axis === "x" ? "minorGridlinesX" : "minorGridlinesY";
        if ((attrs.format?.[key] === true) === intent.on) return null;
        const format: ChartBlockFormat = { ...attrs.format };
        if (intent.on) format[key] = true;
        else delete format[key];
        return committed(attrs, format);
      }
      // Minor gridlines are a calibrated layout's (`c:minorGridlines` on the
      // axis), which every classic writer states for a 2-D column / bar
      // chart; other charts list them greyed with that reason.
      const layout = officeChartLayoutOf({
        ...attrs,
        layout: attrs.format?.layout,
      });
      const key = intent.axis === "x" ? "minorGridlinesX" : "minorGridlinesY";
      if (!layout || layout[key] === intent.on) return null;
      return committed(attrs, {
        ...attrs.format,
        layout: { ...layout, [key]: intent.on },
      });
    }
    case "dataLabels":
      return chartWithDataLabels(attrs, intent.on, intent.placement, intent.callout === true);
    case "dataLabelOverride":
      return chartWithDataLabelOverride(attrs, intent.target, intent.patch);
    case "dataTable":
      return chartWithDataTable(attrs, intent.on, intent.legendKeys);
    case "dataTableBorders":
      return chartWithDataTableBorders(attrs, intent);
    case "pieOptions":
      return chartWithPieOptions(attrs, intent);
    case "roundedCorners": {
      if ((attrs.format?.roundedCorners === true) === intent.on) return null;
      const format: ChartBlockFormat = { ...attrs.format };
      if (intent.on) format.roundedCorners = true;
      else delete format.roundedCorners;
      return committed(attrs, format);
    }
    case "legendOverlay": {
      if (chartLegendShown(attrs) === "none" || (attrs.format?.legendOverlay === true) === intent.overlay) return null;
      const format: ChartBlockFormat = { ...attrs.format };
      if (intent.overlay) format.legendOverlay = true;
      else delete format.legendOverlay;
      return committed(attrs, format);
    }
    case "pointExplosion":
      return chartWithPointExplosion(attrs, intent);
    case "displayBlanksAs":
      return chartWithDisplayBlanksAs(attrs, intent.value);
    case "view3D":
      return chartWithView3D(attrs, intent.view);
    case "errorBars":
      return chartWithErrorBars(attrs, intent.recipe, intent.seriesIndex ?? 0);
    case "legend":
      return chartLegendShown(attrs) === intent.position
        ? null
        : committed(attrs, { ...attrs.format, legend: intent.position });
    case "line":
      return chartWithLine(attrs, intent.line, intent.on);
    case "linesNone":
      return chartWithoutLines(attrs);
    case "trendline": {
      const own =
        intent.seriesIndex === undefined
          ? attrs.format?.trendline
          : chartSeriesTrendline(attrs, intent.seriesIndex);
      let trendline = intent.trendline && chartTrendlineOfKind(own, intent.trendline);
      if (trendline && intent.forward !== undefined) {
        const next: Record<string, unknown> = { ...trendline };
        if (intent.forward > 0) next.forward = intent.forward;
        else delete next.forward;
        trendline = normalizeChartTrendline(next) ?? null;
        if (!trendline) return null;
      }
      return intent.seriesIndex === undefined
        ? chartWithTrendline(attrs, trendline)
        : chartWithSeriesTrendline(attrs, intent.seriesIndex, trendline);
    }
    case "trendlineOption":
      return chartWithTrendlineOption(attrs, intent.change, intent.seriesIndex);
    case "upDownBars":
      return chartWithUpDownBars(attrs, intent.on);
    case "seriesAxis":
      return chartWithSeriesAxis(attrs, intent.seriesIndex, intent.axis);
    case "formatOption":
      return chartWithFormatOption(attrs, intent.option);
    case "seriesOptions":
      return chartTakesGapDepth(attrs.chartType, attrs.chartSubtype)
        ? chartWith3DSeriesOptions(attrs, intent)
        : chartWithSeriesOptions(attrs, intent);
  }
}

/**
 * Format Data Series > Series Options of a 3-D column, bar, line or area
 * chart (Excel: Gap Depth, Gap Width and Column shape; ECMA-376
 * `c:gapDepth`, `c:gapWidth`, ST_Shape): what the chart draws with, its
 * defaults (150 %, 150 %, Box) where it states none, `barShape` being the
 * series `seriesIndex`'s (its own over the chart's). Null for any other chart.
 * A 3-D chart has no Series Overlap (`c:bar3DChart` holds none).
 */
export function chart3DSeriesOptions(
  attrs: ChartBlockAttrsV2,
  seriesIndex?: number,
): { gapDepth: number; gapWidth?: number; barShape?: ChartBarShape } | null {
  if (!chartTakesGapDepth(attrs.chartType, attrs.chartSubtype)) return null;
  const format = attrs.format;
  if (!chartTakesBarShape(attrs.chartType, attrs.chartSubtype)) return { gapDepth: format?.gapDepth ?? 150 };
  const own =
    seriesIndex === undefined ? undefined : format?.seriesStyles?.find((style) => style.seriesIndex === seriesIndex)?.barShape;
  return {
    gapDepth: format?.gapDepth ?? 150,
    gapWidth: format?.barGapWidth ?? 150,
    barShape: own ?? (isChartBarShape(format?.barShape) ? format.barShape : "box"),
  };
}

/**
 * A 3-D chart after a Series Options change: Gap Depth and Gap Width for the
 * chart (whole percents clamped to 0..500, absent at 150), a Column shape for
 * the series `seriesIndex` names -- as Excel's pane applies it to the
 * selected series -- or, without one, for every series (the series' own
 * shapes then give way to it). Null when nothing changes or the chart has no
 * such option.
 */
export function chartWith3DSeriesOptions(
  attrs: ChartBlockAttrsV2,
  change: { gapDepth?: number; gapWidth?: number; barShape?: ChartBarShape; seriesIndex?: number },
): ChartBlockAttrsV2 | null {
  const current = chart3DSeriesOptions(attrs, change.seriesIndex);
  if (!current) return null;
  const whole = (value: number | undefined, fallback: number) =>
    value === undefined || !Number.isFinite(value) ? fallback : Math.min(500, Math.max(0, Math.round(value)));
  const format: ChartBlockFormat = { ...attrs.format };
  const gapDepth = whole(change.gapDepth, current.gapDepth);
  if (gapDepth === 150) delete format.gapDepth;
  else format.gapDepth = gapDepth;
  if (current.gapWidth !== undefined) {
    const gapWidth = whole(change.gapWidth, current.gapWidth);
    if (gapWidth === 150) delete format.barGapWidth;
    else format.barGapWidth = gapWidth;
  }
  if (change.barShape !== undefined && current.barShape !== undefined && isChartBarShape(change.barShape)) {
    const series = (format.seriesStyles ?? []).map((style) => ({ ...style }));
    if (change.seriesIndex !== undefined) {
      const chartShape = isChartBarShape(format.barShape) ? format.barShape : "box";
      const at = series.findIndex((style) => style.seriesIndex === change.seriesIndex);
      const row = at >= 0 ? series[at] : { seriesIndex: change.seriesIndex };
      const { barShape: _old, ...rest } = row as PresentationChartSeriesStyle;
      void _old;
      const next = change.barShape === chartShape ? rest : { ...rest, barShape: change.barShape };
      const keeps = Object.keys(next).some((key) => key !== "seriesIndex");
      if (at >= 0) {
        if (keeps) series[at] = next as PresentationChartSeriesStyle;
        else series.splice(at, 1);
      } else if (keeps) series.push(next as PresentationChartSeriesStyle);
    } else {
      if (change.barShape === "box") delete format.barShape;
      else format.barShape = change.barShape;
      for (let index = series.length - 1; index >= 0; index -= 1) {
        const { barShape: _old, ...rest } = series[index];
        void _old;
        if (Object.keys(rest).some((key) => key !== "seriesIndex")) series[index] = rest as PresentationChartSeriesStyle;
        else series.splice(index, 1);
      }
    }
    if (series.length) format.seriesStyles = series.sort((a, b) => a.seriesIndex - b.seriesIndex);
    else delete format.seriesStyles;
  }
  if (JSON.stringify(format) === JSON.stringify(attrs.format ?? {})) return null;
  return committed(attrs, format);
}

/** Whether a chart's Gap Width / Series Overlap live in `format.barGapWidth`
 * / `barOverlap` rather than its calibrated layout: a 100% stacked chart and
 * a combination always; a clustered or stacked one that states them so and
 * has no layout (a deck's chart read from PowerPoint's part). */
export function chartSpacingInOwnKeys(attrs: ChartBlockAttrsV2): boolean {
  if (!chartTakesBarSpacing(attrs.chartType, attrs.chartSubtype)) return false;
  if (attrs.chartSubtype === "percentStacked" || attrs.chartType === "combo") return true;
  return (
    attrs.format?.layout === undefined &&
    (attrs.format?.barGapWidth !== undefined || attrs.format?.barOverlap !== undefined)
  );
}

/** The Series Options a column or bar chart draws its bars with (Office's
 * defaults when it states none), or null for a chart that has none. */
export function chartSeriesOptions(
  attrs: ChartBlockAttrsV2,
): { overlap: number; gapWidth: number } | null {
  if (chartSpacingInOwnKeys(attrs)) {
    // A 100% stacked chart, a combination's columns and a chart that states
    // its spacing without a calibrated layout keep them in their own keys
    // (absent: the spacing the writer states for the chart).
    const defaults = chartBarSpacingDefaults(attrs.chartType, attrs.chartSubtype, attrs.format?.comboColumnGrouping);
    return {
      overlap: attrs.format?.barOverlap ?? defaults.overlap,
      gapWidth: attrs.format?.barGapWidth ?? defaults.gapWidth,
    };
  }
  const layout = officeChartLayoutOf({ ...attrs, layout: attrs.format?.layout });
  return layout ? { overlap: layout.overlap, gapWidth: layout.gapWidth } : null;
}

/**
 * Format Data Series > Series Options (Microsoft Support, "Change the
 * spacing between bars in a column or bar chart": Series Overlap and Gap
 * Width, for every series of the chart). A clustered or stacked column / bar
 * chart keeps them on its layout, which every classic writer states as
 * `c:overlap` and `c:gapWidth`; a 100% stacked one, which takes no layout,
 * in `format.barOverlap` / `format.barGapWidth`, absent at Office's 100 % /
 * 150 %. The values are whole percents, clamped to Office's ranges.
 */
export function chartWithSeriesOptions(
  attrs: ChartBlockAttrsV2,
  change: { overlap?: number; gapWidth?: number },
): ChartBlockAttrsV2 | null {
  const whole = (value: number | undefined, min: number, max: number, current: number) =>
    value === undefined || !Number.isFinite(value) ? current : Math.min(max, Math.max(min, Math.round(value)));
  if (chartSpacingInOwnKeys(attrs)) {
    const current = chartSeriesOptions(attrs)!;
    const defaults = chartBarSpacingDefaults(attrs.chartType, attrs.chartSubtype, attrs.format?.comboColumnGrouping);
    const overlap = whole(change.overlap, -100, 100, current.overlap);
    const gapWidth = whole(change.gapWidth, 0, 500, current.gapWidth);
    if (overlap === current.overlap && gapWidth === current.gapWidth) return null;
    const format: ChartBlockFormat = { ...attrs.format };
    if (overlap === defaults.overlap) delete format.barOverlap;
    else format.barOverlap = overlap;
    if (gapWidth === defaults.gapWidth) delete format.barGapWidth;
    else format.barGapWidth = gapWidth;
    return committed(attrs, format);
  }
  const layout = officeChartLayoutOf({ ...attrs, layout: attrs.format?.layout });
  if (!layout) return null;
  const overlap = whole(change.overlap, -100, 100, layout.overlap);
  const gapWidth = whole(change.gapWidth, 0, 500, layout.gapWidth);
  if (overlap === layout.overlap && gapWidth === layout.gapWidth) return null;
  return committed(attrs, { ...attrs.format, layout: { ...layout, overlap, gapWidth } });
}
