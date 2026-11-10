/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeOfficeChartText } from "../office-charts/chart-text";
import { normalizeOfficeChartAxesOptions } from "../office-charts/axis-options";
import { spreadsheetChartTitle } from "./chart-title";
import {
  isChartBinIntervalClosed,
  isChartParentLabelLayout,
  normalizeChartBoxWhisker,
  normalizeChartHistogramBins,
  isChartDisplayBlanksAs,
  isChartDataLabelPosition,
  isChartDataLabelSeparator,
  isChartMapLabelLayout,
  type ChartBlockFormat,
} from "../../components/console/editor/extensions/chartBlockShared";
import { normalizeSpreadsheetChartLayout, type SpreadsheetChartLayout } from "./chart-quick-layout";
import {
  normalizePresentationChartAreaStyle,
  normalizePresentationChartAxisStyles,
  normalizePresentationChartAxisTitleStyles,
  normalizePresentationChartSeriesStyles,
  normalizePresentationChartSliceStyles,
  presentationChartAxisStylePart,
  type PresentationChartAreaStyle,
} from "../presentation-chart-area-style";

/** Chart semantics shared with Docs; worksheet binding and native ownership
 * deliberately stay outside this additive formatting envelope. */
export const SPREADSHEET_CHART_AREA_KEYS = [
  "chartAreaStyle", "plotAreaStyle", "chartTitleStyle", "legendStyle",
  // A 3-D chart's walls and floor (Format Back Wall / Side Wall / Floor).
  "backWallStyle", "sideWallStyle", "floorStyle",
  "gridlineStyle", "dataLabelStyle", "dataLabelLeaderLineStyle", "dataTableStyle",
  "gridlineStyleX", "gridlineStyleY", "minorGridlineStyleX", "minorGridlineStyleY",
  "trendlineStyle", "errorBarStyle", "upBarStyle", "downBarStyle",
  "dropLineStyle", "highLowLineStyle", "seriesLineStyle",
] as const satisfies readonly (keyof ChartBlockFormat)[];

const BOOLEAN_KEYS = [
  "dataLabels", "dataLabelValue", "dataLabelLegendKey", "dataLabelCategoryName",
  "dataLabelSeriesName", "dataLabelPercentage", "dataLabelBubbleSize", "dataLabelLeaderLines",
  "dataLabelCallout",
  "axisX", "axisY", "gridlinesX", "gridlinesY", "minorGridlinesX", "minorGridlinesY",
  "dataTable", "dataTableLegendKeys",
  "dataTableHorizontalBorder", "dataTableVerticalBorder", "dataTableOutline",
] as const satisfies readonly (keyof ChartBlockFormat)[];

export type SpreadsheetChartFormat = Pick<ChartBlockFormat,
  typeof SPREADSHEET_CHART_AREA_KEYS[number] | typeof BOOLEAN_KEYS[number] |
  "seriesStyles" | "sliceStyles" | "axisStyles" | "dataLabelPosition" | "dataLabelSeparator" |
  "dataLabelRegionLayout" | "binIntervalClosed" | "parentLabelLayout" | "histogramBins" | "boxWhisker" | "gapWidth" | "displayBlanksAs" | "legend" | "axisOptions" |
  "hiddenSeries" | "hiddenCategories" | "text" | "axisTitleStyles"> & { layout?: SpreadsheetChartLayout };
export type SpreadsheetChartFormatTarget = typeof SPREADSHEET_CHART_AREA_KEYS[number]
  | `series:${number}` | `slice:${number}` | "axis:category" | "axis:value";

const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
// `text` is an element's own font (a title's, a legend's, its labels'), which
// an imported chart states: left off this list, one such legend made the whole

const AREA_FIELDS = new Set(["fill", "outline", "effect", "textFill", "textOutline", "textEffect", "shapePaint", "textPaint", "text"]);

function exactArea(value: unknown): PresentationChartAreaStyle | undefined {
  if (!record(value) || Object.keys(value).some(key => !AREA_FIELDS.has(key))) return undefined;
  const present = Object.fromEntries(Object.entries(value).filter(([, atom]) => atom !== undefined));
  const normalized = normalizePresentationChartAreaStyle(present);
  return normalized && Object.keys(present).length === Object.keys(normalized).length ? normalized : undefined;
}

/** Never turn an unrecognised native/custom atom into a different supported
 * style. Callers retain its original envelope and disable the affected writer. */
export function normalizeSpreadsheetChartFormat(value: unknown): SpreadsheetChartFormat | undefined {
  if (!record(value)) return undefined;
  const result: Record<string, unknown> = {};
  for (const [key, atom] of Object.entries(value)) {
    if (key === "layout") {
      const layout=normalizeSpreadsheetChartLayout(atom); if(!layout)return undefined; result.layout=layout;
    } else if ((SPREADSHEET_CHART_AREA_KEYS as readonly string[]).includes(key)) {
      const area = exactArea(atom);
      if (!area) return undefined;
      result[key] = area;
    } else if ((BOOLEAN_KEYS as readonly string[]).includes(key)) {
      if (typeof atom !== "boolean") return undefined;
      result[key] = atom;
    } else if (key === "legend") {
      if (!["none", "top", "right", "bottom", "left", "topRight"].includes(String(atom))) return undefined;
      result[key] = atom;
    } else if (key === "dataLabelPosition") {
      if (!isChartDataLabelPosition(atom)) return undefined;
      result[key] = atom;
    } else if (key === "dataLabelSeparator") {
      if (!isChartDataLabelSeparator(atom)) return undefined;
      result[key] = atom;
    } else if (key === "dataLabelRegionLayout") {
      // A filled map's Map labels pick. It rides in this envelope rather than
      // in the block's boolean because the boolean cannot say which of Office's
      // two "on" values (`bestFitOnly` / `showAll`) the part carries.
      if (!isChartMapLabelLayout(atom)) return undefined;
      result[key] = atom;
    } else if (key === "axisOptions") {
      // Office's Format Axis > Axis Options, the shared record per axis.
      const normalized = normalizeOfficeChartAxesOptions(atom);
      if (!normalized) return undefined;
      result[key] = normalized;
    } else if (key === "hiddenSeries" || key === "hiddenCategories") {
      // Chart Filters: positions in the chart's series / categories, sorted,
      // kept in the part as Office keeps them (`c15:filtered…Series`, narrowed
      // references with `c15:fullRef`); the shared plan checks them against
      // the chart's data.
      if (
        !Array.isArray(atom) || atom.length === 0 ||
        atom.some((index, at) => !Number.isInteger(index) || index < 0 || (at > 0 && index <= atom[at - 1]))
      ) return undefined;
      result[key] = [...atom];
    } else if (key === "text") {
      // The chart's own text (`c:chartSpace/c:txPr`): Home ▸ Font on the
      // selected chart, inherited by every element that states none.
      const text = normalizeOfficeChartText(atom);
      if (!text || !record(atom) || Object.keys(atom).length !== Object.keys(text).length) return undefined;
      result[key] = text;
    } else if (key === "displayBlanksAs") {
      // Hidden and Empty Cells > Show empty cells as, as the part states it.
      if (!isChartDisplayBlanksAs(atom)) return undefined;
      result[key] = atom;
    } else if (key === "binIntervalClosed") {
      // A histogram's or pareto's closed bucket side, as the part declares it.
      // Carried so an imported chart is drawn with the edges Office draws and
      // written back with the side it arrived with.
      if (!isChartBinIntervalClosed(atom)) return undefined;
      result[key] = atom;
    } else if (key === "parentLabelLayout") {
      // A treemap's Label Options other than Overlapping (Banner, None), as
      // the part states them: drawn by the shared renderer, written back.
      if (!isChartParentLabelLayout(atom)) return undefined;
      result[key] = atom;
    } else if (key === "histogramBins") {
      // A histogram's or pareto's Format Axis > Bins, as the part states them.
      const bins = normalizeChartHistogramBins(atom);
      if (!bins || !record(atom) || Object.keys(atom).length !== Object.keys(bins).length) return undefined;
      result[key] = bins;
    } else if (key === "gapWidth") {
      // A histogram's, pareto's or waterfall's Gap Width, as the part states it.
      if (typeof atom !== "number" || !Number.isInteger(atom) || atom < 0 || atom > 500) return undefined;
      result[key] = atom;
    } else if (key === "boxWhisker") {
      // A box plot's Series Options, as the part states them.
      const options = normalizeChartBoxWhisker(atom);
      if (!options || !record(atom) || Object.keys(atom).length !== Object.keys(options).length) return undefined;
      result[key] = options;
    } else if (key === "axisTitleStyles") {
      // An axis title's own font and paint, as the imported part states it
      // (H08_32readwriteColumnChart4's 14pt titles). Left off this list, one
      // such chart made the whole workbook fail the page schema and open
      // blank and read-only.
      const normalized = normalizePresentationChartAxisTitleStyles(atom);
      if (!normalized || !Array.isArray(atom) || normalized.length !== atom.length) return undefined;
      result[key] = normalized;
    } else if (key === "seriesStyles" || key === "sliceStyles" || key === "axisStyles") {
      const normalized = key === "seriesStyles" ? normalizePresentationChartSeriesStyles(atom)
        : key === "sliceStyles" ? normalizePresentationChartSliceStyles(atom) : normalizePresentationChartAxisStyles(atom);
      if (!normalized || !Array.isArray(atom) || normalized.length !== atom.length ||
        atom.some((entry, index) => !record(entry) || Object.keys(entry).some(field => !Object.hasOwn(normalized[index], field)))) return undefined;
      result[key] = normalized;
    } else return undefined;
  }
  return result as SpreadsheetChartFormat;
}

export function spreadsheetChartAreaStyle(format: SpreadsheetChartFormat | undefined, target: SpreadsheetChartFormatTarget): PresentationChartAreaStyle | undefined {
  if (target.startsWith("series:")) return format?.seriesStyles?.find(item => item.seriesIndex === Number(target.slice(7)));
  if (target.startsWith("slice:")) return format?.sliceStyles?.find(item => item.pointIndex === Number(target.slice(6)));
  if (target.startsWith("axis:")) return format?.axisStyles?.find(item => item.axis === target.slice(5));
  return format?.[target as typeof SPREADSHEET_CHART_AREA_KEYS[number]];
}

/** Legacy menus and the shared pane expose one persisted display authority.
 *
 * The server projection states the same keys from the same format
 * (`editable_page._xlsx_chart_display_toggles`), so one workbook describes
 * itself the same way whichever chain opened it. Either axis showing is "axes
 * on", counting only the axes this chart has: a funnel carries its category
 * axis alone, and reading the axis it lacks as showing would leave its Axes
 * switch stuck on after Office wrote it off. */
export function spreadsheetChartLegacyDisplay(format:SpreadsheetChartFormat|undefined):Record<string,string|boolean> {
  if(!format)return {};
  const axes=(["axisX","axisY"] as const).filter(key=>format[key]!==undefined);
  return {...(format.legend!==undefined?{showLegend:format.legend!=="none",...(format.legend!=="none"?{legendPosition:format.legend}:{})}:{}),
    ...(format.dataLabels!==undefined?{showDataLabels:format.dataLabels}:{}),
    ...(format.dataTable!==undefined?{showDataTable:format.dataTable}:{}),
    ...(axes.length?{showAxisLabels:axes.some(key=>format[key]!==false)}:{}),
    ...(format.gridlinesX!==undefined||format.gridlinesY!==undefined?{showGridlines:format.gridlinesX===true||format.gridlinesY===true}:{})};
}

export function patchSpreadsheetChartArea(
  current: SpreadsheetChartFormat | undefined,
  target: SpreadsheetChartFormatTarget,
  patch: PresentationChartAreaStyle | null,
): SpreadsheetChartFormat | null {
  if (patch) patch = Object.fromEntries(Object.entries(patch).filter(([, atom]) => atom !== undefined));
  const next = { ...current };
  if (target.startsWith("series:")) {
    const seriesIndex = Number(target.slice(7));
    const items = (next.seriesStyles ?? []).filter(item => item.seriesIndex !== seriesIndex);
    if (patch) items.push({ seriesIndex, ...patch } as NonNullable<SpreadsheetChartFormat["seriesStyles"]>[number]);
    if (items.length) next.seriesStyles = items.sort((a, b) => a.seriesIndex - b.seriesIndex);
    else delete next.seriesStyles;
  } else if (target.startsWith("slice:")) {
    const pointIndex = Number(target.slice(6));
    const items = (next.sliceStyles ?? []).filter(item => item.pointIndex !== pointIndex);
    if (patch) items.push({ pointIndex, ...patch });
    if (items.length) next.sliceStyles = items.sort((a, b) => a.pointIndex - b.pointIndex);
    else delete next.sliceStyles;
  } else if (target.startsWith("axis:")) {
    const axis = target.slice(5) as "category" | "value";
    const items = (next.axisStyles ?? []).filter(item => item.axis !== axis);
    if (patch) items.push({ axis, ...presentationChartAxisStylePart(patch) });
    if (items.length) next.axisStyles = items;
    else delete next.axisStyles;
  } else {
    const key = target as typeof SPREADSHEET_CHART_AREA_KEYS[number];
    if (!SPREADSHEET_CHART_AREA_KEYS.includes(key)) return null;
    if (patch) next[key] = patch;
    else delete next[key];
    // Formatting all the major gridlines at once replaces each axis' own paint.
    if (key === "gridlineStyle") {
      delete next.gridlineStyleX;
      delete next.gridlineStyleY;
    }
  }
  return normalizeSpreadsheetChartFormat(next) ?? null;
}

/** A common renderer view only: no cell/formula/source payload is copied into
 * a second chart document. Explicit per-element overrides win over defaults. */
export function spreadsheetChartRenderFormat(
  defaults: ChartBlockFormat,
  value: SpreadsheetChartFormat | undefined,
): ChartBlockFormat {
  if (!value) return defaults;
  const normalized = normalizeSpreadsheetChartFormat(value);
  if (!normalized) return defaults;
  const series = new Map(defaults.seriesStyles?.map(item => [item.seriesIndex, item]));
  for (const item of normalized.seriesStyles ?? []) series.set(item.seriesIndex, { ...series.get(item.seriesIndex), ...item });
  const slices = new Map(defaults.sliceStyles?.map(item => [item.pointIndex, item]));
  for (const item of normalized.sliceStyles ?? []) slices.set(item.pointIndex, { ...slices.get(item.pointIndex), ...item });
  const sharedFormat={...normalized};
  delete sharedFormat.layout;
  if(sharedFormat.dataLabels===false)for(const key of Object.keys(sharedFormat))if(key.startsWith("dataLabel")&&key!=="dataLabels")delete (sharedFormat as Record<string,unknown>)[key];
  if(sharedFormat.dataTable===false){delete sharedFormat.dataTableStyle;delete sharedFormat.dataTableLegendKeys;}
  return { ...defaults, ...sharedFormat,
    ...(series.size ? { seriesStyles: [...series.values()] } : {}),
    ...(slices.size ? { sliceStyles: [...slices.values()] } : {}) };
}

/** Legacy commands and desktop controls edit one authority. Updating an old
 * toggle must not leave a newer override silently winning at render/export. */
export function synchronizeSpreadsheetChartFormat<T extends Record<string, unknown>>(previous:T,next:T):T {
  if(previous.nativeTitlePresent!==undefined&&previous.title!==next.title)next={...next,nativeTitlePresent:typeof next.title === "string" && Boolean(next.title.trim())};
  if (!previous.chartFormat || previous.chartFormat!==next.chartFormat) return next;
  const format=normalizeSpreadsheetChartFormat(previous.chartFormat);
  if (!format) return next;
  const result={...format};
  if (spreadsheetChartTitle(previous) !== spreadsheetChartTitle(next) && result.layout) {
    result.layout = { ...result.layout, title: Boolean(spreadsheetChartTitle(next).trim()) };
  }
  const changed=(key:string)=>previous[key]!==next[key];
  if(changed("showLegend")||changed("legendPosition"))result.legend=next.showLegend===false?"none":(next.legendPosition??"top") as NonNullable<SpreadsheetChartFormat["legend"]>;
  if(changed("showDataLabels")){
    result.dataLabels=next.showDataLabels===true;
    if(!result.dataLabels)for(const key of Object.keys(result))if(key.startsWith("dataLabel")&&key!=="dataLabels")delete (result as Record<string,unknown>)[key];
  }
  if(changed("showAxisLabels")){result.axisX=next.showAxisLabels!==false;result.axisY=next.showAxisLabels!==false;}
  if(changed("showGridlines")){result.gridlinesX=next.showGridlines!==false;result.gridlinesY=next.showGridlines!==false;}
  if(changed("showDataTable")){result.dataTable=next.showDataTable===true;if(!result.dataTable){delete result.dataTableStyle;delete result.dataTableLegendKeys;}}
  if(changed("chartColorPalette")||changed("seriesStyles")){
    const effects=result.seriesStyles?.filter(item=>item.effect).map(item=>({seriesIndex:item.seriesIndex,effect:item.effect}));
    if(effects?.length)result.seriesStyles=effects;else delete result.seriesStyles;
  }
  if(changed("chartColorPalette")||changed("sliceStyles")){
    const effects=result.sliceStyles?.filter(item=>item.effect).map(item=>({pointIndex:item.pointIndex,effect:item.effect}));
    if(effects?.length)result.sliceStyles=effects;else delete result.sliceStyles;
  }
  if(changed("title")&&!next.title)delete result.chartTitleStyle;
  if(changed("type")||changed("chartType")) {
    const {chartAreaStyle,plotAreaStyle,chartTitleStyle}=result;
    return {...next,chartSubtype:undefined,chartFormat:{...(chartAreaStyle?{chartAreaStyle}:{}),...(plotAreaStyle?{plotAreaStyle}:{}),...(chartTitleStyle?{chartTitleStyle}:{})},desktopStyleId:undefined};
  }
  return {...next,chartFormat:result};
}
