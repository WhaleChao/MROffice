/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationChartBlock } from "../components/console/editor/native-document/documentCommands";
import type { DocumentChartAreaStyleKey } from "../components/console/editor/native-document/DocumentChartFormatRibbon";
import {
  normalizePresentationChartAreaStyle,
  type PresentationChartAreaStyle,
  type PresentationChartAreaPaint,
  type PresentationChartEffectPreset,
  chartPointEffectiveStyle,
  chartRowsWithPointStyle,
  chartStyleRowExtras,
} from "./presentation-chart-area-style";
import {
  readPresentationChartV2,
  withPresentationChartV2,
} from "./presentation-chart-v2";
import {
  chartWorkbookData,
  chartWorkbookIsStale,
  normalizeDocumentChartWorkbook,
  seedDocumentChartWorkbook,
  type DocumentChartWorkbook,
} from "../components/console/editor/extensions/documentChartWorkbook";
import {
  chartDataEditOrigins,
  chartFormatFollowingDataEdit,
  type ChartBlockAttrsV2,
  type ChartData,
  type ChartFamily,
  type ChartSubtype,
} from "../components/console/editor/extensions/chartBlockShared";
import {
  OFFICE_CHART_QUICK_LAYOUT_IDS,
  officeChartWithQuickLayout,
  type OfficeChartQuickLayoutId,
  type OfficeChartQuickLayoutWords,
} from "./office-charts/quick-layout";
import { officeChartWithType } from "./office-charts/chart-type";
import { officeChartWithRowsAndColumnsSwitched } from "./office-charts/chart-filters";
/** Quick Layout: the one catalogue Docs and the worksheet apply too. */
export const PPT_DESKTOP_CHART_LAYOUTS = OFFICE_CHART_QUICK_LAYOUT_IDS;
export type PresentationDesktopChartLayout = OfficeChartQuickLayoutId;

export function presentationChartSignature(value: unknown): string {
  return JSON.stringify(value, (_key, item) =>
    item && typeof item === "object" && !Array.isArray(item)
      ? Object.fromEntries(
          Object.entries(item).sort(([a], [b]) => a.localeCompare(b)),
        )
      : item,
  );
}

function labelsAsStored(value: unknown): unknown {
  if (typeof value === "string") return value.trim();
  if (Array.isArray(value)) return value.map(labelsAsStored);
  if (value && typeof value === "object")
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, labelsAsStored(item)]));
  return value;
}

export function readPresentationChartWorkbook(
  value: unknown,
  attrs?: ChartBlockAttrsV2,
): DocumentChartWorkbook | null {
  try {
    const book = normalizeDocumentChartWorkbook(value);
    // The save's own check (`normalize_ppt_chart_workbook`): the chart's data
    // is what the workbook's cells and stamped results state -- read, never
    // recalculated -- unless the results are still being calculated
    // (`calculationStale`: the chart's data lags behind until they land).
    // A label is compared as the save stores it, its edge spaces off (a cell
    
    if (
      attrs &&
      !chartWorkbookIsStale(book) &&
      presentationChartSignature(
        labelsAsStored(chartWorkbookData(attrs, book, { missing: "blank" }).data),
      ) !== presentationChartSignature(labelsAsStored(attrs.data))
    )
      return null;
    return book;
  } catch {
    return null;
  }
}

/**
 * A deck chart's data sheet seeded from its data, laid out as PowerPoint's
 * own (Insert > Chart opens a sheet whose corner cell A1 is blank, the
 * categories down column A and the series names across row 1, PPT-B020).
 */
export function seedPresentationChartWorkbook(
  attrs: ChartBlockAttrsV2,
): DocumentChartWorkbook | null {
  return seedDocumentChartWorkbook(attrs, { categoryHeader: "" });
}

export function presentationChartWorkbook(
  block: PresentationChartBlock,
): DocumentChartWorkbook | null {
  const attrs = readPresentationChartV2(block.chartV2);
  if (!attrs) return null;
  try {
    return block.chartWorkbook === undefined
      ? seedPresentationChartWorkbook(attrs)
      : readPresentationChartWorkbook(block.chartWorkbook, attrs);
  } catch {
    return null;
  }
}

/** Preserve the PPT frame/source envelope. Data+formula caches are one value. */








export function presentationChartAttrsWithData(
  attrs: ChartBlockAttrsV2,
  data: ChartData,
): ChartBlockAttrsV2 {
  const origins = attrs.format ? chartDataEditOrigins(attrs.data, data) : null;
  if (!origins) return { ...attrs, data };
  return { ...attrs, data, format: chartFormatFollowingDataEdit(attrs.format, origins.series, origins.categories) };
}

export function withPresentationChartDesktop(
  block: PresentationChartBlock,
  attrs: ChartBlockAttrsV2,
  workbook?: DocumentChartWorkbook,
): PresentationChartBlock | null {
  if (!readPresentationChartV2(attrs)) return null;
  const next = withPresentationChartV2(block, attrs);
  if (workbook !== undefined) {
    const valid = readPresentationChartWorkbook(workbook, attrs);
    return valid ? { ...next, chartWorkbook: valid } : null;
  }
  if (
    block.chartWorkbook !== undefined &&
    !readPresentationChartWorkbook(block.chartWorkbook, attrs)
  )
    return null;
  return next;
}

export function presentationChartAreaStyle(
  attrs: ChartBlockAttrsV2,
  area: DocumentChartAreaStyleKey,
): PresentationChartAreaStyle | undefined {
  const format = attrs.format;
  if (area.startsWith("errorBarStyle:"))
    return format?.seriesErrorBars?.find((entry) => entry.seriesIndex === Number(area.slice(14)))?.style;
  if (area.startsWith("series:"))
    return format?.seriesStyles?.find(
      (s) => s.seriesIndex === Number(area.slice(7)),
    );
  if (area.startsWith("slice:"))
    return format?.sliceStyles?.find(
      (s) => s.pointIndex === Number(area.slice(6)),
    );
  // One data point of a series (Format Data Point: a column's or a bar's
  // own fill, line and effects), selected by a second click on it.
  const point = /^point:(\d+):(\d+)$/.exec(area);
  if (point) return chartPointEffectiveStyle(format, Number(point[1]), Number(point[2]));
  if (area.startsWith("axis:"))
    return format?.axisStyles?.find((s) => s.axis === area.slice(5));
  if (area.startsWith("axisTitle:"))
    return format?.axisTitleStyles?.find((s) => s.axis === area.slice(10));
  return normalizePresentationChartAreaStyle(
    (format as Record<string, unknown> | undefined)?.[
      area === "gridlines" ? "gridlineStyle" : area
    ],
  );
}

export function setPresentationChartAreaStyle(
  attrs: ChartBlockAttrsV2,
  area: DocumentChartAreaStyleKey,
  style: PresentationChartAreaStyle | null,
): ChartBlockAttrsV2 | null {
  const normalized =
    style === null ? null : normalizePresentationChartAreaStyle(style) ?? null;
  const format = { ...attrs.format };
  if (area.startsWith("errorBarStyle:")) {
    // Another series' error bars' own line (`seriesErrorBars[].style`): no
    // fill, and only on bars the series has.
    const seriesIndex = Number(area.slice(14));
    const entries = format.seriesErrorBars ?? [];
    if (!entries.some((entry) => entry.seriesIndex === seriesIndex)) return null;
    if (normalized?.fill !== undefined) return null;
    if (style !== null && !normalized) return null;
    format.seriesErrorBars = entries.map((entry) => {
      if (entry.seriesIndex !== seriesIndex) return entry;
      const { style: _style, ...rest } = entry;
      return normalized ? { ...rest, style: normalized } : rest;
    });
    return readPresentationChartV2({ ...attrs, format });
  }
  const point = /^point:(\d+):(\d+)$/.exec(area);
  if (point) {
    // One data point of a series: its own row in sliceStyles, by series and
    // point (`c:dPt` in that series), its paint over the series'.
    if (style !== null && !normalized) return null;
    const rows = chartRowsWithPointStyle(format.sliceStyles, Number(point[1]), Number(point[2]), style);
    if (rows) format.sliceStyles = rows;
    else delete format.sliceStyles;
    return readPresentationChartV2({ ...attrs, format });
  }
  const rule = area.startsWith("series:")
    ? (["seriesStyles", "seriesIndex", Number(area.slice(7))] as const)
    : area.startsWith("slice:")
      ? (["sliceStyles", "pointIndex", Number(area.slice(6))] as const)
      : area.startsWith("axis:")
        ? (["axisStyles", "axis", area.slice(5)] as const)
        : area.startsWith("axisTitle:")
          ? (["axisTitleStyles", "axis", area.slice(10)] as const)
          : null;
  if (rule) {
    const [key, indexKey, index] = rule;
    const rows = (format[key] ?? []) as readonly Record<string, unknown>[];
    const existing = rows.find((item) => item[indexKey] === index);
    // A series' marker or Smoothed line (a point's explosion) is a row of
    // its own even with no paint beside it.
    const extras = style === null ? {} : chartStyleRowExtras(key, existing, style);
    if (style !== null && !normalized && !Object.keys(extras).length) return null;
    const next = [
      ...rows.filter((item) => item[indexKey] !== index),
      ...(normalized || Object.keys(extras).length ? [{ [indexKey]: index, ...(normalized ?? {}), ...extras }] : []),
    ];
    if (next.length) Object.assign(format, { [key]: next });
    else delete format[key];
  } else {
    if (style !== null && !normalized) return null;
    const key = area === "gridlines" ? "gridlineStyle" : area;
    if (normalized) Object.assign(format, { [key]: normalized });
    else delete (format as Record<string, unknown>)[key];
    if (key === "gridlineStyle") {
      // All the major gridlines at once: each axis' own paint goes.
      delete format.gridlineStyleX;
      delete format.gridlineStyleY;
    }
  }
  return readPresentationChartV2({ ...attrs, format });
}

/** Same paint transitions as Docs, with the PPT host retaining commit ownership. */
export function presentationChartPaintStyle(
  current: PresentationChartAreaStyle | undefined,
  field: "fill" | "outline",
  paint: PresentationChartAreaPaint,
): PresentationChartAreaStyle {
  const style = { ...current, [field]: paint };
  if (style.shapePaint) {
    style.shapePaint = { ...style.shapePaint };
    if (field === "fill") {
      delete style.shapePaint.fillGradient;
      delete style.shapePaint.fillPattern;
    } else {
      // A solid outline replaces a gradient line.
      delete style.shapePaint.lineGradient;
      style.shapePaint.outlineWidthPt =
        paint === "none" ? 0 : style.shapePaint.outlineWidthPt || 1;
    }
    if (!Object.keys(style.shapePaint).length) delete style.shapePaint;
  }
  return style;
}

export function presentationChartEffectStyle(
  current: PresentationChartAreaStyle | undefined,
  effect: PresentationChartEffectPreset | null,
): PresentationChartAreaStyle | null {
  const style = { ...current };
  if (style.shapePaint?.shadow) {
    style.shapePaint = { ...style.shapePaint };
    delete style.shapePaint.shadow;
  }
  if (effect) style.effect = effect;
  else delete style.effect;
  return Object.keys(style).length ? style : null;
}

/** WordArt presets replace text paint, while retaining the title container. */
export function presentationChartTitleTextStyle(
  current: PresentationChartAreaStyle | undefined,
  patch: PresentationChartAreaStyle,
): PresentationChartAreaStyle | null {
  const style = { ...current };
  delete style.textFill;
  delete style.textOutline;
  delete style.textEffect;
  delete style.textPaint;
  const next = { ...style, ...patch };
  return Object.values(next).some((value) => value !== undefined) ? next : null;
}

/**
 * The owned V2 chart with a Quick Layout applied: the shared recipe
 * (`lib/office-charts/quick-layout`), then the deck's own V2 gate.
 */
export function presentationChartQuickLayout(
  attrs: ChartBlockAttrsV2,
  id: PresentationDesktopChartLayout,
  words?: OfficeChartQuickLayoutWords,
): ChartBlockAttrsV2 | null {
  const next = officeChartWithQuickLayout(attrs, id, words);
  return next ? readPresentationChartV2(next) : null;
}

/**
 * The owned V2 chart as another type: the shared rule
 * (`lib/office-charts/chart-type`), then the deck's own V2 gate.
 */
export function changePresentationChartType(
  attrs: ChartBlockAttrsV2,
  family: ChartFamily,
  subtype: ChartSubtype,
): ChartBlockAttrsV2 | null {
  const next = officeChartWithType(attrs, family, subtype);
  return next ? readPresentationChartV2(next) : null;
}

/**
 * Chart Design > Switch Row/Column on category data: with a data sheet, the
 * sheet's binding plots by the other axis and the chart takes the data it
 * then reads (the filters moved with it); without one, the chart's own data
 * transposed. Null when the chart has no category data or the sheet cannot
 * be read the other way round.
 */
export function presentationChartRowsColumnsSwitched(
  attrs: ChartBlockAttrsV2,
  workbook: DocumentChartWorkbook | null | undefined,
): { attrs: ChartBlockAttrsV2; workbook?: DocumentChartWorkbook } | null {
  if (attrs.data.kind !== "category") return null;
  if (!workbook) {
    const switched = officeChartWithRowsAndColumnsSwitched(attrs);
    return switched ? { attrs: switched } : null;
  }
  try {
    // The sheet read the other way round from its stamped results (a formula
    // without one refuses the switch until the chart has them).
    const next = chartWorkbookData(attrs, {
      ...workbook,
      binding: {
        ...workbook.binding,
        plotBy: workbook.binding.plotBy === "rows" ? "columns" : "rows",
      },
    });
    const switched = officeChartWithRowsAndColumnsSwitched(attrs, next.data);
    return switched ? { attrs: switched, workbook: next.workbook } : null;
  } catch {
    return null;
  }
}
