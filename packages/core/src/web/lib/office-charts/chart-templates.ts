/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's chart templates (Save as Template), the one model all three
 * editors apply. A workbook keeps its own (XLS-A081, `viewOptions.
 * chartTemplates`); a Docs or Slides chart saves to the user's Templates
 * folder (`/api/v1/chart-templates`, `chart-template-store`).
 *
 * Microsoft Support, "Save a custom chart as a template": right-click a chart
 * > Save as Template saves a `.crtx` that "contains chart formatting and
 * stores the colors that are in use"; Insert Chart / Change Chart Type > All
 * Charts > Templates applies one "like any other chart type", and a chart
 * made from it uses the template's colours, not the workbook theme's.
 *
 * A template is the look of one chart of the shared model -- its kind and
 * subtype, style preset and format -- without its data, titles or the
 * data-bound picks (hidden series / categories). The `.crtx` file itself is
 * built by the API's chart compiler (`xlsx/chart_template`).
 */
import {
  isChartFamily,
  isChartStylePreset,
  isChartSubtypeForFamily,
  type ChartBlockAttrsV2,
  type ChartBlockFormat,
  type ChartFamily,
  type ChartStylePresetId,
  type ChartSubtype,
} from "../../components/console/editor/extensions/chartBlockShared";
import { officeChartWithType } from "./chart-type";

export interface OfficeChartTemplate {
  id: string;
  name: string;
  chartType: ChartFamily;
  chartSubtype: ChartSubtype;
  stylePreset: ChartStylePresetId;
  format?: ChartBlockFormat;
}

export const MAX_OFFICE_CHART_TEMPLATES = 64;
export const MAX_CHART_TEMPLATE_NAME_LENGTH = 255;

/** The format keys that describe a chart's data, not its look. */
const DATA_BOUND_FORMAT_KEYS = ["hiddenSeries", "hiddenCategories"] as const;

/** One chart's look as a template named `name`. */
export function chartTemplateFromModel(
  attrs: ChartBlockAttrsV2,
  name: string,
  id: string,
): OfficeChartTemplate {
  const format = attrs.format ? { ...attrs.format } : undefined;
  if (format) for (const key of DATA_BOUND_FORMAT_KEYS) delete (format as Record<string, unknown>)[key];
  return {
    id,
    name: name.trim().slice(0, MAX_CHART_TEMPLATE_NAME_LENGTH),
    chartType: attrs.chartType,
    chartSubtype: attrs.chartSubtype,
    stylePreset: attrs.stylePreset,
    ...(format && Object.keys(format).length > 0 ? { format } : {}),
  };
}

/**
 * The chart with the template applied: the template's kind, style and
 * formatting over the chart's own data and titles; null when the chart's data
 * cannot take that kind (a scatter template on category data).
 */
export function applyChartTemplate(
  attrs: ChartBlockAttrsV2,
  template: OfficeChartTemplate,
): ChartBlockAttrsV2 | null {
  const typed = officeChartWithType(attrs, template.chartType, template.chartSubtype);
  if (!typed) return null;
  const keep: Record<string, unknown> = {};
  for (const key of DATA_BOUND_FORMAT_KEYS) {
    const value = (attrs.format as Record<string, unknown> | undefined)?.[key];
    if (value !== undefined) keep[key] = value;
  }
  const format = { ...(template.format ?? {}), ...keep } as ChartBlockFormat;
  const next: ChartBlockAttrsV2 = { ...typed, stylePreset: template.stylePreset };
  if (Object.keys(format).length > 0) next.format = format;
  else delete next.format;
  return next;
}

/** Templates as persisted, bounded; junk is dropped. */
export function normalizeChartTemplates(raw: unknown): OfficeChartTemplate[] {
  if (!Array.isArray(raw)) return [];
  const out: OfficeChartTemplate[] = [];
  const ids = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_OFFICE_CHART_TEMPLATES) break;
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;
    const record = entry as Record<string, unknown>;
    const id = typeof record.id === "string" ? record.id.trim().slice(0, 200) : "";
    const name = typeof record.name === "string" ? record.name.trim().slice(0, MAX_CHART_TEMPLATE_NAME_LENGTH) : "";
    if (!id || !name || ids.has(id)) continue;
    if (!isChartFamily(record.chartType) || !isChartSubtypeForFamily(record.chartType, record.chartSubtype)) continue;
    if (!isChartStylePreset(record.stylePreset)) continue;
    const format =
      record.format && typeof record.format === "object" && !Array.isArray(record.format)
        ? (record.format as ChartBlockFormat)
        : undefined;
    ids.add(id);
    out.push({
      id,
      name,
      chartType: record.chartType,
      chartSubtype: record.chartSubtype as ChartSubtype,
      stylePreset: record.stylePreset,
      ...(format ? { format } : {}),
    });
  }
  return out;
}

/**
 * A template saved to the user's Templates folder: the chart it was saved
 * from (validated by the API as the shared chart model) as a template; null
 * for a chart that is not one of the model.
 */
export function officeChartTemplateFromSavedChart(saved: {
  id: string;
  name: string;
  attrs: unknown;
}): OfficeChartTemplate | null {
  const attrs = saved.attrs as ChartBlockAttrsV2 | null;
  if (
    !attrs ||
    typeof attrs !== "object" ||
    !isChartFamily(attrs.chartType) ||
    !isChartSubtypeForFamily(attrs.chartType, attrs.chartSubtype) ||
    !isChartStylePreset(attrs.stylePreset)
  ) {
    return null;
  }
  return chartTemplateFromModel(attrs, saved.name, saved.id);
}
