/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart Format > Shape Styles on one chart element, as the Office-profile
 * gallery applies it in the document, the deck and the worksheet: which
 * elements take a fill (lines take outline styles only), which take no shared
 * style at all, and the element style a gallery tile writes. The ribbon
 * (DocumentChartFormatRibbon) and the AI's chart format_element commands run
 * these.
 */

import { officeChartShapeRecipePaint } from "./office-chart-shape-paint";
import type { OfficeShapeStyleRecipe } from "./office-shape-style-catalog";
import type { PresentationChartAreaPaint, PresentationChartAreaStyle } from "./presentation-chart-area-style";

type ChartKind = { readonly chartType?: string; readonly chartSubtype?: string } | null | undefined;

const LINE_AREAS = new Set([
  "gridlines",
  "gridlineStyleX",
  "gridlineStyleY",
  "minorGridlineStyleX",
  "minorGridlineStyleY",
  "trendlineStyle",
  "errorBarStyle",
  "dropLineStyle",
  "highLowLineStyle",
  "seriesLineStyle",
  "dataLabelLeaderLineStyle",
]);

/** Whether the element (a Format tab area key) takes a fill, or outline styles only. */
export function officeChartAreaFillable(attrs: ChartKind, area: string): boolean {
  if (/^series:[0-9]+$/.test(area)) {
    return !(
      attrs?.chartType === "line" ||
      (attrs?.chartType === "radar" && attrs.chartSubtype !== "filled") ||
      (attrs?.chartType === "xy" && !["bubble", "bubble3d", "markers"].includes(attrs.chartSubtype ?? ""))
    );
  }
  if (/^slice:[0-9]+$/.test(area)) return true;
  if (/^axis:/.test(area)) return false;
  return !LINE_AREAS.has(area);
}

/** Elements the shared Shape Styles gallery cannot style in this chart. */
export function officeChartAreaSharedStylesUnavailable(attrs: ChartKind, area: string): boolean {
  return (
    (area === "plotAreaStyle" && !["column", "bar", "line", "area", "xy", "combo", "stock"].includes(attrs?.chartType ?? "")) ||
    (/^series:[0-9]+$/.test(area) && ["radar", "surface", "stock"].includes(attrs?.chartType ?? ""))
  );
}

/**
 * The element's style after a Shape Styles tile: the tile's fill (when the
 * element takes one), outline and shape paint; any earlier effect goes. Null
 * when the tile does not apply (a filled tile on a line, a shared style the
 * element cannot take).
 */
export function officeChartShapeStyleApplied(
  attrs: ChartKind,
  area: string,
  current: PresentationChartAreaStyle | undefined,
  recipe: OfficeShapeStyleRecipe,
): PresentationChartAreaStyle | null {
  const fillable = officeChartAreaFillable(attrs, area);
  if (officeChartAreaSharedStylesUnavailable(attrs, area) || (!fillable && recipe.row !== "outline")) return null;
  return {
    ...(current ?? {}),
    ...(fillable ? { fill: recipe.fill as PresentationChartAreaPaint } : {}),
    outline: recipe.stroke as PresentationChartAreaPaint,
    shapePaint: officeChartShapeRecipePaint(recipe, fillable),
    effect: undefined,
  };
}
