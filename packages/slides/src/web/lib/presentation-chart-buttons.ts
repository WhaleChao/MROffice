/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's chart buttons as pure transforms: Insert > Chart and the
 * one gate every Chart Design edit passes (`commitChartDesktop`). The ribbon
 * and the headless editor commands build the chart attributes with the shared
 * Office chart helpers and land them through these.
 */

import type { PresentationChartBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { ChartBlockAttrsV2 } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { type DocumentChartWorkbook } from "@mrnote/core/web/components/console/editor/extensions/documentChartWorkbook.ts";
import {
  seedPresentationChartWorkbook,
  withPresentationChartDesktop,
} from "@mrnote/core/web/lib/presentation-chart-desktop.ts";
import {
  presentationChartDataEditAllowed,
  presentationChartV2CanEdit,
  presentationChartV2DataWritable,
  presentationChartV2Kind,
  presentationChartTypeTargetLock,
  presentationChartV2TitleWritable,
  presentationChartV2TypeWritable,
} from "@mrnote/core/web/lib/presentation-chart-v2.ts";
import { canEditBlockContent } from "@mrnote/core/web/lib/presentation-text-case.ts";

/**
 * Insert > Chart: a chart block with its seeded data sheet, sized for its
 * title, axes and legend (58% x 55% of the slide) and centred. Geometry is
 * canonical slide-relative, independent of window or viewport zoom.
 */
export function presentationChartBlockForInsert(
  id: string,
  chartV2: ChartBlockAttrsV2,
  bounds: { width: number; height: number },
): PresentationChartBlock {
  const width = Math.max(1, Math.round(bounds.width * 0.58));
  const height = Math.max(1, Math.round(bounds.height * 0.55));
  const chartWorkbook = seedPresentationChartWorkbook(chartV2);
  return {
    id,
    type: "chart" as const,
    // The block's own title is the chart's: the server refuses a chart whose
    // outer title differs from its V2 title, so the sample title the leaf
    // carries must be stated here too (PPT-B399).
    title: chartV2.title ?? "",
    caption: "",
    chartType: presentationChartV2Kind(chartV2),
    chartV2,
    ...(chartWorkbook ? { chartWorkbook } : {}),
    width,
    height,
    left: Math.round((bounds.width - width) / 2),
    top: Math.round((bounds.height - height) / 2),
  } as PresentationChartBlock;
}

/**
 * A Chart Design edit landed on a chart block: the chart's new attributes
 * (and data sheet), or null when the chart cannot take them -- a
 * preserve-only chart (Lock fixes only its frame), a cell-bound title of a linked chart, data or a type
 * the chart's file part owns, a new category or series on cached values.
 */
export function chartBlockWithDesktopAttrs(
  expected: PresentationChartBlock,
  attrs: NonNullable<PresentationChartBlock["chartV2"]>,
  book?: DocumentChartWorkbook,
): PresentationChartBlock | null {
  if (!canEditBlockContent(expected) || !presentationChartV2CanEdit(expected)) return null;
  // A cell-bound title of a linked chart: the save would decline it (`title`).
  if (
    !presentationChartV2TitleWritable(expected) &&
    (attrs.title ?? "") !== (expected.chartV2?.title ?? "")
  )
    return null;
  if (
    !presentationChartV2DataWritable(expected) &&
    (book !== undefined || JSON.stringify(attrs.data) !== JSON.stringify(expected.chartV2?.data))
  )
    return null;
  // A chart holding its data as cached values takes value edits only: the
  // save declines a new category or series of it ('data').
  if (!presentationChartDataEditAllowed(expected, attrs.data)) return null;
  if (
    !presentationChartV2TypeWritable(expected) &&
    (book !== undefined ||
      attrs.chartType !== expected.chartV2?.chartType ||
      attrs.chartSubtype !== expected.chartV2?.chartSubtype)
  )
    return null;
  // An imported classic part stays classic, and one read around a look the
  // model has no words for keeps its family (`_pptx_chart_patch`).
  if (
    attrs.chartType !== expected.chartV2?.chartType &&
    presentationChartTypeTargetLock(expected, attrs.chartType) !== null
  )
    return null;
  return withPresentationChartDesktop(expected, attrs, book);
}

/**
 * Find and Replace may rewrite a Chart V2 chart's drawn title only where the
 * chart's own title editing may (PPT-B031).
 */
export function presentationFindChartTitleWritable(block: Record<string, unknown>): boolean {
  const chart = block as unknown as PresentationChartBlock;
  // A linked chart's cell-bound title is the cell's text (`titleBound`).
  return presentationChartV2CanEdit(chart) && presentationChartV2TitleWritable(chart);
}
