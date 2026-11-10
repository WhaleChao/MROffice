/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the chart engine draws a deck chart from, shared by the slide's chart
 * view (`PresentationChartCanvas`) and the server's ahead-of-print filling of
 * a deck's chart pictures (`deckChartPictureRequests`): one function, so the
 * two ask for the same picture by the same request.
 */
import type { PresentationChartBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { canonicalJson } from "@mrnote/core/web/lib/engine-host/client.ts";
import type { ChartRenderRequest } from "@mrnote/core/web/lib/engine-host/charts-types.ts";
import { readPresentationChartRenderAttrs } from "@mrnote/core/web/lib/presentation-native-chart-view.ts";
import { readPresentationChartV2 } from "@mrnote/core/web/lib/presentation-chart-v2.ts";
import { resolveChartDataTable } from "@mrnote/core/web/lib/presentation-chart-toggles.ts";
import { officeChartTextEChartsStyle } from "@mrnote/core/web/lib/office-charts/chart-text.ts";
import {
  officeChartDataTable,
  officeChartDataTableBand,
  officeChartDataTableValue,
  type OfficeChartDataTable,
} from "@mrnote/core/web/lib/office-charts/data-table.ts";
import { geometryNumber } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { AUTOFIT_EMU_PER_PX } from "./text-layout-css";

/** The picture's own size in slide px: the block's extent, as the chart is drawn. */
function pictureSize(block: PresentationChartBlock): { width: number; height: number } {
  return {
    width:
      typeof block.width === "number" && block.width > 0 ? block.width / 9525 : 320,
    height:
      typeof block.height === "number" && block.height > 0 ? block.height / 9525 : 180,
  };
}

function documentLocale(): string {
  return typeof document === "undefined" ? "en" : document.documentElement.lang || "en";
}

/**
 * What the chart engine draws a deck chart from: the block (where the frame
 * sits never changes the picture, so `left`/`top` stay out), the deck
 * theme's accents and the reader's language. The content key also leaves the
 * size out, so a picture of this chart at another size stands in
 * (unstretched) while the new size is drawn.
 */
export function presentationChartPictureRequest(
  block: PresentationChartBlock,
  themeAccents: readonly string[] | undefined,
  locale = documentLocale(),
  themeFontFamily?: string,
): { request: ChartRenderRequest; contentKey: string } {
  const { width, height } = pictureSize(block);
  const drawn = { ...block, left: undefined, top: undefined } as unknown as Record<string, unknown>;
  const request: ChartRenderRequest = {
    host: "presentation",
    block: drawn,
    ...(themeAccents?.length ? { themeAccents: [...themeAccents] } : {}),
    ...(themeFontFamily ? { themeFontFamily } : {}),
    locale,
    width,
    height,
  };
  const contentKey = canonicalJson({
    ...request,
    block: { ...drawn, width: undefined, height: undefined },
    width: undefined,
    height: undefined,
  });
  return { request, contentKey };
}

/** The raster picture an imported chart's frame carries for when the page
 * cannot plot it (`nativeChartFallbackSrc`), or `null`. */
export function presentationChartFallbackPicture(block: PresentationChartBlock): string | null {
  const src = block.nativeChartFallbackSrc;
  return typeof src === "string" && /^data:image\/(?:png|jpeg|gif);base64,/i.test(src) ? src : null;
}

/** A scratch chart's grid (`data.rows`: category, then one value per series). */
function legacyDataTable(block: PresentationChartBlock): OfficeChartDataTable | null {
  const rows = (block.data?.rows ?? []) as unknown[][];
  const headers = block.data?.headers ?? [];
  const count = Math.max(0, ...rows.map((row) => row.length - 1));
  if (!rows.length || count < 1) return null;
  return {
    categories: rows.map((row) => String(row[0] ?? "")),
    series: Array.from({ length: count }, (_, index) => ({
      name: headers[index + 1]?.trim() || `Series ${index + 1}`,
      values: rows.map((row) => {
        const raw = row[index + 1];
        const value = typeof raw === "number" ? raw : Number(String(raw ?? "").trim() || Number.NaN);
        return Number.isFinite(value) ? officeChartDataTableValue(value) : String(raw ?? "");
      }),
    })),
    legendKeys: true,
  };
}

/** The text style a chart's whole-chart text gives its data table. */
export type PresentationChartTableText = {
  fontSize?: number;
  color?: string;
  fontFamily?: string;
  fontWeight?: unknown;
  fontStyle?: unknown;
};

/**
 * The data table a deck chart shows, from the same reading the chart is drawn
 * from: a Chart V2 or imported chart's `format.dataTable` (what Add Chart
 * Element > Data Table writes), a scratch chart's `showDataTable`. PPT-B140:
 * only the scratch flag was read, so a new chart's table never appeared.
 */
export function presentationChartDataTableBasis(block: PresentationChartBlock): {
  table: OfficeChartDataTable;
  style: unknown;
  inherited: PresentationChartTableText;
} | null {
  let table: OfficeChartDataTable | null;
  let style: unknown;
  let attrs: ReturnType<typeof readPresentationChartRenderAttrs> = null;
  try {
    attrs = readPresentationChartRenderAttrs(block);
  } catch {
    attrs = null;
  }
  // The chart's own text (Home font size and colour on the whole chart)
  // reaches the table where the table states nothing of its own (PPT-B293).
  let inherited: PresentationChartTableText = {};
  if (attrs) {
    table = officeChartDataTable(attrs);
    style = attrs.format?.dataTableStyle;
    const text = officeChartTextEChartsStyle(attrs.format?.text);
    inherited = {
      ...(typeof text.fontSize === "number" ? { fontSize: text.fontSize } : {}),
      ...(typeof text.color === "string" ? { color: text.color } : {}),
      ...(typeof text.fontFamily === "string" ? { fontFamily: text.fontFamily } : {}),
      ...(text.fontWeight !== undefined ? { fontWeight: text.fontWeight } : {}),
      ...(text.fontStyle !== undefined ? { fontStyle: text.fontStyle } : {}),
    };
  } else {
    table = resolveChartDataTable(block) ? legacyDataTable(block) : null;
    style = block.dataTableStyle;
  }
  return table ? { table, style, inherited } : null;
}

/** The band a deck chart's data table takes under its plot, design px. */
export function presentationChartDataTableBand(
  basis: NonNullable<ReturnType<typeof presentationChartDataTableBasis>>,
  frameHeightPx: number,
): ReturnType<typeof officeChartDataTableBand> {
  return officeChartDataTableBand(
    basis.table.series.length + 1,
    frameHeightPx,
    typeof basis.inherited.fontSize === "number" ? basis.inherited.fontSize : undefined,
  );
}

/**
 * The block a deck chart's picture is drawn for: the frame less the band its
 * data table takes from the bottom (the chart is drawn in the rest, not
 * squeezed into it), at the block's own extent or a resize's live one.
 */
export function presentationChartCanvasBlock(
  block: PresentationChartBlock,
  liveSize?: { width: number; height: number },
): PresentationChartBlock {
  const resizing =
    liveSize !== undefined && (liveSize.width !== block.width || liveSize.height !== block.height);
  const frameHeight = resizing ? liveSize!.height : geometryNumber(block, "height");
  const basis = presentationChartDataTableBasis(block);
  const tableEmu = basis
    ? presentationChartDataTableBand(basis, typeof frameHeight === "number" ? frameHeight / AUTOFIT_EMU_PER_PX : 0)
        .heightPx * AUTOFIT_EMU_PER_PX
    : 0;
  const sized = resizing ? { ...block, width: liveSize!.width, height: liveSize!.height } : block;
  return tableEmu > 0 && typeof frameHeight === "number"
    ? { ...sized, height: Math.max(1, frameHeight - tableEmu) }
    : sized;
}

/** Whether a deck chart is drawn by the engine (else its saved picture or nothing). */
export function presentationChartIsDrawn(block: PresentationChartBlock): boolean {
  return !(presentationChartFallbackPicture(block) && !readPresentationChartV2(block.chartV2));
}

/**
 * A deck chart's block as its picture is asked for: with the slide's own
 * overrides and the native appearance recovered for it, its display always
 * stated in `source` (`usePresentationChartPicture`).
 */
export function presentationChartDisplayedBlock(
  sourceBlock: PresentationChartBlock,
  {
    recoveredDisplay,
    inheritedDisplay,
    showLegend,
    showAxisLabels,
    chartStyle,
  }: {
    recoveredDisplay?: unknown;
    inheritedDisplay?: unknown;
    showLegend?: boolean;
    showAxisLabels?: boolean;
    chartStyle?: number;
  } = {},
): PresentationChartBlock {
  return {
    ...sourceBlock,
    ...(showLegend === undefined ? {} : { showLegend }),
    ...(showAxisLabels === undefined ? {} : { showAxisLabels }),
    ...(chartStyle === undefined ? {} : { chartStyle }),
    source: {
      ...sourceBlock.source,
      chartDisplay: sourceBlock.source?.chartDisplay ?? recoveredDisplay ?? inheritedDisplay,
    },
  } as PresentationChartBlock;
}
