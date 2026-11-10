/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Docs HTML export's chart pictures, from the server's chart engine.
 *
 * The export is built synchronously from one JSON snapshot; its charts are
 * drawn first (one batched `charts.render` call), and the projection then
 * looks each chart's picture up by its request. The request is the API
 * export's (`chart_render.document_chart_render_request`): the chart node's
 * attrs (its imported look and its sheet's number format ride in them), the
 * document theme a Change Colors pick resolves against, and the chart's Word
 * extent.
 */
import {
  CHART_HOST_EMU_PER_PX,
  documentChartSize,
} from "@mrnote/core/web/components/console/editor/extensions/documentChartHost.ts";
import type { ChartRenderPicture, ChartRenderRequest } from "@mrnote/core/web/lib/engine-host/charts-types.ts";
import { chartPictureKey, requestChartPicture } from "@mrnote/core/web/lib/office-charts/chart-picture-client.ts";
import type { TipTapNode } from "@/lib/document-export/tiptap-json";

/** The document attrs a chart's plan reads (its theme), nothing else. */
const DOCUMENT_THEME_KEYS = ["documentThemeColors", "designColorSet"] as const;

export type DocumentExportChartPictures = ReadonlyMap<string, ChartRenderPicture>;

export function documentExportChartRequest(
  nodeAttrs: Record<string, unknown>,
  documentAttrs: Record<string, unknown> | undefined,
): ChartRenderRequest {
  let size: ReturnType<typeof documentChartSize>;
  try {
    size = documentChartSize(nodeAttrs);
  } catch {
    size = undefined;
  }
  const theme = Object.fromEntries(
    DOCUMENT_THEME_KEYS.filter((key) => documentAttrs && key in documentAttrs).map((key) => [key, documentAttrs![key]]),
  );
  return {
    attrs: nodeAttrs,
    ...(Object.keys(theme).length ? { documentAttrs: theme } : {}),
    includeTitle: true,
    width: size ? size.widthEmu / CHART_HOST_EMU_PER_PX : 576,
    height: size ? size.heightEmu / CHART_HOST_EMU_PER_PX : 328.32,
  };
}

export function documentExportChartKey(
  nodeAttrs: Record<string, unknown>,
  documentAttrs: Record<string, unknown> | undefined,
): string {
  return chartPictureKey(documentExportChartRequest(nodeAttrs, documentAttrs));
}

function chartNodes(node: TipTapNode | undefined, found: TipTapNode[] = []): TipTapNode[] {
  if (!node || typeof node !== "object") return found;
  if ((node.type === "chart" || node.type === "inlineChart") && node.attrs) found.push(node);
  for (const child of node.content ?? []) chartNodes(child, found);
  return found;
}

/**
 * Every chart of the document drawn by the engine, by its request key. A
 * chart the engine could not draw (offline, refused) is left out; the
 * projection then refuses the export with a reason.
 */
export async function prefetchDocumentExportCharts(doc: TipTapNode): Promise<DocumentExportChartPictures> {
  const documentAttrs = (doc.attrs ?? undefined) as Record<string, unknown> | undefined;
  const pictures = new Map<string, ChartRenderPicture>();
  await Promise.all(
    chartNodes(doc).map(async (node) => {
      const request = documentExportChartRequest(node.attrs as Record<string, unknown>, documentAttrs);
      try {
        pictures.set(chartPictureKey(request), await requestChartPicture(request));
      } catch {
        /* the projection names the chart it cannot draw */
      }
    }),
  );
  return pictures;
}
