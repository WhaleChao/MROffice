/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A deck imported before its charts' native appearance was stored
 * (`source.chartDisplay`, written by the server import since): the server
 * reads that appearance from the deck's original package
 * (`GET /api/v1/pages/{id}/presentation-chart-displays`) and the editor
 * stores it in the blocks ONCE -- the representation-upgrade pattern: outside
 * the undo history, not the member's own edit, saved with the page by the
 * next authored save. The Office XML reading is the server's; the browser
 * only matches blocks by their native slide and shape identity.
 */

import { apiGet } from "@mrnote/core/web/lib/api.ts";
import {
  parsePresentationChartDisplay,
  type PresentationChartDisplay,
} from "@mrnote/core/web/lib/presentation-chart-display.ts";

export interface PresentationChartDisplaysAnswer {
  page_id: string;
  content_version: number;
  state: "resolved" | "unavailable";
  /** Keyed `sld<id>:shp<id>`, as a block's `source.slideId` / `source.shapeId`. */
  displays: Record<string, unknown>;
}

export function fetchPresentationChartDisplays(
  pageId: string,
  signal?: AbortSignal,
): Promise<PresentationChartDisplaysAnswer> {
  return apiGet<PresentationChartDisplaysAnswer>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/presentation-chart-displays`,
    { signal },
  );
}

type ChartLike = {
  type: string;
  source?: { slideId?: unknown; shapeId?: unknown; chartDisplay?: unknown } & Record<string, unknown>;
};

function chartSourceKey(block: ChartLike): string | null {
  const slideId = block.source?.slideId;
  const shapeId = block.source?.shapeId;
  return typeof slideId === "string" && slideId && typeof shapeId === "string" && shapeId
    ? `${slideId}:${shapeId}`
    : null;
}

/** An imported chart that states no appearance of its own yet. */
export function chartBlockNeedsRecoveredDisplay(block: ChartLike): boolean {
  return block.type === "chart" && !block.source?.chartDisplay && chartSourceKey(block) !== null;
}

/**
 * The slides with every chart that lacks an appearance given the one read for
 * its native identity; null when nothing changes. A chart that has one keeps
 * it; a display the schema does not admit is ignored.
 */
export function slidesWithRecoveredChartDisplays<S extends { blocks: readonly ChartLike[] }>(
  slides: readonly S[],
  displays: Record<string, unknown>,
): S[] | null {
  const parsed = new Map<string, PresentationChartDisplay>();
  for (const [key, value] of Object.entries(displays)) {
    const display = parsePresentationChartDisplay(value);
    if (display) parsed.set(key, display);
  }
  if (parsed.size === 0) return null;
  let changed = false;
  const next = slides.map((slide) => {
    let slideChanged = false;
    const blocks = slide.blocks.map((block) => {
      if (!chartBlockNeedsRecoveredDisplay(block)) return block;
      const display = parsed.get(chartSourceKey(block)!);
      if (!display) return block;
      slideChanged = true;
      return { ...block, source: { ...block.source, chartDisplay: display } };
    });
    if (!slideChanged) return slide;
    changed = true;
    return { ...slide, blocks } as S;
  });
  return changed ? next : null;
}
