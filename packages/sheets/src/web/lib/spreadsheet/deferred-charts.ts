/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet } from "@mrnote/core/web/lib/api.ts";
import { rememberFilledChart } from "./deferred-chart-placeholders";

/**
 * Charts imported with their models deferred (CHART-DEFER-DESIGN.md): each
 * stands on the page as a placeholder (`chartPending`) with the id its whole
 * object will have; the page open asks the server for those objects (read
 * off the page's original package) and fills them in for this window; what is
 * stored stays the placeholder until the chart itself changes
 * (deferred-chart-placeholders).
 */
type SheetLike = { id: string; objects?: ReadonlyArray<Record<string, unknown>> };

export function pendingChartIds(sheets: readonly SheetLike[]): string[] {
  const ids: string[] = [];
  for (const sheet of sheets) {
    for (const object of sheet.objects ?? []) {
      if (object?.type === "chart" && object.chartPending === true && typeof object.id === "string") ids.push(object.id);
    }
  }
  return ids;
}








export function deferredChartKey(sheetId: string, objectId: string): string {
  return `${sheetId}\u0000${objectId}`;
}

export type PendingChart = { key: string; sheetId: string; id: string };

export function pendingCharts(sheets: readonly SheetLike[]): PendingChart[] {
  const found: PendingChart[] = [];
  for (const sheet of sheets) {
    for (const object of sheet.objects ?? []) {
      if (object?.type === "chart" && object.chartPending === true && typeof object.id === "string") {
        found.push({ key: deferredChartKey(sheet.id, object.id), sheetId: sheet.id, id: object.id });
      }
    }
  }
  return found;
}

type ChartObjectsAnswer = {
  objects?: Record<string, Record<string, unknown>>;
  objectsBySheet?: Record<string, Record<string, Record<string, unknown>>>;
};

/** The whole objects of these placeholders, by `deferredChartKey`. */
export function deferredChartObjectsByKey(answer: ChartObjectsAnswer | null | undefined, asked: readonly PendingChart[]): Record<string, Record<string, unknown>> {
  const out: Record<string, Record<string, unknown>> = {};
  const bySheet = answer?.objectsBySheet;
  const byId = answer?.objects;
  for (const chart of asked) {
    const whole = bySheet && typeof bySheet === "object"
      ? bySheet[chart.sheetId]?.[chart.id]
      // A server from before the answer by sheet: its objects by id alone.
      : byId && typeof byId === "object" ? byId[chart.id] : undefined;
    if (whole && typeof whole === "object") out[chart.key] = whole;
  }
  return out;
}

export async function fetchDeferredChartObjects(pageId: string, asked: readonly PendingChart[], signal?: AbortSignal): Promise<Record<string, Record<string, unknown>>> {
  const ids = [...new Set(asked.map((chart) => chart.id))];
  // One sheet's (the shown one's first): named, so the server reads that
  
  const sheets = new Set(asked.map((chart) => chart.sheetId));
  const sheet = sheets.size === 1 ? `&sheet=${encodeURIComponent([...sheets][0])}` : "";
  const answer = await apiGet<ChartObjectsAnswer>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/chart-objects?ids=${encodeURIComponent(ids.join(","))}${sheet}`,
    signal ? { signal } : undefined,
  );
  return deferredChartObjectsByKey(answer, asked);
}

/**
 * The sheets with each placeholder replaced by its whole object (`filled`, by
 * `deferredChartKey`), standing where the placeholder stands now; the same array when none is filled in.
 * `normalize` brings a whole object to the page's own shape.
 */
export function withDeferredChartsFilled<S extends SheetLike>(
  sheets: readonly S[],
  filled: Record<string, Record<string, unknown>>,
  normalize: (object: Record<string, unknown>) => Record<string, unknown> | undefined,
): readonly S[] {
  let changed = false;
  const next = sheets.map((sheet) => {
    const objects = sheet.objects;
    if (!objects?.some((object) => object?.chartPending === true)) return sheet;
    let sheetChanged = false;
    const nextObjects = objects.map((object) => {
      if (object?.chartPending !== true || typeof object.id !== "string") return object;
      const whole = filled[deferredChartKey(sheet.id, object.id)];
      if (!whole || whole.id !== object.id || whole.type !== "chart") return object;
      const normalized = normalize({ ...whole, anchor: object.anchor ?? whole.anchor });
      if (!normalized) return object;
      rememberFilledChart(normalized, object);
      sheetChanged = true;
      return normalized;
    });
    if (!sheetChanged) return sheet;
    changed = true;
    return { ...sheet, objects: nextObjects };
  });
  return changed ? next : sheets;
}
