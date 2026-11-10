/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The chart element each editor's chart has selected (the Format pane's
 * Current Selection), by the host's key for the chart: a click on the title,
 * the legend, an axis, its labels or the data labels selects that element
 * (`chart-hit`), the pane's own pick moves it, and deselecting the chart
 * clears it. Home's Font group formats the element (`chart-element-text`)
 * and the chart's picture frames it, in Docs, the deck and the worksheet
 * alike, as Office formats and frames the element selected.
 */
import { useSyncExternalStore } from "react";

import type { ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";
import { officeChartHitSelection, type OfficeChartHit } from "./chart-hit";

const areas = new Map<string, string>();
/**
 * What inside the selected element was clicked, beside its area: one legend
 * entry (the legend selected, an entry clicked again), or the data labels a
 * label click selected -- one series' (its area `series:S`) or one label
 * (`point:S:P`, `slice:P`). The Format pane reads the area as before (the
 * legend's, the series' labels', the point's label); the picture frames the
 * entry or the labels, and Delete removes them rather than the element.
 */
export interface OfficeChartSelectionDetail {
  readonly legendEntry?: number;
  readonly labels?: { readonly series: number; readonly point?: number };
  /** A selected series (every point) or one point, in the picture's own
   * indices (`ecmeta_series_index` / `ecmeta_data_index`). */
  readonly points?: { readonly series: number; readonly point?: number };
}
const details = new Map<string, OfficeChartSelectionDetail>();
const listeners = new Set<() => void>();

function notify(): void {
  for (const listener of listeners) listener();
}

/** The selected element's Format pane area for the chart, or null (the chart area). */
export function officeChartSelectedArea(key: string | null | undefined): string | null {
  return key ? (areas.get(key) ?? null) : null;
}

/** Select an element of the chart (null: back to the chart itself), and
 * what inside it was clicked (`detail`; none for a plain element). */
export function setOfficeChartSelectedArea(
  key: string | null | undefined,
  area: string | null,
  detail: OfficeChartSelectionDetail | null = null,
): void {
  if (!key) return;
  const current = areas.get(key) ?? null;
  const sameDetail = JSON.stringify(details.get(key) ?? null) === JSON.stringify(detail);
  if (current === area && sameDetail) return;
  if (area === null || area === "chartAreaStyle") areas.delete(key);
  else areas.set(key, area);
  if (detail && area !== null && area !== "chartAreaStyle") details.set(key, detail);
  else details.delete(key);
  notify();
}

/** What inside the selected element was clicked, if anything. */
export function officeChartSelectionDetail(key: string | null | undefined): OfficeChartSelectionDetail | null {
  return key ? (details.get(key) ?? null) : null;
}

/** Forget every chart of one host (its keys start with `prefix`) but those
 * named: the host's selection moved on, and Office starts a chart selected
 * again at its chart area. */
export function keepOfficeChartSelections(prefix: string, keys: Iterable<string>): void {
  const kept = new Set(keys);
  let changed = false;
  for (const key of [...areas.keys()]) {
    if (key.startsWith(prefix) && !kept.has(key)) {
      areas.delete(key);
      details.delete(key);
      changed = true;
    }
  }
  if (changed) notify();
}

const hostKeys = new WeakMap<object, string>();
let hostSerial = 0;

/** A stable key prefix for a host object (an editor, a window), so two
 * windows' charts with the same id are two charts here. */
export function officeChartHostKey(host: object): string {
  let key = hostKeys.get(host);
  if (!key) {
    hostSerial += 1;
    key = `h${hostSerial}:`;
    hostKeys.set(host, key);
  }
  return key;
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Select what a click on a chart landed on (`officeChartHitSelection`). */
export function selectOfficeChartHit(key: string | null | undefined, attrs: ChartBlockAttrsV2, hit: OfficeChartHit): void {
  const next = officeChartHitSelection(attrs, hit, officeChartSelectedArea(key), officeChartSelectionDetail(key));
  setOfficeChartSelectedArea(key, next.area, next.detail);
}

/** What inside the selected element was clicked, live. */
export function useOfficeChartSelectionDetail(key: string | null | undefined): OfficeChartSelectionDetail | null {
  return useSyncExternalStore(
    subscribe,
    () => officeChartSelectionDetail(key),
    () => null,
  );
}

/** The chart's selected element area, live. */
export function useOfficeChartSelectedArea(key: string | null | undefined): string | null {
  return useSyncExternalStore(
    subscribe,
    () => officeChartSelectedArea(key),
    () => null,
  );
}
