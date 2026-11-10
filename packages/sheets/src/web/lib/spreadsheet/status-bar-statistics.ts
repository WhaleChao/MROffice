/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which selection statistics the spreadsheet status bar shows, and when.
 *
 * Read off Microsoft Support's "Excel status bar options" and "View summary
 * data on the status bar":
 *
 *  - Excel lists six statistics, always in this order: Average, Count,
 *    Numerical Count, Minimum, Maximum, Sum. Average, Count and Sum are on by
 *    default; Numerical Count, Minimum and Maximum are off.
 *  - Right-clicking the status bar opens "Customize Status Bar", where each one
 *    is switched on or off. The choice belongs to the user, not the workbook,
 *    so it is kept per browser here rather than saved into the page.
 *  - Count is every selected cell that holds something; the other five are
 *    over the numeric cells only.
 *
 * When they show follows Excel (Microsoft Q&A, "Excel status bar not showing
 * sum"): nothing until at least two selected cells are non-empty; the numeric
 * figures only once one of them is a number, so an all-text selection reads
 * just "Count: 2". An error value in the selection still counts, but leaves
 * Average, Minimum, Maximum and Sum off, as =SUM over it would be an error.
 */
import type { SelectionAggregates } from "./grid-ops";

export const STATUS_BAR_STATISTICS = [
  "average",
  "count",
  "numericalCount",
  "minimum",
  "maximum",
  "sum",
] as const;

export type StatusBarStatistic = (typeof STATUS_BAR_STATISTICS)[number];

export type StatusBarStatisticsPreference = Readonly<
  Record<StatusBarStatistic, boolean>
>;

export const DEFAULT_STATUS_BAR_STATISTICS: StatusBarStatisticsPreference =
  Object.freeze({
    average: true,
    count: true,
    numericalCount: false,
    minimum: false,
    maximum: false,
    sum: true,
  });

/** Counts read as plain integers; the rest take the cell's number format. */
export function isStatusBarCountStatistic(id: StatusBarStatistic): boolean {
  return id === "count" || id === "numericalCount";
}

/** The figure Excel would show for `id`, or null when it shows none. */
export function statusBarStatisticValue(
  id: StatusBarStatistic,
  stats: SelectionAggregates,
): number | null {
  // Excel shows the figures for any selection wider than one cell that holds
  // a value: a filtered range whose one visible record is a number reads
  // Average, Count: 1, Sum. The one-cell selection is gated by the host.
  if (stats.count < 1) return null;
  if (id === "count") return stats.count;
  if (stats.countNumbers === 0) return null;
  if (id === "numericalCount") return stats.countNumbers;
  if (stats.errors > 0) return null;
  if (id === "average") return stats.average;
  if (id === "minimum") return stats.min;
  if (id === "maximum") return stats.max;
  return stats.sum;
}

// ---------------------------------------------------------------------------
// Per-user preference (the Customize Status Bar choices)
// ---------------------------------------------------------------------------

const STORAGE_KEY = "mrai.spreadsheet.statusBarStatistics";

let storageWritable = true;
let memoryRaw: string | null = null;
let snapshot: { raw: string | null; value: StatusBarStatisticsPreference } | null =
  null;
const listeners = new Set<() => void>();

function parsePreference(raw: string | null): StatusBarStatisticsPreference {
  if (!raw) return DEFAULT_STATUS_BAR_STATISTICS;
  let stored: unknown;
  try {
    stored = JSON.parse(raw);
  } catch {
    return DEFAULT_STATUS_BAR_STATISTICS;
  }
  if (!stored || typeof stored !== "object") return DEFAULT_STATUS_BAR_STATISTICS;
  const record = stored as Record<string, unknown>;
  const next: Record<StatusBarStatistic, boolean> = {
    ...DEFAULT_STATUS_BAR_STATISTICS,
  };
  for (const id of STATUS_BAR_STATISTICS) {
    const value = record[id];
    if (typeof value === "boolean") next[id] = value;
  }
  return Object.freeze(next);
}

function readRaw(): string | null {
  if (typeof window === "undefined") return null;
  if (!storageWritable) return memoryRaw;
  try {
    return window.localStorage.getItem(STORAGE_KEY);
  } catch {
    return memoryRaw;
  }
}

/** Stable between changes, as useSyncExternalStore requires. */
export function readStatusBarStatisticsPreference(): StatusBarStatisticsPreference {
  const raw = readRaw();
  if (snapshot?.raw !== raw) snapshot = { raw, value: parsePreference(raw) };
  return snapshot.value;
}

export function writeStatusBarStatisticsPreference(
  next: StatusBarStatisticsPreference,
): void {
  const raw = JSON.stringify(next);
  memoryRaw = raw;
  try {
    window.localStorage.setItem(STORAGE_KEY, raw);
    storageWritable = true;
  } catch {
    // Private windows and full storage still get the choice for this session.
    storageWritable = false;
  }
  for (const listener of listeners) listener();
}

/** Every open spreadsheet window follows the same choice, as in Excel, and so
 * does another tab of this browser. */
export function subscribeStatusBarStatisticsPreference(
  listener: () => void,
): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key === null || event.key === STORAGE_KEY) listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined")
      window.removeEventListener("storage", onStorage);
  };
}
