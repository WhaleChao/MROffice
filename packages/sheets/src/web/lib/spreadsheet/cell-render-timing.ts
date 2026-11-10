/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a grid cell's render spends its time, for a probe: when
 * `window.__gridCellTiming` is an object, each mark adds the time since the
 * previous one to its label. Off (one property read per cell) otherwise.
 */
let marks: Record<string, number> | null = null;
let last = 0;

export function cellTimingStart(): void {
  const probe = typeof window === "undefined" ? undefined : (window as unknown as { __gridCellTiming?: unknown }).__gridCellTiming;
  marks = probe && typeof probe === "object" ? (probe as Record<string, number>) : null;
  if (marks) {
    last = performance.now();
    marks.cells = (marks.cells ?? 0) + 1;
  }
}

export function cellTimingMark(label: string): void {
  if (!marks) return;
  const now = performance.now();
  marks[label] = (marks[label] ?? 0) + now - last;
  last = now;
}

/** Adds one to `reason` in `window.__gridCanvasRowsWhy`, when a probe set it to an object. */
export function tallyCanvasProbe(reason: string): void {
  const tally = typeof window === "undefined" ? undefined : (window as unknown as { __gridCanvasRowsWhy?: Record<string, number> }).__gridCanvasRowsWhy;
  if (tally && typeof tally === "object") tally[reason] = (tally[reason] ?? 0) + 1;
}
