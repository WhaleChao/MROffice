/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const CANVAS_COVERED_CHANGED_EVENT = "mrnote:canvas-covered-changed";

let coveredLeft = 0;
/** Each floating piece's reach: the sidebar card, and the pages drawer
 * beside it while it is open. The canvas is covered as far as the furthest. */
const reaches = new Map<string, number>();

/** The sidebar's reach from the window's left edge, in screen pixels; null when it goes. */
export function publishCanvasCoveredLeft(
  px: number | null,
  source: "sidebar" | "drawer" = "sidebar",
): void {
  if (typeof document === "undefined") return;
  if (px === null) reaches.delete(source);
  else reaches.set(source, Math.max(0, px));
  const next = Math.max(0, ...reaches.values());
  if (next === coveredLeft) return;
  coveredLeft = next;
  window.dispatchEvent(new Event(CANVAS_COVERED_CHANGED_EVENT));
}

/**
 * How much of this viewport's left edge is covered, in screen pixels: the
 * sidebar's reach past the viewport's own left edge, never more than half
 * the viewport (a narrow window keeps somewhere to put things).
 */
export function canvasCoveredLeft(viewport: Element | null | undefined): number {
  if (!viewport || coveredLeft <= 0) return 0;
  const reach = coveredLeft - viewport.getBoundingClientRect().left;
  return Math.max(0, Math.min(reach, viewport.clientWidth / 2));
}

/**
 * The sidebar card's own reach over this viewport, without the pages drawer
 * that may lie over the canvas beside it: what something centred on the
 * canvas keeps clear of (the empty canvas's home card), where clearing the
 * drawer too would squeeze it.
 */
export function canvasSidebarCoveredLeft(
  viewport: Element | null | undefined,
): number {
  const sidebar = reaches.get("sidebar") ?? 0;
  if (!viewport || sidebar <= 0) return 0;
  const reach = sidebar - viewport.getBoundingClientRect().left;
  return Math.max(0, Math.min(reach, viewport.clientWidth / 2));
}
