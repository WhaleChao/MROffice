/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Show Source Documents (Word: Review > Compare > Show Source Documents).
 *
 * After Compare or Combine, Word's review window holds the result in the
 * middle and the two source documents beside it, the original above the
 * revised (Microsoft Support, "Combine document revisions": "the middle
 * section shows the combined document, and the right section shows the
 * Original Document and Revised Document"). The submenu chooses which
 * sources show (VBA `Window.ShowSourceDocuments`, `WdShowSourceDocuments`):
 * Hide Source Documents (`wdShowSourceDocumentsNone`), Show Original,
 * Show Revised, Show Both (`wdShowSourceDocumentsBoth`). It is a property of
 * the result's window, and Word greys the submenu out where no compare or
 * combine result is open. The three panes scroll together.
 *
 * Pure parts only (no DOM): the modes, and the synchronized scroll position.
 * The panes themselves are `components/console/editor/SourceDocumentsPane.tsx`.
 */

import type { PMNodeJSON } from "./compare";

export type SourceDocumentsMode = "none" | "original" | "revised" | "both";

export const SOURCE_DOCUMENTS_MODES: readonly SourceDocumentsMode[] = [
  "none",
  "original",
  "revised",
  "both",
];

/** Word opens a compare/combine result with both sources showing. */
export const DEFAULT_SOURCE_DOCUMENTS_MODE: SourceDocumentsMode = "both";

export type SourceDocumentRole = "original" | "revised";

export interface SourceDocument {
  /** What the person picked: "Current document" or a version's name. */
  label: string;
  /** The document as it was compared (stored form). */
  doc: PMNodeJSON;
}

/** A compare or combine result whose two sources are known. */
export interface CompareResultSources {
  /** One result view; a new compare or combine is a new one. */
  id: string;
  kind: "compare" | "combine";
  original: SourceDocument;
  revised: SourceDocument;
}

/** The source panes a mode shows, top to bottom (original above revised). */
export function sourceDocumentPanes(
  mode: SourceDocumentsMode,
): SourceDocumentRole[] {
  switch (mode) {
    case "original":
      return ["original"];
    case "revised":
      return ["revised"];
    case "both":
      return ["original", "revised"];
    default:
      return [];
  }
}

export function isSourceDocumentsMode(value: unknown): value is SourceDocumentsMode {
  return SOURCE_DOCUMENTS_MODES.includes(value as SourceDocumentsMode);
}

// ---------------------------------------------------------------------------
// Synchronized scrolling
// ---------------------------------------------------------------------------

/**
 * One block as laid out in a scroll container: its block id and its box in
 * the container's scroll coordinates (`top` 0 = the top of the content). A
 * block the pages cut into fragments appears once per fragment, under the
 * same id (pagination copies a block's attrs onto each fragment).
 */
export interface ScrollAnchor {
  id: string;
  top: number;
  height: number;
}

export interface ScrollView {
  anchors: readonly ScrollAnchor[];
  scrollTop: number;
  scrollHeight: number;
  clientHeight: number;
}

/** Where in the driving view the panes line up: a little below its top. */
export const SOURCE_SCROLL_PROBE_PX = 24;

function maxScroll(view: Pick<ScrollView, "scrollHeight" | "clientHeight">): number {
  return Math.max(0, view.scrollHeight - view.clientHeight);
}

function fragments(anchors: readonly ScrollAnchor[], id: string): ScrollAnchor[] {
  return anchors
    .filter((anchor) => anchor.id === id && anchor.height > 0)
    .sort((left, right) => left.top - right.top);
}

/**
 * The block at `y` and how far into it `y` is (0..1 over all its fragments).
 * The innermost block wins (a list item, not its whole list), so the panes
 * line up as finely as the two layouts allow.
 */
export function anchorAt(
  anchors: readonly ScrollAnchor[],
  y: number,
): { id: string; fraction: number } | null {
  let best: ScrollAnchor | null = null;
  for (const anchor of anchors) {
    if (anchor.height <= 0) continue;
    if (y < anchor.top || y >= anchor.top + anchor.height) continue;
    if (!best || anchor.height < best.height) best = anchor;
  }
  if (!best) {
    // Between blocks (a page gap, a margin): the nearest block above.
    for (const anchor of anchors) {
      if (anchor.height <= 0 || anchor.top + anchor.height > y) continue;
      if (!best || anchor.top + anchor.height > best.top + best.height) best = anchor;
    }
    if (!best) return null;
    return { id: best.id, fraction: 1 };
  }
  const parts = fragments(anchors, best.id);
  const total = parts.reduce((sum, part) => sum + part.height, 0);
  let before = 0;
  for (const part of parts) {
    if (part === best) break;
    before += part.height;
  }
  const fraction = total > 0 ? (before + (y - best.top)) / total : 0;
  return { id: best.id, fraction: Math.min(1, Math.max(0, fraction)) };
}

/** The y of `fraction` through block `id`'s fragments, or null when absent. */
function positionIn(
  anchors: readonly ScrollAnchor[],
  id: string,
  fraction: number,
): number | null {
  const parts = fragments(anchors, id);
  if (parts.length === 0) return null;
  const total = parts.reduce((sum, part) => sum + part.height, 0);
  let offset = fraction * total;
  for (const part of parts) {
    if (offset <= part.height) return part.top + offset;
    offset -= part.height;
  }
  const last = parts[parts.length - 1];
  return last.top + last.height;
}

/**
 * The scroll position that shows, at the probe line, the same place of the
 * document the driving view shows there. The compared documents share their
 * block structure (Compare and Combine refuse otherwise) and so their block
 * ids; where a block is not found, the panes keep the same proportion.
 */
export function syncedScrollTop(
  driver: ScrollView,
  target: ScrollView,
  probe = SOURCE_SCROLL_PROBE_PX,
): number {
  const targetMax = maxScroll(target);
  const driverMax = maxScroll(driver);
  const proportional = () =>
    driverMax > 0 ? (driver.scrollTop / driverMax) * targetMax : 0;
  if (driver.scrollTop <= 0) return 0;
  if (driverMax > 0 && driver.scrollTop >= driverMax - 0.5) return targetMax;
  const probeY = driver.scrollTop + probe;
  const at = anchorAt(driver.anchors, probeY);
  const y = at ? positionIn(target.anchors, at.id, at.fraction) : null;
  const top = y === null ? proportional() : y - probe;
  return Math.min(targetMax, Math.max(0, top));
}
