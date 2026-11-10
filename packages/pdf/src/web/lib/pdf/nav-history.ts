/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure view-history model for the PDF reader (browser-style Back / Forward).
 *
 * Every programmatic jump (TOC click, link, page input, citation) and every
 * LARGE manual scroll leap (scrollbar drag across pages) records the position
 * being LEFT — `{page, scrollTop}` — on a bounded back stack; going back moves
 * the current position onto the forward stack, mirroring pdf.js's
 * `PDFHistory` (Apache-2.0) in a minimal in-memory form (no URL hash state).
 *
 * The model is immutable-update / DOM-free so the transitions unit-test in
 * isolation; the component keeps the state in a ref and re-syncs the toolbar
 * buttons' disabled flags after each mutation.
 */

/** Max entries kept per stack (oldest evicted first). */
export const PDF_NAV_HISTORY_CAP = 50;

/**
 * How long (ms) after a programmatic scroll the manual-scroll heuristic stays
 * suppressed, so a smooth animated jump is never re-recorded as a user leap.
 * Mirrors pdf.js's UPDATE_VIEWAREA_TIMEOUT idea at a shorter horizon.
 */
export const PDF_NAV_SCROLL_SUPPRESS_MS = 1000;

/**
 * A manual scroll counts as a history-worthy LEAP when a single scroll event
 * moves more than this many viewport heights (scrollbar drags / native
 * Home-End jumps produce one big delta; wheel/trackpad reading produces many
 * small ones that must NOT pollute the history).
 */
export const PDF_NAV_SCROLL_JUMP_FACTOR = 1.5;

/** One remembered reading position. */
export interface PdfNavEntry {
  /** 1-based page number the reader was on. */
  page: number;
  /** Scroll container scrollTop (px) at that moment. */
  scrollTop: number;
  




  scale?: number;
  /** The sticky fit mode in force ("width" / "page"), or null for none. */
  fitMode?: "width" | "page" | null;
}

/** The two history stacks. Treat as immutable; helpers return new objects. */
export interface PdfNavHistory {
  back: PdfNavEntry[];
  forward: PdfNavEntry[];
}

export function emptyNavHistory(): PdfNavHistory {
  return { back: [], forward: [] };
}

/** Validate + normalize an entry; null when unusable. */
function sanitizeNavEntry(entry: unknown): PdfNavEntry | null {
  if (!entry || typeof entry !== "object") return null;
  const { page, scrollTop } = entry as { page?: unknown; scrollTop?: unknown };
  if (typeof page !== "number" || !Number.isInteger(page) || page < 1) {
    return null;
  }
  const top =
    typeof scrollTop === "number" && Number.isFinite(scrollTop) && scrollTop > 0
      ? scrollTop
      : 0;
  const { scale, fitMode } = entry as { scale?: unknown; fitMode?: unknown };
  const clean: PdfNavEntry = { page, scrollTop: top };
  if (typeof scale === "number" && Number.isFinite(scale) && scale > 0) {
    clean.scale = scale;
    clean.fitMode = fitMode === "width" || fitMode === "page" ? fitMode : null;
  }
  return clean;
}

/** Whether two entries are the same position (same page, ~same scrollTop). */
function sameNavEntry(a: PdfNavEntry, b: PdfNavEntry): boolean {
  return (
    a.page === b.page &&
    Math.abs(a.scrollTop - b.scrollTop) < 1 &&
    (a.scale === undefined ||
      b.scale === undefined ||
      (Math.abs(a.scale - b.scale) < 1e-6 && a.fitMode === b.fitMode))
  );
}

/**
 * Record the position being LEFT before a jump. Clears the forward stack
 * (branching invalidates it, like a browser), dedupes an entry identical to
 * the current stack top (repeat jumps from the same spot produce one step),
 * and evicts the oldest entry past {@link PDF_NAV_HISTORY_CAP}. An invalid
 * entry returns the input unchanged.
 */
export function pushNavEntry(
  history: PdfNavHistory,
  entry: PdfNavEntry,
  cap: number = PDF_NAV_HISTORY_CAP,
): PdfNavHistory {
  const clean = sanitizeNavEntry(entry);
  if (!clean) return history;
  const limit = Number.isInteger(cap) && cap > 0 ? cap : PDF_NAV_HISTORY_CAP;
  const top = history.back[history.back.length - 1];
  if (top && sameNavEntry(top, clean)) {
    return history.forward.length === 0
      ? history
      : { back: history.back, forward: [] };
  }
  return { back: [...history.back, clean].slice(-limit), forward: [] };
}

export function canNavBack(history: PdfNavHistory): boolean {
  return history.back.length > 0;
}

export function canNavForward(history: PdfNavHistory): boolean {
  return history.forward.length > 0;
}

/**
 * Step BACK: pop the back stack as the navigation target and push the CURRENT
 * position onto the forward stack (so Forward can return). Returns the new
 * history plus the target, or `null` when there is nowhere to go (state
 * untouched). The forward stack shares the same cap.
 */
export function navBack(
  history: PdfNavHistory,
  current: PdfNavEntry,
  cap: number = PDF_NAV_HISTORY_CAP,
): { history: PdfNavHistory; target: PdfNavEntry } | null {
  if (history.back.length === 0) return null;
  const limit = Number.isInteger(cap) && cap > 0 ? cap : PDF_NAV_HISTORY_CAP;
  const target = history.back[history.back.length - 1];
  const now = sanitizeNavEntry(current);
  const forward = now
    ? [...history.forward, now].slice(-limit)
    : history.forward;
  return { history: { back: history.back.slice(0, -1), forward }, target };
}

/**
 * Step FORWARD: the mirror of {@link navBack} — pop the forward stack as the
 * target and push the current position onto the back stack.
 */
export function navForward(
  history: PdfNavHistory,
  current: PdfNavEntry,
  cap: number = PDF_NAV_HISTORY_CAP,
): { history: PdfNavHistory; target: PdfNavEntry } | null {
  if (history.forward.length === 0) return null;
  const limit = Number.isInteger(cap) && cap > 0 ? cap : PDF_NAV_HISTORY_CAP;
  const target = history.forward[history.forward.length - 1];
  const now = sanitizeNavEntry(current);
  const back = now ? [...history.back, now].slice(-limit) : history.back;
  return { history: { back, forward: history.forward.slice(0, -1) }, target };
}

/**
 * Whether a single scroll event's delta is a history-worthy LEAP: it moved
 * more than {@link PDF_NAV_SCROLL_JUMP_FACTOR} viewport heights in one go.
 * Degenerate inputs (non-finite, non-positive viewport) are never leaps.
 */
export function isLargeScrollJump(
  fromScrollTop: number,
  toScrollTop: number,
  viewportHeight: number,
): boolean {
  if (
    !Number.isFinite(fromScrollTop) ||
    !Number.isFinite(toScrollTop) ||
    !Number.isFinite(viewportHeight) ||
    viewportHeight <= 0
  ) {
    return false;
  }
  return (
    Math.abs(toScrollTop - fromScrollTop) >
    viewportHeight * PDF_NAV_SCROLL_JUMP_FACTOR
  );
}
