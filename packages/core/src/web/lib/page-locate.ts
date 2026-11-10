"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// "Locate" for an AI answer's page source
//
// A source pill's "Locate" opens the page window and asks the page's editor to
// bring the passage the answer read into view. The request travels as a
// window CustomEvent (PAGE_LOCATE_EVENT), dispatched right away and again a
// little later, so an editor that mounts for this very request still hears
// it. It is also parked per page (and mirrored into sessionStorage, so a
// navigation into the notebook keeps it) until an editor finds its target or
// it expires.
//
// The editor resolves the target in its own DOM, most exact first: the block
// with the source's `data-block-id`, else the N-th top-level block, else the
// first block whose text contains the start of the snippet. Nothing found is
// not an error: the page is open, which is the agreed fallback.
//
// The memory graph's evidence "view source" has its own, ProseMirror-level
// protocol (lib/evidence-locator.ts): it selects an exact quote. This one
// needs no editor command, only the rendered blocks.
// ---------------------------------------------------------------------------

export const PAGE_LOCATE_EVENT = "mrai:page-locate";
/** The second dispatch, for an editor that mounted after the first one. */
export const PAGE_LOCATE_REDISPATCH_MS = 400;
/** The class a located block carries while it flashes. */
export const PAGE_LOCATED_CLASS = "is-located";
/**
 * The attribute form of the same marker (value "true"), for a node whose
 * className React owns and a re-render would reset. styles/office-motion.css
 * flashes both.
 */
export const PAGE_LOCATED_ATTRIBUTE = "data-located";
export const PAGE_LOCATE_FLASH_MS = 1200;
/** The snippet prefix the text fallback looks for. */
export const PAGE_LOCATE_SNIPPET_PREFIX = 40;

const PENDING_TTL_MS = 30_000;
const PENDING_STORAGE_KEY = "mrai:page-locate:pending";

/** A worksheet cell: the sheet's name and an A1 reference ("B7"). */
export interface PageLocateCell {
  sheet: string;
  ref: string;
}

export interface PageLocateTarget {
  pageId: string;
  blockId: string | null;
  /** 1-based, in the page's top-level block order. */
  paragraphIndex: number | null;
  snippet: string | null;
  /** Workbook pages: the cell the source points at. Documents ignore it. */
  cell?: PageLocateCell | null;
  /** Deck pages: 0-based slide index. Documents ignore it. */
  slideIndex?: number | null;
}

export interface PageLocateRequest extends PageLocateTarget {
  /** One per request, so the second dispatch never flashes twice. */
  requestId: string;
}

interface PendingEntry {
  request: PageLocateRequest;
  expiresAt: number;
}

const pendingByPageId = new Map<string, PendingEntry>();

function normalizeLocateCell(
  cell: PageLocateCell | null | undefined,
): PageLocateCell | null {
  if (!cell || typeof cell !== "object") return null;
  const sheet = typeof cell.sheet === "string" ? cell.sheet.trim() : "";
  const ref = typeof cell.ref === "string" ? cell.ref.trim().toUpperCase() : "";
  return sheet && ref ? { sheet, ref } : null;
}
let requestCounter = 0;

function readStoredPending(): PendingEntry | null {
  try {
    const raw = window.sessionStorage.getItem(PENDING_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as PendingEntry;
    return parsed && typeof parsed.request?.pageId === "string" ? parsed : null;
  } catch {
    return null;
  }
}

function writeStoredPending(entry: PendingEntry | null): void {
  try {
    if (entry) {
      window.sessionStorage.setItem(PENDING_STORAGE_KEY, JSON.stringify(entry));
    } else {
      window.sessionStorage.removeItem(PENDING_STORAGE_KEY);
    }
  } catch {
    // Storage blocked: the in-memory copy still covers this document.
  }
}

/** Park the request for its page and broadcast it now and once more later. */
export function requestPageLocate(target: PageLocateTarget): PageLocateRequest | null {
  if (typeof window === "undefined") return null;
  const pageId = target.pageId?.trim();
  if (!pageId) return null;
  requestCounter += 1;
  const request: PageLocateRequest = {
    pageId,
    blockId: target.blockId?.trim() || null,
    paragraphIndex:
      Number.isInteger(target.paragraphIndex) && Number(target.paragraphIndex) > 0
        ? Number(target.paragraphIndex)
        : null,
    snippet: target.snippet?.trim() || null,
    requestId: `${Date.now().toString(36)}-${requestCounter}`,
  };
  // The optional workbook / deck coordinates travel only when valid, so a
  // document request keeps exactly its old shape.
  const cell = normalizeLocateCell(target.cell);
  if (cell) request.cell = cell;
  if (Number.isInteger(target.slideIndex) && Number(target.slideIndex) >= 0) {
    request.slideIndex = Number(target.slideIndex);
  }
  const entry = { request, expiresAt: Date.now() + PENDING_TTL_MS };
  pendingByPageId.set(pageId, entry);
  writeStoredPending(entry);
  const dispatch = () =>
    window.dispatchEvent(
      new CustomEvent<PageLocateRequest>(PAGE_LOCATE_EVENT, { detail: request }),
    );
  dispatch();
  window.setTimeout(dispatch, PAGE_LOCATE_REDISPATCH_MS);
  return request;
}

/** The fresh request parked for a page, if any; it stays parked. */
export function peekPendingPageLocate(pageId: string): PageLocateRequest | null {
  const key = pageId?.trim();
  if (!key || typeof window === "undefined") return null;
  let entry = pendingByPageId.get(key) ?? null;
  if (!entry) {
    const stored = readStoredPending();
    entry = stored?.request.pageId === key ? stored : null;
  }
  if (!entry) return null;
  if (entry.expiresAt < Date.now()) {
    clearPendingPageLocate(key);
    return null;
  }
  return entry.request;
}

/** Drops the page's parked request once an editor has located it. */
export function clearPendingPageLocate(pageId: string): void {
  pendingByPageId.delete(pageId);
  if (typeof window === "undefined") return;
  if (readStoredPending()?.request.pageId === pageId) writeStoredPending(null);
}

function normalizeText(value: string | null | undefined): string {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function escapeAttribute(value: string): string {
  return typeof CSS !== "undefined" && typeof CSS.escape === "function"
    ? CSS.escape(value)
    : value.replace(/["\\]/g, "\\$&");
}

/**
 * The page's top-level blocks in reading order. Blocks with identity count
 * once each (pagination cuts one block into fragments that share its id), and
 * a block inside another (a list item) is not top-level. Without identities,
 * the scope's children, looking through page sheets.
 */
export function topLevelPageBlocks(scope: Element): HTMLElement[] {
  const identified: HTMLElement[] = [];
  const seen = new Set<string>();
  for (const element of Array.from(
    scope.querySelectorAll<HTMLElement>("[data-block-id]"),
  )) {
    const outer = element.parentElement?.closest("[data-block-id]");
    if (outer && scope.contains(outer)) continue;
    const id = element.getAttribute("data-block-id") || "";
    if (seen.has(id)) continue;
    seen.add(id);
    identified.push(element);
  }
  if (identified.length > 0) return identified;

  const blocks: HTMLElement[] = [];
  const walk = (parent: Element) => {
    for (const child of Array.from(parent.children)) {
      if (child.hasAttribute("data-page-sheet")) walk(child);
      else if (child instanceof HTMLElement) blocks.push(child);
    }
  };
  walk(scope);
  return blocks;
}

/**
 * The element a locate request points at inside `scope`: the block with its
 * id, else the N-th top-level block, else the first block containing the
 * snippet's first 40 characters. `null` when none matches.
 */
export function resolvePageLocateTarget(
  scope: Element,
  target: Omit<PageLocateTarget, "pageId">,
): HTMLElement | null {
  const blockId = target.blockId?.trim();
  if (blockId) {
    const byId = scope.querySelector<HTMLElement>(
      `[data-block-id="${escapeAttribute(blockId)}"]`,
    );
    if (byId) return byId;
  }
  const blocks = topLevelPageBlocks(scope);
  const index = target.paragraphIndex;
  if (typeof index === "number" && Number.isInteger(index) && index > 0) {
    const byIndex = blocks[index - 1];
    if (byIndex) return byIndex;
  }
  const needle = normalizeText(target.snippet)
    .replace(/^[\s"'“”‘’「」『』…]+/, "")
    .slice(0, PAGE_LOCATE_SNIPPET_PREFIX)
    .trim();
  if (!needle) return null;
  return blocks.find((block) => normalizeText(block.textContent).includes(needle)) ?? null;
}

export interface RevealLocatedElementOptions {
  /**
   * How the flash marks the element: the `is-located` class (default) or
   * data-located="true", for a node whose className React re-renders.
   */
  marker?: "class" | "attribute";
  /**
   * Where the element scrolls to first: the middle of its scroller (default),
   * just into view ("nearest"), or not at all ("none", or `false`, for an
   * element on a fixed stage such as a slide, or one the editor has already
   * placed itself, as the PDF reader's page jump does).
   */
  scroll?: "center" | "nearest" | "none" | false;
}

/** Scrolls the element to the middle of its scroller (unless `scroll` is
 * false) and flashes it once. */
export function revealLocatedElement(
  element: HTMLElement,
  options: RevealLocatedElementOptions = {},
): () => void {
  const reduceMotion =
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const scroll = options.scroll === false ? "none" : (options.scroll ?? "center");
  if (scroll !== "none" && typeof element.scrollIntoView === "function") {
    element.scrollIntoView({
      block: scroll,
      behavior: reduceMotion ? "auto" : "smooth",
    });
  }
  const mark =
    options.marker === "attribute"
      ? () => element.setAttribute(PAGE_LOCATED_ATTRIBUTE, "true")
      : () => element.classList.add(PAGE_LOCATED_CLASS);
  const unmark =
    options.marker === "attribute"
      ? () => element.removeAttribute(PAGE_LOCATED_ATTRIBUTE)
      : () => element.classList.remove(PAGE_LOCATED_CLASS);
  // Restart the flash when the same block is located again.
  unmark();
  void element.offsetWidth;
  mark();
  const timer = window.setTimeout(unmark, PAGE_LOCATE_FLASH_MS);
  return () => {
    window.clearTimeout(timer);
    unmark();
  };
}
