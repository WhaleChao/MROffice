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
// first block whose text contains the start of the snippet (or that the
// snippet starts with). Nothing found is not an error: the page is open, which
// is the agreed fallback. A request that asks for it (announceMiss) says so.
//
// The memory graph's evidence "view source" has its own, ProseMirror-level
// protocol (lib/evidence-locator.ts): it selects an exact quote. This one
// needs no editor command, only the rendered blocks.
// ---------------------------------------------------------------------------

export const PAGE_LOCATE_EVENT = "mrai:page-locate";
/**
 * An editor could not find what an announceMiss request points at, or the
 * request named nothing: the workspace's notice (PageLocateMissNotice) says
 * so. An event, so the editors carry no copy of their own.
 */
export const PAGE_LOCATE_MISS_EVENT = "mrai:page-locate-miss";

export interface PageLocateMiss {
  pageId: string;
  requestId: string;
  /** false when the request named no place at all (a file-level source). */
  named: boolean;
}
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

/** A worksheet cell or range: the sheet's name (and id, when known) and an
 * A1 reference ("B7", "B7:D7"). */
export interface PageLocateCell {
  sheet: string;
  ref: string;
  sheetId?: string | null;
}

/**
 * Cells of a document's table: the table's block id and a range of its grid
 * (0-based rows and columns; a merged cell covers its whole span), as an
 * answer's citation names them (the API's citation_anchors).
 */
export interface PageLocateTableRange {
  blockId: string;
  firstRow: number;
  firstColumn: number;
  lastRow: number;
  lastColumn: number;
}

export interface PageLocateTarget {
  pageId: string;
  blockId: string | null;
  /** 1-based, in the page's top-level block order. */
  paragraphIndex: number | null;
  snippet: string | null;
  /** Workbook pages: the cell the source points at. Documents ignore it. */
  cell?: PageLocateCell | null;
  /** Document pages: the table cells the source points at. */
  table?: PageLocateTableRange | null;
  /** Deck pages: 0-based slide index. Documents ignore it. */
  slideIndex?: number | null;
  /**
   * Deck pages: the slide's id, which outlasts a slide moved or inserted
   * before it (slideIndex is then only the fallback); `blockId` names the
   * object on it and `table` cells of a table object.
   */
  slideId?: string | null;
  /** PDF pages: the 1-based page the passage is on. Others ignore it. */
  pdfPage?: number | null;
  






  announceMiss?: boolean;
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
  if (!sheet || !ref) return null;
  const sheetId = typeof cell.sheetId === "string" && cell.sheetId.trim() ? cell.sheetId.trim() : null;
  return sheetId ? { sheet, ref, sheetId } : { sheet, ref };
}

function normalizeTableRange(
  table: PageLocateTableRange | null | undefined,
): PageLocateTableRange | null {
  if (!table || typeof table !== "object") return null;
  const blockId = typeof table.blockId === "string" ? table.blockId.trim() : "";
  const bounds = [table.firstRow, table.firstColumn, table.lastRow, table.lastColumn];
  if (!blockId || !bounds.every((value) => Number.isInteger(value) && value >= 0)) return null;
  return {
    blockId,
    firstRow: Math.min(table.firstRow, table.lastRow),
    firstColumn: Math.min(table.firstColumn, table.lastColumn),
    lastRow: Math.max(table.firstRow, table.lastRow),
    lastColumn: Math.max(table.firstColumn, table.lastColumn),
  };
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
  const table = normalizeTableRange(target.table);
  if (table) request.table = table;
  if (Number.isInteger(target.slideIndex) && Number(target.slideIndex) >= 0) {
    request.slideIndex = Number(target.slideIndex);
  }
  const slideId = typeof target.slideId === "string" ? target.slideId.trim() : "";
  if (slideId) request.slideId = slideId;
  if (Number.isInteger(target.pdfPage) && Number(target.pdfPage) >= 1) {
    request.pdfPage = Number(target.pdfPage);
  }
  if (target.announceMiss) request.announceMiss = true;
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

/** Whether a request names anything to look for in the page. */
export function pageLocateHasTarget(target: Omit<PageLocateTarget, "pageId">): boolean {
  return Boolean(
    target.blockId?.trim() ||
      (Number.isInteger(target.paragraphIndex) && Number(target.paragraphIndex) > 0) ||
      target.snippet?.trim() ||
      normalizeLocateCell(target.cell) ||
      normalizeTableRange(target.table) ||
      (Number.isInteger(target.slideIndex) && Number(target.slideIndex) >= 0) ||
      Boolean(target.slideId?.trim()) ||
      (Number.isInteger(target.pdfPage) && Number(target.pdfPage) >= 1),
  );
}

/** Shortest block text the snippet's head may stand for on its own. */
const SNIPPET_HEAD_BLOCK_MIN = 6;

function normalizeText(value: string | null | undefined): string {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

/** Text without spaces or leading quotes, for comparing across joined runs. */
function compactText(value: string | null | undefined): string {
  return String(value ?? "")
    .replace(/^[\s"'“”‘’「」『』…]+/, "")
    .replace(/\s+/g, "");
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
  const texts = blocks.map((block) => normalizeText(block.textContent));
  const containing = texts.findIndex((text) => text.includes(needle));
  if (containing >= 0) return blocks[containing];
  // The same compared without spaces: a block's text joins its cells and
  // runs with none ("季度经营报告2026 年…" for a title table), where the
  // snippet, read from the page's plain text, has one between them.
  const compactNeedle = compactText(needle);
  const compactTexts = texts.map(compactText);
  const compactContaining = compactNeedle
    ? compactTexts.findIndex((text) => text.includes(compactNeedle))
    : -1;
  if (compactContaining >= 0) return blocks[compactContaining];
  // A snippet that runs across blocks (a title, its subtitle and the first
  // paragraph, as a page read's opening passage does) is in no single one:
  // the block it starts with is where it is, the outermost first (a table
  // laid out as a title), then one inside it.
  const head = compactText(target.snippet);
  const starting = compactTexts.findIndex(
    (text) => text.length >= SNIPPET_HEAD_BLOCK_MIN && head.startsWith(text),
  );
  if (starting >= 0) return blocks[starting];
  const nested = Array.from(scope.querySelectorAll<HTMLElement>("[data-block-id]")).find((block) => {
    const text = compactText(block.textContent);
    return text.length >= SNIPPET_HEAD_BLOCK_MIN && head.startsWith(text);
  });
  return nested ?? null;
}

/**
 * The cells of a table range, in the page as the editor draws it: the table
 * with the block id (a table a page break splits draws as fragments sharing
 * it, read in order), its grid worked out from each row's cells and their
 * spans, as the API's citation_anchors works it out. Empty when the table or
 * the range is not there (the document changed since the answer).
 */
export function resolveTableRange(scope: Element, range: PageLocateTableRange): HTMLElement[] {
  return tableRangeCells(
    Array.from(
      scope.querySelectorAll<HTMLElement>(`table[data-block-id="${escapeAttribute(range.blockId)}"]`),
    ),
    range,
  );
}

/**
 * The cells of `range` in a table drawn as `fragments` (one `<table>`, or the
 * pieces a page break cut it into, in order): each row's cells laid on the
 * grid with their spans, the way the API's citation_anchors lays a stored
 * table out (a merged cell's covered cells are not in its row).
 */
export function tableRangeCells(
  fragments: readonly HTMLElement[],
  range: Omit<PageLocateTableRange, "blockId">,
): HTMLElement[] {
  if (fragments.length === 0) return [];
  const rows: HTMLTableRowElement[] = [];
  for (const fragment of fragments) {
    for (const row of Array.from(fragment.querySelectorAll<HTMLTableRowElement>("tr"))) {
      if (row.closest("table") !== fragment) continue;
      // A header row a page break repeats is a copy, not a row of its own.
      if (rows.length > 0 && row.closest("table") !== fragments[0] && row.hasAttribute("data-repeated-header")) {
        continue;
      }
      rows.push(row);
    }
  }
  const occupied = new Set<string>();
  const found: HTMLElement[] = [];
  rows.forEach((row, rowIndex) => {
    let column = 0;
    for (const cell of Array.from(row.children)) {
      if (!(cell instanceof HTMLElement) || !/^(TD|TH)$/.test(cell.tagName)) continue;
      while (occupied.has(`${rowIndex}:${column}`)) column += 1;
      const rowSpan = Math.max(1, Number(cell.getAttribute("rowspan")) || 1);
      const columnSpan = Math.max(1, Number(cell.getAttribute("colspan")) || 1);
      for (let r = rowIndex; r < rowIndex + rowSpan; r += 1) {
        for (let c = column; c < column + columnSpan; c += 1) occupied.add(`${r}:${c}`);
      }
      const overlaps =
        rowIndex <= range.lastRow && rowIndex + rowSpan - 1 >= range.firstRow &&
        column <= range.lastColumn && column + columnSpan - 1 >= range.firstColumn;
      if (overlaps) found.push(cell);
      column += columnSpan;
    }
  });
  return found;
}

/** Scrolls the first element into view and flashes them all at once. */
export function revealLocatedElements(
  elements: readonly HTMLElement[],
  options: RevealLocatedElementOptions = {},
): () => void {
  const stops = elements.map((element, index) =>
    revealLocatedElement(element, index === 0 ? options : { ...options, scroll: "none" }),
  );
  return () => stops.forEach((stop) => stop());
}

/** A source locator as the API sends it, read for a locate request. */
export interface LocatorForLocate {
  page_id: string;
  block_id: string | null;
  paragraph_index: number | null;
  anchor?: {
    kind: string;
    /** Absent on a slide anchor that names no table cells. */
    ref?: string;
    table_block_id?: string | null;
    first_row?: number;
    first_column?: number;
    last_row?: number;
    last_column?: number;
    sheet_id?: string | null;
    sheet_name?: string | null;
    slide_id?: string | null;
    slide_number?: number;
    block_id?: string | null;
    page_number?: number;
    line?: string | null;
    first_line?: number;
  } | null;
}

/** The locate request a cited source's locator makes: its table cells,
 * worksheet range or slide object when it names one, its paragraph
 * otherwise. */
export function locateTargetForLocator(
  locator: LocatorForLocate,
  snippet: string | null,
): PageLocateTarget {
  const target: PageLocateTarget = {
    pageId: locator.page_id,
    blockId: locator.block_id,
    paragraphIndex: locator.paragraph_index,
    snippet,
    announceMiss: true,
  };
  const anchor = locator.anchor;
  if (anchor?.kind === "code_lines") {
    // A code page's own line wins over the source's snippet, which may be
    // another line of the same read.
    if (Number.isInteger(anchor.first_line) && Number(anchor.first_line) >= 1) {
      target.paragraphIndex = Number(anchor.first_line);
      target.snippet = null;
    }
  } else if (anchor?.kind === "pdf_page") {
    if (Number.isInteger(anchor.page_number) && Number(anchor.page_number) >= 1) {
      target.pdfPage = Number(anchor.page_number);
      // The cited line on that page, not the source's snippet (often the
      // head of another page of the same read).
      if (anchor.line?.trim()) target.snippet = anchor.line.trim();
    }
  } else if (anchor?.kind === "slide_object") {
    // A deck: its slide (by id, else by place), the object on it and, for a
    // table object, the cells.
    if (anchor.slide_id) target.slideId = anchor.slide_id;
    if (Number.isInteger(anchor.slide_number) && Number(anchor.slide_number) >= 1) {
      target.slideIndex = Number(anchor.slide_number) - 1;
    }
    target.blockId = anchor.block_id ?? null;
    target.paragraphIndex = null;
    if (
      anchor.block_id && anchor.ref &&
      [anchor.first_row, anchor.first_column, anchor.last_row, anchor.last_column].every(
        (value) => Number.isInteger(value),
      )
    ) {
      target.table = {
        blockId: anchor.block_id,
        firstRow: anchor.first_row!,
        firstColumn: anchor.first_column!,
        lastRow: anchor.last_row!,
        lastColumn: anchor.last_column!,
      };
    }
  } else if (anchor?.kind === "sheet_range" && anchor.sheet_name && anchor.ref) {
    target.cell = { sheet: anchor.sheet_name, ref: anchor.ref, sheetId: anchor.sheet_id ?? null };
  } else if (
    anchor?.kind === "table_range" && anchor.table_block_id &&
    [anchor.first_row, anchor.first_column, anchor.last_row, anchor.last_column].every(
      (value) => Number.isInteger(value),
    )
  ) {
    target.table = {
      blockId: anchor.table_block_id,
      firstRow: anchor.first_row!,
      firstColumn: anchor.first_column!,
      lastRow: anchor.last_row!,
      lastColumn: anchor.last_column!,
    };
  }
  return target;
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
