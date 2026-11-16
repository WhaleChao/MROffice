"use client";






















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ChatSource, CitationAnchor, SourceLocator } from "@/components/console/chat/AITurnSources";
import { localizedAppHref } from "../__cut__/m72b8a7b97f";
import { locateTargetForLocator, type LocatorForLocate, type PageLocateTarget } from "./page-locate";

/** The query key a citation link carries its place in. */
export const CITE_PARAM = "cite";
/** Another part of the app asks the open notebook to open a cited place;
 * the notebook cancels the event when it did (NotebookDetailClient). */
export const OPEN_CITED_PLACE_EVENT = "mrai:open-cited-place";

export interface OpenCitedPlaceDetail {
  notebookId: string;
  pageId: string;
  target: PageLocateTarget;
}

const ID = /^[A-Za-z0-9_-]{1,128}$/;
const A1 = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$/;
const SNIPPET_CHARS = 80;
/** Every source marker, as the chips count them (MarkdownContent). */
const SOURCE_MARKER = /[[［]S(\d{1,3})(?:[:：][^\]］\n]{1,60})?[\]］]/g;

interface CompactPlace {
  b?: string;
  i?: number;
  s?: string;
  /**
   * t: a document table's cells (block, grid); s: a worksheet range (name,
   * id); d: a deck's slide (id, number), the object on it and, for a table
   * object, its cells; p: a PDF's page.
   */
  a?: {
    k: "t" | "s" | "d" | "p";
    r?: string;
    tb?: string;
    g?: [number, number, number, number];
    sn?: string;
    si?: string;
    n?: number;
    ob?: string;
  };
}

function base64UrlEncode(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function base64UrlDecode(text: string): string | null {
  try {
    const binary = atob(text.replace(/-/g, "+").replace(/_/g, "/"));
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return new TextDecoder().decode(bytes);
  } catch {
    return null;
  }
}

function gridIndex(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 && value < 1_048_576;
}

function cellRef(value: unknown): value is string {
  return typeof value === "string" && A1.test(value);
}

function ordinal(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= 100_000;
}

/** The place a citation opens, written for a link's query. */
export function citationPlaceToken(locator: LocatorForLocate, snippet?: string | null): string {
  const place: CompactPlace = {};
  if (locator.block_id) place.b = locator.block_id;
  if (locator.paragraph_index) place.i = locator.paragraph_index;
  const anchor = locator.anchor;
  const grid = anchor ? [anchor.first_row, anchor.first_column, anchor.last_row, anchor.last_column] : [];
  const cells = grid.every(gridIndex) ? (grid as [number, number, number, number]) : null;
  if (anchor?.kind === "table_range" && anchor.table_block_id && cellRef(anchor.ref) && cells) {
    place.a = { k: "t", r: anchor.ref, tb: anchor.table_block_id, g: cells };
  } else if (anchor?.kind === "sheet_range" && anchor.sheet_name && cellRef(anchor.ref)) {
    place.a = { k: "s", r: anchor.ref, sn: anchor.sheet_name.slice(0, 80), ...(anchor.sheet_id ? { si: anchor.sheet_id } : {}) };
  } else if (anchor?.kind === "slide_object" && (anchor.slide_id || ordinal(anchor.slide_number))) {
    place.a = {
      k: "d",
      ...(anchor.slide_id ? { si: anchor.slide_id } : {}),
      ...(ordinal(anchor.slide_number) ? { n: anchor.slide_number } : {}),
      ...(anchor.block_id ? { ob: anchor.block_id } : {}),
      ...(anchor.block_id && cellRef(anchor.ref) && cells ? { r: anchor.ref, g: cells } : {}),
    };
  } else if (anchor?.kind === "pdf_page" && ordinal(anchor.page_number)) {
    place.a = { k: "p", n: anchor.page_number };
  }
  // The snippet finds the passage when nothing else names it, and the line
  // on a PDF's page.
  if ((place.a?.k === "p" || (!place.b && !place.i && !place.a)) && snippet?.trim()) {
    place.s = snippet.trim().slice(0, SNIPPET_CHARS);
  }
  return base64UrlEncode(JSON.stringify(place));
}

/** A link's place, read back as the locator it was made from; null when it is none. */
export function readCitationPlaceToken(
  pageId: string,
  token: string,
): { locator: LocatorForLocate; snippet: string | null } | null {
  if (!ID.test(pageId) || token.length > 2048) return null;
  const json = base64UrlDecode(token);
  if (!json) return null;
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return null;
  }
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const place = raw as CompactPlace;
  const locator: LocatorForLocate = {
    page_id: pageId,
    block_id: typeof place.b === "string" && ID.test(place.b) ? place.b : null,
    paragraph_index: typeof place.i === "number" && Number.isInteger(place.i) && place.i > 0 ? place.i : null,
  };
  const anchor = place.a && typeof place.a === "object" ? place.a : null;
  const cells = anchor && Array.isArray(anchor.g) && anchor.g.length === 4 && anchor.g.every(gridIndex) ? anchor.g : null;
  const range = cells
    ? { first_row: cells[0], first_column: cells[1], last_row: cells[2], last_column: cells[3] }
    : null;
  if (anchor?.k === "t" && cellRef(anchor.r) && typeof anchor.tb === "string" && ID.test(anchor.tb) && range) {
    locator.anchor = { kind: "table_range", ref: anchor.r, table_block_id: anchor.tb, ...range };
    locator.block_id = locator.block_id ?? anchor.tb;
  } else if (anchor?.k === "s" && cellRef(anchor.r) && typeof anchor.sn === "string" && anchor.sn.trim()) {
    locator.anchor = {
      kind: "sheet_range", ref: anchor.r, sheet_name: anchor.sn.slice(0, 80),
      sheet_id: typeof anchor.si === "string" && anchor.si.length <= 128 ? anchor.si : null,
    };
  } else if (anchor?.k === "d") {
    const slideId = typeof anchor.si === "string" && ID.test(anchor.si) ? anchor.si : null;
    const number = ordinal(anchor.n) ? anchor.n : undefined;
    if (slideId || number) {
      const objectId = typeof anchor.ob === "string" && ID.test(anchor.ob) ? anchor.ob : null;
      locator.anchor = {
        kind: "slide_object",
        slide_id: slideId,
        ...(number ? { slide_number: number } : {}),
        block_id: objectId,
        ...(objectId && cellRef(anchor.r) && range ? { ref: anchor.r, ...range } : {}),
      };
    }
  } else if (anchor?.k === "p" && ordinal(anchor.n)) {
    locator.anchor = { kind: "pdf_page", page_number: anchor.n };
  }
  const snippet = typeof place.s === "string" ? place.s.slice(0, SNIPPET_CHARS) : null;
  return { locator, snippet };
}

/** The address a citation opens, in this page's origin and locale. */
export function citationHref(options: {
  origin: string;
  notebookId: string;
  locator: LocatorForLocate;
  snippet?: string | null;
}): string | null {
  const { origin, notebookId, locator } = options;
  if (!ID.test(notebookId) || !ID.test(locator.page_id)) return null;
  const query = new URLSearchParams({
    openPage: locator.page_id,
    [CITE_PARAM]: citationPlaceToken(locator, options.snippet),
  });
  return `${origin}${localizedAppHref(`/app/notebooks/${encodeURIComponent(notebookId)}?${query}`)}`;
}

/** A citation link this app made, read back; null for any other address. */
export function parseCitationHref(href: string, origin?: string): OpenCitedPlaceDetail | null {
  let url: URL;
  try {
    url = new URL(href, origin ?? (typeof window !== "undefined" ? window.location.origin : "http://localhost"));
  } catch {
    return null;
  }
  const here = origin ?? (typeof window !== "undefined" ? window.location.origin : null);
  if (here && url.origin !== here) return null;
  const path = /^(?:\/[a-z]{2}(?:-[A-Za-z0-9]{2,4})?)?\/(?:app|workspace)\/notebooks\/([^/?#]+)\/?$/.exec(url.pathname);
  const pageId = url.searchParams.get("openPage") ?? "";
  const token = url.searchParams.get(CITE_PARAM);
  if (!path || !token) return null;
  const notebookId = decodeURIComponent(path[1]);
  if (!ID.test(notebookId)) return null;
  const place = readCitationPlaceToken(pageId, token);
  if (!place) return null;
  return { notebookId, pageId, target: locateTargetForLocator(place.locator, place.snippet) };
}

/**
 * Opens a citation link the app made: in the open notebook's canvas when it
 * is the cited page's notebook (the notebook takes the event), else by
 * going there; the place is parked for the page's editor either way.
 */
export function openCitationHref(href: string): boolean {
  const detail = parseCitationHref(href);
  if (!detail) return false;
  const event = new CustomEvent<OpenCitedPlaceDetail>(OPEN_CITED_PLACE_EVENT, { detail, cancelable: true });
  if (!window.dispatchEvent(event)) return true;
  window.location.assign(href);
  return true;
}

// --- an answer as a report keeps it ----------------------------------------------

export interface CitationFootnote {
  number: number;
  label: string;
  href: string | null;
}

export interface ReportCitations {
  sources: readonly ChatSource[] | null | undefined;
  citationAnchors?: readonly CitationAnchor[] | null;
  /** The notebook of a cited page, when its source does not say (the host's). */
  notebookId?: string | null;
  origin: string;
  /** "关键数据 · B2", as the chip's card names a place. */
  placeLabel: (locator: SourceLocator) => string | null;
}

/** Where a marker points: its own place (citation_anchors), else its source's. */
export function markerLocator(source: ChatSource, occurrence: number, marker: number, anchors: ReportCitations["citationAnchors"]): SourceLocator | null {
  const own = anchors?.find((entry) => entry.occurrence === occurrence && entry.source === marker);
  if (own && (!source.locator || own.locator.page_id === source.locator.page_id || own.locator.page_id === source.target?.page_id)) {
    return own.locator;
  }
  return source.locator ?? null;
}

function placeKey(locator: SourceLocator | null): string {
  if (!locator) return "";
  const anchor = locator.anchor;
  if (!anchor) return `${locator.page_id}|${locator.block_id ?? ""}#${locator.paragraph_index ?? ""}`;
  const holder = anchor.table_block_id ?? anchor.sheet_id ?? anchor.sheet_name ?? anchor.slide_id ?? anchor.slide_number ?? "";
  return `${locator.page_id}|${anchor.kind}:${holder}:${anchor.block_id ?? ""}!${anchor.ref ?? anchor.page_number ?? ""}`;
}

/** Numbers each cited place once, in the order the answer first cites it. */
export class CitationNumbering {
  private readonly byKey = new Map<string, CitationFootnote>();
  readonly footnotes: CitationFootnote[] = [];

  constructor(private readonly options: ReportCitations) {}

  /** The footnote of source `marker` at `locator`; null when the marker names no source. */
  footnote(marker: number, locator: SourceLocator | null): CitationFootnote | null {
    const source = this.options.sources?.[marker - 1];
    if (!source) return null;
    const key = `${marker}|${placeKey(locator)}`;
    const known = this.byKey.get(key);
    if (known) return known;
    const place = locator ? this.options.placeLabel(locator) : null;
    const label = [source.display_name, place].filter(Boolean).join(" · ");
    const notebookId = source.target?.notebook_id ?? this.options.notebookId ?? null;
    const href = source.url
      ? source.url
      : locator && notebookId && !source.unavailable
        ? citationHref({ origin: this.options.origin, notebookId, locator, snippet: source.snippet ?? null })
        : null;
    const footnote = { number: this.footnotes.length + 1, label, href };
    this.byKey.set(key, footnote);
    this.footnotes.push(footnote);
    return footnote;
  }

  /** Each source marker of `text`, from occurrence `first`, as its footnote's link. */
  replaceMarkers(text: string, first: number, link: (footnote: CitationFootnote) => string): { text: string; count: number } {
    let count = 0;
    const out = text.replace(SOURCE_MARKER, (_whole, number: string) => {
      const occurrence = first + count;
      count += 1;
      const marker = Number(number);
      const source = this.options.sources?.[marker - 1];
      if (!source) return "";
      const footnote = this.footnote(marker, markerLocator(source, occurrence, marker, this.options.citationAnchors));
      return footnote ? link(footnote) : "";
    });
    return { text: out, count };
  }
}

/** How many source markers `text` holds, as the chips count them. */
export function countSourceMarkers(text: string): number {
  return [...text.matchAll(SOURCE_MARKER)].length;
}
