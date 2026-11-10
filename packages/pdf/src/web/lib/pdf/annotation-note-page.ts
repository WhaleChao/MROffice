/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure content builder for "annotations → note page" (the Zotero
 * "add note from annotations" pattern): folds a PDF page's highlights and
 * side-panel annotations into a brand-new NATIVE DOCUMENT page's
 * `content_json` (`{ type: "doc", content: [...] }` — the exact shape
 * `lib/file-templates/compile.ts` builds and `documentPageContentSchema`
 * validates), grouped under one heading per PDF page.
 *
 * Every node/mark used here is registered by the document plugin registry
 * (see components/console/editor/native-document/pluginRegistry.ts):
 * heading / paragraph / blockquote / horizontalRule / text, the `link` mark
 * (href limited server-side to http(s)/mailto/… OR scheme-less relative URLs
 * — we emit a same-origin relative URL), the `highlight` mark (TextStyling —
 * strict #rrggbb only, hence the alpha-compositing converter below) and the
 * `reference` block node (ReferenceBlock — the app-native "open this page"
 * card used for the final backlink).
 *
 * Server-side caps this builder must respect (apps/api routers/notebooks.py):
 * MAX_DOCUMENT_NODES 10_000 / text-node chars 65_536 / URL length 4096. The
 * excerpt cap below keeps the worst case well under the node budget.
 *
 * PURE: no React / DOM / network — unit-testable in isolation. All
 * user-facing strings arrive via the `labels` bag (the host resolves i18n).
 */

/** A single source annotation to excerpt (already loaded by the host). */
export interface PdfExcerptAnnotation {
  /** "highlight" = viewer text highlight; "note" = side-panel annotation. */
  kind: "highlight" | "note";
  /** 1-based PDF page number; null/undefined lands in the unpaged group. */
  page?: number | null;
  /** The quoted PDF text (highlight text or a side annotation's excerpt). */
  quote?: string | null;
  /** The user's own note/comment text. */
  note?: string | null;
  /** The highlight's CSS color (palette rgba or free hex); optional. */
  color?: string | null;
}

/** Host-resolved (i18n) strings the builder stamps into the document. */
export interface PdfExcerptLabels {
  /** Full title for the generated page ("<PDF title> — 批注摘录"). */
  pageTitle: string;
  /** Heading for one PDF page's group ("第 {page} 页"). */
  pageHeading: (page: number) => string;
  /** Heading for annotations that carry no page number. */
  unpagedHeading: string;
  /** Per-annotation source-link text with a page ("打开 PDF · 第 {page} 页"). */
  openAtPage: (page: number) => string;
  /** Per-annotation source-link text without a page ("打开 PDF"). */
  openPdf: string;
  /** Intro line naming the source PDF ("来源：{title}"). */
  sourceIntro: string;
  /** Title on the final backlink reference card (the PDF page's title). */
  backlinkTitle: string;
  /** Snippet on the final backlink reference card ("返回 PDF 原文"). */
  backlinkSnippet: string;
}

export interface PdfExcerptNotePageInput {
  notebookId: string;
  /** The PDF notebook page's id (the backlink / open-URL target). */
  pdfPageId: string;
  annotations: readonly PdfExcerptAnnotation[];
  labels: PdfExcerptLabels;
}

/** ProseMirror-style JSON node (same loose shape as lib/file-templates/compile.ts). */
export interface ExcerptDocumentNode {
  type: string;
  attrs?: Record<string, unknown>;
  text?: string;
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>;
  content?: ExcerptDocumentNode[];
}

export interface PdfExcerptNotePage {
  title: string;
  content: { type: "doc"; content: ExcerptDocumentNode[] };
  /** How many annotations made it into the page (post skip-empty). */
  excerptCount: number;
}

/**
 * Hard cap on excerpted annotations. Each excerpt emits ≤ 8 nodes, so 1000
 * excerpts stay far below the server's 10_000-node document cap even with the
 * per-page headings and framing nodes on top.
 */
export const PDF_EXCERPT_MAX_ANNOTATIONS = 1000;

/**
 * Defensive per-run text cap. PDF highlight text is server-capped at 4000 and
 * notes at 2000, but side-panel annotations (e.g. saved AI summaries) are only
 * bounded by the pdf-page schema; keep every run under the server's 65_536
 * text-node limit with margin.
 */
export const PDF_EXCERPT_MAX_TEXT = 20_000;

/** The server's MAX_DOCUMENT_NODES (10_000), less a margin. */
const PDF_EXCERPT_NODE_BUDGET = 9_500;

/**
 * Convert a PDF highlight CSS color (strict hex `#rgb`/`#rrggbb`, or
 * `rgb()`/`rgba()` — the only shapes the viewer's own sanitizer lets through)
 * into the OPAQUE lowercase `#rrggbb` the document editor's `highlight` mark
 * accepts, compositing any alpha over white so the mark's solid background
 * matches what the translucent overlay looks like on the white PDF page.
 * Returns null for anything unparseable (the caller then omits the mark).
 */
export function pdfHighlightColorToHex(color: unknown): string | null {
  if (typeof color !== "string") return null;
  const value = color.trim().toLowerCase();
  if (!value) return null;

  let r: number;
  let g: number;
  let b: number;
  let a = 1;

  const hex = value.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/);
  const rgb = value.match(
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*(0|1|0?\.\d+|1\.0)\s*)?\)$/,
  );
  if (hex) {
    const raw = hex[1];
    const six =
      raw.length === 3
        ? raw
            .split("")
            .map((ch) => ch + ch)
            .join("")
        : raw;
    r = parseInt(six.slice(0, 2), 16);
    g = parseInt(six.slice(2, 4), 16);
    b = parseInt(six.slice(4, 6), 16);
  } else if (rgb) {
    r = Number(rgb[1]);
    g = Number(rgb[2]);
    b = Number(rgb[3]);
    if (rgb[4] !== undefined) a = Number(rgb[4]);
  } else {
    return null;
  }

  if ([r, g, b].some((ch) => !Number.isInteger(ch) || ch < 0 || ch > 255)) {
    return null;
  }
  if (!Number.isFinite(a) || a < 0 || a > 1) return null;

  // Composite over white: c' = c*a + 255*(1-a).
  const blend = (ch: number) => Math.round(ch * a + 255 * (1 - a));
  const toHex = (ch: number) => blend(ch).toString(16).padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}













export function pdfPageOpenUrl(
  notebookId: string,
  pdfPageId: string,
  pdfPage?: number | null,
): string {
  const base = `/app/notebooks/${encodeURIComponent(notebookId)}?openPage=${encodeURIComponent(pdfPageId)}`;
  return typeof pdfPage === "number" && Number.isSafeInteger(pdfPage) && pdfPage >= 1
    ? `${base}&pdfPage=${pdfPage}`
    : base;
}

// ---------------------------------------------------------------------------
// Node authoring helpers (mirror lib/file-templates/compile.ts)
// ---------------------------------------------------------------------------

function textNode(
  value: string,
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>,
): ExcerptDocumentNode {
  return marks && marks.length > 0
    ? { type: "text", text: value, marks }
    : { type: "text", text: value };
}

function paragraph(content: ExcerptDocumentNode[]): ExcerptDocumentNode {
  return { type: "paragraph", content };
}

function heading(level: 1 | 2 | 3, value: string): ExcerptDocumentNode {
  return { type: "heading", attrs: { level }, content: [textNode(value)] };
}

/** Collapse whitespace runs, trim, and cap the run length. */
function cleanText(value: unknown): string {
  if (typeof value !== "string") return "";
  const flat = value.replace(/\s+/g, " ").trim();
  return flat.length > PDF_EXCERPT_MAX_TEXT
    ? flat.slice(0, PDF_EXCERPT_MAX_TEXT)
    : flat;
}







function cleanNoteLines(value: unknown): string[] {
  if (typeof value !== "string") return [];
  const lines: string[] = [];
  let budget = PDF_EXCERPT_MAX_TEXT;
  for (const raw of value.split(/\r\n|\r|\n|\u2028|\u2029/)) {
    if (budget <= 0) break;
    const line = raw.replace(/\s+/g, " ").trim().slice(0, budget);
    if (!line) continue;
    lines.push(line);
    budget -= line.length;
  }
  return lines;
}

interface NormalizedExcerpt {
  page: number | null;
  quote: string;
  note: string[];
  colorHex: string | null;
}

/** Sanitize one source annotation; null = skip (empty of any text). */
function normalizeExcerpt(
  annotation: PdfExcerptAnnotation,
): NormalizedExcerpt | null {
  if (!annotation || typeof annotation !== "object") return null;
  const quote = cleanText(annotation.quote);
  const note = cleanNoteLines(annotation.note);
  if (!quote && note.length === 0) return null;
  const rawPage = annotation.page;
  const page =
    typeof rawPage === "number" && Number.isInteger(rawPage) && rawPage >= 1
      ? rawPage
      : null;
  return {
    page,
    quote,
    note,
    colorHex:
      annotation.kind === "highlight"
        ? pdfHighlightColorToHex(annotation.color)
        : null,
  };
}

/**
 * Build the note page (title + document content) from a PDF's annotations.
 *
 * Structure:
 *   paragraph  — source intro line
 *   per page group (ascending; unpaged group last):
 *     heading(2) — page heading
 *     per annotation:
 *       blockquote — quote paragraph (highlight-colored when the source is a
 *                    colored highlight) + note paragraph (when present)
 *       paragraph  — source link back to the PDF (link mark, relative URL)
 *   horizontalRule
 *   reference block — app-native backlink card to the PDF page itself
 *
 * Annotations with no text at all are skipped. Returns null when NOTHING
 * survives (the host keeps its button disabled / errors before calling the
 * create API).
 */
export function buildPdfExcerptNotePage(
  input: PdfExcerptNotePageInput,
): PdfExcerptNotePage | null {
  const { notebookId, pdfPageId, labels } = input;
  const normalized = (input.annotations ?? [])
    .map(normalizeExcerpt)
    .filter((item): item is NormalizedExcerpt => item !== null)
    .slice(0, PDF_EXCERPT_MAX_ANNOTATIONS);
  if (normalized.length === 0) return null;

  // Group by page, preserving input order within a group. Numbered pages
  // ascending first, the unpaged group (if any) last.
  const byPage = new Map<number | null, NormalizedExcerpt[]>();
  for (const excerpt of normalized) {
    const group = byPage.get(excerpt.page);
    if (group) {
      group.push(excerpt);
    } else {
      byPage.set(excerpt.page, [excerpt]);
    }
  }
  const pages = [...byPage.keys()]
    .filter((page): page is number => page !== null)
    .sort((a, b) => a - b);
  const groups: Array<number | null> = byPage.has(null)
    ? [...pages, null]
    : pages;

  const content: ExcerptDocumentNode[] = [
    paragraph([textNode(labels.sourceIntro)]),
  ];

  // Every excerpt costs at most 8 nodes before its note's extra lines (see
  // PDF_EXCERPT_MAX_ANNOTATIONS); the lines beyond a note's first share what
  // is left of the server's node budget, and past it fold into the last line.
  let spareNodes =
    PDF_EXCERPT_NODE_BUDGET - 8 - groups.length * 2 - normalized.length * 8;

  for (const page of groups) {
    content.push(
      heading(2, page === null ? labels.unpagedHeading : labels.pageHeading(page)),
    );
    // Each group's links name its own page, so following one lands there.
    const linkMark = {
      type: "link",
      attrs: { href: pdfPageOpenUrl(notebookId, pdfPageId, page) },
    };
    for (const excerpt of byPage.get(page) ?? []) {
      const quoteChildren: ExcerptDocumentNode[] = [];
      if (excerpt.quote) {
        quoteChildren.push(
          paragraph([
            textNode(
              excerpt.quote,
              excerpt.colorHex
                ? [{ type: "highlight", attrs: { color: excerpt.colorHex } }]
                : undefined,
            ),
          ]),
        );
      }
      if (excerpt.note.length > 0) {
        const extra = Math.max(
          0,
          Math.min(excerpt.note.length - 1, Math.floor(spareNodes / 2)),
        );
        spareNodes -= extra * 2;
        const lines = excerpt.note.slice(0, extra);
        lines.push(excerpt.note.slice(extra).join(" "));
        for (const line of lines) {
          quoteChildren.push(paragraph([textNode(line)]));
        }
      }
      content.push({ type: "blockquote", content: quoteChildren });
      content.push(
        paragraph([
          textNode(
            page === null ? labels.openPdf : labels.openAtPage(page),
            [linkMark],
          ),
        ]),
      );
    }
  }

  content.push({ type: "horizontalRule" });
  content.push({
    type: "reference",
    attrs: {
      target_type: "page",
      target_id: pdfPageId,
      title: labels.backlinkTitle,
      snippet: labels.backlinkSnippet,
    },
  });

  return {
    title: labels.pageTitle,
    content: { type: "doc", content },
    excerptCount: normalized.length,
  };
}
