/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF *redaction* annotations (Acrobat "Redact" parity).
 *
 * A redaction annotation is a single box on a page (page-relative *percentage*
 * coordinates [0..100], like image overlays / form-field rects / signature
 * boxes) that marks a region whose content should be permanently removed. In
 * the viewer it renders as an OPAQUE BLACK box covering the marked region; on
 * flatten-export each redaction is BAKED as a permanent opaque black rectangle
 * drawn onto the page (so the exported PDF visually removes the content) and the
 * exported redacted pages are rasterized into image-only pages after the visual
 * bake (so copy-paste / text search of the export cannot recover the covered
 * text). See PDFViewer's export path for the rasterization step.
 *
 * IMPORTANT — honest scope of the redaction guarantee:
 *   - The on-page overlay is a removable annotation (it can be deleted before
 *     export); it is a MARK, not the removal itself.
 *   - The EXPORTED redaction is permanent in two senses: (a) an opaque black
 *     rectangle is baked into the page content stream over the region (it is NOT
 *     a removable PDF annotation), and (b) every page carrying a redaction is
 *     rasterized into an image-only page during export, removing the source page's
 *     searchable text layer from that exported page.
 *   - This is strong practical client-side redaction for text extraction and
 *     copy/search recovery, at the cost of rasterizing the affected pages. It is
 *     still not a substitute for a server-side certified redaction pipeline that
 *     parses and rewrites all PDF content streams, image OCR layers, attachments,
 *     incremental-update history, and hidden metadata. Do NOT represent this as
 *     certified / irreversible redaction.
 *
 * Everything here has no DOM / React dependency so it can be unit-tested in
 * isolation and reused by the viewer (and a potential backend mirror). It
 * mirrors `signatures.ts` (a box-shaped annotation):
 *
 *   - the persisted annotation shape + its sanitizer ({@link sanitizeRedactionPayload}),
 *   - position/size clamping (keep the box on-page, bound its dimensions),
 *   - `buildRedactionPayload` to produce exactly what we POST (persisted bytes
 *     match what we render),
 *   - `capRedactionsPerPage` (mirrors `capSignaturesPerPage` / `capStrokesPerPage`).
 *
 * The persisted annotation type is `pdf_redactions` and is additive to the
 * existing `pdf_highlight` / `pdf_note` / `pdf_ink` / `pdf_signatures`
 * annotations and the `pdf_form_values` map — old PDFs with none of these keep
 * loading unchanged, and redaction is entirely optional.
 */

/**
 * The fill used to render + bake a redaction. Redaction is conventionally a
 * solid opaque black box, so this is a single fixed color (NOT a user-chosen
 * allow-list) — there is no per-redaction color to smuggle. Exposed as RGB
 * [0,1] for the pdf-lib bake and as a CSS string for the overlay so both halves
 * reference one source of truth.
 */
export const REDACTION_FILL_RGB: Readonly<{ r: number; g: number; b: number }> = {
  r: 0,
  g: 0,
  b: 0,
};
/** CSS color for the on-page opaque redaction overlay (solid black). */
export const REDACTION_FILL_CSS = "#000000";

/**
 * Max number of redactions kept per page. Mirrors {@link
 * SIGNATURE_MAX_PER_PAGE} in spirit (a generous-but-bounded per-page cap); a
 * page rarely needs hundreds of redaction boxes but a heavily-redacted document
 * can have many.
 */
export const REDACTION_MAX_PER_PAGE = 500;

/** Min / max box size in page-percentage units (resolution-independent). */
export const REDACTION_MIN_SIZE = 0.2;
export const REDACTION_MAX_SIZE = 100;
/** Default box size for a freshly placed redaction (page-percentage units). */
export const DEFAULT_REDACTION_WIDTH = 20;
export const DEFAULT_REDACTION_HEIGHT = 4;

export interface RedactionAnnotation {
  id: string;
  /** 1-based page number. */
  page: number;
  /** Box top-left + size in page-relative percent [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
}

function isFinitePercent(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 100
  );
}

/** Clamp a single coordinate value into the on-page percentage range. */
function clampPercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  if (value < 0) return 0;
  if (value > 100) return 100;
  return value;
}

/** Clamp a box dimension into [{@link REDACTION_MIN_SIZE}, {@link REDACTION_MAX_SIZE}]. */
function clampSize(value: number, fallback: number): number {
  if (!Number.isFinite(value)) return fallback;
  if (value < REDACTION_MIN_SIZE) return REDACTION_MIN_SIZE;
  if (value > REDACTION_MAX_SIZE) return REDACTION_MAX_SIZE;
  return value;
}

/**
 * True when a RAW size value (pre-clamp) is a degenerate redaction dimension:
 * non-finite or below {@link REDACTION_MIN_SIZE}. Used by the sanitizer / build
 * helpers to REJECT a zero-area box rather than {@link clampRedactionBox}'s
 * clamp-UP behavior (which keeps an on-page box valid for rendering but must not
 * resurrect a degenerate one). A non-finite value is degenerate for the
 * sanitizer (corrupt payload), but the build helper treats `undefined` as
 * "use default" before calling this (so an omitted size still works).
 */
function isDegenerateRawSize(value: number): boolean {
  return !Number.isFinite(value) || value < REDACTION_MIN_SIZE;
}

/**
 * Clamp a redaction box so it stays fully on the page: the top-left is clamped
 * into [0,100], the size into the size bounds, then the size is shrunk if the
 * box would overflow the right / bottom edge. Mirrors `clampSignatureBox`.
 *
 * NOTE: like `clampSignatureBox`, a too-small size is clamped UP to
 * {@link REDACTION_MIN_SIZE} (so a valid-but-tiny box stays renderable). Callers
 * that must REJECT a degenerate (zero-area) box check the RAW size first via
 * {@link isDegenerateRawSize} — see {@link sanitizeRedactionPayload} /
 * {@link buildRedactionPayload}.
 */
export function clampRedactionBox(box: {
  x: number;
  y: number;
  width: number;
  height: number;
}): { x: number; y: number; width: number; height: number } {
  const x = clampPercent(box.x);
  const y = clampPercent(box.y);
  let width = clampSize(box.width, DEFAULT_REDACTION_WIDTH);
  let height = clampSize(box.height, DEFAULT_REDACTION_HEIGHT);
  // Shrink so the box never spills past the right / bottom edge.
  if (x + width > 100) width = Math.max(REDACTION_MIN_SIZE, 100 - x);
  if (y + height > 100) height = Math.max(REDACTION_MIN_SIZE, 100 - y);
  return { x, y, width, height };
}

/**
 * Normalize a raw drag-rectangle into a clean top-left + size box. A drag can
 * start at any corner (so width/height could be "negative"); this resolves the
 * two corners into a top-left origin box. Returns null when the resulting box
 * is smaller than {@link REDACTION_MIN_SIZE} in either dimension (a stray click
 * / micro-drag is not a redaction). Inputs are clamped onto the page first.
 */
export function normalizeRedactionDrag(
  start: { x: number; y: number },
  end: { x: number; y: number },
): { x: number; y: number; width: number; height: number } | null {
  if (
    !Number.isFinite(start.x) ||
    !Number.isFinite(start.y) ||
    !Number.isFinite(end.x) ||
    !Number.isFinite(end.y)
  ) {
    return null;
  }
  const sx = clampPercent(start.x);
  const sy = clampPercent(start.y);
  const ex = clampPercent(end.x);
  const ey = clampPercent(end.y);
  const x = Math.min(sx, ex);
  const y = Math.min(sy, ey);
  const width = Math.abs(ex - sx);
  const height = Math.abs(ey - sy);
  if (width < REDACTION_MIN_SIZE || height < REDACTION_MIN_SIZE) return null;
  return clampRedactionBox({ x, y, width, height });
}

/**
 * Validate + normalize a persisted `pdf_redactions` payload into a usable
 * annotation, or return null when it is structurally invalid (mirrors
 * `sanitizeSignaturePayload`). Out-of-range positions are rejected (returns
 * null) rather than silently snapped; the box size is clamped. A box that ends
 * up smaller than {@link REDACTION_MIN_SIZE} in either dimension is rejected so
 * a degenerate (zero-area) redaction is never rendered/baked.
 */
export function sanitizeRedactionPayload(
  id: string,
  payload: Record<string, unknown> | null | undefined,
): RedactionAnnotation | null {
  if (!id || typeof id !== "string") return null;
  if (!payload || typeof payload !== "object") return null;
  const record = payload as Record<string, unknown>;
  const page = Number(record.page);
  if (!Number.isInteger(page) || page < 1 || page > 10000) return null;
  const { x, y } = record;
  if (!isFinitePercent(x) || !isFinitePercent(y)) return null;
  // Reject a degenerate / corrupt box outright (rather than clamping a zero up
  // to a tiny visible box). The persisted payload must carry a real box.
  const rawWidth = Number(record.width);
  const rawHeight = Number(record.height);
  if (isDegenerateRawSize(rawWidth) || isDegenerateRawSize(rawHeight)) {
    return null;
  }
  const box = clampRedactionBox({ x, y, width: rawWidth, height: rawHeight });
  if (box.width < REDACTION_MIN_SIZE || box.height < REDACTION_MIN_SIZE) {
    return null;
  }
  return {
    id,
    page,
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
  };
}

/**
 * Build the payload to POST for a new redaction. Position/size are clamped so
 * what we persist matches what we render. Returns null when the box is smaller
 * than {@link REDACTION_MIN_SIZE} in either dimension (so a blank/degenerate
 * redaction is never persisted).
 */
export function buildRedactionPayload(input: {
  page: number;
  x: number;
  y: number;
  width?: number;
  height?: number;
}): { page: number; x: number; y: number; width: number; height: number } | null {
  // An OMITTED size falls back to the default; an EXPLICIT degenerate size is
  // rejected (so a caller can't persist a zero-area box that clamp would
  // otherwise resurrect to a tiny visible box).
  const rawWidth = input.width ?? DEFAULT_REDACTION_WIDTH;
  const rawHeight = input.height ?? DEFAULT_REDACTION_HEIGHT;
  if (isDegenerateRawSize(rawWidth) || isDegenerateRawSize(rawHeight)) {
    return null;
  }
  const box = clampRedactionBox({
    x: input.x,
    y: input.y,
    width: rawWidth,
    height: rawHeight,
  });
  if (box.width < REDACTION_MIN_SIZE || box.height < REDACTION_MIN_SIZE) {
    return null;
  }
  return {
    page: input.page,
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
  };
}

/**
 * Cap the number of redactions kept for a single page to {@link
 * REDACTION_MAX_PER_PAGE}, dropping the OLDEST first (the tail — most recent —
 * is what the user just added). Returns a new array; input untouched. Mirrors
 * `capSignaturesPerPage` / `capStrokesPerPage`.
 */
export function capRedactionsPerPage<T>(
  redactions: ReadonlyArray<T>,
  max = REDACTION_MAX_PER_PAGE,
): T[] {
  if (!Array.isArray(redactions)) return [];
  const cap = Number.isInteger(max) && max > 0 ? max : REDACTION_MAX_PER_PAGE;
  if (redactions.length <= cap) return redactions.slice();
  return redactions.slice(redactions.length - cap);
}

/**
 * True when a page-relative point (percent, top-left origin) falls inside a
 * redaction box. Used by the export's text-layer removal to decide whether a
 * positioned text item's anchor lies under a redaction (so it is omitted from
 * the exported/searchable text). The box edges are inclusive.
 */
export function pointInRedaction(
  box: { x: number; y: number; width: number; height: number },
  point: { x: number; y: number },
): boolean {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return false;
  return (
    point.x >= box.x &&
    point.x <= box.x + box.width &&
    point.y >= box.y &&
    point.y <= box.y + box.height
  );
}

/**
 * True when two axis-aligned boxes (page-relative percent, top-left origin)
 * overlap by more than a trivial epsilon. Edge-touching (shared border, zero
 * overlap area) does NOT count as covered, so a text item that merely abuts a
 * redaction edge is kept. Used by {@link textItemsUnderRedaction} to decide
 * whether a text item's glyph box is obscured by a redaction.
 */
export function boxesOverlap(
  a: { x: number; y: number; width: number; height: number },
  b: { x: number; y: number; width: number; height: number },
  epsilon = 1e-6,
): boolean {
  const ax2 = a.x + a.width;
  const ay2 = a.y + a.height;
  const bx2 = b.x + b.width;
  const by2 = b.y + b.height;
  const overlapW = Math.min(ax2, bx2) - Math.max(a.x, b.x);
  const overlapH = Math.min(ay2, by2) - Math.max(a.y, b.y);
  return overlapW > epsilon && overlapH > epsilon;
}

/**
 * A text item with a page-relative bounding box (percent, top-left origin). The
 * `index` is the item's position in the source text-item list so the caller can
 * map a decision back onto the original array (e.g. to omit it from the joined
 * text written into a searchable text layer).
 */
export interface RedactableTextItem {
  index: number;
  /** Page-relative bounding box in percent [0..100], top-left origin. */
  box: { x: number; y: number; width: number; height: number };
}

/**
 * Given a page's positioned text items + that page's redaction boxes, return
 * the SET of item indices whose glyph box is covered by ANY redaction (i.e. the
 * items whose text must be OMITTED from the exported / searchable text so a
 * copy-paste or text search of the export cannot recover the redacted text).
 *
 * Pure + DOM-free: the caller projects pdf.js text items into page-relative
 * percent boxes (e.g. via lib/pdf/text-boxes.ts geometry) and passes them here.
 * Coverage uses {@link boxesOverlap} (any non-trivial overlap), so a glyph that
 * is even partially under a redaction is dropped — conservative on the side of
 * removing more text, which is the safe default for redaction.
 */
export function textItemsUnderRedaction(
  items: ReadonlyArray<RedactableTextItem>,
  redactions: ReadonlyArray<{
    x: number;
    y: number;
    width: number;
    height: number;
  }>,
): Set<number> {
  const covered = new Set<number>();
  if (
    !Array.isArray(items) ||
    items.length === 0 ||
    !Array.isArray(redactions) ||
    redactions.length === 0
  ) {
    return covered;
  }
  for (const item of items) {
    if (!item || !item.box) continue;
    for (const red of redactions) {
      if (boxesOverlap(item.box, red)) {
        covered.add(item.index);
        break;
      }
    }
  }
  return covered;
}

/**
 * Build the redacted, copy-safe text for a page: join the text fragments of the
 * items NOT covered by any redaction (the covered items are omitted entirely so
 * their glyphs can't be recovered from the exported text). Items are joined with
 * a single space (mirroring how lib/pdf/text-search.ts `buildPageText` and the
 * viewer's find/search build page text from pdf.js items). This is the text the
 * export should write into a searchable text layer / extracted-text dataset in
 * place of the original page text.
 */
export function redactPageText(
  items: ReadonlyArray<{ index: number; str: string }>,
  coveredIndices: ReadonlySet<number>,
): string {
  if (!Array.isArray(items) || items.length === 0) return "";
  const parts: string[] = [];
  for (const item of items) {
    if (!item || typeof item.str !== "string") continue;
    if (coveredIndices.has(item.index)) continue;
    parts.push(item.str);
  }
  return parts.join(" ").replace(/\s+/g, " ").trim();
}
