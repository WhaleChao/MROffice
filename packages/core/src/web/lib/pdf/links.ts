/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF link-annotation overlays.
 *
 * pdf.js exposes a page's annotations via `page.getAnnotations()`. Link
 * annotations (`subtype === "Link"`) carry a `rect` (the clickable region in PDF
 * user space) plus EITHER an internal `dest` (a named-destination string or an
 * explicit destination array, resolved to a page like the outline does) OR an
 * external `url`. This module turns the loosely-typed raw annotations into a
 * typed, sanitized list the viewer renders as clickable regions over each page,
 * and provides the pure geometry to place those regions.
 *
 * Internal links reuse the outline's `resolveOutlineDestination` resolver
 * (injected pdf.js methods) so there is ONE definition of "dest -> page index".
 * External URLs are guarded by `sanitizeLinkUrl` (http/https only) from
 * `doc-search`, so a hostile PDF can't smuggle a `javascript:` / `data:` href.
 *
 * Everything here is pure (no DOM, no pdf.js runtime) so it can be unit-tested
 * without jsdom or a live document. The viewer converts the PDF-space rect into
 * a viewport rect with pdf.js (rotation-aware) and then calls the pure
 * `viewportRectToPercent` to position the overlay.
 */

import { sanitizeLinkUrl } from "./doc-search";
import {
  nativePdfAnnotationIsInvisible,
  readNativePdfAnnotationFlags,
} from "./native-text-annotations";
import {
  resolveOutlineDestination,
  type OutlineDestResolver,
} from "./outline";

/** A raw link annotation as returned by pdf.js `getAnnotations()`. Loosely typed. */
export interface RawLinkAnnotation {
  /** Annotation type; only `"Link"` is treated as a navigable link. */
  subtype?: unknown;
  /** Clickable region in PDF user space: `[x1, y1, x2, y2]`. */
  rect?: unknown;
  /**
   * Internal destination: a named-destination string or an explicit dest array
   * (`dest[0]` is the page reference). Absent for external/URL links.
   */
  dest?: string | unknown[] | null;
  /** External URL (present on URL links). pdf.js pre-parses this. */
  url?: unknown;
  /** pdf.js id for the annotation (used as a stable React key when present). */
  id?: unknown;
  /** Widget annotations: pdf.js marks an AcroForm push button. */
  pushButton?: unknown;
  /** Widget annotations: the field's /TU, its user-facing name. */
  alternativeText?: unknown;
  /** Widget annotations: the field's other (/AA, JavaScript) actions. */
  actions?: unknown;
  /** Widget annotations: a Named action pdf.js parsed from /A. */
  action?: unknown;
  /** The annotation's /F flags (Hidden, NoView ...). */
  annotationFlags?: unknown;
  /** pdf.js QuadPoints: per quadrilateral its axis-aligned box as
   * `[minX,maxY, maxX,maxY, minX,minY, maxX,minY]`, only when every quad
   * lies inside /Rect (pdf.js drops the whole array otherwise). */
  quadPoints?: unknown;
  /** The annotation's /Contents, which for a link is its description. */
  contentsObj?: unknown;
}

/**
 * The standard Named actions that only turn pages (ISO 32000-1 12.6.4.11,
 * table 212). They need no script, write or outside access, so a reader
 * performs them; any other Named action stays inert.
 */
export const PDF_NAMED_PAGE_ACTIONS = [
  "NextPage",
  "PrevPage",
  "FirstPage",
  "LastPage",
] as const;
export type PdfNamedPageAction = (typeof PDF_NAMED_PAGE_ACTIONS)[number];

export function isPdfNamedPageAction(
  value: unknown,
): value is PdfNamedPageAction {
  return (
    typeof value === "string" &&
    (PDF_NAMED_PAGE_ACTIONS as readonly string[]).includes(value)
  );
}

/**
 * The 1-based visible page a Named page action goes to from `fromOrdinal`,
 * or null when it has nowhere to go (NextPage on the last page).
 */
export function namedPageActionTarget(
  action: PdfNamedPageAction,
  fromOrdinal: number,
  pageCount: number,
): number | null {
  if (!Number.isInteger(pageCount) || pageCount < 1) return null;
  const from = Math.min(Math.max(1, Math.trunc(fromOrdinal) || 1), pageCount);
  const target =
    action === "NextPage"
      ? from + 1
      : action === "PrevPage"
        ? from - 1
        : action === "FirstPage"
          ? 1
          : pageCount;
  return target >= 1 && target <= pageCount && target !== from ? target : null;
}

/** A PDF-space rectangle: `[x1, y1, x2, y2]` (origin bottom-left). */
export type PdfRect = [number, number, number, number];

/** A normalized, sanitized link annotation ready for rendering. */
export interface PdfLinkAnnotation {
  /** Stable id (pdf.js id when present, else a positional fallback). */
  id: string;
  /** Clickable region in PDF user space, normalized to `[x1,y1,x2,y2]`. */
  rect: PdfRect;
  /**
   * Raw internal destination carried over verbatim for later resolution, or
   * `null` when this is an external link (or has no destination).
   */
  dest: string | unknown[] | null;
  /**
   * Sanitized external href (http/https only), or `null` when this is an
   * internal link or the URL was rejected.
   */
  url: string | null;
  /** The link's own description (a link's /Contents, a form push button's
   *  /TU), shown instead of the generic "internal link" label. */
  label?: string;
  /** The link's QuadPoints regions in PDF user space, when it declares more
   *  than its /Rect: only these areas are the clickable link (ISO 32000-1
   *  §12.5.6.5, table 173), the rest of /Rect is not part of it. */
  quads?: PdfRect[];
  /** A standard page-turning Named action, when the link has no dest/url. */
  namedAction?: PdfNamedPageAction;
}

const MAX_LINK_LABEL = 512;








function isSafeGoToPushButton(annotation: RawLinkAnnotation): boolean {
  if (annotation.subtype !== "Widget" || annotation.pushButton !== true) {
    return false;
  }
  if (!hasDest(annotation.dest)) return false;
  if (annotation.url != null || annotation.action != null) return false;
  const actions = annotation.actions;
  if (
    actions != null &&
    (typeof actions !== "object" || Object.keys(actions).length > 0)
  ) {
    return false;
  }
  return true;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/**
 * Normalize a raw `rect` into an ordered `[x1,y1,x2,y2]` (x1<=x2, y1<=y2) tuple,
 * or `null` when it is not 4 finite numbers or has zero area. pdf.js rects are
 * usually already ordered, but malformed PDFs can reverse them.
 */
export function normalizeRect(rect: unknown): PdfRect | null {
  if (!Array.isArray(rect) || rect.length < 4) return null;
  const [a, b, c, d] = rect;
  if (!isFiniteNumber(a) || !isFiniteNumber(b) || !isFiniteNumber(c) || !isFiniteNumber(d)) {
    return null;
  }
  const x1 = Math.min(a, c);
  const y1 = Math.min(b, d);
  const x2 = Math.max(a, c);
  const y2 = Math.max(b, d);
  if (x2 <= x1 || y2 <= y1) return null;
  return [x1, y1, x2, y2];
}

/**
 * True when the annotation's /F says it must not be shown or used: a Hidden
 * annotation is neither displayed nor interactive whatever its type (ISO
 * 32000-1 §12.5.3, table 165), and NoView is not displayed on screen.
 */
export function pdfAnnotationFlagsHideOnScreen(flags: unknown): boolean {
  return nativePdfAnnotationIsInvisible(readNativePdfAnnotationFlags(flags));
}

/**
 * The QuadPoints regions of a link, as PDF-space boxes, or undefined when it
 * has none worth keeping: none declared, malformed, or a single region that
 * is the whole /Rect anyway. pdf.js already checked every quad lies in /Rect.
 */
export function normalizeLinkQuadPoints(
  quadPoints: unknown,
  rect: PdfRect,
): PdfRect[] | undefined {
  const values =
    quadPoints instanceof Float32Array || quadPoints instanceof Float64Array
      ? Array.from(quadPoints)
      : Array.isArray(quadPoints)
        ? quadPoints
        : null;
  if (!values || values.length === 0 || values.length % 8 !== 0) {
    return undefined;
  }
  if (!values.every(isFiniteNumber)) return undefined;
  const epsilon = 0.01;
  const quads: PdfRect[] = [];
  for (let i = 0; i < values.length; i += 8) {
    const xs = [values[i], values[i + 2], values[i + 4], values[i + 6]];
    const ys = [values[i + 1], values[i + 3], values[i + 5], values[i + 7]];
    const box: PdfRect = [
      Math.max(rect[0], Math.min(...xs)),
      Math.max(rect[1], Math.min(...ys)),
      Math.min(rect[2], Math.max(...xs)),
      Math.min(rect[3], Math.max(...ys)),
    ];
    // A quad outside /Rect makes the whole array unusable (ISO: the quads
    // are ignored then and /Rect is the link) - pdf.js applies the same rule.
    if (
      Math.min(...xs) < rect[0] - epsilon ||
      Math.max(...xs) > rect[2] + epsilon ||
      Math.min(...ys) < rect[1] - epsilon ||
      Math.max(...ys) > rect[3] + epsilon
    ) {
      return undefined;
    }
    if (box[2] - box[0] <= 0 || box[3] - box[1] <= 0) continue;
    quads.push(box);
  }
  if (quads.length === 0) return undefined;
  if (
    quads.length === 1 &&
    quads[0].every((value, index) => Math.abs(value - rect[index]) <= epsilon)
  ) {
    return undefined;
  }
  return quads;
}

function cleanLinkLabel(value: unknown): string {
  return typeof value === "string"
    ? value
        .replace(/[\u0000-\u001f\u007f]+/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, MAX_LINK_LABEL)
    : "";
}

/** A link's /Contents description, as pdf.js exposes it. */
export function linkContentsLabel(contentsObj: unknown): string {
  if (!contentsObj || typeof contentsObj !== "object") return "";
  return cleanLinkLabel((contentsObj as { str?: unknown }).str);
}

/**
 * Whether a raw `dest` is a navigable internal destination (non-empty string or
 * non-empty array). Mirrors the outline's notion of navigability.
 */
function hasDest(dest: unknown): dest is string | unknown[] {
  if (typeof dest === "string") return dest.length > 0;
  if (Array.isArray(dest)) return dest.length > 0;
  return false;
}

/**
 * Normalize a page's raw link annotations into a typed, sanitized list.
 *
 *   - keeps only `subtype === "Link"` entries with a valid rect,
 *   - keeps only entries that resolve to SOMETHING actionable: a navigable
 *     internal `dest` OR a sanitized http/https `url` (links with neither — e.g.
 *     a `javascript:` url that got rejected, or a header with no dest — are
 *     dropped so no inert/unsafe region is rendered),
 *   - carries the raw `dest` verbatim for later page resolution,
 *   - assigns a stable id (pdf.js id when present, else a positional fallback
 *     `link-<pageNumber>-<index>`). The page number is folded into the fallback
 *     id so the positional ids stay globally unique across pages — two pages'
 *     first links no longer both collapse to `link-0` (which would clash as
 *     React keys / when the viewer flattens per-page link state).
 *
 * @param rawAnnotations The array from `page.getAnnotations()` (may be
 *   null/undefined/garbage).
 * @param pageNumber The 1-based page this annotation list belongs to. Used to
 *   namespace the positional fallback id so ids are unique across the document.
 *   Non-finite values fall back to `0` so the id is still a stable string.
 */
export function normalizeLinkAnnotations(
  rawAnnotations: unknown,
  pageNumber: number = 0,
): PdfLinkAnnotation[] {
  if (!Array.isArray(rawAnnotations)) return [];
  const pageKey = isFiniteNumber(pageNumber) ? pageNumber : 0;
  const out: PdfLinkAnnotation[] = [];
  rawAnnotations.forEach((raw, index) => {
    if (!raw || typeof raw !== "object") return;
    const annotation = raw as RawLinkAnnotation;
    const goToButton = isSafeGoToPushButton(annotation);
    if (annotation.subtype !== "Link" && !goToButton) return;
    // A Hidden (or NoView) link is not drawn and takes no click or key
    
    if (pdfAnnotationFlagsHideOnScreen(annotation.annotationFlags)) return;
    const rect = normalizeRect(annotation.rect);
    if (!rect) return;

    const dest = hasDest(annotation.dest) ? annotation.dest : null;
    const url = dest == null ? sanitizeLinkUrl(annotation.url) : null;
    // A standard "next page" style link turns the page; it used to be dropped
    
    const namedAction =
      dest == null && url == null && isPdfNamedPageAction(annotation.action)
        ? annotation.action
        : undefined;
    // Drop links that are neither navigable nor safely external.
    if (dest == null && url == null && !namedAction) return;

    const rawId = annotation.id;
    const id =
      typeof rawId === "string" && rawId.length > 0
        ? `link-${rawId}`
        : `link-${pageKey}-${index}`;
    
    // a link's /Contents.
    const label = goToButton
      ? cleanLinkLabel(annotation.alternativeText)
      : linkContentsLabel(annotation.contentsObj);
    const quads = goToButton
      ? undefined
      : normalizeLinkQuadPoints(annotation.quadPoints, rect);
    out.push({
      id,
      rect,
      dest,
      url,
      ...(label ? { label } : {}),
      ...(quads ? { quads } : {}),
      ...(namedAction ? { namedAction } : {}),
    });
  });
  return out;
}

/**
 * Convert a VIEWPORT-space rectangle (already rotation-applied by pdf.js via
 * `viewport.convertToViewportRectangle`) into page-wrapper percentages
 * `{x, y, width, height}` in the SAME coordinate convention the highlight / ink
 * / find overlays use (top-left origin, percent of the rendered page box).
 *
 * `vrect` is `[vx1, vy1, vx2, vy2]` in viewport pixels (pdf.js may return it
 * with y reversed); `viewportWidth`/`viewportHeight` are the rendered page box
 * dimensions. The result is clamped to [0, 100] so a slightly out-of-bounds
 * annotation rect can't paint an overlay outside the page. Returns `null` when
 * the inputs are degenerate (non-finite, zero-size viewport, or zero-area rect).
 */
export function viewportRectToPercent(
  vrect: ReadonlyArray<number>,
  viewportWidth: number,
  viewportHeight: number,
): { x: number; y: number; width: number; height: number } | null {
  if (!Array.isArray(vrect) || vrect.length < 4) return null;
  const [a, b, c, d] = vrect;
  if (!isFiniteNumber(a) || !isFiniteNumber(b) || !isFiniteNumber(c) || !isFiniteNumber(d)) {
    return null;
  }
  if (!isFiniteNumber(viewportWidth) || !isFiniteNumber(viewportHeight)) return null;
  if (viewportWidth <= 0 || viewportHeight <= 0) return null;

  const left = Math.min(a, c);
  const right = Math.max(a, c);
  const top = Math.min(b, d);
  const bottom = Math.max(b, d);
  if (right <= left || bottom <= top) return null;

  const clamp = (v: number) => Math.max(0, Math.min(100, v));
  const x = clamp((left / viewportWidth) * 100);
  const y = clamp((top / viewportHeight) * 100);
  const xEnd = clamp((right / viewportWidth) * 100);
  const yEnd = clamp((bottom / viewportHeight) * 100);
  const width = xEnd - x;
  const height = yEnd - y;
  if (width <= 0 || height <= 0) return null;
  return { x, y, width, height };
}

/**
 * Resolve a link annotation's internal `dest` to a 1-based page NUMBER (what the
 * viewer's `goToPage` expects), or `null` when it has no dest or can't resolve.
 * Reuses the outline's injected resolver so internal links and bookmarks share
 * one resolution path.
 */
export async function resolveLinkPageNumber(
  link: Pick<PdfLinkAnnotation, "dest">,
  resolver: OutlineDestResolver,
): Promise<number | null> {
  if (link.dest == null) return null;
  const index = await resolveOutlineDestination(link.dest, resolver);
  return index == null ? null : index + 1;
}

export { sanitizeLinkUrl };

// --------------------------------------------------------------------------
// USER-CREATED link overlays (Acrobat "Add / Edit Link" parity).
//
// Distinct from the read-only annotations above (those come from the PDF's own
// `getAnnotations()`): these are links the USER draws over a page (select a rect
// → attach a URL or a goto-page target). They persist as a `pdf_link_overlay`
// annotation (same datasets endpoint as shapes / ink), render as a clickable
// region, and bake into the flatten export as a real PDF Link annotation (a
// rect + a URI or GoTo action). Geometry is stored in page-relative PERCENT
// [0..100] (top-left origin) like every other overlay, so it is resolution- and
// rotation-independent; the flatten planner maps it back to user space.
// --------------------------------------------------------------------------

/** Max URL length accepted for a user-created link (defensive bound). */
export const PDF_LINK_OVERLAY_MAX_URL = 2_048;
/** Max user-created link overlays kept per page. */
export const PDF_LINK_OVERLAY_MAX_PER_PAGE = 200;

/**
 * A user-created link overlay. Exactly ONE of `url` / `goToPage` is set:
 *   - `url`: a sanitized http/https external link, or
 *   - `goToPage`: a 1-based page to jump to within this document.
 */
export interface PdfLinkOverlay {
  id: string;
  /** 1-based page the link sits on. */
  page: number;
  /** Clickable region in page-relative percent [0..100], top-left origin. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Sanitized external href (http/https), or null for an internal link. */
  url: string | null;
  /** 1-based goto-page destination, or null for an external link. */
  goToPage: number | null;
}

function isPageNumber(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 10_000
  );
}

/**
 * Build the normalized payload to POST for a new user link, or `null` when the
 * input is unusable (degenerate rect, off-page, no valid destination). Exactly
 * one destination is kept: a sanitized `url` takes precedence; otherwise a valid
 * `goToPage`. The rect is clamped on-page and rejected if it has zero area.
 */
export function buildLinkOverlayPayload(input: {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  url?: string | null;
  goToPage?: number | null;
}): Omit<PdfLinkOverlay, "id"> | null {
  if (!isPageNumber(input.page)) return null;
  const x = Number(input.x);
  const y = Number(input.y);
  const width = Number(input.width);
  const height = Number(input.height);
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(y) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }
  // Clamp the rect fully on-page (top-left origin, percent units).
  const cx = Math.max(0, Math.min(100, x));
  const cy = Math.max(0, Math.min(100, y));
  const cw = Math.min(width, 100 - cx);
  const ch = Math.min(height, 100 - cy);
  if (cw <= 0 || ch <= 0) return null;

  const url =
    typeof input.url === "string" && input.url.length <= PDF_LINK_OVERLAY_MAX_URL
      ? sanitizeLinkUrl(input.url)
      : null;
  const goToPage = url == null && isPageNumber(input.goToPage) ? input.goToPage : null;
  // A link with no actionable destination is meaningless — drop it.
  if (url == null && goToPage == null) return null;

  return { page: input.page, x: cx, y: cy, width: cw, height: ch, url, goToPage };
}

/**
 * Sanitize a persisted/untrusted link-overlay payload into a typed
 * {@link PdfLinkOverlay}, or `null` when it cannot be made safe. Reuses
 * {@link buildLinkOverlayPayload} so persisted bytes are validated identically
 * to freshly-built ones (sanitized URL scheme, on-page rect, single dest).
 */
export function sanitizeLinkOverlayPayload(
  id: string,
  payload: Record<string, unknown>,
): PdfLinkOverlay | null {
  if (typeof id !== "string" || id.length === 0) return null;
  const built = buildLinkOverlayPayload({
    page: Number(payload.page),
    x: Number(payload.x),
    y: Number(payload.y),
    width: Number(payload.width),
    height: Number(payload.height),
    url: typeof payload.url === "string" ? payload.url : null,
    goToPage:
      typeof payload.goToPage === "number" ? payload.goToPage : null,
  });
  if (!built) return null;
  return { id, ...built };
}

/**
 * Cap the number of link overlays kept for a single page to
 * {@link PDF_LINK_OVERLAY_MAX_PER_PAGE}, dropping the OLDEST first. Returns a new
 * array; input untouched. Mirrors `capShapesPerPage`.
 */
export function capLinkOverlaysPerPage(
  overlays: ReadonlyArray<PdfLinkOverlay>,
): PdfLinkOverlay[] {
  if (overlays.length <= PDF_LINK_OVERLAY_MAX_PER_PAGE) return [...overlays];
  return overlays.slice(overlays.length - PDF_LINK_OVERLAY_MAX_PER_PAGE);
}
