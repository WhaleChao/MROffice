/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for INTRA-PAGE destination offsets.
 *
 * PDF outline bookmarks and internal link annotations carry explicit
 * destinations that name not just a PAGE but a position ON the page:
 * `/XYZ left top zoom` and `/FitH top` (and its bounding-box variant `/FitBH`)
 * both specify a `top` coordinate in PDF user space (origin bottom-left).
 * The viewer historically discarded that coordinate and landed every jump at
 * the top of the page; these helpers extract the top coordinate and convert it
 * to a pixel offset inside the rendered page wrapper so a TOC click lands on
 * the actual section heading. Mirrors pdf.js's `PDFViewer#scrollPageIntoView`
 * destination handling (pdf.js, Apache-2.0) in a pure, unit-testable form.
 *
 * Everything here is DOM-free; the component supplies the live pdf.js resolver
 * methods (`getDestination` / `getPageIndex`) exactly like `outline.ts` does.
 */

import { type OutlineDestResolver } from "@mrnote/core/web/lib/pdf/outline.ts";

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/**
 * The view an explicit destination asks for (ISO 32000-1 12.3.2.2, Table
 * 151). Coordinates are PDF user space; `null` is the standard's "unchanged"
 * (for a jump to another page: its left edge / its top, as pdf.js reads it).
 * The bounding-box forms (FitB, FitBH, FitBV) are read as their page forms.
 */
export type PdfDestView =
  | { kind: "XYZ"; left: number | null; top: number | null; zoom: number | null }
  | { kind: "Fit" }
  | { kind: "FitH"; top: number | null }
  | { kind: "FitV"; left: number | null }
  | { kind: "FitR"; left: number; bottom: number; right: number; top: number };

/** A resolved destination: the 1-based target page plus an optional
 * PDF user-space `top` coordinate (null when the dest doesn't carry one). */
export interface PdfDestLocation {
  /** 1-based page number (what the viewer's `goToPage` expects). */
  pageNumber: number;
  /** PDF user-space top coordinate (origin bottom-left), or null. */
  top: number | null;
  /** The whole view the destination names, or null when unreadable. */
  view?: PdfDestView | null;
}

function destKind(explicitDest: unknown[]): string | null {
  const rawKind = explicitDest[1] as { name?: unknown } | string | null;
  return typeof rawKind === "string"
    ? rawKind
    : rawKind && typeof rawKind === "object" && typeof rawKind.name === "string"
      ? rawKind.name
      : null;
}

function finiteOrNull(value: unknown): number | null {
  return isFiniteNumber(value) ? value : null;
}

/** Read the view of an explicit destination array, or null. */
export function extractDestView(explicitDest: unknown): PdfDestView | null {
  if (!Array.isArray(explicitDest) || explicitDest.length < 2) return null;
  switch (destKind(explicitDest)) {
    case "XYZ": {
      const zoom = finiteOrNull(explicitDest[4]);
      return {
        kind: "XYZ",
        left: finiteOrNull(explicitDest[2]),
        top: finiteOrNull(explicitDest[3]),
        // 0 and null both mean "keep the current zoom".
        zoom: zoom !== null && zoom > 0 ? zoom : null,
      };
    }
    case "Fit":
    case "FitB":
      return { kind: "Fit" };
    case "FitH":
    case "FitBH":
      return { kind: "FitH", top: finiteOrNull(explicitDest[2]) };
    case "FitV":
    case "FitBV":
      return { kind: "FitV", left: finiteOrNull(explicitDest[2]) };
    case "FitR": {
      const values = explicitDest.slice(2, 6).map(finiteOrNull);
      if (values.length !== 4 || values.some((value) => value === null)) {
        return null;
      }
      const [x1, y1, x2, y2] = values as number[];
      if (x1 === x2 || y1 === y2) return null;
      return {
        kind: "FitR",
        left: Math.min(x1, x2),
        bottom: Math.min(y1, y2),
        right: Math.max(x1, x2),
        top: Math.max(y1, y2),
      };
    }
    default:
      return null;
  }
}

/** Where a destination view lands, for the viewer to apply. */
export interface PdfDestViewPlan {
  /** The render scale to switch to, or null to keep the current one. */
  scale: number | null;
  /** The sticky fit the view implies (Fit = page, FitH = width), if any. */
  fitMode: "page" | "width" | null;
  /** CSS px from the page wrapper's left edge to put at the viewport's left. */
  offsetX: number;
  /** CSS px from the page wrapper's top edge to put at the viewport's top. */
  offsetY: number;
}

/**
 * Plan the zoom and scroll position of a destination view on a page, the
 * way pdf.js's `scrollPageIntoView` does: the named point (or the rectangle's
 * upper-left corner) goes to the viewport's upper-left, at the zoom the view
 * names or implies. `pageWidth`/`pageHeight` are the page's unrotated size at
 * scale 1; `userUnit` scales user-space coordinates into it; `rotation` is
 * the page's displayed rotation. `viewportWidth` is the width one page may
 * use (spread halves it).
 */
export function planDestView(input: {
  view: PdfDestView;
  pageWidth: number;
  pageHeight: number;
  rotation: number;
  userUnit?: number;
  



  originX?: number;
  originY?: number;
  currentScale: number;
  viewportWidth: number;
  viewportHeight: number;
  minScale: number;
  maxScale: number;
}): PdfDestViewPlan | null {
  const { view, pageWidth: width, pageHeight: height } = input;
  if (
    !isFiniteNumber(width) ||
    !isFiniteNumber(height) ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }
  const unit =
    isFiniteNumber(input.userUnit) && input.userUnit > 0 ? input.userUnit : 1;
  const originX = isFiniteNumber(input.originX) ? input.originX : 0;
  const originY = isFiniteNumber(input.originY) ? input.originY : 0;
  const rotation = (((Math.round(input.rotation / 90) * 90) % 360) + 360) % 360;
  const swapped = rotation === 90 || rotation === 270;
  const shownWidth = swapped ? height : width;
  const shownHeight = swapped ? width : height;
  const clamp = (value: number) =>
    Math.min(input.maxScale, Math.max(input.minScale, +value.toFixed(2)));
  const vw = isFiniteNumber(input.viewportWidth) ? input.viewportWidth : 0;
  const vh = isFiniteNumber(input.viewportHeight) ? input.viewportHeight : 0;

  // Page point (user space) -> unrotated scale-1 page units, y down from top.
  let x = 0;
  let y = 0;
  let scale: number | null = null;
  let fitMode: PdfDestViewPlan["fitMode"] = null;
  switch (view.kind) {
    case "XYZ":
      x = view.left === null ? 0 : (view.left - originX) * unit;
      y = view.top === null ? 0 : height - (view.top - originY) * unit;
      scale = view.zoom !== null ? clamp(view.zoom) : null;
      break;
    case "Fit":
      scale = vw > 0 && vh > 0 ? clamp(Math.min(vw / shownWidth, vh / shownHeight)) : null;
      fitMode = "page";
      break;
    case "FitH":
      y = view.top === null ? 0 : height - (view.top - originY) * unit;
      scale = vw > 0 ? clamp(vw / shownWidth) : null;
      fitMode = "width";
      break;
    case "FitV":
      x = view.left === null ? 0 : (view.left - originX) * unit;
      scale = vh > 0 ? clamp(vh / shownHeight) : null;
      break;
    case "FitR": {
      const rectWidth = (view.right - view.left) * unit;
      const rectHeight = (view.top - view.bottom) * unit;
      const shownRectWidth = swapped ? rectHeight : rectWidth;
      const shownRectHeight = swapped ? rectWidth : rectHeight;
      scale =
        vw > 0 && vh > 0
          ? clamp(Math.min(vw / shownRectWidth, vh / shownRectHeight))
          : null;
      x = (view.left - originX) * unit;
      y = height - (view.top - originY) * unit;
      break;
    }
  }
  const s = scale ?? (isFiniteNumber(input.currentScale) && input.currentScale > 0 ? input.currentScale : 1);
  // Unrotated (x from left, y from top) -> the displayed page, per rotation.
  let shownX = x;
  let shownY = y;
  if (view.kind === "FitR") {
    // The rectangle's displayed upper-left corner.
    const x2 = (view.right - originX) * unit;
    const y2 = height - (view.bottom - originY) * unit;
    const corners = [
      [x, y],
      [x2, y2],
    ].map(([cx, cy]) => rotatePagePoint(cx, cy, width, height, rotation));
    shownX = Math.min(corners[0][0], corners[1][0]);
    shownY = Math.min(corners[0][1], corners[1][1]);
  } else {
    [shownX, shownY] = rotatePagePoint(x, y, width, height, rotation);
    // A coordinate the view leaves open stays at the page edge it starts on.
    if (view.kind === "XYZ" || view.kind === "FitH" || view.kind === "FitV") {
      const leftOpen = view.kind === "FitH" || (view.kind === "XYZ" && view.left === null);
      const topOpen =
        view.kind === "FitV" || ((view.kind === "XYZ" || view.kind === "FitH") && view.top === null);
      if (!swapped) {
        if (leftOpen) shownX = 0;
        if (topOpen) shownY = 0;
      } else {
        if (leftOpen) shownY = 0;
        if (topOpen) shownX = 0;
      }
    }
  }
  return {
    scale,
    fitMode,
    offsetX: Math.max(0, Math.min(shownWidth, shownX)) * s,
    offsetY: Math.max(0, Math.min(shownHeight, shownY)) * s,
  };
}

/**
 * Map an unrotated page point (x from the left, y from the top, scale-1
 * units) to the page as displayed at `rotation` degrees clockwise.
 */
function rotatePagePoint(
  x: number,
  y: number,
  width: number,
  height: number,
  rotation: number,
): [number, number] {
  switch (rotation) {
    case 90:
      return [height - y, x];
    case 180:
      return [width - x, height - y];
    case 270:
      return [y, width - x];
    default:
      return [x, y];
  }
}

/**
 * Extract the PDF user-space `top` coordinate from an EXPLICIT destination
 * array, or `null` when the destination kind carries none.
 *
 * pdf.js explicit destinations look like `[pageRef, {name: "XYZ"}, left, top,
 * zoom]`; the second element names the fit kind. Handled kinds:
 *   - `XYZ`   → `dest[3]` (top; may be null = "keep current position"),
 *   - `FitH` / `FitBH` → `dest[2]` (top).
 * Everything else (`Fit`, `FitV`, `FitR`, garbage) yields null — the caller
 * falls back to the page top, the historical behavior. A plain-string kind is
 * accepted defensively alongside pdf.js's `{name}` object shape.
 */
export function extractDestTop(explicitDest: unknown): number | null {
  if (!Array.isArray(explicitDest) || explicitDest.length < 2) return null;
  const rawKind = explicitDest[1] as { name?: unknown } | string | null;
  const kind =
    typeof rawKind === "string"
      ? rawKind
      : rawKind && typeof rawKind === "object" && typeof rawKind.name === "string"
        ? rawKind.name
        : null;
  if (kind === "XYZ") {
    const top = explicitDest[3];
    return isFiniteNumber(top) ? top : null;
  }
  if (kind === "FitH" || kind === "FitBH") {
    const top = explicitDest[2];
    return isFiniteNumber(top) ? top : null;
  }
  return null;
}

/**
 * Resolve a raw outline/link `dest` to a {@link PdfDestLocation} using the
 * injected resolver, or `null` when it cannot be resolved. Same resolution
 * shape as `resolveOutlineDestination` (named dest → `getDestination`, then
 * `dest[0]` → `getPageIndex`), extended to ALSO surface the top coordinate so
 * outline clicks and embedded links share one location-aware path. Resolver
 * errors are swallowed into `null` so one broken bookmark never throws.
 */
export async function resolveDestLocation(
  dest: string | unknown[] | null | undefined,
  resolver: OutlineDestResolver,
): Promise<PdfDestLocation | null> {
  try {
    let explicit: unknown[] | null = null;
    if (typeof dest === "string") {
      if (dest.length === 0) return null;
      explicit = await resolver.getDestination(dest);
    } else if (Array.isArray(dest)) {
      explicit = dest;
    }
    if (!Array.isArray(explicit) || explicit.length === 0) return null;
    const ref = explicit[0];
    if (ref == null) return null;
    const index = await resolver.getPageIndex(ref);
    if (typeof index !== "number" || !Number.isInteger(index) || index < 0) {
      return null;
    }
    return {
      pageNumber: index + 1,
      top: extractDestTop(explicit),
      view: extractDestView(explicit),
    };
  } catch {
    return null;
  }
}

/**
 * Convert a PDF user-space `top` coordinate into a CSS pixel offset from the
 * TOP of the rendered page wrapper: `(pageHeight − top) × scale` (PDF y grows
 * upward; CSS y grows downward). Clamped to `[0, pageHeight × scale]` so a
 * malformed coordinate can never scroll outside the page. A null/non-finite
 * `top` or degenerate page height yields 0 (land at the page top — the
 * historical behavior). `scale` defaults to 1 when degenerate.
 */
export function destTopToPageOffset(
  top: number | null | undefined,
  pageHeight: number,
  scale: number,
): number {
  if (!isFiniteNumber(top)) return 0;
  if (!isFiniteNumber(pageHeight) || pageHeight <= 0) return 0;
  const s = isFiniteNumber(scale) && scale > 0 ? scale : 1;
  const offset = (pageHeight - top) * s;
  return Math.max(0, Math.min(pageHeight * s, offset));
}

/**
 * The container `scrollTop` at which a page wrapper's top edge sits, from
 * live rect measurements: `containerScrollTop + (wrapperRectTop −
 * containerRectTop) / displayScale`. Pure so the intra-page scroll math is testable without
 * layout; degenerate inputs yield 0.
 */
export function pageWrapperScrollLeft(
  containerScrollLeft: number,
  containerRectLeft: number,
  wrapperRectLeft: number,
  displayScale = 1,
): number {
  return pageWrapperScrollTop(
    containerScrollLeft,
    containerRectLeft,
    wrapperRectLeft,
    displayScale,
  );
}

export function pageWrapperScrollTop(
  containerScrollTop: number,
  containerRectTop: number,
  wrapperRectTop: number,
  displayScale = 1,
): number {
  if (
    !isFiniteNumber(containerScrollTop) ||
    !isFiniteNumber(containerRectTop) ||
    !isFiniteNumber(wrapperRectTop) ||
    !isFiniteNumber(displayScale) || displayScale <= 0
  ) {
    return 0;
  }
  return Math.max(0, containerScrollTop + (wrapperRectTop - containerRectTop) / displayScale);
}
