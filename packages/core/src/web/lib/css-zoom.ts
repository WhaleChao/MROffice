/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The CSS `zoom` an element is drawn at: its own and its ancestors'.
 *
 * On a touch device whole regions of the site take `zoom`
 * (styles/touch-sizing.css). Inside one, getBoundingClientRect() and
 * pointer coordinates are in screen pixels while the element's own lengths
 * (style.width/left, offset/client/scroll sizes, scrollTop) are in its CSS
 * pixels, a quarter smaller. Script that writes a measured distance back
 * into a style or a scroll position divides it by this first. Chromium and
 * WebKit agree on both sides of that line.
 */
export function cssZoomOf(element: Element | null | undefined): number {
  if (!element) return 1;
  const native = (element as Element & { currentCSSZoom?: number })
    .currentCSSZoom;
  if (typeof native === "number" && native > 0) return native;
  if (typeof getComputedStyle !== "function") return 1;
  // An engine without currentCSSZoom: each ancestor's zoom read once a frame.
  // A style read between two font faces added at an open restyled the whole
  
  const now = clockMs();
  const frame = renderedFrame();
  let zoom = 1;
  for (let node: Element | null = element; node; node = node.parentElement) {
    const known = nodeZooms.get(node);
    let own: number;
    if (known && known.epoch === transformEpoch && sameFrame(known, now, frame)) own = known.own;
    else {
      own = Number.parseFloat(getComputedStyle(node).zoom);
      nodeZooms.set(node, { at: now, epoch: transformEpoch, frame, own });
    }
    if (own > 0) zoom *= own;
  }
  return zoom;
}

const nodeZooms = new WeakMap<Element, StyleAnswer>();

/** What one style read of a node answered, and when. */
type StyleAnswer = { at: number; epoch: number; frame: number | undefined; own: number };

function clockMs(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

/** Where no frame is counted: how long one frame's style reads stand. */
const FRAME_MS = 16;










let framesDrawn = 0;
let frameAsked = false;
function renderedFrame(): number | undefined {
  if (typeof requestAnimationFrame !== "function") return undefined;
  if (!frameAsked) {
    frameAsked = true;
    requestAnimationFrame(() => {
      frameAsked = false;
      framesDrawn += 1;
    });
  }
  return framesDrawn;
}

function sameFrame(known: StyleAnswer, now: number, frame: number | undefined): boolean {
  return frame !== undefined && known.frame !== undefined ? known.frame === frame : now - known.at < FRAME_MS;
}

/**
 * The scale Touch mode draws the site's regions at: 1.25 with the reader's
 * Touch mode switch on, else 1 — in the workspace the switch alone decides,
 * on any device; a public page (no `data-workspace-app`) follows the
 * device's primary pointer (styles/touch-sizing.css, --touch-scale, keyed on
 * the same conditions). Read from the page's state, never from computed
 * style, so a per-frame caller forces no style.
 */
export function touchScale(): number {
  if (typeof document === "undefined") return 1;
  const root = document.documentElement;
  if (root.hasAttribute("data-touch-mode")) return TOUCH_SCALE;
  if (root.hasAttribute("data-workspace-app")) return 1;
  try {
    return globalThis.matchMedia?.("(pointer: coarse)").matches ? TOUCH_SCALE : 1;
  } catch {
    return 1;
  }
}

export const TOUCH_SCALE = 1.25;

let zoomScalesText: boolean | undefined;

/**
 * Whether `zoom` reaches font sizes here. iPadOS Safari scales an element's
 * boxes and spacing but draws text at its stated size whatever the zoom, and
 * no style or `text-size-adjust` value changes that (measured on iOS 16 and 18): a
 * sheet at 50% kept its letters and clipped them. Asked of the engine once.
 */
export function cssZoomScalesText(): boolean {
  if (zoomScalesText !== undefined) return zoomScalesText;
  if (typeof document === "undefined" || !document.body) return true;
  const probe = document.createElement("div");
  probe.style.cssText = "position:absolute;visibility:hidden;zoom:0.5;font-size:20px";
  // It needs text of its own: the engine leaves an empty element's size to
  // `zoom` and answers 20px, whatever it does to every element with text.
  probe.textContent = "0";
  document.body.appendChild(probe);
  // Such an engine reports the unzoomed equivalent of what it draws: 40px.
  zoomScalesText = !(Number.parseFloat(getComputedStyle(probe).fontSize) > 24);
  probe.remove();
  return zoomScalesText;
}

/** A font size for text under the sheet's zoom. Where `zoom` leaves text
 * alone it carries the zoom itself, through the stage's `--sheet-text-zoom`.
 * Measuring keeps the stated size: only the drawn one changes. */
export function sheetZoomedFontSize<T>(size: T): T | string {
  return typeof size !== "number" || cssZoomScalesText() ? size : `calc(${size}px * var(--sheet-text-zoom, 1))`;
}

/** The style with its font size under the sheet's zoom. */
export function withSheetZoomedFont<T extends { fontSize?: unknown }>(style: T): T {
  const fontSize = sheetZoomedFontSize(style.fontSize);
  return fontSize === style.fontSize ? style : { ...style, fontSize };
}

/**
 * Screen pixels per unit of an element's own layout answers (offsetWidth,
 * scrollLeft): what a distance measured on screen (getBoundingClientRect,
 * pointer coordinates) is divided by before it is written to scrollLeft.
 * Read from the element itself, so it holds whatever the engine answers
 * under `zoom` (CSS pixels in some, zoomed pixels in others) and under a
 * transform.
 */
export function screenScaleOf(element: HTMLElement): number {
  const own = element.offsetWidth;
  const drawn = element.getBoundingClientRect().width;
  return own > 0 && drawn > 0 ? drawn / own : 1;
}

/** The scale the CSS transforms of `element` and its ancestors draw it at: the
 * notebook desk's zoom, a window's opening scale. */
export function transformScaleOf(element: Element | null | undefined): number {
  let scale = 1;
  if (typeof getComputedStyle !== "function" || typeof DOMMatrixReadOnly === "undefined" || !element) return scale;
  // The grid's canvas asks this at every paint. Each ancestor's own scale is
  // kept: one drawn at its own size for a second, one scaling it (a window's
  // opening, a live zoom's preview) for a frame. The walk read the
  // style of some 25 ancestors (11% of a Page Down), and at a workbook's open,
  // between the font faces added one at a time, every read restyled the whole
  
  // ancestor is seen within a second, at once after a desk zoom or a resize.
  const now = clockMs();
  const frame = renderedFrame();
  for (let node: Element | null = element; node; node = node.parentElement) {
    const known = nodeScales.get(node);
    // A node drawn at its own size (no transform, or a translation only: a
    // window's place on the desk, moved by every drag) keeps its answer for a
    // second; only one scaling what it draws is read again each frame. The
    // window's translation was read again every frame, and every read of an
    
    if (known && known.epoch === transformEpoch && (known.own !== 1 ? sameFrame(known, now, frame) : now - known.at < TRANSFORM_WALK_MS)) {
      scale *= known.own;
      continue;
    }
    const style = getComputedStyle(node);
    const transform = style.transform;
    const own = style.scale;
    const transformed = !((!transform || transform === "none") && (!own || own === "none"));
    const value = transformed ? ownTransformScale(node, transform || "none", own || "none") : 1;
    nodeScales.set(node, { at: now, epoch: transformEpoch, frame, own: value });
    scale *= value;
  }
  return scale;
}










function ownTransformScale(node: Element, transform?: string, own?: string): number {
  if (transform === undefined || own === undefined) {
    const style = getComputedStyle(node);
    transform ??= style.transform;
    own ??= style.scale;
  }
  let scale = 1;
  if (transform && transform !== "none") {
    const matrix = new DOMMatrixReadOnly(transform);
    const fromMatrix = Math.hypot(matrix.a, matrix.b);
    if (fromMatrix > 0) scale *= fromMatrix;
  }
  if (own && own !== "none") {
    // "0.8", "0.8 0.8" or "80%": the horizontal factor, as the matrix's.
    const first = own.trim().split(/\s+/)[0] ?? "";
    const value = first.endsWith("%") ? Number.parseFloat(first) / 100 : Number.parseFloat(first);
    if (Number.isFinite(value) && value > 0) scale *= value;
  }
  return scale;
}

/** How long an ancestor with no transform stands as one. */
const TRANSFORM_WALK_MS = 1000;
const nodeScales = new WeakMap<Element, StyleAnswer>();
let transformEpoch = 0;
if (typeof window !== "undefined" && typeof window.addEventListener === "function") {
  const walkAgain = () => {
    transformEpoch += 1;
  };
  window.addEventListener("mrai:desk-zoom-settled", walkAgain);
  window.addEventListener("mrai:live-zoom-settled", walkAgain);
  window.addEventListener("resize", walkAgain);
}

/** A transform put on or taken off an ancestor by a gesture (a live zoom's
 * preview): the next measure walks the ancestors again instead of the ones
 * found within the last second. */
export function noteAncestorTransformChanged(): void {
  transformEpoch += 1;
}

/** Sent on `window` once a live zoom's preview is taken off its surface (the
 * zoom committed, or not): what drew its own pixels under the preview draws
 * them again without it. */
export const LIVE_ZOOM_SETTLED_EVENT = "mrai:live-zoom-settled";

/** Sent on `window` once the notebook desk's zoom has settled: what draws its
 * own pixels draws them again for the new scale. */
export const DESK_ZOOM_SETTLED_EVENT = "mrai:desk-zoom-settled";

/** The canvas pixels one of an element's own pixels takes on the screen: the
 * device's, the sheet's `zoom` and the transforms above it, so a canvas under a
 * zoomed desk is drawn as sharp as the text beside it. Bounded: a whole sheet's
 * tile at 300% and more would pass what a canvas may hold. */
export function canvasPixelsPerUnit(element: Element, zoom: number): number {
  const device = (typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1) * (zoom > 0 ? zoom : 1);
  return device * Math.min(2, Math.max(0.5, transformScaleOf(element)));
}

/**
 * What to divide a client rect under `element` by for its own lengths, on a
 * sheet drawn at `zoom`. A rect under `zoom` is in screen pixels in some engines
 * and already in the element's own in WebKit, which the element's measure tells.
 * A transform above it counts in both: at a desk zoomed to 104.6% the canvas
 * tiles were laid out in screen pixels inside the zoomed desk, 142 px for 8 rows
 * of 136, so a sheet's text ran ahead of its rows and its row numbers.
 */
export function rectScaleUnderZoom(element: HTMLElement, zoom: number): number {
  const sheetScale = zoom > 0 ? zoom : 1;
  const outer = transformScaleOf(element);
  // At the sheet's own size either answer below is the transforms' scale: the
  // measure (offsetWidth and a client rect, a forced layout at every canvas
  
  if (sheetScale === 1) return outer;
  const measured = screenScaleOf(element) / outer;
  return (Math.abs(measured - sheetScale) <= Math.abs(measured - 1) ? sheetScale : 1) * outer;
}
