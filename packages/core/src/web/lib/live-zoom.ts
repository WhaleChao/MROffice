"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Live zoom: a pinch or a Ctrl/⌘+wheel follows the fingers on every frame,
 * and the editor's real zoom is applied once, when the gesture rests.
 *
 * Why: every editor used to commit each pinch/wheel event as its zoom level.
 * A commit re-renders the whole editor (and, where the level is CSS `zoom`
 * or a render scale, re-lays out and re-rasterises the document), so on an
 * iPad the zoom advanced only when the main thread caught up — in visible
 * steps — and the Presentation stage even moved a whole 10% per wheel event.
 *
 * What the mature apps do, and what this does:
 *
 * - During the gesture the surface (the element the editor scales) gets an
 *   extra CSS `scale` + `translate` — the individual transform properties,
 *   which compose with whatever `transform` the editor already sets and are
 *   never written by React — and `will-change: transform`, so the compositor
 *   scales the existing raster: no React render, no layout, no paint.
 * - The extra transform keeps the content under the fingers (or the surface's
 *   own transform origin, for a host that centres its content) under them,
 *   and follows the fingers if they move while pinching.
 * - When the gesture rests (fingers lift, Safari's gestureend, the wheel
 *   stream pauses, a pen or mouse presses) the host commits the level once;
 *   in the same frame the extra transform is removed and the scroll offsets
 *   are corrected so the point under the fingers stays put — no jump.
 *
 * `createLiveZoomPreview` is the framework-free part (the transform and the
 * anchoring); `useLiveZoomGesture` wires it to the wheel, touches and
 * Safari's gesture events (lib/use-pinch-zoom's attachPinchGestures) and to
 * a host's committed level.
 */

import { useEffect, useLayoutEffect, useRef } from "react";
import { flushSync } from "react-dom";

import { LIVE_ZOOM_SETTLED_EVENT, noteAncestorTransformChanged } from "./css-zoom";
import { EDITOR_ZOOM_STEP, wheelDeltaPixels } from "./editor-zoom";
import {
  attachPinchGestures,
  type PinchPoint,
} from "./use-pinch-zoom";

export interface LiveZoomPoint {
  x: number;
  y: number;
}

/** How the preview keeps the content still while it scales. */
export type LiveZoomAnchor =
  /** The content under the fingers / pointer stays under them. */
  | "pointer"
  /** The surface scales about its own transform origin (a centred slide). */
  | "origin";

/** A pinch zooms once its spread has changed by more than this share. */
export const PINCH_SCALE_DEADBAND = 0.03;
/** A pinch pans once its midpoint has moved more than this (screen px). */
export const PINCH_PAN_DEADBAND_PX = 12;
/** Past the deadband, a zoom catches up with the fingers within this share. */
const PINCH_SCALE_CATCH_UP = 0.15;
/** Past the deadband, a pan catches up with the fingers within this distance (px). */
const PINCH_PAN_CATCH_UP_PX = 40;

export interface LiveZoomPreviewOptions {
  /** The scrolling viewport the surface sits in (its scroll is corrected). */
  viewport: HTMLElement;
  /**
   * The element whose transform the editor's zoom drives — or, where the
   * zoomed content is a run of sibling boxes with no wrapper (a PDF's pages),
   * those boxes: each gets its share of one shared transform.
   */
  surface: HTMLElement | readonly HTMLElement[];
  anchor: LiveZoomAnchor;
  /**
   * Of several surfaces, those near the view (default: all) are composited;
   * the rest only move, unpainted until they come into view.
   */
  composite?: (surface: HTMLElement) => boolean;
  /**
   * The surface's start is where the viewport scrolls from (a sheet's A1
   * and its headers): the preview never carries it past the viewport's
   * start, as the committed zoom cannot scroll before 0. Without it a pinch
   * out at the top-left slid the whole grid, headers and all, down and to
   * the right until the gesture rested.
   */
  clampToScrollStart?: boolean;
}

export interface LiveZoomPreview {
  /** The preview's scale relative to the committed zoom (1 = none). */
  readonly scale: number;
  /**
   * The one screen point the preview has scaled the content about, pans
   * included (null when it only scaled about the surfaces' own origins, or
   * barely scaled): a host that anchors its own zoom there lands exactly on
   * what the preview showed.
   */
  fixedPoint(): LiveZoomPoint | null;
  /** Show the content at `scale` × the committed zoom, anchored at `focal`. */
  update(scale: number, focal?: LiveZoomPoint): void;
  /** Move the content by (dx, dy) screen pixels (two fingers panning as they pinch). */
  translateBy(dx: number, dy: number): void;
  /** Whether the preview has moved the content (`translateBy`), not only scaled it. */
  readonly panned: boolean;
  /**
   * Where the content under `focal` sits from its surface's rendered corner,
   * on screen — taken before the real zoom is committed, so `finish` can put
   * the same content back under the same point.
   */
  capture(focal: LiveZoomPoint | null): LiveZoomCapture | null;
  /**
   * Remove the preview; with a capture, scroll its content back under it.
   * `ratio` is the committed level over the level the preview showed (1
   * unless the host rounded it).
   */
  finish(capture: LiveZoomCapture | null, ratio?: number): void;
}

export interface LiveZoomCapture {
  focal: LiveZoomPoint;
  /** The surface the focal point was on (or nearest to). */
  surface: HTMLElement;
  /** Screen pixels from that surface's rendered top-left corner to `focal`. */
  offsetX: number;
  offsetY: number;
}

function parseOrigin(value: string): [number, number] {
  const parts = value
    .trim()
    .split(" ")
    .filter(Boolean)
    .map((part) => Number.parseFloat(part));
  return [
    Number.isFinite(parts[0]) ? parts[0] : 0,
    Number.isFinite(parts[1]) ? parts[1] : 0,
  ];
}

/** The scale an element's own computed `transform` applies (null when none). */
function ownScale(transform: string): number | null {
  if (!transform || transform === "none") return null;
  if (transform.startsWith("matrix(")) {
    const scale = Number.parseFloat(transform.slice("matrix(".length));
    return Number.isFinite(scale) && scale > 0 ? scale : 1;
  }
  if (typeof DOMMatrixReadOnly !== "undefined") {
    try {
      const scale = new DOMMatrixReadOnly(transform).a;
      return Number.isFinite(scale) && scale > 0 ? scale : 1;
    } catch {
      return 1;
    }
  }
  return 1;
}

interface PreviewedSurface {
  element: HTMLElement;
  saved: {
    scale: string;
    translate: string;
    willChange: string;
    transformOrigin: string;
  };
  setOrigin: boolean;
  composite: boolean;
  /** The reference point minus this surface's transform origin, on screen. */
  dx: number;
  dy: number;
}

/**
 * Start a preview on `surface`. Nothing is written until the first `update`
 * with a scale other than 1; `finish` restores every inline style it touched.
 *
 * The whole preview is one mapping of the screen, X ↦ P + T + k·(X − P),
 * about a reference point P that travels with the viewport's scroll. A
 * surface scaled by `k` about its own transform origin O lands on that
 * mapping when translated by T + (1 − k)·(P − O).
 */
export function createLiveZoomPreview({
  viewport,
  surface,
  anchor,
  composite,
  clampToScrollStart = false,
}: LiveZoomPreviewOptions): LiveZoomPreview {
  const elements: readonly HTMLElement[] = Array.isArray(surface)
    ? (surface as readonly HTMLElement[])
    : [surface as HTMLElement];
  const view = viewport.ownerDocument.defaultView;
  const viewportRect = viewport.getBoundingClientRect();
  // Screen pixels per CSS pixel of the viewport's content: the notebook
  // canvas may itself be zoomed around the window.
  const ancestorScale =
    viewport.offsetWidth > 0 && viewportRect.width > 0
      ? viewportRect.width / viewport.offsetWidth
      : 1;
  const referenceX = viewportRect.left;
  const referenceY = viewportRect.top;
  const surfaces: PreviewedSurface[] = elements.map((element) => {
    const computed = view?.getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    const transformScale = ownScale(computed?.transform ?? "none");
    let originX: number;
    let originY: number;
    let setOrigin = false;
    if (transformScale === null) {
      // No transform of its own (a sheet's CSS `zoom`, a PDF page): scale
      // from the corner.
      setOrigin = true;
      originX = rect.left;
      originY = rect.top;
    } else {
      // A scale about the origin leaves the origin where it was laid out.
      const [ox, oy] = parseOrigin(computed?.transformOrigin ?? "0 0");
      originX = rect.left + ox * transformScale * ancestorScale;
      originY = rect.top + oy * transformScale * ancestorScale;
    }
    return {
      element,
      saved: {
        scale: element.style.getPropertyValue("scale"),
        translate: element.style.getPropertyValue("translate"),
        willChange: element.style.getPropertyValue("will-change"),
        transformOrigin: element.style.getPropertyValue("transform-origin"),
      },
      setOrigin,
      composite: composite ? composite(element) : true,
      dx: referenceX - originX,
      dy: referenceY - originY,
    };
  });
  // The reference point travels with the viewport's scroll (two fingers may
  // pan while they pinch); the surfaces travel with it.
  const scroll0 = { left: viewport.scrollLeft, top: viewport.scrollTop };
  const reference = (): LiveZoomPoint => ({
    x: referenceX - (viewport.scrollLeft - scroll0.left) * ancestorScale,
    y: referenceY - (viewport.scrollTop - scroll0.top) * ancestorScale,
  });

  let scale = 1;
  let tx = 0;
  let ty = 0;
  let written = false;
  let panned = false;

  // The furthest the first surface may move right / down: as far as the
  // viewport is scrolled now, the most a committed zoom can scroll back.
  const clamp = () => {
    const first = surfaces[0];
    if (!clampToScrollStart || !first) return;
    const maxX = viewport.scrollLeft * ancestorScale - (1 - scale) * first.dx;
    const maxY = viewport.scrollTop * ancestorScale - (1 - scale) * first.dy;
    if (tx > maxX) tx = maxX;
    if (ty > maxY) ty = maxY;
  };

  const write = () => {
    const first = !written;
    written = true;
    // The surface takes a `scale` of its own: what measures the transforms
    // above an element looks again (css-zoom's walk keeps what it found for a second).
    if (first) noteAncestorTransformChanged();
    for (const entry of surfaces) {
      const style = entry.element.style;
      if (first) {
        if (entry.composite) style.setProperty("will-change", "transform");
        if (entry.setOrigin) style.setProperty("transform-origin", "0px 0px");
      }
      const x = tx + (1 - scale) * entry.dx;
      const y = ty + (1 - scale) * entry.dy;
      style.setProperty("scale", String(scale));
      style.setProperty(
        "translate",
        `${x / ancestorScale}px ${y / ancestorScale}px`,
      );
    }
  };

  return {
    get scale() {
      return scale;
    },
    get panned() {
      return panned;
    },
    translateBy(dx, dy) {
      if (anchor !== "pointer" || !Number.isFinite(dx) || !Number.isFinite(dy) || (dx === 0 && dy === 0)) return;
      tx += dx;
      ty += dy;
      panned = true;
      clamp();
      write();
    },
    update(next, focal) {
      if (!Number.isFinite(next) || next <= 0) return;
      if (anchor === "pointer" && focal) {
        // The content under `focal` now stays under it at the new scale.
        const p = reference();
        const ratio = next / scale;
        tx = focal.x - p.x - ratio * (focal.x - p.x - tx);
        ty = focal.y - p.y - ratio * (focal.y - p.y - ty);
      } else if (anchor === "origin") {
        // The surface's own origin is the reference: no translation at all.
        tx = 0;
        ty = 0;
      }
      scale = next;
      clamp();
      if (!written && scale === 1 && tx === 0 && ty === 0) return;
      if (anchor === "origin") {
        for (const entry of surfaces) {
          entry.dx = 0;
          entry.dy = 0;
        }
      }
      write();
    },
    fixedPoint() {
      // P + T + k(X − P) = X  ⇔  X = P + T / (1 − k)
      if (anchor !== "pointer" || Math.abs(1 - scale) < 1e-3) return null;
      const p = reference();
      return { x: p.x + tx / (1 - scale), y: p.y + ty / (1 - scale) };
    },
    capture(focal) {
      if (anchor !== "pointer" || !focal || surfaces.length === 0) return null;
      // The surface under the fingers, or the nearest one.
      let best: { element: HTMLElement; box: DOMRect; distance: number } | null =
        null;
      for (const entry of surfaces) {
        const box = entry.element.getBoundingClientRect();
        const distance =
          Math.max(box.left - focal.x, 0, focal.x - box.right) +
          Math.max(box.top - focal.y, 0, focal.y - box.bottom);
        if (!best || distance < best.distance) {
          best = { element: entry.element, box, distance };
        }
        if (distance === 0) break;
      }
      if (!best) return null;
      return {
        focal,
        surface: best.element,
        offsetX: focal.x - best.box.left,
        offsetY: focal.y - best.box.top,
      };
    },
    finish(capture, ratio = 1) {
      if (written) {
        for (const entry of surfaces) {
          const style = entry.element.style;
          const restore = (name: string, value: string) => {
            if (value) style.setProperty(name, value);
            else style.removeProperty(name);
          };
          restore("scale", entry.saved.scale);
          restore("translate", entry.saved.translate);
          if (entry.composite) restore("will-change", entry.saved.willChange);
          if (entry.setOrigin) {
            restore("transform-origin", entry.saved.transformOrigin);
          }
        }
        written = false;
        // What was drawn under the preview (a sheet's canvas tiles, at the
        // preview's size) is drawn again without it.
        noteAncestorTransformChanged();
        if (typeof window !== "undefined") {
          window.dispatchEvent(new Event(LIVE_ZOOM_SETTLED_EVENT));
        }
      }
      scale = 1;
      tx = 0;
      ty = 0;
      panned = false;
      if (!capture || !capture.surface.isConnected) return;
      // The committed surface is drawn at the size the preview showed, from
      // the same corner (a grid whose extent follows the zoom included): the
      // content under the fingers is as far from that corner as it was.
      const box = capture.surface.getBoundingClientRect();
      const scaleRatio = Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
      const dx = box.left + capture.offsetX * scaleRatio - capture.focal.x;
      const dy = box.top + capture.offsetY * scaleRatio - capture.focal.y;
      const left0 = viewport.scrollLeft;
      const top0 = viewport.scrollTop;
      if (Math.abs(dx) >= 0.5) viewport.scrollLeft += dx / ancestorScale;
      if (Math.abs(dy) >= 0.5) viewport.scrollTop += dy / ancestorScale;
      // The scroll stops at its ends (and a page narrower than the window is
      // centred), so what the preview showed past them cannot be kept: the
      // content glides from where the fingers left it to where it can stand,
      // instead of jumping there the instant the zoom lands.
      const restX = Math.abs(dx) >= 0.5 ? dx - (viewport.scrollLeft - left0) * ancestorScale : 0;
      const restY = Math.abs(dy) >= 0.5 ? dy - (viewport.scrollTop - top0) * ancestorScale : 0;
      // restX/Y: how far right/down of the fingers the content now stands.
      glideFrom(surfaces.map((entry) => entry.element), -restX, -restY);
    },
  };
}

/** How the content settles from where a pinch left it (see `finish`). */
const LIVE_ZOOM_GLIDE_MS = 240;
const LIVE_ZOOM_GLIDE_MIN_PX = 1.5;

/** Moves `elements` from `fromX/fromY` px right/down of their place back to it. */
function glideFrom(elements: readonly HTMLElement[], fromX: number, fromY: number): void {
  if (Math.hypot(fromX, fromY) < LIVE_ZOOM_GLIDE_MIN_PX) return;
  if (typeof window === "undefined" || window.matchMedia?.("(prefers-reduced-motion: reduce)").matches) return;
  for (const element of elements) {
    if (typeof element.animate !== "function") continue;
    // Composed on top of whatever translate the element has of its own.
    element.animate(
      [{ translate: `${fromX}px ${fromY}px` }, { translate: "0px 0px" }],
      { duration: LIVE_ZOOM_GLIDE_MS, easing: "cubic-bezier(0.22, 1, 0.36, 1)", composite: "add" },
    );
  }
}

/**
 * Whether a Ctrl/⌘ key is physically held. A trackpad pinch reaches the page
 * as `wheel` events with `ctrlKey` set but no key press; a mouse wheel with
 * Ctrl held has the key down. Only the first follows the fingers 1:1.
 */
let modifierHeld = false;
let modifierListeners = 0;
const onModifierKey = (event: KeyboardEvent) => {
  modifierHeld = event.ctrlKey || event.metaKey;
};
const onModifierReset = () => {
  modifierHeld = false;
};
function watchModifierKeys(): () => void {
  if (typeof window === "undefined") return () => undefined;
  if (modifierListeners++ === 0) {
    window.addEventListener("keydown", onModifierKey, true);
    window.addEventListener("keyup", onModifierKey, true);
    window.addEventListener("blur", onModifierReset);
  }
  return () => {
    if (--modifierListeners === 0) {
      window.removeEventListener("keydown", onModifierKey, true);
      window.removeEventListener("keyup", onModifierKey, true);
      window.removeEventListener("blur", onModifierReset);
      modifierHeld = false;
    }
  };
}

/**
 * Browsers turn a trackpad pinch into Ctrl+wheel deltas of 100·ln(scale),
 * so `exp(-delta / 100)` is the fingers' own scale.
 */
const TRACKPAD_PINCH_RATE = 1 / 100;
/** A held Ctrl/⌘ and a mouse wheel: one 100 px notch is one zoom step. */
const MOUSE_WHEEL_ZOOM_RATE = Math.log(1 + EDITOR_ZOOM_STEP) / 100;

/** The relative scale one Ctrl/⌘+wheel event asks for. */
export function liveWheelZoomFactor(
  event: Pick<WheelEvent, "deltaY" | "deltaMode">,
  keyHeld: boolean,
): number {
  const pixels = wheelDeltaPixels(event.deltaY, event.deltaMode);
  if (!Number.isFinite(pixels) || pixels === 0) return 1;
  const pinch = !keyHeld && event.deltaMode === 0;
  return Math.exp(-pixels * (pinch ? TRACKPAD_PINCH_RATE : MOUSE_WHEEL_ZOOM_RATE));
}

/** How long a Ctrl/⌘+wheel stream may pause before its zoom is committed. */
export const LIVE_ZOOM_WHEEL_SETTLE_MS = 160;
/** How long an asynchronous commit may take before the preview is dropped. */
const LIVE_ZOOM_COMMIT_TIMEOUT_MS = 800;

export interface LiveZoomBounds {
  min: number;
  max: number;
}

export interface LiveZoomGestureOptions {
  /** Where the gesture is heard, and the scroll container that is anchored. */
  viewport: () => HTMLElement | null;
  /**
   * The scroll container that is anchored, when it is not the element the
   * gesture is heard on (default: `viewport`): a host whose zoom shortcut is
   * heard over a whole stage but whose content scrolls in a part of it.
   */
  scroller?: () => HTMLElement | null;
  /** The element(s) the editor's zoom scales (see createLiveZoomPreview). */
  surface: () => HTMLElement | readonly HTMLElement[] | null;
  /** createLiveZoomPreview's `composite`, for a run of surfaces. */
  composite?: (surface: HTMLElement) => boolean;
  /** The committed zoom level. */
  level: number;
  /** The levels a gesture may reach, read when it begins. */
  bounds: () => LiveZoomBounds;
  /**
   * Levels a two-finger pinch is drawn to when it passes near them (within
   * LIVE_ZOOM_SNAP_RATIO): the page exactly as wide as the window, as a
   * note-taking app's pinch settles on the page width. Read when the gesture
   * begins; a trackpad's or a wheel's zoom is not drawn.
   */
  snapLevels?: () => readonly number[];
  /**
   * Commit the level the gesture rested on. Runs inside `flushSync`, so a
   * synchronous state update is in the DOM when it returns. Return a promise
   * when the commit is asynchronous (a command authority); it resolves to
   * whether the level was applied.
   */
  onCommit: (level: number, anchor: LiveZoomPoint | null) => unknown;
  /**
   * Whether the hook scrolls the content under the fingers back under them
   * after the commit (default true). A host that anchors its own zoom at the
   * `anchor` onCommit receives says false.
   */
  correctScroll?: boolean;
  /**
   * A gesture began (inside `flushSync`): undo what its first finger began,
   * leave a view the gesture cannot zoom. The committed level is read after.
   */
  onBegin?: () => void;
  /** Whether a gesture may begin now (default: always). */
  canBegin?: () => boolean;
  anchor?: LiveZoomAnchor;
  /** See `LiveZoomPreviewOptions.clampToScrollStart`. */
  clampToScrollStart?: boolean;
  enabled?: boolean;
  /** Ctrl/⌘+wheel (default true). */
  wheel?: boolean;
  /**
   * The relative scale of one Ctrl/⌘+wheel event (default
   * liveWheelZoomFactor: a trackpad pinch follows the fingers, a mouse wheel
   * with the key held steps 10% per notch).
   */
  wheelFactor?: (event: WheelEvent, keyHeld: boolean) => number;
  /**
   * A Ctrl/⌘+wheel the host swallows without zooming (a momentum tail):
   * it is still kept from the browser's own page zoom.
   */
  ignoreWheel?: (event: WheelEvent) => boolean;
  /** Two-finger touch pinch (default true); Safari's gesture is always heard. */
  touch?: boolean;
  /** Cancel two-finger touchmoves so the content does not pan (default false). */
  preventPanWhilePinching?: boolean;
  /** attachPinchGestures' `accept`, also applied to the wheel. */
  accept?: (event: Event) => boolean;
  /** Re-attach when the viewport element is replaced. */
  attachKey?: unknown;
}

/**
 * How near a pinch must come to a snap level to be drawn to it (a ratio): 4%
 * either way, a detent the fingers feel without it fighting a deliberate
 * zoom past it.
 */
export const LIVE_ZOOM_SNAP_RATIO = 1.04;

/** `scale` drawn to the nearest of `snaps` within LIVE_ZOOM_SNAP_RATIO, else as it is. */
export function liveZoomSnapped(scale: number, snaps: readonly number[]): number {
  let best = scale;
  let nearest = Math.log(LIVE_ZOOM_SNAP_RATIO);
  for (const snap of snaps) {
    if (!(snap > 0)) continue;
    const distance = Math.abs(Math.log(scale / snap));
    if (distance <= nearest) {
      nearest = distance;
      best = snap;
    }
  }
  return best;
}

export function useLiveZoomGesture(options: LiveZoomGestureOptions): void {
  const optionsRef = useRef(options);
  const levelRef = useRef(options.level);
  /** A commit is waiting for the host's level to reach the DOM. */
  const landingRef = useRef<{
    finish: () => void;
    timer: ReturnType<typeof setTimeout> | null;
  } | null>(null);

  useLayoutEffect(() => {
    optionsRef.current = options;
  });

  // The committed level reached the DOM: drop the preview in the same frame.
  useLayoutEffect(() => {
    levelRef.current = options.level;
    const landing = landingRef.current;
    if (landing) landing.finish();
  }, [options.level]);

  const enabled = options.enabled ?? true;
  const wheel = options.wheel ?? true;
  const touch = options.touch ?? true;
  // A pinch anchored at the fingers owns the pan too: the browser's own pan
  // (on the compositor) and the preview's (on the main thread) disagreed by a
  // few pixels every frame.
  const preventPan = options.preventPanWhilePinching ?? (options.anchor ?? "pointer") === "pointer";

  useEffect(() => {
    if (!enabled) return;
    const node = optionsRef.current.viewport();
    if (!node) return;
    const releaseKeys = watchModifierKeys();

    let preview: LiveZoomPreview | null = null;
    let baseLevel = 1;
    let bounds: LiveZoomBounds = { min: 0, max: Infinity };
    /** The snap levels as preview scales (relative to baseLevel), read at begin. */
    let snaps: number[] = [];
    /** The preview scale a pinch's own scale multiplies (Safari may restart it). */
    let pinchBase = 1;
    let focal: LiveZoomPoint | null = null;
    let wheelTimer: ReturnType<typeof setTimeout> | null = null;
    let pinching = false;
    /**
     * The running pinch's intent. Two fingers never move as one: a zoom
     * drifts their midpoint a little, a pan changes their spread a little,
     * and both applied at once made the page twitch. Neither moves the page
     * until it is plainly meant (a spread past PINCH_SCALE_DEADBAND, a
     * midpoint past PINCH_PAN_DEADBAND_PX), and each then starts from where
     * it was let in, with no jump.
     */
    let pinchState: {
      start: LiveZoomPoint;
      scaleFrom: number | null;
      panFrom: LiveZoomPoint | null;
      shown: LiveZoomPoint;
    } | null = null;

    const clearWheelTimer = () => {
      if (wheelTimer !== null) clearTimeout(wheelTimer);
      wheelTimer = null;
    };

    function begin(): boolean {
      if (preview) return true;
      if (landingRef.current) return false;
      const current = optionsRef.current;
      if (current.canBegin && !current.canBegin()) return false;
      if (current.onBegin) flushSync(() => current.onBegin?.());
      const viewport = current.scroller?.() ?? current.viewport();
      const surface = current.surface();
      // (No surface on screen yet: the gesture still zooms, when it rests.)
      if (!viewport || !surface) return false;
      baseLevel = levelRef.current;
      if (!(baseLevel > 0)) return false;
      bounds = current.bounds();
      snaps = (current.snapLevels?.() ?? [])
        .filter((level) => Number.isFinite(level) && level >= bounds.min && level <= bounds.max)
        .map((level) => level / baseLevel);
      preview = createLiveZoomPreview({
        viewport,
        surface,
        anchor: current.anchor ?? "pointer",
        composite: current.composite,
        clampToScrollStart: current.clampToScrollStart,
      });
      return true;
    }

    function setScale(scale: number, point: LiveZoomPoint | null, magnetic = false): void {
      if (!preview) return;
      const min = bounds.min / baseLevel;
      const max = bounds.max / baseLevel;
      const clamped = Math.min(max, Math.max(min, scale));
      const bounded = magnetic && snaps.length > 0 ? liveZoomSnapped(clamped, snaps) : clamped;
      if (point) focal = point;
      preview.update(bounded, focal ?? undefined);
    }

    function settle(): void {
      clearWheelTimer();
      pinching = false;
      const active = preview;
      if (!active) return;
      preview = null;
      const level = baseLevel * active.scale;
      const baseAtCommit = baseLevel;
      const capture =
        optionsRef.current.correctScroll === false
          ? null
          : active.capture(focal);
      const anchorPoint = active.fixedPoint() ?? focal;
      focal = null;
      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        const landing = landingRef.current;
        if (landing?.timer) clearTimeout(landing.timer);
        landingRef.current = null;
        // A refused commit leaves the zoom (and so the scroll) where it was.
        const committed = levelRef.current;
        // A pan the preview showed is kept as the scroll it stood for,
        // whether or not the level changed.
        active.finish(
          committed === baseAtCommit && !active.panned ? null : capture,
          committed / level,
        );
      };
      if (Math.abs(active.scale - 1) < 1e-4) {
        finish();
        return;
      }
      landingRef.current = { finish, timer: null };
      let result: unknown;
      try {
        flushSync(() => {
          result = optionsRef.current.onCommit(level, anchorPoint);
        });
      } catch (error) {
        finish();
        throw error;
      }
      if (done) return;
      if (result && typeof (result as Promise<unknown>).then === "function") {
        const landing = landingRef.current;
        if (landing) {
          landing.timer = setTimeout(finish, LIVE_ZOOM_COMMIT_TIMEOUT_MS);
        }
        (result as Promise<unknown>).then(
          (applied) => {
            if (applied === false) finish();
          },
          () => finish(),
        );
        return;
      }
      // Nothing changed (the level was already there, or was refused).
      finish();
    }

    const accept = optionsRef.current.accept;

    const onWheel = (event: WheelEvent) => {
      if (!wheel) return;
      if (!event.ctrlKey && !event.metaKey) return;
      if (accept && !preview && !accept(event)) return;
      event.preventDefault();
      if (pinching) return;
      if (!preview && optionsRef.current.ignoreWheel?.(event)) return;
      if (!begin()) return;
      const factor = (optionsRef.current.wheelFactor ?? liveWheelZoomFactor)(
        event,
        modifierHeld,
      );
      if (!Number.isFinite(factor) || factor <= 0) return;
      setScale((preview?.scale ?? 1) * factor, {
        x: event.clientX,
        y: event.clientY,
      });
      clearWheelTimer();
      wheelTimer = setTimeout(settle, LIVE_ZOOM_WHEEL_SETTLE_MS);
    };

    // A pen or a mouse pressing mid-gesture works on the real zoom.
    const onPointerDown = (event: PointerEvent) => {
      if (preview && event.pointerType !== "touch") settle();
    };

    node.addEventListener("wheel", onWheel, { passive: false });
    node.addEventListener("pointerdown", onPointerDown, true);
    const detachPinch = attachPinchGestures(
      node,
      {
        onStart: (center: PinchPoint) => {
          clearWheelTimer();
          if (!begin()) return;
          pinching = true;
          pinchBase = preview?.scale ?? 1;
          focal = center;
          pinchState = { start: center, scaleFrom: null, panFrom: null, shown: center };
        },
        onChange: (scale: number, center: PinchPoint) => {
          if (!preview || !pinching) return;
          const state = pinchState;
          if (!state) {
            setScale(pinchBase * scale, center, true);
            return;
          }
          // Each is let in at the deadband's edge (where a slow gesture
          // crosses it), so a quick one that leaps past it is not held back.
          if (state.scaleFrom === null && Math.abs(Math.log(scale)) > Math.log(1 + PINCH_SCALE_DEADBAND)) {
            state.scaleFrom = scale > 1 ? 1 + PINCH_SCALE_DEADBAND : 1 / (1 + PINCH_SCALE_DEADBAND);
          }
          const travel = Math.hypot(center.x - state.start.x, center.y - state.start.y);
          if (state.panFrom === null && travel > PINCH_PAN_DEADBAND_PX) {
            const reach = PINCH_PAN_DEADBAND_PX / travel;
            state.panFrom = {
              x: state.start.x + (center.x - state.start.x) * reach,
              y: state.start.y + (center.y - state.start.y) * reach,
            };
          }
          // Let in, each starts where it was held and catches up with the
          // fingers over the next stretch: no jump, and the content ends
          // under the fingers.
          let scaleNow = 1;
          if (state.scaleFrom !== null) {
            const past = Math.abs(Math.log(scale / state.scaleFrom));
            const lag = Math.max(0, 1 - past / Math.log(1 + PINCH_SCALE_CATCH_UP));
            scaleNow = scale / Math.pow(state.scaleFrom, lag);
          }
          // The page follows the fingers only once a pan is meant (the
          // browser's own pan is held off for the pinch: one mover).
          let shown = state.shown;
          if (state.panFrom !== null && preventPan) {
            const past = Math.hypot(center.x - state.panFrom.x, center.y - state.panFrom.y);
            const lag = Math.max(0, 1 - past / PINCH_PAN_CATCH_UP_PX);
            shown = {
              x: center.x - lag * (state.panFrom.x - state.start.x),
              y: center.y - lag * (state.panFrom.y - state.start.y),
            };
          }
          // Scale about where the content is shown, then carry it along.
          setScale(pinchBase * scaleNow, state.shown, true);
          preview.translateBy(shown.x - state.shown.x, shown.y - state.shown.y);
          state.shown = shown;
          focal = shown;
        },
        onEnd: () => {
          pinchState = null;
          if (pinching) settle();
        },
      },
      { touch, gesture: true, preventPanWhilePinching: preventPan, accept },
    );

    return () => {
      node.removeEventListener("wheel", onWheel);
      node.removeEventListener("pointerdown", onPointerDown, true);
      detachPinch();
      releaseKeys();
      clearWheelTimer();
      const active = preview;
      preview = null;
      active?.finish(null);
    };
  }, [enabled, wheel, touch, preventPan, options.attachKey]);

  // Unmounting mid-landing: nothing will commit any more.
  useEffect(
    () => () => {
      const landing = landingRef.current;
      if (landing) landing.finish();
    },
    [],
  );
}
