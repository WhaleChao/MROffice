"use client";






























/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useLayoutEffect, useRef, type RefObject } from "react";

import {
  clampZoom,
  EDITOR_ZOOM_BOUNDS,
  EDITOR_ZOOM_STEP,
  wheelDeltaPixels,
  zoomAnchoredScroll,
  type EditorZoomBounds,
  type ZoomLevel,
} from "./editor-zoom";
import { fingerTouches, inkPalmGuard, inkPalmTouchKey } from "./ink-palm-guard";
import { yieldTouchHold } from "./touch-hold";

/** One wheel notch (100 px) is one zoom step, multiplicatively (editor-zoom). */
const WHEEL_ZOOM_RATE = Math.log(1 + EDITOR_ZOOM_STEP) / 100;

/** Fingers closer than this (px) carry no usable distance ratio. */
export const PINCH_MIN_DISTANCE_PX = 8;

export type PinchSource = "touch" | "gesture";

export interface PinchPoint {
  x: number;
  y: number;
}

export interface PinchGestureHandlers {
  /** A pinch began (two fingers down, or Safari's gesturestart). */
  onStart?: (center: PinchPoint, source: PinchSource) => void;
  /** `scale` is relative to the pinch's start (1 = unchanged). */
  onChange: (scale: number, center: PinchPoint, source: PinchSource) => void;
  onEnd?: (source: PinchSource) => void;
}

export interface PinchGestureOptions {
  /** Two-finger touch pinch (default true). */
  touch?: boolean;
  /** Safari's gesturestart/change/end (default true). */
  gesture?: boolean;
  /**
   * While two fingers are down, cancel their touchmoves so the content
   * neither pans nor zooms natively under the pinch (default false). It makes
   * the node's touchmove listener non-passive, so the first scroll frame of
   * every touch on it waits for the main thread: leave it off where the
   * main thread is busy (a text editor).
   */
  preventPanWhilePinching?: boolean;
  /** Give the node `touch-action: pan-x pan-y` while attached (default true). */
  setTouchAction?: boolean;
  /**
   * Whether a pinch that begins with this event (a two-finger touchstart, or
   * Safari's gesturestart) is this node's. A host whose descendants pinch
   * their own content (the window canvas: a document window inside it) says
   * no for those, and the whole gesture is left to them. Default: every one.
   */
  accept?: (event: Event) => boolean;
}

/** Safari's pinch events, which are not in the DOM typings. */
interface SafariGestureEvent extends Event {
  readonly scale: number;
  readonly clientX: number;
  readonly clientY: number;
}

/**
 * The two fingers of a pinch: finger contacts only. A pen and the hand that
 * holds it are two contacts, and so are a resting hand's heel and little
 * finger; neither pair is a pinch (lib/ink-palm-guard), and a pinch that
 * read them zoomed the page under the pen while one wrote.
 */
function touchPair(list: TouchList): [Touch, Touch] | null {
  const fingers = fingerTouches(list);
  if (fingers.length !== 2) return null;
  return [fingers[0], fingers[1]];
}

function distanceOf([a, b]: [Touch, Touch]): number {
  return Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
}

function centerOf([a, b]: [Touch, Touch]): PinchPoint {
  return { x: (a.clientX + b.clientX) / 2, y: (a.clientY + b.clientY) / 2 };
}

/**
 * An `accept` for a host whose children pinch their own content: the pinch
 * is the host's only when no finger (nor Safari's gesture target) lands
 * inside `selector` — the window canvas leaves a pinch on a window to it.
 */
export function pinchUnlessInside(selector: string): (event: Event) => boolean {
  return (event) => {
    const list = (event as TouchEvent).touches;
    const targets: Array<EventTarget | null> = list
      ? Array.from({ length: list.length }, (_, index) => list.item(index)?.target ?? event.target)
      : [event.target];
    return targets.every(
      (target) => !(target instanceof Element && target.closest(selector)),
    );
  };
}









let glassTouches: TouchList | null = null;
let glassWatchers = 0;
const watchGlass = (event: TouchEvent) => {
  glassTouches = event.touches;
};
function watchGlassTouches(): () => void {
  if (typeof window === "undefined") return () => {};
  if (glassWatchers === 0) {
    for (const type of ["touchstart", "touchmove", "touchend", "touchcancel"] as const) {
      window.addEventListener(type, watchGlass, { capture: true, passive: true });
    }
  }
  glassWatchers += 1;
  return () => {
    glassWatchers -= 1;
    if (glassWatchers > 0) return;
    for (const type of ["touchstart", "touchmove", "touchend", "touchcancel"] as const) {
      window.removeEventListener(type, watchGlass, { capture: true });
    }
    glassTouches = null;
  };
}

/** Whether contacts are on the glass: a gesture now is a touch screen's. */
function touchesOnGlass(): boolean {
  return (glassTouches?.length ?? 0) > 0;
}

/** Listen for pinches on `node`. Returns the cleanup. */
export function attachPinchGestures(
  node: HTMLElement,
  handlers: PinchGestureHandlers,
  options: PinchGestureOptions = {},
): () => void {
  const touch = options.touch ?? true;
  const gesture = options.gesture ?? true;
  const preventPan = options.preventPanWhilePinching ?? false;
  const accept = options.accept;
  let touchStartDistance = 0;
  let touchActive = false;
  /** Safari reports this pinch as a gesture: the touch path stands down. */
  let gestureActive = false;
  const cleanups: Array<() => void> = [];

  if (touch && (options.setTouchAction ?? true)) {
    const previous = node.style.getPropertyValue("touch-action");
    node.style.setProperty("touch-action", "pan-x pan-y");
    // the attribute carries the same rule (globals.css) for a DOM whose
    // style object does not know the property
    node.setAttribute("data-pinch-zoom-surface", "");
    cleanups.push(() => {
      if (previous) node.style.setProperty("touch-action", previous);
      else node.style.removeProperty("touch-action");
      node.removeAttribute("data-pinch-zoom-surface");
    });
  }

  if (touch) {
    /** The two contacts the running pinch began with, and where it last was. */
    let pinchTouches: [number, number] | null = null;
    let lastCenter: PinchPoint | null = null;
    const endTouch = () => {
      if (!touchActive) return;
      // A "pinch" one of whose fingers was the writing hand never was one:
      // the zoom goes back to where it started.
      if (
        !gestureActive &&
        lastCenter &&
        pinchTouches?.some((identifier) => inkPalmGuard.isPalm(inkPalmTouchKey(identifier)))
      ) {
        handlers.onChange(1, lastCenter, "touch");
      }
      touchActive = false;
      touchStartDistance = 0;
      pinchTouches = null;
      lastCenter = null;
      handlers.onEnd?.("touch");
    };
    const onTouchStart = (event: TouchEvent) => {
      const pair = touchPair(event.touches);
      if (!pair) {
        // a third finger, or one of the two turned out to be the hand
        endTouch();
        return;
      }
      const distance = distanceOf(pair);
      if (distance < PINCH_MIN_DISTANCE_PX) return;
      if (!touchActive && accept && !accept(event)) return;
      // a pinch, not a hold: whatever the first finger began stands down
      yieldTouchHold();
      touchStartDistance = distance;
      touchActive = true;
      pinchTouches = [pair[0].identifier, pair[1].identifier];
      lastCenter = centerOf(pair);
      if (!gestureActive) handlers.onStart?.(lastCenter, "touch");
    };
    const onTouchMove = (event: TouchEvent) => {
      if (!touchActive) return;
      const pair = touchPair(event.touches);
      if (!pair) {
        // the pen arrived: what looked like two fingers was its hand
        if (fingerTouches(event.touches).length < 2) endTouch();
        return;
      }
      if (preventPan && event.cancelable) event.preventDefault();
      if (gestureActive) return;
      const distance = distanceOf(pair);
      if (distance < PINCH_MIN_DISTANCE_PX || touchStartDistance <= 0) return;
      lastCenter = centerOf(pair);
      handlers.onChange(distance / touchStartDistance, lastCenter, "touch");
    };
    const onTouchEnd = (event: TouchEvent) => {
      if (fingerTouches(event.touches).length < 2) endTouch();
    };
    // The pen arriving settles which contacts were its hand; the page hears
    // no more of them, so the pinch they faked ends here.
    cleanups.push(
      inkPalmGuard.subscribe((contacts) => {
        if (!touchActive || !pinchTouches) return;
        const [a, b] = pinchTouches;
        const keys = [inkPalmTouchKey(a), inkPalmTouchKey(b)];
        if (contacts.some((contact) => keys.includes(contact.key))) endTouch();
      }),
    );
    const moveOptions: AddEventListenerOptions = { passive: !preventPan };
    node.addEventListener("touchstart", onTouchStart, { passive: true });
    node.addEventListener("touchmove", onTouchMove, moveOptions);
    node.addEventListener("touchend", onTouchEnd, { passive: true });
    node.addEventListener("touchcancel", onTouchEnd, { passive: true });
    cleanups.push(() => {
      node.removeEventListener("touchstart", onTouchStart);
      node.removeEventListener("touchmove", onTouchMove, moveOptions);
      node.removeEventListener("touchend", onTouchEnd);
      node.removeEventListener("touchcancel", onTouchEnd);
    });
  }

  // Safari's gesture events are listened to whenever touch is, so an iPad's
  // finger pinch (reported both ways) is applied once; handled only when
  // `gesture` is on.
  if (gesture || touch) {
    cleanups.push(watchGlassTouches());
    /** Safari's gesture began somewhere `accept` left to someone else. */
    let gestureRejected = false;
    /**
     * On a touch screen with the touch path listening, the touch path is the
     * pinch (it knows a finger from the hand); Safari's gesture only keeps the
     * browser from zooming the page. Without it, the gesture zooms only while
     * two of its contacts are fingers.
     */
    let touchDriven = false;
    let started = false;
    /** Safari's scale when the pinch began (a late start measures from there). */
    let baseScale = 1;
    const fingersPinch = () => fingerTouches(glassTouches ?? ([] as unknown as TouchList)).length >= 2;
    const onGestureStart = (event: Event) => {
      gestureRejected = Boolean(accept && !accept(event));
      if (gestureRejected) return;
      touchDriven = touch && touchesOnGlass();
      started = false;
      if (touchDriven) {
        if (gesture) event.preventDefault();
        return;
      }
      gestureActive = true;
      if (!gesture) return;
      event.preventDefault();
      if (touchesOnGlass() && !fingersPinch()) return;
      started = true;
      baseScale = 1;
      const g = event as SafariGestureEvent;
      handlers.onStart?.({ x: g.clientX, y: g.clientY }, "gesture");
    };
    const onGestureChange = (event: Event) => {
      if (!gesture || gestureRejected) return;
      event.preventDefault();
      if (touchDriven) return;
      // A pinch Safari reads from a hand (a heel and a little finger) is no
      // pinch; one whose fingers come down later starts then.
      if (touchesOnGlass() && !fingersPinch()) return;
      const g = event as SafariGestureEvent;
      if (!Number.isFinite(g.scale) || g.scale <= 0) return;
      if (!started) {
        started = true;
        baseScale = g.scale;
        handlers.onStart?.({ x: g.clientX, y: g.clientY }, "gesture");
      }
      handlers.onChange(g.scale / baseScale, { x: g.clientX, y: g.clientY }, "gesture");
    };
    const onGestureEnd = (event: Event) => {
      if (gestureRejected) {
        gestureRejected = false;
        return;
      }
      if (touchDriven) {
        touchDriven = false;
        if (gesture) event.preventDefault();
        return;
      }
      gestureActive = false;
      if (!gesture) return;
      event.preventDefault();
      if (!started) return;
      started = false;
      handlers.onEnd?.("gesture");
    };
    const gestureOptions: AddEventListenerOptions = { passive: !gesture };
    node.addEventListener("gesturestart", onGestureStart, gestureOptions);
    node.addEventListener("gesturechange", onGestureChange, gestureOptions);
    node.addEventListener("gestureend", onGestureEnd, gestureOptions);
    cleanups.push(() => {
      node.removeEventListener("gesturestart", onGestureStart, gestureOptions);
      node.removeEventListener("gesturechange", onGestureChange, gestureOptions);
      node.removeEventListener("gestureend", onGestureEnd, gestureOptions);
    });
  }

  return () => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  };
}

export interface PinchZoomOptions {
  /** The scrolling viewport. Its scroll offsets are what anchoring adjusts. */
  viewportRef: RefObject<HTMLElement | null>;
  /** The current committed zoom level. */
  level: ZoomLevel;
  onZoom: (level: ZoomLevel) => void;
  /** Set false to leave the surface entirely alone (no listeners attached). */
  enabled?: boolean;
  /** Two-finger touch pinch (default true). */
  touch?: boolean;
  /**
   * Where the zoom is anchored: under the pointer / pinch centre (default),
   * or nowhere (a host that re-anchors its own way, e.g. on the viewport's
   * centre in a layout effect of its own).
   */
  anchor?: "pointer" | "none";
  /** The editor's zoom bounds (default 50 %–200 %). */
  bounds?: EditorZoomBounds;
  /** A pinch began: undo what its first finger started (a selection …). */
  onPinchStart?: () => void;
  /**
   * Changes whenever `viewportRef` comes to point at a different element.
   * The listeners attach to the element the ref holds when the effect runs;
   * a surface that mounts after the hook's first render (a sheet's grid
   * appears once its workbook has loaded) or is swapped for another must
   * say so, or the pinch is never heard.
   */
  attachKey?: unknown;
}

export function usePinchZoom({
  viewportRef,
  level,
  onZoom,
  enabled = true,
  touch = true,
  anchor = "pointer",
  bounds,
  onPinchStart,
  attachKey,
}: PinchZoomOptions): void {
  const levelRef = useRef(level);
  const onZoomRef = useRef(onZoom);
  const onPinchStartRef = useRef(onPinchStart);
  /** The live gesture level, unsnapped, so small nudges accumulate. */
  const liveRef = useRef<number>(level);
  /** Where a pinch started, since its scale is relative to that. */
  const pinchStartRef = useRef<number>(level);
  /** Scroll offsets waiting for the new scale to reach the DOM. */
  const pendingScrollRef = useRef<{ scrollLeft: number; scrollTop: number } | null>(
    null,
  );
  const boundsRef = useRef(bounds);
  const anchorRef = useRef(anchor);

  // The latest props, for listeners attached once (before any event fires).
  useLayoutEffect(() => {
    levelRef.current = level;
    onZoomRef.current = onZoom;
    onPinchStartRef.current = onPinchStart;
    boundsRef.current = bounds;
    anchorRef.current = anchor;
  });

  // Applied here rather than straight after `onZoom`, because the scrollable
  // area only grows once the new transform is in the DOM. Written any earlier
  // the offset is clamped to the old, smaller extent and the sheet slides out
  // from under the pointer by however much was cut.
  useLayoutEffect(() => {
    const pending = pendingScrollRef.current;
    if (!pending) return;
    pendingScrollRef.current = null;
    const node = viewportRef.current;
    if (!node) return;
    node.scrollLeft = pending.scrollLeft;
    node.scrollTop = pending.scrollTop;
  }, [level, viewportRef]);

  useEffect(() => {
    if (!enabled) return;
    const node = viewportRef.current;
    if (!node) return;

    const clamp = (value: number) => clampZoom(value, boundsRef.current);
    // the live level stays inside the bounds (unsnapped), so pinching past a
    // limit and back turns around at once
    const limit = (value: number) => {
      const bounded = boundsRef.current ?? EDITOR_ZOOM_BOUNDS;
      if (!Number.isFinite(value)) return levelRef.current;
      return Math.min(bounded.max, Math.max(bounded.min, value));
    };

    /**
     * The level a gesture continues from. Re-seeded whenever the level moved
     * by some other route (the ± buttons, reset, a reload), so a pinch never
     * resumes from a stale value.
     */
    function currentLevel(): number {
      if (clamp(liveRef.current) !== levelRef.current) {
        liveRef.current = levelRef.current;
      }
      return liveRef.current;
    }

    function commit(next: number, clientX: number, clientY: number): void {
      liveRef.current = next;
      const snapped = clamp(next);
      const from = levelRef.current;
      if (snapped === from) return;

      const viewport = viewportRef.current;
      if (!viewport) return;
      if (anchorRef.current === "pointer") {
        const rect = viewport.getBoundingClientRect();
        pendingScrollRef.current = zoomAnchoredScroll({
          scrollLeft: viewport.scrollLeft,
          scrollTop: viewport.scrollTop,
          pointerX: clientX - rect.left,
          pointerY: clientY - rect.top,
          from,
          to: snapped,
        });
      }
      levelRef.current = snapped;
      onZoomRef.current(snapped);
    }

    function handleWheel(event: WheelEvent): void {
      // Ctrl is what both a trackpad pinch and an explicit Ctrl+scroll report;
      // Cmd is the macOS habit. A plain wheel stays a plain scroll.
      if (!event.ctrlKey && !event.metaKey) return;
      event.preventDefault();
      const from = currentLevel();
      const pixels = wheelDeltaPixels(event.deltaY, event.deltaMode);
      commit(
        limit(from * Math.exp(-pixels * WHEEL_ZOOM_RATE)),
        event.clientX,
        event.clientY,
      );
    }

    node.addEventListener("wheel", handleWheel, { passive: false });
    const detach = attachPinchGestures(
      node,
      {
        onStart: () => {
          pinchStartRef.current = currentLevel();
          onPinchStartRef.current?.();
        },
        onChange: (scale, center) => {
          commit(
            limit(pinchStartRef.current * scale),
            center.x,
            center.y,
          );
        },
      },
      { touch, gesture: true },
    );
    return () => {
      node.removeEventListener("wheel", handleWheel);
      detach();
      pendingScrollRef.current = null;
    };
  }, [enabled, touch, viewportRef, attachKey]);
}
