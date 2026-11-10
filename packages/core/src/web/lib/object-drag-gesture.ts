"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Object drag gesture — one pointer, captured, cancellable — shared by every
 * editor that moves, resizes or turns an object with the pointer (Docs
 * pictures and shapes, PDF text boxes and pictures, and the deck, sheet and
 * code surfaces that wire it later).
 *
 * Office's touch rule for objects (Office Touch Guide; CONTRACT.md "对象"):
 * tap an object to select it, then hold and drag to move it; drag a handle
 * to resize or turn it; a finger swiped across the page scrolls. So:
 *
 * - A mouse (and a pen on a device that can hover) starts at once, exactly
 *   as before — this is only the mouse path written once.
 * - A finger on something already selected, or on a handle (`immediate`),
 *   starts at once too. Such targets wear `touch-action: none` in CSS while
 *   they are selected so the browser does not take the finger for a scroll.
 * - A finger on anything else must hold still for TOUCH_DRAG_HOLD_MS (300 ms,
 *   within TOUCH_DRAG_SLOP_PX); travelling first is a scroll and the gesture
 *   quietly stands down, and a quick release is a tap. Once lifted the
 *   object takes the pointer: it is captured, the element wears a temporary
 *   `touch-action: none`, and `installObjectDragTouchGuard` cancels the
 *   touchmoves the browser would otherwise turn into a scroll.
 * - Only the pointer that started counts. A second finger before the lift
 *   abandons the gesture (a pinch); after it, the gesture is rolled back.
 * - `pointercancel` (the browser took the touch, a capture was torn down),
 *   Escape and a window blur roll back: `onCancel`, never `onEnd`.
 *
 * Modelled on pdf/usePdfBoxDrag.ts, which already handled cancellation
 * correctly; this is the same loop without the PDF geometry, usable from a
 * React component (`useObjectDragGesture`) or a framework-free node view.
 */

import { useCallback, useEffect, useMemo, useRef } from "react";

import { isTouchLikePointerType } from "./pointer-capability";
import {
  TOUCH_DRAG_HOLD_MS,
  TOUCH_DRAG_SLOP_PX,
  yieldTouchHold,
} from "./touch-hold";

export { TOUCH_DRAG_HOLD_MS, TOUCH_DRAG_SLOP_PX };

/** Travel (px) below which a release is a click, not a drag. */
export const OBJECT_DRAG_CLICK_SLOP_PX = 3;

export type ObjectDragCancelReason =
  /** A touch released before it lifted: an ordinary tap. */
  | "tap"
  /** A touch travelled before it lifted: the page scrolls. */
  | "scroll"
  /** pointercancel, a lost capture or a window blur. */
  | "cancel"
  /** Another finger came down (a pinch takes over). */
  | "second-pointer"
  | "escape"
  /** Another gesture began, or the host called `cancel()`. */
  | "aborted";

/** The parts of a pointer event the gesture reads (a jsdom MouseEvent too). */
export interface ObjectDragPointer {
  readonly clientX: number;
  readonly clientY: number;
  readonly pointerId?: number;
  readonly pointerType?: string;
  readonly button?: number;
  readonly buttons?: number;
  readonly shiftKey?: boolean;
  readonly target?: EventTarget | null;
  readonly currentTarget?: EventTarget | null;
  readonly timeStamp?: number;
}

export interface ObjectDragDelta {
  dx: number;
  dy: number;
}

export interface ObjectDragHandlers {
  /** The gesture has the pointer: at once, or once a held finger lifted. */
  onLift?: (event: ObjectDragPointer) => void;
  /** A move, once the pointer has travelled past the click slop. */
  onMove: (event: PointerEvent, delta: ObjectDragDelta) => void;
  /** The release. `moved` is false for a press that never travelled. */
  onEnd: (event: PointerEvent, moved: boolean, delta: ObjectDragDelta) => void;
  /** The gesture ended without a release to commit: roll the preview back. */
  onCancel?: (reason: ObjectDragCancelReason, lifted: boolean) => void;
}

export interface ObjectDragOptions {
  /**
   * The element that takes the pointer (capture) and wears the temporary
   * `touch-action: none`. Defaults to the event's currentTarget, then target.
   */
  element?: Element | null;
  /**
   * A touch starts at once (a selected object, a handle) instead of after
   * the hold. A mouse always starts at once.
   */
  immediate?: boolean;
  clickSlopPx?: number;
  holdMs?: number;
  holdSlopPx?: number;
  /** Escape rolls the gesture back (default true). */
  escapeCancels?: boolean;
  /**
   * `onMove` runs at most once per frame, with the latest move: for a host
   * whose move does layout reads or hit tests (a finger samples faster than
   * the screen draws on an iPad). The release and the cancel are never
   * delayed, and a move still pending when they come is dropped (the
   * release carries the final position).
   */
  movesPerFrame?: boolean;
}

export interface ObjectDragGesture {
  readonly pointerId: number | undefined;
  readonly touch: boolean;
  /** True once the gesture has the pointer. */
  readonly lifted: boolean;
  /** True once the pointer travelled past the click slop. */
  readonly moved: boolean;
  /** Roll back and stop listening (no-op once ended). */
  cancel: (reason?: ObjectDragCancelReason) => void;
}

/** Touch pointers whose gesture has lifted: their touchmoves never scroll. */
const liftedTouchPointers = new Set<number | undefined>();
/** The one gesture in flight: a new one ends it (one object at a time). */
let current: { cancel: (reason?: ObjectDragCancelReason) => void } | null = null;

function asElement(target: EventTarget | null | undefined): Element | null {
  return typeof Element !== "undefined" && target instanceof Element ? target : null;
}

/** True when a gesture of this kind takes the touch path. */
export function objectDragIsTouch(event: { pointerType?: string }): boolean {
  return isTouchLikePointerType(event.pointerType);
}

/**
 * Start a drag from a pointerdown (React's or the DOM's). Returns the live
 * gesture; the handlers do the host's work.
 */
export function beginObjectDragGesture(
  event: ObjectDragPointer,
  handlers: ObjectDragHandlers,
  options: ObjectDragOptions = {},
): ObjectDragGesture {
  current?.cancel("aborted");
  const touch = objectDragIsTouch(event);
  const pointerId = event.pointerId;
  const startX = event.clientX;
  const startY = event.clientY;
  const element =
    options.element ?? asElement(event.currentTarget) ?? asElement(event.target);
  const clickSlop = options.clickSlopPx ?? OBJECT_DRAG_CLICK_SLOP_PX;
  const holdSlop = options.holdSlopPx ?? TOUCH_DRAG_SLOP_PX;
  const escapeCancels = options.escapeCancels ?? true;
  const holdMs = options.holdMs ?? TOUCH_DRAG_HOLD_MS;
  const held = touch && !options.immediate;
  let lifted = false;
  let moved = false;
  let ended = false;
  let holdTimer: ReturnType<typeof setTimeout> | null = null;
  let restoreTouchAction: (() => void) | null = null;
  const movesPerFrame =
    options.movesPerFrame === true && typeof requestAnimationFrame === "function";
  let pendingMove: PointerEvent | null = null;
  let moveFrame = 0;

  const matches = (candidate: { pointerId?: number }) =>
    pointerId === undefined ||
    candidate.pointerId === undefined ||
    candidate.pointerId === pointerId;
  const deltaOf = (candidate: { clientX: number; clientY: number }) => ({
    dx: candidate.clientX - startX,
    dy: candidate.clientY - startY,
  });

  const stop = () => {
    ended = true;
    if (holdTimer !== null) clearTimeout(holdTimer);
    holdTimer = null;
    if (moveFrame) cancelAnimationFrame(moveFrame);
    moveFrame = 0;
    pendingMove = null;
    liftedTouchPointers.delete(pointerId);
    restoreTouchAction?.();
    restoreTouchAction = null;
    if (pointerId !== undefined) {
      try {
        if (element?.hasPointerCapture?.(pointerId)) {
          element.releasePointerCapture(pointerId);
        }
      } catch {
        /* best effort */
      }
    }
    window.removeEventListener("pointermove", onMove, true);
    window.removeEventListener("pointerup", onUp, true);
    window.removeEventListener("mouseup", onMouseUp, true);
    window.removeEventListener("pointercancel", onPointerCancel, true);
    window.removeEventListener("pointerdown", onOtherDown, true);
    window.removeEventListener("keydown", onKeyDown, true);
    window.removeEventListener("blur", onBlur);
    if (current === gesture) current = null;
  };

  const cancel = (reason: ObjectDragCancelReason = "aborted") => {
    if (ended) return;
    const wasLifted = lifted;
    stop();
    handlers.onCancel?.(reason, wasLifted);
  };

  const lift = () => {
    if (ended || lifted) return;
    lifted = true;
    holdTimer = null;
    if (touch) {
      liftedTouchPointers.add(pointerId);
      // a handle or a selected object: the hold has nothing to show; a
      // movable object marked for it keeps its menu for a still release
      if (!element?.closest?.('[data-touch-hold="drag"]')) {
        yieldTouchHold(pointerId);
      }
      if (element instanceof HTMLElement || element instanceof SVGElement) {
        const style = (element as HTMLElement).style;
        const previous = style.getPropertyValue("touch-action");
        const priority = style.getPropertyPriority("touch-action");
        style.setProperty("touch-action", "none", "important");
        element.setAttribute("data-object-drag-lifted", "true");
        document.documentElement.setAttribute("data-object-drag-touch", "true");
        restoreTouchAction = () => {
          if (previous) style.setProperty("touch-action", previous, priority);
          else style.removeProperty("touch-action");
          element.removeAttribute("data-object-drag-lifted");
          if (liftedTouchPointers.size === 0) {
            document.documentElement.removeAttribute("data-object-drag-touch");
          }
        };
      }
    }
    if (pointerId !== undefined) {
      try {
        element?.setPointerCapture?.(pointerId);
      } catch {
        /* capture is best-effort (jsdom, a detached element) */
      }
    }
    handlers.onLift?.(event);
  };

  function onMove(move: PointerEvent) {
    if (ended || !matches(move)) return;
    const delta = deltaOf(move);
    const distance = Math.hypot(delta.dx, delta.dy);
    if (!lifted) {
      // a finger that travels before the lift is scrolling the page
      if (distance > holdSlop) cancel("scroll");
      return;
    }
    // a mouse button released where we could not see it: a ghost drag
    if (move.pointerType === "mouse" && typeof move.buttons === "number" && (move.buttons & 1) === 0) {
      cancel("cancel");
      return;
    }
    if (!moved) {
      if (distance < clickSlop) return;
      moved = true;
      if (touch) yieldTouchHold(pointerId);
    }
    if (move.cancelable) move.preventDefault();
    if (!movesPerFrame) {
      handlers.onMove(move, delta);
      return;
    }
    pendingMove = move;
    if (!moveFrame) moveFrame = requestAnimationFrame(flushMove);
  }

  function flushMove() {
    moveFrame = 0;
    const move = pendingMove;
    pendingMove = null;
    if (ended || !move) return;
    handlers.onMove(move, deltaOf(move));
  }

  function onUp(up: PointerEvent) {
    if (ended || !matches(up)) return;
    if (!lifted) {
      cancel("tap");
      return;
    }
    // A busy main thread runs the hold's timer before the finger-up waiting
    // behind it: the event's own clock says the finger left in time, so this
    // was a tap and its click is the tap's (a bookmark or a button under a
    // finger did nothing, and the second press felt like lag).
    if (
      held &&
      !moved &&
      typeof event.timeStamp === "number" &&
      event.timeStamp > 0 &&
      up.timeStamp - event.timeStamp < holdMs
    ) {
      cancel("tap");
      return;
    }
    const wasMoved = moved;
    stop();
    handlers.onEnd(up, wasMoved, deltaOf(up));
  }

  // The release of a mouse whose pointerup never reached the window: Safari
  // sends a captured pointer's events to the element that holds the capture
  // even after a re-render took that element off the page, where no window
  // listener hears them. The gesture then never ended and its preview stayed
  // up (a SmartArt shape pressed once stayed 92% opaque, its slide text
  // showing through). The compatibility mouseup is not captured and reaches
  // the window from the element under the pointer.
  function onMouseUp(up: MouseEvent) {
    if (ended || touch) return;
    onUp(up as unknown as PointerEvent);
  }

  function onPointerCancel(cancelled: PointerEvent) {
    if (ended || !matches(cancelled)) return;
    cancel(lifted ? "cancel" : "scroll");
  }

  function onOtherDown(down: PointerEvent) {
    if (ended || down.pointerId === undefined || down.pointerId === pointerId) return;
    cancel("second-pointer");
  }

  function onKeyDown(key: KeyboardEvent) {
    if (!escapeCancels || key.key !== "Escape" || !lifted) return;
    key.preventDefault();
    key.stopPropagation();
    cancel("escape");
  }

  function onBlur() {
    cancel("cancel");
  }

  const gesture: ObjectDragGesture = {
    get pointerId() {
      return pointerId;
    },
    get touch() {
      return touch;
    },
    get lifted() {
      return lifted;
    },
    get moved() {
      return moved;
    },
    cancel,
  };
  current = gesture;

  window.addEventListener("pointermove", onMove, true);
  window.addEventListener("pointerup", onUp, true);
  window.addEventListener("mouseup", onMouseUp, true);
  window.addEventListener("pointercancel", onPointerCancel, true);
  window.addEventListener("pointerdown", onOtherDown, true);
  window.addEventListener("keydown", onKeyDown, true);
  window.addEventListener("blur", onBlur);

  if (!touch || options.immediate) {
    lift();
  } else {
    holdTimer = setTimeout(lift, holdMs);
  }
  return gesture;
}

/**
 * Make `element` a movable object for a finger: a hold lifts it (touch-hold's
 * drag mode, so the menu waits for a still release) and, once a gesture over
 * it has lifted, the browser may not turn the finger's moves into a scroll.
 *
 * The touchmove listener has to be on the element before the finger lands —
 * Chrome decides at touchstart whether a touch's moves can be cancelled at
 * all — and it is non-passive only over the object, so scrolling elsewhere
 * never waits for the main thread. Returns the cleanup.
 */
export function installObjectDragTouchGuard(element: Element): () => void {
  const hadHoldMode = element.getAttribute("data-touch-hold");
  if (hadHoldMode === null) element.setAttribute("data-touch-hold", "drag");
  const onTouchMove = (event: Event) => {
    if (liftedTouchPointers.size > 0 && event.cancelable) event.preventDefault();
  };
  element.addEventListener("touchmove", onTouchMove, { passive: false });
  return () => {
    element.removeEventListener("touchmove", onTouchMove);
    if (hadHoldMode === null && element.getAttribute("data-touch-hold") === "drag") {
      element.removeAttribute("data-touch-hold");
    }
  };
}

/**
 * `installObjectDragTouchGuard` as a stable callback ref (React 19 runs the
 * returned cleanup on detach), for movable objects rendered in a list where
 * a hook per item is not at hand.
 */
export function objectDragTouchGuardRef(element: Element | null): (() => void) | undefined {
  return element ? installObjectDragTouchGuard(element) : undefined;
}

/** True while a touch gesture has lifted an object (tests, diagnostics). */
export function objectDragTouchActive(): boolean {
  return liftedTouchPointers.size > 0;
}

/**
 * React: the gesture as a stable `begin`, ended when the component unmounts
 * (a box taken away mid-gesture rolls back instead of dangling).
 */
export function useObjectDragGesture(): {
  begin: (
    event: ObjectDragPointer,
    handlers: ObjectDragHandlers,
    options?: ObjectDragOptions,
  ) => ObjectDragGesture;
  cancel: (reason?: ObjectDragCancelReason) => void;
} {
  const gestureRef = useRef<ObjectDragGesture | null>(null);
  useEffect(() => () => gestureRef.current?.cancel("aborted"), []);
  const begin = useCallback(
    (event: ObjectDragPointer, handlers: ObjectDragHandlers, options?: ObjectDragOptions) => {
      const gesture = beginObjectDragGesture(event, handlers, options);
      gestureRef.current = gesture;
      return gesture;
    },
    [],
  );
  const cancel = useCallback((reason?: ObjectDragCancelReason) => {
    gestureRef.current?.cancel(reason);
  }, []);
  return useMemo(() => ({ begin, cancel }), [begin, cancel]);
}

/**
 * React: a guard on every element matching `selector` inside `rootRef`,
 * re-installed whenever `revision` changes (the list of objects rendered).
 */
export function useObjectDragTouchGuards(
  rootRef: { readonly current: Element | null },
  selector: string,
  revision: unknown,
): void {
  useEffect(() => {
    const root = rootRef.current;
    if (!root) return;
    const cleanups = Array.from(root.querySelectorAll(selector)).map(
      installObjectDragTouchGuard,
    );
    return () => cleanups.forEach((cleanup) => cleanup());
  }, [rootRef, selector, revision]);
}

/**
 * React: `installObjectDragTouchGuard` as a callback ref, for the element a
 * finger holds to move an object.
 */
export function useObjectDragTouchGuard(): (element: Element | null) => void {
  const cleanupRef = useRef<(() => void) | null>(null);
  useEffect(() => () => cleanupRef.current?.(), []);
  return useCallback((element: Element | null) => {
    cleanupRef.current?.();
    cleanupRef.current = element ? installObjectDragTouchGuard(element) : null;
  }, []);
}
