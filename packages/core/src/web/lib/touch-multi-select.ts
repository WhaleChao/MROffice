"use client";































/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useState, useSyncExternalStore } from "react";

import { isTouchLikePointerType } from "./pointer-capability";
import { TOUCH_DRAG_HOLD_MS } from "./touch-hold";

/** How long the resting finger must have been down before the tap. */
export const TOUCH_MULTI_SELECT_HOLD_MS = TOUCH_DRAG_HOLD_MS;
/** A resting finger that travelled further than this is a pan or a drag. */
export const TOUCH_MULTI_SELECT_SLOP_PX = 12;

export const TOUCH_MULTI_SELECT_ANCHOR_SELECTOR =
  '[data-selected="true"], [aria-selected="true"], [aria-pressed="true"], [data-multiselected="true"]';

interface Finger {
  id: number;
  at: number;
  x: number;
  y: number;
  moved: boolean;
  target: Element | null;
}

const fingers = new Map<number, Finger>();
let tracking = false;

function clock(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

function trackFingers(): void {
  if (tracking || typeof window === "undefined") return;
  tracking = true;
  const options = { capture: true, passive: true } as const;
  window.addEventListener(
    "pointerdown",
    (event) => {
      if (!isTouchLikePointerType(event.pointerType)) return;
      // the first finger of a new touch: none of the others is still down
      // (a lost pointerup must not leave a phantom resting finger)
      if (event.isPrimary) fingers.clear();
      fingers.set(event.pointerId, {
        id: event.pointerId,
        at: clock(),
        x: event.clientX,
        y: event.clientY,
        moved: false,
        target: event.target instanceof Element ? event.target : null,
      });
    },
    options,
  );
  window.addEventListener(
    "pointermove",
    (event) => {
      const finger = fingers.get(event.pointerId);
      if (
        finger &&
        !finger.moved &&
        Math.hypot(event.clientX - finger.x, event.clientY - finger.y) > TOUCH_MULTI_SELECT_SLOP_PX
      ) {
        finger.moved = true;
      }
    },
    options,
  );
  const lift = (event: PointerEvent) => {
    fingers.delete(event.pointerId);
  };
  window.addEventListener("pointerup", lift, options);
  window.addEventListener("pointercancel", lift, options);
}

export interface TouchMultiSelectPress {
  pointerId: number;
  pointerType: string;
}

export interface TouchMultiSelectScope {
  /** Whether the mode is on (the status bar shows). */
  readonly active: boolean;
  /**
   * Ask at an object's pointerdown: true = add / remove this object instead
   * of replacing the selection (and do not start a drag). Turns the mode on
   * when a resting finger made it so.
   */
  isAdditivePress(event: TouchMultiSelectPress): boolean;
  enter(): void;
  exit(): void;
  subscribe(listener: () => void): () => void;
}

export interface TouchMultiSelectOptions {
  /** What counts as a selected object under the resting finger. */
  anchorSelector?: string;
  /** Only resting fingers inside this element count (default: anywhere). */
  root?: () => Element | null;
}

export function createTouchMultiSelectScope(
  options: TouchMultiSelectOptions = {},
): TouchMultiSelectScope {
  trackFingers();
  const anchorSelector = options.anchorSelector ?? TOUCH_MULTI_SELECT_ANCHOR_SELECTOR;
  const listeners = new Set<() => void>();
  let active = false;
  const set = (next: boolean) => {
    if (next === active) return;
    active = next;
    for (const listener of listeners) listener();
  };
  const restingOnSelected = (pointerId: number): boolean => {
    const now = clock();
    const root = options.root?.() ?? null;
    for (const finger of fingers.values()) {
      if (finger.id === pointerId || finger.moved) continue;
      if (now - finger.at < TOUCH_MULTI_SELECT_HOLD_MS) continue;
      const anchor = finger.target?.closest(anchorSelector);
      if (!anchor || (root && !root.contains(anchor))) continue;
      return true;
    }
    return false;
  };
  return {
    get active() {
      return active;
    },
    isAdditivePress(event) {
      if (!isTouchLikePointerType(event.pointerType)) return false;
      if (active) return true;
      if (!restingOnSelected(event.pointerId)) return false;
      set(true);
      return true;
    },
    enter: () => set(true),
    exit: () => set(false),
    subscribe(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  };
}

const scopes = new WeakMap<object, TouchMultiSelectScope>();

/**
 * One scope per surface object (a Docs editor), shared by the code that
 * reads presses (a node view) and the code that draws the bar (React).
 */
export function touchMultiSelectScopeFor(
  key: object,
  options?: TouchMultiSelectOptions,
): TouchMultiSelectScope {
  let scope = scopes.get(key);
  if (!scope) {
    scope = createTouchMultiSelectScope(options);
    scopes.set(key, scope);
  }
  return scope;
}

/** Live `active` of a scope, for render decisions. */
export function useTouchMultiSelectActive(scope: TouchMultiSelectScope): boolean {
  return useSyncExternalStore(
    scope.subscribe,
    () => scope.active,
    () => false,
  );
}

/**
 * A React surface's scope. `count` is how many objects are selected now:
 * the mode ends by itself when it reaches 0 (a tap on empty space).
 */
export function useTouchMultiSelect({
  count,
  ...options
}: TouchMultiSelectOptions & { count: number }): TouchMultiSelectScope & {
  active: boolean;
} {
  const [scope] = useState(() => createTouchMultiSelectScope(options));
  const active = useTouchMultiSelectActive(scope);
  useEffect(() => {
    if (active && count === 0) scope.exit();
  }, [active, count, scope]);
  return {
    active,
    isAdditivePress: scope.isAdditivePress,
    enter: scope.enter,
    exit: scope.exit,
    subscribe: scope.subscribe,
  };
}
