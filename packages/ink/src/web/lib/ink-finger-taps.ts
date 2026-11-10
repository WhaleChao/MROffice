"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";

import { isHandSizedContact, isStylusTouch } from "@mrnote/core/web/lib/ink-palm-guard.ts";
import { inkPointerArbiter } from "@mrnote/core/web/lib/ink-pointer-policy.ts";

/**
 * Finger taps while drawing: two fingers tapped together undo, three redo,
 * the pairing every tablet note-taker teaches the hand, so the pen never
 * travels to the ribbon between a wrong stroke and the next one.
 *
 * A tap is strict, so a pinch never reads as one: every finger lands within
 * {@link INK_TAP_LAND_MS} of the first, none moves more than
 * {@link INK_TAP_SLOP_PX}, and all are off the glass within
 * {@link INK_TAP_MAX_MS} of the first landing. A pen stroke running at any
 * point spoils it (the fingers were the writing hand).
 *
 * Two fingers held still on the glass undo again and again until they lift
 * (Procreate's rapid undo): after {@link INK_RAPID_UNDO_DELAY_MS}, then every
 * {@link INK_RAPID_UNDO_EVERY_MS}. Never the pen's contact (Safari's
 * `Touch.touchType` "stylus") and never a contact the size of a hand.
 *
 * The touch events are only listened to (passive): the pan, the pinch and
 * the palm guard (lib/ink-touch-shield) keep theirs. The listeners are the
 * window's first, attached once when this module loads (the shield imports
 * it): the shield stops a hand's events at the window, and it calls a
 * finger landing within a second of the pen's lift, or any three fingers
 * beside a pen, the hand -- the very taps that undo after a stroke and
 * redo. Heard after the shield, those never arrived.
 */

/** Every finger of a tap lands this close to the first (ms). */
export const INK_TAP_LAND_MS = 150;
/** A tap's fingers are all lifted this soon after the first landed (ms). */
export const INK_TAP_MAX_MS = 300;
/** A finger that travels farther than this (px) is a pinch or a pan, not a tap. */
export const INK_TAP_SLOP_PX = 12;
/** Two fingers held still this long start undoing again and again (ms). */
export const INK_RAPID_UNDO_DELAY_MS = 520;
/** And undo once more this often while they stay (ms). */
export const INK_RAPID_UNDO_EVERY_MS = 170;

export type InkFingerTapCommand = "undo" | "redo";

export interface InkFingerTapTarget {
  /** Where the fingers must land (the page or the slide the pen draws on). */
  readonly root: HTMLElement;
  readonly run: (command: InkFingerTapCommand) => void;
  /** Shown for a moment where the fingers were (the command's own name). */
  readonly labels: Readonly<Record<InkFingerTapCommand, string>>;
}

interface TapSession {
  readonly target: InkFingerTapTarget;
  readonly startAt: number;
  readonly starts: Map<number, { x: number; y: number }>;
  fingers: number;
  spoiled: boolean;
  /** The rapid undo ran: the lift is not a tap as well. */
  held: boolean;
  timer: ReturnType<typeof setTimeout> | null;
}

const targets = new Set<InkFingerTapTarget>();
let session: TapSession | null = null;

function stopHold(current: TapSession): void {
  if (current.timer !== null) clearTimeout(current.timer);
  current.timer = null;
}

function spoil(current: TapSession): void {
  current.spoiled = true;
  stopHold(current);
}

function centre(current: TapSession): { x: number; y: number } {
  const points = [...current.starts.values()];
  return {
    x: points.reduce((sum, point) => sum + point.x, 0) / Math.max(1, points.length),
    y: points.reduce((sum, point) => sum + point.y, 0) / Math.max(1, points.length),
  };
}

function holdToUndo(current: TapSession): void {
  current.timer = setTimeout(
    () => {
      if (session !== current || current.spoiled || inkPointerArbiter.strokeRunning()) return;
      current.held = true;
      current.target.run("undo");
      const at = centre(current);
      flash(current.target.labels.undo, at.x, at.y);
      holdToUndo(current);
    },
    current.held ? INK_RAPID_UNDO_EVERY_MS : INK_RAPID_UNDO_DELAY_MS,
  );
}

/** A touch's contact size (CSS px across), where the browser reports one. */
function contactSize(touch: Touch): { width: number; height: number } {
  const own = touch as Touch & { webkitRadiusX?: number; webkitRadiusY?: number };
  return {
    width: Number(own.radiusX || own.webkitRadiusX || 0) * 2,
    height: Number(own.radiusY || own.webkitRadiusY || 0) * 2,
  };
}

function targetFor(node: EventTarget | null): InkFingerTapTarget | null {
  if (!(node instanceof Node)) return null;
  for (const target of targets) if (target.root.contains(node)) return target;
  return null;
}

/** The command a finished tap of `fingers` fingers stands for. */
export function inkFingerTapCommand(fingers: number): InkFingerTapCommand | null {
  if (fingers === 2) return "undo";
  if (fingers === 3) return "redo";
  return null;
}

function flash(label: string, x: number, y: number): void {
  const chip = document.createElement("div");
  chip.textContent = label;
  chip.setAttribute("data-ink-finger-tap", "");
  chip.setAttribute("aria-hidden", "true");
  Object.assign(chip.style, {
    position: "fixed",
    left: `${x}px`,
    top: `${y}px`,
    transform: "translate(-50%, -50%)",
    padding: "6px 14px",
    borderRadius: "999px",
    background: "rgba(28, 28, 30, 0.82)",
    color: "#fff",
    font: "600 13px/1.2 system-ui, -apple-system, sans-serif",
    pointerEvents: "none",
    zIndex: "2147483000",
  } satisfies Partial<CSSStyleDeclaration>);
  document.body.appendChild(chip);
  const reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
  const animation = chip.animate?.(
    reduced
      ? [{ opacity: 1 }, { opacity: 1, offset: 0.7 }, { opacity: 0 }]
      : [
          { opacity: 0, transform: "translate(-50%, -50%) scale(0.85)" },
          { opacity: 1, transform: "translate(-50%, -50%) scale(1)", offset: 0.18 },
          { opacity: 1, transform: "translate(-50%, -50%) scale(1)", offset: 0.7 },
          { opacity: 0, transform: "translate(-50%, -60%) scale(1)" },
        ],
    { duration: 650, easing: "ease-out" },
  );
  if (animation) animation.onfinish = () => chip.remove();
  else window.setTimeout(() => chip.remove(), 650);
}

function onTouchStart(event: TouchEvent): void {
  if (targets.size === 0) return;
  const now = event.timeStamp;
  if (!session) {
    // A tap starts with nothing else on the glass.
    if (event.touches.length !== event.changedTouches.length) return;
    const target = targetFor(event.target);
    if (!target) return;
    session = { target, startAt: now, starts: new Map(), fingers: 0, spoiled: false, held: false, timer: null };
  } else if (now - session.startAt > INK_TAP_LAND_MS) {
    spoil(session);
  }
  const current = session;
  for (const touch of Array.from(event.changedTouches)) {
    if (isStylusTouch(touch) || isHandSizedContact(contactSize(touch))) spoil(current);
    current.starts.set(touch.identifier, { x: touch.clientX, y: touch.clientY });
  }
  current.fingers = Math.max(current.fingers, event.touches.length);
  if (current.fingers > 3) spoil(current);
  stopHold(current);
  if (!current.spoiled && event.touches.length === 2) holdToUndo(current);
}

function onTouchMove(event: TouchEvent): void {
  if (!session || session.spoiled) return;
  for (const touch of Array.from(event.changedTouches)) {
    const start = session.starts.get(touch.identifier);
    if (start && Math.hypot(touch.clientX - start.x, touch.clientY - start.y) > INK_TAP_SLOP_PX) {
      spoil(session);
      return;
    }
  }
}

function onTouchEnd(event: TouchEvent): void {
  if (!session) return;
  // A finger lifting ends the hold; the others lifting after it are the same gesture.
  stopHold(session);
  if (event.type === "touchcancel") spoil(session);
  if (event.touches.length > 0) return;
  const done = session;
  session = null;
  if (done.spoiled || done.held || event.timeStamp - done.startAt > INK_TAP_MAX_MS) return;
  // The fingers were the hand under a pen.
  if (inkPointerArbiter.strokeRunning()) return;
  const command = inkFingerTapCommand(done.fingers);
  if (!command) return;
  const { x, y } = centre(done);
  done.target.run(command);
  flash(done.target.labels[command], x, y);
}

let attached = false;

/** Once, and for good: the window's first touch listeners (see above). */
function attach(): void {
  if (attached || typeof window === "undefined") return;
  attached = true;
  const options: AddEventListenerOptions = { capture: true, passive: true };
  window.addEventListener("touchstart", onTouchStart, options);
  window.addEventListener("touchmove", onTouchMove, options);
  window.addEventListener("touchend", onTouchEnd, options);
  window.addEventListener("touchcancel", onTouchEnd, options);
}

attach();

/** Listen for finger taps over `target.root`; returns the release. */
export function registerInkFingerTaps(target: InkFingerTapTarget): () => void {
  if (typeof window === "undefined") return () => {};
  attach();
  targets.add(target);
  return () => {
    targets.delete(target);
    if (session?.target === target) {
      stopHold(session);
      session = null;
    }
  };
}

/**
 * While `enabled` (an ink tool is picked), two fingers tapped on `root` (or
 * the element it returns when the effect runs; a function is read once per
 * `enabled` change, so pass a stable one) run
 * `run("undo")` and three run `run("redo")`. The newest `run` and labels are
 * used without re-registering.
 */
export function useInkFingerTaps(
  root: HTMLElement | null | (() => HTMLElement | null),
  enabled: boolean,
  run: (command: InkFingerTapCommand) => void,
  labels: Readonly<Record<InkFingerTapCommand, string>>,
): void {
  const latest = useRef({ run, labels });
  latest.current = { run, labels };
  useEffect(() => {
    const element = typeof root === "function" ? root() : root;
    if (!element || !enabled) return;
    return registerInkFingerTaps({
      root: element,
      run: (command) => latest.current.run(command),
      get labels() {
        return latest.current.labels;
      },
    });
  }, [root, enabled]);
}
