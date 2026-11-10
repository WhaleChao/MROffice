"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { DOCUMENT_FLAGS, holdDocumentFlag } from "@mrnote/core/web/lib/document-flag.ts";
import { useEffect } from "react";

import {
  inkPalmGuard,
  inkPalmPointerKey,
  inkPalmTouchKey,
  isStylusTouch,
  type InkPalmContact,
} from "@mrnote/core/web/lib/ink-palm-guard.ts";
import {
  clearInkBlackBox,
  InkInputTrace,
  reportInkInputStall,
  reportLostInkPage,
  inkCancelsPenPointerDown,
  inkScribbleCancelSwitch,
  inkShieldEnabled,
  inkTraceNote,
  setActiveInkInputTrace,
  writeInkBlackBox,
} from "@mrnote/core/web/lib/ink-input-trace.ts";
import { desktopBridge } from "@mrnote/core/web/lib/open-files/desktop.ts";
import { inkPointerArbiter } from "@mrnote/core/web/lib/ink-pointer-policy.ts";
import { scribbleWritableNear } from "@mrnote/core/web/lib/ink-scribble-guard.ts";
// Its touch listeners attach first: the finger taps hear the fingers the
// shield stops as the hand (lib/ink-finger-taps).
import "./ink-finger-taps";
import { suspendTouchHoldTouchEnd, yieldTouchHold } from "@mrnote/core/web/lib/touch-hold.ts";









































/** A finger moves this far (px) before it pans. */
export const INK_FINGER_PAN_SLOP_PX = 10;
/** Down this long without moving: a resting hand, not a finger about to pan. */
export const INK_FINGER_PAN_REST_MS = 450;
/** A pan shorter than this (px), by a contact that turned out to be the hand, is put back. */
export const INK_FINGER_PAN_REVERT_PX = 32;
/**
 * The glide's decay (ms): iOS's own scroll view loses 0.2% of its speed a
 * millisecond (UIScrollView's normal deceleration rate, 0.998), so a flick
 * travels as far here as the same flick on a page outside the Draw tab.
 */
const GLIDE_TIME_CONSTANT_MS = -1 / Math.log(0.998);
/** Slower than this at the release (px/ms): the finger was placing the page, not flicking it. */
const GLIDE_START_SPEED = 0.12;
const GLIDE_STOP_SPEED = 0.01;
const GLIDE_MAX_SPEED = 8;
/** The finger's speed is read over its last this-many ms. */
const GLIDE_VELOCITY_WINDOW_MS = 100;
/**
 * A flick that lands on a glide still this fast (px/ms) and lets go within
 * {@link GLIDE_BOOST_MS} adds to it, as flicking a native list again does.
 */
const GLIDE_BOOST_MIN_SPEED = 0.3;
const GLIDE_BOOST_MS = 500;
/** A click this close (px, ms) to where a hand lifted is the hand's. */
const GHOST_CLICK_RADIUS_PX = 44;
const GHOST_CLICK_MS = 700;

const LOCK_ATTRIBUTE = "data-ink-finger-pan-lock";
const FREEZE_ATTRIBUTE = "data-ink-palm-freeze";
const WINDOW_SELECTOR = ".wm-window";

interface PanZone {
  /** An element inside the editor's scroll containers. */
  readonly anchor: HTMLElement;
  /** What a finger pans by dragging: the ink surface. */
  readonly surface: HTMLElement;
  locked: HTMLElement[];
}

/** A finger contact the browser may be panning with, and what it could move. */
interface NativePanWatch {
  readonly scrollers: Array<{ element: HTMLElement; left: number; top: number }>;
}

interface FingerPan {
  readonly pointerId: number;
  readonly key: string;
  readonly down: PointerEvent;
  readonly zone: PanZone;
  readonly startX: number;
  readonly startY: number;
  lastX: number;
  lastY: number;
  active: boolean;
  moved: number;
  starts: Map<HTMLElement, { left: number; top: number }>;
  samples: Array<{ at: number; x: number; y: number }>;
  /** The glide this finger landed on, when it was still moving fast. */
  readonly carried: { x: number; y: number } | null;
}

const zones = new Set<PanZone>();
const lockCounts = new WeakMap<HTMLElement, number>();
let installs = 0;
let detach: (() => void) | null = null;
let pan: FingerPan | null = null;
let glideFrame: number | null = null;
/** The running glide's velocity (client px/ms, a finger's direction). */
let glideVelocity: { x: number; y: number } | null = null;
let ghosts: Array<{ x: number; y: number; until: number }> = [];
let penDownAt = Number.NEGATIVE_INFINITY;
let penDownPoint: { x: number; y: number } | null = null;
const nativeWatches = new Map<string, NativePanWatch>();
const frozen = new Map<string, HTMLElement[]>();

function clips(value: string): boolean {
  return value !== "visible" && value !== "clip" && value !== "";
}

function userScrolls(value: string): boolean {
  return value === "auto" || value === "scroll" || value === "overlay";
}

/** The scroll containers around `anchor`, innermost first, up to its window. */
function scrollContainersAround(anchor: HTMLElement): HTMLElement[] {
  const edge = anchor.closest<HTMLElement>(WINDOW_SELECTOR);
  const found: HTMLElement[] = [];
  for (
    let element: HTMLElement | null = anchor;
    element && element !== document.body && element !== document.documentElement;
    element = element.parentElement
  ) {
    const style = getComputedStyle(element);
    if (clips(style.overflowX) || clips(style.overflowY)) found.push(element);
    if (element === edge) break;
  }
  return found;
}

function lock(element: HTMLElement): void {
  const count = lockCounts.get(element) ?? 0;
  lockCounts.set(element, count + 1);
  if (count === 0) element.setAttribute(LOCK_ATTRIBUTE, "");
}

function unlock(element: HTMLElement): void {
  const count = lockCounts.get(element) ?? 0;
  if (count <= 1) {
    lockCounts.delete(element);
    element.removeAttribute(LOCK_ATTRIBUTE);
    return;
  }
  lockCounts.set(element, count - 1);
}

/** Bring the zone's locks up to date with the containers around it now. */
function relock(zone: PanZone): void {
  const next = zone.anchor.isConnected ? scrollContainersAround(zone.anchor) : [];
  const same = next.length === zone.locked.length && next.every((element, index) => element === zone.locked[index]);
  if (same) return;
  for (const element of next) lock(element);
  for (const element of zone.locked) unlock(element);
  zone.locked = next;
}

/**
 * Where a pen inks and a finger pans: the scroll containers around `anchor`
 * stop panning natively and one finger dragged on `surface` (the ink surface;
 * `anchor` itself unless the surface is portaled out of those containers)
 * pans them from here.
 */
export function registerInkPanZone(anchor: HTMLElement, surface: HTMLElement = anchor): () => void {
  const zone: PanZone = { anchor, surface, locked: [] };
  relock(zone);
  zones.add(zone);
  const release = installInkTouchShield();
  let released = false;
  return () => {
    if (released) return;
    released = true;
    zones.delete(zone);
    if (pan?.zone === zone) pan = null;
    for (const element of zone.locked) unlock(element);
    zone.locked = [];
    release();
  };
}

/**
 * The zone a finger landed in, when what it landed on is the zone's to pan:
 * the ink surface, a scroll container's own background, or anything a host
 * marked `data-ink-pan-surface`. Controls and the editor's own objects keep
 * the press.
 */
function zoneFor(target: EventTarget | null): PanZone | null {
  if (!(target instanceof Element)) return null;
  for (const zone of zones) {
    relock(zone);
    if (zone.surface.contains(target)) return zone;
    const outer = zone.locked[zone.locked.length - 1];
    if (!outer?.contains(target)) continue;
    if (zone.locked.includes(target as HTMLElement)) return zone;
    const marked = target.closest("[data-ink-pan-surface]");
    if (marked && outer.contains(marked)) return zone;
  }
  return null;
}

function scaleOf(element: HTMLElement): { x: number; y: number } {
  const rect = element.getBoundingClientRect();
  const x = element.offsetWidth > 0 ? rect.width / element.offsetWidth : 1;
  const y = element.offsetHeight > 0 ? rect.height / element.offsetHeight : 1;
  return { x: x > 0 && Number.isFinite(x) ? x : 1, y: y > 0 && Number.isFinite(y) ? y : 1 };
}

/** The zone's containers a finger can move, innermost first. */
function pannable(zone: PanZone): HTMLElement[] {
  return zone.locked.filter((element) => {
    const style = getComputedStyle(element);
    return (
      (userScrolls(style.overflowY) && element.scrollHeight > element.clientHeight + 1) ||
      (userScrolls(style.overflowX) && element.scrollWidth > element.clientWidth + 1)
    );
  });
}

/**
 * Move the zone by a finger's travel (client px) on one axis: to the
 * innermost container with room that way; what none can take is dropped (the
 * desk behind the window stays put). Returns the travel taken (client px),
 * or null when every container is at its edge that way. A container with
 * room that moved less than asked (WebKit keeps whole pixels) still holds
 * the travel: a sub-pixel step never leaks to the container around it.
 */
function scrollAxisBy(containers: readonly HTMLElement[], axis: "x" | "y", delta: number): number | null {
  if (delta === 0) return 0;
  for (const element of containers) {
    const style = getComputedStyle(element);
    if (!userScrolls(axis === "x" ? style.overflowX : style.overflowY)) continue;
    const scale = axis === "x" ? scaleOf(element).x : scaleOf(element).y;
    const before = axis === "x" ? element.scrollLeft : element.scrollTop;
    if (axis === "x") element.scrollLeft = before - delta / scale;
    else element.scrollTop = before - delta / scale;
    const after = axis === "x" ? element.scrollLeft : element.scrollTop;
    if (after !== before) return (before - after) * scale;
    const range =
      axis === "x" ? element.scrollWidth - element.clientWidth : element.scrollHeight - element.clientHeight;
    // Right-to-left, scrollLeft runs from 0 down to -range.
    const low = axis === "x" && style.direction === "rtl" ? -range : 0;
    const room = delta > 0 ? before > low + 0.5 : before < low + range - 0.5;
    if (room) return 0;
  }
  return null;
}

/** Move the zone by a finger's travel (client px). True when anything moved. */
function scrollZoneBy(containers: readonly HTMLElement[], dx: number, dy: number): boolean {
  const x = scrollAxisBy(containers, "x", dx);
  const y = scrollAxisBy(containers, "y", dy);
  return (x ?? 0) !== 0 || (y ?? 0) !== 0;
}

function stopGlide(): void {
  if (glideFrame !== null) cancelAnimationFrame(glideFrame);
  glideFrame = null;
  glideVelocity = null;
}











function startGlide(containers: readonly HTMLElement[], vx: number, vy: number): void {
  stopGlide();
  const speed = Math.hypot(vx, vy);
  if (speed < GLIDE_START_SPEED || containers.length === 0) return;
  if (speed > GLIDE_MAX_SPEED) {
    vx *= GLIDE_MAX_SPEED / speed;
    vy *= GLIDE_MAX_SPEED / speed;
  }
  const velocity = { x: vx, y: vy };
  const owed = { x: 0, y: 0 };
  let last = performance.now();
  const step = (now: number) => {
    const dt = Math.min(64, Math.max(0, now - last));
    last = now;
    const decay = Math.exp(-dt / GLIDE_TIME_CONSTANT_MS);
    for (const axis of ["x", "y"] as const) {
      if (velocity[axis] === 0) continue;
      owed[axis] += velocity[axis] * GLIDE_TIME_CONSTANT_MS * (1 - decay);
      velocity[axis] *= decay;
      const taken = scrollAxisBy(containers, axis, owed[axis]);
      if (taken === null) {
        // At the edge this way: this axis is done.
        velocity[axis] = 0;
        owed[axis] = 0;
      } else {
        owed[axis] -= taken;
      }
    }
    if (Math.hypot(velocity.x, velocity.y) < GLIDE_STOP_SPEED) {
      glideFrame = null;
      glideVelocity = null;
      return;
    }
    glideFrame = requestAnimationFrame(step);
  };
  glideVelocity = velocity;
  glideFrame = requestAnimationFrame(step);
}

function beginPan(event: PointerEvent): void {
  const gliding = glideVelocity;
  const carried =
    gliding && Math.hypot(gliding.x, gliding.y) >= GLIDE_BOOST_MIN_SPEED ? { x: gliding.x, y: gliding.y } : null;
  stopGlide();
  // A second finger: a pinch's, not a pan's. The first finger's pan stands
  // where it is.
  if (pan || inkPalmGuard.fingersDown("p") > 1) {
    pan = null;
    return;
  }
  const zone = zoneFor(event.target);
  if (!zone) return;
  pan = {
    pointerId: event.pointerId,
    key: inkPalmPointerKey(event.pointerId),
    down: event,
    zone,
    startX: event.clientX,
    startY: event.clientY,
    lastX: event.clientX,
    lastY: event.clientY,
    active: false,
    moved: 0,
    starts: new Map(),
    samples: [],
    carried,
  };
}

function movePan(event: PointerEvent): void {
  const current = pan;
  if (!current || event.pointerId !== current.pointerId) return;
  // Something took the press for itself (an eraser on a finger, a grip).
  if (current.down.defaultPrevented) {
    pan = null;
    return;
  }
  if (!current.active) {
    if (Math.hypot(event.clientX - current.startX, event.clientY - current.startY) < INK_FINGER_PAN_SLOP_PX) {
      return;
    }
    if (event.timeStamp - current.down.timeStamp > INK_FINGER_PAN_REST_MS) {
      // It rested, then slid: a hand settling, for as long as it is down.
      pan = null;
      inkPalmGuard.markPalm(current.key);
      return;
    }
    current.active = true;
    for (const element of current.zone.locked) {
      current.starts.set(element, { left: element.scrollLeft, top: element.scrollTop });
    }
    yieldTouchHold(current.pointerId);
  } else {
    const dx = event.clientX - current.lastX;
    const dy = event.clientY - current.lastY;
    scrollZoneBy(pannable(current.zone), dx, dy);
    current.moved += Math.hypot(dx, dy);
  }
  current.lastX = event.clientX;
  current.lastY = event.clientY;
  current.samples.push({ at: event.timeStamp, x: event.clientX, y: event.clientY });
  if (current.samples.length > 8) current.samples.shift();
}

function endPan(event: PointerEvent, glide: boolean): void {
  const current = pan;
  if (!current || event.pointerId !== current.pointerId) return;
  pan = null;
  if (!current.active || !glide) return;
  const recent = current.samples.filter((sample) => event.timeStamp - sample.at <= GLIDE_VELOCITY_WINDOW_MS);
  const first = recent[0];
  const last = recent[recent.length - 1];
  if (!first || !last || last.at - first.at < 8) return;
  const span = last.at - first.at;
  let vx = (last.x - first.x) / span;
  let vy = (last.y - first.y) / span;
  // Flicked again while it still coasted: the flicks add up, each axis that
  // kept its direction.
  const carried = current.carried;
  if (
    carried &&
    Math.hypot(vx, vy) >= GLIDE_START_SPEED &&
    event.timeStamp - current.down.timeStamp <= GLIDE_BOOST_MS
  ) {
    if (Math.sign(vx) === Math.sign(carried.x)) vx += carried.x;
    if (Math.sign(vy) === Math.sign(carried.y)) vy += carried.y;
  }
  startGlide(pannable(current.zone), vx, vy);
}

/** A contact the page already knew turned out to be the hand. */
function onLatePalm(contacts: readonly InkPalmContact[]): void {
  for (const contact of contacts) {
    if (contact.key.startsWith("t")) {
      stopNativePan(contact.key);
      continue;
    }
    yieldTouchHold(Number(contact.key.slice(1)));
    const current = pan;
    if (!current || current.key !== contact.key) continue;
    pan = null;
    stopGlide();
    // A few pixels of drift before the pen arrived: back where it was. A
    // long pan was meant, by a finger that only happened to stay down.
    if (current.active && current.moved < INK_FINGER_PAN_REVERT_PX) {
      for (const [element, start] of current.starts) {
        element.scrollLeft = start.left;
        element.scrollTop = start.top;
      }
    }
  }
}

/**
 * What the browser could pan for a finger that landed outside every zone:
 * the scrolling ancestors of what it touched, as they stand.
 */
function watchNativePan(key: string, target: EventTarget | null): void {
  if (!(target instanceof Element)) return;
  const scrollers: NativePanWatch["scrollers"] = [];
  let nearest = true;
  for (let element: Element | null = target; element; element = element.parentElement) {
    if (!(element instanceof HTMLElement)) continue;
    const style = getComputedStyle(element);
    if (!clips(style.overflowX) && !clips(style.overflowY)) continue;
    // The nearest scroll container decides whether the browser pans at all:
    // inside a zone it does not, and the finger pan above owns the contact.
    if (nearest && element.hasAttribute(LOCK_ATTRIBUTE)) return;
    nearest = false;
    if (element.scrollHeight <= element.clientHeight + 1 && element.scrollWidth <= element.clientWidth + 1) continue;
    if (!userScrolls(style.overflowX) && !userScrolls(style.overflowY)) continue;
    scrollers.push({ element, left: element.scrollLeft, top: element.scrollTop });
  }
  if (scrollers.length > 0) nativeWatches.set(key, { scrollers });
}

function thaw(key: string): void {
  for (const element of frozen.get(key) ?? []) element.removeAttribute(FREEZE_ATTRIBUTE);
  frozen.delete(key);
}

/**
 * A hand that was panning natively before its pen arrived: the browser's pan
 * cannot be cancelled once begun, so what it moves stops scrolling until the
 * hand lifts, and a few pixels of drift go back where they were.
 */
function stopNativePan(key: string): void {
  const watch = nativeWatches.get(key);
  nativeWatches.delete(key);
  if (!watch) return;
  const moved = watch.scrollers.filter(
    ({ element, left, top }) => element.scrollLeft !== left || element.scrollTop !== top,
  );
  if (moved.length === 0) return;
  for (const { element, left, top } of moved) {
    element.setAttribute(FREEZE_ATTRIBUTE, "");
    if (Math.hypot(element.scrollLeft - left, element.scrollTop - top) < INK_FINGER_PAN_REVERT_PX) {
      element.scrollLeft = left;
      element.scrollTop = top;
    }
  }
  frozen.set(key, moved.map(({ element }) => element));
}

function rememberGhost(contact: InkPalmContact, at: number): void {
  ghosts = ghosts.filter((ghost) => ghost.until > at);
  ghosts.push({ x: contact.last.x, y: contact.last.y, until: at + GHOST_CLICK_MS });
}

function swallow(event: Event): void {
  trace?.swallowed(event);
  // A hand's touchstart on an ink surface goes on to the surface, which
  // cancels it (keepScribbleOff): stopped here it was answered "not
  // cancelled" while the Pencil's touchmove was queued -- WebKit's touch race,
  // its deferring recognizers (and Scribble with them) blocked for seconds
  
  // deferral", each ending a Scribble stall). The rest of the hand's events
  // still stop here.
  if (event.type === "touchstart" && heldScribbleSurfaceOf(event.target)) return;
  event.stopImmediatePropagation();
}

/** What the window heard while a pen wrote, for a page that stops answering. */
let trace: InkInputTrace | null = null;
const TRACED_EVENTS = [
  "pointerdown",
  "pointermove",
  "pointerup",
  "pointercancel",
  "touchstart",
  "touchmove",
  "touchend",
  "touchcancel",
  "click",
  "gesturestart",
] as const;

function traceState(): string {
  const palms = inkPalmGuard.contactsDown("t") - inkPalmGuard.fingersDown("t");
  const agent = typeof navigator === "undefined" ? "" : navigator.userAgent.slice(0, 160);
  return (
    `guard pen=${inkPalmGuard.penOnGlass()} palms=${palms} stroke=${inkPointerArbiter.strokeRunning()} ` +
    `penPD=${inkCancelsPenPointerDown() ? "cancel" : "pass"} ` +
    `visible=${document.visibilityState} ua=${agent}`
  );
}

/**
 * What took a pencil contact the page never heard: the elements under the pen
 * (their editability and touch-action), anything Scribble could write into
 * close by, the focus, the selection and the zoom. iPadOS takes a contact
 * before the page for Scribble, a text selection or a zoom gesture.
 */
function describePenPoint(x: number, y: number): string {
  const doc = document;
  const under = (typeof doc.elementsFromPoint === "function" ? doc.elementsFromPoint(x, y) : [])
    .slice(0, 5)
    .map((element) => {
      const style = getComputedStyle(element);
      const editable = (element as HTMLElement).isContentEditable ? " ce" : "";
      return `${element.tagName.toLowerCase()}.${String(element.className).slice(0, 28)}${editable} ta=${style.touchAction} us=${style.userSelect}`;
    })
    .join(" | ");
  const active = doc.activeElement;
  const focus = active ? `${active.tagName.toLowerCase()}.${String(active.className).slice(0, 24)}${(active as HTMLElement).isContentEditable ? " ce" : ""}` : "none";
  const selection = typeof getSelection === "function" ? getSelection()?.type ?? "none" : "none";
  const zoom = typeof visualViewport !== "undefined" && visualViewport ? visualViewport.scale.toFixed(2) : "?";
  return (
    `at ${Math.round(x)},${Math.round(y)}: ${under}\n` +
    `scribble ${scribbleWritableNear(doc, x, y, 120)} focus=${focus} selection=${selection} zoom=${zoom}`
  );
}











function lockViewportZoom(): () => void {
  if (typeof document === "undefined") return () => {};
  const meta = document.querySelector<HTMLMetaElement>('meta[name="viewport"]');
  if (!meta) return () => {};
  const was = meta.getAttribute("content") ?? "";
  const kept = was
    .split(",")
    .map((part) => part.trim())
    .filter((part) => part && !/^(maximum-scale|minimum-scale|user-scalable)\s*=/i.test(part));
  meta.setAttribute("content", [...kept, "maximum-scale=1", "user-scalable=no"].join(", "));
  return () => {
    if (meta.isConnected) meta.setAttribute("content", was);
  };
}

/** The iPad app's Scribble switch (lib/native-app's bridge), or null in a browser or an older build. */
function setAppInkArmed(armed: boolean): void {
  const bridge = desktopBridge() as unknown as { platform?: string; setInkArmed?: (armed: boolean) => Promise<unknown> } | null;
  if (bridge?.platform !== "ipados" || typeof bridge.setInkArmed !== "function") return;
  bridge.setInkArmed(armed).catch(() => {});
}

function attach(): () => void {
  // Nothing here cancels a touch any more (see the top), so it is CSS that
  // keeps iPadOS from taking a pen's drag over the page as a text selection
  
  // while an ink tool is picked the page's text is not selectable and no
  // callout opens on it (globals.css, html[data-ink-armed]).
  const lowerInkFlag = holdDocumentFlag(DOCUMENT_FLAGS.inkArmed);
  const unlockViewportZoom = lockViewportZoom();
  const resumeTouchHoldTouchEnd = suspendTouchHoldTouchEnd();
  // The iPad app takes Scribble off the page while a tool is picked, which no
  // page can (keepScribbleOff); a browser has no such bridge.
  setAppInkArmed(true);
  const unsubscribe = inkPalmGuard.subscribe(onLatePalm);
  // Registered before the shield's own listeners, so what they swallow is
  // still heard here.
  const recorder = (trace = new InkInputTrace(reportInkInputStall, traceState, undefined, describePenPoint));
  setActiveInkInputTrace(recorder);
  reportLostInkPage();
  const onAnyInput = (event: Event) => {
    recorder.heard(event);
    // After a lift, once the stroke has landed: counting the page's nodes
    // mid-stroke would cost the ink its frames.
    if (event.type === "pointerup" || event.type === "touchend") {
      setTimeout(() => {
        if (trace === recorder) writeInkBlackBox(recorder, traceState());
      }, 300);
    }
  };
  for (const type of TRACED_EVENTS) window.addEventListener(type, onAnyInput, { capture: true, passive: true });
  // Field diagnosis (inkdiag=1): how WebKit was answered for each touch. The
  // window's bubbling listener runs after the page's own, so it reads what
  // the page did: a cancelable touchstart left uncancelled is one WebKit's
  // touch race needs, and the first cancelable touchmove of a contact is the
  // preventable one it coalesces (WebKit PR #59255).
  const touchTag = (target: EventTarget | null) => {
    if (!(target instanceof Element)) return "?";
    const id = target.getAttribute("data-testid");
    const cls = typeof target.className === "string" ? target.className.split(/\s+/)[0] : "";
    return `${target.tagName.toLowerCase()}${id ? `#${id}` : cls ? `.${cls}` : ""}`.slice(0, 48);
  };
  const preventableMoves = new Set<number>();
  const onTouchAnswer = (event: Event) => {
    const touch = event as TouchEvent;
    const kinds = Array.from(touch.changedTouches ?? []).map((item) => (isStylusTouch(item) ? "p" : "f")).join("");
    if (event.type === "touchstart") {
      inkTraceNote(`T:s${kinds}${event.cancelable ? "c" : "u"}${event.defaultPrevented ? "P" : "N"} ${touchTag(event.target)}`);
      return;
    }
    if (event.type === "touchmove") {
      if (!event.cancelable) return;
      for (const item of Array.from(touch.changedTouches ?? [])) {
        if (preventableMoves.has(item.identifier)) continue;
        preventableMoves.add(item.identifier);
        inkTraceNote(`T:m${isStylusTouch(item) ? "p" : "f"}c${event.defaultPrevented ? "P" : "N"} ${touchTag(event.target)}`);
      }
      return;
    }
    for (const item of Array.from(touch.changedTouches ?? [])) preventableMoves.delete(item.identifier);
  };
  for (const type of ["touchstart", "touchmove", "touchend", "touchcancel"]) {
    window.addEventListener(type, onTouchAnswer, { passive: true });
  }

  const onPointerDown = (event: PointerEvent) => {
    if (event.pointerType === "pen") {
      stopGlide();
      penDownAt = event.timeStamp;
      penDownPoint = { x: event.clientX, y: event.clientY };
      inkPalmGuard.penContact(event.timeStamp);
      return;
    }
    if (event.pointerType !== "touch") return;
    const hand = inkPalmGuard.contactDown(
      inkPalmPointerKey(event.pointerId),
      event.timeStamp,
      { x: event.clientX, y: event.clientY },
      { width: event.width, height: event.height },
    );
    if (hand) {
      // The hand coming to rest stops the page under it.
      stopGlide();
      yieldTouchHold(event.pointerId);
      swallow(event);
      return;
    }
    beginPan(event);
  };
  const onPointerMove = (event: PointerEvent) => {
    if (event.pointerType === "pen") {
      // A move with the tip down that no pointerdown announced (the tool was
      // picked mid-stroke) is contact all the same.
      if (event.buttons !== 0 && event.pressure > 0) inkPalmGuard.penContact(event.timeStamp);
      else if (event.buttons === 0) inkPalmGuard.penHovered(event.timeStamp);
      else inkPalmGuard.penMoved(event.timeStamp);
      return;
    }
    if (event.pointerType !== "touch") return;
    const key = inkPalmPointerKey(event.pointerId);
    const point = { x: event.clientX, y: event.clientY };
    if (inkPalmGuard.isPalm(key)) {
      swallow(event);
      return;
    }
    inkPalmGuard.contactMoved(key, point);
    movePan(event);
  };
  const onPointerEnd = (event: PointerEvent) => {
    if (event.pointerType === "pen") {
      inkPalmGuard.penLifted(event.timeStamp);
      return;
    }
    if (event.pointerType !== "touch") return;
    const contact = inkPalmGuard.contactUp(inkPalmPointerKey(event.pointerId));
    if (contact?.palm) {
      rememberGhost(contact, event.timeStamp);
      if (pan?.pointerId === event.pointerId) pan = null;
      // A hand the page never heard of stays unheard; one it knew ends.
      if (!contact.late) swallow(event);
      return;
    }
    endPan(event, event.type === "pointerup");
  };
  const onPointerOut = (event: PointerEvent) => {
    if (event.pointerType === "pen" && event.relatedTarget === null && event.buttons === 0) {
      inkPalmGuard.penLeft(event.timeStamp);
    }
  };

  /** The pen's own contact, where touch events do not say what made them. */
  const isPenTouch = (touch: Touch, at: number) =>
    isStylusTouch(touch) ||
    (penDownPoint !== null &&
      Math.abs(at - penDownAt) < 100 &&
      Math.hypot(touch.clientX - penDownPoint.x, touch.clientY - penDownPoint.y) < 6);
  const penTouches = new Set<number>();
  /** True when the touch event's finger contacts are all the hand (and there is one). */
  const allHand = (event: TouchEvent): boolean => {
    let hands = 0;
    for (const touch of Array.from(event.touches)) {
      if (penTouches.has(touch.identifier) || isStylusTouch(touch)) continue;
      if (!inkPalmGuard.isPalm(inkPalmTouchKey(touch.identifier))) return false;
      hands += 1;
    }
    return hands > 0;
  };
  const hasPen = (event: TouchEvent) =>
    Array.from(event.touches).some((touch) => penTouches.has(touch.identifier) || isStylusTouch(touch));

  /**
   * A touch event lists every contact on the glass: those the guard still
   * holds that are not in it lifted where the page never heard (a cancelled
   * gesture, a hidden page), and a stale one would count toward a hand.
   */
  const forgetLostTouches = (event: TouchEvent) => {
    const live = new Set<string>();
    const ids = new Set<number>();
    for (const touch of Array.from(event.touches)) {
      live.add(inkPalmTouchKey(touch.identifier));
      ids.add(touch.identifier);
    }
    inkPalmGuard.forgetContactsExcept("t", live);
    for (const id of [...penTouches]) if (!ids.has(id)) penTouches.delete(id);
    // A scroller frozen under a lost contact would stay frozen for good.
    for (const key of [...frozen.keys()]) if (!live.has(key)) thaw(key);
    for (const key of [...nativeWatches.keys()]) if (!live.has(key)) nativeWatches.delete(key);
  };

  const onTouchStart = (event: TouchEvent) => {
    forgetLostTouches(event);
    for (const touch of Array.from(event.changedTouches)) {
      if (isPenTouch(touch, event.timeStamp)) {
        penTouches.add(touch.identifier);
        continue;
      }
      const key = inkPalmTouchKey(touch.identifier);
      const hand = inkPalmGuard.contactDown(
        key,
        event.timeStamp,
        { x: touch.clientX, y: touch.clientY },
        { width: (touch.radiusX || 0) * 2, height: (touch.radiusY || 0) * 2 },
      );
      if (!hand) watchNativePan(key, touch.target);
    }
    if (!allHand(event)) return;
    // The pen's own touchstart still reaches the surface's guards.
    if (!hasPen(event)) swallow(event);
  };
  const onTouchMove = (event: TouchEvent) => {
    for (const touch of Array.from(event.changedTouches)) {
      if (penTouches.has(touch.identifier)) continue;
      inkPalmGuard.contactMoved(inkPalmTouchKey(touch.identifier), { x: touch.clientX, y: touch.clientY });
    }
    if (!allHand(event)) return;
    if (!hasPen(event)) swallow(event);
  };
  const onTouchEnd = (event: TouchEvent) => {
    let hands = 0;
    let known = 0;
    let others = 0;
    let penEnded = false;
    for (const touch of Array.from(event.changedTouches)) {
      if (penTouches.delete(touch.identifier) || isStylusTouch(touch)) {
        penEnded = true;
        others += 1;
        continue;
      }
      const key = inkPalmTouchKey(touch.identifier);
      const contact = inkPalmGuard.contactUp(key);
      nativeWatches.delete(key);
      thaw(key);
      if (!contact?.palm) {
        others += 1;
        continue;
      }
      hands += 1;
      if (contact.late) known += 1;
      rememberGhost(contact, event.timeStamp);
    }
    // The Pencil's own touch ending is its lift, whether or not a pointerup
    // follows (a missed one kept every finger the hand).
    if (penEnded && !hasPen(event)) inkPalmGuard.penLifted(event.timeStamp);
    if (hands === 0 || others > 0) return;
    // (The click a hand's lift sends is dropped by `onClick`, by its place.)
    if (known === 0) swallow(event);
  };

  const onClick = (event: MouseEvent) => {
    const now = event.timeStamp;
    ghosts = ghosts.filter((ghost) => ghost.until > now);
    const pointerType = (event as PointerEvent).pointerType;
    if (pointerType === "pen" || pointerType === "mouse") return;
    if (
      !ghosts.some(
        (ghost) => Math.hypot(event.clientX - ghost.x, event.clientY - ghost.y) <= GHOST_CLICK_RADIUS_PX,
      )
    ) {
      return;
    }
    event.preventDefault();
    swallow(event);
  };

  







  const onGesture = (event: Event) => {
    const fingers = inkPalmGuard.fingersDown("t");
    const contacts = inkPalmGuard.contactsDown("t");
    const hands = contacts - fingers;
    const engaged = inkPalmGuard.penOnGlass(event.timeStamp);
    if (!engaged && !(hands > 0 && fingers < 2)) return;
    if (!engaged && contacts === 2 && !inkPointerArbiter.strokeRunning()) return;
    swallow(event);
  };

  /** The page stopped hearing the glass: no lift is coming for the pen. */
  const onBlur = () => {
    recorder.leaving("blur");
    penTouches.clear();
    inkPalmGuard.penLifted(performance.now());
  };
  const onHidden = () => {
    recorder.leaving("hidden");
    clearInkBlackBox();
    onBlur();
    stopGlide();
    pan = null;
    // Whatever was on the glass lifts on another page.
    inkPalmGuard.forgetContacts();
    nativeWatches.clear();
    for (const key of [...frozen.keys()]) thaw(key);
  };
  const onVisibility = () => {
    if (document.visibilityState === "hidden") onHidden();
  };

  const pointerOptions: AddEventListenerOptions = { capture: true, passive: true };
  // Touch and gesture listeners are passive (see the top): WebKit never
  // waits on this page before it lets a native gesture through, and nothing
  // here could cancel one. Clicks are not touches; theirs stay active.
  const activeOptions: AddEventListenerOptions = { capture: true, passive: false };
  window.addEventListener("pointerdown", onPointerDown, pointerOptions);
  window.addEventListener("pointermove", onPointerMove, pointerOptions);
  window.addEventListener("pointerup", onPointerEnd, pointerOptions);
  window.addEventListener("pointercancel", onPointerEnd, pointerOptions);
  window.addEventListener("pointerout", onPointerOut, pointerOptions);
  window.addEventListener("touchstart", onTouchStart, pointerOptions);
  window.addEventListener("touchmove", onTouchMove, pointerOptions);
  window.addEventListener("touchend", onTouchEnd, pointerOptions);
  window.addEventListener("touchcancel", onTouchEnd, pointerOptions);
  for (const type of ["click", "dblclick", "contextmenu", "auxclick"]) {
    window.addEventListener(type, onClick as EventListener, activeOptions);
  }
  for (const type of ["gesturestart", "gesturechange", "gestureend"]) {
    window.addEventListener(type, onGesture, pointerOptions);
  }
  window.addEventListener("blur", onBlur);
  window.addEventListener("pagehide", onHidden);
  document.addEventListener("visibilitychange", onVisibility);
  return () => {
    lowerInkFlag();
    unlockViewportZoom();
    setAppInkArmed(false);
    resumeTouchHoldTouchEnd();
    for (const type of TRACED_EVENTS) window.removeEventListener(type, onAnyInput, { capture: true });
    for (const type of ["touchstart", "touchmove", "touchend", "touchcancel"]) window.removeEventListener(type, onTouchAnswer);
    recorder.dispose();
    clearInkBlackBox();
    if (trace === recorder) {
      trace = null;
      setActiveInkInputTrace(null);
    }
    window.removeEventListener("blur", onBlur);
    window.removeEventListener("pagehide", onHidden);
    document.removeEventListener("visibilitychange", onVisibility);
    unsubscribe();
    stopGlide();
    pan = null;
    ghosts = [];
    penTouches.clear();
    nativeWatches.clear();
    for (const key of [...frozen.keys()]) thaw(key);
    window.removeEventListener("pointerdown", onPointerDown, pointerOptions);
    window.removeEventListener("pointermove", onPointerMove, pointerOptions);
    window.removeEventListener("pointerup", onPointerEnd, pointerOptions);
    window.removeEventListener("pointercancel", onPointerEnd, pointerOptions);
    window.removeEventListener("pointerout", onPointerOut, pointerOptions);
    window.removeEventListener("touchstart", onTouchStart, pointerOptions);
    window.removeEventListener("touchmove", onTouchMove, pointerOptions);
    window.removeEventListener("touchend", onTouchEnd, pointerOptions);
    window.removeEventListener("touchcancel", onTouchEnd, pointerOptions);
    for (const type of ["click", "dblclick", "contextmenu", "auxclick"]) {
      window.removeEventListener(type, onClick as EventListener, activeOptions);
    }
    for (const type of ["gesturestart", "gesturechange", "gestureend"]) {
      window.removeEventListener(type, onGesture, pointerOptions);
    }
    inkPalmGuard.reset();
  };
}

/**
 * The shield stands for as long as someone holds it (an ink surface while
 * its tool is picked). Returns the release.
 */
export function installInkTouchShield(): () => void {
  if (typeof window === "undefined") return () => {};
  installs += 1;
  if (installs === 1) detach = attach();
  let released = false;
  return () => {
    if (released) return;
    released = true;
    installs -= 1;
    if (installs > 0) return;
    detach?.();
    detach = null;
  };
}

/**
 * An ink surface's hold on the shield: while `anchor` is mounted the writing
 * hand is kept out of the page, and with `fingerPans` (a pen inks, a finger
 * pans: Draw with Touch off) the scroll containers around `anchor` are panned
 * from here instead of natively. `surface` is the ink surface when it is
 * portaled out of those containers.
 */
/** The Pencil contacts that start on an ink surface but belong to a control
 * on it (a lasso menu, a form field's button) keep their click. */
const SCRIBBLE_KEEPS_CONTROL =
  'button, a[href], input, select, textarea, label, [role="button"], [role="menuitem"], [role="menu"], [role="dialog"]';































/** The ink surfaces keepScribbleOff holds, each counted by its holders. */
const heldScribbleSurfaces = new Map<HTMLElement, number>();

function heldScribbleSurfaceOf(target: EventTarget | null): HTMLElement | null {
  if (!(target instanceof Node)) return null;
  for (const surface of heldScribbleSurfaces.keys()) if (surface.contains(target)) return surface;
  return null;
}

export function keepScribbleOff(surface: HTMLElement): () => void {
  if (inkScribbleCancelSwitch() === "off") return () => {};
  heldScribbleSurfaces.set(surface, (heldScribbleSurfaces.get(surface) ?? 0) + 1);
  const onTouchStart = (event: TouchEvent) => {
    const target = event.target instanceof Element ? event.target : null;
    const control = target?.closest(SCRIBBLE_KEEPS_CONTROL);
    if (control && surface.contains(control)) return;
    if (!event.cancelable) {
      // Answered already (the surface not yet in WebKit's touch regions):
      // the field diagnosis records each.
      inkTraceNote("S:uncancelable");
      return;
    }
    event.preventDefault();
  };
  // Safari's pinch and rotate gestures over the page, as Excalidraw and
  // tldraw cancel them: the editors zoom themselves.
  const doc = surface.ownerDocument;
  const onGesture = (event: Event) => event.preventDefault();
  surface.addEventListener("touchstart", onTouchStart, { passive: false });
  for (const type of GESTURE_EVENTS) doc.addEventListener(type, onGesture, { passive: false });
  return () => {
    surface.removeEventListener("touchstart", onTouchStart);
    for (const type of GESTURE_EVENTS) doc.removeEventListener(type, onGesture);
    const count = (heldScribbleSurfaces.get(surface) ?? 1) - 1;
    if (count > 0) heldScribbleSurfaces.set(surface, count);
    else heldScribbleSurfaces.delete(surface);
  };
}

const GESTURE_EVENTS = ["gesturestart", "gesturechange", "gestureend"] as const;

export function useInkTouchShield(
  anchor: HTMLElement | null,
  fingerPans: boolean,
  surface: HTMLElement | null = null,
): void {
  // `inkshield=off` (field comparison with Excalidraw): only the armed flag,
  // the viewport lock and the surface's touchstart cancel; no hand guard, no
  // window listeners, no finger pan.
  const shield = inkShieldEnabled();
  useEffect(() => {
    if (!anchor) return undefined;
    if (shield) return installInkTouchShield();
    const lowerInkFlag = holdDocumentFlag(DOCUMENT_FLAGS.inkArmed);
    const unlockViewportZoom = lockViewportZoom();
    return () => {
      lowerInkFlag();
      unlockViewportZoom();
    };
  }, [anchor, shield]);
  useEffect(() => (anchor ? keepScribbleOff(surface ?? anchor) : undefined), [anchor, surface]);
  useEffect(
    () => (anchor && fingerPans && shield ? registerInkPanZone(anchor, surface ?? anchor) : undefined),
    [anchor, fingerPans, surface, shield],
  );
}
