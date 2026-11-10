"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  describeOfficeScreenTip,
  hideOfficeScreenTip,
  showOfficeScreenTip,
  type OfficeScreenTipDescription,
} from "../components/office-ribbon/officeScreenTip";
import {
  isCompatMouseFromTouch,
  isTouchLikePointer,
  latestPointerIsTouchLike,
} from "./pointer-capability";

export { isCompatMouseFromTouch };

/*
 * Press and hold — the touch screen's right-click.
 *
 * Microsoft's touch conventions: "press and hold to access a menu of
 * secondary, contextual commands", and press and hold shows info "without
 * committing to an action or command" (Windows touch interactions; Office
 * Touch Guide: pressing and holding shows more commands, like right-clicking).
 * Chrome on Android/Windows turns a long press into a `contextmenu` event;
 * Safari on an iPad never does, so every right-click menu in the product was
 * unreachable there. One set of window listeners gives every surface the
 * gesture without wiring a call site:
 *
 * - A finger (or a pen on a device that cannot hover) held still for
 *   TOUCH_HOLD_MS — not moved more than TOUCH_HOLD_SLOP_PX, no second finger,
 *   no scroll, no pointercancel — dispatches a `contextmenu` MouseEvent at the
 *   press point on the pressed element. The surface's own onContextMenu opens
 *   its menu there, exactly as for a right-click.
 * - When nothing takes that event (no menu here) and the press is on a
 *   control the editor's chrome tips on hover (OfficeScreenTipProvider, the
 *   Spreadsheet ribbon), the hold shows that same ScreenTip card — name,
 *   shortcut, what it does, and for a greyed control why it cannot act:
 *   "press and hold to learn" (Windows touch), since a finger cannot hover.
 *   Elsewhere a control with a name its face does not show (an icon
 *   button's aria-label or title) gets a card with just that name.
 * - A greyed control is held through what is around it: Safari sends no
 *   pointer events to a disabled form control, so on a device that cannot
 *   hover the chrome's disabled controls let the press through to their
 *   wrapper (OfficeScreenTip.module.css), the hold finds the greyed control
 *   under the finger, and a tap there still does nothing.
 * - Either way the press is used up: the click (and the compatibility mouse
 *   events) a release would send are suppressed, so holding never also
 *   presses the thing held.
 * - Text keeps the platform's long press (select a word, the system's
 *   Copy/Paste callout), as Office on an iPad does: in a document's body
 *   always, in a field while it is being typed in; a hold on text that is
 *   already selected opens the menu for that selection. A field that is not
 *   being typed in (a slide's text box, a sheet cell at rest) is an object
 *   like any other: a hold opens its menu.
 * - Text the platform owns although nobody types in it (a PDF's text
 *   layer, any subtree marked `data-touch-hold="text"`) keeps the platform's
 *   long press too: hold to select a word, drag its selection handles. A
 *   hold on what is already selected there opens the menu for it.
 * - Ink and drawing surfaces, and any subtree marked
 *   `data-touch-hold="off"`, keep the press as it is; a subtree marked
 *   `data-touch-hold-pen="off"` keeps a pen's press (a finger still holds).
 * - An object that moves under a held finger (`data-touch-hold="drag"`) is
 *   held in two stages, as Office's objects are: still for
 *   TOUCH_DRAG_HOLD_MS lifts it (its own gesture, see
 *   lib/object-drag-gesture.ts), a move after that drags it and is not a
 *   hold, and a release without moving opens its menu then. The menu never
 *   pops up under a finger that is about to drag.
 * - A surface that captures the pointer for a gesture of its own (a slider,
 *   a grip, a drawing tool) has taken the press: the hold stands down
 *   (`gotpointercapture` on another element than the one pressed, or an
 *   explicit `yieldTouchHold`).
 * - A mouse, an iPad trackpad and a pen on a hovering device are never
 *   touched: their right-click is the platform's.
 */

export const TOUCH_HOLD_MS = 480;
export const TOUCH_HOLD_SLOP_PX = 10;
/** How long a finger stays still on a movable object before it lifts. */
export const TOUCH_DRAG_HOLD_MS = 300;
/** Travel (px) after which a press on a movable object is a drag or a scroll. */
export const TOUCH_DRAG_SLOP_PX = 8;
/** How long after a used-up hold its release's click is swallowed (a
 * release's click follows it at once; the next press ends it anyway). */
const RELEASE_SUPPRESS_MS = 450;
/** A platform contextmenu this soon after a used-up hold is its duplicate. */
const DUPLICATE_MENU_MS = 1500;
/** A platform contextmenu this soon after a tap is the tap's, not a menu. */
const TAP_MENU_MS = 600;
const TIP_OWNER = "touch-hold";
const TIP_TEST_ID = "touch-hold-screen-tip";
const TIP_HIDE_MS = 4000;

/** Subtrees where a held finger stays the surface's own (ink, drawing). */
export const TOUCH_HOLD_OPT_OUT_SELECTOR = [
  '[data-touch-hold="off"]',
  '[data-testid="document-draw-surface"]',
  '[data-testid="spreadsheet-draw-surface"]',
  '[data-testid="pdf-ink-overlay"]',
  '[data-testid="presentation-record-studio-ink-surface"]',
  '[data-testid="document-picture-draw-canvas"]',
  '[data-table-draw-armed="true"]',
  "[data-table-draw-mode]",
].join(", ");
const OPT_OUT_SELECTOR = TOUCH_HOLD_OPT_OUT_SELECTOR;







export const TOUCH_HOLD_PEN_OFF_ATTRIBUTE = "data-touch-hold-pen";
const PEN_OPT_OUT_SELECTOR = `[${TOUCH_HOLD_PEN_OFF_ATTRIBUTE}="off"]`;

/** Movable objects: hold lifts, a move drags, a still release opens the menu. */
const DRAG_SELECTOR = '[data-touch-hold="drag"]';

/** Text that keeps the platform's long press although it is not editable. */
const PLATFORM_TEXT_SELECTOR = '[data-pdf-text-layer], [data-touch-hold="text"]';

/** Chrome whose greyed controls let a finger's press reach their wrapper
 * (the same scope as the CSS rule in OfficeScreenTip.module.css). */
const DISABLED_REACH_SELECTOR = [
  "[data-office-screen-tips]",
  "[data-office-authoring-ribbon]",
  ".note-editor-ribbon-strip",
].join(", ");
const DISABLED_CONTROL_SELECTOR = "button:disabled, input:disabled, select:disabled";

/** Controls whose name a hold may show. */
const NAMED_CONTROL_SELECTOR = [
  "button",
  "a[href]",
  '[role="button"]',
  '[role="menuitem"]',
  '[role="tab"]',
  '[role="switch"]',
  '[role="checkbox"]',
  '[role="radio"]',
  '[role="option"]',
].join(", ");

interface Press {
  pointerId: number;
  x: number;
  y: number;
  screenX: number;
  screenY: number;
  target: Element;
  editable: boolean;
  timer: ReturnType<typeof setTimeout> | null;
  fired: boolean;
  /** A movable object's press (`data-touch-hold="drag"`). */
  drag: boolean;
  /** A drag press held still long enough that its release is a hold. */
  lifted: boolean;
  /** Pressed on the current text selection: its menu acts on it. */
  onSelectedText: boolean;
}

let press: Press | null = null;
/** Until when a used-up hold's release click / compat events are swallowed. */
let suppressUntil = 0;
/** When a hold was last used up (a menu opened, or a name shown). */
let lastUsedUpAt = 0;
/** The last touch press's down and up, to tell a tap's platform menu. */
let lastTouchDownAt = Number.NEGATIVE_INFINITY;
let lastTouchUpAt = Number.NEGATIVE_INFINITY;
let tipTimer: ReturnType<typeof setTimeout> | null = null;
let installed = false;

function now(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

/** A field, or text the user can type into, at this element. */
export function isTouchHoldEditable(element: Element): boolean {
  const field = element.closest("input, textarea, select");
  if (field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement) {
    return !field.readOnly && !field.disabled;
  }
  if (field instanceof HTMLSelectElement) return !field.disabled;
  const html = element instanceof HTMLElement ? element : element.parentElement;
  if (html && typeof html.isContentEditable === "boolean") {
    return html.isContentEditable;
  }
  // no isContentEditable (a DOM without it): the nearest contenteditable says
  const host = element.closest("[contenteditable]");
  return Boolean(host && host.getAttribute("contenteditable") !== "false");
}

/** The field or editable root the user would be typing into here. */
function editingHost(element: Element): Element | null {
  return (
    element.closest("input, textarea, select") ??
    element.closest('[contenteditable]:not([contenteditable="false"])')
  );
}

/** That host has the focus: the user is editing there. */
function isBeingEdited(host: Element | null): boolean {
  const active = document.activeElement;
  return Boolean(host && active && (host === active || host.contains(active)));
}

/** A focused field has a non-empty selection. */
function fieldHasSelection(host: Element | null): boolean {
  if (!(host instanceof HTMLInputElement || host instanceof HTMLTextAreaElement)) {
    return false;
  }
  try {
    return host.selectionStart !== null && host.selectionStart !== host.selectionEnd;
  } catch {
    return false; // an input type with no selection (number, email …)
  }
}

/** The press point lies on the current (non-empty) text selection. */
function onSelection(x: number, y: number): boolean {
  const selection = typeof window !== "undefined" ? window.getSelection() : null;
  if (!selection || selection.isCollapsed || selection.rangeCount === 0) return false;
  for (let index = 0; index < selection.rangeCount; index += 1) {
    for (const rect of Array.from(selection.getRangeAt(index).getClientRects())) {
      if (x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom) {
        return true;
      }
    }
  }
  return false;
}

/** The name a hold shows for the control under it, or null. */
export function touchHoldControlName(element: Element): {
  control: HTMLElement;
  name: string;
} | null {
  const control = element.closest(NAMED_CONTROL_SELECTOR);
  if (!(control instanceof HTMLElement)) return null;
  const title = control.getAttribute("title")?.replace(/\s+/g, " ").trim() ?? "";
  const label = control.getAttribute("aria-label")?.replace(/\s+/g, " ").trim() ?? "";
  const face = (control.textContent ?? "").replace(/\s+/g, " ").trim();
  // a control whose face already says its name has nothing to add
  if (title && title !== face) return { control, name: title };
  if (label && !face) return { control, name: label };
  return null;
}

/**
 * The greyed control under (x, y) when the press landed on what wraps it
 * (it takes no pointer events there), or null.
 */
export function touchHoldDisabledControlAt(
  element: Element,
  x: number,
  y: number,
): HTMLElement | null {
  if (!element.closest(DISABLED_REACH_SELECTOR)) return null;
  if (element.closest(NAMED_CONTROL_SELECTOR)) return null; // a control itself
  for (const control of element.querySelectorAll(DISABLED_CONTROL_SELECTOR)) {
    if (!(control instanceof HTMLElement)) continue;
    const box = control.getBoundingClientRect();
    if (box.width <= 0 || box.height <= 0) continue;
    if (x >= box.left && x <= box.right && y >= box.top && y <= box.bottom) {
      return control;
    }
  }
  return null;
}

function hideTip() {
  if (tipTimer !== null) clearTimeout(tipTimer);
  tipTimer = null;
  hideOfficeScreenTip(TIP_OWNER);
}

function release() {
  document.documentElement.removeAttribute("data-touch-holding");
}

/** While the finger stays down the platform must not start its own long
 * press (callout, magnifier, a selection) over what the hold opened. */
function holdOn(editable: boolean) {
  document.documentElement.setAttribute(
    "data-touch-holding",
    editable ? "text" : "surface",
  );
}

function cancel() {
  if (press?.timer) clearTimeout(press.timer);
  press = null;
}

function spendHold(editable: boolean) {
  suppressUntil = Number.POSITIVE_INFINITY; // until the finger lifts
  lastUsedUpAt = now();
  holdOn(editable);
}

function fire() {
  const current = press;
  if (!current) return;
  current.timer = null;
  current.fired = true;
  const target = current.target.isConnected
    ? current.target
    : document.elementFromPoint(current.x, current.y);
  if (!target) return;
  const event = new MouseEvent("contextmenu", {
    bubbles: true,
    cancelable: true,
    composed: true,
    clientX: current.x,
    clientY: current.y,
    screenX: current.screenX,
    screenY: current.screenY,
    button: 2,
    buttons: 0,
  });
  const handled = !target.dispatchEvent(event);
  if (handled) {
    spendHold(current.editable || current.onSelectedText);
    // a non-editable surface's native long-press may have begun selecting
    // (a hold on a selection opens the menu FOR it: it stays selected)
    if (!current.editable && !current.onSelectedText) {
      window.getSelection()?.removeAllRanges();
    }
    return;
  }
  const held = touchHoldDisabledControlAt(target, current.x, current.y) ?? target;
  // the card a hover would show (description, greyed reason), else the name
  const named = touchHoldControlName(held);
  const described: OfficeScreenTipDescription | null =
    describeOfficeScreenTip(held) ?? (named ? { element: named.control } : null);
  if (!described) return;
  spendHold(current.editable || current.onSelectedText);
  hideTip();
  const { splitDisabledReason, ...card } = described;
  showOfficeScreenTip(card, 0, {
    testId: TIP_TEST_ID,
    owner: TIP_OWNER,
    splitDisabledReason,
  });
  tipTimer = setTimeout(hideTip, TIP_HIDE_MS);
}

/**
 * A gesture has taken this pointer (or the current press, with no id): the
 * hold stands down — no menu, no name — and its release is left alone.
 */
export function yieldTouchHold(pointerId?: number): void {
  if (!press) return;
  if (pointerId !== undefined && press.pointerId !== pointerId) return;
  cancel();
}

function onPointerDown(event: PointerEvent) {
  hideTip();
  if (isTouchLikePointer(event) && event.isPrimary) lastTouchDownAt = now();
  // a new press: the last release's click, if any, has come and gone
  if (suppressUntil !== Number.POSITIVE_INFINITY) suppressUntil = 0;
  if (press && press.pointerId !== event.pointerId) {
    // a second finger: a pinch or a two-finger gesture, not a hold
    cancel();
    return;
  }
  if (!isTouchLikePointer(event) || !event.isPrimary) return;
  if (event.button > 0) return;
  const target = event.target instanceof Element ? event.target : null;
  if (!target || target.closest(OPT_OUT_SELECTOR)) return;
  if (event.pointerType === "pen" && target.closest(PEN_OPT_OUT_SELECTOR)) return;
  // a PDF's text: hold to select (the platform's), unless already selected
  if (
    target.closest(PLATFORM_TEXT_SELECTOR) &&
    !onSelection(event.clientX, event.clientY)
  ) {
    return;
  }
  const drag = Boolean(target.closest(DRAG_SELECTOR));
  const editable = isTouchHoldEditable(target);
  if (editable) {
    if (target.closest("select")) return;
    // text keeps the platform's long press unless the press is on what is
    // already selected there; a field at rest is held like an object
    const host = editingHost(target);
    const field = host?.matches("input, textarea") ? host : null;
    const platformPress = field
      ? isBeingEdited(field) && !fieldHasSelection(field)
      : !onSelection(event.clientX, event.clientY);
    if (platformPress) return;
  }
  cancel();
  const next: Press = {
    pointerId: event.pointerId,
    x: event.clientX,
    y: event.clientY,
    screenX: event.screenX,
    screenY: event.screenY,
    target,
    editable,
    timer: null,
    fired: false,
    drag,
    lifted: false,
    onSelectedText: onSelection(event.clientX, event.clientY),
  };
  next.timer = drag
    ? setTimeout(() => {
        next.timer = null;
        next.lifted = true;
      }, TOUCH_DRAG_HOLD_MS)
    : setTimeout(fire, TOUCH_HOLD_MS);
  press = next;
}

function onPointerMove(event: PointerEvent) {
  if (!press || press.fired || event.pointerId !== press.pointerId) return;
  if (
    Math.hypot(event.clientX - press.x, event.clientY - press.y) >
    (press.drag ? TOUCH_DRAG_SLOP_PX : TOUCH_HOLD_SLOP_PX)
  ) {
    // a scroll, or (after the lift) the object's drag: not a hold
    cancel();
  }
}

function onGotPointerCapture(event: PointerEvent) {
  if (!press || press.fired || press.drag || event.pointerId !== press.pointerId) {
    return;
  }
  // a touch is implicitly captured by what it pressed; any other element
  // capturing it is a gesture taking the press
  if (event.target !== press.target) cancel();
}

function onPointerEnd(event: PointerEvent) {
  if (event.type === "pointerup" && isTouchLikePointer(event) && event.isPrimary) {
    lastTouchUpAt = now();
  }
  if (!press || event.pointerId !== press.pointerId) return;
  // a movable object held still and let go: its menu, at the release
  if (event.type === "pointerup" && press.drag && press.lifted && !press.fired) {
    fire();
  }
  cancel();
  release();
  if (suppressUntil === Number.POSITIVE_INFINITY) {
    suppressUntil = now() + RELEASE_SUPPRESS_MS;
  }
}

function onScroll() {
  if (press && !press.fired) cancel();
  if (tipTimer !== null) hideTip();
}

function onContextMenu(event: MouseEvent) {
  // our own (untrusted), a mouse's or a trackpad's right-click: as they are
  if (!event.isTrusted || !latestPointerIsTouchLike()) return;
  if (press && !press.fired && press.drag) {
    // a movable object's menu waits for the release (the finger may be
    // about to drag it); the platform's mid-press menu is not shown
    event.preventDefault();
    event.stopImmediatePropagation();
    return;
  }
  if (press && !press.fired) {
    // the platform's own long press (Chrome) came first: it is the menu
    if (press.timer) clearTimeout(press.timer);
    press.timer = null;
    press.fired = true;
    spendHold(press.editable);
    return;
  }
  // the platform's long press after ours (Chrome on Windows fires it on
  // release): the menu is already open
  if (now() - lastUsedUpAt < DUPLICATE_MENU_MS) {
    event.preventDefault();
    event.stopImmediatePropagation();
    return;
  }
  // a tap, not a hold: the finger is already up after a press shorter than
  // a hold. Chrome sends this when the tap lands on text that the tap's own
  // focus has just selected (a sheet cell selects its entry), as its menu
  // for that selection; it is no request for the surface's menu, which then
  // opened on every tap of a filled cell, under the next tap.
  if (
    !press &&
    lastTouchUpAt >= lastTouchDownAt &&
    lastTouchUpAt - lastTouchDownAt < TOUCH_HOLD_MS &&
    now() - lastTouchUpAt < TAP_MENU_MS
  ) {
    event.preventDefault();
    event.stopImmediatePropagation();
  }
}

/** The release of a used-up hold: no click, no compatibility mouse events.
 * And a tap on a greyed control (which passes the press to its wrapper)
 * presses nothing, as on the control itself. */
function swallow(event: Event) {
  if (
    event.type === "click" &&
    event instanceof MouseEvent &&
    event.target instanceof Element &&
    latestPointerIsTouchLike() &&
    touchHoldDisabledControlAt(event.target, event.clientX, event.clientY)
  ) {
    event.preventDefault();
    event.stopImmediatePropagation();
    return;
  }
  if (suppressUntil <= now()) return;
  event.preventDefault();
  event.stopImmediatePropagation();
  if (event.type === "click") suppressUntil = 0;
}

function onTouchEnd(event: TouchEvent) {
  if (suppressUntil > now() && event.cancelable) event.preventDefault();
}

function noop() {}

let touchEndHeld = 0;







export function suspendTouchHoldTouchEnd(): () => void {
  if (typeof window === "undefined") return () => {};
  touchEndHeld += 1;
  if (touchEndHeld === 1 && installed) window.removeEventListener("touchend", onTouchEnd, { capture: true });
  let released = false;
  return () => {
    if (released) return;
    released = true;
    touchEndHeld -= 1;
    if (touchEndHeld === 0 && installed) window.addEventListener("touchend", onTouchEnd, { capture: true, passive: false });
  };
}

/** Install the window listeners (once per page; safe to call repeatedly). */
export function installTouchHold(): void {
  if (installed || typeof window === "undefined") return;
  installed = true;
  const passive = { capture: true, passive: true } as const;
  window.addEventListener("pointerdown", onPointerDown, passive);
  window.addEventListener("pointermove", onPointerMove, passive);
  window.addEventListener("pointerup", onPointerEnd, passive);
  window.addEventListener("pointercancel", onPointerEnd, passive);
  window.addEventListener("gotpointercapture", onGotPointerCapture, passive);
  window.addEventListener("scroll", onScroll, passive);
  window.addEventListener("contextmenu", onContextMenu, true);
  // iOS Safari lets `:active` match only under a touchstart listener: with
  // one on the document a pressed control shows its pressed feedback (the
  // `:active` copies postcss-touch-hover.cjs writes for a touch screen).
  document.addEventListener("touchstart", noop, passive);
  if (touchEndHeld === 0) window.addEventListener("touchend", onTouchEnd, { capture: true, passive: false });
  for (const type of ["mousedown", "mouseup", "click"]) {
    window.addEventListener(type, swallow, true);
  }
}
