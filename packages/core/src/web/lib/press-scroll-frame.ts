/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A press on a zoomed-in canvas that scrolls under it: the canvas rolls when
 * the pointer is dragged to (or past) its edge, as PowerPoint's slide pane
 * does while an object is moved, a box is drawn or ink is laid down, and the
 * gesture keeps working in the place it began in.
 *
 * The editors' gestures measure the canvas once, when they are pressed (the
 * slide's box, its scale), and turn every later move into content
 * coordinates against that measurement. A canvas that scrolls during the
 * press would leave them a move behind: the object would slide off the
 * pointer by the distance scrolled. So, for the length of the press, every
 * move (and the release) the gesture hears is restated in the press's frame:
 * its client point is moved by the distance the canvas has scrolled since the
 * press, and a scroll on its own (the edge roll, a wheel) re-sends the last
 * move so the gesture follows at once. A gesture that measures the canvas
 * anew on every move (`live`) is sent the moves as they are.
 *
 * Framework-free; one install per scroll container.
 */

import { copilotAutoScrollStep } from "../__cut__/m9fe14da694";

export interface PressScrollFrameOptions {
  /** The scroll container (the canvas viewport). */
  scroller: () => HTMLElement | null;
  /** Whether a press starting with this pointerdown takes part. */
  accepts: (event: PointerEvent) => boolean;
  /** Whether this press's gesture measures the canvas again on every move. */
  live?: (event: PointerEvent) => boolean;
  /** Travel (screen px) before a press counts as a drag that may roll the canvas. */
  dragSlopPx?: number;
  /**
   * Runs a re-sent move so its update is drawn in the same frame as the
   * scroll that caused it (a React host passes `flushSync`): otherwise the
   * canvas rolled a frame before the object it carries followed.
   */
  flush?: (run: () => void) => void;
}

const DEFAULT_DRAG_SLOP_PX = 4;

interface Press {
  pointerId: number;
  scroller: HTMLElement;
  left0: number;
  top0: number;
  /** Screen px per CSS px of the scroller's content. */
  scale: number;
  live: boolean;
  startX: number;
  startY: number;
  dragging: boolean;
  last: PointerEvent;
  target: EventTarget | null;
  frame: number;
  /** The shift the gesture last heard a move at. */
  sentX: number;
  sentY: number;
}

/** Events this module sent: they pass through untouched. */
const restated = new WeakSet<Event>();

/** Whether an event is one this module restated (for a host's own listeners). */
export function isRestatedPressEvent(event: Event): boolean {
  return restated.has(event);
}

function restate(
  type: string,
  source: PointerEvent,
  dx: number,
  dy: number,
): PointerEvent {
  const init: PointerEventInit & Record<string, unknown> = {
    bubbles: true,
    cancelable: true,
    composed: true,
    clientX: source.clientX + dx,
    clientY: source.clientY + dy,
    screenX: source.screenX + dx,
    screenY: source.screenY + dy,
    ctrlKey: source.ctrlKey,
    shiftKey: source.shiftKey,
    altKey: source.altKey,
    metaKey: source.metaKey,
    button: source.button,
    buttons: source.buttons,
    detail: source.detail,
    pointerId: source.pointerId,
    pointerType: source.pointerType,
    isPrimary: source.isPrimary,
    pressure: source.pressure,
    tangentialPressure: source.tangentialPressure,
    tiltX: source.tiltX,
    tiltY: source.tiltY,
    twist: source.twist,
    width: source.width,
    height: source.height,
  };
  let event: PointerEvent;
  if (typeof PointerEvent === "function") {
    // The samples a fast stroke batches into one move, moved the same way.
    let coalesced: PointerEvent[] = [];
    try {
      coalesced =
        typeof source.getCoalescedEvents === "function"
          ? source.getCoalescedEvents().map((sample) =>
              new PointerEvent(type, {
                ...init,
                clientX: sample.clientX + dx,
                clientY: sample.clientY + dy,
                screenX: sample.screenX + dx,
                screenY: sample.screenY + dy,
                pressure: sample.pressure,
              }),
            )
          : [];
    } catch {
      coalesced = [];
    }
    try {
      event = new PointerEvent(type, {
        ...init,
        ...(coalesced.length > 0 ? { coalescedEvents: coalesced } : {}),
      } as PointerEventInit);
    } catch {
      event = new PointerEvent(type, init);
    }
  } else {
    // (No PointerEvent: a MouseEvent with its pointer fields.)
    event = new MouseEvent(type, init) as PointerEvent;
    for (const key of [
      "pointerId",
      "pointerType",
      "isPrimary",
      "pressure",
      "tiltX",
      "tiltY",
      "width",
      "height",
    ] as const) {
      Object.defineProperty(event, key, { value: (source as unknown as Record<string, unknown>)[key] });
    }
  }
  restated.add(event);
  return event;
}

export function installPressScrollFrame(
  options: PressScrollFrameOptions,
): () => void {
  if (typeof window === "undefined") return () => undefined;
  const slop = options.dragSlopPx ?? DEFAULT_DRAG_SLOP_PX;
  let press: Press | null = null;

  const shift = (current: Press) => ({
    dx: (current.scroller.scrollLeft - current.left0) * current.scale,
    dy: (current.scroller.scrollTop - current.top0) * current.scale,
  });

  const end = () => {
    const current = press;
    press = null;
    if (!current) return;
    if (current.frame) cancelAnimationFrame(current.frame);
    current.scroller.removeEventListener("scroll", onScroll);
  };

  /** The last move again, in the press's frame (the canvas scrolled under it). */
  const resend = () => {
    const current = press;
    if (!current || !current.dragging) return;
    const target = current.target;
    if (!target || (target instanceof Node && !target.isConnected)) return;
    const { dx, dy } = current.live ? { dx: 0, dy: 0 } : shift(current);
    const moved = shift(current);
    if (moved.dx === current.sentX && moved.dy === current.sentY) return;
    current.sentX = moved.dx;
    current.sentY = moved.dy;
    const event = restate("pointermove", current.last, dx, dy);
    if (options.flush) options.flush(() => target.dispatchEvent(event));
    else target.dispatchEvent(event);
  };

  function onScroll() {
    resend();
  }

  /**
   * The next frame of the edge roll, asked for once. A browser runs a frame
   * callback a frame later, so the roll is a loop of one call per frame that
   * stops when the press ends. A host that runs the callback at once while
   * it is being asked for (a test's requestAnimationFrame) recursed without
   * end and overflowed the stack: such a call does nothing, and the roll
   * stops there instead of looping.
   */
  let scheduling = false;
  const scheduleRoll = (current: Press) => {
    if (scheduling || current.frame || press !== current) return;
    scheduling = true;
    let id = 0;
    try {
      id = requestAnimationFrame(roll);
    } finally {
      scheduling = false;
    }
    if (press === current) current.frame = id;
  };

  /** The edge roll: while a drag is held at or past the canvas's edge. */
  function roll() {
    const current = press;
    if (!current || scheduling) return;
    current.frame = 0;
    // A canvas taken away mid-press (another slide, the view closed) ends it.
    if (!current.scroller.isConnected) {
      end();
      return;
    }
    rollStep(current);
    scheduleRoll(current);
  }

  const rollStep = (current: Press) => {
    if (current.dragging) {
      const scroller = current.scroller;
      const rect = scroller.getBoundingClientRect();
      const s = current.scale;
      const left = rect.left + scroller.clientLeft * s;
      const top = rect.top + scroller.clientTop * s;
      const right = left + scroller.clientWidth * s;
      const bottom = top + scroller.clientHeight * s;
      const x = current.last.clientX;
      const y = current.last.clientY;
      const stepX = copilotAutoScrollStep({ top: left, bottom: right }, x);
      const stepY = copilotAutoScrollStep({ top, bottom }, y);
      if (stepX) scroller.scrollLeft += stepX / s;
      if (stepY) scroller.scrollTop += stepY / s;
      // the gesture follows in this frame, not at the scroll event's
      if (stepX || stepY) resend();
    }
  };

  const onDown = (event: PointerEvent) => {
    if (restated.has(event)) return;
    if (press) {
      // a second finger: a pinch, not this press
      end();
      return;
    }
    const scroller = options.scroller();
    if (!scroller || !options.accepts(event)) return;
    const rect = scroller.getBoundingClientRect();
    const scale =
      scroller.offsetWidth > 0 && rect.width > 0 ? rect.width / scroller.offsetWidth : 1;
    press = {
      pointerId: event.pointerId,
      scroller,
      left0: scroller.scrollLeft,
      top0: scroller.scrollTop,
      scale,
      live: options.live?.(event) ?? false,
      startX: event.clientX,
      startY: event.clientY,
      dragging: false,
      last: event,
      target: event.target,
      frame: 0,
      sentX: 0,
      sentY: 0,
    };
    scroller.addEventListener("scroll", onScroll, { passive: true });
    scheduleRoll(press);
  };

  const onMoveOrEnd = (event: PointerEvent) => {
    const current = press;
    if (!current || restated.has(event) || event.pointerId !== current.pointerId) return;
    const ending = event.type !== "pointermove";
    if (event.type === "pointermove") {
      current.last = event;
      current.target = event.target;
      if (
        !current.dragging &&
        Math.hypot(event.clientX - current.startX, event.clientY - current.startY) > slop
      ) {
        current.dragging = true;
      }
    }
    // A release after a roll the gesture has not yet followed: the move to
    // where the release is comes first (a gesture commits what it last
    // previewed, not the release's point).
    if (ending && current.dragging) resend();
    const moved = shift(current);
    const { dx, dy } = current.live ? { dx: 0, dy: 0 } : moved;
    if (event.type === "pointermove") {
      current.sentX = moved.dx;
      current.sentY = moved.dy;
    }
    if (dx !== 0 || dy !== 0) {
      // The gesture hears this event in the press's frame, and only that.
      event.stopImmediatePropagation();
      const target = event.target ?? current.scroller;
      if (ending) end();
      target.dispatchEvent(restate(event.type, event, dx, dy));
      return;
    }
    if (ending) end();
  };

  const onBlur = () => end();

  window.addEventListener("pointerdown", onDown, true);
  window.addEventListener("pointermove", onMoveOrEnd, true);
  window.addEventListener("pointerup", onMoveOrEnd, true);
  window.addEventListener("pointercancel", onMoveOrEnd, true);
  window.addEventListener("blur", onBlur);
  return () => {
    end();
    window.removeEventListener("pointerdown", onDown, true);
    window.removeEventListener("pointermove", onMoveOrEnd, true);
    window.removeEventListener("pointerup", onMoveOrEnd, true);
    window.removeEventListener("pointercancel", onMoveOrEnd, true);
    window.removeEventListener("blur", onBlur);
  };
}
