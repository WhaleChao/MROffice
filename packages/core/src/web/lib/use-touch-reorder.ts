"use client";






































/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useLayoutEffect, useRef, type RefObject } from "react";
import { screenScaleOf } from "./css-zoom";

import {
  beginObjectDragGesture,
  installObjectDragTouchGuard,
  objectDragIsTouch,
  type ObjectDragGesture,
} from "./object-drag-gesture";

export type TouchReorderOrientation = "vertical" | "horizontal" | "grid";
export type TouchReorderPosition = "before" | "after";

export interface TouchReorderMove {
  /** The items moving, in list order (the pressed one, or its group). */
  readonly ids: readonly string[];
  /** The item they land next to. */
  readonly targetId: string;
  readonly position: TouchReorderPosition;
  readonly source: "touch" | "mouse" | "keyboard";
}

export interface TouchReorderOptions {
  /** The attribute that carries each item's id (e.g. `data-slide-id`). */
  idAttribute: string;
  /** Extra selector the items must also match (default: any element). */
  itemSelector?: string;
  /** Which way the list runs (default vertical). A grid reads left/right. */
  orientation?: TouchReorderOrientation;
  /** The element that scrolls (default: the container). */
  scrollContainer?: () => HTMLElement | null;
  /** What moves when `id` is lifted (default: just `id`). */
  movingIds?: (id: string) => readonly string[];
  /** May this item be lifted (read-only, a locked item …)? Default yes. */
  canLift?: (id: string) => boolean;
  onReorder: (move: TouchReorderMove) => void;
  /** A lift began (the host may select the item). */
  onLift?: (id: string) => void;
  /** Ctrl/⌘+Arrow moves the focused item (default true). */
  keyboard?: boolean;
  /**
   * The mouse reorders through this pointer path too, lifting once it travels
   * past the click slop (a still click stays a click). For a list whose HTML5
   * `draggable` is off: Safari's native drag session could stay attached to
   * the pointer when the list re-rendered under the press (a slide pressed
   * while a text box was in edit), so the tile floated beside the pointer
   * and the next press landed on another slide.
   */
  mouse?: boolean;
  /** Set false to attach nothing. */
  enabled?: boolean;
}

/** Distance (px) from the scrolling edge where auto-scroll begins. */
export const TOUCH_REORDER_EDGE_PX = 48;
/** Auto-scroll speed at the very edge, px per frame. */
export const TOUCH_REORDER_MAX_SCROLL_PX = 18;
const INDICATOR_TEST_ID = "touch-reorder-indicator";

/**
 * The new order after moving `moving` next to `targetId` (`before` / `after`
 * it). The moving block keeps its relative order. A target that is itself
 * moving, or unknown, leaves the order as it was.
 */
export function reorderIds(
  order: readonly string[],
  moving: readonly string[],
  targetId: string,
  position: TouchReorderPosition,
): string[] {
  const movingSet = new Set(moving.filter((id) => order.includes(id)));
  if (movingSet.size === 0 || movingSet.has(targetId) || !order.includes(targetId)) {
    return [...order];
  }
  const block = order.filter((id) => movingSet.has(id));
  const rest = order.filter((id) => !movingSet.has(id));
  const at = rest.indexOf(targetId) + (position === "after" ? 1 : 0);
  return [...rest.slice(0, at), ...block, ...rest.slice(at)];
}

/** Where a point drops among item rects: the nearest item and its side. */
export function touchReorderDropAt(
  items: ReadonlyArray<{ id: string; rect: Pick<DOMRect, "left" | "top" | "right" | "bottom"> }>,
  x: number,
  y: number,
  orientation: TouchReorderOrientation,
): { id: string; position: TouchReorderPosition } | null {
  let best: { id: string; position: TouchReorderPosition; distance: number } | null = null;
  for (const { id, rect } of items) {
    const dx = x < rect.left ? rect.left - x : x > rect.right ? x - rect.right : 0;
    const dy = y < rect.top ? rect.top - y : y > rect.bottom ? y - rect.bottom : 0;
    // a list reads along its axis: the cross axis barely matters
    const distance =
      orientation === "vertical"
        ? dy * 4 + dx
        : orientation === "horizontal"
          ? dx * 4 + dy
          : Math.hypot(dx, dy * 2);
    if (best && distance >= best.distance) continue;
    const centerX = (rect.left + rect.right) / 2;
    const centerY = (rect.top + rect.bottom) / 2;
    const position: TouchReorderPosition =
      orientation === "vertical"
        ? y < centerY ? "before" : "after"
        : x < centerX ? "before" : "after";
    best = { id, position, distance };
  }
  return best ? { id: best.id, position: best.position } : null;
}

/** Auto-scroll speed (px/frame, signed) for a pointer coordinate on one axis. */
export function touchReorderScrollSpeed(
  pointer: number,
  start: number,
  end: number,
  edge = TOUCH_REORDER_EDGE_PX,
  max = TOUCH_REORDER_MAX_SCROLL_PX,
): number {
  if (end - start <= edge * 2) return 0;
  if (pointer < start + edge) {
    return -Math.ceil(max * Math.min(1, (start + edge - pointer) / edge));
  }
  if (pointer > end - edge) {
    return Math.ceil(max * Math.min(1, (pointer - (end - edge)) / edge));
  }
  return 0;
}

function cssEscape(value: string): string {
  return typeof CSS !== "undefined" && typeof CSS.escape === "function"
    ? CSS.escape(value)
    : value.replace(/["\\]/g, "\\$&");
}

/** Attach touch reorder to `container`. Returns the cleanup. */
export function attachTouchReorder(
  container: HTMLElement,
  read: () => TouchReorderOptions,
): () => void {
  const itemSelectorOf = (options: TouchReorderOptions) =>
    `[${options.idAttribute}]${options.itemSelector ?? ""}`;
  const itemOf = (target: EventTarget | null): HTMLElement | null => {
    const options = read();
    if (!(target instanceof Element)) return null;
    const item = target.closest(itemSelectorOf(options));
    return item instanceof HTMLElement && container.contains(item) ? item : null;
  };
  const idOf = (item: Element) => item.getAttribute(read().idAttribute) ?? "";
  const itemById = (id: string) =>
    container.querySelector<HTMLElement>(
      `[${read().idAttribute}="${cssEscape(id)}"]${read().itemSelector ?? ""}`,
    );
  const items = () =>
    Array.from(container.querySelectorAll<HTMLElement>(itemSelectorOf(read())));

  // Each item takes the finger once its gesture lifts (the guard must be on
  // it before the finger lands) and holds in touch-hold's drag mode.
  const guards = new Map<Element, () => void>();
  const syncGuards = () => {
    const live = new Set(items());
    for (const [item, cleanup] of guards) {
      if (!live.has(item as HTMLElement)) {
        cleanup();
        guards.delete(item);
      }
    }
    for (const item of live) {
      if (!guards.has(item)) guards.set(item, installObjectDragTouchGuard(item));
    }
  };
  syncGuards();
  const observer =
    typeof MutationObserver !== "undefined" ? new MutationObserver(syncGuards) : null;
  observer?.observe(container, { childList: true, subtree: true });

  let gesture: ObjectDragGesture | null = null;
  /** A finger is on an item: the list's HTML5 drag must not start. */
  let touchDown = false;
  let drag: {
    ids: readonly string[];
    lifted: HTMLElement[];
    scroller: HTMLElement;
    startScrollLeft: number;
    startScrollTop: number;
    lastX: number;
    lastY: number;
    /** Screen pixels per CSS pixel of the lifted items and of the scroller. */
    liftScale: number;
    scrollScale: number;
    dx: number;
    dy: number;
    drop: { id: string; position: TouchReorderPosition } | null;
    frame: number | null;
    indicator: HTMLDivElement;
  } | null = null;

  const paintLifted = () => {
    if (!drag) return;
    // the finger's travel is in screen pixels and the scroll in the
    // scroller's own; the transform is in the item's own, which a zoom
    // (Touch mode) draws larger than the screen's
    const scrolledX = (drag.scroller.scrollLeft - drag.startScrollLeft) * drag.scrollScale;
    const scrolledY = (drag.scroller.scrollTop - drag.startScrollTop) * drag.scrollScale;
    const x = (drag.dx - scrolledX) / drag.liftScale;
    const y = (drag.dy - scrolledY) / drag.liftScale;
    for (const element of drag.lifted) {
      element.style.setProperty("transform", `translate(${x}px, ${y}px) scale(1.04)`);
    }
  };

  const paintIndicator = () => {
    if (!drag) return;
    const options = read();
    const orientation = options.orientation ?? "vertical";
    const moving = new Set(drag.ids);
    const candidates = items()
      .map((element) => ({ id: idOf(element), rect: element.getBoundingClientRect() }))
      .filter((candidate) => candidate.id && !moving.has(candidate.id));
    drag.drop = touchReorderDropAt(candidates, drag.lastX, drag.lastY, orientation);
    const indicator = drag.indicator;
    const target = drag.drop ? candidates.find((c) => c.id === drag?.drop?.id) : null;
    if (!drag.drop || !target) {
      indicator.style.display = "none";
      return;
    }
    const { rect } = target;
    indicator.style.display = "block";
    indicator.setAttribute("data-target-id", drag.drop.id);
    indicator.setAttribute("data-position", drag.drop.position);
    if (orientation === "vertical") {
      const y = drag.drop.position === "before" ? rect.top : rect.bottom;
      Object.assign(indicator.style, {
        left: `${rect.left}px`,
        top: `${y - 1.5}px`,
        width: `${rect.right - rect.left}px`,
        height: "3px",
      });
    } else {
      const x = drag.drop.position === "before" ? rect.left : rect.right;
      Object.assign(indicator.style, {
        left: `${x - 1.5}px`,
        top: `${rect.top}px`,
        width: "3px",
        height: `${rect.bottom - rect.top}px`,
      });
    }
  };

  const tick = () => {
    if (!drag) return;
    drag.frame = null;
    const options = read();
    const box = drag.scroller.getBoundingClientRect();
    const orientation = options.orientation ?? "vertical";
    const speedY =
      orientation === "horizontal"
        ? 0
        : touchReorderScrollSpeed(drag.lastY, box.top, box.bottom);
    const speedX =
      orientation === "vertical"
        ? 0
        : touchReorderScrollSpeed(drag.lastX, box.left, box.right);
    if (speedX || speedY) {
      drag.scroller.scrollLeft += speedX;
      drag.scroller.scrollTop += speedY;
      paintLifted();
      paintIndicator();
      drag.frame = requestAnimationFrame(tick);
    }
  };

  const endDrag = () => {
    if (!drag) return;
    if (drag.frame !== null) cancelAnimationFrame(drag.frame);
    for (const element of drag.lifted) {
      element.style.removeProperty("transform");
      element.style.removeProperty("z-index");
      element.style.removeProperty("box-shadow");
      element.style.removeProperty("transition");
      element.removeAttribute("data-touch-reorder-lifted");
    }
    drag.indicator.remove();
    container.removeAttribute("data-touch-reordering");
    drag = null;
  };

  const restoreDraggable: Array<() => void> = [];
  const releaseTouch = () => {
    touchDown = false;
    for (const restore of restoreDraggable.splice(0)) restore();
  };

  const onPointerDown = (event: PointerEvent) => {
    const options = read();
    const mouse = options.mouse === true && event.pointerType === "mouse";
    if (
      options.enabled === false ||
      (!objectDragIsTouch(event) && !mouse) ||
      event.isPrimary === false
    ) {
      return;
    }
    if (event.button > 0) return;
    const item = itemOf(event.target);
    if (!item) return;
    const id = idOf(item);
    if (!id || (options.canLift && !options.canLift(id))) return;
    if (!mouse) touchDown = true;
    // iPad Safari lifts a `draggable` element into its own drag-and-drop on a
    // long press: the finger's reorder is this one
    if (item.getAttribute("draggable") === "true") {
      item.setAttribute("draggable", "false");
      restoreDraggable.push(() => {
        if (item.getAttribute("draggable") === "false") item.setAttribute("draggable", "true");
      });
    }
    // A mouse is lifted by the gesture at once (no hold); the tile is only
    // picked up once it travels, so a click selects and nothing moves.
    let started = false;
    const startDrag = () => {
      if (started) return;
      started = true;
      (() => {
          const ids = (options.movingIds?.(id) ?? [id]).filter(Boolean);
          const moving = ids.includes(id) ? ids : [id];
          const lifted = moving
            .map((movingId) => (movingId === id ? item : itemById(movingId)))
            .filter((element): element is HTMLElement => Boolean(element));
          const scroller = options.scrollContainer?.() ?? container;
          const indicator = document.createElement("div");
          indicator.setAttribute("data-testid", INDICATOR_TEST_ID);
          indicator.setAttribute("aria-hidden", "true");
          Object.assign(indicator.style, {
            position: "fixed",
            display: "none",
            pointerEvents: "none",
            zIndex: "2147483000",
            borderRadius: "2px",
            background: "var(--office-accent, #c43e1c)",
          });
          document.body.appendChild(indicator);
          for (const element of lifted) {
            element.setAttribute("data-touch-reorder-lifted", "true");
            element.style.setProperty("z-index", "5");
            element.style.setProperty("box-shadow", "0 8px 24px rgb(0 0 0 / 28%)");
            element.style.setProperty("transition", "transform 60ms linear");
          }
          container.setAttribute("data-touch-reordering", "true");
          drag = {
            ids: moving,
            lifted,
            scroller,
            startScrollLeft: scroller.scrollLeft,
            startScrollTop: scroller.scrollTop,
            lastX: event.clientX,
            lastY: event.clientY,
            liftScale: lifted[0] ? screenScaleOf(lifted[0]) : 1,
            scrollScale: screenScaleOf(scroller),
            dx: 0,
            dy: 0,
            drop: null,
            frame: null,
            indicator,
          };
          paintLifted();
          options.onLift?.(id);
      })();
    };
    gesture = beginObjectDragGesture(
      event,
      {
        onLift: () => {
          if (!mouse) startDrag();
        },
        onMove: (move, delta) => {
          if (mouse) startDrag();
          if (!drag) return;
          drag.lastX = move.clientX;
          drag.lastY = move.clientY;
          drag.dx = delta.dx;
          drag.dy = delta.dy;
          paintLifted();
          paintIndicator();
          if (drag.frame === null) drag.frame = requestAnimationFrame(tick);
        },
        onEnd: (_up, moved) => {
          const current = drag;
          endDrag();
          gesture = null;
          releaseTouch();
          if (!moved || !current?.drop) return;
          read().onReorder({
            ids: current.ids,
            targetId: current.drop.id,
            position: current.drop.position,
            source: mouse ? "mouse" : "touch",
          });
        },
        onCancel: () => {
          endDrag();
          gesture = null;
          releaseTouch();
        },
      },
      { element: item },
    );
  };

  const onPointerEnd = (event: PointerEvent) => {
    // a tap or a scroll that never lifted ends here too
    if (!gesture && touchDown && objectDragIsTouch(event)) releaseTouch();
  };

  const onDragStart = (event: DragEvent) => {
    if (touchDown || drag) event.preventDefault();
  };

  const onKeyDown = (event: KeyboardEvent) => {
    const options = read();
    if (options.enabled === false || options.keyboard === false) return;
    if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
    const orientation = options.orientation ?? "vertical";
    const back =
      event.key === "ArrowUp" || (orientation !== "vertical" && event.key === "ArrowLeft");
    const forward =
      event.key === "ArrowDown" ||
      (orientation !== "vertical" && event.key === "ArrowRight");
    if (!back && !forward) return;
    const item = itemOf(event.target);
    if (!item || event.target !== item) return;
    const id = idOf(item);
    if (!id || (options.canLift && !options.canLift(id))) return;
    const moving = options.movingIds?.(id) ?? [id];
    const ids = moving.includes(id) ? moving : [id];
    const movingSet = new Set(ids);
    const order = items().map(idOf).filter(Boolean);
    const rest = order.filter((candidate) => !movingSet.has(candidate));
    if (rest.length === 0) return;
    let targetId: string | undefined;
    let position: TouchReorderPosition;
    if (event.shiftKey) {
      targetId = back ? rest[0] : rest[rest.length - 1];
      position = back ? "before" : "after";
    } else {
      // one place: past the nearest item beside the block
      const first = order.findIndex((candidate) => movingSet.has(candidate));
      const last = order.length - 1 - [...order].reverse().findIndex((c) => movingSet.has(c));
      targetId = back ? order.slice(0, first).reverse().find((c) => !movingSet.has(c)) : order.slice(last + 1).find((c) => !movingSet.has(c));
      position = back ? "before" : "after";
    }
    event.preventDefault();
    event.stopPropagation();
    if (!targetId) return;
    options.onReorder({ ids, targetId, position, source: "keyboard" });
    requestAnimationFrame(() => itemById(id)?.focus());
  };

  container.addEventListener("pointerdown", onPointerDown, true);
  window.addEventListener("pointerup", onPointerEnd, true);
  window.addEventListener("pointercancel", onPointerEnd, true);
  container.addEventListener("dragstart", onDragStart, true);
  container.addEventListener("keydown", onKeyDown);
  return () => {
    gesture?.cancel("aborted");
    endDrag();
    releaseTouch();
    observer?.disconnect();
    for (const cleanup of guards.values()) cleanup();
    guards.clear();
    container.removeEventListener("pointerdown", onPointerDown, true);
    window.removeEventListener("pointerup", onPointerEnd, true);
    window.removeEventListener("pointercancel", onPointerEnd, true);
    container.removeEventListener("dragstart", onDragStart, true);
    container.removeEventListener("keydown", onKeyDown);
  };
}

/**
 * React: touch reorder on the list inside `container` (an element kept in
 * state by a callback ref, or a ref object). Options are read live (the
 * latest render's), so callbacks need not be stable. `revision` re-attaches
 * when a ref object's element is swapped.
 */
export function useTouchReorder(
  container: RefObject<HTMLElement | null> | HTMLElement | null,
  options: TouchReorderOptions,
  revision?: unknown,
): void {
  const optionsRef = useRef(options);
  useLayoutEffect(() => {
    optionsRef.current = options;
  });
  const enabled = options.enabled !== false;
  useEffect(() => {
    const element =
      container && "current" in container ? container.current : container;
    if (!element || !enabled) return;
    return attachTouchReorder(element, () => optionsRef.current);
  }, [container, enabled, revision]);
}
