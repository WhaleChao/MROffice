/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isDocumentRtl } from "./document-direction";

/** Set on a menu while it is bounded by this observer. */
export const POPOVER_FITTED_ATTRIBUTE = "data-popover-fitted";

/** Menus by role, by the shared menu card, or placed by Radix. */
const DECLARED_SELECTOR = [
  '[role="menu"]',
  '[role="listbox"]',
  '[role="dialog"]:not([aria-modal="true"])',
  ".docs-cupertino-menu",
  "[data-radix-popper-content-wrapper] > *",
  "[data-office-menu-enter]",
  "[data-motion-exit]",
  "[data-motion-surface]",
].join(", ");

/** Hand-built surfaces known only by their name; one counts when it is
 * drawn as a card above the page (see isSurface). */
const NAMED_SELECTOR = [
  '[class*="popover" i]',
  '[class*="dropdown" i]',
  '[class*="flyout" i]',
  '[class*="popup" i]',
  '[class*="picker" i]',
  '[class*="popout" i]',
  '[class*="menu" i]:not([class*="menubar" i])',
].join(", ");

const CANDIDATE_SELECTOR = `${DECLARED_SELECTOR}, ${NAMED_SELECTOR}`;

/** Room kept between a fitted menu and the edge it was fitted to. */
const EDGE_MARGIN = 8;
/** A menu is never squeezed below this; past it the region is not usable. */
const MIN_HEIGHT = 96;
/** How deep inside an added node a menu is looked for (a portal's wrapper). */
const ADDED_DEPTH = 3;

type Box = { top: number; left: number; bottom: number; right: number };

type Fitted = {
  maxHeight: string;
  maxWidth: string;
  overflowY: string;
  overscroll: string;
  touchAction: string;
  translate: string;
};

let installed = false;
const fitted = new WeakMap<HTMLElement, Fitted>();





const dormant = new WeakSet<HTMLElement>();

function isCandidate(element: Element): element is HTMLElement {
  if (!(element instanceof HTMLElement) || !element.matches(CANDIDATE_SELECTOR)) return false;
  if (element.closest('[data-popover-fit="off"]')) return false;
  // a leaving copy (lib/surface-motion) is not the menu
  return !element.hasAttribute("data-motion-ghost");
}

function isFloating(element: HTMLElement, style: CSSStyleDeclaration): boolean {
  if (style.position === "absolute" || style.position === "fixed") return true;
  const wrapper = element.parentElement;
  return Boolean(wrapper?.hasAttribute("data-radix-popper-content-wrapper"));
}

/** A name-matched element is a menu only when it is a card raised over the
 * page — stacked, and painted with a shadow or a ground — never a control
 * or a decoration that merely hangs past a clipped edge. */
function isSurface(element: HTMLElement, style: CSSStyleDeclaration): boolean {
  if (element.matches(DECLARED_SELECTOR)) return true;
  if (element.matches("button, a, input, select, textarea, label, svg, img")) return false;
  const stacked = style.zIndex !== "auto" && Number(style.zIndex) >= 1;
  const painted = style.boxShadow !== "none" || !/^(transparent|rgba\(0, 0, 0, 0\))$/.test(style.backgroundColor);
  return stacked && painted;
}








export function floatingSurfaceOf(target: EventTarget | null, host?: Element | null): HTMLElement | null {
  let node = target instanceof Element ? target : target instanceof Node ? target.parentElement : null;
  for (; node && node !== host && node !== document.body; node = node.parentElement) {
    if (!(node instanceof HTMLElement) || node.hasAttribute("data-motion-ghost")) continue;
    // a native popover (top layer) is a surface whatever its name
    if (node.hasAttribute("popover") && isPopoverOpen(node)) return node;
    if (!node.matches(CANDIDATE_SELECTOR)) continue;
    const style = getComputedStyle(node);
    if (isFloating(node, style) && isSurface(node, style)) return node;
  }
  return null;
}

function isPopoverOpen(element: HTMLElement): boolean {
  try {
    return element.matches(":popover-open");
  } catch {
    // an engine without the Popover API has no top layer to open into
    return false;
  }
}

/** Whether something from `target` up to `surface` can still scroll by this
 * wheel on its main axis. */
function canScrollWithin(event: WheelEvent, target: Element, surface: HTMLElement): boolean {
  const vertical = Math.abs(event.deltaY) >= Math.abs(event.deltaX);
  const delta = vertical ? event.deltaY : event.deltaX;
  if (delta === 0) return false;
  for (let node: Element | null = target; node; node = node === surface ? null : node.parentElement) {
    if (!(node instanceof HTMLElement)) continue;
    const size = vertical ? node.scrollHeight - node.clientHeight : node.scrollWidth - node.clientWidth;
    if (size <= 1) continue;
    const overflow = vertical ? getComputedStyle(node).overflowY : getComputedStyle(node).overflowX;
    if (overflow !== "auto" && overflow !== "scroll") continue;
    const at = Math.abs(vertical ? node.scrollTop : node.scrollLeft);
    if (delta > 0 ? at < size - 1 : at > 0) return true;
  }
  return false;
}

/**
 * A wheel that starts in a floating surface below `host` belongs to that
 * surface: it scrolls there (by the browser) or nowhere -- never the strip,
 * window or sheet the surface hangs from. True when the event was such a
 * wheel; the caller then leaves it alone.
 */
export function wheelKeptByFloatingSurface(event: WheelEvent, host: Element): boolean {
  const surface = floatingSurfaceOf(event.target, host);
  if (!surface) return false;
  const target = event.target instanceof Element ? event.target : surface;
  if (event.cancelable && !canScrollWithin(event, target, surface)) event.preventDefault();
  return true;
}

/** The region of the page the menu can show (null: nothing usable). */
function visibleRegion(element: HTMLElement): Box | null {
  const viewport = window.visualViewport;
  let visible: Box = viewport
    ? {
        top: viewport.offsetTop,
        left: viewport.offsetLeft,
        bottom: viewport.offsetTop + viewport.height,
        right: viewport.offsetLeft + viewport.width,
      }
    : { top: 0, left: 0, bottom: window.innerHeight, right: window.innerWidth };
  // An ancestor's overflow cuts the menu only from the menu's containing
  // block outward: a fixed menu's is the viewport unless a transformed (or
  // contained) ancestor makes one, an absolute menu's is its nearest
  // positioned ancestor.
  const fixed = getComputedStyle(element).position === "fixed";
  let reached = false;
  for (let node = element.parentElement; node && node !== document.body; node = node.parentElement) {
    const style = getComputedStyle(node);
    if (!reached) {
      const makesBlock =
        style.transform !== "none" ||
        style.perspective !== "none" ||
        style.filter !== "none" ||
        /\b(transform|perspective|filter)\b/.test(style.willChange) ||
        /\b(paint|layout|strict|content)\b/.test(style.contain);
      reached = makesBlock || (!fixed && style.position !== "static");
    }
    if (reached && (style.overflowX !== "visible" || style.overflowY !== "visible")) {
      visible = intersect(visible, node.getBoundingClientRect());
    }
    if (reached && style.position === "fixed") break;
  }
  if (visible.bottom - visible.top < MIN_HEIGHT || visible.right - visible.left < 48) return null;
  return visible;
}

function intersect(box: Box, rect: DOMRect): Box {
  return {
    top: Math.max(box.top, rect.top),
    left: Math.max(box.left, rect.left),
    bottom: Math.min(box.bottom, rect.bottom),
    right: Math.min(box.right, rect.right),
  };
}

function release(element: HTMLElement): void {
  const saved = fitted.get(element);
  if (!saved) return;
  fitted.delete(element);
  element.removeAttribute(POPOVER_FITTED_ATTRIBUTE);
  element.style.maxHeight = saved.maxHeight;
  element.style.maxWidth = saved.maxWidth;
  element.style.overflowY = saved.overflowY;
  element.style.overscrollBehavior = saved.overscroll;
  element.style.touchAction = saved.touchAction;
  element.style.translate = saved.translate;
}

/** Asks for another pass (an entrance ended, a menu is not placed yet). */
let again: () => void = () => {};
/** Passes a menu not yet placed (zero-sized, transparent, parked off the
 * region, as Radix holds one for a frame) is looked at again. */
const PLACEMENT_RETRIES = 30;
const unplaced = new WeakMap<HTMLElement, number>();

function waitForPlacement(element: HTMLElement): void {
  dormant.add(element);
  const tries = unplaced.get(element) ?? 0;
  if (tries >= PLACEMENT_RETRIES) return;
  unplaced.set(element, tries + 1);
  again();
}

/** Fits one menu; false when the element is not a floating menu at all
 * (an in-flow list, a plain element with a matching name) and need not be
 * watched. */
function fit(element: HTMLElement): boolean {
  // an entrance still moving it is measured once it has settled (its
  // `finished`: an animationend is not always delivered to the document)
  const moving = (element.getAnimations?.() ?? []).filter(
    (animation) => animation.playState === "running" && animation.effect?.getTiming().iterations !== Infinity,
  );
  if (moving.length > 0) {
    void Promise.allSettled(moving.map((animation) => animation.finished)).then(again);
    return true;
  }
  // What is not (yet) a floating card is let go before anything is written
  // to it. The names match plenty that is not a menu (a menu item, a "New"
  // menu button), and a write to its style, even one put back at once,
  // re-serialises the server's style attribute: React then reports a
  // hydration mismatch on every such element it had not hydrated yet
  
  if (!fitted.has(element)) {
    const style = getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden") {
      dormant.add(element);
      return true;
    }
    if (!isFloating(element, style) || !isSurface(element, style)) return false;
  }
  // Measuring takes the bound off and puts it back in one frame; a menu that
  // transitions its size or position must not animate between the two.
  const transition = element.style.transition;
  element.style.transition = "none";
  try {
    return measureAndFit(element);
  } finally {
    void element.offsetHeight;
    element.style.transition = transition;
  }
}

function measureAndFit(element: HTMLElement): boolean {
  // Measure the menu as its own CSS lays it out.
  release(element);
  const style = getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden") {
    dormant.add(element);
    return true;
  }
  if (!isFloating(element, style) || !isSurface(element, style)) return false;
  // A menu that scrolls keeps its wheel and swipe at its ends too: chained
  // on, they scrolled whatever the menu hangs over.
  if (/^(auto|scroll)$/.test(style.overflowY) && style.overscrollBehaviorY === "auto" && !element.style.overscrollBehaviorY) {
    element.style.overscrollBehaviorY = "contain";
  }
  if ((style.opacity !== "" && Number(style.opacity) === 0) || style.pointerEvents === "none") {
    waitForPlacement(element);
    return true;
  }
  const rect = element.getBoundingClientRect();
  if (rect.width <= 2 || rect.height <= 2) {
    waitForPlacement(element);
    return true;
  }
  // a name-matched chip, badge or caret is no menu, even raised and painted
  if (!element.matches(DECLARED_SELECTOR) && (rect.width < 96 || rect.height < 40)) return false;
  const region = visibleRegion(element);
  if (!region) return true;
  // wholly outside: not placed yet, or a measuring copy / parked surface
  if (rect.bottom <= region.top || rect.top >= region.bottom || rect.right <= region.left || rect.left >= region.right) {
    waitForPlacement(element);
    return true;
  }
  const top = region.top + EDGE_MARGIN;
  const bottom = region.bottom - EDGE_MARGIN;
  const left = region.left + EDGE_MARGIN;
  const right = region.right - EDGE_MARGIN;
  unplaced.delete(element);
  // Only a menu that actually crosses the region's edge is touched; one that
  // merely sits close to it keeps its own placement.
  const overflowsY = rect.top < region.top - 0.5 || rect.bottom > region.bottom + 0.5;
  const overflowsX = rect.left < region.left - 0.5 || rect.right > region.right + 0.5;
  if (!overflowsY && !overflowsX) return true;

  // Lengths written in the menu's own CSS pixels: a touch-mode `zoom` or an
  // entrance scale makes them differ from the page's.
  const ownHeight = parseFloat(style.height) || rect.height;
  const ownWidth = parseFloat(style.width) || rect.width;
  const scaleY = rect.height / ownHeight || 1;
  const scaleX = rect.width / ownWidth || 1;

  fitted.set(element, {
    maxHeight: element.style.maxHeight,
    maxWidth: element.style.maxWidth,
    overflowY: element.style.overflowY,
    overscroll: element.style.overscrollBehavior,
    touchAction: element.style.touchAction,
    translate: element.style.translate,
  });
  element.setAttribute(POPOVER_FITTED_ATTRIBUTE, "");

  // Radix places its own content again whenever its size changes, so a bound
  // read from where it sits now would chase that placement and shrink on
  // every pass: it gets the whole region and Radix picks the side.
  const radixPlaced = Boolean(element.parentElement?.hasAttribute("data-radix-popper-content-wrapper"));
  const room = Math.max(MIN_HEIGHT, bottom - top);
  if (overflowsY) {
    // A hand-placed menu keeps the part that stays in view when it is
    // shortened toward the edge it is anchored to; never more than the region.
    const keep = radixPlaced
      ? room
      : Math.min(room, Math.max(MIN_HEIGHT, Math.min(rect.bottom, bottom) - Math.max(rect.top, top)));
    element.style.maxHeight = `${Math.floor(keep / scaleY)}px`;
    // a bound without a scroll would only cut the rest off
    if (style.overflowY !== "auto" && style.overflowY !== "scroll") element.style.overflowY = "auto";
    element.style.overscrollBehavior = "contain";
    element.style.touchAction = "pan-y";
  }
  const roomX = Math.max(48, right - left);
  if (overflowsX && rect.width > roomX) element.style.maxWidth = `${Math.floor(roomX / scaleX)}px`;

  // Whatever is still outside (a menu anchored at the far edge, or one wider
  // than its place) is moved in.
  const after = element.getBoundingClientRect();
  let dy = 0;
  if (overflowsY && after.top < top) dy = top - after.top;
  else if (overflowsY && after.bottom > bottom) dy = Math.max(top - after.top, bottom - after.bottom);
  let dx = 0;
  if (isDocumentRtl()) {
    // An rtl interface keeps a menu's right (inline-start) edge in view first.
    if (overflowsX && after.right > right) dx = right - after.right;
    else if (overflowsX && after.left < left) dx = Math.min(right - after.right, left - after.left);
  } else if (overflowsX && after.left < left) dx = left - after.left;
  else if (overflowsX && after.right > right) dx = Math.max(left - after.left, right - after.right);
  // (A shift never changes the menu's size, so Radix does not re-place it in
  // answer; under touch-mode zoom its placement can itself be off.)
  if (Math.abs(dx) > 0.5 || Math.abs(dy) > 0.5) {
    element.style.translate = `${Math.round(dx / scaleX)}px ${Math.round(dy / scaleY)}px`;
  }
  return true;
}

/**
 * Start fitting menus in this document. Idempotent; does nothing on the
 * server or in a DOM without layout (a test DOM's boxes are all empty).
 */
export function installPopoverFit(): void {
  if (installed || typeof window === "undefined" || typeof MutationObserver === "undefined") return;
  installed = true;
  const menus = new Set<HTMLElement>();
  let frame = 0;
  // Whether the next pass looks at every menu: everything but a scroll asks
  // it. A scroll asks only for the menus it can move (scrolled).
  let wholePass = false;
  const scrolled = new Set<HTMLElement>();
  const run = () => {
    frame = 0;
    const whole = wholePass;
    const moved = [...scrolled];
    wholePass = false;
    scrolled.clear();
    for (const menu of [...menus]) {
      if (!menu.isConnected) {
        menus.delete(menu);
        resizes?.unobserve(menu);
        continue;
      }
      if (!whole && (dormant.has(menu) || !moved.includes(menu))) continue;
      dormant.delete(menu);
      if (!fit(menu)) {
        menus.delete(menu);
        resizes?.unobserve(menu);
      }
    }
  };
  const schedule = () => {
    wholePass = true;
    if (frame || menus.size === 0) return;
    frame = window.requestAnimationFrame(run);
  };
  // A scroll moves the menus inside what scrolled, the ones placed against an
  // anchor (Radix's, wherever their anchor is) and the ones already bounded;
  // a menu elsewhere stays where it was. Each one measured costs the page a
  // whole style and layout pass, every frame of the scroll.
  const scheduleScroll = (target: EventTarget | null) => {
    const within = target instanceof Element ? target : null;
    for (const menu of menus) {
      if (dormant.has(menu)) continue;
      if (
        !within ||
        within.contains(menu) ||
        fitted.has(menu) ||
        menu.parentElement?.hasAttribute("data-radix-popper-content-wrapper")
      ) {
        scrolled.add(menu);
      }
    }
    if (frame || scrolled.size === 0) return;
    frame = window.requestAnimationFrame(run);
  };
  const resizes =
    typeof ResizeObserver === "undefined"
      ? null
      : // content arriving late (a list that loads) changes what fits; our
        // own bound settles at the same size on the next pass, so it stops
        new ResizeObserver(schedule);
  const track = (element: HTMLElement) => {
    if (menus.has(element)) return;
    menus.add(element);
    resizes?.observe(element);
  };
  const look = (node: Node, depth: number) => {
    if (!(node instanceof HTMLElement)) return;
    if (isCandidate(node)) track(node);
    if (depth >= ADDED_DEPTH) return;
    for (const child of node.children) look(child, depth + 1);
  };
  const mutations = new MutationObserver((records) => {
    const before = menus.size;
    let changed = false;
    for (const record of records) {
      if (record.type === "attributes") {
        if (record.target instanceof HTMLElement && isCandidate(record.target)) {
          unplaced.delete(record.target);
          track(record.target);
          changed = true;
        }
        continue;
      }
      for (const node of record.addedNodes) look(node, 0);
    }
    if (changed || menus.size !== before) schedule();
  });
  mutations.observe(document.documentElement, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ["hidden", "data-state", "open"],
  });
  window.addEventListener("resize", schedule, { passive: true });
  window.visualViewport?.addEventListener("resize", schedule, { passive: true });
  window.visualViewport?.addEventListener("scroll", () => scheduleScroll(null), { passive: true });
  // an anchor scrolled under its open menu
  const onScroll = (event: Event) => {
    if (event.target instanceof Node && [...menus].some((menu) => menu.contains(event.target as Node))) return;
    scheduleScroll(event.target);
  };
  window.addEventListener("scroll", onScroll, { capture: true, passive: true });
  again = schedule;
  for (const element of document.querySelectorAll(CANDIDATE_SELECTOR)) {
    if (isCandidate(element)) track(element);
  }
  schedule();
}
