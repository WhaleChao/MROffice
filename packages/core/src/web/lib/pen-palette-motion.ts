/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  SURFACE_ENTER_EASE,
  SURFACE_EXIT_EASE,
  SURFACE_MOVE_EASE,
  surfaceMotionAllowed,
  surfacePictureOf,
} from "./desk-motion";

/** On the live palette's root element. */
export const PEN_PALETTE_ROOT_ATTRIBUTE = "data-pen-palette-root";
/** On the palette's card ("backdrop") and its face ("body"). */
export const PEN_PALETTE_PART_ATTRIBUTE = "data-pen-palette-part";
const GHOST_ATTRIBUTE = "data-pen-palette-ghost";
/** Set on a root whose arrival a morph has taken over (no plain entrance). */
const MORPHED_ATTRIBUTE = "data-pen-palette-morphed";
const COMPACT_ATTRIBUTE = "data-pen-compact";
const HOST_SELECTOR = ".wm-window";
const ANIMATION_ID = "pen-palette-motion";

/** The ribbon folding into the palette, and the palette growing back into it. */
export const PEN_PALETTE_MORPH_MS = 380;
/** The palette folding, opening, or turning on its side. */
export const PEN_PALETTE_RESHAPE_MS = 280;
const ENTER_MS = 200;
const RIBBON_FADE_MS = 200;
const GHOST_FADE_MS = 140;
/** A palette that arrived this long before its ribbon folded still morphs out of it. */
const MORPH_ARRIVAL_WINDOW_MS = 260;
/** How long a closed palette's picture waits for its ribbon to come back. */
const EXIT_WAIT_MS = 80;

export interface PaletteFrame {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly radius: number;
}

function partOf(root: Element, part: "backdrop" | "body"): HTMLElement | null {
  return root.querySelector<HTMLElement>(`:scope > [${PEN_PALETTE_PART_ATTRIBUTE}="${part}"]`);
}

function radiusOf(element: Element | null): number {
  if (!element) return 0;
  return Number.parseFloat(getComputedStyle(element).borderTopLeftRadius) || 0;
}

function settle(element: Element | null): void {
  for (const animation of element?.getAnimations?.() ?? []) {
    if (animation.id === ANIMATION_ID) animation.cancel();
  }
}

function play(
  element: Element,
  keyframes: Keyframe[],
  timing: KeyframeAnimationOptions,
): Animation | null {
  if (typeof element.animate !== "function") return null;
  try {
    return element.animate(keyframes, { ...timing, id: ANIMATION_ID });
  } catch {
    // An engine without Web Animations shows the new frame at once.
    return null;
  }
}

/** Where the palette's left/top are written: its window, or the viewport. */
function coordinateHostOf(root: HTMLElement): HTMLElement | null {
  return root.style.position === "fixed" ? null : root.parentElement;
}

/**
 * `element`'s frame in the coordinates the palette's left/top are written in:
 * `host`'s own (a window may be zoomed on the canvas: screen px over its
 * scale), or the viewport's when the palette is fixed.
 */
export function frameInHost(host: HTMLElement | null, element: HTMLElement): PaletteFrame {
  const rect = element.getBoundingClientRect();
  const radius = radiusOf(element);
  if (!host) return { x: rect.left, y: rect.top, width: rect.width, height: rect.height, radius };
  const box = host.getBoundingClientRect();
  const scale = host.offsetWidth > 0 ? box.width / host.offsetWidth || 1 : 1;
  return {
    x: (rect.left - box.left) / scale - host.clientLeft,
    y: (rect.top - box.top) / scale - host.clientTop,
    width: rect.width / scale,
    height: rect.height / scale,
    radius,
  };
}

/** The palette's own frame, as its left/top write it (its lift ignored). */
export function paletteFrame(root: HTMLElement): PaletteFrame {
  return {
    x: root.offsetLeft,
    y: root.offsetTop,
    width: root.offsetWidth,
    height: root.offsetHeight,
    radius: radiusOf(partOf(root, "backdrop")),
  };
}

interface MorphOptions {
  readonly duration: number;
  /** Share of the run before the new face shows. */
  readonly reveal: number;
  /** The card fades in over the start: it leaves a ribbon fading the other way. */
  readonly cardFadesIn?: boolean;
}

/**
 * The palette, laid out at its new frame, drawn arriving from `from`: its
 * card glides and reshapes from there, its face shows inside it.
 */
function morphPaletteFrom(root: HTMLElement, from: PaletteFrame, options: MorphOptions): boolean {
  const backdrop = partOf(root, "backdrop");
  const body = partOf(root, "body");
  if (!backdrop || !body) return false;
  const to = paletteFrame(root);
  if (!(to.width > 0 && to.height > 0 && from.width > 0 && from.height > 0)) return false;
  settle(root);
  settle(backdrop);
  settle(body);
  root.setAttribute(MORPHED_ATTRIBUTE, "");
  const dx = from.x - to.x;
  const dy = from.y - to.y;
  const glide: KeyframeAnimationOptions = { duration: options.duration, easing: SURFACE_MOVE_EASE };
  play(
    backdrop,
    [
      {
        left: `${dx}px`,
        top: `${dy}px`,
        width: `${from.width}px`,
        height: `${from.height}px`,
        borderRadius: `${from.radius}px`,
      },
      { left: "0px", top: "0px", width: `${to.width}px`, height: `${to.height}px`, borderRadius: `${to.radius}px` },
    ],
    glide,
  );
  if (options.cardFadesIn) {
    play(backdrop, [{ opacity: 0 }, { opacity: 1, offset: 0.4 }, { opacity: 1 }], {
      duration: options.duration,
      easing: "linear",
    });
  }
  // The face is cut to the card as it goes, and shows once it is nearly there.
  const right = to.width - (dx + from.width);
  const bottom = to.height - (dy + from.height);
  play(
    body,
    [
      { clipPath: `inset(${dy}px ${right}px ${bottom}px ${dx}px round ${from.radius}px)` },
      { clipPath: `inset(0px 0px 0px 0px round ${to.radius}px)` },
    ],
    glide,
  );
  play(
    body,
    [
      { opacity: 0, offset: 0 },
      { opacity: 0, offset: options.reveal, easing: SURFACE_ENTER_EASE },
      { opacity: 1, offset: 1 },
    ],
    { duration: options.duration },
  );
  return true;
}

/** An inert picture of the palette beside it (no ids, roles or focus stops). */
function paletteGhost(root: HTMLElement, withCard: boolean): HTMLElement | null {
  const parent = root.parentElement;
  if (!parent) return null;
  const picture = surfacePictureOf(root, GHOST_ATTRIBUTE);
  const ghost = picture.element;
  ghost.removeAttribute(PEN_PALETTE_ROOT_ATTRIBUTE);
  ghost.style.pointerEvents = "none";
  if (!withCard) {
    const card = partOf(ghost, "backdrop");
    if (card) card.style.visibility = "hidden";
  }
  parent.insertBefore(ghost, root.nextSibling);
  picture.settle();
  return ghost;
}

function fadeAway(ghost: HTMLElement, duration: number): void {
  const animation = play(ghost, [{ opacity: 1 }, { opacity: 0, scale: "0.96" }], {
    duration,
    easing: SURFACE_EXIT_EASE,
    fill: "forwards",
  });
  if (!animation) {
    ghost.remove();
    return;
  }
  animation.finished.then(
    () => ghost.remove(),
    () => ghost.remove(),
  );
}

/**
 * A palette that has just been laid out comes in: a short rise and fade,
 * unless a morph (its ribbon folding) has taken its arrival over. Call it
 * once the palette's root is in the document.
 */
export function playPaletteEntrance(root: HTMLElement): void {
  if (!surfaceMotionAllowed()) return;
  root.dataset.penPaletteArrived = String(performance.now());
  queueMicrotask(() => {
    if (!root.isConnected || root.hasAttribute(MORPHED_ATTRIBUTE)) return;
    settle(root);
    play(
      root,
      [
        { opacity: 0, translate: "0px 6px", scale: "0.98" },
        { opacity: 1, translate: "0px 0px", scale: "1" },
      ],
      { duration: ENTER_MS, easing: SURFACE_ENTER_EASE },
    );
  });
}

export interface PaletteReshape {
  readonly frame: PaletteFrame;
  readonly ghost: HTMLElement | null;
}

/** Before the palette changes shape: where it was, and a picture of its old face. */
export function beginPaletteReshape(root: HTMLElement | null): PaletteReshape | null {
  if (!root || !root.isConnected || !surfaceMotionAllowed()) return null;
  return { frame: paletteFrame(root), ghost: paletteGhost(root, false) };
}

/** Once the palette is laid out in its new shape: the glide from the old one. */
export function finishPaletteReshape(root: HTMLElement | null, reshape: PaletteReshape): void {
  if (reshape.ghost) fadeAway(reshape.ghost, GHOST_FADE_MS);
  if (root?.isConnected) {
    morphPaletteFrom(root, reshape.frame, { duration: PEN_PALETTE_RESHAPE_MS, reveal: 0.25 });
  }
}

interface PendingExit {
  readonly root: HTMLElement;
  readonly ghost: HTMLElement;
  claimed: boolean;
}

const pendingExits = new Map<Element, PendingExit>();

/**
 * The palette is leaving (its root is about to be removed): a picture of it
 * stays where it was, for the ribbon coming back to grow it into itself, or,
 * when no ribbon comes, to fade where it stood.
 */
export function holdPaletteExit(root: HTMLElement): void {
  if (!root.isConnected || !surfaceMotionAllowed()) return;
  const key = root.parentElement;
  if (!key) return;
  const ghost = paletteGhost(root, true);
  if (!ghost) return;
  const record: PendingExit = { root, ghost, claimed: false };
  const previous = pendingExits.get(key);
  if (previous) {
    previous.claimed = true;
    previous.ghost.remove();
  }
  pendingExits.set(key, record);
  const drop = () => {
    if (pendingExits.get(key) === record) pendingExits.delete(key);
  };
  queueMicrotask(() => {
    // Still in the document: not gone after all (a development double mount).
    if (record.claimed || !root.isConnected) return;
    record.claimed = true;
    drop();
    ghost.remove();
  });
  window.setTimeout(() => {
    if (record.claimed) return;
    record.claimed = true;
    drop();
    fadeAway(ghost, GHOST_FADE_MS);
  }, EXIT_WAIT_MS);
}

function claimExit(key: Element | null): PendingExit | null {
  if (!key) return null;
  const record = pendingExits.get(key);
  if (!record || record.claimed || record.root.isConnected) return null;
  record.claimed = true;
  pendingExits.delete(key);
  return record;
}

/** What lies after the ribbon in its window's body: it moves when the ribbon goes. */
function followersOf(ribbon: HTMLElement): HTMLElement[] {
  const out: HTMLElement[] = [];
  const stop = ribbon.closest(".wm-body") ?? ribbon.closest(HOST_SELECTOR);
  let node: HTMLElement | null = ribbon;
  for (let depth = 0; node && node !== stop && depth < 6; depth += 1) {
    for (let sibling = node.nextElementSibling; sibling; sibling = sibling.nextElementSibling) {
      if (
        sibling instanceof HTMLElement &&
        !sibling.hasAttribute(PEN_PALETTE_ROOT_ATTRIBUTE) &&
        !sibling.hasAttribute(GHOST_ATTRIBUTE)
      ) {
        out.push(sibling);
      }
    }
    node = node.parentElement;
  }
  return out;
}

function glideFollowers(
  followers: readonly HTMLElement[],
  before: readonly DOMRect[],
  after: readonly DOMRect[],
  scale: number,
): void {
  followers.forEach((element, index) => {
    const from = before[index];
    const to = after[index];
    if (!from || !to || to.height === 0) return;
    const dx = (from.left - to.left) / scale;
    const dy = (from.top - to.top) / scale;
    if (Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5) return;
    settle(element);
    play(element, [{ translate: `${dx}px ${dy}px` }, { translate: "0px 0px" }], {
      duration: PEN_PALETTE_MORPH_MS,
      easing: SURFACE_MOVE_EASE,
    });
  });
}

function scaleOf(host: HTMLElement | null): number {
  if (!host || host.offsetWidth === 0) return 1;
  return host.getBoundingClientRect().width / host.offsetWidth || 1;
}

/** Ribbons held where they stood while they fade: how to let each go. */
const heldRibbons = new WeakMap<HTMLElement, () => void>();

const HELD_PROPERTIES = [
  "display",
  "position",
  "left",
  "top",
  "width",
  "height",
  "margin",
  "box-sizing",
  "pointer-events",
  "transform-origin",
] as const;

function holdRibbonFading(
  ribbon: HTMLElement,
  display: string,
  box: { left: number; top: number; width: number; height: number },
  toward: { x: number; y: number } | null,
): void {
  heldRibbons.get(ribbon)?.();
  const style = ribbon.style;
  const saved = HELD_PROPERTIES.map(
    (property) => [property, style.getPropertyValue(property), style.getPropertyPriority(property)] as const,
  );
  style.setProperty("display", display);
  style.setProperty("position", "absolute");
  style.setProperty("left", `${box.left}px`);
  style.setProperty("top", `${box.top}px`);
  style.setProperty("width", `${box.width}px`);
  style.setProperty("height", `${box.height}px`);
  style.setProperty("margin", "0");
  style.setProperty("box-sizing", "border-box");
  style.setProperty("pointer-events", "none");
  style.setProperty("transform-origin", toward ? `${toward.x}px ${toward.y}px` : "50% 0%");
  const animation = play(ribbon, [{ opacity: 1 }, { opacity: 0, scale: "0.96" }], {
    duration: RIBBON_FADE_MS,
    easing: SURFACE_EXIT_EASE,
    fill: "forwards",
  });
  const release = () => {
    if (heldRibbons.get(ribbon) !== release) return;
    heldRibbons.delete(ribbon);
    animation?.cancel();
    for (const [property, value, priority] of saved) {
      if (value) style.setProperty(property, value, priority);
      else style.removeProperty(property);
    }
  };
  heldRibbons.set(ribbon, release);
  if (!animation) {
    release();
    return;
  }
  animation.finished.then(release, () => {});
}

function livePaletteIn(host: HTMLElement | null): HTMLElement | null {
  const scope = host ?? document.body;
  return scope.querySelector<HTMLElement>(`:scope > [${PEN_PALETTE_ROOT_ATTRIBUTE}]`);
}

let observer: MutationObserver | null = null;

/** Our own measuring flips the attribute: what that tells the observer is dropped. */
function measureWith<T>(ribbon: HTMLElement, compact: boolean, measure: () => T): T {
  if (compact) ribbon.setAttribute(COMPACT_ATTRIBUTE, "true");
  else ribbon.removeAttribute(COMPACT_ATTRIBUTE);
  try {
    return measure();
  } finally {
    if (compact) ribbon.removeAttribute(COMPACT_ATTRIBUTE);
    else ribbon.setAttribute(COMPACT_ATTRIBUTE, "true");
    observer?.takeRecords();
  }
}

function foldRibbon(ribbon: HTMLElement): void {
  const host = ribbon.closest<HTMLElement>(HOST_SELECTOR);
  const followers = followersOf(ribbon);
  const after = followers.map((element) => element.getBoundingClientRect());
  // The ribbon as it stood, shown once more to be measured.
  const { display, box, frameInWindow, frameOnScreen, before } = measureWith(ribbon, false, () => ({
    display: getComputedStyle(ribbon).display,
    box: { left: ribbon.offsetLeft, top: ribbon.offsetTop, width: ribbon.offsetWidth, height: ribbon.offsetHeight },
    frameInWindow: frameInHost(host, ribbon),
    frameOnScreen: frameInHost(null, ribbon),
    before: followers.map((element) => element.getBoundingClientRect()),
  }));
  if (display === "none" || box.width === 0 || box.height === 0) return;
  glideFollowers(followers, before, after, scaleOf(host));
  const palette = livePaletteIn(host) ?? livePaletteIn(null);
  const arrived = Number(palette?.dataset.penPaletteArrived ?? Number.NaN);
  const fresh = palette !== null && performance.now() - arrived < MORPH_ARRIVAL_WINDOW_MS;
  let toward: { x: number; y: number } | null = null;
  if (palette) {
    const ribbonRect = ribbon.getBoundingClientRect();
    const paletteRect = palette.getBoundingClientRect();
    const scale = scaleOf(host);
    toward = {
      x: (paletteRect.left + paletteRect.width / 2 - ribbonRect.left) / scale,
      y: (paletteRect.top + paletteRect.height / 2 - ribbonRect.top) / scale,
    };
  }
  holdRibbonFading(ribbon, display, box, toward);
  if (palette && fresh) {
    morphPaletteFrom(palette, coordinateHostOf(palette) ? frameInWindow : frameOnScreen, { duration: PEN_PALETTE_MORPH_MS, reveal: 0.45, cardFadesIn: true });
  }
}

function unfoldRibbon(ribbon: HTMLElement): void {
  heldRibbons.get(ribbon)?.();
  const host = ribbon.closest<HTMLElement>(HOST_SELECTOR);
  const followers = followersOf(ribbon);
  const after = followers.map((element) => element.getBoundingClientRect());
  const before = measureWith(ribbon, true, () => followers.map((element) => element.getBoundingClientRect()));
  if (ribbon.offsetWidth === 0 || ribbon.offsetHeight === 0) return;
  glideFollowers(followers, before, after, scaleOf(host));
  const exit = claimExit(host) ?? claimExit(document.body);
  settle(ribbon);
  if (!exit) {
    play(
      ribbon,
      [
        { opacity: 0, translate: "0px -6px" },
        { opacity: 1, translate: "0px 0px" },
      ],
      { duration: ENTER_MS, easing: SURFACE_ENTER_EASE },
    );
    return;
  }
  // The palette's picture grows into the ribbon's frame and hands over to it.
  const { ghost } = exit;
  const card = partOf(ghost, "backdrop");
  const face = partOf(ghost, "body");
  const ghostFrame = paletteFrame(ghost);
  const target = frameInHost(coordinateHostOf(ghost), ribbon);
  const duration = PEN_PALETTE_MORPH_MS;
  if (face) play(face, [{ opacity: 1 }, { opacity: 0, offset: 0.35 }, { opacity: 0 }], { duration });
  let finished: Animation | null = null;
  if (card) {
    finished = play(
      card,
      [
        {
          left: "0px",
          top: "0px",
          width: `${ghostFrame.width}px`,
          height: `${ghostFrame.height}px`,
          borderRadius: `${ghostFrame.radius}px`,
          opacity: 1,
        },
        { opacity: 1, offset: 0.55 },
        {
          left: `${target.x - ghostFrame.x}px`,
          top: `${target.y - ghostFrame.y}px`,
          width: `${target.width}px`,
          height: `${target.height}px`,
          borderRadius: `${target.radius}px`,
          opacity: 0,
        },
      ],
      { duration, easing: SURFACE_MOVE_EASE, fill: "forwards" },
    );
  }
  play(ribbon, [{ opacity: 0 }, { opacity: 0, offset: 0.45 }, { opacity: 1 }], { duration, easing: "linear" });
  if (finished) {
    finished.finished.then(
      () => ghost.remove(),
      () => ghost.remove(),
    );
  } else {
    ghost.remove();
  }
}

/**
 * Watches every ribbon's `data-pen-compact` for the fold and the unfold.
 * Installed once per page, before any ribbon folds (lib/pen-first-draw).
 */
export function installPenPaletteMotion(): void {
  if (observer || typeof document === "undefined" || typeof MutationObserver === "undefined") return;
  observer = new MutationObserver((records) => {
    // Each ribbon once, from what it was before this batch to what it is now.
    const first = new Map<HTMLElement, string | null>();
    for (const record of records) {
      const target = record.target;
      if (target instanceof HTMLElement && !first.has(target)) first.set(target, record.oldValue);
    }
    if (!surfaceMotionAllowed()) return;
    for (const [ribbon, oldValue] of first) {
      if (!ribbon.isConnected) continue;
      const was = oldValue === "true";
      const now = ribbon.getAttribute(COMPACT_ATTRIBUTE) === "true";
      if (was === now) continue;
      if (now) foldRibbon(ribbon);
      else unfoldRibbon(ribbon);
    }
  });
  observer.observe(document.documentElement, {
    attributes: true,
    attributeFilter: [COMPACT_ATTRIBUTE],
    attributeOldValue: true,
    subtree: true,
  });
}
