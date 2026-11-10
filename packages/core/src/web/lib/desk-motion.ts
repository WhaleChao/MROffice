/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How the desk's surfaces come and go: canvas windows, Office task panes and
 * the small overlays of the canvas.
 *
 * One grammar everywhere. Entrances decelerate into place, exits are shorter
 * and accelerate away, and a surface that changes its frame (maximize, an
 * edge snap) glides from the old frame to the new one. Only transform and
 * opacity move: nothing that would lay out an editor again is animated. The
 * stylesheets read the same values as custom properties (`--enter-ease`,
 * `--window-exit`, ...) with these numbers as their fallbacks; scripts
 * (Web Animations) read the constants below.
 *
 * Under reduced motion (the system's or Settings > Reduce motion) and in a
 * test DOM nothing plays: `surfaceMotionAllowed` says no, and every surface
 * simply appears or disappears.
 *
 * `surfacePictureOf` is the picture an exit plays after its surface is gone:
 * an inert copy that looks like the surface did at that moment (its scroll
 * positions and typed values included), cheap enough to take on a large
 * page because what is not on screen is copied as an empty box of the same
 * size.
 */
import { exitMotionAllowed } from "./inert-ghost";

export const SURFACE_ENTER_EASE = "cubic-bezier(0.22, 1, 0.36, 1)";
export const SURFACE_EXIT_EASE = "cubic-bezier(0.4, 0, 1, 1)";
export const SURFACE_MOVE_EASE = "cubic-bezier(0.32, 0.72, 0, 1)";

export const WINDOW_ENTER_MS = 240;
export const WINDOW_EXIT_MS = 180;
export const PANEL_ENTER_MS = 300;
export const PANEL_EXIT_MS = 200;
export const SURFACE_MOVE_MS = 280;
/** The canvas's small overlays (snap preview, guides, labels, the map). */
export const OVERLAY_EXIT_MS = 140;

/** Whether surfaces may move here (a real browser, motion not reduced). */
export function surfaceMotionAllowed(): boolean {
  return typeof window !== "undefined" && exitMotionAllowed(window);
}

/** Attributes that make an element findable or operable; a picture has none
 * (the list lib/inert-ghost strips, for the same reasons). */
const IDENTITY_ATTRIBUTES = [
  "id",
  "data-testid",
  "role",
  "aria-modal",
  "aria-labelledby",
  "aria-describedby",
  "data-modal-layer",
  "name",
  "for",
  "tabindex",
  "data-initial-focus",
  "autofocus",
  "contenteditable",
  "draggable",
] as const;

/** A copied frame or player would load its document again. */
const LIVE_MEDIA = new Set(["iframe", "video", "audio", "object", "embed"]);
const SKIPPED = new Set(["script", "noscript", "template"]);

/** A block at least this tall that is wholly off screen is copied hollow. */
const HOLLOW_MIN_HEIGHT_PX = 96;

export interface SurfacePictureOptions {
  /** Further attributes to strip (ids the caller's own lookups use). */
  readonly stripAttributes?: readonly string[];
  /**
   * How many elements are copied with their contents; past it the rest are
   * empty boxes of their own size. A window's title bar comes first in its
   * tree, so even a small budget keeps it.
   */
  readonly budget?: number;
}

export interface SurfacePicture {
  readonly element: HTMLElement;
  /** Elements copied with their contents (the cost of the picture). */
  readonly copied: number;
  /** Put the copy's scrollers where the surface's were. Call once the
   * picture is in the document. */
  settle(): void;
}

interface PictureWork {
  readonly visible: { left: number; top: number; right: number; bottom: number };
  readonly budget: number;
  readonly strip: readonly string[];
  copied: number;
  readonly scrolls: Array<[Element, number, number]>;
}

function stripIdentity(copy: Element, strip: readonly string[]): void {
  for (const attribute of IDENTITY_ATTRIBUTES) copy.removeAttribute(attribute);
  for (const attribute of strip) copy.removeAttribute(attribute);
}

/** An empty box the size the element is drawn at, so what follows it lies
 * where it did. */
function hollow(source: HTMLElement, copy: HTMLElement): void {
  copy.style.boxSizing = "border-box";
  copy.style.width = `${source.offsetWidth}px`;
  copy.style.height = `${source.offsetHeight}px`;
}

function carryValue(source: Element, copy: Element): void {
  if (source instanceof HTMLInputElement && copy instanceof HTMLInputElement) {
    if (source.type === "checkbox" || source.type === "radio") copy.checked = source.checked;
    else if (source.type !== "file") copy.value = source.value;
  } else if (
    (source instanceof HTMLTextAreaElement && copy instanceof HTMLTextAreaElement) ||
    (source instanceof HTMLSelectElement && copy instanceof HTMLSelectElement)
  ) {
    copy.value = source.value;
  }
}

function copyCanvas(source: HTMLCanvasElement, copy: HTMLCanvasElement): void {
  if (source.width === 0 || source.height === 0) return;
  try {
    copy.getContext("2d")?.drawImage(source, 0, 0);
  } catch {
    // A tainted or lost canvas leaves blank.
  }
}

function copyNode(node: Node, work: PictureWork): Node | null {
  if (node.nodeType === Node.TEXT_NODE) return node.cloneNode(false);
  if (node.nodeType !== Node.ELEMENT_NODE) return null;
  const source = node as Element;
  const tag = source.localName;
  if (SKIPPED.has(tag)) return null;
  if (LIVE_MEDIA.has(tag)) {
    const blank = document.createElement("div");
    blank.className = `${source.getAttribute("class") || ""} wm-dock-blank`.trim();
    blank.style.cssText = (source as HTMLElement).style?.cssText ?? "";
    return blank;
  }
  const copy = source.cloneNode(false) as Element;
  stripIdentity(copy, work.strip);
  if (source instanceof HTMLElement && copy instanceof HTMLElement) {
    const height = source.offsetHeight;
    const width = source.offsetWidth;
    // Not drawn at all (display: none): nothing inside it shows either. A
    // box-less wrapper (display: contents) has no rects but its children do.
    if (
      height === 0 &&
      width === 0 &&
      source.getClientRects().length === 0 &&
      window.getComputedStyle(source).display !== "contents"
    ) {
      return copy;
    }
    if (work.copied >= work.budget) {
      hollow(source, copy);
      return copy;
    }
    if (height >= HOLLOW_MIN_HEIGHT_PX) {
      const rect = source.getBoundingClientRect();
      const { visible } = work;
      if (
        rect.bottom <= visible.top ||
        rect.top >= visible.bottom ||
        rect.right <= visible.left ||
        rect.left >= visible.right
      ) {
        hollow(source, copy);
        return copy;
      }
    }
    if (source.scrollTop !== 0 || source.scrollLeft !== 0) {
      work.scrolls.push([copy, source.scrollTop, source.scrollLeft]);
    }
  } else if (work.copied >= work.budget) {
    // Past the budget a drawing (SVG) keeps its own box, empty.
    return copy;
  }
  work.copied += 1;
  carryValue(source, copy);
  if (source instanceof HTMLCanvasElement && copy instanceof HTMLCanvasElement) {
    copyCanvas(source, copy);
    return copy;
  }
  for (let child = source.firstChild; child; child = child.nextSibling) {
    const childCopy = copyNode(child, work);
    if (childCopy) copy.appendChild(childCopy);
  }
  return copy;
}

/**
 * An inert picture of `source` marked with `marker`: no ids, test ids,
 * roles, names or focus stops (focus, the keyboard, assistive technology and
 * every selector lookup see the surface as gone), no live frames or players.
 * What lies wholly outside the surface's visible part (and the viewport) is
 * copied as an empty box of the same size, so a long document costs what is
 * on screen, not what is in the file.
 */
export function surfacePictureOf(
  source: HTMLElement,
  marker: string,
  options: SurfacePictureOptions = {},
): SurfacePicture {
  const rect = source.getBoundingClientRect();
  const viewWidth = window.innerWidth || document.documentElement.clientWidth;
  const viewHeight = window.innerHeight || document.documentElement.clientHeight;
  const work: PictureWork = {
    visible: {
      left: Math.max(rect.left, 0),
      top: Math.max(rect.top, 0),
      right: Math.min(rect.right, viewWidth),
      bottom: Math.min(rect.bottom, viewHeight),
    },
    budget: options.budget ?? Number.POSITIVE_INFINITY,
    strip: options.stripAttributes ?? [],
    copied: 0,
    scrolls: [],
  };
  const element = (copyNode(source, work) ?? source.cloneNode(false)) as HTMLElement;
  element.setAttribute(marker, "");
  element.setAttribute("aria-hidden", "true");
  element.setAttribute("inert", "");
  return {
    element,
    copied: work.copied,
    settle() {
      for (const [copy, top, left] of work.scrolls) {
        copy.scrollTop = top;
        copy.scrollLeft = left;
      }
    },
  };
}

export interface SurfaceFrame {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * A surface whose frame has just changed from `from` to `to` (both in the
 * coordinates its own position is written in) glides there from where it
 * was: it is laid out once, at its new frame, and drawn moving and scaling
 * from the old one (FLIP). The individual `translate` and `scale` properties
 * carry the motion, so whatever `transform` the surface already has (a
 * wheel bounce, a stylesheet's) composes with it instead of being replaced,
 * and the motion ends exactly where the surface is.
 */
export function glideBetweenFrames(
  element: HTMLElement | null,
  from: SurfaceFrame,
  to: SurfaceFrame,
): Animation | null {
  if (!element || typeof element.animate !== "function" || !surfaceMotionAllowed()) {
    return null;
  }
  if (!(to.width > 0 && to.height > 0 && from.width > 0 && from.height > 0)) return null;
  const dx = from.x + from.width / 2 - (to.x + to.width / 2);
  const dy = from.y + from.height / 2 - (to.y + to.height / 2);
  const sx = from.width / to.width;
  const sy = from.height / to.height;
  if (
    Math.abs(dx) < 0.5 &&
    Math.abs(dy) < 0.5 &&
    Math.abs(sx - 1) < 0.002 &&
    Math.abs(sy - 1) < 0.002
  ) {
    return null;
  }
  try {
    return element.animate(
      [
        {
          transformOrigin: "50% 50%",
          translate: `${dx}px ${dy}px`,
          scale: `${sx} ${sy}`,
        },
        { transformOrigin: "50% 50%", translate: "0px 0px", scale: "1 1" },
      ],
      { duration: SURFACE_MOVE_MS, easing: SURFACE_MOVE_EASE },
    );
  } catch {
    // An engine without Web Animations shows the new frame at once.
    return null;
  }
}

/** A loading card (components/ui/editor-loading-state). */
const LOADING_CARD_SELECTOR = "[data-editor-loading-kind]";
/** On a container whose content is fading in: "playing", then "done". */
export const CONTENT_REVEAL_ATTRIBUTE = "data-content-reveal";
const CONTENT_REVEAL_ANIMATION = "surface-content-reveal";
/** A card that shows up this soon after the container does still counts. */
const CONTENT_REVEAL_WATCH_MS = 1500;
/** A card standing for the content covers at least this much of it (a
 * pane's own small loading card inside loaded content does not count). */
const COVERING_SHARE = 0.4;

/** The border-box area a resize observation reports. */
function observedArea(entry: ResizeObserverEntry): number {
  const box = entry.borderBoxSize?.[0];
  if (box) return box.inlineSize * box.blockSize;
  return entry.contentRect.width * entry.contentRect.height;
}






















export function revealContentWhenLoaded(container: HTMLElement): () => void {
  if (
    !surfaceMotionAllowed() ||
    typeof MutationObserver === "undefined" ||
    typeof ResizeObserver === "undefined"
  ) {
    return () => {};
  }
  let sawCard = false;
  let stopped = false;
  let timer = 0;
  let finishTimer = 0;
  let fallbackFrame = 0;
  let containerArea: number | null = null;
  const cardAreas = new WeakMap<Element, number>();
  const observed = new WeakSet<Element>();
  const onEnd = (event: AnimationEvent) => {
    if (event.animationName === CONTENT_REVEAL_ANIMATION) finish();
  };
  const finish = () => {
    window.clearTimeout(finishTimer);
    container.removeEventListener("animationend", onEnd);
    if (container.getAttribute(CONTENT_REVEAL_ATTRIBUTE) === "playing") {
      // Kept (as "done"): the editors' own mount fades stay off inside it.
      container.setAttribute(CONTENT_REVEAL_ATTRIBUTE, "done");
    }
  };
  /** Whether a card stands for the content; undefined while a size is not
   * reported yet. A container with no area counts as covered. */
  const covers = (card: Element): boolean | undefined => {
    const area = cardAreas.get(card);
    if (area === undefined || containerArea === null) return undefined;
    return !(containerArea > 0) || area >= containerArea * COVERING_SHARE;
  };
  const stop = () => {
    stopped = true;
    observer.disconnect();
    sizes.disconnect();
    window.clearTimeout(timer);
    if (fallbackFrame) window.cancelAnimationFrame(fallbackFrame);
    fallbackFrame = 0;
  };
  const reveal = () => {
    stop();
    container.setAttribute(CONTENT_REVEAL_ATTRIBUTE, "playing");
    container.addEventListener("animationend", onEnd);
    finishTimer = window.setTimeout(finish, 1000);
  };
  const evaluate = () => {
    if (stopped) return;
    const cards = container.querySelectorAll<HTMLElement>(LOADING_CARD_SELECTOR);
    let unknown = false;
    for (const card of Array.from(cards)) {
      if (!observed.has(card)) {
        observed.add(card);
        sizes.observe(card);
      }
      const verdict = covers(card);
      if (verdict === true) {
        sawCard = true;
        return;
      }
      if (verdict === undefined) unknown = true;
    }
    if (unknown) {
      // Reported with the next rendering step. A card the observer never
      // reports (no box at all) is read once, a frame on.
      if (!fallbackFrame) {
        fallbackFrame = window.requestAnimationFrame(() => {
          fallbackFrame = window.requestAnimationFrame(() => {
            fallbackFrame = 0;
            if (stopped) return;
            const box = container.getBoundingClientRect();
            if (containerArea === null) containerArea = box.width * box.height;
            for (const card of Array.from(
              container.querySelectorAll<HTMLElement>(LOADING_CARD_SELECTOR),
            )) {
              if (cardAreas.has(card)) continue;
              const rect = card.getBoundingClientRect();
              cardAreas.set(card, rect.width * rect.height);
            }
            evaluate();
          });
        });
      }
      return;
    }
    if (!sawCard) return;
    reveal();
  };
  const sizes = new ResizeObserver((entries) => {
    for (const entry of entries) {
      if (entry.target === container) containerArea = observedArea(entry);
      else cardAreas.set(entry.target, observedArea(entry));
    }
    evaluate();
  });
  const observer = new MutationObserver(evaluate);
  sizes.observe(container);
  observer.observe(container, { childList: true, subtree: true });
  evaluate();
  timer = window.setTimeout(() => {
    if (!sawCard) stop();
  }, CONTENT_REVEAL_WATCH_MS);
  return () => {
    stop();
    finish();
  };
}

/** Whether `element` is drawn somewhere inside the viewport. */
export function onScreen(element: Element): boolean {
  const rect = element.getBoundingClientRect();
  if (rect.width === 0 || rect.height === 0) return false;
  return (
    rect.right > 0 &&
    rect.bottom > 0 &&
    rect.left < (window.innerWidth || document.documentElement.clientWidth) &&
    rect.top < (window.innerHeight || document.documentElement.clientHeight)
  );
}
