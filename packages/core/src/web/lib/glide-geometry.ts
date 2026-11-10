/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface GlideBox {
  /** From the container's padding-box origin, in the container's CSS pixels. */
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** From the container's padding-box right and bottom edges (CSS `right`/`bottom`). */
  readonly right: number;
  readonly bottom: number;
}

/** Marks the rulers, so a layout watcher does not take them for items. */
const RULER_ATTRIBUTE = "data-glide-ruler";
const RULER_LENGTH = 100;

function ruler(container: HTMLElement, at: "origin" | "moved" | "corner"): HTMLElement {
  const span = container.ownerDocument.createElement("span");
  span.setAttribute(RULER_ATTRIBUTE, "");
  span.setAttribute("aria-hidden", "true");
  span.style.cssText =
    (at === "corner" ? "position:absolute;right:0;bottom:0;" : "position:absolute;top:0;left:0;") +
    "margin:0;padding:0;border:0;" +
    `width:${RULER_LENGTH}px;height:${RULER_LENGTH}px;` +
    "visibility:hidden;pointer-events:none;contain:strict;" +
    `--glide-ruler:${RULER_LENGTH}px;` +
    (at === "moved" ? "transform:translate(var(--glide-ruler),var(--glide-ruler));" : "transform:none;");
  return span;
}

const sane = (value: number, fallback: number) =>
  Number.isFinite(value) && value > 0 ? value : fallback;

interface GlideFrame {
  /** Where the container's padding-box origin and far corner are drawn. */
  readonly originX: number;
  readonly originY: number;
  readonly endX: number;
  readonly endY: number;
  /** Drawn pixels per CSS pixel, along a transform and along a length. */
  readonly stepX: number;
  readonly stepY: number;
  readonly sizeX: number;
  readonly sizeY: number;
}

function frameOf(container: HTMLElement): GlideFrame {
  const still = ruler(container, "origin");
  const moved = ruler(container, "moved");
  const corner = ruler(container, "corner");
  container.append(still, moved, corner);
  const origin = still.getBoundingClientRect();
  const shifted = moved.getBoundingClientRect();
  const end = corner.getBoundingClientRect();
  still.remove();
  moved.remove();
  corner.remove();
  const sizeX = sane(origin.width / RULER_LENGTH, 1);
  const sizeY = sane(origin.height / RULER_LENGTH, sizeX);
  return {
    originX: origin.left,
    originY: origin.top,
    endX: end.right,
    endY: end.bottom,
    sizeX,
    sizeY,
    stepX: sane((shifted.left - origin.left) / RULER_LENGTH, sizeX),
    stepY: sane((shifted.top - origin.top) / RULER_LENGTH, sizeY),
  };
}

/**
 * The item's box inside `container`, in the units an absolutely positioned
 * child of the container is placed in (its CSS pixels from the padding-box
 * origin; a scrolled container's content moves with that origin). The
 * container must be its children's containing block (positioned).
 */
export function boxWithin(item: Element, container: HTMLElement): GlideBox {
  const frame = frameOf(container);
  const inner = item.getBoundingClientRect();
  return {
    x: (inner.left - frame.originX) / frame.stepX,
    y: (inner.top - frame.originY) / frame.stepY,
    width: inner.width / frame.sizeX,
    height: inner.height / frame.sizeY,
    right: (frame.endX - inner.right) / frame.stepX,
    bottom: (frame.endY - inner.bottom) / frame.stepY,
  };
}

/**
 * A point on screen (client coordinates, a rect's edge) as `left`/`top` of
 * an absolutely positioned child of `container`, and the container's own
 * padding-box size, in its CSS pixels.
 */
export function pointWithin(
  container: HTMLElement,
  clientX: number,
  clientY: number,
): { x: number; y: number; width: number; height: number } {
  const frame = frameOf(container);
  return {
    x: (clientX - frame.originX) / frame.stepX,
    y: (clientY - frame.originY) / frame.stepY,
    width: (frame.endX - frame.originX) / frame.sizeX,
    height: (frame.endY - frame.originY) / frame.sizeY,
  };
}

/** Screen (client) coordinates and lengths in `container`'s child units. */
export interface GlidePlacement {
  x(clientX: number): number;
  y(clientY: number): number;
  width(screenWidth: number): number;
  height(screenHeight: number): number;
}

/**
 * `pointWithin` read once, for a caller that places several marks in one
 * frame: where an absolutely positioned child of `container` goes to cover
 * a box measured on screen, whatever lies between the container and the
 * screen (a fixed layer that the browser draws away from the client
 * origin, a zoom, a page zoom).
 */
export function placementWithin(container: HTMLElement): GlidePlacement {
  const frame = frameOf(container);
  return {
    x: (clientX) => (clientX - frame.originX) / frame.stepX,
    y: (clientY) => (clientY - frame.originY) / frame.stepY,
    width: (screenWidth) => screenWidth / frame.sizeX,
    height: (screenHeight) => screenHeight / frame.sizeY,
  };
}

/**
 * How many screen pixels (client coordinates, a rect's edge) one CSS pixel
 * of `container`'s children spans, across and down: what a distance the
 * pointer travelled or two rects differ by is divided by before it is
 * written back as a length or a translate inside the container. Holds
 * whatever zoom, page zoom or transform lies between it and the screen.
 */
export function scaleWithin(container: HTMLElement): { x: number; y: number } {
  const frame = frameOf(container);
  return { x: frame.stepX, y: frame.stepY };
}

/**
 * What the engine's offsetLeft/offsetTop answers for children of `container`
 * are multiplied by: 1 where they come in CSS pixels; the outer zoom where
 * WebKit answers them already zoomed (a zoom inside a zoom: Touch mode's
 * region in a page zoomed with ⌘+). Only positions are affected there;
 * offsetWidth/Height, client sizes and scroll offsets stay CSS pixels.
 * Divide an offsetLeft/offsetTop by it before writing it back as a length.
 * Read off a ruler placed at a known offset in the container, which must be
 * its children's offsetParent (positioned).
 */
export function offsetScaleOf(container: HTMLElement): number {
  const probe = ruler(container, "origin");
  probe.style.top = `${RULER_LENGTH}px`;
  container.append(probe);
  const answered = probe.offsetParent === container ? probe.offsetTop : RULER_LENGTH;
  probe.remove();
  return sane(answered / RULER_LENGTH, 1);
}

/**
 * Calls `onChange` (at most once a frame) whenever the items of `container`
 * may have moved: the container or any of its element children resizing,
 * the window or the visual viewport resizing, the browser zoom or display
 * density changing, fonts arriving, Touch mode switching. Returns the stop.
 */
export function watchGlideLayout(container: HTMLElement, onChange: () => void): () => void {
  if (typeof window === "undefined") return () => undefined;
  let frame = 0;
  let stopped = false;
  const schedule = () => {
    if (stopped || frame) return;
    frame = window.requestAnimationFrame(() => {
      frame = 0;
      if (!stopped) onChange();
    });
  };
  const cleanups: Array<() => void> = [];

  if (typeof ResizeObserver === "function") {
    const resized = new ResizeObserver(schedule);
    resized.observe(container);
    for (const child of Array.from(container.children)) resized.observe(child);
    // items added later (a tier that appears) are observed too; the
    // rulers `boxWithin` puts in for a moment are not items
    const isRuler = (node: Node) => node instanceof Element && node.hasAttribute(RULER_ATTRIBUTE);
    const added = typeof MutationObserver === "function"
      ? new MutationObserver((records) => {
          let moved = false;
          for (const record of records) {
            record.addedNodes.forEach((node) => {
              if (isRuler(node)) return;
              moved = true;
              if (node instanceof Element && node.isConnected) resized.observe(node);
            });
            record.removedNodes.forEach((node) => {
              if (!isRuler(node)) moved = true;
            });
          }
          if (moved) schedule();
        })
      : null;
    added?.observe(container, { childList: true });
    cleanups.push(() => {
      resized.disconnect();
      added?.disconnect();
    });
  }

  window.addEventListener("resize", schedule);
  cleanups.push(() => window.removeEventListener("resize", schedule));
  const viewport = window.visualViewport;
  if (viewport) {
    viewport.addEventListener("resize", schedule);
    cleanups.push(() => viewport.removeEventListener("resize", schedule));
  }

  // browser zoom and moving to another display change the density; the
  // query is re-armed at the new value each time it fires
  let density: MediaQueryList | null = null;
  const onDensity = () => {
    schedule();
    armDensity();
  };
  const armDensity = () => {
    density?.removeEventListener("change", onDensity);
    try {
      density = window.matchMedia(`(resolution: ${window.devicePixelRatio}dppx)`);
      density.addEventListener("change", onDensity);
    } catch {
      density = null;
    }
  };
  armDensity();
  cleanups.push(() => density?.removeEventListener("change", onDensity));

  const fonts = (document as Document & { fonts?: FontFaceSet }).fonts;
  if (fonts && typeof fonts.addEventListener === "function") {
    fonts.addEventListener("loadingdone", schedule);
    cleanups.push(() => fonts.removeEventListener("loadingdone", schedule));
  }

  if (typeof MutationObserver === "function") {
    const mode = new MutationObserver(schedule);
    mode.observe(document.documentElement, { attributes: true, attributeFilter: ["data-touch-mode", "style"] });
    cleanups.push(() => mode.disconnect());
  }

  return () => {
    stopped = true;
    if (frame) window.cancelAnimationFrame(frame);
    for (const cleanup of cleanups) cleanup();
  };
}
