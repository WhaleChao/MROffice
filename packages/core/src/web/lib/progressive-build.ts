/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
interface Entry {
  element: Element;
  build: () => void;
  /** The element's nearest scrolling ancestor, found on its first turn. */
  scroller?: Element | null;
}

const queued = new Set<Entry>();
let frame = 0;

function scrollerOf(element: Element): Element | null {
  const view = element.ownerDocument.defaultView;
  if (!view) return null;
  for (let node = element.parentElement; node; node = node.parentElement) {
    const overflow = view.getComputedStyle(node).overflowY;
    if (overflow === "auto" || overflow === "scroll") return node;
  }
  return null;
}

/** How far the element's box lies outside the visible part of its scroller
 * (0 when any of it is in view). */
function distanceFromView(entry: Entry): number {
  const view = entry.element.ownerDocument.defaultView;
  const box = entry.element.getBoundingClientRect();
  let top = 0;
  let left = 0;
  let bottom = view?.innerHeight ?? 0;
  let right = view?.innerWidth ?? 0;
  if (entry.scroller === undefined) entry.scroller = scrollerOf(entry.element);
  if (entry.scroller?.isConnected) {
    const clip = entry.scroller.getBoundingClientRect();
    top = Math.max(top, clip.top);
    left = Math.max(left, clip.left);
    bottom = Math.min(bottom, clip.bottom);
    right = Math.min(right, clip.right);
  }
  const dy = box.bottom < top ? top - box.bottom : box.top > bottom ? box.top - bottom : 0;
  const dx = box.right < left ? left - box.right : box.left > right ? box.left - right : 0;
  return dy + dx;
}

/** Items already in view built in one frame at most: a screenful of small
 * items (rail tiles) waits for no later frame, and a burst stays bounded. */
const MAX_IN_VIEW_PER_FRAME = 8;
const MIN_IN_VIEW_PER_FRAME = 2;
/** A frame that came this long after the last turn was held by that turn's
 * building: the next turn builds half as many; one this quick, one more. */
const HELD_FRAME_MS = 50;
const QUICK_FRAME_MS = 25;




let inViewPerFrame = MAX_IN_VIEW_PER_FRAME;
let lastTurnAt = 0;

function run(now?: number) {
  frame = 0;
  const at = typeof now === "number" ? now : performance.now();
  if (lastTurnAt > 0) {
    const gap = at - lastTurnAt;
    if (gap > HELD_FRAME_MS) inViewPerFrame = Math.max(MIN_IN_VIEW_PER_FRAME, Math.floor(inViewPerFrame / 2));
    else if (gap < QUICK_FRAME_MS) inViewPerFrame = Math.min(MAX_IN_VIEW_PER_FRAME, inViewPerFrame + 1);
  }
  const inView: Entry[] = [];
  let nearest: Entry | null = null;
  let nearestDistance = Infinity;
  for (const entry of queued) {
    if (!entry.element.isConnected) {
      queued.delete(entry);
      continue;
    }
    const distance = distanceFromView(entry);
    if (distance === 0) {
      inView.push(entry);
      if (inView.length >= inViewPerFrame) break;
    } else if (distance < nearestDistance) {
      nearest = entry;
      nearestDistance = distance;
    }
  }
  // What the reader sees now, else the one nearest it.
  const turn = inView.length > 0 ? inView : nearest ? [nearest] : [];
  for (const entry of turn) {
    queued.delete(entry);
    entry.build();
  }
  if (queued.size > 0) {
    lastTurnAt = turn.length > 0 ? at : 0;
    frame = window.requestAnimationFrame(run);
  } else {
    // A new burst starts from the quick allowance.
    lastTurnAt = 0;
    inViewPerFrame = MAX_IN_VIEW_PER_FRAME;
  }
}

/**
 * Queues `build` for `element`'s turn (what is in view first, then the
 * nearest one per frame). Returns the function that withdraws it if it has not run yet.
 * Without animation frames (a test environment) it builds at once.
 */
export function queueProgressiveBuild(element: Element, build: () => void): () => void {
  if (typeof window === "undefined" || typeof window.requestAnimationFrame !== "function") {
    build();
    return () => {};
  }
  const entry: Entry = { element, build };
  queued.add(entry);
  if (!frame) frame = window.requestAnimationFrame(run);
  return () => {
    queued.delete(entry);
    if (queued.size === 0 && frame) {
      window.cancelAnimationFrame(frame);
      frame = 0;
    }
  };
}
