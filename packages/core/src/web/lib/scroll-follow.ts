"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef, type RefObject } from "react";














/** The box has not scrolled for this long (ms): its glide is over. */
export const SCROLL_STILL_MS = 150;

const movedAt = new WeakMap<Element, number>();
/** Where a move made here left a box, and when. */
const placedBy = new WeakMap<Element, { top: number; at: number }>();
/** What each finger down landed on. */
const touched = new Map<number, Element>();
/** What the wheel last turned over, and when: its scroll has not reached script yet. */
let wheeled: { target: Element; at: number } | null = null;
/** When the person last did what can move a box up: pressed, turned the wheel up, a key, a finger. */
let handledAt = Number.NEGATIVE_INFINITY;
let pressed = false;
/** How long after that a box's top falling is still their doing. */
const HAND_ON_MS = 600;
const MOVING_KEYS = new Set(["ArrowUp", "PageUp", "Home", " ", "Tab"]);
const pending = new WeakMap<Element, Map<unknown, () => void>>();
const timers = new WeakMap<Element, ReturnType<typeof setTimeout>>();
let watching = false;

function watch(): void {
  if (watching || typeof document === "undefined") return;
  watching = true;
  const options: AddEventListenerOptions = { capture: true, passive: true };
  document.addEventListener(
    "touchstart",
    (event) => {
      for (const touch of Array.from(event.changedTouches)) {
        if (touch.target instanceof Element) touched.set(touch.identifier, touch.target);
      }
    },
    options,
  );
  const lift = (event: TouchEvent) => {
    handledAt = performance.now();
    for (const touch of Array.from(event.changedTouches)) touched.delete(touch.identifier);
    // A lift that never reached here (what it landed on left the page).
    if (event.touches.length === 0) touched.clear();
  };
  document.addEventListener("touchend", lift, options);
  document.addEventListener("touchcancel", lift, options);
  document.addEventListener(
    "wheel",
    (event) => {
      const now = performance.now();
      if (event.target instanceof Element) wheeled = { target: event.target, at: now };
      if (event.deltaY < 0) handledAt = now;
    },
    options,
  );
  const press = (down: boolean) => () => {
    pressed = down;
    handledAt = performance.now();
  };
  document.addEventListener("pointerdown", press(true), options);
  document.addEventListener("pointerup", press(false), options);
  document.addEventListener("pointercancel", press(false), options);
  document.addEventListener(
    "keydown",
    (event) => {
      // The key first: `isContentEditable` is a computed style, which WebKit
      // brings up to date to answer -- a style recalculation of a long
      // document's pages before every typed character (R286-D01 in WebKit:
      // this listener's own time was the largest of a key's), for a key that
      // moves nothing. `window.__docsScrollFollowKeyFirst = false` reads it first again (A/B).
      if ((window as Window & { __docsScrollFollowKeyFirst?: boolean }).__docsScrollFollowKeyFirst !== false && !MOVING_KEYS.has(event.key)) return;
      const target = event.target;
      const typing =
        target instanceof HTMLElement &&
        (target.isContentEditable || target.localName === "input" || target.localName === "textarea");
      if (!typing && MOVING_KEYS.has(event.key)) handledAt = performance.now();
    },
    options,
  );
  document.addEventListener(
    "scroll",
    (event) => {
      const box = event.target;
      if (!(box instanceof Element)) return;
      const now = performance.now();
      if (placedHere(box, now)) return;
      movedAt.set(box, now);
    },
    options,
  );
}

/** Whether `box` stands where a move made here just left it. */
function placedHere(box: Element, now: number): boolean {
  const placed = placedBy.get(box);
  return Boolean(placed && Math.abs(box.scrollTop - placed.top) < 1 && now - placed.at < 250);
}

if (typeof document !== "undefined") watch();

/** Whether a finger is on `box`, the wheel just turned over it, or it scrolled (not by a move made here) within the last moment. */
export function scrollInMotion(box: Element, now: number = performance.now()): boolean {
  watch();
  for (const target of touched.values()) if (box.contains(target)) return true;
  if (wheeled && now - wheeled.at < SCROLL_STILL_MS && box.contains(wheeled.target)) return true;
  const at = movedAt.get(box);
  return at !== undefined && now - at < SCROLL_STILL_MS;
}

/** Scroll `box` to `top` from script, as a move the reader did not make. */
export function placeScrollTop(box: HTMLElement, top: number): void {
  watch();
  box.scrollTop = top;
  placedBy.set(box, { top: box.scrollTop, at: performance.now() });
}

function runPending(box: Element): void {
  timers.delete(box);
  const moves = pending.get(box);
  if (!moves || moves.size === 0) return;
  if (scrollInMotion(box)) {
    timers.set(box, setTimeout(() => runPending(box), SCROLL_STILL_MS));
    return;
  }
  pending.delete(box);
  for (const move of moves.values()) move();
}

/**
 * Run `move` (a script scroll of `box`) now if the box is still, else once
 * it is. A later move with the same `key` replaces one still waiting. The
 * move decides again when it runs (the reader may have left the end).
 * Returns a cancel for a move still waiting.
 */
export function moveScrollWhenStill(box: Element, move: () => void, key: unknown = move): () => void {
  if (!scrollInMotion(box) && !pending.get(box)?.size) {
    move();
    return () => {};
  }
  let moves = pending.get(box);
  if (!moves) {
    moves = new Map();
    pending.set(box, moves);
  }
  moves.set(key, move);
  if (!timers.has(box)) timers.set(box, setTimeout(() => runPending(box), SCROLL_STILL_MS));
  return () => {
    const current = pending.get(box);
    if (current?.get(key) === move) current.delete(key);
  };
}




























interface EndFollower {
  /** Whether the reader is with the end. */
  following(): boolean;
  /** The reader asked for the end. */
  follow(smooth?: boolean): void;
  /** The box's height may have changed: keep the reader where they are. */
  keep(): void;
  /** While false nothing is moved (a live log closed); the reader is still read. */
  setEnabled(enabled: boolean): void;
  holdPlace(): void;
  release(): void;
}

/** The element at the box's top edge and the blocks it stands in, deepest first, each with its offset from that edge. */
type ReadingPlace = { element: Element; offset: number }[];

/** How deep a place is looked for: a paragraph, not a word in it. */
const PLACE_DEPTH = 9;
/** How soon after the content changed a fallen top can be that change's layout. */
const LAYOUT_FAULT_MS = 800;

function readingPlace(box: HTMLElement): ReadingPlace {
  const edge = box.getBoundingClientRect().top;
  const place: ReadingPlace = [];
  let parent: Element = box;
  for (let depth = 0; depth < PLACE_DEPTH; depth += 1) {
    let found: { element: Element; top: number } | null = null;
    for (const child of Array.from(parent.children)) {
      const rect = child.getBoundingClientRect();
      if (rect.height > 0 && rect.bottom > edge) {
        found = { element: child, top: rect.top };
        break;
      }
    }
    if (!found) break;
    place.unshift({ element: found.element, offset: found.top - edge });
    // A drawing is written again whole on every token: its box is the place.
    if (found.top >= edge || found.element.localName === "svg") break;
    parent = found.element;
  }
  return place;
}

const followers = new WeakMap<HTMLElement, EndFollower>();
const users = new WeakMap<EndFollower, number>();

/** The box's follower, shared by whatever moves it. Release what you acquire. */
export function acquireEndFollower(
  box: HTMLElement,
  { slack = 48, holdPlace = false }: { slack?: number; holdPlace?: boolean } = {},
): EndFollower {
  const existing = followers.get(box);
  if (existing) {
    users.set(existing, (users.get(existing) ?? 0) + 1);
    if (holdPlace) existing.holdPlace();
    return existing;
  }
  watch();
  let following = box.scrollHeight - box.scrollTop - box.clientHeight <= slack;
  let enabled = true;
  let lastTop = box.scrollTop;
  let lastHeight = box.scrollHeight;
  let holding = false;
  let place: ReadingPlace = [];
  let stopMove: () => void = () => {};
  let changedAt = Number.NEGATIVE_INFINITY;
  const moveKey = {};

  const note = () => {
    lastTop = box.scrollTop;
    lastHeight = box.scrollHeight;
  };
  const toEnd = () => {
    stopMove();
    stopMove = moveScrollWhenStill(
      box,
      () => {
        if (!following || !enabled) return;
        const end = Math.max(0, box.scrollHeight - box.clientHeight);
        if (Math.abs(box.scrollTop - end) >= 1) placeScrollTop(box, end);
        note();
      },
      moveKey,
    );
  };
  const backToPlace = () => {
    if (!holding) return;
    // The reader is moving: where they are is theirs, read at the next scroll.
    if (scrollInMotion(box)) return;
    const standing = place.find(({ element }) => element.isConnected && box.contains(element));
    if (standing) {
      const moved =
        standing.element.getBoundingClientRect().top - box.getBoundingClientRect().top - standing.offset;
      if (Math.abs(moved) >= 1) {
        placeScrollTop(box, box.scrollTop + moved);
        note();
      }
    }
    place = readingPlace(box);
  };
  const keep = () => {
    if (following) toEnd();
    else backToPlace();
    // A box that shrank took its top along by as much; anything lower than
    // that is a scroll whose event has not arrived yet, or layout's (onScroll).
    const height = box.scrollHeight;
    if (height < lastHeight - 0.5) lastTop = Math.max(box.scrollTop, lastTop - (lastHeight - height));
    lastHeight = height;
  };

  const onScroll = () => {
    const top = box.scrollTop;
    const height = box.scrollHeight;
    const now = performance.now();
    if (!placedHere(box, now)) {
      const distance = height - top - box.clientHeight;
      // Where the top stands if only the content's shrinking moved it.
      const expected = lastTop - Math.max(0, lastHeight - height);
      if (top > lastTop + 0.5) {
        if (distance <= slack) following = true;
      } else if (top < expected - 0.5 && distance > 1) {
        // (A top falling at the very end is a rubber band settling.)
        const handOn = pressed || touched.size > 0 || now - handledAt < HAND_ON_MS;
        if (enabled && !handOn && now - changedAt < LAYOUT_FAULT_MS) {
          movedAt.delete(box);
          placeScrollTop(box, following ? Math.max(0, height - box.clientHeight) : expected);
          note();
          if (holding && !following) place = readingPlace(box);
          return;
        }
        following = false;
        stopMove();
      }
    }
    lastTop = top;
    lastHeight = height;
    if (holding && !following) place = readingPlace(box);
  };
  box.addEventListener("scroll", onScroll, { passive: true });

  // A resize observer runs after layout and before paint: the frame that draws
  // the taller content also draws its end.
  const sizes = typeof ResizeObserver === "function" ? new ResizeObserver(keep) : null;
  const watchChildren = () => {
    if (!sizes) return;
    sizes.disconnect();
    sizes.observe(box);
    for (const child of Array.from(box.children)) sizes.observe(child);
  };
  const children = new MutationObserver((records) => {
    changedAt = performance.now();
    if (records.some((record) => record.target === box)) {
      watchChildren();
      keep();
    }
  });
  children.observe(box, { childList: true, subtree: true });
  watchChildren();

  const follower: EndFollower = {
    following: () => following,
    follow: (smooth = false) => {
      following = true;
      if (smooth) {
        // Moves wait for the glide; the last of them lands on the true end.
        movedAt.set(box, performance.now());
        box.scrollTo?.({ top: box.scrollHeight - box.clientHeight, behavior: "smooth" });
      }
      toEnd();
    },
    keep,
    setEnabled: (next) => {
      enabled = next;
    },
    holdPlace: () => {
      holding = true;
      box.style.overflowAnchor = "none";
    },
    release: () => {
      const left = (users.get(follower) ?? 1) - 1;
      users.set(follower, left);
      if (left > 0) return;
      followers.delete(box);
      stopMove();
      box.removeEventListener("scroll", onScroll);
      sizes?.disconnect();
      children.disconnect();
      if (holding) box.style.overflowAnchor = "";
    },
  };
  if (holdPlace) follower.holdPlace();
  users.set(follower, 1);
  followers.set(box, follower);
  return follower;
}

/** The reader asked for the end of `box` (sent a message, pressed ↓): followed again from here. */
export function followScrollEnd(box: HTMLElement | null, { smooth = false }: { smooth?: boolean } = {}): void {
  if (box) followers.get(box)?.follow(smooth);
}

/**
 * A transcript (or a live log) that keeps its newest line in view while its
 * reader is at the end, and leaves a reader who scrolled up where they are
 * (EndFollower). `content` changes with what may have grown the box, where
 * nothing can observe its height; a change of `force` (another conversation
 * opened, a message sent) brings the reader back to the end. Moves wait for
 * the box to be still.
 */
export function useFollowScrollEnd(
  scrollRef: RefObject<HTMLElement | null>,
  content: unknown,
  {
    force,
    enabled = true,
    slack = 48,
    holdPlace = false,
  }: { force?: unknown; enabled?: boolean; slack?: number; holdPlace?: boolean } = {},
): void {
  const attached = useRef<{ box: HTMLElement; follower: EndFollower } | null>(null);
  // The box the ref holds now (it may mount after the hook). A box starts at
  // its end, wherever the browser left it.
  useEffect(() => {
    const box = scrollRef.current;
    if (attached.current?.box === box) return;
    attached.current?.follower.release();
    attached.current = null;
    if (!box) return;
    const follower = acquireEndFollower(box, { slack, holdPlace });
    follower.setEnabled(enabled);
    attached.current = { box, follower };
    follower.follow();
  });
  useEffect(
    () => () => {
      attached.current?.follower.release();
      attached.current = null;
    },
    [],
  );
  const forcedRef = useRef(force);
  useEffect(() => {
    const follower = attached.current?.follower;
    if (!follower) return;
    follower.setEnabled(enabled);
    if (forcedRef.current !== force) {
      forcedRef.current = force;
      follower.follow();
    } else {
      follower.keep();
    }
  }, [content, enabled, force]);
}
