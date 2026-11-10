/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One ResizeObserver, and one IntersectionObserver per root margin, shared by
 * every element of a long list that watches its own size or its distance to
 * the viewport (the presentation rail and slide sorter: a 300-slide deck made
 * 600 observers, each with its own delivery pass per frame).
 *
 * Each call returns the function that stops watching; the shared observer is
 * disconnected when its last element goes. A call returns null when the
 * browser has no such observer, so the caller can take its own fallback.
 * An observer is created from the global constructor of the moment and is
 * replaced when that constructor changes (a test stubbing it per case).
 */

type Pool<O, E> = {
  ctor: unknown;
  observer: O;
  listeners: Map<Element, Set<(entry: E) => void>>;
};

let resizePool: Pool<ResizeObserver, ResizeObserverEntry> | null = null;
const intersectionPools = new Map<
  string,
  Pool<IntersectionObserver, IntersectionObserverEntry>
>();

function dispatch<E extends { target: Element }>(
  listeners: Map<Element, Set<(entry: E) => void>>,
  entries: readonly E[],
) {
  for (const entry of entries) {
    const forTarget = listeners.get(entry.target);
    if (!forTarget) continue;
    for (const listener of [...forTarget]) listener(entry);
  }
}

type Observer = {
  observe(element: Element): void;
  unobserve?(element: Element): void;
  disconnect(): void;
};

function subscribe<O extends Observer, E>(
  pool: Pool<O, E>,
  element: Element,
  listener: (entry: E) => void,
  onEmpty: () => void,
): () => void {
  let forTarget = pool.listeners.get(element);
  if (!forTarget) {
    forTarget = new Set();
    pool.listeners.set(element, forTarget);
    pool.observer.observe(element);
  }
  forTarget.add(listener);
  let active = true;
  return () => {
    if (!active) return;
    active = false;
    const current = pool.listeners.get(element);
    if (!current) return;
    current.delete(listener);
    if (current.size > 0) return;
    pool.listeners.delete(element);
    pool.observer.unobserve?.(element);
    if (pool.listeners.size === 0) {
      pool.observer.disconnect();
      onEmpty();
    }
  };
}

/** Calls `listener` whenever the element's box changes size. */
export function observeElementResize(
  element: Element,
  listener: (entry: ResizeObserverEntry) => void,
): (() => void) | null {
  const Ctor = typeof ResizeObserver === "undefined" ? undefined : ResizeObserver;
  if (!Ctor) return null;
  if (!resizePool || resizePool.ctor !== Ctor) {
    const listeners = new Map<Element, Set<(entry: ResizeObserverEntry) => void>>();
    resizePool = {
      ctor: Ctor,
      listeners,
      observer: new Ctor((entries) => dispatch(listeners, entries)),
    };
  }
  const pool = resizePool;
  return subscribe(pool, element, listener, () => {
    if (resizePool === pool) resizePool = null;
  });
}

/**
 * Calls `listener` with each intersection change of the element against the
 * viewport widened by `rootMargin`.
 */
export function observeElementIntersection(
  element: Element,
  listener: (entry: IntersectionObserverEntry) => void,
  { rootMargin = "0px" }: { rootMargin?: string } = {},
): (() => void) | null {
  const Ctor =
    typeof IntersectionObserver === "undefined" ? undefined : IntersectionObserver;
  if (!Ctor) return null;
  let pool = intersectionPools.get(rootMargin);
  if (!pool || pool.ctor !== Ctor) {
    const listeners = new Map<
      Element,
      Set<(entry: IntersectionObserverEntry) => void>
    >();
    pool = {
      ctor: Ctor,
      listeners,
      observer: new Ctor((entries) => dispatch(listeners, entries), { rootMargin }),
    };
    intersectionPools.set(rootMargin, pool);
  }
  const current = pool;
  return subscribe(current, element, listener, () => {
    if (intersectionPools.get(rootMargin) === current) {
      intersectionPools.delete(rootMargin);
    }
  });
}
