/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { canvasGestureActive } from "./canvas-gesture";
import { createElement, use, type ComponentType, type FunctionComponent } from "react";

// eslint-disable-next-line @typescript-eslint/no-explicit-any -- any component's props
type AnyComponent = ComponentType<any>;

export type PreloadableComponent<C> = (C extends ComponentType<infer P> ? FunctionComponent<P> : never) & {
  preload(): Promise<void>;
};

const OFFLINE_LOAD_RETRY_MS = [1_000, 3_000];
const NOT_HERE = Symbol("part-not-here");

/** A load that failed without a network: asked for again when there is one. What
 * still will not come after that (a bundler may keep the failed chunk for the
 * page's life) is `NOT_HERE`: the part stays out and the window stays up. */
function loadOnceOnline<M>(loader: () => Promise<M>): Promise<M | typeof NOT_HERE> {
  return loader().catch(async (error: unknown) => {
    if (typeof navigator === "undefined" || navigator.onLine !== false) throw error;
    await new Promise<void>((resolve) => window.addEventListener("online", () => resolve(), { once: true }));
    for (const wait of [0, ...OFFLINE_LOAD_RETRY_MS]) {
      if (wait > 0) await new Promise((resolve) => window.setTimeout(resolve, wait));
      try {
        return await loader();
      } catch {
        // goes round again, then the part stays out
      }
    }
    console.warn("[lazy-part] not loaded after the connection came back", error);
    return NOT_HERE;
  });
}

const PartNotHere: FunctionComponent = () => null;

export function lazyNamed<M, K extends keyof M>(loader: () => Promise<M>, name: K): PreloadableComponent<M[K]> {
  let loaded: AnyComponent | null = null;
  let loading: Promise<AnyComponent> | null = null;
  const load = () => {
    if (!loading) {
      loading = loadOnceOnline(loader).then(
        (module) => (loaded = module === NOT_HERE ? PartNotHere : module[name] as unknown as AnyComponent),
        (error: unknown) => {
          loading = null;
          throw error;
        },
      );
    }
    return loading;
  };
  function LazyPart(props: Record<string, unknown>) {
    const component = loaded ?? use(load());
    return createElement(component, props);
  }
  LazyPart.displayName = `Lazy(${String(name)})`;
  const part = LazyPart as unknown as PreloadableComponent<M[K]>;
  part.preload = () => load().then(() => undefined, () => undefined);
  registeredParts.push(part);
  return part;
}

const registeredParts: { preload(): Promise<void> }[] = [];

/** Load every part declared so far (the unit tests' setup: a test draws a
 * Ribbon tab and reads it in the same act, as before the split). */
export function preloadDeclaredParts(): Promise<void> {
  return Promise.all(registeredParts.map((part) => part.preload())).then(() => undefined);
}










export function preloadWhenIdle(
  parts: readonly { preload(): Promise<void> }[],
  delayMs = 3_000,
): () => void {
  if (typeof window === "undefined") return () => undefined;
  let cancelled = false;
  let timer: number | null = null;
  let idle: number | null = null;
  let next = 0;
  let lastInput = 0;
  const noteInput = () => {
    lastInput = performance.now();
  };
  const inputs = ["pointerdown", "keydown", "wheel"] as const;
  for (const type of inputs) window.addEventListener(type, noteInput, { capture: true, passive: true });
  const stop = () => {
    for (const type of inputs) window.removeEventListener(type, noteInput, { capture: true });
  };
  const schedule = (wait = 0) => {
    if (cancelled) return;
    if (next >= parts.length) {
      stop();
      return;
    }
    timer = window.setTimeout(() => {
      timer = null;
      if (typeof window.requestIdleCallback === "function") idle = window.requestIdleCallback(step);
      else step();
    }, wait);
  };
  const step = (deadline?: IdleDeadline) => {
    idle = null;
    if (cancelled) return;
    // A window dragged or sized, the desk zoomed or panned (lib/canvas-gesture)
    // is input too: a part arriving mid-drag is compiled in one task, and
    // in WebKit the frames around it were 0.6-1 s with a deck open.
    if (canvasGestureActive()) noteInput();
    const quietFor = performance.now() - lastInput;
    if (quietFor < PRELOAD_QUIET_MS) {
      schedule(PRELOAD_QUIET_MS - quietFor);
      return;
    }
    if (deadline && !deadline.didTimeout && deadline.timeRemaining() < PRELOAD_IDLE_MS) {
      schedule(50);
      return;
    }
    void parts[next++].preload().finally(() => schedule(50));
  };
  schedule(delayMs);
  return () => {
    cancelled = true;
    stop();
    if (timer !== null) window.clearTimeout(timer);
    if (idle !== null) window.cancelIdleCallback?.(idle);
  };
}

/** How long after the last press or key no part is asked for, and the idle
 * time an idle period must still have for one. */
const PRELOAD_QUIET_MS = 1_200;
const PRELOAD_IDLE_MS = 30;
