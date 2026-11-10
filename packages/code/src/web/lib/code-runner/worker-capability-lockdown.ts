/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * In-worker capability lockdown for execution Workers that run authored code.
 *
 * The Ruby and Shell Workers are module Workers bundled under /_next/static,
 * which the proxy never sees, so no per-path CSP can fence them (the C#, Lua,
 * SQL and Python workers carry one; see proxy.ts). They run on the app origin
 * with everything the creating page could reach: the app's /api with the
 * visitor's cookies, any other origin, the origin's storage and the app's other
 * browsing contexts. Their runtimes give authored code no JS bridge; this is
 * the structural second layer should one ever appear.
 *
 * Call it once, after the runtime has finished loading (its WASM fetched,
 * verified, compiled and instantiated, every chunk imported) and before the
 * first request is accepted. Each capability below is replaced by an
 * undefined, non-writable, non-configurable property on the global AND on
 * every object of its prototype chain, so it can be neither called, restored
 * through assignment or defineProperty, nor recovered from an interface
 * prototype. A warm Worker never needs them again: a run evaluates inside the
 * already-instantiated module.
 *
 * Fails closed: if any capability is still reachable afterwards, it throws,
 * and the Worker reports a load failure instead of accepting authored code.
 *
 * Not removable from inside a Worker: dynamic import() is syntax, not a
 * binding. Reaching it needs the same JS bridge the runtimes do not expose.
 */

const LOCKED: PropertyDescriptor = Object.freeze({
  value: undefined,
  writable: false,
  enumerable: false,
  configurable: false,
});

/** Globals a Worker could use to leave its own thread. */
export const LOCKED_WORKER_GLOBALS = Object.freeze([
  // Network.
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "WebSocketStream",
  "WebTransport",
  "EventSource",
  "RTCPeerConnection",
  // A font load is a network request, and CSS Font Loading reaches Workers.
  "FontFace",
  "fonts",
  // Script loading.
  "importScripts",
  // A nested Worker would start with a fresh, unlocked global.
  "Worker",
  "SharedWorker",
  // Storage the Worker shares with the app's pages.
  "indexedDB",
  "caches",
  "cookieStore",
  "webkitRequestFileSystem",
  "webkitRequestFileSystemSync",
  "webkitResolveLocalFileSystemURL",
  "webkitResolveLocalFileSystemSyncURL",
  // Messaging with the app's other contexts.
  "BroadcastChannel",
] as const);

/** WorkerNavigator members with the same reach (beacons, service worker, OPFS, origin-wide locks). */
export const LOCKED_WORKER_NAVIGATOR_MEMBERS = Object.freeze([
  "sendBeacon",
  "serviceWorker",
  "storage",
  "locks",
] as const);

function seal(target: object, name: string): void {
  for (
    let owner: object | null = target;
    owner !== null;
    owner = Object.getPrototypeOf(owner) as object | null
  ) {
    if (Object.prototype.hasOwnProperty.call(owner, name)) {
      Object.defineProperty(owner, name, LOCKED);
    }
  }
  // Own on the target even where the name never existed, so nothing can
  // define it later.
  if (!Object.prototype.hasOwnProperty.call(target, name)) {
    Object.defineProperty(target, name, LOCKED);
  }
  if (Reflect.get(target, name) !== undefined) {
    throw new Error(`Worker capability "${name}" could not be removed.`);
  }
}

export function lockDownWorkerCapabilities(
  scope: object = globalThis,
): void {
  for (const name of LOCKED_WORKER_GLOBALS) seal(scope, name);
  const navigator: unknown = Reflect.get(scope, "navigator");
  if (navigator && typeof navigator === "object") {
    for (const name of LOCKED_WORKER_NAVIGATOR_MEMBERS) seal(navigator, name);
  }
}
