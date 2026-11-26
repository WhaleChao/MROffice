/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { afterCanvasGesture, canvasGestureActive } from "./canvas-gesture";
export type FontFaceNews = "loading" | "settled";
type FontFaceNewsListener = (faces: readonly FontFace[], news: FontFaceNews) => void;

interface FontSetWatch {
  /**
   * The set itself, held for as long as its document lives. WebKit drops an
   * event listener when the object it was added to is collected, and nothing
   * else keeps `document.fonts` reachable: after a collection the next
   * `document.fonts` is a new object, this guard would be added again BEHIND
   * the listeners registered meanwhile, and they would hear every event.
   */
  fonts: FontFaceSet;
  /** Faces already told as loading, and as loaded or failed. A face is told
   * once for each: one that reads loaded, then unloaded, then loaded again
   * (a `local()` face in WebKit) is not news the second time. */
  loading: WeakSet<FontFace>;
  settled: WeakSet<FontFace>;
  /** The same by description, for an engine that hands out a new object for
   * a face it has shown before: how many faces of each description were
   * told as loading or settled, and as settled. */
  startedByKey: Map<string, number>;
  settledByKey: Map<string, number>;
  listeners: Set<{ listener: FontFaceNewsListener; loading: boolean }>;
  /** Counts the events that brought a face that had not settled before. */
  settledEvents: number;
}

const fontSetWatches = new WeakMap<Document, FontSetWatch>();
/** The new faces of an event the guard let through. */
const newFacesByEvent = new WeakMap<Event, readonly FontFace[]>();

function facesOf(event: Event): readonly FontFace[] | null {
  const faces = (event as Event & { fontfaces?: readonly FontFace[] }).fontfaces;
  return faces && faces.length > 0 ? faces : null;
}

const faceSettled = (face: FontFace) => face.status === "loaded" || face.status === "error";
/**
 * A face's description, read once per face object: an event the set must be
 * walked for (a description seen before) read the five descriptors of every
 * face in it, a binding call each, for every face landing (a Chinese face's
 * slices landing one by one under a document's first layout: 0.2 s of the
 * metrics wait for allometric.docx with 300 faces). A face's descriptors are
 * never written after it is made. `window.__docsFirstScreenTrimOff = true`:
 * read every time, as before.
 */
const faceKeys = new WeakMap<FontFace, string>();
const describe = (face: FontFace) =>
  `${face.family}|${face.style}|${face.weight}|${face.stretch}|${face.unicodeRange}`;
const faceKey = (face: FontFace) => {
  if (typeof window !== "undefined" && (window as Window & { __docsFirstScreenTrimOff?: boolean }).__docsFirstScreenTrimOff === true) {
    return describe(face);
  }
  let key = faceKeys.get(face);
  if (key === undefined) faceKeys.set(face, (key = describe(face)));
  return key;
};

/** How many faces of each description have started (loading or settled) and
 * have settled, in the set as it stands. */
function countFaces(fonts: FontFaceSet): { started: Map<string, number>; settled: Map<string, number> } {
  const started = new Map<string, number>();
  const settled = new Map<string, number>();
  fonts.forEach((face) => {
    const done = faceSettled(face);
    if (!done && face.status !== "loading") return;
    const key = faceKey(face);
    started.set(key, (started.get(key) ?? 0) + 1);
    if (done) settled.set(key, (settled.get(key) ?? 0) + 1);
  });
  return { started, settled };
}

function watchFontSet(document: Document, fonts: FontFaceSet): FontSetWatch {
  const existing = fontSetWatches.get(document);
  if (existing && existing.fonts === fonts) return existing;
  // What is loaded (or loading) now is not news later.
  const counts = countFaces(fonts);
  const watch: FontSetWatch = {
    fonts,
    loading: new WeakSet(),
    settled: new WeakSet(),
    startedByKey: counts.started,
    settledByKey: counts.settled,
    // A set that was collected and came back keeps its subscribers.
    listeners: existing?.listeners ?? new Set(),
    settledEvents: existing?.settledEvents ?? 0,
  };
  fonts.forEach((face) => {
    if (faceSettled(face)) watch.settled.add(face);
    else if (face.status === "loading") watch.loading.add(face);
  });
  const untold = (face: FontFace, news: FontFaceNews) =>
    news === "settled"
      ? faceSettled(face) && !watch.settled.has(face)
      : face.status === "loading" && !watch.loading.has(face) && !watch.settled.has(face);
  /**
   * The event's news read from the faces it names, without walking the set:
   * in WebKit any walk of a document's font set (`forEach`, `size`) restyles
   * the whole document first, about 650 ms with four editor windows open, and
   * fonts settle while windows open. Reading a face (status, descriptors)
   * costs nothing. Null when only the set can tell: a named face is an object
   * not told before whose description was (a new object for a face told
   * before, or a second face of that description).
   */
  const readNamed = (named: readonly FontFace[], news: FontFaceNews): FontFace[] | null => {
    const told = news === "settled" ? watch.settledByKey : watch.startedByKey;
    const fresh: FontFace[] = [];
    for (const face of named) {
      if (!untold(face, news)) continue;
      if ((told.get(faceKey(face)) ?? 0) > 0) return null;
      fresh.push(face);
    }
    // No description of these was ever told: each one is news, as the set
    // holds at least the face itself.
    const byKey = new Map<string, number>();
    for (const face of fresh) {
      const key = faceKey(face);
      byKey.set(key, (byKey.get(key) ?? 0) + 1);
      if (news === "settled") watch.settled.add(face);
      else watch.loading.add(face);
    }
    for (const [key, count] of byKey) {
      watch.startedByKey.set(key, Math.max(watch.startedByKey.get(key) ?? 0, count));
      if (news === "settled") watch.settledByKey.set(key, Math.max(watch.settledByKey.get(key) ?? 0, count));
    }
    return fresh;
  };
  const read = (event: Event, news: FontFaceNews): FontFace[] => {
    const namedFaces = facesOf(event);
    if (namedFaces) {
      const fresh = readNamed(namedFaces, news);
      if (fresh) return fresh;
    }
    // A face is new when it is an object not told before AND the set holds
    // more faces of its description in that state than were told: the first
    // alone would take a face's new object for a new face.
    const counts = countFaces(fonts);
    const now = news === "settled" ? counts.settled : counts.started;
    const told = news === "settled" ? watch.settledByKey : watch.startedByKey;
    const fresh: FontFace[] = [];
    const visit = (face: FontFace) => {
      if (news === "settled") {
        if (!faceSettled(face) || watch.settled.has(face)) return;
        watch.settled.add(face);
      } else {
        if (face.status !== "loading" || watch.loading.has(face) || watch.settled.has(face)) return;
        watch.loading.add(face);
      }
      const key = faceKey(face);
      if ((now.get(key) ?? 0) > (told.get(key) ?? 0)) fresh.push(face);
    };
    if (namedFaces) namedFaces.forEach(visit);
    else fonts.forEach(visit);
    for (const [key, count] of counts.started) {
      if (count > (watch.startedByKey.get(key) ?? 0)) watch.startedByKey.set(key, count);
    }
    if (news === "settled") {
      for (const [key, count] of counts.settled) {
        if (count > (watch.settledByKey.get(key) ?? 0)) watch.settledByKey.set(key, count);
      }
    }
    return fresh;
  };
  const tell = (faces: readonly FontFace[], news: FontFaceNews) => {
    for (const entry of [...watch.listeners]) {
      if (news === "loading" && !entry.loading) continue;
      entry.listener(faces, news);
    }
  };
  // News that came while a canvas window was dragged or the desk zoomed is
  // told once the gesture ends: every listener lays text out again for it
  // (up to a second of WebKit main thread a face while big files open), and
  // the move under the pointer would wait for all of it (lib/canvas-gesture).
  // Replayed as an event of the same type, so plain listeners hear it too.
  const replayed = new WeakMap<Event, readonly FontFace[]>();
  const held: Array<{ type: string; faces: readonly FontFace[] }> = [];
  const onEvent = (event: Event, news: FontFaceNews) => {
    const again = replayed.get(event);
    const faces = again ?? read(event, news);
    if (faces.length === 0) {
      // Nothing new: the listeners registered after this one do not hear it.
      event.stopImmediatePropagation();
      return;
    }
    if (!again && canvasGestureActive()) {
      event.stopImmediatePropagation();
      if (held.length === 0) {
        afterCanvasGesture(() => {
          for (const { type, faces: told } of held.splice(0)) {
            const replay = new Event(type);
            replayed.set(replay, told);
            fonts.dispatchEvent(replay);
          }
        });
      }
      held.push({ type: event.type, faces });
      return;
    }
    newFacesByEvent.set(event, faces);
    if (news === "settled") watch.settledEvents += 1;
    tell(faces, news);
  };
  fonts.addEventListener("loading", (event) => onEvent(event, "loading"));
  fonts.addEventListener("loadingdone", (event) => onEvent(event, "settled"));
  fonts.addEventListener("loadingerror", (event) => onEvent(event, "settled"));
  fontSetWatches.set(document, watch);
  return watch;
}

function watchOf(document: Document | null | undefined): FontSetWatch | null {
  const fonts: FontFaceSet | undefined = document?.fonts;
  if (!document || !fonts || typeof fonts.addEventListener !== "function" || typeof fonts.forEach !== "function") {
    return null;
  }
  return watchFontSet(document, fonts);
}

/**
 * Start reading the document's font events for news, and keep the events
 * that bring none from the listeners registered from now on. Idempotent.
 */
export function installFontFaceNewsGuard(document: Document | null | undefined): void {
  watchOf(document);
}

/**
 * Call `listener` when faces of the document finish loading or fail for the
 * first time (and, with `loading`, when faces begin to load), with those
 * faces alone; never for an event that reports nothing new. Returns the
 * unsubscribe.
 */
export function onFontFaceNews(
  document: Document | null | undefined,
  listener: FontFaceNewsListener,
  options: { loading?: boolean } = {},
): () => void {
  const watch = watchOf(document);
  if (!watch) return () => {};
  const entry = { listener, loading: options.loading === true };
  watch.listeners.add(entry);
  return () => {
    watch.listeners.delete(entry);
  };
}

/**
 * The faces a font event brought that had not been told before, for a plain
 * `loading` / `loadingdone` / `loadingerror` listener: WebKit's `fontfaces`
 * names every face the set ever loaded. Null for an event the guard did not
 * read (no guard on this set), where `event.fontfaces` is all there is.
 */
export function newFontFacesOf(event: Event | null | undefined): readonly FontFace[] | null {
  return (event && newFacesByEvent.get(event)) ?? null;
}

/** How many times a face of the document settled for the first time: a
 * number that changes exactly when measured text may. Null with no set. */
export function fontFaceNewsCount(document: Document | null | undefined): number | null {
  return watchOf(document)?.settledEvents ?? null;
}

/** Callers waiting on the document's `fonts.ready`, read once for all of them. */
const readyWaiters = new WeakMap<Document, Set<() => void>>();

/**
 * Call `callback` once the document's fonts are ready, as
 * `document.fonts.ready.then(callback)` does, for many callers at once.
 * Blink answers a read of `fonts.ready` by laying the page out first (more
 * loads may start), so each of hundreds of callers reading it in one task
 * laid the whole document out once each: the header and footer bands of the
 * 296-page book, one per page, spent 63 s of its open there. The callers of
 * one task are answered by one read, made after them (a microtask later),
 * so it waits for every face they drew with. Returns the cancel.
 */
export function afterFontsReady(document: Document | null | undefined, callback: () => void): () => void {
  const fonts = document?.fonts;
  if (!document || !fonts) return () => {};
  let waiters = readyWaiters.get(document);
  if (!waiters) {
    const batch = new Set<() => void>();
    waiters = batch;
    readyWaiters.set(document, batch);
    queueMicrotask(() => {
      if (readyWaiters.get(document) === batch) readyWaiters.delete(document);
      if (batch.size === 0) return;
      void fonts.ready.then(() => {
        for (const waiting of Array.from(batch)) waiting();
      }, () => undefined);
    });
  }
  const batch = waiters;
  batch.add(callback);
  return () => {
    batch.delete(callback);
  };
}
