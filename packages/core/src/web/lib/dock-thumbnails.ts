/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the side strips' file pictures (components/notebook/page-thumbnails/
 * DockThumbnail) need besides the page thumbnails themselves:
 *
 * - each page's content version, as the notebook's page catalogue last said
 *   it (lib/notebook-page-catalogue), so a picture is keyed by the version it
 *   shows and drawn again once a save moves it on;
 * - one turn at a time to make a picture, each turn begun when the page is
 *   quiet, so a strip of pictures never competes with a cold open or a
 *   window switch for the main thread or the network.
 */
import { useSyncExternalStore } from "react";

import { NOTEBOOK_PAGE_CATALOGUE_EVENT, type PageCatalogueDelta } from "./notebook-page-catalogue";
import { PAGE_REMOTE_SAVED_EVENT } from "./notebook-events";
import { requestQuietCallback } from "./quiet-callback";

/* ------------------------------------------------------ content versions */

const versions = new Map<string, number>();
const listeners = new Set<() => void>();
let listening = false;

/** Pages' content versions, as a page list or the catalogue reads them. A
 *  page's version only ever moves forward. */
export function notePageContentVersions(
  pages: Iterable<{ id: string; content_version?: unknown }>,
): void {
  let moved = false;
  for (const page of pages) {
    const version = page.content_version;
    if (!page.id || typeof version !== "number" || !Number.isSafeInteger(version) || version < 1) continue;
    if ((versions.get(page.id) ?? 0) >= version) continue;
    versions.set(page.id, version);
    moved = true;
  }
  if (moved) for (const listener of listeners) listener();
}

function listen(): void {
  if (listening || typeof window === "undefined") return;
  listening = true;
  window.addEventListener(NOTEBOOK_PAGE_CATALOGUE_EVENT, (event) => {
    const delta = (event as CustomEvent<PageCatalogueDelta>).detail;
    notePageContentVersions(delta?.upserts ?? []);
  });
  window.addEventListener(PAGE_REMOTE_SAVED_EVENT, (event) => {
    const detail = (event as CustomEvent<{ pageId?: string; contentVersion?: number }>).detail;
    if (detail?.pageId) notePageContentVersions([{ id: detail.pageId, content_version: detail.contentVersion }]);
  });
}
// from the first import on: the catalogue's first read comes before any
// strip has a tile to draw
listen();

function subscribe(listener: () => void): () => void {
  listen();
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** The newest content version known for a page (the catalogue's, or the
 *  caller's own row's), or undefined while none is known. */
export function usePageContentVersion(pageId: string, own?: number | null): number | undefined {
  const known = useSyncExternalStore(
    subscribe,
    () => versions.get(pageId),
    () => undefined,
  );
  const mine = typeof own === "number" && own > 0 ? own : undefined;
  if (known === undefined) return mine;
  return mine === undefined ? known : Math.max(known, mine);
}

/* ------------------------------------------------------------------ turns */

/** A turn that never says it is done is over after this long. */
export const DOCK_THUMBNAIL_TURN_MS = 8_000;
/** How long a turn waits for the page to be quiet before it starts anyway. */
const QUIET_WAIT_MS = 3_000;

type Turn = { run: (done: () => void) => void; cancelled: boolean };
const queue: Turn[] = [];
let active: Turn | null = null;
let pump: (() => void) | null = null;

function next(): void {
  if (active || pump) return;
  pump = requestQuietCallback(() => {
    pump = null;
    let turn = queue.shift();
    while (turn?.cancelled) turn = queue.shift();
    if (!turn) return;
    active = turn;
    let over = false;
    const done = () => {
      if (over) return;
      over = true;
      window.clearTimeout(timer);
      if (active === turn) active = null;
      next();
    };
    const timer = window.setTimeout(done, DOCK_THUMBNAIL_TURN_MS);
    try {
      turn.run(done);
    } catch {
      done();
    }
  }, QUIET_WAIT_MS);
}

/**
 * Waits for a quiet moment with no other picture being made, then calls
 * `run`, which calls `done` when its picture is made (or failed). Returns
 * a cancel for a picture no longer wanted (its tile left the strip).
 */
export function requestDockThumbnailTurn(run: (done: () => void) => void): () => void {
  const turn: Turn = { run, cancelled: false };
  queue.push(turn);
  next();
  return () => {
    turn.cancelled = true;
  };
}

export function resetDockThumbnailsForTests(): void {
  versions.clear();
  queue.length = 0;
  active = null;
  pump?.();
  pump = null;
}
