/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * "Open this page there": a place a page is to show when it opens -- a
 * search hit on a handwritten line (the stroke it starts with), a block a
 * link names. Whoever opens the page asks first (`requestPageReveal`); the
 * page's editor takes the request once its document is in (`takePageReveal`)
 * and brings the place into view. A request no editor takes in a minute is
 * dropped, so a page opened later for another reason opens at its start.
 */

export interface PageRevealPlace {
  /** The block to bring into view (a Docs block id: a stroke, a paragraph). */
  readonly blockId: string;
}

const REVEAL_TTL_MS = 60_000;
const EVENT = "mrnote:page-reveal";
const pending = new Map<string, PageRevealPlace & { readonly at: number }>();

export function requestPageReveal(pageId: string, place: PageRevealPlace): void {
  if (!pageId || !place.blockId) return;
  pending.set(pageId, { ...place, at: Date.now() });
  if (typeof window !== "undefined") window.dispatchEvent(new CustomEvent(EVENT, { detail: pageId }));
}

/** The page's place to show, once (null when none is asked or it is stale). */
export function takePageReveal(pageId: string): PageRevealPlace | null {
  const place = pending.get(pageId);
  if (!place) return null;
  pending.delete(pageId);
  return Date.now() - place.at <= REVEAL_TTL_MS ? { blockId: place.blockId } : null;
}

export function hasPageReveal(pageId: string): boolean {
  const place = pending.get(pageId);
  return Boolean(place && Date.now() - place.at <= REVEAL_TTL_MS);
}

/** A page already open is asked again: its editor listens. */
export function subscribePageReveal(listener: (pageId: string) => void): () => void {
  if (typeof window === "undefined") return () => {};
  const onEvent = (event: Event) => listener(String((event as CustomEvent).detail ?? ""));
  window.addEventListener(EVENT, onEvent);
  return () => window.removeEventListener(EVENT, onEvent);
}
