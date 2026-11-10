"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useMemo, useSyncExternalStore } from "react";

import { apiGet } from "@mrnote/core/web/lib/api.ts";
import {
  normalizeCatalogBackground,
  normalizeCatalogCard,
  type CatalogBackground,
  type CatalogCard,
  type SlideMaster,
  type SlideMasterLayout,
} from "@mrnote/core/web/lib/presentation-master.ts";



















interface DrawnCard {
  readonly blocks?: Array<{ id?: string; [key: string]: unknown }>;
  readonly masterGraphics?: Array<{ id?: string; [key: string]: unknown }>;
  readonly background?: CatalogBackground;
  readonly masterBackground?: CatalogBackground;
  readonly catalogCard: CatalogCard;
}

interface PageCards {
  readonly layouts: ReadonlyMap<string, DrawnCard>;
  readonly masters: ReadonlyMap<string, DrawnCard>;
}

const EMPTY: PageCards = { layouts: new Map(), masters: new Map() };
const cardsByPage = new Map<string, PageCards>();
const asked = new Map<string, Set<string>>();
const listeners = new Set<() => void>();
/** The server answers at most this many layout parts at once. */
export const LAYOUT_CARDS_PER_REQUEST = 64;
/** A card drawn without its pictures is asked for again after this long, a few times. */
export const LAYOUT_CARD_PICTURES_RETRY_MS = 4_000;
const LAYOUT_CARD_PICTURES_RETRIES = 5;
const picturesRetries = new Map<string, number>();

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;
}

function blockList(value: unknown): DrawnCard["blocks"] {
  return Array.isArray(value) && value.length > 0
    ? (value.filter((block) => record(block)) as NonNullable<DrawnCard["blocks"]>)
    : undefined;
}

function readCard(value: unknown): DrawnCard | null {
  const card = record(value);
  const catalogCard = normalizeCatalogCard(card?.catalogCard);
  if (!card || !catalogCard || catalogCard === "pending") return null;
  const blocks = blockList(card.blocks);
  const masterGraphics = blockList(card.masterGraphics);
  const background = normalizeCatalogBackground(card.background);
  const masterBackground = normalizeCatalogBackground(card.masterBackground);
  return {
    catalogCard,
    ...(blocks ? { blocks } : {}),
    ...(masterGraphics ? { masterGraphics } : {}),
    ...(background ? { background } : {}),
    ...(masterBackground ? { masterBackground } : {}),
  };
}

function notify(): void {
  for (const listener of listeners) listener();
}

/** The layout parts of `masters` still without their card here. */
export function pendingLayoutCardParts(
  pageId: string,
  masters: readonly SlideMaster[],
): string[] {
  const held = cardsByPage.get(pageId) ?? EMPTY;
  const parts: string[] = [];
  for (const master of masters) {
    for (const layout of master.layouts ?? []) {
      const part = layout.sourcePart ?? layout.id;
      if (layout.catalogCard === "pending" && part.startsWith("ppt/slideLayouts/") && !held.layouts.has(part)) {
        parts.push(part);
      }
    }
    // A master none of whose layouts is used comes with the first of its layouts' cards.
    if (master.catalogCard === "pending" && !held.masters.has(master.sourcePart ?? master.id)) {
      const first = (master.layouts ?? []).map((layout) => layout.sourcePart ?? layout.id)
        .find((part) => part.startsWith("ppt/slideLayouts/"));
      if (first && !parts.includes(first) && !held.layouts.has(first)) parts.push(first);
    }
  }
  return parts;
}

/**
 * Ask the server for the cards of `parts` (each asked once per page), in
 * requests of {@link LAYOUT_CARDS_PER_REQUEST}. A part it cannot draw stays
 * pending: the master view shows its name over an empty card.
 */
export async function requestLayoutCards(pageId: string, parts: readonly string[]): Promise<void> {
  const already = asked.get(pageId) ?? new Set<string>();
  asked.set(pageId, already);
  const wanted = parts.filter((part) => !already.has(part));
  for (const part of wanted) already.add(part);
  for (let start = 0; start < wanted.length; start += LAYOUT_CARDS_PER_REQUEST) {
    const batch = wanted.slice(start, start + LAYOUT_CARDS_PER_REQUEST);
    const query = batch.map((part) => `layout=${encodeURIComponent(part)}`).join("&");
    let response: unknown;
    try {
      response = await apiGet<unknown>(`/api/v1/pages/${encodeURIComponent(pageId)}/pptx-layout-card?${query}`);
    } catch {
      // Asked again on the next open of the page.
      for (const part of batch) already.delete(part);
      continue;
    }
    const body = record(response);
    const current = cardsByPage.get(pageId) ?? EMPTY;
    const layouts = new Map(current.layouts);
    const masters = new Map(current.masters);
    const incomplete: string[] = [];
    for (const [part, value] of Object.entries(record(body?.layouts) ?? {})) {
      const card = readCard(value);
      if (!card) continue;
      layouts.set(part, card);
      if (typeof card.catalogCard === "object" && card.catalogCard.picturesPending) incomplete.push(part);
    }
    for (const [part, value] of Object.entries(record(body?.masters) ?? {})) {
      const card = readCard(value);
      if (card) masters.set(part, card);
    }
    cardsByPage.set(pageId, { layouts, masters });
    notify();
    // Shown without its pictures; asked for again once they are rendered.
    const again = incomplete.filter((part) => {
      const key = `${pageId}\n${part}`;
      const tries = picturesRetries.get(key) ?? 0;
      if (tries >= LAYOUT_CARD_PICTURES_RETRIES) return false;
      picturesRetries.set(key, tries + 1);
      return true;
    });
    if (again.length > 0) {
      setTimeout(() => {
        for (const part of again) already.delete(part);
        void requestLayoutCards(pageId, again);
      }, LAYOUT_CARD_PICTURES_RETRY_MS);
    }
  }
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** The cards held for `pageId` now (outside React: a command). */
export function layoutCardsOf(pageId: string): PageCards {
  return cardsByPage.get(pageId) ?? EMPTY;
}

/** The cards held for `pageId`. */
export function useLayoutCards(pageId: string): PageCards {
  return useSyncExternalStore(
    subscribe,
    () => cardsByPage.get(pageId) ?? EMPTY,
    () => EMPTY,
  );
}

function withCard<T extends SlideMaster | SlideMasterLayout>(target: T, card: DrawnCard | undefined): T {
  if (!card || target.catalogCard !== "pending") return target;
  const { catalogCard, ...drawn } = card;
  return { ...target, ...drawn, catalogCard } as T;
}

/** `masters` with the cards held for them drawn in, for the master view only. */
export function withLayoutCards(masters: readonly SlideMaster[], cards: PageCards): SlideMaster[] {
  if (cards.layouts.size === 0 && cards.masters.size === 0) return masters as SlideMaster[];
  return masters.map((master) => {
    const layouts = master.layouts?.map((layout) => withCard(layout, cards.layouts.get(layout.sourcePart ?? layout.id)));
    const changed = layouts?.some((layout, index) => layout !== master.layouts![index]);
    const drawn = withCard(master, cards.masters.get(master.sourcePart ?? master.id));
    return changed ? { ...drawn, layouts } : drawn;
  });
}

/**
 * The master view's masters: the page's, with the cards drawn so far. While
 * `enabled` (the master view is shown) the missing cards are asked for.
 */
export function useMastersWithLayoutCards(
  pageId: string | null | undefined,
  masters: readonly SlideMaster[],
  enabled: boolean,
): SlideMaster[] {
  const cards = useLayoutCards(pageId ?? "");
  const pending = useMemo(
    () => (pageId && enabled ? pendingLayoutCardParts(pageId, masters).join("\n") : ""),
    [enabled, masters, pageId],
  );
  useEffect(() => {
    if (!pageId || !pending) return;
    const parts = pending.split("\n");
    void requestLayoutCards(pageId, parts);
  }, [pageId, pending]);
  return useMemo(() => withLayoutCards(masters, cards), [cards, masters]);
}

/** Tests: forget every card and request. */
export function resetLayoutCardsForTests(): void {
  cardsByPage.clear();
  asked.clear();
  picturesRetries.clear();
}
