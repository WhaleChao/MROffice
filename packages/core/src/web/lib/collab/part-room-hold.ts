/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pages whose body a document part's live room holds in this tab
 * (lib/document-page-parts). A window remounts its content when someone else
 * saves its page, unless a live session delivers that content; a page room's
 * session says so through its transport (`getPageYjsTransportSnapshot`). A
 * part room is no page room, so it says so here. Without this every save of
 * the room's, by any member, threw the others' editors away: their caret, their
 * Undo and the parts they had loaded.
 */
const held = new Map<string, number>();

/** While the room may deliver the page's content. Returns the release. */
export function holdPageForPartRoom(pageId: string): () => void {
  held.set(pageId, (held.get(pageId) ?? 0) + 1);
  let released = false;
  return () => {
    if (released) return;
    released = true;
    const count = (held.get(pageId) ?? 1) - 1;
    if (count > 0) held.set(pageId, count);
    else held.delete(pageId);
  };
}

export function pageHeldByPartRoom(pageId: string | null | undefined): boolean {
  return Boolean(pageId && held.has(pageId));
}
