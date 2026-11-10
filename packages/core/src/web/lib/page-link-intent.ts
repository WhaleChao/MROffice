/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A page opened by a link (#page=…) rather than by its window coming back.
 *
 * A window restored from minimized returns where it was; a window the
 * notebook brings up because a link named its page goes where the link
 * says (the link's sheet, cell or range). The canvas notes the link it is
 * opening; the page's editor takes the note once, when it loads.
 */
const pending = new Map<string, { hash: string; at: number }>();
const PAGE_LINK_INTENT_MS = 60_000;

export function notePageLinkOpen(pageId: string, hash: string): void {
  pending.set(pageId, { hash, at: Date.now() });
}

/** Whether the page was opened by this link (consumed by the first reader). */
export function takePageLinkOpen(pageId: string, hash: string): boolean {
  const intent = pending.get(pageId);
  if (!intent) return false;
  pending.delete(pageId);
  return intent.hash === hash && Date.now() - intent.at < PAGE_LINK_INTENT_MS;
}

/** The page a notebook URL's hash links to, or null. */
export function linkedPageIdFromHash(hash: string): string | null {
  const params = new URLSearchParams(hash.startsWith("#") ? hash.slice(1) : hash);
  const pageId = params.get("page");
  return pageId && pageId.trim() ? pageId.trim() : null;
}
