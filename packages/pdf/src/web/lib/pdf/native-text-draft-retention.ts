/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Unsaved native-text drafts kept when the source proof failed for good and
 * the reader closed the editor anyway: reopening the same line on the same
 * source bytes brings the typing back. Memory only (never storage, never
 * uploaded), keyed by page + line + source SHA-256. A draft for other bytes
 * than the page now shows (saved elsewhere, undo/redo, a reload) is dropped,
 * never applied to them. Bounded: 32 drafts a page, 64 KiB a draft, oldest
 * out first.
 */

export const MAX_RETAINED_DRAFTS_PER_PAGE = 32;
/** UTF-16 code units: 32 Ki characters = 64 KiB. */
export const MAX_RETAINED_DRAFT_CHARS = 32 * 1024;

interface RetainedDraft {
  readonly sourceSha256: string;
  readonly text: string;
}

const pages = new Map<string, Map<string, RetainedDraft>>();

/** Where the drafts live across a reload of this tab: the update banner
 * promises that unsaved changes are kept, and a draft a line could not
 * take (a blank line typed into it) was kept in memory only, gone after
 * the refresh the banner asked for (user audit C-0070). Session storage is
 * the tab's own and ends with it; a browser without it keeps memory only. */
const STORAGE_KEY = "mrai.pdf.native-text-drafts.v1";
let restored = false;

function restoreFromSession(): void {
  if (restored) return;
  restored = true;
  try {
    const raw = typeof sessionStorage !== "undefined" ? sessionStorage.getItem(STORAGE_KEY) : null;
    if (!raw) return;
    const stored = JSON.parse(raw) as Record<string, Record<string, RetainedDraft>>;
    for (const [pageId, drafts] of Object.entries(stored)) {
      if (!pageId || !drafts || typeof drafts !== "object") continue;
      const map = new Map<string, RetainedDraft>();
      for (const [lineKey, draft] of Object.entries(drafts)) {
        if (!lineKey || !draft || typeof draft.text !== "string" || !/^[a-f0-9]{64}$/.test(String(draft.sourceSha256)) ||
          draft.text.length > MAX_RETAINED_DRAFT_CHARS) continue;
        map.set(lineKey, Object.freeze({ sourceSha256: draft.sourceSha256, text: draft.text }));
        if (map.size >= MAX_RETAINED_DRAFTS_PER_PAGE) break;
      }
      if (map.size) pages.set(pageId, map);
    }
  } catch {
    // Unreadable or blocked storage: memory only.
  }
}

function persistToSession(): void {
  try {
    if (typeof sessionStorage === "undefined") return;
    if (pages.size === 0) {
      sessionStorage.removeItem(STORAGE_KEY);
      return;
    }
    const stored: Record<string, Record<string, RetainedDraft>> = {};
    for (const [pageId, drafts] of pages) stored[pageId] = Object.fromEntries(drafts);
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(stored));
  } catch {
    // Quota or blocked storage: memory still holds the draft.
  }
}

export function retainPdfNativeTextDraft(input: Readonly<{
  pageId: string;
  lineKey: string;
  sourceSha256: string;
  text: string;
}>): boolean {
  if (!input.pageId || !input.lineKey || !/^[a-f0-9]{64}$/.test(input.sourceSha256) ||
    input.text.length > MAX_RETAINED_DRAFT_CHARS) return false;
  restoreFromSession();
  let drafts = pages.get(input.pageId);
  if (!drafts) {
    drafts = new Map();
    pages.set(input.pageId, drafts);
  }
  drafts.delete(input.lineKey);
  drafts.set(input.lineKey, Object.freeze({ sourceSha256: input.sourceSha256, text: input.text }));
  while (drafts.size > MAX_RETAINED_DRAFTS_PER_PAGE) drafts.delete(drafts.keys().next().value as string);
  persistToSession();
  return true;
}

/** The draft kept for this line on these bytes, or null. A draft kept on
 * other bytes is dropped here (the page has moved on). */
export function readRetainedPdfNativeTextDraft(input: Readonly<{
  pageId: string;
  lineKey: string;
  sourceSha256: string;
}>): string | null {
  restoreFromSession();
  const drafts = pages.get(input.pageId);
  const kept = drafts?.get(input.lineKey);
  if (!drafts || !kept) return null;
  if (kept.sourceSha256 !== input.sourceSha256) {
    drafts.delete(input.lineKey);
    persistToSession();
    return null;
  }
  return kept.text;
}

/** The line's draft was saved (or discarded on purpose): forget it. */
export function forgetPdfNativeTextDraft(pageId: string, lineKey: string): void {
  restoreFromSession();
  const drafts = pages.get(pageId);
  drafts?.delete(lineKey);
  if (drafts && drafts.size === 0) pages.delete(pageId);
  persistToSession();
}

/** Every page's drafts on bytes other than `sourceSha256` go (the page's
 * source changed under them). */
export function dropStalePdfNativeTextDrafts(pageId: string, sourceSha256: string): void {
  const drafts = pages.get(pageId);
  if (!drafts) return;
  for (const [key, draft] of drafts) if (draft.sourceSha256 !== sourceSha256) drafts.delete(key);
  if (drafts.size === 0) pages.delete(pageId);
  persistToSession();
}

/** Tests. */
export function clearRetainedPdfNativeTextDrafts(): void {
  pages.clear();
  restored = true;
  persistToSession();
}
