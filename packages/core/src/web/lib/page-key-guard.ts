/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Page-keyed caches and stores take a real page id or nothing (open1s-10 A1b):
 * a document window now exists before its page does (an import's window
 * mounts its editor with the page still pending), and a write keyed by an
 * empty or missing id would be kept under "" or "undefined" -- found by
 * nobody, or by the next pending window. In development and tests such a
 * write throws at once; in production it is not checked (the callers skip a
 * pending page themselves).
 */
const CHECKED = process.env.NODE_ENV !== "production";

function bogus(pageId: unknown): boolean {
  return typeof pageId !== "string" || pageId.trim() === "" || pageId === "undefined" || pageId === "null";
}

/** A write to `store` keyed by `pageId`: the id must be a page's. */
export function assertPageKey(pageId: unknown, store: string): void {
  if (CHECKED && bogus(pageId)) {
    throw new Error(`[page-key-guard] ${store}: a page-keyed write without a page id (${JSON.stringify(pageId)})`);
  }
}

/** A store that skips a missing id itself (no write): only a stringified
 * missing id ("undefined", "null") is a mistake there. */
export function assertNoStringifiedPageKey(pageId: unknown, store: string): void {
  if (CHECKED && (pageId === "undefined" || pageId === "null")) {
    throw new Error(`[page-key-guard] ${store}: a page id that is a stringified missing value (${pageId})`);
  }
}
