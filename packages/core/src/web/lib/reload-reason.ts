/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const KEY = "mrnote:reload-reason";
/** A reason older than this is not the one this load came from. */
const FRESH_MS = 2 * 60_000;
const REASON = /^[a-z0-9_:-]{1,40}$/;

/** The reason the app is about to load the page again (or go to a page). */
export function noteAppReload(reason: string): void {
  const value = REASON.test(reason) ? reason : "app";
  try {
    window.sessionStorage.setItem(KEY, JSON.stringify({ reason: value, at: Date.now() }));
  } catch {
    // Storage blocked: the load is told as the browser saw it.
  }
}

/** Note the reason, then reload. For the app's own "Reload" buttons. */
export function reloadPage(reason: string): void {
  noteAppReload(reason);
  window.location.reload();
}

function navigationType(): string | undefined {
  try {
    const entry = performance.getEntriesByType?.("navigation")?.[0] as
      | PerformanceNavigationTiming
      | undefined;
    return entry?.type || undefined;
  } catch {
    return undefined;
  }
}

/**
 * How this page load came about: `nav` as the browser tells it, and
 * `reload_reason` when the app made it (read once: a later load in the same
 * tab does not inherit it).
 */
export function takeLoadFacts(now: number = Date.now()): { nav?: string; reload_reason?: string } {
  const facts: { nav?: string; reload_reason?: string } = {};
  const nav = navigationType();
  if (nav) facts.nav = nav;
  try {
    const raw = window.sessionStorage.getItem(KEY);
    if (raw) {
      window.sessionStorage.removeItem(KEY);
      const stored = JSON.parse(raw) as { reason?: unknown; at?: unknown };
      if (
        typeof stored.reason === "string" &&
        REASON.test(stored.reason) &&
        typeof stored.at === "number" &&
        now - stored.at >= 0 &&
        now - stored.at < FRESH_MS
      ) {
        facts.reload_reason = stored.reason;
      }
    }
  } catch {
    // Nothing kept: the browser's account alone.
  }
  return facts;
}
