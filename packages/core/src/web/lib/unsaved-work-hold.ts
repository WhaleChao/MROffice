/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An expired session while a window holds work the server does not have yet
 * (Office's "Sign in to save this file"): the tab is not sent to the login
 * page, which would close the work; the windows say so and offer to sign in
 * again beside it (lib/session-reauth). Editors that can keep their edits
 * register here; without a holder a 401 redirects as it always did.
 *
 * Its own module so an editor and a room can register and listen without
 * loading the API client (lib/api decides the 401 and re-exports these).
 */
export const AUTH_SESSION_HELD_EVENT = "mrnote:auth-session-held";
export const AUTH_SESSION_RESTORED_EVENT = "mrnote:auth-session-restored";
const unsavedWorkHolds = new Map<string, () => boolean>();
let authSessionHeld = false;

export function registerUnsavedWorkHold(id: string, holdsWork: () => boolean): () => void {
  unsavedWorkHolds.set(id, holdsWork);
  return () => {
    if (unsavedWorkHolds.get(id) === holdsWork) unsavedWorkHolds.delete(id);
  };
}

export function unsavedWorkIsHeld(): boolean {
  for (const holdsWork of unsavedWorkHolds.values()) {
    try {
      if (holdsWork()) return true;
    } catch {
      // a holder that cannot answer holds nothing
    }
  }
  return false;
}

/** Whether the editors of one page (each registers as `<kind>:<pageId>`)
 * hold work its server copy does not have yet: a ⌘S that writes the file
 * from the server waits for none. */
export function pageHoldsUnsavedWork(pageId: string): boolean {
  for (const [id, holdsWork] of unsavedWorkHolds) {
    if (!id.endsWith(`:${pageId}`)) continue;
    try {
      if (holdsWork()) return true;
    } catch {
      // a holder that cannot answer holds nothing
    }
  }
  return false;
}

/** True from a held 401 until the session is proven back. */
export function isAuthSessionHeld(): boolean {
  return authSessionHeld;
}

export function setAuthSessionHeld(held: boolean): void {
  authSessionHeld = held;
}

/**
 * What an editor holds that is not even in its document yet (a speaker-notes
 * draft written on blur, a typing burst drawn and saved after a pause): each
 * registers how to write it in now, and how to save now. A window about to
 * close or the app about to quit settles its work (`settleUnsavedWork`):
 * every flush and every editor's reload handover first, then the holds are
 * waited on until none holds work or the time is up, then the local recovery
 * copies are written. Not settled, the page still asks before it goes.
 */
const unsavedWorkFlushes = new Map<string, () => unknown>();

export function registerUnsavedWorkFlush(id: string, flush: () => unknown): () => void {
  unsavedWorkFlushes.set(id, flush);
  return () => {
    if (unsavedWorkFlushes.get(id) === flush) unsavedWorkFlushes.delete(id);
  };
}

/** Every registered flush, now (a page hiding: no waiting). */
export function flushUnsavedWork(): void {
  for (const flush of unsavedWorkFlushes.values()) {
    try {
      const done = flush();
      if (done && typeof (done as Promise<unknown>).catch === "function") void (done as Promise<unknown>).catch(() => null);
    } catch {
      // one editor's flush failing leaves the others to theirs
    }
  }
}

export interface UnsavedWorkSettled {
  /** Nothing held unsaved when it returned. */
  readonly settled: boolean;
  readonly waitedMs: number;
}

/** The work written in and saved, as far as `timeoutMs` allows; the local
 * recovery copies written either way (what did not reach the server is on
 * this device for the next open). */
export async function settleUnsavedWork({ timeoutMs = 3000, pollMs = 50 } = {}): Promise<UnsavedWorkSettled> {
  const started = Date.now();
  flushUnsavedWork();
  // The editors that hand their drafts over before a reload (a PDF line
  // being typed: lib/editor-draft-handover) do so here too, within the time.
  try {
    const { handOverEditorDrafts } = await import("./editor-draft-handover");
    await handOverEditorDrafts(Math.max(0, timeoutMs - (Date.now() - started)));
  } catch {
    // an editor's handover failing leaves the rest to the holds below
  }
  // A flush queues its save; the holds read it from state a render later.
  await new Promise((resolve) => setTimeout(resolve, pollMs));
  while (unsavedWorkIsHeld() && Date.now() - started < timeoutMs) {
    await new Promise((resolve) => setTimeout(resolve, pollMs));
  }
  const settled = !unsavedWorkIsHeld();
  try {
    const { flushAutoRecoverWrites } = await import("./autorecover-store");
    await flushAutoRecoverWrites();
  } catch {
    // no recovery store here: what was saved is saved
  }
  return { settled, waitedMs: Date.now() - started };
}
