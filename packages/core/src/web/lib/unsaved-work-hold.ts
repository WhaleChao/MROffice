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

/** True from a held 401 until the session is proven back. */
export function isAuthSessionHeld(): boolean {
  return authSessionHeld;
}

export function setAuthSessionHeld(held: boolean): void {
  authSessionHeld = held;
}
