/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Signing in again beside held work (lib/api `AUTH_SESSION_HELD_EVENT`):
 * the sign-in page opens in a popup (or a new tab), and the held tab learns
 * that the session is back from the popup's announcement or, failing that,
 * by asking the server now and then. The tab never leaves its work.
 */
import { confirmAuthSessionRestored } from "./api";
import { isAuthSessionHeld } from "./unsaved-work-hold";
import { visitorLocalePrefix } from "../i18n/locales";

const CHANNEL = "mrai-auth-reauth";
const STORAGE_KEY = "mrai:auth-reauth-signed-in";
const POLL_MS = 5_000;

/** The sign-in page, returning to the small "signed in" page. */
export function reauthLoginUrl(): string {
  const prefix = visitorLocalePrefix();
  const params = new URLSearchParams({ next: `${prefix}/signed-in` });
  return `${prefix}/login?${params.toString()}`;
}

/** Open the sign-in page beside the work; false when the browser refused. */
export function openReauthWindow(): boolean {
  if (typeof window === "undefined") return false;
  const popup = window.open(reauthLoginUrl(), "mrai-reauth", "popup,width=520,height=760");
  if (popup) {
    try { popup.focus(); } catch { /* another origin's focus rules */ }
    return true;
  }
  return false;
}

/** Called by the "signed in" page: every held tab of this browser resumes. */
export function announceSignedIn(): void {
  if (typeof window === "undefined") return;
  try {
    const channel = new BroadcastChannel(CHANNEL);
    channel.postMessage({ type: "signed-in", at: Date.now() });
    channel.close();
  } catch {
    // no BroadcastChannel: the storage event below
  }
  try {
    window.localStorage.setItem(STORAGE_KEY, String(Date.now()));
  } catch {
    // the held tab's polling still finds the session
  }
}

/**
 * While the session is held, find out when it is back and call
 * `onRestored` once (the caller resumes saving). Returns the unsubscribe.
 */
export function watchSessionRestore(onRestored: () => void): () => void {
  if (typeof window === "undefined") return () => {};
  let stopped = false;
  let checking = false;
  const check = async () => {
    if (stopped || checking || !isAuthSessionHeld()) return;
    checking = true;
    try {
      if (await confirmAuthSessionRestored()) {
        if (!stopped) {
          stopped = true;
          onRestored();
        }
      }
    } finally {
      checking = false;
    }
  };
  let channel: BroadcastChannel | null = null;
  try {
    channel = new BroadcastChannel(CHANNEL);
    channel.onmessage = () => { void check(); };
  } catch {
    channel = null;
  }
  const onStorage = (event: StorageEvent) => {
    if (event.key === STORAGE_KEY) void check();
  };
  const onVisible = () => {
    if (document.visibilityState === "visible") void check();
  };
  const timer = window.setInterval(() => {
    if (document.visibilityState === "visible") void check();
  }, POLL_MS);
  window.addEventListener("storage", onStorage);
  window.addEventListener("focus", onVisible);
  document.addEventListener("visibilitychange", onVisible);
  return () => {
    stopped = true;
    window.clearInterval(timer);
    window.removeEventListener("storage", onStorage);
    window.removeEventListener("focus", onVisible);
    document.removeEventListener("visibilitychange", onVisible);
    channel?.close();
  };
}
