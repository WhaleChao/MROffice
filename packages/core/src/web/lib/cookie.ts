/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { buildClientCookieAttributes } from "./security";

export const CLIENT_COOKIE_CHANGED_EVENT = "mrnote:client-cookie-changed";

function emitCookieChange(): void {
  if (typeof window !== "undefined") {
    window.dispatchEvent(new Event(CLIENT_COOKIE_CHANGED_EVENT));
  }
}

export function subscribeCookieChanges(onChange: () => void): () => void {
  if (typeof window === "undefined") return () => undefined;
  window.addEventListener(CLIENT_COOKIE_CHANGED_EVENT, onChange);
  return () =>
    window.removeEventListener(CLIENT_COOKIE_CHANGED_EVENT, onChange);
}

/**
 * The embedded editor (lib/embed/frame-auth): the cookies the site's code
 * reads that a partner's page cannot hold -- the workspace the frame works
 * in, the session's page-cache principal -- given by the frame's session and
 * kept in this frame's memory only. Empty everywhere else.
 */
const COOKIE_OVERRIDES_GLOBAL = "__mroEmbedCookieOverrides";
// On globalThis: every copy of this module (one per chunk that bundles it)
// reads the same overrides.
const cookieOverrides: Map<string, string> = ((globalThis as Record<string, unknown>)[COOKIE_OVERRIDES_GLOBAL] as
  | Map<string, string>
  | undefined) ?? new Map<string, string>();
(globalThis as Record<string, unknown>)[COOKIE_OVERRIDES_GLOBAL] = cookieOverrides;

export function setCookieOverride(name: string, value: string): void {
  cookieOverrides.set(name, value);
  emitCookieChange();
}

export function readCookie(name: string): string | null {
  const override = cookieOverrides.get(name);
  if (override !== undefined) return override;
  if (typeof document === "undefined") return null;
  const match = document.cookie
    .split(";")
    .map((c) => c.trim())
    .find((c) => c.startsWith(`${name}=`));
  return match ? decodeURIComponent(match.split("=").slice(1).join("=")) : null;
}

export function writeCookie(
  name: string,
  value: string,
  maxAge?: number,
): void {
  if (typeof document === "undefined") return;
  document.cookie = `${name}=${encodeURIComponent(value)}; ${buildClientCookieAttributes(maxAge)}`;
  emitCookieChange();
}

export function clearCookie(name: string): void {
  if (typeof document === "undefined") return;
  document.cookie = `${name}=; ${buildClientCookieAttributes(0)}`;
  emitCookieChange();
}

/**
 * Clears `name` where the server set it with a Domain attribute (this host
 * and each parent domain). To a browser that is a different cookie from the
 * host-only one `writeCookie` sets: neither replaces the other, and the older
 * is read first.
 */
export function clearDomainScopedCookie(name: string): void {
  if (typeof document === "undefined" || typeof window === "undefined") return;
  const host = window.location.hostname;
  if (!host || !host.includes(".") || /^[\d.]+$/.test(host) || host.includes(":")) return;
  const labels = host.split(".");
  for (let index = 0; index < labels.length - 1; index += 1) {
    const domain = labels.slice(index).join(".");
    document.cookie = `${name}=; Domain=${domain}; ${buildClientCookieAttributes(0)}`;
  }
  emitCookieChange();
}
