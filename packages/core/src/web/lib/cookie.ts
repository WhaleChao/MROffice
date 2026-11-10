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

export function readCookie(name: string): string | null {
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
