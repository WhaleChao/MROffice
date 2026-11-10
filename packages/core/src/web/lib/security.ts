/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const DEFAULT_ALLOWED_PROTOCOLS = new Set(["http:", "https:", "blob:"]);
const IMAGE_ALLOWED_PROTOCOLS = new Set(["http:", "https:", "blob:", "data:"]);
const CLICKABLE_LINK_ALLOWED_PROTOCOLS = new Set([
  "http:",
  "https:",
  "mailto:",
  "tel:",
]);
const STRIPE_REDIRECT_ALLOWED_HOSTS = new Set([
  "checkout.stripe.com",
  "billing.stripe.com",
]);

function getWindowOrigin(): string {
  if (typeof window === "undefined") {
    return "http://localhost";
  }
  return window.location.origin;
}

export function shouldUseSecureClientCookie(): boolean {
  return typeof window !== "undefined" && window.location.protocol === "https:";
}

export function buildClientCookieAttributes(maxAge?: number): string {
  let attributes = "Path=/; SameSite=Lax";
  if (maxAge !== undefined) {
    attributes += `; Max-Age=${maxAge}`;
  }
  if (shouldUseSecureClientCookie()) {
    attributes += "; Secure";
  }
  return attributes;
}

export function getSafeNavigationPath(value?: string | null): string | null {
  if (!value || typeof window === "undefined") {
    return null;
  }

  try {
    const url = new URL(value, window.location.origin);
    if (url.origin !== window.location.origin) {
      return null;
    }
    return `${url.pathname}${url.search}${url.hash}`;
  } catch {
    return null;
  }
}

export function getSafeStripeRedirectUrl(value?: string | null): string | null {
  const trimmed = value?.trim();
  if (!trimmed || typeof window === "undefined") {
    return null;
  }

  try {
    const url = new URL(trimmed, window.location.origin);
    if (url.origin === window.location.origin) {
      return `${url.pathname}${url.search}${url.hash}`;
    }
    if (
      url.protocol === "https:" &&
      STRIPE_REDIRECT_ALLOWED_HOSTS.has(url.hostname)
    ) {
      return url.toString();
    }
    return null;
  } catch {
    return null;
  }
}

export function getTrustedPostMessageOrigin(targetUrl?: string | null): string {
  const fallbackOrigin = getWindowOrigin();
  if (!targetUrl) {
    return fallbackOrigin;
  }

  try {
    return new URL(targetUrl, fallbackOrigin).origin;
  } catch {
    return fallbackOrigin;
  }
}

export function getSafeExternalUrl(
  value?: string | null,
  options: { allowData?: boolean } = {},
): string | null {
  if (!value) {
    return null;
  }

  const allowedProtocols = options.allowData
    ? IMAGE_ALLOWED_PROTOCOLS
    : DEFAULT_ALLOWED_PROTOCOLS;

  try {
    const url = new URL(value, getWindowOrigin());
    if (!allowedProtocols.has(url.protocol)) {
      return null;
    }
    return url.toString();
  } catch {
    return null;
  }
}

function stripUrlInsignificantChars(value: string): string {
  let out = "";
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code <= 0x20 || code === 0x7f) {
      continue;
    }
    out += value[index];
  }
  return out;
}

function hasUrlScheme(value: string): boolean {
  return /^[a-z][a-z0-9+\-.]*:/i.test(value);
}

export function getSafeClickableHref(value?: string | null): string | null {
  if (!value) {
    return null;
  }

  const stripped = stripUrlInsignificantChars(value);
  if (!stripped) {
    return null;
  }

  if (!hasUrlScheme(stripped) && !stripped.startsWith("//")) {
    return stripped;
  }

  try {
    const url = new URL(stripped, getWindowOrigin());
    if (!CLICKABLE_LINK_ALLOWED_PROTOCOLS.has(url.protocol)) {
      return null;
    }
    return url.toString();
  } catch {
    return null;
  }
}
