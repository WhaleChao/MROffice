/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Alias app origins: further exact HTTPS origins the same site is served on,
 * e.g. https://cn.mr-note.com (the accelerated entry for mainland China,
 * reached through a relay that forwards to the production host), while the
 * canonical origin (NEXT_PUBLIC_API_BASE_URL) stays https://mr-note.com.
 *
 * NEXT_PUBLIC_APP_ALIAS_ORIGINS lists them, separated by commas or spaces; it
 * is compiled into the bundle like every NEXT_PUBLIC_ value. Unset or empty,
 * nothing here changes behaviour. A page on an alias talks to its own origin
 * (same-origin /api, sockets) instead of the canonical one; the file and
 * sandbox hosts stay canonical and accept the alias as an app origin.
 */

/** The configured alias origins: exact HTTPS origins only, others ignored. */
export function configuredAppAliasOrigins(
  raw: string | undefined = process.env.NEXT_PUBLIC_APP_ALIAS_ORIGINS,
): string[] {
  const aliases = new Set<string>();
  for (const value of (raw ?? "").split(/[\s,]+/)) {
    if (!value) continue;
    try {
      const url = new URL(value);
      if (
        url.protocol === "https:" &&
        url.origin === value &&
        !url.username &&
        !url.password
      ) {
        aliases.add(value);
      }
    } catch {
      continue;
    }
  }
  return Array.from(aliases);
}

/**
 * The page's own origin when it is one of the configured aliases, else null
 * (always null on the server and on the canonical origin).
 */
export function browserAppAliasOrigin(): string | null {
  if (typeof window === "undefined") return null;
  const here = window.location?.origin;
  if (!here) return null;
  return configuredAppAliasOrigins().includes(here) ? here : null;
}

/**
 * The file host of a sister site: an alias on a registrable domain of its
 * own (https://mroffice.ai beside the canonical https://mr-note.com). The
 * canonical file host is on another site there, so the browser would send it
 * none of the page's cookies; the sister pairs with `https://files.<its
 * host>`, the same file edge under its own name. Null for an alias under the
 * canonical host (cn.mr-note.com keeps the canonical file host), for no
 * alias, and when the canonical origin is not configured.
 */
export function sisterFileOrigin(
  alias: string | null | undefined,
  canonical: string | undefined = process.env.NEXT_PUBLIC_API_BASE_URL,
): string | null {
  if (!alias || !canonical) return null;
  try {
    const host = new URL(alias).hostname;
    const main = new URL(canonical).hostname;
    if (!host || !main || host === main || host.endsWith(`.${main}`)) return null;
    return `https://files.${host}`;
  } catch {
    return null;
  }
}
