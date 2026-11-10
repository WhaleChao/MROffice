/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The address rules a document page's save applies to a link's `href` and a
 * picture's `src` (apps/api notebooks.py `_validate_document_url`). Whatever
 * the editor writes into a link must pass them: an address the save refuses
 * makes the WHOLE document unsavable (WUA 024: a link edited to `https://`).
 */

/** The server's `MAX_DOCUMENT_URL_LENGTH`. */
export const MAX_DOCUMENT_URL_LENGTH = 8192;

// Python `urlsplit` raises on control characters it would otherwise strip,
// and the save refuses any whitespace or C0/C1 control inside an address.
const URL_WHITESPACE_OR_CONTROL = /[\u0000- \u007f-\u009f]/;
const URL_SCHEME = /^([A-Za-z][A-Za-z0-9+\-.]*):/;

/**
 * Why a document page's save would refuse `href` as a stated address, or
 * null when it accepts it. Only the shape is checked here; which schemes a
 * caller allows is its own policy.
 */
export function documentUrlProblem(
  href: string,
): "empty" | "too_long" | "whitespace" | "no_host" | "invalid" | null {
  if (href.length > MAX_DOCUMENT_URL_LENGTH) return "too_long";
  if (URL_WHITESPACE_OR_CONTROL.test(href)) return "whitespace";
  if (!href) return "empty";
  const scheme = URL_SCHEME.exec(href)?.[1]?.toLowerCase();
  if (!scheme) return null;
  // `urlsplit` reads a host only after `scheme://`: `https:example.com` and
  // `https://` both have none, and a web address without a host is refused.
  const rest = href.slice(scheme.length + 1);
  const netloc = rest.startsWith("//") ? /^[^/?#]*/.exec(rest.slice(2))?.[0] ?? "" : "";
  const opens = netloc.includes("[");
  const closes = netloc.includes("]");
  if (opens !== closes) return "invalid";
  if (opens) {
    // A bracketed host must be an IPv6 (or IPvFuture) literal.
    const literal = /\[([^\]]*)\]/.exec(netloc)?.[1] ?? "";
    if (!/^(?:[0-9a-f:.]+|v[0-9a-f]+\..+)(?:%.+)?$/i.test(literal)) return "invalid";
  }
  if ((scheme === "http" || scheme === "https") && !netloc) return "no_host";
  return null;
}

/** True when a document page's save accepts `href` as a stated address. */
export function isSavableDocumentUrl(href: unknown): href is string {
  return typeof href === "string" && documentUrlProblem(href) === null;
}
