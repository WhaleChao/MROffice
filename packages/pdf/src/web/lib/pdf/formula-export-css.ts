/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const CSS_URL_RE = /url\(\s*(?:"([^"]+)"|'([^']+)'|([^'")\s]+))\s*\)/g;

function isSelfContainedCssUrl(url: string): boolean {
  return /^(?:data|blob|about):/i.test(url) || url.startsWith("#");
}

export function absolutizeCssUrls(
  cssText: string,
  baseHref: string | null | undefined,
): string {
  if (!baseHref) return cssText;
  return cssText.replace(CSS_URL_RE, (match, double, single, bare) => {
    const raw = String(double ?? single ?? bare ?? "").trim();
    if (!raw || isSelfContainedCssUrl(raw)) return match;
    try {
      return `url("${new URL(raw, baseHref).toString()}")`;
    } catch {
      return match;
    }
  });
}
