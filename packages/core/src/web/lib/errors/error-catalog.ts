/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import ERROR_EN from "../../messages/en/error.json";
import ERROR_ZH from "../../messages/zh/error.json";
import { isLocale } from "../../i18n/locales";

/**
 * The `error` catalogue for the page the visitor is on, for code that cannot
 * use next-intl hooks (the API client, the error-message helper).
 *
 * <ErrorCatalogBridge> hands over the page's own catalogue — the root
 * provider already ships the `error` namespace to every page, in the page's
 * locale — so these lines read in all 20 locales. Registration is
 * client-only: module state on the server would be shared between requests
 * of different locales. Until it happens (and on the server) the bundled
 * catalogues stand in: Chinese for the default and zh paths, English for
 * every other locale.
 */
let pageCatalog: Record<string, string> | null = null;

function flatten(
  value: unknown,
  prefix = "",
  out: Record<string, string> = {},
): Record<string, string> {
  if (typeof value === "string") {
    out[prefix] = value;
  } else if (value && typeof value === "object") {
    for (const [key, child] of Object.entries(value)) {
      flatten(child, prefix ? `${prefix}.${key}` : key, out);
    }
  }
  return out;
}

export function registerErrorCatalog(messages: unknown): void {
  if (typeof window === "undefined" || !messages) return;
  pageCatalog = flatten(messages);
}

/** True on the default (unprefixed) locale and on zh / zh-TW paths. */
export function isChinesePath(): boolean {
  if (typeof window === "undefined") return true;
  const first = window.location.pathname.split("/")[1];
  return !isLocale(first) || first === "zh" || first === "zh-TW";
}

/** The flat catalogue (dotted keys) to read error lines from. */
export function currentErrorCatalog(): Record<string, string> {
  if (pageCatalog) return pageCatalog;
  return (isChinesePath() ? ERROR_ZH : ERROR_EN) as Record<string, string>;
}
