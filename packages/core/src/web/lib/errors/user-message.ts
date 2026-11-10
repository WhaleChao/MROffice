/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * `toUserMessage` — map any caught error to a FRIENDLY, LOCALIZED string that
 * is safe to show the user in a toast / alert / notification.
 *
 * The production audit found ~13 sites putting a raw `error.message` (or a raw
 * backend error string) straight into a user-visible toast description. That
 * leaks backend internals (stack-ish detail, internal identifiers, framework
 * wording) to the end user and reads as broken for a polished launch. This
 * helper is the single funnel those sites now go through:
 *
 *   - For a backend {@link ApiRequestError}, it maps the STABLE machine-readable
 *     `code` to a curated friendly line from the `error.code.*` catalog
 *     (`messages/{en,zh}/error.json`) — e.g. `forbidden` →
 *     "You don't have permission to perform this action." The catalog is read
 *     directly (same self-contained, locale-aware pattern as `lib/api.ts`) so
 *     this helper works from any call site without the caller needing the
 *     `error` translation namespace.
 *   - For anything unmapped / non-Error, it returns the caller's localized
 *     `fallbackKey` (or a generic "Something went wrong. Please try again.").
 *
 * It NEVER returns the raw `error.message`, `error.stack`, or any backend free
 * text. The real detail is still observable — callers log it to the console
 * and/or report it via the error telemetry beacon; only the friendly string
 * reaches the UI.
 */
import { isApiRequestError } from "../api";
import {
  currentErrorCatalog,
  registerErrorCatalog,
} from "./error-catalog";

export { registerErrorCatalog };

/** Same shape next-intl's `useTranslations(ns)` returns; also satisfied by a
 *  plain `(key) => string` in tests. */
export type TranslationFn = (
  key: string,
  values?: Record<string, string | number>,
) => string;

/** Generic last-resort key in the shared `error` namespace ("Something went
 *  wrong. Please try again." / "出了点问题，请重试。"). */
const GENERIC_FALLBACK_KEY = "genericRetry";

/** The flat `error.json` catalog for the current locale (dotted keys). */
function errorCatalog(): Record<string, string> {
  return currentErrorCatalog();
}

/**
 * Locate a structured API error through a small, cycle-safe wrapper chain.
 * Import pipelines attach context with ``Error.cause``; treating only the
 * outer error as authoritative used to discard the stable backend code and
 * replace a truthful, actionable state with a generic retry message.
 */
function apiRequestErrorInCauseChain(error: unknown): unknown {
  const seen = new Set<object>();
  let current = error;
  for (let depth = 0; depth < 5; depth += 1) {
    if (isApiRequestError(current)) return current;
    if (
      (typeof current !== "object" && typeof current !== "function") ||
      !current
    ) {
      return null;
    }
    if (seen.has(current)) return null;
    seen.add(current);
    try {
      current = (current as { cause?: unknown }).cause;
    } catch {
      return null;
    }
  }
  return null;
}

/**
 * Friendly text for a backend error `code`, read straight from the bundled
 * `error.code.*` catalog. Returns `null` when the code has no curated entry, so
 * the caller can fall back rather than surface anything raw.
 */
export function friendlyMessageForCode(code: string | undefined): string | null {
  if (!code) return null;
  const value = errorCatalog()[`code.${code}`];
  return typeof value === "string" && value.length > 0 ? value : null;
}

/**
 * The generic, always-available fallback line in the user's locale. Read from
 * the `error` catalog so it stays translated even though this module can't use
 * a next-intl hook.
 */
export function genericErrorMessage(): string {
  const value = errorCatalog()[GENERIC_FALLBACK_KEY];
  return typeof value === "string" && value.length > 0
    ? value
    : "Something went wrong. Please try again.";
}

/**
 * Map a caught error to a friendly, localized, leak-free user message.
 *
 * @param error       The caught value (any shape; need not be an `Error`).
 * @param t           The caller's scoped translator, used ONLY to resolve
 *                    `fallbackKey`. Optional — omit for non-React call sites
 *                    (e.g. `window.alert`), which then get the generic line.
 * @param fallbackKey A key in the caller's namespace for a context-specific
 *                    friendly fallback (e.g. `"pages.createFailed"`). Used when
 *                    the error has no mapped backend `code`.
 *
 * Resolution order:
 *   1. `ApiRequestError` (direct or through a bounded `cause` chain) with a
 *      curated `code.*` entry → that friendly line.
 *   2. Locally branded Spreadsheet import losses → the existing readable
 *      workbook labels, with a curated rejection heading.
 *   3. `fallbackKey` via `t` (the caller's context-specific friendly copy).
 *   4. The generic localized "Something went wrong" line.
 *
 * The raw `error.message` / `stack` is never returned.
 */
export function toUserMessage(
  error: unknown,
  t?: TranslationFn,
  fallbackKey?: string,
): string {
  const apiError = apiRequestErrorInCauseChain(error);
  if (isApiRequestError(apiError)) {
    const friendly = friendlyMessageForCode(apiError.code);
    if (friendly) {
      return friendly;
    }
  }

  if (t && fallbackKey) {
    try {
      const translated = t(fallbackKey);
      // next-intl returns the key itself (or a marker) on a miss; only use a
      // resolved, non-key string. Anything else falls through to the generic.
      if (translated && translated !== fallbackKey) {
        return translated;
      }
    } catch {
      /* fall through to the generic line */
    }
  }

  return genericErrorMessage();
}
