/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { DEFAULT_LOCALE, isLocale, stripLocalePrefix } from "../i18n/locales";

export const DISCOVER_ENABLED = false;

/**
 * P3.B6.GOV.PR0 — Admin governance write UI master switch.
 *
 * Enabled after the admin write surfaces landed behind their own confirmation
 * dialogs and the backend governance / budget runtime flags moved to their
 * signed-off default-on path. Individual mutations still enforce CSRF,
 * role checks, dual-sign / approver requirements, and backend flag gates.
 */
export const ENABLE_GOVERNANCE_ADMIN_WRITE_UI = true;

























export const SANDBOXED_EXTENSION_FRAME_ENABLED = true;




















export const AI_CUSTOM_TOOL_ENABLED = true;

/**
 * Spotlight command palette — workspace search + AI bridge in Cmd+K.
 *
 * Graduated ON after the command palette and notebook SDK moved to the live
 * grouped ``{ query, duration_ms, results }`` contract, including exact page
 * deep links and explicit loading / empty / error states. Emergency rollback:
 * set ``NEXT_PUBLIC_SPOTLIGHT_SEARCH_ENABLED=0``; navigation and actions remain
 * usable while workspace search + the AI bridge are hidden.
 */
export const SPOTLIGHT_SEARCH_ENABLED =
  process.env.NEXT_PUBLIC_SPOTLIGHT_SEARCH_ENABLED !== "0";

/**
 * Quick Capture — a frictionless Markdown capture surface (Cmd/Ctrl+Shift+K)
 * that saves to a page. Graduated ON after the optional notebook + page write
 * became one transaction with durable idempotent replay, and the client closed
 * loading/error, focus, stale-response, and exact-page navigation boundaries.
 * Set ``NEXT_PUBLIC_QUICK_CAPTURE_ENABLED=0`` for an exact rollback.
 */
export const QUICK_CAPTURE_ENABLED =
  process.env.NEXT_PUBLIC_QUICK_CAPTURE_ENABLED !== "0";








export const WORKSPACE_LAYOUTS_ENABLED = true;

/**
 * Keyboard-shortcuts help overlay (press "?").
 *
 * A reference card of the workspace's keybindings, opened with "?" and closed
 * with Escape / backdrop. Graduated ON after adopting the shared dialog focus
 * trap/return contract, platform-correct Cmd/Ctrl rendering, and a command list
 * derived from the active Workspace feature gates. Emergency rollback: set
 * ``NEXT_PUBLIC_SHORTCUTS_HELP_ENABLED=0``.
 */
export const SHORTCUTS_HELP_ENABLED =
  process.env.NEXT_PUBLIC_SHORTCUTS_HELP_ENABLED !== "0";

/**
 * Workspace-wide "export all notebooks" download. Graduated ON after the
 * archive boundary closed full-list validation, authoritative page hydration,
 * lossless-content rejection, empty workspaces, collision-safe names, and
 * visible UI failures. Pure client + existing ACL-filtered read SDK. Emergency
 * rollback: set ``NEXT_PUBLIC_WORKSPACE_EXPORT_ENABLED=0``.
 */
export const WORKSPACE_EXPORT_ENABLED =
  process.env.NEXT_PUBLIC_WORKSPACE_EXPORT_ENABLED !== "0";

/**
 * Workspace-wide "import notebooks from a zip" restore. The normalized
 * archive now commits every new notebook/page plus a semantic-hash receipt in
 * one server transaction, and both UI entries share one execution flight.
 * Rollback: set ``NEXT_PUBLIC_WORKSPACE_IMPORT_ENABLED=0``.
 */
export const WORKSPACE_IMPORT_ENABLED =
  process.env.NEXT_PUBLIC_WORKSPACE_IMPORT_ENABLED !== "0";

/**
 * Notebook starter templates in the empty-workspace onboarding. Defaults to
 * ``false`` (dark-launch). Template creation currently performs incremental
 * notebook/page writes without rollback. Keep this off until the transactional
 * creation flow is safe.
 */
export const NOTEBOOK_TEMPLATES_ENABLED = false;

export function resolveDiscoverRedirectTarget(
  locale: string,
  from?: string | null,
): string {
  // The default locale is served unprefixed; every other roster locale keeps
  
  // /de reader was bounced to the Chinese console).
  const prefix =
    isLocale(locale) && locale !== DEFAULT_LOCALE ? `/${locale}` : "";
  const fallback = `${prefix}/app`;

  if (
    !from ||
    !from.startsWith("/") ||
    from.startsWith("//") ||
    from.includes("\\")
  ) {
    return fallback;
  }
  try {
    const parsed = new URL(from, "https://mrnote.local");
    if (
      parsed.origin !== "https://mrnote.local" ||
      !stripLocalePrefix(parsed.pathname).startsWith("/app")
    ) {
      return fallback;
    }
    from = `${parsed.pathname}${parsed.search}${parsed.hash}`;
  } catch {
    return fallback;
  }

  // Already locale-prefixed by the caller: hand it back untouched.
  if (stripLocalePrefix(from) !== from) {
    return from;
  }

  return `${prefix}${from}`;
}






















export const SPREADSHEET_DEEP_FORMULA_CHAINS = true;



























export const SPREADSHEET_RANGE_VALUE_CACHE = true;





















export const SPREADSHEET_WARM_RECALC = true;























/** A long Word document goes the source-import road, whose page opens from
 * the document's first part (`docx_progressive_import` on the server). */
export function longWordSourceRoadEnabled(): boolean {
  
  // 本地版就怎么开); its source road reads the file where it is on disk
  // (`/desktop/source-uploads/{id}/from-local`).
  try {
    const browser = typeof window !== "undefined" ? window.localStorage?.getItem("mrai.longWordSourceRoad") : null;
    if (browser === "0") return false;
    if (browser === "1") return true;
  } catch {
    // storage unavailable: the build decides
  }
  return process.env.NEXT_PUBLIC_LONG_WORD_SOURCE_ROAD !== "0";
}

export function pptxLongDeckSourceRoadEnabled(): boolean {
  try {
    const browser = typeof window !== "undefined" ? window.localStorage?.getItem("mrai.pptxLongDeckSourceRoad") : null;
    if (browser === "0") return false;
    if (browser === "1") return true;
  } catch {
    // storage unavailable: the build decides
  }
  return process.env.NEXT_PUBLIC_PPTX_LONG_DECK_SOURCE_ROAD !== "0";
}
