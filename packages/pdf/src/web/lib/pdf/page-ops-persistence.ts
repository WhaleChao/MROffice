/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Best-effort client persistence for PDF READER page operations.
 *
 * Page edits (rotate / delete / reorder / insert blank) produced in the
 * read-only reference viewer are persisted per `dataItemId` so they survive a
 * window close / reopen. The viewer's annotations (highlights / sticky notes)
 * persist to the server's `/data-items/{id}/annotations` endpoint, but that
 * endpoint only accepts `pdf_highlight` / `pdf_note` types — there is no
 * server-side `pdf_page_edits` record yet — so reader page edits mirror that
 * same per-data-item keying against `localStorage` instead. This is additive
 * and non-destructive: a missing / corrupt entry simply yields no edits, and
 * the native editor (PdfPageWindow) keeps using its server-backed page-snapshot
 * path unchanged.
 *
 * The serialized shape is exactly {@link serializePageEditState}'s output, which
 * round-trips through `normalizePageEditState`, so the same bytes could later be
 * lifted into a real server record without a migration.
 *
 * The key builder + (de)serialization are pure and unit-tested; the read/write
 * wrappers swallow every storage error (private-mode quotas, disabled storage,
 * SSR) the same way the highlight/note loaders swallow network errors.
 */

import {
  isEmptyPageEditState,
  normalizePageEditState,
  serializePageEditState,
  type PageEditState,
} from "@mrnote/core/web/lib/pdf/page-ops.ts";

/** Versioned namespace so a future shape change can invalidate old entries. */
export const PAGE_EDITS_STORAGE_PREFIX = "mrnote.pdf.page-edits.v1";

/** Deterministic storage key for a given data item. Empty id ⇒ "". */
export function pageEditsStorageKey(dataItemId: string | null | undefined): string {
  if (!dataItemId || typeof dataItemId !== "string") return "";
  const trimmed = dataItemId.trim();
  if (!trimmed) return "";
  return `${PAGE_EDITS_STORAGE_PREFIX}:${trimmed}`;
}

/** Serialize a state to the JSON string persisted under the storage key. */
export function serializePageEdits(state: PageEditState): string {
  return JSON.stringify(serializePageEditState(state));
}

/**
 * Parse a previously-serialized blob back into a normalized state. Returns null
 * for empty / malformed input so callers can fall back to props cleanly.
 */
export function deserializePageEdits(raw: string | null): PageEditState | null {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!parsed || typeof parsed !== "object") return null;
    return normalizePageEditState(parsed as Record<string, never>);
  } catch {
    return null;
  }
}

/** Read + normalize the persisted edits for a data item (best-effort). */
export function loadPersistedPageEdits(
  dataItemId: string | null | undefined,
): PageEditState | null {
  const key = pageEditsStorageKey(dataItemId);
  if (!key || typeof window === "undefined") return null;
  try {
    return deserializePageEdits(window.localStorage.getItem(key));
  } catch {
    return null;
  }
}

/**
 * Persist (or clear, when empty) the edits for a data item (best-effort).
 * Removing the key on an empty state keeps storage tidy and means "no edits"
 * never shadows future prop-driven edits.
 */
export function persistPageEdits(
  dataItemId: string | null | undefined,
  state: PageEditState,
): void {
  const key = pageEditsStorageKey(dataItemId);
  if (!key || typeof window === "undefined") return;
  try {
    if (isEmptyPageEditState(state)) {
      window.localStorage.removeItem(key);
      return;
    }
    window.localStorage.setItem(key, serializePageEdits(state));
  } catch {
    /* ignore — page edits are best-effort, like highlights/notes */
  }
}
