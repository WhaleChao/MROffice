/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PageItem, PageType } from "./notebook-sdk";

/** Digest a clear signs when its saved baseline was already empty. An empty
 * baseline says nothing about a non-empty stored body (a blank editor after a
 * failed load looks exactly like it), so this never matches a real page. The
 * proof becomes valid only when the autosave hook carries it across this
 * editor's own later save, whose receipt supplies the real digest. */
export const UNPROVEN_EMPTY_SOURCE_SHA256 = "0".repeat(64);

export interface PageSaveAcknowledgementDraft {
  title?: string;
  page_type?: PageType;
  content_json?: Record<string, unknown>;
}

export function pageDocumentJsonEquals(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    return (
      Array.isArray(left) &&
      Array.isArray(right) &&
      left.length === right.length &&
      left.every((value, index) => pageDocumentJsonEquals(value, right[index]))
    );
  }
  if (
    !left ||
    !right ||
    typeof left !== "object" ||
    typeof right !== "object"
  ) {
    return false;
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord);
  const rightKeys = Object.keys(rightRecord);
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key) =>
        Object.prototype.hasOwnProperty.call(rightRecord, key) &&
        pageDocumentJsonEquals(leftRecord[key], rightRecord[key]),
    )
  );
}

/** A page body as the API stores it. PostgreSQL jsonb cannot hold NUL, so the
 * API replaces U+0000 with U+FFFD in every string before the write
 * (`sanitize_jsonb_text`). A PDF glyph whose ToUnicode maps to U+0000 puts
 * that codepoint into native text evidence, so the stored body differs from
 * the authored one by exactly that substitution. Values without a NUL keep
 * their identity. */
export function pageStoredJsonProjection<T>(value: T): T {
  if (typeof value === "string") {
    return (value.includes("\u0000")
      ? value.replaceAll("\u0000", "\ufffd")
      : value) as T;
  }
  if (Array.isArray(value)) {
    let changed = false;
    const next = value.map((item) => {
      const projected = pageStoredJsonProjection(item);
      if (projected !== item) changed = true;
      return projected;
    });
    return (changed ? next : value) as T;
  }
  if (value && typeof value === "object") {
    let changed = false;
    const next: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(value)) {
      const projected = pageStoredJsonProjection(item);
      if (projected !== item) changed = true;
      next[key] = projected;
    }
    return (changed ? next : value) as T;
  }
  return value;
}

/** Whether the API stores `left` and `right` as the same JSON. */
export function pageStoredJsonEquals(left: unknown, right: unknown): boolean {
  return pageDocumentJsonEquals(
    pageStoredJsonProjection(left),
    pageStoredJsonProjection(right),
  );
}

/**
 * Validate the transport-level receipt returned by the exact PATCH request the
 * caller just awaited.
 *
 * The API is allowed to canonicalize authored document JSON before returning
 * it (for example, assigning missing/duplicate block ids). Requiring byte- or
 * object-equality with the submitted draft therefore turns a successful,
 * durable write into a false "unsaved" state. Request identity is already
 * provided by the awaited promise; the receipt only needs to prove that it is
 * for the same page, that its optimistic version did not move backwards, and
 * that non-canonicalized scalar mutations were applied.
 *
 * Equality at the base version is valid: the endpoint intentionally returns it
 * when the submitted draft was already durable. Callers that need to associate
 * this receipt with a particular local editor generation must separately
 * compare against the draft captured for that request.
 */
export function pageSaveResponseAcknowledgesRequest(
  savedPage: PageItem,
  pageId: string,
  requestBaseContentVersion: number,
  draft: PageSaveAcknowledgementDraft,
): boolean {
  const savedContentVersion = savedPage.content_version;
  if (
    savedPage.id !== pageId ||
    typeof savedContentVersion !== "number" ||
    !Number.isSafeInteger(savedContentVersion) ||
    savedContentVersion < requestBaseContentVersion
  ) {
    return false;
  }
  // Content is canonicalized by the server, but these scalar authored fields
  // are not. Empty titles mean "no rename" at the API boundary.
  if (
    typeof draft.title === "string" &&
    draft.title.length > 0 &&
    savedPage.title !== draft.title
  ) {
    return false;
  }
  if (
    draft.page_type !== undefined &&
    savedPage.page_type !== draft.page_type
  ) {
    return false;
  }
  return true;
}

/**
 * A saved page's `editable_page_patchback` receipt is read as data only.
 *
 * A receipt the server answered is a plain enumerable value. An object that
 * offers the property through an accessor is not a receipt the page may act
 * on, and reading it would run that accessor; the window has refused such a
 * receipt since bfc171aae8, and every other comparison of the receipt goes
 * through this reader so none of them invokes it either.
 */
export function pdfPatchbackReceiptDataValue(
  page: object | null | undefined,
): unknown {
  if (!page || typeof page !== "object") return null;
  const descriptor = Object.getOwnPropertyDescriptor(
    page,
    "editable_page_patchback",
  );
  if (!descriptor || !("value" in descriptor) || !descriptor.enumerable) {
    return null;
  }
  return descriptor.value ?? null;
}

/**
 * The page with its receipt as data. A page whose receipt is a data property
 * comes back as it is; one whose receipt is an accessor (or hidden) comes
 * back as a copy of its own data properties with the receipt absent, so no
 * later clone, cache or comparison of the page runs that accessor.
 */
export function pdfPatchbackReceiptDataPage<T extends object>(page: T): T {
  const descriptor = Object.getOwnPropertyDescriptor(
    page,
    "editable_page_patchback",
  );
  if (!descriptor || ("value" in descriptor && descriptor.enumerable)) {
    return page;
  }
  const copy: Record<string, unknown> = {};
  for (const [key, own] of Object.entries(
    Object.getOwnPropertyDescriptors(page),
  )) {
    if (key === "editable_page_patchback") continue;
    if ("value" in own && own.enumerable) copy[key] = own.value;
  }
  copy.editable_page_patchback = null;
  return copy as T;
}
