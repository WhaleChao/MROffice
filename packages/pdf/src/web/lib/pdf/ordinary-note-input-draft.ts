/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  captureSavedPageHotCacheScope,
  isSavedPageHotCacheScopeCurrent,
  type SavedPageHotCacheScope,
} from "@mrnote/core/web/lib/saved-page-hot-cache.ts";
import { sanitizeInsertedPageId } from "@mrnote/core/web/lib/pdf/inserted-page-identity.ts";

export interface PdfOrdinaryNoteInputDraft {
  readonly id: string;
  readonly text: string;
  readonly page?: number;
  /** The inserted page the note is on (`page` is then the page before it):
   * lib/pdf/inserted-page-identity. */
  readonly insertedPageId?: string;
  readonly createdAt?: string;
  readonly submittedText?: string;
}
export interface PdfOrdinaryNoteInputBinding {
  readonly key: string;
  readonly pageId: string;
  readonly dataItemId: string;
  readonly scope: SavedPageHotCacheScope;
}
const MAX_TEXT = 1024 * 1024;
const PREFIX = "mrnote:pdf-note-input:v1:";

/** Match the existing page-draft lifetime: tab-scoped sessionStorage, with
 * the existing actor fingerprint plus exact page/source binding. No page
 * JSON, native transaction or automatic replay is stored here. */
export function bindPdfOrdinaryNoteInput(pageId: string, dataItemId: string | null): PdfOrdinaryNoteInputBinding | null {
  const scope = captureSavedPageHotCacheScope();
  if (!scope.principal || !pageId || !dataItemId) return null;
  return { scope, pageId, dataItemId,
    key: PREFIX + [scope.principal, pageId, dataItemId].map(encodeURIComponent).join(":") };
}

function validDraft(value: unknown): value is PdfOrdinaryNoteInputDraft {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const draft = value as Partial<PdfOrdinaryNoteInputDraft>;
  return typeof draft.id === "string" && /^[a-zA-Z0-9-]{1,128}$/.test(draft.id) &&
    typeof draft.text === "string" && draft.text.length <= MAX_TEXT &&
    (draft.page === undefined || (Number.isSafeInteger(draft.page) && draft.page > 0 && draft.page <= 10_000)) &&
    (draft.insertedPageId === undefined || sanitizeInsertedPageId(draft.insertedPageId) !== null) &&
    (draft.createdAt === undefined || (typeof draft.createdAt === "string" && draft.createdAt.length <= 40 && Number.isFinite(Date.parse(draft.createdAt)))) &&
    (draft.submittedText === undefined || (typeof draft.submittedText === "string" && draft.submittedText.length <= MAX_TEXT));
}

export function readPdfOrdinaryNoteInput(binding: PdfOrdinaryNoteInputBinding): { draft: PdfOrdinaryNoteInputDraft | null; available: boolean } {
  if (!isSavedPageHotCacheScopeCurrent(binding.scope)) return { draft: null, available: false };
  try {
    const raw = window.sessionStorage.getItem(binding.key);
    if (raw === null) return { draft: null, available: true };
    if (raw.length > MAX_TEXT * 3) return { draft: null, available: false };
    const record = JSON.parse(raw);
    if (record?.version !== 1 || record.pageId !== binding.pageId || record.dataItemId !== binding.dataItemId || !validDraft(record.draft)) {
      return { draft: null, available: false };
    }
    return { draft: record.draft, available: true };
  } catch { return { draft: null, available: false }; }
}

export function writePdfOrdinaryNoteInput(binding: PdfOrdinaryNoteInputBinding | null, draft: PdfOrdinaryNoteInputDraft | null): boolean {
  if (!binding || !isSavedPageHotCacheScopeCurrent(binding.scope)) return false;
  try {
    if (!draft || draft.text.length === 0) window.sessionStorage.removeItem(binding.key);
    else {
      if (!validDraft(draft)) return false;
      const serialized = JSON.stringify({ version: 1, pageId: binding.pageId, dataItemId: binding.dataItemId, draft });
      if (serialized.length > MAX_TEXT * 3) return false;
      window.sessionStorage.setItem(binding.key, serialized);
    }
    return true;
  } catch { return false; }
}
