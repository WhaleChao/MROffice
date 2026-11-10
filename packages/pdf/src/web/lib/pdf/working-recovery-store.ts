/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { captureSavedPageHotCacheScope, isSavedPageHotCacheScopeCurrent } from "@mrnote/core/web/lib/saved-page-hot-cache.ts";
import type { PdfWorkingRecovery } from "./native-working-session";

const PREFIX = "mrnote:pdf-working-recovery:v1:";
const MAX_SERIALIZED_CHARS = 4 * 1024 * 1024;

/** Uses the app's principal scope and tab-local storage, but a distinct key:
 * ordinary page-save ACKs cannot retire newer unsubmitted PDF operations.
 * This stores replay recipes, never credentials, PDF bytes or writer leases. */
export function createPdfWorkingRecoveryStore(input: {
  pageId: string;
  dataItemId: string;
  storage?: Storage;
}) {
  const scope = captureSavedPageHotCacheScope();
  if (!scope.principal || !input.pageId || !input.dataItemId) {
    throw new Error("PDF local recovery requires an authenticated document scope.");
  }
  const storage = input.storage ?? window.sessionStorage;
  const key = PREFIX + [scope.principal,input.pageId,input.dataItemId].map(encodeURIComponent).join(":");
  let expected = storage.getItem(key);
  const check = () => {
    if (!isSavedPageHotCacheScopeCurrent(scope)) throw new Error("PDF local recovery account changed.");
    if (storage.getItem(key) !== expected) throw new Error("Another PDF window changed this recovery queue.");
  };
  const decode = (raw: string | null): PdfWorkingRecovery | null => {
    if (raw === null) return null;
    if (raw.length > MAX_SERIALIZED_CHARS) throw new Error("PDF local recovery exceeds its storage budget.");
    const value = JSON.parse(raw);
    if (!value || value.version !== 1 || value.principal !== scope.principal ||
      value.pageId !== input.pageId || value.dataItemId !== input.dataItemId ||
      value.recovery?.schemaVersion !== "mrai.pdf-working-recovery.v1" ||
      value.recovery.pageId !== input.pageId || !Array.isArray(value.recovery.operations) ||
      value.recovery.operations.length > 32) throw new Error("PDF local recovery identity is invalid.");
    // This shape check is not native authority. Restoration independently
    // proves every recipe against its real saved source and output hash.
    return value.recovery as PdfWorkingRecovery;
  };
  return {
    read(): PdfWorkingRecovery | null { check(); return decode(expected); },
    write(recovery: PdfWorkingRecovery): void {
      check();
      const next = JSON.stringify({version:1,principal:scope.principal,pageId:input.pageId,
        dataItemId:input.dataItemId,recovery});
      decode(next);
      // setItem is atomic on quota errors. Do not clear the previous valid
      // queue if serialization, quota or scope validation fails.
      storage.setItem(key,next);
      expected = next;
    },
    /** The reader chose 不保存 for edits AutoSave off held: the queue goes
     * with the window, as WPS discards an unsaved document. */
    discardUnsaved(): void {
      if (!isSavedPageHotCacheScopeCurrent(scope)) return;
      storage.removeItem(key); expected = null;
    },
    clearConfirmed(): void {
      check();
      const current = decode(expected);
      if (current?.operations.length) throw new Error("Unconfirmed PDF operations cannot be cleared.");
      storage.removeItem(key);expected=null;
    },
  };
}
