/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A saved PDF author attempt, read back from the page's recovery JSON: the
 * source it was made against and the journal it ends in. Reading grants
 * nothing; rebuilding the plan to check it (`preparePdfNativeAuthorTextRecovery`)
 * runs on the server's engine host.
 */
import type { PageAuthorOperationInput } from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import type { PdfEngineSourceIdentity } from "./engine/contracts";
import type { PdfSharedStateRecord } from "./engine/shared-state-bridge";

export const isPdfNativeRecoveryRecord = (value: unknown): value is Record<string, JsonValue> =>
  !!value && typeof value === "object" && !Array.isArray(value);

/** Recovery JSON describes an attempt. It grants no lease, saved status or
 * history authority. Callers must reacquire current scope and verify bytes. */
export function readPdfNativeAuthorRecovery(input: PageAuthorOperationInput, pageId: string) {
  const value = input.recovery;
  if (input.family !== "pdf-native-transaction" || !isPdfNativeRecoveryRecord(value) ||
    value.schemaVersion !== "mrai.pdf-native-author-attempt.v1" || value.pageId !== pageId ||
    !isPdfNativeRecoveryRecord(value.source) || !Array.isArray(value.journal) || typeof value.previousJournalAbsent !== "boolean") return null;
  const source = value.source;
  if (typeof source.resourceId !== "string" || !source.resourceId ||
    typeof source.resourceVersionId !== "string" || !source.resourceVersionId ||
    typeof source.sourceSha256 !== "string" || !/^[a-f0-9]{64}$/.test(source.sourceSha256)) return null;
  const previousJournal = value.previousJournalAbsent ? undefined : value.previousJournal;
  if (value.previousJournalAbsent ? value.previousJournal !== null : !Array.isArray(previousJournal)) return null;
  const prefix = previousJournal ?? [];
  if (!Array.isArray(prefix) || value.journal.length !== prefix.length + 1 ||
    !pageDocumentJsonEquals(value.journal.slice(0,-1),prefix)) return null;
  const head = value.journal.at(-1);
  if (!isPdfNativeRecoveryRecord(head) || head.id !== input.id || head.transactionId !== input.id ||
    value.journal.filter(item=>isPdfNativeRecoveryRecord(item)&&item.id===input.id).length !== 1 ||
    !isPdfNativeRecoveryRecord(head.source) || !isPdfNativeRecoveryRecord(head.derivedCandidate) ||
    head.source.resourceId !== source.resourceId || head.source.resourceVersionId !== source.resourceVersionId ||
    head.source.sha256 !== source.sourceSha256 || typeof head.derivedCandidate.sha256 !== "string" ||
    !/^[a-f0-9]{64}$/.test(head.derivedCandidate.sha256)) return null;
  return {source: source as unknown as PdfEngineSourceIdentity,
    journal:value.journal,previousJournal,record:head as PdfSharedStateRecord,
    outputSha256:head.derivedCandidate.sha256};
}

