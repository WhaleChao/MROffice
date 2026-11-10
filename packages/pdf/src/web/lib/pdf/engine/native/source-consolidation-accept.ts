/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  isPdfNativeSourceConsolidationRecord,
  pdfNativeJournalConsolidates,
  readPdfNativeConsolidatedOutputSha256,
} from "./source-consolidation";

const SHA256_PATTERN = /^[0-9a-f]{64}$/u;

/**
 * What an accepted save proves when the server consolidated the client's
 * own candidate: the version is `outputSha256` (`outputByteLength` bytes),
 * the one-section rewrite of the candidate `candidateSha256` this client
 * planned and verified. Consolidation keeps every object under its own
 * number with its bytes, so what the client established about the candidate
 * as a whole (its source risk profile, the page it painted) holds for the
 * version; the version's bytes are still downloaded and bound to the digest.
 */
export interface PdfNativeConsolidatedAcceptProof {
  readonly candidateSha256: string;
  readonly outputSha256: string;
  readonly outputByteLength: number;
}

function readCollection(value: unknown): unknown {
  const content = (value as { content_json?: unknown } | null | undefined)?.content_json;
  return (content as { pdfNativeTransactions?: unknown } | null | undefined)?.pdfNativeTransactions;
}

/**
 * The proof, or null when any part of it does not hold (the accept path then
 * does its full whole-document work):
 *
 * - the journal the client submitted ends in its own record (not a
 *   consolidation) whose derived candidate names the planned digest;
 * - the saved page's journal is exactly the server's consolidation of that
 *   submitted journal (pdfNativeJournalConsolidates), whose candidate is
 *   that planned digest;
 * - the consolidation's derived digest is the receipt's output digest, it
 *   differs from the candidate, and it names a positive byte length.
 */
export function readPdfNativeConsolidatedAcceptProof(input: {
  readonly savedPage: unknown;
  readonly submittedCollection: unknown;
  readonly receiptOutputSha256: string;
}): PdfNativeConsolidatedAcceptProof | null {
  const submitted = input.submittedCollection;
  if (!Array.isArray(submitted) || submitted.length === 0) return null;
  const terminal = submitted[submitted.length - 1] as
    | { derivedCandidate?: { sha256?: unknown } }
    | null
    | undefined;
  if (!terminal || isPdfNativeSourceConsolidationRecord(terminal)) return null;
  const candidateSha256 = terminal.derivedCandidate?.sha256;
  if (typeof candidateSha256 !== "string" || !SHA256_PATTERN.test(candidateSha256)) return null;
  const saved = readCollection(input.savedPage);
  if (!pdfNativeJournalConsolidates(saved, submitted)) return null;
  const outputSha256 = readPdfNativeConsolidatedOutputSha256(input.savedPage, candidateSha256);
  if (
    typeof outputSha256 !== "string" ||
    !SHA256_PATTERN.test(outputSha256) ||
    outputSha256 === candidateSha256 ||
    outputSha256 !== input.receiptOutputSha256
  ) {
    return null;
  }
  const outputByteLength = ((saved as unknown[])[0] as { derivedCandidate?: { byteLength?: unknown } })
    .derivedCandidate?.byteLength;
  if (
    typeof outputByteLength !== "number" ||
    !Number.isSafeInteger(outputByteLength) ||
    outputByteLength < 1
  ) {
    return null;
  }
  return Object.freeze({ candidateSha256, outputSha256, outputByteLength });
}
