/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  isPdfNativeObjectRestoreEvidence,
  isPdfNativeRestoreEvidence,
  isPdfNativeRevisionRestoreEvidence,
  type PdfNativeObjectRestoreEvidence,
  type PdfNativeRestoreEvidence,
  type PdfNativeRevisionRestoreEvidence,
} from "./byte-restore-model";

/**
 * A source consolidation: the server's own rewrite of a native revision
 * chain as one xref section, made when a native write's output nears the
 * parser's ceiling (32 sections). It is never planned here; the server
 * appends it to the page's journal as the one record the next native
 * write chains on, sourced from the client's verified candidate bytes and
 * derived to the consolidated file that became the version. Editor history
 * goes on across it: the records before it are gone with the chain, and the
 * record keeps, for each restorable one, the restore evidence its undo and
 * redo read instead (`restorableHistory`, readPdfNativeRestoreEvidence).
 */
export const MRAI_NATIVE_SOURCE_CONSOLIDATION_ACTION = "native.source.consolidate";
export const MRAI_NATIVE_SOURCE_CONSOLIDATION_KIND = "source-consolidation";
export const MRAI_NATIVE_SOURCE_CONSOLIDATION_SCHEMA_VERSION =
  "mrai.pdf-native-source-consolidation.v1";

/** Whether a journal record is the server's consolidation of the chain. */
export function isPdfNativeSourceConsolidationRecord(value: unknown): boolean {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const record = value as { action?: unknown; kind?: unknown; schemaVersion?: unknown };
  return (
    record.action === MRAI_NATIVE_SOURCE_CONSOLIDATION_ACTION &&
    record.kind === MRAI_NATIVE_SOURCE_CONSOLIDATION_KIND &&
    record.schemaVersion === MRAI_NATIVE_SOURCE_CONSOLIDATION_SCHEMA_VERSION
  );
}





export function pdfNativeJournalEndsInConsolidation(collection: unknown): boolean {
  return (
    Array.isArray(collection) &&
    collection.length > 0 &&
    isPdfNativeSourceConsolidationRecord(collection[collection.length - 1])
  );
}

/**
 * The digest of the consolidated file a persisted save produced for the
 * candidate a writer planned, read from the persisted page's journal: the
 * terminal consolidation record that names that candidate. Null when the
 * save did not consolidate (the candidate itself became the version).
 */
export function readPdfNativeConsolidatedOutputSha256(
  persisted: unknown,
  candidateSha256: string,
): string | null {
  const content = (persisted as { content_json?: unknown } | null | undefined)?.content_json;
  const collection = (content as { pdfNativeTransactions?: unknown } | null | undefined)
    ?.pdfNativeTransactions;
  if (!pdfNativeJournalEndsInConsolidation(collection)) return null;
  const terminal = (collection as unknown[])[(collection as unknown[]).length - 1] as {
    derivedCandidate?: { sha256?: unknown };
    sourceConsolidation?: { candidateSha256?: unknown };
  };
  if (terminal.sourceConsolidation?.candidateSha256 !== candidateSha256) return null;
  const output = terminal.derivedCandidate?.sha256;
  return typeof output === "string" ? output : null;
}

/**
 * Whether a saved journal is the server's consolidation of the journal a
 * save submitted: exactly one record, the consolidation, originating from
 * the submitted terminal record and chaining from its candidate digest.
 * Pure structure; the receipt and bytes are authenticated by the caller.
 */
export function pdfNativeJournalConsolidates(saved: unknown, submitted: unknown): boolean {
  if (!Array.isArray(saved) || !Array.isArray(submitted) || saved.length !== 1) {
    return false;
  }
  const consolidation = saved[0];
  if (!isPdfNativeSourceConsolidationRecord(consolidation)) return false;
  const terminal = submitted[submitted.length - 1] as
    | { id?: unknown; derivedCandidate?: { sha256?: unknown } }
    | null
    | undefined;
  const record = consolidation as {
    source?: { sha256?: unknown };
    sourceConsolidation?: { originTransactionId?: unknown; candidateSha256?: unknown };
  };
  return Boolean(
    terminal &&
      typeof terminal.id === "string" &&
      typeof terminal.derivedCandidate?.sha256 === "string" &&
      record.sourceConsolidation?.originTransactionId === terminal.id &&
      record.sourceConsolidation?.candidateSha256 === terminal.derivedCandidate.sha256 &&
      record.source?.sha256 === terminal.derivedCandidate.sha256,
  );
}

/**
 * Where a history seed reads the write it was planned for. Normally the saved
 * journal, whose last record is that write, with the receipt's output. When
 * the save consolidated the chain (the saved journal is the one
 * consolidation record of the journal this window submitted, and it became
 * the receipt's output), the write is the submitted journal's last record and
 * its output the candidate the consolidation names: the record is gone from
 * the page, and its Undo restores from the evidence the consolidation kept
 * for it.
 */
export function readPdfNativeHistorySeedJournal(
  saved: unknown,
  submitted: unknown,
  outputSha256: string,
): Readonly<{ collection: unknown; outputSha256: string }> {
  if (Array.isArray(saved) && pdfNativeJournalConsolidates(saved, submitted)) {
    const consolidation = saved[0] as { derivedCandidate?: { sha256?: unknown } };
    const terminal = (submitted as Array<{ derivedCandidate?: { sha256?: unknown } }>).at(-1);
    const candidate = terminal?.derivedCandidate?.sha256;
    if (consolidation.derivedCandidate?.sha256 === outputSha256 && typeof candidate === "string") {
      return Object.freeze({ collection: submitted, outputSha256: candidate });
    }
  }
  return Object.freeze({ collection: saved, outputSha256 });
}

/**
 * The revision depth (xref sections of the current source) at which a window
 * that opens a document idle asks the server to consolidate the chain in the
 * background, so the rewrite the save path would otherwise run in the middle
 * of a save near the ceiling (NATIVE_SOURCE_CONSOLIDATION_DEPTH, 24 on the
 * server) happens while nobody is editing and no editor history is lost.
 */
export const PDF_NATIVE_SOURCE_CONSOLIDATION_OPEN_DEPTH = 12;

/**
 * The depth of the source a saved journal's terminal record produced, read
 * from the journal alone: a consolidation is one section; a checkpointed
 * write sourced `xrefSections` and appended one. Null for a legacy journal
 * without checkpoints or an empty one, where nothing is known. The server
 * counts the real sections before it rewrites anything; this only decides
 * whether to ask.
 */
export function readPdfNativeJournalRevisionDepth(collection: unknown): number | null {
  if (!Array.isArray(collection) || collection.length === 0) return null;
  const terminal = collection[collection.length - 1] as {
    sourceCheckpoint?: { schemaVersion?: unknown; xrefSections?: unknown };
  } | null | undefined;
  if (!terminal || typeof terminal !== "object") return null;
  if (isPdfNativeSourceConsolidationRecord(terminal)) return 1;
  const sections = terminal.sourceCheckpoint?.xrefSections;
  if (terminal.sourceCheckpoint?.schemaVersion !== "mrai.pdf-native-source-checkpoint.v1" ||
    typeof sections !== "number" || !Number.isSafeInteger(sections) || sections < 1) return null;
  return sections + 1;
}

/**
 * The restore evidence a journal keeps for `originTransactionId`: one entry
 * of the consolidation record that starts the journal (its records and
 * revisions were compacted away); null when there is none or more than one.
 * Mirror of the server's read_native_restore_evidence.
 */
export function readPdfNativeRestoreEvidence(
  collection: unknown,
  originTransactionId: string,
): PdfNativeRestoreEvidence | null {
  const kept = keptRestoreEvidence(collection, originTransactionId);
  return isPdfNativeRestoreEvidence(kept) ? kept : null;
}

/**
 * The object restore evidence a journal keeps for a font-extension or
 * font-reuse edit `originTransactionId` the consolidation compacted away (as
 * readPdfNativeRestoreEvidence). Mirror of the server's
 * read_native_object_restore_evidence.
 */
export function readPdfNativeObjectRestoreEvidence(
  collection: unknown,
  originTransactionId: string,
): PdfNativeObjectRestoreEvidence | null {
  const kept = keptRestoreEvidence(collection, originTransactionId);
  return isPdfNativeObjectRestoreEvidence(kept) ? kept : null;
}

/**
 * The revision restore evidence a journal keeps for a step published as its
 * bytes (a bytes-publish record `originTransactionId` the consolidation
 * compacted away), as readPdfNativeRestoreEvidence: what
 * planMraiNativeObjectRestoreFromEvidence plans its Undo and Redo from.
 * Mirror of the server's read_native_revision_restore_evidence.
 */
export function readPdfNativeRevisionRestoreEvidence(
  collection: unknown,
  originTransactionId: string,
): PdfNativeRevisionRestoreEvidence | null {
  const kept = keptRestoreEvidence(collection, originTransactionId);
  return isPdfNativeRevisionRestoreEvidence(kept) ? kept : null;
}

function keptRestoreEvidence(collection: unknown, originTransactionId: string): unknown {
  if (!Array.isArray(collection) || collection.length === 0) return null;
  const first = collection[0] as { restorableHistory?: unknown } | null | undefined;
  if (!isPdfNativeSourceConsolidationRecord(first) || !Array.isArray(first?.restorableHistory)) return null;
  const matches = first.restorableHistory.filter((item) =>
    item && typeof item === "object" && (item as { id?: unknown }).id === originTransactionId);
  return matches.length === 1 ? matches[0] : null;
}
