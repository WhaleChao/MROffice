/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The editor's half of a native PDF command: write the planned record through
 * the shared-state bridge (the durable `saveNow` seam) and accept it only on
 * an exact lossless receipt for the candidate the plan named.
 *
 * The planning half (`planMraiNativeCommand` in `mrai-native.ts`) runs on the
 * server's engine host; this module holds no PDF writing knowledge and is the
 * part of the native adapter the browser keeps.
 */

import { readPdfNativeConsolidatedOutputSha256 } from "../native/source-consolidation";
import {
  PdfEngineNotWritableError,
  refusedPdfEngineResult,
  type PdfEngineAdapterContext,
  type PdfEngineCommand,
  type PdfEngineResult,
} from "../contracts";
import type { PdfSharedStateBridge, PdfSharedStateRecord } from "../shared-state-bridge";
import {
  isLosslessDurableNativeSaveReceipt,
  readNativeSaveReceiptSnapshot,
} from "@mrnote/core/web/lib/native-save-receipt.ts";

export const MRAI_NATIVE_TRANSACTIONS_COLLECTION = "pdfNativeTransactions";

export async function commitMraiNativePlannedRecord(input: {
  readonly command: PdfEngineCommand;
  readonly context: PdfEngineAdapterContext;
  readonly record: PdfSharedStateRecord;
  readonly expectedCandidateSha256: string;
  readonly durable: boolean;
  readonly sharedStateBridge: PdfSharedStateBridge;
  readonly collection?: string;
}): Promise<PdfEngineResult> {
  const { command, context } = input;
  let committed: Awaited<ReturnType<PdfSharedStateBridge["commitLocalRecord"]>>;
  try {
    committed = await input.sharedStateBridge.commitLocalRecord({
      collection: input.collection ?? MRAI_NATIVE_TRANSACTIONS_COLLECTION,
      record: input.record,
      transactionId: command.transactionId,
    });
    if (committed.status !== "persisted") {
      return refusedPdfEngineResult(
        "adapter-rejected",
        "Native PDF mutation was not persisted as shared state.",
        command.transactionId,
      );
    }
  } catch (error) {
    if (error instanceof PdfEngineNotWritableError) {
      return refusedPdfEngineResult("not-writable", error.message, command.transactionId);
    }
    return refusedPdfEngineResult(
      "adapter-rejected",
      error instanceof Error ? error.message : "Native PDF shared-state persistence failed.",
      command.transactionId,
    );
  }

  if (input.durable) {
    if (committed.seam !== "saveNow") {
      return refusedPdfEngineResult(
        "adapter-rejected",
        "Native PDF mutation requires the durable saveNow seam.",
        command.transactionId,
      );
    }
    const durableReceipt = readNativeSaveReceiptSnapshot(committed.receipt);
    // Near the parser's revision ceiling the server rewrites the verified
    // candidate as one xref section and keeps that step in the persisted
    // journal (source-consolidation.ts): the version is then that rewrite,
    // and the receipt names it instead of the candidate.
    const consolidatedOutputSha256 = readPdfNativeConsolidatedOutputSha256(
      committed.receipt,
      input.expectedCandidateSha256,
    );
    if (
      !isLosslessDurableNativeSaveReceipt(durableReceipt, {
        expectedAdapter: "pdf",
        expectedResourceId: context.source.resourceId,
        expectedOutputSha256: consolidatedOutputSha256 ?? input.expectedCandidateSha256,
        previousResourceVersionId: context.source.resourceVersionId,
      }) ||
      typeof durableReceipt?.resourceVersionId !== "string" ||
      typeof durableReceipt.outputSha256 !== "string"
    ) {
      return refusedPdfEngineResult(
        "adapter-rejected",
        "Native PDF mutation did not receive an exact lossless PDF resource receipt.",
        command.transactionId,
      );
    }
    return Object.freeze({
      status: "applied" as const,
      transactionId: command.transactionId,
      sessionRevision: command.sessionRevision + 1,
      resourceVersionId: durableReceipt.resourceVersionId,
      sourceSha256: durableReceipt.outputSha256.toLowerCase(),
      undoToken: command.transactionId,
      warnings: Object.freeze([]),
      lossReport: Object.freeze([]),
    });
  }

  return Object.freeze({
    status: "applied" as const,
    transactionId: command.transactionId,
    sessionRevision: command.sessionRevision + 1,
    resourceVersionId: context.source.resourceVersionId,
    sourceSha256: context.source.sourceSha256,
    undoToken: command.transactionId,
    warnings: Object.freeze(["Candidate PDF bytes are derived state and were not durably uploaded."]),
    lossReport: Object.freeze([]),
  });
}
