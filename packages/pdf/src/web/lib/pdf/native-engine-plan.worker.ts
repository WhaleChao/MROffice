/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { readPdfEngineCommand, refusedPdfEngineResult, type PdfEngineResult, type PdfEngineSourceIdentity } from "./engine/contracts";
import {
  createMraiNativePdfAdapter,
  planMraiNativeCommand,
  type MraiNativePlanOptions,
} from "./engine/adapters/mrai-native";
import { preparePdfNativeSourceCheckpoint } from "./engine/native/source-checkpoint";
import { pdfNativePropertyPlanAnswersShared, warmMraiNativeParagraphFormatPlans, type PdfNativeListTextSelectionEvidence } from "./engine/native/list-text-fidelity";
import { MAX_MRAI_NATIVE_SOURCE_BYTES, sha256PdfBytes } from "./engine/native/transaction-error";
import { pdfNativePreparedAnswersShared } from "./engine/native/content-stream-transaction";
import { fetchVerifiedPdfUnicodeFallbackFont } from "./unicode-font-asset";
import type { PdfSharedStateRecord } from "./engine/shared-state-bridge";

export interface PdfNativeEnginePlanWorkerRequest {
  readonly requestId: string;
  /** Only read the bytes for the paragraph plans to come (answered "failed",
   * reason "warmed"). */
  readonly kind?: "warm-paragraph";
  /** With "warm-paragraph": the box whose paragraph is read ahead. */
  readonly selection?: PdfNativeListTextSelectionEvidence | null;
  readonly sourceBytes: ArrayBuffer;
  readonly source: PdfEngineSourceIdentity;
  readonly command: unknown;
  readonly options: Record<string, unknown>;
}

export type PdfNativeEnginePlanWorkerResponse =
  | Readonly<{ requestId: string; status: "planned"; record: PdfSharedStateRecord; expectedCandidateSha256: string;
      candidateBytes: Uint8Array; phases?: PdfNativeEnginePlanPhases }>
  | Readonly<{ requestId: string; status: "refused"; result: PdfEngineResult }>
  | Readonly<{ requestId: string; status: "failed"; reason: string }>;

/** Where a plan's time went in the Worker (ms), and whether a planner answered
 * it with a plan already made (the editor's preview of the same edit). */
export type PdfNativeEnginePlanPhases = Readonly<{ digest: number; checkpoint: number; plan: number; shared: boolean;
  /** A preparation or reopening of the candidate already made answered it. */
  prepared: boolean }>;

const capability = createMraiNativePdfAdapter({
  loadPinnedSource: () => {
    throw new Error("capability check only");
  },
  sharedStateBridge: null as never,
});

let fallbackFont: Promise<Uint8Array> | null = null;
/** The bundled Unicode fallback font, fetched from this origin and verified
 * against its pinned digest (the engine host reads the same file). */
function loadFallbackFont(): Promise<Uint8Array> {
  fallbackFont ??= (async () => {
    const loaded = await fetchVerifiedPdfUnicodeFallbackFont();
    if (!loaded) throw new Error("The bundled Unicode font does not match its digest.");
    return new Uint8Array(loaded);
  })();
  fallbackFont.catch(() => { fallbackFont = null; });
  return fallbackFont.then((bytes) => Uint8Array.from(bytes));
}

function startsWith(bytes: Uint8Array, prefix: Uint8Array): boolean {
  if (bytes.length < prefix.length) return false;
  for (let index = 0; index < prefix.length; index += 1) if (bytes[index] !== prefix[index]) return false;
  return true;
}

/** Testable: plans one request and zeroizes what it was given. */
export async function planAndZeroizePdfNativeEngineCommand(
  input: PdfNativeEnginePlanWorkerRequest,
): Promise<PdfNativeEnginePlanWorkerResponse> {
  const bytes = new Uint8Array(input.sourceBytes);
  try {
    if (input.kind === "warm-paragraph") {
      if (bytes.length && bytes.length <= MAX_MRAI_NATIVE_SOURCE_BYTES) await warmMraiNativeParagraphFormatPlans(bytes, input.selection ?? null);
      return { requestId: input.requestId, status: "failed", reason: "warmed" };
    }
    const startedAt = performance.now();
    if (!bytes.length || bytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES ||
      (await sha256PdfBytes(bytes)) !== input.source.sourceSha256) {
      return { requestId: input.requestId, status: "failed", reason: "source-mismatch" };
    }
    const command = readPdfEngineCommand(input.command);
    if (!command) return { requestId: input.requestId, status: "failed", reason: "invalid-command" };
    if (command.sourceSha256 !== input.source.sourceSha256 || command.resourceId !== input.source.resourceId ||
      command.resourceVersionId !== input.source.resourceVersionId) {
      return { requestId: input.requestId, status: "failed", reason: "source-mismatch" };
    }
    if (!capability.canHandle(command)) {
      return { requestId: input.requestId, status: "refused",
        result: refusedPdfEngineResult("adapter-rejected", "mrai-native cannot handle this command.", command.transactionId) };
    }
    const options = { ...input.options } as MraiNativePlanOptions & { sourceCheckpoint?: unknown };
    const identity = { resourceId: input.source.resourceId, resourceVersionId: input.source.resourceVersionId,
      sourceSha256: input.source.sourceSha256 };
    let sourceCheckpoint: MraiNativePlanOptions["sourceCheckpoint"];
    const digestedAt = performance.now();
    if (options.sourceCheckpoint !== undefined) {
      const asked = (options.sourceCheckpoint && typeof options.sourceCheckpoint === "object"
        ? options.sourceCheckpoint : {}) as { xrefSections?: unknown };
      try {
        sourceCheckpoint = await preparePdfNativeSourceCheckpoint(identity, bytes);
      } catch (error) {
        return { requestId: input.requestId, status: "refused",
          result: refusedPdfEngineResult("unsupported-document",
            error instanceof Error ? error.message : "The PDF checkpoint is unavailable.", command.transactionId) };
      }
      if (asked.xrefSections !== undefined && asked.xrefSections !== sourceCheckpoint.xrefSections) {
        return { requestId: input.requestId, status: "refused",
          result: refusedPdfEngineResult("adapter-rejected", "The source checkpoint is stale or invalid.", command.transactionId) };
      }
    }
    delete options.sourceCheckpoint;
    const checkpointedAt = performance.now();
    const sharedBefore = pdfNativePropertyPlanAnswersShared();
    const preparedBefore = pdfNativePreparedAnswersShared();
    const outcome = await planMraiNativeCommand(command, { source: identity },
      { ...options, ...(sourceCheckpoint ? { sourceCheckpoint } : {}), loadNativeFontExtensionFont: loadFallbackFont },
      () => bytes);
    const phases: PdfNativeEnginePlanPhases = Object.freeze({ digest: Math.round(digestedAt - startedAt),
      checkpoint: Math.round(checkpointedAt - digestedAt), plan: Math.round(performance.now() - checkpointedAt),
      shared: pdfNativePropertyPlanAnswersShared() > sharedBefore,
      prepared: pdfNativePreparedAnswersShared() > preparedBefore });
    if (outcome.status !== "planned") return { requestId: input.requestId, status: "refused", result: outcome };
    try {
      if (!outcome.candidate) return { requestId: input.requestId, status: "failed", reason: "no-candidate" };
      if (outcome.candidate.sourceLength !== bytes.byteLength || !startsWith(outcome.candidate.bytes, bytes)) {
        return { requestId: input.requestId, status: "failed", reason: "candidate-not-incremental" };
      }
      if (outcome.candidate.sha256 !== outcome.expectedCandidateSha256) {
        return { requestId: input.requestId, status: "failed", reason: "candidate-digest" };
      }
      return { requestId: input.requestId, status: "planned",
        record: JSON.parse(JSON.stringify(outcome.record)) as PdfSharedStateRecord,
        expectedCandidateSha256: outcome.expectedCandidateSha256,
        candidateBytes: Uint8Array.from(outcome.candidate.bytes), phases };
    } finally {
      outcome.dispose();
    }
  } catch (error) {
    return { requestId: input.requestId, status: "failed",
      reason: error instanceof Error ? error.message.slice(0, 200) : "plan-failed" };
  } finally {
    bytes.fill(0);
  }
}
