/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The native PDF adapter the editor uses: the command is planned on the
 * server's engine host (`pdf-native/plan`, the same `planMraiNativeCommand`
 * the adapter used to run here) and only the record and the appended revision
 * come back. Committing stays here and unchanged: the record goes through the
 * shared-state bridge's durable saveNow seam and is accepted only on an exact
 * lossless receipt for the candidate the plan named (`mrai-native-commit.ts`),
 * which the server's replay keeps verifying.
 *
 * The options are the local adapter's (`MraiNativePdfAdapterOptions`) plus
 * the page the PDF belongs to; authorities travel as JSON (bytes tagged),
 * the source checkpoint is issued by the engine over the bytes it plans
 * against, and the bundled Unicode font is the engine's own.
 */

import {
  refusedPdfEngineResult,
  type PdfEngineAdapter,
  type PdfEngineAdapterContext,
  type PdfEngineCommand,
  type PdfEngineResult,
  type PdfEngineSourceIdentity,
} from "../contracts";
import type { PdfSharedStateBridge } from "../shared-state-bridge";
import type { MraiNativePdfAdapterOptions } from "./mrai-native";
import { commitMraiNativePlannedRecord } from "./mrai-native-commit";
import { encodePdfNativeWire } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";
import {
  EngineRefusedError,
  EngineUnavailableError,
  isEngineAbort,
} from "@mrnote/core/web/lib/engine-host/client.ts";
import { callPdfNativeEngine, createPdfNativeOffload,
  encodePdfNativeCommand, decodePdfNativeAnswer, type PdfNativeOffload } from "../../pdf-native-remote";

export type MraiNativeRemotePdfAdapterOptions = Omit<
  MraiNativePdfAdapterOptions,
  "loadNativeFontExtensionFont" | "compiledTextPreview" | "sourceCheckpoint"
> & {
  /** The page whose stored PDF the commands edit (the server authorizes it). */
  readonly pageId: string;
  /** Ask the engine to bind the record to the source's revision depth (the
   * local adapter's checkpoint; any truthy value, as the engine issues it). */
  readonly sourceCheckpoint?: unknown;
  /** Accepted for call-site compatibility; the engine has its own font. */
  readonly loadNativeFontExtensionFont?: unknown;
  /** Accepted for call-site compatibility; the engine plans again. */
  readonly compiledTextPreview?: unknown;
  readonly signal?: AbortSignal;
};

/** The option keys the engine reads (everything else is the editor's). */
const PLAN_OPTION_KEYS = [
  "textWindowHistoryAuthority",
  "paragraphHistoryOrigin",
  "fontExtensionHistoryAuthority",
  "allowReceiptBoundPreviousRevisions",
  "listTextHistoryAuthority",
  "textGroupMoveRestoreAuthority",
  "listTextPropertyAuthority",
  "listTextParagraphFrame",
  "textGroupMoveFrameRight",
  "embeddedImageHistoryAuthority",
  "embeddedImageReplacementHistoryAuthority",
  "embeddedImageTransformHistoryAuthority",
] as const;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function sourceMatches(left: PdfEngineSourceIdentity, right: PdfEngineSourceIdentity): boolean {
  return left.resourceId === right.resourceId && left.resourceVersionId === right.resourceVersionId &&
    left.sourceSha256 === right.sourceSha256;
}

/** An engine that could not answer: the command is refused, nothing changed. */
function unavailable(error: unknown, command: PdfEngineCommand): PdfEngineResult {
  // `callPdfNativeEngine` already gave the error its readable text.
  if (error instanceof EngineRefusedError) {
    return refusedPdfEngineResult("invalid-command", error.message, command.transactionId);
  }
  return refusedPdfEngineResult("adapter-unavailable",
    error instanceof EngineUnavailableError ? error.message : "The PDF editor could not answer. Nothing was changed.",
    command.transactionId);
}

export function createRemoteMraiNativePdfAdapter(options: MraiNativeRemotePdfAdapterOptions): PdfEngineAdapter {
  let openedSource: PdfEngineSourceIdentity | null = null;
  const planOptions = async (offloaded: PdfNativeOffload): Promise<Record<string, unknown>> => {
    const out: Record<string, unknown> = {};
    for (const key of PLAN_OPTION_KEYS) if (options[key] !== undefined) out[key] = options[key];
    if (options.sourceCheckpoint) out.sourceCheckpoint = {};
    // A large byte array of an authority (restored bytes) travels beside the JSON.
    return (await encodePdfNativeWire(out, null, undefined, offloaded.offload)) as Record<string, unknown>;
  };

  return Object.freeze({
    authority: "mrai-native" as const,

    availability: () =>
      globalThis.crypto?.subtle
        ? Object.freeze({ available: true })
        : Object.freeze({ available: false, reason: "MRAI native PDF SHA-256 support is unavailable." }),

    // The engine judges the payload's schema; anything else is refused there.
    canHandle: (command: PdfEngineCommand) => isRecord(command.payload) && typeof command.payload.schemaVersion === "string",

    open(source: PdfEngineSourceIdentity) {
      openedSource = Object.freeze({ ...source });
    },

    async execute(command: PdfEngineCommand, context: PdfEngineAdapterContext): Promise<PdfEngineResult> {
      const writableBeforeLoad = context.readWritable();
      if (!writableBeforeLoad.writable) {
        return refusedPdfEngineResult("not-writable", writableBeforeLoad.reason ?? "The PDF page is read-only.", command.transactionId);
      }
      if (!openedSource || !sourceMatches(openedSource, context.source)) {
        return refusedPdfEngineResult("adapter-rejected", "MRAI native adapter is not pinned to this PDF source.", command.transactionId);
      }
      if (!sourceMatches(command, context.source)) {
        return Object.freeze({
          status: "conflict" as const, transactionId: command.transactionId, reason: "source-mismatch" as const,
          latestResourceVersionId: context.source.resourceVersionId, latestSourceSha256: context.source.sourceSha256,
          latestSessionRevision: command.sessionRevision,
        });
      }
      let sourceBytes: Uint8Array | null = null;
      let planned;
      try {
        // The bytes are read first, as the local adapter did: a source the
        // editor no longer holds refuses before anything is asked.
        const held = Uint8Array.from(await options.loadPinnedSource(context.source));
        sourceBytes = held;
        const offloaded = createPdfNativeOffload();
        planned = await callPdfNativeEngine("plan", {
          source: { pageId: options.pageId, ...context.source },
          // A long text the command carries (a stamp's or a picture's
          // base64) travels beside the JSON, as a byte array does: the
          // route bounds the JSON at 1 MiB.
          command: await encodePdfNativeCommand(command, offloaded),
          options: await planOptions(offloaded),
        }, { signal: options.signal, workingBytes: () => held, offloaded });
      } catch (error) {
        sourceBytes?.fill(0);
        if (isEngineAbort(error)) throw error;
        if (error instanceof EngineRefusedError || error instanceof EngineUnavailableError) return unavailable(error, command);
        return refusedPdfEngineResult("adapter-rejected", error instanceof Error ? error.message : "PDF source load failed.", command.transactionId);
      }
      if (planned.status !== "planned") {
        sourceBytes?.fill(0);
        return planned.result;
      }
      let candidate: Uint8Array | null = null;
      try {
        if (planned.display && planned.candidate && options.onCandidatePrepared) {
          candidate = await decodePdfNativeAnswer<Uint8Array>(planned.candidate,
            { sha256: context.source.sourceSha256, bytes: sourceBytes as Uint8Array });
        }
      } catch {
        // A candidate that does not rebuild is not painted; the save still
        // stands or falls on the server's receipt.
        candidate = null;
      } finally {
        sourceBytes?.fill(0);
      }
      const writableBeforePersistence = context.readWritable();
      if (!writableBeforePersistence.writable) {
        candidate?.fill(0);
        return refusedPdfEngineResult("not-writable",
          writableBeforePersistence.reason ?? "The PDF page became read-only before mutation.", command.transactionId);
      }
      if (candidate && planned.candidate && options.onCandidatePrepared) {
        try {
          options.onCandidatePrepared(Object.freeze({ bytes: candidate, sha256: planned.candidate.$append.sha256 }));
        } catch {
          // A display hint never stops the write.
        }
      }
      return commitMraiNativePlannedRecord({
        command,
        context,
        record: planned.record,
        expectedCandidateSha256: planned.expectedCandidateSha256,
        durable: planned.durable,
        sharedStateBridge: options.sharedStateBridge as PdfSharedStateBridge,
        collection: options.collection,
      });
    },

    reload(source: PdfEngineSourceIdentity) {
      openedSource = Object.freeze({ ...source });
    },

    dispose() {
      openedSource = null;
    },
  });
}
