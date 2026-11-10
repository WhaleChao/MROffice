/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfCanonicalId } from "../feature-capabilities";
import {
  pdfCommandRequiresWritable,
  type PdfCommandOwnershipManifest,
} from "./command-ownership";
import {
  PdfEngineNotWritableError,
  readPdfEngineCommand,
  refusedPdfEngineResult,
  type PdfEngineAdapter,
  type PdfEngineAuthority,
  type PdfEngineCommand,
  type PdfEngineResult,
  type PdfEngineSourceIdentity,
  type PdfEngineWritableSnapshot,
} from "./contracts";

export interface PdfEngineSessionOptions {
  readonly ownership: PdfCommandOwnershipManifest;
  readonly adapters: readonly PdfEngineAdapter[];
  readonly readWritable: () => PdfEngineWritableSnapshot;
}

export interface PdfEngineSessionSnapshot extends PdfEngineSourceIdentity {
  readonly sessionRevision: number;
}

export interface PdfEngineSession {
  open(source: PdfEngineSourceIdentity): void;
  snapshot(): PdfEngineSessionSnapshot | null;
  execute(command: unknown): Promise<PdfEngineResult>;
  dispose(): Promise<void>;
}

function isValidSource(
  source: PdfEngineSourceIdentity,
): source is PdfEngineSourceIdentity {
  return (
    source.resourceId.trim().length > 0 &&
    source.resourceVersionId.trim().length > 0 &&
    /^[a-f\d]{64}$/i.test(source.sourceSha256)
  );
}

function sourceConflict(
  command: PdfEngineCommand,
  source: PdfEngineSessionSnapshot,
  reason: "source-mismatch" | "revision-mismatch",
): PdfEngineResult {
  return Object.freeze({
    status: "conflict",
    transactionId: command.transactionId,
    reason,
    latestResourceVersionId: source.resourceVersionId,
    latestSourceSha256: source.sourceSha256,
    latestSessionRevision: source.sessionRevision,
  });
}

function adapterMap(
  adapters: readonly PdfEngineAdapter[],
): ReadonlyMap<PdfEngineAuthority, PdfEngineAdapter> {
  const byAuthority = new Map<PdfEngineAuthority, PdfEngineAdapter>();
  for (const adapter of adapters) {
    if (byAuthority.has(adapter.authority)) {
      throw new Error(`Duplicate PDF engine adapter for ${adapter.authority}.`);
    }
    byAuthority.set(adapter.authority, adapter);
  }
  return byAuthority;
}

export function createPdfEngineSession(
  options: PdfEngineSessionOptions,
): PdfEngineSession {
  const adapters = adapterMap(options.adapters);
  const openedAdapters = new Set<PdfEngineAuthority>();
  let currentSource: PdfEngineSessionSnapshot | null = null;
  let disposed = false;

  const currentContext = () => {
    if (!currentSource) return null;
    return {
      source: {
        resourceId: currentSource.resourceId,
        resourceVersionId: currentSource.resourceVersionId,
        sourceSha256: currentSource.sourceSha256,
      },
      readWritable: options.readWritable,
    } as const;
  };

  return {
    open(source) {
      if (disposed) throw new Error("PDF engine session is disposed.");
      if (currentSource) throw new Error("PDF engine session is already open.");
      if (!isValidSource(source))
        throw new Error("Invalid PDF source identity.");
      currentSource = Object.freeze({
        resourceId: source.resourceId.trim(),
        resourceVersionId: source.resourceVersionId.trim(),
        sourceSha256: source.sourceSha256.toLowerCase(),
        sessionRevision: 0,
      });
    },

    snapshot() {
      return currentSource;
    },

    async execute(rawCommand) {
      const parsed = readPdfEngineCommand(rawCommand);
      if (!parsed) {
        return refusedPdfEngineResult(
          "invalid-command",
          "The PDF command payload is invalid.",
        );
      }
      if (disposed) {
        return refusedPdfEngineResult(
          "session-disposed",
          "The PDF engine session has been disposed.",
          parsed.transactionId,
        );
      }
      if (!currentSource) {
        return refusedPdfEngineResult(
          "session-not-open",
          "The PDF engine session has not been opened.",
          parsed.transactionId,
        );
      }

      if (
        parsed.resourceId !== currentSource.resourceId ||
        parsed.resourceVersionId !== currentSource.resourceVersionId ||
        parsed.sourceSha256 !== currentSource.sourceSha256
      ) {
        return sourceConflict(parsed, currentSource, "source-mismatch");
      }
      if (parsed.sessionRevision !== currentSource.sessionRevision) {
        return sourceConflict(parsed, currentSource, "revision-mismatch");
      }

      const owner = options.ownership.find(
        parsed.featureId as PdfCanonicalId,
        parsed.action,
      );
      if (!owner) {
        return refusedPdfEngineResult(
          "missing-owner",
          `No PDF writer owns ${parsed.featureId}:${parsed.action}.`,
          parsed.transactionId,
        );
      }
      if (owner.output !== parsed.output) {
        return refusedPdfEngineResult(
          "invalid-command",
          "The command output does not match its canonical owner.",
          parsed.transactionId,
        );
      }

      const adapter = adapters.get(owner.authority);
      if (!adapter) {
        return refusedPdfEngineResult(
          "adapter-unavailable",
          `No adapter is registered for ${owner.authority}.`,
          parsed.transactionId,
        );
      }
      const availability = adapter.availability();
      if (!availability.available) {
        return refusedPdfEngineResult(
          "adapter-unavailable",
          availability.reason ?? `${owner.authority} is unavailable.`,
          parsed.transactionId,
        );
      }
      if (!adapter.canHandle(parsed)) {
        return refusedPdfEngineResult(
          "adapter-rejected",
          `${owner.authority} cannot handle this command.`,
          parsed.transactionId,
        );
      }

      let context = currentContext();
      if (!context) {
        return refusedPdfEngineResult(
          "session-not-open",
          "The PDF engine session has not been opened.",
          parsed.transactionId,
        );
      }
      let writable = context.readWritable();
      if (pdfCommandRequiresWritable(owner) && !writable.writable) {
        return refusedPdfEngineResult(
          "not-writable",
          writable.reason ?? "The PDF page is read-only.",
          parsed.transactionId,
        );
      }

      try {
        await adapter.setWritable?.(writable);
        if (!openedAdapters.has(adapter.authority)) {
          await adapter.open(context.source, context);
          openedAdapters.add(adapter.authority);
        }

        context = currentContext();
        if (!context) {
          return refusedPdfEngineResult(
            "session-not-open",
            "The PDF engine session has not been opened.",
            parsed.transactionId,
          );
        }
        writable = context.readWritable();
        await adapter.setWritable?.(writable);
        if (pdfCommandRequiresWritable(owner) && !writable.writable) {
          return refusedPdfEngineResult(
            "not-writable",
            writable.reason ?? "The PDF page is read-only.",
            parsed.transactionId,
          );
        }

        const result = await adapter.execute(parsed, context);
        if (result.status === "applied") {
          if (
            result.transactionId !== parsed.transactionId ||
            result.sessionRevision !== currentSource.sessionRevision + 1 ||
            !result.resourceVersionId.trim() ||
            !/^[a-f\d]{64}$/i.test(result.sourceSha256)
          ) {
            return refusedPdfEngineResult(
              "adapter-rejected",
              "The adapter returned an invalid mutation receipt.",
              parsed.transactionId,
            );
          }
          currentSource = Object.freeze({
            resourceId: currentSource.resourceId,
            resourceVersionId: result.resourceVersionId,
            sourceSha256: result.sourceSha256.toLowerCase(),
            sessionRevision: result.sessionRevision,
          });
        }
        return result;
      } catch (error) {
        if (error instanceof PdfEngineNotWritableError) {
          return refusedPdfEngineResult(
            "not-writable",
            error.message,
            parsed.transactionId,
          );
        }
        return refusedPdfEngineResult(
          "adapter-rejected",
          error instanceof Error ? error.message : "The PDF adapter failed.",
          parsed.transactionId,
        );
      }
    },

    async dispose() {
      if (disposed) return;
      disposed = true;
      for (const authority of openedAdapters) {
        await adapters.get(authority)?.dispose?.();
      }
      openedAdapters.clear();
      currentSource = null;
    },
  };
}
