/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import {
  findPdfFeatureCapability,
  type PdfCanonicalId,
} from "../feature-capabilities";

export type PdfEngineAuthority =
  | "mrai-native"
  | "apryse-client"
  | "apryse-worker"
  | "mrai-service"
  | "browser-platform"
  | "legacy-unmigrated"
  | "blocked";

export type PdfEngineOutput =
  "view" | "current-pdf" | "new-artifact" | "external-effect";

export interface PdfEngineSourceIdentity {
  readonly resourceId: string;
  readonly resourceVersionId: string;
  readonly sourceSha256: string;
}

export interface PdfEngineCommand extends PdfEngineSourceIdentity {
  readonly featureId: PdfCanonicalId;
  readonly controlId: string;
  readonly action: string;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly output: PdfEngineOutput;
  readonly payload: JsonValue;
}

export interface PdfEngineWritableSnapshot {
  readonly writable: boolean;
  readonly reason?: string;
}

export interface PdfEngineAdapterAvailability {
  readonly available: boolean;
  readonly reason?: string;
}

export interface PdfEngineAppliedResult {
  readonly status: "applied";
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly resourceVersionId: string;
  readonly sourceSha256: string;
  readonly undoToken?: string;
  readonly warnings: readonly string[];
  readonly lossReport: readonly string[];
}

export interface PdfEngineRefusedResult {
  readonly status: "refused";
  readonly transactionId: string | null;
  readonly reason:
    | "invalid-command"
    | "session-not-open"
    | "session-disposed"
    | "missing-owner"
    | "adapter-unavailable"
    | "adapter-rejected"
    | "not-writable"
    | "unsupported-document"
    | "verification-failed";
  readonly detail: string;
}

export interface PdfEngineConflictResult {
  readonly status: "conflict";
  readonly transactionId: string;
  readonly reason:
    "source-mismatch" | "revision-mismatch" | "object-precondition-mismatch";
  readonly latestResourceVersionId: string;
  readonly latestSourceSha256: string;
  readonly latestSessionRevision: number;
}

export type PdfEngineResult =
  PdfEngineAppliedResult | PdfEngineRefusedResult | PdfEngineConflictResult;

export class PdfEngineNotWritableError extends Error {
  readonly code = "PDF_ENGINE_NOT_WRITABLE" as const;

  constructor(detail = "The PDF page is read-only.") {
    super(detail);
    this.name = "PdfEngineNotWritableError";
  }
}

export interface PdfEngineAdapterContext {
  readonly source: PdfEngineSourceIdentity;
  readonly readWritable: () => PdfEngineWritableSnapshot;
}

export interface PdfEngineAdapter {
  readonly authority: PdfEngineAuthority;
  availability(): PdfEngineAdapterAvailability;
  canHandle(command: PdfEngineCommand): boolean;
  open(
    source: PdfEngineSourceIdentity,
    context: PdfEngineAdapterContext,
  ): void | Promise<void>;
  execute(
    command: PdfEngineCommand,
    context: PdfEngineAdapterContext,
  ): PdfEngineResult | Promise<PdfEngineResult>;
  setWritable?(writable: PdfEngineWritableSnapshot): void | Promise<void>;
  reload?(
    source: PdfEngineSourceIdentity,
    context: PdfEngineAdapterContext,
  ): void | Promise<void>;
  dispose?(): void | Promise<void>;
}

const ENGINE_OUTPUTS = new Set<PdfEngineOutput>([
  "view",
  "current-pdf",
  "new-artifact",
  "external-effect",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isJsonValue(value: unknown): value is JsonValue {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "boolean"
  ) {
    return true;
  }
  if (typeof value === "number") return Number.isFinite(value);
  if (Array.isArray(value)) return value.every(isJsonValue);
  if (!isRecord(value)) return false;
  return Object.values(value).every(isJsonValue);
}

function readNonEmptyString(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function readRevision(value: unknown): number | null {
  return Number.isSafeInteger(value) && Number(value) >= 0
    ? Number(value)
    : null;
}

export function readPdfEngineCommand(value: unknown): PdfEngineCommand | null {
  if (!isRecord(value)) return null;

  const feature =
    typeof value.featureId === "string"
      ? findPdfFeatureCapability(value.featureId)
      : null;
  const controlId = readNonEmptyString(value.controlId);
  const action = readNonEmptyString(value.action);
  const transactionId = readNonEmptyString(value.transactionId);
  const resourceId = readNonEmptyString(value.resourceId);
  const resourceVersionId = readNonEmptyString(value.resourceVersionId);
  const sourceSha256 = readNonEmptyString(value.sourceSha256);
  const sessionRevision = readRevision(value.sessionRevision);
  const selectionRevision = readRevision(value.selectionRevision);
  const output = ENGINE_OUTPUTS.has(value.output as PdfEngineOutput)
    ? (value.output as PdfEngineOutput)
    : null;

  if (
    !feature ||
    !controlId ||
    !feature.entryPoints.some((entry) => entry.controlId === controlId) ||
    !action ||
    !transactionId ||
    !resourceId ||
    !resourceVersionId ||
    !sourceSha256 ||
    !/^[a-f\d]{64}$/i.test(sourceSha256) ||
    sessionRevision === null ||
    selectionRevision === null ||
    !output ||
    !isJsonValue(value.payload)
  ) {
    return null;
  }

  return Object.freeze({
    featureId: feature.id,
    controlId,
    action,
    transactionId,
    resourceId,
    resourceVersionId,
    sourceSha256: sourceSha256.toLowerCase(),
    sessionRevision,
    selectionRevision,
    output,
    payload: value.payload,
  });
}

export function refusedPdfEngineResult(
  reason: PdfEngineRefusedResult["reason"],
  detail: string,
  transactionId: string | null = null,
): PdfEngineRefusedResult {
  return Object.freeze({ status: "refused", transactionId, reason, detail });
}
