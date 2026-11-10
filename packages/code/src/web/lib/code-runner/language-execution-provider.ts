/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Selector-bounded code execution provider contract.
 *
 * A provider is an execution authority, not merely a syntax label. Registering
 * one can expose Run, so the registry accepts only explicit executable entries
 * from the Code language selector and rejects duplicate ownership. Remote
 * languages stay fail-closed until an authenticated provider registers a real
 * capability receipt.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  isLanguageExecutionErrorCode,
  type LanguageExecutionErrorCode,
} from "@mrnote/core/web/lib/code-runner/execution-error-code.ts";
import type { OnOutput } from "@mrnote/core/web/lib/code-runner/output-coalescer.ts";
import {
  MAX_EXECUTION_HTML_CHARS,
  MAX_EXECUTION_IMAGE_CHARS,
  MAX_EXECUTION_IMAGE_TOTAL_CHARS,
  MAX_EXECUTION_IMAGES,
  MAX_EXECUTION_LOG_CHARS,
  MAX_EXECUTION_LOG_LINE_CHARS,
  MAX_EXECUTION_LOG_LINES,
  MAX_EXECUTION_TEXT_CHARS,
  EXECUTION_OUTPUT_TRUNCATION_MARKER,
  boundExecutionResult,
} from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";
import {
  CODE_LANGUAGE_IDS,
  isExecutableLanguage,
} from "@mrnote/core/web/lib/code-runner/languages.ts";
import {
  isCompileOnlyLanguage,
  isLanguageExecutionArtifact,
  type LanguageExecutionArtifact,
} from "./execution-artifact";
import {
  carriesBuildCache,
  isLanguageExecutionBuildCache,
  type LanguageExecutionBuildCache,
} from "./build-cache";
import {
  carriesOutputFiles,
  isLanguageExecutionOutputFiles,
  type LanguageExecutionOutputFiles,
} from "./output-files";

import {
  MAX_LATEX_PROJECT_FILES,
  MAX_LATEX_PROJECT_TEXT_BYTES,
  compilesLatexProject,
  isLatexProjectAssetId,
  isLatexProjectPath,
  type LatexProjectAsset,
  type LatexProjectFile,
} from "./latex-project";

export type {
  LanguageExecutionArtifact,
  LanguageExecutionBuildCache,
  LatexProjectAsset,
  LatexProjectFile,
};

export const LANGUAGE_EXECUTION_SCHEMA_VERSION = 1 as const;
const SHA256_PATTERN = /^[0-9a-f]{64}$/;
const ID_PATTERN = /^[a-z0-9][a-z0-9._-]{0,127}$/;
const SELECTOR_LANGUAGE_IDS = new Set(CODE_LANGUAGE_IDS);

export type LanguageExecutionMode = "browser" | "remote";
export type LanguageExecutionAvailability =
  | {
      readonly status: "available";
      readonly providerId: string;
      readonly mode: LanguageExecutionMode;
      readonly toolchainProfile: string;
      readonly toolchainDigest: string;
      readonly sandboxProfile: string;
      readonly sandboxPolicyDigest: string;
      readonly supports: LanguageExecutionSupports;
      readonly limits: LanguageExecutionLimits;
    }
  | {
      readonly status: "checking";
      readonly providerId: string;
      readonly mode: LanguageExecutionMode;
      readonly supports: LanguageExecutionSupports;
      readonly reason: "capability-checking";
    }
  | {
      readonly status: "unavailable";
      readonly reason: "provider-unavailable";
    }
  | {
      readonly status: "non-executable";
      readonly reason: "non-executable-language";
    };

export interface LanguageExecutionSupports {
  readonly stdin: boolean;
  readonly partialExecution: boolean;
  readonly streaming: boolean;
  readonly cancel: boolean;
  readonly restart: boolean;
  readonly interactiveConsole: boolean;
  readonly debug: boolean;
}

export interface LanguageExecutionLimits {
  readonly sourceBytes: number;
  readonly stdinBytes: number;
  readonly wallTimeMs: number | null;
  readonly cpuTimeMs: number | null;
  readonly memoryBytes: number | null;
  readonly stdoutBytes: number;
  readonly stderrBytes: number;
  /** Shared aggregate across stdout and stderr, including streamed chunks. */
  readonly outputBytes: number;
}

export interface LanguageExecutionRequest {
  readonly schemaVersion: typeof LANGUAGE_EXECUTION_SCHEMA_VERSION;
  readonly requestId: string;
  readonly pageId: string;
  readonly contentVersion: number | null;
  readonly language: string;
  readonly source: string;
  readonly sourceSha256: string;
  readonly stdin: string;
  readonly phase: "compile-and-run";
  readonly providerId: string;
  readonly toolchainProfile: string;
  readonly toolchainDigest: string;
  readonly sandboxProfile: string;
  readonly sandboxPolicyDigest: string;
  readonly supports: LanguageExecutionSupports;
  readonly limits: LanguageExecutionLimits;
  /** The previous successful compile's intermediate files (LaTeX only). */
  readonly buildCache?: LanguageExecutionBuildCache;
  /**
   * Recompile from scratch: no intermediate files at all, the ones the API
   * keeps for the page included (a build-cache language only).
   */
  readonly freshBuild?: boolean;
  /** The project's other text files, by path (LaTeX only). */
  readonly projectFiles?: readonly LatexProjectFile[];
  /** The project's binary files, by attachment id (LaTeX only). */
  readonly projectAssets?: readonly LatexProjectAsset[];
}

export type LanguageExecutionOutcome =
  | "succeeded"
  | "compile-error"
  | "runtime-error"
  | "timed-out"
  | "interrupted"
  | "quota-exceeded"
  | "provider-unavailable"
  | "sandbox-failure";

export interface LanguageExecutionPhaseReceipt {
  readonly status: "succeeded" | "failed" | "skipped";
  readonly exitCode: number | null;
  readonly signal: string | null;
}

export interface LanguageExecutionReceipt {
  readonly schemaVersion: typeof LANGUAGE_EXECUTION_SCHEMA_VERSION;
  readonly requestId: string;
  readonly pageId: string;
  readonly contentVersion: number | null;
  readonly language: string;
  readonly sourceSha256: string;
  readonly providerId: string;
  readonly toolchainProfile: string;
  readonly toolchainDigest: string;
  readonly sandboxProfile: string;
  readonly sandboxPolicyDigest: string;
  readonly outcome: LanguageExecutionOutcome;
  readonly compile: LanguageExecutionPhaseReceipt;
  readonly run: LanguageExecutionPhaseReceipt;
  readonly output: RunResult;
  readonly startedAt: string;
  readonly completedAt: string;
  /**
   * The verified document of a compile-only language (a LaTeX PDF). Present
   * only on such a language's successful receipt, which skips the run phase.
   */
  readonly artifact?: LanguageExecutionArtifact;
  /**
   * The intermediate files the next compile of this page may start from; on
   * a build-cache language's successful receipt, when the executor had room.
   */
  readonly buildCache?: LanguageExecutionBuildCache;
  /**
   * The compile's log and SyncTeX map (output-files.ts); on an output-file
   * language's successful or failed compile, when the executor had room.
   */
  readonly outputFiles?: LanguageExecutionOutputFiles;
  /**
   * The time a LaTeX compile was granted (its plan's compile timeout, in ms),
   * so a timeout can say which limit it met.
   */
  readonly timeLimitMs?: number;
}

/** The longest time any run may be granted (the protocol's own bound). */
export const MAX_EXECUTION_TIME_LIMIT_MS = 120_000;

export interface LanguageExecutionContext {
  readonly signal: AbortSignal;
  readonly onOutput?: OnOutput;
}

export interface LanguageExecutionProvider {
  readonly id: string;
  readonly mode: LanguageExecutionMode;
  readonly languages: readonly string[];
  /** Notify the registry when an authenticated remote catalog changes. */
  readonly subscribe?: (listener: () => void) => () => void;
  capability(language: string): LanguageExecutionAvailability;
  execute(
    request: LanguageExecutionRequest,
    context: LanguageExecutionContext,
  ): Promise<LanguageExecutionReceipt>;
}

export interface LanguageExecutionDispatch {
  readonly source: string;
  readonly stdin?: string;
  /** Sent only for a build-cache language, and only when well formed. */
  readonly buildCache?: LanguageExecutionBuildCache | null;
  /** Recompile from scratch (a build-cache language only). */
  readonly freshBuild?: boolean;
  /** A LaTeX project's other text files (latex-project.ts). */
  readonly projectFiles?: readonly LatexProjectFile[] | null;
  /** A LaTeX project's binary files, by attachment id. */
  readonly projectAssets?: readonly LatexProjectAsset[] | null;
  readonly signal: AbortSignal;
  readonly onOutput?: OnOutput;
  readonly request: Readonly<
    Pick<
      LanguageExecutionRequest,
      | "schemaVersion"
      | "requestId"
      | "pageId"
      | "contentVersion"
      | "language"
      | "sourceSha256"
      | "phase"
      | "providerId"
      | "toolchainProfile"
      | "toolchainDigest"
      | "sandboxProfile"
      | "sandboxPolicyDigest"
      | "supports"
      | "limits"
    >
  >;
}

export class LanguageExecutionAuthorityError extends Error {
  constructor(message = "Language execution authority is unavailable.") {
    super(message);
    this.name = "LanguageExecutionAuthorityError";
  }
}

export class LanguageExecutionOutputError extends Error {
  constructor(message = "Language execution output was rejected.") {
    super(message);
    this.name = "LanguageExecutionOutputError";
  }
}

export function isLanguageExecutionAuthorityError(
  error: unknown,
): error is LanguageExecutionAuthorityError {
  try {
    return error instanceof LanguageExecutionAuthorityError;
  } catch {
    return false;
  }
}

export function isLanguageExecutionOutputError(
  error: unknown,
): error is LanguageExecutionOutputError {
  try {
    return error instanceof LanguageExecutionOutputError;
  } catch {
    return false;
  }
}

export interface LanguageExecutionRequestInput {
  readonly requestId: string;
  readonly pageId: string;
  readonly contentVersion: number | null;
  readonly language: string;
  readonly source: string;
  readonly stdin?: string;
  readonly capability: Extract<
    LanguageExecutionAvailability,
    { status: "available" }
  >;
}

function validId(value: string): boolean {
  return ID_PATTERN.test(value);
}

function validDigest(value: string): boolean {
  return SHA256_PATTERN.test(value);
}

function validLimits(limits: LanguageExecutionLimits): boolean {
  if (!limits || typeof limits !== "object") return false;
  return (
    Number.isSafeInteger(limits.sourceBytes) &&
    limits.sourceBytes > 0 &&
    Number.isSafeInteger(limits.stdinBytes) &&
    limits.stdinBytes >= 0 &&
    (limits.wallTimeMs === null ||
      (Number.isSafeInteger(limits.wallTimeMs) && limits.wallTimeMs > 0)) &&
    (limits.cpuTimeMs === null ||
      (Number.isSafeInteger(limits.cpuTimeMs) && limits.cpuTimeMs > 0)) &&
    (limits.memoryBytes === null ||
      (Number.isSafeInteger(limits.memoryBytes) && limits.memoryBytes > 0)) &&
    Number.isSafeInteger(limits.stdoutBytes) &&
    limits.stdoutBytes > 0 &&
    limits.stdoutBytes <= MAX_EXECUTION_LOG_CHARS &&
    Number.isSafeInteger(limits.stderrBytes) &&
    limits.stderrBytes > 0 &&
    limits.stderrBytes <= MAX_EXECUTION_LOG_CHARS &&
    Number.isSafeInteger(limits.outputBytes) &&
    limits.outputBytes > 0 &&
    limits.outputBytes <= MAX_EXECUTION_LOG_CHARS &&
    limits.stdoutBytes <= limits.outputBytes &&
    limits.stderrBytes <= limits.outputBytes
  );
}

function validSupports(supports: LanguageExecutionSupports): boolean {
  if (!supports || typeof supports !== "object") return false;
  return (
    typeof supports.stdin === "boolean" &&
    typeof supports.partialExecution === "boolean" &&
    typeof supports.streaming === "boolean" &&
    typeof supports.cancel === "boolean" &&
    typeof supports.restart === "boolean" &&
    typeof supports.interactiveConsole === "boolean" &&
    typeof supports.debug === "boolean"
  );
}

function limitsEqual(
  left: LanguageExecutionLimits,
  right: LanguageExecutionLimits,
): boolean {
  return (
    left.sourceBytes === right.sourceBytes &&
    left.stdinBytes === right.stdinBytes &&
    left.wallTimeMs === right.wallTimeMs &&
    left.cpuTimeMs === right.cpuTimeMs &&
    left.memoryBytes === right.memoryBytes &&
    left.stdoutBytes === right.stdoutBytes &&
    left.stderrBytes === right.stderrBytes &&
    left.outputBytes === right.outputBytes
  );
}

function supportsEqual(
  left: LanguageExecutionSupports,
  right: LanguageExecutionSupports,
): boolean {
  return (
    left.stdin === right.stdin &&
    left.partialExecution === right.partialExecution &&
    left.streaming === right.streaming &&
    left.cancel === right.cancel &&
    left.restart === right.restart &&
    left.interactiveConsole === right.interactiveConsole &&
    left.debug === right.debug
  );
}

function utf8Bytes(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

export type LanguageExecutionInputLimitKind = "source" | "stdin";

/** Identify the exact UTF-8 input budget that a request exceeds. */
export function languageExecutionInputLimitKind(
  limits: Pick<LanguageExecutionLimits, "sourceBytes" | "stdinBytes">,
  source: string,
  stdin = "",
): LanguageExecutionInputLimitKind | null {
  if (utf8Bytes(source) > limits.sourceBytes) return "source";
  return utf8Bytes(stdin) > limits.stdinBytes ? "stdin" : null;
}

function utf8Prefix(value: string, maxBytes: number): string {
  if (maxBytes <= 0) return "";
  if (utf8Bytes(value) <= maxBytes) return value;
  let used = 0;
  let end = 0;
  for (const symbol of value) {
    const next = utf8Bytes(symbol);
    if (used + next > maxBytes) break;
    used += next;
    end += symbol.length;
  }
  return value.slice(0, end);
}

function createCapabilityStreamLimiter(limits: LanguageExecutionLimits): {
  push(chunk: unknown, stream: "stdout" | "stderr"): string | null;
} {
  let chunks = 0;
  let totalBytes = 0;
  let stdoutBytes = 0;
  let stderrBytes = 0;
  let aggregateSealed = false;
  let stdoutSealed = false;
  let stderrSealed = false;
  const markerBytes = utf8Bytes(EXECUTION_OUTPUT_TRUNCATION_MARKER);

  return {
    push(chunk, stream) {
      if (typeof chunk !== "string" || chunk === "" || aggregateSealed) {
        return null;
      }
      const isStderr = stream === "stderr";
      if (isStderr ? stderrSealed : stdoutSealed) return null;
      const streamUsed = isStderr ? stderrBytes : stdoutBytes;
      const streamLimit = isStderr ? limits.stderrBytes : limits.stdoutBytes;
      const aggregateRemaining = limits.outputBytes - totalBytes;
      const streamRemaining = streamLimit - streamUsed;
      const remaining = Math.min(aggregateRemaining, streamRemaining);
      if (remaining <= 0) {
        if (aggregateRemaining <= 0) aggregateSealed = true;
        if (isStderr) stderrSealed = true;
        else stdoutSealed = true;
        return null;
      }

      const lineBounded = chunk.slice(0, MAX_EXECUTION_LOG_LINE_CHARS);
      const exceedsLine = lineBounded.length !== chunk.length;
      const exceedsBytes = utf8Bytes(lineBounded) > remaining;
      const exceedsLines = chunks >= MAX_EXECUTION_LOG_LINES - 1;
      let retained = lineBounded;
      if (exceedsLine || exceedsBytes || exceedsLines) {
        // Preserve an authored line boundary when it fits. Without this byte,
        // the UI coalescer keeps the trusted truncation marker as an
        // unterminated carry and the user cannot see that output was capped
        // until the entire run settles.
        const lineBreak =
          chunk.endsWith("\n") && remaining > markerBytes ? "\n" : "";
        const prefixBudget = Math.max(
          0,
          remaining - markerBytes - utf8Bytes(lineBreak),
        );
        const prefix = exceedsLines ? "" : utf8Prefix(lineBounded, prefixBudget);
        retained =
          prefix +
          utf8Prefix(
            EXECUTION_OUTPUT_TRUNCATION_MARKER,
            Math.max(
              0,
              remaining - utf8Bytes(prefix) - utf8Bytes(lineBreak),
            ),
          ) +
          lineBreak;
        if (aggregateRemaining <= streamRemaining || exceedsLines) {
          aggregateSealed = true;
        }
        if (isStderr) stderrSealed = true;
        else stdoutSealed = true;
      }
      if (retained === "") return null;
      const retainedBytes = utf8Bytes(retained);
      totalBytes += retainedBytes;
      if (isStderr) stderrBytes += retainedBytes;
      else stdoutBytes += retainedBytes;
      chunks += 1;
      return retained;
    },
  };
}

export async function sha256ExecutionSource(source: string): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error("SHA-256 is unavailable for code execution.");
  const bytes = new TextEncoder().encode(source);
  const digest = await subtle.digest("SHA-256", bytes.slice().buffer);
  return Array.from(new Uint8Array(digest), (value) =>
    value.toString(16).padStart(2, "0"),
  ).join("");
}

export async function createLanguageExecutionRequest(
  input: LanguageExecutionRequestInput,
): Promise<LanguageExecutionRequest> {
  if (
    !validId(input.requestId) ||
    input.pageId.trim() === "" ||
    !SELECTOR_LANGUAGE_IDS.has(input.language) ||
    !isExecutableLanguage(input.language) ||
    input.capability.status !== "available" ||
    !validId(input.capability.providerId) ||
    !validId(input.capability.toolchainProfile) ||
    !validDigest(input.capability.toolchainDigest) ||
    !validId(input.capability.sandboxProfile) ||
    !validDigest(input.capability.sandboxPolicyDigest) ||
    !validSupports(input.capability.supports) ||
    !validLimits(input.capability.limits) ||
    (!input.capability.supports.stdin && (input.stdin ?? "") !== "") ||
    languageExecutionInputLimitKind(
      input.capability.limits,
      input.source,
      input.stdin ?? "",
    ) !== null
  ) {
    throw new Error("Invalid language execution authority.");
  }
  return {
    schemaVersion: LANGUAGE_EXECUTION_SCHEMA_VERSION,
    requestId: input.requestId,
    pageId: input.pageId,
    contentVersion:
      Number.isSafeInteger(input.contentVersion) && input.contentVersion! >= 0
        ? input.contentVersion
        : null,
    language: input.language,
    source: input.source,
    sourceSha256: await sha256ExecutionSource(input.source),
    stdin: input.stdin ?? "",
    phase: "compile-and-run",
    providerId: input.capability.providerId,
    toolchainProfile: input.capability.toolchainProfile,
    toolchainDigest: input.capability.toolchainDigest,
    sandboxProfile: input.capability.sandboxProfile,
    sandboxPolicyDigest: input.capability.sandboxPolicyDigest,
    supports: { ...input.capability.supports },
    limits: { ...input.capability.limits },
  };
}

/**
 * Execute an exact request without exposing the provider's full source-bearing
 * envelope to the host component. This keeps the source identity construction
 * and receipt verification inside the registry boundary.
 */
export async function dispatchLanguageExecution(
  registry: LanguageExecutionProviderRegistry,
  dispatch: LanguageExecutionDispatch,
): Promise<LanguageExecutionReceipt> {
  if (dispatch.signal.aborted) {
    throw new LanguageExecutionAuthorityError("Execution request was cancelled.");
  }
  if (
    !validSupports(dispatch.request.supports) ||
    !validLimits(dispatch.request.limits)
  ) {
    throw new LanguageExecutionAuthorityError("Invalid execution output limit.");
  }
  const sourceSha256 = await sha256ExecutionSource(dispatch.source);
  if (dispatch.signal.aborted) {
    throw new LanguageExecutionAuthorityError("Execution request was cancelled.");
  }
  if (sourceSha256 !== dispatch.request.sourceSha256) {
    throw new LanguageExecutionAuthorityError("Execution source identity mismatch.");
  }
  if (
    (!dispatch.request.supports.stdin && (dispatch.stdin ?? "") !== "") ||
    languageExecutionInputLimitKind(
      dispatch.request.limits,
      dispatch.source,
      dispatch.stdin ?? "",
    ) !== null
  ) {
    throw new LanguageExecutionAuthorityError("Execution input limit exceeded.");
  }
  const provider = registry.authorizedProviderFor(dispatch.request);
  if (!provider) throw new LanguageExecutionAuthorityError();
  const freshBuild =
    dispatch.freshBuild === true && carriesBuildCache(dispatch.request.language);
  const buildCache =
    !freshBuild &&
    dispatch.buildCache &&
    carriesBuildCache(dispatch.request.language) &&
    isLanguageExecutionBuildCache(dispatch.buildCache, dispatch.request.language)
      ? dispatch.buildCache
      : null;
  const projectFiles = dispatch.projectFiles?.length ? dispatch.projectFiles : null;
  const projectAssets = dispatch.projectAssets?.length ? dispatch.projectAssets : null;
  if (projectFiles || projectAssets) {
    // Text files and binary files share one tree: a path names one of them.
    const paths = [
      ...(projectFiles ?? []).map((file) => file.path),
      ...(projectAssets ?? []).map((asset) => asset.path),
    ];
    if (
      !compilesLatexProject(dispatch.request.language) ||
      paths.length > MAX_LATEX_PROJECT_FILES ||
      paths.some((path) => !isLatexProjectPath(path)) ||
      new Set(paths).size !== paths.length ||
      (projectAssets ?? []).some((asset) => !isLatexProjectAssetId(asset.attachmentId)) ||
      (projectFiles ?? []).reduce((total, file) => total + utf8Bytes(file.content), 0) >
        MAX_LATEX_PROJECT_TEXT_BYTES
    ) {
      throw new LanguageExecutionAuthorityError("Execution input limit exceeded.");
    }
  }
  const request: LanguageExecutionRequest = {
    ...dispatch.request,
    source: dispatch.source,
    stdin: dispatch.stdin ?? "",
    ...(buildCache ? { buildCache } : {}),
    ...(freshBuild ? { freshBuild } : {}),
    ...(projectFiles ? { projectFiles } : {}),
    ...(projectAssets ? { projectAssets } : {}),
  };
  // Stream budgets come from the exact capability receipt and count encoded
  // bytes, not JavaScript UTF-16 code units. stdout and stderr also share the
  // aggregate authority, so alternating stream names cannot double the limit.
  const streamLimiter = createCapabilityStreamLimiter(request.limits);
  const boundedOnOutput: OnOutput | undefined =
    dispatch.request.supports.streaming && dispatch.onOutput
    ? (chunk, stream) => {
        const retained = streamLimiter.push(chunk, stream);
        if (retained !== null) dispatch.onOutput?.(retained, stream);
      }
    : undefined;
  if (dispatch.signal.aborted) {
    throw new LanguageExecutionAuthorityError("Execution request was cancelled.");
  }
  const untrustedReceipt: unknown = await provider.execute(request, {
    signal: dispatch.signal,
    onOutput: boundedOnOutput,
  });
  if (dispatch.signal.aborted) {
    throw new LanguageExecutionAuthorityError("Execution request was cancelled.");
  }
  const receipt = parseLanguageExecutionReceipt(
    untrustedReceipt,
    provider.mode === "remote",
    request.limits,
  );
  if (!receipt) {
    throw provider.mode === "remote"
      ? new LanguageExecutionOutputError()
      : new LanguageExecutionAuthorityError("Execution receipt is malformed.");
  }
  if (!parsedReceiptMatchesRequest(receipt, request)) {
    throw new LanguageExecutionAuthorityError("Execution receipt identity mismatch.");
  }
  if (!registry.authorizedProviderFor(request, provider)) {
    throw new LanguageExecutionAuthorityError();
  }
  const output = boundExecutionResult(receipt.output);
  return { ...receipt, output };
}

const RECEIPT_KEYS = [
  "schemaVersion",
  "requestId",
  "pageId",
  "contentVersion",
  "language",
  "sourceSha256",
  "providerId",
  "toolchainProfile",
  "toolchainDigest",
  "sandboxProfile",
  "sandboxPolicyDigest",
  "outcome",
  "compile",
  "run",
  "output",
  "startedAt",
  "completedAt",
] as const;
const RECEIPT_ALLOWED_KEYS = [
  ...RECEIPT_KEYS,
  "artifact",
  "buildCache",
  "outputFiles",
  "timeLimitMs",
] as const;
const PHASE_KEYS = ["status", "exitCode", "signal"] as const;
const OUTPUT_KEYS = [
  "logs",
  "logStreams",
  "result",
  "error",
  "executionErrorCode",
  "timedOut",
  "interrupted",
  "images",
  "imageMimes",
  "html",
] as const;
const RECEIPT_OUTCOMES: ReadonlySet<string> = new Set([
  "succeeded",
  "compile-error",
  "runtime-error",
  "timed-out",
  "interrupted",
  "quota-exceeded",
  "provider-unavailable",
  "sandbox-failure",
]);

function ownDataRecord(
  value: unknown,
  allowedKeys: readonly string[],
  requiredKeys: readonly string[] = [],
): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  let descriptors: PropertyDescriptorMap;
  try {
    if (
      Object.getPrototypeOf(value) !== Object.prototype &&
      Object.getPrototypeOf(value) !== null
    ) {
      return null;
    }
    descriptors = Object.getOwnPropertyDescriptors(value);
  } catch {
    return null;
  }
  const allowed = new Set(allowedKeys);
  if (Reflect.ownKeys(descriptors).some((key) => typeof key !== "string")) {
    return null;
  }
  for (const [key, descriptor] of Object.entries(descriptors)) {
    if (!allowed.has(key) || !("value" in descriptor)) return null;
  }
  if (requiredKeys.some((key) => !(key in descriptors))) return null;
  return Object.fromEntries(
    Object.entries(descriptors).map(([key, descriptor]) => [
      key,
      (descriptor as PropertyDescriptor & { value: unknown }).value,
    ]),
  );
}

function ownArray(value: unknown): unknown[] | null {
  if (!Array.isArray(value)) return null;
  let descriptors: Record<PropertyKey, PropertyDescriptor>;
  try {
    descriptors = Object.getOwnPropertyDescriptors(value) as unknown as Record<
      PropertyKey,
      PropertyDescriptor
    >;
  } catch {
    return null;
  }
  const keys = Reflect.ownKeys(descriptors);
  if (
    keys.some(
      (key) =>
        typeof key !== "string" ||
        (key !== "length" && !/^(0|[1-9][0-9]*)$/.test(key)),
    )
  ) {
    return null;
  }
  const lengthDescriptor = descriptors.length;
  if (
    !lengthDescriptor ||
    !("value" in lengthDescriptor) ||
    !Number.isSafeInteger(lengthDescriptor.value) ||
    lengthDescriptor.value < 0
  ) {
    return null;
  }
  const items: unknown[] = [];
  for (let index = 0; index < lengthDescriptor.value; index += 1) {
    const descriptor = descriptors[String(index)];
    if (!descriptor || !("value" in descriptor)) return null;
    items.push(descriptor.value);
  }
  return items;
}

function parsePhaseReceipt(
  value: unknown,
): LanguageExecutionPhaseReceipt | null {
  const raw = ownDataRecord(value, PHASE_KEYS, PHASE_KEYS);
  if (!raw) return null;
  if (
    raw.status !== "succeeded" &&
    raw.status !== "failed" &&
    raw.status !== "skipped"
  ) {
    return null;
  }
  if (
    raw.exitCode !== null &&
    (!Number.isSafeInteger(raw.exitCode) || Number(raw.exitCode) < 0)
  ) {
    return null;
  }
  if (
    raw.signal !== null &&
    (typeof raw.signal !== "string" || raw.signal.length > 128)
  ) {
    return null;
  }
  return {
    status: raw.status,
    exitCode: raw.exitCode as number | null,
    signal: raw.signal as string | null,
  };
}

function parseRemoteExecutionOutput(
  value: unknown,
  limits?: LanguageExecutionLimits,
): RunResult | null {
  const raw = ownDataRecord(value, OUTPUT_KEYS, ["logs"]);
  if (!raw) return null;
  const logs = ownArray(raw.logs);
  if (!logs || logs.length > MAX_EXECUTION_LOG_LINES) return null;
  let logBytes = 0;
  let stdoutBytes = 0;
  let stderrBytes = 0;
  for (const line of logs) {
    if (typeof line !== "string" || line.length > MAX_EXECUTION_LOG_LINE_CHARS) {
      return null;
    }
    const lineBytes = utf8Bytes(line);
    logBytes += lineBytes;
    if (
      logBytes > MAX_EXECUTION_LOG_CHARS ||
      (limits && logBytes > limits.outputBytes)
    ) {
      return null;
    }
  }
  let logStreams: Array<"stdout" | "stderr"> | undefined;
  if (raw.logStreams !== undefined) {
    const streams = ownArray(raw.logStreams);
    if (!streams || streams.length !== logs.length) return null;
    if (streams.some((stream) => stream !== "stdout" && stream !== "stderr")) {
      return null;
    }
    logStreams = streams as Array<"stdout" | "stderr">;
  }
  for (let index = 0; index < logs.length; index += 1) {
    if (logStreams?.[index] === "stderr") {
      stderrBytes += utf8Bytes(String(logs[index]));
    } else {
      stdoutBytes += utf8Bytes(String(logs[index]));
    }
  }
  if (
    limits &&
    (stdoutBytes > limits.stdoutBytes || stderrBytes > limits.stderrBytes)
  ) {
    return null;
  }
  if (
    !optionalBoundedString(raw.result, MAX_EXECUTION_TEXT_CHARS) ||
    !optionalBoundedString(raw.error, MAX_EXECUTION_TEXT_CHARS) ||
    (raw.executionErrorCode !== undefined &&
      !isLanguageExecutionErrorCode(raw.executionErrorCode)) ||
    !optionalBoundedString(raw.html, MAX_EXECUTION_HTML_CHARS) ||
    (raw.timedOut !== undefined && raw.timedOut !== true) ||
    (raw.interrupted !== undefined && raw.interrupted !== true)
  ) {
    return null;
  }
  const aggregateOutputLimit = Math.min(
    limits?.outputBytes ?? MAX_EXECUTION_LOG_CHARS,
    MAX_EXECUTION_LOG_CHARS,
  );
  let aggregateOutputBytes = logBytes;
  for (const field of [raw.result, raw.error, raw.html]) {
    if (typeof field !== "string") continue;
    aggregateOutputBytes += utf8Bytes(field);
    if (aggregateOutputBytes > aggregateOutputLimit) return null;
  }
  let images: string[] | undefined;
  let imageMimes: string[] | undefined;
  if (raw.images !== undefined) {
    const candidates = ownArray(raw.images);
    if (!candidates || candidates.length > MAX_EXECUTION_IMAGES) return null;
    images = [];
    let imageChars = 0;
    for (const image of candidates) {
      if (typeof image !== "string" || image.length > MAX_EXECUTION_IMAGE_CHARS) {
        return null;
      }
      imageChars += image.length;
      if (imageChars > MAX_EXECUTION_IMAGE_TOTAL_CHARS) return null;
      aggregateOutputBytes += utf8Bytes(image);
      if (aggregateOutputBytes > aggregateOutputLimit) return null;
      images.push(image);
    }
    if (raw.imageMimes !== undefined) {
      const mimes = ownArray(raw.imageMimes);
      if (!mimes || mimes.length !== images.length) return null;
      if (
        mimes.some(
          (mime) =>
            mime !== "image/png" &&
            mime !== "image/jpeg" &&
            mime !== "image/gif" &&
            mime !== "image/webp" &&
            mime !== "image/bmp",
        )
      ) {
        return null;
      }
      imageMimes = mimes as string[];
    }
  } else if (raw.imageMimes !== undefined) {
    return null;
  }
  return {
    logs: logs as string[],
    ...(logStreams ? { logStreams } : {}),
    ...(typeof raw.result === "string" ? { result: raw.result } : {}),
    ...(typeof raw.error === "string" ? { error: raw.error } : {}),
    ...(isLanguageExecutionErrorCode(raw.executionErrorCode)
      ? { executionErrorCode: raw.executionErrorCode }
      : {}),
    ...(raw.timedOut === true ? { timedOut: true } : {}),
    ...(raw.interrupted === true ? { interrupted: true } : {}),
    ...(images ? { images } : {}),
    ...(imageMimes ? { imageMimes } : {}),
    ...(typeof raw.html === "string" ? { html: raw.html } : {}),
  } as RunResult;
}

function parseLanguageExecutionReceipt(
  value: unknown,
  strictRemoteOutput = false,
  outputLimits?: LanguageExecutionLimits,
): LanguageExecutionReceipt | null {
  const raw = ownDataRecord(value, RECEIPT_ALLOWED_KEYS, RECEIPT_KEYS);
  if (!raw) return null;
  const hasArtifact = "artifact" in raw;
  if (
    hasArtifact &&
    (typeof raw.language !== "string" ||
      raw.outcome !== "succeeded" ||
      !isLanguageExecutionArtifact(raw.artifact, raw.language))
  ) {
    return null;
  }
  const hasBuildCache = "buildCache" in raw;
  if (
    hasBuildCache &&
    (typeof raw.language !== "string" ||
      raw.outcome !== "succeeded" ||
      !isLanguageExecutionBuildCache(raw.buildCache, raw.language))
  ) {
    return null;
  }
  const hasOutputFiles = "outputFiles" in raw;
  if (
    hasOutputFiles &&
    (typeof raw.language !== "string" ||
      !carriesOutputFiles(raw.language, raw.outcome) ||
      !isLanguageExecutionOutputFiles(raw.outputFiles, raw.language))
  ) {
    return null;
  }
  const hasTimeLimit = "timeLimitMs" in raw;
  if (hasTimeLimit && !validTimeLimit(raw.timeLimitMs)) return null;
  const compile = parsePhaseReceipt(raw.compile);
  const run = parsePhaseReceipt(raw.run);
  const output = strictRemoteOutput
    ? parseRemoteExecutionOutput(raw.output, outputLimits)
    : boundExecutionResult(raw.output);
  if (!compile || !run || !output) return null;
  if (
    raw.schemaVersion !== LANGUAGE_EXECUTION_SCHEMA_VERSION ||
    typeof raw.requestId !== "string" ||
    !validId(raw.requestId) ||
    typeof raw.pageId !== "string" ||
    raw.pageId.trim() === "" ||
    (raw.contentVersion !== null &&
      (!Number.isSafeInteger(raw.contentVersion) || Number(raw.contentVersion) < 0)) ||
    typeof raw.language !== "string" ||
    !SELECTOR_LANGUAGE_IDS.has(raw.language) ||
    typeof raw.sourceSha256 !== "string" ||
    !validDigest(raw.sourceSha256) ||
    typeof raw.providerId !== "string" ||
    !validId(raw.providerId) ||
    typeof raw.toolchainProfile !== "string" ||
    !validId(raw.toolchainProfile) ||
    typeof raw.toolchainDigest !== "string" ||
    !validDigest(raw.toolchainDigest) ||
    typeof raw.sandboxProfile !== "string" ||
    !validId(raw.sandboxProfile) ||
    typeof raw.sandboxPolicyDigest !== "string" ||
    !validDigest(raw.sandboxPolicyDigest) ||
    typeof raw.outcome !== "string" ||
    !RECEIPT_OUTCOMES.has(raw.outcome) ||
    typeof raw.startedAt !== "string" ||
    !Number.isFinite(Date.parse(raw.startedAt)) ||
    typeof raw.completedAt !== "string" ||
    !Number.isFinite(Date.parse(raw.completedAt)) ||
    Date.parse(raw.completedAt) < Date.parse(raw.startedAt)
  ) {
    return null;
  }
  if (
    !receiptOutcomeIsConsistent(
      raw.outcome as LanguageExecutionOutcome,
      compile,
      run,
      output,
      isCompileOnlyLanguage(raw.language),
      hasArtifact,
    )
  ) {
    return null;
  }
  return {
    schemaVersion: LANGUAGE_EXECUTION_SCHEMA_VERSION,
    requestId: raw.requestId,
    pageId: raw.pageId,
    contentVersion: raw.contentVersion as number | null,
    language: raw.language,
    sourceSha256: raw.sourceSha256,
    providerId: raw.providerId,
    toolchainProfile: raw.toolchainProfile,
    toolchainDigest: raw.toolchainDigest,
    sandboxProfile: raw.sandboxProfile,
    sandboxPolicyDigest: raw.sandboxPolicyDigest,
    outcome: raw.outcome as LanguageExecutionOutcome,
    compile,
    run,
    output,
    startedAt: raw.startedAt,
    completedAt: raw.completedAt,
    ...(hasArtifact
      ? { artifact: raw.artifact as LanguageExecutionArtifact }
      : {}),
    ...(hasBuildCache
      ? { buildCache: raw.buildCache as LanguageExecutionBuildCache }
      : {}),
    ...(hasOutputFiles
      ? { outputFiles: raw.outputFiles as LanguageExecutionOutputFiles }
      : {}),
    ...(hasTimeLimit ? { timeLimitMs: raw.timeLimitMs as number } : {}),
  };
}

export function validTimeLimit(value: unknown): value is number {
  return (
    Number.isSafeInteger(value) &&
    Number(value) >= 100 &&
    Number(value) <= MAX_EXECUTION_TIME_LIMIT_MS
  );
}

function receiptOutcomeIsConsistent(
  outcome: LanguageExecutionOutcome,
  compile: LanguageExecutionPhaseReceipt,
  run: LanguageExecutionPhaseReceipt,
  output: RunResult,
  compileOnly = false,
  hasArtifact = false,
): boolean {
  const hasError = typeof output.error === "string" && output.error !== "";
  const timedOut = output.timedOut === true;
  const interrupted = output.interrupted === true;
  const errorCode = output.executionErrorCode;
  const compileFailed = compile.status === "failed" && run.status === "skipped";
  const runFailed = compile.status === "succeeded" && run.status === "failed";
  const notStarted = compile.status === "skipped" && run.status === "skipped";
  // A compile-only language (LaTeX) has no program to run: its document is
  // the result, so a clean compile skips the run phase and never fails it.
  const compiledOnly =
    compileOnly && compile.status === "succeeded" && run.status === "skipped";
  // These are the only legal compile-and-run phase shapes. In particular, a
  // skipped compile cannot be followed by a run, and a failed compile cannot
  // claim that the program ran. This makes outcome a checkable receipt rather
  // than a decorative label supplied by an untrusted remote provider.
  if (
    !(
      (!compileOnly &&
        compile.status === "succeeded" &&
        run.status === "succeeded") ||
      compiledOnly ||
      compileFailed ||
      (!compileOnly && runFailed) ||
      notStarted
    )
  ) {
    return false;
  }
  if (errorCode && !executionErrorCodeMatchesOutcome(errorCode, outcome)) {
    return false;
  }
  if (
    (compile.status === "skipped" &&
      (compile.exitCode !== null || compile.signal !== null)) ||
    (run.status === "skipped" &&
      (run.exitCode !== null || run.signal !== null)) ||
    (compile.status === "succeeded" && compile.signal !== null) ||
    (run.status === "succeeded" && run.signal !== null)
  ) {
    return false;
  }
  if (outcome === "succeeded") {
    if (compileOnly) {
      // The document IS the success; a clean compile without one is not.
      return compiledOnly && hasArtifact && !hasError && !timedOut && !interrupted;
    }
    return (
      compile.status === "succeeded" &&
      run.status === "succeeded" &&
      !hasError &&
      !timedOut &&
      !interrupted
    );
  }
  if (outcome === "compile-error") {
    return compileFailed && hasError && !timedOut && !interrupted;
  }
  if (outcome === "runtime-error") {
    return runFailed && hasError && !timedOut && !interrupted;
  }
  if (outcome === "timed-out") {
    return (compileFailed || runFailed) && timedOut && !interrupted;
  }
  if (outcome === "interrupted") {
    return (compileFailed || runFailed) && interrupted && !timedOut;
  }
  if (outcome === "provider-unavailable") {
    return notStarted && hasError && !timedOut && !interrupted;
  }
  if (outcome === "quota-exceeded") {
    return (
      (compileFailed || runFailed) && hasError && !timedOut && !interrupted
    );
  }
  // A sandbox may fail before compilation starts or while either phase owns
  // the isolated process, but it may never report a successful phase pair.
  return (
    (notStarted || compileFailed || runFailed) &&
    hasError &&
    !timedOut &&
    !interrupted
  );
}

function executionErrorCodeMatchesOutcome(
  code: LanguageExecutionErrorCode,
  outcome: LanguageExecutionOutcome,
): boolean {
  if (outcome === "compile-error") return code === "compile-failed";
  if (outcome === "runtime-error") return code === "runtime-failed";
  if (outcome === "timed-out") {
    return code === "time-limit" || code === "cpu-limit";
  }
  if (outcome === "interrupted") return code === "interrupted";
  if (outcome === "quota-exceeded") {
    return code === "memory-limit" || code === "output-limit";
  }
  if (outcome === "sandbox-failure") {
    return code === "sandbox-policy" || code === "toolchain-failure";
  }
  return false;
}

function optionalBoundedString(
  value: unknown,
  maxChars: number,
): value is string | undefined {
  return value === undefined || (typeof value === "string" && value.length <= maxChars);
}

export function executionReceiptMatchesRequest(
  receipt: unknown,
  request: LanguageExecutionRequest,
): boolean {
  const parsed = parseLanguageExecutionReceipt(receipt, true, request.limits);
  if (!parsed) return false;
  return parsedReceiptMatchesRequest(parsed, request);
}

function parsedReceiptMatchesRequest(
  parsed: LanguageExecutionReceipt,
  request: LanguageExecutionRequest,
): boolean {
  return (
    parsed.requestId === request.requestId &&
    parsed.pageId === request.pageId &&
    parsed.contentVersion === request.contentVersion &&
    parsed.language === request.language &&
    parsed.sourceSha256 === request.sourceSha256 &&
    parsed.providerId === request.providerId &&
    parsed.toolchainProfile === request.toolchainProfile &&
    parsed.toolchainDigest === request.toolchainDigest &&
    parsed.sandboxProfile === request.sandboxProfile &&
    parsed.sandboxPolicyDigest === request.sandboxPolicyDigest
  );
}

export function outcomeForRunResult(result: RunResult): LanguageExecutionOutcome {
  if (result.interrupted) return "interrupted";
  if (result.timedOut) return "timed-out";
  if (result.error) return "runtime-error";
  return "succeeded";
}

export function phaseReceipt(
  status: LanguageExecutionPhaseReceipt["status"],
): LanguageExecutionPhaseReceipt {
  return { status, exitCode: null, signal: null };
}

export class LanguageExecutionProviderRegistry {
  private readonly providers = new Map<string, LanguageExecutionProvider>();
  private readonly languageOwners = new Map<string, string>();
  private readonly providerUnsubscribes = new Map<string, () => void>();
  private readonly listeners = new Set<() => void>();
  private version = 0;

  private emitChange(): void {
    this.version += 1;
    for (const listener of [...this.listeners]) listener();
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getVersion = (): number => this.version;

  register(provider: LanguageExecutionProvider): () => void {
    if (!validId(provider.id) || provider.languages.length === 0) {
      throw new Error("Invalid language execution provider.");
    }
    if (this.providers.has(provider.id)) {
      throw new Error(`Execution provider already registered: ${provider.id}`);
    }
    const languages = [...new Set(provider.languages)];
    for (const language of languages) {
      if (!SELECTOR_LANGUAGE_IDS.has(language) || !isExecutableLanguage(language)) {
        throw new Error(`Execution language is not registerable: ${language}`);
      }
      if (this.languageOwners.has(language)) {
        throw new Error(`Execution language already has a provider: ${language}`);
      }
    }
    this.providers.set(provider.id, provider);
    for (const language of languages) this.languageOwners.set(language, provider.id);
    if (provider.subscribe) {
      const unsubscribe = provider.subscribe(() => this.emitChange());
      this.providerUnsubscribes.set(provider.id, unsubscribe);
    }
    this.emitChange();
    let current = true;
    return () => {
      if (!current || this.providers.get(provider.id) !== provider) return;
      current = false;
      this.providerUnsubscribes.get(provider.id)?.();
      this.providerUnsubscribes.delete(provider.id);
      this.providers.delete(provider.id);
      for (const language of languages) {
        if (this.languageOwners.get(language) === provider.id) {
          this.languageOwners.delete(language);
        }
      }
      this.emitChange();
    };
  }

  providerFor(language: string): LanguageExecutionProvider | null {
    const providerId = this.languageOwners.get(language);
    return providerId ? (this.providers.get(providerId) ?? null) : null;
  }

  capability(language: string): LanguageExecutionAvailability {
    if (!SELECTOR_LANGUAGE_IDS.has(language) || !isExecutableLanguage(language)) {
      return { status: "non-executable", reason: "non-executable-language" };
    }
    const provider = this.providerFor(language);
    if (!provider) {
      return { status: "unavailable", reason: "provider-unavailable" };
    }
    try {
      const capability = provider.capability(language);
      if (
        capability.status === "checking" &&
        capability.providerId === provider.id &&
        capability.mode === provider.mode &&
        validSupports(capability.supports) &&
        capability.supports.stdin === false &&
        capability.supports.partialExecution === false &&
        capability.supports.streaming === false &&
        capability.supports.cancel === false &&
        capability.supports.restart === false &&
        capability.supports.interactiveConsole === false &&
        capability.supports.debug === false
      ) {
        return capability;
      }
      if (
        capability.status !== "available" ||
        capability.providerId !== provider.id ||
        capability.mode !== provider.mode ||
        !validSupports(capability.supports) ||
        !validId(capability.toolchainProfile) ||
        !validDigest(capability.toolchainDigest) ||
        !validId(capability.sandboxProfile) ||
        !validDigest(capability.sandboxPolicyDigest) ||
        !validLimits(capability.limits)
      ) {
        return { status: "unavailable", reason: "provider-unavailable" };
      }
      return capability;
    } catch {
      return { status: "unavailable", reason: "provider-unavailable" };
    }
  }

  authorizedProviderFor(
    request: Pick<
      LanguageExecutionRequest,
      | "language"
      | "providerId"
      | "toolchainProfile"
      | "toolchainDigest"
      | "sandboxProfile"
      | "sandboxPolicyDigest"
      | "supports"
      | "limits"
    >,
    expected?: LanguageExecutionProvider,
  ): LanguageExecutionProvider | null {
    const provider = this.providerFor(request.language);
    if (!provider || (expected && provider !== expected)) return null;
    const capability = this.capability(request.language);
    return capability.status === "available" &&
      capability.providerId === request.providerId &&
      capability.toolchainProfile === request.toolchainProfile &&
      capability.toolchainDigest === request.toolchainDigest &&
      capability.sandboxProfile === request.sandboxProfile &&
      capability.sandboxPolicyDigest === request.sandboxPolicyDigest &&
      supportsEqual(capability.supports, request.supports) &&
      limitsEqual(capability.limits, request.limits)
      ? provider
      : null;
  }
}
