/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  ApiRequestError,
  apiGet,
  isApiRequestError,
  localizedApiErrorCodeMessage,
} from "@mrnote/core/web/lib/api.ts";
import { apiStream } from "@mrnote/core/web/lib/api-stream.ts";
import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  isLanguageExecutionErrorCode,
  type LanguageExecutionErrorCode,
} from "@mrnote/core/web/lib/code-runner/execution-error-code.ts";
import { HELD_EXECUTION_LANGUAGE_IDS } from "@mrnote/core/web/lib/code-runner/languages.ts";
import { decodeExecutionArtifact } from "./execution-artifact";
import {
  carriesOutputFiles,
  outputFilesFromWire,
} from "./output-files";
import {
  buildCacheFromWire,
  buildCacheToWire,
  carriesBuildCache,
} from "./build-cache";
import {
  executionReceiptMatchesRequest,
  type LanguageExecutionAvailability,
  type LanguageExecutionContext,
  type LanguageExecutionLimits,
  type LanguageExecutionProvider,
  type LanguageExecutionReceipt,
  type LanguageExecutionRequest,
  type LanguageExecutionSupports,
  validTimeLimit,
} from "./language-execution-provider";

const PROVIDER_ID = "mrai.remote.gvisor.v1";
const REMOTE_LANGUAGES = [
  "c",
  "cpp",
  "go",
  "java",
  "latex",
  "perl",
  "php",
  "r",
  "rust",
  "swift",
] as const;
type RemoteLanguage = (typeof REMOTE_LANGUAGES)[number];
const capabilityPath = (language: RemoteLanguage): string =>
  `/api/v1/code-execution/capabilities/${language}`;
const EXECUTE_PATH = "/api/v1/code-execution/execute";
const CAPABILITY_REFRESH_MS = 10_000;
const CAPABILITY_REQUEST_TIMEOUT_MS = 5_000;
const CAPABILITY_INTEREST_TTL_MS = CAPABILITY_REFRESH_MS * 3;
const SHA256_PATTERN = /^[0-9a-f]{64}$/u;
const ID_PATTERN = /^[a-z0-9][a-z0-9._-]{0,127}$/u;
const TIMESTAMP_PATTERN =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,6})?Z$/u;
const OUTCOMES = new Set([
  "succeeded",
  "compile-error",
  "runtime-error",
  "timed-out",
  "interrupted",
  "quota-exceeded",
  "sandbox-failure",
]);

class RemoteExecutionReceiptError extends Error {
  constructor() {
    super("Remote execution receipt was rejected.");
    this.name = "RemoteExecutionReceiptError";
  }
}
const SUPPORTS: LanguageExecutionSupports = Object.freeze({
  stdin: true,
  partialExecution: false,
  streaming: false,
  cancel: true,
  restart: true,
  interactiveConsole: false,
  debug: false,
});

type Available = Extract<
  LanguageExecutionAvailability,
  { status: "available" }
>;
type ProviderState =
  | { readonly status: "checking" }
  | { readonly status: "unavailable" }
  | { readonly status: "available"; readonly capability: Available };

export interface RemoteExecutionProvider extends LanguageExecutionProvider {
  dispose(): void;
}

/**
 * Where a run is sent instead of the code page's own route: an Agent's
 * pending code draft, run from its review (the draft's route checks the run
 * is exactly that draft's source under its own page id).
 */
export interface RemoteExecutionTarget {
  readonly path: string;
  /** The draft's workspace, when it is not the current one. */
  readonly headers?: Readonly<Record<string, string>>;
}

export interface RemoteExecutionProviderOptions {
  /**
   * A signed-out window can never satisfy the capability endpoint, which
   * requires an authenticated user and a workspace. Asking it every ten
   * seconds only produces a 401 the provider already knows the answer to, so
   * a guest provider publishes `unavailable` once and schedules nothing.
   */
  readonly guest?: boolean;
  /** Runs go here instead of the code page's route (capabilities do not). */
  readonly target?: RemoteExecutionTarget | null;
}

function isRemoteLanguage(language: string): language is RemoteLanguage {
  if (HELD_EXECUTION_LANGUAGE_IDS.has(language)) return false;
  return (
    language === "c" ||
    language === "cpp" ||
    language === "go" ||
    language === "java" ||
    language === "latex" ||
    language === "perl" ||
    language === "php" ||
    language === "r" ||
    language === "rust" ||
    language === "swift"
  );
}

export function createRemoteExecutionProvider(
  options: RemoteExecutionProviderOptions = {},
): RemoteExecutionProvider {
  const guest = options.guest === true;
  const executePath = options.target?.path ?? EXECUTE_PATH;
  const executeHeaders = options.target?.headers;
  // A guest starts at the answer it would reach anyway, so no window ever
  // reports "checking" for a request it is not allowed to make.
  const initialState: ProviderState = guest
    ? { status: "unavailable" }
    : { status: "checking" };
  const states: Record<RemoteLanguage, ProviderState> = {
    c: initialState,
    cpp: initialState,
    go: initialState,
    java: initialState,
    latex: initialState,
    perl: initialState,
    php: initialState,
    r: initialState,
    rust: initialState,
    swift: initialState,
  };
  let disposed = false;
  const refreshStarted: Record<RemoteLanguage, boolean> = {
    c: false,
    cpp: false,
    go: false,
    java: false,
    latex: false,
    perl: false,
    php: false,
    r: false,
    rust: false,
    swift: false,
  };
  const refreshing: Record<RemoteLanguage, boolean> = {
    c: false,
    cpp: false,
    go: false,
    java: false,
    latex: false,
    perl: false,
    php: false,
    r: false,
    rust: false,
    swift: false,
  };
  const interested: Record<RemoteLanguage, boolean> = {
    c: false,
    cpp: false,
    go: false,
    java: false,
    latex: false,
    perl: false,
    php: false,
    r: false,
    rust: false,
    swift: false,
  };
  const lastInterestedAt: Record<RemoteLanguage, number> = {
    c: 0,
    cpp: 0,
    go: 0,
    java: 0,
    latex: 0,
    perl: 0,
    php: 0,
    r: 0,
    rust: 0,
    swift: 0,
  };
  const refreshTimers: Record<
    RemoteLanguage,
    ReturnType<typeof setTimeout> | null
  > = {
    c: null,
    cpp: null,
    go: null,
    java: null,
    latex: null,
    perl: null,
    php: null,
    r: null,
    rust: null,
    swift: null,
  };
  const refreshControllers: Record<
    RemoteLanguage,
    AbortController | null
  > = {
    c: null,
    cpp: null,
    go: null,
    java: null,
    latex: null,
    perl: null,
    php: null,
    r: null,
    rust: null,
    swift: null,
  };
  const lastCheckedAt: Record<RemoteLanguage, number> = {
    c: 0,
    cpp: 0,
    go: 0,
    java: 0,
    latex: 0,
    perl: 0,
    php: 0,
    r: 0,
    rust: 0,
    swift: 0,
  };
  const refreshGeneration: Record<RemoteLanguage, number> = {
    c: 0,
    cpp: 0,
    go: 0,
    java: 0,
    latex: 0,
    perl: 0,
    php: 0,
    r: 0,
    rust: 0,
    swift: 0,
  };
  const listeners = new Set<() => void>();

  const clearRefresh = (language: RemoteLanguage): void => {
    const timer = refreshTimers[language];
    if (timer !== null) clearTimeout(timer);
    refreshTimers[language] = null;
  };

  const recentlyInterested = (language: RemoteLanguage): boolean =>
    interested[language] &&
    Date.now() - lastInterestedAt[language] < CAPABILITY_INTEREST_TTL_MS;

  const scheduleRefresh = (language: RemoteLanguage): void => {
    if (
      guest ||
      disposed ||
      listeners.size === 0 ||
      !recentlyInterested(language) ||
      refreshing[language] ||
      refreshTimers[language] !== null
    ) {
      return;
    }
    const elapsed = Date.now() - lastCheckedAt[language];
    refreshTimers[language] = setTimeout(
      () => {
        refreshTimers[language] = null;
        if (!recentlyInterested(language)) {
          interested[language] = false;
          return;
        }
        void refresh(language);
      },
      Math.max(0, CAPABILITY_REFRESH_MS - elapsed),
    );
  };

  const publish = (language: RemoteLanguage, next: ProviderState): void => {
    if (disposed) return;
    states[language] = next;
    clearRefresh(language);
    scheduleRefresh(language);
    for (const listener of listeners) listener();
  };

  const refresh = async (language: RemoteLanguage): Promise<void> => {
    if (
      guest ||
      disposed ||
      refreshing[language] ||
      !recentlyInterested(language)
    ) {
      return;
    }
    refreshing[language] = true;
    refreshStarted[language] = true;
    const generation = ++refreshGeneration[language];
    const controller = new AbortController();
    refreshControllers[language] = controller;
    let requestTimeout: ReturnType<typeof setTimeout> | null = null;
    try {
      const request = apiGet<unknown>(capabilityPath(language), {
        cache: "no-store",
        signal: controller.signal,
      });
      const timeout = new Promise<never>((_resolve, reject) => {
        requestTimeout = setTimeout(() => {
          controller.abort();
          reject(new Error("Remote execution capability request timed out."));
        }, CAPABILITY_REQUEST_TIMEOUT_MS);
      });
      const raw = await Promise.race([request, timeout]);
      const capability = parseCapability(raw, language);
      if (disposed || generation !== refreshGeneration[language]) return;
      lastCheckedAt[language] = Date.now();
      publish(
        language,
        capability
          ? { status: "available", capability }
          : { status: "unavailable" },
      );
    } catch {
      if (disposed || generation !== refreshGeneration[language]) return;
      lastCheckedAt[language] = Date.now();
      publish(language, { status: "unavailable" });
    } finally {
      if (requestTimeout !== null) clearTimeout(requestTimeout);
      if (refreshControllers[language] === controller) {
        refreshControllers[language] = null;
      }
      refreshing[language] = false;
      scheduleRefresh(language);
    }
  };

  return {
    id: PROVIDER_ID,
    mode: "remote",
    languages: REMOTE_LANGUAGES.filter(
      (language) => !HELD_EXECUTION_LANGUAGE_IDS.has(language),
    ),
    subscribe(listener) {
      listeners.add(listener);
      for (const language of REMOTE_LANGUAGES) {
        if (!interested[language]) continue;
        if (!refreshStarted[language]) void refresh(language);
        scheduleRefresh(language);
      }
      return () => {
        listeners.delete(listener);
        if (listeners.size === 0) {
          for (const language of REMOTE_LANGUAGES) clearRefresh(language);
        }
      };
    },
    capability(language) {
      if (guest || !isRemoteLanguage(language)) {
        return { status: "unavailable", reason: "provider-unavailable" };
      }
      interested[language] = true;
      lastInterestedAt[language] = Date.now();
      if (!refreshStarted[language]) void refresh(language);
      else scheduleRefresh(language);
      const state = states[language];
      if (state.status === "checking") {
        return {
          status: "checking",
          providerId: PROVIDER_ID,
          mode: "remote",
          supports: SUPPORTS,
          reason: "capability-checking",
        };
      }
      if (state.status === "available") return state.capability;
      scheduleRefresh(language);
      return { status: "unavailable", reason: "provider-unavailable" };
    },
    async execute(
      request: LanguageExecutionRequest,
      context: LanguageExecutionContext,
    ): Promise<LanguageExecutionReceipt> {
      if (
        !isRemoteLanguage(request.language) ||
        request.providerId !== PROVIDER_ID ||
        request.contentVersion === null ||
        states[request.language].status !== "available"
      ) {
        throw new Error("Remote execution authority is unavailable.");
      }
      const capability = states[request.language];
      if (
        capability.status !== "available" ||
        !sameAuthority(request, capability.capability)
      ) {
        throw new Error("Remote execution authority is unavailable.");
      }
      try {
        const raw = await executeStreamed(
          {
            schema_version: request.schemaVersion,
            request_id: request.requestId,
            page_id: request.pageId,
            content_version: request.contentVersion,
            language: request.language,
            source: request.source,
            source_sha256: request.sourceSha256,
            stdin: request.stdin,
            phase: request.phase,
            provider_id: request.providerId,
            toolchain_profile: request.toolchainProfile,
            toolchain_digest: request.toolchainDigest,
            sandbox_profile: request.sandboxProfile,
            sandbox_policy_digest: request.sandboxPolicyDigest,
            limits: snakeLimits(request.limits),
            ...(request.buildCache && carriesBuildCache(request.language)
              ? { build_cache: buildCacheToWire(request.buildCache) }
              : {}),
            ...(request.freshBuild && carriesBuildCache(request.language)
              ? { fresh_build: true }
              : {}),
            ...(request.projectFiles?.length && request.language === "latex"
              ? {
                  project_files: request.projectFiles.map((file) => ({
                    path: file.path,
                    content: file.content,
                  })),
                }
              : {}),
            ...(request.projectAssets?.length && request.language === "latex"
              ? {
                  project_assets: request.projectAssets.map((asset) => ({
                    path: asset.path,
                    attachment_id: asset.attachmentId,
                  })),
                }
              : {}),
          },
          context.signal,
          executePath,
          executeHeaders,
        );
        const receipt = await parseReceipt(raw, request);
        if (receipt.outcome === "sandbox-failure") {
          lastCheckedAt[request.language] = Date.now();
          publish(request.language, { status: "unavailable" });
        }
        return receipt;
      } catch (error) {
        if (
          error instanceof RemoteExecutionReceiptError ||
          (isApiRequestError(error) &&
            (error.status === 0 ||
              error.status === 502 ||
              error.status === 503 ||
              error.code === "execution_authority_conflict"))
        ) {
          lastCheckedAt[request.language] = Date.now();
          publish(request.language, { status: "unavailable" });
        }
        throw error;
      }
    },
    dispose() {
      disposed = true;
      for (const language of REMOTE_LANGUAGES) refreshGeneration[language] += 1;
      for (const language of REMOTE_LANGUAGES) {
        refreshControllers[language]?.abort();
        refreshControllers[language] = null;
      }
      for (const language of REMOTE_LANGUAGES) clearRefresh(language);
      listeners.clear();
    },
  };
}

function parseCapability(
  raw: unknown,
  language: RemoteLanguage,
): Available | null {
  if (!isRecord(raw)) return null;
  if (raw.status === "unavailable") {
    return exactKeys(raw, ["reason", "status"]) &&
      raw.reason === "provider-unavailable"
      ? null
      : null;
  }
  if (
    raw.status !== "available" ||
    !exactKeys(raw, [
      "language",
      "limits",
      "mode",
      "provider_id",
      "sandbox_policy_digest",
      "sandbox_profile",
      "status",
      "supports",
      "toolchain_digest",
      "toolchain_profile",
    ]) ||
    raw.language !== language ||
    raw.provider_id !== PROVIDER_ID ||
    raw.mode !== "remote" ||
    !validId(raw.toolchain_profile) ||
    !validDigest(raw.toolchain_digest) ||
    !validId(raw.sandbox_profile) ||
    !validDigest(raw.sandbox_policy_digest) ||
    !sameSupports(raw.supports)
  ) {
    return null;
  }
  const limits = parseLimits(raw.limits);
  if (!limits) return null;
  return {
    status: "available",
    providerId: PROVIDER_ID,
    mode: "remote",
    toolchainProfile: raw.toolchain_profile,
    toolchainDigest: raw.toolchain_digest,
    sandboxProfile: raw.sandbox_profile,
    sandboxPolicyDigest: raw.sandbox_policy_digest,
    supports: SUPPORTS,
    limits,
  };
}

const RECEIPT_KEYS = [
  "compile",
  "completed_at",
  "content_version",
  "language",
  "outcome",
  "output",
  "page_id",
  "provider_id",
  "request_id",
  "run",
  "sandbox_policy_digest",
  "sandbox_profile",
  "schema_version",
  "source_sha256",
  "started_at",
  "toolchain_digest",
  "toolchain_profile",
] as const;

/**
 * One run, asked for as server-sent events: the route answers at once, says
 * it is still running every few seconds, then sends the receipt, or the error
 * its JSON form would have answered. A long compile (100 s on the largest
 * plan) would otherwise meet the edge proxy's limit on a silent response.
 */
async function executeStreamed(
  body: unknown,
  signal: AbortSignal,
  path: string,
  headers?: Readonly<Record<string, string>>,
): Promise<unknown> {
  let events: AsyncGenerator<{ event: string; data: Record<string, unknown> }>;
  try {
    events = apiStream(path, body, signal, headers ? { ...headers } : undefined);
    for await (const event of events) {
      if (event.event === "receipt") return event.data;
      if (event.event === "error") throw streamedError(event.data);
    }
  } catch (error) {
    if (isApiRequestError(error) || signal.aborted) throw error;
    // The connection failed or dropped: the run's own outcome is unknown.
    throw new ApiRequestError(
      error instanceof Error ? error.message : "Remote execution failed.",
      { status: 0, code: "network_unreachable" },
    );
  }
  throw new ApiRequestError("Remote execution ended without a receipt.", {
    status: 0,
    code: "network_unreachable",
  });
}

function streamedError(data: Record<string, unknown>): ApiRequestError {
  const error = isRecord(data.error) ? data.error : {};
  const status = Number.isSafeInteger(data.status) ? Number(data.status) : 502;
  const details = isRecord(error.details) ? error.details : undefined;
  // A refusal inside the stream reads in the person's language, as one on a
  // plain response does; the server's English text is only the fallback.
  return new ApiRequestError(
    localizedApiErrorCodeMessage(error.code) ??
      (typeof error.message === "string" ? error.message : "Remote execution failed."),
    {
      status,
      ...(typeof error.code === "string" ? { code: error.code } : {}),
      ...(details ? { details } : {}),
    },
  );
}

async function parseReceipt(
  raw: unknown,
  request: LanguageExecutionRequest,
): Promise<LanguageExecutionReceipt> {
  // Only a successful LaTeX receipt carries its compiled document, and it
  // always does; every other receipt keeps exactly the keys it always had.
  const carriesArtifact =
    request.language === "latex" &&
    isRecord(raw) &&
    raw.outcome === "succeeded";
  // Its intermediate files may come with it: absent when the receipt had no
  // room for them, so the next compile simply starts from nothing.
  const carriesCache =
    carriesArtifact && carriesBuildCache(request.language) && "build_cache" in raw;
  // A finished compile (successful or not) may return its log and SyncTeX
  // map; absent when the receipt had no room for them.
  const carriesReports =
    isRecord(raw) &&
    carriesOutputFiles(request.language, raw.outcome) &&
    "output_files" in raw;
  // The time a LaTeX compile was granted: its plan's compile timeout.
  const carriesTimeLimit =
    request.language === "latex" && isRecord(raw) && "time_limit_ms" in raw;
  if (
    !isRecord(raw) ||
    !exactKeys(raw, [
      ...RECEIPT_KEYS,
      ...(carriesArtifact ? ["artifact"] : []),
      ...(carriesCache ? ["build_cache"] : []),
      ...(carriesReports ? ["output_files"] : []),
      ...(carriesTimeLimit ? ["time_limit_ms"] : []),
    ]) ||
    (carriesTimeLimit && !validTimeLimit(raw.time_limit_ms)) ||
    raw.schema_version !== 1 ||
    raw.request_id !== request.requestId ||
    raw.page_id !== request.pageId ||
    raw.content_version !== request.contentVersion ||
    raw.language !== request.language ||
    raw.source_sha256 !== request.sourceSha256 ||
    raw.provider_id !== PROVIDER_ID ||
    raw.toolchain_profile !== request.toolchainProfile ||
    raw.toolchain_digest !== request.toolchainDigest ||
    raw.sandbox_profile !== request.sandboxProfile ||
    raw.sandbox_policy_digest !== request.sandboxPolicyDigest ||
    typeof raw.outcome !== "string" ||
    !OUTCOMES.has(raw.outcome) ||
    !validTimestamp(raw.started_at) ||
    !validTimestamp(raw.completed_at) ||
    Date.parse(raw.completed_at) < Date.parse(raw.started_at)
  ) {
    throw new RemoteExecutionReceiptError();
  }
  const compile = parsePhase(raw.compile);
  const run = parsePhase(raw.run);
  const output = parseOutput(raw.output);
  if (!compile || !run || !output) {
    throw new RemoteExecutionReceiptError();
  }
  // The bytes are decoded and their length, PDF signature and SHA-256 checked
  // here, before the registry or the window ever sees them.
  const artifact = carriesArtifact
    ? await decodeExecutionArtifact(raw.artifact, request.language)
    : null;
  if (carriesArtifact && !artifact) {
    throw new RemoteExecutionReceiptError();
  }
  const buildCache = carriesCache
    ? buildCacheFromWire(raw.build_cache, request.language)
    : null;
  if (carriesCache && !buildCache) {
    throw new RemoteExecutionReceiptError();
  }
  const outputFiles = carriesReports
    ? outputFilesFromWire(raw.output_files, request.language)
    : null;
  if (carriesReports && !outputFiles) {
    throw new RemoteExecutionReceiptError();
  }
  const receipt: LanguageExecutionReceipt = {
    schemaVersion: 1,
    requestId: request.requestId,
    pageId: request.pageId,
    contentVersion: request.contentVersion,
    language: request.language,
    sourceSha256: request.sourceSha256,
    providerId: PROVIDER_ID,
    toolchainProfile: request.toolchainProfile,
    toolchainDigest: request.toolchainDigest,
    sandboxProfile: request.sandboxProfile,
    sandboxPolicyDigest: request.sandboxPolicyDigest,
    outcome: raw.outcome as LanguageExecutionReceipt["outcome"],
    compile,
    run,
    output,
    startedAt: raw.started_at as string,
    completedAt: raw.completed_at as string,
    ...(artifact ? { artifact } : {}),
    ...(buildCache ? { buildCache } : {}),
    ...(outputFiles ? { outputFiles } : {}),
    ...(carriesTimeLimit ? { timeLimitMs: raw.time_limit_ms as number } : {}),
  };
  if (!executionReceiptMatchesRequest(receipt, request)) {
    throw new RemoteExecutionReceiptError();
  }
  return receipt;
}

function parsePhase(raw: unknown): LanguageExecutionReceipt["compile"] | null {
  if (
    !isRecord(raw) ||
    !exactKeys(raw, ["exit_code", "signal", "status"]) ||
    !["succeeded", "failed", "skipped"].includes(String(raw.status)) ||
    (raw.exit_code !== null &&
      (!Number.isSafeInteger(raw.exit_code) ||
        Number(raw.exit_code) < 0 ||
        Number(raw.exit_code) > 255)) ||
    (raw.signal !== null && typeof raw.signal !== "string")
  ) {
    return null;
  }
  return {
    status: raw.status as "succeeded" | "failed" | "skipped",
    exitCode: raw.exit_code as number | null,
    signal: raw.signal as string | null,
  };
}

function parseOutput(raw: unknown): RunResult | null {
  if (!isRecord(raw)) return null;
  const baseKeys = [
    "error",
    "interrupted",
    "log_streams",
    "logs",
    "result",
    "timed_out",
  ] as const;
  const hasErrorCode = Object.prototype.hasOwnProperty.call(raw, "error_code");
  if (
    !exactKeys(raw, hasErrorCode ? [...baseKeys, "error_code"] : baseKeys) ||
    !Array.isArray(raw.logs) ||
    !raw.logs.every((line) => typeof line === "string") ||
    !Array.isArray(raw.log_streams) ||
    raw.log_streams.length !== raw.logs.length ||
    !raw.log_streams.every(
      (stream) => stream === "stdout" || stream === "stderr",
    ) ||
    (raw.result !== null && typeof raw.result !== "string") ||
    (raw.error !== null && typeof raw.error !== "string") ||
    typeof raw.timed_out !== "boolean" ||
    typeof raw.interrupted !== "boolean" ||
    (hasErrorCode &&
      raw.error_code !== null &&
      !isLanguageExecutionErrorCode(raw.error_code))
  ) {
    return null;
  }
  const executionErrorCode = isLanguageExecutionErrorCode(raw.error_code)
    ? (raw.error_code as LanguageExecutionErrorCode)
    : undefined;
  return {
    logs: raw.logs as string[],
    logStreams: raw.log_streams as ("stdout" | "stderr")[],
    ...(typeof raw.result === "string" ? { result: raw.result } : {}),
    ...(typeof raw.error === "string" ? { error: raw.error } : {}),
    ...(executionErrorCode ? { executionErrorCode } : {}),
    ...(raw.timed_out ? { timedOut: true } : {}),
    ...(raw.interrupted ? { interrupted: true } : {}),
  };
}

function parseLimits(raw: unknown): LanguageExecutionLimits | null {
  if (
    !isRecord(raw) ||
    !exactKeys(raw, [
      "cpu_time_ms",
      "memory_bytes",
      "output_bytes",
      "source_bytes",
      "stderr_bytes",
      "stdin_bytes",
      "stdout_bytes",
      "wall_time_ms",
    ])
  ) {
    return null;
  }
  const values = [
    raw.source_bytes,
    raw.stdin_bytes,
    raw.wall_time_ms,
    raw.cpu_time_ms,
    raw.memory_bytes,
    raw.stdout_bytes,
    raw.stderr_bytes,
    raw.output_bytes,
  ];
  if (
    !values.every((value) => Number.isSafeInteger(value) && Number(value) >= 0)
  ) {
    return null;
  }
  if (
    Number(raw.source_bytes) < 1 ||
    Number(raw.wall_time_ms) < 1 ||
    Number(raw.cpu_time_ms) < 1 ||
    Number(raw.memory_bytes) < 1 ||
    Number(raw.stdout_bytes) < 1 ||
    Number(raw.stderr_bytes) < 1 ||
    Number(raw.output_bytes) < 1 ||
    Number(raw.cpu_time_ms) > Number(raw.wall_time_ms) ||
    Number(raw.stdout_bytes) > Number(raw.output_bytes) ||
    Number(raw.stderr_bytes) > Number(raw.output_bytes)
  ) {
    return null;
  }
  return {
    sourceBytes: Number(raw.source_bytes),
    stdinBytes: Number(raw.stdin_bytes),
    wallTimeMs: Number(raw.wall_time_ms),
    cpuTimeMs: Number(raw.cpu_time_ms),
    memoryBytes: Number(raw.memory_bytes),
    stdoutBytes: Number(raw.stdout_bytes),
    stderrBytes: Number(raw.stderr_bytes),
    outputBytes: Number(raw.output_bytes),
  };
}

function snakeLimits(limits: LanguageExecutionLimits): Record<string, number> {
  if (
    limits.wallTimeMs === null ||
    limits.cpuTimeMs === null ||
    limits.memoryBytes === null
  ) {
    throw new Error("Remote execution limits are incomplete.");
  }
  return {
    source_bytes: limits.sourceBytes,
    stdin_bytes: limits.stdinBytes,
    wall_time_ms: limits.wallTimeMs,
    cpu_time_ms: limits.cpuTimeMs,
    memory_bytes: limits.memoryBytes,
    stdout_bytes: limits.stdoutBytes,
    stderr_bytes: limits.stderrBytes,
    output_bytes: limits.outputBytes,
  };
}

function sameAuthority(
  request: LanguageExecutionRequest,
  capability: Available,
): boolean {
  return (
    request.toolchainProfile === capability.toolchainProfile &&
    request.toolchainDigest === capability.toolchainDigest &&
    request.sandboxProfile === capability.sandboxProfile &&
    request.sandboxPolicyDigest === capability.sandboxPolicyDigest &&
    request.supports.stdin === capability.supports.stdin &&
    request.supports.partialExecution ===
      capability.supports.partialExecution &&
    request.supports.streaming === capability.supports.streaming &&
    request.supports.cancel === capability.supports.cancel &&
    request.supports.restart === capability.supports.restart &&
    request.supports.interactiveConsole ===
      capability.supports.interactiveConsole &&
    request.supports.debug === capability.supports.debug &&
    request.limits.sourceBytes === capability.limits.sourceBytes &&
    request.limits.stdinBytes === capability.limits.stdinBytes &&
    request.limits.wallTimeMs === capability.limits.wallTimeMs &&
    request.limits.cpuTimeMs === capability.limits.cpuTimeMs &&
    request.limits.memoryBytes === capability.limits.memoryBytes &&
    request.limits.stdoutBytes === capability.limits.stdoutBytes &&
    request.limits.stderrBytes === capability.limits.stderrBytes &&
    request.limits.outputBytes === capability.limits.outputBytes
  );
}

function sameSupports(raw: unknown): boolean {
  return (
    isRecord(raw) &&
    exactKeys(raw, [
      "cancel",
      "debug",
      "interactive_console",
      "partial_execution",
      "restart",
      "stdin",
      "streaming",
    ]) &&
    raw.stdin === true &&
    raw.partial_execution === false &&
    raw.streaming === false &&
    raw.cancel === true &&
    raw.restart === true &&
    raw.interactive_console === false &&
    raw.debug === false
  );
}

function validTimestamp(value: unknown): value is string {
  return (
    typeof value === "string" &&
    TIMESTAMP_PATTERN.test(value) &&
    Number.isFinite(Date.parse(value))
  );
}

function validId(value: unknown): value is string {
  return typeof value === "string" && ID_PATTERN.test(value);
}

function validDigest(value: unknown): value is string {
  return typeof value === "string" && SHA256_PATTERN.test(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function exactKeys(
  value: Record<string, unknown>,
  keys: readonly string[],
): boolean {
  const actual = Object.keys(value).sort();
  const expected = [...keys].sort();
  return (
    actual.length === expected.length &&
    actual.every((key, index) => key === expected[index])
  );
}

export { PROVIDER_ID as REMOTE_EXECUTION_PROVIDER_ID };







const REMOTE_EXECUTION_REFUSAL_CODES = new Set([
  "ai_allowance_exhausted",
  // A draft's review: the draft was decided or replaced meanwhile.
  "agent_code_draft_changed",
  "agent_code_draft_unavailable",
  "code_execution_busy",
  "code_execution_slots_busy",
  "code_execution_quota",
  "code_execution_daily_budget",
  "code_execution_unavailable",
]);

export function isRemoteExecutionRefusal(error: unknown): error is ApiRequestError {
  return (
    isApiRequestError(error) &&
    typeof error.code === "string" &&
    REMOTE_EXECUTION_REFUSAL_CODES.has(error.code)
  );
}
