/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Parse-only Shell diagnostics backed by the pinned Bashkit WebAssembly parser.
 *
 * This boundary is intentionally separate from Shell execution. Each request
 * owns a fresh Worker generation, the Worker exposes only `Bash.analyze`, and a
 * newer buffer snapshot physically retires the older generation. Authored text
 * therefore has no execute, host-filesystem, custom-builtin or network bridge.
 */

import type {
  LanguageServiceProvider,
  ProviderCapabilities,
  ProviderDiagnostic,
} from "./language-service-provider";
import {
  SHELL_EXECUTION_RUNTIME,
  SHELL_MAX_SOURCE_BYTES,
} from "./shell-execution-runtime";

export const SHELL_ANALYSIS_PROTOCOL = "mrnote.shell-analysis.v1";
export const SHELL_ANALYSIS_RUNTIME = `${SHELL_EXECUTION_RUNTIME}:analyze-only`;
export const SHELL_FILE_LEVEL_DIAGNOSTIC_CODE = "shell.fileLevel";
/**
 * Internal transport marker for the existing CodeMirror lint adapter, which
 * intentionally carries only span/severity/message. The editor strips it
 * before presenting or reporting the Problem, so product copy stays exact.
 */
export const SHELL_FILE_LEVEL_DIAGNOSTIC_MARKER =
  "\u2063shell.fileLevel\u2063";
export const SHELL_FILE_LEVEL_MESSAGE_PREFIX =
  "File-level Shell syntax error (no source location available)";
export const SHELL_ANALYSIS_UNAVAILABLE_MESSAGE =
  "Shell syntax analysis is unavailable for this source.";
export const SHELL_ANALYSIS_SOURCE_LIMIT_MESSAGE =
  "Shell syntax analysis was not run because this source exceeds the 2 MiB limit.";

const READY_TIMEOUT_MS = 20_000;
const REQUEST_TIMEOUT_MS = 5_000;
const MAX_ERROR_CHARS = 2_000;
const encoder = new TextEncoder();

const SHELL_DIAGNOSTIC_CAPABILITIES: ProviderCapabilities = {
  diagnostics: true,
};

export interface ShellLanguageWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

export interface ShellLanguageProviderOptions {
  /** Isolated test seam. Production always uses the module Worker below. */
  workerFactory?: () => ShellLanguageWorkerLike;
  /** Bounded test seams; production callers must not override these. */
  readyTimeoutMs?: number;
  requestTimeoutMs?: number;
}

interface PendingAnalysis {
  id: string;
  source: string;
  resolve: (receipt: AnalysisReceipt | null) => void;
  timer: ReturnType<typeof setTimeout> | null;
  signal?: AbortSignal;
  abortListener?: () => void;
}

interface AnalysisHandle {
  worker: ShellLanguageWorkerLike;
  generation: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: PendingAnalysis | null;
}

type AnalysisReceipt =
  | { status: "ok" }
  | { status: "syntax-error"; message: string };

let testWorkerFactory: (() => ShellLanguageWorkerLike) | null = null;

const defaultWorkerFactory = (): ShellLanguageWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("Shell analysis requires a browser Worker.");
  }
  return new Worker(new URL("./shell-language-worker.ts", import.meta.url), {
    name: "mrnote-shell-analysis",
    type: "module",
  });
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasUnpairedSurrogate(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return true;
      index += 1;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      return true;
    }
  }
  return false;
}

function parseReceipt(
  value: unknown,
  pending: PendingAnalysis,
): AnalysisReceipt | null {
  if (
    !isRecord(value) ||
    value.protocol !== SHELL_ANALYSIS_PROTOCOL ||
    value.runtime !== SHELL_ANALYSIS_RUNTIME ||
    value.type !== "result" ||
    value.id !== pending.id ||
    (value.status !== "ok" && value.status !== "syntax-error")
  ) {
    return null;
  }
  if (value.status === "ok") {
    return value.message === undefined ? { status: "ok" } : null;
  }
  if (
    typeof value.message !== "string" ||
    value.message.length === 0 ||
    value.message.length > MAX_ERROR_CHARS ||
    !/\b(?:parse|syntax) error\b/i.test(value.message)
  ) {
    return null;
  }
  return { status: "syntax-error", message: value.message };
}

/** Convert Bashkit's 1-based Unicode-scalar column into CodeMirror UTF-16. */
function sourceSpanAt(
  source: string,
  requestedLine: number,
  requestedColumn: number,
): { from: number; to: number } | null {
  if (
    !Number.isSafeInteger(requestedLine) ||
    !Number.isSafeInteger(requestedColumn) ||
    requestedLine < 1 ||
    requestedColumn < 1
  ) {
    return null;
  }
  let line = 1;
  let start = 0;
  while (line < requestedLine) {
    const newline = source.indexOf("\n", start);
    if (newline === -1) return null;
    start = newline + 1;
    line += 1;
  }
  const newline = source.indexOf("\n", start);
  const rawEnd = newline === -1 ? source.length : newline;
  const end = rawEnd > start && source.charCodeAt(rawEnd - 1) === 13
    ? rawEnd - 1
    : rawEnd;
  let column = 1;
  let from = start;
  while (column < requestedColumn && from < end) {
    const point = source.codePointAt(from);
    if (point === undefined) return null;
    from += point > 0xffff ? 2 : 1;
    column += 1;
  }
  if (column !== requestedColumn) return null;
  const point = from < end ? source.codePointAt(from) : undefined;
  return {
    from,
    to: point === undefined ? from : from + (point > 0xffff ? 2 : 1),
  };
}

function unavailableDiagnostic(message: string): ProviderDiagnostic[] {
  return [
    {
      span: { from: 0, to: 0 },
      severity: "warning",
      message,
      code: SHELL_FILE_LEVEL_DIAGNOSTIC_CODE,
    },
  ];
}

export function shellDiagnosticFromAnalysis(
  source: string,
  receipt: AnalysisReceipt,
): ProviderDiagnostic[] {
  if (receipt.status === "ok") return [];
  const positioned = receipt.message.match(
    /^parse error at line ([0-9]+), column ([0-9]+):\s*(.+)$/i,
  );
  if (positioned) {
    const span = sourceSpanAt(
      source,
      Number(positioned[1]),
      Number(positioned[2]),
    );
    if (span) {
      return [
        {
          span,
          severity: "error",
          message: receipt.message,
          code: "shell.syntax",
        },
      ];
    }
  }
  const detail = receipt.message
    .replace(/^parse error:\s*/i, "")
    .trim();
  return [
    {
      span: { from: 0, to: 0 },
      severity: "error",
      message: detail
        ? `${SHELL_FILE_LEVEL_MESSAGE_PREFIX}: ${detail}`
        : SHELL_FILE_LEVEL_MESSAGE_PREFIX,
      code: SHELL_FILE_LEVEL_DIAGNOSTIC_CODE,
    },
  ];
}

export function createShellLanguageProvider(
  options: ShellLanguageProviderOptions = {},
): LanguageServiceProvider {
  const workerFactory =
    options.workerFactory ?? testWorkerFactory ?? defaultWorkerFactory;
  const readyTimeoutMs = options.readyTimeoutMs ?? READY_TIMEOUT_MS;
  const requestTimeoutMs = options.requestTimeoutMs ?? REQUEST_TIMEOUT_MS;
  let generation = 0;
  let sequence = 0;
  let diagnosticsEpoch = 0;
  let activeHandle: AnalysisHandle | null = null;

  const settlePending = (
    handle: AnalysisHandle,
    receipt: AnalysisReceipt | null,
  ): void => {
    const pending = handle.pending;
    if (!pending) return;
    handle.pending = null;
    if (pending.timer) clearTimeout(pending.timer);
    if (pending.signal && pending.abortListener) {
      pending.signal.removeEventListener("abort", pending.abortListener);
    }
    pending.resolve(receipt);
  };

  const retire = (handle: AnalysisHandle): void => {
    if (handle.retired) return;
    handle.retired = true;
    clearTimeout(handle.readyTimer);
    if (!handle.readySettled) {
      handle.readySettled = true;
      handle.settleReady(false);
    }
    settlePending(handle, null);
    try {
      handle.worker.terminate();
    } catch {
      // A detached analysis generation is already inert.
    }
    if (activeHandle === handle) activeHandle = null;
  };

  const receive = (handle: AnalysisHandle, value: unknown): void => {
    if (handle.retired || activeHandle !== handle) return;
    if (isRecord(value) && value.type === "fatal") {
      if (
        value.protocol === SHELL_ANALYSIS_PROTOCOL &&
        value.runtime === SHELL_ANALYSIS_RUNTIME
      ) {
        retire(handle);
      }
      return;
    }
    if (isRecord(value) && value.type === "ready") {
      if (
        value.protocol !== SHELL_ANALYSIS_PROTOCOL ||
        value.runtime !== SHELL_ANALYSIS_RUNTIME ||
        handle.ready
      ) {
        retire(handle);
        return;
      }
      handle.ready = true;
      clearTimeout(handle.readyTimer);
      if (!handle.readySettled) {
        handle.readySettled = true;
        handle.settleReady(true);
      }
      return;
    }
    const pending = handle.pending;
    if (!pending || !isRecord(value) || value.id !== pending.id) return;
    const receipt = parseReceipt(value, pending);
    if (!receipt) {
      retire(handle);
      return;
    }
    settlePending(handle, receipt);
    retire(handle);
  };

  const createHandle = (): AnalysisHandle | null => {
    let worker: ShellLanguageWorkerLike;
    try {
      worker = workerFactory();
    } catch {
      return null;
    }
    let settleReady!: (ready: boolean) => void;
    const readyPromise = new Promise<boolean>((resolve) => {
      settleReady = resolve;
    });
    const handle = {
      worker,
      generation: ++generation,
      ready: false,
      retired: false,
      readySettled: false,
      readyPromise,
      settleReady,
      readyTimer: 0 as unknown as ReturnType<typeof setTimeout>,
      pending: null,
    } satisfies AnalysisHandle;
    handle.readyTimer = setTimeout(() => retire(handle), readyTimeoutMs);
    worker.onmessage = (event) => receive(handle, event.data);
    worker.onerror = () => retire(handle);
    activeHandle = handle;
    return handle;
  };

  const requestAnalysis = async (
    source: string,
    signal?: AbortSignal,
  ): Promise<AnalysisReceipt | null> => {
    // New source owns a new generation. Retiring first makes every late receipt
    // from a superseded edit an empty result rather than a stale red marker.
    if (activeHandle) retire(activeHandle);
    if (
      signal?.aborted ||
      hasUnpairedSurrogate(source) ||
      encoder.encode(source).byteLength > SHELL_MAX_SOURCE_BYTES
    ) {
      return null;
    }
    const handle = createHandle();
    if (!handle || !(await handle.readyPromise)) return null;
    if (
      signal?.aborted ||
      handle.retired ||
      activeHandle !== handle ||
      !handle.ready
    ) {
      return null;
    }
    const id = `shell-analysis:${handle.generation}:${++sequence}`;
    return new Promise<AnalysisReceipt | null>((resolve) => {
      const pending: PendingAnalysis = {
        id,
        source,
        resolve,
        timer: null,
        signal,
      };
      pending.timer = setTimeout(() => retire(handle), requestTimeoutMs);
      if (signal) {
        pending.abortListener = () => retire(handle);
        signal.addEventListener("abort", pending.abortListener, { once: true });
      }
      handle.pending = pending;
      if (signal?.aborted) {
        retire(handle);
        return;
      }
      try {
        handle.worker.postMessage({
          protocol: SHELL_ANALYSIS_PROTOCOL,
          type: "analyze",
          id,
          source,
        });
      } catch {
        retire(handle);
      }
    });
  };

  return {
    id: "shell-bashkit-analyze",
    languages: ["bash"],
    capabilities: (language) =>
      language === "bash" ? SHELL_DIAGNOSTIC_CAPABILITIES : {},
    async diagnostics(request) {
      const epoch = ++diagnosticsEpoch;
      if (request.signal?.aborted) return [];
      if (request.text.trim().length === 0) {
        if (activeHandle) retire(activeHandle);
        return [];
      }
      if (encoder.encode(request.text).byteLength > SHELL_MAX_SOURCE_BYTES) {
        if (activeHandle) retire(activeHandle);
        return unavailableDiagnostic(SHELL_ANALYSIS_SOURCE_LIMIT_MESSAGE);
      }
      const receipt = await requestAnalysis(request.text, request.signal);
      if (epoch !== diagnosticsEpoch || request.signal?.aborted) return [];
      return receipt
        ? shellDiagnosticFromAnalysis(request.text, receipt)
        : unavailableDiagnostic(SHELL_ANALYSIS_UNAVAILABLE_MESSAGE);
    },
  };
}

/** Override future provider instances without mutating production registration. */
export function __setShellLanguageWorkerFactoryForTests(
  factory: (() => ShellLanguageWorkerLike) | null,
): void {
  testWorkerFactory = factory;
}

export function __resetShellLanguageProviderForTests(): void {
  testWorkerFactory = null;
}
