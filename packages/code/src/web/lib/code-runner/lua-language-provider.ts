/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Trusted, compile-only Lua diagnostics backed by the pinned Lua 5.5 Worker. */

import type {
  LanguageServiceProvider,
  ProviderDiagnostic,
} from "./language-service-provider";
import {
  LUA_EXECUTION_PROTOCOL,
  LUA_EXECUTION_RUNTIME,
} from "./lua-runner";

export const LUA_FILE_LEVEL_DIAGNOSTIC_CODE = "lua.fileLevel";
export const LUA_COMPILE_UNAVAILABLE_MESSAGE =
  "Lua syntax analysis is unavailable for this source.";
export const LUA_COMPILE_SOURCE_LIMIT_MESSAGE =
  "Lua syntax analysis was not run because this source exceeds the 2 MiB limit.";
export const LUA_FILE_LEVEL_MESSAGE_PREFIX =
  "File-level Lua syntax error (no source location available)";

const READY_TIMEOUT_MS = 20_000;
const REQUEST_TIMEOUT_MS = 5_000;
const MAX_SOURCE_BYTES = 2 * 1024 * 1024;
const MAX_ERROR_CHARS = 2_000;
const encoder = new TextEncoder();

export interface LuaLanguageWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

export interface LuaLanguageProviderOptions {
  /** Isolated test seam. Production always uses the pinned public Worker. */
  workerFactory?: () => LuaLanguageWorkerLike;
  readyTimeoutMs?: number;
  requestTimeoutMs?: number;
}

interface PendingCompile {
  id: string;
  resolve: (receipt: CompileReceipt | null) => void;
  timer: ReturnType<typeof setTimeout> | null;
  signal?: AbortSignal;
  abortListener?: () => void;
}

interface CompileHandle {
  worker: LuaLanguageWorkerLike;
  generation: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: PendingCompile | null;
}

type CompileReceipt =
  | { status: "ok" }
  | { status: "syntax-error"; error: string };

let testWorkerFactory: (() => LuaLanguageWorkerLike) | null = null;

const defaultWorkerFactory = (): LuaLanguageWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("Lua syntax analysis requires a browser Worker.");
  }
  return new Worker("/lua-execution-worker.js", {
    name: "mrnote-lua-compile",
  });
};

function isOwnRecord(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  try {
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) return false;
    return Object.values(Object.getOwnPropertyDescriptors(value)).every(
      (descriptor) => "value" in descriptor,
    );
  } catch {
    return false;
  }
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

function parseReceipt(value: unknown, pending: PendingCompile): CompileReceipt | null {
  if (
    !isOwnRecord(value) ||
    value.protocol !== LUA_EXECUTION_PROTOCOL ||
    value.runtime !== LUA_EXECUTION_RUNTIME ||
    value.type !== "compile-result" ||
    value.id !== pending.id ||
    (value.status !== "ok" && value.status !== "syntax-error")
  ) {
    return null;
  }
  if (value.status === "ok") {
    return value.error === undefined ? { status: "ok" } : null;
  }
  if (
    typeof value.error !== "string" ||
    value.error.length === 0 ||
    value.error.length > MAX_ERROR_CHARS ||
    hasUnpairedSurrogate(value.error)
  ) {
    return null;
  }
  return { status: "syntax-error", error: value.error };
}

function lineSpan(source: string, requestedLine: number): { from: number; to: number } | null {
  if (!Number.isSafeInteger(requestedLine) || requestedLine < 1) return null;
  let line = 1;
  let from = 0;
  while (line < requestedLine) {
    const newline = source.indexOf("\n", from);
    if (newline === -1) return null;
    from = newline + 1;
    line += 1;
  }
  const newline = source.indexOf("\n", from);
  const rawTo = newline === -1 ? source.length : newline;
  const to = rawTo > from && source.charCodeAt(rawTo - 1) === 13
    ? rawTo - 1
    : rawTo;
  return { from, to };
}

function fileLevelDiagnostic(
  message: string,
  severity: "error" | "warning" = "error",
): ProviderDiagnostic[] {
  return [
    {
      span: { from: 0, to: 0 },
      severity,
      message,
      code: LUA_FILE_LEVEL_DIAGNOSTIC_CODE,
    },
  ];
}

function diagnosticsForReceipt(
  source: string,
  receipt: CompileReceipt,
): ProviderDiagnostic[] {
  if (receipt.status === "ok") return [];
  const located = /(?:^|\n)main[.]lua:([1-9][0-9]*):\s*(.+)$/i.exec(
    receipt.error,
  );
  if (located) {
    const span = lineSpan(source, Number(located[1]));
    if (span) {
      return [
        {
          span,
          severity: "error",
          message: receipt.error,
          code: "lua.syntax",
        },
      ];
    }
  }
  const detail = receipt.error.replace(/^main[.]lua:[^:]*:\s*/i, "").trim();
  return fileLevelDiagnostic(
    detail
      ? `${LUA_FILE_LEVEL_MESSAGE_PREFIX}: ${detail}`
      : LUA_FILE_LEVEL_MESSAGE_PREFIX,
  );
}

export function createLuaLanguageProvider(
  options: LuaLanguageProviderOptions = {},
): LanguageServiceProvider {
  const workerFactory =
    options.workerFactory ?? testWorkerFactory ?? defaultWorkerFactory;
  const readyTimeoutMs = options.readyTimeoutMs ?? READY_TIMEOUT_MS;
  const requestTimeoutMs = options.requestTimeoutMs ?? REQUEST_TIMEOUT_MS;
  let generation = 0;
  let sequence = 0;
  let diagnosticsEpoch = 0;
  let activeHandle: CompileHandle | null = null;

  const settlePending = (
    handle: CompileHandle,
    receipt: CompileReceipt | null,
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

  const retire = (handle: CompileHandle): void => {
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
      // Detached before best-effort physical termination.
    }
    if (activeHandle === handle) activeHandle = null;
  };

  const receive = (handle: CompileHandle, value: unknown): void => {
    if (handle.retired || activeHandle !== handle) return;
    if (isOwnRecord(value) && value.type === "fatal") {
      if (
        value.protocol === LUA_EXECUTION_PROTOCOL &&
        value.runtime === LUA_EXECUTION_RUNTIME
      ) {
        retire(handle);
      }
      return;
    }
    if (isOwnRecord(value) && value.type === "ready") {
      if (
        value.protocol !== LUA_EXECUTION_PROTOCOL ||
        value.runtime !== LUA_EXECUTION_RUNTIME ||
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
    if (!pending || !isOwnRecord(value) || value.id !== pending.id) return;
    const receipt = parseReceipt(value, pending);
    if (!receipt) {
      retire(handle);
      return;
    }
    settlePending(handle, receipt);
    retire(handle);
  };

  const createHandle = (): CompileHandle | null => {
    let worker: LuaLanguageWorkerLike;
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
    } satisfies CompileHandle;
    handle.readyTimer = setTimeout(() => retire(handle), readyTimeoutMs);
    worker.onmessage = (event) => receive(handle, event.data);
    worker.onerror = () => retire(handle);
    activeHandle = handle;
    return handle;
  };

  const requestCompile = async (
    source: string,
    signal?: AbortSignal,
  ): Promise<CompileReceipt | null> => {
    if (activeHandle) retire(activeHandle);
    if (signal?.aborted) return null;
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
    const id = `lua-compile:${handle.generation}:${++sequence}`;
    return new Promise<CompileReceipt | null>((resolve) => {
      const pending: PendingCompile = {
        id,
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
          protocol: LUA_EXECUTION_PROTOCOL,
          type: "compile",
          id,
          source,
        });
      } catch {
        retire(handle);
      }
    });
  };

  return {
    id: "lua-5.5-compile",
    languages: ["lua"],
    capabilities: (language) =>
      language === "lua" ? { diagnostics: true } : {},
    async diagnostics(request) {
      const epoch = ++diagnosticsEpoch;
      if (request.signal?.aborted) return [];
      if (request.text.trim().length === 0) {
        if (activeHandle) retire(activeHandle);
        return [];
      }
      if (
        request.text.includes("\0") ||
        hasUnpairedSurrogate(request.text)
      ) {
        if (activeHandle) retire(activeHandle);
        return fileLevelDiagnostic(LUA_COMPILE_UNAVAILABLE_MESSAGE, "warning");
      }
      if (encoder.encode(request.text).byteLength > MAX_SOURCE_BYTES) {
        if (activeHandle) retire(activeHandle);
        return fileLevelDiagnostic(LUA_COMPILE_SOURCE_LIMIT_MESSAGE, "warning");
      }
      const receipt = await requestCompile(request.text, request.signal);
      if (epoch !== diagnosticsEpoch || request.signal?.aborted) return [];
      return receipt
        ? diagnosticsForReceipt(request.text, receipt)
        : fileLevelDiagnostic(LUA_COMPILE_UNAVAILABLE_MESSAGE, "warning");
    },
  };
}

export function __setLuaLanguageWorkerFactoryForTests(
  factory: (() => LuaLanguageWorkerLike) | null,
): void {
  testWorkerFactory = factory;
}

export function __resetLuaLanguageProviderForTests(): void {
  testWorkerFactory = null;
}
