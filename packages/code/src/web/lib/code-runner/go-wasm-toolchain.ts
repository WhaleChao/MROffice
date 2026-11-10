/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser-only Go formatter and single-file compiler diagnostics.
 *
 * The worker hosts an official, version-pinned Go toolchain compiled to WASM.
 * Authored source is parsed, formatted and type-checked; it is never executed.
 * Every request and receipt is bounded and versioned here, at the worker trust
 * boundary, before a provider can expose it to CodeMirror.
 */

import type {
  LanguageServiceProvider,
  ProviderDiagnostic,
  ProviderCapabilities,
} from "./language-service-provider";

export const GO_WASM_PROTOCOL = "mrnote.go-toolchain.v1";
export const GO_WASM_TOOLCHAIN = "go1.26.5";

const GO_WORKER_URL = "/go-toolchain-worker.js";
const READY_TIMEOUT_MS = 20_000;
const REQUEST_TIMEOUT_MS = 15_000;
const MAX_SOURCE_BYTES = 2 * 1024 * 1024;
const MAX_FORMATTED_BYTES = 4 * 1024 * 1024;
const MAX_DIAGNOSTICS = 200;
const MAX_DIAGNOSTIC_MESSAGE_CHARS = 1_000;
const MAX_DIAGNOSTIC_CODE_CHARS = 100;

const encoder = new TextEncoder();

type GoToolchainOperation = "format" | "diagnostics";

interface GoWorkerReady {
  protocol: typeof GO_WASM_PROTOCOL;
  type: "ready";
  toolchain: typeof GO_WASM_TOOLCHAIN;
}

interface ValidatedGoReceipt {
  ok: boolean;
  formatted?: string;
  diagnostics: ProviderDiagnostic[];
}

export interface GoWasmWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

interface PendingRequest {
  id: string;
  op: GoToolchainOperation;
  source: string;
  resolve: (receipt: ValidatedGoReceipt | null) => void;
  timer: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abortListener?: () => void;
}

interface WorkerHandle {
  worker: GoWasmWorkerLike;
  generation: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: Map<string, PendingRequest>;
}

const GO_CAPABILITIES: ProviderCapabilities = {
  diagnostics: true,
  formatting: true,
};

export const GO_TOOLCHAIN_UNAVAILABLE_MESSAGE =
  "Go formatter and compiler diagnostics are unavailable.";

let generation = 0;
let requestSequence = 0;
let activeHandle: WorkerHandle | null = null;

const defaultWorkerFactory = (): GoWasmWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("Go toolchain requires a browser Worker.");
  }
  return new Worker(GO_WORKER_URL, {
    name: "mrnote-go-toolchain",
  });
};

let workerFactory: () => GoWasmWorkerLike = defaultWorkerFactory;

function byteLength(value: string): number {
  return encoder.encode(value).byteLength;
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function settlePending(
  handle: WorkerHandle,
  pending: PendingRequest,
  receipt: ValidatedGoReceipt | null,
): void {
  if (handle.pending.get(pending.id) !== pending) return;
  handle.pending.delete(pending.id);
  clearTimeout(pending.timer);
  if (pending.signal && pending.abortListener) {
    pending.signal.removeEventListener("abort", pending.abortListener);
  }
  pending.resolve(receipt);
}

function retireHandle(handle: WorkerHandle): void {
  if (handle.retired) return;
  handle.retired = true;
  clearTimeout(handle.readyTimer);
  if (!handle.readySettled) {
    handle.readySettled = true;
    handle.settleReady(false);
  }
  for (const pending of [...handle.pending.values()]) {
    settlePending(handle, pending, null);
  }
  try {
    handle.worker.terminate();
  } catch {
    // A broken Worker is already retired; termination is best-effort cleanup.
  }
  if (activeHandle === handle) activeHandle = null;
}

function parseReady(value: unknown): value is GoWorkerReady {
  if (!isRecord(value)) return false;
  return (
    value.protocol === GO_WASM_PROTOCOL &&
    value.type === "ready" &&
    value.toolchain === GO_WASM_TOOLCHAIN
  );
}

function parseDiagnostic(
  value: unknown,
  sourceLength: number,
): ProviderDiagnostic | null {
  if (!isRecord(value)) return null;
  const from = value.from;
  const to = value.to;
  const message = value.message;
  const severity = value.severity;
  const code = value.code;
  if (
    !Number.isSafeInteger(from) ||
    !Number.isSafeInteger(to) ||
    (from as number) < 0 ||
    (to as number) < (from as number) ||
    (to as number) > sourceLength ||
    typeof message !== "string" ||
    message.length === 0 ||
    message.length > MAX_DIAGNOSTIC_MESSAGE_CHARS ||
    (severity !== "error" && severity !== "warning" && severity !== "info") ||
    (code !== undefined &&
      (typeof code !== "string" || code.length > MAX_DIAGNOSTIC_CODE_CHARS))
  ) {
    return null;
  }
  return {
    span: { from: from as number, to: to as number },
    message,
    severity,
    ...(typeof code === "string" && code ? { code } : {}),
  };
}

function parseReceipt(
  value: unknown,
  pending: PendingRequest,
): ValidatedGoReceipt | null {
  if (!isRecord(value)) return null;
  if (
    value.protocol !== GO_WASM_PROTOCOL ||
    value.toolchain !== GO_WASM_TOOLCHAIN ||
    value.id !== pending.id ||
    value.op !== pending.op ||
    typeof value.ok !== "boolean" ||
    !Array.isArray(value.diagnostics) ||
    value.diagnostics.length > MAX_DIAGNOSTICS
  ) {
    return null;
  }
  const diagnostics = value.diagnostics.map((item) =>
    parseDiagnostic(item, pending.source.length),
  );
  if (diagnostics.some((item) => item === null)) return null;

  if (pending.op === "diagnostics") {
    if (value.ok !== true || value.formatted !== undefined) return null;
    return {
      ok: true,
      diagnostics: diagnostics as ProviderDiagnostic[],
    };
  }

  if (value.ok === false) {
    if (value.formatted !== undefined) return null;
    return {
      ok: false,
      diagnostics: diagnostics as ProviderDiagnostic[],
    };
  }
  if (
    typeof value.formatted !== "string" ||
    byteLength(value.formatted) > MAX_FORMATTED_BYTES ||
    (pending.source.trim().length > 0 && value.formatted.trim().length === 0)
  ) {
    return null;
  }
  return {
    ok: true,
    formatted: value.formatted,
    diagnostics: diagnostics as ProviderDiagnostic[],
  };
}

function receiveWorkerMessage(handle: WorkerHandle, value: unknown): void {
  if (handle.retired || activeHandle !== handle) return;
  if (isRecord(value) && value.type === "fatal") {
    if (
      value.protocol === GO_WASM_PROTOCOL &&
      value.toolchain === GO_WASM_TOOLCHAIN
    ) {
      retireHandle(handle);
    }
    return;
  }
  if (isRecord(value) && value.type === "ready") {
    if (!parseReady(value) || handle.ready) {
      retireHandle(handle);
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

  if (!isRecord(value) || typeof value.id !== "string") return;
  const pending = handle.pending.get(value.id);
  if (!pending) return;
  const receipt = parseReceipt(value, pending);
  if (!receipt) {
    // A malformed exact-id receipt means this generation is no longer a
    // trustworthy compiler boundary. Retire siblings with it.
    retireHandle(handle);
    return;
  }
  settlePending(handle, pending, receipt);
}

function createHandle(): WorkerHandle | null {
  let worker: GoWasmWorkerLike;
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
    pending: new Map<string, PendingRequest>(),
  } satisfies WorkerHandle;
  handle.readyTimer = setTimeout(() => retireHandle(handle), READY_TIMEOUT_MS);
  worker.onmessage = (event) => receiveWorkerMessage(handle, event.data);
  worker.onerror = () => retireHandle(handle);
  activeHandle = handle;
  return handle;
}

function getHandle(): WorkerHandle | null {
  if (activeHandle && !activeHandle.retired) return activeHandle;
  return createHandle();
}

async function requestToolchain(
  op: GoToolchainOperation,
  source: string,
  signal?: AbortSignal,
): Promise<ValidatedGoReceipt | null> {
  if (
    signal?.aborted ||
    hasUnpairedSurrogate(source) ||
    byteLength(source) > MAX_SOURCE_BYTES
  ) {
    return null;
  }
  const handle = getHandle();
  if (!handle || !(await handle.readyPromise)) return null;
  if (
    signal?.aborted ||
    handle.retired ||
    activeHandle !== handle ||
    !handle.ready
  ) {
    return null;
  }

  const id = `go:${handle.generation}:${++requestSequence}`;
  return new Promise<ValidatedGoReceipt | null>((resolve) => {
    const pending: PendingRequest = {
      id,
      op,
      source,
      resolve,
      timer: 0 as unknown as ReturnType<typeof setTimeout>,
      signal,
    };
    pending.timer = setTimeout(() => {
      if (handle.pending.get(id) === pending) retireHandle(handle);
    }, REQUEST_TIMEOUT_MS);
    if (signal) {
      pending.abortListener = () => settlePending(handle, pending, null);
      signal.addEventListener("abort", pending.abortListener, { once: true });
    }
    handle.pending.set(id, pending);
    if (signal?.aborted) {
      settlePending(handle, pending, null);
      return;
    }
    try {
      handle.worker.postMessage({
        protocol: GO_WASM_PROTOCOL,
        id,
        op,
        source,
      });
    } catch {
      retireHandle(handle);
    }
  });
}

export async function formatGoWithWasm(
  source: string,
  signal?: AbortSignal,
): Promise<string | null> {
  const receipt = await requestToolchain("format", source, signal);
  return receipt?.ok && typeof receipt.formatted === "string"
    ? receipt.formatted
    : null;
}

export async function diagnoseGoWithWasm(
  source: string,
  signal?: AbortSignal,
): Promise<ProviderDiagnostic[] | null> {
  const receipt = await requestToolchain("diagnostics", source, signal);
  return receipt?.ok ? receipt.diagnostics : null;
}

export function createGoWasmLanguageProvider(): LanguageServiceProvider {
  return {
    id: "go-wasm-toolchain",
    languages: ["go"],
    capabilities: (language) =>
      language === "go" ? GO_CAPABILITIES : {},
    async diagnostics(request) {
      const diagnostics = await diagnoseGoWithWasm(
        request.text,
        request.signal,
      );
      if (diagnostics) return diagnostics;
      return [
        {
          span: { from: 0, to: 0 },
          severity: "warning",
          message: GO_TOOLCHAIN_UNAVAILABLE_MESSAGE,
          code: "go.toolchainUnavailable",
        },
      ];
    },
    formatting(request) {
      return formatGoWithWasm(request.text, request.signal);
    },
  };
}

/** Install an isolated Worker seam and retire the previous generation. Tests only. */
export function __setGoWasmWorkerFactoryForTests(
  factory: (() => GoWasmWorkerLike) | null,
): void {
  if (activeHandle) retireHandle(activeHandle);
  workerFactory = factory ?? defaultWorkerFactory;
}

/** Retire all module state so a later test cannot receive an old generation. */
export function __resetGoWasmToolchainForTests(): void {
  if (activeHandle) retireHandle(activeHandle);
  workerFactory = defaultWorkerFactory;
  requestSequence = 0;
}
