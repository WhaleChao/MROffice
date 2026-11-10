/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Dedicated, analyze-only Bashkit Worker for Shell syntax diagnostics. */

import initBashkit, {
  Bash,
} from "./vendor/bashkit-0.16/bashkit_wasm";
import {
  SHELL_EXECUTION_RUNTIME,
  SHELL_HARDENED_WASM_SHA256,
  SHELL_INTERPRETER_ACCOUNTING_CEILING_BYTES,
  SHELL_MAX_COMMANDS,
  SHELL_MAX_LOOP_ITERATIONS,
  SHELL_MAX_SOURCE_BYTES,
  SHELL_WASM_SHA256,
  hardenShellWasmMemory,
} from "./shell-execution-runtime";
import { lockDownWorkerCapabilities } from "./worker-capability-lockdown";

interface ShellLanguageWorkerScope {
  readonly location: Location;
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  postMessage(message: unknown): void;
  close(): void;
}

const workerScope = self as unknown as ShellLanguageWorkerScope;
const encoder = new TextEncoder();
const emittedShellWasmAsset = "/shell-execution.wasm";
const MAX_ERROR_CHARS = 2_000;

const SHELL_ANALYSIS_PROTOCOL = "mrnote.shell-analysis.v1";
const SHELL_ANALYSIS_RUNTIME = `${SHELL_EXECUTION_RUNTIME}:analyze-only`;

function isPlainRecord(value: unknown): value is Record<string, unknown> {
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

function validRequest(value: unknown): value is {
  protocol: typeof SHELL_ANALYSIS_PROTOCOL;
  type: "analyze";
  id: string;
  source: string;
} {
  return (
    isPlainRecord(value) &&
    value.protocol === SHELL_ANALYSIS_PROTOCOL &&
    value.type === "analyze" &&
    typeof value.id === "string" &&
    value.id.length > 0 &&
    value.id.length <= 160 &&
    typeof value.source === "string" &&
    !value.source.includes("\0") &&
    !hasUnpairedSurrogate(value.source) &&
    encoder.encode(value.source).byteLength <= SHELL_MAX_SOURCE_BYTES
  );
}

function boundedError(error: unknown): string {
  try {
    if (error instanceof Error && error.message.trim()) {
      return error.message.trim().slice(0, MAX_ERROR_CHARS);
    }
    if (typeof error === "string" && error.trim()) {
      return error.trim().slice(0, MAX_ERROR_CHARS);
    }
  } catch {
    // Fall through to the fixed file-level diagnostic.
  }
  return "Shell syntax analysis failed.";
}

async function sha256Hex(value: ArrayBuffer | Uint8Array): Promise<string> {
  const bytes: ArrayBuffer =
    value instanceof Uint8Array
      ? new Uint8Array(value).buffer
      : value;
  const digest = new Uint8Array(
    await globalThis.crypto.subtle.digest("SHA-256", bytes),
  );
  return Array.from(digest, (byte) => byte.toString(16).padStart(2, "0")).join(
    "",
  );
}

async function loadRuntime(): Promise<void> {
  const shellWasmUrl = new globalThis.URL(
    emittedShellWasmAsset,
    workerScope.location.href,
  );
  if (
    shellWasmUrl.origin !== workerScope.location.origin ||
    (shellWasmUrl.protocol !== "http:" && shellWasmUrl.protocol !== "https:")
  ) {
    throw new Error("Shell analysis asset is not same-origin.");
  }
  const response = await fetch(shellWasmUrl, {
    cache: "force-cache",
    // The gate protects this fixed same-origin asset too. Source is parsed
    // only after bootstrap and never receives a network or credential bridge.
    credentials: "same-origin",
    mode: "same-origin",
    redirect: "error",
  });
  if (!response.ok) throw new Error("Shell analysis asset is unavailable.");
  const asset = await response.arrayBuffer();
  if (!globalThis.crypto?.subtle) {
    throw new Error("Shell analysis integrity verification is unavailable.");
  }
  if ((await sha256Hex(asset)) !== SHELL_WASM_SHA256) {
    throw new Error("Shell analysis asset failed integrity verification.");
  }
  const hardenedAsset = hardenShellWasmMemory(asset);
  if ((await sha256Hex(hardenedAsset)) !== SHELL_HARDENED_WASM_SHA256) {
    throw new Error("Shell analysis memory policy failed integrity verification.");
  }
  await initBashkit({
    module_or_path: await WebAssembly.compile(hardenedAsset),
  });
  // Analysis parses inside the instantiated module; nothing is fetched again.
  lockDownWorkerCapabilities();
}

function analyze(id: string, source: string): void {
  const bash = new Bash({
    profile: "hardened",
    username: "analyzer",
    hostname: "mrnote",
    cwd: "/",
    env: { LANG: "C.UTF-8", PATH: "/bin" },
    files: {},
    maxCommands: SHELL_MAX_COMMANDS,
    maxLoopIterations: SHELL_MAX_LOOP_ITERATIONS,
    timeoutMs: 5_000,
    maxMemory: SHELL_INTERPRETER_ACCOUNTING_CEILING_BYTES,
    // No customBuiltins, host filesystem, network or process bridge. `analyze`
    // parses only; a fresh interpreter makes that non-mutation boundary exact.
  });
  try {
    const result = bash.analyze(source);
    if (!isPlainRecord(result) || typeof result.truncated !== "boolean") {
      throw new Error("Shell analyzer returned a malformed result.");
    }
    workerScope.postMessage({
      protocol: SHELL_ANALYSIS_PROTOCOL,
      runtime: SHELL_ANALYSIS_RUNTIME,
      type: "result",
      id,
      status: "ok",
    });
  } catch (error) {
    const message = boundedError(error);
    if (!/\b(?:parse|syntax) error\b/i.test(message)) {
      workerScope.postMessage({
        protocol: SHELL_ANALYSIS_PROTOCOL,
        runtime: SHELL_ANALYSIS_RUNTIME,
        type: "fatal",
        phase: "analyze",
      });
      return;
    }
    workerScope.postMessage({
      protocol: SHELL_ANALYSIS_PROTOCOL,
      runtime: SHELL_ANALYSIS_RUNTIME,
      type: "result",
      id,
      status: "syntax-error",
      message,
    });
  } finally {
    bash.free();
  }
}

void loadRuntime()
  .then(() => {
    let analyzing = false;
    workerScope.onmessage = (event) => {
      const request = event.data;
      if (!validRequest(request) || analyzing) {
        workerScope.postMessage({
          protocol: SHELL_ANALYSIS_PROTOCOL,
          runtime: SHELL_ANALYSIS_RUNTIME,
          type: "fatal",
          phase: "request",
        });
        workerScope.close();
        return;
      }
      analyzing = true;
      try {
        analyze(request.id, request.source);
      } finally {
        workerScope.close();
      }
    };
    workerScope.postMessage({
      protocol: SHELL_ANALYSIS_PROTOCOL,
      runtime: SHELL_ANALYSIS_RUNTIME,
      type: "ready",
    });
  })
  .catch(() => {
    workerScope.postMessage({
      protocol: SHELL_ANALYSIS_PROTOCOL,
      runtime: SHELL_ANALYSIS_RUNTIME,
      type: "fatal",
      phase: "load",
    });
    workerScope.close();
  });
