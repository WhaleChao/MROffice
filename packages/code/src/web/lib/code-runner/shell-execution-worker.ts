/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Dedicated, single-run Bashkit WASM Worker. Authored shell never runs on the page thread. */

import initBashkit, {
  Bash,
} from "./vendor/bashkit-0.16/bashkit_wasm";
import {
  SHELL_EXECUTION_PROTOCOL,
  SHELL_EXECUTION_RUNTIME,
  SHELL_HARDENED_WASM_SHA256,
  SHELL_MAX_COMMANDS,
  SHELL_INTERPRETER_ACCOUNTING_CEILING_BYTES,
  SHELL_MAX_LOOP_ITERATIONS,
  SHELL_MAX_OUTPUT_BYTES,
  SHELL_MAX_SOURCE_BYTES,
  SHELL_MAX_STDIN_BYTES,
  SHELL_STDIN_PATH,
  SHELL_WASM_SHA256,
  classifyShellFailure,
  hardenShellWasmMemory,
  shellCommandWithStdin,
  type ShellExecutionStatus,
} from "./shell-execution-runtime";
import { lockDownWorkerCapabilities } from "./worker-capability-lockdown";

interface ShellWorkerScope {
  readonly location: Location;
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  postMessage(message: unknown): void;
  close(): void;
}

const workerScope = self as unknown as ShellWorkerScope;
const encoder = new TextEncoder();
const MAX_LOG_LINES = 2_000;
const MAX_LOG_LINE_CHARS = 20_000;
const MAX_ERROR_CHARS = 50_000;

// The fixed MRNote rebuild is emitted from `public/` and verified before use.
// Authored source never supplies a URL or receives the bootstrap credentials.
const emittedShellWasmAsset = "/shell-execution.wasm";

class ShellOutputQuotaError extends Error {
  constructor(message = "Shell output exceeded its trusted output budget.") {
    super(message);
    this.name = "ShellOutputQuotaError";
  }
}

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
  protocol: typeof SHELL_EXECUTION_PROTOCOL;
  type: "run";
  id: string;
  source: string;
  stdin: string;
  timeoutMs: number;
  outputBytes: number;
} {
  if (!isPlainRecord(value)) return false;
  return (
    value.protocol === SHELL_EXECUTION_PROTOCOL &&
    value.type === "run" &&
    typeof value.id === "string" &&
    value.id.length > 0 &&
    value.id.length <= 160 &&
    typeof value.source === "string" &&
    !value.source.includes("\0") &&
    !hasUnpairedSurrogate(value.source) &&
    encoder.encode(value.source).byteLength <= SHELL_MAX_SOURCE_BYTES &&
    typeof value.stdin === "string" &&
    !value.stdin.includes("\0") &&
    !hasUnpairedSurrogate(value.stdin) &&
    encoder.encode(value.stdin).byteLength <= SHELL_MAX_STDIN_BYTES &&
    Number.isSafeInteger(value.timeoutMs) &&
    Number(value.timeoutMs) > 0 &&
    Number(value.timeoutMs) <= 60_000 &&
    Number.isSafeInteger(value.outputBytes) &&
    Number(value.outputBytes) > 0 &&
    Number(value.outputBytes) <= SHELL_MAX_OUTPUT_BYTES
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
    // Fall through to the fixed diagnostic.
  }
  return "Shell execution failed.";
}

class ShellOutputCollector {
  readonly logs: string[] = [];
  readonly logStreams: Array<"stdout" | "stderr"> = [];
  private readonly carries = { stdout: "", stderr: "" };
  private readonly complete = { stdout: "", stderr: "" };
  private usedBytes = 0;

  constructor(
    private readonly id: string,
    private readonly outputBytes: number,
  ) {}

  accept(stdout: string, stderr: string): void {
    if (stdout) this.acceptStream(stdout, "stdout");
    if (stderr) this.acceptStream(stderr, "stderr");
  }

  finish(): void {
    for (const stream of ["stdout", "stderr"] as const) {
      const carry = this.carries[stream];
      // A rejected overlong final fragment is not copied into the final
      // receipt. Successful runs can only reach this branch with a bounded
      // carry because acceptStream validated it before returning.
      if (
        carry &&
        carry.length <= MAX_LOG_LINE_CHARS &&
        this.logs.length < MAX_LOG_LINES
      ) {
        this.logs.push(carry.endsWith("\r") ? carry.slice(0, -1) : carry);
        this.logStreams.push(stream);
      }
      this.carries[stream] = "";
    }
  }

  exact(stream: "stdout" | "stderr"): string {
    return this.complete[stream];
  }

  private acceptStream(text: string, stream: "stdout" | "stderr"): void {
    const bytes = encoder.encode(text).byteLength;
    if (this.usedBytes + bytes > this.outputBytes) {
      throw new ShellOutputQuotaError();
    }
    this.usedBytes += bytes;
    this.complete[stream] += text;
    const lines = (this.carries[stream] + text).split("\n");
    this.carries[stream] = lines.pop() ?? "";
    if (this.carries[stream].length > MAX_LOG_LINE_CHARS) {
      throw new ShellOutputQuotaError(
        "Shell output line exceeded its trusted output budget.",
      );
    }
    for (const line of lines) this.appendLine(line, stream);
    workerScope.postMessage({
      protocol: SHELL_EXECUTION_PROTOCOL,
      runtime: SHELL_EXECUTION_RUNTIME,
      type: "output",
      id: this.id,
      chunk: text,
      stream,
    });
  }

  private appendLine(line: string, stream: "stdout" | "stderr"): void {
    if (line.length > MAX_LOG_LINE_CHARS || this.logs.length >= MAX_LOG_LINES) {
      throw new ShellOutputQuotaError();
    }
    this.logs.push(line.endsWith("\r") ? line.slice(0, -1) : line);
    this.logStreams.push(stream);
  }
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
    throw new Error("Shell runtime asset is not same-origin.");
  }
  const response = await fetch(shellWasmUrl, {
    cache: "force-cache",
    // Authorize the fixed same-origin asset through the private-site gate;
    // the interpreter still has no network or JS bridge for authored source.
    credentials: "same-origin",
    mode: "same-origin",
    redirect: "error",
  });
  if (!response.ok) throw new Error("Shell runtime asset is unavailable.");
  const asset = await response.arrayBuffer();
  if (!globalThis.crypto?.subtle) {
    throw new Error("Shell runtime integrity verification is unavailable.");
  }
  const digest = new Uint8Array(
    await globalThis.crypto.subtle.digest("SHA-256", asset),
  );
  const digestHex = Array.from(digest, (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
  if (digestHex !== SHELL_WASM_SHA256) {
    throw new Error("Shell runtime asset failed integrity verification.");
  }
  const hardenedAsset = hardenShellWasmMemory(asset);
  const hardenedDigest = new Uint8Array(
    await globalThis.crypto.subtle.digest("SHA-256", hardenedAsset),
  );
  const hardenedDigestHex = Array.from(hardenedDigest, (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
  if (hardenedDigestHex !== SHELL_HARDENED_WASM_SHA256) {
    throw new Error("Shell runtime memory policy failed integrity verification.");
  }
  const compiledModule = await WebAssembly.compile(hardenedAsset);
  await initBashkit({ module_or_path: compiledModule });
  // Bashkit is instantiated; an interpreter runs inside it with no host
  // bridge. Remove the network and cross-context capabilities before the
  // request is accepted.
  lockDownWorkerCapabilities();
}

function postResult(
  id: string,
  status: ShellExecutionStatus,
  collector: ShellOutputCollector,
  exitCode: number | null,
  error?: string,
): void {
  collector.finish();
  workerScope.postMessage({
    protocol: SHELL_EXECUTION_PROTOCOL,
    runtime: SHELL_EXECUTION_RUNTIME,
    type: "result",
    id,
    status,
    logs: collector.logs,
    logStreams: collector.logStreams,
    exitCode,
    ...(status === "succeeded"
      ? {}
      : { error: error ?? "Shell execution failed." }),
    retireGeneration: true,
  });
}

void loadRuntime()
  .then(() => {
    let executing = false;
    workerScope.onmessage = (event) => {
      const request = event.data;
      if (!validRequest(request) || executing) {
        workerScope.postMessage({
          protocol: SHELL_EXECUTION_PROTOCOL,
          runtime: SHELL_EXECUTION_RUNTIME,
          type: "fatal",
          phase: "request",
        });
        workerScope.close();
        return;
      }
      executing = true;
      const collector = new ShellOutputCollector(request.id, request.outputBytes);
      const bash = new Bash({
        profile: "hardened",
        username: "runner",
        hostname: "mrnote",
        cwd: "/",
        env: { LANG: "C.UTF-8", PATH: "/bin" },
        files: { [SHELL_STDIN_PATH]: request.stdin },
        maxCommands: SHELL_MAX_COMMANDS,
        maxLoopIterations: SHELL_MAX_LOOP_ITERATIONS,
        timeoutMs: request.timeoutMs,
        // Physical linear memory is capped by hardenShellWasmMemory(). Keep
        // Bashkit's advisory counter above that cap: version 0.16 otherwise
        // silently refuses a large assignment yet reports exit 0, which could
        // be mistaken for a complete trusted run.
        maxMemory: SHELL_INTERPRETER_ACCOUNTING_CEILING_BYTES,
        // Intentionally no customBuiltins and no host filesystem. Authored
        // shell therefore has no JS, network, credential, or host-disk bridge.
      });
      void bash
        .executeWithOutput(
          shellCommandWithStdin(request.source),
          (stdout: string, stderr: string) => collector.accept(stdout, stderr),
        )
        .then((result) => {
          if (
            typeof result.stdout !== "string" ||
            typeof result.stderr !== "string" ||
            typeof result.exitCode !== "number" ||
            typeof result.success !== "boolean" ||
            !(result.stdoutBytes instanceof Uint8Array) ||
            !(result.stderrBytes instanceof Uint8Array) ||
            result.stdout !== collector.exact("stdout") ||
            result.stderr !== collector.exact("stderr") ||
            result.stdoutBytes.byteLength !==
              encoder.encode(result.stdout).byteLength ||
            result.stderrBytes.byteLength !==
              encoder.encode(result.stderr).byteLength
          ) {
            throw new Error("Shell runtime returned an inconsistent result.");
          }
          if (result.stdoutTruncated || result.stderrTruncated) {
            throw new ShellOutputQuotaError();
          }
          if (result.success && result.exitCode === 0) {
            postResult(request.id, "succeeded", collector, result.exitCode);
          } else {
            postResult(
              request.id,
              "runtime-error",
              collector,
              result.exitCode,
              `Shell exited with status ${result.exitCode}.`,
            );
          }
        })
        .catch((error) => {
          const detail = boundedError(error);
          postResult(
            request.id,
            classifyShellFailure(detail),
            collector,
            null,
            detail,
          );
        })
        .finally(() => workerScope.close());
    };
    workerScope.postMessage({
      protocol: SHELL_EXECUTION_PROTOCOL,
      runtime: SHELL_EXECUTION_RUNTIME,
      type: "ready",
    });
  })
  .catch(() => {
    workerScope.postMessage({
      protocol: SHELL_EXECUTION_PROTOCOL,
      runtime: SHELL_EXECUTION_RUNTIME,
      type: "fatal",
      phase: "load",
    });
    workerScope.close();
  });
