/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Dedicated Ruby WASI Worker. Authored source never runs on the page thread. */

import {
  RUBY_EXECUTION_PROTOCOL,
  RUBY_EXECUTION_RUNTIME,
  RUBY_MAX_OUTPUT_BYTES,
  RUBY_MAX_SOURCE_BYTES,
  RUBY_MAX_STDIN_BYTES,
  RUBY_WASM_SHA256,
  capRubyWasmMemory,
  createRubyExecutionRuntime,
  resolveRubyWasmAssetUrl,
} from "./ruby-execution-runtime";
import { lockDownWorkerCapabilities } from "./worker-capability-lockdown";

interface RubyWorkerScope {
  readonly location: Location;
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  postMessage(message: unknown): void;
  close(): void;
}

const workerScope = self as unknown as RubyWorkerScope;
const encoder = new TextEncoder();

// Resolve through the package's exported WASM subpath. Traversing into
// node_modules from this source file works on disk, but Turbopack treats that
// relative path as an application module and fails the entire Code route.
// The locked package export lets Next emit the immutable asset into the
// same-origin build; user input never supplies a URL.
const rubyWasmUrl = new URL(
  "@ruby/4.0-wasm-wasi/dist/ruby+stdlib.wasm",
  import.meta.url,
);

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
  protocol: typeof RUBY_EXECUTION_PROTOCOL;
  type: "run";
  id: string;
  source: string;
  stdin: string;
  outputBytes: number;
} {
  if (!isPlainRecord(value)) return false;
  const id = value.id;
  const source = value.source;
  const stdin = value.stdin;
  const outputBytes = value.outputBytes;
  return (
    value.protocol === RUBY_EXECUTION_PROTOCOL &&
    value.type === "run" &&
    typeof id === "string" &&
    id.length > 0 &&
    id.length <= 160 &&
    typeof source === "string" &&
    !source.includes("\0") &&
    !hasUnpairedSurrogate(source) &&
    encoder.encode(source).byteLength <= RUBY_MAX_SOURCE_BYTES &&
    typeof stdin === "string" &&
    !stdin.includes("\0") &&
    !hasUnpairedSurrogate(stdin) &&
    encoder.encode(stdin).byteLength <= RUBY_MAX_STDIN_BYTES &&
    Number.isSafeInteger(outputBytes) &&
    Number(outputBytes) > 0 &&
    Number(outputBytes) <= RUBY_MAX_OUTPUT_BYTES
  );
}

async function loadRuntime() {
  const resolvedRubyWasmUrl = resolveRubyWasmAssetUrl(
    rubyWasmUrl.toString(),
    workerScope.location.origin,
  );
  const response = await fetch(resolvedRubyWasmUrl, {
    cache: "force-cache",
    // Only the fixed, same-origin bootstrap may use the private-site cookie.
    // Authored Ruby cannot fetch or access it: its JS host bridge stays blocked.
    credentials: "same-origin",
    mode: "same-origin",
    redirect: "error",
  });
  if (!response.ok) throw new Error("Ruby runtime asset is unavailable.");
  const asset = await response.arrayBuffer();
  if (!globalThis.crypto?.subtle) {
    throw new Error("Ruby runtime integrity verification is unavailable.");
  }
  const digest = new Uint8Array(
    await globalThis.crypto.subtle.digest("SHA-256", asset),
  );
  const digestHex = Array.from(digest, (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
  if (digestHex !== RUBY_WASM_SHA256) {
    throw new Error("Ruby runtime asset failed integrity verification.");
  }
  const bytes = capRubyWasmMemory(asset);
  const wasmModule = await WebAssembly.compile(Uint8Array.from(bytes).buffer);
  const runtime = await createRubyExecutionRuntime(wasmModule);
  // The VM is instantiated and its WASI and VM glue are imported; every later
  // run, warm or not, only evaluates inside it (the stdlib is embedded in the
  // module). Remove the network and cross-context capabilities before the
  // first request is accepted.
  lockDownWorkerCapabilities();
  return runtime;
}

void loadRuntime()
  .then((runtime) => {
    let executing = false;
    workerScope.onmessage = (event) => {
      const request = event.data;
      if (!validRequest(request) || executing) {
        workerScope.postMessage({
          protocol: RUBY_EXECUTION_PROTOCOL,
          runtime: RUBY_EXECUTION_RUNTIME,
          type: "fatal",
          phase: "request",
        });
        workerScope.close();
        return;
      }
      executing = true;
      try {
        const receipt = runtime.execute(
          request.source,
          request.stdin,
          request.outputBytes,
        );
        workerScope.postMessage({
          protocol: RUBY_EXECUTION_PROTOCOL,
          runtime: RUBY_EXECUTION_RUNTIME,
          type: "result",
          id: request.id,
          ...receipt,
        });
        if (receipt.retireGeneration) workerScope.close();
      } catch {
        workerScope.postMessage({
          protocol: RUBY_EXECUTION_PROTOCOL,
          runtime: RUBY_EXECUTION_RUNTIME,
          type: "fatal",
          phase: "execution",
        });
        workerScope.close();
      } finally {
        executing = false;
      }
    };
    workerScope.postMessage({
      protocol: RUBY_EXECUTION_PROTOCOL,
      runtime: RUBY_EXECUTION_RUNTIME,
      type: "ready",
    });
  })
  .catch(() => {
    workerScope.postMessage({
      protocol: RUBY_EXECUTION_PROTOCOL,
      runtime: RUBY_EXECUTION_RUNTIME,
      type: "fatal",
      phase: "load",
    });
    workerScope.close();
  });
