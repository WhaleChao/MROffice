/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Shared constants for the isolated Bashkit WebAssembly execution boundary. */

export const SHELL_EXECUTION_PROTOCOL = "mrnote.shell-execution.v2";
export const SHELL_EXECUTION_RUNTIME =
  "bashkit-wasm-0.16.0-trusted-counts-hard-memory-v3";
export const SHELL_WASM_SHA256 =
  "ab9f7bda78c55e2b5afe2c1295244b951a79deef13fa07b93139579d0af52aed";
export const SHELL_HARDENED_WASM_SHA256 =
  "72a44b8b85213a2ec6a4789559a990c8fc4bcbde21afa7a456a99f8b2964e919";
export const SHELL_UPSTREAM_COMMIT =
  "667acb8a0beae2a0b01b1d6e207c7399b78e84a9";

export const SHELL_MAX_SOURCE_BYTES = 2 * 1024 * 1024;
export const SHELL_MAX_STDIN_BYTES = 100 * 1024;
export const SHELL_MAX_OUTPUT_BYTES = 300_000;
export const SHELL_MAX_MEMORY_BYTES = 64 * 1024 * 1024;
export const SHELL_WASM_PAGE_BYTES = 64 * 1024;
export const SHELL_MAX_MEMORY_PAGES =
  SHELL_MAX_MEMORY_BYTES / SHELL_WASM_PAGE_BYTES;
/**
 * The pinned MRNote rebuild removes Bashkit's lower variable/array/function
 * count caps because upstream v0.16 silently drops values at those limits.
 * This byte ceiling remains above the physical WebAssembly maximum so memory
 * exhaustion traps the generation instead of signing a partial success.
 */
export const SHELL_INTERPRETER_ACCOUNTING_CEILING_BYTES = 2_000_000_000;
export const SHELL_MAX_COMMANDS = 20_000;
export const SHELL_MAX_LOOP_ITERATIONS = 1_000_000;
export const SHELL_STDIN_PATH = "/.mrnote-stdin";
export const SHELL_TOOLCHAIN_PROFILE =
  "bashkit-0.16.0-wasm-trusted-counts-hard-memory-v3";
export const SHELL_EXECUTION_POLICY_PROFILE =
  "origin-bashkit-wasm-worker-trusted-counts-hard-memory-no-host-bridges-v3";

/** Canonical, hash-bound policy represented in provider capability receipts. */
export const SHELL_EXECUTION_POLICY_MANIFEST = Object.freeze({
  schema: "mrnote.shell-policy.v3",
  runtime: SHELL_EXECUTION_RUNTIME,
  upstreamCommit: SHELL_UPSTREAM_COMMIT,
  rawWasmSha256: SHELL_WASM_SHA256,
  hardenedWasmSha256: SHELL_HARDENED_WASM_SHA256,
  countLimits: "unlimited-behind-physical-wasm-maximum",
  workerIsolation: "dedicated-single-run-module-worker",
  sameOriginAsset: true,
  fetchCredentials: "omit",
  networkBridge: false,
  hostFileSystemBridge: false,
  hostProcessBridge: false,
  customBuiltins: false,
  freshFileSystemPerRun: true,
  physicalMemoryBytes: SHELL_MAX_MEMORY_BYTES,
  interpreterAccountingCeilingBytes:
    SHELL_INTERPRETER_ACCOUNTING_CEILING_BYTES,
  sourceBytes: SHELL_MAX_SOURCE_BYTES,
  stdinBytes: SHELL_MAX_STDIN_BYTES,
  outputBytes: SHELL_MAX_OUTPUT_BYTES,
  maxCommands: SHELL_MAX_COMMANDS,
  maxLoopIterations: SHELL_MAX_LOOP_ITERATIONS,
  defaultWallTimeMs: 5_000,
});
export const SHELL_EXECUTION_POLICY_CANONICAL_JSON = JSON.stringify(
  SHELL_EXECUTION_POLICY_MANIFEST,
);
export const SHELL_EXECUTION_POLICY_SHA256 =
  "c1a9047eda99f0e21f004b84a987109e443f6139aebf84f65f65b96e8efb9a48";

export type ShellExecutionStatus =
  | "succeeded"
  | "compile-error"
  | "runtime-error"
  | "quota-exceeded";

interface VarUintReceipt {
  value: number;
  offset: number;
}

function readVarUint(
  bytes: Uint8Array,
  start: number,
  limit: number,
): VarUintReceipt {
  let offset = start;
  let value = 0;
  let shift = 0;
  while (offset < limit && shift <= 28) {
    const byte = bytes[offset++];
    value += (byte & 0x7f) * 2 ** shift;
    if ((byte & 0x80) === 0 && Number.isSafeInteger(value)) {
      return { value, offset };
    }
    shift += 7;
  }
  throw new Error("Shell runtime contains a malformed WebAssembly integer.");
}

function encodeVarUint(value: number): number[] {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error("Shell WebAssembly memory limit is invalid.");
  }
  const encoded: number[] = [];
  let remaining = value;
  do {
    let byte = remaining % 128;
    remaining = Math.floor(remaining / 128);
    if (remaining !== 0) byte |= 0x80;
    encoded.push(byte);
  } while (remaining !== 0);
  return encoded;
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array<ArrayBuffer> {
  const size = parts.reduce((sum, part) => sum + part.byteLength, 0);
  const result = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

/**
 * Add the missing physical maximum to the pinned module's single defined
 * memory. The raw asset hash is checked before this deterministic transform;
 * callers also check the transformed hash before compiling it.
 */
export function hardenShellWasmMemory(
  input: ArrayBuffer | Uint8Array,
): Uint8Array<ArrayBuffer> {
  const bytes =
    input instanceof Uint8Array
      ? new Uint8Array(input.buffer, input.byteOffset, input.byteLength)
      : new Uint8Array(input);
  if (
    bytes.byteLength < 8 ||
    bytes[0] !== 0x00 ||
    bytes[1] !== 0x61 ||
    bytes[2] !== 0x73 ||
    bytes[3] !== 0x6d ||
    bytes[4] !== 0x01 ||
    bytes[5] !== 0x00 ||
    bytes[6] !== 0x00 ||
    bytes[7] !== 0x00
  ) {
    throw new Error("Shell runtime is not the expected WebAssembly module.");
  }

  let offset = 8;
  let memorySectionStart = -1;
  let memoryPayloadEnd = -1;
  let minimumPages = -1;
  while (offset < bytes.byteLength) {
    const sectionStart = offset;
    const sectionId = bytes[offset++];
    const sectionSize = readVarUint(bytes, offset, bytes.byteLength);
    const payloadStart = sectionSize.offset;
    const payloadEnd = payloadStart + sectionSize.value;
    if (payloadEnd > bytes.byteLength) {
      throw new Error("Shell runtime contains a truncated WebAssembly section.");
    }
    if (sectionId === 5) {
      if (memorySectionStart !== -1) {
        throw new Error("Shell runtime declares multiple memory sections.");
      }
      let cursor = payloadStart;
      const count = readVarUint(bytes, cursor, payloadEnd);
      cursor = count.offset;
      const flags = readVarUint(bytes, cursor, payloadEnd);
      cursor = flags.offset;
      const minimum = readVarUint(bytes, cursor, payloadEnd);
      cursor = minimum.offset;
      if (count.value !== 1 || flags.value !== 0 || cursor !== payloadEnd) {
        throw new Error("Shell runtime memory shape is not the pinned contract.");
      }
      if (minimum.value > SHELL_MAX_MEMORY_PAGES) {
        throw new Error("Shell runtime minimum memory exceeds its hard limit.");
      }
      memorySectionStart = sectionStart;
      memoryPayloadEnd = payloadEnd;
      minimumPages = minimum.value;
    }
    offset = payloadEnd;
  }
  if (memorySectionStart === -1 || memoryPayloadEnd === -1) {
    throw new Error("Shell runtime has no owned WebAssembly memory.");
  }

  const payload = Uint8Array.from([
    1,
    1,
    ...encodeVarUint(minimumPages),
    ...encodeVarUint(SHELL_MAX_MEMORY_PAGES),
  ]);
  return concatBytes([
    bytes.subarray(0, memorySectionStart),
    Uint8Array.from([5, ...encodeVarUint(payload.byteLength)]),
    payload,
    bytes.subarray(memoryPayloadEnd),
  ]);
}

/**
 * Bashkit does not accept a separate stdin argument. Seed stdin into the fresh
 * in-memory VFS and redirect the authored top-level group from that file. The
 * opening group token shares authored line 1, so diagnostics do not gain a
 * synthetic leading line.
 */
export function shellCommandWithStdin(source: string): string {
  return `{ ${source}\n} < ${SHELL_STDIN_PATH}`;
}

export function classifyShellFailure(message: string): ShellExecutionStatus {
  if (/\bparse error\b|\bsyntax error\b|unexpected (?:token|end)/i.test(message)) {
    return "compile-error";
  }
  if (
    /resource limit exceeded|session command limit exceeded|trusted (?:output|memory) budget/i.test(
      message,
    )
  ) {
    return "quota-exceeded";
  }
  return "runtime-error";
}
