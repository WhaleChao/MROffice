/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Trusted Ruby 4.0 WASI runtime used only inside the dedicated Ruby Worker.
 *
 * The official ruby.wasm package supplies CRuby. This boundary supplies an
 * empty in-memory WASI filesystem, bounded stdin/stdout/stderr, no environment
 * variables, no sockets, and a hard maximum on the module's otherwise
 * unbounded WebAssembly memory. Ruby's optional JS bridge is replaced with
 * fail-closed imports so authored code cannot reach Worker globals, fetch,
 * credentials, or the control channel.
 */

export const RUBY_EXECUTION_PROTOCOL = "mrnote.ruby-execution.v1";
export const RUBY_EXECUTION_RUNTIME = "ruby-4.0-wasm-wasi-2.10.1";
export const RUBY_WASM_SHA256 =
  "9fe3c749730a16da9b0d1a973bb77bec164f8d958efcb7f5b81e4275ea8c5439";

export const RUBY_MAX_SOURCE_BYTES = 2 * 1024 * 1024;
export const RUBY_MAX_STDIN_BYTES = 100 * 1024;
export const RUBY_MAX_OUTPUT_BYTES = 300_000;
export const RUBY_MAX_MEMORY_BYTES = 256 * 1024 * 1024;

const RUBY_STATIC_MEDIA_PREFIX = "/_next/static/media/";

const WASM_PAGE_BYTES = 64 * 1024;
const RUBY_MAX_MEMORY_PAGES = RUBY_MAX_MEMORY_BYTES / WASM_PAGE_BYTES;
const MAX_LOG_LINES = 2_000;
const MAX_LOG_LINE_CHARS = 20_000;
const MAX_ERROR_CHARS = 50_000;
const encoder = new TextEncoder();

export type RubyExecutionStatus =
  | "succeeded"
  | "compile-error"
  | "runtime-error"
  | "quota-exceeded";

export interface RubyRuntimeReceipt {
  readonly status: RubyExecutionStatus;
  readonly logs: string[];
  readonly logStreams: Array<"stdout" | "stderr">;
  readonly error?: string;
  /** The VM may no longer be safely reused after this receipt. */
  readonly retireGeneration?: true;
}

export interface RubyExecutionRuntime {
  execute(source: string, stdin: string, outputBytes: number): RubyRuntimeReceipt;
}

/**
 * Bind Next's emitted Ruby asset path to the exact Worker origin.
 *
 * Turbopack represents emitted assets with a URL-like object whose `origin`
 * and `protocol` fields are blank even though `toString()` returns the correct
 * same-origin path. Treat that object only as a path string, then apply the
 * same closed authority checks to the resolved URL that the Worker fetches.
 */
export function resolveRubyWasmAssetUrl(
  assetReference: string,
  workerOrigin: string,
): URL {
  let origin: URL;
  let resolved: URL;
  try {
    if (
      !assetReference ||
      assetReference !== assetReference.trim() ||
      assetReference.includes("\0")
    ) {
      throw new Error("invalid asset reference");
    }
    origin = new URL(workerOrigin);
    if (
      origin.origin !== workerOrigin ||
      (origin.protocol !== "http:" && origin.protocol !== "https:") ||
      origin.username ||
      origin.password ||
      origin.pathname !== "/" ||
      origin.search ||
      origin.hash
    ) {
      throw new Error("invalid worker origin");
    }
    resolved = new URL(assetReference, `${origin.origin}/`);
  } catch {
    throw new Error("Ruby runtime asset is not same-origin.");
  }
  if (
    (resolved.protocol !== "http:" && resolved.protocol !== "https:") ||
    resolved.origin !== origin.origin ||
    resolved.username ||
    resolved.password ||
    resolved.search ||
    resolved.hash ||
    !resolved.pathname.startsWith(RUBY_STATIC_MEDIA_PREFIX) ||
    !resolved.pathname.endsWith(".wasm")
  ) {
    throw new Error("Ruby runtime asset is not same-origin.");
  }
  return resolved;
}

class RubyOutputQuotaError extends Error {
  constructor(message = "Ruby output exceeded its trusted budget.") {
    super(message);
    this.name = "RubyOutputQuotaError";
  }
}

const RUBY_VM_BUGREPORT_FAILURE =
  "Tried to report internal Ruby VM state but failed: ";

/**
 * Ruby's JS wrapper treats a host-side output quota as an unexpected VM error
 * and tries to print a second VM bug report through the same already-sealed
 * stderr sink. That second write necessarily fails, after which the wrapper
 * logs an internal stack to the browser console. Suppress only that exact,
 * expected quota-report failure; every other console error remains untouched.
 */
function runRubyVmOperation<T>(operation: () => T): T {
  const consoleOwner = globalThis.console;
  const originalError = consoleOwner?.error;
  if (!consoleOwner || typeof originalError !== "function") {
    return operation();
  }
  const filteredError = (...args: unknown[]) => {
    if (
      args[0] === RUBY_VM_BUGREPORT_FAILURE &&
      args[1] instanceof RubyOutputQuotaError
    ) {
      return;
    }
    originalError.apply(consoleOwner, args);
  };
  let installed = false;
  try {
    consoleOwner.error = filteredError;
    installed = consoleOwner.error === filteredError;
  } catch {
    // A frozen console still executes correctly; only noise suppression is lost.
  }
  try {
    return operation();
  } finally {
    if (installed && consoleOwner.error === filteredError) {
      try {
        consoleOwner.error = originalError;
      } catch {
        // Do not replace the authored execution result with console cleanup.
      }
    }
  }
}

function decodeU32(
  bytes: Uint8Array,
  offset: number,
  limit = bytes.length,
): { value: number; next: number } {
  let value = 0;
  let shift = 0;
  let cursor = offset;
  while (cursor < limit && cursor - offset < 5) {
    const byte = bytes[cursor++];
    value |= (byte & 0x7f) << shift;
    if ((byte & 0x80) === 0) {
      return { value: value >>> 0, next: cursor };
    }
    shift += 7;
  }
  throw new Error("Ruby WASM contains an invalid unsigned integer.");
}

function encodeU32(value: number): Uint8Array {
  if (!Number.isSafeInteger(value) || value < 0 || value > 0xffffffff) {
    throw new Error("Ruby WASM contains an invalid memory limit.");
  }
  const encoded: number[] = [];
  let remaining = value >>> 0;
  do {
    let byte = remaining & 0x7f;
    remaining >>>= 7;
    if (remaining !== 0) byte |= 0x80;
    encoded.push(byte);
  } while (remaining !== 0);
  return Uint8Array.from(encoded);
}

function joinBytes(parts: readonly Uint8Array[]): Uint8Array {
  const size = parts.reduce((total, part) => total + part.byteLength, 0);
  const joined = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    joined.set(part, offset);
    offset += part.byteLength;
  }
  return joined;
}

/**
 * Add or tighten the module-defined memory maximum without changing any other
 * section. The published Ruby binary declares an initial memory but no maximum;
 * leaving that ambient would make the UI's memory receipt untrue.
 */
export function capRubyWasmMemory(
  input: ArrayBuffer | Uint8Array,
): Uint8Array {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
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
    throw new Error("Ruby runtime asset is not a supported WebAssembly module.");
  }

  const parts: Uint8Array[] = [bytes.slice(0, 8)];
  let cursor = 8;
  let memorySections = 0;
  while (cursor < bytes.byteLength) {
    const sectionStart = cursor;
    const sectionId = bytes[cursor++];
    const size = decodeU32(bytes, cursor);
    const payloadStart = size.next;
    const payloadEnd = payloadStart + size.value;
    if (payloadEnd > bytes.byteLength) {
      throw new Error("Ruby WASM contains a truncated section.");
    }
    cursor = payloadEnd;
    if (sectionId !== 5) {
      parts.push(bytes.slice(sectionStart, payloadEnd));
      continue;
    }

    memorySections += 1;
    const count = decodeU32(bytes, payloadStart, payloadEnd);
    if (count.value !== 1) {
      throw new Error("Ruby WASM must define exactly one memory.");
    }
    const flags = decodeU32(bytes, count.next, payloadEnd);
    // memory64/shared memories need a different host policy and are not part of
    // the pinned Ruby 4.0 artifact accepted by this runtime.
    if ((flags.value & ~1) !== 0) {
      throw new Error("Ruby WASM uses an unsupported memory type.");
    }
    const minimum = decodeU32(bytes, flags.next, payloadEnd);
    let end = minimum.next;
    let maximum = RUBY_MAX_MEMORY_PAGES;
    if ((flags.value & 1) !== 0) {
      const declaredMaximum = decodeU32(bytes, minimum.next, payloadEnd);
      maximum = Math.min(declaredMaximum.value, RUBY_MAX_MEMORY_PAGES);
      end = declaredMaximum.next;
    }
    if (
      end !== payloadEnd ||
      minimum.value > RUBY_MAX_MEMORY_PAGES ||
      maximum < minimum.value
    ) {
      throw new Error("Ruby WASM exceeds the trusted memory budget.");
    }
    const payload = joinBytes([
      encodeU32(1),
      encodeU32(flags.value | 1),
      encodeU32(minimum.value),
      encodeU32(maximum),
    ]);
    parts.push(Uint8Array.of(sectionId), encodeU32(payload.byteLength), payload);
  }
  if (memorySections !== 1) {
    throw new Error("Ruby WASM memory authority is unavailable.");
  }
  return joinBytes(parts);
}

function boundedError(error: unknown): string {
  let message = "Ruby execution failed.";
  try {
    if (error instanceof Error && error.message.trim()) {
      message = error.message.trim();
    } else if (typeof error === "string" && error.trim()) {
      message = error.trim();
    }
  } catch {
    return message;
  }
  return message.slice(0, MAX_ERROR_CHARS);
}

class RubyOutputCollector {
  readonly logs: string[] = [];
  readonly logStreams: Array<"stdout" | "stderr"> = [];
  private readonly decoders = {
    stdout: new TextDecoder("utf-8", { fatal: false }),
    stderr: new TextDecoder("utf-8", { fatal: false }),
  };
  private readonly carries = { stdout: "", stderr: "" };
  private usedBytes = 0;

  constructor(private readonly outputBytes: number) {}

  write(bytes: Uint8Array, stream: "stdout" | "stderr"): void {
    if (this.usedBytes + bytes.byteLength > this.outputBytes) {
      throw new RubyOutputQuotaError();
    }
    this.usedBytes += bytes.byteLength;
    this.acceptText(this.decoders[stream].decode(bytes, { stream: true }), stream);
  }

  finish(): void {
    for (const stream of ["stdout", "stderr"] as const) {
      this.acceptText(this.decoders[stream].decode(), stream);
      const carry = this.carries[stream];
      if (carry) this.appendLine(carry, stream);
      this.carries[stream] = "";
    }
  }

  private acceptText(text: string, stream: "stdout" | "stderr"): void {
    const pieces = (this.carries[stream] + text).split("\n");
    this.carries[stream] = pieces.pop() ?? "";
    if (this.carries[stream].length > MAX_LOG_LINE_CHARS) {
      throw new RubyOutputQuotaError("Ruby output line exceeded its trusted budget.");
    }
    for (const line of pieces) this.appendLine(line, stream);
  }

  private appendLine(line: string, stream: "stdout" | "stderr"): void {
    if (
      line.length > MAX_LOG_LINE_CHARS ||
      this.logs.length >= MAX_LOG_LINES
    ) {
      throw new RubyOutputQuotaError();
    }
    this.logs.push(line.endsWith("\r") ? line.slice(0, -1) : line);
    this.logStreams.push(stream);
  }
}

function blockRubyJsBridge(imports: WebAssembly.Imports): void {
  const host = imports["rb-js-abi-host"];
  if (!host || typeof host !== "object") {
    throw new Error("Ruby JS bridge import authority is unavailable.");
  }
  for (const key of Object.keys(host)) {
    (host as Record<string, WebAssembly.ImportValue>)[key] = () => {
      throw new Error("Ruby host interop is unavailable in this sandbox.");
    };
  }
}

function isFatalRuntimeError(error: unknown): boolean {
  try {
    return (
      error instanceof WebAssembly.RuntimeError ||
      (error instanceof Error &&
        (error.name === "RbFatalError" || error.name === "WASIProcExit"))
    );
  } catch {
    return true;
  }
}

function isRubySyntaxError(error: unknown): boolean {
  try {
    return error instanceof Error && /(?:^|\b)SyntaxError(?:\b|:)/.test(error.message);
  } catch {
    return false;
  }
}

const RUBY_SANDBOX_BOOTSTRAP = String.raw`
module Kernel
  alias_method :__mrai_original_require, :require
  def require(feature)
    if feature == "js" || feature.start_with?("js/")
      raise SecurityError, "Ruby host interop is unavailable in this sandbox."
    end
    __mrai_original_require(feature)
  end
  private :__mrai_original_require
end
`;

export async function createRubyExecutionRuntime(
  module: WebAssembly.Module,
): Promise<RubyExecutionRuntime> {
  if (
    WebAssembly.Module.imports(module).some((entry) => entry.kind === "memory")
  ) {
    throw new Error("Ruby WASM imports an untrusted memory.");
  }
  const [{ RubyVM }, wasiModule] = await Promise.all([
    import("@ruby/wasm-wasi"),
    import("@bjorn3/browser_wasi_shim"),
  ]);
  const { WASI, File, OpenFile, ConsoleStdout, PreopenDirectory } = wasiModule;
  const stdinFile = new File(new Uint8Array(), { readonly: true });
  const stdinFd = new OpenFile(stdinFile);
  let activeOutput: RubyOutputCollector | null = null;
  const stdout = new ConsoleStdout((bytes) =>
    activeOutput?.write(bytes, "stdout"),
  );
  const stderr = new ConsoleStdout((bytes) =>
    activeOutput?.write(bytes, "stderr"),
  );
  const wasi = new WASI(
    [],
    [],
    [stdinFd, stdout, stderr, new PreopenDirectory("/", new Map())],
    { debug: false },
  );
  const { vm } = await RubyVM.instantiateModule({
    module,
    wasip1: wasi,
    addToImports: blockRubyJsBridge,
  });
  vm.eval(RUBY_SANDBOX_BOOTSTRAP);

  return {
    execute(source, stdin, outputBytes) {
      if (activeOutput) {
        return {
          status: "runtime-error",
          logs: [],
          logStreams: [],
          error: "Ruby runtime is already executing.",
          retireGeneration: true,
        };
      }
      stdinFile.data = encoder.encode(stdin);
      stdinFd.file_pos = 0n;
      const output = new RubyOutputCollector(outputBytes);
      activeOutput = output;
      let status: RubyExecutionStatus = "succeeded";
      let errorMessage: string | undefined;
      let retireGeneration = false;
      try {
        runRubyVmOperation(() => vm.eval(source));
      } catch (error) {
        errorMessage = boundedError(error);
        if (error instanceof RubyOutputQuotaError) {
          status = "quota-exceeded";
          retireGeneration = true;
        } else if (isRubySyntaxError(error)) {
          status = "compile-error";
        } else {
          status = "runtime-error";
          retireGeneration = isFatalRuntimeError(error);
        }
      } finally {
        try {
          output.finish();
        } catch (error) {
          status = "quota-exceeded";
          errorMessage = boundedError(error);
          retireGeneration = true;
        }
        activeOutput = null;
      }
      return {
        status,
        logs: [...output.logs],
        logStreams: [...output.logStreams],
        ...(errorMessage ? { error: errorMessage } : {}),
        ...(retireGeneration ? { retireGeneration: true } : {}),
      };
    },
  };
}
