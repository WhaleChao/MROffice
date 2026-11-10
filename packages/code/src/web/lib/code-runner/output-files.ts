/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The compile's own reports a LaTeX receipt returns for the person reading
 * them, never for the next compile: TeX's log (on a failed compile too) and
 * SyncTeX's map between source lines and positions in the PDF.
 *
 * Mirrors the API's OutputFiles and the Worker's internal/protocol/
 * output_files.go: the same names, order and bounds. A log too long to travel
 * whole keeps its end, where TeX stopped, and says how many leading bytes were
 * left out; the SyncTeX map travels whole or not at all.
 *
 * {@link outputFilesFromWire} is the wire boundary (snake_case, base64) and
 * decodes the bytes; {@link isLanguageExecutionOutputFiles} is the registry's
 * synchronous re-check of the decoded value. The bytes live only in memory.
 */
import { base64ToBytes } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";

export const MAX_OUTPUT_FILES_BYTES = 768_000;
export const MAX_OUTPUT_LOG_BYTES = 512_000;
/** Far beyond any log a sandbox could write. */
const MAX_OMITTED_LOG_BYTES = 2 ** 40;

export const OUTPUT_LOG_NAME = "main.log";
export const OUTPUT_SYNCTEX_NAME = "main.synctex.gz";
const OUTPUT_FILE_NAMES: readonly string[] = [OUTPUT_LOG_NAME, OUTPUT_SYNCTEX_NAME];
const OUTPUT_FILE_LANGUAGES: ReadonlySet<string> = new Set(["latex"]);
/** The outcomes whose receipt may carry them: a compile that finished. */
const OUTPUT_FILE_OUTCOMES: ReadonlySet<string> = new Set(["succeeded", "compile-error"]);
const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;
const GZIP_MAGIC = [0x1f, 0x8b];

export interface LanguageExecutionOutputFile {
  readonly name: string;
  readonly bytes: Uint8Array;
  /** The leading bytes of a log too long to return whole (0: none). */
  readonly omittedBytes: number;
}

export interface LanguageExecutionOutputFiles {
  readonly files: readonly LanguageExecutionOutputFile[];
}

export function carriesOutputFiles(language: string, outcome: unknown): boolean {
  return OUTPUT_FILE_LANGUAGES.has(language) && OUTPUT_FILE_OUTCOMES.has(String(outcome));
}

function plainRecord(value: unknown, keys: readonly string[], optional: readonly string[] = []) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return null;
  const own = Object.keys(value);
  if (
    !keys.every((key) => own.includes(key)) ||
    !own.every((key) => keys.includes(key) || optional.includes(key))
  ) {
    return null;
  }
  return value as Record<string, unknown>;
}

/** A Uint8Array from any realm (a worker's, a test environment's). */
function isByteArray(value: unknown): value is Uint8Array {
  return (
    ArrayBuffer.isView(value) &&
    Object.prototype.toString.call(value) === "[object Uint8Array]"
  );
}

function validFile(name: string, bytes: Uint8Array, omittedBytes: number): boolean {
  if (!Number.isSafeInteger(omittedBytes) || omittedBytes < 0 || omittedBytes > MAX_OMITTED_LOG_BYTES) {
    return false;
  }
  if (name === OUTPUT_LOG_NAME) return bytes.byteLength <= MAX_OUTPUT_LOG_BYTES;
  return (
    name === OUTPUT_SYNCTEX_NAME &&
    omittedBytes === 0 &&
    bytes.byteLength >= GZIP_MAGIC.length &&
    GZIP_MAGIC.every((byte, index) => bytes[index] === byte)
  );
}

/** Whether `value` is the decoded output files of a `language` receipt. */
export function isLanguageExecutionOutputFiles(
  value: unknown,
  language: string,
): value is LanguageExecutionOutputFiles {
  if (!OUTPUT_FILE_LANGUAGES.has(language)) return false;
  const record = plainRecord(value, ["files"]);
  if (!record || !Array.isArray(record.files)) return false;
  const files = record.files as unknown[];
  if (files.length < 1 || files.length > OUTPUT_FILE_NAMES.length) return false;
  let previous = "";
  let total = 0;
  for (const entry of files) {
    const file = plainRecord(entry, ["name", "bytes", "omittedBytes"]);
    if (
      !file ||
      typeof file.name !== "string" ||
      !OUTPUT_FILE_NAMES.includes(file.name) ||
      file.name <= previous ||
      !isByteArray(file.bytes) ||
      typeof file.omittedBytes !== "number" ||
      !validFile(file.name, file.bytes, file.omittedBytes)
    ) {
      return false;
    }
    previous = file.name;
    total += file.bytes.byteLength;
    if (total > MAX_OUTPUT_FILES_BYTES) return false;
  }
  return true;
}

/** A receipt's wire-form output files, decoded and validated, or null. */
export function outputFilesFromWire(
  value: unknown,
  language: string,
): LanguageExecutionOutputFiles | null {
  const wire = plainRecord(value, ["files"]);
  if (!wire || !Array.isArray(wire.files)) return null;
  const files: LanguageExecutionOutputFile[] = [];
  for (const entry of wire.files as unknown[]) {
    const file = plainRecord(entry, ["name", "content_base64"], ["omitted_bytes"]);
    if (
      !file ||
      typeof file.name !== "string" ||
      typeof file.content_base64 !== "string" ||
      !BASE64.test(file.content_base64) ||
      file.content_base64.length > Math.ceil(MAX_OUTPUT_FILES_BYTES / 3) * 4 ||
      // The Worker omits a zero count; one written out is not its signed form.
      ("omitted_bytes" in file &&
        (typeof file.omitted_bytes !== "number" || file.omitted_bytes < 1))
    ) {
      return null;
    }
    let bytes: Uint8Array;
    try {
      bytes = base64ToBytes(file.content_base64);
    } catch {
      return null;
    }
    files.push({
      name: file.name,
      bytes,
      omittedBytes: typeof file.omitted_bytes === "number" ? file.omitted_bytes : 0,
    });
  }
  const decoded = { files };
  return isLanguageExecutionOutputFiles(decoded, language) ? decoded : null;
}

/** One of the receipt's output files by name, if it carried it. */
export function outputFile(
  files: LanguageExecutionOutputFiles | null | undefined,
  name: string,
): LanguageExecutionOutputFile | null {
  return files?.files.find((file) => file.name === name) ?? null;
}
