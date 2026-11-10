/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The document a compile-only language hands back instead of program output.
 *
 * LaTeX is the one such language: a successful run typesets `main.tex` on the
 * server runtime and the signed receipt carries `main.pdf` (name, media type,
 * size, SHA-256 and base64 bytes, at most 1 MB). Every other language's receipt
 * omits the field, and the client refuses it there.
 *
 * Two checks, two places:
 *  - {@link decodeExecutionArtifact} is the wire boundary. It takes the
 *    receipt's snake_case object, demands its exact shape, decodes the bytes
 *    and confirms their length, their `%PDF-` signature and their SHA-256
 *    before anything else sees them.
 *  - {@link isLanguageExecutionArtifact} is the registry's synchronous
 *    re-check of the decoded value (shape, length, signature) when it re-parses
 *    a provider's receipt.
 *
 * The bytes live only in memory for the output that produced them; nothing
 * here persists them.
 */
import { base64ToBytes } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";

/** The executor refuses a larger document, so a larger one is forged. */
export const MAX_EXECUTION_ARTIFACT_BYTES = 1_000_000;
const MAX_EXECUTION_ARTIFACT_BASE64_CHARS =
  Math.ceil(MAX_EXECUTION_ARTIFACT_BYTES / 3) * 4;
const SHA256_PATTERN = /^[0-9a-f]{64}$/u;
const BASE64_PATTERN =
  /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/u;

interface ExecutionArtifactKind {
  readonly name: string;
  readonly mediaType: string;
  /** The bytes a valid document of this kind starts with. */
  readonly signature: readonly number[];
}

const PDF_SIGNATURE = [0x25, 0x50, 0x44, 0x46, 0x2d]; // "%PDF-"

const ARTIFACT_KINDS: Readonly<Record<string, ExecutionArtifactKind>> = {
  latex: {
    name: "main.pdf",
    mediaType: "application/pdf",
    signature: PDF_SIGNATURE,
  },
};

export interface LanguageExecutionArtifact {
  readonly name: string;
  readonly mediaType: string;
  readonly sizeBytes: number;
  readonly sha256: string;
  readonly bytes: Uint8Array;
}

/** Whether `language` compiles to a document instead of running a program. */
export function isCompileOnlyLanguage(language: string): boolean {
  return Object.prototype.hasOwnProperty.call(ARTIFACT_KINDS, language);
}

function startsWithSignature(
  bytes: Uint8Array,
  signature: readonly number[],
): boolean {
  if (bytes.length < signature.length) return false;
  return signature.every((byte, index) => bytes[index] === byte);
}

/** A Uint8Array from any realm (a worker's, a test environment's). */
function isByteArray(value: unknown): value is Uint8Array {
  return (
    ArrayBuffer.isView(value) &&
    Object.prototype.toString.call(value) === "[object Uint8Array]"
  );
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

/** The lowercase hex SHA-256 of `bytes` (WebCrypto). */
export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error("SHA-256 is unavailable.");
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  const digest = await subtle.digest("SHA-256", owned.buffer);
  return Array.from(new Uint8Array(digest), (value) =>
    value.toString(16).padStart(2, "0"),
  ).join("");
}

/**
 * Decode and verify the receipt's `artifact` object for `language`, or null
 * when anything about it is wrong. The caller treats null as a malformed
 * receipt.
 */
export async function decodeExecutionArtifact(
  raw: unknown,
  language: string,
): Promise<LanguageExecutionArtifact | null> {
  const kind = isCompileOnlyLanguage(language) ? ARTIFACT_KINDS[language] : null;
  if (
    !kind ||
    typeof raw !== "object" ||
    raw === null ||
    Array.isArray(raw)
  ) {
    return null;
  }
  const record = raw as Record<string, unknown>;
  if (
    !exactKeys(record, [
      "content_base64",
      "media_type",
      "name",
      "sha256",
      "size_bytes",
    ]) ||
    record.name !== kind.name ||
    record.media_type !== kind.mediaType ||
    !Number.isSafeInteger(record.size_bytes) ||
    Number(record.size_bytes) < 1 ||
    Number(record.size_bytes) > MAX_EXECUTION_ARTIFACT_BYTES ||
    typeof record.sha256 !== "string" ||
    !SHA256_PATTERN.test(record.sha256) ||
    typeof record.content_base64 !== "string" ||
    record.content_base64.length === 0 ||
    record.content_base64.length > MAX_EXECUTION_ARTIFACT_BASE64_CHARS ||
    record.content_base64.length % 4 !== 0 ||
    !BASE64_PATTERN.test(record.content_base64)
  ) {
    return null;
  }
  let bytes: Uint8Array;
  try {
    bytes = base64ToBytes(record.content_base64);
  } catch {
    return null;
  }
  const sizeBytes = Number(record.size_bytes);
  if (
    bytes.length !== sizeBytes ||
    !startsWithSignature(bytes, kind.signature)
  ) {
    return null;
  }
  let digest: string;
  try {
    digest = await sha256Hex(bytes);
  } catch {
    return null;
  }
  if (digest !== record.sha256) return null;
  return {
    name: kind.name,
    mediaType: kind.mediaType,
    sizeBytes,
    sha256: digest,
    bytes,
  };
}

/** Synchronous re-check of an already decoded artifact for `language`. */
export function isLanguageExecutionArtifact(
  value: unknown,
  language: string,
): value is LanguageExecutionArtifact {
  const kind = isCompileOnlyLanguage(language) ? ARTIFACT_KINDS[language] : null;
  if (!kind || typeof value !== "object" || value === null) return false;
  const artifact = value as Record<string, unknown>;
  return (
    exactKeys(artifact, ["bytes", "mediaType", "name", "sha256", "sizeBytes"]) &&
    artifact.name === kind.name &&
    artifact.mediaType === kind.mediaType &&
    Number.isSafeInteger(artifact.sizeBytes) &&
    Number(artifact.sizeBytes) >= 1 &&
    Number(artifact.sizeBytes) <= MAX_EXECUTION_ARTIFACT_BYTES &&
    typeof artifact.sha256 === "string" &&
    SHA256_PATTERN.test(artifact.sha256) &&
    isByteArray(artifact.bytes) &&
    artifact.bytes.length === artifact.sizeBytes &&
    startsWithSignature(artifact.bytes, kind.signature)
  );
}
