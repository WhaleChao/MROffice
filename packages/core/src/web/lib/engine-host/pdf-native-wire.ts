/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_WIRE_BYTES = "$u8";
export const PDF_WIRE_APPEND = "$append";
export const PDF_WIRE_MAP = "$map";
export const PDF_WIRE_SET = "$set";
export const PDF_WIRE_TYPED = "$typed";
/** A candidate built on the working head, by its SHA-256: the server holds
 * it (its candidate store) or it travels as an upload of the call. */
export const PDF_WIRE_CANDIDATE = "$candidate";
/** Any other large byte array (a picture), by its SHA-256: an upload of the call. */
export const PDF_WIRE_BLOB = "$blob";
/** A long text, as the `$blob` of its UTF-8 bytes: `{"$text": {"$blob": sha}}`.
 * A record a history call names carries what it installed as base64 (a font
 * extension's narrowed face, ~1.2 MB of JSON), which the call's 1 MiB JSON
 * cannot hold; the text travels beside it as a picture does. */
export const PDF_WIRE_TEXT = "$text";
/** At this size a byte array leaves the JSON (the call's request is 1 MiB),
 * and a text of this many characters. */
export const PDF_WIRE_OFFLOAD_BYTES = 64 * 1024;

const BLOB_SHA = /^[0-9a-f]{64}$/;

/**
 * Replace the `$candidate` / `$blob` stand-ins of a decoded request with the
 * bytes the call named (copies the planner may keep or zero). The engine
 * side of `encodePdfNativeWire`'s `offload`.
 */
export function resolvePdfNativeWireBlobs(value: unknown, blob: (sha256: string) => Uint8Array, depth = 0): unknown {
  if (depth > 64) throw new PdfWireMismatchError("Engine arguments are nested too deeply.");
  if (Array.isArray(value)) return value.map((item) => resolvePdfNativeWireBlobs(item, blob, depth + 1));
  if (value instanceof Map) {
    return new Map([...value].map(([key, item]) => [
      resolvePdfNativeWireBlobs(key, blob, depth + 1), resolvePdfNativeWireBlobs(item, blob, depth + 1)]));
  }
  if (value instanceof Set) return new Set([...value].map((item) => resolvePdfNativeWireBlobs(item, blob, depth + 1)));
  if (!value || typeof value !== "object" || ArrayBuffer.isView(value)) return value;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length === 1 && (keys[0] === PDF_WIRE_CANDIDATE || keys[0] === PDF_WIRE_BLOB)) {
    const sha = record[keys[0]];
    if (typeof sha !== "string" || !BLOB_SHA.test(sha)) throw new PdfWireMismatchError("Malformed engine blob reference.");
    return Uint8Array.from(blob(sha));
  }
  if (keys.length === 1 && keys[0] === PDF_WIRE_TEXT) {
    // The text exactly as it was sent: its UTF-8 bytes, decoded strictly.
    const inner = record[PDF_WIRE_TEXT];
    const sha = isRecord(inner) && Object.keys(inner).length === 1 ? inner[PDF_WIRE_BLOB] : undefined;
    if (typeof sha !== "string" || !BLOB_SHA.test(sha)) throw new PdfWireMismatchError("Malformed engine text reference.");
    try {
      return new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(blob(sha));
    } catch (error) {
      if (error instanceof TypeError) throw new PdfWireMismatchError("Malformed engine text reference.");
      throw error;
    }
  }
  return Object.fromEntries(keys.map((key) => [key, resolvePdfNativeWireBlobs(record[key], blob, depth + 1)]));
}

const TYPED_ARRAYS = {
  Int8Array, Uint8ClampedArray, Int16Array, Uint16Array, Int32Array, Uint32Array, Float32Array, Float64Array,
} as const;
/** Typed arrays other than Uint8Array a proof may hold (glyph tables, widths). */
export const PDF_WIRE_TYPED_KINDS: ReadonlySet<string> = new Set(Object.keys(TYPED_ARRAYS));

export interface PdfWireAppend {
  readonly base: string;
  readonly length: number;
  readonly suffix: string | readonly string[];
  readonly sha256: string;
}

/** The API bounds one JSON string at 1 MiB: a long base64 text (a picture a
 * command carries) travels as a list of parts, joined on arrival. */
const MAX_BASE64_PART = 768 * 1024;
function base64Chunks(text: string): string | string[] {
  if (text.length <= MAX_BASE64_PART) return text;
  const parts: string[] = [];
  for (let offset = 0; offset < text.length; offset += MAX_BASE64_PART) parts.push(text.slice(offset, offset + MAX_BASE64_PART));
  return parts;
}

export function bytesToBase64(bytes: Uint8Array): string {
  const NodeBuffer = (globalThis as { Buffer?: { from(data: Uint8Array): { toString(encoding: string): string } } }).Buffer;
  if (NodeBuffer) return NodeBuffer.from(bytes).toString("base64");
  let binary = "";
  const chunk = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunk) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunk));
  }
  return btoa(binary);
}

export function base64ToBytes(text: string): Uint8Array<ArrayBuffer> {
  const NodeBuffer = (globalThis as { Buffer?: { from(data: string, encoding: string): Uint8Array } }).Buffer;
  if (NodeBuffer) return Uint8Array.from(NodeBuffer.from(text, "base64"));
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  const digest = await crypto.subtle.digest("SHA-256", owned.buffer);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function asBytes(value: unknown): Uint8Array | null {
  if (value instanceof Uint8Array) return value;
  // A byte array from another realm (PDF.js's, a test environment's) is
  // still a byte array.
  if (ArrayBuffer.isView(value) && value.constructor?.name === "Uint8Array") {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  return null;
}

function startsWith(bytes: Uint8Array, prefix: Uint8Array): boolean {
  if (bytes.byteLength < prefix.byteLength) return false;
  const NodeBuffer = (globalThis as { Buffer?: { from(data: ArrayBufferLike, offset: number, length: number): { equals(other: unknown): boolean } } }).Buffer;
  if (NodeBuffer) {
    return NodeBuffer.from(bytes.buffer, bytes.byteOffset, prefix.byteLength)
      .equals(NodeBuffer.from(prefix.buffer, prefix.byteOffset, prefix.byteLength));
  }
  for (let index = 0; index < prefix.byteLength; index += 1) if (bytes[index] !== prefix[index]) return false;
  return true;
}

export interface PdfWireSource {
  readonly sha256: string;
  readonly bytes: Uint8Array;
}

/**
 * Tag every byte array (and map, set, typed array) of a value for the wire.
 * With a `source`, a byte array that starts with it -- the source itself, a
 * candidate appended to it -- travels as `$append` (only its suffix);
 * `keep` names such a candidate (the engine keeps it and answers its digest),
 * else its digest is computed here. Functions and `undefined` are dropped.
 */
export async function encodePdfNativeWire(
  value: unknown,
  source: PdfWireSource | null = null,
  keep?: (bytes: Uint8Array) => string,
  /** A large byte array's stand-in (`{"$candidate": sha}` / `{"$blob": sha}`)
   * when it travels beside the JSON instead of in it; null keeps it inline.
   * `appended`: it starts with `source` (a candidate built on the head). A
   * long text goes the same way as its UTF-8 bytes, under `$text`. */
  offload?: (bytes: Uint8Array, appended: boolean) => Promise<unknown | null>,
): Promise<unknown> {
  const visit = async (item: unknown): Promise<unknown> => {
    // A lone surrogate has no UTF-8 of its own: such a text stays inline.
    if (typeof item === "string" && offload && item.length >= PDF_WIRE_OFFLOAD_BYTES && !/[\uD800-\uDFFF]/u.test(item)) {
      const tag = await offload(new TextEncoder().encode(item), false);
      if (tag !== null) return { [PDF_WIRE_TEXT]: tag };
    }
    const bytes = asBytes(item);
    if (bytes) {
      const appended = Boolean(source && bytes.byteLength >= source.bytes.byteLength && startsWith(bytes, source.bytes));
      if (offload && !(source && appended && bytes.byteLength === source.bytes.byteLength)) {
        const tag = await offload(bytes, appended);
        if (tag !== null) return tag;
      }
      if (source && appended) {
        const sha256 = bytes.byteLength === source.bytes.byteLength
          ? source.sha256
          : keep ? keep(bytes) : await sha256Hex(bytes);
        return {
          [PDF_WIRE_APPEND]: {
            base: source.sha256,
            length: source.bytes.byteLength,
            suffix: base64Chunks(bytesToBase64(bytes.subarray(source.bytes.byteLength))),
            sha256,
          },
        };
      }
      return { [PDF_WIRE_BYTES]: base64Chunks(bytesToBase64(bytes)) };
    }
    if (Array.isArray(item)) {
      const out: unknown[] = new Array(item.length);
      for (let index = 0; index < item.length; index += 1) out[index] = await visit(item[index]);
      return out;
    }
    if (item instanceof Map) {
      const entries: unknown[] = [];
      for (const [key, entry] of item) entries.push([await visit(key), await visit(entry)]);
      return { [PDF_WIRE_MAP]: entries };
    }
    if (item instanceof Set) {
      const entries: unknown[] = [];
      for (const entry of item) entries.push(await visit(entry));
      return { [PDF_WIRE_SET]: entries };
    }
    if (ArrayBuffer.isView(item) && !(item instanceof DataView)) {
      const kind = item.constructor.name;
      if (!PDF_WIRE_TYPED_KINDS.has(kind)) throw new Error(`pdf-native wire cannot carry a ${kind}`);
      return { [PDF_WIRE_TYPED]: { kind, bytes: bytesToBase64(new Uint8Array(item.buffer, item.byteOffset, item.byteLength)) } };
    }
    if (item instanceof ArrayBuffer || typeof item === "bigint") {
      throw new Error("pdf-native wire carries plain JSON, byte arrays, maps and sets only");
    }
    if (item && typeof item === "object") {
      const out: Record<string, unknown> = {};
      for (const [key, entry] of Object.entries(item)) {
        if (entry !== undefined && typeof entry !== "function") out[key] = await visit(entry);
      }
      return out;
    }
    return item;
  };
  return visit(value);
}

/** A request's byte arrays as `$u8`, without a source (a few authorities). */
export function encodePdfNativeWireBytes(value: unknown): Promise<unknown> {
  return encodePdfNativeWire(value, null);
}

/** The answer's bytes do not rebuild what the engine said they were. */
export class PdfWireMismatchError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PdfWireMismatchError";
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Rebuild every tagged byte array of an engine answer. `base(sha256)` answers
 * the caller's own copy of a source the answer appends to (never copied by
 * the caller; read only). A candidate whose rebuilt digest differs is a
 * refusal, never a guess.
 */
export async function decodePdfNativeWire<T>(
  value: unknown,
  base: (sha256: string) => Uint8Array | null | undefined,
): Promise<T> {
  const visit = async (item: unknown): Promise<unknown> => {
    if (Array.isArray(item)) {
      const out: unknown[] = new Array(item.length);
      for (let index = 0; index < item.length; index += 1) out[index] = await visit(item[index]);
      return out;
    }
    if (!isRecord(item)) return item;
    const keys = Object.keys(item);
    if (keys.length === 1 && keys[0] === PDF_WIRE_BYTES && typeof item[PDF_WIRE_BYTES] === "string") {
      return base64ToBytes(item[PDF_WIRE_BYTES] as string);
    }
    if (keys.length === 1 && keys[0] === PDF_WIRE_BYTES && Array.isArray(item[PDF_WIRE_BYTES]) &&
      (item[PDF_WIRE_BYTES] as unknown[]).every((part) => typeof part === "string")) {
      return base64ToBytes((item[PDF_WIRE_BYTES] as string[]).join(""));
    }
    if (keys.length === 1 && keys[0] === PDF_WIRE_TYPED && isRecord(item[PDF_WIRE_TYPED])) {
      const typed = item[PDF_WIRE_TYPED] as { kind?: unknown; bytes?: unknown };
      const Kind = typeof typed.kind === "string" && PDF_WIRE_TYPED_KINDS.has(typed.kind)
        ? TYPED_ARRAYS[typed.kind as keyof typeof TYPED_ARRAYS] : null;
      if (!Kind || typeof typed.bytes !== "string") throw new PdfWireMismatchError("Malformed engine typed array.");
      const bytes = base64ToBytes(typed.bytes);
      if (bytes.byteLength % Kind.BYTES_PER_ELEMENT) throw new PdfWireMismatchError("Malformed engine typed array.");
      return new Kind(bytes.buffer, bytes.byteOffset, bytes.byteLength / Kind.BYTES_PER_ELEMENT);
    }
    if (keys.length === 1 && keys[0] === PDF_WIRE_MAP && Array.isArray(item[PDF_WIRE_MAP])) {
      const map = new Map<unknown, unknown>();
      for (const entry of item[PDF_WIRE_MAP] as unknown[]) {
        if (!Array.isArray(entry) || entry.length !== 2) throw new PdfWireMismatchError("Malformed engine map.");
        map.set(await visit(entry[0]), await visit(entry[1]));
      }
      return map;
    }
    if (keys.length === 1 && keys[0] === PDF_WIRE_SET && Array.isArray(item[PDF_WIRE_SET])) {
      const set = new Set<unknown>();
      for (const entry of item[PDF_WIRE_SET] as unknown[]) set.add(await visit(entry));
      return set;
    }
    if (keys.length === 1 && keys[0] === PDF_WIRE_APPEND && isRecord(item[PDF_WIRE_APPEND])) {
      const append = item[PDF_WIRE_APPEND] as unknown as PdfWireAppend;
      const source = base(append.base);
      if (!source || !Number.isSafeInteger(append.length) || append.length < 0 || append.length > source.byteLength) {
        throw new PdfWireMismatchError("The engine answer names a source this editor does not hold.");
      }
      const suffix = base64ToBytes(typeof append.suffix === "string" ? append.suffix : append.suffix.join(""));
      const bytes = new Uint8Array(append.length + suffix.byteLength);
      bytes.set(source.subarray(0, append.length), 0);
      bytes.set(suffix, append.length);
      if ((await sha256Hex(bytes)) !== append.sha256) {
        bytes.fill(0);
        throw new PdfWireMismatchError("The rebuilt PDF candidate does not match the engine's digest.");
      }
      return bytes;
    }
    const out: Record<string, unknown> = {};
    for (const key of keys) out[key] = await visit(item[key]);
    return out;
  };
  return (await visit(value)) as T;
}
