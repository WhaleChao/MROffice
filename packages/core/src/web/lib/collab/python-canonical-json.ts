/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The canonical JSON text the API hashes a page body as
 * (apps/api/app/core/json_digest.py): Python's
 * `json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)`,
 * reproduced character for character so both sides name one body by one digest.
 *
 * Where the two languages spell a value differently, this follows Python:
 * keys in code point order, an integral number within the safe range as an
 * integer (the API's projections hold those as `int`), any other number as
 * Python's `repr(float)`. Strings need nothing: `JSON.stringify` and Python
 * escape the same characters the same way when non-ASCII text is left raw.
 * A value JSON cannot carry (NaN, an infinity, a BigInt, a lone surrogate the
 * API could not have encoded) throws.
 */
import { isSharedJsonRead } from "./json-crdt";

const SAFE = Number.MAX_SAFE_INTEGER;

/** `repr(float)` for a finite number that is not a safe integer. */
function pythonFloatRepr(value: number): string {
  if (Object.is(value, -0) || value === 0) return Object.is(value, -0) ? "-0.0" : "0.0";
  // Both languages print the shortest digits that read back as this double.
  const [mantissa, exponentText] = value.toExponential().split("e");
  const exponent = Number(exponentText);
  const negative = mantissa.startsWith("-");
  const digits = mantissa.replace("-", "").replace(".", "");
  let text: string;
  if (exponent < -4 || exponent >= 16) {
    const fraction = digits.length > 1 ? `.${digits.slice(1)}` : "";
    const magnitude = Math.abs(exponent);
    text = `${digits[0]}${fraction}e${exponent < 0 ? "-" : "+"}${magnitude < 10 ? "0" : ""}${magnitude}`;
  } else if (exponent < 0) {
    text = `0.${"0".repeat(-exponent - 1)}${digits}`;
  } else if (digits.length <= exponent + 1) {
    text = `${digits}${"0".repeat(exponent + 1 - digits.length)}.0`;
  } else {
    text = `${digits.slice(0, exponent + 1)}.${digits.slice(exponent + 1)}`;
  }
  return negative ? `-${text}` : text;
}

/** UTF-16 code unit mapped so that comparing units compares code points. */
export const compareStringsByCodePoint = (left: string, right: string) => compareByCodePoint(left, right);
const codePointRank = (unit: number) => unit < 0xd800 ? unit : unit < 0xe000 ? unit + 0x2000 : unit - 0x800;

function compareByCodePoint(left: string, right: string): number {
  const length = Math.min(left.length, right.length);
  for (let index = 0; index < length; index += 1) {
    const a = left.charCodeAt(index), b = right.charCodeAt(index);
    if (a !== b) return codePointRank(a) - codePointRank(b);
  }
  return left.length - right.length;
}

// With the `u` flag a well-formed pair is one code point above U+FFFF, so the
// surrogate range matches only a half on its own (no lookbehind, which Safari
// before 16.4 cannot compile).
const LONE_SURROGATE = /[\ud800-\udfff]/u;

function writeString(value: string): string {
  if (LONE_SURROGATE.test(value)) throw new TypeError("A lone surrogate has no canonical JSON form");
  return JSON.stringify(value);
}

/** `json.dumps(value, sort_keys=True)` with Python's defaults: ASCII only,
 * `", "` and `": "` between entries. The API orders a text node's marks by
 * this spelling of their attributes (document_page_causal_crdt._normalize). */
export function pythonDefaultSortedJson(value: unknown): string {
  return pythonCanonicalJson(value, { itemSeparator: ", ", keySeparator: ": ", asciiOnly: true });
}

const NON_ASCII = /[^\x20-\x7e]/gu;
const asciiEscapes = (text: string) => text.replace(NON_ASCII, (char) => {
  let escaped = "";
  for (let index = 0; index < char.length; index += 1) escaped += `\\u${char.charCodeAt(index).toString(16).padStart(4, "0")}`;
  return escaped;
});

/**
 * The canonical text of a frozen object, and how deep it nests: a document's
 * projection keeps its blocks frozen from one save to the next
 * (document-page-causal-model.ts), and the save digested the whole body each
 * time (50-120 ms of an inked page under the pen). A shared read of a room
 * (json-crdt `yToJsonShared`) is read-only in the same way: a deck's body
 * keeps every slide, table row and cell an edit did not touch. Only the API's
 * own form (no options) is kept.
 */
const frozenTexts = new WeakMap<object, { readonly text: string; readonly height: number }>();

export function pythonCanonicalJson(
  value: unknown,
  options: { itemSeparator?: string; keySeparator?: string; asciiOnly?: boolean } = {},
): string {
  const itemSeparator = options.itemSeparator ?? ",", keySeparator = options.keySeparator ?? ":";
  const text = options.asciiOnly
    // JSON.stringify already wrote the control characters, the quote and the
    // backslash as Python does; what is left outside printable ASCII is raw.
    ? (item: string) => asciiEscapes(writeString(item))
    : writeString;
  const keep = options.itemSeparator === undefined && options.keySeparator === undefined && !options.asciiOnly;
  const parts: string[] = [];
  // The deepest level the object being written reaches (for a kept text).
  let reached = 0;
  const write = (item: unknown, depth: number): void => {
    if (keep && item !== null && typeof item === "object" && (Object.isFrozen(item) || isSharedJsonRead(item))) {
      const known = frozenTexts.get(item);
      if (known) {
        if (depth + known.height > 512) throw new RangeError("Canonical JSON nesting is too deep");
        parts.push(known.text);
        reached = Math.max(reached, depth + known.height);
        return;
      }
      const from = parts.length, outer = reached;
      reached = depth;
      writeValue(item, depth);
      frozenTexts.set(item, { text: parts.splice(from).join(""), height: reached - depth });
      parts.push(frozenTexts.get(item)!.text);
      reached = Math.max(outer, reached);
      return;
    }
    writeValue(item, depth);
  };
  const writeValue = (item: unknown, depth: number): void => {
    if (depth > 512) throw new RangeError("Canonical JSON nesting is too deep");
    if (depth > reached) reached = depth;
    if (item === null) { parts.push("null"); return; }
    switch (typeof item) {
      case "boolean": parts.push(item ? "true" : "false"); return;
      case "string": parts.push(text(item)); return;
      case "number":
        if (!Number.isFinite(item)) throw new TypeError("A non-finite number has no canonical JSON form");
        parts.push(Number.isInteger(item) && Math.abs(item) <= SAFE ? String(item === 0 ? 0 : item) : pythonFloatRepr(item));
        return;
      case "object": break;
      default: throw new TypeError(`A ${typeof item} has no canonical JSON form`);
    }
    if (Array.isArray(item)) {
      parts.push("[");
      for (let index = 0; index < item.length; index += 1) {
        if (index) parts.push(itemSeparator);
        // JSON.stringify writes a hole or an undefined entry as null.
        write(item[index] === undefined ? null : item[index], depth + 1);
      }
      parts.push("]");
      return;
    }
    const record = item as Record<string, unknown>;
    const keys = Object.keys(record).filter((key) => record[key] !== undefined).sort(compareByCodePoint);
    parts.push("{");
    for (let index = 0; index < keys.length; index += 1) {
      if (index) parts.push(itemSeparator);
      parts.push(text(keys[index]), keySeparator);
      write(record[keys[index]], depth + 1);
    }
    parts.push("}");
  };
  write(value, 0);
  return parts.join("");
}

/** `canonical_json_sha256` of the API, as lowercase hex. */
export async function pythonCanonicalJsonSha256(value: unknown): Promise<string> {
  const bytes = new TextEncoder().encode(pythonCanonicalJson(value));
  const digest = await globalThis.crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}
