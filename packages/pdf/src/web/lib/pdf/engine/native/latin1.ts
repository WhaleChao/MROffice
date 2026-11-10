/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const CHUNK = 0x8000;

// A long stream's text, kept for the next asking of the same bytes (every
// planner of a page decodes its content stream again: 0.1 s a time for a

const DECODED_TEXTS = 2;
const DECODED_TEXT_MIN_BYTES = 1024 * 1024;
const decodedTexts: { bytes: Uint8Array; text: string }[] = [];

export function decodeLatin1(bytes: Uint8Array): string {
  if (bytes.length < DECODED_TEXT_MIN_BYTES) return decodeLatin1Uncached(bytes);
  const at = decodedTexts.findIndex((entry) => entry.bytes.length === bytes.length &&
    entry.bytes.every((value, index) => value === bytes[index]));
  if (at >= 0) {
    const [entry] = decodedTexts.splice(at, 1);
    decodedTexts.unshift(entry);
    return entry.text;
  }
  const text = decodeLatin1Uncached(bytes);
  decodedTexts.unshift({ bytes: Uint8Array.from(bytes), text });
  decodedTexts.length = Math.min(decodedTexts.length, DECODED_TEXTS);
  return text;
}

function decodeLatin1Uncached(bytes: Uint8Array): string {
  if (bytes.length <= CHUNK) return String.fromCharCode.apply(null, bytes as unknown as number[]);
  const parts: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += CHUNK) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(offset, offset + CHUNK) as unknown as number[]));
  }
  return parts.join("");
}

/** A drop-in for the "latin1" TextDecoder that decodes byte for byte. */
export const LATIN1_EXACT: Readonly<{ decode: (bytes: Uint8Array) => string }> = Object.freeze({
  decode: decodeLatin1,
});
