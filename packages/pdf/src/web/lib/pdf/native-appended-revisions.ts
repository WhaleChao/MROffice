/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const STARTXREF = [0x73, 0x74, 0x61, 0x72, 0x74, 0x78, 0x72, 0x65, 0x66]; // "startxref"
const EOF_MARK = [0x25, 0x25, 0x45, 0x4f, 0x46]; // "%%EOF"

const isSpace = (byte: number) => byte === 0x20 || byte === 0x0a || byte === 0x0d || byte === 0x09 || byte === 0x0c || byte === 0x00;

/**
 * The `startxref` keywords in `output` past `sourceLength`, when what follows
 * the prefix ends as a revision does (`startxref`, an offset, `%%EOF`). Null
 * where it does not: the caller counts the file's sections the long way.
 *
 * A count of one is exact: the bytes end with one revision's own keyword, so
 * no other stands among them (a stream that happens to spell the word makes
 * the count two, and is counted the long way too).
 */
export function countPdfAppendedStartxrefs(output: Uint8Array, sourceLength: number): number | null {
  if (!Number.isSafeInteger(sourceLength) || sourceLength < 0 || sourceLength >= output.length) return null;
  // %%EOF and white space end the file; before it an offset, before that the keyword.
  let end = output.length;
  while (end > sourceLength && isSpace(output[end - 1])) end -= 1;
  if (end - EOF_MARK.length < sourceLength) return null;
  for (let index = 0; index < EOF_MARK.length; index += 1) {
    if (output[end - EOF_MARK.length + index] !== EOF_MARK[index]) return null;
  }
  end -= EOF_MARK.length;
  while (end > sourceLength && isSpace(output[end - 1])) end -= 1;
  let digits = 0;
  while (end > sourceLength && output[end - 1] >= 0x30 && output[end - 1] <= 0x39) { end -= 1; digits += 1; }
  if (digits === 0) return null;
  while (end > sourceLength && isSpace(output[end - 1])) end -= 1;
  if (end - STARTXREF.length < sourceLength) return null;
  for (let index = 0; index < STARTXREF.length; index += 1) {
    if (output[end - STARTXREF.length + index] !== STARTXREF[index]) return null;
  }
  let count = 0;
  const last = output.length - STARTXREF.length;
  for (let at = sourceLength; at <= last; at += 1) {
    if (output[at] !== 0x73) continue;
    let matched = true;
    for (let index = 1; index < STARTXREF.length; index += 1) {
      if (output[at + index] !== STARTXREF[index]) { matched = false; break; }
    }
    if (matched) { count += 1; at += STARTXREF.length - 1; }
  }
  return count;
}

/**
 * The cross-reference sections of `output` -- `sourceLength` bytes of a
 * version of known depth and what was appended to it -- when the appended
 * bytes are exactly one revision. Null where they are not known to be: the
 * depth is then counted from the file.
 */
export function pdfDepthAfterOneAppendedRevision(sourceDepth: number | undefined, output: Uint8Array,
  sourceLength: number): number | null {
  if (sourceDepth === undefined || !Number.isSafeInteger(sourceDepth) || sourceDepth < 1) return null;
  return countPdfAppendedStartxrefs(output, sourceLength) === 1 ? sourceDepth + 1 : null;
}
