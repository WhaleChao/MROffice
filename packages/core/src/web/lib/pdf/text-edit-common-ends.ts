/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where an edit of `before` into `after` starts and ends: the length (in
 * UTF-16 units, as offsets are kept) of the text the two share at their
 * start and at their end, never inside a character. The paragraph rewrite
 * compares whole characters; the editor's part formats, a text box's runs
 * and the box format compared code units, so an emoji typed beside one that
 * shares its first half split the edit a unit apart from the writer.
 */
export function pdfTextEditCommonEnds(before: string, after: string): Readonly<{ prefix: number; suffix: number }> {
  const limit = Math.min(before.length, after.length);
  let prefix = 0;
  while (prefix < limit && before.charCodeAt(prefix) === after.charCodeAt(prefix)) prefix += 1;
  // Not after the first half of a pair the two do not both complete.
  if (prefix > 0 && isHighSurrogate(before.charCodeAt(prefix - 1)) &&
    (prefix >= before.length || prefix >= after.length || before.charCodeAt(prefix) !== after.charCodeAt(prefix))) prefix -= 1;
  let suffix = 0;
  while (suffix < before.length - prefix && suffix < after.length - prefix &&
    before.charCodeAt(before.length - 1 - suffix) === after.charCodeAt(after.length - 1 - suffix)) suffix += 1;
  // Not before the second half of a pair whose first halves differ.
  if (suffix > 0 && isLowSurrogate(before.charCodeAt(before.length - suffix)) &&
    (suffix >= before.length - prefix || suffix >= after.length - prefix ||
      before.charCodeAt(before.length - 1 - suffix) !== after.charCodeAt(after.length - 1 - suffix))) suffix -= 1;
  return Object.freeze({ prefix, suffix });
}

const isHighSurrogate = (code: number) => code >= 0xd800 && code <= 0xdbff;
const isLowSurrogate = (code: number) => code >= 0xdc00 && code <= 0xdfff;
