/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A symbolic TrueType program's own code table, for a simple font that
 * carries neither an /Encoding nor a ToUnicode.
 *
 * PDF 32000-1 9.6.6.4 says how such a font draws a code: through its (3,0)
 * `cmap` subtable, with 0x00 or 0xF0-0xF2 as the high byte, and without one
 * through its (1,0) subtable directly. The (1,0) subtable is the Mac Roman
 * one, so a code the font maps there is the Mac Roman character of that code,
 * and the reader reads a symbolic font without an encoding through
 * MacRomanEncoding exactly so. A code is taken only when the glyph the page
 * draws for it is that Mac Roman glyph: a (3,0) subtable that draws another
 * glyph, or none, for the code says nothing about its text, and a font with
 * no (1,0) subtable names no text at all -- both stay preserve-only.
 */

import { PDF_SIMPLE_FONT_BASE_ENCODINGS } from "./simple-font-encodings";
import { readSfntTableDirectory } from "./truetype-sparse-subset";

type GlyphLookup = (code: number) => number;

function readSubtable(view: DataView, start: number, end: number): GlyphLookup | null {
  if (start + 6 > end) return null;
  const format = view.getUint16(start);
  if (format === 0) {
    if (start + 6 + 256 > end) return null;
    return (code) => (code >= 0 && code < 256 ? view.getUint8(start + 6 + code) : 0);
  }
  if (format === 6) {
    if (start + 10 > end) return null;
    const firstCode = view.getUint16(start + 6);
    const count = view.getUint16(start + 8);
    if (start + 10 + count * 2 > end) return null;
    return (code) =>
      code >= firstCode && code < firstCode + count
        ? view.getUint16(start + 10 + (code - firstCode) * 2)
        : 0;
  }
  if (format === 4) {
    if (start + 14 > end) return null;
    const segments = view.getUint16(start + 6) / 2;
    if (!Number.isInteger(segments) || segments < 1) return null;
    const ends = start + 14;
    const starts = ends + segments * 2 + 2;
    const deltas = starts + segments * 2;
    const rangeOffsets = deltas + segments * 2;
    if (rangeOffsets + segments * 2 > end) return null;
    return (code) => {
      for (let index = 0; index < segments; index += 1) {
        if (view.getUint16(ends + index * 2) < code) continue;
        const first = view.getUint16(starts + index * 2);
        if (first > code) return 0;
        const delta = view.getInt16(deltas + index * 2);
        const rangeOffset = view.getUint16(rangeOffsets + index * 2);
        if (rangeOffset === 0) return (code + delta) & 0xffff;
        const address = rangeOffsets + index * 2 + rangeOffset + (code - first) * 2;
        if (address + 2 > end) return 0;
        const glyph = view.getUint16(address);
        return glyph === 0 ? 0 : (glyph + delta) & 0xffff;
      }
      return 0;
    };
  }
  return null;
}

/** Code -> text of a symbolic TrueType program read the way the page draws
 * it, or null when the program is no TrueType program, cannot be read, or
 * has no Mac Roman subtable. */
export function readTrueTypeSymbolicBuiltinEncoding(
  program: Uint8Array,
): ReadonlyMap<number, string> | null {
  const directory = readSfntTableDirectory(program);
  if (!directory) return null;
  const { view } = directory;
  const cmap = directory.tables.find((table) => table.tag === "cmap");
  const maxp = directory.tables.find((table) => table.tag === "maxp");
  if (!cmap || !maxp || maxp.length < 6 || cmap.length < 4) return null;
  const glyphCount = view.getUint16(maxp.offset + 4);
  const cmapEnd = cmap.offset + cmap.length;
  const records = view.getUint16(cmap.offset + 2);
  if (cmap.offset + 4 + records * 8 > cmapEnd) return null;
  const subtables = new Map<string, GlyphLookup | null>();
  for (let index = 0; index < records; index += 1) {
    const record = cmap.offset + 4 + index * 8;
    const key = `${view.getUint16(record)},${view.getUint16(record + 2)}`;
    if (key !== "1,0" && key !== "3,0") continue;
    // Two subtables for one platform and encoding do not say which one draws.
    if (subtables.has(key)) return null;
    const offset = cmap.offset + view.getUint32(record + 4);
    subtables.set(key, offset < cmapEnd ? readSubtable(view, offset, cmapEnd) : null);
  }
  const mac = subtables.get("1,0");
  const symbol = subtables.get("3,0");
  if (!mac || symbol === null) return null;
  const roman = PDF_SIMPLE_FONT_BASE_ENCODINGS.MacRomanEncoding as Readonly<Record<number, string>>;
  const decode = new Map<number, string>();
  for (let code = 0; code < 256; code += 1) {
    const character = roman[code];
    if (character === undefined) continue;
    const glyph = mac(code);
    if (glyph === 0 || glyph >= glyphCount) continue;
    if (symbol) {
      const drawn = new Set(
        [code, 0xf000 | code, 0xf100 | code, 0xf200 | code]
          .map((candidate) => symbol(candidate))
          .filter((candidate) => candidate !== 0),
      );
      if (drawn.size !== 1 || !drawn.has(glyph)) continue;
    }
    decode.set(code, character);
  }
  return decode;
}
