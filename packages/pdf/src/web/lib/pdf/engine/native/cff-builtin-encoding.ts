/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A bare CFF program's own encoding, for symbolic fonts with no ToUnicode.
 *
 * A Type 1 program states its encoding as PostScript text; a CFF one states
 * the same thing as three binary tables -- an Encoding mapping codes to glyph
 * indexes, a charset mapping glyph indexes to string ids, and a String INDEX
 * naming the ids the standard strings do not already name. Reading them is
 * what lets a symbolic CFF font's codes mean anything at all.
 *
 * Structure per Adobe TN #5176. Nothing here interprets a charstring: only
 * the tables that say which code draws which named glyph are read.
 *
 * Verified against fontTools over every CFF program in the local corpus: 1299
 * of 1308 code assignments identical, and the nine that differ are all code 0
 * of a format 0 encoding, which fontTools drops on a truthiness check while
 * the format's own code array assigns it like any other. The Python replay
 * mirrors this reader rather than fontTools so both sides read the format the
 * same way.
 */

import { CFF_STANDARD_STRINGS } from "./cff-standard-strings";

const MAX_CFF_BYTES = 8 << 20;
const MAX_GLYPHS = 65_535;
const STANDARD_ENCODING_OFFSET = 0;
const EXPERT_ENCODING_OFFSET = 1;

type Reader = Readonly<{ bytes: Uint8Array; view: DataView }>;

function u8(reader: Reader, at: number): number {
  if (at < 0 || at >= reader.bytes.length) throw new RangeError("cff");
  return reader.bytes[at];
}

function offsetAt(reader: Reader, at: number, size: number): number {
  let value = 0;
  for (let index = 0; index < size; index += 1) value = value * 256 + u8(reader, at + index);
  return value;
}

/** One CFF INDEX: its element ranges and the offset just past it. */
function readIndex(
  reader: Reader,
  at: number,
): Readonly<{ items: readonly Readonly<{ start: number; end: number }>[]; end: number }> {
  const count = (u8(reader, at) << 8) | u8(reader, at + 1);
  if (count === 0) return Object.freeze({ items: Object.freeze([]), end: at + 2 });
  const offSize = u8(reader, at + 2);
  if (offSize < 1 || offSize > 4) throw new RangeError("cff");
  const offsets: number[] = [];
  const table = at + 3;
  for (let index = 0; index <= count; index += 1) {
    offsets.push(offsetAt(reader, table + index * offSize, offSize));
  }
  const data = table + (count + 1) * offSize - 1;
  const items = offsets.slice(0, -1).map((start, index) => {
    const end = offsets[index + 1];
    if (start < 1 || end < start || data + end > reader.bytes.length) {
      throw new RangeError("cff");
    }
    return Object.freeze({ start: data + start, end: data + end });
  });
  return Object.freeze({ items: Object.freeze(items), end: data + offsets[count] });
}

/** Operands of one CFF DICT, keyed by operator (1200 + n for escaped ones). */
function readDict(reader: Reader, start: number, end: number): ReadonlyMap<number, readonly number[]> {
  const dict = new Map<number, readonly number[]>();
  let operands: number[] = [];
  let at = start;
  while (at < end) {
    const b0 = u8(reader, at);
    if (b0 <= 21) {
      const operator = b0 === 12 ? 1200 + u8(reader, at + 1) : b0;
      dict.set(operator, Object.freeze(operands));
      operands = [];
      at += b0 === 12 ? 2 : 1;
    } else if (b0 === 28) {
      operands.push(reader.view.getInt16(at + 1));
      at += 3;
    } else if (b0 === 29) {
      operands.push(reader.view.getInt32(at + 1));
      at += 5;
    } else if (b0 === 30) {
      // A real number; its value never addresses a table, so skip its nibbles.
      at += 1;
      for (;;) {
        const pair = u8(reader, at);
        at += 1;
        if ((pair & 0x0f) === 0x0f || (pair >> 4) === 0x0f) break;
      }
      operands.push(Number.NaN);
    } else if (b0 >= 32 && b0 <= 246) {
      operands.push(b0 - 139);
      at += 1;
    } else if (b0 >= 247 && b0 <= 250) {
      operands.push((b0 - 247) * 256 + u8(reader, at + 1) + 108);
      at += 2;
    } else if (b0 >= 251 && b0 <= 254) {
      operands.push(-(b0 - 251) * 256 - u8(reader, at + 1) - 108);
      at += 2;
    } else {
      throw new RangeError("cff");
    }
    if (operands.length > 48) throw new RangeError("cff");
  }
  return dict;
}

function charsetSids(reader: Reader, at: number, glyphs: number): readonly number[] {
  // Offsets 0-2 name the predefined charsets, whose SIDs are their own ids.
  if (at <= 2) return Object.freeze(Array.from({ length: glyphs }, (_, id) => id));
  const format = u8(reader, at);
  const sids = [0];
  if (format === 0) {
    for (let index = 1; index < glyphs; index += 1) {
      sids.push((u8(reader, at + 1 + (index - 1) * 2) << 8) | u8(reader, at + 2 + (index - 1) * 2));
    }
    return Object.freeze(sids);
  }
  if (format !== 1 && format !== 2) throw new RangeError("cff");
  let cursor = at + 1;
  while (sids.length < glyphs) {
    const first = (u8(reader, cursor) << 8) | u8(reader, cursor + 1);
    const left = format === 1
      ? u8(reader, cursor + 2)
      : (u8(reader, cursor + 2) << 8) | u8(reader, cursor + 3);
    cursor += format === 1 ? 3 : 4;
    for (let index = 0; index <= left && sids.length < glyphs; index += 1) {
      sids.push(first + index);
    }
  }
  return Object.freeze(sids);
}

function encodingCodes(reader: Reader, at: number, glyphs: number): ReadonlyMap<number, number> {
  const codes = new Map<number, number>();
  const format = u8(reader, at);
  if ((format & 0x7f) === 0) {
    const count = u8(reader, at + 1);
    for (let index = 1; index <= count && index < glyphs; index += 1) {
      codes.set(u8(reader, at + 1 + index), index);
    }
  } else if ((format & 0x7f) === 1) {
    const ranges = u8(reader, at + 1);
    let glyph = 1;
    for (let index = 0; index < ranges; index += 1) {
      const first = u8(reader, at + 2 + index * 2);
      const left = u8(reader, at + 3 + index * 2);
      for (let step = 0; step <= left && glyph < glyphs; step += 1) {
        codes.set(first + step, glyph);
        glyph += 1;
      }
    }
  } else {
    throw new RangeError("cff");
  }
  if ((format & 0x80) !== 0) {
    // Supplements map extra codes to glyphs already named by a SID.
    const base = (format & 0x7f) === 0 ? at + 2 + u8(reader, at + 1) : at + 2 + u8(reader, at + 1) * 2;
    const supplements = u8(reader, base);
    for (let index = 0; index < supplements; index += 1) {
      const code = u8(reader, base + 1 + index * 3);
      const sid = (u8(reader, base + 2 + index * 3) << 8) | u8(reader, base + 3 + index * 3);
      codes.set(code, -sid);
    }
  }
  return codes;
}

/** Codes a bare CFF program's own encoding assigns, as glyph names. */
export function readCffBuiltinEncodingNames(
  program: Uint8Array,
): ReadonlyMap<number, string> | "standard" | null {
  if (program.length < 4 || program.length > MAX_CFF_BYTES) return null;
  if (program[0] !== 1 && program[0] !== 2) return null;
  try {
    const reader: Reader = Object.freeze({
      bytes: program,
      view: new DataView(program.buffer, program.byteOffset, program.byteLength),
    });
    const headerSize = u8(reader, 2);
    const names = readIndex(reader, headerSize);
    const topDicts = readIndex(reader, names.end);
    const strings = readIndex(reader, topDicts.end);
    if (topDicts.items.length !== 1) return null;
    const top = readDict(reader, topDicts.items[0].start, topDicts.items[0].end);
    const charStringsOffset = top.get(17)?.[0];
    if (charStringsOffset === undefined || !Number.isSafeInteger(charStringsOffset)) return null;
    // A CID-keyed font has no encoding of its own; its codes are CIDs.
    if (top.has(1230)) return null;
    const charStrings = readIndex(reader, charStringsOffset);
    const glyphs = charStrings.items.length;
    if (glyphs < 1 || glyphs > MAX_GLYPHS) return null;
    const encodingOffset = top.get(16)?.[0] ?? STANDARD_ENCODING_OFFSET;
    // A program may name a predefined encoding instead of carrying one. The
    // standard one is a table this editor already has, and saying so is just
    // as definite as spelling the codes out.
    if (encodingOffset === STANDARD_ENCODING_OFFSET) return "standard";
    if (encodingOffset === EXPERT_ENCODING_OFFSET) return null;
    const sids = charsetSids(reader, top.get(15)?.[0] ?? 0, glyphs);
    const nameOf = (sid: number): string | null => {
      if (sid < CFF_STANDARD_STRINGS.length) return CFF_STANDARD_STRINGS[sid];
      const item = strings.items[sid - CFF_STANDARD_STRINGS.length];
      if (!item) return null;
      return String.fromCharCode(...program.subarray(item.start, item.end));
    };
    const result = new Map<number, string>();
    for (const [code, glyph] of encodingCodes(reader, encodingOffset, glyphs)) {
      if (code < 0 || code > 255) continue;
      const sid = glyph < 0 ? -glyph : sids[glyph];
      if (sid === undefined) continue;
      const name = nameOf(sid);
      if (name && name !== ".notdef") result.set(code, name);
    }
    return result.size ? result : null;
  } catch {
    return null;
  }
}

/**
 * The CIDs a CID-keyed CFF program holds a glyph for: its charset, which names
 * each glyph's CID. Bare or in an OpenType wrapper. An Identity font with no
 * ToUnicode reads its text through its ordering, which names every character
 * of the collection; a subset draws only these (a CID outside them draws
 * nothing). Null for a program that is not CID-keyed or cannot be read.
 */
export function readCffCidCoverage(program: Uint8Array): ReadonlySet<number> | null {
  const cff = cffTableOf(program);
  if (!cff || cff.length < 4 || cff.length > MAX_CFF_BYTES || (cff[0] !== 1 && cff[0] !== 2)) return null;
  try {
    const reader: Reader = Object.freeze({ bytes: cff, view: new DataView(cff.buffer, cff.byteOffset, cff.byteLength) });
    const names = readIndex(reader, u8(reader, 2));
    const topDicts = readIndex(reader, names.end);
    if (topDicts.items.length !== 1) return null;
    const top = readDict(reader, topDicts.items[0].start, topDicts.items[0].end);
    const charStringsOffset = top.get(17)?.[0];
    if (!top.has(1230) || charStringsOffset === undefined || !Number.isSafeInteger(charStringsOffset)) return null;
    const glyphs = readIndex(reader, charStringsOffset).items.length;
    if (glyphs < 1 || glyphs > MAX_GLYPHS) return null;
    const charset = top.get(15)?.[0] ?? 0;
    // The predefined charsets are for name-keyed fonts: a CID-keyed one has its own.
    if (charset <= 2) return null;
    return Object.freeze(new Set(charsetSids(reader, charset, glyphs).filter((cid) => cid !== 0)));
  } catch {
    return null;
  }
}

/** The CFF table of an OpenType wrapper, or the program itself. */
function cffTableOf(program: Uint8Array): Uint8Array | null {
  if (!(program[0] === 0x4f && program[1] === 0x54 && program[2] === 0x54 && program[3] === 0x4f)) return program;
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  if (program.length < 12) return null;
  const tables = view.getUint16(4);
  for (let index = 0; index < tables && 12 + (index + 1) * 16 <= program.length; index += 1) {
    const entry = 12 + index * 16;
    if (view.getUint32(entry) !== 0x43464620) continue;
    const start = view.getUint32(entry + 8), length = view.getUint32(entry + 12);
    return start + length <= program.length ? program.subarray(start, start + length) : null;
  }
  return null;
}
