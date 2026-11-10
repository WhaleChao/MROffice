/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A glyph-id preserving sparse TrueType subset.
 *
 * The font extension installs one relocatable Type0 program so a page whose
 * own font cannot carry the replacement text stays natively editable. Carrying
 * the whole canonical program made every such document grow by its full size.
 * Dropping the outlines nothing encodes removes almost all of that, and
 * leaving every glyph id where it was keeps an installed program addressable
 * by the same CIDs -- so a later edit still reuses it instead of installing a
 * second one, and text already written against it keeps rendering.
 *
 * Every table except `glyf` and `loca` is copied byte for byte, including the
 * vertical metrics and layout tables: a subset must stay the same program with
 * fewer outlines, never a differently shaped font.
 */

const SFNT_TRUETYPE = 0x00010000;
const SFNT_TRUE = 0x74727565;
const MAX_TABLES = 64;
const MAX_GLYPHS = 65_535;
const MAX_COMPOSITE_DEPTH = 8;
const CHECKSUM_MAGIC = 0xb1b0afba;

type SfntTable = {
  readonly tag: string;
  readonly offset: number;
  readonly length: number;
};

type SfntDirectory = {
  readonly view: DataView;
  readonly tables: readonly SfntTable[];
};

/** The table directory of a TrueType program, or null when it is not one. */
export function readSfntTableDirectory(font: Uint8Array): SfntDirectory | null {
  return readTableDirectory(font);
}

function readTableDirectory(font: Uint8Array): SfntDirectory | null {
  if (font.length < 12 || font.length > 64 * 1024 * 1024) return null;
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const version = view.getUint32(0);
  if (version !== SFNT_TRUETYPE && version !== SFNT_TRUE) return null;
  const count = view.getUint16(4);
  const dataStart = 12 + count * 16;
  if (count < 1 || count > MAX_TABLES || dataStart > font.length) return null;
  const tables: SfntTable[] = [];
  for (let index = 0; index < count; index += 1) {
    const entry = 12 + index * 16;
    const tag = String.fromCharCode(
      font[entry],
      font[entry + 1],
      font[entry + 2],
      font[entry + 3],
    );
    const offset = view.getUint32(entry + 8);
    const length = view.getUint32(entry + 12);
    if (
      !/^[\x20-\x7e]{4}$/.test(tag) ||
      offset < dataStart ||
      length > font.length ||
      offset + length > font.length ||
      tables.some((table) => table.tag === tag)
    ) {
      return null;
    }
    tables.push(Object.freeze({ tag, offset, length }));
  }
  return { view, tables };
}

type GlyphIndex = {
  readonly numGlyphs: number;
  readonly longFormat: boolean;
  readonly offsets: Uint32Array;
  readonly glyf: SfntTable;
};

function readGlyphIndex(
  font: Uint8Array,
  directory: SfntDirectory,
): GlyphIndex | null {
  const byTag = new Map(directory.tables.map((table) => [table.tag, table]));
  const head = byTag.get("head");
  const maxp = byTag.get("maxp");
  const loca = byTag.get("loca");
  const glyf = byTag.get("glyf");
  if (!head || !maxp || !loca || !glyf || head.length < 54 || maxp.length < 6) {
    return null;
  }
  const { view } = directory;
  const numGlyphs = view.getUint16(maxp.offset + 4);
  const indexToLocFormat = view.getInt16(head.offset + 50);
  if (
    numGlyphs < 1 ||
    numGlyphs > MAX_GLYPHS ||
    (indexToLocFormat !== 0 && indexToLocFormat !== 1)
  ) {
    return null;
  }
  const longFormat = indexToLocFormat === 1;
  const entries = numGlyphs + 1;
  if (loca.length < entries * (longFormat ? 4 : 2)) return null;
  const offsets = new Uint32Array(entries);
  let previous = 0;
  for (let index = 0; index < entries; index += 1) {
    const value = longFormat
      ? view.getUint32(loca.offset + index * 4)
      : view.getUint16(loca.offset + index * 2) * 2;
    if (value < previous || value > glyf.length) return null;
    offsets[index] = value;
    previous = value;
  }
  return { numGlyphs, longFormat, offsets, glyf };
}

/** Add a composite glyph's components, so a kept outline never references a
 * dropped one. A malformed record refuses the subset rather than emitting a
 * program whose glyph would silently render as a blank. */
function addGlyphComponents(
  view: DataView,
  index: GlyphIndex,
  glyphId: number,
  kept: Set<number>,
  depth: number,
): boolean {
  if (depth > MAX_COMPOSITE_DEPTH) return false;
  const start = index.glyf.offset + index.offsets[glyphId];
  const end = index.glyf.offset + index.offsets[glyphId + 1];
  if (end - start === 0) return true;
  if (end - start < 10 || view.getInt16(start) >= 0) return true;
  let cursor = start + 10;
  for (;;) {
    if (cursor + 4 > end) return false;
    const flags = view.getUint16(cursor);
    const component = view.getUint16(cursor + 2);
    cursor += 4;
    cursor += (flags & 0x0001) !== 0 ? 4 : 2;
    if ((flags & 0x0008) !== 0) cursor += 2;
    else if ((flags & 0x0040) !== 0) cursor += 4;
    else if ((flags & 0x0080) !== 0) cursor += 8;
    if (cursor > end || component >= index.numGlyphs) return false;
    if (!kept.has(component)) {
      kept.add(component);
      if (!addGlyphComponents(view, index, component, kept, depth + 1)) {
        return false;
      }
    }
    if ((flags & 0x0020) === 0) return true;
  }
}

function tableChecksum(data: Uint8Array): number {
  let sum = 0;
  for (let index = 0; index < data.length; index += 4) {
    const word =
      ((data[index] ?? 0) << 24) |
      ((data[index + 1] ?? 0) << 16) |
      ((data[index + 2] ?? 0) << 8) |
      (data[index + 3] ?? 0);
    sum = (sum + word) >>> 0;
  }
  return sum;
}

/** A `cmap` table mapping exactly `entries` (code point -> glyph id): a
 * Windows BMP format 4 subtable and a full-repertoire format 12 one. Null when
 * a BMP subtable would not fit its 16-bit length. */
function buildCmapTable(entries: ReadonlyMap<number, number>): Uint8Array | null {
  const sorted = [...entries].sort((left, right) => left[0] - right[0]);
  const groups: { start: number; end: number; glyph: number }[] = [];
  for (const [codePoint, glyphId] of sorted) {
    const last = groups.at(-1);
    if (last && codePoint === last.end + 1 && glyphId === last.glyph + codePoint - last.start) {
      last.end = codePoint;
    } else {
      groups.push({ start: codePoint, end: codePoint, glyph: glyphId });
    }
  }
  const bmp = groups
    .filter((group) => group.start <= 0xfffe)
    .map((group) => ({ ...group, end: Math.min(group.end, 0xfffe) }));
  const segments = [...bmp, { start: 0xffff, end: 0xffff, glyph: 0 }];
  const segCount = segments.length;
  const format4Length = 16 + segCount * 8;
  if (format4Length > 0xffff) return null;
  const format12Length = 16 + groups.length * 12;
  const headerLength = 4 + 2 * 8;
  const output = new Uint8Array(headerLength + format4Length + format12Length);
  const view = new DataView(output.buffer);
  view.setUint16(0, 0);
  view.setUint16(2, 2);
  view.setUint16(4, 3);
  view.setUint16(6, 1);
  view.setUint32(8, headerLength);
  view.setUint16(12, 3);
  view.setUint16(14, 10);
  view.setUint32(16, headerLength + format4Length);
  let cursor = headerLength;
  const entrySelector = Math.floor(Math.log2(segCount));
  const searchRange = 2 * 2 ** entrySelector;
  view.setUint16(cursor, 4);
  view.setUint16(cursor + 2, format4Length);
  view.setUint16(cursor + 4, 0);
  view.setUint16(cursor + 6, segCount * 2);
  view.setUint16(cursor + 8, searchRange);
  view.setUint16(cursor + 10, entrySelector);
  view.setUint16(cursor + 12, segCount * 2 - searchRange);
  const ends = cursor + 14;
  const starts = ends + segCount * 2 + 2;
  const deltas = starts + segCount * 2;
  const rangeOffsets = deltas + segCount * 2;
  segments.forEach((segment, position) => {
    view.setUint16(ends + position * 2, segment.end);
    view.setUint16(starts + position * 2, segment.start);
    // The closing segment maps 0xFFFF to glyph 0: 0xFFFF + 1 wraps to 0.
    const delta = segment.start === 0xffff ? 1 : (segment.glyph - segment.start) & 0xffff;
    view.setUint16(deltas + position * 2, delta);
    view.setUint16(rangeOffsets + position * 2, 0);
  });
  cursor += format4Length;
  view.setUint16(cursor, 12);
  view.setUint16(cursor + 2, 0);
  view.setUint32(cursor + 4, format12Length);
  view.setUint32(cursor + 8, 0);
  view.setUint32(cursor + 12, groups.length);
  groups.forEach((group, position) => {
    const entry = cursor + 16 + position * 12;
    view.setUint32(entry, group.start);
    view.setUint32(entry + 4, group.end);
    view.setUint32(entry + 8, group.glyph);
  });
  return output;
}

/** Rewrite `font` carrying only the requested glyphs, at their own glyph ids.
 * Returns null when the program cannot be narrowed without changing it; the
 * caller then keeps the full program rather than shipping a damaged one.
 *
 * With `cmap`, the program's character map is narrowed to exactly those code
 * points too. Glyph ids never move, so every CID keeps drawing the same
 * outline; what shrinks is what the program claims to cover, and with it the
 * widths and ToUnicode entries an embedder derives from that claim. */
export function pdfNativeSparseTrueTypeSubset(
  font: Uint8Array,
  glyphIds: Iterable<number>,
  cmap?: ReadonlyMap<number, number>,
): Uint8Array | null {
  const directory = readTableDirectory(font);
  if (!directory) return null;
  const index = readGlyphIndex(font, directory);
  if (!index) return null;
  const { view } = directory;
  const kept = new Set<number>([0]);
  for (const glyphId of glyphIds) {
    if (
      !Number.isSafeInteger(glyphId) ||
      glyphId < 0 ||
      glyphId >= index.numGlyphs
    ) {
      return null;
    }
    kept.add(glyphId);
  }
  for (const glyphId of [...kept]) {
    if (!addGlyphComponents(view, index, glyphId, kept, 0)) return null;
  }

  const align = (value: number): number => (value + 3) & ~3;
  let glyfLength = 0;
  for (const glyphId of kept) {
    glyfLength += align(index.offsets[glyphId + 1] - index.offsets[glyphId]);
  }
  // A subset that is not smaller buys nothing and would look like tampering
  // to the independent server check, which requires it below the full size.
  if (glyfLength >= index.glyf.length) return null;
  const locaLength = (index.numGlyphs + 1) * (index.longFormat ? 4 : 2);
  const glyfData = new Uint8Array(glyfLength);
  const locaData = new Uint8Array(locaLength);
  const locaView = new DataView(locaData.buffer);
  let cursor = 0;
  for (let glyphId = 0; glyphId <= index.numGlyphs; glyphId += 1) {
    if (index.longFormat) locaView.setUint32(glyphId * 4, cursor);
    else if (cursor % 2 !== 0 || cursor / 2 > 0xffff) return null;
    else locaView.setUint16(glyphId * 2, cursor / 2);
    if (glyphId === index.numGlyphs) break;
    if (!kept.has(glyphId)) continue;
    const start = index.offsets[glyphId];
    const end = index.offsets[glyphId + 1];
    glyfData.set(
      font.subarray(index.glyf.offset + start, index.glyf.offset + end),
      cursor,
    );
    cursor += align(end - start);
  }

  let cmapData: Uint8Array | null = null;
  if (cmap) {
    for (const [codePoint, glyphId] of cmap) {
      if (
        !Number.isSafeInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff ||
        (codePoint >= 0xd800 && codePoint <= 0xdfff) || !kept.has(glyphId)
      ) {
        return null;
      }
    }
    cmapData = buildCmapTable(cmap);
    if (!cmapData || !directory.tables.some((table) => table.tag === "cmap")) return null;
  }

  const tables = [...directory.tables]
    .map((table) => {
      if (table.tag === "glyf") return { tag: table.tag, data: glyfData };
      if (table.tag === "loca") return { tag: table.tag, data: locaData };
      if (table.tag === "cmap" && cmapData) return { tag: table.tag, data: cmapData };
      const data = font.slice(table.offset, table.offset + table.length);
      // The file checksum below only closes over a head whose adjustment slot
      // is zero, exactly as the format defines it.
      if (table.tag === "head" && data.length >= 12) data.fill(0, 8, 12);
      return { tag: table.tag, data };
    })
    .sort((left, right) => (left.tag < right.tag ? -1 : left.tag > right.tag ? 1 : 0));

  const count = tables.length;
  const dataStart = 12 + count * 16;
  let total = dataStart;
  const placements = tables.map((table) => {
    const offset = total;
    total += align(table.data.length);
    return { ...table, offset };
  });
  const output = new Uint8Array(total);
  const outputView = new DataView(output.buffer);
  const entrySelector = Math.floor(Math.log2(count));
  const searchRange = 16 * 2 ** entrySelector;
  outputView.setUint32(0, view.getUint32(0));
  outputView.setUint16(4, count);
  outputView.setUint16(6, searchRange);
  outputView.setUint16(8, entrySelector);
  outputView.setUint16(10, count * 16 - searchRange);
  let headOffset: number | null = null;
  placements.forEach((table, position) => {
    const entry = 12 + position * 16;
    for (let byte = 0; byte < 4; byte += 1) {
      output[entry + byte] = table.tag.charCodeAt(byte);
    }
    outputView.setUint32(entry + 4, tableChecksum(table.data));
    outputView.setUint32(entry + 8, table.offset);
    outputView.setUint32(entry + 12, table.data.length);
    output.set(table.data, table.offset);
    if (table.tag === "head") headOffset = table.offset;
  });
  if (headOffset === null) return null;
  outputView.setUint32(
    headOffset + 8,
    (CHECKSUM_MAGIC - tableChecksum(output)) >>> 0,
  );
  return output;
}
