/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A TrueType program read glyph by glyph: the open faces MRAI serves are
 * WOFF2 slices (lib/office-fonts), and the PDF editor embeds subsets of them
 * (sfnt-subset.ts). This reads what that needs -- the glyph outlines exactly
 * as the font stores them (points, on-curve flags, contours, components),
 * the metrics, the character map and the variation tables -- from a WOFF2
 * or a plain TrueType file.
 *
 * WOFF2 is decoded per the W3C WOFF2 specification: the tables are one
 * Brotli stream (decompressed by fontkit, which already carries a decoder),
 * the `glyf`/`loca` pair may be stored transformed (section 5.1) and `hmtx`
 * may omit side bearings (section 5.4). The server reads the same slice with
 * fontTools and proves every embedded glyph against it, so nothing here is
 * trusted beyond producing a candidate.
 */

import { readCffOutlines, type CffOutlines } from "./cff-outlines";

export type PdfNativeSfntErrorCode =
  | "invalid-font"
  | "unsupported-outlines"
  | "unsupported-variation"
  | "missing-glyph";

export class PdfNativeSfntError extends Error {
  constructor(readonly code: PdfNativeSfntErrorCode, message?: string) {
    super(message ?? `PDF font program refused: ${code}.`);
    this.name = "PdfNativeSfntError";
  }
}

export function sfntFail(code: PdfNativeSfntErrorCode, message?: string): never {
  throw new PdfNativeSfntError(code, message);
}

/** Big-endian reads over one table. */
export class SfntReader {
  private readonly view: DataView;
  offset: number;

  constructor(readonly bytes: Uint8Array, offset = 0) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.offset = offset;
  }

  private need(count: number): void {
    if (this.offset < 0 || this.offset + count > this.bytes.length) sfntFail("invalid-font", "Font table is truncated.");
  }

  u8(): number { this.need(1); return this.view.getUint8(this.offset++); }
  i8(): number { this.need(1); return this.view.getInt8(this.offset++); }
  u16(): number { this.need(2); const value = this.view.getUint16(this.offset); this.offset += 2; return value; }
  i16(): number { this.need(2); const value = this.view.getInt16(this.offset); this.offset += 2; return value; }
  u32(): number { this.need(4); const value = this.view.getUint32(this.offset); this.offset += 4; return value; }
  i32(): number { this.need(4); const value = this.view.getInt32(this.offset); this.offset += 4; return value; }
  /** F2DOT14 as a float. */
  f2dot14(): number { return this.i16() / 16384; }
  /** 16.16 Fixed as a float. */
  fixed(): number { return this.i32() / 65536; }
  tag(): string { this.need(4); const text = String.fromCharCode(...this.bytes.subarray(this.offset, this.offset + 4)); this.offset += 4; return text; }
  take(count: number): Uint8Array { this.need(count); const out = this.bytes.subarray(this.offset, this.offset + count); this.offset += count; return out; }
  /** WOFF2 255UInt16. */
  u255(): number {
    const code = this.u8();
    if (code === 253) return this.u16();
    if (code === 255) return this.u8() + 253;
    if (code === 254) return this.u8() + 506;
    return code;
  }
  at(offset: number): SfntReader { return new SfntReader(this.bytes, offset); }
}

/** One component of a composite glyph, as stored. */
export interface SfntComponent {
  readonly flags: number;
  readonly glyphIndex: number;
  /** x/y offsets when ARGS_ARE_XY_VALUES, else the two point numbers. */
  readonly arg1: number;
  readonly arg2: number;
  /** The scale/2x2 transform, raw F2DOT14 values (0, 1 or 4 of them). */
  readonly transform: readonly number[];
}

export interface SfntSimpleGlyph {
  readonly kind: "simple";
  /** Last point index of each contour. */
  readonly endPoints: readonly number[];
  readonly onCurve: readonly boolean[];
  readonly xs: readonly number[];
  readonly ys: readonly number[];
  /** OVERLAP_SIMPLE on the first point's flags. */
  readonly overlap: boolean;
  readonly bbox: readonly [number, number, number, number];
}

export interface SfntCompositeGlyph {
  readonly kind: "composite";
  readonly components: readonly SfntComponent[];
  readonly bbox: readonly [number, number, number, number];
}

export type SfntGlyph = SfntSimpleGlyph | SfntCompositeGlyph;

export const COMPOSITE_ARG_1_AND_2_ARE_WORDS = 0x0001;
export const COMPOSITE_ARGS_ARE_XY_VALUES = 0x0002;
export const COMPOSITE_ROUND_XY_TO_GRID = 0x0004;
export const COMPOSITE_WE_HAVE_A_SCALE = 0x0008;
export const COMPOSITE_MORE_COMPONENTS = 0x0020;
export const COMPOSITE_WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
export const COMPOSITE_WE_HAVE_A_TWO_BY_TWO = 0x0080;
export const COMPOSITE_WE_HAVE_INSTRUCTIONS = 0x0100;
export const COMPOSITE_USE_MY_METRICS = 0x0200;
export const COMPOSITE_OVERLAP_COMPOUND = 0x0400;

export interface SfntProgram {
  /** Raw sfnt tables other than glyf/loca/hmtx (already in sfnt form). */
  readonly tables: ReadonlyMap<string, Uint8Array>;
  readonly numGlyphs: number;
  readonly unitsPerEm: number;
  readonly ascent: number;
  readonly descent: number;
  /** Advance and left side bearing of every glyph. */
  readonly advances: readonly number[];
  readonly leftSideBearings: readonly number[];
  /** null for an empty glyph (no contours). */
  glyph(glyphId: number): SfntGlyph | null;
  /** The best Unicode map (fontTools' getBestCmap order). */
  readonly cmap: ReadonlyMap<number, number>;
  readonly postscriptName: string | null;
}

// --- WOFF2 -----------------------------------------------------------------

interface Woff2Entry {
  readonly tag: string;
  readonly offset: number;
  readonly length: number;
  readonly transformLength?: number;
  readonly transformed: boolean;
  readonly transformVersion: number;
}

interface FontkitWoff2 {
  _decompress(): void;
  stream: { buffer: Uint8Array };
  directory: { tables: Record<string, Woff2Entry>; flavor: number };
}

async function loadFontkit(): Promise<{ create(bytes: Uint8Array): unknown }> {
  const fontkitModule: unknown = await import("@pdf-lib/fontkit");
  return ((fontkitModule as { default?: unknown }).default ?? fontkitModule) as { create(bytes: Uint8Array): unknown };
}

/** The decompressed table data of a WOFF2 file, table by table. fontkit
 * decodes the header and inflates the Brotli stream; it lays the tables end
 * to end in directory order, each at its (transformed) length. */
async function woff2Tables(bytes: Uint8Array): Promise<Readonly<{ flavor: number; tables: Map<string, Woff2Entry & { data: Uint8Array }> }>> {
  const fontkit = await loadFontkit();
  let font: FontkitWoff2;
  try {
    font = fontkit.create(bytes) as FontkitWoff2;
    font._decompress();
  } catch {
    return sfntFail("invalid-font", "The WOFF2 program could not be decompressed.");
  }
  const data = font.stream?.buffer;
  const directory = font.directory?.tables;
  if (!(data instanceof Uint8Array) || !directory) sfntFail("invalid-font");
  const tables = new Map<string, Woff2Entry & { data: Uint8Array }>();
  for (const [tag, entry] of Object.entries(directory)) {
    const length = entry.transformed ? entry.transformLength : entry.length;
    if (!Number.isSafeInteger(entry.offset) || !Number.isSafeInteger(length) || entry.offset + (length as number) > data.length) {
      sfntFail("invalid-font", "WOFF2 table directory is inconsistent.");
    }
    tables.set(tag, { ...entry, data: data.subarray(entry.offset, entry.offset + (length as number)) });
  }
  return { flavor: font.directory.flavor, tables };
}

const ON_CURVE = 0x01;
const X_SHORT = 0x02;
const Y_SHORT = 0x04;
const REPEAT = 0x08;
const X_SAME = 0x10;
const Y_SAME = 0x20;
const OVERLAP_SIMPLE = 0x40;

function boundsOf(xs: readonly number[], ys: readonly number[]): [number, number, number, number] {
  if (xs.length === 0) return [0, 0, 0, 0];
  let xMin = Infinity, yMin = Infinity, xMax = -Infinity, yMax = -Infinity;
  for (let index = 0; index < xs.length; index += 1) {
    xMin = Math.min(xMin, xs[index]); xMax = Math.max(xMax, xs[index]);
    yMin = Math.min(yMin, ys[index]); yMax = Math.max(yMax, ys[index]);
  }
  return [xMin, yMin, xMax, yMax];
}

function withSign(flag: number, value: number): number {
  return flag & 1 ? value : -value;
}

/** The composite records of one glyph (shared by the plain and the WOFF2
 * transformed glyf readers): every component until MORE_COMPONENTS stops. */
function readComponents(reader: SfntReader): { components: SfntComponent[]; instructions: boolean } {
  const components: SfntComponent[] = [];
  let flags: number;
  let instructions = false;
  do {
    flags = reader.u16();
    const glyphIndex = reader.u16();
    const words = (flags & COMPOSITE_ARG_1_AND_2_ARE_WORDS) !== 0;
    const signed = (flags & COMPOSITE_ARGS_ARE_XY_VALUES) !== 0;
    const arg1 = words ? (signed ? reader.i16() : reader.u16()) : (signed ? reader.i8() : reader.u8());
    const arg2 = words ? (signed ? reader.i16() : reader.u16()) : (signed ? reader.i8() : reader.u8());
    const transform: number[] = [];
    if (flags & COMPOSITE_WE_HAVE_A_SCALE) transform.push(reader.i16());
    else if (flags & COMPOSITE_WE_HAVE_AN_X_AND_Y_SCALE) transform.push(reader.i16(), reader.i16());
    else if (flags & COMPOSITE_WE_HAVE_A_TWO_BY_TWO) transform.push(reader.i16(), reader.i16(), reader.i16(), reader.i16());
    if (flags & COMPOSITE_WE_HAVE_INSTRUCTIONS) instructions = true;
    components.push(Object.freeze({ flags, glyphIndex, arg1, arg2, transform: Object.freeze(transform) }));
    if (components.length > 1_024) sfntFail("invalid-font", "Composite glyph has too many components.");
  } while (flags & COMPOSITE_MORE_COMPONENTS);
  return { components, instructions };
}

/** The `glyf` table reconstructed from its WOFF2 transform (section 5.1). */
function decodeTransformedGlyf(data: Uint8Array): { glyphs: (SfntGlyph | null)[]; xMins: number[] } {
  const header = new SfntReader(data);
  header.u16(); // reserved
  const optionFlags = header.u16();
  const numGlyphs = header.u16();
  header.u16(); // indexFormat
  const sizes = [header.u32(), header.u32(), header.u32(), header.u32(), header.u32(), header.u32(), header.u32()];
  let cursor = header.offset;
  const stream = (size: number) => {
    const out = new SfntReader(data.subarray(cursor, cursor + size));
    if (cursor + size > data.length) sfntFail("invalid-font", "WOFF2 glyf stream is truncated.");
    cursor += size;
    return out;
  };
  const nContours = stream(sizes[0]);
  const nPoints = stream(sizes[1]);
  const flags = stream(sizes[2]);
  const glyphStream = stream(sizes[3]);
  const composites = stream(sizes[4]);
  const bboxStream = stream(sizes[5]);
  stream(sizes[6]); // instructions: not carried into an embedded subset
  const overlapBitmap = optionFlags & 1 ? data.subarray(cursor, cursor + ((numGlyphs + 7) >> 3)) : null;
  const bboxBitmapLength = 4 * ((numGlyphs + 31) >> 5);
  const bboxBitmap = bboxStream.take(bboxBitmapLength);
  const glyphs: (SfntGlyph | null)[] = [];
  const xMins: number[] = [];
  for (let glyphId = 0; glyphId < numGlyphs; glyphId += 1) {
    const explicitBox = (bboxBitmap[glyphId >> 3] & (0x80 >> (glyphId & 7))) !== 0;
    const contours = nContours.i16();
    if (contours === 0) {
      if (explicitBox) sfntFail("invalid-font", "An empty WOFF2 glyph names a bounding box.");
      glyphs.push(null);
      xMins.push(0);
      continue;
    }
    if (contours > 0) {
      const endPoints: number[] = [];
      let total = 0;
      for (let index = 0; index < contours; index += 1) {
        total += nPoints.u255();
        endPoints.push(total - 1);
      }
      const onCurve: boolean[] = [];
      const xs: number[] = [];
      const ys: number[] = [];
      let x = 0, y = 0;
      for (let index = 0; index < total; index += 1) {
        let flag = flags.u8();
        onCurve.push((flag >> 7) === 0);
        flag &= 0x7f;
        let dx: number, dy: number;
        if (flag < 10) {
          dx = 0;
          dy = withSign(flag, ((flag & 14) << 7) + glyphStream.u8());
        } else if (flag < 20) {
          dx = withSign(flag, (((flag - 10) & 14) << 7) + glyphStream.u8());
          dy = 0;
        } else if (flag < 84) {
          const b0 = flag - 20;
          const b1 = glyphStream.u8();
          dx = withSign(flag, 1 + (b0 & 0x30) + (b1 >> 4));
          dy = withSign(flag >> 1, 1 + ((b0 & 0x0c) << 2) + (b1 & 0x0f));
        } else if (flag < 120) {
          const b0 = flag - 84;
          dx = withSign(flag, 1 + (Math.floor(b0 / 12) << 8) + glyphStream.u8());
          dy = withSign(flag >> 1, 1 + (((b0 % 12) >> 2) << 8) + glyphStream.u8());
        } else if (flag < 124) {
          const b1 = glyphStream.u8();
          const b2 = glyphStream.u8();
          const b3 = glyphStream.u8();
          dx = withSign(flag, (b1 << 4) + (b2 >> 4));
          dy = withSign(flag >> 1, ((b2 & 0x0f) << 8) + b3);
        } else {
          dx = withSign(flag, glyphStream.u16());
          dy = withSign(flag >> 1, glyphStream.u16());
        }
        x += dx;
        y += dy;
        xs.push(x);
        ys.push(y);
      }
      glyphStream.u255(); // instruction length; the bytes are in their own stream
      const bbox = explicitBox
        ? ([bboxStream.i16(), bboxStream.i16(), bboxStream.i16(), bboxStream.i16()] as [number, number, number, number])
        : boundsOf(xs, ys);
      const overlap = overlapBitmap !== null && (overlapBitmap[glyphId >> 3] & (0x80 >> (glyphId & 7))) !== 0;
      glyphs.push(Object.freeze({ kind: "simple" as const, endPoints, onCurve, xs, ys, overlap, bbox }));
      xMins.push(bbox[0]);
      continue;
    }
    if (contours !== -1) sfntFail("invalid-font", "WOFF2 glyph contour count is invalid.");
    if (!explicitBox) sfntFail("invalid-font", "A WOFF2 composite glyph has no bounding box.");
    const { components, instructions } = readComponents(composites);
    if (instructions) glyphStream.u255();
    const bbox = [bboxStream.i16(), bboxStream.i16(), bboxStream.i16(), bboxStream.i16()] as [number, number, number, number];
    glyphs.push(Object.freeze({ kind: "composite" as const, components, bbox }));
    xMins.push(bbox[0]);
  }
  return { glyphs, xMins };
}

/** One glyph of a plain `glyf` table. */
function decodeGlyfRecord(record: Uint8Array): SfntGlyph | null {
  if (record.length === 0) return null;
  const reader = new SfntReader(record);
  const contours = reader.i16();
  const bbox = [reader.i16(), reader.i16(), reader.i16(), reader.i16()] as [number, number, number, number];
  if (contours < 0) {
    return Object.freeze({ kind: "composite" as const, components: readComponents(reader).components, bbox });
  }
  if (contours === 0) return null;
  const endPoints: number[] = [];
  for (let index = 0; index < contours; index += 1) endPoints.push(reader.u16());
  const total = endPoints[endPoints.length - 1] + 1;
  reader.take(reader.u16()); // instructions
  const pointFlags: number[] = [];
  while (pointFlags.length < total) {
    const flag = reader.u8();
    pointFlags.push(flag);
    if (flag & REPEAT) {
      const count = reader.u8();
      for (let repeat = 0; repeat < count; repeat += 1) pointFlags.push(flag);
    }
  }
  if (pointFlags.length !== total) sfntFail("invalid-font", "Glyph flags overrun its points.");
  const coordinates = (short: number, same: number): number[] => {
    const values: number[] = [];
    let value = 0;
    for (const flag of pointFlags) {
      if (flag & short) value += flag & same ? reader.u8() : -reader.u8();
      else if (!(flag & same)) value += reader.i16();
      values.push(value);
    }
    return values;
  };
  const xs = coordinates(X_SHORT, X_SAME);
  const ys = coordinates(Y_SHORT, Y_SAME);
  return Object.freeze({
    kind: "simple" as const,
    endPoints,
    onCurve: pointFlags.map((flag) => (flag & ON_CURVE) !== 0),
    xs,
    ys,
    overlap: (pointFlags[0] & OVERLAP_SIMPLE) !== 0,
    bbox,
  });
}

function decodePlainGlyf(glyf: Uint8Array, loca: Uint8Array, numGlyphs: number, longOffsets: boolean): (SfntGlyph | null)[] {
  const offsets = new SfntReader(loca);
  const at: number[] = [];
  for (let index = 0; index <= numGlyphs; index += 1) at.push(longOffsets ? offsets.u32() : offsets.u16() * 2);
  const glyphs: (SfntGlyph | null)[] = [];
  for (let index = 0; index < numGlyphs; index += 1) {
    if (at[index + 1] < at[index] || at[index + 1] > glyf.length) sfntFail("invalid-font", "loca points outside glyf.");
    glyphs.push(decodeGlyfRecord(glyf.subarray(at[index], at[index + 1])));
  }
  return glyphs;
}

/** hmtx, reconstructed from its WOFF2 transform when it is stored so. */
function decodeHmtx(
  data: Uint8Array,
  transformed: boolean,
  numberOfHMetrics: number,
  numGlyphs: number,
  xMins: readonly number[],
): { advances: number[]; leftSideBearings: number[] } {
  const reader = new SfntReader(data);
  const advances: number[] = [];
  const leftSideBearings: number[] = [];
  if (!transformed) {
    for (let index = 0; index < numberOfHMetrics; index += 1) {
      advances.push(reader.u16());
      leftSideBearings.push(reader.i16());
    }
    for (let index = numberOfHMetrics; index < numGlyphs; index += 1) {
      advances.push(advances[numberOfHMetrics - 1]);
      leftSideBearings.push(reader.i16());
    }
    return { advances, leftSideBearings };
  }
  const flags = reader.u8();
  for (let index = 0; index < numberOfHMetrics; index += 1) advances.push(reader.u16());
  for (let index = 0; index < numberOfHMetrics; index += 1) {
    leftSideBearings.push(flags & 1 ? xMins[index] : reader.i16());
  }
  for (let index = numberOfHMetrics; index < numGlyphs; index += 1) {
    advances.push(advances[numberOfHMetrics - 1]);
    leftSideBearings.push(flags & 2 ? xMins[index] : reader.i16());
  }
  return { advances, leftSideBearings };
}

// --- cmap ------------------------------------------------------------------

function readCmapSubtable(table: Uint8Array, offset: number): Map<number, number> | null {
  const reader = new SfntReader(table, offset);
  const format = reader.u16();
  const map = new Map<number, number>();
  if (format === 4) {
    reader.u16(); reader.u16();
    const segments = reader.u16() / 2;
    reader.offset += 6;
    const base = reader.offset;
    const ends = new SfntReader(table, base);
    const starts = new SfntReader(table, base + segments * 2 + 2);
    const deltas = new SfntReader(table, base + segments * 4 + 2);
    const rangeOffsetsAt = base + segments * 6 + 2;
    const rangeOffsets = new SfntReader(table, rangeOffsetsAt);
    for (let segment = 0; segment < segments; segment += 1) {
      const end = ends.u16();
      const start = starts.u16();
      const delta = deltas.i16();
      const rangeOffset = rangeOffsets.u16();
      for (let code = start; code <= end && code !== 0xffff; code += 1) {
        let glyph: number;
        if (rangeOffset === 0) glyph = (code + delta) & 0xffff;
        else {
          const at = rangeOffsetsAt + segment * 2 + rangeOffset + (code - start) * 2;
          glyph = new SfntReader(table, at).u16();
          if (glyph !== 0) glyph = (glyph + delta) & 0xffff;
        }
        if (glyph !== 0) map.set(code, glyph);
      }
    }
    return map;
  }
  if (format === 12) {
    reader.u16(); reader.u32(); reader.u32();
    const groups = reader.u32();
    for (let group = 0; group < groups; group += 1) {
      const start = reader.u32();
      const end = reader.u32();
      const glyph = reader.u32();
      if (end < start || end - start > 0x10ffff) sfntFail("invalid-font", "cmap group is invalid.");
      for (let code = start; code <= end; code += 1) if (glyph + code - start !== 0) map.set(code, glyph + code - start);
    }
    return map;
  }
  if (format === 0) {
    reader.u16(); reader.u16();
    for (let code = 0; code < 256; code += 1) {
      const glyph = reader.u8();
      if (glyph !== 0) map.set(code, glyph);
    }
    return map;
  }
  if (format === 6) {
    reader.u16(); reader.u16();
    const first = reader.u16();
    const count = reader.u16();
    for (let index = 0; index < count; index += 1) {
      const glyph = reader.u16();
      if (glyph !== 0) map.set(first + index, glyph);
    }
    return map;
  }
  return null;
}

/** fontTools' getBestCmap preference, so both sides read one map. */
const CMAP_PREFERENCES: readonly (readonly [number, number])[] = [[3, 10], [0, 6], [0, 4], [3, 1], [0, 3], [0, 2], [0, 1], [0, 0]];

function readBestCmap(table: Uint8Array | undefined): Map<number, number> {
  if (!table) return new Map();
  const reader = new SfntReader(table);
  reader.u16();
  const count = reader.u16();
  const records = new Map<string, number>();
  for (let index = 0; index < count; index += 1) {
    const platform = reader.u16();
    const encoding = reader.u16();
    const offset = reader.u32();
    const key = `${platform}:${encoding}`;
    if (!records.has(key)) records.set(key, offset);
  }
  for (const [platform, encoding] of CMAP_PREFERENCES) {
    const offset = records.get(`${platform}:${encoding}`);
    if (offset === undefined) continue;
    const map = readCmapSubtable(table, offset);
    if (map) return map;
  }
  return new Map();
}

function postscriptNameOf(table: Uint8Array | undefined): string | null {
  if (!table) return null;
  const reader = new SfntReader(table);
  reader.u16();
  const count = reader.u16();
  const storage = reader.u16();
  for (let index = 0; index < count; index += 1) {
    const platform = reader.u16();
    const encoding = reader.u16();
    reader.u16();
    const nameId = reader.u16();
    const length = reader.u16();
    const offset = reader.u16();
    if (nameId !== 6) continue;
    const raw = table.subarray(storage + offset, storage + offset + length);
    if (platform === 3 || platform === 0) {
      let text = "";
      for (let at = 0; at + 1 < raw.length; at += 2) text += String.fromCharCode((raw[at] << 8) | raw[at + 1]);
      return text;
    }
    if (platform === 1 && encoding === 0) return String.fromCharCode(...raw);
  }
  return null;
}

// --- program ---------------------------------------------------------------

function plainTables(bytes: Uint8Array): { flavor: number; tables: Map<string, Uint8Array> } {
  const reader = new SfntReader(bytes);
  const flavor = reader.u32();
  const count = reader.u16();
  reader.offset += 6;
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < count; index += 1) {
    const tag = reader.tag();
    reader.u32();
    const offset = reader.u32();
    const length = reader.u32();
    if (offset + length > bytes.length) sfntFail("invalid-font", "sfnt table is outside the file.");
    tables.set(tag, bytes.subarray(offset, offset + length));
  }
  return { flavor, tables };
}

/** Read a WOFF2 or TrueType program. CFF outlines are refused: the embedded
 * subset is a TrueType program (CIDFontType2), which they cannot become
 * without conversion. */
export async function readSfntProgram(bytes: Uint8Array): Promise<SfntProgram> {
  if (bytes.length < 12) sfntFail("invalid-font");
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  let tables: Map<string, Uint8Array>;
  let glyphs: (SfntGlyph | null)[];
  let hmtx: { advances: number[]; leftSideBearings: number[] };
  const tableOf = (map: Map<string, Uint8Array>, tag: string) => map.get(tag) ?? sfntFail("invalid-font", `The program has no ${tag} table.`);
  if (signature === "wOF2") {
    const woff2 = await woff2Tables(bytes);
    if (woff2.tables.has("CFF ") || woff2.tables.has("CFF2")) sfntFail("unsupported-outlines", "CFF outlines cannot be embedded as TrueType.");
    tables = new Map([...woff2.tables].filter(([tag]) => tag !== "glyf" && tag !== "loca" && tag !== "hmtx").map(([tag, entry]) => [tag, entry.data]));
    const head = new SfntReader(tableOf(tables, "head"));
    head.offset = 50;
    const longOffsets = head.i16() === 1;
    const numGlyphs = new SfntReader(tableOf(tables, "maxp"), 4).u16();
    const glyf = woff2.tables.get("glyf") ?? sfntFail("unsupported-outlines", "The program has no TrueType outlines.");
    const loca = woff2.tables.get("loca") ?? sfntFail("invalid-font");
    let xMins: number[];
    if (glyf.transformVersion === 0) {
      const decoded = decodeTransformedGlyf(glyf.data);
      if (decoded.glyphs.length !== numGlyphs) sfntFail("invalid-font", "WOFF2 glyf glyph count differs from maxp.");
      glyphs = decoded.glyphs;
      xMins = decoded.xMins;
    } else {
      glyphs = decodePlainGlyf(glyf.data, loca.data, numGlyphs, longOffsets);
      xMins = glyphs.map((glyph) => glyph?.bbox[0] ?? 0);
    }
    const hhea = new SfntReader(tableOf(tables, "hhea"), 34);
    const hmtxEntry = woff2.tables.get("hmtx") ?? sfntFail("invalid-font");
    hmtx = decodeHmtx(hmtxEntry.data, hmtxEntry.transformed, hhea.u16(), numGlyphs, xMins);
  } else if (signature === "\u0000\u0001\u0000\u0000" || signature === "true") {
    const plain = plainTables(bytes);
    tables = new Map([...plain.tables].filter(([tag]) => tag !== "glyf" && tag !== "loca" && tag !== "hmtx"));
    const head = new SfntReader(tableOf(plain.tables, "head"));
    head.offset = 50;
    const longOffsets = head.i16() === 1;
    const numGlyphs = new SfntReader(tableOf(plain.tables, "maxp"), 4).u16();
    glyphs = decodePlainGlyf(tableOf(plain.tables, "glyf"), tableOf(plain.tables, "loca"), numGlyphs, longOffsets);
    const hhea = new SfntReader(tableOf(plain.tables, "hhea"), 34);
    hmtx = decodeHmtx(tableOf(plain.tables, "hmtx"), false, hhea.u16(), numGlyphs, []);
  } else if (signature === "OTTO") {
    return sfntFail("unsupported-outlines", "CFF outlines cannot be embedded as TrueType.");
  } else {
    return sfntFail("invalid-font", "The program is neither WOFF2 nor TrueType.");
  }
  const head = new SfntReader(tableOf(tables, "head"), 18);
  const unitsPerEm = head.u16();
  const hhea = new SfntReader(tableOf(tables, "hhea"), 4);
  const ascent = hhea.i16();
  const descent = hhea.i16();
  if (!(unitsPerEm >= 16 && unitsPerEm <= 16_384) || glyphs.length < 1) sfntFail("invalid-font");
  const cmap = readBestCmap(tables.get("cmap"));
  const frozen = Object.freeze(glyphs);
  return Object.freeze({
    tables,
    numGlyphs: glyphs.length,
    unitsPerEm,
    ascent,
    descent,
    advances: Object.freeze(hmtx.advances),
    leftSideBearings: Object.freeze(hmtx.leftSideBearings),
    glyph: (glyphId: number) => {
      if (!Number.isSafeInteger(glyphId) || glyphId < 0 || glyphId >= frozen.length) sfntFail("missing-glyph");
      return frozen[glyphId];
    },
    cmap,
    postscriptName: postscriptNameOf(tables.get("name")),
  });
}

// --- CFF read as TrueType ---------------------------------------------------

type CubicPoint = readonly [number, number];

function cubicAt(p0: CubicPoint, p1: CubicPoint, p2: CubicPoint, p3: CubicPoint, t: number): CubicPoint {
  const u = 1 - t;
  const a = u * u * u, b = 3 * u * u * t, c = 3 * u * t * t, d = t * t * t;
  return [a * p0[0] + b * p1[0] + c * p2[0] + d * p3[0], a * p0[1] + b * p1[1] + c * p2[1] + d * p3[1]];
}

function cubicSlope(p0: CubicPoint, p1: CubicPoint, p2: CubicPoint, p3: CubicPoint, t: number): CubicPoint {
  const u = 1 - t;
  const a = 3 * u * u, b = 6 * u * t, c = 3 * t * t;
  return [
    a * (p1[0] - p0[0]) + b * (p2[0] - p1[0]) + c * (p3[0] - p2[0]),
    a * (p1[1] - p0[1]) + b * (p2[1] - p1[1]) + c * (p3[1] - p2[1]),
  ];
}

/** A cubic as quadratics, each within `tolerance` of it (fontTools' cu2qu
 * bound): the curve split into equal pieces until every piece's quadratic
 * (its control where the piece's two tangents agree on average) stays that
 * close at its quarter points. Each pair is a control and an end point. */
function cubicAsQuadratics(
  p0: CubicPoint, p1: CubicPoint, p2: CubicPoint, p3: CubicPoint, tolerance: number,
): Array<readonly [CubicPoint, CubicPoint]> {
  let pieces: Array<readonly [CubicPoint, CubicPoint]> = [];
  for (let count = 1; count <= 32; count += 1) {
    pieces = [];
    let close = true;
    for (let index = 0; index < count; index += 1) {
      const t0 = index / count, t1 = (index + 1) / count, span = (t1 - t0) / 3;
      const q0 = cubicAt(p0, p1, p2, p3, t0), q3 = cubicAt(p0, p1, p2, p3, t1);
      const s0 = cubicSlope(p0, p1, p2, p3, t0), s1 = cubicSlope(p0, p1, p2, p3, t1);
      const q1: CubicPoint = [q0[0] + s0[0] * span, q0[1] + s0[1] * span];
      const q2: CubicPoint = [q3[0] - s1[0] * span, q3[1] - s1[1] * span];
      const control: CubicPoint = [(3 * (q1[0] + q2[0]) - (q0[0] + q3[0])) / 4, (3 * (q1[1] + q2[1]) - (q0[1] + q3[1])) / 4];
      for (const t of [0.25, 0.5, 0.75]) {
        const cubic = cubicAt(q0, q1, q2, q3, t);
        const u = 1 - t;
        const quadratic: CubicPoint = [
          u * u * q0[0] + 2 * u * t * control[0] + t * t * q3[0],
          u * u * q0[1] + 2 * u * t * control[1] + t * t * q3[1],
        ];
        if (Math.hypot(cubic[0] - quadratic[0], cubic[1] - quadratic[1]) > tolerance) close = false;
      }
      pieces.push([control, q3]);
    }
    if (close) break;
  }
  return pieces;
}

/** A CFF glyph (cff-outlines.ts) as a TrueType simple glyph: contours
 * turned quadratic and clockwise, points on the font's integer grid. */
function cffGlyphAsTrueType(outlines: CffOutlines, glyphId: number, tolerance: number): SfntSimpleGlyph | null {
  const contours: { x: number; y: number; on: boolean }[][] = [];
  let contour: { x: number; y: number; on: boolean }[] = [];
  let current: CubicPoint = [0, 0];
  const close = () => {
    if (contour.length > 1) {
      const first = contour[0], last = contour[contour.length - 1];
      if (last.on && first.on && Math.round(last.x) === Math.round(first.x) && Math.round(last.y) === Math.round(first.y)) contour.pop();
    }
    if (contour.length) contours.push([contour[0], ...contour.slice(1).reverse()]);
    contour = [];
  };
  for (const command of outlines.path(glyphId)) {
    if (command.kind === "M") {
      close();
      current = [command.x, command.y];
      contour.push({ x: command.x, y: command.y, on: true });
    } else if (command.kind === "L") {
      current = [command.x, command.y];
      contour.push({ x: command.x, y: command.y, on: true });
    } else if (command.kind === "C") {
      for (const [control, end] of cubicAsQuadratics(current, [command.x1, command.y1], [command.x2, command.y2], [command.x, command.y], tolerance)) {
        contour.push({ x: control[0], y: control[1], on: false }, { x: end[0], y: end[1], on: true });
      }
      current = [command.x, command.y];
    } else {
      close();
    }
  }
  close();
  if (!contours.length) return null;
  const xs: number[] = [], ys: number[] = [], onCurve: boolean[] = [], endPoints: number[] = [];
  for (const points of contours) {
    for (const point of points) {
      xs.push(Math.floor(point.x + 0.5));
      ys.push(Math.floor(point.y + 0.5));
      onCurve.push(point.on);
    }
    endPoints.push(xs.length - 1);
  }
  return Object.freeze({
    kind: "simple" as const,
    endPoints: Object.freeze(endPoints),
    onCurve: Object.freeze(onCurve),
    xs: Object.freeze(xs),
    ys: Object.freeze(ys),
    overlap: false,
    bbox: Object.freeze([Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] as const),
  });
}

/**
 * A CFF-outlined OpenType program read as a TrueType one: each glyph turned
 * quadratic only when it is asked for (a subset of a 50,000-glyph CJK face
 * converts the few thousand it names). For a font from the reader's own
 * computer (PingFang is one), whose subset the page keeps: converted here,
 * the server never spends seconds on the whole face.
 */
export async function readCffSfntProgramAsTrueType(bytes: Uint8Array): Promise<SfntProgram> {
  if (bytes.length < 12 || String.fromCharCode(...bytes.subarray(0, 4)) !== "OTTO") sfntFail("invalid-font");
  const plain = plainTables(bytes);
  const tableOf = (tag: string) => plain.tables.get(tag) ?? sfntFail("invalid-font", `The program has no ${tag} table.`);
  const outlines = readCffOutlines(plain.tables.get("CFF ") ?? sfntFail("unsupported-outlines", "The program has no CFF outlines."));
  const numGlyphs = new SfntReader(tableOf("maxp"), 4).u16();
  if (numGlyphs !== outlines.glyphCount) sfntFail("invalid-font", "The CFF charstrings and maxp count differ.");
  const hmtx = decodeHmtx(tableOf("hmtx"), false, new SfntReader(tableOf("hhea"), 34).u16(), numGlyphs, []);
  // A TrueType maxp (version 1.0): the subset recounts its glyphs.
  const maxp = new Uint8Array(32);
  const maxpView = new DataView(maxp.buffer);
  maxpView.setUint32(0, 0x00010000);
  maxpView.setUint16(4, numGlyphs);
  maxpView.setUint16(14, 2);
  const tables = new Map([...plain.tables].filter(([tag]) =>
    !["CFF ", "CFF2", "VORG", "glyf", "loca", "hmtx", "maxp"].includes(tag)));
  tables.set("maxp", maxp);
  const unitsPerEm = new SfntReader(tableOf("head"), 18).u16();
  const hhea = new SfntReader(tableOf("hhea"), 4);
  const ascent = hhea.i16();
  const descent = hhea.i16();
  if (!(unitsPerEm >= 16 && unitsPerEm <= 16_384) || numGlyphs < 1) sfntFail("invalid-font");
  const tolerance = unitsPerEm / 1000;
  const converted = new Map<number, SfntGlyph | null>();
  return Object.freeze({
    tables,
    numGlyphs,
    unitsPerEm,
    ascent,
    descent,
    advances: Object.freeze(hmtx.advances),
    leftSideBearings: Object.freeze(hmtx.leftSideBearings),
    glyph: (glyphId: number) => {
      if (!Number.isSafeInteger(glyphId) || glyphId < 0 || glyphId >= numGlyphs) sfntFail("missing-glyph");
      if (!converted.has(glyphId)) {
        try {
          converted.set(glyphId, cffGlyphAsTrueType(outlines, glyphId, tolerance));
        } catch (error) {
          if (error instanceof PdfNativeSfntError) throw error;
          sfntFail("invalid-font", "A CFF glyph could not be drawn.");
        }
      }
      return converted.get(glyphId)!;
    },
    cmap: readBestCmap(tables.get("cmap")),
    postscriptName: postscriptNameOf(tables.get("name")),
  });
}

// --- metrics only -------------------------------------------------------------

/**
 * What text measurement needs from a font program and nothing more: the
 * advance of every glyph, the Unicode map and the vertical metrics. Read
 * from a WOFF2 (TrueType or CFF outlines) or a plain sfnt without decoding a
 * single outline, so a face is ready to measure in a fraction of the time
 * {@link readSfntProgram} takes (lib/text-engine).
 */
export interface SfntMetrics {
  readonly unitsPerEm: number;
  readonly numGlyphs: number;
  /** Advance width of every glyph, in font units. */
  readonly advances: Uint16Array;
  /** The best Unicode map (fontTools' getBestCmap order). */
  readonly cmap: ReadonlyMap<number, number>;
  readonly hhea: { readonly ascender: number; readonly descender: number; readonly lineGap: number };
  /** null when the program has no OS/2 table. */
  readonly os2: {
    readonly typoAscender: number;
    readonly typoDescender: number;
    readonly typoLineGap: number;
    readonly winAscent: number;
    readonly winDescent: number;
    /** fsSelection bit 7. */
    readonly useTypoMetrics: boolean;
  } | null;
  readonly postscriptName: string | null;
  /** The raw variation tables (fvar, avar, HVAR, STAT) of a variable
   * program, for its advances at an instance (sfnt-variation.ts). */
  readonly variationTables?: ReadonlyMap<string, Uint8Array>;
  /** The raw GPOS and kern tables, for pair kerning (lib/text-engine). */
  readonly kerningTables?: ReadonlyMap<string, Uint8Array>;
}

function metricsFromTables(tables: ReadonlyMap<string, Uint8Array>, hmtx: Uint8Array, hmtxTransformed: boolean): SfntMetrics {
  const tableOf = (tag: string) => tables.get(tag) ?? sfntFail("invalid-font", `The program has no ${tag} table.`);
  const unitsPerEm = new SfntReader(tableOf("head"), 18).u16();
  const numGlyphs = new SfntReader(tableOf("maxp"), 4).u16();
  const hheaReader = new SfntReader(tableOf("hhea"), 4);
  const ascender = hheaReader.i16();
  const descender = hheaReader.i16();
  const lineGap = hheaReader.i16();
  const numberOfHMetrics = new SfntReader(tableOf("hhea"), 34).u16();
  if (!(unitsPerEm >= 16 && unitsPerEm <= 16_384) || numGlyphs < 1 || numberOfHMetrics < 1 || numberOfHMetrics > numGlyphs) sfntFail("invalid-font");
  // An untransformed hmtx is (advance, lsb) pairs; the WOFF2 transform (5.4)
  // puts a flags byte first and every advance ahead of the side bearings.
  const reader = new SfntReader(hmtx);
  const stride = hmtxTransformed ? 0 : 2;
  if (hmtxTransformed) reader.u8();
  const advances = new Uint16Array(numGlyphs);
  for (let index = 0; index < numberOfHMetrics; index += 1) {
    advances[index] = reader.u16();
    reader.offset += stride;
  }
  advances.fill(advances[numberOfHMetrics - 1], numberOfHMetrics);
  const os2Table = tables.get("OS/2");
  let os2: SfntMetrics["os2"] = null;
  if (os2Table && os2Table.length >= 78) {
    const os2Reader = new SfntReader(os2Table, 62);
    const fsSelection = os2Reader.u16();
    os2Reader.offset = 68;
    os2 = {
      typoAscender: os2Reader.i16(),
      typoDescender: os2Reader.i16(),
      typoLineGap: os2Reader.i16(),
      winAscent: os2Reader.u16(),
      winDescent: os2Reader.u16(),
      useTypoMetrics: (fsSelection & 0x80) !== 0,
    };
  }
  return Object.freeze({
    unitsPerEm,
    numGlyphs,
    advances,
    cmap: readBestCmap(tables.get("cmap")),
    hhea: Object.freeze({ ascender, descender, lineGap }),
    os2: os2 && Object.freeze(os2),
    postscriptName: postscriptNameOf(tables.get("name")),
    ...(tables.has("GPOS") || tables.has("kern")
      ? { kerningTables: new Map(["GPOS", "kern"].flatMap((tag) => (tables.has(tag) ? [[tag, tables.get(tag)!] as const] : []))) }
      : {}),
    ...(tables.has("fvar")
      ? { variationTables: new Map(["fvar", "avar", "HVAR", "gvar"].flatMap((tag) => (tables.has(tag) ? [[tag, tag === "gvar" ? new Uint8Array(0) : tables.get(tag)!] as const] : []))) }
      : {}),
  });
}

/** Read a font program's metrics (see {@link SfntMetrics}); no outline is decoded. */
export async function readSfntMetrics(bytes: Uint8Array): Promise<SfntMetrics> {
  if (bytes.length < 12) sfntFail("invalid-font");
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  if (signature === "wOF2") {
    const woff2 = await woff2Tables(bytes);
    const tables = new Map([...woff2.tables].map(([tag, entry]) => [tag, entry.data] as const));
    const hmtx = woff2.tables.get("hmtx") ?? sfntFail("invalid-font", "The program has no hmtx table.");
    return metricsFromTables(tables, hmtx.data, hmtx.transformed);
  }
  if (signature === "\u0000\u0001\u0000\u0000" || signature === "true" || signature === "OTTO") {
    const plain = plainTables(bytes);
    return metricsFromTables(plain.tables, plain.tables.get("hmtx") ?? sfntFail("invalid-font", "The program has no hmtx table."), false);
  }
  return sfntFail("invalid-font", "The program is neither WOFF2 nor an sfnt.");
}
