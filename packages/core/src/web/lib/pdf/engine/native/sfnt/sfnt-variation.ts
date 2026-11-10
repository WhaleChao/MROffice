/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A glyph of a variable TrueType program at one weight. The open faces MRAI
 * serves are often variable fonts whose default instance is not the weight a
 * text asks for (Noto Sans SC's default is Thin), so the glyph is computed at
 * the face's location the way fontTools' glyph set computes it -- the server
 * proves the embedded outline against exactly that (native_font_extension.py):
 * the user location normalized through `fvar` and `avar` (v1), every `gvar`
 * tuple scaled by its support and its untouched points inferred (IUP), the
 * four phantom points moved with them, and the advance from `HVAR` when the
 * program has one, else from the phantom points.
 *
 * The coordinates stay unrounded here; the subset writer rounds them once.
 */

import {
  COMPOSITE_ARGS_ARE_XY_VALUES,
  SfntReader,
  sfntFail,
  type SfntComponent,
  type SfntGlyph,
  type SfntProgram,
} from "./sfnt-reader";

export interface SfntVariationAxis {
  readonly tag: string;
  readonly min: number;
  readonly defaultValue: number;
  readonly max: number;
}

/** A glyph at a location, unrounded. */
export type SfntInstancedGlyph =
  | Readonly<{
      kind: "simple";
      endPoints: readonly number[];
      onCurve: readonly boolean[];
      xs: readonly number[];
      ys: readonly number[];
      overlap: boolean;
    }>
  | Readonly<{ kind: "composite"; components: readonly SfntComponent[]; dxs: readonly number[]; dys: readonly number[] }>;

export interface SfntInstance {
  readonly glyph: SfntInstancedGlyph | null;
  /** Left and right phantom point x at the location. */
  readonly leftSideX: number;
  readonly rightSideX: number;
  /** The advance fontTools reports at the location (unrounded with HVAR). */
  readonly advance: number;
}

export function sfntVariationAxes(program: SfntProgram): readonly SfntVariationAxis[] {
  const fvar = program.tables.get("fvar");
  if (!fvar) return [];
  const reader = new SfntReader(fvar);
  reader.u16(); reader.u16();
  const axesOffset = reader.u16();
  reader.u16();
  const axisCount = reader.u16();
  const axisSize = reader.u16();
  const axes: SfntVariationAxis[] = [];
  for (let index = 0; index < axisCount; index += 1) {
    const axis = reader.at(axesOffset + index * axisSize);
    const tag = axis.tag();
    const min = axis.fixed();
    const defaultValue = axis.fixed();
    const max = axis.fixed();
    axes.push(Object.freeze({ tag, min, defaultValue, max }));
  }
  return Object.freeze(axes);
}

function normalizeValue(value: number, axis: SfntVariationAxis): number {
  const { min: lower, defaultValue, max: upper } = axis;
  const clamped = Math.max(Math.min(value, upper), lower);
  if (clamped === defaultValue || lower === upper) return 0;
  if ((clamped < defaultValue && lower !== defaultValue) || (clamped > defaultValue && upper === defaultValue)) {
    return (clamped - defaultValue) / (defaultValue - lower);
  }
  return (clamped - defaultValue) / (upper - defaultValue);
}

function piecewiseLinearMap(value: number, mapping: ReadonlyMap<number, number>): number {
  if (mapping.size === 0) return value;
  if (mapping.has(value)) return mapping.get(value)!;
  const keys = [...mapping.keys()];
  const lowest = Math.min(...keys);
  if (value < lowest) return value + mapping.get(lowest)! - lowest;
  const highest = Math.max(...keys);
  if (value > highest) return value + mapping.get(highest)! - highest;
  const a = Math.max(...keys.filter((key) => key < value));
  const b = Math.min(...keys.filter((key) => key > value));
  const va = mapping.get(a)!;
  const vb = mapping.get(b)!;
  return va + ((vb - va) * (value - a)) / (b - a);
}

/** The normalized location of a user location (fontTools normalizeLocation
 * with the avar v1 mapping). An axis the location leaves out is at its
 * default. */
export function normalizeSfntLocation(
  program: SfntProgram,
  location: Readonly<Record<string, number>>,
): Readonly<Record<string, number>> {
  const axes = sfntVariationAxes(program);
  const normalized: Record<string, number> = {};
  for (const axis of axes) normalized[axis.tag] = normalizeValue(location[axis.tag] ?? axis.defaultValue, axis);
  const avar = program.tables.get("avar");
  if (avar) {
    const reader = new SfntReader(avar);
    const major = reader.u16();
    reader.u16(); reader.u16();
    if (major !== 1) sfntFail("unsupported-variation", "Only avar version 1 is read.");
    const axisCount = reader.u16();
    if (axisCount !== axes.length) sfntFail("invalid-font", "avar axis count differs from fvar.");
    for (const axis of axes) {
      const count = reader.u16();
      const mapping = new Map<number, number>();
      for (let index = 0; index < count; index += 1) {
        const from = reader.f2dot14();
        const to = reader.f2dot14();
        mapping.set(from, to);
      }
      normalized[axis.tag] = piecewiseLinearMap(normalized[axis.tag], mapping);
    }
  }
  return Object.freeze(normalized);
}

type Support = ReadonlyMap<string, readonly [number, number, number]>;

/** fontTools varLib.models.supportScalar with ot=True. */
function supportScalar(location: Readonly<Record<string, number>>, support: Support): number {
  let scalar = 1;
  for (const [axis, [lower, peak, upper]] of support) {
    if (peak === 0) continue;
    if (lower > peak || peak > upper) continue;
    if (lower < 0 && upper > 0) continue;
    const value = location[axis] ?? 0;
    if (value === peak) continue;
    if (value <= lower || upper <= value) return 0;
    scalar *= value < peak ? (value - lower) / (peak - lower) : (value - upper) / (peak - upper);
  }
  return scalar;
}

interface GvarTable {
  readonly axisCount: number;
  readonly sharedTuples: readonly (readonly number[])[];
  readonly offsets: readonly number[];
  readonly dataStart: number;
  readonly bytes: Uint8Array;
}

function readGvar(program: SfntProgram): GvarTable | null {
  const bytes = program.tables.get("gvar");
  if (!bytes) return null;
  const reader = new SfntReader(bytes);
  reader.u16(); reader.u16();
  const axisCount = reader.u16();
  const sharedTupleCount = reader.u16();
  const sharedTuplesOffset = reader.u32();
  const glyphCount = reader.u16();
  const flags = reader.u16();
  const dataStart = reader.u32();
  const offsets: number[] = [];
  for (let index = 0; index <= glyphCount; index += 1) offsets.push(flags & 1 ? reader.u32() : reader.u16() * 2);
  const shared = reader.at(sharedTuplesOffset);
  const sharedTuples: number[][] = [];
  for (let index = 0; index < sharedTupleCount; index += 1) {
    const tuple: number[] = [];
    for (let axis = 0; axis < axisCount; axis += 1) tuple.push(shared.f2dot14());
    sharedTuples.push(tuple);
  }
  return { axisCount, sharedTuples, offsets, dataStart, bytes };
}

function readPointNumbers(reader: SfntReader): number[] | null {
  let count = reader.u8();
  if (count === 0) return null; // every point
  if (count & 0x80) count = ((count & 0x7f) << 8) | reader.u8();
  const points: number[] = [];
  let last = 0;
  while (points.length < count) {
    const control = reader.u8();
    const run = (control & 0x7f) + 1;
    for (let index = 0; index < run && points.length < count; index += 1) {
      last += control & 0x80 ? reader.u16() : reader.u8();
      points.push(last);
    }
  }
  return points;
}

function readDeltas(reader: SfntReader, count: number): number[] {
  const deltas: number[] = [];
  while (deltas.length < count) {
    const control = reader.u8();
    const run = (control & 0x3f) + 1;
    const kind = control & 0xc0;
    for (let index = 0; index < run; index += 1) {
      deltas.push(kind === 0x80 ? 0 : kind === 0x40 ? reader.i16() : kind === 0xc0 ? reader.i32() : reader.i8());
    }
  }
  if (deltas.length !== count) sfntFail("invalid-font", "gvar deltas overrun their points.");
  return deltas;
}

/** fontTools varLib.iup.iup_segment over one axis pair. */
function iupSegment(
  coords: readonly (readonly [number, number])[],
  rc1: readonly [number, number], rd1: readonly [number, number],
  rc2: readonly [number, number], rd2: readonly [number, number],
): [number, number][] {
  const out: [number, number][] = coords.map(() => [0, 0]);
  for (const axis of [0, 1] as const) {
    let x1 = rc1[axis], x2 = rc2[axis], d1 = rd1[axis], d2 = rd2[axis];
    if (x1 === x2) {
      const value = d1 === d2 ? d1 : 0;
      coords.forEach((_, index) => { out[index][axis] = value; });
      continue;
    }
    if (x1 > x2) {
      [x1, x2] = [x2, x1];
      [d1, d2] = [d2, d1];
    }
    const scale = (d2 - d1) / (x2 - x1);
    coords.forEach((point, index) => {
      const x = point[axis];
      out[index][axis] = x <= x1 ? d1 : x >= x2 ? d2 : d1 + (x - x1) * scale;
    });
  }
  return out;
}

function iupContour(
  deltas: readonly ((readonly [number, number]) | null)[],
  coords: readonly (readonly [number, number])[],
): (readonly [number, number])[] {
  if (deltas.every((delta) => delta !== null)) return deltas as (readonly [number, number])[];
  const indices = deltas.flatMap((delta, index) => (delta === null ? [] : [index]));
  if (indices.length === 0) return deltas.map(() => [0, 0] as const);
  const out: (readonly [number, number])[] = [];
  let start = indices[0];
  if (start !== 0) {
    const last = indices[indices.length - 1];
    out.push(...iupSegment(coords.slice(0, start), coords[start], deltas[start]!, coords[last], deltas[last]!));
  }
  out.push(deltas[start]!);
  for (const end of indices.slice(1)) {
    if (end - start > 1) out.push(...iupSegment(coords.slice(start + 1, end), coords[start], deltas[start]!, coords[end], deltas[end]!));
    out.push(deltas[end]!);
    start = end;
  }
  if (start !== deltas.length - 1) {
    out.push(...iupSegment(coords.slice(start + 1), coords[start], deltas[start]!, coords[indices[0]], deltas[indices[0]]!));
  }
  return out;
}

/** fontTools varLib.iup.iup_delta: each contour (and each phantom point on
 * its own) infers the deltas of the points its tuple left out. */
function iupDelta(
  deltas: readonly ((readonly [number, number]) | null)[],
  coords: readonly (readonly [number, number])[],
  ends: readonly number[],
): (readonly [number, number])[] {
  const n = coords.length;
  const all = [...ends, n - 4, n - 3, n - 2, n - 1];
  const out: (readonly [number, number])[] = [];
  let start = 0;
  for (const endIndex of all) {
    const end = endIndex + 1;
    out.push(...iupContour(deltas.slice(start, end), coords.slice(start, end)));
    start = end;
  }
  return out;
}

interface ItemVariationStore {
  readonly regions: readonly Support[];
  readonly data: readonly Readonly<{ regionIndexes: readonly number[]; rows: readonly (readonly number[])[] }>[];
}

function readItemVariationStore(bytes: Uint8Array, offset: number, axes: readonly SfntVariationAxis[]): ItemVariationStore {
  const reader = new SfntReader(bytes, offset);
  if (reader.u16() !== 1) sfntFail("unsupported-variation", "ItemVariationStore format is unknown.");
  const regionListOffset = reader.u32();
  const dataCount = reader.u16();
  const dataOffsets: number[] = [];
  for (let index = 0; index < dataCount; index += 1) dataOffsets.push(reader.u32());
  const regionReader = new SfntReader(bytes, offset + regionListOffset);
  const axisCount = regionReader.u16();
  const regionCount = regionReader.u16();
  const regions: Support[] = [];
  for (let index = 0; index < regionCount; index += 1) {
    const support = new Map<string, readonly [number, number, number]>();
    for (let axis = 0; axis < axisCount; axis += 1) {
      const start = regionReader.f2dot14();
      const peak = regionReader.f2dot14();
      const end = regionReader.f2dot14();
      support.set(axes[axis]?.tag ?? `#${axis}`, [start, peak, end] as const);
    }
    regions.push(support);
  }
  const data = dataOffsets.map((dataOffset) => {
    const item = new SfntReader(bytes, offset + dataOffset);
    const itemCount = item.u16();
    const wordDeltaCount = item.u16();
    const longWords = (wordDeltaCount & 0x8000) !== 0;
    const wordCount = wordDeltaCount & 0x7fff;
    const regionIndexCount = item.u16();
    const regionIndexes: number[] = [];
    for (let index = 0; index < regionIndexCount; index += 1) regionIndexes.push(item.u16());
    const rows: number[][] = [];
    for (let row = 0; row < itemCount; row += 1) {
      const values: number[] = [];
      for (let column = 0; column < regionIndexCount; column += 1) {
        const word = column < wordCount;
        values.push(longWords ? (word ? item.i32() : item.i16()) : (word ? item.i16() : item.i8()));
      }
      rows.push(values);
    }
    return Object.freeze({ regionIndexes, rows });
  });
  return { regions, data };
}

function readDeltaSetIndexMap(bytes: Uint8Array, offset: number): (glyphId: number) => number {
  const reader = new SfntReader(bytes, offset);
  const format = reader.u8();
  const entryFormat = reader.u8();
  const count = format === 1 ? reader.u32() : reader.u16();
  const entrySize = ((entryFormat & 0x30) >> 4) + 1;
  const innerBits = (entryFormat & 0x0f) + 1;
  const entries: number[] = [];
  for (let index = 0; index < count; index += 1) {
    let value = 0;
    for (let byte = 0; byte < entrySize; byte += 1) value = value * 256 + reader.u8();
    const outer = Math.floor(value / 2 ** innerBits);
    const inner = value % 2 ** innerBits;
    entries.push(outer * 65536 + inner);
  }
  return (glyphId) => (entries.length === 0 ? glyphId : entries[Math.min(glyphId, entries.length - 1)]);
}

/** Reads a program's variation tables once and answers glyphs at one
 * normalized location. */
export function createSfntInstancer(program: SfntProgram, location: Readonly<Record<string, number>>) {
  const axes = sfntVariationAxes(program);
  if (axes.length === 0) sfntFail("unsupported-variation", "The program has no variation axes.");
  if (program.tables.has("CFF2") || program.tables.has("VARC")) sfntFail("unsupported-variation");
  const gvar = readGvar(program);
  if (gvar && gvar.axisCount !== axes.length) sfntFail("invalid-font", "gvar axis count differs from fvar.");
  const hvarBytes = program.tables.get("HVAR");
  let advanceDelta: ((glyphId: number) => number) | null = null;
  if (hvarBytes) {
    const reader = new SfntReader(hvarBytes);
    reader.u16(); reader.u16();
    const storeOffset = reader.u32();
    const advanceMapOffset = reader.u32();
    const store = readItemVariationStore(hvarBytes, storeOffset, axes);
    const map = advanceMapOffset ? readDeltaSetIndexMap(hvarBytes, advanceMapOffset) : (glyphId: number) => glyphId;
    const scalars = new Map<number, number>();
    const regionScalar = (index: number) => {
      let scalar = scalars.get(index);
      if (scalar === undefined) {
        scalar = supportScalar(location, store.regions[index] ?? new Map());
        scalars.set(index, scalar);
      }
      return scalar;
    };
    advanceDelta = (glyphId) => {
      const varIndex = map(glyphId);
      if (varIndex === 0xffffffff) return 0;
      const outer = Math.floor(varIndex / 65536);
      const inner = varIndex % 65536;
      const data = store.data[outer];
      const row = data?.rows[inner];
      if (!data || !row) sfntFail("invalid-font", "HVAR names a delta set that is not there.");
      let delta = 0;
      data.regionIndexes.forEach((regionIndex, column) => {
        const scalar = regionScalar(regionIndex);
        if (scalar) delta += row[column] * scalar;
      });
      return delta;
    };
  }

  const tupleVariations = (glyphId: number, pointCount: number) => {
    if (!gvar || glyphId + 1 >= gvar.offsets.length) return [];
    const start = gvar.offsets[glyphId];
    const end = gvar.offsets[glyphId + 1];
    if (start === end) return [];
    const base = gvar.dataStart + start;
    const reader = new SfntReader(gvar.bytes, base);
    const tupleCountField = reader.u16();
    const dataOffset = reader.u16();
    const tupleCount = tupleCountField & 0x0fff;
    const data = new SfntReader(gvar.bytes, base + dataOffset);
    const sharedPoints = tupleCountField & 0x8000 ? readPointNumbers(data) : undefined;
    const variations: { support: Support; points: number[] | null; deltas: [number, number][] }[] = [];
    for (let index = 0; index < tupleCount; index += 1) {
      const size = reader.u16();
      const tupleIndex = reader.u16();
      const peak = tupleIndex & 0x8000
        ? Array.from({ length: gvar.axisCount }, () => reader.f2dot14())
        : gvar.sharedTuples[tupleIndex & 0x0fff] ?? sfntFail("invalid-font", "gvar names a shared tuple that is not there.");
      let startTuple: number[] | null = null;
      let endTuple: number[] | null = null;
      if (tupleIndex & 0x4000) {
        startTuple = Array.from({ length: gvar.axisCount }, () => reader.f2dot14());
        endTuple = Array.from({ length: gvar.axisCount }, () => reader.f2dot14());
      }
      const support = new Map<string, readonly [number, number, number]>();
      axes.forEach((axis, axisIndex) => {
        const p = peak[axisIndex];
        const s = startTuple ? startTuple[axisIndex] : Math.min(p, 0);
        const e = endTuple ? endTuple[axisIndex] : Math.max(p, 0);
        if (s !== 0 || p !== 0 || e !== 0) support.set(axis.tag, [s, p, e] as const);
      });
      const tupleData = new SfntReader(gvar.bytes, data.offset);
      const points = tupleIndex & 0x2000 ? readPointNumbers(tupleData) : (sharedPoints ?? null);
      const count = points === null ? pointCount : points.length;
      const xs = readDeltas(tupleData, count);
      const ys = readDeltas(tupleData, count);
      data.offset += size;
      variations.push({ support, points, deltas: xs.map((x, at) => [x, ys[at]]) });
    }
    return variations;
  };

  /** The glyph at the location, with the phantom points it moved. */
  const instance = (glyphId: number): SfntInstance => {
    const glyph: SfntGlyph | null = program.glyph(glyphId);
    const advance = program.advances[glyphId];
    const lsb = program.leftSideBearings[glyphId];
    const xMin = glyph ? glyph.bbox[0] : 0;
    // fontTools' phantom points: left at xMin - lsb, right an advance on;
    // the vertical pair moves nothing horizontal and is carried along.
    const leftSideX = xMin - lsb;
    const base: [number, number][] = glyph?.kind === "simple"
      ? glyph.xs.map((x, index) => [x, glyph.ys[index]])
      : glyph?.kind === "composite"
        ? glyph.components.map((component) =>
            component.flags & COMPOSITE_ARGS_ARE_XY_VALUES ? [component.arg1, component.arg2] : [0, 0])
        : [];
    const coords: [number, number][] = [...base, [leftSideX, 0], [leftSideX + advance, 0], [0, 0], [0, 0]];
    const ends = glyph?.kind === "simple"
      ? glyph.endPoints
      : glyph?.kind === "composite" ? glyph.components.map((_, index) => index) : [];
    const moved = coords.map(([x, y]) => [x, y] as [number, number]);
    for (const variation of tupleVariations(glyphId, coords.length)) {
      const scalar = supportScalar(location, variation.support);
      if (!scalar) continue;
      let deltas: (readonly [number, number])[];
      if (variation.points === null) {
        deltas = variation.deltas;
      } else {
        const sparse: ((readonly [number, number]) | null)[] = coords.map(() => null);
        variation.points.forEach((point, index) => {
          if (point < sparse.length) sparse[point] = variation.deltas[index];
        });
        deltas = sparse.every((delta) => delta !== null)
          ? (sparse as (readonly [number, number])[])
          : iupDelta(sparse, coords, ends);
      }
      if (deltas.length !== coords.length) sfntFail("invalid-font", "gvar deltas do not match the glyph's points.");
      deltas.forEach(([dx, dy], index) => {
        moved[index][0] += dx * scalar;
        moved[index][1] += dy * scalar;
      });
    }
    const pointCount = base.length;
    const left = moved[pointCount][0];
    const right = moved[pointCount + 1][0];
    const instanced: SfntInstancedGlyph | null = glyph === null
      ? null
      : glyph.kind === "simple"
        ? Object.freeze({
            kind: "simple" as const,
            endPoints: glyph.endPoints,
            onCurve: glyph.onCurve,
            xs: moved.slice(0, pointCount).map(([x]) => x),
            ys: moved.slice(0, pointCount).map(([, y]) => y),
            overlap: glyph.overlap,
          })
        : Object.freeze({
            kind: "composite" as const,
            components: glyph.components,
            dxs: moved.slice(0, pointCount).map(([x]) => x),
            dys: moved.slice(0, pointCount).map(([, y]) => y),
          });
    const advanceAtLocation = advanceDelta
      ? advance + advanceDelta(glyphId)
      : Math.floor(right - left + 0.5);
    return Object.freeze({ glyph: instanced, leftSideX: left, rightSideX: right, advance: advanceAtLocation });
  };
  return Object.freeze({ axes, instance });
}

/**
 * The advance deltas of a variable program at a user location, from its
 * HVAR table: what text measurement needs, without decoding a glyph
 * (lib/text-engine). Null when the program has no HVAR (its advances would
 * come from the outlines' phantom points). `tables` needs only fvar, avar
 * and HVAR.
 */
export function createSfntAdvanceDeltas(
  tables: ReadonlyMap<string, Uint8Array>,
  userLocation: Readonly<Record<string, number>>,
): ((glyphId: number) => number) | null {
  const program = { tables } as unknown as SfntProgram;
  const axes = sfntVariationAxes(program);
  const hvarBytes = tables.get("HVAR");
  if (!axes.length || !hvarBytes) return null;
  const location = normalizeSfntLocation(program, userLocation);
  const reader = new SfntReader(hvarBytes);
  reader.u16(); reader.u16();
  const storeOffset = reader.u32();
  const advanceMapOffset = reader.u32();
  const store = readItemVariationStore(hvarBytes, storeOffset, axes);
  const map = advanceMapOffset ? readDeltaSetIndexMap(hvarBytes, advanceMapOffset) : (glyphId: number) => glyphId;
  const scalars = store.regions.map((region) => supportScalar(location, region));
  return (glyphId) => {
    const varIndex = map(glyphId);
    if (varIndex === 0xffffffff) return 0;
    const data = store.data[Math.floor(varIndex / 65536)];
    const row = data?.rows[varIndex % 65536];
    if (!data || !row) return 0;
    let delta = 0;
    data.regionIndexes.forEach((regionIndex, column) => {
      const scalar = scalars[regionIndex] ?? 0;
      if (scalar) delta += row[column] * scalar;
    });
    return delta;
  };
}
