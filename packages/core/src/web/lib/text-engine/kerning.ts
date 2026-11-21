/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pair kerning from a face's own tables, for the editors whose documents ask
 * for it (PowerPoint's `kern` from a size up, Word's `w:kern`): the GPOS
 * `kern` feature's pair adjustments (PairPos formats 1 and 2, through
 * Extension lookups) of the Latin script's default language system (else
 * DFLT's, else every `kern` feature's), else the legacy `kern` table
 * (format 0). Applied as OpenType applies them: every lookup of the feature
 * in turn, the adjustments adding up; within a lookup the first subtable
 * that applies -- a format 1 subtable only when it lists the pair (Calibri's
 * lists `A` with a few partners and leaves `AV` to the class subtable after
 * it), a format 2 one whenever it covers the first glyph. Only the first
 * glyph's x-advance is adjusted, which is all horizontal kerning changes. Excel never kerns cell text, so the
 * spreadsheet never asks.
 */
import { SfntReader } from "../pdf/engine/native/sfnt/sfnt-reader";
import type { FaceMetrics } from "./face-metrics";

export type PairKerning = (left: number, right: number) => number;
/** A subtable's answer for a left glyph it does not cover. */
const NOT_COVERED = Number.NaN;

const cache = new WeakMap<FaceMetrics, PairKerning | null>();
const legacyCache = new WeakMap<FaceMetrics, PairKerning | null>();














export function pairKerning(metrics: FaceMetrics, { legacyKernOnly = false }: { legacyKernOnly?: boolean } = {}): PairKerning | null {
  if (legacyKernOnly) {
    if (legacyCache.has(metrics)) return legacyCache.get(metrics)!;
    let legacy: PairKerning | null = null;
    try {
      const kern = metrics.kerningTables?.get("kern");
      legacy = kern ? legacyKerning(kern) : null;
    } catch {
      legacy = null;
    }
    legacyCache.set(metrics, legacy);
    return legacy;
  }
  if (cache.has(metrics)) return cache.get(metrics)!;
  let result: PairKerning | null = null;
  try {
    const gpos = metrics.kerningTables?.get("GPOS");
    result = gpos ? gposKerning(gpos) : null;
    if (!result) {
      const kern = metrics.kerningTables?.get("kern");
      result = kern ? legacyKerning(kern) : null;
    }
  } catch {
    result = null;
  }
  cache.set(metrics, result);
  return result;
}

function coverage(bytes: Uint8Array, offset: number): Map<number, number> {
  const reader = new SfntReader(bytes, offset);
  const format = reader.u16();
  const out = new Map<number, number>();
  if (format === 1) {
    const count = reader.u16();
    for (let index = 0; index < count; index += 1) out.set(reader.u16(), index);
  } else if (format === 2) {
    const count = reader.u16();
    for (let index = 0; index < count; index += 1) {
      const start = reader.u16();
      const end = reader.u16();
      const startIndex = reader.u16();
      for (let glyph = start; glyph <= end; glyph += 1) out.set(glyph, startIndex + glyph - start);
    }
  }
  return out;
}

function classDef(bytes: Uint8Array, offset: number): (glyph: number) => number {
  if (!offset) return () => 0;
  const reader = new SfntReader(bytes, offset);
  const format = reader.u16();
  const map = new Map<number, number>();
  if (format === 1) {
    const start = reader.u16();
    const count = reader.u16();
    for (let index = 0; index < count; index += 1) map.set(start + index, reader.u16());
  } else if (format === 2) {
    const count = reader.u16();
    for (let index = 0; index < count; index += 1) {
      const start = reader.u16();
      const end = reader.u16();
      const cls = reader.u16();
      for (let glyph = start; glyph <= end; glyph += 1) map.set(glyph, cls);
    }
  }
  return (glyph) => map.get(glyph) ?? 0;
}

const VALUE_SIZE = (format: number) => {
  let size = 0;
  for (let bit = 0; bit < 8; bit += 1) if (format & (1 << bit)) size += 2;
  return size;
};

/** The x-advance field of a ValueRecord at `offset` (0 when absent). */
function xAdvance(bytes: Uint8Array, offset: number, format: number): number {
  if (!(format & 0x0004)) return 0;
  let at = offset;
  if (format & 0x0001) at += 2;
  if (format & 0x0002) at += 2;
  return new SfntReader(bytes, at).i16();
}

function pairSubtable(bytes: Uint8Array, offset: number): PairKerning | null {
  const reader = new SfntReader(bytes, offset);
  const format = reader.u16();
  const coverageOffset = reader.u16();
  const format1 = reader.u16();
  const format2 = reader.u16();
  const covered = coverage(bytes, offset + coverageOffset);
  const size1 = VALUE_SIZE(format1);
  const size2 = VALUE_SIZE(format2);
  if (format === 1) {
    const setCount = reader.u16();
    const setOffsets: number[] = [];
    for (let index = 0; index < setCount; index += 1) setOffsets.push(reader.u16());
    const sets = new Map<number, Map<number, number>>();
    return (left, right) => {
      const index = covered.get(left);
      if (index === undefined || index >= setOffsets.length) return NOT_COVERED;
      let pairs = sets.get(index);
      if (!pairs) {
        pairs = new Map();
        const set = new SfntReader(bytes, offset + setOffsets[index]);
        const count = set.u16();
        for (let pair = 0; pair < count; pair += 1) {
          const second = set.u16();
          const at = set.offset;
          pairs.set(second, xAdvance(bytes, at, format1));
          set.offset += size1 + size2;
        }
        sets.set(index, pairs);
      }
      return pairs.get(right) ?? NOT_COVERED;
    };
  }
  if (format === 2) {
    const classDef1 = classDef(bytes, offset + reader.u16());
    const classDef2 = classDef(bytes, offset + reader.u16());
    reader.u16(); // class1Count
    const class2Count = reader.u16();
    const base = reader.offset;
    const recordSize = size1 + size2;
    return (left, right) => {
      if (!covered.has(left)) return NOT_COVERED;
      const at = base + (classDef1(left) * class2Count + classDef2(right)) * recordSize;
      return at + recordSize <= bytes.length ? xAdvance(bytes, at, format1) : 0;
    };
  }
  return null;
}

/** The `kern` lookups a language system names (feature indexes into the FeatureList). */
function langSysKernLookups(bytes: Uint8Array, langSysOffset: number, kernFeatures: ReadonlyMap<number, readonly number[]>): number[] {
  const reader = new SfntReader(bytes, langSysOffset);
  reader.u16(); // lookupOrder
  const required = reader.u16();
  const count = reader.u16();
  const indexes = [required, ...Array.from({ length: count }, () => reader.u16())];
  return indexes.flatMap((index) => kernFeatures.get(index) ?? []);
}

function gposKerning(bytes: Uint8Array): PairKerning | null {
  const header = new SfntReader(bytes);
  header.u32();
  const scriptListOffset = header.u16();
  const featureListOffset = header.u16();
  const lookupListOffset = header.u16();
  const features = new SfntReader(bytes, featureListOffset);
  const featureCount = features.u16();
  const kernFeatures = new Map<number, number[]>();
  for (let index = 0; index < featureCount; index += 1) {
    const tag = features.tag();
    const offset = features.u16();
    if (tag !== "kern") continue;
    const feature = new SfntReader(bytes, featureListOffset + offset);
    feature.u16();
    const count = feature.u16();
    kernFeatures.set(index, Array.from({ length: count }, () => feature.u16()));
  }
  if (!kernFeatures.size) return null;
  // The lookups the Latin script's default language system names, else
  // DFLT's: the ones a shaper applies to Latin text.
  const byScript = new Map<string, number[]>();
  if (scriptListOffset) {
    const scripts = new SfntReader(bytes, scriptListOffset);
    const scriptCount = scripts.u16();
    for (let index = 0; index < scriptCount; index += 1) {
      const tag = scripts.tag();
      const scriptOffset = scriptListOffset + scripts.u16();
      const defaultLangSys = new SfntReader(bytes, scriptOffset).u16();
      if (defaultLangSys) byScript.set(tag, langSysKernLookups(bytes, scriptOffset + defaultLangSys, kernFeatures));
    }
  }
  const named = byScript.get("latn") ?? byScript.get("DFLT") ?? [...kernFeatures.values()].flat();
  const lookupIndexes = [...new Set(named)].sort((a, b) => a - b);
  if (!lookupIndexes.length) return null;
  const lookups = new SfntReader(bytes, lookupListOffset);
  const lookupCount = lookups.u16();
  const lookupOffsets: number[] = [];
  for (let index = 0; index < lookupCount; index += 1) lookupOffsets.push(lookups.u16());
  const applied: PairKerning[][] = [];
  for (const lookupIndex of lookupIndexes) {
    if (lookupIndex >= lookupOffsets.length) continue;
    const lookupOffset = lookupListOffset + lookupOffsets[lookupIndex];
    const lookup = new SfntReader(bytes, lookupOffset);
    const type = lookup.u16();
    lookup.u16();
    const count = lookup.u16();
    const subtables: PairKerning[] = [];
    for (let sub = 0; sub < count; sub += 1) {
      let subOffset = lookupOffset + lookup.u16();
      let subType = type;
      if (type === 9) {
        const extension = new SfntReader(bytes, subOffset);
        extension.u16();
        subType = extension.u16();
        subOffset += extension.u32();
      }
      if (subType !== 2) continue;
      const pair = pairSubtable(bytes, subOffset);
      if (pair) subtables.push(pair);
    }
    if (subtables.length) applied.push(subtables);
  }
  if (!applied.length) return null;
  return (left, right) => {
    let total = 0;
    for (const subtables of applied) {
      for (const subtable of subtables) {
        const value = subtable(left, right);
        if (!Number.isNaN(value)) {
          total += value;
          break;
        }
      }
    }
    return total;
  };
}

function legacyKerning(bytes: Uint8Array): PairKerning | null {
  const reader = new SfntReader(bytes);
  reader.u16();
  const tables = reader.u16();
  const pairs = new Map<number, number>();
  for (let table = 0; table < tables; table += 1) {
    const start = reader.offset;
    reader.u16();
    const length = reader.u16();
    const coverageField = reader.u16();
    if ((coverageField >> 8) === 0 && (coverageField & 1)) {
      const count = reader.u16();
      reader.offset += 6;
      for (let index = 0; index < count; index += 1) {
        const left = reader.u16();
        const right = reader.u16();
        pairs.set(left * 65536 + right, reader.i16());
      }
    }
    reader.offset = start + length;
  }
  return pairs.size ? (left, right) => pairs.get(left * 65536 + right) ?? 0 : null;
}
