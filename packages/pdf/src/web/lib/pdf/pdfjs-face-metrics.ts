/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const SFNT_HEADER_BYTES = 12;
const SFNT_RECORD_BYTES = 16;
const MAX_FACE_BYTES = 64 * 1024 * 1024;

export interface PdfJsFaceMetrics {
  readonly unitsPerEm: number;
  readonly numGlyphs: number;
  readonly numberOfHMetrics: number;
  /** Advance width per glyph id, in font units. */
  readonly advances: Uint16Array;
  /** Code point → glyph id, every Unicode-keyed cmap subtable merged. */
  readonly cmap: ReadonlyMap<number, number>;
}

interface SfntTable {
  readonly tag: string;
  readonly offset: number;
  readonly length: number;
}

function readTables(bytes: Uint8Array): SfntTable[] | null {
  if (bytes.length < SFNT_HEADER_BYTES || bytes.length > MAX_FACE_BYTES) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const version = view.getUint32(0);
  // TrueType, Apple TrueType, OpenType CFF.
  if (version !== 0x00010000 && version !== 0x74727565 && version !== 0x4f54544f) return null;
  const numTables = view.getUint16(4);
  if (numTables < 1 || numTables > 64) return null;
  if (bytes.length < SFNT_HEADER_BYTES + numTables * SFNT_RECORD_BYTES) return null;
  const tables: SfntTable[] = [];
  for (let index = 0; index < numTables; index += 1) {
    const at = SFNT_HEADER_BYTES + index * SFNT_RECORD_BYTES;
    const tag = String.fromCharCode(bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]);
    const offset = view.getUint32(at + 8);
    const length = view.getUint32(at + 12);
    if (offset > bytes.length || length > bytes.length - offset) return null;
    tables.push({ tag, offset, length });
  }
  return tables;
}

function tableNamed(tables: readonly SfntTable[], tag: string): SfntTable | null {
  return tables.find((table) => table.tag === tag) ?? null;
}

function readCmap(bytes: Uint8Array, table: SfntTable): Map<number, number> {
  const view = new DataView(bytes.buffer, bytes.byteOffset + table.offset, table.length);
  const merged = new Map<number, number>();
  if (table.length < 4) return merged;
  const numSubtables = view.getUint16(2);
  const subtables: Array<{ rank: number; offset: number }> = [];
  for (let index = 0; index < numSubtables && 4 + index * 8 + 8 <= table.length; index += 1) {
    const at = 4 + index * 8;
    const platformId = view.getUint16(at);
    const encodingId = view.getUint16(at + 2);
    const offset = view.getUint32(at + 4);
    if (offset >= table.length) continue;
    // Unicode-keyed subtables only: a (1,0) or (3,0) table keys glyphs by
    // the font's own codes, which the draft's text never carries.
    const rank =
      platformId === 3 && encodingId === 10 ? 0
      : platformId === 3 && encodingId === 1 ? 1
      : platformId === 0 ? 2
      : -1;
    if (rank >= 0) subtables.push({ rank, offset });
  }
  subtables.sort((a, b) => a.rank - b.rank);
  const put = (code: number, glyph: number) => {
    if (glyph !== 0 && !merged.has(code)) merged.set(code, glyph);
  };
  for (const { offset } of subtables) {
    if (offset + 4 > table.length) continue;
    const format = view.getUint16(offset);
    try {
      if (format === 4) {
        const segCountX2 = view.getUint16(offset + 6);
        const segCount = segCountX2 >>> 1;
        const endCodes = offset + 14;
        const startCodes = endCodes + segCountX2 + 2;
        const idDeltas = startCodes + segCountX2;
        const idRangeOffsets = idDeltas + segCountX2;
        for (let segment = 0; segment < segCount; segment += 1) {
          const end = view.getUint16(endCodes + segment * 2);
          const start = view.getUint16(startCodes + segment * 2);
          const delta = view.getInt16(idDeltas + segment * 2);
          const rangeOffsetAt = idRangeOffsets + segment * 2;
          const rangeOffset = view.getUint16(rangeOffsetAt);
          if (start > end || end === 0xffff && start === 0xffff) continue;
          for (let code = start; code <= end && code !== 0xffff; code += 1) {
            let glyph: number;
            if (rangeOffset === 0) glyph = (code + delta) & 0xffff;
            else {
              const glyphAt = rangeOffsetAt + rangeOffset + (code - start) * 2;
              if (glyphAt + 2 > table.length) continue;
              const raw = view.getUint16(glyphAt);
              glyph = raw === 0 ? 0 : (raw + delta) & 0xffff;
            }
            put(code, glyph);
          }
        }
      } else if (format === 12) {
        const groups = view.getUint32(offset + 12);
        for (let group = 0; group < groups; group += 1) {
          const at = offset + 16 + group * 12;
          if (at + 12 > table.length) break;
          const start = view.getUint32(at);
          const end = view.getUint32(at + 4);
          const startGlyph = view.getUint32(at + 8);
          if (end < start || end - start > 0x10ffff) continue;
          for (let code = start; code <= end; code += 1) put(code, startGlyph + (code - start));
        }
      } else if (format === 6) {
        const first = view.getUint16(offset + 6);
        const count = view.getUint16(offset + 8);
        for (let index = 0; index < count; index += 1) {
          const at = offset + 10 + index * 2;
          if (at + 2 > table.length) break;
          put(first + index, view.getUint16(at));
        }
      } else if (format === 0) {
        for (let code = 0; code < 256; code += 1) {
          const at = offset + 6 + code;
          if (at >= table.length) break;
          put(code, bytes[table.offset + at]);
        }
      }
    } catch {
      // A malformed subtable contributes nothing; the others still do.
    }
  }
  return merged;
}

/** The face's own metrics, or null when the bytes are not an sfnt the
 * browser would load (pdf.js registers TrueType and OpenType faces). */
export function readPdfJsFaceMetrics(bytes: Uint8Array): PdfJsFaceMetrics | null {
  const tables = readTables(bytes);
  if (!tables) return null;
  const head = tableNamed(tables, "head");
  const maxp = tableNamed(tables, "maxp");
  const hhea = tableNamed(tables, "hhea");
  const hmtx = tableNamed(tables, "hmtx");
  const cmap = tableNamed(tables, "cmap");
  if (!head || !maxp || !hhea || !hmtx || !cmap || head.length < 54 || maxp.length < 6 || hhea.length < 36) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const unitsPerEm = view.getUint16(head.offset + 18);
  const numGlyphs = view.getUint16(maxp.offset + 4);
  const numberOfHMetrics = view.getUint16(hhea.offset + 34);
  if (!(unitsPerEm >= 16 && unitsPerEm <= 16_384) || numGlyphs < 1 || numberOfHMetrics < 1 ||
    numberOfHMetrics > numGlyphs || hmtx.length < numberOfHMetrics * 4) return null;
  const advances = new Uint16Array(numGlyphs);
  for (let glyph = 0; glyph < numGlyphs; glyph += 1) {
    const index = Math.min(glyph, numberOfHMetrics - 1);
    advances[glyph] = view.getUint16(hmtx.offset + index * 4);
  }
  return Object.freeze({ unitsPerEm, numGlyphs, numberOfHMetrics, advances, cmap: readCmap(bytes, cmap) });
}

function tableChecksum(data: Uint8Array): number {
  let sum = 0;
  for (let at = 0; at < data.length; at += 4) {
    const word = ((data[at] ?? 0) << 24) | ((data[at + 1] ?? 0) << 16) | ((data[at + 2] ?? 0) << 8) | (data[at + 3] ?? 0);
    sum = (sum + (word >>> 0)) >>> 0;
  }
  return sum;
}

/**
 * The same face with the given advances (font units, by glyph id); every
 * other table is copied. Glyphs not named keep their own advance. The
 * horizontal metrics are written one per glyph so a glyph past the face's
 * numberOfHMetrics can take its own.
 */
export function rebuildPdfJsFaceAdvances(
  bytes: Uint8Array,
  advances: ReadonlyMap<number, number>,
): Uint8Array | null {
  return rebuildPdfJsFace(bytes, { advances });
}

/** Unicode-keyed cmap subtables for `entries` (code point → glyph id): a
 * format 4 of the Basic Multilingual Plane under (3,1) where it fits, and a
 * format 12 of all of them under (3,10). */
export function pdfJsFaceCmapTable(entries: ReadonlyMap<number, number>): Uint8Array {
  const sorted = [...entries].filter(([code, glyph]) => code >= 0 && code <= 0x10ffff && glyph > 0 && glyph <= 0xffff)
    .sort((a, b) => a[0] - b[0]);
  const runs = (list: readonly (readonly [number, number])[]) => {
    const out: { start: number; end: number; glyph: number }[] = [];
    for (const [code, glyph] of list) {
      const last = out.at(-1);
      if (last && code === last.end + 1 && glyph === last.glyph + (code - last.start)) last.end = code;
      else out.push({ start: code, end: code, glyph });
    }
    return out;
  };
  const groups = runs(sorted);
  const format12 = new Uint8Array(16 + groups.length * 12);
  const f12 = new DataView(format12.buffer);
  f12.setUint16(0, 12); f12.setUint32(4, format12.length); f12.setUint32(12, groups.length);
  groups.forEach((group, index) => {
    f12.setUint32(16 + index * 12, group.start);
    f12.setUint32(20 + index * 12, group.end);
    f12.setUint32(24 + index * 12, group.glyph);
  });
  const segments = [...runs(sorted.filter(([code]) => code < 0xffff)), { start: 0xffff, end: 0xffff, glyph: 0 }];
  const segCount = segments.length;
  const format4 = 16 + segCount * 8 <= 0xffff ? new Uint8Array(16 + segCount * 8) : null;
  if (format4) {
    const f4 = new DataView(format4.buffer);
    const entrySelector = Math.floor(Math.log2(segCount));
    const searchRange = 2 * 2 ** entrySelector;
    f4.setUint16(0, 4); f4.setUint16(2, format4.length); f4.setUint16(6, segCount * 2);
    f4.setUint16(8, searchRange); f4.setUint16(10, entrySelector); f4.setUint16(12, segCount * 2 - searchRange);
    segments.forEach((segment, index) => {
      f4.setUint16(14 + index * 2, segment.end);
      f4.setUint16(16 + segCount * 2 + index * 2, segment.start);
      // The last segment maps 0xFFFF to glyph 0 (delta 1 wraps it there).
      f4.setUint16(16 + segCount * 4 + index * 2, segment.start === 0xffff ? 1 : (segment.glyph - segment.start) & 0xffff);
      f4.setUint16(16 + segCount * 6 + index * 2, 0);
    });
  }
  const subtables = format4 ? [[3, 1, format4], [3, 10, format12]] as const : [[3, 10, format12]] as const;
  const header = 4 + subtables.length * 8;
  const table = new Uint8Array(header + subtables.reduce((total, [, , data]) => total + data.length, 0));
  const view = new DataView(table.buffer);
  view.setUint16(2, subtables.length);
  let at = header;
  subtables.forEach(([platform, encoding, data], index) => {
    view.setUint16(4 + index * 8, platform); view.setUint16(6 + index * 8, encoding); view.setUint32(8 + index * 8, at);
    table.set(data, at);
    at += data.length;
  });
  return table;
}

/**
 * The same face with the given advances (font units, by glyph id) and, given
 * one, another cmap (code point → glyph id); every other table is copied.
 */
export function rebuildPdfJsFace(
  bytes: Uint8Array,
  changes: Readonly<{ advances?: ReadonlyMap<number, number>; cmap?: ReadonlyMap<number, number> }>,
): Uint8Array | null {
  const advances = changes.advances ?? new Map<number, number>();
  const tables = readTables(bytes);
  const metrics = readPdfJsFaceMetrics(bytes);
  if (!tables || !metrics) return null;
  const hhea = tableNamed(tables, "hhea")!;
  const hmtx = tableNamed(tables, "hmtx")!;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const nextHmtx = new Uint8Array(metrics.numGlyphs * 4);
  const nextHmtxView = new DataView(nextHmtx.buffer);
  for (let glyph = 0; glyph < metrics.numGlyphs; glyph += 1) {
    const wanted = advances.get(glyph);
    const advance = wanted !== undefined && Number.isFinite(wanted)
      ? Math.max(0, Math.min(0xffff, Math.round(wanted))) : metrics.advances[glyph];
    const leftSideBearing = glyph < metrics.numberOfHMetrics
      ? view.getInt16(hmtx.offset + glyph * 4 + 2)
      : hmtx.length >= metrics.numberOfHMetrics * 4 + (glyph - metrics.numberOfHMetrics + 1) * 2
        ? view.getInt16(hmtx.offset + metrics.numberOfHMetrics * 4 + (glyph - metrics.numberOfHMetrics) * 2)
        : 0;
    nextHmtxView.setUint16(glyph * 4, advance);
    nextHmtxView.setInt16(glyph * 4 + 2, leftSideBearing);
  }
  const nextHhea = bytes.slice(hhea.offset, hhea.offset + hhea.length);
  new DataView(nextHhea.buffer).setUint16(34, metrics.numGlyphs);
  const replaced = new Map<string, Uint8Array>([["hmtx", nextHmtx], ["hhea", nextHhea],
    ...(changes.cmap ? [["cmap", pdfJsFaceCmapTable(changes.cmap)] as const] : [])]);
  const padded = (length: number) => (length + 3) & ~3;
  let total = SFNT_HEADER_BYTES + tables.length * SFNT_RECORD_BYTES;
  const payloads = tables.map((table) => {
    const data = replaced.get(table.tag) ?? bytes.subarray(table.offset, table.offset + table.length);
    const at = total;
    total += padded(data.length);
    return { table, data, at };
  });
  const out = new Uint8Array(total);
  out.set(bytes.subarray(0, SFNT_HEADER_BYTES));
  const outView = new DataView(out.buffer);
  let headAt = -1;
  payloads.forEach(({ table, data, at }, index) => {
    const record = SFNT_HEADER_BYTES + index * SFNT_RECORD_BYTES;
    out.set(bytes.subarray(SFNT_HEADER_BYTES + index * SFNT_RECORD_BYTES, SFNT_HEADER_BYTES + index * SFNT_RECORD_BYTES + 4), record);
    out.set(data, at);
    if (table.tag === "head") { headAt = at; outView.setUint32(at + 8, 0); }
    outView.setUint32(record + 4, tableChecksum(out.subarray(at, at + padded(data.length))));
    outView.setUint32(record + 8, at);
    outView.setUint32(record + 12, data.length);
  });
  if (headAt >= 0) outView.setUint32(headAt + 8, (0xb1b0afba - tableChecksum(out)) >>> 0);
  return out;
}

/** A glyph's advance in the PDF (1/1000 text space) against the face's. */
function expectedAdvanceUnits(width1000: number, unitsPerEm: number): number {
  return Math.round((width1000 * unitsPerEm) / 1000);
}

/**
 * The advances the face needs so the draft's glyphs stand where the page
 * draws them: by glyph id, font units; empty when the face already agrees
 * with the page for every glyph it draws.
 */
export function pdfJsFaceAdvanceRepairs(
  metrics: PdfJsFaceMetrics,
  widthsByFontChar: ReadonlyMap<number, number>,
): Map<number, number> {
  const repairs = new Map<number, number>();
  const tolerance = Math.max(2, Math.round(metrics.unitsPerEm * 0.01));
  for (const [code, width] of widthsByFontChar) {
    const glyph = metrics.cmap.get(code);
    if (glyph === undefined || glyph >= metrics.numGlyphs || !Number.isFinite(width) || width < 0) continue;
    const expected = expectedAdvanceUnits(width, metrics.unitsPerEm);
    if (Math.abs(expected - metrics.advances[glyph]) > tolerance) repairs.set(glyph, expected);
  }
  return repairs;
}

/** pdf.js's operator list, read for the glyphs each face draws and the
 * width the page gives each (text space × 1000): face → font char → width. */
export function collectPdfJsGlyphWidths(
  operatorList: { fnArray: ArrayLike<number>; argsArray: ArrayLike<unknown> },
  ops: { setFont: number; showText: number },
): Map<string, Map<number, number>> {
  const byFace = new Map<string, Map<number, number>>();
  let face: string | null = null;
  for (let index = 0; index < operatorList.fnArray.length; index += 1) {
    const fn = operatorList.fnArray[index];
    const args = operatorList.argsArray[index];
    if (fn === ops.setFont) {
      face = Array.isArray(args) && typeof args[0] === "string" ? args[0] : null;
      continue;
    }
    if (fn !== ops.showText || !face || !Array.isArray(args) || !Array.isArray(args[0])) continue;
    let widths = byFace.get(face);
    if (!widths) { widths = new Map(); byFace.set(face, widths); }
    for (const glyph of args[0] as unknown[]) {
      if (!glyph || typeof glyph !== "object") continue;
      const { fontChar, width } = glyph as { fontChar?: unknown; width?: unknown };
      if (typeof fontChar !== "string" || !fontChar || typeof width !== "number") continue;
      const code = fontChar.codePointAt(0)!;
      if (!widths.has(code)) widths.set(code, width);
    }
  }
  return byFace;
}

/** pdf.js's operator list, read for the character each glyph a face draws
 * stands for: face → code point → the font char pdf.js draws it by. */
export function collectPdfJsGlyphUnicodes(
  operatorList: { fnArray: ArrayLike<number>; argsArray: ArrayLike<unknown> },
  ops: { setFont: number; showText: number },
): Map<string, Map<number, number>> {
  const byFace = new Map<string, Map<number, number>>();
  let face: string | null = null;
  for (let index = 0; index < operatorList.fnArray.length; index += 1) {
    const fn = operatorList.fnArray[index];
    const args = operatorList.argsArray[index];
    if (fn === ops.setFont) {
      face = Array.isArray(args) && typeof args[0] === "string" ? args[0] : null;
      continue;
    }
    if (fn !== ops.showText || !face || !Array.isArray(args) || !Array.isArray(args[0])) continue;
    let unicodes = byFace.get(face);
    if (!unicodes) { unicodes = new Map(); byFace.set(face, unicodes); }
    for (const glyph of args[0] as unknown[]) {
      if (!glyph || typeof glyph !== "object") continue;
      const { fontChar, unicode } = glyph as { fontChar?: unknown; unicode?: unknown };
      if (typeof fontChar !== "string" || typeof unicode !== "string") continue;
      const chars = [...unicode], fontChars = [...fontChar];
      if (chars.length !== 1 || fontChars.length !== 1) continue;
      const code = chars[0].codePointAt(0)!;
      if (!unicodes.has(code)) unicodes.set(code, fontChars[0].codePointAt(0)!);
    }
  }
  return byFace;
}

const isPrivateUse = (code: number) => (code >= 0xe000 && code <= 0xf8ff) || code >= 0xf0000;











export function pdfJsFaceUnicodeCmap(
  metrics: PdfJsFaceMetrics,
  unicodeToFontChar: ReadonlyMap<number, number>,
): Map<number, number> | null {
  const entries = new Map<number, number>();
  let scrambled = false;
  for (const [code, fontChar] of unicodeToFontChar) {
    const glyph = metrics.cmap.get(fontChar);
    if (glyph === undefined || glyph <= 0 || glyph >= metrics.numGlyphs || isPrivateUse(code)) continue;
    const own = metrics.cmap.get(code);
    if (own === glyph) continue;
    if (own !== undefined) scrambled = true;
    entries.set(code, glyph);
  }
  if (!entries.size) return null;
  const cmap = new Map<number, number>();
  for (const [code, glyph] of metrics.cmap) if (!scrambled || isPrivateUse(code)) cmap.set(code, glyph);
  for (const [code, glyph] of entries) cmap.set(code, glyph);
  return cmap;
}

const REPAIRED_SUFFIX = "-mrai";
const repairedFamilies = new Map<string, string>();
const repairsInFlight = new Map<string, Promise<string | null>>();
const repairReports = new Map<string, { glyphs: number; repaired: number; family: string | null }>();

/** The family registered for a pdf.js face whose advances were repaired. */
export function repairedPdfJsFaceFamily(loadedName: string): string | null {
  return repairedFamilies.get(loadedName) ?? null;
}

/** Development and probes: what each face needed. */
export function readPdfJsFaceRepairReport(): Record<string, { glyphs: number; repaired: number; family: string | null }> {
  return Object.fromEntries(repairReports);
}

const PDF_JS_FACE_TOKEN = /^"?(g_d\d+_f\d+)"?$/;

/** A font stack with every pdf.js face replaced by its repaired family. */
export function withRepairedPdfJsFaces(fontFamily: string): string {
  if (!fontFamily || repairedFamilies.size === 0 || !/g_d\d+_f\d+/.test(fontFamily)) return fontFamily;
  return fontFamily.split(",").map((token) => {
    const match = PDF_JS_FACE_TOKEN.exec(token.trim());
    const repaired = match ? repairedFamilies.get(match[1]) : undefined;
    return repaired ? ` "${repaired}"` : token;
  }).join(",").trimStart();
}

/** What each face is drawn with on the pages read so far: its glyphs' widths
 * and the characters they stand for, and which copy of it stands now. */
const seenByFace = new Map<string, { widths: Map<number, number>; unicodes: Map<number, number>; version: number;
  face: FontFace | null }>();

/**
 * Registers a repaired copy of the face when the pages draw its glyphs at
 * other advances than the face carries, or draw characters it does not key
 * to those glyphs (pdfJsFaceUnicodeCmap); resolves the repaired family, or
 * null when the face agrees with the pages, cannot be read, or cannot be
 * loaded. A page that draws glyphs or characters of it the pages before did
 * not makes a new copy (the draft of a line on page 3 found page 1's only).
 */
export function ensureRepairedPdfJsFace(input: {
  readonly loadedName: string;
  readonly data: Uint8Array | null | undefined;
  readonly widthsByFontChar: ReadonlyMap<number, number>;
  /** The characters the page draws in the face (collectPdfJsGlyphUnicodes). */
  readonly unicodeToFontChar?: ReadonlyMap<number, number>;
  




  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly target?: Document | null;
}): Promise<string | null> {
  let seen = seenByFace.get(input.loadedName);
  const first = !seen;
  if (!seen) {
    seen = { widths: new Map(), unicodes: new Map(), version: 0, face: null };
    seenByFace.set(input.loadedName, seen);
  }
  let grew = first;
  for (const [code, width] of input.widthsByFontChar) if (!seen.widths.has(code)) { seen.widths.set(code, width); grew = true; }
  for (const [code, fontChar] of input.unicodeToFontChar ?? []) {
    if (!seen.unicodes.has(code)) { seen.unicodes.set(code, fontChar); grew = true; }
  }
  const inFlight = repairsInFlight.get(input.loadedName);
  if (!grew) return inFlight ?? Promise.resolve(repairedFamilies.get(input.loadedName) ?? null);
  const target = input.target ?? (typeof document === "undefined" ? null : document);
  const state = seen;
  const build = async (): Promise<string | null> => {
    const widths = state.widths;
    try {
      if (!input.data?.length || !target?.fonts || typeof FontFace === "undefined") {
        repairReports.set(input.loadedName, { glyphs: widths.size, repaired: -2, family: null });
        return null;
      }
      const metrics = readPdfJsFaceMetrics(input.data);
      if (!metrics) {
        repairReports.set(input.loadedName, { glyphs: widths.size, repaired: -1, family: null });
        return null;
      }
      const repairs = pdfJsFaceAdvanceRepairs(metrics, widths);
      const cmap = state.unicodes.size ? pdfJsFaceUnicodeCmap(metrics, state.unicodes) : null;
      repairReports.set(input.loadedName, { glyphs: widths.size, repaired: repairs.size, family: null });
      if (process.env.NODE_ENV !== "production") {
        try {
          console.debug("[pdf-face] " + JSON.stringify({ face: input.loadedName, bytes: input.data.length,
            unitsPerEm: metrics.unitsPerEm, glyphs: metrics.numGlyphs, cmap: metrics.cmap.size,
            drawn: widths.size, repairs: repairs.size, characters: state.unicodes.size, keyed: cmap?.size ?? 0 }));
        } catch { /* diagnostics never throw */ }
      }
      if (repairs.size === 0 && !cmap && !input.bold && !input.italic) return repairedFamilies.get(input.loadedName) ?? null;
      const rebuilt = rebuildPdfJsFace(input.data, { advances: repairs, ...(cmap ? { cmap } : {}) });
      if (!rebuilt) return repairedFamilies.get(input.loadedName) ?? null;
      state.version += 1;
      const family = `${input.loadedName}${REPAIRED_SUFFIX}${state.version > 1 ? state.version : ""}`;
      // FontFace takes an ArrayBuffer (the typed array's view could sit on a
      // shared buffer, which the DOM type refuses).
      const face = new FontFace(family,
        rebuilt.buffer.slice(rebuilt.byteOffset, rebuilt.byteOffset + rebuilt.byteLength) as ArrayBuffer,
        { display: "block", weight: input.bold ? "700" : "400", style: input.italic ? "italic" : "normal" });
      await face.load();
      target.fonts.add(face);
      // The copy before it goes once a draft set in it has had time to take
      // the new one.
      const before = state.face;
      state.face = face;
      if (before) setTimeout(() => { try { target.fonts.delete(before); } catch { /* gone with its document */ } }, 30_000);
      repairedFamilies.set(input.loadedName, family);
      repairReports.set(input.loadedName, { glyphs: widths.size, repaired: repairs.size, family });
      return family;
    } catch {
      return repairedFamilies.get(input.loadedName) ?? null;
    }
  };
  // One build of a face at a time, in the order the pages came.
  const promise = (inFlight ?? Promise.resolve(null)).then(build).finally(() => {
    if (repairsInFlight.get(input.loadedName) === promise) repairsInFlight.delete(input.loadedName);
  });
  repairsInFlight.set(input.loadedName, promise);
  return promise;
}

if (process.env.NODE_ENV !== "production" && typeof window !== "undefined") {
  // Probes read what each face needed.
  (window as unknown as { __pdfJsFaceRepairs?: () => unknown }).__pdfJsFaceRepairs = readPdfJsFaceRepairReport;
}
