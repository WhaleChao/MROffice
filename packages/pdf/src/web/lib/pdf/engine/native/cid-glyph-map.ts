/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The glyphs a TrueType program holds: its maxp count. */
export function readPdfTrueTypeGlyphCount(font: Uint8Array): number {
  const fail = (): never => { throw new Error("CIDToGIDMap or its TrueType glyph table is invalid."); };
  if (font.length < 12) fail();
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  if (![0x00010000, 0x74727565].includes(view.getUint32(0))) fail();
  const count = view.getUint16(4);
  if (!count || count > 4096 || 12 + count * 16 > font.length) fail();
  let glyphCount: number | null = null;
  for (let i = 0; i < count; i++) {
    const entry = 12 + i * 16;
    if (view.getUint32(entry) !== 0x6d617870) continue; // maxp
    const offset = view.getUint32(entry + 8), length = view.getUint32(entry + 12);
    if (glyphCount !== null || length < 6 || offset < 12 + count * 16 || offset + length > font.length) fail();
    glyphCount = view.getUint16(offset + 4);
  }
  const glyphLimit = glyphCount ?? fail();
  if (!glyphLimit) fail();
  return glyphLimit;
}

/** Bounded existing-font coverage, not a font conversion or replacement. */
export function readPdfCidGlyphCoverage(mapping: Uint8Array, font: Uint8Array): ReadonlySet<number> {
  if (!mapping.length || mapping.length > 131_072 || mapping.length % 2) {
    throw new Error("CIDToGIDMap or its TrueType glyph table is invalid.");
  }
  const glyphLimit = readPdfTrueTypeGlyphCount(font);
  const covered = new Set<number>();
  for (let code = 0; code < mapping.length / 2; code++) {
    const glyph = mapping[code * 2] * 256 + mapping[code * 2 + 1];
    // An unusable slot affects that character only. Never manufacture a
    // glyph or assume CID == glyph id for a stream-backed mapping.
    if (glyph > 0 && glyph < glyphLimit) covered.add(code);
  }
  return covered;
}












export function readPdfSimpleTrueTypeGlyphCodes(
  font: Uint8Array,
  unicodeOf: (code: number) => string | undefined,
): ReadonlySet<number> | null {
  try {
    const glyphLimit = readPdfTrueTypeGlyphCount(font);
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    const count = view.getUint16(4);
    let cmap: { offset: number; length: number } | null = null;
    for (let i = 0; i < count; i++) {
      const entry = 12 + i * 16;
      if (view.getUint32(entry) !== 0x636d6170) continue; // cmap
      cmap = { offset: view.getUint32(entry + 8), length: view.getUint32(entry + 12) };
    }
    if (!cmap || cmap.length < 4 || cmap.offset + cmap.length > font.length) return null;
    const end = cmap.offset + cmap.length;
    const u16 = (at: number) => { if (at + 2 > end) throw new Error("cmap"); return view.getUint16(at); };
    const u32 = (at: number) => { if (at + 4 > end) throw new Error("cmap"); return view.getUint32(at); };
    /** One subtable as a lookup, or null for a format this does not read. */
    const readSubtable = (at: number): ((code: number) => number) | null => {
      const format = u16(at);
      if (format === 0) return (code) => (code >= 0 && code < 256 ? font[at + 6 + code] ?? 0 : 0);
      if (format === 6) {
        const first = u16(at + 6), entries = u16(at + 8);
        return (code) => (code >= first && code < first + entries ? u16(at + 10 + (code - first) * 2) : 0);
      }
      if (format === 4) {
        const segments = u16(at + 6) / 2;
        const ends = at + 14, starts = ends + segments * 2 + 2, deltas = starts + segments * 2, ranges = deltas + segments * 2;
        return (code) => {
          for (let s = 0; s < segments; s++) {
            if (code > u16(ends + s * 2)) continue;
            const start = u16(starts + s * 2);
            if (code < start) return 0;
            const delta = u16(deltas + s * 2), range = u16(ranges + s * 2);
            if (range === 0) return (code + delta) & 0xffff;
            const glyph = u16(ranges + s * 2 + range + (code - start) * 2);
            return glyph === 0 ? 0 : (glyph + delta) & 0xffff;
          }
          return 0;
        };
      }
      if (format === 12) {
        const groups = u32(at + 12);
        return (code) => {
          for (let g = 0; g < groups; g++) {
            const group = at + 16 + g * 12, first = u32(group), last = u32(group + 4);
            if (code >= first && code <= last) return u32(group + 8) + (code - first);
          }
          return 0;
        };
      }
      return null;
    };
    const byCode: Array<(code: number) => number> = [];
    const byUnicode: Array<(code: number) => number> = [];
    const tables = u16(cmap.offset + 2);
    for (let t = 0; t < tables; t++) {
      const record = cmap.offset + 4 + t * 8;
      const platform = u16(record), encoding = u16(record + 2);
      const symbolOrMac = (platform === 1 && encoding === 0) || (platform === 3 && encoding === 0);
      const unicode = platform === 0 || (platform === 3 && (encoding === 1 || encoding === 10));
      if (!symbolOrMac && !unicode) continue;
      const lookup = readSubtable(cmap.offset + u32(record + 4));
      if (!lookup) return null;
      if (platform === 3 && encoding === 0) byCode.push((code) => lookup(0xf000 + code) || lookup(code));
      else if (symbolOrMac) byCode.push(lookup);
      else byUnicode.push(lookup);
    }
    if (!byCode.length && !byUnicode.length) return null;
    const drawn = (glyph: number) => glyph > 0 && glyph < glyphLimit;
    const covered = new Set<number>();
    for (let code = 0; code < 256; code++) {
      if (byCode.some((lookup) => drawn(lookup(code)))) { covered.add(code); continue; }
      const text = unicodeOf(code);
      const point = text && [...text].length === 1 ? text.codePointAt(0)! : null;
      if (point !== null && byUnicode.some((lookup) => drawn(lookup(point)))) covered.add(code);
    }
    return covered;
  } catch {
    return null;
  }
}

/** One table of a TrueType program by its tag, or null where the directory
 * does not hold it whole. */
function trueTypeTable(font: Uint8Array, view: DataView, tag: number): { offset: number; length: number } | null {
  const count = view.getUint16(4);
  for (let i = 0; i < count; i++) {
    const entry = 12 + i * 16;
    if (view.getUint32(entry) !== tag) continue;
    const offset = view.getUint32(entry + 8), length = view.getUint32(entry + 12);
    return offset >= 12 + count * 16 && offset + length <= font.length ? { offset, length } : null;
  }
  return null;
}





















export function readPdfTrueTypeGlyphInk(font: Uint8Array): ((glyph: number) => boolean) | null {
  try {
    const glyphLimit = readPdfTrueTypeGlyphCount(font);
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    const head = trueTypeTable(font, view, 0x68656164); // head
    const loca = trueTypeTable(font, view, 0x6c6f6361); // loca
    const glyf = trueTypeTable(font, view, 0x676c7966); // glyf
    if (!head || !loca || !glyf || head.length < 54) return null;
    const format = view.getInt16(head.offset + 50);
    if (format !== 0 && format !== 1) return null;
    const width = format === 1 ? 4 : 2;
    if (loca.length < (glyphLimit + 1) * width) return null;
    const at = (glyph: number) => format === 1
      ? view.getUint32(loca.offset + glyph * 4)
      : view.getUint16(loca.offset + glyph * 2) * 2;
    const inked = (glyph: number, depth: number): boolean => {
      if (!Number.isInteger(glyph) || glyph < 0 || glyph >= glyphLimit) return false;
      const start = at(glyph), end = at(glyph + 1);
      // A glyph's record opens with its contour count (none is no outline,
      // negative is a composite) and the box of its outline.
      if (end <= start || end > glyf.length || end - start < 10) return false;
      const record = glyf.offset + start, limit = glyf.offset + end;
      const contours = view.getInt16(record);
      if (contours === 0) return false;
      if (view.getInt16(record + 6) > view.getInt16(record + 2) &&
        view.getInt16(record + 8) > view.getInt16(record + 4)) return true;
      try {
        return contours > 0
          ? trueTypeSimpleGlyphHasExtent(view, record + 10, limit, contours)
          : depth < 4 && trueTypeCompositeParts(view, record + 10, limit).some((part) => inked(part, depth + 1));
      } catch {
        // A record that cannot be read through says nothing: not ink, so the
        // font is read as it always was.
        return false;
      }
    };
    return (glyph) => inked(glyph, 0);
  } catch {
    return null;
  }
}

/** Whether a simple glyph's points span some width and some height. */
function trueTypeSimpleGlyphHasExtent(view: DataView, from: number, limit: number, contours: number): boolean {
  const need = (at: number, bytes: number) => { if (at + bytes > limit) throw new Error("glyf"); };
  let at = from;
  need(at, contours * 2 + 2);
  const points = view.getUint16(at + (contours - 1) * 2) + 1;
  at += contours * 2;
  at += 2 + view.getUint16(at);
  const flags: number[] = [];
  while (flags.length < points) {
    need(at, 1);
    const flag = view.getUint8(at++);
    flags.push(flag);
    if (flag & 0x08) {
      need(at, 1);
      for (let repeat = view.getUint8(at++); repeat > 0 && flags.length < points; repeat -= 1) flags.push(flag);
    }
  }
  const spans = (short: number, same: number) => {
    let value = 0, low = 0, high = 0, first = true;
    for (const flag of flags) {
      if (flag & short) {
        need(at, 1);
        const delta = view.getUint8(at++);
        value += flag & same ? delta : -delta;
      } else if (!(flag & same)) {
        need(at, 2);
        value += view.getInt16(at);
        at += 2;
      }
      if (first) { low = high = value; first = false; }
      else { low = Math.min(low, value); high = Math.max(high, value); }
    }
    return high > low;
  };
  // Both are read: the y coordinates follow the x ones.
  const wide = spans(0x02, 0x10);
  const tall = spans(0x04, 0x20);
  return wide && tall;
}

/** The glyphs a composite glyph is made of. */
function trueTypeCompositeParts(view: DataView, from: number, limit: number): number[] {
  const parts: number[] = [];
  let at = from;
  for (let guard = 0; guard < 64; guard += 1) {
    if (at + 4 > limit) throw new Error("glyf");
    const flags = view.getUint16(at);
    parts.push(view.getUint16(at + 2));
    at += 4 + (flags & 0x0001 ? 4 : 2) + (flags & 0x0008 ? 2 : flags & 0x0040 ? 4 : flags & 0x0080 ? 8 : 0);
    if (!(flags & 0x0020)) break;
  }
  return parts;
}

/**
 * The characters a TrueType program's own Unicode `cmap` names each glyph
 * by: glyph id to its code points, lowest first. A subset keeps the names of
 * the glyphs it kept, so the program still says what a glyph is where the
 * font's ToUnicode does not. Null where the program holds no Unicode subtable
 * this reads.
 */
export function readPdfTrueTypeGlyphCharacters(font: Uint8Array): ReadonlyMap<number, readonly number[]> | null {
  try {
    const glyphLimit = readPdfTrueTypeGlyphCount(font);
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    const cmap = trueTypeTable(font, view, 0x636d6170); // cmap
    if (!cmap || cmap.length < 4) return null;
    const end = cmap.offset + cmap.length;
    const u16 = (at: number) => { if (at + 2 > end) throw new Error("cmap"); return view.getUint16(at); };
    const u32 = (at: number) => { if (at + 4 > end) throw new Error("cmap"); return view.getUint32(at); };
    const named = new Map<number, Set<number>>();
    let budget = 400_000;
    const name = (point: number, glyph: number) => {
      if ((budget -= 1) < 0) throw new Error("cmap");
      if (glyph <= 0 || glyph >= glyphLimit || point > 0x10ffff || (point >= 0xd800 && point <= 0xdfff)) return;
      let points = named.get(glyph);
      if (!points) named.set(glyph, points = new Set());
      points.add(point);
    };
    let read = false;
    const tables = u16(cmap.offset + 2);
    for (let t = 0; t < tables; t++) {
      const record = cmap.offset + 4 + t * 8;
      const platform = u16(record), encoding = u16(record + 2);
      if (!(platform === 0 || (platform === 3 && (encoding === 1 || encoding === 10)))) continue;
      const at = cmap.offset + u32(record + 4);
      const format = u16(at);
      if (format === 4) {
        const segments = u16(at + 6) / 2;
        const ends = at + 14, starts = ends + segments * 2 + 2, deltas = starts + segments * 2, ranges = deltas + segments * 2;
        for (let s = 0; s < segments; s++) {
          const last = u16(ends + s * 2), first = u16(starts + s * 2);
          const delta = u16(deltas + s * 2), range = u16(ranges + s * 2);
          for (let point = first; point <= last && point < 0xffff; point++) {
            if (range === 0) { name(point, (point + delta) & 0xffff); continue; }
            const glyph = u16(ranges + s * 2 + range + (point - first) * 2);
            if (glyph !== 0) name(point, (glyph + delta) & 0xffff);
          }
        }
        read = true;
      } else if (format === 12) {
        const groups = u32(at + 12);
        for (let g = 0; g < groups; g++) {
          const group = at + 16 + g * 12, first = u32(group), last = u32(group + 4), start = u32(group + 8);
          for (let point = first; point <= last; point++) name(point, start + (point - first));
        }
        read = true;
      } else if (format === 6) {
        const first = u16(at + 6), entries = u16(at + 8);
        for (let i = 0; i < entries; i++) name(first + i, u16(at + 10 + i * 2));
        read = true;
      }
    }
    if (!read) return null;
    return new Map([...named].map(([glyph, points]) => [glyph, Object.freeze([...points].sort((a, b) => a - b))]));
  } catch {
    return null;
  }
}

/**
 * What a glyph the font's ToUnicode calls white space really is, by the
 * program's own name for it: its lowest code point that is a character a
 * reader shows (not white space, a control or a format character, nor one
 * of the private use area). Null where the program names it no such thing.
 */
export function pdfTrueTypeGlyphName(points: readonly number[] | undefined): string | null {
  for (const point of points ?? []) {
    const text = String.fromCodePoint(point);
    if (/^[\s\p{C}]$/u.test(text)) continue;
    return text;
  }
  return null;
}
