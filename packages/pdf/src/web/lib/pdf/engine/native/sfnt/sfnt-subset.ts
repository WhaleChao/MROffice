/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A small TrueType program holding only the glyphs a text needs, taken from
 * a served slice (sfnt-reader.ts), at the face's weight when the slice is a
 * variable font (sfnt-variation.ts). It is what the PDF embeds as the
 * FontFile2 of a CIDFontType2 whose CIDs are its glyph indices.
 *
 * Glyph 0 is the slice's .notdef; the text's glyphs follow in the order the
 * text first names them, and the components of composite glyphs after them.
 * The outlines keep the slice's points, on-curve flags and contours exactly;
 * at a location they are the instance rounded once, so the server (which
 * draws the slice at the same location with fontTools) finds every point
 * within rounding of its own. Hinting is not carried: the instructions of a
 * variable glyph no longer describe its instance, and a PDF viewer draws
 * unhinted TrueType well.
 */

import {
  COMPOSITE_ARG_1_AND_2_ARE_WORDS,
  COMPOSITE_ARGS_ARE_XY_VALUES,
  COMPOSITE_MORE_COMPONENTS,
  COMPOSITE_WE_HAVE_INSTRUCTIONS,
  sfntFail,
  type SfntComponent,
  type SfntProgram,
} from "@mrnote/core/web/lib/pdf/engine/native/sfnt/sfnt-reader.ts";
import { createSfntInstancer, normalizeSfntLocation, sfntVariationAxes } from "@mrnote/core/web/lib/pdf/engine/native/sfnt/sfnt-variation.ts";

/** fontTools' otRound. */
export function otRound(value: number): number {
  return Math.floor(value + 0.5);
}

interface SubsetSimple {
  readonly kind: "simple";
  readonly endPoints: readonly number[];
  readonly onCurve: readonly boolean[];
  readonly xs: readonly number[];
  readonly ys: readonly number[];
  readonly overlap: boolean;
}

interface SubsetComposite {
  readonly kind: "composite";
  readonly components: readonly SfntComponent[];
}

interface SubsetGlyph {
  readonly sourceId: number;
  readonly outline: SubsetSimple | SubsetComposite | null;
  readonly advance: number;
  /** The left side bearing to write; null: from the instance's left
   * phantom point (`leftSideX`). */
  readonly leftSideBearing: number | null;
  readonly leftSideX: number;
  /** The source's stored bounds, kept for an outline copied unchanged. */
  readonly storedBounds: readonly [number, number, number, number] | null;
}

export interface PdfNativeSfntSubsetGlyph {
  /** The glyph's index in the subset: its CID in the PDF. */
  readonly cid: number;
  readonly text: string;
  readonly advance: number;
}

/** A glyph `buildPdfNativeSfntSubset` adds after the subset's own. */
export type PdfNativeSfntGlyphCopy =
  | Readonly<{ glyph: number; empty: boolean }>
  | Readonly<{ components: readonly Readonly<{ glyph: number; x: number; y: number }>[]; advance: number }>;

export interface PdfNativeSfntSubset {
  readonly bytes: Uint8Array;
  readonly unitsPerEm: number;
  readonly ascent: number;
  readonly descent: number;
  /** One per distinct character asked for, in first-use order. */
  readonly glyphs: readonly PdfNativeSfntSubsetGlyph[];
  /** Each glyph asked for by its index (a shaped run's), by that index: its
   * CID in the subset and its advance. */
  readonly shapedGlyphs?: ReadonlyMap<number, Readonly<{ cid: number; advance: number }>>;
  /** Each copy asked for (`glyphCopies`), in that order: its CID and advance. */
  readonly glyphCopies?: readonly Readonly<{ cid: number; advance: number }>[];
}

class Writer {
  private readonly parts: number[] = [];
  u8(value: number) { this.parts.push(value & 0xff); }
  i8(value: number) { this.u8(value < 0 ? value + 256 : value); }
  u16(value: number) { this.parts.push((value >> 8) & 0xff, value & 0xff); }
  i16(value: number) {
    if (!Number.isInteger(value) || value < -32_768 || value > 32_767) sfntFail("invalid-font", "A coordinate does not fit a TrueType glyph.");
    this.u16(value < 0 ? value + 65_536 : value);
  }
  u32(value: number) { this.parts.push((value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff); }
  bytes(values: Uint8Array | readonly number[]) { for (const value of values) this.parts.push(value); }
  pad4() { while (this.parts.length % 4) this.parts.push(0); }
  get length() { return this.parts.length; }
  done(): Uint8Array { return Uint8Array.from(this.parts); }
}

function encodeSimple(glyph: SubsetSimple, bounds: readonly [number, number, number, number]): Uint8Array {
  const out = new Writer();
  out.i16(glyph.endPoints.length);
  bounds.forEach((value) => out.i16(value));
  glyph.endPoints.forEach((value) => out.u16(value));
  out.u16(0); // no instructions
  const flags: number[] = [];
  const xBytes = new Writer();
  const yBytes = new Writer();
  let lastX = 0, lastY = 0;
  glyph.xs.forEach((x, index) => {
    const y = glyph.ys[index];
    let flag = glyph.onCurve[index] ? 0x01 : 0;
    if (index === 0 && glyph.overlap) flag |= 0x40;
    const dx = x - lastX;
    const dy = y - lastY;
    if (dx === 0) flag |= 0x10;
    else if (dx >= -255 && dx <= 255) {
      flag |= 0x02;
      if (dx > 0) flag |= 0x10;
      xBytes.u8(Math.abs(dx));
    } else xBytes.i16(dx);
    if (dy === 0) flag |= 0x20;
    else if (dy >= -255 && dy <= 255) {
      flag |= 0x04;
      if (dy > 0) flag |= 0x20;
      yBytes.u8(Math.abs(dy));
    } else yBytes.i16(dy);
    flags.push(flag);
    lastX = x;
    lastY = y;
  });
  for (let index = 0; index < flags.length; ) {
    let run = 1;
    while (index + run < flags.length && flags[index + run] === flags[index] && run < 256) run += 1;
    if (run > 1) {
      out.u8(flags[index] | 0x08);
      out.u8(run - 1);
    } else out.u8(flags[index]);
    index += run;
  }
  out.bytes(xBytes.done());
  out.bytes(yBytes.done());
  out.pad4();
  return out.done();
}

function encodeComposite(
  glyph: SubsetComposite,
  bounds: readonly [number, number, number, number],
  remap: (sourceId: number) => number,
): Uint8Array {
  const out = new Writer();
  out.i16(-1);
  bounds.forEach((value) => out.i16(value));
  glyph.components.forEach((component, index) => {
    let flags = component.flags & ~(COMPOSITE_WE_HAVE_INSTRUCTIONS | COMPOSITE_MORE_COMPONENTS | COMPOSITE_ARG_1_AND_2_ARE_WORDS);
    if (index < glyph.components.length - 1) flags |= COMPOSITE_MORE_COMPONENTS;
    const xy = (component.flags & COMPOSITE_ARGS_ARE_XY_VALUES) !== 0;
    const fitsByte = xy
      ? component.arg1 >= -128 && component.arg1 <= 127 && component.arg2 >= -128 && component.arg2 <= 127
      : component.arg1 <= 255 && component.arg2 <= 255;
    if (!fitsByte) flags |= COMPOSITE_ARG_1_AND_2_ARE_WORDS;
    out.u16(flags);
    out.u16(remap(component.glyphIndex));
    if (fitsByte) {
      if (xy) { out.i8(component.arg1); out.i8(component.arg2); }
      else { out.u8(component.arg1); out.u8(component.arg2); }
    } else if (xy) { out.i16(component.arg1); out.i16(component.arg2); }
    else { out.u16(component.arg1); out.u16(component.arg2); }
    component.transform.forEach((value) => out.i16(value));
  });
  out.pad4();
  return out.done();
}

/** The points of a glyph drawn with its components placed, for bounds. */
function decomposedPoints(
  glyph: SubsetGlyph,
  glyphs: ReadonlyMap<number, SubsetGlyph>,
  depth = 0,
): [number, number][] {
  if (depth > 16) sfntFail("invalid-font", "Composite glyphs nest too deeply.");
  const outline = glyph.outline;
  if (!outline) return [];
  if (outline.kind === "simple") return outline.xs.map((x, index) => [x, outline.ys[index]]);
  const points: [number, number][] = [];
  for (const component of outline.components) {
    const child = glyphs.get(component.glyphIndex);
    if (!child) sfntFail("invalid-font", "A component glyph is missing.");
    const [a, b, c, d] = component.transform.length === 1
      ? [component.transform[0] / 16384, 0, 0, component.transform[0] / 16384]
      : component.transform.length === 2
        ? [component.transform[0] / 16384, 0, 0, component.transform[1] / 16384]
        : component.transform.length === 4
          ? component.transform.map((value) => value / 16384)
          : [1, 0, 0, 1];
    if (!(component.flags & COMPOSITE_ARGS_ARE_XY_VALUES)) {
      sfntFail("unsupported-outlines", "Composite glyphs anchored by point are not embedded.");
    }
    for (const [x, y] of decomposedPoints(child, glyphs, depth + 1)) {
      points.push([x * a + y * c + component.arg1, x * b + y * d + component.arg2]);
    }
  }
  return points;
}

function intBounds(points: readonly (readonly [number, number])[]): [number, number, number, number] {
  if (points.length === 0) return [0, 0, 0, 0];
  let xMin = Infinity, yMin = Infinity, xMax = -Infinity, yMax = -Infinity;
  for (const [x, y] of points) {
    xMin = Math.min(xMin, x); xMax = Math.max(xMax, x);
    yMin = Math.min(yMin, y); yMax = Math.max(yMax, y);
  }
  return [otRound(xMin), otRound(yMin), otRound(xMax), otRound(yMax)];
}

function tableChecksum(bytes: Uint8Array): number {
  let sum = 0;
  for (let offset = 0; offset < bytes.length; offset += 4) {
    const word = ((bytes[offset] ?? 0) << 24) | ((bytes[offset + 1] ?? 0) << 16) | ((bytes[offset + 2] ?? 0) << 8) | (bytes[offset + 3] ?? 0);
    sum = (sum + (word >>> 0)) >>> 0;
  }
  return sum;
}

function cmapTable(entries: readonly (readonly [number, number])[]): Uint8Array {
  const sorted = [...entries].sort((left, right) => left[0] - right[0]);
  const bmp = sorted.filter(([code]) => code <= 0xffff);
  const subtables: Array<{ platform: number; encoding: number; bytes: Uint8Array }> = [];
  // Format 4, one segment per character, then the terminal segment.
  const format4 = new Writer();
  const segments = [...bmp.map(([code, glyph]) => [code, code, (glyph - code) & 0xffff] as const), [0xffff, 0xffff, 1] as const];
  const segX2 = segments.length * 2;
  let search = 1, selector = 0;
  while (search * 2 <= segments.length) { search *= 2; selector += 1; }
  format4.u16(4);
  format4.u16(16 + segments.length * 8);
  format4.u16(0);
  format4.u16(segX2);
  format4.u16(search * 2);
  format4.u16(selector);
  format4.u16(segX2 - search * 2);
  segments.forEach(([, end]) => format4.u16(end));
  format4.u16(0);
  segments.forEach(([start]) => format4.u16(start));
  segments.forEach(([, , delta]) => format4.u16(delta));
  segments.forEach(() => format4.u16(0));
  subtables.push({ platform: 3, encoding: 1, bytes: format4.done() });
  if (sorted.some(([code]) => code > 0xffff)) {
    const format12 = new Writer();
    format12.u16(12);
    format12.u16(0);
    format12.u32(16 + sorted.length * 12);
    format12.u32(0);
    format12.u32(sorted.length);
    sorted.forEach(([code, glyph]) => { format12.u32(code); format12.u32(code); format12.u32(glyph); });
    subtables.push({ platform: 3, encoding: 10, bytes: format12.done() });
  }
  const out = new Writer();
  out.u16(0);
  out.u16(subtables.length);
  let offset = 4 + subtables.length * 8;
  for (const subtable of subtables) {
    out.u16(subtable.platform);
    out.u16(subtable.encoding);
    out.u32(offset);
    offset += subtable.bytes.length;
  }
  for (const subtable of subtables) out.bytes(subtable.bytes);
  return out.done();
}

function patched(source: Uint8Array, minimumLength: number, edits: (view: DataView) => void): Uint8Array {
  if (source.length < minimumLength) sfntFail("invalid-font", "A required table is too short.");
  const copy = Uint8Array.from(source);
  edits(new DataView(copy.buffer));
  return copy;
}

/**
 * The subset program for `text` from `program`. `location` (user values,
 * e.g. `{ wght: 700 }`) sets a variable program's instance; leave it out for
 * the program's own default. Fails with `missing-glyph` when a character has
 * no glyph in the program.
 */
export function buildPdfNativeSfntSubset(input: {
  readonly program: SfntProgram;
  readonly text: string;
  readonly location?: Readonly<Record<string, number>> | null;
  /** A larger bound for the subset a page keeps of a font from the reader's
   * computer (its common characters; page-fonts-browser.ts), at most 8,000. */
  readonly maxCharacters?: number;
  /** Glyphs asked for by index beside the characters: a run HarfBuzz shaped
   * (an Arabic word's joined forms, a Devanagari conjunct), which no
   * character of the cmap names. */
  readonly glyphIds?: readonly number[];
  /**
   * More glyphs after the subset's own: a shaped glyph again under a CID of
   * its own (one glyph standing for other letters elsewhere), `empty` a
   * glyph with no outline and no advance (an invisible carrier of the marks
   * a reader puts after its base letter), or a composite of shaped glyphs
   * placed as the shaper placed them (a cluster drawn as one glyph: a
   * letter and its dots, a conjunct and its vowel sign), its advance the
   * cluster's. Every glyph a copy names must be one of `glyphIds`.
   */
  readonly glyphCopies?: readonly PdfNativeSfntGlyphCopy[];
}): PdfNativeSfntSubset {
  const { program } = input;
  const characters = [...new Set([...input.text])];
  const shapedIds = [...new Set(input.glyphIds ?? [])];
  const maxCharacters = Math.min(8_000, input.maxCharacters ?? 2_000);
  if ((characters.length === 0 && shapedIds.length === 0) || characters.length > maxCharacters ||
    shapedIds.length > maxCharacters) sfntFail("invalid-font", "The subset text is empty or too long.");
  if (shapedIds.some((id) => !Number.isInteger(id) || id <= 0 || id >= program.advances.length)) {
    sfntFail("missing-glyph", "A shaped glyph is not in the program.");
  }
  const copies = input.glyphCopies ?? [];
  const namesShaped = (copy: PdfNativeSfntGlyphCopy) => "components" in copy
    ? copy.components.length > 0 && copy.components.length <= 64 && Number.isInteger(copy.advance) &&
      copy.advance >= 0 && copy.advance <= 65_535 && copy.components.every((component) =>
        shapedIds.includes(component.glyph) && Number.isInteger(component.x) && Number.isInteger(component.y) &&
        Math.abs(component.x) <= 32_767 && Math.abs(component.y) <= 32_767)
    : shapedIds.includes(copy.glyph);
  if (copies.length > maxCharacters || !copies.every(namesShaped)) {
    sfntFail("invalid-font", "A glyph copy names no shaped glyph.");
  }
  const location = input.location && Object.keys(input.location).length > 0 ? input.location : null;
  const instancer = location
    ? createSfntInstancer(program, normalizeSfntLocation(program, location))
    : null;
  if (location && sfntVariationAxes(program).length === 0) sfntFail("unsupported-variation");
  const order: number[] = [0];
  const ordered = new Set(order);
  const characterGlyphs = characters.map((character) => {
    const glyphId = program.cmap.get(character.codePointAt(0)!);
    if (glyphId === undefined || glyphId === 0) sfntFail("missing-glyph", `The program has no glyph for U+${character.codePointAt(0)!.toString(16).toUpperCase()}.`);
    if (!ordered.has(glyphId)) { ordered.add(glyphId); order.push(glyphId); }
    return glyphId;
  });
  for (const glyphId of shapedIds) {
    if (!ordered.has(glyphId)) { ordered.add(glyphId); order.push(glyphId); }
  }
  const glyphs = new Map<number, SubsetGlyph>();
  for (let index = 0; index < order.length; index += 1) {
    if (order.length > Math.max(4_096, maxCharacters * 2)) sfntFail("invalid-font", "The subset grows too large.");
    const sourceId = order[index];
    const stored = program.glyph(sourceId);
    if (stored?.kind === "composite") {
      for (const component of stored.components) {
        if (!ordered.has(component.glyphIndex)) { ordered.add(component.glyphIndex); order.push(component.glyphIndex); }
      }
    }
    if (!instancer) {
      glyphs.set(sourceId, Object.freeze({
        sourceId,
        outline: stored === null ? null
          : stored.kind === "simple"
            ? Object.freeze({ kind: "simple" as const, endPoints: stored.endPoints, onCurve: stored.onCurve, xs: stored.xs, ys: stored.ys, overlap: stored.overlap })
            : Object.freeze({ kind: "composite" as const, components: stored.components }),
        advance: program.advances[sourceId],
        leftSideBearing: program.leftSideBearings[sourceId],
        leftSideX: 0,
        storedBounds: stored?.bbox ?? null,
      }));
      continue;
    }
    const instance = instancer.instance(sourceId);
    const glyph = instance.glyph;
    let outline: SubsetGlyph["outline"] = null;
    if (glyph?.kind === "simple") {
      outline = Object.freeze({
        kind: "simple" as const,
        endPoints: glyph.endPoints,
        onCurve: glyph.onCurve,
        xs: glyph.xs.map(otRound),
        ys: glyph.ys.map(otRound),
        overlap: glyph.overlap,
      });
    } else if (glyph?.kind === "composite") {
      outline = Object.freeze({
        kind: "composite" as const,
        components: glyph.components.map((component, componentIndex) =>
          component.flags & COMPOSITE_ARGS_ARE_XY_VALUES
            ? Object.freeze({ ...component, arg1: otRound(glyph.dxs[componentIndex]), arg2: otRound(glyph.dys[componentIndex]) })
            : component),
      });
    }
    glyphs.set(sourceId, Object.freeze({
      sourceId,
      outline,
      advance: otRound(instance.advance),
      // fontTools: otRound(xMin - leftSideX), xMin the rounded outline's.
      leftSideBearing: null,
      leftSideX: instance.leftSideX,
      storedBounds: null,
    }));
  }
  const newIndex = new Map(order.map((sourceId, index) => [sourceId, index]));
  const remap = (sourceId: number) => newIndex.get(sourceId) ?? sfntFail("invalid-font", "A component glyph is missing.");
  const records: Uint8Array[] = [];
  const metrics: Array<[number, number]> = [];
  let fontBounds: [number, number, number, number] | null = null;
  const copyOf = (copy: PdfNativeSfntGlyphCopy): SubsetGlyph => "components" in copy
    ? Object.freeze({
      sourceId: copy.components[0].glyph,
      outline: Object.freeze({ kind: "composite" as const, components: copy.components.map((component) => Object.freeze({
        flags: COMPOSITE_ARGS_ARE_XY_VALUES, glyphIndex: component.glyph, arg1: component.x, arg2: component.y, transform: [],
      })) }),
      advance: copy.advance,
      leftSideBearing: null,
      leftSideX: 0,
      storedBounds: null,
    })
    : copy.empty
      ? Object.freeze({ sourceId: copy.glyph, outline: null, advance: 0, leftSideBearing: 0, leftSideX: 0, storedBounds: null })
      : glyphs.get(copy.glyph)!;
  for (const glyph of [...order.map((sourceId) => glyphs.get(sourceId)!), ...copies.map(copyOf)]) {
    const outline = glyph.outline;
    const bounds = glyph.storedBounds ?? (outline ? intBounds(decomposedPoints(glyph, glyphs)) : [0, 0, 0, 0] as [number, number, number, number]);
    if (outline) {
      fontBounds = fontBounds
        ? [Math.min(fontBounds[0], bounds[0]), Math.min(fontBounds[1], bounds[1]), Math.max(fontBounds[2], bounds[2]), Math.max(fontBounds[3], bounds[3])]
        : [...bounds] as [number, number, number, number];
    }
    records.push(outline === null ? new Uint8Array(0)
      : outline.kind === "simple" ? encodeSimple(outline, bounds)
        : encodeComposite(outline, bounds, remap));
    const lsb = glyph.leftSideBearing ?? otRound((outline ? bounds[0] : 0) - glyph.leftSideX);
    if (!Number.isInteger(glyph.advance) || glyph.advance < 0 || glyph.advance > 65_535) sfntFail("invalid-font", "A glyph advance does not fit hmtx.");
    metrics.push([glyph.advance, lsb]);
  }
  const loca = new Writer();
  const glyf = new Writer();
  for (const record of records) {
    loca.u32(glyf.length);
    glyf.bytes(record);
  }
  loca.u32(glyf.length);
  const hmtx = new Writer();
  metrics.forEach(([advance, lsb]) => { hmtx.u16(advance); hmtx.i16(lsb); });
  const bounds = fontBounds ?? [0, 0, 0, 0];
  const head = patched(program.tables.get("head") ?? sfntFail("invalid-font"), 54, (view) => {
    view.setUint32(8, 0);
    view.setInt16(36, bounds[0]);
    view.setInt16(38, bounds[1]);
    view.setInt16(40, bounds[2]);
    view.setInt16(42, bounds[3]);
    view.setInt16(50, 1);
  });
  const hhea = patched(program.tables.get("hhea") ?? sfntFail("invalid-font"), 36, (view) => {
    view.setUint16(10, Math.max(0, ...metrics.map(([advance]) => advance)));
    view.setUint16(34, metrics.length);
  });
  const maxp = patched(program.tables.get("maxp") ?? sfntFail("invalid-font"), 6, (view) => {
    view.setUint16(4, order.length + copies.length);
  });
  const sourcePost = program.tables.get("post") ?? sfntFail("invalid-font");
  const post = patched(sourcePost.subarray(0, 32), 32, (view) => {
    view.setUint32(0, 0x00030000);
  });
  const tables = new Map<string, Uint8Array>([
    ["cmap", cmapTable(characters.map((character, index) => [character.codePointAt(0)!, remap(characterGlyphs[index])] as const))],
    ["glyf", glyf.done()],
    ["head", head],
    ["hhea", hhea],
    ["hmtx", hmtx.done()],
    ["loca", loca.done()],
    ["maxp", maxp],
    ["post", post],
  ]);
  const os2 = program.tables.get("OS/2");
  if (os2) tables.set("OS/2", Uint8Array.from(os2));
  const name = program.tables.get("name");
  if (name) tables.set("name", Uint8Array.from(name));
  const tags = [...tables.keys()].sort();
  const directory = new Writer();
  let search = 1, selector = 0;
  while (search * 2 <= tags.length) { search *= 2; selector += 1; }
  directory.u32(0x00010000);
  directory.u16(tags.length);
  directory.u16(search * 16);
  directory.u16(selector);
  directory.u16(tags.length * 16 - search * 16);
  let offset = 12 + tags.length * 16;
  const body = new Writer();
  for (const tag of tags) {
    const bytes = tables.get(tag)!;
    for (const character of tag) directory.u8(character.charCodeAt(0));
    directory.u32(tableChecksum(bytes));
    directory.u32(offset);
    directory.u32(bytes.length);
    body.bytes(bytes);
    body.pad4();
    offset = 12 + tags.length * 16 + body.length;
  }
  const font = new Uint8Array(directory.length + body.length);
  font.set(directory.done(), 0);
  font.set(body.done(), directory.length);
  const headOffset = 12 + tags.length * 16 + tags.slice(0, tags.indexOf("head")).reduce((sum, tag) => sum + Math.ceil(tables.get(tag)!.length / 4) * 4, 0);
  const adjustment = (0xb1b0afba - tableChecksum(font)) >>> 0;
  new DataView(font.buffer).setUint32(headOffset + 8, adjustment);
  return Object.freeze({
    bytes: font,
    unitsPerEm: program.unitsPerEm,
    ascent: program.ascent,
    descent: program.descent,
    glyphs: Object.freeze(characters.map((character, index) => Object.freeze({
      cid: remap(characterGlyphs[index]),
      text: character,
      advance: metrics[remap(characterGlyphs[index])][0],
    }))),
    ...(shapedIds.length ? { shapedGlyphs: new Map(shapedIds.map((glyphId) =>
      [glyphId, Object.freeze({ cid: remap(glyphId), advance: metrics[remap(glyphId)][0] })] as const)) } : {}),
    ...(copies.length ? { glyphCopies: Object.freeze(copies.map((_copy, index) =>
      Object.freeze({ cid: order.length + index, advance: metrics[order.length + index][0] }))) } : {}),
  });
}
