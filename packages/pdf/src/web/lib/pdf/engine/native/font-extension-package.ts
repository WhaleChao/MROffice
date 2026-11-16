/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { LATIN1_EXACT } from "./latin1";
import {
  CustomFontEmbedder,
  CustomFontSubsetEmbedder,
  PDFContext,
  PDFName,
  PDFRawStream,
  PDFRef,
  PDFString,
  type PDFObject,
} from "pdf-lib";

import {
  PDF_UNICODE_FALLBACK_FONT_SHA256,
  verifyPdfUnicodeFallbackFontBytes,
} from "../../unicode-font-asset";
import { unzlibSync } from "fflate";
import { pdfNativeFontRepertoireTiers } from "./font-repertoire";
import { alignSubsetGlyphRecords } from "../../fontkit-subset-glyph-records";
import { pdfNativeSparseTrueTypeSubset } from "./truetype-sparse-subset";
import type { PdfNativeSfntSubset } from "./sfnt/sfnt-subset";

export const MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION =
  "mrai.native-font-extension-package.v1" as const;

const RESOURCE_NAME_PATTERN = /^[A-Za-z][A-Za-z0-9_.+\-]{0,63}$/;
const MAX_TEXT_CODE_UNITS = 2_000;
const MAX_FONT_OBJECTS = 16;
const MAX_FONT_PACKAGE_BYTES = 16 * 1024 * 1024;
const MAX_OBJECT_NUMBER = 2_147_483_600;

export type PdfNativeFontExtensionPackageErrorCode =
  "invalid-request" | "invalid-font" | "unsupported-glyph" | "package-invalid";

export class PdfNativeFontExtensionPackageError extends Error {
  constructor(readonly code: PdfNativeFontExtensionPackageErrorCode) {
    super(`PDF native font extension package failed: ${code}.`);
    this.name = "PdfNativeFontExtensionPackageError";
  }
}

export interface PdfNativeFontExtensionObject {
  readonly objectNumber: number;
  readonly generation: 0;
  readonly sha256: string;
  readonly bytes: Uint8Array;
}

/** The face of a served Office substitute a package embeds a subset of
 * (lib/office-fonts/embeddable-faces.ts): the server proves the package's
 * glyphs against the slice `fontSha256` names, drawn at this face's weight. */
export interface PdfNativeEmbeddedFontFace {
  readonly family: string;
  readonly weight: number;
  readonly italic: boolean;
  /** A font from the reader's computer the page keeps (page-fonts.ts): the
   * server proves the glyphs against the page's copy of `fontSha256`. */
  readonly page?: true;
}

/** The fixed name every program this module embeds from the bundled
 * fallback goes under, subset or whole (no random subset tag): the name the
 * server validates. */
export const PDF_NATIVE_FALLBACK_FONT_NAME = "MRNoteNotoSansSC";

export interface PdfNativeFontExtensionPackage {
  readonly schemaVersion: typeof MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION;
  /** The canonical program the glyphs are proven against: the bundled CJK
   * fallback, or a served slice (then `programFace` names its face). */
  readonly fontSha256: string;
  readonly programFace?: PdfNativeEmbeddedFontFace;
  readonly resourceName: string;
  readonly font: Readonly<{ objectNumber: number; generation: 0 }>;
  readonly encodedHex: string;
  readonly widthAt1000: number;
  readonly unitsPerEm: number;
  readonly ascent: number;
  readonly descent: number;
  readonly objectByteLength: number;
  readonly objects: readonly PdfNativeFontExtensionObject[];
}

/** One object's bytes on the wire: base64, split into chunks when long.
 *
 * The record lands in the page's content JSON, which admits no string longer
 * than 1,048,576 characters. An installed CJK program is larger than that in
 * base64, so its object travels as consecutive chunks of whole base64 quads;
 * a short object keeps the single string it always had. */
export type PdfNativeFontObjectWireBytes =
  | { readonly bytesBase64: string }
  | { readonly bytesBase64Chunks: readonly string[] };

export const PDF_NATIVE_FONT_OBJECT_BASE64_CHUNK_CHARS = 512 * 1024;

export function pdfNativeFontObjectWireBytes(
  base64: string,
): PdfNativeFontObjectWireBytes {
  if (base64.length <= PDF_NATIVE_FONT_OBJECT_BASE64_CHUNK_CHARS) {
    return Object.freeze({ bytesBase64: base64 });
  }
  const chunks: string[] = [];
  for (
    let offset = 0;
    offset < base64.length;
    offset += PDF_NATIVE_FONT_OBJECT_BASE64_CHUNK_CHARS
  ) {
    chunks.push(
      base64.slice(offset, offset + PDF_NATIVE_FONT_OBJECT_BASE64_CHUNK_CHARS),
    );
  }
  return Object.freeze({ bytesBase64Chunks: Object.freeze(chunks) });
}

/** The whole base64 text of a wire object, or null when it carries neither
 * form, both, or chunks that are not whole quads before the last. */
export function pdfNativeFontObjectWireBase64(
  item: Readonly<Record<string, unknown>>,
): string | null {
  const single = item.bytesBase64;
  const chunks = item.bytesBase64Chunks;
  if (single !== undefined) {
    return chunks === undefined && typeof single === "string" ? single : null;
  }
  if (
    !Array.isArray(chunks) ||
    chunks.length < 2 ||
    chunks.length > 64 ||
    chunks.some(
      (chunk, index) =>
        typeof chunk !== "string" ||
        chunk.length < 1 ||
        chunk.length > PDF_NATIVE_FONT_OBJECT_BASE64_CHUNK_CHARS ||
        (index < chunks.length - 1 &&
          chunk.length !== PDF_NATIVE_FONT_OBJECT_BASE64_CHUNK_CHARS),
    )
  ) {
    return null;
  }
  return chunks.join("");
}

export interface PdfNativeFontExtensionWirePackage extends Omit<
  PdfNativeFontExtensionPackage,
  "objects"
> {
  readonly objects: readonly Readonly<
    Omit<PdfNativeFontExtensionObject, "bytes"> & PdfNativeFontObjectWireBytes
  >[];
}

export interface PdfNativeFontExtensionTextMetrics {
  readonly text: string;
  readonly widthAt1000: number;
  readonly unitsPerEm: number;
  readonly ascent: number;
  readonly descent: number;
  readonly glyphs: readonly Readonly<{
    glyphId: number;
    text: string;
    codePoints: readonly number[];
    advanceAt1000: number;
  }>[];
}

export interface PdfNativeMergedToUnicodeCMap {
  readonly mappingCount: number;
  readonly streamBytes: Uint8Array;
}

function fail(code: PdfNativeFontExtensionPackageErrorCode): never {
  throw new PdfNativeFontExtensionPackageError(code);
}

function readToUnicodeMap(bytes: Uint8Array): Map<number, string> {
  if (bytes.length < 1 || bytes.length > 4 * 1024 * 1024) {
    return fail("package-invalid");
  }
  const source = new TextDecoder("ascii", { fatal: true }).decode(bytes);
  const codeSpaces = [
    ...source.matchAll(
      /\d+\s+begincodespacerange([\s\S]*?)endcodespacerange/gu,
    ),
  ];
  if (
    codeSpaces.length !== 1 ||
    !/^\s*<0000>\s*<FFFF>\s*$/iu.test(codeSpaces[0][1]) ||
    /beginbfrange/u.test(source)
  ) {
    return fail("package-invalid");
  }
  const result = new Map<number, string>();
  const entryPattern = /<([0-9A-Fa-f]{4})>\s*<([0-9A-Fa-f]{4,16})>/gu;
  const blocks = [
    ...source.matchAll(/(\d+)\s+beginbfchar([\s\S]*?)endbfchar/gu),
  ];
  for (const block of blocks) {
    const entries = [...block[2].matchAll(entryPattern)];
    if (
      Number(block[1]) !== entries.length ||
      block[2].replace(entryPattern, "").trim()
    ) {
      return fail("package-invalid");
    }
    for (const entry of entries) {
      const cid = Number.parseInt(entry[1], 16);
      const unicodeHex = entry[2].toUpperCase();
      if (unicodeHex.length % 4 !== 0) return fail("package-invalid");
      try {
        const unicodeBytes = Uint8Array.from(
          unicodeHex.match(/../gu) ?? [],
          (pair) => Number.parseInt(pair, 16),
        );
        if (
          !new TextDecoder("utf-16be", { fatal: true }).decode(unicodeBytes)
        ) {
          return fail("package-invalid");
        }
      } catch {
        return fail("package-invalid");
      }
      const current = result.get(cid);
      if (current !== undefined && current !== unicodeHex) {
        return fail("package-invalid");
      }
      result.set(cid, unicodeHex);
    }
  }
  if (result.size < 1 || result.size > 65_535) {
    return fail("package-invalid");
  }
  return result;
}

function serializeToUnicodeMap(
  mapping: ReadonlyMap<number, string>,
): Uint8Array {
  const entries = [...mapping.entries()].sort(
    (left, right) => left[0] - right[0],
  );
  const blocks: string[] = [];
  for (let offset = 0; offset < entries.length; offset += 100) {
    const block = entries.slice(offset, offset + 100);
    blocks.push(
      `${block.length} beginbfchar\n${block
        .map(
          ([cid, unicode]) =>
            `<${cid.toString(16).toUpperCase().padStart(4, "0")}> <${unicode}>`,
        )
        .join("\n")}\nendbfchar`,
    );
  }
  return new TextEncoder().encode(
    `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${blocks.join("\n")}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`,
  );
}

/**
 * How the face a package installs reads: its codes, two bytes each, as the
 * text its own ToUnicode names for them (what a reader reads the installed
 * glyphs as). Null when the package carries no ToUnicode it can read (an
 * alias names one the document already holds); the decoder answers null
 * for a code the map does not name.
 */
export function pdfNativeFontExtensionPackageDecoder(
  fontPackage: Pick<PdfNativeFontExtensionPackage, "objects">,
): ((codes: Uint8Array) => string | null) | null {
  for (const object of fontPackage.objects) {
    const source = LATIN1_EXACT.decode(object.bytes);
    const open = /\bstream(?:\r\n|\n)/u.exec(source);
    const length = open ? /\/Length\s+(\d{1,9})\b/u.exec(source.slice(0, open.index)) : null;
    if (!open || !length) continue;
    const start = open.index + open[0].length;
    let data = object.bytes.subarray(start, start + Number(length[1]));
    const dictionary = source.slice(0, open.index);
    if (/\/Filter\s*\/FlateDecode\b/u.test(dictionary)) {
      try {
        data = unzlibSync(data);
      } catch {
        continue;
      }
    } else if (/\/Filter\b/u.test(dictionary)) {
      continue;
    }
    if (!LATIN1_EXACT.decode(data).includes("beginbfchar")) continue;
    let map: Map<number, string>;
    try {
      map = readToUnicodeMap(data);
    } catch {
      continue;
    }
    return (codes) => {
      if (codes.length % 2) return null;
      let text = "";
      for (let index = 0; index < codes.length; index += 2) {
        const unicode = map.get((codes[index] << 8) | codes[index + 1]);
        if (unicode === undefined) return null;
        for (let offset = 0; offset < unicode.length; offset += 4) {
          text += String.fromCharCode(Number.parseInt(unicode.slice(offset, offset + 4), 16));
        }
      }
      return text;
    };
  }
  return null;
}







function withoutNotdefInToUnicode(context: PDFContext): void {
  for (const [ref, object] of context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFRawStream)) continue;
    const filter = object.dict.get(PDFName.of("Filter"));
    let data: Uint8Array = object.contents;
    if (filter === PDFName.of("FlateDecode")) {
      try { data = unzlibSync(data); } catch { continue; }
    } else if (filter !== undefined) continue;
    if (!LATIN1_EXACT.decode(data).includes("beginbfchar")) continue;
    let map: Map<number, string>;
    try { map = readToUnicodeMap(data); } catch { continue; }
    if (!map.delete(0) || map.size === 0) continue;
    // As text: pdf-lib takes a byte array of another realm for a string.
    context.assign(ref, context.flateStream(LATIN1_EXACT.decode(serializeToUnicodeMap(map))));
  }
}

export function mergePdfNativeFontExtensionToUnicode(
  streams: readonly Uint8Array[],
): PdfNativeMergedToUnicodeCMap {
  if (!Array.isArray(streams) || streams.length < 1 || streams.length > 16) {
    return fail("invalid-request");
  }
  const merged = new Map<number, string>();
  for (const stream of streams) {
    for (const [cid, unicode] of readToUnicodeMap(stream)) {
      const current = merged.get(cid);
      if (current !== undefined && current !== unicode) {
        return fail("package-invalid");
      }
      merged.set(cid, unicode);
    }
  }
  return Object.freeze({
    mappingCount: merged.size,
    streamBytes: serializeToUnicodeMap(merged),
  });
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(
    parts.reduce((total, part) => total + part.length, 0),
  );
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

export function serializePdfNativeFontExtensionPackage(
  fontPackage: PdfNativeFontExtensionPackage,
): PdfNativeFontExtensionWirePackage {
  return Object.freeze({
    schemaVersion: fontPackage.schemaVersion,
    fontSha256: fontPackage.fontSha256,
    ...(fontPackage.programFace ? { programFace: Object.freeze({ ...fontPackage.programFace }) } : {}),
    resourceName: fontPackage.resourceName,
    font: fontPackage.font,
    encodedHex: fontPackage.encodedHex,
    widthAt1000: fontPackage.widthAt1000,
    unitsPerEm: fontPackage.unitsPerEm,
    ascent: fontPackage.ascent,
    descent: fontPackage.descent,
    objectByteLength: fontPackage.objectByteLength,
    objects: Object.freeze(
      fontPackage.objects.map((object) =>
        Object.freeze({
          objectNumber: object.objectNumber,
          generation: object.generation,
          sha256: object.sha256,
          ...pdfNativeFontObjectWireBytes(bytesToBase64(object.bytes)),
        }),
      ),
    ),
  });
}

async function sha256(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) return fail("package-invalid");
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", owned.buffer);
  return [...new Uint8Array(digest)]
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("");
}

function serializeIndirectObject(ref: PDFRef, object: PDFObject): Uint8Array {
  const body = new Uint8Array(object.sizeInBytes());
  if (object.copyBytesInto(body, 0) !== body.length) {
    return fail("package-invalid");
  }
  return concat([
    new TextEncoder().encode(`${ref.objectNumber} 0 obj\n`),
    body,
    new TextEncoder().encode("\nendobj"),
  ]);
}

function exactText(value: string): string {
  if (
    !value ||
    value.length > MAX_TEXT_CODE_UNITS ||
    /[\u0000-\u001f\u007f]/u.test(value) ||
    [...value].some((character) => {
      const codePoint = character.codePointAt(0)!;
      return codePoint >= 0xd800 && codePoint <= 0xdfff;
    })
  ) {
    return fail("invalid-request");
  }
  return value;
}

/**
 * Layout must stay cmap-exact: the server proves every encoded code against
 * the canonical font's cmap glyph and sums plain advances for `/W`, so
 * localized forms (Noto Sans SC swaps ASCII digits under `locl`), ligatures,
 * kerning and mark positioning are all switched off.
 */
const CMAP_EXACT_LAYOUT_FEATURES: Readonly<Record<string, boolean>> =
  Object.freeze(
    Object.fromEntries(
      [
        "rvrn", "ccmp", "locl", "rlig", "mark", "mkmk", "calt", "clig", "liga",
        "rclt", "curs", "kern", "dist", "dlig", "pwid", "fwid", "hwid", "vert",
        "vrt2", "aalt", "hist", "ruby",
      ].map((tag) => [tag, false]),
    ),
  );

async function loadFontkit(): Promise<unknown> {
  const fontkitModule: unknown = await import("@pdf-lib/fontkit");
  return (fontkitModule as { default?: unknown }).default ?? fontkitModule;
}






















async function narrowedFontEmbedder(
  canonical: CustomFontEmbedder,
  text: string,
  retainedProgram: Uint8Array | null | undefined,
): Promise<CustomFontEmbedder | null> {
  // A document already carrying the whole program keeps it: narrowing it now
  // would rewrite a font other revisions already reuse byte for byte.
  if (retainedProgram && retainedProgram.length >= canonical.fontData.length) {
    return canonical;
  }
  const textGlyphs = canonical.font.layout(text, canonical.fontFeatures).glyphs;
  const candidates: Uint8Array[] = [];
  for (const tier of pdfNativeFontRepertoireTiers()) {
    const glyphIds = new Set<number>();
    const cmap = new Map<number, number>();
    for (const codePoint of tier.codePoints) {
      if (canonical.font.hasGlyphForCodePoint(codePoint)) {
        const glyphId = canonical.font.glyphForCodePoint(codePoint).id;
        glyphIds.add(glyphId);
        cmap.set(codePoint, glyphId);
      }
    }
    // Only the repertoire proves a glyph is there, so a text reaching outside
    // it moves on to the next one.
    if (textGlyphs.some((glyph) => !glyphIds.has(glyph.id))) continue;
    const narrowedBytes = pdfNativeSparseTrueTypeSubset(
      canonical.fontData,
      glyphIds,
      cmap,
    );
    if (!narrowedBytes) continue;
    candidates.push(narrowedBytes);
    // Without an installed program the smallest repertoire is the answer; with
    // one, only the repertoire it was built from lets this edit reuse it.
    if (!retainedProgram || bytesEqual(narrowedBytes, retainedProgram)) break;
  }
  const chosen =
    (retainedProgram &&
      candidates.find((candidate) => bytesEqual(candidate, retainedProgram))) ||
    candidates[0];
  if (!chosen) return null;
  const narrowed = await CustomFontEmbedder.for(
    (await loadFontkit()) as Parameters<typeof CustomFontEmbedder.for>[0],
    chosen,
    PDF_NATIVE_FALLBACK_FONT_NAME,
    { ...CMAP_EXACT_LAYOUT_FEATURES },
  );
  return narrowed.encodeText(text).asString() ===
      canonical.encodeText(text).asString() &&
    Math.abs(
      narrowed.widthOfTextAtSize(text, 1_000) -
        canonical.widthOfTextAtSize(text, 1_000),
    ) < 0.000_000_1
    ? narrowed
    : null;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

async function verifiedFontEmbedder(
  value: Uint8Array | ArrayBuffer,
  subset = false,
): Promise<CustomFontEmbedder> {
  const fontBytes = await verifyPdfUnicodeFallbackFontBytes(value);
  if (!fontBytes) return fail("invalid-font");
  if (subset) return subsetFontEmbedder(fontBytes);
  return CustomFontEmbedder.for(
    (await loadFontkit()) as Parameters<typeof CustomFontEmbedder.for>[0],
    fontBytes,
    PDF_NATIVE_FALLBACK_FONT_NAME,
    { ...CMAP_EXACT_LAYOUT_FEATURES },
  );
}

/** Subset embedding carries only the glyphs the text encodes, under the fixed
 * custom name the server validates (no random subset tag). It renumbers the
 * glyphs it keeps, so only the shows this edit writes can use it; a reusable
 * installed font is narrowed by narrowedFontEmbedder instead. `fontBytes`
 * are the verified canonical program. */
async function subsetFontEmbedder(fontBytes: Uint8Array): Promise<CustomFontEmbedder> {
  const embedder = await CustomFontSubsetEmbedder.for(
    (await loadFontkit()) as Parameters<typeof CustomFontSubsetEmbedder.for>[0],
    fontBytes,
    PDF_NATIVE_FALLBACK_FONT_NAME,
    { ...CMAP_EXACT_LAYOUT_FEATURES },
  );
  alignSubsetGlyphRecords(embedder);
  return embedder;
}

/** Exact logical-order advances used by the bounded fixed-frame reflow layer. */
export async function measurePdfNativeFontExtensionText(input: {
  readonly fontBytes: Uint8Array | ArrayBuffer;
  readonly text: string;
}): Promise<PdfNativeFontExtensionTextMetrics> {
  const text = exactText(input.text);
  const embedder = await verifiedFontEmbedder(input.fontBytes);
  if (
    [...text].some(
      (character) =>
        !embedder.font.hasGlyphForCodePoint(character.codePointAt(0)!),
    )
  ) {
    return fail("unsupported-glyph");
  }
  const run = embedder.font.layout(text, embedder.fontFeatures);
  if (
    run.glyphs.length < 1 ||
    run.glyphs.length > MAX_TEXT_CODE_UNITS ||
    run.positions.length !== run.glyphs.length
  ) {
    return fail("package-invalid");
  }
  const glyphs = run.glyphs.map((glyph) => {
    const glyphText = String.fromCodePoint(...glyph.codePoints);
    const advanceAt1000 = glyph.advanceWidth * embedder.scale;
    if (
      !glyphText ||
      !Number.isSafeInteger(glyph.id) ||
      glyph.id < 0 ||
      !Number.isFinite(advanceAt1000) ||
      advanceAt1000 < 0 ||
      glyph.codePoints.some(
        (codePoint) =>
          !Number.isSafeInteger(codePoint) ||
          codePoint < 0 ||
          codePoint > 0x10ffff,
      )
    ) {
      return fail("package-invalid");
    }
    return Object.freeze({
      glyphId: glyph.id,
      text: glyphText,
      codePoints: Object.freeze([...glyph.codePoints]),
      advanceAt1000,
    });
  });
  if (glyphs.map((glyph) => glyph.text).join("") !== text) {
    return fail("package-invalid");
  }
  const widthAt1000 = glyphs.reduce(
    (total, glyph) => total + glyph.advanceAt1000,
    0,
  );
  const measuredWidth = embedder.widthOfTextAtSize(text, 1_000);
  if (
    !Number.isFinite(widthAt1000) ||
    widthAt1000 <= 0 ||
    Math.abs(widthAt1000 - measuredWidth) > 0.000001
  ) {
    return fail("package-invalid");
  }
  return Object.freeze({
    text,
    widthAt1000,
    unitsPerEm: embedder.font.unitsPerEm,
    ascent: embedder.font.ascent,
    descent: embedder.font.descent,
    glyphs: Object.freeze(glyphs),
  });
}

/** Build a relocatable, source-independent Type0 font object package. */
export async function buildPdfNativeFontExtensionPackage(input: {
  readonly fontBytes: Uint8Array | ArrayBuffer;
  readonly text: string;
  readonly resourceName: string;
  readonly firstObjectNumber: number;
  /** Embed only the encoded glyphs instead of the whole canonical program. */
  readonly subset?: boolean;
  /** A program already installed in the document, whose glyphs this package
   * must keep so an existing reference to it stays reusable. */
  readonly retainedProgram?: Uint8Array | null;
}): Promise<PdfNativeFontExtensionPackage> {
  const text = exactText(input.text);
  if (
    !RESOURCE_NAME_PATTERN.test(input.resourceName) ||
    !Number.isSafeInteger(input.firstObjectNumber) ||
    input.firstObjectNumber < 1 ||
    input.firstObjectNumber > MAX_OBJECT_NUMBER
  ) {
    return fail("invalid-request");
  }
  const verified = await verifiedFontEmbedder(
    input.fontBytes,
    input.subset === true,
  );
  if (
    [...text].some(
      (character) =>
        !verified.font.hasGlyphForCodePoint(character.codePointAt(0)!),
    )
  ) {
    return fail("unsupported-glyph");
  }
  const embedder =
    input.subset === true
      ? verified
      : (await narrowedFontEmbedder(verified, text, input.retainedProgram)) ??
        (await subsetFontEmbedder(verified.fontData));
  return assemblePackage({
    embedder,
    text,
    resourceName: input.resourceName,
    firstObjectNumber: input.firstObjectNumber,
    fontSha256: PDF_UNICODE_FALLBACK_FONT_SHA256,
  });
}

/** The package's objects, written by pdf-lib's Type0 embedder from the
 * relocated first object number on. */
async function assemblePackage(input: {
  readonly embedder: CustomFontEmbedder;
  readonly text: string;
  readonly resourceName: string;
  readonly firstObjectNumber: number;
  readonly fontSha256: string;
  readonly programFace?: PdfNativeEmbeddedFontFace;
}): Promise<PdfNativeFontExtensionPackage> {
  const { embedder, text } = input;
  const encodedHex = embedder.encodeText(text).asString().toUpperCase();
  const widthAt1000 = embedder.widthOfTextAtSize(text, 1_000);
  if (
    !encodedHex ||
    encodedHex.length % 2 !== 0 ||
    !/^[0-9A-F]+$/.test(encodedHex) ||
    !Number.isFinite(widthAt1000) ||
    widthAt1000 <= 0 ||
    widthAt1000 > 10_000_000
  ) {
    return fail("package-invalid");
  }
  // A character the program has no glyph for is laid out as glyph 0, the
  // .notdef box (Noto's is drawn): such a package would write the box into
  
  if (/^(?:[0-9A-F]{4})*0000/u.test(encodedHex)) return fail("unsupported-glyph");

  const context = PDFContext.create();
  context.largestObjectNumber = input.firstObjectNumber;
  const fontRef = PDFRef.of(input.firstObjectNumber, 0);
  await embedder.embedIntoContext(context, fontRef);
  withoutNotdefInToUnicode(context);
  const entries = context
    .enumerateIndirectObjects()
    .sort((left, right) => left[0].objectNumber - right[0].objectNumber);
  if (
    entries.length < 4 ||
    entries.length > MAX_FONT_OBJECTS ||
    !entries.some(([ref]) => ref === fontRef) ||
    entries.some(
      ([ref]) =>
        ref.generationNumber !== 0 ||
        ref.objectNumber < input.firstObjectNumber ||
        ref.objectNumber > input.firstObjectNumber + MAX_FONT_OBJECTS,
    )
  ) {
    return fail("package-invalid");
  }
  const serialized = entries.map(([ref, object]) => ({
    ref,
    bytes: serializeIndirectObject(ref, object),
  }));
  const objectByteLength = serialized.reduce(
    (total, entry) => total + entry.bytes.length,
    0,
  );
  if (objectByteLength > MAX_FONT_PACKAGE_BYTES) {
    return fail("package-invalid");
  }
  const objects = await Promise.all(
    serialized.map(async ({ ref, bytes }) =>
      Object.freeze({
        objectNumber: ref.objectNumber,
        generation: 0 as const,
        sha256: await sha256(bytes),
        bytes,
      }),
    ),
  );
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION,
    fontSha256: input.fontSha256,
    ...(input.programFace ? { programFace: Object.freeze({ ...input.programFace }) } : {}),
    resourceName: input.resourceName,
    font: Object.freeze({
      objectNumber: fontRef.objectNumber,
      generation: 0 as const,
    }),
    encodedHex,
    widthAt1000,
    unitsPerEm: embedder.font.unitsPerEm,
    ascent: embedder.font.ascent,
    descent: embedder.font.descent,
    objectByteLength,
    objects: Object.freeze(objects),
  });
}

const EMBEDDED_FONT_STYLE_NAMES: Readonly<Record<number, string>> = Object.freeze({
  100: "Thin", 200: "ExtraLight", 300: "Light", 350: "SemiLight", 400: "Regular",
  500: "Medium", 600: "SemiBold", 700: "Bold", 800: "ExtraBold", 900: "Black",
});

/** The name an embedded subset of a served face carries, after its subset
 * tag: the substitute's own family and style (`Arimo-Bold`,
 * `NotoSansSC-Regular`). Mirror of the server's office_font_postscript_name. */
export function pdfEmbeddedFontBaseName(face: PdfNativeEmbeddedFontFace): string {
  const family = face.family.replace(/[^A-Za-z0-9]/gu, "") || "Font";
  let style = EMBEDDED_FONT_STYLE_NAMES[face.weight] ?? `W${face.weight}`;
  if (face.italic) style = style === "Regular" ? "Italic" : `${style}Italic`;
  return `${family}-${style}`;
}

/** Six capitals derived from the program's own bytes: two subsets in one
 * document never share a name (PDF 32000-1 9.6.4). Mirror of the server's
 * office_font_subset_tag. */
export async function pdfEmbeddedFontSubsetTag(program: Uint8Array): Promise<string> {
  const digest = await sha256(program);
  let tag = "";
  for (let index = 0; index < 6; index += 1) {
    tag += String.fromCharCode(65 + (Number.parseInt(digest.slice(index * 2, index * 2 + 2), 16) % 26));
  }
  return tag;
}

/**
 * A package that embeds a subset of a served face (sfnt-subset.ts built the
 * program from the slice `programSha256`): the whole small program is the
 * FontFile2, its glyph indices the CIDs, under the face's own name with a
 * subset tag. `text` is every character the package draws, in the order the
 * write draws them; the encoded codes are its shows' codes end to end.
 */
export async function buildPdfNativeEmbeddedFontPackage(input: {
  readonly subset: PdfNativeSfntSubset;
  readonly programSha256: string;
  readonly face: PdfNativeEmbeddedFontFace;
  readonly text: string;
  readonly resourceName: string;
  readonly firstObjectNumber: number;
}): Promise<PdfNativeFontExtensionPackage> {
  const text = exactText(input.text);
  if (
    !RESOURCE_NAME_PATTERN.test(input.resourceName) ||
    !/^[0-9a-f]{64}$/u.test(input.programSha256) ||
    !Number.isSafeInteger(input.firstObjectNumber) ||
    input.firstObjectNumber < 1 ||
    input.firstObjectNumber > MAX_OBJECT_NUMBER
  ) {
    return fail("invalid-request");
  }
  const baseFontName = `${await pdfEmbeddedFontSubsetTag(input.subset.bytes)}+${pdfEmbeddedFontBaseName(input.face)}`;
  let embedder: CustomFontEmbedder;
  try {
    embedder = await CustomFontEmbedder.for(
      (await loadFontkit()) as Parameters<typeof CustomFontEmbedder.for>[0],
      input.subset.bytes,
      baseFontName,
      { ...CMAP_EXACT_LAYOUT_FEATURES },
    );
  } catch {
    return fail("invalid-font");
  }
  const cids = new Map(input.subset.glyphs.map((glyph) => [glyph.text, glyph.cid]));
  const run = embedder.font.layout(text, embedder.fontFeatures);
  const characters = [...text];
  if (
    run.glyphs.length !== characters.length ||
    run.glyphs.some((glyph, index) => glyph.id !== cids.get(characters[index]))
  ) {
    return fail("unsupported-glyph");
  }
  return assemblePackage({
    embedder,
    text,
    resourceName: input.resourceName,
    firstObjectNumber: input.firstObjectNumber,
    fontSha256: input.programSha256,
    programFace: input.face,
  });
}

/** One glyph of a shaped run as the package draws it: its CID in the
 * subset (a shaped glyph's, or a copy's: `buildPdfNativeSfntSubset`
 * `glyphCopies`), and what it reads as. A CID reads as one thing wherever
 * it is drawn: a glyph that stands for other letters elsewhere is drawn
 * from a copy of its own. */
export interface PdfNativeShapedPackageGlyph {
  readonly cid: number;
  readonly text: string;
}

function utf16Hex(text: string): string {
  let hex = "";
  for (let index = 0; index < text.length; index += 1) hex += text.charCodeAt(index).toString(16).toUpperCase().padStart(4, "0");
  return hex;
}

/**
 * A package that draws a shaped run (HarfBuzz's glyphs for Arabic, Hebrew,
 * Thai, Devanagari or Bengali words, `app.services.pdf_text_shaping`) in a
 * subset of its packaged face: a Type 0 font over a CIDFontType2 whose CIDs
 * are the subset's glyph indices (Identity-H, CIDToGIDMap Identity), each
 * glyph's width its advance, and a ToUnicode that names a glyph the letters
 * of its cluster (a joined form or a conjunct has no character of its own;
 * the run's ActualText keeps the words in reading order beside it).
 * `glyphs` is every glyph the write draws, in drawing order: the encoded
 * codes are its shows' codes end to end.
 */
export async function buildPdfNativeShapedFontPackage(input: {
  readonly subset: PdfNativeSfntSubset;
  /** The packaged face's pinned digest (`PDF_COMPLEX_SCRIPT_FONTS`). */
  readonly programSha256: string;
  readonly baseName: string;
  readonly glyphs: readonly PdfNativeShapedPackageGlyph[];
  readonly resourceName: string;
  readonly firstObjectNumber: number;
}): Promise<PdfNativeFontExtensionPackage> {
  const shaped = input.subset.shapedGlyphs;
  if (
    !RESOURCE_NAME_PATTERN.test(input.resourceName) ||
    !/^[0-9a-f]{64}$/u.test(input.programSha256) ||
    !/^[A-Za-z][A-Za-z0-9-]{0,60}$/u.test(input.baseName) ||
    !Number.isSafeInteger(input.firstObjectNumber) ||
    input.firstObjectNumber < 1 ||
    input.firstObjectNumber > MAX_OBJECT_NUMBER ||
    input.glyphs.length === 0 ||
    input.glyphs.length > MAX_TEXT_CODE_UNITS * 4 ||
    !shaped
  ) {
    return fail("invalid-request");
  }
  const unitsPerEm = input.subset.unitsPerEm;
  const at1000 = (advance: number) => Math.round((advance * 1000) / unitsPerEm);
  const width = new Map<number, number>();
  for (const entry of shaped.values()) width.set(entry.cid, at1000(entry.advance));
  for (const entry of input.subset.glyphCopies ?? []) width.set(entry.cid, at1000(entry.advance));
  const cids: number[] = [];
  // What each code reads as: every code drawn reads as something (a code
  // left out of the ToUnicode reads as its own number), and as one thing.
  const letters = new Map<number, string>();
  for (const glyph of input.glyphs) {
    if (!width.has(glyph.cid) || glyph.cid <= 0 || !glyph.text) return fail("unsupported-glyph");
    const held = letters.get(glyph.cid);
    if (held !== undefined && held !== glyph.text) return fail("invalid-request");
    letters.set(glyph.cid, glyph.text);
    cids.push(glyph.cid);
  }
  const encodedHex = cids.map((cid) => cid.toString(16).toUpperCase().padStart(4, "0")).join("");
  const widthAt1000 = cids.reduce((total, cid) => total + (width.get(cid) ?? 0), 0);
  if (!(widthAt1000 > 0) || widthAt1000 > 10_000_000) return fail("package-invalid");
  const baseFont = `${await pdfEmbeddedFontSubsetTag(input.subset.bytes)}+${input.baseName}`;
  const context = PDFContext.create();
  const first = input.firstObjectNumber;
  const [typeZero, cidFont, descriptor, fontFile, toUnicode] = [0, 1, 2, 3, 4].map((offset) => PDFRef.of(first + offset, 0));
  const scale = (value: number) => Math.round((value * 1000) / unitsPerEm);
  const sortedCids = [...width.keys()].sort((left, right) => left - right);
  const toUnicodeBody = [
    "/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
    "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
    "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def",
    "1 begincodespacerange", "<0000> <FFFF>", "endcodespacerange",
    ...(() => {
      const mapped = sortedCids.filter((cid) => letters.has(cid));
      const blocks: string[] = [];
      for (let start = 0; start < mapped.length; start += 100) {
        const block = mapped.slice(start, start + 100);
        blocks.push(`${block.length} beginbfchar`,
          ...block.map((cid) => `<${cid.toString(16).toUpperCase().padStart(4, "0")}> <${utf16Hex(letters.get(cid)!)}>`),
          "endbfchar");
      }
      return blocks;
    })(),
    "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end",
  ].join("\n");
  context.assign(fontFile, context.flateStream(Uint8Array.from(input.subset.bytes), { Length1: input.subset.bytes.length }));
  context.assign(toUnicode, context.flateStream(toUnicodeBody));
  context.assign(descriptor, context.obj({
    Type: "FontDescriptor", FontName: PDFName.of(baseFont), Flags: 4,
    FontBBox: [0, scale(input.subset.descent), 1000, scale(input.subset.ascent)],
    ItalicAngle: 0, Ascent: scale(input.subset.ascent), Descent: scale(input.subset.descent),
    CapHeight: scale(input.subset.ascent), StemV: 80, FontFile2: fontFile,
  }));
  context.assign(cidFont, context.obj({
    Type: "Font", Subtype: "CIDFontType2", BaseFont: PDFName.of(baseFont),
    CIDSystemInfo: context.obj({ Registry: PDFString.of("Adobe"), Ordering: PDFString.of("Identity"), Supplement: 0 }),
    FontDescriptor: descriptor, CIDToGIDMap: "Identity", DW: 0,
    W: sortedCids.flatMap((cid) => [cid, [width.get(cid) ?? 0]]),
  }));
  context.assign(typeZero, context.obj({
    Type: "Font", Subtype: "Type0", BaseFont: PDFName.of(baseFont), Encoding: "Identity-H",
    DescendantFonts: [cidFont], ToUnicode: toUnicode,
  }));
  const serialized = [typeZero, cidFont, descriptor, fontFile, toUnicode].map((ref) => ({
    ref, bytes: serializeIndirectObject(ref, context.lookup(ref)!),
  }));
  const objectByteLength = serialized.reduce((total, entry) => total + entry.bytes.length, 0);
  if (objectByteLength > MAX_FONT_PACKAGE_BYTES) return fail("package-invalid");
  const objects = await Promise.all(serialized.map(async ({ ref, bytes }) => Object.freeze({
    objectNumber: ref.objectNumber, generation: 0 as const, sha256: await sha256(bytes), bytes,
  })));
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION,
    fontSha256: input.programSha256,
    resourceName: input.resourceName,
    font: Object.freeze({ objectNumber: typeZero.objectNumber, generation: 0 as const }),
    encodedHex,
    widthAt1000,
    unitsPerEm,
    ascent: input.subset.ascent,
    descent: input.subset.descent,
    objectByteLength,
    objects: Object.freeze(objects),
  });
}
