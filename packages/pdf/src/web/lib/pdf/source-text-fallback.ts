/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { LATIN1_EXACT } from "./engine/native/latin1";
import { MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE } from "./engine/native/vocabulary";
import type {
  PdfNativeListTextSelectionEvidence,
  PdfNativeListTextSelectionItemEvidence,
  PdfNativeListTextSelectionProof,
  PdfNativeTextWordsFixedReason,
} from "./engine/native/list-text-fidelity";
import { MRAI_NATIVE_REORDERED_TEXT_MESSAGE, pdfTextHasRightToLeftRun } from "./engine/native/reader-order";

export const PDF_FALLBACK_TEXT_SCHEMA_VERSION = "mrai.pdf-source-text-fallback.v1" as const;

export type PdfFallbackTextRefusalCode =
  | "type3"
  | "encrypted"
  | "not-found"
  | "shared"
  | "unreadable"
  | "split-unproven"
  | "glyphs-read-as-blanks"
  | "text-box-source"
  | "letters-read-reordered";

import { appendMraiNativeObjectRevisions } from "./engine/native/content-stream-transaction";
import { MRAI_NATIVE_BOX_MARK_TAG as MRAI_BOX_BLOCK_TAG } from "./engine/native/box-mark";
import { PDF_FALLBACK_NOT_FOUND_MESSAGE, PDF_FALLBACK_TEXT_BOX_SOURCE_MESSAGE, PDF_FALLBACK_TYPE3_MESSAGE } from "./source-text-fallback-copy";
import { pdfLibObjectsPerTick, pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

export { PDF_FALLBACK_TYPE3_MESSAGE };

export class PdfFallbackTextRefusal extends Error {
  constructor(readonly code: PdfFallbackTextRefusalCode, message: string) {
    super(message);
    this.name = "PdfFallbackTextRefusal";
  }
}

export interface PdfFallbackTextPiece {
  readonly object: number;
  readonly generation: number;
  readonly start: number;
  readonly end: number;
  readonly sha256: string;
  /** Present when only some codes of one show are the box's. */
  readonly codeRanges?: readonly (readonly [number, number])[];
  readonly codeWidth?: number;
  /** The show's bytes (base64), for a split piece: Undo rebuilds them. */
  readonly operator?: string;
  /** A show split across two page content streams: this piece is its
   * operands at the end of the first, the continuation its operator at the
   * start of the next. */
  readonly continuation?: Readonly<{ object: number; generation: number; end: number; sha256: string }>;
}

export interface PdfFallbackTextProofData {
  readonly schemaVersion: typeof PDF_FALLBACK_TEXT_SCHEMA_VERSION;
  readonly pageNumber: number;
  readonly pieces: readonly PdfFallbackTextPiece[];
  readonly displayText: string;
  /** Glyphs the reader items hold that no piece covers (text the page draws
   * some other way); the box still opens, the rest is written. */
  readonly unmatchedItems: number;
  readonly fontSize: number;
  /** The first piece's bytes (base64, at most 16 KiB): the binding's
   * `operator`, informational for a span binding. */
  readonly firstPieceHead: string;
  




  readonly renderingMode?: number;
}

// ---------------------------------------------------------------------------
// Content lexer (mirror of source_text_fallback.py lex_content)

type TokenKind = "num" | "name" | "str" | "hex" | "op" | "[" | "]" | "<<" | ">>" | "{" | "}";

interface Token {
  readonly kind: TokenKind;
  readonly start: number;
  readonly end: number;
  readonly text?: string;
  readonly bytes?: Uint8Array;
}

const isWhite = (byte: number) => byte === 0 || byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32;
const isDelimiter = (byte: number) =>
  byte === 0x28 || byte === 0x29 || byte === 0x3c || byte === 0x3e || byte === 0x5b || byte === 0x5d ||
  byte === 0x7b || byte === 0x7d || byte === 0x2f || byte === 0x25;
const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)$/u;
const LATIN1 = LATIN1_EXACT;

function fail(code: PdfFallbackTextRefusalCode, message: string): never {
  throw new PdfFallbackTextRefusal(code, message);
}

function literalBytes(data: Uint8Array, start: number): [Uint8Array, number] {
  const out: number[] = [];
  let depth = 1;
  let index = start + 1;
  const size = data.length;
  while (index < size) {
    const byte = data[index];
    if (byte === 0x5c) {
      index += 1;
      if (index >= size) break;
      const escaped = data[index];
      const simple: Record<number, number> = { 0x6e: 0x0a, 0x72: 0x0d, 0x74: 0x09, 0x62: 0x08, 0x66: 0x0c,
        0x28: 0x28, 0x29: 0x29, 0x5c: 0x5c };
      if (escaped in simple) { out.push(simple[escaped]); index += 1; }
      else if (escaped >= 0x30 && escaped <= 0x37) {
        let value = 0;
        let count = 0;
        while (count < 3 && index < size && data[index] >= 0x30 && data[index] <= 0x37) {
          value = value * 8 + data[index] - 0x30; index += 1; count += 1;
        }
        out.push(value & 0xff);
      } else if (escaped === 0x0d) { index += 1; if (index < size && data[index] === 0x0a) index += 1; }
      else if (escaped === 0x0a) index += 1;
      else { out.push(escaped); index += 1; }
      continue;
    }
    if (byte === 0x28) depth += 1;
    else if (byte === 0x29) {
      depth -= 1;
      if (depth === 0) return [Uint8Array.from(out), index + 1];
    }
    if (byte === 0x0d) {
      out.push(0x0a); index += 1;
      if (index < size && data[index] === 0x0a) index += 1;
      continue;
    }
    out.push(byte);
    index += 1;
  }
  fail("unreadable", "A string in the page content is not terminated.");
}

function hexBytes(data: Uint8Array, start: number, end: number): Uint8Array {
  let digits = "";
  for (let index = start; index < end; index += 1) {
    const byte = data[index];
    if (isWhite(byte)) continue;
    const char = String.fromCharCode(byte);
    if (!/^[0-9A-Fa-f]$/u.test(char)) fail("unreadable", "A hex string in the page content is malformed.");
    digits += char;
  }
  if (digits.length % 2) digits += "0";
  const out = new Uint8Array(digits.length / 2);
  for (let index = 0; index < out.length; index += 1) out[index] = parseInt(digits.slice(index * 2, index * 2 + 2), 16);
  return out;
}

export function lexPdfFallbackContent(data: Uint8Array): readonly Token[] {
  const tokens: Token[] = [];
  let index = 0;
  const size = data.length;
  while (index < size) {
    const byte = data[index];
    if (isWhite(byte)) { index += 1; continue; }
    if (byte === 0x25) {
      while (index < size && data[index] !== 0x0a && data[index] !== 0x0d) index += 1;
      continue;
    }
    const start = index;
    if (byte === 0x28) {
      const [bytes, next] = literalBytes(data, index);
      index = next;
      tokens.push({ kind: "str", start, end: index, bytes });
      continue;
    }
    if (byte === 0x3c) {
      if (data[index + 1] === 0x3c) { tokens.push({ kind: "<<", start, end: index + 2 }); index += 2; continue; }
      let end = index + 1;
      while (end < size && data[end] !== 0x3e) end += 1;
      if (end >= size) fail("unreadable", "A hex string in the page content is not terminated.");
      const bytes = hexBytes(data, index + 1, end);
      index = end + 1;
      tokens.push({ kind: "hex", start, end: index, bytes });
      continue;
    }
    if (byte === 0x3e) {
      if (data[index + 1] === 0x3e) { tokens.push({ kind: ">>", start, end: index + 2 }); index += 2; continue; }
      fail("unreadable", "The page content has a stray delimiter.");
    }
    if (byte === 0x5b || byte === 0x5d || byte === 0x7b || byte === 0x7d) {
      tokens.push({ kind: String.fromCharCode(byte) as TokenKind, start, end: index + 1 });
      index += 1;
      continue;
    }
    if (byte === 0x2f) {
      index += 1;
      while (index < size && !isWhite(data[index]) && !isDelimiter(data[index])) index += 1;
      tokens.push({ kind: "name", start, end: index, text: LATIN1.decode(data.subarray(start + 1, index)) });
      continue;
    }
    if (byte === 0x29) fail("unreadable", "The page content has a stray delimiter.");
    while (index < size && !isWhite(data[index]) && !isDelimiter(data[index])) index += 1;
    const word = LATIN1.decode(data.subarray(start, index));
    if (NUMBER.test(word)) { tokens.push({ kind: "num", start, end: index, text: word }); continue; }
    tokens.push({ kind: "op", start, end: index, text: word });
    if (word === "ID") {
      let position = index + 1;
      for (;;) {
        let found = -1;
        for (let probe = position; probe + 1 < size; probe += 1) {
          if (data[probe] === 0x45 && data[probe + 1] === 0x49) { found = probe; break; }
        }
        if (found < 0) fail("unreadable", "An inline image in the page content is not terminated.");
        const beforeOk = found > 0 && isWhite(data[found - 1]);
        const after = found + 2;
        const afterOk = after >= size || isWhite(data[after]) || isDelimiter(data[after]);
        if (beforeOk && afterOk) {
          tokens.push({ kind: "op", start: found, end: after, text: "EI" });
          index = after;
          break;
        }
        position = found + 1;
      }
    }
  }
  return tokens;
}

// ---------------------------------------------------------------------------
// Hidden form of a piece (mirror of source_text_fallback.py hidden_piece_bytes)

type ShowElement = Readonly<{ kind: "s"; bytes: Uint8Array } | { kind: "n"; text: string }>;

const SHOW_OPERATORS = new Set(["Tj", "TJ", "'", "\""]);
const ascii = (text: string) => Uint8Array.from(text, (character) => character.charCodeAt(0));

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) { out.set(part, offset); offset += part.length; }
  return out;
}

function parseShow(raw: Uint8Array): { operator: string; lead: string; elements: readonly ShowElement[] } {
  const tokens = lexPdfFallbackContent(raw);
  const last = tokens.at(-1);
  if (!tokens.length || tokens[0].start !== 0 || !last || last.end !== raw.length || last.kind !== "op") {
    fail("unreadable", "The show is not one text operator.");
  }
  const operator = last.text!;
  const operands = tokens.slice(0, -1);
  const isString = (token: Token) => token.kind === "str" || token.kind === "hex";
  if ((operator === "Tj" || operator === "'") && operands.length === 1 && isString(operands[0])) {
    return { operator, lead: operator === "'" ? "T*" : "", elements: [{ kind: "s", bytes: operands[0].bytes! }] };
  }
  if (operator === "\"" && operands.length === 3 && operands[0].kind === "num" && operands[1].kind === "num" &&
    isString(operands[2])) {
    return { operator, lead: `${operands[0].text} Tw ${operands[1].text} Tc T*`,
      elements: [{ kind: "s", bytes: operands[2].bytes! }] };
  }
  if (operator === "TJ" && operands.length >= 2 && operands[0].kind === "[" && operands.at(-1)!.kind === "]") {
    const elements: ShowElement[] = [];
    for (const token of operands.slice(1, -1)) {
      if (isString(token)) elements.push({ kind: "s", bytes: token.bytes! });
      else if (token.kind === "num") elements.push({ kind: "n", text: token.text! });
      else fail("unreadable", "The show is not one text operator.");
    }
    return { operator, lead: "", elements };
  }
  fail("unreadable", "The show is not one text operator.");
}

export function pdfFallbackShowCodeCount(raw: Uint8Array, codeWidth: number): number {
  let total = 0;
  for (const element of parseShow(raw).elements) {
    if (element.kind !== "s") continue;
    if (element.bytes.length % codeWidth) fail("split-unproven", "The show's codes do not divide by the font's code width.");
    total += element.bytes.length / codeWidth;
  }
  return total;
}

const hex = (bytes: Uint8Array) => Array.from(bytes, (value) => value.toString(16).padStart(2, "0")).join("").toUpperCase();

export function hiddenPdfFallbackPieceBytes(
  raw: Uint8Array,
  codeRanges: readonly (readonly [number, number])[] | null | undefined,
  codeWidth: number | null | undefined,
  priorMode: number,
): Uint8Array {
  if (!Number.isInteger(priorMode) || priorMode < 0 || priorMode > 7) fail("unreadable", "The rendering mode is invalid.");
  const hidden = priorMode >= 4 ? 7 : 3;
  if (!codeRanges?.length) return concat([ascii(`\n${hidden} Tr\n`), raw, ascii(`\n${priorMode} Tr\n`)]);
  if (codeWidth !== 1 && codeWidth !== 2 && codeWidth !== 3 && codeWidth !== 4) fail("split-unproven", "The code width is invalid.");
  const show = parseShow(raw);
  const total = pdfFallbackShowCodeCount(raw, codeWidth);
  const bounds = [...new Set([0, total, ...codeRanges.flat()])].sort((a, b) => a - b);
  if (bounds[0] !== 0 || bounds.at(-1) !== total) fail("split-unproven", "The code range is outside the show.");
  const segments: [number, number][] = [];
  for (let index = 0; index + 1 < bounds.length; index += 1) {
    if (bounds[index + 1] > bounds[index]) segments.push([bounds[index], bounds[index + 1]]);
  }
  const segmentOf = (code: number) => {
    const found = segments.findIndex(([low, high]) => low <= code && code < high);
    return found < 0 ? segments.length - 1 : found;
  };
  const items: string[][] = segments.map(() => []);
  let code = 0;
  for (const element of show.elements) {
    if (element.kind === "n") { items[segmentOf(code)].push(element.text); continue; }
    let chunk: number[] = [];
    let current = segmentOf(code);
    for (let offset = 0; offset < element.bytes.length; offset += codeWidth) {
      const target = segmentOf(code);
      if (target !== current && chunk.length) {
        items[current].push(`<${hex(Uint8Array.from(chunk))}>`);
        chunk = [];
      }
      current = target;
      chunk.push(...element.bytes.subarray(offset, offset + codeWidth));
      code += 1;
    }
    if (chunk.length) items[current].push(`<${hex(Uint8Array.from(chunk))}>`);
  }
  const parts = segments.map(([low], index) => {
    const body = `[${items[index].join(" ")}] TJ`;
    return codeRanges.some(([start, end]) => start <= low && low < end)
      ? `${hidden} Tr\n${body}\n${priorMode} Tr` : body;
  });
  return ascii(`\n${show.lead ? `${show.lead}\n` : ""}${parts.join("\n")}\n`);
}

// ---------------------------------------------------------------------------
// Page model (pdf-lib), fonts and the content interpreter

type PdfLib = typeof import("pdf-lib");
type Matrix = readonly [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
const multiply = (a: Matrix, b: Matrix): Matrix => [
  a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
  a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
  a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5],
];
const apply = (m: Matrix, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

interface FontInfo {
  readonly kind: "simple" | "type0" | "type3" | "unknown";
  /** Byte length of the code starting at `offset`. */
  codeLength(bytes: Uint8Array, offset: number): number;
  /** Advance of a code in text space per unit font size. */
  width(code: number): number;
  readonly vertical: boolean;
  /** Bytes per code when every code has the same length. */
  readonly codeWidth: number | null;
  /** Codes the font's ToUnicode calls white space whose glyphs draw
   * (readMraiNativeInkedBlankCodes): the reader shows blanks where the page
   * shows letters, so a box drawn from the reader's text cannot stand in for
   * the line they are on. */
  readonly inkedBlankCodes?: ReadonlySet<number>;
  /** The codes its ToUnicode reads as right-to-left letters, and those it
   * reads as several of them (readMraiNativeRightToLeftLetterCodes): the
   * reader turns the letters of such a glyph round. */
  readonly rightToLeftLetterCodes?: Readonly<{ letters: ReadonlySet<number>; several: ReadonlySet<number> }>;
  /** What its ToUnicode reads each code as, for a font of right-to-left
   * letters (pdfInvisibleRightToLeftLineText reads a line in the order its
   * shows hold it). */
  readonly decode?: ReadonlyMap<number, string>;
  /** Its /BaseFont (what a diagnosis names it by). */
  readonly name?: string;
}

interface StreamRef { readonly objectNumber: number; readonly generation: number }

interface ShowRecord {
  readonly ref: StreamRef;
  readonly key: string;
  readonly start: number;
  readonly end: number;
  readonly codeCount: number;
  readonly codeWidth: number | null;
  readonly font: FontInfo;
  readonly mode: number;
  /** Per code: center in page space, text direction, and whether it inks. */
  readonly glyphs: readonly Readonly<{ x: number; y: number; ux: number; uy: number; ink: boolean;
    /** Its character code in the show's font. */
    code: number;
    /** The reader reads it as a blank and it draws (FontInfo.inkedBlankCodes). */
    readAsBlank?: boolean;
    



    size: number }>[];
  /** A show whose operands end the stream `ref` and whose operator starts
   * the next page content stream: `end` is then the first stream's length. */
  readonly continuation?: Readonly<{ ref: StreamRef; key: string; end: number }>;
  /** Where its first glyph starts, in page space (its pen on its baseline). */
  readonly origin: readonly [number, number];
  /** The block of a box written here that draws it (`/MRAIBox BMC` ... `EMC`:
   * a text box made of a line, a paragraph set again), counted along the
   * page's drawing; absent for the page's other text. */
  readonly box?: number;
}

interface StreamRecord {
  readonly ref: StreamRef;
  readonly key: string;
  readonly data: Uint8Array;
  readonly tokens: readonly Token[];
  /** Form streams: how many times the page draws it. */
  invocations: number;
  readonly form: boolean;
}

interface PageReading {
  readonly lib: PdfLib;
  readonly doc: import("pdf-lib").PDFDocument;
  readonly page: StreamRef;
  readonly streams: Map<string, StreamRecord>;
  readonly shows: ShowRecord[];
}

const refKey = (ref: StreamRef) => `${ref.objectNumber} ${ref.generation}`;

async function loadDocument(sourceBytes: Uint8Array) {
  const lib = await import("pdf-lib");
  const doc = await lib.PDFDocument.load(Uint8Array.from(sourceBytes), { parseSpeed: pdfLibParseSpeed(), ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
  if (doc.isEncrypted) fail("encrypted", "The PDF is encrypted; its text cannot be written.");
  return { lib, doc };
}

function numberOf(lib: PdfLib, value: unknown, fallback: number): number {
  return value instanceof lib.PDFNumber ? value.asNumber() : fallback;
}

function parseCodespace(text: string): { bytes: number; low: number; high: number }[] {
  const ranges: { bytes: number; low: number; high: number }[] = [];
  for (const block of text.matchAll(/begincodespacerange([\s\S]*?)endcodespacerange/gu)) {
    for (const pair of block[1].matchAll(/<([0-9A-Fa-f\s]+)>\s*<([0-9A-Fa-f\s]+)>/gu)) {
      const low = pair[1].replace(/\s/gu, ""), high = pair[2].replace(/\s/gu, "");
      ranges.push({ bytes: low.length / 2, low: parseInt(low, 16), high: parseInt(high, 16) });
    }
  }
  return ranges;
}

function codespaceLength(ranges: readonly { bytes: number; low: number; high: number }[], fallback: number) {
  if (!ranges.length) return () => fallback;
  const widths = [...new Set(ranges.map((range) => range.bytes))].sort((a, b) => a - b);
  return (bytes: Uint8Array, offset: number) => {
    for (const width of widths) {
      if (offset + width > bytes.length) break;
      let value = 0;
      for (let index = 0; index < width; index += 1) value = value * 256 + bytes[offset + index];
      if (ranges.some((range) => range.bytes === width && range.low <= value && value <= range.high)) return width;
    }
    return widths[0];
  };
}







function missingFontInfo(lib: PdfLib): FontInfo {
  const standard = standardFontMetrics(lib, "Helvetica", lib.PDFName.of("WinAnsiEncoding"), (entry) => entry);
  return { kind: "unknown", codeLength: () => 1, vertical: false, codeWidth: 1,
    width: (code) => standard?.(code) ?? (code === 32 ? 0.25 : 0.52) };
}

async function readFont(lib: PdfLib, doc: import("pdf-lib").PDFDocument, value: unknown): Promise<FontInfo> {
  const dict = value instanceof lib.PDFRef ? doc.context.lookup(value) : value;
  const unknown = missingFontInfo(lib);
  if (!(dict instanceof lib.PDFDict)) return unknown;
  const get = (target: import("pdf-lib").PDFDict, key: string) => {
    const entry = target.get(lib.PDFName.of(key));
    return entry instanceof lib.PDFRef ? doc.context.lookup(entry) : entry;
  };
  const resolveEntry = (entry: unknown) => (entry instanceof lib.PDFRef ? doc.context.lookup(entry) : entry);
  const subtype = get(dict, "Subtype");
  const subtypeName = subtype instanceof lib.PDFName ? subtype.decodeText() : "";
  if (subtypeName === "Type0") {
    const encoding = get(dict, "Encoding");
    let codeLength: FontInfo["codeLength"] = () => 2;
    let codeWidth: number | null = 2;
    let vertical = false;
    if (encoding instanceof lib.PDFName) {
      const name = encoding.decodeText();
      vertical = /-V$/u.test(name);
      if (!/^Identity-[HV]$/u.test(name) && !/-(?:UCS2|UTF16)(?:-HW)?-[HV]$/u.test(name)) {
        try {
          const { readMraiPredefinedCMapCodespace } = await import("./engine/native/predefined-cmap");
          const ranges = (await readMraiPredefinedCMapCodespace(name)).map((range) =>
            ({ bytes: range.byteLength, low: range.low, high: range.high }));
          const widths = new Set(ranges.map((range) => range.bytes));
          codeLength = codespaceLength(ranges, 2);
          codeWidth = widths.size === 1 ? [...widths][0] : null;
        } catch { codeWidth = null; }
      }
    } else if (encoding instanceof lib.PDFRawStream) {
      try {
        const text = LATIN1.decode(lib.decodePDFRawStream(encoding).decode());
        const ranges = parseCodespace(text);
        const widths = new Set(ranges.map((range) => range.bytes));
        codeLength = codespaceLength(ranges, 2);
        codeWidth = widths.size === 1 ? [...widths][0] : null;
        vertical = /\/WMode\s+1/u.test(text);
      } catch { codeWidth = null; }
    }
    const descendants = get(dict, "DescendantFonts");
    const descendant = descendants instanceof lib.PDFArray ? descendants.lookup(0) : null;
    let defaultWidth = 1;
    const widths = new Map<number, number>();
    if (descendant instanceof lib.PDFDict) {
      defaultWidth = numberOf(lib, get(descendant, "DW"), 1000) / 1000;
      const w = get(descendant, "W");
      if (w instanceof lib.PDFArray) {
        const entries = w.asArray().map((entry) => (entry instanceof lib.PDFRef ? doc.context.lookup(entry) : entry));
        for (let index = 0; index < entries.length;) {
          const first = numberOf(lib, entries[index], NaN);
          const next = entries[index + 1];
          if (next instanceof lib.PDFArray) {
            next.asArray().forEach((entry, offset) => widths.set(first + offset, numberOf(lib, entry, 0) / 1000));
            index += 2;
          } else {
            const last = numberOf(lib, next, NaN), value = numberOf(lib, entries[index + 2], 0) / 1000;
            if (Number.isFinite(first) && Number.isFinite(last) && last - first < 65_536) {
              for (let cid = first; cid <= last; cid += 1) widths.set(cid, value);
            }
            index += 3;
          }
        }
      }
    }
    const identity = encoding instanceof lib.PDFName && /^Identity-[HV]$/u.test(encoding.decodeText());
    let inkedBlankCodes: ReadonlySet<number> | null = null;
    if (identity && descendant instanceof lib.PDFDict) {
      // The code is the glyph id only under an Identity CIDToGIDMap.
      const descendantSubtype = get(descendant, "Subtype");
      const glyphMap = get(descendant, "CIDToGIDMap");
      const descriptor = get(descendant, "FontDescriptor");
      const program = descriptor instanceof lib.PDFDict ? get(descriptor, "FontFile2") : null;
      const toUnicode = get(dict, "ToUnicode");
      if (descendantSubtype instanceof lib.PDFName && descendantSubtype.decodeText() === "CIDFontType2" &&
        (glyphMap === undefined || (glyphMap instanceof lib.PDFName && glyphMap.decodeText() === "Identity")) &&
        program instanceof lib.PDFRawStream && toUnicode instanceof lib.PDFRawStream) {
        try {
          const { readMraiNativeInkedBlankCodes } = await import("./engine/native/list-text-fidelity");
          inkedBlankCodes = readMraiNativeInkedBlankCodes(
            lib.decodePDFRawStream(toUnicode).decode(), lib.decodePDFRawStream(program).decode());
        } catch { inkedBlankCodes = null; }
      }
    }
    const named = get(dict, "BaseFont");
    // (Two-byte codes: what the reader reads each as is asked by the code.)
    let rightToLeftLetterCodes: FontInfo["rightToLeftLetterCodes"] | null = null;
    let decode: ReadonlyMap<number, string> | null = null;
    const table = get(dict, "ToUnicode");
    if (codeWidth === 2 && table instanceof lib.PDFRawStream) {
      try {
        const { readMraiNativeRightToLeftLetterCodes, readMraiNativeToUnicodeDecode } =
          await import("./engine/native/list-text-fidelity");
        const tableBytes = lib.decodePDFRawStream(table).decode();
        rightToLeftLetterCodes = readMraiNativeRightToLeftLetterCodes(tableBytes, 2);
        if (rightToLeftLetterCodes) decode = readMraiNativeToUnicodeDecode(tableBytes, 2);
      } catch { rightToLeftLetterCodes = null; decode = null; }
    }
    return { kind: "type0", codeLength, vertical, codeWidth,
      width: (code) => (identity ? widths.get(code) ?? defaultWidth : defaultWidth),
      ...(inkedBlankCodes?.size ? { inkedBlankCodes } : {}),
      ...(rightToLeftLetterCodes ? { rightToLeftLetterCodes } : {}),
      ...(decode ? { decode } : {}),
      ...(named instanceof lib.PDFName ? { name: named.decodeText() } : {}) };
  }
  const firstChar = numberOf(lib, get(dict, "FirstChar"), 0);
  const widthsValue = get(dict, "Widths");
  const widthList = widthsValue instanceof lib.PDFArray
    ? widthsValue.asArray().map((entry) => numberOf(lib, entry instanceof lib.PDFRef ? doc.context.lookup(entry) : entry, 0))
    : null;
  const descriptor = get(dict, "FontDescriptor");
  const missing = descriptor instanceof lib.PDFDict ? numberOf(lib, get(descriptor, "MissingWidth"), 0) : 0;
  if (subtypeName === "Type3") {
    const matrix = get(dict, "FontMatrix");
    const scale = matrix instanceof lib.PDFArray ? numberOf(lib, matrix.lookup(0), 0.001) : 0.001;
    return { kind: "type3", codeLength: () => 1, vertical: false, codeWidth: 1,
      width: (code) => (widthList?.[code - firstChar] ?? 0) * scale };
  }
  const baseFont = get(dict, "BaseFont");
  const baseName = baseFont instanceof lib.PDFName ? baseFont.decodeText().replace(/^[A-Z]{6}\+/u, "") : "";
  const monospace = /Courier/iu.test(baseName);
  // A standard-14 face named without /Widths is measured by its own AFM
  // metrics (pdf-lib carries them), through its encoding and /Differences.
  const standard = widthList ? null : standardFontMetrics(lib, baseName, get(dict, "Encoding"), resolveEntry);
  let rightToLeftLetterCodes: FontInfo["rightToLeftLetterCodes"] | null = null;
  let decode: ReadonlyMap<number, string> | null = null;
  const table = get(dict, "ToUnicode");
  if (table instanceof lib.PDFRawStream) {
    try {
      const { readMraiNativeRightToLeftLetterCodes, readMraiNativeToUnicodeDecode } =
        await import("./engine/native/list-text-fidelity");
      const tableBytes = lib.decodePDFRawStream(table).decode();
      rightToLeftLetterCodes = readMraiNativeRightToLeftLetterCodes(tableBytes, 1);
      if (rightToLeftLetterCodes) decode = readMraiNativeToUnicodeDecode(tableBytes, 1);
    } catch { rightToLeftLetterCodes = null; decode = null; }
  }
  return { kind: "simple", codeLength: () => 1, vertical: false, codeWidth: 1,
    width: (code) => {
      const width = widthList?.[code - firstChar];
      if (width !== undefined) return width / 1000;
      if (widthList) return missing / 1000;
      const measured = standard?.(code);
      if (measured !== undefined) return measured;
      return monospace ? 0.6 : code === 32 ? 0.25 : 0.52;
    },
    ...(rightToLeftLetterCodes ? { rightToLeftLetterCodes } : {}),
    ...(decode ? { decode } : {}),
    ...(baseFont instanceof lib.PDFName ? { name: baseFont.decodeText() } : {}) };
}

function standardFontMetrics(
  lib: PdfLib,
  baseName: string,
  encoding: unknown,
  resolveEntry: (value: unknown) => unknown,
): ((code: number) => number | undefined) | null {
  const bold = /Bold|Black|Heavy|Semibold/iu.test(baseName);
  const italic = /Italic|Oblique/iu.test(baseName);
  const family = /Courier|Mono/iu.test(baseName) ? "Courier"
    : /Times|Serif|Roman|Georgia|Garamond/iu.test(baseName) ? "Times"
      : /Symbol/iu.test(baseName) ? "Symbol"
        : /Dingbat/iu.test(baseName) ? "ZapfDingbats" : "Helvetica";
  const names: Record<string, string> = {
    Courier: bold ? (italic ? "Courier-BoldOblique" : "Courier-Bold") : italic ? "Courier-Oblique" : "Courier",
    Times: bold ? (italic ? "Times-BoldItalic" : "Times-Bold") : italic ? "Times-Italic" : "Times-Roman",
    Helvetica: bold ? (italic ? "Helvetica-BoldOblique" : "Helvetica-Bold") : italic ? "Helvetica-Oblique" : "Helvetica",
    Symbol: "Symbol",
    ZapfDingbats: "ZapfDingbats",
  };
  try {
    const embedder = lib.StandardFontEmbedder.for(names[family] as unknown as Parameters<typeof lib.StandardFontEmbedder.for>[0]);
    const glyphNames = new Map<number, string>();
    for (const [, [code, name]] of Object.entries((embedder.encoding as unknown as { unicodeMappings: Record<string, [number, string]> }).unicodeMappings)) {
      if (!glyphNames.has(code)) glyphNames.set(code, name);
    }
    const encodingDict = resolveEntry(encoding);
    if (encodingDict instanceof lib.PDFDict) {
      const differences = resolveEntry(encodingDict.get(lib.PDFName.of("Differences")));
      if (differences instanceof lib.PDFArray) {
        let code = 0;
        for (const entry of differences.asArray()) {
          const value = resolveEntry(entry);
          if (value instanceof lib.PDFNumber) code = value.asNumber();
          else if (value instanceof lib.PDFName) { glyphNames.set(code, value.decodeText()); code += 1; }
        }
      }
    }
    return (code) => {
      const name = glyphNames.get(code);
      if (!name) return undefined;
      const width = embedder.font.getWidthOfGlyph(name);
      return typeof width === "number" && Number.isFinite(width) ? width / 1000 : undefined;
    };
  } catch {
    return null;
  }
}

interface TextState {
  ctm: Matrix; font: string | null; size: number; tc: number; tw: number; tz: number; tl: number; ts: number; tr: number;
  fontInfo: FontInfo | null;
}

async function readPage(sourceBytes: Uint8Array, pageNumber: number): Promise<PageReading> {
  const { lib, doc } = await loadDocument(sourceBytes);
  const page = doc.getPages()[pageNumber - 1];
  if (!page) fail("not-found", "The page is not in the document.");
  const pageRef = page.ref;
  const streams = new Map<string, StreamRecord>();
  const shows: ShowRecord[] = [];
  const fontCache = new Map<unknown, Promise<FontInfo>>();
  const resolve = (value: unknown) => (value instanceof lib.PDFRef ? doc.context.lookup(value) : value);
  const decode = (stream: unknown): Uint8Array => {
    if (stream instanceof lib.PDFRawStream) return lib.decodePDFRawStream(stream).decode();
    fail("unreadable", "A content stream cannot be read.");
  };
  const fontFor = async (resources: unknown, name: string): Promise<FontInfo | null> => {
    const dict = resolve(resources);
    if (!(dict instanceof lib.PDFDict)) return null;
    const fonts = resolve(dict.get(lib.PDFName.of("Font")));
    if (!(fonts instanceof lib.PDFDict)) return null;
    const entry = fonts.get(lib.PDFName.of(name));
    if (!entry) return null;
    let pending = fontCache.get(entry);
    if (!pending) { pending = readFont(lib, doc, entry); fontCache.set(entry, pending); }
    return pending;
  };
  const state: TextState = { ctm: IDENTITY, font: null, size: 0, tc: 0, tw: 0, tz: 100, tl: 0, ts: 0, tr: 0, fontInfo: null };
  const missingFont = missingFontInfo(lib);
  const stack: TextState[] = [];
  let tm: Matrix = IDENTITY, tlm: Matrix = IDENTITY;
  // The marked content open where the drawing stands: a block of a box
  // written here by its count, anything else null.
  const marked: (number | null)[] = [];
  let boxBlocks = 0;
  const boxOpen = () => { for (let at = marked.length - 1; at >= 0; at -= 1) if (marked[at] !== null) return marked[at]!; return undefined; };

  // Operands a page content stream ends with belong to the operator the
  // next one starts with: the page's streams are one drawing.
  type Carried = { operands: Token[]; ref: StreamRef; key: string; length: number };
  const walk = async (ref: StreamRef, data: Uint8Array, resources: unknown, depth: number, form: boolean,
    carried: Carried | null = null): Promise<Carried | null> => {
    const key = refKey(ref);
    let record = streams.get(key);
    if (record) { record.invocations += 1; return null; }
    record = { ref, key, data, tokens: lexPdfFallbackContent(data), invocations: 1, form };
    streams.set(key, record);
    const tokens = record.tokens;
    const operands: Token[] = carried ? [...carried.operands] : [];
    let carriedCount = carried?.operands.length ?? 0;
    const nums = () => operands.filter((token) => token.kind === "num").map((token) => Number(token.text));
    for (let index = 0; index < tokens.length; index += 1) {
      const token = tokens[index];
      if (token.kind !== "op") { operands.push(token); continue; }
      const op = token.text!;
      const values = nums();
      switch (op) {
        case "q": stack.push({ ...state }); break;
        case "Q": { const saved = stack.pop(); if (saved) Object.assign(state, saved); break; }
        case "cm": if (values.length === 6) state.ctm = multiply(values as unknown as Matrix, state.ctm); break;
        case "BT": tm = IDENTITY; tlm = IDENTITY; break;
        case "Tf": {
          const name = [...operands].reverse().find((entry) => entry.kind === "name")?.text ?? null;
          state.font = name; state.size = values.at(-1) ?? state.size;
          state.fontInfo = name ? await fontFor(resources, name) : null;
          break;
        }
        case "Tc": if (values.length) state.tc = values.at(-1)!; break;
        case "Tw": if (values.length) state.tw = values.at(-1)!; break;
        case "Tz": if (values.length) state.tz = values.at(-1)!; break;
        case "TL": if (values.length) state.tl = values.at(-1)!; break;
        case "Ts": if (values.length) state.ts = values.at(-1)!; break;
        case "Tr": if (values.length && values.at(-1)! >= 0 && values.at(-1)! <= 7) state.tr = Math.trunc(values.at(-1)!); break;
        case "Td": case "TD":
          if (values.length >= 2) {
            const [tx, ty] = values.slice(-2);
            if (op === "TD") state.tl = -ty;
            tlm = multiply([1, 0, 0, 1, tx, ty], tlm); tm = tlm;
          }
          break;
        case "Tm": if (values.length === 6) { tlm = values as unknown as Matrix; tm = tlm; } break;
        case "T*": tlm = multiply([1, 0, 0, 1, 0, -state.tl], tlm); tm = tlm; break;
        case "BMC": case "BDC":
          marked.push(operands.find((entry) => entry.kind === "name")?.text === MRAI_BOX_BLOCK_TAG ? (boxBlocks += 1) : null);
          break;
        case "EMC": marked.pop(); break;
        case "Tj": case "TJ": case "'": case "\"": {
          if (op === "\"" && values.length >= 2) { state.tw = values[0]; state.tc = values[1]; }
          if (op === "'" || op === "\"") { tlm = multiply([1, 0, 0, 1, 0, -state.tl], tlm); tm = tlm; }
          const first = operands[0] ?? token;
          const font = state.fontInfo ?? missingFont;
          const size = state.size, th = state.tz / 100;
          const glyphs: { x: number; y: number; ux: number; uy: number; ink: boolean; code: number; size: number; readAsBlank?: boolean }[] = [];
          const origin = apply(multiply(tm, state.ctm), 0, state.ts);
          const box = boxOpen();
          const elements = op === "TJ"
            ? operands.filter((entry) => entry.kind === "num" || entry.kind === "str" || entry.kind === "hex")
            : operands.filter((entry) => entry.kind === "str" || entry.kind === "hex").slice(-1);
          for (const element of elements) {
            if (element.kind === "num") {
              const shift = -Number(element.text) / 1000 * size;
              tm = font.vertical ? multiply([1, 0, 0, 1, 0, shift], tm) : multiply([1, 0, 0, 1, shift * th, 0], tm);
              continue;
            }
            const bytes = element.bytes!;
            for (let offset = 0; offset < bytes.length;) {
              const length = Math.max(1, Math.min(4, font.codeLength(bytes, offset)));
              let code = 0;
              for (let at = 0; at < length && offset + at < bytes.length; at += 1) code = code * 256 + bytes[offset + at];
              const w0 = font.width(code);
              const space = length === 1 && code === 32 ? state.tw : 0;
              const m = multiply(tm, state.ctm);
              if (font.vertical) {
                const ty = -(1 * size) + state.tc + space;
                const [x, y] = apply(m, 0, state.ts + ty / 2);
                const len = Math.hypot(m[0], m[1]) || 1;
                glyphs.push({ x, y, ux: m[0] / len, uy: m[1] / len, ink: true, code, size: size * Math.hypot(m[2], m[3]),
                  ...(font.inkedBlankCodes?.has(code) ? { readAsBlank: true } : {}) });
                tm = multiply([1, 0, 0, 1, 0, ty], tm);
              } else {
                const tx = (w0 * size + state.tc + space) * th;
                const [x, y] = apply(m, (w0 * size * th) / 2, state.ts + 0.3 * size);
                const len = Math.hypot(m[0], m[1]) || 1;
                glyphs.push({ x, y, ux: m[0] / len, uy: m[1] / len, code, size: size * Math.hypot(m[2], m[3]),
                  ink: !(font.kind === "simple" && length === 1 && code === 32),
                  ...(font.inkedBlankCodes?.has(code) ? { readAsBlank: true } : {}) });
                tm = multiply([1, 0, 0, 1, tx, 0], tm);
              }
              offset += length;
            }
          }
          if (carriedCount > 0 && carried) {
            // The show starts in the stream before this one.
            shows.push({ ref: carried.ref, key: carried.key, start: first.start, end: carried.length,
              continuation: { ref, key, end: token.end }, codeCount: glyphs.length,
              codeWidth: font.codeWidth, font, mode: state.tr, glyphs, origin, ...(box !== undefined ? { box } : {}) });
          } else {
            shows.push({ ref, key, start: first.start, end: token.end, codeCount: glyphs.length,
              codeWidth: font.codeWidth, font, mode: state.tr, glyphs, origin, ...(box !== undefined ? { box } : {}) });
          }
          break;
        }
        case "Do": {
          const name = [...operands].reverse().find((entry) => entry.kind === "name")?.text;
          const dict = resolve(resources);
          const xobjects = dict instanceof lib.PDFDict ? resolve(dict.get(lib.PDFName.of("XObject"))) : null;
          const entry = xobjects instanceof lib.PDFDict && name ? xobjects.get(lib.PDFName.of(name)) : null;
          const xobject = resolve(entry);
          if (depth < 12 && entry instanceof lib.PDFRef && xobject instanceof lib.PDFRawStream) {
            const subtype = xobject.dict.get(lib.PDFName.of("Subtype"));
            // A form that draws a table, chart or SmartArt made here as the
            // page's content (managed_object_content.py, /MRNoteObject): its
            // words are the object's, never a line of the page to bind.
            if (subtype instanceof lib.PDFName && subtype.decodeText() === "Form" &&
              !xobject.dict.has(lib.PDFName.of("MRNoteObject"))) {
              const saved = { ...state }, savedTm = tm, savedTlm = tlm;
              const matrix = resolve(xobject.dict.get(lib.PDFName.of("Matrix")));
              if (matrix instanceof lib.PDFArray && matrix.size() === 6) {
                state.ctm = multiply(matrix.asArray().map((value) => numberOf(lib, resolve(value), 0)) as unknown as Matrix, state.ctm);
              }
              const own = xobject.dict.get(lib.PDFName.of("Resources"));
              let data: Uint8Array | null = null;
              try { data = decode(xobject); } catch { data = null; }
              if (data) {
                await walk({ objectNumber: entry.objectNumber, generation: entry.generationNumber }, data,
                  own ?? resources, depth + 1, true);
              }
              Object.assign(state, saved); tm = savedTm; tlm = savedTlm;
            }
          }
          break;
        }
        default: break;
      }
      operands.length = 0;
      carriedCount = 0;
    }
    return !form && operands.length ? { operands: [...operands], ref, key, length: data.length } : null;
  };
  // A page's Contents is a stream or an array of streams, and either may be
  // named by reference (ISO 32000 7.7.3.3): `/Contents 41 0 R` with object
  // 41 the array is the page's streams, not a stream that cannot be read.
  // (Read as one, every line of such a page was refused here as unreadable
  // and bound by its one show instead, which the server proves only on a
  // page of one content stream: walk 0515, a court form.)
  const named = page.node.get(lib.PDFName.of("Contents"));
  const array = named instanceof lib.PDFRef ? doc.context.lookup(named) : named;
  const contents = array instanceof lib.PDFArray ? array : named;
  const refs = contents instanceof lib.PDFRef ? [contents]
    : contents instanceof lib.PDFArray ? contents.asArray().filter((entry): entry is import("pdf-lib").PDFRef => entry instanceof lib.PDFRef)
      : [];
  const resources = page.node.Resources();
  let carried: Awaited<ReturnType<typeof walk>> = null;
  for (const ref of refs) {
    const stream = doc.context.lookup(ref);
    carried = await walk({ objectNumber: ref.objectNumber, generation: ref.generationNumber }, decode(stream),
      resources, 0, false, carried);
  }
  return { lib, doc, page: { objectNumber: pageRef.objectNumber, generation: pageRef.generationNumber }, streams, shows };
}

// ---------------------------------------------------------------------------
// Matching the reader's items to glyphs, and pieces

function itemRegion(item: PdfNativeListTextSelectionItemEvidence, slack = 1, alongSlack = 0.2 * slack) {
  const [a, b, c, d, e, f] = item.transform;
  let ux = a, uy = b;
  let length = Math.hypot(ux, uy);
  let vx = c, vy = d;
  let vLength = Math.hypot(vx, vy);
  if (!(length > 0) && vLength > 0) { ux = vy / vLength; uy = -vx / vLength; length = 1; }
  if (!(vLength > 0) && length > 0) { vx = -uy / length; vy = ux / length; vLength = 1; }
  if (!(length > 0) || !(vLength > 0)) return null;
  ux /= length; uy /= length; vx /= vLength; vy /= vLength;
  const height = Math.max(item.height, Math.hypot(c, d), 0.5);
  const det = ux * vy - uy * vx;
  if (Math.abs(det) < 1e-6) return null;
  return (x: number, y: number, gx: number, gy: number): boolean => {
    const px = x - e, py = y - f;
    const s = (px * vy - py * vx) / det;
    const t = (ux * py - uy * px) / det;
    return gx * ux + gy * uy > 0.5 &&
      s >= -alongSlack * height && s <= item.width + alongSlack * height &&
      t >= -0.45 * slack * height && t <= 1.25 * slack * height;
  };
}

const base64 = (bytes: Uint8Array) => {
  let text = "";
  for (let index = 0; index < bytes.length; index += 0x8000) text += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  return btoa(text);
};

async function sha256(bytes: Uint8Array): Promise<string> {
  const { sha256PdfBytes } = await import("./engine/native/content-stream-transaction");
  return sha256PdfBytes(bytes);
}

let readingCache: { bytes: Uint8Array; length: number; page: number; reading: Promise<PageReading> } | null = null;

/** Boxes of one page read the same page: the last reading is kept for the
 * same bytes (the same array, never mutated here) and page. */
function pageReading(sourceBytes: Uint8Array, pageNumber: number): Promise<PageReading> {
  if (!readingCache || readingCache.bytes !== sourceBytes || readingCache.page !== pageNumber ||
    readingCache.length !== sourceBytes.length) {
    readingCache = { bytes: sourceBytes, length: sourceBytes.length, page: pageNumber,
      reading: readPage(sourceBytes, pageNumber) };
    readingCache.reading.catch(() => { readingCache = null; });
  }
  return readingCache.reading;
}

/** The items of a box that draw something. */
const boxItems = (selection: PdfNativeListTextSelectionEvidence) =>
  (selection.items?.length ? selection.items : [selection.item]).filter((item) => /\S/u.test(item.str));






/**
 * Whether a point is where a show of the item starts: on the item's
 * baseline (a tenth of its height, half a point at least), from where the
 * item begins to where it ends along it. The reader reads an item as one
 * run along one baseline, so every show it read into the item starts there.
 */
function itemStart(item: PdfNativeListTextSelectionItemEvidence, anywhereAlong = false) {
  const [a, b, c, d, e, f] = item.transform;
  let ux = a, uy = b;
  let length = Math.hypot(ux, uy);
  let vx = c, vy = d;
  let vLength = Math.hypot(vx, vy);
  if (!(length > 0) && vLength > 0) { ux = vy / vLength; uy = -vx / vLength; length = 1; }
  if (!(vLength > 0) && length > 0) { vx = -uy / length; vy = ux / length; vLength = 1; }
  if (!(length > 0) || !(vLength > 0)) return null;
  ux /= length; uy /= length; vx /= vLength; vy /= vLength;
  const height = Math.max(item.height, Math.hypot(c, d), 0.5);
  const det = ux * vy - uy * vx;
  if (Math.abs(det) < 1e-6) return null;
  const off = Math.max(0.5, 0.1 * height);
  return (x: number, y: number): boolean => {
    const px = x - e, py = y - f;
    const s = (px * vy - py * vx) / det;
    const t = (ux * py - uy * px) / det;
    return Math.abs(t) <= off && (anywhereAlong || (s >= -0.2 * height - 0.5 && s <= item.width + 0.2 * height + 0.5));
  };
}
function sizedLike(items: readonly PdfNativeListTextSelectionItemEvidence[]) {
  const heights = items.map((item) => Math.max(item.height, Math.hypot(item.transform[2], item.transform[3]), 0.5));
  return (show: ShowRecord, glyph: ShowRecord["glyphs"][number], at: number) =>
    show.font.kind === "type3" || !(glyph.size > 0) ||
    (glyph.size <= 1.6 * heights[at] && glyph.size >= 0.5 * heights[at]);
}









function readAsBlankOnLines(reading: PageReading, items: readonly PdfNativeListTextSelectionItemEvidence[],
  taken?: (show: ShowRecord, code: number) => boolean,
  /** How far about each item a glyph is asked for (itemRegion's slack): the
   * room a box made of the line may come to cover, or a line's own reach. */
  slack = 3, alongSlack = 2): boolean {
  const sized = sizedLike(items);
  const lineRegions = items.map((item) => itemRegion(item, slack, alongSlack));
  return reading.shows.some((show) => Boolean(show.font.inkedBlankCodes) && show.mode !== 3 && show.mode !== 7 &&
    show.glyphs.some((glyph, code) => Boolean(glyph.readAsBlank) && (Boolean(taken?.(show, code)) ||
      lineRegions.some((inside, at) => Boolean(inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy)) && sized(show, glyph, at)))));
}

























function hiddenShowsAreATextBoxSource(
  reading: PageReading,
  items: readonly PdfNativeListTextSelectionItemEvidence[],
  shows: readonly ShowRecord[],
): boolean {
  const { PDFArray, PDFDict, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString } = reading.lib;
  const page = reading.doc.context.lookup(PDFRef.of(reading.page.objectNumber, reading.page.generation));
  const annots = page instanceof PDFDict ? page.lookup(PDFName.of("Annots")) : null;
  if (!(annots instanceof PDFArray)) return false;
  const streams = new Set(shows.map((show) => show.key));
  const origins = items.map((item) => [Number(item.transform[4]), Number(item.transform[5])] as const)
    .filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
  for (let index = 0; index < annots.size(); index += 1) {
    const annotation = annots.lookup(index);
    if (!(annotation instanceof PDFDict)) continue;
    const stored = annotation.lookup(PDFName.of("MRNoteSourceTextBinding"));
    if (!(stored instanceof PDFString) && !(stored instanceof PDFHexString)) continue;
    // A record that cannot be read names no stream: its rectangle alone
    // says whether the box stands over the line.
    let named: Set<string> | null = null;
    let boundAt: readonly [number, number] | null = null;
    try {
      const record = JSON.parse(stored.decodeText()) as {
        object?: unknown; generation?: unknown; pieces?: unknown;
        originalProof?: { evidence?: { selection?: { item?: { transform?: unknown } } } };
      };
      named = new Set<string>();
      const name = (object: unknown, generation: unknown) => {
        if (Number.isSafeInteger(object) && Number.isSafeInteger(generation)) named!.add(`${object} ${generation}`);
      };
      name(record.object, record.generation);
      if (Array.isArray(record.pieces)) {
        for (const piece of record.pieces as { object?: unknown; generation?: unknown }[]) name(piece?.object, piece?.generation);
      }
      const transform = record.originalProof?.evidence?.selection?.item?.transform;
      if (Array.isArray(transform) && transform.length === 6 && Number.isFinite(transform[4]) && Number.isFinite(transform[5])) {
        boundAt = [Number(transform[4]), Number(transform[5])];
      }
    } catch { named = null; }
    if (named && ![...named].some((key) => streams.has(key))) continue;
    const rect = annotation.lookup(PDFName.of("Rect"));
    const corners = rect instanceof PDFArray && rect.size() === 4
      ? [0, 1, 2, 3].map((at) => { const value = rect.lookup(at); return value instanceof PDFNumber ? value.asNumber() : Number.NaN; })
      : null;
    const within = corners && corners.every(Number.isFinite) ? {
      left: Math.min(corners[0], corners[2]) - 2, right: Math.max(corners[0], corners[2]) + 2,
      bottom: Math.min(corners[1], corners[3]) - 2, top: Math.max(corners[1], corners[3]) + 2,
    } : null;
    if (origins.some(([x, y]) =>
      (within !== null && x >= within.left && x <= within.right && y >= within.bottom && y <= within.top) ||
      (boundAt !== null && Math.abs(x - boundAt[0]) < 1 && Math.abs(y - boundAt[1]) < 1))) return true;
  }
  return false;
}

export async function pdfTextLinesDrawGlyphsReadAsBlanks(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
  /** The bytes are a copy made for this question: its reading is not kept
   * (the kept one holds its bytes until the next box is read). */
  transient?: boolean;
}): Promise<boolean> {
  const reading = await (input.transient
    ? readPage(input.sourceBytes, input.selection.pageNumber)
    : pageReading(input.sourceBytes, input.selection.pageNumber));
  return readAsBlankOnLines(reading, boxItems(input.selection));
}

/**
 * Why the words of a box are fixed (PdfNativeListTextSelectionProof.wordsFixed)
 * where the page draws a glyph the reader read as a blank on or near its
 * lines -- the reading pdfTextLinesDrawGlyphsReadAsBlanks answers, which no
 * line may be made a text box against -- and null where it draws none.
 * "glyphs-read-as-blanks": such a glyph stands on a line of the box itself
 * (within a line's own reach: the dots over and under its letters).
 * "glyphs-read-as-blanks-near": only on the lines around it, which a box
 * made of it and grown over them would paint over.
 */
export async function pdfTextLinesWordsFixedReason(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
}): Promise<PdfNativeTextWordsFixedReason | null> {
  const reading = await pageReading(input.sourceBytes, input.selection.pageNumber);
  const items = boxItems(input.selection);
  if (!readAsBlankOnLines(reading, items)) return null;
  return readAsBlankOnLines(reading, items, undefined, 1, 0.2) ? "glyphs-read-as-blanks" : "glyphs-read-as-blanks-near";
}
































const REORDERED_LETTERS_SHARE = 0.25;
const REORDERED_LETTERS_FEWEST = 20;

const READ_REORDERED_FONTS = new WeakMap<PageReading, ReadonlySet<FontInfo>>();

/** The fonts of a page whose lines the reader reads as other words than the
 * page shows (REORDERED_LETTERS_SHARE). */
function fontsReadReordered(reading: PageReading): ReadonlySet<FontInfo> {
  const known = READ_REORDERED_FONTS.get(reading);
  if (known) return known;
  const drawn = new Map<FontInfo, { letters: number; several: number }>();
  for (const show of reading.shows) {
    const codes = show.font.rightToLeftLetterCodes;
    if (!codes || show.mode === 3 || show.mode === 7) continue;
    let use = drawn.get(show.font);
    if (!use) { use = { letters: 0, several: 0 }; drawn.set(show.font, use); }
    for (const glyph of show.glyphs) {
      if (codes.letters.has(glyph.code)) use.letters += 1;
      if (codes.several.has(glyph.code)) use.several += 1;
    }
  }
  const fonts = new Set<FontInfo>();
  for (const [font, use] of drawn) {
    if (use.letters >= REORDERED_LETTERS_FEWEST && use.several >= REORDERED_LETTERS_SHARE * use.letters) fonts.add(font);
  }
  READ_REORDERED_FONTS.set(reading, fonts);
  return fonts;
}














export async function pdfRightToLeftLinesReadReordered(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
}): Promise<boolean> {
  // (Its own text, as the reader read it: a colon beside such a line, whose
  // box a letter of the line stands in, is no line of right-to-left text.)
  if (!pdfTextHasRightToLeftRun(input.selection.item.str)) return false;
  const reading = await pageReading(input.sourceBytes, input.selection.pageNumber);
  const items = boxItems(input.selection);
  const regions = items.map((item) => itemRegion(item));
  const sized = sizedLike(items);
  // A recognised text layer drawn in the invisible mode (a scan's OCR)
  // holds its right-to-left words in reading order, as the recognition
  // engine writes them, each drawn left to right; the reader takes every
  // show for the order it is drawn in and reads each such word turned round
  // (pdf-walk-1000 0480: the layer's "القرآن" read "نآرقلا"). Written from
  // the reader's words, the box made of the line held the letters reversed,
  // drawn over the scan and saved so (once the box writer shaped Arabic
  // script, 67cc34cee06).
  {
    let letters = 0, hidden = 0;
    for (const show of reading.shows) {
      const codes = show.font.rightToLeftLetterCodes;
      if (!codes) continue;
      for (const glyph of show.glyphs) {
        if (!codes.letters.has(glyph.code) ||
          !regions.some((inside, at) => Boolean(inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy)) && sized(show, glyph, at))) continue;
        letters += 1;
        if (show.mode === 3 || show.mode === 7) hidden += 1;
      }
    }
    if (hidden > 0 && hidden * 2 > letters) return true;
  }
  const fonts = fontsReadReordered(reading);
  if (!fonts.size) return false;
  let letters = 0, reordered = 0;
  for (const show of reading.shows) {
    const codes = show.font.rightToLeftLetterCodes;
    if (!codes || show.mode === 3 || show.mode === 7) continue;
    for (const glyph of show.glyphs) {
      if (!codes.letters.has(glyph.code) ||
        !regions.some((inside, at) => Boolean(inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy)) && sized(show, glyph, at))) continue;
      letters += 1;
      if (fonts.has(show.font)) reordered += 1;
    }
  }
  return letters > 0 && reordered * 2 > letters;
}











export async function pdfRightToLeftLineWordsFixedReason(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
}): Promise<PdfNativeTextWordsFixedReason | null> {
  if (!pdfTextHasRightToLeftRun(input.selection.item.str)) return null;
  // A box that lists a line's items as they stand, not as the source draws
  // them, holds a word the page draws in pieces of several runs (the
  // reader's joinRunsOfOneJoinedWord: an elongation set as a run of its own
  // inside a word). Its words are read in another order than they are
  // drawn, by the reader's own joining of them: no writer places a typed
  // letter among such pieces -- typed into one of them when each was a box,
  // a letter was drawn alone, a size smaller, outside the word it was typed
  
  // written where its text stands, as every such line's.
  return await pdfTextLinesWordsFixedReason(input) ??
    (selectionListsAWordDrawnInPieces(input.selection) || await pdfRightToLeftLinesReadReordered(input)
      ? "letters-read-reordered" : null);
}
function selectionListsAWordDrawnInPieces(selection: PdfNativeListTextSelectionEvidence): boolean {
  return selection.items?.some((item, index, items) => index > 0 && item.itemIndex < items[index - 1].itemIndex) ?? false;
}

/**
 * A line of right-to-left text a scan's recognised layer draws in the
 * invisible mode alone (GlyphLessFont, `3 Tr`), as the layer holds it: the
 * characters its shows' codes stand for through the font's ToUnicode, in the
 * order the shows hold them -- its reading order, as the recognition engine
 * writes it. The reader reads each such word turned round (pdf-walk-1000
 * 0480: the layer's "القرآن" read "نآرقلا", pdfRightToLeftLinesReadReordered),
 * so this, not the reader's text, is what the line's box holds and writes.
 *
 * Null for anything else: a line any of whose glyphs paint (visible
 * right-to-left text the reader reads in its visual order, and reads right),
 * a code the table does not read, or a reading whose letters are not the
 * reader's letters of the box (another order of the same characters is all
 * this may change). Shows on another baseline are a line break where the
 * reader's text has one, a space where it has none.
 */
export async function pdfInvisibleRightToLeftLineText(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
}): Promise<string | null> {
  const read = input.selection.item.str;
  if (!pdfTextHasRightToLeftRun(read)) return null;
  const reading = await pageReading(input.sourceBytes, input.selection.pageNumber);
  const items = boxItems(input.selection);
  const regions = items.map((item) => itemRegion(item));
  const sized = sizedLike(items);
  const breaks = /\n/u.test(read);
  let text = "";
  let letters = 0;
  let baseline: number | null = null;
  for (const show of reading.shows) {
    let part = "";
    let firstY: number | null = null;
    let size = 0;
    for (const glyph of show.glyphs) {
      if (!regions.some((inside, at) => Boolean(inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy)) && sized(show, glyph, at))) continue;
      // Every glyph of it drawn in the invisible mode, by a font whose table
      // reads every code.
      if (show.mode !== 3 && show.mode !== 7) return null;
      const decoded = show.font.decode?.get(glyph.code);
      if (decoded === undefined) return null;
      if (show.font.rightToLeftLetterCodes?.letters.has(glyph.code)) letters += 1;
      firstY ??= glyph.y;
      size = Math.max(size, glyph.size);
      part += decoded;
    }
    if (firstY === null) continue;
    if (text) {
      const newLine = baseline !== null && Math.abs(firstY - baseline) > Math.max(size, 1) * 0.5;
      if (newLine) text = `${text.trimEnd()}${breaks ? "\n" : " "}`;
      else if (!/\s$/u.test(text) && !/^\s/u.test(part)) text += " ";
    }
    baseline = firstY;
    text += part;
  }
  if (!letters) return null;
  const tidy = text.split("\n").map((line) => line.replace(/[^\S\n]+/gu, " ").trim()).join("\n").trim();
  // The same characters as the reader's, in another order.
  const counted = (value: string) => [...value.replace(/\s+/gu, "")].sort().join("");
  return tidy && counted(tidy) === counted(read) ? tidy : null;
}

/** What the page content draws of a box, as pieces to hide by span. */
export async function locatePdfFallbackText(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
}): Promise<PdfFallbackTextProofData> {
  const selection = input.selection;
  const reading = await pageReading(input.sourceBytes, selection.pageNumber);
  const items = boxItems(selection);
  const regions = items.map((item) => itemRegion(item));
  const sized = sizedLike(items);
  const matchedItems = new Set<number>();
  // Codes of each show inside the box.
  const selected = new Map<ShowRecord, Set<number>>();
  const starts = items.map((item) => itemStart(item));
  const baselines = items.map((item) => itemStart(item, true));
  for (const show of reading.shows) {
    // A show of a block written here (a text box made of a line, a
    // paragraph set again) is that block's box's: it is this box's only
    // where it starts on one of this box's items -- on the item's baseline,
    // as the reader reads every row of a block. Standing in an item's region
    // alone (the region reaches a line's height above the item's baseline,
    // for the scripts and marks set off it) it is another box's row drawn
    // close over this line: the word typed into the line above, wrapped to a
    // second row 9 pt over this one, was named a piece of this line and
    // emptied when this line was made a box (pdf-walk-1000 0337 p.1,
    
    if (show.box !== undefined && !starts.some((onItem) => onItem?.(show.origin[0], show.origin[1]))) continue;
    // So is a show of a text layer drawn in the invisible mode (a scan's
    // recognised words) that starts on no baseline of this box's items:
    // such a layer sets each word on its own line's baseline, with no
    // scripts or marks off it, and its lines stand closer than a line's
    // reach below its baseline. The next line's first word, its frame
    // running half a line into this one's, was named a piece of this line
    // and hidden with it when the line was written as a box -- "RA" gone
    // from the page and the saved file (pdf-walk-1000 ru0516g, 0516 p.1:
    // "...possono. ra" over "RA telefono 35-178."). Anywhere along a
    // baseline of the box is the box's (a line drawn by one show the
    // reader reads as several items).
    if ((show.mode === 3 || show.mode === 7) &&
      !baselines.some((onLine) => onLine?.(show.origin[0], show.origin[1]))) continue;
    show.glyphs.forEach((glyph, code) => {
      // Every item the glyph stands in reads it: a superscript set right
      // after a word also stands in the word's slack, and crediting only the
      
      // author line "Frostig¹, Benjamini¹" could not be deleted).
      let found = false;
      regions.forEach((inside, at) => {
        if (!inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy) || !sized(show, glyph, at)) return;
        found = true;
        if (glyph.ink) matchedItems.add(at);
      });
      if (!found) return;
      let codes = selected.get(show);
      if (!codes) { codes = new Set(); selected.set(show, codes); }
      codes.add(code);
    });
  }
  const inked = [...selected].filter(([show, codes]) => [...codes].some((code) => show.glyphs[code].ink));
  if (inked.some(([show]) => show.font.kind === "type3")) fail("type3", PDF_FALLBACK_TYPE3_MESSAGE);
  if (!inked.length) {
    // A Type3 face measures its glyphs by its own matrix; one near the box
    // that the reader places differently is still the box's face.
    const near = items.map((item) => itemRegion(item, 3));
    if (reading.shows.some((show) => show.font.kind === "type3" &&
      show.glyphs.some((glyph) => near.some((inside) => inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy))))) {
      fail("type3", PDF_FALLBACK_TYPE3_MESSAGE);
    }
    fail("not-found", PDF_FALLBACK_NOT_FOUND_MESSAGE);
  }
  // A line that paints nothing because a text box of the page draws its
  // words is that box's, not a line to open, delete or bind again.
  if (inked.every(([show]) => show.mode === 3 || show.mode === 7) &&
    hiddenShowsAreATextBoxSource(reading, items, inked.map(([show]) => show))) {
    fail("text-box-source", PDF_FALLBACK_TEXT_BOX_SOURCE_MESSAGE);
  }
  // Each show: the runs of its codes the box holds, a gap of spaces joined.
  const nearRegions = items.map((item) => itemRegion(item, 3));
  const plans = new Map<ShowRecord, (readonly [number, number])[] | "whole">();
  for (const [show, codes] of inked) {
    const runs: [number, number][] = [];
    for (let code = 0; code < show.codeCount; code += 1) {
      if (!codes.has(code)) continue;
      const last = runs.at(-1);
      if (last && (last[1] === code || show.glyphs.slice(last[1], code).every((glyph) => !glyph.ink))) last[1] = code + 1;
      else runs.push([code, code + 1]);
    }
    // Spaces at either end of the show go with the box: blanks, and glyphs
    // standing just past the line's items (a CID font's space is not code
    // 32 and reads as ink; a Gurmukhi line's trailing space left its show
    
    // r05-12). Glyphs further away stay another item's.
    const first = runs[0], last = runs.at(-1)!;
    const besideLine = (glyph: ShowRecord["glyphs"][number]) =>
      !glyph.ink || nearRegions.some((inside, at) => inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy) && sized(show, glyph, at));
    if (show.glyphs.slice(0, first[0]).every(besideLine)) first[0] = 0;
    if (show.glyphs.slice(last[1]).every(besideLine)) last[1] = show.codeCount;
    plans.set(show, runs.length === 1 && runs[0][0] === 0 && runs[0][1] === show.codeCount ? "whole" : runs);
  }
  // A glyph the reader read as a blank that draws (FontInfo.inkedBlankCodes):
  // hidden with the box, it is gone, since the box is drawn from the reader's
  // text and holds a space there; left beside it, it stays where it was while
  // the box's text is set, coloured or moved without it. Either way the line
  
  // text, its capital A's, colons and quotes named U+0020).
  if (readAsBlankOnLines(reading, items, (show, code) => {
    const plan = plans.get(show);
    return plan === "whole" || Boolean(plan?.some(([from, to]) => code >= from && code < to));
  })) fail("glyphs-read-as-blanks", MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE);
  const pieces: PdfFallbackTextPiece[] = [];
  const byStream = new Map<string, ShowRecord[]>();
  for (const show of reading.shows) {
    const list = byStream.get(show.key) ?? [];
    list.push(show);
    byStream.set(show.key, list);
  }
  for (const [key, shows] of byStream) {
    const stream = reading.streams.get(key)!;
    if (!shows.some((show) => plans.has(show))) continue;
    if (stream.invocations > 1) fail("shared", "This text is drawn from a form the page draws more than once.");
    shows.sort((left, right) => left.start - right.start);
    // A window may run over another operator only when that operator keeps
    // the invisible mode inside it: no Tr, no Do, and q/Q balanced (mirror
    // of the server's _check_window). Indexed once per stream.
    const barriers: number[] = [];
    const nesting: { at: number; depth: number }[] = [];
    let running = 0;
    for (const token of stream.tokens) {
      if (token.kind !== "op") continue;
      const op = token.text!;
      if (op === "Tr" || op === "Do" || op === "BI" || op === "ID" || op === "EI" || op === "d0" || op === "d1") barriers.push(token.start);
      else if (op === "q" || op === "Q") { running += op === "q" ? 1 : -1; nesting.push({ at: token.start, depth: running }); }
    }
    const firstAtOrAfter = (list: readonly number[], position: number) => {
      let low = 0, high = list.length;
      while (low < high) { const mid = (low + high) >> 1; if (list[mid] < position) low = mid + 1; else high = mid; }
      return low;
    };
    const nestingStarts = nesting.map((entry) => entry.at);
    const depthBefore = (position: number) => {
      const index = firstAtOrAfter(nestingStarts, position) - 1;
      return index >= 0 ? nesting[index].depth : 0;
    };
    const windowSafe = (from: number, to: number) => {
      const barrier = firstAtOrAfter(barriers, from);
      if (barrier < barriers.length && barriers[barrier] < to) return false;
      const base = depthBefore(from);
      for (let index = firstAtOrAfter(nestingStarts, from); index < nesting.length && nesting[index].at < to; index += 1) {
        if (nesting[index].depth < base) return false;
      }
      return depthBefore(to) === base;
    };
    let window: { start: number; end: number } | null = null;
    const flush = async () => {
      if (!window) return;
      const bytes = stream.data.subarray(window.start, window.end);
      pieces.push({ object: stream.ref.objectNumber, generation: stream.ref.generation,
        start: window.start, end: window.end, sha256: await sha256(bytes) });
      window = null;
    };
    for (const show of shows) {
      const plan = plans.get(show);
      if (show.continuation) {
        // A show across two page content streams stands alone: its operands
        // end this stream, its operator starts the next.
        await flush();
        if (!plan) continue;
        if (plan !== "whole") fail("split-unproven", "Part of this text shares one show, drawn across two content streams, with other text.");
        const next = reading.streams.get(show.continuation.key)!;
        pieces.push({ object: stream.ref.objectNumber, generation: stream.ref.generation, start: show.start,
          end: show.end, sha256: await sha256(stream.data.subarray(show.start, show.end)),
          continuation: { object: next.ref.objectNumber, generation: next.ref.generation, end: show.continuation.end,
            sha256: await sha256(next.data.subarray(0, show.continuation.end)) } });
        continue;
      }
      if (plan === "whole") {
        if (window && windowSafe(window.end, show.start)) { window.end = show.end; continue; }
        await flush();
        window = { start: show.start, end: show.end };
        continue;
      }
      if (!plan) {
        // A show that draws nothing may stand inside a window.
        if (show.codeCount === 0 && window && windowSafe(window.end, show.start)) { window.end = show.end; continue; }
        // So may a show of blanks alone between two shows of the box: the
        // window goes on over it when the box's next show comes (checked
        // then, from the window's end to that show), and ends before it
        // where none does. A text layer that sets every word by itself with
        // a blank's show after it (an OCR layer: `(word) Tj 100 Tz ( ) Tj`)
        // was a piece a word -- 259 pieces for a paragraph of thirty lines,
        // a binding too long to be taken (pdf-walk-1000 0435 p.1,
        
        if (window && show.codeCount > 0 && show.glyphs.every((glyph) => !glyph.ink)) continue;
        await flush();
        continue;
      }
      await flush();
      if (!show.codeWidth) fail("split-unproven", "Part of this text shares one show with other text in a font whose codes vary in length.");
      const raw = stream.data.subarray(show.start, show.end);
      pieces.push({ object: stream.ref.objectNumber, generation: stream.ref.generation, start: show.start, end: show.end,
        sha256: await sha256(raw), codeRanges: plan.map(([a, b]) => [a, b] as const), codeWidth: show.codeWidth,
        operator: base64(raw) });
    }
    await flush();
  }
  const fontSize = Math.max(...items.map((item) => Math.hypot(item.transform[2], item.transform[3]) || item.height), 1);
  if (!pieces.length) fail("not-found", PDF_FALLBACK_NOT_FOUND_MESSAGE);
  const head = pieces[0];
  const headStream = reading.streams.get(refKey({ objectNumber: head.object, generation: head.generation }))!;
  const renderingMode = sourceRenderingMode(reading, pieces);
  return Object.freeze({
    ...(renderingMode !== undefined ? { renderingMode } : {}),
    firstPieceHead: base64(headStream.data.subarray(head.start, Math.min(head.end, head.start + 16_384))),
    schemaVersion: PDF_FALLBACK_TEXT_SCHEMA_VERSION,
    pageNumber: selection.pageNumber,
    pieces: Object.freeze(pieces),
    displayText: selectionText(selection),
    unmatchedItems: items.length - matchedItems.size,
    fontSize,
  });
}





function sourceRenderingMode(reading: PageReading, pieces: readonly PdfFallbackTextPiece[]): number | undefined {
  const modes = new Set<number>();
  const shown = new Map(reading.shows.map((show) => [`${show.key}@${show.start}`, show.mode]));
  for (const piece of pieces) {
    const key = refKey({ objectNumber: piece.object, generation: piece.generation });
    const data = reading.streams.get(key)?.data;
    if (!data) return undefined;
    const marker = LATIN1.decode(data.subarray(Math.max(0, piece.start - 6), piece.start));
    const restored = /^\n([0-7]) Tr\n/u.exec(LATIN1.decode(data.subarray(piece.end, piece.end + 6)));
    const wrapped = !piece.codeRanges && !piece.continuation && restored && (marker === "\n3 Tr\n" || marker === "\n7 Tr\n");
    const mode = wrapped && restored ? Number(restored[1]) : shown.get(`${key}@${piece.start}`);
    if (mode === undefined) return undefined;
    modes.add(mode);
  }
  return modes.size === 1 ? [...modes][0] : undefined;
}

function selectionText(selection: PdfNativeListTextSelectionEvidence): string {
  return selection.item.str;
}

/** The source with every piece hidden: what the editor paints under the
 * draft while the box is open (never a saved write). */
export async function applyPdfFallbackTextHide(
  sourceBytes: Uint8Array,
  data: PdfFallbackTextProofData,
): Promise<Uint8Array> {
  const reading = await readPage(sourceBytes, data.pageNumber);
  const { lib, doc } = reading;
  // Each stream's edits: (start, end, bytes), applied back to front.
  const edits = new Map<string, [number, number, Uint8Array][]>();
  const add = (key: string, edit: [number, number, Uint8Array]) => edits.set(key, [...(edits.get(key) ?? []), edit]);
  const modes = new Map(reading.shows.map((show) => [`${show.key}@${show.start}`, show.mode]));
  for (const piece of data.pieces) {
    const key = refKey({ objectNumber: piece.object, generation: piece.generation });
    const stream = reading.streams.get(key);
    if (!stream) fail("not-found", "A piece's stream is not on this page.");
    const raw = stream.data.subarray(piece.start, piece.end);
    if (await sha256(raw) !== piece.sha256) fail("not-found", "The source changed under this text.");
    const mode = modes.get(`${key}@${piece.start}`) ?? 0;
    if (piece.continuation) {
      const next = refKey({ objectNumber: piece.continuation.object, generation: piece.continuation.generation });
      add(key, [piece.start, piece.start, ascii(`\n${mode >= 4 ? 7 : 3} Tr\n`)]);
      add(next, [piece.continuation.end, piece.continuation.end, ascii(`\n${mode} Tr\n`)]);
      continue;
    }
    add(key, [piece.start, piece.end, hiddenPdfFallbackPieceBytes(raw, piece.codeRanges, piece.codeWidth, mode)]);
  }
  // Each stream written again, unfiltered, as an object of a revision
  // appended to the source: the hidden page only paints, and a whole rewrite
  // of a big file did not fit the engine's answer.
  const revisions: { objectNumber: number; generation: number; bytes: Uint8Array }[] = [];
  for (const [key, streamEdits] of edits) {
    const stream = reading.streams.get(key);
    if (!stream) fail("not-found", "A piece's stream is not on this page.");
    const parts: Uint8Array[] = [];
    let cursor = 0;
    for (const [start, end, bytes] of [...streamEdits].sort((left, right) => left[0] - right[0])) {
      parts.push(stream.data.subarray(cursor, start), bytes);
      cursor = end;
    }
    parts.push(stream.data.subarray(cursor));
    const ref = lib.PDFRef.of(stream.ref.objectNumber, stream.ref.generation);
    const previous = doc.context.lookup(ref);
    const dict = previous instanceof lib.PDFRawStream ? previous.dict.clone(doc.context) : doc.context.obj({});
    dict.delete(lib.PDFName.of("Filter"));
    dict.delete(lib.PDFName.of("DecodeParms"));
    const data = concat(parts);
    dict.set(lib.PDFName.of("Length"), lib.PDFNumber.of(data.length));
    doc.context.assign(ref, lib.PDFRawStream.of(dict, data));
    revisions.push({
      objectNumber: stream.ref.objectNumber,
      generation: stream.ref.generation,
      bytes: concat([
        ascii(`${stream.ref.objectNumber} ${stream.ref.generation} obj\n${dict.toString()}\nstream\n`),
        data,
        ascii("\nendstream\nendobj"),
      ]),
    });
  }
  try {
    return appendMraiNativeObjectRevisions(sourceBytes, revisions);
  } catch {
    // A source whose cross-references the native reader does not take is
    // written whole, as before.
    return doc.save({ objectsPerTick: pdfLibObjectsPerTick(), useObjectStreams: false, updateFieldAppearances: false });
  }
}

/** Whether every piece is hidden in these bytes (a reload of a saved box). */
export async function isPdfFallbackTextHidden(
  sourceBytes: Uint8Array,
  data: Pick<PdfFallbackTextProofData, "pageNumber" | "pieces">,
): Promise<boolean> {
  try {
    const reading = await readPage(sourceBytes, data.pageNumber);
    for (const piece of data.pieces) {
      const stream = reading.streams.get(refKey({ objectNumber: piece.object, generation: piece.generation }));
      if (!stream) return false;
      const length = piece.end - piece.start;
      let found = false;
      if (piece.continuation) {
        // Operands at the end of this stream, after the invisible mode; the
        // operator at the start of the next, followed by the mode before.
        const next = reading.streams.get(refKey({ objectNumber: piece.continuation.object,
          generation: piece.continuation.generation }));
        const start = stream.data.length - length;
        const marker = LATIN1.decode(stream.data.subarray(start - 6, start));
        found = Boolean(next) && start >= 6 && (marker === "\n3 Tr\n" || marker === "\n7 Tr\n") &&
          await sha256(stream.data.subarray(start)) === piece.sha256 &&
          await sha256(next!.data.subarray(0, piece.continuation.end)) === piece.continuation.sha256 &&
          /^\n[0-7] Tr\n/u.test(LATIN1.decode(next!.data.subarray(piece.continuation.end, piece.continuation.end + 6)));
      } else if (!piece.codeRanges) {
        for (const token of stream.tokens) {
          if (token.kind !== "num" || (token.text !== "3" && token.text !== "7")) continue;
          const at = token.start - 1;
          const marker = ascii(`\n${token.text} Tr\n`);
          if (at < 0 || !marker.every((byte, index) => stream.data[at + index] === byte)) continue;
          const raw = stream.data.subarray(at + marker.length, at + marker.length + length);
          if (raw.length === length && await sha256(raw) === piece.sha256) { found = true; break; }
        }
      } else if (piece.operator) {
        const raw = Uint8Array.from(atob(piece.operator), (character) => character.charCodeAt(0));
        const text = LATIN1.decode(stream.data);
        for (let mode = 0; mode < 8 && !found; mode += 1) {
          found = text.includes(LATIN1.decode(hiddenPdfFallbackPieceBytes(raw, piece.codeRanges, piece.codeWidth, mode)));
        }
      }
      if (!found) return false;
    }
    return true;
  } catch {
    return false;
  }
}










export async function preparePdfFallbackTextOpen(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
  /** False skips the page painted without the box (a probe needs no surface). */
  blank?: boolean;
}): Promise<{ proof: PdfNativeListTextSelectionProof; blankedSource: Uint8Array | null }> {
  let data: PdfFallbackTextProofData;
  try {
    data = await locatePdfFallbackText(input);
  } catch (error) {
    if (!(error instanceof PdfFallbackTextRefusal) || error.code !== "glyphs-read-as-blanks") throw error;
    return { proof: await pdfWordsFixedTextProof(input, await pdfTextLinesWordsFixedReason(input) ?? undefined),
      blankedSource: null };
  }
  
  // words than the page shows: its box would be a line of other letters.)
  // (And a line with a word the page draws in pieces of several runs:
  // pdfRightToLeftLineWordsFixedReason.)
  // (A line of a scan's invisible layer is opened in its reading order, as
  // the layer holds it: pdfInvisibleRightToLeftLineText.)
  const layerText = await pdfInvisibleRightToLeftLineText(input).catch(() => null);
  if (layerText !== null) {
    return {
      proof: pdfFallbackTextProof({ ...data, displayText: layerText }, input.selection, layerText),
      blankedSource: input.blank === false ? null
        : await applyPdfFallbackTextHide(input.sourceBytes, data).catch(() => null),
    };
  }
  if ((pdfTextHasRightToLeftRun(input.selection.item.str) && selectionListsAWordDrawnInPieces(input.selection)) ||
    await pdfRightToLeftLinesReadReordered(input)) {
    try {
      return { proof: await pdfWordsFixedTextProof(input, "letters-read-reordered"), blankedSource: null };
    } catch (error) {
      // A line drawn in the invisible mode alone (a scan's OCR layer) has no
      // show to set a format on either: it is refused, and says why.
      if (error instanceof PdfFallbackTextRefusal && error.code === "not-found") {
        fail("letters-read-reordered", MRAI_NATIVE_REORDERED_TEXT_MESSAGE);
      }
      throw error;
    }
  }
  return {
    proof: pdfFallbackTextProof(data, input.selection),
    blankedSource: input.blank === false ? null
      : await applyPdfFallbackTextHide(input.sourceBytes, data).catch(() => null),
  };
}

export function pdfFallbackTextProof(
  data: PdfFallbackTextProofData,
  selection: PdfNativeListTextSelectionEvidence,
  /** The box's text where it is not the reader's (pdfInvisibleRightToLeftLineText). */
  text?: string,
): PdfNativeListTextSelectionProof {
  const first = data.pieces[0];
  return Object.freeze({
    ...proofOfNoRun(selection, { object: first.object, generation: first.generation, start: first.start, end: first.end,
      sha256: first.sha256 }, data.fontSize, data.renderingMode, text),
    fallbackText: data,
  });
}

/** A proof that states no run: the box as the reader read it, and where its
 * first show stands in the page's content. */
function proofOfNoRun(
  selection: PdfNativeListTextSelectionEvidence,
  first: Readonly<{ object: number; generation: number; start: number; end: number; sha256: string }>,
  fontSize: number,
  renderingMode?: number,
  boxText?: string,
) {
  const items = selection.items?.length ? selection.items : [selection.item];
  const baselines = items.map((item) => item.transform[5]);
  const lineSpacing = baselines.length > 1 ? Math.abs(baselines[0] - baselines[1]) : fontSize * 1.2;
  const text = boxText ?? selectionText(selection);
  return {
    target: Object.freeze({ objectNumber: first.object, generation: first.generation }),
    evidence: Object.freeze({
      schemaVersion: PDF_FALLBACK_TEXT_SCHEMA_VERSION,
      layoutKind: "paragraph-line",
      originalText: text,
      originalDisplayText: text,
      selection,
      fontSize,
      lineSpacing,
      textOperatorSpan: Object.freeze({ start: first.start, end: first.end, sha256: first.sha256 }),
      ...(renderingMode !== undefined ? { renderingMode } : {}),
    }) as unknown as PdfNativeListTextSelectionProof["evidence"],
    properties: null,
  };
}















export async function pdfWordsFixedTextProof(input: {
  sourceBytes: Uint8Array;
  selection: PdfNativeListTextSelectionEvidence;
}, reason: PdfNativeTextWordsFixedReason = "glyphs-read-as-blanks"): Promise<PdfNativeListTextSelectionProof> {
  // (A paragraph that opens as a block for its formats alone may begin with a
  // picture glyph, and a box bold by a Type3 show drawn again with that show:
  // no word of either is set by this proof, in any face.)
  const anyFace = reason === "too-large-to-set-again" || reason === "weight-drawn-again";
  const { selection } = input;
  const reading = await pageReading(input.sourceBytes, selection.pageNumber);
  const items = boxItems(selection);
  const regions = items.map((item) => itemRegion(item));
  const sized = sizedLike(items);
  // The first show, in the page's order, that inks in the box.
  const first = reading.shows.find((show) => show.mode !== 3 && show.mode !== 7 && show.glyphs.some((glyph) =>
    glyph.ink && regions.some((inside, at) => Boolean(inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy)) && sized(show, glyph, at))));
  if (!first) return fail("not-found", PDF_FALLBACK_NOT_FOUND_MESSAGE);
  if (first.font.kind === "type3" && !anyFace) fail("type3", PDF_FALLBACK_TYPE3_MESSAGE);
  const stream = reading.streams.get(first.key);
  if (!stream) return fail("not-found", PDF_FALLBACK_NOT_FOUND_MESSAGE);
  const fontSize = Math.max(...items.map((item) => Math.hypot(item.transform[2], item.transform[3]) || item.height), 1);
  return Object.freeze({
    ...proofOfNoRun(selection, { object: first.ref.objectNumber, generation: first.ref.generation, start: first.start,
      end: first.end, sha256: await sha256(stream.data.subarray(first.start, first.end)) }, fontSize),
    wordsFixed: reason,
  });
}

/** Diagnosis only: every show the page draws, with its glyph centers. */
export async function readPdfFallbackShowsForDiagnosis(sourceBytes: Uint8Array, pageNumber: number) {
  const reading = await readPage(sourceBytes, pageNumber);
  return reading.shows.map((show) => ({ key: show.key, start: show.start, end: show.end, font: show.font.kind,
    mode: show.mode, glyphs: show.glyphs, fontName: show.font.name ?? null,
    fontInkedBlankCodes: show.font.inkedBlankCodes?.size ?? 0,
    letters: show.font.rightToLeftLetterCodes ? show.glyphs.filter((glyph) => show.font.rightToLeftLetterCodes!.letters.has(glyph.code)).length : 0,
    several: show.font.rightToLeftLetterCodes ? show.glyphs.filter((glyph) => show.font.rightToLeftLetterCodes!.several.has(glyph.code)).length : 0 }));
}
