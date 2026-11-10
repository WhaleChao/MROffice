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
} from "./engine/native/list-text-fidelity";

export const PDF_FALLBACK_TEXT_SCHEMA_VERSION = "mrai.pdf-source-text-fallback.v1" as const;

export type PdfFallbackTextRefusalCode =
  | "type3"
  | "encrypted"
  | "not-found"
  | "shared"
  | "unreadable"
  | "split-unproven"
  | "glyphs-read-as-blanks";

import { appendMraiNativeObjectRevisions } from "./engine/native/content-stream-transaction";
import { PDF_FALLBACK_NOT_FOUND_MESSAGE, PDF_FALLBACK_TYPE3_MESSAGE } from "./source-text-fallback-copy";

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
    /** The reader reads it as a blank and it draws (FontInfo.inkedBlankCodes). */
    readAsBlank?: boolean;
    



    size: number }>[];
  /** A show whose operands end the stream `ref` and whose operator starts
   * the next page content stream: `end` is then the first stream's length. */
  readonly continuation?: Readonly<{ ref: StreamRef; key: string; end: number }>;
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
  const doc = await lib.PDFDocument.load(Uint8Array.from(sourceBytes), { ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
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
    return { kind: "type0", codeLength, vertical, codeWidth,
      width: (code) => (identity ? widths.get(code) ?? defaultWidth : defaultWidth),
      ...(inkedBlankCodes?.size ? { inkedBlankCodes } : {}) };
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
  return { kind: "simple", codeLength: () => 1, vertical: false, codeWidth: 1,
    width: (code) => {
      const width = widthList?.[code - firstChar];
      if (width !== undefined) return width / 1000;
      if (widthList) return missing / 1000;
      const measured = standard?.(code);
      if (measured !== undefined) return measured;
      return monospace ? 0.6 : code === 32 ? 0.25 : 0.52;
    } };
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
        case "Tj": case "TJ": case "'": case "\"": {
          if (op === "\"" && values.length >= 2) { state.tw = values[0]; state.tc = values[1]; }
          if (op === "'" || op === "\"") { tlm = multiply([1, 0, 0, 1, 0, -state.tl], tlm); tm = tlm; }
          const first = operands[0] ?? token;
          const font = state.fontInfo ?? missingFont;
          const size = state.size, th = state.tz / 100;
          const glyphs: { x: number; y: number; ux: number; uy: number; ink: boolean; size: number; readAsBlank?: boolean }[] = [];
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
                glyphs.push({ x, y, ux: m[0] / len, uy: m[1] / len, ink: true, size: size * Math.hypot(m[2], m[3]),
                  ...(font.inkedBlankCodes?.has(code) ? { readAsBlank: true } : {}) });
                tm = multiply([1, 0, 0, 1, 0, ty], tm);
              } else {
                const tx = (w0 * size + state.tc + space) * th;
                const [x, y] = apply(m, (w0 * size * th) / 2, state.ts + 0.3 * size);
                const len = Math.hypot(m[0], m[1]) || 1;
                glyphs.push({ x, y, ux: m[0] / len, uy: m[1] / len, size: size * Math.hypot(m[2], m[3]),
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
              codeWidth: font.codeWidth, font, mode: state.tr, glyphs });
          } else {
            shows.push({ ref, key, start: first.start, end: token.end, codeCount: glyphs.length,
              codeWidth: font.codeWidth, font, mode: state.tr, glyphs });
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
            if (subtype instanceof lib.PDFName && subtype.decodeText() === "Form") {
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
  const contents = page.node.get(lib.PDFName.of("Contents"));
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






function sizedLike(items: readonly PdfNativeListTextSelectionItemEvidence[]) {
  const heights = items.map((item) => Math.max(item.height, Math.hypot(item.transform[2], item.transform[3]), 0.5));
  return (show: ShowRecord, glyph: ShowRecord["glyphs"][number], at: number) =>
    show.font.kind === "type3" || !(glyph.size > 0) ||
    (glyph.size <= 1.6 * heights[at] && glyph.size >= 0.5 * heights[at]);
}









function readAsBlankOnLines(reading: PageReading, items: readonly PdfNativeListTextSelectionItemEvidence[],
  taken?: (show: ShowRecord, code: number) => boolean): boolean {
  const sized = sizedLike(items);
  const lineRegions = items.map((item) => itemRegion(item, 3, 2));
  return reading.shows.some((show) => Boolean(show.font.inkedBlankCodes) && show.mode !== 3 && show.mode !== 7 &&
    show.glyphs.some((glyph, code) => Boolean(glyph.readAsBlank) && (Boolean(taken?.(show, code)) ||
      lineRegions.some((inside, at) => Boolean(inside?.(glyph.x, glyph.y, glyph.ux, glyph.uy)) && sized(show, glyph, at)))));
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
  for (const show of reading.shows) {
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
    return doc.save({ useObjectStreams: false, updateFieldAppearances: false });
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
  const data = await locatePdfFallbackText(input);
  return {
    proof: pdfFallbackTextProof(data, input.selection),
    blankedSource: input.blank === false ? null
      : await applyPdfFallbackTextHide(input.sourceBytes, data).catch(() => null),
  };
}

export function pdfFallbackTextProof(
  data: PdfFallbackTextProofData,
  selection: PdfNativeListTextSelectionEvidence,
): PdfNativeListTextSelectionProof {
  const first = data.pieces[0];
  const items = selection.items?.length ? selection.items : [selection.item];
  const baselines = items.map((item) => item.transform[5]);
  const lineSpacing = baselines.length > 1 ? Math.abs(baselines[0] - baselines[1]) : data.fontSize * 1.2;
  return Object.freeze({
    target: Object.freeze({ objectNumber: first.object, generation: first.generation }),
    evidence: Object.freeze({
      schemaVersion: PDF_FALLBACK_TEXT_SCHEMA_VERSION,
      layoutKind: "paragraph-line",
      originalText: data.displayText,
      originalDisplayText: data.displayText,
      selection,
      fontSize: data.fontSize,
      lineSpacing,
      textOperatorSpan: Object.freeze({ start: first.start, end: first.end, sha256: first.sha256 }),
      ...(data.renderingMode !== undefined ? { renderingMode: data.renderingMode } : {}),
    }) as unknown as PdfNativeListTextSelectionProof["evidence"],
    properties: null,
    fallbackText: data,
  });
}

/** Diagnosis only: every show the page draws, with its glyph centers. */
export async function readPdfFallbackShowsForDiagnosis(sourceBytes: Uint8Array, pageNumber: number) {
  const reading = await readPage(sourceBytes, pageNumber);
  return reading.shows.map((show) => ({ key: show.key, start: show.start, end: show.end, font: show.font.kind,
    mode: show.mode, glyphs: show.glyphs }));
}
