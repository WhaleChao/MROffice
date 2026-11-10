/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { LATIN1_EXACT } from "./latin1";
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFObject,
  PDFRef,
  PDFString,
} from "pdf-lib";

import {
  decodeMraiNativePdfStringToken,
  decodeMraiNativePdfTextString,
  encodeMraiNativePdfTextStringToken,
  inspectMraiNativeObject,
  mraiNativeInlineImageDataEnd,
  readMraiNativeStructElementActualText,
  sha256PdfBytes,
  type PdfNativeStructActualTextRevision,
} from "./content-stream-transaction";

/** The refusal for an edit of text whose structure element's /ActualText
 * cannot be written in step with it: the element is shared or reached in a
 * way no one element can be named, its text is a replacement the edit cannot
 * be placed in, or the text it names cannot be read. */
export const MRAI_NATIVE_STRUCT_ACTUAL_TEXT_MESSAGE =
  "The selected text has a replacement text on its structure element (/ActualText); an edit would leave that text saying something else.";

/** `"<page object>:<generation>:<MCID>"` -> the element's ActualText, and
 * the same under `"p<page index>:<MCID>"` for a caller that knows the page
 * by its place in the document. */
export type PdfStructActualTextCoverage = ReadonlyMap<string, string>;

/** The replacement texts structure elements give to sequences on one page
 * (0-based index), each once. */
export function pdfStructActualTextsOnPage(
  coverage: PdfStructActualTextCoverage,
  pageIndex: number,
): readonly string[] {
  const prefix = `p${pageIndex}:`;
  const texts = new Set<string>();
  for (const [key, text] of coverage) {
    if (key.startsWith(prefix)) texts.add(text);
  }
  return [...texts];
}

export function pdfStructActualTextKey(
  page: Readonly<{ objectNumber: number; generation: number }>,
  mcid: number,
): string {
  return `${page.objectNumber}:${page.generation}:${mcid}`;
}

type ObjectRef = Readonly<{ objectNumber: number; generation: number }>;

/** One structure element that names an /ActualText. */
interface StructActualTextElement {
  /** Null when the element is a direct dictionary inside another object. */
  readonly ref: ObjectRef | null;
  readonly text: string;
  /** The marked content its subtree names, in /K order (its own MCIDs and
   * its child elements', an /ActualText standing for everything inside it):
   * the page each is on (a reference's /Pg, else the nearest element's,
   * inherited per ISO 32000-1 14.7.2; null when none names one), and whether
   * a reference names a stream of its own (/Stm, a Form's marked content).
   * Object references (annotations) name no text and are not listed. */
  readonly mcids: Readonly<{ mcid: number; page: ObjectRef | null; viaStream: boolean }>[];
  /** How many parents reach it. */
  parents: number;
}

/** What the structure tree says of one page's MCID. */
interface StructMcidOwner {
  /** How many elements name it in their own /K. */
  count: number;
  /** The nearest element (the owner or an ancestor) naming an /ActualText. */
  actual: StructActualTextElement | null;
  /** Every element naming an /ActualText over it: the owner and its
   * ancestors, for every element naming it (a malformed tree may name one
   * MCID from several). Each one's text follows its content. */
  readonly actuals: StructActualTextElement[];
}

interface StructActualTextIndex {
  /** Keyed by pdfStructActualTextKey. */
  readonly owners: ReadonlyMap<string, StructMcidOwner>;
  /** Pages (by object key) holding an MCID an /ActualText covers. */
  readonly pagesWithActualText: ReadonlySet<string>;
  /** Page object key -> 0-based page index. */
  readonly pageIndexes: ReadonlyMap<string, number>;
  /** Page object key -> the MCID each named property list of its
   * /Resources /Properties carries (`/Span /MC0 BDC`), null for one naming
   * none. */
  readonly namedProperties: ReadonlyMap<string, ReadonlyMap<string, number | null>>;
}

const MAX_NODES = 200_000;

function refOf(value: PDFRef): ObjectRef {
  return Object.freeze({ objectNumber: value.objectNumber, generation: value.generationNumber });
}

function refKey(value: ObjectRef): string {
  return `${value.objectNumber}:${value.generation}`;
}

function readIndex(doc: PDFDocument): StructActualTextIndex {
  const owners = new Map<string, StructMcidOwner>();
  const pagesWithActualText = new Set<string>();
  const pageIndexes = new Map(doc.getPages().map((page, index) => [refKey(refOf(page.ref)), index]));
  const namedProperties = new Map<string, ReadonlyMap<string, number | null>>();
  const index = Object.freeze({ owners, pagesWithActualText, pageIndexes, namedProperties });
  const lookup = (value: PDFObject | undefined): PDFObject | undefined =>
    value instanceof PDFRef ? doc.context.lookup(value) : value;
  for (const page of doc.getPages()) {
    const properties = lookup(page.node.Resources()?.get(PDFName.of("Properties")));
    if (!(properties instanceof PDFDict)) continue;
    const named = new Map<string, number | null>();
    for (const [name, value] of properties.entries()) {
      const list = lookup(value);
      const mcid = list instanceof PDFDict ? lookup(list.get(PDFName.of("MCID"))) : undefined;
      named.set(name.decodeText(), mcid instanceof PDFNumber ? mcid.asNumber() : null);
    }
    namedProperties.set(refKey(refOf(page.ref)), named);
  }
  const root = lookup(doc.catalog.get(PDFName.of("StructTreeRoot")));
  if (!(root instanceof PDFDict)) return index;
  let visited = 0;
  // Every element reached by reference, so one reached twice is known to be
  // shared (and is walked once).
  const elements = new Map<string, StructActualTextElement | null>();
  const pageOf = (dict: PDFDict): ObjectRef | null => {
    const page = dict.get(PDFName.of("Pg"));
    return page instanceof PDFRef ? refOf(page) : null;
  };
  const claim = (page: ObjectRef | null, mcid: number, chain: readonly StructActualTextElement[], viaStream: boolean) => {
    for (const element of chain) element.mcids.push(Object.freeze({ mcid, page, viaStream }));
    if (!page || !Number.isSafeInteger(mcid) || mcid < 0) return;
    const key = pdfStructActualTextKey(page, mcid);
    const actual = chain.at(-1) ?? null;
    const known = owners.get(key);
    if (known) {
      known.count += 1;
      // The nearest element's text wins, as a reader would read it.
      known.actual ??= actual;
      for (const element of chain) if (!known.actuals.includes(element)) known.actuals.push(element);
    } else {
      owners.set(key, { count: 1, actual, actuals: [...chain] });
    }
    if (actual) pagesWithActualText.add(refKey(page));
  };
  /** One kid of an element's /K (or the root's): a bare MCID, a reference
   * dictionary, an object reference, or an element. */
  const walk = (
    rawNode: PDFObject | undefined,
    page: ObjectRef | null,
    chain: readonly StructActualTextElement[],
  ): void => {
    if (visited++ > MAX_NODES) throw new Error("Structure tree is too large.");
    const node = lookup(rawNode);
    if (node instanceof PDFArray) {
      for (const entry of node.asArray()) walk(entry, page, chain);
      return;
    }
    if (node instanceof PDFNumber) {
      claim(page, node.asNumber(), chain, false);
      return;
    }
    if (!(node instanceof PDFDict)) return;
    const mcid = lookup(node.get(PDFName.of("MCID")));
    if (mcid instanceof PDFNumber) {
      claim(pageOf(node) ?? page, mcid.asNumber(), chain, node.get(PDFName.of("Stm")) !== undefined);
      return;
    }
    const type = lookup(node.get(PDFName.of("Type")));
    if (type instanceof PDFName && type.asString() === "/OBJR") return;
    if (rawNode instanceof PDFRef) {
      const key = refKey(refOf(rawNode));
      if (elements.has(key)) {
        const known = elements.get(key);
        if (known) known.parents += 1;
        return;
      }
    }
    const elementPage = pageOf(node) ?? page;
    const actualValue = lookup(node.get(PDFName.of("ActualText")));
    const kids = lookup(node.get(PDFName.of("K")));
    let self: StructActualTextElement | null = null;
    if (actualValue instanceof PDFString || actualValue instanceof PDFHexString) {
      self = {
        ref: rawNode instanceof PDFRef ? refOf(rawNode) : null,
        text: actualValue.decodeText(),
        mcids: [],
        parents: 1,
      };
    }
    if (rawNode instanceof PDFRef) elements.set(refKey(refOf(rawNode)), self);
    walk(kids, elementPage, self ? [...chain, self] : chain);
  };
  walk(root.get(PDFName.of("K")), null, []);
  return index;
}

const INDEX_LIMIT = 4;
const indexes = new Map<string, Promise<StructActualTextIndex>>();

async function readPdfStructActualTextIndex(sourceBytes: Uint8Array): Promise<StructActualTextIndex> {
  const key = await sha256PdfBytes(sourceBytes);
  const cached = indexes.get(key);
  if (cached) {
    indexes.delete(key);
    indexes.set(key, cached);
    return cached;
  }
  const loading = (async () => {
    // A structure tree this reader cannot walk names no element it can
    // write: the edit is written and the tree left as it stands, as WPS
    
    try {
      const doc = await PDFDocument.load(Uint8Array.from(sourceBytes), { updateMetadata: false });
      return readIndex(doc);
    } catch {
      return Object.freeze({ owners: new Map(), pagesWithActualText: new Set<string>(), pageIndexes: new Map(),
        namedProperties: new Map() }) satisfies StructActualTextIndex;
    }
  })();
  indexes.set(key, loading);
  if (indexes.size > INDEX_LIMIT) indexes.delete(indexes.keys().next().value!);
  try {
    return await loading;
  } catch (error) {
    indexes.delete(key);
    throw error;
  }
}

/**
 * Every (page, MCID) a structure element with an /ActualText covers, for a
 * document; read once per source and kept for the last few sources.
 */
export async function readPdfStructActualTextCoverage(
  sourceBytes: Uint8Array,
): Promise<PdfStructActualTextCoverage> {
  const index = await readPdfStructActualTextIndex(sourceBytes);
  const coverage = new Map<string, string>();
  for (const [key, owner] of index.owners) {
    if (!owner.actual) continue;
    coverage.set(key, owner.actual.text);
    const [objectNumber, generation, mcid] = key.split(":");
    const pageIndex = index.pageIndexes.get(`${objectNumber}:${generation}`);
    if (pageIndex !== undefined) coverage.set(`p${pageIndex}:${mcid}`, owner.actual.text);
  }
  return coverage;
}

// -- what marked content reads as ---------------------------------------------

/** One text show of a content stream and the marked content it stands in. */
interface MarkedShow {
  /** The MCID of the innermost sequence naming one; null outside any. */
  readonly mcid: number | null;
  /** It stands in a sequence whose properties are a named resource, so
   * which MCID it belongs to is not read here. */
  readonly unattributed: boolean;
  /** It stands in a sequence with an inline /ActualText. */
  readonly inlineActualText: boolean;
  /** That sequence's /ActualText (the outermost one), which a reader reads
   * in place of the sequence's shows; null when it cannot be read. */
  readonly inlineText: string | null;
  /** The first show of that sequence: the one the text is read at. */
  readonly inlineFirst: boolean;
  readonly font: string | null;
  readonly codes: Uint8Array;
}

interface MarkedContentReading {
  readonly shows: readonly MarkedShow[];
  /** Every MCID a sequence of the stream names, with or without shows. */
  readonly sequences: ReadonlySet<number>;
}

type Operand =
  | Readonly<{ kind: "string"; bytes: Uint8Array }>
  | Readonly<{ kind: "name"; name: string }>
  | Readonly<{ kind: "dict"; source: string }>
  | Readonly<{ kind: "array"; items: readonly Operand[] }>
  | Readonly<{ kind: "other" }>;

const PDF_WHITESPACE = new Set(["\x00", "\t", "\n", "\f", "\r", " "]);
const PDF_DELIMITERS = new Set(["(", ")", "<", ">", "[", "]", "{", "}", "/", "%"]);
const MCID_PROPERTY = /\/MCID\s+(\d{1,9})(?=$|[\s[\]()<>/{}%])/u;
const ACTUAL_TEXT_PROPERTY = /\/ActualText\s*/u;

/** The text an inline `/ActualText` of a marked-content property list names:
 * null when the list names none, "unreadable" when it names one this cannot
 * read. */
function inlineActualText(properties: string): string | null | "unreadable" {
  const key = ACTUAL_TEXT_PROPERTY.exec(properties);
  if (!key) return null;
  const start = key.index + key[0].length;
  const end = properties[start] === "(" ? literalEnd(properties, start)
    : properties[start] === "<" ? properties.indexOf(">", start) + 1 : -1;
  if (end <= start) return "unreadable";
  const bytes = decodeMraiNativePdfStringToken(Uint8Array.from(properties.slice(start, end), (character) => character.charCodeAt(0)));
  const text = bytes ? decodeMraiNativePdfTextString(bytes, false) : null;
  return text ?? "unreadable";
}

function decodePdfName(source: string): string {
  return source.replace(/#([0-9A-Fa-f]{2})/g, (_match, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
}

/** The end of the literal string at `start`, or -1. */
function literalEnd(source: string, start: number): number {
  let depth = 0;
  for (let cursor = start; cursor < source.length; cursor += 1) {
    const character = source[cursor];
    if (character === "\\") { cursor += 1; continue; }
    if (character === "(") depth += 1;
    else if (character === ")") {
      depth -= 1;
      if (depth === 0) return cursor + 1;
    }
  }
  return -1;
}

/** The end of the dictionary at `start` (its `>>`), or -1. */
function dictionaryEnd(source: string, start: number): number {
  let depth = 0;
  for (let cursor = start; cursor < source.length; cursor += 1) {
    const character = source[cursor];
    if (character === "(") {
      const end = literalEnd(source, cursor);
      if (end < 0) return -1;
      cursor = end - 1;
      continue;
    }
    if (character === "%") {
      while (cursor < source.length && source[cursor] !== "\n" && source[cursor] !== "\r") cursor += 1;
      continue;
    }
    if (source.startsWith("<<", cursor)) { depth += 1; cursor += 1; continue; }
    // A hexadecimal string ends at its own `>`: in `<</ActualText<FEFF2003>>>`
    // (PLOS, PDFlib) that `>` is not half of the dictionary's `>>`, and the
    // page read as unreadable, so every edit of a PLOS paper went to the
    
    if (character === "<") {
      const end = source.indexOf(">", cursor);
      if (end < 0) return -1;
      cursor = end;
      continue;
    }
    if (source.startsWith(">>", cursor)) {
      depth -= 1;
      cursor += 1;
      if (depth === 0) return cursor + 1;
      continue;
    }
  }
  return -1;
}

/**
 * The text shows of a decoded content stream, each with the marked content
 * it stands in and the font it is drawn in: what a reader of the structure
 * tree reads an MCID's content by. Null when the stream cannot be read.
 */
function readMarkedContent(
  stream: Uint8Array,
  named: ReadonlyMap<string, number | null> = new Map(),
): MarkedContentReading | null {
  const source = LATIN1_EXACT.decode(stream);
  const shows: MarkedShow[] = [];
  const sequences = new Set<number>();
  const marks: { mcid: number | null; unattributed: boolean; inline: string | null | "unreadable"; emitted: boolean }[] = [];
  const fonts: (string | null)[] = [];
  let font: string | null = null;
  let operands: Operand[] = [];
  const arrays: Operand[][] = [];
  const push = (operand: Operand) => (arrays.length ? arrays[arrays.length - 1] : operands).push(operand);
  const show = (codes: Uint8Array) => {
    let mcid: number | null = null;
    let unattributed = false;
    for (let index = marks.length - 1; index >= 0; index -= 1) {
      if (marks[index].unattributed) { unattributed = true; break; }
      if (marks[index].mcid !== null) { mcid = marks[index].mcid; break; }
    }
    // The outermost inline /ActualText stands for everything inside it.
    const replaced = marks.find((mark) => mark.inline !== null);
    const inlineFirst = Boolean(replaced && !replaced.emitted);
    if (replaced) replaced.emitted = true;
    shows.push(Object.freeze({
      mcid, unattributed, inlineActualText: Boolean(replaced),
      inlineText: replaced && replaced.inline !== "unreadable" ? replaced.inline : null, inlineFirst, font, codes,
    }));
  };
  const stringBytes = (operand: Operand | undefined): Uint8Array | null =>
    operand?.kind === "string" ? operand.bytes : null;
  let cursor = 0;
  while (cursor < source.length) {
    const character = source[cursor];
    if (PDF_WHITESPACE.has(character)) { cursor += 1; continue; }
    if (character === "%") {
      while (cursor < source.length && source[cursor] !== "\n" && source[cursor] !== "\r") cursor += 1;
      continue;
    }
    if (character === "(") {
      const end = literalEnd(source, cursor);
      if (end < 0) return null;
      const bytes = decodeMraiNativePdfStringToken(stream.subarray(cursor, end));
      if (!bytes) return null;
      push({ kind: "string", bytes });
      cursor = end;
      continue;
    }
    if (source.startsWith("<<", cursor)) {
      const end = dictionaryEnd(source, cursor);
      if (end < 0) return null;
      push({ kind: "dict", source: source.slice(cursor, end) });
      cursor = end;
      continue;
    }
    if (character === "<") {
      const end = source.indexOf(">", cursor);
      if (end < 0) return null;
      const bytes = decodeMraiNativePdfStringToken(stream.subarray(cursor, end + 1));
      if (!bytes) return null;
      push({ kind: "string", bytes });
      cursor = end + 1;
      continue;
    }
    if (character === "[") { arrays.push([]); cursor += 1; continue; }
    if (character === "]") {
      const items = arrays.pop();
      if (!items) return null;
      push({ kind: "array", items });
      cursor += 1;
      continue;
    }
    if (character === "/") {
      let end = cursor + 1;
      while (end < source.length && !PDF_WHITESPACE.has(source[end]) && !PDF_DELIMITERS.has(source[end])) end += 1;
      push({ kind: "name", name: decodePdfName(source.slice(cursor + 1, end)) });
      cursor = end;
      continue;
    }
    if (PDF_DELIMITERS.has(character)) return null;
    let end = cursor;
    while (end < source.length && !PDF_WHITESPACE.has(source[end]) && !PDF_DELIMITERS.has(source[end])) end += 1;
    const word = source.slice(cursor, end);
    cursor = end;
    if (/^[+\-.\d]/u.test(word) || word === "true" || word === "false" || word === "null") {
      push({ kind: "other" });
      continue;
    }
    if (arrays.length) return null;
    if (word === "BI") {
      const id = /\sID[\x00\t\n\f\r ]/u.exec(source.slice(cursor));
      if (!id) return null;
      const dataStart = cursor + id.index + id[0].length;
      const dataEnd = mraiNativeInlineImageDataEnd(source, dataStart, source.slice(cursor, cursor + id.index));
      if (dataEnd === null) return null;
      cursor = dataEnd;
      operands = [];
      continue;
    }
    const last = operands[operands.length - 1];
    switch (word) {
      case "q": fonts.push(font); break;
      case "Q": if (fonts.length) font = fonts.pop()!; break;
      case "Tf": {
        const name = operands[operands.length - 2];
        font = name?.kind === "name" ? name.name : null;
        break;
      }
      case "BDC": {
        if (last?.kind === "dict") {
          const mcid = MCID_PROPERTY.exec(last.source);
          const value = mcid ? Number(mcid[1]) : null;
          if (value !== null) sequences.add(value);
          marks.push({ mcid: value, unattributed: false, inline: inlineActualText(last.source), emitted: false });
        } else if (last?.kind === "name" && named.has(last.name)) {
          // A property list the page names in /Resources /Properties: its
          // MCID is the sequence's, read from there.
          const value = named.get(last.name) ?? null;
          if (value !== null) sequences.add(value);
          marks.push({ mcid: value, unattributed: false, inline: null, emitted: false });
        } else {
          marks.push({ mcid: null, unattributed: true, inline: null, emitted: false });
        }
        break;
      }
      case "BMC": marks.push({ mcid: null, unattributed: false, inline: null, emitted: false }); break;
      // A sequence opened in an earlier stream of the page closes here.
      case "EMC": marks.pop(); break;
      case "Tj":
      case "'": {
        const bytes = stringBytes(last);
        if (!bytes) return null;
        show(bytes);
        break;
      }
      case "\"": {
        const bytes = stringBytes(last);
        if (!bytes) return null;
        show(bytes);
        break;
      }
      case "TJ": {
        if (last?.kind !== "array") return null;
        const parts = last.items.flatMap((item) => (item.kind === "string" ? [item.bytes] : []));
        const codes = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
        let offset = 0;
        for (const part of parts) { codes.set(part, offset); offset += part.length; }
        show(codes);
        break;
      }
      default: break;
    }
    operands = [];
  }
  return Object.freeze({ shows: Object.freeze(shows), sequences });
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) if (left[index] !== right[index]) return false;
  return true;
}

/** Whether the shows of one MCID are the same, font and codes, in both. */
function sameShows(left: readonly MarkedShow[], right: readonly MarkedShow[]): boolean {
  return left.length === right.length && left.every((show, index) =>
    show.font === right[index].font && show.inlineActualText === right[index].inlineActualText &&
    show.inlineText === right[index].inlineText && sameBytes(show.codes, right[index].codes));
}

// -- the element's text after the edit -----------------------------------------

const WHITE_SPACE = /\s/u;

function commonPrefix(left: string, right: string): number {
  const limit = Math.min(left.length, right.length);
  let length = 0;
  while (length < limit && left[length] === right[length]) length += 1;
  return length;
}

function commonSuffix(left: string, right: string, prefix: number): number {
  const limit = Math.min(left.length, right.length) - prefix;
  let length = 0;
  while (length < limit && left[left.length - 1 - length] === right[right.length - 1 - length]) length += 1;
  return length;
}

/** Where in `actual` the content of `reading` before `position` ends, read
 * by the characters that are not white space: the place after the `count`th
 * such character (0 for none). */
function afterNonSpaces(text: string, count: number): number {
  if (count === 0) return 0;
  let seen = 0;
  for (let index = 0; index < text.length; index += 1) {
    if (!WHITE_SPACE.test(text[index])) {
      seen += 1;
      if (seen === count) return index + 1;
    }
  }
  return -1;
}

/** The place of the `count + 1`th character that is not white space, or the
 * end of the text when there is none. */
function atNonSpace(text: string, count: number): number {
  let seen = 0;
  for (let index = 0; index < text.length; index += 1) {
    if (!WHITE_SPACE.test(text[index])) {
      if (seen === count) return index;
      seen += 1;
    }
  }
  return text.length;
}

function nonSpaces(text: string): number {
  let count = 0;
  for (const character of text) if (!WHITE_SPACE.test(character)) count += 1;
  return count;
}

/**
 * The /ActualText an element should hold after its marked content read
 * `before` and now reads `after`, from the /ActualText `actual` it holds
 * now; null when that cannot be decided.
 *
 * - Where `actual` is the text its content reads as (Word and PowerPoint
 *   repeat each run's text), it becomes the new reading. White space is not
 *   part of that comparison -- a producer may draw a word space as a gap and
 *   still name it -- and the edit is placed among `actual`'s own spaces.
 * - Where it is a real replacement (a ligature's letters, a hyphenation
 *   joined, a math label), the edit is spliced into it only where the two
 *   agree around it: the changed part of the reading lies in a stretch the
 *   replacement spells the same.
 */
export function pdfNativeStructActualTextAfterEdit(actual: string, before: string, after: string): string | null {
  if (before === after) return actual;
  if (actual === before) return after;
  const prefix = commonPrefix(before, after);
  const suffix = commonSuffix(before, after, prefix);
  const removedEnd = before.length - suffix;
  const inserted = after.slice(prefix, after.length - suffix);
  // The replacement agrees with the reading before the edit and after it.
  const agreedPrefix = commonPrefix(actual, before);
  const agreedSuffix = commonSuffix(actual, before, agreedPrefix);
  if (removedEnd <= agreedPrefix) {
    return actual.slice(0, prefix) + inserted + actual.slice(removedEnd);
  }
  if (prefix >= before.length - agreedSuffix) {
    const shift = actual.length - before.length;
    return actual.slice(0, prefix + shift) + inserted + actual.slice(removedEnd + shift);
  }
  // The same text but for its white space: the edit is placed by the
  // characters that are not.
  const strip = (text: string) => text.replace(/\s+/gu, "");
  if (strip(actual) !== strip(before)) return null;
  const removed = before.slice(prefix, removedEnd);
  const kept = nonSpaces(before.slice(0, prefix));
  const removedCount = nonSpaces(removed);
  let start: number;
  let end: number;
  if (removed === "") {
    start = end = afterNonSpaces(actual, kept);
  } else {
    start = WHITE_SPACE.test(removed[0]) ? afterNonSpaces(actual, kept) : atNonSpace(actual, kept);
    end = WHITE_SPACE.test(removed[removed.length - 1])
      ? atNonSpace(actual, kept + removedCount)
      : afterNonSpaces(actual, kept + removedCount);
  }
  if (start < 0 || end < 0 || end < start) return null;
  return actual.slice(0, start) + inserted + actual.slice(end);
}

/**
 * The /ActualText an element should hold when its content cannot be read but
 * the edited line's text can (as the reader reads it, before and after):
 * the line's text where it stands in `actual`, once, replaced -- exactly, or
 * else by the characters that are not white space. Null when the line does
 * not stand in `actual` exactly once.
 */
export function pdfNativeStructActualTextAfterLineEdit(actual: string, before: string, after: string): string | null {
  if (before === after) return actual;
  if (!before.trim()) return null;
  const exact = actual.indexOf(before);
  if (exact >= 0) {
    return actual.indexOf(before, exact + 1) < 0 ? actual.slice(0, exact) + after + actual.slice(exact + before.length) : null;
  }
  const kept: number[] = [];
  let stripped = "";
  for (let index = 0; index < actual.length; index += 1) {
    if (!WHITE_SPACE.test(actual[index])) {
      kept.push(index);
      stripped += actual[index];
    }
  }
  const needle = before.replace(/\s+/gu, "");
  const at = stripped.indexOf(needle);
  if (at < 0 || stripped.indexOf(needle, at + 1) >= 0) return null;
  return actual.slice(0, kept[at]) + after + actual.slice(kept[at + needle.length - 1] + 1);
}

// -- the revision ---------------------------------------------------------------

/** Reads a show's codes as text in the font a resource name selects; null
 * when the font or one of its codes cannot be read, "absent" when the page
 * names no such font (a face the write itself installs). */
export type PdfNativeStructTextDecoder = (
  fontName: string,
) => Promise<((codes: Uint8Array) => string | null) | null | "absent">;

/** Whether a document's structure gives any marked content of a page an
 * /ActualText (the cheap question asked before reading any stream). */
export async function pdfNativePageHasStructActualText(
  sourceBytes: Uint8Array,
  page: ObjectRef,
): Promise<boolean> {
  return (await readPdfStructActualTextIndex(sourceBytes)).pagesWithActualText.has(refKey(page));
}












export async function planMraiNativeStructActualTextRevisions(input: {
  readonly sourceBytes: Uint8Array;
  readonly page: ObjectRef;
  readonly sourceStream: Uint8Array;
  readonly editedStream: Uint8Array;
  readonly decoder: PdfNativeStructTextDecoder;
  /** A face the page does not name is one the write installs, and the
   * writer that installs it plans the elements again with it: answer none
   * rather than refuse. */
  readonly deferAbsentFonts?: boolean;
  /** The edited line's text as the reader reads it, before and after: the
   * element's text where its content's own fonts cannot be read
   * (pdfNativeStructActualTextAfterLineEdit). */
  readonly lineTexts?: Readonly<{ before: string; after: string }>;
}): Promise<readonly PdfNativeStructActualTextRevision[] | "deferred"> {
  const index = await readPdfStructActualTextIndex(input.sourceBytes);
  if (!index.pagesWithActualText.has(refKey(input.page))) return [];
  const named = index.namedProperties.get(refKey(input.page));
  const before = readMarkedContent(input.sourceStream, named);
  const after = readMarkedContent(input.editedStream, named);
  if (!before || !after) return Object.freeze([]);
  const byMcid = (reading: MarkedContentReading) => {
    const groups = new Map<number, MarkedShow[]>();
    for (const show of reading.shows) {
      if (show.mcid === null) continue;
      const group = groups.get(show.mcid) ?? [];
      group.push(show);
      groups.set(show.mcid, group);
    }
    return groups;
  };
  const beforeShows = byMcid(before);
  const afterShows = byMcid(after);
  // Text whose MCID is not read here may be any element's.
  const unattributed = (reading: MarkedContentReading) => reading.shows.filter((show) => show.unattributed);
  if (!sameShows(unattributed(before), unattributed(after))) return Object.freeze([]);
  const changed = new Set<number>();
  for (const mcid of new Set([...beforeShows.keys(), ...afterShows.keys()])) {
    if (!sameShows(beforeShows.get(mcid) ?? [], afterShows.get(mcid) ?? [])) changed.add(mcid);
  }
  const elements = new Map<string, StructActualTextElement>();
  for (const mcid of changed) {
    const owner = index.owners.get(pdfStructActualTextKey(input.page, mcid));
    // Every element naming the MCID follows its content; one reached from
    // several parents is one object, written once.
    for (const element of owner?.actuals ?? []) {
      // Not an object of its own, or reaching content beyond this page's
      // stream: left as it stands.
      if (
        !element.ref ||
        element.mcids.length === 0 ||
        element.mcids.some((entry) => entry.viaStream || !entry.page || refKey(entry.page) !== refKey(input.page)) ||
        !element.mcids.some((entry) => entry.mcid === mcid)
      ) {
        continue;
      }
      elements.set(refKey(element.ref), element);
    }
  }
  const decoders = new Map<string, ((codes: Uint8Array) => string | null) | null | "absent">();
  let absent = false;
  const read = async (shows: readonly MarkedShow[]): Promise<string | null> => {
    let text = "";
    for (const show of shows) {
      // A sequence with an inline /ActualText reads as that text, once.
      if (show.inlineActualText) {
        if (show.inlineText === null) return null;
        if (show.inlineFirst) text += show.inlineText;
        continue;
      }
      if (show.font === null) return null;
      if (!decoders.has(show.font)) decoders.set(show.font, await input.decoder(show.font));
      const decode = decoders.get(show.font);
      if (decode === "absent") absent = true;
      const part = decode && decode !== "absent" ? decode(show.codes) : null;
      if (part === null) return null;
      text += part;
    }
    return text;
  };
  const revisions: PdfNativeStructActualTextRevision[] = [];
  for (const element of elements.values()) {
    const ref = element.ref!;
    // Naming marked content another stream draws: left as it stands.
    if (element.mcids.some((entry) => !before.sequences.has(entry.mcid))) continue;
    const ordered = element.mcids.map((entry) => entry.mcid);
    const readingBefore = await read(ordered.flatMap((mcid) => beforeShows.get(mcid) ?? []));
    const readingAfter = await read(ordered.flatMap((mcid) => afterShows.get(mcid) ?? []));
    const unread = readingBefore === null || readingAfter === null;
    if (unread && absent && input.deferAbsentFonts) return "deferred";
    if (!unread && readingBefore === readingAfter) continue;
    const object = await inspectMraiNativeObject(input.sourceBytes, ref);
    const found = readMraiNativeStructElementActualText(object.objectBytes);
    const previousBytes = found ? decodeMraiNativePdfStringToken(found.token) : null;
    const actual = previousBytes ? decodeMraiNativePdfTextString(previousBytes, false) : null;
    if (!found || actual === null || object.streamBytes !== null) continue;
    let text: string;
    if (unread) {
      // A font that names no characters for its codes: the edited line's
      // own text, as the reader reads it, is placed where it stands in the
      // element's text.
      const placed = input.lineTexts
        ? pdfNativeStructActualTextAfterLineEdit(actual, input.lineTexts.before, input.lineTexts.after)
        : null;
      if (placed === null) continue;
      text = placed;
    } else {
      // A replacement the edit cannot be placed in (it spells the edited
      // part its own way) takes what the content now reads: the element then
      // says what the page shows, rather than the edit going to a text box.
      text = pdfNativeStructActualTextAfterEdit(actual, readingBefore!, readingAfter!) ?? readingAfter!;
    }
    if (text === actual) continue;
    const value = encodeMraiNativePdfTextStringToken(text, found.token);
    if (!value) continue;
    revisions.push(Object.freeze({
      objectNumber: ref.objectNumber,
      generation: ref.generation,
      expectedObjectSha256: object.objectSha256,
      page: input.page,
      mcids: Object.freeze([...new Set(ordered)].sort((left, right) => left - right)),
      previous: found.token,
      value,
      text,
    }));
  }
  return Object.freeze(revisions.sort((left, right) =>
    left.objectNumber - right.objectNumber || left.generation - right.generation));
}
