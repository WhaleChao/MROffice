/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Delete one original text object from a page, natively.
 *
 * The selected run's whole `BT … ET` object is replaced by an empty `BT ET`:
 * a text object that shows nothing. Every other byte of the page is kept, the
 * replacement is exactly five bytes long whatever the object held, and Undo
 * writes the recorded bytes back over that empty object.
 *
 * A text object that draws several shows (TeX and Skia put a whole paragraph
 * in one) keeps the others: only the selected shows are emptied, each into the
 * form of its own operator that shows nothing (`<> Tj`, `[] TJ`, `() '`), so
 * the state and the line moves the operators carried stay in place. That is only
 * exact when the text after the show is placed afresh rather than continuing
 * from where the show ended, which the planner proves before it writes. A
 * single show the next one draws on from is emptied into `[n] TJ` instead, a
 * show of nothing but its own advance, with `n` read off the run's text proof.
 */
import { LATIN1_EXACT } from "./latin1";
import {
  MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
} from "./vocabulary";
export {
  MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
};
import {
  inspectMraiNativeFormHostedContentStream,
  pdfNativeListTextSourceRunSpan,
  pdfNativeTextShowEdgeAdjustments,
  proveMraiNativeListTextSelectionForWriter,
  type PdfNativeListTextFidelityEvidence,
  type PdfNativeListTextSelectionEvidence,
} from "./list-text-fidelity";
import {
  inspectMraiNativeConsolidatedContentStream,
  inspectMraiNativeObject,
  inspectMraiNativePageContentStream,
  inspectMraiNativePageContentStreams,
  mraiNativePageSharesContent,
  nextMraiNativeIndirectObjectNumber,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeFormInstanceIsolationPlan,
  type PdfNativeObjectPrecondition,
  type PdfNativePageContentConsolidationPlan,
  type PdfNativePreparedContentStreamTransaction,
} from "./content-stream-transaction";




export const MRAI_NATIVE_EMPTY_TEXT_OBJECT = "BT ET" as const;

export type PdfNativeTextObjectDeleteDirection = "delete" | "restore";

/** Whether the whole `BT … ET` object was emptied, one show inside it,
 * several whole objects one run spans (a line LibreOffice or Word writes with
 * a link or a bold word in objects of their own), or the shows of a run that
 * starts or ends partway through the objects it crosses (a browser writes a
 * link's text and the full stop after it in objects of their own). */
export type PdfNativeTextObjectDeleteScope = "object" | "show" | "objects" | "run";

/**
 * The page content stream that owns the selected run. A page drawn from
 * several streams is one drawing, so the span alone does not say which part
 * of it the run sits in; the text proof that produced the span names this.
 */
export type PdfNativeTextObjectStreamRef = Readonly<{
  objectNumber: number;
  generation: number;
}>;

export interface PdfNativeTextObjectDeletePlan {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION;
  readonly direction: PdfNativeTextObjectDeleteDirection;
  readonly scope: PdfNativeTextObjectDeleteScope;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  /** The deleted bytes (the object, or the one show): removed by a delete,
   * written by a restore. */
  readonly objectBytes: readonly number[];
  /** A show deleted inside a line whose next show draws on from its end is
   * left as a show of nothing but its own advance (`[n] TJ`), so that next
   * show stays where it was; the proof of the run says what the advance is. */
  readonly penRestore?: Readonly<{
    number: string;
    sourceFidelity: PdfNativeListTextFidelityEvidence;
    /** The character spacing the show is drawn under, when it is not zero:
     * the pen is written where the reader moves as the show moved it
     * (mraiNativeShowPenForm). */
    characterSpacing?: number;
    /** The shows before the last one that also keep their advance, in the
     * order the span holds them. A reader hands over as one item what the
     * page may draw as several shows -- a typesetter starts a new one
     * wherever the font changes, which pdfTeX does for a digit inside a
     * word -- and each of them is emptied where it is and left as its own
     * advance, measured in its own font. They are the shows immediately
     * before the last: a show whose text is placed afresh is one every
     * show before it is placed afresh behind too. */
    precedingShows?: readonly Readonly<{
      number: string;
      sourceFidelity: PdfNativeListTextFidelityEvidence;
      characterSpacing?: number;
    }>[];
  }>;
  /** Shows the enclosing text object holds. */
  readonly showCount: number;
  readonly request: PdfNativeContentStreamTransactionRequest;
}

const LATIN1 = LATIN1_EXACT;
const SHOW_OPERATOR = /(?:\bT[jJ]\b|'|")/g;

function fail(
  code: "invalid-request" | "unsupported-document" | "precondition-mismatch",
  message: string,
): never {
  throw new PdfNativeTransactionError(code, message);
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (character) => character.charCodeAt(0) & 0xff);
}

interface ResolvedTextPage extends PdfNativeTextRunDrawing {
  readonly source: string;
}

async function resolveTextPage(
  sourceBytes: Uint8Array,
  pageNumber: number,
  contentStream: PdfNativeTextObjectStreamRef | null = null,
  runSpan: Readonly<{ start: number; end: number }> = { start: 0, end: 0 },
): Promise<ResolvedTextPage> {
  const drawing = await resolveMraiNativeTextRunDrawing(
    sourceBytes,
    pageNumber,
    contentStream,
    runSpan,
    "Deleting a text object needs one exact page content stream.",
  );
  return Object.freeze({ ...drawing, source: LATIN1.decode(drawing.streamBytes) });
}

export interface PdfNativeTextObjectSpan {
  readonly start: number;
  readonly end: number;
  readonly showCount: number;
}

/** The drawing a run lives in, as a rewrite addresses it. */
export interface PdfNativeTextRunDrawing {
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly streamBytes: Uint8Array;
  readonly identity: PdfNativeContentStreamTransactionRequest["identity"];
  readonly objectSha256: string;
  readonly streamSha256: string;
  readonly untouchedObjects: PdfNativeContentStreamTransactionRequest["untouchedObjects"];
  /** The run's span in this drawing. */
  readonly runSpan: Readonly<{ start: number; end: number }>;
  /** Set when the drawing is the page's streams joined as one new stream. */
  readonly pageContentConsolidation?: PdfNativePageContentConsolidationPlan;
  /** Set when the drawing is a Form XObject the document draws in more than
   * one place: the chain cloned so the change stays on this page. */
  readonly formInstanceIsolation?: PdfNativeFormInstanceIsolationPlan;
}

/**
 * The drawing a run is rewritten in: the content stream (or form) its span is
 * an offset into -- unless the text object the run starts in is not closed in
 * that stream. A producer may split a page's content into several streams
 * anywhere, even inside a text object; a reader draws them as one sequence,
 * so the page is rewritten as that one stream (the text replacement and the
 * picture arrange consolidate the same way) and the run's span moves into it.
 */
export async function resolveMraiNativeTextRunDrawing(
  sourceBytes: Uint8Array,
  pageNumber: number,
  contentStream: Readonly<{ objectNumber: number; generation: number }> | null,
  runSpan: Readonly<{ start: number; end: number }>,
  singleStreamRefusal: string,
): Promise<PdfNativeTextRunDrawing> {
  const inspected = await (async () => {
    try {
      return await inspectMraiNativePageContentStream(
        sourceBytes,
        pageNumber,
        contentStream,
        singleStreamRefusal,
      );
    } catch (error) {
      // The run may be drawn by a Form XObject the page invokes; its span is
      // then an offset into the form's own stream.
      if (
        contentStream &&
        error instanceof PdfNativeTransactionError &&
        error.message === "The named content stream does not draw this PDF page."
      ) {
        return inspectMraiNativeFormHostedContentStream(sourceBytes, pageNumber, contentStream);
      }
      throw error;
    }
  })();
  const own = Object.freeze({
    pageNumber: inspected.pageNumber,
    page: inspected.page,
    contentStream: inspected.contentStream,
    streamBytes: inspected.streamBytes,
    identity: inspected.identity,
    objectSha256: inspected.objectSha256,
    streamSha256: inspected.streamSha256,
    untouchedObjects: inspected.untouchedObjects,
    ...("formInstanceIsolation" in inspected && inspected.formInstanceIsolation
      ? { formInstanceIsolation: inspected.formInstanceIsolation }
      : {}),
    runSpan: Object.freeze({ start: runSpan.start, end: runSpan.end }),
  });
  const splitsTextObject = (() => {
    try {
      enclosingTextObject(LATIN1.decode(inspected.streamBytes), runSpan.start);
      return false;
    } catch (error) {
      return (
        error instanceof PdfNativeTransactionError &&
        error.message === "PDF text object nesting is unbalanced."
      );
    }
  })();
  // A page that draws a stream another page draws too is rewritten on its
  // own joined copy the same way, so the change stays on this page.
  const drawn = contentStream ?? inspected.contentStream;
  const shared = !("formInstanceIsolation" in inspected && inspected.formInstanceIsolation) &&
    await mraiNativePageSharesContent(sourceBytes, inspected.pageNumber);
  if ((!splitsTextObject || !contentStream) && !shared) return own;
  const streams = await inspectMraiNativePageContentStreams(sourceBytes, pageNumber);
  const index = streams.findIndex(
    (item) =>
      item.contentStream.objectNumber === drawn.objectNumber &&
      item.contentStream.generation === drawn.generation,
  );
  if ((streams.length < 2 && !shared) || index < 0) return own;
  const joined = new Uint8Array(streams.reduce((sum, item) => sum + item.streamBytes.length, 0));
  let offset = 0;
  let joinedStart = 0;
  for (const [position, item] of streams.entries()) {
    if (position === index) joinedStart = offset;
    joined.set(item.streamBytes, offset);
    offset += item.streamBytes.length;
  }
  const pageObject = await inspectMraiNativeObject(sourceBytes, inspected.page);
  const plan: PdfNativePageContentConsolidationPlan = Object.freeze({
    pageNumber: inspected.pageNumber,
    page: Object.freeze({ ...inspected.page, expectedObjectSha256: pageObject.objectSha256 }),
    sourceStreams: Object.freeze(
      streams.map((item) =>
        Object.freeze({ ...item.contentStream, expectedObjectSha256: item.objectSha256 }),
      ),
    ),
    joinedDecodedSha256: await sha256PdfBytes(joined),
    stream: Object.freeze({
      objectNumber: nextMraiNativeIndirectObjectNumber(sourceBytes),
      generation: 0 as const,
    }),
  });
  const consolidated = await inspectMraiNativeConsolidatedContentStream(sourceBytes, plan);
  // The page is rewritten to draw the new stream, so it is no longer pinned;
  // every source stream stays a witness by its hash (mirror of the picture
  // arrange's consolidation).
  const pageKey = `${inspected.page.objectNumber}:${inspected.page.generation}`;
  const pins = new Map<string, PdfNativeObjectPrecondition>();
  for (const item of streams.flatMap((entry) => entry.untouchedObjects ?? [])) {
    const key = `${item.objectNumber}:${item.generation}`;
    if (key !== pageKey) pins.set(key, item);
  }
  for (const item of streams) {
    pins.set(
      `${item.contentStream.objectNumber}:${item.contentStream.generation}`,
      Object.freeze({
        ...item.contentStream,
        expectedObjectSha256: item.objectSha256,
        role: "untouched-object" as const,
      }),
    );
  }
  return Object.freeze({
    pageNumber: inspected.pageNumber,
    page: inspected.page,
    contentStream: Object.freeze({ ...plan.stream }),
    streamBytes: consolidated.streamBytes,
    identity: consolidated.identity,
    objectSha256: consolidated.objectSha256,
    streamSha256: consolidated.streamSha256,
    untouchedObjects: Object.freeze([...pins.values()]),
    runSpan: Object.freeze({ start: joinedStart + runSpan.start, end: joinedStart + runSpan.end }),
    pageContentConsolidation: plan,
  });
}

import { MRAI_NATIVE_TEXT_RULES_NOT_CARRIED_MESSAGE, mraiNativeTextObjectOperators, pdfNativeTextRulesStanding, type PdfNativeTextRuleGeometry, type PdfNativeTextRuleStyle } from "./text-rule-block";
export { MRAI_NATIVE_TEXT_RULES_NOT_CARRIED_MESSAGE, mraiNativeTextObjectOperators };

/**
 * The `BT … ET` object that owns one byte offset. Nested or unterminated text
 * objects are refused rather than guessed at.
 */
export function enclosingTextObject(
  source: string,
  offset: number,
): PdfNativeTextObjectSpan {
  let open: number | null = null;
  for (const { index, operator } of mraiNativeTextObjectOperators(source)) {
    if (operator === "BT") {
      if (open !== null) {
        fail("unsupported-document", "Nested PDF text objects are unsupported.");
      }
      open = index;
      continue;
    }
    if (open === null) {
      fail("unsupported-document", "PDF text object nesting is unbalanced.");
    }
    const end = index + 2;
    if (offset >= open && offset < end) {
      const body = source.slice(open + 2, index);
      return Object.freeze({
        start: open,
        end,
        showCount: (body.match(SHOW_OPERATOR) ?? []).length,
      });
    }
    open = null;
  }
  if (open !== null) {
    fail("unsupported-document", "PDF text object nesting is unbalanced.");
  }
  fail("precondition-mismatch", "Selected text is not inside a text object.");
}

/**
 * End of the content token that starts at `index`, or null when the bytes
 * there are not one token this writer understands. Strings, hex strings,
 * arrays and inline dictionaries are skipped as wholes.
 */
function contentTokenEnd(source: string, index: number): number | null {
  const first = source[index];
  if (first === "(") {
    let depth = 0;
    for (let i = index; i < source.length; i += 1) {
      const character = source[i];
      if (character === "\\") {
        i += 1;
        continue;
      }
      if (character === "(") depth += 1;
      else if (character === ")") {
        depth -= 1;
        if (depth === 0) return i + 1;
      }
    }
    return null;
  }
  if (first === "<") {
    if (source[index + 1] === "<") {
      // A dictionary holds tokens of its own -- a hex string, a nested
      // dictionary (`/Span <</ActualText <FEFF005F>>> BDC`) -- so its end
      // is the `>>` after its last token, not the first `>>` in sight.
      let i = index + 2;
      for (;;) {
        while (i < source.length && /\s/.test(source[i])) i += 1;
        if (i >= source.length) return null;
        if (source.startsWith(">>", i)) return i + 2;
        const end = contentTokenEnd(source, i);
        if (end === null) return null;
        i = end;
      }
    }
    const close = source.indexOf(">", index + 1);
    if (close < 0 || !/^[0-9A-Fa-f\s]*$/.test(source.slice(index + 1, close))) {
      return null;
    }
    return close + 1;
  }
  if (first === "[") {
    let i = index + 1;
    for (;;) {
      while (i < source.length && /\s/.test(source[i])) i += 1;
      if (i >= source.length) return null;
      if (source[i] === "]") return i + 1;
      const end = contentTokenEnd(source, i);
      if (end === null) return null;
      i = end;
    }
  }
  const rest = source.slice(index);
  if (first === "/") return index + /^\/[^\s[\]()<>{}/%]*/.exec(rest)![0].length;
  if (/[+\-.0-9]/.test(first)) return index + /^[+\-.0-9]+/.exec(rest)![0].length;
  const operator = /^[A-Za-z*'"][A-Za-z0-9*'"]*/.exec(rest);
  return operator ? index + operator[0].length : null;
}

interface ContentToken {
  readonly start: number;
  readonly end: number;
  readonly operator: boolean;
}

/** The tokens of `source[from, to)`, or null when any of it is not readable. */
/** The content tokens of `source[from, to)`, or null when the span holds a
 * comment or a token this reader does not understand. */
export function mraiNativeContentTokens(
  source: string,
  from: number,
  to: number,
): readonly Readonly<{ start: number; end: number; operator: boolean }>[] | null {
  return contentTokens(source, from, to);
}

function contentTokens(
  source: string,
  from: number,
  to: number,
): readonly ContentToken[] | null {
  const tokens: ContentToken[] = [];
  let i = from;
  while (i < to) {
    if (/\s/.test(source[i])) {
      i += 1;
      continue;
    }
    if (source[i] === "%") return null;
    const end = contentTokenEnd(source, i);
    if (end === null || end > to) return null;
    tokens.push({ start: i, end, operator: /^[A-Za-z'"]/.test(source[i]) });
    i = end;
  }
  return tokens;
}

/** One show operator inside a span: its operands and operator, as a whole. */
interface ShowOperator {
  readonly start: number;
  readonly end: number;
  readonly emptyForm: string;
}

/**
 * The text-showing operators of `source[from, to)`, each with the bytes that
 * show nothing in its place while keeping every other effect the operator
 * had: `'` still moves to the next line, `"` still sets the spacings. Null
 * when the span is not readable or a show's operands are not what its
 * operator takes.
 */
function showOperatorsIn(
  source: string,
  from: number,
  to: number,
): readonly ShowOperator[] | null {
  const tokens = contentTokens(source, from, to);
  if (!tokens) return null;
  const shows: ShowOperator[] = [];
  let operands: ContentToken[] = [];
  // A literal string is a string whatever its first character is: `(<) Tj`
  // draws the glyph whose code is `<` (a LaTeX table sets its "]" with it).
  // Only a hex string is told from a dictionary by a second `<`. (Mirror of
  // the replay's _show_operators_in.)
  const isString = (token: ContentToken): boolean =>
    source[token.start] === "(" ||
    (source[token.start] === "<" && source[token.start + 1] !== "<");
  const isNumber = (token: ContentToken): boolean =>
    /^[+\-.0-9]+$/.test(source.slice(token.start, token.end));
  for (const token of tokens) {
    if (!token.operator) {
      operands.push(token);
      continue;
    }
    const operator = source.slice(token.start, token.end);
    let emptyForm: string | null = null;
    if (operator === "Tj" || operator === "'") {
      if (operands.length !== 1 || !isString(operands[0])) return null;
      emptyForm = operator === "Tj" ? "<> Tj" : "() '";
    } else if (operator === "TJ") {
      if (operands.length !== 1 || source[operands[0].start] !== "[") return null;
      emptyForm = "[] TJ";
    } else if (operator === '"') {
      if (
        operands.length !== 3 ||
        !isNumber(operands[0]) ||
        !isNumber(operands[1]) ||
        !isString(operands[2])
      ) {
        return null;
      }
      emptyForm = `${source.slice(operands[0].start, operands[1].end)} () "`;
    }
    if (emptyForm !== null) {
      shows.push({
        start: operands.length > 0 ? operands[0].start : token.start,
        end: token.end,
        emptyForm,
      });
    }
    operands = [];
  }
  return operands.length === 0 ? shows : null;
}

/**
 * The bytes that show nothing where the shows of `span` showed text: every
 * show is replaced by the empty form of its own operator and every other
 * byte of the span is kept. Null when the span holds no show, is not readable,
 * or a show's operands are not what its operator takes.
 */
export function mraiNativeEmptyShowsForm(
  span: string,
  /** The advances the last shows of the span keep, in the order the span
   * holds them: each of those shows is left as `[n] TJ` instead of as a show
   * of nothing, so the text that draws on from the span stays where it was.
   * A show that keeps its advance is one every show after it keeps one too,
   * because they all look past the span for the text that follows. */
  pens: readonly string[] = [],
  /** The character spacing each of those shows is drawn under, in the same
   * order (none: zero): the pen is written where the reader moves as the
   * show moved it (mraiNativeShowPenForm). */
  spacings: readonly number[] = [],
): string | null {
  const shows = showOperatorsIn(span, 0, span.length);
  if (!shows || shows.length === 0 || pens.length > shows.length || spacings.length > pens.length) return null;
  const first = shows.length - pens.length;
  let out = "";
  let cursor = 0;
  for (const [index, show] of shows.entries()) {
    const pen = index >= first
      ? mraiNativeShowPenForm(span.slice(show.start, show.end), pens[index - first], spacings[index - first] ?? 0)
      : show.emptyForm;
    if (pen === null) return null;
    out += span.slice(cursor, show.start) + pen;
    cursor = show.end;
  }
  return out + span.slice(cursor);
}

/** A show left as nothing but its own advance `pen`: `[pen] TJ`, written
 * where the reader moves its pen as the show moved it under the character
 * spacing `spacing` it is drawn under (readerPenForm). pdf.js reads a
 * non-zero number with no string of codes before it as an empty chunk and
 * moves by the spacing as well, so a show whose own elements hold no such
 * number is left as `0 Tc [pen] TJ <spacing> Tc` under a non-zero spacing.
 * Null when it cannot be written. (Mirror of replay's _show_pen_form.) */
function mraiNativeShowPenForm(show: string, pen: string, spacing: number): string | null {
  const elements = showElements(show);
  return readerPenForm([], [{ text: pen, codeBytes: null }], [], elements, 0, elements?.length ?? 0, spacing);
}

/** The one show of `span` left as its own advance `pen`, for the paragraph
 * rewrite: `[pen] TJ` (mraiNativeEmptyShowsForm), written where the reader
 * moves its pen as the show moved it under the character spacing `spacing`
 * it is drawn under (readerPenForm). Null when it cannot be. (Mirror of
 * replay's _empty_show_pen_form.) */
export function mraiNativeEmptyShowPenForm(span: string, pen: string, spacing: number): string | null {
  const shows = showOperatorsIn(span, 0, span.length);
  if (!shows || shows.length !== 1) return null;
  const [show] = shows;
  const form = mraiNativeShowPenForm(span.slice(show.start, show.end), pen, spacing);
  return form === null ? null : span.slice(0, show.start) + form + span.slice(show.end);
}

/** The elements of one show as a TJ array: a `Tj`, `'` or `"` draws its one
 * string the way a `TJ` of that string alone does. Null when it is none. */
function showElements(show: string): readonly Readonly<{ text: string; codeBytes: number | null }>[] | null {
  const array = tjArrayElements(show);
  if (array) return array;
  const single = /^\s*(?:[+\-.0-9]+\s+[+\-.0-9]+\s+)?([(<][\s\S]*[)>])\s*(?:Tj|'|")\s*$/u.exec(show);
  return single ? tjArrayElements(`[${single[1]}] TJ`) : null;
}

/** How many numbers of `elements[from, to)` the reader takes as a chunk of
 * no glyphs. pdf.js (buildTextContentItem, its `!chars` branch) reads a
 * non-zero TJ number with no string of codes between it and the array's
 * start or the non-zero number before it as an empty chunk, and moves its
 * pen for it by the character spacing as well as by the number -- a move
 * the renderer never makes (a zero number it passes over). */
function readerEmptyChunks(
  elements: readonly Readonly<{ text: string; codeBytes: number | null }>[],
  from: number,
  to: number,
): number {
  let count = 0;
  let glyphs = false;
  for (const [index, element] of elements.entries()) {
    if (index >= to) break;
    if (element.codeBytes !== null) {
      if (element.codeBytes > 0) glyphs = true;
      continue;
    }
    if (Number(element.text) === 0) continue;
    if (!glyphs && index >= from) count += 1;
    glyphs = false;
  }
  return count;
}

/** A character spacing as this writer spells it back: at most six decimals,
 * and only a value that spelling reads back exactly. */
function exactSpacing(value: number): string | null {
  if (!Number.isFinite(value) || Math.abs(value) > 1_000_000) return null;
  const spelled = value.toFixed(6).replace(/\.0+$/u, "").replace(/(\.\d*?)0+$/u, "$1");
  return Number(spelled) === value ? spelled : null;
}

/**
 * The array of a show emptied to a pen, written so that the reader moves
 * its pen over it as the renderer does. `before`, `region` and `after` are
 * the kept elements before the pen, the pen with the numbers that follow
 * it up to the next string of codes, and the rest; `source[from, to)` are
 * the elements of the source show the region stands for. Under a non-zero
 * character spacing the reader moves by that spacing once more for every
 * number it reads as an empty chunk (readerEmptyChunks), so the text after
 * the pen reads where it read before only while the region holds as many
 * of those as the source did: a region that holds more is written as its
 * own array under a zero spacing, which the reader reads as the renderer
 * draws it, and the spacing is set back after it. One that holds fewer
 * cannot be written; null, as is a spacing this writer cannot set back
 * exactly. (Mirror of replay's _reader_pen_form.)
 */
function readerPenForm(
  before: readonly Readonly<{ text: string; codeBytes: number | null }>[],
  region: readonly Readonly<{ text: string; codeBytes: number | null }>[],
  after: readonly Readonly<{ text: string; codeBytes: number | null }>[],
  source: readonly Readonly<{ text: string; codeBytes: number | null }>[] | null,
  from: number,
  to: number,
  spacing: number,
): string | null {
  const array = (elements: readonly Readonly<{ text: string }>[]) => `[${elements.map((element) => element.text).join(" ")}] TJ`;
  const whole = [...before, ...region, ...after];
  if (spacing === 0) return array(whole);
  if (source === null) return null;
  const written = readerEmptyChunks(whole, before.length, before.length + region.length);
  const read = readerEmptyChunks(source, from, to);
  if (written === read) return array(whole);
  if (read !== 0) return null;
  const spelled = exactSpacing(spacing);
  if (spelled === null) return null;
  return `${before.length ? `${array(before)} ` : ""}0 Tc ${array(region)} ${spelled} Tc${after.length ? ` ${array(after)}` : ""}`;
}

/** The elements of one `[...] TJ` array as the source spells them: each
 * string with the number of code bytes it holds, each number; null when the
 * show is not one TJ array. */
function tjArrayElements(
  span: string,
): readonly Readonly<{ text: string; codeBytes: number | null }>[] | null {
  const trimmed = span.trim();
  if (!trimmed.startsWith("[")) return null;
  const close = trimmed.lastIndexOf("]");
  if (close < 0 || !/^\]\s*TJ$/.test(trimmed.slice(close))) return null;
  const inner = trimmed.slice(1, close);
  const elements: { text: string; codeBytes: number | null }[] = [];
  let cursor = 0;
  while (cursor < inner.length) {
    const character = inner[cursor];
    if (/[\s\0]/.test(character)) { cursor += 1; continue; }
    if (character === "(") {
      let depth = 1;
      let bytes = 0;
      let index = cursor + 1;
      while (index < inner.length && depth > 0) {
        const c = inner[index];
        if (c === "\\") {
          const next = inner[index + 1];
          if (next === undefined) return null;
          if (/[0-7]/.test(next)) {
            let digits = 1;
            while (digits < 3 && /[0-7]/.test(inner[index + 1 + digits] ?? "")) digits += 1;
            bytes += 1;
            index += 1 + digits;
          } else if (next === "\r") {
            index += inner[index + 2] === "\n" ? 3 : 2;
          } else if (next === "\n") {
            index += 2;
          } else {
            bytes += 1;
            index += 2;
          }
          continue;
        }
        if (c === "(") depth += 1;
        else if (c === ")") { depth -= 1; if (depth === 0) { index += 1; break; } }
        if (c === "\r") { bytes += 1; index += inner[index + 1] === "\n" ? 2 : 1; continue; }
        bytes += 1;
        index += 1;
      }
      if (depth !== 0) return null;
      elements.push({ text: inner.slice(cursor, index), codeBytes: bytes });
      cursor = index;
      continue;
    }
    if (character === "<") {
      const end = inner.indexOf(">", cursor + 1);
      if (end < 0) return null;
      const digits = inner.slice(cursor + 1, end).replace(/[\s\0]/g, "");
      if (!/^[0-9A-Fa-f]*$/.test(digits)) return null;
      elements.push({ text: inner.slice(cursor, end + 1), codeBytes: Math.ceil(digits.length / 2) });
      cursor = end + 1;
      continue;
    }
    const number = /^[+-]?(?:\d+\.?\d*|\.\d+)/.exec(inner.slice(cursor));
    if (!number) return null;
    elements.push({ text: number[0], codeBytes: null });
    cursor += number[0].length;
  }
  return elements;
}

/** The code bytes one TJ string element holds, as its spelling decodes
 * them: a hex string digit pair by digit pair (an odd last digit padded with
 * 0), a literal string with its escapes read and every end of line a line
 * feed (PDF 32000-1 7.3.4.2). */
function tjStringBytes(text: string): number[] | null {
  if (text.startsWith("<")) {
    const digits = text.slice(1, -1).replace(/\s+/gu, "");
    if (!/^[0-9A-Fa-f]*$/u.test(digits)) return null;
    const padded = digits.length % 2 ? `${digits}0` : digits;
    const bytes: number[] = [];
    for (let index = 0; index < padded.length; index += 2) bytes.push(parseInt(padded.slice(index, index + 2), 16));
    return bytes;
  }
  if (!text.startsWith("(") || !text.endsWith(")")) return null;
  const inner = text.slice(1, -1);
  const escapes: Record<string, number> = { n: 10, r: 13, t: 9, b: 8, f: 12, "(": 40, ")": 41, "\\": 92 };
  const bytes: number[] = [];
  for (let index = 0; index < inner.length; index += 1) {
    const character = inner[index];
    if (character === "\\") {
      const next = inner[index + 1];
      if (next === undefined) return null;
      if (/[0-7]/u.test(next)) {
        let digits = next;
        while (digits.length < 3 && /[0-7]/u.test(inner[index + 1 + digits.length] ?? "")) digits += inner[index + 1 + digits.length];
        bytes.push(parseInt(digits, 8) & 0xff);
        index += digits.length;
      } else if (next === "\r") {
        index += inner[index + 2] === "\n" ? 2 : 1;
      } else if (next === "\n") {
        index += 1;
      } else {
        bytes.push(escapes[next] ?? (next.charCodeAt(0) & 0xff));
        index += 1;
      }
      continue;
    }
    if (character === "\r") {
      bytes.push(10);
      if (inner[index + 1] === "\n") index += 1;
      continue;
    }
    bytes.push(character.charCodeAt(0) & 0xff);
  }
  return bytes;
}

const hexTjString = (bytes: readonly number[]) =>
  `<${bytes.map((byte) => byte.toString(16).padStart(2, "0").toUpperCase()).join("")}>`;

/** A TJ show with one contiguous range of its codes emptied to a pen: the
 * strings of `[codeStart, codeEnd)` and the numbers between them become the
 * one number `pen` (their advance, measured by the caller), every other
 * element keeps its own spelling, and the array is written `[a b c] TJ`.
 * The range may begin or end inside a string -- the reader cuts a show at a
 * run of spaces inside one too (a manual's line leaves spaces where an
 * inline icon stands, and goes on after it) -- and the part of that string
 * outside the range is kept, written as a hex string. Null when the show is
 * not one TJ array or the range is not inside it. The paragraph rewrite
 * empties a show it owns only part of this way (a table of contents line
 * and its page number, an equation and its number, share one show), and
 * replay rebuilds the same form. */
export function mraiNativeEmptyChunkForm(
  span: string,
  codeStart: number,
  codeEnd: number,
  pen: string,
  /** The character spacing the show is drawn under (see readerPenForm). */
  spacing = 0,
): string | null {
  // A `Tj` draws its one string the way a `TJ` of that string alone does
  // (Tc and Tw included), so a show the reader cut at a space it reads wide
  // (Word sets a table cell's two labels with a large Tw) is emptied as one.
  const single = /^\s*([(<][\s\S]*[)>])\s*Tj\s*$/u.exec(span);
  const elements = tjArrayElements(span) ?? (single ? tjArrayElements(`[${single[1]}] TJ`) : null);
  if (!elements || !(codeStart < codeEnd)) return null;
  type Element = Readonly<{ text: string; codeBytes: number | null }>;
  const kept: Element[] = [];
  let offset = 0;
  let inside = false;
  let penAt = -1;
  let closed = false;
  // The source elements the pen stands for: from the string the range opens
  // in up to the first string whose codes are kept after it.
  let from = -1;
  let to = elements.length;
  for (const [index, element] of elements.entries()) {
    if (element.codeBytes === null) {
      if (!inside) kept.push(element);
      continue;
    }
    const start = offset;
    const end = offset + element.codeBytes;
    offset = end;
    if (closed && to === elements.length && element.codeBytes > 0) to = index;
    const opens = penAt < 0 && codeStart >= start && codeStart < end;
    if (!inside && !opens) {
      kept.push(element);
      continue;
    }
    const closes = codeEnd <= end;
    let bytes: number[] | null = null;
    if ((opens && codeStart > start) || (closes && codeEnd < end)) {
      bytes = tjStringBytes(element.text);
      if (!bytes || bytes.length !== element.codeBytes) return null;
    }
    if (opens) {
      if (codeStart > start) kept.push({ text: hexTjString(bytes!.slice(0, codeStart - start)), codeBytes: codeStart - start });
      penAt = kept.length;
      from = index;
      kept.push({ text: pen, codeBytes: null });
      inside = true;
    }
    if (closes) {
      if (codeEnd < end) {
        kept.push({ text: hexTjString(bytes!.slice(codeEnd - start)), codeBytes: end - codeEnd });
        to = index;
      }
      inside = false;
      closed = true;
    }
  }
  if (penAt < 0 || !closed) return null;
  // The pen and the numbers after it up to the next string of codes.
  let regionEnd = penAt + 1;
  while (regionEnd < kept.length && !(kept[regionEnd].codeBytes! > 0)) regionEnd += 1;
  return readerPenForm(kept.slice(0, penAt), kept.slice(penAt, regionEnd), kept.slice(regionEnd),
    elements, from, to, spacing);
}

/** The advances a plan's pen restore keeps, in the order the span holds
 * them: the shows before the last one, then the last one's own. */
function penRestoreNumbers(
  penRestore: PdfNativeTextObjectDeletePlan["penRestore"] | null | undefined,
): readonly string[] {
  if (!penRestore) return [];
  return [...(penRestore.precedingShows ?? []).map((show) => show.number), penRestore.number];
}

/** The character spacings those advances are drawn under, in the same order
 * (zero where a pen names none). */
function penRestoreSpacings(
  penRestore: PdfNativeTextObjectDeletePlan["penRestore"] | null | undefined,
): readonly number[] {
  if (!penRestore) return [];
  return [...(penRestore.precedingShows ?? []), penRestore].map((show) => show.characterSpacing ?? 0);
}

/** A pen's character spacing as a record names it: only when it is not
 * zero. */
function penSpacing(evidence: PdfNativeListTextFidelityEvidence): { characterSpacing?: number } {
  const spacing = evidence.sameLineLayoutState?.characterSpacing ?? 0;
  return spacing !== 0 ? { characterSpacing: spacing } : {};
}

/** The text-showing operators of `source[from, to)` (see showOperatorsIn),
 * for the paragraph rewrite, which empties the shows it owns one by one. */
export function mraiNativeShowOperatorsIn(
  source: string,
  from: number,
  to: number,
): readonly Readonly<{ start: number; end: number; emptyForm: string }>[] | null {
  return showOperatorsIn(source, from, to);
}

/** Whether emptying the show that ends at `from` moves nothing after it (see
 * textAfterShowIsPlacedAfresh); the paragraph rewrite asks it of every show
 * it empties. */
export function mraiNativeTextAfterShowIsPlacedAfresh(
  source: string,
  from: number,
  deletedEnd: number,
  objectEnd: number,
): boolean {
  return textAfterShowIsPlacedAfresh(source, from, deletedEnd, objectEnd);
}

/** Operators that place the text position afresh: after one of them, the
 * text no longer depends on how far the show before it reached. */
const TEXT_POSITION_RESETTING_OPERATORS = new Set(["Td", "TD", "Tm", "T*", "'", '"', "ET"]);
/** Operators that only change state; they may sit between a show and the
 * operator that places the next text. */
const TEXT_STATE_ONLY_OPERATORS = new Set([
  "Tf", "Tc", "Tw", "Tz", "TL", "Ts", "Tr",
  "g", "rg", "k", "G", "RG", "K", "cs", "sc", "scn", "CS", "SC", "SCN", "gs",
  // The line's own graphics state (a run drawn stroked for its weight sets
  // its line width, `2 Tr 0.42 w`, as the paragraph rewrite writes a
  // synthetic bold): none of it moves the text position. Read as text
  // drawn on from the show before, a line this program had set again
  
  "w", "J", "j", "M", "d", "ri", "i",
  "BMC", "BDC", "EMC", "MP", "DP",
]);

/**
 * Whether the text after `from` (inside the object that ends at `objectEnd`)
 * is placed afresh before anything else shows, so emptying the show before
 * it moves nothing.
 */
function textAfterShowIsPlacedAfresh(
  source: string,
  from: number,
  deletedEnd: number,
  objectEnd: number,
): boolean {
  const tokens = contentTokens(source, from, objectEnd);
  if (!tokens) return false;
  let operand: ContentToken | null = null;
  for (const token of tokens) {
    if (!token.operator) {
      operand = token;
      continue;
    }
    const operator = source.slice(token.start, token.end);
    const shown = operand ? source.slice(operand.start, operand.end) : null;
    operand = null;
    if (TEXT_POSITION_RESETTING_OPERATORS.has(operator)) return true;
    if (TEXT_STATE_ONLY_OPERATORS.has(operator)) continue;
    // A show that is emptied along with this one no longer needs the position.
    if ((operator === "Tj" || operator === "TJ") && token.end <= deletedEnd) continue;
    // A show of nothing (`()Tj`, `<>Tj`, `[]TJ`) draws nothing wherever it is.
    if ((operator === "Tj" || operator === "TJ") && shown !== null && /^(?:\(\)|<>|\[\s*\])$/.test(shown)) continue;
    return false;
  }
  return false;
}

/** Operators an emptied text object loses: its shows and the moves of its
 * text position, which end with it. */
const EMPTIED_TEXT_OBJECT_DROPPED = new Set(["Tj", "TJ", "'", "Td", "Tm", "T*"]);

/** Text state parameters: of each, only the setting in force after `ET`
 * matters. */
const TEXT_STATE_PARAMETER_OPERATORS = new Set(["Tf", "Tc", "Tw", "Tz", "TL", "Ts", "Tr"]);

/** A number token negated as text, a zero left unsigned. */
function negatedPdfNumber(value: string): string {
  if (/^[+-]?(?:0+\.?0*|\.0+)$/.test(value)) return value.replace(/^[+-]/, "");
  if (value.startsWith("-")) return value.slice(1);
  if (value.startsWith("+")) return `-${value.slice(1)}`;
  return `-${value}`;
}













export function mraiNativeEmptiedTextObject(object: string): string {
  if (!object.startsWith("BT") || !object.endsWith("ET") || object.length < 4) return MRAI_NATIVE_EMPTY_TEXT_OBJECT;
  const tokens = contentTokens(object, 2, object.length - 2);
  if (!tokens) return MRAI_NATIVE_EMPTY_TEXT_OBJECT;
  let kept: { operator: string; written: string }[] = [];
  let operands: string[] = [];
  const keep = (operator: string, written: string) => {
    if (TEXT_STATE_PARAMETER_OPERATORS.has(operator)) kept = kept.filter((item) => item.operator !== operator);
    kept.push({ operator, written });
  };
  for (const token of tokens) {
    const text = object.slice(token.start, token.end);
    if (!token.operator || text === "true" || text === "false" || text === "null") {
      operands.push(text);
      continue;
    }
    if (text === "TD") {
      if (operands.length !== 2) return MRAI_NATIVE_EMPTY_TEXT_OBJECT;
      keep("TL", `${negatedPdfNumber(operands[1])} TL`);
    } else if (text === '"') {
      if (operands.length !== 3) return MRAI_NATIVE_EMPTY_TEXT_OBJECT;
      keep("Tw", `${operands[0]} Tw`);
      keep("Tc", `${operands[1]} Tc`);
    } else if (!EMPTIED_TEXT_OBJECT_DROPPED.has(text)) {
      keep(text, [...operands, text].join(" "));
    }
    operands = [];
  }
  if (operands.length) return MRAI_NATIVE_EMPTY_TEXT_OBJECT;
  return kept.length ? `BT ${kept.map((item) => item.written).join(" ")} ET` : MRAI_NATIVE_EMPTY_TEXT_OBJECT;
}

/** The bytes a delete leaves behind for `plan`. */
/**
 * The bytes a run of whole text objects leaves behind: every `BT … ET` in the
 * span emptied to `BT ET`, and whatever the page draws between them -- the
 * colours, saves and the rule under a link -- kept as it is. Null unless the
 * span starts with one object, ends with another and holds only whole,
 * unnested objects between.
 */
export function mraiNativeEmptiedTextObjectsForm(
  span: string,
   
  legacy = false,
): string | null {
  const objects: { start: number; end: number }[] = [];
  let open: number | null = null;
  for (const { index, operator } of mraiNativeTextObjectOperators(span)) {
    if (operator === "BT") {
      if (open !== null) return null;
      open = index;
      continue;
    }
    if (open === null) return null;
    objects.push({ start: open, end: index + 2 });
    open = null;
  }
  if (
    open !== null ||
    objects.length < 2 ||
    objects[0].start !== 0 ||
    objects[objects.length - 1].end !== span.length
  ) {
    return null;
  }
  let emptied = "";
  let cursor = 0;
  for (const object of objects) {
    emptied += span.slice(cursor, object.start) +
      (legacy ? MRAI_NATIVE_EMPTY_TEXT_OBJECT : mraiNativeEmptiedTextObject(span.slice(object.start, object.end)));
    cursor = object.end;
  }
  return emptied;
}

/** A TJ pen number, written the way the text writer writes its numbers. */
function formatPenNumber(value: number): string {
  if (!Number.isFinite(value) || Math.abs(value) > 1_000_000) {
    fail("unsupported-document", "The deleted run's advance is not a number this writer writes.");
  }
  const normalized = Math.abs(value) < 0.0000005 ? 0 : value;
  return normalized
    .toFixed(6)
    .replace(/\.0+$/, "")
    .replace(/(\.\d*?)0+$/, "$1");
}

/**
 * The bytes a run crossing text objects leaves behind: every show of the
 * span emptied to its operator's empty form, and everything else -- the
 * objects' ends and starts, the colours between them -- kept as it is. With
 * a pen number the last show is left as a show of nothing but that advance,
 * so the text its object draws after it stays where it was. Null unless the
 * span is readable, holds a show and crosses an object's end.
 */
export function mraiNativeDeletedRunForm(
  span: string,
  pen: string | null,
  /** The character spacing the last show is drawn under (mraiNativeShowPenForm). */
  spacing = 0,
): string | null {
  const shows = showOperatorsIn(span, 0, span.length);
  if (!shows || shows.length === 0 || !/(?:^|[^/\w])ET\b/.test(span)) return null;
  let out = "";
  let cursor = 0;
  for (const [index, show] of shows.entries()) {
    const last = index === shows.length - 1;
    const form = last && pen ? mraiNativeShowPenForm(span.slice(show.start, show.end), pen, spacing) : show.emptyForm;
    if (form === null) return null;
    out += span.slice(cursor, show.start) + form;
    cursor = show.end;
  }
  return out + span.slice(cursor);
}

function deletedFormOf(
  scope: PdfNativeTextObjectDeleteScope,
  objectBytes: readonly number[],
  /** The advances the last shows of the span keep, in the order the span
   * holds them; a run crossing text objects keeps only its last show's. */
  pens: readonly string[] = [],
  /** The character spacings those shows are drawn under, in the same order. */
  spacings: readonly number[] = [],
): string {
  if (scope === "run") {
    const form = mraiNativeDeletedRunForm(
      LATIN1.decode(Uint8Array.from(objectBytes)),
      pens.at(-1) ?? null,
      spacings.at(-1) ?? 0,
    );
    if (form === null) fail("invalid-request", "The recorded span is not a run of shows.");
    return form;
  }
  if (scope === "show") {
    const empty = mraiNativeEmptyShowsForm(LATIN1.decode(Uint8Array.from(objectBytes)), pens, spacings);
    if (empty === null) fail("invalid-request", "The recorded text span holds no readable show.");
    return empty;
  }
  if (scope === "object") return mraiNativeEmptiedTextObject(LATIN1.decode(Uint8Array.from(objectBytes)));
  if (scope === "objects") {
    const emptied = mraiNativeEmptiedTextObjectsForm(
      LATIN1.decode(Uint8Array.from(objectBytes)),
    );
    if (emptied === null) fail("invalid-request", "The recorded span is not a run of whole text objects.");
    return emptied;
  }
  return MRAI_NATIVE_EMPTY_TEXT_OBJECT;
}





function deletedFormsOf(
  scope: PdfNativeTextObjectDeleteScope,
  objectBytes: readonly number[],
  pens: readonly string[] = [],
  spacings: readonly number[] = [],
): readonly string[] {
  const current = deletedFormOf(scope, objectBytes, pens, spacings);
  const legacy = scope === "object"
    ? MRAI_NATIVE_EMPTY_TEXT_OBJECT
    : scope === "objects" && pens.length === 0
      ? mraiNativeEmptiedTextObjectsForm(LATIN1.decode(Uint8Array.from(objectBytes)), true) ?? current
      : current;
  return legacy === current ? [current] : [current, legacy];
}

export async function planMraiNativeTextObjectDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  /** Byte span of the selected run inside the content stream that owns it. */
  readonly runSpan: Readonly<{ start: number; end: number }>;
  /** Which of the page's content streams that span is an offset into. */
  readonly contentStream?: PdfNativeTextObjectStreamRef | null;
  /** The reader selection the run was proven from, when the caller has one:
   * a show whose next show draws on from its end is then deleted keeping its
   * advance. */
  readonly textSelection?: PdfNativeListTextSelectionEvidence | null;
  /** The underline and strikethrough the editor read beside the run, with
   * the geometry they were drawn from: taken away in this same write. A run
   * drawn as several text objects refuses (MRAI_NATIVE_TEXT_RULES_NOT_CARRIED_MESSAGE)
   * and the caller takes them away first. */
  readonly keepTextRules?: Readonly<{ styles: readonly PdfNativeTextRuleStyle[]; geometry: PdfNativeTextRuleGeometry }> | null;
}): Promise<PdfNativeTextObjectDeletePlan> {
  if (
    !Number.isSafeInteger(input.runSpan.start) ||
    !Number.isSafeInteger(input.runSpan.end) ||
    input.runSpan.start < 0 ||
    input.runSpan.end <= input.runSpan.start
  ) {
    fail("invalid-request", "Selected text span is invalid.");
  }
  const page = await resolveTextPage(
    input.sourceBytes,
    input.pageNumber,
    input.contentStream ?? null,
    input.runSpan,
  );
  // The span as the drawing being rewritten reads it: the stream's own, or
  // moved into the page's streams joined as one.
  const runSpan = page.runSpan;
  if (runSpan.end > page.streamBytes.length) {
    fail("invalid-request", "Selected text span is invalid.");
  }
  const consolidation = page.pageContentConsolidation
    ? { pageContentConsolidation: page.pageContentConsolidation }
    : {};
  const object = enclosingTextObject(page.source, runSpan.start);
  if (runSpan.end > object.end) {
    if (input.keepTextRules?.styles.length) {
      fail("unsupported-document", MRAI_NATIVE_TEXT_RULES_NOT_CARRIED_MESSAGE);
    }
    return planTextObjectsDelete({ ...input, runSpan }, page, object);
  }
  // The rules drawn beside the object go in the same write (walk s2r-c:
  // two writes of their own before the text's, each proving the run again).
  const textRules = input.keepTextRules?.styles.length
    ? pdfNativeTextRulesStanding(page.streamBytes, object, input.keepTextRules.styles, input.keepTextRules.geometry)
    : [];
  const rulesCarried = textRules.length ? { textRules } : {};
  const scope: PdfNativeTextObjectDeleteScope =
    object.showCount === 1 ? "object" : "show";
  let span = { start: object.start, end: object.end };
  let replacementForm: string = mraiNativeEmptiedTextObject(page.source.slice(object.start, object.end));
  if (scope === "show") {
    const shows = showOperatorsIn(page.source, runSpan.start, runSpan.end);
    if (!shows || shows.length === 0) {
      fail(
        "precondition-mismatch",
        "Selected text span does not hold readable text-showing operators.",
      );
    }
    let penRestore: PdfNativeTextObjectDeletePlan["penRestore"] | undefined;
    // The pen number is read off the font's widths, so the objects the proof
    // read them from are pinned beside the page.
    let penPins: readonly PdfNativeObjectPrecondition[] = [];
    // The text after a show is looked for past the whole span, so the shows
    // that draw on from their end are the last of it.
    const needing = shows.filter(
      (show) => !textAfterShowIsPlacedAfresh(page.source, show.end, runSpan.end, object.end),
    ).length;
    if (needing > 0) {
      // The text after the span starts where the span ended. Deleting a
      // proven show leaves its advance behind, so that text stays put.
      const proven =
        shows.length === 1
          ? await provenPenRestore(input, page)
          : await provenShowPens(input, page, shows, needing);
      penRestore = proven?.pen;
      penPins = proven?.pins ?? [];
      if (!penRestore) {
        fail(
          "unsupported-document",
          "The text after this run continues from where it ends.",
        );
      }
    }
    span = { start: runSpan.start, end: runSpan.end };
    const emptied = mraiNativeEmptyShowsForm(
      page.source.slice(runSpan.start, runSpan.end),
      penRestoreNumbers(penRestore),
      penRestoreSpacings(penRestore),
    );
    if (emptied === null) {
      fail("unsupported-document", "The text after this run cannot keep its place under its character spacing.");
    }
    replacementForm = emptied;
    if (page.source.slice(span.start, span.end) === replacementForm) {
      fail("unsupported-document", "This text object is already empty.");
    }
    const objectBytes = page.streamBytes.slice(span.start, span.end);
    const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
      identity: page.identity,
      expectedObjectSha256: page.objectSha256,
      expectedStreamSha256: page.streamSha256,
      operatorSpan: Object.freeze({
        start: span.start,
        end: span.end,
        expectedSha256: await sha256PdfBytes(objectBytes),
      }),
      replacement: asciiBytes(replacementForm),
      ...consolidation,
      ...rulesCarried,
      untouchedObjects: Object.freeze([
        ...(page.untouchedObjects ?? []),
        ...penPins.filter(
          (pin) =>
            !(page.untouchedObjects ?? []).some(
              (item) =>
                item.objectNumber === pin.objectNumber && item.generation === pin.generation,
            ) &&
            !(
              pin.objectNumber === page.contentStream.objectNumber &&
              pin.generation === page.contentStream.generation
            ),
        ),
      ]),
      ...(page.formInstanceIsolation
        ? { formInstanceIsolation: page.formInstanceIsolation }
        : {}),
    });
    return Object.freeze({
      schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
      direction: "delete",
      scope,
      pageNumber: page.pageNumber,
      page: page.page,
      contentStream: page.contentStream,
      operatorSpan: request.operatorSpan,
      objectBytes: Object.freeze([...objectBytes]),
      showCount: object.showCount,
      ...(penRestore ? { penRestore } : {}),
      request,
    });
  }
  const objectBytes = page.streamBytes.slice(span.start, span.end);
  const replacement = asciiBytes(replacementForm);
  if (objectBytes.length <= replacement.length) {
    fail("unsupported-document", "This text object is already empty.");
  }
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: page.identity,
    expectedObjectSha256: page.objectSha256,
    expectedStreamSha256: page.streamSha256,
    operatorSpan: Object.freeze({
      start: span.start,
      end: span.end,
      expectedSha256: await sha256PdfBytes(objectBytes),
    }),
    replacement,
    ...consolidation,
    // A rule beside a run inside a shared form is taken away on its own.
    ...(page.formInstanceIsolation ? {} : rulesCarried),
    untouchedObjects: page.untouchedObjects,
    ...(page.formInstanceIsolation
      ? { formInstanceIsolation: page.formInstanceIsolation }
      : {}),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
    direction: "delete",
    scope,
    pageNumber: page.pageNumber,
    page: page.page,
    contentStream: page.contentStream,
    operatorSpan: request.operatorSpan,
    objectBytes: Object.freeze([...objectBytes]),
    showCount: object.showCount,
    request,
  });
}

/**
 * A run a reader reads as one item across several whole text objects -- the
 * shows the run holds are every show those objects hold -- is deleted object
 * by object; nothing between the objects changes, so nothing after them
 * moves. A run that starts or ends partway through an object's shows is
 * still refused.
 */
async function planTextObjectsDelete(
  input: Readonly<{
    sourceBytes: Uint8Array;
    runSpan: Readonly<{ start: number; end: number }>;
    textSelection?: PdfNativeListTextSelectionEvidence | null;
  }>,
  page: Awaited<ReturnType<typeof resolveTextPage>>,
  first: PdfNativeTextObjectSpan,
): Promise<PdfNativeTextObjectDeletePlan> {
  const last = enclosingTextObject(page.source, input.runSpan.end - 1);
  const leading = showOperatorsIn(page.source, first.start + 2, input.runSpan.start);
  const trailing = showOperatorsIn(page.source, input.runSpan.end, last.end - 2);
  if (!leading || !trailing) {
    fail("precondition-mismatch", "Selected text crosses its text object.");
  }
  if (leading.length > 0 || trailing.length > 0) {
    // The run starts or ends partway through an object: its own shows go,
    // and the objects' other shows stay.
    return planTextRunDelete(input, page, last);
  }
  const recorded = page.source.slice(first.start, last.end);
  const emptied = mraiNativeEmptiedTextObjectsForm(recorded);
  if (emptied === null) {
    fail("precondition-mismatch", "Selected text crosses its text object.");
  }
  const objectBytes = page.streamBytes.slice(first.start, last.end);
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: page.identity,
    expectedObjectSha256: page.objectSha256,
    expectedStreamSha256: page.streamSha256,
    operatorSpan: Object.freeze({
      start: first.start,
      end: last.end,
      expectedSha256: await sha256PdfBytes(objectBytes),
    }),
    replacement: asciiBytes(emptied),
    ...(page.pageContentConsolidation
      ? { pageContentConsolidation: page.pageContentConsolidation }
      : {}),
    untouchedObjects: page.untouchedObjects,
    ...(page.formInstanceIsolation
      ? { formInstanceIsolation: page.formInstanceIsolation }
      : {}),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
    direction: "delete",
    scope: "objects",
    pageNumber: page.pageNumber,
    page: page.page,
    contentStream: page.contentStream,
    operatorSpan: request.operatorSpan,
    objectBytes: Object.freeze([...objectBytes]),
    showCount: (recorded.match(SHOW_OPERATOR) ?? []).length,
    request,
  });
}

/**
 * A run that crosses text objects and starts or ends partway through them:
 * each of its shows is emptied where it is. A show whose object draws on
 * from its end keeps its advance, which only the run's last show can need
 * (every other one is followed by its object's end or by the run's next
 * show); its advance is the one the proof of the run measured for that show.
 */
async function planTextRunDelete(
  input: Readonly<{
    sourceBytes: Uint8Array;
    runSpan: Readonly<{ start: number; end: number }>;
    textSelection?: PdfNativeListTextSelectionEvidence | null;
  }>,
  page: Awaited<ReturnType<typeof resolveTextPage>>,
  last: PdfNativeTextObjectSpan,
): Promise<PdfNativeTextObjectDeletePlan> {
  const span = input.runSpan;
  const shows = showOperatorsIn(page.source, span.start, span.end);
  if (!shows || shows.length === 0) {
    fail("precondition-mismatch", "Selected text crosses its text object.");
  }
  const finalShow = shows[shows.length - 1];
  let pen: PdfNativeTextObjectDeletePlan["penRestore"] | undefined;
  let penPins: readonly PdfNativeObjectPrecondition[] = [];
  for (const show of shows) {
    const objectEnd = enclosingTextObject(page.source, show.start).end;
    if (textAfterShowIsPlacedAfresh(page.source, show.end, span.end, objectEnd)) continue;
    const alone = shows.filter((item) => item.start >= last.start).length === 1;
    const proven =
      show === finalShow && alone && !page.pageContentConsolidation
        ? await provenRunLastShowPen(input, page, finalShow)
        : undefined;
    if (!proven) {
      fail("unsupported-document", "The text after this run continues from where it ends.");
    }
    pen = proven.pen;
    penPins = proven.pins;
  }
  const objectBytes = page.streamBytes.slice(span.start, span.end);
  const form = mraiNativeDeletedRunForm(page.source.slice(span.start, span.end), pen?.number ?? null,
    pen?.characterSpacing ?? 0);
  if (form === null) {
    fail("precondition-mismatch", "Selected text crosses its text object.");
  }
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: page.identity,
    expectedObjectSha256: page.objectSha256,
    expectedStreamSha256: page.streamSha256,
    operatorSpan: Object.freeze({
      start: span.start,
      end: span.end,
      expectedSha256: await sha256PdfBytes(objectBytes),
    }),
    replacement: asciiBytes(form),
    ...(page.pageContentConsolidation
      ? { pageContentConsolidation: page.pageContentConsolidation }
      : {}),
    untouchedObjects: Object.freeze([
      ...(page.untouchedObjects ?? []),
      ...penPins.filter(
        (pin) =>
          !(page.untouchedObjects ?? []).some(
            (item) => item.objectNumber === pin.objectNumber && item.generation === pin.generation,
          ) &&
          !(
            pin.objectNumber === page.contentStream.objectNumber &&
            pin.generation === page.contentStream.generation
          ),
      ),
    ]),
    ...(page.formInstanceIsolation
      ? { formInstanceIsolation: page.formInstanceIsolation }
      : {}),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
    direction: "delete",
    scope: "run",
    pageNumber: page.pageNumber,
    page: page.page,
    contentStream: page.contentStream,
    operatorSpan: request.operatorSpan,
    objectBytes: Object.freeze([...objectBytes]),
    showCount: shows.length,
    ...(pen ? { penRestore: pen } : {}),
    request,
  });
}

/**
 * The advance the last show of a run crossing text objects leaves behind:
 * the proof of the run measures it as one of the runs it groups, in that
 * show's own size and scale.
 */
async function provenRunLastShowPen(
  input: Readonly<{
    sourceBytes: Uint8Array;
    textSelection?: PdfNativeListTextSelectionEvidence | null;
  }>,
  page: Awaited<ReturnType<typeof resolveTextPage>>,
  show: Readonly<{ start: number; end: number }>,
): Promise<
  | Readonly<{
      pen: NonNullable<PdfNativeTextObjectDeletePlan["penRestore"]>;
      pins: readonly PdfNativeObjectPrecondition[];
    }>
  | undefined
> {
  if (!input.textSelection) return undefined;
  let proof: Awaited<ReturnType<typeof proveMraiNativeListTextSelectionForWriter>>;
  try {
    proof = await proveMraiNativeListTextSelectionForWriter({
      sourceBytes: input.sourceBytes,
      selection: {
        ...input.textSelection,
        sourceSha256: await sha256PdfBytes(input.sourceBytes),
      },
    });
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return undefined;
    throw error;
  }
  const stream = proof.sourceRun?.contentStream ?? proof.target;
  const leaf = proof.evidence.positionedRunGroup?.runs.at(-1)?.evidence;
  const scale = leaf?.sameLineLayoutState?.horizontalScale;
  if (
    !leaf ||
    leaf.textOperatorSpan.start !== show.start ||
    proofShowEnd(leaf) !== show.end ||
    stream.objectNumber !== page.contentStream.objectNumber ||
    stream.generation !== page.contentStream.generation ||
    leaf.positionedRunGroup ||
    leaf.linePlan ||
    scale === undefined ||
    !(scale > 0) ||
    !(leaf.fontSize > 0) ||
    !(leaf.originalAdvance > 0)
  ) {
    return undefined;
  }
  const edges = pdfNativeTextShowEdgeAdjustments(page.source.slice(show.start, show.end));
  if (edges === null) return undefined;
  return Object.freeze({
    pen: Object.freeze({
      number: formatPenNumber(
        (-leaf.originalAdvance * 1_000) / (leaf.fontSize * (scale / 100)) + edges,
      ),
      sourceFidelity: await penSourceFidelity(leaf, page.source),
      ...penSpacing(leaf),
    }),
    pins: proof.request.untouchedObjects ?? [],
  });
}

/** A run's proof as a pen number's authority: what the run is and how it is
 * set, over the run's own show, without the growth proofs (free space, a line
 * shift and the operands it would move) a writer that widened it would need;
 * the number is the run's own advance. */
async function penSourceFidelity(
  evidence: PdfNativeListTextFidelityEvidence,
  source: string,
): Promise<PdfNativeListTextFidelityEvidence> {
  const { lineShift, sameLineCapacity: _sameLineCapacity, ...own } = evidence;
  const span = evidence.textOperatorSpan;
  if (!lineShift || !(span.start < lineShift.showEnd && lineShift.showEnd < span.end)) {
    const { lineShift: _dropped, ...kept } = own as PdfNativeListTextFidelityEvidence;
    return Object.freeze(kept) as PdfNativeListTextFidelityEvidence;
  }
  return Object.freeze({
    ...own,
    textOperatorSpan: Object.freeze({
      start: span.start,
      end: lineShift.showEnd,
      sha256: await sha256PdfBytes(asciiBytes(source.slice(span.start, lineShift.showEnd))),
    }),
  }) as PdfNativeListTextFidelityEvidence;
}

/** Where a proof's run ends: its show's end, before the operands a line
 * shift would move. */
function proofShowEnd(evidence: PdfNativeListTextFidelityEvidence): number {
  const span = evidence.textOperatorSpan;
  const showEnd = evidence.lineShift?.showEnd;
  return showEnd !== undefined && span.start < showEnd && showEnd < span.end ? showEnd : span.end;
}

/**
 * The advances a span the page draws as several shows leaves behind.
 *
 * A reader hands over as one item what a typesetter may have written as more
 * than one show -- pdfTeX starts a new one wherever the font changes, which
 * it does for a digit or a symbol inside a word -- and the proof of such a
 * selection is a group with one run per show. Each show is emptied where it
 * is and left as its own advance, measured in its own font: no advance has
 * to be carried across a font change, and the pen ends where it ended.
 *
 * Only the shows whose text is not placed afresh need one, and those are the
 * last of the span (a show that draws on from the one before it draws on
 * from every show before that too). The last show always needs it when any
 * does, because they all look past the span for the text that follows.
 */
async function provenShowPens(
  input: Readonly<{
    sourceBytes: Uint8Array;
    runSpan: Readonly<{ start: number; end: number }>;
    textSelection?: PdfNativeListTextSelectionEvidence | null;
  }>,
  page: Awaited<ReturnType<typeof resolveTextPage>>,
  shows: readonly Readonly<{ start: number; end: number }>[],
  needing: number,
): Promise<
  | Readonly<{
      pen: NonNullable<PdfNativeTextObjectDeletePlan["penRestore"]>;
      pins: readonly PdfNativeObjectPrecondition[];
    }>
  | undefined
> {
  if (!input.textSelection || needing < 1 || needing > shows.length) return undefined;
  let proof: Awaited<ReturnType<typeof proveMraiNativeListTextSelectionForWriter>>;
  try {
    proof = await proveMraiNativeListTextSelectionForWriter({
      sourceBytes: input.sourceBytes,
      selection: {
        ...input.textSelection,
        sourceSha256: await sha256PdfBytes(input.sourceBytes),
      },
    });
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return undefined;
    throw error;
  }
  const stream = proof.sourceRun?.contentStream ?? proof.target;
  const sourceStream = page.pageContentConsolidation
    ? (input as Readonly<{ contentStream?: PdfNativeTextObjectStreamRef | null }>).contentStream
    : page.contentStream;
  // A positioned group's runs, or a positioned paragraph's lines (a line
  // whose parts the writer set in fonts of their own: each part a show of
  
  // the delete of a line typed with part formats was refused).
  // A line drawn by several shows is proved as a group of its own: its runs
  // stand for its shows.
  const runs: readonly Readonly<{ evidence: PdfNativeListTextFidelityEvidence }>[] | undefined =
    proof.evidence.positionedRunGroup?.runs ?? proof.evidence.positionedLines?.lines
      .flatMap((line) => line.evidence.positionedRunGroup?.runs ?? [line]);
  if (
    !runs ||
    !sourceStream ||
    stream.objectNumber !== sourceStream.objectNumber ||
    stream.generation !== sourceStream.generation ||
    proof.evidence.linePlan
  ) {
    return undefined;
  }
  if (runs.length !== shows.length) return undefined;
  // Each run is the proof of the show it names, and the shows it names are
  // the ones being emptied: a group that covers anything else says nothing
  // about what this delete removes.
  const pens: { number: string; sourceFidelity: PdfNativeListTextFidelityEvidence; characterSpacing?: number }[] = [];
  for (const [index, show] of shows.entries()) {
    const leaf = runs[index].evidence;
    if (leaf.textOperatorSpan.start !== show.start || proofShowEnd(leaf) !== show.end) {
      return undefined;
    }
    if (index < shows.length - needing) continue;
    const scale = leaf.sameLineLayoutState?.horizontalScale;
    if (
      leaf.positionedRunGroup ||
      leaf.linePlan ||
      scale === undefined ||
      !(scale > 0) ||
      !(leaf.fontSize > 0) ||
      !(leaf.originalAdvance > 0)
    ) {
      return undefined;
    }
    const edges = pdfNativeTextShowEdgeAdjustments(page.source.slice(show.start, show.end));
    if (edges === null) return undefined;
    pens.push({
      number: formatPenNumber(
        (-leaf.originalAdvance * 1_000) / (leaf.fontSize * (scale / 100)) + edges,
      ),
      sourceFidelity: await penSourceFidelity(leaf, page.source),
      ...penSpacing(leaf),
    });
  }
  const last = pens.pop();
  if (!last) return undefined;
  return Object.freeze({
    pen: Object.freeze({
      ...last,
      ...(pens.length > 0 ? { precedingShows: Object.freeze(pens.map((pen) => Object.freeze(pen))) } : {}),
    }),
    pins: proof.request.untouchedObjects ?? [],
  });
}

/**
 * The advance a deleted show leaves behind, from the proof of the run: the
 * same run, in the same stream, measured by the text route in its own size
 * and scale, as the TJ number that moves the pen that far.
 */
async function provenPenRestore(
  input: Readonly<{
    sourceBytes: Uint8Array;
    runSpan: Readonly<{ start: number; end: number }>;
    textSelection?: PdfNativeListTextSelectionEvidence | null;
  }>,
  page: Awaited<ReturnType<typeof resolveTextPage>>,
): Promise<
  | Readonly<{
      pen: NonNullable<PdfNativeTextObjectDeletePlan["penRestore"]>;
      pins: readonly PdfNativeObjectPrecondition[];
    }>
  | undefined
> {
  if (!input.textSelection) return undefined;
  let proof: Awaited<ReturnType<typeof proveMraiNativeListTextSelectionForWriter>>;
  try {
    // The selection is re-proven against the bytes being edited now (a redo
    // runs after an undo rewrote them); the proof checks everything again.
    proof = await proveMraiNativeListTextSelectionForWriter({
      sourceBytes: input.sourceBytes,
      selection: {
        ...input.textSelection,
        sourceSha256: await sha256PdfBytes(input.sourceBytes),
      },
    });
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return undefined;
    throw error;
  }
  const span = pdfNativeListTextSourceRunSpan(proof);
  const stream = proof.sourceRun?.contentStream ?? proof.target;
  const evidence = proof.evidence;
  const scale = evidence.sameLineLayoutState?.horizontalScale;
  // The proof names the run in its source stream, and the operator where it
  // is written: the stream itself, or the page's streams joined as one when
  // this delete rewrites them that way. Replay proves it where it is written.
  const sourceStream = page.pageContentConsolidation
    ? (input as Readonly<{ contentStream?: PdfNativeTextObjectStreamRef | null }>).contentStream
    : page.contentStream;
  if (
    !span ||
    !sourceStream ||
    span.start !== input.runSpan.start ||
    span.end !== input.runSpan.end ||
    evidence.textOperatorSpan.start !== page.runSpan.start ||
    proofShowEnd(evidence) !== page.runSpan.end ||
    stream.objectNumber !== sourceStream.objectNumber ||
    stream.generation !== sourceStream.generation ||
    evidence.positionedRunGroup ||
    evidence.linePlan ||
    scale === undefined ||
    !(scale > 0) ||
    !(evidence.fontSize > 0) ||
    !(evidence.originalAdvance > 0)
  ) {
    return undefined;
  }
  // The numbers before the first glyph and after the last move the pen too,
  // though the run's measured advance leaves them out.
  const edges = pdfNativeTextShowEdgeAdjustments(page.source.slice(span.start, span.end));
  if (edges === null) return undefined;
  return Object.freeze({
    pen: Object.freeze({
      number: formatPenNumber(
        (-evidence.originalAdvance * 1_000) / (evidence.fontSize * (scale / 100)) + edges,
      ),
      sourceFidelity: await penSourceFidelity(evidence, page.source),
      ...penSpacing(evidence),
    }),
    pins: proof.request.untouchedObjects ?? [],
  });
}

/** Put a deleted text object back over the empty one the delete wrote. */
export async function planMraiNativeTextObjectDeleteRestore(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly offset: number;
  readonly objectBytes: readonly number[];
  /** Which of the page's content streams that offset is an offset into. */
  readonly contentStream?: PdfNativeTextObjectStreamRef | null;
  /** What the delete emptied; records before shows were deletable say the object. */
  readonly scope?: PdfNativeTextObjectDeleteScope;
  /** The pen number a show deleted keeping its advance was left as; a
   * record written before a span of several shows could keep theirs. */
  readonly penRestore?: string | null;
  /** The advances the last shows of the span were left as, in the order the
   * span holds them. */
  readonly penRestores?: readonly string[] | null;
  /** The character spacing each of those advances is drawn under, in the
   * same order (zero where it is none); absent on records from before the
   * pen was written where the reader reads it. */
  readonly penCharacterSpacings?: readonly number[] | null;
}): Promise<PdfNativeTextObjectDeletePlan> {
  const page = await resolveTextPage(
    input.sourceBytes,
    input.pageNumber,
    input.contentStream ?? null,
  );
  const scope: PdfNativeTextObjectDeleteScope = input.scope ?? "object";
  const pens = input.penRestores ?? (input.penRestore ? [input.penRestore] : []);
  const spacings = input.penCharacterSpacings ?? [];
  if (
    (pens.length > 0 &&
      ((scope !== "show" && scope !== "run") ||
        (scope === "run" && pens.length !== 1) ||
        pens.some((pen) => !/^-?\d+(?:\.\d+)?$/.test(pen)))) ||
    (spacings.length > 0 && (spacings.length !== pens.length || spacings.some((value) => !Number.isFinite(value))))
  ) {
    fail("invalid-request", "The recorded pen restore is not a show's advance.");
  }
  
  const empty = Number.isSafeInteger(input.offset) && input.offset >= 0
    ? deletedFormsOf(scope, input.objectBytes, pens, spacings)
      .find((form) => page.source.slice(input.offset, input.offset + form.length) === form) ??
      deletedFormOf(scope, input.objectBytes, pens, spacings)
    : deletedFormOf(scope, input.objectBytes, pens, spacings);
  const end = input.offset + empty.length;
  if (
    !Number.isSafeInteger(input.offset) ||
    input.offset < 0 ||
    end > page.streamBytes.length ||
    page.source.slice(input.offset, end) !== empty
  ) {
    fail(
      "precondition-mismatch",
      "The deleted text object is no longer where it was removed.",
    );
  }
  const objectBytes = Uint8Array.from(input.objectBytes);
  const restored = LATIN1.decode(objectBytes);
  if (
    // A kept advance may be written longer than the show it replaces (one
    // under a character spacing sets the spacing around it), so a span of
    // shows is held to its own empty form, as replay's record reader holds it.
    (scope === "run"
      ? restored === empty
      : scope === "show"
        ? objectBytes.length <= (mraiNativeEmptyShowsForm(restored) ?? restored).length
        : objectBytes.length <= empty.length) ||
    (scope !== "show" && scope !== "run" && (!restored.startsWith("BT") || !restored.endsWith("ET")))
  ) {
    fail("invalid-request", "The recorded text object is not one BT…ET block.");
  }
  if (scope === "show") {
    const object = enclosingTextObject(page.source, input.offset);
    if (input.offset <= object.start || end > object.end - 2) {
      fail("precondition-mismatch", "The emptied shows are no longer inside their text object.");
    }
  }
  // The emptied show must still sit inside a text object; the object's show
  // count is unchanged because the empty form is still a show operator.
  const showCount =
    scope === "show"
      ? enclosingTextObject(page.source, input.offset).showCount
      : scope === "run"
        ? (showOperatorsIn(restored, 0, restored.length) ?? []).length
        : (restored.match(SHOW_OPERATOR) ?? []).length;
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: page.identity,
    expectedObjectSha256: page.objectSha256,
    expectedStreamSha256: page.streamSha256,
    operatorSpan: Object.freeze({
      start: input.offset,
      end,
      expectedSha256: await sha256PdfBytes(asciiBytes(empty)),
    }),
    replacement: objectBytes,
    untouchedObjects: page.untouchedObjects,
    ...(page.formInstanceIsolation
      ? { formInstanceIsolation: page.formInstanceIsolation }
      : {}),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
    direction: "restore",
    scope,
    pageNumber: page.pageNumber,
    page: page.page,
    contentStream: page.contentStream,
    operatorSpan: request.operatorSpan,
    objectBytes: Object.freeze([...objectBytes]),
    showCount,
    request,
  });
}

export async function prepareMraiNativeTextObjectDelete(
  sourceBytes: Uint8Array,
  plan: PdfNativeTextObjectDeletePlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (plan.schemaVersion !== MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION) {
    fail("invalid-request", "Text object delete plan authority is invalid.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeTextObjectDeleteCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeTextObjectDeletePlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<Readonly<{ verified: true; candidateSha256: string }>> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  // A run inside a Form the document draws in more than one place is written
  // on a clone: in the candidate this page draws the copy, and the form the
  // plan named draws it no longer. The change is read back where it was
  // written -- and at the span this plan wrote, because a producer may split
  // one text object across two content streams anywhere on the page: the
  // drawing a run is rewritten in is the stream's own when the run sits
  // inside a closed text object there, and the page's streams joined as one
  // when it does not. Asking that question at offset zero asks it of a split
  // the run has nothing to do with, and the page then reopens as a joined
  // stream whose offsets are not the ones this plan wrote.
  const reopened = await resolveTextPage(
    prepared.candidateBytes,
    plan.pageNumber,
    prepared.formInstanceIsolation?.primaryClone ?? plan.contentStream,
    Object.freeze({ start: plan.operatorSpan.start, end: plan.operatorSpan.end }),
  );
  const written =
    plan.direction === "delete"
      ? deletedFormOf(plan.scope, plan.objectBytes, penRestoreNumbers(plan.penRestore),
        penRestoreSpacings(plan.penRestore))
      : LATIN1.decode(Uint8Array.from(plan.objectBytes));
  // A rule taken away before the object moves it that much to the left.
  const writtenStart = plan.operatorSpan.start - (plan.request.textRules ?? [])
    .filter((rule) => rule.end <= plan.operatorSpan.start)
    .reduce((sum, rule) => sum + rule.removed.length - rule.added.length, 0);
  if (
    reopened.source.slice(
      writtenStart,
      writtenStart + written.length,
    ) !== written
  ) {
    prepared.candidateBytes.fill(0);
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF text object does not match the requested change.",
    );
  }
  return Object.freeze({ verified: true, candidateSha256: prepared.candidateSha256 });
}
