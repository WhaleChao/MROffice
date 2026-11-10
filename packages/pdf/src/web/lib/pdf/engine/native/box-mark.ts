/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { LATIN1_EXACT } from "./latin1";



































/** The tag of the sequence (native-text-hard-groups.ts reads it). */
export const MRAI_NATIVE_BOX_MARK_TAG = "MRAIBox";
const OPEN_TEXT = `/${MRAI_NATIVE_BOX_MARK_TAG} BMC\n`;
const CLOSE_TEXT = "\nEMC\n";

/** The bytes written where the box starts and where it ends. */
export const MRAI_NATIVE_BOX_MARK_OPEN: Uint8Array = Uint8Array.from(OPEN_TEXT, (character) => character.charCodeAt(0));
export const MRAI_NATIVE_BOX_MARK_CLOSE: Uint8Array = Uint8Array.from(CLOSE_TEXT, (character) => character.charCodeAt(0));

/** Where a box's sequence opens and closes, in the source stream. */
export interface PdfNativeBoxMark {
  readonly start: number;
  readonly end: number;
}

type Span = Readonly<{ start: number; end: number }>;

/** One operation: its first operand's start, its operator's end. */
interface BoxMarkOperation {
  readonly start: number;
  readonly end: number;
  readonly operator: string;
  /** The first operand's bytes (a marked-content tag). */
  readonly firstOperand?: string;
}

const WHITESPACE = new Set(["\u0000", "\t", "\n", "\f", "\r", " "]);
const DELIMITERS = new Set(["(", ")", "<", ">", "[", "]", "{", "}", "/", "%"]);
const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)$/u;
const SHOWS = new Set(["Tj", "TJ", "'", '"']);
const OPENS_BOX = new Set(["BT", "q", "BMC", "BDC"]);
const CLOSES_BOX = new Set(["ET", "Q", "EMC"]);

function literalStringEnd(source: string, start: number): number | null {
  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    const character = source[index];
    if (character === "\\") {
      index += 1;
    } else if (character === "(") {
      depth += 1;
    } else if (character === ")") {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
  }
  return null;
}

/**
 * The operations of `source[from, to)`, each from its first operand to its
 * operator, or null when the bytes are not plain operations: a comment, an
 * inline image, an unclosed string or dictionary, a stray delimiter, or
 * operands no operator takes (mirror of native_box_mark.py
 * box_mark_operations).
 */
export function mraiNativeBoxMarkOperations(
  source: string,
  from: number,
  to: number,
): readonly BoxMarkOperation[] | null {
  const operations: BoxMarkOperation[] = [];
  let first: number | null = null;
  let firstOperand: string | undefined;
  let index = from;
  const operand = (start: number, end: number) => {
    if (first === null) {
      first = start;
      firstOperand = source.slice(start, end);
    }
  };
  while (index < to) {
    const character = source[index];
    if (WHITESPACE.has(character)) {
      index += 1;
      continue;
    }
    if (character === "%" || character === ")" || character === ">") return null;
    let end: number;
    if (character === "(") {
      const close = literalStringEnd(source, index);
      if (close === null) return null;
      end = close;
    } else if (character === "<" && source[index + 1] === "<") {
      let depth = 0;
      let cursor = index;
      end = -1;
      while (cursor < to) {
        if (source[cursor] === "<" && source[cursor + 1] === "<") {
          depth += 1;
          cursor += 2;
        } else if (source[cursor] === ">" && source[cursor + 1] === ">") {
          depth -= 1;
          cursor += 2;
          if (depth === 0) {
            end = cursor;
            break;
          }
        } else if (source[cursor] === "(") {
          const close = literalStringEnd(source, cursor);
          if (close === null) return null;
          cursor = close;
        } else if (source[cursor] === "<") {
          const close = source.indexOf(">", cursor + 1);
          if (close < 0) return null;
          cursor = close + 1;
        } else {
          cursor += 1;
        }
      }
      if (end < 0) return null;
    } else if (character === "<") {
      const close = source.indexOf(">", index + 1);
      if (close < 0) return null;
      end = close + 1;
    } else if (character === "[" || character === "]" || character === "{" || character === "}") {
      end = index + 1;
    } else {
      end = index + 1;
      while (end < to && !WHITESPACE.has(source[end]) && !DELIMITERS.has(source[end])) end += 1;
      const word = source.slice(index, end);
      if (character !== "/" && !NUMBER.test(word) && word !== "true" && word !== "false" && word !== "null") {
        if (word === "BI" || word === "ID" || word === "EI") return null;
        operations.push(Object.freeze({
          start: first ?? index,
          end,
          operator: word,
          ...(firstOperand !== undefined ? { firstOperand } : {}),
        }));
        first = null;
        firstOperand = undefined;
        index = end;
        continue;
      }
    }
    if (end > to) return null;
    operand(index, end);
    index = end;
  }
  return first === null ? operations : null;
}

/**
 * Why a box mark cannot stand where it says, or null when it can (mirror of
 * native_box_mark.py box_mark_issue). `changed` is the stream the write
 * leaves with the mark in it; the edit's operator span lies inside the mark,
 * and the edit changes the stream's length by `delta` there.
 */
export function mraiNativeBoxMarkIssue(
  source: string,
  changed: string,
  mark: PdfNativeBoxMark,
  operatorSpan: Span,
  delta: number,
): string | null {
  if (
    !Number.isSafeInteger(mark.start) || !Number.isSafeInteger(mark.end) ||
    mark.start < 0 || mark.end <= mark.start ||
    mark.start > operatorSpan.start || operatorSpan.end > mark.end || mark.end > source.length
  ) {
    return "The box mark does not hold the edit.";
  }
  const before = mark.start > 0 ? source[mark.start - 1] : " ";
  const after = mark.end < source.length ? source[mark.end] : " ";
  if (!(WHITESPACE.has(before) || ")>]}".includes(before)) || !(WHITESPACE.has(after) || DELIMITERS.has(after))) {
    return "The box mark does not stand between two operations.";
  }
  const innerStart = mark.start + OPEN_TEXT.length;
  const innerEnd = mark.end + delta + OPEN_TEXT.length;
  if (
    changed.slice(mark.start, innerStart) !== OPEN_TEXT ||
    changed.slice(innerEnd, innerEnd + CLOSE_TEXT.length) !== CLOSE_TEXT
  ) {
    return "The box mark is not written where it says.";
  }
  const operations = mraiNativeBoxMarkOperations(changed, innerStart, innerEnd);
  if (
    !operations?.length ||
    !OPENS_BOX.has(operations[0].operator) ||
    !CLOSES_BOX.has(operations[operations.length - 1].operator)
  ) {
    return "The box mark does not hold whole operations.";
  }
  let text = 0;
  let saved = 0;
  let marked = 0;
  for (const { operator } of operations) {
    if (operator === "BT") text += 1;
    else if (operator === "ET") text -= 1;
    else if (operator === "q") saved += 1;
    else if (operator === "Q") saved -= 1;
    else if (operator === "BMC" || operator === "BDC") marked += 1;
    else if (operator === "EMC") marked -= 1;
    if (text < 0 || text > 1 || saved < 0 || marked < 0) return "The box mark crosses a nesting it is in.";
  }
  return text === 0 && saved === 0 && marked === 0 ? null : "The box mark crosses a nesting it is in.";
}

/** The stream a write leaves: its operator span replaced and, with a box
 * mark, the mark's two halves written round the box (mirror of the replay's
 * edited spans). */
export function applyMraiNativeBoxMarkedEdit(
  stream: Uint8Array,
  operatorSpan: Span,
  replacement: Uint8Array,
  mark: PdfNativeBoxMark | undefined,
): Uint8Array {
  const parts = mark
    ? [
        stream.subarray(0, mark.start),
        MRAI_NATIVE_BOX_MARK_OPEN,
        stream.subarray(mark.start, operatorSpan.start),
        replacement,
        stream.subarray(operatorSpan.end, mark.end),
        MRAI_NATIVE_BOX_MARK_CLOSE,
        stream.subarray(mark.end),
      ]
    : [stream.subarray(0, operatorSpan.start), replacement, stream.subarray(operatorSpan.end)];
  const out = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

const LATIN1 = LATIN1_EXACT;

/**
 * Where to mark the box a write sets, or null when it is not wrapped alone.
 *
 * The box is the text objects that draw `boxShows` (each show from its first
 * operand to its operator, in the source stream) and the edit's own span;
 * the range is widened to whole saves and whole marked-content sequences
 * that open outside it and close inside it, or the other way round (Word
 * draws each run inside `q ... Q` and `/Span <<...>> BDC ... EMC`). Every
 * show the range then holds must be one of the box's. A range the page
 * already marks as a box is not marked again.
 */
export function planMraiNativeBoxMark(input: {
  readonly streamBytes: Uint8Array;
  readonly operatorSpan: Span;
  readonly replacement: Uint8Array;
  readonly boxShows: readonly Span[];
}): PdfNativeBoxMark | null {
  const { streamBytes, operatorSpan, replacement } = input;
  if (input.boxShows.length === 0 || operatorSpan.end <= operatorSpan.start) return null;
  const source = LATIN1.decode(streamBytes);
  const operations = mraiNativeBoxMarkOperations(source, 0, source.length);
  if (!operations) return null;
  // Each opener's closer and each closer's opener, by operation index, and
  // the text object each operation stands in; a nesting the stream leaves
  // unbalanced is not wrapped.
  const partner = new Map<number, number>();
  const textObject: (number | null)[] = [];
  const stacks = { text: [] as number[], saved: [] as number[], marked: [] as number[] };
  for (const [index, { operator }] of operations.entries()) {
    const kind = operator === "BT" || operator === "ET" ? "text"
      : operator === "q" || operator === "Q" ? "saved"
      : operator === "BMC" || operator === "BDC" || operator === "EMC" ? "marked" : null;
    if (operator === "BT") {
      if (stacks.text.length > 0) return null;
      stacks.text.push(index);
    }
    textObject.push(stacks.text[0] ?? null);
    if (!kind || operator === "BT") continue;
    if (operator === "q" || operator === "BMC" || operator === "BDC") {
      stacks[kind].push(index);
      continue;
    }
    const opener = stacks[kind].pop();
    if (opener === undefined) return null;
    partner.set(opener, index);
    partner.set(index, opener);
  }
  if (stacks.text.length > 0) return null;
  const operationAt = new Map(operations.map((operation, index) => [operation.start, index]));
  // The edit's span, and the text objects the box's shows stand in.
  let first = -1;
  let last = -1;
  for (const [index, operation] of operations.entries()) {
    if (operation.end > operatorSpan.start && operation.start < operatorSpan.end) {
      if (first < 0) first = index;
      last = index;
    }
  }
  if (first < 0) return null;
  for (const show of input.boxShows) {
    const at = operationAt.get(show.start);
    const open = at === undefined ? null : textObject[at];
    const close = open === null ? undefined : partner.get(open);
    if (at === undefined || operations[at].end !== show.end || open === null || close === undefined) return null;
    first = Math.min(first, open);
    last = Math.max(last, close);
  }
  // Widen to every nesting that opens outside the range and closes in it,
  // or opens in it and closes outside.
  for (let widened = true; widened;) {
    widened = false;
    for (let index = first; index <= last; index += 1) {
      const other = partner.get(index);
      if (other === undefined) {
        if (OPENS_BOX.has(operations[index].operator) || CLOSES_BOX.has(operations[index].operator)) return null;
        continue;
      }
      if (other < first) {
        first = other;
        widened = true;
      } else if (other > last) {
        last = other;
        widened = true;
      }
    }
  }
  // The range holds the box's text and no other.
  const owned = new Set(input.boxShows.map((show) => `${show.start}:${show.end}`));
  for (let index = first; index <= last; index += 1) {
    const operation = operations[index];
    if (SHOWS.has(operation.operator) && !owned.has(`${operation.start}:${operation.end}`)) return null;
  }
  // Already one box: the sequence just round the range is a box mark.
  const opener = operations[first - 1];
  if (
    opener?.operator === "BMC" && opener.firstOperand === `/${MRAI_NATIVE_BOX_MARK_TAG}` &&
    partner.get(first - 1) === last + 1
  ) {
    return null;
  }
  const mark = Object.freeze({
    start: Math.min(operations[first].start, operatorSpan.start),
    end: Math.max(operations[last].end, operatorSpan.end),
  });
  const changed = LATIN1.decode(applyMraiNativeBoxMarkedEdit(streamBytes, operatorSpan, replacement, mark));
  const delta = replacement.length - (operatorSpan.end - operatorSpan.start);
  return mraiNativeBoxMarkIssue(source, changed, mark, operatorSpan, delta) === null ? mark : null;
}
