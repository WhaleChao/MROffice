/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The bytes of one underline or strikethrough: a filled quadrilateral in the
 * space the run's text matrix establishes, drawn beside the run's own
 * `BT … ET` object (text-decoration.ts places it; a property edit that
 * moves the run rewrites it in the same write, list-text-fidelity.ts).
 */
import { PDF_NATIVE_RULE_THICKNESS_EM, PDF_NATIVE_STRIKETHROUGH_OFFSET_EM, PDF_NATIVE_UNDERLINE_OFFSET_EM } from "../../native-text-rule-range";
import { PdfNativeTransactionError } from "./transaction-error";
import { MRAI_NATIVE_TEXT_RULES_NOT_CARRIED_MESSAGE } from "./vocabulary";
export { MRAI_NATIVE_TEXT_RULES_NOT_CARRIED_MESSAGE };
import {
  PDF_NATIVE_TEXT_RULE_RANGE_EPSILON,
  pdfNativeTextRuleRangeIsWhole,
  type PdfNativeTextRuleRange,
} from "../../native-text-rule-range";

/**
 * The `BT` and `ET` operators of `source`, found the way a reader lexes a
 * content stream: strings, hex strings, names, comments and inline image
 * data are skipped whole, so letters that spell an operator inside them (a
 * word such as "COMPLETED" drawn one kerned piece at a time, `(ET)`) are no
 * operator. Mirror of the replay's _text_object_operators.
 */
export function mraiNativeTextObjectOperators(
  source: string,
): readonly Readonly<{ index: number; operator: "BT" | "ET" }>[] {
  // Asked for every show of a paragraph (enclosingTextObject), each of which
  // read the whole stream again: 1 s for a UN report page's long figure
  
  const known = textObjectOperatorReadings.find((entry) => entry.source === source);
  if (known) return known.operators;
  const operators = readMraiNativeTextObjectOperators(source);
  textObjectOperatorReadings.unshift({ source, operators });
  textObjectOperatorReadings.length = Math.min(textObjectOperatorReadings.length, 4);
  return operators;
}

const textObjectOperatorReadings: { source: string; operators: readonly Readonly<{ index: number; operator: "BT" | "ET" }>[] }[] = [];

function readMraiNativeTextObjectOperators(
  source: string,
): readonly Readonly<{ index: number; operator: "BT" | "ET" }>[] {
  const operators: { index: number; operator: "BT" | "ET" }[] = [];
  const length = source.length;
  const white = (character: string | undefined) =>
    character === " " || character === "\n" || character === "\r" || character === "\t" ||
    character === "\f" || character === "\0";
  const delimiter = (character: string | undefined) =>
    character !== undefined && "()<>[]{}/%".includes(character);
  let index = 0;
  while (index < length) {
    const character = source[index];
    if (white(character)) {
      index += 1;
    } else if (character === "%") {
      while (index < length && source[index] !== "\n" && source[index] !== "\r") index += 1;
    } else if (character === "(") {
      let depth = 1;
      index += 1;
      while (index < length && depth > 0) {
        const inner = source[index];
        if (inner === "\\") {
          index += 2;
          continue;
        }
        if (inner === "(") depth += 1;
        else if (inner === ")") depth -= 1;
        index += 1;
      }
    } else if (character === "<") {
      if (source[index + 1] === "<") {
        index += 2;
      } else {
        const end = source.indexOf(">", index + 1);
        index = end < 0 ? length : end + 1;
      }
    } else if (character === "/") {
      index += 1;
      while (index < length && !white(source[index]) && !delimiter(source[index])) index += 1;
    } else if (delimiter(character)) {
      index += 1;
    } else {
      const start = index;
      while (index < length && !white(source[index]) && !delimiter(source[index])) index += 1;
      const token = source.slice(start, index);
      if (token === "BT" || token === "ET") {
        operators.push({ index: start, operator: token });
      } else if (token === "ID") {
        // Inline image data: one white-space byte, then any bytes up to a
        // white-space byte, `EI` and the end of that token.
        index += 1;
        while (index < length) {
          if (white(source[index]) && source[index + 1] === "E" && source[index + 2] === "I" &&
            (index + 3 >= length || white(source[index + 3]) || delimiter(source[index + 3]))) {
            index += 3;
            break;
          }
          index += 1;
        }
      }
    }
  }
  return Object.freeze(operators);
}

/**
 * The `BT … ET` object holding one byte offset of `source`: its start (the
 * `BT`) and its end (just after the `ET`).
 */
export function pdfNativeTextObjectEdgesAt(
  source: string,
  offset: number,
): Readonly<{ start: number; end: number }> {
  let open: number | null = null;
  for (const { index, operator } of mraiNativeTextObjectOperators(source)) {
    if (operator === "BT") {
      if (open !== null) throw new PdfNativeTransactionError("unsupported-document", "Nested PDF text objects are unsupported.");
      open = index;
      continue;
    }
    if (open === null) throw new PdfNativeTransactionError("unsupported-document", "PDF text object nesting is unbalanced.");
    const end = index + 2;
    if (offset >= open && offset < end) return Object.freeze({ start: open, end });
    open = null;
  }
  throw new PdfNativeTransactionError("precondition-mismatch", "Selected text is not inside a text object.");
}

export type PdfNativeTextRuleStyle = "underline" | "strikethrough";

export interface PdfNativeTextRuleGeometry {
  /** Text matrix of the run, which also carries its rotation and scale. */
  readonly textMatrix: readonly [number, number, number, number, number, number];
  readonly fontSize: number;
  /** The run's `Ts`: a raised run's rules stand with its glyphs. */
  readonly rise?: number;
  /** Advance of the run in the same units the text matrix maps. */
  readonly advance: number;
  readonly color: Readonly<{
    kind: "gray" | "rgb" | "cmyk";
    components: readonly number[];
  }>;
}

// Where a rule sits and how thick it is, as a fraction of the run's font
// size: the long-standing typographic proportions.
const UNDERLINE_OFFSET = PDF_NATIVE_UNDERLINE_OFFSET_EM;
const STRIKETHROUGH_OFFSET = PDF_NATIVE_STRIKETHROUGH_OFFSET_EM;
const RULE_THICKNESS = PDF_NATIVE_RULE_THICKNESS_EM;

function fail(message: string): never {
  throw new PdfNativeTransactionError("invalid-request", message);
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (character) => character.charCodeAt(0) & 0xff);
}

function pdfNumber(value: number): string {
  if (!Number.isFinite(value) || Math.abs(value) > 1_000_000) {
    fail("Native text decoration number is invalid.");
  }
  const normalized = Math.abs(value) < 0.0000005 ? 0 : value;
  return normalized
    .toFixed(6)
    .replace(/\.0+$/, "")
    .replace(/(\.\d*?)0+$/, "$1");
}

function colorOperator(color: PdfNativeTextRuleGeometry["color"]): string {
  const expected = color.kind === "gray" ? 1 : color.kind === "rgb" ? 3 : 4;
  if (
    color.components.length !== expected ||
    color.components.some((value) => !Number.isFinite(value) || value < 0 || value > 1)
  ) {
    fail("Native text decoration colour is invalid.");
  }
  const components = color.components.map(pdfNumber).join(" ");
  return `${components} ${color.kind === "gray" ? "g" : color.kind === "rgb" ? "rg" : "k"}`;
}

/** The rule's four corners, mapped through the run's own text matrix. */
export function pdfNativeTextRuleBlock(
  geometry: PdfNativeTextRuleGeometry,
  style: PdfNativeTextRuleStyle,
): Uint8Array {
  const [a, b, c, d, e, f] = geometry.textMatrix;
  if (
    [a, b, c, d, e, f].some((value) => !Number.isFinite(value)) ||
    !Number.isFinite(geometry.fontSize) ||
    geometry.fontSize <= 0 ||
    !Number.isFinite(geometry.advance) ||
    geometry.advance <= 0
  ) {
    fail("Native text decoration geometry is invalid.");
  }
  const offset = (geometry.rise ?? 0) + geometry.fontSize * (style === "underline" ? UNDERLINE_OFFSET : STRIKETHROUGH_OFFSET);
  const thickness = geometry.fontSize * RULE_THICKNESS;
  const point = (u: number, v: number) =>
    `${pdfNumber(a * u + c * v + e)} ${pdfNumber(b * u + d * v + f)}`;
  return asciiBytes(
    `\nq ${colorOperator(geometry.color)} ${point(0, offset)} m ` +
      `${point(geometry.advance, offset)} l ` +
      `${point(geometry.advance, offset + thickness)} l ` +
      `${point(0, offset + thickness)} l h f Q\n`,
  );
}

/**
 * One rule a text edit rewrites in the same write (mirror of the server's
 * NativeTextRuleSpan): the bytes `removed` stand at `start..end` of the
 * stream, beside the edited run's text object (an underline ends at its
 * `BT`, a strikethrough starts after its `ET`), and `added` is written
 * there instead; either may be empty.
 */
export interface PdfNativeTextRuleSpan {
  readonly style: PdfNativeTextRuleStyle;
  readonly start: number;
  readonly end: number;
  readonly removed: Uint8Array;
  readonly added: Uint8Array;
}


// One rule block as pdfNativeTextRuleBlock writes it: its colour, then the
// four corners. The numbers are written by pdfNumber (no exponent).
const RULE_NUMBER = String.raw`-?\d+(?:\.\d+)?`;
const RULE_BLOCK =
  String.raw`\nq ((?:${RULE_NUMBER} ){1,4})(g|rg|k) (${RULE_NUMBER}) (${RULE_NUMBER}) m ` +
  String.raw`(${RULE_NUMBER}) (${RULE_NUMBER}) l (${RULE_NUMBER}) (${RULE_NUMBER}) l ` +
  String.raw`(${RULE_NUMBER}) (${RULE_NUMBER}) l h f Q\n`;
const RULE_BLOCK_ENDING = new RegExp(`${RULE_BLOCK}$`);
const RULE_BLOCK_STARTING = new RegExp(`^${RULE_BLOCK}`);
/** Longer than any block pdfNativeTextRuleBlock writes (twelve numbers of at
 * most fourteen characters, the colour and the operators). */
const MAX_RULE_BLOCK_LENGTH = 512;

/** One MRAI rule standing beside a run's text object. */
export interface PdfNativeTextRuleStanding {
  readonly style: PdfNativeTextRuleStyle;
  /** Byte span of the rule block in the stream. */
  readonly start: number;
  readonly end: number;
  readonly bytes: Uint8Array;
  /** The stretch of the run it runs under (native-text-rule-range.ts). */
  readonly range: PdfNativeTextRuleRange;
}

/**
 * The MRAI rule of `style` standing against `edge` -- ending there for an
 * underline (the run's `BT`), starting there for a strikethrough (just after
 * its `ET`), as the decoration writer draws them -- when it is a rule of the
 * run `geometry` describes: in the run's colour, on its underline or
 * strikethrough band, under the whole run or one stretch of it (WPS: two
 * characters underlined). Null when no such rule stands there.
 */
/** The most rule blocks one text object stands beside: a paragraph set again
 * draws a rule for each ruled run before (after) its one text object. */
const MAX_STACKED_RULES = 64;

export function pdfNativeTextRuleStandingAt(
  streamBytes: Uint8Array,
  edge: number,
  style: PdfNativeTextRuleStyle,
  geometry: PdfNativeTextRuleGeometry,
): PdfNativeTextRuleStanding | null {
  // The rule standing against the text object, or -- several rules stacked
  // there (one for each ruled run of a rewritten paragraph) -- the one of them
  // that is this run's: the others are passed over, nearest first.
  let at = edge;
  for (let passed = 0; passed < MAX_STACKED_RULES; passed += 1) {
    const found = ruleStandingAgainst(streamBytes, at, style, geometry);
    if (found.standing) return found.standing;
    if (found.blockStart === null) return null;
    at = style === "underline" ? found.blockStart : found.blockEnd!;
  }
  return null;
}

function ruleStandingAgainst(
  streamBytes: Uint8Array,
  edge: number,
  style: PdfNativeTextRuleStyle,
  geometry: PdfNativeTextRuleGeometry,
): { standing: PdfNativeTextRuleStanding | null; blockStart: number | null; blockEnd: number | null } {
  const underline = style === "underline";
  const from = underline ? Math.max(0, edge - MAX_RULE_BLOCK_LENGTH) : edge;
  const to = underline ? edge : Math.min(streamBytes.length, edge + MAX_RULE_BLOCK_LENGTH);
  const none = { standing: null, blockStart: null, blockEnd: null };
  if (!(edge >= 0 && edge <= streamBytes.length) || to <= from) return none;
  let text = "";
  for (let index = from; index < to; index += 1) text += String.fromCharCode(streamBytes[index]);
  const match = (underline ? RULE_BLOCK_ENDING : RULE_BLOCK_STARTING).exec(text);
  if (!match) return none;
  // This block's extent, whatever run it is the rule of: the next one
  // stacked against the text object is read past it.
  const blockStart = from + match.index;
  const blockEnd = blockStart + match[0].length;
  const other = { standing: null, blockStart, blockEnd };
  let colour: string;
  try {
    colour = colorOperator(geometry.color);
  } catch {
    return other;
  }
  if (`${match[1]}${match[2]}` !== colour) return other;
  const [a, b, c, d, e, f] = geometry.textMatrix;
  const determinant = a * d - b * c;
  if (
    [a, b, c, d, e, f, geometry.fontSize, geometry.advance].some((value) => !Number.isFinite(value)) ||
    !(Math.abs(determinant) > 1e-12) || !(geometry.advance > 0) || !(geometry.fontSize > 0)
  ) {
    return none;
  }
  const corners = match.slice(3, 11).map(Number);
  if (corners.some((value) => !Number.isFinite(value))) return other;
  // Each corner back in the run's own text space.
  const local = (index: number): readonly [number, number] => {
    const x = corners[index] - e;
    const y = corners[index + 1] - f;
    return [(d * x - c * y) / determinant, (-b * x + a * y) / determinant];
  };
  const [p0, p1, p2, p3] = [0, 2, 4, 6].map(local);
  const offset = (geometry.rise ?? 0) + geometry.fontSize * (underline ? UNDERLINE_OFFSET : STRIKETHROUGH_OFFSET);
  const thickness = geometry.fontSize * RULE_THICKNESS;
  // The corners carry six decimals in the space the matrix maps into.
  const slack = 2e-6 / Math.sqrt(Math.abs(determinant));
  const across = geometry.fontSize * 1e-4 + slack;
  const along = geometry.advance * 1e-4 + slack;
  const near = (value: number, expected: number, tolerance: number) => Math.abs(value - expected) <= tolerance;
  if (
    !near(p0[1], offset, across) || !near(p1[1], offset, across) ||
    !near(p2[1], offset + thickness, across) || !near(p3[1], offset + thickness, across) ||
    !near(p0[0], p3[0], along) || !near(p1[0], p2[0], along) ||
    !(p1[0] - p0[0] > along) || p0[0] < -along || p1[0] > geometry.advance + along
  ) {
    return other;
  }
  const share = (value: number) => Math.min(1, Math.max(0, value / geometry.advance));
  const stretch = { from: share(p0[0]), to: share(p1[0]) };
  if (!(stretch.to - stretch.from > PDF_NATIVE_TEXT_RULE_RANGE_EPSILON / 10)) return other;
  return { blockStart, blockEnd, standing: Object.freeze({
    style,
    start: blockStart,
    end: blockEnd,
    bytes: streamBytes.slice(blockStart, blockEnd),
    range: pdfNativeTextRuleRangeIsWhole(stretch) ? Object.freeze({ from: 0, to: 1 }) : Object.freeze(stretch),
  }) };
}

/**
 * Every rule block of this writer's form stacked against the text object at
 * `edges` -- ending at its `BT` (underlines), starting just after its `ET`
 * (strikethroughs) -- whatever run each is the rule of, nearest first.
 */
export function pdfNativeTextRuleBlocksBeside(
  streamBytes: Uint8Array,
  edges: Readonly<{ start: number; end: number }>,
): readonly Readonly<{ style: PdfNativeTextRuleStyle; start: number; end: number }>[] {
  const blocks: Readonly<{ style: PdfNativeTextRuleStyle; start: number; end: number }>[] = [];
  const read = (from: number, to: number) => {
    let text = "";
    for (let index = from; index < to; index += 1) text += String.fromCharCode(streamBytes[index]);
    return text;
  };
  for (let at = edges.start, passed = 0; passed < MAX_STACKED_RULES && at > 0 && at <= streamBytes.length; passed += 1) {
    const from = Math.max(0, at - MAX_RULE_BLOCK_LENGTH);
    const match = RULE_BLOCK_ENDING.exec(read(from, at));
    if (!match) break;
    const start = from + match.index;
    blocks.push(Object.freeze({ style: "underline" as const, start, end: at }));
    at = start;
  }
  for (let at = edges.end, passed = 0; passed < MAX_STACKED_RULES && at >= 0 && at < streamBytes.length; passed += 1) {
    const match = RULE_BLOCK_STARTING.exec(read(at, Math.min(streamBytes.length, at + MAX_RULE_BLOCK_LENGTH)));
    if (!match) break;
    blocks.push(Object.freeze({ style: "strikethrough" as const, start: at, end: at + match[0].length }));
    at += match[0].length;
  }
  return Object.freeze(blocks);
}

/**
 * The rules of `styles` standing beside the text object at `edges` that are
 * rules of the run `geometry` describes -- under the whole run or one
 * stretch of it (pdfNativeTextRuleStandingAt) -- as spans to take away
 * (`added` empty). A rule not standing there is not named.
 */
export function pdfNativeTextRulesStanding(
  streamBytes: Uint8Array,
  edges: Readonly<{ start: number; end: number }>,
  styles: readonly PdfNativeTextRuleStyle[],
  geometry: PdfNativeTextRuleGeometry,
): readonly PdfNativeTextRuleSpan[] {
  const rules: PdfNativeTextRuleSpan[] = [];
  for (const style of new Set(styles)) {
    const standing = pdfNativeTextRuleStandingAt(
      streamBytes, style === "underline" ? edges.start : edges.end, style, geometry);
    if (!standing) continue;
    rules.push(Object.freeze({
      style, start: standing.start, end: standing.end, removed: standing.bytes, added: new Uint8Array(0),
    }));
  }
  return Object.freeze(rules);
}
