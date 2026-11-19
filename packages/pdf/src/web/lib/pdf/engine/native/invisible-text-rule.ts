/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A rule beside a scanned line's invisible words.
 *
 * A scanned page is a picture with an OCR layer set in `3 Tr`: text a reader
 * selects and nobody sees. A reader who underlines such a line wants a line
 * under the scanned words, not the OCR's reading of them set again as a text
 * box. This module says where that rule goes and whether it would be seen:
 *
 *  - each word's place comes from the planner's rules-only proof
 *    (proveMraiNativeInvisibleTextRunStands), which offers no glyph property;
 *  - the words on one baseline are one line, and take one rule, its extent
 *    the words' own advances;
 *  - the rule is drawn beside the line's text object, in a `q … Q` of its own
 *    with its own colour (text-rule-block.ts), so `3 Tr` -- a text state --
 *    never reaches it. What could still hide it is asked here and refused:
 *    a clip, a graphics state or an optional-content group in force where it
 *    is drawn, and anything painted after it (a page whose image is drawn
 *    over its OCR text).
 */
import { inspectMraiNativePageContentStreams } from "./content-stream-transaction";
import { LATIN1_EXACT } from "./latin1";
import {
  proveMraiNativeInvisibleTextRunStands,
  type PdfNativeInvisibleRunStand,
  type PdfNativeListTextSelectionEvidence,
} from "./list-text-fidelity";
import { pdfNativeBoxRowStarts } from "../../text-edit/box-rows";
import {
  decorationAnchor,
  planMraiNativeTextDecoration,
  planMraiNativeTextDecorationRemoval,
  prepareMraiNativeTextDecoration,
  verifyMraiNativeTextDecorationCandidate,
  type PdfNativeTextDecorationGeometry,
  type PdfNativeTextDecorationPlan,
  type PdfNativeTextDecorationStyle,
} from "./text-decoration";
import { pdfNativeStreamHoldsTextRule, pdfNativeTextRuleStandingAt } from "./text-rule-block";

/** One line of the selection: the rule under the whole of it, and where
 * each of its items lies along it (in the units its text matrix maps). */
export interface PdfNativeInvisibleTextRuleLine {
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  /** The whole line's rule: the first word's matrix on the line's middle
   * baseline, the middle size of its words, the advance to the end of the
   * last word, black. */
  readonly geometry: PdfNativeTextDecorationGeometry;
  /** Null when a rule drawn beside the line would be seen; else why not. */
  readonly hidden: string | null;
}

export interface PdfNativeInvisibleTextRulesProof {
  /** One an item of the selection: its line and its stretch of it. Null for
   * a blank item and for one whose run is not proved. */
  readonly items: readonly (Readonly<{ line: number; from: number; to: number }> | null)[];
  /** The items with text whose run is not proved. */
  readonly unproven: readonly number[];
  readonly lines: readonly PdfNativeInvisibleTextRuleLine[];
}

const BLACK = Object.freeze({ kind: "gray" as const, components: Object.freeze([0]) });

/** The stands as lines, each item's stretch along its line's rule. An OCR
 * layer sets every word by its own box: its own size, and a baseline a
 * little above or below its neighbours'; and a page scanned askew sets each
 * word level, a step below (or above) the one before. The line's rule is one
 * straight rule, so it takes the middle size of the line's words and the
 * line through their baselines -- climbing or falling with the scan --
 * rather than stepping from word to word (0337 of the walk: a line falling
 * 9 pt over its 300; a level rule under its middle struck its last words
 * through). */
export function pdfNativeInvisibleTextRuleLines(
  stands: readonly (PdfNativeInvisibleRunStand | null)[],
): Readonly<{
  items: PdfNativeInvisibleTextRulesProof["items"];
  lines: readonly Readonly<{ contentStream: PdfNativeInvisibleRunStand["contentStream"]; geometry: PdfNativeTextDecorationGeometry }>[];
}> {
  interface Word { readonly item: number; readonly from: number; readonly to: number; readonly baseline: number; readonly size: number }
  const lines: Array<{ first: PdfNativeInvisibleRunStand; start: number; end: number; to: number; words: Word[] }> = [];
  for (const [item, stand] of stands.entries()) {
    // (A stand of no size or no advance names no rule: its item is unproven.)
    if (!stand || !(Math.abs(stand.fontSize) > 0) || !(stand.advance > 0)) continue;
    const line = lines[lines.length - 1];
    const placed = line ? stretchAlong(line.first, stand) : null;
    // The words of one scanned line stand within about half a line of the
    // word before -- a line of the taller of the word and the line's middle
    // size: a short word is set small and a bracket or a capital large and
    // lower -- while the line below starts a whole line down, and back at
    // the left. On the line, then, while it starts no further back than the
    // line has run (a word-wide tolerance for words an OCR layer lets overlap).
    const middleSize = line ? middle(line.words.map((word) => word.size)) : 0;
    if (line && placed &&
      Math.abs(placed.baseline - line.words[line.words.length - 1].baseline) <= Math.max(middleSize, placed.size) * 0.6 &&
      placed.from >= line.to - middleSize * 0.75 && stand.runSpan.start >= line.end) {
      line.end = Math.max(line.end, stand.runSpan.end);
      line.to = Math.max(line.to, placed.to);
      line.words.push({ item, ...placed });
      continue;
    }
    lines.push({
      first: stand,
      start: stand.runSpan.start,
      end: stand.runSpan.end,
      to: stand.advance,
      words: [{ item, from: 0, to: stand.advance, baseline: 0, size: Math.abs(stand.fontSize) }],
    });
  }
  const items: Array<Readonly<{ line: number; from: number; to: number }> | null> = stands.map(() => null);
  const geometries = lines.map((line, index) => {
    const [a, b, c, d, e, f] = line.first.textMatrix;
    const size = middle(line.words.map((word) => word.size));
    // The line through the words' baselines, in the first word's text
    // space: the middle of the slopes between its words (a word set off the
    // line by its own box does not turn it), level when the words do not
    // climb or fall by more than the eye reads.
    const slopes: number[] = [];
    for (const [at, left] of line.words.entries()) {
      for (const right of line.words.slice(at + 1)) {
        const run = (right.from + right.to - left.from - left.to) / 2;
        if (run > size * 2) slopes.push((right.baseline - left.baseline) / run);
      }
    }
    const fitted = slopes.length >= 3 ? middle(slopes) : 0;
    const slope = Math.abs(fitted) < 1e-3 ? 0 : fitted;
    const lift = middle(line.words.map((word) => word.baseline - slope * (word.from + word.to) / 2));
    // The rule's own axes in that space: along the line, and square to it.
    const length = Math.hypot(1, slope);
    const along = { x: 1 / length, y: slope / length };
    const across = { x: -slope / length, y: 1 / length };
    const place = (x: number, y: number) => x * along.x + (y - lift) * along.y;
    let advance = 0;
    for (const word of line.words) {
      const from = Math.max(0, place(word.from, word.baseline));
      const to = place(word.to, word.baseline);
      advance = Math.max(advance, to);
      items[word.item] = Object.freeze({ line: index, from, to });
    }
    return Object.freeze({
      contentStream: line.first.contentStream,
      geometry: Object.freeze({
        textMatrix: Object.freeze([
          a * along.x + c * along.y, b * along.x + d * along.y,
          a * across.x + c * across.y, b * across.x + d * across.y,
          e + c * lift, f + d * lift,
        ]) as unknown as PdfNativeTextDecorationGeometry["textMatrix"],
        fontSize: size,
        advance,
        color: BLACK,
        runSpan: Object.freeze({ start: line.start, end: line.end }),
      }),
    });
  });
  return Object.freeze({ items: Object.freeze(items), lines: Object.freeze(geometries) });
}

function middle(values: readonly number[]): number {
  const sorted = [...values].sort((left, right) => left - right);
  const half = sorted.length >> 1;
  return sorted.length === 0 ? 0 : sorted.length % 2 === 1 ? sorted[half] : (sorted[half - 1] + sorted[half]) / 2;
}

/** Where `stand` lies along `first`'s baseline, in `first`'s text space, how
 * far off that baseline it starts and how tall it is there: null when it is
 * another stream's, set in another page transform or turned another way. */
function stretchAlong(
  first: PdfNativeInvisibleRunStand,
  stand: PdfNativeInvisibleRunStand,
): Readonly<{ from: number; to: number; baseline: number; size: number }> | null {
  if (first.contentStream.objectNumber !== stand.contentStream.objectNumber ||
    first.contentStream.generation !== stand.contentStream.generation ||
    first.graphicsMatrix.some((value, index) => Math.abs(value - stand.graphicsMatrix[index]) > 1e-6)) return null;
  const [a, b, c, d, e, f] = first.textMatrix;
  const det = a * d - b * c;
  if (!(Math.abs(det) > 1e-12)) return null;
  // A point of the shared page transform's space in the first word's text space.
  const into = (x: number, y: number) => ({
    x: (d * (x - e) - c * (y - f)) / det,
    y: (-b * (x - e) + a * (y - f)) / det,
  });
  const [sa, sb, sc, sd, se, sf] = stand.textMatrix;
  const start = into(se, sf);
  const end = into(se + sa * stand.advance, sf + sb * stand.advance);
  const unit = Math.hypot(c, d);
  if (!(unit > 0)) return null;
  const size = Math.abs(stand.fontSize) * Math.hypot(sc, sd) / unit;
  // Turned another way: it climbs or falls along the first word's baseline.
  if (!(size > 0) || !(end.x > start.x) || Math.abs(end.y - start.y) > size * 0.1) return null;
  return { from: start.x, to: end.x, baseline: start.y, size };
}

interface ContentOperator {
  readonly operator: string;
  readonly start: number;
  readonly operands: readonly string[];
}

/** The operators of a page's drawing, each with its operands; null where
 * the bytes are not ones this reader reads (an inline image's data). */
function contentOperators(source: string): readonly ContentOperator[] | null {
  const operators: ContentOperator[] = [];
  let operands: string[] = [];
  let i = 0;
  const delimiter = /[\s()<>[\]{}/%]/;
  while (i < source.length) {
    const ch = source[i];
    if (/\s/.test(ch)) {
      i += 1;
    } else if (ch === "%") {
      while (i < source.length && source[i] !== "\n" && source[i] !== "\r") i += 1;
    } else if (ch === "(") {
      const start = i;
      let depth = 0;
      for (; i < source.length; i += 1) {
        if (source[i] === "\\") i += 1;
        else if (source[i] === "(") depth += 1;
        else if (source[i] === ")" && (depth -= 1) === 0) break;
      }
      if (i >= source.length) return null;
      i += 1;
      operands.push(source.slice(start, i));
    } else if (ch === "<" && source[i + 1] !== "<") {
      const end = source.indexOf(">", i);
      if (end < 0) return null;
      operands.push(source.slice(i, end + 1));
      i = end + 1;
    } else if (ch === "<" || ch === ">") {
      if (source[i + 1] !== ch) return null;
      operands.push(ch + ch);
      i += 2;
    } else if (ch === "[" || ch === "]" || ch === "{" || ch === "}") {
      operands.push(ch);
      i += 1;
    } else {
      const start = i;
      i += 1;
      while (i < source.length && !delimiter.test(source[i])) i += 1;
      const word = source.slice(start, i);
      if (ch === "/" || /^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(word) || word === "true" || word === "false" || word === "null") {
        operands.push(word);
      } else {
        // An inline image's data is not content this reader walks.
        if (word === "BI" || word === "ID") return null;
        operators.push({ operator: word, start, operands });
        operands = [];
      }
    }
  }
  return operators;
}

const PATH_PAINTS: ReadonlySet<string> = new Set(["f", "F", "f*", "B", "B*", "b", "b*", "S", "s"]);

/**
 * Whether a rule drawn beside the text objects of `runSpan` would be seen:
 * null when it would, else the reason it might not be.
 *
 * `streams` are the page's content streams in the order a reader draws them
 * and `index` the one the span is an offset into. The rule is written just
 * before the first text object of the span (an underline) or just after its
 * last (a strikethrough); it is seen when nothing in force there cuts or
 * blends it away and nothing drawn later lies over it. Later text (drawn or
 * not) and this writer's own rules are no cover; any other paint is taken
 * for one, without asking where it falls.
 */
export function pdfNativeRuleBesideRunHidden(
  streams: readonly string[],
  index: number,
  runSpan: Readonly<{ start: number; end: number }>,
): string | null {
  if (index < 0 || index >= streams.length) return "The run is not drawn by the page's own content.";
  const offset = streams.slice(0, index).reduce((sum, stream) => sum + stream.length + 1, 0);
  const from = offset + runSpan.start, to = offset + runSpan.end;
  const operators = contentOperators(streams.join("\n"));
  if (!operators) return "The page's content is not readable around the run.";
  interface Frame { clip: boolean; state: boolean; ops: string[] }
  let frame: Frame = { clip: false, state: false, ops: [] };
  const saved: Frame[] = [];
  const marked: boolean[] = [];
  const inForce = (): string | null =>
    frame.clip ? "A clip is in force where the rule would be drawn."
      : frame.state ? "A graphics state is in force where the rule would be drawn."
        : marked.some(Boolean) ? "The run is drawn in an optional-content group." : null;
  // The state at the text object the span starts in, and after the one it
  // ends in.
  let atBT: string | null = null;
  let started = false, ended = false;
  for (const entry of operators) {
    const { operator } = entry;
    if (operator === "Tr" && Number(entry.operands[entry.operands.length - 1]) >= 4) {
      return "The page's text is added to the clip.";
    }
    if (!started && entry.start >= from) {
      started = true;
      if (atBT) return atBT;
    }
    if (operator === "q") {
      saved.push(frame);
      frame = { clip: frame.clip, state: frame.state, ops: [] };
    } else if (operator === "Q") {
      frame = saved.pop() ?? { clip: false, state: false, ops: [] };
    } else {
      if (operator === "W" || operator === "W*") frame.clip = true;
      else if (operator === "gs") frame.state = true;
      else if (operator === "BDC" || operator === "BMC") marked.push(entry.operands[0] === "/OC");
      else if (operator === "EMC") marked.pop();
      else if (operator === "BT") atBT = inForce();
      if (started && !ended && operator === "ET" && entry.start >= to) {
        ended = true;
        const after = inForce();
        if (after) return after;
      }
      if (started && (operator === "Do" || operator === "sh")) {
        return "The page draws a picture after the run.";
      }
      if (started && PATH_PAINTS.has(operator)) {
        const ops = frame.ops;
        const rule = operator === "f" && saved.length > 0 && ops.length === 6 &&
          (ops[0] === "g" || ops[0] === "rg" || ops[0] === "k") &&
          ops[1] === "m" && ops[2] === "l" && ops[3] === "l" && ops[4] === "l" && ops[5] === "h";
        if (!rule) return "The page paints after the run.";
      }
      frame.ops.push(operator);
    }
  }
  return started && ended ? null : "The run's text objects were not found in the page's content.";
}

type RuleStyle = PdfNativeTextDecorationStyle;
const RULE_STYLES: readonly RuleStyle[] = ["underline", "strikethrough"];

interface PageStream {
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  bytes: Uint8Array;
  text: string;
}

const sameStream = (
  left: Readonly<{ objectNumber: number; generation: number }>,
  right: Readonly<{ objectNumber: number; generation: number }>,
) => left.objectNumber === right.objectNumber && left.generation === right.generation;

/** The proof, and the page's streams it was read on (one inspection serves
 * the proof, the rules standing and the changes). */
async function proveRules(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
}): Promise<Readonly<{ proof: PdfNativeInvisibleTextRulesProof; streams: PageStream[] }>> {
  const selected = input.selection.items ?? [input.selection.item];
  const stands = await proveMraiNativeInvisibleTextRunStands(input);
  const { items, lines } = pdfNativeInvisibleTextRuleLines(stands);
  const unproven = selected.flatMap((item, index) => (item.str.trim() && !items[index] ? [index] : []));
  let streams: PageStream[] = [];
  if (lines.length > 0) {
    const inspected = await inspectMraiNativePageContentStreams(input.sourceBytes, input.selection.pageNumber);
    streams = inspected.map((stream) => ({
      contentStream: stream.contentStream,
      bytes: stream.streamBytes,
      text: LATIN1_EXACT.decode(stream.streamBytes),
    }));
  }
  const texts = streams.map((stream) => stream.text);
  return Object.freeze({
    streams,
    proof: Object.freeze({
      items,
      unproven: Object.freeze(unproven),
      lines: Object.freeze(lines.map((line) => {
        const index = streams.findIndex((stream) => sameStream(stream.contentStream, line.contentStream));
        let hidden = pdfNativeRuleBesideRunHidden(texts, index, line.geometry.runSpan);
        // Where the rule would be written: beside the line's own text
        // object(s), in this one stream. A line whose text object the
        // page's streams split, or whose objects stand in different
        // transforms, has no such place.
        if (hidden === null) {
          try {
            for (const style of RULE_STYLES) decorationAnchor(streams[index].text, line.geometry.runSpan, style);
          } catch (error) {
            hidden = error instanceof Error && error.message ? error.message : "The run's text object is not one a rule stands beside.";
          }
        }
        return Object.freeze({ contentStream: line.contentStream, geometry: line.geometry, hidden });
      })),
    }),
  });
}

/**
 * The rules-only proof of a scanned line's selection: its lines, where each
 * item lies along its line, and whether a rule beside each line would be
 * seen. It grants nothing but a rule: no property of the words is read.
 */
export async function proveMraiNativeInvisibleTextRules(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
}): Promise<PdfNativeInvisibleTextRulesProof> {
  return (await proveRules(input)).proof;
}

/** The most items of one box this reader proves: each is a plan of its own
 * (a paragraph of the walk's scans holds up to 37 OCR words). A longer box
 * keeps the text box for its rules. */
export const PDF_NATIVE_INVISIBLE_TEXT_RULE_MAX_ITEMS = 64;

/** Stretches of a box's text, in its characters: sorted, apart. */
export type PdfNativeInvisibleTextRuleStretches = readonly Readonly<{ start: number; end: number }>[];

export interface PdfNativeInvisibleTextRulesStanding {
  /** Whether every line of the box takes a rule beside its invisible run:
   * every item with text proved, every line's rule one that would be seen. */
  readonly proven: boolean;
  /** Why not (the log's, not the reader's). */
  readonly reason: string | null;
  /** The characters of the box's text each rule runs under now. */
  readonly underline: PdfNativeInvisibleTextRuleStretches;
  readonly strikethrough: PdfNativeInvisibleTextRuleStretches;
}

/** The format asked: the whole box's rules and parts' (offsets in the box's
 * text, or -- `row` given -- in that item's own). Nothing but rules. */
export interface PdfNativeInvisibleTextRuleFormat {
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  readonly segments?: readonly Readonly<{
    row?: number; start: number; end: number; underline?: boolean; strikethrough?: boolean;
  }>[];
}

export interface PdfNativeInvisibleTextRuleOperation {
  readonly style: RuleStyle;
  readonly direction: "add" | "remove";
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  /** An add: the rule's own geometry on the bytes the operations before it left. */
  readonly geometry?: PdfNativeTextDecorationGeometry;
  /** A removal: the rule's bytes, where they stand on those bytes. */
  readonly history?: Readonly<{ offset: number; ruleBytes: readonly number[] }>;
}

export interface PdfNativeInvisibleTextRuleChanges {
  /** In order, each planned on the bytes the one before it wrote. None:
   * the page already draws what was asked. */
  readonly operations: readonly PdfNativeInvisibleTextRuleOperation[];
  /** The bytes' digest once all of them are written (the source's for none). */
  readonly resultSha256: string;
  /** The rules standing then, read back from those bytes. */
  readonly standing: PdfNativeInvisibleTextRulesStanding;
}

interface Stretch { readonly from: number; readonly to: number }
type LineRules = Record<RuleStyle, Stretch | null>;

interface RuleState {
  readonly proof: PdfNativeInvisibleTextRulesProof;
  readonly streams: PageStream[];
  readonly reason: string | null;
  /** Where each item's text starts in the box's text. */
  readonly starts: readonly number[];
  readonly boxText: string;
  readonly texts: readonly string[];
}

async function readRuleState(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
}): Promise<RuleState | Readonly<{ proof: null; reason: string }>> {
  const selected = input.selection.items ?? [input.selection.item];
  if (selected.length > PDF_NATIVE_INVISIBLE_TEXT_RULE_MAX_ITEMS) {
    return { proof: null, reason: `The box holds ${selected.length} items.` };
  }
  const boxText = input.selection.item.str;
  const starts = pdfNativeBoxRowStarts(boxText, selected);
  if (!starts) return { proof: null, reason: "The box's text does not hold its items in order." };
  const { proof, streams } = await proveRules(input);
  const hidden = proof.lines.find((line) => line.hidden !== null);
  const reason = proof.lines.length === 0 ? "No run of the box is proved."
    : proof.unproven.length > 0 ? `${proof.unproven.length} of the box's items are not proved.`
      : hidden ? hidden.hidden : null;
  return { proof, streams, reason, starts, boxText, texts: selected.map((item) => item.str) };
}

/** The rule of `style` standing beside `line`, as a stretch of its advance. */
function standingBeside(
  stream: PageStream | undefined,
  geometry: PdfNativeTextDecorationGeometry,
  style: RuleStyle,
): Readonly<{ offset: number; bytes: Uint8Array; stretch: Stretch }> | null {
  if (!stream) return null;
  try {
    const anchor = decorationAnchor(stream.text, geometry.runSpan, style);
    const standing = pdfNativeTextRuleStandingAt(stream.bytes, anchor, style, geometry);
    return standing ? {
      offset: standing.start,
      bytes: standing.bytes,
      stretch: { from: standing.range.from * geometry.advance, to: standing.range.to * geometry.advance },
    } : null;
  } catch {
    return null;
  }
}

function linesStanding(state: RuleState, spans?: readonly Readonly<{ start: number; end: number }>[]): LineRules[] {
  // Nearly every page holds no rule of this writer's: one scan says so.
  const holds = state.streams.map((stream) => pdfNativeStreamHoldsTextRule(stream.bytes));
  return state.proof.lines.map((line, index) => {
    const at = state.streams.findIndex((stream) => sameStream(stream.contentStream, line.contentStream));
    const geometry = spans ? { ...line.geometry, runSpan: spans[index] } : line.geometry;
    const read = (style: RuleStyle) => (at >= 0 && holds[at] ? standingBeside(state.streams[at], geometry, style)?.stretch ?? null : null);
    return { underline: read("underline"), strikethrough: read("strikethrough") };
  });
}

/** Where character `at` (0..its length) of an item stands along its line:
 * an OCR word is one show fitted to the scanned word's box, so its
 * characters share its advance evenly. */
function characterPlace(place: Readonly<{ from: number; to: number }>, length: number, at: number): number {
  return place.from + (place.to - place.from) * (length > 0 ? at / length : 0);
}

/** The characters of the box's text under each line's stretch. */
function charactersUnder(state: RuleState, rules: readonly LineRules[], style: RuleStyle): Uint8Array {
  const on = new Uint8Array(state.boxText.length);
  for (const [index, place] of state.proof.items.entries()) {
    const stretch = place ? rules[place.line]?.[style] : null;
    if (!place || !stretch) continue;
    const slack = state.proof.lines[place.line].geometry.advance * 1e-3;
    const length = state.texts[index].length;
    for (let at = 0; at < length; at += 1) {
      const middle = characterPlace(place, length, at + 0.5);
      if (middle >= stretch.from - slack && middle <= stretch.to + slack) on[state.starts[index] + at] = 1;
    }
  }
  return on;
}

/** The characters on, as stretches: a blank between two ruled characters is
 * ruled with them (Word underlines the space between two underlined words). */
function stretchesOf(on: Uint8Array, text: string): PdfNativeInvisibleTextRuleStretches {
  const stretches: { start: number; end: number }[] = [];
  let open: { start: number; end: number } | null = null;
  for (let at = 0; at < on.length; at += 1) {
    if (on[at]) {
      if (open && !text.slice(open.end, at).trim()) open.end = at + 1;
      else stretches.push(open = { start: at, end: at + 1 });
    }
  }
  return Object.freeze(stretches.map((stretch) => Object.freeze(stretch)));
}

function standingOf(state: RuleState, rules: readonly LineRules[]): PdfNativeInvisibleTextRulesStanding {
  return Object.freeze({
    proven: state.reason === null,
    reason: state.reason,
    underline: stretchesOf(charactersUnder(state, rules, "underline"), state.boxText),
    strikethrough: stretchesOf(charactersUnder(state, rules, "strikethrough"), state.boxText),
  });
}

const NOT_PROVEN = (reason: string): PdfNativeInvisibleTextRulesStanding =>
  Object.freeze({ proven: false, reason, underline: Object.freeze([]), strikethrough: Object.freeze([]) });







export async function readMraiNativeInvisibleTextRulesStanding(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
}): Promise<PdfNativeInvisibleTextRulesStanding> {
  const state = await readRuleState(input);
  if (!state.proof) return NOT_PROVEN(state.reason);
  return standingOf(state, state.reason === null ? linesStanding(state) : state.proof.lines.map(() => ({ underline: null, strikethrough: null })));
}

/**
 * The rule changes a format asks of a scanned box, each a decoration the
 * engine writes (planMraiNativeTextDecoration / …Removal): a rule drawn
 * beside the line's invisible run in a `q … Q` of its own, the scan and its
 * OCR text untouched. One rule of a style a line: a stretch asked apart
 * from the one standing is not the page's to draw (as beside a drawn run).
 *
 * The operations are tried here, each on the bytes the one before wrote,
 * and the result is proved again and its rules read back: null when the
 * box is not proven, when the format is not one of rules, or when what
 * would stand is not what was asked.
 */
export async function planMraiNativeInvisibleTextRuleChanges(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly format: PdfNativeInvisibleTextRuleFormat;
}): Promise<PdfNativeInvisibleTextRuleChanges | null> {
  // Rules and nothing else: a format that also names a property of the
  // glyphs (a weight, a colour, a size) is not half written here.
  const named = (record: object, allowed: readonly string[]) => Object.entries(record)
    .every(([key, value]) => value === undefined || allowed.includes(key));
  if (!named(input.format, ["underline", "strikethrough", "segments"]) ||
    (input.format.segments ?? []).some((segment) =>
      !named(segment, ["row", "start", "end", "underline", "strikethrough"]))) return null;
  const state = await readRuleState(input);
  if (!state.proof || state.reason !== null) return null;
  const { proof, streams, starts, texts, boxText } = state;
  const before = linesStanding(state);
  // What is asked, in the box's characters, over what stands.
  const wanted: Record<RuleStyle, Uint8Array> = {
    underline: charactersUnder(state, before, "underline"),
    strikethrough: charactersUnder(state, before, "strikethrough"),
  };
  const touched: Record<RuleStyle, Uint8Array> = {
    underline: new Uint8Array(boxText.length),
    strikethrough: new Uint8Array(boxText.length),
  };
  const set = (style: RuleStyle, start: number, end: number, value: boolean) => {
    for (let at = Math.max(0, start); at < Math.min(boxText.length, end); at += 1) {
      wanted[style][at] = value ? 1 : 0;
      touched[style][at] = 1;
    }
  };
  for (const style of RULE_STYLES) {
    const whole = input.format[style];
    if (whole !== undefined) set(style, 0, boxText.length, whole);
  }
  for (const segment of input.format.segments ?? []) {
    const base = segment.row === undefined || segment.row < 0 ? 0 : starts[segment.row];
    if (base === undefined || !(segment.end > segment.start)) return null;
    for (const style of RULE_STYLES) {
      const value = segment[style];
      if (value !== undefined) set(style, base + segment.start, base + segment.end, value);
    }
  }
  // Each line's stretch of each style asked of it.
  const asked: Array<Partial<Record<RuleStyle, Stretch | null>>> = proof.lines.map(() => ({}));
  for (const [lineIndex, line] of proof.lines.entries()) {
    const advance = line.geometry.advance;
    for (const style of RULE_STYLES) {
      let any = false, from = Infinity, to = -Infinity, first = -1, last = -1, count = 0, gaps = false;
      for (const [index, place] of proof.items.entries()) {
        if (!place || place.line !== lineIndex) continue;
        const length = texts[index].length;
        for (let at = 0; at < length; at += 1) {
          if (!texts[index][at].trim()) continue;
          const character = starts[index] + at;
          any ||= touched[style][character] === 1;
          if (wanted[style][character]) {
            if (first < 0) first = count;
            else if (last !== count - 1) gaps = true;
            last = count;
            from = Math.min(from, characterPlace(place, length, at));
            to = Math.max(to, characterPlace(place, length, at + 1));
          }
          count += 1;
        }
      }
      if (!any) continue;
      // Two stretches of one style apart on one line.
      if (gaps) return null;
      asked[lineIndex][style] = first < 0 ? null
        : first === 0 && last === count - 1 ? { from: 0, to: advance }
          : { from: Math.max(0, from), to: Math.min(advance, to) };
    }
  }
  const near = (left: Stretch | null, right: Stretch | null, advance: number) =>
    left === null || right === null ? left === right
      : Math.abs(left.from - right.from) <= advance * 2e-3 && Math.abs(left.to - right.to) <= advance * 2e-3;
  const spans = proof.lines.map((line) => ({ ...line.geometry.runSpan }));
  const shift = (contentStream: PageStream["contentStream"], at: number, by: number) => {
    for (const [index, line] of proof.lines.entries()) {
      if (!sameStream(line.contentStream, contentStream)) continue;
      if (spans[index].start >= at) spans[index].start += by;
      if (spans[index].end > at) spans[index].end += by;
    }
  };
  const operations: PdfNativeInvisibleTextRuleOperation[] = [];
  let bytes = input.sourceBytes;
  let resultSha256 = input.selection.sourceSha256;
  const write = async (plan: PdfNativeTextDecorationPlan, stream: PageStream): Promise<boolean> => {
    // One stream's bytes, in place: a write that joins the page's streams
    // or copies a Form moves every span this plan counts by.
    if (plan.request.pageContentConsolidation || plan.request.formInstanceIsolation ||
      !sameStream(plan.contentStream, stream.contentStream)) return false;
    const prepared = await prepareMraiNativeTextDecoration(bytes, plan);
    await verifyMraiNativeTextDecorationCandidate(bytes, plan, prepared);
    if (bytes !== input.sourceBytes) bytes.fill(0);
    bytes = prepared.candidateBytes;
    resultSha256 = prepared.candidateSha256;
    const rule = Uint8Array.from(plan.ruleBytes);
    const next = new Uint8Array(stream.bytes.length + (plan.direction === "add" ? rule.length : -rule.length));
    next.set(stream.bytes.subarray(0, plan.ruleOffset));
    if (plan.direction === "add") {
      next.set(rule, plan.ruleOffset);
      next.set(stream.bytes.subarray(plan.ruleOffset), plan.ruleOffset + rule.length);
    } else {
      next.set(stream.bytes.subarray(plan.ruleOffset + rule.length), plan.ruleOffset);
    }
    stream.bytes = next;
    stream.text = LATIN1_EXACT.decode(next);
    shift(stream.contentStream, plan.direction === "add" ? plan.ruleOffset : plan.ruleOffset + rule.length,
      plan.direction === "add" ? rule.length : -rule.length);
    return true;
  };
  try {
    for (const [lineIndex, line] of proof.lines.entries()) {
      const stream = streams.find((candidate) => sameStream(candidate.contentStream, line.contentStream));
      if (!stream) return null;
      for (const style of RULE_STYLES) {
        const want = asked[lineIndex][style];
        if (want === undefined || near(before[lineIndex][style], want, line.geometry.advance)) continue;
        const standing = standingBeside(stream, { ...line.geometry, runSpan: spans[lineIndex] }, style);
        if (standing) {
          const history = Object.freeze({ offset: standing.offset, ruleBytes: Object.freeze([...standing.bytes]) });
          const plan = await planMraiNativeTextDecorationRemoval({
            sourceBytes: bytes, pageNumber: input.selection.pageNumber, style,
            offset: history.offset, ruleBytes: history.ruleBytes, contentStream: line.contentStream,
          });
          if (!await write(plan, stream)) return null;
          operations.push(Object.freeze({ style, direction: "remove", contentStream: line.contentStream, history }));
        }
        if (want) {
          // The stretch's own rule: the line's, started `from` along it.
          const [a, b, c, d, e, f] = line.geometry.textMatrix;
          const geometry: PdfNativeTextDecorationGeometry = Object.freeze({
            textMatrix: Object.freeze([a, b, c, d, e + a * want.from, f + b * want.from]) as unknown as
              PdfNativeTextDecorationGeometry["textMatrix"],
            fontSize: line.geometry.fontSize,
            advance: want.to - want.from,
            color: line.geometry.color,
            runSpan: Object.freeze({ ...spans[lineIndex] }),
          });
          const plan = await planMraiNativeTextDecoration({
            sourceBytes: bytes, pageNumber: input.selection.pageNumber, style, geometry, contentStream: line.contentStream,
          });
          if (!await write(plan, stream)) return null;
          operations.push(Object.freeze({ style, direction: "add", contentStream: line.contentStream, geometry }));
        }
      }
    }
    if (operations.length === 0) {
      return Object.freeze({ operations: Object.freeze([]), resultSha256, standing: standingOf(state, before) });
    }
    // What stands now, proved again on the bytes written: every line where
    // it was, each rule the one asked and the others as they were.
    const after = await readRuleState({
      sourceBytes: bytes,
      selection: Object.freeze({ ...input.selection, sourceSha256: resultSha256 }),
    });
    if (!after.proof || after.reason !== null || after.proof.lines.length !== proof.lines.length) return null;
    const read = linesStanding(after);
    for (const [lineIndex, line] of proof.lines.entries()) {
      for (const style of RULE_STYLES) {
        const want = asked[lineIndex][style];
        if (!near(read[lineIndex][style], want === undefined ? before[lineIndex][style] : want, line.geometry.advance)) return null;
      }
    }
    return Object.freeze({ operations: Object.freeze(operations), resultSha256, standing: standingOf(after, read) });
  } catch {
    return null;
  } finally {
    if (bytes !== input.sourceBytes) bytes.fill(0);
  }
}
