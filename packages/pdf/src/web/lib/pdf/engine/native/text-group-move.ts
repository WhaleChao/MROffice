/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Move or resize a whole text box natively: every show a box's proof names --
 * the lines of a paragraph, the runs of a line, across text objects, marked
 * content, graphics states and colour changes -- drawn somewhere else, or
 * scaled, as the page draws it.
 *
 * A box is moved by one affine map of the page's user space, `x' = a x + e`,
 * `y' = d y + f` (a pure move when a = d = 1). The page draws text through
 * the text line matrix its positioning operators set (`Tm`, `Td`, `TD`, `T*`,
 * `'`, `"`); a show is placed where that matrix and the pen after the shows
 * before it on the line put it. So the map is carried by the positioning
 * operators alone, and every glyph show, every font, colour, spacing and
 * marked-content operator keeps its bytes:
 *
 * - A text object is cut into segments at each positioning operator (and at
 *   `BT`). A segment whose shows are all the box's is set in the moved space:
 *   its line matrix is the source's times G = CTM x M x CTM^-1, the map read
 *   in the object's own space. A segment of other text keeps the source's.
 *   A segment holding both is refused: its box text continues a line of other
 *   text, which no positioning operator places on its own.
 * - Where two segments of the same kind follow each other, a relative
 *   operator carries the map on by itself (T(t) x (L x G) = (T(t) x L) x G)
 *   and keeps its bytes. At a boundary between the kinds the operator is
 *   rewritten to reach the other space from where the rewritten stream
 *   stands: its translation when the linear part is the same (a move), else
 *   an absolute `Tm` (a resize). `TD` keeps the leading it sets, `'` and `"`
 *   become the positioning plus the same string shown by `Tj`.
 * - A `Tm` of the box is written times G; one of other text keeps its bytes.
 *   A box segment that `BT` opens gets a `Tm` of its own after `BT`.
 *
 * Nothing outside text objects is touched, so no graphics state set around
 * the text leaks: the text keeps the clip, colour and transparency it was
 * drawn under, and the moved box has to land where that clip still shows it.
 *
 * Where a clip of its own would hide the landing (a web page clips each
 * block to its element box, a Google Docs export each column to its half of
 * the page), the box is written outside every clip instead: each of its
 * shows is taken out of its place (`'` and `"` keep their move to the next
 * line), and a block appended at the tail of the stream draws the same
 * shows again, run by run: `q`, the text object's CTM, every graphics and
 * text state operator in force where the run starts (the last writer of each
 * state, in the order the page wrote them), `BT`, the run's line matrix
 * times G, and the run's own operators with its marked content left out.
 * Only state that can be read off the stream is carried: a run whose state
 * at the tail the block cannot set back (an ExtGState it does not name), a
 * pattern colour, marked content naming a replacement text or a layer, a
 * Type3 font or a soft mask is refused, as is a box sharing a line with
 * other text.
 *
 * Numbers are written in micro-units with one rounding rule (the text move's,
 * formatMraiNativeTextMoveMicro), and the rewrite is derived from the stream,
 * the owned shows and the map alone, so replay derives the same bytes
 * (native_text_group_move.py). An undo writes the source span's own bytes back, and
 * is proven by deriving the move again from them.
 */
import { LATIN1_EXACT } from "./latin1";
import { MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS, invalid, mraiNativeTextGroupTransformIsIdentity, readMraiNativeTextGroupMoveDraft } from "./text-group-move-model";
export { MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS, mraiNativeTextGroupTransformIsIdentity, readMraiNativeTextGroupMoveDraft };
import {
  MRAI_NATIVE_TEXT_GROUP_MOVE_KIND,
  MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION,
} from "./vocabulary";
export {
  MRAI_NATIVE_TEXT_GROUP_MOVE_KIND,
  MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION,
};
import {
  MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION,
  mraiNativeReaderChunkRestIsBlank,
  mraiNativeShowIsBlankInLeafFont,
  pdfNativeTextFontIdentity,
  prepareMraiNativeTextEditingSelection,
  proveMraiNativeListTextSelectionForWriter,
  readMraiNativeVisiblePageBounds,
  qtFakeBoldTwinSpan,
  type PdfNativeListTextFidelityEvidence,
  type PdfNativeListTextSelectionEvidence,
  type PdfNativeListTextSelectionItemEvidence,
  type PdfNativeSourceTextAppearanceRun,
  type PdfNativeTextPropertyDraft,
} from "./list-text-fidelity";
import {
  inspectMraiNativePageContentStreams,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  readMraiNativePageResourceDictionaries,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativePageContentStreamInspection,
  type PdfNativePreparedContentStreamTransaction,
  type PdfNativeReplacementPiece,
} from "./content-stream-transaction";
import { formatMraiNativeTextMoveMicro } from "./text-move";
import { pdfNativeTextRuleBlocksBeside } from "./text-rule-block";
import { MRAI_NATIVE_BOX_MARK_OPEN, planMraiNativeBoxMark, type PdfNativeBoxMark } from "./box-mark";
import { decodePdfSourceName, pdfSourceNamesRead } from "./pdf-source-name";
import { PDF_NATIVE_SYNTHETIC_BOLD_STROKE_EM } from "../../native-synthetic-style";

/** Record kind of a group move; it rides the text edit command. */

/** A resize is refused past this factor either way (the editor's own bound). */
const MIN_GROUP_SCALE = 0.05;
const MAX_GROUP_SCALE = 20;
const MAX_GROUP_MOVE_DISTANCE = 100_000;
const MAX_GROUP_MICRO = 1_000_000 * 1_000_000;
/** At most this many operators are rewritten by one group move. */
const MAX_GROUP_EDITS = 4_096;
/** A record carries at most this many replacement bytes (the replay's bound). */
const MAX_GROUP_REPLACEMENT_BYTES = 64 * 1024;
/** A box may lose at most this share of its visible area to a clip or the page
 * edge: its bounds are the reader items' boxes padded for descenders, not ink. */
const CLIP_VISIBLE_SLACK = 0.1;
/** Across its lines a box's bounds are its glyphs: a clip of its own may hide
 * no more of its width at the landing than it did before, to this (points). */
const CLIP_WIDTH_TOLERANCE = 0.5;


type Span = Readonly<{ start: number; end: number }>;
/** A show the box owns: whole, or only the range of its codes a reader chunk
 * holds when the rest of it is other text (codeStart/codeEnd, byte offsets). */
type GroupShow = Readonly<{ start: number; end: number; codeStart?: number; codeEnd?: number }>;
type Matrix = readonly [number, number, number, number, number, number];

/** The map of the page's user space a group move applies. `b` and `c`, when
 * present, turn the box (WPS turns a text block with its frame's handle);
 * absent, the map keeps the page's axes. */
export interface PdfNativeTextGroupTransform {
  readonly a: number;
  readonly b?: number;
  readonly c?: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

/** The map's linear part, turn included. */
function groupLinear(transform: PdfNativeTextGroupTransform): readonly [number, number, number, number] {
  return [transform.a, transform.b ?? 0, transform.c ?? 0, transform.d];
}

/** A map that only moves: its linear part is the identity. */
function groupMapMovesOnly(transform: PdfNativeTextGroupTransform): boolean {
  const [a, b, c, d] = groupLinear(transform);
  return a === 1 && b === 0 && c === 0 && d === 1;
}

/** A show of the box (or the range of its code bytes `codeStart..codeEnd`)
 * a group bold is set on. */
export type PdfNativeTextGroupShow = GroupShow;

/**
 * A synthetic bold of the box, the way WPS (and Word) draw a bold the face
 * has none of (lib/pdf/native-synthetic-style.ts): `on` strokes each show's
 * glyph outlines in its own fill colour, `2 Tr` under a width of 0.035 em,
 * and puts back after it the rendering mode, stroke width and stroke colour
 * in force; `!on` fills a show the page strokes so (`0 Tr`, the `2 Tr` put
 * back after it). Every show of the box, or `shows` only (a part of it).
 */
export interface PdfNativeTextGroupBold {
  readonly on: boolean;
  readonly shows?: readonly GroupShow[];
}

/** What an editor draft asks of a group: its map, and a weight. */
export interface PdfNativeTextGroupEdit extends PdfNativeTextGroupTransform {
  readonly bold?: PdfNativeTextGroupBold;
}

/** What the editor needs to offer a native group move or resize. */
export interface PdfNativeTextGroupMoveProof {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION;
  /** The box's own state as the page draws it; a group move changes none of
   * it. Its `syntheticBold` is the group's weight when the group can take a
   * synthetic bold on or off (every show filled, or every show stroked in
   * place), absent when it cannot. */
  readonly baseline: PdfNativeTextPropertyDraft;
}

export interface PdfNativeTextGroupMoveEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  /** The source span the move rewrites; its sha is the source bytes'. */
  readonly operatorSpan: Readonly<{ start: number; end: number; expectedSha256: string }>;
  readonly transform: PdfNativeTextGroupTransform;
  /** Every show the box owns, in stream order (source offsets). */
  readonly shows: readonly GroupShow[];
  /** The no-op proof of the box: which shows it is. */
  readonly sourceFidelity: PdfNativeListTextFidelityEvidence;
  /** The no-op request's bytes, when a writer proof owns the box. */
  readonly sourceProofReplacement?: readonly number[];
  readonly sourceDuplicateSpans?: readonly Span[];
  /** A paragraph owned by its items proved alone (the paragraph rewrite's
   * read proof): each item's proof, `sourceFidelity` being the first. */
  readonly sourceItems?: readonly PdfNativeListTextFidelityEvidence[];
  /** Set on the exact byte restore of a group move (its undo). */
  readonly restore?: Readonly<{ originTransactionId: string; direction: "undo" }>;
  /** The box is written outside every clip: its shows taken out of their
   * place and drawn again at the tail of the stream (a clip of their own
   * would hide the landing). The span runs to the stream's end. */
  readonly outsideClip?: true;
  /** A synthetic bold put on or taken off the box's shows with the map (an
   * identity map then writes the weight alone). */
  readonly bold?: PdfNativeTextGroupBold;
  /** The box mark the move wrote round the box (box-mark.ts), the record's
   * own: its restore finds the move's bytes inside it, and leaves it. */
  readonly boxMark?: PdfNativeBoxMark;
}

export interface PdfNativeTextGroupMovePlan {
  readonly evidence: PdfNativeTextGroupMoveEvidence;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  /** The selection as the reader reads the moved box, bound to the output. */
  readonly movedSelection: PdfNativeListTextSelectionEvidence;
  /** The source span's own bytes, which the move's undo writes back. */
  readonly originalBytes: Uint8Array;
  /** The box's selection as the proof that owns it read it. */
  readonly proofSelection: PdfNativeListTextSelectionEvidence;
}

function refuse(message: string): never {
  throw new PdfNativeTransactionError("unsupported-document", message);
}


// ---------------------------------------------------------------------------
// Tokens (mirror of _group_tokens). PDF white space and delimiters only.

const WHITE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20]);
const DELIMITER = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25]);
const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;

interface Token {
  readonly start: number;
  readonly end: number;
  readonly operator: boolean;
}

function skipWhite(source: string, index: number): number {
  let i = index;
  for (;;) {
    while (i < source.length && WHITE.has(source.charCodeAt(i))) i += 1;
    if (i < source.length && source.charCodeAt(i) === 0x25) {
      while (i < source.length && source.charCodeAt(i) !== 0x0a && source.charCodeAt(i) !== 0x0d) i += 1;
      continue;
    }
    return i;
  }
}

/** End of the token at `index`, or -1 when it is not one this reader knows. */
function tokenEnd(source: string, index: number): number {
  const first = source.charCodeAt(index);
  if (first === 0x28) {
    let depth = 0;
    for (let i = index; i < source.length; i += 1) {
      const c = source.charCodeAt(i);
      if (c === 0x5c) {
        i += 1;
      } else if (c === 0x28) {
        depth += 1;
      } else if (c === 0x29) {
        depth -= 1;
        if (depth === 0) return i + 1;
      }
    }
    return -1;
  }
  if (first === 0x3c) {
    if (source.charCodeAt(index + 1) === 0x3c) {
      let i = index + 2;
      for (;;) {
        i = skipWhite(source, i);
        if (i >= source.length) return -1;
        if (source.charCodeAt(i) === 0x3e && source.charCodeAt(i + 1) === 0x3e) return i + 2;
        const end = tokenEnd(source, i);
        if (end < 0) return -1;
        i = end;
      }
    }
    for (let i = index + 1; i < source.length; i += 1) {
      const c = source.charCodeAt(i);
      if (c === 0x3e) return i + 1;
      const hex = (c >= 0x30 && c <= 0x39) || (c >= 0x41 && c <= 0x46) || (c >= 0x61 && c <= 0x66);
      if (!hex && !WHITE.has(c)) return -1;
    }
    return -1;
  }
  if (first === 0x5b) {
    let i = index + 1;
    for (;;) {
      i = skipWhite(source, i);
      if (i >= source.length) return -1;
      if (source.charCodeAt(i) === 0x5d) return i + 1;
      const end = tokenEnd(source, i);
      if (end < 0) return -1;
      i = end;
    }
  }
  if (first === 0x29 || first === 0x3e || first === 0x5d || first === 0x7b || first === 0x7d) return -1;
  let i = index + 1;
  while (i < source.length) {
    const c = source.charCodeAt(i);
    if (WHITE.has(c) || DELIMITER.has(c)) break;
    i += 1;
  }
  return i;
}

function isOperatorStart(code: number): boolean {
  return (code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a) || code === 0x27 || code === 0x22;
}

/**
 * The tokens of a content stream, or null when any of it is not readable.
 * An inline image's data (`BI … ID <data> EI`) is skipped whole: it ends at
 * the first `EI` with white space on both sides.
 */
function groupTokens(source: string): Token[] | null {
  const tokens: Token[] = [];
  let i = 0;
  for (;;) {
    i = skipWhite(source, i);
    if (i >= source.length) return tokens;
    const end = tokenEnd(source, i);
    if (end < 0) return null;
    const operator = isOperatorStart(source.charCodeAt(i));
    tokens.push({ start: i, end, operator });
    i = end;
    if (operator && end - tokens[tokens.length - 1].start === 2 && source.startsWith("ID", end - 2)) {
      let j = end + 1;
      for (;;) {
        j = source.indexOf("EI", j);
        if (j < 0) return null;
        const before = source.charCodeAt(j - 1);
        const after = j + 2 >= source.length ? 0x20 : source.charCodeAt(j + 2);
        if (WHITE.has(before) && WHITE.has(after)) break;
        j += 2;
      }
      tokens.push({ start: j, end: j + 2, operator: true });
      i = j + 2;
    }
  }
}

function numberAt(source: string, token: Token): number {
  const text = source.slice(token.start, token.end);
  if (!NUMBER.test(text)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
  const value = Number(text);
  if (!Number.isFinite(value)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
  return value;
}

/** The bytes a string token writes (mirror of _string_bytes); null when it
 * is not one this reader decodes exactly. */
function stringBytes(source: string, token: Token): number[] | null {
  const text = source.slice(token.start, token.end);
  const out: number[] = [];
  if (text.startsWith("<")) {
    const digits = text.slice(1, -1).replace(/[\x00\t\n\f\r ]/g, "");
    if (!/^[0-9A-Fa-f]*$/.test(digits)) return null;
    const even = digits.length % 2 === 0 ? digits : `${digits}0`;
    for (let index = 0; index < even.length; index += 2) out.push(parseInt(even.slice(index, index + 2), 16));
    return out;
  }
  if (!text.startsWith("(")) return null;
  const body = text.slice(1, -1);
  for (let index = 0; index < body.length; index += 1) {
    const c = body.charCodeAt(index);
    if (c === 0x0d) return null;
    if (c !== 0x5c) {
      out.push(c & 0xff);
      continue;
    }
    index += 1;
    if (index >= body.length) return null;
    const next = body[index];
    const simple: Record<string, number> = { n: 0x0a, r: 0x0d, t: 0x09, b: 0x08, f: 0x0c, "(": 0x28, ")": 0x29, "\\": 0x5c };
    if (next in simple) {
      out.push(simple[next]);
    } else if (next >= "0" && next <= "7") {
      let digits = next;
      while (digits.length < 3 && index + 1 < body.length && body[index + 1] >= "0" && body[index + 1] <= "7") {
        index += 1;
        digits += body[index];
      }
      out.push(parseInt(digits, 8) & 0xff);
    } else if (next === "\n") {
      // A line continuation writes nothing.
    } else if (next === "\r") {
      return null;
    } else {
      out.push(next.charCodeAt(0) & 0xff);
    }
  }
  return out;
}

function hexString(bytes: readonly number[]): string {
  return `<${bytes.map((byte) => byte.toString(16).toUpperCase().padStart(2, "0")).join("")}>`;
}

/**
 * A show's elements cut at a range of its code bytes (mirror of
 * _split_show_payload): the elements before it, in it and after it, strings
 * written as hex, numbers as they were written. A number between two pieces
 * goes with the piece its pen position is in.
 */
function splitShowPayload(
  source: string,
  operand: Token,
  array: boolean,
  part: readonly [number, number],
): Readonly<{ before: string; owned: string; after: string }> {
  const elements: Token[] = [];
  if (array) {
    let index = operand.start + 1;
    for (;;) {
      index = skipWhite(source, index);
      if (index >= operand.end - 1) break;
      const end = tokenEnd(source, index);
      if (end < 0 || end > operand.end - 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      elements.push({ start: index, end, operator: false });
      index = end;
    }
  } else {
    elements.push(operand);
  }
  const [cut, stop] = part;
  const pieces: [string[], string[], string[]] = [[], [], []];
  let position = 0;
  const pieceAt = (at: number) => (at < cut ? 0 : at < stop ? 1 : 2);
  for (const element of elements) {
    const first = source.charCodeAt(element.start);
    if (first === 0x28 || first === 0x3c) {
      const bytes = stringBytes(source, element);
      if (!bytes) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      let from = 0;
      while (from < bytes.length) {
        const piece = pieceAt(position + from);
        const limit = piece === 0 ? cut : piece === 1 ? stop : Number.POSITIVE_INFINITY;
        const to = Math.min(bytes.length, limit - position);
        pieces[piece].push(hexString(bytes.slice(from, to)));
        from = to;
      }
      position += bytes.length;
    } else {
      numberAt(source, element);
      pieces[pieceAt(position)].push(source.slice(element.start, element.end));
    }
  }
  if (!(cut < stop) || stop > position || pieces[1].length === 0) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.partial);
  return Object.freeze({ before: pieces[0].join(" "), owned: pieces[1].join(" "), after: pieces[2].join(" ") });
}

// ---------------------------------------------------------------------------
// Matrices (row vectors, as the PDF writes them) and numbers.

function multiply(left: Matrix, right: Matrix): Matrix {
  const [a, b, c, d, e, f] = left;
  const [ta, tb, tc, td, te, tf] = right;
  return [
    a * ta + b * tc,
    a * tb + b * td,
    c * ta + d * tc,
    c * tb + d * td,
    e * ta + f * tc + te,
    e * tb + f * td + tf,
  ];
}

function translate(tx: number, ty: number, matrix: Matrix): Matrix {
  return [
    matrix[0],
    matrix[1],
    matrix[2],
    matrix[3],
    tx * matrix[0] + ty * matrix[2] + matrix[4],
    tx * matrix[1] + ty * matrix[3] + matrix[5],
  ];
}

/** The map read in a space drawn under `ctm`: CTM x M x CTM^-1. A move keeps
 * its linear part exactly the identity, so a relative operator of the box
 * keeps its bytes and its translation is the text move's offset. */
function groupMapUnder(ctm: Matrix, transform: PdfNativeTextGroupTransform): Matrix {
  const [a, b, c, d, e, f] = ctm;
  const determinant = a * d - b * c;
  if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) {
    refuse("Native text move requires an invertible graphics state.");
  }
  if (groupMapMovesOnly(transform)) {
    // Solve [x y] x CTM_linear = [e f] (mirror of the text move's offset).
    return [
      1,
      0,
      0,
      1,
      (transform.e * d - transform.f * c) / determinant,
      (transform.f * a - transform.e * b) / determinant,
    ];
  }
  const inverse: Matrix = [
    d / determinant,
    -b / determinant,
    -c / determinant,
    a / determinant,
    (c * f - d * e) / determinant,
    (b * e - a * f) / determinant,
  ];
  const [ta, tb, tc, td] = groupLinear(transform);
  return multiply(multiply(ctm, [ta, tb, tc, td, transform.e, transform.f]), inverse);
}

function micro(value: number): number {
  const result = Math.floor(value * 1_000_000 + 0.5);
  if (!Number.isSafeInteger(result) || Math.abs(result) > MAX_GROUP_MICRO) {
    refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.transform);
  }
  return result === 0 ? 0 : result;
}

/** A value written as the rewrite writes it, and the value a reader reads back. */
function written(value: number): Readonly<{ text: string; value: number }> {
  const m = micro(value);
  return { text: formatMraiNativeTextMoveMicro(m), value: m / 1_000_000 };
}

// ---------------------------------------------------------------------------
// Graphics state (mirror of _GroupState).

export type PdfNativeTextGroupRect = readonly [number, number, number, number];
type Rect = PdfNativeTextGroupRect;
/** null: no clip; else the rectangle in the page's user space the clip is
 * (a rectangle) or lies inside (any other path: its bounds). */
export type PdfNativeTextGroupClip = Rect | null;
type Clip = PdfNativeTextGroupClip;

interface GraphicsState {
  ctm: Matrix;
  leading: number;
  mode: number;
  clip: Clip;
  /** Tf's size, Tz, and Ts with the text it was written as. */
  fontSize: number;
  hscale: number;
  rise: number;
  riseText: string;
  /** Every state operator in force, the last writer of each key, in the
   * order the page wrote them: written again in that order, they set the
   * state the walk reads here (mirror of _State.entries). */
  entries: readonly StateEntry[];
  /** The font resource name Tf last set ("" before any). */
  font: string;
}

/** One state operator in force: what it sets (its key) and how it is written. */
interface StateEntry {
  readonly key: string;
  readonly text: string;
}

const FILL = "fill";
const FILL_COLOR = "fill color";
const STROKE = "stroke";
const STROKE_COLOR = "stroke color";

/** What a state key is set back to where the run never set it (the initial
 * graphics state). A font has none: a run that shows text sets its own. */
const STATE_DEFAULTS: Readonly<Record<string, string>> = Object.freeze({
  w: "1 w",
  J: "0 J",
  j: "0 j",
  M: "10 M",
  d: "[] 0 d",
  ri: "/RelativeColorimetric ri",
  i: "0 i",
  [FILL]: "0 g",
  [STROKE]: "0 G",
  Tc: "0 Tc",
  Tw: "0 Tw",
  Tz: "100 Tz",
  TL: "0 TL",
  Tr: "0 Tr",
  Ts: "0 Ts",
});

/** The marked-content operators: a block drawn again outside the clip
 * leaves them out (its glyphs are drawn once more, not tagged twice). */
const MARKED = new Set(["BMC", "BDC", "EMC", "MP", "DP"]);

function operatorText(source: string, operator: string, operands: readonly Token[]): string {
  return [...operands.map((token) => source.slice(token.start, token.end)), operator].join(" ");
}

/** The state entry an operator writes (and the key it clears: a colour
 * space sets its colour back), or null for any other operator (mirror of
 * _state_entry). */
function stateEntry(
  source: string,
  operator: string,
  operands: readonly Token[],
): Readonly<{ key: string; text: string; clears?: string }> | null {
  switch (operator) {
    case "w":
    case "J":
    case "j":
    case "M":
    case "d":
    case "ri":
    case "i":
    case "Tc":
    case "Tw":
    case "Tz":
    case "TL":
    case "Tf":
    case "Tr":
    case "Ts":
      return { key: operator, text: operatorText(source, operator, operands) };
    case "gs":
      if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return { key: `gs ${source.slice(operands[0].start, operands[0].end)}`, text: operatorText(source, operator, operands) };
    case "g":
    case "rg":
    case "k":
    case "cs":
      return { key: FILL, text: operatorText(source, operator, operands), clears: FILL_COLOR };
    case "sc":
    case "scn":
      return { key: FILL_COLOR, text: operatorText(source, operator, operands) };
    case "G":
    case "RG":
    case "K":
    case "CS":
      return { key: STROKE, text: operatorText(source, operator, operands), clears: STROKE_COLOR };
    case "SC":
    case "SCN":
      return { key: STROKE_COLOR, text: operatorText(source, operator, operands) };
    default:
      return null;
  }
}

function withEntry(
  entries: readonly StateEntry[],
  entry: Readonly<{ key: string; text: string; clears?: string }>,
): readonly StateEntry[] {
  return Object.freeze([
    ...entries.filter((item) => item.key !== entry.key && item.key !== entry.clears),
    Object.freeze({ key: entry.key, text: entry.text }),
  ]);
}

/** A colour set by a pattern (its last operand a name) is placed in the
 * page's space, not the text's: it is not carried out of a clip. */
function isPatternColor(text: string): boolean {
  return /\/[^\s/]*\s+(?:scn|SCN)$/.test(text);
}

// ---------------------------------------------------------------------------
// A synthetic bold of the box's shows (mirror of _bold_wrap).

/** The stroke operator that sets the colour a fill operator sets. */
const STROKE_SPELLING: Readonly<Record<string, string>> = Object.freeze({
  g: "G", rg: "RG", k: "K", cs: "CS", sc: "SC", scn: "SCN",
});

/** A fill state entry spelt as the stroke one that sets the same colour
 * (`0 0 0 rg` is `0 0 0 RG`, `/Cs1 cs` is `/Cs1 CS`). */
function strokeSpelling(text: string): string {
  const at = text.lastIndexOf(" ");
  const stroke = STROKE_SPELLING[text.slice(at + 1)];
  if (!stroke) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
  return `${text.slice(0, at + 1)}${stroke}`;
}

/**
 * The stroke width in force, as the page wrote it: the last `w` (else the
 * initial `1 w`), unless a `gs` after it sets one -- its `/LW`, as the state
 * spells it (`widths`, by the name the `gs` gives). undefined when a `gs`
 * after the last width is one the page's resources do not state plainly.
 */
function lineWidthInForce(
  entries: readonly StateEntry[],
  widths: ReadonlyMap<string, string | null> | undefined,
): string | undefined {
  let text = "1 w";
  let known = true;
  for (const entry of entries) {
    if (entry.key === "w") {
      text = entry.text;
      known = true;
    } else if (entry.key.startsWith("gs ")) {
      const width = widths?.get(entry.key.slice(3));
      if (width === undefined) known = false;
      else if (width !== null) {
        text = `${width} w`;
        known = true;
      }
    }
  }
  return known ? text : undefined;
}

/**
 * What a group bold writes around one show (mirror of _bold_wrap), or null
 * when the show is drawn that way already. On: `2 Tr`, a width of 0.035 em
 * of the show (its size times its line matrix's x axis, so the width scales
 * with the glyphs), the stroke colour set to its fill; after it, the mode,
 * width and stroke colour in force put back, each only where it differs.
 * Off: `0 Tr`, the `2 Tr` put back after it.
 */
function boldWrap(
  on: boolean,
  mode: number,
  entries: readonly StateEntry[],
  fontSize: number,
  linear: Matrix,
  widths: ReadonlyMap<string, string | null> | undefined,
): Readonly<{ pre: string; post: string }> | null {
  if (mode !== 0 && mode !== 2) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldMode);
  if (!on) return mode === 2 ? Object.freeze({ pre: "0 Tr", post: "2 Tr" }) : null;
  if (mode === 2) return null;
  const inForce = lineWidthInForce(entries, widths);
  if (inForce === undefined) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldWidth);
  const em = Math.abs(fontSize) * Math.sqrt(linear[0] * linear[0] + linear[1] * linear[1]);
  const width = written(PDF_NATIVE_SYNTHETIC_BOLD_STROKE_EM * em);
  if (!(width.value > 0)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldWidth);
  const read = (key: string) => entries.find((entry) => entry.key === key)?.text;
  const fill = read(FILL);
  const fillColor = read(FILL_COLOR);
  const stroke = read(STROKE);
  const strokeColor = read(STROKE_COLOR);
  const wanted = [fill === undefined ? "0 G" : strokeSpelling(fill),
    ...(fillColor === undefined ? [] : [strokeSpelling(fillColor)])].join(" ");
  const standing = [stroke ?? "0 G", ...(strokeColor === undefined ? [] : [strokeColor])].join(" ");
  const widthText = `${width.text} w`;
  const pre = ["2 Tr", ...(widthText === inForce ? [] : [widthText]), ...(wanted === standing ? [] : [wanted])];
  const post = ["0 Tr", ...(widthText === inForce ? [] : [inForce]), ...(wanted === standing ? [] : [standing])];
  return Object.freeze({ pre: pre.join(" "), post: post.join(" ") });
}

/** The `gs` names a stream and the streams before it apply, as each `gs`
 * spells its operand (mirror of _ext_g_state_names). */
export function mraiNativeTextGroupExtGStateNames(streams: readonly string[]): readonly string[] {
  const names = new Set<string>();
  for (const stream of streams) {
    const tokens = groupTokens(stream);
    if (!tokens) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
    tokens.forEach((token, index) => {
      if (token.operator && stream.slice(token.start, token.end) === "gs" && index > 0 && !tokens[index - 1].operator) {
        names.add(stream.slice(tokens[index - 1].start, tokens[index - 1].end));
      }
    });
  }
  return Object.freeze([...names].sort());
}

// PDF white space only, as the replay reads it (mirror of _LINE_WIDTH_KEY / _VALUE).
const LINE_WIDTH_KEY = /\/LW(?![^\x00\t\n\f\r /<>[\]()%{}])/g;
const LINE_WIDTH_VALUE =
  /^[\x00\t\n\f\r ]*([+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+))(?=$|[\x00\t\n\f\r /<>[\]()%{}])(?![\x00\t\n\f\r ]+[0-9]+[\x00\t\n\f\r ]+R(?![^\x00\t\n\f\r /<>[\]()%{}]))/;

/** The `/LW` an ExtGState dictionary sets, as it spells it: null when it
 * sets none, undefined when it sets one this reader does not read plainly
 * (twice, by reference, or not a number) (mirror of _ext_g_state_line_width). */
export function mraiNativeExtGStateLineWidth(dictionary: string): string | null | undefined {
  const read = pdfSourceNamesRead(dictionary);
  const keys = [...read.matchAll(LINE_WIDTH_KEY)];
  if (keys.length === 0) return null;
  if (keys.length > 1) return undefined;
  const value = LINE_WIDTH_VALUE.exec(read.slice(keys[0].index! + 3));
  return value ? value[1] : undefined;
}

/** The stroke widths the `gs` names of a page set, by the name each `gs`
 * gives (see lineWidthInForce); a name the page's resources do not hold as
 * one plain dictionary is left out. */
/** Whether each show is drawn in a bold face, by the BaseFont of the font
 * it names (pdfNativeTextFontIdentity, as the editor names a face bold). */
async function readGroupBoldFaces(
  sourceBytes: Uint8Array,
  pageNumber: number,
  fonts: readonly string[],
): Promise<readonly boolean[]> {
  const names = [...new Set(fonts.filter(Boolean))];
  const dictionaries = await readMraiNativePageResourceDictionaries(
    sourceBytes, pageNumber, "Font", names.map((name) => decodePdfSourceName(name)));
  const bold = new Map<string, boolean>();
  for (const name of names) {
    const dictionary = dictionaries.get(decodePdfSourceName(name));
    const baseFont = dictionary ? /\/BaseFont\s*(\/[^\s/<>[\]()%{}]+)/u.exec(dictionary)?.[1] : undefined;
    bold.set(name, baseFont ? pdfNativeTextFontIdentity(decodePdfSourceName(baseFont)).bold : false);
  }
  return fonts.map((font) => bold.get(font) ?? false);
}

async function readGroupLineWidths(
  sourceBytes: Uint8Array,
  pageNumber: number,
  names: readonly string[],
): Promise<ReadonlyMap<string, string | null>> {
  const dictionaries = await readMraiNativePageResourceDictionaries(
    sourceBytes, pageNumber, "ExtGState", names.map((name) => decodePdfSourceName(name)));
  const widths = new Map<string, string | null>();
  for (const name of names) {
    const dictionary = dictionaries.get(decodePdfSourceName(name));
    if (!dictionary) continue;
    const width = mraiNativeExtGStateLineWidth(dictionary);
    if (width !== undefined) widths.set(name, width);
  }
  return widths;
}

/** `TD`'s second operand negated as written: the leading it sets. */
function negatedText(text: string): string {
  return text.startsWith("-") ? text.slice(1) : text.startsWith("+") ? `-${text.slice(1)}` : `-${text}`;
}

/** Whether a marked-content operator opens a sequence a copy cannot carry:
 * properties named in the page's resources (a layer, or anything else) or
 * a replacement text of its glyphs. */
function marksBlocked(source: string, operands: readonly Token[]): boolean {
  if (operands.length !== 2) return true;
  const properties = source.slice(operands[1].start, operands[1].end);
  // Keys are read as their values: `/Actual#54ext` is /ActualText.
  return !properties.startsWith("<<") || /\/ActualText(?![^\s/<>[\]()%{}])/.test(pdfSourceNamesRead(properties));
}

/** The marked-content sequences open at a point of the walk: whether each
 * is one a copy cannot carry (mirror of _step_marks). */
function stepMarks(source: string, operator: string, operands: readonly Token[], marks: boolean[]): void {
  if (operator === "BMC") marks.push(false);
  else if (operator === "BDC") marks.push(marksBlocked(source, operands));
  else if (operator === "EMC") marks.pop();
}

function intersectClip(current: Clip, next: Clip): Clip {
  if (next === null) return current;
  if (current === null) return next;
  return [
    Math.max(current[0], next[0]),
    Math.max(current[1], next[1]),
    Math.min(current[2], next[2]),
    Math.min(current[3], next[3]),
  ];
}

function applyPoint(matrix: Matrix, x: number, y: number): readonly [number, number] {
  return [x * matrix[0] + y * matrix[2] + matrix[4], x * matrix[1] + y * matrix[3] + matrix[5]];
}

/** The axis-aligned rectangle a closed polygon of these points draws, if it is one. */
function polygonRect(points: readonly (readonly [number, number])[]): Rect | null {
  let list = points;
  if (list.length === 5 && list[4][0] === list[0][0] && list[4][1] === list[0][1]) list = list.slice(0, 4);
  if (list.length !== 4) return null;
  for (let index = 0; index < 4; index += 1) {
    const [x0, y0] = list[index];
    const [x1, y1] = list[(index + 1) % 4];
    if (x0 !== x1 && y0 !== y1) return null;
  }
  const xs = list.map((point) => point[0]);
  const ys = list.map((point) => point[1]);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

interface PathState {
  /** The path's subpaths in the page's user space; a curve adds its control
   * points, which bound it. */
  subpaths: (readonly [number, number])[][];
  curved: boolean;
  pendingClip: boolean;
}

const PAINT = new Set(["n", "f", "F", "f*", "B", "B*", "b", "b*", "S", "s"]);

/** One walk's handling of an operator outside the text rewrite: the graphics
 * state, the state operators in force and the path a clip is read from. */
function stepGraphics(
  source: string,
  operator: string,
  operands: readonly Token[],
  state: GraphicsState,
  stack: GraphicsState[],
  path: PathState,
): GraphicsState {
  const next = stepGraphicsState(source, operator, operands, state, stack, path);
  const entry = operator === "TD" && operands.length === 2
    ? { key: "TL", text: `${negatedText(source.slice(operands[1].start, operands[1].end))} TL` }
    : stateEntry(source, operator, operands);
  if (!entry) return next;
  return {
    ...next,
    entries: withEntry(next.entries, entry),
    ...(operator === "Tf" ? { font: source.slice(operands[0].start, operands[0].end) } : {}),
  };
}

function stepGraphicsState(
  source: string,
  operator: string,
  operands: readonly Token[],
  state: GraphicsState,
  stack: GraphicsState[],
  path: PathState,
): GraphicsState {
  switch (operator) {
    case "q":
      stack.push({ ...state });
      return state;
    case "Q": {
      const restored = stack.pop();
      if (!restored) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return { ...restored };
    }
    case "cm": {
      if (operands.length !== 6) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      const values = operands.map((token) => numberAt(source, token)) as unknown as Matrix;
      return { ...state, ctm: multiply(values, state.ctm) };
    }
    case "TL":
      if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return { ...state, leading: numberAt(source, operands[0]) };
    case "Tf":
      if (operands.length !== 2) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return { ...state, fontSize: numberAt(source, operands[1]) };
    case "Tz":
      if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return { ...state, hscale: numberAt(source, operands[0]) };
    case "Ts":
      if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return {
        ...state,
        rise: numberAt(source, operands[0]),
        riseText: source.slice(operands[0].start, operands[0].end),
      };
    case "TD":
      if (operands.length !== 2) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return { ...state, leading: -numberAt(source, operands[1]) };
    case "Tr": {
      if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      const mode = numberAt(source, operands[0]);
      if (!Number.isInteger(mode) || mode < 0 || mode > 7) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      return { ...state, mode };
    }
    case "re": {
      if (operands.length !== 4) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      const [x, y, w, h] = operands.map((token) => numberAt(source, token));
      path.subpaths.push([
        applyPoint(state.ctm, x, y),
        applyPoint(state.ctm, x + w, y),
        applyPoint(state.ctm, x + w, y + h),
        applyPoint(state.ctm, x, y + h),
      ]);
      return state;
    }
    case "m":
      if (operands.length !== 2) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      path.subpaths.push([applyPoint(state.ctm, numberAt(source, operands[0]), numberAt(source, operands[1]))]);
      return state;
    case "l":
    case "c":
    case "v":
    case "y": {
      const count = operator === "l" ? 2 : operator === "c" ? 6 : 4;
      if (operands.length !== count) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      if (path.subpaths.length === 0) path.subpaths.push([]);
      const current = path.subpaths[path.subpaths.length - 1];
      for (let index = 0; index < count; index += 2) {
        current.push(applyPoint(state.ctm, numberAt(source, operands[index]), numberAt(source, operands[index + 1])));
      }
      if (operator !== "l") path.curved = true;
      return state;
    }
    case "h":
      return state;
    case "W":
    case "W*":
      path.pendingClip = true;
      return state;
    default:
      if (PAINT.has(operator)) {
        let next = state;
        if (path.pendingClip) {
          // A rectangle clips exactly to itself; any other path to at most its
          // bounds, which the landing is then held to.
          const points = path.subpaths.flat();
          const rect = !path.curved && path.subpaths.length === 1 ? polygonRect(path.subpaths[0]) : null;
          const bounds: Rect | null = points.length === 0 ? null : [
            Math.min(...points.map((point) => point[0])), Math.min(...points.map((point) => point[1])),
            Math.max(...points.map((point) => point[0])), Math.max(...points.map((point) => point[1])),
          ];
          const clip = rect ?? bounds;
          if (clip) next = { ...state, clip: intersectClip(state.clip, clip) };
        }
        path.subpaths = [];
        path.curved = false;
        path.pendingClip = false;
        return next;
      }
      return state;
  }
}

// ---------------------------------------------------------------------------
// The rewrite (mirror of derive_text_group_move).

interface PositionEvent {
  readonly kind: "Tm" | "Td" | "TD" | "T*" | "'" | '"';
  /** From its first operand to the end of its operator. */
  readonly start: number;
  readonly end: number;
  readonly operands: readonly Token[];
  /** The line matrix it sets in the source. */
  readonly lineMatrix: Matrix;
  /** The leading in force when it runs (T*, ', "). */
  readonly leading: number;
}

/** A show of the box, with the text state it is drawn in. */
interface OwnedShow {
  readonly start: number;
  readonly end: number;
  /** The byte range of its codes the box owns, when it owns only part. */
  readonly part?: readonly [number, number];
  readonly kind: string;
  readonly operands: readonly Token[];
  readonly fontSize: number;
  readonly hscale: number;
  readonly rise: number;
  readonly riseText: string;
}

interface Segment {
  readonly opener: PositionEvent | null;
  owned: number;
  other: number;
  flag: boolean | null;
  /** Box shows and other text on one line: the box's shows move by their
   * own kerning and rise, the line stays placed as it was (a move only). */
  mixed: boolean;
  readonly shows: OwnedShow[];
  /** The state operators in force once its opener has run. */
  entries: readonly StateEntry[];
}

/** One operator of a text object, as the block outside the clip copies it. */
interface ObjectOperator {
  readonly op: string;
  readonly start: number;
  readonly end: number;
  readonly operands: readonly Token[];
  /** The segment it is part of (an opener: the one it opens). */
  readonly segment: number;
  readonly opens: boolean;
  readonly owned: boolean;
  /** The font a show is drawn in. */
  readonly font: string;
}

/** A run of a box's segments drawn again outside the clip: one text object's
 * consecutive segments whose shows are all the box's. */
interface OutsideRun {
  readonly ctm: Matrix;
  readonly entries: readonly StateEntry[];
  readonly lines: readonly string[];
}

/** A rectangle filled outside a text object: a rule another writer may have
 * drawn under a run (its `re` operands, where they and its fill end). */
interface RuleCandidate {
  readonly start: number;
  readonly reEnd: number;
  readonly end: number;
  readonly numbers: readonly number[];
  readonly ctm: Matrix;
  readonly fill: string | null;
  readonly disruptions: number;
}

/** One of this writer's rule blocks (pdfNativeTextRuleBlock: its colour and
 * four corners) with its corners taken through `map`, written as it writes
 * one; null where it is not one. */
function movedRuleBlock(block: string, map: Matrix): string | null {
  const number = String.raw`-?\d+(?:\.\d+)?`;
  const match = new RegExp(String.raw`^\nq ((?:${number} ){1,4})(g|rg|k) (${number}) (${number}) m ` +
    String.raw`(${number}) (${number}) l (${number}) (${number}) l (${number}) (${number}) l h f Q\n$`, "u").exec(block);
  if (!match) return null;
  const corners: string[] = [];
  for (let at = 3; at < 11; at += 2) {
    const x = Number(match[at]), y = Number(match[at + 1]);
    corners.push(`${written(map[0] * x + map[2] * y + map[4]).text} ${written(map[1] * x + map[3] * y + map[5]).text}`);
  }
  return `\nq ${match[1]}${match[2]} ${corners[0]} m ${corners[1]} l ${corners[2]} l ${corners[3]} l h f Q\n`;
}

/** Colours, marked content and the rule's own path: what stands between a
 * box's runs and the rules drawn beside them. */
const RULE_NEIGHBOUR_OPERATORS = new Set(["G", "RG", "K", "BDC", "BMC", "EMC", "MP", "DP"]);

/**
 * The rules a box's runs carry (a web print's link underlines): thin
 * horizontal rectangles filled in a colour of their own, under the matrix
 * the box's text is drawn under, among its text objects or right after the
 * last of them (nothing but colours and marked content before them).
 */
function carriedGroupRules(
  candidates: readonly RuleCandidate[],
  owned: readonly Readonly<{ start: number; end: number; ctm: Matrix; disruptions: number }>[],
): readonly RuleCandidate[] {
  const first = owned[0];
  const last = owned.at(-1);
  if (!first || !last) return [];
  const same = (left: Matrix, right: Matrix) => left.every((value, at) => Math.abs(value - right[at]) <= 1e-9);
  if (!owned.every((object) => same(object.ctm, first.ctm))) return [];
  return candidates.filter((rule) => {
    if (rule.start < first.end || !rule.fill || !same(rule.ctm, first.ctm)) return false;
    if (rule.start > last.end && rule.disruptions !== last.disruptions) return false;
    const [a, b, c, d] = rule.ctm;
    const [x, y, w, h] = rule.numbers;
    const corners = [[x, y], [x + w, y], [x, y + h], [x + w, y + h]].map(([px, py]) => [a * px + c * py, b * px + d * py]);
    const xs = corners.map(([px]) => px), ys = corners.map(([, py]) => py);
    const across = Math.max(...ys) - Math.min(...ys), along = Math.max(...xs) - Math.min(...xs);
    return across > 0 && across <= 2.5 && along > 3 * across;
  });
}

/** The rewrite of one text object, and what the clip check needs of it. */
interface ObjectRewrite {
  readonly edits: { start: number; end: number; text: string }[];
  readonly clip: Clip;
}

export interface PdfNativeTextGroupMoveRewrite {
  readonly span: Span;
  readonly replacement: string;
  /** The clip in force at each text object the box is drawn in (outside
   * every clip: the one in force at the tail of the stream). */
  readonly clips: readonly Clip[];
  /** Written outside every clip: the replacement is `middle` (the text
   * objects with the box's shows taken out, up to `middleEnd`), the source
   * from there to its end, then `block`. */
  readonly outside?: Readonly<{
    middle: string;
    middleEnd: number;
    block: string;
    /** Font and ExtGState resource names the block draws with (`/F1`). */
    fonts: readonly string[];
    extGStates: readonly string[];
  }>;
  /** The rendering mode each show of the box is drawn in, in stream order. */
  readonly modes: readonly number[];
  /** The font resource each show of the box is drawn in (`/F1`), in stream
   * order. */
  readonly showFonts: readonly string[];
}

/** A show a group bold wraps: where it is, what it is, the range of its
 * codes it is set on (a part), and what is written around it. */
interface BoldWrap {
  readonly start: number;
  readonly end: number;
  readonly kind: string;
  readonly operands: readonly Token[];
  readonly range?: readonly [number, number];
  readonly pre: string;
  readonly post: string;
}

/** A wrapped show as an edit of its own: the show (or the part of it the
 * bold is set on, the rest drawn on either side as it was) between the
 * state the bold sets and the state put back (mirror of _bold_wrapped). */
function boldWrapped(source: string, wrap: BoldWrap): string {
  if (!wrap.range) return ` ${wrap.pre} ${source.slice(wrap.start, wrap.end)} ${wrap.post}`;
  if (wrap.kind !== "Tj" && wrap.kind !== "TJ") refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
  const pieces = splitShowPayload(source, wrap.operands[0], wrap.kind === "TJ", wrap.range);
  return ` ${[
    ...(pieces.before ? [`[${pieces.before}] TJ`] : []),
    `${wrap.pre} [${pieces.owned}] TJ ${wrap.post}`,
    ...(pieces.after ? [`[${pieces.after}] TJ`] : []),
  ].join(" ")}`;
}

function formatMatrix(values: readonly Readonly<{ text: string }>[]): string {
  return values.map((value) => value.text).join(" ");
}

/**
 * The span of `source` a group move of `shows` by `transform` rewrites and
 * what it writes there. `prefix` are the page's streams drawn before it.
 * `outsideClip` writes the box outside every clip (the stream must be the
 * page's last: the block it appends closes what the stream leaves open).
 */
export function deriveMraiNativeTextGroupMoveRewrite(input: {
  readonly prefix: readonly string[];
  readonly source: string;
  readonly shows: readonly GroupShow[];
  readonly transform: PdfNativeTextGroupTransform;
  readonly outsideClip?: boolean;
  /** A synthetic bold put on or taken off with the map. */
  readonly bold?: PdfNativeTextGroupBold;
  /** The stroke widths the page's `gs` names set (lineWidthInForce). */
  readonly lineWidths?: ReadonlyMap<string, string | null>;
  /** The page draws on in a stream after this one. */
  readonly continued?: boolean;
}): PdfNativeTextGroupMoveRewrite {
  const { source, transform, bold } = input;
  const outsideClip = input.outsideClip === true;
  const owned = new Set(input.shows.map((show) => `${show.start}:${show.end}`));
  const parts = new Map(input.shows.filter((show) => show.codeStart !== undefined)
    .map((show) => [`${show.start}:${show.end}`, [show.codeStart!, show.codeEnd!] as const]));
  // The shows the bold is set on (every show of the box when it names none),
  // each with the range of its codes when it is a part.
  const boldShows = bold?.shows
    ? new Map(bold.shows.map((show) => [`${show.start}:${show.end}`, show]))
    : null;
  if (boldShows && (boldShows.size !== bold!.shows!.length || boldShows.size === 0)) {
    invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
  }
  const scaled = !groupMapMovesOnly(transform);
  const modes: number[] = [];
  const showFonts: string[] = [];
  let boldFound = 0;
  let state: GraphicsState = {
    ctm: [1, 0, 0, 1, 0, 0], leading: 0, mode: 0, clip: null, fontSize: 1, hscale: 100, rise: 0, riseText: "0",
    entries: Object.freeze([]), font: "",
  };
  const stack: GraphicsState[] = [];
  const path: PathState = { subpaths: [], curved: false, pendingClip: false };
  const marks: boolean[] = [];
  // The streams drawn before this one only set the state it starts in.
  for (const earlier of input.prefix) {
    const tokens = groupTokens(earlier);
    if (!tokens) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
    let operands: Token[] = [];
    for (const token of tokens) {
      if (!token.operator) {
        operands.push(token);
        continue;
      }
      const operator = earlier.slice(token.start, token.end);
      state = stepGraphics(earlier, operator, operands, state, stack, path);
      stepMarks(earlier, operator, operands, marks);
      operands = [];
    }
  }
  const tokens = groupTokens(source);
  if (!tokens) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
  const objects: ObjectRewrite[] = [];
  const runs: OutsideRun[] = [];
  // The rules another writer drew among the box's text objects (a web
  // print's link underlines: `x y w h re f` beside each run, under the same
  // matrix), read as the walk passes them; carried with the box below.
  const ruleCandidates: RuleCandidate[] = [];
  let pendingRect: Omit<RuleCandidate, "end" | "fill" | "disruptions"> | null = null;
  let fillOperator: string | null = null;
  // Operators other than colours, marked content and rules: a rule after the
  // box's last text object with one of these before it is not the box's.
  let disruptions = 0;
  const ownedObjects: { start: number; end: number; ctm: Matrix; clip: Clip; disruptions: number;
    runs: readonly [number, number] }[] = [];
  const fonts = new Set<string>();
  const extGStates = new Set<string>();
  let found = 0;
  let operands: Token[] = [];
  let index = 0;
  while (index < tokens.length) {
    const token = tokens[index];
    if (!token.operator) {
      operands.push(token);
      index += 1;
      continue;
    }
    const operator = source.slice(token.start, token.end);
    if (operator !== "BT") {
      const numbers = operands.map((operand) => Number(source.slice(operand.start, operand.end)));
      if (operator === "re" && numbers.length === 4 && numbers.every(Number.isFinite)) {
        pendingRect = { start: operands[0].start, reEnd: token.end, numbers, ctm: state.ctm };
      } else {
        if ((operator === "f" || operator === "F" || operator === "f*") && pendingRect) {
          ruleCandidates.push({ ...pendingRect, end: token.end, fill: fillOperator, disruptions });
        } else if ((operator === "g" || operator === "rg" || operator === "k") && operands.length) {
          fillOperator = source.slice(operands[0].start, token.end);
        } else if (!RULE_NEIGHBOUR_OPERATORS.has(operator)) {
          disruptions += 1;
        }
        pendingRect = null;
      }
      state = stepGraphics(source, operator, operands, state, stack, path);
      stepMarks(source, operator, operands, marks);
      operands = [];
      index += 1;
      continue;
    }
    pendingRect = null;
    // One text object: its events, read from the source.
    const bt = token;
    const objectState = state;
    let leading = state.leading;
    let mode = state.mode;
    let lineMatrix: Matrix = [1, 0, 0, 1, 0, 0];
    const segments: Segment[] = [
      { opener: null, owned: 0, other: 0, flag: null, mixed: false, shows: [], entries: state.entries },
    ];
    const objectOperators: ObjectOperator[] = [];
    let fontSize = state.fontSize;
    let hscale = state.hscale;
    let rise = state.rise;
    let riseText = state.riseText;
    let entries = state.entries;
    let font = state.font;
    let ownedHere = 0;
    // The shows of this object a bold wraps, and the map read in its space.
    const wraps: BoldWrap[] = [];
    let objectMap: Matrix | null = null;
    operands = [];
    index += 1;
    let closed = false;
    for (; index < tokens.length; index += 1) {
      const inner = tokens[index];
      if (!inner.operator) {
        operands.push(inner);
        continue;
      }
      const op = source.slice(inner.start, inner.end);
      const start = operands.length > 0 ? operands[0].start : inner.start;
      const segmentCount = segments.length;
      let ownedShow = false;
      const event = (kind: PositionEvent["kind"], matrix: Matrix): void => {
        lineMatrix = matrix;
        segments.push({
          opener: Object.freeze({ kind, start, end: inner.end, operands: [...operands], lineMatrix: matrix, leading }),
          owned: 0,
          other: 0,
          flag: null,
          mixed: false,
          shows: [],
          entries,
        });
      };
      if (op === "ET") {
        closed = true;
        index += 1;
        break;
      }
      // The state operators in force, kept for a copy drawn outside the clip.
      const entry = op === "TD" && operands.length === 2
        ? { key: "TL", text: `${negatedText(source.slice(operands[1].start, operands[1].end))} TL` }
        : stateEntry(source, op, operands);
      if (entry) entries = withEntry(entries, entry);
      if (op === '"' && operands.length === 3) {
        entries = withEntry(entries, { key: "Tw", text: `${source.slice(operands[0].start, operands[0].end)} Tw` });
        entries = withEntry(entries, { key: "Tc", text: `${source.slice(operands[1].start, operands[1].end)} Tc` });
      }
      if (op === "Tf" && operands.length === 2) font = source.slice(operands[0].start, operands[0].end);
      stepMarks(source, op, operands, marks);
      if (op === "BT" || op === "q" || op === "Q" || op === "cm") {
        refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      } else if (op === "Tm") {
        if (operands.length !== 6) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        event("Tm", operands.map((operand) => numberAt(source, operand)) as unknown as Matrix);
      } else if (op === "Td" || op === "TD") {
        if (operands.length !== 2) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        const tx = numberAt(source, operands[0]);
        const ty = numberAt(source, operands[1]);
        event(op, translate(tx, ty, lineMatrix));
        if (op === "TD") leading = -ty;
      } else if (op === "T*") {
        event("T*", translate(0, -leading, lineMatrix));
      } else if (op === "TL") {
        if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        leading = numberAt(source, operands[0]);
      } else if (op === "Tr") {
        if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        mode = numberAt(source, operands[0]);
        if (!Number.isInteger(mode) || mode < 0 || mode > 7) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      } else if (op === "Tf") {
        if (operands.length !== 2) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        fontSize = numberAt(source, operands[1]);
      } else if (op === "Tz") {
        if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        hscale = numberAt(source, operands[0]);
      } else if (op === "Ts") {
        if (operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        rise = numberAt(source, operands[0]);
        riseText = source.slice(operands[0].start, operands[0].end);
      }
      if (op === "'" || op === '"') {
        if (op === '"' && operands.length !== 3) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        if (op === "'" && operands.length !== 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        event(op, translate(0, -leading, lineMatrix));
      }
      if (op === "Tj" || op === "TJ" || op === "'" || op === '"') {
        const isOwned = owned.has(`${start}:${inner.end}`);
        const segment = segments[segments.length - 1];
        if (isOwned) {
          segment.owned += 1;
          ownedHere += 1;
          ownedShow = true;
          if (mode >= 4) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.clip);
          const part = parts.get(`${start}:${inner.end}`);
          // A show the box owns only part of is other text too.
          if (part) segment.other += 1;
          segment.shows.push(Object.freeze({ start, end: inner.end, kind: op, operands: [...operands], fontSize, hscale,
            rise, riseText, ...(part ? { part } : {}) }));
          modes.push(mode);
          showFonts.push(font);
          const boldShow = boldShows ? boldShows.get(`${start}:${inner.end}`) : bold ? true : undefined;
          if (bold && boldShow !== undefined) {
            boldFound += 1;
            // The part the bold is set on lies in the part the box owns.
            const range: readonly [number, number] | undefined = boldShow !== true && boldShow.codeStart !== undefined
              ? [boldShow.codeStart, boldShow.codeEnd!] : part;
            if (range && (range[0] >= range[1] || (part && (range[0] < part[0] || range[1] > part[1])))) {
              invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
            }
            // The width follows the glyphs where a resize scales them.
            if (scaled && !objectMap) objectMap = groupMapUnder(objectState.ctm, transform);
            const wrap = boldWrap(bold.on, mode, entries, fontSize,
              scaled ? multiply(lineMatrix, objectMap!) : lineMatrix, input.lineWidths);
            if (wrap) {
              wraps.push(Object.freeze({ start, end: inner.end, kind: op, operands: [...operands],
                ...(range ? { range } : {}), ...wrap }));
            }
          }
          // A copy drawn outside the clip carries no marked content: none
          // that names a replacement text or a layer may hold the box.
          if (outsideClip && marks.some(Boolean)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.outsideMarked);
        } else {
          segment.other += 1;
        }
      }
      objectOperators.push(Object.freeze({
        op, start, end: inner.end, operands: [...operands], segment: segments.length - 1,
        opens: segments.length > segmentCount, owned: ownedShow, font,
      }));
      operands = [];
    }
    if (!closed) {
      // A page split across content streams may end one inside a text
      // object the next stream closes (05-en-financial-tables: the array
      // here, its `TJ ET` there). One that draws none of the box is the last
      // this stream holds: nothing of it moves and nothing after it is here.
      if (input.continued && ownedHere === 0 && index >= tokens.length) break;
      refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
    }
    // Leading and mode set inside the object stay in force after it. A clip
    // by glyph outlines (Tr 4-7) has no shape this reads: the clip before it
    // stands, and the landing is held to that.
    state = { ...state, leading, mode, fontSize, hscale, rise, riseText, entries, font };
    operands = [];
    if (ownedHere === 0) {
      disruptions += 1;
      continue;
    }
    found += ownedHere;
    const ownedObject = { start: bt.start, end: tokens[index - 1].end, ctm: objectState.ctm, clip: objectState.clip,
      disruptions, runs: [runs.length, runs.length] as readonly [number, number] };
    ownedObjects.push(ownedObject);
    if (outsideClip) {
      const map = groupMapUnder(objectState.ctm, transform);
      objects.push({
        edits: outsideEdits(source, segments, objectOperators, map, objectState.ctm, runs, fonts, extGStates, wraps),
        clip: objectState.clip,
      });
      ownedObject.runs = [ownedObject.runs[0], runs.length];
      continue;
    }
    if (groupMapMovesOnly(transform) && transform.e === 0 && transform.f === 0) {
      // The box stays where it is (a weight alone): nothing is placed again,
      // only its bolded shows are written.
      objects.push({
        edits: wraps.map((wrap) => ({ start: wrap.start, end: wrap.end, text: boldWrapped(source, wrap) })),
        clip: objectState.clip,
      });
      continue;
    }
    const pureMove = groupMapMovesOnly(transform);
    for (const segment of segments) {
      if (segment.owned > 0 && segment.other > 0) {
        // Box text on a line of other text moves by its own kerning and
        // rise; a resize would change its advance, which only a font's
        // widths could give back.
        if (!pureMove) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.shared);
        segment.mixed = true;
        segment.flag = false;
        continue;
      }
      segment.flag = segment.owned > 0 ? true : segment.other > 0 ? false : null;
    }
    // A segment that shows nothing stays in the space before it, so the
    // operator that opens the next segment with text is the boundary.
    let previous = false;
    for (const segment of segments) {
      if (segment.flag === null) segment.flag = previous;
      previous = segment.flag;
    }
    const map = groupMapUnder(objectState.ctm, transform);
    const edits: { start: number; end: number; text: string }[] = [];
    // A box show on a line of other text, drawn where the map puts it by a
    // kerning before it (taken back after it) and a rise (set back after it):
    // the pen after it, and every other show, stay where they were.
    const kerned = (show: OwnedShow, lineLinear: Matrix): string => {
      const [c0, c1, c2, c3] = objectState.ctm;
      const m00 = lineLinear[0] * c0 + lineLinear[1] * c2;
      const m01 = lineLinear[0] * c1 + lineLinear[1] * c3;
      const m10 = lineLinear[2] * c0 + lineLinear[3] * c2;
      const m11 = lineLinear[2] * c1 + lineLinear[3] * c3;
      const determinant = m00 * m11 - m01 * m10;
      const scale = (show.fontSize * show.hscale) / 100;
      if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12 || !(Math.abs(scale) > 1e-9)) {
        refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.shared);
      }
      const u = (transform.e * m11 - transform.f * m10) / determinant;
      const v = (transform.f * m00 - transform.e * m01) / determinant;
      const kern = micro((-u * 1_000) / scale);
      const lift = written(show.rise + v);
      const before = formatMraiNativeTextMoveMicro(kern);
      const after = formatMraiNativeTextMoveMicro(kern === 0 ? 0 : -kern);
      const text = (at: number) => source.slice(show.operands[at].start, show.operands[at].end);
      if (show.part) {
        // Only this range of the show's codes is the box's: the show is cut
        // there, the rest drawn on from where it was.
        const pieces = splitShowPayload(source, show.operands[show.kind === '"' ? 2 : 0], show.kind === "TJ", show.part);
        return [
          ...(pieces.before ? [`[${pieces.before}] TJ`] : []),
          `${lift.text} Ts [${before} ${pieces.owned} ${after}] TJ ${show.riseText} Ts`,
          ...(pieces.after ? [`[${pieces.after}] TJ`] : []),
        ].join(" ");
      }
      const payload = show.kind === "TJ" ? text(0).slice(1, -1) : text(show.kind === '"' ? 2 : 0);
      return `${lift.text} Ts [${before} ${payload} ${after}] TJ ${show.riseText} Ts`;
    };
    for (const segment of segments) {
      if (!segment.mixed) continue;
      const opener = segment.opener;
      const linear = opener ? opener.lineMatrix : ([1, 0, 0, 1, 0, 0] as Matrix);
      for (const show of segment.shows) {
        // A quote show opens its segment: the opener's edit writes it.
        if (opener && show.start === opener.start && show.end === opener.end) continue;
        edits.push({ start: show.start, end: show.end, text: kerned(show, linear) });
      }
    }
    let current: Matrix = [1, 0, 0, 1, 0, 0];
    if (segments[0].flag) {
      const values = map.map(written);
      edits.push({ start: bt.end, end: bt.end, text: ` ${formatMatrix(values)} Tm` });
      current = values.map((value) => value.value) as unknown as Matrix;
    }
    for (let k = 1; k < segments.length; k += 1) {
      const opener = segments[k].opener!;
      const flag = segments[k].flag!;
      const desired = flag ? multiply(opener.lineMatrix, map) : opener.lineMatrix;
      if (opener.kind === "Tm") {
        if (!flag) {
          current = opener.lineMatrix;
          continue;
        }
        const values = desired.map((value, position) =>
          value === opener.lineMatrix[position]
            ? { text: source.slice(opener.operands[position].start, opener.operands[position].end), value }
            : written(value),
        );
        edits.push({ start: opener.start, end: opener.end, text: `${formatMatrix(values)} Tm` });
        current = values.map((value) => value.value) as unknown as Matrix;
        continue;
      }
      const quoted = segments[k].mixed && (opener.kind === "'" || opener.kind === '"')
        ? segments[k].shows.find((show) => show.start === opener.start && show.end === opener.end)
        : undefined;
      if (flag === segments[k - 1].flag! && !quoted) {
        // The same space on both sides: the operator carries the map on.
        current =
          opener.kind === "Td" || opener.kind === "TD"
            ? translate(numberAt(source, opener.operands[0]), numberAt(source, opener.operands[1]), current)
            : translate(0, -opener.leading, current);
        continue;
      }
      // A boundary: reach the other space from where the rewrite stands (a
      // quote show of the box on a line of other text keeps its own move to
      // the next line).
      let position: string;
      if (flag === segments[k - 1].flag!) {
        position = "T*";
        current = translate(0, -opener.leading, current);
      } else if (
        current[0] === desired[0] &&
        current[1] === desired[1] &&
        current[2] === desired[2] &&
        current[3] === desired[3]
      ) {
        const determinant = current[0] * current[3] - current[1] * current[2];
        if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) {
          refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
        }
        const dx = desired[4] - current[4];
        const dy = desired[5] - current[5];
        const tx = written((dx * current[3] - dy * current[2]) / determinant);
        const ty = written((dy * current[0] - dx * current[1]) / determinant);
        position = `${tx.text} ${ty.text} Td`;
        current = translate(tx.value, ty.value, current);
      } else {
        const values = desired.map(written);
        position = `${formatMatrix(values)} Tm`;
        current = values.map((value) => value.value) as unknown as Matrix;
      }
      const operandText = (at: number) => source.slice(opener.operands[at].start, opener.operands[at].end);
      let text: string;
      if (opener.kind === "Td" || opener.kind === "T*") {
        text = position;
      } else if (opener.kind === "TD") {
        const ty = operandText(1);
        const negated = ty.startsWith("-") ? ty.slice(1) : ty.startsWith("+") ? `-${ty.slice(1)}` : `-${ty}`;
        text = `${negated} TL ${position}`;
      } else if (opener.kind === "'") {
        text = `${position} ${quoted ? kerned(quoted, opener.lineMatrix) : `${operandText(0)} Tj`}`;
      } else {
        text = `${operandText(0)} Tw ${operandText(1)} Tc ${position} ${
          quoted ? kerned(quoted, opener.lineMatrix) : `${operandText(2)} Tj`}`;
      }
      edits.push({ start: opener.start, end: opener.end, text });
    }
    // A bolded show between the state the bold sets and the state put back:
    // around the edit that writes it already (a quote show at a boundary, a
    // show kerned on a line of other text), else as an edit of its own.
    for (const wrap of wraps) {
      const covering = edits.find((edit) => edit.start === wrap.start && edit.end === wrap.end);
      if (!covering) {
        edits.push({ start: wrap.start, end: wrap.end, text: boldWrapped(source, wrap) });
        continue;
      }
      if (wrap.range) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
      covering.text = ` ${wrap.pre} ${covering.text} ${wrap.post}`;
    }
    objects.push({ edits, clip: objectState.clip });
  }
  if (found !== owned.size || owned.size !== input.shows.length || owned.size === 0) {
    refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.proof);
  }
  // A bold names shows of the box only.
  if (boldShows && boldFound !== boldShows.size) invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
  // The box's link underlines go with it: rewritten in place under the same
  // map, or -- the box written outside every clip -- taken out and drawn
  // after the text of its run in the block (where a paragraph's later
  // rewrite finds them under its runs again). Left behind, a paragraph
  // moved and set in a wider frame showed each one through whatever line
  
  const ruleLines: string[] = [];
  let ruleRunCtm: Matrix | null = null;
  const besideRuns = new Map<number, { before: string[]; after: string[] }>();
  if (!(groupMapMovesOnly(transform) && transform.e === 0 && transform.f === 0)) {
    const carried = carriedGroupRules(ruleCandidates, ownedObjects);
    const ruleEdits: { start: number; end: number; text: string }[] = [];
    for (const rule of carried) {
      const map = groupMapUnder(rule.ctm, transform);
      if (Math.abs(map[1]) > 1e-9 || Math.abs(map[2]) > 1e-9) break;
      const [x, y, w, h] = rule.numbers;
      const xs = [x, x + w].map((value) => map[0] * value + map[4]);
      const ys = [y, y + h].map((value) => map[3] * value + map[5]);
      const placed = [Math.min(...xs), Math.min(...ys), Math.max(...xs) - Math.min(...xs),
        Math.max(...ys) - Math.min(...ys)].map((value) => written(value).text).join(" ");
      if (!outsideClip) {
        ruleEdits.push({ start: rule.start, end: rule.reEnd, text: `${placed} re` });
        continue;
      }
      let run: OutsideRun | undefined;
      for (let at = runs.length - 1; at >= 0 && !run; at -= 1) {
        if (runs[at].ctm.every((value, index) => Math.abs(value - rule.ctm[index]) <= 1e-9)) run = runs[at];
      }
      if (!run) continue;
      ruleRunCtm = run.ctm;
      ruleEdits.push({ start: rule.start, end: rule.end, text: " " });
      ruleLines.push(rule.fill!, `${placed} re`, "f");
    }
    // This writer's own rules beside the box's runs (pdfNativeTextRuleBlock:
    // an underline ends at its run's BT, a strikethrough starts after its
    // ET) go with them too, beside the same runs: left, the paragraph's
    // rewrite after the move no longer read them as its runs' and they
    
    const sourceBytes = Uint8Array.from(source, (character) => character.charCodeAt(0) & 0xff);
    const seen = new Set<number>();
    for (const object of ownedObjects) {
      const map = groupMapUnder(object.ctm, transform);
      for (const block of pdfNativeTextRuleBlocksBeside(sourceBytes, object)) {
        if (seen.has(block.start)) continue;
        seen.add(block.start);
        const moved = movedRuleBlock(source.slice(block.start, block.end), map);
        if (!moved) continue;
        if (!outsideClip) {
          ruleEdits.push({ start: block.start, end: block.end, text: moved });
          continue;
        }
        const [firstRun, endRun] = object.runs;
        if (endRun <= firstRun) continue;
        const at = block.style === "underline" ? firstRun : endRun - 1;
        const beside = besideRuns.get(at) ?? { before: [], after: [] };
        (block.style === "underline" ? beside.before : beside.after).push(moved.trim());
        besideRuns.set(at, beside);
        ruleEdits.push({ start: block.start, end: block.end, text: " " });
      }
    }
    if (ruleEdits.length) objects.push({ edits: ruleEdits, clip: ownedObjects[0]!.clip });
  }
  const edits = objects.flatMap((object) => object.edits).sort((left, right) => left.start - right.start);
  if (edits.length === 0) invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unchanged);
  if (edits.length > MAX_GROUP_EDITS) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.size);
  const start = edits[0].start;
  const end = Math.max(...edits.map((edit) => edit.end));
  let replacement = "";
  let cursor = start;
  // An edit never runs into the bytes beside it: "T*" followed by a written
  // "332.96 Ts" read as one unknown token "T*332.96" to pdfium and WPS (the
  
  // native_text_group_move.py.
  const regular = (character: string | undefined) =>
    character !== undefined && !/[\s()<>[\]{}/%]/u.test(character);
  for (const edit of edits) {
    const gap = source.slice(cursor, edit.start);
    const previous = (replacement + gap).at(-1) ?? (start > 0 ? source[start - 1] : undefined);
    const text = (regular(previous) && regular(edit.text[0]) ? " " : "") + edit.text +
      (regular(edit.text.at(-1)) && regular(source[edit.end]) ? " " : "");
    replacement += gap + text;
    cursor = edit.end;
  }
  replacement += source.slice(cursor, end);
  if (!outsideClip) {
    return Object.freeze({
      span: Object.freeze({ start, end }),
      replacement,
      clips: Object.freeze(objects.map((object) => object.clip)),
      modes: Object.freeze(modes),
      showFonts: Object.freeze(showFonts),
    });
  }
  // The block lands at the tail, where the state is the one the stream's
  // first open `q` saved (or the last one, with none open): it closes them,
  // then draws each run in a `q … Q` of its own.
  const tail = stack.length > 0 ? stack[0] : state;
  const block = outsideBlock(source, runs, tail, stack.length, marks.length,
    ruleRunCtm ? { ctm: ruleRunCtm, lines: ruleLines } : null, besideRuns);
  return Object.freeze({
    span: Object.freeze({ start, end: source.length }),
    replacement: replacement + source.slice(end) + block,
    clips: Object.freeze([tail.clip]),
    outside: Object.freeze({
      middle: replacement,
      middleEnd: end,
      block,
      fonts: Object.freeze([...fonts].sort()),
      extGStates: Object.freeze([...extGStates].sort()),
    }),
    modes: Object.freeze(modes),
    showFonts: Object.freeze(showFonts),
  });
}

/**
 * One text object of a box written outside the clip (mirror of
 * _outside_edits): each of the box's shows taken out of its place -- `'` and
 * `"` keep their move to the next line, and `"` the spacing it sets -- and
 * each run of the box's segments recorded for the block. A segment holding
 * other text too is refused: its shows place the text after them.
 */
function outsideEdits(
  source: string,
  segments: readonly Segment[],
  operators: readonly ObjectOperator[],
  map: Matrix,
  ctm: Matrix,
  runs: OutsideRun[],
  fonts: Set<string>,
  extGStates: Set<string>,
  wraps: readonly BoldWrap[],
): { start: number; end: number; text: string }[] {
  for (const segment of segments) {
    if (segment.shows.some((show) => show.part)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.partial);
    if (segment.owned > 0 && segment.other > 0) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.shared);
  }
  // A bolded show is drawn again between the state its bold sets and the
  // state put back; a part of a show is not drawn again outside the clip.
  const wrapOf = new Map(wraps.map((wrap) => [`${wrap.start}:${wrap.end}`, wrap]));
  if (wraps.some((wrap) => wrap.range)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
  const wrapped = (start: number, end: number, line: string): string => {
    const wrap = wrapOf.get(`${start}:${end}`);
    return wrap ? `${wrap.pre} ${line} ${wrap.post}` : line;
  };
  const text = (token: Token) => source.slice(token.start, token.end);
  const edits: { start: number; end: number; text: string }[] = [];
  for (const operator of operators) {
    if (!operator.owned) continue;
    const removed = operator.op === "'"
      ? " T* "
      : operator.op === '"'
      ? ` ${text(operator.operands[0])} Tw ${text(operator.operands[1])} Tc T* `
      : "";
    edits.push({ start: operator.start, end: operator.end, text: removed });
  }
  let first = 0;
  while (first < segments.length) {
    if (segments[first].owned === 0) {
      first += 1;
      continue;
    }
    // The run: this segment and the next ones with no other text, to the
    // last with a show of the box.
    let last = first;
    for (let next = first + 1; next < segments.length && segments[next].other === 0; next += 1) {
      if (segments[next].owned > 0) last = next;
    }
    const opener = segments[first].opener;
    const lineMatrix: Matrix = opener ? opener.lineMatrix : [1, 0, 0, 1, 0, 0];
    const lines = [`${formatMatrix(multiply(lineMatrix, map).map(written))} Tm`];
    if (opener?.kind === "'") lines.push(wrapped(opener.start, opener.end, `${text(opener.operands[0])} Tj`));
    if (opener?.kind === '"') {
      lines.push(wrapped(opener.start, opener.end,
        `${text(opener.operands[0])} Tw ${text(opener.operands[1])} Tc ${text(opener.operands[2])} Tj`));
    }
    let lastShow = -1;
    operators.forEach((operator, at) => {
      if (operator.segment === last && operator.owned) lastShow = at;
    });
    for (let at = 0; at <= lastShow; at += 1) {
      const operator = operators[at];
      if (operator.segment < first || (operator.segment === first && operator.opens)) {
        if (operator.segment === first && operator.owned) fonts.add(operator.font);
        continue;
      }
      if (operator.opens && operator.op === "Tm") {
        const opened = segments[operator.segment].opener!;
        lines.push(`${formatMatrix(multiply(opened.lineMatrix, map).map(written))} Tm`);
        continue;
      }
      if (MARKED.has(operator.op)) continue;
      const copied = operatorText(source, operator.op, operator.operands);
      if (operator.opens || operator.owned) {
        if (operator.owned) fonts.add(operator.font);
        lines.push(operator.owned ? wrapped(operator.start, operator.end, copied) : copied);
        continue;
      }
      const entry = stateEntry(source, operator.op, operator.operands);
      if (!entry) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
      if (isPatternColor(entry.text)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.outsideState);
      if (entry.key.startsWith("gs ")) extGStates.add(entry.key.slice(3));
      lines.push(copied);
    }
    const entries = segments[first].entries;
    for (const entry of entries) {
      if (isPatternColor(entry.text)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.outsideState);
      if (entry.key.startsWith("gs ")) extGStates.add(entry.key.slice(3));
    }
    runs.push(Object.freeze({ ctm, entries, lines: Object.freeze(lines) }));
    first = last + 1;
  }
  return edits;
}

/**
 * The block appended at the tail of the stream (mirror of _outside_block):
 * the marked content and the `q` levels the stream leaves open closed, then
 * each run in a `q … Q` of its own -- its CTM read from the tail's, the
 * state the tail sets but the run never did set back, the run's state in
 * force, and its text object.
 */
function outsideBlock(
  source: string,
  runs: readonly OutsideRun[],
  tail: GraphicsState,
  depth: number,
  openMarks: number,
  /** The box's link underlines, drawn after the text of its last run under
   * this matrix (carriedGroupRules). */
  rules: Readonly<{ ctm: Matrix; lines: readonly string[] }> | null = null,
  /** This writer's own rule blocks of a run, by its index: an underline
   * before its text object, a strikethrough after it (as written beside it). */
  beside: ReadonlyMap<number, Readonly<{ before: readonly string[]; after: readonly string[] }>> = new Map(),
): string {
  const [a, b, c, d, e, f] = tail.ctm;
  const determinant = a * d - b * c;
  if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) {
    refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unreadable);
  }
  const inverse: Matrix = [
    d / determinant,
    -b / determinant,
    -c / determinant,
    a / determinant,
    (c * f - d * e) / determinant,
    (b * e - a * f) / determinant,
  ];
  const lines: string[] = [...Array<string>(openMarks).fill("EMC"), ...Array<string>(depth).fill("Q")];
  let ruleRun = -1;
  for (let at = runs.length - 1; rules && at >= 0 && ruleRun < 0; at -= 1) if (runs[at].ctm === rules.ctm) ruleRun = at;
  for (const [runIndex, run] of runs.entries()) {
    const keys = new Set(run.entries.map((entry) => entry.key));
    const defaults: string[] = [];
    for (const entry of tail.entries) {
      if (keys.has(entry.key) || entry.key === "Tf") continue;
      if (entry.key.startsWith("gs ")) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.outsideState);
      const group = entry.key === FILL_COLOR ? FILL : entry.key === STROKE_COLOR ? STROKE : entry.key;
      if (keys.has(group)) continue;
      const reset = STATE_DEFAULTS[group];
      if (reset === undefined) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.outsideState);
      if (!defaults.includes(reset)) defaults.push(reset);
    }
    const relative = multiply(run.ctm, inverse).map(written);
    const identity = relative.map((value) => value.text).join(" ") === "1 0 0 1 0 0";
    lines.push(
      "q",
      ...(identity ? [] : [`${formatMatrix(relative)} cm`]),
      ...defaults,
      ...run.entries.map((entry) => entry.text),
      ...(beside.get(runIndex)?.before ?? []),
      "BT",
      ...run.lines,
      "ET",
      ...(beside.get(runIndex)?.after ?? []),
      ...(runIndex === ruleRun ? rules!.lines : []),
      "Q",
    );
  }
  const separator = source.length === 0 || WHITE.has(source.charCodeAt(source.length - 1)) ? "" : "\n";
  return `${separator}${lines.join("\n")}\n`;
}

// ---------------------------------------------------------------------------
// What a proof owns.

/** The leaves of a no-op proof that hold only a reader chunk of their show
 * (mirror of _group_proof_chunk_leaves): each show moves whole only when the
 * rest of it is white space. */
export function mraiNativeTextGroupProofChunkLeaves(
  evidence: PdfNativeListTextFidelityEvidence,
): readonly PdfNativeListTextFidelityEvidence[] {
  return groupProofShowLeaves(evidence).filter((leaf) => leaf.readerChunk);
}

/** The leaves of a no-op proof that each prove one show of their own (no
 * line plan): the show, and the font it is set in. */
function groupProofShowLeaves(
  evidence: PdfNativeListTextFidelityEvidence,
): readonly PdfNativeListTextFidelityEvidence[] {
  const read = paragraphReadItems(evidence, true);
  if (read) return read.flatMap((item) => groupProofShowLeaves(item));
  if (evidence.positionedLines) {
    return evidence.positionedLines.lines.flatMap((line) => groupProofShowLeaves(line.evidence));
  }
  if (evidence.positionedRunGroup) {
    return evidence.positionedRunGroup.runs.flatMap((run) => groupProofShowLeaves(run.evidence));
  }
  return evidence.linePlan ? [] : [evidence];
}

/** What may stand between a show of the box and a white-space show after it
 * on its line: marked content and spacing. Anything else (a font, a mode, a
 * colour, a move to another line) ends what the box's show proves of it. */
const LINE_BLANK_NEIGHBOUR_OPERATORS = new Set(["BMC", "BDC", "EMC", "MP", "DP", "Tc", "Tw"]);














async function groupLineBlankShows(
  sourceBytes: Uint8Array,
  streamBytes: Uint8Array,
  shows: readonly GroupShow[],
  leaves: ReadonlyMap<string, PdfNativeListTextFidelityEvidence>,
): Promise<readonly GroupShow[]> {
  if (leaves.size === 0) return [];
  const source = LATIN1.decode(streamBytes);
  const tokens = groupTokens(source);
  if (!tokens) return [];
  const owned = new Map(shows.map((show) => [`${show.start}:${show.end}`, show]));
  const blanks: GroupShow[] = [];
  let inText = false;
  // The leaf of the box's show the pen comes from, while nothing since has
  // changed the line or the state it was drawn in.
  let anchor: PdfNativeListTextFidelityEvidence | null = null;
  let operands: Token[] = [];
  for (const token of tokens) {
    if (!token.operator) {
      operands.push(token);
      continue;
    }
    const op = source.slice(token.start, token.end);
    const start = operands.length > 0 ? operands[0].start : token.start;
    if (op === "BT" || op === "ET") {
      inText = op === "BT";
      anchor = null;
    } else if (!inText) {
      // Nothing outside a text object is a show.
    } else if (op === "Tj" || op === "TJ" || op === "'" || op === '"') {
      const key = `${start}:${token.end}`;
      const show = owned.get(key);
      if (show) {
        anchor = show.codeStart === undefined ? leaves.get(key) ?? null : null;
      } else if (
        anchor && (op === "Tj" || op === "TJ") && operands.length === 1 &&
        (await mraiNativeShowIsBlankInLeafFont(sourceBytes, anchor, streamBytes, { start, end: token.end }))
      ) {
        blanks.push(Object.freeze({ start, end: token.end }));
      } else {
        anchor = null;
      }
    } else if (op === "Td" || op === "TD") {
      const height = operands.length === 2 ? source.slice(operands[1].start, operands[1].end) : "";
      if (!NUMBER.test(height) || Number(height) !== 0) anchor = null;
    } else if (!LINE_BLANK_NEIGHBOUR_OPERATORS.has(op)) {
      anchor = null;
    }
    operands = [];
  }
  return blanks;
}

/** The items of a paragraph rewrite's read proof (no block), each proved
 * alone; null for any other proof (mirror of _paragraph_read_items). */
function paragraphReadItems(
  evidence: PdfNativeListTextFidelityEvidence,
  top: boolean,
): readonly PdfNativeListTextFidelityEvidence[] | null {
  const rewrite = evidence.paragraphRewrite;
  if (!top || !rewrite || rewrite.block || rewrite.history) return null;
  const items = rewrite.items.filter((item): item is PdfNativeListTextFidelityEvidence => item !== null);
  if (items.length === 0) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.proof);
  return items;
}

/** The shows a no-op proof of a box names (mirror of _group_proof_shows). A
 * show named by several reader chunks is named once. */
export function mraiNativeTextGroupProofShows(
  evidence: PdfNativeListTextFidelityEvidence,
  top = true,
): readonly Span[] {
  // A paragraph the surgical routes cannot prove whole opens by its items
  // proved alone (the paragraph rewrite's read proof): they name its shows.
  const read = paragraphReadItems(evidence, top);
  if (read) return read.flatMap((item) => mraiNativeTextGroupProofShows(item));
  if (
    (top &&
      evidence.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION &&
      evidence.schemaVersion !== MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION) ||
    (top && !evidence.selection) ||
    evidence.paragraphRewrite ||
    evidence.paragraphReflow ||
    evidence.propertyMutation ||
    evidence.consolidatedRuns ||
    evidence.emptiedShow ||
    evidence.positionedRestore ||
    evidence.listAnchor ||
    evidence.originalText !== evidence.replacementText ||
    evidence.originalDisplayText !== evidence.replacementDisplayText ||
    JSON.stringify(evidence.originalCodeBytes) !== JSON.stringify(evidence.replacementCodeBytes)
  ) {
    refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.proof);
  }
  if (evidence.positionedLines) {
    return evidence.positionedLines.lines.flatMap((line) => mraiNativeTextGroupProofShows(line.evidence, false));
  }
  if (evidence.positionedRunGroup) {
    return evidence.positionedRunGroup.runs.flatMap((run) => mraiNativeTextGroupProofShows(run.evidence, false));
  }
  if (evidence.linePlan) {
    return evidence.linePlan.lines.map((line) =>
      Object.freeze({ start: line.operatorSpan.start, end: line.operatorSpan.end }),
    );
  }
  const span = evidence.textOperatorSpan;
  const showEnd = evidence.lineShift?.showEnd;
  return [
    Object.freeze({
      start: span.start,
      end: showEnd !== undefined && span.start < showEnd && showEnd < span.end ? showEnd : span.end,
    }),
  ];
}

const LATIN1 = LATIN1_EXACT;

interface ResolvedGroup {
  readonly streams: readonly PdfNativePageContentStreamInspection[];
  readonly index: number;
  readonly shows: readonly GroupShow[];
  /** The shows the proof names alone, where `shows` holds white-space shows
   * of the box's lines as well (groupLineBlankShows): what the box moves as
   * where it cannot move with those. */
  readonly named?: readonly GroupShow[];
  readonly duplicates: readonly Span[] | undefined;
}

/** What owns a box's shows: its writer proof (and the no-op request that
 * proves it), or the paragraph rewrite's read proof of its items alone. */
interface GroupOwnership {
  readonly evidence: PdfNativeListTextFidelityEvidence;
  readonly target: Readonly<{ objectNumber: number; generation: number }>;
  readonly request?: PdfNativeContentStreamTransactionRequest;
  /** The proof reads the page's streams joined as one (its offsets are the
   * joined drawing's). */
  readonly consolidated?: boolean;
}

async function resolveGroup(
  sourceBytes: Uint8Array,
  ownership: GroupOwnership,
): Promise<ResolvedGroup> {
  const { evidence } = ownership;
  const request: Partial<PdfNativeContentStreamTransactionRequest> = ownership.request ?? {};
  const named = new Map<string, GroupShow>();
  for (const show of mraiNativeTextGroupProofShows(evidence)) named.set(`${show.start}:${show.end}`, show);
  let shows: GroupShow[] = [...named.values()];
  const consolidated = Boolean(request.pageContentConsolidation) || ownership.consolidated === true;
  if (
    request.formInstanceIsolation ||
    request.pageExtGStateMutation ||
    request.pageFontExtensionMutation ||
    request.pageXObjectUnmap ||
    request.contentAppend ||
    request.replacementPieces
  ) {
    refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.source);
  }
  const streams = await inspectMraiNativePageContentStreams(sourceBytes, evidence.selection!.pageNumber);
  let index = streams.findIndex(
    (stream) =>
      stream.identity.objectNumber === ownership.target.objectNumber &&
      stream.identity.generation === ownership.target.generation,
  );
  // The bytes the proof's offsets address, and where the stream the box is
  // drawn in starts among them.
  let coordinateBytes: Uint8Array;
  let offset = 0;
  if (index < 0) {
    // A page whose streams a reader joins (the proof reads them as one, as a
    // text object straddles two of them): the box moves in the one stream of
    // them that draws all its shows.
    if (!consolidated || shows.length === 0) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.source);
    // The proof's offsets are the consolidated stream's: the streams end to
    // end, nothing between them (buildMraiNativeConsolidatedContentStream).
    // `drawingOffset` counts a separator byte after each stream, so every
    // stream after the first was looked for a byte late and every box on
    // such a page lost its group: its size, bold and move (the title of
    // 05-en-financial-tables went to a text box, PDF AI eval V1-T04).
    const starts: number[] = [];
    let at = 0;
    for (const entry of streams) {
      starts.push(at);
      at += entry.streamBytes.length;
    }
    coordinateBytes = new Uint8Array(at);
    streams.forEach((entry, k) => coordinateBytes.set(entry.streamBytes, starts[k]));
    const span = evidence.textOperatorSpan;
    if (span.end > at || (await sha256PdfBytes(coordinateBytes.slice(span.start, span.end))) !== span.sha256) {
      refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.source);
    }
    const low = Math.min(...shows.map((show) => show.start));
    const high = Math.max(...shows.map((show) => show.end));
    index = streams.findIndex((entry, k) => starts[k] <= low && high <= starts[k] + entry.streamBytes.length);
    if (index < 0) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.source);
    offset = starts[index];
  } else {
    const found = streams[index];
    if (
      ownership.request &&
      (found.streamSha256 !== ownership.request.expectedStreamSha256 ||
        found.objectSha256 !== ownership.request.expectedObjectSha256)
    ) {
      throw new PdfNativeTransactionError("precondition-mismatch", "The text's content stream changed.");
    }
    coordinateBytes = found.streamBytes;
  }
  // A show the proof holds only a reader chunk of moves whole when the rest
  // of it is white space the reader never writes; other text in it (a
  // contents line's page number) would move with the box.
  const chunkLeaves = mraiNativeTextGroupProofChunkLeaves(evidence);
  if (chunkLeaves.length > 0) {
    const byShow = new Map<string, { leaf: PdfNativeListTextFidelityEvidence; chunks: { codeStart: number; codeEnd: number }[] }>();
    for (const leaf of chunkLeaves) {
      const key = `${leaf.textOperatorSpan.start}`;
      const entry = byShow.get(key) ?? { leaf, chunks: [] };
      entry.chunks.push({ codeStart: leaf.readerChunk!.codeStart, codeEnd: leaf.readerChunk!.codeEnd });
      byShow.set(key, entry);
    }
    for (const { leaf, chunks } of byShow.values()) {
      if (await mraiNativeReaderChunkRestIsBlank(sourceBytes, leaf, coordinateBytes, chunks)) continue;
      // Other text shares the show (a contents line's page number): the box
      // owns the range its chunks span, when the gaps between them are white
      // space, and the show is cut around it.
      const codeStart = Math.min(...chunks.map((chunk) => chunk.codeStart));
      const codeEnd = Math.max(...chunks.map((chunk) => chunk.codeEnd));
      if (!(await mraiNativeReaderChunkRestIsBlank(sourceBytes, leaf, coordinateBytes, chunks,
        { start: codeStart, end: codeEnd }))) {
        refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.partial);
      }
      const span = leaf.textOperatorSpan;
      const showEnd = leaf.lineShift?.showEnd;
      const end = showEnd !== undefined && span.start < showEnd && showEnd < span.end ? showEnd : span.end;
      const key = `${span.start}:${end}`;
      if (!named.has(key)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.proof);
      named.set(key, Object.freeze({ start: span.start, end, codeStart, codeEnd }));
    }
    shows = [...named.values()];
  }
  let duplicates: Span[] | undefined;
  if (request.duplicateSpans?.length) {
    // A fake-bold overprint pass moves with its box (the text move's rule).
    const source = LATIN1.decode(coordinateBytes);
    const span = evidence.textOperatorSpan;
    const twin = shows.length === 1 ? qtFakeBoldTwinSpan(source, span, evidence.fontSize) : null;
    if (
      request.duplicateSpans.length !== 1 ||
      !twin ||
      twin.start !== request.duplicateSpans[0].start ||
      twin.end !== request.duplicateSpans[0].end
    ) {
      refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.twin);
    }
    const shift = twin.start - span.start;
    duplicates = [Object.freeze({ start: twin.start - offset, end: twin.end - offset })];
    shows = [...shows, ...shows.map((show) => Object.freeze({ start: show.start + shift, end: show.end + shift }))];
  }
  // Offsets into the stream the box is drawn in.
  if (offset !== 0) {
    shows = shows.map((show) => Object.freeze({ ...show, start: show.start - offset, end: show.end - offset }));
    const stream = streams[index];
    if (shows.some((show) => show.start < 0 || show.end > stream.streamBytes.length)) {
      refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.source);
    }
  }
  shows.sort((left, right) => left.start - right.start);
  for (let k = 1; k < shows.length; k += 1) {
    if (shows[k - 1].end > shows[k].start) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.proof);
  }
  // The white-space shows of the box's lines go with it (a fake-bold twin's
  // box is its one show and that show's copy).
  const leaves = new Map<string, PdfNativeListTextFidelityEvidence>();
  if (!duplicates) {
    for (const leaf of groupProofShowLeaves(evidence)) {
      const span = leaf.textOperatorSpan;
      const showEnd = leaf.lineShift?.showEnd;
      const end = showEnd !== undefined && span.start < showEnd && showEnd < span.end ? showEnd : span.end;
      leaves.set(`${span.start - offset}:${end - offset}`, leaf);
    }
  }
  const blanks = await groupLineBlankShows(sourceBytes, streams[index].streamBytes, shows, leaves);
  if (blanks.length === 0) return Object.freeze({ streams, index, shows: Object.freeze(shows), duplicates });
  const withBlanks = [...shows, ...blanks].sort((left, right) => left.start - right.start);
  return Object.freeze({ streams, index, shows: Object.freeze(withBlanks), named: Object.freeze(shows), duplicates });
}

/** Where the box's reader items are on the page: their union, in user space. */
function selectionBounds(selection: PdfNativeListTextSelectionEvidence): Rect | null {
  const items = selection.items ?? [selection.item];
  let bounds: [number, number, number, number] | null = null;
  for (const item of items) {
    if (!/\S/u.test(item.str)) continue;
    const [a, b, c, d, e, f] = item.transform;
    const along = Math.hypot(a, b);
    const up = Math.hypot(c, d);
    if (!(along > 0) || !(up > 0)) continue;
    const ux = a / along;
    const uy = b / along;
    const vx = c / up;
    const vy = d / up;
    const height = item.height > 0 ? item.height : up;
    const corners = [
      [e - vx * height * 0.25, f - vy * height * 0.25],
      [e + ux * item.width - vx * height * 0.25, f + uy * item.width - vy * height * 0.25],
      [e + vx * height, f + vy * height],
      [e + ux * item.width + vx * height, f + uy * item.width + vy * height],
    ];
    for (const [x, y] of corners) {
      bounds = bounds
        ? [Math.min(bounds[0], x), Math.min(bounds[1], y), Math.max(bounds[2], x), Math.max(bounds[3], y)]
        : [x, y, x, y];
    }
  }
  return bounds;
}

function mapRect(rect: Rect, transform: PdfNativeTextGroupTransform): Rect {
  const [a, b, c, d] = groupLinear(transform);
  const corners = [[rect[0], rect[1]], [rect[2], rect[1]], [rect[0], rect[3]], [rect[2], rect[3]]]
    .map(([x, y]) => [x * a + y * c + transform.e, x * b + y * d + transform.f]);
  const xs = corners.map(([x]) => x), ys = corners.map(([, y]) => y);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

function visibleShare(rect: Rect, clip: Rect): number {
  const area = Math.max(0, rect[2] - rect[0]) * Math.max(0, rect[3] - rect[1]);
  if (!(area > 0)) return 1;
  const width = Math.max(0, Math.min(rect[2], clip[2]) - Math.max(rect[0], clip[0]));
  const height = Math.max(0, Math.min(rect[3], clip[3]) - Math.max(rect[1], clip[1]));
  return (width * height) / area;
}

/**
 * The landing is proven where every clip the box is drawn under shows as
 * much of it as it did before the move. A clip whose shape is not a
 * rectangle cannot be measured and is refused.
 */
export function assertMraiNativeTextGroupLandingVisible(
  clips: readonly Clip[],
  bounds: Rect | null,
  transform: PdfNativeTextGroupTransform,
): void {
  if (!landingVisible(clips, bounds, transform)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.landing);
}

function landingVisible(clips: readonly Clip[], bounds: Rect | null, transform: PdfNativeTextGroupTransform): boolean {
  for (const clip of clips) {
    if (clip === null || !bounds) continue;
    const before = visibleShare(bounds, clip);
    const after = visibleShare(mapRect(bounds, transform), clip);
    if (after < Math.min(1, before) - CLIP_VISIBLE_SLACK) return false;
  }
  return true;
}










function landingKeepsWidth(clips: readonly Clip[], bounds: Rect | null, transform: PdfNativeTextGroupTransform,
  frameRight?: number): boolean {
  if (!bounds) return true;
  const moved = mapRect(bounds, transform);
  const landing: Rect = frameRight === undefined || !Number.isFinite(frameRight) ? moved
    : [moved[0], moved[1], Math.max(moved[2], frameRight), moved[3]];
  const hidden = (rect: Rect, clip: Rect) =>
    (rect[2] - rect[0]) - Math.max(0, Math.min(rect[2], clip[2]) - Math.max(rect[0], clip[0]));
  for (const clip of clips) {
    if (clip === null) continue;
    if (hidden(landing, clip) > Math.max(0, hidden(bounds, clip)) + CLIP_WIDTH_TOLERANCE) return false;
  }
  return true;
}

const TYPE3_FONT = /\/Subtype\s*\/Type3(?![^\s/<>[\]()%{}])/;
const SOFT_MASK = /\/SMask(?![^\s/<>[\]()%{}])\s*(\/[^\s/<>[\]()%{}]*|[^\s/])/g;

/**
 * The resources a block drawn outside the clip draws with must draw the
 * same there (mirror of _assert_outside_resources): no Type3 font, and no
 * ExtGState with a soft mask -- a mask is placed where its `gs` ran and
 * hides what lands outside it, as a clip does.
 */
export async function assertMraiNativeTextGroupOutsideResources(
  sourceBytes: Uint8Array,
  pageNumber: number,
  fonts: readonly string[],
  extGStates: readonly string[],
): Promise<void> {
  // The resources name each by its value: `/GS#31` is /GS1.
  const name = (value: string) => decodePdfSourceName(value);
  const fontDictionaries = await readMraiNativePageResourceDictionaries(
    sourceBytes, pageNumber, "Font", fonts.map(name));
  for (const font of fonts) {
    const dictionary = fontDictionaries.get(name(font));
    if (!dictionary || TYPE3_FONT.test(dictionary)) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.outsideResource);
  }
  const stateDictionaries = await readMraiNativePageResourceDictionaries(
    sourceBytes, pageNumber, "ExtGState", extGStates.map(name));
  for (const state of extGStates) {
    const dictionary = stateDictionaries.get(name(state));
    if (!dictionary || [...dictionary.matchAll(SOFT_MASK)].some((match) => match[1] !== "/None")) {
      refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.outsideResource);
    }
  }
}

/** A group bold as a record or payload carries it (mirror of the server's
 * NativeTextGroupBold). */
export function readMraiNativeTextGroupBold(value: unknown): PdfNativeTextGroupBold {
  const record = value as Record<string, unknown> | null;
  if (!record || typeof record !== "object" || Array.isArray(record) || typeof record.on !== "boolean" ||
    Object.keys(record).some((key) => key !== "on" && key !== "shows")) {
    invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldUnavailable);
  }
  if (record.shows === undefined) return Object.freeze({ on: record.on });
  if (!Array.isArray(record.shows) || record.shows.length === 0 || record.shows.length > MAX_GROUP_EDITS) {
    invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
  }
  const integer = (entry: unknown) => typeof entry === "number" && Number.isSafeInteger(entry) && entry >= 0;
  let last = -1;
  const shows = record.shows.map((entry: unknown) => {
    const show = entry as Record<string, unknown> | null;
    if (!show || typeof show !== "object" || !integer(show.start) || !integer(show.end) ||
      (show.end as number) <= (show.start as number) || (show.start as number) < last ||
      (show.codeStart === undefined) !== (show.codeEnd === undefined) ||
      (show.codeStart !== undefined && (!integer(show.codeStart) || !integer(show.codeEnd) ||
        (show.codeEnd as number) <= (show.codeStart as number))) ||
      Object.keys(show).some((key) => !["start", "end", "codeStart", "codeEnd"].includes(key))) {
      invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldPart);
    }
    last = show.end as number;
    return Object.freeze({
      start: show.start as number,
      end: show.end as number,
      ...(show.codeStart !== undefined ? { codeStart: show.codeStart as number, codeEnd: show.codeEnd as number } : {}),
    });
  });
  return Object.freeze({ on: record.on, shows: Object.freeze(shows) });
}

/** A group edit as a payload carries it: its map, and the bold beside it
 * (`bold`, absent for a map alone). A map that leaves the box where it was
 * is an edit only with a bold. */
export function readMraiNativeTextGroupTransform(value: unknown, bold?: unknown): PdfNativeTextGroupEdit {
  const record = value as Record<string, unknown> | null;
  if (!record || typeof record !== "object") invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.transform);
  const numbers = ["a", "d", "e", "f"].map((key) => record[key]);
  const turn = ["b", "c"].map((key) => record[key] ?? 0);
  if ([...numbers, ...turn].some((entry) => typeof entry !== "number" || !Number.isFinite(entry))) {
    invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.transform);
  }
  const [a, d, e, f] = numbers as number[];
  const [b, c] = turn as number[];
  // The box's scale along each of its own axes, whichever way it is turned;
  // a map that mirrors it is no move of a text block.
  const alongX = Math.hypot(a, b), alongY = Math.hypot(c, d);
  if (
    alongX < MIN_GROUP_SCALE ||
    alongX > MAX_GROUP_SCALE ||
    alongY < MIN_GROUP_SCALE ||
    alongY > MAX_GROUP_SCALE ||
    !(a * d - b * c > 0) ||
    ((b !== 0 || c !== 0) ? false : a < MIN_GROUP_SCALE || d < MIN_GROUP_SCALE) ||
    Math.abs(e) > MAX_GROUP_MOVE_DISTANCE ||
    Math.abs(f) > MAX_GROUP_MOVE_DISTANCE
  ) {
    invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.transform);
  }
  const turned = b !== 0 || c !== 0 ? { b, c } : {};
  if (bold !== undefined) return Object.freeze({ a, ...turned, d, e, f, bold: readMraiNativeTextGroupBold(bold) });
  if (a === 1 && d === 1 && b === 0 && c === 0 && e === 0 && f === 0) invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.unchanged);
  return Object.freeze({ a, ...turned, d, e, f });
}

function mappedItem(
  item: PdfNativeListTextSelectionItemEvidence,
  transform: PdfNativeTextGroupTransform,
): PdfNativeListTextSelectionItemEvidence {
  const [a, b, c, d, e, f] = item.transform;
  const [ta, tb, tc, td] = groupLinear(transform);
  const mapped = [
    a * ta + b * tc,
    a * tb + b * td,
    c * ta + d * tc,
    c * tb + d * td,
    e * ta + f * tc + transform.e,
    e * tb + f * td + transform.f,
  ] as const;
  const along = Math.hypot(a, b), across = Math.hypot(c, d);
  return Object.freeze({
    ...item,
    transform: Object.freeze([...mapped]) as PdfNativeListTextSelectionItemEvidence["transform"],
    width: along > 0 ? (item.width * Math.hypot(mapped[0], mapped[1])) / along : item.width,
    height: across > 0 ? (item.height * Math.hypot(mapped[2], mapped[3])) / across : item.height * td,
  });
}

function groupBaselineDraft(
  evidence: PdfNativeListTextFidelityEvidence,
  appearance: readonly PdfNativeSourceTextAppearanceRun[] | undefined,
): PdfNativeTextPropertyDraft {
  const run = appearance?.find((entry) => entry.text.trim()) ?? appearance?.[0];
  return Object.freeze({
    fontResourceName: evidence.resourceName,
    fontSize: evidence.fontSize,
    color:
      run?.color ??
      Object.freeze({ kind: "gray" as const, components: Object.freeze([0]) as readonly [number] }),
    alignment: "left" as const,
    characterSpacing: run?.characterSpacing ?? 0,
    wordSpacing: run?.wordSpacing ?? 0,
    lineSpacing: evidence.lineSpacing,
    leftIndent: 0,
    rightIndent: 0,
    firstLineIndent: 0,
    horizontalScale: run?.horizontalScale ?? 100,
    textRise: 0,
    textMatrixLinear: Object.freeze([1, 0, 0, 1]) as readonly [number, number, number, number],
  });
}

/**
 * Whether a proven box can be moved and resized as a group, read off the
 * proof it opens with; null when it cannot. It grants nothing: the move is
 * planned again, from its own proof, when it is previewed or saved.
 */
export async function proveMraiNativeTextGroupMove(
  sourceBytes: Uint8Array,
  proof: Readonly<{
    request?: PdfNativeContentStreamTransactionRequest;
    target?: Readonly<{ objectNumber: number; generation: number }>;
    evidence: PdfNativeListTextFidelityEvidence;
    sourceAppearance?: readonly PdfNativeSourceTextAppearanceRun[];
    consolidated?: boolean;
  }>,
  /** Whether the group's weight is offered too (a box no single-run
   * property proof covers: its bold is the group's). */
  options: Readonly<{ bold?: boolean }> = {},
): Promise<PdfNativeTextGroupMoveProof | null> {
  const probed = await probeGroup(sourceBytes, proof, options.bold === true);
  if (typeof probed === "string") return null;
  const baseline = groupBaselineDraft(proof.evidence, proof.sourceAppearance);
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION,
    baseline: probed.bold === undefined ? baseline : Object.freeze({ ...baseline, syntheticBold: probed.bold }),
  });
}

/** Why a proven box cannot be moved as a group, or null when it can: the
 * refusal proveMraiNativeTextGroupMove reads (for diagnosis). */
export async function mraiNativeTextGroupMoveRefusal(
  sourceBytes: Uint8Array,
  proof: Readonly<{
    request?: PdfNativeContentStreamTransactionRequest;
    target?: Readonly<{ objectNumber: number; generation: number }>;
    evidence: PdfNativeListTextFidelityEvidence;
    consolidated?: boolean;
  }>,
): Promise<string | null> {
  const probed = await probeGroup(sourceBytes, proof, false);
  return typeof probed === "string" ? probed : null;
}

/**
 * Whether a box moves as a group (else why not), and its weight when asked:
 * a group whose every show is filled (`0 Tr`) takes a synthetic bold, one
 * whose every show is stroked in place (`2 Tr`: WPS's and Word's own bold of
 * a face with none) has it taken off -- each where that bold derives (the
 * stroke width in force known, every part one a show can be cut at).
 */
async function probeGroup(
  sourceBytes: Uint8Array,
  proof: Readonly<{
    request?: PdfNativeContentStreamTransactionRequest;
    target?: Readonly<{ objectNumber: number; generation: number }>;
    evidence: PdfNativeListTextFidelityEvidence;
    consolidated?: boolean;
  }>,
  weight: boolean,
): Promise<string | Readonly<{ bold?: boolean }>> {
  let input: Parameters<typeof deriveMraiNativeTextGroupMoveRewrite>[0];
  let pageNumber: number;
  let modes: readonly number[];
  let showFonts: readonly string[];
  try {
    const target = proof.request?.identity ?? proof.target;
    if (!target) return "The box names no content stream.";
    const resolved = await resolveGroup(sourceBytes, { evidence: proof.evidence, target, request: proof.request,
      consolidated: proof.consolidated });
    pageNumber = resolved.streams[resolved.index].pageNumber;
    input = {
      prefix: resolved.streams.slice(0, resolved.index).map((stream) => LATIN1.decode(stream.streamBytes)),
      source: LATIN1.decode(resolved.streams[resolved.index].streamBytes),
      shows: resolved.shows,
      transform: { a: 1, d: 1, e: 1, f: 1 },
      continued: resolved.index < resolved.streams.length - 1,
    };
    try {
      ({ modes, showFonts } = deriveMraiNativeTextGroupMoveRewrite(input));
    } catch (error) {
      // Not with its lines' white-space shows: as the shows its proof names.
      if (!resolved.named || !(error instanceof PdfNativeTransactionError)) throw error;
      input = { ...input, shows: resolved.named };
      ({ modes, showFonts } = deriveMraiNativeTextGroupMoveRewrite(input));
    }
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
  if (!weight || modes.length === 0 || !modes.every((mode) => mode === 0 || mode === 2)) return Object.freeze({});
  // A line bolded with its bold word left alone (planMraiNativeTextGroupMove)
  // reads bold: its other shows are stroked, that word is in a bold face.
  const stroked = modes.some((mode) => mode === 2);
  const faces = stroked && !modes.every((mode) => mode === 2)
    ? await readGroupBoldFaces(sourceBytes, pageNumber, showFonts) : null;
  const bold = modes.every((mode, index) => mode === 2 || (faces?.[index] ?? false));
  try {
    deriveMraiNativeTextGroupMoveRewrite({
      ...input,
      transform: { a: 1, d: 1, e: 0, f: 0 },
      bold: { on: !bold },
      ...(bold ? {} : {
        lineWidths: await readGroupLineWidths(sourceBytes, pageNumber,
          mraiNativeTextGroupExtGStateNames([...input.prefix, input.source])),
      }),
    });
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return Object.freeze({});
  }
  return Object.freeze({ bold });
}




function asBytes(text: string): Uint8Array {
  return Uint8Array.from(text, (character) => character.charCodeAt(0) & 0xff);
}

/** An edit's replacement as the bytes it writes and, by their place in the
 * span, the bytes it keeps: `head`, the span's `restStart..restEnd`, then
 * `tail` (an empty one left out). */
function outsidePieces(head: string, restStart: number, restEnd: number, tail: string): readonly PdfNativeReplacementPiece[] {
  return Object.freeze([
    ...(head ? [Object.freeze({ literal: asBytes(head) })] : []),
    Object.freeze({ spanStart: restStart, spanEnd: restEnd }),
    ...(tail ? [Object.freeze({ literal: asBytes(tail) })] : []),
  ]);
}








async function groupOwnership(
  sourceBytes: Uint8Array,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<GroupOwnership> {
  try {
    const proof = await proveMraiNativeListTextSelectionForWriter({ sourceBytes, selection });
    return Object.freeze({ evidence: proof.evidence, target: proof.request.identity, request: proof.request });
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError) || error.code !== "unsupported-document") throw error;
    // The proof the box opened with (the read path's, which may prove what
    // the writer's cannot re-prove) owns its shows as well.
    const read = await prepareMraiNativeTextEditingSelection({ sourceBytes, selection });
    return Object.freeze({ evidence: read.proof.evidence, target: read.proof.target,
      consolidated: Boolean(read.proof.sourceRun) });
  }
}

/**
 * Plan one native group move of the box a selection proves. The candidate is
 * prepared and reopened; the moved box's selection is bound to it.
 */
export async function planMraiNativeTextGroupMove(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
  /** The map, and the weight set with it (a bold alone keeps the map the
   * identity). */
  readonly transform: PdfNativeTextGroupEdit;
  /** The right edge (the page's user space) of the frame the moved
   * paragraph is set again in after the move: a clip that would cut its
   * lines there has the box written outside every clip. */
  readonly frameRight?: number;
}): Promise<PdfNativeTextGroupMovePlan> {
  const edit = readMraiNativeTextGroupTransform(input.transform, input.transform.bold);
  const transform: PdfNativeTextGroupTransform = Object.freeze({ a: edit.a,
    ...(edit.b ? { b: edit.b } : {}), ...(edit.c ? { c: edit.c } : {}), d: edit.d, e: edit.e, f: edit.f });
  const ownership = await groupOwnership(input.sourceBytes, input.selection);
  const resolved = await resolveGroup(input.sourceBytes, ownership);
  try {
    return await planResolvedGroupMove(input, transform, edit.bold, ownership, resolved, resolved.shows);
  } catch (error) {
    // The box does not move with its lines' white-space shows (one drawn in
    // a mode or under a mark the rewrite refuses): it moves as the shows its
    // proof names, as it did before they were read.
    if (!resolved.named || !(error instanceof PdfNativeTransactionError)) throw error;
    return planResolvedGroupMove(input, transform, edit.bold, ownership, resolved, resolved.named);
  }
}

/** The move of a resolved box as `shows`: every show it owns, or the shows
 * its proof names alone (planMraiNativeTextGroupMove). */
async function planResolvedGroupMove(
  input: Parameters<typeof planMraiNativeTextGroupMove>[0],
  transform: PdfNativeTextGroupTransform,
  requestedBold: PdfNativeTextGroupBold | undefined,
  ownership: GroupOwnership,
  resolved: ResolvedGroup,
  shows: readonly GroupShow[],
): Promise<PdfNativeTextGroupMovePlan> {
  let bold = requestedBold;
  const stream = resolved.streams[resolved.index];
  const source = LATIN1.decode(stream.streamBytes);
  const prefix = resolved.streams.slice(0, resolved.index).map((entry) => LATIN1.decode(entry.streamBytes));
  const continued = resolved.index < resolved.streams.length - 1;
  if (bold?.on && !bold.shows) {
    // A bold put on a line of mixed weights strokes only its shows in a
    // regular face: a word the page draws in a bold face keeps its weight,
    // as WPS 加粗 on 「Quarterly Report to Shareholders」 (its first word in
    // Calibri-Bold) bolds the rest. Stroked as well, that word was drawn
    // heavier than any bold of its face.
    const read = deriveMraiNativeTextGroupMoveRewrite({ prefix, source, shows,
      transform: { a: 1, d: 1, e: 1, f: 1 }, continued });
    const faces = await readGroupBoldFaces(input.sourceBytes, stream.pageNumber, read.showFonts);
    if (faces.length === shows.length && faces.some(Boolean) && !faces.every(Boolean)) {
      bold = Object.freeze({ on: true, shows: Object.freeze(shows.filter((_, index) => !faces[index])) });
    }
  }
  // A bold put on puts back the stroke width in force, which a `gs` may set.
  const lineWidths = bold?.on
    ? await readGroupLineWidths(input.sourceBytes, stream.pageNumber, mraiNativeTextGroupExtGStateNames([...prefix, source]))
    : undefined;
  const weight = bold ? { bold, ...(lineWidths ? { lineWidths } : {}) } : {};
  let rewrite = deriveMraiNativeTextGroupMoveRewrite({ prefix, source, shows, transform, continued,
    ...weight });
  const selection = ownership.evidence.selection!;
  const readItems = paragraphReadItems(ownership.evidence, true);
  // The landing keeps every clip the box is drawn under, and the page.
  const page = await readMraiNativeVisiblePageBounds(input.sourceBytes, selection.pageNumber);
  const pageClips: Clip[] = page ? [[page.left, page.bottom, page.right, page.top]] : [];
  const bounds = selectionBounds(selection);
  const visible = landingVisible([...rewrite.clips, ...pageClips], bounds, transform);
  const last = resolved.index === resolved.streams.length - 1;
  if (!visible || (last && !landingKeepsWidth(rewrite.clips, bounds, transform, input.frameRight))) {
    // A clip of the box's own would hide it where the page shows it: the box
    // is written outside every clip, at the tail of the page's last stream,
    // where that is the same drawing. Where only a sliver of its width would
    // go and it cannot be written so, it lands under its clips as before.
    const inClip = rewrite;
    try {
      assertMraiNativeTextGroupLandingVisible(pageClips, bounds, transform);
      if (!last) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.landing);
      rewrite = deriveMraiNativeTextGroupMoveRewrite({ prefix, source, shows, transform, outsideClip: true,
        ...weight });
      assertMraiNativeTextGroupLandingVisible([...rewrite.clips, ...pageClips], bounds, transform);
      await assertMraiNativeTextGroupOutsideResources(
        input.sourceBytes, selection.pageNumber, rewrite.outside!.fonts, rewrite.outside!.extGStates);
    } catch (error) {
      // With its lines' white-space shows the box is not written outside
      // the clip (one of them under a mark a copy cannot carry): it is
      // planned again as the shows its proof names, which may be.
      if (!visible || !(error instanceof PdfNativeTransactionError) || shows !== (resolved.named ?? shows)) throw error;
      rewrite = inClip;
    }
  }
  const outside = rewrite.outside;
  // An edit outside the clip carries only the bytes it writes: the rest of
  // the stream it runs over is named by its place.
  const literalLength = outside ? outside.middle.length + outside.block.length : rewrite.replacement.length;
  if (literalLength > MAX_GROUP_REPLACEMENT_BYTES) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.size);
  const operatorSpan = Object.freeze({
    start: rewrite.span.start,
    end: rewrite.span.end,
    expectedSha256: await sha256PdfBytes(stream.streamBytes.slice(rewrite.span.start, rewrite.span.end)),
  });
  // The moved box is wrapped in a box mark of its own, so the page names it
  // one box where it lands (box-mark.ts): a box moved in place, whole shows
  // only, with no weight set with it. A move the mark would cost is written
  // without one, as before.
  const boxMark = !outside && !bold && shows.every((show) => show.codeStart === undefined)
    ? planMraiNativeBoxMark({
        streamBytes: stream.streamBytes,
        operatorSpan: rewrite.span,
        replacement: asBytes(rewrite.replacement),
        boxShows: [...shows, ...(resolved.duplicates ?? [])],
      })
    : null;
  const write = async (mark: PdfNativeBoxMark | null) => {
    const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
      identity: stream.identity,
      expectedObjectSha256: stream.objectSha256,
      expectedStreamSha256: stream.streamSha256,
      operatorSpan,
      replacement: asBytes(rewrite.replacement),
      ...(outside
        ? {
            replacementPieces: outsidePieces(
              outside.middle, outside.middleEnd - rewrite.span.start, rewrite.span.end - rewrite.span.start, outside.block),
          }
        : {}),
      ...(mark ? { boxMark: mark } : {}),
      // A joined page's proof pins the joined drawing; the move writes one
      // stream of it, with the page and its sibling streams pinned.
      untouchedObjects: ownership.request && !ownership.request.pageContentConsolidation
        ? ownership.request.untouchedObjects ?? []
        : stream.untouchedObjects,
    });
    const prepared = await prepareMraiNativeContentStreamTransaction(input.sourceBytes, request);
    try {
      await verifyMraiNativeContentStreamCandidate(input.sourceBytes, prepared);
    } catch (error) {
      prepared.candidateBytes.fill(0);
      throw error;
    }
    return { request, prepared, mark };
  };
  let written: Awaited<ReturnType<typeof write>>;
  try {
    written = await write(boxMark);
  } catch (error) {
    if (!boxMark || !(error instanceof PdfNativeTransactionError)) throw error;
    written = await write(null);
  }
  const { request, prepared } = written;
  // A box of several items is their sum, as a selection states it (the
  // writer checks it exactly): each width mapped apart summed to another
  // last digit than the box's width mapped (229.10487127753436 against
  // ...344), and the moved paragraph could not be set again in a frame
  
  const movedItems = selection.items
    ? Object.freeze(selection.items.map((item) => mappedItem(item, transform))) : null;
  const movedItem = mappedItem(selection.item, transform);
  const movedSelection: PdfNativeListTextSelectionEvidence = Object.freeze({
    ...selection,
    sourceSha256: prepared.candidateSha256,
    item: movedItems
      ? Object.freeze({ ...movedItem, width: movedItems.reduce((total, entry) => total + entry.width, 0) })
      : movedItem,
    ...(movedItems ? { items: movedItems } : {}),
  });
  const evidence: PdfNativeTextGroupMoveEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION,
    sourceSha256: stream.identity.sourceSha256,
    pageNumber: stream.pageNumber,
    page: Object.freeze({ ...stream.page }),
    contentStream: Object.freeze({ ...stream.contentStream }),
    operatorSpan,
    transform,
    shows,
    // A read proof of items alone is written as those items' own proofs,
    // which the record carries as they are.
    ...(readItems
      ? { sourceFidelity: readItems[0], sourceItems: Object.freeze([...readItems]) }
      : { sourceFidelity: ownership.evidence }),
    ...(ownership.request ? { sourceProofReplacement: Object.freeze([...ownership.request.replacement]) } : {}),
    ...(resolved.duplicates ? { sourceDuplicateSpans: resolved.duplicates } : {}),
    ...(outside ? { outsideClip: true as const } : {}),
    ...(bold ? { bold } : {}),
    ...(written.mark ? { boxMark: written.mark } : {}),
  });
  return Object.freeze({
    evidence,
    request,
    prepared,
    movedSelection,
    originalBytes: stream.streamBytes.slice(rewrite.span.start, rewrite.span.end),
    proofSelection: selection,
  });
}

/**
 * The undo of a group move: the source span's own bytes written back where
 * the move wrote its rewrite. It is proven by deriving the move again from
 * the stream it restores: that must give the bytes it replaces.
 */
export async function planMraiNativeTextGroupMoveRestore(input: {
  readonly sourceBytes: Uint8Array;
  readonly origin: PdfNativeTextGroupMoveEvidence;
  readonly originTransactionId: string;
  /** The source span's bytes, as the move read them. */
  readonly originalBytes: Uint8Array;
  /** The bytes the move wrote in their place. */
  readonly rewrittenBytes: Uint8Array;
}): Promise<Readonly<{
  evidence: PdfNativeTextGroupMoveEvidence;
  request: PdfNativeContentStreamTransactionRequest;
  prepared: PdfNativePreparedContentStreamTransaction;
}>> {
  const { origin } = input;
  if ((await sha256PdfBytes(input.originalBytes)) !== origin.operatorSpan.expectedSha256) {
    invalid("The restored bytes are not the ones the move replaced.");
  }
  const streams = await inspectMraiNativePageContentStreams(input.sourceBytes, origin.pageNumber);
  const index = streams.findIndex(
    (stream) =>
      stream.contentStream.objectNumber === origin.contentStream.objectNumber &&
      stream.contentStream.generation === origin.contentStream.generation,
  );
  if (index < 0) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.source);
  const stream = streams[index];
  // A move that marked its box wrote its bytes after the mark's opening
  // half; the restore writes the source bytes back there and keeps the mark.
  const start = origin.operatorSpan.start + (origin.boxMark ? MRAI_NATIVE_BOX_MARK_OPEN.length : 0);
  const end = start + input.rewrittenBytes.length;
  const current = LATIN1.decode(stream.streamBytes);
  const rewritten = LATIN1.decode(input.rewrittenBytes);
  if (end > current.length || current.slice(start, end) !== rewritten) {
    refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.restore);
  }
  // A box written outside the clip was written at the tail of the page's
  // last stream, and is put back from there.
  if (origin.outsideClip && index !== streams.length - 1) refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.restore);
  const original = LATIN1.decode(input.originalBytes);
  const prefix = streams.slice(0, index).map((entry) => LATIN1.decode(entry.streamBytes));
  // A bold the move put on is derived again with the widths the page's
  // states set (its resources are the move's own: only the stream changed).
  const lineWidths = origin.bold?.on
    ? await readGroupLineWidths(input.sourceBytes, origin.pageNumber,
      mraiNativeTextGroupExtGStateNames([...prefix, current.slice(0, start) + original + current.slice(end)]))
    : undefined;
  const derived = verifyMraiNativeTextGroupMoveRestore({
    prefix,
    current,
    span: { start, end },
    original,
    origin,
    continued: index < streams.length - 1,
    ...(lineWidths ? { lineWidths } : {}),
  });
  const operatorSpan = Object.freeze({
    start,
    end,
    expectedSha256: await sha256PdfBytes(input.rewrittenBytes),
  });
  // Outside the clip, the restore writes the text objects' own bytes back
  // and keeps the rest of the stream by its place, without the block.
  const outside = derived.outside;
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: stream.identity,
    expectedObjectSha256: stream.objectSha256,
    expectedStreamSha256: stream.streamSha256,
    operatorSpan,
    replacement: Uint8Array.from(input.originalBytes),
    ...(outside
      ? {
          replacementPieces: outsidePieces(
            original.slice(0, outside.middleEnd - start),
            outside.middle.length,
            outside.middle.length + (origin.operatorSpan.end - outside.middleEnd),
            "",
          ),
        }
      : {}),
    untouchedObjects: stream.untouchedObjects,
  });
  const prepared = await prepareMraiNativeContentStreamTransaction(input.sourceBytes, request);
  try {
    await verifyMraiNativeContentStreamCandidate(input.sourceBytes, prepared);
  } catch (error) {
    prepared.candidateBytes.fill(0);
    throw error;
  }
  const evidence: PdfNativeTextGroupMoveEvidence = Object.freeze({
    ...origin,
    sourceSha256: stream.identity.sourceSha256,
    restore: Object.freeze({ originTransactionId: input.originTransactionId, direction: "undo" as const }),
  });
  return Object.freeze({ evidence, request, prepared });
}

/**
 * A restore is the move's own source: the stream with the original span put
 * back derives, by the same rule, exactly the bytes the restore replaces
 * (mirror of _validate_text_group_restore). Nothing but positioning differs
 * between the two, so a restore can bring back no glyph the move did not
 * carry.
 */
export function verifyMraiNativeTextGroupMoveRestore(input: {
  readonly prefix: readonly string[];
  readonly current: string;
  readonly span: Span;
  readonly original: string;
  readonly origin: PdfNativeTextGroupMoveEvidence;
  /** The stroke widths the page's `gs` names set, for a bold put on. */
  readonly lineWidths?: ReadonlyMap<string, string | null>;
  /** The page draws on in a stream after this one. */
  readonly continued?: boolean;
}): PdfNativeTextGroupMoveRewrite {
  const { origin, span } = input;
  const restored = input.current.slice(0, span.start) + input.original + input.current.slice(span.end);
  // The box mark's opening half stands before every show the move named.
  const shift = origin.boxMark ? MRAI_NATIVE_BOX_MARK_OPEN.length : 0;
  const derived = deriveMraiNativeTextGroupMoveRewrite({
    prefix: input.prefix,
    source: restored,
    shows: shift
      ? origin.shows.map((show) => Object.freeze({ ...show, start: show.start + shift, end: show.end + shift }))
      : origin.shows,
    transform: origin.transform,
    outsideClip: origin.outsideClip === true,
    ...(origin.bold ? { bold: origin.bold } : {}),
    ...(input.lineWidths ? { lineWidths: input.lineWidths } : {}),
    ...(input.continued ? { continued: true } : {}),
  });
  if (
    derived.span.start !== origin.operatorSpan.start + shift ||
    derived.span.end !== origin.operatorSpan.end + shift ||
    derived.replacement !== input.current.slice(span.start, span.end)
  ) {
    refuse(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.restore);
  }
  return derived;
}
