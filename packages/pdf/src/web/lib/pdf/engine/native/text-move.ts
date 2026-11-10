/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Move one reader item's own text on the page, natively.
 *
 * A move is a translation of the text as the page draws it: the same glyph
 * shows, font, size, spacing and colour, set somewhere else. A text object
 * places its text by a few operators that set the text matrix afresh -- a
 * `Tm`, or the first `Td` after `BT` -- and every later `Td`, `TD`, `T*` and
 * show only moves on from there. So the move rewrites the translation those
 * operators carry (`Tm`'s e and f, `Td`'s tx and ty) and nothing else; every
 * glyph show keeps its bytes. A producer that sets each glyph with its own
 * `Td` (Qt writes every item as `Tm`, `Td <glyph> Tj`, `Td <glyph> Tj`, ...)
 * moves whole, because only the first placement is absolute.
 *
 * The move owns a segment of a text object: the operators from one placing
 * operator to the next `Tm` (or `ET`). It is isolated only when every show of
 * each segment it rewrites is one of the item's own shows, so no other text
 * moves with it; a later segment starts from its own `Tm` and stays. A page
 * that draws the item twice for a fake bold (Qt's overprint pass under an
 * empty /ActualText span) moves both passes by the same amount, as one unit.
 *
 * Numbers are written in micro-units with one rounding rule on both sides
 * (mirror of _text_move_micro / _format_text_move_micro in
 * native_surgical_transaction.py), so replay derives the same bytes and an
 * undo by the opposite move writes the original values back.
 */
import { LATIN1_EXACT } from "./latin1";
import { invalid, readMraiNativeTextMoveDraft } from "./text-move-model";
export { readMraiNativeTextMoveDraft };
import {
  MRAI_NATIVE_TEXT_MOVE_KIND,
  MRAI_NATIVE_TEXT_MOVE_REFUSALS,
  MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION,
} from "./vocabulary";
export {
  MRAI_NATIVE_TEXT_MOVE_KIND,
  MRAI_NATIVE_TEXT_MOVE_REFUSALS,
  MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION,
};
import {
  MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION,
  proveMraiNativeListTextSelectionForWriter,
  qtFakeBoldTwinSpan,
  type PdfNativeListTextFidelityEvidence,
  type PdfNativeListTextSelectionEvidence,
  pageAlignedClipBox,
  pageAlignedPolygonClipBox,
  type PdfClipBox,
  type PdfNativeListTextSelectionItemEvidence,
  type PdfNativeSourceTextAppearanceRun,
  type PdfNativeTextPropertyDraft,
} from "./list-text-fidelity";
import {
  enclosingTextObject,
  mraiNativeContentTokens,
} from "./text-object-delete";
import { planMraiNativeBoxMark } from "./box-mark";
import {
  inspectMraiNativePageContentStreams,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativePageContentStreamInspection,
  type PdfNativePreparedContentStreamTransaction,
} from "./content-stream-transaction";

/** Record kind of a move; it rides the text edit command (its action). */

/** At most this many placing operators are rewritten by one move. */
const MAX_TEXT_MOVE_ANCHORS = 256;
/** A move farther than this, in points, is not a drag on any page. */
const MAX_TEXT_MOVE_DISTANCE = 100_000;
/** Written operands stay within the range every writer here accepts. */
const MAX_TEXT_MOVE_MICRO = 1_000_000 * 1_000_000;


type Span = Readonly<{ start: number; end: number }>;
type Matrix = readonly [number, number, number, number, number, number];

/** What the editor needs to offer a native move of a proven item. */
export interface PdfNativeTextMoveProof {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION;
  /**
   * The starting point of a move draft: the item's own state as the page
   * draws it. A move changes none of it; the draft only carries the move.
   */
  readonly baseline: PdfNativeTextPropertyDraft;
}

export interface PdfNativeTextMoveEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  /** The requested translation in PDF user space (points, y up). */
  readonly move: Readonly<{ x: number; y: number }>;
  /** The no-op proof of the moved item: which shows it is. */
  readonly sourceFidelity: PdfNativeListTextFidelityEvidence;
  /** The proof's own replacement, which replay proves the item with. */
  readonly sourceProofReplacement: readonly number[];
  /** The fake-bold overprint pass the proof named as the item's copy. */
  readonly sourceDuplicateSpans?: readonly Span[];
}

export interface PdfNativeTextMovePlan {
  readonly evidence: PdfNativeTextMoveEvidence;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  /** The page-space translation the glyphs actually move by. */
  readonly displacement: Readonly<{ x: number; y: number }>;
  /** The selection as the reader reads the moved text, bound to the output. */
  readonly movedSelection: PdfNativeListTextSelectionEvidence;
}

interface TextMoveAnchor {
  readonly operator: "Tm" | "Td";
  /** The operand tokens a move rewrites: Tm's e and f, or Td's tx and ty. */
  readonly x: Span;
  readonly y: Span;
}

interface TextMoveLayout {
  readonly object: Span;
  readonly anchors: readonly TextMoveAnchor[];
}

const LATIN1 = LATIN1_EXACT;
const NUMBER_TOKEN = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;

function refuse(message: string): never {
  throw new PdfNativeTransactionError("unsupported-document", message);
}


function numberToken(source: string, token: Span): number | null {
  const text = source.slice(token.start, token.end);
  if (!NUMBER_TOKEN.test(text)) return null;
  const value = Number(text);
  return Number.isFinite(value) ? value : null;
}

/** A value in micro-units, rounded half up (mirror of _text_move_micro). */
function textMoveMicro(value: number): number {
  const micro = Math.floor(value * 1_000_000 + 0.5);
  if (!Number.isSafeInteger(micro) || Math.abs(micro) > MAX_TEXT_MOVE_MICRO) {
    invalid(MRAI_NATIVE_TEXT_MOVE_REFUSALS.distance);
  }
  return micro === 0 ? 0 : micro;
}

/** Micro-units written as a PDF number: at most six decimals, no trailing
 * zeros (mirror of _format_text_move_micro). */
export function formatMraiNativeTextMoveMicro(micro: number): string {
  if (micro === 0) return "0";
  const magnitude = Math.abs(micro);
  const whole = Math.floor(magnitude / 1_000_000);
  const fraction = magnitude - whole * 1_000_000;
  const text =
    fraction === 0
      ? `${whole}`
      : `${whole}.${String(fraction).padStart(6, "0").replace(/0+$/, "")}`;
  return micro < 0 ? `-${text}` : text;
}

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

/**
 * The page transform and text rendering mode in force at `offset` of the last
 * source: every `cm` and `Tr` before it, undone by the `Q` that closes its
 * save. A page's streams are one drawing, so the ones drawn before the target
 * are read first.
 */
function graphicsStateAt(
  sources: readonly string[],
  offset: number,
): Readonly<{ matrix: Matrix; renderingMode: number; clipBox: PdfClipBox | null }> {
  let matrix: Matrix = [1, 0, 0, 1, 0, 0];
  let renderingMode = 0;
  // The page-aligned clip box in force: a path of rectangles clipped with
  // W/W* cuts to their bounding box at most; any other path leaves that clip
  // unknown (mirror of the graphics analysis's clipBoxes).
  let clipBox: PdfClipBox | null = null;
  let pathRects: PdfClipBox[] | null = [];
  let pathPolygons: [number, number][][] = [];
  let pendingClip = false;
  const saved: { matrix: Matrix; renderingMode: number; clipBox: PdfClipBox | null }[] = [];
  for (const [index, source] of sources.entries()) {
    const end = index === sources.length - 1 ? offset : source.length;
    const tokens = mraiNativeContentTokens(source, 0, end);
    if (!tokens) refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
    let operands: Span[] = [];
    for (const token of tokens) {
      if (!token.operator) {
        operands.push(token);
        continue;
      }
      const operator = source.slice(token.start, token.end);
      if (operator === "q") saved.push({ matrix, renderingMode, clipBox });
      else if (operator === "Q") {
        const restored = saved.pop();
        if (!restored) refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
        ({ matrix, renderingMode, clipBox } = restored);
      } else if (operator === "cm") {
        const values = operands.map((operand) => numberToken(source, operand));
        if (values.length !== 6 || values.some((value) => value === null)) {
          refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
        }
        matrix = multiply(values as unknown as Matrix, matrix);
      } else if (operator === "re") {
        if (pathRects !== null) {
          const values = operands.map((operand) => numberToken(source, operand));
          const rect = values.length === 4 && values.every((value) => value !== null)
            ? pageAlignedClipBox(matrix, ...(values as [number, number, number, number]))
            : null;
          if (rect) pathRects.push(rect);
          else pathRects = null;
        }
      } else if (operator === "m" || operator === "l") {
        if (pathRects !== null) {
          const values = operands.map((operand) => numberToken(source, operand));
          const point = values.length === 2 && values.every((value) => value !== null)
            ? [values[0] as number, values[1] as number] as [number, number] : null;
          if (!point) pathRects = null;
          else if (operator === "m") pathPolygons.push([point]);
          else if (pathPolygons.length > 0) pathPolygons[pathPolygons.length - 1].push(point);
          else pathRects = null;
        }
      } else if (operator === "c" || operator === "v" || operator === "y") {
        pathRects = null;
      } else if (operator === "W" || operator === "W*") {
        pendingClip = true;
      } else if (CLIP_PATH_PAINTING_OPERATORS.has(operator)) {
        if (pathRects !== null) {
          for (const polygon of pathPolygons) {
            const rect = pageAlignedPolygonClipBox(matrix, polygon);
            if (rect) pathRects.push(rect);
            else { pathRects = null; break; }
          }
        }
        if (pendingClip && pathRects !== null && pathRects.length > 0) {
          const bounding: PdfClipBox = [
            Math.min(...pathRects.map((rect) => rect[0])), Math.min(...pathRects.map((rect) => rect[1])),
            Math.max(...pathRects.map((rect) => rect[2])), Math.max(...pathRects.map((rect) => rect[3])),
          ];
          clipBox = clipBox === null ? bounding : [
            Math.max(clipBox[0], bounding[0]), Math.max(clipBox[1], bounding[1]),
            Math.min(clipBox[2], bounding[2]), Math.min(clipBox[3], bounding[3]),
          ];
        }
        pendingClip = false;
        pathRects = [];
        pathPolygons = [];
      } else if (operator === "Tr") {
        const mode =
          operands.length === 1 ? numberToken(source, operands[0]) : null;
        if (mode === null || !Number.isInteger(mode) || mode < 0 || mode > 7) {
          refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
        }
        renderingMode = mode;
      }
      operands = [];
    }
  }
  if (!matrix.every(Number.isFinite))
    refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
  return Object.freeze({ matrix, renderingMode, clipBox });
}

const CLIP_PATH_PAINTING_OPERATORS: ReadonlySet<string> = new Set(["n", "f", "F", "f*", "B", "B*", "b", "b*", "S", "s"]);

/** The share of a reader item's box on the page that a page-aligned clip
 * box leaves visible; 1 for a box of no area. */
function visibleShare(item: PdfNativeListTextSelectionItemEvidence, clip: PdfClipBox): number {
  const [a, b, c, d, e, f] = item.transform;
  const along = Math.hypot(a, b), up = Math.hypot(c, d);
  if (!(along > 0) || !(up > 0)) return 1;
  const xs: number[] = [], ys: number[] = [];
  for (const u of [0, item.width / along]) for (const v of [0, item.height / up]) {
    xs.push(e + u * a + v * c);
    ys.push(f + u * b + v * d);
  }
  const box = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
  const area = (box[2] - box[0]) * (box[3] - box[1]);
  if (!(area > 0)) return 1;
  const width = Math.min(box[2], clip[2]) - Math.max(box[0], clip[0]);
  const height = Math.min(box[3], clip[3]) - Math.max(box[1], clip[1]);
  return width > 0 && height > 0 ? (width * height) / area : 0;
}

/**
 * The item's own shows, as the proof reads them: every run of a positioned
 * group, every line of a line plan, or the one show. A proof that holds part
 * of a show (a reader chunk) or moves other text with an edit names no show a
 * move can take whole.
 */
function provenShows(
  evidence: PdfNativeListTextFidelityEvidence,
): readonly Span[] {
  if (
    (evidence.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION &&
      evidence.schemaVersion !== MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION) ||
    !evidence.selection ||
    evidence.paragraphRewrite ||
    evidence.paragraphReflow ||
    evidence.propertyMutation ||
    evidence.positionedLines ||
    evidence.consolidatedRuns ||
    evidence.lineShift ||
    evidence.readerChunk ||
    evidence.emptiedShow ||
    evidence.positionedRestore ||
    evidence.listAnchor ||
    evidence.originalText !== evidence.replacementText ||
    evidence.originalDisplayText !== evidence.replacementDisplayText ||
    JSON.stringify(evidence.originalCodeBytes) !==
      JSON.stringify(evidence.replacementCodeBytes)
  ) {
    refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.proof);
  }
  if (evidence.positionedRunGroup) {
    return Object.freeze(
      evidence.positionedRunGroup.runs.map((run) => {
        const leaf = run.evidence;
        if (
          leaf.readerChunk ||
          leaf.emptiedShow ||
          leaf.lineShift ||
          leaf.consolidatedRuns ||
          leaf.positionedRunGroup ||
          leaf.linePlan
        )
          refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.proof);
        return Object.freeze({
          start: leaf.textOperatorSpan.start,
          end: leaf.textOperatorSpan.end,
        });
      }),
    );
  }
  if (evidence.linePlan) {
    return Object.freeze(
      evidence.linePlan.lines.map((line) =>
        Object.freeze({
          start: line.operatorSpan.start,
          end: line.operatorSpan.end,
        }),
      ),
    );
  }
  return Object.freeze([
    Object.freeze({
      start: evidence.textOperatorSpan.start,
      end: evidence.textOperatorSpan.end,
    }),
  ]);
}

/**
 * The placing operators a move of these shows rewrites (mirror of
 * _text_move_layout). The text object is cut into segments at every `Tm`;
 * a segment that `BT` opens is placed by its first `Td`, which must come
 * before its first show and before any other operator that moves the line.
 * Every segment that draws one of the moved shows is rewritten whole, so each
 * of its shows must be a moved one, drawn in a mode that clips nothing.
 */
function textMoveLayout(input: {
  readonly source: string;
  readonly renderingMode: number;
  readonly object: Span;
  readonly shows: readonly Span[];
}): TextMoveLayout {
  const { source, object } = input;
  const moved = new Set(input.shows.map((show) => `${show.start}:${show.end}`));
  const tokens = mraiNativeContentTokens(
    source,
    object.start + 2,
    object.end - 2,
  );
  if (!tokens) refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
  interface Segment {
    anchor: TextMoveAnchor | null;
    unanchorable: boolean;
    shows: { key: string; mode: number }[];
  }
  const segments: Segment[] = [
    { anchor: null, unanchorable: false, shows: [] },
  ];
  let renderingMode = input.renderingMode;
  let operands: Span[] = [];
  const numbers = (count: number): Span[] => {
    if (
      operands.length !== count ||
      operands.some((operand) => numberToken(source, operand) === null)
    ) {
      refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
    }
    return operands;
  };
  for (const token of tokens) {
    if (!token.operator) {
      operands.push(token);
      continue;
    }
    const operator = source.slice(token.start, token.end);
    const current = segments[segments.length - 1];
    const opening = segments.length === 1 && current.anchor === null;
    if (
      operator === "q" ||
      operator === "Q" ||
      operator === "cm" ||
      operator === "BT" ||
      operator === "ET"
    ) {
      refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.transform);
    } else if (operator === "Tm") {
      const values = numbers(6);
      segments.push({
        anchor: Object.freeze({ operator: "Tm", x: values[4], y: values[5] }),
        unanchorable: false,
        shows: [],
      });
    } else if (operator === "Td") {
      const values = numbers(2);
      if (opening && !current.unanchorable && current.shows.length === 0) {
        current.anchor = Object.freeze({
          operator: "Td",
          x: values[0],
          y: values[1],
        });
      }
    } else if (operator === "TD" || operator === "T*") {
      if (opening) current.unanchorable = true;
    } else if (
      operator === "Tj" ||
      operator === "TJ" ||
      operator === "'" ||
      operator === '"'
    ) {
      // `'` and `"` move to the next line before they show.
      if (opening) current.unanchorable = true;
      const start = operands.length > 0 ? operands[0].start : token.start;
      current.shows.push({ key: `${start}:${token.end}`, mode: renderingMode });
    } else if (operator === "Tr") {
      const [mode] = numbers(1);
      const value = numberToken(source, mode)!;
      if (!Number.isInteger(value) || value < 0 || value > 7)
        refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
      renderingMode = value;
    }
    operands = [];
  }
  let found = 0;
  const anchors: TextMoveAnchor[] = [];
  for (const segment of segments) {
    const owned = segment.shows.filter((show) => moved.has(show.key));
    if (owned.length === 0) continue;
    found += owned.length;
    if (owned.length !== segment.shows.length)
      refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.shared);
    if (segment.shows.some((show) => show.mode >= 4))
      refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.clip);
    if (!segment.anchor || segment.unanchorable)
      refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.anchor);
    anchors.push(segment.anchor);
  }
  if (found !== moved.size || moved.size !== input.shows.length)
    refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.object);
  if (anchors.length === 0 || anchors.length > MAX_TEXT_MOVE_ANCHORS)
    refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.anchor);
  return Object.freeze({ object, anchors: Object.freeze(anchors) });
}

/** The page's own stream that draws a proven item, its segments' placing
 * operators, and the page transform they are drawn under. */
interface ResolvedTextMove {
  readonly stream: PdfNativePageContentStreamInspection;
  readonly source: string;
  readonly layout: TextMoveLayout;
  readonly matrix: Matrix;
  readonly twin: Span | null;
  /** The page-aligned clip box in force at the text object, when one is. */
  readonly clipBox: PdfClipBox | null;
  /** Every show that moves: the item's own, and its overprint pass. */
  readonly shows: readonly Span[];
}

async function resolveTextMove(
  sourceBytes: Uint8Array,
  proof: Readonly<{
    request: PdfNativeContentStreamTransactionRequest;
    evidence: PdfNativeListTextFidelityEvidence;
  }>,
): Promise<ResolvedTextMove> {
  const { request, evidence } = proof;
  const shows = provenShows(evidence);
  if (
    request.pageContentConsolidation ||
    request.formInstanceIsolation ||
    request.pageExtGStateMutation ||
    request.pageFontExtensionMutation ||
    request.pageXObjectUnmap ||
    request.contentAppend ||
    request.replacementPieces
  ) {
    refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.source);
  }
  const streams = await inspectMraiNativePageContentStreams(
    sourceBytes,
    evidence.selection!.pageNumber,
  );
  const index = streams.findIndex(
    (stream) =>
      stream.identity.objectNumber === request.identity.objectNumber &&
      stream.identity.generation === request.identity.generation,
  );
  if (index < 0) refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.source);
  const stream = streams[index];
  if (
    stream.streamSha256 !== request.expectedStreamSha256 ||
    stream.objectSha256 !== request.expectedObjectSha256
  ) {
    throw new PdfNativeTransactionError(
      "precondition-mismatch",
      "The text's content stream changed.",
    );
  }
  const source = LATIN1.decode(stream.streamBytes);
  const span = evidence.textOperatorSpan;
  const duplicates = request.duplicateSpans ?? [];
  let twin: Span | null = null;
  if (duplicates.length > 0) {
    const expected = qtFakeBoldTwinSpan(source, span, evidence.fontSize);
    if (
      duplicates.length !== 1 ||
      !expected ||
      expected.start !== duplicates[0].start ||
      expected.end !== duplicates[0].end
    ) {
      refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.twin);
    }
    twin = Object.freeze({ start: expected.start, end: expected.end });
  }
  const object = enclosingTextObject(source, span.start);
  for (const part of twin ? [span, twin] : [span]) {
    if (part.start < object.start + 2 || part.end > object.end - 2)
      refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.object);
  }
  const state = graphicsStateAt(
    [
      ...streams.slice(0, index).map((item) => LATIN1.decode(item.streamBytes)),
      source,
    ],
    object.start,
  );
  const shift = twin ? twin.start - span.start : 0;
  const moved = twin
    ? [
        ...shows,
        ...shows.map((show) =>
          Object.freeze({ start: show.start + shift, end: show.end + shift }),
        ),
      ]
    : shows;
  const layout = textMoveLayout({
    source,
    renderingMode: state.renderingMode,
    object: Object.freeze({ start: object.start, end: object.end }),
    shows: moved,
  });
  return Object.freeze({
    stream, source, layout, matrix: state.matrix, twin, clipBox: state.clipBox, shows: Object.freeze(moved),
  });
}

/** The translation a move writes into the placing operators, in their own
 * space and in micro-units (mirror of _text_move_offset). */
function textMoveOffset(
  move: Readonly<{ x: number; y: number }>,
  matrix: Matrix,
): Readonly<{ x: number; y: number }> {
  if (
    !Number.isFinite(move.x) ||
    !Number.isFinite(move.y) ||
    Math.abs(move.x) > MAX_TEXT_MOVE_DISTANCE ||
    Math.abs(move.y) > MAX_TEXT_MOVE_DISTANCE
  ) {
    invalid(MRAI_NATIVE_TEXT_MOVE_REFUSALS.distance);
  }
  const [a, b, c, d] = matrix;
  const determinant = a * d - b * c;
  if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-9) {
    refuse("Native text move requires an invertible graphics state.");
  }
  return Object.freeze({
    x: textMoveMicro((move.x * d - move.y * c) / determinant),
    y: textMoveMicro((move.y * a - move.x * b) / determinant),
  });
}

/** The span a move rewrites and what it writes there: each placing operand
 * that the offset changes, set to its own value plus the offset; every other
 * byte, every glyph show among them, as it was. */
function textMoveRewrite(
  source: string,
  layout: TextMoveLayout,
  offset: Readonly<{ x: number; y: number }>,
): Readonly<{ span: Span; replacement: Uint8Array }> {
  const edits: { start: number; end: number; text: string }[] = [];
  for (const anchor of layout.anchors) {
    for (const [token, delta] of [
      [anchor.x, offset.x],
      [anchor.y, offset.y],
    ] as const) {
      if (delta === 0) continue;
      const value = numberToken(source, token);
      if (value === null) refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unreadable);
      edits.push({
        start: token.start,
        end: token.end,
        text: formatMraiNativeTextMoveMicro(textMoveMicro(value) + delta),
      });
    }
  }
  if (edits.length === 0) invalid(MRAI_NATIVE_TEXT_MOVE_REFUSALS.unchanged);
  edits.sort((left, right) => left.start - right.start);
  const start = edits[0].start;
  const end = edits[edits.length - 1].end;
  let written = "";
  let cursor = start;
  for (const edit of edits) {
    written += source.slice(cursor, edit.start) + edit.text;
    cursor = edit.end;
  }
  written += source.slice(cursor, end);
  return Object.freeze({
    span: Object.freeze({ start, end }),
    replacement: Uint8Array.from(
      written,
      (character) => character.charCodeAt(0) & 0xff,
    ),
  });
}

function movedItem(
  item: PdfNativeListTextSelectionItemEvidence,
  displacement: Readonly<{ x: number; y: number }>,
): PdfNativeListTextSelectionItemEvidence {
  const [a, b, c, d, e, f] = item.transform;
  return Object.freeze({
    ...item,
    transform: Object.freeze([
      a,
      b,
      c,
      d,
      e + displacement.x,
      f + displacement.y,
    ]) as PdfNativeListTextSelectionItemEvidence["transform"],
  });
}

function moveBaselineDraft(
  evidence: PdfNativeListTextFidelityEvidence,
  appearance: readonly PdfNativeSourceTextAppearanceRun[] | undefined,
): PdfNativeTextPropertyDraft {
  const run = appearance?.find((entry) => entry.text.trim()) ?? appearance?.[0];
  return Object.freeze({
    fontResourceName: evidence.resourceName,
    fontSize: evidence.fontSize,
    color:
      run?.color ??
      Object.freeze({
        kind: "gray" as const,
        components: Object.freeze([0]) as readonly [number],
      }),
    alignment: "left" as const,
    characterSpacing: run?.characterSpacing ?? 0,
    wordSpacing: run?.wordSpacing ?? 0,
    lineSpacing: evidence.lineSpacing,
    leftIndent: 0,
    rightIndent: 0,
    firstLineIndent: 0,
    horizontalScale: run?.horizontalScale ?? 100,
    textRise: 0,
    textMatrixLinear: Object.freeze([1, 0, 0, 1]) as readonly [
      number,
      number,
      number,
      number,
    ],
  });
}



/**
 * Whether a proven item can be moved natively, read off the proof it opens
 * with; null when it cannot. It grants nothing: the move is planned again,
 * from its own proof, when it is previewed or saved. Opening the box never
 * depends on it, so any failure here only leaves the move to the managed
 * text box.
 */
export async function proveMraiNativeTextMove(
  sourceBytes: Uint8Array,
  proof: Readonly<{
    request: PdfNativeContentStreamTransactionRequest;
    evidence: PdfNativeListTextFidelityEvidence;
    sourceAppearance?: readonly PdfNativeSourceTextAppearanceRun[];
  }>,
): Promise<PdfNativeTextMoveProof | null> {
  try {
    await resolveTextMove(sourceBytes, proof);
  } catch {
    return null;
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION,
    baseline: moveBaselineDraft(proof.evidence, proof.sourceAppearance),
  });
}

/**
 * Plan one native move of the text a selection proves, by `move` points in
 * PDF user space. The candidate is prepared and reopened, and the moved text
 * is proven again where it lands, so the same box can be opened (and the move
 * undone) from there.
 */
export async function planMraiNativeTextMove(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly move: Readonly<{ x: number; y: number }>;
}): Promise<PdfNativeTextMovePlan> {
  const move = Object.freeze({ x: input.move.x, y: input.move.y });
  if (
    !Number.isFinite(move.x) ||
    !Number.isFinite(move.y) ||
    Math.abs(move.x) > MAX_TEXT_MOVE_DISTANCE ||
    Math.abs(move.y) > MAX_TEXT_MOVE_DISTANCE
  ) {
    invalid(MRAI_NATIVE_TEXT_MOVE_REFUSALS.distance);
  }
  const proof = await proveMraiNativeListTextSelectionForWriter({
    sourceBytes: input.sourceBytes,
    selection: input.selection,
  });
  const resolved = await resolveTextMove(input.sourceBytes, proof);
  const offset = textMoveOffset(move, resolved.matrix);
  const { span, replacement } = textMoveRewrite(
    resolved.source,
    resolved.layout,
    offset,
  );
  const stream = resolved.stream;
  const operatorSpan = Object.freeze({
    start: span.start,
    end: span.end,
    expectedSha256: await sha256PdfBytes(
      stream.streamBytes.slice(span.start, span.end),
    ),
  });
  const [a, b, c, d] = resolved.matrix;
  const displacement = Object.freeze({
    x: (offset.x * a + offset.y * c) / 1_000_000,
    y: (offset.x * b + offset.y * d) / 1_000_000,
  });
  const selection = proof.evidence.selection!;
  // The clip in force at the text object does not move with it: text that
  // lands under it is drawn no more (an invoice's margin clip). An item the
  // clip already cuts may not lose more of itself (mirror of replay).
  if (resolved.clipBox) {
    const clip = resolved.clipBox;
    for (const item of selection.items ?? [selection.item]) {
      if (visibleShare(movedItem(item, displacement), clip) + 0.001 < Math.min(1, visibleShare(item, clip))) {
        refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.clipLanding);
      }
    }
  }
  const movedSelection: PdfNativeListTextSelectionEvidence = Object.freeze({
    ...selection,
    item: movedItem(selection.item, displacement),
    ...(selection.items
      ? {
          items: Object.freeze(
            selection.items.map((item) => movedItem(item, displacement)),
          ),
        }
      : {}),
  });
  // The moved text is wrapped in a box mark of its own, so the page names
  // it one box where it lands (box-mark.ts); a move the mark would cost is
  // written without one, as before.
  const boxMark = planMraiNativeBoxMark({
    streamBytes: stream.streamBytes,
    operatorSpan: span,
    replacement,
    boxShows: resolved.shows,
  });
  const write = async (mark: typeof boxMark) => {
    const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
      identity: stream.identity,
      expectedObjectSha256: stream.objectSha256,
      expectedStreamSha256: stream.streamSha256,
      operatorSpan,
      replacement,
      ...(mark ? { boxMark: mark } : {}),
      untouchedObjects: proof.request.untouchedObjects ?? [],
    });
    const prepared = await prepareMraiNativeContentStreamTransaction(
      input.sourceBytes,
      request,
    );
    const landedSelection = Object.freeze({ ...movedSelection, sourceSha256: prepared.candidateSha256 });
    try {
      await verifyMraiNativeContentStreamCandidate(input.sourceBytes, prepared);
      // The same box opens where the text landed: the moved selection proves
      // the same codes there, and a move back from it is a move like this one.
      let landed: Awaited<
        ReturnType<typeof proveMraiNativeListTextSelectionForWriter>
      >;
      try {
        landed = await proveMraiNativeListTextSelectionForWriter({
          sourceBytes: prepared.candidateBytes,
          selection: landedSelection,
        });
      } catch (error) {
        if (error instanceof PdfNativeTransactionError)
          refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.landing);
        throw error;
      }
      if (
        JSON.stringify(landed.evidence.originalCodeBytes) !==
        JSON.stringify(proof.evidence.originalCodeBytes)
      ) {
        refuse(MRAI_NATIVE_TEXT_MOVE_REFUSALS.landing);
      }
    } catch (error) {
      prepared.candidateBytes.fill(0);
      throw error;
    }
    return Object.freeze({ request, prepared, landedSelection });
  };
  let written: Awaited<ReturnType<typeof write>>;
  try {
    written = await write(boxMark);
  } catch (error) {
    if (!boxMark || !(error instanceof PdfNativeTransactionError)) throw error;
    written = await write(null);
  }
  const { request, prepared } = written;
  const evidence: PdfNativeTextMoveEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION,
    sourceSha256: stream.identity.sourceSha256,
    pageNumber: stream.pageNumber,
    page: Object.freeze({ ...stream.page }),
    contentStream: Object.freeze({ ...stream.contentStream }),
    operatorSpan,
    move,
    sourceFidelity: proof.evidence,
    sourceProofReplacement: Object.freeze([...proof.request.replacement]),
    ...(proof.request.duplicateSpans
      ? {
          sourceDuplicateSpans: Object.freeze(
            proof.request.duplicateSpans.map((item) =>
              Object.freeze({ start: item.start, end: item.end }),
            ),
          ),
        }
      : {}),
  });
  return Object.freeze({
    evidence,
    request,
    prepared,
    displacement,
    movedSelection: written.landedSelection,
  });
}
