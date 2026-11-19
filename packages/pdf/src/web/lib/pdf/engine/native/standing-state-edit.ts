/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PdfNativeTransactionError } from "./content-stream-transaction";
import { mraiNativeContentTokens } from "./text-object-delete";

type Matrix = readonly [number, number, number, number, number, number];

function refuse(message: string): never {
  throw new PdfNativeTransactionError("unsupported-document", message);
}



export const MRAI_NATIVE_STANDING_STATE_UNREADABLE_MESSAGE =
  "The page's content cannot be read operator by operator; a format was not written into it.";
/** The refusal where the state to put back after the text is not known (a
 * line width a graphics state dictionary sets that the page does not state
 * plainly, a colour the content inherits). */
export const MRAI_NATIVE_STANDING_STATE_UNKNOWN_MESSAGE =
  "The line width or colour in force after this text is not stated where this writer can read it, so it could not be put back; the format was not written.";

/** The paint state a format changes. `w`, `stroke` and `fill` are spelled as
 * the page wrote them (`0.5 w`, `0 0 0 RG`, `/Cs1 cs 0.2 scn`); null: not
 * known here. */
export type StandingPaintState = Readonly<{
  tr: string;
  w: string | null;
  stroke: string | null;
  fill: string | null;
}>;
type Kind = keyof StandingPaintState;
const KINDS: readonly Kind[] = ["tr", "w", "stroke", "fill"];

/** null: no clip; a rectangle in the page's space; "unknown": a clip this
 * walk does not read as one rectangle. */
export type StandingClip = readonly [number, number, number, number] | "unknown" | null;

type Token = Readonly<{ start: number; end: number; operator: boolean }>;

export type StandingShow = Readonly<{
  /** The show's span: its first operand to its operator. */
  start: number;
  end: number;
  /** Its first operand and its operator among the walk's tokens. */
  first: number;
  operator: number;
  state: StandingPaintState;
  ctm: Matrix;
  clip: StandingClip;
}>;

export type StandingWalk = Readonly<{
  source: string;
  tokens: readonly Token[];
  /** The paint state before each token (and, last, after them all). */
  states: readonly StandingPaintState[];
  shows: readonly StandingShow[];
  /** Each painting operator in order: its token, and the state it paints under. */
  paints: readonly Readonly<{ token: number; operator: string; state: StandingPaintState }>[];
  /** How many `q` stand open at the end. */
  depth: number;
}>;

const SHOW_OPERATORS = new Set(["Tj", "TJ", "'", '"']);
const PAINT_OPERATORS = new Set(["S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "sh", "Do"]);
const PATH_END_OPERATORS = new Set(["S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n"]);
const STROKE_SPACE = new Set(["G", "RG", "K", "CS"]);
const STROKE_COLOR = new Set(["SC", "SCN"]);
const FILL_SPACE = new Set(["g", "rg", "k", "cs"]);
const FILL_COLOR = new Set(["sc", "scn"]);
/** The operators a format's state is written with (a `gs` is not one: what
 * it sets is read, never written). */
const kindOf = (operator: string): Kind | null =>
  operator === "Tr" ? "tr" : operator === "w" ? "w"
    : STROKE_SPACE.has(operator) || STROKE_COLOR.has(operator) ? "stroke"
      : FILL_SPACE.has(operator) || FILL_COLOR.has(operator) ? "fill" : null;
/** What may stand between two shows of a box without reading the state. */
const NEUTRAL_OPERATORS = new Set(["Tf", "Td", "TD", "Tm", "T*", "Tc", "Tw", "Tz", "TL", "Ts", "BT", "ET",
  "BDC", "BMC", "EMC", "MP", "DP"]);

const INITIAL: StandingPaintState = Object.freeze({ tr: "0", w: "1 w", stroke: "0 G", fill: "0 g" });
const UNKNOWN: StandingPaintState = Object.freeze({ tr: "0", w: null, stroke: null, fill: null });

const compose = (first: Matrix, then: Matrix): Matrix => [
  first[0] * then[0] + first[1] * then[2], first[0] * then[1] + first[1] * then[3],
  first[2] * then[0] + first[3] * then[2], first[2] * then[1] + first[3] * then[3],
  first[4] * then[0] + first[5] * then[2] + then[4], first[4] * then[1] + first[5] * then[3] + then[5],
];

/** A fill colour spelled as the stroke colour that is the same colour. */
export function standingStrokeSpelling(fill: string): string {
  const spelled: Record<string, string> = { g: "G", rg: "RG", k: "K", cs: "CS", sc: "SC", scn: "SCN" };
  return fill.split(" ").map((part) => spelled[part] ?? part).join(" ");
}

/**
 * The stream walked for its paint state. `prefix`: the page's streams drawn
 * before it (they set the state it starts in); `inherits`: the stream is
 * drawn in a state its page does not show (a form), so nothing is known
 * until the stream sets it. `lineWidths`: the `/LW` each `gs` name sets (null:
 * none), a name left out being one the page does not state plainly.
 */
export function walkStandingStream(input: Readonly<{
  source: string;
  prefix?: readonly string[];
  inherits?: boolean;
  lineWidths: ReadonlyMap<string, string | null>;
}>): StandingWalk {
  let state: StandingPaintState = input.inherits ? UNKNOWN : INITIAL;
  let ctm: Matrix = [1, 0, 0, 1, 0, 0];
  let clip: StandingClip = null;
  const stack: { state: StandingPaintState; ctm: Matrix; clip: StandingClip }[] = [];
  let subpaths: (readonly [number, number])[][] = [];
  let curved = false;
  let pendingClip = false;
  const set = (kind: Kind, value: string | null) => {
    if (state[kind] !== value) state = Object.freeze({ ...state, [kind]: value });
  };
  const run = (source: string, record: boolean): StandingWalk | null => {
    const tokens = mraiNativeContentTokens(source, 0, source.length);
    if (!tokens) return refuse(MRAI_NATIVE_STANDING_STATE_UNREADABLE_MESSAGE);
    const states: StandingPaintState[] = [];
    const shows: StandingShow[] = [];
    const paints: { token: number; operator: string; state: StandingPaintState }[] = [];
    let operands: number[] = [];
    for (const [index, token] of tokens.entries()) {
      if (record) states.push(state);
      if (!token.operator) {
        operands.push(index);
        continue;
      }
      const operator = source.slice(token.start, token.end);
      const texts = operands.map((at) => source.slice(tokens[at].start, tokens[at].end));
      const numbers = texts.map(Number);
      const spelled = [...texts, operator].join(" ");
      if (operator === "BI" || operator === "ID" || operator === "EI") return refuse(MRAI_NATIVE_STANDING_STATE_UNREADABLE_MESSAGE);
      if (operator === "q") stack.push({ state, ctm, clip });
      else if (operator === "Q") {
        const saved = stack.pop();
        // (A stream that closes what it never opened: read as the page draws it, from the state it began in.)
        if (saved) ({ state, ctm, clip } = saved);
      } else if (operator === "cm") {
        if (numbers.length !== 6 || numbers.some((value) => !Number.isFinite(value))) return refuse(MRAI_NATIVE_STANDING_STATE_UNREADABLE_MESSAGE);
        ctm = compose(numbers as unknown as Matrix, ctm);
      } else if (operator === "Tr") set("tr", texts.length === 1 ? texts[0] : "?");
      else if (operator === "w") set("w", texts.length === 1 && Number.isFinite(numbers[0]) ? spelled : null);
      else if (STROKE_SPACE.has(operator)) set("stroke", spelled);
      else if (STROKE_COLOR.has(operator)) {
        // (A colour in the space last named: spelled after it.)
        // (In a space a device operator named, or none named here: not spelled whole.)
        const space = state.stroke?.split(" ").slice(0, 2).join(" ");
        set("stroke", state.stroke && /^\/\S+ CS$/u.test(space ?? "") ? `${space} ${spelled}` : null);
      } else if (FILL_SPACE.has(operator)) set("fill", spelled);
      else if (FILL_COLOR.has(operator)) {
        const space = state.fill?.split(" ").slice(0, 2).join(" ");
        set("fill", state.fill && /^\/\S+ cs$/u.test(space ?? "") ? `${space} ${spelled}` : null);
      } else if (operator === "gs") {
        const width = texts.length === 1 ? input.lineWidths.get(texts[0]) : undefined;
        if (width === undefined) set("w", null);
        else if (width !== null) set("w", `${width} w`);
      } else if (operator === "m" || operator === "l") {
        if (numbers.length === 2 && numbers.every(Number.isFinite)) {
          const point = [numbers[0] * ctm[0] + numbers[1] * ctm[2] + ctm[4], numbers[0] * ctm[1] + numbers[1] * ctm[3] + ctm[5]] as const;
          if (operator === "m") subpaths.push([point]);
          else if (subpaths.length) subpaths[subpaths.length - 1].push(point);
          else curved = true;
        } else curved = true;
      } else if (operator === "re") {
        if (numbers.length === 4 && numbers.every(Number.isFinite)) {
          const [x, y, width, height] = numbers;
          subpaths.push([[x, y], [x + width, y], [x + width, y + height], [x, y + height]].map(([px, py]) =>
            [px * ctm[0] + py * ctm[2] + ctm[4], px * ctm[1] + py * ctm[3] + ctm[5]] as const));
        } else curved = true;
      } else if (operator === "c" || operator === "v" || operator === "y") curved = true;
      else if (operator === "W" || operator === "W*") pendingClip = true;
      if (SHOW_OPERATORS.has(operator) && record) {
        if (operands.length < 1) return refuse(MRAI_NATIVE_STANDING_STATE_UNREADABLE_MESSAGE);
        shows.push(Object.freeze({ start: tokens[operands[0]].start, end: token.end, first: operands[0], operator: index, state, ctm, clip }));
      }
      if (record && (SHOW_OPERATORS.has(operator) || PAINT_OPERATORS.has(operator))) paints.push({ token: index, operator, state });
      if (PATH_END_OPERATORS.has(operator)) {
        if (pendingClip) {
          const rect = !curved && subpaths.length === 1 ? axisRect(subpaths[0]) : null;
          clip = !rect || clip === "unknown" ? "unknown"
            : clip === null ? rect
              : [Math.max(clip[0], rect[0]), Math.max(clip[1], rect[1]), Math.min(clip[2], rect[2]), Math.min(clip[3], rect[3])];
        }
        pendingClip = false;
        subpaths = [];
        curved = false;
      }
      operands = [];
    }
    if (operands.length) return refuse(MRAI_NATIVE_STANDING_STATE_UNREADABLE_MESSAGE);
    if (!record) return null;
    states.push(state);
    return Object.freeze({ source, tokens, states, shows, paints, depth: stack.length });
  };
  for (const earlier of input.prefix ?? []) run(earlier, false);
  return run(input.source, true)!;
}

/** The rectangle a closed path of four corners is, its sides along the
 * page's axes; null for any other path. */
function axisRect(points: readonly (readonly [number, number])[]): readonly [number, number, number, number] | null {
  const near = (left: number, right: number) => Math.abs(left - right) < 1e-6;
  const corners = points.length === 5 && near(points[4][0], points[0][0]) && near(points[4][1], points[0][1])
    ? points.slice(0, 4) : points;
  if (corners.length !== 4) return null;
  const xs = corners.map((point) => point[0]), ys = corners.map((point) => point[1]);
  const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] as const;
  if (!(rect[2] > rect[0]) || !(rect[3] > rect[1])) return null;
  // Each corner is a corner of the box, and each side runs along one axis:
  // the four are the box's four corners, in order round it.
  for (let index = 0; index < 4; index += 1) {
    const [x, y] = corners[index], [nx, ny] = corners[(index + 1) % 4];
    if (!(near(x, rect[0]) || near(x, rect[2])) || !(near(y, rect[1]) || near(y, rect[3]))) return null;
    if (near(x, nx) === near(y, ny)) return null;
  }
  return rect;
}

/** One edit of the stream: `text` in place of `start..end` (an insertion
 * where they are equal). */
export type StandingEdit = Readonly<{ start: number; end: number; text: string }>;

/** What a show of the box is asked to be drawn under: each kind as it must
 * stand at the show, or undefined to leave it as it stands ("free": whatever
 * costs no operator, where the kind draws nothing of the show). */
export type StandingTarget = Readonly<Partial<Record<Kind, string | "free">>>;

const isSpace = (character: string | undefined) => character !== undefined && /[\x00\t\n\f\r ]/u.test(character);
const isBreak = (character: string | undefined) => character === undefined || /[\x00\t\n\f\r ()<>[\]{}/%]/u.test(character);

/**
 * The edits that draw each named show under its target and leave every
 * other painting operator of the stream under the state it had.
 *
 * The shows are taken in chains: shows of the box one after another with
 * nothing between them that reads the paint state (only text positioning,
 * text state and the operators of the four kinds themselves). Inside a chain
 * nothing but the box's shows reads the state, so each kind is simply made
 * what each show wants: where a show's want differs from what stands there,
 * the operators of that kind between it and the show before are taken away
 * and the one that sets the want is written before the show. Before a
 * chain's first show only the state operators that stand directly against it
 * are its own to change; after its last show the state the page had there is
 * put back, again by taking away the operators directly after it that no
 * longer say the right thing and writing the ones that do.
 */
export function planStandingStateEdits(
  walk: StandingWalk,
  targets: ReadonlyMap<number, StandingTarget>,
): readonly StandingEdit[] {
  const { source, tokens, states } = walk;
  const owned = walk.shows.filter((show) => targets.has(show.end));
  if (owned.length !== targets.size) refuse("A show of this text is not where the page's content puts it.");
  const edits: StandingEdit[] = [];
  const text = (token: Token) => source.slice(token.start, token.end);
  /** The state operators among tokens `from..to`, each with its operands; null
   * where anything else stands there. `neutral`: text operators may too. */
  const stateOperators = (from: number, to: number, neutral: boolean) => {
    const found: { kind: Kind; start: number; end: number }[] = [];
    let first = from;
    for (let index = from; index < to; index += 1) {
      if (!tokens[index].operator) continue;
      const kind = kindOf(text(tokens[index]));
      if (kind) found.push({ kind, start: tokens[first].start, end: tokens[index].end });
      else if (!neutral || !NEUTRAL_OPERATORS.has(text(tokens[index]))) return null;
      first = index + 1;
    }
    return first === to ? found : null;
  };
  // An operator taken away with the white space after it (before a show) or
  // before it (after one), so that what a format wrote goes as it came.
  const removeBefore = (start: number, end: number) =>
    edits.push({ start, end: isSpace(source[end]) && (start === 0 || isBreak(source[start - 1])) ? end + 1 : end, text: "" });
  const removeAfter = (start: number, end: number) =>
    edits.push({ start: isSpace(source[start - 1]) && isBreak(source[end]) ? start - 1 : start, end, text: "" });
  const chains: StandingShow[][] = [];
  for (const show of owned) {
    const chain = chains.at(-1);
    const before = chain?.at(-1);
    if (chain && before && stateOperators(before.operator + 1, show.first, true)) chain.push(show);
    else chains.push([show]);
  }
  for (const chain of chains) {
    // The operators directly before the first show, and the state before them.
    let lead = chain[0].first;
    for (;;) {
      let at = lead - 1;
      if (at < 0 || !tokens[at].operator || !kindOf(text(tokens[at]))) break;
      while (at > 0 && !tokens[at - 1].operator) at -= 1;
      lead = at;
    }
    let region = stateOperators(lead, chain[0].first, false) ?? [];
    // What stands before the show in the stream as written, and as it will be.
    let standing: StandingPaintState = states[lead];
    const written: Record<Kind, string | null> = { ...standing };
    for (const [order, show] of chain.entries()) {
      if (order > 0) {
        region = stateOperators(chain[order - 1].operator + 1, show.first, true)!;
        standing = chain[order - 1].state;
      }
      const target = targets.get(show.end)!;
      if (show.state.tr !== "0" && show.state.tr !== "2") refuse("A show of this text is not drawn filled.");
      const adds: string[] = [];
      for (const kind of KINDS) {
        const asked = target[kind];
        const want = asked === undefined ? show.state[kind] : asked === "free" ? written[kind] : asked;
        // (Untouched: the operators between stand, and say what they said.)
        if (want === show.state[kind] && written[kind] === standing[kind]) {
          written[kind] = show.state[kind];
          continue;
        }
        for (const operator of region) if (operator.kind === kind) removeBefore(operator.start, operator.end);
        if (want !== written[kind]) {
          if (want === null) return refuse(MRAI_NATIVE_STANDING_STATE_UNKNOWN_MESSAGE);
          adds.push(kind === "tr" ? `${want} Tr` : want);
        }
        written[kind] = want;
      }
      if (adds.length) edits.push({ start: show.start, end: show.start, text: `${adds.join(" ")} ` });
    }
    // After the last show: what the page's other content read there.
    const last = chain.at(-1)!;
    let trail = last.operator + 1;
    for (;;) {
      let at = trail;
      while (at < tokens.length && !tokens[at].operator) at += 1;
      if (at >= tokens.length || !kindOf(text(tokens[at]))) break;
      trail = at + 1;
    }
    const after = stateOperators(last.operator + 1, trail, false) ?? [];
    const following = states[trail];
    // Whether anything reads a kind of the state after them before it is set
    // again or given back: nothing does where a `Q` comes first (the state
    // saved before these shows returns), or an operator that sets the kind
    // whole. Put back there, the state would be written for no one.
    const read = (kind: Kind): boolean => {
      for (let index = trail; index < tokens.length; index += 1) {
        if (!tokens[index].operator) continue;
        const operator = text(tokens[index]);
        if (operator === "Q") return false;
        if (operator === "q" || SHOW_OPERATORS.has(operator) || PAINT_OPERATORS.has(operator)) return true;
        if (kind === "w" && operator === "gs") return true;
        if ((kind === "tr" && operator === "Tr") || (kind === "w" && operator === "w") ||
          (kind === "stroke" && STROKE_SPACE.has(operator)) || (kind === "fill" && FILL_SPACE.has(operator))) return false;
      }
      return true;
    };
    const puts: string[] = [];
    for (const kind of KINDS) {
      if (written[kind] === last.state[kind]) continue;
      for (const operator of after) if (operator.kind === kind) removeAfter(operator.start, operator.end);
      if (following[kind] !== written[kind] && read(kind)) {
        if (following[kind] === null) return refuse(MRAI_NATIVE_STANDING_STATE_UNKNOWN_MESSAGE);
        puts.push(kind === "tr" ? `${following[kind]} Tr` : following[kind]!);
      }
    }
    if (puts.length) edits.push({ start: last.end, end: last.end, text: ` ${puts.join(" ")}` });
  }
  return normalizeStandingEdits(edits);
}

/** Edits in stream order, none over another; what is written at one place
 * joined in the order it was asked. */
export function normalizeStandingEdits(edits: readonly StandingEdit[]): readonly StandingEdit[] {
  const ordered = edits.map((edit, order) => ({ ...edit, order })).sort((left, right) =>
    left.start - right.start || left.end - right.end || left.order - right.order);
  const out: StandingEdit[] = [];
  for (const edit of ordered) {
    const before = out.at(-1);
    if (before && edit.start < before.end) refuse("The format's edits of the page's content run into one another.");
    if (before && before.start === before.end && edit.start === edit.end && before.start === edit.start) {
      out[out.length - 1] = { start: before.start, end: before.end, text: before.text + edit.text };
    } else out.push({ start: edit.start, end: edit.end, text: edit.text });
  }
  return Object.freeze(out);
}

/** The stream with the edits made. */
export function applyStandingEdits(source: string, edits: readonly StandingEdit[]): string {
  let out = "";
  let cursor = 0;
  for (const edit of edits) {
    out += source.slice(cursor, edit.start) + edit.text;
    cursor = edit.end;
  }
  return out + source.slice(cursor);
}

/**
 * Whether the stream, edited, paints anything but the named shows under
 * another state than it did: the proof of planStandingStateEdits, read from
 * the bytes it wrote and not from how it wrote them. `expected`: by show (in
 * the order the stream draws them), what a named show must now stand under.
 * Returns what differs, or null.
 */
export function standingStateTraceDiffers(
  before: StandingWalk,
  after: StandingWalk,
  expected: ReadonlyMap<number, StandingTarget>,
): string | null {
  if (before.paints.length !== after.paints.length || before.shows.length !== after.shows.length) return "the page paints another number of things";
  if (before.depth !== after.depth) return "the graphics state is saved another number of times";
  const key = (state: StandingPaintState) => KINDS.map((kind) => state[kind] ?? "?").join("|");
  const named = new Map<number, StandingTarget>();
  for (const [order, show] of before.shows.entries()) {
    const target = expected.get(show.end);
    if (target) named.set(after.shows[order].operator, target);
  }
  for (const [order, paint] of before.paints.entries()) {
    const now = after.paints[order];
    if (paint.operator !== now.operator) return "the page paints in another order";
    const target = named.get(now.token);
    if (!target) {
      if (key(paint.state) !== key(now.state)) return `a ${paint.operator} that is not this text's is painted under another state`;
      continue;
    }
    for (const kind of KINDS) {
      const asked = target[kind];
      if (asked === "free") continue;
      if ((asked ?? paint.state[kind]) !== now.state[kind]) return `a show of this text is not drawn as asked (${kind})`;
    }
  }
  const [end, endNow] = [before.states.at(-1)!, after.states.at(-1)!];
  return key(end) === key(endNow) ? null : "the state the content ends in is another";
}

/** The edits as the pieces of one replacement of `spanStart..spanEnd`: the
 * ranges of the span that stay, and the bytes written between them. */
export function standingEditPieces(
  edits: readonly StandingEdit[],
  spanStart: number,
  spanEnd: number,
): readonly (Readonly<{ spanStart: number; spanEnd: number }> | Readonly<{ literal: Uint8Array }>)[] {
  const pieces: (Readonly<{ spanStart: number; spanEnd: number }> | Readonly<{ literal: Uint8Array }>)[] = [];
  let cursor = spanStart;
  for (const edit of edits) {
    if (edit.start < cursor || edit.end > spanEnd) refuse("The format's edits lie outside the span they are written in.");
    if (edit.start > cursor) pieces.push(Object.freeze({ spanStart: cursor - spanStart, spanEnd: edit.start - spanStart }));
    if (edit.text) pieces.push(Object.freeze({ literal: Uint8Array.from(edit.text, (character) => character.charCodeAt(0) & 0xff) }));
    cursor = edit.end;
  }
  if (spanEnd > cursor) pieces.push(Object.freeze({ spanStart: cursor - spanStart, spanEnd: spanEnd - spanStart }));
  return Object.freeze(pieces);
}

// What a format written where the text stands may write between the ranges
// of its span, and take away from it: the operators that set the text
// rendering mode, the line width, the stroke colour and the fill colour, each
// with its operands, and the rules of pdfNativeTextRuleBlock. Nothing that
// shows, places or selects text, no path but a rule's, no other operator.
// (Mirror of the server's standing_bytes_are_state_or_rules,
// native_paragraph_rewrite_contract.py.)
const STATE_NUMBER = String.raw`[+-]?(?:\d+\.?\d*|\.\d+)`;
const STATE_NAME = String.raw`\/[^\x00\t\n\f\r /<>\[\]()%{}]+`;
const RULE_NUMBER = String.raw`-?\d+(?:\.\d+)?`;
const RULE_BLOCK = new RegExp(String.raw`\nq (?:${RULE_NUMBER} ){1,4}(?:g|rg|k) ${RULE_NUMBER} ${RULE_NUMBER} m ` +
  String.raw`(?:${RULE_NUMBER} ${RULE_NUMBER} l ){3}h f Q\n`, "g");
const STATE_OPERATOR = String.raw`(?:[02] Tr|${STATE_NUMBER} w|${STATE_NUMBER} [Gg]|(?:${STATE_NUMBER} ){3}(?:RG|rg)|` +
  String.raw`(?:${STATE_NUMBER} ){4}[Kk]|${STATE_NAME} (?:CS|cs)|(?:(?:${STATE_NUMBER}|${STATE_NAME}) ){1,33}(?:SCN|SC|scn|sc))`;
const STATE_RUN = new RegExp(String.raw`^[\x00\t\n\f\r ]*(?:${STATE_OPERATOR}(?:[\x00\t\n\f\r ]+|$))*$`);

/** Whether `bytes` (a stream's bytes as text) is nothing but state operators
 * of such a format and rules of this writer's form. */
export function mraiNativeStandingBytesAreStateOrRules(bytes: string): boolean {
  let cursor = 0;
  for (const block of bytes.matchAll(RULE_BLOCK)) {
    if (!STATE_RUN.test(bytes.slice(cursor, block.index))) return false;
    cursor = block.index + block[0].length;
  }
  return STATE_RUN.test(bytes.slice(cursor));
}

/** Where a show is cut in two before the code byte `at` of its codes, and
 * what is written there to end the first show and begin the second. */
export type StandingShowCut = Readonly<{ offset: number; close: string; open: string }>;

/**
 * Where a `Tj` or `TJ` show may be cut before each of the code bytes `cuts`
 * (ascending, inside the show's codes), by writing alone: between two
 * elements of a `TJ` array (`] TJ [`: the numbers after the last string go
 * with the show that follows, which moves the pen by them first), or inside
 * a string (`)] TJ [(`, `>] TJ [<`; `) Tj (`, `> Tj <` for a `Tj`). A show
 * cut so draws every glyph where it drew it: the pen a show ends at is the
 * pen the next begins at. Null where the show is not read that surely (an
 * odd hexadecimal string, a cut inside an escape, another operator).
 */
export function standingShowCuts(
  source: string,
  show: Readonly<{ start: number; end: number }>,
  cuts: readonly number[],
): readonly StandingShowCut[] | null {
  const body = source.slice(show.start, show.end);
  const operator = /(TJ|Tj)$/u.exec(body)?.[1];
  if (!operator) return null;
  const array = operator === "TJ";
  // Each code byte's place in the source, and what kind of string holds it.
  const bytes: { offset: number; kind: "(" | "<"; first: boolean }[] = [];
  let index = 0;
  const end = body.length - 2;
  const skipSpace = () => { while (index < end && /[\x00\t\n\f\r ]/u.test(body[index])) index += 1; };
  skipSpace();
  if (array) {
    if (body[index] !== "[") return null;
    index += 1;
  }
  let closed = !array;
  let strings = 0;
  while (index < end) {
    skipSpace();
    if (index >= end) break;
    const character = body[index];
    if (array && character === "]") {
      closed = true;
      index += 1;
      continue;
    }
    if (closed && (array || strings > 0)) return null;
    if (character === "(") {
      strings += 1;
      let depth = 1;
      let first = true;
      index += 1;
      while (index < end && depth > 0) {
        const at = index;
        const now = body[index];
        if (now === "\\") {
          const next = body[index + 1];
          if (next === undefined) return null;
          if (/[0-7]/u.test(next)) index += 1 + (/^[0-7]{1,3}/u.exec(body.slice(index + 1))![0].length);
          else if (next === "\r") { index += body[index + 2] === "\n" ? 3 : 2; continue; }
          else if (next === "\n") { index += 2; continue; }
          else index += 2;
        } else if (now === "(") { depth += 1; index += 1; }
        else if (now === ")") {
          depth -= 1;
          index += 1;
          if (depth === 0) break;
        } else if (now === "\r" && body[index + 1] === "\n") index += 2;
        else index += 1;
        bytes.push({ offset: show.start + at, kind: "(", first });
        first = false;
      }
      if (depth !== 0) return null;
    } else if (character === "<") {
      strings += 1;
      const close = body.indexOf(">", index);
      const digits = close < 0 ? null : body.slice(index + 1, close);
      if (digits === null || !/^[0-9A-Fa-f]*$/u.test(digits) || digits.length % 2 !== 0) return null;
      for (let at = 0; at < digits.length; at += 2) bytes.push({ offset: show.start + index + 1 + at, kind: "<", first: at === 0 });
      index = close + 1;
    } else if (array && /[-+.0-9]/u.test(character)) {
      while (index < end && /[-+.0-9]/u.test(body[index])) index += 1;
    } else return null;
    if (!array) closed = true;
  }
  if (!closed) return null;
  const out: StandingShowCut[] = [];
  for (const cut of cuts) {
    const at = bytes[cut];
    if (!at || cut <= 0) return null;
    if (at.first && array) {
      // Between two elements: right after the string before it.
      let before = at.offset - show.start - 1;
      while (before > 0 && body[before - 1] !== ")" && body[before - 1] !== ">") before -= 1;
      if (before <= 0) return null;
      out.push({ offset: show.start + before, close: "] TJ", open: "[" });
    } else if (at.first) return null;
    else {
      const [left, right] = at.kind === "(" ? [")", "("] : [">", "<"];
      out.push(array ? { offset: at.offset, close: `${left}] TJ`, open: `[${right}` }
        : { offset: at.offset, close: `${left} Tj`, open: right });
    }
  }
  return Object.freeze(out);
}

const SHOW_CUT = new RegExp(String.raw`^(?:\)\] TJ[\x00\t\n\f\r ]+(?:${STATE_OPERATOR}[\x00\t\n\f\r ]+)*\[\(|` +
  String.raw`>\] TJ[\x00\t\n\f\r ]+(?:${STATE_OPERATOR}[\x00\t\n\f\r ]+)*\[<|` +
  String.raw`\] TJ[\x00\t\n\f\r ]+(?:${STATE_OPERATOR}[\x00\t\n\f\r ]+)*\[|` +
  String.raw`\) Tj[\x00\t\n\f\r ]+(?:${STATE_OPERATOR}[\x00\t\n\f\r ]+)*\(|` +
  String.raw`> Tj[\x00\t\n\f\r ]+(?:${STATE_OPERATOR}[\x00\t\n\f\r ]+)*<)$`);

/** Whether `bytes` cuts a show in two (standingShowCuts), with nothing but
 * state operators between the two shows. (Mirror of the server's
 * standing_bytes_cut_a_show.) */
export function mraiNativeStandingBytesCutAShow(bytes: string): boolean {
  return SHOW_CUT.test(bytes);
}
