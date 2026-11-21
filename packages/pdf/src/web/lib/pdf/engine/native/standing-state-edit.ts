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
        // Where the page itself sets the kind once directly before the show
        // (`0 g` `0 G` `(...) Tj`), the want is written in that operator's
        // place: the state is set where the page set it, in the order it
        // set its kinds. Taken away and written anew against the show, a
        // colour set and then set back left the page's own two operators
        // the other way round (`0 G` `0 g`) -- the same paint, other bytes,
        // and a format taken back gives the page's bytes back (pdf-walk-1000
        
        // own is kept so even where the state before it is the want already
        // (the page's operator said it again; what was set in its place is
        // set back in its place); one set against the show on its line, as
        // this writer sets what it adds, goes as it came where it is no
        // longer wanted.
        const own = region.filter((operator) => operator.kind === kind);
        if (want !== null && own.length === 1 &&
          (want !== written[kind] || /[\n\r]/u.test(source.slice(own[0].end, show.start)))) {
          const set = kind === "tr" ? `${want} Tr` : want;
          if (source.slice(own[0].start, own[0].end) !== set) edits.push({ start: own[0].start, end: own[0].end, text: set });
          written[kind] = want;
          continue;
        }
        for (const operator of own) removeBefore(operator.start, operator.end);
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
// (And, for a lean: a text matrix or a move with its operands, a `T*` -- what
// planStandingLeanEdits sets before and after the box's shows, and takes away
// again. They place text; that no show moves is read back by the writer. No
// part of what may stand between the two pieces of a cut show: SHOW_CUT.)
const PLACING_OPERATOR = String.raw`(?:(?:${STATE_NUMBER} ){6}Tm|${STATE_NUMBER} ${STATE_NUMBER} Td|T\*)`;
const STATE_RUN = new RegExp(String.raw`^[\x00\t\n\f\r ]*(?:(?:${STATE_OPERATOR}|${PLACING_OPERATOR})(?:[\x00\t\n\f\r ]+|$))*$`);

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

// (Between the two pieces: state operators, and -- for a lean of the box's
// stretch of the show -- the text matrix the piece after starts under, set
// where it stands. No move: a `Td` or a `T*` there would carry the piece.)
const CUT_BETWEEN = String.raw`[\x00\t\n\f\r ]+(?:(?:${STATE_OPERATOR}|(?:${STATE_NUMBER} ){6}Tm)[\x00\t\n\f\r ]+)*`;
const SHOW_CUT = new RegExp(String.raw`^(?:\)\] TJ${CUT_BETWEEN}\[\(|>\] TJ${CUT_BETWEEN}\[<|\] TJ${CUT_BETWEEN}\[|` +
  String.raw`\) Tj${CUT_BETWEEN}\(|> Tj${CUT_BETWEEN}<)$`);

/** Whether `bytes` cuts a show in two (standingShowCuts), with nothing but
 * state operators -- and the text matrix of the piece after -- between the
 * two shows. (Mirror of the server's standing_bytes_cut_a_show.) */
export function mraiNativeStandingBytesCutAShow(bytes: string): boolean {
  return SHOW_CUT.test(bytes);
}

// ---------------------------------------------------------------------------
// A lean written where the text stands.
//
// A lean is a shear of text space: x += delta * y. No operator shears the
// text matrix a show is drawn under; it is whatever the last `Tm` set, moved
// on by the `Td`, `TD` and `T*` since (the line matrix) and by the shows
// drawn (the pen). So the lean of a box's shows is written into the
// positioning operators that bring the text to them:
//
//  - a `Tm` that sets the matrix of shows of the box is written leaned (its
//    c and d operands; where it stands is as it was);
//  - a move down a line under a leaned matrix would carry the text along the
//    line too (ty times the lean): a move of its own along the line, written
//    right after it (`x 0 Td`), takes that back. The page's own move -- its
//    operands as the page spelled them, a `T*` as a `T*` -- is not touched, so
//    that the lean taken off leaves the bytes the page had: the move this
//    writer set goes, and nothing has to be spelled again;
//  - where the box's shows are reached from other text by a move alone, the
//    leaned line matrix is set before them by a `Tm` of its own, and where
//    other text is reached from them by a move alone the line matrix is set
//    back before the move;
//  - where a show runs on from one that leans otherwise (a word of the box
//    after a word of another, a formula's letter in its italic face among
//    leaned words), the matrix it starts under is set there, and the next
//    move takes back how far along its line that was (again by a move of
//    its own after it).
//
// Taken off, what a lean set goes the way it came: a matrix that sets what
// the matrix already is once the lean is off is taken away (with the cut of
// the show it stood in, where it stood in one: the show is one show again),
// and a move of this writer's that then moves nowhere with it.
//
// No show's bytes change and no show moves; every show of the stream that is
// not the box's stands under the matrix it had. That is checked by the
// caller from the bytes written, as its paint state is.
// ---------------------------------------------------------------------------

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
const sheared = (matrix: Matrix, delta: number): Matrix =>
  [matrix[0], matrix[1], matrix[2] + delta * matrix[0], matrix[3] + delta * matrix[1], matrix[4], matrix[5]];
const moved = (matrix: Matrix, tx: number, ty: number): Matrix => [matrix[0], matrix[1], matrix[2], matrix[3],
  tx * matrix[0] + ty * matrix[2] + matrix[4], tx * matrix[1] + ty * matrix[3] + matrix[5]];
const sameMatrix = (left: Matrix, right: Matrix) => left.every((value, at) =>
  Math.abs(value - right[at]) <= 2e-6 * Math.max(1, Math.abs(value), Math.abs(right[at])));
const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)$/u;
const linearOf = (matrix: Matrix): Matrix => [matrix[0], matrix[1], matrix[2], matrix[3], 0, 0];

// A show cut in two with the matrix of the piece after set between the two
// (standingShowCuts, and the lean's own matrix): what closes the first piece
// and what opens the second, as this writer writes them.
const CUT_ENDS: readonly (readonly [string, string])[] = [[")] TJ", "[("], [">] TJ", "[<"], ["] TJ", "["], [") Tj", "("], ["> Tj", "<"]];
const CUT_MATRIX_LINE_END = /Tm\n(\[\(|\[<|\[|\(|<)$/u;
/** How far from the pen a matrix this writer set at it may stand when the
 * lean is taken off: the digits it was stated in, and those of each move this
 * writer set on the lines before it (a paragraph of many lines adds them up). */
const PEN_SET_REACH = 1e-4;

/** The bytes of such a cut round the matrix at `from..to` (its operands to
 * its operator), or null where the matrix stands in none: taken away, the
 * two pieces are one show again. The cut and its matrix stand on one line
 * (composeStandingCutEdits); two shows of the page with a matrix this writer
 * set between them do not. (Two strings of an array that stood side by side
 * with no number between them read back as one string: the same glyphs from
 * the same pen.) */
function standingCutAround(source: string, from: number, to: number): Readonly<{ start: number; end: number }> | null {
  for (const [close, open] of CUT_ENDS) {
    if (source.slice(from - close.length - 1, from) === `${close} ` && source.slice(to, to + 1 + open.length) === ` ${open}`) {
      return { start: from - close.length - 1, end: to + 1 + open.length };
    }
  }
  return null;
}

type LeanOperator = Readonly<{
  kind: "Tm" | "Td" | "TD" | "T*" | "next-line-show";
  /** Its operator among the walk's tokens, and its operands. */
  token: number;
  operands: readonly number[];
  /** The line matrix before it, as the stream stands (null: not known). */
  before: Matrix | null;
  /** The leading in force at it (null: not known). */
  leading: number | null;
}>;

/**
 * The edits that lean the shows of `owned` (by the end of each show's span:
 * the shear `delta` of text space it takes) where they stand in the stream
 * `walk` reads. A string instead: why this stream cannot take the lean so (the
 * caller draws the box again, or refuses in its own words).
 */
export function planStandingLeanEdits(input: Readonly<{
  walk: StandingWalk;
  owned: ReadonlyMap<number, number>;
  /** Shows of `owned` that hang off their line -- a sign set above or below
   * the baseline of the text it belongs to (by the end of the show's span):
   * how far along its line each moves with the lean, in the units of the
   * matrix it is drawn under. A show on its line's baseline leans about its
   * own origin and moves nowhere; one that hangs off it is carried along by
   * the lean of the line, as far as it hangs. */
  hangs?: ReadonlyMap<number, number>;
  /** The text matrix each show of the stream starts under, by the end of its span. */
  pens: ReadonlyMap<number, Matrix>;
  /** Whether the stream starts with no leading set (the page's first stream). */
  leadingKnown: boolean;
  format: (value: number) => string;
  /** The text matrix the show after the bytes `from..to` of the stream (a
   * matrix set alone between two shows) starts under without them: where the
   * show before it leaves the pen. Null: not read. Asked of a matrix that may
   * set nothing once the lean is off. */
  penWithout?: (from: number, to: number) => Matrix | null;
}>): readonly StandingEdit[] | string {
  const { walk, owned, format } = input;
  const { source, tokens } = walk;
  const text = (index: number) => source.slice(tokens[index].start, tokens[index].end);
  const numberAt = (index: number): number | null => {
    const spelled = index >= 0 && !tokens[index].operator ? text(index) : "";
    return NUMBER.test(spelled) ? Number(spelled) : null;
  };
  const showAt = new Map(walk.shows.map((show) => [show.first, show] as const));
  const edits: StandingEdit[] = [];
  const matrixText = (matrix: Matrix) => matrix.map(format).join(" ");
  const space = (at: number) => at >= 0 && at < source.length && /[\x00\t\n\f\r ]/u.test(source[at]);

  let leading: number | null = input.leadingKnown ? 0 : null;
  const leadings: (number | null)[] = [];
  let inText = false;
  let line: Matrix | null = null;
  // What was drawn before the positioning operators gathered in `gap`: the
  // text object's start, a show of the box (and the shear it takes), or
  // another show.
  let before: "start" | "other" | Readonly<{ delta: number; hang: number }> = "start";
  let beforeEnd = 0;
  let gap: LeanOperator[] = [];

  // How far along its line the line matrix stands from where the stream as
  // it was has it: a matrix this writer sets where a show starts mid-line
  // (below) is that show's, and the next move from it takes the difference.
  let shift = 0;
  const unread = "The matrix this text is drawn under cannot be read from the page's content.";

  // A move, made to land where it did: under a leaned line matrix it would
  // carry the text along the line too (ty times the lean), and from a line
  // matrix set mid-line it would start that much further on. The move stays
  // as the page wrote it; a move along the line alone is written after it
  // (`correction`: the one already there -- this writer's, set by a lean
  // before -- which is then set anew, or taken away where it would move
  // nowhere).
  // (`hang`: how far further along the line the show the move reaches is to
  // land -- a sign carried along by the lean of the line it hangs on.)
  const landed = (operator: LeanOperator, correction: LeanOperator | null, lean: number | null, hang = 0): string | null => {
    const delta = lean ?? 0;
    const carried = shift - hang;
    shift = 0;
    if (operator.kind === "Tm") return hang === 0 ? null : "A sign that hangs on a line of this box is placed by a matrix after a move.";
    if (operator.kind === "next-line-show") {
      if (carried === 0 && (delta === 0 || operator.leading === 0)) return null;
      return "A show of other text moves to its line as it is drawn, after text of this box.";
    }
    let down = 0;
    if (delta !== 0) {
      if (operator.kind === "T*") {
        if (operator.leading === null) return "The line spacing a `T*` of this text moves by is not stated where this writer can read it.";
        down = -operator.leading;
      } else down = numberAt(operator.operands[1])!;
    }
    const further = -delta * down - carried;
    if (correction) {
      const had = numberAt(correction.operands[0])!;
      const to = had + further;
      if (Math.abs(to - had) <= 1e-12) return null;
      const first = tokens[correction.operands[0]];
      // (Nowhere, to the digits this writer states a move and a matrix in: a
      // move in the line's own units, a matrix on the page.)
      const along = operator.before ? Math.hypot(operator.before[0], operator.before[1]) : 1;
      if (Math.abs(to) <= 1e-6 * (1 + 1 / Math.max(along, 1e-9)) || format(to) === "0") {
        edits.push({ start: space(first.start - 1) ? first.start - 1 : first.start, end: tokens[correction.token].end, text: "" });
      } else edits.push({ start: first.start, end: first.end, text: format(to) });
      return null;
    }
    if (Math.abs(further) <= 1e-12 || format(further) === "0") return null;
    const after = tokens[operator.token].end;
    edits.push({ start: after, end: after, text: ` ${format(further)} 0 Td` });
    return null;
  };
  /** A move along the line alone, right after another move: what this writer sets after a move (above). */
  const alongOnly = (operator: LeanOperator | undefined): operator is LeanOperator =>
    operator !== undefined && operator.kind === "Td" && numberAt(operator.operands[1]) === 0;
  /** The moves of `operators`, each made to land where it did (the last of them `hang` further on). */
  const landedAll = (operators: readonly LeanOperator[], lean: number | null, hang = 0): string | null => {
    for (let at = 0; at < operators.length; at += 1) {
      const operator = operators[at];
      const moves = operator.kind === "Td" || operator.kind === "TD" || operator.kind === "T*";
      const correction = moves && alongOnly(operators[at + 1]) ? operators[at + 1] : null;
      const last = at + (correction ? 1 : 0) === operators.length - 1;
      const why = landed(operator, correction, lean, last ? hang : 0);
      if (why) return why;
      if (correction) at += 1;
    }
    return null;
  };
  const sameLean = (left: number | null, right: number | null) => left === right ||
    (left !== null && right !== null && Math.abs(left - right) <= 1e-12 * Math.max(1, Math.abs(left)));

  /** The operators gathered are done: `next` is drawn after them (null: the
   * text object ends), its span from `nextStart` to `nextEnd`. */
  const close = (next: "other" | Readonly<{ delta: number; hang: number }> | null, nextStart: number, nextEnd: number): string | null => {
    const lean = next !== null && next !== "other" ? next.delta : null;
    const leaned = before !== "start" && before !== "other" ? before.delta : null;
    // A sign that hangs on its line is carried along it by the lean: the
    // move (or the matrix) that reaches it lands that much further on, and
    // what is drawn after it is reached from there as it was (`shift`). It is
    // reached by a move or a matrix of its own -- one that runs on from the
    // show before it, or that a show runs on from, is not carried so.
    const hang = next !== null && next !== "other" ? next.hang : 0;
    const hungBefore = before !== "start" && before !== "other" ? before.hang : 0;
    if (hungBefore !== 0 && gap.length === 0 && next !== null) return "A show runs on from a sign that hangs on a line of this box.";
    if (hang !== 0 && gap.length === 0) return "A sign that hangs on a line of this box runs on from the show before it.";
    if (hang !== 0) {
      const lastSetOfSign = gap.map((operator) => operator.kind).lastIndexOf("Tm");
      if (lastSetOfSign >= 0) {
        // Placed by a matrix: its origin set that much further along the line it is drawn on.
        const set = gap[lastSetOfSign];
        const matrix = set.operands.map((operand) => numberAt(operand)!) as unknown as Matrix;
        const leanedMatrix = sheared(matrix, lean ?? 0);
        const carriedTo = moved(leanedMatrix, hang, 0);
        shift = 0;
        for (const at of [2, 3, 4, 5]) {
          if (Math.abs(carriedTo[at] - matrix[at]) <= 1e-12) continue;
          edits.push({ start: tokens[set.operands[at]].start, end: tokens[set.operands[at]].end, text: format(carriedTo[at]) });
        }
        const why = landedAll(gap.slice(lastSetOfSign + 1), lean);
        if (why) return why;
        shift = hang;
        return null;
      }
      if (!sameLean(lean, leaned)) return "A sign that hangs on a line of this box is reached from text leaned otherwise.";
      const why = landedAll(gap, lean, hang);
      if (why) return why;
      shift = hang;
      return null;
    }
    const lastSet = gap.map((operator) => operator.kind).lastIndexOf("Tm");
    if (lastSet >= 0) {
      const set = gap[lastSet];
      const matrix = set.operands.map((operand) => numberAt(operand)!) as unknown as Matrix;
      const carried = shift;
      shift = 0;
      const leanedMatrix = sheared(matrix, lean ?? 0);
      // The frame the line matrix stands in before this matrix, once the
      // lean is written: the frame of the show before it.
      const frame = set.before ? sheared(set.before, leaned ?? 0) : null;
      const only = gap.filter((operator) => operator.kind === "Tm").length === 1;
      const from = tokens[set.operands[0]].start, to = tokens[set.token].end;
      // Alone between two shows, where it stands now it changes the frame the
      // text is drawn in at the pen (a show that runs on from one leaned
      // otherwise: this writer set it, and cut the show where it stands in
      // one). With the lean written it sets what the matrix already is: it
      // goes, the cut with it, and the next move starts from the line matrix
      // that was there before it.
      // (This writer's, by how it is written: the matrix in the digits this
      // writer states one in, at the pen to those digits, ending its line or
      // standing in a cut. A matrix of the page's own that this writer leaned
      // is spelled otherwise -- below -- and stays.)
      const between = only && gap.length === 1 && before !== "start" && next !== null && set.before !== null && frame !== null;
      const cut = between ? standingCutAround(source, from, to) : null;
      if (between && !sameMatrix(linearOf(matrix), linearOf(set.before!)) && sameMatrix(linearOf(leanedMatrix), linearOf(frame!)) &&
        source.slice(from, to) === `${matrixText(matrix)} Tm` && (cut !== null || source[to] === "\n")) {
        const pen = input.penWithout?.(from, to + 1) ?? null;
        const [a, b] = [set.before![0], set.before![1]];
        const along = a * a + b * b;
        const [dx, dy] = [matrix[4] - set.before![4], matrix[5] - set.before![5]];
        const reach = along > 0 ? (dx * a + dy * b) / along : 0;
        if (pen && along > 0 && sameMatrix(linearOf(pen), linearOf(set.before!)) &&
          Math.abs(pen[4] - matrix[4]) <= PEN_SET_REACH && Math.abs(pen[5] - matrix[5]) <= PEN_SET_REACH &&
          Math.hypot(dx - reach * a, dy - reach * b) <= 1e-6 * Math.max(1, Math.hypot(dx, dy))) {
          edits.push({ ...(cut ?? { start: from, end: to + 1 }), text: "" });
          shift = carried - reach;
          return null;
        }
      }
      // It sets what the line matrix already is once the lean is written
      // (as this writer set it: back before the moves other text is reached
      // from by, or leaned before the box's shows where they are reached by
      // moves), the pen at the line's start or the text moved on from it:
      // it goes.
      if (only && frame && (lastSet > 0 || before === "start" || lastSet < gap.length - 1) &&
        !sameMatrix(matrix, set.before!) && sameMatrix(leanedMatrix, moved(frame, carried, 0))) {
        edits.push(lastSet === 0 && gap.length > 1 ? { start: space(from - 1) ? from - 1 : from, end: to, text: "" }
          : { start: from, end: space(to) ? to + 1 : to, text: "" });
        return landedAll(gap.slice(lastSet + 1), lean);
      }
      if (lean !== null) {
        // A matrix of the page's own between two shows of one frame that the
        // lean parts is, leaned, what a matrix this writer sets at a pen
        // looks like (above). Its leaned operand is spelled as this writer
        // spells no number -- a nought after its last digit -- so that with
        // the lean off it is the page's still: set upright where it stands.
        const ownsPlace = between && sameMatrix(linearOf(matrix), linearOf(set.before!)) && !sameLean(lean, leaned);
        for (const at of [2, 3]) {
          if (Math.abs(leanedMatrix[at] - matrix[at]) <= 1e-12) continue;
          const spelled = format(leanedMatrix[at]);
          edits.push({ start: tokens[set.operands[at]].start, end: tokens[set.operands[at]].end,
            text: ownsPlace && spelled !== "0" ? `${spelled}${spelled.includes(".") ? "" : "."}0` : spelled });
        }
      }
      return landedAll(gap.slice(lastSet + 1), lean);
    }
    // (Moves the text object ends on bring text to nothing.)
    if (next === null) return null;
    // The same lean before and after (or none), by moves alone.
    if (sameLean(lean, leaned)) return landedAll(gap, lean);
    if (gap.length === 0) {
      // The lean changes where a show runs on from the show before it: the
      // matrix it starts under is set there, leaned as it takes it. (The first
      // show of a text object that places nothing starts under the line matrix.)
      const pen = input.pens.get(nextEnd) ?? (before === "start" ? line : null);
      if (!pen || !line) return unread;
      const [a, b] = [line[0], line[1]];
      const along = a * a + b * b;
      if (!(along > 0) || !sameMatrix([pen[0], pen[1], pen[2], pen[3], 0, 0], [line[0], line[1], line[2], line[3], 0, 0])) return unread;
      const [dx, dy] = [pen[4] - line[4], pen[5] - line[5]];
      const from = (dx * a + dy * b) / along;
      if (Math.hypot(dx - from * a, dy - from * b) > 1e-6 * Math.max(1, Math.hypot(dx, dy))) return unread;
      edits.push({ start: nextStart, end: nextStart, text: `${matrixText(lean === null ? pen : sheared(pen, lean))} Tm\n` });
      shift = from;
      return null;
    }
    if (lean !== null) {
      // Reached by moves from text leaned otherwise (or not): the line matrix
      // is set, leaned, before the show -- where the stream as it was has it,
      // whatever the moves before it now start from. (They are left as the
      // page wrote them: with the lean off and this matrix gone they land
      // where they did.)
      if (!line) return unread;
      shift = 0;
      edits.push({ start: nextStart, end: nextStart, text: `${matrixText(sheared(line, lean))} Tm\n` });
      return null;
    }
    // Other text reached from the box's by moves: the line matrix set back before them.
    const set = gap[0].before;
    if (!set) return unread;
    edits.push({ start: beforeEnd, end: beforeEnd, text: `\n${matrixText(set)} Tm` });
    shift = 0;
    return null;
  };

  for (let index = 0; index < tokens.length; index += 1) {
    const show = showAt.get(index);
    if (show) {
      if (!inText) return "A show of this page is drawn outside a text object.";
      const operator = text(show.operator);
      const delta = owned.get(show.end);
      if (operator === "'" || operator === '"') {
        if (delta !== undefined) return "A show of this box moves to its line as it is drawn.";
        gap.push({ kind: "next-line-show", token: show.operator, operands: [], before: line, leading });
        line = line && leading !== null ? moved(line, 0, -leading) : null;
      }
      const next = delta !== undefined ? Object.freeze({ delta, hang: input.hangs?.get(show.end) ?? 0 }) : "other" as const;
      const why = close(next, tokens[show.first].start, show.end);
      if (why) return why;
      before = next;
      beforeEnd = tokens[show.operator].end;
      gap = [];
      index = show.operator;
      continue;
    }
    if (!tokens[index].operator) continue;
    const operator = text(index);
    if (operator === "q") leadings.push(leading);
    else if (operator === "Q") leading = leadings.length ? leadings.pop()! : null;
    else if (operator === "TL") leading = numberAt(index - 1);
    else if (operator === "BT") {
      inText = true;
      line = IDENTITY;
      shift = 0;
      before = "start";
      gap = [];
    } else if (operator === "ET") {
      const why = close(null, tokens[index].start, tokens[index].end);
      if (why) return why;
      inText = false;
      before = "start";
      gap = [];
    } else if (operator === "Tm" && inText) {
      const operands = [6, 5, 4, 3, 2, 1].map((back) => index - back);
      if (operands.some((operand) => numberAt(operand) === null)) return "A text matrix of this page is not six plain numbers.";
      gap.push({ kind: "Tm", token: index, operands, before: line, leading });
      line = operands.map((operand) => numberAt(operand)!) as unknown as Matrix;
    } else if ((operator === "Td" || operator === "TD") && inText) {
      const operands = [index - 2, index - 1];
      const [tx, ty] = operands.map(numberAt);
      if (tx === null || ty === null) return "A text move of this page is not two plain numbers.";
      gap.push({ kind: operator, token: index, operands, before: line, leading });
      line = line ? moved(line, tx, ty) : null;
      if (operator === "TD") leading = -ty;
    } else if (operator === "T*" && inText) {
      gap.push({ kind: "T*", token: index, operands: [], before: line, leading });
      line = line && leading !== null ? moved(line, 0, -leading) : null;
    }
  }
  if (inText) return "A text object of this page is not closed.";
  return normalizeStandingEdits(edits);
}

/**
 * Edits of a stream whose shows were cut first (`cuts`: what each cut wrote,
 * by where in the stream as it was), as edits of the stream as it was: an
 * edit that stands in what a cut wrote is written with the cut, the others
 * where they stood before the cuts. Null where an edit takes away part of
 * what a cut wrote, or runs across one.
 */
export function composeStandingCutEdits(
  cuts: readonly StandingEdit[],
  edits: readonly StandingEdit[],
): readonly StandingEdit[] | null {
  let grown = 0;
  const written = cuts.map((cut) => {
    if (cut.start !== cut.end) return null;
    const at = cut.start + grown;
    grown += cut.text.length;
    return { at, end: at + cut.text.length, source: cut.start, text: cut.text, splices: [] as { at: number; text: string }[] };
  });
  if (written.some((entry) => entry === null)) return null;
  const out: StandingEdit[] = [];
  for (const edit of edits) {
    const inside = written.find((entry) => entry!.at <= edit.start && edit.end <= entry!.end);
    if (inside) {
      if (edit.start !== edit.end) return null;
      inside.splices.push({ at: edit.start - inside.at, text: edit.text });
      continue;
    }
    if (written.some((entry) => edit.start < entry!.end && entry!.at < edit.end)) return null;
    const before = written.reduce((sum, entry) => (entry!.end <= edit.start ? sum + entry!.text.length : sum), 0);
    out.push({ start: edit.start - before, end: edit.end - before, text: edit.text });
  }
  for (const entry of written) {
    let text = entry!.text;
    // (From the last place to the first, so that each place is still where it was; two at one place in their order.)
    for (const splice of entry!.splices.map((value, order) => ({ ...value, order }))
      .sort((left, right) => right.at - left.at || right.order - left.order)) {
      text = text.slice(0, splice.at) + splice.text + text.slice(splice.at);
    }
    // The matrix of the piece after, set in the cut: on one line with it. A
    // matrix set before a show the page drew apart ends its line (the lean's
    // own form), and the two are told by that when the lean is taken off --
    // a cut is one show again, two shows of the page stay two.
    text = text.replace(CUT_MATRIX_LINE_END, "Tm $1");
    out.push({ start: entry!.source, end: entry!.source, text });
  }
  return normalizeStandingEdits(out);
}

/** Whether `text` is one plain number: an operand of a positioning operator
 * a lean rewrites (a `Tm`'s c or d, a move's tx), written for the number it
 * was. Mirror of the server's standing_bytes_are_a_number. */
export function mraiNativeStandingBytesAreANumber(text: string): boolean {
  return NUMBER.test(text);
}
