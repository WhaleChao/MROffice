/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * UnicodeMath → OMML model, following Word's Build Up. The rules come from
 * Unicode Technical Note #28 v3.2 and were checked case by case against
 * Word 16.86's own build-up of the same linear text (operand extent, space
 * terminators, bracket matching, function recognition, n-ary operands).
 */

import type { MathList, MathNode, MathRunStyle } from "./ast";
import {
  ABOVE,
  BELOW,
  CASES,
  CLOSE_DELIMITER,
  CLOSE_GROUP,
  CLOSERS,
  DELIMITER_CHARACTERS,
  EQUATION_ARRAY,
  FRACTION_OPERATORS,
  FUNCTION_APPLY,
  GROUP_CHARACTERS,
  LIMIT_FUNCTIONS,
  MATRIX_FENCES,
  NARY_AND,
  NARY_OPERATORS,
  NBSP,
  OPEN_DELIMITER,
  OPEN_GROUP,
  OPENERS,
  PHANTOMS,
  RECOGNIZED_FUNCTIONS,
  SELF_DELIMITERS,
  STRETCH_ARROWS,
  VERTICAL_SEPARATOR,
  decodeMathCharacter,
  isCombiningMark,
  isItalicByDefault,
  isOperandCharacter,
  unicodeSubscript,
  unicodeSuperscript,
} from "./characters";
import { UNICODE_MATH_KEYWORDS } from "./keywords";

interface Token {
  readonly ch: string;
  readonly style: MathRunStyle;
  /** Escaped with a backslash: always ordinary text. */
  readonly lit: boolean;
  /** Ordinary text written in ASCII double quotes. */
  readonly nor?: string;
}

const PRIMES: Readonly<Record<string, string>> = {
  "'": "′",
  "′": "′",
  "″": "″",
  "‴": "‴",
  "⁗": "⁗",
};
const SCRIPT_SIGNS = new Set(["-", "−", "+", "±", "∓"]);
const MIDDLE = "ⓜ";
const DIVIDES = "∣";

function tokenize(text: string): Token[] {
  const chars = Array.from(text);
  const tokens: Token[] = [];
  for (let index = 0; index < chars.length; ) {
    const character = chars[index];
    if (character === '"') {
      let cursor = index + 1;
      let body = "";
      let closed = false;
      while (cursor < chars.length) {
        if (
          chars[cursor] === "\\" &&
          (chars[cursor + 1] === '"' || chars[cursor + 1] === "\\")
        ) {
          body += chars[cursor + 1];
          cursor += 2;
          continue;
        }
        if (chars[cursor] === '"') {
          closed = true;
          break;
        }
        body += chars[cursor];
        cursor += 1;
      }
      if (closed && body) {
        tokens.push({
          ch: "",
          style: { nor: true },
          lit: false,
          nor: Array.from(body, (item) => decodeMathCharacter(item).character).join(""),
        });
        index = cursor + 1;
        continue;
      }
      if (closed) {
        tokens.push({ ch: '"', style: {}, lit: true }, { ch: '"', style: {}, lit: true });
        index = cursor + 1;
        continue;
      }
      tokens.push({ ch: '"', style: {}, lit: true });
      index += 1;
      continue;
    }
    if (character === "\\" && index + 1 < chars.length) {
      const next = chars[index + 1];
      if (/^[A-Za-z]$/.test(next)) {
        let cursor = index + 1;
        let word = "";
        while (cursor < chars.length && /^[A-Za-z]$/.test(chars[cursor])) {
          word += chars[cursor];
          cursor += 1;
        }
        if (Object.prototype.hasOwnProperty.call(UNICODE_MATH_KEYWORDS, word)) {
          const glyph = UNICODE_MATH_KEYWORDS[word];
          if (isItalicByDefault(glyph)) {
            tokens.push({ ch: glyph, style: {}, lit: false });
          } else {
            const decoded = decodeMathCharacter(glyph);
            tokens.push({ ch: decoded.character, style: styleOf(decoded), lit: false });
          }
          index = cursor + (chars[cursor] === " " ? 1 : 0);
          continue;
        }
        tokens.push({ ch: "\\", style: {}, lit: true });
        for (const letter of word) tokens.push({ ch: letter, style: {}, lit: true });
        index = cursor + (chars[cursor] === " " ? 1 : 0);
        continue;
      }
      tokens.push({ ch: next, style: {}, lit: true });
      index += 2;
      continue;
    }
    const decoded = decodeMathCharacter(character);
    // Word stores the ASCII hyphen-minus for a typed or pasted U+2212.
    const value = decoded.character === "\u2212" ? "-" : decoded.character;
    tokens.push({ ch: value, style: styleOf(decoded), lit: false });
    index += 1;
  }
  return tokens;
}

function styleOf(decoded: ReturnType<typeof decodeMathCharacter>): MathRunStyle {
  const style: { sty?: MathRunStyle["sty"]; scr?: MathRunStyle["scr"] } = {};
  if (decoded.sty) style.sty = decoded.sty;
  if (decoded.scr) style.scr = decoded.scr;
  return style;
}

function sameStyle(a: MathRunStyle, b: MathRunStyle): boolean {
  return (
    a.sty === b.sty &&
    a.scr === b.scr &&
    Boolean(a.nor) === Boolean(b.nor) &&
    Boolean(a.lit) === Boolean(b.lit)
  );
}

/** Append text, merging with a preceding run of the same style. */
function pushText(list: MathList, text: string, style: MathRunStyle): void {
  const last = list.at(-1);
  if (last && last.kind === "text" && sameStyle(last.style, style)) {
    list[list.length - 1] = { ...last, text: last.text + text };
    return;
  }
  list.push({ kind: "text", text, style });
}

function appendList(target: MathList, nodes: MathList): void {
  for (const node of nodes) {
    if (node.kind === "text") pushText(target, node.text, node.style);
    else target.push(node);
  }
}

const isObject = (node: MathNode | undefined) =>
  Boolean(node && (node.kind !== "text" || node.style.nor));

type FactorKind = "paren" | "group" | "other";
interface Factor {
  nodes: MathList;
  kind: FactorKind;
}
interface Operand {
  nodes: MathList;
  factors: Factor[];
}

type Stop = (token: Token) => boolean;

const plain = (token: Token | undefined, character: string) =>
  Boolean(token && !token.lit && token.nor === undefined && token.ch === character);

class Parser {
  private pos = 0;

  constructor(private readonly tokens: readonly Token[]) {}

  parse(): MathList {
    const list = this.parseExp(() => false);
    // `#` numbers the equation; Word builds the zone into an equation array.
    if (this.tokens.some((token) => !token.lit && token.nor === undefined && token.ch === "#")) {
      return [{ kind: "eqArr", rows: [list] }];
    }
    // Anything a stop left behind (an unmatched closer) is ordinary text.
    while (this.pos < this.tokens.length) {
      const token = this.tokens[this.pos++];
      if (token.nor !== undefined) pushText(list, token.nor, { nor: true });
      else pushText(list, token.ch, token.lit ? { ...token.style, lit: true } : token.style);
      appendList(list, this.parseExp(() => false));
    }
    return list;
  }

  private peek(offset = 0): Token | undefined {
    return this.tokens[this.pos + offset];
  }

  private isSpaceAt(offset = 0): boolean {
    return plain(this.peek(offset), " ");
  }

  private startsOperand(offset: number, stops: Stop): boolean {
    const token = this.peek(offset);
    if (!token || stops(token)) return false;
    if (token.nor !== undefined) return true;
    if (token.lit) return isOperandCharacter(token.ch);
    const character = token.ch;
    return (
      this.isOperandToken(token) ||
      OPENERS.has(character) ||
      SELF_DELIMITERS.has(character) ||
      character === OPEN_DELIMITER ||
      character === OPEN_GROUP ||
      isCombiningMark(character) ||
      PREFIX.has(character) ||
      MATRIX_FENCES.has(character) ||
      character === EQUATION_ARRAY ||
      character === CASES ||
      NARY_OPERATORS.has(character) ||
      ((character === "_" || character === "^" || character === BELOW || character === ABOVE) &&
        offset === 0 &&
        this.atContainerStart()) ||
      (STRETCH_ARROWS.has(character) &&
        (plain(this.peek(offset + 1), ABOVE) || plain(this.peek(offset + 1), BELOW)))
    );
  }

  private isOperandToken(token: Token): boolean {
    if (token.nor !== undefined) return false;
    // An escaped operator (`\/`, `\(`) is literal text that still ends an operand.
    if (token.lit) return isOperandCharacter(token.ch);
    const character = token.ch;
    if (isCombiningMark(character) || PRIMES[character]) return false;
    if (unicodeSuperscript(character) || unicodeSubscript(character)) return false;
    return isOperandCharacter(character);
  }

  /* -------------------------------------------------------------------- */

  private parseExp(stops: Stop): MathList {
    const out: MathList = [];
    let spaceUsed = false;
    while (this.pos < this.tokens.length) {
      const token = this.tokens[this.pos];
      if (stops(token)) break;
      if (plain(token, " ")) {
        this.pos += 1;
        const following = this.peek();
        if (following && !following.lit && isCombiningMark(following.ch)) continue;
        if (!spaceUsed && isObject(out.at(-1))) {
          spaceUsed = true;
          continue;
        }
        pushText(out, " ", {});
        continue;
      }
      spaceUsed = false;
      const before = this.pos;
      if (this.startsOperand(0, stops) || this.isFractionOperator(token)) {
        const nodes = this.parseFractionChain(stops);
        if (this.pos > before) {
          appendList(out, nodes);
          continue;
        }
      }
      this.pos += 1;
      if (token.nor !== undefined) pushText(out, token.nor, { nor: true });
      else pushText(out, token.ch, token.lit ? { ...token.style, lit: true } : token.style);
      // Any character can carry a script: `→^+`.
      if (
        token.nor === undefined &&
        token.ch !== " " &&
        (plain(this.peek(), "_") || plain(this.peek(), "^"))
      ) {
        const { rest, unit } = splitUnit(out, "other");
        out.length = 0;
        appendList(out, rest);
        out.push(this.parseScripts(unit, stops));
      }
    }
    return out;
  }

  private isFractionOperator(token: Token | undefined): boolean {
    return Boolean(token && !token.lit && token.nor === undefined && FRACTION_OPERATORS.has(token.ch));
  }

  private parseFractionChain(stops: Stop): MathList {
    let left: Operand = { nodes: [], factors: [] };
    if (!this.isFractionOperator(this.peek())) {
      left = this.parseOperand(stops, null, true);
      if (left.nodes.length === 0) return [];
    }
    while (this.pos < this.tokens.length && this.isFractionOperator(this.peek()) && !stops(this.peek()!)) {
      const operator = this.tokens[this.pos++].ch;
      const right = this.startsOperand(0, stops)
        ? this.parseOperand(stops, null)
        : { nodes: [], factors: [] };
      let node: MathNode = {
        kind: "f",
        type: FRACTION_OPERATORS.get(operator) ?? "bar",
        num: stripOperand(left),
        den: stripOperand(right),
      };
      if (operator === "⒞") {
        node = { kind: "d", beg: "(", end: ")", sep: "|", args: [[node]] };
      }
      left = { nodes: [node], factors: [{ nodes: [node], kind: "other" }] };
    }
    return left.nodes;
  }

  private parseOperand(stops: Stop, forbidden: string | null, continues = false): Operand {
    const nodes: MathList = [];
    const factors: Factor[] = [];
    while (this.pos < this.tokens.length) {
      const token = this.tokens[this.pos];
      if (stops(token)) break;
      if (plain(token, " ")) {
        // A space after a built-up object terminates that object; the
        // operand continues when another operand follows.
        if (
          continues &&
          isObject(nodes.at(-1)) &&
          this.startsOperand(1, stops) &&
          !this.isFractionOperator(this.peek(1))
        ) {
          this.pos += 1;
          continue;
        }
        break;
      }
      const before = this.pos;
      const factor = this.parseFactor(stops, forbidden);
      if (!factor || this.pos === before) {
        this.pos = before;
        break;
      }
      appendList(nodes, factor.nodes);
      factors.push(factor);
    }
    return { nodes, factors };
  }

  /* -------------------------------------------------------------------- */

  private parseFactor(stops: Stop, forbidden: string | null): Factor | null {
    const token = this.peek();
    if (!token || stops(token)) return null;
    let base: Factor | null = null;
    if (token.nor !== undefined) {
      this.pos += 1;
      base = { nodes: [{ kind: "text", text: token.nor, style: { nor: true } }], kind: "other" };
    } else if (this.isOperandToken(token)) {
      base = this.parseAtoms(stops);
    } else {
      const character = token.ch;
      if (character === "(" && (plain(this.peek(1), "_") || plain(this.peek(1), "^"))) {
        base = this.parsePrescript(stops) ?? this.parseBracket();
      } else if (OPENERS.has(character) || character === OPEN_DELIMITER) base = this.parseBracket();
      else if (isCombiningMark(character)) base = { nodes: [], kind: "group" };
      else if (
        (character === "_" || character === "^" || character === BELOW || character === ABOVE) &&
        character !== forbidden &&
        this.atContainerStart()
      ) {
        base = { nodes: [], kind: "group" };
      } else if (SELF_DELIMITERS.has(character)) {
        base = this.hasPartnerAhead(character) || this.closerAhead() ? this.parseBracket() : null;
      } else if (character === OPEN_GROUP) base = this.parseGroup();
      else if (PREFIX.has(character)) base = this.parsePrefix(stops);
      else if (MATRIX_FENCES.has(character) || character === EQUATION_ARRAY || character === CASES) {
        base = this.parseArray();
      } else if (NARY_OPERATORS.has(character)) base = this.parseNary(stops);
      else if (
        STRETCH_ARROWS.has(character) &&
        (plain(this.peek(1), ABOVE) || plain(this.peek(1), BELOW))
      ) {
        this.pos += 1;
        base = { nodes: [{ kind: "text", text: character, style: token.style }], kind: "other" };
      }
    }
    if (!base) return null;
    return this.parsePostfix(base, stops, forbidden);
  }

  private parseAtoms(stops: Stop): Factor {
    const run: Token[] = [];
    while (this.pos < this.tokens.length) {
      const token = this.tokens[this.pos];
      if (stops(token)) break;
      if (this.isOperandToken(token)) {
        run.push(token);
        this.pos += 1;
        continue;
      }
      // A decimal point or comma between ASCII digits stays in the number.
      const next = this.peek(1);
      if (
        (plain(token, ".") || plain(token, ",")) &&
        run.length > 0 &&
        /^[0-9]$/.test(run[run.length - 1].ch) &&
        next &&
        !next.lit &&
        /^[0-9]$/.test(next.ch)
      ) {
        run.push(token);
        this.pos += 1;
        continue;
      }
      break;
    }
    // A recognized function name: the run's trailing upright ASCII word.
    let start = run.length;
    while (
      start > 0 &&
      !run[start - 1].lit &&
      run[start - 1].style.sty === "p" &&
      !run[start - 1].style.scr &&
      /^[A-Za-z]$/.test(run[start - 1].ch)
    ) {
      start -= 1;
    }
    const word = run.slice(start).map((token) => token.ch).join("");
    if (word && RECOGNIZED_FUNCTIONS.has(word) && this.functionFollows(stops)) {
      const nodes = tokensToText(run.slice(0, start));
      const name: MathList = [{ kind: "text", text: word, style: { sty: "p" } }];
      nodes.push(this.parseFunction(name, word, stops));
      return { nodes, kind: "other" };
    }
    return { nodes: tokensToText(run), kind: "other" };
  }

  private postfixFollows(): boolean {
    const next = this.peek();
    if (!next || next.lit || next.nor !== undefined) return false;
    return (
      ["_", "^", BELOW, ABOVE, NBSP].includes(next.ch) ||
      isCombiningMark(next.ch) ||
      Boolean(PRIMES[next.ch]) ||
      unicodeSuperscript(next.ch) !== null ||
      unicodeSubscript(next.ch) !== null
    );
  }

  private functionFollows(stops: Stop): boolean {
    const next = this.peek();
    if (!next || next.lit || next.nor !== undefined || stops(next)) return false;
    if ([FUNCTION_APPLY, "_", "^", BELOW, ABOVE, OPEN_GROUP].includes(next.ch)) return true;
    if (OPENERS.has(next.ch)) return true;
    if (next.ch === " ") return this.startsOperand(1, stops);
    return false;
  }

  private parseFunction(name: MathList, word: string | null, stops: Stop): MathNode {
    let current: MathList = name;
    for (;;) {
      const token = this.peek();
      if (plain(token, "_") || plain(token, "^")) {
        if (word && LIMIT_FUNCTIONS.has(word) && plain(token, "_")) {
          this.pos += 1;
          current = [{ kind: "limLow", e: current, lim: this.parseScriptArgument("_", stops) }];
          continue;
        }
        current = [this.parseScripts(current, stops)];
        continue;
      }
      if (plain(token, BELOW) || plain(token, ABOVE)) {
        this.pos += 1;
        current = [this.parseLimit(current, token!.ch, stops)];
        continue;
      }
      break;
    }
    if (plain(this.peek(), FUNCTION_APPLY)) this.pos += 1;
    else if (this.isSpaceAt() && this.startsOperand(1, stops)) this.pos += 1;
    return { kind: "func", name: current, e: this.parseFunctionArgument(stops) };
  }

  private parseFunctionArgument(stops: Stop): MathList {
    const token = this.peek();
    if (!token || stops(token)) return [];
    if (plain(token, OPEN_GROUP)) {
      const start = this.pos;
      const group = this.parseGroup();
      if (group && !this.postfixFollows()) return group.nodes;
      this.pos = start;
    }
    if (!token.lit && token.nor === undefined && OPENERS.has(token.ch)) {
      const bracket = this.parseBracket();
      if (bracket) return this.parsePostfix(bracket, stops, null).nodes;
    }
    if (!this.startsOperand(0, stops)) return [];
    return this.parseOperand(stops, null).nodes;
  }

  /* -------------------------------------------------------------------- */

  private parsePostfix(base: Factor, stops: Stop, forbidden: string | null): Factor {
    let nodes = base.nodes;
    let kind = base.kind;
    for (;;) {
      const token = this.peek();
      if (!token || token.lit || token.nor !== undefined || stops(token)) break;
      const character = token.ch;
      if (character === "_" || character === "^") {
        if (forbidden === character) break;
        const { rest, unit } = splitUnit(nodes, kind);
        nodes = [...rest, this.parseScripts(unit, stops, kind !== "group")];
        kind = "other";
        continue;
      }
      if (character === BELOW || character === ABOVE) {
        this.pos += 1;
        const { rest, unit } = splitUnit(nodes, kind);
        nodes = [...rest, this.parseLimit(unit, character, stops)];
        kind = "other";
        continue;
      }
      const spacedMark = this.peek(1);
      if (
        character === " " &&
        spacedMark &&
        !spacedMark.lit &&
        isCombiningMark(spacedMark.ch) &&
        nodes.length > 0 &&
        (isObject(nodes.at(-1)) || kind === "paren")
      ) {
        this.pos += 2;
        const { rest, unit } = splitUnit(nodes, kind);
        nodes = [...rest, { kind: "acc", chr: spacedMark.ch, e: unit }];
        kind = "other";
        continue;
      }
      const nestedAccent =
        character === " " &&
        plain(this.peek(1), NBSP) &&
        nodes.at(-1)?.kind === "acc" &&
        this.peek(2) !== undefined &&
        isCombiningMark(this.peek(2)!.ch);
      if (nestedAccent) {
        this.pos += 1;
        continue;
      }
      const markAt = character === NBSP ? 1 : 0;
      const mark = this.peek(markAt);
      if (
        mark &&
        !mark.lit &&
        mark.nor === undefined &&
        isCombiningMark(mark.ch) &&
        (markAt === 0 || nodes.length > 0)
      ) {
        this.pos += markAt + 1;
        const { rest, unit } = splitUnit(nodes, kind);
        const accentBase =
          markAt === 1 && kind === "paren" && unit.length === 1 && unit[0].kind === "d"
            ? unit[0].args[0]
            : unit;
        nodes = [...rest, { kind: "acc", chr: mark.ch, e: accentBase }];
        kind = "other";
        continue;
      }
      if (unicodeSuperscript(character) || unicodeSubscript(character)) {
        const superscript = Boolean(unicodeSuperscript(character));
        let text = "";
        while (this.peek()) {
          const next = this.peek()!;
          const mapped = next.lit
            ? null
            : superscript
              ? unicodeSuperscript(next.ch)
              : unicodeSubscript(next.ch);
          if (!mapped) break;
          text += mapped;
          this.pos += 1;
        }
        const { rest, unit } = splitUnit(nodes, kind);
        const script: MathList = [{ kind: "text", text, style: {} }];
        nodes = [
          ...rest,
          superscript ? { kind: "sSup", e: unit, sup: script } : { kind: "sSub", e: unit, sub: script },
        ];
        kind = "other";
        continue;
      }
      if (PRIMES[character]) {
        let text = "";
        while (this.peek() && !this.peek()!.lit && PRIMES[this.peek()!.ch]) {
          text += PRIMES[this.peek()!.ch];
          this.pos += 1;
        }
        const { rest, unit } = splitUnit(nodes, kind);
        const primed: MathNode = { kind: "sSup", e: unit, sup: [{ kind: "text", text, style: {} }] };
        PRIMED.add(primed);
        nodes = [...rest, primed];
        kind = "other";
        continue;
      }
      if (character === FUNCTION_APPLY) {
        this.pos += 1;
        const { rest, unit } = splitName(nodes, kind);
        const name = unit.map((node) =>
          node.kind === "text" && !node.style.nor && !node.style.sty && !node.style.scr
            ? { ...node, style: { ...node.style, sty: "p" as const } }
            : node,
        );
        nodes = [...rest, { kind: "func", name, e: this.parseFunctionArgument(stops) }];
        kind = "other";
        continue;
      }
      break;
    }
    return { nodes, kind };
  }

  private parseScripts(unit: MathList, stops: Stop, joinPrime = false): MathNode {
    const operator = this.tokens[this.pos++].ch;
    const first = this.parseScriptArgument(operator, stops);
    const primed =
      joinPrime && unit.length === 1 && unit[0].kind === "sSup" && PRIMED.has(unit[0]) ? unit[0] : null;
    if (
      operator === "^" &&
      primed &&
      primed.sup.length === 1 &&
      primed.sup[0].kind === "text" &&
      /^[′″‴⁗]+$/u.test(primed.sup[0].text) &&
      !plain(this.peek(), "_")
    ) {
      const sup: MathList = [primed.sup[0]];
      appendList(sup, first);
      return { kind: "sSup", e: primed.e, sup };
    }
    const other = operator === "_" ? "^" : "_";
    const groupCharacter = unit.length === 1 && unit[0].kind === "groupChr";
    if (groupCharacter && !plain(this.peek(), other)) {
      return operator === "^"
        ? { kind: "limUpp", e: unit, lim: first }
        : { kind: "limLow", e: unit, lim: first };
    }
    if (plain(this.peek(), other)) {
      this.pos += 1;
      const second = this.parseScriptArgument(other, stops);
      return operator === "_"
        ? { kind: "sSubSup", e: unit, sub: first, sup: second }
        : { kind: "sSubSup", e: unit, sub: second, sup: first };
    }
    return operator === "_"
      ? { kind: "sSub", e: unit, sub: first }
      : { kind: "sSup", e: unit, sup: first };
  }

  private parseScriptArgument(operator: string, outer: Stop): MathList {
    const other = operator === "_" ? "^" : operator === "^" ? "_" : null;
    // Function apply binds looser than scripts and limits: `lim┬(n→∞)⁡a`.
    const stops: Stop = (token) => plain(token, FUNCTION_APPLY) || outer(token);
    const lead: MathList = [];
    const sign = this.peek();
    if (sign && !sign.lit && sign.nor === undefined && SCRIPT_SIGNS.has(sign.ch)) {
      this.pos += 1;
      pushText(lead, sign.ch, sign.style);
    }
    if (!lead.length && this.peek() && !this.peek()!.lit && PRIMES[this.peek()!.ch]) {
      let primes = "";
      while (this.peek() && !this.peek()!.lit && PRIMES[this.peek()!.ch]) {
        primes += PRIMES[this.peek()!.ch];
        this.pos += 1;
      }
      return [{ kind: "text", text: primes, style: {} }];
    }
    const single = this.peek();
    if (
      !lead.length &&
      single &&
      !single.lit &&
      single.nor === undefined &&
      !stops(single) &&
      !this.startsOperand(0, stops) &&
      single.ch !== " " &&
      !CLOSERS.has(single.ch) &&
      !this.isFractionOperator(single) &&
      !["_", "^", "&", "@", NARY_AND, FUNCTION_APPLY, CLOSE_DELIMITER, CLOSE_GROUP, VERTICAL_SEPARATOR, BELOW, ABOVE].includes(single.ch) &&
      !isCombiningMark(single.ch)
    ) {
      this.pos += 1;
      return [{ kind: "text", text: single.ch, style: single.style }];
    }
    const operand = this.startsOperand(0, stops)
      ? this.parseOperand(stops, other)
      : { nodes: [], factors: [] };
    if (lead.length) {
      appendList(lead, operand.nodes);
      return lead;
    }
    return stripOperand(operand);
  }

  private parseLimit(unit: MathList, operator: string, stops: Stop): MathNode {
    const next = this.peek();
    let argument: MathList;
    if (next && !next.lit && next.nor === undefined && STRETCH_ARROWS.has(next.ch)) {
      this.pos += 1;
      argument = [{ kind: "text", text: next.ch, style: next.style }];
    } else argument = this.parseScriptArgument(operator, stops);
    const arrow = (list: MathList) =>
      list.length === 1 &&
      list[0].kind === "text" &&
      !list[0].style.nor &&
      Array.from(list[0].text).length === 1 &&
      STRETCH_ARROWS.has(list[0].text)
        ? list[0].text
        : null;
    const baseArrow = arrow(unit);
    if (baseArrow) {
      return operator === ABOVE
        ? { kind: "groupChr", chr: baseArrow, pos: "bot", vertJc: "bot", e: argument }
        : { kind: "groupChr", chr: baseArrow, pos: "top", vertJc: "top", e: argument };
    }
    const limitArrow = next && !next.lit && STRETCH_ARROWS.has(next.ch) ? arrow(argument) : null;
    if (limitArrow) {
      return operator === ABOVE
        ? { kind: "groupChr", chr: limitArrow, pos: "top", vertJc: "bot", e: unit }
        : { kind: "groupChr", chr: limitArrow, pos: "bot", vertJc: "top", e: unit };
    }
    return operator === ABOVE
      ? { kind: "limUpp", e: unit, lim: argument }
      : { kind: "limLow", e: unit, lim: argument };
  }

  /* -------------------------------------------------------------------- */

  private atContainerStart(): boolean {
    const previous = this.tokens[this.pos - 1];
    if (!previous) return true;
    if (previous.lit || previous.nor !== undefined) return false;
    return (
      previous.ch === OPEN_GROUP ||
      OPENERS.has(previous.ch) ||
      ["&", "@", NARY_AND, FUNCTION_APPLY, VERTICAL_SEPARATOR, "_", "^", BELOW, ABOVE].includes(previous.ch)
    );
  }

  /** Whether a closing bracket follows at this bracket level. */
  private closerAhead(): boolean {
    let depth = 0;
    for (let index = this.pos + 1; index < this.tokens.length; index += 1) {
      const token = this.tokens[index];
      if (token.lit || token.nor !== undefined) continue;
      const value = token.ch;
      if (OPENERS.has(value) || value === OPEN_DELIMITER || value === OPEN_GROUP) depth += 1;
      else if (CLOSERS.has(value) || value === CLOSE_DELIMITER) {
        if (depth === 0) return true;
        depth -= 1;
      } else if (value === CLOSE_GROUP) {
        if (depth === 0) return false;
        depth -= 1;
      }
    }
    return false;
  }

  /** Whether another `character` closes a |…| pair at this bracket level. */
  private hasPartnerAhead(character: string): boolean {
    let depth = 0;
    for (let index = this.pos + 1; index < this.tokens.length; index += 1) {
      const token = this.tokens[index];
      if (token.lit || token.nor !== undefined) continue;
      const value = token.ch;
      if (OPENERS.has(value) || value === OPEN_DELIMITER || value === OPEN_GROUP) depth += 1;
      else if (CLOSERS.has(value) || value === CLOSE_DELIMITER || value === CLOSE_GROUP) {
        if (depth === 0) return false;
        depth -= 1;
      } else if (depth === 0 && value === character) return true;
    }
    return false;
  }

  private parseBracket(): Factor | null {
    const start = this.pos;
    const opener = this.tokens[this.pos++];
    let beg = opener.ch;
    let explicit = false;
    if (opener.ch === OPEN_DELIMITER) {
      explicit = true;
      const next = this.peek();
      if (plain(next, " ")) {
        beg = "";
        this.pos += 1;
      } else if (next && !next.lit && next.nor === undefined && DELIMITER_CHARACTERS.has(next.ch) && !OPENERS.has(next.ch)) {
        beg = next.ch;
        this.pos += 1;
      } else beg = "";
    }
    const args: MathList[] = [];
    let sep: string | null = null;
    let end: string | null = null;
    for (;;) {
      const argumentStart = this.pos;
      const inner: Stop = (token) => {
        if (token.lit || token.nor !== undefined) return false;
        const value = token.ch;
        if (CLOSERS.has(value) || value === CLOSE_DELIMITER) return true;
        if (value === VERTICAL_SEPARATOR || value === DIVIDES || value === MIDDLE) return true;
        if (SELF_DELIMITERS.has(value)) {
          if (value === beg && !explicit) return this.pos > argumentStart;
          return !this.hasPartnerAhead(value);
        }
        return false;
      };
      args.push(this.parseExp(inner));
      const token = this.peek();
      if (!token || token.lit || token.nor !== undefined) break;
      const value = token.ch;
      if (value === VERTICAL_SEPARATOR || value === DIVIDES) {
        sep ??= value === VERTICAL_SEPARATOR ? "|" : DIVIDES;
        this.pos += 1;
        continue;
      }
      if (value === MIDDLE && this.peek(1)) {
        sep ??= this.peek(1)!.ch;
        this.pos += 2;
        continue;
      }
      if (value === CLOSE_DELIMITER) {
        this.pos += 1;
        const next = this.peek();
        const empty = args.every((arg) => arg.length === 0);
        if (
          next &&
          !next.lit &&
          next.nor === undefined &&
          DELIMITER_CHARACTERS.has(next.ch) &&
          (!CLOSERS.has(next.ch) || (empty && !explicit))
        ) {
          // `(┤)` is Word's spelling of an empty pair: the closer is the end.
          end = next.ch;
          this.pos += 1;
        } else end = "";
        explicit = true;
        break;
      }
      if (CLOSERS.has(value) || (SELF_DELIMITERS.has(value) && value === beg)) {
        end = value;
        this.pos += 1;
        break;
      }
      if (SELF_DELIMITERS.has(value)) {
        // `(a|b)` separates; `(a|` with no closer left ends the pair.
        if (!this.closerAhead()) {
          end = value;
          this.pos += 1;
          break;
        }
        sep ??= value;
        this.pos += 1;
        continue;
      }
      break;
    }
    if (end === null) {
      this.pos = start;
      return null;
    }
    const node: MathNode = { kind: "d", beg, end, sep: sep ?? "|", args };
    const paren = !explicit && beg === "(" && end === ")" && args.length === 1;
    return { nodes: [node], kind: paren ? "paren" : "other" };
  }

  /** `(_sub^sup)base` — Word's prescript. */
  private parsePrescript(stops: Stop): Factor | null {
    const start = this.pos;
    this.pos += 1;
    const inner: Stop = (token) => plain(token, ")") || stops(token);
    let sub: MathList | null = null;
    let sup: MathList | null = null;
    for (let index = 0; index < 2; index += 1) {
      if (sub === null && plain(this.peek(), "_")) {
        this.pos += 1;
        sub = this.parseScriptArgument("_", inner);
      } else if (sup === null && plain(this.peek(), "^")) {
        this.pos += 1;
        sup = this.parseScriptArgument("^", inner);
      }
    }
    if (sub === null || sup === null || !plain(this.peek(), ")")) {
      this.pos = start;
      return null;
    }
    this.pos += 1;
    const core = this.parseFactorCore(stops);
    let e: MathList = core ? stripFactor(core) : [];
    let rest: MathList = [];
    if (core && core.kind === "other" && e.length === 1 && e[0].kind === "text") {
      const characters = Array.from(e[0].text);
      if (characters.length > 1) {
        rest = [{ ...e[0], text: characters.slice(1).join("") }];
        e = [{ ...e[0], text: characters[0] }];
      }
    }
    return { nodes: [{ kind: "sPre", e, sub, sup }, ...rest], kind: "other" };
  }

  private parseGroup(): Factor | null {
    const start = this.pos;
    this.pos += 1;
    const content = this.parseExp((token) => plain(token, CLOSE_GROUP));
    if (!plain(this.peek(), CLOSE_GROUP)) {
      this.pos = start;
      return null;
    }
    this.pos += 1;
    return { nodes: content, kind: "group" };
  }

  private parsePrefix(stops: Stop): Factor {
    const operator = this.tokens[this.pos++].ch;
    if (operator === "√" && plain(this.peek(), "(")) {
      const root = this.tryMaskedArgument(stops);
      if (root) {
        return { nodes: [{ kind: "rad", deg: root.first, e: root.second }], kind: "other" };
      }
    }
    let mask: number | null = null;
    if ((operator === "□" || operator === "▭" || PHANTOMS.has(operator)) && plain(this.peek(), "(")) {
      const masked = this.tryMaskedArgument(stops);
      if (masked) {
        const digits = masked.first.length === 1 && masked.first[0].kind === "text" ? masked.first[0].text : "";
        mask = /^[0-9]+$/.test(digits) ? Number(digits) : 0;
        return { nodes: [this.prefixNode(operator, masked.second, mask)], kind: "other" };
      }
    }
    const tight = operator === "¯" || operator === "▁" || GROUP_CHARACTERS.has(operator);
    let argument: MathList = [];
    if (tight) {
      const factor = this.parseFactorCore(stops);
      if (factor) argument = stripFactor(factor);
    } else if (this.startsOperand(0, stops)) {
      argument = stripOperand(this.parseOperand(stops, null));
    }
    return { nodes: [this.prefixNode(operator, argument, mask)], kind: "other" };
  }

  /** `(first&second)` after a prefix operator; null when there is no `&`. */
  private tryMaskedArgument(stops: Stop): { first: MathList; second: MathList } | null {
    const start = this.pos;
    this.pos += 1;
    const first = this.parseExp((token) => plain(token, "&") || plain(token, ")") || stops(token));
    if (!plain(this.peek(), "&")) {
      this.pos = start;
      return null;
    }
    this.pos += 1;
    const second = this.parseExp((token) => plain(token, ")") || stops(token));
    if (!plain(this.peek(), ")")) {
      this.pos = start;
      return null;
    }
    this.pos += 1;
    return { first, second };
  }

  private prefixNode(operator: string, argument: MathList, mask: number | null): MathNode {
    switch (operator) {
      case "√":
        return { kind: "rad", deg: [], e: argument };
      case "∛":
        return { kind: "rad", deg: [{ kind: "text", text: "3", style: {} }], e: argument };
      case "∜":
        return { kind: "rad", deg: [{ kind: "text", text: "4", style: {} }], e: argument };
      case "□":
        return { kind: "box", e: argument };
      case "▭":
        return { kind: "borderBox", flags: [], e: argument };
      case "¯":
        return { kind: "bar", pos: "top", e: argument };
      case "▁":
        return { kind: "bar", pos: "bot", e: argument };
    }
    const group = GROUP_CHARACTERS.get(operator);
    if (group) return { kind: "groupChr", chr: operator, pos: group.pos, vertJc: group.vertJc, e: argument };
    const phantom =
      mask !== null && operator === "⟡"
        ? {
            show: Boolean(mask & 1),
            zeroWid: Boolean(mask & 2),
            zeroAsc: Boolean(mask & 4),
            zeroDesc: Boolean(mask & 8),
          }
        : (PHANTOMS.get(operator) ?? PHANTOMS.get("⟡")!);
    return { kind: "phant", ...phantom, transp: false, e: argument };
  }

  /** One factor without postfix operators (the operand of ¯ ▁ ⏞ ⏟). */
  private parseFactorCore(stops: Stop): Factor | null {
    const token = this.peek();
    if (!token || stops(token)) return null;
    if (token.nor !== undefined) {
      this.pos += 1;
      return { nodes: [{ kind: "text", text: token.nor, style: { nor: true } }], kind: "other" };
    }
    if (this.isOperandToken(token)) return this.parseAtoms(stops);
    const character = token.ch;
    if (OPENERS.has(character) || character === OPEN_DELIMITER) return this.parseBracket();
    if (SELF_DELIMITERS.has(character) && (this.hasPartnerAhead(character) || this.closerAhead())) {
      return this.parseBracket();
    }
    if (character === OPEN_GROUP) return this.parseGroup();
    if (PREFIX.has(character)) return this.parsePrefix(stops);
    if (MATRIX_FENCES.has(character) || character === EQUATION_ARRAY || character === CASES) {
      return this.parseArray();
    }
    if (NARY_OPERATORS.has(character)) return this.parseNary(stops);
    return null;
  }

  private parseArray(): Factor | null {
    const start = this.pos;
    const introducer = this.tokens[this.pos++].ch;
    if (!plain(this.peek(), "(")) {
      this.pos = start;
      return null;
    }
    this.pos += 1;
    const equationArray = introducer === EQUATION_ARRAY || introducer === CASES;
    const rows: MathList[][] = [[]];
    for (;;) {
      const cell = this.parseExp(
        (token) => plain(token, "@") || plain(token, ")") || (!equationArray && plain(token, "&")),
      );
      rows[rows.length - 1].push(cell);
      const token = this.peek();
      if (plain(token, "&") && !equationArray) {
        this.pos += 1;
        continue;
      }
      if (plain(token, "@")) {
        this.pos += 1;
        rows.push([]);
        continue;
      }
      if (plain(token, ")")) {
        this.pos += 1;
        break;
      }
      this.pos = start;
      return null;
    }
    let node: MathNode;
    if (equationArray) {
      node = { kind: "eqArr", rows: rows.map((row) => row[0] ?? []) };
      if (introducer === CASES) node = { kind: "d", beg: "{", end: "", sep: "|", args: [[node]] };
      return { nodes: [node], kind: "other" };
    }
    const columns = Math.max(...rows.map((row) => row.length));
    node = {
      kind: "m",
      rows: rows.map((row) => [...row, ...Array.from({ length: columns - row.length }, () => [])]),
      columnJc: [],
    };
    const fence = MATRIX_FENCES.get(introducer);
    if (fence) node = { kind: "d", beg: fence[0], end: fence[1], sep: "|", args: [[node]] };
    return { nodes: [node], kind: "other" };
  }

  private parseNary(stops: Stop): Factor {
    const chr = this.tokens[this.pos++].ch;
    let limLoc: "undOvr" | "subSup" | null = null;
    if (plain(this.peek(), "1")) {
      limLoc = "undOvr";
      this.pos += 1;
    } else if (plain(this.peek(), "2")) {
      limLoc = "subSup";
      this.pos += 1;
    }
    let sub: MathList | null = null;
    let sup: MathList | null = null;
    for (let index = 0; index < 2; index += 1) {
      if (sub === null && plain(this.peek(), "_")) {
        this.pos += 1;
        sub = this.parseScriptArgument("_", stops);
      } else if (sup === null && plain(this.peek(), "^")) {
        this.pos += 1;
        sup = this.parseScriptArgument("^", stops);
      }
    }
    let e: MathList = [];
    if (plain(this.peek(), NARY_AND)) {
      this.pos += 1;
      if (this.isSpaceAt()) this.pos += 1;
      else if (this.startsOperand(0, stops)) e = groupOnly(this.parseOperand(stops, null));
    } else if (this.isSpaceAt() && this.startsOperand(1, stops)) {
      this.pos += 1;
      e = groupOnly(this.parseOperand(stops, null));
    }
    return { nodes: [{ kind: "nary", chr, limLoc, sub, sup, e }], kind: "other" };
  }
}

/** Superscripts written as trailing primes (x′), which a following ^ joins. */
const PRIMED = new WeakSet<MathNode>();

const PREFIX = new Set([
  "√", "∛", "∜", "□", "▭", "▁", "¯", ...PHANTOMS.keys(), ...GROUP_CHARACTERS.keys(),
]);

function tokensToText(tokens: readonly Token[]): MathList {
  const list: MathList = [];
  for (const token of tokens) {
    pushText(list, token.ch, token.lit ? { ...token.style, lit: true } : token.style);
  }
  return list;
}

function stripFactor(factor: Factor): MathList {
  if (factor.kind === "paren" && factor.nodes[0]?.kind === "d") return factor.nodes[0].args[0];
  return factor.nodes;
}

/** An argument's outer parentheses (or 〖〗) disappear on build-up. */
function stripOperand(operand: Operand): MathList {
  if (operand.factors.length === 1) return stripFactor(operand.factors[0]);
  return operand.nodes;
}

function groupOnly(operand: Operand): MathList {
  return operand.nodes;
}

/** The trailing unit a postfix operator applies to. */
function splitUnit(nodes: MathList, kind: FactorKind): { rest: MathList; unit: MathList } {
  if (kind === "group") return { rest: [], unit: nodes };
  const last = nodes.at(-1);
  if (!last) return { rest: [], unit: [] };
  if (last.kind !== "text" || last.style.nor) return { rest: nodes.slice(0, -1), unit: [last] };
  const characters = Array.from(last.text);
  let start = characters.length - 1;
  if (/^[0-9]$/.test(characters[start] ?? "")) {
    while (start > 0 && /^[0-9.,]$/.test(characters[start - 1])) start -= 1;
    while (start < characters.length - 1 && !/^[0-9]$/.test(characters[start])) start += 1;
  }
  const head = characters.slice(0, start).join("");
  const rest = nodes.slice(0, -1);
  if (head) rest.push({ ...last, text: head });
  return { rest, unit: [{ ...last, text: characters.slice(start).join("") }] };
}

/** The function name before an explicit function-apply character. */
function splitName(nodes: MathList, kind: FactorKind): { rest: MathList; unit: MathList } {
  if (kind === "group") return { rest: [], unit: nodes };
  const last = nodes.at(-1);
  if (!last || last.kind !== "text" || last.style.nor) return splitUnit(nodes, kind);
  const characters = Array.from(last.text);
  let start = characters.length;
  while (start > 0 && /^\p{L}$/u.test(characters[start - 1])) start -= 1;
  if (start === characters.length) return splitUnit(nodes, kind);
  const rest = nodes.slice(0, -1);
  if (start > 0) rest.push({ ...last, text: characters.slice(0, start).join("") });
  return { rest, unit: [{ ...last, text: characters.slice(start).join("") }] };
}

/** Build up UnicodeMath linear text into the OMML model. */
export function parseUnicodeMath(text: string): MathList {
  // A line break in a display equation's linear text breaks the equation
  // (UTN #28 §3.23): the lines become the rows of an equation array.
  const lines = text.split(/\r\n|\r|\n/u).filter((line) => line.trim() !== "");
  if (lines.length > 1) {
    return [{ kind: "eqArr", rows: lines.map((line) => new Parser(tokenize(line)).parse()) }];
  }
  return new Parser(tokenize(lines[0] ?? "")).parse();
}
