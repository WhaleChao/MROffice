/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * OMML model → UnicodeMath, following Word's Linearize: argument
 * parenthesization, 〖〗 grouping, operand-terminating spaces and the
 * code points Word writes for styled letters were taken from Word 16.86's
 * linearization of ~1,600 sample equations.
 *
 * "safe" mode wraps every compound argument; callers use it when Word's
 * terser spelling would not build back up to the same equation.
 */

import type { MathList, MathNode, MathRunStyle } from "./ast";
import { pairedPositions } from "./brackets";
import {
  ABOVE,
  BELOW,
  DELIMITER_CHARACTERS,
  BUILD_UP_CHARACTERS,
  CLOSE_DELIMITER,
  CLOSE_GROUP,
  CLOSERS,
  FUNCTION_APPLY,
  GROUP_CHARACTERS,
  NARY_AND,
  NBSP,
  OPEN_DELIMITER,
  OPEN_GROUP,
  OPENERS,
  PHANTOMS,
  RECOGNIZED_FUNCTIONS,
  SELF_DELIMITERS,
  STRETCH_ARROWS,
  VERTICAL_SEPARATOR,
  defaultLimitLocation,
  encodeMathCharacter,
  isCombiningMark,
  isOperandCharacter,
  unicodeSubscript,
  unicodeSuperscript,
} from "./characters";

export class UnsupportedUnicodeMath extends Error {}

export type SerializeMode = "word" | "safe";

type Category = "operator" | "operand" | "open" | "close" | "object" | "nor" | "space" | "none";

interface Piece {
  readonly text: string;
  readonly first: Category;
  readonly last: Category;
  readonly lead: number;
  readonly trail: ReadonlySet<Category>;
  readonly endSpace: boolean;
  readonly zoneEndSpace: boolean;
  readonly consumesSpace: boolean;
  /** A delimiter object: no end space inside another delimiter or function. */
  readonly delimiter?: boolean;
}

type Container =
  | "zone"
  | "num"
  | "den"
  | "sup"
  | "sub"
  | "base"
  | "rad"
  | "d"
  | "naryE"
  | "narySub"
  | "funcE"
  | "accE"
  | "barE"
  | "lim"
  | "cell"
  | "cellEnd"
  | "row"
  | "boxE"
  | "deg"
  | "sPreE"
  | "groupE"
  | "fname";

const END_SPACE: ReadonlySet<Container> = new Set([
  "den", "sup", "sub", "rad", "d", "naryE", "funcE", "accE", "barE", "lim",
  "cellEnd", "boxE", "sPreE", "groupE",
]);
const OPEN_TRAIL: ReadonlySet<Category> = new Set(["operand", "open", "object", "nor"]);
const NO_TRAIL: ReadonlySet<Category> = new Set();

/* Word's parenthesization of a one-object argument, by container. */
const NUM_PAREN = new Set(["sSubSup", "nary", "naryEmpty", "f", "limLow", "d", "limUpp", "skw", "groupArrow"]);
const DEN_PAREN = new Set(["sSubSup", "nary", "naryEmpty", "f", "d", "skw"]);
const SCRIPT_PAREN = new Set(["sSup", "sSub", "sSubSup", "f", "d", "skw"]);
const BASE_BARE = new Set(["acc", "bar", "d", "groupChr", "m", "eqArr"]);
const FUNC_GROUP = new Set(["sSup", "sSub", "sSubSup", "f", "skw"]);
const LIM_PAREN = new Set(["sSup", "sSub", "sSubSup", "f", "d", "limLow", "limUpp", "groupChr", "skw", "groupArrow"]);
const NARY_SUB_PAREN = new Set(["sSup", "sSub", "sSubSup", "rad", "f", "d", "skw", "radDeg"]);

const SPACING_ACCENTS: Readonly<Record<string, string>> = {
  "^": "̂",
  "ˆ": "̂",
  "~": "̃",
  "˜": "̃",
  "¯": "̅",
  "‾": "̅",
  "˙": "̇",
  "¨": "̈",
  "ˇ": "̌",
  "˘": "̆",
  "´": "́",
  "`": "̀",
  "˚": "̊",
  "→": "⃗",
  "←": "⃖",
  "↔": "⃡",
};

export function combiningAccent(chr: string): string | null {
  if (isCombiningMark(chr)) return chr;
  return SPACING_ACCENTS[chr] ?? null;
}

const PRIME_CHARACTERS = new Set(["'", "′", "″", "‴", "⁗"]);
const WORD_BRACKETS: Readonly<Record<string, string>> = { "〈": "⟨", "〉": "⟩" };

function characterCategory(character: string | undefined): Category {
  if (character === undefined) return "none";
  if (character === " ") return "space";
  if (OPENERS.has(character)) return "open";
  if (CLOSERS.has(character)) return "close";
  if (isOperandCharacter(character) || PRIME_CHARACTERS.has(character)) return "operand";
  return "operator";
}

function isOperandText(node: MathNode): boolean {
  if (node.kind !== "text" || node.style.nor || node.style.lit) return false;
  const characters = Array.from(node.text);
  if (characters.length === 0) return false;
  return characters.every((character, index) => {
    if (isOperandCharacter(character)) return true;
    if ((character === "." || character === ",") && index > 0 && index < characters.length - 1) {
      return /^[0-9]$/.test(characters[index - 1]) && /^[0-9]$/.test(characters[index + 1]);
    }
    return false;
  });
}

/** An argument Word writes without parentheses: plain operand characters. */
function isSimpleText(list: MathList): boolean {
  return list.length > 0 && list.every(isOperandText);
}

function textLength(list: MathList): number {
  return list.reduce((total, node) => total + (node.kind === "text" ? Array.from(node.text).length : 0), 0);
}

function isUprightFunctionWord(list: MathList): boolean {
  return (
    list.length === 1 &&
    list[0].kind === "text" &&
    !list[0].style.nor &&
    list[0].style.sty === "p" &&
    (list[0].style.scr === undefined || list[0].style.scr === "roman") &&
    RECOGNIZED_FUNCTIONS.has(list[0].text)
  );
}

function arrowForm(node: MathNode): "above" | "below" | "overArrow" | "underArrow" | null {
  if (node.kind !== "groupChr" || !STRETCH_ARROWS.has(node.chr)) return null;
  if (node.pos === "top" && node.vertJc === "bot") return "above"; // base┴→
  if (node.pos === "bot" && node.vertJc === "top") return "below"; // base┬→
  if (node.pos === "bot" && node.vertJc === "bot") return "overArrow"; // →┴base
  return "underArrow"; // →┬base
}

function nodeClass(node: MathNode): string {
  switch (node.kind) {
    case "nary":
      return node.e.length === 0 ? "naryEmpty" : "nary";
    case "f":
      return node.type === "skw" || node.type === "lin" ? "skw" : "f";
    case "rad":
      return node.deg.length > 0 ? "radDeg" : "rad";
    case "func":
      return node.name.length === 1 && (node.name[0].kind === "limLow" || node.name[0].kind === "limUpp")
        ? "funcLim"
        : "func";
    case "groupChr":
      return arrowForm(node) ? "groupArrow" : "groupChr";
    case "text":
      return node.style.nor ? "nor" : "text";
    default:
      return node.kind;
  }
}

export class Serializer {
  private readonly absorb: boolean;

  constructor(private readonly mode: SerializeMode, options: { absorbNaryands?: boolean } = {}) {
    this.absorb = options.absorbNaryands ?? true;
  }

  zone(list: MathList): string {
    return this.list(list, "zone");
  }

  /* -------------------------------------------------------------------- */

  private list(input: MathList, container: Container, eqArrRow = false): string {
    const merged = mergeRuns(input);
    const list = this.absorb ? absorbNaryands(merged) : merged;
    const pairs = pairedPositions(list);
    const pieces = list.map((node, index) => this.piece(node, index, pairs, eqArrRow));
    let out = "";
    pieces.forEach((piece, index) => {
      const previous = pieces[index - 1];
      if (previous) {
        let spaces = 0;
        if (previous.trail.has(piece.first)) spaces += 1;
        if (piece.lead && previous.last !== "operator" && previous.last !== "space") {
          spaces += piece.lead;
          // One space after an object only ends that object; the operand
          // (and so a numerator) would continue across it.
          if (previous.consumesSpace && !previous.trail.has(piece.first)) spaces += 1;
        }
        if (piece.first === "space" && previous.consumesSpace) spaces += 1;
        if (spaces === 0 && joinsNumber(out, piece.text, previous, piece)) spaces = 1;
        if (spaces === 0 && out.endsWith(CLOSE_DELIMITER) && DELIMITER_CHARACTERS.has(Array.from(piece.text)[0] ?? "")) {
          spaces = 1;
        }
        out += " ".repeat(spaces);
      } else if (piece.first === "space" && container !== "zone") {
        // nothing: a leading space inside an argument is literal
      }
      out += piece.text;
    });
    const last = pieces.at(-1);
    const delimiterInside = last?.delimiter && (container === "d" || container === "funcE");
    if (last && ((END_SPACE.has(container) && last.endSpace && !delimiterInside) || last.zoneEndSpace)) out += " ";
    return out;
  }

  /** Serialize an argument with Word's parentheses/grouping decision. */
  private argument(list: MathList, container: Container): string {
    const body = this.list(list, container);
    const wrap = this.wrapFor(list, container);
    if (wrap === "paren") return `(${body})`;
    if (wrap === "group") return `${OPEN_GROUP}${body}${CLOSE_GROUP}`;
    return body;
  }

  private wrapFor(list: MathList, container: Container): "bare" | "paren" | "group" {
    if (list.length === 0) return "bare";
    if (["d", "cell", "cellEnd", "row", "deg", "sPreE", "fname", "zone"].includes(container)) return "bare";
    const safe = this.mode === "safe";
    const single = list.length === 1 ? list[0] : null;
    const kind = single ? nodeClass(single) : "multi";
    if (
      container === "base" &&
      list.length === 1 &&
      list[0].kind === "text" &&
      !list[0].style.nor &&
      Array.from(list[0].text).length === 1 &&
      list[0].text !== " "
    ) {
      return "bare";
    }
    if (
      (container === "sup" || container === "sub") &&
      list.length === 1 &&
      list[0].kind === "text" &&
      !list[0].style.nor &&
      Array.from(list[0].text).length === 1 &&
      !OPENERS.has(list[0].text) &&
      !CLOSERS.has(list[0].text) &&
      list[0].text !== " " &&
      !BUILD_UP_CHARACTERS.has(list[0].text)
    ) {
      return "bare";
    }
    if (isSimpleText(list)) {
      const length = textLength(list);
      if (container === "base") {
        if (length === 1 || /^[0-9]+$/.test(list.map((node) => (node.kind === "text" ? node.text : "")).join(""))) return "bare";
        return isUprightFunctionWord(list) ? "bare" : "group";
      }
      if (container === "accE" || container === "barE" || container === "groupE") {
        return length === 1 ? "bare" : "paren";
      }
      return "bare";
    }
    switch (container) {
      case "num":
        return !safe && single && !NUM_PAREN.has(kind) && kind !== "text" ? "bare" : "paren";
      case "den":
        return !safe && single && !DEN_PAREN.has(kind) && kind !== "text" ? "bare" : "paren";
      case "sup":
      case "sub":
        return !safe && single && !SCRIPT_PAREN.has(kind) && kind !== "text" ? "bare" : "paren";
      case "base":
        return single && (BASE_BARE.has(kind) || (!safe && kind === "acc")) ? "bare" : "group";
      case "rad":
      case "boxE":
        return "paren";
      case "accE":
        return single && kind === "acc" ? "bare" : "paren";
      case "barE":
      case "groupE":
        return "paren";
      case "lim":
        return !safe && single && !LIM_PAREN.has(kind) && kind !== "text" ? "bare" : "paren";
      case "narySub":
        return !safe && single && !NARY_SUB_PAREN.has(kind) && kind !== "text" ? "bare" : "paren";
      case "funcE":
        if (single) return FUNC_GROUP.has(kind) || kind === "text" || safe ? "group" : "bare";
        return !safe && list.every((node) => isOperandText(node) || node.kind === "d" || (node.kind === "text" && node.style.nor))
          ? "bare"
          : "group";
      case "naryE":
        if (single) return kind === "text" ? "group" : "bare";
        return !safe && list.every((node) => isOperandText(node) || node.kind === "d" || (node.kind === "text" && node.style.nor))
          ? "bare"
          : "group";
      default:
        return "paren";
    }
  }

  /* -------------------------------------------------------------------- */

  private piece(node: MathNode, index: number, pairs: Set<string>, eqArrRow: boolean): Piece {
    switch (node.kind) {
      case "text":
        return node.style.nor
          ? this.norPiece(node.text, node.style)
          : this.textPiece(node.text, node.style, index, pairs, eqArrRow);
      case "f": {
        const operator = node.type === "skw" ? "⁄" : node.type === "lin" ? "∕" : node.type === "noBar" ? "¦" : "/";
        return objectPiece(`${this.argument(node.num, "num")}${operator}${this.argument(node.den, "den")}`, { lead: 1 });
      }
      case "sSup":
      case "sSub":
      case "sSubSup": {
        const base = node.e.length === 0 ? "" : this.argument(node.e, "base");
        let scripts = "";
        if (node.kind === "sSub" || node.kind === "sSubSup") scripts += `_${this.argument(node.sub, "sub")}`;
        if (node.kind === "sSup" || node.kind === "sSubSup") scripts += `^${this.argument(node.sup, "sup")}`;
        const text =
          node.e.length === 0
            ? `${OPEN_GROUP}${scripts}${node.kind === "sSubSup" ? " " : ""}${CLOSE_GROUP}`
            : base + scripts;
        return objectPiece(text, { endSpace: true, trail: node.e.length === 0 ? NO_TRAIL : OPEN_TRAIL });
      }
      case "sPre": {
        const base = node.e.length === 1 && node.e[0].kind === "text" && textLength(node.e) > 1
          ? `${OPEN_GROUP}${this.list(node.e, "sPreE")}${CLOSE_GROUP}`
          : this.list(node.e, "sPreE");
        const empty = node.e.length === 0 ? " " : "";
        return objectPiece(`(_${this.argument(node.sub, "sub")}^${this.argument(node.sup, "sup")})${base}${empty}`, {
          endSpace: true,
          trail: new Set<Category>(["open", "object"]),
        });
      }
      case "rad": {
        const quirk = prefixArgumentSpace(node.e);
        const degree = node.deg.length === 1 && node.deg[0].kind === "text" && !node.deg[0].style.nor ? node.deg[0].text : null;
        if (node.deg.length === 0) return objectPiece(`√${this.argument(node.e, "rad")}${quirk}`);
        if (degree === "3") return objectPiece(`∛${this.argument(node.e, "rad")}${quirk}`);
        if (degree === "4") return objectPiece(`∜${this.argument(node.e, "rad")}${quirk}`);
        return objectPiece(`√(${this.list(node.deg, "deg")}&${this.list(node.e, "cellEnd")})`);
      }
      case "nary": {
        let text = node.chr;
        const location = node.limLoc ?? defaultLimitLocation(node.chr);
        if (location !== defaultLimitLocation(node.chr)) text += location === "undOvr" ? "1" : "2";
        if (node.sub && node.sub.length) text += `_${this.argument(node.sub, "narySub")}`;
        if (node.sup && node.sup.length) text += `^${this.argument(node.sup, "narySub")}`;
        if (node.e.length) text += `${NARY_AND}${this.argument(node.e, "naryE")}`;
        return objectPiece(text);
      }
      case "d":
        return this.delimiterPiece(node);
      case "m": {
        if (node.columnJc.some((jc) => jc !== "center")) throw new UnsupportedUnicodeMath("matrix column alignment");
        return objectPiece(`■(${this.rows(node.rows)})`, { trail: NO_TRAIL, last: "close" });
      }
      case "eqArr":
        return objectPiece(
          `█(${node.rows.map((row, rowIndex) => this.list(row, rowIndex === node.rows.length - 1 ? "cellEnd" : "row", true)).join("@")})`,
          { trail: NO_TRAIL, last: "close" },
        );
      case "func": {
        const limit = nodeClass(node) === "funcLim";
        return objectPiece(`${this.list(node.name, "fname")}${FUNCTION_APPLY}${this.argument(node.e, "funcE")}`, {
          lead: limit ? 2 : 1,
          endSpace: true,
        });
      }
      case "limLow":
      case "limUpp": {
        const operator = node.kind === "limLow" ? BELOW : ABOVE;
        if (node.e.length === 0) {
          return objectPiece(`${OPEN_GROUP}${operator}${this.argument(node.lim, "lim")}${CLOSE_GROUP}`, {
            endSpace: true,
            trail: NO_TRAIL,
          });
        }
        return objectPiece(`${this.limitBase(node.e)}${operator}${this.argument(node.lim, "lim")}`, {
          lead: 1,
          endSpace: true,
        });
      }
      case "acc": {
        const mark = combiningAccent(node.chr);
        if (!mark) throw new UnsupportedUnicodeMath("accent");
        const single = node.e.length === 1 && node.e[0].kind === "text" && !node.e[0].style.nor && Array.from(node.e[0].text).length === 1;
        const nested = node.e.length === 1 && node.e[0].kind === "acc";
        const base = single
          ? this.list(node.e, "accE").trimEnd()
          : nested
            ? `${this.list(node.e, "accE").trimEnd()} `
            : `(${this.list(node.e, "accE")})`;
        return objectPiece(`${base}${NBSP}${mark}`, { trail: NO_TRAIL, endSpace: true });
      }
      case "bar":
        return objectPiece(`${node.pos === "top" ? "¯" : "▁"}${this.argument(node.e, "barE")}${prefixArgumentSpace(node.e)}`);
      case "groupChr": {
        const form = arrowForm(node);
        if (form === "above") return objectPiece(`${this.limitBase(node.e)}${ABOVE}${node.chr}`, { lead: 1, endSpace: true });
        if (form === "below") return objectPiece(`${this.limitBase(node.e)}${BELOW}${node.chr}`, { lead: 1, endSpace: true });
        if (form === "overArrow") return objectPiece(`${node.chr}${ABOVE}${this.argument(node.e, "lim")}`, { lead: 1, endSpace: true });
        if (form === "underArrow") return objectPiece(`${node.chr}${BELOW}${this.argument(node.e, "lim")}`, { lead: 1, endSpace: true });
        const group = GROUP_CHARACTERS.get(node.chr);
        if (!group || group.pos !== node.pos || group.vertJc !== node.vertJc) {
          throw new UnsupportedUnicodeMath("group character");
        }
        return objectPiece(`${node.chr}${this.argument(node.e, "groupE")}`, { endSpace: true });
      }
      case "borderBox":
        if (node.flags.length) throw new UnsupportedUnicodeMath("border box options");
        return objectPiece(`▭${this.argument(node.e, "boxE")}`);
      case "box":
        return objectPiece(`□${this.argument(node.e, "boxE")}${prefixArgumentSpace(node.e)}`);
      case "phant": {
        const operator = [...PHANTOMS].find(
          ([, flags]) =>
            flags.show === node.show &&
            flags.zeroWid === node.zeroWid &&
            flags.zeroAsc === node.zeroAsc &&
            flags.zeroDesc === node.zeroDesc,
        )?.[0];
        if (operator && !node.transp) return objectPiece(`${operator}${this.argument(node.e, "boxE")}`);
        const mask =
          (node.show ? 1 : 0) + (node.zeroWid ? 2 : 0) + (node.zeroAsc ? 4 : 0) + (node.zeroDesc ? 8 : 0) + (node.transp ? 16 : 0);
        return objectPiece(`⟡(${mask}&${this.list(node.e, "cellEnd")})`);
      }
    }
  }

  private rows(rows: MathList[][]): string {
    return rows
      .map((row, rowIndex) =>
        row
          .map((cell, cellIndex) =>
            this.list(cell, rowIndex === rows.length - 1 && cellIndex === row.length - 1 ? "cellEnd" : "cell"),
          )
          .join("&"),
      )
      .join("@");
  }

  private limitBase(base: MathList): string {
    if (isUprightFunctionWord(base)) return this.list(base, "base");
    // ┬ and ┴ need an operand base; an operator character is grouped.
    const only = base.length === 1 && base[0].kind === "text" && !base[0].style.nor ? base[0].text : null;
    if (only && Array.from(only).length === 1 && !isOperandCharacter(only)) {
      return `${OPEN_GROUP}${this.list(base, "base")}${CLOSE_GROUP}`;
    }
    return this.argument(base, "base");
  }

  private delimiterPiece(input: Extract<MathNode, { kind: "d" }>): Piece {
    let node = input;
    // Word's cases: a brace on one side of left-aligned rows.
    const only = node.args.length === 1 && node.args[0].length === 1 ? node.args[0][0] : null;
    if (
      only &&
      only.kind === "m" &&
      node.beg === "{" &&
      node.end === "" &&
      only.rows.every((row) => row.length <= 2) &&
      only.columnJc.length > 0 &&
      only.columnJc.every((jc) => jc === "left")
    ) {
      const rows = only.rows.map((row) => row.map((cell) => this.list(cell, "row", true)).join("&")).join("@");
      return objectPiece(`{█(${rows})${CLOSE_DELIMITER}`, { trail: NO_TRAIL, last: "close", endSpace: true });
    }
    const empty = node.args.every((arg) => arg.length === 0);
    node = { ...node, beg: WORD_BRACKETS[node.beg] ?? node.beg, end: WORD_BRACKETS[node.end] ?? node.end };
    const selfPair = SELF_DELIMITERS.has(node.beg) && node.beg === node.end;
    let open: string;
    if (node.beg === "") open = `${OPEN_DELIMITER} `;
    else if (OPENERS.has(node.beg) || (selfPair && !empty)) open = node.beg;
    else open = `${OPEN_DELIMITER}${node.beg}`;
    let close: string;
    if (node.end === "") close = CLOSE_DELIMITER;
    else if (empty && node.args.length === 1 && CLOSERS.has(node.end) && OPENERS.has(node.beg)) {
      close = `${CLOSE_DELIMITER}${node.end}`;
    }
    else if (CLOSERS.has(node.end) || (selfPair && !empty)) close = node.end;
    else close = `${CLOSE_DELIMITER}${node.end}`;
    const separator =
      node.sep === "|" ? VERTICAL_SEPARATOR : node.sep === "∣" ? "∣" : `ⓜ${node.sep}`;
    let body = node.args
      .map((arg, index) => this.list(arg, index === node.args.length - 1 ? "d" : "cell"))
      .join(separator);
    // `┤` takes a following delimiter character as its own end.
    if (body.endsWith(CLOSE_DELIMITER) && DELIMITER_CHARACTERS.has(Array.from(close)[0] ?? "")) body += " ";
    return {
      text: `${open}${body}${close}`,
      first: "open",
      last: node.end === "" ? "object" : "close",
      lead: 0,
      trail: new Set<Category>(node.end === "" ? ["open"] : ["object"]),
      endSpace: true,
      delimiter: true,
      zoneEndSpace: false,
      consumesSpace: true,
    };
  }

  private norPiece(text: string, style: MathRunStyle): Piece {
    const encoded = Array.from(text)
      .map((character) => {
        if (character === '"' || character === "\\") return `\\${character}`;
        return encodeMathCharacter(character, { sty: style.sty, scr: style.scr });
      })
      .join("");
    return {
      text: `"${encoded}"`,
      first: "nor",
      last: "nor",
      lead: 0,
      trail: new Set<Category>(["operand", "open", "object", "nor", "close"]),
      endSpace: true,
      zoneEndSpace: true,
      consumesSpace: true,
    };
  }

  private textPiece(
    text: string,
    style: MathRunStyle,
    index: number,
    pairs: Set<string>,
    eqArrRow: boolean,
  ): Piece {
    const characters = Array.from(text);
    let out = "";
    characters.forEach((character, position) => {
      const paired = pairs.has(`${index}:${position}`);
      if (paired) {
        // Word does not pair the CJK angle brackets; ⟨ ⟩ are the same glyphs.
        out += WORD_BRACKETS[character] ?? character;
        return;
      }
      const special =
        BUILD_UP_CHARACTERS.has(character) ||
        isCombiningMark(character) ||
        unicodeSuperscript(character) !== null ||
        unicodeSubscript(character) !== null ||
        character === "@";
      if (special && !(eqArrRow && character === "&")) {
        out += `\\${character}`;
        return;
      }
      if (character === "-" && !style.lit) {
        out += "−";
        return;
      }
      out += encodeMathCharacter(character, style);
    });
    const pairedFirst = pairs.has(`${index}:0`);
    const pairedLast = pairs.has(`${index}:${characters.length - 1}`) && characters.length > 0;
    let first = characterCategory(characters[0]);
    let last = characterCategory(characters.at(-1));
    if (pairedFirst && (OPENERS.has(characters[0]) || SELF_DELIMITERS.has(characters[0]))) first = "open";
    else if (first === "open") first = "operator";
    if (pairedLast && (CLOSERS.has(characters.at(-1)!) || SELF_DELIMITERS.has(characters.at(-1)!))) last = "close";
    else if (last === "close") last = "operator";
    // A paired closing bracket builds up into a delimiter object: it takes
    // the delimiter's terminating space before a following object.
    const closes = last === "close";
    return {
      text: out,
      first,
      last,
      lead: 0,
      trail: closes ? new Set<Category>(["object"]) : NO_TRAIL,
      endSpace: false,
      zoneEndSpace: false,
      consumesSpace: closes,
    };
  }
}

function objectPiece(
  text: string,
  options: {
    lead?: number;
    trail?: ReadonlySet<Category>;
    endSpace?: boolean;
    last?: Category;
  } = {},
): Piece {
  return {
    text,
    first: "object",
    last: options.last ?? "object",
    lead: options.lead ?? 0,
    trail: options.trail ?? OPEN_TRAIL,
    endSpace: options.endSpace ?? false,
    zoneEndSpace: false,
    consumesSpace: true,
  };
}

/**
 * Word writes a summand inside the n-ary object and a limit function's
 * argument after the function-apply character; LaTeX (and our exporter)
 * leaves both after the operator. Move the following term into an empty
 * n-ary operand, or make `lim┬…` a function with it as the argument,
 * stopping at the first operator or space outside brackets.
 */
function absorbNaryands(input: MathList): MathList {
  const list = splitFunctionWords(input);
  const out: MathList = [];
  for (let index = 0; index < list.length; index += 1) {
    const node = list[index];
    const nary = node.kind === "nary" && node.e.length === 0;
    const limit =
      (node.kind === "limLow" ||
        node.kind === "limUpp" ||
        node.kind === "sSub" ||
        node.kind === "sSup" ||
        node.kind === "sSubSup" ||
        node.kind === "text") &&
      (node.kind === "text" ? isUprightFunctionWord([node]) : isUprightFunctionWord(node.e));
    if (!nary && !limit) {
      out.push(node);
      continue;
    }
    const absorbed: MathList = [];
    let depth = 0;
    let cursor = index + 1;
    let remainder: MathNode | null = null;
    scan: for (; cursor < list.length; cursor += 1) {
      const next = list[cursor];
      if (next.kind !== "text" || next.style.nor) {
        absorbed.push(next);
        continue;
      }
      const characters = Array.from(next.text);
      for (let position = 0; position < characters.length; position += 1) {
        const character = characters[position];
        if (OPENERS.has(character)) depth += 1;
        else if (CLOSERS.has(character)) {
          if (depth === 0) {
            if (position > 0) absorbed.push({ ...next, text: characters.slice(0, position).join("") });
            remainder = { ...next, text: characters.slice(position).join("") };
            break scan;
          }
          depth -= 1;
        } else if (
          depth === 0 &&
          (next.style.lit ||
            (!isOperandCharacter(character) && !PRIME_CHARACTERS.has(character) && !SELF_DELIMITERS.has(character)))
        ) {
          if (position > 0) absorbed.push({ ...next, text: characters.slice(0, position).join("") });
          remainder = { ...next, text: characters.slice(position).join("") };
          break scan;
        }
      }
      absorbed.push(next);
    }
    if (depth !== 0 || absorbed.length === 0 || pairedPositions(absorbed).size !== countBrackets(absorbed)) {
      out.push(node);
      continue;
    }
    out.push(nary ? { ...node, e: absorbed } : { kind: "func", name: [node], e: absorbed });
    if (remainder) out.push(remainder);
    index = remainder ? cursor : cursor - 1;
  }
  return out;
}

/** An upright function name at the end of a run becomes its own run (`dsin` → `d`, `sin`). */
function splitFunctionWords(list: MathList): MathList {
  const out: MathList = [];
  for (const node of list) {
    if (node.kind !== "text" || node.style.nor || node.style.lit || node.style.sty !== "p") {
      out.push(node);
      continue;
    }
    const match = /^(.*?[^A-Za-z])?([A-Za-z]+)$/su.exec(node.text);
    const word = match?.[2];
    if (!match || !word || !match[1] || !RECOGNIZED_FUNCTIONS.has(word)) {
      out.push(node);
      continue;
    }
    out.push({ ...node, text: match[1] }, { ...node, text: word });
  }
  return out;
}

/** Adjacent runs of one style are one run (OMML often stores a run per character). */
function mergeRuns(list: MathList): MathList {
  const out: MathList = [];
  for (const node of list) {
    const last = out.at(-1);
    if (
      node.kind === "text" &&
      last?.kind === "text" &&
      last.style.sty === node.style.sty &&
      last.style.scr === node.style.scr &&
      Boolean(last.style.nor) === Boolean(node.style.nor) &&
      Boolean(last.style.lit) === Boolean(node.style.lit)
    ) {
      out[out.length - 1] = { ...last, text: last.text + node.text };
      continue;
    }
    out.push(node);
  }
  return out;
}

/** Bracket characters in plain runs (all must pair inside an absorbed term). */
function countBrackets(list: MathList): number {
  let count = 0;
  for (const node of list) {
    if (node.kind !== "text" || node.style.nor || node.style.lit) continue;
    for (const character of node.text) {
      if (OPENERS.has(character) || CLOSERS.has(character) || SELF_DELIMITERS.has(character)) count += 1;
    }
  }
  return count;
}

export function serializeUnicodeMath(
  list: MathList,
  mode: SerializeMode = "word",
  options: { absorbNaryands?: boolean } = {},
): string {
  return new Serializer(mode, options).zone(list);
}

/**
 * Digits joined by `.` or `,` read as one number (12.5, 1,000). Keep an
 * object's digits apart from neighbouring text digits: `1/2, 1/2`.
 */
function joinsNumber(before: string, text: string, previous: Piece, piece: Piece): boolean {
  if (previous.first !== "object" && piece.first !== "object" && !previous.consumesSpace) return false;
  return (/[0-9]$/u.test(before) && /^[.,][0-9]/u.test(text)) || (/[0-9][.,]$/u.test(before) && /^[0-9]/u.test(text)) ||
    (/[.,]$/u.test(before) && /^[0-9]/u.test(text) && previous.first === "object");
}

/** Word writes `√(√x) ` — a space after a parenthesized prefix-object argument. */
function prefixArgumentSpace(argument: MathList): string {
  const last = argument.at(-1);
  if (!last || isSimpleText(argument)) return "";
  if (argument.length === 1 && last.kind === "text") return "";
  return last.kind === "rad" || last.kind === "box" || last.kind === "phant" || last.kind === "borderBox" ? " " : "";
}
