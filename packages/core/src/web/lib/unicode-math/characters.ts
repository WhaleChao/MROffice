/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Character tables for Word's linear format (UnicodeMath, Unicode Technical
 * Note #28 v3.2 and the behaviour of Word's own Linearize/Build Up commands).
 *
 * In linear text a letter's math style is carried by its code point exactly as
 * Word writes it: an italic variable is a Mathematical Italic letter (𝑥), an
 * upright one is the plain letter (x), bold is 𝐱, script 𝒳, and so on.
 */

import type { MathRunStyle, MathScript, MathSty } from "./ast";

const UPPER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const LOWER = "abcdefghijklmnopqrstuvwxyz";
const DIGITS = "0123456789";
/** The 58-character layout of every Mathematical Greek block. */
const GREEK =
  "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ∇αβγδεζηθικλμνξοπρςστυφχψω∂ϵϑϰϕϱϖ";

/** Letterlike characters that fill the holes of the alphanumeric blocks. */
const HOLES: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  italic: { h: "ℎ" },
  script: {
    B: "ℬ", E: "ℰ", F: "ℱ", H: "ℋ", I: "ℐ", L: "ℒ", M: "ℳ", R: "ℛ",
    e: "ℯ", g: "ℊ", l: "ℓ", o: "ℴ",
  },
  fraktur: { C: "ℭ", H: "ℌ", I: "ℑ", R: "ℜ", Z: "ℨ" },
  doubleStruck: { C: "ℂ", H: "ℍ", N: "ℕ", P: "ℙ", Q: "ℚ", R: "ℝ", Z: "ℤ" },
};

type Alphabet =
  | "bold"
  | "italic"
  | "boldItalic"
  | "script"
  | "boldScript"
  | "fraktur"
  | "doubleStruck"
  | "boldFraktur"
  | "sans"
  | "sansBold"
  | "sansItalic"
  | "sansBoldItalic"
  | "mono";

const LATIN_BASE: Readonly<Record<Alphabet, number>> = {
  bold: 0x1d400,
  italic: 0x1d434,
  boldItalic: 0x1d468,
  script: 0x1d49c,
  boldScript: 0x1d4d0,
  fraktur: 0x1d504,
  doubleStruck: 0x1d538,
  boldFraktur: 0x1d56c,
  sans: 0x1d5a0,
  sansBold: 0x1d5d4,
  sansItalic: 0x1d608,
  sansBoldItalic: 0x1d63c,
  mono: 0x1d670,
};
const GREEK_BASE: Partial<Record<Alphabet, number>> = {
  bold: 0x1d6a8,
  italic: 0x1d6e2,
  boldItalic: 0x1d71c,
  sansBold: 0x1d756,
  sansBoldItalic: 0x1d790,
};
const DIGIT_BASE: Partial<Record<Alphabet, number>> = {
  bold: 0x1d7ce,
  doubleStruck: 0x1d7d8,
  sans: 0x1d7e2,
  sansBold: 0x1d7ec,
  mono: 0x1d7f6,
};

/** styled code point → [plain character, alphabet] */
const DECODE = new Map<string, readonly [string, Alphabet]>();
/** `${alphabet}:${plain}` → styled code point */
const ENCODE = new Map<string, string>();

function register(alphabet: Alphabet, plain: string, styled: string) {
  ENCODE.set(`${alphabet}:${plain}`, styled);
  if (!DECODE.has(styled)) DECODE.set(styled, [plain, alphabet]);
}

for (const [alphabet, base] of Object.entries(LATIN_BASE) as [Alphabet, number][]) {
  const holes =
    alphabet === "italic"
      ? HOLES.italic
      : alphabet === "script"
        ? HOLES.script
        : alphabet === "fraktur"
          ? HOLES.fraktur
          : alphabet === "doubleStruck"
            ? HOLES.doubleStruck
            : {};
  [...UPPER, ...LOWER].forEach((letter, index) => {
    register(alphabet, letter, holes[letter] ?? String.fromCodePoint(base + index));
  });
}
for (const [alphabet, base] of Object.entries(GREEK_BASE) as [Alphabet, number][]) {
  [...GREEK].forEach((letter, index) => {
    register(alphabet, letter, String.fromCodePoint(base + index));
  });
}
for (const [alphabet, base] of Object.entries(DIGIT_BASE) as [Alphabet, number][]) {
  [...DIGITS].forEach((digit, index) => {
    register(alphabet, digit, String.fromCodePoint(base + index));
  });
}
register("italic", "ı", "𝚤");
register("italic", "ȷ", "𝚥");
register("bold", "Ϝ", "𝟊");
register("bold", "ϝ", "𝟋");
// Greek Theta symbol (ϑ) shares the capital Theta variant slot ϴ; Word writes
// U+03F4 only for ϴ itself. Keep both decodable.

/** Characters that take the math italic style by default (Word, OMML). */
export function isItalicByDefault(character: string): boolean {
  return (
    /^[A-Za-z]$/.test(character) ||
    GREEK.includes(character) ||
    character === "ı" ||
    character === "ȷ"
  );
}

function alphabetFor(style: MathRunStyle, character: string): Alphabet | null {
  const sty: MathSty =
    style.sty ?? (isItalicByDefault(character) ? "i" : "p");
  const bold = sty === "b" || sty === "bi";
  const italic = sty === "i" || sty === "bi";
  switch (style.scr) {
    case "script":
      return bold ? "boldScript" : "script";
    case "fraktur":
      return bold ? "boldFraktur" : "fraktur";
    case "double-struck":
      return "doubleStruck";
    case "sans-serif":
      return bold
        ? italic
          ? "sansBoldItalic"
          : "sansBold"
        : italic
          ? "sansItalic"
          : "sans";
    case "monospace":
      return "mono";
    default:
      return bold ? (italic ? "boldItalic" : "bold") : italic ? "italic" : null;
  }
}

/** The code point Word's linear format uses for one run character. */
export function encodeMathCharacter(character: string, style: MathRunStyle): string {
  const alphabet = alphabetFor(style, character);
  if (!alphabet) return character;
  const styled = ENCODE.get(`${alphabet}:${character}`);
  if (styled) return styled;
  // A bold-italic digit has no code point: Word draws it bold.
  if (alphabet === "boldItalic") return ENCODE.get(`bold:${character}`) ?? character;
  if (alphabet === "sansItalic") return ENCODE.get(`sans:${character}`) ?? character;
  if (alphabet === "sansBoldItalic") return ENCODE.get(`sansBold:${character}`) ?? character;
  return character;
}

const ALPHABET_STYLE: Readonly<Record<Alphabet, { sty?: MathSty; scr?: MathScript }>> = {
  bold: { sty: "b" },
  italic: {},
  boldItalic: { sty: "bi" },
  script: { scr: "script" },
  boldScript: { scr: "script", sty: "b" },
  fraktur: { scr: "fraktur" },
  doubleStruck: { scr: "double-struck" },
  boldFraktur: { scr: "fraktur", sty: "b" },
  sans: { scr: "sans-serif", sty: "p" },
  sansBold: { scr: "sans-serif", sty: "b" },
  sansItalic: { scr: "sans-serif" },
  sansBoldItalic: { scr: "sans-serif", sty: "bi" },
  mono: { scr: "monospace" },
};

/**
 * Decode one linear-format character to its run character and style. Plain
 * letters are upright (Word writes italic variables as Mathematical Italic).
 */
export function decodeMathCharacter(character: string): {
  character: string;
  sty?: MathSty;
  scr?: MathScript;
} {
  const decoded = DECODE.get(character);
  if (decoded) {
    const [plain, alphabet] = decoded;
    const style = ALPHABET_STYLE[alphabet];
    // Bold digits are plain bold; double-struck/mono digits carry no sty.
    return { character: plain, ...style };
  }
  if (isItalicByDefault(character)) return { character, sty: "p" };
  return { character };
}

export function isStyledMathCharacter(character: string): boolean {
  return DECODE.has(character);
}

/**
 * Letterlike symbols that fill holes of the alphanumeric blocks (ℜ, ℝ, ℎ …)
 * are the styled letter itself: Word builds ℝ up as a double-struck R.
 */
export function letterlikeLetter(character: string): { character: string; sty?: MathSty; scr?: MathScript } | null {
  const code = character.codePointAt(0) ?? 0;
  if (code >= 0x1d400) return null;
  const decoded = DECODE.get(character);
  if (!decoded) return null;
  const [plain, alphabet] = decoded;
  return { character: plain, ...ALPHABET_STYLE[alphabet] };
}

/* ------------------------------------------------------------------------ */
/* Linear-format character classes.                                          */
/* ------------------------------------------------------------------------ */

export const FUNCTION_APPLY = "⁡";
export const NARY_AND = "▒"; // ▒
export const OPEN_GROUP = "〖"; // 〖
export const CLOSE_GROUP = "〗"; // 〗
export const OPEN_DELIMITER = "├"; // ├
export const CLOSE_DELIMITER = "┤"; // ┤
export const BELOW = "┬"; // ┬
export const ABOVE = "┴"; // ┴
export const VERTICAL_SEPARATOR = "│"; // │
export const NBSP = " ";

export const OPENERS: ReadonlyMap<string, string> = new Map([
  ["(", ")"],
  ["[", "]"],
  ["{", "}"],
  ["⟨", "⟩"],
  ["〈", "〉"],
  ["⌊", "⌋"],
  ["⌈", "⌉"],
  ["⟦", "⟧"],
  ["⟮", "⟯"],
  ["⌜", "⌝"],
  ["⌞", "⌟"],
  ["⟪", "⟫"],
  ["❲", "❳"],
]);
export const CLOSERS: ReadonlySet<string> = new Set(OPENERS.values());
/** Characters that open and close themselves (matched by count). */
export const SELF_DELIMITERS: ReadonlySet<string> = new Set(["|", "‖"]);
/** Characters Word accepts after ├ / ┤ as an explicit delimiter. */
export const DELIMITER_CHARACTERS: ReadonlySet<string> = new Set([
  ...OPENERS.keys(),
  ...CLOSERS,
  ...SELF_DELIMITERS,
  "/",
  "\\",
  "↑",
  "↓",
  "↕",
  "⇑",
  "⇓",
  "⇕",
]);

export const NARY_OPERATORS: ReadonlySet<string> = new Set([
  ..."∑∏∐∫∬∭⨌∮∯∰∱⨑∲∳⋀⋁⋂⋃⨀⨁⨂⨃⨄⨅⨆",
]);
/** Integrals default to limits beside the sign; the rest above/below. */
export function defaultLimitLocation(chr: string): "undOvr" | "subSup" {
  return /[∫∬∭⨌∮∯∰∱⨑∲∳]/u.test(chr) ? "subSup" : "undOvr";
}

/** Horizontal brackets Word builds up as group characters. */
export const GROUP_CHARACTERS: ReadonlyMap<string, { pos: "top" | "bot"; vertJc: "top" | "bot" }> =
  new Map([
    ["⏞", { pos: "top", vertJc: "bot" }],
    ["⏟", { pos: "bot", vertJc: "top" }],
    ["⏜", { pos: "top", vertJc: "bot" }],
    ["⏝", { pos: "bot", vertJc: "top" }],
  ]);

/** Stretchy arrows that stack with ┴ / ┬ into a group character. */
export const STRETCH_ARROWS: ReadonlySet<string> = new Set([
  ..."←→↔⇐⇒⇔↼⇀↽⇁⟵⟶⟷⟸⟹⟺↩↪↤↦⇋⇌",
]);

/** Prefix operators that build up with one operand. */
export const PREFIX_OPERATORS: ReadonlySet<string> = new Set([
  "√",
  "∛",
  "∜",
  "□",
  "▭",
  "▁",
  "¯",
  "⟡",
  "⬄",
  "⇳",
  "⬍",
  "⬆",
  "⬇",
  "⬌",
  ...GROUP_CHARACTERS.keys(),
]);

/** Phantom operators and their OMML switches. */
export const PHANTOMS: ReadonlyMap<
  string,
  { show: boolean; zeroWid: boolean; zeroAsc: boolean; zeroDesc: boolean }
> = new Map([
  ["⟡", { show: false, zeroWid: false, zeroAsc: false, zeroDesc: false }],
  ["⬄", { show: false, zeroWid: false, zeroAsc: true, zeroDesc: true }],
  ["⇳", { show: false, zeroWid: true, zeroAsc: false, zeroDesc: false }],
  ["⬍", { show: true, zeroWid: false, zeroAsc: true, zeroDesc: true }],
  ["⬆", { show: true, zeroWid: false, zeroAsc: true, zeroDesc: false }],
  ["⬇", { show: true, zeroWid: false, zeroAsc: false, zeroDesc: true }],
  ["⬌", { show: true, zeroWid: true, zeroAsc: false, zeroDesc: false }],
]);

/** Matrix and array introducers. */
export const MATRIX_FENCES: ReadonlyMap<string, readonly [string, string] | null> = new Map([
  ["■", null],
  ["⒨", ["(", ")"]],
  ["ⓢ", ["[", "]"]],
  ["Ⓢ", ["{", "}"]],
  ["⒱", ["|", "|"]],
  ["⒩", ["‖", "‖"]],
]);
export const EQUATION_ARRAY = "█";
export const CASES = "Ⓒ";

export const FRACTION_OPERATORS: ReadonlyMap<string, "bar" | "skw" | "lin" | "noBar"> =
  new Map([
    ["/", "bar"],
    ["⁄", "skw"],
    ["∕", "lin"],
    ["¦", "noBar"],
    ["⒞", "noBar"],
  ]);

export function isCombiningMark(character: string): boolean {
  return /^[̀-ͯ⃐-⃿︠-︯]$/u.test(character);
}

const SUPERSCRIPTS: Readonly<Record<string, string>> = {
  "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5", "⁶": "6",
  "⁷": "7", "⁸": "8", "⁹": "9", "⁺": "+", "⁻": "-", "⁼": "=", "⁽": "(",
  "⁾": ")", "ⁱ": "i", "ⁿ": "n",
};
const SUBSCRIPTS: Readonly<Record<string, string>> = {
  "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4", "₅": "5", "₆": "6",
  "₇": "7", "₈": "8", "₉": "9", "₊": "+", "₋": "-", "₌": "=", "₍": "(",
  "₎": ")",
};
export function unicodeSuperscript(character: string): string | null {
  return SUPERSCRIPTS[character] ?? null;
}
export function unicodeSubscript(character: string): string | null {
  return SUBSCRIPTS[character] ?? null;
}

/**
 * Operand characters: they extend an operand (a script argument, a numerator)
 * instead of terminating it. Everything else is an operator for build-up.
 */
export function isOperandCharacter(character: string): boolean {
  if (/^[\p{L}\p{N}\p{M}]$/u.test(character)) return true;
  return "∞∂∇ℏ∅!…⋯⋮⋱⋰°%$".includes(character) || isStyledMathCharacter(character);
}

/** Word's default Math AutoCorrect "Recognized Functions". */
export const RECOGNIZED_FUNCTIONS: ReadonlySet<string> = new Set([
  "arccos", "arccot", "arccsc", "arcsec", "arcsin", "arctan", "arg", "cos",
  "cosh", "cot", "coth", "csc", "csch", "def", "deg", "det", "dim", "exp",
  "gcd", "hom", "inf", "ker", "lg", "lim", "liminf", "limsup", "ln", "log",
  "max", "min", "Pr", "sec", "sech", "sin", "sinh", "sup", "tan", "tanh",
]);
/** Functions whose `_` limit is written below the name (m:limLow). */
export const LIMIT_FUNCTIONS: ReadonlySet<string> = new Set([
  "lim", "liminf", "limsup", "max", "min", "sup", "inf",
]);

/** Characters that would build up when written literally in a run. */
export const BUILD_UP_CHARACTERS: ReadonlySet<string> = new Set([
  "/", "⁄", "∕", "¦", "⒞", "^", "_", BELOW, ABOVE, NARY_AND, "&", "@",
  VERTICAL_SEPARATOR, "∣", OPEN_DELIMITER, CLOSE_DELIMITER, OPEN_GROUP,
  CLOSE_GROUP, FUNCTION_APPLY, '"', "\\", "#", "ⓜ", ...OPENERS.keys(), ...CLOSERS,
  ...SELF_DELIMITERS, ...PREFIX_OPERATORS, ...MATRIX_FENCES.keys(),
  EQUATION_ARRAY, CASES, ...NARY_OPERATORS, ...Object.keys(SUPERSCRIPTS),
  ...Object.keys(SUBSCRIPTS),
]);
