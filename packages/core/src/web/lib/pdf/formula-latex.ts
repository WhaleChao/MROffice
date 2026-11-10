/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { replaceBehind } from "../regexp-behind";

const GREEK_AND_SYMBOL_LATEX: Record<string, string> = {
  α: "\\alpha",
  β: "\\beta",
  γ: "\\gamma",
  δ: "\\delta",
  ε: "\\epsilon",
  ζ: "\\zeta",
  η: "\\eta",
  θ: "\\theta",
  κ: "\\kappa",
  λ: "\\lambda",
  μ: "\\mu",
  ν: "\\nu",
  ξ: "\\xi",
  π: "\\pi",
  ρ: "\\rho",
  σ: "\\sigma",
  τ: "\\tau",
  φ: "\\phi",
  χ: "\\chi",
  ψ: "\\psi",
  ω: "\\omega",
  Γ: "\\Gamma",
  Δ: "\\Delta",
  Θ: "\\Theta",
  Λ: "\\Lambda",
  Ξ: "\\Xi",
  Π: "\\Pi",
  Σ: "\\Sigma",
  Φ: "\\Phi",
  Ψ: "\\Psi",
  Ω: "\\Omega",
  ℝ: "\\mathbb{R}",
  ℕ: "\\mathbb{N}",
  ℤ: "\\mathbb{Z}",
  ℚ: "\\mathbb{Q}",
  ℂ: "\\mathbb{C}",
  "∅": "\\emptyset",
  "∞": "\\infty",
  "∑": "\\sum",
  "∏": "\\prod",
  "∫": "\\int",
  "∂": "\\partial",
  "∇": "\\nabla",
  "∆": "\\Delta",
  "⋯": "\\cdots",
  "…": "\\cdots",
  "⋮": "\\vdots",
  "⋱": "\\ddots",
  "√": "\\sqrt{}",
  "±": "\\pm",
  "∓": "\\mp",
  "×": "\\times",
  "÷": "\\div",
  "·": "\\cdot",
  "≤": "\\le",
  "≥": "\\ge",
  "≈": "\\approx",
  "≠": "\\ne",
  "≡": "\\equiv",
  "∝": "\\propto",
  "∈": "\\in",
  "∉": "\\notin",
  "∀": "\\forall",
  "∃": "\\exists",
  "∧": "\\land",
  "∨": "\\lor",
  "∩": "\\cap",
  "∪": "\\cup",
  "⊂": "\\subset",
  "⊆": "\\subseteq",
  "⊃": "\\supset",
  "⊇": "\\supseteq",
  "→": "\\to",
  "←": "\\leftarrow",
  "↔": "\\leftrightarrow",
};

const LATEX_COMMAND_TEXT: Record<string, string> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  epsilon: "ε",
  varepsilon: "ε",
  zeta: "ζ",
  eta: "η",
  theta: "θ",
  vartheta: "θ",
  kappa: "κ",
  lambda: "λ",
  mu: "μ",
  nu: "ν",
  xi: "ξ",
  pi: "π",
  rho: "ρ",
  sigma: "σ",
  tau: "τ",
  phi: "φ",
  varphi: "φ",
  chi: "χ",
  psi: "ψ",
  omega: "ω",
  Gamma: "Γ",
  Delta: "Δ",
  Theta: "Θ",
  Lambda: "Λ",
  Xi: "Ξ",
  Pi: "Π",
  Sigma: "Σ",
  Phi: "Φ",
  Psi: "Ψ",
  Omega: "Ω",
  emptyset: "∅",
  infty: "∞",
  sum: "∑",
  prod: "∏",
  int: "∫",
  partial: "∂",
  nabla: "∇",
  cdots: "…",
  vdots: "⋮",
  ddots: "⋱",
  pm: "±",
  mp: "∓",
  times: "×",
  div: "÷",
  cdot: "·",
  le: "≤",
  leq: "≤",
  ge: "≥",
  geq: "≥",
  approx: "≈",
  ne: "≠",
  neq: "≠",
  equiv: "≡",
  propto: "∝",
  in: "∈",
  notin: "∉",
  forall: "∀",
  exists: "∃",
  land: "∧",
  lor: "∨",
  cap: "∩",
  cup: "∪",
  subset: "⊂",
  subseteq: "⊆",
  supset: "⊃",
  supseteq: "⊇",
  to: "→",
  rightarrow: "→",
  leftarrow: "←",
  leftrightarrow: "↔",
  arg: "arg ",
  max: "max",
  min: "min",
  lim: "lim",
  Pr: "Pr",
  det: "det",
  sin: "sin",
  cos: "cos",
  tan: "tan",
  log: "log",
  ln: "ln",
  exp: "exp",
  ldots: "…",
};

const MATHBB_TEXT: Record<string, string> = {
  R: "ℝ",
  N: "ℕ",
  Z: "ℤ",
  Q: "ℚ",
  C: "ℂ",
};

const ASCII_GREEK_LATEX: Record<string, string> = {
  alpha: "\\alpha",
  beta: "\\beta",
  gamma: "\\gamma",
  delta: "\\delta",
  epsilon: "\\epsilon",
  theta: "\\theta",
  lambda: "\\lambda",
  mu: "\\mu",
  pi: "\\pi",
  sigma: "\\sigma",
  omega: "\\omega",
  Gamma: "\\Gamma",
  Delta: "\\Delta",
  Theta: "\\Theta",
  Lambda: "\\Lambda",
  Pi: "\\Pi",
  Sigma: "\\Sigma",
  Omega: "\\Omega",
};

const ASCII_UPPER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const ASCII_LOWER = "abcdefghijklmnopqrstuvwxyz";
const MATH_ITALIC_LOWER = [
  "a",
  "b",
  "c",
  "d",
  "e",
  "f",
  "g",
  "",
  "i",
  "j",
  "k",
  "l",
  "m",
  "n",
  "o",
  "p",
  "q",
  "r",
  "s",
  "t",
  "u",
  "v",
  "w",
  "x",
  "y",
  "z",
] as const;
const MATH_GREEK_UPPER_LATEX = [
  "A",
  "B",
  "\\Gamma",
  "\\Delta",
  "E",
  "Z",
  "H",
  "\\Theta",
  "I",
  "K",
  "\\Lambda",
  "M",
  "N",
  "\\Xi",
  "O",
  "\\Pi",
  "P",
  "\\Theta",
  "\\Sigma",
  "T",
  "Y",
  "\\Phi",
  "X",
  "\\Psi",
  "\\Omega",
  "\\nabla",
] as const;
const MATH_GREEK_LOWER_LATEX = [
  "\\alpha",
  "\\beta",
  "\\gamma",
  "\\delta",
  "\\epsilon",
  "\\zeta",
  "\\eta",
  "\\theta",
  "\\iota",
  "\\kappa",
  "\\lambda",
  "\\mu",
  "\\nu",
  "\\xi",
  "o",
  "\\pi",
  "\\rho",
  "\\varsigma",
  "\\sigma",
  "\\tau",
  "\\upsilon",
  "\\phi",
  "\\chi",
  "\\psi",
  "\\omega",
  "\\partial",
  "\\varepsilon",
  "\\vartheta",
  "\\varkappa",
  "\\varphi",
  "\\varrho",
  "\\varpi",
] as const;

const SUPERSCRIPT_DIGITS: Record<string, string> = {
  "⁰": "0",
  "¹": "1",
  "²": "2",
  "³": "3",
  "⁴": "4",
  "⁵": "5",
  "⁶": "6",
  "⁷": "7",
  "⁸": "8",
  "⁹": "9",
  "⁺": "+",
  "⁻": "-",
  "⁽": "(",
  "⁾": ")",
  ᵀ: "T",
  ᵃ: "a",
  ᵇ: "b",
  ᶜ: "c",
  ᵈ: "d",
  ᵉ: "e",
  ᶠ: "f",
  ᵍ: "g",
  ʰ: "h",
  ⁱ: "i",
  ʲ: "j",
  ᵏ: "k",
  ˡ: "l",
  ᵐ: "m",
  ⁿ: "n",
  ᵒ: "o",
  ᵖ: "p",
  ʳ: "r",
  ˢ: "s",
  ᵗ: "t",
  ᵘ: "u",
  ᵛ: "v",
  ʷ: "w",
  ˣ: "x",
  ʸ: "y",
  ᶻ: "z",
};

const SUBSCRIPT_DIGITS: Record<string, string> = {
  "₀": "0",
  "₁": "1",
  "₂": "2",
  "₃": "3",
  "₄": "4",
  "₅": "5",
  "₆": "6",
  "₇": "7",
  "₈": "8",
  "₉": "9",
  "₊": "+",
  "₋": "-",
  "₍": "(",
  "₎": ")",
  ₐ: "a",
  ₑ: "e",
  ₕ: "h",
  ᵢ: "i",
  ⱼ: "j",
  ₖ: "k",
  ₗ: "l",
  ₘ: "m",
  ₙ: "n",
  ₒ: "o",
  ₚ: "p",
  ᵣ: "r",
  ₛ: "s",
  ₜ: "t",
  ᵤ: "u",
  ᵥ: "v",
  ₓ: "x",
};

const SUPERSCRIPT_RE = /[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ]+/g;
const SUBSCRIPT_RE = /[₀₁₂₃₄₅₆₇₈₉₊₋₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ]+/g;
const SI_UNIT_LATEX_TOKENS = [
  "mol",
  "Hz",
  "Pa",
  "kg",
  "ms",
  "km",
  "cm",
  "mm",
  "Ω",
  "N",
  "J",
  "W",
  "A",
  "V",
  "K",
  "g",
  "m",
  "s",
] as const;
const SINGLE_LETTER_SI_UNITS = new Set(["N", "J", "W", "A", "V", "K", "g", "m", "s"]);

export interface PdfFormulaRun {
  left: number;
  right?: number;
  centerY: number;
  baselineY?: number;
  fontSize: number;
  text: string;
}

function scriptGroupToLatex(
  text: string,
  map: Record<string, string>,
  marker: "^" | "_",
): string {
  const body = Array.from(text)
    .map((ch) => map[ch] ?? ch)
    .join("");
  return `${marker}{${body}}`;
}

function latexTokenNeedsTrailingSpace(token: string): boolean {
  return token.startsWith("\\") && /[A-Za-z]$/.test(token);
}

function appendLatexToken(out: string, token: string): string {
  return latexTokenNeedsTrailingSpace(token)
    ? `${out}${token} `
    : `${out}${token}`;
}

function codePointCharAt(text: string, index: number): string {
  // NOT `/^./u`: dot excludes line terminators, so a "\n" in the text made
  // this return "" — and callers advancing by `ch.length` then advanced by
  // ZERO, an infinite loop on the main thread. Multi-line formula clusters
  // join their boxes with "\n", so any math-font PDF with a multi-line
  
  // codePointAt handles every character including "\n", and is O(1) where
  // the slice+regex was O(n) from an always-linear caller loop.
  const cp = text.codePointAt(index);
  return cp == null ? "" : String.fromCodePoint(cp);
}

function previousCodePointChar(text: string, index: number): string {
  // O(1) equivalent of `Array.from(text.slice(0, index)).at(-1)` — that form
  // re-materialized the whole prefix per call and was a measured hot spot.
  if (index <= 0) return "";
  const trailing = text.charCodeAt(index - 1);
  if (trailing >= 0xdc00 && trailing <= 0xdfff && index >= 2) {
    const leading = text.charCodeAt(index - 2);
    if (leading >= 0xd800 && leading <= 0xdbff) {
      return text.slice(index - 2, index);
    }
  }
  return text[index - 1] ?? "";
}

function isAsciiUnitNeighbor(ch: string): boolean {
  return /[A-Za-z_]/.test(ch);
}

function isAsciiAlnum(ch: string): boolean {
  return /[A-Za-z0-9]/.test(ch);
}

function hasNumericUnitContextBefore(text: string, start: number): boolean {
  const before = text.slice(0, start).trimEnd();
  if (/\d$/.test(before)) return true;
  const scriptBase = trailingNumericLatexScriptBase(before);
  if (scriptBase && /\d$/.test(scriptBase)) return true;
  return trailingNumericDelimitedContext(before);
}

function unitLatexToken(token: string): string {
  return token === "Ω" ? "\\Omega" : `\\mathrm{${token}}`;
}

function trailingNumericLatexScriptBase(text: string): string | null {
  const match = /\^\{\s*[+\-−]?\d+(?:\.\d+)?\s*\}$/.exec(text);
  if (!match) return null;
  return text.slice(0, match.index).trimEnd();
}

function trailingNumericDelimitedContext(text: string): boolean {
  const close = text.at(-1);
  const open = close === ")" ? "(" : close === "]" ? "[" : "";
  if (!open || !close) return false;
  let depth = 0;
  for (let index = text.length - 1; index >= 0; index -= 1) {
    const ch = text[index];
    if (ch === close) {
      depth += 1;
    } else if (ch === open) {
      depth -= 1;
      if (depth === 0) {
        const body = text.slice(index + 1, text.length - 1).trim();
        return /^[+\-−]?\d+(?:\.\d+)?(?:\s*[,;]\s*[+\-−]?\d+(?:\.\d+)?)*$/.test(
          body,
        );
      }
    }
  }
  return false;
}

function readUnitScript(
  text: string,
  start: number,
): { latex: string; end: number } | null {
  const ch = codePointCharAt(text, start);
  if (SUPERSCRIPT_DIGITS[ch]) {
    let cursor = start;
    let script = "";
    while (cursor < text.length) {
      const next = codePointCharAt(text, cursor);
      if (!SUPERSCRIPT_DIGITS[next]) break;
      script += next;
      cursor += next.length;
    }
    return {
      latex: scriptGroupToLatex(script, SUPERSCRIPT_DIGITS, "^"),
      end: cursor,
    };
  }

  if (text[start] !== "^") return null;
  const cursor = skipLatexSpaces(text, start + 1);
  const group = readLatexGroup(text, cursor);
  if (group) {
    return { latex: `^{${group.body.trim()}}`, end: group.end };
  }
  const next = codePointCharAt(text, cursor);
  if (!next || !/[A-Za-z0-9+\-−]/u.test(next)) return null;
  return { latex: `^{${next === "−" ? "-" : next}}`, end: cursor + next.length };
}

function readSiUnitToken(
  text: string,
  start: number,
): { latex: string; end: number; singleLetter: boolean } | null {
  for (const token of SI_UNIT_LATEX_TOKENS) {
    if (!text.startsWith(token, start)) continue;
    const end = start + token.length;
    const next = codePointCharAt(text, end);
    if (isAsciiAlnum(next) || next === "_") continue;

    const script = readUnitScript(text, end);
    return {
      latex: `${unitLatexToken(token)}${script?.latex ?? ""}`,
      end: script?.end ?? end,
      singleLetter: SINGLE_LETTER_SI_UNITS.has(token),
    };
  }
  return null;
}

function readSiUnitExpression(
  text: string,
  start: number,
): { latex: string; end: number } | null {
  const previous = previousCodePointChar(text, start);
  if (previous === "\\" || isAsciiUnitNeighbor(previous)) return null;

  const first = readSiUnitToken(text, start);
  if (!first) return null;

  const parts = [first.latex];
  let cursor = skipLatexSpaces(text, first.end);
  let end = first.end;
  let hasSlash = false;
  while (text[cursor] === "/" || text[cursor] === "／") {
    const slashIndex = cursor;
    cursor = skipLatexSpaces(text, cursor + 1);
    const next = readSiUnitToken(text, cursor);
    if (!next) {
      cursor = slashIndex;
      break;
    }
    hasSlash = true;
    parts.push("/", next.latex);
    end = next.end;
    cursor = skipLatexSpaces(text, next.end);
  }

  const next = codePointCharAt(text, end);
  if (isAsciiAlnum(next) || next === "_") return null;
  if (first.singleLetter && !hasSlash && !hasNumericUnitContextBefore(text, start)) {
    return null;
  }

  return { latex: parts.join(""), end };
}

function mathLatinTokenFromCodePoint(cp: number): string | null {
  if (cp >= 0x1d400 && cp <= 0x1d419) {
    return ASCII_UPPER[cp - 0x1d400] ?? null;
  }
  if (cp >= 0x1d41a && cp <= 0x1d433) {
    return ASCII_LOWER[cp - 0x1d41a] ?? null;
  }
  if (cp >= 0x1d434 && cp <= 0x1d44d) {
    return ASCII_UPPER[cp - 0x1d434] ?? null;
  }
  if (cp >= 0x1d44e && cp <= 0x1d467) {
    return MATH_ITALIC_LOWER[cp - 0x1d44e] || null;
  }
  if (cp >= 0x1d468 && cp <= 0x1d481) {
    return ASCII_UPPER[cp - 0x1d468] ?? null;
  }
  if (cp >= 0x1d482 && cp <= 0x1d49b) {
    return ASCII_LOWER[cp - 0x1d482] ?? null;
  }
  if (cp >= 0x1d5a0 && cp <= 0x1d5b9) {
    return ASCII_UPPER[cp - 0x1d5a0] ?? null;
  }
  if (cp >= 0x1d5ba && cp <= 0x1d5d3) {
    return ASCII_LOWER[cp - 0x1d5ba] ?? null;
  }
  if (cp >= 0x1d5d4 && cp <= 0x1d5ed) {
    return ASCII_UPPER[cp - 0x1d5d4] ?? null;
  }
  if (cp >= 0x1d5ee && cp <= 0x1d607) {
    return ASCII_LOWER[cp - 0x1d5ee] ?? null;
  }
  if (cp >= 0x1d608 && cp <= 0x1d621) {
    return ASCII_UPPER[cp - 0x1d608] ?? null;
  }
  if (cp >= 0x1d622 && cp <= 0x1d63b) {
    return ASCII_LOWER[cp - 0x1d622] ?? null;
  }
  if (cp >= 0x1d63c && cp <= 0x1d655) {
    return ASCII_UPPER[cp - 0x1d63c] ?? null;
  }
  if (cp >= 0x1d656 && cp <= 0x1d66f) {
    return ASCII_LOWER[cp - 0x1d656] ?? null;
  }
  if (cp >= 0x1d670 && cp <= 0x1d689) {
    return ASCII_UPPER[cp - 0x1d670] ?? null;
  }
  if (cp >= 0x1d68a && cp <= 0x1d6a3) {
    return ASCII_LOWER[cp - 0x1d68a] ?? null;
  }
  return null;
}

function mathGreekTokenFromCodePoint(cp: number): string | null {
  for (const start of [0x1d6a8, 0x1d6e2, 0x1d71c, 0x1d756, 0x1d790]) {
    if (cp >= start && cp < start + MATH_GREEK_UPPER_LATEX.length) {
      return MATH_GREEK_UPPER_LATEX[cp - start] ?? null;
    }
  }
  for (const start of [0x1d6c2, 0x1d6fc, 0x1d736, 0x1d770, 0x1d7aa]) {
    if (cp >= start && cp < start + MATH_GREEK_LOWER_LATEX.length) {
      return MATH_GREEK_LOWER_LATEX[cp - start] ?? null;
    }
  }
  return null;
}

function mathAlphanumericLatexToken(ch: string): string | null {
  if (ch === "ℎ") return "h";
  const cp = ch.codePointAt(0);
  if (cp == null) return null;
  return mathLatinTokenFromCodePoint(cp) ?? mathGreekTokenFromCodePoint(cp);
}

function normalizeLatexSpacing(text: string): string {
  return text
    .replace(/[−–—]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s+([_^])/g, "$1")
    .replace(/\s+([,.;:!)\]}])/g, "$1")
    .replace(/([(\[{])\s+/g, "$1")
    .replace(/([A-Za-z0-9}\\\]])\s+([(\[])/g, "$1$2")
    .replace(/\\([A-Za-z]+)\s+(?=[A-Za-z])/g, "\\$1 ")
    .trim();
}

function normalizeFormulaFallbackSpacing(text: string): string {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/[ \t\f\v]+/g, " ")
    .replace(/[ \t]*\n[ \t]*/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/\s+([,.;:!)\]}])/g, "$1")
    .replace(/([(\[{])\s+/g, "$1")
    .replace(/[ \t\f\v]*([_^])[ \t\f\v]*/g, "$1")
    .replace(/[ \t\f\v]*\/[ \t\f\v]*/g, "/")
    .replace(/[ \t\f\v]*([=+×÷<>≤≥≈≠])[ \t\f\v]*/g, " $1 ")
    .replace(/[ \t\f\v]+/g, " ")
    .replace(/ *\n */g, "\n")
    .trim();
}

function compactGroupedScriptFallbackSpacing(text: string): string {
  return text.replace(
    /([_^])\(([^()\n]*)\)/g,
    (_, marker: string, body: string) => {
      const compact = body.replace(/\s*([=+×÷<>≤≥≈≠→←↔])\s*/g, "$1");
      return `${marker}(${compact})`;
    },
  );
}

function joinLatexTextEditSegments(
  ...segments: Array<string | undefined>
): string {
  return normalizeLatexSpacing(
    segments
      .map((segment) => segment?.trim())
      .filter((segment): segment is string => Boolean(segment))
      .join(" "),
  );
}

function removeLatexTextEditSuffix(
  latex: string,
  suffix: string,
): string | null {
  const base = latex.trimEnd();
  const trimmedSuffix = suffix.trim();
  if (!trimmedSuffix || !base.endsWith(trimmedSuffix)) return null;
  return normalizeLatexSpacing(
    base.slice(0, base.length - trimmedSuffix.length).trimEnd(),
  );
}

function uniqueLatexTextReplace(
  latex: string,
  search: string,
  replacement: string,
): string | null {
  if (!search) return null;
  const index = latex.indexOf(search);
  if (index < 0) return null;
  if (latex.indexOf(search, index + search.length) >= 0) return null;
  return normalizeLatexSpacing(
    `${latex.slice(0, index)}${replacement}${latex.slice(
      index + search.length,
    )}`,
  );
}

function textEditDelta(
  previousText: string,
  nextText: string,
): { removed: string; inserted: string } {
  let prefix = 0;
  while (
    prefix < previousText.length &&
    prefix < nextText.length &&
    previousText[prefix] === nextText[prefix]
  ) {
    prefix += 1;
  }

  let previousEnd = previousText.length;
  let nextEnd = nextText.length;
  while (
    previousEnd > prefix &&
    nextEnd > prefix &&
    previousText[previousEnd - 1] === nextText[nextEnd - 1]
  ) {
    previousEnd -= 1;
    nextEnd -= 1;
  }

  return {
    removed: previousText.slice(prefix, previousEnd),
    inserted: nextText.slice(prefix, nextEnd),
  };
}

function normalizedFormulaEditText(text: string): string {
  return normalizeFormulaFallbackSpacing(text).toLowerCase();
}

function latexEditTokens(text: string): Set<string> {
  const normalized = text
    .replace(/\\begin\{[^}]+\}|\\end\{[^}]+\}/g, " ")
    .replace(/\\operatorname\{([^}]+)\}/g, " $1 ")
    .replace(/\\mathrm\{([^}]+)\}/g, " $1 ")
    .replace(/\\text\{([^}]+)\}/g, " $1 ")
    .replace(/\\mathbb\{([^}]+)\}/g, " $1 ")
    .replace(/\\([A-Za-z]+)/g, " $1 ")
    .toLowerCase();
  return new Set(normalized.match(/[a-zα-ωΑ-Ω]+|\d+(?:\.\d+)?/g) ?? []);
}

function latexLooksRelatedToPreviousText(
  latex: string,
  previousText: string,
): boolean {
  const generatedPrevious = latexFromPdfFormulaText(previousText);
  const fallbackPrevious = fallbackTextFromPdfFormulaLatex(latex);
  if (latex === generatedPrevious) return true;
  if (
    normalizedFormulaEditText(fallbackPrevious) ===
    normalizedFormulaEditText(previousText)
  ) {
    return true;
  }

  const sourceTokens = latexEditTokens(latex);
  const generatedTokens = latexEditTokens(generatedPrevious);
  if (sourceTokens.size === 0 || generatedTokens.size === 0) return false;
  let overlap = 0;
  for (const token of generatedTokens) {
    if (sourceTokens.has(token)) overlap += 1;
  }
  return overlap >= Math.min(2, generatedTokens.size);
}

/**
 * Preserve semantic LaTeX while a user or AI edits the visible fallback text.
 *
 * PDF extraction often knows more than the plain text can express, e.g.
 * `\frac{1}{2}` renders as `1/2`. When the visible text changes from `1/2` to
 * `1/3`, blindly regenerating LaTeX would lose the structured fraction. This
 * helper patches the existing source only when it can prove the source matches
 * the previous visible fallback and the changed fragment is unique.
 */
export function preservePdfFormulaLatexForTextEdit({
  currentLatex,
  previousText,
  nextText,
}: {
  currentLatex?: string | null;
  previousText: string;
  nextText: string;
}): string | null {
  const latex = typeof currentLatex === "string" ? currentLatex.trim() : "";
  if (!latex) return null;
  if (nextText === previousText) return latex;

  if (!latexLooksRelatedToPreviousText(latex, previousText)) return null;

  if (nextText.startsWith(previousText)) {
    const inserted = nextText.slice(previousText.length);
    return joinLatexTextEditSegments(latex, latexFromPdfFormulaText(inserted));
  }

  if (nextText.endsWith(previousText)) {
    const inserted = nextText.slice(0, nextText.length - previousText.length);
    return joinLatexTextEditSegments(latexFromPdfFormulaText(inserted), latex);
  }

  if (previousText.startsWith(nextText)) {
    const deleted = previousText.slice(nextText.length);
    return removeLatexTextEditSuffix(latex, latexFromPdfFormulaText(deleted));
  }

  const delta = textEditDelta(previousText, nextText);
  if (!delta.removed) return null;

  const insertedLatex = delta.inserted
    ? latexFromPdfFormulaText(delta.inserted)
    : "";
  const removedLatex = latexFromPdfFormulaText(delta.removed);
  const candidates = [removedLatex, delta.removed]
    .map((candidate) => candidate.trim())
    .filter(
      (candidate, index, all) => candidate && all.indexOf(candidate) === index,
    );

  for (const candidate of candidates) {
    const patched = uniqueLatexTextReplace(latex, candidate, insertedLatex);
    if (patched != null) return patched;
  }

  return null;
}

export function latexForPdfFormulaTextEdit({
  currentLatex,
  previousText,
  nextText,
}: {
  currentLatex?: string | null;
  previousText: string;
  nextText: string;
}): string {
  return (
    preservePdfFormulaLatexForTextEdit({
      currentLatex,
      previousText,
      nextText,
    }) ?? latexFromPdfFormulaText(nextText)
  );
}

function skipLatexSpaces(text: string, start: number): number {
  let index = start;
  while (index < text.length && /\s/.test(text[index] ?? "")) index += 1;
  return index;
}

function readLatexGroup(
  text: string,
  start: number,
): { body: string; end: number } | null {
  if (text[start] !== "{") return null;
  let depth = 0;
  for (let index = start; index < text.length; index += 1) {
    const char = text[index];
    if (char === "\\" && index + 1 < text.length) {
      index += 1;
      continue;
    }
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) {
        return { body: text.slice(start + 1, index), end: index + 1 };
      }
    }
  }
  return null;
}

function readDelimitedText(
  text: string,
  start: number,
  open: string,
  close: string,
): { body: string; end: number } | null {
  if (text[start] !== open) return null;
  let depth = 0;
  for (let index = start; index < text.length; index += 1) {
    const char = text[index];
    if (char === "\\" && index + 1 < text.length) {
      index += 1;
      continue;
    }
    if (char === open) depth += 1;
    if (char === close) {
      depth -= 1;
      if (depth === 0) {
        return { body: text.slice(start + 1, index), end: index + 1 };
      }
    }
  }
  return null;
}

function readLatexBracket(
  text: string,
  start: number,
): { body: string; end: number } | null {
  return readDelimitedText(text, start, "[", "]");
}

function splitTopLevelComma(text: string): [string, string] | null {
  let braceDepth = 0;
  let parenDepth = 0;
  let bracketDepth = 0;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index] ?? "";
    if (char === "\\" && index + 1 < text.length) {
      index += 1;
      continue;
    }
    if (char === "{") braceDepth += 1;
    else if (char === "}") braceDepth = Math.max(0, braceDepth - 1);
    else if (char === "(") parenDepth += 1;
    else if (char === ")") parenDepth = Math.max(0, parenDepth - 1);
    else if (char === "[") bracketDepth += 1;
    else if (char === "]") bracketDepth = Math.max(0, bracketDepth - 1);
    else if (
      char === "," &&
      braceDepth === 0 &&
      parenDepth === 0 &&
      bracketDepth === 0
    ) {
      return [text.slice(0, index), text.slice(index + 1)];
    }
  }
  return null;
}

function friendlyFunctionAt(
  text: string,
  start: number,
  name: string,
): { body: string; end: number } | null {
  if (!text.startsWith(name, start)) return null;
  const before = text[start - 1] ?? "";
  const after = text[start + name.length] ?? "";
  if (before === "\\" || /[A-Za-z0-9]/.test(before)) return null;
  if (/[A-Za-z0-9]/.test(after)) return null;
  const groupStart = skipLatexSpaces(text, start + name.length);
  return readDelimitedText(text, groupStart, "(", ")");
}

function tryFriendlyFormulaFunction(
  text: string,
  start: number,
): { latex: string; end: number } | null {
  const sqrtCall = friendlyFunctionAt(text, start, "sqrt");
  if (sqrtCall) {
    return {
      latex: `\\sqrt{${normalizePdfLatexSourceSegment(sqrtCall.body, false)}}`,
      end: sqrtCall.end,
    };
  }

  const rootCall = friendlyFunctionAt(text, start, "root");
  if (rootCall) {
    const parts = splitTopLevelComma(rootCall.body);
    if (!parts) return null;
    return {
      latex: `\\sqrt[${normalizePdfLatexSourceSegment(
        parts[0].trim(),
        false,
      )}]{${normalizePdfLatexSourceSegment(parts[1].trim(), false)}}`,
      end: rootCall.end,
    };
  }

  for (const command of ["frac", "dfrac", "tfrac"] as const) {
    const call = friendlyFunctionAt(text, start, command);
    if (!call) continue;
    const parts = splitTopLevelComma(call.body);
    if (!parts) return null;
    return {
      latex: `\\${command}{${normalizePdfLatexSourceSegment(
        parts[0].trim(),
        false,
      )}}{${normalizePdfLatexSourceSegment(parts[1].trim(), false)}}`,
      end: call.end,
    };
  }

  return null;
}

function partLooksLikeUnitOrText(part: string): boolean {
  return (
    /\\(?:text|mathrm)\s*\{/.test(part) ||
    /[\uE000\uE001]/u.test(part) ||
    /[\u3400-\u9fff]/u.test(part)
  );
}

function safeFractionPart(part: string): string | null {
  const normalized = normalizeLatexSpacing(part.trim());
  if (!normalized || partLooksLikeUnitOrText(normalized)) return null;
  return normalized;
}

function safeSlashFraction(
  numerator: string,
  denominator: string,
): string | null {
  const top = safeFractionPart(numerator);
  const bottom = safeFractionPart(denominator);
  if (!top || !bottom) return null;
  const signedTop = /^([+\-−])\s*(.+)$/.exec(top);
  if (signedTop) {
    const sign = signedTop[1] === "+" ? "" : "-";
    return `${sign}\\frac{${signedTop[2]}}{${bottom}}`;
  }
  return `\\frac{${top}}{${bottom}}`;
}

function protectTextModeLatexGroups(
  text: string,
  transform: (input: string) => string,
): string {
  const placeholders: string[] = [];
  let protectedText = "";
  let index = 0;

  while (index < text.length) {
    if (text[index] !== "\\") {
      protectedText += text[index] ?? "";
      index += 1;
      continue;
    }

    const command = readLatexCommandToken(text, index);
    if (!command || !TEXT_MODE_LATEX_COMMANDS.has(command.name)) {
      protectedText += text[index] ?? "";
      index += 1;
      continue;
    }

    const groupStart = skipLatexSpaces(text, command.end);
    const group = readLatexGroup(text, groupStart);
    if (!group) {
      protectedText += command.token;
      index = command.end;
      continue;
    }

    const placeholder = `\uE000${placeholders.length}\uE001`;
    placeholders.push(text.slice(index, group.end));
    protectedText += placeholder;
    index = group.end;
  }

  const transformed = transform(protectedText);
  return transformed.replace(
    /\uE000(\d+)\uE001/g,
    (_match, rawIndex: string) => {
      const replacement = placeholders[Number(rawIndex)];
      return replacement ?? "";
    },
  );
}

function normalizeSafeSlashFractions(text: string): string {
  return protectTextModeLatexGroups(text, (input) => {
    const grouped = input
      .replace(
        /\(([^()\n]+)\)\s*\/\s*\(([^()\n]+)\)/g,
        (match: string, numerator: string, denominator: string) =>
          safeSlashFraction(numerator, denominator) ?? match,
      )
      .replace(
        /([+\-−]?\d+(?:\.\d+)?|[A-Za-z](?:\^\{[^{}]+\}|_\{[^{}]+\})?)\s*\/\s*\(([^()\n]+)\)/g,
        (match: string, numerator: string, denominator: string) =>
          safeSlashFraction(numerator, denominator) ?? match,
      )
      .replace(
        /([+\-−]?\d+(?:\.\d+)?)\s*\/\s*([A-Za-z])\(([^()\n]+)\)/g,
        (match: string, numerator: string, fn: string, args: string) =>
          safeSlashFraction(numerator, `${fn}(${args})`) ?? match,
      );
    // A plain number fraction not after a letter or a backslash, tested in
    // code: Safari before 16.4 cannot compile the lookbehind.
    return replaceBehind(
      grouped,
      /([+\-−]?\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)(?![A-Za-z])/g,
      { char: /[A-Za-z\\]/, not: true },
      ([match, numerator, denominator]) => safeSlashFraction(numerator, denominator) ?? match,
    );
  });
}

function replaceTwoGroupLatexCommand(
  text: string,
  command: string,
  render: (first: string, second: string) => string,
): string {
  const token = `\\${command}`;
  let out = "";
  let index = 0;
  while (index < text.length) {
    if (text.startsWith(token, index)) {
      let cursor = skipLatexSpaces(text, index + token.length);
      const first = readLatexGroup(text, cursor);
      cursor = first ? skipLatexSpaces(text, first.end) : cursor;
      const second = first ? readLatexGroup(text, cursor) : null;
      if (first && second) {
        out += render(
          fallbackTextFromPdfFormulaLatex(first.body),
          fallbackTextFromPdfFormulaLatex(second.body),
        );
        index = second.end;
        continue;
      }
    }
    out += text[index] ?? "";
    index += 1;
  }
  return out;
}

function replaceOneGroupLatexCommand(
  text: string,
  command: string,
  render: (body: string) => string,
): string {
  const token = `\\${command}`;
  let out = "";
  let index = 0;
  while (index < text.length) {
    if (text.startsWith(token, index)) {
      const group = readLatexGroup(
        text,
        skipLatexSpaces(text, index + token.length),
      );
      if (group) {
        out += render(fallbackTextFromPdfFormulaLatex(group.body));
        index = group.end;
        continue;
      }
    }
    out += text[index] ?? "";
    index += 1;
  }
  return out;
}

function replaceSqrtLatexCommand(text: string): string {
  const token = "\\sqrt";
  let out = "";
  let index = 0;
  while (index < text.length) {
    if (text.startsWith(token, index)) {
      let cursor = skipLatexSpaces(text, index + token.length);
      const rootIndex = readLatexBracket(text, cursor);
      cursor = rootIndex ? skipLatexSpaces(text, rootIndex.end) : cursor;
      const group = readLatexGroup(text, cursor);
      if (group) {
        const body = fallbackTextFromPdfFormulaLatex(group.body);
        if (rootIndex) {
          out += `√[${fallbackTextFromPdfFormulaLatex(rootIndex.body)}](${body})`;
        } else {
          out += `√(${body})`;
        }
        index = group.end;
        continue;
      }
    }
    out += text[index] ?? "";
    index += 1;
  }
  return out;
}

function applyCombiningAccent(body: string, combiningMark: string): string {
  return body ? `${body}${combiningMark}` : "";
}

function splitTopLevelLatexColumns(text: string): string[] {
  const columns: string[] = [];
  let start = 0;
  let depth = 0;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === "\\" && index + 1 < text.length) {
      index += 1;
      continue;
    }
    if (char === "{") {
      depth += 1;
      continue;
    }
    if (char === "}") {
      depth = Math.max(0, depth - 1);
      continue;
    }
    if (char === "&" && depth === 0) {
      columns.push(text.slice(start, index).trim());
      start = index + 1;
    }
  }
  columns.push(text.slice(start).trim());
  return columns;
}

function skipLatexRowOption(text: string, start: number): number {
  let index = skipLatexSpaces(text, start);
  if (text[index] !== "[") return index;
  let depth = 0;
  for (; index < text.length; index += 1) {
    const char = text[index];
    if (char === "\\" && index + 1 < text.length) {
      index += 1;
      continue;
    }
    if (char === "[") depth += 1;
    if (char === "]") {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
  }
  return start;
}

function splitTopLevelLatexRows(text: string): string[] {
  const rows: string[] = [];
  let start = 0;
  let depth = 0;
  let index = 0;
  while (index < text.length) {
    const char = text[index];
    if (char === "\\" && text[index + 1] === "\\" && depth === 0) {
      rows.push(text.slice(start, index).trim());
      index = skipLatexRowOption(text, index + 2);
      start = index;
      continue;
    }
    if (char === "\\" && index + 1 < text.length) {
      index += 2;
      continue;
    }
    if (char === "{") depth += 1;
    if (char === "}") depth = Math.max(0, depth - 1);
    index += 1;
  }
  rows.push(text.slice(start).trim());
  return rows.filter(Boolean);
}

function latexRowLooksLikeCase(row: string): boolean {
  const columns = splitTopLevelLatexColumns(row);
  if (columns.length < 2) return false;
  const expression = fallbackTextFromPdfFormulaLatex(columns[0]);
  const condition = fallbackTextFromPdfFormulaLatex(columns.slice(1).join(" "));
  if (!expression || !condition) return false;
  return /[=<>≤≥≈≠]|\b(?:otherwise|if|when|for)\b/i.test(condition);
}

function latexRowsLookLikeCases(rows: readonly string[]): boolean {
  return rows.length > 0 && rows.every(latexRowLooksLikeCase);
}

function renderLatexCasesFallback(rows: readonly string[]): string {
  return rows
    .map((row) => {
      const columns = splitTopLevelLatexColumns(row);
      if (columns.length < 2) return fallbackTextFromPdfFormulaLatex(row);
      const expression = fallbackTextFromPdfFormulaLatex(columns[0]);
      const condition = fallbackTextFromPdfFormulaLatex(
        columns.slice(1).join(" "),
      );
      return `${expression}, ${condition}`.trim();
    })
    .filter(Boolean)
    .join("\n");
}

function renderLatexMatrixFallback(
  env: string,
  rows: readonly string[],
): string {
  const matrixRows = rows.map((row) =>
    splitTopLevelLatexColumns(row)
      .map((cell) => fallbackTextFromPdfFormulaLatex(cell).replace(/\n+/g, " "))
      .filter(Boolean)
      .join(" "),
  );
  const matrixBody = matrixRows.filter(Boolean).join("\n");
  if (env === "pmatrix") return `(${matrixBody})`;
  if (env === "vmatrix") return `|${matrixBody}|`;
  if (env === "Vmatrix") return `‖${matrixBody}‖`;
  if (env === "matrix" || env === "smallmatrix" || env === "array") {
    return matrixBody;
  }
  return `[${matrixBody}]`;
}

function renderLatexEnvironmentFallback(env: string, body: string): string {
  const rows = splitTopLevelLatexRows(body);
  if (env === "cases") {
    return renderLatexCasesFallback(rows);
  }
  return renderLatexMatrixFallback(env, rows);
}

function stripArrayAlignmentGroup(body: string): string {
  const start = skipLatexSpaces(body, 0);
  const alignmentGroup = readLatexGroup(body, start);
  if (!alignmentGroup) return body;
  const alignment = alignmentGroup.body.trim();
  if (!/^[clr|@!<>{}*\s.]+$/.test(alignment)) return body;
  return body.slice(alignmentGroup.end);
}

function stripTrailingCasesDelimiter(text: string): {
  text: string;
  stripped: boolean;
} {
  const stripped = text.replace(/(?:\\\{|[{])\s*$/, "");
  return { text: stripped, stripped: stripped !== text };
}

function replaceStructuredLatexEnvironments(text: string): string {
  const envPattern =
    /\\begin\{(cases|array|bmatrix|Bmatrix|pmatrix|matrix|smallmatrix|vmatrix|Vmatrix)\}/g;
  let out = "";
  let cursor = 0;
  let match: RegExpExecArray | null;
  while ((match = envPattern.exec(text)) != null) {
    const env = match[1];
    let bodyStart = envPattern.lastIndex;
    if (env === "array") {
      const alignmentGroup = readLatexGroup(
        text,
        skipLatexSpaces(text, bodyStart),
      );
      if (alignmentGroup) {
        bodyStart = alignmentGroup.end;
      }
    }
    const endToken = `\\end{${env}}`;
    const end = text.indexOf(endToken, bodyStart);
    if (end < 0) {
      out += text.slice(cursor, bodyStart);
      cursor = bodyStart;
      continue;
    }
    const body =
      env === "array"
        ? stripArrayAlignmentGroup(text.slice(bodyStart, end))
        : text.slice(bodyStart, end);
    const rows = splitTopLevelLatexRows(body);
    const isArrayCases = env === "array" && latexRowsLookLikeCases(rows);
    let prefix = text.slice(cursor, match.index);
    let strippedCasesDelimiter = false;
    if (isArrayCases) {
      const stripped = stripTrailingCasesDelimiter(prefix);
      prefix = stripped.text;
      strippedCasesDelimiter = stripped.stripped;
    }

    out += prefix;
    out += isArrayCases
      ? renderLatexCasesFallback(rows)
      : renderLatexEnvironmentFallback(env, body);
    cursor = end + endToken.length;
    if (strippedCasesDelimiter && text[cursor] === ".") {
      cursor += 1;
    }
    envPattern.lastIndex = cursor;
  }
  return out + text.slice(cursor);
}

function needsFractionGrouping(text: string): boolean {
  return /[\s=+\-−×÷<>≤≥≈≠]/.test(text);
}

function groupedFractionPart(text: string): string {
  const trimmed = text.trim();
  if (!trimmed) return "";
  return needsFractionGrouping(trimmed) ? `(${trimmed})` : trimmed;
}

function replaceLatexScripts(text: string): string {
  let out = "";
  let index = 0;
  while (index < text.length) {
    const marker = text[index];
    if (marker === "^" || marker === "_") {
      const cursor = skipLatexSpaces(text, index + 1);
      const group = readLatexGroup(text, cursor);
      if (group) {
        out += `${marker}${readableScriptFallbackBody(
          fallbackTextFromPdfFormulaLatex(group.body),
        )}`;
        index = group.end;
        continue;
      }
      const next = text[cursor] ?? "";
      if (next) {
        out += `${marker}${readableScriptFallbackBody(
          fallbackTextFromPdfFormulaLatex(next),
        )}`;
        index = cursor + 1;
        continue;
      }
    }
    out += marker ?? "";
    index += 1;
  }
  return out;
}

function readableScriptFallbackBody(text: string): string {
  const normalized = normalizeFormulaFallbackSpacing(text).replace(
    /\s*([=+×÷<>≤≥≈≠→←↔])\s*/g,
    "$1",
  );
  if (!normalized) return "";
  const nonLeadingMinus = normalized.slice(1).match(/[-−]/);
  if (/[\s=+×÷<>≤≥≈≠→←↔,]/.test(normalized) || nonLeadingMinus) {
    return `(${normalized})`;
  }
  return normalized;
}

function joinLatexFormulaParts(parts: readonly string[]): string {
  return normalizeLatexSpacing(parts.filter(Boolean).join(" "));
}

function normalizeMatrixBracketGlyphs(text: string): string {
  return text
    .replace(/[⎡⎢⎣]/g, "[")
    .replace(/[⎤⎥⎦]/g, "]")
    .replace(/[⎛⎜⎝]/g, "(")
    .replace(/[⎞⎟⎠]/g, ")");
}

function isRootAtomChar(char: string): boolean {
  return /[A-Za-z0-9αβγδθλμνπρσφψωΑ-Ω₀₁₂₃₄₅₆₇₈₉₊₋₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ\u{1d400}-\u{1d7ff}]/u.test(
    char,
  );
}

function readVisualRootTarget(
  text: string,
  start: number,
): { latex: string; end: number } | null {
  let cursor = skipLatexSpaces(text, start);
  let rootIndex = "";
  const indexGroup = readDelimitedText(text, cursor, "[", "]");
  if (indexGroup) {
    rootIndex = `[${indexGroup.body.trim()}]`;
    cursor = skipLatexSpaces(text, indexGroup.end);
  }

  const grouped = readDelimitedText(text, cursor, "(", ")");
  if (grouped) {
    return {
      latex: `\\sqrt${rootIndex}{${grouped.body.trim()}}`,
      end: grouped.end,
    };
  }

  let end = cursor;
  while (end < text.length) {
    const nextChar = text.slice(end).match(/^./u)?.[0] ?? "";
    if (!nextChar || !isRootAtomChar(nextChar)) break;
    end += nextChar.length;
  }
  if (end === cursor) return null;
  return {
    latex: `\\sqrt${rootIndex}{${text.slice(cursor, end)}}`,
    end,
  };
}

function normalizeRootNotation(text: string): string {
  let out = "";
  let index = 0;
  while (index < text.length) {
    if (text[index] !== "√") {
      out += text[index] ?? "";
      index += 1;
      continue;
    }
    const target = readVisualRootTarget(text, index + 1);
    if (!target) {
      out += text[index] ?? "";
      index += 1;
      continue;
    }
    out += target.latex;
    index = target.end;
  }
  return out;
}

function normalizeCollapsedScripts(text: string): string {
  const scriptBoundary = "(?=$|[\\s,.;:+\\-−=*/<>≤≥≈≠)\\]}_])";
  return text
    .replace(
      new RegExp(`([)\\]}])\\s*([0-9]+)${scriptBoundary}`, "g"),
      "$1^{$2}",
    )
    .replace(
      new RegExp(
        `([αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ])\\s*([0-9]+)${scriptBoundary}`,
        "g",
      ),
      "$1^{$2}",
    );
}

function normalizeMathLatinGlyphs(text: string): string {
  let out = "";
  for (let index = 0; index < text.length; ) {
    const ch = codePointCharAt(text, index);
    const cp = ch.codePointAt(0);
    out += cp == null ? ch : (mathLatinTokenFromCodePoint(cp) ?? ch);
    // Minimum advance of 1: a zero-length `ch` must never freeze the loop —
    // this exact loop spun forever when codePointCharAt returned "" for "\n".
    index += Math.max(1, ch.length);
  }
  return out;
}

function separateTightNamedOperatorTokens(text: string): string {
  const normalized = normalizeMathLatinGlyphs(text)
    .replace(
      /\blim(?=(?:sin|cos|tan|log|ln|exp|max|min)[A-Za-z]|\d|[A-Za-z]\s*[\[(])/g,
      "lim ",
    )
    .replace(
      /\b(sin|cos|tan|log|ln|exp|max|min)(?=[A-Za-z])/g,
      "$1 ",
    );
  // Before a named operator that follows a one-letter word (`(?<=\b[A-Za-z])`,
  // tested in code: Safari before 16.4 cannot compile a lookbehind).
  return replaceBehind(
    normalized,
    /(?=(?:sin|cos|tan|log|ln|exp|max|min)(?:[A-Za-z]|\\[A-Za-z]+|[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]))/g,
    (source, index) => index >= 1 && /[A-Za-z]/.test(source[index - 1]) && (index < 2 || !/\w/.test(source[index - 2])),
    () => " ",
  )
    .replace(
      /\b(sin|cos|tan|log|ln|exp|max|min)(?=[A-Za-z])/g,
      "$1 ",
    );
}

function normalizeNamedOperators(text: string): string {
  const named = separateTightNamedOperatorTokens(text)
    .replace(/\barg\s*max\b/gi, "\\arg\\max")
    .replace(/\barg\s*min\b/gi, "\\arg\\min")
    .replace(/\blim\b/g, "\\lim")
    .replace(/\bPr\b/g, "\\Pr")
    .replace(/\bdet\b/g, "\\det")
    .replace(/\b(?:Var|var)\b/g, (match) => `\\operatorname{${match}}`)
    .replace(/\b(?:Cov|cov)\b/g, (match) => `\\operatorname{${match}}`);
  // Not after a backslash: tested in code, no lookbehind (Safari < 16.4).
  return replaceBehind(named, /\b(?:sin|cos|tan|log|ln|exp|max|min)\b/g, { char: /\\/, not: true }, ([match]) => `\\${match}`);
}

function normalizeAsciiGreekSymbols(text: string): string {
  return text.replace(
    /(^|[^A-Za-z0-9\\])(alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|omega|Gamma|Delta|Theta|Lambda|Pi|Sigma|Omega)(?=$|[^A-Za-z0-9]|[_^])/g,
    (_match, prefix: string, name: string) =>
      `${prefix}${ASCII_GREEK_LATEX[name] ?? name}`,
  );
}

function normalizeExplicitAsciiScripts(text: string): string {
  const base = "(?:\\\\[A-Za-z]+|[A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ])";
  return text
    .replace(new RegExp(`(${base})\\s*_\\s*([A-Za-z0-9]+)`, "g"), "$1_{$2}")
    .replace(new RegExp(`(${base})\\s*\\^\\s*([A-Za-z0-9]+)`, "g"), "$1^{$2}");
}

function escapeLatexTextBody(text: string): string {
  return text.replace(/([\\{}])/g, "\\$1");
}

const TEXT_MODE_LATEX_COMMANDS = new Set([
  "text",
  "textrm",
  "textsf",
  "texttt",
  "textbf",
  "textit",
  "mathrm",
  "mbox",
  "operatorname",
]);

function isCjkLatexChar(ch: string): boolean {
  return /[\u3400-\u9fff]/u.test(ch);
}

function readLatexCommandToken(
  text: string,
  start: number,
): { token: string; name: string; end: number } | null {
  if (text[start] !== "\\") return null;
  const next = text[start + 1] ?? "";
  if (!next) return { token: "\\", name: "", end: start + 1 };
  if (!/[A-Za-z]/.test(next)) {
    return {
      token: text.slice(start, start + 2),
      name: next,
      end: start + 2,
    };
  }

  let end = start + 1;
  while (end < text.length && /[A-Za-z]/.test(text[end] ?? "")) end += 1;
  return {
    token: text.slice(start, end),
    name: text.slice(start + 1, end),
    end,
  };
}

function normalizePdfLatexSourceSegment(
  text: string,
  inTextMode: boolean,
): string {
  let out = "";
  let index = 0;

  while (index < text.length) {
    const ch = codePointCharAt(text, index);

    if (ch === "\\") {
      const command = readLatexCommandToken(text, index);
      if (!command) {
        out += ch;
        index += 1;
        continue;
      }

      out += command.token;
      index = command.end;
      if (TEXT_MODE_LATEX_COMMANDS.has(command.name)) {
        const groupStart = skipLatexSpaces(text, index);
        out += text.slice(index, groupStart);
        const group = readLatexGroup(text, groupStart);
        if (group) {
          out += `{${normalizePdfLatexSourceSegment(group.body, true)}}`;
          index = group.end;
        }
      }
      continue;
    }

    if (ch === "{") {
      const group = readLatexGroup(text, index);
      if (group) {
        out += `{${normalizePdfLatexSourceSegment(group.body, inTextMode)}}`;
        index = group.end;
        continue;
      }
    }

    if (!inTextMode) {
      const friendly = tryFriendlyFormulaFunction(text, index);
      if (friendly) {
        out += friendly.latex;
        index = friendly.end;
        continue;
      }
      if (SUPERSCRIPT_DIGITS[ch]) {
        let cursor = index;
        let script = "";
        while (cursor < text.length && SUPERSCRIPT_DIGITS[text[cursor] ?? ""]) {
          script += text[cursor] ?? "";
          cursor += 1;
        }
        out += scriptGroupToLatex(script, SUPERSCRIPT_DIGITS, "^");
        index = cursor;
        continue;
      }
      if (SUBSCRIPT_DIGITS[ch]) {
        let cursor = index;
        let script = "";
        while (cursor < text.length && SUBSCRIPT_DIGITS[text[cursor] ?? ""]) {
          script += text[cursor] ?? "";
          cursor += 1;
        }
        out += scriptGroupToLatex(script, SUBSCRIPT_DIGITS, "_");
        index = cursor;
        continue;
      }
      const unit = readSiUnitExpression(text, index);
      if (unit) {
        out += unit.latex;
        index = unit.end;
        continue;
      }
      const mathToken = mathAlphanumericLatexToken(ch);
      if (mathToken) {
        out = appendLatexToken(out, mathToken);
        index += ch.length;
        continue;
      }
      if (isCjkLatexChar(ch)) {
        let cursor = index;
        let run = "";
        while (cursor < text.length && isCjkLatexChar(text[cursor] ?? "")) {
          run += text[cursor] ?? "";
          cursor += 1;
        }
        out += `\\text{${escapeLatexTextBody(run)}}`;
        index = cursor;
        continue;
      }
      const mapped = GREEK_AND_SYMBOL_LATEX[ch];
      if (mapped) {
        out = appendLatexToken(out, mapped);
        index += ch.length;
        continue;
      }
    }

    out += ch === "\u00a0" ? " " : ch;
    index += ch.length;
  }

  return out;
}

/**
 * Normalize hand-written or persisted LaTeX before rendering/committing it.
 *
 * Generated formula seeds already go through `latexFromPdfFormulaText`, but old
 * saved edits and direct user input can still contain PDF glyph text such as
 * `∆𝑥米` or `\frac{8米}{8秒}`. KaTeX renders those with warnings or missing
 * metrics. This pass keeps existing LaTeX commands intact while normalizing raw
 * unicode math symbols and wrapping CJK unit runs in `\text{...}`.
 */
export function normalizePdfFormulaLatexSource(latex: string): string {
  const normalized = normalizePdfLatexSourceSegment(
    normalizeRootNotation(latex),
    false,
  );
  return normalizeLatexSpacing(normalizeSafeSlashFractions(normalized));
}

function latexForFormulaRow(row: string): string {
  const withRoots = normalizeRootNotation(
    normalizeNamedOperators(
      normalizeExplicitAsciiScripts(
        normalizeAsciiGreekSymbols(
          normalizeCollapsedScripts(normalizeMatrixBracketGlyphs(row)),
        ),
      ),
    ),
  );
  const withAccents = withRoots
    .replace(/˜\s*([A-Za-z])/g, "\\tilde{$1}")
    .replace(/\b([A-Za-z])\s*˜(?=\s*\()/g, "\\tilde{$1}");
  const withScripts = withAccents
    .replace(SUPERSCRIPT_RE, (match) =>
      scriptGroupToLatex(match, SUPERSCRIPT_DIGITS, "^"),
    )
    .replace(SUBSCRIPT_RE, (match) =>
      scriptGroupToLatex(match, SUBSCRIPT_DIGITS, "_"),
    );

  let out = "";
  let index = 0;
  while (index < withScripts.length) {
    const unit = readSiUnitExpression(withScripts, index);
    if (unit) {
      out += unit.latex;
      index = unit.end;
      continue;
    }
    const ch = codePointCharAt(withScripts, index);
    const mathToken = mathAlphanumericLatexToken(ch);
    if (mathToken) {
      out = appendLatexToken(out, mathToken);
      index += ch.length;
      continue;
    }
    const mapped = GREEK_AND_SYMBOL_LATEX[ch];
    if (mapped) {
      out = appendLatexToken(out, mapped);
      index += ch.length;
      continue;
    }
    if (ch === "\u00a0") {
      out += " ";
      index += ch.length;
      continue;
    }
    out += ch;
    index += ch.length;
  }
  return normalizePdfFormulaLatexSource(out);
}

function runRight(run: PdfFormulaRun): number {
  if (typeof run.right === "number" && Number.isFinite(run.right)) {
    return run.right;
  }
  return (
    run.left + Math.max(run.fontSize, run.text.length * run.fontSize * 0.55)
  );
}

function runCenterX(run: PdfFormulaRun): number {
  return (run.left + runRight(run)) / 2;
}

function rowCenter(row: readonly PdfFormulaRun[]): number {
  return (
    row.reduce((sum, run) => sum + run.centerY, 0) / Math.max(1, row.length)
  );
}

function runBaselineY(run: PdfFormulaRun): number {
  return typeof run.baselineY === "number" && Number.isFinite(run.baselineY)
    ? run.baselineY
    : run.centerY;
}

function rowBaselineY(row: readonly PdfFormulaRun[]): number {
  return (
    row.reduce((sum, run) => sum + runBaselineY(run), 0) /
    Math.max(1, row.length)
  );
}

function rowFont(row: readonly PdfFormulaRun[]): number {
  return Math.max(...row.map((run) => run.fontSize), 1);
}

function rowLeft(row: readonly PdfFormulaRun[]): number {
  return Math.min(...row.map((run) => run.left));
}

function rowRight(row: readonly PdfFormulaRun[]): number {
  return Math.max(...row.map(runRight));
}

function rowWidth(row: readonly PdfFormulaRun[]): number {
  return Math.max(1, rowRight(row) - rowLeft(row));
}

function latexFromGeometryRootRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  if (rows.length !== 1) return null;

  const row = [...rows[0]].sort((a, b) => a.left - b.left);
  const parts: string[] = [];
  let usedGeometry = false;

  for (let index = 0; index < row.length; index += 1) {
    const run = row[index];
    if (!run.text.includes("√")) {
      parts.push(latexForFormulaRow(run.text));
      continue;
    }
    if (run.text.replace(/√/g, "").trim()) return null;

    const spanRight = runRight(run);
    if (spanRight - run.left < run.fontSize * 1.45) {
      parts.push(latexForFormulaRow(run.text));
      continue;
    }

    const pad = run.fontSize * 0.35;
    const radicand: PdfFormulaRun[] = [];
    let cursor = index + 1;
    while (cursor < row.length) {
      const next = row[cursor];
      if (next.left < run.left || runCenterX(next) > spanRight + pad) break;
      radicand.push(next);
      cursor += 1;
    }

    if (radicand.length === 0) {
      parts.push(latexForFormulaRow(run.text));
      continue;
    }
    parts.push(`\\sqrt{${latexRowText(radicand)}}`);
    usedGeometry = true;
    index = cursor - 1;
  }

  return usedGeometry ? joinLatexFormulaParts(parts) : null;
}

function rowsFromFormulaRuns(
  runs: readonly PdfFormulaRun[],
): PdfFormulaRun[][] {
  const sorted = [...runs]
    .map((run) => ({ ...run, text: run.text.trim() }))
    .filter((run) => run.text.length > 0)
    .sort((a, b) => a.centerY - b.centerY || a.left - b.left);

  const rows: PdfFormulaRun[][] = [];
  for (const run of sorted) {
    const row = rows[rows.length - 1];
    if (!row) {
      rows.push([run]);
      continue;
    }

    const largerFont = Math.max(rowFont(row), run.fontSize);
    const smallerFont = Math.min(rowFont(row), run.fontSize);
    const center = rowCenter(row);
    const centerDelta = Math.abs(run.centerY - center);
    const baselineDelta = Math.abs(runBaselineY(run) - rowBaselineY(row));
    const likelyScriptOffset =
      smallerFont <= largerFont * 0.72 &&
      centerDelta > smallerFont * 0.15 &&
      baselineDelta > Math.max(smallerFont * 0.42, largerFont * 0.24);
    if (centerDelta > largerFont * 0.42 || likelyScriptOffset) {
      rows.push([run]);
      continue;
    }

    row.push(run);
  }
  return rows;
}

function plainRowText(row: readonly PdfFormulaRun[]): string {
  const fragments = [...row]
    .sort((a, b) => a.left - b.left)
    .map((run) => run.text.trim())
    .filter(Boolean);
  return joinFormulaRunTextFragments(fragments);
}

function shouldJoinFormulaRunTextWithoutSpace(
  previous: string,
  next: string,
): boolean {
  if (!previous || !next) return false;
  if (/^[,.;:!?，。；：、！？）)\]}]/.test(next)) return true;
  if (/[（([{]$/.test(previous)) return true;
  if (/^[_^]/.test(next)) return true;
  if (
    /^[([{]/.test(next) &&
    /[A-Za-z0-9αβγδθλμνπρσφψωΑ-Ω₀₁₂₃₄₅₆₇₈₉₊₋₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ)\]}]$/u.test(
      previous,
    )
  ) {
    return true;
  }
  return false;
}

function joinFormulaRunTextFragments(fragments: readonly string[]): string {
  let text = "";
  for (const fragment of fragments) {
    const next = fragment.trim();
    if (!next) continue;
    if (!text) {
      text = next;
      continue;
    }
    text += shouldJoinFormulaRunTextWithoutSpace(text, next) ? next : ` ${next}`;
  }
  return text.replace(/\s+/g, " ").trim();
}

function latexRowText(row: readonly PdfFormulaRun[]): string {
  return latexForFormulaRow(plainRowText(row));
}

function alignedLatexFromRows(rows: readonly PdfFormulaRun[][]): string {
  return `\\begin{aligned}${rows
    .map((row) => latexRowText(row))
    .filter(Boolean)
    .join(" \\\\ ")}\\end{aligned}`;
}

function compactLatexSignal(text: string): string {
  return text.replace(/\\begin\{aligned\}|\\end\{aligned\}|\\\\|\s+/g, "");
}

function latexContentWeight(text: string): number {
  return compactLatexSignal(text).replace(/[{}_^]/g, "").length;
}

function preservesMostFormulaContent(
  candidate: string,
  fallback: string,
): boolean {
  const candidateSignal = compactLatexSignal(candidate);
  const fallbackSignal = compactLatexSignal(fallback);
  if (fallbackSignal.length < 36) return true;
  return candidateSignal.length >= fallbackSignal.length * 0.45;
}

function isFractionBarText(text: string): boolean {
  return /^[\-_−–—―─━]+$/.test(text.replace(/\s+/g, ""));
}

function isFractionBarRun(run: PdfFormulaRun): boolean {
  return (
    isFractionBarText(run.text) &&
    runRight(run) - run.left >= run.fontSize * 1.8
  );
}

function isFractionBarRow(row: readonly PdfFormulaRun[]): boolean {
  const text = plainRowText(row);
  return isFractionBarText(text) && rowWidth(row) >= rowFont(row) * 1.8;
}

function rowIsCenteredOnBar(
  row: readonly PdfFormulaRun[],
  bar: readonly PdfFormulaRun[],
): boolean {
  const rowCenterX =
    (Math.min(...row.map((run) => run.left)) + Math.max(...row.map(runRight))) /
    2;
  const barLeft = Math.min(...bar.map((run) => run.left));
  const barRight = Math.max(...bar.map(runRight));
  const pad = Math.max(rowFont(row), rowFont(bar)) * 2;
  return rowCenterX >= barLeft - pad && rowCenterX <= barRight + pad;
}

function isStandaloneRootRun(run: PdfFormulaRun): boolean {
  return (
    run.text.includes("√") && run.text.replace(/√/g, "").trim().length === 0
  );
}

function isRootIndexRow(
  row: readonly PdfFormulaRun[],
  root: PdfFormulaRun,
  rootRow: readonly PdfFormulaRun[],
): boolean {
  if (row.length !== 1) return false;
  const text = plainRowText(row);
  if (!text || !/^[+\-−]?(?:\d+|[A-Za-z])$/.test(text)) return false;

  const rootFont = Math.max(root.fontSize, rowFont(rootRow), 1);
  if (rowFont(row) > rootFont * 0.75) return false;
  if (rowCenter(row) >= rowCenter(rootRow)) return false;

  const rootLeft = root.left;
  const rootRight = runRight(root);
  return (
    rowRight(row) >= rootLeft - rootFont * 1.1 &&
    rowLeft(row) <= rootRight + rootFont * 0.25 &&
    runCenterX(row[0]) <= rootLeft + rootFont * 0.45
  );
}

function rowLooksLikePositionedScript(
  row: readonly PdfFormulaRun[],
  baseRow: readonly PdfFormulaRun[],
): boolean {
  const baseFont = rowFont(baseRow);
  const text = plainRowText(row);
  return (
    rowFont(row) <= baseFont * 0.82 ||
    /^[\d+\-−()n\s]+$/.test(text)
  );
}

function rowHorizontallyOverlaps(
  row: readonly PdfFormulaRun[],
  left: number,
  right: number,
  pad: number,
): boolean {
  return rowRight(row) >= left - pad && rowLeft(row) <= right + pad;
}

function latexFromRowWithPositionedScripts(
  baseRow: readonly PdfFormulaRun[],
  scriptRows: readonly PdfFormulaRun[][],
): string | null {
  if (scriptRows.length === 0) return latexRowText(baseRow);

  const baseY = rowCenter(baseRow);
  const chars = baseCharacters(baseRow);
  const insertions = new Map<number, { sup: string[]; sub: string[] }>();

  for (const row of scriptRows) {
    if (!rowLooksLikePositionedScript(row, baseRow)) return null;
    for (const run of row) {
      const script = latexForFormulaRow(run.text);
      if (!script) continue;
      const anchor = anchorIndexForScript(chars, runCenterX(run));
      const bucket = insertions.get(anchor) ?? { sup: [], sub: [] };
      if (rowCenter(row) < baseY) bucket.sup.push(script);
      else bucket.sub.push(script);
      insertions.set(anchor, bucket);
    }
  }
  if (insertions.size === 0) return null;

  let raw = "";
  chars.forEach(({ char }, index) => {
    raw += char;
    const bucket = insertions.get(index);
    if (!bucket) return;
    if (bucket.sub.length > 0) raw += `_{${bucket.sub.join(" ")}}`;
    if (bucket.sup.length > 0) raw += `^{${bucket.sup.join(" ")}}`;
  });
  return latexForFormulaRow(raw);
}

function latexFromVisualRootRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  const rootRows = rows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => row.some(isStandaloneRootRun));
  if (rootRows.length !== 1) return null;

  const { row: rootRow, index: rootRowIndex } = rootRows[0];
  const sortedRootRow = [...rootRow].sort((a, b) => a.left - b.left);
  const rootRuns = sortedRootRow.filter(isStandaloneRootRun);
  if (rootRuns.length === 0) return null;

  const explicitBarCandidates = rows.flatMap((row, rowIndex) =>
    row
      .filter(isFractionBarRun)
      .filter(() => row.every(isFractionBarRun) && rowIndex < rootRowIndex)
      .map((bar) => ({ bar, rowIndex })),
  );

  type VisualRootBar = {
    left: number;
    right: number;
    fontSize: number;
    rowIndex: number | null;
  };
  type VisualRootDescriptor = {
    root: PdfFormulaRun;
    bar: VisualRootBar;
    radicand: PdfFormulaRun[];
    rootIndex?: string;
    rootIndexRow?: PdfFormulaRun[];
    scriptRows: PdfFormulaRun[][];
  };

  const usedExplicitBars = new Set<PdfFormulaRun>();
  const descriptors: VisualRootDescriptor[] = [];
  for (const root of rootRuns) {
    const rootFont = Math.max(root.fontSize, 1);
    const bestExplicitBar = explicitBarCandidates
      .filter(({ bar, rowIndex }) => {
        if (usedExplicitBars.has(bar)) return false;
        const barRight = runRight(bar);
        const verticalDistance = rowCenter(rootRow) - rowCenter(rows[rowIndex]);
        return (
          verticalDistance > 0 &&
          verticalDistance <= rootFont * 1.6 &&
          barRight > root.left + rootFont * 0.75 &&
          bar.left <= runRight(root) + rootFont * 1.8
        );
      })
      .sort(
        (a, b) =>
          Math.abs(rowCenter(rows[a.rowIndex]) - rowCenter(rootRow)) -
            Math.abs(rowCenter(rows[b.rowIndex]) - rowCenter(rootRow)) ||
          Math.abs(a.bar.left - runRight(root)) -
            Math.abs(b.bar.left - runRight(root)),
      )[0];
    const carriedBarRight = runRight(root);
    const carriedBar =
      !bestExplicitBar && carriedBarRight - root.left >= rootFont * 1.45
        ? {
            left: root.left + rootFont * 0.75,
            right: carriedBarRight,
            fontSize: rootFont,
            rowIndex: null,
          }
        : null;
    const bar = bestExplicitBar
      ? {
          left: bestExplicitBar.bar.left,
          right: runRight(bestExplicitBar.bar),
          fontSize: bestExplicitBar.bar.fontSize,
          rowIndex: bestExplicitBar.rowIndex,
        }
      : carriedBar;
    if (!bar) return null;
    if (bestExplicitBar) usedExplicitBars.add(bestExplicitBar.bar);

    descriptors.push({ root, bar, radicand: [], scriptRows: [] });
  }

  const usedBarRowIndexes = new Set(
    descriptors
      .map((descriptor) => descriptor.bar.rowIndex)
      .filter((index): index is number => index !== null),
  );
  const extraRows = rows.filter(
    (_, index) => index !== rootRowIndex && !usedBarRowIndexes.has(index),
  );

  const radicandOwner = new Map<PdfFormulaRun, VisualRootDescriptor>();
  for (const descriptor of descriptors) {
    const rootFont = Math.max(descriptor.root.fontSize, 1);
    const pad = Math.max(rootFont, descriptor.bar.fontSize) * 0.45;
    const radicand = sortedRootRow.filter(
      (run) =>
        run !== descriptor.root &&
        !isStandaloneRootRun(run) &&
        run.left >= descriptor.root.left + rootFont * 0.25 &&
        runCenterX(run) >= descriptor.bar.left - pad &&
        runCenterX(run) <= descriptor.bar.right + pad,
    );
    if (radicand.length === 0) return null;
    for (const run of radicand) {
      if (radicandOwner.has(run)) return null;
      radicandOwner.set(run, descriptor);
    }
    descriptor.radicand = radicand;
  }

  const rootIndexRowSet = new Set<PdfFormulaRun[]>();
  for (const descriptor of descriptors) {
    const rootIndexRows = extraRows
      .map((row) => ({ row, latex: latexRowText(row) }))
      .filter(
        ({ row, latex }) =>
          latex.length > 0 && isRootIndexRow(row, descriptor.root, rootRow),
      );
    if (rootIndexRows.length > 1) return null;
    const rootIndex = rootIndexRows[0];
    if (!rootIndex) continue;
    if (rootIndexRowSet.has(rootIndex.row)) return null;
    rootIndexRowSet.add(rootIndex.row);
    descriptor.rootIndex = rootIndex.latex;
    descriptor.rootIndexRow = rootIndex.row;
  }

  for (const row of extraRows) {
    if (rootIndexRowSet.has(row)) continue;
    const matchingDescriptors = descriptors.filter((descriptor) => {
      const rootFont = Math.max(descriptor.root.fontSize, 1);
      return (
        rowLooksLikePositionedScript(row, descriptor.radicand) &&
        rowHorizontallyOverlaps(
          row,
          descriptor.bar.left,
          descriptor.bar.right,
          rootFont * 0.65,
        )
      );
    });
    if (matchingDescriptors.length > 1) return null;
    if (matchingDescriptors.length === 1) {
      matchingDescriptors[0].scriptRows.push(row);
    }
  }

  const allowedExtraRows = new Set<PdfFormulaRun[]>([
    ...rootIndexRowSet,
    ...descriptors.flatMap((descriptor) => descriptor.scriptRows),
  ]);
  const descriptorByRoot = new Map(
    descriptors.map((descriptor) => [descriptor.root, descriptor]),
  );
  const radicandSet = new Set(radicandOwner.keys());
  const radicandLatexByRoot = new Map<PdfFormulaRun, string>();
  for (const descriptor of descriptors) {
    const radicandLatex =
      latexFromRowWithPositionedScripts(
        descriptor.radicand,
        descriptor.scriptRows,
      ) ?? latexRowText(descriptor.radicand);
    radicandLatexByRoot.set(descriptor.root, radicandLatex);
  }

  const parts: string[] = [];
  for (const run of sortedRootRow) {
    const descriptor = descriptorByRoot.get(run);
    if (descriptor) {
      const radicandLatex = radicandLatexByRoot.get(run) ?? "";
      parts.push(
        descriptor.rootIndex
          ? `\\sqrt[${descriptor.rootIndex}]{${radicandLatex}}`
          : `\\sqrt{${radicandLatex}}`,
      );
    } else if (!radicandSet.has(run)) {
      parts.push(latexForFormulaRow(run.text));
    }
  }

  if (
    extraRows.some(
      (row) =>
        row.some((run) => !isFractionBarRun(run)) && !allowedExtraRows.has(row),
    )
  ) {
    return null;
  }

  return normalizeLatexSpacing(joinLatexFormulaParts(parts));
}

function latexFromFractionRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  if (rows.length !== 3) return null;
  const barIndex = rows.findIndex(isFractionBarRow);
  if (barIndex <= 0 || barIndex >= rows.length - 1) return null;

  const numerator = rows[barIndex - 1];
  const bar = rows[barIndex];
  const denominator = rows[barIndex + 1];
  if (bar.filter(isFractionBarRun).length !== 1) return null;
  if (
    !rowIsCenteredOnBar(numerator, bar) ||
    !rowIsCenteredOnBar(denominator, bar)
  ) {
    return null;
  }

  return normalizeLatexSpacing(
    `\\frac{${latexRowText(numerator)}}{${latexRowText(denominator)}}`,
  );
}

function rowFitsBarSpan(
  row: readonly PdfFormulaRun[],
  bar: PdfFormulaRun,
): boolean {
  const left = Math.min(...row.map((run) => run.left));
  const right = Math.max(...row.map(runRight));
  const pad = Math.max(rowFont(row), bar.fontSize) * 2.5;
  return left >= bar.left - pad && right <= runRight(bar) + pad;
}

function rowFractionRunsForBar(
  row: readonly PdfFormulaRun[],
  bar: PdfFormulaRun,
): PdfFormulaRun[] {
  const pad = Math.max(rowFont(row), bar.fontSize) * 2.5;
  const barLeft = bar.left - pad;
  const barRight = runRight(bar) + pad;
  return row.filter((run) => {
    const center = runCenterX(run);
    return center >= barLeft && center <= barRight;
  });
}

function nearestFractionRow(
  rows: readonly PdfFormulaRun[][],
  barRowIndex: number,
  bar: PdfFormulaRun,
  direction: "above" | "below",
): PdfFormulaRun[] | null {
  const barCenter = rowCenter(rows[barRowIndex]);
  return (
    rows
      .map((row, index) => ({ row, index }))
      .filter(({ index }) =>
        direction === "above" ? index < barRowIndex : index > barRowIndex,
      )
      .map(({ row, index }) => ({
        row: rowFractionRunsForBar(row, bar),
        index,
      }))
      .filter(({ row }) => row.length > 0)
      .filter(({ row }) => rowIsCenteredOnBar(row, [bar]))
      .filter(({ row }) => rowFitsBarSpan(row, bar))
      .sort(
        (a, b) =>
          Math.abs(rowCenter(a.row) - barCenter) -
          Math.abs(rowCenter(b.row) - barCenter),
      )[0]?.row ?? null
  );
}

function latexFromLocalFractionRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  const barCandidates = rows.flatMap((row, rowIndex) =>
    row.filter(isFractionBarRun).map((run) => ({ row, rowIndex, run })),
  );
  if (barCandidates.length === 0) return null;

  const fractionRunsByBar = new Map<PdfFormulaRun, PdfFormulaRun>();
  const consumed = new Set<PdfFormulaRun>();

  for (const { row: barRow, rowIndex, run: bar } of barCandidates) {
    const numerator = nearestFractionRow(rows, rowIndex, bar, "above");
    const denominator = nearestFractionRow(rows, rowIndex, bar, "below");
    if (!numerator || !denominator) return null;
    if (
      numerator.some((run) => consumed.has(run)) ||
      denominator.some((run) => consumed.has(run))
    ) {
      return null;
    }

    const fractionLatex = `\\frac{${latexRowText(numerator)}}{${latexRowText(
      denominator,
    )}}`;
    const fractionRun: PdfFormulaRun = {
      ...bar,
      text: fractionLatex,
      left: Math.min(bar.left, rowLeft(numerator), rowLeft(denominator)),
      right: Math.max(runRight(bar), rowRight(numerator), rowRight(denominator)),
      centerY: rowCenter(barRow),
      fontSize: Math.max(
        rowFont(barRow),
        rowFont(numerator),
        rowFont(denominator),
      ),
    };
    fractionRunsByBar.set(bar, fractionRun);
    consumed.add(bar);
    for (const run of numerator) consumed.add(run);
    for (const run of denominator) consumed.add(run);
  }

  const fractionRuns = new Set(fractionRunsByBar.values());
  const transformedRows = rows
    .map((row) => {
      if (row.some((run) => fractionRunsByBar.has(run))) {
        return row
          .flatMap((run) => {
            const fractionRun = fractionRunsByBar.get(run);
            if (fractionRun) return [fractionRun];
            return consumed.has(run) ? [] : [run];
          })
          .sort((a, b) => a.left - b.left);
      }
      return row.filter((run) => !consumed.has(run));
    })
    .filter((row) => row.length > 0);
  const transformedBaseRow = transformedRows.find((row) =>
    row.some((run) => fractionRuns.has(run)),
  );
  if (!transformedBaseRow) return null;

  if (transformedRows.length === 1) {
    return latexRowText(transformedBaseRow);
  }

  const scripted = latexWithPositionedScripts(transformedRows);
  return scripted?.trim() ? normalizeLatexSpacing(scripted) : null;
}

function latexFromImplicitLocalFractionRowsForBase(
  rows: readonly PdfFormulaRun[][],
  baseRow: readonly PdfFormulaRun[],
): string | null {
  const baseFont = rowFont(baseRow);
  const baseCenter = rowCenter(baseRow);
  const above = rows
    .filter((row) => row !== baseRow && rowCenter(row) < baseCenter)
    .sort((a, b) => baseCenter - rowCenter(a) - (baseCenter - rowCenter(b)))[0];
  const below = rows
    .filter((row) => row !== baseRow && rowCenter(row) > baseCenter)
    .sort((a, b) => rowCenter(a) - baseCenter - (rowCenter(b) - baseCenter))[0];
  if (!above || !below) return null;
  if (rowFont(above) > baseFont * 1.1 || rowFont(below) > baseFont * 1.1) {
    return null;
  }
  if (
    Math.abs(rowCenter(above) - baseCenter) > baseFont * 1.4 ||
    Math.abs(rowCenter(below) - baseCenter) > baseFont * 1.4
  ) {
    return null;
  }

  const fractionLeft = Math.min(rowLeft(above), rowLeft(below));
  const fractionRight = Math.max(rowRight(above), rowRight(below));
  const fractionCenter = (fractionLeft + fractionRight) / 2;
  const numeratorCenter = (rowLeft(above) + rowRight(above)) / 2;
  const denominatorCenter = (rowLeft(below) + rowRight(below)) / 2;
  if (Math.abs(numeratorCenter - denominatorCenter) > baseFont * 2.2) {
    return null;
  }

  const baseText = plainRowText(baseRow);
  const fraction = `\\frac{${latexRowText(above)}}{${latexRowText(below)}}`;

  const missingFractionSlot = /^(.*?)=\s*=\s*(.+)$/.exec(
    baseText.replace(/\s+/g, " ").trim(),
  );
  if (missingFractionSlot) {
    const before = latexForFormulaRow(missingFractionSlot[1].trim());
    const after = latexForFormulaRow(missingFractionSlot[2].trim());
    return joinLatexFormulaParts([before, "=", fraction, "=", after]);
  }

  const beforeRuns = baseRow.filter(
    (run) => runRight(run) <= fractionLeft + baseFont * 0.55,
  );
  const afterRuns = baseRow.filter(
    (run) => run.left >= fractionRight - baseFont * 0.55,
  );
  const baseHasVisibleFractionSlot =
    /[=+\-−×÷*/()[\]]/.test(baseText) ||
    /(?:^|[^A-Za-z])(?:ln|log|exp|sin|cos|tan|Pr)\s*[\[(]/.test(baseText);
  if (!baseHasVisibleFractionSlot && beforeRuns.length === 0 && afterRuns.length === 0) {
    return null;
  }
  if (beforeRuns.length === 0 && afterRuns.length === 0) return null;
  if (
    beforeRuns.length === 0 &&
    fractionCenter < rowRight(baseRow) + baseFont * 0.25
  ) {
    return null;
  }

  const before = beforeRuns.length > 0 ? latexRowText(beforeRuns) : "";
  const after = afterRuns.length > 0 ? latexRowText(afterRuns) : "";
  return joinLatexFormulaParts([before, fraction, after]);
}

function latexFromImplicitLocalFractionRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  if (rows.length !== 3) return null;

  const dominant = dominantBaseRow(rows);
  const candidateBaseRows = [
    dominant,
    ...rows.filter((row) => row !== dominant),
  ];
  for (const baseRow of candidateBaseRows) {
    const candidate = latexFromImplicitLocalFractionRowsForBase(rows, baseRow);
    if (candidate) return candidate;
  }
  return null;
}

function rowLooksLikeFractionPart(row: readonly PdfFormulaRun[]): boolean {
  const text = plainRowText(row);
  if (!text || /[=<>≤≥≈≠]/.test(text)) return false;
  if ((text.match(/[A-Za-z]{3,}/g) ?? []).length > 0) return false;
  return latexRowText(row).trim().length > 0;
}

function rowLooksLikeSafeStackedDenominator(
  row: readonly PdfFormulaRun[],
): boolean {
  const compact = plainRowText(row).replace(/\s+/g, "");
  if (!compact) return false;
  if (/^[A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]$/.test(compact)) {
    return false;
  }
  return (
    /^[+\-−]?\d/.test(compact) ||
    /[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ∂∇]/.test(compact) ||
    /[+\-−*/×÷()[\]{}^_]/.test(compact)
  );
}

function isCorruptedZeroGlyphRun(run: PdfFormulaRun): boolean {
  return run.text.trim() === "!";
}

function rowsLookLikeCorruptedZeroStackedFraction(
  above: readonly PdfFormulaRun[],
  below: readonly PdfFormulaRun[],
): boolean {
  if (above.length !== 1 || below.length !== 1) return false;
  const numerator = above[0];
  const denominator = below[0];
  if (
    !isCorruptedZeroGlyphRun(numerator) ||
    !isCorruptedZeroGlyphRun(denominator)
  ) {
    return false;
  }

  const baseFont = Math.max(numerator.fontSize, denominator.fontSize, 1);
  const smallerFont = Math.max(
    0.001,
    Math.min(numerator.fontSize, denominator.fontSize),
  );
  if (Math.max(numerator.fontSize, denominator.fontSize) / smallerFont > 1.25) {
    return false;
  }

  const verticalDistance = denominator.centerY - numerator.centerY;
  if (verticalDistance < baseFont * 0.75 || verticalDistance > baseFont * 3.6) {
    return false;
  }

  const numeratorWidth = runRight(numerator) - numerator.left;
  const denominatorWidth = runRight(denominator) - denominator.left;
  if (Math.max(numeratorWidth, denominatorWidth) > baseFont * 1.25) {
    return false;
  }

  return (
    Math.abs(runCenterX(numerator) - runCenterX(denominator)) <=
    Math.max(baseFont * 0.65, Math.max(numeratorWidth, denominatorWidth))
  );
}

function latexFromStackedFractionRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  if (rows.length !== 2) return null;

  const [above, below] = [...rows].sort((a, b) => rowCenter(a) - rowCenter(b));
  if (rowsLookLikeCorruptedZeroStackedFraction(above, below)) {
    return "\\frac{0}{0}";
  }

  if (!rowLooksLikeFractionPart(above) || !rowLooksLikeFractionPart(below)) {
    return null;
  }
  if (!rowLooksLikeSafeStackedDenominator(below)) return null;

  const baseFont = Math.max(rowFont(above), rowFont(below));
  const verticalDistance = rowCenter(below) - rowCenter(above);
  if (verticalDistance < baseFont * 0.75 || verticalDistance > baseFont * 3.6) {
    return null;
  }

  const numeratorCenter = (rowLeft(above) + rowRight(above)) / 2;
  const denominatorCenter = (rowLeft(below) + rowRight(below)) / 2;
  if (Math.abs(numeratorCenter - denominatorCenter) > baseFont * 2.2) {
    return null;
  }
  const widthImbalance = rowWidth(above) / Math.max(1, rowWidth(below));
  if (widthImbalance > 3.2 && rowWidth(above) > rowWidth(below) + baseFont * 3) {
    return null;
  }

  return normalizeLatexSpacing(
    `\\frac{${latexRowText(above)}}{${latexRowText(below)}}`,
  );
}

function dominantBaseRow(rows: readonly PdfFormulaRun[][]): PdfFormulaRun[] {
  return [...rows].sort((a, b) => {
    const widthDelta = rowWidth(b) - rowWidth(a);
    if (Math.abs(widthDelta) > Math.max(rowFont(a), rowFont(b)) * 0.4) {
      return widthDelta;
    }
    return rowFont(b) - rowFont(a);
  })[0];
}

function operatorLatex(operator: string): string {
  return operator === "∑" ? "\\sum" : operator === "∏" ? "\\prod" : "\\int";
}

type LargeOperator = "∫" | "∑" | "∏";

interface OperatorAnchor {
  operator: LargeOperator;
  centerX: number;
}

interface OperatorSlot {
  anchor: OperatorAnchor;
  left: number;
  right: number;
}

interface OperatorSlice {
  text: string;
  left: number;
  right: number;
  centerX: number;
  fontSize: number;
  width: number;
}

interface NamedLimitOperator {
  raw: string;
  latex: string;
}

const NAMED_LIMIT_OPERATORS: readonly NamedLimitOperator[] = [
  { raw: "arg max", latex: "\\arg\\max" },
  { raw: "argmin", latex: "\\arg\\min" },
  { raw: "argmax", latex: "\\arg\\max" },
  { raw: "lim", latex: "\\lim" },
  { raw: "max", latex: "\\max" },
  { raw: "min", latex: "\\min" },
];

function namedLimitMatch(
  text: string,
): { operator: NamedLimitOperator; index: number; rawLength: number } | null {
  for (const operator of NAMED_LIMIT_OPERATORS) {
    const pattern = operator.raw.includes(" ")
      ? operator.raw.replace(/\s+/g, "\\s+")
      : operator.raw;
    const match = new RegExp(`(^|\\s)(${pattern})(?=\\s|$|[_({[])`, "i").exec(
      text,
    );
    if (match?.index != null) {
      return {
        operator,
        index: match.index + match[1].length,
        rawLength: match[2].length,
      };
    }
  }
  return null;
}

function namedOperatorCenterX(
  row: readonly PdfFormulaRun[],
  start: number,
  length: number,
): number {
  const left = Math.min(...row.map((run) => run.left));
  const width = rowWidth(row);
  const textLength = Math.max(1, plainRowText(row).length);
  return left + (width / textLength) * (start + length / 2);
}

function isNamedOperatorLimitRow(
  row: readonly PdfFormulaRun[],
  opCenter: number,
  opFont: number,
): boolean {
  const left = Math.min(...row.map((run) => run.left));
  const right = Math.max(...row.map(runRight));
  const font = rowFont(row);
  const width = rowWidth(row);
  return (
    opCenter >= left - font * 2.5 &&
    opCenter <= right + font * 2.5 &&
    font <= opFont * 0.95 &&
    width <= Math.max(font * 8, opFont * 4.5)
  );
}

function operatorCenterX(
  row: readonly PdfFormulaRun[],
  operator: string,
): number {
  const run = row.find((candidate) => candidate.text.includes(operator));
  if (!run) return runCenterX(row[0]);
  const chars = Array.from(run.text);
  const index = Math.max(0, chars.indexOf(operator));
  const width = Math.max(run.fontSize, runRight(run) - run.left);
  return run.left + (width / Math.max(1, chars.length)) * (index + 0.5);
}

function operatorRowCandidate(
  rows: readonly PdfFormulaRun[][],
): { row: PdfFormulaRun[]; operator: LargeOperator } | null {
  for (const operator of ["∫", "∑", "∏"] as const) {
    const row = rows.find((candidate) =>
      candidate.some((run) => run.text.includes(operator)),
    );
    if (row) return { row, operator };
  }
  return null;
}

function operatorAnchorsFromRow(
  row: readonly PdfFormulaRun[],
): OperatorAnchor[] {
  return baseCharacters(row)
    .filter((char): char is { char: LargeOperator; x: number } =>
      char.char === "∫" || char.char === "∑" || char.char === "∏",
    )
    .map((char) => ({
      operator: char.char,
      centerX: char.x,
    }))
    .sort((a, b) => a.centerX - b.centerX);
}

function operatorSlotsFromAnchors(
  anchors: readonly OperatorAnchor[],
): OperatorSlot[] {
  return anchors.map((anchor, index) => {
    const previous = anchors[index - 1];
    const next = anchors[index + 1];
    return {
      anchor,
      left: previous ? (previous.centerX + anchor.centerX) / 2 : -Infinity,
      right: next ? (anchor.centerX + next.centerX) / 2 : Infinity,
    };
  });
}

function rowSliceForOperatorSlot(
  row: readonly PdfFormulaRun[],
  slot: OperatorSlot,
  edgePadding = 0,
): OperatorSlice | null {
  const chars = baseCharacters(row).filter(
    (char) => char.x >= slot.left - edgePadding && char.x < slot.right + edgePadding,
  );
  const text = chars
    .map((char) => char.char)
    .join("")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return null;

  const geometryChars = chars.filter((char) => char.char.trim().length > 0);
  if (geometryChars.length === 0) return null;

  const fontSize = rowFont(row);
  const left = Math.min(...geometryChars.map((char) => char.x));
  const right = Math.max(...geometryChars.map((char) => char.x));
  const width = Math.max(fontSize, right - left);
  return {
    text,
    left,
    right,
    centerX: (left + right) / 2,
    fontSize,
    width,
  };
}

function isOperatorLimitSlice(
  slice: OperatorSlice,
  slot: OperatorSlot,
  opFont: number,
): boolean {
  return (
    Math.abs(slice.centerX - slot.anchor.centerX) <=
      Math.max(opFont * 1.65, slice.fontSize * 2.5) &&
    slice.fontSize <= opFont * 0.9 &&
    slice.width <= Math.max(slice.fontSize * 5, opFont * 3.2)
  );
}

function nearestOperatorSlotLimitText(
  rows: readonly PdfFormulaRun[][],
  opRow: readonly PdfFormulaRun[],
  slot: OperatorSlot,
  side: "above" | "below",
  opFont: number,
): string {
  const opCenterY = rowCenter(opRow);
  const candidates = rows
    .filter((row) => {
      if (row === opRow) return false;
      return side === "above"
        ? rowCenter(row) < opCenterY
        : rowCenter(row) > opCenterY;
    })
    .sort(
      (a, b) =>
        Math.abs(rowCenter(a) - opCenterY) -
        Math.abs(rowCenter(b) - opCenterY),
    );

  for (const row of candidates) {
    const slice = rowSliceForOperatorSlot(row, slot);
    if (!slice || !isOperatorLimitSlice(slice, slot, opFont)) continue;
    return latexForFormulaRow(slice.text);
  }
  return "";
}

function nearestOperatorSlotBodyText(
  rows: readonly PdfFormulaRun[][],
  opRow: readonly PdfFormulaRun[],
  slot: OperatorSlot,
  opFont: number,
): string {
  const opCenterY = rowCenter(opRow);
  const candidates = rows
    .filter((row) => row !== opRow)
    .flatMap((row) => {
      const slice = rowSliceForOperatorSlot(row, slot);
      if (
        !slice ||
        isOperatorLimitSlice(slice, slot, opFont) ||
        slice.fontSize < opFont * 0.65
      ) {
        return [];
      }
      return [{ row, slice }];
    })
    .sort((a, b) => {
      const fontDelta = b.slice.fontSize - a.slice.fontSize;
      if (Math.abs(fontDelta) > opFont * 0.1) return fontDelta;
      return (
        Math.abs(rowCenter(a.row) - opCenterY) -
        Math.abs(rowCenter(b.row) - opCenterY)
      );
    });
  return candidates[0] ? latexForFormulaRow(candidates[0].slice.text) : "";
}

function splitOperatorSlotText(
  text: string,
  operator: LargeOperator,
): { before: string; afterRaw: string; after: string } | null {
  const index = text.indexOf(operator);
  if (index < 0) return null;
  const before = latexForFormulaRow(text.slice(0, index).trim());
  const afterRaw = text.slice(index + operator.length).trim();
  const after = latexForFormulaRow(afterRaw);
  return { before, afterRaw, after };
}

function splitTightIntegralUpperAndBody(
  raw: string,
): { upper: string; body: string } | null {
  const match =
    /^(∞|[+\-−]?\d+(?:\.\d+)?|[A-Za-z])(?:\s+(.+))?$/.exec(raw.trim());
  if (!match) return null;
  return {
    upper: latexForFormulaRow(match[1]),
    body: match[2] ? latexForFormulaRow(match[2].trim()) : "",
  };
}

function latexFromMultiOperatorRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  const candidate = rows
    .map((row) => ({ row, anchors: operatorAnchorsFromRow(row) }))
    .find(({ anchors }) => anchors.length > 1);
  if (!candidate) return null;

  const opFont = rowFont(candidate.row);
  const slots = operatorSlotsFromAnchors(candidate.anchors);
  const parts: string[] = [];

  for (const slot of slots) {
    const opSlice = rowSliceForOperatorSlot(candidate.row, slot, opFont * 0.65);
    if (!opSlice) return null;
    const split = splitOperatorSlotText(opSlice.text, slot.anchor.operator);
    if (!split) return null;

    const lower = nearestOperatorSlotLimitText(
      rows,
      candidate.row,
      slot,
      "below",
      opFont,
    );
    const upperFromAbove = nearestOperatorSlotLimitText(
      rows,
      candidate.row,
      slot,
      "above",
      opFont,
    );
    const bodyFromOtherRow = nearestOperatorSlotBodyText(
      rows,
      candidate.row,
      slot,
      opFont,
    );

    if (slot.anchor.operator === "∫") {
      const tight = !upperFromAbove
        ? splitTightIntegralUpperAndBody(split.afterRaw)
        : null;
      const upper = upperFromAbove || tight?.upper || "";
      const body = tight?.body || split.after || bodyFromOtherRow;
      if (!body) return null;
      const lowerPart = lower ? `_{${lower}}` : "";
      const upperPart = upper ? `^{${upper}}` : "";
      parts.push(`${operatorLatex(slot.anchor.operator)}${lowerPart}${upperPart} ${body}`);
      continue;
    }

    const upper = upperFromAbove || split.before;
    const body = split.after || bodyFromOtherRow;
    if (!body) return null;
    const lowerPart = lower ? `_{${lower}}` : "";
    const upperPart = upper ? `^{${upper}}` : "";
    parts.push(`${operatorLatex(slot.anchor.operator)}${lowerPart}${upperPart} ${body}`);
  }

  return parts.length > 1
    ? normalizeLatexSpacing(joinLatexFormulaParts(parts))
    : null;
}

function nearestRowText(
  rows: readonly PdfFormulaRun[][],
  predicate: (row: PdfFormulaRun[]) => boolean,
): string {
  const row = rows.find(predicate);
  return row ? latexRowText(row) : "";
}

function isOperatorLimitRow(
  row: readonly PdfFormulaRun[],
  opCenter: number,
  opFont: number,
): boolean {
  const left = Math.min(...row.map((run) => run.left));
  const right = Math.max(...row.map(runRight));
  const font = rowFont(row);
  const width = rowWidth(row);
  const center = (left + right) / 2;
  return (
    Math.abs(center - opCenter) <= Math.max(opFont * 1.65, font * 2.5) &&
    font <= opFont * 0.9 &&
    width <= Math.max(font * 5, opFont * 3.2)
  );
}

function operatorBodyRow(
  rows: readonly PdfFormulaRun[][],
  opRow: readonly PdfFormulaRun[],
  opCenter: number,
  opFont: number,
): PdfFormulaRun[] | null {
  const candidates = rows.filter(
    (row) => row !== opRow && !isOperatorLimitRow(row, opCenter, opFont),
  );
  return (
    [...candidates].sort((a, b) => {
      const score = (row: readonly PdfFormulaRun[]) => {
        const text = plainRowText(row);
        return rowFont(row) * 3 + Math.min(text.length, 80);
      };
      return score(b) - score(a);
    })[0] ?? null
  );
}

function latexFromNamedOperatorRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  const candidate = rows
    .map((row) => {
      const text = plainRowText(row);
      const match = namedLimitMatch(text);
      return match ? { row, text, match } : null;
    })
    .find((item): item is NonNullable<typeof item> => Boolean(item));
  if (!candidate) return null;

  const { row: opRow, text: opText, match } = candidate;
  const opCenter = namedOperatorCenterX(opRow, match.index, match.rawLength);
  const opFont = rowFont(opRow);
  const before = latexForFormulaRow(opText.slice(0, match.index).trim());
  const after = latexForFormulaRow(
    opText.slice(match.index + match.rawLength).trim(),
  );
  const belowRows = [...rows]
    .filter((row) => row !== opRow && rowCenter(row) > rowCenter(opRow))
    .sort(
      (a, b) =>
        Math.abs(rowCenter(a) - rowCenter(opRow)) -
        Math.abs(rowCenter(b) - rowCenter(opRow)),
    );
  const lower = nearestRowText(belowRows, (row) =>
    isNamedOperatorLimitRow(row, opCenter, opFont),
  );
  if (!lower) return null;

  return joinLatexFormulaParts([
    before,
    `${match.operator.latex}_{${lower}}`,
    after,
  ]);
}

function splitCaseLatexRow(
  row: string,
): { expression: string; condition: string } | null {
  let expression = "";
  let condition = "";
  const commaIndex = row.indexOf(",");
  if (commaIndex >= 0) {
    expression = row.slice(0, commaIndex).trim();
    condition = row.slice(commaIndex + 1).trim();
  } else {
    const otherwise =
      /^(.+?)\s+(otherwise|否则|其他|其它|\\text\{(?:否则|其他|其它)\})$/i.exec(
        row,
      );
    const conditional =
      /^(.+?)\s+(?:if|when|for|若|当|如果|\\text\{(?:若|当|如果)\})\s+(.+)$/i.exec(
        row,
      );
    if (otherwise) {
      expression = otherwise[1].trim();
      condition = otherwise[2].trim();
    } else if (conditional) {
      expression = conditional[1].trim();
      condition = conditional[2].trim();
    }
  }
  if (!expression || !condition) return null;
  if ((expression.match(/[A-Za-z]{3,}/g) ?? []).length > 1) return null;
  if (
    !/[=<>≤≥≈≠]|\\le|\\ge|\\lt|\\gt|\botherwise\b|^(?:否则|其他|其它)$|\\text\{(?:其他|否则|其它)\}/i.test(
      condition,
    )
  ) {
    return null;
  }
  return { expression, condition };
}

function isMatrixEllipsisLatexCell(cell: string): boolean {
  return /^\\(?:c|v|d)dots$/.test(cell.trim());
}

function isMatrixLatexAtom(cell: string): boolean {
  if (isMatrixEllipsisLatexCell(cell)) return true;
  const base =
    "(?:\\d+(?:\\.\\d+)?|[A-Za-z]|\\\\[A-Za-z]+(?:\\{[A-Za-z0-9]+\\})?)";
  const script =
    "(?:\\^\\{[^{}]+\\}|_\\{[^{}]+\\}|\\^[A-Za-z0-9]+|_[A-Za-z0-9]+)";
  return new RegExp(`^[-+−]?${base}${script}*$`).test(cell);
}

function isMatrixLatexCell(cell: string): boolean {
  if (isMatrixEllipsisLatexCell(cell)) return true;
  if (isMatrixLatexAtom(cell)) return true;
  for (let index = 1; index < cell.length - 1; index += 1) {
    if (!/[+\-−*/×÷]/.test(cell[index] ?? "")) continue;
    const left = cell.slice(0, index).trim();
    const right = cell.slice(index + 1).trim();
    if (isMatrixLatexAtom(left) && isMatrixLatexAtom(right)) return true;
  }
  const parts = cell.split(/\s+/).filter(Boolean);
  if (parts.length !== 3) return false;
  return (
    isMatrixLatexAtom(parts[0]) &&
    /^[+\-−*/×÷]$/.test(parts[1]) &&
    isMatrixLatexAtom(parts[2])
  );
}

function formatMatrixLatexCell(cell: string): string {
  for (let index = 1; index < cell.length - 1; index += 1) {
    const operator = cell[index] ?? "";
    if (!/[+\-−*/×÷]/.test(operator)) continue;
    const left = cell.slice(0, index).trim();
    const right = cell.slice(index + 1).trim();
    if (isMatrixLatexAtom(left) && isMatrixLatexAtom(right)) {
      return `${left} ${operator} ${right}`;
    }
  }
  return cell;
}

function isCasesDelimiterRun(run: PdfFormulaRun): boolean {
  return /^[{}⎧⎨⎩⎪⎫⎬⎭]+$/.test(run.text.trim());
}

function latexFromCasesRows(rows: readonly PdfFormulaRun[][]): string | null {
  if (rows.length < 2) return null;

  let prefix = "";
  const rowLatex = rows.map((row) =>
    latexRowText(row.filter((run) => !isCasesDelimiterRun(run))),
  );
  const normalizedRows = rowLatex
    .map((row) => row.trim())
    .filter((row) => row.length > 0);
  for (let index = 0; index < normalizedRows.length; index += 1) {
    const row = normalizedRows[index];
    const assignment = /^(.+?)=\s*(.+)$/.exec(row);
    if (!assignment) continue;
    if (!splitCaseLatexRow(assignment[2])) continue;
    prefix = `${assignment[1].trim()} = `;
    normalizedRows[index] = assignment[2].trim();
    break;
  }

  const cases = normalizedRows.map(splitCaseLatexRow);
  if (cases.some((row) => row == null)) return null;
  return normalizeLatexSpacing(
    `${prefix}\\begin{cases}${cases
      .map((row) => `${row?.expression} & ${row?.condition}`)
      .join(" \\\\ ")}\\end{cases}`,
  );
}

function matrixCellsFromLatexRow(row: string): string[] | null {
  const cleaned = row
    .replace(/^[\s[(]+/, "")
    .replace(/[\s)\]]+$/, "")
    .trim();
  if (!cleaned || /[,;]|[<>≤≥≈≠]/.test(cleaned)) return null;
  const tokens = cleaned.split(/\s+/).filter(Boolean);
  const cells: string[] = [];
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (
      index + 2 < tokens.length &&
      isMatrixLatexAtom(token) &&
      /^[+\-−*/×÷]$/.test(tokens[index + 1]) &&
      isMatrixLatexAtom(tokens[index + 2])
    ) {
      cells.push(`${token} ${tokens[index + 1]} ${tokens[index + 2]}`);
      index += 2;
      continue;
    }
    cells.push(token);
  }
  if (cells.length < 2) return null;
  if (cells.some((cell) => !isMatrixLatexCell(cell))) {
    return null;
  }
  return cells;
}

function latexFromMatrixRows(rows: readonly PdfFormulaRun[][]): string | null {
  if (rows.length < 2) return null;

  let prefix = "";
  const rowLatex = rows.map((row) => latexRowText(row));
  const normalizedRows = rowLatex.map((row) => row.trim());
  const assignment = /^(.+?)=\s*(.+)$/.exec(normalizedRows[0] ?? "");
  if (assignment) {
    prefix = `${assignment[1].trim()} = `;
    normalizedRows[0] = assignment[2].trim();
  }

  const first = normalizedRows[0] ?? "";
  const last = normalizedRows[normalizedRows.length - 1] ?? "";
  const hasSquareBrackets = /^\[/.test(first) || /\]$/.test(last);
  const hasRoundBrackets = /^\(/.test(first) || /\)$/.test(last);
  if (!hasSquareBrackets && !hasRoundBrackets) return null;

  const rowsAsCells = normalizedRows.map(matrixCellsFromLatexRow);
  if (rowsAsCells.some((row) => row == null)) return null;
  const width = rowsAsCells[0]?.length ?? 0;
  if (width < 2 || rowsAsCells.some((row) => row?.length !== width)) {
    return null;
  }

  const env = hasRoundBrackets && !hasSquareBrackets ? "pmatrix" : "bmatrix";
  return normalizeLatexSpacing(
    `${prefix}\\begin{${env}}${rowsAsCells
      .map((row) => row?.map(formatMatrixLatexCell).join(" & "))
      .join(" \\\\ ")}\\end{${env}}`,
  );
}

function latexFromOperatorRows(
  rows: readonly PdfFormulaRun[][],
): string | null {
  const multiOperatorLatex = latexFromMultiOperatorRows(rows);
  if (multiOperatorLatex) return multiOperatorLatex;

  const candidate = operatorRowCandidate(rows);
  if (!candidate) return null;

  const { row: opRow, operator } = candidate;
  const opText = plainRowText(opRow);
  const opIndex = opText.indexOf(operator);
  if (opIndex < 0) return null;

  const before = latexForFormulaRow(opText.slice(0, opIndex).trim());
  const after = latexForFormulaRow(opText.slice(opIndex + 1).trim());
  const afterRaw = opText.slice(opIndex + 1).trim();
  const opCenter = operatorCenterX(opRow, operator);
  const opFont = rowFont(opRow);
  const bodyRow = operatorBodyRow(rows, opRow, opCenter, opFont);
  const bodyText = bodyRow ? latexRowText(bodyRow) : "";
  const belowRows = [...rows]
    .filter((row) => rowCenter(row) > rowCenter(opRow))
    .sort(
      (a, b) =>
        Math.abs(rowCenter(a) - rowCenter(opRow)) -
        Math.abs(rowCenter(b) - rowCenter(opRow)),
    );
  const aboveRows = [...rows]
    .filter((row) => rowCenter(row) < rowCenter(opRow))
    .sort(
      (a, b) =>
        Math.abs(rowCenter(a) - rowCenter(opRow)) -
        Math.abs(rowCenter(b) - rowCenter(opRow)),
    );
  const lower = nearestRowText(belowRows, (row) => {
    return isOperatorLimitRow(row, opCenter, opFont);
  });
  const upperFromAbove = nearestRowText(aboveRows, (row) => {
    return isOperatorLimitRow(row, opCenter, opFont);
  });

  if (operator === "∫") {
    const afterLooksLikeLimit =
      lower &&
      after &&
      after.length <= 18 &&
      (/^[∞+\-−\d\s.,]+$/.test(afterRaw) || /^\\infty$/.test(after));
    const upper = upperFromAbove || (afterLooksLikeLimit ? after : "");
    const body =
      bodyText ||
      (!afterLooksLikeLimit ? after : latexRowText(dominantBaseRow(rows)));
    const lowerPart = lower ? `_{${lower}}` : "";
    const upperPart = upper ? `^{${upper}}` : "";
    return normalizeLatexSpacing(
      `${operatorLatex(operator)}${lowerPart}${upperPart} ${body}`,
    );
  }

  const afterLooksLikeTightLimit =
    afterRaw.length > 0 &&
    after.length <= 10 &&
    /^[∞+\-−\d\s.,]+$/.test(afterRaw);
  if (afterLooksLikeTightLimit && bodyText.length > 18) {
    return null;
  }
  if (
    after &&
    bodyText.length > 18 &&
    latexContentWeight(bodyText) > latexContentWeight(after) * 1.5
  ) {
    return null;
  }

  const upper = before || upperFromAbove;
  const body = after || bodyText;
  const lowerPart = lower ? `_{${lower}}` : "";
  const upperPart = upper ? `^{${upper}}` : "";
  return normalizeLatexSpacing(
    `${operatorLatex(operator)}${lowerPart}${upperPart} ${body}`,
  );
}

function baseCharacters(row: readonly PdfFormulaRun[]): Array<{
  char: string;
  x: number;
}> {
  const chars: Array<{ char: string; x: number }> = [];
  const sorted = [...row].sort((a, b) => a.left - b.left);
  sorted.forEach((run, runIndex) => {
    const previous = sorted[runIndex - 1];
    if (previous) {
      const gap = run.left - runRight(previous);
      if (gap > Math.max(previous.fontSize, run.fontSize) * 0.25) {
        chars.push({ char: " ", x: run.left - gap / 2 });
      }
    }
    const text = Array.from(run.text);
    const width = Math.max(run.fontSize, runRight(run) - run.left);
    const step = width / Math.max(1, text.length);
    text.forEach((char, index) => {
      chars.push({ char, x: run.left + step * (index + 0.5) });
    });
  });
  return chars;
}

function anchorIndexForScript(
  chars: readonly { char: string; x: number }[],
  scriptCenterX: number,
): number {
  let index = 0;
  let distance = Number.POSITIVE_INFINITY;
  chars.forEach((char, nextIndex) => {
    const nextDistance = Math.abs(char.x - scriptCenterX);
    if (nextDistance < distance) {
      distance = nextDistance;
      index = nextIndex;
    }
  });
  if (/[)\]}]/.test(chars[index]?.char ?? "")) {
    const close = chars[index]?.char;
    const open = close === ")" ? "(" : close === "]" ? "[" : "{";
    let depth = 0;
    for (let i = index; i >= 0; i -= 1) {
      if (chars[i]?.char === close) depth += 1;
      if (chars[i]?.char === open) depth -= 1;
      if (depth === 0) return index;
    }
  }
  while (
    index > 0 &&
    /[\s)\]}.,;:=+\-−<>≤≥≈≠]/.test(chars[index]?.char ?? "")
  ) {
    index -= 1;
  }
  return index;
}

function latexWithPositionedScripts(
  rows: readonly PdfFormulaRun[][],
): string | null {
  const baseRow = dominantBaseRow(rows);
  const baseFont = rowFont(baseRow);
  const baseY = rowCenter(baseRow);
  const chars = baseCharacters(baseRow);
  const insertions = new Map<number, { sup: string[]; sub: string[] }>();
  let consumedScriptRows = 0;

  for (const row of rows) {
    if (row === baseRow) continue;
    const center = rowCenter(row);
    const isScriptRow =
      rowFont(row) <= baseFont * 0.82 ||
      /^[\d+\-−()n\s]+$/.test(plainRowText(row));
    if (!isScriptRow) continue;

    consumedScriptRows += 1;
    for (const run of row) {
      const script = latexForFormulaRow(run.text);
      if (!script) continue;
      const anchor = anchorIndexForScript(chars, runCenterX(run));
      const bucket = insertions.get(anchor) ?? { sup: [], sub: [] };
      if (center < baseY) bucket.sup.push(script);
      else bucket.sub.push(script);
      insertions.set(anchor, bucket);
    }
  }
  if (consumedScriptRows !== rows.length - 1 || insertions.size === 0) {
    return null;
  }

  let raw = "";
  chars.forEach(({ char }, index) => {
    raw += char;
    const bucket = insertions.get(index);
    if (!bucket) return;
    if (bucket.sub.length > 0) raw += `_{${bucket.sub.join(" ")}}`;
    if (bucket.sup.length > 0) raw += `^{${bucket.sup.join(" ")}}`;
  });
  return latexForFormulaRow(raw);
}

/**
 * Best-effort conversion from PDF text extraction's visual formula text into a
 * user-editable LaTeX seed. It intentionally does not pretend to reconstruct a
 * full math AST: uncertain layout is preserved as aligned rows, while reliable
 * symbols/scripts are normalized so KaTeX, AI, and users have a structured
 * starting point instead of flattened OCR-like text.
 */
export function latexFromPdfFormulaText(text: string): string {
  const rows = text
    .split(/\r?\n/g)
    .map((row) => latexForFormulaRow(row))
    .filter(Boolean);

  if (rows.length === 0) return "";
  if (rows.length === 1) return rows[0];
  return `\\begin{aligned}${rows.join(" \\\\ ")}\\end{aligned}`;
}

/**
 * Best-effort readable fallback for edited formula source.
 *
 * `latex` is the semantic source of truth for formulas, but many downstream
 * paths (inline box text, PDF flattening, search/replace, and simple AI text
 * context) still need a plain string. This keeps that string synchronized with
 * source edits instead of leaving stale OCR-like text behind.
 */
export function fallbackTextFromPdfFormulaLatex(latex: string): string {
  if (typeof latex !== "string") return "";
  let text = latex.trim();
  if (!text) return "";

  text = text
    .replace(/^\$\$([\s\S]*)\$\$$/, "$1")
    .replace(/^\$([\s\S]*)\$$/, "$1")
    .replace(/^\\\[([\s\S]*)\\\]$/, "$1")
    .replace(/^\\\(([\s\S]*)\\\)$/, "$1")
    .replace(/\\left|\\right/g, "");

  text = replaceStructuredLatexEnvironments(text);

  text = text
    .replace(/\\begin\{[A-Za-z*]+\}/g, "")
    .replace(/\\end\{[A-Za-z*]+\}/g, "")
    .replace(/\\\\/g, "\n")
    .replace(/&/g, " ");

  for (const command of ["frac", "dfrac", "tfrac"] as const) {
    text = replaceTwoGroupLatexCommand(
      text,
      command,
      (numerator, denominator) =>
        `${groupedFractionPart(numerator)}/${groupedFractionPart(denominator)}`,
    );
  }
  text = replaceSqrtLatexCommand(text);
  text = replaceOneGroupLatexCommand(text, "operatorname", (body) => body);
  text = replaceOneGroupLatexCommand(text, "mathrm", (body) => body);
  text = replaceOneGroupLatexCommand(text, "text", (body) => body);
  text = replaceOneGroupLatexCommand(text, "mathcal", (body) => body);
  text = replaceOneGroupLatexCommand(text, "mathbf", (body) => body);
  text = replaceOneGroupLatexCommand(text, "mathit", (body) => body);
  text = replaceOneGroupLatexCommand(text, "mathsf", (body) => body);
  text = replaceOneGroupLatexCommand(text, "mathtt", (body) => body);
  text = replaceOneGroupLatexCommand(
    text,
    "mathbb",
    (body) => MATHBB_TEXT[body.trim()] ?? body,
  );
  text = replaceOneGroupLatexCommand(text, "tilde", (body) =>
    applyCombiningAccent(body, "\u0303"),
  );
  text = replaceOneGroupLatexCommand(text, "hat", (body) =>
    applyCombiningAccent(body, "\u0302"),
  );
  text = replaceOneGroupLatexCommand(text, "bar", (body) =>
    applyCombiningAccent(body, "\u0304"),
  );
  text = replaceOneGroupLatexCommand(text, "vec", (body) =>
    applyCombiningAccent(body, "\u20d7"),
  );
  text = replaceLatexScripts(text);
  text = text
    .replace(/\\[,;:!]/g, " ")
    .replace(/\\ /g, " ")
    .replace(/\\([{}_$%#&])/g, "$1")
    .replace(
      /\\([A-Za-z]+)(?![A-Za-z])/g,
      (_, command: string) => LATEX_COMMAND_TEXT[command] ?? command,
    )
    .replace(/[{}]/g, "");

  return compactGroupedScriptFallbackSpacing(
    normalizeFormulaFallbackSpacing(text),
  );
}

/**
 * Geometry-aware LaTeX seed for PDF formula fragments. Use this when the caller
 * still has each visual run's position: it can reconstruct common limits and
 * scripts that are impossible to infer from flattened text alone.
 */
export function latexFromPdfFormulaRuns(
  runs: readonly PdfFormulaRun[],
): string {
  const rows = rowsFromFormulaRuns(runs);
  if (rows.length === 0) return "";
  const geometryRoot = latexFromGeometryRootRows(rows);
  if (geometryRoot) return geometryRoot;
  if (rows.length === 1) return latexRowText(rows[0]);

  const alignedFallback = alignedLatexFromRows(rows);
  const matrix = latexFromMatrixRows(rows);
  if (matrix && preservesMostFormulaContent(matrix, alignedFallback)) {
    return matrix;
  }

  const cases = latexFromCasesRows(rows);
  if (cases && preservesMostFormulaContent(cases, alignedFallback)) {
    return cases;
  }

  const visualRoot = latexFromVisualRootRows(rows);
  if (visualRoot) return visualRoot;

  const fraction = latexFromFractionRows(rows);
  if (fraction && preservesMostFormulaContent(fraction, alignedFallback)) {
    return fraction;
  }

  const localFraction = latexFromLocalFractionRows(rows);
  if (
    localFraction &&
    preservesMostFormulaContent(localFraction, alignedFallback)
  ) {
    return localFraction;
  }

  const implicitLocalFraction = latexFromImplicitLocalFractionRows(rows);
  if (
    implicitLocalFraction &&
    preservesMostFormulaContent(implicitLocalFraction, alignedFallback)
  ) {
    return implicitLocalFraction;
  }

  const stackedFraction = latexFromStackedFractionRows(rows);
  if (
    stackedFraction &&
    preservesMostFormulaContent(stackedFraction, alignedFallback)
  ) {
    return stackedFraction;
  }

  const operatorLatexSeed = latexFromOperatorRows(rows);
  if (
    operatorLatexSeed &&
    preservesMostFormulaContent(operatorLatexSeed, alignedFallback)
  ) {
    return operatorLatexSeed;
  }

  const namedOperatorLatexSeed = latexFromNamedOperatorRows(rows);
  if (
    namedOperatorLatexSeed &&
    preservesMostFormulaContent(namedOperatorLatexSeed, alignedFallback)
  ) {
    return namedOperatorLatexSeed;
  }

  const scripted = latexWithPositionedScripts(rows);
  if (
    scripted?.trim() &&
    preservesMostFormulaContent(scripted, alignedFallback)
  ) {
    return scripted;
  }

  return alignedFallback;
}
