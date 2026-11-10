/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * LaTeX -> OMML (Office Math Markup Language), synchronous and pure.
 *
 * Word documents are written by the server (apps/api latex_omml.py is the
 * converter its .docx writer uses). This browser copy remains only because
 * the equation editor's Unicode linear format (lib/unicode-math) converts
 * between LaTeX and UnicodeMath through OMML on every keystroke and cannot
 * wait for a request. It is the converter the retired browser .docx writer
 * used, moved here unchanged; keep it in step with latex_omml.py.
 */

// ---------------------------------------------------------------------------
// LaTeX → Word linear-format (UnicodeMath-ish) translation for math export.
// ---------------------------------------------------------------------------

/** Reverse symbol map: LaTeX control words → Unicode (Word linear format). */
const LATEX_TO_UNICODE: Record<string, string> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  varepsilon: "ε",
  epsilon: "ϵ",
  zeta: "ζ",
  eta: "η",
  theta: "θ",
  vartheta: "ϑ",
  iota: "ι",
  kappa: "κ",
  lambda: "λ",
  mu: "μ",
  nu: "ν",
  xi: "ξ",
  pi: "π",
  varpi: "ϖ",
  rho: "ρ",
  varrho: "ϱ",
  sigma: "σ",
  varsigma: "ς",
  tau: "τ",
  upsilon: "υ",
  varphi: "φ",
  phi: "ϕ",
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
  Upsilon: "Υ",
  Phi: "Φ",
  Psi: "Ψ",
  Omega: "Ω",
  pm: "±",
  mp: "∓",
  times: "×",
  div: "÷",
  cdot: "⋅",
  ast: "∗",
  circ: "∘",
  bullet: "∙",
  oplus: "⊕",
  ominus: "⊖",
  otimes: "⊗",
  odot: "⊙",
  le: "≤",
  leq: "≤",
  ge: "≥",
  geq: "≥",
  ne: "≠",
  neq: "≠",
  approx: "≈",
  equiv: "≡",
  cong: "≅",
  simeq: "≃",
  sim: "∼",
  propto: "∝",
  ll: "≪",
  gg: "≫",
  prec: "≺",
  succ: "≻",
  in: "∈",
  notin: "∉",
  ni: "∋",
  subset: "⊂",
  supset: "⊃",
  subseteq: "⊆",
  supseteq: "⊇",
  cup: "∪",
  cap: "∩",
  setminus: "∖",
  emptyset: "∅",
  varnothing: "∅",
  wedge: "∧",
  vee: "∨",
  neg: "¬",
  forall: "∀",
  exists: "∃",
  nexists: "∄",
  perp: "⊥",
  parallel: "∥",
  vdash: "⊢",
  models: "⊨",
  sqcup: "⊔",
  sqcap: "⊓",
  sum: "∑",
  prod: "∏",
  coprod: "∐",
  int: "∫",
  iint: "∬",
  iiint: "∭",
  oint: "∮",
  bigwedge: "⋀",
  bigvee: "⋁",
  bigcap: "⋂",
  bigcup: "⋃",
  bigodot: "⨀",
  bigoplus: "⨁",
  bigotimes: "⨂",
  to: "→",
  rightarrow: "→",
  leftarrow: "←",
  leftrightarrow: "↔",
  Rightarrow: "⇒",
  Leftarrow: "⇐",
  Leftrightarrow: "⇔",
  mapsto: "↦",
  hookrightarrow: "↪",
  rightharpoonup: "⇀",
  uparrow: "↑",
  downarrow: "↓",
  longrightarrow: "⟶",
  Longrightarrow: "⟹",
  nearrow: "↗",
  searrow: "↘",
  partial: "∂",
  nabla: "∇",
  infty: "∞",
  hbar: "ℏ",
  ell: "ℓ",
  wp: "℘",
  Re: "ℜ",
  Im: "ℑ",
  aleph: "ℵ",
  angle: "∠",
  triangle: "△",
  square: "□",
  blacksquare: "∎",
  dots: "…",
  cdots: "⋯",
  vdots: "⋮",
  ddots: "⋱",
  langle: "〈",
  rangle: "〉",
  Vert: "‖",
  lfloor: "⌊",
  rfloor: "⌋",
  lceil: "⌈",
  rceil: "⌉",
  dagger: "†",
  ddagger: "‡",
  star: "★",
  checkmark: "✓",
  lim: "lim",
  sin: "sin",
  cos: "cos",
  tan: "tan",
  cot: "cot",
  sec: "sec",
  csc: "csc",
  exp: "exp",
  log: "log",
  ln: "ln",
  det: "det",
  dim: "dim",
  ker: "ker",
  arg: "arg",
  gcd: "gcd",
  min: "min",
  max: "max",
  sup: "sup",
  inf: "inf",
  backslash: "\\",
  "%": "%",
  "&": "&",
  "#": "#",
  $: "$",
  _: "_",
  "{": "{",
  "}": "}",
};

const LITERAL_MATH_ACTIVE_CODE_POINTS = new Set([0x005e, 0x007e]);
const LITERAL_MATH_SPACING_CODE_POINTS = new Set([
  0x00a0, 0x2002, 0x2003, 0x2004, 0x2005, 0x2006, 0x2007, 0x2008, 0x2009,
  0x200a, 0x202f, 0x205f, 0x3000,
]);

// ---------------------------------------------------------------------------
// LaTeX → structured OMML (Office Math Markup Language) export.
//
// Word / WPS require OMML for proper math display in .docx. A plain text run
// inside <m:oMath> degrades to UnicodeMath linear (Word can rebuild it), but
// structured OMML (<m:f>, <m:sSup>, <m:rad>, <m:nary>, <m:d>, …) is the
// canonical form that round-trips cleanly through the OMML readers (the
// server's omml_to_latex.py on import, lib/unicode-math/omml-to-latex.ts in
// the equation editor).
//
// STRATEGY: a mini recursive-descent scanner over the LaTeX string that
// recognises the constructs the editor emits and the importer can read back.
// Everything else falls back to a literal <m:r><m:t> text run (never crashes).
// PURE — no DOM, no side effects.
// ---------------------------------------------------------------------------

/**
 * Escape a string for use inside an XML text node or attribute value inside
 * a math run.  Only the five XML metacharacters need escaping.
 */
function escapeOmml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Emit a plain math run (fallback for anything we cannot structure). */
function ommlRun(text: string): string {
  if (!text) return "";
  return `<m:r><m:t xml:space="preserve">${escapeOmml(text)}</m:t></m:r>`;
}

/**
 * LaTeX scanner position: an immutable cursor into a string.  We pass it by
 * value (a number index) and return the new position alongside each result.
 */

/** Skip ASCII whitespace at position i. */
function skipWs(src: string, i: number): number {
  while (
    i < src.length &&
    (src[i] === " " || src[i] === "\t" || src[i] === "\n" || src[i] === "\r")
  )
    i++;
  return i;
}

/**
 * Read a single LaTeX token starting at `i`.  A token is either:
 *   - a `\command` word (backslash + letters), or
 *   - a `\.` escape (backslash + one non-letter character), or
 *   - a single character.
 * Returns {token, end} where end is the index after the token.
 */
function readToken(
  src: string,
  i: number,
): { token: string; end: number } | null {
  i = skipWs(src, i);
  if (i >= src.length) return null;
  if (src[i] === "\\") {
    const j = i + 1;
    if (j >= src.length) return { token: "\\", end: j };
    if (/[A-Za-z]/.test(src[j])) {
      let k = j;
      while (k < src.length && /[A-Za-z]/.test(src[k])) k++;
      // Consume one trailing space after a control word (TeX convention).
      if (k < src.length && src[k] === " ") k++;
      return { token: src.slice(i, k).trimEnd(), end: k };
    }
    // \<non-letter> — single special char escape
    return { token: src.slice(i, j + 1), end: j + 1 };
  }
  return { token: src[i], end: i + 1 };
}

/**
 * Read a balanced `{…}` group starting at `i` (which must point at `{`).
 * Returns {body, end} where body is the content without the outer braces and
 * end is the index after the closing `}`.  Returns null if not a `{` group.
 */
function readGroup(
  src: string,
  i: number,
): { body: string; end: number } | null {
  i = skipWs(src, i);
  if (i >= src.length || src[i] !== "{") return null;
  let depth = 0;
  for (let j = i; j < src.length; j++) {
    if (src[j] === "\\") {
      j++;
      continue;
    }
    if (src[j] === "{") depth++;
    else if (src[j] === "}") {
      depth--;
      if (depth === 0) return { body: src.slice(i + 1, j), end: j + 1 };
    }
  }
  return null;
}

/**
 * Read either a `{…}` group (returns body) or a single token (returns that
 * token).  Used for the argument of `^`, `_`, `\sqrt[…]`, etc.
 */
function readArg(src: string, i: number): { body: string; end: number } | null {
  const g = readGroup(src, i);
  if (g) return g;
  const t = readToken(src, i);
  if (!t) return null;
  return { body: t.token, end: t.end };
}

/**
 * Recognise the exact canonical LaTeX emitted by the OMML importer for an
 * Office skewed fraction. Requiring two braced operands and both explicit
 * negative-space controls keeps an ordinary `a/b` (or a malformed lookalike)
 * on the existing literal/linear path.
 */
function readNativeSkewedFraction(
  src: string,
  i: number,
): { numerator: string; denominator: string; end: number } | null {
  const numerator = readGroup(src, i);
  const marker = "\\!/\\!";
  if (
    !numerator ||
    !src.startsWith(marker, numerator.end) ||
    src[numerator.end + marker.length] !== "{"
  ) {
    return null;
  }
  const denominator = readGroup(src, numerator.end + marker.length);
  if (!denominator) return null;
  return {
    numerator: numerator.body,
    denominator: denominator.body,
    end: denominator.end,
  };
}

/**
 * Recognise the exact canonical LaTeX the OMML importer emits for Word's
 * bar-less fraction (`<m:f><m:fPr><m:type m:val="noBar"/>`) — the binomial
 * coefficient stack Word's own "Binomial Theorem" gallery entry is built from.
 * Only the literal `\genfrac{}{}{0pt}{}{num}{den}` spelling converts; any other
 * `\genfrac` variation keeps the pre-existing literal/linear degradation rather
 * than guessing at a bar thickness or math style Word does not store.
 */
function readNativeBarlessFraction(
  src: string,
  i: number,
): { numerator: string; denominator: string; end: number } | null {
  let pos = i;
  for (const expected of ["", "", "0pt", ""]) {
    const group = readGroup(src, pos);
    if (!group || group.body !== expected) return null;
    pos = group.end;
  }
  const numerator = readGroup(src, pos);
  if (!numerator) return null;
  const denominator = readGroup(src, numerator.end);
  if (!denominator) return null;
  return {
    numerator: numerator.body,
    denominator: denominator.body,
    end: denominator.end,
  };
}

/** Exact canonical form produced by the importer for `<m:sPre>`. */
function readNativePrescript(
  src: string,
  i: number,
): {
  subscript: string;
  superscript: string;
  base: string;
  end: number;
} | null {
  const emptyBase = readGroup(src, i);
  if (!emptyBase || emptyBase.body !== "" || src[emptyBase.end] !== "_") {
    return null;
  }
  const subStart = emptyBase.end + 1;
  if (src[subStart] !== "{") return null;
  const subscript = readGroup(src, subStart);
  if (!subscript || src[subscript.end] !== "^") return null;
  const supStart = subscript.end + 1;
  if (src[supStart] !== "{") return null;
  const superscript = readGroup(src, supStart);
  if (!superscript || src[superscript.end] !== "{") return null;
  const base = readGroup(src, superscript.end);
  if (!base) return null;
  return {
    subscript: subscript.body,
    superscript: superscript.body,
    base: base.body,
    end: base.end,
  };
}

const NATIVE_FUNCTION_APPLY_NAMES = new Set(["sin", "cos", "log"]);

function isNativeFunctionDelimiter(token: string): boolean {
  return Object.prototype.hasOwnProperty.call(DELIM_LATEX_TO_CHAR, token);
}

/** Validate every scalable delimiter in one already-balanced brace group. */
function hasBalancedNativeFunctionDelimiters(src: string): boolean {
  let depth = 0;
  let position = 0;
  while (position < src.length) {
    const token = readToken(src, position);
    if (!token) break;
    position = token.end;
    if (token.token !== "\\left" && token.token !== "\\right") continue;
    const delimiter = readToken(src, position);
    if (!delimiter || !isNativeFunctionDelimiter(delimiter.token)) return false;
    position = delimiter.end;
    if (token.token === "\\left") {
      depth += 1;
    } else {
      if (depth === 0) return false;
      depth -= 1;
    }
  }
  return depth === 0;
}

/** Read the one balanced argument atom accepted by native `m:func`. */
function readNativeFunctionArgument(
  src: string,
  i: number,
): { latex: string; end: number } | null {
  const start = skipWs(src, i);
  const group = readGroup(src, start);
  if (group) {
    if (!hasBalancedNativeFunctionDelimiters(group.body)) return null;
    return { latex: group.body, end: group.end };
  }
  const token = readToken(src, start);
  if (!token || token.token !== "\\left") return null;
  const delimited = readLeftRight(src, token.end);
  if (
    !delimited ||
    !isNativeFunctionDelimiter(delimited.beg) ||
    !isNativeFunctionDelimiter(delimited.end)
  ) {
    return null;
  }
  return { latex: src.slice(start, delimited.endPos), end: delimited.endPos };
}

function nativeFunctionNameOmml(name: string): string {
  return `<m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>${escapeOmml(name)}</m:t></m:r>`;
}

/** LaTeX N-ary operator commands → OMML chr attribute. */
const NARY_LATEX_TO_CHR: Record<string, string> = {
  "\\sum": "∑",
  "\\prod": "∏",
  "\\coprod": "∐",
  "\\int": "∫",
  "\\iint": "∬",
  "\\iiint": "∭",
  "\\oint": "∮",
  "\\bigwedge": "⋀",
  "\\bigvee": "⋁",
  "\\bigcap": "⋂",
  "\\bigcup": "⋃",
  "\\bigodot": "⨀",
  "\\bigoplus": "⨁",
  "\\bigotimes": "⨂",
  "\\biguplus": "⨄",
  "\\bigsqcup": "⨆",
  "\\oiint": "∯",
  "\\oiiint": "∰",
};

/** Delimiter character → LaTeX equivalent for import round-trip. */
const DELIM_LATEX_TO_CHAR: Record<string, string> = {
  "(": "(",
  ")": ")",
  "[": "[",
  "]": "]",
  "\\{": "{",
  "\\}": "}",
  "\\lbrace": "{",
  "\\rbrace": "}",
  "|": "|",
  "\\vert": "|",
  "\\Vert": "‖",
  "\\langle": "〈",
  "\\rangle": "〉",
  "\\lfloor": "⌊",
  "\\rfloor": "⌋",
  "\\lceil": "⌈",
  "\\rceil": "⌉",
  "\\lvert": "|",
  "\\rvert": "|",
  "\\lVert": "‖",
  "\\rVert": "‖",
  "\\|": "‖",
  "\\lbrack": "[",
  "\\rbrack": "]",
  "\\backslash": "\\",
  "/": "/",
  "<": "〈",
  ">": "〉",
  ".": "", // \left. or \right. = invisible delimiter
};

/** Accent commands → OMML <m:acc><m:accPr><m:chr m:val="…"/></m:accPr> chr. */
const ACCENT_LATEX_TO_CHR: Record<string, string> = {
  "\\hat": "̂", // combining circumflex
  "\\check": "̌", // combining caron
  "\\tilde": "̃", // combining tilde
  "\\acute": "́", // combining acute
  "\\grave": "̀", // combining grave
  "\\breve": "̆", // combining breve
  "\\bar": "̅", // combining overline
  "\\vec": "⃗", // combining right arrow above
  "\\dot": "̇", // combining dot above
  "\\ddot": "̈", // combining diaeresis
  "\\overline": "̅", // reuse overline
  "\\widehat": "̂",
  "\\widetilde": "̃",
  "\\mathring": "̊", // combining ring above
  "\\dddot": "⃛", // combining three dots above
};

/**
 * Convert a LaTeX sub/superscript argument string into OMML.
 * Used as the content of <m:sub>, <m:sup>, etc.
 */
function latexArgToOmml(latex: string): string {
  return latexToOmmlInner(latex.trim());
}

// ---------------------------------------------------------------------------
// Native Office Math structures beyond fractions, scripts and radicals.
//
// The shapes follow ECMA-376 Part 1 §22.1 and match the server writer
// (apps/api/app/services/latex_omml.py), so the server OMML reader
// (file_adapters/docx/omml_to_latex.py) reads a browser-exported equation back
// to equivalent LaTeX and Word shows it built up instead of as LaTeX source:
//   - m:m has no delimiters of its own (§22.1.2.60): pmatrix & co. are an m:d
//     fence around the m:m. m:mPr/m:mcs groups adjacent columns of one m:mcJc
//     justification; an array's column spec becomes those justifications.
//   - cases is a brace on one side (m:d begChr "{" endChr "") around a matrix
//     of left-justified columns.
//   - aligned/gathered/split are an m:eqArr with one m:e per row. Inside an
//     equation array an "&" in m:t is Word's alignment mark (§22.1.2.34), so a
//     LaTeX column separator becomes an "&" run and an escaped \& cannot be
//     written faithfully.
//   - over/under braces and stretchy labelled arrows are an m:groupChr
//     (chr/pos/vertJc) with the label as an m:limUpp / m:limLow around it.
//   - \boxed is an m:borderBox; \cancel is a borderless one with a diagonal
//     strike; \phantom is an m:phant with show=0.
//   - math styles are m:rPr run properties: m:sty (b, bi, i, p) for \mathbf,
//     \boldsymbol, \mathit, \mathrm and operator names, m:scr (double-struck,
//     script, fraktur, sans-serif, monospace) for the math alphabets and
//     m:nor for \text.
// Anything outside this subset keeps the literal-run degradation.
// ---------------------------------------------------------------------------

/** Environment bodies above this size stay literal text. */
const OMML_ENVIRONMENT_BODY_LIMIT = 32_768;
/** `m:count` is an ST_Integer255. */
const OMML_MAX_MATRIX_COLUMNS = 255;

/** Symbols the server writer knows that the linear table above lacks. */
const OMML_SYMBOL_SUPPLEMENT: Readonly<Record<string, string>> = {
  omicron: "ο",
  varkappa: "ϰ",
  digamma: "ϝ",
  oslash: "⊘",
  asymp: "≍",
  land: "∧",
  lor: "∨",
  lnot: "¬",
  gets: "←",
  implies: "⇒",
  iff: "⇔",
  longleftarrow: "⟵",
  Longleftarrow: "⟸",
  longleftrightarrow: "⟷",
  Longleftrightarrow: "⟺",
  longmapsto: "⟼",
  hookleftarrow: "↩",
  leftharpoonup: "↼",
  rightleftharpoons: "⇌",
  leadsto: "⇝",
  Uparrow: "⇑",
  Downarrow: "⇓",
  updownarrow: "↕",
  Updownarrow: "⇕",
  swarrow: "↙",
  nwarrow: "↖",
  ldots: "…",
  dotsc: "…",
  dotso: "…",
  dotsb: "⋯",
  dotsm: "⋯",
  dotsi: "⋯",
  prime: "′",
  mid: "∣",
  nmid: "∤",
  nparallel: "∦",
  dashv: "⊣",
  Vdash: "⊩",
  top: "⊤",
  bot: "⊥",
  therefore: "∴",
  because: "∵",
  imath: "ı",
  jmath: "ȷ",
  vert: "|",
  lvert: "|",
  rvert: "|",
  lVert: "‖",
  rVert: "‖",
  "|": "‖",
  lbrace: "{",
  rbrace: "}",
  degree: "°",
  colon: ":",
  coloneqq: "≔",
  eqqcolon: "≕",
  triangleq: "≜",
  doteq: "≐",
  preceq: "⪯",
  succeq: "⪰",
  sqsubseteq: "⊑",
  sqsupseteq: "⊒",
  subsetneq: "⊊",
  supsetneq: "⊋",
  nsubseteq: "⊈",
  nsupseteq: "⊉",
  lesssim: "≲",
  gtrsim: "≳",
  leqslant: "⩽",
  geqslant: "⩾",
  nless: "≮",
  ngtr: "≯",
  nleq: "≰",
  ngeq: "≱",
  nsim: "≁",
  ncong: "≇",
  beth: "ℶ",
  gimel: "ℷ",
  complement: "∁",
  mho: "℧",
  Box: "□",
  diamond: "⋄",
  lozenge: "◊",
  uplus: "⊎",
  amalg: "⨿",
  triangleleft: "◁",
  triangleright: "▷",
  lhd: "⊲",
  rhd: "⊳",
  unlhd: "⊴",
  unrhd: "⊵",
  // Equation palette symbols (equationCatalog.ts) with the palette's glyphs.
  daleth: "ℸ",
  hslash: "ℏ",
  eth: "ð",
  wr: "≀",
  dotplus: "∔",
  boxplus: "⊞",
  boxminus: "⊟",
  boxtimes: "⊠",
  boxdot: "⊡",
  ltimes: "⋉",
  rtimes: "⋊",
  leftthreetimes: "⋋",
  rightthreetimes: "⋌",
  curlyvee: "⋎",
  curlywedge: "⋏",
  circledast: "⊛",
  circledcirc: "⊚",
  circleddash: "⊝",
  lessdot: "⋖",
  gtrdot: "⋗",
  lll: "⋘",
  ggg: "⋙",
  lessapprox: "⪅",
  gtrapprox: "⪆",
  bowtie: "⋈",
  pitchfork: "⋔",
  leftharpoondown: "↽",
  rightharpoondown: "⇁",
  nprec: "⊀",
  nsucc: "⊁",
  nsubset: "⊄",
  nsupset: "⊅",
  nvdash: "⊬",
  nvDash: "⊭",
  rightangle: "∟",
  measuredangle: "∡",
  sphericalangle: "∢",
  blacktriangle: "▲",
  triangledown: "▽",
  blacktriangledown: "▼",
  blacklozenge: "◆",
  bigcirc: "○",
};

/** The Unicode symbol a control word stands for (own keys only). */
function ommlSymbolForCommand(word: string): string | null {
  if (Object.prototype.hasOwnProperty.call(OMML_SYMBOL_SUPPLEMENT, word)) {
    return OMML_SYMBOL_SUPPLEMENT[word];
  }
  if (Object.prototype.hasOwnProperty.call(LATEX_TO_UNICODE, word)) {
    return LATEX_TO_UNICODE[word];
  }
  return null;
}

function ommlGreekCharacter(word: string): string | null {
  const value = ommlSymbolForCommand(word);
  return value !== null && /^[\u0370-\u03ff]$/u.test(value) ? value : null;
}

/** A math run with run properties; xml:space only where the text needs it. */
function ommlPropertyRun(text: string, properties: string): string {
  if (!text) return "";
  const space =
    text !== text.trim() || text.includes("  ") ? ' xml:space="preserve"' : "";
  return `<m:r><m:rPr>${properties}</m:rPr><m:t${space}>${escapeOmml(text)}</m:t></m:r>`;
}

/** A styled math run (m:sty: p upright, b bold, i italic, bi bold italic). */
function ommlStyleRun(text: string, style: "p" | "b" | "i" | "bi"): string {
  return ommlPropertyRun(text, `<m:sty m:val="${style}"/>`);
}

/** Single-character escapes that stand for literal text in a text argument. */
const OMML_TEXT_ESCAPES: Readonly<Record<string, string>> = {
  "\\{": "{",
  "\\}": "}",
  "\\|": "‖",
  "\\$": "$",
  "\\%": "%",
  "\\&": "&",
  "\\#": "#",
  "\\_": "_",
  "\\ ": " ",
  "\\,": " ",
  "\\;": " ",
  "\\:": " ",
  "\\!": "",
  "\\-": "",
};

/**
 * The literal text of a plain argument (`\mathrm{…}`, `\operatorname{…}`,
 * `\text{…}`): characters, escapes and symbol control words. Null when the
 * argument holds structure (groups, scripts, other commands), which callers
 * leave on their existing unstyled path.
 */
function decodeOmmlPlainText(
  body: string,
  preserveSpaces: boolean,
): string | null {
  let out = "";
  let index = 0;
  while (index < body.length) {
    const ch = body[index];
    if (/\s/u.test(ch)) {
      if (preserveSpaces) out += " ";
      index += 1;
      continue;
    }
    if ("{}^_&".includes(ch)) return null;
    if (ch !== "\\") {
      out += ch;
      index += 1;
      continue;
    }
    const next = body[index + 1];
    if (next === undefined) return null;
    if (/[A-Za-z]/.test(next)) {
      let end = index + 1;
      while (end < body.length && /[A-Za-z]/.test(body[end])) end += 1;
      const value = ommlSymbolForCommand(body.slice(index + 1, end));
      if (value === null) return null;
      out += value;
      // TeX swallows the blanks after a control word.
      while (end < body.length && /\s/u.test(body[end])) end += 1;
      index = end;
      continue;
    }
    const escaped = OMML_TEXT_ESCAPES[`\\${next}`];
    if (escaped === undefined) return null;
    out += escaped;
    index += 2;
  }
  return out;
}

/**
 * A style command's argument as LaTeX reads it: a `{group}`, one character
 * (`\mathrm d`) or one Greek letter (`\boldsymbol\theta`).
 */
function readOmmlStyleText(
  src: string,
  i: number,
  allowGreek: boolean,
): { text: string; end: number } | null {
  const start = skipWs(src, i);
  if (start >= src.length) return null;
  if (src[start] === "{") {
    const group = readGroup(src, start);
    if (!group) return null;
    const text = decodeOmmlPlainText(group.body, false);
    return text === null ? null : { text, end: group.end };
  }
  if (src[start] === "\\") {
    if (!allowGreek) return null;
    const token = readToken(src, start);
    const greek = token ? ommlGreekCharacter(token.token.slice(1)) : null;
    return token && greek ? { text: greek, end: token.end } : null;
  }
  if ("}^_&".includes(src[start])) return null;
  const character = String.fromCodePoint(src.codePointAt(start) ?? 0);
  return { text: character, end: start + character.length };
}

/** Math alphabets Word keeps as an m:scr run property (§22.1.2.92). */
const OMML_MATH_ALPHABETS: ReadonlyMap<
  string,
  { readonly script: string; readonly digits: boolean }
> = new Map([
  ["\\mathbb", { script: "double-struck", digits: true }],
  ["\\mathcal", { script: "script", digits: false }],
  ["\\mathscr", { script: "script", digits: false }],
  ["\\mathfrak", { script: "fraktur", digits: false }],
  ["\\mathsf", { script: "sans-serif", digits: true }],
  ["\\mathtt", { script: "monospace", digits: true }],
]);

/** Run styles Word keeps as an m:sty run property. */
const OMML_RUN_STYLES: ReadonlyMap<string, "p" | "b" | "i" | "bi"> = new Map([
  ["\\mathrm", "p"],
  ["\\mathbf", "b"],
  ["\\mathit", "i"],
  ["\\boldsymbol", "bi"],
  ["\\bm", "bi"],
]);

const OMML_TEXT_COMMANDS = new Set([
  "\\text",
  "\\textrm",
  "\\textnormal",
  "\\mbox",
]);

/**
 * Style wrappers whose argument is not plain text (`\mathbf{\frac{a}{b}}`):
 * the contents are written without the style rather than as LaTeX source.
 */
const OMML_PASSTHROUGH_COMMANDS = new Set([
  ...OMML_MATH_ALPHABETS.keys(),
  ...OMML_RUN_STYLES.keys(),
  ...OMML_TEXT_COMMANDS,
  "\\operatorname",
]);

/** Operator names Word sets upright (server writer's function table). */
const OMML_UPRIGHT_FUNCTION_NAMES = new Set([
  "sin", "cos", "tan", "cot", "sec", "csc",
  "arcsin", "arccos", "arctan",
  "sinh", "cosh", "tanh", "coth",
  "log", "ln", "lg", "exp", "det", "gcd", "deg",
  "dim", "hom", "ker", "arg", "Pr", "bmod", "mod",
  "cosec", "sh", "ch", "th", "tg", "ctg", "cth", "cotg", "arctg", "arcctg",
]);

/** Operator names whose `_` subscript is a limit under them (m:limLow). */
const OMML_LIMIT_FUNCTION_NAMES = new Set([
  "lim", "limsup", "liminf", "max", "min", "sup", "inf",
]);

/** Stretchy arrows over or under text: [character, its position, baseline]. */
const OMML_GROUP_ARROWS: ReadonlyMap<
  string,
  readonly [string, "top" | "bot", "top" | "bot"]
> = new Map([
  ["\\overrightarrow", ["→", "top", "bot"]],
  ["\\overleftarrow", ["←", "top", "bot"]],
  ["\\overleftrightarrow", ["↔", "top", "bot"]],
  ["\\underrightarrow", ["→", "bot", "top"]],
  ["\\underleftarrow", ["←", "bot", "top"]],
  ["\\underleftrightarrow", ["↔", "bot", "top"]],
]);

/** Extensible arrows with a label above and/or below (`\xrightarrow[b]{a}`). */
const OMML_X_ARROWS: ReadonlyMap<string, string> = new Map([
  ["\\xrightarrow", "→"],
  ["\\xleftarrow", "←"],
  ["\\xleftrightarrow", "↔"],
  ["\\xRightarrow", "⇒"],
  ["\\xLeftarrow", "⇐"],
  ["\\xLeftrightarrow", "⇔"],
  ["\\xmapsto", "↦"],
]);

/** `\cancel` family: the strikes of a borderless m:borderBox. */
const OMML_CANCELS: ReadonlyMap<string, readonly string[]> = new Map([
  ["\\cancel", ["strikeBLTR"]],
  ["\\bcancel", ["strikeTLBR"]],
  ["\\xcancel", ["strikeBLTR", "strikeTLBR"]],
]);

/** `\phantom` family: m:phantPr flags besides the hidden contents. */
const OMML_PHANTOMS: ReadonlyMap<string, readonly string[]> = new Map([
  ["\\phantom", []],
  ["\\hphantom", ["zeroAsc", "zeroDesc"]],
  ["\\vphantom", ["zeroWid"]],
]);

/** `\not` before a relation: the negated relation's own character. */
const OMML_NEGATED_CHARACTERS: ReadonlyMap<string, string> = new Map([
  ["=", "≠"],
  ["<", "≮"],
  [">", "≯"],
]);
const OMML_NEGATED_COMMANDS: ReadonlyMap<string, string> = new Map([
  ["\\in", "∉"],
  ["\\ni", "∌"],
  ["\\exists", "∄"],
  ["\\equiv", "≢"],
  ["\\approx", "≉"],
  ["\\sim", "≁"],
  ["\\cong", "≇"],
  ["\\mid", "∤"],
  ["\\parallel", "∦"],
  ["\\le", "≰"],
  ["\\leq", "≰"],
  ["\\ge", "≱"],
  ["\\geq", "≱"],
  ["\\subset", "⊄"],
  ["\\supset", "⊅"],
  ["\\subseteq", "⊈"],
  ["\\supseteq", "⊉"],
]);

/** Style switches with no OMML counterpart: the content stands as written. */
const OMML_IGNORED_COMMANDS = new Set([
  "\\displaystyle",
  "\\textstyle",
  "\\nonumber",
  "\\notag",
]);

/** `\bigl(` and kin: a sizing hint; the delimiter stays an ordinary run. */
const OMML_BIG_DELIMITER_COMMANDS = new Set([
  "\\big", "\\Big", "\\bigg", "\\Bigg",
  "\\bigl", "\\bigr", "\\Bigl", "\\Bigr",
  "\\biggl", "\\biggr", "\\Biggl", "\\Biggr",
  "\\bigm", "\\Bigm", "\\biggm", "\\Biggm",
]);

/** env name -> fence characters around the m:m (null: no fence). */
const OMML_MATRIX_FENCES: ReadonlyMap<string, readonly [string, string] | null> =
  new Map([
    ["matrix", null],
    ["smallmatrix", null],
    ["pmatrix", ["(", ")"]],
    ["bmatrix", ["[", "]"]],
    ["Bmatrix", ["{", "}"]],
    ["vmatrix", ["|", "|"]],
    ["Vmatrix", ["‖", "‖"]],
  ]);

/** Cases: a brace on one side of left-justified columns. */
const OMML_CASES_FENCES: ReadonlyMap<string, readonly [string, string]> =
  new Map([
    ["cases", ["{", ""]],
    ["dcases", ["{", ""]],
    ["rcases", ["", "}"]],
    ["drcases", ["", "}"]],
  ]);

const OMML_ALIGNED_ENVIRONMENTS = new Set([
  "aligned", "align", "align*", "split", "alignedat", "alignat", "alignat*",
]);
const OMML_GATHERED_ENVIRONMENTS = new Set(["gathered", "gather", "gather*"]);
const OMML_COUNTED_ENVIRONMENTS = new Set(["alignedat", "alignat", "alignat*"]);
const OMML_EQUATION_ENVIRONMENTS = new Set(["equation", "equation*"]);
const OMML_ARRAY_ENVIRONMENTS = new Set(["array", "darray"]);
const OMML_COLUMN_JUSTIFICATIONS: ReadonlyMap<string, string> = new Map([
  ["l", "left"],
  ["c", "center"],
  ["r", "right"],
]);

/** A delimiter token's OMML character; null for anything else. */
function ommlDelimiterCharacter(token: string): string | null {
  return Object.prototype.hasOwnProperty.call(DELIM_LATEX_TO_CHAR, token)
    ? DELIM_LATEX_TO_CHAR[token]
    : null;
}

/** Read `[…]` (bracket-depth aware) at `i`; null when absent or unbalanced. */
function readOmmlBracketGroup(
  src: string,
  i: number,
): { body: string; end: number } | null {
  const start = skipWs(src, i);
  if (src[start] !== "[") return null;
  let depth = 0;
  for (let index = start; index < src.length; index += 1) {
    if (src[index] === "\\") {
      index += 1;
      continue;
    }
    if (src[index] === "[") depth += 1;
    else if (src[index] === "]") {
      depth -= 1;
      if (depth === 0) return { body: src.slice(start + 1, index), end: index + 1 };
    }
  }
  return null;
}

/**
 * Locate the `\end{name}` matching an environment opened just before `start`,
 * counting nested environments of the same name.
 */
function readOmmlEnvironmentBody(
  src: string,
  start: number,
  name: string,
): { body: string; end: number } | null {
  let depth = 1;
  let position = start;
  while (position < src.length) {
    const backslash = src.indexOf("\\", position);
    if (backslash < 0) return null;
    const token = readToken(src, backslash);
    if (!token) return null;
    if (token.token === "\\begin" || token.token === "\\end") {
      const group = readGroup(src, token.end);
      if (group && group.body.trim() === name) {
        depth += token.token === "\\begin" ? 1 : -1;
        if (depth === 0) {
          return { body: src.slice(start, backslash), end: group.end };
        }
      }
    }
    position = token.end;
  }
  return null;
}

/** Skip `\\[2pt]`-style row spacing, which has no OMML counterpart. */
function skipOmmlRowSpacing(body: string, position: number): number {
  const bracket = readOmmlBracketGroup(body, position);
  if (
    bracket &&
    /^\s*-?(?:\d+\.?\d*|\.\d+)\s*(?:pt|em|ex|mm|cm|in|bp|pc|sp|dd|cc|mu)\s*$/.test(
      bracket.body,
    )
  ) {
    return bracket.end;
  }
  return position;
}

/** Skip an array's `\hline` / `\hdashline`: OMML matrices draw no rules. */
function skipOmmlArrayRules(body: string, position: number): number {
  const rule = /\\h(?:dash)?line(?![A-Za-z])/y;
  let index = position;
  for (;;) {
    rule.lastIndex = skipWs(body, index);
    const match = rule.exec(body);
    if (!match) return index;
    index = rule.lastIndex;
  }
}

/**
 * Split an environment body into rows (`\\`) of cells (`&`) at the top level:
 * outside braces, `\left…\right` pairs and nested environments.
 */
function splitOmmlEnvironmentRows(
  body: string,
  skipRules: boolean,
): string[][] | null {
  if (body.length > OMML_ENVIRONMENT_BODY_LIMIT) return null;
  const rows: string[][] = [];
  let row: string[] = [];
  let braceDepth = 0;
  let delimiterDepth = 0;
  let environmentDepth = 0;
  let index = skipRules ? skipOmmlArrayRules(body, 0) : 0;
  let cellStart = index;
  while (index < body.length) {
    const ch = body[index];
    if (ch === "{") {
      braceDepth += 1;
      index += 1;
      continue;
    }
    if (ch === "}") {
      braceDepth -= 1;
      if (braceDepth < 0) return null;
      index += 1;
      continue;
    }
    const topLevel =
      braceDepth === 0 && delimiterDepth === 0 && environmentDepth === 0;
    if (ch === "&") {
      if (topLevel) {
        row.push(body.slice(cellStart, index).trim());
        cellStart = index + 1;
      }
      index += 1;
      continue;
    }
    if (ch !== "\\") {
      index += 1;
      continue;
    }
    const token = readToken(body, index);
    if (!token) return null;
    if (token.token === "\\left") delimiterDepth += 1;
    else if (token.token === "\\right") {
      delimiterDepth -= 1;
      if (delimiterDepth < 0) return null;
    } else if (token.token === "\\begin") environmentDepth += 1;
    else if (token.token === "\\end") {
      environmentDepth -= 1;
      if (environmentDepth < 0) return null;
    } else if (token.token === "\\\\" && topLevel) {
      row.push(body.slice(cellStart, index).trim());
      rows.push(row);
      row = [];
      let next = skipOmmlRowSpacing(body, token.end);
      if (skipRules) next = skipOmmlArrayRules(body, next);
      index = next;
      cellStart = next;
      continue;
    }
    index = token.end;
  }
  if (braceDepth !== 0 || delimiterDepth !== 0 || environmentDepth !== 0) {
    return null;
  }
  row.push(body.slice(cellStart).trim());
  rows.push(row);
  return rows;
}

/** LaTeX ignores a `\\` right before `\end{…}`. */
function withoutTrailingEmptyOmmlRow(rows: string[][] | null): string[][] | null {
  if (!rows) return null;
  const last = rows[rows.length - 1];
  return rows.length > 1 && last.length === 1 && last[0] === ""
    ? rows.slice(0, -1)
    : rows;
}

/**
 * One environment cell as OMML, or null when the cell holds something this
 * converter leaves as literal LaTeX (a run still spelling a backslash), or —
 * inside an equation array, where "&" is Word's alignment mark — an escaped
 * `\&`. Either way the whole environment stays literal text instead of
 * becoming a native structure Word or the importers would read differently.
 */
function ommlEnvironmentCell(
  cell: string,
  insideEquationArray: boolean,
): string | null {
  const omml = latexArgToOmml(cell);
  for (const text of omml.matchAll(/<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g)) {
    const value = text[1] ?? "";
    if (value.includes("\\")) return null;
    if (insideEquationArray && value.includes("&amp;")) return null;
  }
  return omml;
}

/** An m:m whose m:mcs groups adjacent columns of one justification. */
function ommlMatrix(
  rows: readonly (readonly string[])[],
  justifications: readonly string[],
): string | null {
  if (
    justifications.length < 1 ||
    justifications.length > OMML_MAX_MATRIX_COLUMNS
  ) {
    return null;
  }
  const groups: Array<[string, number]> = [];
  for (const justification of justifications) {
    const last = groups[groups.length - 1];
    if (last && last[0] === justification) last[1] += 1;
    else groups.push([justification, 1]);
  }
  let out = "<m:m><m:mPr><m:mcs>";
  for (const [justification, count] of groups) {
    out +=
      `<m:mc><m:mcPr><m:count m:val="${count}"/>` +
      `<m:mcJc m:val="${justification}"/></m:mcPr></m:mc>`;
  }
  out += "</m:mcs></m:mPr>";
  for (const row of rows) {
    out += "<m:mr>";
    for (const cell of row) {
      const cellOmml = ommlEnvironmentCell(cell, false);
      if (cellOmml === null) return null;
      out += `<m:e>${cellOmml}</m:e>`;
    }
    out += "</m:mr>";
  }
  return `${out}</m:m>`;
}

function ommlFenced(beg: string, end: string, content: string): string {
  return `<m:d>${buildDPr(beg, end)}<m:e>${content}</m:e></m:d>`;
}

function padOmmlRows(rows: readonly string[][], columns: number): string[][] {
  return rows.map((cells) => [
    ...cells,
    ...Array.from({ length: columns - cells.length }, () => ""),
  ]);
}

function ommlRowsAreEmpty(rows: readonly string[][]): boolean {
  return rows.every((cells) => cells.every((cell) => cell === ""));
}

function matrixEnvironmentToOmml(
  body: string,
  fence: readonly [string, string] | null,
): string | null {
  const split = splitOmmlEnvironmentRows(body, false);
  if (!split) return null;
  // A blank template (every cell empty, e.g. the 2×1 \begin{matrix}\\\end{matrix})
  // keeps every row: its trailing \\ is a row of the template, not a typo.
  const blank = ommlRowsAreEmpty(split);
  const rows = blank ? split : withoutTrailingEmptyOmmlRow(split);
  if (!rows) return null;
  const columns = rows[0].length;
  // A matrix is a rectangular grid; a ragged row would have to be guessed
  // into one (and a lone empty one-column cell is no matrix at all).
  if (rows.some((cells) => cells.length !== columns)) return null;
  if (!blank && columns === 1 && rows[rows.length - 1][0] === "") return null;
  const matrix = ommlMatrix(rows, Array(columns).fill("center"));
  if (!matrix) return null;
  return fence ? ommlFenced(fence[0], fence[1], matrix) : matrix;
}

function casesEnvironmentToOmml(
  body: string,
  fence: readonly [string, string],
): string | null {
  const rows = withoutTrailingEmptyOmmlRow(splitOmmlEnvironmentRows(body, false));
  if (!rows || ommlRowsAreEmpty(rows)) return null;
  // A value column and a condition column.
  const columns = Math.max(...rows.map((cells) => cells.length));
  if (columns > 2) return null;
  const matrix = ommlMatrix(
    padOmmlRows(rows, columns),
    Array(columns).fill("left"),
  );
  return matrix ? ommlFenced(fence[0], fence[1], matrix) : null;
}

function equationArrayToOmml(name: string, body: string): string | null {
  let content = body;
  if (OMML_COUNTED_ENVIRONMENTS.has(name)) {
    const count = readGroup(body, 0);
    if (!count || !/^\s*\d+\s*$/.test(count.body)) return null;
    content = body.slice(count.end);
  }
  const rows = withoutTrailingEmptyOmmlRow(
    splitOmmlEnvironmentRows(content, false),
  );
  if (!rows || ommlRowsAreEmpty(rows)) return null;
  const gathered = OMML_GATHERED_ENVIRONMENTS.has(name);
  let out = "<m:eqArr>";
  for (const cells of rows) {
    if (gathered && cells.length > 1) return null;
    const parts: string[] = [];
    for (const cell of cells) {
      const cellOmml = ommlEnvironmentCell(cell, true);
      if (cellOmml === null) return null;
      parts.push(cellOmml);
    }
    // Word's alignment mark between the columns of one row.
    out += `<m:e>${parts.join(ommlRun("&"))}</m:e>`;
  }
  return `${out}</m:eqArr>`;
}

function arrayEnvironmentToOmml(body: string): string | null {
  const spec = readGroup(body, 0);
  if (!spec) return null;
  const justifications: string[] = [];
  for (const character of spec.body) {
    const justification = OMML_COLUMN_JUSTIFICATIONS.get(character);
    if (justification) justifications.push(justification);
    // Vertical rules have no OMML counterpart.
    else if (!/[\s|:]/u.test(character)) return null;
  }
  const rows = withoutTrailingEmptyOmmlRow(
    splitOmmlEnvironmentRows(body.slice(spec.end), true),
  );
  if (!rows || ommlRowsAreEmpty(rows)) return null;
  if (rows.some((cells) => cells.length > justifications.length)) return null;
  return ommlMatrix(padOmmlRows(rows, justifications.length), justifications);
}

/** Native OMML for one `\begin{name}…\end{name}`, or null to keep it literal. */
function environmentToOmml(name: string, body: string): string | null {
  const fence = OMML_MATRIX_FENCES.get(name);
  if (fence !== undefined) return matrixEnvironmentToOmml(body, fence);
  const casesFence = OMML_CASES_FENCES.get(name);
  if (casesFence) return casesEnvironmentToOmml(body, casesFence);
  if (
    OMML_ALIGNED_ENVIRONMENTS.has(name) ||
    OMML_GATHERED_ENVIRONMENTS.has(name)
  ) {
    return equationArrayToOmml(name, body);
  }
  if (OMML_EQUATION_ENVIRONMENTS.has(name)) {
    const rows = splitOmmlEnvironmentRows(body, false);
    if (!rows || rows.length !== 1 || rows[0].length !== 1 || !rows[0][0]) {
      return null;
    }
    return ommlEnvironmentCell(rows[0][0], false);
  }
  if (OMML_ARRAY_ENVIRONMENTS.has(name)) return arrayEnvironmentToOmml(body);
  return null;
}

/**
 * `\left( a \middle| b \right)`: the segments around one repeated separator,
 * which Word keeps as one m:d with an m:sepChr and an m:e per segment. Null
 * when there is no top-level `\middle` or the separators differ.
 */
function splitOmmlMiddleSegments(
  inner: string,
): { separator: string; parts: string[] } | null {
  const parts: string[] = [];
  let separator: string | null = null;
  let depth = 0;
  let position = 0;
  let segmentStart = 0;
  while (position < inner.length) {
    const tokenStart = skipWs(inner, position);
    const token = readToken(inner, position);
    if (!token) break;
    if (token.token === "\\left") depth += 1;
    else if (token.token === "\\right") depth -= 1;
    else if (token.token === "\\middle" && depth === 0) {
      const delimiter = readToken(inner, token.end);
      const character = delimiter ? ommlDelimiterCharacter(delimiter.token) : null;
      if (!delimiter || !character) return null;
      if (separator !== null && character !== separator) return null;
      separator = character;
      parts.push(inner.slice(segmentStart, tokenStart));
      segmentStart = delimiter.end;
      position = delimiter.end;
      continue;
    }
    position = token.end;
  }
  if (separator === null) return null;
  parts.push(inner.slice(segmentStart));
  return { separator, parts };
}

function ommlGroupChr(
  character: string,
  position: "top" | "bot",
  baseline: "top" | "bot",
  body: string,
): string {
  return (
    `<m:groupChr><m:groupChrPr><m:chr m:val="${escapeOmml(character)}"/>` +
    `<m:pos m:val="${position}"/><m:vertJc m:val="${baseline}"/></m:groupChrPr>` +
    `<m:e>${body}</m:e></m:groupChr>`
  );
}

function ommlLimit(
  tag: "limLow" | "limUpp",
  base: string,
  limit: string,
): string {
  return `<m:${tag}><m:e>${base}</m:e><m:lim>${limit}</m:lim></m:${tag}>`;
}

/** `\lim_{x}` / `\operatorname*{argmax}\limits_{x}`: the under-limit argument. */
function readOmmlUnderLimit(
  src: string,
  i: number,
): { latex: string; end: number } | null {
  let position = skipWs(src, i);
  const limits = readToken(src, position);
  if (limits?.token === "\\limits") position = skipWs(src, limits.end);
  if (src[position] !== "_") return null;
  const argument = readArg(src, position + 1);
  return argument ? { latex: argument.body, end: argument.end } : null;
}

function ommlUnderLimited(
  label: string,
  src: string,
  i: number,
): { xml: string; end: number } {
  const under = readOmmlUnderLimit(src, i);
  return under
    ? {
        xml: ommlLimit("limLow", label, latexArgToOmml(under.latex)),
        end: under.end,
      }
    : { xml: label, end: i };
}

function ommlFraction(
  numerator: string,
  denominator: string,
  barless: boolean,
): string {
  const type = barless ? '<m:type m:val="noBar"/>' : "";
  return (
    `<m:f><m:fPr>${type}<m:ctrlPr/></m:fPr>` +
    `<m:num>${latexArgToOmml(numerator)}</m:num>` +
    `<m:den>${latexArgToOmml(denominator)}</m:den></m:f>`
  );
}

/** A `\genfrac` delimiter argument: `{(}`, `{}` or a bare token. */
function readOmmlGenfracDelimiter(
  src: string,
  i: number,
): { character: string; end: number } | null {
  const start = skipWs(src, i);
  if (src[start] === "{") {
    const group = readGroup(src, start);
    if (!group) return null;
    if (!group.body.trim()) return { character: "", end: group.end };
    const token = readToken(group.body, 0);
    if (!token || group.body.slice(token.end).trim()) return null;
    const character = ommlDelimiterCharacter(token.token);
    return character === null ? null : { character, end: group.end };
  }
  const token = readToken(src, start);
  const character = token ? ommlDelimiterCharacter(token.token) : null;
  return token && character !== null ? { character, end: token.end } : null;
}

/** `\genfrac{l}{r}{thickness}{style}{num}{den}`; a zero rule is Word's noBar. */
function readOmmlGenfrac(
  src: string,
  i: number,
): { xml: string; end: number } | null {
  const left = readOmmlGenfracDelimiter(src, i);
  const right = left ? readOmmlGenfracDelimiter(src, left.end) : null;
  const thickness = right ? readGroup(src, right.end) : null;
  const style = thickness ? readGroup(src, thickness.end) : null;
  if (!left || !right || !thickness || !style) return null;
  if (!/^\s*[0-3]?\s*$/.test(style.body)) return null;
  const dimension =
    /^\s*(\d+\.?\d*|\.\d+)\s*(?:pt|em|ex|mm|cm|in|bp|mu)?\s*$/.exec(
      thickness.body,
    );
  if (thickness.body.trim() && !dimension) return null;
  const numerator = readArg(src, style.end);
  const denominator = numerator ? readArg(src, numerator.end) : null;
  if (!numerator || !denominator) return null;
  const fraction = ommlFraction(
    numerator.body,
    denominator.body,
    dimension !== null && Number(dimension[1]) === 0,
  );
  return {
    xml:
      left.character || right.character
        ? ommlFenced(left.character, right.character, fraction)
        : fraction,
    end: denominator.end,
  };
}

/**
 * Native OMML for one control word handled here (styles, operator names,
 * braces, arrows, boxes, …), or null to let the caller's existing handling
 * (symbols, literal runs) take it. `xml` may be "" for a dropped switch.
 */
function nativeCommandToOmml(
  cmd: string,
  src: string,
  i: number,
): { xml: string; end: number } | null {
  const word = cmd.slice(1);

  if (OMML_LIMIT_FUNCTION_NAMES.has(word)) {
    return ommlUnderLimited(ommlStyleRun(word, "p"), src, i);
  }
  if (OMML_UPRIGHT_FUNCTION_NAMES.has(word)) {
    return { xml: ommlStyleRun(word === "bmod" ? "mod" : word, "p"), end: i };
  }

  if (cmd === "\\operatorname") {
    let start = skipWs(src, i);
    const starred = src[start] === "*";
    if (starred) start += 1;
    const group = readGroup(src, start);
    const name = group ? decodeOmmlPlainText(group.body, false) : null;
    if (!group || !name) return null;
    const label = ommlStyleRun(name, "p");
    return starred
      ? ommlUnderLimited(label, src, group.end)
      : { xml: label, end: group.end };
  }

  const alphabet = OMML_MATH_ALPHABETS.get(cmd);
  if (alphabet) {
    const letters = readOmmlStyleText(src, i, false);
    const pattern = alphabet.digits ? /^[A-Za-z0-9]+$/ : /^[A-Za-z]+$/;
    if (!letters || !pattern.test(letters.text)) return null;
    return {
      xml: ommlPropertyRun(
        letters.text,
        `<m:scr m:val="${alphabet.script}"/><m:sty m:val="p"/>`,
      ),
      end: letters.end,
    };
  }

  const style = OMML_RUN_STYLES.get(cmd);
  if (style) {
    if (style === "b" || style === "bi") {
      // Bold script (\boldsymbol{\mathcal{A}}): the script alphabet, bold.
      const group = readGroup(src, i);
      const boldScript = group
        ? /^\s*\\math(?:cal|scr)\s*(?:\{\s*([A-Za-z]+)\s*\}|([A-Za-z]))\s*$/.exec(
            group.body,
          )
        : null;
      if (group && boldScript) {
        return {
          xml: ommlPropertyRun(
            boldScript[1] ?? boldScript[2] ?? "",
            '<m:scr m:val="script"/><m:sty m:val="b"/>',
          ),
          end: group.end,
        };
      }
    }
    const text = readOmmlStyleText(src, i, true);
    // Only letters and digits carry the style; a lone symbol such as the
    // importer's \mathrm{∑} literal-glyph spelling stays a plain run.
    if (!text || !/[\p{L}\p{N}]/u.test(text.text)) return null;
    return { xml: ommlStyleRun(text.text, style), end: text.end };
  }

  if (OMML_TEXT_COMMANDS.has(cmd)) {
    const argument = readArg(src, i);
    const text = argument ? decodeOmmlPlainText(argument.body, true) : null;
    if (!argument || text === null) return null;
    return { xml: ommlPropertyRun(text, "<m:nor/>"), end: argument.end };
  }

  if (cmd === "\\overset" || cmd === "\\underset" || cmd === "\\stackrel") {
    const script = readArg(src, i);
    const base = script ? readArg(src, script.end) : null;
    if (!script || !base) return null;
    return {
      xml: ommlLimit(
        cmd === "\\underset" ? "limLow" : "limUpp",
        latexArgToOmml(base.body),
        latexArgToOmml(script.body),
      ),
      end: base.end,
    };
  }

  if (cmd === "\\overbrace" || cmd === "\\underbrace") {
    const body = readArg(src, i);
    if (!body) return null;
    const over = cmd === "\\overbrace";
    // Word keeps the text on the baseline: an overbrace's (bot) and an
    // underbrace's (top); the label is the limit above or below the brace.
    const group = over
      ? ommlGroupChr("⏞", "top", "bot", latexArgToOmml(body.body))
      : ommlGroupChr("⏟", "bot", "top", latexArgToOmml(body.body));
    const scriptAt = skipWs(src, body.end);
    if (src[scriptAt] === (over ? "^" : "_")) {
      const label = readArg(src, scriptAt + 1);
      if (label) {
        return {
          xml: ommlLimit(
            over ? "limUpp" : "limLow",
            group,
            latexArgToOmml(label.body),
          ),
          end: label.end,
        };
      }
    }
    return { xml: group, end: body.end };
  }

  const groupArrow = OMML_GROUP_ARROWS.get(cmd);
  if (groupArrow) {
    const body = readArg(src, i);
    if (!body) return null;
    return {
      xml: ommlGroupChr(
        groupArrow[0],
        groupArrow[1],
        groupArrow[2],
        latexArgToOmml(body.body),
      ),
      end: body.end,
    };
  }

  const xArrow = OMML_X_ARROWS.get(cmd);
  if (xArrow) {
    let position = i;
    let below = "";
    if (src[skipWs(src, i)] === "[") {
      const bracket = readOmmlBracketGroup(src, i);
      if (!bracket) return null;
      below = bracket.body.trim();
      position = bracket.end;
    }
    const above = readArg(src, position);
    if (!above) return null;
    if (!above.body.trim() && below) {
      // Only a lower label: the arrow above it, both on the baseline.
      return {
        xml: ommlGroupChr(xArrow, "top", "top", latexArgToOmml(below)),
        end: above.end,
      };
    }
    // A stretchy arrow under its label; a lower label as its limit.
    const group = ommlGroupChr(xArrow, "bot", "bot", latexArgToOmml(above.body));
    return {
      xml: below ? ommlLimit("limLow", group, latexArgToOmml(below)) : group,
      end: above.end,
    };
  }

  if (cmd === "\\underline") {
    const body = readArg(src, i);
    if (!body) return null;
    // Word's underbar object (UnicodeMath ▁); m:pos bot is the default.
    return {
      xml: `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${latexArgToOmml(body.body)}</m:e></m:bar>`,
      end: body.end,
    };
  }

  if (cmd === "\\boxed") {
    const body = readArg(src, i);
    if (!body) return null;
    return {
      xml: `<m:borderBox><m:e>${latexArgToOmml(body.body)}</m:e></m:borderBox>`,
      end: body.end,
    };
  }

  const strikes = OMML_CANCELS.get(cmd);
  if (strikes) {
    const body = readArg(src, i);
    if (!body) return null;
    const flags = ["hideTop", "hideBot", "hideLeft", "hideRight", ...strikes]
      .map((flag) => `<m:${flag} m:val="1"/>`)
      .join("");
    return {
      xml:
        `<m:borderBox><m:borderBoxPr>${flags}</m:borderBoxPr>` +
        `<m:e>${latexArgToOmml(body.body)}</m:e></m:borderBox>`,
      end: body.end,
    };
  }

  const phantomFlags = OMML_PHANTOMS.get(cmd);
  if (phantomFlags) {
    const body = readArg(src, i);
    if (!body) return null;
    const flags = phantomFlags.map((flag) => `<m:${flag} m:val="1"/>`).join("");
    return {
      xml:
        `<m:phant><m:phantPr><m:show m:val="0"/>${flags}</m:phantPr>` +
        `<m:e>${latexArgToOmml(body.body)}</m:e></m:phant>`,
      end: body.end,
    };
  }

  if (cmd === "\\binom" || cmd === "\\dbinom" || cmd === "\\tbinom") {
    const top = readArg(src, i);
    const bottom = top ? readArg(src, top.end) : null;
    if (!top || !bottom) return null;
    return {
      xml: ommlFenced("(", ")", ommlFraction(top.body, bottom.body, true)),
      end: bottom.end,
    };
  }

  if (cmd === "\\genfrac") return readOmmlGenfrac(src, i);

  if (cmd === "\\not") {
    const next = readToken(src, i);
    if (!next) return null;
    const negated = next.token.startsWith("\\")
      ? OMML_NEGATED_COMMANDS.get(next.token)
      : OMML_NEGATED_CHARACTERS.get(next.token);
    return negated ? { xml: ommlRun(negated), end: next.end } : null;
  }

  if (cmd === "\\pmod") {
    const argument = readArg(src, i);
    if (!argument) return null;
    return {
      xml:
        ommlRun("(") +
        ommlStyleRun("mod", "p") +
        ommlRun(" ") +
        latexArgToOmml(argument.body) +
        ommlRun(")"),
      end: argument.end,
    };
  }

  if (cmd === "\\ket" || cmd === "\\bra" || cmd === "\\braket") {
    const argument = readArg(src, i);
    if (!argument) return null;
    const [beg, end] =
      cmd === "\\ket" ? ["|", "〉"] : cmd === "\\bra" ? ["〈", "|"] : ["〈", "〉"];
    return {
      xml: ommlFenced(beg, end, latexArgToOmml(argument.body)),
      end: argument.end,
    };
  }

  if (cmd === "\\color") {
    // A colour switch: the formula stands, its colour is not exported.
    const colour = readGroup(src, i);
    return colour ? { xml: "", end: colour.end } : null;
  }
  if (cmd === "\\uline") {
    // Word's run underline (w:u on those runs): the runs stand; this
    // conversion keeps no run formatting, as it keeps no colour.
    const body = readArg(src, i);
    return body ? { xml: latexArgToOmml(body.body), end: body.end } : null;
  }
  if (cmd === "\\textcolor") {
    const colour = readGroup(src, i);
    const body = colour ? readArg(src, colour.end) : null;
    if (!body) return null;
    return { xml: latexArgToOmml(body.body), end: body.end };
  }

  if (OMML_IGNORED_COMMANDS.has(cmd)) return { xml: "", end: i };

  if (OMML_BIG_DELIMITER_COMMANDS.has(cmd)) {
    const delimiter = readToken(src, i);
    const character = delimiter ? ommlDelimiterCharacter(delimiter.token) : null;
    if (!delimiter || character === null) return null;
    return { xml: ommlRun(character), end: delimiter.end };
  }

  if (cmd === "\\substack") {
    // Stacked limits (\sum_{\substack{i<n\\j<m}}): Word keeps them as an
    // equation array in the limit.
    const group = readGroup(src, i);
    const rows = group
      ? withoutTrailingEmptyOmmlRow(splitOmmlEnvironmentRows(group.body, false))
      : null;
    if (!group || !rows || ommlRowsAreEmpty(rows)) return null;
    if (rows.some((cells) => cells.length !== 1)) return null;
    let out = "<m:eqArr>";
    for (const [cell] of rows) {
      const cellOmml = ommlEnvironmentCell(cell, true);
      if (cellOmml === null) return null;
      out += `<m:e>${cellOmml}</m:e>`;
    }
    return { xml: `${out}</m:eqArr>`, end: group.end };
  }

  return null;
}

/**
 * Read a `\left … \right` block starting just after `\left`.  Returns
 * {inner, beg, end, endPos} or null if malformed.
 */
function readLeftRight(
  src: string,
  i: number,
): { inner: string; beg: string; end: string; endPos: number } | null {
  // read the opening delimiter token
  const begTok = readToken(src, i);
  if (!begTok) return null;
  const begRaw = begTok.token;
  let pos = begTok.end;
  // Collect until matching \right (depth-aware for nested \left...\right)
  let depth = 1;
  const start = pos;
  let innerEnd = pos;
  while (pos < src.length) {
    const tok = readToken(src, pos);
    if (!tok) break;
    if (tok.token === "\\left") depth++;
    else if (tok.token === "\\right") {
      depth--;
      if (depth === 0) {
        innerEnd = pos;
        pos = tok.end;
        break;
      }
    }
    pos = tok.end;
  }
  if (depth !== 0) return null; // unmatched
  const inner = src.slice(start, innerEnd);
  // read the closing delimiter token
  const endTok = readToken(src, pos);
  if (!endTok) return null;
  const endRaw = endTok.token;
  return { inner, beg: begRaw, end: endRaw, endPos: endTok.end };
}

/**
 * Convert a delimiter token (e.g. `(`, `\\{`, `\\langle`, `.`) to the
 * character stored in OMML `m:val`.  An empty string means "no delimiter".
 */
function delimChar(tok: string): string {
  return DELIM_LATEX_TO_CHAR[tok] ?? tok;
}

/**
 * Core LaTeX → structured OMML converter.  Returns a string of OMML XML
 * elements (WITHOUT the outer <m:oMath> wrapper) suitable for insertion
 * inside an `<m:oMath>` zone.
 *
 * Uses a two-phase approach:
 *   1. Scan tokens into an array of OMML element strings.
 *   2. A second pass folds `^` and `_` "pending script" markers into the
 *      preceding element as `<m:sSup>`, `<m:sSub>`, or `<m:sSubSup>`.
 *
 * Handles: fractions (including native skew), post- and pre-scripts, radicals,
 * n-ary (sum/prod/int/…), \left…\middle…\right delimiters, accents, matrix /
 * cases / aligned / array environments, braces, stretchy arrows, boxes,
 * phantoms, operator names and math styles (see the native-structure section
 * above). Everything else degrades to an <m:r><m:t> text run — never crashes.
 *
 * PURE.
 *
 * @param latex - Raw LaTeX string (no outer $ delimiters).
 * @returns OMML element string (no outer <m:oMath>).
 */
export function latexToOmmlInner(latex: string): string {
  // We collect OMML "items" — either real element strings or script markers.
  // Script markers are sentinel objects processed in pass 2.
  type OmmlItem =
    | { kind: "elem"; xml: string }
    | { kind: "sup"; arg: string } // ^ arg
    | { kind: "sub"; arg: string }; // _ arg

  const items: OmmlItem[] = [];
  const src = latex;
  let i = 0;

  /** Push an OMML element. */
  const pushElem = (xml: string) => {
    if (xml) items.push({ kind: "elem", xml });
  };

  while (i < src.length) {
    i = skipWs(src, i);
    if (i >= src.length) break;

    const ch = src[i];

    // ── Grouping braces {…} — recurse into contents ──────────────────────
    if (ch === "{") {
      const prescript = readNativePrescript(src, i);
      if (prescript) {
        pushElem(
          `<m:sPre><m:sPrePr><m:ctrlPr/></m:sPrePr><m:sub>${latexArgToOmml(prescript.subscript)}</m:sub><m:sup>${latexArgToOmml(prescript.superscript)}</m:sup><m:e>${latexArgToOmml(prescript.base)}</m:e></m:sPre>`,
        );
        i = prescript.end;
        continue;
      }
      const skewed = readNativeSkewedFraction(src, i);
      if (skewed) {
        pushElem(
          `<m:f><m:fPr><m:type m:val="skw"/><m:ctrlPr/></m:fPr><m:num>${latexArgToOmml(skewed.numerator)}</m:num><m:den>${latexArgToOmml(skewed.denominator)}</m:den></m:f>`,
        );
        i = skewed.end;
        continue;
      }
      const g = readGroup(src, i);
      if (g) {
        pushElem(latexArgToOmml(g.body));
        i = g.end;
        continue;
      }
      pushElem(ommlRun("{"));
      i++;
      continue;
    }

    // ── Superscript ^ ────────────────────────────────────────────────────
    if (ch === "^") {
      i++;
      const arg = readArg(src, i);
      if (!arg) {
        pushElem(ommlRun("^"));
        continue;
      }
      items.push({ kind: "sup", arg: arg.body });
      i = arg.end;
      continue;
    }

    // ── Subscript _ ──────────────────────────────────────────────────────
    if (ch === "_") {
      i++;
      const sub = readArg(src, i);
      if (!sub) {
        pushElem(ommlRun("_"));
        continue;
      }
      items.push({ kind: "sub", arg: sub.body });
      i = sub.end;
      continue;
    }

    // ── Control sequences \… ─────────────────────────────────────────────
    if (ch === "\\") {
      const commandStart = i;
      const tok = readToken(src, i);
      if (!tok) {
        pushElem(ommlRun("\\"));
        i++;
        continue;
      }
      const cmd = tok.token;
      i = tok.end;

      const functionName = cmd.startsWith("\\") ? cmd.slice(1) : cmd;
      if (NATIVE_FUNCTION_APPLY_NAMES.has(functionName)) {
        const argument = readNativeFunctionArgument(src, i);
        if (argument) {
          pushElem(
            `<m:func><m:funcPr><m:ctrlPr/></m:funcPr><m:fName>${nativeFunctionNameOmml(functionName)}</m:fName><m:e>${latexArgToOmml(argument.latex)}</m:e></m:func>`,
          );
          i = argument.end;
          continue;
        }
      }

      // ── \frac{num}{den} ─────────────────────────────────────────────
      if (
        cmd === "\\frac" ||
        cmd === "\\dfrac" ||
        cmd === "\\tfrac" ||
        cmd === "\\cfrac"
      ) {
        const numG = readGroup(src, i);
        if (!numG) {
          pushElem(ommlRun(cmd));
          continue;
        }
        i = numG.end;
        const denG = readGroup(src, i);
        if (!denG) {
          pushElem(ommlRun(cmd + "{" + numG.body + "}"));
          continue;
        }
        i = denG.end;
        pushElem(
          `<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num>${latexArgToOmml(numG.body)}</m:num><m:den>${latexArgToOmml(denG.body)}</m:den></m:f>`,
        );
        continue;
      }

      // ── \genfrac{}{}{0pt}{}{num}{den} — Word's bar-less fraction ─────
      if (cmd === "\\genfrac") {
        const barless = readNativeBarlessFraction(src, i);
        if (barless) {
          pushElem(
            `<m:f><m:fPr><m:type m:val="noBar"/><m:ctrlPr/></m:fPr><m:num>${latexArgToOmml(barless.numerator)}</m:num><m:den>${latexArgToOmml(barless.denominator)}</m:den></m:f>`,
          );
          i = barless.end;
          continue;
        }
      }

      // ── \sqrt[n]{x} or \sqrt{x} ─────────────────────────────────────
      if (cmd === "\\sqrt") {
        let degOmml = "";
        const afterSqrt = skipWs(src, i);
        if (afterSqrt < src.length && src[afterSqrt] === "[") {
          let j = afterSqrt + 1;
          let degEnd = -1;
          let depth = 1;
          while (j < src.length) {
            if (src[j] === "[") depth++;
            else if (src[j] === "]") {
              depth--;
              if (depth === 0) {
                degEnd = j;
                break;
              }
            }
            j++;
          }
          if (degEnd > 0) {
            degOmml = latexArgToOmml(src.slice(afterSqrt + 1, degEnd));
            i = degEnd + 1;
          }
        }
        const radicand = readGroup(src, i);
        if (!radicand) {
          pushElem(ommlRun("√"));
          continue;
        }
        i = radicand.end;
        const radPr = degOmml
          ? `<m:radPr><m:degHide m:val="0"/><m:ctrlPr/></m:radPr><m:deg>${degOmml}</m:deg>`
          : `<m:radPr><m:degHide m:val="1"/><m:ctrlPr/></m:radPr><m:deg/>`;
        pushElem(
          `<m:rad>${radPr}<m:e>${latexArgToOmml(radicand.body)}</m:e></m:rad>`,
        );
        continue;
      }

      // ── N-ary operators: \sum \prod \int \oint … ────────────────────
      if (NARY_LATEX_TO_CHR[cmd] !== undefined) {
        const chr = NARY_LATEX_TO_CHR[cmd];
        let subOmml = "";
        let supOmml = "";
        let j = i;
        // \oint\limits_C / \sum\nolimits_i: an explicit limit placement.
        let placement: "undOvr" | "subSup" | null = null;
        const placementToken = readToken(src, j);
        if (
          placementToken?.token === "\\limits" ||
          placementToken?.token === "\\nolimits"
        ) {
          placement =
            placementToken.token === "\\limits" ? "undOvr" : "subSup";
          j = placementToken.end;
        }
        for (let pass = 0; pass < 2; pass++) {
          const jj = skipWs(src, j);
          if (jj >= src.length) break;
          if (src[jj] === "_" && !subOmml) {
            const arg = readArg(src, jj + 1);
            if (arg) {
              subOmml = latexArgToOmml(arg.body);
              j = arg.end;
              continue;
            }
          }
          if (src[jj] === "^" && !supOmml) {
            const arg = readArg(src, jj + 1);
            if (arg) {
              supOmml = latexArgToOmml(arg.body);
              j = arg.end;
              continue;
            }
          }
          break;
        }
        i = j;
        const limLoc =
          placement ?? (chr === "∑" || chr === "∏" ? "undOvr" : "subSup");
        const naryPr = `<m:naryPr><m:chr m:val="${escapeOmml(chr)}"/><m:limLoc m:val="${limLoc}"/><m:ctrlPr/></m:naryPr>`;
        pushElem(
          `<m:nary>${naryPr}<m:sub>${subOmml}</m:sub><m:sup>${supOmml}</m:sup><m:e></m:e></m:nary>`,
        );
        continue;
      }

      // ── \left … \right delimiters ───────────────────────────────────
      if (cmd === "\\left") {
        const lr = readLeftRight(src, i);
        if (lr) {
          i = lr.endPos;
          const begChar = delimChar(lr.beg);
          const endChar = delimChar(lr.end);
          // \left( a \middle| b \right): one m:d, an m:e per segment and the
          // bar as its m:sepChr. A fenced matrix (\left[\begin{matrix}…) is
          // the same m:d around the environment's m:m.
          const middle = splitOmmlMiddleSegments(lr.inner);
          const dPr = buildDPr(begChar, endChar, middle?.separator);
          const segments = middle ? middle.parts : [lr.inner];
          pushElem(
            `<m:d>${dPr}${segments
              .map((segment) => `<m:e>${latexArgToOmml(segment)}</m:e>`)
              .join("")}</m:d>`,
          );
          continue;
        }
        pushElem(ommlRun("("));
        continue;
      }
      if (cmd === "\\right") {
        pushElem(ommlRun(")"));
        continue;
      }

      // ── \begin{env}…\end{env} environments ──────────────────────────
      if (cmd === "\\begin") {
        const envG = readGroup(src, i);
        const envName = envG?.body.trim() ?? "";
        const environment =
          envG && /^[A-Za-z]+\*?$/.test(envName)
            ? readOmmlEnvironmentBody(src, envG.end, envName)
            : null;
        if (environment) {
          i = environment.end;
          // An environment this converter cannot build natively stays one
          // literal run holding its LaTeX source.
          pushElem(
            environmentToOmml(envName, environment.body) ??
              ommlRun(src.slice(commandStart, i)),
          );
          continue;
        }
        pushElem(ommlRun(src.slice(commandStart)));
        i = src.length;
        continue;
      }

      // ── Accent commands ──────────────────────────────────────────────
      if (ACCENT_LATEX_TO_CHR[cmd] !== undefined) {
        const accentChr = ACCENT_LATEX_TO_CHR[cmd];
        const arg = readArg(src, i);
        if (arg) {
          i = arg.end;
          pushElem(
            `<m:acc><m:accPr><m:chr m:val="${escapeOmml(accentChr)}"/><m:ctrlPr/></m:accPr><m:e>${latexArgToOmml(arg.body)}</m:e></m:acc>`,
          );
          continue;
        }
      }

      // ── Styles, operator names, braces, arrows, boxes, \not, … ───────
      const native = nativeCommandToOmml(cmd, src, i);
      if (native) {
        pushElem(native.xml);
        i = native.end;
        continue;
      }

      // ── Style wrappers whose argument is not plain text: contents only ─
      if (OMML_PASSTHROUGH_COMMANDS.has(cmd)) {
        let argumentStart = i;
        if (cmd === "\\operatorname" && src[skipWs(src, i)] === "*") {
          argumentStart = skipWs(src, i) + 1;
        }
        const arg = readArg(src, argumentStart);
        if (arg) {
          i = arg.end;
          if (OMML_TEXT_COMMANDS.has(cmd)) {
            pushElem(ommlRun(arg.body));
          } else {
            pushElem(latexArgToOmml(arg.body));
          }
          continue;
        }
      }

      // KaTeX and MathLive both preserve this bounded Unicode form. It keeps
      // Office-width spaces distinct instead of collapsing them to the ASCII
      // space emitted by the generic spacing-command branch below.
      if (cmd === "\\unicode") {
        const arg = readArg(src, i);
        const codePointMatch = arg
          ? /^x([0-9a-f]{1,6})$/iu.exec(arg.body)
          : null;
        const codePoint = codePointMatch
          ? Number.parseInt(codePointMatch[1]!, 16)
          : Number.NaN;
        if (
          arg &&
          (LITERAL_MATH_ACTIVE_CODE_POINTS.has(codePoint) ||
            LITERAL_MATH_SPACING_CODE_POINTS.has(codePoint))
        ) {
          i = arg.end;
          pushElem(ommlRun(String.fromCodePoint(codePoint)));
          continue;
        }
      }

      // ── Spacing commands — emit thin space ───────────────────────────
      if (
        cmd === "\\," ||
        cmd === "\\;" ||
        cmd === "\\:" ||
        cmd === "\\!" ||
        cmd === "\\quad" ||
        cmd === "\\qquad" ||
        cmd === "\\thinspace" ||
        cmd === "\\ " ||
        cmd === "\\medspace"
      ) {
        pushElem(ommlRun(" "));
        continue;
      }

      // ── Line break in align / array — skip ──────────────────────────
      if (cmd === "\\\\") {
        pushElem(ommlRun(" "));
        continue;
      }

      // ── \limits — skip (used after \sum etc., already handled above) ─
      if (cmd === "\\limits" || cmd === "\\nolimits") continue;

      // ── All other control words: map to their Unicode symbol ─────────
      // (own keys only: `\constructor` must not reach Object.prototype).
      const cmdWord = cmd.startsWith("\\") ? cmd.slice(1) : cmd;
      const unicode = ommlSymbolForCommand(cmdWord);
      if (unicode !== null) {
        pushElem(ommlRun(unicode));
        continue;
      }
      // Truly unknown — emit the full control word as literal text
      pushElem(ommlRun(cmd));
      continue;
    }

    // ── Plain character ───────────────────────────────────────────────
    pushElem(ommlRun(ch));
    i++;
  }

  // ── Pass 2: fold script markers into preceding elements ─────────────────
  //
  // We scan items left to right. Whenever we see a "sub" or "sup" marker we
  // pop the preceding element and wrap it:
  //   prev + sub               → <m:sSub><m:e>prev</m:e><m:sub>…</m:sub></m:sSub>
  //   prev + sup               → <m:sSup><m:e>prev</m:e><m:sup>…</m:sup></m:sSup>
  //   prev + sub + sup         → <m:sSubSup>…</m:sSubSup>
  //   (nothing) + sub/sup      → <m:sSub/sSup> with empty <m:e/>
  //
  const out: string[] = [];
  for (let k = 0; k < items.length; k++) {
    const item = items[k];
    if (item.kind === "elem") {
      out.push(item.xml);
      continue;
    }
    // Script marker: pop the last element as the base.
    const base = out.length > 0 ? out.pop()! : "";
    const baseE = `<m:e>${base}</m:e>`;

    if (item.kind === "sub") {
      const subOmml = latexArgToOmml(item.arg);
      // Check if the NEXT item is a "sup" (x_i^2 = sub then sup)
      if (k + 1 < items.length && items[k + 1].kind === "sup") {
        const supItem = items[k + 1] as { kind: "sup"; arg: string };
        const supOmml = latexArgToOmml(supItem.arg);
        out.push(
          `<m:sSubSup><m:sSubSupPr><m:ctrlPr/></m:sSubSupPr>${baseE}<m:sub>${subOmml}</m:sub><m:sup>${supOmml}</m:sup></m:sSubSup>`,
        );
        k++; // consume the sup
      } else {
        out.push(
          `<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr>${baseE}<m:sub>${subOmml}</m:sub></m:sSub>`,
        );
      }
    } else {
      // sup
      const supOmml = latexArgToOmml(item.arg);
      // Check if the next item is a "sub" (x^2_i = sup then sub)
      if (k + 1 < items.length && items[k + 1].kind === "sub") {
        const subItem = items[k + 1] as { kind: "sub"; arg: string };
        const subOmml = latexArgToOmml(subItem.arg);
        out.push(
          `<m:sSubSup><m:sSubSupPr><m:ctrlPr/></m:sSubSupPr>${baseE}<m:sub>${subOmml}</m:sub><m:sup>${supOmml}</m:sup></m:sSubSup>`,
        );
        k++; // consume the sub
      } else {
        out.push(
          `<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr>${baseE}<m:sup>${supOmml}</m:sup></m:sSup>`,
        );
      }
    }
  }

  return out.join("");
}

/** Build an <m:dPr> block for a delimiter pair. */
function buildDPr(beg: string, end: string, separator?: string): string {
  const begPart = beg
    ? `<m:begChr m:val="${escapeOmml(beg)}"/>`
    : `<m:begChr m:val=""/>`;
  // CT_DPr order: begChr, sepChr, endChr, …, ctrlPr.
  const separatorPart = separator
    ? `<m:sepChr m:val="${escapeOmml(separator)}"/>`
    : "";
  const endPart = end
    ? `<m:endChr m:val="${escapeOmml(end)}"/>`
    : `<m:endChr m:val=""/>`;
  return `<m:dPr>${begPart}${separatorPart}${endPart}<m:ctrlPr/></m:dPr>`;
}

/**
 * Convert a LaTeX equation string to a complete `<m:oMath>…</m:oMath>` OMML
 * zone string with structured markup.  Exported for unit testing.
 *
 * Unlike the old `latexToWordLinear` (UnicodeMath text) approach, this emits
 * real OMML elements: `<m:f>` fractions, `<m:sSup>`/`<m:sSub>`/`<m:sSubSup>`
 * scripts, `<m:rad>` radicals, `<m:nary>` n-ary operators, `<m:d>` delimiters,
 * `<m:acc>` accents, `<m:m>` matrices.  Anything unrecognised degrades to an
 * `<m:r><m:t>` run — the round-trip importer handles those already.
 *
 * PURE.
 */
export function latexToOmml(latex: string): string {
  const trimmed = latex.trim();
  if (!trimmed) return "";
  const inner = latexToOmmlInner(trimmed);
  if (!inner) return "";
  return `<m:oMath>${inner}</m:oMath>`;
}

