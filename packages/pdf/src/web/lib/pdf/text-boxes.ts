/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Convert pdf.js text content into editable text-box models (WPS-style).
 *
 * pdf.js `getTextContent()` returns items each with a text-space transform
 * matrix; applying the page viewport transform yields CSS-space geometry. We
 * project every item, merge items on the same baseline into a line box, and
 * normalise to page-relative fractions (0..1) so the editable overlay is
 * resolution-independent. Pure + framework-free for unit testing.
 */

import { normalizeBoxAngle, rotatePointAround } from "@mrnote/core/web/lib/pdf/rotated-box.ts";
import {
  fallbackTextFromPdfFormulaLatex,
  latexFromPdfFormulaRuns,
  latexFromPdfFormulaText,
  type PdfFormulaRun,
} from "@mrnote/core/web/lib/pdf/formula-latex.ts";
import { isWeakScalarMetricText } from "./text-statistics";
import { hasFormulaLatex, isFormulaTextBoxKind } from "@mrnote/core/web/lib/pdf/text-box-kind.ts";
import type { PdfTextLineLayout } from "@mrnote/core/web/lib/pdf/text-line-layout.ts";
import {
  mergePdfTextStyleRuns,
  textFromPdfTextStyleRuns,
  type PdfTextStyleRun,
} from "@mrnote/core/web/lib/pdf/text-style-runs.ts";

/**
 * Cache a pure text predicate.
 *
 * Line grouping asks these questions once per candidate PAIR, so the same
 * string is re-tested against every neighbour it might merge with — a
 * quadratic number of calls over a linear number of distinct strings. A
 * 6-page LaTeX paper spent 12.5s of blocked main thread here before this
 * cache, because equation-number detection alone recompiled its regex on
 * every one of those calls.
 *
 * The cache is sound only because every memoized function below is a pure
 * function of its single string argument. It is bounded so that a long
 * session across many documents cannot retain unbounded text.
 */
const TEXT_PREDICATE_CACHE_LIMIT = 4096;

function memoizeTextPredicate<T>(compute: (text: string) => T) {
  const cache = new Map<string, T>();
  return (text: string): T => {
    const hit = cache.get(text);
    if (hit !== undefined || cache.has(text)) return hit as T;
    const value = compute(text);
    if (cache.size >= TEXT_PREDICATE_CACHE_LIMIT) cache.clear();
    cache.set(text, value);
    return value;
  };
}


export type Matrix = [number, number, number, number, number, number];

export interface PdfTextItem {
  str: string;
  /** Text-space transform [a,b,c,d,e,f] from pdf.js. */
  transform: number[];
  /** Item advance width in text space. */
  width: number;
  height?: number;
  /**
   * CSS font-family for this run, resolved by the caller from pdf.js
   * `textContent.styles[item.fontName].fontFamily` — which references the
   * embedded/standard font pdf.js loaded into `document.fonts`, so the editable
   * overlay can render the SAME font as the canvas glyph.
   */
  fontFamily?: string;
  /**
   * Optional live CSS family for the current pdf.js document. This may include
   * pdf.js synthetic font-face names such as `"g_d0_f1"` and must not be
   * persisted; `fontFamily` above is the stable save/export family.
   */
  displayFontFamily?: string;
  /** Original PDF/PostScript font name when pdf.js exposes it. */
  sourceFontName?: string;
  /** True when the run's font is bold / italic (from the pdf.js font object). */
  bold?: boolean;
  italic?: boolean;
  /** Best-effort CSS text colour sampled from the rendered page canvas. */
  color?: string;
  /** Best-effort block classification for editor behaviour and AI context. */
  kind?: "text" | "formula";
  /** Best-effort LaTeX seed for display math extracted from the PDF. */
  latex?: string;
  /** True when the run uses a dedicated math face (e.g. Cambria Math). */
  mathFont?: boolean;
  /**
   * True when the run's extracted characters are raw-charcode fallbacks from a
   * broken ToUnicode CMap (see lib/pdf/math-font-trust.ts): the glyphs render
   * correctly but the text is unreliable and must not seed LaTeX.
   */
  untrustedText?: boolean;
}

export interface PdfViewportLike {
  width: number;
  height: number;
  /** Viewport transform [a,b,c,d,e,f] from page.getViewport(). */
  transform: number[];
}

export interface TextBox {
  id: string;
  /** Page-relative fractions (0..1). */
  x: number;
  y: number;
  width: number;
  height: number;
  /**
   * Clockwise angle of the run on screen, in degrees, about the rect's CENTRE.
   * Absent for the overwhelming majority of text — a rect with no angle is
   * exactly what every box was before this existed.
   */
  angle?: number;
  text: string;
  /** Font size as a fraction of page height. */
  fontSize: number;
  /** CSS font-family of the line (so an edited box matches the canvas glyph). */
  fontFamily?: string;
  /**
   * Live render-only family for the current PDF page. It can include pdf.js
   * synthetic names and is intentionally excluded from persisted edits.
   */
  displayFontFamily?: string;
  /** Original PDF/PostScript font name carried for export-time face matching. */
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
  /**
   * Character runs for mixed-format PDF paragraphs. Their concatenation equals
   * `text`; each run keeps its source face/weight so entering edit mode never
   * repaints the whole paragraph with only the dominant run's style.
   */
  runs?: PdfTextStyleRun[];
  /** Original per-line indentation and vertical placement within this box. */
  lineLayout?: PdfTextLineLayout[];
  /** Best-effort CSS text colour sampled from the rendered page canvas. */
  color?: string;
  /** Best-effort block classification for editor behaviour and AI context. */
  kind?: "text" | "formula";
  /** Best-effort LaTeX seed for display math extracted from the PDF. */
  latex?: string;
  /**
   * True for formula boxes whose source characters are unreliable (broken
   * math-font ToUnicode): no LaTeX seed is derivable from the PDF text and the
   * user (or AI vision) must transcribe the formula from its rendered glyphs.
   */
  transcriptionNeeded?: boolean;
}

/** Multiply two affine matrices (pdf.js Util.transform order: m1 ∘ m2). */
export function multiplyMatrix(m1: number[], m2: number[]): Matrix {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

/**
 * Below this many degrees, text is upright.
 *
 * Extracted matrices carry rounding from the file and from the viewport
 * composition, so an exactly-horizontal baseline routinely measures as a
 * millionth of a degree off. Treating that as a rotation would put an `angle`
 * on every box in every document.
 */
const UPRIGHT_ANGLE_EPSILON = 0.05;

interface ProjectedItem {
  x: number; // CSS px, top-left-ish (baseline x)
  baselineY: number; // CSS px
  /**
   * Clockwise angle of the baseline on screen, in degrees. Zero for the
   * overwhelming majority of text; a watermark or a caption set along a
   * diagram's slope is not zero, and grouping it with upright neighbours by
   * baseline alone produces a line that never existed.
   */
  angle: number;
  /** Distance along the baseline; equals `x` for upright text. */
  along: number;
  /** Distance across the baseline; equals `baselineY` for upright text. */
  across: number;
  fontSize: number; // CSS px
  width: number; // CSS px
  str: string;
  fontFamily?: string;
  displayFontFamily?: string;
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  kind?: "text" | "formula";
  latex?: string;
  mathFont?: boolean;
  untrustedText?: boolean;
}

/** Same run shape as the public model, but fontSize is still CSS px here. */
type ProjectedTextStyleRun = PdfTextStyleRun;

function projectItem(
  item: PdfTextItem,
  viewport: PdfViewportLike,
): ProjectedItem {
  const m = multiplyMatrix(viewport.transform, item.transform);
  const fontSize = Math.hypot(m[2], m[3]) || Math.abs(m[3]) || 1;
  const scaleX = Math.hypot(viewport.transform[0], viewport.transform[1]) || 1;
  // The composed matrix is already in viewport space, where y grows DOWN, so
  // the baseline direction (m[0], m[1]) gives the on-screen clockwise angle
  // directly — the same convention `lib/pdf/rotated-box` works in.
  const rawAngle = normalizeBoxAngle((Math.atan2(m[1], m[0]) * 180) / Math.PI);
  const angle = Math.abs(rawAngle) < UPRIGHT_ANGLE_EPSILON ? 0 : rawAngle;
  // Distances along and across the baseline. For upright text these ARE x and
  // baselineY, so every existing document groups exactly as before; for a
  // turned run they are what "next word" and "same line" actually mean.
  const radians = (angle * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return {
    x: m[4],
    baselineY: m[5],
    angle,
    along: angle === 0 ? m[4] : m[4] * cos + m[5] * sin,
    across: angle === 0 ? m[5] : -m[4] * sin + m[5] * cos,
    fontSize,
    width: Math.max(0, item.width) * scaleX,
    str: item.str,
    fontFamily: item.fontFamily,
    displayFontFamily: item.displayFontFamily,
    sourceFontName: item.sourceFontName,
    bold: item.bold,
    italic: item.italic,
    color: item.color,
    kind: item.kind,
    latex: item.latex,
    mathFont: item.mathFont,
    untrustedText: item.untrustedText,
  };
}

function styleKey(item: {
  fontFamily?: string;
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
}): string {
  return [
    item.fontFamily ?? "",
    item.sourceFontName ?? "",
    item.bold ? "b" : "",
    item.italic ? "i" : "",
  ].join("|");
}

/**
 * Whether two runs read as one style. A run whose source font the reader has
 * not named yet (a font subset still loading, the sibling subset of the same
 * face) is no style change: reportlab and Word draw one CJK line in several
 * subsets of one face, and an unnamed subset beside a named one split a
 * paragraph into six fragments at every digit (PDF quality eval P7).
 */
function sameRunStyle(
  a: { fontFamily?: string; sourceFontName?: string; bold?: boolean; italic?: boolean },
  b: { fontFamily?: string; sourceFontName?: string; bold?: boolean; italic?: boolean },
): boolean {
  if (styleKey(a) === styleKey(b)) return true;
  if (a.sourceFontName && b.sourceFontName) return false;
  return (
    (a.fontFamily ?? "") === (b.fontFamily ?? "") &&
    Boolean(a.bold) === Boolean(b.bold) &&
    Boolean(a.italic) === Boolean(b.italic)
  );
}

function fontSizeRatio(a: number, b: number): number {
  const min = Math.max(0.001, Math.min(a, b));
  return Math.max(a, b) / min;
}

const isDocumentIndexLine = memoizeTextPredicate((text: string): boolean => {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (/(?:\.\s*){3,}\d{1,4}\s*$/.test(trimmed)) return true;
  return false;
});


const FULLWIDTH_OPEN_PAREN = /[（]/g;
const FULLWIDTH_CLOSE_PAREN = /[）]/g;
const FULLWIDTH_OPEN_BRACKET = /[［【]/g;
const FULLWIDTH_CLOSE_BRACKET = /[］】]/g;
const FULLWIDTH_PERIOD = /[．｡]/g;
const FULLWIDTH_COMMA = /[，]/g;
const FULLWIDTH_COLON = /[：]/g;
const FULLWIDTH_SEMICOLON = /[；]/g;
const SPACE_AROUND_PUNCTUATION = /\s*([()[\].,:;-])\s*/g;
const ALL_WHITESPACE = /\s+/g;

const canonicalEquationNumberLineText = memoizeTextPredicate(
  (text: string): string =>
    text
      .trim()
      .normalize("NFKC")
      .replace(FULLWIDTH_OPEN_PAREN, "(")
      .replace(FULLWIDTH_CLOSE_PAREN, ")")
      .replace(FULLWIDTH_OPEN_BRACKET, "[")
      .replace(FULLWIDTH_CLOSE_BRACKET, "]")
      .replace(FULLWIDTH_PERIOD, ".")
      .replace(FULLWIDTH_COMMA, ",")
      .replace(FULLWIDTH_COLON, ":")
      .replace(FULLWIDTH_SEMICOLON, ";")
      .replace(SPACE_AROUND_PUNCTUATION, "$1")
      .replace(ALL_WHITESPACE, ""),
);

const EQUATION_NUMBER_LINE_PATTERN = (() => {
  const suffix = "(?:[a-z])?(?:['′’*★])?";
  const numeric = `\\d+(?:\\.\\d+)*${suffix}`;
  const prefixed = `[A-Z]{1,4}(?:\\.|-)?(?:\\d+(?:[.-]\\d+)*)${suffix}`;
  const cjkPrefixed = `(?:式|公式|方程)\\d+(?:\\.\\d+)*${suffix}`;
  const eqPrefixed = `Eq\\.?\\d+(?:\\.\\d+)*${suffix}`;
  const label = `(?:${numeric}|${prefixed}|${cjkPrefixed}|${eqPrefixed})`;
  return new RegExp(`^(?:\\(${label}\\)|\\[${label}\\])\\s*[.,:;]?$`);
})();

const isEquationNumberLine = memoizeTextPredicate((text: string): boolean =>
  EQUATION_NUMBER_LINE_PATTERN.test(canonicalEquationNumberLineText(text)),
);

const hasStrongMathOperator = memoizeTextPredicate((text: string): boolean => {
  return /[∑∫√∞∂∇∏]/.test(text);
});

const strongMathOperatorCount = memoizeTextPredicate((text: string): number => {
  return (text.match(/[∑∫√∞∂∇∏]/g) ?? []).length;
});

const hasNamedMathOperator = memoizeTextPredicate((text: string): boolean => {
  return /\b(?:lim|arg\s*max|arg\s*min|argmax|argmin|Pr|Cov|Var|det)\b/.test(
    text,
  );
});

const isVisualFractionBarText = memoizeTextPredicate((text: string): boolean => {
  return /^[\-_−–—―─━]+$/.test(text.replace(/\s+/g, ""));
});

const isDocumentIndexLeaderTokenText = memoizeTextPredicate((text: string): boolean => {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact) return false;
  return /^[.\-_−–—―─━]*\d{1,4}[.\-_−–—―─━]*$/.test(compact);
});

const isDocumentIndexLeaderLineText = memoizeTextPredicate((text: string): boolean => {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact) return false;
  return (
    isVisualFractionBarText(compact) || isDocumentIndexLeaderTokenText(compact)
  );
});

const proseWordCount = memoizeTextPredicate((text: string): number => {
  return (text.match(/[A-Za-z]{3,}/g) ?? []).length;
});

const formulaProseWordCount = memoizeTextPredicate((text: string): number => {
  const words = text.match(/[A-Za-z]{3,}/g) ?? [];
  return words.filter(
    (word) =>
      !/^(?:sin|cos|tan|exp|log|ln|max|min|arg|var|lim|det|Pr|Cov|Var)$/i.test(
        word,
      ),
  ).length;
});

const cjkCharCount = memoizeTextPredicate((text: string): number => {
  return (text.match(/[\u3400-\u9fff]/g) ?? []).length;
});

const hasProseText = memoizeTextPredicate((text: string): boolean => {
  return proseWordCount(text) > 0 || cjkCharCount(text) > 0;
});

const MATH_LETTER_CLASS =
  "A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\\u{1d400}-\\u{1d7ff}";
const MATH_ALPHANUMERIC_CLASS = `${MATH_LETTER_CLASS}0-9`;
const UNICODE_SCRIPT_CLASS =
  "₀₁₂₃₄₅₆₇₈₉₊₋₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ";
const MATH_LETTER_RE = new RegExp(`^[${MATH_LETTER_CLASS}]$`, "u");
const MATH_FUNCTION_ATOM_RE = new RegExp(
  `^[${MATH_LETTER_CLASS}]\\([${MATH_ALPHANUMERIC_CLASS},]+\\)$`,
  "u",
);
const MATH_SUBSCRIPT_ATOM_RE = new RegExp(
  `^[${MATH_LETTER_CLASS}]_[${MATH_ALPHANUMERIC_CLASS}]+$`,
  "u",
);
const MATH_SCRIPTED_ATOM_RE = new RegExp(
  `^[${MATH_LETTER_CLASS}](?:[_^][${MATH_ALPHANUMERIC_CLASS}]+|[${UNICODE_SCRIPT_CLASS}])+$`,
  "u",
);

const FORMULA_UNIT_PATTERN =
  "(?:毫米|厘米|千米|公里|米|秒|分钟|小时|分|克|千克|牛|焦|瓦|安|伏|欧|帕|赫|摩尔|开|度|弧度|mm|cm|km|m|s|ms|kg|g|N|J|W|A|V|Ω|Pa|Hz|mol|K)";

const FORMULA_UNIT_RE = new RegExp(
  `^${FORMULA_UNIT_PATTERN}(?:[/／]${FORMULA_UNIT_PATTERN})?$`,
  "i",
);

const isFormulaUnitText = memoizeTextPredicate((text: string): boolean => {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact) return false;
  return FORMULA_UNIT_RE.test(compact);
});

function formulaUnitCompactText(text: string): string | null {
  const compact = text.trim().replace(/\s+/g, "").replace(/／/g, "/");
  return compact && isFormulaUnitText(compact) ? compact.toLowerCase() : null;
}

function formulaUnitSuffix(text: string): string | null {
  const compact = text.trim().replace(/\s+/g, "");
  const units = [
    "分钟",
    "小时",
    "毫米",
    "厘米",
    "千米",
    "公里",
    "千克",
    "摩尔",
    "弧度",
    "mol",
    "ms",
    "kg",
    "mm",
    "cm",
    "km",
    "Pa",
    "Hz",
    "米",
    "秒",
    "分",
    "克",
    "牛",
    "焦",
    "瓦",
    "安",
    "伏",
    "欧",
    "帕",
    "赫",
    "开",
    "度",
    "m",
    "s",
    "g",
    "N",
    "J",
    "W",
    "A",
    "V",
    "Ω",
    "K",
  ];
  return units.find((unit) => compact.endsWith(unit)) ?? null;
}

function isFormulaQuantityBeforeUnitText(text: string): boolean {
  if (/\d/.test(text)) return true;
  if (/[∆Δ∂∇αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}]/u.test(text)) {
    return true;
  }
  if (/^[A-Za-z]{1,2}(?:[_^][A-Za-z0-9]+)?$/.test(text)) return true;
  return /^(?:alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|omega)(?:[_^][A-Za-z0-9]+)?$/i.test(
    text,
  );
}

function isBareAsciiSingleUnitCollision(
  source: string,
  quantity: string,
  unit: string,
): boolean {
  if (!/^[A-Za-z]$/.test(unit)) return false;
  if (!/^[A-Za-z]+$/.test(quantity)) return false;
  const compact = source.trim().replace(/\s+/g, "");
  if (!/^[A-Za-z]+$/.test(compact)) return false;
  // Bare words such as "as", "is", "Am", "gas", and "bus" are more likely
  // prose/OCR fragments than unit expressions. Real unit formulas should carry
  // numeric, Greek, delta, slash, whitespace, or explicit math context.
  return !/[\s/／_^]/.test(source.trim());
}

function isQuantityWithFormulaUnitText(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "");
  const unit = formulaUnitSuffix(compact);
  if (!unit) return false;
  const quantity = compact.slice(0, -unit.length);
  if (!quantity || cjkCharCount(quantity) > 0 || quantity.length > 12) {
    return false;
  }
  if (isBareAsciiSingleUnitCollision(text, quantity, unit)) return false;
  return isFormulaQuantityBeforeUnitText(quantity);
}

function isShortFormulaUnitFragmentText(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "");
  return (
    isQuantityWithFormulaUnitText(text) ||
    (isFormulaUnitText(text) && /[/／]/.test(compact))
  );
}

const mathSignalScore = memoizeTextPredicate((text: string): number => {
  let score = 0;
  score += (text.match(/[=<>≤≥≈≠±∓×÷·→←↔∝]/g) ?? []).length;
  score += (text.match(/[∑∫√∞∂∇∏]/g) ?? []).length * 3;
  score += (text.match(/[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/g) ?? []).length;
  score += (text.match(/[∈∉∀∃∧∨∩∪⊂⊆⊃⊇≡ℝℕℤℚℂ∅]/g) ?? []).length * 2;
  score += (text.match(/[|()[\]{}]/g) ?? []).length > 2 ? 1 : 0;
  if (
    /\b(?:sin|cos|tan|exp|log|ln|max|min|arg|var|lim|det|Pr|Cov|Var)\b/.test(
      text,
    )
  ) {
    score += 2;
  }
  if (hasNamedMathOperator(text)) score += 2;
  if (/\b(?:p|N|E|P)\s*[\[(]/.test(text)) score += 2;
  if (/\b[a-zA-Z]\s*=\s*/.test(text)) score += 1;
  return score;
});

function isCjkProseWithInlineMathText(text: string): boolean {
  const trimmed = text.trim();
  const cjkCount = cjkCharCount(trimmed);
  if (cjkCount < 4) return false;
  const score = mathSignalScore(trimmed);
  const hasCjkSentenceMarker =
    /[，。；：]/.test(trimmed) ||
    /(?:如果|那么|对于|任意|存在|定义|满足|类似|只需|同时|其中|因此|因为|当|若|使|该|由)/.test(
      trimmed,
    );
  if (cjkCount >= 8 && hasCjkSentenceMarker) return true;
  if (score < 8) return true;
  const strongCount = strongMathOperatorCount(trimmed);
  return (
    strongCount >= 2 &&
    cjkCount >= Math.max(6, strongCount * 2 + 4) &&
    cjkCount >= score
  );
}

function isShortCjkNamedOperatorProseFragmentText(text: string): boolean {
  const trimmed = text.trim();
  if (cjkCharCount(trimmed) === 0) return false;
  if (!hasNamedMathOperator(trimmed)) return false;
  if (hasStrongMathOperator(trimmed) || /[=<>≤≥≈≠]/.test(trimmed)) {
    return false;
  }
  return Array.from(trimmed).length <= 24;
}

function isNamedOperatorOnlyFragmentText(text: string): boolean {
  const trimmed = text.trim();
  if (!hasNamedMathOperator(trimmed)) return false;
  if (hasStrongMathOperator(trimmed) || /[=<>≤≥≈≠]/.test(trimmed)) {
    return false;
  }
  const withoutOperators = trimmed
    .replace(
      /\b(?:lim|arg\s*max|arg\s*min|argmax|argmin|Pr|Cov|Var|det)\b/g,
      "",
    )
    .replace(/\b(?:sin|cos|tan|exp|log|ln|max|min|var)\b/gi, "")
    .replace(/[\s()[\]{},.;:，。；：|]+/g, "");
  return withoutOperators === "";
}

function isDanglingNamedOperatorFormulaFragmentText(text: string): boolean {
  const trimmed = text
    .trim()
    .replace(/[,.，。;；:：]+$/g, "")
    .trim();
  if (!trimmed || !hasNamedMathOperator(trimmed)) return false;

  const namedOperator =
    "(?:lim|arg\\s*max|arg\\s*min|argmax|argmin|Pr|Cov|Var|det)";
  const elementaryOperator = "(?:sin|cos|tan|exp|log|ln|max|min|var)";
  const operator = `(?:${namedOperator}|${elementaryOperator})`;
  const comparison = "[=<>≤≥≈≠]";
  const strongSymbol = "[∑∫√∞∂∇∏]";

  if (new RegExp(`^${operator}\\s+${operator}$`, "i").test(trimmed)) {
    return true;
  }
  if (new RegExp(`^${operator}\\s*${comparison}\\s*$`, "i").test(trimmed)) {
    return true;
  }
  if (
    new RegExp(`^${operator}\\s*${comparison}\\s*${strongSymbol}$`, "i").test(
      trimmed,
    )
  ) {
    return true;
  }
  if (new RegExp(`${comparison}\\s*${operator}\\s*$`, "i").test(trimmed)) {
    return true;
  }
  return false;
}

function isSplitNamedOperatorPrefixFormulaText(text: string): boolean {
  const trimmed = formulaBoundaryText(text)
    .replace(/[,.，。;；:：]+$/g, "")
    .trim();
  if (!trimmed || cjkCharCount(trimmed) > 0) return false;
  if (formulaProseWordCount(trimmed) > 0) return false;
  if (/[=<>≤≥≈≠]/.test(trimmed)) return false;
  if (
    startsWithFormulaContinuation(trimmed) ||
    endsWithFormulaContinuation(trimmed)
  ) {
    return false;
  }

  const operator =
    "(?:lim|arg\\s*max|arg\\s*min|argmax|argmin|Pr|Cov|Var|det|sin|cos|tan|exp|log|ln|max|min|var)";
  const tail = `[${MATH_LETTER_CLASS}](?:[_^][${MATH_ALPHANUMERIC_CLASS}]+|[${UNICODE_SCRIPT_CLASS}]+)?`;
  return new RegExp(`^${operator}\\s+${tail}$`, "iu").test(trimmed);
}

function startsWithFormulaArgumentContinuation(text: string): boolean {
  const trimmed = formulaBoundaryText(text);
  if (!trimmed) return false;
  return new RegExp(
    `^[([{]\\s*[${MATH_ALPHANUMERIC_CLASS}αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]`,
    "u",
  ).test(trimmed);
}

function sameRowSplitNamedOperatorLooksLikeFormula(
  previousText: string,
  nextText: string,
): boolean {
  if (!isSplitNamedOperatorPrefixFormulaText(previousText)) return false;
  if (!startsWithFormulaArgumentContinuation(nextText)) return false;
  if (isEquationNumberLine(nextText)) return false;

  const combinedText = joinFormulaTextFragments([previousText, nextText]);
  if (!combinedText || cjkCharCount(combinedText) > 0) return false;
  if (formulaProseWordCount(combinedText) > 0) return false;
  if (isGarbledFormulaExtractionFragmentText(combinedText)) return false;
  return (
    isCompleteRelationFormulaText(combinedText) ||
    hasCompactFormulaStructure(combinedText) ||
    mathSignalScore(combinedText) >= 4
  );
}

function isGarbledFormulaExtractionFragmentText(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (/[#@]/.test(trimmed)) return true;
  if (/^=\s*[,，)\]}]/.test(trimmed)) return true;
  if (/=\s*[,，]\s*\d/.test(trimmed)) return true;
  return false;
}

function isLowConfidenceFormulaFragmentText(text: string): boolean {
  return (
    isShortCjkNamedOperatorProseFragmentText(text) ||
    isNamedOperatorOnlyFragmentText(text) ||
    isDanglingNamedOperatorFormulaFragmentText(text) ||
    isGarbledFormulaExtractionFragmentText(text)
  );
}

function isNonFormulaProseOrFragmentText(text: string): boolean {
  return (
    isEnglishProseWithInlineMathText(text) ||
    isShortProseWithInlineMathText(text) ||
    isCjkProseWithInlineMathText(text) ||
    isLowConfidenceFormulaFragmentText(text)
  );
}

function isEnglishProseWithInlineMathText(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (formulaProseWordCount(trimmed) < 3) return false;
  return (
    mathSignalScore(trimmed) > 0 ||
    /[∑∫√∞∂∇∏αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(trimmed) ||
    /[A-Za-z]\s*=\s*/.test(trimmed)
  );
}

function isShortProseWithInlineMathText(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  const hasInlineMath =
    mathSignalScore(trimmed) > 0 ||
    /[∑∫√∞∂∇∏αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(trimmed) ||
    /[A-Za-z]\s*=\s*/.test(trimmed);
  if (!hasInlineMath) return false;
  if (
    /^(?:where|using|given|let|suppose|assume|when|if|for|from|by|via|because|therefore|thus)\b/i.test(
      trimmed,
    ) ||
    /^according\s+to\b/i.test(trimmed)
  ) {
    return true;
  }
  if (/^(?:求|令|设|若|当|其中|因此|因为|由|根据|利用|通过)/.test(trimmed)) {
    return true;
  }
  return false;
}

const isCompactMathAtom = memoizeTextPredicate((text: string): boolean => {
  const compact = text.trim().replace(/\s+/g, "").normalize("NFKC");
  if (!compact) return false;
  if (/^[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]$/.test(compact)) {
    return true;
  }
  if (MATH_FUNCTION_ATOM_RE.test(compact)) return true;
  if (MATH_SUBSCRIPT_ATOM_RE.test(compact)) return true;
  if (MATH_SCRIPTED_ATOM_RE.test(compact)) return true;
  return false;
});

function isShortAxisLabelText(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "");
  const canonical = compact.normalize("NFKC");
  if (!compact || canonical.length > 12) return false;
  if (/[=<>≤≥≈≠+\-−*/×÷·∑∫√∞∂∇∏]/.test(canonical)) return false;
  if (proseWordCount(canonical) > 0 || cjkCharCount(canonical) > 0) {
    return false;
  }
  if (isBareMathAtomText(canonical)) return true;
  if (MATH_FUNCTION_ATOM_RE.test(canonical)) return true;
  if (MATH_SUBSCRIPT_ATOM_RE.test(canonical)) return true;
  if (
    MATH_SCRIPTED_ATOM_RE.test(compact) ||
    MATH_SCRIPTED_ATOM_RE.test(canonical)
  ) {
    return true;
  }
  return false;
}

function hasCompactFormulaStructure(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "").normalize("NFKC");
  if (!compact) return false;
  return (
    MATH_FUNCTION_ATOM_RE.test(compact) ||
    MATH_SUBSCRIPT_ATOM_RE.test(compact) ||
    MATH_SCRIPTED_ATOM_RE.test(compact)
  );
}

const isBareMathAtomText = memoizeTextPredicate((text: string): boolean => {
  const compact = text.trim().replace(/\s+/g, "").normalize("NFKC");
  return MATH_LETTER_RE.test(compact);
});

function isShortInlineMathAtomSequenceText(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact || compact.length > 12) return false;
  if (proseWordCount(compact) > 0 || cjkCharCount(compact) > 0) return false;
  if (isNumericTickText(compact) || isDocumentIndexLeaderTokenText(compact)) {
    return false;
  }
  if (
    !/[∆Δ∂∇αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}]/u.test(compact)
  ) {
    return false;
  }
  return /^[A-Za-z0-9∆Δ∂∇αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}_^+\-−*/×÷·()[\]{}]+$/u.test(
    compact,
  );
}

function normalizeMatrixBracketGlyphs(text: string): string {
  return text
    .replace(/[⎡⎢⎣]/g, "[")
    .replace(/[⎤⎥⎦]/g, "]")
    .replace(/[⎛⎜⎝]/g, "(")
    .replace(/[⎞⎟⎠]/g, ")");
}

function matrixLikeRowBody(text: string): string {
  return normalizeMatrixBracketGlyphs(text)
    .trim()
    .replace(/^[A-Za-z][A-Za-z0-9_]*\s*=\s*/, "")
    .replace(/^[([{]\s*/, "")
    .replace(/\s*[)\]}]$/, "")
    .trim();
}

const MATRIX_ELLIPSIS_RE = /^(?:⋯|…|⋮|⋱|\.{3})$/;

const { MATRIX_ATOM_RE, MATRIX_BINARY_RE } = (() => {
  const base =
    "(?:\\d+(?:\\.\\d+)?|[A-Za-z]|alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|omega|[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\\u{1d400}-\\u{1d7ff}])";
  const script =
    "(?:_[A-Za-z0-9]+|\\^[A-Za-z0-9]+|[₀₁₂₃₄₅₆₇₈₉₊₋₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ]+)*";
  const atom = `[-+−]?${base}${script}`;
  return {
    MATRIX_ATOM_RE: new RegExp(`^${atom}$`, "u"),
    MATRIX_BINARY_RE: new RegExp(`^${atom}[+\\-−*/×÷·]${atom}$`, "u"),
  };
})();

const isMatrixLikeCellText = memoizeTextPredicate((text: string): boolean => {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact) return false;
  if (MATRIX_ELLIPSIS_RE.test(compact)) return true;
  if (MATRIX_ATOM_RE.test(compact)) return true;
  return MATRIX_BINARY_RE.test(compact);
});

const isMatrixLikeRowText = memoizeTextPredicate((text: string): boolean => {
  const body = matrixLikeRowBody(text);
  if (!body) return false;
  const cells = body.split(/\s+/).filter(Boolean);
  return cells.length >= 2 && cells.every(isMatrixLikeCellText);
});

function matrixLikeRowCells(text: string): string[] {
  const body = matrixLikeRowBody(text);
  if (!body || !isMatrixLikeRowText(text)) return [];
  return body.split(/\s+/).filter(Boolean);
}

function matrixAssignmentLeftHandSide(text: string): string | null {
  return (
    /^([A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}][A-Za-z0-9_αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}]*)\s*=/u.exec(
      normalizeMatrixBracketGlyphs(text).trim(),
    )?.[1] ?? null
  );
}

function matrixAssignmentLeftHandSideLooksMathematical(text: string): boolean {
  const lhs = matrixAssignmentLeftHandSide(text);
  if (!lhs) return false;
  return /^[A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}](?:\d+|_[A-Za-z0-9]+)?$/u.test(
    lhs,
  );
}

function isNonMathematicalMatrixAssignmentRowText(text: string): boolean {
  const trimmed = normalizeMatrixBracketGlyphs(text).trim();
  if (!isMatrixLikeRowText(trimmed)) return false;
  if (!matrixAssignmentLeftHandSide(trimmed)) return false;
  if (/^[([{\s]*[[(]/.test(trimmed) || /[\])\]}]\s*$/.test(trimmed)) {
    return false;
  }
  return !matrixAssignmentLeftHandSideLooksMathematical(trimmed);
}

function isWeakScalarMetricLine(text: string): boolean {
  return isWeakScalarMetricText(text);
}

function isNumericTickText(text: string): boolean {
  return /^[+\-−]?\d+(?:\.\d+)?%?$/.test(text.trim());
}

function isShortPlotLabelText(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact) return false;
  return isNumericTickText(compact) || isShortAxisLabelText(compact);
}

function isLowEvidenceFormulaFragment(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  return isWeakScalarMetricLine(text) || isShortPlotLabelText(text);
}

function isOnlyLowEvidenceFormulaCluster(
  lines: readonly LineBoxInternal[],
): boolean {
  if (!lines.every(isLowEvidenceFormulaFragment)) return false;
  return !linesHavePositionedScriptFormulaGeometry(lines);
}

function latinContinuationSuffix(text: string): string | null {
  const trimmed = text.trim();
  const direct =
    /^([A-Za-z\u00c0-\u024f\u00df-\u00f6\u00f8-\u00ff]{1,5})$/.exec(trimmed);
  if (direct) return direct[1];

  const noisy =
    /^[;:,.，。；：·]+([A-Za-z\u00c0-\u024f\u00df-\u00f6\u00f8-\u00ff]{1,5})$/.exec(
      trimmed,
    );
  return noisy?.[1] ?? null;
}

function isNoisyLatinContinuationSuffix(text: string): boolean {
  return /^[;:,.，。；：·]+[A-Za-z\u00c0-\u024f\u00df-\u00f6\u00f8-\u00ff]{1,5}$/.test(
    text.trim(),
  );
}

function latinContinuationTailLooksLikeSuffix(suffix: string): boolean {
  return /^(?:s|es|ed|er|al|ial|ic|ical|ly|ity|ies|ing|ion|ions|tion|ation|ment)$/i.test(
    suffix,
  );
}

function trailingLatinWord(text: string): string | null {
  return /([A-Za-z\u00c0-\u024f]{4,})$/.exec(text.trimEnd())?.[1] ?? null;
}

function trailingDigitRun(text: string): string | null {
  return /(\d{1,4})$/.exec(text.trimEnd())?.[1] ?? null;
}

function isRomanNumeralContinuationText(text: string): boolean {
  return /^[IVXLCDM]{1,6}$/.test(text.trim());
}

function continuationGapLooksReasonable(
  gap: number,
  fontSize: number,
): boolean {
  return gap >= -fontSize * 0.4 && gap <= fontSize * 3.75;
}

function isAlphabeticWordContinuationWithGap({
  previousText,
  nextText,
  gap,
  fontSize,
  fontSizeRatioValue,
}: {
  previousText: string;
  nextText: string;
  gap: number;
  fontSize: number;
  fontSizeRatioValue: number;
}): boolean {
  const previous = previousText.trimEnd();
  if (!/[A-Za-z\u00c0-\u024f]$/.test(previous)) return false;
  const suffix = latinContinuationSuffix(nextText);
  if (!suffix) return false;
  if (fontSizeRatioValue > 1.35) return false;
  const suffixIsLowercase = /^[a-z\u00df-\u00f6\u00f8-\u00ff]+$/.test(suffix);
  const trailing = trailingLatinWord(previous);
  const suffixLooksMorphological = latinContinuationTailLooksLikeSuffix(suffix);
  if (
    !suffixIsLowercase &&
    !(trailing && trailing.length >= 5 && suffixLooksMorphological)
  ) {
    return false;
  }

  const normalGap = gap >= -fontSize * 0.4 && gap <= fontSize * 0.75;
  if (normalGap) return true;

  const wideGap = gap > fontSize * 0.75 && gap <= fontSize * 3.75;
  if (!wideGap) return false;

  if (!trailing) return false;
  return (
    isNoisyLatinContinuationSuffix(nextText) ||
    latinContinuationTailLooksLikeSuffix(suffix)
  );
}

function isNumericContinuationWithGap({
  previousText,
  nextText,
  gap,
  fontSize,
  fontSizeRatioValue,
}: {
  previousText: string;
  nextText: string;
  gap: number;
  fontSize: number;
  fontSizeRatioValue: number;
}): boolean {
  if (fontSizeRatioValue > 1.35) return false;
  const previousDigits = trailingDigitRun(previousText);
  const nextDigits = /^\d{1,2}$/.exec(nextText.trim())?.[0] ?? null;
  if (!previousDigits || !nextDigits) return false;
  if (previousDigits.length + nextDigits.length > 4) return false;
  return continuationGapLooksReasonable(gap, fontSize);
}

function isCjkContinuationWithGap({
  previousText,
  nextText,
  gap,
  fontSize,
  fontSizeRatioValue,
}: {
  previousText: string;
  nextText: string;
  gap: number;
  fontSize: number;
  fontSizeRatioValue: number;
}): boolean {
  if (fontSizeRatioValue > 1.35) return false;
  if (!/[\u3400-\u9fff]$/.test(previousText.trimEnd())) return false;
  if (!/^[\u3400-\u9fff]{1,3}$/.test(nextText.trim())) return false;
  return continuationGapLooksReasonable(gap, fontSize);
}

function isRomanNumeralContinuationWithGap({
  previousText,
  nextText,
  gap,
  fontSize,
  fontSizeRatioValue,
}: {
  previousText: string;
  nextText: string;
  gap: number;
  fontSize: number;
  fontSizeRatioValue: number;
}): boolean {
  if (fontSizeRatioValue > 1.35) return false;
  const previous = previousText.trimEnd();
  const roman = nextText.trim();
  if (!/[A-Za-z0-9]$/.test(previous)) return false;
  if (!isRomanNumeralContinuationText(roman)) return false;
  if (
    roman.length < 2 &&
    !/\b(?:appendix|chapter|part|section)\s*$/i.test(previous)
  ) {
    return false;
  }
  return continuationGapLooksReasonable(gap, fontSize);
}

function isLooseTitleContinuationWithGap(params: {
  previousText: string;
  nextText: string;
  gap: number;
  fontSize: number;
  fontSizeRatioValue: number;
}): boolean {
  return (
    isNumericContinuationWithGap(params) ||
    isCjkContinuationWithGap(params) ||
    isRomanNumeralContinuationWithGap(params)
  );
}

function isSameTextRunContinuationWithGap(params: {
  previousText: string;
  nextText: string;
  gap: number;
  fontSize: number;
  fontSizeRatioValue: number;
}): boolean {
  return (
    isAlphabeticWordContinuationWithGap(params) ||
    isLooseTitleContinuationWithGap(params)
  );
}

function appendLatinContinuationSuffix(text: string, suffix: string): string {
  if (text.toLocaleLowerCase().endsWith(suffix.toLocaleLowerCase())) {
    return text;
  }
  return `${text}${suffix}`;
}

interface CaseLikeRowParts {
  expression: string;
  condition: string;
}

function parseCaseLikeRowText(text: string): CaseLikeRowParts | null {
  const trimmed = text.trim();
  let expression = "";
  let condition = "";
  const commaIndex = trimmed.indexOf(",");
  if (commaIndex >= 0) {
    expression = trimmed.slice(0, commaIndex).trim();
    condition = trimmed.slice(commaIndex + 1).trim();
  } else {
    const otherwise = /^(.+?)\s+(otherwise|否则|其他|其它)$/i.exec(trimmed);
    const conditional = /^(.+?)\s+(?:if|when|for|若|当|如果)\s+(.+)$/i.exec(
      trimmed,
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
  if (proseWordCount(expression) > 1) return null;
  if (
    !/[=<>≤≥≈≠]|\b(?:otherwise|if|for|when)\b|^(?:否则|其他|其它)$/i.test(
      condition,
    )
  ) {
    return null;
  }
  return { expression, condition };
}

function isCaseLikeRowText(text: string): boolean {
  return parseCaseLikeRowText(text) != null;
}

function isCaseConditionFragmentText(text: string): boolean {
  return parseCaseLikeRowText(`0, ${text.trim()}`) != null;
}

function splitCaseAssignmentFragment(
  text: string,
): { prefix: string; expression: string } | null {
  const match = /^(.+?=\s*)(.*)$/.exec(text.trim());
  if (!match) return null;
  const lhs = match[1].replace(/=\s*$/, "").trim();
  if (!lhs || proseWordCount(lhs) > 1) return null;
  return {
    prefix: `${lhs} =`,
    expression: match[2].trim(),
  };
}

function caseAssignmentPrefixLooksPiecewise(prefix: string): boolean {
  const lhs = prefix.replace(/=\s*$/, "").trim();
  if (!lhs) return false;
  if (/[()[\]{}|]/.test(lhs)) return true;
  if (/[\u{1d400}-\u{1d7ff}αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/u.test(lhs)) {
    return true;
  }
  return /^[A-Za-z]\s*[_^]?\s*[A-Za-z0-9]+$/.test(lhs);
}

function isSplitCaseExpressionFragmentText(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (proseWordCount(trimmed) > 1) return false;
  return hasMathExpressionEvidence(trimmed);
}

function hasMathExpressionEvidence(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact) return false;
  return (
    isCompactMathAtom(text) ||
    isFormulaLikeText(text) ||
    /[=<>≤≥≈≠±∓×÷·∑∫√∞∂∇∏]/.test(text) ||
    /[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(text) ||
    /[⁰¹²³⁴⁵⁶⁷⁸⁹ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ₀₁₂₃₄₅₆₇₈₉]/.test(text) ||
    /^[+\-−]?\d+(?:\.\d+)?$/.test(compact) ||
    /^[A-Za-z]\([A-Za-z0-9,]+\)$/.test(compact)
  );
}

function isStructuralCaseFormulaRowText(text: string): boolean {
  const row = parseCaseLikeRowText(text);
  return row != null && hasMathExpressionEvidence(row.expression);
}

function isStructuralMatrixFormulaRowText(text: string): boolean {
  const trimmed = normalizeMatrixBracketGlyphs(text).trim();
  if (!isMatrixLikeRowText(trimmed)) return false;
  return (
    matrixAssignmentLeftHandSideLooksMathematical(trimmed) ||
    /^[([{\s]*[[(]/.test(trimmed) ||
    /[\])\]}]\s*$/.test(trimmed)
  );
}

function isMatrixLikeNumericRowText(text: string): boolean {
  const body = text
    .replace(/[⎡⎢⎣]/g, "[")
    .replace(/[⎤⎥⎦]/g, "]")
    .replace(/[⎛⎜⎝]/g, "(")
    .replace(/[⎞⎟⎠]/g, ")")
    .trim()
    .replace(/^[A-Za-z][A-Za-z0-9_]*\s*=\s*/, "")
    .replace(/^[([{]\s*/, "")
    .replace(/\s*[)\]}]$/, "")
    .trim();
  if (!body) return false;
  return /^[+\-−]?\d+(?:\.\d+)?(?:\s+[+\-−]?\d+(?:\.\d+)?)+$/.test(body);
}

function isShortTechnicalOrIndexLabelText(text: string): boolean {
  const trimmed = text.trim();
  const compact = trimmed.replace(/\s+/g, "");
  if (!trimmed) return false;
  if (/^[A-Za-z]\+\+$/.test(compact)) return true;
  if (isCompactGradeOrPeriodLabelToken(trimmed)) return true;
  const tokens = trimmed.split(/\s+/).filter(Boolean);
  if (
    tokens.length >= 2 &&
    tokens.length <= 8 &&
    tokens.every(isCompactGradeOrPeriodLabelToken)
  ) {
    return true;
  }
  if (/^[A-Za-z]\s*[-‐‑‒–—−]\s*[A-Za-z]{2,}$/.test(trimmed)) {
    return true;
  }
  if (/^[A-Z]\s*[-‐‑‒–—−]\s*\d+[A-Za-z]?$/.test(trimmed)) {
    return true;
  }
  if (
    /^(?:Appendix|Chapter|Part|Section)\s+[A-Z0-9IVXLCDM.-]+$/i.test(trimmed)
  ) {
    return true;
  }
  return false;
}

function isCompactGradeOrPeriodLabelToken(text: string): boolean {
  const compact = text.trim().replace(/\s+/g, "");
  if (!compact || compact.length > 8) return false;
  if (/^[A-Z]{1,3}\d{0,2}[+＋]$/.test(compact)) return true;
  if (/^[A-Z]{1,3}\d{0,2}[-‐‑‒–—−]$/.test(compact)) return true;
  return /^[A-Z]{1,4}\d{0,2}[+＋\-‐‑‒–—−]\d{1,3}[A-Za-z]?$/.test(compact);
}

function isFormulaLikeText(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed || isEquationNumberLine(trimmed)) return false;
  if (isVisualFractionBarText(trimmed)) return false;
  if (isDocumentIndexLeaderTokenText(trimmed)) return false;
  if (isShortTechnicalOrIndexLabelText(trimmed)) return false;
  if (isShortPlotLabelText(trimmed)) return false;
  if (isClearlyDanglingFormulaSourceText(trimmed)) {
    return false;
  }
  const score = mathSignalScore(trimmed);
  const words = proseWordCount(trimmed);
  const cjkCount = cjkCharCount(trimmed);
  if (isNonFormulaProseOrFragmentText(trimmed)) return false;
  if (isShortFormulaUnitFragmentText(trimmed) && /[=<>≤≥≈≠]/.test(trimmed)) {
    return true;
  }
  if (cjkCount >= 2 && score < 4 && !hasStrongMathOperator(trimmed)) {
    return false;
  }
  const strongOperator = hasStrongMathOperator(trimmed);
  const namedOperator = hasNamedMathOperator(trimmed);
  if (words >= 5 && !strongOperator) return false;
  if (words >= 4 && !strongOperator && !/[=<>≤≥≈≠]/.test(trimmed)) {
    return false;
  }
  if (strongOperator) {
    return words < 3 || (score >= 6 && words <= 4);
  }
  if (isCompactMathAtom(trimmed)) return true;
  if (namedOperator && score >= 3 && words <= 4) return true;
  if (/^[A-Za-z]\s*[\[(]\s*[A-Za-z0-9]+\s*[\])]\s*=/.test(trimmed)) {
    return true;
  }
  if (trimmed.length <= 8 && /[=+\-]/.test(trimmed) && words <= 1) {
    return true;
  }
  if (score >= 4 && words <= 8) return true;
  if (score >= 3 && words <= 3) return true;
  return false;
}

function isCjkInlineMathProseContinuation({
  previousText,
  nextText,
  gap,
  fontSize,
}: {
  previousText: string;
  nextText: string;
  gap: number;
  fontSize: number;
}): boolean {
  if (gap > fontSize * 0.8) return false;
  const combined = `${previousText} ${nextText}`;
  if (cjkCharCount(combined) < 2) return false;
  return (
    mathSignalScore(combined) > 0 ||
    /[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(combined) ||
    /[A-Za-z]\s*[\[(]\s*[A-Za-z0-9]+\s*[\])]\s*=/.test(combined)
  );
}

function shouldStartNewInlineSegment(
  previous: ProjectedItem,
  next: ProjectedItem,
): boolean {
  const previousEnd = previous.x + previous.width;
  const gap = next.x - previousEnd;
  const fontSize = Math.max(previous.fontSize, next.fontSize);
  const ratio = fontSizeRatio(previous.fontSize, next.fontSize);
  const textRunContinuation = isSameTextRunContinuationWithGap({
    previousText: previous.str,
    nextText: next.str,
    gap,
    fontSize,
    fontSizeRatioValue: ratio,
  });
  const previousFormula =
    previous.kind === "formula" || Boolean(previous.latex);
  const nextFormula = next.kind === "formula" || Boolean(next.latex);
  const proseLeadContinuation =
    isInlineProseLeadFragmentText(previous.str) &&
    (proseWordCount(next.str) > 0 || cjkCharCount(next.str) > 0) &&
    gap >= -fontSize * 0.4 &&
    gap <= fontSize * 1.2;

  if (gap > fontSize * 1.75 && !textRunContinuation) return true;
  if (
    previousFormula !== nextFormula &&
    !isCjkInlineMathProseContinuation({
      previousText: previous.str,
      nextText: next.str,
      gap,
      fontSize,
    })
  ) {
    return true;
  }
  if (ratio > 1.8) return true;
  if (textRunContinuation) return false;
  if (
    !sameRunStyle(previous, next) &&
    gap > fontSize * 0.18 &&
    !proseLeadContinuation
  ) {
    return true;
  }
  return false;
}

function projectedRunForText(
  item: ProjectedItem,
  text: string,
): ProjectedTextStyleRun {
  return {
    text,
    fontSize: item.fontSize,
    fontFamily: item.fontFamily,
    displayFontFamily: item.displayFontFamily,
    sourceFontName: item.sourceFontName,
    bold: item.bold,
    italic: item.italic,
    color: item.color,
    sampleX: item.x,
    sampleY: item.baselineY - item.fontSize,
    sampleWidth: Math.max(item.width, item.fontSize * 0.2),
    sampleHeight: item.fontSize * 1.25,
  };
}

function trimProjectedRunsEnd(
  runs: readonly ProjectedTextStyleRun[],
  length: number,
): ProjectedTextStyleRun[] {
  const out: ProjectedTextStyleRun[] = [];
  let remaining = Math.max(0, length);
  for (const run of runs) {
    if (remaining <= 0) break;
    const text = run.text.slice(0, remaining);
    if (text) out.push({ ...run, text });
    remaining -= text.length;
  }
  return mergePdfTextStyleRuns(out);
}

function textAndRunsFromProjectedItems(items: readonly ProjectedItem[]): {
  text: string;
  runs: ProjectedTextStyleRun[];
} {
  let text = "";
  const runs: ProjectedTextStyleRun[] = [];
  const append = (value: string, item: ProjectedItem) => {
    if (!value) return;
    text += value;
    runs.push(projectedRunForText(item, value));
  };
  items.forEach((item, index) => {
    const previous = items[index - 1];
    if (previous) {
      const gap = item.x - (previous.x + previous.width);
      const fontSize = Math.max(previous.fontSize, item.fontSize);
      const ratio = fontSizeRatio(previous.fontSize, item.fontSize);
      const wordContinuation = isAlphabeticWordContinuationWithGap({
        previousText: previous.str,
        nextText: item.str,
        gap,
        fontSize,
        fontSizeRatioValue: ratio,
      });
      const noSpaceContinuation =
        wordContinuation ||
        isNumericContinuationWithGap({
          previousText: previous.str,
          nextText: item.str,
          gap,
          fontSize,
          fontSizeRatioValue: ratio,
        }) ||
        isCjkContinuationWithGap({
          previousText: previous.str,
          nextText: item.str,
          gap,
          fontSize,
          fontSizeRatioValue: ratio,
        });
      if (
        !noSpaceContinuation &&
        gap > previous.fontSize * 0.25 &&
        !/\s$/.test(text)
      ) {
        append(" ", previous);
      }
    }
    if (previous) {
      const suffix = latinContinuationSuffix(item.str);
      const gap = item.x - (previous.x + previous.width);
      const fontSize = Math.max(previous.fontSize, item.fontSize);
      const wordContinuation = isAlphabeticWordContinuationWithGap({
        previousText: previous.str,
        nextText: item.str,
        gap,
        fontSize,
        fontSizeRatioValue: fontSizeRatio(previous.fontSize, item.fontSize),
      });
      if (suffix && wordContinuation) {
        const nextText = appendLatinContinuationSuffix(text, suffix);
        append(nextText.slice(text.length), item);
        return;
      }
    }
    append(item.str, item);
  });
  const trimmed = text.trimEnd();
  return {
    text: trimmed,
    runs: trimProjectedRunsEnd(runs, trimmed.length),
  };
}

function textFromProjectedItems(items: readonly ProjectedItem[]): string {
  return textAndRunsFromProjectedItems(items).text;
}

function shouldDetachPlotLabelFromInlineMetric(
  line: readonly ProjectedItem[],
  next: ProjectedItem,
): boolean {
  const currentText = textFromProjectedItems(line).trim();
  const nextText = next.str.trim();
  if (!currentText || !nextText) return false;
  if (isWeakScalarMetricLine(currentText) && isShortPlotLabelText(nextText)) {
    return true;
  }
  if (!isShortPlotLabelText(nextText) || isWeakScalarMetricLine(nextText)) {
    return false;
  }
  const previous = line[line.length - 1];
  const gap = next.x - (previous.x + previous.width);
  const fontSize = Math.max(previous.fontSize, next.fontSize, 1);
  return gap > fontSize * 1.35 && hasMathExpressionEvidence(currentText);
}

interface LineBoxInternal {
  left: number;
  right: number;
  top: number;
  bottom: number;
  /** Clockwise angle of the run on screen; absent means upright. */
  angle?: number;
  fontSize: number;
  text: string;
  fontFamily?: string;
  displayFontFamily?: string;
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  /** Styled text runs; fontSize is CSS px until final page normalization. */
  textRuns?: ProjectedTextStyleRun[];
  lineLayout?: PdfTextLineLayout[];
  kind?: "text" | "formula";
  latex?: string;
  explicitLatex?: string;
  formulaRuns?: PdfFormulaRun[];
  equationAnchored?: boolean;
}

function lineAsFormulaRun(line: LineBoxInternal): PdfFormulaRun {
  return {
    left: line.left,
    right: line.right,
    centerY: lineCenterY(line),
    fontSize: line.fontSize,
    text: line.text,
  };
}

function formulaRunsFromLines(
  lines: readonly LineBoxInternal[],
): PdfFormulaRun[] {
  return lines.flatMap((line) => line.formulaRuns ?? [lineAsFormulaRun(line)]);
}

function formulaLatexFromProjectedLineItems(
  line: readonly ProjectedItem[],
  formulaRuns: readonly PdfFormulaRun[],
  trustedExplicitLatex: string,
): string {
  if (!trustedExplicitLatex.trim()) {
    return latexFromPdfFormulaRuns(formulaRuns);
  }

  const parts: string[] = [];
  let pendingRuns: PdfFormulaRun[] = [];

  const flushPendingRuns = () => {
    if (pendingRuns.length === 0) return;
    const text = pendingRuns.map((run) => run.text.trim()).join(" ");
    const latex =
      latexFromPdfFormulaRuns(pendingRuns) || latexFromPdfFormulaText(text);
    if (latex.trim()) parts.push(latex.trim());
    pendingRuns = [];
  };

  line.forEach((item, index) => {
    const explicit = item.latex?.trim();
    if (explicit) {
      flushPendingRuns();
      parts.push(explicit);
      return;
    }

    if (!item.str.trim()) return;
    const run = formulaRuns[index];
    if (run) {
      pendingRuns.push(run);
      return;
    }

    flushPendingRuns();
    const latex = latexFromPdfFormulaText(item.str);
    if (latex.trim()) parts.push(latex.trim());
  });
  flushPendingRuns();

  return parts.length > 0
    ? parts.join(" ").replace(/\s+/g, " ").trim()
    : trustedExplicitLatex.trim();
}

function lineBoxFromItems(line: ProjectedItem[]): LineBoxInternal {
  const fontSize = Math.max(...line.map((p) => p.fontSize));
  const dominant = line.reduce(
    (a, b) => (b.fontSize > a.fontSize ? b : a),
    line[0],
  );
  // Measured along the line's OWN axes. Upright text has along === x and
  // across === baselineY, so this is the previous measurement unchanged.
  const angle = line[0]?.angle ?? 0;
  const localLeft = Math.min(...line.map((p) => p.along));
  const localRight = Math.max(...line.map((p) => p.along + p.width));
  const localBaseline = Math.max(...line.map((p) => p.across));
  const localTop = localBaseline - fontSize;
  // Back to page axes: the rect keeps the size it has in its own frame, and
  // its CENTRE is where that frame puts it — which is exactly the model an
  // `angle` plus an axis-aligned rect describes.
  const localCentre = {
    x: (localLeft + localRight) / 2,
    y: (localTop + localTop + fontSize * 1.25) / 2,
  };
  const centre =
    angle === 0
      ? localCentre
      : rotatePointAround(localCentre, { x: 0, y: 0 }, angle);
  const width = localRight - localLeft;
  const left = angle === 0 ? localLeft : centre.x - width / 2;
  const right = angle === 0 ? localRight : centre.x + width / 2;
  const top = angle === 0 ? localTop : centre.y - (fontSize * 1.25) / 2;
  const explicitLatex = line
    .map((p) => p.latex?.trim())
    .filter((latex): latex is string => Boolean(latex))
    .join(" ");
  const projectedText = textAndRunsFromProjectedItems(line);
  const rawText = projectedText.text;
  const latexFallbackText =
    rawText.trim() || !explicitLatex.trim()
      ? ""
      : fallbackTextFromPdfFormulaLatex(explicitLatex).trim();
  const text = rawText.trim() ? rawText.trimEnd() : latexFallbackText;
  const isProtectedProseOrFragment = isNonFormulaProseOrFragmentText(text);
  const isWeakScalarMetric = isWeakScalarMetricLine(text);
  const isDanglingFormulaSource = isClearlyDanglingFormulaSourceText(text);
  const trustedExplicitLatex = isProtectedProseOrFragment ? "" : explicitLatex;
  const hasTrustedExplicitLatex = trustedExplicitLatex.trim().length > 0;
  // A line set mostly in a dedicated math face (Cambria Math, CMMI, …) is a
  // formula regardless of what its characters look like: broken math-font
  // ToUnicode CMaps make the extracted text unrecognizable as math (`x→c`
  // reads as `2→3`), so character heuristics alone systematically miss these.
  const nonSpaceChars = (value: string) => value.replace(/\s+/g, "").length;
  const totalLineChars = line.reduce((sum, p) => sum + nonSpaceChars(p.str), 0);
  const mathFontLineChars = line.reduce(
    (sum, p) => sum + (p.mathFont ? nonSpaceChars(p.str) : 0),
    0,
  );
  const isMostlyMathFontLine =
    totalLineChars > 0 && mathFontLineChars / totalLineChars >= 0.5;
  // An overwhelmingly math-font line bypasses even the prose-fragment guard:
  // that guard exists to keep sentences out of formula boxes, but prose is
  // never typeset in a dedicated math face — while short operator fragments
  // ("lim") and charcode garbage ("!→#") routinely trip it.
  const isOverwhelminglyMathFontLine =
    totalLineChars > 0 && mathFontLineChars / totalLineChars >= 0.9;
  const kind =
    !isWeakScalarMetric &&
    !isDanglingFormulaSource &&
    (isFormulaLikeText(text) ||
      isOverwhelminglyMathFontLine ||
      ((line.some((p) => p.kind === "formula") ||
        hasTrustedExplicitLatex ||
        isMostlyMathFontLine) &&
        !isProtectedProseOrFragment))
      ? "formula"
      : "text";
  const formulaRuns: PdfFormulaRun[] = line.map((p) => ({
    left: p.x,
    right: p.x + p.width,
    centerY: p.baselineY - p.fontSize * 0.375,
    baselineY: p.baselineY,
    fontSize: p.fontSize,
    text: p.str,
  }));
  const latex =
    kind === "formula"
      ? formulaLatexFromProjectedLineItems(
          line,
          formulaRuns,
          trustedExplicitLatex,
        )
      : undefined;
  return {
    left,
    right,
    top,
    bottom: top + fontSize * 1.25,
    ...(angle === 0 ? {} : { angle }),
    fontSize,
    text,
    fontFamily: dominant?.fontFamily,
    displayFontFamily: dominant?.displayFontFamily,
    sourceFontName: dominant?.sourceFontName,
    bold: dominant?.bold,
    italic: dominant?.italic,
    color: dominant?.color,
    textRuns:
      text === rawText && textFromPdfTextStyleRuns(projectedText.runs) === text
        ? projectedText.runs
        : undefined,
    kind,
    latex,
    explicitLatex: trustedExplicitLatex ? latex : undefined,
    formulaRuns: kind === "formula" ? formulaRuns : undefined,
  };
}

function splitInlinePlotMetricSuffixLine(
  line: LineBoxInternal,
): LineBoxInternal[] {
  const text = line.text.trim();
  const match = /^(.*\d(?:\.\d+)?%?)([xyt])$/i.exec(text);
  if (!match) return [line];

  const prefix = match[1].trimEnd();
  const suffix = match[2];
  if (!isWeakScalarMetricLine(prefix) || !isBareMathAtomText(suffix)) {
    return [line];
  }

  const width = line.right - line.left;
  const chars = Math.max(1, Array.from(text).length);
  const suffixChars = Math.max(1, Array.from(suffix).length);
  const suffixWidth = Math.min(
    width * 0.4,
    Math.max(line.fontSize * 0.45, width * (suffixChars / chars)),
  );
  if (suffixWidth <= 0 || width - suffixWidth <= line.fontSize * 0.8) {
    return [line];
  }

  const splitX = line.right - suffixWidth;
  const centerY = lineCenterY(line);
  const prefixRun: PdfFormulaRun = {
    left: line.left,
    right: splitX,
    centerY,
    fontSize: line.fontSize,
    text: prefix,
  };
  return [
    {
      ...line,
      right: splitX,
      text: prefix,
      kind: "formula",
      latex: latexFromPdfFormulaText(prefix),
      explicitLatex: undefined,
      formulaRuns: [prefixRun],
    },
    {
      ...line,
      left: splitX,
      text: suffix,
      kind: "text",
      latex: undefined,
      explicitLatex: undefined,
      formulaRuns: undefined,
    },
  ];
}

function splitLooseTrailingFormulaLabelLine(
  line: LineBoxInternal,
): LineBoxInternal[] {
  const runs = line.formulaRuns;
  if (!runs || runs.length < 2) return [line];

  const orderedRuns = [...runs].sort(
    (a, b) => a.left - b.left || (a.right ?? a.left) - (b.right ?? b.left),
  );
  const trailingEquationRun = isEquationNumberLine(
    orderedRuns[orderedRuns.length - 1]?.text.trim() ?? "",
  )
    ? orderedRuns[orderedRuns.length - 1]
    : null;
  const bodyRuns = trailingEquationRun ? orderedRuns.slice(0, -1) : orderedRuns;
  if (bodyRuns.length < 2) return [line];

  const suffixRun = bodyRuns[bodyRuns.length - 1];
  const previousRun = bodyRuns[bodyRuns.length - 2];
  const suffixText = suffixRun.text.trim();
  if (!isShortPlotLabelText(suffixText) || isWeakScalarMetricLine(suffixText)) {
    return [line];
  }

  const gap = suffixRun.left - (previousRun.right ?? previousRun.left);
  const fontSize = Math.max(line.fontSize, suffixRun.fontSize, 1);
  if (gap <= fontSize * 1.35) return [line];

  const prefixRuns = bodyRuns.slice(0, -1);
  const prefixText = prefixRuns
    .map((run) => run.text.trim())
    .filter(Boolean)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
  if (!prefixText || !isFormulaLikeText(prefixText)) return [line];

  const prefixRight = Math.max(
    ...prefixRuns.map((run) => run.right ?? run.left),
  );
  const suffixWidth = Math.max(
    1,
    (suffixRun.right ?? suffixRun.left) - suffixRun.left,
  );
  const suffixLine: LineBoxInternal = {
    ...line,
    left: suffixRun.left,
    right: suffixRun.left + suffixWidth,
    text: suffixText,
    kind: "text",
    latex: undefined,
    explicitLatex: undefined,
    formulaRuns: undefined,
  };
  const prefixLine: LineBoxInternal = {
    ...line,
    right: prefixRight,
    text: prefixText,
    kind: "formula",
    latex: latexFromPdfFormulaRuns(prefixRuns),
    explicitLatex: undefined,
    formulaRuns: prefixRuns,
  };
  if (!trailingEquationRun) return [prefixLine, suffixLine];

  const equationWidth = Math.max(
    1,
    (trailingEquationRun.right ?? trailingEquationRun.left) -
      trailingEquationRun.left,
  );
  const equationLine: LineBoxInternal = {
    ...line,
    left: trailingEquationRun.left,
    right: trailingEquationRun.left + equationWidth,
    text: trailingEquationRun.text.trim(),
    kind: "text",
    latex: undefined,
    explicitLatex: undefined,
    formulaRuns: undefined,
  };
  return [prefixLine, suffixLine, equationLine];
}

function embeddedEquationNumberSuffix(
  text: string,
): { prefix: string; label: string } | null {
  const trimmed = text.trimEnd();
  const match =
    /^(.*?)(\s*[(\[（［【][^)\]）］】]+[)\]）］】]\s*[.,:;，：；]?)$/.exec(
      trimmed,
    );
  if (!match) return null;

  const prefix = match[1].trimEnd();
  const label = match[2].trim();
  if (!prefix || !isEquationNumberLine(label)) return null;
  if (hasProseText(prefix) && !isFormulaLikeText(prefix)) return null;
  if (
    !isFormulaLikeText(prefix) &&
    !isCompleteRelationFormulaText(prefix) &&
    !hasStrongMathOperator(prefix) &&
    mathSignalScore(prefix) < 2
  ) {
    return null;
  }
  return { prefix, label };
}

function splitEmbeddedEquationNumberLine(
  line: LineBoxInternal,
): LineBoxInternal[] {
  if (isEquationNumberLine(line.text)) return [line];

  const embedded = embeddedEquationNumberSuffix(line.text);
  if (!embedded) return [line];

  const width = line.right - line.left;
  const chars = Math.max(1, Array.from(line.text.trimEnd()).length);
  const labelChars = Math.max(1, Array.from(embedded.label).length);
  const labelWidth = Math.min(
    width * 0.48,
    Math.max(line.fontSize * 1.6, width * (labelChars / chars)),
  );
  if (labelWidth <= 0 || width - labelWidth <= line.fontSize * 1.2) {
    return [line];
  }

  const splitX = line.right - labelWidth;
  const centerY = lineCenterY(line);
  const prefixRun: PdfFormulaRun = {
    left: line.left,
    right: splitX,
    centerY,
    fontSize: line.fontSize,
    text: embedded.prefix,
  };

  return [
    {
      ...line,
      right: splitX,
      text: embedded.prefix,
      kind: "formula",
      latex: latexFromPdfFormulaText(embedded.prefix),
      explicitLatex: undefined,
      formulaRuns: [prefixRun],
    },
    {
      ...line,
      left: splitX,
      text: embedded.label,
      kind: "text",
      latex: undefined,
      explicitLatex: undefined,
      formulaRuns: undefined,
    },
  ];
}

function isEquationNumberFragmentLine(line: LineBoxInternal): boolean {
  const canonical = canonicalEquationNumberLineText(line.text);
  if (!canonical || canonical.length > 18) return false;
  if (!/^[()[\]A-Za-z0-9.,:;\-'′’*★式公式方程]+$/u.test(canonical)) {
    return false;
  }
  if (/[a-z]/.test(canonical) && !/^Eq\.?$/i.test(canonical)) return false;
  return true;
}

function mergeEquationNumberFragmentLines(
  group: readonly LineBoxInternal[],
): LineBoxInternal {
  const dominant = group.reduce(
    (a, b) => (b.fontSize > a.fontSize ? b : a),
    group[0],
  );
  return {
    left: Math.min(...group.map((line) => line.left)),
    right: Math.max(...group.map((line) => line.right)),
    top: Math.min(...group.map((line) => line.top)),
    bottom: Math.max(...group.map((line) => line.bottom)),
    fontSize: dominant.fontSize,
    text: group.map((line) => line.text.trim()).join(""),
    fontFamily: dominant.fontFamily,
    displayFontFamily: dominant.displayFontFamily,
    sourceFontName: dominant.sourceFontName,
    bold: dominant.bold,
    italic: dominant.italic,
    color: dominant.color,
    kind: "text",
    latex: undefined,
    explicitLatex: undefined,
    formulaRuns: undefined,
  };
}

function mergeSplitEquationNumberLines(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] {
  const merged: LineBoxInternal[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    const first = lines[index];
    if (!isEquationNumberFragmentLine(first)) {
      merged.push(first);
      continue;
    }

    let found: { endIndex: number; line: LineBoxInternal } | null = null;
    const group: LineBoxInternal[] = [];
    for (
      let cursor = index;
      cursor < Math.min(lines.length, index + 6);
      cursor += 1
    ) {
      const line = lines[cursor];
      const previous = group[group.length - 1];
      if (!isEquationNumberFragmentLine(line)) break;
      if (
        previous &&
        (!sameInlineRow(previous, line) || !inlineFragmentGapOk(previous, line))
      ) {
        break;
      }
      group.push(line);
      const candidate = mergeEquationNumberFragmentLines(group);
      if (group.length > 1 && isEquationNumberLine(candidate.text)) {
        found = { endIndex: cursor, line: candidate };
        break;
      }
    }

    if (found) {
      merged.push(found.line);
      index = found.endIndex;
    } else {
      merged.push(first);
    }
  }
  return merged;
}

function lineOverlapsOrNests(
  group: readonly LineBoxInternal[],
  next: LineBoxInternal,
): boolean {
  const left = Math.min(...group.map((line) => line.left));
  const right = Math.max(...group.map((line) => line.right));
  const fontSize = Math.max(
    next.fontSize,
    ...group.map((line) => line.fontSize),
  );
  const overlap = Math.min(right, next.right) - Math.max(left, next.left);
  const nextWidth = Math.max(1, next.right - next.left);
  const groupWidth = Math.max(1, right - left);
  if (overlap > 0) return true;

  const nextCenter = (next.left + next.right) / 2;
  const groupCenter = (left + right) / 2;
  const centerDistance = Math.abs(nextCenter - groupCenter);
  const minWidth = Math.min(groupWidth, nextWidth);
  if (centerDistance <= Math.max(fontSize * 1.75, minWidth * 0.55)) {
    return true;
  }
  const hasStructuralEvidence =
    group.some(hasStructuralFormulaEvidence) ||
    hasStructuralFormulaEvidence(next);
  if (
    hasStructuralEvidence &&
    centerDistance <=
      Math.max(fontSize * 3.5, Math.max(groupWidth, nextWidth) * 0.35)
  ) {
    return true;
  }
  return (
    Math.min(Math.abs(next.left - right), Math.abs(left - next.right)) <=
    fontSize * 4
  );
}

function shouldMergeFormulaLines(
  group: readonly LineBoxInternal[],
  next: LineBoxInternal,
): boolean {
  if (isEquationNumberLine(next.text)) return false;
  if (next.kind !== "formula") return false;
  if (!group.some((line) => line.kind === "formula")) return false;
  if (next.equationAnchored || group.some((line) => line.equationAnchored)) {
    return false;
  }
  if (crossesIndependentUnnumberedFormulaBoundary(group, next)) {
    return false;
  }
  if (
    isSelfContainedFormulaBoundaryLine(group[group.length - 1]) &&
    isSelfContainedFormulaBoundaryLine(next) &&
    !repeatedLeftHandSideLooksLikeDerivation([...group, next]) &&
    !chainedRelationLooksLikeDerivation([...group, next])
  ) {
    return false;
  }
  if (
    [...group, next].every((line) => isWeakScalarMetricLine(line.text.trim()))
  ) {
    return false;
  }
  const groupHasWeakScalarMetric = group.some((line) =>
    isWeakScalarMetricLine(line.text.trim()),
  );
  const nextIsWeakScalarMetric = isWeakScalarMetricLine(next.text.trim());
  if (groupHasWeakScalarMetric !== nextIsWeakScalarMetric) {
    return false;
  }
  const hasExplicitFormulaSource = [...group, next].some((line) =>
    Boolean(line.explicitLatex?.trim()),
  );
  if (
    !hasExplicitFormulaSource &&
    isOnlyLowEvidenceFormulaCluster([...group, next])
  ) {
    return false;
  }

  const bottom = Math.max(...group.map((line) => line.bottom));
  const top = Math.min(...group.map((line) => line.top));
  const fontSize = Math.max(
    next.fontSize,
    ...group.map((line) => line.fontSize),
  );
  const previous = group[group.length - 1];
  if (
    previous &&
    sameInlineRow(previous, next) &&
    !sameRowFormulaBoundaryContinuesExpression(group, next)
  ) {
    return false;
  }
  if (isVerticalIndependentRelationCluster([...group, next])) {
    return false;
  }
  const verticalGap = next.top - bottom;
  if (verticalGap > fontSize * 1.35) return false;
  if (next.bottom < top - fontSize * 0.4) return false;
  return lineOverlapsOrNests(group, next);
}

function lineCenterY(line: LineBoxInternal): number {
  return (line.top + line.bottom) / 2;
}

function lineCenterX(line: LineBoxInternal): number {
  return (line.left + line.right) / 2;
}

function lineWidth(line: LineBoxInternal): number {
  return Math.max(0, line.right - line.left);
}

function canonicalFootnoteMarkerText(text: string): string {
  return text.trim().replace(/\s+/g, "").normalize("NFKC");
}

function isFootnoteMarkerText(text: string): boolean {
  const compact = canonicalFootnoteMarkerText(text);
  return /^(?:[0-9]|[1-9][0-9]|[A-Za-z]|[*†‡§]|注(?:[0-9]|[1-9][0-9]))$/.test(
    compact,
  );
}

function isConservativeFootnoteMarkerText(text: string): boolean {
  const compact = canonicalFootnoteMarkerText(text);
  return /^(?:1|[*†‡§])$/.test(compact);
}

function lineLooksLikeTrailingFormulaFootnoteMarker(
  line: LineBoxInternal,
  formulaLines: readonly LineBoxInternal[],
): boolean {
  if (!isFootnoteMarkerText(line.text)) return false;
  const candidates = formulaLines.filter((candidate) => candidate !== line);
  if (candidates.length === 0) return false;

  const fontSize = Math.max(
    ...candidates.map((candidate) => candidate.fontSize),
    1,
  );
  if (line.fontSize > fontSize * 0.72) return false;

  const clusterRight = Math.max(
    ...candidates.map((candidate) => candidate.right),
  );
  const clusterCenterY =
    candidates.reduce((sum, candidate) => sum + lineCenterY(candidate), 0) /
    candidates.length;
  const trailingGap = line.left - clusterRight;
  const conservativeMarker = isConservativeFootnoteMarkerText(line.text);
  if (trailingGap < line.fontSize * (conservativeMarker ? 0.45 : 0.9)) {
    return false;
  }
  if (!conservativeMarker && trailingGap < fontSize * 1.2) {
    return false;
  }
  return lineCenterY(line) < clusterCenterY - line.fontSize * 0.2;
}

function stripTrailingFormulaFootnoteMarkers(
  cluster: readonly { line: LineBoxInternal; index: number }[],
): Array<{ line: LineBoxInternal; index: number }> {
  return cluster.filter(
    ({ line }) =>
      !lineLooksLikeTrailingFormulaFootnoteMarker(
        line,
        cluster.map((item) => item.line),
      ),
  );
}

function isVisualFractionBarLine(line: LineBoxInternal): boolean {
  return (
    isVisualFractionBarText(line.text.trim()) &&
    lineWidth(line) >= Math.max(1, line.fontSize * 1.8)
  );
}

function isFractionPartLine(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (!text || proseWordCount(text) > 0 || isWeakScalarMetricLine(text)) {
    return false;
  }
  if (isMatrixLikeNumericRowText(text)) return false;
  if (isQuantityWithFormulaUnitText(text)) return true;

  const compact = text.replace(/\s+/g, "");
  if (!compact) return false;
  if (/^[+\-−]?\d+(?:\.\d+)?$/.test(compact)) return true;
  if (isBareMathAtomText(text) || isCompactMathAtom(text)) return true;
  if (/[=+\-−*/×÷^_()[\]{}|]/.test(text)) return true;
  if (hasMathTypographyEvidence(text)) return true;
  if (mathSignalScore(text) > 0) return true;
  return /^[A-Za-z0-9αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]+$/.test(compact);
}

function lineFitsFractionBarSpan(
  line: LineBoxInternal,
  bar: LineBoxInternal,
): boolean {
  const pad = Math.max(line.fontSize, bar.fontSize) * 2.5;
  return line.left >= bar.left - pad && line.right <= bar.right + pad;
}

function lineIsCenteredOnFractionBar(
  line: LineBoxInternal,
  bar: LineBoxInternal,
): boolean {
  const pad = Math.max(line.fontSize, bar.fontSize) * 2;
  const center = lineCenterX(line);
  return center >= bar.left - pad && center <= bar.right + pad;
}

function hasVisualFractionClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  const lines = cluster.map(({ line }) => line);
  return lines.some((bar) => {
    if (!isVisualFractionBarLine(bar)) return false;
    const barCenterY = lineCenterY(bar);
    const partsForDirection = (direction: "above" | "below") =>
      lines.filter((line) => {
        if (line === bar || !isFractionPartLine(line)) return false;
        const centerY = lineCenterY(line);
        if (direction === "above" && centerY >= barCenterY) return false;
        if (direction === "below" && centerY <= barCenterY) return false;
        return (
          lineFitsFractionBarSpan(line, bar) &&
          lineIsCenteredOnFractionBar(line, bar)
        );
      });
    const aboveParts = partsForDirection("above");
    const belowParts = partsForDirection("below");
    return aboveParts.some((above) =>
      belowParts.some((below) =>
        visualFractionCandidateHasMathEvidence(lines, above, below),
      ),
    );
  });
}

function visualFractionCandidateHasMathEvidence(
  lines: readonly LineBoxInternal[],
  above: LineBoxInternal,
  below: LineBoxInternal,
): boolean {
  const aboveText = above.text.trim();
  const belowText = below.text.trim();
  if (!isNumericTickText(aboveText) || !isNumericTickText(belowText)) {
    return true;
  }
  return lines.some((line) => {
    if (isVisualFractionBarLine(line)) return false;
    const text = line.text.trim();
    return (
      line.kind === "formula" ||
      Boolean(line.explicitLatex?.trim()) ||
      hasStrongMathOperator(text) ||
      mathSignalScore(text) >= 2 ||
      isQuantityWithFormulaUnitText(text) ||
      /[A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}]/u.test(text)
    );
  });
}

function lineLooksLikeSafeStackedFractionDenominator(
  line: LineBoxInternal,
): boolean {
  const compact = line.text.trim().replace(/\s+/g, "");
  if (!compact) return false;
  if (isBareMathAtomText(compact)) return false;
  return (
    /^[+\-−]?\d/.test(compact) ||
    /[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(compact) ||
    /[+\-−*/×÷()[\]{}^_]/.test(compact)
  );
}

function hasStackedFractionClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  const lines = cluster
    .map(({ line }) => line)
    .filter((line) => !isVisualFractionBarLine(line));
  if (lines.length !== 2) return false;

  const [above, below] = [...lines].sort(
    (a, b) => lineCenterY(a) - lineCenterY(b),
  );
  if (!isFractionPartLine(above) || !isFractionPartLine(below)) return false;
  if (
    isSelfContainedFormulaBoundaryLine(above) &&
    isSelfContainedFormulaBoundaryLine(below)
  ) {
    return false;
  }
  if (!lineLooksLikeSafeStackedFractionDenominator(below)) return false;
  if (/[=<>≤≥≈≠]/.test(above.text) || /[=<>≤≥≈≠]/.test(below.text)) {
    return false;
  }

  const fontSize = Math.max(above.fontSize, below.fontSize, 1);
  const verticalDistance = lineCenterY(below) - lineCenterY(above);
  if (verticalDistance < fontSize * 0.75 || verticalDistance > fontSize * 3.6) {
    return false;
  }
  if (Math.abs(lineCenterX(above) - lineCenterX(below)) > fontSize * 2.2) {
    return false;
  }
  if (lineWidth(above) > lineWidth(below) * 1.45 + fontSize) {
    return false;
  }
  return true;
}

function lineHasStackedFractionMathEvidence(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (isWeakScalarMetricLine(text)) return false;
  if (line.kind === "formula" || line.explicitLatex?.trim()) return true;
  return (
    isQuantityWithFormulaUnitText(text) ||
    isCompactMathAtom(text) ||
    mathSignalScore(text) >= 1 ||
    /[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(text) ||
    /[⁰¹²³⁴⁵⁶⁷⁸⁹ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ₀₁₂₃₄₅₆₇₈₉]/.test(text)
  );
}

function hasUnnumberedStackedFractionClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  if (!hasStackedFractionClusterSeed(cluster)) return false;
  const lines = cluster
    .map(({ line }) => line)
    .filter((line) => !isVisualFractionBarLine(line));
  if (lines.length !== 2) return false;
  if (isOnlyLowEvidenceFormulaCluster(lines)) return false;
  return lines.some(lineHasStackedFractionMathEvidence);
}

function isCorruptedZeroGlyphLine(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (text !== "!") return false;
  if (line.kind === "formula" || line.explicitLatex?.trim()) return false;
  return lineWidth(line) <= Math.max(line.fontSize * 1.25, 1);
}

function hasCorruptedZeroStackedFractionGeometry(
  above: LineBoxInternal,
  below: LineBoxInternal,
): boolean {
  if (!isCorruptedZeroGlyphLine(above) || !isCorruptedZeroGlyphLine(below)) {
    return false;
  }
  const baseFont = Math.max(above.fontSize, below.fontSize, 1);
  if (fontSizeRatio(above.fontSize, below.fontSize) > 1.25) return false;

  const verticalDistance = lineCenterY(below) - lineCenterY(above);
  if (verticalDistance < baseFont * 0.75 || verticalDistance > baseFont * 3.6) {
    return false;
  }

  const maxWidth = Math.max(lineWidth(above), lineWidth(below));
  return (
    Math.abs(lineCenterX(above) - lineCenterX(below)) <=
      Math.max(baseFont * 0.65, maxWidth) && maxWidth <= baseFont * 1.25
  );
}

function corruptedZeroStackedFractionPair(
  a: LineBoxInternal,
  b: LineBoxInternal,
): [LineBoxInternal, LineBoxInternal] | null {
  const [above, below] = lineCenterY(a) <= lineCenterY(b) ? [a, b] : [b, a];
  return hasCorruptedZeroStackedFractionGeometry(above, below)
    ? [above, below]
    : null;
}

function isDisplayFormulaFragment(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (!text || isEquationNumberLine(text)) return false;
  if (isVisualFractionBarText(text)) return true;
  if (isShortFormulaUnitFragmentText(text)) return true;
  if (isCaseLikeRowText(text)) return true;
  if (isMatrixLikeRowText(text)) {
    return !isNonMathematicalMatrixAssignmentRowText(text);
  }
  if (!line.explicitLatex?.trim() && isNonFormulaProseOrFragmentText(text)) {
    return false;
  }
  if (line.kind === "formula" || isFormulaLikeText(text)) return true;
  const proseWords = text.match(/[A-Za-z]{3,}/g) ?? [];
  if (
    proseWords.some(
      (word) =>
        !/^(?:exp|log|ln|sin|cos|tan|var|max|min|arg|lim|det|pr|cov)$/i.test(
          word,
        ) &&
        (word.length > 3 || mathSignalScore(text) < 2),
    )
  ) {
    return false;
  }
  if (mathSignalScore(text) > 0) return true;
  if (/^[(){}\[\]|.,;:+\-=−<>≤≥≈≠±∓×÷·\d]+$/.test(text)) return true;
  if (/^[A-Za-z]$/.test(text)) return true;
  if (/^[([{]?[A-Za-z][)\]}]?$/.test(text)) return true;
  if (isCompactMathAtom(text)) return true;
  if (/^d[A-Za-z]$/.test(text)) return true;
  if (/^[A-Za-z]\/\d+$/.test(text)) return true;
  if (
    /^(?:exp|log|ln|sin|cos|tan|var|max|min|lim|det|Pr|Cov|Var)$/.test(text)
  ) {
    return true;
  }
  if (isMatrixLikeRowText(text)) return true;
  if (/^[+\-−]?\d+(?:\.\d+)?(?:\s+[+\-−]?\d+(?:\.\d+)?)+$/.test(text)) {
    return true;
  }
  if (/^\d+(?:\/\d+)?$/.test(text)) return true;
  return false;
}

function hasDisplayFormulaSeed(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (!isDisplayFormulaFragment(line)) return false;
  if (isLowEvidenceFormulaFragment(line)) return false;
  if (
    isStructuralCaseFormulaRowText(text) ||
    isStructuralMatrixFormulaRowText(text)
  ) {
    return true;
  }
  return (
    (line.kind === "formula" && proseWordCount(text) === 0) ||
    hasStrongMathOperator(text) ||
    mathSignalScore(text) >= 3
  );
}

function isNumberedSingletonFormulaSeed(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (!isDisplayFormulaFragment(line)) return false;
  if (isBareMathAtomText(text)) return true;
  if (isCompactMathAtom(text)) return true;
  if (isLowEvidenceFormulaFragment(line)) return false;
  if (proseWordCount(text) > 0 && !isCompactMathAtom(text)) return false;
  return (
    line.kind === "formula" ||
    (isMatrixLikeNumericRowText(text) && /[=[\](){}]/.test(text)) ||
    isCompactMathAtom(text) ||
    hasStrongMathOperator(text) ||
    mathSignalScore(text) >= 2
  );
}

function hasMatrixLikeNumberedClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  const matrixRows = cluster.filter(({ line }) =>
    isMatrixLikeRowText(line.text),
  );
  if (matrixRows.length < 2) return false;
  return cluster.every(({ line }) => {
    const text = line.text.trim();
    return isMatrixLikeRowText(text) || mathSignalScore(text) > 0;
  });
}

function isMatrixClusterWithoutStructuralEvidence(
  lines: readonly LineBoxInternal[],
): boolean {
  const matrixRows = lines.filter((line) => isMatrixLikeRowText(line.text));
  if (matrixRows.length < 2) return false;
  if (hasUnbracketedNumberedMatrixGeometryEvidence(matrixRows)) return false;
  return !lines.some(
    (line) =>
      isStructuralMatrixFormulaRowText(line.text) ||
      Boolean(line.explicitLatex?.trim()),
  );
}

function hasUnbracketedNumberedMatrixGeometryEvidence(
  rows: readonly LineBoxInternal[],
): boolean {
  if (rows.length < 2 || rows.length > 6) return false;
  if (!rows.every((row) => isMatrixLikeRowText(row.text))) return false;

  const cellsByRow = rows.map((row) => matrixLikeRowCells(row.text));
  const columnCount = cellsByRow[0]?.length ?? 0;
  if (columnCount < 2 || columnCount > 6) return false;
  if (!cellsByRow.every((cells) => cells.length === columnCount)) return false;

  const allCells = cellsByRow.flat();
  const allNumeric = allCells.every((cell) =>
    /^[+\-−]?\d+(?:\.\d+)?$/.test(cell.trim()),
  );
  if (allNumeric) {
    const allSingleDigit = allCells.every((cell) =>
      /^[+\-−]?\d$/.test(cell.trim()),
    );
    if (!allSingleDigit || rows.length !== columnCount) return false;
  }

  const fontSize = Math.max(...rows.map((row) => row.fontSize), 1);
  const lefts = rows.map((row) => row.left);
  const rights = rows.map((row) => row.right);
  const widths = rows.map((row) => lineWidth(row));
  const leftSpread = Math.max(...lefts) - Math.min(...lefts);
  const rightSpread = Math.max(...rights) - Math.min(...rights);
  const maxWidth = Math.max(...widths, 1);
  const minWidth = Math.min(...widths);
  if (leftSpread > fontSize * 1.25) return false;
  if (rightSpread > fontSize * 1.75) return false;
  if (minWidth / maxWidth < 0.72) return false;

  const ordered = [...rows].sort(
    (a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left,
  );
  const gaps = ordered
    .slice(1)
    .map((row, index) => lineCenterY(row) - lineCenterY(ordered[index]));
  if (gaps.some((gap) => gap < fontSize * 0.75 || gap > fontSize * 2.4)) {
    return false;
  }
  if (gaps.length >= 2) {
    const maxGap = Math.max(...gaps);
    const minGap = Math.min(...gaps);
    if (minGap / maxGap < 0.58) return false;
  }
  return true;
}

function latexForMatrixCellText(cell: string): string {
  return latexFromPdfFormulaText(cell).trim() || cell.trim();
}

function unbracketedMatrixLatexFromLines(
  lines: readonly LineBoxInternal[],
): string | null {
  if (!hasUnbracketedNumberedMatrixGeometryEvidence(lines)) return null;
  const rows = [...lines]
    .sort((a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left)
    .map((line) =>
      matrixLikeRowCells(line.text).map(latexForMatrixCellText).join(" & "),
    )
    .filter(Boolean);
  if (rows.length < 2) return null;
  return `\\begin{bmatrix}${rows.join(" \\\\ ")}\\end{bmatrix}`;
}

function unbracketedMatrixTextFromLines(
  lines: readonly LineBoxInternal[],
): string | null {
  if (!hasUnbracketedNumberedMatrixGeometryEvidence(lines)) return null;
  const rows = [...lines]
    .sort((a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left)
    .map((line) => matrixLikeRowCells(line.text).join(" "))
    .filter(Boolean);
  if (rows.length < 2) return null;
  return rows.join("\n");
}

function hasCasesLikeNumberedClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  return cluster.filter(({ line }) => isCaseLikeRowText(line.text)).length >= 2;
}

function hasCasesLikeUnnumberedClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  return (
    cluster.filter(({ line }) => isStructuralCaseFormulaRowText(line.text))
      .length >= 2
  );
}

function hasMatrixLikeUnnumberedClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  const matrixRows = cluster.filter(({ line }) =>
    isMatrixLikeRowText(line.text),
  );
  if (matrixRows.length < 2) return false;
  return matrixRows.some(({ line }) =>
    isStructuralMatrixFormulaRowText(line.text),
  );
}

function hasUnnumberedFormulaClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  const lines = cluster.map(({ line }) => line);
  return (
    hasVisualFractionClusterSeed(cluster) ||
    hasUnnumberedStackedFractionClusterSeed(cluster) ||
    linesHavePositionedScriptFormulaGeometry(lines) ||
    hasCasesLikeUnnumberedClusterSeed(cluster) ||
    hasMatrixLikeUnnumberedClusterSeed(cluster) ||
    cluster.some(({ line }) => hasDisplayFormulaSeed(line))
  );
}

function isDocumentIndexLeaderCluster(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  if (cluster.length < 5) return false;
  const lines = cluster.map(({ line }) => line);
  if (!lines.every((line) => isDocumentIndexLeaderLineText(line.text))) {
    return false;
  }
  if (
    lines.filter((line) => isDocumentIndexLeaderTokenText(line.text)).length < 3
  ) {
    return false;
  }
  if (
    lines.some((line) =>
      /[=<>≤≥≈≠±∓×÷·∑∫√∞∂∇∏()[\]{}|A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(
        line.text,
      ),
    )
  ) {
    return false;
  }
  const fontSize = Math.max(...lines.map((line) => line.fontSize), 1);
  const clusterWidth =
    Math.max(...lines.map((line) => line.right)) -
    Math.min(...lines.map((line) => line.left));
  return clusterWidth <= fontSize * 8;
}

function hasNumberedFormulaClusterSeed(
  cluster: readonly { line: LineBoxInternal }[],
): boolean {
  return (
    hasVisualFractionClusterSeed(cluster) ||
    hasStackedFractionClusterSeed(cluster) ||
    hasMatrixLikeNumberedClusterSeed(cluster) ||
    hasCasesLikeNumberedClusterSeed(cluster) ||
    cluster.some(
      ({ line }) =>
        hasDisplayFormulaSeed(line) || isNumberedSingletonFormulaSeed(line),
    )
  );
}

function hasStructuralFormulaEvidence(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (!text) return false;
  if (isVisualFractionBarText(text)) return true;
  if (isCaseLikeRowText(text) || isMatrixLikeRowText(text)) return true;
  if (hasStrongMathOperator(text)) return true;
  if (hasCompactFormulaStructure(text)) return true;
  if (mathSignalScore(text) >= 3 && !isLowEvidenceFormulaFragment(line)) {
    return true;
  }
  if (line.latex && line.latex.trim() !== latexFromPdfFormulaText(text)) {
    return true;
  }
  return false;
}

function canAttachLowEvidenceFormulaFragment(
  line: LineBoxInternal,
  cluster: readonly LineBoxInternal[],
): boolean {
  if (!isLowEvidenceFormulaFragment(line)) return true;
  const canAttachAsPositionedScript = cluster.some((body) =>
    linePairHasPositionedScriptFormulaGeometry(body, line),
  );
  const canAttachToVisualFraction = cluster.some(isVisualFractionBarLine);
  if (
    !cluster.some(hasStructuralFormulaEvidence) &&
    !canAttachAsPositionedScript &&
    !canAttachToVisualFraction
  ) {
    return false;
  }
  if (isWeakScalarMetricLine(line.text)) return false;

  const text = line.text.trim();
  const clusterFont = Math.max(...cluster.map((item) => item.fontSize), 1);
  const smallRelativeToCluster = line.fontSize <= clusterFont * 0.82;
  if (isCompactMathAtom(text) && !isBareMathAtomText(text)) {
    return canAttachAsPositionedScript || canAttachToVisualFraction;
  }
  if (isNumericTickText(text)) {
    return smallRelativeToCluster && lineOverlapsOrNests(cluster, line);
  }
  if (isBareMathAtomText(text) && !smallRelativeToCluster) {
    return (
      cluster.some(isVisualFractionBarLine) ||
      cluster.some((body) =>
        linePairHasPositionedScriptFormulaGeometry(body, line),
      ) ||
      cluster.some(
        (body) =>
          sameInlineRow(body, line) &&
          hasStructuralFormulaEvidence(body) &&
          lineHorizontallyNearCluster(line, cluster, clusterFont),
      )
    );
  }
  return lineOverlapsOrNests(cluster, line);
}

function isDetachedShortFormulaLabel(
  line: LineBoxInternal,
  cluster: readonly LineBoxInternal[],
): boolean {
  if (!isShortPlotLabelText(line.text) || isWeakScalarMetricLine(line.text)) {
    return false;
  }
  const clusterFont = Math.max(...cluster.map((item) => item.fontSize), 1);
  if (line.fontSize <= clusterFont * 0.82) return false;
  if (cluster.some(isVisualFractionBarLine)) return false;
  if (
    cluster.some((body) =>
      linePairHasPositionedScriptFormulaGeometry(body, line),
    )
  ) {
    return false;
  }
  return true;
}

function clusterContainsDetachedShortFormulaLabel(
  lines: readonly LineBoxInternal[],
): boolean {
  const cluster = lines.map((line) => ({ line }));
  if (
    hasUnnumberedStackedFractionClusterSeed(cluster) ||
    hasStackedFractionClusterSeed(cluster)
  ) {
    return false;
  }
  return lines.some((line, index) => {
    const others = lines.filter((_, otherIndex) => otherIndex !== index);
    return others.length > 0 && isDetachedShortFormulaLabel(line, others);
  });
}

function leadingRelationContinuationText(text: string): string | null {
  const trimmed = formulaBoundaryText(text);
  const match = /^([=<>≤≥≈≠])\s*(.+)$/.exec(trimmed);
  const body = match?.[2]?.trim() ?? "";
  if (!body || endsWithFormulaContinuation(body)) return null;
  if (formulaProseWordCount(body) > 0) return null;
  return body;
}

function isLinewiseFormulaRelationRow(line: LineBoxInternal): boolean {
  const text = formulaBoundaryText(line.text);
  if (!text || formulaProseWordCount(text) > 0) return false;
  return (
    isCompleteRelationFormulaText(text) ||
    leadingRelationContinuationText(text) != null
  );
}

function linewiseAlignedDerivationRows(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] | null {
  if (lines.length < 2) return null;
  if (
    lines.some(
      (line) =>
        isVisualFractionBarText(line.text.trim()) ||
        isCaseLikeRowText(line.text) ||
        isMatrixLikeRowText(line.text),
    )
  ) {
    return null;
  }

  const ordered = [...lines].sort(
    (a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left,
  );
  if (!ordered.every(isLinewiseFormulaRelationRow)) return null;
  if (
    repeatedLeftHandSideLooksLikeDerivation(ordered) ||
    chainedRelationLooksLikeDerivation(ordered)
  ) {
    return ordered;
  }
  if (!isCompleteRelationFormulaText(ordered[0].text)) return null;
  if (
    !ordered
      .slice(1)
      .every((line) => leadingRelationContinuationText(line.text))
  ) {
    return null;
  }
  return ordered;
}

function linewiseFormulaRowLatex(line: LineBoxInternal): string {
  return (
    line.explicitLatex?.trim() ||
    latexFromPdfFormulaText(line.text).trim() ||
    line.latex?.trim() ||
    latexFromPdfFormulaRuns(line.formulaRuns ?? [lineAsFormulaRun(line)])
  );
}

function formulaTextFromLines(lines: readonly LineBoxInternal[]): string {
  const unbracketedMatrixText = unbracketedMatrixTextFromLines(lines);
  if (unbracketedMatrixText) return unbracketedMatrixText;
  const derivationRows = linewiseAlignedDerivationRows(lines);
  if (derivationRows) {
    return derivationRows
      .map((line) =>
        readableFormulaTextFromLatex(
          linewiseFormulaRowLatex(line),
          line.text.trim(),
        ),
      )
      .filter(Boolean)
      .join("\n")
      .trim();
  }
  if (lines.length === 1) {
    const runs = lines[0].formulaRuns;
    if (runs && runs.length > 1 && formulaRunsHaveScriptOffsets(runs)) {
      const visualText = formulaTextFromVisualRuns(
        runs.map((run) => ({
          left: run.left,
          centerY: run.baselineY ?? run.centerY,
          fontSize: run.fontSize,
          text: run.text,
        })),
      );
      if (visualText) return visualText;
    }
    return lines[0].text.trim();
  }
  return formulaTextFromVisualRuns(
    lines.map((line) => ({
      left: line.left,
      centerY: lineCenterY(line),
      fontSize: line.fontSize,
      text: line.text,
    })),
  );
}

function formulaRunsHaveScriptOffsets(runs: readonly PdfFormulaRun[]): boolean {
  for (let i = 0; i < runs.length; i += 1) {
    for (let j = i + 1; j < runs.length; j += 1) {
      const a = runs[i];
      const b = runs[j];
      const largerFont = Math.max(a.fontSize, b.fontSize);
      const smallerFont = Math.min(a.fontSize, b.fontSize);
      if (smallerFont > largerFont * 0.72) continue;
      const centerDelta = Math.abs(a.centerY - b.centerY);
      const baselineDelta = Math.abs(
        (a.baselineY ?? a.centerY) - (b.baselineY ?? b.centerY),
      );
      if (
        centerDelta > smallerFont * 0.15 &&
        baselineDelta > Math.max(smallerFont * 0.42, largerFont * 0.24)
      ) {
        return true;
      }
    }
  }
  return false;
}

function formulaLatexFromLines(lines: readonly LineBoxInternal[]): string {
  if (lines.length === 1 && lines[0].latex) return lines[0].latex;
  const unbracketedMatrixLatex = unbracketedMatrixLatexFromLines(lines);
  if (unbracketedMatrixLatex) return unbracketedMatrixLatex;
  const derivationRows = linewiseAlignedDerivationRows(lines);
  if (derivationRows) {
    const rowLatex = derivationRows
      .map(linewiseFormulaRowLatex)
      .filter(Boolean);
    if (rowLatex.length <= 1) return rowLatex[0] ?? "";
    return `\\begin{aligned}${rowLatex.join(" \\\\ ")}\\end{aligned}`;
  }
  if (lines.some((line) => line.explicitLatex?.trim())) {
    const sorted = [...lines].sort(
      (a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left,
    );
    const rows: LineBoxInternal[][] = [];
    for (const line of sorted) {
      const row = rows[rows.length - 1];
      const rowFont = row
        ? Math.max(...row.map((item) => item.fontSize), line.fontSize)
        : line.fontSize;
      const rowCenter = row
        ? row.reduce((sum, item) => sum + lineCenterY(item), 0) / row.length
        : lineCenterY(line);
      if (!row || Math.abs(lineCenterY(line) - rowCenter) > rowFont * 0.45) {
        rows.push([line]);
      } else {
        row.push(line);
      }
    }

    const rowLatex = rows
      .map((row) =>
        row
          .sort((a, b) => a.left - b.left)
          .map((line) => {
            const explicit = line.explicitLatex?.trim();
            if (explicit) return explicit;
            if (line.latex?.trim()) return line.latex.trim();
            return latexFromPdfFormulaRuns(
              line.formulaRuns ?? [lineAsFormulaRun(line)],
            );
          })
          .filter(Boolean)
          .join(" ")
          .replace(/\s+/g, " ")
          .trim(),
      )
      .filter(Boolean);
    if (rowLatex.length <= 1) return rowLatex[0] ?? "";
    return `\\begin{aligned}${rowLatex.join(" \\\\ ")}\\end{aligned}`;
  }
  return latexFromPdfFormulaRuns(formulaRunsFromLines(lines));
}

function readableFormulaTextFromLatex(
  latex: string | undefined,
  visualText: string,
): string {
  const fallback =
    typeof latex === "string" ? fallbackTextFromPdfFormulaLatex(latex) : "";
  const readable = fallback.trim();
  const visual = visualText.trim();
  if (
    readable &&
    visual &&
    /[\wαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]\^1[)\]}]\s*=/.test(readable)
  ) {
    return visual;
  }
  if (
    readable &&
    visual &&
    /\b(?:exp|log|ln|sin|cos|tan)\^-?\d/.test(readable)
  ) {
    return visual;
  }
  return readable || visual;
}

function readableFormulaTextFromLineGroup(
  lines: readonly LineBoxInternal[],
  latex: string | undefined = formulaLatexFromLines(lines),
): string {
  const visualText = formulaTextFromLines(lines);
  return linewiseAlignedDerivationRows(lines)
    ? visualText
    : readableFormulaTextFromLatex(latex, visualText);
}

function hasLargeSameRowGap(
  lines: readonly LineBoxInternal[],
  maxGap: number,
): boolean {
  const fragments: LineBoxInternal[] = lines.flatMap((line) => {
    if (!line.formulaRuns || line.formulaRuns.length <= 1) return [line];
    return line.formulaRuns.map((run) => ({
      left: run.left,
      right: run.right ?? run.left,
      top: run.centerY - run.fontSize / 2,
      bottom: run.centerY + run.fontSize / 2,
      fontSize: run.fontSize,
      text: run.text,
      kind: "formula" as const,
    }));
  });
  const maxFont = Math.max(...fragments.map((line) => line.fontSize), 0.001);
  const bodyFontFloor = maxFont * 0.65;
  const sorted = [...fragments].sort(
    (a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left,
  );
  const rows: LineBoxInternal[][] = [];
  for (const line of sorted) {
    const row = rows[rows.length - 1];
    const rowFont = row
      ? Math.max(...row.map((item) => item.fontSize), line.fontSize)
      : line.fontSize;
    const rowCenter = row
      ? row.reduce((sum, item) => sum + lineCenterY(item), 0) / row.length
      : lineCenterY(line);
    if (!row || Math.abs(lineCenterY(line) - rowCenter) > rowFont * 0.45) {
      rows.push([line]);
    } else {
      row.push(line);
    }
  }
  return rows.some((row) => {
    const ordered = row.sort((a, b) => a.left - b.left);
    return ordered.some((line, index) => {
      const previous = ordered[index - 1];
      return previous &&
        previous.fontSize >= bodyFontFloor &&
        line.fontSize >= bodyFontFloor
        ? line.left - previous.right > maxGap
        : false;
    });
  });
}

function shouldMergeInlineFormulaAffix(
  left: LineBoxInternal,
  right: LineBoxInternal,
): boolean {
  if (isEquationNumberLine(left.text) || isEquationNumberLine(right.text)) {
    return false;
  }
  const fontSize = Math.max(left.fontSize, right.fontSize, 1);
  if (Math.abs(lineCenterY(left) - lineCenterY(right)) > fontSize * 0.45) {
    return false;
  }
  const gap = right.left - left.right;
  if (gap < -fontSize * 0.4 || gap > fontSize * 1.35) return false;

  const leftText = left.text.trim();
  const rightText = right.text.trim();
  const leftCanTakeFactorial =
    (left.kind === "formula" ||
      isBareMathAtomText(leftText) ||
      isCompactMathAtom(leftText)) &&
    !/[=<>≤≥≈≠+\-−*/×÷]$/.test(leftText) &&
    !/[=<>≤≥≈≠]/.test(leftText) &&
    proseWordCount(leftText) === 0 &&
    cjkCharCount(leftText) === 0;
  if (rightText === "!" && leftCanTakeFactorial) {
    return true;
  }
  if (isNumericTickText(leftText) && isFormulaUnitText(rightText)) {
    return true;
  }
  if (isNumericTickText(rightText) && isFormulaUnitText(leftText)) {
    return true;
  }
  if (
    isBareMathAtomText(leftText) &&
    right.kind === "formula" &&
    /^[=<>≤≥≈≠+\-−]/.test(rightText)
  ) {
    return true;
  }
  if (
    left.kind === "formula" &&
    /[=<>≤≥≈≠+\-−]$/.test(leftText) &&
    isBareMathAtomText(rightText)
  ) {
    return true;
  }
  return false;
}

function mergeInlineFormulaAffixes(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] {
  const merged: LineBoxInternal[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    const current = lines[index];
    const next = lines[index + 1];
    if (next && shouldMergeInlineFormulaAffix(current, next)) {
      const mergedKind =
        current.kind === "formula" || next.kind === "formula"
          ? "formula"
          : "text";
      const mergedLine =
        mergedKind === "formula"
          ? paragraphBoxFromLines([current, next])
          : inlineTextBoxFromLines([current, next]);
      const mergedLatex =
        mergedKind === "formula"
          ? formulaLatexFromLines([current, next])
          : undefined;
      merged.push({
        ...mergedLine,
        text:
          mergedKind === "formula"
            ? readableFormulaTextFromLatex(mergedLatex, mergedLine.text)
            : mergedLine.text,
        kind: mergedKind,
        latex: mergedLatex,
        formulaRuns: undefined,
      });
      index += 1;
    } else {
      merged.push(current);
    }
  }
  return merged;
}

function lineBoxForCorruptedZeroStackedFraction(
  pair: readonly [LineBoxInternal, LineBoxInternal],
): LineBoxInternal {
  const dominant = pair.reduce(
    (a, b) => (b.fontSize > a.fontSize ? b : a),
    pair[0],
  );
  const left = Math.min(...pair.map((line) => line.left));
  const right = Math.max(...pair.map((line) => line.right));
  const top = Math.min(...pair.map((line) => line.top));
  const bottom = Math.max(...pair.map((line) => line.bottom));
  return {
    left,
    right,
    top,
    bottom,
    fontSize: dominant.fontSize,
    text: "0/0",
    fontFamily: dominant.fontFamily,
    displayFontFamily: dominant.displayFontFamily,
    sourceFontName: dominant.sourceFontName,
    bold: dominant.bold,
    italic: dominant.italic,
    color: dominant.color,
    kind: "formula",
    latex: "\\frac{0}{0}",
    formulaRuns: pair.map(lineAsFormulaRun),
  };
}

function mergeCorruptedZeroStackedFractionFragments(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, LineBoxInternal>();

  lines.forEach((line, index) => {
    if (consumed.has(index) || !isCorruptedZeroGlyphLine(line)) return;

    const fontSize = Math.max(line.fontSize, 1);
    let match: {
      index: number;
      pair: [LineBoxInternal, LineBoxInternal];
    } | null = null;
    for (let cursor = index + 1; cursor < lines.length; cursor += 1) {
      if (consumed.has(cursor)) continue;
      const candidate = lines[cursor];
      if (lineCenterY(candidate) - lineCenterY(line) > fontSize * 4) break;
      const pair = corruptedZeroStackedFractionPair(line, candidate);
      if (!pair) continue;
      match = { index: cursor, pair };
      break;
    }

    if (!match) return;
    replacements.set(index, lineBoxForCorruptedZeroStackedFraction(match.pair));
    consumed.add(match.index);
  });

  return lines
    .map((line, index) => replacements.get(index) ?? line)
    .filter((_, index) => !consumed.has(index));
}

function isShortInlineMathFragmentText(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (isWeakScalarMetricLine(trimmed)) return true;
  if (isCompactMathAtom(trimmed) || isBareMathAtomText(trimmed)) return true;
  if (isShortInlineMathAtomSequenceText(trimmed)) return true;
  if (trimmed.length <= 20 && proseWordCount(trimmed) === 0) {
    return /^[A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]\s*(?:=|<|>|≤|≥|≈|≠)\s*[+\-−]?(?:\d+(?:\.\d+)?|[A-Za-zαβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ])$/.test(
      trimmed,
    );
  }
  return false;
}

function isInlineProseAbsorbableMath(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (isBareMathAtomText(text) && proseWordCount(text) === 0) return true;
  if (isShortInlineMathAtomSequenceText(text)) return true;
  return line.kind === "formula" && isShortInlineMathFragmentText(text);
}

function isInlineProseAbsorbableEmphasis(line: LineBoxInternal): boolean {
  const text = line.text.trim();
  if (!text || line.kind === "formula" || isEquationNumberLine(text)) {
    return false;
  }
  if (line.bold) return false;
  if (
    !line.italic &&
    !/italic|oblique/i.test(
      `${line.fontFamily ?? ""} ${line.sourceFontName ?? ""}`,
    )
  ) {
    return false;
  }
  return proseWordCount(text) > 0 && proseWordCount(text) <= 4;
}

function isInlineProseAbsorbableSegment(line: LineBoxInternal): boolean {
  return (
    isInlineProseAbsorbableMath(line) || isInlineProseAbsorbableEmphasis(line)
  );
}

function isInlineProseLeadFragmentText(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (/^(?:由|根据|利用|通过|其中|令|设|求)$/.test(trimmed)) {
    return true;
  }
  return /^(?:where|using|given|according|from|by|via)$/i.test(trimmed);
}

function sameInlineRow(a: LineBoxInternal, b: LineBoxInternal): boolean {
  const fontSize = Math.max(a.fontSize, b.fontSize, 1);
  return Math.abs(lineCenterY(a) - lineCenterY(b)) <= fontSize * 0.45;
}

function inlineFragmentGapOk(
  left: LineBoxInternal,
  right: LineBoxInternal,
): boolean {
  const fontSize = Math.max(left.fontSize, right.fontSize, 1);
  const gap = right.left - left.right;
  return gap >= -fontSize * 0.4 && gap <= fontSize * 1.35;
}

/** Whether two line boxes are set at the same angle. */
function sameRunAngle(a: LineBoxInternal, b: LineBoxInternal): boolean {
  return Math.abs((a.angle ?? 0) - (b.angle ?? 0)) < UPRIGHT_ANGLE_EPSILON;
}

function canContinueInlineProseRun(
  run: readonly LineBoxInternal[],
  next: LineBoxInternal,
): boolean {
  if (run.length === 0 || isEquationNumberLine(next.text)) return false;
  const previous = run[run.length - 1];
  if (isEquationNumberLine(previous.text)) return false;
  // Two runs set at different angles are never the same paragraph, however
  // close their rectangles land: once each run's rect is placed by its own
  // frame, a watermark and the text it crosses can share a centre.
  if (!sameRunAngle(previous, next)) return false;
  if (!sameInlineRow(previous, next) || !inlineFragmentGapOk(previous, next)) {
    return false;
  }

  const previousSegment = isInlineProseAbsorbableSegment(previous);
  const nextSegment = isInlineProseAbsorbableSegment(next);
  const runHasInlineSegment = run.some(isInlineProseAbsorbableSegment);
  const combinedText = joinInlineTextFragments(
    [...run, next].map((line) => line.text),
  );
  const combinedIsShortInlineProseMath =
    isShortProseWithInlineMathText(combinedText);
  const canCarryProseLead =
    run.some((line) => isInlineProseLeadFragmentText(line.text)) &&
    (proseWordCount(next.text) > 0 || cjkCharCount(next.text) > 0);
  if (
    !previousSegment &&
    !nextSegment &&
    !runHasInlineSegment &&
    !combinedIsShortInlineProseMath &&
    !canCarryProseLead
  ) {
    return false;
  }

  const nextHasProse =
    proseWordCount(next.text) > 0 || cjkCharCount(next.text) > 0;
  if (!nextSegment && !nextHasProse) return false;
  if (
    !sameRunStyle(previous, next) &&
    !previousSegment &&
    !nextSegment &&
    !combinedIsShortInlineProseMath &&
    !canCarryProseLead
  ) {
    return false;
  }
  if (combinedIsShortInlineProseMath) return true;
  return true;
}

function shouldJoinInlineTextWithoutSpace(
  previous: string,
  next: string,
): boolean {
  if (!previous || !next) return false;
  if (/^[,.;:!?，。；：、！？）)\]}]/.test(next)) return true;
  if (/[（([{]$/.test(previous)) return true;
  if (/[，。；：、！？]$/.test(previous)) return true;
  return false;
}

function joinInlineTextFragments(fragments: readonly string[]): string {
  let text = "";
  for (const fragment of fragments) {
    const next = fragment.trim();
    if (!next) continue;
    if (!text) {
      text = next;
      continue;
    }
    text += shouldJoinInlineTextWithoutSpace(text, next) ? next : ` ${next}`;
  }
  return text.replace(/\s+/g, " ").trim();
}

function shouldJoinFormulaTextWithoutSpace(
  previous: string,
  next: string,
): boolean {
  if (shouldJoinInlineTextWithoutSpace(previous, next)) return true;
  const previousText = previous.trimEnd();
  const nextText = next.trimStart();
  if (!previousText || !nextText) return false;
  if (/^[_^]/.test(nextText)) return true;
  if (
    /^[([{]/.test(nextText) &&
    new RegExp(
      `[${MATH_ALPHANUMERIC_CLASS}${UNICODE_SCRIPT_CLASS})\\]}]$`,
      "u",
    ).test(previousText)
  ) {
    return true;
  }
  return false;
}

function joinFormulaTextFragments(fragments: readonly string[]): string {
  let text = "";
  for (const fragment of fragments) {
    const next = fragment.trim();
    if (!next) continue;
    if (!text) {
      text = next;
      continue;
    }
    text += shouldJoinFormulaTextWithoutSpace(text, next) ? next : ` ${next}`;
  }
  return text.replace(/\s+/g, " ").trim();
}

function inlineTextFromLines(lines: readonly LineBoxInternal[]): string {
  const fragments = [...lines]
    .sort((a, b) => a.left - b.left)
    .map((line) => line.text.trim())
    .filter(Boolean);
  return joinInlineTextFragments(fragments);
}

function projectedRunsForLine(
  line: LineBoxInternal,
  text = line.text,
): ProjectedTextStyleRun[] {
  if (
    line.textRuns &&
    textFromPdfTextStyleRuns(line.textRuns) === line.text &&
    text === line.text
  ) {
    return line.textRuns.map((run) => ({ ...run }));
  }
  return [
    {
      text,
      fontSize: line.fontSize,
      fontFamily: line.fontFamily,
      displayFontFamily: line.displayFontFamily,
      sourceFontName: line.sourceFontName,
      bold: line.bold,
      italic: line.italic,
      color: line.color,
    },
  ];
}

function sliceProjectedRuns(
  runs: readonly ProjectedTextStyleRun[],
  start: number,
  end: number,
): ProjectedTextStyleRun[] {
  const out: ProjectedTextStyleRun[] = [];
  let cursor = 0;
  for (const run of runs) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;
    if (runEnd <= start || runStart >= end) continue;
    const localStart = Math.max(start, runStart) - runStart;
    const localEnd = Math.min(end, runEnd) - runStart;
    const text = run.text.slice(localStart, localEnd);
    if (text) out.push({ ...run, text });
  }
  return mergePdfTextStyleRuns(out);
}

function projectedRunsForTrimmedLine(
  line: LineBoxInternal,
): ProjectedTextStyleRun[] {
  const start = line.text.length - line.text.trimStart().length;
  const end = line.text.trimEnd().length;
  return sliceProjectedRuns(projectedRunsForLine(line), start, end);
}

function inlineTextAndRunsFromLines(lines: readonly LineBoxInternal[]): {
  text: string;
  runs: ProjectedTextStyleRun[];
} {
  let text = "";
  const runs: ProjectedTextStyleRun[] = [];
  for (const line of [...lines].sort((a, b) => a.left - b.left)) {
    const fragment = line.text.trim();
    if (!fragment) continue;
    if (text && !shouldJoinInlineTextWithoutSpace(text, fragment)) {
      const previous = runs[runs.length - 1] ?? projectedRunsForLine(line)[0];
      runs.push({ ...previous, text: " " });
      text += " ";
    }
    text += fragment;
    runs.push(...projectedRunsForTrimmedLine(line));
  }
  return { text, runs: mergePdfTextStyleRuns(runs) };
}

function inlineTextBoxFromLines(
  lines: readonly LineBoxInternal[],
): LineBoxInternal {
  const dominant = lines.reduce(
    (a, b) => (b.fontSize > a.fontSize ? b : a),
    lines[0],
  );
  const inline = inlineTextAndRunsFromLines(lines);
  return {
    left: Math.min(...lines.map((line) => line.left)),
    right: Math.max(...lines.map((line) => line.right)),
    top: Math.min(...lines.map((line) => line.top)),
    bottom: Math.max(...lines.map((line) => line.bottom)),
    fontSize: dominant.fontSize,
    text: inline.text,
    fontFamily: dominant.fontFamily,
    displayFontFamily: dominant.displayFontFamily,
    sourceFontName: dominant.sourceFontName,
    bold: dominant.bold,
    italic: dominant.italic,
    color: dominant.color,
    textRuns: inline.runs,
    kind: "text",
  };
}

function mergeInlineProseTextRuns(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] {
  const merged: LineBoxInternal[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    const run = [lines[index]];
    let cursor = index + 1;
    while (
      cursor < lines.length &&
      canContinueInlineProseRun(run, lines[cursor])
    ) {
      run.push(lines[cursor]);
      cursor += 1;
    }

    const hasInlineSegment = run.some(isInlineProseAbsorbableSegment);
    const hasProse = run.some(
      (line) => proseWordCount(line.text) > 0 || cjkCharCount(line.text) > 0,
    );
    if (run.length > 1 && hasInlineSegment && hasProse) {
      merged.push(inlineTextBoxFromLines(run));
      index = cursor - 1;
    } else {
      merged.push(...run);
      index = cursor - 1;
    }
  }
  return merged;
}

function lineHasProseText(line: LineBoxInternal): boolean {
  return proseWordCount(line.text) > 0 || cjkCharCount(line.text) > 0;
}

function isOrphanPunctuationFragmentLine(line: LineBoxInternal): boolean {
  if (line.kind === "formula" || line.explicitLatex?.trim()) return false;
  const compact = line.text.trim().replace(/\s+/g, "");
  if (!compact || compact.length > 3) return false;
  if (
    isEquationNumberLine(compact) ||
    isFootnoteMarkerText(compact) ||
    isVisualFractionBarText(compact)
  ) {
    return false;
  }
  return /^[!?,.;:，。！？；：、"'“”‘’]+$/u.test(compact);
}

function isTrailingPunctuationFragment(text: string): boolean {
  return /^[,.;:!?，。！？；：、]+$/u.test(text.trim());
}

function isFormulaTerminalPunctuationFragment(text: string): boolean {
  return /^[,.;:，。；：、]+$/u.test(text.trim());
}

function isLeadingPunctuationFragment(text: string): boolean {
  return /^["'“‘]+$/u.test(text.trim());
}

function attachFormulaTerminalPunctuation(
  formula: LineBoxInternal,
  punctuation: LineBoxInternal,
): LineBoxInternal {
  const lines = [formula, punctuation].sort((a, b) => a.left - b.left);
  const latex = formulaLatexFromLines(lines);
  const merged = paragraphBoxFromLines(lines);
  return {
    ...merged,
    text: readableFormulaTextFromLatex(latex, inlineTextFromLines(lines)),
    kind: "formula",
    latex,
    explicitLatex: formula.explicitLatex?.trim() ? latex : undefined,
    formulaRuns: formulaRunsFromLines(lines),
    equationAnchored: formula.equationAnchored,
  };
}

function canAttachTrailingPunctuationToFormula(
  previous: LineBoxInternal,
  punctuation: LineBoxInternal,
): boolean {
  if (previous.kind !== "formula") return false;
  if (!sameInlineRow(previous, punctuation)) return false;
  if (!isFormulaTerminalPunctuationFragment(punctuation.text)) return false;
  const fontSize = Math.max(previous.fontSize, punctuation.fontSize, 1);
  const gap = lineFragmentGap(previous, punctuation);
  return gap >= -fontSize * 0.4 && gap <= fontSize * 1.5;
}

function lineFragmentGap(
  left: LineBoxInternal,
  right: LineBoxInternal,
): number {
  return right.left - left.right;
}

function canAttachTrailingPunctuationToProse(
  previous: LineBoxInternal,
  punctuation: LineBoxInternal,
): boolean {
  if (!lineHasProseText(previous)) return false;
  if (!sameInlineRow(previous, punctuation)) return false;
  if (!isTrailingPunctuationFragment(punctuation.text)) return false;
  const fontSize = Math.max(previous.fontSize, punctuation.fontSize, 1);
  const gap = lineFragmentGap(previous, punctuation);
  return gap >= -fontSize * 0.4 && gap <= fontSize * 1.5;
}

function canAttachLeadingPunctuationToProse(
  punctuation: LineBoxInternal,
  next: LineBoxInternal,
): boolean {
  if (!lineHasProseText(next)) return false;
  if (!sameInlineRow(punctuation, next)) return false;
  if (!isLeadingPunctuationFragment(punctuation.text)) return false;
  const fontSize = Math.max(punctuation.fontSize, next.fontSize, 1);
  const gap = lineFragmentGap(punctuation, next);
  return gap >= -fontSize * 0.4 && gap <= fontSize * 1.5;
}

function lineIsNearContent(
  line: LineBoxInternal,
  neighbor: LineBoxInternal | undefined,
): boolean {
  if (!neighbor) return false;
  const fontSize = Math.max(line.fontSize, neighbor.fontSize, 1);
  if (sameInlineRow(line, neighbor)) {
    const gap =
      line.left <= neighbor.left
        ? lineFragmentGap(line, neighbor)
        : lineFragmentGap(neighbor, line);
    return gap >= -fontSize * 0.5 && gap <= fontSize * 4;
  }
  const verticalOverlap =
    Math.min(line.bottom, neighbor.bottom) - Math.max(line.top, neighbor.top);
  if (verticalOverlap <= fontSize * 0.12) return false;
  return (
    Math.abs(line.right - neighbor.left) <= fontSize * 4 ||
    Math.abs(neighbor.right - line.left) <= fontSize * 4 ||
    (line.left < neighbor.right && line.right > neighbor.left)
  );
}

function mergeOrDropOrphanPunctuationFragments(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] {
  const merged: LineBoxInternal[] = [];

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    if (!isOrphanPunctuationFragmentLine(line)) {
      merged.push(line);
      continue;
    }

    const previous = merged[merged.length - 1];
    if (previous && canAttachTrailingPunctuationToProse(previous, line)) {
      merged[merged.length - 1] = inlineTextBoxFromLines([previous, line]);
      continue;
    }

    if (previous && canAttachTrailingPunctuationToFormula(previous, line)) {
      merged[merged.length - 1] = attachFormulaTerminalPunctuation(
        previous,
        line,
      );
      continue;
    }

    const next = lines[index + 1];
    if (next && canAttachLeadingPunctuationToProse(line, next)) {
      merged.push(inlineTextBoxFromLines([line, next]));
      index += 1;
      continue;
    }

    if (lineIsNearContent(line, previous) || lineIsNearContent(line, next)) {
      continue;
    }

    merged.push(line);
  }

  return merged;
}

type EquationFormulaSide = "before" | "after";

function lineIsOnEquationFormulaSide(
  line: LineBoxInternal,
  equationNumber: LineBoxInternal,
  side: EquationFormulaSide,
  fontSize: number,
): boolean {
  const pad = fontSize * 0.35;
  return side === "before"
    ? line.right <= equationNumber.left - pad
    : line.left >= equationNumber.right + pad;
}

function formulaClusterGapToEquationNumber(
  cluster: readonly { line: LineBoxInternal }[],
  equationNumber: LineBoxInternal,
  side: EquationFormulaSide,
): number {
  const clusterLeft = Math.min(...cluster.map(({ line }) => line.left));
  const clusterRight = Math.max(...cluster.map(({ line }) => line.right));
  return side === "before"
    ? equationNumber.left - clusterRight
    : clusterLeft - equationNumber.right;
}

function formulaClusterInsertionIndex(
  cluster: readonly { line: LineBoxInternal; index: number }[],
  side: EquationFormulaSide,
  equationIndex: number,
): number {
  const candidates =
    side === "after"
      ? cluster.filter(({ index }) => index > equationIndex)
      : [...cluster];
  const ranked = (candidates.length > 0 ? candidates : [...cluster]).sort(
    (a, b) => b.line.fontSize - a.line.fontSize || a.index - b.index,
  );
  return ranked[0]?.index ?? Math.min(...cluster.map(({ index }) => index));
}

function verticalGapToCluster(
  line: LineBoxInternal,
  cluster: readonly LineBoxInternal[],
): number {
  const top = Math.min(...cluster.map((item) => item.top));
  const bottom = Math.max(...cluster.map((item) => item.bottom));
  if (line.bottom < top) return top - line.bottom;
  if (line.top > bottom) return line.top - bottom;
  return 0;
}

function hasCloserEquationNumber(
  line: LineBoxInternal,
  equationNumbers: readonly { line: LineBoxInternal; index: number }[],
  equationIndex: number,
  equationNumber: LineBoxInternal,
): boolean {
  const distanceToThisNumber = Math.abs(
    lineCenterY(line) - lineCenterY(equationNumber),
  );
  return equationNumbers.some(
    ({ line: otherNumber, index: otherIndex }) =>
      otherIndex !== equationIndex &&
      Math.abs(lineCenterY(line) - lineCenterY(otherNumber)) <
        distanceToThisNumber,
  );
}

function expandNumberedFormulaCluster(
  seedCluster: readonly { line: LineBoxInternal; index: number }[],
  lines: readonly LineBoxInternal[],
  equationNumbers: readonly { line: LineBoxInternal; index: number }[],
  equationIndex: number,
  equationNumber: LineBoxInternal,
  side: EquationFormulaSide,
  consumed: ReadonlySet<number>,
  fontSize: number,
): Array<{ line: LineBoxInternal; index: number }> {
  const selected = new Map<number, { line: LineBoxInternal; index: number }>();
  for (const item of seedCluster) selected.set(item.index, item);

  let changed = true;
  while (changed) {
    changed = false;
    const clusterLines = [...selected.values()].map(({ line }) => line);
    const clusterFont = Math.max(
      fontSize,
      ...clusterLines.map((line) => line.fontSize),
    );

    lines.forEach((line, index) => {
      if (
        index === equationIndex ||
        consumed.has(index) ||
        selected.has(index)
      ) {
        return;
      }
      if (isEquationNumberLine(line.text)) return;
      if (!lineIsOnEquationFormulaSide(line, equationNumber, side, fontSize)) {
        return;
      }
      if (
        hasColumnFormulaBarrierBetweenLineAndEquationNumber(
          line,
          equationNumber,
          lines,
          fontSize,
        )
      ) {
        return;
      }
      if (
        hasCloserEquationNumber(
          line,
          equationNumbers,
          equationIndex,
          equationNumber,
        )
      ) {
        return;
      }
      if (!isDisplayFormulaFragment(line)) return;
      if (
        verticalGapToCluster(line, clusterLines) >
        Math.max(clusterFont * 1.55, 8)
      ) {
        return;
      }
      if (
        !lineOverlapsOrNests(clusterLines, line) &&
        !(
          clusterLines.some((clusterLine) =>
            sameInlineRow(clusterLine, line),
          ) && lineHorizontallyNearCluster(line, clusterLines, clusterFont)
        )
      ) {
        return;
      }
      if (lineLooksLikeTrailingFormulaFootnoteMarker(line, clusterLines)) {
        return;
      }
      if (!canAttachLowEvidenceFormulaFragment(line, clusterLines)) return;
      selected.set(index, { line, index });
      changed = true;
    });
  }

  return [...selected.values()];
}

function formulaTextFromVisualRuns(
  runs: readonly {
    left: number;
    centerY: number;
    fontSize: number;
    text: string;
  }[],
): string {
  const sorted = [...runs]
    .map((run) => ({ ...run, text: run.text.trim() }))
    .filter((run) => run.text.length > 0)
    .sort((a, b) => a.centerY - b.centerY || a.left - b.left);
  if (sorted.length === 0) return "";

  const rows: Array<typeof sorted> = [];
  for (const run of sorted) {
    const row = rows[rows.length - 1];
    if (!row) {
      rows.push([run]);
      continue;
    }

    const rowFont = Math.max(...row.map((item) => item.fontSize));
    const largerFont = Math.max(rowFont, run.fontSize);
    const smallerFont = Math.min(rowFont, run.fontSize);
    const rowCenter =
      row.reduce((sum, item) => sum + item.centerY, 0) / row.length;
    const centerDelta = Math.abs(run.centerY - rowCenter);
    const likelyScriptOffset =
      smallerFont <= largerFont * 0.72 &&
      centerDelta > Math.max(smallerFont * 0.42, largerFont * 0.24);
    const tallOperatorBody =
      row.length === 1 &&
      hasStrongMathOperator(row[0].text) &&
      !hasStrongMathOperator(run.text) &&
      run.left >= row[0].left &&
      run.fontSize >= rowFont * 0.5 &&
      centerDelta <= rowFont * 0.72;
    if (
      (centerDelta > largerFont * 0.42 && !tallOperatorBody) ||
      likelyScriptOffset
    ) {
      rows.push([run]);
      continue;
    }

    row.push(run);
  }

  return rows
    .map((row) =>
      joinFormulaTextFragments(
        row.sort((a, b) => a.left - b.left).map((run) => run.text),
      ),
    )
    .filter(Boolean)
    .join("\n")
    .trim();
}

function residualFormulaTextFromBoxes(
  boxes: readonly TextBox[],
  seed: TextBox,
): string {
  const seedCenter = seed.y + seed.height / 2;
  const text = formulaTextFromVisualRuns(
    boxes.map((box) => ({
      left: box.x,
      centerY: box.y + box.height / 2,
      fontSize: box.fontSize,
      text: box.text,
    })),
  );
  if (text) return text;
  return [...boxes]
    .sort((a, b) => {
      const dy =
        Math.abs(a.y + a.height / 2 - seedCenter) -
        Math.abs(b.y + b.height / 2 - seedCenter);
      return dy || a.x - b.x;
    })
    .map((box) => box.text.trim())
    .filter(Boolean)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function residualFormulaLatexFromBoxes(boxes: readonly TextBox[]): string {
  const lines = boxes.map(textBoxToLine);
  if (lines.some((line) => line.explicitLatex?.trim())) {
    return formulaLatexFromLines(lines);
  }
  const coordinateScale = 1000;
  const runs: PdfFormulaRun[] = boxes.map((box) => ({
    left: box.x * coordinateScale,
    right: (box.x + box.width) * coordinateScale,
    centerY: (box.y + box.height / 2) * coordinateScale,
    fontSize: box.fontSize * coordinateScale,
    text: box.text,
  }));
  const latex = latexFromPdfFormulaRuns(runs);
  if (latex) return latex;
  return latexFromPdfFormulaText(boxes.map((box) => box.text).join("\n"));
}

function mergeNumberedDisplayFormulaFragments(
  lines: readonly LineBoxInternal[],
  pageWidth: number,
): LineBoxInternal[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, LineBoxInternal>();
  const equationNumbers = lines
    .map((line, index) => ({ line, index }))
    .filter(({ line }) => isEquationNumberLine(line.text));

  equationNumbers.forEach(({ line: equationNumber, index: equationIndex }) => {
    if (!isEquationNumberLine(equationNumber.text)) return;

    const fontSize = Math.max(1, equationNumber.fontSize);
    const centerY = lineCenterY(equationNumber);
    const verticalWindow = fontSize * 2.4 + 4;
    const candidateClusters = (["before", "after"] as const).flatMap((side) => {
      const rowCandidatesRaw = lines
        .map((line, index) => ({ line, index }))
        .filter(({ line, index }) => {
          if (index === equationIndex || consumed.has(index)) return false;
          if (
            !lineIsOnEquationFormulaSide(line, equationNumber, side, fontSize)
          ) {
            return false;
          }
          if (
            hasColumnFormulaBarrierBetweenLineAndEquationNumber(
              line,
              equationNumber,
              lines,
              fontSize,
            )
          ) {
            return false;
          }
          if (Math.abs(lineCenterY(line) - centerY) > verticalWindow) {
            return false;
          }
          if (
            hasCloserEquationNumber(
              line,
              equationNumbers,
              equationIndex,
              equationNumber,
            )
          ) {
            return false;
          }
          return isDisplayFormulaFragment(line);
        });
      const rowCandidateLines = rowCandidatesRaw.map(({ line }) => line);
      const rowCandidates = rowCandidatesRaw.filter(
        ({ line }) =>
          !lineLooksLikeTrailingFormulaFootnoteMarker(line, rowCandidateLines),
      );

      if (
        rowCandidates.length < 2 &&
        !rowCandidates.some(({ line }) => isNumberedSingletonFormulaSeed(line))
      ) {
        return [];
      }

      const sortedCandidates = rowCandidates.sort(
        (a, b) => a.line.left - b.line.left || a.index - b.index,
      );
      const clusters: Array<Array<{ line: LineBoxInternal; index: number }>> =
        [];
      for (const candidate of sortedCandidates) {
        const cluster = clusters[clusters.length - 1];
        const clusterRight = cluster
          ? Math.max(...cluster.map(({ line }) => line.right))
          : 0;
        if (!cluster || candidate.line.left - clusterRight > fontSize * 4) {
          clusters.push([candidate]);
        } else {
          cluster.push(candidate);
        }
      }
      const formulaClusters = clusters
        .map(stripTrailingFormulaFootnoteMarkers)
        .filter((cluster) => cluster.length > 0);

      const sameRowNonFormula = lines.filter((line, index) => {
        if (index === equationIndex || consumed.has(index)) return false;
        if (rowCandidates.some((candidate) => candidate.index === index)) {
          return false;
        }
        if (
          !lineIsOnEquationFormulaSide(line, equationNumber, side, fontSize)
        ) {
          return false;
        }
        if (Math.abs(lineCenterY(line) - centerY) > verticalWindow) {
          return false;
        }
        return hasProseText(line.text);
      });

      return formulaClusters
        .filter(
          (cluster) =>
            cluster.length >= 2 ||
            cluster.some(({ line }) => isNumberedSingletonFormulaSeed(line)),
        )
        .filter((cluster) => hasNumberedFormulaClusterSeed(cluster))
        .map((cluster) =>
          expandNumberedFormulaCluster(
            cluster,
            lines,
            equationNumbers,
            equationIndex,
            equationNumber,
            side,
            consumed,
            fontSize,
          ),
        )
        .filter(
          (cluster) =>
            !hasLargeSameRowGap(
              cluster.map(({ line }) => line),
              fontSize * 10,
            ),
        )
        .filter(
          (cluster) =>
            !isMatrixClusterWithoutStructuralEvidence(
              cluster.map(({ line }) => line),
            ),
        )
        .filter(
          (cluster) =>
            !isSameRowIndependentRelationCluster(
              cluster.map(({ line }) => line),
            ),
        )
        .filter(
          (cluster) =>
            !isVerticalIndependentRelationCluster(
              cluster.map(({ line }) => line),
            ),
        )
        .filter((cluster) => {
          const clusterLeft = Math.min(...cluster.map(({ line }) => line.left));
          const clusterRight = Math.max(
            ...cluster.map(({ line }) => line.right),
          );
          const clusterTop = Math.min(...cluster.map(({ line }) => line.top));
          const clusterBottom = Math.max(
            ...cluster.map(({ line }) => line.bottom),
          );
          const gapToNumber = formulaClusterGapToEquationNumber(
            cluster,
            equationNumber,
            side,
          );
          if (
            gapToNumber < 0 ||
            gapToNumber > Math.max(fontSize * 16, pageWidth * 0.45)
          ) {
            return false;
          }
          return !sameRowNonFormula.some((line) => {
            const verticalOverlap =
              Math.min(line.bottom, clusterBottom) -
              Math.max(line.top, clusterTop);
            if (verticalOverlap <= fontSize * 0.15) return false;
            const nearHorizontally =
              Math.abs(line.right - clusterLeft) <= fontSize * 3 ||
              Math.abs(line.left - clusterRight) <= fontSize * 3 ||
              (line.left < clusterRight && line.right > clusterLeft);
            return nearHorizontally;
          });
        })
        .map((cluster) => ({ cluster, side }));
    });

    const bestCandidate = candidateClusters.sort((a, b) => {
      const gapDelta =
        formulaClusterGapToEquationNumber(a.cluster, equationNumber, a.side) -
        formulaClusterGapToEquationNumber(b.cluster, equationNumber, b.side);
      if (Math.abs(gapDelta) > fontSize * 0.5) return gapDelta;
      const aWidth =
        Math.max(...a.cluster.map(({ line }) => line.right)) -
        Math.min(...a.cluster.map(({ line }) => line.left));
      const bWidth =
        Math.max(...b.cluster.map(({ line }) => line.right)) -
        Math.min(...b.cluster.map(({ line }) => line.left));
      return bWidth - aWidth;
    })[0];
    const bestCluster = bestCandidate?.cluster;

    if (!bestCluster) return;

    const formulaLines = bestCluster
      .sort((a, b) => a.index - b.index)
      .map(({ line }) => line);
    const formulaLatex = formulaLatexFromLines(formulaLines);
    const formulaBox = {
      ...paragraphBoxFromLines(formulaLines),
      text: readableFormulaTextFromLineGroup(formulaLines, formulaLatex),
      kind: "formula" as const,
      latex: formulaLatex,
      equationAnchored: true,
    };

    const firstIndex = formulaClusterInsertionIndex(
      bestCluster,
      bestCandidate.side,
      equationIndex,
    );
    replacements.set(firstIndex, formulaBox);
    for (const { index } of bestCluster) {
      consumed.add(index);
    }
  });

  const merged: LineBoxInternal[] = [];
  lines.forEach((line, index) => {
    const replacement = replacements.get(index);
    if (replacement) {
      merged.push(replacement);
    }
    if (!consumed.has(index)) {
      merged.push(line);
    }
  });
  return merged;
}

function hasNearbyEquationNumber(
  line: LineBoxInternal,
  lines: readonly LineBoxInternal[],
  pageWidth: number,
): boolean {
  if (pageWidth <= 0) return false;
  const centerY = lineCenterY(line);
  const fontSize = Math.max(1, line.fontSize);
  const maxGap = Math.min(
    pageWidth * 0.35,
    Math.max(fontSize * 18, pageWidth * 0.12),
  );
  return lines.some((candidate) => {
    if (!isEquationNumberLine(candidate.text)) return false;
    if (Math.abs(lineCenterY(candidate) - centerY) > fontSize * 1.35) {
      return false;
    }
    if (
      hasColumnFormulaBarrierBetweenLineAndEquationNumber(
        line,
        candidate,
        lines,
        fontSize,
      )
    ) {
      return false;
    }
    if (candidate.left >= line.right) {
      return candidate.left - line.right <= maxGap;
    }
    if (line.left >= candidate.right) {
      return line.left - candidate.right <= maxGap;
    }
    return false;
  });
}

function hasColumnFormulaBarrierBetweenLineAndEquationNumber(
  line: LineBoxInternal,
  equationNumber: LineBoxInternal,
  lines: readonly LineBoxInternal[],
  fontSize: number,
): boolean {
  const rowWindow = Math.max(fontSize * 2.4, 8);
  const columnGap = fontSize * 10;
  const equationNumberCenterY = lineCenterY(equationNumber);

  // Cheap rejections first — see the note on the TextBox twin above. The row
  // window and the column geometry discard nearly every line in constant time;
  // `isDisplayFormulaFragment` is the expensive one and now runs only for the
  // handful of lines that actually sit in the gap.
  const isBarrier = (
    other: LineBoxInternal,
    inGap: (candidate: LineBoxInternal) => boolean,
  ) => {
    if (other === line || other === equationNumber) return false;
    if (Math.abs(lineCenterY(other) - equationNumberCenterY) > rowWindow) {
      return false;
    }
    if (!inGap(other)) return false;
    if (isEquationNumberLine(other.text)) return false;
    return isDisplayFormulaFragment(other);
  };

  if (equationNumber.left >= line.right) {
    return lines.some((other) =>
      isBarrier(
        other,
        (candidate) =>
          candidate.left >= line.right &&
          candidate.right <= equationNumber.left &&
          candidate.left - line.right > columnGap,
      ),
    );
  }
  if (line.left >= equationNumber.right) {
    return lines.some((other) =>
      isBarrier(
        other,
        (candidate) =>
          candidate.right <= line.left &&
          candidate.left >= equationNumber.right &&
          line.left - candidate.right > columnGap,
      ),
    );
  }
  return false;
}

function lineHorizontallyNearCluster(
  line: LineBoxInternal,
  cluster: readonly LineBoxInternal[],
  fontSize: number,
): boolean {
  const left = Math.min(...cluster.map((item) => item.left));
  const right = Math.max(...cluster.map((item) => item.right));
  return (
    Math.abs(line.right - left) <= fontSize * 8 ||
    Math.abs(line.left - right) <= fontSize * 8 ||
    (line.left < right && line.right > left)
  );
}

function hasFormulaClusterProseBlocker(
  cluster: readonly LineBoxInternal[],
  lines: readonly LineBoxInternal[],
  consumed: ReadonlySet<number>,
  clusterIndexes: ReadonlySet<number>,
): boolean {
  const left = Math.min(...cluster.map((line) => line.left));
  const right = Math.max(...cluster.map((line) => line.right));
  const top = Math.min(...cluster.map((line) => line.top));
  const bottom = Math.max(...cluster.map((line) => line.bottom));
  const fontSize = Math.max(...cluster.map((line) => line.fontSize), 1);

  return lines.some((line, index) => {
    if (consumed.has(index) || clusterIndexes.has(index)) return false;
    if (isEquationNumberLine(line.text)) return false;
    if (!hasProseText(line.text)) return false;
    if (isDisplayFormulaFragment(line)) return false;
    if (canAttachDanglingNamedOperatorBridge(line, cluster)) return false;
    const verticalOverlap =
      Math.min(line.bottom, bottom) - Math.max(line.top, top);
    if (verticalOverlap <= fontSize * 0.12) return false;
    return (
      Math.abs(line.right - left) <= fontSize * 4 ||
      Math.abs(line.left - right) <= fontSize * 4 ||
      (line.left < right && line.right > left)
    );
  });
}

function sameRowDisplayFormulaGapOk(
  left: LineBoxInternal,
  right: LineBoxInternal,
  pageWidth: number,
): boolean {
  const fontSize = Math.max(left.fontSize, right.fontSize, 1);
  const gap = right.left - left.right;
  if (gap < -fontSize * 0.4) return false;
  return gap <= Math.min(fontSize * 8, pageWidth * 0.28);
}

interface SplitCaseColumnRow {
  indexes: number[];
  line: LineBoxInternal;
  expressionLeft: number;
  conditionLeft: number;
  hasAssignmentPrefix: boolean;
  hasPiecewisePrefix: boolean;
  hasOtherwiseCondition: boolean;
}

function lineBoxForSplitCaseRow(
  lines: readonly LineBoxInternal[],
  text: string,
): LineBoxInternal {
  const dominant = lines.reduce(
    (a, b) => (b.fontSize > a.fontSize ? b : a),
    lines[0],
  );
  const left = Math.min(...lines.map((line) => line.left));
  const right = Math.max(...lines.map((line) => line.right));
  const top = Math.min(...lines.map((line) => line.top));
  const bottom = Math.max(...lines.map((line) => line.bottom));
  const fontSize = dominant.fontSize;
  return {
    left,
    right,
    top,
    bottom,
    fontSize,
    text,
    fontFamily: dominant.fontFamily,
    displayFontFamily: dominant.displayFontFamily,
    sourceFontName: dominant.sourceFontName,
    bold: dominant.bold,
    italic: dominant.italic,
    color: dominant.color,
    kind: "formula",
    latex: latexFromPdfFormulaText(text),
    formulaRuns: [
      {
        left,
        right,
        centerY: (top + bottom) / 2,
        fontSize,
        text,
      },
    ],
  };
}

function splitCaseColumnRowFromSameLine(
  row: readonly { line: LineBoxInternal; index: number }[],
): SplitCaseColumnRow | null {
  const ordered = [...row].sort((a, b) => a.line.left - b.line.left);
  if (ordered.length < 2) return null;

  for (
    let conditionIndex = ordered.length - 1;
    conditionIndex >= 1;
    conditionIndex -= 1
  ) {
    const conditionLine = ordered[conditionIndex].line;
    const condition = conditionLine.text.trim();
    if (!isCaseConditionFragmentText(condition)) continue;

    const expressionParts = ordered.slice(0, conditionIndex);
    const expressionText = expressionParts
      .map(({ line }) => line.text.trim())
      .filter(Boolean)
      .join(" ");
    if (!expressionText) continue;

    const assignment = splitCaseAssignmentFragment(expressionText);
    const prefix = assignment?.prefix ?? "";
    const expression = assignment?.expression || expressionText;
    if (!isSplitCaseExpressionFragmentText(expression)) continue;

    const rowText = prefix
      ? `${prefix} ${expression}, ${condition}`
      : `${expression}, ${condition}`;
    const rowLines = ordered
      .slice(0, conditionIndex + 1)
      .map(({ line }) => line);
    return {
      indexes: ordered.slice(0, conditionIndex + 1).map(({ index }) => index),
      line: lineBoxForSplitCaseRow(rowLines, rowText),
      expressionLeft: expressionParts[0].line.left,
      conditionLeft: conditionLine.left,
      hasAssignmentPrefix: Boolean(prefix),
      hasPiecewisePrefix: Boolean(
        prefix && caseAssignmentPrefixLooksPiecewise(prefix),
      ),
      hasOtherwiseCondition: /^(?:otherwise|否则|其他|其它)$/i.test(condition),
    };
  }

  return null;
}

function splitCaseRowsAlign(
  previous: SplitCaseColumnRow,
  next: SplitCaseColumnRow,
): boolean {
  const fontSize = Math.max(previous.line.fontSize, next.line.fontSize, 1);
  const verticalGap = lineCenterY(next.line) - lineCenterY(previous.line);
  if (verticalGap <= 0 || verticalGap > fontSize * 2.35) return false;
  if (Math.abs(previous.conditionLeft - next.conditionLeft) > fontSize * 2) {
    return false;
  }
  if (previous.hasAssignmentPrefix || next.hasAssignmentPrefix) return true;
  return (
    Math.abs(previous.expressionLeft - next.expressionLeft) <= fontSize * 2
  );
}

function mergeSplitCaseColumnLines(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] {
  const rowGroups: Array<Array<{ line: LineBoxInternal; index: number }>> = [];
  for (const [index, line] of lines.entries()) {
    const row = rowGroups[rowGroups.length - 1];
    const anchor = row?.[0]?.line;
    if (anchor && sameInlineRow(anchor, line)) {
      row.push({ line, index });
    } else {
      rowGroups.push([{ line, index }]);
    }
  }

  const candidates = rowGroups.map(splitCaseColumnRowFromSameLine);
  const consumed = new Set<number>();
  const replacements = new Map<number, LineBoxInternal>();

  for (let rowIndex = 0; rowIndex < candidates.length; rowIndex += 1) {
    const first = candidates[rowIndex];
    if (!first || first.indexes.some((index) => consumed.has(index))) continue;

    const cluster = [first];
    for (
      let nextIndex = rowIndex + 1;
      nextIndex < candidates.length;
      nextIndex += 1
    ) {
      const next = candidates[nextIndex];
      if (!next || next.indexes.some((index) => consumed.has(index))) break;
      if (!splitCaseRowsAlign(cluster[cluster.length - 1], next)) break;
      cluster.push(next);
    }

    if (cluster.length < 2) continue;
    if (!cluster.some((row) => row.hasPiecewisePrefix)) continue;
    if (
      !cluster.some((row) => !row.hasAssignmentPrefix) &&
      !cluster.some((row) => row.hasOtherwiseCondition)
    ) {
      continue;
    }

    for (const row of cluster) {
      const [firstIndex, ...restIndexes] = row.indexes;
      replacements.set(firstIndex, row.line);
      for (const index of restIndexes) consumed.add(index);
    }
  }

  return lines
    .map((line, index) => replacements.get(index) ?? line)
    .filter((_, index) => !consumed.has(index));
}

const formulaBoundaryText = memoizeTextPredicate((text: string): string => {
  return text.trim().replace(/\s+/g, " ");
});

function hasMathTypographyEvidence(text: string): boolean {
  return /[∆Δ∂∇αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}⁰¹²³⁴⁵⁶⁷⁸⁹ᵀᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ₀₁₂₃₄₅₆₇₈₉]/u.test(
    text,
  );
}

function startsWithFormulaContinuation(text: string): boolean {
  const compact = formulaBoundaryText(text).replace(/\s+/g, "");
  return /^[=<>≤≥≈≠+\-−×÷*/·,;:)\]}|]/.test(compact);
}

function isClearlyDanglingFormulaSourceText(text: string): boolean {
  const trimmed = formulaBoundaryText(text);
  if (
    !startsWithFormulaContinuation(trimmed) &&
    !endsWithFormulaContinuation(trimmed)
  ) {
    return false;
  }
  return hasNamedMathOperator(trimmed) || hasStrongMathOperator(trimmed);
}

function endsWithFormulaContinuation(text: string): boolean {
  const compact = formulaBoundaryText(text).replace(/\s+/g, "");
  return /(?:[=<>≤≥≈≠+\-−×÷*/·,;:([{|]|\\(?:to|left|right))$/.test(compact);
}

function isCompleteRelationFormulaText(text: string): boolean {
  const trimmed = formulaBoundaryText(text);
  if (!trimmed || formulaProseWordCount(trimmed) > 0) return false;
  if (!/[=<>≤≥≈≠]/.test(trimmed)) return false;
  if (
    startsWithFormulaContinuation(trimmed) ||
    endsWithFormulaContinuation(trimmed)
  ) {
    return false;
  }
  return /[^\s=<>≤≥≈≠]/.test(trimmed);
}

function hasGarbledNamedOperatorClusterText(text: string): boolean {
  const compact = formulaBoundaryText(text).replace(/\s+/g, "");
  return /[#@]/.test(compact) || /\^\(?[<>]/.test(compact);
}

function canBeClosedSameRowNamedOperatorMember(line: LineBoxInternal): boolean {
  const text = formulaBoundaryText(line.text);
  if (!text || isEquationNumberLine(text)) return false;
  if (cjkCharCount(text) > 0 || formulaProseWordCount(text) > 0) return false;
  if (isWeakScalarMetricLine(text) || isShortPlotLabelText(text)) return false;
  if (hasGarbledNamedOperatorClusterText(text)) return false;
  return (
    hasNamedMathOperator(text) ||
    isDisplayFormulaFragment(line) ||
    startsWithFormulaArgumentContinuation(text) ||
    startsWithFormulaContinuation(text) ||
    endsWithFormulaContinuation(text) ||
    isCompleteRelationFormulaText(text) ||
    hasMathTypographyEvidence(text)
  );
}

function lineHasClosedNamedOperatorEvidence(line: LineBoxInternal): boolean {
  const text = formulaBoundaryText(line.text);
  if (hasMathTypographyEvidence(text) || line.explicitLatex?.trim())
    return true;
  return line.kind === "formula" && !isLowConfidenceFormulaFragmentText(text);
}

function isDanglingNamedOperatorBridgeCandidate(
  line: LineBoxInternal,
): boolean {
  const text = formulaBoundaryText(line.text);
  if (!text || isEquationNumberLine(text)) return false;
  if (!isDanglingNamedOperatorFormulaFragmentText(text)) return false;
  if (!hasNamedMathOperator(text)) return false;
  if (cjkCharCount(text) > 0 || formulaProseWordCount(text) > 0) return false;
  if (isWeakScalarMetricLine(text) || isShortPlotLabelText(text)) return false;
  if (hasGarbledNamedOperatorClusterText(text)) return false;
  return (
    hasMathTypographyEvidence(text) ||
    Boolean(line.explicitLatex?.trim()) ||
    (line.kind === "formula" && !isLowConfidenceFormulaFragmentText(text))
  );
}

function clusterHasDanglingNamedOperatorBridgeContext(
  cluster: readonly LineBoxInternal[],
): boolean {
  if (cluster.length === 0) return false;
  const wrapped = cluster.map((line) => ({ line }));
  return (
    cluster.some(isVisualFractionBarLine) ||
    hasVisualFractionClusterSeed(wrapped) ||
    hasUnnumberedStackedFractionClusterSeed(wrapped) ||
    linesHavePositionedScriptFormulaGeometry(cluster) ||
    cluster.some(hasStructuralFormulaEvidence)
  );
}

function canAttachDanglingNamedOperatorBridge(
  line: LineBoxInternal,
  cluster: readonly LineBoxInternal[],
): boolean {
  if (!isDanglingNamedOperatorBridgeCandidate(line)) return false;
  return clusterHasDanglingNamedOperatorBridgeContext(cluster);
}

function isClosedSameRowNamedOperatorFormulaRun(
  lines: readonly LineBoxInternal[],
): boolean {
  if (lines.length < 2 || lines.length > 4) return false;
  const ordered = [...lines].sort(
    (a, b) => a.left - b.left || lineCenterY(a) - lineCenterY(b),
  );
  if (!ordered.some((line) => hasNamedMathOperator(line.text))) return false;
  if (!ordered.some(lineHasClosedNamedOperatorEvidence)) return false;
  if (!ordered.every(canBeClosedSameRowNamedOperatorMember)) return false;

  const text = joinFormulaTextFragments(ordered.map((line) => line.text));
  if (!text || !hasNamedMathOperator(text)) return false;
  if (cjkCharCount(text) > 0 || formulaProseWordCount(text) > 0) return false;
  if (hasGarbledNamedOperatorClusterText(text)) return false;
  if (!isCompleteRelationFormulaText(text)) return false;
  if (
    !ordered.some(
      (line, index) =>
        index > 0 && startsWithFormulaArgumentContinuation(line.text),
    )
  ) {
    return false;
  }
  return mathSignalScore(text) >= 4;
}

function relationLeftHandSideKey(text: string): string | null {
  const match = /^(.+?)[=<>≤≥≈≠]/.exec(formulaBoundaryText(text));
  const left = match?.[1]?.trim() ?? "";
  if (!left) return null;
  return left.replace(/\s+/g, "").toLocaleLowerCase();
}

function relationRightHandSideText(text: string): string {
  const match = /^.+?[=<>≤≥≈≠](.+)$/.exec(formulaBoundaryText(text));
  return match?.[1]?.trim() ?? "";
}

function relationDerivationSideKey(text: string): string | null {
  const compact = formulaBoundaryText(text)
    .replace(/\s+/g, "")
    .toLocaleLowerCase();
  if (!compact || startsWithFormulaContinuation(compact)) return null;
  if (endsWithFormulaContinuation(compact)) return null;
  if (formulaProseWordCount(compact) > 0) return null;
  return compact;
}

function repeatedLeftHandSideLooksLikeDerivation(
  lines: readonly LineBoxInternal[],
): boolean {
  const keys = lines.map((line) => relationLeftHandSideKey(line.text));
  if (keys.some((key) => !key)) return false;
  const unique = new Set(keys);
  if (unique.size !== 1) return false;
  const key = keys[0] ?? "";
  const lhsLooksStructured =
    key.length > 2 ||
    /[()[\]{}]/.test(key) ||
    /[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(key);
  if (lhsLooksStructured) return true;
  return lines.some((line) => {
    const rhs = relationRightHandSideText(line.text);
    return (
      hasStrongMathOperator(rhs) ||
      hasNamedMathOperator(rhs) ||
      mathSignalScore(rhs) >= 3
    );
  });
}

function chainedRelationLooksLikeDerivation(
  lines: readonly LineBoxInternal[],
): boolean {
  if (lines.length < 2) return false;
  return lines.slice(1).every((line, index) => {
    const previous = lines[index];
    const previousRight = relationDerivationSideKey(
      relationRightHandSideText(previous.text),
    );
    const nextLeft = relationDerivationSideKey(
      relationLeftHandSideKey(line.text) ?? "",
    );
    if (!previousRight || !nextLeft || previousRight !== nextLeft) {
      return false;
    }
    const nextRight = relationRightHandSideText(line.text);
    return (
      hasStrongMathOperator(nextRight) ||
      hasNamedMathOperator(nextRight) ||
      mathSignalScore(nextRight) >= 2 ||
      isCompleteRelationFormulaText(line.text)
    );
  });
}

function hasSelfContainedFormulaExpressionEvidence(text: string): boolean {
  const compact = formulaBoundaryText(text).replace(/\s+/g, "");
  if (!compact) return false;
  const hasOperand =
    /[A-Za-z0-9αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}]/u.test(
      compact,
    );
  if (!hasOperand) return false;
  const operand =
    "[A-Za-z0-9αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\\u{1d400}-\\u{1d7ff}\\])}]";
  const nextOperand =
    "[A-Za-z0-9αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\\u{1d400}-\\u{1d7ff}\\[({]";
  if (new RegExp(`${operand}[+\\-−*/×÷·]${nextOperand}`, "u").test(compact)) {
    return true;
  }
  if (
    /[∑∫√∏]/.test(compact) &&
    /[A-Za-z0-9αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\u{1d400}-\u{1d7ff}]/u.test(
      compact.replace(/[∑∫√∏]/g, ""),
    )
  ) {
    return true;
  }
  return (
    hasNamedMathOperator(text) ||
    hasCompactFormulaStructure(text) ||
    mathSignalScore(text) >= 2
  );
}

function isSelfContainedFormulaBoundaryLine(line: LineBoxInternal): boolean {
  const text = formulaBoundaryText(line.text);
  if (!text || formulaProseWordCount(text) > 0) return false;
  if (
    /\b(?:if|otherwise|when|for)\b|(?:若|当|如果|否则|其他|其它)/i.test(text)
  ) {
    return false;
  }
  if (
    startsWithFormulaContinuation(text) ||
    endsWithFormulaContinuation(text)
  ) {
    return false;
  }
  if (!isDisplayFormulaFragment(line) || isLowEvidenceFormulaFragment(line)) {
    return false;
  }
  return (
    isCompleteRelationFormulaText(text) ||
    (!isCaseLikeRowText(text) &&
      !isMatrixLikeRowText(text) &&
      hasSelfContainedFormulaExpressionEvidence(text))
  );
}

function sameRowFormulaBoundaryContinuesExpression(
  run: readonly LineBoxInternal[],
  next: LineBoxInternal,
): boolean {
  const previous = run[run.length - 1];
  if (!previous) return false;
  const previousText = previous.text.trim();
  const nextText = next.text.trim();
  if (!previousText || !nextText) return false;
  if (endsWithFormulaContinuation(previousText)) return true;
  if (startsWithFormulaContinuation(nextText)) return true;
  if (sameRowSplitNamedOperatorLooksLikeFormula(previousText, nextText)) {
    return true;
  }
  if (
    isCompleteRelationFormulaText(previousText) &&
    isCompleteRelationFormulaText(nextText)
  ) {
    return false;
  }
  if (
    isShortFormulaUnitFragmentText(previousText) ||
    isShortFormulaUnitFragmentText(nextText)
  ) {
    return true;
  }
  if (
    isSelfContainedFormulaBoundaryLine(previous) &&
    isSelfContainedFormulaBoundaryLine(next)
  ) {
    return false;
  }
  return true;
}

function isSameRowIndependentRelationCluster(
  lines: readonly LineBoxInternal[],
): boolean {
  if (lines.length < 2) return false;
  if (lines.some((line) => isCaseLikeRowText(line.text))) return false;
  const ordered = [...lines].sort((a, b) => a.left - b.left);
  const fontSize = Math.max(...ordered.map((line) => line.fontSize), 1);
  const first = ordered[0];
  if (
    !ordered.every(
      (line) =>
        Math.abs(lineCenterY(line) - lineCenterY(first)) <= fontSize * 0.45,
    )
  ) {
    return false;
  }
  return ordered.every((line) => isCompleteRelationFormulaText(line.text));
}

function isVerticalIndependentRelationCluster(
  lines: readonly LineBoxInternal[],
): boolean {
  if (lines.length < 2) return false;
  if (lines.some((line) => isCaseLikeRowText(line.text))) return false;
  const ordered = [...lines].sort(
    (a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left,
  );
  const fontSize = Math.max(...ordered.map((line) => line.fontSize), 1);
  if (!ordered.every((line) => isCompleteRelationFormulaText(line.text))) {
    return false;
  }
  if (!ordered.every((line) => line.fontSize >= fontSize * 0.85)) {
    return false;
  }
  if (
    repeatedLeftHandSideLooksLikeDerivation(ordered) ||
    chainedRelationLooksLikeDerivation(ordered)
  ) {
    return false;
  }
  return ordered.every((line, index) => {
    if (index === 0) return true;
    const previous = ordered[index - 1];
    if (sameInlineRow(previous, line)) return false;
    return lineCenterY(line) - lineCenterY(previous) <= fontSize * 2.2;
  });
}

function linesFormSelfContainedStructuredFormula(
  lines: readonly LineBoxInternal[],
): boolean {
  if (lines.length === 0) return false;
  if (linesHavePositionedScriptFormulaGeometry(lines)) return true;
  const latex = formulaLatexFromLines(lines).trim();
  if (!latex || /^\\begin\{aligned\}/.test(latex)) return false;
  return /(?:\^\{|_\{|\\frac|\\sqrt|\\sum|\\int|\\prod|\\lim)/.test(latex);
}

function linesHavePositionedScriptFormulaGeometry(
  lines: readonly LineBoxInternal[],
): boolean {
  for (let i = 0; i < lines.length; i += 1) {
    for (let j = i + 1; j < lines.length; j += 1) {
      if (linePairHasPositionedScriptFormulaGeometry(lines[i], lines[j])) {
        return true;
      }
    }
  }
  return false;
}

function linePairHasPositionedScriptFormulaGeometry(
  a: LineBoxInternal,
  b: LineBoxInternal,
): boolean {
  const body = a.fontSize >= b.fontSize ? a : b;
  const script = body === a ? b : a;
  if (script.fontSize > body.fontSize * 0.72) return false;
  if (
    body.kind !== "formula" &&
    !hasMathExpressionEvidence(body.text) &&
    !isDisplayFormulaFragment(body)
  ) {
    return false;
  }
  if (!isLowEvidenceFormulaFragment(script)) return false;
  const centerDelta = Math.abs(lineCenterY(body) - lineCenterY(script));
  if (centerDelta <= Math.max(script.fontSize * 0.42, body.fontSize * 0.24)) {
    return false;
  }
  const horizontalGap = Math.max(
    body.left - script.right,
    script.left - body.right,
    0,
  );
  return horizontalGap <= body.fontSize * 1.6;
}

function scriptFragmentHasCloserFormulaBody(
  script: LineBoxInternal,
  cluster: readonly LineBoxInternal[],
  lines: readonly LineBoxInternal[],
): boolean {
  if (!isLowEvidenceFormulaFragment(script)) return false;
  const clusterDistance = Math.min(
    ...cluster
      .filter((line) =>
        linePairHasPositionedScriptFormulaGeometry(line, script),
      )
      .map((line) => Math.abs(lineCenterY(line) - lineCenterY(script))),
  );
  if (!Number.isFinite(clusterDistance)) return false;

  const fontSize = Math.max(script.fontSize, 1);
  return lines.some((line) => {
    if (line === script || cluster.includes(line)) return false;
    if (isEquationNumberLine(line.text)) return false;
    if (isWeakScalarMetricLine(line.text) || isNumericTickText(line.text)) {
      return false;
    }
    if (!linePairHasPositionedScriptFormulaGeometry(line, script)) {
      return false;
    }
    return (
      Math.abs(lineCenterY(line) - lineCenterY(script)) <
      clusterDistance - fontSize * 0.2
    );
  });
}

function crossesIndependentUnnumberedFormulaBoundary(
  cluster: readonly LineBoxInternal[],
  next: LineBoxInternal,
): boolean {
  if (
    cluster.some(lineContainsCaseLikeFormulaRow) ||
    lineContainsCaseLikeFormulaRow(next)
  ) {
    return false;
  }
  if (!isIndependentFormulaBoundaryLine(next)) return false;
  if (!linesFormSelfContainedStructuredFormula(cluster)) return false;

  const fontSize = Math.max(1, ...cluster.map((line) => line.fontSize));
  const clusterBottomCenter = Math.max(...cluster.map(lineCenterY));
  if (lineCenterY(next) <= clusterBottomCenter + fontSize * 0.45) {
    return false;
  }

  const ordered = [...cluster].sort(
    (a, b) => lineCenterY(a) - lineCenterY(b) || a.left - b.left,
  );
  const previous = ordered[ordered.length - 1];
  if (!previous) return false;
  if (endsWithFormulaContinuation(previous.text)) return false;
  if (startsWithFormulaContinuation(next.text)) return false;
  return true;
}

function lineContainsCaseLikeFormulaRow(line: LineBoxInternal): boolean {
  if (/\\begin\{cases\}/.test(line.latex ?? "")) return true;
  return line.text.split(/\r?\n/g).some((row) => isCaseLikeRowText(row.trim()));
}

function isIndependentFormulaBoundaryLine(line: LineBoxInternal): boolean {
  const text = formulaBoundaryText(line.text);
  if (isCompleteRelationFormulaText(text)) return true;
  if (
    startsWithFormulaContinuation(text) ||
    endsWithFormulaContinuation(text)
  ) {
    return false;
  }
  if (formulaProseWordCount(text) > 0) return false;
  return isSelfContainedFormulaBoundaryLine(line);
}

function canAttachSameRowDisplayFormulaFragment(
  line: LineBoxInternal,
  cluster: readonly LineBoxInternal[],
): boolean {
  if (!isLowEvidenceFormulaFragment(line)) return true;
  if (!cluster.some(hasStructuralFormulaEvidence)) return false;
  if (isWeakScalarMetricLine(line.text)) return false;

  const text = line.text.trim();
  if (
    isBareMathAtomText(text) ||
    isCompactMathAtom(text) ||
    isNumericTickText(text)
  ) {
    return true;
  }
  return lineOverlapsOrNests(cluster, line);
}

function orderedSameRowDisplayFormulaRun(
  run: readonly LineBoxInternal[],
  next: LineBoxInternal,
  pageWidth: number,
): LineBoxInternal[] | null {
  const ordered = [...run, next].sort(
    (a, b) => a.left - b.left || lineCenterY(a) - lineCenterY(b),
  );
  for (let index = 1; index < ordered.length; index += 1) {
    const previous = ordered[index - 1];
    const current = ordered[index];
    if (!sameInlineRow(previous, current)) return null;
    if (!sameRowDisplayFormulaGapOk(previous, current, pageWidth)) return null;
    if (
      !sameRowFormulaBoundaryContinuesExpression(
        ordered.slice(0, index),
        current,
      )
    ) {
      return null;
    }
  }
  return ordered;
}

function canContinueSameRowDisplayFormulaRun(
  run: readonly LineBoxInternal[],
  next: LineBoxInternal,
  lines: readonly LineBoxInternal[],
  pageWidth: number,
): boolean {
  if (run.length === 0) return false;
  if (isEquationNumberLine(next.text)) return false;
  if (next.equationAnchored || run.some((line) => line.equationAnchored)) {
    return false;
  }
  if (hasNearbyEquationNumber(next, lines, pageWidth)) return false;
  if (!isDisplayFormulaFragment(next)) return false;

  const orderedRun = orderedSameRowDisplayFormulaRun(run, next, pageWidth);
  if (!orderedRun) return false;
  const visualCluster = orderedRun.filter((line) => line !== next);
  return canAttachSameRowDisplayFormulaFragment(next, visualCluster);
}

function lineBoxForClosedSameRowNamedOperatorFormula(
  lines: readonly LineBoxInternal[],
): LineBoxInternal {
  const ordered = [...lines].sort(
    (a, b) => a.left - b.left || lineCenterY(a) - lineCenterY(b),
  );
  const latex = formulaLatexFromLines(ordered);
  return {
    ...paragraphBoxFromLines(ordered),
    text: readableFormulaTextFromLineGroup(ordered, latex),
    kind: "formula",
    latex,
    formulaRuns: formulaRunsFromLines(ordered),
  };
}

function canContinueClosedSameRowNamedOperatorRun(
  run: readonly LineBoxInternal[],
  next: LineBoxInternal,
  pageWidth: number,
): boolean {
  const previous = run[run.length - 1];
  if (!previous) return false;
  if (isEquationNumberLine(next.text) || next.equationAnchored) return false;
  if (!sameInlineRow(previous, next)) return false;
  if (!sameRowDisplayFormulaGapOk(previous, next, pageWidth)) return false;
  return canBeClosedSameRowNamedOperatorMember(next);
}

function mergeClosedSameRowNamedOperatorFormulaFragments(
  lines: readonly LineBoxInternal[],
  pageWidth: number,
): LineBoxInternal[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, LineBoxInternal>();

  for (let index = 0; index < lines.length; index += 1) {
    if (consumed.has(index)) continue;
    const seed = lines[index];
    if (!canBeClosedSameRowNamedOperatorMember(seed)) continue;
    if (!hasNamedMathOperator(seed.text)) continue;
    if (isEquationNumberLine(seed.text) || seed.equationAnchored) continue;

    const run = [seed];
    let cursor = index + 1;
    while (
      cursor < lines.length &&
      run.length < 4 &&
      !consumed.has(cursor) &&
      canContinueClosedSameRowNamedOperatorRun(run, lines[cursor], pageWidth)
    ) {
      run.push(lines[cursor]);
      if (isClosedSameRowNamedOperatorFormulaRun(run)) {
        replacements.set(
          index,
          lineBoxForClosedSameRowNamedOperatorFormula(run),
        );
        for (
          let consumedIndex = index + 1;
          consumedIndex <= cursor;
          consumedIndex += 1
        ) {
          consumed.add(consumedIndex);
        }
        index = cursor;
        break;
      }
      cursor += 1;
    }
  }

  return lines
    .map((line, index) => replacements.get(index) ?? line)
    .filter((_, index) => !consumed.has(index));
}

function isNamedOperatorLimitConditionLine(line: LineBoxInternal): boolean {
  const text = formulaBoundaryText(line.text);
  if (!text || isEquationNumberLine(text)) return false;
  if (hasProseText(text) || isVisualFractionBarText(text)) return false;
  if (isWeakScalarMetricLine(text) || isShortPlotLabelText(text)) return false;
  if (!hasMathTypographyEvidence(text)) return false;
  const compact = text.replace(/\s+/g, "");
  return /(?:→|->|\\to|∞|0)/.test(compact);
}

function lineCanServeAsBridgeLimit(
  line: LineBoxInternal,
  bridge: LineBoxInternal,
): boolean {
  if (!isNamedOperatorLimitConditionLine(line)) return false;
  if (line.fontSize > bridge.fontSize * 0.9) return false;
  if (lineCenterY(line) <= lineCenterY(bridge)) return false;
  return (
    lineCenterX(line) >= bridge.left - bridge.fontSize &&
    lineCenterX(line) <= bridge.right + bridge.fontSize
  );
}

function nearestFractionPartForBar(
  lines: readonly LineBoxInternal[],
  bar: LineBoxInternal,
  direction: "above" | "below",
): LineBoxInternal | null {
  const barCenter = lineCenterY(bar);
  return (
    lines
      .filter((line) => {
        if (line === bar || !isFractionPartLine(line)) return false;
        const centerY = lineCenterY(line);
        if (direction === "above" && centerY >= barCenter) return false;
        if (direction === "below" && centerY <= barCenter) return false;
        return (
          lineFitsFractionBarSpan(line, bar) &&
          lineIsCenteredOnFractionBar(line, bar)
        );
      })
      .sort(
        (a, b) =>
          Math.abs(lineCenterY(a) - barCenter) -
          Math.abs(lineCenterY(b) - barCenter),
      )[0] ?? null
  );
}

function formulaLineForDanglingNamedOperatorFractionBridge(
  bridge: LineBoxInternal,
  limit: LineBoxInternal,
  numerator: LineBoxInternal,
  bar: LineBoxInternal,
  denominator: LineBoxInternal,
): LineBoxInternal {
  const lines = [bridge, limit, numerator, bar, denominator];
  const bridgeLatex = latexFromPdfFormulaText(bridge.text).trim();
  const limitLatex = latexFromPdfFormulaText(limit.text).trim();
  const numeratorLatex = latexFromPdfFormulaText(numerator.text).trim();
  const denominatorLatex = latexFromPdfFormulaText(denominator.text).trim();
  const fractionLatex = `\\frac{${numeratorLatex}}{${denominatorLatex}}`;
  const latex =
    bridgeLatex && limitLatex && numeratorLatex && denominatorLatex
      ? `${bridgeLatex}_{${limitLatex}} ${fractionLatex}`
      : formulaLatexFromLines(lines);
  return {
    ...paragraphBoxFromLines(lines),
    text: readableFormulaTextFromLatex(latex, formulaTextFromLines(lines)),
    kind: "formula",
    latex,
    formulaRuns: formulaRunsFromLines(lines),
  };
}

function mergeDanglingNamedOperatorFractionBridges(
  lines: readonly LineBoxInternal[],
): LineBoxInternal[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, LineBoxInternal>();

  lines.forEach((bridge, bridgeIndex) => {
    if (consumed.has(bridgeIndex)) return;
    if (!isDanglingNamedOperatorBridgeCandidate(bridge)) return;

    const bridgeFont = Math.max(bridge.fontSize, 1);
    const bars = lines
      .map((line, index) => ({ line, index }))
      .filter(({ line, index }) => {
        if (index === bridgeIndex || consumed.has(index)) return false;
        if (!isVisualFractionBarLine(line)) return false;
        if (line.left < bridge.right - bridgeFont * 2) return false;
        if (line.left - bridge.right > bridgeFont * 8) return false;
        return (
          Math.abs(lineCenterY(line) - lineCenterY(bridge)) <= bridgeFont * 0.85
        );
      })
      .sort(
        (a, b) =>
          Math.abs(a.line.left - bridge.right) -
          Math.abs(b.line.left - bridge.right),
      );
    const barEntry = bars[0];
    if (!barEntry) return;

    const numerator = nearestFractionPartForBar(lines, barEntry.line, "above");
    const denominator = nearestFractionPartForBar(
      lines,
      barEntry.line,
      "below",
    );
    if (!numerator || !denominator) return;
    const numeratorIndex = lines.indexOf(numerator);
    const denominatorIndex = lines.indexOf(denominator);
    if (numeratorIndex < 0 || denominatorIndex < 0) return;
    if (consumed.has(numeratorIndex) || consumed.has(denominatorIndex)) return;

    const limitEntry = lines
      .map((line, index) => ({ line, index }))
      .filter(({ line, index }) => {
        if (
          index === bridgeIndex ||
          index === barEntry.index ||
          index === numeratorIndex ||
          index === denominatorIndex ||
          consumed.has(index)
        ) {
          return false;
        }
        return lineCanServeAsBridgeLimit(line, bridge);
      })
      .sort(
        (a, b) =>
          Math.abs(lineCenterY(a.line) - lineCenterY(bridge)) -
          Math.abs(lineCenterY(b.line) - lineCenterY(bridge)),
      )[0];
    if (!limitEntry) return;

    const merged = formulaLineForDanglingNamedOperatorFractionBridge(
      bridge,
      limitEntry.line,
      numerator,
      barEntry.line,
      denominator,
    );
    replacements.set(
      Math.min(
        bridgeIndex,
        limitEntry.index,
        numeratorIndex,
        barEntry.index,
        denominatorIndex,
      ),
      merged,
    );
    for (const index of [
      bridgeIndex,
      limitEntry.index,
      numeratorIndex,
      barEntry.index,
      denominatorIndex,
    ]) {
      if (
        index !==
        Math.min(
          bridgeIndex,
          limitEntry.index,
          numeratorIndex,
          barEntry.index,
          denominatorIndex,
        )
      ) {
        consumed.add(index);
      }
    }
  });

  return lines
    .map((line, index) => replacements.get(index) ?? line)
    .filter((_, index) => !consumed.has(index));
}

function mergeUnnumberedSameRowDisplayFormulaFragments(
  lines: readonly LineBoxInternal[],
  pageWidth: number,
): LineBoxInternal[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, LineBoxInternal>();

  for (let index = 0; index < lines.length; index += 1) {
    if (consumed.has(index)) continue;
    const seed = lines[index];
    if (!hasDisplayFormulaSeed(seed)) continue;
    if (isEquationNumberLine(seed.text)) continue;
    if (seed.equationAnchored) continue;
    if (hasNearbyEquationNumber(seed, lines, pageWidth)) continue;

    const run = [seed];
    let cursor = index + 1;
    while (
      cursor < lines.length &&
      !consumed.has(cursor) &&
      canContinueSameRowDisplayFormulaRun(run, lines[cursor], lines, pageWidth)
    ) {
      run.push(lines[cursor]);
      cursor += 1;
    }

    if (run.length < 2) continue;
    const cluster = run.map((line, offset) => ({
      line,
      index: index + offset,
    }));
    if (isDocumentIndexLeaderCluster(cluster)) continue;
    if (cluster.every(({ line }) => isWeakScalarMetricLine(line.text))) {
      continue;
    }
    if (isOnlyLowEvidenceFormulaCluster(cluster.map(({ line }) => line))) {
      continue;
    }
    if (!hasUnnumberedFormulaClusterSeed(cluster)) continue;
    if (
      cluster.some(({ line }) => isWeakScalarMetricLine(line.text)) &&
      cluster.some(({ line }) => isBareMathAtomText(line.text))
    ) {
      continue;
    }
    if (
      hasLargeSameRowGap(
        cluster.map(({ line }) => line),
        Math.max(seed.fontSize, 1) * 8,
      )
    ) {
      continue;
    }

    const clusterIndexes = new Set(cluster.map(({ index }) => index));
    if (
      hasFormulaClusterProseBlocker(
        cluster.map(({ line }) => line),
        lines,
        consumed,
        clusterIndexes,
      )
    ) {
      continue;
    }

    const formulaLines = cluster.map(({ line }) => line);
    const formulaLatex = formulaLatexFromLines(formulaLines);
    replacements.set(index, {
      ...paragraphBoxFromLines(formulaLines),
      text: readableFormulaTextFromLineGroup(formulaLines, formulaLatex),
      kind: "formula",
      latex: formulaLatex,
    });
    for (const { index: consumedIndex } of cluster.slice(1)) {
      consumed.add(consumedIndex);
    }
    index = cursor - 1;
  }

  return lines
    .map((line, index) => replacements.get(index) ?? line)
    .filter((_, index) => !consumed.has(index));
}

function expandUnnumberedFormulaCluster(
  seed: LineBoxInternal,
  seedIndex: number,
  lines: readonly LineBoxInternal[],
  consumed: ReadonlySet<number>,
  pageWidth: number,
): Array<{ line: LineBoxInternal; index: number }> {
  const selected = new Map<number, { line: LineBoxInternal; index: number }>();
  selected.set(seedIndex, { line: seed, index: seedIndex });

  let changed = true;
  while (changed) {
    changed = false;
    const clusterLines = [...selected.values()].map(({ line }) => line);
    const clusterFont = Math.max(
      1,
      ...clusterLines.map((line) => line.fontSize),
    );

    lines.forEach((line, index) => {
      if (index === seedIndex || consumed.has(index) || selected.has(index)) {
        return;
      }
      if (isEquationNumberLine(line.text)) return;
      if (line.equationAnchored) return;
      if (hasNearbyEquationNumber(line, lines, pageWidth)) return;
      const canAttachNamedOperatorBridge = canAttachDanglingNamedOperatorBridge(
        line,
        clusterLines,
      );
      if (!isDisplayFormulaFragment(line) && !canAttachNamedOperatorBridge) {
        return;
      }
      const stackedFractionCluster = [...clusterLines, line].map((item) => ({
        line: item,
      }));
      const canFormStackedFraction = hasUnnumberedStackedFractionClusterSeed(
        stackedFractionCluster,
      );
      if (
        verticalGapToCluster(line, clusterLines) >
          Math.max(clusterFont * 1.7, 10) &&
        !canFormStackedFraction &&
        !canAttachNamedOperatorBridge
      ) {
        return;
      }
      if (
        !lineOverlapsOrNests(clusterLines, line) &&
        !lineHorizontallyNearCluster(line, clusterLines, clusterFont)
      ) {
        return;
      }
      if (
        !canAttachLowEvidenceFormulaFragment(line, clusterLines) &&
        !canFormStackedFraction &&
        !canAttachNamedOperatorBridge
      ) {
        return;
      }
      if (
        !canFormStackedFraction &&
        !canAttachNamedOperatorBridge &&
        isDetachedShortFormulaLabel(line, clusterLines)
      ) {
        return;
      }
      if (lineLooksLikeTrailingFormulaFootnoteMarker(line, clusterLines)) {
        return;
      }
      const currentClusterLines = [...selected.values()].map(
        (item) => item.line,
      );
      const previous = currentClusterLines[currentClusterLines.length - 1];
      if (
        previous &&
        isSelfContainedFormulaBoundaryLine(previous) &&
        isSelfContainedFormulaBoundaryLine(line) &&
        !repeatedLeftHandSideLooksLikeDerivation([
          ...currentClusterLines,
          line,
        ]) &&
        !chainedRelationLooksLikeDerivation([...currentClusterLines, line])
      ) {
        return;
      }
      if (
        crossesIndependentUnnumberedFormulaBoundary(currentClusterLines, line)
      ) {
        return;
      }
      if (
        scriptFragmentHasCloserFormulaBody(line, currentClusterLines, lines)
      ) {
        return;
      }
      selected.set(index, { line, index });
      changed = true;
    });
  }

  return [...selected.values()];
}

function mergeUnnumberedDisplayFormulaFragments(
  lines: readonly LineBoxInternal[],
  pageWidth: number,
): LineBoxInternal[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, LineBoxInternal>();

  lines.forEach((seed, seedIndex) => {
    if (consumed.has(seedIndex)) return;
    if (
      !hasDisplayFormulaSeed(seed) &&
      !isVisualFractionBarLine(seed) &&
      !lineHasStackedFractionMathEvidence(seed)
    ) {
      return;
    }
    if (isEquationNumberLine(seed.text)) return;
    if (seed.equationAnchored) return;
    if (hasNearbyEquationNumber(seed, lines, pageWidth)) return;

    const fontSize = Math.max(1, seed.fontSize);
    const centerY = lineCenterY(seed);
    const cluster = expandUnnumberedFormulaCluster(
      seed,
      seedIndex,
      lines,
      consumed,
      pageWidth,
    );

    if (cluster.length < 2) return;
    if (isDocumentIndexLeaderCluster(cluster)) return;
    if (cluster.every(({ line }) => isWeakScalarMetricLine(line.text))) return;
    if (isOnlyLowEvidenceFormulaCluster(cluster.map(({ line }) => line))) {
      return;
    }
    if (
      clusterContainsDetachedShortFormulaLabel(cluster.map(({ line }) => line))
    ) {
      return;
    }
    if (isVerticalIndependentRelationCluster(cluster.map(({ line }) => line))) {
      return;
    }
    if (!hasUnnumberedFormulaClusterSeed(cluster)) return;
    if (
      cluster.some(({ line }) => isWeakScalarMetricLine(line.text)) &&
      cluster.some(({ line }) => isBareMathAtomText(line.text))
    ) {
      return;
    }
    if (
      !cluster.some(
        ({ line }) => Math.abs(lineCenterY(line) - centerY) > fontSize * 0.45,
      )
    ) {
      return;
    }
    if (
      hasLargeSameRowGap(
        cluster.map(({ line }) => line),
        fontSize * 10,
      )
    ) {
      return;
    }

    const clusterIndexes = new Set(cluster.map(({ index }) => index));
    if (
      hasFormulaClusterProseBlocker(
        cluster.map(({ line }) => line),
        lines,
        consumed,
        clusterIndexes,
      )
    ) {
      return;
    }

    const ordered = cluster.sort((a, b) => a.index - b.index);
    const firstIndex = ordered[0].index;
    replacements.set(firstIndex, {
      ...paragraphBoxFromLines(ordered.map(({ line }) => line)),
      kind: "formula",
      latex: formulaLatexFromLines(ordered.map(({ line }) => line)),
    });
    for (const { index } of ordered.slice(1)) consumed.add(index);
  });

  return lines
    .map((line, index) => replacements.get(index) ?? line)
    .filter((_, index) => !consumed.has(index));
}

function textBoxToLine(box: TextBox): LineBoxInternal {
  const formulaRun: PdfFormulaRun = {
    left: box.x,
    right: box.x + box.width,
    centerY: box.y + box.height / 2,
    fontSize: box.fontSize,
    text: box.text,
  };
  return {
    left: box.x,
    right: box.x + box.width,
    top: box.y,
    bottom: box.y + box.height,
    fontSize: box.fontSize,
    text: box.text,
    fontFamily: box.fontFamily,
    displayFontFamily: box.displayFontFamily,
    sourceFontName: box.sourceFontName,
    bold: box.bold,
    italic: box.italic,
    color: box.color,
    kind: box.kind,
    latex: box.latex,
    explicitLatex:
      box.kind === "formula" &&
      box.latex &&
      box.latex.trim() !== latexFromPdfFormulaText(box.text)
        ? box.latex
        : undefined,
    formulaRuns: box.kind === "formula" ? [formulaRun] : undefined,
  };
}

function textBoxCenterY(box: TextBox): number {
  return box.y + box.height / 2;
}

function textBoxRight(box: TextBox): number {
  return box.x + box.width;
}

function textBoxIsOnEquationFormulaSide(
  box: TextBox,
  equationNumber: TextBox,
  side: EquationFormulaSide,
  fontSize: number,
): boolean {
  const pad = fontSize * 0.35;
  return side === "before"
    ? textBoxRight(box) <= equationNumber.x - pad
    : box.x >= textBoxRight(equationNumber) + pad;
}

function hasColumnFormulaBarrierBetweenTextBoxAndEquationNumber(
  box: TextBox,
  equationNumber: TextBox,
  boxes: readonly TextBox[],
  fontSize: number,
): boolean {
  const rowWindow = Math.max(fontSize * 8, 0.03);
  const columnGap = fontSize * 10;
  const boxRight = textBoxRight(box);
  const equationNumberCenterY = textBoxCenterY(equationNumber);

  /**
   * Reject `other` on the cheap tests first.
   *
   * Every test here is a side-effect-free rejection, so their order does not
   * change the result — but it changes the cost by orders of magnitude.
   * `textBoxPairCoveredByFormulaContainer` scans all boxes, making this
   * already-O(n) callback O(n^2), and the caller runs it per candidate pair.
   * Geometry answers "different row" or "wrong side" in constant time for the
   * overwhelming majority of boxes, so the scan now runs only for the few
   * that are genuinely between the box and its equation number.
   */
  const isBarrier = (other: TextBox, inGap: (other: TextBox) => boolean) => {
    if (other === box || other === equationNumber) return false;
    if (Math.abs(textBoxCenterY(other) - equationNumberCenterY) > rowWindow) {
      return false;
    }
    if (!inGap(other)) return false;
    if (isEquationNumberLine(other.text)) return false;
    if (!isDisplayFormulaFragment(textBoxToLine(other))) return false;
    return !textBoxPairCoveredByFormulaContainer(box, other, boxes, fontSize);
  };

  if (equationNumber.x >= boxRight) {
    return boxes.some((other) =>
      isBarrier(
        other,
        (candidate) =>
          candidate.x >= boxRight &&
          textBoxRight(candidate) <= equationNumber.x &&
          candidate.x - boxRight > columnGap,
      ),
    );
  }
  if (box.x >= textBoxRight(equationNumber)) {
    const equationNumberRight = textBoxRight(equationNumber);
    return boxes.some((other) =>
      isBarrier(
        other,
        (candidate) =>
          textBoxRight(candidate) <= box.x &&
          candidate.x >= equationNumberRight &&
          box.x - textBoxRight(candidate) > columnGap,
      ),
    );
  }
  return false;
}

function textBoxPairCoveredByFormulaContainer(
  a: TextBox,
  b: TextBox,
  boxes: readonly TextBox[],
  fontSize: number,
): boolean {
  const left = Math.min(a.x, b.x);
  const right = Math.max(textBoxRight(a), textBoxRight(b));
  const top = Math.min(a.y, b.y);
  const bottom = Math.max(a.y + a.height, b.y + b.height);
  const pad = fontSize * 1.5;
  return boxes.some((container) => {
    if (container === a || container === b) return false;
    if (!isFormulaTextBoxKind(container)) return false;
    if (container.x > left + pad) return false;
    if (textBoxRight(container) < right - pad) return false;
    const overlap =
      Math.min(container.y + container.height, bottom) -
      Math.max(container.y, top);
    return overlap > fontSize * 0.2;
  });
}

function textBoxClusterGapToEquationNumber(
  cluster: readonly { box: TextBox }[],
  equationNumber: TextBox,
  side: EquationFormulaSide,
): number {
  const clusterLeft = Math.min(...cluster.map(({ box }) => box.x));
  const clusterRight = Math.max(...cluster.map(({ box }) => textBoxRight(box)));
  return side === "before"
    ? equationNumber.x - clusterRight
    : clusterLeft - textBoxRight(equationNumber);
}

function textBoxVerticalGapToCluster(
  box: TextBox,
  cluster: readonly TextBox[],
): number {
  const top = Math.min(...cluster.map((item) => item.y));
  const bottom = Math.max(...cluster.map((item) => item.y + item.height));
  if (box.y + box.height < top) return top - (box.y + box.height);
  if (box.y > bottom) return box.y - bottom;
  return 0;
}

function horizontallyNearTextBoxCluster(
  box: TextBox,
  left: number,
  right: number,
  fontSize: number,
): boolean {
  return (
    Math.abs(textBoxRight(box) - left) <= fontSize * 4 ||
    Math.abs(box.x - right) <= fontSize * 4 ||
    (box.x < right && textBoxRight(box) > left)
  );
}

function textBoxOverlapsOrNestsCluster(
  box: TextBox,
  cluster: readonly TextBox[],
): boolean {
  return lineOverlapsOrNests(cluster.map(textBoxToLine), textBoxToLine(box));
}

function textBoxLooksLikeTrailingFormulaFootnoteMarker(
  box: TextBox,
  formulaBoxes: readonly TextBox[],
): boolean {
  if (!isFootnoteMarkerText(box.text)) return false;
  const candidates = formulaBoxes.filter((candidate) => candidate !== box);
  if (candidates.length === 0) return false;

  const fontSize = Math.max(
    ...candidates.map((candidate) => candidate.fontSize),
    0.001,
  );
  if (box.fontSize > fontSize * 0.72) return false;

  const clusterRight = Math.max(
    ...candidates.map((candidate) => textBoxRight(candidate)),
  );
  const clusterCenterY =
    candidates.reduce((sum, candidate) => sum + textBoxCenterY(candidate), 0) /
    candidates.length;
  const trailingGap = box.x - clusterRight;
  if (trailingGap < box.fontSize * 0.9) return false;
  if (
    !isConservativeFootnoteMarkerText(box.text) &&
    trailingGap < fontSize * 1.2
  ) {
    return false;
  }
  return textBoxCenterY(box) < clusterCenterY - box.fontSize * 0.2;
}

function textBoxLooksLikeDetachedLowEvidenceSameRowLabel(
  box: TextBox,
  cluster: readonly TextBox[],
  fontSize: number,
): boolean {
  const line = textBoxToLine(box);
  if (!isLowEvidenceFormulaFragment(line)) return false;
  if (
    cluster.some((body) =>
      linePairHasPositionedScriptFormulaGeometry(textBoxToLine(body), line),
    )
  ) {
    return false;
  }
  const sameRow = cluster.filter((body) =>
    sameInlineRow(textBoxToLine(body), line),
  );
  if (sameRow.length === 0) return false;
  const gap = Math.min(
    ...sameRow.map((body) =>
      Math.max(body.x - textBoxRight(box), box.x - textBoxRight(body), 0),
    ),
  );
  return gap > fontSize * 1.35;
}

function hasCloserTextBoxEquationNumber(
  box: TextBox,
  equationNumbers: readonly { box: TextBox; index: number }[],
  equationIndex: number,
  equationNumber: TextBox,
): boolean {
  const distanceToThisNumber = Math.abs(
    textBoxCenterY(box) - textBoxCenterY(equationNumber),
  );
  return equationNumbers.some(
    ({ box: otherNumber, index: otherIndex }) =>
      otherIndex !== equationIndex &&
      Math.abs(textBoxCenterY(box) - textBoxCenterY(otherNumber)) <
        distanceToThisNumber,
  );
}

function expandResidualNumberedFormulaCluster(
  seedCluster: readonly { box: TextBox; index: number }[],
  boxes: readonly TextBox[],
  equationNumbers: readonly { box: TextBox; index: number }[],
  equationIndex: number,
  equationNumber: TextBox,
  side: EquationFormulaSide,
  consumed: ReadonlySet<number>,
  fontSize: number,
): Array<{ box: TextBox; index: number }> {
  const selected = new Map<number, { box: TextBox; index: number }>();
  for (const item of seedCluster) selected.set(item.index, item);

  let changed = true;
  while (changed) {
    changed = false;
    const clusterBoxes = [...selected.values()].map(({ box }) => box);
    const clusterFont = Math.max(
      fontSize,
      ...clusterBoxes.map((box) => box.fontSize),
    );
    const clusterLeft = Math.min(...clusterBoxes.map((box) => box.x));
    const clusterRight = Math.max(
      ...clusterBoxes.map((box) => textBoxRight(box)),
    );

    boxes.forEach((box, index) => {
      if (
        index === equationIndex ||
        consumed.has(index) ||
        selected.has(index)
      ) {
        return;
      }
      if (isEquationNumberLine(box.text)) return;
      if (
        !textBoxIsOnEquationFormulaSide(box, equationNumber, side, fontSize)
      ) {
        return;
      }
      if (
        hasColumnFormulaBarrierBetweenTextBoxAndEquationNumber(
          box,
          equationNumber,
          boxes,
          fontSize,
        )
      ) {
        return;
      }
      if (
        hasCloserTextBoxEquationNumber(
          box,
          equationNumbers,
          equationIndex,
          equationNumber,
        )
      ) {
        return;
      }
      if (!isDisplayFormulaFragment(textBoxToLine(box))) return;
      if (
        textBoxVerticalGapToCluster(box, clusterBoxes) >
        Math.max(clusterFont * 2.4, 0.02)
      ) {
        return;
      }
      if (
        !textBoxOverlapsOrNestsCluster(box, clusterBoxes) &&
        !horizontallyNearTextBoxCluster(
          box,
          clusterLeft,
          clusterRight,
          fontSize,
        )
      ) {
        return;
      }
      if (textBoxLooksLikeTrailingFormulaFootnoteMarker(box, clusterBoxes)) {
        return;
      }
      if (
        textBoxLooksLikeDetachedLowEvidenceSameRowLabel(
          box,
          clusterBoxes,
          clusterFont,
        )
      ) {
        return;
      }
      if (
        !canAttachLowEvidenceFormulaFragment(
          textBoxToLine(box),
          clusterBoxes.map(textBoxToLine),
        )
      ) {
        return;
      }
      selected.set(index, { box, index });
      changed = true;
    });
  }

  return [...selected.values()];
}

function hasStructuralResidualFormulaEvidence(box: TextBox): boolean {
  const line = textBoxToLine(box);
  const text = box.text.trim();
  if (!text) return false;
  if (isVisualFractionBarText(text)) return true;
  if (isCaseLikeRowText(text) || isMatrixLikeRowText(text)) return true;
  if (hasStrongMathOperator(text)) return true;
  if (hasCompactFormulaStructure(text)) return true;
  if (mathSignalScore(text) >= 3) return true;
  if (box.latex && box.latex.trim() !== latexFromPdfFormulaText(text)) {
    return true;
  }
  if (line.kind === "formula" && !isWeakScalarMetricLine(text)) {
    return (
      /[=<>≤≥≈≠]/.test(text) ||
      /[()[\]{}]/.test(text) ||
      isCompactMathAtom(text)
    );
  }
  return false;
}

function isWeakResidualFormulaCluster(
  cluster: readonly { box: TextBox; index: number }[],
): boolean {
  const boxes = cluster.map(({ box }) => box);
  if (boxes.length === 0) return true;
  if (
    isDocumentIndexLeaderCluster(
      boxes.map((box) => ({ line: textBoxToLine(box) })),
    )
  ) {
    return true;
  }
  if (boxes.every((box) => isWeakScalarMetricLine(box.text))) {
    return true;
  }
  if (boxes.some(hasStructuralResidualFormulaEvidence)) {
    return false;
  }
  return boxes.every(
    (box) => isWeakScalarMetricLine(box.text) || isShortPlotLabelText(box.text),
  );
}

function textBoxOverlapArea(a: TextBox, b: TextBox): number {
  const left = Math.max(a.x, b.x);
  const right = Math.min(textBoxRight(a), textBoxRight(b));
  const top = Math.max(a.y, b.y);
  const bottom = Math.min(a.y + a.height, b.y + b.height);
  return Math.max(0, right - left) * Math.max(0, bottom - top);
}

function textBoxArea(box: TextBox): number {
  return Math.max(0, box.width) * Math.max(0, box.height);
}

function textBoxCenterInside(box: TextBox, container: TextBox): boolean {
  const centerX = box.x + box.width / 2;
  const centerY = box.y + box.height / 2;
  return (
    centerX >= container.x &&
    centerX <= textBoxRight(container) &&
    centerY >= container.y &&
    centerY <= container.y + container.height
  );
}

function comparableMathGlyphTokens(text: string): Set<string> {
  const tokens = new Set<string>();
  for (const raw of Array.from(text)) {
    for (const ch of Array.from(raw.normalize("NFKC"))) {
      if (ch === "∆" || ch === "Δ") {
        tokens.add("Delta");
        continue;
      }
      if (/[A-Za-z0-9]/.test(ch)) {
        tokens.add(ch.toLowerCase());
        continue;
      }
      if (/[\u{1d400}-\u{1d7ff}αβγδεζηθκλμνξπρστφχψωΓΘΛΞΠΣΦΨΩ∂∇]/u.test(ch)) {
        tokens.add(ch);
      }
    }
  }
  return tokens;
}

function escapeRegexLiteral(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function formulaVisibleTextContainsUnit(
  formulaVisibleText: string,
  residualUnit: string,
): boolean {
  const compactFormulaUnits = formulaVisibleText
    .replace(/\s+/g, "")
    .replace(/／/g, "/")
    .toLowerCase();
  if (residualUnit.length !== 1) {
    return compactFormulaUnits.includes(residualUnit);
  }
  const unit = escapeRegexLiteral(residualUnit);
  const readableFormulaText = formulaVisibleText
    .replace(/／/g, "/")
    .toLowerCase();
  return new RegExp(
    `(?:\\d\\s*${unit}\\b|${unit}\\s*/|/\\s*${unit}\\b)`,
    "i",
  ).test(readableFormulaText);
}

function formulaCoversResidualText(
  formula: TextBox,
  residual: TextBox,
): boolean {
  const formulaVisibleText = [
    formula.text,
    formula.latex ? fallbackTextFromPdfFormulaLatex(formula.latex) : "",
  ]
    .filter(Boolean)
    .join(" ");
  const residualUnit = formulaUnitCompactText(residual.text);
  if (residualUnit) {
    return formulaVisibleTextContainsUnit(formulaVisibleText, residualUnit);
  }
  const residualTokens = comparableMathGlyphTokens(residual.text);
  if (residualTokens.size === 0) return false;
  const formulaTokens = comparableMathGlyphTokens(formulaVisibleText);
  for (const token of residualTokens) {
    if (!formulaTokens.has(token)) return false;
  }
  return true;
}

function formulaCoveredResidualOverlapThreshold(box: TextBox): number {
  if (formulaUnitCompactText(box.text)) return 0.55;
  const tokenCount = comparableMathGlyphTokens(box.text).size;
  if (tokenCount <= 1 || isShortPlotLabelText(box.text)) return 0.9;
  if (tokenCount === 2) return 0.7;
  return 0.55;
}

function formulaCoveredResidualCenterThreshold(box: TextBox): number | null {
  if (formulaUnitCompactText(box.text)) return 0;
  const tokenCount = comparableMathGlyphTokens(box.text).size;
  if (tokenCount <= 1 || isShortPlotLabelText(box.text)) return null;
  if (tokenCount === 2) return 0.45;
  return 0.35;
}

function isFormulaCoveredResidualBox(box: TextBox): boolean {
  if (isFormulaTextBoxKind(box)) {
    return false;
  }
  const line = textBoxToLine(box);
  const text = box.text.trim();
  if (!text) return false;
  if (hasProseText(text) && !isFormulaUnitText(text)) return false;
  return (
    isDisplayFormulaFragment(line) ||
    mathSignalScore(text) > 0 ||
    isBareMathAtomText(text) ||
    comparableMathGlyphTokens(text).size >= 2
  );
}

export function pruneFormulaCoveredResidualTextBoxes(
  boxes: readonly TextBox[],
): TextBox[] {
  const formulaBoxes = boxes.filter((box) => isFormulaTextBoxKind(box));
  if (formulaBoxes.length === 0) return [...boxes];

  return boxes.filter((box) => {
    if (!isFormulaCoveredResidualBox(box)) return true;
    const area = textBoxArea(box);
    if (area <= 0) return true;
    return !formulaBoxes.some((formula) => {
      if (!formulaCoversResidualText(formula, box)) return false;
      const overlapRatio = textBoxOverlapArea(box, formula) / area;
      const overlapThreshold = formulaCoveredResidualOverlapThreshold(box);
      if (overlapRatio >= overlapThreshold) return true;
      const centerThreshold = formulaCoveredResidualCenterThreshold(box);
      return (
        centerThreshold !== null &&
        textBoxCenterInside(box, formula) &&
        overlapRatio >= centerThreshold
      );
    });
  });
}

/**
 * Browser pdf.js can expose a few superscripts/subscripts as independent text
 * boxes even after line-level formula grouping. Run a conservative second pass
 * over final boxes near right-aligned equation numbers so editable PDF mode does
 * not leave formulas as many tiny text fields.
 */
export function mergeResidualNumberedFormulaTextBoxes(
  boxes: readonly TextBox[],
): TextBox[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, TextBox>();
  const equationNumbers = boxes
    .map((box, index) => ({ box, index }))
    .filter(({ box }) => isEquationNumberLine(box.text));

  for (const { box: equationNumber, index: equationIndex } of equationNumbers) {
    const fontSize = Math.max(equationNumber.fontSize, 0.006);
    const centerY = textBoxCenterY(equationNumber);
    const verticalWindow = Math.max(fontSize * 8, 0.12);
    const candidateClusters = (["before", "after"] as const).flatMap((side) => {
      const rawCandidates = boxes
        .map((box, index) => ({ box, index }))
        .filter(({ box, index }) => {
          if (index === equationIndex || consumed.has(index)) return false;
          if (
            !textBoxIsOnEquationFormulaSide(box, equationNumber, side, fontSize)
          ) {
            return false;
          }
          if (
            hasColumnFormulaBarrierBetweenTextBoxAndEquationNumber(
              box,
              equationNumber,
              boxes,
              fontSize,
            )
          ) {
            return false;
          }
          if (Math.abs(textBoxCenterY(box) - centerY) > verticalWindow) {
            return false;
          }
          if (
            hasCloserTextBoxEquationNumber(
              box,
              equationNumbers,
              equationIndex,
              equationNumber,
            )
          ) {
            return false;
          }
          return isDisplayFormulaFragment(textBoxToLine(box));
        });
      const candidateBoxes = rawCandidates.map(({ box }) => box);
      const candidates = rawCandidates.filter(
        ({ box }) =>
          !textBoxLooksLikeTrailingFormulaFootnoteMarker(box, candidateBoxes),
      );

      const seed = candidates
        .filter(({ box }) => {
          const line = textBoxToLine(box);
          return hasDisplayFormulaSeed(line) || isVisualFractionBarLine(line);
        })
        .sort((a, b) =>
          side === "before"
            ? textBoxRight(b.box) - textBoxRight(a.box)
            : a.box.x - b.box.x,
        )[0];
      if (!seed) return [];

      const cluster = expandResidualNumberedFormulaCluster(
        [seed],
        boxes,
        equationNumbers,
        equationIndex,
        equationNumber,
        side,
        consumed,
        fontSize,
      );
      if (cluster.length < 2) return [];
      if (isWeakResidualFormulaCluster(cluster)) return [];
      if (
        hasLargeSameRowGap(
          cluster.map(({ box }) => textBoxToLine(box)),
          fontSize * 10,
        )
      ) {
        return [];
      }
      if (
        isMatrixClusterWithoutStructuralEvidence(
          cluster.map(({ box }) => textBoxToLine(box)),
        )
      ) {
        return [];
      }
      if (
        isSameRowIndependentRelationCluster(
          cluster.map(({ box }) => textBoxToLine(box)),
        )
      ) {
        return [];
      }
      if (
        !hasNumberedFormulaClusterSeed(
          cluster.map(({ box }) => ({ line: textBoxToLine(box) })),
        )
      ) {
        return [];
      }
      const gapToNumber = textBoxClusterGapToEquationNumber(
        cluster,
        equationNumber,
        side,
      );
      if (gapToNumber < 0 || gapToNumber > Math.max(fontSize * 20, 0.36)) {
        return [];
      }
      return [{ cluster, side, seed }];
    });

    const best = candidateClusters.sort((a, b) => {
      const gapDelta =
        textBoxClusterGapToEquationNumber(a.cluster, equationNumber, a.side) -
        textBoxClusterGapToEquationNumber(b.cluster, equationNumber, b.side);
      if (Math.abs(gapDelta) > fontSize * 0.5) return gapDelta;
      const aWidth =
        Math.max(...a.cluster.map(({ box }) => textBoxRight(box))) -
        Math.min(...a.cluster.map(({ box }) => box.x));
      const bWidth =
        Math.max(...b.cluster.map(({ box }) => textBoxRight(box))) -
        Math.min(...b.cluster.map(({ box }) => box.x));
      return bWidth - aWidth;
    })[0];
    if (!best) continue;
    const { cluster, seed, side } = best;

    const clusterLeft = Math.min(...cluster.map(({ box }) => box.x));
    const clusterRight = Math.max(
      ...cluster.map(({ box }) => textBoxRight(box)),
    );
    const clusterTop = Math.min(...cluster.map(({ box }) => box.y));
    const clusterBottom = Math.max(
      ...cluster.map(({ box }) => box.y + box.height),
    );
    const proseBlocker = boxes.some((box, index) => {
      if (index === equationIndex || consumed.has(index)) return false;
      if (cluster.some((item) => item.index === index)) return false;
      if (
        !textBoxIsOnEquationFormulaSide(box, equationNumber, side, fontSize)
      ) {
        return false;
      }
      if (!hasProseText(box.text)) return false;
      const overlap =
        Math.min(box.y + box.height, clusterBottom) -
        Math.max(box.y, clusterTop);
      if (overlap <= fontSize * 0.15) return false;
      return horizontallyNearTextBoxCluster(
        box,
        clusterLeft,
        clusterRight,
        fontSize,
      );
    });
    if (proseBlocker) continue;

    const ordered = cluster.sort((a, b) => a.index - b.index);
    const firstIndex = ordered[0].index;
    const seedBox = seed.box;
    const visualText = residualFormulaTextFromBoxes(
      ordered.map(({ box }) => box),
      seedBox,
    );
    const latex = residualFormulaLatexFromBoxes(ordered.map(({ box }) => box));
    replacements.set(firstIndex, {
      ...ordered[0].box,
      x: clusterLeft,
      y: clusterTop,
      width: Math.max(0, clusterRight - clusterLeft),
      height: Math.max(0, clusterBottom - clusterTop),
      text: readableFormulaTextFromLatex(latex, visualText),
      fontSize: Math.max(...ordered.map(({ box }) => box.fontSize)),
      fontFamily: seedBox.fontFamily ?? ordered[0].box.fontFamily,
      displayFontFamily:
        seedBox.displayFontFamily ?? ordered[0].box.displayFontFamily,
      sourceFontName: seedBox.sourceFontName ?? ordered[0].box.sourceFontName,
      bold: seedBox.bold ?? ordered[0].box.bold,
      italic: seedBox.italic ?? ordered[0].box.italic,
      color: seedBox.color ?? ordered[0].box.color,
      kind: "formula",
      latex,
    });
    for (const { index } of ordered.slice(1)) consumed.add(index);
  }

  const merged = boxes
    .map((box, index) => replacements.get(index) ?? box)
    .filter((_, index) => !consumed.has(index));
  return demoteDetachedLowEvidenceFormulaLabelsNearNumberedFormulas(
    splitMergedTrailingLowEvidenceFormulaLabels(merged),
  );
}

function splitMergedTrailingLowEvidenceFormulaLabels(
  boxes: readonly TextBox[],
): TextBox[] {
  const equationNumbers = boxes.filter((box) => isEquationNumberLine(box.text));
  if (equationNumbers.length === 0) return [...boxes];

  return boxes.flatMap((box) => {
    if (!isFormulaTextBoxKind(box)) return [box];
    const match =
      /^(.*[=<>≤≥≈≠].*[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ\]\)])\s+([A-Za-z])$/u.exec(
        box.text.trim(),
      );
    if (!match) return [box];
    const [, prefix, label] = match;
    if (!isShortPlotLabelText(label) || !isFormulaLikeText(prefix))
      return [box];
    const sameRowEquation = equationNumbers.some(
      (equationNumber) =>
        Math.abs(textBoxCenterY(equationNumber) - textBoxCenterY(box)) <=
        Math.max(box.fontSize * 2, 0.02),
    );
    if (!sameRowEquation) return [box];

    const totalChars = Math.max(1, Array.from(box.text.trim()).length);
    const labelChars = Math.max(1, Array.from(label).length);
    const labelWidth = Math.min(
      box.width * 0.28,
      Math.max(box.fontSize * 0.7, box.width * (labelChars / totalChars)),
    );
    const prefixWidth = Math.max(box.width - labelWidth, box.width * 0.5);
    return [
      {
        ...box,
        width: prefixWidth,
        text: prefix.trim(),
        latex: latexFromPdfFormulaText(prefix.trim()),
      },
      {
        ...box,
        id: `${box.id}-label`,
        x: box.x + prefixWidth,
        width: Math.max(0, box.width - prefixWidth),
        text: label,
        kind: "text",
        latex: undefined,
      },
    ];
  });
}

function demoteDetachedLowEvidenceFormulaLabelsNearNumberedFormulas(
  boxes: readonly TextBox[],
): TextBox[] {
  const equationNumbers = boxes.filter((box) => isEquationNumberLine(box.text));
  if (equationNumbers.length === 0) return [...boxes];

  return boxes.map((box) => {
    if (!isFormulaTextBoxKind(box)) return box;
    const line = textBoxToLine(box);
    if (!isLowEvidenceFormulaFragment(line)) return box;
    const fontSize = Math.max(box.fontSize, 0.006);
    const hasCompetingFormula = equationNumbers.some((equationNumber) => {
      if (
        Math.abs(textBoxCenterY(box) - textBoxCenterY(equationNumber)) >
        fontSize * 2
      ) {
        return false;
      }
      return boxes.some((other) => {
        if (other === box || other === equationNumber) return false;
        if (!isFormulaTextBoxKind(other)) return false;
        if (isLowEvidenceFormulaFragment(textBoxToLine(other))) return false;
        if (
          Math.abs(textBoxCenterY(other) - textBoxCenterY(box)) >
          fontSize * 0.75
        ) {
          return false;
        }
        const gap = Math.max(
          other.x - textBoxRight(box),
          box.x - textBoxRight(other),
          0,
        );
        return gap > fontSize * 1.35;
      });
    });
    if (!hasCompetingFormula) return box;
    return { ...box, kind: "text", latex: undefined };
  });
}

function zeroQuantityUnitPart(
  text: string,
): { quantity: string; unit: string } | null {
  const compact = text.trim().replace(/\s+/g, "");
  const unit = formulaUnitSuffix(compact);
  if (!unit) return null;
  const quantity = compact.slice(0, -unit.length);
  if (quantity !== "0") return null;
  return { quantity, unit };
}

function textBoxHorizontallyNear(
  box: TextBox,
  left: number,
  right: number,
  fontSize: number,
): boolean {
  return (
    textBoxRight(box) >= left - fontSize * 4 && box.x <= right + fontSize * 4
  );
}

function textBoxVerticallyNear(
  box: TextBox,
  top: number,
  bottom: number,
  fontSize: number,
): boolean {
  return (
    box.y + box.height >= top - fontSize * 3 && box.y <= bottom + fontSize * 3
  );
}

function zeroResidualBoxesLookStacked(
  boxes: readonly TextBox[],
  fontSize: number,
): boolean {
  if (boxes.length < 2) return false;
  const ordered = [...boxes].sort(
    (a, b) => textBoxCenterY(a) - textBoxCenterY(b),
  );
  for (let index = 0; index < ordered.length - 1; index += 1) {
    const above = ordered[index];
    const below = ordered[index + 1];
    const verticalDistance = textBoxCenterY(below) - textBoxCenterY(above);
    if (verticalDistance < fontSize * 0.75 || verticalDistance > fontSize * 4) {
      continue;
    }
    const centerDelta = Math.abs(
      above.x + above.width / 2 - (below.x + below.width / 2),
    );
    if (
      centerDelta <=
      Math.max(fontSize * 1.2, Math.max(above.width, below.width))
    ) {
      return true;
    }
  }
  return false;
}

function formulaBoxLooksLikeCorruptedZeroFraction(box: TextBox): boolean {
  if (!isFormulaTextBoxKind(box)) return false;
  const source = `${box.text} ${box.latex ?? ""}`;
  return /!!|_\{\s*!!\s*\}/.test(source);
}

function variableFromCorruptedZeroFormula(box: TextBox): string {
  const source = `${box.text} ${box.latex ?? ""}`;
  const match =
    /([A-Za-z𝑎-𝑧𝐚-𝐳])\s*!{2}/u.exec(source) ??
    /([A-Za-z𝑎-𝑧𝐚-𝐳])_!{2}/u.exec(source) ??
    /([A-Za-z𝑎-𝑧𝐚-𝐳])_\{\s*!{2}\s*\}/u.exec(source);
  return match?.[1] ?? "v";
}

function latexForFormulaUnit(unit: string): string {
  return latexFromPdfFormulaText(unit).trim() || `\\text{${unit}}`;
}

function repairedCorruptedZeroFractionLatex(
  formula: TextBox,
  numeratorUnit: string,
  denominatorUnit: string,
): string {
  const variable = latexFromPdfFormulaText(
    variableFromCorruptedZeroFormula(formula),
  );
  const numerator = `0${latexForFormulaUnit(numeratorUnit)}`;
  const denominator = `0${latexForFormulaUnit(denominatorUnit)}`;
  const resultUnit = `${latexForFormulaUnit(numeratorUnit)}/${latexForFormulaUnit(
    denominatorUnit,
  )}`;
  return `${variable} = \\frac{${numerator}}{${denominator}} = ${resultUnit}`;
}

function zeroUnitPartsFromFormulaText(text: string): string[] {
  const unitPattern = new RegExp(`0\\s*(${FORMULA_UNIT_PATTERN})`, "gi");
  return Array.from(text.matchAll(unitPattern), (match) => match[1]).filter(
    Boolean,
  );
}

type CorruptedZeroUnitBox = {
  box: TextBox;
  index: number;
  part: { quantity: string; unit: string };
};

function repairCorruptedZeroFractionFormulaTextBoxes(
  boxes: readonly TextBox[],
): TextBox[] {
  const consumed = new Set<number>();
  const replacements = new Map<number, TextBox>();

  boxes.forEach((formula, formulaIndex) => {
    if (consumed.has(formulaIndex)) return;
    if (!formulaBoxLooksLikeCorruptedZeroFraction(formula)) return;

    const fontSize = Math.max(formula.fontSize, 0.001);
    const left = formula.x;
    const right = textBoxRight(formula);
    const top = formula.y;
    const bottom = formula.y + formula.height;
    const nearby = boxes
      .map((box, index) => ({ box, index }))
      .filter(({ box, index }) => {
        if (index === formulaIndex || consumed.has(index)) return false;
        if (!textBoxHorizontallyNear(box, left, right, fontSize)) return false;
        return textBoxVerticallyNear(box, top, bottom, fontSize);
      });

    const zeroUnitBoxes = nearby
      .map((item): CorruptedZeroUnitBox | null => {
        const part = zeroQuantityUnitPart(item.box.text);
        return part ? { ...item, part } : null;
      })
      .filter((item): item is CorruptedZeroUnitBox => item != null)
      .sort((a, b) => textBoxCenterY(a.box) - textBoxCenterY(b.box));
    const zeroBoxes = nearby.filter(({ box }) => box.text.trim() === "0");
    const mergedFormulaUnits = zeroUnitPartsFromFormulaText(formula.text);
    const canRepairFromResidualBoxes =
      zeroUnitBoxes.length >= 2 &&
      zeroResidualBoxesLookStacked(
        zeroBoxes.map(({ box }) => box),
        fontSize,
      );
    const numeratorUnit = canRepairFromResidualBoxes
      ? zeroUnitBoxes[0].part.unit
      : mergedFormulaUnits[0];
    const denominatorUnit = canRepairFromResidualBoxes
      ? zeroUnitBoxes[zeroUnitBoxes.length - 1].part.unit
      : mergedFormulaUnits[mergedFormulaUnits.length - 1];
    if (
      !numeratorUnit ||
      !denominatorUnit ||
      numeratorUnit === denominatorUnit
    ) {
      return;
    }

    const repairLatex = repairedCorruptedZeroFractionLatex(
      formula,
      numeratorUnit,
      denominatorUnit,
    );
    const repairText = fallbackTextFromPdfFormulaLatex(repairLatex);
    const repairMembers = canRepairFromResidualBoxes
      ? [
          { box: formula, index: formulaIndex },
          zeroUnitBoxes[0],
          zeroUnitBoxes[zeroUnitBoxes.length - 1],
          ...zeroBoxes,
        ]
      : [{ box: formula, index: formulaIndex }];
    const repairLeft = Math.min(...repairMembers.map(({ box }) => box.x));
    const repairRight = Math.max(
      ...repairMembers.map(({ box }) => textBoxRight(box)),
    );
    const repairTop = Math.min(...repairMembers.map(({ box }) => box.y));
    const repairBottom = Math.max(
      ...repairMembers.map(({ box }) => box.y + box.height),
    );

    replacements.set(formulaIndex, {
      ...formula,
      x: repairLeft,
      y: repairTop,
      width: Math.max(0, repairRight - repairLeft),
      height: Math.max(0, repairBottom - repairTop),
      text: repairText,
      latex: repairLatex,
      kind: "formula",
    });
    for (const { index } of repairMembers) {
      if (index !== formulaIndex) consumed.add(index);
    }
  });

  return boxes
    .map((box, index) => replacements.get(index) ?? box)
    .filter((_, index) => !consumed.has(index));
}

function shouldMergeParagraphLines(
  previous: LineBoxInternal,
  next: LineBoxInternal,
): boolean {
  if (previous.kind === "formula" || next.kind === "formula") return false;
  if (isDocumentIndexLine(previous.text) || isDocumentIndexLine(next.text)) {
    return false;
  }
  if (
    isWeakScalarMetricLine(previous.text) ||
    isWeakScalarMetricLine(next.text)
  ) {
    return false;
  }
  if (isShortPlotLabelText(previous.text) && isShortPlotLabelText(next.text)) {
    return false;
  }
  if (
    isNonMathematicalMatrixAssignmentRowText(previous.text) ||
    isNonMathematicalMatrixAssignmentRowText(next.text)
  ) {
    return false;
  }

  const fontSize = Math.max(previous.fontSize, next.fontSize);
  const verticalGap = next.top - previous.bottom;
  if (verticalGap < -fontSize * 0.35 || verticalGap > fontSize * 0.95) {
    return false;
  }
  if (fontSizeRatio(previous.fontSize, next.fontSize) > 1.25) return false;
  // Paragraph lines keep the strict test: a heading in another weight of
  // the same face is not its paragraph's first line.
  if (styleKey(previous) !== styleKey(next)) return false;

  const leftDelta = Math.abs(next.left - previous.left);
  const overlap =
    Math.min(previous.right, next.right) - Math.max(previous.left, next.left);
  const minWidth = Math.max(
    1,
    Math.min(previous.right - previous.left, next.right - next.left),
  );
  const overlapRatio = overlap / minWidth;

  return (
    leftDelta <= fontSize * 2 ||
    (overlapRatio >= 0.65 && leftDelta <= fontSize * 6)
  );
}

function paragraphBoxFromLines(lines: LineBoxInternal[]): LineBoxInternal {
  const dominant = lines.reduce(
    (a, b) => (b.fontSize > a.fontSize ? b : a),
    lines[0],
  );
  const kind = lines.some((line) => line.kind === "formula")
    ? "formula"
    : "text";
  const latex = kind === "formula" ? formulaLatexFromLines(lines) : undefined;
  const text =
    kind === "formula"
      ? readableFormulaTextFromLineGroup(lines, latex)
      : lines.map((line) => line.text).join("\n");
  const textRuns =
    kind === "text"
      ? mergePdfTextStyleRuns(
          lines.flatMap((line, index) => {
            const lineRuns = projectedRunsForLine(line);
            if (index === lines.length - 1) return lineRuns;
            const separatorStyle = lineRuns[lineRuns.length - 1] ?? {
              text: "",
              fontSize: line.fontSize,
              fontFamily: line.fontFamily,
              displayFontFamily: line.displayFontFamily,
              sourceFontName: line.sourceFontName,
              bold: line.bold,
              italic: line.italic,
              color: line.color,
            };
            return [...lineRuns, { ...separatorStyle, text: "\n" }];
          }),
        )
      : undefined;
  const left = Math.min(...lines.map((line) => line.left));
  const right = Math.max(...lines.map((line) => line.right));
  const top = Math.min(...lines.map((line) => line.top));
  const bottom = Math.max(...lines.map((line) => line.bottom));
  const boxWidth = Math.max(0.0001, right - left);
  const boxHeight = Math.max(0.0001, bottom - top);
  let textOffset = 0;
  const lineLayout: PdfTextLineLayout[] | undefined =
    kind === "text"
      ? lines.map((line, index) => {
          const start = textOffset;
          const end = start + line.text.length;
          textOffset = end + (index < lines.length - 1 ? 1 : 0);
          return {
            start,
            end,
            x: Math.max(0, (line.left - left) / boxWidth),
            y: Math.max(0, (line.top - top) / boxHeight),
            width: Math.max(0.000001, (line.right - line.left) / boxWidth),
            height: Math.max(0.000001, (line.bottom - line.top) / boxHeight),
          };
        })
      : undefined;
  return {
    left,
    right,
    top,
    bottom,
    // Every line in a paragraph is at the same angle — the pipeline runs once
    // per angle — so the paragraph is too.
    ...(dominant.angle ? { angle: dominant.angle } : {}),
    fontSize: dominant.fontSize,
    text,
    fontFamily: dominant.fontFamily,
    displayFontFamily: dominant.displayFontFamily,
    sourceFontName: dominant.sourceFontName,
    bold: dominant.bold,
    italic: dominant.italic,
    color: dominant.color,
    textRuns:
      textRuns && textFromPdfTextStyleRuns(textRuns) === text
        ? textRuns
        : undefined,
    lineLayout,
    kind,
    latex,
    explicitLatex:
      kind === "formula" && lines.some((line) => line.explicitLatex?.trim())
        ? latex
        : undefined,
    formulaRuns: kind === "formula" ? formulaRunsFromLines(lines) : undefined,
    equationAnchored: lines.some((line) => line.equationAnchored),
  };
}

/**
 * Group projected items into line boxes (same baseline, left-to-right) and
 * normalise to page fractions.
 */
export function groupTextBoxes(
  items: PdfTextItem[],
  viewport: PdfViewportLike,
  idPrefix = "tb",
): TextBox[] {
  const projected = items
    .map((item) => projectItem(item, viewport))
    .filter(
      (p) =>
        (p.str.trim() !== "" || hasFormulaLatex(p.latex)) &&
        Number.isFinite(p.x) &&
        Number.isFinite(p.baselineY),
    );
  // Sort top-to-bottom, then left-to-right.
  // Ordered along each run's own axes. Upright text has along === x and
  // across === baselineY, so this is the previous sort unchanged; a turned run
  // reads in its own direction instead of the page's.
  projected.sort(
    (a, b) => a.angle - b.angle || a.across - b.across || a.along - b.along,
  );

  const lines: ProjectedItem[][] = [];
  for (const p of projected) {
    const line = lines[lines.length - 1];
    const last = line?.[line.length - 1];
    // Same visual run: baselines align, x order is sane, and it has not crossed
    // a WPS-style block boundary (large tab/page-number gap or clear style jump).
    if (
      last &&
      // Text at a different angle is not the same visual run, whatever its
      // baseline reads as: a rotated watermark crossing a paragraph would
      // otherwise be spliced into its middle.
      Math.abs(p.angle - last.angle) < UPRIGHT_ANGLE_EPSILON &&
      Math.abs(p.across - last.across) <= last.fontSize * 0.6 &&
      p.along >= last.along - last.fontSize &&
      !shouldStartNewInlineSegment(last, p) &&
      !shouldDetachPlotLabelFromInlineMetric(line, p)
    ) {
      line.push(p);
    } else {
      lines.push([p]);
    }
  }

  const W = viewport.width || 1;
  const H = viewport.height || 1;
  // Every merge below reasons about rectangles, and once each run's rect is
  // placed by its own frame, a watermark and the text it crosses can share a
  // centre. Rather than teach a dozen mergers about angles, the pipeline runs
  // once per angle. A document with only upright text has exactly one group,
  // so it runs exactly as it always did.
  const byAngle = new Map<number, ProjectedItem[][]>();
  for (const line of lines) {
    const key = line[0]?.angle ?? 0;
    const bucket = byAngle.get(key);
    if (bucket) bucket.push(line);
    else byAngle.set(key, [line]);
  }
  // Paragraph assembly and the residue passes reason about rectangles too, so
  // the WHOLE pipeline runs per angle rather than just the merges.
  const groups = [...byAngle.entries()];
  return groups.flatMap(([angle, angleLines], groupIndex) => {
    // The upright group keeps the caller's prefix, so ids in an ordinary
    // document are exactly what they have always been.
    const prefix = angle === 0 || groups.length === 1
      ? idPrefix
      : `${idPrefix}-a${groupIndex}`;
    const angleItems = projected.filter(
      (item) => Math.abs(item.angle - angle) < UPRIGHT_ANGLE_EPSILON,
    );
    const boxes = lineBoxesToTextBoxes(
      mergeLineBoxes(angleLines, W),
      angleItems,
      W,
      H,
      prefix,
    );
    return angle === 0 ? absorbFormulaFragments(boxes, angleItems, W, H) : boxes;
  });
}

/** What a fragment of a formula may read: digits, operators, brackets and
 * one- or two-letter symbols -- never a word, a CJK character or an
 * equation number. */
function isFormulaFragmentText(text: string): boolean {
  const trimmed = text.trim();
  return (
    trimmed.length > 0 &&
    trimmed.length <= 16 &&
    !isEquationNumberLine(trimmed) &&
    proseWordCount(trimmed) === 0 &&
    cjkCharCount(trimmed) === 0 &&
    /^[\s0-9A-Za-z+\-−=<>≤≥≈≠×÷·*/().,'′^_|]+$/.test(trimmed) &&
    !/[A-Za-z]{3,}/.test(trimmed)
  );
}











function absorbFormulaFragments(
  boxes: TextBox[],
  items: readonly ProjectedItem[],
  W: number,
  H: number,
): TextBox[] {
  type Rect = { left: number; top: number; right: number; bottom: number };
  const rectOf = (box: TextBox): Rect => ({
    left: box.x * W, top: box.y * H, right: (box.x + box.width) * W, bottom: (box.y + box.height) * H,
  });
  const out = [...boxes];
  const gone = new Set<TextBox>();
  for (let index = 0; index < out.length; index += 1) {
    const formula = out[index];
    if (formula.kind !== "formula" || gone.has(formula)) continue;
    let rect = rectOf(formula);
    const em = Math.max(formula.fontSize * H, 1);
    const parts: TextBox[] = [];
    const taken = () => [rectOf(formula), ...parts.map(rectOf)];
    const take = (box: TextBox, b: Rect) => {
      parts.push(box);
      rect = {
        left: Math.min(rect.left, b.left), top: Math.min(rect.top, b.top),
        right: Math.max(rect.right, b.right), bottom: Math.max(rect.bottom, b.bottom),
      };
    };
    const candidates = () => out.filter((box) => box !== formula && !gone.has(box) && !parts.includes(box) &&
      box.kind !== "formula" && isFormulaFragmentText(box.text));
    for (let grew = true; grew;) {
      grew = false;
      // The rest of one of its rows: an operator and what follows it, just
      // after the formula (never a bare label or an equation number).
      for (const box of candidates()) {
        const b = rectOf(box);
        const after = (share: number) => [formula, ...parts].find((part) => {
          const row = rectOf(part);
          return Math.min(row.bottom, b.bottom) - Math.max(row.top, b.top) > (b.bottom - b.top) * share &&
            b.left - row.right >= -em * 0.4 && b.left - row.right <= em * 1.35;
        });
        // An operator and what follows it, on one of its rows; or the operand
        // a relation or an operator left waiting -- "y =" then "x", or "D ="
        // then a numerator raised over the row.
        const onRow = after(0.5);
        const beside = after(0);
        const opens = Boolean(onRow) && /^[=<>≤≥≈≠+\-−×÷·*/]/.test(box.text.trim());
        const operand = Boolean(beside) && /[=<>≤≥≈≠+\-−×÷·*/]$/.test(beside!.text.trim()) &&
          /^[0-9A-Za-z(]/.test(box.text.trim());
        if (opens || operand) { take(box, b); grew = true; }
      }
      // A numerator or denominator within its width, stacked on a part of it.
      for (const box of candidates()) {
        const b = rectOf(box);
        if (b.left < rect.left - em * 0.2 || b.right > rect.right + em * 0.2) continue;
        const stackedOn = taken().some((part) => {
          const overlap = Math.min(part.right, b.right) - Math.max(part.left, b.left);
          const gap = Math.max(b.top - part.bottom, part.top - b.bottom);
          return overlap >= (b.right - b.left) * 0.3 && gap <= em * 0.6 &&
            Math.abs((part.top + part.bottom) / 2 - (b.top + b.bottom) / 2) >= em * 0.5;
        });
        if (stackedOn && /^[0-9A-Za-z+\-−().,\s]+$/.test(box.text.trim())) { take(box, b); grew = true; }
      }
      // An exponent or index inside it: smaller, and with the formula going on
      // after it (a marker after its end is a footnote's, not an exponent).
      for (const box of candidates()) {
        const b = rectOf(box);
        if (box.fontSize * H > em * 0.82) continue;
        const inside = b.left >= rect.left && b.right <= rect.right &&
          taken().some((part) => part.left >= b.right - em * 0.1 &&
            Math.min(part.bottom, b.bottom) - Math.max(part.top, b.top) > 0);
        if (inside) { take(box, b); grew = true; }
      }
    }
    if (parts.length === 0) continue;
    const pad = em * 0.1;
    const runs: PdfFormulaRun[] = items
      .filter((item) => {
        const centerX = item.x + item.width / 2;
        const centerY = item.baselineY - item.fontSize * 0.35;
        return item.str.trim() !== "" && centerX >= rect.left - pad && centerX <= rect.right + pad &&
          centerY >= rect.top - pad && centerY <= rect.bottom + pad;
      })
      .map((item) => ({
        left: item.x,
        right: item.x + item.width,
        centerY: item.baselineY - item.fontSize * 0.35,
        baselineY: item.baselineY,
        fontSize: item.fontSize,
        text: item.str,
      }));
    const latex = latexFromPdfFormulaRuns(runs).trim();
    if (!latex) continue;
    const visualText = [formula, ...parts].map((box) => box.text).join(" ");
    const { runs: _runs, lineLayout: _lineLayout, ...kept } = formula;
    out[index] = {
      ...kept,
      x: rect.left / W,
      y: rect.top / H,
      width: (rect.right - rect.left) / W,
      height: (rect.bottom - rect.top) / H,
      text: readableFormulaTextFromLatex(latex, visualText),
      latex,
    };
    for (const part of parts) gone.add(part);
  }
  return out.filter((box) => !gone.has(box));
}

/** One angle's worth of lines, through the whole merge pipeline. */
function mergeLineBoxes(
  lines: ProjectedItem[][],
  W: number,
): LineBoxInternal[] {
  const rawLineBoxes = mergeSplitEquationNumberLines(
    lines.flatMap((line) =>
      splitInlinePlotMetricSuffixLine(lineBoxFromItems(line))
        .flatMap(splitLooseTrailingFormulaLabelLine)
        .flatMap(splitEmbeddedEquationNumberLine),
    ),
  );

  const lineBoxes = mergeInlineProseTextRuns(
    mergeUnnumberedDisplayFormulaFragments(
      mergeUnnumberedSameRowDisplayFormulaFragments(
        mergeNumberedDisplayFormulaFragments(
          mergeSplitCaseColumnLines(
            mergeCorruptedZeroStackedFractionFragments(
              mergeDanglingNamedOperatorFractionBridges(
                mergeClosedSameRowNamedOperatorFormulaFragments(
                  mergeInlineFormulaAffixes(rawLineBoxes),
                  W,
                ),
              ),
            ),
          ),
          W,
        ),
        W,
      ),
      W,
    ),
  );
  return mergeOrDropOrphanPunctuationFragments(lineBoxes);
}

/** Paragraph assembly and the final projection into page fractions. */
function lineBoxesToTextBoxes(
  cleanedLineBoxes: LineBoxInternal[],
  projected: ProjectedItem[],
  W: number,
  H: number,
  idPrefix: string,
): TextBox[] {
  const paragraphs: LineBoxInternal[][] = [];
  let lastFormulaParagraph: LineBoxInternal[] | null = null;
  for (const line of cleanedLineBoxes) {
    if (isEquationNumberLine(line.text)) {
      paragraphs.push([line]);
      lastFormulaParagraph = null;
      continue;
    }

    if (
      lastFormulaParagraph &&
      shouldMergeFormulaLines(lastFormulaParagraph, line)
    ) {
      lastFormulaParagraph.push(line);
      continue;
    }

    const paragraph = paragraphs[paragraphs.length - 1];
    const previous = paragraph?.[paragraph.length - 1];
    if (paragraph && shouldMergeFormulaLines(paragraph, line)) {
      paragraph.push(line);
      lastFormulaParagraph = paragraph;
    } else if (previous && shouldMergeParagraphLines(previous, line)) {
      paragraph.push(line);
    } else {
      const nextParagraph = [line];
      paragraphs.push(nextParagraph);
      if (line.kind === "formula" && !isEquationNumberLine(line.text)) {
        lastFormulaParagraph = nextParagraph;
      }
    }
  }

  const boxes: TextBox[] = [];
  paragraphs.forEach((paragraph, index) => {
    const box = paragraphBoxFromLines(paragraph);
    const runs = box.textRuns
      ?.map((run) => {
        const sampleX = Math.max(0, Math.min(1, Number(run.sampleX) / W));
        const sampleY = Math.max(0, Math.min(1, Number(run.sampleY) / H));
        const sampleWidth = Math.max(
          0,
          Math.min(1 - sampleX, Number(run.sampleWidth) / W),
        );
        const sampleHeight = Math.max(
          0,
          Math.min(1 - sampleY, Number(run.sampleHeight) / H),
        );
        return {
          ...run,
          fontSize:
            typeof run.fontSize === "number" ? run.fontSize / H : undefined,
          ...(Number.isFinite(sampleX) &&
          Number.isFinite(sampleY) &&
          Number.isFinite(sampleWidth) &&
          Number.isFinite(sampleHeight) &&
          sampleWidth > 0 &&
          sampleHeight > 0
            ? { sampleX, sampleY, sampleWidth, sampleHeight }
            : {}),
        };
      })
      .filter((run) => run.text.length > 0);
    boxes.push({
      id: `${idPrefix}-${index}`,
      x: box.left / W,
      y: box.top / H,
      width: Math.max(0, box.right - box.left) / W,
      height: Math.max(0, box.bottom - box.top) / H,
      text: box.text,
      fontSize: box.fontSize / H,
      fontFamily: box.fontFamily,
      displayFontFamily: box.displayFontFamily,
      sourceFontName: box.sourceFontName,
      bold: box.bold,
      italic: box.italic,
      ...(runs && runs.length > 0 && textFromPdfTextStyleRuns(runs) === box.text
        ? { runs }
        : {}),
      ...(box.lineLayout ? { lineLayout: box.lineLayout } : {}),
      ...(box.angle ? { angle: box.angle } : {}),
      color: box.color,
      kind: box.kind,
      latex: box.latex,
    });
  });
  return absorbFormulaBoxesNestedInLargerFormulas(
    markUntrustedFormulaBoxes(
      pruneFormulaCoveredResidualTextBoxes(
        repairCorruptedZeroFractionFormulaTextBoxes(
          mergeResidualNumberedFormulaTextBoxes(boxes),
        ),
      ),
      projected,
      W,
      H,
    ),
  );
}

/**
 * Drop tiny formula fragments whose box sits inside a much larger formula
 * box. `pruneFormulaCoveredResidualTextBoxes` only removes covered *text*
 * residue, so a stray single-glyph run classified as a formula (common with
 * broken-ToUnicode math fonts) survives as a duplicate box on top of the
 * display formula that already covers it. Only untranscribable fragments are
 * dropped — a nested fragment that still carries its own LaTeX is kept, since
 * removing it would lose user-actionable content.
 */
export function absorbFormulaBoxesNestedInLargerFormulas(
  boxes: TextBox[],
): TextBox[] {
  const formulas = boxes.filter((box) => box.kind === "formula");
  if (formulas.length < 2) return boxes;
  return boxes.filter((box) => {
    if (box.kind !== "formula") return true;
    if (box.latex?.trim()) return true;
    const area = box.width * box.height;
    const centerX = box.x + box.width / 2;
    const centerY = box.y + box.height / 2;
    const host = formulas.find((other) => {
      if (other === box) return false;
      const otherArea = other.width * other.height;
      if (otherArea < area * 6) return false;
      return (
        centerX >= other.x &&
        centerX <= other.x + other.width &&
        centerY >= other.y &&
        centerY <= other.y + other.height
      );
    });
    return host == null;
  });
}

/**
 * Final honesty pass: formula boxes built (even partially) from runs whose
 * characters are raw-charcode fallbacks of a broken math-font ToUnicode CMap
 * must not carry a LaTeX seed — the derived LaTeX would be wrong-but-plausible
 * (`\lim_{2\to 3}` for a rendered `\lim_{x\to c}`). Geometry is matched
 * against the ORIGINAL projected runs so intermediate merges/splits cannot
 * launder an untrusted run into a trusted box.
 */
function markUntrustedFormulaBoxes(
  boxes: TextBox[],
  projected: readonly ProjectedItem[],
  pageWidth: number,
  pageHeight: number,
): TextBox[] {
  const untrusted = projected.filter((p) => p.untrustedText && p.str.trim());
  if (untrusted.length === 0) return boxes;
  return boxes.map((box) => {
    if (box.kind !== "formula") return box;
    const left = box.x * pageWidth;
    const top = box.y * pageHeight;
    const right = left + box.width * pageWidth;
    const bottom = top + box.height * pageHeight;
    const tainted = untrusted.some((p) => {
      const centerX = p.x + p.width / 2;
      const centerY = p.baselineY - p.fontSize * 0.375;
      return (
        centerX >= left &&
        centerX <= right &&
        centerY >= top &&
        centerY <= bottom
      );
    });
    if (!tainted) return box;
    return { ...box, latex: undefined, transcriptionNeeded: true };
  });
}
