/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  fallbackTextFromPdfFormulaLatex,
  normalizePdfFormulaLatexSource,
} from "./formula-latex";
import { renderKatexToString } from "../katex-render";
import type { BoxRect } from "./text-box-geometry";

export const PDF_FORMULA_KATEX_RENDER_OPTIONS = {
  displayMode: true,
  throwOnError: true,
  strict: "ignore",
} as const;

export const PDF_FORMULA_MIN_READABLE_COMPACT_SCALE = 0.9;

export interface PdfFormulaCompactFit {
  scale: number;
  rawScale: number;
  overflows: boolean;
}

export function normalizePdfFormulaForRender(latex: string): string {
  return normalizePdfFormulaLatexSource(latex);
}

export function pdfFormulaLatexValidationError(latex: string): string | null {
  const normalized = normalizePdfFormulaForRender(latex);
  if (normalized.trim() === "") return "Formula source is empty";
  if (hasDegradedPdfFormulaSource(normalized)) {
    return "Formula source is incomplete";
  }
  try {
    renderKatexToString(normalized, PDF_FORMULA_KATEX_RENDER_OPTIONS);
    return null;
  } catch (error) {
    return error instanceof Error ? error.message : "Invalid formula source";
  }
}

export function pdfFormulaCompactFitScale({
  availableWidth,
  availableHeight,
  contentWidth,
  contentHeight,
}: {
  availableWidth: number;
  availableHeight: number;
  contentWidth: number;
  contentHeight: number;
}): PdfFormulaCompactFit {
  if (
    !Number.isFinite(availableWidth) ||
    !Number.isFinite(availableHeight) ||
    !Number.isFinite(contentWidth) ||
    !Number.isFinite(contentHeight) ||
    availableWidth <= 0 ||
    availableHeight <= 0 ||
    contentWidth <= 0 ||
    contentHeight <= 0
  ) {
    return { scale: 1, rawScale: 1, overflows: false };
  }

  const rawScale = Math.min(
    1,
    availableWidth / contentWidth,
    availableHeight / contentHeight,
  );
  const scale =
    rawScale < PDF_FORMULA_MIN_READABLE_COMPACT_SCALE
      ? PDF_FORMULA_MIN_READABLE_COMPACT_SCALE
      : rawScale;
  return {
    scale,
    rawScale,
    overflows: scale > rawScale + 0.001,
  };
}

const STRUCTURED_FORMULA_ENV =
  "(?:aligned|alignedat|gathered|split|cases|matrix|bmatrix|Bmatrix|pmatrix|vmatrix|Vmatrix|smallmatrix|array)";

const STRUCTURED_FORMULA_ENV_RE = new RegExp(
  `\\\\begin\\{${STRUCTURED_FORMULA_ENV}\\}`,
);

const AUTO_RENDER_SAFE_STRUCTURED_ENV_RE =
  /\\begin\{(?:gathered|cases|matrix|bmatrix|Bmatrix|pmatrix|vmatrix|Vmatrix|smallmatrix|array)\}/;

const ALIGNMENT_REQUIRED_STRUCTURED_ENV_RE =
  /\\begin\{(?:aligned|alignedat|split)\}/;

const STRUCTURAL_FORMULA_COMMAND_RE =
  /\\(?:frac|dfrac|tfrac|sqrt|sum|prod|int|lim|arg|max|min)\b/;

const LIMIT_OPERATOR_WITH_SCRIPT_RE =
  /(?:\\(?:lim|sum|prod|int|max|min)\s*(?:_|\^)|\\arg\s*(?:max|min)\s*(?:_|\^))/;

const NAMED_FORMULA_OPERATOR_RE =
  /\\(?:operatorname|Pr|Cov|Var|det|log|ln|exp|sin|cos|tan)\b/;

const GREEK_OR_SET_FORMULA_COMMAND_RE =
  /\\(?:alpha|beta|gamma|delta|epsilon|varepsilon|zeta|eta|theta|vartheta|kappa|lambda|mu|nu|xi|pi|rho|sigma|tau|phi|varphi|chi|psi|omega|Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|Phi|Psi|Omega|mathbb|infty|partial|nabla|in|notin|forall|exists|land|lor|cap|cup|subset|subseteq|supset|supseteq|to|rightarrow|leftarrow|leftrightarrow)\b/;

const LATEX_RELATION_RE =
  /(?:=|<|>|≤|≥|≈|≠|\\(?:le|leq|ge|geq|approx|ne|neq|equiv|propto|in|notin|subset|subseteq|supset|supseteq)\b)/;

const INFIX_FORMULA_OPERATOR_RE =
  /(?:[+\-−*/×÷=<>≤≥≈≠]|\\(?:times|div|cdot|pm|mp)\b)/;

const RELATION_TOKEN_RE =
  /(?:=|<|>|≤|≥|≈|≠|\\(?:le|leq|ge|geq|approx|ne|neq|equiv|propto|in|notin|subset|subseteq|supset|supseteq)\b)/g;

const TRAILING_INFIX_FORMULA_OPERATOR_RE =
  /(?:[=<>≤≥≈≠+\-−*/×÷]|\\(?:times|div|cdot|pm|mp|le|leq|ge|geq|approx|ne|neq|equiv|propto|in|notin|subset|subseteq|supset|supseteq)\b)\s*$/;

const LEADING_INFIX_FORMULA_OPERATOR_RE =
  /^(?:[=<>≤≥≈≠+*/×÷]|\\(?:times|div|cdot|le|leq|ge|geq|approx|ne|neq|equiv|propto|in|notin|subset|subseteq|supset|supseteq)\b)/;

const LEADING_SIGN_OPERATOR_RE = /^(?:[+\-−]|\\(?:pm|mp)\b)/;

function hasFormulaFunctionCall(text: string): boolean {
  return /(?:^|[^A-Za-z])(?:[A-Za-z]|\\(?:Pr|operatorname\{[A-Za-z]+\}))\s*[\[(]/.test(
    text,
  );
}

function isHighConfidenceRelationFormulaLatex(text: string): boolean {
  if (!LATEX_RELATION_RE.test(text)) return false;
  if (
    NAMED_FORMULA_OPERATOR_RE.test(text) ||
    GREEK_OR_SET_FORMULA_COMMAND_RE.test(text) ||
    hasFormulaFunctionCall(text)
  ) {
    return true;
  }
  return false;
}

function formulaRowsForSourceQuality(normalized: string): string[] {
  const withoutEnvironments = normalized.replace(
    /\\(?:begin|end)\{[^{}]+\}/g,
    "",
  );
  return withoutEnvironments
    .split(/\\\\/)
    .map((row) => row.trim())
    .filter(Boolean);
}

function relationCommandStartsAt(row: string, index: number): boolean {
  if (row[index] !== "\\") return false;
  return /^\\(?:le|leq|ge|geq|approx|ne|neq|equiv|propto|in|notin|subset|subseteq|supset|supseteq)\b/.test(
    row.slice(index),
  );
}

function hasFormulaOperandSignal(text: string): boolean {
  const cleaned = text
    .replace(/&/g, " ")
    .replace(/\\(?:,|;|:|!|quad|qquad)\b/g, " ")
    .replace(/[{}]/g, " ")
    .trim();
  if (!cleaned) return false;
  return /(?:[A-Za-z0-9]|[\u3400-\u9fff]|[αβγδεζηθκλμνξπρστφχψωΓΔΘΛΞΠΣΦΨΩ]|[)\]]|\\(?:frac|dfrac|tfrac|sqrt|sum|prod|int|lim|arg|max|min|operatorname|Pr|Cov|Var|det|log|ln|exp|sin|cos|tan|alpha|beta|gamma|delta|epsilon|varepsilon|zeta|eta|theta|vartheta|kappa|lambda|mu|nu|xi|pi|rho|sigma|tau|phi|varphi|chi|psi|omega|Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|Phi|Psi|Omega|mathbb|infty|partial|nabla)\b)/.test(
    cleaned,
  );
}

function rightSideHasOperandAfterLeadingOperator(text: string): boolean {
  const right = text.trim();
  if (!right) return false;
  if (LEADING_INFIX_FORMULA_OPERATOR_RE.test(right)) return false;
  const sign = LEADING_SIGN_OPERATOR_RE.exec(right);
  if (sign) {
    return hasFormulaOperandSignal(right.slice(sign[0].length));
  }
  return hasFormulaOperandSignal(right);
}

function rowHasDanglingRelationOrOperator(
  row: string,
  allowLeadingRelationContinuation: boolean,
): boolean {
  const trimmed = row.trim();
  if (!trimmed) return true;
  if (TRAILING_INFIX_FORMULA_OPERATOR_RE.test(trimmed)) return true;
  if (!allowLeadingRelationContinuation && LEADING_INFIX_FORMULA_OPERATOR_RE.test(trimmed)) {
    return true;
  }

  RELATION_TOKEN_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = RELATION_TOKEN_RE.exec(row))) {
    const index = match.index;
    if (relationCommandStartsAt(row, index - 1)) continue;
    const left = row.slice(0, index);
    const right = row.slice(index + match[0].length);
    const hasLeftOperand =
      hasFormulaOperandSignal(left) ||
      (allowLeadingRelationContinuation && left.replace(/&/g, "").trim() === "");
    if (!hasLeftOperand) return true;
    if (!rightSideHasOperandAfterLeadingOperator(right)) return true;
  }

  return false;
}

function hasMissingOperandAroundInfixOperator(row: string): boolean {
  return (
    /(?:^|[=+\-−*/×÷<>≤≥≈≠])\s*\\(?:times|div|cdot)\b/.test(row) ||
    /\\(?:times|div|cdot)\b\s*(?:[=+\-−*/×÷<>≤≥≈≠]|$)/.test(row) ||
    /[+\-−*/×÷]\s*[=<>≤≥≈≠]/.test(row)
  );
}

function rowLooksLikeRepeatedBareAtoms(row: string): boolean {
  if (INFIX_FORMULA_OPERATOR_RE.test(row)) return false;
  if (/[()[\]{}]/.test(row)) return false;
  const tokens = row.split(/\s+/).filter(Boolean);
  if (tokens.length < 3) return false;
  return tokens.every((token) =>
    /^(?:[A-Za-z0-9]|\\(?:alpha|beta|gamma|delta|theta|lambda|mu|pi|sigma|omega|Gamma|Delta|Theta|Lambda|Pi|Sigma|Omega))$/.test(
      token,
    ),
  );
}

function rowHasWeakBareLimitSource(row: string): boolean {
  const compact = row.trim().replace(/\s+/g, " ");
  if (!/\\lim\b/.test(compact)) return false;
  if (/\\lim\s*(?:_|\^|\{)/.test(compact)) return false;
  if (/\\lim\s+(?:\\?[A-Za-z]|\\operatorname\{[A-Za-z]+\})\s*[\[(]/.test(compact)) {
    return false;
  }
  return /(?:^|[=<>≤≥≈≠]\s*)\\lim\s+[A-Za-z]{3,}\s*$/.test(compact);
}

function hasEmptyRequiredFormulaGroup(row: string): boolean {
  return (
    /\\(?:sqrt|operatorname|mathrm|text|mathcal|mathbf|mathit|mathsf|mathtt|mathbb)\s*\{\s*\}/.test(
      row,
    ) ||
    /\\(?:frac|dfrac|tfrac)\s*\{\s*\}/.test(row) ||
    /\\(?:frac|dfrac|tfrac)\s*\{[^{}]*\}\s*\{\s*\}/.test(row)
  );
}

function hasDegradedPdfFormulaSource(normalized: string): boolean {
  if (/\\[()]/.test(normalized)) return true;
  const allowsLeadingRelationContinuation =
    STRUCTURED_FORMULA_ENV_RE.test(normalized) || /&/.test(normalized);
  const rows = formulaRowsForSourceQuality(normalized);
  if (rows.length === 0) return true;
  return rows.some(
    (row) =>
      rowHasDanglingRelationOrOperator(
        row,
        allowsLeadingRelationContinuation,
      ) ||
      hasEmptyRequiredFormulaGroup(row) ||
      hasMissingOperandAroundInfixOperator(row) ||
      rowHasWeakBareLimitSource(row) ||
      rowLooksLikeRepeatedBareAtoms(row),
  );
}

export function isStructuredPdfFormulaLatex(latex: string): boolean {
  const normalized = normalizePdfFormulaForRender(latex).replace(/\s+/g, " ");
  if (normalized.trim() === "") return false;
  if (hasDegradedPdfFormulaSource(normalized)) return false;
  const hasSafeStructuredEnvironment =
    AUTO_RENDER_SAFE_STRUCTURED_ENV_RE.test(normalized) ||
    (ALIGNMENT_REQUIRED_STRUCTURED_ENV_RE.test(normalized) && /&/.test(normalized));
  if (
    ALIGNMENT_REQUIRED_STRUCTURED_ENV_RE.test(normalized) &&
    !hasSafeStructuredEnvironment
  ) {
    return false;
  }
  return (
    STRUCTURAL_FORMULA_COMMAND_RE.test(normalized) ||
    hasSafeStructuredEnvironment ||
    /(?:\^|_)(?:\{[^{}]+\}|[A-Za-z0-9+\-−])/.test(normalized) ||
    isHighConfidenceRelationFormulaLatex(normalized)
  );
}

export function isRenderablePdfFormulaLatex(latex: string): boolean {
  return (
    isStructuredPdfFormulaLatex(latex) &&
    pdfFormulaLatexValidationError(latex) == null
  );
}








export function isTypesettablePdfFormulaLatex(latex: string): boolean {
  return latex.trim().length > 0 && pdfFormulaLatexValidationError(latex) == null;
}

/** Whether a formula box is drawn as its LaTeX: one the user made whenever
 * the LaTeX parses, one found on the page only when it reads as maths. */
export function pdfFormulaBoxDrawsLatex(box: Readonly<{ created?: boolean }>, latex: string): boolean {
  return box.created === true ? isTypesettablePdfFormulaLatex(latex) : isRenderablePdfFormulaLatex(latex);
}

export function pdfFormulaPreviewHeightPx(
  latex: string,
  fontSizePx: number,
  baseHeightPx: number,
  pageHeightPx: number,
): number {
  const normalized = normalizePdfFormulaForRender(latex).replace(/\s+/g, " ");
  let multiplier = 1.65;
  const hasFraction = /\\(?:frac|dfrac|tfrac)\b/.test(normalized);
  const fractionCount = normalized.match(/\\(?:frac|dfrac|tfrac)\b/g)?.length ?? 0;
  const hasRadical = /\\sqrt\b/.test(normalized);
  const hasLimits = LIMIT_OPERATOR_WITH_SCRIPT_RE.test(normalized);
  if (hasFraction || hasRadical) {
    multiplier = 2.85;
  } else if (/\\(?:sum|prod|int)\b|[_^]/.test(normalized)) {
    multiplier = 2.35;
  }
  if (hasLimits) {
    multiplier = Math.max(multiplier, 2.65);
  }
  if (hasLimits && hasFraction) {
    multiplier = Math.max(multiplier, 4.1);
  } else if (fractionCount > 1) {
    multiplier = Math.max(multiplier, 3.25);
  }
  if (STRUCTURED_FORMULA_ENV_RE.test(normalized)) {
    const rows = (normalized.match(/\\\\/g) ?? []).length + 1;
    multiplier = Math.max(multiplier, Math.min(5.6, 1.75 + rows * 0.95));
  }
  return Math.min(
    Math.max(baseHeightPx, fontSizePx * multiplier),
    Math.max(baseHeightPx, pageHeightPx),
  );
}

function rowTextWidthUnits(text: string): number {
  let units = 0;
  for (const char of Array.from(text)) {
    if (/\s/.test(char)) {
      units += 0.28;
    } else if (/[\u4e00-\u9fff]/.test(char)) {
      units += 1;
    } else if (/[=+\-−×÷*/|()[\]{}.,;:]/.test(char)) {
      units += 0.46;
    } else if (/[A-Z0-9]/.test(char)) {
      units += 0.66;
    } else {
      units += 0.58;
    }
  }
  return units;
}

function formulaRowsForWidthEstimate(normalized: string): string[] {
  return formulaRowsForSourceQuality(normalized);
}

export function pdfFormulaPreviewWidthPx(
  latex: string,
  fontSizePx: number,
  baseWidthPx: number,
  pageWidthPx: number,
): number {
  const normalized = normalizePdfFormulaForRender(latex).replace(/\s+/g, " ");
  if (normalized.trim() === "") return baseWidthPx;

  const rows = formulaRowsForWidthEstimate(normalized);
  const visibleRows = rows.length > 0 ? rows : [normalized];
  const maxRowUnits = Math.max(
    2.6,
    ...visibleRows.map((row) =>
      rowTextWidthUnits(fallbackTextFromPdfFormulaLatex(row)),
    ),
  );
  const structuralPaddingEm =
    (STRUCTURAL_FORMULA_COMMAND_RE.test(normalized)
      ? 1.2
      : 0) +
    (LIMIT_OPERATOR_WITH_SCRIPT_RE.test(normalized) ? 0.8 : 0) +
    (STRUCTURED_FORMULA_ENV_RE.test(normalized) ? 0.9 : 0) +
    (/(?:\^|_)(?:\{[^{}]+\}|[A-Za-z0-9+\-−])/.test(normalized) ? 0.4 : 0);
  const estimatedWidthPx = fontSizePx * (maxRowUnits + structuralPaddingEm);
  return Math.min(
    Math.max(baseWidthPx, estimatedWidthPx),
    Math.max(baseWidthPx, pageWidthPx),
  );
}

export interface PdfFormulaPreviewRectPx {
  left: number;
  top: number;
  width: number;
  height: number;
}

export function pdfFormulaPreviewRectPx({
  latex,
  fontSizePx,
  baseLeftPx,
  baseTopPx,
  baseWidthPx,
  baseHeightPx,
  pageWidthPx,
  pageHeightPx,
}: {
  latex: string;
  fontSizePx: number;
  baseLeftPx: number;
  baseTopPx: number;
  baseWidthPx: number;
  baseHeightPx: number;
  pageWidthPx: number;
  pageHeightPx: number;
}): PdfFormulaPreviewRectPx {
  const safePageWidth = Math.max(1, pageWidthPx);
  const safePageHeight = Math.max(1, pageHeightPx);
  const safeBaseWidth = Math.max(1, baseWidthPx);
  const safeBaseHeight = Math.max(1, baseHeightPx);
  const displayWidth = pdfFormulaPreviewWidthPx(
    latex,
    fontSizePx,
    safeBaseWidth,
    safePageWidth,
  );
  const displayHeight = pdfFormulaPreviewHeightPx(
    latex,
    fontSizePx,
    safeBaseHeight,
    safePageHeight,
  );
  const centeredLeft = baseLeftPx + safeBaseWidth / 2 - displayWidth / 2;
  const centeredTop = baseTopPx + safeBaseHeight / 2 - displayHeight / 2;

  return {
    left: Math.max(
      0,
      Math.min(centeredLeft, Math.max(0, safePageWidth - displayWidth)),
    ),
    top: Math.max(
      0,
      Math.min(centeredTop, Math.max(0, safePageHeight - displayHeight)),
    ),
    width: displayWidth,
    height: displayHeight,
  };
}

/**
 * The frame a box takes once it is set as this formula: grown to the
 * equation the preview draws at the box's size (never shrunk), kept on the
 * page. The overlay's formula boxes and 转换为公式 on a page line alike.
 */
export function pdfFormulaBoxRect({
  rect,
  latex,
  fontSizeFraction,
  pageWidth,
  pageHeight,
}: {
  rect: BoxRect;
  latex: string;
  fontSizeFraction: number;
  pageWidth: number;
  pageHeight: number;
}): BoxRect {
  if (!isTypesettablePdfFormulaLatex(latex)) return rect;

  const baseHeightPx = Math.max(8, rect.height * pageHeight);
  const baseWidthPx = Math.max(8, rect.width * pageWidth);
  const fontSizePx = Math.max(6, fontSizeFraction * pageHeight);
  const formulaRectPx = pdfFormulaPreviewRectPx({
    latex,
    fontSizePx,
    baseLeftPx: rect.x * pageWidth,
    baseTopPx: rect.y * pageHeight,
    baseWidthPx,
    baseHeightPx,
    pageWidthPx: pageWidth,
    pageHeightPx: pageHeight,
  });
  const nextHeight = Math.min(1, formulaRectPx.height / pageHeight);
  const nextWidth = Math.min(1, formulaRectPx.width / pageWidth);
  if (nextHeight <= rect.height + 0.0001 && nextWidth <= rect.width + 0.0001) {
    return rect;
  }

  return {
    ...rect,
    x: Math.max(0, Math.min(formulaRectPx.left / pageWidth, 1 - nextWidth)),
    y: Math.max(0, Math.min(formulaRectPx.top / pageHeight, 1 - nextHeight)),
    width: nextWidth,
    height: nextHeight,
  };
}
