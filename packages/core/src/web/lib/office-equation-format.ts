/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Character formatting on an equation, as Office draws it.
 *
 * Office Math keeps formatting on each math run (`m:r`): its math style
 * (`m:rPr/m:sty`, plain / bold / italic / bold italic) and the host's own run
 * properties (`a:rPr` on a slide, `w:rPr` in Word: underline, strikethrough,
 * colour, size, highlight). Selecting an equation, or part of one, and
 * pressing Bold or Underline formats those runs, so an underlined fraction is
 * underlined under its numerator and under its denominator, each on its own
 * baseline, not under the equation's box.
 *
 * The renderers draw that here, from the formatting the equation carries:
 * bold is KaTeX's bold math (`\boldsymbol`, the math alphabet Office uses
 * for `m:sty="b"`/`"bi"`), not a CSS weight the math fonts do not have; an
 * underline or a strikethrough is drawn on every glyph run of the rendered
 * formula (`data-equation-decoration`, `office-equation-format.css`), since a
 * CSS decoration on the equation's box never reaches the inline-blocks KaTeX
 * lays the glyphs out in.
 */

import { equationLatexWithLetterFormat } from "./office-equation-letter-format";

export type EquationCharacterFormat = {
  bold?: boolean;
  underline?: boolean;
  strike?: boolean;
};

/** The attribute the decoration rules key on. */
export const EQUATION_DECORATION_ATTRIBUTE = "data-equation-decoration";

/** Whether a CSS `font-weight` value is bold. */
export function cssWeightIsBold(weight: unknown): boolean {
  if (weight === "bold" || weight === "bolder") return true;
  const value = typeof weight === "number" ? weight : typeof weight === "string" ? Number(weight) : NaN;
  return Number.isFinite(value) && value >= 600;
}

/** The format a CSS style (a run's, a mark's) states for an equation in it. */
export function equationFormatFromCss(style: {
  fontWeight?: unknown;
  textDecoration?: unknown;
  textDecorationLine?: unknown;
} | null | undefined): EquationCharacterFormat {
  if (!style) return {};
  const decoration = `${typeof style.textDecoration === "string" ? style.textDecoration : ""} ${
    typeof style.textDecorationLine === "string" ? style.textDecorationLine : ""
  }`;
  return {
    ...(cssWeightIsBold(style.fontWeight) ? { bold: true } : {}),
    ...(/\bunderline\b/.test(decoration) ? { underline: true } : {}),
    ...(/\bline-through\b/.test(decoration) ? { strike: true } : {}),
  };
}

/**
 * The LaTeX an equation is drawn from under `format`: bold as bold math on
 * its letters and digits, operators and delimiters at normal weight, as
 * Office Math bolds a zone (`lib/office-equation-letter-format`).
 */
export function equationLatexForFormat(latex: string, format: EquationCharacterFormat | null | undefined): string {
  if (!format?.bold || !latex.trim()) return latex;
  return equationLatexWithLetterFormat(latex, "bold", true);
}

/** The decoration attribute value for `format`, or null when it draws none. */
export function equationDecorationValue(format: EquationCharacterFormat | null | undefined): string | null {
  const lines = [format?.underline ? "underline" : null, format?.strike ? "line-through" : null].filter(Boolean);
  return lines.length ? lines.join(" ") : null;
}
