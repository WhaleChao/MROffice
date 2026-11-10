/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word-compatible UnicodeMath (the "Unicode" linear format of Word's
 * Equation tab) for MRAI equations, whose stored source is LaTeX.
 *
 * LaTeX → OMML (`document-export/latex-omml`) and OMML → LaTeX
 * (`./omml-to-latex`) are the single-equation mappings Word's own files use
 * (the server's DOCX import and export read and write the same), so an
 * equation typed in UnicodeMath is stored exactly as the same equation
 * opened from a Word document would be. Both directions are verified by
 * building the result back up; anything that would change the equation
 * returns null and the caller keeps the formula unchanged.
 */

import { latexToOmml } from "../document-export/latex-omml";
import { ommlToLatex } from "./omml-to-latex";
import { astFromOmml, astToOmml, type MathList, type MathNode } from "./ast";
import { canonicalMath } from "./normalize";
import { parseUnicodeMath } from "./parse";
import { serializeUnicodeMath, UnsupportedUnicodeMath } from "./serialize";

export { parseUnicodeMath } from "./parse";
export { serializeUnicodeMath } from "./serialize";
export { astFromOmml, astToOmml } from "./ast";
export { canonicalMath } from "./normalize";

export interface ReversibleEquationConversion {
  readonly latex: string;
  readonly linear: string;
}

/** Linear text for a model, or null when it does not build back up the same. */
export function linearizeMath(list: MathList): string | null {
  const target = canonicalMath(list);
  const attempts = [
    { mode: "word", absorbNaryands: false },
    { mode: "word", absorbNaryands: true },
    { mode: "safe", absorbNaryands: true },
  ] as const;
  for (const attempt of attempts) {
    let text: string;
    try {
      text = serializeUnicodeMath(list, attempt.mode, { absorbNaryands: attempt.absorbNaryands });
    } catch (error) {
      if (error instanceof UnsupportedUnicodeMath) return null;
      throw error;
    }
    if (canonicalMath(parseUnicodeMath(text)) === target) return text;
  }
  return null;
}

/** Word's linear text for an `<m:oMath>` element string. */
export function ommlToUnicodeMath(omml: string): string | null {
  const list = astFromOmml(omml);
  return list ? linearizeMath(list) : null;
}

function exportedModel(latex: string): MathList | null {
  let omml: string;
  try {
    omml = latexToOmml(latex);
  } catch {
    return null;
  }
  return astFromOmml(omml);
}

/**
 * The importer reads a matrix natively only when its cells hold what the
 * exporter writes for them, and the exporter writes `(…)` in a cell as run
 * text. Plain parentheses there are the same equation either way.
 */
function literalParenthesesInCells(list: MathList, inCell = false): MathList {
  const out: MathList = [];
  for (const node of list) {
    if (
      inCell &&
      node.kind === "d" &&
      node.beg === "(" &&
      node.end === ")" &&
      node.args.length === 1
    ) {
      out.push({ kind: "text", text: "(", style: {} });
      out.push(...literalParenthesesInCells(node.args[0], true));
      out.push({ kind: "text", text: ")", style: {} });
      continue;
    }
    if (node.kind === "text" && node.style.lit) {
      // m:lit only stops Word from building the character up again.
      const style = { ...node.style };
      delete (style as { lit?: boolean }).lit;
      out.push({ ...node, style });
      continue;
    }
    out.push(mapChildren(node, (child) => literalParenthesesInCells(child, inCell || node.kind === "m")));
  }
  return out;
}

function mapChildren(node: MathNode, map: (list: MathList) => MathList): MathNode {
  switch (node.kind) {
    case "text":
      return node;
    case "f":
      return { ...node, num: map(node.num), den: map(node.den) };
    case "sSup":
      return { ...node, e: map(node.e), sup: map(node.sup) };
    case "sSub":
      return { ...node, e: map(node.e), sub: map(node.sub) };
    case "sSubSup":
    case "sPre":
      return { ...node, e: map(node.e), sub: map(node.sub), sup: map(node.sup) };
    case "rad":
      return { ...node, deg: map(node.deg), e: map(node.e) };
    case "nary":
      return { ...node, sub: node.sub && map(node.sub), sup: node.sup && map(node.sup), e: map(node.e) };
    case "d":
      return { ...node, args: node.args.map(map) };
    case "m":
      return { ...node, rows: node.rows.map((row) => row.map(map)) };
    case "eqArr":
      return { ...node, rows: node.rows.map(map) };
    case "func":
      return { ...node, name: map(node.name), e: map(node.e) };
    case "limLow":
    case "limUpp":
      return { ...node, e: map(node.e), lim: map(node.lim) };
    default:
      return { ...node, e: map(node.e) };
  }
}

function importedLatex(list: MathList): string | null {
  if (typeof DOMParser === "undefined") return null;
  const document = new DOMParser().parseFromString(
    astToOmml(literalParenthesesInCells(list)),
    "application/xml",
  );
  const root = document.documentElement;
  if (!root || root.getElementsByTagName("parsererror").length > 0) return null;
  return ommlToLatex(root);
}

/**
 * LaTeX whose rendering carries something the DOCX exporter (and Word's
 * equation model) does not keep — colour, a forced math style, a fixed
 * delimiter size, links and tags: the equation would change when its
 * UnicodeMath is built back up. (\tfrac/\dfrac and \widehat are not here:
 * Word has one fraction and one stretching accent object for them.)
 */
const LOSSY_LATEX_COMMANDS = new Set([
  "color", "textcolor", "colorbox", "fcolorbox", "displaystyle", "textstyle",
  "scriptstyle", "scriptscriptstyle", "big", "Big", "bigg", "Bigg", "bigl",
  "Bigl", "biggl", "Biggl", "bigr", "Bigr", "biggr", "Biggr", "bigm", "Bigm",
  "biggm", "Biggm", "href", "url", "htmlClass", "htmlId", "htmlStyle",
  "htmlData", "tag", "label",
]);

function usesLossyLatex(latex: string): boolean {
  for (const match of latex.matchAll(/\\([A-Za-z]+)/g)) {
    if (LOSSY_LATEX_COMMANDS.has(match[1])) return true;
  }
  return false;
}

/** A backslash left in run text is LaTeX the exporter could not convert. */
function hasLiteralLatex(list: MathList): boolean {
  return JSON.stringify(list).includes("\\\\");
}

/** UnicodeMath for a stored formula, only when it round-trips losslessly. */
export function latexToReversibleWordLinear(latex: string): ReversibleEquationConversion | null {
  if (latex.trim() === "") return { latex, linear: "" };
  if (usesLossyLatex(latex)) return null;
  const list = exportedModel(latex);
  if (!list || hasLiteralLatex(list)) return null;
  const linear = linearizeMath(list);
  return linear === null ? null : { latex, linear };
}

/**
 * Build up edited UnicodeMath into the LaTeX the editor stores. The LaTeX is
 * the importer's spelling of the built-up equation and must export back to
 * the same equation, otherwise null (the formula is left unchanged).
 */
export function wordLinearToLatex(linear: string): ReversibleEquationConversion | null {
  if (linear.trim() === "") return { latex: "", linear };
  const list = parseUnicodeMath(linear);
  if (hasLiteralLatex(list)) return null;
  const latex = importedLatex(list);
  if (latex === null) return null;
  const exported = exportedModel(latex);
  if (!exported || canonicalMath(exported) !== canonicalMath(list)) return null;
  return { latex, linear };
}
