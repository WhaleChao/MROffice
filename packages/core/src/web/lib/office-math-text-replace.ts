/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A replacement inside an equation's characters, made in its LaTeX too.
 *
 * An equation run states its characters (`text`, the equation read out:
 * "Dose=weight×15 mg/kg") and its LaTeX ("Dose=weight\times 15 mg/kg"). Find
 * & Replace and the AI's text edits find words in the characters; a change
 * there is the same change in the LaTeX when the replaced characters are
 * spelled there exactly once per occurrence, in order ("×" as `\times`,
 * "%" as `\%`). Anything else (a fraction's bar, a character LaTeX spells as
 * structure) is not followed: the caller leaves the equation as it is and
 * says so, never writing the characters as plain text in its place.
 *
 * The same rule as the server's deck studio (`_math_run_edited`,
 * presentation_artifact_studio.py), so an equation edited by the AI and one
 * replaced in the editor read alike.
 */

const MATH_TEXT_LATEX: Readonly<Record<string, string>> = {
  "×": "\\times", "÷": "\\div", "·": "\\cdot", "−": "-", "±": "\\pm", "≤": "\\le", "≥": "\\ge",
  "≠": "\\ne", "%": "\\%", "&": "\\&", "#": "\\#", $: "\\$", _: "\\_",
};

/** Characters a replacement may hold inside an equation (no LaTeX syntax of its own). */
const MATH_TEXT_SAFE = /^[0-9A-Za-z.,:;!?'()[\]+=<>/|\- ·×÷−±≤≥≠%&#$_一-鿿]*$/u;

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function fragmentPattern(fragment: string, spacesInsignificant = false): string {
  return Array.from(fragment)
    .map((character) => {
      const latex = MATH_TEXT_LATEX[character];
      if (latex !== undefined) return escapeRegExp(latex) + (/[a-z]$/i.test(latex) ? "\\s*" : "");
      if (character === " ") return "\\s*";
      return escapeRegExp(character);
    })
    .join(spacesInsignificant ? "\\s*" : "");
}

function fragmentLatex(fragment: string): string {
  return Array.from(fragment)
    .map((character) => {
      const latex = MATH_TEXT_LATEX[character];
      if (latex === undefined) return character;
      return /[a-z]$/i.test(latex) ? `${latex} ` : latex;
    })
    .join("");
}

function occurrences(text: string, fragment: string, before = text.length): number {
  let count = 0;
  for (let at = text.indexOf(fragment); at >= 0 && at < before; at = text.indexOf(fragment, at + fragment.length)) count += 1;
  return count;
}

/**
 * The equation (`text`, `latex`) with `text[start, end)` replaced by
 * `replacement`, the LaTeX changed to match; null when the LaTeX cannot
 * follow (the equation is then left as it is).
 *
 * `spacesInsignificant`: `text` is the equation as drawn, where a space in
 * the LaTeX's math mode draws nothing ("15 mg" draws "15mg", Word's math
 * runs read the same), so the replaced characters are found in the LaTeX
 * with any spacing between them. The caller checks the result draws the
 * expected characters.
 */
export function mathTextReplaced(
  text: string,
  latex: string,
  start: number,
  end: number,
  replacement: string,
  options: { spacesInsignificant?: boolean } = {},
): { text: string; latex: string } | null {
  const old = text.slice(start, end);
  if (!old || !MATH_TEXT_SAFE.test(old) || !MATH_TEXT_SAFE.test(replacement)) return null;
  const occurrence = occurrences(text, old, start);
  const found = Array.from(latex.matchAll(new RegExp(fragmentPattern(old, options.spacesInsignificant === true), "gu")));
  if (found.length !== occurrences(text, old) || occurrence >= found.length) return null;
  const match = found[occurrence];
  const matchStart = match.index ?? 0;
  const matchEnd = matchStart + match[0].length;
  let spelled = fragmentLatex(replacement);
  if (spelled.endsWith(" ") && !/^[A-Za-z]/.test(latex.slice(matchEnd, matchEnd + 1))) spelled = spelled.trimEnd();
  return {
    text: text.slice(0, start) + replacement + text.slice(end),
    latex: latex.slice(0, matchStart) + spelled + latex.slice(matchEnd),
  };
}
