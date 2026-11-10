/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which TeX engine a LaTeX compile runs, as the main file says it.
 *
 * The sandbox's wrapper (mrai-latex in docker/code-executor) reads the first
 * `% !TEX program = <engine>` (or `% !TEX TS-program = <engine>`) line among
 * the first 20 lines of main.tex and otherwise picks the engine from what the
 * document uses. Overleaf keeps the same choice as a project setting with
 * pdfLaTeX as its default; here the line itself is the setting, so it travels
 * with the file, is saved and undone like any edit, and a hand-written line
 * is read back the same way.
 *
 * {@link findLatexMagicComment} follows the wrapper's own pattern exactly
 * (the engine name is case-sensitive there), so the control never shows an
 * engine the compile would not run.
 */

export const LATEX_ENGINES = ["pdflatex", "xelatex", "lualatex", "latex"] as const;
export type LatexEngine = (typeof LATEX_ENGINES)[number];
/** "auto": no magic comment, the wrapper detects the engine. */
export type LatexCompilerChoice = "auto" | LatexEngine;
export const LATEX_COMPILER_CHOICES: readonly LatexCompilerChoice[] = [
  "auto",
  ...LATEX_ENGINES,
];

/** The wrapper reads `head -n 20 main.tex`. */
const MAGIC_COMMENT_LINES = 20;
// The wrapper's sed expression: [[:space:]] within a line, "TeX" in any case,
// at least one space after it, "TS-" optional, "program" in lower case, and
// the engine as the run of letters after "=".
const MAGIC_COMMENT =
  /^[^\S\n]*%[^\S\n]*![^\S\n]*[Tt][Ee][Xx][^\S\n]+(?:TS-)?program[^\S\n]*=[^\S\n]*([A-Za-z]*)/;

export interface LatexMagicComment {
  /** The engine as written; empty when the line names none. */
  readonly program: string;
  /** The line, without its line break. */
  readonly lineFrom: number;
  readonly lineTo: number;
  /** The engine name within the line. */
  readonly valueFrom: number;
  readonly valueTo: number;
}

export interface LatexCompilerChange {
  readonly from: number;
  readonly to: number;
  readonly insert: string;
}

function isLatexEngine(value: string): value is LatexEngine {
  return (LATEX_ENGINES as readonly string[]).includes(value);
}

/** The first magic comment the wrapper would read, or null. */
export function findLatexMagicComment(source: string): LatexMagicComment | null {
  let lineFrom = 0;
  for (let index = 0; index < MAGIC_COMMENT_LINES && lineFrom <= source.length; index += 1) {
    const newline = source.indexOf("\n", lineFrom);
    const lineEnd = newline < 0 ? source.length : newline;
    const line = source.slice(lineFrom, lineEnd);
    const match = MAGIC_COMMENT.exec(line);
    if (match) {
      const program = match[1] ?? "";
      const valueTo = lineFrom + match[0].length;
      return {
        program,
        lineFrom,
        lineTo: lineEnd > lineFrom && source.charCodeAt(lineEnd - 1) === 13 ? lineEnd - 1 : lineEnd,
        valueFrom: valueTo - program.length,
        valueTo,
      };
    }
    if (newline < 0) break;
    lineFrom = newline + 1;
  }
  return null;
}

/**
 * The compiler the main file asks for: "auto" with no magic comment (or one
 * that names no engine), null when it names an engine the compile does not
 * offer (the compile then stops and says so).
 */
export function latexCompilerChoice(source: string): LatexCompilerChoice | null {
  const comment = findLatexMagicComment(source);
  if (!comment || comment.program === "") return "auto";
  return isLatexEngine(comment.program) ? comment.program : null;
}

/**
 * The one edit that makes `source` ask for `choice`, or null when it already
 * does. An engine replaces the name in the existing line (keeping how the
 * line is written) or adds `% !TEX program = <engine>` as the first line;
 * Auto removes the line with its line break.
 */
export function latexCompilerChange(
  source: string,
  choice: LatexCompilerChoice,
): LatexCompilerChange | null {
  const comment = findLatexMagicComment(source);
  if (choice === "auto") {
    if (!comment) return null;
    if (source.startsWith("\r\n", comment.lineTo)) {
      return { from: comment.lineFrom, to: comment.lineTo + 2, insert: "" };
    }
    if (source.startsWith("\n", comment.lineTo)) {
      return { from: comment.lineFrom, to: comment.lineTo + 1, insert: "" };
    }
    // The last line: take the line break before it instead.
    const before =
      comment.lineFrom >= 2 && source.startsWith("\r\n", comment.lineFrom - 2)
        ? 2
        : comment.lineFrom >= 1
          ? 1
          : 0;
    return { from: comment.lineFrom - before, to: comment.lineTo, insert: "" };
  }
  if (comment) {
    if (comment.program === choice) return null;
    return { from: comment.valueFrom, to: comment.valueTo, insert: choice };
  }
  const lineBreak = /\r\n/.test(source.slice(0, source.indexOf("\n") + 1)) ? "\r\n" : "\n";
  return { from: 0, to: 0, insert: `% !TEX program = ${choice}${lineBreak}` };
}
