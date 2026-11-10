/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** PowerPoint's advance grid: 1/8 pt, in CSS px (96 per inch, 72 pt). */
export const POWERPOINT_ADVANCE_QUANTUM_PX = 96 / 72 / 8;

/** One glyph as the paragraph draws it. */
export interface LineFitGlyph {
  readonly ch: string;
  /** Its advance alone, in CSS px (no kerning against its neighbours). */
  readonly advance: number;
}

/** A place a line may end after: a word with its trailing spaces. */
export interface LineFitToken {
  /** Index of its first glyph. */
  readonly start: number;
  /** Index after its last glyph, trailing spaces included. */
  readonly end: number;
  /** Index after its last non-space glyph (the part a line end counts). */
  readonly inkEnd: number;
  /** A hard line break follows it. */
  readonly hardBreak: boolean;
}

const SPACE = /[ \t\u00a0\u3000]/;
// A line may end between any two East Asian characters.
const EAST_ASIAN =
  /[\u1100-\u11ff\u2e80-\u303f\u3040-\u30ff\u3100-\u31ff\u3400-\u4dbf\u4e00-\u9fff\ua960-\ua97f\uac00-\ud7af\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef]/u;
// Characters the browser may end a line after, besides spaces (a hyphen,
// a dash, a slash): inside a word, only a hyphen between letters is
// modelled (a break after it); a paragraph with any other is not.
const OTHER_BREAKS = /[-\u2010\u2012-\u2014\/?!}\]|]/;
const HYPHENS = /[-\u2010]/;
const LETTER = /\p{L}/u;
const DIGIT = /\p{Nd}/u;

/** The paragraph's words, split where a line may end. */
export function lineFitTokens(glyphs: readonly LineFitGlyph[]): LineFitToken[] | null {
  const tokens: LineFitToken[] = [];
  let index = 0;
  while (index < glyphs.length) {
    const start = index;
    if (glyphs[index].ch === "\n") {
      tokens.push({ start, end: index + 1, inkEnd: start, hardBreak: true });
      index += 1;
      continue;
    }
    if (EAST_ASIAN.test(glyphs[index].ch)) {
      index += 1;
    } else {
      while (
        index < glyphs.length &&
        glyphs[index].ch !== "\n" &&
        !SPACE.test(glyphs[index].ch) &&
        !EAST_ASIAN.test(glyphs[index].ch)
      ) {
        const ch = glyphs[index].ch;
        index += 1;
        const next = glyphs[index]?.ch;
        // Ending its word (before a space, a break or the end), a mark is
        // part of the word.
        if (!OTHER_BREAKS.test(ch) || next === undefined || next === "\n" || SPACE.test(next)) continue;
        // A hyphen before a digit (a range, a negative number) is no break
        // for either (Unicode line breaking, LB25: HY x NU) ...
        if (HYPHENS.test(ch) && DIGIT.test(next)) continue;
        // ... and between two letters both break after it.
        if (HYPHENS.test(ch) && LETTER.test(glyphs[index - 2]?.ch ?? "") && LETTER.test(next)) break;
        // Any other break inside a word is the browser's own call.
        return null;
      }
    }
    const inkEnd = index;
    while (index < glyphs.length && SPACE.test(glyphs[index].ch)) index += 1;
    const hardBreak = index < glyphs.length && glyphs[index].ch === "\n";
    if (hardBreak) index += 1;
    tokens.push({ start, end: index, inkEnd, hardBreak });
  }
  return tokens;
}

function roundedAdvance(advance: number): number {
  return Math.round(advance / POWERPOINT_ADVANCE_QUANTUM_PX) * POWERPOINT_ADVANCE_QUANTUM_PX;
}

/** One line: the tokens `[first, last]` (inclusive). */
export interface LineFitLine {
  readonly first: number;
  readonly last: number;
}

/**
 * Where PowerPoint ends each line: tokens are added while the rounded
 * advances up to the new token's last inked glyph stay within the line
 * (trailing spaces hang past the edge). `null` when a single token is wider
 * than its line (PowerPoint and the browser then cut inside the word, each
 * its own way).
 */
export function powerPointLines(
  glyphs: readonly LineFitGlyph[],
  tokens: readonly LineFitToken[],
  width: number,
  firstLineIndent = 0,
): LineFitLine[] | null {
  const rounded = glyphs.map((glyph) => roundedAdvance(glyph.advance));
  const lines: LineFitLine[] = [];
  let first = 0;
  while (first < tokens.length) {
    const available = width - (lines.length === 0 ? firstLineIndent : 0);
    let used = 0;
    let last = first;
    for (let index = first; index < tokens.length; index += 1) {
      const token = tokens[index];
      let ink = used;
      for (let glyph = token.start; glyph < token.inkEnd; glyph += 1) ink += rounded[glyph];
      if (index > first && ink > available + 1e-6) break;
      if (index === first && ink > available + 1e-6) return null;
      last = index;
      for (let glyph = token.inkEnd; glyph < token.end; glyph += 1) {
        if (glyphs[glyph].ch !== "\n") ink += rounded[glyph];
      }
      used = ink;
      if (token.hardBreak) break;
    }
    lines.push({ first, last });
    first = last + 1;
  }
  return lines;
}

/**
 * The width range the browser's lines end where PowerPoint's do:
 * `[low, high)`, from the browser's own (exact, kerned) measure of a run of
 * glyphs. Every line must fit (`low`) and no line may take the next word
 * (`high`). `null` when no width does (the exact and rounded measures order
 * two lines differently).
 */
export function browserWidthRange(
  tokens: readonly LineFitToken[],
  lines: readonly LineFitLine[],
  measure: (start: number, end: number) => number,
  firstLineIndent = 0,
): { low: number; high: number } | null {
  let low = 0;
  let high = Number.POSITIVE_INFINITY;
  lines.forEach((line, index) => {
    const indent = index === 0 ? firstLineIndent : 0;
    const start = tokens[line.first].start;
    low = Math.max(low, indent + measure(start, tokens[line.last].inkEnd));
    const next = tokens[line.last + 1];
    if (next && !tokens[line.last].hardBreak) {
      high = Math.min(high, indent + measure(start, next.inkEnd));
    }
  });
  return low < high ? { low, high } : null;
}

/**
 * The width to give the browser's text column so its lines end where
 * PowerPoint's do: the current width when it already does, else the nearest
 * width inside the range. `null` when nothing should change.
 */
export function powerPointFitWidth(
  width: number,
  range: { low: number; high: number },
): number | null {
  if (range.low <= width + 1e-6 && width < range.high - 1e-6) return null;
  const margin = Math.min(0.02, (range.high - range.low) / 2);
  return width < range.low ? range.low + margin : range.high - margin;
}
