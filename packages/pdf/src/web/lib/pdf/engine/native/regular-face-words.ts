/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where the glyphs of a box drawn in a bold face stand when its words are
 * set again in the family's regular face: the arithmetic of the paragraph
 * rewrite's "bold face set regular where its words stand"
 * (paragraph-rewrite.ts), kept apart from it so that it can be held to its
 * rules on numbers alone.
 *
 * Text the reader reorders (Arabic, Urdu, Hebrew) is drawn glyph by glyph in
 * the order the eye meets it, each glyph a presentation form the producer
 * chose. Its bold is taken off by drawing the same forms of the regular face
 * -- and the regular face has other widths (pdf-walk-1000 0979: 95 of the
 * bold Times New Roman's 95 Arabic glyphs, by -0.087 to +0.075 em). Kept on
 * the bold face's origins, joined letters come apart and letters that do not
 * join run into one another (71 and 33 of 477 joins on that file's first six
 * pages). So a word is set as the regular face sets it: its glyphs one after
 * another at the regular face's own advances, no spacing between them, which
 * is how the face joins them; and the word stands where it stood, its ink
 * centred where the bold word's ink was centred, so the gaps either side of
 * it change by half the difference each.
 *
 * Nothing is written unless every word of the box can be set so:
 *  - every pair of neighbours in a word that the bold face joins is joined
 *    by the regular face, and no pair it does not join runs into one another
 *    further than the bold face's own design sets them;
 *  - no word's edge moves more than `MAX_EDGE_MOVE_POINTS`;
 *  - no gap between a word and its neighbour on the line -- a word of the
 *    box, or other text of the page -- closes to less than half of what it
 *    was;
 *  - no word runs on into text the box does not hold (a word an elongation
 *    cut in two boxes for the reader): half of it set again would join the
 *    other half in another weight. (A full stop or a Latin word that stands
 *    against it joins nothing: its gap is held like any other.)
 * Whether two glyphs join is read off their forms, not off the gap the page
 * leaves between them: a letter that joins the next reaches its advance, one
 * that is joined starts before its origin. (By the gap alone, a producer's
 * own character spacing closes a join to a hair or opens a pair that never
 * joined to a touch: 0979 at ±0.06 pt.)
 */

/** A glyph's advance and where its ink starts and ends from its origin, in
 * thousandths of the em. */
export type PdfRegularFaceGlyphMetrics = Readonly<{ advance: number; left: number; right: number }>;

/** One glyph of the box as the page draws it, and the regular face's glyph
 * of the same form. */
export type PdfRegularFaceGlyph = Readonly<{
  /** The line it is drawn on (any number that is one line's alone). */
  line: number;
  /** Where its origin stands along the line, in points. */
  at: number;
  /** One em of it along the line, in points (its size on the page). */
  em: number;
  bold: PdfRegularFaceGlyphMetrics;
  regular: PdfRegularFaceGlyphMetrics;
}>;

/** Other ink on one of the box's lines that stays where it is: text of the
 * page the box does not hold, or text of the box that is not set again.
 * `joinsBefore`: its first glyph along the line is a form joined to the
 * glyph before it; `joinsAfter`: its last reaches on to join the next (a
 * full stop, a Latin word, a digit: neither). */
export type PdfRegularFaceFixedInk = Readonly<{ line: number; left: number; right: number; joinsBefore: boolean; joinsAfter: boolean }>;

export type PdfRegularFaceWordsInput = Readonly<{
  glyphs: readonly PdfRegularFaceGlyph[];
  /** Where each blank the box draws stands (a space shown between two words). */
  blanks: readonly Readonly<{ line: number; at: number }>[];
  fixed: readonly PdfRegularFaceFixedInk[];
}>;

/** Why the words are not set again. */
export type PdfRegularFaceWordsRefusal =
  | "no-glyph" | "join-forms-differ" | "join-opens" | "glyphs-run-together" | "edge-moves-too-far" | "gap-closes"
  | "word-runs-out-of-the-box";

export type PdfRegularFaceWordsPlan =
  | Readonly<{ ok: true; /** Each glyph's new origin along its line, in the order given. */ at: readonly number[];
    words: number; pairs: number; joins: number; /** The furthest an edge of a word moves, in points. */ edgeMove: number }>
  | Readonly<{ ok: false; reason: PdfRegularFaceWordsRefusal; /** What stood in the way, for a probe. */ detail?: string }>;

/** How far apart two glyphs' ink stands for them to be two words, as a share
 * of the em (a space is a quarter of it and more; the widest gap inside a
 * word -- an alef beside a lam -- a tenth). */
export const PDF_REGULAR_FACE_WORD_GAP_EM = 0.18;
/** How far an edge of a word may move, in points. */
export const PDF_REGULAR_FACE_MAX_EDGE_MOVE_POINTS = 2;
/** The least of a gap between two words that stays. */
export const PDF_REGULAR_FACE_MIN_GAP_SHARE = 0.5;
/** How short of its advance a glyph's ink may end and still join the next,
 * and how far two joined glyphs overlap at the least, in thousandths. */
const JOIN_REACH = 2;
const JOIN_OVERLAP = 5;

/** Whether the face joins `before` to `after`, the glyph drawn next along the line. */
export function pdfRegularFaceGlyphsJoin(before: PdfRegularFaceGlyphMetrics, after: PdfRegularFaceGlyphMetrics): boolean {
  return before.right >= before.advance - JOIN_REACH && after.left <= -JOIN_OVERLAP;
}

export function planPdfRegularFaceWords(input: PdfRegularFaceWordsInput): PdfRegularFaceWordsPlan {
  const { glyphs, blanks, fixed } = input;
  if (!glyphs.length) return { ok: false, reason: "no-glyph" };
  const at = glyphs.map((glyph) => glyph.at);
  const unit = (glyph: PdfRegularFaceGlyph) => glyph.em / 1000;
  const boldInk = (index: number) => {
    const glyph = glyphs[index];
    return [glyph.at + glyph.bold.left * unit(glyph), glyph.at + glyph.bold.right * unit(glyph)] as const;
  };
  // A glyph and the regular face's glyph for it are one form: each joins on
  // the sides the other joins on (the letter's isolated form is not set in
  // the place of its final one because both read as the same character).
  const joinsBefore = (glyph: PdfRegularFaceGlyphMetrics) => glyph.left <= -JOIN_OVERLAP;
  const joinsAfter = (glyph: PdfRegularFaceGlyphMetrics) => glyph.right >= glyph.advance - JOIN_REACH;
  if (glyphs.some((glyph) => joinsBefore(glyph.bold) !== joinsBefore(glyph.regular) ||
    joinsAfter(glyph.bold) !== joinsAfter(glyph.regular))) return { ok: false, reason: "join-forms-differ" };
  let words = 0, pairs = 0, joins = 0, edgeMove = 0;
  for (const line of new Set(glyphs.map((glyph) => glyph.line))) {
    const row = glyphs.flatMap((glyph, index) => (glyph.line === line ? [index] : []))
      .sort((left, right) => glyphs[left].at - glyphs[right].at || left - right);
    const cuts = blanks.filter((blank) => blank.line === line).map((blank) => blank.at);
    // Its words: neighbours along the line with no blank drawn between them
    // and their ink no further apart than a word holds its letters.
    const rowWords: number[][] = [];
    for (const index of row) {
      const word = rowWords.at(-1);
      const before = word?.at(-1);
      if (word && before !== undefined) {
        const apart = boldInk(index)[0] - boldInk(before)[1];
        const blank = cuts.some((cut) => cut > glyphs[before].at + 1e-6 && cut <= glyphs[index].at + 1e-6);
        if (!blank && apart <= PDF_REGULAR_FACE_WORD_GAP_EM * Math.max(glyphs[before].em, glyphs[index].em)) {
          word.push(index);
          continue;
        }
      }
      rowWords.push([index]);
    }
    const still = fixed.filter((ink) => ink.line === line);
    /** Each word's ink as it stood and as it is set, along the line. */
    const extents: { was: readonly [number, number]; is: readonly [number, number]; em: number }[] = [];
    for (const word of rowWords) {
      words += 1;
      // As the regular face sets it: each glyph where the one before ends.
      const set = [0];
      for (let order = 0; order < word.length - 1; order += 1) {
        const glyph = glyphs[word[order]];
        set.push(set[order] + glyph.regular.advance * unit(glyph));
      }
      for (let order = 0; order < word.length - 1; order += 1) {
        const before = glyphs[word[order]], after = glyphs[word[order + 1]];
        pairs += 1;
        const joined = pdfRegularFaceGlyphsJoin(before.bold, after.bold);
        if (joined !== pdfRegularFaceGlyphsJoin(before.regular, after.regular)) return { ok: false, reason: "join-forms-differ" };
        // (In thousandths of the em of the glyph drawn first.)
        const gap = (set[order + 1] + after.regular.left * unit(after) - (set[order] + before.regular.right * unit(before))) /
          unit(before);
        if (joined) {
          joins += 1;
          if (gap > -JOIN_OVERLAP) return { ok: false, reason: "join-opens" };
        } else {
          // Not into one another, or no further than the bold face's own
          // design sets the two.
          const designed = before.bold.advance - before.bold.right + after.bold.left * (unit(after) / unit(before));
          if (gap < Math.min(0, designed) - 1e-6) return { ok: false, reason: "glyphs-run-together" };
        }
      }
      const wasLeft = Math.min(...word.map((index) => boldInk(index)[0]));
      const wasRight = Math.max(...word.map((index) => boldInk(index)[1]));
      const setLeft = Math.min(...word.map((index, order) => set[order] + glyphs[index].regular.left * unit(glyphs[index])));
      const setRight = Math.max(...word.map((index, order) => set[order] + glyphs[index].regular.right * unit(glyphs[index])));
      // Its ink centred where the bold word's was.
      const origin = (wasLeft + wasRight) / 2 - (setLeft + setRight) / 2;
      for (const [order, index] of word.entries()) at[index] = origin + set[order];
      const is = [origin + setLeft, origin + setRight] as const;
      const moved = Math.max(Math.abs(is[0] - wasLeft), Math.abs(is[1] - wasRight));
      if (moved > PDF_REGULAR_FACE_MAX_EDGE_MOVE_POINTS + 1e-9) return { ok: false, reason: "edge-moves-too-far" };
      edgeMove = Math.max(edgeMove, moved);
      const em = Math.max(...word.map((index) => glyphs[index].em));
      // A word that runs on into ink the box does not set again.
      // A word that runs on into ink the box does not set again: a glyph
      // there joined to the word's last, or its first joined to one there.
      const firstBold = glyphs[word[0]].bold, lastBold = glyphs[word.at(-1)!].bold;
      const reach = PDF_REGULAR_FACE_WORD_GAP_EM * em;
      const blankIn = (from: number, to: number) => cuts.some((cut) => cut >= from - 1e-6 && cut <= to + 1e-6);
      const beside = still.find((ink) =>
        (ink.left >= wasRight - reach && ink.left <= wasRight + reach && ink.joinsBefore && joinsAfter(lastBold) &&
          !blankIn(wasRight, ink.left)) ||
        (ink.right <= wasLeft + reach && ink.right >= wasLeft - reach && ink.joinsAfter && joinsBefore(firstBold) &&
          !blankIn(ink.right, wasLeft)));
      if (beside) {
        return { ok: false, reason: "word-runs-out-of-the-box",
          detail: `word ${wasLeft.toFixed(2)}..${wasRight.toFixed(2)} beside ink ${beside.left.toFixed(2)}..${beside.right.toFixed(2)}` };
      }
      extents.push({ was: [wasLeft, wasRight], is, em });
    }
    // The gaps of the line: between its words, and between a word and the
    // ink beside it that stays.
    const keeps = (was: number, is: number) => was <= 0 ? is >= was - 1e-9 : is >= was * PDF_REGULAR_FACE_MIN_GAP_SHARE - 1e-9;
    for (const [order, word] of extents.entries()) {
      const next = extents[order + 1];
      if (next && !keeps(next.was[0] - word.was[1], next.is[0] - word.is[1])) return { ok: false, reason: "gap-closes" };
      for (const ink of still) {
        if (ink.left >= word.was[1] && !keeps(ink.left - word.was[1], ink.left - word.is[1])) return { ok: false, reason: "gap-closes" };
        if (ink.right <= word.was[0] && !keeps(word.was[0] - ink.right, word.is[0] - ink.right)) return { ok: false, reason: "gap-closes" };
      }
    }
  }
  return { ok: true, at: Object.freeze(at), words, pairs, joins, edgeMove };
}
