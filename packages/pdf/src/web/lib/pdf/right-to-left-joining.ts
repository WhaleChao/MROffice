/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether two stretches of Arabic-script text, one read after the other, are
 * joined where they meet: the last letter of the first runs on into the
 * first letter of the second, as the letters of one word do. Read off the
 * characters alone -- what the page's reader has of a run -- for the box
 * builder, which keeps the pieces of one word in one box
 * (native-paragraph-edit-boxes).
 *
 * A PDF's font most often names its glyphs by their presentation forms
 * (U+FE70..FEFF, U+FB50..FBFF): the form says how the glyph joins -- an
 * initial or a medial form runs on, a medial or a final form is run on to.
 * A tatweel joins on both sides. Marks stand on their letter and are passed
 * over. Anything else joins nothing here: a digit, a punctuation mark, a
 * blank, a letter of another script, a ligature whose form is not listed --
 * and a letter the font names by itself (U+0600..06FF), whose form on the
 * page is then not known. Its class would say that it may join, not that
 * it does: where a face sets its words edge to edge (Nastaleeq, with no
 * space glyph between them) the last letter of one word and the first of
 * the next read as a pair that joins, and two words would be taken for one
 * (pdf-walk-1000 0339, 0944, 0948: «ششماہی» and «امتحان»).
 */

/** Presentation forms as their blocks list them: a letter's run of code
 * points and how many forms it has there. Two: isolated, final. Four:
 * isolated, final, initial, medial. One: isolated. */
const FORM_RUNS: readonly (readonly [number, number])[] = [
  // Arabic Presentation Forms-A (the letters; its ligatures are not read here).
  [0xfb50, 2], [0xfb52, 4], [0xfb56, 4], [0xfb5a, 4], [0xfb5e, 4], [0xfb62, 4], [0xfb66, 4], [0xfb6a, 4], [0xfb6e, 4],
  [0xfb72, 4], [0xfb76, 4], [0xfb7a, 4], [0xfb7e, 4], [0xfb82, 2], [0xfb84, 2], [0xfb86, 2], [0xfb88, 2], [0xfb8a, 2],
  [0xfb8c, 2], [0xfb8e, 4], [0xfb92, 4], [0xfb96, 4], [0xfb9a, 4], [0xfb9e, 2], [0xfba0, 4], [0xfba4, 2], [0xfba6, 4],
  [0xfbaa, 4], [0xfbae, 2], [0xfbb0, 2], [0xfbd3, 4], [0xfbd7, 2], [0xfbd9, 2], [0xfbdb, 2], [0xfbdd, 1], [0xfbde, 2],
  [0xfbe0, 2], [0xfbe2, 2], [0xfbe4, 4], [0xfbfc, 4],
  // Arabic Presentation Forms-B.
  [0xfe80, 1], [0xfe81, 2], [0xfe83, 2], [0xfe85, 2], [0xfe87, 2], [0xfe89, 4], [0xfe8d, 2], [0xfe8f, 4], [0xfe93, 2],
  [0xfe95, 4], [0xfe99, 4], [0xfe9d, 4], [0xfea1, 4], [0xfea5, 4], [0xfea9, 2], [0xfeab, 2], [0xfead, 2], [0xfeaf, 2],
  [0xfeb1, 4], [0xfeb5, 4], [0xfeb9, 4], [0xfebd, 4], [0xfec1, 4], [0xfec5, 4], [0xfec9, 4], [0xfecd, 4], [0xfed1, 4],
  [0xfed5, 4], [0xfed9, 4], [0xfedd, 4], [0xfee1, 4], [0xfee5, 4], [0xfee9, 4], [0xfeed, 2], [0xfeef, 2], [0xfef1, 4],
  // (Lam with alef: isolated, final.)
  [0xfef5, 2], [0xfef7, 2], [0xfef9, 2], [0xfefb, 2],
];
/** FBE8 and FBE9: the Uighur yeh's initial and medial forms alone. */
const LONE_FORMS = new Map<number, Readonly<{ on: boolean; to: boolean }>>([
  [0xfbe8, { on: true, to: false }], [0xfbe9, { on: true, to: true }],
]);

const TATWEEL = 0x0640;
const MARK = /^\p{M}$/u;

/** How one character joins: `on` -- it runs on into the letter read after
 * it; `to` -- the letter read before it runs on into it. Null for a mark. */
function joining(character: string): Readonly<{ on: boolean; to: boolean }> | null {
  if (MARK.test(character)) return null;
  const point = character.codePointAt(0)!;
  if (point === TATWEEL || point === 0x200d) return { on: true, to: true };
  const lone = LONE_FORMS.get(point);
  if (lone) return lone;
  if ((point >= 0xfb50 && point <= 0xfdff) || (point >= 0xfe70 && point <= 0xfeff)) {
    for (const [start, forms] of FORM_RUNS) {
      if (point < start || point >= start + forms) continue;
      const form = point - start;
      // isolated, final, initial, medial
      return { on: form >= 2, to: form === 1 || form === 3 };
    }
    return { on: false, to: false };
  }
  return { on: false, to: false };
}

/**
 * Whether `first` runs on into `second`, read after it: the last letter of
 * `first` and the first letter of `second` (the marks on them aside) are
 * joined to each other.
 */
export function pdfRightToLeftTextJoins(first: string, second: string): boolean {
  let last: ReturnType<typeof joining> = null;
  for (const character of [...first].reverse()) {
    last = joining(character);
    if (last) break;
  }
  let next: ReturnType<typeof joining> = null;
  for (const character of second) {
    next = joining(character);
    if (next) break;
  }
  return Boolean(last?.on && next?.to);
}
