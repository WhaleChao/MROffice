/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The height of one line as CSS lays it out (CSS 2.1 §10.8), for a line the
 * page has already styled: the block's strut and each inline box on the line,
 * with the line-height the page computed for it, and each atom (an
 * inline-block such as a formula's placeholder) with the box it is given.
 *
 * Word's rules are not modelled here: the page states them as CSS (the auto
 * line spacing projection, the grid line heights, the formula placeholders),
 * and this reads the lines that CSS draws from them. A text box takes its
 * font's ascent and descent plus half the leading its line-height leaves on
 * each side; an atom takes its margin box, its baseline at its bottom edge
 * (an inline-block that holds no line, or hides its overflow), moved by its
 * vertical-align length. The line runs from the highest top to the lowest
 * bottom. No DOM, so it can run on a block nobody has laid out.
 */

/** A font's box on the line: its metrics at its size, px, ascent up. */
export interface CssTextBox {
  readonly kind: "text";
  readonly ascent: number;
  readonly descent: number;
  /** The font's own line gap, for `line-height: normal`. */
  readonly lineGap: number;
  /** Computed line-height in px, or "normal". */
  readonly lineHeight: number | "normal";
  /** vertical-align as a length (px, positive up); 0 for baseline. */
  readonly shift?: number;
}

/** An atomic inline box: its margin box and where its baseline is. */
export interface CssAtomBox {
  readonly kind: "atom";
  readonly height: number;
  /** How far its bottom edge stands below the line's baseline before
   * vertical-align (0: the bottom edge on the baseline). */
  readonly depth: number;
  readonly shift?: number;
}

export type CssLineItem = CssTextBox | CssAtomBox;

export interface CssLineBox {
  /** Extent above the baseline and below it, px. */
  readonly above: number;
  readonly below: number;
  readonly height: number;
}

function extent(item: CssLineItem): { above: number; below: number } {
  const shift = item.shift ?? 0;
  if (item.kind === "atom") return { above: item.height - item.depth + shift, below: item.depth - shift };
  const { ascent, descent } = item;
  const content = ascent + descent;
  const line = item.lineHeight === "normal" ? content + item.lineGap : item.lineHeight;
  const half = (line - content) / 2;
  return { above: ascent + half + shift, below: descent + half - shift };
}

/** The line box of a block's strut and the inline boxes on the line. */
export function cssLineBox(strut: CssTextBox, items: readonly CssLineItem[]): CssLineBox {
  let { above, below } = extent(strut);
  for (const item of items) {
    const box = extent(item);
    if (box.above > above) above = box.above;
    if (box.below > below) below = box.below;
  }
  return { above, below, height: above + below };
}
