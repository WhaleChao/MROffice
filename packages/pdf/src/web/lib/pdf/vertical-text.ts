/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Laying out vertical CJK text.
 *
 * This is not a rotated box. Rotating a box 90 degrees lays every character on
 * its side; vertical writing keeps the glyphs upright, advances DOWN the
 * column, and stacks columns RIGHT to LEFT. The two are different enough that
 * sharing an implementation would get both wrong.
 *
 * Latin runs are the exception: with no tate-chu-yoko, the convention is to
 * turn a Latin run a quarter turn clockwise and let it read down the column.
 * That is what a reader of a vertical Chinese or Japanese page expects to see
 * for "PDF" or "2026".
 *
 * Positions are in points from the box's TOP-RIGHT corner, with `x` growing
 * LEFT — the direction columns actually advance. The caller converts to
 * whatever space it draws in; keeping the model in the writing's own directions
 * is what makes the rules above readable.
 */

/** Advance between columns, as a multiple of the font size. */
export const VERTICAL_COLUMN_RATIO = 1.5;

export interface VerticalGlyphRun {
  /** The characters drawn as one unit — one upright glyph, or a Latin run. */
  text: string;
  /** 0 is the rightmost column. */
  column: number;
  /** Distance LEFT from the box's right edge to the run's column centre. */
  x: number;
  /** Distance DOWN from the box's top to the run's start. */
  y: number;
  /** Drawn turned a quarter clockwise (a Latin run, or a bracket). */
  rotated: boolean;
  /** How far along the column this run advances, in points. */
  advance: number;
  /**
   * Shift within the em box, as a fraction of the font size: `+x` toward the
   * column's right edge, `+y` upward. Corner punctuation uses it — a comma or
   * full stop belongs in the TOP-RIGHT of its em box in vertical writing,
   * where a horizontal one sits bottom-left.
   */
  offsetX?: number;
  offsetY?: number;
  /**
   * Tate-chu-yoko: a short Latin run kept UPRIGHT and fitted into one em, so a
   * year or a page number reads normally instead of lying on its side. The
   * value is the scale to draw it at.
   */
  squeeze?: number;
}

export interface VerticalLayoutOptions {
  fontSize: number;
  /** Column length in points; a column that would overrun starts a new one. */
  boxHeight: number;
  /** Overrides the default column advance. */
  columnAdvance?: number;
  /**
   * Width of a Latin run in points at `fontSize`. The caller has the font, so
   * it measures; without this a run falls back to half an em per character,
   * which is close for digits and wrong for prose.
   */
  measure?: (text: string, fontSize: number) => number;
  /** Stop after this many columns. */
  maxColumns?: number;
}

export interface VerticalLayout {
  runs: VerticalGlyphRun[];
  /** Columns actually used. */
  columns: number;
  /** Text that did not fit within `maxColumns`. */
  overflow: string;
}

/**
 * Characters that stay upright in vertical writing.
 *
 * CJK ideographs, kana, Hangul and full-width forms are upright; so is
 * ideographic punctuation. Everything else — Latin, digits, ASCII punctuation —
 * turns with its run.
 */
const UPRIGHT = /[ᄀ-ᇿ⺀-〿぀-ヿ㄰-㆏㇀-鿿ꥠ-꥿가-퟿豈-﫿︐-︟︰-﹏＀-￯]/u;

/**
 * Ideographic punctuation that turns with the column.
 *
 * Brackets, dashes and ellipses are drawn along the writing direction, so in a
 * vertical column they lie the other way from how they look horizontally.
 * Unicode has dedicated vertical forms (U+FE35 and friends), but a font that
 * lacks them draws a notdef box — worse than the character it replaced — so
 * these are ROTATED instead, which no font can fail at.
 */
const ROTATE_IN_VERTICAL = new Set(
  Array.from("（）｛｝［］〈〉《》「」『』【】〔〕〖〗ー～－—―…‥＝"),
);

/**
 * Punctuation that moves to the top-right of its em box in vertical writing.
 *
 * Horizontally these sit at the bottom-left; vertically the convention puts
 * them in the opposite corner. Leaving them centred is the giveaway that a
 * renderer is only stacking horizontal glyphs.
 */
const CORNER_PUNCTUATION = new Set(Array.from("、。，．"));

/**
 * Small kana (拗音・促音) shift toward the TOP-RIGHT of their em in vertical
 * writing, by convention about a quarter em — less than punctuation's half,
 * because they are letters, not marks. Centred small kana are the second
 * giveaway (after centred 、) that a renderer only stacks horizontal glyphs.
 */
const SMALL_KANA = new Set(
  Array.from("ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ"),
);

/** Characters kept upright and squeezed into one em when a run is short. */
const TATE_CHU_YOKO = /^[0-9A-Za-z%°]{1,2}$/;

/** Longest Latin run that stays upright as tate-chu-yoko. */
export const TATE_CHU_YOKO_MAX = 2;


/** Whether a character keeps its upright orientation in a vertical column. */
export function isUprightInVerticalText(char: string): boolean {
  return UPRIGHT.test(char);
}

/** An explicit line break starts a new column, as it does horizontally. */
function isBreak(char: string): boolean {
  return char === "\n" || char === "\r";
}

type RunKind =
  | "upright"
  | "rotate"
  | "corner"
  | "small-kana"
  | "tcy"
  | "break";

/**
 * Split into runs.
 *
 * One run per upright character; consecutive non-ideographic characters group
 * so a Latin word turns as a word rather than letter by letter — unless the
 * group is short enough for tate-chu-yoko, in which case it stays upright.
 */
function runsOf(text: string): Array<{ text: string; kind: RunKind }> {
  const out: Array<{ text: string; kind: RunKind }> = [];
  let pending = "";
  const flush = () => {
    if (!pending) return;
    out.push({
      text: pending,
      // "12" or "AD" reads upright, squeezed into one em; anything longer
      // turns, because squeezing it would make it unreadable.
      kind: TATE_CHU_YOKO.test(pending) ? "tcy" : "rotate",
    });
    pending = "";
  };
  for (const char of Array.from(text)) {
    if (isBreak(char)) {
      flush();
      out.push({ text: char, kind: "break" });
      continue;
    }
    if (isUprightInVerticalText(char)) {
      flush();
      out.push({
        text: char,
        kind: ROTATE_IN_VERTICAL.has(char)
          ? "rotate"
          : CORNER_PUNCTUATION.has(char)
            ? "corner"
            : SMALL_KANA.has(char)
              ? "small-kana"
              : "upright",
      });
      continue;
    }
    // A space between two Latin words belongs to the same run; a space between
    // ideographs is its own upright advance.
    pending += char;
  }
  flush();
  return out;
}

export function layoutVerticalText(
  text: string,
  options: VerticalLayoutOptions,
): VerticalLayout {
  const { fontSize, boxHeight } = options;
  const empty: VerticalLayout = { runs: [], columns: 0, overflow: "" };
  if (!text || !(fontSize > 0) || !(boxHeight > 0)) return empty;

  const columnAdvance =
    options.columnAdvance && options.columnAdvance > 0
      ? options.columnAdvance
      : fontSize * VERTICAL_COLUMN_RATIO;
  const maxColumns = Math.max(
    1,
    options.maxColumns ?? Number.MAX_SAFE_INTEGER,
  );
  const measure =
    options.measure ?? ((run: string, size: number) => run.length * size * 0.5);

  const runs: VerticalGlyphRun[] = [];
  let column = 0;
  let y = 0;
  let overflowFrom = -1;
  const pieces = runsOf(text);

  for (let index = 0; index < pieces.length; index += 1) {
    const piece = pieces[index];
    if (piece.kind === "break") {
      column += 1;
      y = 0;
      if (column >= maxColumns) {
        overflowFrom = index + 1;
        break;
      }
      continue;
    }
    const rotated = piece.kind === "rotate";
    // Everything upright takes one em, including a squeezed tate-chu-yoko run:
    // fitting it into the column's own rhythm is the point of squeezing it.
    const advance = rotated
      ? Math.max(0, measure(piece.text, fontSize))
      : fontSize;
    // A run longer than the whole column would loop forever if it always
    // wrapped; it starts a column of its own and is allowed to overrun.
    if (y > 0 && y + advance > boxHeight + 1e-9) {
      column += 1;
      y = 0;
      if (column >= maxColumns) {
        overflowFrom = index;
        break;
      }
    }
    const measured =
      piece.kind === "tcy" ? Math.max(0, measure(piece.text, fontSize)) : 0;
    runs.push({
      text: piece.text,
      column,
      // Columns advance leftward from the right edge; the centre of the first
      // column sits half an advance in.
      x: column * columnAdvance + columnAdvance / 2,
      y,
      rotated,
      advance,
      // Half an em up and right moves the ink from the em box's bottom-left
      // quadrant, where a horizontal comma lives, to its top-right; small
      // kana take a quarter, being letters rather than marks.
      ...(piece.kind === "corner"
        ? { offsetX: 0.5, offsetY: 0.5 }
        : piece.kind === "small-kana"
          ? { offsetX: 0.25, offsetY: 0.25 }
          : {}),
      ...(piece.kind === "tcy" && measured > fontSize
        ? { squeeze: fontSize / measured }
        : piece.kind === "tcy"
          ? { squeeze: 1 }
          : {}),
    });
    y += advance;
  }

  return {
    runs,
    columns: runs.length === 0 ? 0 : runs[runs.length - 1].column + 1,
    overflow:
      overflowFrom < 0
        ? ""
        : pieces
            .slice(overflowFrom)
            .map((piece) => piece.text)
            .join(""),
  };
}
