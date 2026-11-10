/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's 竖排 (Text Direction > Vertical) for an MRAI text box: the box's
 * text set in columns down from the frame's top, stacking from its right
 * edge leftward, as the browser lays out `writing-mode: vertical-rl` with
 * `text-orientation: mixed`. Every writer of the box (the raster appearance,
 * the flattened export, the server's appearance in native_edit_primitives)
 * sets the same columns from these rules:
 *
 * - ideographs, kana, Hangul and full-width forms stand upright, one em
 *   down the column; a Latin word turns a quarter clockwise and reads down
 *   it; brackets, dashes and the long-vowel mark turn too;
 * - a column breaks between two upright characters or after a word's
 *   blanks, never before closing punctuation (the character before it goes
 *   along); a word longer than the column breaks between its letters;
 * - 、。，． stand in the top-right of their em, small kana a quarter em so.
 *
 * Imported vertical writing keeps its own layout (vertical-text.ts).
 */
import { isUprightInVerticalText } from "./vertical-text";

export type PdfTextBoxVerticalTokenKind = "upright" | "rotated" | "corner" | "small-kana";

export interface PdfTextBoxVerticalToken {
  readonly text: string;
  readonly kind: PdfTextBoxVerticalTokenKind;
  /** Along the column, at the size laid out. */
  readonly advance: number;
}

export interface PdfTextBoxVerticalColumn {
  readonly tokens: readonly PdfTextBoxVerticalToken[];
  /** The column ends its paragraph (Justify leaves it as it is). */
  readonly paragraphEnd: boolean;
}

const ROTATED = new Set(Array.from("（）｛｝［］〈〉《》「」『』【】〔〕〖〗ー～－—―…‥＝"));
const CORNER = new Set(Array.from("、。，．"));
const SMALL_KANA = new Set(Array.from("ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ"));
const NO_START = new Set(Array.from("、。，．,.)）」』】〕〉》！？：；!?:;…‥ーぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ"));

/** The em box's middle over the alphabetic baseline: an upright ideograph
 * and a turned word are both centred on their column by it. */
export const PDF_TEXT_BOX_VERTICAL_MIDDLE = 0.38;

export function layoutPdfTextBoxVerticalColumns(
  text: string,
  options: Readonly<{ measure: (value: string) => number; size: number; columnLength: number }>,
): PdfTextBoxVerticalColumn[] {
  const { measure, size } = options;
  const columnLength = Math.max(1, options.columnLength);
  const columns: PdfTextBoxVerticalColumn[] = [];
  for (const paragraph of text.replace(/\r\n?/gu, "\n").split("\n")) {
    const tokens: PdfTextBoxVerticalToken[] = [];
    let latin = "";
    const flush = () => {
      for (const word of latin.match(/\S+\s*|\s+/gu) ?? []) tokens.push({ text: word, kind: "rotated", advance: measure(word) });
      latin = "";
    };
    for (const character of Array.from(paragraph)) {
      if (isUprightInVerticalText(character)) {
        flush();
        const kind: PdfTextBoxVerticalTokenKind = ROTATED.has(character) ? "rotated"
          : CORNER.has(character) ? "corner" : SMALL_KANA.has(character) ? "small-kana" : "upright";
        tokens.push({ text: character, kind, advance: kind === "rotated" ? measure(character) : size });
      } else {
        latin += character;
      }
    }
    flush();
    let column: PdfTextBoxVerticalToken[] = [];
    let used = 0;
    let index = 0;
    while (index < tokens.length) {
      const token = tokens[index]!;
      // Trailing blanks hang past the column's foot, as the browser's do.
      const ink = token.kind === "rotated" ? measure(token.text.trimEnd()) : token.advance;
      if (column.length && used + ink > columnLength + 1e-6) {
        const carried: PdfTextBoxVerticalToken[] = [];
        if (NO_START.has(token.text.slice(0, 1)) && column.length > 1) carried.push(column.pop()!);
        columns.push({ tokens: column, paragraphEnd: false });
        column = carried;
        used = carried.reduce((sum, item) => sum + item.advance, 0);
        continue;
      }
      if (!column.length && ink > columnLength + 1e-6 && Array.from(token.text).length > 1) {
        let head = "";
        for (const character of Array.from(token.text)) {
          if (head && measure(head + character) > columnLength + 1e-6) break;
          head += character;
        }
        const rest = token.text.slice(head.length);
        tokens.splice(index, 1, { text: head, kind: token.kind, advance: measure(head) },
          { text: rest, kind: token.kind, advance: measure(rest) });
        continue;
      }
      column.push(token);
      used += token.advance;
      index += 1;
    }
    columns.push({ tokens: column, paragraphEnd: true });
  }
  return columns;
}

export interface PdfTextBoxVerticalGlyph {
  readonly text: string;
  /** Turned a quarter clockwise (its glyph tops point right). */
  readonly rotated: boolean;
  /** Where its alphabetic baseline starts, from the frame's top-left, y down:
   * an upright glyph's left end, a turned word's top end. */
  readonly x: number;
  readonly y: number;
}

/**
 * The glyphs of a vertical box placed in its frame (x right, y down from the
 * frame's top-left): each column's centre one pitch further left, Align on
 * the column (top, centre, foot, Justify spreading all but a paragraph's
 * last column), Align Text across the frame (right, middle, left).
 * `width` measures an upright glyph across.
 */
export function placePdfTextBoxVerticalColumns(
  columns: readonly PdfTextBoxVerticalColumn[],
  options: Readonly<{
    size: number;
    pitch: number;
    frameWidth: number;
    frameHeight: number;
    align: "left" | "center" | "right" | "justify";
    verticalAlign?: "middle" | "bottom";
    measure: (value: string) => number;
  }>,
): Readonly<{ glyphs: PdfTextBoxVerticalGlyph[]; bands: Readonly<{ centre: number; top: number; length: number }>[] }> {
  const { size, pitch, frameWidth, frameHeight, align, verticalAlign, measure } = options;
  const spare = frameWidth - columns.length * pitch;
  const shift = !(spare > 0) ? 0 : verticalAlign === "middle" ? spare / 2 : verticalAlign === "bottom" ? spare : 0;
  const right = frameWidth - shift;
  const middle = PDF_TEXT_BOX_VERTICAL_MIDDLE * size;
  const glyphs: PdfTextBoxVerticalGlyph[] = [];
  const bands: { centre: number; top: number; length: number }[] = [];
  columns.forEach((column, index) => {
    const centre = right - (index + 0.5) * pitch;
    const tokens = column.tokens.map((token, at) => {
      if (at !== column.tokens.length - 1 || token.kind !== "rotated" || token.text === token.text.trimEnd()) return token;
      const text = token.text.trimEnd();
      return { ...token, text, advance: measure(text) };
    });
    const used = tokens.reduce((sum, token) => sum + token.advance, 0);
    const room = Math.max(0, frameHeight - used);
    let start = 0;
    let gap = 0;
    if (align === "center") start = room / 2;
    else if (align === "right") start = room;
    else if (align === "justify" && !column.paragraphEnd && tokens.length > 1) gap = room / (tokens.length - 1);
    let y = start;
    for (const token of tokens) {
      if (token.kind === "rotated") {
        glyphs.push({ text: token.text, rotated: true, x: centre - middle, y });
      } else {
        const lift = token.kind === "corner" ? 0.5 * size : token.kind === "small-kana" ? 0.25 * size : 0;
        glyphs.push({
          text: token.text, rotated: false,
          x: centre - measure(token.text) / 2 + lift,
          y: y + size * 0.5 + middle - lift,
        });
      }
      y += token.advance + gap;
    }
    if (tokens.length) bands.push({ centre, top: start, length: used + gap * Math.max(0, tokens.length - 1) });
  });
  return { glyphs, bands };
}
