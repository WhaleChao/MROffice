/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { hangingMarginTabTarget } from "../document-tabstops";

import { KINSOKU_EAST_ASIAN, type KinsokuList } from "./kinsoku";
import { lineBreakOpportunities } from "./line-break";
import { symbolAdvanceEm } from "./symbol-advances";

/** PowerPoint's advance grid: 1/8 pt. */
export const OFFICE_ADVANCE_QUANTUM_PT = 1 / 8;
/** Office's single line, as a multiple of the type size. */
export const OFFICE_SINGLE_LINE = 1.2;
/**
 * Where PowerPoint's baseline stands in a spaced line at the least: this much
 * of its pitch under its top. PowerPoint for Mac's P1-pitch PDF (Arial 8-32
 * pt at 150 %, every baseline on a whole point): 3/4 of the pitch for all 15
 * sizes, where 0.7547 put the 24, 28 and 32 pt first lines a point low.
 */
export const OFFICE_BASELINE_SHARE = 0.75;
/** A table cell's spaced line ({@link OfficeTextBodyInput.cellLines}). */
const CELL_BASELINE_SHARE = 0.7547;
/** Where PowerPoint's baseline stands in an exactly spaced line. */
export const OFFICE_EXACT_BASELINE_SHARE = 0.75;

export interface OfficeTextFont {
  /** CSS font-family list the text is drawn in. */
  readonly stack: string;
  readonly weight: number;
  readonly italic: boolean;
  readonly sizePt: number;
  /** Pair kerning from the face (a run whose `a:rPr@kern` its size reaches). */
  readonly kerning?: boolean;
}

export interface OfficeMeasuredGrapheme {
  readonly text: string;
  /** Advance in pt, as the face file states it at the size (unrounded). */
  readonly advance: number;
}

export type OfficeMeasureAnswer<T> =
  | { readonly status: "ok"; readonly value: T }
  | { readonly status: "pending" }
  | { readonly status: "unsupported"; readonly reason: string };

/**
 * Where a face's ink box sits, per em: its win ascent and descent, or its
 * typo ones where the file says to use them (`office-text-measurer.ts`).
 */
export interface OfficeFaceVerticals {
  readonly ascent: number;
  readonly descent: number;
}

export interface OfficeTextMeasurer {
  graphemes(text: string, font: OfficeTextFont): OfficeMeasureAnswer<readonly OfficeMeasuredGrapheme[]>;
  verticals(font: OfficeTextFont): OfficeMeasureAnswer<OfficeFaceVerticals>;
}

export interface OfficeTextRunInput {
  readonly text: string;
  readonly font: OfficeTextFont;
  /** `a:rPr@spc`, pt, after every glyph. */
  readonly letterSpacingPt?: number;
  /** All capitals: measured and drawn upper-cased. */
  readonly allCaps?: boolean;
}

export interface OfficeMarkerInput {
  /** The marker's characters (a bullet, "1.", "(iv)"). */
  readonly text: string;
  readonly font: OfficeTextFont;
  /** A picture marker's width instead of the text's. */
  readonly widthPt?: number;
  /** The bullet's own symbol-font character (`a:buFont` + `a:buChar`), which
   * `text` may stand in for: its width is the symbol glyph's. */
  readonly symbol?: { readonly family: string; readonly char: string };
  /** Where the marker stands on the first line, pt from the text area's left (marL + indent). */
  readonly leftPt: number;
  /** The marker's box: the words start past it, or past the marker when it is wider. */
  readonly boxPt: number;
  /** Between the marker's box and the words. */
  readonly gapPt: number;
  /** PowerPoint's hanging marker: the following lines go back to marL. Otherwise
   * every line starts where the first line's words do (the editor's own lists). */
  readonly hanging: boolean;
}

export type OfficeLineSpacing =
  | { readonly kind: "percent"; /** 1 = 100 % */ readonly value: number }
  | { readonly kind: "points"; readonly value: number };

export type OfficeTextAlign = "l" | "ctr" | "r" | "just" | "dist";

export interface OfficeTabStop {
  /** From the text area's left edge, pt. */
  readonly positionPt: number;
  readonly align: "l" | "ctr" | "r" | "dec";
}

export interface OfficeParagraphInput {
  readonly runs: readonly OfficeTextRunInput[];
  /** The face and size of the paragraph's end mark (an empty row's line). */
  readonly endFont: OfficeTextFont;
  readonly align: OfficeTextAlign;
  readonly marginLeftPt: number;
  readonly marginRightPt: number;
  readonly indentPt: number;
  readonly lineSpacing: OfficeLineSpacing;
  readonly spaceBeforePt: number;
  readonly spaceAfterPt: number;
  readonly marker?: OfficeMarkerInput;
  readonly tabStops?: readonly OfficeTabStop[];
  /** Default tab interval (`a:pPr@defTabSz`), pt. */
  readonly defaultTabPt?: number;
  readonly kinsoku?: KinsokuList | null;
  readonly latinLineBreak?: boolean;
}

export interface OfficeTextBodyInput {
  /** The text area's width (the shape's width less its insets), pt. */
  readonly widthPt: number;
  /** `a:bodyPr@wrap`: false keeps every line whole. */
  readonly wrap: boolean;
  readonly paragraphs: readonly OfficeParagraphInput[];
  /** `a:normAutofit@lnSpcReduction` as a fraction (0.2 = 20 %). */
  readonly lineSpacingReduction?: number;
  


  readonly keepLongWords?: boolean;
  /**
   * A table cell's lines (PowerPoint for Mac's P2-cell PDF: Arial, Calibri,
   * Times New Roman and Meiryo UI, 8-18 pt, 1-3 lines, top / middle / bottom):
   * the first baseline at the face's box centred in the 1.2 line (or, for a
   * face whose box is taller than the line, Calibri's, the line shared in the
   * box's proportion), on whole points like a text box's; a cell is centred
   * on its whole lines (`linesBottom`), the anchor's offset not rounded.
   * Arial 10 pt: 9 pt under the top (a text box: 10), 23 in the middle of a
   * 40 pt row, 37 at its foot.
   */
  readonly cellLines?: boolean;
}

/** A stretch of one run on one line. */
export interface OfficeLaidSegment {
  readonly run: number;
  /** UTF-16 offsets into the run's text. */
  readonly start: number;
  readonly end: number;
  /** Left edge from the text area's left, pt (spacing for justification included). */
  readonly x: number;
  /** Advance, pt, as laid (rounded advances, spacing, justification). */
  readonly width: number;
  /** The same graphemes at their unrounded advances (what a browser draws with no help). */
  readonly naturalWidth: number;
  /** Graphemes in the segment. */
  readonly graphemes: number;
  /** Spaces in the segment that take a justified line's extra. */
  readonly spaces: number;
  /** A tab's stretch: drawn as its gap. */
  readonly tab?: boolean;
  /** Each grapheme of the stretch and its left edge from the text area's left, pt. */
  readonly glyphs: readonly { readonly text: string; readonly x: number }[];
}

export interface OfficeLaidLine {
  readonly segments: readonly OfficeLaidSegment[];
  /** Baseline, pt from the top of the text area. */
  readonly baseline: number;
  /** The line's top and bottom, pt from the top of the text area. */
  readonly top: number;
  readonly bottom: number;
  /** Ink extent: first glyph's left to last non-space glyph's right. */
  readonly inkLeft: number;
  readonly inkRight: number;
  /** Extra per space / per character gap this line is drawn with (just, dist). */
  readonly wordSpacing: number;
  readonly letterSpacing: number;
  /** The largest type size on the line. */
  readonly sizePt: number;
}

export interface OfficeLaidMarker {
  readonly x: number;
  readonly width: number;
}

export interface OfficeLaidParagraph {
  readonly lines: readonly OfficeLaidLine[];
  readonly marker?: OfficeLaidMarker;
  /** Top of the paragraph's first line, bottom of its last (pt). */
  readonly top: number;
  readonly bottom: number;
}

export type OfficeTextBodyLayout =
  | {
      readonly status: "ok";
      readonly paragraphs: readonly OfficeLaidParagraph[];
      /** From the text area's top to under the last line's descent, pt. */
      readonly height: number;
    }
  | { readonly status: "pending" }
  | { readonly status: "unsupported"; readonly reason: string };

const EPSILON = 1e-6;

export function roundOfficeAdvance(advance: number): number {
  return Math.round(advance / OFFICE_ADVANCE_QUANTUM_PT) * OFFICE_ADVANCE_QUANTUM_PT;
}

interface Glyph {
  readonly text: string;
  readonly run: number;
  /** UTF-16 offset in the run's text. */
  readonly offset: number;
  readonly length: number;
  /** Laid advance (rounded, spacing included). */
  readonly advance: number;
  readonly natural: number;
  readonly kind: "ink" | "space" | "tab" | "break";
}

const SPACES = new Set([" ", "\u00a0", "\u3000"]);
/** A line break inside a paragraph (`a:br`), as the page keeps it. */
const BREAKS = new Set(["\n", "\u000b", "\u2028"]);

function upper(text: string): string {
  const shown = text.toUpperCase();
  return shown.length === text.length ? shown : text;
}

function glyphsOf(
  paragraph: OfficeParagraphInput,
  measurer: OfficeTextMeasurer,
): readonly Glyph[] | { status: "pending" } | { status: "unsupported"; reason: string } {
  const glyphs: Glyph[] = [];
  let pending = false;
  paragraph.runs.forEach((run, runIndex) => {
    if (!run.text) return;
    const shown = run.allCaps ? upper(run.text) : run.text;
    // Hard breaks and tabs are the layout's own; the measurer sees the rest.
    let offset = 0;
    const pieces = shown.split(/([\n\u000b\u2028\t])/);
    for (const piece of pieces) {
      if (!piece) continue;
      if (BREAKS.has(piece)) {
        glyphs.push({ text: piece, run: runIndex, offset, length: 1, advance: 0, natural: 0, kind: "break" });
        offset += 1;
        continue;
      }
      if (piece === "\t") {
        glyphs.push({ text: piece, run: runIndex, offset, length: 1, advance: 0, natural: 0, kind: "tab" });
        offset += 1;
        continue;
      }
      const answer = measurer.graphemes(piece, run.font);
      if (answer.status === "pending") {
        pending = true;
        offset += piece.length;
        continue;
      }
      if (answer.status === "unsupported") throw new UnsupportedText(answer.reason);
      const spacing = run.letterSpacingPt ?? 0;
      let local = 0;
      for (const grapheme of answer.value) {
        const kind = SPACES.has(grapheme.text) ? "space" : "ink";
        glyphs.push({
          text: grapheme.text,
          run: runIndex,
          offset: offset + local,
          length: grapheme.text.length,
          advance: roundOfficeAdvance(grapheme.advance) + spacing,
          natural: grapheme.advance + spacing,
          kind,
        });
        local += grapheme.text.length;
      }
      offset += piece.length;
    }
  });
  return pending ? { status: "pending" } : glyphs;
}

/**
 * What a body the engine leaves to the browser throws to its caller
 * (`layoutOfficeTextBody`): a plain value, not an Error, so no stack is
 * taken -- a page re-lays every body it draws on each keystroke, and the
 * bodies it cannot lay out answered by a stack trace each time.
 */
class UnsupportedText {
  constructor(readonly message: string) {}
}













function nextTabStop(x: number, paragraph: OfficeParagraphInput, hang?: number): { position: number; align: OfficeTabStop["align"] } {
  const custom = (paragraph.tabStops ?? []).find((stop) => stop.positionPt > x + EPSILON);
  const margin = hangingMarginTabTarget(x, custom?.positionPt ?? null, hang ?? null);
  if (margin !== null) return { position: margin, align: "l" };
  if (custom) return { position: custom.positionPt, align: custom.align };
  const step = paragraph.defaultTabPt && paragraph.defaultTabPt > 0 ? paragraph.defaultTabPt : 72;
  return { position: (Math.floor(x / step + EPSILON) + 1) * step, align: "l" };
}

interface LineRange {
  readonly start: number;
  /** Index after the line's last glyph (its hanging spaces and break included). */
  readonly end: number;
  /** Index after the last ink glyph. */
  readonly inkEnd: number;
  readonly hardEnd: boolean;
  /** Left edge of the line's words, pt from the text area's left. */
  readonly left: number;
  /** The width the line may fill, from `left`. */
  readonly available: number;
  /** Laid advances of glyphs [start, inkEnd) with tabs resolved. */
  readonly advances: readonly number[];
}

/** The line's tabs resolved from `left`, and the ink width up to each glyph. */
function resolveTabs(glyphs: readonly Glyph[], start: number, end: number, left: number, paragraph: OfficeParagraphInput, hang?: number): number[] {
  const out: number[] = [];
  let x = left;
  for (let index = start; index < end; index += 1) {
    const glyph = glyphs[index];
    let advance = glyph.advance;
    if (glyph.kind === "tab") advance = Math.max(0, nextTabStop(x, paragraph, hang).position - x);
    out.push(advance);
    x += advance;
  }
  return out;
}

function fillLines(
  glyphs: readonly Glyph[],
  paragraph: OfficeParagraphInput,
  widthPt: number,
  wrap: boolean,
  markerWidth: number,
  keepLongWords = false,
): LineRange[] {
  const texts = glyphs.map((glyph) => glyph.text);
  const opportunities = lineBreakOpportunities(texts, { kinsoku: paragraph.kinsoku ?? null, powerpoint: true });
  if (paragraph.latinLineBreak) {
    for (let index = 1; index < glyphs.length; index += 1) {
      if (glyphs[index].kind === "ink" && glyphs[index - 1].kind === "ink") opportunities[index] = true;
    }
  }
  // A tab ends a word: a line may break after it.
  for (let index = 1; index < glyphs.length; index += 1) {
    if (glyphs[index - 1].kind === "tab" && glyphs[index].kind !== "space") opportunities[index] = true;
  }
  const right = widthPt - paragraph.marginRightPt;
  const lines: LineRange[] = [];
  let start = 0;
  for (;;) {
    const first = lines.length === 0;
    let left = first ? paragraph.marginLeftPt + paragraph.indentPt : paragraph.marginLeftPt;
    const marker = paragraph.marker;
    if (marker && (first || !marker.hanging)) {
      // The marker hangs at marL + indent; the words start past its box (at
      // marL), or past a marker wider than its hang.
      left = marker.leftPt + Math.max(marker.boxPt, markerWidth) + marker.gapPt;
    }
    left = Math.max(0, left);
    // A hanging first line (no marker in its hang): its margin is a stop.
    const hang = first && !marker && paragraph.indentPt < -EPSILON && paragraph.marginLeftPt > left + EPSILON ? paragraph.marginLeftPt : undefined;
    const available = right - left;
    let x = left;
    let lastBreak = -1;
    let overflow = -1;
    let end = glyphs.length;
    let hardEnd = false;
    for (let index = start; index < glyphs.length; index += 1) {
      const glyph = glyphs[index];
      if (index > start && opportunities[index]) lastBreak = index;
      if (glyph.kind === "break") {
        end = index + 1;
        hardEnd = true;
        break;
      }
      let advance = glyph.advance;
      if (glyph.kind === "tab") advance = Math.max(0, nextTabStop(x, paragraph, hang).position - x);
      if (glyph.kind === "space") {
        x += advance;
        continue;
      }
      if (wrap && index > start && x + advance - left > available + EPSILON) {
        overflow = index;
        break;
      }
      x += advance;
    }
    if (overflow >= 0) {
      if (lastBreak > start) end = lastBreak;
      else if (keepLongWords) {
        // The word runs on whole: the line ends at the next place a line may.
        end = overflow;
        while (end < glyphs.length && !opportunities[end] && glyphs[end].kind !== "break") end += 1;
      } else end = overflow;
    }
    let inkEnd = end;
    while (inkEnd > start && (glyphs[inkEnd - 1].kind === "space" || glyphs[inkEnd - 1].kind === "break")) inkEnd -= 1;
    lines.push({ start, end, inkEnd, hardEnd: hardEnd && overflow < 0, left, available, advances: resolveTabs(glyphs, start, inkEnd, left, paragraph, hang) });
    if (end >= glyphs.length) {
      // A paragraph that ends on a break opens one more (empty) line.
      if (hardEnd && overflow < 0 && end === glyphs.length) {
        lines.push({ start: end, end, inkEnd: end, hardEnd: false, left: paragraph.marginLeftPt, available: right - paragraph.marginLeftPt, advances: [] });
      }
      return lines;
    }
    start = end;
  }
}

interface LineVerticals {
  /** From the line's top to its baseline. */
  readonly above: number;
  /** From the baseline to the next line's top. */
  readonly below: number;
  /** The face's own descent under the baseline (what the last line keeps). */
  readonly descent: number;
  readonly sizePt: number;
}

function lineVerticals(
  fonts: readonly OfficeTextFont[],
  spacing: OfficeLineSpacing,
  reduction: number,
  measurer: OfficeTextMeasurer,
  cell = false,
): LineVerticals | { status: "pending" } {
  let size = 0;
  let single = 0;
  let descent = 0;
  // The part of the single line under the baseline (its descent share).
  let singleBelow = 0;
  let pending = false;
  for (const font of fonts) {
    const answer = measurer.verticals(font);
    if (answer.status === "pending") {
      pending = true;
      continue;
    }
    if (answer.status === "unsupported") throw new UnsupportedText(answer.reason);
    const { ascent: a, descent: d } = answer.value;
    size = Math.max(size, font.sizePt);
    // Office's single line shared between the face's ascent and descent in
    // their proportion. Rounds 400-442's PowerPoint PDFs, the first line of a
    // top-anchored body at single spacing, rounded to a whole point (below):
    // within 0.25 pt of PowerPoint for 3317 of 3477 first lines in Arial,
    // Calibri, Times New Roman and Verdana, against 1138 with the face's ink
    // box centred in the line (Arial 16 pt: 16 pt under the top, not 15.14;
    // Calibri 40 pt: 37, not 37.66; Times 12 pt: 12, not 11.23).
    const share = cell && a + d <= OFFICE_SINGLE_LINE
      ? (OFFICE_SINGLE_LINE - (a + d)) / 2 + a
      : a + d > 0 ? (OFFICE_SINGLE_LINE * a) / (a + d) : a;
    single = Math.max(single, share * font.sizePt);
    singleBelow = Math.max(singleBelow, (OFFICE_SINGLE_LINE - share) * font.sizePt);
    descent = Math.max(descent, d * font.sizePt);
  }
  if (pending) return { status: "pending" };
  if (spacing.kind === "points") {
    const pitch = spacing.value;
    const above = OFFICE_EXACT_BASELINE_SHARE * pitch;
    return { above, below: pitch - above, descent, sizePt: size };
  }
  const multiple = Math.max(0.01, spacing.value - (reduction > 0 && reduction < 1 ? reduction : 0));
  const pitch = multiple * OFFICE_SINGLE_LINE * size;
  // Below single the line loses its height from above the baseline: the
  // single line's descent share stays under it. P1-pitch (PowerPoint for
  // Mac's PDF, five lines a box, whole-point baselines): Arial 90 % 8-32 pt
  // and Calibri 90 % 8-40 pt all within their feasible first baseline, where
  // the pitch less the face's own descent put Arial 32 pt a point low; and
  
  // stands 9.52 under its top, so its lines step 12, 12, 13, 12 as there).
  const above = cell
    ? multiple >= 1 - EPSILON
      ? Math.max(single, CELL_BASELINE_SHARE * pitch)
      : Math.max(pitch - descent, CELL_BASELINE_SHARE * pitch)
    // At single spacing the single-line ascent alone, however short: P3-faces
    // (PowerPoint for Mac) stands Meiryo (win 2171/901, 0.848 em) 10 pt's
    // first baseline 8 pt under the top and 11 pt's lines on 9, 23, 36, 49,
    // 62, where 3/4 of the pitch put every one of its 13 boxes a point low.
    : multiple <= 1 + EPSILON && multiple >= 1 - EPSILON
      ? single
      : multiple > 1
        ? Math.max(single, OFFICE_BASELINE_SHARE * pitch)
        : Math.max(pitch - singleBelow, OFFICE_BASELINE_SHARE * pitch);
  return { above, below: pitch - above, descent, sizePt: size };
}

/** Lay a text body out. */
export function layoutOfficeTextBody(input: OfficeTextBodyInput, measurer: OfficeTextMeasurer): OfficeTextBodyLayout {
  try {
    return layoutBody(input, measurer);
  } catch (error) {
    if (error instanceof UnsupportedText) return { status: "unsupported", reason: error.message };
    throw error;
  }
}

function layoutBody(input: OfficeTextBodyInput, measurer: OfficeTextMeasurer): OfficeTextBodyLayout {
  const reduction = input.lineSpacingReduction ?? 0;
  const laid: OfficeLaidParagraph[] = [];
  let pending = false;
  let cursor = 0;
  let lastDescent = 0;
  input.paragraphs.forEach((paragraph) => {
    const glyphs = glyphsOf(paragraph, measurer);
    let markerWidth = 0;
    const hasText = paragraph.runs.some((run) => run.text.length > 0);
    const marker = hasText ? paragraph.marker : undefined;
    if (marker) {
      const symbolEm = marker.symbol ? symbolAdvanceEm(marker.symbol.family, marker.symbol.char) : undefined;
      if (marker.widthPt !== undefined) markerWidth = marker.widthPt;
      // A bullet in an Office symbol font (Wingdings, Symbol): the page draws
      
      // Wingdings arrows, 13.5 pt over a 7.2 pt hang, move the words past).
      else if (symbolEm !== undefined) markerWidth = roundOfficeAdvance(symbolEm * marker.font.sizePt);
      else {
        const answer = measurer.graphemes(marker.text, marker.font);
        if (answer.status === "pending") pending = true;
        else if (answer.status === "unsupported") {
          // A hanging bullet in a face the page does not serve and no record
          // has: not 0, which would start the words at marL however wide the
          // bullet is, but an en dash of the text's own face, a bullet's
          // usual width, or else the hang itself.
          if (!marker.hanging) throw new UnsupportedText(answer.reason);
          const textFont = paragraph.runs.find((run) => run.text.length > 0)?.font ?? paragraph.endFont;
          const dash = measurer.graphemes("\u2013", { ...textFont, sizePt: marker.font.sizePt, kerning: false });
          if (dash.status === "pending") pending = true;
          markerWidth = dash.status === "ok" ? roundOfficeAdvance(dash.value[0]?.advance ?? 0) : marker.boxPt;
        } else markerWidth = answer.value.reduce((sum, grapheme) => sum + roundOfficeAdvance(grapheme.advance), 0);
      }
    }
    if (!Array.isArray(glyphs)) {
      if ((glyphs as { status: string }).status === "pending") {
        pending = true;
        return;
      }
    }
    if (pending) return;
    const list = glyphs as readonly Glyph[];
    const ranges = list.length ? fillLines(list, { ...paragraph, marker }, input.widthPt, input.wrap, markerWidth, input.keepLongWords) : [];
    const lineRanges = ranges.length ? ranges : [null];
    const lines: OfficeLaidLine[] = [];
    // Paragraph spacing adds up: the previous paragraph's after, this one's
    // before (the caller passes the first row's before as PowerPoint keeps it).
    cursor += paragraph.spaceBeforePt;
    let paragraphTop = Number.NaN;
    for (let lineIndex = 0; lineIndex < lineRanges.length; lineIndex += 1) {
      const range = lineRanges[lineIndex];
      const fonts: OfficeTextFont[] = [];
      if (range && range.inkEnd > range.start) {
        const runs = new Set<number>();
        for (let index = range.start; index < range.inkEnd; index += 1) runs.add(list[index].run);
        for (const run of runs) fonts.push(paragraph.runs[run].font);
      } else if (range && range.end > range.start) {
        fonts.push(paragraph.runs[list[range.start].run].font);
      }
      if (!fonts.length) fonts.push(paragraph.endFont);
      const verticals = lineVerticals(fonts, paragraph.lineSpacing, reduction, measurer, input.cellLines === true);
      if ("status" in verticals) {
        pending = true;
        return;
      }
      // The line's top is where the previous line's pitch ended.
      const top = cursor;
      // On a whole point from the body's top (PowerPoint for Mac's PDFs:
      // first baselines 19.01, 22.01, 25.99 pt under the inset; Arial 12 pt
      // lines 14 and 15 pt apart, never 14.4), the pitches adding up exactly.
      // A half point rounds down never: the pitches add up in binary (Arial
      // 150 % 14 pt's fourth line at 94.5 summed to 94.4999...).
      const baseline = Math.round(top + verticals.above + 1e-6);
      const bottom = top + verticals.above + verticals.below;
      if (Number.isNaN(paragraphTop)) paragraphTop = top;
      lines.push(placeLine(list, range, paragraph, input.widthPt, baseline, top, bottom, verticals.sizePt, lineIndex === lineRanges.length - 1));
      cursor = bottom;
      lastDescent = verticals.descent;
    }
    const lastBaseline = lines[lines.length - 1].baseline;
    laid.push({
      lines,
      marker: marker ? { x: marker.leftPt, width: markerWidth } : undefined,
      top: paragraphTop,
      bottom: lastBaseline + lastDescent,
    });
    cursor += paragraph.spaceAfterPt;
  });
  if (pending) return { status: "pending" };
  // The last paragraph's space after is not the body's: PowerPoint anchors
  
  
  // centred diagram labels with 9.24 pt, 4.7 pt).
  const last = laid[laid.length - 1];
  const height = last ? last.bottom : 0;
  return { status: "ok", paragraphs: laid, height };
}

function isEastAsian(text: string): boolean {
  return KINSOKU_EAST_ASIAN.test(text);
}

function placeLine(
  glyphs: readonly Glyph[],
  range: LineRange | null,
  paragraph: OfficeParagraphInput,
  widthPt: number,
  baseline: number,
  top: number,
  bottom: number,
  sizePt: number,
  lastOfParagraph: boolean,
): OfficeLaidLine {
  if (!range || range.end === range.start) {
    const left = paragraph.marginLeftPt;
    return { segments: [], baseline, top, bottom, inkLeft: left, inkRight: left, wordSpacing: 0, letterSpacing: 0, sizePt };
  }
  const inkWidth = range.advances.reduce((sum, advance) => sum + advance, 0);
  const extra = range.available - inkWidth;
  let offset = 0;
  let wordSpacing = 0;
  let letterSpacing = 0;
  // Spaces inside the ink, and the character gaps, for just / dist.
  let innerSpaces = 0;
  let firstInk = range.start;
  while (firstInk < range.inkEnd && glyphs[firstInk].kind === "space") firstInk += 1;
  for (let index = firstInk; index < range.inkEnd; index += 1) if (glyphs[index].kind === "space") innerSpaces += 1;
  const gaps = Math.max(0, range.inkEnd - range.start - 1);
  const justify = paragraph.align === "just" && !lastOfParagraph && !range.hardEnd && extra > 0;
  if (paragraph.align === "dist" && extra > 0) {
    if (gaps > 0) letterSpacing = extra / gaps;
    else offset = extra / 2;
  } else if (justify) {
    let eastAsian = false;
    for (let index = range.start; index < range.inkEnd && !eastAsian; index += 1) eastAsian = isEastAsian(glyphs[index].text);
    if (innerSpaces > 0 && !eastAsian) wordSpacing = extra / innerSpaces;
    else if (gaps > 0) letterSpacing = extra / gaps;
  } else if (paragraph.align === "ctr") {
    offset = extra / 2;
  } else if (paragraph.align === "r") {
    offset = extra;
  }
  const segments: OfficeLaidSegment[] = [];
  let x = range.left + offset;
  const inkLeft = x;
  let current: { run: number; start: number; end: number; x: number; width: number; natural: number; graphemes: number; spaces: number; tab: boolean; glyphs: { text: string; x: number }[] } | null = null;
  const flush = () => {
    if (current) {
      segments.push({
        run: current.run, start: current.start, end: current.end, x: current.x, width: current.width,
        naturalWidth: current.natural, graphemes: current.graphemes, spaces: current.spaces, ...(current.tab ? { tab: true } : {}),
        glyphs: current.glyphs,
      });
    }
    current = null;
  };
  let inkRight = x;
  for (let index = range.start; index < range.end; index += 1) {
    const glyph = glyphs[index];
    if (glyph.kind === "break") continue;
    const inside = index < range.inkEnd;
    const tab = glyph.kind === "tab";
    let advance = inside ? range.advances[index - range.start] : glyph.advance;
    if (inside) {
      if (glyph.kind === "space" && index >= firstInk) advance += wordSpacing;
      if (index < range.inkEnd - 1) advance += letterSpacing;
    }
    if (!current || current.run !== glyph.run || tab || current.tab) {
      flush();
      current = { run: glyph.run, start: glyph.offset, end: glyph.offset, x, width: 0, natural: 0, graphemes: 0, spaces: 0, tab, glyphs: [] };
    }
    const segment = current as NonNullable<typeof current>;
    segment.end = glyph.offset + glyph.length;
    segment.width += advance;
    segment.natural += tab ? advance : glyph.natural;
    segment.graphemes += 1;
    segment.glyphs.push({ text: glyph.text, x });
    if (glyph.kind === "space" && inside && index >= firstInk) segment.spaces += 1;
    x += advance;
    if (inside) inkRight = x;
  }
  flush();
  void widthPt;
  return { segments, baseline, top, bottom, inkLeft, inkRight, wordSpacing, letterSpacing, sizePt };
}
