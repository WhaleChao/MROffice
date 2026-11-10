/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { foldTextForSearch, type FoldTextOptions, type PdfPageSearchText } from "./find-normalize";

/** A text box as the reader draws it: page fractions of the displayed page. */
export interface FindTextBoxOverlay {
  readonly id: string;
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** How the live view sets the box's text (text-box-view-overlay), so a
   * match can be placed at its letters; absent, the whole box is lit. */
  readonly fontSize?: number;
  readonly fontFamily?: string;
  readonly fontWeight?: string;
  readonly fontStyle?: string;
  readonly textAlign?: string;
  readonly angle?: number;
  readonly typography?: unknown;
  readonly lineLayout?: ReadonlyArray<
    Readonly<{ start: number; end: number; x: number; y: number; width: number; height: number }>
  >;
}

/** The width of `text` set as `box` draws it, as a fraction of the displayed
 * page's width; null when it cannot be measured. */
export type FindTextBoxMeasure = (text: string, box: FindTextBoxOverlay) => number | null;

/** A stored text-box edit; only what tells which reader items it replaced. */
export interface FindTextBoxEdit {
  readonly id?: unknown;
  readonly sourceText?: unknown;
}

export interface FindTextBoxSegment {
  /** Folded offsets of the box's text in the page's searchable text. */
  readonly start: number;
  readonly end: number;
  readonly box: FindTextBoxOverlay;
}

export interface FindPageWithTextBoxes extends PdfPageSearchText {
  /** Where the boxes' text begins (the page's own text ends before it). */
  readonly boxBase: number;
  readonly segments: readonly FindTextBoxSegment[];
}

const NATIVE_TEXT_BOX_ID = /^native-text:(\d+):(\d+):(\d+)$/;

const collapse = (value: string) => value.replace(/\s+/gu, "");

/**
 * The reader items (indexes into the page's text items) the page's text
 * boxes replaced: a box bound to its source line names them in its id, and
 * an item is taken only while it still reads as part of the line the box
 * was proved on (an index that no longer names that line is left alone).
 */
export function textBoxReplacedItemIndexes(
  edits: readonly FindTextBoxEdit[] | null | undefined,
  pageNumber: number,
  itemText: (index: number) => string | null,
): number[] {
  const indexes = new Set<number>();
  for (const edit of edits ?? []) {
    const source = edit?.sourceText as
      | { boxId?: unknown; originalProof?: { evidence?: { originalText?: unknown } } }
      | null
      | undefined;
    if (!source || typeof source !== "object") continue;
    const id = typeof source.boxId === "string" ? source.boxId : typeof edit.id === "string" ? edit.id : "";
    const match = NATIVE_TEXT_BOX_ID.exec(id);
    if (!match || Number(match[1]) !== pageNumber) continue;
    const first = Number(match[2]);
    const last = Number(match[3]);
    if (!(last >= first) || last - first > 500) continue;
    const original = source.originalProof?.evidence?.originalText;
    const line = typeof original === "string" ? collapse(original) : null;
    for (let index = first; index <= last; index += 1) {
      const text = itemText(index);
      if (text === null) continue;
      const own = collapse(text);
      if (!own) continue;
      if (line !== null && !line.includes(own)) continue;
      indexes.add(index);
    }
  }
  return [...indexes].sort((a, b) => a - b);
}

/** The raw ranges of those items in the page text (each up to the next). */
export function replacedItemRawRanges(
  entry: Pick<PdfPageSearchText, "rawText" | "itemStarts">,
  itemIndexes: readonly number[],
): Array<[number, number]> {
  const ranges: Array<[number, number]> = [];
  for (const index of itemIndexes) {
    const start = entry.itemStarts[index];
    if (start === undefined) continue;
    const end = entry.itemStarts[index + 1] ?? entry.rawText.length;
    if (end <= start) continue;
    const last = ranges[ranges.length - 1];
    if (last && last[1] >= start) last[1] = Math.max(last[1], end);
    else ranges.push([start, end]);
  }
  return ranges;
}

/**
 * The page's searchable text (`visibleText`: the folded text with whatever
 * is not on the page blanked in place) followed by each box's text, one line
 * each, with the raw text and offset maps extended to match.
 */
export function pageSearchTextWithTextBoxes(
  entry: PdfPageSearchText,
  visibleText: string,
  boxes: readonly FindTextBoxOverlay[],
  options: FoldTextOptions = {},
): FindPageWithTextBoxes {
  let text = visibleText;
  let rawText = entry.rawText;
  const normStarts = [...entry.normStarts];
  const normEnds = [...entry.normEnds];
  const segments: FindTextBoxSegment[] = [];
  const boxBase = text.length;
  for (const box of boxes) {
    if (!box.text.trim()) continue;
    // A line break before the box, in both texts: no match runs into it.
    normStarts.push(rawText.length);
    normEnds.push(rawText.length + 1);
    text += "\n";
    rawText += "\n";
    const folded = foldTextForSearch(box.text, options);
    const start = text.length;
    for (let index = 0; index < folded.text.length; index += 1) {
      normStarts.push(rawText.length + (folded.normStarts[index] ?? 0));
      normEnds.push(rawText.length + (folded.normEnds[index] ?? 0));
    }
    text += folded.text;
    rawText += box.text;
    segments.push(Object.freeze({ start, end: text.length, box }));
  }
  return Object.freeze({
    text,
    rawText,
    itemStarts: entry.itemStarts,
    normStarts,
    normEnds,
    boxBase,
    segments: Object.freeze(segments),
  });
}

/** Whether a page's searched text carries text boxes after the page's own. */
export function isFindPageWithTextBoxes(
  page: PdfPageSearchText,
): page is FindPageWithTextBoxes {
  return "boxBase" in page && Array.isArray((page as Partial<FindPageWithTextBoxes>).segments);
}

type TextBoxMatchPage = Pick<FindPageWithTextBoxes, "boxBase" | "segments"> &
  Partial<Pick<FindPageWithTextBoxes, "normStarts" | "normEnds">>;

/**
 * The rects (page fractions) of the match's letters in the box, measured as
 * the view sets them: a short word lit the whole two-line box (user audit
 * C-0039). Null when the box is not set the way this can measure (turned,
 * scaled, or lines the browser breaks itself), and the whole box is lit.
 */
function textBoxWordRects(
  page: TextBoxMatchPage,
  segment: FindTextBoxSegment,
  match: Readonly<{ start: number; end: number }>,
  measure: FindTextBoxMeasure,
): Array<{ x: number; y: number; width: number; height: number }> | null {
  const { box } = segment;
  const { normStarts, normEnds } = page;
  if (!normStarts || !normEnds || box.angle || box.typography) return null;
  const base = normStarts[segment.start];
  const first = normStarts[Math.max(match.start, segment.start)];
  const last = normEnds[Math.min(match.end, segment.end) - 1];
  if (base === undefined || first === undefined || last === undefined) return null;
  const rawStart = first - base;
  const rawEnd = last - base;
  if (!(rawEnd > rawStart)) return null;
  const lines = box.lineLayout ??
    (box.text.includes("\n") ? null : [{ start: 0, end: box.text.length, x: 0, y: 0, width: 1, height: 1 }]);
  if (!lines) return null;
  const rects: Array<{ x: number; y: number; width: number; height: number }> = [];
  for (const line of lines) {
    if (line.end <= rawStart || line.start >= rawEnd) continue;
    const lineText = box.text.slice(line.start, line.end);
    const from = Math.max(rawStart, line.start) - line.start;
    const to = Math.min(rawEnd, line.end) - line.start;
    const lineWidth = measure(lineText, box);
    const before = measure(lineText.slice(0, from), box);
    const width = measure(lineText.slice(from, to), box);
    if (lineWidth === null || before === null || width === null) return null;
    const room = line.width * box.width;
    // Wider than its line: the browser breaks it where this cannot tell.
    if (lineWidth > room * 1.02) return null;
    const lead = box.textAlign === "center" ? (room - lineWidth) / 2 : box.textAlign === "right" ? room - lineWidth : 0;
    rects.push({
      x: box.x + line.x * box.width + lead + before,
      y: box.y + line.y * box.height,
      width,
      height: line.height * box.height,
    });
  }
  return rects.length > 0 ? rects : null;
}

/** Highlight rects (percent of the page) of a match in a box's text: at its
 * letters when `measure` sets the box's text as the view does, else the box. */
export function textBoxMatchRects(
  page: TextBoxMatchPage,
  match: Readonly<{ start: number; end: number }>,
  matchIndex: number,
  measure?: FindTextBoxMeasure,
): Array<{ matchIndex: number; x: number; y: number; width: number; height: number; textBox: true }> {
  if (match.start < page.boxBase) return [];
  return page.segments
    .filter((segment) => segment.start < match.end && match.start < segment.end)
    .flatMap((segment) => {
      const rects = (measure ? textBoxWordRects(page, segment, match, measure) : null) ??
        [{ x: segment.box.x, y: segment.box.y, width: segment.box.width, height: segment.box.height }];
      return rects.map((rect) => ({
        matchIndex,
        x: rect.x * 100,
        y: rect.y * 100,
        width: rect.width * 100,
        height: rect.height * 100,
        textBox: true as const,
      }));
    });
}
