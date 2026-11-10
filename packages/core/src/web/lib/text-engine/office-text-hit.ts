/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a point falls in a text body the engine laid out and drew (as SVG,
 * which the browser's own caret-from-point does not map back to the body's
 * characters): the text offset of the nearest place between two graphemes,
 * from the engine's own lines. A drawn body registers its layout on its
 * element; `officeTextOffsetAtPoint` answers for any element containing it.
 */
import type { OfficeTextBodyLayout } from "./office-text-body";

export interface OfficeTextHitMap {
  readonly layout: Extract<OfficeTextBodyLayout, { status: "ok" }>;
  /** Each paragraph's runs' lengths in the body's text (UTF-16), as laid out. */
  readonly runLengths: readonly (readonly number[])[];
}

const maps = new WeakMap<Element, OfficeTextHitMap>();

/** The element the engine's body is drawn in, with what it drew (or null to forget it). */
export function registerOfficeTextHitMap(element: Element, map: OfficeTextHitMap | null): void {
  if (map) maps.set(element, map);
  else maps.delete(element);
}

export const OFFICE_TEXT_HIT_ATTRIBUTE = "data-office-text-hit";

/**
 * The offset in the body's text (its paragraphs joined by "\n") nearest the
 * point, or null when `root` holds no body the engine drew, or the point is
 * outside it. `widthPt` is the body's laid width: the drawn element's width
 * gives the scale the slide is shown at.
 */
export function officeTextOffsetAtPoint(root: Element, clientX: number, clientY: number): number | null {
  const element = root.matches?.(`[${OFFICE_TEXT_HIT_ATTRIBUTE}]`) ? root : root.querySelector?.(`[${OFFICE_TEXT_HIT_ATTRIBUTE}]`);
  if (!element) return null;
  const map = maps.get(element);
  if (!map) return null;
  const rect = element.getBoundingClientRect();
  const heightPt = map.layout.height;
  if (!(rect.height > 0) || !(heightPt > 0)) return null;
  // pt per screen px: the body's drawn height against its laid height.
  const scale = heightPt / rect.height;
  const x = (clientX - rect.left) * scale;
  const y = (clientY - rect.top) * scale;
  let paragraphStart = 0;
  let best: { offset: number; distance: number } | null = null;
  map.layout.paragraphs.forEach((paragraph, paragraphIndex) => {
    const lengths = map.runLengths[paragraphIndex] ?? [];
    const runStarts: number[] = [];
    let total = 0;
    for (const length of lengths) {
      runStarts.push(total);
      total += length;
    }
    for (const line of paragraph.lines) {
      const dy = y < line.top ? line.top - y : y > line.bottom ? y - line.bottom : 0;
      // Each place between graphemes on the line: its x and its offset.
      const places: { x: number; offset: number }[] = [];
      for (const segment of line.segments) {
        let offset = paragraphStart + (runStarts[segment.run] ?? 0) + segment.start;
        segment.glyphs.forEach((glyph, index) => {
          places.push({ x: glyph.x, offset });
          offset += glyph.text.length;
          if (index === segment.glyphs.length - 1) {
            places.push({ x: segment.x + segment.width, offset });
          }
        });
      }
      if (!places.length) places.push({ x: line.inkLeft, offset: paragraphStart });
      for (const place of places) {
        const distance = dy * 1000 + Math.abs(place.x - x);
        if (!best || distance < best.distance) best = { offset: place.offset, distance };
      }
    }
    paragraphStart += total + 1;
  });
  return (best as { offset: number } | null)?.offset ?? null;
}

/** A caret's place in a laid body, pt from the text area's top-left. */
export interface OfficeTextCaret {
  readonly x: number;
  readonly top: number;
  readonly bottom: number;
  readonly baseline: number;
  /** The line's largest type size, pt. */
  readonly sizePt: number;
  readonly paragraph: number;
  readonly line: number;
}

interface PlacedLine {
  readonly paragraph: number;
  readonly line: number;
  /** Body offsets the line's segments cover. */
  readonly start: number;
  readonly end: number;
  readonly top: number;
  readonly bottom: number;
  readonly baseline: number;
  readonly sizePt: number;
  readonly inkLeft: number;
  /** Each place between graphemes: its body offset and its x. */
  readonly places: readonly { readonly offset: number; readonly x: number }[];
}

const placedLines = new WeakMap<OfficeTextHitMap, readonly PlacedLine[]>();

/** Every line of the body with the body offsets its graphemes stand at (paragraphs joined by "\n"). */
function placeLines(map: OfficeTextHitMap): readonly PlacedLine[] {
  const known = placedLines.get(map);
  if (known) return known;
  const out: PlacedLine[] = [];
  let paragraphStart = 0;
  map.layout.paragraphs.forEach((paragraph, paragraphIndex) => {
    const lengths = map.runLengths[paragraphIndex] ?? [];
    const runStarts: number[] = [];
    let total = 0;
    for (const length of lengths) {
      runStarts.push(total);
      total += length;
    }
    let previousEnd = paragraphStart;
    paragraph.lines.forEach((line, lineIndex) => {
      const places: { offset: number; x: number }[] = [];
      for (const segment of line.segments) {
        let offset = paragraphStart + (runStarts[segment.run] ?? 0) + segment.start;
        segment.glyphs.forEach((glyph) => {
          places.push({ offset, x: glyph.x });
          offset += glyph.text.length;
        });
        places.push({ offset, x: segment.x + segment.width });
      }
      // A line with no graphemes (an empty paragraph, the line a break opens)
      // holds its caret where the line's words would start: past the break
      // that ended the line before it.
      const start = places.length ? places[0].offset : lineIndex === 0 ? paragraphStart : previousEnd + 1;
      const end = places.length ? places[places.length - 1].offset : start;
      if (!places.length) places.push({ offset: start, x: line.inkLeft });
      out.push({ paragraph: paragraphIndex, line: lineIndex, start, end, top: line.top, bottom: line.bottom, baseline: line.baseline, sizePt: line.sizePt, inkLeft: line.inkLeft, places });
      previousEnd = end;
    });
    paragraphStart += total + 1;
  });
  placedLines.set(map, out);
  return out;
}

function xAt(line: PlacedLine, offset: number): number {
  let x = line.places[0].x;
  for (const place of line.places) {
    if (place.offset > offset) break;
    x = place.x;
  }
  return x;
}

/**
 * Where a caret at `offset` (the body's text, its paragraphs joined by "\n")
 * stands in the engine's lines: on the line that holds the character after
 * it, as a browser places a caret at a soft line end (at the start of the
 * next line); at the end of a paragraph's last line, or of a line a break
 * ends, on that line. Null when the body has no lines.
 */
export function officeTextCaretAtOffset(map: OfficeTextHitMap, offset: number): OfficeTextCaret | null {
  const lines = placeLines(map);
  if (!lines.length) return null;
  let chosen: PlacedLine | null = null;
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    if (offset < line.start) continue;
    const next = lines[index + 1];
    const lastOfParagraph = !next || next.paragraph !== line.paragraph;
    if (offset < line.end || (offset === line.end && (lastOfParagraph || next.start > offset))) {
      chosen = line;
      break;
    }
    if (lastOfParagraph && offset <= line.end) {
      chosen = line;
      break;
    }
  }
  chosen ??= offset <= lines[0].start ? lines[0] : lines[lines.length - 1];
  const x = xAt(chosen, Math.max(chosen.start, Math.min(chosen.end, offset)));
  return { x, top: chosen.top, bottom: chosen.bottom, baseline: chosen.baseline, sizePt: chosen.sizePt, paragraph: chosen.paragraph, line: chosen.line };
}

/** The rectangles a selection of the body's [start, end) covers on the engine's lines, pt. */
export function officeTextSelectionRects(
  map: OfficeTextHitMap,
  start: number,
  end: number,
): { readonly x: number; readonly y: number; readonly width: number; readonly height: number }[] {
  const from = Math.min(start, end);
  const to = Math.max(start, end);
  if (to <= from) return [];
  const lines = placeLines(map);
  const rects: { x: number; y: number; width: number; height: number }[] = [];
  lines.forEach((line, index) => {
    const next = lines[index + 1];
    const lastOfParagraph = !next || next.paragraph !== line.paragraph;
    // The paragraph's end ("\n" after its last line) is selected too.
    const reach = lastOfParagraph ? line.end + 1 : next.start;
    if (to <= line.start || from >= reach) return;
    if (from >= line.end && !(lastOfParagraph && from === line.end && to > line.end)) return;
    const x0 = xAt(line, Math.max(line.start, from));
    let x1 = xAt(line, Math.min(line.end, to));
    // A selected paragraph end shows as a sliver past the line's words, as Office shows it.
    if (lastOfParagraph && to > line.end) x1 += Math.max(2, (line.bottom - line.top) * 0.25);
    if (x1 > x0) rects.push({ x: x0, y: line.top, width: x1 - x0, height: line.bottom - line.top });
  });
  return rects;
}
