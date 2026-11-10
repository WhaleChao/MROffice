/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfRangeGlyphRect } from "./range-glyph-rect";

export type PdfTextCaretPoint = Readonly<{ x: number; y: number }>;

/** Re-evaluate entry after the CURRENT text and fonts mount. Page fractions
 * survive scrolling during async preparation; no old source-text offset. */
export function pdfTextCaretOffsetAtPagePoint(
  draft: HTMLElement | null,
  page: HTMLElement | null,
  point: PdfTextCaretPoint | null | undefined,
  angle = 0,
): number | null {
  if (
    !draft?.isConnected ||
    !page?.isConnected ||
    !point ||
    ![point.x, point.y, angle].every(Number.isFinite) ||
    point.x < 0 ||
    point.x > 1 ||
    point.y < 0 ||
    point.y > 1
  )
    return null;
  const rect = page.getBoundingClientRect();
  if (!(rect.width > 0 && rect.height > 0)) return null;
  const direction = getComputedStyle(draft).direction === "rtl" ? -1 : 1;
  const radians = (angle * Math.PI) / 180;
  return textOffsetFromCharacterRects(
    draft,
    rect.left + point.x * rect.width,
    rect.top + point.y * rect.height,
    { x: direction * Math.cos(radians), y: direction * Math.sin(radians) },
  );
}

/** Shared with the FreeText overlay: measure all styled runs by grapheme,
 * keeping the caret out of surrogate pairs and combining sequences. */
export function textOffsetFromCharacterRects(
  node: HTMLElement,
  clientX: number,
  clientY: number,
  flowAxis?: Readonly<{ x: number; y: number }>,
): number | null {
  const doc = node.ownerDocument;
  let globalOffset = 0;
  const glyphs: Array<{
    start: number;
    end: number;
    centerX: number;
    centerY: number;
  }> = [];

  const visit = (current: Node): void => {
    if (current.nodeType === Node.TEXT_NODE) {
      const textNode = current as Text;
      const text = textNode.textContent ?? "";
      const boundaries = characterBoundaryIndexes(text);
      for (let index = 0; index < boundaries.length - 1; index += 1) {
        const start = boundaries[index] ?? 0;
        const end = boundaries[index + 1] ?? text.length;
        if (flowAxis && /[\r\n]/.test(text.slice(start, end))) continue;
        const range = doc.createRange();
        range.setStart(textNode, start);
        range.setEnd(textNode, end);
        if (typeof range.getBoundingClientRect !== "function") {
          range.detach();
          continue;
        }
        const rect = pdfRangeGlyphRect(range);
        range.detach();
        if (rect.width <= 0 && rect.height <= 0) continue;
        glyphs.push({
          start: globalOffset + start,
          end: globalOffset + end,
          centerX: rect.left + rect.width / 2,
          centerY: rect.top + rect.height / 2,
        });
      }
      globalOffset += text.length;
      return;
    }
    for (const child of Array.from(current.childNodes)) {
      visit(child);
    }
  };

  visit(node);
  if (glyphs.length === 0) return null;

  let nearestIndex = 0;
  let nearestDistance = Number.POSITIVE_INFINITY;
  let nearestLineDistance = Number.POSITIVE_INFINITY;
  for (let index = 0; index < glyphs.length; index += 1) {
    const glyph = glyphs[index];
    if (!glyph) continue;
    const dx = clientX - glyph.centerX;
    const dy = clientY - glyph.centerY;
    const distance = dx * dx + dy * dy;
    const lineDistance = flowAxis
      ? Math.abs(dx * flowAxis.y - dy * flowAxis.x)
      : 0;
    if (
      lineDistance < nearestLineDistance - 1 ||
      (Math.abs(lineDistance - nearestLineDistance) <= 1 &&
        distance < nearestDistance)
    ) {
      nearestIndex = index;
      nearestDistance = distance;
      nearestLineDistance = lineDistance;
    }
  }

  const nearest = glyphs[nearestIndex];
  if (!nearest) return null;
  const previous = glyphs[nearestIndex - 1];
  const next = glyphs[nearestIndex + 1];
  let axisX =
    (next?.centerX ?? nearest.centerX) - (previous?.centerX ?? nearest.centerX);
  let axisY =
    (next?.centerY ?? nearest.centerY) - (previous?.centerY ?? nearest.centerY);
  if (flowAxis) {
    axisX = flowAxis.x;
    axisY = flowAxis.y;
  }
  if (Math.hypot(axisX, axisY) < 0.01) {
    const rotationMatch = node.style.transform.match(
      /rotate\(([-+\d.eE]+)deg\)/,
    );
    const angle = Number(rotationMatch?.[1] ?? 0) * (Math.PI / 180);
    const vertical = node.style.writingMode.startsWith("vertical");
    const localAxisX = vertical ? 0 : 1;
    const localAxisY = vertical ? 1 : 0;
    axisX = localAxisX * Math.cos(angle) - localAxisY * Math.sin(angle);
    axisY = localAxisX * Math.sin(angle) + localAxisY * Math.cos(angle);
    if (!vertical && node.style.direction === "rtl") {
      axisX *= -1;
      axisY *= -1;
    }
  }
  const projection =
    (clientX - nearest.centerX) * axisX + (clientY - nearest.centerY) * axisY;
  return projection < 0 ? nearest.start : nearest.end;
}
type GraphemeSegmenter = {
  segment: (input: string) => Iterable<{ index: number }>;
};
type GraphemeSegmenterConstructor = new (
  locales?: string | string[],
  options?: { granularity?: "grapheme" },
) => GraphemeSegmenter;
export function characterBoundaryIndexes(text: string): number[] {
  const Segmenter = (
    Intl as unknown as {
      Segmenter?: GraphemeSegmenterConstructor;
    }
  ).Segmenter;
  if (Segmenter) {
    try {
      const boundaries = new Set<number>([0, text.length]);
      for (const segment of new Segmenter(undefined, {
        granularity: "grapheme",
      }).segment(text)) {
        boundaries.add(segment.index);
      }
      return Array.from(boundaries).sort((a, b) => a - b);
    } catch {
      // Fall through to code-point boundaries when Intl.Segmenter is absent or
      // unavailable in the test/browser runtime.
    }
  }

  const boundaries = [0];
  let index = 0;
  for (const character of text) {
    index += character.length;
    boundaries.push(index);
  }
  return boundaries;
}
