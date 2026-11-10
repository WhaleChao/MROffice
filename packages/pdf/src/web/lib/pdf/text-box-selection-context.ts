/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { NormalizedRect } from "./normalizeHighlightRects";
import type { TextBoxViewOverlay } from "./text-box-view-overlay";

const MIN_OVERLAY_COVERAGE = 0.28;

function normalizeWhitespace(text: string): string {
  return text
    .split("\n")
    .map((line) => line.replace(/[ \t]{2,}/g, " ").trim())
    .filter(Boolean)
    .join("\n")
    .trim();
}

function rectArea(rect: NormalizedRect): number {
  return Math.max(0, rect.width) * Math.max(0, rect.height);
}

function intersectionArea(a: NormalizedRect, b: NormalizedRect): number {
  const x0 = Math.max(a.x, b.x);
  const y0 = Math.max(a.y, b.y);
  const x1 = Math.min(a.x + a.width, b.x + b.width);
  const y1 = Math.min(a.y + a.height, b.y + b.height);
  return Math.max(0, x1 - x0) * Math.max(0, y1 - y0);
}

function overlayRect(overlay: TextBoxViewOverlay): NormalizedRect {
  return {
    x: overlay.x * 100,
    y: overlay.y * 100,
    width: overlay.width * 100,
    height: overlay.height * 100,
  };
}

function overlayMaskRect(overlay: TextBoxViewOverlay): NormalizedRect | null {
  if (!overlay.mask) return null;
  return {
    x: overlay.mask.x * 100,
    y: overlay.mask.y * 100,
    width: overlay.mask.width * 100,
    height: overlay.mask.height * 100,
  };
}

function rectsOverlapSelection(
  selectionRects: readonly NormalizedRect[],
  candidate: NormalizedRect,
): boolean {
  const area = rectArea(candidate);
  if (area <= 0) return false;
  const covered = selectionRects.reduce(
    (sum, rect) => sum + intersectionArea(rect, candidate),
    0,
  );
  if (covered / area >= MIN_OVERLAY_COVERAGE) return true;

  const center = {
    x: candidate.x + candidate.width / 2,
    y: candidate.y + candidate.height / 2,
  };
  return selectionRects.some(
    (rect) =>
      center.x >= rect.x &&
      center.x <= rect.x + rect.width &&
      center.y >= rect.y &&
      center.y <= rect.y + rect.height,
  );
}

function overlayOverlapsSelection(
  selectionRects: readonly NormalizedRect[],
  overlay: TextBoxViewOverlay,
): boolean {
  return [overlayMaskRect(overlay), overlayRect(overlay)]
    .filter((rect): rect is NormalizedRect => Boolean(rect))
    .some((rect) => rectsOverlapSelection(selectionRects, rect));
}

/** A selection reaching this close to a row's edges selects the whole row. */
const FULL_ROW_TOLERANCE = 0.03;

interface OverlayRow {
  rect: NormalizedRect;
  start: number;
  end: number;
}

function overlayRows(
  overlay: TextBoxViewOverlay,
  rect: NormalizedRect,
  useLineLayout: boolean,
): OverlayRow[] {
  const length = overlay.text.length;
  const layout = useLineLayout ? overlay.lineLayout : undefined;
  if (
    Array.isArray(layout) &&
    layout.length > 0 &&
    layout.every(
      (line) =>
        Number.isInteger(line.start) &&
        Number.isInteger(line.end) &&
        line.start >= 0 &&
        line.end >= line.start &&
        line.end <= length &&
        [line.x, line.y, line.width, line.height].every(Number.isFinite),
    )
  ) {
    return layout.map((line) => ({
      rect: {
        x: rect.x + line.x * rect.width,
        y: rect.y + line.y * rect.height,
        width: line.width * rect.width,
        height: line.height * rect.height,
      },
      start: line.start,
      end: line.end,
    }));
  }
  return [{ rect, start: 0, end: length }];
}

const WORD_CHARACTER = /[\p{L}\p{N}\p{M}_'’-]/u;
const SINGLE_CHARACTER_WORD = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;

/** Offsets where a word starts or ends (every CJK character is a word). */
function isWordBoundary(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) return true;
  const before = text[offset - 1];
  const after = text[offset];
  if (SINGLE_CHARACTER_WORD.test(before) || SINGLE_CHARACTER_WORD.test(after)) {
    return true;
  }
  return WORD_CHARACTER.test(before) !== WORD_CHARACTER.test(after);
}

function nearestWordBoundary(
  text: string,
  offset: number,
  from: number,
  to: number,
): number {
  const target = Math.max(from, Math.min(to, Math.round(offset)));
  for (let distance = 0; distance <= to - from; distance += 1) {
    if (target - distance >= from && isWordBoundary(text, target - distance)) {
      return target - distance;
    }
    if (target + distance <= to && isWordBoundary(text, target + distance)) {
      return target + distance;
    }
  }
  return target;
}










function overlaySelectedRange(
  selectionRects: readonly NormalizedRect[],
  overlay: TextBoxViewOverlay,
  rawText: string,
): { start: number; end: number } | null {
  const text = overlay.text;
  if (!text) return null;
  const mask = overlayMaskRect(overlay);
  const box = overlayRect(overlay);
  const onMask = Boolean(mask && rectsOverlapSelection(selectionRects, mask));
  const rect = onMask && mask ? mask : box;
  const sameGeometry =
    !mask ||
    (Math.abs(mask.x - box.x) < 0.01 &&
      Math.abs(mask.y - box.y) < 0.01 &&
      Math.abs(mask.width - box.width) < 0.01 &&
      Math.abs(mask.height - box.height) < 0.01);
  const rows = overlayRows(overlay, rect, !onMask || sameGeometry);

  let start = Infinity;
  let end = -Infinity;
  let whole = true;
  let hitRows = 0;
  for (const row of rows) {
    if (row.rect.width <= 0 || row.rect.height <= 0) continue;
    let left = Infinity;
    let right = -Infinity;
    for (const selected of selectionRects) {
      const top = Math.max(selected.y, row.rect.y);
      const bottom = Math.min(
        selected.y + selected.height,
        row.rect.y + row.rect.height,
      );
      if (bottom - top < Math.min(selected.height, row.rect.height) / 2) {
        continue;
      }
      left = Math.min(left, Math.max(selected.x, row.rect.x));
      right = Math.max(
        right,
        Math.min(selected.x + selected.width, row.rect.x + row.rect.width),
      );
    }
    if (!(right > left)) {
      whole = false;
      continue;
    }
    hitRows += 1;
    const from = (left - row.rect.x) / row.rect.width;
    const to = (right - row.rect.x) / row.rect.width;
    if (from > FULL_ROW_TOLERANCE || to < 1 - FULL_ROW_TOLERANCE) whole = false;
    const span = row.end - row.start;
    start = Math.min(start, row.start + from * span);
    end = Math.max(end, row.start + to * span);
  }
  if (hitRows === 0 || whole) return null;

  // A phrase the edit left alone is quoted exactly as selected.
  const raw = normalizeWhitespace(rawText);
  if (raw && !raw.includes("\n")) {
    let best = -1;
    for (
      let at = text.indexOf(raw);
      at >= 0;
      at = text.indexOf(raw, at + 1)
    ) {
      if (best < 0 || Math.abs(at - start) < Math.abs(best - start)) best = at;
    }
    if (best >= 0) return { start: best, end: best + raw.length };
  }

  let from = nearestWordBoundary(text, start, 0, text.length);
  let to = nearestWordBoundary(text, end, 0, text.length);
  if (to <= from) {
    // Less than a word under the selection: the word at its middle.
    const middle = Math.min(text.length - 1, Math.max(0, Math.floor((start + end) / 2)));
    from = middle;
    to = middle + 1;
    while (from > 0 && !isWordBoundary(text, from)) from -= 1;
    while (to < text.length && !isWordBoundary(text, to)) to += 1;
  }
  return { start: from, end: to };
}

function overlaySelectionLine(
  overlay: TextBoxViewOverlay,
  selectionRects: readonly NormalizedRect[],
  rawText: string,
): string | null {
  const latex = normalizeWhitespace(overlay.latex ?? "");
  if (overlay.kind === "formula" && latex) {
    const text = normalizeWhitespace(overlay.text);
    return text && text !== latex ? `${text}\nLaTeX: ${latex}` : latex;
  }
  const range = overlaySelectedRange(selectionRects, overlay, rawText);
  const text = normalizeWhitespace(
    range ? overlay.text.slice(range.start, range.end) : overlay.text,
  );
  return text || null;
}

export function selectionTextWithTextBoxOverrides({
  rawText,
  pageNumber,
  selectionRects,
  overlays,
}: {
  rawText: string;
  pageNumber: number | null;
  selectionRects: readonly NormalizedRect[];
  overlays: ReadonlyArray<TextBoxViewOverlay> | null | undefined;
}): string {
  const fallback = normalizeWhitespace(rawText);
  if (!pageNumber || selectionRects.length === 0 || !Array.isArray(overlays)) {
    return fallback;
  }

  const overrideLines = overlays
    .filter((overlay) => overlay.page === pageNumber)
    .filter((overlay) => overlayOverlapsSelection(selectionRects, overlay))
    .sort((a, b) => a.y - b.y || a.x - b.x)
    .map((overlay) => overlaySelectionLine(overlay, selectionRects, rawText))
    .filter((line): line is string => Boolean(line));

  if (overrideLines.length === 0) return fallback;
  return normalizeWhitespace(overrideLines.join("\n"));
}
