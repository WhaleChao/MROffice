/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Persisted per-line geometry for an editable PDF paragraph. */
export interface PdfTextLineLayout {
  /** UTF-16 offsets into the flat text, excluding the separating newline. */
  start: number;
  end: number;
  /** Geometry relative to the paragraph box. Values may exceed 1 after edits. */
  x: number;
  y: number;
  width: number;
  height: number;
}

const MAX_TEXT_LINES = 2_000;
const MAX_RELATIVE_GEOMETRY = 10;

function finiteBetween(
  value: unknown,
  min: number,
  max: number,
): number | null {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric >= min && numeric <= max
    ? numeric
    : null;
}

/** Validate line ranges and geometry against the exact flat text. */
export function normalizePdfTextLineLayout(
  value: unknown,
  text: string,
): PdfTextLineLayout[] | undefined {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    value.length > MAX_TEXT_LINES
  ) {
    return undefined;
  }
  const lines: PdfTextLineLayout[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object") return undefined;
    const record = entry as Record<string, unknown>;
    const start = Number(record.start);
    const end = Number(record.end);
    const x = finiteBetween(record.x, 0, MAX_RELATIVE_GEOMETRY);
    const y = finiteBetween(record.y, 0, MAX_RELATIVE_GEOMETRY);
    const width = finiteBetween(record.width, 0.000001, MAX_RELATIVE_GEOMETRY);
    const height = finiteBetween(
      record.height,
      0.000001,
      MAX_RELATIVE_GEOMETRY,
    );
    if (
      !Number.isInteger(start) ||
      !Number.isInteger(end) ||
      start < 0 ||
      end < start ||
      end > text.length ||
      x == null ||
      y == null ||
      width == null ||
      height == null
    ) {
      return undefined;
    }
    lines.push({ start, end, x, y, width, height });
  }
  let expectedStart = 0;
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    if (line.start !== expectedStart) return undefined;
    if (index < lines.length - 1) {
      if (text[line.end] !== "\n") return undefined;
      expectedStart = line.end + 1;
    } else if (line.end !== text.length) {
      return undefined;
    }
  }
  return lines;
}

/**
 * Keep source line geometry while text changes. Ordinary typing only changes
 * offsets. Added/removed explicit newlines reuse the nearest line geometry and
 * extend new lines by the previous leading, so no content becomes unpositioned.
 */
export function reconcilePdfTextLineLayout(
  value: unknown,
  previousText: string,
  nextText: string,
): PdfTextLineLayout[] | undefined {
  const previous = normalizePdfTextLineLayout(value, previousText);
  if (!previous) return undefined;
  if (previousText === nextText) return previous;

  const parts = nextText.split("\n");
  const fallbackHeight =
    previous.at(-1)?.height ?? 1 / Math.max(1, parts.length);
  let offset = 0;
  const next = parts.map((part, index) => {
    const source = previous[Math.min(index, previous.length - 1)];
    const extraIndex = Math.max(0, index - (previous.length - 1));
    const line: PdfTextLineLayout = source
      ? {
          ...source,
          y: source.y + extraIndex * source.height,
          start: offset,
          end: offset + part.length,
        }
      : {
          start: offset,
          end: offset + part.length,
          x: 0,
          y: index * fallbackHeight,
          width: 1,
          height: fallbackHeight,
        };
    offset = line.end + (index < parts.length - 1 ? 1 : 0);
    return line;
  });
  return normalizePdfTextLineLayout(next, nextText);
}
