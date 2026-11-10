/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Formats of part of a PDF text run (WPS: two characters of a line made
 * bold change only those two). The editor keeps them as segments of its
 * draft -- ranges of the run's text, each with the fields it sets apart from
 * the run -- and the planner writes the run as its parts
 * (list-text-fidelity.ts, planPropertySegmentParts). This module is the
 * editor's bookkeeping: put a format on a range, follow typing, read the
 * format a range has.
 */

import type {
  PdfNativeTextPropertySegmentDraft,
} from "./engine/native/list-text-fidelity";

export type PdfNativeTextSegmentFields = Omit<PdfNativeTextPropertySegmentDraft, "start" | "end">;

const FIELDS = ["fontResourceName", "fontSize", "color", "syntheticBold", "slanted", "textRise", "characterSpacing"] as const;

function sameValue(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function fieldsOf(segment: PdfNativeTextPropertySegmentDraft): PdfNativeTextSegmentFields {
  const fields: Record<string, unknown> = {};
  for (const key of FIELDS) {
    if (segment[key] !== undefined) fields[key] = segment[key];
  }
  return fields as PdfNativeTextSegmentFields;
}

/**
 * The segments in order, each holding only what differs from the run's own
 * values (`base`: what a field left out means), empty ones dropped and
 * neighbours that set the same fields joined.
 */
export function normalizePdfNativeTextSegments(
  segments: readonly PdfNativeTextPropertySegmentDraft[],
  base: PdfNativeTextSegmentFields,
): readonly PdfNativeTextPropertySegmentDraft[] {
  const out: PdfNativeTextPropertySegmentDraft[] = [];
  for (const segment of [...segments].sort((left, right) => left.start - right.start)) {
    if (segment.end <= segment.start) continue;
    const fields: Record<string, unknown> = {};
    for (const key of FIELDS) {
      const value = segment[key];
      if (value === undefined) continue;
      if (base[key] !== undefined ? sameValue(value, base[key]) : (key === "syntheticBold" || key === "slanted") && value === false) continue;
      fields[key] = value;
    }
    if (Object.keys(fields).length === 0) continue;
    const last = out[out.length - 1];
    if (last && last.end === segment.start && sameValue(fieldsOf(last), fields)) {
      out[out.length - 1] = Object.freeze({ ...last, end: segment.end });
      continue;
    }
    out.push(Object.freeze({ start: segment.start, end: segment.end, ...fields }));
  }
  return Object.freeze(out);
}

/** The segments with `fields` put on `[start, end)`, over whatever the
 * range already had (a field set again replaces the old value). */
export function applyPdfNativeTextSegmentFormat(
  segments: readonly PdfNativeTextPropertySegmentDraft[],
  range: Readonly<{ start: number; end: number }>,
  fields: PdfNativeTextSegmentFields,
  base: PdfNativeTextSegmentFields,
): readonly PdfNativeTextPropertySegmentDraft[] {
  const pieces: PdfNativeTextPropertySegmentDraft[] = [];
  let covered = range.start;
  for (const segment of [...segments].sort((left, right) => left.start - right.start)) {
    if (segment.end <= range.start || segment.start >= range.end) {
      pieces.push(segment);
      continue;
    }
    // The part of the segment outside the range keeps what it had.
    if (segment.start < range.start) pieces.push({ ...segment, end: range.start });
    if (segment.end > range.end) pieces.push({ ...segment, start: range.end });
    const start = Math.max(segment.start, range.start);
    const end = Math.min(segment.end, range.end);
    if (start > covered) pieces.push({ start: covered, end: start, ...fields });
    pieces.push({ ...segment, start, end, ...fields });
    covered = end;
  }
  if (covered < range.end) pieces.push({ start: covered, end: range.end, ...fields });
  return normalizePdfNativeTextSegments(pieces, base);
}

/**
 * The segments after the text was edited from `before` to `after` (one
 * contiguous change, found by the common prefix and suffix): text typed in
 * a segment, or right after its end, takes its format (Word's rule: new
 * text takes the format before it); text deleted takes its segment part.
 */
export function pdfNativeTextSegmentsAfterEdit(
  segments: readonly PdfNativeTextPropertySegmentDraft[],
  before: string,
  after: string,
  base: PdfNativeTextSegmentFields,
): readonly PdfNativeTextPropertySegmentDraft[] {
  if (before === after || segments.length === 0) return segments;
  let prefix = 0;
  while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix]) prefix += 1;
  let suffix = 0;
  while (
    suffix < before.length - prefix && suffix < after.length - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  ) suffix += 1;
  const removedEnd = before.length - suffix;
  const inserted = after.length - suffix - prefix;
  const delta = after.length - before.length;
  const moved = segments.flatMap((segment) => {
    // Past the change: moves with the text after it. Text typed at the very
    // start takes the first character's format, there being none before it.
    if (segment.start >= removedEnd && !(segment.start === 0 && prefix === 0 && removedEnd === 0)) {
      return [{ ...segment, start: segment.start + delta, end: segment.end + delta }];
    }
    // Before the change, and not reaching it.
    if (segment.end < prefix || (segment.end === prefix && inserted === 0)) return [segment];
    // Over the change: what it lost is gone, what was typed in it (or right
    // after its end) is its own.
    const start = segment.start <= prefix ? segment.start : prefix + inserted;
    const keptAfter = Math.max(0, segment.end - removedEnd);
    const end = segment.end <= removedEnd
      ? (segment.end >= prefix ? prefix + inserted : segment.end)
      : prefix + inserted + keptAfter;
    return end > start ? [{ ...segment, start, end }] : [];
  });
  return normalizePdfNativeTextSegments(moved, base);
}

/** What a range of the run's text has for each field: the one value every
 * character of it has (the run's own where no segment sets it), or
 * "mixed". */
export function pdfNativeTextSegmentFieldsAt(
  segments: readonly PdfNativeTextPropertySegmentDraft[],
  range: Readonly<{ start: number; end: number }>,
  base: PdfNativeTextSegmentFields,
): Readonly<Record<(typeof FIELDS)[number], unknown>> {
  const result: Record<string, unknown> = {};
  for (const key of FIELDS) {
    const values: unknown[] = [];
    let cursor = range.start;
    for (const segment of [...segments].sort((left, right) => left.start - right.start)) {
      if (segment.end <= range.start || segment.start >= range.end) continue;
      if (segment.start > cursor) values.push(base[key]);
      values.push(segment[key] ?? base[key]);
      cursor = segment.end;
    }
    if (cursor < range.end) values.push(base[key]);
    result[key] = values.every((value) => sameValue(value, values[0])) ? values[0] : "mixed";
  }
  return result as Record<(typeof FIELDS)[number], unknown>;
}
