/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfTextEditCommonEnds } from "@mrnote/core/web/lib/pdf/text-edit-common-ends.ts";



























export type PdfTextPartScript = "superscript" | "subscript" | "none";

export interface PdfTextPartFields {
  /** The value the font list gave (a family name, one of the PDF's own
   * fonts by its resource name, a page font's token). */
  readonly family?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  readonly script?: PdfTextPartScript;
  /** The size on paper, in points. */
  readonly sizePoints?: number;
  /** `#rrggbb`. */
  readonly color?: string;
}

export interface PdfTextPartFormat extends PdfTextPartFields {
  readonly start: number;
  readonly end: number;
}

/** Sorted, disjoint parts, each with at least one field. */
export type PdfTextPartFormats = readonly PdfTextPartFormat[];

export const PDF_TEXT_PART_FIELDS = [
  "family", "bold", "italic", "underline", "strikethrough", "script", "sizePoints", "color",
] as const satisfies readonly (keyof PdfTextPartFields)[];

type FieldKey = (typeof PDF_TEXT_PART_FIELDS)[number];

function fieldsOf(part: PdfTextPartFields): PdfTextPartFields {
  const fields: Record<string, unknown> = {};
  for (const key of PDF_TEXT_PART_FIELDS) if (part[key] !== undefined) fields[key] = part[key];
  return fields as PdfTextPartFields;
}

function sameFields(left: PdfTextPartFields, right: PdfTextPartFields): boolean {
  return PDF_TEXT_PART_FIELDS.every((key) => left[key] === right[key]);
}

function hasFields(part: PdfTextPartFields): boolean {
  return PDF_TEXT_PART_FIELDS.some((key) => part[key] !== undefined);
}

/** The parts in order, adjacent parts of one format joined, empty ones gone. */
export function normalizePdfTextPartFormats(parts: readonly PdfTextPartFormat[]): PdfTextPartFormats {
  const sorted = parts
    .filter((part) => Number.isFinite(part.start) && Number.isFinite(part.end) && part.end > part.start && hasFields(part))
    .map((part) => Object.freeze({ ...fieldsOf(part), start: part.start, end: part.end }))
    .sort((left, right) => left.start - right.start);
  const joined: PdfTextPartFormat[] = [];
  for (const part of sorted) {
    const previous = joined[joined.length - 1];
    if (previous && previous.end === part.start && sameFields(previous, part)) {
      joined[joined.length - 1] = Object.freeze({ ...previous, end: part.end });
    } else {
      joined.push(part);
    }
  }
  return Object.freeze(joined);
}

/**
 * The parts after `fields` were pressed on [start, end): where the press
 * covers a part, its fields are set over that part's (other fields kept);
 * where it covers nothing yet, a part of its own. A field given as
 * `undefined` is taken off (the line's own again).
 */
export function applyPdfTextPartFormat(
  parts: PdfTextPartFormats,
  range: Readonly<{ start: number; end: number }>,
  fields: PdfTextPartFields,
): PdfTextPartFormats {
  const start = Math.max(0, Math.trunc(range.start));
  const end = Math.trunc(range.end);
  if (!(end > start)) return parts;
  const pressed = Object.keys(fields) as FieldKey[];
  if (pressed.length === 0) return parts;
  const out: PdfTextPartFormat[] = [];
  let cursor = start;
  const press = (base: PdfTextPartFields, from: number, to: number) => {
    const next: Record<string, unknown> = { ...fieldsOf(base) };
    for (const key of pressed) {
      if (fields[key] === undefined) delete next[key];
      else next[key] = fields[key];
    }
    out.push({ ...(next as PdfTextPartFields), start: from, end: to });
  };
  for (const part of [...parts].sort((left, right) => left.start - right.start)) {
    if (part.end <= start || part.start >= end) {
      out.push(part);
      continue;
    }
    if (part.start < start) out.push({ ...part, end: start });
    const from = Math.max(part.start, start), to = Math.min(part.end, end);
    if (from > cursor) press({}, cursor, from);
    press(part, from, to);
    cursor = to;
    if (part.end > end) out.push({ ...part, start: end });
  }
  if (cursor < end) press({}, cursor, end);
  return normalizePdfTextPartFormats(out);
}

/**
 * The value of each field over [start, end) as the parts press it: set when
 * every non-blank character of the range has it pressed alike (blanks take
 * any), absent otherwise. A collapsed range reads the character before it
 * (Word), the first character at the start.
 */
export function pdfTextPartFieldsOver(
  parts: PdfTextPartFormats,
  range: Readonly<{ start: number; end: number }> | null,
  text: string,
): PdfTextPartFields {
  if (parts.length === 0) return {};
  let start = range ? range.start : 0;
  let end = range ? range.end : text.length;
  if (!(end > start)) {
    start = Math.max(0, start - 1);
    end = Math.min(text.length, start + 1);
    if (!(end > start)) return {};
  }
  const out: Record<string, unknown> = {};
  for (const key of PDF_TEXT_PART_FIELDS) {
    let value: unknown;
    let seen = false;
    let uniform = true;
    for (let index = start; index < end && uniform; index += 1) {
      if (!(text[index] ?? "").trim() && end - start > 1) continue;
      const part = parts.find((entry) => entry.start <= index && index < entry.end);
      const own = part?.[key];
      if (own === undefined) { uniform = false; break; }
      if (!seen) { value = own; seen = true; }
      else if (value !== own) uniform = false;
    }
    if (uniform && seen) out[key] = value;
  }
  return out as PdfTextPartFields;
}

/**
 * The parts moved through one edit of the text (`before` to `after`): the
 * longest common head and tail stand; characters typed between take the
 * fields of the character before them (or after them, at the start); the
 * characters taken out take their parts with them.
 */
export function pdfTextPartFormatsAfterEdit(
  parts: PdfTextPartFormats,
  before: string,
  after: string,
): PdfTextPartFormats {
  if (before === after || parts.length === 0) return parts;
  const { prefix, suffix } = pdfTextEditCommonEnds(before, after);
  const removedEnd = before.length - suffix;
  const insertedEnd = after.length - suffix;
  const shift = after.length - before.length;
  const inheritAt = prefix > 0 ? prefix - 1 : removedEnd;
  const inherited = parts.find((part) => part.start <= inheritAt && inheritAt < part.end) ?? null;
  const out: PdfTextPartFormat[] = [];
  for (const part of parts) {
    // The head of the part before the edit; its tail after it, shifted.
    if (part.start < prefix) out.push({ ...part, end: Math.min(part.end, prefix) });
    if (part.end > removedEnd) out.push({ ...part, start: Math.max(part.start, removedEnd) + shift, end: part.end + shift });
  }
  if (inherited && insertedEnd > prefix) out.push({ ...fieldsOf(inherited), start: prefix, end: insertedEnd });
  return normalizePdfTextPartFormats(out);
}

/** The parts that press `key`, with that field's value. */
export function pdfTextPartFormatsWith<Key extends FieldKey>(
  parts: PdfTextPartFormats,
  key: Key,
): readonly Readonly<{ start: number; end: number; value: NonNullable<PdfTextPartFields[Key]> }>[] {
  return parts.flatMap((part) => part[key] === undefined ? []
    : [Object.freeze({ start: part.start, end: part.end, value: part[key] as NonNullable<PdfTextPartFields[Key]> })]);
}

/** The rule (underline, strikethrough) parts as on/off stretches, joined. */
export function pdfTextPartRuleStretches(
  parts: PdfTextPartFormats,
  style: "underline" | "strikethrough",
): readonly Readonly<{ start: number; end: number; on: boolean }>[] {
  const stretches: { start: number; end: number; on: boolean }[] = [];
  for (const part of pdfTextPartFormatsWith(parts, style)) {
    const previous = stretches[stretches.length - 1];
    if (previous && previous.end === part.start && previous.on === part.value) previous.end = part.end;
    else stretches.push({ start: part.start, end: part.end, on: part.value });
  }
  return Object.freeze(stretches.map((stretch) => Object.freeze(stretch)));
}

/**
 * The characters a rule runs under once the stretches pressed are set over
 * the ones it runs under now (`standing`, a [start, end) of the same text,
 * null for none): one stretch, `null` for none, or `"apart"` when it would
 * leave two stretches apart -- a run carries one rule of a style (the
 * decoration writer), and Word's two underlined words with a plain one
 * between them is a format the page cannot draw beside one run. Blanks
 * between two stretches join them (Word underlines the space between two
 * underlined words).
 */
export function pdfTextRuleCharactersAfter(
  standing: Readonly<{ start: number; end: number }> | null,
  stretches: readonly Readonly<{ start: number; end: number; on: boolean }>[],
  text: string,
): Readonly<{ start: number; end: number }> | null | "apart" {
  const on = new Uint8Array(text.length);
  if (standing) for (let index = Math.max(0, standing.start); index < Math.min(text.length, standing.end); index += 1) on[index] = 1;
  for (const stretch of stretches) {
    for (let index = Math.max(0, stretch.start); index < Math.min(text.length, stretch.end); index += 1) {
      on[index] = stretch.on ? 1 : 0;
    }
  }
  let first = -1, last = -1;
  for (let index = 0; index < text.length; index += 1) {
    if (!on[index]) continue;
    if (first < 0) first = index;
    last = index;
  }
  if (first < 0) return null;
  for (let index = first; index <= last; index += 1) {
    if (!on[index] && (text[index] ?? "").trim()) return "apart";
  }
  return Object.freeze({ start: first, end: last + 1 });
}
