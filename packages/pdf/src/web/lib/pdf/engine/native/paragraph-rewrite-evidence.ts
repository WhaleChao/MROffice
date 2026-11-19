/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextFidelityEvidence } from "./list-text-fidelity";
import type { PdfNativeParagraphRewriteAlignment, PdfNativeParagraphRewriteEvidence } from "./paragraph-rewrite";

export const MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION =
  "mrai.pdf-native-paragraph-rewrite.v1" as const;

export function readRecord(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null;
}

export function readFinite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

export function readSpan(value: unknown): Readonly<{ start: number; end: number }> | null {
  const span = readRecord(value);
  return span && Number.isSafeInteger(span.start) && Number.isSafeInteger(span.end) &&
    (span.start as number) >= 0 && (span.end as number) > (span.start as number)
    ? Object.freeze({ start: span.start as number, end: span.end as number }) : null;
}

export function readByteArray(value: unknown, limit: number): readonly number[] | null {
  return Array.isArray(value) && value.length <= limit &&
    value.every((byte) => Number.isInteger(byte) && byte >= 0 && byte <= 255)
    ? Object.freeze([...value as number[]]) : null;
}

/** A stored rewrite read back field by field; null when anything is off. */
export function readMraiNativeParagraphRewriteEvidence(
  value: unknown,
  readEvidence: (value: unknown) => PdfNativeListTextFidelityEvidence | null,
): PdfNativeParagraphRewriteEvidence | null {
  const record = readRecord(value);
  const target = readRecord(record?.target);
  if (!record || record.schemaVersion !== MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION || !target ||
    !Number.isSafeInteger(target.objectNumber) || (target.objectNumber as number) < 1 ||
    !Number.isSafeInteger(target.generation) || (target.generation as number) < 0 ||
    !Array.isArray(record.items) || record.items.length < 1 || record.items.length > 256 ||
    !Array.isArray(record.shows) || record.shows.length < 1 || record.shows.length > 4_096 ||
    !Array.isArray(record.pens) || record.pens.length > record.shows.length) return null;
  const items: (PdfNativeListTextFidelityEvidence | null)[] = [];
  for (const item of record.items) {
    if (item === null) { items.push(null); continue; }
    const read = readEvidence(item);
    if (!read || read.paragraphRewrite) return null;
    items.push(read);
  }
  const shows: Readonly<{ start: number; end: number }>[] = [];
  for (const show of record.shows) {
    const span = readSpan(show);
    if (!span || (shows.length && span.start < shows[shows.length - 1].end)) return null;
    shows.push(span);
  }
  const pens: PdfNativeParagraphRewriteEvidence["pens"][number][] = [];
  for (const value of record.pens) {
    const pen = readRecord(value);
    const span = readSpan(pen);
    const sourceFidelity = pen ? readEvidence(pen.sourceFidelity) : null;
    const chunk = pen?.chunk === undefined || pen.chunk === null ? null : readRecord(pen.chunk);
    const spacing = pen?.characterSpacing;
    if (!pen || !span || typeof pen.number !== "string" || !/^-?\d+(?:\.\d+)?$/u.test(pen.number) ||
      pen.number.length > 32 || !sourceFidelity ||
      !shows.some((show) => show.start === span.start && show.end === span.end) ||
      (pen.chunk !== undefined && pen.chunk !== null && (!chunk || !Number.isSafeInteger(chunk.codeStart) ||
        !Number.isSafeInteger(chunk.codeEnd) || (chunk.codeStart as number) < 0 ||
        !((chunk.codeStart as number) < (chunk.codeEnd as number)))) ||
      (spacing !== undefined && spacing !== null && (typeof spacing !== "number" || !Number.isFinite(spacing) ||
        spacing === 0))) return null;
    pens.push(Object.freeze({ ...span, number: pen.number, sourceFidelity,
      ...(chunk ? { chunk: Object.freeze({ codeStart: chunk.codeStart as number, codeEnd: chunk.codeEnd as number }) } : {}),
      ...(typeof spacing === "number" ? { characterSpacing: spacing } : {}) }));
  }
  const originalBytes = readByteArray(record.originalBytes, 16 * 1024 * 1024);
  if (!originalBytes || !Array.isArray(record.pins) || record.pins.length > 256) return null;
  const pins: PdfNativeParagraphRewriteEvidence["pins"][number][] = [];
  for (const value of record.pins) {
    const pin = readRecord(value);
    if (!pin || !Number.isSafeInteger(pin.objectNumber) || (pin.objectNumber as number) < 1 ||
      !Number.isSafeInteger(pin.generation) || (pin.generation as number) < 0 ||
      (pin.role !== "resource" && pin.role !== "untouched-object")) return null;
    pins.push(Object.freeze({ objectNumber: pin.objectNumber as number, generation: pin.generation as number, role: pin.role }));
  }
  let block: PdfNativeParagraphRewriteEvidence["block"] = null;
  if (record.block != null) {
    const raw = readRecord(record.block);
    const numbers = raw ? ["frameLeft", "frameRight", "firstLeft", "restLeft", "firstBaseline", "lineStep"]
      .map((key) => readFinite(raw[key])) : [];
    const basis = raw && Array.isArray(raw.basis) && raw.basis.length === 4 ? raw.basis.map(readFinite) : null;
    const bytes = raw ? readByteArray(raw.bytes, 4 * 1024 * 1024) : null;
    if (!raw || !Number.isSafeInteger(raw.start) || !bytes || !bytes.length || !basis ||
      basis.some((entry) => entry === null) || numbers.some((entry) => entry === null) ||
      !["left", "center", "right", "justify"].includes(String(raw.alignment)) ||
      !Number.isSafeInteger(raw.originalLineCount) || (raw.originalLineCount as number) < 1 ||
      !Number.isSafeInteger(raw.lineCount) || (raw.lineCount as number) < 1 ||
      !Array.isArray(raw.lines) || raw.lines.length !== raw.lineCount ||
      (raw.frameRequested !== undefined && raw.frameRequested !== null && raw.frameRequested !== true) ||
      (raw.lineStepRequested !== undefined && raw.lineStepRequested !== null && raw.lineStepRequested !== true) ||
      (raw.sizeScale !== undefined && raw.sizeScale !== null && readFinite(raw.sizeScale) === null)) return null;
    const lines: Readonly<{ start: number; end: number }>[] = [];
    for (const line of raw.lines) {
      const span = readSpan(line);
      if (!span || span.start !== (lines.at(-1)?.end ?? 0)) return null;
      lines.push(span);
    }
    // Each line's baseline, where a line set larger stands lower: one per
    // line, each below the one before.
    let lineBaselines: number[] | null = null;
    if (raw.lineBaselines !== undefined && raw.lineBaselines !== null) {
      if (!Array.isArray(raw.lineBaselines) || raw.lineBaselines.length !== raw.lineCount) return null;
      lineBaselines = raw.lineBaselines.map(readFinite) as number[];
      if (lineBaselines.some((baseline, index) => baseline === null ||
        (index > 0 && !(baseline < lineBaselines![index - 1])))) return null;
    }
    const [frameLeft, frameRight, firstLeft, restLeft, firstBaseline, lineStep] = numbers as number[];
    block = Object.freeze({
      start: raw.start as number, bytes, alignment: raw.alignment as PdfNativeParagraphRewriteAlignment,
      basis: Object.freeze(basis as number[]) as unknown as readonly [number, number, number, number],
      frameLeft, frameRight, firstLeft, restLeft, firstBaseline, lineStep,
      originalLineCount: raw.originalLineCount as number, lineCount: raw.lineCount as number,
      lines: Object.freeze(lines),
      ...(lineBaselines ? { lineBaselines: Object.freeze(lineBaselines) } : {}),
      ...(raw.frameRequested === true ? { frameRequested: true as const } : {}),
      ...(raw.lineStepRequested === true ? { lineStepRequested: true as const } : {}),
      ...(typeof raw.sizeScale === "number" ? { sizeScale: raw.sizeScale } : {}),
    });
  }
  let history: PdfNativeParagraphRewriteEvidence["history"];
  if (record.history != null) {
    const raw = readRecord(record.history);
    if (!raw || typeof raw.originTransactionId !== "string" || !raw.originTransactionId ||
      raw.originTransactionId.length > 200 || (raw.direction !== "undo" && raw.direction !== "redo")) return null;
    history = Object.freeze({ originTransactionId: raw.originTransactionId, direction: raw.direction });
  }
  // A format written by state where the text stands: no block, no pen, no
  // source bytes (mirror of the server's NativeParagraphRewrite._in_place).
  const inPlace = record.inPlace === true;
  if (record.inPlace != null && !inPlace) return null;
  if (inPlace && (block || history || pens.length > 0 || originalBytes.length > 0 || record.twins != null)) return null;
  if (!inPlace && originalBytes.length < 1) return null;
  if (!block && !history && !inPlace) return null;
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION,
    target: Object.freeze({ objectNumber: target.objectNumber as number, generation: target.generation as number }),
    items: Object.freeze(items), shows: Object.freeze(shows), pens: Object.freeze(pens), originalBytes,
    pins: Object.freeze(pins), block,
    ...(history ? { history } : {}),
    ...(inPlace ? { inPlace: true as const } : {}),
  });
}
