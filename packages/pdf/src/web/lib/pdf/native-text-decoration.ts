/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Underline and strikethrough for one original text run, written as page
 * content next to the run's own text object.
 */
import { fail, pdfNativeTextDecorationHistory } from "./native-text-decoration-model";
export { pdfNativeTextDecorationHistory };
import {
  MRAI_NATIVE_TEXT_DECORATION_ACTION,
  PDF_NATIVE_TEXT_DECORATION_CONTROL_ID,
} from "./engine/native/vocabulary";
export {
  MRAI_NATIVE_TEXT_DECORATION_ACTION,
  PDF_NATIVE_TEXT_DECORATION_CONTROL_ID,
};
import {
  MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
  findMraiNativeTextDecoration,
  planMraiNativeTextDecoration,
  planMraiNativeTextDecorationRemoval,
  prepareMraiNativeTextDecoration,
  verifyMraiNativeTextDecorationCandidate,
  type PdfNativeTextDecorationGeometry,
  type PdfNativeTextDecorationPlan,
  type PdfNativeTextDecorationStyle,
} from "./engine/native/text-decoration";
import {
  sha256PdfBytes,
  type PdfNativePreparedContentStreamTransaction,
} from "./engine/native/content-stream-transaction";

export {
  MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
  findMraiNativeTextDecoration,
  type PdfNativeTextDecorationGeometry,
  type PdfNativeTextDecorationStyle,
};


export interface PdfNativeTextDecorationHistory {
  readonly offset: number;
  readonly ruleBytes: readonly number[];
}

/** One of the page's content streams, by object identity. */
export type PdfNativeTextDecorationStream = Readonly<{
  objectNumber: number;
  generation: number;
}>;

export interface PdfNativeTextDecorationPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION;
  readonly pageNumber: number;
  readonly style: PdfNativeTextDecorationStyle;
  readonly direction: "add" | "remove";
  /**
   * Which of the page's content streams the offsets below belong to. A page
   * drawn from several streams is one drawing, and both the run's span and
   * the rule's own offset are offsets into one part of it. Absent means the
   * page draws from one stream, which is then the only answer.
   */
  readonly contentStream?: PdfNativeTextDecorationStream;
  /** Add only: the run's own matrix, size, advance and colour. */
  readonly geometry?: PdfNativeTextDecorationGeometry;
  /** Remove only: where the rule sits and the exact bytes to take out. */
  readonly history?: PdfNativeTextDecorationHistory;
}

export interface PdfNativeTextDecorationEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly direction: "add" | "remove";
  readonly style: PdfNativeTextDecorationStyle;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly ruleBytes: readonly number[];
  readonly ruleOffset: number;
}

export interface PdfNativeTextDecorationPlanResult {
  readonly payload: PdfNativeTextDecorationPayload;
  readonly evidence: PdfNativeTextDecorationEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly plan: PdfNativeTextDecorationPlan;
}


function plainRecord(value: unknown): Record<string, unknown> | null {
  return typeof value === "object" && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function wholeNumber(value: unknown, label: string, allowZero = true): number {
  if (
    typeof value !== "number" ||
    !Number.isSafeInteger(value) ||
    value < (allowZero ? 0 : 1)
  ) {
    fail(`Invalid native PDF text rule ${label}.`);
  }
  return value;
}

function finiteNumber(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail(`Invalid native PDF text rule ${label}.`);
  }
  return value;
}

function byteList(value: unknown, label: string): readonly number[] {
  if (
    !Array.isArray(value) ||
    value.length < 8 ||
    value.length > 64 * 1024 ||
    value.some(
      (item) =>
        typeof item !== "number" ||
        !Number.isSafeInteger(item) ||
        item < 0 ||
        item > 255,
    )
  ) {
    fail(`Invalid native PDF text rule ${label}.`);
  }
  return Object.freeze(value.map(Number));
}

function readGeometry(value: unknown): PdfNativeTextDecorationGeometry {
  const record = plainRecord(value);
  const matrix = record?.textMatrix;
  const color = plainRecord(record?.color);
  const runSpan = plainRecord(record?.runSpan);
  if (
    !record ||
    !Array.isArray(matrix) ||
    matrix.length !== 6 ||
    !color ||
    !runSpan ||
    !Array.isArray(color.components) ||
    (color.kind !== "gray" && color.kind !== "rgb" && color.kind !== "cmyk")
  ) {
    fail("Invalid native PDF text rule geometry.");
  }
  return Object.freeze({
    textMatrix: Object.freeze(
      matrix.map((entry, index) => finiteNumber(entry, `matrix ${index}`)),
    ) as PdfNativeTextDecorationGeometry["textMatrix"],
    fontSize: finiteNumber(record.fontSize, "font size"),
    advance: finiteNumber(record.advance, "advance"),
    color: Object.freeze({
      kind: color.kind,
      components: Object.freeze(
        (color.components as unknown[]).map((entry, index) =>
          finiteNumber(entry, `colour ${index}`),
        ),
      ),
    }),
    runSpan: Object.freeze({
      start: wholeNumber(runSpan.start, "run start"),
      end: wholeNumber(runSpan.end, "run end", false),
    }),
  });
}

function readContentStream(value: unknown): PdfNativeTextDecorationStream {
  const record = plainRecord(value);
  if (!record) fail("Invalid native PDF text rule content stream.");
  return Object.freeze({
    objectNumber: wholeNumber(record.objectNumber, "stream object", false),
    generation: wholeNumber(record.generation, "stream generation"),
  });
}

export function readPdfNativeTextDecorationPayload(
  value: unknown,
): PdfNativeTextDecorationPayload {
  const payload = plainRecord(value);
  if (
    !payload ||
    payload.schemaVersion !== MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION ||
    (payload.style !== "underline" && payload.style !== "strikethrough") ||
    (payload.direction !== "add" && payload.direction !== "remove")
  ) {
    fail("Unsupported native PDF text rule payload.");
  }
  const base = Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
    pageNumber: wholeNumber(payload.pageNumber, "page", false),
    style: payload.style,
    direction: payload.direction,
    ...(payload.contentStream === undefined
      ? {}
      : { contentStream: readContentStream(payload.contentStream) }),
  });
  if (payload.direction === "add") {
    if (payload.history !== undefined) {
      fail("Native PDF text rule carries history it cannot use.");
    }
    return Object.freeze({ ...base, geometry: readGeometry(payload.geometry) });
  }
  const history = plainRecord(payload.history);
  if (!history || payload.geometry !== undefined) {
    fail("Invalid native PDF text rule history.");
  }
  return Object.freeze({
    ...base,
    history: Object.freeze({
      offset: wholeNumber(history.offset, "offset"),
      ruleBytes: byteList(history.ruleBytes, "bytes"),
    }),
  });
}

export async function planPdfNativeTextDecoration(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeTextDecorationPayload;
}): Promise<PdfNativeTextDecorationPlanResult> {
  const payload = readPdfNativeTextDecorationPayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  const plan =
    payload.direction === "remove"
      ? await planMraiNativeTextDecorationRemoval({
          sourceBytes: input.sourceBytes,
          pageNumber: payload.pageNumber,
          style: payload.style,
          offset: payload.history!.offset,
          ruleBytes: payload.history!.ruleBytes,
          contentStream: payload.contentStream ?? null,
        })
      : await planMraiNativeTextDecoration({
          sourceBytes: input.sourceBytes,
          pageNumber: payload.pageNumber,
          style: payload.style,
          geometry: payload.geometry!,
          contentStream: payload.contentStream ?? null,
        });
  const prepared = await prepareMraiNativeTextDecoration(
    input.sourceBytes,
    plan,
  );
  await verifyMraiNativeTextDecorationCandidate(
    input.sourceBytes,
    plan,
    prepared,
  );
  const evidence: PdfNativeTextDecorationEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
    sourceSha256,
    direction: plan.direction,
    style: plan.style,
    pageNumber: plan.pageNumber,
    page: Object.freeze({ ...plan.page }),
    contentStream: Object.freeze({ ...plan.contentStream }),
    operatorSpan: Object.freeze({ ...plan.operatorSpan }),
    ruleBytes: plan.ruleBytes,
    ruleOffset: plan.ruleOffset,
  });
  return Object.freeze({ payload, evidence, prepared, plan });
}

