/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Delete one original text object from a page and put it back.
 *
 * The selected run's whole `BT … ET` object becomes an empty `BT ET`; Undo
 * writes the recorded object back over it. Both directions are one operator
 * span in the page's own content stream.
 */
import { fail, offset, pdfNativeTextObjectDeleteHistory } from "./native-text-object-delete-model";
export { pdfNativeTextObjectDeleteHistory };
import {
  MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION,
  PDF_NATIVE_TEXT_OBJECT_DELETE_CONTROL_ID,
} from "./engine/native/vocabulary";
export {
  MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION,
  PDF_NATIVE_TEXT_OBJECT_DELETE_CONTROL_ID,
};
import {
  MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
  planMraiNativeTextObjectDelete,
  planMraiNativeTextObjectDeleteRestore,
  prepareMraiNativeTextObjectDelete,
  verifyMraiNativeTextObjectDeleteCandidate,
  type PdfNativeTextObjectDeletePlan,
  type PdfNativeTextObjectDeleteScope,
} from "./engine/native/text-object-delete";
import {
  sha256PdfBytes,
  type PdfNativePreparedContentStreamTransaction,
} from "./engine/native/content-stream-transaction";
import {
  planMraiNativeStructActualTextForRequest,
  type PdfNativeListTextFidelityEvidence,
  type PdfNativeListTextSelectionEvidence,
} from "./engine/native/list-text-fidelity";

export { MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION };


export interface PdfNativeTextObjectDeleteHistory {
  readonly offset: number;
  readonly objectBytes: readonly number[];
  /** What the delete emptied; absent on records from before shows were deletable. */
  readonly scope?: PdfNativeTextObjectDeleteScope;
  /** The pen number a show deleted keeping its advance was left as; written
   * before a span of several shows could each keep theirs. */
  readonly penRestore?: string;
  /** The advances the last shows of the span were left as, in the order the
   * span holds them. */
  readonly penRestores?: readonly string[];
  /** The character spacing each of those advances is drawn under, in the
   * same order (zero where it is none): a pen under a non-zero spacing is
   * written where the reader reads it, and the restore finds it there. */
  readonly penCharacterSpacings?: readonly number[];
}

/** One of the page's content streams, by object identity. */
export type PdfNativeTextObjectDeleteStream = Readonly<{
  objectNumber: number;
  generation: number;
}>;

export interface PdfNativeTextObjectDeletePayload {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION;
  readonly selection: Readonly<{
    pageNumber: number;
    runSpan: Readonly<{ start: number; end: number }>;
    /**
     * Which of the page's content streams the span above is an offset into.
     * A page drawn from several streams is one drawing, so the span alone
     * does not name the part of it that holds the run. Absent means the page
     * draws from one stream, which is then the only answer.
     */
    contentStream?: PdfNativeTextObjectDeleteStream;
    /** The reader selection the run was proven from: a show whose next show
     * draws on from its end is deleted keeping its advance. */
    textSelection?: PdfNativeListTextSelectionEvidence;
    /** The underline and strikethrough read beside the run, with the
     * geometry they were drawn from: taken away in the same write. */
    keepTextRules?: Readonly<{
      styles: readonly ("underline" | "strikethrough")[];
      geometry: Readonly<{
        textMatrix: readonly [number, number, number, number, number, number];
        fontSize: number;
        advance: number;
        color: Readonly<{ kind: "gray" | "rgb" | "cmyk"; components: readonly number[] }>;
      }>;
    }>;
  }>;
  readonly direction: "delete" | "restore";
  /** Restore only: where the empty object sits and what replaces it. */
  readonly history?: PdfNativeTextObjectDeleteHistory;
}

export interface PdfNativeTextObjectDeleteEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly direction: "delete" | "restore";
  readonly scope: PdfNativeTextObjectDeleteScope;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly objectBytes: readonly number[];
  readonly showCount: number;
  /** A show deleted keeping its advance: the pen number it was left as, and
   * (on a delete) the proof of the run that number measures. */
  readonly penRestore?: Readonly<{
    number: string;
    sourceFidelity?: PdfNativeListTextFidelityEvidence;
    characterSpacing?: number;
    precedingShows?: readonly Readonly<{
      number: string;
      sourceFidelity?: PdfNativeListTextFidelityEvidence;
      characterSpacing?: number;
    }>[];
  }>;
}

export interface PdfNativeTextObjectDeletePlanResult {
  readonly payload: PdfNativeTextObjectDeletePayload;
  readonly evidence: PdfNativeTextObjectDeleteEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly plan: PdfNativeTextObjectDeletePlan;
}


function plainRecord(value: unknown): Record<string, unknown> | null {
  return typeof value === "object" && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}


function readScope(value: unknown): PdfNativeTextObjectDeleteScope {
  if (value !== "object" && value !== "show" && value !== "objects" && value !== "run") {
    fail("Invalid native PDF text object delete scope.");
  }
  return value;
}

function byteList(value: unknown): readonly number[] {
  if (
    !Array.isArray(value) ||
    value.length < 6 ||
    value.length > 4 * 1024 * 1024 ||
    value.some(
      (item) =>
        typeof item !== "number" ||
        !Number.isSafeInteger(item) ||
        item < 0 ||
        item > 255,
    )
  ) {
    fail("Invalid native PDF text object bytes.");
  }
  return Object.freeze(value.map(Number));
}

function readContentStream(value: unknown): PdfNativeTextObjectDeleteStream {
  const record = plainRecord(value);
  if (!record) fail("Invalid native PDF text object content stream.");
  return Object.freeze({
    objectNumber: offset(record.objectNumber, "stream object", false),
    generation: offset(record.generation, "stream generation"),
  });
}

function readKeepTextRules(value: unknown): NonNullable<PdfNativeTextObjectDeletePayload["selection"]["keepTextRules"]> {
  const record = plainRecord(value);
  const geometry = record ? plainRecord(record.geometry) : null;
  const color = geometry ? plainRecord(geometry.color) : null;
  const number = (item: unknown) => typeof item === "number" && Number.isFinite(item);
  if (
    !record || !geometry || !color ||
    !Array.isArray(record.styles) || record.styles.length === 0 || record.styles.length > 2 ||
    record.styles.some((style) => style !== "underline" && style !== "strikethrough") ||
    !Array.isArray(geometry.textMatrix) || geometry.textMatrix.length !== 6 || !geometry.textMatrix.every(number) ||
    !number(geometry.fontSize) || !number(geometry.advance) ||
    (color.kind !== "gray" && color.kind !== "rgb" && color.kind !== "cmyk") ||
    !Array.isArray(color.components) || !color.components.every(number)
  ) {
    fail("Invalid native PDF text rule request.");
  }
  return Object.freeze({
    styles: Object.freeze([...(record.styles as ("underline" | "strikethrough")[])]),
    geometry: Object.freeze({
      textMatrix: Object.freeze([...(geometry.textMatrix as number[])]) as unknown as readonly [number, number, number, number, number, number],
      fontSize: geometry.fontSize as number,
      advance: geometry.advance as number,
      color: Object.freeze({ kind: color.kind as "gray" | "rgb" | "cmyk", components: Object.freeze([...(color.components as number[])]) }),
    }),
  });
}

export function readPdfNativeTextObjectDeletePayload(
  value: unknown,
): PdfNativeTextObjectDeletePayload {
  const payload = plainRecord(value);
  if (
    !payload ||
    payload.schemaVersion !== MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION
  ) {
    fail("Unsupported native PDF text object delete schema.");
  }
  const selection = plainRecord(payload.selection);
  const runSpan = selection ? plainRecord(selection.runSpan) : null;
  if (!selection || !runSpan) fail("Invalid native PDF text object selection.");
  if (payload.direction !== "delete" && payload.direction !== "restore") {
    fail("Invalid native PDF text object delete direction.");
  }
  const contentStream =
    selection.contentStream === undefined
      ? undefined
      : readContentStream(selection.contentStream);
  const start = offset(runSpan.start, "span start");
  const end = offset(runSpan.end, "span end", false);
  if (end <= start) fail("Invalid native PDF text object span.");
  const textSelection = selection.textSelection === undefined
    ? undefined
    : plainRecord(selection.textSelection);
  if (selection.textSelection !== undefined && !textSelection) {
    fail("Invalid native PDF text object selection.");
  }
  const keepTextRules = selection.keepTextRules === undefined ? undefined : readKeepTextRules(selection.keepTextRules);
  const base = Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
    selection: Object.freeze({
      pageNumber: offset(selection.pageNumber, "page", false),
      runSpan: Object.freeze({ start, end }),
      ...(contentStream === undefined ? {} : { contentStream }),
      // Proven against the source by the planner; only carried here.
      ...(textSelection
        ? { textSelection: textSelection as unknown as PdfNativeListTextSelectionEvidence }
        : {}),
      ...(keepTextRules ? { keepTextRules } : {}),
    }),
    direction: payload.direction,
  });
  if (payload.direction === "delete") {
    if (payload.history !== undefined) {
      fail("Native PDF text object delete carries history it cannot use.");
    }
    return base;
  }
  const history = plainRecord(payload.history);
  if (!history) fail("Invalid native PDF text object delete history.");
  return Object.freeze({
    ...base,
    history: Object.freeze({
      offset: offset(history.offset, "history offset"),
      objectBytes: byteList(history.objectBytes),
      ...(history.scope === undefined
        ? {}
        : { scope: readScope(history.scope) }),
      ...(history.penRestore === undefined
        ? {}
        : typeof history.penRestore === "string"
          ? { penRestore: history.penRestore }
          : fail("Invalid native PDF text object delete history.")),
      ...(history.penRestores === undefined
        ? {}
        : Array.isArray(history.penRestores) &&
            history.penRestores.every((pen) => typeof pen === "string")
          ? { penRestores: Object.freeze([...(history.penRestores as string[])]) }
          : fail("Invalid native PDF text object delete history.")),
      ...(history.penCharacterSpacings === undefined
        ? {}
        : Array.isArray(history.penCharacterSpacings) &&
            history.penCharacterSpacings.every((value) => typeof value === "number" && Number.isFinite(value))
          ? { penCharacterSpacings: Object.freeze([...(history.penCharacterSpacings as number[])]) }
          : fail("Invalid native PDF text object delete history.")),
    }),
  });
}

export async function planPdfNativeTextObjectDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeTextObjectDeletePayload;
}): Promise<PdfNativeTextObjectDeletePlanResult> {
  const payload = readPdfNativeTextObjectDeletePayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  const planned =
    payload.direction === "restore"
      ? await planMraiNativeTextObjectDeleteRestore({
          sourceBytes: input.sourceBytes,
          pageNumber: payload.selection.pageNumber,
          offset: payload.history!.offset,
          objectBytes: payload.history!.objectBytes,
          contentStream: payload.selection.contentStream ?? null,
          scope: payload.history!.scope ?? "object",
          penRestore: payload.history!.penRestore ?? null,
          penRestores: payload.history!.penRestores ?? null,
          penCharacterSpacings: payload.history!.penCharacterSpacings ?? null,
        })
      : await planMraiNativeTextObjectDelete({
          sourceBytes: input.sourceBytes,
          pageNumber: payload.selection.pageNumber,
          runSpan: payload.selection.runSpan,
          contentStream: payload.selection.contentStream ?? null,
          textSelection: payload.selection.textSelection ?? null,
          keepTextRules: payload.selection.keepTextRules ?? null,
        });
  // The structure elements whose /ActualText names the text taken away (or
  // put back) are written in the same revision: each then reads what its
  
  const request = await planMraiNativeStructActualTextForRequest({
    sourceBytes: input.sourceBytes,
    request: planned.request,
    pageNumber: planned.pageNumber,
  });
  const plan = request === planned.request ? planned : Object.freeze({ ...planned, request });
  const prepared = await prepareMraiNativeTextObjectDelete(
    input.sourceBytes,
    plan,
  );
  await verifyMraiNativeTextObjectDeleteCandidate(
    input.sourceBytes,
    plan,
    prepared,
  );
  const evidence: PdfNativeTextObjectDeleteEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
    sourceSha256,
    direction: plan.direction,
    scope: plan.scope,
    pageNumber: plan.pageNumber,
    page: Object.freeze({ ...plan.page }),
    contentStream: Object.freeze({ ...plan.contentStream }),
    operatorSpan: Object.freeze({ ...plan.operatorSpan }),
    objectBytes: plan.objectBytes,
    showCount: plan.showCount,
    ...(plan.penRestore
      ? { penRestore: Object.freeze({ ...plan.penRestore }) }
      : plan.direction === "restore" && payload.history?.penRestore
        ? { penRestore: restoredPens(payload.history) }
        : {}),
  });
  return Object.freeze({ payload, evidence, prepared, plan });
}

/** The pens a restore finds the delete left, as its record names them: every
 * advance with the character spacing it was written under (mirror of
 * replay's _pen_numbers and _pen_spacings). */
function restoredPens(
  history: PdfNativeTextObjectDeleteHistory,
): NonNullable<PdfNativeTextObjectDeleteEvidence["penRestore"]> {
  const numbers = history.penRestores ?? [history.penRestore!];
  const pens = numbers.map((number, index) => {
    const spacing = history.penCharacterSpacings?.[index] ?? 0;
    return Object.freeze({ number, ...(spacing !== 0 ? { characterSpacing: spacing } : {}) });
  });
  const last = pens.pop()!;
  return Object.freeze({ ...last, ...(pens.length > 0 ? { precedingShows: Object.freeze(pens) } : {}) });
}

/**
 * The page's content streams of two files of the same content (the version
 * the server rewrote as another file, and the bytes a queued delete wrote
 * when it was planned): what a delete that no longer reads as its run is
 * stood on (native-text-delete-retarget: sameContentStreams).
 */
export async function readPdfNativeSameContentDeleteStreams(input: {
  readonly savedBytes: Uint8Array;
  readonly plannedBytes: Uint8Array;
  readonly pageNumber: number;
}): Promise<Readonly<{
  saved: readonly Readonly<{ ref: Readonly<{ objectNumber: number; generation: number }>; bytes: Uint8Array }>[];
  planned: readonly Uint8Array[];
}>> {
  const { inspectMraiNativePageContentStreams } = await import("./engine/native/content-stream-transaction");
  const saved = await inspectMraiNativePageContentStreams(input.savedBytes, input.pageNumber);
  const planned = await inspectMraiNativePageContentStreams(input.plannedBytes, input.pageNumber);
  return Object.freeze({
    saved: saved.map((stream) => Object.freeze({ ref: Object.freeze({ objectNumber: stream.contentStream.objectNumber,
      generation: stream.contentStream.generation }), bytes: Uint8Array.from(stream.streamBytes) })),
    planned: planned.map((stream) => Uint8Array.from(stream.streamBytes)),
  });
}
