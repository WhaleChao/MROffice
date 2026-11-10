/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Native picture z-order: move an embedded image to the front or the back of
 * its page by rotating one span of the page content stream around the block
 * that paints it. Every other byte of the file is kept, and the inverse is the
 * same rotation, so Undo puts the original stream back byte for byte.
 */
import { fail, pdfNativeEmbeddedImageReorderHistory } from "./native-embedded-image-reorder-model";
export { pdfNativeEmbeddedImageReorderHistory };
/** A picture moved past its clips drawn whole where it is (one revision of
 * its content stream, published as its bytes). */
export { planMraiNativeEmbeddedImageClipEscape } from "./engine/native/vector-content-transaction";
import {
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION,
} from "./engine/native/vocabulary";
export {
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION,
};
import {
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
  planMraiNativeEmbeddedImageReorder,
  planMraiNativeEmbeddedImageReorderRestore,
  prepareMraiNativeEmbeddedImageReorder,
  pdfNativeEmbeddedImageIsAlreadyArranged,
  verifyMraiNativeEmbeddedImageReorderCandidate,
  MRAI_NATIVE_EMBEDDED_IMAGE_CARRIED_PAINT_STATE_LIMIT,
  type PdfNativeEmbeddedImageMoveLift,
  type PdfNativeEmbeddedImageReorderPlan,
} from "./engine/native/vector-content-transaction";
import {
  sha256PdfBytes,
  type PdfNativePreparedContentStreamTransaction,
} from "./engine/native/content-stream-transaction";

export { MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION, pdfNativeEmbeddedImageIsAlreadyArranged };

export const PDF_NATIVE_EMBEDDED_IMAGE_REORDER_CONTROL_ID =
  "PDFCTL.PAGE.PICTURE.ARRANGE" as const;

export type PdfNativeEmbeddedImageReorderIntent = "front" | "back" | "forward" | "backward";

export interface PdfNativeEmbeddedImageReorderHistory {
  readonly imageOrderAfter: number;
  readonly spanStart: number;
  readonly spanEnd: number;
  readonly blockLength: number;
  readonly blockLeads: boolean;
  /** The transform the forward move wrapped the block in, if it needed one. */
  readonly compensation?: readonly number[];
  /** The graphics states the forward move wrote inside that wrapper, if the
   * picture had to take any with it. */
  readonly carriedStates?: readonly string[];
  /** The groups the forward move took the block out of, if it left any. */
  readonly lift?: PdfNativeEmbeddedImageMoveLift;
  /** The places the forward move emptied of the marked-content markers it
   * took with the picture, and the bytes it left at the span's outer edge;
   * absent when it copied every section it left. */
  readonly emptiedMarkers?: readonly Readonly<{ start: number; end: number }>[];
  readonly stationaryBytes?: number;
  /** True when the forward move was one step past one unit of content. */
  readonly stepped?: boolean;
}

export interface PdfNativeEmbeddedImageReorderPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION;
  readonly selection: Readonly<{ pageNumber: number; imageOrder: number }>;
  readonly direction: PdfNativeEmbeddedImageReorderIntent | "restore";
  /** A front move out from under the clips the picture is drawn in: a
   * picture moved past a clip the page draws everything under (a printed
   * page's printable area) is drawn whole, after the groups that set it. */
  readonly leaveClips?: true;
  /** Restore only: the rotation the forward move wrote. */
  readonly history?: PdfNativeEmbeddedImageReorderHistory;
}

export interface PdfNativeEmbeddedImageReorderEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly direction: PdfNativeEmbeddedImageReorderIntent | "restore";
  /** True for a one-step move and for the restore of one. */
  readonly stepped: boolean;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly imageOrderAfter: number;
  readonly resourceName: string;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceOwner: Readonly<{ objectNumber: number; generation: number }>;
  readonly image: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly blockLength: number;
  readonly blockLeads: boolean;
  /** The transform the moved block is wrapped in so it keeps its place
   * while passing top-level `cm`s; absent when the move passes none. */
  readonly compensation?: readonly number[];
  /** The graphics states the moved block names again inside its wrapper so
   * the picture paints as it was drawn; absent when the drawing it passes
   * leaves its paint alone. */
  readonly carriedStates?: readonly string[];
  /** The `q…Q` groups and marked-content sections the moved block left,
   * carrying their state in its wrapper; absent when it left none. */
  readonly lift?: PdfNativeEmbeddedImageMoveLift;
  /** The places this move emptied of the markers of the sections that
   * travelled with the picture, and the bytes it left at the span's outer
   * edge; absent when every section it left was copied instead. */
  readonly emptiedMarkers?: readonly Readonly<{ start: number; end: number }>[];
  readonly stationaryBytes?: number;
}

export interface PdfNativeEmbeddedImageReorderPlanResult {
  readonly payload: PdfNativeEmbeddedImageReorderPayload;
  readonly evidence: PdfNativeEmbeddedImageReorderEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly plan: PdfNativeEmbeddedImageReorderPlan;
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
    fail(`Invalid native PDF picture ${label}.`);
  }
  return value;
}

export function readPdfNativeEmbeddedImageReorderPayload(
  value: unknown,
): PdfNativeEmbeddedImageReorderPayload {
  const payload = plainRecord(value);
  if (
    !payload ||
    payload.schemaVersion !== MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION
  ) {
    fail("Unsupported native PDF picture arrange schema.");
  }
  const selection = plainRecord(payload.selection);
  if (!selection) fail("Invalid native PDF picture selection.");
  if (
    payload.direction !== "front" &&
    payload.direction !== "back" &&
    payload.direction !== "forward" &&
    payload.direction !== "backward" &&
    payload.direction !== "restore"
  ) {
    fail("Invalid native PDF picture arrange direction.");
  }
  const base = Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
    selection: Object.freeze({
      pageNumber: wholeNumber(selection.pageNumber, "page", false),
      imageOrder: wholeNumber(selection.imageOrder, "order"),
    }),
    direction: payload.direction,
  });
  if (payload.leaveClips !== undefined && (payload.leaveClips !== true || payload.direction !== "front")) {
    fail("Invalid native PDF picture arrange clip move.");
  }
  if (payload.direction !== "restore") {
    if (payload.history !== undefined) {
      fail("Native PDF picture arrange carries history it cannot use.");
    }
    return payload.leaveClips ? Object.freeze({ ...base, leaveClips: true as const }) : base;
  }
  const history = plainRecord(payload.history);
  if (!history || typeof history.blockLeads !== "boolean") {
    fail("Invalid native PDF picture arrange history.");
  }
  if (history.stepped !== undefined && typeof history.stepped !== "boolean") {
    fail("Invalid native PDF picture arrange history.");
  }
  const compensation = history.compensation;
  if (
    compensation !== undefined &&
    (!Array.isArray(compensation) ||
      compensation.length !== 6 ||
      compensation.some((value) => typeof value !== "number" || !Number.isFinite(value)))
  ) {
    fail("Invalid native PDF picture arrange compensation.");
  }
  const carried = history.carriedStates;
  if (
    carried !== undefined &&
    (!Array.isArray(carried) ||
      carried.length === 0 ||
      carried.length > MRAI_NATIVE_EMBEDDED_IMAGE_CARRIED_PAINT_STATE_LIMIT ||
      carried.some(
        (text) => typeof text !== "string" || !/^\/[^\s/<>[\]()%]{1,127}\s+gs$/.test(text),
      ))
  ) {
    fail("Invalid native PDF picture arrange carried graphics state.");
  }
  const lift = history.lift === undefined ? undefined : plainRecord(history.lift);
  if (history.lift !== undefined) {
    if (!lift) fail("Invalid native PDF picture arrange lift.");
    const save = wholeNumber(lift.save, "lift");
    const marked = wholeNumber(lift.marked, "lift");
    if (save + marked === 0 || save > 64 || marked > 64) {
      fail("Invalid native PDF picture arrange lift.");
    }
  }
  const emptied = history.emptiedMarkers;
  if (emptied !== undefined) {
    if (
      !Array.isArray(emptied) ||
      emptied.length === 0 ||
      emptied.length > 128 ||
      emptied.some((range) => {
        const plain = plainRecord(range);
        return (
          !plain ||
          typeof plain.start !== "number" ||
          typeof plain.end !== "number" ||
          !Number.isInteger(plain.start) ||
          !Number.isInteger(plain.end) ||
          plain.start < 0 ||
          plain.end <= plain.start
        );
      })
    ) {
      fail("Invalid native PDF picture arrange section markers.");
    }
    wholeNumber(history.stationaryBytes, "stationary bytes");
  } else if (history.stationaryBytes !== undefined) {
    fail("Invalid native PDF picture arrange section markers.");
  }
  return Object.freeze({
    ...base,
    history: Object.freeze({
      imageOrderAfter: wholeNumber(history.imageOrderAfter, "order"),
      spanStart: wholeNumber(history.spanStart, "span start"),
      spanEnd: wholeNumber(history.spanEnd, "span end", false),
      blockLength: wholeNumber(history.blockLength, "block length", false),
      blockLeads: history.blockLeads,
      ...(history.stepped === true ? { stepped: true } : {}),
      ...(emptied !== undefined
        ? {
            emptiedMarkers: Object.freeze(
              (emptied as { start: number; end: number }[]).map((range) =>
                Object.freeze({ start: range.start, end: range.end }),
              ),
            ),
            stationaryBytes: history.stationaryBytes as number,
          }
        : {}),
      ...(compensation !== undefined
        ? { compensation: Object.freeze([...(compensation as number[])]) }
        : {}),
      ...(carried !== undefined
        ? { carriedStates: Object.freeze([...(carried as string[])]) }
        : {}),
      ...(lift
        ? { lift: Object.freeze({ save: lift.save as number, marked: lift.marked as number }) }
        : {}),
    }),
  });
}

export async function planPdfNativeEmbeddedImageReorder(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeEmbeddedImageReorderPayload;
}): Promise<PdfNativeEmbeddedImageReorderPlanResult> {
  const payload = readPdfNativeEmbeddedImageReorderPayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  const plan =
    payload.direction === "restore"
      ? await planMraiNativeEmbeddedImageReorderRestore({
          sourceBytes: input.sourceBytes,
          pageNumber: payload.selection.pageNumber,
          imageOrder: payload.selection.imageOrder,
          imageOrderAfter: payload.history!.imageOrderAfter,
          spanStart: payload.history!.spanStart,
          spanEnd: payload.history!.spanEnd,
          blockLength: payload.history!.blockLength,
          blockLeads: payload.history!.blockLeads,
          compensation: payload.history!.compensation ?? null,
          carriedStates: payload.history!.carriedStates ?? null,
          lift: payload.history!.lift ?? null,
          emptiedMarkers: payload.history!.emptiedMarkers ?? null,
          stationaryBytes: payload.history!.stationaryBytes ?? null,
          stepped: payload.history!.stepped ?? false,
        })
      : await planMraiNativeEmbeddedImageReorder({
          sourceBytes: input.sourceBytes,
          pageNumber: payload.selection.pageNumber,
          imageOrder: payload.selection.imageOrder,
          direction: payload.direction,
          ...(payload.leaveClips ? { leaveClips: true } : {}),
        });
  const prepared = await prepareMraiNativeEmbeddedImageReorder(
    input.sourceBytes,
    plan,
  );
  await verifyMraiNativeEmbeddedImageReorderCandidate(
    input.sourceBytes,
    plan,
    prepared,
  );
  const evidence: PdfNativeEmbeddedImageReorderEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
    sourceSha256,
    direction: plan.direction,
    stepped: plan.stepped,
    pageNumber: plan.pageNumber,
    imageOrder: plan.imageOrder,
    imageOrderAfter: plan.imageOrderAfter,
    resourceName: plan.resourceName,
    page: Object.freeze({ ...plan.page }),
    resourceOwner: Object.freeze({ ...plan.resourceOwner }),
    image: Object.freeze({ ...plan.image }),
    contentStream: Object.freeze({ ...plan.contentStream }),
    operatorSpan: Object.freeze({ ...plan.operatorSpan }),
    blockLength: plan.blockLength,
    blockLeads: plan.blockLeads,
    ...(plan.compensation ? { compensation: Object.freeze([...plan.compensation]) } : {}),
    ...(plan.carriedStates && plan.carriedStates.length > 0
      ? { carriedStates: Object.freeze([...plan.carriedStates]) }
      : {}),
    ...(plan.lift ? { lift: Object.freeze({ ...plan.lift }) } : {}),
    ...(plan.emptiedMarkers && plan.emptiedMarkers.length > 0
      ? {
          emptiedMarkers: Object.freeze(
            plan.emptiedMarkers.map((range) => Object.freeze({ ...range })),
          ),
          stationaryBytes: plan.stationaryBytes ?? 0,
        }
      : {}),
  });
  return Object.freeze({ payload, evidence, prepared, plan });
}

