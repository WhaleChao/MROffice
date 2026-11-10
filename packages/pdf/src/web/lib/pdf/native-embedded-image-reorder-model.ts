/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeEmbeddedImageReorderEvidence, PdfNativeEmbeddedImageReorderHistory } from "./native-embedded-image-reorder";

export function fail(message: string): never {
  throw new Error(message);
}

/**
 * The history an Undo needs to rotate a forward move's span back. A Redo just
 * plans the same move again, which re-proves the block against current bytes.
 */
export function pdfNativeEmbeddedImageReorderHistory(
  evidence: PdfNativeEmbeddedImageReorderEvidence,
): PdfNativeEmbeddedImageReorderHistory {
  if (evidence.direction === "restore") {
    fail("Native PDF picture arrange cannot be undone twice.");
  }
  return Object.freeze({
    imageOrderAfter: evidence.imageOrder,
    spanStart: evidence.operatorSpan.start,
    spanEnd: evidence.operatorSpan.end,
    blockLength: evidence.blockLength,
    blockLeads: evidence.blockLeads,
    ...(evidence.stepped ? { stepped: true } : {}),
    ...(evidence.compensation ? { compensation: evidence.compensation } : {}),
    ...(evidence.carriedStates && evidence.carriedStates.length > 0
      ? { carriedStates: evidence.carriedStates }
      : {}),
    ...(evidence.lift ? { lift: evidence.lift } : {}),
    ...(evidence.emptiedMarkers && evidence.emptiedMarkers.length > 0
      ? {
          emptiedMarkers: evidence.emptiedMarkers,
          stationaryBytes: evidence.stationaryBytes ?? 0,
        }
      : {}),
  });
}
