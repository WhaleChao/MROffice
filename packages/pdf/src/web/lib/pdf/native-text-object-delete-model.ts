/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeTextObjectDeleteEvidence, PdfNativeTextObjectDeleteHistory } from "./native-text-object-delete";

export function fail(message: string): never {
  throw new Error(message);
}

export function offset(value: unknown, label: string, allowZero = true): number {
  if (
    typeof value !== "number" ||
    !Number.isSafeInteger(value) ||
    value < (allowZero ? 0 : 1)
  ) {
    fail(`Invalid native PDF text object ${label}.`);
  }
  return value;
}

/** What an Undo needs to write the deleted object back where it was. */
export function pdfNativeTextObjectDeleteHistory(
  evidence: PdfNativeTextObjectDeleteEvidence,
): PdfNativeTextObjectDeleteHistory {
  if (evidence.direction !== "delete") {
    fail("Native PDF text object delete cannot be undone twice.");
  }
  return Object.freeze({
    offset: evidence.operatorSpan.start,
    objectBytes: evidence.objectBytes,
    scope: evidence.scope,
    ...(evidence.penRestore ? { penRestore: evidence.penRestore.number } : {}),
    ...(evidence.penRestore?.precedingShows?.length
      ? {
          penRestores: Object.freeze([
            ...evidence.penRestore.precedingShows.map((show) => show.number),
            evidence.penRestore.number,
          ]),
        }
      : {}),
    ...(evidence.penRestore &&
    [...(evidence.penRestore.precedingShows ?? []), evidence.penRestore].some((pen) => pen.characterSpacing)
      ? {
          penCharacterSpacings: Object.freeze(
            [...(evidence.penRestore.precedingShows ?? []), evidence.penRestore].map((pen) => pen.characterSpacing ?? 0),
          ),
        }
      : {}),
  });
}
