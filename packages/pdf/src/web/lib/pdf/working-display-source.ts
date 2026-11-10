/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MAX_MRAI_NATIVE_SOURCE_BYTES, sha256PdfBytes } from "./engine/native/transaction-error";

/** Track requested displays, not only completed PDF.js documents: a second
 * load may supersede the first while the completed-document ref is empty. */
export function createPdfWorkingViewContinuity() {
  let previous: {dataItemId: string | undefined; working: boolean} | null = null;
  return {
    begin(dataItemId: string | undefined, working: boolean): boolean {
      const retain = Boolean(dataItemId && previous?.dataItemId === dataItemId && (working || previous.working));
      previous = {dataItemId,working};
      return retain;
    },
  };
}

/** A local display identity, never a ResourceVersion or saved acknowledgement.
 * The byte owner may stop serving a retired generation; readers must not fall
 * back to the server URL when that happens. */
export interface PdfWorkingDisplaySource {
  readonly kind: "local-working-pdf";
  readonly pageId: string;
  readonly dataItemId: string;
  readonly sessionId: string;
  readonly workingRevision: number;
  readonly sha256: string;
  readonly displayKey: string;
  /** Intent continuity only; never a resource-version or write receipt. */
  readonly originDocumentRevision?: string | number;
  /** The saved file these bytes extend (when the display was made): a page
   * structure written in that file applies to them, one written later does
   * not (lib/pdf/structured-document). */
  readonly baseSha256?: string;
  copyBytes(): Uint8Array | null;
}

export async function copyPdfWorkingDisplayBytes(source: PdfWorkingDisplaySource, dataItemId?: string): Promise<Uint8Array> {
  if (!isPdfWorkingDisplaySourceFor(source, dataItemId)) throw new Error("PDF local display scope is invalid.");
  const bytes = source.copyBytes();
  if (!bytes) throw new Error("PDF local display generation is retired.");
  try {
    if (!bytes.length || bytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES || await sha256PdfBytes(bytes) !== source.sha256) {
      throw new Error("PDF local display bytes do not match their identity.");
    }
    return bytes;
  } catch (error) { bytes.fill(0); throw error; }
}

/** A working revision's bytes for a reader that holds them to the
 * revision's digest itself (the source inspection, which throws on a
 * mismatch before anything reads them): hashed here as well, every
 * revision was hashed twice (~200 ms each in Safari). */
export function copyPdfWorkingDisplayBytesForInspection(source: PdfWorkingDisplaySource, dataItemId?: string): Uint8Array {
  if (!isPdfWorkingDisplaySourceFor(source, dataItemId)) throw new Error("PDF local display scope is invalid.");
  const bytes = source.copyBytes();
  if (!bytes) throw new Error("PDF local display generation is retired.");
  if (!bytes.length || bytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES) {
    bytes.fill(0);
    throw new Error("PDF local display bytes do not match their identity.");
  }
  return bytes;
}

export function isPdfWorkingDisplaySourceFor(
  source: PdfWorkingDisplaySource | null | undefined,
  dataItemId: string | null | undefined,
): source is PdfWorkingDisplaySource {
  return Boolean(source && source.kind === "local-working-pdf" &&
    source.dataItemId === dataItemId && dataItemId && source.pageId &&
    source.sessionId && Number.isSafeInteger(source.workingRevision) && source.workingRevision >= 0 &&
    /^[a-f0-9]{64}$/.test(source.sha256) &&
    source.displayKey === `mrai-pdf-working:${source.sessionId}:${source.workingRevision}` &&
    typeof source.copyBytes === "function");
}
