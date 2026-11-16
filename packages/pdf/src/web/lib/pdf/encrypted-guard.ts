/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Fail-closed loading for every pdf-lib mutation path.
 *
 * pdf-lib does not decrypt. `ignoreEncryption: true` parses an encrypted
 * document's object graph while its stream contents stay ciphertext — and a
 * later `save()` copies that ciphertext into a file whose trailer no longer
 * declares encryption. Every reader shows blank or garbage pages, and no step
 * ever errors: the classic "pretend the decryption worked" failure.
 *
 * The dangerous case is not the password-prompt flow (those bytes arrive
 * already decrypted by the server, and `isEncrypted` is false). It is the
 * PERMISSIONS-ONLY document — encrypted with an empty user password, so
 * pdf.js renders it without ever prompting, the decrypted-bytes cache is
 * never seeded, and the raw encrypted bytes flow straight into compress,
 * merge, flatten or form export. Before this guard, those paths silently
 * destroyed the file. Refusing is the only honest behaviour pdf-lib affords.
 */

import { PDFDocument } from "pdf-lib";
import { pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

/** Thrown when a mutation path receives encrypted bytes. Matched by name so
 *  it survives minification and realm boundaries, like the password errors. */
export class PdfEncryptedError extends Error {
  constructor() {
    super(
      "This PDF is encrypted; editing it would corrupt the saved file. Remove the protection first.",
    );
    this.name = "PdfEncryptedError";
  }
}

export function isPdfEncryptedError(err: unknown): boolean {
  return (
    err instanceof Error &&
    (err.name === "PdfEncryptedError" ||
      // pdf-lib's own refusal when ignoreEncryption is NOT passed.
      err.name === "EncryptedPDFError")
  );
}

type LoadOptions = NonNullable<Parameters<typeof PDFDocument.load>[1]>;

/**
 * Load a document for MUTATION. Encrypted bytes are refused with a typed
 * error instead of being silently mangled; everything else behaves exactly
 * like `PDFDocument.load(bytes, { ignoreEncryption: true, ...options })`,
 * which is what every mutation path used before.
 *
 * `ignoreEncryption` stays on for the LOAD so detection is ours: without it,
 * pdf-lib throws its own error before `isEncrypted` can be read, and the
 * message is not one we would show a user.
 */
export async function loadPdfForMutation(
  bytes: Uint8Array | ArrayBuffer,
  options: Omit<LoadOptions, "ignoreEncryption"> = {},
): Promise<PDFDocument> {
  const doc = await PDFDocument.load(bytes, {
    parseSpeed: pdfLibParseSpeed(),
    ...options,
    ignoreEncryption: true,
  });
  if (doc.isEncrypted) throw new PdfEncryptedError();
  return doc;
}

/**
 * Load a document for READ-ONLY structural inspection (walking annotations,
 * reading the calculation order). Ciphertext streams make string and name
 * primitives unreadable too, so an encrypted document yields nothing useful —
 * but refusing a pure read would break surfaces that degrade gracefully.
 * Returns null for encrypted bytes so callers show their honest empty state.
 */
export async function loadPdfForInspection(
  bytes: Uint8Array | ArrayBuffer,
  options: Omit<LoadOptions, "ignoreEncryption"> = {},
): Promise<PDFDocument | null> {
  const doc = await PDFDocument.load(bytes, {
    parseSpeed: pdfLibParseSpeed(),
    ...options,
    ignoreEncryption: true,
  });
  return doc.isEncrypted ? null : doc;
}
