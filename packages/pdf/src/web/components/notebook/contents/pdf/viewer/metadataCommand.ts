/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfMetadata } from "../../../../../lib/pdf/document-properties";

export const PDF_METADATA_COMMAND_KEYS = [
  "title",
  "author",
  "subject",
  "keywords",
  "language",
] as const satisfies readonly (keyof PdfMetadata)[];

export function snapshotPdfMetadataCommandPayload(
  value: PdfMetadata,
): Readonly<PdfMetadata> | null {
  try {
    if (!value || typeof value !== "object") return null;
    const descriptors = Object.getOwnPropertyDescriptors(value);
    const keys = Reflect.ownKeys(value);
    if (
      keys.length !== PDF_METADATA_COMMAND_KEYS.length ||
      keys.some(
        (key) =>
          typeof key !== "string" ||
          !PDF_METADATA_COMMAND_KEYS.includes(key as keyof PdfMetadata),
      )
    ) {
      return null;
    }
    const snapshot = {} as PdfMetadata;
    for (const key of PDF_METADATA_COMMAND_KEYS) {
      const descriptor = descriptors[key];
      if (
        !descriptor ||
        !("value" in descriptor) ||
        !descriptor.enumerable ||
        typeof descriptor.value !== "string"
      ) {
        return null;
      }
      snapshot[key] = descriptor.value;
    }
    return Object.freeze(snapshot);
  } catch {
    return null;
  }
}
