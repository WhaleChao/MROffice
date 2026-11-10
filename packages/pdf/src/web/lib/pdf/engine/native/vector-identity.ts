/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A proof's id: the page, the stream and the span of its operators, and --
 * for a Form the page shows more than once -- which drawing it reads (mirror
 * of the server's NativeVectorWriteSelection id). */
export function pdfNativeVectorProofId(input: {
  readonly family: "shape" | "subpath";
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly objectNumber: number;
  readonly generation: number;
  readonly start: number;
  readonly end: number;
  readonly formInstance?: string;
}): string {
  return `pdf-native-vector${input.family === "subpath" ? "-subpath" : ""}:${input.sourceSha256}:p${input.pageNumber}:${input.objectNumber}:${input.generation}:${input.start}:${input.end}${
    input.formInstance !== undefined ? `:i${input.formInstance}` : ""
  }`;
}

/** A drawing key as a proof carries it: the offsets of the `Do`s on its way. */
export const PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN = /^\d{1,9}(?:\.\d{1,9}){0,7}$/;
