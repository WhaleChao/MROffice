/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { sha256PdfBytes } from "./engine/native/transaction-error";
import { z } from "zod";
import type { PdfTaggedTextHistoryOrigin, PdfTaggedTextOperation } from "./native-tagged-text-roundtrip";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const PDF_TAGGED_TEXT_SCHEMA = "mrai.pdf-native-tagged-text.v1" as const;

export const shaSchema = z.string().regex(/^[a-f0-9]{64}$/);

export const refSchema = z
  .object({
    objectNumber: z.number().int().min(1),
    generation: z.number().int().min(0).max(65535),
  })
  .strict();

export const valueSchema = z
  .object({
    text: z.string(),
    rawHex: z
      .string()
      .regex(/^(?:[0-9a-f]{2})*$/)
      .max(32000),
  })
  .strict();

export const inlineSchema = valueSchema
  .extend({ start: z.number().int().min(0), end: z.number().int().min(1) })
  .strict();

export const targetSchema = z
  .object({
    pageNumber: z.number().int().min(1),
    page: refSchema,
    node: refSchema,
    stream: refSchema,
    parentTreeKey: z.number().int().min(0),
    mcid: z.number().int().min(0).max(1_000_000_000),
    role: z.string().min(1).max(256),
    nodeSha256: shaSchema,
    streamSha256: shaSchema,
    alt: valueSchema.nullable(),
    actualText: valueSchema.nullable(),
    inlineAlt: inlineSchema.nullable(),
    inlineActualText: inlineSchema.nullable(),
  })
  .strict();

export const scalarText = z
  .string()
  .max(4000)
  .refine(
    (value) =>
      !Array.from(value).some((char) => {
        const code = char.codePointAt(0)!;
        return (
          (code >= 0xd800 && code <= 0xdfff) ||
          (code < 32 && !["\t", "\n", "\r"].includes(char))
        );
      }),
  );

export const propertySchema = z.enum(["Alt", "ActualText"]);

export const originSchema = z
  .object({
    sourceLength: z
      .number()
      .int()
      .positive()
      .max(PDF_EDITABLE_MAX_BYTES),
    candidateLength: z
      .number()
      .int()
      .positive()
      .max(PDF_EDITABLE_MAX_BYTES),
    sourceSha256: shaSchema,
    candidateSha256: shaSchema,
    target: targetSchema,
    property: propertySchema,
    value: scalarText,
  })
  .strict();

export const historySchema = z
  .object({ direction: z.enum(["undo", "redo"]), origin: originSchema })
  .strict();

export const patchSchema = z
  .object({
    schemaVersion: z.literal(PDF_TAGGED_TEXT_SCHEMA),
    id: z.string().min(1).max(128),
    sourceSha256: shaSchema,
    target: targetSchema,
    property: propertySchema,
    value: scalarText,
    history: historySchema.nullable().optional(),
  })
  .strict();

export const operationSchema = z
  .object({
    sourceResourceId: z.string().min(1).max(128),
    sourceResourceVersionId: z.string().min(1).max(128),
    patch: patchSchema,
  })
  .strict();

export function readPdfTaggedTextOperations(
  content: unknown,
): readonly PdfTaggedTextOperation[] {
  if (!content || typeof content !== "object" || Array.isArray(content))
    throw Error("PDF content cannot be read safely.");
  const raw = (content as Record<string, unknown>).pdfTaggedTextOperations;
  const parsed = z
    .array(operationSchema)
    .max(64)
    .safeParse(raw === undefined ? [] : raw);
  if (
    !parsed.success ||
    new Set(parsed.data.map((op) => op.patch.id)).size !== parsed.data.length
  )
    throw Error("PDF semantic-text history cannot be read safely.");
  return parsed.data;
}

export async function createPdfTaggedTextHistoryOrigin(
  operation: PdfTaggedTextOperation,
  before: Uint8Array,
  after: Uint8Array,
  outputSha256: string,
): Promise<PdfTaggedTextHistoryOrigin> {
  if (
    after.length <= before.length ||
    !before.every((value, index) => after[index] === value) ||
    (await sha256PdfBytes(before)) !== operation.patch.sourceSha256 ||
    (await sha256PdfBytes(after)) !== outputSha256
  )
    throw Error("PDF semantic-text save did not preserve its source receipt.");
  return originSchema.parse({
    sourceLength: before.length,
    candidateLength: after.length,
    sourceSha256: operation.patch.sourceSha256,
    candidateSha256: outputSha256,
    target: operation.patch.target,
    property: operation.patch.property,
    value: operation.patch.value,
  });
}
