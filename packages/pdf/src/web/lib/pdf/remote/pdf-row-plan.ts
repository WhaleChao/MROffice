/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiPost } from "@mrnote/core/web/lib/api.ts";
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { sha256PdfBytes } from "../engine/native/transaction-error";

/** The revision a page's annotation rows make on the working copy's head,
 * as the server writes it (engine unification, stage 4; the API's
 * pdf_row_plan): the head with the revision appended, its digest, and the
 * server's signature that it is its writing of exactly these rows on exactly
 * that head (the bytes record carries it as `rowsBinding`). */
export interface PdfRowPlanResult {
  readonly outputBytes: Uint8Array;
  readonly outputSha256: string;
  readonly rowsBinding: string;
  /** What the page takes beside the edit's own companions (its labels
   * renumbered with its pages). */
  readonly writtenCompanions?: Readonly<Record<string, JsonValue>>;
}

interface PdfRowPlanWire {
  output_sha256: string;
  byte_length: number;
  source_length: number;
  tail_base64: string;
  rows_binding: string;
  written_companions?: Record<string, JsonValue>;
}

function base64Bytes(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

/** Plan `rows` on `headBytes` (which holds `previousRows`). The head must be
 * the page's saved version or a head the server holds as the editor's
 * candidate (an engine call sends it first). */
export async function planPdfRowsOnWorkingCopy(input: Readonly<{
  pageId: string;
  headBytes: Uint8Array;
  headSha256: string;
  savedSource: Readonly<{ resourceId: string; resourceVersionId: string; sourceSha256: string }>;
  previousRows: readonly JsonValue[];
  rows: readonly JsonValue[];
  /** Page collections the edit changes beside its rows (the ink strokes an
   * ink edit's rows are written from, form values, the document's
   * decorations as one value): the page takes them with the rows. */
  companions?: Readonly<Record<string, JsonValue>>;
  /** Those collections as the head holds them: a form value is written only
   * where it changed. */
  previousCompanions?: Readonly<Record<string, JsonValue>>;
  /** The structure the head holds (pages deleted, moved, added): rows are
   * placed by it. */
  headStructure?: Readonly<Record<string, JsonValue>>;
}>): Promise<PdfRowPlanResult> {
  const answer = await apiPost<PdfRowPlanWire>(`/api/v1/pages/${encodeURIComponent(input.pageId)}/pdf-row-plan`, {
    head_sha256: input.headSha256,
    ...(input.headSha256 === input.savedSource.sourceSha256
      ? { resource_id: input.savedSource.resourceId, resource_version_id: input.savedSource.resourceVersionId }
      : {}),
    previous_annotations: input.previousRows,
    annotations: input.rows,
    ...(input.headStructure ? { head_structure: input.headStructure } : {}),
    ...(input.companions && Object.keys(input.companions).length ? { companions: input.companions,
      ...(input.previousCompanions ? { previous_companions: input.previousCompanions } : {}) } : {}),
  });
  const tail = base64Bytes(answer.tail_base64);
  if (answer.source_length !== input.headBytes.length || answer.byte_length !== input.headBytes.length + tail.length) {
    throw new Error("The planned annotation revision does not append to the working copy.");
  }
  const outputBytes = new Uint8Array(answer.byte_length);
  outputBytes.set(input.headBytes, 0);
  outputBytes.set(tail, input.headBytes.length);
  if (await sha256PdfBytes(outputBytes) !== answer.output_sha256) {
    outputBytes.fill(0);
    throw new Error("The planned annotation revision is not the one the server named.");
  }
  const written = answer.written_companions;
  return Object.freeze({ outputBytes, outputSha256: answer.output_sha256, rowsBinding: answer.rows_binding,
    ...(written && typeof written === "object" && !Array.isArray(written) && Object.keys(written).length
      ? { writtenCompanions: Object.freeze({ ...written }) } : {}) });
}
