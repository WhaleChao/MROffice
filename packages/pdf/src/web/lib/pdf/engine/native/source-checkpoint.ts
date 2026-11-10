/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfNativeCheckpointChainSha256 } from "./source-checkpoint-chain";
export { pdfNativeCheckpointChainSha256 };
import {
  PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA,
} from "./vocabulary";
export {
  PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA,
};
import type { PdfEngineSourceIdentity } from "../contracts";
import {
  readMraiNativePdfXrefSectionCount,
  sha256PdfBytes,
} from "./content-stream-transaction";

export type PdfNativeSourceCheckpoint = Readonly<{
  source: PdfEngineSourceIdentity;
  xrefSections: number;
}>;
const issued = new WeakSet<object>();

/** Local source evidence, not server seed authority. The API independently
 * authenticates custody and checks its own baseline and actual xref depth. */
export async function preparePdfNativeSourceCheckpoint(
  source: PdfEngineSourceIdentity,
  bytes: Uint8Array,
): Promise<PdfNativeSourceCheckpoint> {
  if (!source.resourceId || !source.resourceVersionId ||
    await sha256PdfBytes(bytes) !== source.sourceSha256) throw new Error("The PDF checkpoint source changed.");
  const xrefSections = readMraiNativePdfXrefSectionCount(bytes);
  if (xrefSections < 1 || xrefSections >= 32) throw new Error("This PDF has reached its supported revision limit.");
  const checkpoint = Object.freeze({ source: Object.freeze({ ...source }), xrefSections });
  issued.add(checkpoint);
  return checkpoint;
}

export function pdfNativeSourceCheckpointClaim(
  checkpoint: PdfNativeSourceCheckpoint | undefined,
  source: PdfEngineSourceIdentity, xrefSections: number,
): Readonly<{schemaVersion: typeof PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA; xrefSections: number}> | null {
  if (!checkpoint || !issued.has(checkpoint) || checkpoint.source.resourceId !== source.resourceId ||
    checkpoint.source.resourceVersionId !== source.resourceVersionId || checkpoint.source.sourceSha256 !== source.sourceSha256 ||
    checkpoint.xrefSections !== xrefSections) return null;
  return Object.freeze({ schemaVersion: PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA, xrefSections });
}

/**
 * The source checkpoint the window issues before a native write, for a write
 * the server plans (the AI's PDF commands): bound to the depth the page save
 * will authenticate, so a source another tool already revised is written too.
 */
export async function pdfNativeServerSourceCheckpoint(
  request: Readonly<{ source: PdfEngineSourceIdentity; checkpointDepth: number }>,
  sourceBytes: Uint8Array,
): Promise<PdfNativeSourceCheckpoint> {
  const checkpoint = await preparePdfNativeSourceCheckpoint(request.source, sourceBytes);
  if (checkpoint.xrefSections !== request.checkpointDepth) {
    throw new Error("The PDF source revisions differ from the page's authenticated depth.");
  }
  return checkpoint;
}


