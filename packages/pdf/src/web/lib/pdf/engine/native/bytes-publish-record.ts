/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import type { PdfEngineSourceIdentity } from "../contracts";
import type { PdfSharedStateRecord } from "../shared-state-bridge";
import {
  MRAI_NATIVE_BYTES_PUBLISH_ACTION,
  MRAI_NATIVE_BYTES_PUBLISH_KIND,
  MRAI_NATIVE_BYTES_PUBLISH_SCHEMA_VERSION,
} from "./vocabulary";

/** Who the record says made the edit (the feature and control its history
 * and notices name); the server reads neither. */
export const PDF_NATIVE_BYTES_PUBLISH_FEATURE_ID = "PDF-EDT-ENGINE-01" as const;
export const PDF_NATIVE_BYTES_PUBLISH_CONTROL_ID = "PDFCTL.RIBBON.EDIT.EDIT_CONTENT" as const;

/** The journal record of one revision the engine appended, published as its
 * bytes (engine unification, stage 2): the version it appends to and its
 * depth, and the file it makes. The bytes themselves reach the server as the
 * engine host's candidate for the page (the editor holds them there first).
 * Mirror of NativeBytesPublishRecord (native_surgical_contract.py). */
export function buildMraiNativeBytesPublishRecord(input: Readonly<{
  id: string;
  source: PdfEngineSourceIdentity;
  /** Cross-reference sections of the source (the saved version). */
  sourceXrefSections: number;
  sourceLength: number;
  output: Readonly<{ sha256: string; byteLength: number }>;
  baseSessionRevision: number;
  selectionRevision?: number;
  label?: string;
  /** The server's signature that this revision is its writing of the
   * page's annotation rows the same save carries (stage 4, pdf-row-plan). */
  rowsBinding?: string;
  /** Where it moved a run of the page's own text (its plan's
   * textFidelity.propertyMutation): a PDF line has no paragraph indent, so
   * the indent is read back from the journal (native-indent-chain). */
  propertyMutation?: JsonValue;
  /** The text blocks the edit set, the box it wrote first, then its page's
   * box map (native-page-box-map.ts): a page published as its bytes is
   * read again with them, each box the box it was. */
  keptTextBlocks?: readonly Readonly<{
    pageNumber: number; rows: readonly string[]; origin?: Readonly<{ x: number; y: number }>;
    frame?: Readonly<{ left: number; right: number; firstLeft: number }>; moved?: true; pinned?: true;
  }>[];
}>): PdfSharedStateRecord {
  if (!input.id || input.id.length > 200 || !/^[0-9a-f]{64}$/.test(input.output.sha256) ||
    input.output.sha256 === input.source.sourceSha256 || !Number.isSafeInteger(input.sourceLength) ||
    input.sourceLength < 1 || !Number.isSafeInteger(input.output.byteLength) || input.output.byteLength <= input.sourceLength ||
    !Number.isSafeInteger(input.sourceXrefSections) || input.sourceXrefSections < 1 || input.sourceXrefSections > 31 ||
    !Number.isSafeInteger(input.baseSessionRevision) || input.baseSessionRevision < 0) {
    throw new Error("The published PDF revision is not one revision appended to its source.");
  }
  const label = input.label?.slice(0, 200);
  const record: Record<string, JsonValue> = {
    id: input.id,
    schemaVersion: MRAI_NATIVE_BYTES_PUBLISH_SCHEMA_VERSION,
    kind: MRAI_NATIVE_BYTES_PUBLISH_KIND,
    action: MRAI_NATIVE_BYTES_PUBLISH_ACTION,
    transactionId: input.id,
    featureId: PDF_NATIVE_BYTES_PUBLISH_FEATURE_ID,
    controlId: PDF_NATIVE_BYTES_PUBLISH_CONTROL_ID,
    baseSessionRevision: input.baseSessionRevision,
    selectionRevision: Math.max(0, input.selectionRevision ?? 0),
    source: { resourceId: input.source.resourceId, resourceVersionId: input.source.resourceVersionId,
      sha256: input.source.sourceSha256 },
    sourceCheckpoint: { schemaVersion: "mrai.pdf-native-source-checkpoint.v1", xrefSections: input.sourceXrefSections },
    derivedCandidate: { sha256: input.output.sha256, byteLength: input.output.byteLength,
      sourceLength: input.sourceLength },
    ...(label ? { label } : {}),
    ...(input.rowsBinding && /^[0-9a-f]{64}$/.test(input.rowsBinding) ? { rowsBinding: input.rowsBinding } : {}),
    ...(input.propertyMutation && typeof input.propertyMutation === "object" && !Array.isArray(input.propertyMutation) &&
      JSON.stringify(input.propertyMutation).length <= 4096 ? { propertyMutation: input.propertyMutation } : {}),
    ...keptTextBlocksField(input.keptTextBlocks),
  };
  return Object.freeze(record) as unknown as PdfSharedStateRecord;
}

/** The most blocks, and bytes of them, one record carries (mirror of
 * NativeBytesPublishRecord.keptTextBlocks): a page past them is published
 * with the box its edit wrote alone. */
const KEPT_TEXT_BLOCKS_MAX = 1_000;
const KEPT_TEXT_BLOCKS_MAX_BYTES = 256 * 1024;

function keptTextBlocksField(blocks: Parameters<typeof buildMraiNativeBytesPublishRecord>[0]["keptTextBlocks"]):
  Record<string, JsonValue> {
  const valid = (blocks ?? []).filter((block) => Number.isSafeInteger(block.pageNumber) && block.pageNumber >= 1 &&
    block.rows.length > 0 && block.rows.every((row) => typeof row === "string" && row.trim()) &&
    (!block.origin || (Number.isFinite(block.origin.x) && Number.isFinite(block.origin.y))));
  const json = (list: typeof valid) => list.map((block) => ({
    pageNumber: block.pageNumber, rows: [...block.rows],
    ...(block.origin ? { origin: { x: block.origin.x, y: block.origin.y } } : {}),
    ...(block.frame ? { frame: { left: block.frame.left, right: block.frame.right, firstLeft: block.frame.firstLeft } } : {}),
    ...(block.moved ? { moved: true } : {}),
    ...(block.pinned ? { pinned: true } : {}),
  })) as JsonValue[];
  for (const list of [valid.slice(0, KEPT_TEXT_BLOCKS_MAX), valid.slice(0, 1)]) {
    const value = json(list);
    if (value.length && JSON.stringify(value).length <= KEPT_TEXT_BLOCKS_MAX_BYTES) return { keptTextBlocks: value };
  }
  return {};
}

/** A journal entry that is a published revision. */
export function isMraiNativeBytesPublishRecord(value: unknown): value is PdfSharedStateRecord & Readonly<{
  derivedCandidate: Readonly<{ sha256: string; byteLength: number; sourceLength: number }>;
}> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  const output = record.derivedCandidate as Record<string, unknown> | undefined;
  return record.schemaVersion === MRAI_NATIVE_BYTES_PUBLISH_SCHEMA_VERSION &&
    record.kind === MRAI_NATIVE_BYTES_PUBLISH_KIND && !!output && typeof output === "object" &&
    typeof output.sha256 === "string" && Number.isSafeInteger(output.byteLength) &&
    Number.isSafeInteger(output.sourceLength);
}
