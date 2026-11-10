/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Stable shared-state authority for native PDF page-import mutations.
 *
 * Imported PDF bytes never live in `content_json`. The source receipt names
 * the workspace-scoped DataItem and its immutable public identity. The server
 * must independently resolve DataItem -> Dataset -> Notebook -> Workspace,
 * lock the authoritative storage object, and verify its size/hash before it
 * replays the ordered page import against the target source receipt.
 */

import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";
export const PDF_NATIVE_PAGE_IMPORTS_COLLECTION =
  "pdfNativePageImports" as const;
export const PDF_NATIVE_PAGE_IMPORT_SCHEMA_VERSION = 1 as const;
export const PDF_NATIVE_PAGE_REPLACE_SCHEMA_VERSION = 2 as const;
export const PDF_NATIVE_PAGE_REPLACE_RANGE_SCHEMA_VERSION = 3 as const;
export const MAX_PDF_NATIVE_PAGE_IMPORT_RECORDS = 64;
export const MAX_PDF_NATIVE_PAGE_IMPORT_PAGES = 256;
export const MAX_PDF_NATIVE_PAGE_IMPORT_SOURCE_BYTES = PDF_EDITABLE_MAX_BYTES;
export const MAX_PDF_NATIVE_PAGE_IMPORT_AGGREGATE_SOURCE_BYTES =
  256 * 1024 * 1024;
export const MAX_PDF_NATIVE_PAGE_INDEX = 9_999;
export const MAX_PDF_NATIVE_TARGET_INSERTION_INDEX = 10_000;

export interface PdfNativePageImportSource {
  dataItemId: string;
  datasetId: string;
  workspaceId: string;
  notebookId: string;
  sha256: string;
  sizeBytes: number;
}

export interface PdfNativePageImportTargetSourceReceipt {
  resourceId: string;
  resourceVersionId: string;
  sha256: string;
}

/**
 * Delete is a full tombstone. Source custody, selected pages, insertion point,
 * and target resource identity stay immutable so delete/Undo cannot silently
 * retarget a different PDF or source file. Only operation and the target
 * version/SHA receipt advance between transitions.
 */
export interface PdfNativePageImportRecord {
  /**
   * v1 inserts pages at `targetInsertionIndex`; v2 replaces one page with the
   * source page's presentation; v3 pairs ordered source/target ranges while
   * preserving every target page's stable identity, boxes, and rotation.
   */
  schemaVersion:
    | typeof PDF_NATIVE_PAGE_IMPORT_SCHEMA_VERSION
    | typeof PDF_NATIVE_PAGE_REPLACE_SCHEMA_VERSION
    | typeof PDF_NATIVE_PAGE_REPLACE_RANGE_SCHEMA_VERSION;
  /** `settled`: the pages stay as ordinary pages of the document, their
   * import markers gone (before a page-structure edit); no later state. */
  operation: "upsert" | "delete" | "settled";
  id: string;
  source: PdfNativePageImportSource;
  targetSourceReceipt: PdfNativePageImportTargetSourceReceipt;
  /** Zero-based, strictly increasing source-page indices. */
  sourcePageIndices: number[];
  /** Zero-based gap in the target page tree; pageCount means append. */
  targetInsertionIndex: number;
  /** v3 replacement targets, paired by ordinal with `sourcePageIndices`. */
  targetPageIndices?: number[];
  /** v3 keeps every target page box/rotation and fits source presentation. */
  contentPlacement?: "fit-center";
}

export type PdfNativePageImportAuthority =
  | { state: "absent" }
  | { state: "complete"; records: PdfNativePageImportRecord[] }
  | {
      state: "malformed";
      raw: unknown;
      reason:
        | "invalid_content"
        | "invalid_collection"
        | "unsupported_schema"
        | "invalid_record"
        | "duplicate_id"
        | "target_resource_mismatch"
        | "source_budget"
        | "patch_limit";
    };

export interface PdfNativePageImportMutationDiff {
  index: number;
  before?: PdfNativePageImportRecord;
  after: PdfNativePageImportRecord;
}

const SHA256_PATTERN = /^[0-9a-f]{64}$/;
const STABLE_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._~:%-]{0,191}$/;
const BOUNDED_IDENTITY_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._~:/%-]{0,255}$/;
const RECORD_KEYS = [
  "schemaVersion",
  "operation",
  "id",
  "source",
  "targetSourceReceipt",
  "sourcePageIndices",
  "targetInsertionIndex",
] as const;
const RANGE_REPLACEMENT_RECORD_KEYS = [
  ...RECORD_KEYS,
  "targetPageIndices",
  "contentPlacement",
] as const;
const SOURCE_KEYS = [
  "dataItemId",
  "datasetId",
  "workspaceId",
  "notebookId",
  "sha256",
  "sizeBytes",
] as const;
const TARGET_RECEIPT_KEYS = [
  "resourceId",
  "resourceVersionId",
  "sha256",
] as const;

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function hasExactKeys(
  value: Record<string, unknown>,
  expected: readonly string[],
): boolean {
  const actual = Object.keys(value).sort();
  const wanted = [...expected].sort();
  return (
    actual.length === wanted.length &&
    actual.every((key, index) => key === wanted[index])
  );
}

function isBoundedIdentity(value: unknown): value is string {
  return typeof value === "string" && BOUNDED_IDENTITY_PATTERN.test(value);
}

function sanitizeSource(value: unknown): PdfNativePageImportSource | null {
  if (!isPlainObject(value) || !hasExactKeys(value, SOURCE_KEYS)) return null;
  if (
    !isBoundedIdentity(value.dataItemId) ||
    !isBoundedIdentity(value.datasetId) ||
    !isBoundedIdentity(value.workspaceId) ||
    !isBoundedIdentity(value.notebookId) ||
    typeof value.sha256 !== "string" ||
    !SHA256_PATTERN.test(value.sha256) ||
    typeof value.sizeBytes !== "number" ||
    !Number.isSafeInteger(value.sizeBytes) ||
    value.sizeBytes < 1 ||
    value.sizeBytes > MAX_PDF_NATIVE_PAGE_IMPORT_SOURCE_BYTES
  ) {
    return null;
  }
  return {
    dataItemId: value.dataItemId,
    datasetId: value.datasetId,
    workspaceId: value.workspaceId,
    notebookId: value.notebookId,
    sha256: value.sha256,
    sizeBytes: value.sizeBytes,
  };
}

function sanitizeTargetSourceReceipt(
  value: unknown,
): PdfNativePageImportTargetSourceReceipt | null {
  if (!isPlainObject(value) || !hasExactKeys(value, TARGET_RECEIPT_KEYS)) {
    return null;
  }
  if (
    !isBoundedIdentity(value.resourceId) ||
    !isBoundedIdentity(value.resourceVersionId) ||
    typeof value.sha256 !== "string" ||
    !SHA256_PATTERN.test(value.sha256)
  ) {
    return null;
  }
  return {
    resourceId: value.resourceId,
    resourceVersionId: value.resourceVersionId,
    sha256: value.sha256,
  };
}

function sanitizeSourcePageIndices(value: unknown): number[] | null {
  if (
    !Array.isArray(value) ||
    value.length < 1 ||
    value.length > MAX_PDF_NATIVE_PAGE_IMPORT_PAGES
  ) {
    return null;
  }
  const indices: number[] = [];
  let previous = -1;
  for (const index of value) {
    if (
      typeof index !== "number" ||
      !Number.isSafeInteger(index) ||
      index < 0 ||
      index > MAX_PDF_NATIVE_PAGE_INDEX ||
      index <= previous
    ) {
      return null;
    }
    indices.push(index);
    previous = index;
  }
  return indices;
}

function isSupportedSchemaVersion(
  value: unknown,
): value is PdfNativePageImportRecord["schemaVersion"] {
  return (
    value === PDF_NATIVE_PAGE_IMPORT_SCHEMA_VERSION ||
    value === PDF_NATIVE_PAGE_REPLACE_SCHEMA_VERSION ||
    value === PDF_NATIVE_PAGE_REPLACE_RANGE_SCHEMA_VERSION
  );
}

export function isPdfNativePageReplacementRecord(
  record: PdfNativePageImportRecord,
): boolean {
  return record.schemaVersion !== PDF_NATIVE_PAGE_IMPORT_SCHEMA_VERSION;
}

export function pdfNativePageReplacementTargetIndices(
  record: PdfNativePageImportRecord,
): number[] {
  return record.schemaVersion === PDF_NATIVE_PAGE_REPLACE_RANGE_SCHEMA_VERSION
    ? [...(record.targetPageIndices ?? [])]
    : [record.targetInsertionIndex];
}

/** Strict parser: validates persisted authority without repairing it. */
export function sanitizePdfNativePageImportRecord(
  value: unknown,
): PdfNativePageImportRecord | null {
  if (!isPlainObject(value) || !isSupportedSchemaVersion(value.schemaVersion)) {
    return null;
  }
  const isRangeReplacement =
    value.schemaVersion === PDF_NATIVE_PAGE_REPLACE_RANGE_SCHEMA_VERSION;
  if (
    !hasExactKeys(
      value,
      isRangeReplacement ? RANGE_REPLACEMENT_RECORD_KEYS : RECORD_KEYS,
    )
  ) {
    return null;
  }
  const source = sanitizeSource(value.source);
  const targetSourceReceipt = sanitizeTargetSourceReceipt(
    value.targetSourceReceipt,
  );
  const sourcePageIndices = sanitizeSourcePageIndices(value.sourcePageIndices);
  const targetPageIndices = isRangeReplacement
    ? sanitizeSourcePageIndices(value.targetPageIndices)
    : null;
  if (
    (value.operation !== "upsert" && value.operation !== "delete" && value.operation !== "settled") ||
    typeof value.id !== "string" ||
    !STABLE_ID_PATTERN.test(value.id) ||
    !source ||
    !targetSourceReceipt ||
    !sourcePageIndices ||
    typeof value.targetInsertionIndex !== "number" ||
    !Number.isSafeInteger(value.targetInsertionIndex) ||
    value.targetInsertionIndex < 0 ||
    value.targetInsertionIndex > MAX_PDF_NATIVE_TARGET_INSERTION_INDEX ||
    (value.schemaVersion === PDF_NATIVE_PAGE_REPLACE_SCHEMA_VERSION &&
      sourcePageIndices.length !== 1) ||
    (isRangeReplacement &&
      (!targetPageIndices ||
        targetPageIndices.length !== sourcePageIndices.length ||
        targetPageIndices[0] !== value.targetInsertionIndex ||
        value.contentPlacement !== "fit-center"))
  ) {
    return null;
  }
  const record: PdfNativePageImportRecord = {
    schemaVersion: value.schemaVersion,
    operation: value.operation,
    id: value.id,
    source,
    targetSourceReceipt,
    sourcePageIndices,
    targetInsertionIndex: value.targetInsertionIndex,
  };
  if (isRangeReplacement) {
    record.targetPageIndices = targetPageIndices!;
    record.contentPlacement = "fit-center";
  }
  return record;
}

export function extractPdfNativePageImportCollection(
  content: unknown,
): unknown | undefined {
  if (!isPlainObject(content)) return undefined;
  return Object.prototype.hasOwnProperty.call(
    content,
    PDF_NATIVE_PAGE_IMPORTS_COLLECTION,
  )
    ? content[PDF_NATIVE_PAGE_IMPORTS_COLLECTION]
    : undefined;
}

/**
 * Absent is distinct from an authoritative empty array. A malformed or future
 * reserved collection blocks native saves so unrelated edits cannot erase it.
 */
export function readPdfNativePageImportAuthority(
  content: unknown,
): PdfNativePageImportAuthority {
  if (!isPlainObject(content)) {
    return { state: "malformed", raw: content, reason: "invalid_content" };
  }
  if (
    !Object.prototype.hasOwnProperty.call(
      content,
      PDF_NATIVE_PAGE_IMPORTS_COLLECTION,
    )
  ) {
    return { state: "absent" };
  }
  const raw = content[PDF_NATIVE_PAGE_IMPORTS_COLLECTION];
  if (!Array.isArray(raw)) {
    return { state: "malformed", raw, reason: "invalid_collection" };
  }
  if (raw.length > MAX_PDF_NATIVE_PAGE_IMPORT_RECORDS) {
    return { state: "malformed", raw, reason: "patch_limit" };
  }
  const records: PdfNativePageImportRecord[] = [];
  const ids = new Set<string>();
  let targetResourceId: string | null = null;
  let mutationKind: "insert" | "replace" | null = null;
  for (const row of raw) {
    if (
      isPlainObject(row) &&
      Object.prototype.hasOwnProperty.call(row, "schemaVersion") &&
      !isSupportedSchemaVersion(row.schemaVersion)
    ) {
      return { state: "malformed", raw, reason: "unsupported_schema" };
    }
    const record = sanitizePdfNativePageImportRecord(row);
    if (!record) {
      return { state: "malformed", raw, reason: "invalid_record" };
    }
    if (ids.has(record.id)) {
      return { state: "malformed", raw, reason: "duplicate_id" };
    }
    if (
      targetResourceId !== null &&
      record.targetSourceReceipt.resourceId !== targetResourceId
    ) {
      return { state: "malformed", raw, reason: "target_resource_mismatch" };
    }
    const recordMutationKind = isPdfNativePageReplacementRecord(record)
      ? "replace"
      : "insert";
    if (mutationKind !== null && recordMutationKind !== mutationKind) {
      return { state: "malformed", raw, reason: "invalid_record" };
    }
    ids.add(record.id);
    records.push(record);
    targetResourceId = record.targetSourceReceipt.resourceId;
    mutationKind = recordMutationKind;
  }
  if (!staysWithinUniqueSourceBudget(records)) {
    return { state: "malformed", raw, reason: "source_budget" };
  }
  return { state: "complete", records };
}

function sourceEqual(
  left: PdfNativePageImportSource,
  right: PdfNativePageImportSource,
): boolean {
  return (
    left.dataItemId === right.dataItemId &&
    left.datasetId === right.datasetId &&
    left.workspaceId === right.workspaceId &&
    left.notebookId === right.notebookId &&
    left.sha256 === right.sha256 &&
    left.sizeBytes === right.sizeBytes
  );
}

function targetReceiptEqual(
  left: PdfNativePageImportTargetSourceReceipt,
  right: PdfNativePageImportTargetSourceReceipt,
): boolean {
  return (
    left.resourceId === right.resourceId &&
    left.resourceVersionId === right.resourceVersionId &&
    left.sha256 === right.sha256
  );
}

function pageIndicesEqual(left: readonly number[], right: readonly number[]) {
  return (
    left.length === right.length &&
    left.every((value, index) => value === right[index])
  );
}

function recordsEqual(
  left: PdfNativePageImportRecord,
  right: PdfNativePageImportRecord,
): boolean {
  return (
    left.schemaVersion === right.schemaVersion &&
    left.operation === right.operation &&
    left.id === right.id &&
    sourceEqual(left.source, right.source) &&
    targetReceiptEqual(left.targetSourceReceipt, right.targetSourceReceipt) &&
    pageIndicesEqual(left.sourcePageIndices, right.sourcePageIndices) &&
    left.targetInsertionIndex === right.targetInsertionIndex &&
    pageIndicesEqual(
      left.targetPageIndices ?? [],
      right.targetPageIndices ?? [],
    ) &&
    left.contentPlacement === right.contentPlacement
  );
}

function immutableIdentityEqual(
  left: PdfNativePageImportRecord,
  right: PdfNativePageImportRecord,
): boolean {
  return (
    left.schemaVersion === right.schemaVersion &&
    left.id === right.id &&
    sourceEqual(left.source, right.source) &&
    left.targetSourceReceipt.resourceId ===
      right.targetSourceReceipt.resourceId &&
    pageIndicesEqual(left.sourcePageIndices, right.sourcePageIndices) &&
    left.targetInsertionIndex === right.targetInsertionIndex &&
    pageIndicesEqual(
      left.targetPageIndices ?? [],
      right.targetPageIndices ?? [],
    ) &&
    left.contentPlacement === right.contentPlacement
  );
}

function isValidTombstoneTransition(
  before: PdfNativePageImportRecord,
  after: PdfNativePageImportRecord,
): boolean {
  return (
    before.operation !== after.operation &&
    // Settled is the last state, and only pages on the document settle.
    before.operation !== "settled" &&
    (after.operation !== "settled" || before.operation === "upsert") &&
    immutableIdentityEqual(before, after) &&
    !targetReceiptEqual(before.targetSourceReceipt, after.targetSourceReceipt)
  );
}

function hasOneTargetResource(
  records: readonly PdfNativePageImportRecord[],
): boolean {
  const ids = new Set(
    records.map((record) => record.targetSourceReceipt.resourceId),
  );
  return ids.size <= 1;
}

function hasOneMutationFamily(
  records: readonly PdfNativePageImportRecord[],
): boolean {
  return (
    new Set(
      records.map((record) =>
        isPdfNativePageReplacementRecord(record) ? "replace" : "insert",
      ),
    ).size <= 1
  );
}

function staysWithinUniqueSourceBudget(
  records: readonly PdfNativePageImportRecord[],
): boolean {
  const seen = new Set<string>();
  let total = 0;
  for (const record of records) {
    const source = record.source;
    const key = [
      source.dataItemId,
      source.datasetId,
      source.workspaceId,
      source.notebookId,
      source.sha256,
      source.sizeBytes,
    ].join("\u0000");
    if (seen.has(key)) continue;
    seen.add(key);
    total += source.sizeBytes;
    if (total > MAX_PDF_NATIVE_PAGE_IMPORT_AGGREGATE_SOURCE_BYTES) {
      return false;
    }
  }
  return true;
}

/** Apply exactly one append or same-id tombstone transition. */
export function applyPdfNativePageImportMutation(
  authority: PdfNativePageImportAuthority,
  mutation: unknown,
): PdfNativePageImportRecord[] | null {
  if (authority.state === "malformed") return null;
  const nextRecord = sanitizePdfNativePageImportRecord(mutation);
  if (!nextRecord) return null;
  const current = authority.state === "complete" ? authority.records : [];
  if (
    current.length > 0 &&
    (current[0].targetSourceReceipt.resourceId !==
      nextRecord.targetSourceReceipt.resourceId ||
      isPdfNativePageReplacementRecord(current[0]) !==
        isPdfNativePageReplacementRecord(nextRecord))
  ) {
    return null;
  }
  const index = current.findIndex((record) => record.id === nextRecord.id);
  if (index < 0) {
    if (
      nextRecord.operation !== "upsert" ||
      current.length >= MAX_PDF_NATIVE_PAGE_IMPORT_RECORDS
    ) {
      return null;
    }
    const appended = [...current, nextRecord];
    return staysWithinUniqueSourceBudget(appended) ? appended : null;
  }
  if (
    recordsEqual(current[index], nextRecord) ||
    !isValidTombstoneTransition(current[index], nextRecord)
  ) {
    return null;
  }
  const replaced = current.map((record, recordIndex) =>
    recordIndex === index ? nextRecord : record,
  );
  return staysWithinUniqueSourceBudget(replaced) ? replaced : null;
}

/**
 * Authenticate an exact one-record transition. Array order and every untouched
 * record are preconditions; ID-map comparisons must not accept reordered state.
 */
export function diffPdfNativePageImportMutation(
  beforeRaw: unknown,
  afterRaw: unknown,
): PdfNativePageImportMutationDiff | null {
  if (!Array.isArray(beforeRaw) || !Array.isArray(afterRaw)) return null;
  if (
    beforeRaw.length > MAX_PDF_NATIVE_PAGE_IMPORT_RECORDS ||
    afterRaw.length > MAX_PDF_NATIVE_PAGE_IMPORT_RECORDS
  ) {
    return null;
  }
  const before = beforeRaw.map(sanitizePdfNativePageImportRecord);
  const after = afterRaw.map(sanitizePdfNativePageImportRecord);
  if (before.some((record) => !record) || after.some((record) => !record)) {
    return null;
  }
  const beforeRecords = before as PdfNativePageImportRecord[];
  const afterRecords = after as PdfNativePageImportRecord[];
  if (
    new Set(beforeRecords.map((record) => record.id)).size !==
      beforeRecords.length ||
    new Set(afterRecords.map((record) => record.id)).size !==
      afterRecords.length ||
    !hasOneTargetResource(beforeRecords) ||
    !hasOneTargetResource(afterRecords) ||
    !hasOneMutationFamily(beforeRecords) ||
    !hasOneMutationFamily(afterRecords) ||
    !staysWithinUniqueSourceBudget(beforeRecords) ||
    !staysWithinUniqueSourceBudget(afterRecords)
  ) {
    return null;
  }
  if (afterRecords.length === beforeRecords.length + 1) {
    if (
      !beforeRecords.every((record, index) =>
        recordsEqual(record, afterRecords[index]),
      ) ||
      afterRecords[afterRecords.length - 1].operation !== "upsert" ||
      (beforeRecords.length > 0 &&
        beforeRecords[0].targetSourceReceipt.resourceId !==
          afterRecords[afterRecords.length - 1].targetSourceReceipt.resourceId)
    ) {
      return null;
    }
    return {
      index: afterRecords.length - 1,
      after: afterRecords[afterRecords.length - 1],
    };
  }
  if (afterRecords.length !== beforeRecords.length) return null;
  let changedIndex = -1;
  for (let index = 0; index < beforeRecords.length; index += 1) {
    if (recordsEqual(beforeRecords[index], afterRecords[index])) continue;
    if (
      changedIndex >= 0 ||
      beforeRecords[index].id !== afterRecords[index].id
    ) {
      return null;
    }
    changedIndex = index;
  }
  if (
    changedIndex < 0 ||
    !isValidTombstoneTransition(
      beforeRecords[changedIndex],
      afterRecords[changedIndex],
    )
  ) {
    return null;
  }
  return {
    index: changedIndex,
    before: beforeRecords[changedIndex],
    after: afterRecords[changedIndex],
  };
}
