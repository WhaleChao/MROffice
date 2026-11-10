/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Stable shared-state authority for native PDF file-attachment mutations.
 *
 * Attachment bytes never live in `content_json`. The record binds an already
 * uploaded, workspace-scoped DataItem to the exact PDF source receipt against
 * which the mutation was prepared. The server independently reopens both
 * objects and verifies every identity before it writes `/FileAttachment`,
 * `/Filespec`, `/EF`, and the EmbeddedFiles name tree.
 */

import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const PDF_NATIVE_FILE_ATTACHMENTS_COLLECTION =
  "pdfNativeFileAttachments" as const;
export const PDF_NATIVE_FILE_ATTACHMENT_SCHEMA_VERSION = 1 as const;
export const MAX_PDF_NATIVE_FILE_ATTACHMENTS = 400;
/**
 * An attachment has no size of its own: it is written into the PDF, and the
 * PDF must stay within its document total (lib/document-size-limits), which
 * the insert checks against the PDF's own size. One attachment, or all of a
 * PDF's together, can therefore never be larger than that total.
 */
export const MAX_PDF_NATIVE_FILE_ATTACHMENT_BYTES = PDF_EDITABLE_MAX_BYTES;
export const MAX_ACTIVE_PDF_NATIVE_FILE_ATTACHMENT_BYTES = PDF_EDITABLE_MAX_BYTES;

export const PDF_NATIVE_FILE_ATTACHMENT_ICONS = [
  "Paperclip",
  "PushPin",
  "Graph",
  "Tag",
] as const;

export type PdfNativeFileAttachmentIcon =
  (typeof PDF_NATIVE_FILE_ATTACHMENT_ICONS)[number];

export interface PdfNativeFileAttachmentSourceReceipt {
  resourceId: string;
  resourceVersionId: string;
  sha256: string;
}

/**
 * A delete is a full tombstone rather than a bare id. Retaining the immutable
 * attachment and source identities lets the server prove exactly which object
 * is being removed and keeps CRDT merges deterministic.
 */
export interface PdfNativeFileAttachmentRecord {
  schemaVersion: typeof PDF_NATIVE_FILE_ATTACHMENT_SCHEMA_VERSION;
  operation: "upsert" | "delete";
  id: string;
  /** One-based source PDF page. */
  page: number;
  /** Page-relative percentages. */
  x: number;
  y: number;
  filename: string;
  embeddedName: string;
  mimeType: string;
  sizeBytes: number;
  sha256: string;
  attachmentDataItemId: string;
  icon: PdfNativeFileAttachmentIcon;
  sourceReceipt: PdfNativeFileAttachmentSourceReceipt;
}

export type PdfNativeFileAttachmentAuthority =
  | { state: "absent" }
  | { state: "complete"; records: PdfNativeFileAttachmentRecord[] }
  | {
      state: "malformed";
      raw: unknown;
      reason:
        | "invalid_content"
        | "invalid_collection"
        | "unsupported_schema"
        | "invalid_record"
        | "duplicate_id"
        | "active_attachments_too_large"
        | "patch_limit";
    };

export interface PdfNativeFileAttachmentMutationDiff {
  index: number;
  before?: PdfNativeFileAttachmentRecord;
  after: PdfNativeFileAttachmentRecord;
}

const SHA256_PATTERN = /^[0-9a-f]{64}$/;
const STABLE_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._~:%-]{0,191}$/;
const BOUNDED_IDENTITY_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._~:/%-]{0,255}$/;
const RECORD_KEYS = [
  "schemaVersion",
  "operation",
  "id",
  "page",
  "x",
  "y",
  "filename",
  "embeddedName",
  "mimeType",
  "sizeBytes",
  "sha256",
  "attachmentDataItemId",
  "icon",
  "sourceReceipt",
] as const;
const RECEIPT_KEYS = ["resourceId", "resourceVersionId", "sha256"] as const;

const ALLOWED_EXTENSIONS_BY_MIME: Readonly<Record<string, readonly string[]>> =
  Object.freeze({
    "application/pdf": ["pdf"],
    "image/png": ["png"],
    "image/jpeg": ["jpg", "jpeg"],
    "image/webp": ["webp"],
    "image/gif": ["gif"],
    "image/bmp": ["bmp"],
    "image/tiff": ["tif", "tiff"],
    "text/plain": ["txt", "log"],
    "text/markdown": ["md"],
    "text/csv": ["csv"],
    "text/tab-separated-values": ["tsv"],
    "application/json": ["json"],
    "application/rtf": ["rtf"],
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [
      "docx",
    ],
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
      "xlsx",
    ],
    "application/vnd.openxmlformats-officedocument.presentationml.presentation":
      ["pptx"],
  });

const BLOCKED_EXTENSIONS = new Set([
  "7z",
  "app",
  "apk",
  "appimage",
  "bat",
  "bin",
  "bz2",
  "cab",
  "cmd",
  "com",
  "cpl",
  "deb",
  "dll",
  "dmg",
  "docm",
  "exe",
  "gz",
  "hta",
  "htm",
  "html",
  "img",
  "iso",
  "jar",
  "js",
  "jse",
  "lnk",
  "msi",
  "msp",
  "pif",
  "pkg",
  "pptm",
  "ps1",
  "rar",
  "reg",
  "rpm",
  "scr",
  "sh",
  "svg",
  "sys",
  "tar",
  "vbs",
  "vbe",
  "wsf",
  "xlam",
  "xlsm",
  "xml",
  "xz",
  "zip",
]);

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

function isFinitePercent(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 100
  );
}

function isValidPage(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 10_000
  );
}

function isBoundedIdentity(value: unknown): value is string {
  return typeof value === "string" && BOUNDED_IDENTITY_PATTERN.test(value);
}

function filenameExtension(value: string): string {
  const index = value.lastIndexOf(".");
  return index <= 0 ? "" : value.slice(index + 1).toLowerCase();
}

/**
 * Canonicalize a user-visible or embedded filename without ever manufacturing
 * a different extension. The caller must separately pass the MIME/extension
 * allowlist before the record becomes authoritative.
 */
export function sanitizePdfNativeFileAttachmentName(
  value: unknown,
): string | null {
  if (typeof value !== "string") return null;
  const normalized = value
    .normalize("NFKC")
    .replace(/[\u0000-\u001f\u007f]/g, "_")
    .replace(/[\u202a-\u202e\u2066-\u2069]/g, "_")
    .replace(/[\\/]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^[. ]+|[. ]+$/g, "");
  if (
    !normalized ||
    normalized === "." ||
    normalized === ".." ||
    !/[\p{L}\p{N}]/u.test(normalized)
  ) {
    return null;
  }
  if (normalized.length > 240) return null;
  return normalized;
}

export function isAllowedPdfNativeFileAttachmentType(
  filename: string,
  mimeType: string,
): boolean {
  if (mimeType !== mimeType.trim().toLowerCase()) return false;
  const extension = filenameExtension(filename);
  if (!extension || BLOCKED_EXTENSIONS.has(extension)) return false;
  return Boolean(ALLOWED_EXTENSIONS_BY_MIME[mimeType]?.includes(extension));
}

/**
 * Imported attachments have no trusted MIME contract. They may be offered for
 * an explicit download only when their visible filename is already canonical,
 * has an extension and is not one of the executable/archive/script families
 * rejected by the managed attachment writer.
 */
export function isSafePdfFileAttachmentDownloadName(value: string): boolean {
  const normalized = sanitizePdfNativeFileAttachmentName(value);
  if (!normalized || normalized !== value) return false;
  const extension = filenameExtension(normalized);
  return Boolean(extension) && !BLOCKED_EXTENSIONS.has(extension);
}

function sanitizeSourceReceipt(
  value: unknown,
): PdfNativeFileAttachmentSourceReceipt | null {
  if (!isPlainObject(value) || !hasExactKeys(value, RECEIPT_KEYS)) return null;
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

/** Strict parser: it validates but never repairs persisted authority. */
export function sanitizePdfNativeFileAttachmentRecord(
  value: unknown,
): PdfNativeFileAttachmentRecord | null {
  if (!isPlainObject(value) || !hasExactKeys(value, RECORD_KEYS)) return null;
  const filename = sanitizePdfNativeFileAttachmentName(value.filename);
  const embeddedName = sanitizePdfNativeFileAttachmentName(value.embeddedName);
  const sourceReceipt = sanitizeSourceReceipt(value.sourceReceipt);
  if (
    value.schemaVersion !== PDF_NATIVE_FILE_ATTACHMENT_SCHEMA_VERSION ||
    (value.operation !== "upsert" && value.operation !== "delete") ||
    typeof value.id !== "string" ||
    !STABLE_ID_PATTERN.test(value.id) ||
    !isValidPage(value.page) ||
    !isFinitePercent(value.x) ||
    !isFinitePercent(value.y) ||
    filename === null ||
    filename !== value.filename ||
    embeddedName === null ||
    embeddedName !== value.embeddedName ||
    typeof value.mimeType !== "string" ||
    !isAllowedPdfNativeFileAttachmentType(filename, value.mimeType) ||
    typeof value.sizeBytes !== "number" ||
    !Number.isSafeInteger(value.sizeBytes) ||
    value.sizeBytes < 1 ||
    value.sizeBytes > MAX_PDF_NATIVE_FILE_ATTACHMENT_BYTES ||
    typeof value.sha256 !== "string" ||
    !SHA256_PATTERN.test(value.sha256) ||
    !isBoundedIdentity(value.attachmentDataItemId) ||
    !(PDF_NATIVE_FILE_ATTACHMENT_ICONS as readonly unknown[]).includes(
      value.icon,
    ) ||
    !sourceReceipt
  ) {
    return null;
  }
  return {
    schemaVersion: PDF_NATIVE_FILE_ATTACHMENT_SCHEMA_VERSION,
    operation: value.operation,
    id: value.id,
    page: value.page,
    x: value.x,
    y: value.y,
    filename,
    embeddedName,
    mimeType: value.mimeType,
    sizeBytes: value.sizeBytes,
    sha256: value.sha256,
    attachmentDataItemId: value.attachmentDataItemId,
    icon: value.icon as PdfNativeFileAttachmentIcon,
    sourceReceipt,
  };
}

export function extractPdfNativeFileAttachmentCollection(
  content: unknown,
): unknown | undefined {
  if (!isPlainObject(content)) return undefined;
  return Object.prototype.hasOwnProperty.call(
    content,
    PDF_NATIVE_FILE_ATTACHMENTS_COLLECTION,
  )
    ? content[PDF_NATIVE_FILE_ATTACHMENTS_COLLECTION]
    : undefined;
}

/**
 * Preserve all three authority states. `absent` is not an authoritative empty
 * array; a present malformed/future collection must stop every native save so
 * an unrelated edit cannot silently erase information this version cannot read.
 */
export function readPdfNativeFileAttachmentAuthority(
  content: unknown,
): PdfNativeFileAttachmentAuthority {
  if (!isPlainObject(content)) {
    return { state: "malformed", raw: content, reason: "invalid_content" };
  }
  if (
    !Object.prototype.hasOwnProperty.call(
      content,
      PDF_NATIVE_FILE_ATTACHMENTS_COLLECTION,
    )
  ) {
    return { state: "absent" };
  }
  const raw = content[PDF_NATIVE_FILE_ATTACHMENTS_COLLECTION];
  if (!Array.isArray(raw)) {
    return { state: "malformed", raw, reason: "invalid_collection" };
  }
  if (raw.length > MAX_PDF_NATIVE_FILE_ATTACHMENTS) {
    return { state: "malformed", raw, reason: "patch_limit" };
  }
  const records: PdfNativeFileAttachmentRecord[] = [];
  const ids = new Set<string>();
  let activeBytes = 0;
  for (const row of raw) {
    if (
      isPlainObject(row) &&
      Object.prototype.hasOwnProperty.call(row, "schemaVersion") &&
      row.schemaVersion !== PDF_NATIVE_FILE_ATTACHMENT_SCHEMA_VERSION
    ) {
      return { state: "malformed", raw, reason: "unsupported_schema" };
    }
    const record = sanitizePdfNativeFileAttachmentRecord(row);
    if (!record) {
      return { state: "malformed", raw, reason: "invalid_record" };
    }
    if (ids.has(record.id)) {
      return { state: "malformed", raw, reason: "duplicate_id" };
    }
    ids.add(record.id);
    records.push(record);
    if (record.operation === "upsert") {
      activeBytes += record.sizeBytes;
      if (activeBytes > MAX_ACTIVE_PDF_NATIVE_FILE_ATTACHMENT_BYTES) {
        return {
          state: "malformed",
          raw,
          reason: "active_attachments_too_large",
        };
      }
    }
  }
  return { state: "complete", records };
}

function recordsEqual(
  left: PdfNativeFileAttachmentRecord,
  right: PdfNativeFileAttachmentRecord,
): boolean {
  return (
    left.schemaVersion === right.schemaVersion &&
    left.operation === right.operation &&
    left.id === right.id &&
    left.page === right.page &&
    left.x === right.x &&
    left.y === right.y &&
    left.filename === right.filename &&
    left.embeddedName === right.embeddedName &&
    left.mimeType === right.mimeType &&
    left.sizeBytes === right.sizeBytes &&
    left.sha256 === right.sha256 &&
    left.attachmentDataItemId === right.attachmentDataItemId &&
    left.icon === right.icon &&
    left.sourceReceipt.resourceId === right.sourceReceipt.resourceId &&
    left.sourceReceipt.resourceVersionId ===
      right.sourceReceipt.resourceVersionId &&
    left.sourceReceipt.sha256 === right.sourceReceipt.sha256
  );
}

function attachmentIdentityEqual(
  left: PdfNativeFileAttachmentRecord,
  right: PdfNativeFileAttachmentRecord,
): boolean {
  return (
    left.schemaVersion === right.schemaVersion &&
    left.id === right.id &&
    left.page === right.page &&
    left.x === right.x &&
    left.y === right.y &&
    left.filename === right.filename &&
    left.embeddedName === right.embeddedName &&
    left.mimeType === right.mimeType &&
    left.sizeBytes === right.sizeBytes &&
    left.sha256 === right.sha256 &&
    left.attachmentDataItemId === right.attachmentDataItemId &&
    left.icon === right.icon
  );
}

function isValidTombstoneTransition(
  before: PdfNativeFileAttachmentRecord,
  after: PdfNativeFileAttachmentRecord,
): boolean {
  if (before.operation !== "delete" && after.operation !== "delete") {
    return true;
  }
  return (
    before.operation !== after.operation &&
    attachmentIdentityEqual(before, after)
  );
}

/** Apply exactly one append or same-id replacement without reordering peers. */
export function applyPdfNativeFileAttachmentMutation(
  authority: PdfNativeFileAttachmentAuthority,
  mutation: unknown,
): PdfNativeFileAttachmentRecord[] | null {
  if (authority.state === "malformed") return null;
  const nextRecord = sanitizePdfNativeFileAttachmentRecord(mutation);
  if (!nextRecord) return null;
  const current = authority.state === "complete" ? authority.records : [];
  const index = current.findIndex((record) => record.id === nextRecord.id);
  if (index < 0) {
    if (
      nextRecord.operation !== "upsert" ||
      current.length >= MAX_PDF_NATIVE_FILE_ATTACHMENTS
    ) {
      return null;
    }
    const appended = [...current, nextRecord];
    return activeAttachmentBytes(appended) <=
      MAX_ACTIVE_PDF_NATIVE_FILE_ATTACHMENT_BYTES
      ? appended
      : null;
  }
  if (recordsEqual(current[index], nextRecord)) return null;
  if (!isValidTombstoneTransition(current[index], nextRecord)) return null;
  const replaced = current.map((record, recordIndex) =>
    recordIndex === index ? nextRecord : record,
  );
  return activeAttachmentBytes(replaced) <=
    MAX_ACTIVE_PDF_NATIVE_FILE_ATTACHMENT_BYTES
    ? replaced
    : null;
}

function activeAttachmentBytes(
  records: readonly PdfNativeFileAttachmentRecord[],
): number {
  return records.reduce(
    (total, record) =>
      record.operation === "upsert" ? total + record.sizeBytes : total,
    0,
  );
}

/**
 * Authenticate an exact one-record transition. Array order and every untouched
 * value are part of the precondition; an ID-map comparison would incorrectly
 * accept reordered concurrent state.
 */
export function diffPdfNativeFileAttachmentMutation(
  beforeRaw: unknown,
  afterRaw: unknown,
): PdfNativeFileAttachmentMutationDiff | null {
  if (!Array.isArray(beforeRaw) || !Array.isArray(afterRaw)) return null;
  const before = beforeRaw.map(sanitizePdfNativeFileAttachmentRecord);
  const after = afterRaw.map(sanitizePdfNativeFileAttachmentRecord);
  if (before.some((record) => !record) || after.some((record) => !record)) {
    return null;
  }
  const beforeRecords = before as PdfNativeFileAttachmentRecord[];
  const afterRecords = after as PdfNativeFileAttachmentRecord[];
  if (
    new Set(beforeRecords.map((record) => record.id)).size !==
      beforeRecords.length ||
    new Set(afterRecords.map((record) => record.id)).size !==
      afterRecords.length ||
    activeAttachmentBytes(beforeRecords) >
      MAX_ACTIVE_PDF_NATIVE_FILE_ATTACHMENT_BYTES ||
    activeAttachmentBytes(afterRecords) >
      MAX_ACTIVE_PDF_NATIVE_FILE_ATTACHMENT_BYTES
  ) {
    return null;
  }
  if (afterRecords.length === beforeRecords.length + 1) {
    if (
      !beforeRecords.every((record, index) =>
        recordsEqual(record, afterRecords[index]),
      ) ||
      afterRecords[afterRecords.length - 1].operation !== "upsert"
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
  if (changedIndex < 0) return null;
  if (
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
