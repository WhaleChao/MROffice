/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { resolvePageImportCapability } from "@mrnote/core/web/lib/page-import-capabilities.ts";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const PDF_SOURCE_REPLACEMENT_AUDIT_KIND =
  "mrnote.pdf-source-replacement-audit.v1" as const;
// Exclusive cap, aligned with the API's PDF projection and native patchback
// workers. Keeping this stricter than the server would make a valid existing
// PDF impossible to re-import through the editor.
export const PDF_SOURCE_REPLACEMENT_MAX_BYTES = PDF_EDITABLE_MAX_BYTES;
export const PDF_SOURCE_REPLACEMENT_CLEANUP_LEDGER_KEY =
  "mrnote:pdf-source-replacement-cleanup:v1" as const;
export const PDF_SOURCE_REPLACEMENT_CLEANUP_KIND =
  "mrnote.pdf-source-replacement-cleanup.v1" as const;
export const PDF_SOURCE_REPLACEMENT_CLEANUP_RECOVERY_LEASE_MS = 30_000;
export const PDF_SOURCE_REPLACEMENT_CLEANUP_HEARTBEAT_MS = 10_000;

export type PdfSourcePreflightFailure =
  "empty" | "invalid-signature" | "not-pdf" | "too-large";

export type PdfSourcePreflightResult =
  | { ok: true }
  | {
      ok: false;
      code: PdfSourcePreflightFailure;
      maxBytes?: number;
    };

export type PdfSourceReplacementLossKind =
  | "annotations"
  | "form-fields"
  | "form-values"
  | "page-transforms"
  | "deleted-pages"
  | "inserted-pages"
  | "page-moves"
  | "text-edits"
  | "image-overlays"
  | "ink-strokes"
  | "crop-boxes"
  | "sticky-notes"
  | "ink-annotations"
  | "signatures"
  | "redactions"
  | "reader-annotations";

export interface PdfSourceReplacementLossEntry {
  kind: PdfSourceReplacementLossKind;
  count: number;
}

export interface PdfSourceReplacementLossManifest {
  entries: PdfSourceReplacementLossEntry[];
  total: number;
}

export interface PdfSourceReplacementAudit {
  kind: typeof PDF_SOURCE_REPLACEMENT_AUDIT_KIND;
  replacedAt: string;
  sourceRetained: true;
  previous: {
    dataItemId: string;
    filename: string;
    sha256: string | null;
    riskProfileDigest: string | null;
    sizeBytes: number | null;
    mediaType: string | null;
    contentVersion: number | null;
    pageVersionId: string | null;
    pageVersionNo: number | null;
  };
  replacement: {
    filename: string;
    sha256: string | null;
    riskProfileDigest: string | null;
    sizeBytes: number;
    mediaType: string | null;
  };
  lossDigest: string | null;
  losses: PdfSourceReplacementLossEntry[];
  totalLossEntries: number;
}

export interface PdfSourceReplacementCleanupHandle {
  readonly kind: typeof PDF_SOURCE_REPLACEMENT_CLEANUP_KIND;
  readonly notebookId: string;
  readonly pageId: string;
  readonly sourceDataItemId: string;
  readonly sourceSha256: string;
  readonly sourceDocumentRevision: string | number;
  readonly replacementSha256: string;
  readonly uploadedDataItemId: string;
  /**
   * Null until the uploaded object is submitted in a page CAS. Once set, a
   * missing/failed client receipt is ambiguous: without an atomic server
   * reference guard, the client must retain the object unless the authoritative
   * page already proves that this exact upload is attached.
   */
  readonly publicationAttemptedAt: number | null;
  readonly recordedAt: number;
}

const PDF_SIGNATURE = [0x25, 0x50, 0x44, 0x46, 0x2d] as const;
const PDF_SIGNATURE_SCAN_BYTES = 1024;
const MAX_LOSS_ENTRY_COUNT = 100_000;
const MAX_AUDIT_STRING_LENGTH = 512;
const MAX_CLEANUP_LEDGER_ENTRIES = 32;
// The writer cap is deliberately lower than the reader cap. Independent tabs
// can both observe the last free slot before either atomic per-upload key is
// visible. Recovery must keep both obligations readable instead of turning
// that benign overrun into a lost cleanup handle.
const MAX_CLEANUP_LEDGER_RECOVERY_ENTRIES = 128;
const MAX_CLEANUP_ID_LENGTH = 160;
const MAX_LOSS_DIGEST_LENGTH = 4096;
const SHA256_RE = /^[a-f0-9]{64}$/;
// One state letter per save-policy risk key. The profile has nine keys since
// 499fcad8b4 split certification from approval signatures; audits persisted
// before that carry the eight-key digest and must stay readable after reopen.
const RISK_PROFILE_DIGEST_RE = /^pdf-risk-v1:[npu]{8,9}$/;
const LOSS_DIGEST_RE =
  /^pdf-loss-v1:(?:[A-Z0-9][A-Z0-9_.-]*(?:,[A-Z0-9][A-Z0-9_.-]*)*)?$/;
const LOSS_KINDS = new Set<PdfSourceReplacementLossKind>([
  "annotations",
  "form-fields",
  "form-values",
  "page-transforms",
  "deleted-pages",
  "inserted-pages",
  "page-moves",
  "text-edits",
  "image-overlays",
  "ink-strokes",
  "crop-boxes",
  "sticky-notes",
  "ink-annotations",
  "signatures",
  "redactions",
  "reader-annotations",
]);
const MAX_TOTAL_LOSS_ENTRY_COUNT = MAX_LOSS_ENTRY_COUNT * LOSS_KINDS.size;
const activeCleanupDataItemIds = new Set<string>();
const releasedCleanupDataItemIds = new Set<string>();
const volatileCleanupHandles = new Map<
  string,
  PdfSourceReplacementCleanupHandle
>();
const volatileOnlyCleanupDataItemIds = new Set<string>();
let volatileCleanupBeforeUnloadInstalled = false;

interface PdfSourceFileLike {
  name: string;
  type?: string;
  size: number;
  slice(start?: number, end?: number): Blob;
}

function exactCleanupString(value: unknown, label: string): string {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > MAX_CLEANUP_ID_LENGTH ||
    value !== value.trim() ||
    value.includes("\0")
  ) {
    throw new Error(`Invalid ${label}`);
  }
  return value;
}

function exactCleanupRevision(value: unknown): string | number {
  if (typeof value === "number" && Number.isSafeInteger(value) && value >= 0) {
    return value;
  }
  return exactCleanupString(value, "PDF cleanup source revision");
}

function exactCleanupTimestamp(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0) {
    throw new Error(`Invalid ${label}`);
  }
  return value;
}

function normalizeCleanupHandle(
  value: Omit<
    PdfSourceReplacementCleanupHandle,
    "kind" | "publicationAttemptedAt" | "recordedAt"
  > & {
    readonly kind?: typeof PDF_SOURCE_REPLACEMENT_CLEANUP_KIND;
    readonly publicationAttemptedAt?: number | null;
    readonly recordedAt?: number;
  },
): PdfSourceReplacementCleanupHandle {
  if (
    value.kind !== undefined &&
    value.kind !== PDF_SOURCE_REPLACEMENT_CLEANUP_KIND
  ) {
    throw new Error("Invalid PDF cleanup handle kind");
  }
  const sourceSha256 = exactCleanupString(
    value.sourceSha256,
    "PDF cleanup source SHA-256",
  ).toLowerCase();
  const replacementSha256 = exactCleanupString(
    value.replacementSha256,
    "PDF cleanup replacement SHA-256",
  ).toLowerCase();
  if (!SHA256_RE.test(sourceSha256) || !SHA256_RE.test(replacementSha256)) {
    throw new Error("Invalid PDF cleanup SHA-256");
  }
  const recordedAt = value.recordedAt ?? Date.now();
  const publicationAttemptedAt =
    value.publicationAttemptedAt == null
      ? null
      : exactCleanupTimestamp(
          value.publicationAttemptedAt,
          "PDF cleanup publication timestamp",
        );
  return Object.freeze({
    kind: PDF_SOURCE_REPLACEMENT_CLEANUP_KIND,
    notebookId: exactCleanupString(value.notebookId, "PDF cleanup notebook ID"),
    pageId: exactCleanupString(value.pageId, "PDF cleanup page ID"),
    sourceDataItemId: exactCleanupString(
      value.sourceDataItemId,
      "PDF cleanup source data item ID",
    ),
    sourceSha256,
    sourceDocumentRevision: exactCleanupRevision(value.sourceDocumentRevision),
    replacementSha256,
    uploadedDataItemId: exactCleanupString(
      value.uploadedDataItemId,
      "PDF cleanup uploaded data item ID",
    ),
    publicationAttemptedAt,
    recordedAt: exactCleanupTimestamp(recordedAt, "PDF cleanup timestamp"),
  });
}

function volatileCleanupBeforeUnload(event: BeforeUnloadEvent): void {
  event.preventDefault();
  event.returnValue = "";
}

function syncVolatileCleanupBeforeUnload(): void {
  if (typeof window === "undefined") return;
  const required = volatileOnlyCleanupDataItemIds.size > 0;
  if (required === volatileCleanupBeforeUnloadInstalled) return;
  if (required) {
    window.addEventListener("beforeunload", volatileCleanupBeforeUnload);
  } else {
    window.removeEventListener("beforeunload", volatileCleanupBeforeUnload);
  }
  volatileCleanupBeforeUnloadInstalled = required;
}

function cleanupStorages(storage?: Storage): readonly Storage[] {
  if (storage) return Object.freeze([storage]);
  if (typeof window === "undefined") {
    return Object.freeze([]);
  }
  const candidates: Storage[] = [];
  for (const read of [() => window.localStorage, () => window.sessionStorage]) {
    try {
      const candidate = read();
      if (!candidates.includes(candidate)) candidates.push(candidate);
    } catch {
      // Continue to the independent per-tab store. The in-memory custody map
      // remains the final same-session fallback and installs an unload guard.
    }
  }
  return Object.freeze(candidates);
}

function cleanupStorageKey(uploadedDataItemId: string): string {
  return `${PDF_SOURCE_REPLACEMENT_CLEANUP_LEDGER_KEY}:${encodeURIComponent(
    exactCleanupString(uploadedDataItemId, "PDF cleanup uploaded data item ID"),
  )}`;
}

function parseStoredCleanupHandle(
  key: string,
  raw: string,
): PdfSourceReplacementCleanupHandle {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error("Invalid PDF source-replacement cleanup handle");
  }
  if (!parsed || typeof parsed !== "object") {
    throw new Error("Invalid PDF source-replacement cleanup handle");
  }
  const row = parsed as Record<string, unknown>;
  const handle = normalizeCleanupHandle({
    kind: row.kind as typeof PDF_SOURCE_REPLACEMENT_CLEANUP_KIND,
    notebookId: row.notebookId as string,
    pageId: row.pageId as string,
    sourceDataItemId: row.sourceDataItemId as string,
    sourceSha256: row.sourceSha256 as string,
    sourceDocumentRevision: row.sourceDocumentRevision as string | number,
    replacementSha256: row.replacementSha256 as string,
    uploadedDataItemId: row.uploadedDataItemId as string,
    // Rows created before the publication-state hardening are conservatively
    // treated as having attempted publication; they must never be deleted on
    // the assumption that an old client did not submit the page CAS.
    publicationAttemptedAt:
      row.publicationAttemptedAt === undefined
        ? (row.recordedAt as number)
        : (row.publicationAttemptedAt as number | null),
    recordedAt: row.recordedAt as number,
  });
  if (key !== cleanupStorageKey(handle.uploadedDataItemId)) {
    throw new Error("Mismatched PDF cleanup data item custody");
  }
  return handle;
}

function sameCleanupHandle(
  left: PdfSourceReplacementCleanupHandle,
  right: PdfSourceReplacementCleanupHandle,
): boolean {
  return (
    left.kind === right.kind &&
    left.notebookId === right.notebookId &&
    left.pageId === right.pageId &&
    left.sourceDataItemId === right.sourceDataItemId &&
    left.sourceSha256 === right.sourceSha256 &&
    left.sourceDocumentRevision === right.sourceDocumentRevision &&
    left.replacementSha256 === right.replacementSha256 &&
    left.uploadedDataItemId === right.uploadedDataItemId
  );
}

function mergeCleanupHandleMetadata(
  left: PdfSourceReplacementCleanupHandle,
  right: PdfSourceReplacementCleanupHandle,
): PdfSourceReplacementCleanupHandle {
  if (!sameCleanupHandle(left, right)) {
    throw new Error("Conflicting PDF cleanup data item custody");
  }
  return normalizeCleanupHandle({
    ...left,
    recordedAt: Math.max(left.recordedAt, right.recordedAt),
    publicationAttemptedAt:
      left.publicationAttemptedAt === null
        ? right.publicationAttemptedAt
        : right.publicationAttemptedAt === null
          ? left.publicationAttemptedAt
          : Math.max(left.publicationAttemptedAt, right.publicationAttemptedAt),
  });
}

function persistCleanupHandle(
  handle: PdfSourceReplacementCleanupHandle,
  storage?: Storage,
): boolean {
  let lastError: unknown = null;
  for (const target of cleanupStorages(storage)) {
    try {
      target.setItem(
        cleanupStorageKey(handle.uploadedDataItemId),
        JSON.stringify(handle),
      );
      volatileOnlyCleanupDataItemIds.delete(handle.uploadedDataItemId);
      syncVolatileCleanupBeforeUnload();
      return true;
    } catch (error) {
      lastError = error;
    }
  }
  volatileOnlyCleanupDataItemIds.add(handle.uploadedDataItemId);
  syncVolatileCleanupBeforeUnload();
  if (storage && lastError instanceof Error) throw lastError;
  return false;
}

export function createPdfSourceReplacementCleanupHandle(
  value: Omit<
    PdfSourceReplacementCleanupHandle,
    "kind" | "publicationAttemptedAt" | "recordedAt"
  > & {
    readonly publicationAttemptedAt?: number | null;
    readonly recordedAt?: number;
  },
): PdfSourceReplacementCleanupHandle {
  return normalizeCleanupHandle(value);
}

export function readPdfSourceReplacementCleanupLedger(
  storage?: Storage,
): readonly PdfSourceReplacementCleanupHandle[] {
  const prefix = `${PDF_SOURCE_REPLACEMENT_CLEANUP_LEDGER_KEY}:`;
  const entries = new Map<string, PdfSourceReplacementCleanupHandle>();
  // localStorage itself is quota bounded. Each matching entry is additionally
  // size/shape bounded by normalizeCleanupHandle before it can enter recovery.
  for (const target of cleanupStorages(storage)) {
    for (let index = 0; index < target.length; index += 1) {
      const key = target.key(index);
      if (!key?.startsWith(prefix)) continue;
      const raw = target.getItem(key);
      if (raw === null) continue;
      const parsed = parseStoredCleanupHandle(key, raw);
      const existing = entries.get(parsed.uploadedDataItemId);
      entries.set(
        parsed.uploadedDataItemId,
        existing ? mergeCleanupHandleMetadata(existing, parsed) : parsed,
      );
      if (entries.size > MAX_CLEANUP_LEDGER_RECOVERY_ENTRIES) {
        throw new Error("PDF cleanup ledger recovery capacity exceeded");
      }
    }
  }
  for (const entry of volatileCleanupHandles.values()) {
    const existing = entries.get(entry.uploadedDataItemId);
    entries.set(
      entry.uploadedDataItemId,
      existing ? mergeCleanupHandleMetadata(existing, entry) : entry,
    );
    if (entries.size > MAX_CLEANUP_LEDGER_RECOVERY_ENTRIES) {
      throw new Error("PDF cleanup ledger recovery capacity exceeded");
    }
  }
  const result = [...entries.values()];
  result.sort(
    (left, right) =>
      left.recordedAt - right.recordedAt ||
      left.uploadedDataItemId.localeCompare(right.uploadedDataItemId),
  );
  return Object.freeze(result);
}

export function recordPdfSourceReplacementCleanupHandle(
  handle: PdfSourceReplacementCleanupHandle,
  storage?: Storage,
): void {
  const normalized = normalizeCleanupHandle(handle);
  let current = normalized;
  const volatileExisting = volatileCleanupHandles.get(
    normalized.uploadedDataItemId,
  );
  if (volatileExisting && !sameCleanupHandle(volatileExisting, normalized)) {
    throw new Error("Conflicting PDF cleanup data item custody");
  }
  if (volatileExisting) {
    current = mergeCleanupHandleMetadata(volatileExisting, current);
  }
  const ledger = readPdfSourceReplacementCleanupLedger(storage);
  const existing = ledger.find(
    (entry) => entry.uploadedDataItemId === normalized.uploadedDataItemId,
  );
  if (existing) {
    if (!sameCleanupHandle(existing, normalized)) {
      throw new Error("Conflicting PDF cleanup data item custody");
    }
    current = mergeCleanupHandleMetadata(existing, current);
  }
  volatileCleanupHandles.set(normalized.uploadedDataItemId, current);
  if (
    ledger.filter(
      (entry) => entry.uploadedDataItemId !== normalized.uploadedDataItemId,
    ).length >= MAX_CLEANUP_LEDGER_ENTRIES
  ) {
    volatileOnlyCleanupDataItemIds.add(normalized.uploadedDataItemId);
    syncVolatileCleanupBeforeUnload();
    throw new Error("PDF cleanup ledger capacity exceeded");
  }
  // Each upload owns one key. Two tabs recording different uploads therefore
  // never overwrite each other's read-modify-write array.
  // Re-recording a stale snapshot is monotonic: a non-null publication marker
  // observed in any durable or volatile copy can never be downgraded to null.
  if (!persistCleanupHandle(current, storage)) {
    throw new Error("PDF cleanup storage is unavailable");
  }
  activeCleanupDataItemIds.add(normalized.uploadedDataItemId);
  releasedCleanupDataItemIds.delete(normalized.uploadedDataItemId);
}

export function claimPdfSourceReplacementCleanupHandle(
  handle: PdfSourceReplacementCleanupHandle,
  now = Date.now(),
  storage?: Storage,
): boolean {
  const normalized = normalizeCleanupHandle(handle);
  if (activeCleanupDataItemIds.has(normalized.uploadedDataItemId)) return false;
  if (
    !releasedCleanupDataItemIds.has(normalized.uploadedDataItemId) &&
    now - normalized.recordedAt <
      PDF_SOURCE_REPLACEMENT_CLEANUP_RECOVERY_LEASE_MS
  ) {
    return false;
  }
  const stored = readPdfSourceReplacementCleanupLedger(storage).find(
    (entry) => entry.uploadedDataItemId === normalized.uploadedDataItemId,
  );
  if (!stored || !sameCleanupHandle(stored, normalized)) return false;
  // Publish a fresh per-upload lease before any async recovery side effect.
  // A second tab that wakes from an older timer will re-read this timestamp and
  // defer instead of racing the same authoritative reconciliation.
  const claimed = normalizeCleanupHandle({ ...stored, recordedAt: now });
  volatileCleanupHandles.set(claimed.uploadedDataItemId, claimed);
  persistCleanupHandle(claimed, storage);
  activeCleanupDataItemIds.add(normalized.uploadedDataItemId);
  releasedCleanupDataItemIds.delete(normalized.uploadedDataItemId);
  return true;
}

export function renewPdfSourceReplacementCleanupHandleLease(
  handle: PdfSourceReplacementCleanupHandle,
  storage?: Storage,
  now = Date.now(),
): PdfSourceReplacementCleanupHandle {
  const normalized = normalizeCleanupHandle(handle);
  const stored = readPdfSourceReplacementCleanupLedger(storage).find(
    (entry) => entry.uploadedDataItemId === normalized.uploadedDataItemId,
  );
  if (!stored || !sameCleanupHandle(stored, normalized)) {
    throw new Error("Missing PDF cleanup data item custody");
  }
  const renewed = normalizeCleanupHandle({ ...stored, recordedAt: now });
  volatileCleanupHandles.set(renewed.uploadedDataItemId, renewed);
  if (!persistCleanupHandle(renewed, storage)) {
    throw new Error("PDF cleanup lease custody is not durable");
  }
  activeCleanupDataItemIds.add(renewed.uploadedDataItemId);
  releasedCleanupDataItemIds.delete(renewed.uploadedDataItemId);
  return renewed;
}

export function markPdfSourceReplacementCleanupPublicationAttempted(
  handle: PdfSourceReplacementCleanupHandle,
  storage?: Storage,
  now = Date.now(),
): PdfSourceReplacementCleanupHandle {
  const normalized = normalizeCleanupHandle(handle);
  const current = readPdfSourceReplacementCleanupLedger(storage).find(
    (entry) => entry.uploadedDataItemId === normalized.uploadedDataItemId,
  );
  if (!current || !sameCleanupHandle(current, normalized)) {
    throw new Error("Missing PDF cleanup data item custody");
  }
  const attempted = normalizeCleanupHandle({
    ...current,
    publicationAttemptedAt: exactCleanupTimestamp(
      now,
      "PDF cleanup publication timestamp",
    ),
    recordedAt: now,
  });
  volatileCleanupHandles.set(attempted.uploadedDataItemId, attempted);
  if (!persistCleanupHandle(attempted, storage)) {
    throw new Error("PDF cleanup publication custody is not durable");
  }
  activeCleanupDataItemIds.add(attempted.uploadedDataItemId);
  releasedCleanupDataItemIds.delete(attempted.uploadedDataItemId);
  return attempted;
}

export function releasePdfSourceReplacementCleanupHandle(
  handle: PdfSourceReplacementCleanupHandle,
): void {
  activeCleanupDataItemIds.delete(handle.uploadedDataItemId);
  releasedCleanupDataItemIds.add(handle.uploadedDataItemId);
}

export function clearPdfSourceReplacementCleanupHandle(
  handle: PdfSourceReplacementCleanupHandle,
  storage?: Storage,
): void {
  clearPdfSourceReplacementCleanupHandleInternal(handle, storage, false);
}

export function clearAttachedPdfSourceReplacementCleanupHandle(
  handle: PdfSourceReplacementCleanupHandle,
  storage?: Storage,
): void {
  clearPdfSourceReplacementCleanupHandleInternal(handle, storage, true);
}

function clearPdfSourceReplacementCleanupHandleInternal(
  handle: PdfSourceReplacementCleanupHandle,
  storage: Storage | undefined,
  authoritativeAttachment: boolean,
): void {
  const normalized = normalizeCleanupHandle(handle);
  const existing = readPdfSourceReplacementCleanupLedger(storage).find(
    (entry) => entry.uploadedDataItemId === normalized.uploadedDataItemId,
  );
  if (existing && !sameCleanupHandle(existing, normalized)) {
    throw new Error("Conflicting PDF cleanup data item custody");
  }
  if (
    existing &&
    !authoritativeAttachment &&
    existing.publicationAttemptedAt !== normalized.publicationAttemptedAt
  ) {
    throw new Error("Stale PDF cleanup publication custody");
  }
  let removalError: unknown = null;
  const targets = cleanupStorages(storage);
  if (
    targets.length === 0 &&
    !volatileOnlyCleanupDataItemIds.has(normalized.uploadedDataItemId)
  ) {
    throw new Error("PDF cleanup storage is unavailable");
  }
  for (const target of targets) {
    try {
      target.removeItem(cleanupStorageKey(normalized.uploadedDataItemId));
    } catch (error) {
      removalError ??= error;
    }
  }
  if (removalError) {
    if (removalError instanceof Error) throw removalError;
    throw new Error("PDF cleanup storage could not be cleared");
  }
  volatileCleanupHandles.delete(normalized.uploadedDataItemId);
  volatileOnlyCleanupDataItemIds.delete(normalized.uploadedDataItemId);
  syncVolatileCleanupBeforeUnload();
  activeCleanupDataItemIds.delete(normalized.uploadedDataItemId);
  releasedCleanupDataItemIds.delete(normalized.uploadedDataItemId);
}

export function readRecoverablePdfSourceReplacementCleanupHandles(
  storage?: Storage,
  now = Date.now(),
): readonly PdfSourceReplacementCleanupHandle[] {
  return Object.freeze(
    readPdfSourceReplacementCleanupLedger(storage).filter(
      (entry) =>
        !activeCleanupDataItemIds.has(entry.uploadedDataItemId) &&
        (releasedCleanupDataItemIds.has(entry.uploadedDataItemId) ||
          now - entry.recordedAt >=
            PDF_SOURCE_REPLACEMENT_CLEANUP_RECOVERY_LEASE_MS),
    ),
  );
}

function boundedCount(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    return 0;
  }
  return Math.min(MAX_LOSS_ENTRY_COUNT, Math.floor(value));
}

function boundedString(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim();
  if (!normalized || normalized.length > MAX_AUDIT_STRING_LENGTH) return null;
  return normalized;
}

function nullableBoundedString(value: unknown): string | null {
  if (value == null || value === "") return null;
  return boundedString(value);
}

function normalizeOptionalSha256(value: unknown): string | null {
  if (value == null || value === "") return null;
  if (typeof value !== "string") throw new Error("Invalid PDF source SHA-256");
  const normalized = value.toLowerCase();
  if (!SHA256_RE.test(normalized))
    throw new Error("Invalid PDF source SHA-256");
  return normalized;
}

function normalizeOptionalRiskProfileDigest(value: unknown): string | null {
  if (value == null || value === "") return null;
  if (typeof value !== "string" || !RISK_PROFILE_DIGEST_RE.test(value)) {
    throw new Error("Invalid PDF risk-profile digest");
  }
  return value;
}

function normalizeOptionalLossDigest(value: unknown): string | null {
  if (value == null || value === "") return null;
  if (
    typeof value !== "string" ||
    value.length > MAX_LOSS_DIGEST_LENGTH ||
    !LOSS_DIGEST_RE.test(value)
  ) {
    throw new Error("Invalid PDF loss digest");
  }
  return value;
}

function parseOptionalBoundIdentity<T>(
  value: unknown,
  normalize: (candidate: unknown) => T | null,
): { valid: true; value: T | null } | { valid: false } {
  try {
    return { valid: true, value: normalize(value) };
  } catch {
    return { valid: false };
  }
}

function nullableNonNegativeInteger(value: unknown): number | null {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < 0 ||
    !Number.isInteger(value)
  ) {
    return null;
  }
  return value;
}

function boundedTotalCount(value: unknown): number | null {
  const count = nullableNonNegativeInteger(value);
  if (count === null || count > MAX_TOTAL_LOSS_ENTRY_COUNT) return null;
  return count;
}

async function readBlobBytes(blob: Blob): Promise<Uint8Array> {
  if (typeof blob.arrayBuffer === "function") {
    return new Uint8Array(await blob.arrayBuffer());
  }
  if (typeof FileReader === "undefined") {
    throw new Error("FileReader is unavailable");
  }
  return await new Promise<Uint8Array>((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () =>
      reject(reader.error ?? new Error("File read failed"));
    reader.onload = () => {
      if (!(reader.result instanceof ArrayBuffer)) {
        reject(new Error("File read failed"));
        return;
      }
      resolve(new Uint8Array(reader.result));
    };
    reader.readAsArrayBuffer(blob);
  });
}

function containsPdfSignature(bytes: Uint8Array): boolean {
  if (bytes.length < PDF_SIGNATURE.length) return false;
  const lastStart = bytes.length - PDF_SIGNATURE.length;
  for (let offset = 0; offset <= lastStart; offset += 1) {
    let matches = true;
    for (let index = 0; index < PDF_SIGNATURE.length; index += 1) {
      if (bytes[offset + index] !== PDF_SIGNATURE[index]) {
        matches = false;
        break;
      }
    }
    if (matches) return true;
  }
  return false;
}

/**
 * Early client preflight for source replacement. Server-side signature, ACL,
 * quota and adapter validation remain authoritative after upload.
 */
export async function preflightPdfSourceFile(
  file: PdfSourceFileLike,
): Promise<PdfSourcePreflightResult> {
  if (resolvePageImportCapability(file)?.id !== "pdf") {
    return { ok: false, code: "not-pdf" };
  }
  if (!Number.isFinite(file.size) || file.size <= 0) {
    return { ok: false, code: "empty" };
  }
  if (file.size >= PDF_SOURCE_REPLACEMENT_MAX_BYTES) {
    return {
      ok: false,
      code: "too-large",
      maxBytes: PDF_SOURCE_REPLACEMENT_MAX_BYTES,
    };
  }
  try {
    const bytes = await readBlobBytes(
      file.slice(0, Math.min(PDF_SIGNATURE_SCAN_BYTES, file.size)),
    );
    return containsPdfSignature(bytes)
      ? { ok: true }
      : { ok: false, code: "invalid-signature" };
  } catch {
    return { ok: false, code: "invalid-signature" };
  }
}

function firstArrayLength(
  content: Record<string, unknown>,
  keys: readonly string[],
): number {
  for (const key of keys) {
    const value = content[key];
    if (Array.isArray(value)) return boundedCount(value.length);
  }
  return 0;
}

function firstObjectSize(
  content: Record<string, unknown>,
  keys: readonly string[],
): number {
  for (const key of keys) {
    const value = content[key];
    if (value && typeof value === "object" && !Array.isArray(value)) {
      return boundedCount(Object.keys(value).length);
    }
  }
  return 0;
}

export function buildPdfSourceReplacementLossManifest(
  content: Record<string, unknown> | undefined,
  readerAnnotationCount = 0,
): PdfSourceReplacementLossManifest {
  const safeContent = content ?? {};
  const candidates: PdfSourceReplacementLossEntry[] = [
    {
      kind: "annotations",
      count: firstArrayLength(safeContent, ["annotations"]),
    },
    {
      kind: "form-fields",
      count: firstArrayLength(safeContent, ["formFields", "form_fields"]),
    },
    {
      kind: "form-values",
      count: firstObjectSize(safeContent, ["formValues", "form_values"]),
    },
    {
      kind: "page-transforms",
      count: firstArrayLength(safeContent, [
        "pageTransforms",
        "page_transforms",
      ]),
    },
    {
      kind: "deleted-pages",
      count: firstArrayLength(safeContent, ["deletedPages", "deleted_pages"]),
    },
    {
      kind: "inserted-pages",
      count: firstArrayLength(safeContent, ["insertedPages", "inserted_pages"]),
    },
    {
      kind: "page-moves",
      count: firstArrayLength(safeContent, ["pageMoves", "page_moves"]),
    },
    {
      kind: "text-edits",
      count: firstArrayLength(safeContent, ["textBoxEdits", "text_box_edits"]),
    },
    {
      kind: "image-overlays",
      count: firstArrayLength(safeContent, ["imageOverlays", "image_overlays"]),
    },
    {
      kind: "ink-strokes",
      count: firstArrayLength(safeContent, ["inkStrokes", "ink_strokes"]),
    },
    {
      kind: "crop-boxes",
      count: firstArrayLength(safeContent, ["cropBoxes", "crop_boxes"]),
    },
    {
      kind: "sticky-notes",
      count: firstArrayLength(safeContent, ["stickyNotes", "sticky_notes"]),
    },
    {
      kind: "ink-annotations",
      count: firstArrayLength(safeContent, [
        "inkAnnotations",
        "ink_annotations",
      ]),
    },
    {
      kind: "signatures",
      count: firstArrayLength(safeContent, [
        "signatureAnnotations",
        "signature_annotations",
      ]),
    },
    {
      kind: "redactions",
      count: firstArrayLength(safeContent, [
        "redactionAnnotations",
        "redaction_annotations",
      ]),
    },
    {
      kind: "reader-annotations",
      count: boundedCount(readerAnnotationCount),
    },
  ];
  const entries = candidates.filter((entry) => entry.count > 0);
  return {
    entries,
    total: entries.reduce((total, entry) => total + entry.count, 0),
  };
}

export function buildPdfSourceReplacementAudit(options: {
  replacedAt?: string;
  previous: {
    dataItemId: string;
    filename: string;
    sha256?: string | null;
    riskProfileDigest?: string | null;
    sizeBytes?: number | null;
    mediaType?: string | null;
    contentVersion?: number | null;
    pageVersionId?: string | null;
    pageVersionNo?: number | null;
  };
  replacement: {
    filename: string;
    sha256?: string | null;
    riskProfileDigest?: string | null;
    sizeBytes: number;
    mediaType?: string | null;
  };
  lossDigest?: string | null;
  lossManifest: PdfSourceReplacementLossManifest;
}): PdfSourceReplacementAudit {
  const losses = options.lossManifest.entries.map((entry) => ({ ...entry }));
  return {
    kind: PDF_SOURCE_REPLACEMENT_AUDIT_KIND,
    replacedAt: options.replacedAt ?? new Date().toISOString(),
    sourceRetained: true,
    previous: {
      dataItemId: options.previous.dataItemId,
      filename: options.previous.filename,
      sha256: normalizeOptionalSha256(options.previous.sha256),
      riskProfileDigest: normalizeOptionalRiskProfileDigest(
        options.previous.riskProfileDigest,
      ),
      sizeBytes: options.previous.sizeBytes ?? null,
      mediaType: options.previous.mediaType ?? null,
      contentVersion: options.previous.contentVersion ?? null,
      pageVersionId: options.previous.pageVersionId ?? null,
      pageVersionNo: options.previous.pageVersionNo ?? null,
    },
    replacement: {
      filename: options.replacement.filename,
      sha256: normalizeOptionalSha256(options.replacement.sha256),
      riskProfileDigest: normalizeOptionalRiskProfileDigest(
        options.replacement.riskProfileDigest,
      ),
      sizeBytes: Math.max(0, Math.floor(options.replacement.sizeBytes)),
      mediaType: options.replacement.mediaType || null,
    },
    lossDigest: normalizeOptionalLossDigest(options.lossDigest),
    losses,
    totalLossEntries: losses.reduce((total, entry) => total + entry.count, 0),
  };
}

function parseLossEntries(
  value: unknown,
): PdfSourceReplacementLossEntry[] | null {
  if (!Array.isArray(value) || value.length > LOSS_KINDS.size) return null;
  const entries: PdfSourceReplacementLossEntry[] = [];
  const seen = new Set<PdfSourceReplacementLossKind>();
  for (const item of value) {
    if (!item || typeof item !== "object") return null;
    const row = item as Record<string, unknown>;
    if (
      typeof row.kind !== "string" ||
      !LOSS_KINDS.has(row.kind as PdfSourceReplacementLossKind) ||
      seen.has(row.kind as PdfSourceReplacementLossKind)
    ) {
      return null;
    }
    const count = boundedCount(row.count);
    if (count <= 0) return null;
    const kind = row.kind as PdfSourceReplacementLossKind;
    seen.add(kind);
    entries.push({ kind, count });
  }
  return entries;
}

/** Strictly reads the bounded audit shown after a page is reopened. */
export function readPdfSourceReplacementAudit(
  content: Record<string, unknown> | undefined,
): PdfSourceReplacementAudit | null {
  const source =
    content?.source && typeof content.source === "object"
      ? (content.source as Record<string, unknown>)
      : null;
  const value =
    source?.replacementAudit && typeof source.replacementAudit === "object"
      ? (source.replacementAudit as Record<string, unknown>)
      : null;
  if (
    !value ||
    value.kind !== PDF_SOURCE_REPLACEMENT_AUDIT_KIND ||
    value.sourceRetained !== true
  ) {
    return null;
  }
  const previous =
    value.previous && typeof value.previous === "object"
      ? (value.previous as Record<string, unknown>)
      : null;
  const replacement =
    value.replacement && typeof value.replacement === "object"
      ? (value.replacement as Record<string, unknown>)
      : null;
  const replacedAt = boundedString(value.replacedAt);
  const previousDataItemId = boundedString(previous?.dataItemId);
  const previousFilename = boundedString(previous?.filename);
  const previousPageVersionId = nullableBoundedString(previous?.pageVersionId);
  const previousPageVersionNo = nullableNonNegativeInteger(
    previous?.pageVersionNo,
  );
  const replacementFilename = boundedString(replacement?.filename);
  const replacementSizeBytes = nullableNonNegativeInteger(
    replacement?.sizeBytes,
  );
  const previousSha256 = parseOptionalBoundIdentity(
    previous?.sha256,
    normalizeOptionalSha256,
  );
  const previousRiskProfileDigest = parseOptionalBoundIdentity(
    previous?.riskProfileDigest,
    normalizeOptionalRiskProfileDigest,
  );
  const replacementSha256 = parseOptionalBoundIdentity(
    replacement?.sha256,
    normalizeOptionalSha256,
  );
  const replacementRiskProfileDigest = parseOptionalBoundIdentity(
    replacement?.riskProfileDigest,
    normalizeOptionalRiskProfileDigest,
  );
  const lossDigest = parseOptionalBoundIdentity(
    value.lossDigest,
    normalizeOptionalLossDigest,
  );
  const losses = parseLossEntries(value.losses);
  if (
    !replacedAt ||
    !previousDataItemId ||
    !previousFilename ||
    !replacementFilename ||
    replacementSizeBytes === null ||
    !previousSha256.valid ||
    !previousRiskProfileDigest.valid ||
    !replacementSha256.valid ||
    !replacementRiskProfileDigest.valid ||
    !lossDigest.valid ||
    losses === null
  ) {
    return null;
  }
  if (
    (previousPageVersionId === null) !== (previousPageVersionNo === null) ||
    previousPageVersionNo === 0 ||
    (previousSha256.value === null) !==
      (previousRiskProfileDigest.value === null) ||
    (replacementSha256.value === null) !==
      (replacementRiskProfileDigest.value === null)
  ) {
    return null;
  }
  const totalLossEntries = losses.reduce(
    (total, entry) => total + entry.count,
    0,
  );
  if (boundedTotalCount(value.totalLossEntries) !== totalLossEntries) {
    return null;
  }
  return {
    kind: PDF_SOURCE_REPLACEMENT_AUDIT_KIND,
    replacedAt,
    sourceRetained: true,
    previous: {
      dataItemId: previousDataItemId,
      filename: previousFilename,
      sha256: previousSha256.value,
      riskProfileDigest: previousRiskProfileDigest.value,
      sizeBytes: nullableNonNegativeInteger(previous?.sizeBytes),
      mediaType: nullableBoundedString(previous?.mediaType),
      contentVersion: nullableNonNegativeInteger(previous?.contentVersion),
      pageVersionId: previousPageVersionId,
      pageVersionNo: previousPageVersionNo,
    },
    replacement: {
      filename: replacementFilename,
      sha256: replacementSha256.value,
      riskProfileDigest: replacementRiskProfileDigest.value,
      sizeBytes: replacementSizeBytes,
      mediaType: nullableBoundedString(replacement?.mediaType),
    },
    lossDigest: lossDigest.value,
    losses,
    totalLossEntries,
  };
}
