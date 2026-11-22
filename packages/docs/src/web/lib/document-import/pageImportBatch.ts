"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiFetchRaw, apiGet, isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import { preloadPageEditorCode } from "@mrnote/core/web/lib/editor-code-preload.ts";
import { desktopLocalUploadPart, linkDesktopOriginal } from "@mrnote/core/web/lib/desktop-local-files.ts";
import {
  PageImportProjectionBudget,
  PageImportProjectionBudgetError,
} from "./pageImportProjectionBudget";

const IMPORT_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/;
const UUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const MAX_BATCH_ITEMS = 20;
const MAX_RECOVERY_POLLS = 8;
const DEFAULT_RECOVERY_DELAY_MS = 250;
/** How long a probe waits on a batch the server is still converting. A long
 * document's conversion outlives the request that started it (a proxy cuts
 * the connection, the server carries on and commits); the page is waited
 * for, not reported as a failed import. The server's lease is two hours. */
const LIVE_OWNER_WAIT_MS = 45 * 60_000;
const LIVE_OWNER_MIN_POLL_MS = 250;
const LIVE_OWNER_MAX_POLL_MS = 10_000;
const PENDING_BATCH_STORAGE_PREFIX = "mrai.page-import-batch.pending.v1";
// The server retains terminal receipts for 24 hours. Expire the browser hint
// one hour earlier so a stored id is never intentionally reused at the exact
// boundary where its authoritative receipt can be purged.
const PENDING_BATCH_MAX_AGE_MS = 23 * 60 * 60 * 1_000;

interface PendingBatchIdentity {
  storageKey: string;
  notebookHash: string;
  manifestHash: string;
  filesHash: string;
}

interface PendingBatchRecord {
  version: 1;
  notebook_hash: string;
  manifest_hash: string;
  files_hash: string;
  batch_id: string;
  created_at_ms: number;
}

export type PageImportBatchPageType =
  | "document"
  | "spreadsheet"
  | "presentation"
  | "pdf"
  | "code";

export interface PageImportBatchFileItem {
  client_id: string;
  operation_id: string;
}

export interface PageImportBatchAppendOperation {
  operation_id: string;
  kind: "append_to_page";
  page_id: string;
  base_content_version: number;
  title?: string | null;
  nodes: Array<Record<string, unknown>>;
}

export interface PageImportBatchCreateOperation {
  operation_id: string;
  kind: "create_page";
  title: string;
  page_type: PageImportBatchPageType;
  content_json: Record<string, unknown>;
  sort_order: number;
}

export type PageImportSourceKind = "pdf" | "docx" | "xlsx" | "pptx";

export interface PageImportBatchSourceCreateOperation {
  operation_id: string;
  kind: "create_source_page";
  title: string;
  source_kind: PageImportSourceKind;
  sort_order: number;
}

export type PageImportBatchOperation =
  | PageImportBatchAppendOperation
  | PageImportBatchCreateOperation
  | PageImportBatchSourceCreateOperation;

export interface PageImportBatchManifest {
  items: PageImportBatchFileItem[];
  operations: PageImportBatchOperation[];
  /** The sidebar folder every created page lands in (absent: the root). */
  folder_id?: string;
}

export interface PageImportBatchAttachmentReceipt {
  client_id: string;
  attachment_id: string;
  filename: string;
  mime_type: string;
  size_bytes: number;
  attachment_type: string;
}

export interface PageImportBatchOperationReceipt {
  operation_id: string;
  kind: PageImportBatchOperation["kind"];
  page_id: string;
  content_version: number;
  attachment_ids: string[];
  source_client_id?: string;
  data_item_id?: string;
  resource_id?: string;
  page_type?: PageImportBatchPageType;
}

export interface PageImportBatchReceipt {
  batch_id: string;
  status: "preparing" | "committed" | "failed";
  attachments: PageImportBatchAttachmentReceipt[];
  items: PageImportBatchOperationReceipt[];
  error_code: string | null;
}

export class PageImportBatchProtocolError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PageImportBatchProtocolError";
  }
}

export class PageImportBatchFailedError extends Error {
  constructor(readonly errorCode: string | null) {
    super(errorCode || "The import batch failed.");
    this.name = "PageImportBatchFailedError";
  }
}

export class PageImportBatchLiveOwnerError extends Error {
  constructor(
    readonly batchId: string,
    readonly retryAfterMs: number,
    readonly requestId: string | null,
  ) {
    super("The import batch is still owned by a live request.");
    this.name = "PageImportBatchLiveOwnerError";
  }
}

interface SubmitPageImportBatchInput {
  notebookId: string;
  batchId: string;
  manifest: PageImportBatchManifest;
  files: readonly File[];
  signal?: AbortSignal;
  /** Test seam and UI recovery tuning; production defaults to 250 ms. */
  recoveryDelayMs?: number;
  /** Told the batch id the PUT goes out under (a recovered or replaced id
   * included), once, as it goes out: the page a source operation makes is
   * named by it (the API's ``_deterministic_import_page_id``). */
  onSubmitting?: (batchId: string) => void;
}

function protocolError(message: string): never {
  throw new PageImportBatchProtocolError(message);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function canonicalJsonStringify(value: unknown): string {
  const ancestors = new WeakSet<object>();
  const serialize = (
    current: unknown,
    inArray: boolean,
  ): string | undefined => {
    if (current === null) return "null";
    if (typeof current === "string" || typeof current === "boolean") {
      return JSON.stringify(current);
    }
    if (typeof current === "number") {
      return Number.isFinite(current) ? JSON.stringify(current) : "null";
    }
    if (
      current === undefined ||
      typeof current === "function" ||
      typeof current === "symbol"
    ) {
      return inArray ? "null" : undefined;
    }
    if (typeof current === "bigint") {
      throw new TypeError("BigInt is not valid JSON");
    }
    if (typeof current !== "object") return undefined;
    if (ancestors.has(current)) {
      throw new TypeError("Converting circular structure to JSON");
    }
    ancestors.add(current);
    try {
      if (Array.isArray(current)) {
        return `[${current
          .map((item) => serialize(item, true) ?? "null")
          .join(",")}]`;
      }
      const fields: string[] = [];
      for (const key of Object.keys(current).sort()) {
        const encoded = serialize(
          (current as Record<string, unknown>)[key],
          false,
        );
        if (encoded !== undefined) {
          fields.push(`${JSON.stringify(key)}:${encoded}`);
        }
      }
      return `{${fields.join(",")}}`;
    } finally {
      ancestors.delete(current);
    }
  };
  const encoded = serialize(value, false);
  if (encoded === undefined) {
    throw new TypeError("Value is not valid JSON");
  }
  return encoded;
}

function stableFingerprint(value: string): string {
  // This is an identity/cache hash, not a security decision. The server's
  // byte-derived SHA-256 remains authoritative before any replay is accepted.
  let left = 0x811c9dc5;
  let right = 0x9e3779b9;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    left = Math.imul(left ^ code, 0x01000193) >>> 0;
    right = Math.imul(right ^ code, 0x85ebca6b) >>> 0;
  }
  return `${left.toString(16).padStart(8, "0")}${right
    .toString(16)
    .padStart(8, "0")}`;
}

function pendingBatchIdentity(
  notebookId: string,
  manifestJson: string,
  files: readonly File[],
): PendingBatchIdentity {
  const notebookHash = stableFingerprint(notebookId);
  const manifestHash = stableFingerprint(manifestJson);
  const fileMetadata = files.map((file) => ({
    last_modified: Number.isSafeInteger(file.lastModified)
      ? file.lastModified
      : 0,
    name: file.name,
    size: file.size,
    type: file.type.toLowerCase(),
  }));
  const filesHash = stableFingerprint(canonicalJsonStringify(fileMetadata));
  return {
    storageKey: `${PENDING_BATCH_STORAGE_PREFIX}:${notebookHash}:${manifestHash}:${filesHash}`,
    notebookHash,
    manifestHash,
    filesHash,
  };
}

function browserSessionStorage(): Storage | null {
  if (typeof window === "undefined") return null;
  try {
    return window.sessionStorage;
  } catch {
    return null;
  }
}

function loadPendingBatchId(identity: PendingBatchIdentity): string | null {
  const storage = browserSessionStorage();
  if (!storage) return null;
  try {
    const raw = storage.getItem(identity.storageKey);
    if (!raw) return null;
    const record = JSON.parse(raw) as unknown;
    if (
      !isRecord(record) ||
      record.version !== 1 ||
      record.notebook_hash !== identity.notebookHash ||
      record.manifest_hash !== identity.manifestHash ||
      record.files_hash !== identity.filesHash ||
      typeof record.batch_id !== "string" ||
      !UUID_PATTERN.test(record.batch_id) ||
      !Number.isSafeInteger(record.created_at_ms) ||
      Date.now() - Number(record.created_at_ms) < 0 ||
      Date.now() - Number(record.created_at_ms) > PENDING_BATCH_MAX_AGE_MS
    ) {
      storage.removeItem(identity.storageKey);
      return null;
    }
    return record.batch_id;
  } catch {
    try {
      storage.removeItem(identity.storageKey);
    } catch {
      // Storage can become unavailable between reads (privacy mode/quota).
    }
    return null;
  }
}

function storePendingBatch(
  identity: PendingBatchIdentity,
  batchId: string,
): void {
  const storage = browserSessionStorage();
  if (!storage) return;
  const record: PendingBatchRecord = {
    version: 1,
    notebook_hash: identity.notebookHash,
    manifest_hash: identity.manifestHash,
    files_hash: identity.filesHash,
    batch_id: batchId,
    created_at_ms: Date.now(),
  };
  try {
    storage.setItem(identity.storageKey, JSON.stringify(record));
  } catch {
    // Idempotency still holds for this invocation; only refresh recovery is
    // unavailable when the browser denies session storage.
  }
}

function clearPendingBatch(
  identity: PendingBatchIdentity,
  batchId: string,
): void {
  const storage = browserSessionStorage();
  if (!storage) return;
  try {
    const current = loadPendingBatchId(identity);
    if (current === batchId) storage.removeItem(identity.storageKey);
  } catch {
    // Best effort only; stale records are age-bounded and validated on read.
  }
}

function requireBoundedString(
  value: unknown,
  field: string,
  maxLength = 512,
): string {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > maxLength
  ) {
    protocolError(`Invalid import batch ${field}.`);
  }
  return value;
}

function validateManifest(
  manifest: PageImportBatchManifest,
  files: readonly File[],
): string {
  if (!isRecord(manifest)) {
    protocolError("The import batch manifest must be an object.");
  }
  if (
    !Array.isArray(manifest.items) ||
    manifest.items.length > MAX_BATCH_ITEMS ||
    manifest.items.length !== files.length
  ) {
    protocolError(
      "The import batch file manifest does not match its multipart files.",
    );
  }
  if (
    !Array.isArray(manifest.operations) ||
    manifest.operations.length === 0 ||
    manifest.operations.length > MAX_BATCH_ITEMS
  ) {
    protocolError(
      "The import batch must contain between one and twenty operations.",
    );
  }

  const operationKinds = new Map<string, PageImportBatchOperation["kind"]>();
  const appendTargets = new Set<string>();
  const projectionBudget = new PageImportProjectionBudget();
  for (const operation of manifest.operations) {
    if (
      !isRecord(operation) ||
      !IMPORT_ID_PATTERN.test(String(operation.operation_id))
    ) {
      protocolError("The import batch contains an invalid operation id.");
    }
    const operationId = String(operation.operation_id);
    if (operationKinds.has(operationId)) {
      protocolError("The import batch contains duplicate operation ids.");
    }
    if (
      operation.kind !== "append_to_page" &&
      operation.kind !== "create_page" &&
      operation.kind !== "create_source_page"
    ) {
      protocolError("The import batch contains an unknown operation kind.");
    }
    operationKinds.set(operationId, operation.kind);
    if (operation.kind === "append_to_page") {
      const target = requireBoundedString(operation.page_id, "page id", 128);
      if (appendTargets.has(target)) {
        protocolError("A page can be appended to only once per import batch.");
      }
      appendTargets.add(target);
      if (
        !Number.isSafeInteger(operation.base_content_version) ||
        operation.base_content_version < 1 ||
        !Array.isArray(operation.nodes) ||
        operation.nodes.length === 0
      ) {
        protocolError("The import batch append operation is invalid.");
      }
      try {
        projectionBudget.consumeValue(operation.nodes);
      } catch (error) {
        if (error instanceof PageImportProjectionBudgetError) {
          protocolError("The combined import projection is too large.");
        }
        throw error;
      }
    } else if (
      operation.kind === "create_page" &&
      !isRecord(operation.content_json)
    ) {
      protocolError("The import batch create operation has invalid content.");
    } else if (operation.kind === "create_page") {
      try {
        projectionBudget.consumeValue(operation.content_json);
      } catch (error) {
        if (error instanceof PageImportProjectionBudgetError) {
          protocolError("The combined import projection is too large.");
        }
        throw error;
      }
    } else if (
      typeof operation.title !== "string" ||
      operation.title.length > 500 ||
      !["pdf", "docx", "xlsx", "pptx"].includes(operation.source_kind) ||
      !Number.isSafeInteger(operation.sort_order)
    ) {
      protocolError("The import batch source operation is invalid.");
    }
  }

  const clientIds = new Set<string>();
  for (const item of manifest.items) {
    if (
      !isRecord(item) ||
      !IMPORT_ID_PATTERN.test(String(item.client_id)) ||
      !IMPORT_ID_PATTERN.test(String(item.operation_id))
    ) {
      protocolError("The import batch contains an invalid file item.");
    }
    if (clientIds.has(item.client_id)) {
      protocolError("The import batch contains duplicate file ids.");
    }
    if (!operationKinds.has(item.operation_id)) {
      protocolError("An import file names an unknown operation.");
    }
    clientIds.add(item.client_id);
  }

  const sourceOperations = manifest.operations.filter(
    (operation) => operation.kind === "create_source_page",
  );
  if (
    sourceOperations.length > 0 &&
    (manifest.operations.length !== 1 ||
      manifest.items.length !== 1 ||
      manifest.items[0]?.operation_id !== sourceOperations[0]?.operation_id)
  ) {
    protocolError(
      "A source-custody import must contain exactly one operation and one file.",
    );
  }

  try {
    return canonicalJsonStringify(manifest);
  } catch {
    protocolError("The import batch manifest is not serializable.");
  }
}

function validateReceipt(
  payload: unknown,
  options: {
    expectedBatchId: string;
    manifest: PageImportBatchManifest;
  },
): PageImportBatchReceipt {
  const { expectedBatchId, manifest } = options;
  if (!isRecord(payload) || payload.batch_id !== expectedBatchId) {
    protocolError(
      "The import batch response does not match the submitted batch.",
    );
  }
  if (
    payload.status !== "preparing" &&
    payload.status !== "committed" &&
    payload.status !== "failed"
  ) {
    protocolError("The import batch response has an invalid status.");
  }
  if (
    !Array.isArray(payload.attachments) ||
    payload.attachments.length > MAX_BATCH_ITEMS ||
    !Array.isArray(payload.items) ||
    payload.items.length > MAX_BATCH_ITEMS
  ) {
    protocolError("The import batch response has invalid result arrays.");
  }
  if (
    payload.error_code !== null &&
    (typeof payload.error_code !== "string" || payload.error_code.length > 120)
  ) {
    protocolError("The import batch response has an invalid error code.");
  }

  const expectedOperations = new Map(
    manifest.operations.map((operation) => [operation.operation_id, operation]),
  );
  const expectedFiles = new Set(manifest.items.map((item) => item.client_id));
  const sourceFileIds = new Set(
    manifest.items
      .filter(
        (item) =>
          expectedOperations.get(item.operation_id)?.kind ===
          "create_source_page",
      )
      .map((item) => item.client_id),
  );
  const seenFiles = new Set<string>();
  const attachmentIdsByClientId = new Map<string, string>();
  const seenAttachmentIds = new Set<string>();
  for (const attachment of payload.attachments) {
    if (!isRecord(attachment)) {
      protocolError("The import batch response has an invalid attachment.");
    }
    const clientId = requireBoundedString(
      attachment.client_id,
      "client id",
      64,
    );
    const attachmentId = requireBoundedString(
      attachment.attachment_id,
      "attachment id",
      256,
    );
    requireBoundedString(attachment.filename, "filename", 1024);
    requireBoundedString(attachment.mime_type, "MIME type", 256);
    requireBoundedString(attachment.attachment_type, "attachment type", 64);
    if (
      seenFiles.has(clientId) ||
      seenAttachmentIds.has(attachmentId) ||
      !expectedFiles.has(clientId) ||
      sourceFileIds.has(clientId) ||
      !Number.isSafeInteger(attachment.size_bytes) ||
      Number(attachment.size_bytes) < 0
    ) {
      protocolError(
        "The import batch response has an invalid attachment mapping.",
      );
    }
    seenFiles.add(clientId);
    seenAttachmentIds.add(attachmentId);
    attachmentIdsByClientId.set(clientId, attachmentId);
  }

  const seenOperations = new Set<string>();
  const seenPageIds = new Set<string>();
  for (const item of payload.items) {
    if (!isRecord(item)) {
      protocolError(
        "The import batch response has an invalid operation result.",
      );
    }
    const operationId = requireBoundedString(
      item.operation_id,
      "operation id",
      64,
    );
    const pageId = requireBoundedString(item.page_id, "page id", 256);
    const expectedOperation = expectedOperations.get(operationId);
    const expectedContentVersion =
      expectedOperation?.kind === "append_to_page"
        ? expectedOperation.base_content_version + 1
        : expectedOperation?.kind === "create_page"
          ? 1
          : expectedOperation?.kind === "create_source_page"
            ? 1
            : null;
    const attachmentIds = Array.isArray(item.attachment_ids)
      ? item.attachment_ids
      : [];
    if (
      seenOperations.has(operationId) ||
      seenPageIds.has(pageId) ||
      expectedOperation?.kind !== item.kind ||
      !Number.isSafeInteger(item.content_version) ||
      item.content_version !== expectedContentVersion ||
      !Array.isArray(item.attachment_ids) ||
      attachmentIds.length > MAX_BATCH_ITEMS ||
      attachmentIds.some(
        (value) =>
          typeof value !== "string" || value.length === 0 || value.length > 256,
      ) ||
      new Set(attachmentIds).size !== attachmentIds.length
    ) {
      protocolError(
        "The import batch response has an invalid operation mapping.",
      );
    }
    if (
      expectedOperation?.kind === "append_to_page" &&
      expectedOperation.page_id !== pageId
    ) {
      protocolError(
        "The import batch response targets a different page than requested.",
      );
    }
    if (expectedOperation?.kind === "create_source_page") {
      const expectedSourceItem = manifest.items.find(
        (manifestItem) => manifestItem.operation_id === operationId,
      );
      const expectedPageType: PageImportBatchPageType =
        expectedOperation.source_kind === "pdf"
          ? "pdf"
          : expectedOperation.source_kind === "docx"
            ? "document"
            : expectedOperation.source_kind === "xlsx"
              ? "spreadsheet"
              : "presentation";
      const sourceClientId = requireBoundedString(
        item.source_client_id,
        "source client id",
        64,
      );
      requireBoundedString(item.data_item_id, "data item id", 256);
      requireBoundedString(item.resource_id, "resource id", 256);
      if (
        expectedSourceItem?.client_id !== sourceClientId ||
        item.page_type !== expectedPageType ||
        attachmentIds.length !== 0 ||
        seenFiles.has(sourceClientId)
      ) {
        protocolError(
          "The import batch response has an invalid source-custody mapping.",
        );
      }
      seenFiles.add(sourceClientId);
    } else if (
      "source_client_id" in item ||
      "data_item_id" in item ||
      "resource_id" in item ||
      "page_type" in item
    ) {
      protocolError(
        "A non-source import result cannot contain source-custody identities.",
      );
    }
    const expectedAttachmentIds = manifest.items
      .filter((manifestItem) => manifestItem.operation_id === operationId)
      .map((manifestItem) =>
        attachmentIdsByClientId.get(manifestItem.client_id),
      )
      .filter((value): value is string => typeof value === "string")
      .sort();
    const actualAttachmentIds = [...attachmentIds].sort();
    if (
      expectedAttachmentIds.length !== actualAttachmentIds.length ||
      expectedAttachmentIds.some(
        (attachmentId, index) => attachmentId !== actualAttachmentIds[index],
      )
    ) {
      protocolError(
        "The import batch response assigns an attachment to the wrong operation.",
      );
    }
    seenOperations.add(operationId);
    seenPageIds.add(pageId);
  }

  if (payload.status === "committed") {
    if (
      seenFiles.size !== expectedFiles.size ||
      seenOperations.size !== expectedOperations.size ||
      payload.error_code !== null
    ) {
      protocolError("The committed import batch receipt is incomplete.");
    }
  } else if (payload.attachments.length !== 0 || payload.items.length !== 0) {
    protocolError(
      "An unfinished import batch cannot contain committed results.",
    );
  }
  return payload as unknown as PageImportBatchReceipt;
}

function importBatchPath(notebookId: string, batchId: string): string {
  return `/api/v1/notebooks/${encodeURIComponent(notebookId)}/import-batches/${batchId}`;
}

async function buildMultipart(
  manifestJson: string,
  files: readonly File[],
): Promise<FormData> {
  const form = new FormData();
  form.append("manifest", manifestJson);
  // MRNote Classic: a file on this computer travels as a token its API
  // reads where it is (lib/desktop-local-files); elsewhere the file itself.
  for (const file of await Promise.all(files.map(desktopLocalUploadPart))) {
    form.append("files", file);
  }
  return form;
}

/** MRNote Classic: each page made from a file on this computer is linked to
 * that file, so Save writes the page back to it. */
function linkLocalOriginals(
  notebookId: string,
  manifest: PageImportBatchManifest,
  files: readonly File[],
  receipt: PageImportBatchReceipt,
): void {
  for (const item of receipt.items) {
    if (item.kind !== "create_source_page" || !item.source_client_id) continue;
    const index = manifest.items.findIndex((entry) => entry.client_id === item.source_client_id);
    const file = index >= 0 ? files[index] : undefined;
    if (file) linkDesktopOriginal(item.page_id, file, notebookId);
  }
}

function isAmbiguousMutationFailure(error: unknown): boolean {
  if (error instanceof DOMException && error.name === "AbortError") return true;
  if (!isApiRequestError(error)) return false;
  return (
    error.status === 0 ||
    error.status >= 500 ||
    (error.status === 409 &&
      ["import_batch_in_progress", "import_batch_superseded"].includes(
        error.code || "",
      ))
  );
}

function isRecoveredBatchIdentityConflict(error: unknown): boolean {
  if (!isApiRequestError(error)) return false;
  if (error.status === 404) return true;
  return (
    error.status === 409 &&
    [
      "import_batch_idempotency_conflict",
      "import_batch_id_unavailable",
    ].includes(error.code || "")
  );
}

type ImportBatchOwnerProof = {
  ownerState: "live" | "orphan";
  retryAfterMs: number;
  requestId: string | null;
};

function boundedRequestId(value: unknown): string | null {
  return typeof value === "string" &&
    value.length > 0 &&
    value.length <= 128 &&
    value.trim() === value &&
    !/[\u0000-\u001f\u007f]/.test(value)
    ? value
    : null;
}

/**
 * Parse only the authenticated GET-status owner proof. An arbitrary 409 body
 * must never authorize replay: the batch id, stable code, bounded retry delay,
 * and exact owner-state vocabulary all have to match this request.
 */
function requestIdFromApiError(error: unknown): string | null {
  if (!isApiRequestError(error)) return null;
  const topLevelRequestId = (error as unknown as { requestId?: unknown })
    .requestId;
  const detailsRequestId = error.details?.request_id;
  if (
    (topLevelRequestId !== undefined &&
      boundedRequestId(topLevelRequestId) === null) ||
    (detailsRequestId !== undefined &&
      boundedRequestId(detailsRequestId) === null)
  ) {
    return null;
  }
  return (
    boundedRequestId(topLevelRequestId) ?? boundedRequestId(detailsRequestId)
  );
}

/**
 * Owner state is authoritative only when it arrives in the authenticated
 * status GET payload. A PUT conflict body is useful for a display-only request
 * id, but it can never authorize replay by itself.
 */
function importBatchOwnerProof(
  payload: unknown,
  receipt: PageImportBatchReceipt,
  options: {
    fallbackRetryAfterMs: number;
    requestId: string | null;
  },
): ImportBatchOwnerProof | null {
  if (!isRecord(payload)) return null;
  const ownerState = payload.owner_state;
  if (ownerState === undefined) {
    if ("retry_after_ms" in payload || "request_id" in payload) {
      protocolError("The import batch owner status is incomplete.");
    }
    return null;
  }
  if (
    receipt.status !== "preparing" ||
    (ownerState !== "live" && ownerState !== "orphan")
  ) {
    protocolError("The import batch response has an invalid owner status.");
  }

  const rawRetryAfterMs = payload.retry_after_ms;
  if (
    rawRetryAfterMs !== undefined &&
    (!Number.isSafeInteger(rawRetryAfterMs) ||
      Number(rawRetryAfterMs) < 0 ||
      Number(rawRetryAfterMs) > 60_000)
  ) {
    protocolError("The import batch response has an invalid retry delay.");
  }
  const rawRequestId = payload.request_id;
  if (
    rawRequestId !== undefined &&
    boundedRequestId(rawRequestId) === null
  ) {
    protocolError("The import batch response has an invalid request id.");
  }
  return {
    ownerState,
    retryAfterMs:
      rawRetryAfterMs === undefined
        ? options.fallbackRetryAfterMs
        : Number(rawRetryAfterMs),
    requestId:
      options.requestId ??
      (rawRequestId === undefined ? null : boundedRequestId(rawRequestId)),
  };
}

function wait(delayMs: number): Promise<void> {
  if (delayMs <= 0) return Promise.resolve();
  return new Promise((resolve) => window.setTimeout(resolve, delayMs));
}

type ProbeResult =
  | { kind: "receipt"; receipt: PageImportBatchReceipt }
  | { kind: "orphan" }
  | { kind: "not_found" }
  | { kind: "unknown" };

async function probeImportBatch(
  path: string,
  options: {
    expectedBatchId: string;
    manifest: PageImportBatchManifest;
    delayMs: number;
    requestId?: string | null;
    /** Stops waiting on a live owner (the server still finishes the batch). */
    signal?: AbortSignal;
  },
): Promise<ProbeResult> {
  const { expectedBatchId, manifest, delayMs } = options;
  let everyProbeWasNotFound = true;
  const liveOwnerDeadline = Date.now() + LIVE_OWNER_WAIT_MS;
  for (let attempt = 0; attempt < MAX_RECOVERY_POLLS; attempt += 1) {
    try {
      const payload = await apiGet<unknown>(path);
      const receipt = validateReceipt(payload, {
        expectedBatchId,
        manifest,
      });
      everyProbeWasNotFound = false;
      const ownerProof = importBatchOwnerProof(payload, receipt, {
        fallbackRetryAfterMs: delayMs,
        requestId: options.requestId ?? null,
      });
      if (ownerProof?.ownerState === "orphan") {
        return { kind: "orphan" };
      }
      if (
        ownerProof?.ownerState === "live" &&
        Date.now() < liveOwnerDeadline &&
        !options.signal?.aborted
      ) {
        await wait(
          Math.min(
            LIVE_OWNER_MAX_POLL_MS,
            Math.max(LIVE_OWNER_MIN_POLL_MS, ownerProof.retryAfterMs),
          ),
        );
        // The server is converting it: a wait spends no recovery probe.
        attempt -= 1;
        continue;
      }
      if (ownerProof?.ownerState === "live") {
        throw new PageImportBatchLiveOwnerError(
          expectedBatchId,
          ownerProof.retryAfterMs,
          ownerProof.requestId,
        );
      }
      if (receipt.status === "committed") {
        return { kind: "receipt", receipt };
      }
      if (receipt.status === "failed") {
        throw new PageImportBatchFailedError(receipt.error_code);
      }
    } catch (error) {
      if (error instanceof PageImportBatchProtocolError) throw error;
      if (error instanceof PageImportBatchFailedError) throw error;
      if (error instanceof PageImportBatchLiveOwnerError) throw error;
      if (!isApiRequestError(error) || error.status !== 404) {
        everyProbeWasNotFound = false;
      }
    }
    if (attempt + 1 < MAX_RECOVERY_POLLS) {
      await wait(delayMs);
    }
  }
  return everyProbeWasNotFound ? { kind: "not_found" } : { kind: "unknown" };
}

/**
 * Commit one notebook import as a durable idempotent server transaction.
 *
 * A lost/aborted PUT acknowledgement is always probed without the caller's
 * cancelled signal. Only a proven 404 permits a second PUT, and that replay
 * uses the same batch id, manifest, file order and bytes.
 */
export async function submitPageImportBatch(
  input: SubmitPageImportBatchInput,
): Promise<PageImportBatchReceipt> {
  if (!UUID_PATTERN.test(input.batchId)) {
    protocolError("The import batch id must be a UUID.");
  }
  const notebookId = requireBoundedString(input.notebookId, "notebook id", 256);
  const manifestJson = validateManifest(input.manifest, input.files);
  // The editors the imported files open in load while the import runs.
  for (const operation of input.manifest.operations) {
    if (operation.kind === "create_source_page") preloadPageEditorCode(operation.source_kind);
  }
  const pendingIdentity = pendingBatchIdentity(
    notebookId,
    manifestJson,
    input.files,
  );
  const recoveredBatchId = loadPendingBatchId(pendingIdentity);
  let effectiveBatchId = recoveredBatchId || input.batchId;
  let requiresRecoveredByteVerification = recoveredBatchId !== null;
  let path = importBatchPath(notebookId, effectiveBatchId);
  if (!recoveredBatchId) {
    storePendingBatch(pendingIdentity, effectiveBatchId);
  }
  const delayMs = Math.max(
    0,
    Math.min(2_000, input.recoveryDelayMs ?? DEFAULT_RECOVERY_DELAY_MS),
  );

  const committed = (receipt: PageImportBatchReceipt) => {
    clearPendingBatch(pendingIdentity, effectiveBatchId);
    return receipt;
  };

  const replaceRecoveredBatchId = () => {
    clearPendingBatch(pendingIdentity, effectiveBatchId);
    effectiveBatchId =
      input.batchId === effectiveBatchId
        ? createPageImportBatchId()
        : input.batchId;
    path = importBatchPath(notebookId, effectiveBatchId);
    storePendingBatch(pendingIdentity, effectiveBatchId);
    requiresRecoveredByteVerification = false;
  };

  // A same-selection retry after navigation/refresh must first inspect the
  // durable operation it already started. A terminal failure releases the id.
  // A committed GET is only a candidate: the storage key intentionally uses
  // cheap file metadata, so the selected bytes still must pass a same-id PUT
  // and the server's authoritative SHA-256 comparison before we may return it.
  if (recoveredBatchId) {
    try {
      await probeImportBatch(path, {
        expectedBatchId: effectiveBatchId,
        manifest: input.manifest,
        delayMs,
        signal: input.signal,
      });
    } catch (error) {
      if (!(error instanceof PageImportBatchFailedError)) throw error;
      replaceRecoveredBatchId();
    }
  }

  let lastError: unknown;
  try {
    let putAttempt = 0;
    while (putAttempt < 2) {
      try {
        if (putAttempt === 0 && input.onSubmitting) {
          try { input.onSubmitting(effectiveBatchId); } catch { /* an observer never stops the import */ }
        }
        const response = await apiFetchRaw(
          path,
          {
            method: "PUT",
            body: await buildMultipart(manifestJson, input.files),
            signal: input.signal,
          },
          { contentType: "" },
        );
        const receipt = validateReceipt(await response.json(), {
          expectedBatchId: effectiveBatchId,
          manifest: input.manifest,
        });
        if (receipt.status !== "committed") {
          protocolError(
            "The import batch PUT did not return a committed receipt.",
          );
        }
        linkLocalOriginals(notebookId, input.manifest, input.files, receipt);
        return committed(receipt);
      } catch (error) {
        if (error instanceof PageImportBatchProtocolError) throw error;
        if (error instanceof PageImportBatchFailedError) throw error;
        if (
          requiresRecoveredByteVerification &&
          isRecoveredBatchIdentityConflict(error)
        ) {
          // The browser hint collided with a different byte payload/scope. The
          // old operation is terminal or unavailable, so a fresh id is safe;
          // never mutate or downgrade the old receipt.
          replaceRecoveredBatchId();
          putAttempt = 0;
          continue;
        }
        if (!isAmbiguousMutationFailure(error)) throw error;
        lastError = error;

        const probe = await probeImportBatch(path, {
          expectedBatchId: effectiveBatchId,
          manifest: input.manifest,
          delayMs,
          requestId: requestIdFromApiError(error),
          signal: input.signal,
        });
        if (probe.kind === "receipt") {
          if (!requiresRecoveredByteVerification) {
            return committed(probe.receipt);
          }
          // The receipt may belong to a metadata-colliding file. Only a direct
          // successful PUT response proves this invocation's bytes matched it.
          putAttempt += 1;
          continue;
        }
        if (input.signal?.aborted) throw error;
        if (probe.kind === "orphan") {
          // A GET status probe proved that no live owner holds this exact
          // batch. Replay the identical multipart immediately under the same
          // id; do not replace its pending recovery hint or idempotency key.
          putAttempt += 1;
          continue;
        }
        putAttempt += 1;
        if (probe.kind !== "not_found" || putAttempt >= 2) break;
      }
    }
  } catch (error) {
    if (error instanceof PageImportBatchFailedError) {
      clearPendingBatch(pendingIdentity, effectiveBatchId);
    }
    throw error;
  }
  throw lastError instanceof Error
    ? lastError
    : new Error("The import batch result could not be confirmed.");
}

export function createPageImportBatchId(): string {
  if (
    typeof crypto === "undefined" ||
    typeof crypto.randomUUID !== "function"
  ) {
    throw new PageImportBatchProtocolError(
      "Secure UUID generation is unavailable here.",
    );
  }
  return crypto.randomUUID();
}
