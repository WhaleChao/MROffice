/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfCanonicalId } from "../../feature-capabilities";
import type { PdfSourceRiskProfile } from "../../save-policy";
import { MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION, MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA } from "./vocabulary";
import { MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION, MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION, MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION, MRAI_NATIVE_LIST_TEXT_ACTION, MRAI_NATIVE_OUTLINE_ACTION } from "./vocabulary";
import { MRAI_NATIVE_FONT_REUSE_TEXT_ACTION, MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION, MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION, MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION, MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION, MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION } from "./vocabulary";
import {
  validateManagedOutlineTree,
  type ManagedOutlineNode,
} from "@mrnote/core/web/lib/pdf/outline.ts";
import { readPdfCommandOwnershipManifest } from "../command-ownership";
import { type PdfNativeTextDecorationGeometry, type PdfNativeTextDecorationHistory, type PdfNativeTextDecorationStyle } from "../../native-text-decoration";
import { MRAI_NATIVE_TEXT_DECORATION_ACTION, MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION, PDF_NATIVE_TEXT_DECORATION_CONTROL_ID } from "./vocabulary";
import { type PdfNativeTextObjectDeleteHistory } from "../../native-text-object-delete";
import { MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION, MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION, PDF_NATIVE_TEXT_OBJECT_DELETE_CONTROL_ID } from "./vocabulary";
import { type PdfNativeEmbeddedImageReorderHistory, type PdfNativeEmbeddedImageReorderIntent } from "../../native-embedded-image-reorder";
import { MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION, MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION } from "./vocabulary";
import { MRAI_NATIVE_CONTENT_APPEND_ACTION, MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION } from "./vocabulary";
import { MRAI_NATIVE_BYTES_PUBLISH_ACTION, MRAI_NATIVE_BYTES_PUBLISH_KIND, MRAI_NATIVE_BYTES_PUBLISH_SCHEMA_VERSION } from "./vocabulary";
import { MRAI_NATIVE_PAGE_BACKGROUND_ACTION, MRAI_NATIVE_PAGE_BACKGROUND_KIND, MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION, MRAI_NATIVE_PAGE_SPLIT_ACTION, MRAI_NATIVE_PAGE_SPLIT_KIND, MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION } from "./vocabulary";
import { isPdfNativeSourceConsolidationRecord } from "./source-consolidation";
import {
  readPdfEngineCommand,
  type PdfEngineCommand,
  type PdfEngineSourceIdentity,
  type PdfEngineWritableSnapshot,
} from "../contracts";
import { type PdfNativeListTextSelectionEvidence } from "./list-text-fidelity";
import { MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION } from "./vocabulary";
import { MRAI_NATIVE_TEXT_MOVE_KIND, MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION } from "./vocabulary";
import { type PdfNativeTextGroupEdit } from "./text-group-move";
import { MRAI_NATIVE_TEXT_GROUP_MOVE_KIND, MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION } from "./vocabulary";
import { mraiNativeTextGroupTransformIsIdentity } from "./text-group-move-model";
import { type PdfNativeEmbeddedImageRemovalEvidence, type PdfNativeEmbeddedImageTransformEvidence, type PdfNativeEmbeddedImageTransformIntent, type PdfNativeEmbeddedImageReplacementEvidence } from "./content-stream-transaction";
import { MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION, MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION, MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION, MRAI_NATIVE_OUTLINE_SCHEMA_VERSION } from "./vocabulary";
import { MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES } from "./transaction-error";

export const PDF_NATIVE_LIST_TEXT_FEATURE_ID =
  "PDF-EDT-TEXT-01" as const satisfies PdfCanonicalId;
export const PDF_NATIVE_LIST_TEXT_EDIT_CONTROL_ID =
  "PDFCTL.RIBBON.EDIT.EDIT_CONTENT" as const;
export const PDF_NATIVE_LIST_TEXT_HOME_CONTROL_ID =
  "PDFCTL.RIBBON.HOME.EDIT_CONTENT" as const;
export const PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID =
  "PDF-EDT-IMAGE-02" as const satisfies PdfCanonicalId;
export const PDF_NATIVE_EMBEDDED_IMAGE_DELETE_CONTROL_ID =
  "PDFCTL.CONTEXT.PICTURE.DELETE_SOURCE" as const;
export const PDF_NATIVE_EMBEDDED_IMAGE_REPLACE_CONTROL_ID =
  "PDFCTL.CONTEXT.PICTURE.REPLACE_SOURCE" as const;
export const PDF_NATIVE_EMBEDDED_IMAGE_TRANSFORM_CONTROL_ID =
  "PDFCTL.CONTEXT.PICTURE.TRANSFORM_SOURCE" as const;
export const PDF_NATIVE_EMBEDDED_IMAGE_REORDER_CONTROL_ID =
  "PDFCTL.CONTEXT.PICTURE.ARRANGE_SOURCE" as const;
export const PDF_NATIVE_OUTLINE_FEATURE_ID =
  "PDF-HOM-NAV-05" as const satisfies PdfCanonicalId;
export const PDF_NATIVE_OUTLINE_CONTROL_ID =
  "PDFCTL.OUTLINE.MANAGED_LIFECYCLE" as const;

export type PdfNativeListTextControlId =
  | typeof PDF_NATIVE_LIST_TEXT_EDIT_CONTROL_ID
  | typeof PDF_NATIVE_LIST_TEXT_HOME_CONTROL_ID;

export type PdfNativeListTextLifecycleCode =
  | "read-only"
  | "source-not-ready"
  | "encrypted"
  /** A certification (DocMDP) signature; an approval signature is not one. */
  | "signed"
  | "native-replay-pending"
  /** The last change the server refused to write into the PDF (the page's
   * patchback receipt failed): the journal cannot be current until it is
   * repaired, and a reload does not change that. */
  | "native-writeback-failed"
  /** A legacy journal (records without source checkpoints, which no source
   * consolidation compacts) at the parser's ceiling: every native write
   * appends one xref section and the native parser reads at most 32. A
   * checkpointed chain is consolidated long before it. */
  | "native-chain-exhausted"
  /** The selected text is on a page an active import or replacement holds as
   * its own (patchbackReceipts.ts pdfNativeImportedPageNumbers); every other
   * page stays editable. */
  | "imported-page"
  | "selection-required"
  | "page-drift"
  | "selection-drift"
  | "source-mismatch"
  | "invalid-replacement";

export class PdfNativeListTextLifecycleError extends Error {
  readonly code: PdfNativeListTextLifecycleCode;

  constructor(code: PdfNativeListTextLifecycleCode, message: string) {
    super(message);
    this.name = "PdfNativeListTextLifecycleError";
    this.code = code;
  }
}

export interface PdfNativeListTextEditLease {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly controlId: PdfNativeListTextControlId;
  readonly source: PdfEngineSourceIdentity;
  readonly selection: PdfNativeListTextSelectionEvidence;
}

export interface PdfNativeListTextAvailabilityInput {
  readonly writable: boolean;
  readonly sourceReady: boolean;
  readonly riskProfile: PdfSourceRiskProfile | null;
  readonly hasExistingNativePatchback: boolean;
  readonly nativeTransactionCount: number | null;
  /**
   * True only when the complete stable-id journal is an append-only SHA chain
   * whose last derived candidate is the exact lossless receipt currently
   * rendered by the host. Historical rows are then acknowledged history, not
   * work that may be replayed against the latest source a second time.
   */
  readonly nativeReplayCurrent?: boolean;
  readonly selection: PdfNativeListTextSelectionEvidence | null;
  readonly currentPage: number;
  readonly expectedSourceSha256: string | null;
}

export type PdfNativeListTextPatchbackMetadataKind =
  | "generic"
  | "native-journal"
  | "unknown";

export interface PdfNativeListTextPatchbackAuthority {
  readonly resourceId: string;
  readonly resourceVersionId: string;
  readonly outputSha256: string;
}

export type PdfNativeListTextReplayAuthority = Readonly<
  | {
      state: "empty";
      transactionCount: 0;
      patchback: null;
    }
  | {
      state: "current";
      transactionCount: number;
      patchback: PdfNativeListTextPatchbackAuthority;
      latestDerivedSha256: string;
    }
  | {
      state: "exhausted";
      transactionCount: 31;
      patchback: PdfNativeListTextPatchbackAuthority;
      latestDerivedSha256: string;
    }
  | {
      state: "pending";
      transactionCount: number | null;
      patchback: PdfNativeListTextPatchbackAuthority | null;
    }
>;

const SHA256_PATTERN = /^[0-9a-f]{64}$/;
// The native parser accepts at most 32 xref sections. Starting from the
// narrowest one-section source, 31 append-only edits are the absolute ceiling;
// once 31 receipts exist the host must not open a 32nd edit. Sources that
// already contain more sections are still re-authorized independently by the
// planner/replay and can refuse earlier without persisting anything.
const MAX_ACKNOWLEDGED_NATIVE_TEXT_TRANSACTIONS = 31;

function ownDataRecord(
  value: unknown,
): Readonly<Record<string, unknown>> | null {
  try {
    if (!value || typeof value !== "object" || Array.isArray(value))
      return null;
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) return null;
    const descriptors = Object.getOwnPropertyDescriptors(value);
    const result: Record<string, unknown> = {};
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== "string") return null;
      const descriptor = descriptors[key];
      if (!descriptor?.enumerable || !("value" in descriptor)) return null;
      result[key] = descriptor.value;
    }
    return result;
  } catch {
    return null;
  }
}

function nonEmptyBoundedString(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length > 0 &&
    value.length <= 256 &&
    value === value.trim()
  );
}

function normalizedSha256(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.toLowerCase();
  return SHA256_PATTERN.test(normalized) ? normalized : null;
}

/**
 * Classify the exact ResourceVersion behind a durable PDF receipt without
 * treating every original-package writer as native-list-text history.
 *
 * Presence of either native-journal metadata key is deliberately enough to
 * classify the version as native text. Missing or malformed journal payloads
 * must therefore remain fail-closed in the replay layer, while a verified
 * generic version (for example an AcroForm-only write) can become the next
 * source baseline for a new text journal.
 */
export function readPdfNativeListTextPatchbackMetadataKind(input: {
  readonly resourceId: string;
  readonly resourceVersionId: string;
  readonly outputSha256: string;
  readonly versions: readonly unknown[];
}): PdfNativeListTextPatchbackMetadataKind {
  const expectedOutputSha256 = normalizedSha256(input.outputSha256);
  if (
    !nonEmptyBoundedString(input.resourceId) ||
    !nonEmptyBoundedString(input.resourceVersionId) ||
    !expectedOutputSha256 ||
    !Array.isArray(input.versions)
  ) {
    return "unknown";
  }
  const matches = input.versions.filter((value) => {
    const version = ownDataRecord(value);
    const metadata = ownDataRecord(version?.metadata_json);
    return Boolean(
      version &&
        metadata &&
        version.id === input.resourceVersionId &&
        version.resource_id === input.resourceId &&
        version.version_kind === "export" &&
        normalizedSha256(metadata.output_sha256) === expectedOutputSha256,
    );
  });
  if (matches.length !== 1) return "unknown";
  const metadata = ownDataRecord(ownDataRecord(matches[0])?.metadata_json);
  if (!metadata) return "unknown";
  return Object.hasOwn(metadata, "pdf_native_journal") ||
    Object.hasOwn(metadata, "pdf_native_transaction_id")
    ? "native-journal"
    : "generic";
}

/**
 * Authenticate the browser-visible cumulative journal without trusting a
 * derived count. Every historical row must form one source-SHA -> candidate-
 * SHA chain. A non-empty journal is usable only after the latest durable PDF
 * receipt and the bytes rendered by pdf.js both acknowledge its last output.
 */
export function readPdfNativeListTextReplayAuthority(input: {
  readonly collection: unknown;
  readonly patchback: PdfNativeListTextPatchbackAuthority | null;
  readonly renderedSourceSha256: string | null;
}): PdfNativeListTextReplayAuthority {
  const { collection, patchback } = input;
  if (
    collection === undefined ||
    (Array.isArray(collection) && collection.length === 0)
  ) {
    return patchback
      ? Object.freeze({ state: "pending", transactionCount: 0, patchback })
      : Object.freeze({ state: "empty", transactionCount: 0, patchback: null });
  }
  if (
    !Array.isArray(collection) ||
    collection.length < 1 ||
    collection.length > MAX_ACKNOWLEDGED_NATIVE_TEXT_TRANSACTIONS
  ) {
    return Object.freeze({
      state: "pending",
      transactionCount: null,
      patchback,
    });
  }

  const ids = new Set<string>();
  let resourceId: string | null = null;
  let previousDerivedSha256: string | null = null;
  let previousSourceVersionId: string | null = null;
  for (const value of collection) {
    const record = ownDataRecord(value);
    const source = ownDataRecord(record?.source);
    const derived = ownDataRecord(record?.derivedCandidate);
    const id = record?.id;
    const transactionId = record?.transactionId;
    const nextResourceId = source?.resourceId;
    const nextResourceVersionId = source?.resourceVersionId;
    const sourceSha256 = normalizedSha256(source?.sha256);
    const derivedSha256 = normalizedSha256(derived?.sha256);
    if (
      !record ||
      !source ||
      !derived ||
      !nonEmptyBoundedString(id) ||
      id !== transactionId ||
      ids.has(id) ||
      !(
        (record.action === MRAI_NATIVE_LIST_TEXT_ACTION &&
          (record.kind === "text-replacement" ||
            record.kind === "list-text-replacement" ||
            (record.kind === MRAI_NATIVE_TEXT_MOVE_KIND && record.schemaVersion === MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION) ||
            (record.kind === MRAI_NATIVE_TEXT_GROUP_MOVE_KIND &&
              record.schemaVersion === MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION) ||
            (record.kind === "text-window-history" && record.schemaVersion === "mrai.pdf-native-text.window-history.v1"))) ||
        (record.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION &&
          record.kind === "font-extension-text-replacement") ||
        (record.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION &&
          record.kind === "font-extension-text-history") ||
        (record.action === MRAI_NATIVE_FONT_REUSE_TEXT_ACTION &&
          record.kind === "font-reuse-text-replacement" &&
          record.schemaVersion === MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION) ||
        (record.action === MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION &&
          record.kind === "paragraph-properties" &&
          record.schemaVersion === MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA) ||
        (record.action === MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION &&
          record.kind === "embedded-image-removal") ||
        (record.action === MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION &&
          record.kind === "embedded-image-replacement") ||
        (record.action === MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION &&
          record.kind === "embedded-image-transform" &&
          record.schemaVersion === MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION) ||
        (record.action === MRAI_NATIVE_TEXT_DECORATION_ACTION &&
          record.kind === "text-decoration" &&
          record.schemaVersion === MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION) ||
        (record.action === MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION &&
          record.kind === "text-object-delete" &&
          record.schemaVersion === MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION) ||
        (record.action === MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION &&
          record.kind === "embedded-image-reorder" &&
          record.schemaVersion === MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION) ||
        (record.action === MRAI_NATIVE_OUTLINE_ACTION &&
          record.kind === "outline-mutation") ||
        (record.action === MRAI_NATIVE_CONTENT_APPEND_ACTION &&
          record.kind === "content-append" &&
          record.schemaVersion === MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION) ||
        (record.action === MRAI_NATIVE_PAGE_SPLIT_ACTION &&
          record.kind === MRAI_NATIVE_PAGE_SPLIT_KIND &&
          record.schemaVersion === MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION) ||
        (record.action === MRAI_NATIVE_PAGE_BACKGROUND_ACTION &&
          record.kind === MRAI_NATIVE_PAGE_BACKGROUND_KIND &&
          record.schemaVersion === MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION) ||
        // A revision the engine appended, published as its bytes.
        (record.action === MRAI_NATIVE_BYTES_PUBLISH_ACTION &&
          record.kind === MRAI_NATIVE_BYTES_PUBLISH_KIND &&
          record.schemaVersion === MRAI_NATIVE_BYTES_PUBLISH_SCHEMA_VERSION) ||
        // The server's consolidation of the chain (source-consolidation.ts):
        // the one record a rewritten chain keeps, chaining the verified
        // candidate to the one-section file that became the version.
        isPdfNativeSourceConsolidationRecord(record)
      ) ||
      !nonEmptyBoundedString(nextResourceId) ||
      !nonEmptyBoundedString(nextResourceVersionId) ||
      !sourceSha256 ||
      !derivedSha256 ||
      sourceSha256 === derivedSha256 ||
      (resourceId !== null && resourceId !== nextResourceId) ||
      (previousDerivedSha256 !== null &&
        sourceSha256 !== previousDerivedSha256) ||
      (previousSourceVersionId !== null &&
        previousSourceVersionId === nextResourceVersionId)
    ) {
      return Object.freeze({
        state: "pending",
        transactionCount: collection.length,
        patchback,
      });
    }
    ids.add(id);
    resourceId = nextResourceId;
    previousSourceVersionId = nextResourceVersionId;
    previousDerivedSha256 = derivedSha256;
  }

  const renderedSourceSha256 = normalizedSha256(input.renderedSourceSha256);
  if (
    resourceId === null ||
    previousSourceVersionId === null ||
    previousDerivedSha256 === null ||
    !patchback ||
    patchback.resourceId !== resourceId ||
    !nonEmptyBoundedString(patchback.resourceVersionId) ||
    patchback.resourceVersionId === previousSourceVersionId ||
    normalizedSha256(patchback.outputSha256) !== previousDerivedSha256 ||
    renderedSourceSha256 !== previousDerivedSha256
  ) {
    return Object.freeze({
      state: "pending",
      transactionCount: collection.length,
      patchback,
    });
  }
  return collection.length === MAX_ACKNOWLEDGED_NATIVE_TEXT_TRANSACTIONS
    ? Object.freeze({
        state: "exhausted",
        transactionCount: MAX_ACKNOWLEDGED_NATIVE_TEXT_TRANSACTIONS,
        patchback,
        latestDerivedSha256: previousDerivedSha256,
      })
    : Object.freeze({
        state: "current",
        transactionCount: collection.length,
        patchback,
        latestDerivedSha256: previousDerivedSha256,
      });
}

function recognizedControlId(
  value: string,
): value is PdfNativeListTextControlId {
  return (
    value === PDF_NATIVE_LIST_TEXT_EDIT_CONTROL_ID ||
    value === PDF_NATIVE_LIST_TEXT_HOME_CONTROL_ID
  );
}

function exactSource(
  left: PdfEngineSourceIdentity,
  right: PdfEngineSourceIdentity,
): boolean {
  return (
    left.resourceId === right.resourceId &&
    left.resourceVersionId === right.resourceVersionId &&
    left.sourceSha256 === right.sourceSha256
  );
}

export function samePdfNativeListTextSelection(
  left: PdfNativeListTextSelectionEvidence | null,
  right: PdfNativeListTextSelectionEvidence | null,
): boolean {
  if (!left || !right) return left === right;
  const sameItem = (
    leftItem: PdfNativeListTextSelectionEvidence["item"],
    rightItem: PdfNativeListTextSelectionEvidence["item"],
  ) =>
    leftItem.itemIndex === rightItem.itemIndex &&
    leftItem.str === rightItem.str &&
    leftItem.fontName === rightItem.fontName &&
    leftItem.transform.every(
      (value, index) => value === rightItem.transform[index],
    ) &&
    leftItem.width === rightItem.width &&
    leftItem.height === rightItem.height &&
    leftItem.hasEOL === rightItem.hasEOL;
  return (
    left.sourceSha256 === right.sourceSha256 &&
    left.selectionRevision === right.selectionRevision &&
    left.pageNumber === right.pageNumber &&
    left.startOffset === right.startOffset &&
    left.endOffset === right.endOffset &&
    sameItem(left.item, right.item) &&
    (left.items?.length ?? 0) === (right.items?.length ?? 0) &&
    (left.items?.every(
      (item, index) =>
        right.items !== undefined && sameItem(item, right.items[index]),
    ) ??
      true)
  );
}

export function readPdfNativeListTextAvailabilityCode(
  input: PdfNativeListTextAvailabilityInput,
): PdfNativeListTextLifecycleCode | null {
  if (!input.writable) return "read-only";
  if (
    input.nativeTransactionCount === null ||
    (input.nativeTransactionCount === 0
      ? input.hasExistingNativePatchback
      : !input.hasExistingNativePatchback || input.nativeReplayCurrent !== true)
  ) {
    return "native-replay-pending";
  }
  if (!input.sourceReady || !input.riskProfile) return "source-not-ready";
  // Only encryption and a certification (DocMDP) signature keep native text
  
  // signs the revision it was made over and the write appends a new one; an
  // XFA form's page content is written as any page's; an inspection that
  // could not read some other part of the catalog is background. Encryption
  // or a certification the inspection could not settle still waits.
  if (input.riskProfile.encryption === "present") return "encrypted";
  if (input.riskProfile.certifiedSignatures === "present") return "signed";
  if (
    input.riskProfile.encryption !== "none" ||
    input.riskProfile.certifiedSignatures !== "none"
  ) {
    return "source-not-ready";
  }
  if (!input.selection) return "selection-required";
  if (input.selection.pageNumber !== input.currentPage) return "page-drift";
  if (
    !input.expectedSourceSha256 ||
    input.selection.sourceSha256 !== input.expectedSourceSha256
  ) {
    return "source-mismatch";
  }
  return null;
}

function freezeSelection(
  selection: PdfNativeListTextSelectionEvidence,
): PdfNativeListTextSelectionEvidence {
  return Object.freeze({
    ...selection,
    item: Object.freeze({
      ...selection.item,
      transform: Object.freeze([
        ...selection.item.transform,
      ]) as PdfNativeListTextSelectionEvidence["item"]["transform"],
    }),
    ...(selection.items
      ? {
          items: Object.freeze(
            selection.items.map((item) =>
              Object.freeze({
                ...item,
                transform: Object.freeze([
                  ...item.transform,
                ]) as PdfNativeListTextSelectionEvidence["item"]["transform"],
              }),
            ),
          ),
        }
      : {}),
  });
}

export function createPdfNativeListTextEditLease(input: {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly controlId: string;
  readonly source: PdfEngineSourceIdentity;
  readonly selection: PdfNativeListTextSelectionEvidence;
}): PdfNativeListTextEditLease {
  if (
    !input.pageId.trim() ||
    !Number.isSafeInteger(input.pageContentVersion) ||
    input.pageContentVersion < 0 ||
    !recognizedControlId(input.controlId) ||
    input.selection.sourceSha256 !== input.source.sourceSha256
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The selected text is not bound to the current PDF source.",
    );
  }
  return Object.freeze({
    pageId: input.pageId,
    pageContentVersion: input.pageContentVersion,
    controlId: input.controlId,
    source: Object.freeze({ ...input.source }),
    selection: freezeSelection(input.selection),
  });
}

export function createPdfNativeListTextOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
      action: MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
      action: MRAI_NATIVE_LIST_TEXT_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
      action: MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
      action: MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export interface PdfNativeEmbeddedImageEditLease {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly source: PdfEngineSourceIdentity;
  readonly pageNumber: number;
  readonly imageOrder: number;
}

export function createPdfNativeEmbeddedImageOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
      action: MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
      action: MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
      action: MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
      action: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export interface PdfNativeTextObjectDeleteLease {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly source: PdfEngineSourceIdentity;
  readonly pageNumber: number;
  readonly runSpan: Readonly<{ start: number; end: number }>;
  /** The page content stream the span above is an offset into. */
  readonly contentStream?: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  /** The reader selection the run was proven from. */
  readonly textSelection?: PdfNativeListTextSelectionEvidence;
  /** The rules read beside the run, taken away in the same write. */
  readonly keepTextRules?: PdfNativeTextObjectDeleteKeepTextRules;
}

export type PdfNativeTextObjectDeleteKeepTextRules = Readonly<{
  styles: readonly ("underline" | "strikethrough")[];
  geometry: Readonly<{
    textMatrix: readonly [number, number, number, number, number, number];
    fontSize: number;
    advance: number;
    color: Readonly<{ kind: "gray" | "rgb" | "cmyk"; components: readonly number[] }>;
  }>;
}>;

export function createPdfNativeTextObjectDeleteOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
      action: MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function buildPdfNativeTextObjectDeleteCommand(input: {
  readonly lease: PdfNativeTextObjectDeleteLease;
  readonly currentPageId: string;
  readonly currentPageContentVersion: number;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly transactionId: string;
  readonly direction: "delete" | "restore";
  readonly history?: PdfNativeTextObjectDeleteHistory;
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    input.currentPageId !== input.lease.pageId ||
    input.currentPageContentVersion !== input.lease.pageContentVersion ||
    !exactSource(input.currentSource, input.lease.source)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource changed after the text deletion was staged.",
    );
  }
  if ((input.direction === "restore") !== (input.history !== undefined)) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native text deletion needs exactly one authority.",
    );
  }
  const command = readPdfEngineCommand({
    ...input.lease.source,
    featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
    controlId: PDF_NATIVE_TEXT_OBJECT_DELETE_CONTROL_ID,
    action: MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
      selection: {
        pageNumber: input.lease.pageNumber,
        runSpan: { ...input.lease.runSpan },
        ...(input.lease.contentStream
          ? { contentStream: { ...input.lease.contentStream } }
          : {}),
        ...(input.lease.textSelection
          ? { textSelection: JSON.parse(JSON.stringify(input.lease.textSelection)) }
          : {}),
        ...(input.lease.keepTextRules
          ? { keepTextRules: JSON.parse(JSON.stringify(input.lease.keepTextRules)) }
          : {}),
      },
      direction: input.direction,
      ...(input.history
        ? {
            // Everything the restore needs: what the delete emptied and the
            // pen number it may have left, not just where and what.
            history: {
              offset: input.history.offset,
              objectBytes: [...input.history.objectBytes],
              ...(input.history.scope ? { scope: input.history.scope } : {}),
              ...(input.history.penRestore
                ? { penRestore: input.history.penRestore }
                : {}),
              ...(input.history.penRestores
                ? { penRestores: [...input.history.penRestores] }
                : {}),
              ...(input.history.penCharacterSpacings
                ? { penCharacterSpacings: [...input.history.penCharacterSpacings] }
                : {}),
            },
          }
        : {}),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native text deletion failed canonical validation.",
    );
  }
  return command;
}

export function createPdfNativeTextObjectDeleteLease(input: {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly source: PdfEngineSourceIdentity;
  readonly pageNumber: number;
  readonly runSpan: Readonly<{ start: number; end: number }>;
  readonly contentStream?: Readonly<{
    objectNumber: number;
    generation: number;
  }> | null;
  readonly textSelection?: PdfNativeListTextSelectionEvidence | null;
  readonly keepTextRules?: PdfNativeTextObjectDeleteKeepTextRules | null;
}): PdfNativeTextObjectDeleteLease {
  return Object.freeze({
    pageId: input.pageId,
    pageContentVersion: input.pageContentVersion,
    source: Object.freeze({ ...input.source }),
    pageNumber: input.pageNumber,
    runSpan: Object.freeze({ ...input.runSpan }),
    ...(input.contentStream
      ? { contentStream: Object.freeze({ ...input.contentStream }) }
      : {}),
    ...(input.textSelection ? { textSelection: input.textSelection } : {}),
    ...(input.keepTextRules ? { keepTextRules: input.keepTextRules } : {}),
  });
}

export function createPdfNativeTextDecorationOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
      action: MRAI_NATIVE_TEXT_DECORATION_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function buildPdfNativeTextDecorationCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly transactionId: string;
  readonly pageNumber: number;
  readonly style: PdfNativeTextDecorationStyle;
  readonly direction: "add" | "remove";
  /** The page content stream the run's span and the rule's offset belong to. */
  readonly contentStream?: Readonly<{
    objectNumber: number;
    generation: number;
  }> | null;
  readonly geometry?: PdfNativeTextDecorationGeometry;
  readonly history?: PdfNativeTextDecorationHistory;
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (!exactSource(input.currentSource, input.source)) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource changed after the text rule was staged.",
    );
  }
  if ((input.direction === "remove") !== (input.history !== undefined)) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native text rule needs exactly one authority.",
    );
  }
  const command = readPdfEngineCommand({
    ...input.source,
    featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
    controlId: PDF_NATIVE_TEXT_DECORATION_CONTROL_ID,
    action: MRAI_NATIVE_TEXT_DECORATION_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
      pageNumber: input.pageNumber,
      style: input.style,
      direction: input.direction,
      ...(input.contentStream
        ? { contentStream: { ...input.contentStream } }
        : {}),
      ...(input.geometry
        ? {
            geometry: {
              textMatrix: [...input.geometry.textMatrix],
              fontSize: input.geometry.fontSize,
              advance: input.geometry.advance,
              color: {
                kind: input.geometry.color.kind,
                components: [...input.geometry.color.components],
              },
              runSpan: { ...input.geometry.runSpan },
            },
          }
        : {}),
      ...(input.history
        ? {
            history: {
              offset: input.history.offset,
              ruleBytes: [...input.history.ruleBytes],
            },
          }
        : {}),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native text rule failed canonical validation.",
    );
  }
  return command;
}

export function buildPdfNativeEmbeddedImageReorderCommand(input: {
  readonly lease: PdfNativeEmbeddedImageEditLease;
  readonly currentPageId: string;
  readonly currentPageContentVersion: number;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly transactionId: string;
  readonly direction: PdfNativeEmbeddedImageReorderIntent | "restore";
  readonly history?: PdfNativeEmbeddedImageReorderHistory;
  /** A front move out from under the picture's clips (the payload's). */
  readonly leaveClips?: boolean;
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    input.currentPageId !== input.lease.pageId ||
    input.currentPageContentVersion !== input.lease.pageContentVersion ||
    !exactSource(input.currentSource, input.lease.source)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource changed after the picture arrange was staged.",
    );
  }
  if ((input.direction === "restore") !== (input.history !== undefined)) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native picture arrange command needs exactly one authority.",
    );
  }
  const command = readPdfEngineCommand({
    ...input.lease.source,
    featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
    controlId: PDF_NATIVE_EMBEDDED_IMAGE_REORDER_CONTROL_ID,
    action: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
      selection: {
        pageNumber: input.lease.pageNumber,
        imageOrder: input.lease.imageOrder,
      },
      direction: input.direction,
      ...(input.leaveClips ? { leaveClips: true } : {}),
      ...(input.history ? { history: { ...input.history } } : {}),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native picture arrange command failed canonical validation.",
    );
  }
  return command;
}

export interface PdfNativeOutlineEditLease {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly source: PdfEngineSourceIdentity;
  readonly expectedManagedIds: readonly string[];
}

function managedOutlineIds(
  tree: readonly ManagedOutlineNode[],
): readonly string[] {
  const ids: string[] = [];
  const walk = (nodes: readonly ManagedOutlineNode[]) =>
    nodes.forEach((node) => {
      ids.push(node.id);
      walk(node.children);
    });
  walk(tree);
  return Object.freeze(ids);
}

export function createPdfNativeOutlineOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_OUTLINE_FEATURE_ID,
      action: MRAI_NATIVE_OUTLINE_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function createPdfNativeOutlineEditLease(input: {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly source: PdfEngineSourceIdentity;
  readonly currentTree: readonly ManagedOutlineNode[];
}): PdfNativeOutlineEditLease {
  if (
    !input.pageId.trim() ||
    !Number.isSafeInteger(input.pageContentVersion) ||
    input.pageContentVersion < 0
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The managed outline is not bound to the current PDF source.",
    );
  }
  const tree = validateManagedOutlineTree(input.currentTree);
  return Object.freeze({
    pageId: input.pageId,
    pageContentVersion: input.pageContentVersion,
    source: Object.freeze({ ...input.source }),
    expectedManagedIds: managedOutlineIds(tree),
  });
}

export function buildPdfNativeOutlineMutationCommand(input: {
  readonly lease: PdfNativeOutlineEditLease;
  readonly currentPageId: string;
  readonly currentPageContentVersion: number;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly transactionId: string;
  readonly operation:
    | "add"
    | "rename"
    | "update-destination"
    | "move"
    | "delete"
    | "rename-source"
    | "delete-source";
  readonly managedTree: readonly ManagedOutlineNode[];
  /** rename-source / delete-source: one of the file's own bookmarks, by its
   * place in the outline (sibling indexes), and a rename's new title. */
  readonly sourceTarget?: Readonly<{ path: readonly number[]; title: string | null; expectedTitle?: string }>;
  /** The window shows a saved file holding the page's structure: the tree
   * names the upload's pages, this is the file's map (for each file page,
   * the upload's page it is; null for an added one). */
  readonly pageMap?: readonly (number | null)[];
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    input.currentPageId !== input.lease.pageId ||
    input.currentPageContentVersion !== input.lease.pageContentVersion ||
    !exactSource(input.currentSource, input.lease.source)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource changed after the outline edit was staged.",
    );
  }
  const managedTree = validateManagedOutlineTree(input.managedTree);
  const command = readPdfEngineCommand({
    ...input.lease.source,
    featureId: PDF_NATIVE_OUTLINE_FEATURE_ID,
    controlId: PDF_NATIVE_OUTLINE_CONTROL_ID,
    action: MRAI_NATIVE_OUTLINE_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_OUTLINE_SCHEMA_VERSION,
      operation: input.operation,
      expectedManagedIds: [...input.lease.expectedManagedIds],
      managedTree,
      ...(input.sourceTarget
        ? { sourceTarget: {
            path: [...input.sourceTarget.path],
            title: input.sourceTarget.title,
            ...(input.sourceTarget.expectedTitle !== undefined ? { expectedTitle: input.sourceTarget.expectedTitle } : {}),
          } }
        : {}),
      ...(input.pageMap ? { pageMap: [...input.pageMap] } : {}),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native outline command failed canonical validation.",
    );
  }
  return command;
}

/** Read the latest authenticated journal projection for panel hydration. */
export function hasPdfNativeManagedOutlineMutation(collection: unknown): boolean {
  if (!Array.isArray(collection)) return false;
  return collection.some((value) => {
    const record = ownDataRecord(value);
    return (
      record?.action === MRAI_NATIVE_OUTLINE_ACTION ||
      record?.kind === "outline-mutation"
    );
  });
}

export function readPdfNativeManagedOutlineProjection(
  collection: unknown,
): readonly ManagedOutlineNode[] {
  if (!Array.isArray(collection) || collection.length === 0) return Object.freeze([]);
  for (let index = collection.length - 1; index >= 0; index -= 1) {
    const record = ownDataRecord(collection[index]);
    if (
      record?.action !== MRAI_NATIVE_OUTLINE_ACTION ||
      record.kind !== "outline-mutation"
    ) {
      continue;
    }
    const mutation = ownDataRecord(record.outlineMutation);
    if (
      mutation?.schemaVersion !== MRAI_NATIVE_OUTLINE_SCHEMA_VERSION ||
      !Array.isArray(mutation.managedTree)
    ) {
      throw new PdfNativeListTextLifecycleError(
        "native-replay-pending",
        "The managed outline journal is malformed.",
      );
    }
    // Written on a file holding the page's structure: its tree is on that
    // file's pages, read back as the upload's (the numbers the window keeps).
    const pageMap = mutation.pageMap;
    if (
      pageMap !== undefined &&
      (!Array.isArray(pageMap) ||
        !pageMap.every((page) => page === null || (Number.isSafeInteger(page) && (page as number) >= 1)))
    ) {
      throw new PdfNativeListTextLifecycleError(
        "native-replay-pending",
        "The managed outline journal page map is malformed.",
      );
    }
    const shownPage = (filePage: unknown): unknown => {
      if (!Array.isArray(pageMap)) return filePage;
      const upload = Number.isSafeInteger(filePage) ? pageMap[(filePage as number) - 1] : undefined;
      if (typeof upload !== "number") {
        throw new PdfNativeListTextLifecycleError(
          "native-replay-pending",
          "The managed outline journal names a page outside its map.",
        );
      }
      return upload;
    };
    const project = (nodes: unknown[]): ManagedOutlineNode[] =>
      nodes.map((value) => {
        const node = ownDataRecord(value);
        const destination = ownDataRecord(node?.destination);
        if (!node || !destination || !Array.isArray(node.children)) {
          throw new PdfNativeListTextLifecycleError(
            "native-replay-pending",
            "The managed outline journal projection is malformed.",
          );
        }
        return {
          id: node.id as string,
          title: node.title as string,
          destination: {
            pageNumber: shownPage(destination.pageNumber) as number,
            left: destination.left as number | null,
            top: destination.top as number | null,
            zoom: destination.zoom as number | null,
          },
          children: project(node.children),
        };
      });
    return validateManagedOutlineTree(project(mutation.managedTree));
  }
  return Object.freeze([]);
}

/** The bookmarks a file holding the page's structure keeps (the upload's
 * page numbers): one on a page deleted from the file (`deletedPages`, the
 * upload's numbers) is gone and its children take its place, as the save's
 * remap writes them (remap_managed_outline_after_page_structure). */
export function managedOutlineWithoutDeletedPages(
  tree: readonly ManagedOutlineNode[],
  deletedPages: readonly number[],
): readonly ManagedOutlineNode[] {
  if (deletedPages.length === 0) return tree;
  const deleted = new Set(deletedPages);
  let dropped = false;
  const within = (nodes: readonly ManagedOutlineNode[]): ManagedOutlineNode[] =>
    nodes.flatMap((node) => {
      const children = within(node.children);
      if (deleted.has(node.destination.pageNumber)) {
        dropped = true;
        return children;
      }
      return [{ ...node, children }];
    });
  const result = within(tree);
  return dropped ? validateManagedOutlineTree(result) : tree;
}

export function createPdfNativeEmbeddedImageEditLease(input: {
  readonly pageId: string;
  readonly pageContentVersion: number;
  readonly source: PdfEngineSourceIdentity;
  readonly pageNumber: number;
  readonly imageOrder: number;
}): PdfNativeEmbeddedImageEditLease {
  if (
    !input.pageId.trim() ||
    !Number.isSafeInteger(input.pageContentVersion) ||
    input.pageContentVersion < 0 ||
    !Number.isSafeInteger(input.pageNumber) ||
    input.pageNumber < 1 ||
    !Number.isSafeInteger(input.imageOrder) ||
    input.imageOrder < 0
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The selected image is not bound to the current PDF source.",
    );
  }
  return Object.freeze({
    ...input,
    source: Object.freeze({ ...input.source }),
  });
}

export function buildPdfNativeEmbeddedImageRemoveCommand(input: {
  readonly lease: PdfNativeEmbeddedImageEditLease;
  readonly currentPageId: string;
  readonly currentPageContentVersion: number;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly transactionId: string;
  readonly historyEvidence?: PdfNativeEmbeddedImageRemovalEvidence;
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    input.currentPageId !== input.lease.pageId ||
    input.currentPageContentVersion !== input.lease.pageContentVersion ||
    !exactSource(input.currentSource, input.lease.source)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource changed after the image edit was staged.",
    );
  }
  const command = readPdfEngineCommand({
    ...input.lease.source,
    featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
    controlId: PDF_NATIVE_EMBEDDED_IMAGE_DELETE_CONTROL_ID,
    action: MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION,
      selection: {
        pageNumber: input.lease.pageNumber,
        imageOrder: input.lease.imageOrder,
      },
      ...(input.historyEvidence
        ? { historySchemaVersion: input.historyEvidence.schemaVersion }
        : {}),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native image command failed canonical validation.",
    );
  }
  return command;
}

export function buildPdfNativeEmbeddedImageTransformCommand(input: {
  readonly lease: PdfNativeEmbeddedImageEditLease;
  readonly currentPageId: string;
  readonly currentPageContentVersion: number;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly transactionId: string;
  readonly transform?: PdfNativeEmbeddedImageTransformIntent;
  readonly historyEvidence?: PdfNativeEmbeddedImageTransformEvidence;
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    input.currentPageId !== input.lease.pageId ||
    input.currentPageContentVersion !== input.lease.pageContentVersion ||
    !exactSource(input.currentSource, input.lease.source)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource changed after the image edit was staged.",
    );
  }
  if (Boolean(input.transform) === Boolean(input.historyEvidence)) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native image transform command needs exactly one authority.",
    );
  }
  const command = readPdfEngineCommand({
    ...input.lease.source,
    featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
    controlId: PDF_NATIVE_EMBEDDED_IMAGE_TRANSFORM_CONTROL_ID,
    action: MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION,
      selection: {
        pageNumber: input.lease.pageNumber,
        imageOrder: input.lease.imageOrder,
      },
      ...(input.transform
        ? {
            intent: {
              ctm: [...input.transform.ctm],
              frame: { ...input.transform.frame },
              quarterTurns: input.transform.quarterTurns,
              flipX: input.transform.flipX,
              flipY: input.transform.flipY,
              crop: input.transform.crop ? { ...input.transform.crop } : null,
              ...(input.transform.opacity === undefined ||
              input.transform.opacity === null
                ? {}
                : { opacity: input.transform.opacity }),
              ...(input.transform.pageTransform
                ? { pageTransform: [...input.transform.pageTransform] }
                : {}),
              ...(input.transform.cropMode ? { cropMode: input.transform.cropMode } : {}),
            },
          }
        : { historySchemaVersion: input.historyEvidence!.schemaVersion }),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native image transform command failed canonical validation.",
    );
  }
  return command;
}

function nativeImageBytesBase64(bytes: Uint8Array): string {
  if (
    bytes.byteLength < 1 ||
    bytes.byteLength > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The replacement image exceeds the native limit.",
    );
  }
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

export function buildPdfNativeEmbeddedImageReplaceCommand(input: {
  readonly lease: PdfNativeEmbeddedImageEditLease;
  readonly currentPageId: string;
  readonly currentPageContentVersion: number;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
  readonly transactionId: string;
  readonly replacement?: Readonly<{
    bytes: Uint8Array;
    mimeType: "image/jpeg" | "image/png";
  }>;
  readonly historyEvidence?: PdfNativeEmbeddedImageReplacementEvidence;
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    input.currentPageId !== input.lease.pageId ||
    input.currentPageContentVersion !== input.lease.pageContentVersion ||
    !exactSource(input.currentSource, input.lease.source) ||
    Boolean(input.replacement) === Boolean(input.historyEvidence)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource changed after the image replacement was staged.",
    );
  }
  const command = readPdfEngineCommand({
    ...input.lease.source,
    featureId: PDF_NATIVE_EMBEDDED_IMAGE_FEATURE_ID,
    controlId: PDF_NATIVE_EMBEDDED_IMAGE_REPLACE_CONTROL_ID,
    action: MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION,
      selection: {
        pageNumber: input.lease.pageNumber,
        imageOrder: input.lease.imageOrder,
      },
      ...(input.replacement
        ? {
            replacementMimeType: input.replacement.mimeType,
            replacementDataBase64: nativeImageBytesBase64(
              input.replacement.bytes,
            ),
          }
        : {
            historySchemaVersion: input.historyEvidence!.schemaVersion,
          }),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native image replacement command failed canonical validation.",
    );
  }
  return command;
}

export function buildPdfNativeListTextEditCommand(input: {
  readonly lease: PdfNativeListTextEditLease;
  readonly currentPageId: string;
  readonly currentPageContentVersion: number;
  readonly currentSource: PdfEngineSourceIdentity;
  readonly currentSelection: PdfNativeListTextSelectionEvidence | null;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly transactionId: string;
  readonly replacementText: string;
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    input.currentPageId !== input.lease.pageId ||
    input.currentPageContentVersion !== input.lease.pageContentVersion
  ) {
    throw new PdfNativeListTextLifecycleError(
      "page-drift",
      "The PDF page changed after the edit dialog opened.",
    );
  }
  if (!exactSource(input.currentSource, input.lease.source)) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The PDF resource version changed after the edit dialog opened.",
    );
  }
  if (
    !samePdfNativeListTextSelection(
      input.currentSelection,
      input.lease.selection,
    )
  ) {
    throw new PdfNativeListTextLifecycleError(
      "selection-drift",
      "The PDF text selection changed after the edit dialog opened.",
    );
  }
  if (
    typeof input.replacementText !== "string" ||
    input.replacementText.length === 0 ||
    input.replacementText.length > 2_000 ||
    // A typed line break ("\n", Enter) starts a new line of the paragraph
    // (the paragraph rewrite sets it); a carriage return is no text.
    /\r/.test(input.replacementText)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "invalid-replacement",
      "Enter one non-empty line of replacement text.",
    );
  }

  const command = readPdfEngineCommand({
    ...input.lease.source,
    featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
    controlId: input.lease.controlId,
    action: MRAI_NATIVE_LIST_TEXT_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.lease.selection.selectionRevision,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION,
      selection: input.lease.selection,
      replacementText: input.replacementText,
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native text command failed canonical validation.",
    );
  }
  return command;
}

/**
 * The text edit command carrying a native move of the selected text instead
 * of new words (text-move.ts): the same lease, drift and writability checks,
 * with the translation in PDF user space as its payload.
 */
export function buildPdfNativeTextMoveCommand(
  input: Omit<Parameters<typeof buildPdfNativeListTextEditCommand>[0], "replacementText"> & {
    readonly move: Readonly<{ x: number; y: number }>;
  },
): PdfEngineCommand {
  // A move keeps the words, so the edit's one-line text rule has nothing to
  // check; the payload below carries no text at all.
  const canonical = buildPdfNativeListTextEditCommand({
    ...input,
    replacementText: "move",
  });
  if (
    !Number.isFinite(input.move.x) ||
    !Number.isFinite(input.move.y) ||
    (input.move.x === 0 && input.move.y === 0)
  ) {
    throw new PdfNativeListTextLifecycleError(
      "invalid-replacement",
      "The text is already at this position.",
    );
  }
  const command = readPdfEngineCommand({
    ...canonical,
    payload: {
      schemaVersion: MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION,
      selection: input.lease.selection,
      move: { x: input.move.x, y: input.move.y },
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native text move command failed canonical validation.",
    );
  }
  return command;
}

/**
 * The text edit command carrying a native group move or resize of a whole
 * box (text-group-move.ts), or the exact restore of one (its undo): the same
 * lease, drift and writability checks, with the map of the page -- or the
 * move it puts back -- as its payload.
 */
export function buildPdfNativeTextGroupMoveCommand(
  input: Omit<Parameters<typeof buildPdfNativeListTextEditCommand>[0], "replacementText"> & (
    | { readonly transform: PdfNativeTextGroupEdit; readonly restore?: undefined }
    | { readonly restore: Readonly<{ originTransactionId: string }>; readonly transform?: undefined }
  ),
): PdfEngineCommand {
  const canonical = buildPdfNativeListTextEditCommand({
    ...input,
    replacementText: "move",
  });
  if (input.transform && mraiNativeTextGroupTransformIsIdentity(input.transform)) {
    throw new PdfNativeListTextLifecycleError(
      "invalid-replacement",
      "The text is already at this position and size.",
    );
  }
  const command = readPdfEngineCommand({
    ...canonical,
    payload: {
      schemaVersion: MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION,
      selection: input.lease.selection,
      ...(input.transform
        ? {
            transform: { a: input.transform.a, d: input.transform.d, e: input.transform.e, f: input.transform.f },
            // A weight set with the map (a bold alone keeps it the identity).
            ...(input.transform.bold ? { bold: JSON.parse(JSON.stringify(input.transform.bold)) } : {}),
          }
        : { restore: { originTransactionId: input.restore!.originTransactionId } }),
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native group move command failed canonical validation.",
    );
  }
  return command;
}

export function buildPdfNativeFontExtensionTextEditCommand(
  input: Parameters<typeof buildPdfNativeListTextEditCommand>[0],
): PdfEngineCommand {
  const canonical = buildPdfNativeListTextEditCommand(input);
  const command = readPdfEngineCommand({
    ...canonical,
    action: MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
    payload: {
      schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
      selection: input.lease.selection,
      replacementText: input.replacementText,
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native font extension command failed canonical validation.",
    );
  }
  return command;
}

export function buildPdfNativeFontExtensionTextHistoryCommand(input: {
  readonly currentSource: PdfEngineSourceIdentity;
  readonly writable: PdfEngineWritableSnapshot;
  readonly sessionRevision: number;
  readonly transactionId: string;
  readonly originTransactionId: string;
  readonly direction: "undo" | "redo";
}): PdfEngineCommand {
  if (!input.writable.writable) {
    throw new PdfNativeListTextLifecycleError(
      "read-only",
      input.writable.reason ?? "The PDF page is read-only.",
    );
  }
  if (
    !input.originTransactionId ||
    !input.transactionId ||
    (input.direction !== "undo" && input.direction !== "redo")
  ) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native font extension history authority is invalid.",
    );
  }
  const command = readPdfEngineCommand({
    ...input.currentSource,
    featureId: PDF_NATIVE_LIST_TEXT_FEATURE_ID,
    controlId: PDF_NATIVE_LIST_TEXT_EDIT_CONTROL_ID,
    action: MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: 0,
    output: "current-pdf",
    payload: {
      schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION,
      originTransactionId: input.originTransactionId,
      direction: input.direction,
    },
  });
  if (!command) {
    throw new PdfNativeListTextLifecycleError(
      "source-mismatch",
      "The native font extension history command failed canonical validation.",
    );
  }
  return command;
}
