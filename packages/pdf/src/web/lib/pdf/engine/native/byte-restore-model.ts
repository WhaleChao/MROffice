/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Json } from "./byte-restore-history";

/** Writers whose history is not a byte restore of one content-stream span. */
export const NOT_RESTORABLE = new Set([
  "mrai.pdf-native-outline.mutate.v1",
  "mrai.pdf-native-embedded-image.replace.v1",
  "mrai.pdf-native-text.font-extension-history.v1",
  "mrai.pdf-native-content-append.v1",
  "mrai.pdf-native-page-split.v1",
  "mrai.pdf-native-page-background.v1",
]);

export function record(value: unknown): Json | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Json) : null;
}

export function int(value: unknown): number | null {
  return Number.isSafeInteger(value) && (value as number) >= 0 ? (value as number) : null;
}

/** Whether a journal record's edit can be undone by writing its bytes back. */
export function mraiNativeRecordIsByteRestorable(originRecord: unknown): boolean {
  const origin = record(originRecord);
  const derived = record(origin?.derivedCandidate);
  // A picture removal that also unmapped the picture's resource name is not
  // undone by its content bytes alone: the restored `Do` would name nothing.
  const unmapped = record(record(origin?.embeddedImageRemoval)?.resourceUnmap);
  return Boolean(origin && typeof origin.id === "string" && origin.id && !NOT_RESTORABLE.has(String(origin.schemaVersion)) && !unmapped &&
    record(origin.target) && record(origin.operatorSpan) && typeof origin.replacementSha256 === "string" &&
    (Array.isArray(origin.replacement) || Array.isArray(origin.replacementPieces)) &&
    int(derived?.sourceLength) !== null && int(derived?.byteLength) !== null && typeof derived?.sha256 === "string");
}

/**
 * A forward edit's byte restore once its revisions are gone (mirror of the
 * server's NativeRestoreEvidence, native_surgical_contract.py). A source
 * consolidation rewrites the chain as one section; before it does, the
 * server reads from each restorable record's revisions what the restore
 * reads from them -- the object the forward wrote, the bytes it replaced and
 * wrote (base64 of a zlib stream), and per direction the window those bytes
 * held with the digests of the stream before its end and after its start --
 * and the consolidation record keeps the list (`restorableHistory`).
 */
export const MRAI_NATIVE_RESTORE_EVIDENCE_SCHEMA_VERSION = "mrai.pdf-native-restore-evidence.v1";

// The evidence is journal JSON: object type aliases (not interfaces) keep it
// assignable where a record's JSON is expected.

export type PdfNativeRestoreWindow = {
  readonly length: number;
  readonly start: number;
  readonly end: number;
  readonly prefixSha256: string;
  readonly suffixSha256: string;
};

export type PdfNativeRestoreEvidence = {
  readonly schemaVersion: typeof MRAI_NATIVE_RESTORE_EVIDENCE_SCHEMA_VERSION;
  /** The forward record's transaction id (the history's origin). */
  readonly id: string;
  readonly kind: string;
  readonly originSchemaVersion: string;
  readonly resourceId: string;
  readonly originSourceSha256: string;
  readonly originCandidateSha256: string;
  readonly target: { readonly objectNumber: number; readonly generation: number };
  readonly replaced: string;
  readonly written: string;
  readonly undo: PdfNativeRestoreWindow;
  readonly redo: PdfNativeRestoreWindow;
};

const SHA256 = /^[a-f0-9]{64}$/;

function restoreWindow(value: unknown): value is PdfNativeRestoreWindow {
  const window = record(value);
  const length = int(window?.length), start = int(window?.start), end = int(window?.end);
  return Boolean(window && length !== null && start !== null && end !== null && start < end && end <= length &&
    typeof window.prefixSha256 === "string" && SHA256.test(window.prefixSha256) &&
    typeof window.suffixSha256 === "string" && SHA256.test(window.suffixSha256));
}

/** Whether a value has the shape of restore evidence (the server validated
 * its bytes when it wrote the consolidation record). */
export function isPdfNativeRestoreEvidence(value: unknown): value is PdfNativeRestoreEvidence {
  const evidence = record(value);
  const target = record(evidence?.target);
  return Boolean(evidence && evidence.schemaVersion === MRAI_NATIVE_RESTORE_EVIDENCE_SCHEMA_VERSION &&
    typeof evidence.id === "string" && evidence.id && typeof evidence.kind === "string" &&
    typeof evidence.originSchemaVersion === "string" && !NOT_RESTORABLE.has(evidence.originSchemaVersion) &&
    typeof evidence.resourceId === "string" && evidence.resourceId &&
    typeof evidence.originSourceSha256 === "string" && SHA256.test(evidence.originSourceSha256) &&
    typeof evidence.originCandidateSha256 === "string" && SHA256.test(evidence.originCandidateSha256) &&
    target && int(target.objectNumber) !== null && int(target.generation) !== null &&
    typeof evidence.replaced === "string" && typeof evidence.written === "string" &&
    restoreWindow(evidence.undo) && restoreWindow(evidence.redo));
}

/** Whether a history origin -- a journal record, or the evidence a
 * consolidation kept for one -- can be restored by writing bytes back. */
export function mraiNativeRestoreOriginIsUsable(origin: unknown): boolean {
  return mraiNativeRecordIsByteRestorable(origin) || isPdfNativeRestoreEvidence(origin);
}

/**
 * A font-extension or font-reuse edit's history once its revisions are gone
 * (mirror of the server's NativeObjectRestoreEvidence). That history is not a
 * byte window: it swaps whole objects back -- the edited content stream, and
 * for an extension the page whose font resources it rewired -- and checks
 * that the font objects it installed still stand. The consolidation keeps
 * each of those objects as a history revision writes it ("N G obj ...
 * endobj", packed like the window bytes), the digests of the streams the two
 * targets carry, and the installed objects' digests.
 */
export const MRAI_NATIVE_OBJECT_RESTORE_EVIDENCE_SCHEMA_VERSION = "mrai.pdf-native-object-restore-evidence.v1";

export type PdfNativeObjectRestoreEvidence = {
  readonly schemaVersion: typeof MRAI_NATIVE_OBJECT_RESTORE_EVIDENCE_SCHEMA_VERSION;
  /** The forward record's transaction id (the history's origin). */
  readonly id: string;
  readonly kind: string;
  readonly originSchemaVersion: string;
  /** The forward's transaction schema, which its history record repeats. */
  readonly transactionSchemaVersion: string;
  readonly resourceId: string;
  readonly originSourceSha256: string;
  readonly originCandidateSha256: string;
  readonly originKind: "extension" | "reuse";
  readonly target: { readonly objectNumber: number; readonly generation: number };
  readonly page: { readonly objectNumber: number; readonly generation: number };
  readonly originalTarget: string;
  readonly extendedTarget: string;
  readonly originalPage: string;
  readonly extendedPage: string;
  readonly originalTargetStreamSha256: string;
  readonly extendedTargetStreamSha256: string;
  readonly protected: readonly { readonly objectNumber: number; readonly generation: number; readonly sha256: string }[];
};

function objectRef(value: unknown): boolean {
  const ref = record(value);
  return Boolean(ref && int(ref.objectNumber) !== null && (ref.objectNumber as number) >= 1 && int(ref.generation) !== null);
}

/** Whether a value has the shape of a font history's object restore
 * evidence (the server validated its bytes when it wrote the record). */
export function isPdfNativeObjectRestoreEvidence(value: unknown): value is PdfNativeObjectRestoreEvidence {
  const evidence = record(value);
  const text = (key: string) => typeof evidence?.[key] === "string" && Boolean(evidence[key]);
  const digest = (key: string) => typeof evidence?.[key] === "string" && SHA256.test(evidence[key] as string);
  return Boolean(evidence && evidence.schemaVersion === MRAI_NATIVE_OBJECT_RESTORE_EVIDENCE_SCHEMA_VERSION &&
    text("id") && text("kind") && text("originSchemaVersion") && text("transactionSchemaVersion") && text("resourceId") &&
    digest("originSourceSha256") && digest("originCandidateSha256") &&
    (evidence.originKind === "extension" || evidence.originKind === "reuse") &&
    objectRef(evidence.target) && objectRef(evidence.page) &&
    text("originalTarget") && text("extendedTarget") && text("originalPage") && text("extendedPage") &&
    digest("originalTargetStreamSha256") && digest("extendedTargetStreamSha256") &&
    Array.isArray(evidence.protected) &&
    evidence.protected.every((item) => objectRef(item) && typeof record(item)?.sha256 === "string" &&
      SHA256.test(record(item)!.sha256 as string)));
}

/**
 * A revision the engine appended and published as its bytes, once its two
 * sides are gone (mirror of the server's NativeRevisionRestoreEvidence). Its
 * Undo and Redo are an object restore (planMraiNativeObjectRestore), which
 * reads the revision's sides as prefixes of the current file; a source
 * consolidation drops those prefixes and keeps instead, per bytes-publish
 * record, every object the revision wrote over one that already stood (its
 * body after the revision and before it, packed like the window bytes, with
 * the generation each side wrote it under) and every object it created (by
 * digest only: the restore leaves those where they are). A body is what
 * stands between `N G obj` and `endobj` without the white space round it.
 * planMraiNativeObjectRestoreFromEvidence plans the same restore from it.
 */
export const MRAI_NATIVE_REVISION_RESTORE_EVIDENCE_SCHEMA_VERSION = "mrai.pdf-native-revision-restore-evidence.v1";

export type PdfNativeRevisionRestoreObject = {
  readonly objectNumber: number;
  /** The generation the revision wrote it under (a redo writes it). */
  readonly generation: number;
  /** The generation it stood under before the revision (an undo writes it). */
  readonly beforeGeneration: number;
  /** Its body as the revision wrote it, and as it stood before. */
  readonly after: string;
  readonly before: string;
};

export type PdfNativeRevisionCreatedObject = {
  readonly objectNumber: number;
  readonly generation: number;
  /** SHA-256 of the body the revision wrote. */
  readonly sha256: string;
};

export type PdfNativeRevisionRestoreEvidence = {
  readonly schemaVersion: typeof MRAI_NATIVE_REVISION_RESTORE_EVIDENCE_SCHEMA_VERSION;
  /** The bytes-publish record's id (the history's origin). */
  readonly id: string;
  readonly kind: string;
  readonly originSchemaVersion: string;
  readonly resourceId: string;
  readonly originSourceSha256: string;
  readonly originCandidateSha256: string;
  /** Ascending by object number; an object appears in one list only. */
  readonly objects: readonly PdfNativeRevisionRestoreObject[];
  readonly created: readonly PdfNativeRevisionCreatedObject[];
};

const MAX_REVISION_RESTORE_OBJECTS = 512;

function generationOf(value: unknown): boolean {
  return int(value) !== null && (value as number) <= 65_535;
}

/** Whether a value has the shape of a published revision's restore evidence
 * (the server validated its bytes when it wrote the consolidation record). */
export function isPdfNativeRevisionRestoreEvidence(value: unknown): value is PdfNativeRevisionRestoreEvidence {
  const evidence = record(value);
  const text = (key: string) => typeof evidence?.[key] === "string" && Boolean(evidence[key]);
  const digest = (key: string) => typeof evidence?.[key] === "string" && SHA256.test(evidence[key] as string);
  if (!evidence || evidence.schemaVersion !== MRAI_NATIVE_REVISION_RESTORE_EVIDENCE_SCHEMA_VERSION ||
    !text("id") || !text("kind") || !text("originSchemaVersion") || !text("resourceId") ||
    !digest("originSourceSha256") || !digest("originCandidateSha256") ||
    evidence.originSourceSha256 === evidence.originCandidateSha256 ||
    !Array.isArray(evidence.objects) || !Array.isArray(evidence.created) || evidence.objects.length < 1 ||
    evidence.objects.length + evidence.created.length > MAX_REVISION_RESTORE_OBJECTS) {
    return false;
  }
  const objects = evidence.objects as unknown[], createdObjects = evidence.created as unknown[];
  const restored = objects.every((item) => {
    const object = record(item);
    return Boolean(object && objectRef(object) && generationOf(object.generation) &&
      generationOf(object.beforeGeneration) && typeof object.after === "string" && object.after.length >= 4 &&
      typeof object.before === "string" && object.before.length >= 4);
  });
  const created = createdObjects.every((item) => {
    const object = record(item);
    return Boolean(object && objectRef(object) && generationOf(object.generation) &&
      typeof object.sha256 === "string" && SHA256.test(object.sha256));
  });
  if (!restored || !created) return false;
  const numberOf = (item: unknown) => record(item)!.objectNumber as number;
  const ascending = (items: readonly unknown[]) =>
    items.every((item, index) => index === 0 || numberOf(item) > numberOf(items[index - 1]));
  const restoredNumbers = new Set(objects.map(numberOf));
  return ascending(objects) && ascending(createdObjects) &&
    createdObjects.every((item) => !restoredNumbers.has(numberOf(item)));
}
