/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import {
  PdfEngineNotWritableError,
  type PdfEngineWritableSnapshot,
} from "./contracts";

export type PdfSharedStateDocument = Record<string, JsonValue>;
export type PdfSharedStateRecord = Record<string, JsonValue> & {
  readonly id: string;
};

export interface PdfSharedStatePersistence {
  readonly seam: "queueSave" | "saveNow";
  persist(nextContentJson: PdfSharedStateDocument): unknown | Promise<unknown>;
}

export interface PdfSharedStateBridgeOptions {
  readonly readContentJson: () => PdfSharedStateDocument;
  readonly readWritable: () => PdfEngineWritableSnapshot;
  readonly persistence: PdfSharedStatePersistence;
  readonly projectRemote: (
    contentJson: PdfSharedStateDocument,
    transactionId: string,
  ) => void | Promise<void>;
}

export interface PdfSharedStateMutation {
  readonly collection: string;
  readonly record: PdfSharedStateRecord;
  readonly transactionId: string;
  readonly origin?: "local" | "remote";
}

export type PdfSharedStateCommitResult =
  | Readonly<{
      status: "persisted";
      collection: string;
      id: string;
      seam: PdfSharedStatePersistence["seam"];
      contentJson: PdfSharedStateDocument;
      receipt: unknown;
    }>
  | Readonly<{
      status: "echo-suppressed";
      transactionId: string;
    }>;

export interface PdfSharedStateBridge {
  commitLocalRecord(
    mutation: PdfSharedStateMutation,
  ): Promise<PdfSharedStateCommitResult>;
  applyRemoteProjection(
    contentJson: PdfSharedStateDocument,
    transactionId: string,
  ): Promise<void>;
}

function isRecord(value: unknown): value is Record<string, JsonValue> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function cloneJson(value: JsonValue): JsonValue {
  if (Array.isArray(value)) return value.map(cloneJson);
  if (isRecord(value)) {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [key, cloneJson(entry)]),
    );
  }
  return value;
}

function jsonValueEquals(left: JsonValue, right: JsonValue): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) && Array.isArray(right)) {
    return (
      left.length === right.length &&
      left.every((entry, index) => jsonValueEquals(entry, right[index]))
    );
  }
  if (isRecord(left) && isRecord(right)) {
    const leftKeys = Object.keys(left);
    const rightKeys = Object.keys(right);
    return (
      leftKeys.length === rightKeys.length &&
      leftKeys.every(
        (key) => key in right && jsonValueEquals(left[key], right[key]),
      )
    );
  }
  return false;
}







function storedJsonValueEquals(left: JsonValue, right: JsonValue): boolean {
  if (typeof left === "string" && typeof right === "string") {
    return left.replaceAll("\u0000", "\ufffd") === right.replaceAll("\u0000", "\ufffd");
  }
  if (Array.isArray(left) && Array.isArray(right)) {
    return (
      left.length === right.length &&
      left.every((entry, index) => storedJsonValueEquals(entry, right[index]))
    );
  }
  if (isRecord(left) && isRecord(right)) {
    const leftKeys = Object.keys(left);
    return (
      leftKeys.length === Object.keys(right).length &&
      leftKeys.every(
        (key) => key in right && storedJsonValueEquals(left[key], right[key]),
      )
    );
  }
  return Object.is(left, right);
}

function isStableId(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length > 0 &&
    value.length <= 256 &&
    value === value.trim()
  );
}

function containsOpaqueXfdf(value: JsonValue): boolean {
  if (typeof value === "string") return /<\s*xfdf\b/i.test(value);
  if (Array.isArray(value)) return value.some(containsOpaqueXfdf);
  if (!isRecord(value)) return false;
  return Object.entries(value).some(
    ([key, entry]) => /xfdf/i.test(key) || containsOpaqueXfdf(entry),
  );
}

function assertKeyedCollection(
  collection: string,
  value: JsonValue,
): asserts value is PdfSharedStateRecord[] {
  if (!Array.isArray(value)) {
    throw new Error(
      `PDF shared-state collection ${collection} is not an array.`,
    );
  }
  const ids = new Set<string>();
  for (const item of value) {
    if (!isRecord(item) || !isStableId(item.id)) {
      throw new Error(
        `PDF shared-state collection ${collection} contains a record without a stable id.`,
      );
    }
    if (ids.has(item.id)) {
      throw new Error(
        `PDF shared-state collection ${collection} contains duplicate id ${item.id}.`,
      );
    }
    ids.add(item.id);
  }
}

function assertSharedDocument(contentJson: PdfSharedStateDocument): void {
  if (containsOpaqueXfdf(contentJson)) {
    throw new Error("Opaque XFDF cannot be stored as PDF shared state.");
  }
  for (const [key, value] of Object.entries(contentJson)) {
    if (
      Array.isArray(value) &&
      value.some((item) => typeof item === "object" && item !== null)
    ) {
      assertKeyedCollection(key, value);
    }
  }
}

function readCollectionName(value: string): string | null {
  const name = value.trim();
  return /^[A-Za-z][A-Za-z0-9_]*$/.test(name) && !/xfdf/i.test(name)
    ? name
    : null;
}

function upsertSharedRecord(
  contentJson: PdfSharedStateDocument,
  collection: string,
  record: PdfSharedStateRecord,
): PdfSharedStateDocument {
  const existing = contentJson[collection] ?? [];
  assertKeyedCollection(collection, existing);
  const index = existing.findIndex((item) => item.id === record.id);
  const nextCollection = [...existing];
  if (index >= 0) nextCollection[index] = record;
  else nextCollection.push(record);
  return { ...contentJson, [collection]: nextCollection };
}

function findSharedRecord(
  contentJson: PdfSharedStateDocument,
  collection: string,
  id: string,
): PdfSharedStateRecord | null {
  const value = contentJson[collection];
  if (value === undefined) return null;
  assertKeyedCollection(collection, value);
  return value.find((record) => record.id === id) ?? null;
}

function pendingRecordKey(collection: string, id: string): string {
  return `${collection}\u0000${id}`;
}

/**
 * Whether an observed collection acknowledges a local record it no longer
 * holds: the server consolidated the chain at that record, keeping one
 * consolidation record that names it as its origin (source-consolidation.ts).
 * Replaying the local record over that journal would append it a second
 * time, which the next save refuses as a non-append-only history.
 */
function sharedRecordConsolidated(
  contentJson: PdfSharedStateDocument,
  collection: string,
  id: string,
): boolean {
  const value = contentJson[collection];
  if (!Array.isArray(value) || value.length === 0) return false;
  if (value.some((record) => isRecord(record) && record.id === id)) return false;
  // The record a consolidation started from, and every record it folded and
  // kept the restore evidence of (restorableHistory).
  return value.some((record) => {
    if (!isRecord(record) || record.kind !== "source-consolidation") return false;
    const consolidation = record as {
      sourceConsolidation?: { originTransactionId?: unknown };
      restorableHistory?: unknown;
    };
    return consolidation.sourceConsolidation?.originTransactionId === id ||
      (Array.isArray(consolidation.restorableHistory) &&
        consolidation.restorableHistory.some((entry) => isRecord(entry) && entry.id === id));
  });
}

export function createPdfSharedStateBridge(
  options: PdfSharedStateBridgeOptions,
): PdfSharedStateBridge {
  const remoteTransactions = new Set<string>();
  const unacknowledgedLocalRecords = new Map<
    string,
    { collection: string; record: PdfSharedStateRecord }
  >();

  return {
    commitLocalRecord(mutation) {
      if (
        mutation.origin === "remote" ||
        remoteTransactions.has(mutation.transactionId)
      ) {
        return Promise.resolve(
          Object.freeze({
            status: "echo-suppressed",
            transactionId: mutation.transactionId,
          }),
        );
      }

      const collection = readCollectionName(mutation.collection);
      if (!collection) {
        return Promise.reject(
          new Error("Invalid PDF shared-state collection name."),
        );
      }
      if (!isStableId(mutation.record.id)) {
        return Promise.reject(
          new Error("PDF shared-state records require a stable id."),
        );
      }
      if (containsOpaqueXfdf(mutation.record)) {
        return Promise.reject(
          new Error("Opaque XFDF cannot be stored as PDF shared state."),
        );
      }

      let current = cloneJson(
        options.readContentJson(),
      ) as PdfSharedStateDocument;
      assertSharedDocument(current);

      // queueSave returns void before React/content_json necessarily exposes
      // the draft. A local record is acknowledged only after a later observed
      // snapshot contains the same id AND value. Until then, replay the latest
      // value for each collection+id over the newest observed base.
      for (const [key, pending] of unacknowledgedLocalRecords) {
        const observed = findSharedRecord(
          current,
          pending.collection,
          pending.record.id,
        );
        if (observed && storedJsonValueEquals(observed, pending.record)) {
          unacknowledgedLocalRecords.delete(key);
          continue;
        }
        if (sharedRecordConsolidated(current, pending.collection, pending.record.id)) {
          unacknowledgedLocalRecords.delete(key);
          continue;
        }
        current = upsertSharedRecord(
          current,
          pending.collection,
          pending.record,
        );
      }

      const record = cloneJson(mutation.record) as PdfSharedStateRecord;
      const nextContentJson = upsertSharedRecord(current, collection, record);
      assertSharedDocument(nextContentJson);

      const writable = options.readWritable();
      if (!writable.writable) {
        throw new PdfEngineNotWritableError(
          writable.reason ?? "The PDF page became read-only before mutation.",
        );
      }

      // queueSave/saveNow must be entered synchronously in the committing
      // gesture. The returned promise is observed only after that single call.
      const key = pendingRecordKey(collection, record.id);
      unacknowledgedLocalRecords.set(key, {
        collection,
        record,
      });
      const rollbackRejectedRecord = () => {
        const pending = unacknowledgedLocalRecords.get(key);
        if (pending && jsonValueEquals(pending.record, record)) {
          unacknowledgedLocalRecords.delete(key);
        }
      };
      let pendingPersistence: unknown | Promise<unknown>;
      try {
        pendingPersistence = options.persistence.persist(nextContentJson);
      } catch (error) {
        rollbackRejectedRecord();
        throw error;
      }
      return Promise.resolve(pendingPersistence).then(
        (receipt) =>
          Object.freeze({
            status: "persisted" as const,
            collection,
            id: record.id,
            seam: options.persistence.seam,
            contentJson: nextContentJson,
            receipt,
          }),
        (error) => {
          rollbackRejectedRecord();
          throw error;
        },
      );
    },

    async applyRemoteProjection(contentJson, transactionId) {
      if (!transactionId.trim()) {
        throw new Error("Remote PDF projection requires a transaction id.");
      }
      const cloned = cloneJson(contentJson) as PdfSharedStateDocument;
      assertSharedDocument(cloned);
      remoteTransactions.add(transactionId);
      try {
        await options.projectRemote(cloned, transactionId);
      } finally {
        remoteTransactions.delete(transactionId);
      }
    },
  };
}
