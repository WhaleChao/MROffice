/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";

import { appendManagedNativeNoteTombstone, type NativePdfPageAnnotation } from "@mrnote/core/web/lib/pdf/native-note-roundtrip.ts";

import { mergePageDocumentDelta } from "@mrnote/core/web/hooks/usePageContentSync.ts";
import {
  applyJsonDelta,
  createJsonCollaborationHistory,
  jsonEquals,
  yToJson,
  type JsonValue,
} from "@mrnote/core/web/lib/collab/json-crdt.ts";

interface NoteIdentity {
  id: string;
  managedMarkup?: unknown;
  managedProofing?: unknown;
  nativeSource?: { kind: string };
}

export interface PdfOrdinaryNoteSlice<T extends NoteIdentity> {
  annotations: T[];
  tombstones: T[];
}

export interface PdfOrdinaryNoteIntent<T extends NoteIdentity> {
  readonly before: PdfOrdinaryNoteSlice<T>;
  readonly after: PdfOrdinaryNoteSlice<T>;
}

function byId<T extends NoteIdentity>(notes: readonly T[]): Record<string, JsonValue> {
  const ids = new Set<string>();
  const entries = notes.map((note) => {
    if (!note.id || ids.has(note.id)) throw new Error("Invalid PDF note identity");
    ids.add(note.id);
    return [note.id, note as unknown as JsonValue];
  });
  return Object.fromEntries(entries);
}

function asDocument<T extends NoteIdentity>(slice: PdfOrdinaryNoteSlice<T>): Record<string, JsonValue> {
  byId(slice.annotations);
  byId(slice.tombstones);
  return slice as unknown as Record<string, JsonValue>;
}

/** Capture before the source/lease awaits. Managed markup/proofing keeps its
 * existing exact native writer path; this adapter owns ordinary Text notes. */
export function capturePdfOrdinaryNoteIntent<T extends NoteIdentity>(
  before: PdfOrdinaryNoteSlice<T>,
  after: PdfOrdinaryNoteSlice<T>,
): PdfOrdinaryNoteIntent<T> | null {
  for (const key of ["annotations", "tombstones"] as const) {
    const old = byId(before[key]);
    const next = byId(after[key]);
    for (const id of new Set([...Object.keys(old), ...Object.keys(next)])) {
      if (jsonEquals(old[id], next[id])) continue;
      for (const note of [old[id], next[id]]) {
        const value = note as unknown as NoteIdentity | undefined;
        if (value && (value.managedMarkup || value.managedProofing ||
          (value.nativeSource && value.nativeSource.kind !== "pdf_text_annotation"))) return null;
      }
    }
  }
  return structuredClone({ before, after });
}

/** Reuse the shared per-id/property delta; unknown peer notes/replies stay. */
export function applyPdfOrdinaryNoteIntents<T extends NoteIdentity>(
  current: PdfOrdinaryNoteSlice<T>,
  intents: readonly PdfOrdinaryNoteIntent<T>[],
): PdfOrdinaryNoteSlice<T> {
  return intents.reduce((value, intent) => mergePageDocumentDelta(
    asDocument(value), asDocument(intent.after), asDocument(intent.before),
  ) as unknown as PdfOrdinaryNoteSlice<T>, structuredClone(current));
}

/** A merged ACK may contain other authors' notes and a different concurrent
 * list order. Require this invocation's changed IDs/properties, not old arrays. */
export function pdfOrdinaryNoteIntentIsPresent<T extends NoteIdentity>(
  current: PdfOrdinaryNoteSlice<T>,
  intent: PdfOrdinaryNoteIntent<T>,
): boolean {
  for (const key of ["annotations", "tombstones"] as const) {
    const value = byId(current[key]);
    if (!jsonEquals(value, mergePageDocumentDelta(value, byId(intent.after[key]), byId(intent.before[key])))) return false;
  }
  return true;
}

/** Roll back only this optimistic local operation. The shared history guard
 * preserves peer edits, replies and peer-used creations; no old list restore. */
export function revertPdfOrdinaryNoteIntent<T extends NoteIdentity>(
  current: PdfOrdinaryNoteSlice<T>,
  intent: PdfOrdinaryNoteIntent<T>,
): PdfOrdinaryNoteSlice<T> {
  const doc = new Y.Doc();
  const root = doc.getMap<unknown>("pdf-note-intent");
  const localOrigin = Symbol("pdf-optimistic-note");
  doc.transact(() => applyJsonDelta(root, asDocument(intent.before), {}));
  const history = createJsonCollaborationHistory(root, { localOrigin, canWrite: () => true });
  try {
    doc.transact(() => applyJsonDelta(root, asDocument(intent.after), asDocument(intent.before)), localOrigin);
    history.stopCapturing();
    doc.transact(() => applyJsonDelta(root, asDocument(current), asDocument(intent.after)), "peer-after-note");
    history.undo();
    return yToJson(root) as unknown as PdfOrdinaryNoteSlice<T>;
  } finally {
    history.destroy();
    doc.destroy();
  }
}

export function pdfOrdinaryNoteIntentTouchesBlockedIds<T extends NoteIdentity>(
  intent: PdfOrdinaryNoteIntent<T>, blocked: ReadonlySet<string>,
): boolean {
  for (const key of ["annotations", "tombstones"] as const) {
    const before = byId(intent.before[key]);
    const after = byId(intent.after[key]);
    for (const id of blocked) if (!jsonEquals(before[id], after[id])) return true;
  }
  return false;
}

/** A durable Undo deletes a /Text object through the existing native directive;
 * optimistic failure rollback above must never fabricate a persisted deletion. */
export function revertPdfOrdinaryNoteHistory<T extends NativePdfPageAnnotation>(
  current: PdfOrdinaryNoteSlice<T>,
  intent: PdfOrdinaryNoteIntent<T>,
): PdfOrdinaryNoteSlice<T> {
  const target = revertPdfOrdinaryNoteIntent(current, intent);
  const retainedIds = new Set(target.annotations.map(note => note.id));
  target.tombstones = target.tombstones.filter(note => !retainedIds.has(note.id));
  for (const note of current.annotations) {
    if (retainedIds.has(note.id)) continue;
    const tombstones = appendManagedNativeNoteTombstone(target.tombstones, note);
    if (!tombstones) throw new Error("The PDF note deletion cannot be preserved.");
    target.tombstones = tombstones;
  }
  // List order is not a user's note mutation. A guarded inverse may rebuild
  // the array without changing any stable ID; keep that a true no-op.
  if (jsonEquals(byId(current.annotations), byId(target.annotations)) &&
    jsonEquals(byId(current.tombstones), byId(target.tombstones))) return structuredClone(current);
  return target;
}

/** The peer-rebased journal entrance is only for stable managed /NM notes.
 * Imported object-reference writers retain the exact native receipt path. */
export function pdfOrdinaryNoteIntentUsesManagedIds<T extends NoteIdentity>(intent: PdfOrdinaryNoteIntent<T>): boolean {
  for (const key of ["annotations", "tombstones"] as const) {
    const before = byId(intent.before[key]);
    const after = byId(intent.after[key]);
    for (const id of new Set([...Object.keys(before), ...Object.keys(after)])) {
      if (jsonEquals(before[id], after[id])) continue;
      for (const value of [before[id], after[id]]) {
        const note = value as unknown as NoteIdentity | undefined;
        if (note?.nativeSource || note?.managedMarkup || note?.managedProofing) return false;
      }
    }
  }
  return true;
}
