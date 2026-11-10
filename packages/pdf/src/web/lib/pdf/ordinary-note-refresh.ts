/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import { mergePageDocumentDelta, type PageContentDocument } from "@mrnote/core/web/hooks/usePageContentSync.ts";

const EMPTY_SLICES: Record<string, string | null> = {
  formFields: "form_fields", pageTransforms: "page_transforms", deletedPages: "deleted_pages",
  insertedPages: "inserted_pages", pageMoves: "page_moves", textBoxEdits: "text_box_edits",
  imageOverlays: "image_overlays", inkStrokes: "ink_strokes", pageSizes: null,
  cropBoxes: "crop_boxes", pdfFormAuthoringOperations: null,
};

function notesById(content: Record<string, unknown>): Record<string, Record<string, unknown>> | null {
  if (content.annotations !== undefined && !Array.isArray(content.annotations)) return null;
  const result: Record<string, Record<string, unknown>> = Object.create(null);
  for (const note of (content.annotations ?? []) as unknown[]) {
    if (!note || typeof note !== "object" || Array.isArray(note)) return null;
    const value = note as Record<string, unknown>;
    if (typeof value.id !== "string" || !value.id || Object.hasOwn(result, value.id)) return null;
    result[value.id] = value;
  }
  return result;
}

function withoutNotes(content: Record<string, unknown>, peers: Record<string, unknown>[]): Record<string, unknown> {
  return Object.fromEntries(Object.entries(content).filter(([key, value]) => {
    if (key === "annotations") return false;
    if (!Object.hasOwn(EMPTY_SLICES, key) || !Array.isArray(value) || value.length !== 0) return true;
    const alias = EMPTY_SLICES[key];
    return alias !== null && peers.some(peer => Object.hasOwn(peer, alias));
  }));
}

/** Semantic admission only. The caller must separately authenticate the actual
 * ResourceVersion bytes and install its prepared reader before acknowledging. */
export function preparePdfOrdinaryNoteRefresh(
  sourcePage: PageItem, incoming: PageItem, live: Record<string, unknown>,
): PageContentDocument | null {
  const before = sourcePage.content_json;
  const after = incoming.content_json;
  if (!before || !after || sourcePage.id !== incoming.id ||
    sourcePage.notebook_id !== incoming.notebook_id || sourcePage.title !== incoming.title ||
    sourcePage.page_type !== "pdf" || incoming.page_type !== "pdf" ||
    !Number.isSafeInteger(sourcePage.content_version) || Number(sourcePage.content_version) < 1 ||
    !Number.isSafeInteger(incoming.content_version) || Number(incoming.content_version) <= Number(sourcePage.content_version) ||
    typeof before.dataItemId !== "string" || before.dataItemId !== after.dataItemId) return null;
  const peers = [before, after, live];
  const native = withoutNotes(before, peers);
  if (!pageDocumentJsonEquals(native, withoutNotes(after, peers)) ||
    !pageDocumentJsonEquals(native, withoutNotes(live, peers))) return null;
  const oldNotes = notesById(before), nextNotes = notesById(after), liveNotes = notesById(live);
  if (!oldNotes || !nextNotes || !liveNotes) return null;
  for (const id of new Set([...Object.keys(oldNotes), ...Object.keys(nextNotes)])) {
    if (pageDocumentJsonEquals(oldNotes[id], nextNotes[id])) continue;
    for (const note of [oldNotes[id], nextNotes[id]]) {
      if (!note) continue;
      // This entrance cannot adopt another native writer's metadata as notes.
      if (note.managedMarkup || note.managedProofing || note.nativeSource ||
        (typeof note.text === "string" && note.text.startsWith("[[MRNOTE_") &&
          !note.text.startsWith("[[MRNOTE_NATIVE_NOTE_V1|"))) return null;
    }
    if (!pageDocumentJsonEquals(liveNotes[id], oldNotes[id]) &&
      !pageDocumentJsonEquals(liveNotes[id], nextNotes[id])) return null;
  }
  return mergePageDocumentDelta(live as PageContentDocument, after as PageContentDocument, before as PageContentDocument);
}











export function pdfOrdinaryPeerRefreshPlan(input: {
  sourcePage: PageItem; incoming: PageItem; live: Record<string, unknown>;
  requestedVersion: number; reloadedVersion: number | null;
}): { kind: "settled" } | { kind: "merge"; content: PageContentDocument } | { kind: "reload" } | { kind: "wait" } {
  const { sourcePage, incoming, live, requestedVersion, reloadedVersion } = input;
  const version = Number(incoming.content_version);
  if (!(version > Number(sourcePage.content_version))) return { kind: "settled" };
  if (version < requestedVersion) return { kind: "wait" };
  const content = preparePdfOrdinaryNoteRefresh(sourcePage, incoming, live);
  if (content) return { kind: "merge", content };
  return reloadedVersion === version ? { kind: "wait" } : { kind: "reload" };
}
