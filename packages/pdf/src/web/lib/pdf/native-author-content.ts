/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeTextPropertyDraft, PdfNativeTextPropertyState } from "./engine/native/list-text-fidelity";
import type { PageContentDocument } from "@mrnote/core/web/hooks/usePageContentSync.ts";
import { pageDocumentJsonEquals, pageStoredJsonEquals, pdfPatchbackReceiptDataValue } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import type { PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";

// Native source changes become collaborative only through a verified server
// source/receipt. A pending head must not hitch a ride on a peer's notes save.
export const PDF_NATIVE_AUTHOR_PRIVATE_KEYS = Object.freeze(["pdfNativeTransactions"]);

/** Called synchronously INSIDE saveTransformedContent. The current author's
 * notes and other fields are already merged by the shared save owner. A PDF
 * compiler may append its one operation only if that native source prefix is
 * still current; a newer native head needs reconciliation, not a JSON merge. */
export function appendPdfNativeAuthorTransaction(
  currentAuthor: PageContentDocument,
  expectedJournal: JsonValue | undefined,
  nextJournal: JsonValue,
): PageContentDocument | null {
  const before = expectedJournal === undefined ? [] : expectedJournal;
  if (!Array.isArray(before) || !Array.isArray(nextJournal) ||
    nextJournal.length !== before.length + 1 ||
    !pageStoredJsonEquals(nextJournal.slice(0, -1), before)) return null;
  const ids = new Set<string>();
  for (const record of nextJournal) {
    if (!record || typeof record !== "object" || Array.isArray(record) ||
      typeof record.id !== "string" || !record.id.trim() || ids.has(record.id)) return null;
    ids.add(record.id);
  }
  // An explicit retry reuses the exact operation id and bytes. Its admitted
  // native head may already be in shared recovery while newer notes follow it.
  // Resubmit that same head; this equality is not a saved acknowledgement.
  const held = currentAuthor.pdfNativeTransactions;
  if (pageDocumentJsonEquals(held, nextJournal)) {
    return { ...currentAuthor, pdfNativeTransactions: structuredClone(nextJournal) };
  }
  // The page holds the journal as the API stored it (U+0000 as U+FFFD,
  // pageStoredJsonProjection); the caller's copy may still be the one it
  // authored. The append keeps the page's own prefix, so the save owner sees
  // exactly one record added to the journal it holds.
  if (!pageStoredJsonEquals(held, expectedJournal)) {
    return pageStoredJsonEquals(held, nextJournal)
      ? { ...currentAuthor, pdfNativeTransactions: structuredClone(held) }
      : null;
  }
  const prefix = Array.isArray(held) ? held : [];
  return {
    ...currentAuthor,
    pdfNativeTransactions: structuredClone([...prefix, nextJournal.at(-1)!]),
  };
}

export function pdfNativeTextPropertyDraftFromState(
  state: PdfNativeTextPropertyState,
): PdfNativeTextPropertyDraft {
  return Object.freeze({
    fontResourceName: state.font.resourceName,
    fontSize: state.fontSize,
    color: state.color,
    alignment: state.alignment,
    characterSpacing: state.characterSpacing,
    wordSpacing: state.wordSpacing,
    lineSpacing: state.lineSpacing,
    leftIndent: state.leftIndent,
    rightIndent: state.rightIndent,
    firstLineIndent: state.firstLineIndent,
    horizontalScale: state.horizontalScale,
    textRise: state.textRise,
    textMatrixLinear: state.textMatrixLinear,
  });
}

/** The move is authenticated by the record's own offset and output bytes, not
 * by the property draft. Comparing a draft against a saved state must drop it,
 * and must rebuild the draft in the mapping's own key order. */
export function pdfNativeTextPropertyDraftWithoutMove(
  draft: PdfNativeTextPropertyDraft,
): PdfNativeTextPropertyDraft {
  return Object.freeze({
    fontResourceName: draft.fontResourceName,
    fontSize: draft.fontSize,
    color: draft.color,
    alignment: draft.alignment,
    characterSpacing: draft.characterSpacing,
    wordSpacing: draft.wordSpacing,
    lineSpacing: draft.lineSpacing,
    leftIndent: draft.leftIndent,
    rightIndent: draft.rightIndent,
    firstLineIndent: draft.firstLineIndent,
    horizontalScale: draft.horizontalScale,
    textRise: draft.textRise,
    textMatrixLinear: draft.textMatrixLinear,
    ...(draft.syntheticBold !== undefined ? { syntheticBold: draft.syntheticBold } : {}),
    // The rules a Clear Formatting takes away, or the Format Painter draws,
    // ride the draft to the write.
    ...(draft.dropTextRules?.length ? { dropTextRules: draft.dropTextRules } : {}),
    ...(draft.addTextRules?.length ? { addTextRules: draft.addTextRules } : {}),
    // Rules pressed on characters of the run ride the same write.
    ...(draft.textRuleParts?.length ? { textRuleParts: draft.textRuleParts } : {}),
    ...(draft.textRuleRanges?.length ? { textRuleRanges: draft.textRuleRanges } : {}),
    // Formats of part of the run split it: they are part of what is drawn.
    ...(draft.segments?.length ? { segments: draft.segments } : {}),
  });
}

/** A room note update is a view/author delta, not new PDF byte authority.
 * Only an unchanged saved identity and an already-owned journal may preserve
 * the active native editor. All other changes take the full reload path. */
export function pdfRemoteUpdateChangesOnlyNotes(
  current: PageItem | null,
  incoming: PageItem,
  admittedJournal?: JsonValue,
): boolean {
  if (!current || current.id !== incoming.id || current.notebook_id !== incoming.notebook_id ||
    current.page_type !== "pdf" || incoming.page_type !== "pdf" ||
    !Number.isSafeInteger(current.content_version) || Number(current.content_version) < 1 ||
    current.content_version !== incoming.content_version ||
    !pageDocumentJsonEquals(pdfPatchbackReceiptDataValue(current), pdfPatchbackReceiptDataValue(incoming)) ||
    !current.content_json || !incoming.content_json) return false;
  const previous = current.content_json;
  const next = incoming.content_json;
  const previousJournal = previous.pdfNativeTransactions === undefined ? [] : previous.pdfNativeTransactions;
  if (!pageStoredJsonEquals(previous.pdfNativeTransactions, next.pdfNativeTransactions) &&
    (admittedJournal === undefined || !pageStoredJsonEquals(admittedJournal, next.pdfNativeTransactions) ||
      !Array.isArray(admittedJournal) ||
      !Array.isArray(previousJournal) || admittedJournal.length !== previousJournal.length + 1 ||
      !pageStoredJsonEquals(admittedJournal.slice(0, -1), previousJournal))) return false;
  // The importer omits optional empty slices; buildCurrentContent materializes
  // these same defaults. A peer doing that does not change PDF bytes. Do not
  // generalize this to deletion-bearing collections (page labels/imports), or
  // let an empty canonical field mask a legacy alias containing source state.
  const emptyDefaults: Record<string, string | null> = {
    cropBoxes: "crop_boxes", formFields: "form_fields", imageOverlays: "image_overlays",
    inkStrokes: "ink_strokes", pageSizes: null, textBoxEdits: "text_box_edits",
    pdfFormAuthoringOperations: null,
  };
  const withoutNotes = (content: Record<string, unknown>) => Object.fromEntries(
    Object.entries(content).filter(([key, value]) => {
      if (key === "annotations" || key === "pdfNativeTransactions") return false;
      if (!Object.hasOwn(emptyDefaults, key) || !Array.isArray(value) || value.length !== 0) return true;
      const alias = emptyDefaults[key];
      return alias !== null && (Object.hasOwn(previous, alias) || Object.hasOwn(next, alias));
    }));
  return pageDocumentJsonEquals(withoutNotes(previous), withoutNotes(next));
}
