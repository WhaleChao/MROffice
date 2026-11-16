/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The text the PDF editor shows for a failed edit. Most failures carry their
 * own sentence (a planner's refusal); an engine that could not answer --
 * offline, busy, timed out, refused at the door -- reads from the message
 * catalog in the user's language (`engineHost.*`, console-notebooks), whether
 * it arrives as the engine's error or as the sentence the PDF engine session
 * carried in a refusal. The error itself is untouched: callers keep its class
 * and `reason`.
 */
import { EngineRefusedError, EngineUnavailableError, engineErrorMessageKey } from "@mrnote/core/web/lib/engine-host/client.ts";
import { pdfEngineFailureMessageKey } from "./pdf-native-remote";
import { MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE, MRAI_NATIVE_TEXT_MOVE_REFUSALS } from "./engine/native/vocabulary";
import { PDF_FALLBACK_NOT_FOUND_MESSAGE, PDF_FALLBACK_TEXT_BOX_SOURCE_MESSAGE } from "./source-text-fallback-copy";
import { pdfReaderRefusalText } from "./reader-refusal-text";



const MOVE_REFUSAL_KEYS: ReadonlyMap<string, string> = new Map(
  Object.entries(MRAI_NATIVE_TEXT_MOVE_REFUSALS).map(([cause, message]) => [message,
    cause === "clipLanding" ? "pdf.nativeTextEdit.moveRefusedClipLanding"
      : cause === "distance" ? "pdf.nativeTextEdit.moveRefusedDistance"
        : cause === "unchanged" ? "pdf.nativeTextEdit.moveRefusedUnchanged"
          : cause === "withText" ? "pdf.pageWindow.moveWithTextUnavailable"
            : "pdf.nativeTextEdit.moveRefusedStructure"]),
);

/** The preparation refusals that mean the text's stored state no longer
 * matches the bytes (a box whose last save did not land): a sentence for
 * the reader, never the engine's English. */
const PLANNER_SELECTION_REFUSAL = "Selected text does not resolve to one native operator";
const SOURCE_STATE_REFUSALS = new Set(["Appearance mismatch", "Source text owner mismatch", "Source mismatch"]);

const WORKING_SESSION_RECONCILIATION = new Set([
  "PDF save requires reconciliation before retry.",
  "Queued PDF operations require history or reconciliation.",
]);

const WORKING_SESSION_HELD_TOO_MUCH = new Set([
  "PDF pending operation memory budget is exhausted.",
  "PDF pending edits must finish saving before this operation can be staged.",
]);

export function pdfErrorText(error: unknown, t: (key: string) => string, fallback: string): string {
  if (error instanceof EngineUnavailableError || error instanceof EngineRefusedError) {
    return t(engineErrorMessageKey(error));
  }
  if (!(error instanceof Error) || !error.message) return fallback;
  // The page window's read of the source's metadata (a busy or restarting
  // server, a dropped connection) carried its English onto every format
  // button (walk fix-s1c: "PDF source metadata request timed out").
  if (error.name === "PdfNativeSourceMetadataTimeoutError") return t("engineHost.busy");
  if (error.name === "PdfNativeSourceMetadataTransportError") return t("engineHost.needsConnection");
  if (SOURCE_STATE_REFUSALS.has(error.message)) return t("pdf.nativeTextEdit.sourceStateChanged");
  // The working session's own refusals while an earlier save is unsettled
  
  // English on a move's toast).
  if (WORKING_SESSION_RECONCILIATION.has(error.message)) return t("pdf.pageWindow.previousPdfWriteStillRequires");
  // The unsaved edits a window may hold are bounded in memory (twice the
  // largest file): past that the next edit is not taken until they are
  // saved. Its English, or a sentence that named no way out, reached the
  
  if (WORKING_SESSION_HELD_TOO_MUCH.has(error.message)) return t("pdf.pageWindow.heldEditsTooLarge");
  // The working copy's write authority moved under an edit (the page's
  // version or access changed): its English reached the reader as-is
  
  if (error.message === "PDF working session authority changed.") return t("pdf.pageWindow.workingSessionChanged");
  // The planner's summary of a selection no route could write (the causes it
  // heard follow, in English, for the engine's logs): one sentence for the
  // reader (walk t-s6-b: a colour on the 装配理论 paragraph).
  
  // (alone, or as a cause in the planner's summary below).
  if (error.message.includes(MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE)) return t("pdf.nativeTextEdit.glyphsReadAsBlanks");
  if (error.message.startsWith(PLANNER_SELECTION_REFUSAL)) return t("pdf.nativeTextEdit.selectionNotRewritable");
  
  // form object drawn twice): its English went onto every format button of
  
  if (error.message === PDF_FALLBACK_NOT_FOUND_MESSAGE) return t("pdf.nativeTextEdit.noGlyphWhereText");
  // A line under a text box whose record cannot be read: the same sentence
  // (how the text is drawn cannot be read), not the engine's English.
  if (error.message === PDF_FALLBACK_TEXT_BOX_SOURCE_MESSAGE) return t("pdf.nativeTextEdit.noGlyphWhereText");
  const moveRefusal = MOVE_REFUSAL_KEYS.get(error.message);
  if (moveRefusal) return t(moveRefusal);
  const key = pdfEngineFailureMessageKey(error.message);
  return key ? t(key) : pdfReaderRefusalText(error.message, t);
}
