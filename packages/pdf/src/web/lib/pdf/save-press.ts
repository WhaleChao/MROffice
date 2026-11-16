/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type WindowSaveStatus = "saved" | "saving" | "unsaved" | "conflict";

/** How often the window is read while a press settles. */
export const PDF_SAVE_PRESS_POLL_MS = 200;
/** Held work still unsaved this long after the try is tried once more. */
export const PDF_SAVE_PRESS_RERUN_AFTER_MS = 600;
/** A window unsaved this long with nothing left to try is told why. */
export const PDF_SAVE_PRESS_SAY_AFTER_MS = 2_500;
/** A write still running after this long is left to its own status. */
export const PDF_SAVE_PRESS_GIVE_UP_MS = 15_000;
/** 已保存 read this many times in a row is the press's answer. */
export const PDF_SAVE_PRESS_SAVED_READS = 3;

export type PdfSavePressStep =
  /** Read the window again in a moment. */
  | "wait"
  /** Everything is written: the press saved. */
  | "saved"
   
  | "rerun"
  /** Nothing more to try and the window is unsaved: say why. */
  | "say"
  /** Something else is already saying it (a conflict, a write still running). */
  | "quiet";

export function pdfSavePressStep(seen: Readonly<{
  /** What the window says (its shown status). */
  status: WindowSaveStatus;
  /** The page's live room holds every edit: the window reads 已同步. */
  synced: boolean;
  /** Reads of 已保存 in a row, this one included. */
  savedReads: number;
  /** How long the window has read 未保存 without a break. */
  unsavedMs: number;
  /** Since the press's own try ended. */
  waitedMs: number;
   
  dirty: boolean;
  /** The one further try is still to be made (and this press may make it). */
  rerunLeft: boolean;
  /** Rows kept as the page's drafts are all that is unsaved: no save is coming. */
  draftsOnly: boolean;
}>): PdfSavePressStep {
  if (seen.status === "conflict" || seen.synced) return "quiet";
  if (seen.status === "saved") return seen.savedReads >= PDF_SAVE_PRESS_SAVED_READS ? "saved" : "wait";
  if (seen.status === "saving") return seen.waitedMs >= PDF_SAVE_PRESS_GIVE_UP_MS ? "quiet" : "wait";
  if (seen.draftsOnly) return "say";
  if (seen.dirty && seen.rerunLeft) return seen.unsavedMs >= PDF_SAVE_PRESS_RERUN_AFTER_MS ? "rerun" : "wait";
  return seen.unsavedMs >= PDF_SAVE_PRESS_SAY_AFTER_MS || seen.waitedMs >= PDF_SAVE_PRESS_GIVE_UP_MS ? "say" : "wait";
}
