/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What one edit of a box is, read once for its preview and its Leave
 * (docs/pdf/text-edit-unified-model.md §2.1): the way it is written follows
 * from it. They read it apart -- the preview of a move and its frame, or a
 * move and its words, chose nothing while the Leave wrote both -- and so
 * showed one thing and wrote another.
 */
export interface PdfNativeTextEditShape {
  /** The words differ from the box's. */
  readonly textChanged: boolean;
  /** A run's format draft changes it (pdfNativePropertyDraftChanges). */
  readonly formatChanged: boolean;
  /** Formats asked of a box no run property states (its rows' or paragraph's). */
  readonly boxFormat: boolean;
  /** The box moved along the page or as a whole (its group). */
  readonly moved: boolean;
  /** Its paragraph set again in a resized frame. */
  readonly frame: boolean;
  /** Its paragraph's own settings (alignment, indents, spacing). */
  readonly paragraphSettings: boolean;
   
  readonly layout: boolean;
}

export type PdfNativeTextEditPath =
  | "unchanged"
  | "layout"
  | "paragraph-settings"
  | "move-then-frame"
  | "move-then-words"
  | "move"
  | "frame"
  | "box-format"
  | "format"
  | "words";

/** The way an edit of this shape is written, and previewed. */
export function pdfNativeTextEditPath(edit: PdfNativeTextEditShape): PdfNativeTextEditPath {
  if (edit.layout) return "layout";
  if (edit.paragraphSettings) return "paragraph-settings";
  if (edit.moved) return edit.frame ? "move-then-frame" : edit.textChanged ? "move-then-words" : "move";
  if (edit.frame) return "frame";
  if (edit.boxFormat) return "box-format";
  if (edit.formatChanged) return "format";
  return edit.textChanged ? "words" : "unchanged";
}
