/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The cell remark and protection writes behind the Review tab (Notes,
 * Comments, Protect Sheet / Workbook, Lock Cell) and Insert > Comment. Moved
 * out of SpreadsheetWindow so the window and the spreadsheet editor commands
 * (lib/editor-commands/spreadsheet) run the same code; behaviour is
 * unchanged. Receipts, the live-sheet checks, the password prompt, focus and
 * the comment surfaces stay in the window.
 */

import { cellKind, copyCellMap, hasCellMetadata, type CellMap, type SpreadsheetCell } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  appendReply,
  hasNote,
  isNoteComment,
  isThreadComment,
  setNote,
  toggleNoteVisible,
  toggleResolved,
  type CellComment,
  type CommentMessage,
} from "@mrnote/core/web/lib/spreadsheet/comments.ts";
import {
  workbookStructureProtectionPassword,
  type SheetProtectionOptions,
  type WorkbookStructureProtection,
} from "@mrnote/core/web/lib/spreadsheet/protection.ts";

/** The part of a sheet a remark write touches. */
export interface RemarkSheet {
  cells: CellMap;
}

/**
 * Apply a threaded-comment transition to one cell and re-prune the cell. The
 * transition is computed by a pure helper in lib/spreadsheet/comments off the
 * cell's CURRENT comment, so migrate-legacy / append-reply / toggle-resolve /
 * edit-draft all funnel through one cell-rewrite path. When the resulting
 * comment is `undefined` AND the cell has no value or other metadata, the cell
 * is dropped so a note-less cell stays byte-identical to before.
 */
export function applyCellCommentChange<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
  transition: (comment: CellComment | undefined) => CellComment | undefined,
): S {
  const key = `${row}:${col}`;
  const nextCells = copyCellMap(sheet.cells);
  const previous = nextCells[key] || { value: "", kind: "text" as const };
  const nextComment = transition(previous.comment);
  const nextCell: SpreadsheetCell = {
    ...previous,
    value: String(previous.value || ""),
    kind: previous.kind || cellKind(String(previous.value || "")),
  };
  if (nextComment) {
    nextCell.comment = nextComment;
    nextCells[key] = nextCell;
  } else {
    delete nextCell.comment;
    if (nextCell.value || hasCellMetadata(nextCell)) {
      nextCells[key] = nextCell;
    } else {
      delete nextCells[key];
    }
  }
  return { ...sheet, cells: nextCells };
}

/**
 * Commit the cell's pending draft (or an explicit text) as a new reply message.
 * A blank draft is a no-op (no empty message is appended). `ts` is supplied by
 * the caller so the pure helper stays clock-free / deterministic in tests.
 */
export function appendCellCommentReply<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
  text: string,
  author?: string,
  ts?: number,
  actorId?: string,
  messageId?: string,
): S {
  return applyCellCommentChange(sheet, row, col, (comment) =>
    appendReply(
      isThreadComment(comment) ? comment : undefined,
      text,
      author,
      ts,
      actorId,
      messageId,
    ),
  );
}

/** Toggle the resolved flag on a cell's comment thread (kept, never deleted). */
export function toggleCellCommentResolved<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
): S {
  return applyCellCommentChange(sheet, row, col, (comment) =>
    toggleResolved(isThreadComment(comment) ? comment : undefined),
  );
}

export function setCellNote<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
  text: string,
  author?: string,
): S {
  return applyCellCommentChange(sheet, row, col, (comment) => {
    const previous = isNoteComment(comment) ? comment : undefined;
    return setNote(text, previous ? previous.author : author, previous);
  });
}

/**
 * New Note / Edit Note committed: the note's text written to the cell, or
 * null when the cell holds a threaded comment, a remark kept in native
 * custody, or no note and nothing to write. A note emptied of text goes.
 */
export function sheetWithCellNote<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
  text: string,
  author?: string,
): S | null {
  const comment = sheet.cells[`${row}:${col}`]?.comment;
  if (isThreadComment(comment) || comment?.nativeBlocker) return null;
  if (!text.trim() && !hasNote(comment)) return null;
  return setCellNote(sheet, row, col, text, author);
}

/**
 * Comments > Post: a new thread on a cell without one, or a reply on an
 * open thread. Null for a blank message, a cell that holds a note, a
 * resolved thread or a remark kept in native custody.
 */
export function sheetWithPostedComment<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
  text: string,
  author?: string,
  ts?: number,
  actorId?: string,
  messageId?: string,
): S | null {
  const comment = sheet.cells[`${row}:${col}`]?.comment;
  if (comment?.nativeBlocker) return null;
  if (!text.trim() || isNoteComment(comment)) return null;
  if (isThreadComment(comment) && comment.done) return null;
  return appendCellCommentReply(sheet, row, col, text, author, ts, actorId, messageId);
}

/** Comments > Resolve / Reopen: null when the cell holds no thread. */
export function sheetWithCommentResolveToggled<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
): S | null {
  const comment = sheet.cells[`${row}:${col}`]?.comment;
  if (comment?.nativeBlocker || !isThreadComment(comment)) return null;
  return toggleCellCommentResolved(sheet, row, col);
}

/**
 * Review > Delete: the cell's threaded comment, or its note when it holds a
 * note and no thread. Null when there is neither, or the remark is kept in
 * native custody.
 */
export function sheetWithoutCellRemark<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
): S | null {
  const comment = sheet.cells[`${row}:${col}`]?.comment;
  if (comment?.nativeBlocker) return null;
  if (!isThreadComment(comment) && !hasNote(comment)) return null;
  return applyCellCommentChange(sheet, row, col, () => undefined);
}

/** Notes > Show/Hide Note: null when the cell holds no note. */
export function sheetWithNoteVisibilityToggled<S extends RemarkSheet>(
  sheet: S,
  row: number,
  col: number,
): S | null {
  const comment = sheet.cells[`${row}:${col}`]?.comment;
  if (comment?.nativeBlocker || !hasNote(comment)) return null;
  return applyCellCommentChange(sheet, row, col, toggleNoteVisible);
}

/**
 * Whether a thread message is the given actor's own: by MRAI actor id when
 * the message has one; never an imported Office message; otherwise by the
 * author name, ignoring case. Only the author may edit or delete it.
 */
export function commentMessageOwnedBy(
  message: CommentMessage,
  actorId: string | undefined,
  actorName: string | undefined,
): boolean {
  if (message.actorId) return message.actorId === actorId;
  if (message.nativeId || message.personId) return false;
  return Boolean(
    message.author &&
      actorName &&
      message.author.trim().toLocaleLowerCase() ===
        actorName.toLocaleLowerCase(),
  );
}

/** A sheet's protection fields (the stored permissions are only ever true). */
export interface ProtectionFields {
  protected?: boolean;
  protectionOptions?: SheetProtectionOptions;
}

/**
 * Protect Sheet's OK: the sheet protected with the dialog's permissions and
 * optional password verifier (absent options = Excel's defaults). Null when
 * the sheet is already protected.
 */
export function sheetWithProtection<S extends ProtectionFields>(
  sheet: S,
  options: SheetProtectionOptions | undefined,
): S | null {
  if (sheet.protected === true) return null;
  const nextSheet: S = { ...sheet, protected: true };
  if (options) {
    (nextSheet as ProtectionFields).protectionOptions = options;
  } else {
    delete (nextSheet as ProtectionFields).protectionOptions;
  }
  return nextSheet;
}

/** Unprotect Sheet: protection and its permissions removed. */
export function sheetWithoutProtection<S extends ProtectionFields>(
  sheet: S,
): S {
  const nextSheet: S = { ...sheet };
  delete (nextSheet as ProtectionFields).protected;
  delete (nextSheet as ProtectionFields).protectionOptions;
  return nextSheet;
}

/**
 * Protect Workbook (structure) toggled: protected without a password, or
 * unprotected (undefined). Null when the workbook keeps a native protection
 * the editor cannot change. A password-protected structure must have its
 * password verified by the caller before unprotecting.
 */
export function workbookStructureProtectionToggled(
  current: WorkbookStructureProtection | undefined | null,
): { next: WorkbookStructureProtection | undefined } | null {
  if (current?.nativeBlocker) return null;
  // A password-protected structure is unprotected only after the password
  // is verified, which this function cannot do.
  const active = current?.locked === true || current?.windowsLocked === true;
  if (active && workbookStructureProtectionPassword(current)) return null;
  return {
    next: active ? undefined : { locked: true, passwordKind: "none" },
  };
}
