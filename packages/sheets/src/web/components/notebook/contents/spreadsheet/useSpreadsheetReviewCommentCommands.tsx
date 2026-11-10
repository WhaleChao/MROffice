"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { sheetWithPostedComment, sheetWithCommentResolveToggled, applyCellCommentChange, commentMessageOwnedBy } from "../../../../lib/spreadsheet/review-buttons";
import { useCallback, type RefObject } from "react";
import { spreadsheetReviewCommentWriterReceiptMatches } from "../../../../lib/spreadsheet/visible-command-authority";
import { isThreadComment, type CommentMessage, type ThreadComment, editCommentMessage, deleteCommentReply } from "@mrnote/core/web/lib/spreadsheet/comments.ts";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type ReviewCommentMutationReceipt, type PageSetupSelectionSnapshot, type ReviewCommentDialogAuthority, type LiveReviewCommentPageAuthority } from "./spreadsheetMutationReceipts";
import { type SheetViewMappedCellMutation, type SheetViewMappedDrawingMutation, type SheetViewMappedLineMutation, type SheetViewMappedPageMutation } from "./spreadsheetSheetViews";
import { type CommentMentionsController } from "@mrnote/core/web/components/office-comments/CommentMentions.tsx";
import { type Dispatch, type SetStateAction } from "react";

/** What useSpreadsheetReviewCommentCommands reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetReviewCommentCommandsInputs {
  closeCellCommentSurface: () => void;
  commentEditDraft: string;
  commentMentions: CommentMentionsController | null;
  isSheetProtectionEffectiveRef: RefObject<(candidate: SheetModel | null | undefined) => boolean>;
  liveActiveSheetIndexRef: RefObject<number>;
  liveSheetsRef: RefObject<SheetModel[]>;
  readPageSetupSelectionSnapshot: () => PageSetupSelectionSnapshot | null;
  replaceActiveSheet: (nextSheet: SheetModel, nextTitle?: string, calculationMutation?: "content" | "coordinate", sheetViewMutation?: SheetViewMappedCellMutation | SheetViewMappedDrawingMutation | SheetViewMappedLineMutation | SheetViewMappedPageMutation, followOnOtherSheets?: (candidate: SheetModel) => SheetModel) => SheetModel[];
  reviewCommentActorId: string;
  reviewCommentActorName: string;
  reviewCommentDialogAuthority: ReviewCommentDialogAuthority | null;
  reviewCommentMutationCommandRef: RefObject<ReviewCommentMutationReceipt | null>;
  reviewCommentPageAuthorityIsCurrent: (authority: ReviewCommentDialogAuthority | null | undefined) => boolean;
  reviewCommentPageAuthorityRef: RefObject<LiveReviewCommentPageAuthority | null>;
  setCommentDraft: Dispatch<SetStateAction<string>>;
  setCommentEditDraft: Dispatch<SetStateAction<string>>;
  setCommentEditingIndex: Dispatch<SetStateAction<number | null>>;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetReviewCommentCommands({
  closeCellCommentSurface,
  commentEditDraft,
  commentMentions,
  isSheetProtectionEffectiveRef,
  liveActiveSheetIndexRef,
  liveSheetsRef,
  readPageSetupSelectionSnapshot,
  replaceActiveSheet,
  reviewCommentActorId,
  reviewCommentActorName,
  reviewCommentDialogAuthority,
  reviewCommentMutationCommandRef,
  reviewCommentPageAuthorityIsCurrent,
  reviewCommentPageAuthorityRef,
  setCommentDraft,
  setCommentEditDraft,
  setCommentEditingIndex,
}: SpreadsheetReviewCommentCommandsInputs) {
  const executeReviewCommentReceiptWriter = useCallback(
    (
      expectedControlId: ReviewCommentMutationReceipt["controlId"],
    ): boolean => {
      const receipt = reviewCommentMutationCommandRef.current;
      try {
        if (!receipt) return false;
        const liveSelection = readPageSetupSelectionSnapshot();
        const liveWorkbook = liveSheetsRef.current;
        const liveIndex = liveActiveSheetIndexRef.current;
        const liveSheet = liveWorkbook[liveIndex];
        const livePageAuthority = reviewCommentPageAuthorityRef.current;
        if (!liveSelection || !liveSheet || !livePageAuthority?.allowed) {
          return false;
        }
        const liveComment =
          liveSheet.cells[
            `${liveSelection.active.row}:${liveSelection.active.col}`
          ]?.comment;
        const liveParameter =
          expectedControlId === "XLSCTL.REVIEW.COMMENT.RESOLVE_TOGGLE" &&
          isThreadComment(liveComment)
            ? String(!liveComment.done)
            : receipt.parameter;
        if (
          !spreadsheetReviewCommentWriterReceiptMatches({
            expectedControlId,
            receiptControlId: receipt.controlId,
            sourceWorkbook: receipt.workbook,
            sourceActiveSheetIndex: receipt.activeSheetIndex,
            sourceSheet: receipt.sheet,
            sourceSheetId: receipt.sheetId,
            sourceSelectionEpoch: receipt.selection.epoch,
            sourceRow: receipt.selection.active.row,
            sourceCol: receipt.selection.active.col,
            sourceComment: receipt.comment,
            sourceParameter: receipt.parameter,
            sourcePageId: receipt.pageId,
            sourcePageAuthorityGeneration: receipt.pageAuthorityGeneration,
            sourceActorId: receipt.actorId,
            sourceActorName: receipt.actorName,
            liveWorkbook,
            liveActiveSheetIndex: liveIndex,
            liveSheet,
            liveSheetId: liveSheet.id,
            liveSelectionEpoch: liveSelection.epoch,
            liveRow: liveSelection.active.row,
            liveCol: liveSelection.active.col,
            liveComment,
            liveParameter,
            livePageId: livePageAuthority.pageId,
            livePageAuthorityGeneration: livePageAuthority.generation,
            liveActorId: livePageAuthority.actorId,
            liveActorName: livePageAuthority.actorName,
          }) ||
          isSheetProtectionEffectiveRef.current(liveSheet) ||
          liveComment?.nativeBlocker
        ) {
          return false;
        }
        const row = liveSelection.active.row;
        const col = liveSelection.active.col;
        if (expectedControlId === "XLSCTL.REVIEW.COMMENT.POST") {
          const posted = sheetWithPostedComment(
            liveSheet,
            row,
            col,
            receipt.parameter,
            receipt.actorName,
            Date.now(),
            receipt.actorId,
            globalThis.crypto?.randomUUID?.(),
          );
          if (!posted) return false;
          replaceActiveSheet(posted);
          setCommentDraft("");
          return true;
        }
        if (
          expectedControlId === "XLSCTL.REVIEW.COMMENT.RESOLVE_TOGGLE"
        ) {
          const toggled = sheetWithCommentResolveToggled(liveSheet, row, col);
          if (!toggled) return false;
          replaceActiveSheet(toggled);
          return true;
        }
        if (!isThreadComment(liveComment)) return false;
        replaceActiveSheet(
          applyCellCommentChange(liveSheet, row, col, () => undefined),
        );
        closeCellCommentSurface();
        return true;
      } finally {
        reviewCommentMutationCommandRef.current = null;
      }
    },
    [closeCellCommentSurface, readPageSetupSelectionSnapshot, replaceActiveSheet],
  );

  const reviewCommentMessageOwnedByActor = useCallback(
    (message: CommentMessage): boolean => {
      if (!reviewCommentPageAuthorityRef.current?.allowed) return false;
      return commentMessageOwnedBy(
        message,
        reviewCommentActorId,
        reviewCommentActorName,
      );
    },
    [reviewCommentActorId, reviewCommentActorName],
  );

  const resolveLiveReviewCommentMutation = useCallback((): {
    comment: ThreadComment;
    col: number;
    liveSheet: SheetModel;
    row: number;
  } | null => {
    const authority = reviewCommentDialogAuthority;
    const pageAuthority = reviewCommentPageAuthorityRef.current;
    const liveSelection = readPageSetupSelectionSnapshot();
    const liveWorkbook = liveSheetsRef.current;
    const liveIndex = liveActiveSheetIndexRef.current;
    const liveSheet = liveWorkbook[liveIndex];
    if (
      !authority ||
      !pageAuthority?.allowed ||
      !reviewCommentPageAuthorityIsCurrent(authority) ||
      !liveSelection ||
      !liveSheet ||
      authority.workbook !== liveWorkbook ||
      authority.activeSheetIndex !== liveIndex ||
      authority.sheet !== liveSheet ||
      authority.sheetId !== liveSheet.id ||
      authority.selection.epoch !== liveSelection.epoch ||
      authority.selection.active.row !== liveSelection.active.row ||
      authority.selection.active.col !== liveSelection.active.col ||
      authority.actorId !== pageAuthority.actorId ||
      authority.actorName !== pageAuthority.actorName ||
      isSheetProtectionEffectiveRef.current(liveSheet)
    ) {
      return null;
    }
    const row = liveSelection.active.row;
    const col = liveSelection.active.col;
    const liveComment = liveSheet.cells[`${row}:${col}`]?.comment;
    if (
      liveComment !== authority.comment ||
      !isThreadComment(liveComment) ||
      liveComment.nativeBlocker
    ) {
      return null;
    }
    return { comment: liveComment, col, liveSheet, row };
  }, [
    readPageSetupSelectionSnapshot,
    reviewCommentDialogAuthority,
    reviewCommentPageAuthorityIsCurrent,
  ]);

  const startEditingReviewCommentMessage = useCallback(
    (messageIndex: number) => {
      const live = resolveLiveReviewCommentMutation();
      const message = live?.comment.thread[messageIndex];
      if (
        !live ||
        !message ||
        live.comment.done ||
        !reviewCommentMessageOwnedByActor(message)
      ) {
        return;
      }
      setCommentEditingIndex(messageIndex);
      setCommentEditDraft(message.text);
    }, [resolveLiveReviewCommentMutation, reviewCommentMessageOwnedByActor],
  );

  const saveEditedReviewCommentMessage = useCallback(
    (messageIndex: number) => {
      const live = resolveLiveReviewCommentMutation();
      const message = live?.comment.thread[messageIndex];
      if (
        !live ||
        !message ||
        live.comment.done ||
        !reviewCommentMessageOwnedByActor(message)
      ) {
        return;
      }
      const nextComment = editCommentMessage(
        live.comment,
        messageIndex,
        commentEditDraft,
      );
      if (!nextComment || nextComment === live.comment) {
        setCommentEditingIndex(null);
        setCommentEditDraft("");
        return;
      }
      replaceActiveSheet(
        applyCellCommentChange(live.liveSheet, live.row, live.col, () =>
          nextComment,
        ),
      );
      commentMentions?.notifyPosted({
        threadRef: `${live.liveSheet.id}:${live.row}:${live.col}`,
        messageRef: message.id,
        text: commentEditDraft,
        previousText: message.text,
      });
      setCommentEditingIndex(null);
      setCommentEditDraft("");
    }, [
      commentEditDraft,
      commentMentions,
      replaceActiveSheet,
      resolveLiveReviewCommentMutation,
      reviewCommentMessageOwnedByActor,
    ],
  );

  const deleteReviewCommentReply = useCallback(
    (messageIndex: number) => {
      const live = resolveLiveReviewCommentMutation();
      const message = live?.comment.thread[messageIndex];
      if (
        !live ||
        !message ||
        messageIndex <= 0 ||
        !reviewCommentMessageOwnedByActor(message)
      ) {
        return;
      }
      const nextComment = deleteCommentReply(live.comment, messageIndex);
      if (!nextComment || nextComment === live.comment) return;
      replaceActiveSheet(
        applyCellCommentChange(live.liveSheet, live.row, live.col, () =>
          nextComment,
        ),
      );
      setCommentEditingIndex(null);
      setCommentEditDraft("");
    }, [
      replaceActiveSheet,
      resolveLiveReviewCommentMutation,
      reviewCommentMessageOwnedByActor,
    ],
  );
  return {
    deleteReviewCommentReply,
    executeReviewCommentReceiptWriter,
    reviewCommentMessageOwnedByActor,
    saveEditedReviewCommentMessage,
    startEditingReviewCommentMessage,
  };
}
