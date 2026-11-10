/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeTextPropertyDraft } from "./list-text-fidelity";
import { PdfNativeTransactionError } from "./transaction-error";
import type { PdfNativeTextMoveProof } from "./text-move";

export function invalid(message: string): never {
  throw new PdfNativeTransactionError("invalid-request", message);
}

/** A draft's own state in one key order, without its move. */
export function draftState(draft: PdfNativeTextPropertyDraft): string {
  return JSON.stringify([
    draft.fontResourceName,
    draft.fontSize,
    draft.color,
    draft.alignment,
    draft.characterSpacing,
    draft.wordSpacing,
    draft.lineSpacing,
    draft.leftIndent,
    draft.rightIndent,
    draft.firstLineIndent,
    draft.horizontalScale,
    draft.textRise ?? 0,
    draft.textMatrixLinear ?? null,
  ]);
}

/**
 * The move an editor draft asks of an item that only a move proof covers:
 * null when the draft is not such a move (a property proof owns it, or the
 * item has no move proof). A draft that changes anything but the position is
 * refused: the item's text is drawn as the page draws it, and only a move
 * keeps that exactly.
 */
export function readMraiNativeTextMoveDraft(
  draft: PdfNativeTextPropertyDraft | null,
  proof: Readonly<{
    properties: unknown;
    move?: PdfNativeTextMoveProof;
  }> | null,
): Readonly<{ x: number; y: number }> | null {
  if (!draft || !proof || proof.properties || !proof.move) return null;
  if (draftState(draft) !== draftState(proof.move.baseline)) {
    invalid(
      "Only the position of this text can change here; it keeps the font, size and colour the page draws it in.",
    );
  }
  return Object.freeze({ x: draft.moveX ?? 0, y: draft.moveY ?? 0 });
}
