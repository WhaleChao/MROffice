/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextSelectionProof, PdfNativeTextPropertyDraft } from "../engine/native/list-text-fidelity";
import { pdfNativeTextPropertyDraftFromState } from "../native-author-content";
import {
  mraiNativeTextGroupTransformIsIdentity,
  readMraiNativeTextGroupMoveDraft,
} from "../engine/native/text-group-move-model";
import { readMraiNativeTextMoveDraft } from "../engine/native/text-move-model";

export type PdfNativeTextDraftMove = Readonly<{
  /** The map of the whole box a group proof covers (text-group-move.ts). */
  group: ReturnType<typeof readMraiNativeTextGroupMoveDraft>;
  /** The move of the text as the page draws it, for an item only a move
   * proof covers (text-move.ts); null when the box is a group. */
  move: Readonly<{ x: number; y: number }> | null;
  /** The draft takes the box anywhere (a map that is not the identity, a
   * move that is not zero). */
  changed: boolean;
}>;

/**
 * The move a box's draft asks, read one way wherever a Leave, its preview or
 * its "nothing changed" test needs it (six readers had their own copy).
 * Throws as the readers do for a draft that changes what such a proof keeps.
 */
export function readPdfNativeTextDraftMove(
  draft: PdfNativeTextPropertyDraft | null,
  proof: PdfNativeListTextSelectionProof | null,
): PdfNativeTextDraftMove {
  const group = readMraiNativeTextGroupMoveDraft(draft, proof,
    proof?.properties ? pdfNativeTextPropertyDraftFromState(proof.properties.original) : null);
  const move = group ? null : readMraiNativeTextMoveDraft(draft, proof);
  const changed = group ? !mraiNativeTextGroupTransformIsIdentity(group)
    : Boolean(move && (move.x !== 0 || move.y !== 0));
  return Object.freeze({ group, move, changed });
}

/** As readPdfNativeTextDraftMove; a draft the proof refuses reads as no move. */
export function readPdfNativeTextDraftMoveOrNone(
  draft: PdfNativeTextPropertyDraft | null,
  proof: PdfNativeListTextSelectionProof | null,
): PdfNativeTextDraftMove {
  let group: PdfNativeTextDraftMove["group"] = null;
  try {
    group = readMraiNativeTextGroupMoveDraft(draft, proof,
      proof?.properties ? pdfNativeTextPropertyDraftFromState(proof.properties.original) : null);
  } catch { /* no group map */ }
  let move: PdfNativeTextDraftMove["move"] = null;
  if (!group) {
    try { move = readMraiNativeTextMoveDraft(draft, proof); } catch { /* no move */ }
  }
  const changed = group ? !mraiNativeTextGroupTransformIsIdentity(group)
    : Boolean(move && (move.x !== 0 || move.y !== 0));
  return Object.freeze({ group, move, changed });
}
