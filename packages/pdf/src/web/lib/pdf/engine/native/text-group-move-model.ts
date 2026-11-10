/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeTextPropertyDraft } from "./list-text-fidelity";
import { PdfNativeTransactionError } from "./transaction-error";
import type { PdfNativeTextGroupEdit, PdfNativeTextGroupMoveProof } from "./text-group-move";

export const MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS = Object.freeze({
  source:
    "Moving this text natively needs it drawn directly by the page's own content stream.",
  proof: "Moving this text natively needs one exact proof of the text it moves.",
  unreadable: "The content stream this text is drawn in cannot be read exactly.",
  shared:
    "This text continues a line of other text, so it cannot be moved on its own.",
  clip: "This text clips what the page draws after it, so it cannot be moved on its own.",
  landing:
    "The moved text would land where the page's clip hides it, so nothing was moved.",
  transform: "The move is not a finite map of the page.",
  unchanged: "The text is already at this position and size.",
  twin: "This text is drawn again in its place in a way a move cannot follow.",
  size: "This move rewrites more of the page than one edit may.",
  restore: "The moved text no longer stands where the move wrote it, so it cannot be put back exactly.",
  partial: "Part of a line this text is drawn in belongs to other text, so it cannot be moved on its own.",
  // A box whose landing a clip of its own would hide is drawn again outside
  // every clip, at the tail of the page drawing; these keep it where that
  // cannot be shown to draw the same glyphs the same way.
  outsideState:
    "The moved text is drawn in a state the page sets in a way that cannot be carried out of its clip, so nothing was moved.",
  outsideMarked:
    "The moved text is marked as a replacement text or a layer that cannot be carried out of its clip, so nothing was moved.",
  outsideResource:
    "The moved text is drawn by a Type3 font or under a soft mask that cannot be carried out of its clip, so nothing was moved.",
  // A synthetic bold of the box (2 Tr in its fill colour) puts back the state
  // in force after each show; these keep it where that cannot be done.
  boldWidth:
    "This text is stroked under a width the page sets in a way that cannot be put back, so its weight cannot change here.",
  boldMode: "This text is not drawn by filling its glyphs, so its weight cannot change here.",
  boldPart: "This part of the text cannot take a weight of its own here.",
  boldUnavailable: "The weight of this text cannot change here.",
});

export function invalid(message: string): never {
  throw new PdfNativeTransactionError("invalid-request", message);
}

/** A draft's own state in one key order, without its move or scale. */
export function groupDraftState(draft: PdfNativeTextPropertyDraft): string {
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
    draft.syntheticBold ?? null,
    // Rules a Clear Formatting takes away are a change of the page, not a
    // move: a draft carrying them is never read as the move alone.
    draft.dropTextRules?.length ? draft.dropTextRules : null,
    draft.addTextRules?.length ? draft.addTextRules : null,
    draft.textRuleParts?.length ? draft.textRuleParts : null,
    draft.textRuleRanges?.length ? draft.textRuleRanges : null,
    draft.segments?.length ? draft.segments : null,
  ]);
}

















export function readMraiNativeTextGroupMoveDraft(
  draft: PdfNativeTextPropertyDraft | null,
  proof: Readonly<{
    properties: unknown;
    move?: unknown;
    groupMove?: PdfNativeTextGroupMoveProof;
  }> | null,
  propertyBaseline: PdfNativeTextPropertyDraft | null = null,
): PdfNativeTextGroupEdit | null {
  if (!draft || !proof || !proof.groupMove) return null;
  const baseline = proof.groupMove.baseline;
  const bold = !proof.properties && draft.syntheticBold !== undefined &&
    draft.syntheticBold !== (baseline.syntheticBold ?? false)
    ? draft.syntheticBold : undefined;
  if (proof.move && !draft.groupScale && !draft.groupLinear && bold === undefined) return null;
  if (proof.properties) {
    if (!propertyBaseline || groupDraftState(draft) !== groupDraftState(propertyBaseline)) return null;
  } else if (
    groupDraftState({ ...draft, syntheticBold: baseline.syntheticBold }) !== groupDraftState(baseline)
  ) {
    invalid(
      "Only the position, size and weight of this text can change here; it keeps the font, spacing and colour the page draws it in.",
    );
  } else if (bold !== undefined && baseline.syntheticBold === undefined) {
    invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.boldUnavailable);
  }
  const [a, b, c, d] = draft.groupLinear ?? [draft.groupScale?.x ?? 1, 0, 0, draft.groupScale?.y ?? 1];
  const e = draft.moveX ?? 0;
  const f = draft.moveY ?? 0;
  if (![a, b, c, d, e, f].every(Number.isFinite)) invalid(MRAI_NATIVE_TEXT_GROUP_MOVE_REFUSALS.transform);
  return Object.freeze({ a, ...(b !== 0 || c !== 0 ? { b, c } : {}), d, e, f,
    ...(bold !== undefined ? { bold: Object.freeze({ on: bold }) } : {}) });
}

/** Whether a group edit leaves the box as it was: the map leaves it where it
 * was and it sets no weight. */
export function mraiNativeTextGroupTransformIsIdentity(transform: PdfNativeTextGroupEdit): boolean {
  return transform.a === 1 && transform.d === 1 && !transform.b && !transform.c &&
    transform.e === 0 && transform.f === 0 && !transform.bold;
}

/** Whether a group edit turns the box (a map only a revision published as
 * its bytes carries: the server's record replay keeps the page's axes). */
export function mraiNativeTextGroupTransformTurns(transform: PdfNativeTextGroupEdit): boolean {
  return Boolean(transform.b || transform.c);
}
