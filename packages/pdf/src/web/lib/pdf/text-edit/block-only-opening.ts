/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// A paragraph that opens as a block for its formats alone.
//

// edit carries, and holding what no text box may stand for -- is refused
// whole, and the press that found it so opens the line under it: its words

// text stands still reads such a paragraph whole. So the window, once the
// paragraph was refused, asks the engine whether it opens as a block that
// takes formats and no words (the preparation request's `formatsWhole`;
// native-text-edit-preparation.worker's proofOfParagraphAsFormatsBlock) and,
// where it does, offers the whole paragraph from the editor of each of its
// lines (PdfNativeParagraphBlockOffer). Taken, the line is let go and the
// paragraph opens selected as a block, read-only, where 加粗, 倾斜, 下划线,
// 删除线 and a colour are set on all of it in one write
// (PdfNativeListTextSelectionProof.wordsFixed, "too-large-to-set-again").
// A double press on the block lets it go and opens the line under the press
// for its words.
//
// Nothing a press did before is done otherwise: the block is entered from a
// line of it only, by name.
import type { PdfNativeListTextSelectionEvidence, PdfNativeListTextSelectionProof } from "../engine/native/list-text-fidelity";
import { pdfNativeRefusedTextOpening } from "./refused-opening";

/** The engine-side sentence for a box asked for as such a block that is none
 * (the window offers the block only where the engine has answered with it). */
export const PDF_NATIVE_TEXT_NOT_A_FORMATS_BLOCK_MESSAGE =
  "This text does not open as a block whose formats are written where it stands.";

/** Whether the box opens as a block that takes formats and no words. */
export function pdfNativeTextOpensAsBlockOnly(
  proof: Pick<PdfNativeListTextSelectionProof, "wordsFixed"> | null | undefined,
): boolean {
  return proof?.wordsFixed === "too-large-to-set-again";
}

type Selection = Pick<PdfNativeListTextSelectionEvidence, "pageNumber" | "sourceSha256" | "item" | "items">;

/**
 * A paragraph refused whole that the engine opens as a block for its formats:
 * what the window keeps of it, for the bytes it was read in. `opening` is
 * what the paragraph's own press brought (its selection, the frame the reader
 * gave it), opaque here.
 */
export type PdfNativeParagraphBlockOffer<Opening = unknown> = Readonly<{
  pageNumber: number;
  sourceSha256: string;
  /** Where its runs stand on the page (PDF units), a line's height around them. */
  bounds: Readonly<{ left: number; right: number; bottom: number; top: number }>;
  /** How many reader items it holds. */
  items: number;
  opening: Opening;
}>;

/** The offer of `selection` (a paragraph of several reader items, refused
 * whole); null for one item, and where its runs state no place. */
export function pdfNativeParagraphBlockOffer<Opening>(selection: Selection, opening: Opening): PdfNativeParagraphBlockOffer<Opening> | null {
  const items = selection.items?.length ?? 0;
  if (items < 2) return null;
  const placed = pdfNativeRefusedTextOpening("", selection, 0);
  return placed ? Object.freeze({ pageNumber: selection.pageNumber, sourceSha256: selection.sourceSha256,
    bounds: placed.bounds, items, opening }) : null;
}

/**
 * Whether `line` -- the box open now -- is a part of the paragraph offered:
 * the same page of the same bytes, every run of it standing where the
 * paragraph's runs stand, and fewer runs than the paragraph (the paragraph
 * itself is no line of it). A page written since is other bytes: the offer
 * is void, and made again when the paragraph is pressed and refused again.
 */
export function pdfNativeParagraphBlockOfferedFor(
  offer: PdfNativeParagraphBlockOffer | null | undefined,
  line: Selection | null | undefined,
): boolean {
  if (!offer || !line || offer.pageNumber !== line.pageNumber || offer.sourceSha256 !== line.sourceSha256) return false;
  const runs = line.items?.length ? line.items : [line.item];
  if (runs.length >= offer.items) return false;
  return runs.every((item) => {
    const [, , , , x, y] = item.transform;
    return x >= offer.bounds.left && x <= offer.bounds.right && y >= offer.bounds.bottom && y <= offer.bounds.top;
  });
}

/** How long after the reader asked for the whole paragraph the window's own
 * click on the line may still be answered with the paragraph (the line is
 * let go first, which may wait for the page). */
export const PDF_NATIVE_PARAGRAPH_BLOCK_ASK_MS = 15_000;
