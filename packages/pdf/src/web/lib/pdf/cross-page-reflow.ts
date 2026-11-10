/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The cross-page half of reflow (R3's remainder).
 *
 * The per-page overlay honestly sees one page; the seam belongs to the
 * continuous editor. The logic lives here as pure functions over the
 * editor's own data — edits, per-page box targets — so the part that decides
 * "offer it" and the part that builds the edits are testable without a
 * browser, and the component keeps only the rendering.
 */

import { chainForBox, inferBoxChains, type BoxChain } from "./box-chain";
import type { PdfTextStyleRun } from "@mrnote/core/web/lib/pdf/text-style-runs.ts";
import type { PdfTextLineLayout } from "@mrnote/core/web/lib/pdf/text-line-layout.ts";
import {
  reflowChainStyled,
  type StyledReflowBox,
} from "./styled-reflow";

/** The slice of a text-box command target this module reads. */
export interface CrossPageTarget {
  id: string;
  /** 0-based page index, as the editor stores them. */
  page: number;
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize?: number;
  fontFamily?: string;
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
}

/** The slice of a persisted edit this module reads and writes. */
export interface CrossPageEdit {
  id: string;
  page: number;
  /** Present on real edit records; cleared when reflow moves the text. */
  runs?: PdfTextStyleRun[];
  lineLayout?: PdfTextLineLayout[];
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  fontSize?: number;
  fontFamily?: string;
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
  originX?: number;
  originY?: number;
  originWidth?: number;
  originHeight?: number;
}

export interface CrossPageReflowOffer {
  editId: string;
  /** 0-based index of the page BEFORE the seam. */
  pageIndex: number;
  chain: BoxChain;
}

const DEFAULT_FONT_SIZE = 0.02;

/**
 * The first edit whose box chains onto the next page.
 *
 * Only an edit whose text actually differs from the extracted box offers —
 * moving text around because a box was merely selected would be reflow
 * without an edit to justify it.
 */
export function findCrossPageReflowOffer(
  edits: ReadonlyArray<CrossPageEdit>,
  targetsByPage: Readonly<Record<number, ReadonlyArray<CrossPageTarget>>>,
  dismissedEditIds: ReadonlySet<string>,
): CrossPageReflowOffer | null {
  for (const edit of edits) {
    if (dismissedEditIds.has(edit.id)) continue;
    const here = targetsByPage[edit.page];
    const there = targetsByPage[edit.page + 1];
    if (!here || !there) continue;
    const target = here.find((candidate) => candidate.id === edit.id);
    if (!target || (edit.text ?? "") === (target.text ?? "")) continue;
    const chains = inferBoxChains(
      [...here, ...there].map((box) => ({
        id: box.id,
        page: box.page + 1,
        x: box.x,
        y: box.y,
        width: box.width,
        height: box.height,
        fontSize: box.fontSize ?? DEFAULT_FONT_SIZE,
        text: box.text || "",
      })),
    );
    const chain = chainForBox(chains, edit.id);
    if (!chain) continue;
    if (!chain.links.some((link) => link.kind === "cross-page")) continue;
    const at = chain.boxIds.indexOf(edit.id);
    if (at < 0 || at >= chain.boxIds.length - 1) continue;
    return { editId: edit.id, pageIndex: edit.page, chain };
  }
  return null;
}

/** Character-width estimate in page fractions, adequate at page scale. */
export function estimateReflowWidth(text: string, fontSize: number): number {
  return text.length * fontSize * 0.55;
}

/**
 * The edit records an accepted offer produces.
 *
 * Merging is replace-by-id over the CURRENT edits — an edit the user already
 * made on a downstream box keeps its styling and origin, only its text moves.
 * Returns null when the offer no longer matches the data it was made on.
 */
export function buildCrossPageReflowEdits(
  offer: CrossPageReflowOffer,
  edits: ReadonlyArray<CrossPageEdit>,
  targetsByPage: Readonly<Record<number, ReadonlyArray<CrossPageTarget>>>,
  measure: (text: string, fontSize: number) => number = estimateReflowWidth,
): CrossPageEdit[] | null {
  const editsById = new Map(edits.map((edit) => [edit.id, edit]));
  const targetsById = new Map(
    [
      ...(targetsByPage[offer.pageIndex] ?? []),
      ...(targetsByPage[offer.pageIndex + 1] ?? []),
    ].map((box) => [box.id, box]),
  );
  const chainBoxes: StyledReflowBox[] = offer.chain.boxIds
    .map((id) => targetsById.get(id))
    .filter((box): box is CrossPageTarget => !!box)
    .map((box) => {
      const boxEdit = editsById.get(box.id);
      return {
        id: box.id,
        width: box.width,
        height: box.height,
        fontSize: boxEdit?.fontSize ?? box.fontSize ?? DEFAULT_FONT_SIZE,
        text: (boxEdit?.text ?? box.text) || "",
        runs: boxEdit?.runs,
      };
    });
  const edited = editsById.get(offer.editId);
  if (!edited || chainBoxes.length < 2) return null;

  const result = reflowChainStyled(
    chainBoxes,
    offer.editId,
    edited.text ?? "",
    edited.runs,
    measure,
  );
  if (!result) return null;

  const out: CrossPageEdit[] = [];
  for (const boxId of result.changedBoxIds) {
    const target = targetsById.get(boxId);
    const output = result.outputs.get(boxId);
    if (!target || output === undefined) continue;
    const nextText = output.text;
    const existing = editsById.get(boxId);
    if ((existing?.text ?? target.text) === nextText) continue;
    out.push({
      ...existing,
      id: boxId,
      page: target.page,
      x: existing?.x ?? target.x,
      y: existing?.y ?? target.y,
      width: existing?.width ?? target.width,
      height: existing?.height ?? target.height,
      fontSize: existing?.fontSize ?? target.fontSize,
      fontFamily: existing?.fontFamily ?? target.fontFamily,
      sourceFontName: existing?.sourceFontName ?? target.sourceFontName,
      bold: existing?.bold ?? target.bold,
      italic: existing?.italic ?? target.italic,
      // Styling rides with the characters; only the layout is recomputed.
      runs: output.runs,
      lineLayout: undefined,
      text: nextText,
      // The mask pins to the box's own extracted rect.
      originX: existing?.originX ?? target.x,
      originY: existing?.originY ?? target.y,
      originWidth: existing?.originWidth ?? target.width,
      originHeight: existing?.originHeight ?? target.height,
    });
  }
  return out;
}
