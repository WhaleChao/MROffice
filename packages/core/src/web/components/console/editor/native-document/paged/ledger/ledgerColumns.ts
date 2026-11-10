/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Unequal-width newspaper columns inside one sheet.
 *
 * CSS multicol only knows equal tracks, so `unequalColumnProjection` places
 * every direct child of such a sheet into an exact track itself and reports
 * the blocks that straddle a track boundary as split candidates. Word cuts
 * those blocks at the boundary; this step materialises that cut as an
 * IN-SHEET continuation seam (both fragments stay in the same pageBlock, the
 * projector paints the tail in the next track) with the same lossless split
 * helpers a page edge uses. The seam is structural: the ledger reads it as a
 * column seam and never merges it while the two fragments sit in different
 * tracks.
 *
 * Planning happens at the SOURCE track width: the candidate is currently
 * painted whole in the following track, so the projector temporarily
 * measures it at the width it is being cut for.
 */
import type { Transaction } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import { normalizeWordCompatibilityMode } from "../../documentOverflowPunctuation";
import {
  NOTE_LIST_NODE,
  isSplittableBlock,
  isTextSplittableBlock,
  measureCodeLineMetrics,
  minimumTextSplitLines,
  planHeadLines,
  planNoteHeadCursor,
  planTextCutByChars,
} from "../blockSplitting";
import { applyContainerColumnSplit, isSplittableContainer, planContainerHead } from "../containerSplitting";
import {
  applyCodeBlockColumnSplit,
  applyNoteListColumnSplit,
  applyTableColumnSplit,
  applyTextBlockColumnSplit,
  collectPages,
  planRenderedTableFragment,
  REBALANCE_META,
} from "../pageRebalanceExtension";
import { measureUnequalColumnBlockAtWidth, unequalColumnProjectionForSheet } from "../unequalColumnProjection";
import { sheetBlockElements } from "./ledgerAtoms";
import { measureOutsideEditorObservation } from "../../documentLayoutStatus";
import type { LedgerMeasureCache } from "./ledgerMeasureCache";
import { MRNOTE_BLOCK_ID_ATTR } from "../../blockIds";
import { sheetsOf } from "../sheetMeasureMemo";

/**
 * Column cuts the ledger took back. The ledger joins a seam between two
 * fragments on one sheet unless the projector placed the tail in a later
 * track. A cut the projector then placed with both pieces in one track
 * (R13-D13: a paragraph cut for the narrow first column, both pieces drawn
 * in the second) was joined by the next pass and made again by the one
 * after, for as long as the passes lasted: the background never got past
 * it. A block the ledger joined so is not cut on that sheet again until the
 * layout's inputs change (an edit, fonts, the width): the ledger's own
 * composition stands.
 */
const refusedColumnCuts = new WeakMap<EditorView, Set<string>>();

export function refuseColumnCuts(view: EditorView, keys: readonly string[]): void {
  if (keys.length === 0) return;
  let refused = refusedColumnCuts.get(view);
  if (!refused) refusedColumnCuts.set(view, (refused = new Set()));
  for (const key of keys) refused.add(key);
}

export function clearRefusedColumnCuts(view: EditorView): void {
  refusedColumnCuts.delete(view);
}

/**
 * Cut the first block that straddles an unequal-track boundary. Returns the
 * transaction to dispatch, or null when no candidate yields a legal cut.
 */
export function planUnequalColumnSplit(view: EditorView, cache?: LedgerMeasureCache | null): Transaction | null {
  const pages = collectPages(view.state.doc);
  if (!pages) return null;
  // The measurement width is a property of the sheet, set and put back:
  // ProseMirror must not read either write as a change of the page, nor the
  // measurement cache as a reason to measure the sheet again.
  const writeUnobserved = (apply: () => void) =>
    measureOutsideEditorObservation(view, () => (cache ? cache.unrecorded(apply) : apply()));
  const sheets = sheetsOf(view.dom);
  if (sheets.length !== pages.length) return null;
  const refused = refusedColumnCuts.get(view);
  for (let sheetIndex = 0; sheetIndex < pages.length; sheetIndex += 1) {
    const sheet = sheets[sheetIndex];
    const projection = unequalColumnProjectionForSheet(sheet);
    if (!projection) continue;
    const page = pages[sheetIndex];
    for (const candidate of projection.splitCandidates) {
      if (candidate.blockIndex < 0 || candidate.blockIndex >= page.node.childCount) continue;
      const block = page.node.child(candidate.blockIndex);
      const blockId = block.attrs?.[MRNOTE_BLOCK_ID_ATTR];
      if (refused && typeof blockId === "string" && refused.has(`${blockId}@${sheetIndex}`)) continue;
      const dom = sheetBlockElements(sheet)[candidate.blockIndex];
      if (!(dom instanceof HTMLElement)) continue;
      const isCode = isSplittableBlock(block);
      const isText = isTextSplittableBlock(block);
      const isNotes = block.type.name === NOTE_LIST_NODE;
      const isContainer = isSplittableContainer(block);
      const isTable = block.type.name === "table";
      if (!isCode && !isText && !isNotes && !isContainer && !isTable) continue;
      let blockPos = page.pos + 1;
      for (let index = 0; index < candidate.blockIndex; index += 1) blockPos += page.node.child(index).nodeSize;
      const tr = view.state.tr;
      const changed = measureUnequalColumnBlockAtWidth(dom, candidate.measurementWidthPx, () => {
        if (isCode) {
          const metrics = measureCodeLineMetrics(dom, block, dom.offsetHeight);
          const headLines = planHeadLines(candidate.availablePx, metrics);
          return headLines !== null && applyCodeBlockColumnSplit(tr, sheetIndex, candidate.blockIndex, headLines);
        }
        if (isText) {
          // The same head the ledger's column-aware cut plans (ledgerAtoms'
          // text oracle): the fragment rule for the last row, Word's
          // page-bottom credit at a foot on the bottom margin, and no
          // widow/orphan floor for a block opening its column.
          const minLines = candidate.opensColumn ? 1 : minimumTextSplitLines(block.attrs?.widowControl);
          const cut = planTextCutByChars(dom, candidate.availablePx, minLines, true, null, {
            footCreditPx: candidate.footCreditPx ?? 0,
          });
          if (!cut) return false;
          let cutPos: number;
          try {
            cutPos = view.posAtDOM(cut.node, cut.offset);
          } catch {
            return false;
          }
          const from = blockPos + 1;
          const to = blockPos + block.nodeSize - 1;
          if (cutPos <= from || cutPos >= to) return false;
          return applyTextBlockColumnSplit(tr, sheetIndex, candidate.blockIndex, cutPos);
        }
        if (isNotes) {
          const cut = planNoteHeadCursor(dom, candidate.availablePx);
          return cut !== null && applyNoteListColumnSplit(tr, sheetIndex, candidate.blockIndex, cut);
        }
        if (isContainer) {
          const plan = planContainerHead(dom, block, candidate.availablePx, {
            containerPos: blockPos,
            posAtDOM: (node, offset) => {
              try {
                return view.posAtDOM(node, offset);
              } catch {
                return null;
              }
            },
          });
          return plan !== null && applyContainerColumnSplit(tr, sheetIndex, candidate.blockIndex, plan);
        }
        const mode = normalizeWordCompatibilityMode(view.state.doc.attrs.wordCompatibilityMode);
        const planned = planRenderedTableFragment(dom, block, candidate.availablePx, projection.capacityPx, 0, {
          overPageCantSplitStands: mode !== null && mode < 15,
          cellWidowControl: mode !== null && mode >= 15,
        });
        return planned !== null && applyTableColumnSplit(tr, sheetIndex, candidate.blockIndex, planned.plan, planned.cuts);
      }, writeUnobserved);
      if (!changed || !tr.docChanged) continue;
      tr.setMeta(REBALANCE_META, true);
      tr.setMeta("addToHistory", false);
      return tr;
    }
  }
  return null;
}
