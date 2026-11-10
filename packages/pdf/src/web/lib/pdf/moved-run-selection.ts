/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { summarizePdfNativeSelectionItems } from "./engine/native/list-text-model";
import type { PdfNativeListTextSelectionItemEvidence } from "./engine/native/list-text-fidelity";
import type { PdfNativeParagraphEditBox } from "./native-paragraph-edit-boxes";

/**
 * The selection of a run -- one an edit just moved along its line, or one
 * standing where it was -- that the reader does not report as one item with
 * the run's own text.
 *
 * Character or word spacing makes pdf.js split a single show into pieces and
 * put spaces between its glyphs ("L a t i n", " ", "d i s p l a y", ...), so
 * no item's text equals the run's. The page's edit boxes group those pieces
 * back into the line the reader opens on a click; the box whose text is the
 * run's, spaces aside, on the nearest baseline, is that run. The selection is
 * built exactly as a click on the box builds it, and the planner still has to
 * prove one source operator for it.
 */
export type MovedRunSelection = Readonly<{
  item: PdfNativeListTextSelectionItemEvidence;
  items?: readonly PdfNativeListTextSelectionItemEvidence[];
  endOffset: number;
}>;

const withoutSpaces = (text: string) => text.replace(/\s+/gu, "");

/** How far a size, rise or spacing change may carry a run's baseline (pt). */
const MOVED_RUN_BASELINE_REACH = 30;

export function movedRunSelection(
  boxes: readonly PdfNativeParagraphEditBox[],
  text: string,
  baselineY: number,
  place: Readonly<{
    /** A run that has not moved: it starts exactly here, on that baseline. */
    exactX?: number;
    /** A run turned about its box centre: its origin left the baseline, so
     * the one box on the page with its text is taken, and only that. */
    anywhere?: boolean;
  }> = {},
): MovedRunSelection | null {
  const wanted = withoutSpaces(text);
  if (!wanted) return null;
  const near = boxes
    .flatMap((box) => {
      const first = box.items[0];
      if (!box.editable || !first) return [];
      const summary = box.selectionItem?.str ?? summarizePdfNativeSelectionItems(box.items);
      if (withoutSpaces(summary) !== wanted) return [];
      const dy = Math.abs(first.transform[5] - baselineY);
      if (place.exactX !== undefined) {
        return dy < 0.01 && Math.abs(first.transform[4] - place.exactX) < 0.01 ? [{ box, summary, dy }] : [];
      }
      return place.anywhere || dy < MOVED_RUN_BASELINE_REACH ? [{ box, summary, dy }] : [];
    })
    .sort((left, right) => left.dy - right.dy);
  if (place.anywhere && near.length !== 1) return null;
  // The nearest baseline, and no other one as near: two lines with the same
  // text an equal distance away are an ambiguous target.
  if (!near.length || (near.length > 1 && near[1].dy - near[0].dy <= 1)) return null;
  const { box, summary } = near[0];
  const item = box.selectionItem ?? Object.freeze({
    ...box.items[0],
    str: summary,
    width: box.items.reduce((total, entry) => total + entry.width, 0),
    hasEOL: box.items.at(-1)?.hasEOL ?? false,
  });
  return Object.freeze({
    item,
    ...(box.items.length > 1 && !box.selectionItem ? { items: Object.freeze([...box.items]) } : {}),
    endOffset: summary.length,
  });
}

/**
 * The run whose text is `text` as the leading items of the box on its line:
 * a paragraph set again, read with the text after it as one box until the
 * write's own kept block parts them (the candidate before it is committed:
 * the eiffel.pdf paragraph whose last line the reader joined to the
 * paragraph's next box). Exact as movedRunSelection, its items cut where the
 * text ends, at an item's end.
 */
export function leadingRunSelection(
  boxes: readonly PdfNativeParagraphEditBox[],
  text: string,
  baselineY: number,
): MovedRunSelection | null {
  const wanted = withoutSpaces(text);
  if (!wanted) return null;
  const found: MovedRunSelection[] = [];
  for (const box of boxes) {
    const first = box.items[0];
    if (!box.editable || !first || box.selectionItem || box.items.length < 2 ||
      Math.abs(first.transform[5] - baselineY) >= 0.5) continue;
    let read = "";
    for (let count = 1; count <= box.items.length; count += 1) {
      read += withoutSpaces(box.items[count - 1].str);
      if (!wanted.startsWith(read)) break;
      if (read.length < wanted.length) continue;
      if (count === box.items.length) break;
      const items = box.items.slice(0, count);
      const summary = summarizePdfNativeSelectionItems(items);
      if (withoutSpaces(summary) !== wanted) break;
      found.push(Object.freeze({
        item: Object.freeze({ ...items[0], str: summary, width: items.reduce((total, entry) => total + entry.width, 0),
          hasEOL: items.at(-1)?.hasEOL ?? false }),
        items: Object.freeze([...items]),
        endOffset: summary.length,
      }));
      break;
    }
  }
  return found.length === 1 ? found[0] : null;
}
