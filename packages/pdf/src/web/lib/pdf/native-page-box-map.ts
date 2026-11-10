/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeKeptTextBlock } from "./native-kept-text-blocks";
import type { PdfNativeParagraphEditBox } from "./native-paragraph-edit-boxes";
import type { PdfNativeListTextSelectionItemEvidence } from "./engine/native/list-text-fidelity";

/** How many boxes one page's map holds at most: a page with more is read by
 * its look past them (a page of a table set cell by cell). */
export const PDF_NATIVE_PAGE_BOX_MAP_MAX_BOXES = 600;
/** The longest text one pinned box carries. */
const MAX_PINNED_TEXT = 20_000;

type Item = PdfNativeListTextSelectionItemEvidence;

/** A box's text items in the order the page draws them. */
function sourceOrdered(items: readonly Item[]): Item[] {
  return [...items].sort((left, right) => left.itemIndex - right.itemIndex);
}

/** The rows a run of items draws, top to bottom as the page draws them: a
 * row opens after an item the reader ended its line on, or off the row's
 * baseline (mirror of native-kept-text-blocks rowsOfItems). */
function rowsOf(items: readonly Item[]): string[] {
  const rows: string[] = [];
  let baseline: number | null = null;
  let ended = false;
  for (const item of items) {
    const y = Number(item.transform[5]);
    if (baseline === null || ended || Math.abs(y - baseline) > Math.max(Number(item.height) || 0, 1) / 2) {
      rows.push("");
      baseline = y;
    }
    rows[rows.length - 1] += item.str;
    ended = item.hasEOL === true;
  }
  return rows.map((row) => row.trimEnd()).filter((row) => row.trim());
}

/** The pinned block of a box, or null when it cannot name one. */
export function pdfNativePinnedBoxOf(
  pageNumber: number,
  items: readonly Item[],
): PdfNativeKeptTextBlock | null {
  const ordered = sourceOrdered(items);
  const head = ordered[0];
  if (!head || !head.str.trim()) return null;
  const x = Number(head.transform[4]), y = Number(head.transform[5]);
  const rows = rowsOf(ordered);
  if (!Number.isFinite(x) || !Number.isFinite(y) || rows.length === 0 ||
    rows.reduce((total, row) => total + row.length, 0) > MAX_PINNED_TEXT) return null;
  return Object.freeze({ pageNumber, rows: Object.freeze(rows), origin: Object.freeze({ x, y }), pinned: true as const });
}

/**
 * The box a write sets, named by what it writes there: its text, from where
 * its first item stood (a write in place keeps where its text starts). A
 * write that moves where the text starts (a line set centred again) names
 * nothing the page draws and the box is read by its marks and its look.
 */
export function pdfNativePinnedBoxOfWrite(
  pageNumber: number,
  items: readonly Item[],
  text: string,
): PdfNativeKeptTextBlock | null {
  const head = sourceOrdered(items).find((item) => item.str.trim());
  const rows = text.split(/\n/u).map((row) => row.trimEnd()).filter((row) => row.trim());
  if (!head || rows.length === 0 || rows.reduce((total, row) => total + row.length, 0) > MAX_PINNED_TEXT) return null;
  const x = Number(head.transform[4]), y = Number(head.transform[5]);
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  return Object.freeze({ pageNumber, rows: Object.freeze(rows), origin: Object.freeze({ x, y }), pinned: true as const });
}

const boxKey = (box: Pick<PdfNativeParagraphEditBox, "items">) =>
  sourceOrdered(box.items).map((item) => item.itemIndex).join(",");
/** A box as it reads: its items, and its text in the order it shows it. */
const boxReading = (box: Pick<PdfNativeParagraphEditBox, "items" | "text">) => `${boxKey(box)}\u0000${box.text}`;

/**
 * The map of a page as `read` reads it: a pinned block per box that,
 * read with the others, comes back as that box. `read` reads the page with
 * the blocks given (on top of the blocks the page is read with already).
 */
export function pdfNativePageBoxMap(
  pageNumber: number,
  boxes: readonly PdfNativeParagraphEditBox[],
  read: (blocks: readonly PdfNativeKeptTextBlock[]) => readonly PdfNativeParagraphEditBox[],
): readonly PdfNativeKeptTextBlock[] {
  const wanted = new Set(boxes.map(boxReading));
  let pins = boxes.slice(0, PDF_NATIVE_PAGE_BOX_MAP_MAX_BOXES).flatMap((box) => {
    const pin = pdfNativePinnedBoxOf(pageNumber, box.items);
    return pin ? [{ pin, key: boxReading(box) }] : [];
  });
  // A pin that does not read back as its box (its items, its text in its
  // order) is dropped, and the rest read again without it: what one claimed
  // wrongly may have kept another from its own items.
  for (let round = 0; round < 4 && pins.length; round += 1) {
    const read_ = new Set(read(pins.map((entry) => entry.pin)).map(boxReading));
    const kept = pins.filter((entry) => read_.has(entry.key) && wanted.has(entry.key));
    if (kept.length === pins.length) break;
    pins = kept;
  }
  return Object.freeze(pins.map((entry) => entry.pin));
}
