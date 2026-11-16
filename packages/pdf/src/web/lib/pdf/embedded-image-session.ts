/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The state behind picking an embedded image and saying what to do with it.
 *
 * The edit model (`embedded-image-edits`) and the applier
 * (`embedded-image-apply`) are both pure functions over a list of edits. What
 * sits between them is a small amount of bookkeeping that is easy to get wrong
 * in a component: one edit per image rather than a stack of masks, replacement
 * bytes that stay reachable while an edit refers to them and are released the
 * moment it does not, and an order that does not depend on the order the user
 * happened to click in.
 *
 * Keeping it here means those rules are testable without a browser.
 */

import type { EmbeddedImageEdit, FractionRect } from "./embedded-image-edits";
import type { ReplacementBytes } from "./embedded-image-apply";
import type { PlacedImage } from "./embedded-images";

/** Edits keyed by page and paint order, so a second click on the same image
 *  replaces its edit instead of stacking another mask on top. */
export type EmbeddedImageEdits = Readonly<Record<string, EmbeddedImageEdit>>;

/** Bytes the user chose locally. They are only uploaded if the file is saved
 *  back to the notebook; a plain download never needs them to leave the tab. */
export interface LocalReplacement {
  id: string;
  bytes: Uint8Array;
  mimeType: string;
  /** The file's own name, for the status line. */
  name: string;
}

export function imageEditKey(page: number, order: number): string {
  return `${page}:${order}`;
}

export function imageEditAt(
  edits: EmbeddedImageEdits,
  page: number,
  order: number,
): EmbeddedImageEdit | undefined {
  return edits[imageEditKey(page, order)];
}

export function withImageEdit(
  edits: EmbeddedImageEdits,
  edit: EmbeddedImageEdit,
): EmbeddedImageEdits {
  return { ...edits, [imageEditKey(edit.page, edit.order)]: edit };
}

export function withoutImageEdit(
  edits: EmbeddedImageEdits,
  page: number,
  order: number,
): EmbeddedImageEdits {
  const key = imageEditKey(page, order);
  if (!(key in edits)) return edits;
  const next = { ...edits };
  delete next[key];
  return next;
}

/**
 * The staged edits once a write refused one of them: only that edit goes.
 *
 * The others were not part of the refused write and stay staged, and an edit
 * staged for the same picture since is a newer one, not the refused one.
 */
export function withoutRefusedImageEdit(
  edits: EmbeddedImageEdits,
  refused: EmbeddedImageEdit,
): EmbeddedImageEdits {
  if (imageEditAt(edits, refused.page, refused.order) !== refused) return edits;
  return withoutImageEdit(edits, refused.page, refused.order);
}

/**
 * Edits in page then paint order.
 *
 * The applier draws in the order it is given, and two edits on one page can
 * overlap — a mask for one image can land on top of a replacement for another.
 * Paint order is the document's own answer to which is on top, so following it
 * makes the result independent of the order the user clicked.
 */
export function sortedImageEdits(
  edits: EmbeddedImageEdits,
): EmbeddedImageEdit[] {
  return Object.values(edits)
    .slice()
    .sort((a, b) => a.page - b.page || a.order - b.order);
}

/**
 * Replacement ids that no edit refers to any more.
 *
 * Replacing the same image twice, or undoing a replace, strands the earlier
 * bytes: they stay in the map, and any object URL made from them stays alive
 * for the life of the tab. The caller revokes what this returns.
 */
export function orphanReplacementIds(
  edits: EmbeddedImageEdits,
  known: Iterable<string>,
): string[] {
  const live = new Set(
    Object.values(edits)
      .map((edit) => edit.dataItemId)
      .filter((id): id is string => !!id),
  );
  const orphans: string[] = [];
  for (const id of known) if (!live.has(id)) orphans.push(id);
  return orphans;
}

/** The applier's input map, from the local replacements the user picked. */
export function replacementBytesFor(
  replacements: Iterable<LocalReplacement>,
): ReplacementBytes {
  const map = new Map<string, { bytes: Uint8Array; mimeType: string }>();
  for (const item of replacements) {
    map.set(item.id, { bytes: item.bytes, mimeType: item.mimeType });
  }
  return map;
}

export interface ImageEditSummary {
  deleted: number;
  replaced: number;
  moved: number;
  total: number;
}

export function summariseImageEdits(
  edits: EmbeddedImageEdits,
): ImageEditSummary {
  const summary: ImageEditSummary = {
    deleted: 0,
    replaced: 0,
    moved: 0,
    total: 0,
  };
  for (const edit of Object.values(edits)) {
    summary.total += 1;
    if (edit.kind === "delete") summary.deleted += 1;
    else if (edit.kind === "replace") summary.replaced += 1;
    else if (edit.kind === "move") summary.moved += 1;
  }
  return summary;
}

/**
 * Which image a click at `point` (page fractions) is on.
 *
 * Images overlap — a figure drawn over a background panel is two placements at
 * the same spot — and the one a reader would say they clicked is the one on
 * top. That is the LAST in paint order, and the smallest is preferred at equal
 * depth so a full-page background never swallows a click meant for the figure
 * on it.
 */
export function imageAtPoint(
  images: ReadonlyArray<PlacedImage>,
  point: { x: number; y: number },
): PlacedImage | undefined {
  // Every comparison below is false for NaN, so an unusable point would fail
  // to be OUTSIDE any image and select whichever came first.
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return undefined;
  let best: PlacedImage | undefined;
  let bestArea = Infinity;
  for (const image of images) {
    if (
      point.x < image.x ||
      point.y < image.y ||
      point.x > image.x + image.width ||
      point.y > image.y + image.height
    ) {
      continue;
    }
    const area = image.width * image.height;
    if (!best || area < bestArea || (area === bestArea && image.order > best.order)) {
      best = image;
      bestArea = area;
    }
  }
  return best;
}

/** The placement an edit currently shows at — where a move put it, or home. */
export function currentRectFor(
  image: PlacedImage,
  edit: EmbeddedImageEdit | undefined,
): FractionRect {
  if (edit?.rect) return edit.rect;
  return {
    x: image.x,
    y: image.y,
    width: image.width,
    height: image.height,
  };
}
