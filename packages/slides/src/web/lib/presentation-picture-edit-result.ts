/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { BackgroundRemovalSpec } from "@mrnote/core/web/lib/office-picture-background-removal.ts";
import { isPictureObjectUrl } from "@mrnote/core/web/lib/page-picture-objects.ts";

/**
 * Asset custody on a picture states that its inline `src` is a cache of the
 * authenticated uploaded file (the save validator requires the cache MIME to
 * equal `mimeType`; the export hashes the cache against the stored object).
 * Remove Background and Erase draw a new PNG, so the result no longer is that
 * file: keeping the custody refused the save of a JPEG ("cache MIME must match
 * mimeType") and the export of a PNG ("no longer matches its authorized
 * source"). The result is an inline picture; the original stays on the page as
 * `backgroundRemovalOriginalSrc`, which the export writes as the HD Photo
 * layer. The intrinsic pixel pair stays: the edit keeps the aspect ratio that
 * Crop Fill / Fit reads.
 */
const PICTURE_ASSET_CUSTODY_KEYS = [
  "sourceKind",
  "assetId",
  "mimeType",
  "sizeBytes",
  "hasAlpha",
  "exifOrientation",
] as const;

type PictureEditTarget = Record<string, unknown> & {
  src?: unknown;
  source?: unknown;
  backgroundRemoval?: unknown;
  backgroundRemovalOriginalSrc?: unknown;
};

export type PresentationPictureEdit =
  | {
      readonly mode: "removal";
      readonly src: string;
      readonly spec: BackgroundRemovalSpec;
      /** The package original fetched for this session, if any. */
      readonly originalSrc?: string | null;
    }
  | { readonly mode: "erase"; readonly src: string };

/** The picture block after Keep Changes of Remove Background or Erase. */
export function presentationPictureEditResult<T extends object>(
  block: T,
  edit: PresentationPictureEdit,
): T {
  const target = block as unknown as PictureEditTarget;
  const next: PictureEditTarget = { ...target, src: edit.src };
  for (const key of PICTURE_ASSET_CUSTODY_KEYS) delete next[key];
  if (edit.mode === "removal") {
    next.backgroundRemoval = edit.spec;
    const source = target.source as Record<string, unknown> | undefined;
    const imported = typeof source?.shapeId === "string";
    // A picture stored beside the deck is drawn from its URL: no bytes the
    // page can keep as the original (the stored picture stays where it is).
    const stored = isPictureObjectUrl(target.src);
    if (!target.backgroundRemoval && !target.backgroundRemovalOriginalSrc && !imported && !stored) {
      // The first removal on a picture this deck owns keeps its original on
      // the page, so the next Remove Background starts from it.
      next.backgroundRemovalOriginalSrc = target.src;
    } else if (edit.originalSrc && !target.backgroundRemovalOriginalSrc) {
      // The package original fetched for this session: a deck written from
      // scratch keeps it too (as the layer's HD Photo original).
      next.backgroundRemovalOriginalSrc = edit.originalSrc;
    }
  } else {
    // Erase makes a new picture: the removal session no longer describes it.
    delete next.backgroundRemoval;
    delete next.backgroundRemovalOriginalSrc;
  }
  return next as unknown as T;
}
