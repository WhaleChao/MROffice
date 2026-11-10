/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An imported picture's Office original (`a14:imgLayer`): after an Artistic
 * Effect or Remove Background, PowerPoint keeps the RESULT as the picture and
 * the original as an HD Photo part, and a different effect, Reset Picture or
 * the next removal starts from that original.
 *
 * The page names the original only by digest
 * (`source.pictureLayerOriginalSha256`); the bytes stay in the deck's own
 * package and the API decodes them on request
 * (`GET /pages/{id}/presentation-picture-originals/{block}`). The browser
 * never reads OOXML: it asks for a PNG and puts it on the picture.
 */

import { apiFetchRaw } from "@mrnote/core/web/lib/api.ts";
import type { PresentationImageBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";

const DIGEST = /^[0-9a-f]{64}$/;

/** The digest the page names for the picture's layer original, or null. */
export function pictureLayerOriginalDigest(block: PresentationImageBlock | null | undefined): string | null {
  const source = (block as { source?: unknown } | null | undefined)?.source;
  const digest =
    source && typeof source === "object"
      ? (source as Record<string, unknown>).pictureLayerOriginalSha256
      : undefined;
  return typeof digest === "string" && DIGEST.test(digest) ? digest : null;
}

/**
 * Whether a picture whose Artistic Effect the file drew can go back to its
 * original: the package keeps one and no Remove Background result sits on
 * top of it (going back to the original would silently undo the removal).
 */
export function pictureEffectCanReturnToOriginal(block: PresentationImageBlock | null | undefined): boolean {
  if (!block || (block as Record<string, unknown>).artisticEffectRendered !== true) return false;
  if ((block as Record<string, unknown>).backgroundRemoval) return false;
  return pictureLayerOriginalDigest(block) !== null;
}

/**
 * The picture drawn from its original again with `effect` (null: none), as
 * picking another Artistic Effect -- or Reset Picture -- does in PowerPoint.
 * The effect is now MRAI's to draw, so the "drawn by the file" mark goes.
 */
export function withPictureOriginal(
  block: PresentationImageBlock,
  originalSrc: string,
  effect: string | null,
): PresentationImageBlock {
  const next = { ...block, src: originalSrc } as Record<string, unknown>;
  delete next.artisticEffectRendered;
  if (effect) next.artisticEffect = effect;
  else delete next.artisticEffect;
  return next as unknown as PresentationImageBlock;
}

/** The URL the original is read from. */
export function pictureLayerOriginalUrl(pageId: string, blockId: string): string {
  return `/api/v1/pages/${encodeURIComponent(pageId)}/presentation-picture-originals/${encodeURIComponent(blockId)}`;
}

/** The original as a PNG data URL, or null when it cannot be read. */
export async function fetchPictureLayerOriginal(pageId: string, blockId: string): Promise<string | null> {
  try {
    const response = await apiFetchRaw(pictureLayerOriginalUrl(pageId, blockId), {}, {
      suppressUnauthorizedHandling: true,
    });
    if (!response.ok) return null;
    const blob = await response.blob();
    if (blob.type !== "image/png") return null;
    return await new Promise<string | null>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : null);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}
