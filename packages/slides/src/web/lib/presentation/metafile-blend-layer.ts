/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An EMF picture's blend layer (readings row 1021, the API's
 * `source.pictureMetafileBlend` {mode, layer, over}).
 *
 * Windows PowerPoint applies a metafile's lone SRCPAINT / SRCAND bitmap to
 * the slide under the picture: the bitmap is ORed / ANDed into what is
 * already drawn there (its slide show and Slide.Export). The import renders
 * the picture without those pixels and renders them apart as the layer,
 * which the slide lays over the picture in CSS `mix-blend-mode` lighten /
 * darken -- OR / AND exactly for 0 / 255 channels. Display only: the
 * package keeps the metafile.
 *
 * A blend sees only what its stacking context holds, and every object on
 * the slide is one (`blockGeometryStyle`'s `zIndex: 0`): a picture with a
 * layer is no stacking context of its own (`blendsWithSlide`), so the layer
 * blends with the slide and the objects under it. The import gives a layer
 * only to a picture neither turned nor mirrored (a transform is a stacking
 * context); one turned since draws without its layer.
 */
import { pictureObjectDisplaySrc, pictureObjectTokenForUrl } from "@mrnote/core/web/lib/page-picture-objects.ts";

export type MetafileBlendMode = "lighten" | "darken";

export interface MetafileBlendLayer {
  readonly mode: MetafileBlendMode;
  /** The layer's stored picture, `mrai-picture:<sha256>` (in a print, the print's name for it). */
  readonly token: string;
  /** What an `<img>` draws it from. */
  readonly src: string;
}

type Rec = Record<string, unknown>;

function asRec(value: unknown): Rec | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Rec) : null;
}

/**
 * The block's blend layer, or null: none, or the picture is no longer the one
 * it was rendered for (`over`: a picture replaced since keeps its source).
 */
export function metafileBlendLayer(block: unknown): MetafileBlendLayer | null {
  const record = asRec(block);
  const blend = asRec(asRec(record?.source)?.pictureMetafileBlend);
  if (!record || !blend) return null;
  const mode = blend.mode;
  if (mode !== "lighten" && mode !== "darken") return null;
  // The page names it by its token; a deck's print hands the print page every
  // picture -- this one, and the `over` beside the picture's own src -- under
  // the print's own names (`inline_presentation_media`, `with_print_pictures`:
  // a data URL, or a path of the print page's origin).
  const token = blend.layer;
  if (typeof token !== "string" || !/^(mrai-picture:|data:image\/|\/)/.test(token)) return null;
  if (typeof blend.over === "string") {
    const src = typeof record.src === "string" ? pictureObjectTokenForUrl(record.src) : null;
    if (src !== pictureObjectTokenForUrl(blend.over)) return null;
  }
  return { mode, token, src: pictureObjectDisplaySrc(token) };
}

/**
 * Whether a block's frame must not be a stacking context of its own: a picture
 * whose blend layer is drawn, on a frame neither turned nor mirrored (that
 * frame's transform is a stacking context the layer cannot see out of).
 */
export function blendsWithSlide(block: unknown, transform: string | undefined): boolean {
  return transform === undefined && asRec(block)?.type === "image" && metafileBlendLayer(block) !== null;
}
