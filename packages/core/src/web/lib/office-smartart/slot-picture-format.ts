/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format on a SmartArt picture slot: Crop (Crop > Fill / Fit), the
 * Picture Styles gallery and Reset Picture, written once for the three
 * editors.
 *
 * Office: a picture placeholder holding a picture gets the Picture Format tab
 * (Microsoft Support, "Crop a picture to fit in a shape" -- Crop > Fill / Fit;
 * "Add a border or another style to a picture" -- Picture Styles). Office 2010
 * enabled cropping for pictures in SmartArt graphics, with Fill and Fit, the
 * shape staying where the layout put it ("Office 2010 Improvements for
 * Cropping Pictures", Microsoft Learn archive). In the file the slot's picture
 * is the presenting point's FILL (`dgm:spPr/a:blipFill`, measured on Word's
 * own WC014 / R43-D20 data parts), so:
 *
 * - Crop is that fill's source rectangle, `a:srcRect` (ECMA-376 §20.1.8.55),
 *   in the editors' shared `CropRect` (`lib/presentation-crop-group`: the kept
 *   part of the source as fractions of it, negative / > 1 for Fit's outset).
 * - A Picture Style is the recipe Office writes into the picture's `spPr`
 *   (`office_picture_styles.py`): geometry, line, effects and 3-D, on the same
 *   point as the fill.
 * - Reset Picture removes both.
 *
 * Fields: a worksheet object and a Docs group member state them as
 * {@link SMARTART_PICTURE_CROP_KEY} / {@link SMARTART_PICTURE_STYLE_KEY}; a
 * slide's slot IS a picture block, which keeps its own `cropRect` and
 * `pictureStylePreset` and draws them as any picture does. The engine's image
 * node is that slide block, so it carries the slide's spelling and each
 * projection renames.
 */

import {
  normalizeSignedCropRect,
  pictureFillCropRect,
  pictureFitCropRect,
  type CropRect,
} from "../presentation-crop-group";
import {
  isOfficePictureStyleId,
  OFFICE_PICTURE_STYLE_CATALOG,
  type OfficePictureStyleId,
} from "../office-picture-styles/catalog";
import { nativePictureStyle } from "../office-picture-styles/native";

const OFFICE_PICTURE_STYLE_IDS: readonly OfficePictureStyleId[] = OFFICE_PICTURE_STYLE_CATALOG.map(
  (entry) => entry.id,
);

export const SMARTART_PICTURE_CROP_KEY = "smartArtPictureCrop";
export const SMARTART_PICTURE_STYLE_KEY = "smartArtPictureStyle";

export interface SmartArtSlotPictureFormat {
  readonly crop?: CropRect;
  readonly style?: OfficePictureStyleId;
}

/** A slot's crop, or `null` for none / anything that is not one. */
export function readSmartArtSlotPictureCrop(value: unknown): CropRect | null {
  const crop = normalizeSignedCropRect(value);
  return crop && !(crop.x === 0 && crop.y === 0 && crop.w === 1 && crop.h === 1) ? crop : null;
}

/** A slot's Picture Style id, or `null`. */
export function readSmartArtSlotPictureStyle(value: unknown): OfficePictureStyleId | null {
  return isOfficePictureStyleId(value) ? value : null;
}

/** What a host's slot member states (its `smartArtPicture…` fields). */
export function smartArtSlotPictureFormatOf(member: unknown): SmartArtSlotPictureFormat {
  const record = (member ?? {}) as Record<string, unknown>;
  const crop = readSmartArtSlotPictureCrop(record[SMARTART_PICTURE_CROP_KEY]);
  const style = readSmartArtSlotPictureStyle(record[SMARTART_PICTURE_STYLE_KEY]);
  return { ...(crop ? { crop } : {}), ...(style ? { style } : {}) };
}

/** The same, off an engine image node / slide block (`cropRect`, `pictureStylePreset`). */
export function smartArtSlotPictureFormatOfBlock(block: unknown): SmartArtSlotPictureFormat {
  const record = (block ?? {}) as Record<string, unknown>;
  const crop = readSmartArtSlotPictureCrop(record.cropRect);
  const style = readSmartArtSlotPictureStyle(record.pictureStylePreset);
  return { ...(crop ? { crop } : {}), ...(style ? { style } : {}) };
}

/** A member's `smartArtPicture…` fields for `format` (the keys to spread). */
export function smartArtSlotPictureFields(
  format: SmartArtSlotPictureFormat,
): { smartArtPictureCrop?: CropRect; smartArtPictureStyle?: OfficePictureStyleId } {
  return {
    ...(format.crop ? { [SMARTART_PICTURE_CROP_KEY]: format.crop } : {}),
    ...(format.style ? { [SMARTART_PICTURE_STYLE_KEY]: format.style } : {}),
  };
}

export type SmartArtSlotPictureCommand =
  | { readonly kind: "crop"; readonly mode: "fill" | "fit" }
  | { readonly kind: "style"; readonly style: OfficePictureStyleId }
  | { readonly kind: "reset" };

/**
 * The slot's next Picture Format after one command, or `null` when the command
 * changes nothing or cannot be computed (Fill / Fit need the picture's own
 * pixel size and the slot's box). A Picture Style keeps the crop, as Office's
 * does; Reset Picture removes both (Office: "Reset Picture" discards the
 * formatting changes made to the picture).
 */
export function smartArtSlotPictureFormatAfter(
  current: SmartArtSlotPictureFormat,
  command: SmartArtSlotPictureCommand,
  sizes?: {
    readonly picture?: { readonly width: number; readonly height: number } | null;
    readonly slot?: { readonly width: number; readonly height: number } | null;
  },
): SmartArtSlotPictureFormat | null {
  if (command.kind === "reset") {
    return current.crop || current.style ? {} : null;
  }
  if (command.kind === "style") {
    if (current.style === command.style) return null;
    return { ...current, style: command.style };
  }
  const picture = sizes?.picture;
  const slot = sizes?.slot;
  if (!picture || !slot) return null;
  const crop =
    command.mode === "fill"
      ? pictureFillCropRect(picture.width, picture.height, slot.width, slot.height)
      : pictureFitCropRect(picture.width, picture.height, slot.width, slot.height);
  if (!crop) return null;
  const kept = readSmartArtSlotPictureCrop(crop);
  const { crop: _previous, ...rest } = current;
  return kept ? { ...rest, crop: kept } : rest;
}

/**
 * Where the picture is drawn inside the slot's box, as percentages of the
 * box: the source's kept rectangle stretched over the box (`a:srcRect` with
 * `a:stretch`). No crop: the whole picture over the box, centred and
 * covering it (Office's Fill, which is how a slot shows a picture put in it).
 */
export function smartArtSlotPictureBox(
  crop: CropRect | null | undefined,
): { left: number; top: number; width: number; height: number } | null {
  const kept = readSmartArtSlotPictureCrop(crop);
  if (!kept) return null;
  return {
    left: (-kept.x / kept.w) * 100,
    top: (-kept.y / kept.h) * 100,
    width: (1 / kept.w) * 100,
    height: (1 / kept.h) * 100,
  };
}

/**
 * The Picture Style a point's `spPr` states, recognised by the parts every
 * recipe differs in (geometry, line width and colour, each effect's kind and
 * offsets, the camera and the bevel): unique over Office's 28 styles. `null`
 * when it is none of them.
 */
export function smartArtPictureStyleFromSignature(signature: {
  readonly geometry: string | null;
  readonly lineWidth: number;
  readonly lineColor: string;
  readonly effects: readonly { kind: string; blurRad: string; dist: string; dir: string }[];
  readonly camera: string;
  readonly bevelWidth: string;
}): OfficePictureStyleId | null {
  if (!signature.geometry) return null;
  const key = (parts: {
    geometry: string;
    lineWidth: number;
    lineColor: string;
    effects: readonly { kind: string; blurRad?: unknown; dist?: unknown; dir?: unknown }[];
    camera: string;
    bevelWidth: string;
  }) =>
    [
      parts.geometry,
      parts.lineWidth,
      parts.lineColor.toLowerCase(),
      parts.effects
        .map((effect) => [effect.kind, effect.blurRad ?? "", effect.dist ?? "", effect.dir ?? ""].join(":"))
        .sort()
        .join("+"),
      parts.camera,
      parts.bevelWidth,
    ].join("|");
  const wanted = key({ ...signature, geometry: signature.geometry });
  for (const id of OFFICE_PICTURE_STYLE_IDS) {
    const style = nativePictureStyle(id);
    if (!style) continue;
    const candidate = key({
      geometry: style.geometry,
      lineWidth: Number(style.line?.width ?? 0) || 0,
      lineColor: style.line?.color ?? "",
      effects: (style.effects ?? []).map((effect) => ({
        kind: String(effect.kind),
        blurRad: effect.blurRad ?? "",
        dist: effect.dist ?? "",
        dir: effect.dir ?? "",
      })),
      camera: style.camera?.prst ?? "",
      bevelWidth: style.threeD?.bevel?.w ?? "",
    });
    if (candidate === wanted) return id;
  }
  return null;
}

/** The style's geometry, as the clip a host draws the slot in (`rect` when none). */
export function smartArtSlotPictureStyleGeometry(style: unknown): string | null {
  return nativePictureStyle(style)?.geometry ?? null;
}
