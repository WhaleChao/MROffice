/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Shape > Fill > Picture or texture fill (and the ribbon's Shape Fill
 * > Picture...) for the Docs and Spreadsheet editors' shapes: the file picker
 * and the inline picture a shape is filled with.
 *
 * Office stretches the picture over the shape's box and draws it inside the
 * geometry (DrawingML `a:blipFill` + `a:stretch/a:fillRect` in the shape's
 * `spPr`, in the fill slot), the words on top. The picture is held inline as
 * a PNG or JPEG `data:` URL -- the two raster codecs every MRAI writer turns
 * into a blip, and the spelling both shape models already validate
 * (`office-smartart/picture`) -- fitted to the page envelope's bound
 * (`office-picture-envelope`), as an inserted picture is.
 */

import { readFittedPicture, officePictureFileByteCeiling } from "./office-picture-envelope";
import { smartArtPictureSrc } from "./office-smartart/picture";
import type { OfficePictureFillTile } from "./office-picture-fill-options";
import { OFFICE_TEXTURE_TILE, officeTextureDataUrl, officeTextureSourceId } from "./office-textures";

/** The kinds the picker offers: what a fill can hold. */
export const OFFICE_PICTURE_FILL_ACCEPT = "image/png,image/jpeg";

/** The fill picture `value` names, or null when it is not one a shape can hold. */
export function officePictureFillSrc(value: unknown): string | null {
  return smartArtPictureSrc(value);
}

/**
 * A picked file as a shape's fill picture: fitted to the page bound and held
 * to PNG / JPEG, or null when it cannot be.
 */
export async function readOfficePictureFill(file: Blob): Promise<string | null> {
  if (file.size > officePictureFileByteCeiling(file.type)) return null;
  const type = file.type.trim().toLowerCase();
  if (type !== "image/png" && type !== "image/jpeg" && type !== "image/jpg") return null;
  const fitted = await readFittedPicture(file);
  return officePictureFillSrc(fitted);
}

/** Office's "Picture source": Insert... (a file), Clipboard, or a Texture (`texture:<id>`). */
export type OfficePictureFillSource = "file" | "clipboard" | `texture:${string}`;

/**
 * Picture source > Clipboard: the first PNG or JPEG picture on the system
 * clipboard as a fill (null when there is none, or the browser refuses to
 * read it).
 */
export async function readOfficePictureFillFromClipboard(): Promise<string | null> {
  const clipboard = typeof navigator === "undefined" ? undefined : navigator.clipboard;
  if (!clipboard || typeof clipboard.read !== "function") return null;
  const items = await clipboard.read();
  for (const item of items) {
    const type = item.types.find((candidate) => candidate === "image/png" || candidate === "image/jpeg");
    if (type) return readOfficePictureFill(await item.getType(type));
  }
  return null;
}

/**
 * Opens the file picker (or reads the clipboard) and answers the picked picture as a fill (null when
 * nothing usable was picked). The host re-checks what it fills once the
 * picture is read: the selection may have moved while the picker was open.
 */
export function pickOfficePictureFill(
  onPicked: (src: string | null, tile?: OfficePictureFillTile) => void,
  source: OfficePictureFillSource = "file",
): void {
  // A texture is drawn here and tiled as Office tiles one (100 %, top left).
  const texture = officeTextureSourceId(source);
  if (texture) {
    const src = officeTextureDataUrl(texture);
    onPicked(src ? officePictureFillSrc(src) : null, OFFICE_TEXTURE_TILE);
    return;
  }
  if (source === "clipboard") {
    void readOfficePictureFillFromClipboard().then(onPicked, () => onPicked(null));
    return;
  }
  if (typeof document === "undefined") return;
  const input = document.createElement("input");
  input.type = "file";
  input.accept = OFFICE_PICTURE_FILL_ACCEPT;
  input.addEventListener("change", () => {
    const file = input.files?.[0];
    if (!file) return;
    void readOfficePictureFill(file).then(onPicked, () => onPicked(null));
  });
  input.click();
}
