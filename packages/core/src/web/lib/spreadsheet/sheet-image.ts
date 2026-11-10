/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isValidInlineImageSource } from "../inline-image-source";
import {
  pictureObjectTokenSha,
  pictureObjectUrlForToken,
} from "../page-picture-objects";

/**
 * Insert Picture (Excel for the web "Insert ▸ Pictures") — pure helpers for
 * adding an image object to a sheet. New pictures must be embedded data images:
 * the native XLSX writer cannot package an ordinary network URL, so accepting
 * one here would create a picture that disappears after export/reopen.
 *
 * The shared inline-image guard verifies base64 shape, declared MIME and magic
 * bytes. Keeping that one parser authoritative also aligns Spreadsheet with the
 * existing Docs and Presentation local-picture paths.
 */

import {
  OFFICE_PICTURE_FILE_MAX_BYTES,
  OFFICE_PICTURE_SRC_MAX_LENGTH,
} from "../office-picture-envelope";

export const SPREADSHEET_IMAGE_FILE_MIME_TYPES = Object.freeze([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp",
  "image/bmp",
] as const);

/**
 * Max length of an accepted image source: the page envelope's longest single
 * string (lib/office-picture-envelope). A data: URL past it is not a picture
 * the save trims but a workbook the save refuses whole, so nothing this long
 * may reach a sheet (sheet-image-fit re-encodes a larger picture down to it).
 */
export const MAX_IMAGE_SRC_LENGTH = OFFICE_PICTURE_SRC_MAX_LENGTH;

/**
 * The largest local file whose base64 data: URL fits MAX_IMAGE_SRC_LENGTH:
 * the ceiling for a file the fit step cannot re-encode (GIF, BMP).
 */
export const MAX_SPREADSHEET_IMAGE_FILE_BYTES = OFFICE_PICTURE_FILE_MAX_BYTES;

/**
 * The shape of an image object this module produces. Mirrors the spreadsheet's
 * `SpreadsheetObject` image variant (id/type/src/alt/caption/anchor) without
 * importing the component's type, so the helper stays dependency-light. The
 * anchor's `from` is the cell the picture is dropped onto.
 */
export interface SheetImageObject {
  id: string;
  type: "image";
  src: string;
  alt: string;
  caption: string;
  anchor: { from: { row: number; column: number } };
}

/**
 * Validate + normalize an embedded image source. Ordinary HTTP(S) sources are
 * deliberately rejected because they are not packageable by the XLSX writer.
 * The length and decoded bytes are bounded, and the MIME must match the image
 * magic bytes.
 */
export function sanitizeImageSource(raw: unknown): string | null {
  if (typeof raw !== "string") return null;
  const value = raw.trim();
  if (value === "" || value.length > MAX_IMAGE_SRC_LENGTH) return null;
  const match = /^data:(image\/[a-z0-9.+-]+);base64,/i.exec(value);
  if (!match) return null;
  const mimeType =
    match[1].toLowerCase() === "image/jpg"
      ? "image/jpeg"
      : match[1].toLowerCase();
  if (
    !SPREADSHEET_IMAGE_FILE_MIME_TYPES.includes(
      mimeType as (typeof SPREADSHEET_IMAGE_FILE_MIME_TYPES)[number],
    )
  ) {
    return null;
  }
  return isValidInlineImageSource(value) ? value : null;
}

/**
 * A floating picture's source as the sheet keeps it: an embedded data image
 * ({@link sanitizeImageSource}), or a picture an imported workbook stores
 * beside its page, named by its token (`mrai-picture:<sha256>`,
 * lib/page-picture-objects). The token is kept as it is; the server requires
 * it to be one of the page's own pictures (or copies one the user can read).
 * A floating picture and a cell picture may be a token; a background and a
 * header/footer picture stay embedded.
 */
export function sanitizeSheetPictureSource(raw: unknown): string | null {
  if (typeof raw === "string" && pictureObjectTokenSha(raw.trim()) !== null) {
    return raw.trim();
  }
  return sanitizeImageSource(raw);
}

/**
 * A stored picture's bytes as an embedded data image fitted to the page's
 * single-string bound, for a place that holds only embedded pictures (Place
 * in Cell). `null` when the picture cannot be read or fitted.
 */
export async function embeddedSheetPictureSource(token: string): Promise<string | null> {
  const url = pictureObjectUrlForToken(token);
  if (!url || typeof fetch === "undefined") return null;
  try {
    const response = await fetch(url, { credentials: "same-origin" });
    if (!response.ok) return null;
    const blob = await response.blob();
    const { readFittedPicture } = await import("../office-picture-envelope");
    return sanitizeImageSource(await readFittedPicture(blob));
  } catch {
    return null;
  }
}

let imageSeq = 0;

/** An SVG file (Insert > Pictures takes one, as Excel does). */
export function isSpreadsheetSvgPictureFile(file: { type: string; name: string }): boolean {
  const type = file.type.trim().toLowerCase();
  return type === "image/svg+xml" || (!type && /\.svg$/i.test(file.name));
}

/**
 * An SVG picture object: the passive SVG itself (`officeGraphicKind: "svg"`,
 * lib/presentation-office-graphic-file `preparePassiveSvgPicture`), drawn as
 * it is. The workbook writer stores it as Excel does -- the SVG beside a PNG
 * the server draws ([MS-ODRAWXML] asvg:svgBlip). Null for a source that is
 * no base64 SVG data URL within the page's single-string bound.
 */
export function buildSvgImageObject(
  src: string,
  anchorCell: { row: number; col: number },
  alt?: string,
): (SheetImageObject & { officeGraphicKind: "svg" }) | null {
  if (
    typeof src !== "string" ||
    src.length > MAX_IMAGE_SRC_LENGTH ||
    !/^data:image\/svg\+xml;base64,[A-Za-z0-9+/]+={0,2}$/.test(src)
  ) {
    return null;
  }
  imageSeq += 1;
  return {
    id: `image-${imageSeq}-${Date.now().toString(36)}`,
    type: "image",
    src,
    alt: typeof alt === "string" ? alt.trim().slice(0, 500) : "",
    caption: "",
    anchor: { from: { row: clampCell(anchorCell.row), column: clampCell(anchorCell.col) } },
    officeGraphicKind: "svg",
  };
}

/**
 * Build an image SpreadsheetObject from a (sanitized) source and an anchor cell.
 * Returns null when the source is not a safe image reference (see
 * {@link sanitizeImageSource}). The id is unique within the process; alt/caption
 * default to empty (the user edits them in the existing image panel). The anchor
 * row/col are floored and clamped to ≥ 0 so a stray negative never positions the
 * object off-grid.
 *
 * @param src the user-supplied source (validated here)
 * @param anchorCell the cell the picture is dropped onto (0-based row/col)
 * @param alt optional alt text (trimmed, length-bounded)
 */
export function buildImageObject(
  src: unknown,
  anchorCell: { row: number; col: number },
  alt?: string,
): SheetImageObject | null {
  const safeSrc = sanitizeImageSource(src);
  if (safeSrc === null) return null;
  const row = clampCell(anchorCell.row);
  const col = clampCell(anchorCell.col);
  imageSeq += 1;
  const altText =
    typeof alt === "string" ? alt.trim().slice(0, 500) : "";
  return {
    id: `image-${imageSeq}-${Date.now().toString(36)}`,
    type: "image",
    src: safeSrc,
    alt: altText,
    caption: "",
    anchor: { from: { row, column: col } },
  };
}

function clampCell(raw: number): number {
  if (!Number.isFinite(raw)) return 0;
  return Math.max(0, Math.floor(raw));
}
