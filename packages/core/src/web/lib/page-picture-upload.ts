/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A picture a page held as a `data:` URI, stored beside the page instead.
 *
 * POST /pages/{id}/pictures keeps the bytes as one of the page's picture
 * objects and answers its token, `mrai-picture:<sha256>` (idempotent per
 * digest); the page names the picture by it from then on
 * (lib/page-picture-objects). Docs (lib/document-picture-objects-upload) and
 * the workbook (spreadsheet/useSpreadsheetPictureObjects) store their pictures
 * this way, so a page of photos is not megabytes of JSON every save carries.
 */

import { apiPostFormData } from "./api";

/** Smaller pictures stay where they are: an upload costs more than they do. */
export const PAGE_PICTURE_OBJECT_MIN_CHARS = 16 * 1024;

/** The types the server keeps as picture objects (`sniffed_picture_media_type`). */
const STORED_PICTURE_DATA_URL = /^data:image\/(?:png|jpe?g|gif|bmp|webp)[;,]/i;

/** Whether a source is a `data:` picture worth storing beside the page. */
export function isStorablePictureDataUrl(src: unknown): src is string {
  return typeof src === "string" && src.length >= PAGE_PICTURE_OBJECT_MIN_CHARS && STORED_PICTURE_DATA_URL.test(src);
}

/** Store a `data:` picture as one of the page's picture objects; its token,
 * or null when the server did not take it. */
export async function uploadPagePicture(pageId: string, dataUrl: string): Promise<string | null> {
  const match = /^data:(image\/[a-z0-9.+-]+)(;[^,]*)?,/i.exec(dataUrl);
  if (!match) return null;
  const blob = dataUrlBlob(dataUrl, match[1], match[2] ?? "", match[0].length);
  if (!blob) return null;
  const form = new FormData();
  const extension = match[1].split("/")[1]?.replace("jpeg", "jpg").replace("+xml", "") || "png";
  form.append("file", blob, `picture.${extension}`);
  const answer = await apiPostFormData<{ src?: unknown }>(`/api/v1/pages/${encodeURIComponent(pageId)}/pictures`, form);
  return typeof answer?.src === "string" && answer.src.startsWith("mrai-picture:") ? answer.src : null;
}

/** The bytes of a `data:` URL, decoded here: `fetch(dataUrl)` is a connection
 * the page's Content Security Policy refuses (`connect-src` names no
 * `data:`; corpus R289-D09 / R301-D16 / R307-D19 logged "Connecting to
 * 'data:image/png;...' violates connect-src" and the picture was never
 * stored beside the page). Null when the payload does not decode. */
function dataUrlBlob(dataUrl: string, type: string, parameters: string, payloadStart: number): Blob | null {
  if (typeof Blob === "undefined") return null;
  const payload = dataUrl.slice(payloadStart);
  try {
    if (!/;base64$/i.test(parameters)) {
      return new Blob([decodeURIComponent(payload)], { type });
    }
    const binary = atob(payload);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes.length ? new Blob([bytes], { type }) : null;
  } catch {
    return null;
  }
}
