/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The ONE spelling a SmartArt picture is stored in, for every host.
 *
 * A picture layout's slot holds its image as an inline `data:` URL — the
 * spelling the presentation author model already validates, the one the shared
 * DiagramML writer already decodes (`office_smartart_diagram_write._inline_image`),
 * and therefore the one the Docs group member and the worksheet shape object
 * adopt rather than each inventing a second.
 *
 * This module is deliberately tiny and imports nothing: the document shape
 * model, the worksheet shape model and the native conversion all need the same
 * answer, and none of them may depend on the SmartArt engine to get it.
 */

/**
 * Office's own two raster codecs for a diagram blip, and the only two the
 * writer decodes. A member carrying anything else is not a picture slot MRAI
 * can export, so it is refused rather than stored and silently dropped later.
 */
const SMARTART_PICTURE_SRC =
  /^data:image\/(?:png|jpeg);base64,[A-Za-z0-9+/]+={0,2}$/;

/**
 * The same bound the API's page-content validator puts on every inline image
 * (`MAX_DOCUMENT_DATA_IMAGE_URL_LENGTH`). Stated here so the browser refuses a
 * payload the server would refuse, instead of letting the author fill a slot
 * that fails to save.
 */
export const SMARTART_MAX_PICTURE_SRC_LENGTH = 2_000_000;

/** The member's picture, or `null` when it declares none this model accepts. */
export function smartArtPictureSrc(value: unknown): string | null {
  return typeof value === "string" &&
    value.length <= SMARTART_MAX_PICTURE_SRC_LENGTH &&
    SMARTART_PICTURE_SRC.test(value)
    ? value
    : null;
}
