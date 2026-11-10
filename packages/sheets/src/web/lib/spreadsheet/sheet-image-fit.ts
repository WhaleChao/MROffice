/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The spreadsheet's view of the shared inline-picture envelope
 * (lib/office-picture-envelope): an inserted or replaced picture is fitted
 * to the page's single-string bound before it reaches a sheet.
 */

export {
  fitOfficePictureSource as fitSpreadsheetPictureSource,
  officePictureFileByteCeiling as spreadsheetPictureFileByteCeiling,
  FITTABLE_PICTURE_MIME_TYPES,
  MAX_FITTABLE_PICTURE_FILE_BYTES,
  FittedPictureReader,
} from "@mrnote/core/web/lib/office-picture-envelope.ts";
