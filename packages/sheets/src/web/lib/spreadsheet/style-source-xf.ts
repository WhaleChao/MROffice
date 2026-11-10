/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const SOURCE_XF_STYLE_KEY = "sourceXf";

/** Whether a cell's style states any format; a style holding only the mark does not. */
export function styleStatesAFormat(style: object | null | undefined): boolean {
  if (!style) return false;
  for (const key of Object.keys(style)) {
    if (key !== SOURCE_XF_STYLE_KEY) return true;
  }
  return false;
}
