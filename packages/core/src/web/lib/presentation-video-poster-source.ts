/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Video Format > Poster Frame (PPT-A155): what a poster the user set is on the
 * page -- an inline PNG or JPEG data URL within the API's one-string bound.
 * Dependency-free so the page schema can import it.
 */
export const MAX_VIDEO_POSTER_SRC_LENGTH = 1_048_576;

export function isVideoPosterSource(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= MAX_VIDEO_POSTER_SRC_LENGTH &&
    /^data:image\/(png|jpeg);base64,[A-Za-z0-9+/]+={0,2}$/.test(value)
  );
}
