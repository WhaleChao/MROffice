/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Front-end feature gates for additive Docs domain nodes. */

/**
 * Rollout gate for the native Docs online-video node. The schema is always
 * registered so an existing node remains loadable/preservable while OFF.
 * Authoring and user-initiated third-party playback default ON; an explicit
 * `0` is the emergency rollback.
 */
export function isDocumentOnlineVideoV1Enabled(): boolean {
  return process.env.NEXT_PUBLIC_DOCS_ONLINE_VIDEO_V1 !== "0";
}
