"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

import { getEmbedAccess, subscribeEmbedAccess, type EmbedAccess } from "./session-access";

/** The embed session's access (lib/embed/session-access), re-rendering when
 * it changes (a new token's mode); null outside an embedded editor. */
export function useEmbedAccess(): EmbedAccess | null {
  return useSyncExternalStore(subscribeEmbedAccess, getEmbedAccess, () => null);
}
