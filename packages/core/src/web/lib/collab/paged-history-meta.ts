/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PAGED_AWAITS_FOLLOW_UP_META = "mrnotePagedAwaitsFollowUp";
/** What the layout derived from the step that awaits it (see `PAGED_AWAITS_FOLLOW_UP_META`). */
export const PAGED_FOLLOW_UP_META = "mrnotePagedFollowUp";
/** Dispatched by the paged bridge once its own Undo/Redo has landed in the
 * view (value: "undo" | "redo"), while the operation still counts as this
 * view's history: a plugin may append the maintenance that operation calls
 * for, and it is published outside personal history. */
export const PAGED_HISTORY_APPLIED_META = "mrnotePagedHistoryApplied";
