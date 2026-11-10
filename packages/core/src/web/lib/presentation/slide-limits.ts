/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The most slides a deck page holds: the page schema, the load edge and every
 * structural command (New Slide, Duplicate, Paste, Reuse) enforce it, so an
 * in-memory slide past it can never be saved and then disappear on reload.
 * The server's twin is `MAX_PRESENTATION_SLIDES` in `apps/api/app/routers/notebooks.py`.
 *
 * It bounds work, not size: what a deck may hold is the page's JSON budget
 * (10 MiB, 400,000 nodes), which a deck's own content reaches first once each
 * receipt its slides repeat is stored once (`presentation-deck-receipts`).
 */
export const MAX_PRESENTATION_SLIDES = 2000;
