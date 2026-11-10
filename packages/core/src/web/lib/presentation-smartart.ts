/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation-facing façade over the shared SmartArt MODEL.
 *
 * The generator lives in `lib/office-smartart/engine.ts` and runs on the
 * server (the engine host); the deck asks for every layout through
 * `lib/office-smartart/remote.ts`. What the deck reads off its own blocks --
 * the kinds, the member markers, the outline -- is re-exported here under the
 * names it always had, so no presentation import path changes and persisted
 * `smartArtKind` markers keep their legacy ids.
 */

export * from "./office-smartart/model";
export * from "./office-smartart/outline";
