/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw with Trackpad's one motion: the switch's knob crosses its track while
 * the drawing frame gathers from the edges of the document, over the same
 * time on the same curve. Keep the switch's CSS (OfficeInkRibbon.module.css,
 * .inkSwitchKnob) on these values.
 */
export const OFFICE_INK_TRACKPAD_TRANSITION_MS = 460;
export const OFFICE_INK_TRACKPAD_EASING = "cubic-bezier(0.32, 0.72, 0, 1)";
