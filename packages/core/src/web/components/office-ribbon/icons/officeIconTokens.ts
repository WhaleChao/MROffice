/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Shared geometry and state tokens for original MRAI Office ribbon icons. */
export const OFFICE_ICON_VIEW_BOX = "0 0 24 24";

export const OFFICE_ICON_OPTICAL_SAFE_AREA = Object.freeze({
  min: 2,
  max: 22,
  size: 20,
});

export const OFFICE_ICON_STROKE_WIDTH = 1;
export const OFFICE_ICON_EMPHASIS_STROKE_WIDTH = 1.4;

export const OFFICE_ICON_RENDER_SIZES = [16, 18, 24, 32] as const;

export const OFFICE_ICON_VISUAL_STATES = [
  "default",
  "hover",
  "active",
  "pressed",
  "disabled",
  "focus",
] as const;

export type OfficeIconVisualState = (typeof OFFICE_ICON_VISUAL_STATES)[number];

/**
 * Public custom-property names. Hosts may override these roles without
 * coupling semantic glyph geometry to an editor theme implementation.
 */
export const OFFICE_ICON_CSS_TOKENS = Object.freeze({
  foreground: "--mrai-office-icon-foreground",
  foregroundSoft: "--mrai-office-icon-foreground-soft",
  surface: "--mrai-office-icon-surface",
  accent: "--mrai-office-icon-accent",
  accentSoft: "--mrai-office-icon-accent-soft",
  warm: "--mrai-office-icon-warm",
  warmSoft: "--mrai-office-icon-warm-soft",
  disabled: "--mrai-office-icon-disabled",
  transitionDuration: "--mrai-office-icon-transition-duration",
});
