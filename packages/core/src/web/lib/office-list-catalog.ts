/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shared Office list-gallery catalog.
 *
 * Docs owns the complete Word list library while Presentation exposes the
 * smaller PowerPoint-for-the-web subset. Both read the same stable style ids
 * and glyphs from this pure module so an editor adapter can filter the catalog
 * without copying or renaming an item.
 */

export const OFFICE_BULLET_STYLE_GLYPHS = {
  disc: "•",
  circle: "○",
  square: "▪",
  diamond: "◆",
  triangle: "‣",
  arrow: "➤",
  check: "✓",
  star: "✦",
  dash: "–",
  "black-circle": "●",
  "black-square": "■",
  "white-square": "□",
  "five-point-star": "★",
  "right-arrow": "→",
  plus: "+",
  "black-diamond": "♦",
  lozenge: "◊",
  "four-diamond": "❖",
  arrowhead: "➢",
  smile: "☺",
  frown: "☹",
  cross: "✖",
  "shadow-square": "❒",
} as const;

export type OfficeBulletStyleType = keyof typeof OFFICE_BULLET_STYLE_GLYPHS;

export const OFFICE_BULLET_STYLE_TYPES = Object.freeze(
  Object.keys(OFFICE_BULLET_STYLE_GLYPHS) as OfficeBulletStyleType[],
);

/** Frozen order from Microsoft's PowerPoint-for-the-web bullet gallery. */
export const POWERPOINT_WEB_BULLET_STYLE_TYPES = [
  "disc",
  "circle",
  "square",
  "white-square",
  "diamond",
  "arrow",
  "check",
] as const satisfies readonly OfficeBulletStyleType[];

/** Frozen order from Microsoft's PowerPoint-for-the-web numbering gallery. */
export const POWERPOINT_WEB_NUMBER_FORMATS = [
  "decimal",
  "decimal-paren",
  "upper-roman",
  "upper-alpha",
  "lower-alpha-paren",
  "lower-alpha",
  "lower-roman",
] as const;

/**
 * Every glyph a PowerPoint paragraph may durably carry: the seven the web
 * gallery shows, plus the two legacy MRAI choices existing saved pages use.
 * The middle dot is the level-3 DEFAULT and deliberately not authorable here.
 */
export const POWERPOINT_BULLET_CHARS = [
  "•",
  "○",
  "▪",
  "□",
  "◆",
  "➤",
  "✓",
  "–",
  "→",
] as const;

/** PowerPoint's own five-level ladder: solid -> en dash -> middle dot. */
export const POWERPOINT_DEFAULT_BULLET_GLYPH_BY_LEVEL = [
  "•",
  "–",
  "·",
  "•",
  "–",
] as const;

/** The matching numbered ladder. */
export const POWERPOINT_DEFAULT_NUMBER_FORMAT_BY_LEVEL = [
  "decimal",
  "lower-alpha",
  "lower-roman",
  "decimal",
  "lower-alpha",
] as const;

/**
 * The editor's number format <-> the DrawingML `ST_TextAutonumberScheme` the
 * writers put back. Seven canonical spellings; a similar-but-different native
 * scheme stays package-authoritative through the `nativeBullet` receipt rather
 * than being collapsed onto one of these.
 */
export const POWERPOINT_AUTONUM_BY_NUMBER_FORMAT = {
  decimal: "arabicPeriod",
  "decimal-paren": "arabicParenR",
  "lower-alpha": "alphaLcPeriod",
  "lower-alpha-paren": "alphaLcParenR",
  "lower-roman": "romanLcPeriod",
  "upper-alpha": "alphaUcPeriod",
  "upper-roman": "romanUcPeriod",
} as const;

export const POWERPOINT_NUMBER_FORMAT_BY_AUTONUM = Object.freeze(
  Object.fromEntries(
    Object.entries(POWERPOINT_AUTONUM_BY_NUMBER_FORMAT).map(
      ([format, scheme]) => [scheme, format],
    ),
  ) as {
    [S in (typeof POWERPOINT_AUTONUM_BY_NUMBER_FORMAT)[keyof typeof POWERPOINT_AUTONUM_BY_NUMBER_FORMAT]]: {
      [F in keyof typeof POWERPOINT_AUTONUM_BY_NUMBER_FORMAT]: (typeof POWERPOINT_AUTONUM_BY_NUMBER_FORMAT)[F] extends S
        ? F
        : never;
    }[keyof typeof POWERPOINT_AUTONUM_BY_NUMBER_FORMAT];
  },
);
