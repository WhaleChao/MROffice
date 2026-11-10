/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A Docs type size as the document stores it: CSS px rounded to four
 * decimals and kept to six significant digits, the spelling the DOCX import
 * writes (projection.py `_typography_font_size`: `f"{round(px, 4):g}px"`).
 * Every writer of a size uses it, so one size has one spelling: an imported
 * 11pt run said "14.6667px", the ribbon's 11pt "14.67px" and a stored "11pt"
 * read as "14.667px" — three values for one size wherever sizes are compared.
 */
import { ptToPx } from "./office-units";

/** A size in CSS px, spelled as the document stores it. */
export function canonicalDocumentFontSizePx(px: number): string {
  const rounded = Math.round(px * 10_000) / 10_000;
  return `${Number(rounded.toPrecision(6))}px`;
}

/** A size in points, spelled as the document stores it. */
export function documentFontSizePx(points: number): string {
  return canonicalDocumentFontSizePx(ptToPx(points));
}
