/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PX_PER_PT, pxToPt } from "./office-units";
/**
 * CSS absolute lengths (CSS Values 4, section 6.2) in CSS pixels.
 *
 * Pasted and imported HTML states sizes in any of them: `24pt`, `0.5in`,
 * `32px` are the same length family and one cannot be read while another is
 * dropped. Relative units (em, %, rem, vw) depend on context this module does
 * not have and yield `null`, as does anything that is not one plain number
 * with one unit.
 */
const PX_PER_UNIT: Readonly<Record<string, number>> = {
  px: 1,
  pt: PX_PER_PT,
  pc: 16,
  in: 96,
  cm: 96 / 2.54,
  mm: 96 / 25.4,
  q: 96 / 101.6,
};

const ABSOLUTE_LENGTH_RE = /^([+-]?(?:\d+(?:\.\d+)?|\.\d+))(px|pt|pc|in|cm|mm|q)$/;

export function cssAbsoluteLengthPx(value: unknown): number | null {
  if (typeof value !== "string") return null;
  const match = ABSOLUTE_LENGTH_RE.exec(value.trim().toLowerCase());
  if (!match) return null;
  const px = Number.parseFloat(match[1]!) * PX_PER_UNIT[match[2]!]!;
  return Number.isFinite(px) ? px : null;
}

/** The same length in points (1pt = 1/72in = 4/3 px). */
export function cssAbsoluteLengthPt(value: unknown): number | null {
  const px = cssAbsoluteLengthPx(value);
  return px === null ? null : pxToPt(px);
}
