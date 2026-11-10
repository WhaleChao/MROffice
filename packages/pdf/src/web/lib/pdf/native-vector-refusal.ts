/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfNativeVectorRefusal =
  | "offPage"
  | "leavesPage"
  | "leavesClip"
  | "shared"
  | "stale"
  | "zeroAxis"
  | "collapse"
  | "noVertices"
  | "lineFill"
  | "notWritable"
  | "unchanged";

/** Ordered: the first family whose pattern the refusal matches. */
const REFUSAL_PATTERNS: readonly (readonly [PdfNativeVectorRefusal, RegExp])[] = [
  ["offPage", /drawn outside (its|their) page/i],
  [
    "leavesPage",
    /outside its CropBox|is outside the page|keep part of the shape on the page|would leave the CropBox|leaves the CropBox/i,
  ],
  ["leavesClip", /leaves? the clip it is drawn under/i],
  ["shared", /shared stream|content stream is shared|drawn more than once/i],
  ["zeroAxis", /zero-thickness/i],
  ["collapse", /collapse|degenerate|singular/i],
  ["noVertices", /no editable vertices/i],
  ["lineFill", /no fill area|open segment cannot be repainted/i],
  ["unchanged", /is unchanged|writes nothing/i],
  [
    "stale",
    /stale|ambiguous|source changed|no longer|changed since|CropBox or rotation changed|proof index changed|does not match|superseded/i,
  ],
  [
    "notWritable",
    /not independently writable|clipped, compound, or partial|unselected PDF vector changed|cannot be deleted|not adjacent|unbalanced|unmatched|unsupported/i,
  ],
];

/** The family a refusal (an error, its message or a server detail) is in. */
export function pdfNativeVectorRefusal(
  refusal: unknown,
): PdfNativeVectorRefusal | null {
  const text =
    refusal instanceof Error
      ? refusal.message
      : typeof refusal === "string"
        ? refusal
        : null;
  if (!text) return null;
  for (const [family, pattern] of REFUSAL_PATTERNS) {
    if (pattern.test(text)) return family;
  }
  return null;
}

/** The i18n key that says why, or null for a refusal this does not know. */
export function pdfNativeVectorRefusalMessageKey(refusal: unknown): string | null {
  const family = pdfNativeVectorRefusal(refusal);
  return family ? `pdf.vectorRefusal.${family}` : null;
}
