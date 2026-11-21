/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Regional indicators as Windows Excel prints them (lookwin p254-lw): Excel
// does not join two regional indicator symbols (U+1F1E6-1F1FF) into a flag;
// the Segoe UI Emoji its cell's face falls back to prints each as its letter,
// a small semibold capital half an em high, at its own advance (R464-X05's
// 🇧🇪 "BE", 0.9 em, where the browser drew a 1.2 em flag). The face
// (public/fonts/regional-letters, scripts/build-regional-letters-font.py) is
// Selawik Semibold's capitals at those advances, the same the API's sheet PDF
// writer builds (office_regional_letters.py). A cell stack names it right
// behind the cell's face, so only those characters fall to it; it is the
// spreadsheet's alone (no Docs or presentation stack names it).

import { noteFontFaceSource } from "../font-face-declarations";

/** Segoe UI Emoji's advance of each regional indicator letter A-Z, em (p254-lw). */
export const OFFICE_REGIONAL_LETTER_ADVANCES: readonly number[] = [
  0.469, 0.484, 0.462, 0.55, 0.417, 0.417, 0.529, 0.577, 0.32, 0.298, 0.497, 0.385, 0.702,
  0.586, 0.586, 0.462, 0.586, 0.454, 0.424, 0.41, 0.534, 0.479, 0.723, 0.459, 0.417, 0.452,
];

/** The family; its "MRAI Excel " prefix keeps it out of the Office family lookups (servedStack). */
export const OFFICE_REGIONAL_LETTERS_FAMILY = "MRAI Excel regional letters";
const FILE = "/fonts/regional-letters/regional-letters-600.woff2";
const RANGE = "U+1F1E6-1F1FF";

/** Excel's advance of a regional indicator, em; undefined for any other character. */
export function officeRegionalLetterAdvance(character: string): number | undefined {
  const code = character.codePointAt(0) ?? 0;
  if ([...character].length !== 1 || code < 0x1f1e6 || code > 0x1f1ff) return undefined;
  return OFFICE_REGIONAL_LETTER_ADVANCES[code - 0x1f1e6];
}

const declared = new WeakSet<Document>();

/** Declares the face once per document (regular and bold alike: Excel's letters have no bold). */
export function ensureOfficeRegionalLettersFace(
  doc: Document | undefined = typeof document === "undefined" ? undefined : document,
): void {
  if (!doc?.fonts || typeof FontFace === "undefined" || declared.has(doc)) return;
  declared.add(doc);
  const src = `url(${FILE}) format("woff2")`;
  for (const weight of ["400", "700"]) {
    try {
      const face = new FontFace(OFFICE_REGIONAL_LETTERS_FAMILY, src, { unicodeRange: RANGE, weight });
      doc.fonts.add(face);
      // Noted, so the text engine reads the face as the browser draws it
      // (an undeclared family in a stack stopped the engine's measure).
      noteFontFaceSource(face, src, undefined, {
        doc,
        declared: { family: OFFICE_REGIONAL_LETTERS_FAMILY, weight, style: "normal", unicodeRange: RANGE },
      });
    } catch {
      // No face: the stack's next family draws them, as before.
    }
  }
}

/** The quoted family, comma-terminated, a cell stack carries right behind its face. */
export function officeRegionalLettersStack(): string {
  ensureOfficeRegionalLettersFace();
  return `"${OFFICE_REGIONAL_LETTERS_FAMILY}", `;
}
