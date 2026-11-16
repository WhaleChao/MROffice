/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Which Office faces join which Latin letters into a ligature no code point
// maps to (office-latin-ligatures.json, the API's office_latin_ligatures.json:
// the same table, measured from the Office faces by lookwin
// tools/ligature_table.py). Windows Excel's real print (File > Print) draws a
// cell's text unclipped -- past the page's margins to the paper's edge, on
// across the next page from the paper's left edge, over the values in its
// way -- when the text holds a line break, does not wrap, and its words take
// such a ligature of the cell's face (lookwin p100-lw .. p102h-lw, p104-lw:
// Calibri's ti, tt, ft, fb; fi, fl, ff, ffi map to U+FB00..FB04 and are cut
// at the margins as ever). The API's print route draws the same
// (notebooks.py cell look "unclipped").
import table from "./office-latin-ligatures.json";

const faces = table.faces as Readonly<Record<string, readonly string[]>>;

/** The Office face a CSS family or stack names first, as the table keys it. */
function faceKey(family: string | null | undefined): string {
  const first = (family ?? "").split(",")[0] ?? "";
  return first.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();
}

/** Whether `text` set in the Office face `family` forms a ligature no code point maps to. */
export function officeTextTakesLigature(family: string | null | undefined, text: string): boolean {
  if (!text) return false;
  const sequences = faces[faceKey(family)];
  return Boolean(sequences?.length) && sequences.some((sequence) => text.includes(sequence));
}

/** Whether Windows Excel's real print draws a cell's text unclipped (above). */
export function excelPrintsTextUnclipped(family: string | null | undefined, text: string, wraps: boolean): boolean {
  return !wraps && /[\r\n]/.test(text) && officeTextTakesLigature(family, text);
}
