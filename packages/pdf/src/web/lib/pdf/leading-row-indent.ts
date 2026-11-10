/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfLeadingRowIndent = Readonly<{
  /** How far right of the paragraph's left edge the first rows start (px). */
  indent: number;
  /** How many rows start there. */
  rows: number;
  /** The distance from one row's baseline to the next (px). */
  pitch: number;
}>;

type Character = Readonly<{
  quad: readonly number[];
  origin: readonly number[];
}>;

export function pdfLeadingRowIndent(characters: readonly Character[]): PdfLeadingRowIndent | null {
  const rows: Array<{ y: number; left: number; size: number }> = [];
  for (const character of characters) {
    const xs = [character.quad[0], character.quad[2], character.quad[4], character.quad[6]];
    const ys = [character.quad[1], character.quad[3], character.quad[5], character.quad[7]];
    if (![...xs, ...ys, character.origin[1]].every(Number.isFinite)) continue;
    // The glyph box's own height: the quads are in the caller's pixels, the
    // font size in the PDF's units.
    const size = Math.max(...ys) - Math.min(...ys) || 1;
    const last = rows[rows.length - 1];
    if (!last || Math.abs(character.origin[1] - last.y) > Math.max(last.size, size) * 0.5) {
      rows.push({ y: character.origin[1], left: Math.min(...xs), size });
    } else {
      last.left = Math.min(last.left, ...xs);
      last.size = Math.max(last.size, size);
    }
  }
  if (rows.length < 2) return null;
  const base = Math.min(...rows.map((row) => row.left));
  const size = Math.max(...rows.map((row) => row.size));
  const threshold = Math.max(4, size * 0.6);
  let leading = 0;
  while (leading < rows.length && rows[leading].left - base > threshold) leading += 1;
  // Some rows indented and the rest at the paragraph's edge: not a first-line
  // indent alone (that the draft's own first-line indent covers), not a
  // whole paragraph set further right.
  if (leading < 1 || leading >= rows.length) return null;
  if (!rows.slice(leading).every((row) => row.left - base <= size * 0.3)) return null;
  if (leading === 1) return null;
  const gaps = rows.slice(1).map((row, index) => Math.abs(row.y - rows[index].y)).sort((a, b) => a - b);
  const pitch = gaps[Math.floor(gaps.length / 2)];
  if (!(pitch > 0)) return null;
  return Object.freeze({
    indent: Math.min(...rows.slice(0, leading).map((row) => row.left)) - base,
    rows: leading,
    pitch,
  });
}
