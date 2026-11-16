/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A line set in the middle of its page (its frame's middle within 3% of
 * the page's width of the page's), in page fractions: a title Word or WPS
 * would hold centred. The editor sets such a line typed on again centred
 * (centredRetype), and the window reads its paragraph ahead for it. */
export function pdfNativeLineCentredOnPage(rect: Readonly<{ x: number; width: number }>): boolean {
  return Math.abs(rect.x + rect.width / 2 - 0.5) < 0.03;
}












export function pdfNativeLineOfPageBody(host: Element | null | undefined, fontSizePixels: number | null): boolean {
  const page = host?.closest(".pdf-viewer__page");
  if (!host || !page) return false;
  const pageBox = page.getBoundingClientRect();
  const own = host.getBoundingClientRect();
  if (!(pageBox.width > 0) || !(own.width > 0)) return false;
  if (fontSizePixels !== null && fontSizePixels > 0 && own.height > fontSizePixels * 1.6) return true;
  const left = (own.left - pageBox.left) / pageBox.width;
  // (Two points of a page 612 pt wide, as two lines ending together are told.)
  return [...page.querySelectorAll<HTMLElement>('span[data-pdf-native-text-edit-box="true"]')]
    .some((box) => {
      if (box === host || !box.textContent?.trim()) return false;
      const other = box.getBoundingClientRect();
      return other.width > 0 && other.height > 0 && Math.abs((other.left - pageBox.left) / pageBox.width - left) < 0.003;
    });
}











export function pdfTextSourceBlockCentred(
  appearance: Readonly<{ characters: readonly Readonly<{ text: string; origin: readonly number[]; quad: readonly number[] }>[] }> | null | undefined,
): boolean {
  const characters = appearance?.characters.filter((character) => character.text.trim() !== "") ?? [];
  if (characters.length < 2) return false;
  const size = Math.max(...characters.map((character) =>
    Math.hypot(character.quad[4] - character.quad[0], character.quad[5] - character.quad[1])));
  if (!(size > 0)) return false;
  const rows: { left: number; right: number }[] = [];
  let baseline = Number.NaN;
  for (const character of characters) {
    const xs = [character.quad[0], character.quad[2], character.quad[4], character.quad[6]];
    const left = Math.min(...xs), right = Math.max(...xs);
    // (A label set tight keeps its rows 0.85 em apart; a raised mark stands
    // less than half an em off its row.)
    if (!rows.length || Math.abs(character.origin[1] - baseline) > size * 0.6) {
      rows.push({ left, right });
      baseline = character.origin[1];
    } else {
      const row = rows[rows.length - 1];
      row.left = Math.min(row.left, left);
      row.right = Math.max(row.right, right);
    }
  }
  if (rows.length < 2 || !rows.every((row) => Number.isFinite(row.left) && Number.isFinite(row.right))) return false;
  const centres = rows.map((row) => (row.left + row.right) / 2);
  const lefts = rows.map((row) => row.left);
  return Math.max(...centres) - Math.min(...centres) <= size * 0.2 &&
    Math.max(...lefts) - Math.min(...lefts) >= size * 0.1;
}



export const PDF_TEXT_BOX_HANDOFF_FIELD_ATTR = "data-pdf-textbox-handoff-field";





export function pdfNativeTextBoxHandoffOpening(document: Document | null | undefined): boolean {
  return Boolean(document?.querySelector(`[${PDF_TEXT_BOX_HANDOFF_FIELD_ATTR}]`));
}

/**
 * Whether a line stands flush with the page's right margin, read from the
 * boxes of its page as they are shown: in the page's right half, its right
 * edge where another line of the page ends too (a cover's code, date and
 * duration set one under another), or where the page's text starts from the
 * left, mirrored. Word would hold such a line right-aligned: typed longer it
 * grows to the left. The editor sets it so (alignedRetype), as a centred
 * line keeps its middle. `fontSizePixels`: the box's font size as shown,
 * by which a box of several rows is told.
 */
export function pdfNativeLineRightFlushOnPage(host: Element | null | undefined, fontSizePixels: number | null = null): boolean {
  const page = host?.closest(".pdf-viewer__page");
  if (!host || !page) return false;
  const pageBox = page.getBoundingClientRect();
  const own = host.getBoundingClientRect();
  if (!(pageBox.width > 0) || !(own.width > 0)) return false;
  // A box of several rows is no line: an author's name over the affiliation,
  // the last of a row of such blocks, ends where the page's text ends -- and
  // set again right-aligned with the affiliation typed longer, the name
  
  
  if (fontSizePixels !== null && fontSizePixels > 0 && own.height > fontSizePixels * 1.6) return false;
  const left = (own.left - pageBox.left) / pageBox.width;
  const right = (own.right - pageBox.left) / pageBox.width;
  // A label, not a line of a column of text (a two-column page's right
  // column ends its justified lines at the margin too).
  if (left <= 0.5 || right > 0.99 || right - left > 0.35 ||
    pdfNativeLineCentredOnPage({ x: left, width: right - left })) return false;
  const others = [...page.querySelectorAll<HTMLElement>('span[data-pdf-native-text-edit-box="true"]')]
    .filter((box) => box !== host && box.textContent?.trim())
    .map((box) => box.getBoundingClientRect())
    .filter((box) => box.width > 0 && box.height > 0)
    .map((box) => ({ left: (box.left - pageBox.left) / pageBox.width, right: (box.right - pageBox.left) / pageBox.width,
      rows: box.height / own.height }));
  if (!others.length) return false;
  // Two lines end within two points of each other (a page 612 pt wide; the
  // boxes are the glyphs' ink, whose last side bearings differ).
  const sameEnd = others.some((box) => Math.abs(box.right - right) < 0.003 && box.left > 0.5 && box.rows < 1.8);
  const textLeft = Math.min(...others.map((box) => box.left));
  const mirrored = Math.abs(right - (1 - textLeft)) < 0.01;
  return sameEnd || mirrored;
}
