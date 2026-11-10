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
 * line keeps its middle.
 */
export function pdfNativeLineRightFlushOnPage(host: Element | null | undefined): boolean {
  const page = host?.closest(".pdf-viewer__page");
  if (!host || !page) return false;
  const pageBox = page.getBoundingClientRect();
  const own = host.getBoundingClientRect();
  if (!(pageBox.width > 0) || !(own.width > 0)) return false;
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
