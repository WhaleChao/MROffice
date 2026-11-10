/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Line-number labels as the page can hold them (user audit PDF 0041). The
 * labels were 10 px at every zoom, 14 px high with their padding: at 57 %
 * fourteen pairs of neighbouring numbers lay over each other. Like Word's
 * line numbers they now scale with the page (never below a legible 7 px),
 * and a label that would still overlap the one above it is left out, as
 * Word's "Count by" shows every n-th number of dense text.
 */
export interface PdfLineNumberLine {
  readonly number: number;
  /** Line centre, percent of the page height. */
  readonly topPercent: number;
}

/** The label font at a zoom: 10 px at 100 %, between 7 and 12 px. */
export function pdfLineNumberFontSize(scale: number): number {
  const size = 10 * (Number.isFinite(scale) && scale > 0 ? scale : 1);
  return Math.round(Math.min(12, Math.max(7, size)) * 10) / 10;
}

/** The labels to draw, top to bottom, none closer than a label's height. */
export function pdfLineNumberLabels<T extends PdfLineNumberLine>(
  lines: readonly T[],
  pageHeightPx: number,
  fontSize: number,
): T[] {
  if (!(pageHeightPx > 0)) return [...lines];
  // Line height 1.2 plus the 1 px padding above and below.
  const labelHeight = fontSize * 1.2 + 2;
  const shown: T[] = [];
  let lastTop = -Infinity;
  for (const line of [...lines].sort((left, right) => left.topPercent - right.topPercent)) {
    const top = (line.topPercent / 100) * pageHeightPx;
    if (top - lastTop < labelHeight) continue;
    shown.push(line);
    lastTop = top;
  }
  return shown;
}
