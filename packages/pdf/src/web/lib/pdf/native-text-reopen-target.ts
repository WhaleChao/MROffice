/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { findPdfNativeTextBoxByLineId } from "./native-text-line-lineage";

/** The box a line is opened again on, for words carried back into it
 * (PdfPageWindow's reopenNativeTextLineAtRef), and how it is clicked.
 *
 * `found` is the box lying topmost at the line's point. Where lines overlap
 * -- a scan's recognised text, whose frames run into the next line's --
 * that may be another line's box, and a click at the point, even on the
 * line's own box, is taken by the viewer for the box its geometry finds
 * there (the smallest frame holding the point: handleNativeTextEditBoxClick):
 * another line opened, the words waited two minutes for a line that never
 * opened and were kept as a draft, the page unsaved for good (pdf-walk-1000
 * rx0516: "...possono. ra" opened again as "RA telefono", ru0516c after the
 * first fix: the line's own box clicked, "RA" opened).
 * The line the words are for (`lineId`, its lineage id) is opened where the
 * page has it (`ownLine`) by a click with no pointer: on the box itself,
 * which the viewer takes as it is sent. */
export function pdfNativeTextReopenTarget(
  pageHost: ParentNode | null | undefined,
  found: HTMLElement | null,
  lineId: string | null | undefined,
): Readonly<{ box: HTMLElement | null; ownLine: boolean }> {
  const named = lineId ? findPdfNativeTextBoxByLineId(pageHost, lineId) : null;
  if (named) return { box: named, ownLine: true };
  return { box: found, ownLine: false };
}
