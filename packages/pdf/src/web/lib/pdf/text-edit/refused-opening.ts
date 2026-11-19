/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextSelectionEvidence } from "../engine/native/list-text-fidelity";

/**
 * A box whose opening was refused for good, kept for the few seconds in
 * which the window may open a part of it in its place: the viewer opens the
 * run under the press of a paragraph refused whole. That opening continues
 * the refused one -- the presses held for it are held for the run, and
 * answered by the run's proof (PdfNativeTextEditSessionHost's `continues`).
 */
export type PdfNativeRefusedTextOpening = Readonly<{
  /** The session key of the editor the refused box opened in. */
  sessionKey: string;
  pageNumber: number;
  sourceSha256: string;
  /** Where its runs stand on the page (PDF units), a line's height around them. */
  bounds: Readonly<{ left: number; right: number; bottom: number; top: number }>;
  /** When it was refused (ms). */
  at: number;
}>;

/** How long after a refusal an opening may still continue it. */
export const PDF_NATIVE_REFUSED_OPENING_CONTINUES_MS = 5_000;

type Selection = Pick<PdfNativeListTextSelectionEvidence, "pageNumber" | "sourceSha256" | "item" | "items">;

const runsOf = (selection: Selection) => (selection.items?.length ? selection.items : [selection.item]);

/** The refused opening of `selection`, as it is kept; null where its runs
 * state no place. By where the runs stand, not by their item numbers: a
 * paragraph's items may each be several of the page's, and the run that
 * opens in its place is one of those. */
export function pdfNativeRefusedTextOpening(sessionKey: string, selection: Selection, at: number): PdfNativeRefusedTextOpening | null {
  const frames = runsOf(selection).flatMap((item) => {
    const [a, b, , , x, y] = item.transform;
    if (![a, b, x, y, item.width, item.height].every(Number.isFinite)) return [];
    const along = Math.hypot(a, b) || 1, reach = Math.abs(item.height) || 1;
    const [endX, endY] = [x + (a / along) * item.width, y + (b / along) * item.width];
    return [[Math.min(x, endX) - reach, Math.max(x, endX) + reach, Math.min(y, endY) - reach, Math.max(y, endY) + reach]];
  });
  if (!frames.length) return null;
  return Object.freeze({
    sessionKey, pageNumber: selection.pageNumber, sourceSha256: selection.sourceSha256, at,
    bounds: Object.freeze({ left: Math.min(...frames.map((frame) => frame[0])), right: Math.max(...frames.map((frame) => frame[1])),
      bottom: Math.min(...frames.map((frame) => frame[2])), top: Math.max(...frames.map((frame) => frame[3])) }),
  });
}

/**
 * The session key of the refused opening that an opening of `selection`
 * continues, or undefined: the same page of the same bytes, within
 * PDF_NATIVE_REFUSED_OPENING_CONTINUES_MS of the refusal, every run of it
 * standing where the refused box's runs stand. Another box opened meanwhile
 * (a line elsewhere on the page) continues nothing.
 */
export function pdfNativeTextOpeningContinues(refused: PdfNativeRefusedTextOpening | null, selection: Selection,
  now: number): string | undefined {
  if (!refused || refused.pageNumber !== selection.pageNumber || refused.sourceSha256 !== selection.sourceSha256 ||
    now - refused.at < 0 || now - refused.at >= PDF_NATIVE_REFUSED_OPENING_CONTINUES_MS) return undefined;
  const inside = runsOf(selection).every((item) => {
    const [, , , , x, y] = item.transform;
    return x >= refused.bounds.left && x <= refused.bounds.right && y >= refused.bounds.bottom && y <= refused.bounds.top;
  });
  return inside ? refused.sessionKey : undefined;
}
