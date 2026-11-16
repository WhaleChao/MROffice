/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PRESENTATION_FIRST_SLIDE_PAINTED_EVENT = "mrai:presentation-first-slide-painted";

export interface PresentationFirstSlidePainted {
  pageId: string;
  slideIndex: number;
}

/** Announce it once the fonts are in and a frame is drawn; the returned function cancels. */
export function announceFirstSlidePainted(
  detail: PresentationFirstSlidePainted,
  onSent?: () => void,
  win: Window = window,
): () => void {
  let active = true;
  let frame: number | null = null;
  const fonts = (win.document as Document & { fonts?: FontFaceSet }).fonts;
  const fontsReady: Promise<unknown> = fonts?.ready ?? Promise.resolve();
  const send = () => {
    frame = null;
    if (!active) return;
    active = false;
    onSent?.();
    win.dispatchEvent(new CustomEvent<PresentationFirstSlidePainted>(PRESENTATION_FIRST_SLIDE_PAINTED_EVENT, { detail }));
  };
  void fontsReady.then(
    () => {
      // The frame after the fonts: the slide as it now stays.
      if (active) frame = win.requestAnimationFrame(() => (frame = win.requestAnimationFrame(send)));
    },
    () => {
      if (active) frame = win.requestAnimationFrame(send);
    },
  );
  return () => {
    active = false;
    if (frame !== null) win.cancelAnimationFrame(frame);
  };
}
