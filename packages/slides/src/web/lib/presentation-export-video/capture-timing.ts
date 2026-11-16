/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How PowerPoint's Create a Video lays a slide's transition on its frames.
 *
 * Measured on PowerPoint's own recording of the probe deck (slideshow-1110
 * truth/probe-slideshow.mov, 30 frames a second): a transition is drawn on
 * whole frames, its first at the start (nothing of the new slide yet) and
 * its last at its end, and the slide's timeline -- its first builds and the
 * time its After counts -- starts on the frame after that. A 10 ms Fade puts
 * the slide's first fade-in two frames (67 ms) after the transition starts
 * and keeps the slide up 4,067 ms for an After of 4 s; a slide with no
 * transition starts its timeline on its first frame.
 */

/** The frames a transition of `leadMs` holds the slide's timeline back, at `frameMs` a frame. */
export function captureTransitionLeadMs(leadMs: number, frameMs: number): number {
  if (!(leadMs > 0) || !(frameMs > 0)) return 0;
  // A transition that ends on a frame's time ends on that frame.
  const frames = Math.ceil(leadMs / frameMs - 1e-6);
  return (frames + 1) * frameMs;
}
