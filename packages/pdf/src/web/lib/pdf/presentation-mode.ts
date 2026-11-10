/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation (fullscreen reading) mode state transitions for the PDF reader.
 *
 * Mirrors pdf.js `web/pdf_presentation_mode.js` (Apache-2.0): entering
 * fullscreen saves the current view state and switches to a one-page-per-screen
 * layout; leaving restores exactly what the reader had before. jsdom cannot run
 * a real Fullscreen API, so the *decisions* live here as pure snapshot →
 * snapshot functions — the component only wires them to `requestFullscreen()`
 * and the `fullscreenchange` event.
 */

import {
  PDF_FIT_MODES,
  clampScale,
  isPdfViewMode,
  PDF_DEFAULT_VIEW_MODE,
  type PdfFitMode,
  type PdfViewMode,
} from "./view-modes";

/** The ephemeral view state presentation mode saves and restores. */
export interface PdfViewStateSnapshot {
  viewMode: PdfViewMode;
  fitMode: PdfFitMode | null;
  scale: number;
}

/** What to apply when ENTERING presentation mode, plus the state to stash. */
export interface PdfPresentationEnter {
  /** Sanitized copy of the pre-presentation state, for the exit restore. */
  saved: PdfViewStateSnapshot;
  /** One page per screen, fitted whole — pdf.js's page-fit presentation. */
  apply: { viewMode: PdfViewMode; fitMode: PdfFitMode };
}

/** What to apply when EXITING presentation mode. */
export interface PdfPresentationExit {
  viewMode: PdfViewMode;
  fitMode: PdfFitMode | null;
  scale: number;
  /**
   * True when a sticky fit mode was armed before entering: the restored scale
   * must be RECOMPUTED against the post-fullscreen container size (the saved
   * numeric scale fit the old window, which fullscreen resized twice since).
   * False → apply the saved numeric scale directly.
   */
  refit: boolean;
}

/** Defensive copy: unknown junk degrades to the reader defaults. */
function sanitizeSnapshot(state: PdfViewStateSnapshot): PdfViewStateSnapshot {
  return {
    viewMode: isPdfViewMode(state?.viewMode)
      ? state.viewMode
      : PDF_DEFAULT_VIEW_MODE,
    fitMode: (PDF_FIT_MODES as readonly unknown[]).includes(state?.fitMode)
      ? (state.fitMode as PdfFitMode)
      : null,
    scale: clampScale(state?.scale),
  };
}

/**
 * The transition for entering presentation mode: stash the current view state
 * and switch to single-page + fit-page (both routed through the existing view
 * mode / sticky-fit paths, which then handle the fullscreen resize).
 */
export function presentationEnterTransition(
  current: PdfViewStateSnapshot,
): PdfPresentationEnter {
  return {
    saved: sanitizeSnapshot(current),
    apply: { viewMode: "single", fitMode: "page" },
  };
}

/**
 * The transition for exiting presentation mode: restore the saved snapshot.
 * `null` saved state (exit without a recorded entry — e.g. the component
 * remounted while fullscreen) → null: nothing to restore, keep whatever the
 * reader currently shows rather than jumping to fabricated defaults.
 */
export function presentationExitTransition(
  saved: PdfViewStateSnapshot | null,
): PdfPresentationExit | null {
  if (!saved) return null;
  const clean = sanitizeSnapshot(saved);
  return {
    viewMode: clean.viewMode,
    fitMode: clean.fitMode,
    scale: clean.scale,
    refit: clean.fitMode != null,
  };
}
