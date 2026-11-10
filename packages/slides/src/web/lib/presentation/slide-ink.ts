/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeHexColor } from "@mrnote/core/web/lib/office-color-palette.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

/**
 * The colour of words that state none (no run, block, master or theme colour
 * reaches them): a new text box's words, a table cell's, an equation's, a
 * placeholder's -- the theme's text colour (Text 1, `tx1`: the theme's Dark 1,
 * black in the Office theme), as PowerPoint draws them.
 *
 * The slide states it itself, on every surface that draws it -- the editor
 * canvas, the slide show, the presenter's slides, Reading View, the
 * thumbnails. Inherited from the page instead, each surface drew its own
 * chrome's ink: the canvas the deck window's grey, the show the console's
 * teal (and the canvas a light grey on the white slide in dark mode).
 */
export function presentationSlideInk(theme: DeckTheme | null | undefined): string {
  // An imported deck's own theme states its twelve slots on the import's theme
  // source only (as its Hyperlink slot, `presentationHyperlinkStyle`): its
  // Text 1 is that Dark 1, not the theme's body colour, which the import reads
  
  // 2007 theme's 1F497D blue; PowerPoint draws them black).
  const identity = theme?.identity;
  const nativeDark1 =
    identity?.selectionKind === "native" && (identity.variantId ?? "base") === "base"
      ? identity.source?.colorSlots?.dk1
      : undefined;
  return (
    normalizeHexColor(theme?.colorSlots?.dk1) ??
    normalizeHexColor(nativeDark1) ??
    normalizeHexColor(theme?.bodyTextColor) ??
    PRESENTATION_SLIDE_INK
  );
}

/** Text 1 of a deck that states no theme: the Office theme's black. */
export const PRESENTATION_SLIDE_INK = "#000000";

/** A layout placeholder's words that state no colour: the same Text 1. */
export const PRESENTATION_PLACEHOLDER_INK = PRESENTATION_SLIDE_INK;

/**
 * The custom property the canvas's placeholder rule reads, so the stylesheet
 * and the show draw one value.
 */
export const PRESENTATION_PLACEHOLDER_INK_VAR = "--presentation-placeholder-ink";

/** Whether a block is a layout placeholder (the canvas marks it `data-placeholder-role`). */
export function presentationBlockIsPlaceholder(block: unknown): boolean {
  return (
    !!block &&
    typeof block === "object" &&
    typeof (block as { role?: unknown }).role === "string"
  );
}
