/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeHexColor } from "@mrnote/core/web/lib/office-color-palette.ts";
import { presentationThemeColorMap } from "@mrnote/core/web/lib/presentation-office-background-styles.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import {
  presentationMasterColorMap,
  presentationMasterMappedColor,
  type PresentationMasterColorMapRecord,
} from "@mrnote/core/web/lib/presentation/master-color-map.ts";

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
export function presentationSlideInk(
  theme: DeckTheme | null | undefined,
  master?: PresentationMasterColorMapRecord | null,
): string {
  // A slide on a master with its own colour map (reading row 850, a deck
  
  // Text 1 through that master: the colour its own theme names, else the
  // slot its map names in the deck's slots.
  const mastersOwn = presentationMasterMappedColor(theme, master, "tx1");
  if (mastersOwn) return normalizeHexColor(mastersOwn) ?? mastersOwn;
  const masterMap = presentationMasterColorMap(theme, master);
  if (masterMap) theme = { ...theme, colorMap: masterMap };
  // An imported deck's own theme states its twelve slots on the import's theme
  // source only (as its Hyperlink slot, `presentationHyperlinkStyle`): its
  // Text 1 is that Dark 1, not the theme's body colour, which the import reads
  
  // 2007 theme's 1F497D blue; PowerPoint draws them black).
  // Text 1 is the slot the master's colour map names (`p:clrMap` tx1): Light 1
  // on a dark master (Berlin, a dark variant, an imported dark deck), where
  // read as Dark 1 a new text box wrote black words on the dark ground
  // (Windows PowerPoint draws them white, census50-work/berlin).
  const text1 = presentationThemeColorMap(theme).tx1;
  const identity = theme?.identity;
  const nativeText1 =
    identity?.selectionKind === "native" && (identity.variantId ?? "base") === "base"
      ? identity.source?.colorSlots?.[text1]
      : undefined;
  return (
    normalizeHexColor(theme?.colorSlots?.[text1]) ??
    normalizeHexColor(nativeText1) ??
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
