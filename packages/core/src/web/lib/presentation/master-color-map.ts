/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeMasterColorMapFlipped } from "../presentation-office-background-styles";
import type {
  DECK_THEME_MAPPED_COLORS,
  DeckTheme,
  DeckThemeColorMap,
} from "../presentation-theme";

/** What a slide master states of its own colour map (reading row 850). */
export interface PresentationMasterColorMapRecord {
  colorMap?: DeckThemeColorMap;
  mappedColors?: Readonly<Record<(typeof DECK_THEME_MAPPED_COLORS)[number], string>>;
}

type MappedName = (typeof DECK_THEME_MAPPED_COLORS)[number];

function sameSlots(
  left: Readonly<Record<string, string | undefined>> | undefined,
  right: Readonly<Record<string, string | undefined>> | undefined,
): boolean {
  if (!left || !right) return false;
  const keys = new Set([...Object.keys(left), ...Object.keys(right)]);
  for (const key of keys) {
    if ((left[key] ?? "").toLowerCase() !== (right[key] ?? "").toLowerCase()) return false;
  }
  return true;
}

/**
 * Whether the deck's slides still sit on the file's own masters: the theme is
 * the import's (`native` and its base, or the Office theme the import
 * recognised, `origin: "package"`) and no Themes or Variants pick replaced its
 * colours. PowerPoint's Design > Themes and Variants put one master in for
 * every slide (its own colour map with it); Design > Variants > Colors only
 * recolours the theme, each master keeping its map.
 */
export function presentationThemeKeepsFileMasters(theme: DeckTheme | null | undefined): boolean {
  const identity = theme?.identity;
  if (!identity) return false;
  const own =
    (identity.selectionKind === "native" && (identity.variantId ?? "base") === "base") ||
    (identity.selectionKind === "office" && identity.origin === "package");
  if (!own) return false;
  // A Colors pick states its scheme; any other change of the slots is a
  // variant's (an Office variant writes its own twelve colours).
  if (theme?.colorSchemeId !== undefined) return true;
  return theme?.colorSlots === undefined || sameSlots(theme.colorSlots, identity.source?.colorSlots);
}

/**
 * The colour map a slide on `master` draws through, when the master states
 * its own (reading row 850) and the deck still sits on the file's masters;
 * undefined otherwise (the deck theme's map is the slide's).
 */
export function presentationMasterColorMap(
  theme: DeckTheme | null | undefined,
  master: PresentationMasterColorMapRecord | null | undefined,
): DeckThemeColorMap | undefined {
  if (!master?.colorMap || !master.mappedColors) return undefined;
  return presentationThemeKeepsFileMasters(theme) ? master.colorMap : undefined;
}

/**
 * The colour Background 1 / Text 1 / Background 2 / Text 2 is on a slide of
 * `master`, as that master's own theme states it, while nothing recoloured
 * the theme (a Colors pick, a Background Styles swap): then the slide's slots
 * are the deck theme's and its map the master's. Undefined when the master
 * states none.
 */
export function presentationMasterMappedColor(
  theme: DeckTheme | null | undefined,
  master: PresentationMasterColorMapRecord | null | undefined,
  name: MappedName,
): string | undefined {
  if (!presentationMasterColorMap(theme, master)) return undefined;
  if (theme?.colorSchemeId !== undefined || officeMasterColorMapFlipped(theme)) return undefined;
  return master?.mappedColors?.[name];
}
