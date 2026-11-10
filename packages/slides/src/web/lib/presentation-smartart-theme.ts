/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
/**
 * The theme a slide's SmartArt is drawn through: the slide master's
 * colours and body face, and the style matrix its quick style's
 * `a:fillRef` / `a:effectRef` name -- Design > Variants > Effects, else the
 * file's own theme lists (`hostTheme`) while the master still wears the
 * file's theme. One reading for the window's conversion, rebuilds and
 * restyles and for the agent's SmartArt commands.
 */

import type { SmartArtColorTheme } from "@mrnote/core/web/lib/office-smartart/colors.ts";
import { nativeSmartArtHostTheme } from "@mrnote/core/web/lib/office-smartart/document-native-model.ts";
import {
  smartArtHostThemeSource,
  smartArtThemeEffectShadows,
  smartArtThemeFillStyles,
} from "@mrnote/core/web/lib/office-smartart/theme-effects.ts";
import { OFFICE_THEME_GRID_SLOTS } from "@mrnote/core/web/lib/office-theme-color-grid.ts";
import { officeThemeFontGalleryEntry } from "@mrnote/core/web/lib/office-theme-fonts.ts";
import type { mergeMasterTheme } from "@mrnote/core/web/lib/presentation-master.ts";
import { presentationThemeGridColors } from "@mrnote/core/web/lib/presentation-theme-color-binding.ts";

/** A diagram's file theme lists: the import's (`hostTheme`) or the one its conversion keeps. */
export type PresentationSmartArtHostTheme = { readonly effects?: string; readonly xml: string };

/**
 * The file's theme lists a converted diagram keeps (`smartArtHostTheme`, on
 * its first member), so a later Add Shape or style draws its gradients and
 * shadows as the conversion did; undefined for one that keeps none.
 */
export function presentationSmartArtMembersHostTheme(
  members: readonly unknown[],
): PresentationSmartArtHostTheme | undefined {
  for (const member of members) {
    const value = member && typeof member === "object" ? (member as Record<string, unknown>).smartArtHostTheme : undefined;
    const read = nativeSmartArtHostTheme(value);
    if (read) return read;
  }
  return undefined;
}

/**
 * The theme colours a SmartArt colour variation resolves against, from one
 * master's effective theme and its colour grid; null for a deck with no theme
 * of its own. The body face lines a diagram's words (a DengXian deck's labels
 * stand at DengXian's line, as PowerPoint fits them).
 */
export function presentationSmartArtColorTheme(
  effective: ReturnType<typeof mergeMasterTheme>,
  /** The diagram's file theme lists, as imported or kept by its conversion. */
  hostTheme?: PresentationSmartArtHostTheme | null,
  grid: ReturnType<typeof presentationThemeGridColors> = presentationThemeGridColors(effective),
): SmartArtColorTheme | null {
  if (!effective) return null;
  const bodyFace =
    effective.fontScheme?.minorLatin ??
    officeThemeFontGalleryEntry(effective.fontSchemeId)?.bodyLatin;
  return {
    dk1: grid.tx1,
    lt1: grid.bg1,
    dk2: grid.tx2,
    lt2: grid.bg2,
    accents: OFFICE_THEME_GRID_SLOTS.slice(4).map((slot) => grid[slot]),
    ...(bodyFace ? { bodyFace } : {}),
    // Design > Variants > Effects, else the deck's own theme as the diagram's
    // view read it while the deck still wears that theme: what a quick
    // style's `a:fillRef` / `a:effectRef` draw (3-D Polished's shadow and
    // rising ramp under the 2007 Office theme, `theme-effects`). A theme or
    // variant picked in Design is not the file's: its effects are its own.
    ...(() => {
      const identity = effective.identity;
      const fileTheme = !identity || identity.origin === "package" || identity.selectionKind === "native";
      const source = smartArtHostThemeSource(
        effective.effectSchemeId ? { id: effective.effectSchemeId } : null,
        fileTheme ? hostTheme : null,
      );
      const effectShadows = smartArtThemeEffectShadows(source);
      const fillStyles = smartArtThemeFillStyles(source);
      return { ...(effectShadows ? { effectShadows } : {}), ...(fillStyles ? { fillStyles } : {}) };
    })(),
  };
}
