/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { presentationThemeColorMap } from "@mrnote/core/web/lib/presentation-office-background-styles.ts";
import type { TableStyleSchemeColors } from "@mrnote/core/web/lib/presentation-table-web-gallery.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

/**
 * The deck's own colours a table style resolves besides its accents
 * (`TableStyleSchemeColors`): the theme's Dark 1 / Light 1 / Dark 2 / Light 2
 * and link slots -- an applied Office theme's, else an imported deck's own
 * (stated on its import's theme source, as the slide's ink reads them) -- and
 * the master's colour map (`presentationThemeColorMap`, reading row 826).
 * Undefined for a deck that states neither: the gallery then resolves in
 * Office's own theme, as before.
 */
export function presentationTableStyleSchemeColors(
  theme: DeckTheme | null | undefined,
): TableStyleSchemeColors | undefined {
  if (!theme) return undefined;
  const identity = theme.identity;
  const nativeSlots =
    identity?.selectionKind === "native" && (identity.variantId ?? "base") === "base"
      ? identity.source?.colorSlots
      : undefined;
  const source = (theme.colorSlots ?? nativeSlots) as Readonly<Record<string, unknown>> | undefined;
  const slots: Record<string, string> = {};
  for (const slot of ["dk1", "lt1", "dk2", "lt2", "hlink", "folHlink"] as const) {
    const color = source?.[slot];
    if (typeof color === "string" && /^#[0-9a-f]{6}$/i.test(color)) slots[slot] = color.toLowerCase();
  }
  const map = presentationThemeColorMap(theme);
  const light = map.tx1 === "dk1" && map.bg1 === "lt1" && map.tx2 === "dk2" && map.bg2 === "lt2";
  if (!Object.keys(slots).length && light) return undefined;
  return {
    ...(Object.keys(slots).length ? { slots } : {}),
    ...(light ? {} : { map: { tx1: map.tx1, tx2: map.tx2, bg1: map.bg1, bg2: map.bg2 } }),
  };
}
