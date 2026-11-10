/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The masters of an Office theme deck after Edit Master ▸ Insert Slide Master
 * (PPT-B030), as the one-master edits of Slide Master view address them.
 *
 * Master one is the theme's own (`theme.officeMaster`); every further master
 * is an entry of `theme.officeMaster.additionalMasters` -- Office's default
 * theme with Slide Master view's edits in the same keys. Each Slide Master
 * view transform (lib/presentation-office-master*, -layout-edit) is written
 * for one master; this adapter runs it on an inserted master by handing it
 * that master as a deck theme of its own (the default Office theme with the
 * entry as its `officeMaster`) and writing the result back into the entry.
 * Slides on an inserted master's layouts carry `<key>--<layout>`; the
 * adapter hands the transform those slides with the master's own layout ids
 * and puts the refs back.
 *
 * Dependency-free beyond the theme model, so every master module can use it.
 */

import {
  OFFICE_INSERTED_MASTER_THEME_ID,
  OFFICE_INSERTED_MASTER_VARIANT_ID,
  OFFICE_PRIMARY_MASTER_KEY,
  normalizeDeckOfficeMaster,
  officeMasterLayoutRef,
  parseOfficeMasterLayoutRef,
  type DeckOfficeMaster,
  type OfficeAdditionalMaster,
} from "@mrnote/core/web/lib/presentation-office-master-model.ts";
import { applyOfficeThemeSelection, normalizeTheme, type DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

/** The masters Insert Slide Master added, in thumbnail order. */
export function officeAdditionalMasters(theme: DeckTheme | null | undefined): OfficeAdditionalMaster[] {
  return normalizeDeckOfficeMaster(theme?.officeMaster)?.additionalMasters ?? [];
}

/**
 * An inserted master as a deck theme of its own: Office's default theme, the
 * entry's edits as its `officeMaster`. Null when the deck has no such master.
 */
export function officeAdditionalMasterTheme(
  theme: DeckTheme | null | undefined,
  masterKey: string,
): DeckTheme | null {
  const entry = officeAdditionalMasters(theme).find((candidate) => candidate.key === masterKey);
  if (!entry) return null;
  const base = applyOfficeThemeSelection(undefined, OFFICE_INSERTED_MASTER_THEME_ID, OFFICE_INSERTED_MASTER_VARIANT_ID);
  if (!base) return null;
  // The entry's key is the deck's address of the master, not one of its edits.
  const edits: Partial<OfficeAdditionalMaster> = { ...entry };
  delete edits.key;
  const officeMaster = normalizeDeckOfficeMaster(edits);
  return normalizeTheme({ ...base, ...(officeMaster ? { officeMaster } : {}) }) ?? null;
}

/** The theme a master's Slide Master view card draws and edits under: the deck's, or an inserted master's own. */
export function officeMasterTheme(theme: DeckTheme | null | undefined, masterKey: string): DeckTheme | null {
  if (masterKey === OFFICE_PRIMARY_MASTER_KEY) return theme ?? null;
  return officeAdditionalMasterTheme(theme, masterKey);
}

/**
 * The deck theme with an inserted master's edits taken from `next` (its
 * one-master theme after a transform). Null when `next` is null (a no-op)
 * or the master is gone.
 */
export function themeWithOfficeAdditionalMaster(
  theme: DeckTheme | null | undefined,
  masterKey: string,
  next: DeckTheme | null,
): DeckTheme | null {
  if (!theme || !next) return null;
  const master: DeckOfficeMaster = { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
  const masters = [...(master.additionalMasters ?? [])];
  const index = masters.findIndex((entry) => entry.key === masterKey);
  if (index < 0) return null;
  const edits = normalizeDeckOfficeMaster(next.officeMaster) ?? {};
  delete edits.additionalMasters;
  delete edits.background;
  delete edits.backgroundStyle;
  masters[index] = { key: masterKey, ...edits };
  master.additionalMasters = masters;
  const result = normalizeTheme({ ...theme, officeMaster: normalizeDeckOfficeMaster(master) });
  return result && JSON.stringify(result) !== JSON.stringify(normalizeTheme(theme)) ? result : null;
}

/**
 * A layout ref of an inserted master the deck has, split into its master's
 * theme and the master's own layout id; null for a layout of master one.
 */
export function routeOfficeLayout(
  theme: DeckTheme | null | undefined,
  layoutRef: unknown,
): { masterKey: string; theme: DeckTheme; layoutId: string } | null {
  const ref = parseOfficeMasterLayoutRef(layoutRef);
  if (!ref) return null;
  const masterTheme = officeAdditionalMasterTheme(theme, ref.masterKey);
  return masterTheme ? { masterKey: ref.masterKey, theme: masterTheme, layoutId: ref.layoutId } : null;
}

/** Which master a slide is on: an inserted master's key, else master one's. */
export function officeSlideMasterKey(
  slide: { layout?: unknown },
  theme?: DeckTheme | null,
): string {
  const ref = parseOfficeMasterLayoutRef(slide.layout);
  if (!ref) return OFFICE_PRIMARY_MASTER_KEY;
  // A slide whose master is gone falls back to master one's layout.
  if (theme && !officeAdditionalMasters(theme).some((entry) => entry.key === ref.masterKey)) {
    return OFFICE_PRIMARY_MASTER_KEY;
  }
  return ref.masterKey;
}

/**
 * The slides on one inserted master, with their layouts as the master's own
 * ids, and where each sits in the deck. Slides on another master are not
 * handed to the master's transform, so a layout id of one master never
 * matches a slide of another.
 */
export function officeMasterLocalSlides<S extends { layout?: unknown }>(
  slides: readonly S[],
  masterKey: string,
): { slides: S[]; positions: number[] } {
  const local: S[] = [];
  const positions: number[] = [];
  slides.forEach((slide, position) => {
    const ref = parseOfficeMasterLayoutRef(slide.layout);
    if (ref?.masterKey !== masterKey) return;
    local.push({ ...slide, layout: ref.layoutId });
    positions.push(position);
  });
  return { slides: local, positions };
}

/** The deck's slides with a master's transformed local slides put back, their refs restored. */
export function slidesFromOfficeMasterLocal<S extends { layout?: unknown }>(
  slides: readonly S[],
  masterKey: string,
  local: { slides: readonly S[]; positions: readonly number[] },
  edited: readonly S[],
): S[] {
  const next = [...slides];
  local.positions.forEach((position, index) => {
    const slide = edited[index];
    // An untouched slide stays the deck's own object.
    if (!slide || slide === local.slides[index]) return;
    next[position] = {
      ...slide,
      layout: officeMasterLayoutRef(masterKey, String(slide.layout ?? "blank")),
    };
  });
  return next;
}
