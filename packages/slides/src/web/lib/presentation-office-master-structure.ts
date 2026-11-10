/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide Master view's remaining structure commands on an Office theme deck
 * (PPT-B030), as pure transforms of the deck theme:
 *
 * - Background ▸ Hide Background Graphics on a layout: the layout's slides
 *   do not draw the master's graphics (`p:sldLayout@showMasterSp="0"`;
 *   Microsoft Support, "Hide background graphics": in Slide Master view,
 *   select the layout, then Hide Background Graphics).
 * - Master Layout: the master's own placeholders -- title, text, date, slide
 *   number, footer -- deleted in Slide Master view (select one, Delete), and
 *   the dialog's check boxes put a deleted one back. PowerPoint's dialog
 *   shows a placeholder the master still has as checked and unavailable; it
 *   only restores (Microsoft Support, "Restore placeholders on the slide
 *   master").
 * - Edit Master ▸ Insert Layout: a new layout under the master, "Custom
 *   Layout" with a title and the footers (Microsoft Support, "Apply a slide
 *   layout": Insert Layout adds a layout with title and footer
 *   placeholders). It is stored as a `custom-<letters>` id in
 *   `officeMaster.customLayouts` and takes Title Only's frames, so every
 *   other Slide Master view command (Insert Placeholder, Title, Footers,
 *   Rename, Delete, Hide Background Graphics) works on it unchanged.
 *
 * Both live in `theme.officeMaster` (lib/presentation-office-master-model)
 * and the server writes them into the download's master and layout parts.
 */

import {
  OFFICE_MASTER_MAX_ADDITIONAL,
  OFFICE_MASTER_MAX_CUSTOM_LAYOUTS,
  OFFICE_PRIMARY_MASTER_KEY,
  nextOfficeMasterKey,
  normalizeOfficeMasterName,
  officeMasterLayoutRef,
  OFFICE_MASTER_PLACEHOLDER_FAMILIES,
  nextOfficeCustomLayoutId,
  normalizeDeckOfficeMaster,
  type DeckOfficeMaster,
  type OfficeMasterPlaceholderFamily,
} from "@mrnote/core/web/lib/presentation-office-master-model.ts";
import { officeMasterLayoutIds, officeMasterSelection } from "./presentation-office-master";
import {
  officeAdditionalMasters,
  officeMasterTheme,
  officeSlideMasterKey,
  routeOfficeLayout,
  themeWithOfficeAdditionalMaster,
} from "./presentation-office-master-routing";
import { normalizeTheme, type DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

/** Run a one-master edit on master `masterKey` (master one, or an inserted master). */
function onOfficeMaster(
  theme: DeckTheme | null | undefined,
  masterKey: string | undefined,
  edit: (masterTheme: DeckTheme | null | undefined) => DeckTheme | null,
): DeckTheme | null {
  if (!masterKey || masterKey === OFFICE_PRIMARY_MASTER_KEY) return edit(theme);
  return themeWithOfficeAdditionalMaster(theme, masterKey, edit(officeMasterTheme(theme, masterKey)));
}

function currentMaster(theme: DeckTheme): DeckOfficeMaster {
  return { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
}

function withMaster(theme: DeckTheme, master: DeckOfficeMaster): DeckTheme | null {
  const next: DeckTheme = { ...theme };
  const normalized = normalizeDeckOfficeMaster(master);
  if (normalized) next.officeMaster = normalized;
  else delete next.officeMaster;
  const result = normalizeTheme(next);
  return result && JSON.stringify(result) !== JSON.stringify(normalizeTheme(theme)) ? result : null;
}

/** Whether a slide on `layoutId` hides the master's graphics. */
export function officeLayoutHidesBackgroundGraphics(
  theme: DeckTheme | null | undefined,
  layoutId: unknown,
): boolean {
  if (typeof layoutId !== "string" || !officeMasterSelection(theme)) return false;
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) return officeLayoutHidesBackgroundGraphics(routed.theme, routed.layoutId);
  const id = layoutId === "titleSlide" ? "title" : layoutId;
  return (normalizeDeckOfficeMaster(theme?.officeMaster)?.layoutsHidingBackgroundGraphics ?? []).includes(id);
}

/** Background ▸ Hide Background Graphics on the layout. Null for a no-op. */
export function themeWithOfficeLayoutBackgroundGraphics(
  theme: DeckTheme | null | undefined,
  layoutId: string,
  hide: boolean,
): DeckTheme | null {
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) {
    return themeWithOfficeAdditionalMaster(
      theme,
      routed.masterKey,
      themeWithOfficeLayoutBackgroundGraphics(routed.theme, routed.layoutId, hide),
    );
  }
  if (!theme || !officeMasterSelection(theme)) return null;
  const master = currentMaster(theme);
  const hiding = new Set(master.layoutsHidingBackgroundGraphics ?? []);
  if (hide) hiding.add(layoutId);
  else hiding.delete(layoutId);
  master.layoutsHidingBackgroundGraphics = [...hiding];
  if (master.layoutsHidingBackgroundGraphics.length === 0) delete master.layoutsHidingBackgroundGraphics;
  return withMaster(theme, master);
}

/** The master's placeholders Slide Master view deleted. */
export function officeMasterPlaceholdersOff(
  theme: DeckTheme | null | undefined,
  masterKey?: string,
): ReadonlySet<OfficeMasterPlaceholderFamily> {
  const masterTheme = masterKey ? officeMasterTheme(theme, masterKey) : theme;
  return new Set(normalizeDeckOfficeMaster(masterTheme?.officeMaster)?.masterPlaceholdersOff ?? []);
}

/**
 * Delete (`on: false`) or restore (`on: true`) one of the master's own
 * placeholders. Null for a no-op or a deck without an Office master.
 */
export function themeWithOfficeMasterPlaceholder(
  theme: DeckTheme | null | undefined,
  family: OfficeMasterPlaceholderFamily,
  on: boolean,
  masterKey?: string,
): DeckTheme | null {
  if (masterKey && masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    return onOfficeMaster(theme, masterKey, (masterTheme) =>
      themeWithOfficeMasterPlaceholder(masterTheme, family, on),
    );
  }
  if (!theme || !officeMasterSelection(theme)) return null;
  const master = currentMaster(theme);
  const off = new Set(master.masterPlaceholdersOff ?? []);
  if (on) off.delete(family);
  else off.add(family);
  master.masterPlaceholdersOff = OFFICE_MASTER_PLACEHOLDER_FAMILIES.filter((candidate) => off.has(candidate));
  if (master.masterPlaceholdersOff.length === 0) delete master.masterPlaceholdersOff;
  return withMaster(theme, master);
}

/** The master card's block role → the placeholder family it is. */
export function officeMasterPlaceholderFamilyForRole(role: unknown): OfficeMasterPlaceholderFamily | null {
  if (role === "title") return "title";
  if (role === "body" || role === "content" || role === "subtitle") return "body";
  return null;
}

/**
 * Edit Master ▸ Insert Layout: a custom layout right after `afterLayoutId`
 * (the selected layout, as PowerPoint inserts it; after the last layout when
 * the master or nothing is selected), named `name` (PowerPoint's "Custom
 * Layout"). Null when the deck has no Office master or holds the maximum.
 */
export function themeWithInsertedOfficeLayout(
  theme: DeckTheme | null | undefined,
  name: string,
  afterLayoutId?: string | null,
  masterKey?: string,
): { theme: DeckTheme; layoutId: string } | null {
  // On an inserted master (its card or one of its layouts selected): the
  // layout joins that master (PPT-B030).
  const routed = afterLayoutId ? routeOfficeLayout(theme, afterLayoutId) : null;
  const target = routed?.masterKey ?? (afterLayoutId ? undefined : masterKey);
  if (target && target !== OFFICE_PRIMARY_MASTER_KEY) {
    const inserted = themeWithInsertedOfficeLayout(officeMasterTheme(theme, target), name, routed?.layoutId ?? null);
    const next = inserted ? themeWithOfficeAdditionalMaster(theme, target, inserted.theme) : null;
    return inserted && next ? { theme: next, layoutId: officeMasterLayoutRef(target, inserted.layoutId) } : null;
  }
  if (!theme || !officeMasterSelection(theme)) return null;
  const master = currentMaster(theme);
  const custom = master.customLayouts ?? [];
  if (custom.length >= OFFICE_MASTER_MAX_CUSTOM_LAYOUTS) return null;
  const layoutId = nextOfficeCustomLayoutId(custom);
  const order: string[] = officeMasterLayoutIds(theme);
  const anchor = afterLayoutId ? order.indexOf(afterLayoutId) : -1;
  order.splice(anchor >= 0 ? anchor + 1 : order.length, 0, layoutId);
  master.customLayouts = [...custom, layoutId];
  // Only an insert before the end needs an order of its own.
  if (anchor >= 0 && anchor + 1 < order.length - 1) master.layoutOrder = order;
  else if (master.layoutOrder) master.layoutOrder = order;
  const trimmed = name.trim();
  if (trimmed) master.layoutNames = { ...(master.layoutNames ?? {}), [layoutId]: trimmed };
  const next = withMaster(theme, master);
  return next ? { theme: next, layoutId } : null;
}

/**
 * Edit Master ▸ Insert Slide Master (PPT-B030): a master after every other,
 * as PowerPoint inserts it -- Office's default theme "without theme colors or
 * effects" with its eleven standard layouts (Microsoft Support, "Use multiple
 * slide masters in one presentation"), named `name` (PowerPoint's "Custom
 * Design", the new theme's name) and preserved, as PowerPoint writes it
 * (`p:sldMaster@preserve="1"`). Null on a deck without an Office master or at
 * the maximum.
 */
export function themeWithInsertedOfficeSlideMaster(
  theme: DeckTheme | null | undefined,
  name: string,
): { theme: DeckTheme; masterKey: string } | null {
  if (!theme || !officeMasterSelection(theme)) return null;
  const master = currentMaster(theme);
  const masters = [...(master.additionalMasters ?? [])];
  if (masters.length >= OFFICE_MASTER_MAX_ADDITIONAL) return null;
  const masterKey = nextOfficeMasterKey(masters.map((entry) => entry.key));
  const trimmed = normalizeOfficeMasterName(name);
  masters.push({ key: masterKey, ...(trimmed ? { name: trimmed } : {}), preserve: true });
  master.additionalMasters = masters;
  const next = withMaster(theme, master);
  return next ? { theme: next, masterKey } : null;
}

/** The masters of the deck in thumbnail order, by key: master one, then the inserted ones. */
export function officeSlideMasterKeys(theme: DeckTheme | null | undefined): string[] {
  if (!officeMasterSelection(theme)) return [];
  return [OFFICE_PRIMARY_MASTER_KEY, ...officeAdditionalMasters(theme).map((entry) => entry.key)];
}

/**
 * Why Edit Master ▸ Delete cannot remove the master, or null when it can:
 * PowerPoint deletes a master only while another remains and no slide uses
 * any of its layouts.
 */
export function officeSlideMasterDeleteBlocker(
  theme: DeckTheme | null | undefined,
  slides: ReadonlyArray<{ layout?: unknown }>,
  masterKey: string,
): "only" | "in-use" | "unavailable" | null {
  const keys = officeSlideMasterKeys(theme);
  if (!keys.includes(masterKey)) return "unavailable";
  if (keys.length <= 1) return "only";
  if (slides.some((slide) => officeSlideMasterKey(slide, theme) === masterKey)) return "in-use";
  return null;
}

/**
 * Edit Master ▸ Delete on a master no slide uses. An inserted master is
 * simply gone. Master one is the deck theme's own: the next master takes
 * its place -- the deck becomes Office's default theme with that master's
 * edits -- and its slides' layout refs become master one's. Null when the
 * master cannot go.
 */
export function deckWithDeletedOfficeSlideMaster<S extends { layout?: unknown }>(
  theme: DeckTheme | null | undefined,
  slides: readonly S[],
  masterKey: string,
): { theme: DeckTheme; slides: S[] } | null {
  if (!theme || officeSlideMasterDeleteBlocker(theme, slides, masterKey) !== null) return null;
  const master = currentMaster(theme);
  const masters = master.additionalMasters ?? [];
  if (masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    master.additionalMasters = masters.filter((entry) => entry.key !== masterKey);
    if (master.additionalMasters.length === 0) delete master.additionalMasters;
    const next = withMaster(theme, master);
    return next ? { theme: next, slides: [...slides] } : null;
  }
  const [promoted, ...rest] = masters;
  if (!promoted) return null;
  const promotedTheme = officeMasterTheme(theme, promoted.key);
  if (!promotedTheme) return null;
  const promotedMaster = { ...(normalizeDeckOfficeMaster(promotedTheme.officeMaster) ?? {}) };
  if (rest.length > 0) promotedMaster.additionalMasters = rest;
  const nextTheme = normalizeTheme({ ...promotedTheme, officeMaster: normalizeDeckOfficeMaster(promotedMaster) });
  if (!nextTheme) return null;
  const prefix = `${promoted.key}--`;
  const nextSlides = slides.map((slide) =>
    typeof slide.layout === "string" && slide.layout.startsWith(prefix)
      ? { ...slide, layout: slide.layout.slice(prefix.length) }
      : slide,
  );
  return { theme: nextTheme, slides: nextSlides };
}

/** Rename Master's and Preserve's master: the selected master card's key, or the master of the selected layout. */
export function officeSlideMasterKeyOf(cardKey: string): string {
  if (cardKey === OFFICE_PRIMARY_MASTER_KEY) return cardKey;
  const routedKey = /^(master-[a-z]{1,2})(?:--.*)?$/.exec(cardKey)?.[1];
  return routedKey ?? OFFICE_PRIMARY_MASTER_KEY;
}

export { onOfficeMaster };
