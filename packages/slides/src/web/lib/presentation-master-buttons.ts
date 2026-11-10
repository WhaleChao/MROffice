/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide Master view edits on the deck's `masters`, as pure transforms: a
 * master's name, its title / body text styles and other theme overrides
 * (`master.theme`, layered over the deck theme by `mergeMasterTheme`), and
 * its child layouts (Insert / Rename / Delete Layout). The window's Slide
 * Master pane and an editor command both commit what these return.
 */

import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import type { LayoutId } from "@mrnote/core/web/lib/presentation-layouts.ts";
import {
  MAX_MASTER_LAYOUTS,
  normalizeMasters,
  type SlideMaster,
  type SlideMasterLayout,
} from "@mrnote/core/web/lib/presentation-master.ts";
import { normalizeTheme, type DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

/**
 * A slide with its native master/layout references dropped (keys removed,
 * not set to `undefined`, so the persisted slide stays byte-clean): what a
 * slide becomes when its master is deleted or the assignment is cleared.
 */
export function slideWithoutMaster<S extends { masterId?: unknown; layoutId?: unknown }>(slide: S): S {
  if (slide.masterId === undefined && slide.layoutId === undefined) return slide;
  const next = { ...slide };
  delete next.masterId;
  delete next.layoutId;
  return next;
}

/** Rename a master (blank falls back to the master id at the serialize edge). */
export function mastersWithName(masters: readonly SlideMaster[], masterId: string, name: string): SlideMaster[] {
  return masters.map((master) => (master.id === masterId ? { ...master, name } : master));
}

/**
 * Patch fields of a master's style-override theme, merging onto its current
 * theme. A `null` / empty value clears that one field; an empty resulting
 * theme is dropped so the master stays byte-clean.
 */
export function mastersWithThemeField(
  masters: readonly SlideMaster[],
  masterId: string,
  patch: Partial<DeckTheme>,
): SlideMaster[] {
  return masters.map((master) => {
    if (master.id !== masterId) return master;
    const merged: DeckTheme = { ...(master.theme ?? {}) };
    for (const [key, value] of Object.entries(patch)) {
      if (value === undefined || value === null || value === "") {
        delete merged[key as keyof DeckTheme];
      } else {
        (merged as Record<string, unknown>)[key] = value;
      }
    }
    const next: SlideMaster = { ...master };
    if (Object.keys(merged).length > 0) next.theme = merged;
    else delete next.theme;
    return next;
  });
}

/**
 * One master's own text styles (Slide Master view: the title or body
 * placeholder's size, colour or bold), normalized into `master.theme`.
 * The normalized masters, or null when nothing changes (a no-op click must
 * not push an undo step or queue a save).
 */
export function mastersWithTextStyle(
  masters: readonly SlideMaster[],
  masterId: string,
  patch: Partial<DeckTheme>,
): SlideMaster[] | null {
  let changed = false;
  const nextMasters = masters.map((master) => {
    if (master.id !== masterId) return master;
    const merged = normalizeTheme({ ...(master.theme ?? {}), ...patch });
    if (pageDocumentJsonEquals(merged ?? null, master.theme ?? null)) return master;
    changed = true;
    const next: SlideMaster = { ...master };
    if (merged) next.theme = merged;
    else delete next.theme;
    return next;
  });
  return changed ? normalizeMasters(nextMasters) : null;
}

/** A layout or master that came from an imported file keeps its package part: its layout list is the file's. */
function isNative(item: { sourcePart?: unknown }): boolean {
  return typeof item.sourcePart === "string" && item.sourcePart.length > 0;
}

/**
 * Insert Layout: a new child layout of an authored master, drawn with a
 * built-in layout's placeholders. Null for an imported master (its layouts
 * are the file's parts) or a master already holding the most layouts.
 */
export function mastersWithNewLayout(
  masters: readonly SlideMaster[],
  masterId: string,
  layout: { id: string; name: string; builtInLayoutId: LayoutId },
): SlideMaster[] | null {
  const master = masters.find((candidate) => candidate.id === masterId);
  const layouts = master?.layouts ?? [];
  if (!master || isNative(master) || layouts.length >= MAX_MASTER_LAYOUTS || layouts.some((item) => item.id === layout.id)) {
    return null;
  }
  const added: SlideMasterLayout = { id: layout.id, name: layout.name, builtInLayoutId: layout.builtInLayoutId };
  return normalizeMasters(
    masters.map((candidate) => (candidate === master ? { ...master, layouts: [...layouts, added] } : candidate)),
  );
}

/** Rename Layout on an authored layout; null for an imported one or an unknown id. */
export function mastersWithLayoutName(
  masters: readonly SlideMaster[],
  masterId: string,
  layoutId: string,
  name: string,
): SlideMaster[] | null {
  const master = masters.find((candidate) => candidate.id === masterId);
  const layout = master?.layouts?.find((item) => item.id === layoutId);
  if (!master || !layout || isNative(layout)) return null;
  return normalizeMasters(
    masters.map((candidate) =>
      candidate === master
        ? { ...master, layouts: (master.layouts ?? []).map((item) => (item === layout ? { ...item, name } : item)) }
        : candidate,
    ),
  );
}

/**
 * Delete Layout: an authored layout no slide uses (PowerPoint offers Delete
 * only for a layout no slide is based on). Null otherwise.
 */
export function mastersWithoutLayout(
  masters: readonly SlideMaster[],
  slides: ReadonlyArray<{ layoutId?: unknown }>,
  masterId: string,
  layoutId: string,
): SlideMaster[] | null {
  const master = masters.find((candidate) => candidate.id === masterId);
  const layout = master?.layouts?.find((item) => item.id === layoutId);
  if (!master || !layout || isNative(layout) || slides.some((slide) => slide.layoutId === layoutId)) return null;
  return normalizeMasters(
    masters.map((candidate) =>
      candidate === master ? { ...master, layouts: (master.layouts ?? []).filter((item) => item !== layout) } : candidate,
    ),
  );
}
