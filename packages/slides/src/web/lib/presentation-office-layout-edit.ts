/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide Master view's structural edits of an Office theme deck's layouts
 * (PPT-B030), as pure transforms of the deck theme and slides:
 *
 * - Master Layout ▸ Insert Placeholder: Content, Content (Vertical), Text,
 *   Text (Vertical), Picture, Chart, Table, SmartArt, Media, Online Image.
 *   The placeholder joins the layout, and slides made from the layout from
 *   then on get it; slides already on the layout keep what they have
 *   (Microsoft Support, "Add, edit, or remove a placeholder on a slide
 *   layout": an existing slide shows a layout change once the layout is
 *   re-applied).
 * - Deleting a placeholder of a layout (select it, press Delete): the layout
 *   loses it. A slide's placeholder poured from it stops being a placeholder
 *   -- an empty one goes, one holding text stays as a text box -- so every
 *   other placeholder of the slide keeps pairing with its layout slot.
 * - Master Layout ▸ Title: the same delete for the layout's title, and its
 *   restore.
 * - Master Layout ▸ Footers: whether the layout keeps its date, footer and
 *   slide-number placeholders.
 * - Edit Master ▸ Delete on a layout: only a layout no slide uses can go,
 *   as in PowerPoint, where Delete is unavailable for a layout in use.
 *
 * All of it is stored in `theme.officeMaster` (lib/presentation-office-master-model)
 * and written into the download's layout parts by the server.
 */

import {
  LAYOUT_DEFINITIONS,
  layoutFrameSourceForTheme,
  normalizeSlideLayoutRef,
  officeLayoutSlots,
  type LayoutId,
} from "@mrnote/core/web/lib/presentation-layouts.ts";
import {
  OFFICE_MASTER_MAX_ADDED_PLACEHOLDERS,
  isOfficeCustomLayoutId,
  normalizeDeckOfficeMaster,
  type DeckOfficeMaster,
  type OfficePlaceholderKind,
} from "@mrnote/core/web/lib/presentation-office-master-model.ts";
import { officeMasterLayoutIds, officeMasterSelection } from "./presentation-office-master";
import {
  officeMasterLocalSlides,
  routeOfficeLayout,
  slidesFromOfficeMasterLocal,
  themeWithOfficeAdditionalMaster,
} from "./presentation-office-master-routing";
import { normalizeTheme, type DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

const CATALOG_WIDTH = 12_192_000;
const CATALOG_HEIGHT = 6_858_000;

interface SlideLike {
  layout?: unknown;
  blocks?: unknown;
  source?: unknown;
}

/** Whether the slide is on the layout (a Title Slide may be stored as `titleSlide`). */
export function slideUsesOfficeLayout(slide: SlideLike, layoutId: string): boolean {
  // A slide with no (or an unknown) layout is on Blank, as the download
  // writes it; a custom layout's slides are its own (PPT-B030).
  const own = slide.layout === "titleSlide" ? "title" : normalizeSlideLayoutRef(slide.layout);
  return own === layoutId;
}

function isImportedSlide(slide: SlideLike): boolean {
  const source = slide.source as Record<string, unknown> | undefined;
  return Boolean(source && typeof source.slideId === "string");
}

/** A placeholder block the deck poured from its layout (the download pairs these with layout slots). */
function isOwnedPlaceholder(block: unknown): block is Record<string, unknown> {
  const record = block as Record<string, unknown> | null;
  if (!record || record.type !== "text") return false;
  if (typeof record.placeholder !== "string" || typeof record.role !== "string") return false;
  const source = record.source as Record<string, unknown> | undefined;
  return !(source && typeof source.slideId === "string");
}

function holdsText(block: Record<string, unknown>): boolean {
  const filled = (value: unknown) => typeof value === "string" && value.trim() !== "";
  if (filled(block.text)) return true;
  const lists = [block.runs, block.paragraphs];
  return lists.some(
    (list) =>
      Array.isArray(list) &&
      list.some((item) => {
        const entry = item as Record<string, unknown> | null;
        if (!entry) return false;
        if (filled(entry.text)) return true;
        return Array.isArray(entry.runs) && entry.runs.some((run) => filled((run as Record<string, unknown>)?.text));
      }),
  );
}

function withMaster(theme: DeckTheme, master: DeckOfficeMaster): DeckTheme | null {
  const next: DeckTheme = { ...theme };
  const normalized = normalizeDeckOfficeMaster(master);
  if (normalized) next.officeMaster = normalized;
  else delete next.officeMaster;
  const result = normalizeTheme(next);
  return result && JSON.stringify(result) !== JSON.stringify(normalizeTheme(theme)) ? result : null;
}

function currentMaster(theme: DeckTheme): DeckOfficeMaster {
  return { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
}

/** The layout's slot refs in pour order, as stored or the catalog's. */
function slotRefs(theme: DeckTheme, layoutId: string): string[] | null {
  const slots = officeLayoutSlots(layoutFrameSourceForTheme(theme), layoutId);
  return slots ? slots.map((slot) => slot.ref) : null;
}

/** Whether the layout can take structural edits here (an Office deck's layout with placeholders). */
export function officeLayoutEditable(theme: DeckTheme | null | undefined, layoutId: string): boolean {
  // A layout of an inserted master is edited on that master (PPT-B030).
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) return officeLayoutEditable(routed.theme, routed.layoutId);
  if (!theme || !officeMasterSelection(theme) || layoutId === "blank") return false;
  return officeMasterLayoutIds(theme).includes(layoutId as LayoutId) && slotRefs(theme, layoutId) !== null;
}

/**
 * Insert Placeholder on a layout: the kind's placeholder in the middle of the
 * slide (PowerPoint has the author drag it out; here it lands centred, each
 * further one stepped down and right, and is then moved and sized like any
 * frame). Null when the layout cannot take one or holds the maximum.
 */
export function themeWithInsertedOfficePlaceholder(
  theme: DeckTheme | null | undefined,
  layoutId: string,
  kind: OfficePlaceholderKind,
): DeckTheme | null {
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) {
    return themeWithOfficeAdditionalMaster(
      theme,
      routed.masterKey,
      themeWithInsertedOfficePlaceholder(routed.theme, routed.layoutId, kind),
    );
  }
  if (!theme || !officeLayoutEditable(theme, layoutId)) return null;
  const refs = slotRefs(theme, layoutId);
  if (!refs) return null;
  const master = currentMaster(theme);
  const added = [...(master.addedPlaceholders?.[layoutId] ?? [])];
  if (added.length >= OFFICE_MASTER_MAX_ADDED_PLACEHOLDERS) return null;
  const step = 228_600 * (added.length % 8);
  const width = Math.round(CATALOG_WIDTH / 2);
  const height = Math.round(CATALOG_HEIGHT * 0.4);
  added.push({
    kind,
    frame: [
      Math.round((CATALOG_WIDTH - width) / 2) + step,
      Math.round((CATALOG_HEIGHT - height) / 2) + step,
      width,
      height,
    ],
  });
  master.addedPlaceholders = { ...(master.addedPlaceholders ?? {}), [layoutId]: added };
  master.layoutSlots = { ...(master.layoutSlots ?? {}), [layoutId]: [...refs, `a${added.length - 1}`] };
  return withMaster(theme, master);
}

export interface OfficeDeckEdit<S> {
  theme: DeckTheme;
  slides: S[];
}

/**
 * A deck edit of an inserted master's layout (PPT-B030): the transform runs on
 * that master and its slides, and both come back into the deck. Null when the
 * layout is master one's.
 */
function routedOfficeDeckEdit<S extends SlideLike>(
  theme: DeckTheme | null | undefined,
  slides: readonly S[],
  layoutId: string,
  edit: (masterTheme: DeckTheme, localSlides: S[], localLayoutId: string) => OfficeDeckEdit<S> | null,
): { edit: OfficeDeckEdit<S> | null } | null {
  const routed = routeOfficeLayout(theme, layoutId);
  if (!routed) return null;
  const local = officeMasterLocalSlides(slides, routed.masterKey);
  const result = edit(routed.theme, local.slides, routed.layoutId);
  if (!result) return { edit: null };
  const nextTheme =
    themeWithOfficeAdditionalMaster(theme, routed.masterKey, result.theme) ?? (theme ? (theme as DeckTheme) : null);
  if (!nextTheme) return { edit: null };
  return {
    edit: {
      theme: nextTheme,
      slides: slidesFromOfficeMasterLocal(slides, routed.masterKey, local, result.slides),
    },
  };
}

/**
 * Delete the layout's placeholder at `position` (its place in pour order):
 * the layout loses it and every slide on the layout lets go of the
 * placeholder it poured from it (removed when empty, kept as a text box when
 * it holds text). Null when there is nothing to delete.
 */
export function deckWithDeletedOfficePlaceholder<S extends SlideLike>(
  theme: DeckTheme | null | undefined,
  slides: readonly S[],
  layoutId: string,
  position: number,
): OfficeDeckEdit<S> | null {
  const routed = routedOfficeDeckEdit(theme, slides, layoutId, (masterTheme, localSlides, local) =>
    deckWithDeletedOfficePlaceholder(masterTheme, localSlides, local, position),
  );
  if (routed) return routed.edit;
  if (!theme || !officeLayoutEditable(theme, layoutId)) return null;
  const refs = slotRefs(theme, layoutId);
  if (!refs || position < 0 || position >= refs.length) return null;
  const master = currentMaster(theme);
  master.layoutSlots = {
    ...(master.layoutSlots ?? {}),
    [layoutId]: refs.filter((_, index) => index !== position),
  };
  const nextTheme = withMaster(theme, master);
  if (!nextTheme) return null;
  const nextSlides = slides.map((slide) => {
    if (!slideUsesOfficeLayout(slide, layoutId) || isImportedSlide(slide) || !Array.isArray(slide.blocks)) {
      return slide;
    }
    let owned = -1;
    let changed = false;
    const blocks: unknown[] = [];
    for (const block of slide.blocks as unknown[]) {
      if (!isOwnedPlaceholder(block)) {
        blocks.push(block);
        continue;
      }
      owned += 1;
      if (owned !== position) {
        blocks.push(block);
        continue;
      }
      changed = true;
      if (!holdsText(block)) continue;
      const loose: Record<string, unknown> = { ...block };
      delete loose.placeholder;
      delete loose.role;
      blocks.push(loose);
    }
    return changed ? ({ ...slide, blocks } as S) : slide;
  });
  return { theme: nextTheme, slides: nextSlides };
}

/** Whether the layout keeps a title placeholder. */
export function officeLayoutHasTitle(theme: DeckTheme | null | undefined, layoutId: string): boolean {
  if (!theme) return false;
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) return officeLayoutHasTitle(routed.theme, routed.layoutId);
  return (officeLayoutSlots(layoutFrameSourceForTheme(theme), layoutId) ?? []).some(
    (slot) => slot.frame.role === "title",
  );
}

/**
 * Master Layout ▸ Title: off deletes the layout's title placeholder (as
 * {@link deckWithDeletedOfficePlaceholder}); on gives the catalog title back,
 * last in pour order so the slides already poured keep their pairing.
 */
export function deckWithOfficeLayoutTitle<S extends SlideLike>(
  theme: DeckTheme | null | undefined,
  slides: readonly S[],
  layoutId: string,
  on: boolean,
): OfficeDeckEdit<S> | null {
  const routed = routedOfficeDeckEdit(theme, slides, layoutId, (masterTheme, localSlides, local) =>
    deckWithOfficeLayoutTitle(masterTheme, localSlides, local, on),
  );
  if (routed) return routed.edit;
  if (!theme || !officeLayoutEditable(theme, layoutId)) return null;
  const slots = officeLayoutSlots(layoutFrameSourceForTheme(theme), layoutId) ?? [];
  if (!on) {
    const position = slots.findIndex((slot) => slot.frame.role === "title");
    return position < 0 ? null : deckWithDeletedOfficePlaceholder(theme, slides, layoutId, position);
  }
  if (slots.some((slot) => slot.frame.role === "title")) return null;
  const selection = officeMasterSelection(theme);
  // A custom layout's catalog is Title Only's (PPT-B030).
  const catalog = selection
    ? officeLayoutSlots(
        { kind: "office", themeId: selection.theme.id, variantId: selection.variant.id },
        isOfficeCustomLayoutId(layoutId) ? "title-only" : layoutId,
      )
    : undefined;
  const title = catalog?.find((slot) => slot.frame.role === "title");
  if (!title) return null;
  const master = currentMaster(theme);
  master.layoutSlots = {
    ...(master.layoutSlots ?? {}),
    [layoutId]: [...slots.map((slot) => slot.ref), title.ref],
  };
  const nextTheme = withMaster(theme, master);
  return nextTheme ? { theme: nextTheme, slides: [...slides] } : null;
}

/** Whether the layout keeps its date, footer and slide-number placeholders. */
export function officeLayoutHasFooters(theme: DeckTheme | null | undefined, layoutId: string): boolean {
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) return officeLayoutHasFooters(routed.theme, routed.layoutId);
  return !(normalizeDeckOfficeMaster(theme?.officeMaster)?.layoutsWithoutFooters ?? []).includes(layoutId);
}

/** Master Layout ▸ Footers on the layout. Null for a no-op. */
export function themeWithOfficeLayoutFooters(
  theme: DeckTheme | null | undefined,
  layoutId: string,
  on: boolean,
): DeckTheme | null {
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) {
    return themeWithOfficeAdditionalMaster(
      theme,
      routed.masterKey,
      themeWithOfficeLayoutFooters(routed.theme, routed.layoutId, on),
    );
  }
  if (!theme || !officeLayoutEditable(theme, layoutId)) return null;
  const master = currentMaster(theme);
  const without = new Set(master.layoutsWithoutFooters ?? []);
  if (on) without.delete(layoutId);
  else without.add(layoutId);
  master.layoutsWithoutFooters = [...without];
  if (master.layoutsWithoutFooters.length === 0) delete master.layoutsWithoutFooters;
  return withMaster(theme, master);
}

/** Why Edit Master ▸ Delete cannot remove the layout, or null when it can. */
export function officeLayoutDeleteBlocker(
  theme: DeckTheme | null | undefined,
  slides: readonly SlideLike[],
  layoutId: string,
): "in-use" | "last" | "unavailable" | null {
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) {
    return officeLayoutDeleteBlocker(
      routed.theme,
      officeMasterLocalSlides(slides, routed.masterKey).slides,
      routed.layoutId,
    );
  }
  if (!theme || !officeMasterSelection(theme)) return "unavailable";
  const layouts = officeMasterLayoutIds(theme);
  if (!layouts.includes(layoutId as LayoutId)) return "unavailable";
  if (slides.some((slide) => slideUsesOfficeLayout(slide, layoutId))) return "in-use";
  if (layouts.length <= 1) return "last";
  return null;
}

/** Edit Master ▸ Delete on an unused layout. Null when it cannot go. */
export function themeWithDeletedOfficeLayout(
  theme: DeckTheme | null | undefined,
  slides: readonly SlideLike[],
  layoutId: string,
): DeckTheme | null {
  const routed = routeOfficeLayout(theme, layoutId);
  if (routed) {
    return themeWithOfficeAdditionalMaster(
      theme,
      routed.masterKey,
      themeWithDeletedOfficeLayout(routed.theme, officeMasterLocalSlides(slides, routed.masterKey).slides, routed.layoutId),
    );
  }
  if (!theme || officeLayoutDeleteBlocker(theme, slides, layoutId) !== null) return null;
  const master = currentMaster(theme);
  if (isOfficeCustomLayoutId(layoutId)) {
    // A custom layout Insert Layout made is simply gone (PPT-B030).
    master.customLayouts = (master.customLayouts ?? []).filter((id) => id !== layoutId);
  } else {
    master.deletedLayouts = [...new Set([...(master.deletedLayouts ?? []), layoutId])];
  }
  // What Slide Master view kept for the layout goes with it.
  for (const key of ["layoutFrames", "layoutNames", "addedPlaceholders", "layoutSlots"] as const) {
    const map = master[key] as Record<string, unknown> | undefined;
    if (map && layoutId in map) {
      const rest = { ...map };
      delete rest[layoutId];
      (master as Record<string, unknown>)[key] = Object.keys(rest).length > 0 ? rest : undefined;
    }
  }
  for (const key of ["layoutsWithoutFooters", "layoutsHidingBackgroundGraphics", "layoutOrder"] as const) {
    const list = master[key];
    if (list) master[key] = list.filter((id) => id !== layoutId);
  }
  return withMaster(theme, master);
}

/** The layouts a new slide can be made from, in the master's order (deleted ones gone). */
export function officeAvailableLayoutIds(theme: DeckTheme | null | undefined): LayoutId[] {
  if (!officeMasterSelection(theme)) return LAYOUT_DEFINITIONS.map((layout) => layout.id);
  return officeMasterLayoutIds(theme);
}
