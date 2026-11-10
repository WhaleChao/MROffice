/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide Master view on a deck that uses one of PowerPoint's Office themes
 * (PPT-B030): the master and layouts it shows, and the edits it makes to the
 * master, as pure transforms of the deck theme.
 *
 * The deck has one master -- the theme's -- whose parts the server writes from
 * its catalog on download (`apply_office_theme_to_package`). Slide Master view
 * edits it through `theme.officeMaster` (lib/presentation-office-master-model):
 *
 * - a text style level (title level one, body levels one to five: size,
 *   bold, italic, colour) reaches every slide through the layout templates
 *   (`layoutFrameSourceForTheme` carries it, so `slidesRestyledForTheme`
 *   restyles every placeholder still holding the old template value, and the
 *   body ladder draws levels two to five from it); the download writes it into
 *   `p:txStyles` of the slide master and states nothing on the slides;
 * - the master's background is the theme's background fields, which every
 *   slide without a fill of its own already draws; the download writes it into
 *   the master's `p:bg`.
 *
 * Microsoft Support, "What is a slide master?": a change to the master applies
 * to every slide based on it, and the master is the top thumbnail with its
 * layouts below it.
 */

import {
  LAYOUT_DEFINITIONS,
  layoutBlocks,
  layoutFrameSourceForTheme,
  officeFrameWithMasterEdits,
  officeLayoutSlots,
  layoutPlaceholders,
  type LayoutDefinition,
  type LayoutFrameSource,
  type LayoutId,
} from "@mrnote/core/web/lib/presentation-layouts.ts";
import {
  OFFICE_MASTER_BODY_LEVELS,
  OFFICE_PRIMARY_MASTER_KEY,
  isOfficeAdditionalMasterKey,
  isOfficeCustomLayoutId,
  officeMasterLayoutRef,
  parseOfficeMasterLayoutRef,
  normalizeDeckOfficeMaster,
  normalizeOfficeMasterName,
  normalizeOfficeMasterTextValues,
  type DeckOfficeMaster,
  type OfficeMasterBodyLevel,
  type OfficeMasterPlaceholderFamily,
  type OfficeMasterTextValues,
  type OfficePlaceholderKind,
} from "@mrnote/core/web/lib/presentation-office-master-model.ts";
import { officeBodyLadder } from "@mrnote/core/web/lib/presentation-office-ladder-catalog.ts";
import {
  presentationOfficeFrameTextColor,
  presentationOfficeLayoutFrames,
  presentationOfficeTheme,
  presentationOfficeThemeVariant,
  type PresentationOfficeTheme,
  type PresentationOfficeThemeVariant,
} from "@mrnote/core/web/lib/presentation-office-themes.ts";
import { normalizeTheme, type DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import {
  officeAdditionalMasters,
  officeMasterLocalSlides,
  officeMasterTheme,
  routeOfficeLayout,
  themeWithOfficeAdditionalMaster,
} from "./presentation-office-master-routing";

/** The Office theme a deck's master view edits, or null when it edits none. */
export function officeMasterSelection(
  theme: DeckTheme | null | undefined,
): { theme: PresentationOfficeTheme; variant: PresentationOfficeThemeVariant } | null {
  const identity = theme?.identity;
  // An imported deck keeps its package's own masters (the native catalog).
  if (!identity || identity.selectionKind !== "office" || identity.source) return null;
  const office = presentationOfficeTheme(identity.themeId);
  const variant = presentationOfficeThemeVariant(identity.themeId, identity.variantId);
  return office && variant ? { theme: office, variant } : null;
}

/**
 * The layouts of the theme's master, in PowerPoint's thumbnail order (the
 * catalog's: Title Slide first), as MRAI layout ids that have frames under
 * this variant.
 */
export function officeMasterLayoutIds(theme: DeckTheme | null | undefined): LayoutId[] {
  const selection = officeMasterSelection(theme);
  if (!selection) return [];
  // Edit Master > Delete took a layout off the master (PPT-B030).
  const deleted = new Set(normalizeDeckOfficeMaster(theme?.officeMaster)?.deletedLayouts ?? []);
  const catalog = LAYOUT_DEFINITIONS.map((layout) => layout.id).filter(
    (id) =>
      !deleted.has(id) &&
      (id === "blank" ||
        presentationOfficeLayoutFrames(selection.theme.id, selection.variant.id, id) !== undefined),
  );
  // Insert Layout's custom layouts follow, in the order they were inserted
  // (PPT-B030); the type carries their `custom-<letters>` ids as LayoutIds.
  const master = normalizeDeckOfficeMaster(theme?.officeMaster);
  const all = [...catalog, ...((master?.customLayouts ?? []) as unknown as LayoutId[])];
  // Insert Layout put a layout after the selected one: the stored order
  // first, anything it does not name after it (PPT-B030).
  const order = master?.layoutOrder;
  if (!order) return all;
  const present = new Set<string>(all);
  const ordered = order.filter((id) => present.has(id)) as LayoutId[];
  const named = new Set<string>(ordered);
  return [...ordered, ...all.filter((id) => !named.has(id))];
}

export type OfficeMasterTextStyleName = "title" | "body";

/** The level a text-style edit addresses: title level one, body levels one to five. */
export type OfficeMasterTextLevel = { style: "title"; level: "1" } | { style: "body"; level: OfficeMasterBodyLevel };

/** What the master's text style currently states at `target` over the theme. */
export function officeMasterTextValues(
  theme: DeckTheme | null | undefined,
  target: OfficeMasterTextLevel,
  masterKey?: string,
): OfficeMasterTextValues {
  if (masterKey && masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    return officeMasterTextValues(officeMasterTheme(theme, masterKey), target);
  }
  const styles = normalizeDeckOfficeMaster(theme?.officeMaster)?.textStyles;
  return (
    (target.style === "title" ? styles?.title?.["1"] : styles?.body?.[target.level]) ?? {}
  );
}

/**
 * The theme with one text-style level of the master changed. A patch value
 * of `null` clears that attribute (back to the theme's). Null when nothing
 * changes or the deck has no Office master to edit.
 */
export function themeWithOfficeMasterTextStyle(
  theme: DeckTheme | null | undefined,
  target: OfficeMasterTextLevel,
  patch: { [K in keyof OfficeMasterTextValues]?: OfficeMasterTextValues[K] | null },
  masterKey?: string,
): DeckTheme | null {
  if (masterKey && masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    return themeWithOfficeAdditionalMaster(
      theme,
      masterKey,
      themeWithOfficeMasterTextStyle(officeMasterTheme(theme, masterKey), target, patch),
    );
  }
  if (!theme || !officeMasterSelection(theme)) return null;
  const master: DeckOfficeMaster = normalizeDeckOfficeMaster(theme.officeMaster) ?? {};
  const current = officeMasterTextValues(theme, target);
  const merged: Record<string, unknown> = { ...current };
  for (const [key, value] of Object.entries(patch)) {
    if (value === null || value === undefined) delete merged[key];
    else merged[key] = value;
  }
  const values = normalizeOfficeMasterTextValues(merged);
  const textStyles = { ...(master.textStyles ?? {}) };
  if (target.style === "title") {
    if (values) textStyles.title = { "1": values };
    else delete textStyles.title;
  } else {
    const body = { ...(textStyles.body ?? {}) };
    if (values) body[target.level] = values;
    else delete body[target.level];
    if (Object.keys(body).length > 0) textStyles.body = body;
    else delete textStyles.body;
  }
  const nextMaster: DeckOfficeMaster = { ...master };
  if (Object.keys(textStyles).length > 0) nextMaster.textStyles = textStyles;
  else delete nextMaster.textStyles;
  return withOfficeMaster(theme, nextMaster);
}

/** The background fields a master's Format Background writes. */
export interface OfficeMasterBackground {
  background?: string;
  backgroundGradient?: unknown;
  backgroundPattern?: unknown;
}

/**
 * The theme with the master's background set (Format Background on the
 * master): exactly one fill in the theme's background fields and the master
 * marked as stating it. A picture fill is not a master fill the download can
 * write yet, so it is refused (null), as is a no-op.
 */
export function themeWithOfficeMasterBackground(
  theme: DeckTheme | null | undefined,
  fill: OfficeMasterBackground,
): DeckTheme | null {
  if (!theme || !officeMasterSelection(theme)) return null;
  const next: DeckTheme = { ...theme };
  delete next.backgroundGradient;
  delete next.backgroundPattern;
  delete next.backgroundImage;
  if (typeof fill.background === "string") next.background = fill.background;
  if (fill.backgroundGradient !== undefined) next.backgroundGradient = fill.backgroundGradient;
  if (fill.backgroundPattern !== undefined) next.backgroundPattern = fill.backgroundPattern;
  const master = normalizeDeckOfficeMaster(theme.officeMaster) ?? {};
  return withOfficeMaster(next, { ...master, background: true });
}

/**
 * Reset Background on the master: the theme's own ground again (the
 * variant's, which the catalog master paints), and the master states none.
 */
export function themeWithoutOfficeMasterBackground(
  theme: DeckTheme | null | undefined,
): DeckTheme | null {
  const selection = officeMasterSelection(theme);
  if (!theme || !selection) return null;
  const next: DeckTheme = { ...theme, background: selection.variant.background };
  delete next.backgroundGradient;
  delete next.backgroundPattern;
  delete next.backgroundImage;
  const master = { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
  delete master.background;
  return withOfficeMaster(next, master);
}

function withOfficeMaster(theme: DeckTheme, master: DeckOfficeMaster): DeckTheme | null {
  const next: DeckTheme = { ...theme };
  const normalized = normalizeDeckOfficeMaster(master);
  if (normalized) next.officeMaster = normalized;
  else delete next.officeMaster;
  const result = normalizeTheme(next);
  return result && JSON.stringify(result) !== JSON.stringify(normalizeTheme(theme)) ? result : null;
}

/** PowerPoint's master prompts, in the viewer's language. */
export interface OfficeMasterPrompts {
  /** "Click to edit Master title style". */
  title: string;
  /** "Click to edit Master subtitle style". */
  subtitle: string;
  /** The five body levels, "Click to edit Master text styles" .. "Fifth level". */
  textLevels: readonly string[];
  /** What an inserted Chart, Table, SmartArt, Media or Online Image placeholder says. */
  kinds?: Partial<Record<OfficePlaceholderKind, string>>;
}

/** One card of the master view's thumbnail pane: the master or one of its layouts. */
export interface OfficeMasterViewEntry {
  kind: "master" | "layout";
  /**
   * `master`, or the MRAI layout id; an inserted master's card is its key
   * and its layouts `<key>--<layout>` (PPT-B030).
   */
  key: string;
  /** The master the card belongs to: `master` (master one) or an inserted master's key. */
  masterKey: string;
  layoutId: LayoutId;
  /** The layout's message key suffix under `presentation.layout.*`; none for the master. */
  nameKey?: string;
  /** Slides of the deck on this layout (all slides for the master). */
  usage: number;
  /** A slide-shaped preview the slide renderer draws. */
  slide: {
    id: string;
    title: string;
    layout: LayoutId;
    widthEmu?: number;
    heightEmu?: number;
    blocks: Array<Record<string, unknown>>;
  };
}

/**
 * The master and its layouts as Slide Master view shows them for an Office
 * theme deck: each card is the layout's placeholders on the deck's slide
 * size, holding PowerPoint's master prompts ("Click to edit Master title
 * style"; the body's five levels) so the master's text styles and the list
 * ladder are drawn exactly as the slides inherit them. The master's own card
 * uses the Title and Content frames, where the Office masters keep their
 * title and body placeholders. Empty when the deck has no Office master.
 */
export function officeMasterViewEntries(
  theme: DeckTheme | null | undefined,
  slides: ReadonlyArray<{ layout?: unknown; widthEmu?: number | null; heightEmu?: number | null }>,
  prompts: OfficeMasterPrompts,
): OfficeMasterViewEntry[] {
  if (!officeMasterSelection(theme)) return [];
  // Each master in thumbnail order with its layouts beneath it: master one,
  // then every master Insert Slide Master added (PPT-B030). A master's card
  // counts the slides on its own layouts.
  const additional = officeAdditionalMasters(theme);
  const primarySlides = slides.filter((slide) => !isOfficeAdditionalMasterLayout(slide.layout, additional));
  const entries = singleOfficeMasterEntries(theme, primarySlides, prompts, OFFICE_PRIMARY_MASTER_KEY, slides);
  if (additional.length > 0) entries[0] = { ...entries[0], usage: primarySlides.length };
  for (const master of additional) {
    const masterTheme = officeMasterTheme(theme, master.key);
    if (!masterTheme) continue;
    const local = officeMasterLocalSlides(slides, master.key).slides;
    for (const entry of singleOfficeMasterEntries(masterTheme, local, prompts, master.key, slides)) {
      entries.push({
        ...entry,
        usage: entry.kind === "master" ? local.length : entry.usage,
        key: entry.kind === "master" ? master.key : officeMasterLayoutRef(master.key, entry.key),
        layoutId: officeMasterLayoutRef(master.key, entry.layoutId) as LayoutId,
      });
    }
  }
  return entries;
}

function isOfficeAdditionalMasterLayout(
  layout: unknown,
  masters: ReadonlyArray<{ key: string }>,
): boolean {
  return (
    typeof layout === "string" &&
    masters.some((master) => layout.startsWith(`${master.key}--`))
  );
}

function singleOfficeMasterEntries(
  theme: DeckTheme | null | undefined,
  slides: ReadonlyArray<{ layout?: unknown; widthEmu?: number | null; heightEmu?: number | null }>,
  prompts: OfficeMasterPrompts,
  masterKey: string,
  sizeSlides: ReadonlyArray<{ widthEmu?: number | null; heightEmu?: number | null }>,
): OfficeMasterViewEntry[] {
  if (!officeMasterSelection(theme)) return [];
  const sized = sizeSlides.find((slide) => slide.widthEmu && slide.heightEmu);
  const widthEmu = sized?.widthEmu ?? undefined;
  const heightEmu = sized?.heightEmu ?? undefined;
  const source = layoutFrameSourceForTheme(theme);
  // The master's own placeholders deleted in Slide Master view (PPT-B030).
  const masterOff = new Set(normalizeDeckOfficeMaster(theme?.officeMaster)?.masterPlaceholdersOff ?? []);
  // An inserted master's previews carry its key, so no two cards share a slide id.
  const idKey = (key: string) => (masterKey === OFFICE_PRIMARY_MASTER_KEY ? key : `${masterKey}-${key}`);
  const preview = (key: string, layoutId: LayoutId): OfficeMasterViewEntry["slide"] => ({
    id: `office-master-${idKey(key)}`,
    title: "",
    layout: layoutId,
    ...(widthEmu && heightEmu ? { widthEmu, heightEmu } : {}),
    blocks: layoutBlocks(layoutId, { width: widthEmu, height: heightEmu }, source).flatMap((block, index) => {
      if (key === "master" && masterOff.has(block.role === "title" ? "title" : "body")) return [];
      const id = `office-master-${idKey(key)}-${index}`;
      if (block.role === "title") return [{ ...block, id, text: prompts.title }];
      if (block.role === "subtitle") return [{ ...block, id, text: prompts.subtitle }];
      if (block.role === "picture") return [{ ...block, id }];
      const kind = key === "master" ? undefined : officeLayoutSlots(source, layoutId)?.[index]?.kind;
      const kindPrompt = kind ? prompts.kinds?.[kind] : undefined;
      if (kindPrompt) return [{ ...block, id, text: kindPrompt }];
      const levels = prompts.textLevels.slice(0, 5);
      return [
        {
          ...block,
          id,
          text: levels.join("\n"),
          paragraphs: levels.map((text, level) => ({ text, level, bullet: "bullet" })),
        },
      ];
    }),
  });
  const usage = (layoutId: LayoutId) =>
    slides.filter((slide) => slide.layout === layoutId || (layoutId === "title" && slide.layout === "titleSlide")).length;
  const entries: OfficeMasterViewEntry[] = [
    {
      kind: "master",
      key: "master",
      masterKey,
      layoutId: "title-content",
      usage: slides.length,
      slide: preview("master", "title-content"),
    },
  ];
  for (const layoutId of officeMasterLayoutIds(theme)) {
    const definition = LAYOUT_DEFINITIONS.find((layout) => layout.id === layoutId);
    entries.push({
      kind: "layout",
      key: layoutId,
      masterKey,
      layoutId,
      // Insert Layout's layouts are "Custom Layout" until renamed (PPT-B030).
      ...(definition ? { nameKey: definition.nameKey } : isOfficeCustomLayoutId(layoutId) ? { nameKey: "custom" } : {}),
      usage: usage(layoutId),
      slide: preview(layoutId, layoutId),
    });
  }
  return entries;
}

export { OFFICE_MASTER_BODY_LEVELS };

/**
 * What the master's text style draws at `target`: the master's own value
 * where Slide Master view set one, else the theme's -- the title of the
 * master's Title and Content frames, or the body ladder's level -- so the
 * controls show a value, never an empty box.
 */
export function officeMasterEffectiveTextValues(
  theme: DeckTheme | null | undefined,
  target: OfficeMasterTextLevel,
  masterKey?: string,
): OfficeMasterTextValues {
  if (masterKey && masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    return officeMasterEffectiveTextValues(officeMasterTheme(theme, masterKey), target);
  }
  const selection = officeMasterSelection(theme);
  if (!selection) return {};
  const own = officeMasterTextValues(theme, target);
  const frames = presentationOfficeLayoutFrames(selection.theme.id, selection.variant.id, "title-content");
  const frame = frames?.[target.style === "title" ? 0 : 1];
  const base: OfficeMasterTextValues = {};
  if (target.style === "title") {
    if (frame?.sz) base.fontSize = frame.sz / 100;
    if (frame?.b !== undefined) base.bold = frame.b;
  } else {
    const level = officeBodyLadder(selection.theme.id, selection.variant.id, "title-content", 1)?.[
      Number(target.level) - 1
    ];
    if (level?.sizePt !== undefined) base.fontSize = level.sizePt;
  }
  const color = frame ? presentationOfficeFrameTextColor(frame, selection.variant) : null;
  if (color) base.color = color;
  return { bold: false, italic: false, ...base, ...own };
}

/** "1-3, 5": slide numbers as PowerPoint lists the slides a layout is used by. */
export function slideNumberRanges(numbers: readonly number[]): string {
  const sorted = [...new Set(numbers)].sort((a, b) => a - b);
  const parts: string[] = [];
  for (let index = 0; index < sorted.length; ) {
    let end = index;
    while (end + 1 < sorted.length && sorted[end + 1] === sorted[end] + 1) end += 1;
    parts.push(end > index ? `${sorted[index]}-${sorted[end]}` : String(sorted[index]));
    index = end + 1;
  }
  return parts.join(", ");
}

/** The catalog slide every Office theme frame is recorded on. */
const CATALOG_WIDTH = 12_192_000;
const CATALOG_HEIGHT = 6_858_000;

/** One placeholder frame of a Slide Master view card, as fractions of the slide. */
export interface OfficeMasterCardFrame {
  /** `title` / `body` on the master card; the slot index on a layout card. */
  readonly target:
    | { kind: "master"; family: OfficeMasterPlaceholderFamily; masterKey?: string }
    | { kind: "layout"; layoutId: LayoutId; slot: number }
    | { kind: "added"; layoutId: LayoutId; index: number };
  readonly role: string;
  /** The Insert Placeholder kind of an added placeholder. */
  readonly kind?: OfficePlaceholderKind;
  /**
   * A frame the card shows and Delete takes off, but that does not move:
   * the master's date, footer and slide number, whose frames the deck does
   * not store (the download keeps the theme's own).
   */
  readonly fixed?: boolean;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

/**
 * The master's date, footer and slide number placeholders, where the Office
 * themes' 16:9 slide master puts them (`p:sp` type dt / ftr / sldNum on
 * slideMaster1: x 838200 / 4038600 / 8610600, y 6356350, cx 2743200 /
 * 4114800 / 2743200, cy 365125), in the Master Layout dialog's order. Slide
 * Master view shows them so they can be selected and deleted, as PowerPoint
 * does; the dialog's check boxes only put a deleted one back (PPT-B030).
 */
const OFFICE_MASTER_FOOTER_FRAMES: ReadonlyArray<{
  family: Extract<OfficeMasterPlaceholderFamily, "date" | "footer" | "slideNumber">;
  role: string;
  x: number;
  cx: number;
}> = [
  { family: "date", role: "date", x: 838_200, cx: 2_743_200 },
  { family: "footer", role: "footer", x: 4_038_600, cx: 4_114_800 },
  { family: "slideNumber", role: "slideNumber", x: 8_610_600, cx: 2_743_200 },
];
const OFFICE_MASTER_FOOTER_Y = 6_356_350;
const OFFICE_MASTER_FOOTER_CY = 365_125;

/**
 * The placeholder frames the selected card shows, with Slide Master view's
 * moves applied: the master's title and body (drawn on the Title and Content
 * frames, which inherit them), or every placeholder of a layout.
 */
export function officeMasterCardFrames(
  theme: DeckTheme | null | undefined,
  entry: Pick<OfficeMasterViewEntry, "kind" | "layoutId"> & { masterKey?: string },
): OfficeMasterCardFrame[] {
  // A card of an inserted master: its frames on that master, the targets
  // addressing it (PPT-B030).
  if (entry.masterKey && entry.masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    const masterKey = entry.masterKey;
    const local = routeOfficeLayout(theme, entry.layoutId)?.layoutId ?? entry.layoutId;
    return officeMasterCardFrames(officeMasterTheme(theme, masterKey), {
      kind: entry.kind,
      layoutId: local as LayoutId,
    }).map((frame) => ({
      ...frame,
      target:
        frame.target.kind === "master"
          ? { ...frame.target, masterKey }
          : { ...frame.target, layoutId: officeMasterLayoutRef(masterKey, frame.target.layoutId) as LayoutId },
    }));
  }
  const selection = officeMasterSelection(theme);
  if (!selection) return [];
  const source = layoutFrameSourceForTheme(theme);
  if (entry.kind === "master") {
    // The master card is the title and body placeholders the layouts inherit,
    // drawn on the Title and Content frames.
    const frames = presentationOfficeLayoutFrames(selection.theme.id, selection.variant.id, entry.layoutId) ?? [];
    // A placeholder deleted from the master is not on its card (PPT-B030).
    const off = new Set(normalizeDeckOfficeMaster(theme?.officeMaster)?.masterPlaceholdersOff ?? []);
    const text = frames.slice(0, 2).flatMap((catalogFrame, slot): OfficeMasterCardFrame[] => {
      if (off.has(catalogFrame.role === "title" ? "title" : "body")) return [];
      const frame = officeFrameWithMasterEdits(source, entry.layoutId, slot, catalogFrame);
      return [
        {
          role: frame.role,
          left: frame.x / CATALOG_WIDTH,
          top: frame.y / CATALOG_HEIGHT,
          width: frame.cx / CATALOG_WIDTH,
          height: frame.cy / CATALOG_HEIGHT,
          target: { kind: "master", family: frame.role === "title" ? "title" : "body" },
        },
      ];
    });
    // ...and its date, footer and slide number, until deleted.
    const footers = OFFICE_MASTER_FOOTER_FRAMES.filter((frame) => !off.has(frame.family)).map(
      (frame): OfficeMasterCardFrame => ({
        role: frame.role,
        left: frame.x / CATALOG_WIDTH,
        top: OFFICE_MASTER_FOOTER_Y / CATALOG_HEIGHT,
        width: frame.cx / CATALOG_WIDTH,
        height: OFFICE_MASTER_FOOTER_CY / CATALOG_HEIGHT,
        fixed: true,
        target: { kind: "master", family: frame.family },
      }),
    );
    return [...text, ...footers];
  }
  return (officeLayoutSlots(source, entry.layoutId) ?? []).map(
    (slot): OfficeMasterCardFrame => ({
      role: slot.frame.role,
      left: slot.frame.x / CATALOG_WIDTH,
      top: slot.frame.y / CATALOG_HEIGHT,
      width: slot.frame.cx / CATALOG_WIDTH,
      height: slot.frame.cy / CATALOG_HEIGHT,
      ...(slot.kind ? { kind: slot.kind } : {}),
      target:
        slot.addedIndex !== undefined
          ? { kind: "added", layoutId: entry.layoutId, index: slot.addedIndex }
          : { kind: "layout", layoutId: entry.layoutId, slot: slot.catalogSlot ?? 0 },
    }),
  );
}

/**
 * The theme with one placeholder frame moved or resized in Slide Master view
 * (`rect` as fractions of the slide). Every slide placeholder still on the old
 * frame follows it in the same step (`slidesRestyledForTheme`), and the
 * download writes the frame into the master or the layout, so PowerPoint's
 * inheriting placeholders follow it too. Null for a no-op.
 */
export function themeWithOfficeMasterFrame(
  theme: DeckTheme | null | undefined,
  target: OfficeMasterCardFrame["target"],
  rect: { left: number; top: number; width: number; height: number },
): DeckTheme | null {
  const masterKey =
    target.kind === "master" ? target.masterKey : routeOfficeLayout(theme, target.layoutId)?.masterKey;
  if (masterKey && masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    const local =
      target.kind === "master"
        ? { kind: "master" as const, family: target.family }
        : { ...target, layoutId: (routeOfficeLayout(theme, target.layoutId)?.layoutId ?? target.layoutId) as LayoutId };
    return themeWithOfficeAdditionalMaster(
      theme,
      masterKey,
      themeWithOfficeMasterFrame(officeMasterTheme(theme, masterKey), local, rect),
    );
  }
  if (!theme || !officeMasterSelection(theme)) return null;
  const frame: [number, number, number, number] = [
    Math.round(rect.left * CATALOG_WIDTH),
    Math.round(rect.top * CATALOG_HEIGHT),
    Math.max(1, Math.round(rect.width * CATALOG_WIDTH)),
    Math.max(1, Math.round(rect.height * CATALOG_HEIGHT)),
  ];
  const master: DeckOfficeMaster = { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
  if (target.kind === "master") {
    // Only the title and text frames are stored; the footers keep the theme's.
    if (target.family !== "title" && target.family !== "body") return null;
    master.masterFrames = { ...(master.masterFrames ?? {}), [target.family]: frame };
  } else if (target.kind === "added") {
    const list = [...(master.addedPlaceholders?.[target.layoutId] ?? [])];
    const entry = list[target.index];
    if (!entry) return null;
    list[target.index] = { ...entry, frame };
    master.addedPlaceholders = { ...(master.addedPlaceholders ?? {}), [target.layoutId]: list };
  } else {
    const layouts = { ...(master.layoutFrames ?? {}) };
    layouts[target.layoutId] = { ...(layouts[target.layoutId] ?? {}), [String(target.slot)]: frame };
    master.layoutFrames = layouts;
  }
  return withOfficeMaster(theme, master);
}

/**
 * Whether Slide Master view gave the master a background of its own (Format
 * Background or a Background Styles pick): the theme's own ground artwork is
 * then not the master's background any more and is not drawn under it.
 */
export function officeMasterStatesBackground(theme: DeckTheme | null | undefined): boolean {
  const master = normalizeDeckOfficeMaster(theme?.officeMaster);
  return master?.background === true || master?.backgroundStyle !== undefined;
}

/**
 * Rename Slide Master / Rename Layout (PPT-B030): `key` is `master` or the
 * MRAI layout id. A blank name gives the theme's own name back. Null for a
 * no-op.
 */
export function themeWithOfficeMasterName(
  theme: DeckTheme | null | undefined,
  key: string,
  name: string,
): DeckTheme | null {
  // An inserted master's card (its key) or one of its layouts (PPT-B030).
  const routed = routeOfficeLayout(theme, key);
  if (routed || isOfficeAdditionalMasterKey(key)) {
    const masterKey = routed?.masterKey ?? key;
    const local = routed?.layoutId ?? "master";
    // An inserted master always has a name (it is its theme's): a blank
    // one is refused rather than cleared.
    if (local === "master" && !normalizeOfficeMasterName(name)) return null;
    return themeWithOfficeAdditionalMaster(
      theme,
      masterKey,
      themeWithOfficeMasterName(officeMasterTheme(theme, masterKey), local, name),
    );
  }
  if (!theme || !officeMasterSelection(theme)) return null;
  const master: DeckOfficeMaster = { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
  const trimmed = normalizeOfficeMasterName(name);
  if (key === "master") {
    if (trimmed) master.name = trimmed;
    else delete master.name;
  } else {
    const names = { ...(master.layoutNames ?? {}) };
    if (trimmed) names[key] = trimmed;
    else delete names[key];
    if (Object.keys(names).length > 0) master.layoutNames = names;
    else delete master.layoutNames;
  }
  return withOfficeMaster(theme, master);
}

/** The name Rename gave the master (`master`) or a layout, if any. */
export function officeMasterGivenName(
  theme: DeckTheme | null | undefined,
  key: string,
): string | undefined {
  const routed = routeOfficeLayout(theme, key);
  if (routed) return officeMasterGivenName(routed.theme, routed.layoutId);
  if (isOfficeAdditionalMasterKey(key)) return officeMasterGivenName(officeMasterTheme(theme, key), "master");
  const master = normalizeDeckOfficeMaster(theme?.officeMaster);
  return key === "master" ? master?.name : master?.layoutNames?.[key];
}

/** Edit Master ▸ Preserve on the Office master: on, or off again. Null for a no-op. */
export function themeWithOfficeMasterPreserved(
  theme: DeckTheme | null | undefined,
  preserve: boolean,
  masterKey?: string,
): DeckTheme | null {
  if (masterKey && masterKey !== OFFICE_PRIMARY_MASTER_KEY) {
    return themeWithOfficeAdditionalMaster(
      theme,
      masterKey,
      themeWithOfficeMasterPreserved(officeMasterTheme(theme, masterKey), preserve),
    );
  }
  if (!theme || !officeMasterSelection(theme)) return null;
  const master: DeckOfficeMaster = { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
  if (preserve) master.preserve = true;
  else delete master.preserve;
  return withOfficeMaster(theme, master);
}

/** Whether the master (master one, or an inserted master by key) is preserved. */
export function officeMasterPreserved(theme: DeckTheme | null | undefined, masterKey: string): boolean {
  return normalizeDeckOfficeMaster(officeMasterTheme(theme, masterKey)?.officeMaster)?.preserve === true;
}

/** One group of the New Slide and Layout galleries: a master's name and its layouts. */
export interface OfficeLayoutGalleryGroup {
  readonly masterKey: string;
  /** The layouts as slides address them (`<key>--<layout>` on an inserted master). */
  readonly layoutIds: readonly string[];
}

/**
 * The New Slide and Layout galleries of an Office theme deck, one group per
 * master in thumbnail order (PPT-B030): PowerPoint heads each group with its
 * master's name and lists that master's layouts, and a slide may take a
 * layout of any master. One group (master one's) when the deck has one
 * master; none on a deck without an Office master.
 */
export function officeLayoutGalleryGroups(
  theme: DeckTheme | null | undefined,
  primaryLayoutIds: readonly string[],
): OfficeLayoutGalleryGroup[] {
  if (!officeMasterSelection(theme)) return [];
  return [
    { masterKey: OFFICE_PRIMARY_MASTER_KEY, layoutIds: primaryLayoutIds },
    ...officeAdditionalMasters(theme).map((master) => ({
      masterKey: master.key,
      layoutIds: officeMasterLayoutIds(officeMasterTheme(theme, master.key)).map((layoutId) =>
        officeMasterLayoutRef(master.key, layoutId),
      ),
    })),
  ];
}

/**
 * The gallery card of a layout of an inserted master (`<key>--<layout>`,
 * PPT-B030): the layout's name key and its placeholders under its master.
 */
export function officeLayoutGalleryDefinition(layoutRef: string, source: LayoutFrameSource): LayoutDefinition {
  const own = parseOfficeMasterLayoutRef(layoutRef)?.layoutId ?? layoutRef;
  const builtIn = LAYOUT_DEFINITIONS.find((layout) => layout.id === own);
  return {
    id: layoutRef as LayoutId,
    nameKey: builtIn?.nameKey ?? (isOfficeCustomLayoutId(own) ? "custom" : "blank"),
    placeholders: layoutPlaceholders(layoutRef, source),
  };
}
