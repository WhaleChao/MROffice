/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Slide Master edits of a deck on one of PowerPoint's Office themes
 * (PPT-B030): `theme.officeMaster`.
 *
 * A from-scratch deck on an Office theme has ONE slide master, the theme's,
 * which the server writes from its catalog on download. What Slide Master
 * view changes on it is stored here and only here, the way Design ▸ Variants
 * ▸ Colors stores its pick (`colorSchemeId`):
 *
 * - `textStyles`: what the master's text styles state over the theme's
 *   (`p:txStyles`, ECMA-376 §19.3.1.52) -- the title style's level one and
 *   the body style's levels one to five, each its size in points, bold,
 *   italic and colour. Only stated attributes appear; the rest inherit the
 *   theme's.
 * - `background: true`: the master's own background (`p:cSld/p:bg`) is the
 *   theme's background fields (`background` / `backgroundGradient` /
 *   `backgroundPattern`), set in Slide Master view, rather than the catalog
 *   theme's own fill.
 * - `backgroundStyle`: a Background Styles pick (PPT-A073).
 *
 * Only Slide Master view writes it; picking a theme or a variant drops it
 * (a theme replaces the master, as in PowerPoint). Dependency-free so the
 * theme model and the zod schema can import it.
 */

export const OFFICE_MASTER_BODY_LEVELS = ["1", "2", "3", "4", "5"] as const;
export type OfficeMasterBodyLevel = (typeof OFFICE_MASTER_BODY_LEVELS)[number];

/** What one text-style level states. `fontSize` is in points. */
export interface OfficeMasterTextValues {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

export interface OfficeMasterTextStyles {
  title?: { "1"?: OfficeMasterTextValues };
  body?: Partial<Record<OfficeMasterBodyLevel, OfficeMasterTextValues>>;
}

export interface DeckOfficeMaster {
  textStyles?: OfficeMasterTextStyles;
  background?: true;
  /**
   * Design ▸ Variants ▸ Background Styles (PPT-A073): one of the theme's
   * twelve background styles, 1..12 in the gallery's order -- the master's
   * `p:bgRef` to a fill of the theme's `a:bgFillStyleLst` and the colour map
   * that goes with it (the dark styles map bg1/tx1 to dk1/lt1).
   */
  backgroundStyle?: number;
  /**
   * Placeholder frames moved or resized in Slide Master view (PPT-B030),
   * `[x, y, cx, cy]` in EMU on the theme's 16:9 catalog slide (12192000 x
   * 6858000; the deck's size scales them like every layout frame): the
   * master's title / body placeholder -- which every layout placeholder that
   * states no frame of its own inherits -- and a layout's placeholder by its
   * slot (the editor's pour order).
   */
  masterFrames?: Partial<Record<"title" | "body", OfficeMasterFrame>>;
  layoutFrames?: Record<string, Record<string, OfficeMasterFrame>>;
  /** Rename Slide Master / Rename Layout (the parts' `p:cSld@name`). */
  name?: string;
  layoutNames?: Record<string, string>;
  /** Edit Master ▸ Preserve (`p:sldMaster@preserve`). */
  preserve?: true;
  /**
   * Master Layout ▸ Insert Placeholder (PPT-B030): the placeholders Slide
   * Master view added to a layout, by layout id, in the order they were
   * added. An entry is never removed or reordered (a slot ref `a<n>` names it
   * by position); deleting one drops its ref from `layoutSlots`.
   */
  addedPlaceholders?: Record<string, OfficeMasterAddedPlaceholder[]>;
  /**
   * A layout's placeholders after Slide Master view inserted or deleted one,
   * in the order a new slide pours them: `c<n>` is the catalog layout's slot
   * n, `a<n>` the n-th added placeholder. Absent: the catalog's slots in their
   * order. Deleting appends nothing; restoring the title (the Title check box)
   * appends its slot, so the slots earlier slides were poured from keep their
   * positions and the download pairs them as before.
   */
  layoutSlots?: Record<string, string[]>;
  /** Edit Master ▸ Delete on a layout no slide uses: the layout is gone. */
  deletedLayouts?: string[];
  /** Master Layout ▸ Footers cleared: the layout keeps no date, footer or slide number. */
  layoutsWithoutFooters?: string[];
  /**
   * Background ▸ Hide Background Graphics on a layout (PPT-B030): its slides
   * do not draw the master's graphics (`p:sldLayout@showMasterSp="0"`).
   */
  layoutsHidingBackgroundGraphics?: string[];
  /**
   * The master's own placeholders deleted in Slide Master view, which the
   * Master Layout dialog puts back (PPT-B030): title, text, date, slide
   * number, footer.
   */
  masterPlaceholdersOff?: OfficeMasterPlaceholderFamily[];
  /**
   * Edit Master ▸ Insert Layout (PPT-B030): the master's custom layouts, in
   * the order they were inserted, by id (`custom-<letters>`). Each starts as
   * PowerPoint's Custom Layout -- a title and the footers -- and takes
   * placeholders, names, frames and Hide Background Graphics like any layout.
   */
  customLayouts?: string[];
  /**
   * The master's layouts in thumbnail order once Insert Layout put one after
   * the selected layout (PPT-B030; PowerPoint inserts after the selection).
   * Absent: the catalog order, then the custom layouts.
   */
  layoutOrder?: string[];
  /**
   * Edit Master ▸ Insert Slide Master (PPT-B030): the masters after the
   * theme's own, in thumbnail order. The theme's master stays master one and
   * is this record itself; each further master is PowerPoint's inserted one
   * -- the default Office theme ("without theme colors or effects",
   * Microsoft Support, "Use multiple slide masters in one presentation")
   * with its eleven standard layouts -- and keeps its own Slide Master view
   * edits in the same keys as this record. Its layouts are addressed as
   * `<key>--<layout>` ({@link officeMasterLayoutRef}), on slides too.
   */
  additionalMasters?: OfficeAdditionalMaster[];
}

/**
 * One master Insert Slide Master added: its key and what Slide Master view
 * states over it (the keys of {@link DeckOfficeMaster} that address a master
 * and its layouts; its background is the theme's own).
 */
export interface OfficeAdditionalMaster
  extends Omit<DeckOfficeMaster, "additionalMasters" | "background" | "backgroundStyle"> {
  /** `master-<letters>`: master-b, master-c, ... (the theme's own master is `master`). */
  key: string;
}

/** The key of the theme's own master, master one. */
export const OFFICE_PRIMARY_MASTER_KEY = "master";
/** An inserted master's key. */
export const OFFICE_ADDITIONAL_MASTER_KEY = /^master-[a-z]{1,2}$/;
/** At most this many inserted masters (PowerPoint has no small cap; bounded here). */
export const OFFICE_MASTER_MAX_ADDITIONAL = 16;
/** The theme an inserted master carries: Office's default theme, its base variant. */
export const OFFICE_INSERTED_MASTER_THEME_ID = "office";
export const OFFICE_INSERTED_MASTER_VARIANT_ID = "v1";

const OFFICE_MASTER_LAYOUT_REF = /^(master-[a-z]{1,2})--([a-z][a-z-]{0,40})$/;

export function isOfficeAdditionalMasterKey(value: unknown): value is string {
  return typeof value === "string" && OFFICE_ADDITIONAL_MASTER_KEY.test(value);
}

/** A layout of an inserted master as slides and Slide Master view address it. */
export function officeMasterLayoutRef(masterKey: string, layoutId: string): string {
  return masterKey === OFFICE_PRIMARY_MASTER_KEY ? layoutId : `${masterKey}--${layoutId}`;
}

/** `<key>--<layout>` split into the inserted master and its own layout id; null for any other value. */
export function parseOfficeMasterLayoutRef(value: unknown): { masterKey: string; layoutId: string } | null {
  if (typeof value !== "string") return null;
  const match = OFFICE_MASTER_LAYOUT_REF.exec(value);
  return match ? { masterKey: match[1], layoutId: match[2] } : null;
}

/** The next free inserted-master key after `existing`: master-b, master-c, ..., master-z, master-ba. */
export function nextOfficeMasterKey(existing: readonly string[]): string {
  const taken = new Set(existing);
  for (let index = 1; index < 26 * 27; index += 1) {
    let n = index;
    let letters = "";
    do {
      letters = String.fromCharCode(97 + (n % 26)) + letters;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    const key = `master-${letters}`;
    if (!taken.has(key)) return key;
  }
  return "master-zz";
}

/** The master's placeholders the Master Layout dialog lists, in its order. */
export const OFFICE_MASTER_PLACEHOLDER_FAMILIES = [
  "title",
  "body",
  "date",
  "slideNumber",
  "footer",
] as const;
export type OfficeMasterPlaceholderFamily = (typeof OFFICE_MASTER_PLACEHOLDER_FAMILIES)[number];

/** A custom layout's id: `custom-` and letters (the layout-key alphabet). */
export const OFFICE_CUSTOM_LAYOUT_ID = /^custom-[a-z]{1,3}$/;
/** At most this many custom layouts on the master. */
export const OFFICE_MASTER_MAX_CUSTOM_LAYOUTS = 32;

export function isOfficeCustomLayoutId(value: unknown): value is string {
  return typeof value === "string" && OFFICE_CUSTOM_LAYOUT_ID.test(value);
}

/** The next free custom layout id after `existing`: custom-a, custom-b, ..., custom-aa. */
export function nextOfficeCustomLayoutId(existing: readonly string[]): string {
  const taken = new Set(existing);
  for (let index = 0; index < 18_278; index += 1) {
    let n = index;
    let letters = "";
    do {
      letters = String.fromCharCode(97 + (n % 26)) + letters;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    const id = `custom-${letters}`;
    if (!taken.has(id)) return id;
  }
  return "custom-zzz";
}

/**
 * The kinds of Master Layout ▸ Insert Placeholder, in PowerPoint's menu
 * order: Content, Content (Vertical), Text, Text (Vertical), Picture, Chart,
 * Table, SmartArt, Media, Online Image.
 */
export const OFFICE_PLACEHOLDER_KINDS = [
  "content",
  "content-vertical",
  "text",
  "text-vertical",
  "picture",
  "chart",
  "table",
  "smartart",
  "media",
  "online-image",
] as const;
export type OfficePlaceholderKind = (typeof OFFICE_PLACEHOLDER_KINDS)[number];

export interface OfficeMasterAddedPlaceholder {
  kind: OfficePlaceholderKind;
  /** `[x, y, cx, cy]` in EMU on the 16:9 catalog slide, like every frame here. */
  frame: OfficeMasterFrame;
}

/** At most this many placeholders added to one layout (PowerPoint has no small cap; bounded here). */
export const OFFICE_MASTER_MAX_ADDED_PLACEHOLDERS = 16;
/** A slot ref of {@link DeckOfficeMaster.layoutSlots}. */
export const OFFICE_MASTER_SLOT_REF = /^[ca](\d{1,2})$/;

export const OFFICE_MASTER_NAME_MAX_LENGTH = 200;

/** A master or layout name as PowerPoint keeps it: trimmed, one line, bounded. */
export function normalizeOfficeMasterName(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;
  const name = raw.replace(/[\u0000-\u001f\u007f]/g, " ").trim();
  return name && name.length <= OFFICE_MASTER_NAME_MAX_LENGTH ? name : undefined;
}

export type OfficeMasterFrame = [number, number, number, number];

const FRAME_LIMIT = 51_206_400;
const LAYOUT_FRAME_KEY = /^[a-z][a-z-]{0,40}$/;
const LAYOUT_KEYS_MAX = 64;

/** A frame as `[x, y, cx, cy]` integers in bounds, or undefined. */
export function normalizeOfficeMasterFrame(raw: unknown): OfficeMasterFrame | undefined {
  if (!Array.isArray(raw) || raw.length !== 4) return undefined;
  if (!raw.every((value) => typeof value === "number" && Number.isInteger(value))) return undefined;
  const [x, y, cx, cy] = raw as number[];
  if (Math.abs(x) > FRAME_LIMIT || Math.abs(y) > FRAME_LIMIT) return undefined;
  if (cx <= 0 || cy <= 0 || cx > FRAME_LIMIT || cy > FRAME_LIMIT) return undefined;
  return [x, y, cx, cy];
}

export const OFFICE_BACKGROUND_STYLE_COUNT = 12;

const HEX_COLOR = /^#[0-9a-f]{6}$/i;

/** One level's values, or undefined when it states nothing usable. */
export function normalizeOfficeMasterTextValues(
  raw: unknown,
): OfficeMasterTextValues | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const values: OfficeMasterTextValues = {};
  const size = record.fontSize;
  if (typeof size === "number" && Number.isFinite(size) && size >= 1 && size <= 4000) {
    values.fontSize = Math.round(size * 100) / 100;
  }
  if (typeof record.bold === "boolean") values.bold = record.bold;
  if (typeof record.italic === "boolean") values.italic = record.italic;
  if (typeof record.color === "string" && HEX_COLOR.test(record.color.trim())) {
    values.color = record.color.trim().toLowerCase();
  }
  return Object.keys(values).length > 0 ? values : undefined;
}

/** `theme.officeMaster` normalized to its fixed key set, or undefined when empty. */
export function normalizeDeckOfficeMaster(raw: unknown, nested?: boolean): DeckOfficeMaster | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const master: DeckOfficeMaster = {};
  const styles = record.textStyles;
  if (styles && typeof styles === "object" && !Array.isArray(styles)) {
    const textStyles: OfficeMasterTextStyles = {};
    const title = (styles as Record<string, unknown>).title;
    const titleOne =
      title && typeof title === "object" && !Array.isArray(title)
        ? normalizeOfficeMasterTextValues((title as Record<string, unknown>)["1"])
        : undefined;
    if (titleOne) textStyles.title = { "1": titleOne };
    const body = (styles as Record<string, unknown>).body;
    if (body && typeof body === "object" && !Array.isArray(body)) {
      const levels: Partial<Record<OfficeMasterBodyLevel, OfficeMasterTextValues>> = {};
      for (const level of OFFICE_MASTER_BODY_LEVELS) {
        const values = normalizeOfficeMasterTextValues((body as Record<string, unknown>)[level]);
        if (values) levels[level] = values;
      }
      if (Object.keys(levels).length > 0) textStyles.body = levels;
    }
    if (Object.keys(textStyles).length > 0) master.textStyles = textStyles;
  }
  if (record.background === true) master.background = true;
  const rawMasterFrames = record.masterFrames;
  if (rawMasterFrames && typeof rawMasterFrames === "object" && !Array.isArray(rawMasterFrames)) {
    const frames: Partial<Record<"title" | "body", OfficeMasterFrame>> = {};
    for (const family of ["title", "body"] as const) {
      const frame = normalizeOfficeMasterFrame((rawMasterFrames as Record<string, unknown>)[family]);
      if (frame) frames[family] = frame;
    }
    if (Object.keys(frames).length > 0) master.masterFrames = frames;
  }
  const rawLayoutFrames = record.layoutFrames;
  if (rawLayoutFrames && typeof rawLayoutFrames === "object" && !Array.isArray(rawLayoutFrames)) {
    const layouts: Record<string, Record<string, OfficeMasterFrame>> = {};
    for (const [layoutId, slots] of Object.entries(rawLayoutFrames as Record<string, unknown>)) {
      if (!LAYOUT_FRAME_KEY.test(layoutId) || !slots || typeof slots !== "object" || Array.isArray(slots)) continue;
      const frames: Record<string, OfficeMasterFrame> = {};
      for (const [slot, value] of Object.entries(slots as Record<string, unknown>)) {
        const frame = /^\d{1,2}$/.test(slot) ? normalizeOfficeMasterFrame(value) : undefined;
        if (frame) frames[slot] = frame;
      }
      if (Object.keys(frames).length > 0) layouts[layoutId] = frames;
    }
    if (Object.keys(layouts).length > 0) master.layoutFrames = layouts;
  }
  const name = normalizeOfficeMasterName(record.name);
  if (name) master.name = name;
  const rawLayoutNames = record.layoutNames;
  if (rawLayoutNames && typeof rawLayoutNames === "object" && !Array.isArray(rawLayoutNames)) {
    const names: Record<string, string> = {};
    for (const [layoutId, value] of Object.entries(rawLayoutNames as Record<string, unknown>)) {
      const layoutName = LAYOUT_FRAME_KEY.test(layoutId) ? normalizeOfficeMasterName(value) : undefined;
      if (layoutName) names[layoutId] = layoutName;
    }
    if (Object.keys(names).length > 0) master.layoutNames = names;
  }
  if (record.preserve === true) master.preserve = true;
  const rawAdded = record.addedPlaceholders;
  if (rawAdded && typeof rawAdded === "object" && !Array.isArray(rawAdded)) {
    const added: Record<string, OfficeMasterAddedPlaceholder[]> = {};
    for (const [layoutId, list] of Object.entries(rawAdded as Record<string, unknown>).slice(0, LAYOUT_KEYS_MAX)) {
      if (!LAYOUT_FRAME_KEY.test(layoutId) || !Array.isArray(list)) continue;
      const entries: OfficeMasterAddedPlaceholder[] = [];
      for (const item of list.slice(0, OFFICE_MASTER_MAX_ADDED_PLACEHOLDERS)) {
        const entry = item as Record<string, unknown> | null;
        const kind = entry?.kind;
        const frame = normalizeOfficeMasterFrame(entry?.frame);
        // A bad entry ends the list: later refs `a<n>` must keep naming the same entry.
        if (typeof kind !== "string" || !(OFFICE_PLACEHOLDER_KINDS as readonly string[]).includes(kind) || !frame) break;
        entries.push({ kind: kind as OfficePlaceholderKind, frame });
      }
      if (entries.length > 0) added[layoutId] = entries;
    }
    if (Object.keys(added).length > 0) master.addedPlaceholders = added;
  }
  const rawSlots = record.layoutSlots;
  if (rawSlots && typeof rawSlots === "object" && !Array.isArray(rawSlots)) {
    const slots: Record<string, string[]> = {};
    for (const [layoutId, list] of Object.entries(rawSlots as Record<string, unknown>).slice(0, LAYOUT_KEYS_MAX)) {
      if (!LAYOUT_FRAME_KEY.test(layoutId) || !Array.isArray(list)) continue;
      const addedCount = master.addedPlaceholders?.[layoutId]?.length ?? 0;
      const refs: string[] = [];
      for (const ref of list) {
        const match = typeof ref === "string" ? OFFICE_MASTER_SLOT_REF.exec(ref) : null;
        if (!match || refs.includes(ref as string)) continue;
        if ((ref as string).startsWith("a") && Number(match[1]) >= addedCount) continue;
        refs.push(ref as string);
      }
      slots[layoutId] = refs.slice(0, 32);
    }
    if (Object.keys(slots).length > 0) master.layoutSlots = slots;
  }
  for (const key of ["deletedLayouts", "layoutsWithoutFooters", "layoutsHidingBackgroundGraphics"] as const) {
    const raw = record[key];
    if (!Array.isArray(raw)) continue;
    const ids = [...new Set(raw.filter((id): id is string => typeof id === "string" && LAYOUT_FRAME_KEY.test(id)))]
      .slice(0, LAYOUT_KEYS_MAX)
      .sort();
    if (ids.length > 0) master[key] = ids;
  }
  const style = record.backgroundStyle;
  if (typeof style === "number" && Number.isInteger(style) && style >= 1 && style <= OFFICE_BACKGROUND_STYLE_COUNT) {
    master.backgroundStyle = style;
  }
  if (Array.isArray(record.masterPlaceholdersOff)) {
    const off = OFFICE_MASTER_PLACEHOLDER_FAMILIES.filter((family) =>
      (record.masterPlaceholdersOff as unknown[]).includes(family),
    );
    if (off.length > 0) master.masterPlaceholdersOff = off;
  }
  if (Array.isArray(record.customLayouts)) {
    // Insertion order is the thumbnail order: kept, not sorted.
    const custom = [...new Set(record.customLayouts.filter(isOfficeCustomLayoutId))].slice(
      0,
      OFFICE_MASTER_MAX_CUSTOM_LAYOUTS,
    );
    if (custom.length > 0) master.customLayouts = custom;
  }
  if (Array.isArray(record.layoutOrder)) {
    const order = [
      ...new Set(record.layoutOrder.filter((id): id is string => typeof id === "string" && LAYOUT_FRAME_KEY.test(id))),
    ].slice(0, LAYOUT_KEYS_MAX);
    if (order.length > 0) master.layoutOrder = order;
  }
  if (nested !== true && Array.isArray(record.additionalMasters)) {
    // Insert Slide Master's masters, in thumbnail order (PPT-B030). Each is
    // normalized like this record, without a background of its own or
    // further masters; one without a valid, unique key is dropped.
    const masters: OfficeAdditionalMaster[] = [];
    for (const item of record.additionalMasters.slice(0, OFFICE_MASTER_MAX_ADDITIONAL)) {
      const key = (item as Record<string, unknown> | null)?.key;
      if (!isOfficeAdditionalMasterKey(key) || masters.some((entry) => entry.key === key)) continue;
      const edits = normalizeDeckOfficeMaster(item, true) ?? {};
      delete edits.background;
      delete edits.backgroundStyle;
      masters.push({ key, ...edits });
    }
    if (masters.length > 0) master.additionalMasters = masters;
  }
  return Object.keys(master).length > 0 ? master : undefined;
}
