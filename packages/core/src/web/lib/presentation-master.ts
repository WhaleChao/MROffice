/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Deck-level MASTER LAYOUT inheritance (PowerPoint slide-master parity).
 *
 * A *master* is a NAMED, deck-level bundle of two things a slide can INHERIT:
 *   1. a layout id (the placeholder geometry template, see presentation-layouts),
 *   2. a set of style overrides (background / font / title+body colours) layered
 *      ON TOP of the deck theme.
 *
 * A slide references a master by `masterId`. Editing a master cascades to every
 * slide that uses it — unless the slide (or one of its blocks) supplies its own
 * explicit override. The full precedence (most specific wins) is therefore:
 *
 *   per-block / per-slide explicit override
 *     >  master override
 *     >  deck theme
 *     >  built-in default
 *
 * This module is the SINGLE home for the pure, side-effect-free resolution rules
 * so the zod schema, the editor canvas, and the read-only present-mode overlay
 * all agree. Keep it dependency-free of React / the DOM (it may import the sister
 * pure modules `presentation-theme` + `presentation-layouts`) so it can be unit
 * tested in isolation and imported from the schema.
 *
 * Everything here is ADDITIVE and OPTIONAL: a deck with no `masters` (and slides
 * with no `masterId`) resolves to EXACTLY the prior theme-only behaviour, so
 * existing decks render byte-for-byte the same as before this feature.
 */

import {
  DEFAULT_LAYOUT_ID,
  isLayoutId,
  normalizeLayoutId,
  type LayoutId,
} from "./presentation-layouts";
import {
  isNativePresentationThemePartPath,
  normalizeTheme,
  normalizeThemeColorMap,
  presentationBodyFontStyle,
  themeFontSchemeFamily,
  normalizeThemeColor,
  resolveBlockTextStyle,
  resolveSlideBackground,
  resolveTitleTextStyle,
  DECK_THEME_MAPPED_COLORS,
  type DeckTheme,
  type DeckThemeColorMap,
  type ResolvedBlockTextStyle,
} from "./presentation-theme";
import { SLIDE_BG_FIELDS } from "./presentation-slide-bg";
import { presentationMasterColorMap } from "./presentation/master-color-map";
import { officeThemeFontStack } from "./office-theme-fonts";

/** Max length for a master id / name (mirrors the schema's structural guard). */
const MASTER_ID_MAX_LENGTH = 200;
const MASTER_NAME_MAX_LENGTH = 200;
const MASTER_LAYOUT_ID_MAX_LENGTH = 512;
/** Hard cap on the number of masters a deck may carry. */
export const MAX_MASTERS = 50;
/** PowerPoint/WPS decks commonly carry 9–12 layouts per master. */
export const MAX_MASTER_LAYOUTS = 64;

/**
 * One native child layout owned by a slide master.
 *
 * `id` is an opaque stable reference. Imported PPTX decks use the OOXML part
 * path (for example `ppt/slideLayouts/slideLayout2.xml`); locally-created
 * layouts use a UUID-derived id. `builtInLayoutId` is only the editor preview
 * geometry and never replaces the native identity.
 */
export interface SlideMasterLayout {
  id: string;
  name: string;
  builtInLayoutId?: LayoutId;
  sourcePart?: string;
  nativeType?: string;
  matchingName?: string;
  preserve?: boolean;
  showMasterShapes?: boolean;
  userDrawn?: boolean;
  blocks?: Array<{ id?: string; [key: string]: unknown }>;
  /**
   * The MASTER's graphics as THIS layout's colour map draws them, present only
   * when the layout carries a real `a:overrideClrMapping` (ISO/IEC 29500-1
   * §20.1.6.5): the override "is used in place of the already defined color
   * mapping, or master color mapping", so a dark layout under a light master
   * draws the master's `schemeClr` shapes in other colours than the master
   * card does. Every other layout composes the master card's own blocks, which
   * is one copy for the whole deck (§73.4, §74).
   */
  masterGraphics?: Array<{ id?: string; [key: string]: unknown }>;
  /** The layout part's OWN `p:cSld/p:bg`, in the slide's background vocabulary. */
  background?: CatalogBackground;
  /**
   * The MASTER's background as THIS layout's colour map draws it, present
   * only when the layout has no `p:bg` of its own and carries a real
   * `a:overrideClrMapping`: PowerPoint paints a `bg1`/`bg2` master background
   * in the overriding colours on every slide of that layout (§75.3,
   * PDF-sampled), so the master card's own background is the wrong one here.
   */
  masterBackground?: CatalogBackground;
  /**
   * The background a slide that follows its layout's (no `p:bg` of its own:
   * `source.slideBackgroundFollowsLayout`) shows on THIS layout, in the
   * slide's own background fields exactly as the import projects that slide
   * (readings row 503). Empty: the deck theme's background. Home > Layout puts
   * it on such a slide at once ({@link slideWithNativeLayoutBackground}).
   */
  slideBackground?: Partial<Record<(typeof SLIDE_BG_FIELDS)[number], unknown>>;
  /** See {@link CatalogCard}. */
  catalogCard?: CatalogCard;
}

/**
 * Whether an imported layout's or master's card (its drawn blocks and
 * background, the thumbnail of the master view) came with the page: a
 * layout no slide uses, and a master none of whose layouts is used, is read
 * without it (`"pending"`), its card drawn by the server when it is first
 * shown (`lib/presentation-layout-cards`, `GET /pages/{id}/pptx-layout-card`);
 * the card the server drew names itself by `sha256`. Kept as read so a save
 * states it as the page holds it.
 */
export type CatalogCard =
  | "pending"
  | {
      readonly state: "drawn";
      readonly sha256: string;
      /** Drawn without its WMF/EMF pictures, which the server is still rendering. */
      readonly picturesPending?: true;
    };

export function normalizeCatalogCard(value: unknown): CatalogCard | undefined {
  if (value === "pending") return "pending";
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const { state, sha256, picturesPending } = value as { state?: unknown; sha256?: unknown; picturesPending?: unknown };
  return state === "drawn" && typeof sha256 === "string" && /^[a-f0-9]{64}$/.test(sha256)
    ? { state: "drawn", sha256, ...(picturesPending === true ? { picturesPending: true as const } : {}) }
    : undefined;
}

/**
 * A native master's / layout's own background (SHAPE-IMPORT-AUDIT.md §73):
 * exactly the keys a slide's background carries, read by the same projection
 * reader, so a card spreads it onto a slide-shaped preview and the slide
 * renderer draws it. An EMPTY object is meaningful: the part declares a
 * background the page draws as the theme's, which stops the card from
 * inheriting a master colour PowerPoint would not show. Absent means "declares
 * none" and the card inherits: layout → master → theme, PowerPoint's order.
 */
export interface CatalogBackground {
  backgroundColor?: string;
  backgroundGradient?: Record<string, unknown>;
  backgroundImage?: Record<string, unknown>;
  backgroundPattern?: Record<string, unknown>;
}

export const MASTER_TEXT_STYLE_NAMES = ["title", "body", "other"] as const;
export type MasterTextStyleName = (typeof MASTER_TEXT_STYLE_NAMES)[number];
export const MASTER_FONT_SCRIPTS = ["latin", "ea", "cs"] as const;
export type MasterFontScript = (typeof MASTER_FONT_SCRIPTS)[number];
export type MasterFontSlots = Partial<Record<MasterFontScript, string>>;
/**
 * `p:txStyles` has NINE levels and each `a:lvlNpPr/a:defRPr` may name its own
 * faces (ISO/IEC 29500-1 §19.3.1.52). Level one is the flat `latin`/`ea`/`cs`
 * keys; levels two to nine appear under `levels` only when they state a face
 * of their own -- a level that states none INHERITS, and inventing one there
 * would offer the user a font the master never chose (§74).
 */
export const MASTER_FONT_LEVEL_KEYS = ["2", "3", "4", "5", "6", "7", "8", "9"] as const;
export type MasterFontLevelKey = (typeof MASTER_FONT_LEVEL_KEYS)[number];
export interface MasterFontFaces extends MasterFontSlots {
  levels?: Partial<Record<MasterFontLevelKey, MasterFontSlots>>;
}

/**
 * What one `p:txStyles` level's `a:defRPr` STATES for size (points), bold,
 * italic and colour (§75.2). PowerPoint edits these from the Home tab with
 * that level's text selected in Slide Master view. Only stated attributes
 * appear: an absent one inherits.
 */
export interface MasterTextStyleLevelValues {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}
export const MASTER_TEXT_STYLE_LEVEL_KEYS = ["1", "2", "3", "4", "5", "6", "7", "8", "9"] as const;
export type MasterTextStyleLevelKey = (typeof MASTER_TEXT_STYLE_LEVEL_KEYS)[number];
export type MasterTextStyleLevels = Partial<
  Record<MasterTextStyleName, Partial<Record<MasterTextStyleLevelKey, MasterTextStyleLevelValues>>>
>;

function normalizeMasterTextStyleLevelValues(
  raw: unknown,
): MasterTextStyleLevelValues | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const values: MasterTextStyleLevelValues = {};
  const size = record.fontSize;
  if (typeof size === "number" && Number.isFinite(size) && size >= 1 && size <= 4000) {
    values.fontSize = Math.round(size * 100) / 100;
  }
  if (typeof record.bold === "boolean") values.bold = record.bold;
  if (typeof record.italic === "boolean") values.italic = record.italic;
  const color = normalizeThemeColor(record.color);
  if (color) values.color = color;
  return Object.keys(values).length > 0 ? values : undefined;
}

/** Normalize a persisted `textStyleLevels` map to its fixed key set (§75.2). */
export function normalizeMasterTextStyleLevels(
  raw: unknown,
): MasterTextStyleLevels | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const result: MasterTextStyleLevels = {};
  for (const name of MASTER_TEXT_STYLE_NAMES) {
    const style = record[name];
    if (!style || typeof style !== "object" || Array.isArray(style)) continue;
    const levels: Partial<Record<MasterTextStyleLevelKey, MasterTextStyleLevelValues>> = {};
    for (const key of MASTER_TEXT_STYLE_LEVEL_KEYS) {
      const values = normalizeMasterTextStyleLevelValues(
        (style as Record<string, unknown>)[key],
      );
      if (values) levels[key] = values;
    }
    if (Object.keys(levels).length > 0) result[name] = levels;
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

const MASTER_FONT_FACE_MAX_LENGTH = 200;

/** Whether a projected face is a theme font reference rather than a typeface. */
export function isThemeFontReference(face: string): boolean {
  return face.startsWith("+");
}

function normalizeMasterFontSlots(raw: unknown): MasterFontSlots | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const faces: MasterFontSlots = {};
  for (const script of MASTER_FONT_SCRIPTS) {
    const value = record[script];
    if (typeof value !== "string") continue;
    const face = value.trim();
    if (!face || face.length > MASTER_FONT_FACE_MAX_LENGTH || /[<\n]/.test(face)) continue;
    faces[script] = face;
  }
  return Object.keys(faces).length > 0 ? faces : undefined;
}

function normalizeMasterFontFaces(
  raw: unknown,
  withLevels: boolean,
): MasterFontFaces | undefined {
  const slots = normalizeMasterFontSlots(raw);
  if (!withLevels) return slots;
  const record = raw as Record<string, unknown> | null;
  const rawLevels =
    record && typeof record === "object" && !Array.isArray(record)
      ? record.levels
      : undefined;
  let levels: Partial<Record<MasterFontLevelKey, MasterFontSlots>> | undefined;
  if (rawLevels && typeof rawLevels === "object" && !Array.isArray(rawLevels)) {
    const source = rawLevels as Record<string, unknown>;
    for (const key of MASTER_FONT_LEVEL_KEYS) {
      const level = normalizeMasterFontSlots(source[key]);
      if (!level) continue;
      levels = { ...(levels ?? {}), [key]: level };
    }
  }
  if (!slots && !levels) return undefined;
  return { ...(slots ?? {}), ...(levels ? { levels } : {}) };
}

/** Normalize a persisted `textStyleFonts` / `themeFonts` map to its fixed key set. */
export function normalizeMasterFontMap<K extends string>(
  raw: unknown,
  names: readonly K[],
  options: { levels?: boolean } = {},
): Partial<Record<K, MasterFontFaces>> | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const map: Partial<Record<K, MasterFontFaces>> = {};
  for (const name of names) {
    const faces = normalizeMasterFontFaces(record[name], options.levels === true);
    if (faces) map[name] = faces;
  }
  return Object.keys(map).length > 0 ? map : undefined;
}

const CATALOG_BACKGROUND_OBJECT_KEYS = [
  "backgroundGradient",
  "backgroundImage",
  "backgroundPattern",
] as const;

/** Normalize a persisted catalog background; `undefined` when it is not one. */
export function normalizeCatalogBackground(
  raw: unknown,
): CatalogBackground | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const background: CatalogBackground = {};
  if (typeof record.backgroundColor === "string") {
    const color = normalizeThemeColor(record.backgroundColor);
    if (color) background.backgroundColor = color;
  }
  for (const key of CATALOG_BACKGROUND_OBJECT_KEYS) {
    const value = record[key];
    if (value && typeof value === "object" && !Array.isArray(value)) {
      background[key] = value as Record<string, unknown>;
    }
  }
  return background;
}

/**
 * One deck-level master. `id` is the stable reference a slide stores in
 * `masterId`. `name` is the human-facing label shown in the master editor.
 * `layoutId` (optional) names the placeholder geometry template a slide inherits
 * when it has none of its own. `theme` (optional) is the master's style override
 * bundle, layered on top of the deck theme. Both are optional so a bare master
 * (just an id + name) is valid and contributes nothing beyond grouping.
 */
export interface SlideMaster {
  id: string;
  name: string;
  /** See {@link CatalogCard}. */
  catalogCard?: CatalogCard;
  /** Original OOXML slide-master part path when this master was imported. */
  sourcePart?: string;
  /** Native OOXML theme part id inherited by this master. */
  themeId?: string;
  /** PowerPoint/WPS "Preserve master" semantic. */
  preserve?: boolean;
  /** Native child layouts retained in authoring order. */
  layouts?: SlideMasterLayout[];
  /** The master part's OWN `p:cSld/p:bg`, in the slide's background vocabulary. */
  background?: CatalogBackground;
  /**
   * The faces level one of each `p:txStyles` style names for `a:latin` /
   * `a:ea` / `a:cs`, verbatim (§73, Replace Fonts). A theme reference
   * (`+mj-lt`, `+mn-ea` ...) stays a reference: it is the THEME's font
   * (Design ▸ Fonts), which Replace Fonts leaves alone; only a literal face
   * here is the master's own and may be replaced.
   */
  textStyleFonts?: Partial<Record<MasterTextStyleName, MasterFontFaces>>;
  /** The theme font scheme those references resolve through; read-only. */
  themeFonts?: Partial<Record<"major" | "minor", MasterFontFaces>>;
  /**
   * The master theme's Hyperlink / Followed Hyperlink colours, stated by the
   * import when the deck's masters name different ones; read-only. The slide
   * surface draws its links in `hlink` (`presentationHyperlinkStyle`).
   */
  linkColors?: { hlink?: string; folHlink?: string };
  /**
   * The master's own colour map (`p:clrMap`: Background 1 / Text 1 /
   * Background 2 / Text 2 to a theme slot) and the colour each names in the
   * master's own theme, stated by the import when the deck's masters do not
   * all draw them alike (reading row 850); read-only. A slide on the master
   * draws its unstated words and Text 1 through them
   * (`presentationSlideInk`, `presentationThemeGridColors`), while the deck's
   * theme is still the file's own.
   */
  colorMap?: DeckThemeColorMap;
  mappedColors?: Readonly<Record<(typeof DECK_THEME_MAPPED_COLORS)[number], string>>;
  /** Size / bold / italic / colour each `p:txStyles` level states (§75.2). */
  textStyleLevels?: MasterTextStyleLevels;
  /** Inherited layout id (placeholder geometry). Absent → no layout contribution. */
  layoutId?: LayoutId;
  /** Master style overrides layered on top of the deck theme. */
  theme?: DeckTheme;
  /**
   * Placeholder blocks edited ON the master that propagate to every slide using
   * it (PowerPoint master text). A slide overrides an inherited placeholder by
   * providing its own block of the same id — see lib/presentation-master-cascade
   * `resolveSlideRenderBlocks`. Absent/empty → the master contributes no text.
   */
  blocks?: Array<{ id?: string; [key: string]: unknown }>;
}

function normalizeMasterMappedColors(
  raw: unknown,
): SlideMaster["mappedColors"] | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  if (Object.keys(record).length !== DECK_THEME_MAPPED_COLORS.length) return undefined;
  const result = {} as Record<(typeof DECK_THEME_MAPPED_COLORS)[number], string>;
  for (const name of DECK_THEME_MAPPED_COLORS) {
    const value = record[name];
    if (typeof value !== "string" || !/^#[0-9a-f]{6}$/i.test(value)) return undefined;
    result[name] = value.toLowerCase();
  }
  return result;
}

/**
 * Normalize a single persisted value into a clean {@link SlideMaster}, or
 * `undefined` when it has no usable id. Keeps only valid fields:
 *  - `id` is required (trimmed, length-capped); a missing id ⇒ undefined.
 *  - `name` falls back to the id when absent / blank so the editor always has a
 *    label.
 *  - `layoutId` is dropped unless it is a KNOWN layout id (so a stale value never
 *    references a missing layout). Note: unlike `normalizeLayoutId`, an unknown
 *    value here is OMITTED (not coerced to `blank`) so "no layout contribution"
 *    stays distinct from "explicitly the blank layout".
 *  - `theme` is run through `normalizeTheme`, dropping invalid colours / fonts.
 */
export function normalizeMaster(
  raw: unknown,
  normalizeBlock?: (block: unknown, index: number) => unknown,
): SlideMaster | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const id = normalizeMasterId(record.id);
  if (!id) return undefined;
  const master: SlideMaster = { id, name: id };
  const name =
    typeof record.name === "string" ? record.name.trim() : "";
  if (name) master.name = name.slice(0, MASTER_NAME_MAX_LENGTH);
  const sourcePart = normalizeOoxmlPartPath(
    record.sourcePart,
    "ppt/slideMasters/",
  );
  if (sourcePart) master.sourcePart = sourcePart;
  const rawThemeId =
    typeof record.themeId === "string" ? record.themeId.trim() : "";
  if (isNativePresentationThemePartPath(rawThemeId)) {
    master.themeId = rawThemeId;
  }
  if (record.preserve === true) master.preserve = true;
  const background = normalizeCatalogBackground(record.background);
  if (background) master.background = background;
  const textStyleFonts = normalizeMasterFontMap(
    record.textStyleFonts,
    MASTER_TEXT_STYLE_NAMES,
    { levels: true },
  );
  if (textStyleFonts) master.textStyleFonts = textStyleFonts;
  const themeFonts = normalizeMasterFontMap(record.themeFonts, ["major", "minor"] as const);
  if (themeFonts) master.themeFonts = themeFonts;
  if (record.linkColors && typeof record.linkColors === "object" && !Array.isArray(record.linkColors)) {
    const links = record.linkColors as Record<string, unknown>;
    const linkColors: { hlink?: string; folHlink?: string } = {};
    for (const slot of ["hlink", "folHlink"] as const) {
      const value = links[slot];
      if (typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value)) linkColors[slot] = value.toLowerCase();
    }
    if (linkColors.hlink || linkColors.folHlink) master.linkColors = linkColors;
  }
  // Both or neither: a map without its colours (or the colours without the
  // map) cannot be drawn exactly, and the deck theme's is then the one.
  const colorMap = normalizeThemeColorMap(record.colorMap);
  const mappedColors = normalizeMasterMappedColors(record.mappedColors);
  if (colorMap && mappedColors) {
    master.colorMap = colorMap;
    master.mappedColors = mappedColors;
  }
  const textStyleLevels = normalizeMasterTextStyleLevels(record.textStyleLevels);
  if (textStyleLevels) master.textStyleLevels = textStyleLevels;
  // Keep ONLY a known layout id; an unknown / stale value is omitted entirely so
  // the master simply contributes no geometry (rather than silently forcing
  // `blank`). A value that IS a recognized synonym is canonicalized.
  if (record.layoutId !== undefined) {
    const normalized = normalizeLayoutId(record.layoutId);
    // `normalizeLayoutId` collapses garbage to `blank`; only retain the result
    // when the raw value was itself a real layout id / known synonym (i.e. the
    // normalized value is not just the garbage fallback). We detect this by
    // re-checking: if the raw trimmed string is a known id OR the normalization
    // changed it to a non-default id, keep it. A raw value that maps to `blank`
    // ONLY because it was unrecognized is dropped.
    if (masterLayoutIsMeaningful(record.layoutId, normalized)) {
      master.layoutId = normalized;
    }
  }
  const theme = normalizeTheme(record.theme);
  if (theme) master.theme = theme;
  const layouts = normalizeMasterLayouts(record.layouts, normalizeBlock);
  if (layouts.length > 0) master.layouts = layouts;
  // Placeholder blocks: when a per-block normalizer is supplied (the deck-load
  // path passes the SAME normalizePresentationBlock that slide blocks run
  // through), each block is sanitized so a malformed/forged block — e.g. a table
  // whose `rows` is not an array — can't reach the renderer and crash deck load.
  // Without it the array passes through verbatim, so a master without master text
  // stays byte-for-byte unchanged. Only a non-empty array is kept.
  if (Array.isArray(record.blocks) && record.blocks.length > 0) {
    const rawBlocks = record.blocks;
    master.blocks = (
      normalizeBlock
        ? rawBlocks.map((block, index) => normalizeBlock(block, index))
        : rawBlocks
    ) as SlideMaster["blocks"];
  }
  const catalogCard = normalizeCatalogCard(record.catalogCard);
  if (catalogCard) master.catalogCard = catalogCard;
  return master;
}

/**
 * Decide whether a normalized layout id should be RETAINED on a master.
 * `normalizeLayoutId` maps every unknown value to `blank`, which would make an
 * accidental garbage value indistinguishable from an explicit `blank`. We keep
 * the layout only when the raw value genuinely denotes a layout:
 *  - it is exactly a known layout id (after trim/lowercase), or
 *  - it normalized to a NON-default id (i.e. a recognized synonym), or
 *  - it is one of the explicit blank synonyms.
 */
function masterLayoutIsMeaningful(
  raw: unknown,
  normalized: LayoutId,
): boolean {
  if (typeof raw !== "string") return false;
  const candidate = raw.trim().toLowerCase();
  if (candidate === "") return false;
  if (isLayoutId(candidate)) return true;
  if (normalized !== DEFAULT_LAYOUT_ID) return true;
  // Explicit blank synonyms (otherwise an unrecognized string also maps to
  // `blank` and we must NOT keep it).
  return (
    candidate === "empty" ||
    candidate === "none" ||
    candidate === "blank"
  );
}

/**
 * Normalize an arbitrary persisted master id into a clean, safe string, or
 * `undefined` when missing / unusable. Trims surrounding whitespace and caps
 * length so a stale or hand-edited value can never be unbounded. Returns the id
 * VERBATIM (case preserved) — ids are opaque references, not enums.
 */
export function normalizeMasterId(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.slice(0, MASTER_ID_MAX_LENGTH);
}

/** Normalize an opaque native/local slide-layout reference. */
export function normalizeMasterLayoutId(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.slice(0, MASTER_LAYOUT_ID_MAX_LENGTH);
}

function normalizeOoxmlPartPath(
  value: unknown,
  prefix: string,
): string | undefined {
  const normalized = normalizeMasterLayoutId(value);
  if (
    !normalized ||
    !normalized.startsWith(prefix) ||
    !normalized.endsWith(".xml") ||
    normalized.split("/").includes("..")
  ) {
    return undefined;
  }
  return normalized;
}

/** Normalize one native/local child layout without flattening its identity. */
export function normalizeMasterLayout(
  raw: unknown,
  normalizeBlock?: (block: unknown, index: number) => unknown,
): SlideMasterLayout | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const id = normalizeMasterLayoutId(record.id);
  if (!id) return undefined;
  const rawName = typeof record.name === "string" ? record.name.trim() : "";
  const layout: SlideMasterLayout = {
    id,
    name: (rawName || id).slice(0, MASTER_NAME_MAX_LENGTH),
  };
  if (record.builtInLayoutId !== undefined) {
    const normalized = normalizeLayoutId(record.builtInLayoutId);
    if (masterLayoutIsMeaningful(record.builtInLayoutId, normalized)) {
      layout.builtInLayoutId = normalized;
    }
  }
  const sourcePart = normalizeOoxmlPartPath(
    record.sourcePart ?? record.id,
    "ppt/slideLayouts/",
  );
  if (sourcePart) layout.sourcePart = sourcePart;
  for (const key of ["nativeType", "matchingName"] as const) {
    const value = typeof record[key] === "string" ? record[key].trim() : "";
    if (value) layout[key] = value.slice(0, MASTER_NAME_MAX_LENGTH);
  }
  if (record.preserve === true) layout.preserve = true;
  if (record.showMasterShapes === false) layout.showMasterShapes = false;
  if (record.userDrawn === true) layout.userDrawn = true;
  const background = normalizeCatalogBackground(record.background);
  if (background) layout.background = background;
  const masterBackground = normalizeCatalogBackground(record.masterBackground);
  if (masterBackground) layout.masterBackground = masterBackground;
  const slideBackground = record.slideBackground;
  if (slideBackground && typeof slideBackground === "object" && !Array.isArray(slideBackground)) {
    // As the import wrote it (a slide's own fields): the save compares the
    // slide's background with what the package projects, value for value.
    const fields: NonNullable<SlideMasterLayout["slideBackground"]> = {};
    for (const field of SLIDE_BG_FIELDS) {
      const value = (slideBackground as Record<string, unknown>)[field];
      if (value !== undefined) fields[field] = value;
    }
    layout.slideBackground = fields;
  }
  for (const key of ["blocks", "masterGraphics"] as const) {
    const blocks = record[key];
    if (!Array.isArray(blocks) || blocks.length === 0) continue;
    layout[key] = (
      normalizeBlock
        ? blocks.map((block, index) => normalizeBlock(block, index))
        : blocks
    ) as SlideMasterLayout["blocks"];
  }
  const catalogCard = normalizeCatalogCard(record.catalogCard);
  if (catalogCard) layout.catalogCard = catalogCard;
  return layout;
}

/**
 * The blocks a slide INHERITS from its child layout.
 *
 * A layout the editor authored hands its `blocks` to every slide that uses it
 * (`resolveSlideRenderBlocks`). A NATIVE layout — one that names a package
 * part — hands them to nobody: an imported slide was projected with everything
 * it inherits already on it (`source.inheritedFrom`), and PowerPoint draws a
 * layout's placeholders on a slide only where the slide has one of its own. So
 * a native layout's `blocks` are the import's read-only preview of the layout
 * part, drawn by the master view and the Design gallery and nowhere else.
 */
export function layoutBlocksInheritedBySlides(
  layout: SlideMasterLayout | undefined,
): SlideMasterLayout["blocks"] {
  if (!layout || layout.sourcePart) return undefined;
  return layout.blocks;
}

type CatalogBlock = { id?: string; [key: string]: unknown };

/**
 * Whether a master's / layout's block is the import's read-only preview of
 * the part (stamped `source.inheritedFrom` by the projection) rather than
 * content the editor authored on it (the master-text block carries no
 * `source` at all).
 */
export function isCatalogPreviewBlock(block: unknown): boolean {
  if (!block || typeof block !== "object") return false;
  const source = (block as { source?: unknown }).source;
  if (!source || typeof source !== "object") return false;
  const origin = (source as { inheritedFrom?: unknown }).inheritedFrom;
  return origin === "master" || origin === "layout";
}

/**
 * Whether a SLIDE block is a layout's or master's own graphic the import drew
 * on the slide (stamped `source.inheritedFrom`, as the catalog's preview
 * blocks are) -- a decoration, not a placeholder. PowerPoint draws such a
 * shape on the slide but never selects it there: a click, a marquee, Tab and
 * Select All pass it by, and Slide Master view edits it (the CDC deck's title
 * band: a click selected its six layout shapes, with no frame to show for
 * them). An inherited placeholder stays selectable: formatting it makes it
 * the slide's own (`materializedPlaceholder`).
 */
export function isLayoutGraphicOnSlide(block: unknown): boolean {
  if (!isCatalogPreviewBlock(block)) return false;
  const record = block as { placeholder?: unknown; placeholderType?: unknown; source?: { placeholderType?: unknown } };
  return !record.placeholder && !record.placeholderType && !record.source?.placeholderType;
}

/**
 * The blocks a slide INHERITS from its master (`resolveSlideRenderBlocks`).
 *
 * The editor's own master text cascades to every slide using the master. A
 * NATIVE master's card (§73) shares the same `blocks` array, and those preview
 * blocks cascade to nobody: an imported slide was projected with the master
 * graphics it shows already on it. The two are told apart by the stamp, not
 * by the master's identity, so a master-text block written onto an imported
 * master keeps cascading exactly as before.
 */
export function masterBlocksInheritedBySlides(
  master: SlideMaster | undefined,
): SlideMaster["blocks"] {
  const blocks = master?.blocks?.filter((block) => !isCatalogPreviewBlock(block));
  return blocks && blocks.length > 0 ? blocks : undefined;
}

/**
 * The master graphics a LAYOUT's card draws under its own blocks: the
 * master card's preview blocks that are not placeholders. PowerPoint draws a
 * layout with the master's graphics behind it but never the master's
 * placeholders (the layout has its own), and not at all when the layout says
 * `showMasterSp="0"`. Projected once on the master and composed here, instead
 * of copied into every layout's preview.
 */
export function masterGraphicsInheritedByLayout(
  master: SlideMaster | undefined,
  layout: SlideMasterLayout | undefined,
): CatalogBlock[] {
  if (!layout || layout.showMasterShapes === false) return [];
  // A layout that really re-points the scheme names draws the master's
  // graphics in its OWN colours, and the projection gave it its own copy
  // (§74). Everything else composes the one copy on the master card.
  const source = layout.masterGraphics ?? master?.blocks;
  if (!source) return [];
  return source.filter(
    (block) =>
      isCatalogPreviewBlock(block) &&
      !Object.prototype.hasOwnProperty.call(block, "placeholderType"),
  );
}

/** The master / layout record a slide names (`masterId` / `layoutId`: part paths of an import). */
function nativeSlideLayoutRecords(
  masters: readonly SlideMaster[] | undefined,
  masterId: unknown,
  layoutId: unknown,
): { master: SlideMaster; layout: SlideMasterLayout } | null {
  // Not `Array.isArray`: on a readonly array it narrows to `any[]`.
  if (!masters || typeof layoutId !== "string" || !layoutId) return null;
  for (const master of masters) {
    if (typeof masterId === "string" && master.id !== masterId && master.sourcePart !== masterId) continue;
    const layout = master.layouts?.find((entry) => entry.id === layoutId || entry.sourcePart === layoutId);
    if (layout) return { master, layout };
  }
  return null;
}

/**
 * The graphics a slide on a native layout shows from it, PowerPoint's order
 * (ECMA-376 §19.3.1.38 `sld`, §19.3.1.39 `sldLayout` `showMasterSp`): the
 * master's non-placeholder shapes, unless the layout says `showMasterSp="0"`
 * (as the layout's colour map draws them when it re-points the scheme), then
 * the layout's own non-placeholder shapes -- what the import draws on every
 * slide of that layout (`source.inheritedFrom`). Null while the layout's or the
 * master's card is not here (`catalogCard: "pending"`): its drawing is unknown.
 */
export function nativeLayoutGraphicsForSlide(
  masters: readonly SlideMaster[] | undefined,
  masterId: unknown,
  layoutId: unknown,
): CatalogBlock[] | null {
  const records = nativeSlideLayoutRecords(masters, masterId, layoutId);
  if (!records || !records.layout.sourcePart) return null;
  const { master, layout } = records;
  // A card not drawn yet, or drawn while its pictures were still rendering
  // (`picturesPending`: no blocks yet), does not say what the layout draws.
  const unknown = (card: CatalogCard | undefined) =>
    card === "pending" || (typeof card === "object" && card.picturesPending === true);
  if (unknown(layout.catalogCard)) return null;
  if (layout.showMasterShapes !== false && !layout.masterGraphics && unknown(master.catalogCard)) return null;
  return [
    ...masterGraphicsInheritedByLayout(master, layout),
    ...(layout.blocks ?? []).filter((block) => isLayoutGraphicOnSlide(block)),
  ];
}














export function slideWithNativeLayoutGraphics<S extends {
  id?: unknown;
  layoutId?: unknown;
  masterId?: unknown;
  source?: unknown;
  hideBackgroundGraphics?: unknown;
  blocks?: unknown;
}>(slide: S, masters: readonly SlideMaster[] | undefined): S {
  const graphics = nativeLayoutGraphicsForSlide(masters, slide.masterId, slide.layoutId);
  if (!graphics || !Array.isArray(slide.blocks)) return slide;
  const source =
    slide.source && typeof slide.source === "object" && !Array.isArray(slide.source)
      ? (slide.source as Record<string, unknown>)
      : {};
  const hidden =
    slide.hideBackgroundGraphics === true || source.slideMasterShapesOriginalShow === false;
  const slideSourceId = typeof source.slideId === "string" ? source.slideId : null;
  const slideId = typeof slide.id === "string" ? slide.id : "slide";
  const own = (slide.blocks as unknown[]).filter((block) => !isLayoutGraphicOnSlide(block));
  const counts = { theme: 0, layout: 0 };
  const groupNames = { theme: new Map<string, string>(), layout: new Map<string, string>() };
  const drawn = hidden
    ? []
    : graphics.map((block) => {
        const copy = JSON.parse(JSON.stringify(block)) as CatalogBlock & { groupPath?: unknown };
        delete copy.zIndex;
        delete copy.sourceOrder;
        const blockSource = (copy.source ?? {}) as Record<string, unknown>;
        // The ids and group names the save reads a relaid slide's graphics
        // under (`pptx_theme_settle.settled_decoration_id`, `_expected_graphics`):
        // `<slide>-theme-<n>` for the master's, `<slide>-layout-<n>` for the
        // layout's, each group `<slide>-<theme|layout>-group-<n>` in order of
        // first use. Under any other id the save took them for the slide's own
        // shapes and declined them (`inherited:`), as the page drew them.
        const label = blockSource.inheritedFrom === "master" ? "theme" : "layout";
        const groups = groupNames[label];
        const group = (name: unknown): unknown => {
          if (typeof name !== "string" || !name) return name;
          if (!groups.has(name)) groups.set(name, `${slideId}-${label}-group-${groups.size + 1}`);
          return groups.get(name);
        };
        if ("groupId" in copy) copy.groupId = group(copy.groupId);
        if (Array.isArray(copy.groupPath)) copy.groupPath = copy.groupPath.map(group);
        counts[label] += 1;
        return {
          ...copy,
          id: `${slideId}-${label}-${counts[label]}`,
          source: { ...blockSource, ...(slideSourceId ? { slideId: slideSourceId } : {}) },
        };
      });
  return { ...slide, blocks: [...drawn, ...own] };
}

/** A slide's background fields, comparable (`SLIDE_BG_FIELDS`, keys sorted). */
function slideBackgroundSignature(record: Record<string, unknown>): string {
  const sorted = (value: unknown): unknown =>
    Array.isArray(value)
      ? value.map(sorted)
      : value && typeof value === "object"
        ? Object.fromEntries(
            Object.keys(value as Record<string, unknown>)
              .sort()
              .map((key) => [key, sorted((value as Record<string, unknown>)[key])]),
          )
        : value;
  return JSON.stringify(SLIDE_BG_FIELDS.map((field) => sorted(record[field]) ?? null));
}

/**
 * A native slide that changed layout (Home > Layout) with the new layout's
 * background, as PowerPoint draws it at once (`Slide.FollowMasterBackground`):
 * a slide with no background of its own shows its layout's, else its
 * master's, so it takes the NEW layout's; one with its own keeps it.
 *
 * `before` is the slide on its old layout, `after` the same slide on the new
 * one. The slide follows when the import said so
 * (`source.slideBackgroundFollowsLayout`, readings row 503) AND it still
 * shows what its old layout gives a follower (the old card's
 * `slideBackground`): Format Background since gave it one of its own, which
 * the save writes on the slide, and that one stays. A background equal to the
 * layout's is no background of its own to the save either: it writes none,
 * and the slide follows in the file as it does here. `after` itself when
 * either card is not here (`catalogCard: "pending"`), or when `after`'s
 * background is no longer `before`'s.
 */
export function slideWithNativeLayoutBackground<S extends {
  layoutId?: unknown;
  masterId?: unknown;
  source?: unknown;
}>(before: S, after: S, masters: readonly SlideMaster[] | undefined): S {
  const source =
    before.source && typeof before.source === "object" && !Array.isArray(before.source)
      ? (before.source as Record<string, unknown>)
      : null;
  if (source?.slideBackgroundFollowsLayout !== true || before.layoutId === after.layoutId) return after;
  const old = nativeSlideLayoutRecords(masters, before.masterId, before.layoutId)?.layout.slideBackground;
  const next = nativeSlideLayoutRecords(masters, after.masterId, after.layoutId)?.layout.slideBackground;
  if (!old || !next) return after;
  const beforeFields = before as unknown as Record<string, unknown>;
  const afterFields = after as unknown as Record<string, unknown>;
  const shown = slideBackgroundSignature(beforeFields);
  if (shown !== slideBackgroundSignature(afterFields) || shown !== slideBackgroundSignature(old)) return after;
  if (slideBackgroundSignature(next) === shown) return after;
  const result = { ...afterFields };
  for (const field of SLIDE_BG_FIELDS) delete result[field];
  for (const field of SLIDE_BG_FIELDS) {
    if (next[field] !== undefined) result[field] = JSON.parse(JSON.stringify(next[field]));
  }
  return result as unknown as S;
}

/**
 * The background a catalog card draws, PowerPoint's order: the layout's own,
 * else the master's -- as the layout's colour map draws it when the layout
 * re-points the scheme (`masterBackground`), else as the master's own map
 * does -- else nothing here (the deck theme's, which the slide renderer
 * already falls back to). An explicit empty object stops the walk.
 */
export function catalogPreviewBackground(
  master: SlideMaster | undefined,
  layout?: SlideMasterLayout,
): CatalogBackground {
  return layout?.background ?? layout?.masterBackground ?? master?.background ?? {};
}

/** Normalize, de-duplicate, and cap a master's native child layouts. */
export function normalizeMasterLayouts(
  raw: unknown,
  normalizeBlock?: (block: unknown, index: number) => unknown,
): SlideMasterLayout[] {
  if (!Array.isArray(raw)) return [];
  const layouts: SlideMasterLayout[] = [];
  const seen = new Set<string>();
  for (const entry of raw) {
    if (layouts.length >= MAX_MASTER_LAYOUTS) break;
    const layout = normalizeMasterLayout(entry, normalizeBlock);
    if (!layout || seen.has(layout.id)) continue;
    seen.add(layout.id);
    layouts.push(layout);
  }
  return layouts;
}

/**
 * Normalize an arbitrary persisted `masters` value into a clean, de-duplicated
 * array of {@link SlideMaster}. Non-array / missing input yields `[]`. Entries
 * without a usable id are dropped; duplicate ids keep the FIRST occurrence
 * (later duplicates are discarded so a `masterId` reference is unambiguous). The
 * result is capped at {@link MAX_MASTERS}.
 */
export function normalizeMasters(
  raw: unknown,
  normalizeBlock?: (block: unknown, index: number) => unknown,
): SlideMaster[] {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const masters: SlideMaster[] = [];
  for (const entry of raw) {
    if (masters.length >= MAX_MASTERS) break;
    const master = normalizeMaster(entry, normalizeBlock);
    if (!master) continue;
    if (seen.has(master.id)) continue;
    seen.add(master.id);
    masters.push(master);
  }
  return masters;
}

/**
 * Resolve the master a slide INHERITS from, or `undefined` when the slide has no
 * (valid) `masterId` — in which case the slide keeps the prior theme-only
 * behaviour. An id that does not match any master in `masters` falls back to
 * `undefined` (so a stale / dangling reference degrades gracefully rather than
 * crashing or silently picking the wrong master).
 */
export function resolveSlideMaster(
  slide: { masterId?: unknown; layoutId?: unknown } | null | undefined,
  masters: readonly SlideMaster[] | null | undefined,
): SlideMaster | undefined {
  const id = normalizeMasterId(slide?.masterId);
  if (!masters || masters.length === 0) return undefined;
  if (id) {
    const explicit = masters.find((master) => master.id === id);
    if (explicit) return explicit;
  }
  const layoutId = normalizeMasterLayoutId(slide?.layoutId);
  if (!layoutId) return undefined;
  return masters.find((master) =>
    master.layouts?.some((layout) => layout.id === layoutId),
  );
}

/** Resolve the slide's native child layout, retaining its stable OOXML id. */
export function resolveSlideNativeLayout(
  slide: { masterId?: unknown; layoutId?: unknown } | null | undefined,
  masters: readonly SlideMaster[] | null | undefined,
): SlideMasterLayout | undefined {
  const layoutId = normalizeMasterLayoutId(slide?.layoutId);
  if (!layoutId || !masters || masters.length === 0) return undefined;
  const master = resolveSlideMaster(slide, masters);
  if (master) {
    const match = master.layouts?.find((layout) => layout.id === layoutId);
    if (match) return match;
  }
  for (const candidate of masters) {
    const match = candidate.layouts?.find((layout) => layout.id === layoutId);
    if (match) return match;
  }
  return undefined;
}

/**
 * Layer a master's style overrides ON TOP of the deck theme, producing the
 * EFFECTIVE theme a slide/block should resolve against. Field-by-field, a
 * present master field wins over the same deck-theme field; an absent master
 * field falls through to the deck theme. The result is re-normalized so it is
 * always render-safe. Returns `undefined` only when BOTH inputs contribute
 * nothing (so a master-less, theme-less deck stays byte-clean).
 */
export function mergeMasterTheme(
  theme: DeckTheme | null | undefined,
  master: SlideMaster | null | undefined,
): DeckTheme | undefined {
  const base = normalizeTheme(theme);
  const overlay = master ? normalizeTheme(master.theme) : undefined;
  if (!base && !overlay) return undefined;
  // A background is ONE fill (solid, gradient, picture or pattern): a master
  // that states its own replaces the deck's whole fill. Field by field, a
  // reused slide's white source background lost to the destination's
  
  const ownFill =
    overlay !== undefined &&
    MASTER_BACKGROUND_FIELDS.some((field) => overlay[field] !== undefined);
  const under: DeckTheme = { ...(base ?? {}) };
  if (ownFill) for (const field of MASTER_BACKGROUND_FIELDS) delete under[field];
  // The master's own colour map (reading row 850): its slides' Text 1 and
  // Background 1 name the slots it maps, as PowerPoint draws them, while the
  // deck still sits on the file's masters.
  const ownMap = presentationMasterColorMap(base, master);
  return normalizeTheme({ ...under, ...(overlay ?? {}), ...(ownMap ? { colorMap: ownMap } : {}) });
}











export function presentationSlideChartBodyFont(
  slideTheme: DeckTheme | null | undefined,
  master: Pick<SlideMaster, "themeFonts"> | null | undefined,
): string | undefined {
  if (themeFontSchemeFamily(slideTheme, false)) return presentationBodyFontStyle(slideTheme)["--presentation-body-font"];
  const minor = master?.themeFonts?.minor;
  const latin = typeof minor?.latin === "string" ? minor.latin.trim() : "";
  if (latin && !latin.startsWith("+")) {
    const ea = typeof minor?.ea === "string" && !minor.ea.startsWith("+") ? minor.ea : "";
    return officeThemeFontStack({ latin, ea, scripts: {} });
  }
  return presentationBodyFontStyle(slideTheme)["--presentation-body-font"];
}

const MASTER_BACKGROUND_FIELDS = [
  "background",
  "backgroundGradient",
  "backgroundImage",
  "backgroundPattern",
] as const satisfies readonly (keyof DeckTheme)[];

/**
 * What a master states of its own that Design > Themes takes away: its
 * background, and Slide Master > Fonts (`fontFamily`), which is the fonts of
 * the master's theme part -- a theme picked replaces that part whole
 * (ISO/IEC 29500-1 §20.1.6.10), its font scheme with it.
 */
const MASTER_THEME_PICK_FIELDS = [
  ...MASTER_BACKGROUND_FIELDS,
  "fontFamily",
] as const satisfies readonly (keyof DeckTheme)[];












export function mastersTakingThemePick(
  masters: readonly SlideMaster[] | null | undefined,
): SlideMaster[] | null {
  if (!Array.isArray(masters)) return null;
  let changed = false;
  const next = masters.map((master) => {
    const theme = master.theme;
    if (!theme || !MASTER_THEME_PICK_FIELDS.some((field) => theme[field] !== undefined)) return master;
    changed = true;
    const kept: DeckTheme = { ...theme };
    for (const field of MASTER_THEME_PICK_FIELDS) delete kept[field];
    const out: SlideMaster = { ...master };
    if (Object.keys(kept).length > 0) out.theme = kept;
    else delete out.theme;
    return out;
  });
  return changed ? next : null;
}

/**
 * The fully resolved layout id for a slide, applying:
 *   slide.layout (explicit override) > master.layoutId > built-in default.
 *
 * A slide that pins its own `layout` keeps it (the author re-arranged this slide
 * specifically). Otherwise it inherits the master's `layoutId`. With neither, it
 * falls through to {@link DEFAULT_LAYOUT_ID} (`blank`), exactly as a master-less
 * slide did before. The slide value is re-normalized so a stale id is safe.
 */
export function resolveSlideLayoutId(
  slide:
    | { layout?: unknown; layoutId?: unknown; masterId?: unknown }
    | null
    | undefined,
  masters: readonly SlideMaster[] | null | undefined,
): LayoutId {
  // An explicit per-slide layout always wins. We only treat the slide's layout
  // as an override when it is actually present (not undefined), so a slide with
  // no layout truly inherits the master rather than forcing `blank`.
  if (slide && (slide as { layout?: unknown }).layout !== undefined) {
    return normalizeLayoutId((slide as { layout?: unknown }).layout);
  }
  const nativeLayout = resolveSlideNativeLayout(slide, masters);
  if (nativeLayout?.builtInLayoutId) return nativeLayout.builtInLayoutId;
  const master = resolveSlideMaster(slide, masters);
  if (master?.layoutId) return master.layoutId;
  return DEFAULT_LAYOUT_ID;
}

/** The resolved style + geometry context for a slide under master inheritance. */
export interface ResolvedSlideContext {
  /** The master this slide inherits from (if any). */
  master?: SlideMaster;
  /** Native child layout selected by the slide, when present. */
  nativeLayout?: SlideMasterLayout;
  /** Deck theme with the master's overrides layered on top. */
  effectiveTheme?: DeckTheme;
  /** Resolved layout id (slide override > master > default). */
  layoutId: LayoutId;
  /** Resolved slide background (slide override > master/theme > default). */
  background?: string;
  /** Resolved standalone slide-title text style (master/theme > default). */
  titleStyle: ResolvedBlockTextStyle;
}

/**
 * Resolve the full slide-level context under master inheritance in one call:
 * the effective (theme+master) theme, the resolved layout id, the slide
 * background, and the slide-title text style. Each piece honours the
 * override > master > theme > default chain. A slide with no master + a deck
 * with no theme yields `{ layoutId: 'blank', titleStyle: {} }` and an absent
 * background — identical to the prior behaviour.
 */
export function resolveSlideContext(
  slide:
    | {
        backgroundColor?: unknown;
        layout?: unknown;
        layoutId?: unknown;
        masterId?: unknown;
      }
    | null
    | undefined,
  masters: readonly SlideMaster[] | null | undefined,
  theme: DeckTheme | null | undefined,
): ResolvedSlideContext {
  const master = resolveSlideMaster(slide, masters);
  const nativeLayout = resolveSlideNativeLayout(slide, masters);
  const effectiveTheme = mergeMasterTheme(theme, master);
  const context: ResolvedSlideContext = {
    layoutId: resolveSlideLayoutId(slide, masters),
    background: resolveSlideBackground(slide, effectiveTheme),
    titleStyle: resolveTitleTextStyle(effectiveTheme),
  };
  if (master) context.master = master;
  if (nativeLayout) context.nativeLayout = nativeLayout;
  if (effectiveTheme) context.effectiveTheme = effectiveTheme;
  return context;
}

/**
 * Resolve the EFFECTIVE text style for a block under master inheritance,
 * applying the full chain:
 *
 *   block.style.color / .fontFamily (explicit override)
 *     >  master override
 *     >  deck theme
 *     >  built-in default
 *
 * Internally this layers the master's theme onto the deck theme (so a master
 * colour beats the deck theme but loses to a per-block override) and then reuses
 * the shared `resolveBlockTextStyle` precedence (which already drops a present
 * per-block override in favour of the block's own value). Non-text blocks yield
 * an empty result, exactly as the theme-only resolver does.
 */
export function resolveWithMaster(
  slide: { masterId?: unknown } | null | undefined,
  block:
    | {
        type?: unknown;
        placeholder?: unknown;
        style?: {
          color?: unknown;
          fontFamily?: unknown;
          fontSize?: unknown;
          bold?: unknown;
        } | null;
      }
    | null
    | undefined,
  masters: readonly SlideMaster[] | null | undefined,
  theme: DeckTheme | null | undefined,
): ResolvedBlockTextStyle {
  const master = resolveSlideMaster(slide, masters);
  const effectiveTheme = mergeMasterTheme(theme, master);
  return resolveBlockTextStyle(block, effectiveTheme);
}
