/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's Design tab as pure transforms: a theme from the Themes
 * gallery or a Variant, the table restyle that rides a theme change, and the
 * gate that decides when an imported deck's theme may change. The ribbon
 * (`PresentationWindow`) and the headless editor commands call the same
 * functions.
 */

import {
  nativeBlockFollowingTheme,
  nativeSlideThemeFollow,
  nativeThemeFollow,
  nativeThemeFontFollow,
} from "./presentation-native-theme-follow";
import { designPxToPoints } from "@mrnote/core/web/lib/presentation/text-units.ts";
import {
  presentationBlockKeepingThemeBindings,
  presentationBlockThemeColorBindings,
  presentationBlockUnderTheme,
  presentationThemeGridColors,
  presentationThemeSlotPicks,
  reconcilePresentationThemeColors,
} from "@mrnote/core/web/lib/presentation-theme-color-binding.ts";
import {
  presentationShapeMatchesStyle,
  presentationShapeStyleRecipes,
  presentationTextFrameMatchesStyle,
} from "@mrnote/core/web/lib/presentation-shape-style-adapter.ts";
import { blocksWithShapeStyleRecipe } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import { officeShapeStyleRecipeSlots } from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTableBlock,
  PresentationTextBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { isValidInlineImageSource } from "@mrnote/core/web/lib/inline-image-source.ts";
import {
  DEFAULT_PRESENTATION_NOTES_SIZE,
  normalizeFirstSlideNumber,
  normalizeHeaderFooter,
  normalizePresentationNotesSize,
  presentationNotesSizesEqual,
  type DeckHeaderFooter,
  type PresentationNotesSize,
} from "@mrnote/core/web/lib/presentation-header-footer.ts";
import type { SlideMaster } from "@mrnote/core/web/lib/presentation-master.ts";
import { slideDimension } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import {
  projectNativeSlideSizeDeck,
  resizeSlideSizeBlock,
  resizeSlideSizeBlocks,
  sameSlideSize,
  scaleSlideSizeBlockContent,
  slideSizeContentScale,
  stampNativeSlideSizeIntent,
  type SlideSizeEmu,
  type SlideSizeFitStrategy,
  type SlideSizeGeometryBlock,
} from "@mrnote/core/web/lib/presentation-slide-size.ts";
import { presentationTextUsesPoints } from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import { resolvePresentationTextStyle } from "./presentation-text-style-resolve";
import { SHAPE_LABEL_DEFAULT_FONT_PX } from "./presentation/text-layout-css";
import { mergeMasterTheme } from "@mrnote/core/web/lib/presentation-master.ts";
import { presentationTableStyleSchemeColors } from "./presentation/table-style-scheme";
import {
  normalizeBgGradient,
  normalizeBgPatternFill,
  normalizeBgPicture,
  normalizeSolidBackgroundTransparency,
  type SlideBgKind,
} from "@mrnote/core/web/lib/presentation-slide-bg.ts";
import {
  layoutFrameSourceForTheme,
  layoutFrameSourcesEqual,
  reflowPlaceholders,
} from "@mrnote/core/web/lib/presentation-layouts.ts";
import { presentationDeckHasNativePptxSource, type PresentationButtonSlide } from "@mrnote/core/web/lib/presentation-slide-buttons.ts";
import {
  fallbackWebTableStyleLookForTheme,
  isNativeWebTableStyleId,
  webTableStyleLookForTheme,
} from "@mrnote/core/web/lib/presentation-table-web-gallery.ts";
import {
  isOfficeThemeVariantId,
  PRESENTATION_OFFICE_THEME_IDS,
  presentationOfficeTheme,
} from "@mrnote/core/web/lib/presentation-office-themes.ts";
import {
  applyOfficeThemeSelection,
  applyPresetThemeSelection,
  applyThemeColorSchemeSelection,
  applyThemeVariantSelection,
  withoutColorSchemePick,
  PRESET_THEMES,
  resolveThemeAccentColors,
  type DeckTheme,
  type PresentationThemeVariantId,
} from "@mrnote/core/web/lib/presentation-theme.ts";

/**
 * MRAI presets whose id is also an Office theme's (`facet`) are chosen as
 * `preset-<id>`; every other preset keeps its own id.
 */
export const PRESENTATION_PRESET_CHOICE_PREFIX = "preset-";

/** The gallery / command id of an MRAI preset. */
export function presentationPresetChoiceId(presetId: string): string {
  return PRESENTATION_OFFICE_THEME_IDS.includes(presetId)
    ? `${PRESENTATION_PRESET_CHOICE_PREFIX}${presetId}`
    : presetId;
}

/**
 * Every Themes-gallery choice: the Office themes in gallery order, then the
 * MRAI presets (without "none").
 */
export const PRESENTATION_THEME_CHOICE_IDS: readonly string[] = Object.freeze([
  ...PRESENTATION_OFFICE_THEME_IDS,
  ...PRESET_THEMES.filter((preset) => preset.id !== "none").map((preset) =>
    presentationPresetChoiceId(preset.id),
  ),
]);

export function presentationThemeSelectionIsWritable(
  slides: readonly PresentationButtonSlide[],
  theme: DeckTheme | undefined,
): boolean {
  if (!presentationDeckHasNativePptxSource(slides)) return true;
  const source = theme?.identity?.source;
  // A deck of several master themes is written theme by theme (row 231's
  // `themeParts`), as PowerPoint applies a Design pick to all slides.
  const themes = source?.themeCount ?? 1;
  return Boolean(
    source &&
    (themes === 1 || source.themeParts?.length === themes) &&
    /^[0-9a-f]{64}$/u.test(source.masterDesignFingerprint ?? ""),
  );
}

/**
 * The theme a Themes gallery pick makes, or null when the pick changes
 * nothing or cannot be represented. An Office theme id applies that theme
 * (re-picking the current Office theme keeps its variant); `preset-<id>`
 * or any other id names an MRAI preset. "none" clears only a from-scratch
 * deck's theme (clearing an imported native theme would discard its
 * patchback target); re-picking Classic on a themeless deck is a no-op,
 * since the historical themeless deck already renders as Classic.
 */
export function themeAfterPreset(
  current: DeckTheme | undefined,
  choiceId: string,
): { theme: DeckTheme | undefined } | null {
  const office = choiceId.startsWith(PRESENTATION_PRESET_CHOICE_PREFIX)
    ? undefined
    : presentationOfficeTheme(choiceId);
  if (office) {
    const identity = current?.identity;
    const variantId =
      identity?.selectionKind === "office" &&
      identity.themeId === office.id &&
      isOfficeThemeVariantId(identity.variantId)
        ? identity.variantId
        : "v1";
    const theme = applyOfficeThemeSelection(current, office.id, variantId);
    return theme ? { theme } : null;
  }
  const presetId =
    choiceId.startsWith(PRESENTATION_PRESET_CHOICE_PREFIX) &&
    PRESENTATION_OFFICE_THEME_IDS.includes(
      choiceId.slice(PRESENTATION_PRESET_CHOICE_PREFIX.length),
    )
      ? choiceId.slice(PRESENTATION_PRESET_CHOICE_PREFIX.length)
      : choiceId;
  const preset = PRESET_THEMES.find((entry) => entry.id === presetId);
  if (!preset) return null;
  if (preset.id === "none") {
    return current?.identity?.source ? null : { theme: undefined };
  }
  if (!current && preset.id === "classic") return null;
  return { theme: applyPresetThemeSelection(current, preset.id) };
}

/**
 * The theme a Variants pick makes (on Classic for a themeless deck), or null
 * for a no-op. An Office theme takes only its own `v1`..`vN`; every other
 * selection only base / soft / deep / contrast.
 */
export function themeAfterVariant(
  current: DeckTheme | undefined,
  variantId: PresentationThemeVariantId,
): { theme: DeckTheme | undefined } | null {
  const officeSelection = current?.identity?.selectionKind === "office";
  if (isOfficeThemeVariantId(variantId) !== officeSelection) return null;
  if (!current && variantId === "base") return null;
  const selectedTheme =
    current ?? applyPresetThemeSelection(undefined, "classic");
  return {
    theme: withoutColorSchemePick(
      applyThemeVariantSelection(selectedTheme, variantId),
    ),
  };
}

/**
 * The theme a Variants > Colors pick makes (PPT-A073), or null for a no-op
 * or an id that is not a Colors-gallery entry.
 */
export function themeAfterColorScheme(
  current: DeckTheme | undefined,
  schemeId: string,
): { theme: DeckTheme } | null {
  const theme = applyThemeColorSchemeSelection(current, schemeId);
  return theme ? { theme } : null;
}

/** Office's own theme accents (2023), the ones an Office-themed file's table
 * styles were resolved against (`catalog.defaults`). */
const OFFICE_TABLE_STYLE_ACCENTS = ["#156082", "#e97132", "#196b24", "#0f9ed5", "#a02b93", "#4ea72e"] as const;

/** Two table looks that say the same, whatever order their keys are in. */
function sameTableStyleLook(left: unknown, right: unknown): boolean {
  const canonical = (value: unknown): unknown =>
    Array.isArray(value)
      ? value.map(canonical)
      : value && typeof value === "object"
        ? Object.fromEntries(
            Object.keys(value as Record<string, unknown>)
              .sort()
              .map((key) => [key, canonical((value as Record<string, unknown>)[key])]),
          )
        : typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value)
          ? value.toLowerCase()
          : value;
  return left !== null && left !== undefined && JSON.stringify(canonical(left)) === JSON.stringify(canonical(right));
}

/**
 * A table made here with a built-in style follows the theme: its stored
 * look is resolved again against the new accents, in the same change as the
 * theme (PPT-B274). With `previousTheme` given, the layout placeholders
 * follow the theme too: when the Office theme or variant (or the switch
 * between an Office theme and anything else) changes the layouts' frames,
 * every placeholder still on its old frame moves to the new one and takes
 * the new layout's text facts (`reflowPlaceholders`; an imported deck keeps
 * its native layouts). The same slides when nothing changes.
 */











function nativePaintStatesNoThemeColour(block: PresentationBlock): boolean {
  const record = block as unknown as {
    source?: { shapeId?: unknown; slideId?: unknown; nativeInheritedPaint?: { fields?: unknown } };
    themeColors?: Record<string, unknown>;
  };
  const source = record.source;
  if (!source || (typeof source.shapeId !== "string" && typeof source.slideId !== "string")) return false;
  const paintFields = new Set(["fill", "fillTo", "fillGradient", "stroke"]);
  const bindings = presentationBlockThemeColorBindings(block);
  if (bindings && Object.keys(bindings).some((field) => paintFields.has(field))) return false;
  const inherited = source.nativeInheritedPaint?.fields;
  return !(Array.isArray(inherited) && inherited.some((field) => paintFields.has(String(field))));
}

function blockIsUnpaintedNative(block: PresentationBlock): boolean {
  const record = block as unknown as { source?: { shapeId?: unknown }; fill?: unknown; stroke?: unknown };
  const unpainted = (value: unknown) => value === undefined || value === null || value === "none";
  return typeof record.source?.shapeId === "string" && unpainted(record.fill) && unpainted(record.stroke);
}

/** Light / dark pairs named either way (`bg1` is `lt1`, `tx1` is `dk1`, ...). */
const SLOT_ALIASES: Readonly<Record<string, string>> = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" };

/** Whether every theme colour a block states for its fill and line is the slot the recipe paints it in. */
function blockThemeSlotsAgreeWith(
  block: PresentationBlock,
  slots: { readonly fill?: string; readonly stroke?: string },
): boolean {
  const bindings = presentationBlockThemeColorBindings(block) as Record<string, { slot?: unknown } | undefined> | undefined;
  if (!bindings) return true;
  const same = (stated: unknown, painted: string | undefined) =>
    typeof stated !== "string" || painted === undefined || (SLOT_ALIASES[stated] ?? stated) === painted;
  return (
    same(bindings.fill?.slot, slots.fill) &&
    same(bindings.stroke?.slot, slots.stroke) &&
    same(bindings.outlineColor?.slot, slots.stroke)
  );
}

export function slidesRestyledForTheme<S extends PresentationButtonSlide>(
  slides: readonly S[],
  nextTheme: DeckTheme | undefined,
  previousTheme?: DeckTheme | undefined,
): S[] {
  const accents = resolveThemeAccentColors(mergeMasterTheme(nextTheme ?? undefined, undefined));
  // Dark 1 / Light 1 and the master's colour map: the words and lines of a
  // style that draws them in Text 1 (white on a dark master).
  const scheme = presentationTableStyleSchemeColors(mergeMasterTheme(nextTheme ?? undefined, undefined));
  let restyled = false;
  const before = layoutFrameSourceForTheme(previousTheme);
  const after = layoutFrameSourceForTheme(nextTheme);
  const reflowed =
    previousTheme === undefined && nextTheme === undefined
      ? slides
      : !layoutFrameSourcesEqual(before, after) && !presentationDeckHasNativePptxSource(slides)
        ? reflowPlaceholders(slides, before, after)
        : slides;
  if (reflowed !== slides) restyled = true;
  // A colour picked from the theme follows it (PPT-A037): its reference is
  // resolved again against the new theme's colours.
  const themeColors = presentationThemeGridColors(nextTheme);
  // An Office theme's colour map is known (the grid above); an imported master's
  // is told by what its references were resolved to.
  const previousColorsForMap =
    previousTheme && previousTheme.identity?.selectionKind !== "office" && nextTheme?.identity?.selectionKind !== "office"
      ? presentationThemeGridColors(previousTheme)
      : undefined;
  // A shape or text box still painted by one of the theme's Shape Styles
  // takes the same style in the new theme, as PowerPoint re-resolves a style
  // reference: an Accent 1 gradient kept the old theme's blues while its
  
  const stylesBefore = presentationShapeStyleRecipes(mergeMasterTheme(previousTheme, undefined));
  const stylesAfter = new Map(
    presentationShapeStyleRecipes(mergeMasterTheme(nextTheme, undefined)).map((recipe) => [recipe.id, recipe]),
  );
  const underThemeStyle = (block: PresentationBlock): PresentationBlock => {
    if (block.type !== "shape" && block.type !== "text") return block;
    // An imported object that paints nothing is painted by no style: a title
    // whose words inherit their colour matched "no fill, no line", and the
    // recipe stamped the new theme's text colour on its runs, an explicit
    
    
    if (blockIsUnpaintedNative(block) || nativePaintStatesNoThemeColour(block)) return block;
    const before = stylesBefore.find((recipe) =>
      block.type === "shape"
        ? presentationShapeMatchesStyle(block as PresentationShapeBlock, recipe)
        : presentationTextFrameMatchesStyle(block as PresentationTextBlock, recipe),
    );
    const after = before ? stylesAfter.get(before.id) : undefined;
    if (!before || !after || JSON.stringify(after) === JSON.stringify(before)) return block;
    // The style the object wears is the one its own theme colours name: a
    // fill bound to Background 1 is no Accent 3 style because the old scheme's
    
    // restyled such text boxes in the new Accent 3 grey).
    if (!blockThemeSlotsAgreeWith(block, officeShapeStyleRecipeSlots(before))) return block;
    const restyled = blocksWithShapeStyleRecipe([block], [block], after)?.[0];
    if (!restyled) return block;
    // The style's theme colours stay bound, now to the new theme's, and so
    // does a reference the object stated that the restyled colour still is
    
    return presentationBlockKeepingThemeBindings(
      block,
      reconcilePresentationThemeColors(undefined, restyled, presentationThemeSlotPicks(themeColors)),
      themeColors,
    );
  };
  // A deck opened from a file: what its words inherit, drawn at once (the server's
  // reading of the written file then corrects what this got wrong).
  const nativeFollow = presentationDeckHasNativePptxSource(slides)
    ? (nativeThemeFollow(previousTheme, nextTheme) ?? nativeThemeFontFollow(previousTheme, nextTheme))
    : null;
  const nextSlides = reflowed.map((slide) => {
    let changed = false;
    const slideFollow = nativeFollow ? nativeSlideThemeFollow(nativeFollow, slide) : null;
    const blocks = (slide.blocks ?? []).map((input) => {
      const block = slideFollow ? nativeBlockFollowingTheme(input, slideFollow) : input;
      if (block !== input) changed = true;
      // An SVG graphic's Graphics Fill / Outline is a theme pick too
      // (PPT-A161): it stayed on the old theme's colour because only shapes
      // and text boxes were resolved again (acceptance step 35).
      // A member of a diagram the page preserves as drawn is the file's: the
      // written theme repaints it, and a member the page restates is a declined
      
      // "smartArtFormat" on every text of a SmartArt, for a theme alone).
      if ((block as { nativeGroupDisposition?: unknown }).nativeGroupDisposition === "preserve-only") return block;
      if (block.type === "shape" || block.type === "text" || block.type === "image") {
        const recoloured = presentationBlockUnderTheme(underThemeStyle(block), themeColors, previousColorsForMap);
        if (recoloured !== block) changed = true;
        return recoloured;
      }
      if (block.type !== "table") return block;
      const table = block as PresentationTableBlock;
      if (!table.tableStyleLook) return block;
      // An imported table's look was resolved from its file's style; it
      // follows the theme only when it IS the built-in style the GUID names,
      // resolved against the old theme (its file then states theme colours
      // the new theme repaints, as PowerPoint and the downloaded file show:
      
      // A deck whose theme changed before this rule still holds the look its
      // file resolved under Office's own theme; that one is the built-in
      // style too. So is the recipe's look, applied while the measured looks
      // were away (a guest page).
      // (The import resolved it in the deck's own Dark 1 / Light 1 and colour
      // map too: those first, then the accents alone, as before.)
      const previousMerged = previousTheme === undefined ? null : mergeMasterTheme(previousTheme, undefined);
      if (
        (table as { source?: unknown }).source &&
        ![
          previousMerged === null
            ? null
            : ([resolveThemeAccentColors(previousMerged), presentationTableStyleSchemeColors(previousMerged)] as const),
          previousMerged === null ? null : ([resolveThemeAccentColors(previousMerged), undefined] as const),
          [OFFICE_TABLE_STYLE_ACCENTS, undefined] as const,
        ].some(
          (before) =>
            before !== null &&
            (sameTableStyleLook(webTableStyleLookForTheme(table.tableStyleId, before[0], before[1]), table.tableStyleLook) ||
              sameTableStyleLook(fallbackWebTableStyleLookForTheme(table.tableStyleId, before[0], before[1]), table.tableStyleLook)),
        )
      ) {
        return block;
      }
      const look = webTableStyleLookForTheme(table.tableStyleId, accents, scheme);
      if (!look || sameTableStyleLook(look, table.tableStyleLook)) return block;
      changed = true;
      return { ...table, tableStyleLook: look } as PresentationBlock;
    });
    if (!changed) return slide;
    restyled = true;
    return { ...slide, blocks };
  });
  return restyled ? nextSlides : (slides as S[]);
}

/**
 * The measured looks landed on a deck whose tables were styled while they
 * were away (a guest page; the seconds before they arrive): a table that
 * carries the recipe's look for its built-in style
 * (`office-table-style-fallback.ts`), in the deck's accents or the gallery's
 * own, takes the style's own look in the deck's accents, as a theme change
 * resolves it (PPT-B274). Every other look stays. The same slides when
 * nothing changes.
 */
export function slidesWithServedTableStyleLooks<S extends PresentationButtonSlide>(
  slides: readonly S[],
  theme: DeckTheme | undefined,
): S[] {
  const accents = resolveThemeAccentColors(mergeMasterTheme(theme ?? undefined, undefined));
  const scheme = presentationTableStyleSchemeColors(mergeMasterTheme(theme ?? undefined, undefined));
  let restyled = false;
  const nextSlides = slides.map((slide) => {
    let changed = false;
    const blocks = (slide.blocks ?? []).map((block) => {
      if (block.type !== "table") return block;
      const table = block as PresentationTableBlock;
      if (!table.tableStyleLook || !isNativeWebTableStyleId(table.tableStyleId)) return block;
      const carriesFallback = ([[accents, scheme], [accents, undefined], [OFFICE_TABLE_STYLE_ACCENTS, undefined]] as const).some(
        ([before, beforeScheme]) =>
          sameTableStyleLook(fallbackWebTableStyleLookForTheme(table.tableStyleId, before, beforeScheme), table.tableStyleLook),
      );
      if (!carriesFallback) return block;
      const look = webTableStyleLookForTheme(table.tableStyleId, accents, scheme);
      if (!look || sameTableStyleLook(look, table.tableStyleLook)) return block;
      changed = true;
      return { ...table, tableStyleLook: look } as PresentationBlock;
    });
    if (!changed) return slide;
    restyled = true;
    return { ...slide, blocks };
  });
  return restyled ? nextSlides : (slides as S[]);
}

/** The background fill a slide states: gradient, picture, pattern, solid or none. */
export function slideBackgroundKind(slide: Record<string, unknown>): SlideBgKind {
  return normalizeBgGradient(slide.backgroundGradient)
    ? "gradient"
    : normalizeBgPicture(slide.backgroundImage, isValidInlineImageSource) ||
        (slide.backgroundImage && typeof slide.backgroundImage === "object")
      ? "picture"
      : normalizeBgPatternFill(slide.backgroundPattern)
        ? "pattern"
        : typeof slide.backgroundColor === "string"
          ? "solid"
          : "none";
}

/**
 * Format Background's fill kind: switching clears the other kinds' fields so
 * one fill is persisted and an unfilled slide stays byte-clean; `solid`
 * keeps its colour (white when unset) and transparency, `gradient` and
 * `pattern` seed a visible default, `picture` an empty url awaiting one.
 */
export function slideWithBackgroundKind<S extends object>(slide: S, kind: SlideBgKind): S {
  // The kind the slide already has keeps what it holds, and is the slide
  // itself: choosing Picture again emptied the picture, and Gradient or
  // Pattern again put the seed back over the author's (PPT-B323).
  if (kind === slideBackgroundKind(slide as Record<string, unknown>)) return slide;
  const base = { ...slide } as Record<string, unknown>;
  const solidTransparency = normalizeSolidBackgroundTransparency(
    base.backgroundTransparency,
  );
  delete base.backgroundGradient;
  delete base.backgroundImage;
  delete base.backgroundPattern;
  delete base.backgroundTransparency;
  if (kind === "none") {
    delete base.backgroundColor;
  } else if (kind === "solid") {
    if (typeof base.backgroundColor !== "string") {
      base.backgroundColor = "#ffffff";
    }
    if (solidTransparency > 0) {
      base.backgroundTransparency = solidTransparency;
    }
  } else if (kind === "gradient") {
    const seedFrom =
      typeof base.backgroundColor === "string"
        ? (base.backgroundColor as string)
        : "#2563eb";
    base.backgroundGradient = { from: seedFrom, to: "#0ea5e9", angle: 90 };
  } else if (kind === "picture") {
    base.backgroundImage = { url: "", fit: "cover" };
  } else if (kind === "pattern") {
    base.backgroundPattern = {
      pattern: "dots",
      color: "#cbd5e1",
      background: "#ffffff",
    };
  }
  return base as S;
}

/** Reset Background: every fill field of the slide goes; it shows the theme's. */
export function slideWithoutBackground<S extends object>(slide: S): S {
  const next = { ...slide } as Record<string, unknown>;
  delete next.backgroundColor;
  delete next.backgroundGradient;
  delete next.backgroundPattern;
  delete next.backgroundImage;
  delete next.backgroundTransparency;
  return next as S;
}

/** Hide Background Graphics: stated only when on. */
export function slideWithHiddenBackgroundGraphics<S extends object>(slide: S, hidden: boolean): S {
  const nextSlide = { ...slide } as Record<string, unknown>;
  if (hidden) {
    nextSlide.hideBackgroundGraphics = true;
  } else {
    delete nextSlide.hideBackgroundGraphics;
  }
  return nextSlide as S;
}

/**
 * Slide Size: every slide (and master and layout) to the new size with
 * PowerPoint's content-fit strategy, plus the first slide number and notes
 * page the dialog sets. An imported deck is projected through its native
 * size authority; a scratch deck scales its content with the slide
 * (PPT-B089). "unchanged" when nothing changes, null when the deck's size
 * cannot change here.
 */
export function deckAfterSlideSize<
  S extends PresentationButtonSlide,
  M extends { blocks?: unknown[]; layouts?: Array<{ blocks?: unknown[] }> },
>(input: {
  slides: S[];
  masters: M[];
  theme: DeckTheme | undefined;
  headerFooter: DeckHeaderFooter | undefined;
  activeIndex: number;
  fallbackSize: SlideSizeEmu;
  size: SlideSizeEmu;
  strategy: SlideSizeFitStrategy;
  firstSlideNumber: number;
  notesSize: PresentationNotesSize;
}): { slides: S[]; masters: M[]; headerFooter: DeckHeaderFooter | undefined } | "unchanged" | null {
  const { slides: liveSlides, size, strategy } = input;
  const liveSlide = liveSlides[input.activeIndex];
  const liveWidth = slideDimension(liveSlide, "widthEmu");
  const liveHeight = slideDimension(liveSlide, "heightEmu");
  const current =
    liveWidth && liveHeight
      ? { widthEmu: liveWidth, heightEmu: liveHeight }
      : input.fallbackSize;
  const currentFirstSlideNumber =
    input.headerFooter?.firstSlideNumber ?? 1;
  const firstSlideNumber =
    normalizeFirstSlideNumber(input.firstSlideNumber) ?? 1;
  const firstSlideNumberChanged =
    currentFirstSlideNumber !== firstSlideNumber;
  const currentNotesSize =
    input.headerFooter?.notesSize ?? DEFAULT_PRESENTATION_NOTES_SIZE;
  const notesSizeChanged = !presentationNotesSizesEqual(
    currentNotesSize,
    input.notesSize,
  );
  const nextHeaderFooter = normalizeHeaderFooter({
    ...(input.headerFooter ?? {}),
    firstSlideNumber:
      normalizeFirstSlideNumber(input.firstSlideNumber),
    notesSize: normalizePresentationNotesSize(input.notesSize),
  });
  if (
    sameSlideSize(current, size) &&
    !firstSlideNumberChanged &&
    !notesSizeChanged
  ) {
    return "unchanged";
  }
  if (sameSlideSize(current, size)) {
    return { slides: liveSlides, masters: input.masters, headerFooter: nextHeaderFooter };
  }
  const importedDeck = liveSlides.some((slide) =>
    Boolean(
      slide.source &&
        typeof slide.source === "object" &&
        typeof (slide.source as Record<string, unknown>).slideId === "string",
    ),
  );
  // The content scales with the slide, not just its frames (PPT-B089).
  const contentScale = slideSizeContentScale(current, size, strategy);
  const theme = input.theme;
  const scaledContent = <B extends SlideSizeGeometryBlock>(
    blocks: readonly B[] | undefined,
    drawnFontSize: (block: B) => number | undefined,
  ) =>
    blocks?.map((block) =>
      scaleSlideSizeBlockContent(
        resizeSlideSizeBlock(block, current, size, strategy),
        contentScale,
        drawnFontSize(block),
      ),
    );
  // What a block's text draws at when it states no size: the master's or
  // theme's size, else the canvas default, in the block's own unit.
  const drawnFontSize = (
    slide: S,
    block: PresentationBlock,
  ): number | undefined => {
    // A shape's words that state none draw at PowerPoint's 18 pt
    // (SHAPE_LABEL_DEFAULT_FONT_PX); 12 pt / 16 px scaled them from too small.
    if (block.type === "shape") return presentationTextUsesPoints(block) ? 18 : SHAPE_LABEL_DEFAULT_FONT_PX;
    if (block.type !== "text") return undefined;
    const inherited = resolvePresentationTextStyle(slide, block, input.masters as unknown as SlideMaster[], theme).fontSize;
    if (typeof inherited === "number") return inherited;
    return presentationTextUsesPoints(block) ? designPxToPoints(17) : 17;
  };
  const nextSlides = importedDeck
    ? projectNativeSlideSizeDeck(liveSlides, size, strategy)
    : stampNativeSlideSizeIntent(
        liveSlides.map((slide) => ({
          ...slide,
          widthEmu: size.widthEmu,
          heightEmu: size.heightEmu,
          blocks:
            scaledContent(slide.blocks as unknown as SlideSizeGeometryBlock[], (block) => drawnFontSize(slide, block as unknown as PresentationBlock)) ??
            slide.blocks,
        })),
        strategy,
      );
  if (!nextSlides) return null;
  const nextMasters = input.masters.map((master) => {
    const blocks = importedDeck
      ? resizeSlideSizeBlocks(master.blocks as SlideSizeGeometryBlock[] | undefined, current, size, strategy)
      : scaledContent(master.blocks as SlideSizeGeometryBlock[] | undefined, () => undefined);
    const layouts = master.layouts?.map((layout) => {
      const layoutBlocks = importedDeck
        ? resizeSlideSizeBlocks(layout.blocks as SlideSizeGeometryBlock[] | undefined, current, size, strategy)
        : scaledContent(layout.blocks as SlideSizeGeometryBlock[] | undefined, () => undefined);
      return {
        ...layout,
        ...(layoutBlocks ? { blocks: layoutBlocks } : {}),
      };
    });
    return {
      ...master,
      ...(blocks ? { blocks } : {}),
      ...(layouts ? { layouts } : {}),
    };
  });
  return { slides: nextSlides as S[], masters: nextMasters, headerFooter: nextHeaderFooter };
}
