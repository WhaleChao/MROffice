/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's Slides buttons as pure slide transforms: New Slide (with
 * a layout), Layout, Reset, and the gates that decide when an imported deck
 * may take them. The ribbon (`PresentationWindow`) and the headless editor
 * commands call the same functions on the same slide records.
 */

import { pruneSequence, type TimelineStep } from "./presentation-anim-timeline";
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import { presentationCommentAuthoritySchema } from "./native-page-schemas";
import { MAX_PRESENTATION_SLIDES } from "./presentation-conveniences";
import { officeMasterLayoutRef, parseOfficeMasterLayoutRef } from "./presentation-office-master-model";
import {
  applyNativeLayoutToSlide,
  BUILTIN_LAYOUT_FRAMES,
  createNativeSlideInsertionReceipt,
  DEFAULT_LAYOUT_ID,
  LAYOUT_DEFINITIONS,
  layoutBlocks,
  layoutTextStyleAfterTemplate,
  nativeLayoutCanApply,
  nativeLayoutCatalogEntryForValue,
  nativeLayoutModelLayoutId,
  nativeLayoutRef,
  parseNativeLayoutRef,
  nativeLayoutIdFromType,
  nativeLayoutInsertionBlocks,
  type NativeLayoutCatalogMasters,
  nativeResetCanApply,
  nativeSlideCanInsertLayout,
  filledPlaceholderAtTemplate,
  isFilledPlaceholderBlock,
  normalizeSlideLayoutRef,
  parseNativeLayoutCascadeReceipt,
  parseNativeSlideInsertionReceipt,
  resetNativeSlideToLayout,
  resetSlideToLayout,
  type LayoutFrameSource,
  type LayoutId,
  type OfficeMasterLayoutRef,
} from "./presentation-layouts";
import {
  normalizeHeaderFooter,
  type DeckDateField,
  type DeckHeaderFooter,
  type PresentationDateFieldFormat,
  type PresentationDateFieldLocale,
  type SlideHeaderFooterOverride,
} from "./presentation-header-footer";
import {
  slideWithNativeLayoutBackground,
  slideWithNativeLayoutGraphics,
  type SlideMaster,
} from "./presentation-master";
import { canRequestNativeSlideCopy } from "./presentation-native-object-copy";
import { deckSlideSizeFields, slideBounds } from "./presentation-object-buttons";

/** The fields of a slide these buttons read and write. */
export interface PresentationButtonSlide {
  id: string;
  title: string;
  blocks: PresentationBlock[];
  source?: Record<string, unknown>;
  layout?: LayoutId | OfficeMasterLayoutRef;
  layoutId?: string;
  masterId?: string;
  commentAuthority?: unknown;
  [key: string]: unknown;
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/** True for a scratch text block explicitly authored as a free text box. */
export function isTextBoxBlock(block: PresentationBlock): boolean {
  return (
    block.type === "text" &&
    (block as Record<string, unknown>).kind === "textbox"
  );
}

function placeholderHoldsText(block: PresentationBlock): boolean {
  const record = block as Record<string, unknown>;
  return typeof record.text === "string" && record.text.length > 0;
}

/**
 * Re-pour one slide into a built-in layout preview without dropping authored
 * content. Native layouts keep their stable `layoutId`; this helper only updates
 * the editable preview geometry mirrored by the legacy `layout` field.
 */
export function reflowSlideToLayout<S extends PresentationButtonSlide>(
  slide: S,
  layoutValue: unknown,
  bounds: { width?: number | null; height?: number | null },
  source: LayoutFrameSource = BUILTIN_LAYOUT_FRAMES,
): S {
  const layout = normalizeSlideLayoutRef(layoutValue);
  const templates = layoutBlocks(layout, bounds, source);
  // The places the slide's placeholders came from: a text value a block still
  // holds as its old layout stated it follows the new layout; the author's
  // own stays (the same rule as a theme change, `reflowPlaceholders`).
  const previous = layoutBlocks(slide.layout, bounds, source);
  // Only the slide's placeholders take the new layout's places. A text box the
  // user inserted (a text box with no placeholder key or role; a layout's own
  // places carry both) is an object of its own: PowerPoint's Layout keeps it
  // where it is, its size and its role (PPT-B094).
  const pours = (block: PresentationBlock) => {
    // A table or chart a content placeholder was filled with is that place's
    
    if (isFilledPlaceholderBlock(block)) return true;
    if (block.type !== "text") return false;
    const record = block as Record<string, unknown>;
    return !(
      isTextBoxBlock(block) &&
      record.placeholder === undefined &&
      record.role === undefined
    );
  };
  const textBlocks = slide.blocks.filter(pours);
  const placed: PresentationBlock[] = templates.map((template, index) => {
    const source = textBlocks[index];
    if (source && isFilledPlaceholderBlock(source)) {
      return filledPlaceholderAtTemplate(source as unknown as Record<string, unknown>, template) as unknown as PresentationBlock;
    }
    const { style: templateStyle, ...frame } = template;
    const style = layoutTextStyleAfterTemplate(
      (source as Record<string, unknown> | undefined)?.style as
        | Record<string, unknown>
        | undefined,
      previous[index]?.style as Record<string, unknown> | undefined,
      templateStyle as Record<string, unknown> | undefined,
    );
    const next = {
      ...(source ?? {}),
      ...frame,
      id: source?.id ?? template.id,
      text: typeof source?.text === "string" ? source.text : "",
    } as unknown as Record<string, unknown>;
    if (style) next.style = style;
    else delete next.style;
    return next as unknown as PresentationBlock;
  });
  // Each placeholder takes its new place where it stands in the stacking
  // order: a shape the author sent behind the title stays behind it
  // (PPT-B412; this used to put every placeholder in front of every other
  // object). A place the new layout adds goes just in front of the slide's
  // last placeholder, or at the back of a slide that had none.
  const pourIndex = new Map(textBlocks.map((block, index) => [block, index]));
  const blocks: PresentationBlock[] = [];
  let lastPlaceholder = -1;
  for (const block of slide.blocks) {
    const index = pourIndex.get(block);
    if (index === undefined) {
      blocks.push(block);
      continue;
    }
    // A placeholder the new layout has no place for stays where it is when
    // it holds words and goes when it is empty, as PowerPoint's Layout does
    // (a Title Only slide keeps no empty "Click to add text" boxes).
    if (index >= placed.length && !placeholderHoldsText(block) && !isFilledPlaceholderBlock(block)) continue;
    blocks.push(index < placed.length ? placed[index] : block);
    lastPlaceholder = blocks.length - 1;
  }
  blocks.splice(lastPlaceholder + 1, 0, ...placed.slice(textBlocks.length));
  const next: S = { ...slide, blocks };
  // Blank is stated too: a slide put on it may then be emptied, as
  // PowerPoint's Blank slide can (`blocksLeftOnSlide`).
  next.layout = layout;
  return next;
}

export function presentationDeckHasNativePptxSource(
  slides: readonly PresentationButtonSlide[],
): boolean {
  return slides.some((slide) => {
    const slideSource = asRecord(slide.source);
    if (
      typeof slideSource?.slideId === "string" &&
      slideSource.slideId.trim() !== ""
    ) {
      return true;
    }
    if (parseNativeSlideInsertionReceipt(slideSource?.nativeSlideInsertion)) {
      return true;
    }
    return slide.blocks.some((block) => {
      const source = asRecord((block as Record<string, unknown>).source);
      return (
        typeof source?.slideId === "string" && source.slideId.trim() !== ""
      );
    });
  });
}

export function presentationDeckNativeNewSlideGraphIsWritable(
  slides: readonly PresentationButtonSlide[],
): boolean {
  // New Slide adds a comment-free slide whatever the deck's comment graph:
  // one the page edits or one it only keeps (classic comments, or a graph
  // the reader could not authenticate) -- the slide takes the deck's receipt
  // and no comment part changes. Only receipts that disagree (no one deck
  // graph to give the slide) hold it back. The CDC TB deck, an author list
  // and no comment, had New Slide greyed out on every slide.
  const authorities = slides.map((slide) => slide.commentAuthority);
  if (authorities.every((authority) => authority === undefined)) return true;
  const parsed = authorities.map((authority) =>
    presentationCommentAuthoritySchema.safeParse(authority),
  );
  if (parsed.some((result) => !result.success)) return false;
  const canonical = JSON.stringify(parsed[0]?.data);
  return parsed.every(
    (result) => result.success && JSON.stringify(result.data) === canonical,
  );
}

export function presentationNewSlideFromLayoutIsWritable(
  slides: readonly PresentationButtonSlide[],
  activeSlide: PresentationButtonSlide | undefined,
  layoutValue: unknown,
): boolean {
  if (!presentationDeckHasNativePptxSource(slides)) return true;
  return Boolean(
    activeSlide &&
    presentationDeckNativeNewSlideGraphIsWritable(slides) &&
    nativeSlideCanInsertLayout(activeSlide, layoutValue),
  );
}

export function presentationDefaultNewSlideLayout(
  slides: readonly PresentationButtonSlide[],
  activeSlide: PresentationButtonSlide | undefined,
): LayoutId | null {
  // New Slide after a slide on a master Insert Slide Master added takes that
  // master's Title and Content, as PowerPoint adds it under the current
  // slide's master (PPT-B030).
  const masterRef = parseOfficeMasterLayoutRef(activeSlide?.layout);
  if (masterRef) {
    const sameMaster = officeMasterLayoutRef(masterRef.masterKey, "title-content") as LayoutId;
    if (presentationNewSlideFromLayoutIsWritable(slides, activeSlide, sameMaster)) return sameMaster;
  }
  if (
    presentationNewSlideFromLayoutIsWritable(
      slides,
      activeSlide,
      "title-content",
    )
  ) {
    return "title-content";
  }
  const receipt = presentationSlideNativeLayoutReceipt(activeSlide);
  const currentLayoutPart =
    typeof activeSlide?.layoutId === "string"
      ? activeSlide.layoutId
      : receipt?.layoutPart;
  const currentLayout = receipt?.layouts.find(
    (entry) => entry.layoutPart === currentLayoutPart,
  );
  const currentLayoutId = nativeLayoutIdFromType(currentLayout?.layoutType);
  if (
    currentLayoutId &&
    presentationNewSlideFromLayoutIsWritable(
      slides,
      activeSlide,
      currentLayoutId,
    )
  ) {
    return currentLayoutId;
  }
  const builtIn = LAYOUT_DEFINITIONS.find((layout) =>
    presentationNewSlideFromLayoutIsWritable(slides, activeSlide, layout.id),
  )?.id;
  if (builtIn) return builtIn;
  // Else the file's own layout the slide is on, then any it can take: a
  // deck's New Slide is never closed while its gallery has a layout open.
  const nativeChoices = [
    ...(currentLayoutPart ? [currentLayoutPart] : []),
    ...(receipt?.layouts.map((entry) => entry.layoutPart) ?? []),
  ].map((part) => nativeLayoutRef(part) as LayoutId);
  return (
    nativeChoices.find((value) =>
      presentationNewSlideFromLayoutIsWritable(slides, activeSlide, value),
    ) ?? null
  );
}

export function presentationDuplicateSlideIsWritable(
  slides: readonly PresentationButtonSlide[],
  activeSlide: PresentationButtonSlide | undefined,
): boolean {
  if (!presentationDeckHasNativePptxSource(slides)) return Boolean(activeSlide);
  return Boolean(
    activeSlide && canRequestNativeSlideCopy(activeSlide),
  );
}

export function presentationSlideHasNativePptxSource(slide: PresentationButtonSlide | undefined): boolean {
  const source = asRecord(slide?.source);
  return Boolean(
    typeof source?.slideId === "string" && source.slideId.trim() !== "",
  );
}

export function presentationSlideNativeLayoutReceipt(slide: PresentationButtonSlide | undefined) {
  const source = asRecord(slide?.source);
  return parseNativeLayoutCascadeReceipt(source?.layoutCascade);
}

export function presentationSlideLayoutCommandsAreWritable(
  slide: PresentationButtonSlide | undefined,
): boolean {
  if (!presentationSlideHasNativePptxSource(slide)) return true;
  const receipt = presentationSlideNativeLayoutReceipt(slide);
  if (!receipt) return false;
  const currentLayoutPart =
    typeof slide?.layoutId === "string" ? slide.layoutId : receipt.layoutPart;
  // A slide on a custom layout (no built-in type, as Google Slides and many
  // templates write them) changes layout as any other: the type names the
  // gallery's built-in it resembles, nothing the change needs.
  const currentLayout = receipt.layouts.find(
    (entry) => entry.layoutPart === currentLayoutPart,
  );
  if (!currentLayout) return false;
  return receipt.layouts.some(
    (target) =>
      target.layoutPart !== currentLayoutPart &&
      nativeLayoutCanApply(receipt, target),
  );
}

export function presentationSlideResetIsWritable(
  slide: PresentationButtonSlide | undefined,
): boolean {
  if (!slide) return false;
  if (!presentationSlideHasNativePptxSource(slide)) return true;
  const receipt = presentationSlideNativeLayoutReceipt(slide);
  return Boolean(
    receipt &&
      nativeResetCanApply(
        slide as unknown as Record<string, unknown>,
        receipt,
      ),
  );
}


/**
 * New Slide: a slide with `layoutValue` (else the deck's default new-slide
 * layout) inserted after the slide at `anchorIndex` (an empty deck takes it
 * first). On an imported deck the slide is made from the native layout with
 * its insertion receipt, or refused when the deck cannot take it. Returns the
 * new slide list and the new slide's index, or null.
 */
export function slidesWithNewSlideFromLayout<S extends PresentationButtonSlide>(
  liveSlides: readonly S[],
  insertionAnchorIndex: number,
  layoutValue: string | undefined,
  makeId: () => string,
  source: LayoutFrameSource = BUILTIN_LAYOUT_FRAMES,
   
  masters?: NativeLayoutCatalogMasters,
): { slides: S[]; index: number } | null {
  if (liveSlides.length >= MAX_PRESENTATION_SLIDES) return null;
  const liveSource = liveSlides[insertionAnchorIndex];
  // A non-empty deck with no exact live insertion anchor is stale. Never
  // append or fall back to another slide because that can cross sections.
  if (liveSlides.length > 0 && !liveSource) return null;
  const requestedLayout =
    layoutValue ??
    presentationDefaultNewSlideLayout(liveSlides, liveSource);
  if (!requestedLayout) return null;
  const nativeDeck = presentationDeckHasNativePptxSource(liveSlides);
  // An imported deck's own layout (`native-layout:<part>`) is found by its
  // part; the slide model still reads the built-in layout it amounts to.
  const nativeLayoutValue = parseNativeLayoutRef(requestedLayout) ? requestedLayout : null;
  const nativeTarget =
    nativeLayoutValue && liveSource
      ? (() => {
          const receipt = presentationSlideNativeLayoutReceipt(liveSource);
          return receipt ? nativeLayoutCatalogEntryForValue(receipt, nativeLayoutValue) : null;
        })()
      : null;
  if (nativeLayoutValue && !nativeTarget) return null;
  const layout = nativeTarget
    ? nativeLayoutModelLayoutId(nativeTarget)
    : normalizeSlideLayoutRef(requestedLayout);
  const insertionLayout = nativeLayoutValue ?? layout;
  if (
    nativeDeck &&
    (!liveSource ||
      !presentationNewSlideFromLayoutIsWritable(
        liveSlides,
        liveSource,
        insertionLayout,
      ))
  ) {
    return null;
  }
  const bounds = slideBounds(liveSource);
  const modelSlideId = makeId();
  const nativeInsertionReceipt =
    nativeDeck && liveSource
      ? createNativeSlideInsertionReceipt(
          liveSource,
          modelSlideId,
          "new",
          insertionLayout,
        )
      : null;
  const blocks =
    nativeDeck && liveSource
      ? nativeLayoutInsertionBlocks(liveSource, insertionLayout, makeId, masters)
      : layoutBlocks(layout, bounds, source);
  if (nativeDeck && (!nativeInsertionReceipt || blocks === null)) return null;
  const sourceSize = deckSlideSizeFields(liveSource);
  const next = {
    id: modelSlideId,
    title: "",
    // The Blank layout has no placeholder: its slide starts empty and says it
    // is on Blank, as in PowerPoint (it used to get a body prompt nothing in
    // the layout asked for; `blocksLeftOnSlide`).
    blocks: (blocks ?? []) as unknown as PresentationBlock[],
    ...(layout !== DEFAULT_LAYOUT_ID || !blocks?.length ? { layout } : {}),
    ...(nativeInsertionReceipt
      ? {
          layoutId: nativeInsertionReceipt.layoutPart,
          masterId: nativeInsertionReceipt.masterPart,
          source: { nativeSlideInsertion: nativeInsertionReceipt },
        }
      : {}),
    ...sourceSize,
  } as unknown as S;
  const nextIndex = liveSource ? insertionAnchorIndex + 1 : 0;
  return {
    slides: [
      ...liveSlides.slice(0, nextIndex),
      next,
      ...liveSlides.slice(nextIndex),
    ],
    index: nextIndex,
  };
}

/**
 * Layout: the slide re-poured into `layoutValue`. An imported slide takes the
 * native layout through its layout-cascade receipt, or is refused (null); a
 * slide already on it comes back as itself.
 */
export function slideWithLayout<S extends PresentationButtonSlide>(
  liveSlide: S,
  layoutValue: string,
  source: LayoutFrameSource = BUILTIN_LAYOUT_FRAMES,
  /** The deck's masters, with the layout cards drawn so far: an imported
   * slide takes its new layout's graphics from them (`slideWithNativeLayoutGraphics`). */
  masters?: readonly SlideMaster[],
): S | null {
  const liveSource = asRecord(liveSlide.source);
  const rawLayoutCascade = liveSource?.layoutCascade;
  const layoutCascade = parseNativeLayoutCascadeReceipt(rawLayoutCascade);
  const layout = normalizeSlideLayoutRef(layoutValue);
  if (
    (rawLayoutCascade !== undefined && !layoutCascade) ||
    (presentationSlideHasNativePptxSource(liveSlide) && !layoutCascade)
  ) {
    return null;
  }
  if (layoutCascade) {
    const nativeLayout = nativeLayoutCatalogEntryForValue(layoutCascade, layoutValue);
    if (!nativeLayout || !nativeLayoutCanApply(layoutCascade, nativeLayout)) {
      return null;
    }
    const applied = applyNativeLayoutToSlide(
      liveSlide as unknown as Record<string, unknown>,
      layoutCascade,
      nativeLayout,
    ) as unknown as S;
    // PowerPoint redraws the slide on its new layout at once: the layout's
    // and master's graphics, and the layout's background on a slide that
    // has none of its own.
    return applied !== liveSlide && masters
      ? slideWithNativeLayoutBackground(liveSlide, slideWithNativeLayoutGraphics(applied, masters), masters)
      : applied;
  }
  const next = reflowSlideToLayout(
    liveSlide,
    layout,
    slideBounds(liveSlide),
    source,
  );
  const nextSlide: S = { ...liveSlide, blocks: next.blocks };
  if (next.layout) nextSlide.layout = next.layout;
  else delete nextSlide.layout;
  // An empty placeholder the new layout has no place for goes, and its
  // animation with it: a step left pointing at it cost the show a click
  
  const sequence = (liveSlide as { animationSequence?: TimelineStep[] }).animationSequence;
  if (Array.isArray(sequence) && sequence.length > 0) {
    const pruned = pruneSequence(sequence, new Set(next.blocks.map((block) => block.id)));
    if (pruned.length !== sequence.length || pruned.some((step, index) => step !== sequence[index])) {
      (nextSlide as { animationSequence?: TimelineStep[] }).animationSequence = pruned;
    }
  }
  return nextSlide;
}

/**
 * Reset: the slide's placeholders back to their layout's places and format
 * (an imported slide through its native receipt); the slide itself when the
 * reset does not apply.
 */
export function slideAfterReset<S extends PresentationButtonSlide>(
  liveSlide: S,
  source: LayoutFrameSource = BUILTIN_LAYOUT_FRAMES,
): S {
  const nativeReceipt = presentationSlideNativeLayoutReceipt(liveSlide);
  return presentationSlideHasNativePptxSource(liveSlide)
    ? nativeReceipt &&
      nativeResetCanApply(
        liveSlide as unknown as Record<string, unknown>,
        nativeReceipt,
      )
      ? (resetNativeSlideToLayout(
          liveSlide as unknown as Record<string, unknown>,
          nativeReceipt,
        ) as unknown as S)
      : liveSlide
    : (resetSlideToLayout(
        liveSlide as unknown as Record<string, unknown>,
        slideBounds(liveSlide),
        undefined,
        source,
      ) as unknown as S);
}

/** What the Header & Footer dialog holds. */
export interface PresentationHeaderFooterDialogDraft {
  slideNumber: boolean;
  footerEnabled: boolean;
  footerText: string;
  dateEnabled: boolean;
  dateMode: "auto" | "fixed";
  fixedDateText: string;
  dateFormat: PresentationDateFieldFormat;
  dateLocale: PresentationDateFieldLocale;
  excludeTitleSlide: boolean;
}

/** The dialog's fields as a deck header/footer value. */
export function presentationHeaderFooterDialogValue(
  draft: PresentationHeaderFooterDialogDraft,
): DeckHeaderFooter | undefined {
  const dateField: DeckDateField | undefined =
    !draft.dateEnabled
      ? undefined
      : draft.dateMode === "auto"
      ? {
          mode: "auto",
          format: draft.dateFormat,
          locale: draft.dateLocale,
        }
      : draft.dateMode === "fixed"
        ? { mode: "fixed", text: draft.fixedDateText }
        : undefined;
  const fields = normalizeHeaderFooter({
    slideNumber: draft.slideNumber,
    footerText: draft.footerEnabled ? draft.footerText : undefined,
    dateField,
  });
  return fields
    ? {
        ...fields,
        ...(draft.excludeTitleSlide ? { excludeTitleSlide: true as const } : {}),
      }
    : undefined;
}

/**
 * Header & Footer's Apply to All (`slideIds` null) or Apply to the given
 * slides: Apply to All replaces the deck defaults and authenticates one
 * uniform result, so stale per-slide overrides must not keep winning; Apply
 * writes each slide's exact override and the title-slide choice.
 */
export function headerFooterAfterDialog(
  current: DeckHeaderFooter | undefined,
  draft: PresentationHeaderFooterDialogDraft,
  slideIds: readonly string[] | null,
): DeckHeaderFooter | undefined {
  const draftValue = presentationHeaderFooterDialogValue(draft);
  if (slideIds === null) {
    // Apply to All replaces the deck's footer, date and numbering. What the
    // dialog has no field for -- the first slide's number (Slide Size) and
    // the notes page size -- stays the deck's: replaced whole, a start at 5
    // went back to 1 with the footer (PPT-B329).
    const kept = normalizeHeaderFooter(current);
    if (
      kept?.firstSlideNumber === undefined &&
      kept?.notesSize === undefined &&
      kept?.notesHandouts === undefined
    ) {
      return draftValue;
    }
    return normalizeHeaderFooter({
      ...(draftValue ?? {}),
      ...(kept.firstSlideNumber === undefined ? {} : { firstSlideNumber: kept.firstSlideNumber }),
      ...(kept.notesSize === undefined ? {} : { notesSize: kept.notesSize }),
      // The Notes and Handouts page is the dialog's other page (PPT-B006).
      ...(kept.notesHandouts === undefined ? {} : { notesHandouts: kept.notesHandouts }),
    });
  }
  const exactOverride: SlideHeaderFooterOverride = {
    slideNumber: draftValue?.slideNumber === true,
    footerText: draftValue?.footerText ?? null,
    dateField: draftValue?.dateField ?? null,
  };
  const normalized = normalizeHeaderFooter(current);
  const slideOverrides = { ...(normalized?.slideOverrides ?? {}) };
  for (const slideId of slideIds) slideOverrides[slideId] = structuredClone(exactOverride);
  return normalizeHeaderFooter({
    ...(normalized ?? {}),
    excludeTitleSlide: draft.excludeTitleSlide ? true : undefined,
    slideOverrides,
  });
}

/**
 * The longest speaker notes the Notes pane keeps: the same ceiling the PPTX
 * import applies (`_MAX_PPTX_NOTES_CHARS`, editable_page.py), so a one-letter
 * edit of an imported script never cuts its tail on blur (PPT-B381).
 */
export const SPEAKER_NOTES_MAX_LENGTH = 10_000;

/**
 * The slide with the Notes pane's text committed: cut to the pane's length
 * and trimmed; empty notes are an absent field, so a cleared or never-noted
 * slide stays byte-clean. The slide itself when the text did not change.
 */
export function slideWithNotes<S extends PresentationButtonSlide>(slide: S, draft: string): S {
  const trimmed = draft.slice(0, SPEAKER_NOTES_MAX_LENGTH);
  const current = typeof slide.notes === "string" ? slide.notes : "";
  if (trimmed === current) return slide;
  const next = trimmed.trim();
  if (next) return { ...slide, notes: next };
  const cleared = { ...slide };
  delete cleared.notes;
  return cleared;
}
