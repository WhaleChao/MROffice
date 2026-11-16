/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The presentation editor's deck model and its load-time normalizers: the
 * `SlideModel` shape, the per-block normalizer that every stored, pasted or
 * generated block passes through, the slide-list normalizer, and the readers
 * that turn stored page content into the editor's deck (themes, masters,
 * sections, custom shows, sensitivity) and back (`buildContent`).
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change); that file re-exports the symbols it
 * used to export so existing import paths keep working.
 */

import { randomUUID } from "../random-uuid";
import { runSlideActionTarget } from "../presentation-runs";
import {
  normalizePresentationViewSettings,
  type PresentationViewSettings,
} from "../presentation-grid";
import { officeFillPatternOf } from "../office-chart-shape-paint";
import type { OfficeFontFace } from "../../components/office-ribbon/officeFontFace";
import { repairBlockTextLayers } from "../presentation-collab-text";
import {
  normalizePresentationFontFaces,
  presentationFontFacesForPersist,
} from "../presentation-font-faces";
import {
  presentationChartBlockForSave,
  upgradePresentationChartBlock,
} from "../presentation-chart-upgrade";
import { hasAnyBullet, textFromParagraphs } from "../presentation-paragraphs";
import { normalizeRuns, textFromRuns } from "../presentation-point-runs";
import {
  PRESENTATION_CONTENT_TYPE,
  presentationPageContentObjectSchema,
  presentationPageContentSchema,
  presentationSlideSchema,
  refinePresentationPageContent,
} from "../native-page-schemas";
import { z } from "zod";
import { PRESENTATION_MEDIA_KEY } from "../presentation-page-media";
import { presentationListInPaintOrder } from "../presentation-paint-order";
import { PRESENTATION_DECK_RECEIPTS_KEY } from "../presentation-deck-receipts";
import {
  DEFAULT_SLIDE_TRANSITION,
  isDirectionalSlideTransition,
  isSlideTransitionDirectionForKind,
  isTransitionAdvanceAfterMs,
  normalizeSlideTransition,
  normalizeTransitionDuration,
  type SlideTransitionDirection,
  type SlideTransition,
  type SlideTransitionZoomElement,
} from "../presentation-transitions";
import {
  normalizePresentationSlideShowSettings,
  presentationSlideShowSettingsWithinSlides,
  type PresentationSlideShowSettings,
} from "./rehearse-timings";
import {
  normalizePresentationCustomShows,
  presentationCustomShowsForPersist,
  type PresentationCustomShows,
} from "../presentation-custom-shows";
import {
  STROKE_STYLES,
  DEFAULT_SHAPE_FILL,
  DEFAULT_SHAPE_STROKE,
  DEFAULT_STROKE_WIDTH,
  normalizeShapeColor,
  isOpenBasicShapeKind,
  isLineCalloutShapeKind,
  isPresentationConnectorShapeKind,
  normalizeShapeFillImage,
  normalizeShapeKind,
  normalizeStrokeWidth,
  presentationInkStrokeWidthPx,
  type StrokeStyle,
} from "../presentation-shapes";
import {
  normalizePresentationCustomPath,
  presentationCustomGeometryPathAdmitted,
  presentationCustomPathIsOpenLineOnly,
  presentationInkPointsWithinEnvelope,
} from "../presentation-freeform";
import {
  normalizePresentationPrintSettings,
  type PresentationPrintSettings,
} from "../presentation-print-settings";
import { normalizeBgPatternFill } from "../presentation-slide-bg";
import {
  buildOccupancyGrid,
  MAX_TABLE_SPLIT_COLUMNS,
  MAX_TABLE_SPLIT_ROWS,
  tableLogicalWidth,
} from "../presentation-table-merge";
import { normalizeAxisSizes } from "../presentation-table-size";
import { MAX_PRESENTATION_SLIDES } from "../presentation-conveniences";
import {
  IMAGE_EFFECT_KEYS,
  isImageRecolorMode,
  persistedImageEffects,
  pictureCorrectionsSourceForLoad,
} from "../presentation-image-effects";
import {
  CELL_STYLE_KEYS,
  normalizeTableStyleId,
  persistedCellStyle,
} from "../presentation-table-style";
import {
  assetBackedPresentationMediaSource,
  normalizeEmbedKind,
  resolveEmbed,
} from "../presentation-embed";
import {
  DEFAULT_LAYOUT_ID,
  layoutFrameSourceForTheme,
  layoutBlocks,
  blocksLeftOnSlide,
  normalizeSlideLayoutRef,
  type LayoutId,
  type OfficeMasterLayoutRef,
} from "../presentation-layouts";
import {
  DEFAULT_BLOCK_ANIMATION,
  isPresentationActionButtonShapeKind,
  presentationHasAbsentNativeAction,
  normalizeBlockAnimation,
  normalizePresentationClickAction,
  normalizePresentationRelativeClickAction,
  sanitizeBlockHref,
  type PresentationRelativeClickAction,
} from "../presentation-block-fx";
import { normalizeParagraphsWithRuns } from "../presentation-text-buttons";
import { presentationShapeOutlineIsPreserveOnly } from "../presentation-shape-buttons";
import {
  assignTimelineStepIds,
  migrateLegacyAnimationSequence,
  normalizeAnimationSequence,
  POWERPOINT_ANIMATION_TIMING,
  pruneSequence,
  type TimelineStep,
} from "../presentation-anim-timeline";
import { nativeAutofitResizesShape, normalizeAutofitMode } from "../presentation-autofit";
import { normalizeComments, type SlideComment } from "../presentation-comments";
import {
  defaultPresentationTheme,
  materializeCatalogThemeIdentity,
  normalizeTheme,
  type DeckTheme,
} from "../presentation-theme";
import {
  normalizeMasterLayoutId,
  normalizeMasterId,
  normalizeMasters,
  type SlideMaster,
} from "../presentation-master";
import { compactPresentationMedia, expandPresentationMedia } from "../presentation-page-media";
import { normalizeHeaderFooter, type DeckHeaderFooter } from "../presentation-header-footer";
import {
  sectionsInStoredForm,
  storedDeckSections,
  type SlideSection,
} from "../presentation-sections";
import { normalizeHiddenFlag } from "../presentation-slide-visibility";
import { isIdentityCrop, normalizeSignedCropRect } from "../presentation-crop-group";
import { presentationHasNativeShapeDrawing } from "../presentation-native-theme-paint";
import type {
  PresentationBlock,
  PresentationChartBlock,
  PresentationImageBlock,
  PresentationTableCell,
} from "../../components/console/editor/native-document/documentCommands";
import { MAX_PRESENTATION_SLIDE_BLOCKS } from "./slide-block-budget";
import { asRecord } from "../as-record";

/**
 * This editor's own insert gate. It was a literal 200 — the same number the
 * page schema, the page validator and the server projection each wrote out
 * separately, which is how a deck could arrive with 200 blocks on a slide and
 * never take one more even though the page had room for hundreds.
 *
 * It is now a bound on work only, shared with the page schema. What a slide
 * may really carry is its share of the page's bytes, and that share is spent
 * where the blocks are actually produced — the server projection — exactly as
 * the workbook does it (`SpreadsheetWindow` takes `SPREADSHEET_MAX_OBJECTS`,
 * the count, while `SPREADSHEET_OBJECT_BYTE_BUDGET` binds the import). Asking
 * the byte budget on a render path would re-measure the whole slide on every
 * keystroke, because React hands this component a new array each time.
 */
export const MAX_PRESENTATION_BLOCKS_PER_SLIDE = MAX_PRESENTATION_SLIDE_BLOCKS;

/**
 * Deck SENSITIVITY levels (PPT 敏感度): an information-classification label on
 * the whole deck. Persisted as `sensitivity` in content_json only when set;
 * the export writes it into docProps core properties as the category, which
 * is where Office tooling looks for a classification.
 */
const PRESENTATION_SENSITIVITY_LEVELS = [
  ["public", "公开"],
  ["internal", "内部"],
  ["confidential", "机密"],
  ["restricted", "绝密"],
] as const;

export type PresentationSensitivity =
  (typeof PRESENTATION_SENSITIVITY_LEVELS)[number][0];

export function normalizeSensitivity(
  value: unknown,
): PresentationSensitivity | undefined {
  return PRESENTATION_SENSITIVITY_LEVELS.some(([level]) => level === value)
    ? (value as PresentationSensitivity)
    : undefined;
}

export interface SlideModel {
  id: string;
  title: string;
  blocks: PresentationBlock[];
  source?: Record<string, unknown>;
  /**
   * A stub of a deck held as parts (lib/deck-page-parts): a slide of a part
   * the editor does not hold (still being imported, or not loaded yet). It
   * has no blocks, is never saved as a slide, and is only moved, hidden or
   * deleted until its part arrives in its place.
   */
  importPending?: boolean;
  notes?: string;
  backgroundColor?: string;
  /** PowerPoint solid-fill transparency percentage; omission is the exact 0%. */
  backgroundTransparency?: number;
  /**
   * Optional RICH background fill (gradient / picture / pattern) overriding the
   * solid `backgroundColor` (and the theme) at render time. Resolved by
   * `lib/presentation-slide-bg`; absent ⇒ the solid colour path, byte-stable.
   */
  backgroundGradient?: { from?: string; to?: string; angle?: number };
  backgroundImage?: { url?: string; fit?: string };
  backgroundPattern?: { pattern?: string; color?: string; background?: string };
  /** Hide layout/master graphics without changing the slide's own blocks. */
  hideBackgroundGraphics?: boolean;
  /**
   * Hide / skip this slide in a slideshow (PowerPoint / WPS "Hide Slide"). The
   * slide stays in the deck (edits + serializes) but present-mode advance /
   * retreat / jump skip it. Absent / `false` → a normal, presentable slide.
   */
  hidden?: boolean;
  /** Present-mode entrance transition. Absent / `none` plays no animation. */
  transition?: SlideTransition;
  /**
   * Present-mode transition DURATION (ms). Absent → the CSS default duration
   * (byte-for-byte as before); when set it drives an inline animation-duration
   * override, clamped to [MIN..MAX]_TRANSITION_DURATION_MS at the render edge.
   */
  transitionDurationMs?: number;
  /** Native Push/Wipe direction or Shape geometry Effect Options identity. */
  transitionDirection?: SlideTransitionDirection;
  /** OOXML `advClick`; omission is the default `true`, so persist false only. */
  transitionAdvanceOnClick?: boolean;
  /** OOXML `advTm` in milliseconds; omission means automatic advance is off. */
  transitionAdvanceAfterMs?: number;
  /**
   * Read only: when each click came in PowerPoint's Rehearse Timings (the
   * slide's tag TIMING), ms, each after the click before (the first after
   * the slide's start). Use Timings and Create a Video play the clicks then
   * (`slide-click-schedule`).
   */
  rehearsedClickTimingsMs?: readonly number[];
  /**
   * Which element spells this slide's Zoom. PowerPoint's own Zoom gallery
   * entry writes Office 2010's `p14:warp`, and so does a Zoom picked new in
   * this editor; absent means the 2006 `p:zoom` an older slide was written
   * with. Carried so an edit writes the element the file used.
   */
  transitionZoomElement?: SlideTransitionZoomElement;
  /**
   * Master-slide layout id this slide was created from / arranged to. A slide
   * on an inserted master's layout carries `<master key>--<layout id>`
   * (presentation-office-master-model.ts: officeMasterLayoutRef).
   */
  layout?: LayoutId | OfficeMasterLayoutRef;
  /**
   * Stable native child-layout reference. Imported PPTX decks retain the OOXML
   * part path; locally-created masters use an opaque generated id.
   */
  layoutId?: string;
  /** Deck-level master this slide inherits layout + style overrides from. */
  masterId?: string;
  /**
   * Present-mode ANIMATION TIMELINE: an ORDERED sequence of animation steps
   * (PowerPoint animation-pane parity). Absent / empty plays no scripted
   * sequence (per-block `block.animation` entrance fx still apply on enter).
   */
  animationSequence?: TimelineStep[];
  






  animationTiming?: typeof POWERPOINT_ANIMATION_TIMING;
  /**
   * An imported slide's animations no reader recognised, laid out for
   * playback by the server (`animation_timing_tree`, render only): played by
   * `presentation-native-timing-player` when the slide has no sequence.
   */
  nativeTimingTree?: unknown;
  /**
   * Persistent REVIEW COMMENTS (PowerPoint / WPS "Review → Comments"): a threaded
   * comment list pinned to this slide (or to a specific object via a comment's
   * `blockId`). Persisted on the deck — distinct from the ephemeral present-mode
   * ink. Absent / empty ⇒ no comments (byte-stable). Normalized at the load edge
   * by `lib/presentation-comments` (`normalizeComments`).
   */
  comments?: SlideComment[];
  [key: string]: unknown;
}

export function defaultSlide(): SlideModel {
  return {
    id: randomUUID(),
    title: "",
    layout: "title",
    // PowerPoint and WPS start a new deck on a real Title Slide layout. Using
    // positioned placeholders here keeps the editor in canvas semantics from
    // the first frame instead of falling back to the document-like title/body
    // form that made an empty presentation look like a note editor.
    // The new deck's theme (an Office theme by default) places them where
    // PowerPoint's Title Slide would.
    blocks: layoutBlocks(
      "title",
      undefined,
      layoutFrameSourceForTheme(defaultPresentationTheme()),
    ) as unknown as PresentationBlock[],
  };
}

/**
 * A deck is "pristine" when it still looks factory-new: exactly one untitled
 * slide whose blocks are all empty text. Whole-deck inserts (AI generation /
 * design templates) REPLACE a pristine deck and APPEND to anything else so an
 * authored deck — including one holding only images/shapes — is never clobbered.
 */
export function isPristineDeck(slides: SlideModel[]): boolean {
  return (
    slides.length === 1 &&
    !slides[0].title &&
    slides[0].blocks.every((block) => block.type === "text" && !block.text)
  );
}

/**
 * One fully-undoable unit of deck state: the slides AND the deck-level masters
 * captured together. Every history entry holds a {@link DeckSnapshot} so undo /
 * redo restores BOTH dimensions consistently — a master mutation (create /
 * rename / layout / theme / delete / assign) and the slide edits it triggers
 * (e.g. clearing a deleted master's dangling `masterId`) revert as one step.
 */
export interface DeckSnapshot {
  fontFaces?: OfficeFontFace[];
  slides: SlideModel[];
  masters: SlideMaster[];
  /**
   * Active slide identity at the time of the snapshot. Structural edits can
   * shift indices, so undo / redo restores the exact slide rather than merely
   * clamping the post-mutation index into the restored deck.
   */
  activeSlideId: string | undefined;
  /**
   * Deck-level sections at the time of the snapshot. Captured alongside slides +
   * masters so undo / redo reverts the section dimension consistently (e.g. a
   * slide delete that re-anchors a section, or an add/rename/delete-section
   * step, reverts as ONE history entry).
   */
  sections: SlideSection[];
  /**
   * Deck theme at the time of the snapshot. A commit that ALSO changes the theme
   * (applying a template bundles a new background/fonts/colors WITH a slide
   * change) captures the PRE-apply theme here so undo reverts both as ONE step.
   * A theme-unchanged commit captures the current theme, so reverting it is a
   * no-op (a plain slide edit is unaffected).
   */
  theme: DeckTheme | undefined;
  /** Deck defaults + per-slide Header & Footer overrides, restored atomically. */
  headerFooter: DeckHeaderFooter | undefined;
  /** Set Up Slide Show's native choices, restored atomically. */
  slideShowSettings: PresentationSlideShowSettings | undefined;
  /** Named custom-show registries and their exact playback order. */
  customSlideShows: PresentationCustomShows;
}

export { asRecord };








function blockHrefKept(value: unknown): string | null {
  return typeof value === "string" && runSlideActionTarget(value.trim()) ? value.trim() : sanitizeBlockHref(value);
}

function omitRecordKeys(
  record: Record<string, unknown>,
  keys: ReadonlySet<string>,
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(record).filter(([key]) => !keys.has(key)),
  );
}

/** Keys handled explicitly by {@link normalizeBlockFx}; never copied via extras. */
const BLOCK_FX_KEYS: ReadonlySet<string> = new Set([
  "animation",
  "href",
  "clickAction",
  "locked",
]);

/**
 * Resolve the shared block-level effects (entrance animation + hyperlink) from a
 * raw persisted record into a normalized, render-safe partial. Unknown / stale
 * animations collapse to the default (and are omitted so default decks stay
 * byte-clean); an unsafe or malformed `href` (e.g. `javascript:`) is dropped
 * entirely so it can never reach the renderer. Spread onto every normalized
 * block so all block kinds support the same effects.
 */






function wrappedTurnAngle(angle: number): number {
  // Only an angle out of the turn moves: `(a % 360 + 360) % 360` turned the
  // importer's 219.9 into 219.89999999999998.
  if (angle >= 0 && angle < 360) return angle;
  const wrapped = angle % 360;
  const turned = wrapped < 0 ? wrapped + 360 : wrapped;
  return turned >= 360 ? 0 : turned;
}

function normalizeBlockFx(record: Record<string, unknown>): {
  animation?: string;
  href?: string;
  clickAction?: PresentationRelativeClickAction;
  locked?: true;
} {
  const fx: {
    animation?: string;
    href?: string;
    clickAction?: PresentationRelativeClickAction;
    locked?: true;
  } = {};
  const animation = normalizeBlockAnimation(record.animation);
  if (animation !== DEFAULT_BLOCK_ANIMATION) fx.animation = animation;
  const href = blockHrefKept(record.href);
  if (href) fx.href = href;
  // Action Settings' slide jump on any object (a text box, a picture, a
  // chart, a table): it was dropped on reopen and gone with the next save.
  // A shape states its own (media actions are an Action Button's).
  const clickAction = normalizePresentationRelativeClickAction(record.clickAction);
  if (clickAction && !href) fx.clickAction = clickAction;
  // LOCK: persist ONLY when strictly `true` so an unlocked block carries no
  // `locked` key (byte-clean) and a stray truthy value never sticks as a lock.
  if (record.locked === true) fx.locked = true;
  return fx;
}

function normalizeTableCell(value: unknown): PresentationTableCell {
  if (typeof value === "string" || typeof value === "number") {
    return { text: String(value) };
  }
  const record = asRecord(value) ?? {};
  const { text, colspan, rowspan, ...extras } = record;
  for (const key of CELL_STYLE_KEYS) delete extras[key];
  return {
    ...extras,
    text: String(text ?? ""),
    ...(typeof colspan === "number" && colspan > 1 ? { colspan } : {}),
    ...(typeof rowspan === "number" && rowspan > 1 ? { rowspan } : {}),
    // Persist per-cell formatting only when non-default (mirrors the
    // shape/image normalization) so an un-styled cell stays `{ text }` clean.
    ...persistedCellStyle(record),
  };
}

function normalizeChartData(
  value: unknown,
): NonNullable<PresentationChartBlock["data"]> | undefined {
  const record = asRecord(value);
  if (!record) return undefined;
  const headers = Array.isArray(record.headers)
    ? record.headers
        .map((header) => String(header ?? "").slice(0, 500))
        .slice(0, 20)
    : [];
  const rows = Array.isArray(record.rows)
    ? record.rows
        .filter((row): row is unknown[] => Array.isArray(row))
        .slice(0, 200)
        .map((row) =>
          row.map((cell) => String(cell ?? "").slice(0, 2_000)).slice(0, 20),
        )
    : [];
  if (!headers.length && !rows.length) return undefined;
  return {
    ...(headers.length ? { headers } : {}),
    ...(rows.length ? { rows } : {}),
  };
}

const HOISTED_TEXT_STYLE_KEYS = [
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
  "letterSpacing",
  "writingMode",
  "autofit",
  "align",
  "verticalAlign",
  "anchorCenter",
  // The importer's box-wide indent (every row's `a:pPr@marL`, nearest
  
  // title sits 284pt in from its frame in PowerPoint and flush here.
  "indentLevel",
  "textShadow",
  // The words' shadow as they state it (PPTX readings row 805), beside
  // `textShadow`: drawn by `presentationTextShadowCss` from `style`.
  "textShadowEffect",
  // `a:bodyPr@wrap="none"`: the importer states it at the top level; the
  // text drew wrapped without it (PPT-B360).
  "textWrap",
  // Custom tab stops (`a:tabLst`, PPT-B366): the importer states them at the
  // top level, `textBlockTabSizeCss` reads `style.tabStops`, so an imported
  
  "tabStops",
  // The paragraphs' default tab interval (`@defTabSz`), where it is not an inch.
  "defaultTabPx",
  // `a:bodyPr@spcFirstLastPara`: whether the first row's Spacing Before is
  
  "spaceFirstLastPara",
  // `a:pPr@latinLnBrk="1"` on every row: Latin words break anywhere.
  "latinLineBreak",
] as const;

function hoistImportedTextStyle(
  record: Record<string, unknown>,
): Record<string, unknown> {
  if (!HOISTED_TEXT_STYLE_KEYS.some((key) => key in record)) return record;
  const style = {
    ...(asRecord(record.style) ?? {}),
  } as Record<string, unknown>;
  const next = { ...record };
  for (const key of HOISTED_TEXT_STYLE_KEYS) {
    if (key in next) {
      if (!(key in style)) style[key] = next[key];
      delete next[key];
    }
  }
  next.style = style;
  return next;
}

/** Keep autofit in its one canonical block-level save/inspector location. */
function canonicalizeTextAutofit(
  record: Record<string, unknown>,
): Record<string, unknown> {
  const style = asRecord(record.style);
  const hasStyledAutofit = Boolean(style && "autofit" in style);
  if (!hasStyledAutofit && !("autofit" in record)) return record;

  const effectiveAutofit = normalizeAutofitMode(
    hasStyledAutofit ? style?.autofit : record.autofit,
  );
  const next = { ...record };
  if (style && "autofit" in style) {
    const nextStyle = { ...style };
    delete nextStyle.autofit;
    if (Object.keys(nextStyle).length > 0) next.style = nextStyle;
    else delete next.style;
  }
  delete next.autofit;
  const persisted = persistedAutofitMode(effectiveAutofit, record);
  if (persisted !== undefined) next.autofit = persisted;
  return next;
}

/**
 * The `autofit` a text block stores: Shrink text on overflow and Resize shape
 * to fit text are stated; Do not Autofit is stated only where it overrides
 * an imported `a:spAutoFit` (`effectiveAutofitMode`), else omitted so an
 * un-opted block stays byte-clean.
 */
function persistedAutofitMode(
  mode: ReturnType<typeof normalizeAutofitMode>,
  record: Record<string, unknown>,
): "shrink" | "resize-shape" | "none" | undefined {
  if (mode === "shrink" || mode === "resize-shape") return mode;
  return nativeAutofitResizesShape(record) ? "none" : undefined;
}













export interface PresentationReadIssue {
  readonly slideId?: string;
  readonly blockId: string;
  readonly code: string;
}

/** Where a degraded block keeps the stored settings it could not honour. */
export const PRESENTATION_READ_PRESERVED_KEY = "__readPreserved";

interface PresentationBlockReadOptions {
  /** Read a contradictory block instead of throwing; say what was set aside. */
  readonly onIssue?: (code: string) => void;
}

function readIssue(options: PresentationBlockReadOptions, code: string): void {
  if (!options.onIssue) throw new Error(code);
  options.onIssue(code);
}

export function normalizePresentationBlock(
  value: unknown,
  fallbackIndex: number,
  options: PresentationBlockReadOptions = {},
): PresentationBlock {
  const rawRecord = asRecord(value) ?? {};
  const hoistedRecord =
    rawRecord.type === "text" || rawRecord.type === undefined
      ? hoistImportedTextStyle(rawRecord)
      : rawRecord;
  const record =
    rawRecord.type === "text" || rawRecord.type === undefined
      ? canonicalizeTextAutofit(hoistedRecord)
      : hoistedRecord;
  const id = String(record.id || randomUUID());
  if (record.type === "image" && typeof record.src === "string" && record.src) {
    const {
      src,
      altTitle,
      alt,
      decorative,
      caption,
      width,
      height,
      cropRect,
    } = record;
    const extras = omitRecordKeys(
      record,
      new Set([
        "id",
        "type",
        "src",
        "altTitle",
        "alt",
        "decorative",
        "caption",
        "width",
        "height",
        "cropRect",
        ...IMAGE_EFFECT_KEYS,
        ...BLOCK_FX_KEYS,
      ]),
    );
    const canonicalCorrectionsSource = pictureCorrectionsSourceForLoad(
      record.source,
      record.brightness,
      record.contrast,
    );
    if (canonicalCorrectionsSource !== record.source) {
      extras.source = canonicalCorrectionsSource;
    }
    // Persisted Fill/Fit can carry signed source-space outsets. Validate those
    // without the free-gesture clamp; identity/invalid input is omitted so a
    // default deck remains byte-clean and malformed CSS never reaches render.
    const signedCrop = normalizeSignedCropRect(cropRect);
    const persistedCrop =
      signedCrop && !isIdentityCrop(signedCrop) ? { cropRect: signedCrop } : {};
    const unsupportedLegacyRecolor =
      record.recolor !== undefined && !isImageRecolorMode(record.recolor)
        ? {
            recolor: record.recolor as PresentationImageBlock["recolor"],
          }
        : {};
    return {
      ...extras,
      id,
      type: "image",
      src,
      ...(typeof altTitle === "string" ? { altTitle } : {}),
      ...(typeof alt === "string" ? { alt } : {}),
      ...(typeof decorative === "boolean" ? { decorative } : {}),
      ...(typeof caption === "string" ? { caption } : {}),
      ...(typeof width === "number" && width > 0 ? { width } : {}),
      ...(typeof height === "number" && height > 0 ? { height } : {}),
      ...persistedCrop,
      // Persist image effects only when meaningful (non-default) so an image
      // without effects serializes byte-for-byte as before.
      ...persistedImageEffects(record),
      // Keep any unknown pre-schema value through unrelated normalization. The
      // known legacy recolors above are preserved by persistedImageEffects;
      // native Picture Color still accepts only its separate two-value domain.
      ...unsupportedLegacyRecolor,
      ...normalizeBlockFx(record),
    };
  }
  if (record.type === "table" && Array.isArray(record.rows)) {
    const {
      rows,
      caption,
      borders,
      headerRow,
      colWidths,
      rowHeights,
      tableStyleId,
    } = record;
    const extras = omitRecordKeys(
      record,
      new Set([
        "id",
        "type",
        "rows",
        "caption",
        "borders",
        "headerRow",
        "colWidths",
        "rowHeights",
        "tableStyleId",
        ...BLOCK_FX_KEYS,
      ]),
    );
    const cellRows = rows
      .filter((row): row is unknown[] => Array.isArray(row))
      .slice(0, MAX_TABLE_SPLIT_ROWS)
      .map((row) => row.slice(0, MAX_TABLE_SPLIT_COLUMNS).map(normalizeTableCell));
    // A row with no cell of its own is still a LOGICAL row when a rowspan from
    // above covers it (a 2x2 table merged into one cell keeps two rows; its
    // second row array is legitimately empty). Only a row nothing reaches is
    // junk, and dropping it drops its height with it so the axes stay aligned.
    const occupancy = buildOccupancyGrid(cellRows);
    const keptRowIndexes = cellRows
      .map((_row, index) => index)
      .filter(
        (index) =>
          cellRows[index].length > 0 ||
          (occupancy?.[index]?.some((slot) => slot !== null) ?? false),
      );
    const normalizedRows = keptRowIndexes.map((index) => cellRows[index]);
    if (normalizedRows.length > 0) {
      // Bound + trim the optional per-axis sizes against the LOGICAL grid shape
      // (column count = the colspan/rowspan-aware width, so a row merged into
      // one wide cell keeps every hidden column's width; row count = surviving
      // rows). An all-auto / absent array drops out so an un-sized table stays
      // byte-for-byte clean.
      const columnCount = tableLogicalWidth(normalizedRows);
      const normalizedColWidths = normalizeAxisSizes(colWidths, columnCount);
      const normalizedRowHeights = normalizeAxisSizes(
        Array.isArray(rowHeights)
          ? keptRowIndexes.map((index) => rowHeights[index])
          : rowHeights,
        normalizedRows.length,
      );
      const normalizedTableStyleId = normalizeTableStyleId(tableStyleId);
      return {
        ...extras,
        id,
        type: "table",
        rows: normalizedRows,
        ...(typeof caption === "string" ? { caption } : {}),
        // Persist the table-level flags only when on so an un-styled table
        // serializes byte-for-byte as before.
        ...(borders === true ? { borders: true } : {}),
        ...(headerRow === true ? { headerRow: true } : {}),
        ...(normalizedTableStyleId
          ? { tableStyleId: normalizedTableStyleId }
          : {}),
        ...(normalizedColWidths ? { colWidths: normalizedColWidths } : {}),
        ...(normalizedRowHeights ? { rowHeights: normalizedRowHeights } : {}),
        ...normalizeBlockFx(record),
      };
    }
  }
  if (record.type === "shape") {
    const {
      shape,
      fill,
      fillTo,
      gradientAngle,
      fillImage,
      fillPattern,
      stroke,
      strokeTo,
      strokeGradientAngle,
      strokeWidth,
      strokeStyle,
      shadow,
      customPath,
      clickAction,
    } = record;
    const extras = omitRecordKeys(
      record,
      new Set([
        "id",
        "type",
        "shape",
        "fill",
        "fillTo",
        "gradientAngle",
        "fillImage",
        "fillPattern",
        "stroke",
        "strokeTo",
        "strokeGradientAngle",
        "strokeWidth",
        "strokeStyle",
        "shadow",
        "customPath",
        "clickAction",
        ...BLOCK_FX_KEYS,
      ]),
    );
    const kind = normalizeShapeKind(shape);
    const normalizedHref = blockHrefKept(record.href);
    const normalizedClickAction = normalizePresentationClickAction(clickAction);
    const isActionButton = isPresentationActionButtonShapeKind(kind);
    const absentNativeAction = presentationHasAbsentNativeAction(record);
    const hasRawHref = record.href !== undefined;
    const hasRawClickAction = clickAction !== undefined;
    let preservedClickAction = false;
    let preservedCustomPath = false;
    if (
      (hasRawHref && hasRawClickAction) ||
      (hasRawClickAction && normalizedClickAction === null) ||
      (isActionButton && hasRawHref && normalizedHref === null) ||
      (isActionButton && !hasRawHref && !hasRawClickAction && !absentNativeAction) ||
      ((kind === "action-button-movie" ||
        kind === "action-button-sound") &&
        normalizedClickAction?.kind !== "play-media" && !absentNativeAction) ||
      // Any shape takes a slide action (Action Settings, PowerPoint's for
      // every object); only the Movie / Sound buttons play media. A plain
      // rectangle jumping to the previous slide made the whole deck
      
      (!isActionButton && normalizedClickAction?.kind === "play-media")
    ) {
      readIssue(options, "presentation_action_button_target_invalid");
      preservedClickAction = true;
    }
    if (!preservedClickAction && normalizedClickAction?.kind === "play-media") {
      const expectedKind =
        kind === "action-button-movie"
          ? "video"
          : kind === "action-button-sound"
            ? "audio"
            : null;
      if (normalizedClickAction.mediaKind !== expectedKind) {
        readIssue(options, "presentation_action_button_media_invalid");
        preservedClickAction = true;
      }
    }
    const normalizedCustomPath = normalizePresentationCustomPath(customPath);
    const hasCustomPath = customPath !== undefined;
    const hasUnsupportedCustomGeometryText =
      (kind === "ink" || kind === "custom-geometry") &&
      ((typeof record.text === "string" && record.text.length > 0) ||
        (Array.isArray(record.runs) && record.runs.length > 0));
    if (hasUnsupportedCustomGeometryText) {
      // The package writer intentionally has no text-bearing custGeom route.
      // Refuse stale/raw-page fallback data here instead of mounting an editor
      // operation whose text would disappear on export.
      readIssue(options, "presentation_custom_geometry_text_unsupported");
    }
    const customPathCombinationValid =
      kind === "custom-geometry"
        ? normalizedCustomPath !== null &&
          presentationCustomGeometryPathAdmitted(normalizedCustomPath)
        : kind === "ink"
          ? !hasCustomPath ||
            (normalizedCustomPath !== null &&
              presentationCustomPathIsOpenLineOnly(normalizedCustomPath))
          : !hasCustomPath;
    if (!customPathCombinationValid) {
      // `presentationPageContentSchema.safeParse` deliberately has a raw-page
      // fallback for older decks, so this normalizer is a second authority
      // boundary. Refuse contradictory custom geometry rather than mounting a
      // false rectangle or dropping the path on the next autosave.
      readIssue(options, "presentation_custom_geometry_invalid");
      preservedCustomPath = true;
    }
    // Highlighter ink used to be stored with a 0-1 fraction in this 0-100
    // percentage, which renders as an invisible 0; read such a stroke back as
    // the percentage it meant. Nothing else writes a fractional ink opacity
    // (PPT-B157).
    const legacyInkOpacity =
      kind === "ink" &&
      typeof extras.opacity === "number" &&
      extras.opacity > 0 &&
      extras.opacity < 1
        ? Math.round(extras.opacity * 100)
        : undefined;
    // Record Studio stored its ink's width in EMU, 9525 times the design
    // pixels every other ink states (a pen 41 148 wide on a 16:9 slide), and
    // every sample of its gesture. An ink wider than the thickest outline is
    // such a stroke (nothing else writes one: the Draw tab, Shape Format and
    // the importer all stay within it), read back as the pixels it meant; its
    // points are held to the ink envelope the canvas draws and the server
    // stores (`presentationInkPointsWithinEnvelope`). The page schema reads
    // such a stroke the same way, so a deck holding one still parses.
    const recordedInkWidth =
      kind === "ink" ? presentationInkStrokeWidthPx(strokeWidth) : undefined;
    const inkPoints =
      kind === "ink" ? presentationInkPointsWithinEnvelope(extras.points) : extras.points;
    const isLinear =
      isPresentationConnectorShapeKind(kind) ||
      kind === "ink" ||
      (kind === "custom-geometry" && normalizedCustomPath?.closed === false);
    // A Cameo is a live camera frame: PowerPoint gives it no shape fill and no
    // default outline (Camera Format offers only a border the user picks), and
    // the page validator refuses fill, gradient and transparency on it. The
    // default colours below were added on every reopen, so a deck holding a
    // Cameo could never be saved again (acceptance 51).
    const isCameo = record.cameo === true;
    // A picture fill is meaningful only on a box / polygon shape (never a line);
    // an empty / unsafe url normalizes to null (dropped, byte-clean).
    const normalizedFillImage = isLinear || isCameo
      ? null
      : normalizeShapeFillImage(fillImage);
    // Office's preset form (the Format pane's 48 patterns) is kept as it is;
    // read as the slide background's four, a reopened Small Checker came
    
    const normalizedFillPattern = isLinear || isCameo
      ? undefined
      : fillPattern && typeof fillPattern === "object" && "preset" in fillPattern
        ? officeFillPatternOf(fillPattern) ?? undefined
        : normalizeBgPatternFill(fillPattern);
    // A shape painted with a picture or a pattern states its fill already: a
    // default colour invented under it draws nothing, and on an imported
    // layout / master shape (whose source carries no shapeId, so the native
    // drawing check above cannot vouch for it) it was saved as the user's
    
    const fillStatedByPaint =
      fill === undefined &&
      (normalizedFillImage !== null || normalizedFillPattern !== undefined);
    // An imported OPEN freeform DrawingML fills all the same (`openPathFilled`,
    
    
    // two navy circles whose paths end without `a:close`).
    const openPathFilled =
      kind === "custom-geometry" && normalizedCustomPath?.closed === false && record.openPathFilled === true;
    const normalizedFill = isCameo || (fill === undefined && presentationHasNativeShapeDrawing(record)) || fillStatedByPaint || (isLinear && !openPathFilled) || (isOpenBasicShapeKind(kind) && fill !== "none") || (isLineCalloutShapeKind(kind) && fill === undefined)
      ? null
      : normalizeShapeColor(fill, DEFAULT_SHAPE_FILL);
    const outlineIsPreserveOnly =
      presentationShapeOutlineIsPreserveOnly(record);
    const normalizedStroke = normalizeShapeColor(
      stroke,
      outlineIsPreserveOnly || isCameo ? null : DEFAULT_SHAPE_STROKE,
    );
    const normalizedStrokeTo =
      normalizedStroke && normalizedStroke !== "none" && !isCameo
        ? normalizeShapeColor(strokeTo, null)
        : null;
    // A gradient is only meaningful on a filled box shape with a distinct stop.
    const normalizedFillTo =
      normalizedFill === null ? null : normalizeShapeColor(fillTo, null);
    const normalizedWidth = normalizeStrokeWidth(recordedInkWidth ?? strokeWidth);
    const normalizedStrokeStyle: StrokeStyle =
      typeof strokeStyle === "string" &&
      (STROKE_STYLES as readonly string[]).includes(strokeStyle)
        ? (strokeStyle as StrokeStyle)
        : "solid";
    // Wrap (not clamp) to [0, 360) so the persisted angle matches the value the
    // renderer derives via normalizeGradientAngle — otherwise 360 would store as
    // 360 but render as 0, mutating the deck on a no-op round trip. Kept to the
    // hundredth the importer reads (`a:lin@ang` is 60000ths of a degree): a
    // file's 219.9° rounded to 220° on load read as an edit to the gradient,
    // which no in-place writer expresses, and every save of the deck -- the
    
    
    const normalizedAngle =
      typeof gradientAngle === "number" && Number.isFinite(gradientAngle)
        ? wrappedTurnAngle(gradientAngle)
        : undefined;
    const normalizedStrokeAngle =
      typeof strokeGradientAngle === "number" &&
      Number.isFinite(strokeGradientAngle)
        ? wrappedTurnAngle(strokeGradientAngle)
        : undefined;
    return {
      ...extras,
      ...(legacyInkOpacity !== undefined ? { opacity: legacyInkOpacity } : {}),
      ...(typeof inkPoints === "string" && inkPoints !== extras.points ? { points: inkPoints } : {}),
      id,
      type: "shape",
      shape: kind,
      ...(normalizedFill ? { fill: normalizedFill } : {}),
      ...(normalizedFillTo ? { fillTo: normalizedFillTo } : {}),
      ...(normalizedFillTo && normalizedAngle !== undefined
        ? { gradientAngle: normalizedAngle }
        : {}),
      // Persist a validated picture fill (PowerPoint "Picture or texture fill");
      // an absent / unsafe one is dropped so a colour-fill shape stays byte-clean.
      ...(normalizedFillImage ? { fillImage: normalizedFillImage } : {}),
      ...(normalizedFillPattern ? { fillPattern: normalizedFillPattern } : {}),
      ...(normalizedStroke ? { stroke: normalizedStroke } : {}),
      ...(normalizedStrokeTo ? { strokeTo: normalizedStrokeTo } : {}),
      ...(normalizedStrokeTo && normalizedStrokeAngle !== undefined
        ? { strokeGradientAngle: normalizedStrokeAngle }
        : {}),
      // Keep an explicit finite weight, including the default, so history
      // replay does not change an acknowledged shape or SmartArt seed.
      ...((!outlineIsPreserveOnly || typeof strokeWidth === "number") &&
      (normalizedWidth !== DEFAULT_STROKE_WIDTH ||
        typeof strokeWidth === "number" && Number.isFinite(strokeWidth))
        ? { strokeWidth: normalizedWidth }
        : {}),
      ...((!outlineIsPreserveOnly || typeof strokeStyle === "string") &&
      normalizedStrokeStyle !== "solid"
        ? { strokeStyle: normalizedStrokeStyle }
        : {}),
      ...(shadow === true ? { shadow: true } : {}),
      ...(normalizedCustomPath && !preservedCustomPath ? { customPath: normalizedCustomPath } : {}),
      ...normalizeBlockFx(record),
      ...(normalizedClickAction && !preservedClickAction ? { clickAction: normalizedClickAction } : {}),
      ...(preservedClickAction || preservedCustomPath
        ? {
            [PRESENTATION_READ_PRESERVED_KEY]: {
              ...(preservedClickAction && hasRawClickAction ? { clickAction } : {}),
              ...(preservedCustomPath && hasCustomPath ? { customPath } : {}),
            },
          }
        : {}),
    };
  }
  if (record.type === "chart") {
    const { title, caption, chartType, ranges, data } = record;
    const extras = omitRecordKeys(
      record,
      new Set([
        "id",
        "type",
        "title",
        "caption",
        "chartType",
        "ranges",
        "data",
        ...BLOCK_FX_KEYS,
      ]),
    );
    const normalizedRanges = Array.isArray(ranges)
      ? ranges
          .map((range) => String(range || "").trim())
          .filter(Boolean)
          .slice(0, 50)
      : [];
    // One chart model: a block stored before Chart V2 is upgraded as the
    // deck loads it (`upgradePresentationChartBlock`), so every chart the
    // editor holds has `chartV2` and the one ribbon, canvas and workbook.
    return upgradePresentationChartBlock({
      ...extras,
      id,
      type: "chart",
      ...(typeof title === "string" ? { title } : {}),
      ...(typeof caption === "string" ? { caption } : {}),
      ...(typeof chartType === "string" ? { chartType } : {}),
      ...(normalizedRanges.length > 0 ? { ranges: normalizedRanges } : {}),
      ...(normalizeChartData(data) ? { data: normalizeChartData(data) } : {}),
      ...normalizeBlockFx(record),
    } as PresentationChartBlock);
  }
  if (record.type === "embed") {
    const { title } = record;
    const hasAssetCustody = [
      "sourceKind",
      "assetId",
      "mimeType",
      "sizeBytes",
    ].some((field) => field in record);
    const assetMedia = assetBackedPresentationMediaSource(record);
    if (hasAssetCustody && !assetMedia) {
      // Its asset fields stay on the block as stored; it is drawn from its
      // address (or as an unplayable clip) and can be replaced.
      readIssue(options, "presentation_video_asset_invalid");
    }
    const extras = omitRecordKeys(
      record,
      new Set(["id", "type", "kind", "url", "src", "title", ...BLOCK_FX_KEYS]),
    );
    // Resolve the kind first, then the URL against THAT kind. A persisted URL
    // that is unsafe/missing for the resolved kind yields `url: null`; we store
    // the sanitized (or original-string) value verbatim so the block round-
    // trips and the render edge shows a fallback rather than a broken element.
    // `kind`/`url` accept the canonical names or the `src` alias an imported
    // deck may have used.
    const resolved = resolveEmbed(record);
    const rawUrl =
      typeof record.url === "string"
        ? record.url
        : typeof record.src === "string"
          ? record.src
          : "";
    return {
      ...extras,
      id,
      type: "embed",
      kind: assetMedia ? assetMedia.kind : resolved.kind,
      // Prefer the sanitized URL; if the persisted value is unsafe for this
      // kind keep the raw string so the editor input still shows what the
      // author typed (the renderer re-sanitizes and refuses to mount it).
      url: assetMedia ? "" : (resolved.url ?? rawUrl),
      ...(assetMedia
        ? {
            sourceKind: assetMedia.sourceKind,
            assetId: assetMedia.assetId,
            mimeType: assetMedia.mimeType,
            sizeBytes: assetMedia.sizeBytes,
          }
        : {}),
      ...(typeof title === "string" ? { title } : {}),
      ...normalizeBlockFx(record),
    };
  }
  const { text } = record;
  const extras = omitRecordKeys(
    record,
    new Set([
      "id",
      "type",
      "text",
      "paragraphs",
      "runs",
      "autofit",
      ...BLOCK_FX_KEYS,
    ]),
  );
  // Optional TEXT AUTOFIT mode (PowerPoint "Shrink text on overflow"): persisted
  // ONLY when `shrink` so an un-opted block stays byte-for-byte unchanged (no
  // `autofit` key); any other / garbage value normalizes to `none` = omitted.
  const autofitMode =
    record.autofit === undefined ? undefined : persistedAutofitMode(normalizeAutofitMode(record.autofit), record);
  const autofitPatch = autofitMode !== undefined ? { autofit: autofitMode } : {};
  // Optional list paragraphs persist ONLY when present + usable; a flat-text
  // block (no `paragraphs`) stays byte-for-byte unchanged. When paragraphs ARE
  // present the flat `text` is kept in sync (newline-joined) so any non-list
  // consumer (plain-text index, AI context, legacy export) still reads it.
  const normalizedParagraphs = normalizeParagraphsWithRuns(record.paragraphs);
  const flatText = String(
    text ?? (fallbackIndex === 0 ? (record.content ?? "") : ""),
  );
  if (normalizedParagraphs) {
    // The presentation-specific normalizer authenticates the core list fields
    // while retaining per-run formatting and per-paragraph DrawingML spacing.
    const withRuns = normalizedParagraphs;
    const anyParagraphRuns = withRuns.some((paragraph) =>
      Boolean(normalizeRuns(paragraph.runs)),
    );
    // An INDENTED-but-unbulleted block (paragraphs carry indent/level but no
    // visible marker) renders through the FLAT path (`!hasAnyBullet`), so its
    // per-run formatting lives in a BLOCK-level `runs` array (what the present
    // overlay + inspector preview read), NOT in per-paragraph `runs`. Re-attach
    // that block-level runs here so it survives the load edge; without this the
    // paragraphs branch silently drops it (data loss). Only when NO paragraph
    // carried its own runs (otherwise the per-paragraph model is authoritative)
    // and the block has no visible bullet (the per-paragraph model owns runs for
    // real lists). A flat-equivalent / plain block adds no `runs` (byte-stable).
    const blockRuns =
      !anyParagraphRuns && !hasAnyBullet(withRuns)
        ? normalizeRuns(record.runs)
        : undefined;
    return {
      ...extras,
      id,
      type: "text",
      text: textFromParagraphs(withRuns),
      paragraphs: withRuns,
      ...(blockRuns ? { runs: blockRuns } : {}),
      ...autofitPatch,
      ...normalizeBlockFx(record),
    };
  }
  // Flat (no paragraphs): optional block-level per-run `runs`. When present the
  // flat `text` is kept in sync as the runs' concatenation; a flat-equivalent
  // block (no real per-run variation) drops `runs` and stays byte-stable.
  const normalizedRuns = normalizeRuns(record.runs);
  return {
    ...extras,
    id,
    type: "text",
    text: normalizedRuns ? textFromRuns(normalizedRuns) : flatText,
    ...(normalizedRuns ? { runs: normalizedRuns } : {}),
    ...autofitPatch,
    ...normalizeBlockFx(record),
  };
}

function normalizeStoredSlide(
  slide: unknown,
  index: number,
  /** The whole deck read as its schema says (`presentationPageContentSchema`). */
  parsedSuccess: boolean,
  issues?: PresentationReadIssue[],
): SlideModel {
  const slideRecord = asRecord(slide) ?? {};
  const {
    blocks: rawBlocksValue,
    id,
    notes,
    title,
    transition,
    transitionDurationMs: rawTransitionDurationMs,
    transitionDirection: rawTransitionDirection,
    transitionAdvanceOnClick: rawTransitionAdvanceOnClick,
    transitionAdvanceAfterMs: rawTransitionAdvanceAfterMs,
    layout,
    layoutId,
    masterId,
    animationSequence: rawSequence,
    animationTiming: rawAnimationTiming,
    hidden: rawHidden,
    comments: rawComments,
    commentAuthority: rawCommentAuthority,
    ...slideExtras
  } = slideRecord;
  const rawBlocks = Array.isArray(rawBlocksValue) ? rawBlocksValue : [];
  const slideId = typeof id === "string" ? id : undefined;
  const blocks = rawBlocks
    .slice(0, MAX_PRESENTATION_BLOCKS_PER_SLIDE)
    .map((block, blockIndex) => {
      try {
        return normalizePresentationBlock(block, blockIndex);
      } catch {
        const codes: string[] = [];
        const read = normalizePresentationBlock(block, blockIndex, {
          onIssue: (code) => codes.push(code),
        });
        for (const code of codes) issues?.push({ slideId, blockId: read.id, code });
        return read;
      }
    });
  // Normalize the persisted animation timeline: drop/clamp invalid steps and
  // prune any step whose blockId no longer matches a block on this slide so
  // the persisted sequence never dangles. Only retain a non-empty result so
  // decks that never authored a timeline stay byte-clean (no key persisted).
  const blockIds = new Set(blocks.map((block) => block.id));
  for (const block of blocks) {
    if (block.type !== "shape") continue;
    const clickAction = block.clickAction;
    if (clickAction?.kind !== "play-media") continue;
    const target = blocks.find(
      (candidate) => candidate.id === clickAction.targetBlockId,
    );
    if (
      !target ||
      target.type !== "embed" ||
      normalizeEmbedKind(target.kind) !== clickAction.mediaKind ||
      !resolveEmbed({ kind: target.kind, url: target.url }).url
    ) {
      // The button stays, without the play action its clip cannot take;
      // the action is kept as stored and written back on save.
      const index = blocks.indexOf(block);
      const { clickAction: storedAction, ...rest } = block;
      const preserved = asRecord(
        (block as unknown as Record<string, unknown>)[PRESENTATION_READ_PRESERVED_KEY],
      );
      blocks[index] = {
        ...rest,
        [PRESENTATION_READ_PRESERVED_KEY]: { ...(preserved ?? {}), clickAction: storedAction },
      } as typeof block;
      issues?.push({
        slideId,
        blockId: block.id,
        code: "presentation_action_button_media_target_invalid",
      });
    }
  }
  // A step that was stored without an identity (decks from before steps
  // were born with one) gets a deterministic id derived from the slide,
  // its block and its position, so every collaborator loading this deck
  // derives the same ids and the animation list merges per step instead
  // of as one whole value. The ids persist through the normal save.
  const prunedSequence = pruneSequence(normalizeAnimationSequence(rawSequence), blockIds);
  const normalizedSequence = assignTimelineStepIds(
    // A slide saved under the timing before PowerPoint's: every step
    // keeps its start (`migrateLegacyAnimationSequence`).
    rawAnimationTiming === POWERPOINT_ANIMATION_TIMING
      ? prunedSequence
      : migrateLegacyAnimationSequence(prunedSequence),
    String(id || `slide-${index + 1}`),
  );
  // Normalize the persisted transition against the allowed set; only retain
  // it when it differs from the default so existing decks stay byte-clean.
  const normalizedTransition = normalizeSlideTransition(transition);
  // Normalize the persisted transition duration: a usable value is clamped
  // to the supported range; absent / non-numeric → undefined so a slide that
  // never set a duration persists NO field (CSS default, byte-clean).
  const normalizedTransitionDurationMs = normalizeTransitionDuration(
    rawTransitionDurationMs,
  );
  const normalizedTransitionDirection =
    isDirectionalSlideTransition(normalizedTransition) &&
    isSlideTransitionDirectionForKind(
      normalizedTransition,
      rawTransitionDirection,
    )
      ? rawTransitionDirection
      : undefined;
  const normalizedTransitionAdvanceOnClick =
    rawTransitionAdvanceOnClick === false ? false : undefined;
  const normalizedTransitionAdvanceAfterMs =
    isTransitionAdvanceAfterMs(rawTransitionAdvanceAfterMs)
      ? rawTransitionAdvanceAfterMs
      : undefined;
  // Normalize the persisted layout id against the allowed set; only retain
  // it when it differs from the default (`blank`) so decks that never chose a
  // layout stay byte-clean and a stale value can't reference a missing layout.
  const normalizedLayout =
    // An Office deck's custom layout id (PPT-B030) is kept as is.
    layout === undefined ? undefined : normalizeSlideLayoutRef(layout);
  // A native explicit blank is part of the source snapshot. Retain it so
  // loading and exporting do not manufacture a representation-only Save.
  // A slide that SAYS it is on the Blank layout keeps saying so too: an
  // empty one is an empty canvas, as PowerPoint's Blank layout has no
  // placeholder (`blocksLeftOnSlide`), and only the stated layout tells
  // it from a legacy slide that lost its last block.
  const keepStatedBlankLayout = layout === DEFAULT_LAYOUT_ID;
  // Normalize the persisted master reference: keep only a usable (non-empty,
  // length-capped) id, dropped otherwise so a stale value never lingers. A
  // dangling id (no matching master) degrades to no master at render time.
  const normalizedMasterId = normalizeMasterId(masterId);
  const normalizedNativeLayoutId = normalizeMasterLayoutId(layoutId);
  // Persistent review comments: drop malformed entries + cap; only retain a
  // non-empty list so a slide that never authored a comment stays byte-clean
  // (no `comments` key). A comment's `blockId` anchor is NOT pruned against
  // the slide's blocks here (a comment may intentionally outlive a deleted
  // object as a "what happened to X" note); the render edge degrades a
  // dangling anchor to a slide-level comment.
  const normalizedComments =
    parsedSuccess && Array.isArray(rawComments)
      ? (JSON.parse(JSON.stringify(rawComments)) as SlideComment[])
      : normalizeComments(rawComments);
  // An imported slide stored before the importer listed its objects at
  // their `p:spTree` places (a SmartArt diagram after every other object,
  // while it painted at its frame's place) is read in paint order, so the
  // Selection Pane and Arrange see the slide the canvas shows; the next save
  
  // A slide that records reorders of its imported objects already lists
  // them as they paint.
  const listedAsPainted =
    slideExtras.nativeShapeOrderTarget === undefined &&
    typeof asRecord(slideExtras.source)?.slideId === "string"
      ? presentationListInPaintOrder(blocks)
      : null;
  const listedBlocks = listedAsPainted
    ? listedAsPainted.flatMap((at) => {
        const block = blocks[at];
        return block ? [block] : [];
      })
    : blocks;
  return {
    ...slideExtras,
    ...(typeof notes === "string" ? { notes } : {}),
    ...(normalizedComments.length > 0
      ? { comments: normalizedComments }
      : {}),
    ...(rawCommentAuthority !== undefined
      ? { commentAuthority: rawCommentAuthority }
      : {}),
    ...(normalizedTransition !== DEFAULT_SLIDE_TRANSITION
      ? { transition: normalizedTransition }
      : {}),
    ...(normalizedTransitionDurationMs !== undefined
      ? { transitionDurationMs: normalizedTransitionDurationMs }
      : {}),
    ...(normalizedTransitionDirection !== undefined
      ? { transitionDirection: normalizedTransitionDirection }
      : {}),
    ...(normalizedTransitionAdvanceOnClick === false
      ? { transitionAdvanceOnClick: false }
      : {}),
    ...(normalizedTransitionAdvanceAfterMs !== undefined
      ? { transitionAdvanceAfterMs: normalizedTransitionAdvanceAfterMs }
      : {}),
    ...(normalizedLayout && (normalizedLayout !== DEFAULT_LAYOUT_ID || keepStatedBlankLayout)
      ? { layout: normalizedLayout }
      : {}),
    ...(normalizedNativeLayoutId
      ? { layoutId: normalizedNativeLayoutId }
      : {}),
    ...(normalizedMasterId ? { masterId: normalizedMasterId } : {}),
    ...(normalizedSequence.length > 0
      ? { animationSequence: normalizedSequence, animationTiming: POWERPOINT_ANIMATION_TIMING }
      : {}),
    // Only an explicit `true` hides a slide; a visible slide persists no
    // `hidden` key (byte-clean). See `normalizeHiddenFlag`.
    ...(normalizeHiddenFlag(rawHidden) ? { hidden: true } : {}),
    id: String(id || `slide-${index + 1}`),
    title: String(title ?? ""),
    // An empty Blank slide stays empty; the editor draws it as an empty
    // canvas (`slideUsesCanvas`). Any other empty slide gets its body.
    blocks: blocksLeftOnSlide({ layout }, listedBlocks, () => randomUUID()),
  };
}

/** The deck's slides read as a whole: its content parsed by its schema at once. */
function readSlidesWhole(
  storedContent: Record<string, unknown> | undefined,
  issues?: PresentationReadIssue[],
): SlideModel[] {
  // A picture every slide inherits is stored once (`presentation-page-media`);
  // the editor works on its bytes and `buildContent` stores it once again.
  const content = expandPresentationMedia(storedContent);
  const parsed = presentationPageContentSchema.safeParse(content);
  const rawSlides = parsed.success
    ? parsed.data.slides
    : Array.isArray(content?.slides)
      ? content.slides
      : [];
  return rawSlides
    .slice(0, MAX_PRESENTATION_SLIDES)
    .map((slide, index) => normalizeStoredSlide(slide, index, parsed.success, issues));
}

// ── Reading a deck slide by slide ─────────────────────────────────────────
//
// A deck's room hands the window its whole content at every collaborator's
// key, and reading it whole parsed every slide by the deck's schema each time:


// every slide it did not change as the same object, so a slide is read once
// per stored object: expanded with the deck's media and receipts, parsed by
// the slide schema, normalized. The deck's own fields are read once per set of
// stored values, and the deck-wide rules run over the whole deck so read each
// time. The result is the whole read's, slide for slide
// (`tests/unit/presentation-read-slide-by-slide.test.ts`): the deck is valid
// exactly when its own fields, its slide count, every slide and the deck-wide
// rules are; a deck that is not is read from its raw slides, as before.

interface StoredSlideRead {
  readonly media: unknown;
  readonly receipts: unknown;
  /** The slide with the deck's media and receipts put in (`expandPresentationMedia`). */
  readonly expanded: unknown;
  readonly parse: { success: true; data: unknown } | { success: false };
}

interface StoredSlideNormalized {
  readonly read: StoredSlideRead;
  readonly parsedSuccess: boolean;
  /** What the slide's identity comes from when it has none of its own. */
  readonly position: number | null;
  readonly slide: SlideModel;
  readonly issues: readonly PresentationReadIssue[];
}

const storedSlideReads = new WeakMap<object, StoredSlideRead>();
const storedSlideNormals = new WeakMap<object, StoredSlideNormalized>();
let lastDeckFieldsRead: { keys: unknown[]; parse: { success: true; data: Record<string, unknown> } | { success: false } } | null = null;

const deckFieldsSchema = presentationPageContentObjectSchema.omit({ slides: true });
const deckRulesSchema = z.any().superRefine(
  refinePresentationPageContent as (value: unknown, ctx: z.RefinementCtx) => void,
);

function knownStoredSlideRead(stored: Record<string, unknown>, slide: unknown): StoredSlideRead | null {
  const known = slide !== null && typeof slide === "object" ? storedSlideReads.get(slide) : undefined;
  return known && known.media === stored[PRESENTATION_MEDIA_KEY] && known.receipts === stored[PRESENTATION_DECK_RECEIPTS_KEY]
    ? known : null;
}

function readStoredSlide(stored: Record<string, unknown>, slide: unknown, expandedWithDeck?: unknown): StoredSlideRead {
  const media = stored[PRESENTATION_MEDIA_KEY];
  const receipts = stored[PRESENTATION_DECK_RECEIPTS_KEY];
  const key = slide !== null && typeof slide === "object" ? slide : null;
  const known = knownStoredSlideRead(stored, slide);
  if (known) return known;
  let expanded = expandedWithDeck;
  if (expanded === undefined) {
    const one: Record<string, unknown> = { slides: [slide] };
    if (media !== undefined) one[PRESENTATION_MEDIA_KEY] = media;
    if (receipts !== undefined) one[PRESENTATION_DECK_RECEIPTS_KEY] = receipts;
    expanded = (expandPresentationMedia(one).slides as unknown[])[0];
  }
  const parsed = presentationSlideSchema.safeParse(expanded);
  const read: StoredSlideRead = {
    media, receipts, expanded,
    parse: parsed.success ? { success: true, data: parsed.data } : { success: false },
  };
  if (key) storedSlideReads.set(key, read);
  return read;
}

function readDeckFields(stored: Record<string, unknown>) {
  const keys = Object.keys(stored).filter((key) => key !== "slides").sort().flatMap((key) => [key, stored[key]]);
  const last = lastDeckFieldsRead;
  if (last && last.keys.length === keys.length && last.keys.every((value, at) => value === keys[at])) return last.parse;
  const fields: Record<string, unknown> = {};
  for (const key of Object.keys(stored)) if (key !== "slides") fields[key] = stored[key];
  const expanded = expandPresentationMedia({ ...fields, slides: [] }) as Record<string, unknown>;
  const { slides: _slides, ...withoutSlides } = expanded;
  const parsed = deckFieldsSchema.safeParse(withoutSlides);
  const parse = parsed.success
    ? { success: true as const, data: parsed.data as Record<string, unknown> }
    : { success: false as const };
  lastDeckFieldsRead = { keys, parse };
  return parse;
}

function readSlidesOneByOne(
  storedContent: Record<string, unknown> | undefined,
  issues?: PresentationReadIssue[],
): SlideModel[] | null {
  const stored = asRecord(storedContent);
  if (!stored || !Array.isArray(stored.slides)) return null;
  const storedSlides = stored.slides as unknown[];
  // A deck opened (most slides not read yet) is expanded once, whole: one
  // slide at a time it is a pass over the deck's receipts per slide.
  const unread = storedSlides.filter((slide) => !knownStoredSlideRead(stored, slide)).length;
  const expandedWhole = unread > 8
    ? ((expandPresentationMedia(stored) as Record<string, unknown>).slides as unknown[] | undefined) : undefined;
  const reads = storedSlides.map((slide, index) => readStoredSlide(stored, slide, expandedWhole?.[index]));
  const fields = readDeckFields(stored);
  let success = fields.success && storedSlides.length >= 1 && storedSlides.length <= MAX_PRESENTATION_SLIDES &&
    reads.every((read) => read.parse.success);
  if (success && fields.success) {
    const slides = reads.map((read) => (read.parse as { data: unknown }).data);
    success = deckRulesSchema.safeParse({ ...fields.data, slides }).success;
  }
  return storedSlides.slice(0, MAX_PRESENTATION_SLIDES).map((slide, index) => {
    const read = reads[index];
    const record = asRecord(read.expanded);
    const position = typeof record?.id === "string" && record.id ? null : index;
    const key = slide !== null && typeof slide === "object" ? slide : null;
    const known = key ? storedSlideNormals.get(key) : undefined;
    if (known && known.read === read && known.parsedSuccess === success && known.position === position) {
      issues?.push(...known.issues);
      return known.slide;
    }
    const own: PresentationReadIssue[] = [];
    const normalized = normalizeStoredSlide(
      success ? (read.parse as { data: unknown }).data : read.expanded, index, success, own,
    );
    issues?.push(...own);
    if (key) storedSlideNormals.set(key, { read, parsedSuccess: success, position, slide: normalized, issues: own });
    return normalized;
  });
}

/** @internal the whole-deck reading, for the slide-by-slide read's comparison. */
export const normalizeSlidesWholeForTests = readSlidesWhole;

export function normalizeSlides(
  storedContent: Record<string, unknown> | undefined,
  /** Blocks read without a setting they contradict (see `PresentationReadIssue`). */
  issues?: PresentationReadIssue[],
): SlideModel[] {
  const slides = readSlidesOneByOne(storedContent, issues) ?? readSlidesWhole(storedContent, issues);
  if (slides.length === 0) return [defaultSlide()];
  // The backend seeds a brand-new presentation page as ONE slide holding ONE
  // empty legacy body-text block (no geometry, no placeholder role). That shape
  // renders in the document-like stacked editor — and the first canvas insert
  // then flips the whole layout mid-edit, which read as "my slide got replaced"
  
  // seed — never typed into, never styled, still one slide — upgrades to the
  // same Title Slide layout `defaultSlide()` starts with, so a fresh page is
  // canvas-native from the first frame. Any authored content at all (text, a
  // second block or slide, styling, notes, a transition) leaves the deck alone:
  // the check is key-exhaustive, not key-specific, so a NEW feature's key can
  // never be mistaken for pristine.
  if (slides.length === 1) {
    const only = slides[0] as unknown as Record<string, unknown>;
    const pristineSlide =
      only.title === "" &&
      Object.keys(only).every((key) =>
        ["id", "title", "blocks"].includes(key),
      ) &&
      Array.isArray(only.blocks) &&
      only.blocks.length === 1;
    if (pristineSlide) {
      const block = asRecord((only.blocks as unknown[])[0]) ?? {};
      const pristineBlock =
        block.type === "text" &&
        String(block.text ?? "").trim() === "" &&
        Object.keys(block).every((key) => ["id", "type", "text"].includes(key));
      if (pristineBlock) return [defaultSlide()];
    }
  }
  return slides;
}

/**
 * Read the deck-level theme out of a persisted page's content, normalizing it
 * through the shared resolver so only valid fields survive. Returns `undefined`
 * for a themeless deck (the common case) so the editor + present mode fall back
 * to the historical built-in defaults.
 */
export function parseDeckTheme(
  content: Record<string, unknown> | undefined,
): DeckTheme | undefined {
  return materializeCatalogThemeIdentity(normalizeTheme(content?.theme));
}

/**
 * Whether a page's content is exactly the server's seed for a new deck
 * (`default_content_for_page_type`): one untitled Title Slide holding the
 * two empty title / subtitle placeholders, and nothing else at all -- no
 * theme, no masters, no other deck field. Key-exhaustive, so any authored
 * content or new field leaves a deck alone.
 */
export function isUntouchedNewDeckSeed(content: unknown): boolean {
  const record = asRecord(content);
  if (
    !record ||
    record.type !== "mrnote.presentation.v1" ||
    !Object.keys(record).every((key) => key === "type" || key === "slides") ||
    !Array.isArray(record.slides) ||
    record.slides.length !== 1
  ) {
    return false;
  }
  const slide = asRecord(record.slides[0]);
  if (
    !slide ||
    slide.title !== "" ||
    slide.layout !== "title" ||
    !Object.keys(slide).every((key) =>
      ["id", "title", "layout", "blocks"].includes(key),
    ) ||
    !Array.isArray(slide.blocks) ||
    slide.blocks.length !== 2
  ) {
    return false;
  }
  const placeholders = ["title", "subtitle"];
  return slide.blocks.every((raw, index) => {
    const block = asRecord(raw);
    return (
      !!block &&
      block.type === "text" &&
      block.text === "" &&
      block.placeholder === placeholders[index] &&
      Object.keys(block).every((key) =>
        [
          "id",
          "type",
          "kind",
          "text",
          "placeholder",
          "role",
          "left",
          "top",
          "width",
          "height",
        ].includes(key),
      )
    );
  });
}

/**
 * Read the deck-level masters out of a persisted page's content, normalizing
 * each through the shared resolver so only valid fields survive (and dangling /
 * duplicate ids are dropped). Returns `[]` for a deck with no masters (the
 * common case) so the editor + present mode fall back to the theme-only chain.
 */
export function parseDeckMasters(
  content: Record<string, unknown> | undefined,
): SlideMaster[] {
  // Sanitize each master placeholder block with the SAME normalizer slide blocks
  // run through, so a forged / malformed persisted master block (e.g. a table
  // with non-array rows) is repaired instead of crashing the renderer on load.
  return normalizeMasters(
    expandPresentationMedia(content)?.masters,
    normalizePresentationBlock,
  );
}

/**
 * Read + normalize the deck-level SECTIONS out of a persisted page's content
 * against the deck's slide ids. A malformed value (bad ids, a dangling anchor,
 * duplicate ids / anchors, an unsorted list) is dropped / repaired here so deck
 * load can never crash; `[]` for a deck with no sections (the common case).
 */
export function parseDeckSections(
  content: Record<string, unknown> | undefined,
  slides: SlideModel[],
): SlideSection[] {
  
  return storedDeckSections(
    content?.sections,
    slides as unknown as (SlideModel & { section?: unknown })[],
  );
}

export function parseDeckCustomSlideShows(
  content: Record<string, unknown> | undefined,
  slides: readonly SlideModel[],
): PresentationCustomShows {
  return normalizePresentationCustomShows(
    content?.customSlideShows,
    slides.map((slide) => slide.id),
  );
}

/**
 * A block read without a setting it contradicted (`normalizeSlides`) goes
 * back with that setting as stored, unless the user has set a new one: what
 * the editor could not honour is never lost by opening and saving the deck.
 */
function withReadPreservedSettings<T extends PresentationBlock>(block: T): T {
  const record = block as unknown as Record<string, unknown>;
  const preserved = asRecord(record[PRESENTATION_READ_PRESERVED_KEY]);
  if (!preserved && !(PRESENTATION_READ_PRESERVED_KEY in record)) return block;
  const { [PRESENTATION_READ_PRESERVED_KEY]: _preserved, ...rest } = record;
  void _preserved;
  for (const [key, value] of Object.entries(preserved ?? {})) {
    if (!(key in rest)) rest[key] = value;
  }
  return rest as unknown as T;
}









const savedThemes = new WeakMap<object, ReturnType<typeof normalizeTheme>>();
const savedMasters = new WeakMap<object, ReturnType<typeof normalizeMasters>>();

function savedTheme(theme: DeckTheme | undefined): ReturnType<typeof normalizeTheme> {
  if (!theme || typeof theme !== "object") return normalizeTheme(theme);
  if (savedThemes.has(theme)) return savedThemes.get(theme);
  const normalized = normalizeTheme(theme);
  savedThemes.set(theme, normalized);
  return normalized;
}

function savedMastersOf(
  masters: readonly SlideMaster[] | undefined,
  forSave: Parameters<typeof normalizeMasters>[1],
): ReturnType<typeof normalizeMasters> {
  if (!masters || typeof masters !== "object") return normalizeMasters(masters, forSave);
  const cached = savedMasters.get(masters);
  if (cached) return cached;
  const normalized = normalizeMasters(masters, forSave);
  savedMasters.set(masters, normalized);
  return normalized;
}

/** A master's or layout's chart preview goes back as its stored reading. */
function masterBlockForSave(block: unknown): unknown {
  const record = block && typeof block === "object" ? (block as { type?: unknown }) : null;
  return record?.type === "chart"
    ? presentationChartBlockForSave(record as unknown as PresentationChartBlock)
    : block;
}

/** The slide with `animationTiming` stated exactly when it has a sequence. */
export function slideWithAnimationTiming<T extends { animationSequence?: unknown; animationTiming?: unknown }>(slide: T): T {
  const timed = Array.isArray(slide.animationSequence) && slide.animationSequence.length > 0;
  if (timed && slide.animationTiming !== POWERPOINT_ANIMATION_TIMING) {
    return { ...slide, animationTiming: POWERPOINT_ANIMATION_TIMING };
  }
  if (!timed && slide.animationTiming !== undefined) {
    const { animationTiming: _dropped, ...rest } = slide;
    void _dropped;
    return rest as T;
  }
  return slide;
}

export function buildContent(
  slides: SlideModel[],
  theme?: DeckTheme,
  masters?: readonly SlideMaster[],
  sections?: readonly SlideSection[],
  headerFooter?: DeckHeaderFooter,
  printSettings?: PresentationPrintSettings,
  sensitivity?: PresentationSensitivity,
  slideShowSettings?: PresentationSlideShowSettings,
  customSlideShows?: PresentationCustomShows,
  fontFaces?: readonly OfficeFontFace[],
  /** View > Grid Settings (`lib/presentation-grid`), the deck's own. */
  viewSettings?: PresentationViewSettings,
): Record<string, unknown> {
  // Normalize the deck theme at the serialize edge so only valid fields persist;
  // an empty / absent theme is omitted entirely so default decks stay byte-clean
  // and existing pages keep validating exactly as before.
  const normalizedFontFaces = presentationFontFacesForPersist(fontFaces);
  const normalizedTheme = savedTheme(theme);
  // Same for masters: re-normalize + drop an empty array so a deck that never
  // defined a master stays byte-clean (no `masters` key persisted).
  // A master's or layout's chart preview goes back as its stored reading, as
  // a slide's does below (`presentationChartBlockForSave`): the load composed
  // `chartV2` onto it for the canvas, and the server refuses a block whose
  
  // stored as "doughnut" refused every save of the deck).
  const normalizedMasters = savedMastersOf(masters, masterBlockForSave as Parameters<typeof normalizeMasters>[1]);
  // Same for sections: re-normalize against the CURRENT slide ids (so a now-
  // dangling anchor cannot persist) + drop an empty list so a deck that never
  // defined a section stays byte-for-byte unchanged (no `sections` key).
  // The sections as the deck stores them: per-slide names for a deck whose
  // slides name them (one read from a PPTX: the native save and the PPTX
  
  // with the names too when two neighbouring sections share one.
  const storedSections = sectionsInStoredForm(
    slides as unknown as (SlideModel & { section?: unknown })[],
    sections ?? [],
  );
  const sectionedSlides = storedSections.slides as unknown as SlideModel[];
  const normalizedSections = storedSections.sections;
  // Same for the slide number + footer: re-normalize (strip control bytes, drop a
  // `false` flag / empty footer) so a deck that never configured a header/footer
  // stays byte-for-byte unchanged (no `headerFooter` key, no `slideNumber: false`,
  // no empty `footerText`).
  const normalizedHeaderFooter = normalizeHeaderFooter(headerFooter);
  // Same for the print / handout settings: re-normalize (drop fields equal to
  // the export default — 2 slides/page, 100 % scale) so a deck that never
  // configured print settings stays byte-for-byte unchanged (no `printSettings`
  // key persisted) and the export keeps using handout-2.
  const normalizedPrintSettings =
    normalizePresentationPrintSettings(printSettings);
  // Set Up Show's From/To are slide numbers: a range past the deck's last
  // slide (slides deleted since) ends at the last slide.
  const normalizedSlideShowSettings = presentationSlideShowSettingsWithinSlides(
    normalizePresentationSlideShowSettings(slideShowSettings),
    slides.length,
  );
  const normalizedCustomSlideShows = presentationCustomShowsForPersist(
    customSlideShows,
    slides.map((slide) => slide.id),
  );
  // A chart awaiting its V2 receipt goes back as its stored reading: the
  // attrs the load composed for it are the page's view, not the file's.
  // Words and their formatting saved together: a body whose runs or
  
  const persistedSlides = sectionedSlides.map((slide) => {
    const blocks = slide.blocks.map((block) =>
      block.type === "chart"
        ? presentationChartBlockForSave(block)
        : repairBlockTextLayers(withReadPreservedSettings(block)),
    );
    // A sequence is saved with the timing it is stated in (PowerPoint's): a
    // slide without the mark is read as one saved before it.
    const timed = slideWithAnimationTiming(slide);
    return blocks.some((block, index) => block !== slide.blocks[index])
      ? { ...timed, blocks }
      : timed;
  });
  return compactPresentationMedia({
    type: PRESENTATION_CONTENT_TYPE,
    slides: persistedSlides,
    ...(normalizedFontFaces ? { fontFaces: normalizedFontFaces } : {}),
    ...(normalizedTheme ? { theme: normalizedTheme } : {}),
    ...(normalizedMasters.length > 0 ? { masters: normalizedMasters } : {}),
    ...(normalizedSections ? { sections: normalizedSections } : {}),
    ...(normalizedHeaderFooter ? { headerFooter: normalizedHeaderFooter } : {}),
    ...(normalizedPrintSettings
      ? { printSettings: normalizedPrintSettings }
      : {}),
    ...(normalizedSlideShowSettings
      ? { slideShowSettings: normalizedSlideShowSettings }
      : {}),
    ...(normalizedCustomSlideShows
      ? { customSlideShows: normalizedCustomSlideShows }
      : {}),
    ...(normalizeSensitivity(sensitivity)
      ? { sensitivity: normalizeSensitivity(sensitivity) }
      : {}),
    ...(normalizePresentationViewSettings(viewSettings)
      ? { viewSettings: normalizePresentationViewSettings(viewSettings) }
      : {}),
  });
}

/**
 * The body a deck is stored as once this editor has opened it: the page
 * load's own normalisation (`normalizeSlides` and the `parseDeck*` readers,
 * as `handleLoadedPage` and the guest seed apply them) serialized by the same
 * `buildContent` every autosave writes. A caller that seeds a page with a
 * deck it did not author in this editor — the dev dialog galleries' scratch
 * pages — writes exactly what the editor would save, so the stored-page
 * validators judge an editor-shaped body rather than a hand-written one.
 */
export function presentationPageContentFromDeck(
  content: Record<string, unknown>,
): Record<string, unknown> {
  const slides = normalizeSlides(content);
  return buildContent(
    slides,
    parseDeckTheme(content),
    parseDeckMasters(content),
    parseDeckSections(content, slides),
    normalizeHeaderFooter(content.headerFooter),
    normalizePresentationPrintSettings(content.printSettings),
    normalizeSensitivity(content.sensitivity),
    normalizePresentationSlideShowSettings(content.slideShowSettings),
    parseDeckCustomSlideShows(content, slides),
    normalizePresentationFontFaces(content.fontFaces),
    normalizePresentationViewSettings(content.viewSettings),
  );
}
