/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Predefined slide layouts (PowerPoint / WPS master-slide parity subset).
 *
 * A *layout* describes the default set of placeholder blocks — and their
 * geometry — for a new slide. The supported set mirrors the small, familiar
 * PowerPoint family: Title, Title + Content, Two Content, Section Header, and
 * Blank.
 *
 * Pure, side-effect-free helpers shared by the editor (and unit tests) so the
 * allowed layout ids, their placeholder templates, and the
 * fractional → absolute geometry resolution live in exactly one place. Keep this
 * module dependency-free (no React / no DOM) so it can be unit tested in
 * isolation.
 *
 * Geometry is authored as *fractions* of the slide canvas (0..1) and resolved to
 * EMUs (English Metric Units) against the slide's bounds, matching the existing
 * presentation geometry (`left`/`top`/`width`/`height` on every block). This
 * keeps a layout correct regardless of the deck's actual slide dimensions.
 */

import { normalizeOfficeShapeShadow } from "./office-shape-paint";
import { nativeSlideSizeLayoutProjection } from "./presentation-slide-size";
import { ptToPx } from "./office-units";
import {
  dropNativeParagraphReceipts,
  isFlatEquivalentParagraphs,
  type SlideParagraph,
} from "./presentation-paragraphs";
import {
  PRESENTATION_OFFICE_CATALOG_SLIDE_HEIGHT,
  PRESENTATION_OFFICE_CATALOG_SLIDE_WIDTH,
  isOfficeThemeVariantId,
  presentationOfficeFrameTextColor,
  presentationOfficeLayoutFrames,
  presentationOfficeThemeVariant,
  type PresentationOfficePlaceholderFrame,
  type PresentationOfficeThemeVariant,
  type PresentationOfficeVariantId,
} from "./presentation-office-themes";
import {
  officeSlotOwnAttributes,
  officeSlotOwnsGeometry,
  type OfficeRunAttribute,
} from "./presentation-office-ladder-catalog";
import {
  flippedColorSlot,
  officeMasterColorMapFlipped,
} from "./presentation-office-background-styles";
import type { DeckTheme } from "./presentation-theme";
import { type TextRun } from "./presentation-runs";
import { normalizeRuns } from "./presentation-point-runs";
import {
  OFFICE_INSERTED_MASTER_THEME_ID,
  OFFICE_INSERTED_MASTER_VARIANT_ID,
  isOfficeCustomLayoutId,
  normalizeDeckOfficeMaster,
  parseOfficeMasterLayoutRef,
  type DeckOfficeMaster,
  type OfficePlaceholderKind,
  type OfficeMasterTextStyles,
  type OfficeMasterTextValues,
} from "./presentation-office-master-model";
export * from "./presentation-layout-id";
import { DEFAULT_LAYOUT_ID, LAYOUT_IDS, isLayoutId, normalizeLayoutId } from "./presentation-layout-id";
import type { LayoutId } from "./presentation-layout-id";

/** A 4 : 3-ish default slide canvas, mirroring PresentationWindow defaults. */
const DEFAULT_SLIDE_WIDTH_EMU = 9_144_000;
const DEFAULT_SLIDE_HEIGHT_EMU = 5_143_500;

/**
 * A layout of an inserted slide master (PPT-B030): `<master key>--<layout id>`,
 * e.g. `master-b--title-content` (presentation-office-master-model.ts:
 * officeMasterLayoutRef). Master one's layouts carry the bare id.
 */
export type OfficeMasterLayoutRef = `master-${string}--${LayoutId}`;

/**
 * The role a layout placeholder plays. Used purely so callers / tests can
 * reason about a layout's structure; it does not affect persisted block shape
 * (every placeholder is a positioned text block).
 */
export type PlaceholderRole =
  | "title"
  | "subtitle"
  | "body"
  | "content"
  | "picture";

/**
 * A single placeholder template within a layout, authored in fractional
 * coordinates (0..1 of the slide canvas). `role` is descriptive metadata;
 * `placeholderKey` selects the prompt text shown in an empty placeholder.
 */
export interface LayoutPlaceholder {
  role: PlaceholderRole;
  /** i18n key suffix under `presentation.layoutPlaceholder.*` for empty prompt. */
  placeholderKey: string;
  /** CJK vertical text (PPT 竖排): the block ships with vertical-rl style. */
  vertical?: boolean;
  /** Fractional geometry within the slide (each value in [0, 1]). */
  left: number;
  top: number;
  width: number;
  height: number;
  /**
   * The text the layout gives the placeholder (an Office theme layout's
   * level-1 facts), written on the block's `style` when it is made. The
   * built-in layouts state none.
   */
  text?: LayoutPlaceholderText;
}

/**
 * Block-style values a layout states for a placeholder, in the block's own
 * vocabulary: `fontSize` in the editor's design px (the layout's point size
 * x 4/3), the rest exactly as `style.align` / `style.verticalAlign` /
 * `style.bold` / `style.caps` / `style.color` are persisted. A value the
 * layout leaves at its default is absent.
 */
export interface LayoutPlaceholderText {
  fontSize?: number;
  align?: "center" | "right" | "justify" | "distribute";
  verticalAlign?: "middle" | "bottom";
  bold?: boolean;
  italic?: boolean;
  caps?: "all" | "small";
  color?: string;
}

/** The block-style keys a layout can state ({@link LayoutPlaceholderText}). */
export const LAYOUT_TEXT_STYLE_KEYS = [
  "fontSize",
  "align",
  "verticalAlign",
  "bold",
  "italic",
  "caps",
  "color",
] as const;

/**
 * Where a layout's placeholder frames come from: MRAI's own
 * {@link LAYOUT_DEFINITIONS} (`builtin`), or the layouts of a PowerPoint
 * Office theme variant (Design > Themes), whose frames the catalog records
 * (`presentation-office-themes.json`). A deck on an Office theme places new
 * placeholders where PowerPoint would.
 */
export type LayoutFrameSource =
  | { readonly kind: "builtin" }
  | {
      readonly kind: "office";
      readonly themeId: string;
      readonly variantId: PresentationOfficeVariantId;
      /**
       * What Slide Master view changed in the master's text styles
       * (`theme.officeMaster.textStyles`, PPT-B030): its level-one values
       * reach every placeholder whose own list style does not state them.
       */
      readonly masterText?: OfficeMasterTextStyles;
      /**
       * A Background Styles pick put the master on the other colour map
       * (PPT-A073): text the layout draws in tx1 / tx2 takes the other slot
       * of its pair (dk1 <-> lt1, dk2 <-> lt2).
       */
      readonly flipColorMap?: true;
      /** Placeholder frames moved in Slide Master view (PPT-B030). */
      readonly masterFrames?: DeckOfficeMaster["masterFrames"];
      readonly layoutFrames?: DeckOfficeMaster["layoutFrames"];
      /** Placeholders inserted into / deleted from a layout in Slide Master view (PPT-B030). */
      readonly addedPlaceholders?: DeckOfficeMaster["addedPlaceholders"];
      readonly layoutSlots?: DeckOfficeMaster["layoutSlots"];
      /** Insert Layout's custom layouts on the master (PPT-B030). */
      readonly customLayouts?: DeckOfficeMaster["customLayouts"];
      /**
       * Insert Slide Master's masters (PPT-B030), by key: a layout ref
       * `<key>--<layout>` takes its frames from that master's source
       * ({@link resolveOfficeLayoutRef}).
       */
      readonly additionalMasters?: Readonly<Record<string, LayoutFrameSource>>;
    };

export const BUILTIN_LAYOUT_FRAMES: LayoutFrameSource = Object.freeze({
  kind: "builtin",
}) as LayoutFrameSource;

/**
 * The frame source a deck theme implies: an Office selection names its theme
 * and variant; every other theme (MRAI presets, custom, native, none) uses
 * the built-in layouts. Reads only the identity, so any theme record works.
 */
export function layoutFrameSourceForTheme(
  theme:
    | {
        identity?: {
          selectionKind?: unknown;
          themeId?: unknown;
          variantId?: unknown;
        } | null;
        officeMaster?: unknown;
      }
    | null
    | undefined,
): LayoutFrameSource {
  const identity = theme?.identity;
  if (
    identity?.selectionKind === "office" &&
    typeof identity.themeId === "string" &&
    isOfficeThemeVariantId(identity.variantId) &&
    presentationOfficeThemeVariant(identity.themeId, identity.variantId)
  ) {
    const officeMaster = normalizeDeckOfficeMaster(theme?.officeMaster);
    const flipColorMap = officeMasterColorMapFlipped(theme as DeckTheme);
    const source = officeFrameSource(identity.themeId, identity.variantId, officeMaster, flipColorMap);
    const additional = officeMaster?.additionalMasters ?? [];
    if (additional.length === 0) return source;
    // Insert Slide Master's masters are Office's default theme, each with
    // its own Slide Master view edits (PPT-B030).
    return {
      ...source,
      additionalMasters: Object.fromEntries(
        additional.map((entry) => [
          entry.key,
          officeFrameSource(OFFICE_INSERTED_MASTER_THEME_ID, OFFICE_INSERTED_MASTER_VARIANT_ID, entry, false),
        ]),
      ),
    };
  }
  return BUILTIN_LAYOUT_FRAMES;
}

function officeFrameSource(
  themeId: string,
  variantId: PresentationOfficeVariantId,
  officeMaster: Omit<DeckOfficeMaster, "additionalMasters"> | undefined,
  flipColorMap: boolean,
): Extract<LayoutFrameSource, { kind: "office" }> {
  const masterText = officeMaster?.textStyles;
  return {
    kind: "office",
    themeId,
    variantId,
    ...(masterText ? { masterText } : {}),
    ...(flipColorMap ? { flipColorMap: true as const } : {}),
    ...(officeMaster?.masterFrames ? { masterFrames: officeMaster.masterFrames } : {}),
    ...(officeMaster?.layoutFrames ? { layoutFrames: officeMaster.layoutFrames } : {}),
    ...(officeMaster?.addedPlaceholders ? { addedPlaceholders: officeMaster.addedPlaceholders } : {}),
    ...(officeMaster?.layoutSlots ? { layoutSlots: officeMaster.layoutSlots } : {}),
    ...(officeMaster?.customLayouts ? { customLayouts: officeMaster.customLayouts } : {}),
  };
}

/**
 * The source and the master's own layout id a layout ref addresses: a layout
 * of an inserted master (`<key>--<layout>`, PPT-B030) is that master's
 * layout under its source; a master the deck no longer has falls back to the
 * same layout of master one; any other ref is the source's own.
 */
export function resolveOfficeLayoutRef(
  source: LayoutFrameSource,
  value: unknown,
): { source: LayoutFrameSource; layoutId: unknown } {
  const ref = parseOfficeMasterLayoutRef(value);
  if (!ref) return { source, layoutId: value };
  const inserted = source.kind === "office" ? source.additionalMasters?.[ref.masterKey] : undefined;
  return { source: inserted ?? source, layoutId: ref.layoutId };
}

/**
 * The catalog layout whose frames a layout starts from: a custom layout
 * Insert Layout made is PowerPoint's Custom Layout -- a title and the
 * footers, which is the Title Only layout's -- and every other layout is its
 * own (PPT-B030).
 */
export function officeCatalogLayoutId(source: LayoutFrameSource, layoutId: string): string {
  const routed = resolveOfficeLayoutRef(source, layoutId);
  if (routed.source !== source || routed.layoutId !== layoutId) {
    const inner = officeCatalogLayoutId(routed.source, routed.layoutId as string);
    return inner === routed.layoutId ? layoutId : inner;
  }
  return source.kind === "office" &&
    isOfficeCustomLayoutId(layoutId) &&
    (source.customLayouts ?? []).includes(layoutId)
    ? "title-only"
    : layoutId;
}

export function layoutFrameSourcesEqual(
  left: LayoutFrameSource,
  right: LayoutFrameSource,
): boolean {
  if (left.kind !== right.kind) return false;
  if (left.kind === "builtin" || right.kind === "builtin") return true;
  return (
    left.themeId === right.themeId &&
    left.variantId === right.variantId &&
    left.flipColorMap === right.flipColorMap &&
    JSON.stringify(left.masterFrames ?? null) === JSON.stringify(right.masterFrames ?? null) &&
    JSON.stringify(left.layoutFrames ?? null) === JSON.stringify(right.layoutFrames ?? null) &&
    JSON.stringify(left.addedPlaceholders ?? null) === JSON.stringify(right.addedPlaceholders ?? null) &&
    JSON.stringify(left.layoutSlots ?? null) === JSON.stringify(right.layoutSlots ?? null) &&
    JSON.stringify(left.customLayouts ?? null) === JSON.stringify(right.customLayouts ?? null) &&
    JSON.stringify(left.masterText ?? null) === JSON.stringify(right.masterText ?? null) &&
    JSON.stringify(left.additionalMasters ?? null) === JSON.stringify(right.additionalMasters ?? null)
  );
}

export interface LayoutDefinition {
  id: LayoutId;
  /** i18n key suffix under `presentation.layout.*` for the display name. */
  nameKey: string;
  placeholders: LayoutPlaceholder[];
}

/**
 * The layout catalog. Geometry is fractional and intentionally kept within the
 * [0, 1] canvas with comfortable margins so resolved blocks always land fully
 * on-slide. `blank` has no placeholders by design.
 */
export const LAYOUT_DEFINITIONS: readonly LayoutDefinition[] = [
  // PowerPoint's own order (the Office Theme's slideLayout1..11, PPT-A034):
  // every layout gallery and menu lists the layouts in this order.
  {
    id: "title",
    nameKey: "title",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.1,
        top: 0.34,
        width: 0.8,
        height: 0.2,
      },
      {
        role: "subtitle",
        placeholderKey: "subtitle",
        left: 0.1,
        top: 0.56,
        width: 0.8,
        height: 0.12,
      },
    ],
  },
  {
    id: "title-content",
    nameKey: "titleContent",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.06,
        top: 0.06,
        width: 0.88,
        height: 0.16,
      },
      {
        role: "content",
        placeholderKey: "content",
        left: 0.06,
        top: 0.26,
        width: 0.88,
        height: 0.66,
      },
    ],
  },
  {
    id: "section-header",
    nameKey: "sectionHeader",
    placeholders: [
      {
        role: "title",
        placeholderKey: "sectionTitle",
        left: 0.08,
        top: 0.4,
        width: 0.84,
        height: 0.18,
      },
      {
        role: "body",
        placeholderKey: "sectionText",
        left: 0.08,
        top: 0.6,
        width: 0.84,
        height: 0.1,
      },
    ],
  },
  {
    id: "two-content",
    nameKey: "twoContent",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.06,
        top: 0.06,
        width: 0.88,
        height: 0.16,
      },
      {
        role: "content",
        placeholderKey: "content",
        left: 0.06,
        top: 0.26,
        width: 0.43,
        height: 0.66,
      },
      {
        role: "content",
        placeholderKey: "content",
        left: 0.51,
        top: 0.26,
        width: 0.43,
        height: 0.66,
      },
    ],
  },
  {
    // Comparison: a title, then two side-by-side columns, each a small HEADING
    // above a content area (PowerPoint "Comparison").
    id: "comparison",
    nameKey: "comparison",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.06,
        top: 0.06,
        width: 0.88,
        height: 0.16,
      },
      {
        role: "body",
        placeholderKey: "heading",
        left: 0.06,
        top: 0.26,
        width: 0.43,
        height: 0.1,
      },
      {
        role: "content",
        placeholderKey: "content",
        left: 0.06,
        top: 0.38,
        width: 0.43,
        height: 0.54,
      },
      {
        role: "body",
        placeholderKey: "heading",
        left: 0.51,
        top: 0.26,
        width: 0.43,
        height: 0.1,
      },
      {
        role: "content",
        placeholderKey: "content",
        left: 0.51,
        top: 0.38,
        width: 0.43,
        height: 0.54,
      },
    ],
  },
  {
    // Title Only: a single title bar across the top, leaving the body free for
    // hand-placed content (PowerPoint "Title Only").
    id: "title-only",
    nameKey: "titleOnly",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.06,
        top: 0.06,
        width: 0.88,
        height: 0.16,
      },
    ],
  },
  {
    id: "blank",
    nameKey: "blank",
    placeholders: [],
  },
  {
    // Content with Caption: a large content area on the right with a title +
    // caption text stacked on the left (PowerPoint "Content with Caption").
    id: "content-caption",
    nameKey: "contentCaption",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.06,
        top: 0.1,
        width: 0.34,
        height: 0.16,
      },
      {
        role: "body",
        placeholderKey: "caption",
        left: 0.06,
        top: 0.28,
        width: 0.34,
        height: 0.62,
      },
      {
        role: "content",
        placeholderKey: "content",
        left: 0.44,
        top: 0.1,
        width: 0.5,
        height: 0.8,
      },
    ],
  },
  {
    // Picture with Caption: a large PICTURE region on the right with a title +
    // caption stacked on the left (PowerPoint "Picture with Caption").
    id: "picture-caption",
    nameKey: "pictureCaption",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.06,
        top: 0.1,
        width: 0.34,
        height: 0.16,
      },
      {
        role: "body",
        placeholderKey: "caption",
        left: 0.06,
        top: 0.28,
        width: 0.34,
        height: 0.62,
      },
      {
        role: "picture",
        placeholderKey: "picture",
        left: 0.44,
        top: 0.1,
        width: 0.5,
        height: 0.8,
      },
    ],
  },
  {
    // PPT 标题和竖排文字: horizontal title, vertical-rl body below.
    id: "title-vertical-text",
    nameKey: "titleVerticalText",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        left: 0.05,
        top: 0.05,
        width: 0.9,
        height: 0.16,
      },
      {
        role: "content",
        placeholderKey: "content",
        vertical: true,
        left: 0.05,
        top: 0.25,
        width: 0.9,
        height: 0.68,
      },
    ],
  },
  {
    // PPT 竖排标题与文本: vertical title column on the RIGHT, vertical body
    // filling the left — the classic CJK book order.
    id: "vertical-title-text",
    nameKey: "verticalTitleText",
    placeholders: [
      {
        role: "title",
        placeholderKey: "title",
        vertical: true,
        left: 0.78,
        top: 0.05,
        width: 0.17,
        height: 0.88,
      },
      {
        role: "content",
        placeholderKey: "content",
        vertical: true,
        left: 0.05,
        top: 0.05,
        width: 0.7,
        height: 0.88,
      },
    ],
  },
] as const;

const LAYOUT_BY_ID: ReadonlyMap<LayoutId, LayoutDefinition> = new Map(
  LAYOUT_DEFINITIONS.map((layout) => [layout.id, layout]),
);

/**
 * A slide's `layout` as the deck keeps it: one of {@link LAYOUT_IDS}, or a
 * custom layout Slide Master view inserted on an Office theme deck
 * (`custom-<letters>`, PPT-B030), which the type carries as a LayoutId. Every
 * other value normalizes as {@link normalizeLayoutId} does.
 */
export function normalizeSlideLayoutRef(value: unknown): LayoutId {
  if (isOfficeCustomLayoutId(value)) return value as LayoutId;
  // A layout of a master Insert Slide Master added: `<key>--<layout>` (PPT-B030).
  const ref = parseOfficeMasterLayoutRef(value);
  if (ref && (isLayoutId(ref.layoutId) || isOfficeCustomLayoutId(ref.layoutId))) return value as LayoutId;
  return normalizeLayoutId(value);
}

/**
 * The blocks a slide keeps once everything on it is removed. A slide on the
 * Blank layout (one that says so: `layout: "blank"`) stays empty, as
 * PowerPoint's does -- the layout has no placeholder to leave behind. Any
 * other slide keeps the one empty body text block it has always been left
 * with, which a legacy stacked slide needs to type into.
 */
export function blocksLeftOnSlide<B>(
  slide: { readonly layout?: unknown },
  remaining: B[],
  makeId: () => string,
): B[] {
  if (remaining.length > 0 || slide.layout === DEFAULT_LAYOUT_ID) return remaining;
  return [{ id: makeId(), type: "text", text: "" } as unknown as B];
}

/** Immutable browser/server custody for one imported slide-layout cascade. */
export const PPTX_LAYOUT_CASCADE_RECEIPT_VERSION =
  "pptx-layout-cascade-v1" as const;
export const PPTX_PLACEHOLDER_RESET_RECEIPT_VERSION =
  "pptx-placeholder-reset-v1" as const;
export const PPTX_NATIVE_SLIDE_INSERTION_RECEIPT_VERSION =
  "pptx-native-slide-insertion-v1" as const;

export type NativePlaceholderResetOverrideKind =
  | "geometry"
  | "shape-fill"
  | "shape-line"
  | "shape-effect"
  | "body-format"
  | "paragraph-format"
  | "run-format";

export interface NativePlaceholderResetOverrideReceipt {
  kind: NativePlaceholderResetOverrideKind;
  path: string;
  action: "remove" | "clear" | "clear-format";
  sha256: string;
}

export interface NativePlaceholderResetReceipt {
  shapeId: string;
  fingerprint: string;
  overrides: NativePlaceholderResetOverrideReceipt[];
}

export interface NativePlaceholderResetAuthorityReceipt {
  version: typeof PPTX_PLACEHOLDER_RESET_RECEIPT_VERSION;
  disposition: "native" | "preserve-only";
  blocker: string | null;
  placeholders: NativePlaceholderResetReceipt[];
}

export interface NativeLayoutPlaceholderReceipt {
  shapeId: string;
  placeholderType: string;
  placeholderIndex: string;
  nativeIndexPresent: boolean;
  owner: "slide" | "layout";
  localXfrm: boolean;
  localStyle: boolean;
  layoutShapeId?: string | null;
  layoutGeometry: Record<string, number>;
  layoutStyle: Record<string, unknown>;
}

export interface NativeLayoutCatalogEntryReceipt {
  layoutPart: string;
  layoutFingerprint: string;
  masterPart: string;
  masterRelType: string;
  layoutType: string | null;
  layoutName: string | null;
  placeholders: NativeLayoutPlaceholderReceipt[];
}

export interface NativeLayoutCascadeReceipt {
  schemaVersion: typeof PPTX_LAYOUT_CASCADE_RECEIPT_VERSION;
  packageFingerprint: string;
  presentationPart: string;
  slideId: string;
  slidePart: string;
  layoutRelId: string;
  layoutRelType: string;
  layoutPart: string;
  layoutFingerprint: string;
  masterPart: string;
  masterRelType: string;
  masterFingerprint: string;
  themePart: string | null;
  themeRelType: string;
  themeFingerprint: string | null;
  layoutType: string | null;
  layoutName: string | null;
  placeholders: NativeLayoutPlaceholderReceipt[];
  layouts: NativeLayoutCatalogEntryReceipt[];
  resetAuthority?: NativePlaceholderResetAuthorityReceipt;
}

/** Durable same-package authority for one source-less model slide. */
export interface NativeSlideInsertionReceipt {
  schemaVersion: typeof PPTX_NATIVE_SLIDE_INSERTION_RECEIPT_VERSION;
  kind: "new" | "duplicate";
  modelSlideId: string;
  packageFingerprint: string;
  presentationPart: string;
  anchorSlideId: string;
  anchorSlidePart: string;
  layoutPart: string;
  layoutFingerprint: string;
  masterPart: string;
  masterFingerprint: string;
}

interface NativeSlideInsertionHost {
  id: string;
  blocks: Array<{ type?: unknown }>;
  notes?: unknown;
  comments?: unknown;
  commentAuthority?: unknown;
  source?: unknown;
  layoutId?: unknown;
  [key: string]: unknown;
}

/** Exact ST_SlideLayoutType subset represented by the existing Web gallery. */
const NATIVE_LAYOUT_TYPE_TO_ID: Readonly<Record<string, LayoutId>> =
  Object.freeze({
    title: "title",
    obj: "title-content",
    tx: "title-content",
    twoObj: "two-content",
    twoTxTwoObj: "comparison",
    titleOnly: "title-only",
    objAndTx: "content-caption",
    picTx: "picture-caption",
    secHead: "section-header",
    vertTx: "title-vertical-text",
    vertTitleAndTx: "vertical-title-text",
    blank: "blank",
  });

export function nativeLayoutIdFromType(value: unknown): LayoutId | null {
  return typeof value === "string"
    ? (NATIVE_LAYOUT_TYPE_TO_ID[value] ?? null)
    : null;
}

export function nativeLayoutCatalogEntryForLayout(
  receipt: NativeLayoutCascadeReceipt,
  layoutId: LayoutId,
): NativeLayoutCatalogEntryReceipt | null {
  // The file's layout of that built-in type, else the one its placeholders
  // make it (a custom "Title and Content"), the slide's own master's first:
  // PowerPoint's New Slide takes its master's first such layout. A deck with
  // two layouts of a type, or only custom layouts, used to have none, and
  
  const sameMasterFirst = (entries: NativeLayoutCatalogEntryReceipt[]) =>
    entries.find((entry) => entry.masterPart === receipt.masterPart) ?? entries[0] ?? null;
  const typed = receipt.layouts.filter(
    (entry) => nativeLayoutIdFromType(entry.layoutType) === layoutId,
  );
  if (typed.length > 0) return sameMasterFirst(typed);
  return sameMasterFirst(
    receipt.layouts.filter(
      (entry) =>
        !nativeLayoutIdFromType(entry.layoutType) &&
        nativeLayoutModelLayoutId(entry) === layoutId,
    ),
  );
}

/**
 * One of an imported file's OWN layouts, as the New Slide and Layout
 * galleries name it: `native-layout:<layout part>`. PowerPoint's galleries
 * list the layouts the deck's master has (by their names, custom ones
 * included), not Office's built-in eleven; a value naming the exact part
 * also tells two layouts of one type apart.
 */
export const NATIVE_LAYOUT_REF_PREFIX = "native-layout:";

export function nativeLayoutRef(layoutPart: string): string {
  return `${NATIVE_LAYOUT_REF_PREFIX}${layoutPart}`;
}

export function parseNativeLayoutRef(value: unknown): string | null {
  return typeof value === "string" &&
    value.startsWith(NATIVE_LAYOUT_REF_PREFIX) &&
    value.length > NATIVE_LAYOUT_REF_PREFIX.length
    ? value.slice(NATIVE_LAYOUT_REF_PREFIX.length)
    : null;
}

/** The file layout a gallery value names: an exact part, or the one layout of a built-in type. */
export function nativeLayoutCatalogEntryForValue(
  receipt: NativeLayoutCascadeReceipt,
  value: unknown,
): NativeLayoutCatalogEntryReceipt | null {
  const part = parseNativeLayoutRef(value);
  if (part) return receipt.layouts.find((entry) => entry.layoutPart === part) ?? null;
  return nativeLayoutCatalogEntryForLayout(receipt, normalizeLayoutId(value));
}

/**
 * The built-in layout a file layout reads as on the slide model: its own
 * type's, else what its placeholders make it (a custom layout with a title
 * and a body reads as Title and Content, a title alone as Title Only, none
 * as Blank).
 */
export function nativeLayoutModelLayoutId(entry: NativeLayoutCatalogEntryReceipt): LayoutId {
  const typed = nativeLayoutIdFromType(entry.layoutType);
  if (typed) return typed;
  const roles = entry.placeholders
    .filter((placeholder) => !NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES.has(placeholder.placeholderType))
    .map((placeholder) => nativePlaceholderRole(placeholder.placeholderType)?.role);
  if (roles.length === 0) return "blank";
  if (roles.includes("title") && roles.some((role) => role === "body" || role === "content")) {
    return "title-content";
  }
  if (roles.includes("title")) return "title-only";
  return "title-content";
}

/** One gallery card for a file layout: its thumbnail definition and the name the file gives it. */
export interface NativeLayoutGalleryEntry {
  readonly layoutPart: string;
  readonly name: string | null;
  readonly definition: LayoutDefinition;
}

/**
 * The New Slide / Layout gallery of an imported deck: every layout of the
 * anchor slide's master, in the file's order, drawn from its placeholders'
 * geometry (EMU, made fractions of the slide size) and named as the file
 * names it.
 */
export function nativeLayoutGalleryEntries(
  receipt: NativeLayoutCascadeReceipt,
  slideSize: { readonly width: number; readonly height: number },
): NativeLayoutGalleryEntry[] {
  const width = slideSize.width > 0 ? slideSize.width : 1;
  const height = slideSize.height > 0 ? slideSize.height : 1;
  const fraction = (value: number, total: number) =>
    Number.isFinite(value) ? Math.min(1, Math.max(0, value / total)) : 0;
  return receipt.layouts
    .filter((entry) => entry.masterPart === receipt.masterPart)
    .map((entry) => {
      const typed = nativeLayoutIdFromType(entry.layoutType);
      const placeholders: LayoutPlaceholder[] = [];
      for (const placeholder of entry.placeholders) {
        if (NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES.has(placeholder.placeholderType)) continue;
        const role = nativePlaceholderRole(placeholder.placeholderType);
        if (!role || !nativeLayoutGeometryIsComplete(placeholder.layoutGeometry)) continue;
        const geometry = placeholder.layoutGeometry;
        placeholders.push({
          role: role.role,
          placeholderKey: role.placeholder,
          left: fraction(geometry.left, width),
          top: fraction(geometry.top, height),
          width: fraction(geometry.width, width),
          height: fraction(geometry.height, height),
        });
      }
      return {
        layoutPart: entry.layoutPart,
        name: typeof entry.layoutName === "string" && entry.layoutName.trim() ? entry.layoutName.trim() : null,
        definition: {
          id: nativeLayoutRef(entry.layoutPart) as LayoutId,
          nameKey: LAYOUT_DEFINITIONS.find((layout) => layout.id === typed)?.nameKey ?? "custom",
          placeholders,
        },
      };
    });
}

function nativePlaceholderKey(
  placeholder: Pick<
    NativeLayoutPlaceholderReceipt,
    "placeholderType" | "placeholderIndex"
  >,
): string {
  return `${placeholder.placeholderType}:${Number(placeholder.placeholderIndex)}`;
}

function nativeLayoutTargetForPlaceholder(
  target: NativeLayoutCatalogEntryReceipt,
  source: NativeLayoutPlaceholderReceipt,
): NativeLayoutPlaceholderReceipt | null {
  const exact = target.placeholders.filter(
    (candidate) => nativePlaceholderKey(candidate) === nativePlaceholderKey(source),
  );
  if (exact.length === 1) return exact[0];
  if (
    exact.length === 0 &&
    (source.placeholderType === "title" || source.placeholderType === "ctrTitle")
  ) {
    const compatible = target.placeholders.filter(
      (candidate) =>
        Number(candidate.placeholderIndex) === Number(source.placeholderIndex) &&
        (candidate.placeholderType === "title" ||
          candidate.placeholderType === "ctrTitle"),
    );
    return compatible.length === 1 ? compatible[0] : null;
  }
  return null;
}

/** Placeholder types that hold a slide's body content and stand in for one
 * another when a slide changes layout (the API's `_PPTX_LAYOUT_CONTENT_TYPES`). */
const NATIVE_LAYOUT_CONTENT_TYPES = new Set([
  "body",
  "obj",
  "subTitle",
  "tbl",
  "chart",
  "dgm",
  "media",
  "clipArt",
]);

function nativeLayoutTitleLike(type: string): boolean {
  return type === "title" || type === "ctrTitle";
}

/**
 * Where each of the slide's placeholders goes when the slide takes `target`
 * (its shape id to the target layout's placeholder), as PowerPoint's Layout
 * moves content: the placeholder of the same type and index, else a title to
 * the layout's one title, else body content to the content placeholder of the
 * same index (a subtitle to Title and Content's content box). Each target
 * takes one placeholder; one with no place stays where it is on the slide
 * (`applyNativeLayoutToSlide`, and the save pins its frame). The same walk as
 * the API's `_pptx_layout_placeholder_mapping`.
 */
export function nativeLayoutPlaceholderMapping(
  receipt: Pick<NativeLayoutCascadeReceipt, "placeholders">,
  target: NativeLayoutCatalogEntryReceipt,
): Map<string, NativeLayoutPlaceholderReceipt> {
  const mapping = new Map<string, NativeLayoutPlaceholderReceipt>();
  const used = new Set<string>();
  const free = (candidate: NativeLayoutPlaceholderReceipt) =>
    !used.has(nativePlaceholderKey(candidate)) &&
    !NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES.has(candidate.placeholderType) &&
    (!nativePlaceholderRole(candidate.placeholderType) ||
      nativeLayoutGeometryIsComplete(candidate.layoutGeometry));
  const passes: Array<(source: NativeLayoutPlaceholderReceipt, candidate: NativeLayoutPlaceholderReceipt) => boolean> = [
    (source, candidate) => nativePlaceholderKey(candidate) === nativePlaceholderKey(source),
    (source, candidate) =>
      nativeLayoutTitleLike(source.placeholderType) &&
      nativeLayoutTitleLike(candidate.placeholderType),
    (source, candidate) =>
      NATIVE_LAYOUT_CONTENT_TYPES.has(source.placeholderType) &&
      NATIVE_LAYOUT_CONTENT_TYPES.has(candidate.placeholderType) &&
      Number(candidate.placeholderIndex) === Number(source.placeholderIndex),
  ];
  for (const matches of passes) {
    for (const source of receipt.placeholders) {
      if (mapping.has(source.shapeId) || NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES.has(source.placeholderType)) continue;
      const candidates = target.placeholders.filter((candidate) => free(candidate) && matches(source, candidate));
      if (candidates.length !== 1) continue;
      mapping.set(source.shapeId, candidates[0]);
      used.add(nativePlaceholderKey(candidates[0]));
    }
  }
  // A band placeholder (date, footer, slide number) keeps its own kind.
  for (const source of receipt.placeholders) {
    if (!NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES.has(source.placeholderType)) continue;
    const candidates = target.placeholders.filter(
      (candidate) =>
        !used.has(nativePlaceholderKey(candidate)) &&
        nativePlaceholderKey(candidate) === nativePlaceholderKey(source),
    );
    if (candidates.length !== 1) continue;
    mapping.set(source.shapeId, candidates[0]);
    used.add(nativePlaceholderKey(candidates[0]));
  }
  return mapping;
}

const NATIVE_LAYOUT_HASH = /^[0-9a-f]{64}$/;
// The most layouts one master's catalog in a layout receipt may list. A
// master has no layout limit in PowerPoint; 64 refused the receipt the


// `_PPTX_LAYOUT_CATALOG_MAX` in the save planner.
export const NATIVE_LAYOUT_CATALOG_MAX = 1024;





export const NATIVE_LAYOUT_CATALOG_PLACEHOLDERS_MAX = 256;
const NATIVE_SLIDE_ID = /^sld[1-9][0-9]*$/;
const NATIVE_SLIDE_INSERTION_KEYS = new Set([
  "schemaVersion",
  "kind",
  "modelSlideId",
  "packageFingerprint",
  "presentationPart",
  "anchorSlideId",
  "anchorSlidePart",
  "layoutPart",
  "layoutFingerprint",
  "masterPart",
  "masterFingerprint",
]);
const NATIVE_LAYOUT_REL_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout",
]);
const NATIVE_MASTER_REL_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster",
]);
const NATIVE_THEME_REL_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/theme",
]);
const NATIVE_LAYOUT_RECEIPT_KEYS = new Set([
  "schemaVersion",
  "packageFingerprint",
  "presentationPart",
  "slidePart",
  "slideId",
  "layoutRelId",
  "layoutRelType",
  "layoutPart",
  "layoutFingerprint",
  "masterPart",
  "masterRelType",
  "masterFingerprint",
  "themePart",
  "themeRelType",
  "themeFingerprint",
  "layoutType",
  "layoutName",
  "placeholders",
  "layouts",
  "resetAuthority",
]);
const NATIVE_LAYOUT_REQUIRED_RECEIPT_KEYS = new Set(
  [...NATIVE_LAYOUT_RECEIPT_KEYS].filter((key) => key !== "resetAuthority"),
);
const NATIVE_LAYOUT_CATALOG_KEYS = new Set([
  "layoutPart",
  "layoutFingerprint",
  "masterPart",
  "masterRelType",
  "layoutType",
  "layoutName",
  "placeholders",
]);
const NATIVE_LAYOUT_PLACEHOLDER_KEYS = new Set([
  "shapeId",
  "placeholderType",
  "placeholderIndex",
  "nativeIndexPresent",
  "owner",
  "localXfrm",
  "localStyle",
  "layoutGeometry",
  "layoutStyle",
]);
const NATIVE_SLIDE_PLACEHOLDER_KEYS = new Set([
  ...NATIVE_LAYOUT_PLACEHOLDER_KEYS,
  "layoutShapeId",
]);
const NATIVE_LAYOUT_GEOMETRY_KEYS = new Set([
  "left",
  "top",
  "width",
  "height",
  "rotation",
]);
const NATIVE_LAYOUT_NESTED_STYLE_KEYS = [
  "fontFamily",
  "fontSize",
  "color",
  "bold",
  "italic",
  "underline",
  "strike",
  "align",
  "verticalAlign",
  "writingMode",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
  "letterSpacing",
  "textShadow",
  // The shadow as the layout's text style states it (PPTX readings row 805).
  "textShadowEffect",
  "indentLevel",
  // A master title drawn in capitals (`a:defRPr@cap`, PPT-B362) and a
  // double underline (`@u="dbl"`), as the box's own style states them.
  "caps",
  "underlineStyle",
  // `a:bodyPr@wrap`: the editor keeps it in the box's style (PPT-B360,
  // `HOISTED_TEXT_STYLE_KEYS`), wherever the importer states it.
  "textWrap",
] as const;
const NATIVE_LAYOUT_TOP_LEVEL_STYLE_KEYS = [
  "verticalOverflow",
  "textInsets",
  "autofit",
  "fill",
  // A layout text frame's solid-fill Transparency (API readings row 155).
  // Missing here, the editor refused every receipt of such a layout, so New
  
  "fillTransparency",
  "fillTo",
  "gradientAngle",
  "outlineColor",
  "outlineColorTo",
  "outlineGradientAngle",
  "outlineWidth",
  "outlineStyle",
  // `a:ln@cmpd` / `a:ln@algn` of a layout box's border (PPT-B368).
  "outlineCompound",
  "outlineAlign",
  "shadow",
  "bulletChar",
  "numberFormat",
  // The theme slot of a layout box's solid fill / outline beside its colour.
  "themeColors",
] as const;
/**
 * Every field a layout baseline (`layoutStyle`) may state: the server's
 * `layout_receipt_fields.LAYOUT_STYLE_FIELDS`, which writes and checks it.
 * One field missing here refused the WHOLE receipt, and New Slide, Layout
 * and Reset were unavailable on the deck: every Google Slides export
 * (`textWrap`), every deck on PowerPoint 2007's default template (its
 * Section Header title is in capitals) and every layout box in a theme
 * colour. The tables are pinned to the server's by a test.
 */
export const NATIVE_LAYOUT_STYLE_FIELDS: readonly string[] = [
  ...NATIVE_LAYOUT_NESTED_STYLE_KEYS,
  ...NATIVE_LAYOUT_TOP_LEVEL_STYLE_KEYS,
];








const NATIVE_LAYOUT_CARRIED_STYLE_KEYS = [
  "tabStops",
  "latinLineBreak",
  // A layout box's theme references with a shade or the like (its `p:style`
  // outline reference, API readings row 891), carried beside `themeColors`
  // and not restated by Layout / Reset (the server's
  // `LAYOUT_STYLE_CARRIED_FIELDS`). Refused here, every slide on a layout
  // whose placeholder its `p:style` paints with a shaded outline was a
  
  "themeColorsTransformed",
  // Whether the layout's face follows the theme's heading or body font
  // (API readings row 843), carried beside `fontFamily`.
  "themeFontRole",
] as const;
const NATIVE_LAYOUT_STYLE_KEYS: ReadonlySet<string> = new Set([
  ...NATIVE_LAYOUT_STYLE_FIELDS,
  ...NATIVE_LAYOUT_CARRIED_STYLE_KEYS,
]);

/** A baseline's tab stops as the importer states them: up to 20 positions in
 * design px, a left stop bare, any other `{ pos, align }`. */
function nativeLayoutTabStops(value: unknown): boolean {
  if (!Array.isArray(value) || value.length > 20) return false;
  return value.every((stop) => {
    const pos = stop && typeof stop === "object" && !Array.isArray(stop) ? (stop as Record<string, unknown>).pos : stop;
    if (typeof pos !== "number" || !Number.isFinite(pos) || pos <= 0 || pos > 4096) return false;
    if (typeof stop === "number") return true;
    const record = stop as Record<string, unknown>;
    return (
      Object.keys(record).every((key) => key === "pos" || key === "align") &&
      ["center", "right", "decimal"].includes(String(record.align))
    );
  });
}

/** A layout box's theme colour references: `{ fill: { slot, rgb }, ... }`. */
function nativeLayoutThemeColors(value: unknown): boolean {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const entries = Object.entries(value as Record<string, unknown>);
  return (
    entries.length <= 8 &&
    JSON.stringify(value).length <= 2_048 &&
    entries.every(
      ([key, reference]) =>
        key.length <= 64 &&
        !!reference &&
        typeof reference === "object" &&
        !Array.isArray(reference) &&
        Object.values(reference as Record<string, unknown>).every(
          (item) =>
            typeof item === "boolean" ||
            (typeof item === "number" && Number.isFinite(item)) ||
            (typeof item === "string" && item.length <= 128 && !/[\u0000-\u001f\u007f]/.test(item)),
        ),
    )
  );
}
/** A layout box's theme references with transforms:
 * `{ outlineColor: { slot, transforms: [[kind, value], ...], rgb } }`, the
 * page schema's `presentationThemeColorsTransformedSchema`. */
function nativeLayoutThemeColorsTransformed(value: unknown): boolean {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const entries = Object.entries(value as Record<string, unknown>);
  return (
    entries.length > 0 &&
    entries.every(([field, reference]) => {
      if (!["fill", "stroke", "outlineColor"].includes(field)) return false;
      if (!reference || typeof reference !== "object" || Array.isArray(reference)) return false;
      const { slot, transforms, rgb, ...rest } = reference as Record<string, unknown>;
      return (
        Object.keys(rest).length === 0 &&
        typeof slot === "string" &&
        /^[A-Za-z0-9]{1,16}$/.test(slot) &&
        typeof rgb === "string" &&
        /^#[0-9a-fA-F]{6}$/.test(rgb) &&
        Array.isArray(transforms) &&
        transforms.length >= 1 &&
        transforms.length <= 8 &&
        transforms.every(
          (step) =>
            Array.isArray(step) &&
            step.length === 2 &&
            typeof step[0] === "string" &&
            /^[A-Za-z]{1,16}$/.test(step[0]) &&
            typeof step[1] === "number" &&
            Number.isFinite(step[1]),
        )
      );
    })
  );
}
const NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES = new Set([
  "dt",
  "ftr",
  "hdr",
  "sldNum",
]);

/** Strictly parse the source-less slide receipt persisted by New/Duplicate. */
export function parseNativeSlideInsertionReceipt(
  value: unknown,
): NativeSlideInsertionReceipt | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (
    record.schemaVersion !== PPTX_NATIVE_SLIDE_INSERTION_RECEIPT_VERSION ||
    (record.kind !== "new" && record.kind !== "duplicate") ||
    Object.keys(record).length !== NATIVE_SLIDE_INSERTION_KEYS.size ||
    Object.keys(record).some((key) => !NATIVE_SLIDE_INSERTION_KEYS.has(key))
  ) {
    return null;
  }
  const modelSlideId = exactNativeReceiptString(record.modelSlideId, 256);
  const packageFingerprint = exactNativeReceiptString(
    record.packageFingerprint,
    64,
  );
  const presentationPart = exactNativeReceiptString(
    record.presentationPart,
    512,
  );
  const anchorSlideId = exactNativeReceiptString(record.anchorSlideId, 128);
  const anchorSlidePart = exactNativeReceiptString(record.anchorSlidePart, 512);
  const layoutPart = exactNativeReceiptString(record.layoutPart, 512);
  const layoutFingerprint = exactNativeReceiptString(
    record.layoutFingerprint,
    64,
  );
  const masterPart = exactNativeReceiptString(record.masterPart, 512);
  const masterFingerprint = exactNativeReceiptString(
    record.masterFingerprint,
    64,
  );
  if (
    !modelSlideId ||
    !packageFingerprint ||
    !NATIVE_LAYOUT_HASH.test(packageFingerprint) ||
    presentationPart !== "ppt/presentation.xml" ||
    !anchorSlideId ||
    !NATIVE_SLIDE_ID.test(anchorSlideId) ||
    !anchorSlidePart ||
    !layoutPart ||
    !layoutFingerprint ||
    !NATIVE_LAYOUT_HASH.test(layoutFingerprint) ||
    !masterPart ||
    !masterFingerprint ||
    !NATIVE_LAYOUT_HASH.test(masterFingerprint)
  ) {
    return null;
  }
  return {
    schemaVersion: PPTX_NATIVE_SLIDE_INSERTION_RECEIPT_VERSION,
    kind: record.kind,
    modelSlideId,
    packageFingerprint,
    presentationPart,
    anchorSlideId,
    anchorSlidePart,
    layoutPart,
    layoutFingerprint,
    masterPart,
    masterFingerprint,
  };
}
const NATIVE_RESET_AUTHORITY_KEYS = new Set([
  "version",
  "disposition",
  "blocker",
  "placeholders",
]);
const NATIVE_RESET_PLACEHOLDER_KEYS = new Set([
  "shapeId",
  "fingerprint",
  "overrides",
]);
const NATIVE_RESET_OVERRIDE_KEYS = new Set([
  "kind",
  "path",
  "action",
  "sha256",
]);
const NATIVE_RESET_OVERRIDE_CONTRACT = Object.freeze({
  geometry: { action: "remove", path: /^(?:xfrm|spPr\/xfrm)$/ },
  "shape-fill": { action: "remove", path: /^spPr\/(?:solidFill|noFill)$/ },
  "shape-line": { action: "remove", path: /^spPr\/ln$/ },
  "shape-effect": { action: "remove", path: /^spPr\/effectLst$/ },
  "body-format": { action: "clear", path: /^txBody\/bodyPr$/ },
  "paragraph-format": {
    action: "remove",
    path: /^txBody\/p\[(?:0|[1-9][0-9]{0,5})\]\/pPr$/,
  },
  "run-format": {
    action: "clear-format",
    path:
      /^txBody\/p\[(?:0|[1-9][0-9]{0,5})\]\/run\[(?:0|[1-9][0-9]{0,5})\]\/rPr$/,
  },
} satisfies Record<
  NativePlaceholderResetOverrideKind,
  { action: NativePlaceholderResetOverrideReceipt["action"]; path: RegExp }
>);

function nativeLayoutGeometry(value: unknown): Record<string, number> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  // A placeholder's flips (`flipH` / `flipV`, booleans the import states
  // beside its box) are no part of the box a layout places: a flipped title
  
  
  if (["flipH", "flipV"].some((key) => key in record && typeof record[key] !== "boolean")) return null;
  const keys = Object.keys(record).filter((key) => key !== "flipH" && key !== "flipV");
  if (
    keys.some((key) => !NATIVE_LAYOUT_GEOMETRY_KEYS.has(key)) ||
    keys.some(
      (key) =>
        typeof record[key] !== "number" ||
        !Number.isFinite(record[key]) ||
        Math.abs(record[key] as number) > 100_000_000_000,
    ) ||
    (("width" in record || "height" in record) &&
      [record.width, record.height].some(
        (entry) => entry !== undefined && (entry as number) <= 0,
      ))
  ) {
    return null;
  }
  return Object.fromEntries(keys.map((key) => [key, record[key] as number]));
}

function nativeLayoutGeometryIsComplete(
  value: Record<string, number>,
): boolean {
  return (
    typeof value.left === "number" &&
    Number.isFinite(value.left) &&
    typeof value.top === "number" &&
    Number.isFinite(value.top) &&
    typeof value.width === "number" &&
    Number.isFinite(value.width) &&
    value.width > 0 &&
    typeof value.height === "number" &&
    Number.isFinite(value.height) &&
    value.height > 0 &&
    (value.rotation === undefined ||
      (typeof value.rotation === "number" && Number.isFinite(value.rotation)))
  );
}

function nativeLayoutStyle(value: unknown): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.some((key) => !NATIVE_LAYOUT_STYLE_KEYS.has(key))) return null;
  for (const key of keys) {
    const item = record[key];
    if (key === "themeColors") {
      if (!nativeLayoutThemeColors(item)) return null;
      continue;
    }
    if (key === "tabStops") {
      if (!nativeLayoutTabStops(item)) return null;
      continue;
    }
    if (key === "themeColorsTransformed") {
      if (!nativeLayoutThemeColorsTransformed(item)) return null;
      continue;
    }
    if (key === "themeFontRole") {
      if (item !== "heading" && item !== "body") return null;
      continue;
    }
    if (key === "textShadowEffect") {
      // The shadow the layout's text style states (reading row 805), in the
      // shape shadow model: an object, refused as a scalar here, which made
      // every slide on a layout whose placeholder casts a stated shadow a
      
      
      // "native layout cascade receipt is malformed" on every slide).
      try {
        if (!normalizeOfficeShapeShadow(item)) return null;
      } catch {
        return null;
      }
      continue;
    }
    if (key === "textInsets") {
      if (!item || typeof item !== "object" || Array.isArray(item)) return null;
      const inset = item as Record<string, unknown>;
      if (
        Object.keys(inset).some(
          (side) => !["left", "top", "right", "bottom"].includes(side),
        ) ||
        Object.values(inset).some(
          (entry) =>
            typeof entry !== "number" ||
            !Number.isFinite(entry) ||
            entry < 0 ||
            entry > 100_000_000_000,
        )
      ) {
        return null;
      }
      continue;
    }
    if (
      !(
        typeof item === "boolean" ||
        (typeof item === "number" &&
          Number.isFinite(item) &&
          Math.abs(item) <= 100_000_000_000) ||
        (typeof item === "string" &&
          item.length <= 512 &&
          !/[\u0000-\u001f\u007f]/.test(item))
      )
    ) {
      return null;
    }
  }
  return { ...record };
}

function exactNativeReceiptString(
  value: unknown,
  maxLength: number,
): string | null {
  return typeof value === "string" &&
    value.length > 0 &&
    value.length <= maxLength &&
    value.trim() === value &&
    !/[\u0000-\u001f\u007f]/.test(value)
    ? value
    : null;
}









function nativeLayoutDisplayName(value: unknown): string | null {
  return typeof value === "string" && value.length > 0 && value.length <= 4_000 ? value : null;
}

function nativePlaceholderResetAuthority(
  value: unknown,
  slidePlaceholders: readonly NativeLayoutPlaceholderReceipt[],
): NativePlaceholderResetAuthorityReceipt | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (
    Object.keys(record).length !== NATIVE_RESET_AUTHORITY_KEYS.size ||
    Object.keys(record).some((key) => !NATIVE_RESET_AUTHORITY_KEYS.has(key)) ||
    record.version !== PPTX_PLACEHOLDER_RESET_RECEIPT_VERSION ||
    (record.disposition !== "native" && record.disposition !== "preserve-only") ||
    !Array.isArray(record.placeholders) ||
    record.placeholders.length > 64
  ) {
    return null;
  }
  const blocker =
    record.blocker === null
      ? null
      : exactNativeReceiptString(record.blocker, 160);
  if (
    (record.disposition === "native" && blocker !== null) ||
    (record.disposition === "preserve-only" &&
      (!blocker || record.placeholders.length !== 0))
  ) {
    return null;
  }
  const placeholders: NativePlaceholderResetReceipt[] = [];
  const seenShapeIds = new Set<string>();
  for (const value of record.placeholders) {
    if (!value || typeof value !== "object" || Array.isArray(value)) return null;
    const item = value as Record<string, unknown>;
    const shapeId = exactNativeReceiptString(item.shapeId, 64);
    const fingerprint = exactNativeReceiptString(item.fingerprint, 64);
    if (
      Object.keys(item).length !== NATIVE_RESET_PLACEHOLDER_KEYS.size ||
      Object.keys(item).some((key) => !NATIVE_RESET_PLACEHOLDER_KEYS.has(key)) ||
      !shapeId ||
      !/^shp[1-9][0-9]*$/.test(shapeId) ||
      seenShapeIds.has(shapeId) ||
      !fingerprint ||
      !NATIVE_LAYOUT_HASH.test(fingerprint) ||
      !Array.isArray(item.overrides) ||
      item.overrides.length > 256
    ) {
      return null;
    }
    const overrides: NativePlaceholderResetOverrideReceipt[] = [];
    const seenPaths = new Set<string>();
    for (const value of item.overrides) {
      if (!value || typeof value !== "object" || Array.isArray(value)) return null;
      const override = value as Record<string, unknown>;
      const kind = override.kind as NativePlaceholderResetOverrideKind;
      const contract = NATIVE_RESET_OVERRIDE_CONTRACT[kind];
      const path = exactNativeReceiptString(override.path, 160);
      const sha256 = exactNativeReceiptString(override.sha256, 64);
      if (
        Object.keys(override).length !== NATIVE_RESET_OVERRIDE_KEYS.size ||
        Object.keys(override).some((key) => !NATIVE_RESET_OVERRIDE_KEYS.has(key)) ||
        !contract ||
        override.action !== contract.action ||
        !path ||
        !contract.path.test(path) ||
        seenPaths.has(path) ||
        !sha256 ||
        !NATIVE_LAYOUT_HASH.test(sha256)
      ) {
        return null;
      }
      seenPaths.add(path);
      overrides.push({
        kind,
        path,
        action: override.action as NativePlaceholderResetOverrideReceipt["action"],
        sha256,
      });
    }
    seenShapeIds.add(shapeId);
    placeholders.push({ shapeId, fingerprint, overrides });
  }
  if (record.disposition === "native") {
    const expected = new Set(slidePlaceholders.map((placeholder) => placeholder.shapeId));
    if (
      placeholders.length !== expected.size ||
      placeholders.some((placeholder) => !expected.has(placeholder.shapeId))
    ) {
      return null;
    }
  }
  return {
    version: PPTX_PLACEHOLDER_RESET_RECEIPT_VERSION,
    disposition: record.disposition,
    blocker,
    placeholders,
  };
}

function nativePlaceholderIdentity(
  value: unknown,
  owner: "slide" | "layout",
): NativeLayoutPlaceholderReceipt | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const allowedKeys =
    owner === "slide"
      ? NATIVE_SLIDE_PLACEHOLDER_KEYS
      : NATIVE_LAYOUT_PLACEHOLDER_KEYS;
  if (Object.keys(record).some((key) => !allowedKeys.has(key))) return null;
  const shapeId = exactNativeReceiptString(record.shapeId, 64);
  const placeholderType = exactNativeReceiptString(
    record.placeholderType,
    32,
  );
  const placeholderIndex = record.placeholderIndex;
  const nativeIndexPresent = record.nativeIndexPresent;
  const validIndex =
    typeof placeholderIndex === "string" &&
    /^(?:0|[1-9][0-9]{0,9})$/.test(placeholderIndex) &&
    Number(placeholderIndex) <= 4_294_967_295;
  if (
    !shapeId ||
    !/^shp[1-9][0-9]*$/.test(shapeId) ||
    !placeholderType ||
    !validIndex ||
    typeof nativeIndexPresent !== "boolean" ||
    record.owner !== owner ||
    (!nativeIndexPresent && placeholderIndex !== "0") ||
    typeof record.localXfrm !== "boolean" ||
    typeof record.localStyle !== "boolean"
  ) {
    return null;
  }
  const layoutShapeId =
    owner === "slide"
      ? exactNativeReceiptString(record.layoutShapeId, 64)
      : undefined;
  const layoutGeometry = nativeLayoutGeometry(record.layoutGeometry);
  const layoutStyle = nativeLayoutStyle(record.layoutStyle);
  if (
    owner === "slide" &&
    (!layoutShapeId || !/^shp[1-9][0-9]*$/.test(layoutShapeId)) ||
    !layoutGeometry ||
    !layoutStyle
  ) {
    return null;
  }
  return {
    shapeId,
    placeholderType,
    placeholderIndex: placeholderIndex as string,
    nativeIndexPresent,
    owner,
    localXfrm: record.localXfrm as boolean,
    localStyle: record.localStyle as boolean,
    ...(owner === "slide" ? { layoutShapeId } : {}),
    layoutGeometry,
    layoutStyle,
  };
}

/**
 * Each catalog parsed once per master: every slide of an imported deck holds
 * its master's catalog as the same array (`presentation-deck-receipts`), so a
 * deck of 2,000 slides parses it once, not 2,000 times. The parsed entries are
 * shared and must not be changed.
 */
const parsedNativeLayoutCatalogs = new WeakMap<
  object,
  Map<string, readonly NativeLayoutCatalogEntryReceipt[] | null>
>();

function parseNativeLayoutCatalog(
  rawLayouts: readonly unknown[],
  masterPart: string,
): readonly NativeLayoutCatalogEntryReceipt[] | null {
  let byMaster = parsedNativeLayoutCatalogs.get(rawLayouts);
  if (!byMaster) {
    byMaster = new Map();
    parsedNativeLayoutCatalogs.set(rawLayouts, byMaster);
  }
  if (byMaster.has(masterPart)) return byMaster.get(masterPart) ?? null;
  const parsed = parseNativeLayoutCatalogEntries(rawLayouts, masterPart);
  byMaster.set(masterPart, parsed);
  return parsed;
}

function parseNativeLayoutCatalogEntries(
  rawLayouts: readonly unknown[],
  masterPart: string,
): NativeLayoutCatalogEntryReceipt[] | null {
  const layouts: NativeLayoutCatalogEntryReceipt[] = [];
  const catalogIds = new Set<string>();
  for (const value of rawLayouts) {
    if (!value || typeof value !== "object" || Array.isArray(value)) return null;
    const item = value as Record<string, unknown>;
    if (
      Object.keys(item).some((key) => !NATIVE_LAYOUT_CATALOG_KEYS.has(key)) ||
      Object.keys(item).length !== NATIVE_LAYOUT_CATALOG_KEYS.size
    ) {
      return null;
    }
    const catalogLayoutPart = exactNativeReceiptString(item.layoutPart, 512);
    const catalogFingerprint = exactNativeReceiptString(
      item.layoutFingerprint,
      64,
    );
    const catalogMasterPart = exactNativeReceiptString(item.masterPart, 512);
    const catalogMasterRelType = exactNativeReceiptString(
      item.masterRelType,
      160,
    );
    const catalogLayoutType =
      item.layoutType === null
        ? null
        : exactNativeReceiptString(item.layoutType, 64);
    const catalogLayoutName =
      item.layoutName === null
        ? null
        : nativeLayoutDisplayName(item.layoutName);
    if (
      !catalogLayoutPart ||
      !catalogFingerprint ||
      !NATIVE_LAYOUT_HASH.test(catalogFingerprint) ||
      !catalogMasterPart ||
      catalogMasterPart !== masterPart ||
      !catalogMasterRelType ||
      !NATIVE_MASTER_REL_TYPES.has(catalogMasterRelType) ||
      (item.layoutType !== null && !catalogLayoutType) ||
      (item.layoutName !== null && !catalogLayoutName) ||
      catalogIds.has(catalogLayoutPart) ||
      !Array.isArray(item.placeholders) ||
      item.placeholders.length > NATIVE_LAYOUT_CATALOG_PLACEHOLDERS_MAX
    ) {
      return null;
    }
    const catalogPlaceholders = item.placeholders.map((placeholder) =>
      nativePlaceholderIdentity(placeholder, "layout"),
    );
    if (catalogPlaceholders.some((entry) => entry === null)) return null;
    // A layout's placeholders are told apart by their type and index, as
    // PowerPoint inherits by them. Not by the shape id: LibreOffice writes
    // its layouts' band placeholders under the ids of its title and body
    
    // catalog for that made every slide's receipt malformed -- the whole page
    // refused by every catalog command (the AI's edits, the Automation API).
    // The writer still fails closed on an ambiguous layout part.
    const keys = new Set<string>();
    for (const entry of catalogPlaceholders) {
      const key = `${entry!.placeholderType}:${Number(entry!.placeholderIndex)}`;
      if (keys.has(key)) return null;
      keys.add(key);
    }
    catalogIds.add(catalogLayoutPart);
    layouts.push({
      layoutPart: catalogLayoutPart,
      layoutFingerprint: catalogFingerprint,
      masterPart: catalogMasterPart,
      masterRelType: catalogMasterRelType,
      layoutType: catalogLayoutType,
      layoutName: catalogLayoutName,
      placeholders: catalogPlaceholders as Array<{
        shapeId: string;
        placeholderType: string;
        placeholderIndex: string;
        nativeIndexPresent: boolean;
        owner: "layout";
        localXfrm: boolean;
        localStyle: boolean;
        layoutGeometry: Record<string, number>;
        layoutStyle: Record<string, unknown>;
      }>,
    });
  }
  return layouts;
}

/**
 * Strictly parse the bounded native receipt used by New/Layout/Reset.
 * Unknown or partial custody is never repaired into editable authority.
 */
export function parseNativeLayoutCascadeReceipt(
  value: unknown,
): NativeLayoutCascadeReceipt | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (
    record.schemaVersion !== PPTX_LAYOUT_CASCADE_RECEIPT_VERSION ||
    Object.keys(record).some((key) => !NATIVE_LAYOUT_RECEIPT_KEYS.has(key)) ||
    [...NATIVE_LAYOUT_REQUIRED_RECEIPT_KEYS].some(
      (key) => !Object.hasOwn(record, key),
    )
  ) {
    return null;
  }
  const packageFingerprint = exactNativeReceiptString(
    record.packageFingerprint,
    64,
  );
  const presentationPart = exactNativeReceiptString(
    record.presentationPart,
    512,
  );
  const slideId = exactNativeReceiptString(record.slideId, 128);
  const slidePart = exactNativeReceiptString(record.slidePart, 512);
  const layoutRelId = exactNativeReceiptString(record.layoutRelId, 128);
  const layoutRelType = exactNativeReceiptString(record.layoutRelType, 160);
  const layoutPart = exactNativeReceiptString(record.layoutPart, 512);
  const layoutFingerprint = exactNativeReceiptString(
    record.layoutFingerprint,
    64,
  );
  const masterPart = exactNativeReceiptString(record.masterPart, 512);
  const masterRelType = exactNativeReceiptString(record.masterRelType, 160);
  const masterFingerprint = exactNativeReceiptString(
    record.masterFingerprint,
    64,
  );
  const themePart =
    record.themePart === null
      ? null
      : exactNativeReceiptString(record.themePart, 512);
  const themeRelType = exactNativeReceiptString(record.themeRelType, 160);
  const themeFingerprint =
    record.themeFingerprint === null
      ? null
      : exactNativeReceiptString(record.themeFingerprint, 64);
  const layoutType =
    record.layoutType === null
      ? null
      : exactNativeReceiptString(record.layoutType, 64);
  const layoutName =
    record.layoutName === null
      ? null
      : nativeLayoutDisplayName(record.layoutName);
  if (
    !packageFingerprint ||
    !NATIVE_LAYOUT_HASH.test(packageFingerprint) ||
    !presentationPart ||
    !slideId ||
    !slidePart ||
    !layoutRelId ||
    !layoutRelType ||
    !NATIVE_LAYOUT_REL_TYPES.has(layoutRelType) ||
    !layoutPart ||
    !layoutFingerprint ||
    !NATIVE_LAYOUT_HASH.test(layoutFingerprint) ||
    !masterPart ||
    !masterRelType ||
    !NATIVE_MASTER_REL_TYPES.has(masterRelType) ||
    !masterFingerprint ||
    !NATIVE_LAYOUT_HASH.test(masterFingerprint) ||
    (record.themePart !== null && !themePart) ||
    !themeRelType ||
    !NATIVE_THEME_REL_TYPES.has(themeRelType) ||
    (record.themeFingerprint !== null &&
      (!themeFingerprint || !NATIVE_LAYOUT_HASH.test(themeFingerprint))) ||
    (themePart === null) !== (themeFingerprint === null) ||
    (record.layoutType !== null && !layoutType) ||
    (record.layoutName !== null && !layoutName)
  ) {
    return null;
  }
  if (
    !Array.isArray(record.placeholders) ||
    record.placeholders.length > 64 ||
    !Array.isArray(record.layouts) ||
    record.layouts.length === 0 ||
    record.layouts.length > NATIVE_LAYOUT_CATALOG_MAX
  ) {
    return null;
  }
  const placeholders: NativeLayoutPlaceholderReceipt[] = [];
  const placeholderKeys = new Set<string>();
  const placeholderShapeIds = new Set<string>();
  for (const value of record.placeholders) {
    const identity = nativePlaceholderIdentity(value, "slide");
    const key = identity
      ? `${identity.placeholderType}:${Number(identity.placeholderIndex)}`
      : "";
    if (
      !identity ||
      placeholderKeys.has(key) ||
      placeholderShapeIds.has(identity.shapeId)
    ) {
      return null;
    }
    placeholderKeys.add(key);
    placeholderShapeIds.add(identity.shapeId);
    placeholders.push(identity);
  }
  const parsedLayouts = parseNativeLayoutCatalog(record.layouts, masterPart);
  if (!parsedLayouts) return null;
  const layouts = parsedLayouts as NativeLayoutCatalogEntryReceipt[];
  const currentCatalogEntry = layouts.find(
    (entry) => entry.layoutPart === layoutPart,
  );
  if (
    !currentCatalogEntry ||
    currentCatalogEntry.layoutFingerprint !== layoutFingerprint ||
    currentCatalogEntry.layoutType !== layoutType ||
    currentCatalogEntry.layoutName !== layoutName
  ) {
    return null;
  }
  const currentLayoutPlaceholders = new Map(
    currentCatalogEntry.placeholders.map((placeholder) => [
      `${placeholder.placeholderType}:${Number(placeholder.placeholderIndex)}`,
      placeholder,
    ]),
  );
  for (const placeholder of placeholders) {
    const layoutPlaceholder = currentLayoutPlaceholders.get(
      `${placeholder.placeholderType}:${Number(placeholder.placeholderIndex)}`,
    );
    if (
      placeholder.layoutShapeId !== layoutPlaceholder?.shapeId
    ) {
      return null;
    }
  }
  const resetAuthority =
    record.resetAuthority === undefined
      ? undefined
      : nativePlaceholderResetAuthority(record.resetAuthority, placeholders);
  if (record.resetAuthority !== undefined && !resetAuthority) return null;
  return {
    schemaVersion: PPTX_LAYOUT_CASCADE_RECEIPT_VERSION,
    packageFingerprint,
    presentationPart,
    slideId,
    slidePart,
    layoutRelId,
    layoutRelType,
    layoutPart,
    layoutFingerprint,
    masterPart,
    masterRelType,
    masterFingerprint,
    themePart,
    themeRelType,
    themeFingerprint,
    layoutType,
    layoutName,
    placeholders,
    layouts,
    ...(resetAuthority ? { resetAuthority } : {}),
  };
}

/**
 * Look up a layout definition by id, falling back to the {@link DEFAULT_LAYOUT_ID}
 * (`blank`) definition for unknown / stale values. Always returns a definition.
 */
export function getLayoutDefinition(value: unknown): LayoutDefinition {
  const id = normalizeLayoutId(value);
  // The map always contains every LAYOUT_ID, so this is non-null; the `?? blank`
  // is defensive only.
  return (
    LAYOUT_BY_ID.get(id) ??
    LAYOUT_BY_ID.get(DEFAULT_LAYOUT_ID) ??
    LAYOUT_DEFINITIONS[LAYOUT_DEFINITIONS.length - 1]
  );
}

export interface LayoutBlock {
  id: string;
  type: "text";
  /**
   * Layout placeholders are positioned text boxes so they integrate with the
   * editor's existing canvas (drag / resize / present-mode) text-box rendering.
   */
  kind: "textbox";
  text: string;
  /** Carries the placeholder prompt key for empty placeholders. */
  placeholder: string;
  /** Descriptive role of the placeholder within its layout. */
  role: PlaceholderRole;
  left: number;
  top: number;
  width: number;
  height: number;
  [key: string]: unknown;
}

/** Resolve a positive slide dimension, falling back to the standard canvas. */
function resolveBound(value: number | null | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : fallback;
}

/** Stable, dependency-free id generator with a crypto.randomUUID fallback. */
function makeId(index: number): string {
  const g = (
    globalThis as { crypto?: { randomUUID?: () => string } }
  ).crypto;
  if (g?.randomUUID) return g.randomUUID();
  return `layout-block-${index}-${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * Materialize a layout's placeholder template into concrete presentation text
 * blocks positioned (in EMUs) within `slideBounds`. Each block is an empty
 * positioned text block carrying a `placeholder` prompt key + `role` so the
 * editor can render a hint and the slide enters canvas (geometry) mode.
 *
 * - An unknown / stale layout id falls back to {@link DEFAULT_LAYOUT_ID}
 *   (`blank`) and yields an EMPTY array (no placeholders).
 * - Geometry is clamped to the slide so a block can never overflow the canvas
 *   even with rounding, and width/height stay strictly positive.
 * - The result is deterministic in geometry; only block `id`s are random.
 */
export function layoutBlocks(
  value: unknown,
  slideBounds?: { width?: number | null; height?: number | null },
  source: LayoutFrameSource = BUILTIN_LAYOUT_FRAMES,
): LayoutBlock[] {
  const width = resolveBound(slideBounds?.width, DEFAULT_SLIDE_WIDTH_EMU);
  const height = resolveBound(slideBounds?.height, DEFAULT_SLIDE_HEIGHT_EMU);

  return layoutPlaceholders(value, source).map((placeholder, index) => {
    // Clamp fractional origin to [0, 1) so a block always starts on-slide.
    const fLeft = clampFraction(placeholder.left, 0, 0.99);
    const fTop = clampFraction(placeholder.top, 0, 0.99);
    // Width/height must stay positive and fit within the remaining canvas.
    const fWidth = clampFraction(placeholder.width, 0.01, 1 - fLeft);
    const fHeight = clampFraction(placeholder.height, 0.01, 1 - fTop);

    const left = Math.round(width * fLeft);
    const top = Math.round(height * fTop);
    const blockWidth = Math.max(1, Math.round(width * fWidth));
    const blockHeight = Math.max(1, Math.round(height * fHeight));

    return {
      id: makeId(index),
      type: "text" as const,
      kind: "textbox" as const,
      text: "",
      placeholder: placeholder.placeholderKey,
      role: placeholder.role,
      left: Math.min(left, width - blockWidth),
      top: Math.min(top, height - blockHeight),
      width: blockWidth,
      height: blockHeight,
      // A text placeholder shrinks its words to fit when they overflow it
      // ("Shrink text on overflow", `a:normAutofit`), as every placeholder
      // of PowerPoint's own layouts does; a picture placeholder holds none.
      ...(placeholder.role === "picture" ? {} : { autofit: "shrink" as const }),
      // The vertical layouts ship their placeholders already in vertical-rl,
      // and an Office layout its level-1 text facts; everything downstream
      // (editor, export) treats both as ordinary style.
      ...layoutPlaceholderStyle(placeholder),
    };
  });
}

function layoutPlaceholderStyle(
  placeholder: LayoutPlaceholder,
): { style: Record<string, unknown> } | Record<string, never> {
  const style: Record<string, unknown> = {};
  if (placeholder.vertical) style.writingMode = "vertical-rl";
  for (const key of LAYOUT_TEXT_STYLE_KEYS) {
    const value = placeholder.text?.[key];
    if (value !== undefined) style[key] = value;
  }
  return Object.keys(style).length > 0 ? { style } : {};
}

const OFFICE_ALIGN: Readonly<
  Record<string, LayoutPlaceholderText["align"] | undefined>
> = { l: undefined, ctr: "center", r: "right", just: "justify", dist: "distribute" };
const OFFICE_ANCHOR: Readonly<
  Record<string, LayoutPlaceholderText["verticalAlign"] | undefined>
> = { t: undefined, ctr: "middle", b: "bottom" };

/** The block's design px of a DrawingML hundredths-of-a-point size (960 px per 10 in slide). */
function officeFontSizePx(sz: number): number {
  return Math.round(ptToPx((sz / 100)) * 100) / 100;
}

function officePlaceholder(
  frame: PresentationOfficePlaceholderFrame,
  variant: PresentationOfficeThemeVariant,
  builtIn: LayoutPlaceholder | undefined,
  master?: {
    values: OfficeMasterTextValues | undefined;
    own: ReadonlySet<OfficeRunAttribute> | undefined;
  },
  flipColorMap = false,
): LayoutPlaceholder {
  const text: LayoutPlaceholderText = {};
  if (frame.sz) text.fontSize = officeFontSizePx(frame.sz);
  const align = frame.algn ? OFFICE_ALIGN[frame.algn] : undefined;
  if (align) text.align = align;
  const verticalAlign = frame.anchor ? OFFICE_ANCHOR[frame.anchor] : undefined;
  if (verticalAlign) text.verticalAlign = verticalAlign;
  if (frame.b === true) text.bold = true;
  if (frame.cap === "all" || frame.cap === "small") text.caps = frame.cap;
  const color = presentationOfficeFrameTextColor(
    flipColorMap && frame.clr && !frame.clr.startsWith("#")
      ? { ...frame, clr: flippedColorSlot(frame.clr) }
      : frame,
    variant,
  );
  if (color) text.color = color;
  // The master's text style, where the placeholder's own list style is
  // silent: PowerPoint resolves each attribute from the closest statement.
  const values = master?.values;
  if (values) {
    const own = master?.own;
    if (values.fontSize !== undefined && !own?.has("s")) {
      text.fontSize = officeFontSizePx(Math.round(values.fontSize * 100));
    }
    if (values.bold !== undefined && !own?.has("b")) {
      if (values.bold) text.bold = true;
      else delete text.bold;
    }
    if (values.italic !== undefined && !own?.has("i")) {
      if (values.italic) text.italic = true;
      else delete text.italic;
    }
    if (values.color !== undefined && !own?.has("c")) text.color = values.color;
  }
  return {
    role: frame.role,
    // The prompt of the built-in placeholder in the same slot when it plays
    // the same role (a Comparison heading, a caption), else the role's own.
    placeholderKey:
      builtIn && builtIn.role === frame.role
        ? builtIn.placeholderKey
        : OFFICE_PLACEHOLDER_KEY[frame.role],
    ...(frame.vert ? { vertical: true } : {}),
    left: frame.x / PRESENTATION_OFFICE_CATALOG_SLIDE_WIDTH,
    top: frame.y / PRESENTATION_OFFICE_CATALOG_SLIDE_HEIGHT,
    width: frame.cx / PRESENTATION_OFFICE_CATALOG_SLIDE_WIDTH,
    height: frame.cy / PRESENTATION_OFFICE_CATALOG_SLIDE_HEIGHT,
    ...(Object.keys(text).length > 0 ? { text } : {}),
  };
}

const OFFICE_PLACEHOLDER_KEY: Readonly<Record<PlaceholderRole, string>> = {
  title: "title",
  subtitle: "subtitle",
  body: "content",
  content: "content",
  picture: "picture",
};

/**
 * A catalog frame with Slide Master view's moves applied (PPT-B030): the
 * layout placeholder's own moved frame, else -- for a placeholder that states
 * no frame of its own and so inherits the master's -- the master title's or
 * body's moved frame, else the catalog's.
 */
export function officeFrameWithMasterEdits(
  source: LayoutFrameSource,
  layoutId: string,
  slot: number,
  frame: PresentationOfficePlaceholderFrame,
): PresentationOfficePlaceholderFrame {
  const routed = resolveOfficeLayoutRef(source, layoutId);
  if (routed.source !== source) return officeFrameWithMasterEdits(routed.source, routed.layoutId as string, slot, frame);
  if (source.kind !== "office") return frame;
  const own = source.layoutFrames?.[layoutId]?.[String(slot)];
  const inherited =
    !own &&
    source.masterFrames &&
    !officeSlotOwnsGeometry(source.themeId, source.variantId, officeCatalogLayoutId(source, layoutId), slot)
      ? source.masterFrames[frame.role === "title" ? "title" : "body"]
      : undefined;
  const rect = own ?? inherited;
  if (!rect) return frame;
  return { ...frame, x: rect[0], y: rect[1], cx: rect[2], cy: rect[3] };
}

/**
 * The placeholders of a layout under a frame source, in the editor's pour
 * order: the Office theme layout's frames (catalog) when the source names a
 * theme that has the layout, else the built-in definition's. `blank` and an
 * unknown id yield none.
 */
export function layoutPlaceholders(
  value: unknown,
  source: LayoutFrameSource = BUILTIN_LAYOUT_FRAMES,
): LayoutPlaceholder[] {
  const routed = resolveOfficeLayoutRef(source, value);
  if (routed.layoutId !== value) return layoutPlaceholders(routed.layoutId, routed.source);
  const custom = typeof value === "string" && officeCatalogLayoutId(source, value) !== value;
  const layout = getLayoutDefinition(custom ? "title-only" : value);
  if (source.kind === "office" && layout.id !== DEFAULT_LAYOUT_ID) {
    // A custom layout pours its own slots (PPT-B030); its catalog slots are
    // the Title Only layout's.
    const slots = officeLayoutSlots(source, custom ? (value as string) : layout.id);
    const variant = presentationOfficeThemeVariant(source.themeId, source.variantId);
    if (slots && variant) {
      return slots.map((slot) => {
        const frame = slot.frame;
        const masterText = source.masterText;
        const values = !masterText
          ? undefined
          : frame.role === "title"
            ? masterText.title?.["1"]
            : masterText.body?.["1"];
        const ladder = slot.ladder;
        return officePlaceholder(
          frame,
          variant,
          slot.catalogSlot !== undefined ? layout.placeholders[slot.catalogSlot] : undefined,
          values
            ? {
                values,
                own: officeSlotOwnAttributes(
                  source.themeId,
                  source.variantId,
                  officeCatalogLayoutId(source, ladder.layoutId),
                  ladder.slot,
                )[0],
              }
            : undefined,
          source.flipColorMap === true,
        );
      });
    }
  }
  return layout.placeholders;
}

/** One placeholder of an Office theme layout as Slide Master view left it. */
export interface OfficeLayoutSlot {
  /** `c<n>` (the catalog layout's slot n) or `a<n>` (the n-th added placeholder). */
  readonly ref: string;
  readonly catalogSlot?: number;
  readonly addedIndex?: number;
  readonly kind?: OfficePlaceholderKind;
  /** The frame with Slide Master view's moves applied, on the catalog slide. */
  readonly frame: PresentationOfficePlaceholderFrame;
  /** Where the slot's own list style and body ladder are read: itself, or the master body's. */
  readonly ladder: { readonly layoutId: string; readonly slot: number };
}

/** PowerPoint's `p:ph@type` of each Insert Placeholder kind; content has none (obj). */
export const OFFICE_PLACEHOLDER_KIND_TYPES: Readonly<Record<OfficePlaceholderKind, string | undefined>> = {
  content: undefined,
  "content-vertical": undefined,
  text: "body",
  "text-vertical": "body",
  picture: "pic",
  chart: "chart",
  table: "tbl",
  smartart: "dgm",
  media: "media",
  "online-image": "clipArt",
};

/** The first `p:ph@idx` an added placeholder takes (the footers hold 10..12). */
export const OFFICE_ADDED_PLACEHOLDER_FIRST_IDX = 13;

/**
 * An Office layout's placeholders in pour order, after Slide Master view's
 * inserts, deletes and moves (PPT-B030); undefined when the theme has no such
 * layout. An added placeholder takes the master body's text (the content
 * slot of Title and Content), as PowerPoint's inserted placeholders inherit
 * the master's body style.
 */
export function officeLayoutSlots(
  source: LayoutFrameSource,
  layoutId: string,
): OfficeLayoutSlot[] | undefined {
  const routed = resolveOfficeLayoutRef(source, layoutId);
  if (routed.layoutId !== layoutId) return officeLayoutSlots(routed.source, routed.layoutId as string);
  if (source.kind !== "office") return undefined;
  if (isOfficeCustomLayoutId(layoutId) && !(source.customLayouts ?? []).includes(layoutId)) return undefined;
  const catalog = presentationOfficeLayoutFrames(
    source.themeId,
    source.variantId,
    officeCatalogLayoutId(source, layoutId),
  );
  if (!catalog) return undefined;
  const added = source.addedPlaceholders?.[layoutId] ?? [];
  const refs = source.layoutSlots?.[layoutId] ?? catalog.map((_, index) => `c${index}`);
  const bodyBase = presentationOfficeLayoutFrames(source.themeId, source.variantId, "title-content")?.[1];
  const slots: OfficeLayoutSlot[] = [];
  for (const ref of refs) {
    const index = Number(ref.slice(1));
    if (ref.startsWith("c")) {
      const catalogFrame = catalog[index];
      if (!catalogFrame) continue;
      slots.push({
        ref,
        catalogSlot: index,
        frame: officeFrameWithMasterEdits(source, layoutId, index, catalogFrame),
        ladder: { layoutId, slot: index },
      });
      continue;
    }
    const entry = added[index];
    if (!entry) continue;
    const type = OFFICE_PLACEHOLDER_KIND_TYPES[entry.kind];
    const role: PresentationOfficePlaceholderFrame["role"] =
      entry.kind === "picture" ? "picture" : type === "body" ? "body" : "content";
    const [x, y, cx, cy] = entry.frame;
    slots.push({
      ref,
      addedIndex: index,
      kind: entry.kind,
      frame: {
        ...(bodyBase && role !== "picture"
          ? {
              ...(bodyBase.sz ? { sz: bodyBase.sz } : {}),
              ...(bodyBase.algn ? { algn: bodyBase.algn } : {}),
              ...(bodyBase.clr ? { clr: bodyBase.clr } : {}),
              ...(bodyBase.mods ? { mods: bodyBase.mods } : {}),
            }
          : {}),
        role,
        ...(type ? { type } : {}),
        idx: String(OFFICE_ADDED_PLACEHOLDER_FIRST_IDX + index),
        x,
        y,
        cx,
        cy,
        ...(entry.kind.endsWith("-vertical") ? { vert: true } : {}),
      },
      // A picture placeholder holds no list; the others read the master body's.
      ladder: role === "picture" ? { layoutId, slot: -1 } : { layoutId: "title-content", slot: 1 },
    });
  }
  return slots;
}

/**
 * A placeholder's text style after its layout template changes from `before`
 * to `after`: every {@link LAYOUT_TEXT_STYLE_KEYS} value the block still holds
 * as the old template stated it (both absent counts) takes the new template's
 * value; a value the author set is kept. Returns the next style, or undefined
 * when nothing is left on it.
 */
export function layoutTextStyleAfterTemplate(
  current: Record<string, unknown> | undefined,
  before: Record<string, unknown> | undefined,
  after: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  const next: Record<string, unknown> = { ...(current ?? {}) };
  for (const key of [...LAYOUT_TEXT_STYLE_KEYS, "writingMode"] as const) {
    if (next[key] !== before?.[key]) continue;
    if (after?.[key] === undefined) delete next[key];
    else next[key] = after[key];
  }
  return Object.keys(next).length > 0 ? next : undefined;
}

/** A block's four geometry values when all are finite numbers. */
function blockRect(
  block: Record<string, unknown>,
): { left: number; top: number; width: number; height: number } | null {
  const read = (key: string) => {
    const value = block[key];
    return typeof value === "number" && Number.isFinite(value) ? value : null;
  };
  const left = read("left");
  const top = read("top");
  const width = read("width");
  const height = read("height");
  return left === null || top === null || width === null || height === null
    ? null
    : { left, top, width, height };
}

/** Is `block` a layout placeholder (made by {@link layoutBlocks}) the deck still owns? */







export function isFilledPlaceholderBlock(block: unknown): boolean {
  const record = block as Record<string, unknown> | null;
  if (!record || record.type === "text") return false;
  if (typeof record.placeholder !== "string" || typeof record.role !== "string") return false;
  const source = record.source as Record<string, unknown> | undefined;
  return !(source && typeof source.slideId === "string");
}

/** A filled placeholder moved to its place in the layout (its own size kept). */
export function filledPlaceholderAtTemplate<T extends Record<string, unknown>>(
  block: T,
  template: { left: number; top: number },
): T {
  return { ...block, left: template.left, top: template.top };
}

function isLayoutPlaceholderBlock(block: unknown): block is Record<string, unknown> {
  const record = block as Record<string, unknown> | null;
  if (!record || record.type !== "text") return false;
  if (typeof record.placeholder !== "string" || typeof record.role !== "string") return false;
  const source = record.source as Record<string, unknown> | undefined;
  return !(source && typeof source.slideId === "string");
}

/** Where a layout placeholder takes its frame, within one EMU (the rounding of two computations). */
function followsTemplate(
  block: Record<string, unknown>,
  template: { left: number; top: number; width: number; height: number },
): boolean {
  const rect = blockRect(block);
  return (
    rect !== null &&
    Math.abs(rect.left - template.left) <= 1 &&
    Math.abs(rect.top - template.top) <= 1 &&
    Math.abs(rect.width - template.width) <= 1 &&
    Math.abs(rect.height - template.height) <= 1
  );
}

/**
 * The slides after their layout frames change source (a theme or variant
 * applied): every placeholder block that still sits on the OLD source's frame
 * moves to the new source's frame for the same slot, and every text value it
 * still holds as the old layout stated it takes the new layout's. A block the
 * author moved, resized or formatted keeps what the author did; blocks that
 * are not layout placeholders, `blank` slides and imported (native) slides
 * are untouched. Pure; the same array comes back when nothing changes.
 */
export function reflowPlaceholders<
  S extends {
    layout?: unknown;
    blocks?: unknown;
    widthEmu?: number | null;
    heightEmu?: number | null;
    source?: unknown;
  },
>(slides: readonly S[], before: LayoutFrameSource, after: LayoutFrameSource): S[] {
  if (layoutFrameSourcesEqual(before, after)) return slides as S[];
  let changed = false;
  const next = slides.map((slide) => {
    // A custom layout's slides reflow too (PPT-B030).
    const layoutId = normalizeSlideLayoutRef(slide.layout);
    const slideSource = slide.source as Record<string, unknown> | undefined;
    if (layoutId === DEFAULT_LAYOUT_ID || !Array.isArray(slide.blocks)) return slide;
    if (slideSource && typeof slideSource.slideId === "string") return slide;
    const bounds = { width: slide.widthEmu, height: slide.heightEmu };
    const oldTemplates = layoutBlocks(layoutId, bounds, before);
    const newTemplates = layoutBlocks(layoutId, bounds, after);
    let slot = 0;
    let slideChanged = false;
    const blocks = (slide.blocks as unknown[]).map((block) => {
      if (!isLayoutPlaceholderBlock(block)) return block;
      const index = slot;
      slot += 1;
      const oldTemplate = oldTemplates[index];
      const newTemplate = newTemplates[index];
      if (!oldTemplate || !newTemplate) return block;
      const result: Record<string, unknown> = { ...block };
      if (followsTemplate(block, oldTemplate)) {
        result.left = newTemplate.left;
        result.top = newTemplate.top;
        result.width = newTemplate.width;
        result.height = newTemplate.height;
      }
      const style = layoutTextStyleAfterTemplate(
        block.style as Record<string, unknown> | undefined,
        oldTemplate.style as Record<string, unknown> | undefined,
        newTemplate.style as Record<string, unknown> | undefined,
      );
      if (style) result.style = style;
      else delete result.style;
      if (JSON.stringify(result) === JSON.stringify(block)) return block;
      slideChanged = true;
      return result;
    });
    if (!slideChanged) return slide;
    changed = true;
    return { ...slide, blocks };
  });
  return changed ? next : (slides as S[]);
}

function nativePlaceholderRole(type: string): {
  role: PlaceholderRole;
  placeholder: string;
} | null {
  if (type === "title" || type === "ctrTitle") {
    return { role: "title", placeholder: "title" };
  }
  if (type === "subTitle") {
    return { role: "subtitle", placeholder: "subtitle" };
  }
  if (type === "body") return { role: "body", placeholder: "content" };
  if (type === "pic") return { role: "picture", placeholder: "picture" };
  if (["obj", "chart", "tbl", "clipArt", "dgm", "media", "sldImg"].includes(type)) {
    return { role: "content", placeholder: "content" };
  }
  return null;
}










export function nativeLayoutCanApply(
  receipt: NativeLayoutCascadeReceipt,
  target: NativeLayoutCatalogEntryReceipt,
): boolean {
  return receipt.layouts.some((entry) => entry.layoutPart === target.layoutPart);
}

function nativeSlideInsertionLayoutTarget(
  slide: NativeSlideInsertionHost,
  layoutValue: unknown,
  kind: NativeSlideInsertionReceipt["kind"],
): {
  receipt: NativeLayoutCascadeReceipt;
  target: NativeLayoutCatalogEntryReceipt;
} | null {
  const source =
    slide.source &&
    typeof slide.source === "object" &&
    !Array.isArray(slide.source)
      ? (slide.source as Record<string, unknown>)
      : null;
  const receipt = parseNativeLayoutCascadeReceipt(source?.layoutCascade);
  if (
    !receipt ||
    source?.slideId !== receipt.slideId ||
    !NATIVE_SLIDE_ID.test(receipt.slideId)
  ) {
    return null;
  }
  const target =
    kind === "new"
      ? nativeLayoutCatalogEntryForValue(receipt, layoutValue)
      : (receipt.layouts.find(
          (entry) =>
            entry.layoutPart ===
            (typeof slide.layoutId === "string"
              ? slide.layoutId
              : receipt.layoutPart),
        ) ?? null);
  if (
    !target ||
    target.masterPart !== receipt.masterPart ||
    target.masterRelType !== receipt.masterRelType ||
    target.placeholders.some(
      (placeholder) =>
        !NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES.has(
          placeholder.placeholderType,
        ) &&
        (!nativePlaceholderRole(placeholder.placeholderType) ||
          !nativeLayoutGeometryIsComplete(placeholder.layoutGeometry)),
    )
  ) {
    return null;
  }
  const placeholderKeys = target.placeholders.map(nativePlaceholderKey);
  if (new Set(placeholderKeys).size !== placeholderKeys.length) return null;
  return { receipt, target };
}

/** Whether New Slide can use this exact layout inside the anchor's native master. */
export function nativeSlideCanInsertLayout(
  slide: NativeSlideInsertionHost,
  layoutValue: unknown,
): boolean {
  return nativeSlideInsertionLayoutTarget(slide, layoutValue, "new") !== null;
}

/** Whether a native source slide can be cloned without foreign package graphs. */
export function nativeSlideCanDuplicate(
  slide: NativeSlideInsertionHost,
): boolean {
  if (
    (typeof slide.notes === "string" && slide.notes.trim() !== "") ||
    (slide.comments !== undefined &&
      (!Array.isArray(slide.comments) || slide.comments.length > 0)) ||
    !Array.isArray(slide.blocks) ||
    slide.blocks.some(
      (block) =>
        !block ||
        !["text", "image", "shape", "table"].includes(String(block.type ?? "")),
    )
  ) {
    return false;
  }
  const resolved = nativeSlideInsertionLayoutTarget(
    slide,
    undefined,
    "duplicate",
  );
  return resolved !== null;
}

/** Build the exact persisted intent the API reauthenticates against the source. */
export function createNativeSlideInsertionReceipt(
  slide: NativeSlideInsertionHost,
  modelSlideId: string,
  kind: NativeSlideInsertionReceipt["kind"],
  layoutValue?: unknown,
): NativeSlideInsertionReceipt | null {
  const normalizedModelId = modelSlideId.trim();
  if (
    !normalizedModelId ||
    normalizedModelId.length > 256 ||
    (kind === "duplicate" && !nativeSlideCanDuplicate(slide))
  ) {
    return null;
  }
  const resolved = nativeSlideInsertionLayoutTarget(slide, layoutValue, kind);
  if (!resolved) return null;
  const { receipt, target } = resolved;
  return {
    schemaVersion: PPTX_NATIVE_SLIDE_INSERTION_RECEIPT_VERSION,
    kind,
    modelSlideId: normalizedModelId,
    packageFingerprint: receipt.packageFingerprint,
    presentationPart: receipt.presentationPart,
    anchorSlideId: receipt.slideId,
    anchorSlidePart: receipt.slidePart,
    layoutPart: target.layoutPart,
    layoutFingerprint: target.layoutFingerprint,
    masterPart: target.masterPart,
    masterFingerprint: receipt.masterFingerprint,
  };
}

/**
 * The deck's masters catalog as New Slide reads it: each imported master's
 * layouts and the blocks the importer projected from each layout's own
 * placeholders (their prompt rows, `paragraphs`, carry the list marker every
 * level inherits).
 */
export type NativeLayoutCatalogMasters = readonly unknown[] | null | undefined;

function catalogRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/** The projected blocks of one native layout of one master, or null. */
function nativeLayoutCatalogBlocks(
  masters: NativeLayoutCatalogMasters,
  masterPart: string,
  layoutPart: string,
): Record<string, unknown>[] | null {
  if (!Array.isArray(masters)) return null;
  for (const value of masters) {
    const master = catalogRecord(value);
    if (!master || (master.id !== masterPart && master.sourcePart !== masterPart)) continue;
    if (!Array.isArray(master.layouts)) return null;
    for (const entry of master.layouts) {
      const layout = catalogRecord(entry);
      if (!layout || (layout.id !== layoutPart && layout.sourcePart !== layoutPart)) continue;
      return Array.isArray(layout.blocks)
        ? layout.blocks.flatMap((block) => {
            const record = catalogRecord(block);
            return record ? [record] : [];
          })
        : null;
    }
    return null;
  }
  return null;
}











export function nativeLayoutInheritedBullet(
  placeholder: Pick<NativeLayoutPlaceholderReceipt, "placeholderType" | "layoutGeometry" | "layoutStyle">,
  catalogBlocks: readonly Record<string, unknown>[] | null,
): "none" | "bullet" | "number" {
  const stated = nativeLayoutParagraphBullet(placeholder.layoutStyle);
  if (stated !== "none" || !catalogBlocks) return stated;
  const geometry = placeholder.layoutGeometry;
  const markers = new Set(
    catalogBlocks
      .filter(
        (block) =>
          block.placeholderType === placeholder.placeholderType &&
          (["left", "top", "width", "height"] as const).every(
            (key) => block[key] === geometry[key],
          ),
      )
      .map((block) => {
        const first = Array.isArray(block.paragraphs)
          ? catalogRecord(block.paragraphs[0])
          : null;
        return first && (first.level ?? 0) === 0 &&
          (first.bullet === "bullet" || first.bullet === "number")
          ? first.bullet
          : "none";
      }),
  );
  return markers.size === 1 ? ([...markers][0] as "none" | "bullet" | "number") : "none";
}

/** Immediate editor projection for an authenticated native New Slide layout. */
export function nativeLayoutInsertionBlocks(
  slide: NativeSlideInsertionHost,
  layoutValue: unknown,
  nextId: () => string,
  masters?: NativeLayoutCatalogMasters,
): LayoutBlock[] | null {
  const resolved = nativeSlideInsertionLayoutTarget(slide, layoutValue, "new");
  if (!resolved) return null;
  const catalogBlocks = nativeLayoutCatalogBlocks(
    masters,
    resolved.target.masterPart,
    resolved.target.layoutPart,
  );
  const blocks: LayoutBlock[] = [];
  for (const placeholder of resolved.target.placeholders) {
    if (NATIVE_LAYOUT_BAND_PLACEHOLDER_TYPES.has(placeholder.placeholderType)) {
      continue;
    }
    const role = nativePlaceholderRole(placeholder.placeholderType);
    if (!role || !nativeLayoutGeometryIsComplete(placeholder.layoutGeometry)) {
      return null;
    }
    const style = Object.fromEntries(
      NATIVE_LAYOUT_NESTED_STYLE_KEYS.filter((key) =>
        Object.hasOwn(placeholder.layoutStyle, key),
      ).map((key) => [key, placeholder.layoutStyle[key]]),
    );
    const directStyle = Object.fromEntries(
      NATIVE_LAYOUT_TOP_LEVEL_STYLE_KEYS.filter((key) =>
        Object.hasOwn(placeholder.layoutStyle, key),
      ).map((key) => [key, placeholder.layoutStyle[key]]),
    );
    const bullet = nativeLayoutInheritedBullet(placeholder, catalogBlocks);
    blocks.push({
      id: nextId(),
      type: "text",
      kind: "textbox",
      text: "",
      placeholder: role.placeholder,
      role: role.role,
      left: placeholder.layoutGeometry.left,
      top: placeholder.layoutGeometry.top,
      width: placeholder.layoutGeometry.width,
      height: placeholder.layoutGeometry.height,
      // The layout's turn too: a slide's placeholder takes its layout's whole
      // `a:xfrm`, `rot` included, as PowerPoint draws a New Slide. Left out,
      // a template's tilted title and body stood upright on the new slide and
      // the save, finding the box no longer the layout's, declined the slide
      
      // slideLayout2's title at -75 degrees, its body at 15).
      ...(placeholder.layoutGeometry.rotation
        ? { rotation: placeholder.layoutGeometry.rotation }
        : {}),
      ...(Object.keys(style).length > 0 ? { style } : {}),
      ...directStyle,
      // The empty prompt row in the marker it inherits, as the importer
      // writes a reopened slide's empty placeholder: a Two Content body
      
      ...(bullet !== "none" ? { paragraphs: [{ text: "", level: 0, bullet }] } : {}),
      // The layout states its sizes in points, as every imported block
      // does: drawn as design pixels, a new slide's 44 pt title was 44 px,
      
      source: { textFontSizeUnit: "point" },
    });
  }
  return blocks;
}

/** Clamp a fraction to [min, max], collapsing non-finite values to `min`. */
function clampFraction(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  const hi = Math.max(min, max);
  return Math.min(Math.max(value, min), hi);
}

/**
 * The per-block FORMATTING keys a "Reset Slide" clears so a placeholder snaps
 * back to the layout's DEFAULT look (PowerPoint / WPS "Reset" parity). The
 * block's identity (`id`/`type`/`kind`), its placeholder `text`, and the
 * layout-owned geometry/`role`/`placeholder` are NOT in this list — only the
 * author-applied overrides (per-character runs, list structure, fill/outline,
 * shadow, animation, alignment, anchor, …) are stripped.
 */
const RESET_CLEARED_BLOCK_KEYS = [
  "style",
  "runs",
  "paragraphs",
  "bulletChar",
  "numberFormat",
  "fill",
  "fillTo",
  "gradientAngle",
  "outlineColor",
  "outlineWidth",
  "outlineStyle",
  "shadow",
  "animation",
  "align",
  "verticalAnchor",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
] as const;

/**
 * The block's runs with only what is content kept: a link and its ScreenTip,
 * a field, an equation. Reset and Clear Formatting take the look off text,
 * not its hyperlinks (PPT-B413; a link was wiped with the runs before), and a
 * text box's own link (`href`) is no formatting either. `undefined` when
 * nothing but plain text is left.
 */
function contentOnlyRuns(value: unknown): TextRun[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const runs = (value as TextRun[]).map((run) => {
    const href = run?.style?.href;
    const linkTooltip = run?.style?.linkTooltip;
    const next: TextRun = href
      ? { text: run.text, style: { href, ...(linkTooltip ? { linkTooltip } : {}) } }
      : { text: run?.text };
    if (run?.field) next.field = run.field;
    if (run?.math) next.math = run.math;
    return next;
  });
  return normalizeRuns(runs);
}

function withoutAuthoredFormatting<T extends Record<string, unknown>>(block: T): T {
  const next = { ...block } as Record<string, unknown>;
  for (const key of RESET_CLEARED_BLOCK_KEYS) {
    delete next[key];
  }
  const runs = contentOnlyRuns(block.runs);
  if (runs) next.runs = runs;
  return next as T;
}

/**
 * Strip every author-applied FORMATTING override off one block — the per-block
 * form of "Reset", reused by the Home ribbon's Clear Formatting (PowerPoint's
 * font-group eraser). Shares {@link RESET_CLEARED_BLOCK_KEYS} with the slide
 * reset so the two commands can never disagree about what "formatting" means.
 * Pure: returns a NEW object; content (`text`), identity, and geometry stay.
 */
export function clearAuthoredBlockFormatting<
  T extends Record<string, unknown>,
>(block: T): T {
  return withoutAuthoredFormatting(block);
}

/** A text box the author drew (no layout placeholder: it names no role). */
function isFreeTextBox(block: Record<string, unknown>): boolean {
  return (
    block.kind === "textbox" &&
    typeof block.placeholder !== "string" &&
    typeof block.role !== "string"
  );
}

/** A presentation slide as far as the reset helper needs to reason about it. */
export interface ResettableSlide {
  layout?: unknown;
  blocks?: unknown;
  widthEmu?: number | null;
  heightEmu?: number | null;
  [key: string]: unknown;
}

function nativeLayoutStyleValueEquals(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (
    !left ||
    !right ||
    typeof left !== "object" ||
    typeof right !== "object" ||
    Array.isArray(left) ||
    Array.isArray(right)
  ) {
    return false;
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord).sort();
  const rightKeys = Object.keys(rightRecord).sort();
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key, index) =>
        key === rightKeys[index] &&
        nativeLayoutStyleValueEquals(leftRecord[key], rightRecord[key]),
    )
  );
}

function nativeLayoutLiveValueInherits(
  live: Record<string, unknown>,
  original: Record<string, unknown>,
  key: string,
): boolean {
  const liveHas = Object.hasOwn(live, key);
  const originalHas = Object.hasOwn(original, key);
  return (
    (!liveHas && !originalHas) ||
    (liveHas &&
      originalHas &&
      nativeLayoutStyleValueEquals(live[key], original[key]))
  );
}

function nativeLayoutTextDetailRecords(
  block: Record<string, unknown>,
): Record<string, unknown>[] {
  const details: Record<string, unknown>[] = [];
  const collectRunStyles = (value: unknown): void => {
    if (!Array.isArray(value)) return;
    for (const run of value) {
      if (!run || typeof run !== "object" || Array.isArray(run)) continue;
      const style = (run as Record<string, unknown>).style;
      if (style && typeof style === "object" && !Array.isArray(style)) {
        details.push(style as Record<string, unknown>);
      }
    }
  };
  collectRunStyles(block.runs);
  if (Array.isArray(block.paragraphs)) {
    for (const value of block.paragraphs) {
      if (!value || typeof value !== "object" || Array.isArray(value)) continue;
      const paragraph = value as Record<string, unknown>;
      details.push(paragraph);
      collectRunStyles(paragraph.runs);
    }
  }
  return details;
}

function nativeLayoutNestedValueLocation(
  block: Record<string, unknown>,
  liveStyle: Record<string, unknown>,
  original: Record<string, unknown>,
  key: string,
): "block" | "details" | "absent" | "authored" {
  if (Object.hasOwn(liveStyle, key)) {
    return Object.hasOwn(original, key) &&
      nativeLayoutStyleValueEquals(liveStyle[key], original[key])
      ? "block"
      : "authored";
  }
  if (!Object.hasOwn(original, key)) return "absent";
  return nativeLayoutTextDetailRecords(block).some(
    (detail) =>
      Object.hasOwn(detail, key) &&
      nativeLayoutStyleValueEquals(detail[key], original[key]),
  )
    ? "details"
    : "authored";
}

function reconcileInheritedRunStyle(
  value: unknown,
  original: Record<string, unknown>,
): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) return value;
  const run = value as Record<string, unknown>;
  const rawStyle = run.style;
  if (!rawStyle || typeof rawStyle !== "object" || Array.isArray(rawStyle)) {
    return value;
  }
  const style = { ...(rawStyle as Record<string, unknown>) };
  let changed = false;
  for (const key of NATIVE_LAYOUT_NESTED_STYLE_KEYS) {
    if (
      Object.hasOwn(original, key) &&
      Object.hasOwn(style, key) &&
      nativeLayoutStyleValueEquals(style[key], original[key])
    ) {
      delete style[key];
      changed = true;
    }
  }
  if (!changed) return value;
  const next = { ...run };
  if (Object.keys(style).length > 0) next.style = style;
  else delete next.style;
  return next;
}

function nativeLayoutParagraphBullet(
  style: Record<string, unknown>,
): "none" | "bullet" | "number" {
  if (Object.hasOwn(style, "numberFormat")) return "number";
  if (Object.hasOwn(style, "bulletChar")) return "bullet";
  return "none";
}

function reconcileInheritedTextDetails(
  block: Record<string, unknown>,
  original: Record<string, unknown>,
  target: Record<string, unknown>,
  listStyleInherits: boolean,
): void {
  if (Array.isArray(block.runs)) {
    block.runs = block.runs.map((run) =>
      reconcileInheritedRunStyle(run, original),
    );
  }
  if (!Array.isArray(block.paragraphs)) return;
  block.paragraphs = block.paragraphs.map((value) => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return value;
    const paragraph = { ...(value as Record<string, unknown>) };
    if (
      listStyleInherits &&
      paragraph.bullet === nativeLayoutParagraphBullet(original)
    ) {
      paragraph.bullet = nativeLayoutParagraphBullet(target);
    }
    for (const key of [
      "align",
      "lineSpacing",
      "spaceBefore",
      "spaceAfter",
    ] as const) {
      if (
        Object.hasOwn(original, key) &&
        Object.hasOwn(paragraph, key) &&
        nativeLayoutStyleValueEquals(paragraph[key], original[key])
      ) {
        delete paragraph[key];
      }
    }
    if (Array.isArray(paragraph.runs)) {
      paragraph.runs = paragraph.runs.map((run) =>
        reconcileInheritedRunStyle(run, original),
      );
    }
    return paragraph;
  });
}

/** Apply one authenticated catalog layout without re-pouring freeform text. */
export function applyNativeLayoutToSlide(
  slide: ResettableSlide,
  receipt: NativeLayoutCascadeReceipt,
  target: NativeLayoutCatalogEntryReceipt,
): ResettableSlide {
  // A custom layout (no built-in type) reads as what its placeholders make it.
  const layout = nativeLayoutModelLayoutId(target);
  const currentLayoutPart =
    typeof slide.layoutId === "string" ? slide.layoutId : receipt.layoutPart;
  if (currentLayoutPart === target.layoutPart) return slide;
  const currentLayout = receipt.layouts.find(
    (entry) => entry.layoutPart === currentLayoutPart,
  );
  if (
    !currentLayout ||
    !receipt.layouts.some(
      (entry) =>
        entry.layoutPart === target.layoutPart &&
        entry.layoutFingerprint === target.layoutFingerprint &&
        entry.masterPart === target.masterPart &&
        entry.masterRelType === target.masterRelType,
    )
  ) {
    return slide;
  }
  if (!nativeLayoutCanApply(receipt, target)) return slide;
  const originalByShapeId = new Map(
    receipt.placeholders.map((placeholder) => [placeholder.shapeId, placeholder]),
  );
  const blocks = Array.isArray(slide.blocks)
    ? (slide.blocks as Record<string, unknown>[])
    : [];
  // A placeholder with no place on the new layout stays where it is (it is
  // left out of both maps below and its block is returned as it was).
  const targetForSource = nativeLayoutPlaceholderMapping(receipt, target);
  const currentForSource = new Map<string, NativeLayoutPlaceholderReceipt>();
  for (const original of receipt.placeholders) {
    const currentPlaceholder = nativeLayoutTargetForPlaceholder(
      currentLayout,
      original,
    );
    if (currentPlaceholder) currentForSource.set(original.shapeId, currentPlaceholder);
  }
  const nextBlocks = blocks.map((block) => {
    const source =
      block.source && typeof block.source === "object" && !Array.isArray(block.source)
        ? (block.source as Record<string, unknown>)
        : null;
    const original =
      typeof source?.shapeId === "string"
        ? originalByShapeId.get(source.shapeId)
        : undefined;
    if (
      !original ||
      source?.placeholderType !== original.placeholderType ||
      source.placeholderIndex !== original.placeholderIndex ||
      source.layoutPart !== receipt.layoutPart ||
      source.masterPart !== receipt.masterPart
    ) {
      return block;
    }
    const baseline = targetForSource.get(original.shapeId);
    const currentBaseline = currentForSource.get(original.shapeId);
    if (!baseline?.layoutGeometry || !currentBaseline) return block;
    const next = { ...block };
    // A frame the author set on the slide (`localXfrm`) stays: PowerPoint's
    // Layout keeps it, and Reset is what takes the layout's again.
    if (!original.localXfrm) {
      for (const key of NATIVE_LAYOUT_GEOMETRY_KEYS) delete next[key];
      Object.assign(next, baseline.layoutGeometry);
    }
    if (next.type === "text") {
      const originalLayoutStyle = currentBaseline.layoutStyle;
      const targetLayoutStyle = baseline.layoutStyle;
      const currentStyle =
        next.style && typeof next.style === "object" && !Array.isArray(next.style)
          ? (next.style as Record<string, unknown>)
          : {};
      const style = { ...currentStyle };
      const listStyleInherits = ["bulletChar", "numberFormat"].every((key) =>
        nativeLayoutLiveValueInherits(next, originalLayoutStyle, key),
      );
      for (const key of NATIVE_LAYOUT_NESTED_STYLE_KEYS) {
        const location = nativeLayoutNestedValueLocation(
          next,
          style,
          originalLayoutStyle,
          key,
        );
        if (location === "authored") {
          continue;
        }
        if (Object.hasOwn(targetLayoutStyle, key)) style[key] = targetLayoutStyle[key];
        else delete style[key];
      }
      if (Object.keys(style).length > 0) next.style = style;
      else delete next.style;

      for (const key of NATIVE_LAYOUT_TOP_LEVEL_STYLE_KEYS) {
        if (!nativeLayoutLiveValueInherits(next, originalLayoutStyle, key)) {
          continue;
        }
        if (Object.hasOwn(targetLayoutStyle, key)) next[key] = targetLayoutStyle[key];
        else delete next[key];
      }
      reconcileInheritedTextDetails(
        next,
        originalLayoutStyle,
        targetLayoutStyle,
        listStyleInherits,
      );
    }
    return next;
  });
  const next: ResettableSlide = {
    ...slide,
    blocks: nextBlocks,
    layoutId: target.layoutPart,
    masterId: target.masterPart,
  };
  if (layout === DEFAULT_LAYOUT_ID) delete next.layout;
  else next.layout = layout;
  return next;
}

function nativeResetRevision(slide: ResettableSlide): number | null {
  const source =
    slide.source && typeof slide.source === "object" && !Array.isArray(slide.source)
      ? (slide.source as Record<string, unknown>)
      : null;
  const value = source?.layoutResetRevision ?? 0;
  return typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 0 &&
    value <= 2_147_483_647
    ? value
    : null;
}

function nativeResetCurrentLayout(
  slide: ResettableSlide,
  receipt: NativeLayoutCascadeReceipt,
): NativeLayoutCatalogEntryReceipt | null {
  const layoutPart =
    typeof slide.layoutId === "string" ? slide.layoutId : receipt.layoutPart;
  const candidates = receipt.layouts.filter(
    (entry) =>
      entry.layoutPart === layoutPart &&
      entry.masterPart === receipt.masterPart,
  );
  return candidates.length === 1 ? candidates[0] : null;
}

/** Whether Reset has complete authenticated native custody and a real delta. */
export function nativeResetCanApply(
  slide: ResettableSlide,
  receipt: NativeLayoutCascadeReceipt,
): boolean {
  const authority = receipt.resetAuthority;
  const revision = nativeResetRevision(slide);
  if (
    !authority ||
    authority.disposition !== "native" ||
    authority.blocker !== null ||
    !authority.placeholders.some((placeholder) => placeholder.overrides.length > 0) ||
    revision !== 0
  ) {
    return false;
  }
  const currentLayout = nativeResetCurrentLayout(slide, receipt);
  if (!currentLayout) return false;
  if (
    slide.masterId !== undefined &&
    slide.masterId !== currentLayout.masterPart
  ) {
    return false;
  }
  const originalByShapeId = new Map(
    receipt.placeholders.map((placeholder) => [placeholder.shapeId, placeholder]),
  );
  const seenProjectedShapeIds = new Set<string>();
  for (const value of Array.isArray(slide.blocks) ? slide.blocks : []) {
    if (!value || typeof value !== "object" || Array.isArray(value)) continue;
    const block = value as Record<string, unknown>;
    const source =
      block.source && typeof block.source === "object" && !Array.isArray(block.source)
        ? (block.source as Record<string, unknown>)
        : null;
    if (typeof source?.shapeId !== "string") continue;
    const original = originalByShapeId.get(source.shapeId);
    if (!original) continue;
    if (
      seenProjectedShapeIds.has(original.shapeId) ||
      source.placeholderType !== original.placeholderType ||
      source.placeholderIndex !== original.placeholderIndex ||
      source.layoutPart !== receipt.layoutPart ||
      source.masterPart !== receipt.masterPart
    ) {
      return false;
    }
    const baseline = nativeLayoutTargetForPlaceholder(currentLayout, original);
    if (!baseline || !nativeLayoutGeometryIsComplete(baseline.layoutGeometry)) {
      return false;
    }
    seenProjectedShapeIds.add(original.shapeId);
  }
  return true;
}

function nativeResetTextDetails(
  block: Record<string, unknown>,
  target: Record<string, unknown>,
): void {
  const resetRun = (value: unknown): unknown => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return value;
    const run = { ...(value as Record<string, unknown>) };
    const rawStyle = run.style;
    if (rawStyle && typeof rawStyle === "object" && !Array.isArray(rawStyle)) {
      const style = { ...(rawStyle as Record<string, unknown>) };
      for (const key of NATIVE_LAYOUT_NESTED_STYLE_KEYS) delete style[key];
      if (Object.keys(style).length > 0) run.style = style;
      else delete run.style;
    }
    return run;
  };
  if (Array.isArray(block.runs)) block.runs = block.runs.map(resetRun);
  // `a:pPr@lvl` is a slide-local selector.  The authenticated Reset writer
  // removes the complete direct pPr, so the reopened paragraph starts again
  // at DrawingML's default inheritance level (0).  `nativeListLevels` is the
  // immutable projection of levels 5..8 and `nativeBullet` the marker the
  // package drew; retaining either after that removal would make the immediate
  // canvas disagree with the reopened package.
  dropNativeParagraphReceipts(block);
  if (!Array.isArray(block.paragraphs)) return;
  block.paragraphs = block.paragraphs.map((value) => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return value;
    const paragraph = { ...(value as Record<string, unknown>) };
    for (const key of [
      "align",
      "lineSpacing",
      "spaceBefore",
      "spaceAfter",
      "bulletChar",
      "numberFormat",
      // `marL`/`indent` are `a:pPr` attributes too, and the Reset writer takes
      // the whole element: a row that kept them would claim an indentation the
      // reopened package no longer states.
      "marL",
      "indent",
    ]) {
      delete paragraph[key];
    }
    paragraph.level = 0;
    paragraph.bullet = nativeLayoutParagraphBullet(target);
    if (Array.isArray(paragraph.runs)) {
      paragraph.runs = paragraph.runs.map(resetRun);
    }
    return paragraph;
  });
  // A Reset can empty the rows of everything paragraph-local, and the importer
  // does not write a `paragraphs` record for rows that say nothing: a plain
  // single paragraph with no list structure, spacing, alignment or indent has
  // its runs surfaced at BLOCK level instead. Keeping the emptied array here
  // would leave the canvas asserting `level: 0, bullet: "none"` for a paragraph
  // the reopened package answers with no record at all — a field written in a
  // spelling its own reader never produces, which is what made the reset fail
  // to be a fixed point.
  const reset = block.paragraphs as SlideParagraph[];
  if (!isFlatEquivalentParagraphs(reset)) return;
  const rowRuns = reset.map((paragraph) =>
    (paragraph as unknown as Record<string, unknown>).runs,
  );
  const anyRowRuns = rowRuns.some(
    (runs) => Array.isArray(runs) && runs.length > 0,
  );
  if (reset.length > 1 && anyRowRuns) return;
  if (
    reset.length === 1 &&
    anyRowRuns &&
    (!Array.isArray(block.runs) || block.runs.length === 0)
  ) {
    block.runs = rowRuns[0];
  }
  delete block.paragraphs;
}

/** Apply one authenticated native Reset projection without changing content. */
export function resetNativeSlideToLayout(
  slide: ResettableSlide,
  receipt: NativeLayoutCascadeReceipt,
): ResettableSlide {
  if (!nativeResetCanApply(slide, receipt)) return slide;
  const currentLayout = nativeResetCurrentLayout(slide, receipt);
  if (!currentLayout) return slide;
  const originalByShapeId = new Map(
    receipt.placeholders.map((placeholder) => [placeholder.shapeId, placeholder]),
  );
  // The layout as the deck now draws it: after a Slide Size, scaled with the
  // slides as the package's layouts are written.
  const resized = nativeSlideSizeLayoutProjection(slide);
  const blocks = (Array.isArray(slide.blocks) ? slide.blocks : []).map((value) => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return value;
    const block = value as Record<string, unknown>;
    const source =
      block.source && typeof block.source === "object" && !Array.isArray(block.source)
        ? (block.source as Record<string, unknown>)
        : null;
    const original =
      typeof source?.shapeId === "string"
        ? originalByShapeId.get(source.shapeId)
        : undefined;
    const baseline = original
      ? nativeLayoutTargetForPlaceholder(currentLayout, original)
      : null;
    if (!baseline) return value;
    // Reset clears what the save clears: a placeholder whose formatting the
    // API cannot clear exactly is reset to its layout's frame alone (its
    // receipt names only a `geometry` override), and one it names nothing
    // for stays as it is.
    const kinds = new Set(
      (receipt.resetAuthority?.placeholders.find((item) => item.shapeId === original!.shapeId)?.overrides ?? []).map(
        (override) => override.kind,
      ),
    );
    if (kinds.size === 0) return value;
    const next = { ...block };
    if (kinds.has("geometry")) {
      for (const key of NATIVE_LAYOUT_GEOMETRY_KEYS) delete next[key];
      Object.assign(next, resized ? resized.geometry(baseline.layoutGeometry) : baseline.layoutGeometry);
    }
    if (next.type === "text" && [...kinds].some((kind) => kind !== "geometry")) {
      const style =
        next.style && typeof next.style === "object" && !Array.isArray(next.style)
          ? { ...(next.style as Record<string, unknown>) }
          : {};
      for (const key of NATIVE_LAYOUT_NESTED_STYLE_KEYS) delete style[key];
      for (const key of NATIVE_LAYOUT_NESTED_STYLE_KEYS) {
        if (Object.hasOwn(baseline.layoutStyle, key)) {
          const value = baseline.layoutStyle[key];
          style[key] =
            key === "fontSize" && resized && typeof value === "number" && Number.isFinite(value)
              ? resized.scaleFont(value)
              : value;
        }
      }
      if (Object.keys(style).length > 0) next.style = style;
      else delete next.style;
      for (const key of NATIVE_LAYOUT_TOP_LEVEL_STYLE_KEYS) delete next[key];
      for (const key of NATIVE_LAYOUT_TOP_LEVEL_STYLE_KEYS) {
        if (Object.hasOwn(baseline.layoutStyle, key)) {
          next[key] = baseline.layoutStyle[key];
        }
      }
      nativeResetTextDetails(next, baseline.layoutStyle);
    }
    return next;
  });
  const source = {
    ...((slide.source &&
    typeof slide.source === "object" &&
    !Array.isArray(slide.source)
      ? slide.source
      : {}) as Record<string, unknown>),
    layoutResetRevision: 1,
  };
  return { ...slide, blocks, source };
}

/**
 * Restore a slide's PLACEHOLDERS to its layout default — the "Reset Slide"
 * affordance (PowerPoint / WPS). Pure + side-effect-free: returns a NEW slide
 * (the input is never mutated) so the caller can commit it as one undoable step.
 *
 * Behaviour (mirrors PowerPoint "Reset", which only touches placeholders):
 *  - The slide's existing TEXT content is re-poured, in document order, into the
 *    layout's placeholder boxes (so wording is never lost on reset), exactly
 *    like {@link layoutBlocks} positions them. Each placed block is reset to the
 *    layout's default geometry/`role`/`placeholder` and has its author
 *    formatting overrides ({@link RESET_CLEARED_BLOCK_KEYS}) CLEARED.
 *  - Text beyond the placeholder count is preserved as plain reset text blocks
 *    (geometry/formatting stripped) so no content is dropped.
 *  - NON-text blocks (images / shapes / tables / charts …) are preserved
 *    verbatim — Reset does not delete media.
 *  - Slide-level properties (background, transition, notes, size, master, hidden)
 *    are NOT touched; Reset is a placeholder-only operation.
 *  - A `blank` layout (no placeholders) keeps every existing block but strips the
 *    text blocks' formatting overrides (a true "reset formatting").
 *
 * `bounds` is the resolved slide canvas the caller renders against (so reset
 * placeholders land where the live canvas expects); when omitted it falls back
 * to the slide's own `widthEmu`/`heightEmu` (and then the standard canvas).
 * `makeBlockId` lets the caller inject a deterministic id factory (tests); it
 * defaults to the same crypto-or-fallback generator the layout materializer uses.
 */
export function resetSlideToLayout(
  slide: ResettableSlide,
  bounds?: { width?: number | null; height?: number | null },
  makeBlockId: () => string = () => makeId(0),
  source: LayoutFrameSource = BUILTIN_LAYOUT_FRAMES,
): ResettableSlide {
  const resolvedBounds = bounds ?? {
    width: slide.widthEmu,
    height: slide.heightEmu,
  };
  const templates = layoutBlocks(slide.layout, resolvedBounds, source);
  const existing: Record<string, unknown>[] = Array.isArray(slide.blocks)
    ? (slide.blocks as Record<string, unknown>[])
    : [];
  // A free text box (Insert > Text Box, WordArt) is no placeholder: Reset
  // leaves it as it is, as PowerPoint's touches placeholders only. It took a
  
  const textBlocks = existing.filter(
    (b) => (b?.type === "text" && !isFreeTextBox(b)) || isFilledPlaceholderBlock(b),
  );
  const stripFormatting = withoutAuthoredFormatting;

  // Pour existing text (in order) into the layout placeholders, resetting each
  // to default geometry + role + placeholder and clearing author formatting.
  const placed = templates.map((template, index) => {
    const source = textBlocks[index];
    if (source && isFilledPlaceholderBlock(source)) return filledPlaceholderAtTemplate(source, template);
    const text = typeof source?.text === "string" ? source.text : "";
    const runs = contentOnlyRuns(source?.runs);
    const href = source?.href;
    const linkTooltip = source?.linkTooltip;
    return {
      ...template,
      id: typeof source?.id === "string" ? source.id : makeBlockId(),
      text,
      ...(runs ? { runs } : {}),
      ...(typeof href === "string" ? { href } : {}),
      ...(typeof href === "string" && typeof linkTooltip === "string" ? { linkTooltip } : {}),
    };
  });
  // Text beyond the placeholder count: keep the content but strip formatting.
  const leftoverText = textBlocks
    .slice(templates.length)
    .map((block) => (isFilledPlaceholderBlock(block) ? block : stripFormatting(block)));

  // Each text block is reset where it stands in the stacking order, as
  // Layout re-pours it (PPT-B412): a shape sent behind the title stays there.
  // A place the layout adds goes just in front of the last text block.
  const textIndex = new Map(textBlocks.map((block, index) => [block, index]));
  const ordered: Record<string, unknown>[] = [];
  let lastText = -1;
  for (const block of existing) {
    const index = textIndex.get(block);
    if (index === undefined) {
      ordered.push(block);
      continue;
    }
    ordered.push(index < placed.length ? placed[index] : leftoverText[index - placed.length]);
    lastText = ordered.length - 1;
  }
  ordered.splice(lastText + 1, 0, ...placed.slice(textBlocks.length));
  const nextBlocks =
    templates.length > 0
      ? ordered
      : // `blank` layout: no placeholders — keep blocks, strip text formatting.
        existing.map((block) =>
          block?.type === "text" && !isFreeTextBox(block)
            ? stripFormatting(block)
            : block,
        );

  // A Blank slide with nothing on it stays empty: PowerPoint's Blank layout
  // has no placeholder to put back.
  return {
    ...slide,
    blocks: blocksLeftOnSlide(slide, nextBlocks, makeBlockId),
  };
}
