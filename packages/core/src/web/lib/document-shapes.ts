/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Native Docs shape contract.
 *
 * This is a thin Word adapter over the existing presentation geometry
 * renderer.  It deliberately owns a separate persisted schema: presentation
 * slide objects and flowing Word drawings do not share storage semantics, but
 * their preview/path geometry is identical and therefore remains in
 * `presentation-shapes` as the one rendering authority.
 */

import { normalizeOfficeCustomPathPaints, type OfficeCustomPathPaint } from "./office-custom-path-paints";
import type { Node as PmNode } from "@tiptap/pm/model";
import { pictureObjectTokenSha } from "./page-picture-objects";
import { SheetFact } from "./editor/sheet-facts";
import { normalizeOfficeTextWarp, type OfficeTextWarp } from "./office-text-warp";
import { WORD_LINE_MULTIPLE_MAX, WORD_LINE_MULTIPLE_MIN } from "./office-units";
import { OFFICE_OUTLINE_SHOWN_WEIGHT_PT, officeOutlineWeightWhenShown } from "./office-line-weight";
import {
  isPresentationConnectorShapeKind,
  normalizePresentationConnectorSpec,
  PRESENTATION_CONNECTOR_END_SIZES,
  PRESENTATION_CONNECTOR_ENDS,
  type PresentationConnectorEnd,
  type PresentationConnectorEndSize,
  type StrokeStyle,
} from "./presentation-shapes";
import { officeFramePlace } from "./office-object-frame";
// One custom-geometry model for every Office surface, rather than a Docs-only
// spelling of the same path.
import {
  DOCUMENT_CUSTOM_PATH_MAX_COMMANDS,
  normalizePresentationCustomPath,
  type PresentationCustomPath,
} from "./presentation-freeform";
import { normalizeOfficeTextPaint, patchOfficeTextPaint, type OfficeTextPaint } from "./office-text-paint";
import { DEFAULT_LINE_HEIGHT, estimateTextHeight } from "./presentation-autofit";
import { nearestAnchorPair } from "./presentation-connectors";
import {
  normalizeOfficeShapeFontFamily,
  OFFICE_SHAPE_MAX_FONT_FAMILY_LENGTH,
} from "./office-shape-text-font";
import {
  officeGeometryTakesLineEnds,
  officeShapeLineEndsPatch,
  readOfficeLineEnds,
  type OfficeLineEndsState,
} from "./office-line-ends";
import { presentationPresetGeometryForKind } from "./presentation-authored-preset-geometry";
import {normalizeOfficeShapeGradient,normalizeOfficeShapeShadow,type OfficeShapeGradient,type OfficeShapeShadow} from "./office-shape-paint";
import {
  normalizeOfficePictureEffects,
  OFFICE_PICTURE_EFFECT_KINDS,
  type OfficePictureEffectKind,
  type OfficePictureEffects,
} from "./office-picture-effects";
import { normalizeOfficeLineSketch, type OfficeLineSketch } from "./office-shape-effects-paint";
import {
  OFFICE_SWITCH_SHADOW,
  officeEffectsOfLegacyPreset,
  officeShadowFromShapeShadow,
} from "./office-shape-legacy-effects";
import { officeStrokeAlign, type OfficeStrokeAlign } from "./office-shape-line-align";
import {
  OFFICE_MENU_DASHES,
  officeMenuDashFields,
  officeStrokeDash,
  type OfficeMenuDash,
  type OfficeStrokeDash,
} from "./office-stroke-dash";
import { officeStrokeDetail, type OfficeStrokeDetail } from "./office-stroke-detail";
import {
  INK_EFFECT_MAX_CHARS,
  OFFICE_INK_EFFECTS,
  isInkEffect,
  normalizeInkEffect,
  type InkEffectId,
} from "./office-ink-effects";
// The SmartArt quick-style vocabulary is shared with Presentation and the
// worksheet (`lib/office-smartart/styles`); Docs stores the SAME payloads on a
// group member rather than declaring a second 3-D dialect. That module is pure
// and imports nothing from here, so the dependency stays one-way.
import {
  nativeSmartArtTextBlocker,
  type NativeSmartArtTextBlocker,
} from "./office-smartart/native-text-blocker";
// The one spelling a SmartArt picture is stored in, shared with the worksheet
// shape model and the native conversion. Pure, imports nothing.
import { smartArtPictureSrc } from "./office-smartart/picture";
import { smartArtColorVariant } from "./office-smartart/colors";
import { nativeSmartArtAliasOf } from "./office-smartart/native-layouts";
import { isSmartArtHierBranch, SMARTART_OFFICE_DEFINITION_KIND } from "./office-smartart/model";
import { normalizeOfficePictureFillOptions, type OfficePictureFillOptions } from "./office-picture-fill-options";
import { officeGroupFillMemberPaint, type OfficeGroupFillBox } from "./office-shape-group-fill";
import {
  readSmartArtCustom,
  type SmartArtCustomFactors,
} from "./office-smartart/custom-geometry";
import {
  readSmartArtTextBox,
  type SmartArtTextBox,
} from "./office-smartart/text-box";
import {
  isNativeScene3d,
  isNativeShape3d,
  isSmartArtStyleId,
  isSmartArtTextShape3d,
  type SmartArtScene3d,
  type SmartArtShape3d,
  type SmartArtStyleId,
  type SmartArtTextShape3d,
} from "./office-smartart/styles-model";
import { toggleFlip } from "./presentation-drag";
import type { TextRun } from "./presentation-runs";
import { normalizeDocumentShapeTextRuns } from "./document-shape-text-runs";
import type { CropRect } from "./presentation-crop-group";
import type { OfficePictureStyleId } from "./office-picture-styles/catalog";
import { smartArtSlotPictureFields, smartArtSlotPictureFormatOf } from "./office-smartart/slot-picture-format";
import {
  normalizeOfficeChartFillPattern,
  type OfficeChartFillPattern,
  officeFillPatternOf,
} from "./office-chart-shape-paint";
import {
  OFFICE_SHAPE_CATEGORIES,
  OFFICE_SHAPE_CATEGORY_IDS,
  OFFICE_SHAPE_DEFINITIONS,
  OFFICE_SHAPE_PRESETS,
  isOfficeShapePreset,
  officeShapeDefinition,
  officeShapeMatchesSearch,
  officeShapePresetFromDrawingMl,
  parseOfficeShapeRecents,
  pushOfficeShapeRecent,
  type OfficeShapeCategory,
  type OfficeShapeCategoryId,
  type OfficeShapeDefinition,
  type OfficeShapePreset,
} from "./office-shape-catalog";

export const DOCUMENT_SHAPE_SCHEMA_VERSION = 2 as const;
export const DOCUMENT_SHAPE_EMU_PER_PX = 9_525;
export const DOCUMENT_SHAPE_MIN_EDGE_EMU = 16 * DOCUMENT_SHAPE_EMU_PER_PX;
/**
 * The smallest edge a shape that encloses an area may STATE: Word keeps a
 * 0.13 in (118,745 EMU) rule as drawn, and a thinner one too. The 16 px above
 * is the size a new shape is drawn and grabbed at, not a bound on a document's
 * geometry; raising a thin rule to it on load changed the page, and a save
 * wrote the invented height into the file. One EMU, not zero: a shape with no
 * extent on an axis encloses nothing. The server twin is
 * `DOCUMENT_SHAPE_MIN_AREA_EDGE_EMU` (office_smartart_document_members.py).
 */
export const DOCUMENT_SHAPE_MIN_AREA_EDGE_EMU = 1;

/**
 * The minimum edge one preset is held to. One EMU for every shape that
 * encloses an area (`DOCUMENT_SHAPE_MIN_AREA_EDGE_EMU`); ZERO for a one-dimensional route — the straight line, its
 * arrowed spellings, every elbow and curve — because a horizontal line is as
 * tall as nothing. Word states that directly (DrawingML `a:ext cy="0"`, VML
 * two endpoints with the same `y`), and flooring it does not thicken the line,
 * it makes it DIAGONAL: a line preset is drawn corner to corner in its box.
 * The server twin is `UniversalDocumentShape.validate_extent_floor`
 * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §21).
 */
export function documentShapeMinEdgeEmu(preset: OfficeShapePreset): number {
  // Either spelling answers: the modern kind (`line`, `elbow-connector`) is
  // itself a connector kind, and a legacy preset (`lineArrow`) reaches the
  // same answer through the catalog's own `previewKind`.
  return isPresentationConnectorShapeKind(preset) ||
    isPresentationConnectorShapeKind(officeShapeDefinition(preset).previewKind)
    ? 0
    : DOCUMENT_SHAPE_MIN_AREA_EDGE_EMU;
}
export const DOCUMENT_SHAPE_MAX_EDGE_EMU = 5_000 * DOCUMENT_SHAPE_EMU_PER_PX;
export const DOCUMENT_SHAPE_DEFAULT_WIDTH_EMU = 220 * DOCUMENT_SHAPE_EMU_PER_PX;
export const DOCUMENT_SHAPE_DEFAULT_HEIGHT_EMU =
  140 * DOCUMENT_SHAPE_EMU_PER_PX;
export const DOCUMENT_SHAPE_DEFAULT_FILL = "#dbeafe";
export const DOCUMENT_SHAPE_DEFAULT_OUTLINE = "#2563eb";
export const DOCUMENT_SHAPE_DEFAULT_OUTLINE_WIDTH_PT = 1.5;
export const DOCUMENT_SHAPE_DEFAULT_TEXT_FILL = "#0f172a";
export const DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_FAMILY = "Arial";
export const DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_SIZE_PT = 14;
/** Word's largest font size, a shape's words' as any text's. */
export const DOCUMENT_SHAPE_MAX_TEXT_FONT_SIZE_PT = 1638;
export const DOCUMENT_SHAPE_MAX_TEXT_LENGTH = 32_000;
export const DOCUMENT_SHAPE_MAX_ALT_TEXT_LENGTH = 1_000;
export const DOCUMENT_SHAPE_MAX_FONT_FAMILY_LENGTH = OFFICE_SHAPE_MAX_FONT_FAMILY_LENGTH;
/**
 * How many shapes one native `wpg:wgp` group may hold.
 *
 * It is a capacity bound, not a fidelity one: a group is drawn, validated and
 * written member by member, so the only thing the number protects is how much
 * one atom can cost. Word's own diagrams routinely pass 32 — a picture-accent
 * process is four shapes per step, and the import corpus holds real diagrams of
 * 33, 44 and 45 members — and every one of those used to lose its group and
 * fall back to a preserved drawing or a flat picture. A drawn flow chart in a
 * canvas holds 79 (DOCX corpus R07-D03), so the bound is 256: still a bound on
 * one atom's cost, no longer a cap real drawings meet.
 *
 * The API mirrors this in `office_smartart_document_members.py` and enforces it
 * in the page validator; all of them read one number so they cannot drift.
 */
/** A group member's own story (band rich blocks), kept as the server wrote
 * it: an array of plain block records, bounded. The renderer reads it through
 * the band's own normalizer. */
export function documentShapeMemberTextStory(
  value: unknown,
): readonly Readonly<Record<string, unknown>>[] | null {
  if (!Array.isArray(value) || value.length === 0 || value.length > 200) return null;
  if (!value.every((block) => block && typeof block === "object" && !Array.isArray(block))) {
    return null;
  }
  return value as readonly Readonly<Record<string, unknown>>[];
}

export const DOCUMENT_SHAPE_MAX_GROUP_MEMBERS = 2048;
export const DOCUMENT_SHAPE_MAX_DRAWING_ID = 2_147_483_647;
/** DrawingML `wp:anchor/@relativeHeight` is an unsigned 32-bit integer. */
export const DOCUMENT_SHAPE_MAX_RELATIVE_HEIGHT = 0xffff_ffff;
/** Word's stable default for a newly-authored floating shape. */
export const DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT = 251_658_240;
/** One keyboard/menu nudge, expressed in the persisted DrawingML coordinate. */
export const DOCUMENT_SHAPE_POSITION_STEP_EMU = 8 * DOCUMENT_SHAPE_EMU_PER_PX;

export type DocumentShapePositionAction =
  | "left"
  | "right"
  | "up"
  | "down"
  | "reset";
export type DocumentShapeAlignAction = "left" | "center" | "right";
export type DocumentShapeRotationAction = -90 | 0 | 90;
export type DocumentShapeFlipAxis = "flipH" | "flipV";
export type DocumentShapeSizeAxis = "width" | "height";
export type DocumentShapeEffectPreset =
  | "none"
  | "shadow"
  | "reflection"
  | "glow"
  | "softEdge"
  | "bevel"
  | "threeD";
export type DocumentShapeTextDirection =
  | "horizontal"
  | "rotate90"
  | "rotate270"
  | "stacked";
export type DocumentShapeTextVerticalAlign = "top" | "middle" | "bottom";
/** `w:u` on the shape's text runs, as Home's Underline offers it. */
export type DocumentShapeTextUnderline = "none" | "single" | "double";
/** `w:jc` of the paragraphs inside the shape: the words' own alignment. */
export type DocumentShapeTextHorizontalAlign =
  | "left"
  | "center"
  | "right"
  | "justify";
/** `w:jc` of the paragraph an "In Line with Text" shape stands in alone. */
export type DocumentShapeParagraphAlign = "left" | "center" | "right";
export type DocumentShapeWrapMode =
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText"
  /**
   * "In Line with Text": a `wp:inline` drawing, which sits in its paragraph like
   * a glyph. It has no position and no `relativeHeight`, so it is never moved
   * and never takes part in the floating z-order stack. Only an imported shape
   * holds it; the writer keeps the inline host and never converts it.
   */
  | "inline";
export type DocumentShapeZOrderAction =
  | "bringToFront"
  | "bringForward"
  | "sendBackward"
  | "sendToBack";

export interface DocumentShapeStackEntry {
  readonly pos: number;
  readonly relativeHeight: number;
}

export interface DocumentShapeStackResolution {
  readonly complete: boolean;
  /** Lowest/backmost first. Equal source values use stable document order. */
  readonly backToFront: readonly DocumentShapeStackEntry[];
}

interface DocumentShapeStackNode {
  readonly type: Readonly<{ name: string }>;
  readonly attrs: Readonly<Record<string, unknown>>;
}

interface DocumentShapeStackDocument {
  descendants(
    callback: (node: DocumentShapeStackNode, pos: number) => boolean | void,
  ): void;
}

/** Strict UInt32 parser shared by model normalization and z-order admission. */
export function documentShapeRelativeHeight(value: unknown): number | null {
  return typeof value === "number" &&
    Number.isSafeInteger(value) &&
    value >= 0 &&
    value <= DOCUMENT_SHAPE_MAX_RELATIVE_HEIGHT
    ? value
    : null;
}

/**
 * The floating stack's next place: a new object (a drawn shape, a stroke, a
 * picture just floated) paints over everything the document already states.
 */
export function nextDocumentFloatingHeight(doc: DocumentShapeStackDocument): number {
  const heights = documentFloatingStack(doc).heights;
  let top = DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT - 1;
  if (heights.length) {
    top = -Infinity;
    for (const height of heights) if (height > top) top = height;
  }
  return Math.min(
    DOCUMENT_SHAPE_MAX_RELATIVE_HEIGHT,
    Math.max(top + 1, DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT),
  );
}

/**
 * The place a picture floated IN THE EDITOR holds in the floating stack: its
 * own `relativeHeight` (given when it was floated), as a shape holds its own.
 * An imported picture keeps its `wp:anchor` reading instead (`floatAnchor`),
 * and a picture from before heights were given states none: both answer null.
 */
export function documentPictureStackHeight(
  attrs: Readonly<Record<string, unknown>>,
): number | null {
  if (attrs.floatAnchor && typeof attrs.floatAnchor === "object") return null;
  if (typeof attrs.dataFloat !== "string" || attrs.dataFloat.length === 0) return null;
  return documentShapeRelativeHeight(attrs.relativeHeight);
}

/**
 * The attributes a z-order patch writes on a member of the stack: a shape's
 * or an editor-floated picture's own `relativeHeight`. Null for a node that
 * is not a renumberable member (the stack was then not complete).
 */
export function documentStackPatchAttrs(
  node: DocumentShapeStackNode,
  relativeHeight: number,
): Record<string, unknown> | null {
  const name = node.type.name;
  if (name === "documentShape") return { ...node.attrs, relativeHeight };
  if (name === "image" && documentPictureStackHeight(node.attrs) !== null) {
    return { ...node.attrs, relativeHeight };
  }
  // An imported anchored picture or a floating text box frame holds its
  // place in its anchor reading, which the page paints by and the save
  // writes back into the file's `wp:anchor` (`_docx_floating_stack_heights`).
  const anchor = documentAnchoredStackAnchor(node);
  if (anchor) return { ...node.attrs, floatAnchor: { ...anchor, relativeHeight } };
  // A floating chart: its own anchor, which the chart writer writes whole.
  const chartAnchor = name === "chart" ? node.attrs.floatAnchor : null;
  if (
    chartAnchor &&
    typeof chartAnchor === "object" &&
    documentShapeRelativeHeight((chartAnchor as { relativeHeight?: unknown }).relativeHeight) !== null
  ) {
    return { ...node.attrs, floatAnchor: { ...(chartAnchor as Record<string, unknown>), relativeHeight } };
  }
  return null;
}

/**
 * The anchor reading of an imported member of the floating stack the page
 * may reorder: an anchored picture or a floating text box frame whose
 * import found the page showing the whole stack (`zOrderStackComplete`).
 */
function documentAnchoredStackAnchor(
  node: DocumentShapeStackNode,
): Record<string, unknown> | null {
  const name = node.type.name;
  if (name !== "image" && name !== "textBox") return null;
  const anchor = node.attrs.floatAnchor;
  if (!anchor || typeof anchor !== "object" || node.attrs.zOrderStackComplete !== true) {
    return null;
  }
  return documentShapeRelativeHeight((anchor as { relativeHeight?: unknown }).relativeHeight) === null
    ? null
    : (anchor as Record<string, unknown>);
}

/**
 * Every floating object's stated `relativeHeight`, for the WHOLE document.
 *
 * Word keeps ONE floating stack per document: `wp:anchor/@relativeHeight`
 * orders shapes, text boxes, pictures and groups against each other, and
 * `behindDoc` puts a member of that stack under the text layer. So paint order
 * has to be ranked over the whole list. Ranking each family among itself
 * cannot compare a shape with a text box, and the browser then decides that
 * pair by DOM order — which is how a rectangle Word paints on top of a text
 * box ended up painted underneath it.
 *
 * This answers the PAINT question, from the heights the file states.
 * `resolveDocumentShapeStack` below answers a different one — whether the
 * reorder COMMAND may renumber the stack — and stays fail-closed, because
 * renumbering only the shapes in a document that also holds a text box would
 * move the text box in the stack without writing anything about it.
 *
 * A picture takes its place from its imported anchor's `relativeHeight`, or,
 * floated in the editor, from the `relativeHeight` it was given when floated
 * (`documentPictureStackHeight`). A picture that states neither (floated
 * before heights were given) has no place to take: it is counted, never
 * invented.
 */
export function documentFloatingStack(doc: DocumentShapeStackDocument): DocumentFloatingStack {
  // A document is immutable: its stack is read once. Every floating object's
  // paint asks for it (each stroke on an inked page), and walking the whole
  // document for each made painting n objects cost n walks.
  const cached = floatingStacks.get(doc);
  if (cached) return cached;
  const stack = readDocumentFloatingStack(doc);
  floatingStacks.set(doc, stack);
  return stack;
}

interface DocumentFloatingStack {
  /** Stated heights, document-wide, across every floating family (frozen). */
  readonly heights: readonly number[];
  /** Anchored pictures that state no height and so cannot be ranked. */
  readonly unrankedPictures: number;
}

const floatingStacks = new WeakMap<DocumentShapeStackDocument, DocumentFloatingStack>();

const STACK_OBJECTS: ReadonlySet<string> = new Set(["documentShape", "textBox", "image", "chart"]);
const HOLDS_STACK_OBJECT = new SheetFact((node) => STACK_OBJECTS.has(node.type.name));

type StackNode = Parameters<Parameters<DocumentShapeStackDocument["descendants"]>[0]>[0];
interface StackCollect { heights: number[]; unrankedPictures: number }

/** One node's place in the floating stack; true to read inside it. */
function collectStackNode(node: StackNode, out: StackCollect): boolean {
  const name = node.type.name;
  if (name === "documentShape") {
    // "In Line with Text" is not in the floating stack at all.
    if (node.attrs.wrapMode === "inline") return false;
    const height = documentShapeRelativeHeight(node.attrs.relativeHeight);
    if (height !== null) out.heights.push(height);
    return false;
  }
  if (name === "textBox") {
    // The frame keeps its whole `wp:anchor` reading; an in-line frame has
    // none and, like an inline shape, is not in the stack.
    const anchor = node.attrs.floatAnchor;
    const height =
      anchor && typeof anchor === "object"
        ? documentShapeRelativeHeight(
            (anchor as { relativeHeight?: unknown }).relativeHeight,
          )
        : null;
    if (height !== null) out.heights.push(height);
    return false;
  }
  if (
    name === "image" &&
    (node.attrs.officeDrawingKind === "anchor" ||
      (typeof node.attrs.dataFloat === "string" &&
        node.attrs.dataFloat.length > 0))
  ) {
    // An imported picture keeps its `wp:anchor` reading and so its place in
    // the stack; one floated in the editor states no height.
    // An imported picture keeps its `wp:anchor` reading and so its place in
    // the stack; one floated in the editor carries the `relativeHeight` it
    // was given when floated (pictureButtons / objectPlacement). One from
    // before heights were given states none and is counted, never invented.
    const anchor = node.attrs.floatAnchor;
    const height =
      anchor && typeof anchor === "object"
        ? documentShapeRelativeHeight(
            (anchor as { relativeHeight?: unknown }).relativeHeight,
          )
        : documentShapeRelativeHeight(node.attrs.relativeHeight);
    if (height !== null) out.heights.push(height);
    else out.unrankedPictures += 1;
    return false;
  }
  if (name === "chart") {
    // A floating chart keeps its `wp:anchor` reading, as a picture does,
    // and with it its place in the one stack; one in line is not in it.
    const anchor = node.attrs.floatAnchor;
    const height =
      anchor && typeof anchor === "object"
        ? documentShapeRelativeHeight((anchor as { relativeHeight?: unknown }).relativeHeight)
        : null;
    if (height !== null) out.heights.push(height);
    return false;
  }
  return true;
}

/** Each sheet's own stack (a document's top-level child), kept with the
 * sheet node: a stroke landing changes one sheet, and the document's stack is
 * the other sheets' kept lists and that one read again. Walking every sheet
 * of a 3,000-stroke note for each landing was 4-6 ms of an iPad's frame. */
const sheetStacks = new WeakMap<PmNode, { readonly heights: readonly number[]; readonly unrankedPictures: number }>();

function sheetFloatingStack(sheet: PmNode): { readonly heights: readonly number[]; readonly unrankedPictures: number } {
  let known = sheetStacks.get(sheet);
  if (!known) {
    const out: StackCollect = { heights: [], unrankedPictures: 0 };
    if (collectStackNode(sheet as unknown as StackNode, out)) {
      sheet.descendants((node) => collectStackNode(node as unknown as StackNode, out));
    }
    known = out;
    sheetStacks.set(sheet, known);
  }
  return known;
}

function readDocumentFloatingStack(doc: DocumentShapeStackDocument): DocumentFloatingStack {
  // An editor document asks only the sheets that hold such an object, each
  // sheet's own list read once per sheet node (`__mraiStackWalkWhole` walks
  // every sheet again).
  if (typeof (doc as { nodeAt?: unknown }).nodeAt === "function" &&
      !(globalThis as { __mraiStackWalkWhole?: boolean }).__mraiStackWalkWhole) {
    const pm = doc as unknown as PmNode;
    const heights: number[] = [];
    let unrankedPictures = 0;
    pm.forEach((sheet) => {
      if (!HOLDS_STACK_OBJECT.sheet(sheet)) return;
      const own = sheetFloatingStack(sheet);
      for (const height of own.heights) heights.push(height);
      unrankedPictures += own.unrankedPictures;
    });
    return Object.freeze({ heights: Object.freeze(heights), unrankedPictures });
  }
  const out: StackCollect = { heights: [], unrankedPictures: 0 };
  const walk: DocumentShapeStackDocument["descendants"] = typeof (doc as { nodeAt?: unknown }).nodeAt === "function"
    ? (visit) => HOLDS_STACK_OBJECT.descendants(doc as unknown as PmNode, visit as unknown as Parameters<SheetFact["descendants"]>[1])
    : (visit) => doc.descendants(visit);
  walk((node) => collectStackNode(node, out));
  return Object.freeze({ heights: Object.freeze(out.heights), unrankedPictures: out.unrankedPictures });
}

/**
 * Resolve the complete floating stack the editor can represent. Anchored
 * pictures and DrawingML object families outside the native WPS shape subset
 * block reordering: changing only the visible subset would lie about Word's
 * global floating-object order.
 */
const shapeStacks = new WeakMap<object, DocumentShapeStackResolution>();

/** The floating stack of `doc`. An editor document (immutable) keeps its
 * answer: the toolbar asks on every render, a selection change too. */
export function resolveDocumentShapeStack(
  doc: DocumentShapeStackDocument,
): DocumentShapeStackResolution {
  const immutable = typeof (doc as { nodeAt?: unknown }).nodeAt === "function";
  const known = immutable ? shapeStacks.get(doc) : undefined;
  if (known) return known;
  const resolved = findDocumentShapeStack(doc);
  if (immutable) shapeStacks.set(doc, resolved);
  return resolved;
}

function findDocumentShapeStack(
  doc: DocumentShapeStackDocument,
): DocumentShapeStackResolution {
  const entries: DocumentShapeStackEntry[] = [];
  let complete = true;
  doc.descendants((node, pos) => {
    const name = node.type.name;
    if (name === "documentShape") {
      // An inline shape does not float, exactly like an inline picture below:
      // it neither completes nor breaks the floating stack.
      if (node.attrs.wrapMode === "inline") return false;
      const relativeHeight = documentShapeRelativeHeight(
        node.attrs.relativeHeight,
      );
      if (
        relativeHeight === null ||
        node.attrs.zOrderStackComplete !== true
      ) {
        complete = false;
      } else {
        entries.push({ pos, relativeHeight });
      }
      return false;
    }
    if (
      name === "image" &&
      (node.attrs.officeDrawingKind === "anchor" ||
        (typeof node.attrs.dataFloat === "string" &&
          node.attrs.dataFloat.length > 0))
    ) {
      // A picture floated in the editor with its own height is a member the
      // command may renumber, like a shape; so is an imported anchored
      // picture whose import found the page showing the whole stack (its
      // anchor reading carries the height the save writes back). One that
      // states no height, or an import that saw members the page does not
      // show, still blocks: moving around it would move it unwritten.
      const anchor = documentAnchoredStackAnchor(node);
      const relativeHeight = anchor
        ? documentShapeRelativeHeight((anchor as { relativeHeight?: unknown }).relativeHeight)
        : documentPictureStackHeight(node.attrs);
      if (relativeHeight === null) complete = false;
      else entries.push({ pos, relativeHeight });
      return false;
    }
    // A chart in line with its text is not in the stack. A floating chart
    // (Layout Options, `chartPlacement`) holds its place by its own
    // `floatAnchor`, as an anchored picture does, and the save writes that
    // anchor whole -- a member like any other: Word's one stack orders it
    // against the shapes, frames and pictures. Left out, Bring to Front on a
    // shape under a floating chart found the shape alone at the top and
    // greyed the button. A source chart Word anchored that the import could
    // not take is already counted by the import: it leaves the shapes' stack
    // receipts (`zOrderStackComplete`) incomplete.
    if (name === "chart") {
      const anchor = node.attrs.floatAnchor;
      if (!anchor || typeof anchor !== "object") return false;
      const relativeHeight = documentShapeRelativeHeight(
        (anchor as { relativeHeight?: unknown }).relativeHeight,
      );
      if (relativeHeight === null) complete = false;
      else entries.push({ pos, relativeHeight });
      return false;
    }
    if (name === "textBox") {
      // A frame in line with text is not in the stack; a floating one is a
      // member when its import saw the whole stack (`zOrderStackComplete`).
      if (!node.attrs.floatAnchor) return false;
      const anchor = documentAnchoredStackAnchor(node);
      const relativeHeight = anchor
        ? documentShapeRelativeHeight((anchor as { relativeHeight?: unknown }).relativeHeight)
        : null;
      if (relativeHeight === null) complete = false;
      else entries.push({ pos, relativeHeight });
      return false;
    }
    if (name === "smartArt") {
      complete = false;
      return false;
    }
    return true;
  });
  if (!complete) return { complete: false, backToFront: [] };
  entries.sort(
    (left, right) =>
      left.relativeHeight - right.relativeHeight || left.pos - right.pos,
  );
  return { complete: true, backToFront: entries };
}

/**
 * Plan one bounded z-order command. The returned patches canonicalize the
 * complete representable stack into consecutive UInt32 values, so equal source
 * heights become deterministic and every NodeView can paint a small CSS rank.
 */
/** The selected member already sits where the action would put it (Send to
 *  Back on the backmost one): the button changes nothing. */
export function documentShapeZOrderIsCurrent(
  stack: DocumentShapeStackResolution,
  selectedPos: number,
  action: DocumentShapeZOrderAction,
): boolean {
  if (!stack.complete) return false;
  const index = stack.backToFront.findIndex((entry) => entry.pos === selectedPos);
  if (index < 0) return false;
  const front = action === "bringToFront" || action === "bringForward";
  return front ? index === stack.backToFront.length - 1 : index === 0;
}

export function documentShapeZOrderPatches(
  stack: DocumentShapeStackResolution,
  selectedPos: number,
  action: DocumentShapeZOrderAction,
): readonly DocumentShapeStackEntry[] | null {
  if (!stack.complete || stack.backToFront.length === 0) return null;
  const ordered = [...stack.backToFront];
  const index = ordered.findIndex((entry) => entry.pos === selectedPos);
  if (index < 0) return null;
  let destination = index;
  if (action === "bringToFront") destination = ordered.length - 1;
  else if (action === "bringForward") destination = Math.min(index + 1, ordered.length - 1);
  else if (action === "sendBackward") destination = Math.max(index - 1, 0);
  else destination = 0;
  if (destination === index) return null;

  // Word moves only the selected member (Word 16.86, measured on a shape, a
  // diagram and a picture of one stack): to the front, the highest height
  // plus 1024; to the back, the lowest minus 1025; one step, one less than
  // the midpoint of its new neighbours. Every other member keeps its height.
  const others = ordered.filter((_, position) => position !== index);
  const below = destination > 0 ? others[destination - 1]?.relativeHeight : undefined;
  const above = destination < others.length ? others[destination]?.relativeHeight : undefined;
  const wordHeight =
    above === undefined
      ? (below ?? ordered[index].relativeHeight) + 1024
      : below === undefined
        ? above - 1025
        : Math.floor((below + above) / 2) - 1;
  if (
    Number.isInteger(wordHeight) &&
    wordHeight >= 0 &&
    wordHeight <= DOCUMENT_SHAPE_MAX_RELATIVE_HEIGHT &&
    (below === undefined || wordHeight > below) &&
    (above === undefined || wordHeight < above)
  ) {
    return [{ pos: ordered[index].pos, relativeHeight: wordHeight }];
  }
  // No room between the neighbours (heights a step apart): the stack is
  // renumbered into consecutive heights instead.
  const [selected] = ordered.splice(index, 1);
  ordered.splice(destination, 0, selected);
  const base = Math.min(
    DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
    DOCUMENT_SHAPE_MAX_RELATIVE_HEIGHT - ordered.length,
  );
  return ordered
    .map((entry, rank) => ({
      pos: entry.pos,
      relativeHeight: base + rank,
    }))
    .filter(
      (entry) =>
        stack.backToFront.find((current) => current.pos === entry.pos)
          ?.relativeHeight !== entry.relativeHeight,
    );
}

/**
 * The z-order commands on SEVERAL selected members (Shift / Ctrl + click):
 * Office moves the selection as a set and keeps its own order. Bring to Front
 * / Send to Back carry every selected member past the others; Bring Forward /
 * Send Backward move each one past the next unselected member, and a member
 * already against a selected neighbour (or the edge) stays. Each member moves
 * as the single command moves it (`documentShapeZOrderPatches`, Word's
 * heights), one after another; the heights that changed, or null when nothing
 * moves, the stack cannot be renumbered, or a position is not a member.
 */
export function documentShapeZOrderPatchesForSelection(
  stack: DocumentShapeStackResolution,
  positions: readonly number[],
  action: DocumentShapeZOrderAction,
): readonly DocumentShapeStackEntry[] | null {
  if (!stack.complete || positions.length === 0) return null;
  const selected = new Set(positions);
  const members = stack.backToFront.filter((entry) => selected.has(entry.pos));
  if (members.length !== selected.size) return null;
  if (members.length === 1) {
    const only = members[0];
    return only ? documentShapeZOrderPatches(stack, only.pos, action) : null;
  }
  // Already the whole top (bottom) block: To Front (To Back) changes nothing.
  const edge =
    action === "bringToFront"
      ? stack.backToFront.slice(-members.length)
      : action === "sendToBack"
        ? stack.backToFront.slice(0, members.length)
        : null;
  if (edge && edge.every((entry) => selected.has(entry.pos))) return null;
  // To the front, the backmost goes first (each lands on top, so the last
  // one moved -- the frontmost -- stays on top); one step forward, the
  // frontmost goes first (so it makes room for the one under it).
  const order =
    action === "bringToFront" || action === "sendBackward"
      ? members
      : [...members].reverse();
  let current: DocumentShapeStackEntry[] = [...stack.backToFront];
  for (const member of order) {
    const index = current.findIndex((entry) => entry.pos === member.pos);
    if (index < 0) return null;
    if (action === "bringForward") {
      const next = current[index + 1];
      if (!next || selected.has(next.pos)) continue;
    } else if (action === "sendBackward") {
      const previous = index > 0 ? current[index - 1] : undefined;
      if (!previous || selected.has(previous.pos)) continue;
    }
    const patches = documentShapeZOrderPatches(
      { complete: true, backToFront: current },
      member.pos,
      action,
    );
    if (!patches?.length) continue;
    const moved = new Map(patches.map((patch) => [patch.pos, patch.relativeHeight]));
    current = current
      .map((entry) => {
        const height = moved.get(entry.pos);
        return height === undefined ? entry : { pos: entry.pos, relativeHeight: height };
      })
      .sort((left, right) => left.relativeHeight - right.relativeHeight || left.pos - right.pos);
  }
  const before = new Map(stack.backToFront.map((entry) => [entry.pos, entry.relativeHeight]));
  const changed = current.filter((entry) => before.get(entry.pos) !== entry.relativeHeight);
  return changed.length > 0 ? changed : null;
}

/**
 * Shape Format > Reorder Objects on the document's floating stack: `frontFirst`
 * names some of the stack's members (by position), front first, in their new
 * order. The listed members swap the heights they held between them, so every
 * other floating object keeps its own depth (PowerPoint's dialog, run on
 * Word's one floating stack). The heights that change, or null when nothing
 * moves, the stack cannot be renumbered, or the list names other objects.
 */
export function documentShapeStackReorderPatches(
  stack: DocumentShapeStackResolution,
  frontFirst: readonly number[],
): readonly DocumentShapeStackEntry[] | null {
  if (!stack.complete || frontFirst.length < 2) return null;
  if (new Set(frontFirst).size !== frontFirst.length) return null;
  const listed = new Set(frontFirst);
  const members = stack.backToFront.filter((entry) => listed.has(entry.pos));
  if (members.length !== frontFirst.length) return null;
  let heights = members.map((entry) => entry.relativeHeight);
  // Equal heights paint in document order; renumber them apart from the
  // lowest so the new order is stated, not left to document order.
  if (heights.some((height, index) => index > 0 && height <= heights[index - 1])) {
    const base = Math.min(heights[0], DOCUMENT_SHAPE_MAX_RELATIVE_HEIGHT - heights.length);
    heights = heights.map((_, index) => base + index);
  }
  const current = new Map(members.map((entry) => [entry.pos, entry.relativeHeight]));
  const patches = [...frontFirst]
    .reverse()
    .map((pos, index) => ({ pos, relativeHeight: heights[index] }))
    .filter((entry) => current.get(entry.pos) !== entry.relativeHeight);
  return patches.length > 0 ? patches : null;
}

const DOCUMENT_CONNECTION_SITE_ANCHORS = ["top", "right", "bottom", "left"] as const;

/**
 * Shape Format > Edit Shape > Reroute Connectors on a group's native
 * connectors (Word draws glued connectors only inside a drawing canvas or a
 * group -- Microsoft Support's note on Reroute Connectors): each connector
 * glued at both ends moves to the closest pair of connection sites of the two
 * members it joins, the slide's rule (`nearestAnchorPair`), and its route is
 * redrawn between them. Null when no connector changes.
 */
export function documentShapeReroutedConnectorMembers(
  members: readonly DocumentShapeGroupMember[],
): readonly DocumentShapeGroupMember[] | null {
  const byId = new Map(members.map((member) => [member.memberId, member]));
  const rect = (member: DocumentShapeGroupMember) => ({
    left: member.xEmu,
    top: member.yEmu,
    width: member.widthEmu,
    height: member.heightEmu,
  });
  let changed = false;
  const next = members.map((member) => {
    if (member.kind !== "connector") return member;
    const start = byId.get(member.startMemberId ?? "");
    const end = byId.get(member.endMemberId ?? "");
    if (!start || !end || start.kind === "connector" || end.kind === "connector") return member;
    const pair = nearestAnchorPair(rect(start), rect(end));
    const startSite = DOCUMENT_CONNECTION_SITE_ANCHORS.indexOf(pair.start as "top");
    const endSite = DOCUMENT_CONNECTION_SITE_ANCHORS.indexOf(pair.end as "top");
    if (startSite < 0 || endSite < 0) return member;
    if (startSite === member.startConnectionSite && endSite === member.endConnectionSite) return member;
    const route = documentShapeConnectorRoute(start, startSite, end, endSite);
    if (!route) return member;
    changed = true;
    return { ...member, startConnectionSite: startSite, endConnectionSite: endSite, ...route };
  });
  return changed ? next : null;
}

/** Whether a shape is a group holding a connector glued at both ends. */
export function documentShapeHasGluedConnectors(attrs: Pick<DocumentShapeAttrs, "groupMembers">): boolean {
  return Boolean(
    attrs.groupMembers?.some(
      (member) => member.kind === "connector" && member.startMemberId && member.endMemberId,
    ),
  );
}

export type DocumentShapePreset = OfficeShapePreset;
export type DocumentShapeCategoryId = OfficeShapeCategoryId;
export type DocumentShapeDefinition = OfficeShapeDefinition;
export type DocumentShapeCategory = OfficeShapeCategory;

export const DOCUMENT_SHAPE_CATEGORY_IDS = OFFICE_SHAPE_CATEGORY_IDS;
export const DOCUMENT_SHAPE_DEFINITIONS = OFFICE_SHAPE_DEFINITIONS;
export const DOCUMENT_SHAPE_PRESETS = OFFICE_SHAPE_PRESETS;
export const DOCUMENT_SHAPE_CATEGORIES = OFFICE_SHAPE_CATEGORIES;
export const documentShapeDefinition = officeShapeDefinition;
export const isDocumentShapePreset = isOfficeShapePreset;
/** A document shape's preset is all it knows of the route, so the route is kept. */
export const documentShapePresetFromDrawingMl = (
  drawingPreset: unknown,
  lineEnd: unknown = null,
  lineStart: unknown = null,
): OfficeShapePreset | null =>
  officeShapePresetFromDrawingMl(drawingPreset, lineEnd, lineStart, {
    keepRoute: true,
  });

export const DOCUMENT_SHAPE_RECENT_LIMIT = 18;
export const DOCUMENT_SHAPE_RECENT_STORAGE_KEY =
  "mrnote.document.shape-recents.drawingml-prstgeom-v2";

export const parseDocumentShapeRecents = parseOfficeShapeRecents;
export const pushDocumentShapeRecent = pushOfficeShapeRecent;
export const documentShapeMatchesSearch = officeShapeMatchesSearch;

function finiteNumber(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}









function normalizeDocumentShapeOutlineWidthPt(value: unknown, fallback: number): number {
  return Math.round(Math.min(24, Math.max(0, finiteNumber(value, fallback))) * 100) / 100;
}

function boundedInteger(
  value: unknown,
  fallback: number,
  min: number,
  max: number,
): number {
  return Math.round(
    Math.min(max, Math.max(min, finiteNumber(value, fallback))),
  );
}

export function normalizeDocumentShapeColor(
  value: unknown,
  fallback: string,
): string {
  if (value === "none") return "none";
  return typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value)
    ? value.toLowerCase()
    : fallback;
}

export function normalizeDocumentShapeText(value: unknown): string {
  if (typeof value !== "string") return "";
  return value
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "")
    .slice(0, DOCUMENT_SHAPE_MAX_TEXT_LENGTH);
}

/**
 * One `w:p` of an imported shape's `w:txbxContent`. Word lays a shape's text
 * out as real paragraphs — each with its own `w:jc` (ECMA-376 17.3.1.13),
 * size and `w:spacing` — inside the body box `bodyPr` describes.
 */
export interface DocumentShapeTextBodyParagraph {
  readonly text: string;
  readonly align?: "left" | "center" | "right" | "justify";
  readonly fontSizePt?: number;
  /** The face the paragraph is set in, when the source states one. */
  readonly fontFamily?: string;
  readonly spaceBeforePt?: number;
  readonly spaceAfterPt?: number;
  readonly lineHeight?: number;
  /** An exact or at-least line (`w:lineRule` `exact` / `atLeast`), in points
   * with its rule: both or neither. */
  readonly lineHeightPt?: number;
  readonly lineHeightRule?: "exact" | "atLeast";
  /** The paragraph's own `w:ind`, in points (a hanging indent is a negative first line). */
  readonly indentLeftPt?: number;
  readonly indentRightPt?: number;
  readonly firstLineIndentPt?: number;
  /** The paragraph's own `w:pStyle`, which the full writer states again. */
  readonly styleId?: string;
}

/**
 * An imported shape's text body: `bodyPr`'s insets and vertical anchoring
 * (ECMA-376 Part 1 §21.1.2.1.1 `ST_TextAnchoringType`) plus its paragraphs.
 *
 * Package-owned, exactly like {@link DocumentShapeAttrs} `floatAnchor`: no
 * writer reads it, and the shape's flat `text` stays the editable string a
 * save writes back. A shape without one draws from `text` as before.
 */
export interface DocumentShapeTextBody {
  readonly anchor: "t" | "ctr" | "b";
  readonly insetLeftEmu: number;
  readonly insetTopEmu: number;
  readonly insetRightEmu: number;
  readonly insetBottomEmu: number;
  readonly paragraphs: readonly DocumentShapeTextBodyParagraph[];
}

const shapeTextBodyNumber = (
  value: unknown,
  low: number,
  high: number,
): number | null =>
  typeof value === "number" && Number.isFinite(value) && value >= low && value <= high
    ? value
    : null;

export function normalizeDocumentShapeTextBody(
  value: unknown,
): DocumentShapeTextBody | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const input = value as Record<string, unknown>;
  const anchor = input.anchor;
  if (anchor !== "t" && anchor !== "ctr" && anchor !== "b") return null;
  const insets: Record<string, number> = {};
  for (const [key, fallback] of [
    ["insetLeftEmu", 91_440],
    ["insetTopEmu", 45_720],
    ["insetRightEmu", 91_440],
    ["insetBottomEmu", 45_720],
  ] as const) {
    const inset = shapeTextBodyNumber(input[key], 0, DOCUMENT_SHAPE_MAX_EDGE_EMU);
    insets[key] = inset === null ? fallback : Math.round(inset);
  }
  const raw = input.paragraphs;
  if (!Array.isArray(raw) || raw.length < 1 || raw.length > 200) return null;
  const paragraphs: DocumentShapeTextBodyParagraph[] = [];
  for (const entry of raw) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) return null;
    const row = entry as Record<string, unknown>;
    const align = row.align;
    const fontSizePt = shapeTextBodyNumber(row.fontSizePt, 1, 1_638);
    const spaceBeforePt = shapeTextBodyNumber(row.spaceBeforePt, 0, 1_584);
    const spaceAfterPt = shapeTextBodyNumber(row.spaceAfterPt, 0, 1_584);
    const lineHeight = shapeTextBodyNumber(row.lineHeight, WORD_LINE_MULTIPLE_MIN, WORD_LINE_MULTIPLE_MAX);
    const lineHeightPt = shapeTextBodyNumber(row.lineHeightPt, 0.05, 1_584);
    const lineHeightRule = row.lineHeightRule === "exact" || row.lineHeightRule === "atLeast" ? row.lineHeightRule : null;
    const indentLeftPt = shapeTextBodyNumber(row.indentLeftPt, -1_584, 1_584);
    const indentRightPt = shapeTextBodyNumber(row.indentRightPt, -1_584, 1_584);
    const firstLineIndentPt = shapeTextBodyNumber(row.firstLineIndentPt, -1_584, 1_584);
    paragraphs.push({
      text: normalizeDocumentShapeText(row.text),
      ...(align === "left" || align === "center" || align === "right" || align === "justify"
        ? { align }
        : {}),
      ...(fontSizePt === null ? {} : { fontSizePt }),
      ...(typeof row.fontFamily === "string" && row.fontFamily.trim() && row.fontFamily.length <= 128
        ? { fontFamily: row.fontFamily.trim() }
        : {}),
      ...(spaceBeforePt === null ? {} : { spaceBeforePt }),
      ...(spaceAfterPt === null ? {} : { spaceAfterPt }),
      ...(lineHeight === null ? {} : { lineHeight }),
      ...(lineHeightPt === null || lineHeightRule === null ? {} : { lineHeightPt, lineHeightRule }),
      ...(indentLeftPt === null ? {} : { indentLeftPt }),
      ...(indentRightPt === null ? {} : { indentRightPt }),
      ...(firstLineIndentPt === null ? {} : { firstLineIndentPt }),
      ...(typeof row.styleId === "string" &&
      row.styleId.length > 0 &&
      row.styleId.length <= 253 &&
      !/[\x00-\x1f<>"&]/.test(row.styleId)
        ? { styleId: row.styleId }
        : {}),
    });
  }
  return {
    anchor,
    insetLeftEmu: insets.insetLeftEmu,
    insetTopEmu: insets.insetTopEmu,
    insetRightEmu: insets.insetRightEmu,
    insetBottomEmu: insets.insetBottomEmu,
    paragraphs,
  };
}

/** A Selection Pane name: one line, no control characters, at most 255. */
export function normalizeDocumentShapeName(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const name = value.replace(/[\u0000-\u001f\u007f]/g, " ").trim().slice(0, 255);
  return name || null;
}

export function normalizeDocumentShapeAltText(value: unknown): string {
  if (typeof value !== "string") return "";
  return value
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "")
    .slice(0, DOCUMENT_SHAPE_MAX_ALT_TEXT_LENGTH);
}

export function normalizeDocumentShapeEffect(
  value: unknown,
): DocumentShapeEffectPreset {
  return value === "shadow" ||
    value === "reflection" ||
    value === "glow" ||
    value === "softEdge" ||
    value === "bevel" ||
    value === "threeD"
    ? value
    : "none";
}

export function normalizeDocumentShapeTextDirection(
  value: unknown,
): DocumentShapeTextDirection {
  return value === "rotate90" ||
    value === "rotate270" ||
    value === "stacked"
    ? value
    : "horizontal";
}

export function normalizeDocumentShapeTextVerticalAlign(
  value: unknown,
): DocumentShapeTextVerticalAlign {
  return value === "top" || value === "bottom" ? value : "middle";
}

export function normalizeDocumentShapeTextUnderline(
  value: unknown,
): DocumentShapeTextUnderline {
  return value === "single" || value === "double" ? value : "none";
}

/** Word's Add Text writes `w:jc="center"`, which is also how the canvas draws a label. */
export function normalizeDocumentShapeTextHorizontalAlign(
  value: unknown,
): DocumentShapeTextHorizontalAlign {
  return value === "left" || value === "right" || value === "justify"
    ? value
    : "center";
}

export function normalizeDocumentShapeParagraphAlign(
  value: unknown,
): DocumentShapeParagraphAlign | undefined {
  return value === "left" || value === "center" || value === "right"
    ? value
    : undefined;
}

/**
 * The bounded adjust-handle map, or null. Word bounds a shape at eight
 * adjustments ("Adjustments object (Word)"), and the values are deliberately
 * not clamped to 0..100000: several presets are authored outside that range
 * (a callout elbow at -8333), and the preset's own formulas pin each handle.
 */
export function normalizeDocumentShapeAdjustments(
  value: unknown,
): Readonly<Record<string, number>> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const entries = Object.entries(value as Record<string, unknown>);
  if (entries.length === 0 || entries.length > 8) return null;
  const bounded: Record<string, number> = {};
  for (const [name, amount] of entries) {
    if (
      !/^[A-Za-z0-9_]{1,81}$/.test(name) ||
      typeof amount !== "number" ||
      !Number.isSafeInteger(amount) ||
      amount < -2_147_483_648 ||
      amount > 2_147_483_647
    ) {
      return null;
    }
    bounded[name] = amount;
  }
  return bounded;
}

export function normalizeDocumentShapeWrapMode(
  value: unknown,
): DocumentShapeWrapMode {
  return value === "tight" ||
    value === "through" ||
    value === "topAndBottom" ||
    value === "behindText" ||
    value === "inFrontOfText" ||
    value === "inline"
    ? value
    : "square";
}

export function normalizeDocumentShapeFontFamily(value: unknown): string {
  // The one family sanitiser every Office shape label uses; Docs only adds
  // its default face for "no family".
  return (
    normalizeOfficeShapeFontFamily(value) ??
    DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_FAMILY
  );
}

function documentShapeLinkedTextBoxCoordinate(
  value: unknown,
  min: number,
): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= min && number <= 0xffff
    ? number
    : null;
}

/** A whole-percent Transparency 1..100, or null (none). */
function documentShapeWholePercent(value: unknown): number | null {
  return typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= 100 ? value : null;
}

/**
 * A shape's picture fill: the inline PNG / JPEG picture, stretched (with the
 * stretch's offsets) or tiled, and the shared Format Shape options
 * (`lib/office-picture-fill-options`: Transparency, Rotate with shape).
 * Exactly one of `stretch` / `tile`.
 */
export interface DocumentShapePictureFill extends OfficePictureFillOptions {
  readonly src: string;
  readonly stretch?: true;
}

const DOCUMENT_SHAPE_PICTURE_FILL_KEYS = new Set(["src", "stretch", "tile", "fillRect", "transparency", "rotWithShape"]);

/** `value` as a shape's picture fill, or null when it is not one. */
export function normalizeDocumentShapePictureFill(value: unknown): DocumentShapePictureFill | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const src = smartArtPictureSrc(record.src);
  if (!src || Object.keys(record).some((key) => !DOCUMENT_SHAPE_PICTURE_FILL_KEYS.has(key))) return null;
  if ((record.stretch === true) === (record.tile !== undefined)) return null;
  const options = normalizeOfficePictureFillOptions({
    tile: record.tile,
    fillRect: record.fillRect,
    transparency: record.transparency,
    rotWithShape: record.rotWithShape,
  });
  if (!options) return null;
  return options.tile ? { src, ...options } : { src, stretch: true, ...options };
}

export interface DocumentShapeAttrs {
  textPaint?: OfficeTextPaint | null;
  fillGradient?: OfficeShapeGradient | null;
  shapeShadow?: OfficeShapeShadow | null;
  /**
   * Shape Effects beyond the outer shadow -- glow, reflection, soft edges and
   * an inner shadow -- in the editors' shared effect model
   * (`office-picture-effects`; the server's `document_docx_shape_effects`).
   * The one Shape Effects preset (`effect`) replaces them when picked.
   */
  shapeEffects?: OfficePictureEffects | null;
  /** Shape Outline > Sketched. Absent: the page states none (an older page
   * keeps the file's); null: no sketch, written out. */
  outlineSketch?: OfficeLineSketch | null;
  /**
   * Format Shape > Fill & Line beyond the colours, as the server reads and
   * writes them (`a:pattFill`, the solid fill's and line's `a:alpha`, a
   * gradient line `a:ln/a:gradFill`). Absent: the page states none (an older
   * page keeps the file's); null: none, written out.
   */
  fillPattern?: OfficeChartFillPattern | null;
  /**
   * Format Shape > Fill > Picture or texture fill (`a:blipFill` with
   * `a:stretch/a:fillRect` in `wps:spPr`): the picture, stretched over the
   * shape's box inside its geometry, the words drawn on top. The stretched
   * form of a text box's `DocumentTextBoxPictureFill`. It IS the fill while
   * present (`fill` is the colour underneath, kept for when another kind is
   * picked). Absent: the page states none (an older page keeps the file's);
   * null: none, written out.
   */
  fillPicture?: DocumentShapePictureFill | null;
  /** Whole percent 1..100 on a solid fill. */
  fillTransparency?: number | null;
  /** Whole percent 1..100 on a solid line. */
  outlineTransparency?: number | null;
  outlineGradient?: OfficeShapeGradient | null;
  schemaVersion: typeof DOCUMENT_SHAPE_SCHEMA_VERSION;
  preset: DocumentShapePreset;
  /**
   * The shape's OWN outline, in the shared custom-geometry model every Office
   * surface persists (`lib/presentation-freeform`, 0..100000 in the shape's
   * own box). Present on a shape that is not any preset — a pyramid's band, a
   * funnel's vessel, a segmented cycle's wedge — and it WINS over
   * {@link preset}, which then only names the nearest auto-shape for a host or
   * a package that can still draw only presets.
   *
   * Geometry, not identity: it survives the identity strip that turns a
   * converted diagram back into an ordinary group.
   */
  customPath?: PresentationCustomPath | null;
  /**
   * How each of {@link customPath}'s paths is painted when they paint apart
   * (`a:path@fill` / `@stroke`: a bent arrow's darker fold, an arc's unstroked
   * wedge), as `lib/office-custom-path-paints` reads it. A group member's,
   * from its package; absent: the whole outline takes the shape's paint.
   */
  readonly customPathPaints?: readonly OfficeCustomPathPaint[];
  xEmu: number;
  yEmu: number;
  widthEmu: number;
  heightEmu: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /**
   * The shape's dragged adjust handles — the yellow ones — as `a:avLst` holds
   * them: the guide name the preset declares mapped to its literal value
   * (`{ adj: 40000 }`). Absent means the preset's own defaults. The canvas
   * evaluates the same catalog program the slide canvas does with these values
   * substituted, so a dragged corner keeps its radius at every size.
   */
  adjustments?: Readonly<Record<string, number>>;
  /**
   * `a:ln/a:headEnd|a:tailEnd`. {@link preset} names WHICH ends are decorated;
   * these name WHICH decoration (all six `a:ST_LineEndType` members) and its
   * optional `@w` / `@len`. A value the preset already implies is ABSENT, and
   * an absent size stays absent on the way out: no attribute of
   * `CT_LineEndProperties` has a schema default, and Word writes
   * `<a:tailEnd type="triangle"/>` bare where a slide writes `w="med"
   * len="med"` (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §7.4).
   */
  headEnd?: PresentationConnectorEnd;
  tailEnd?: PresentationConnectorEnd;
  headEndWidth?: PresentationConnectorEndSize;
  headEndLength?: PresentationConnectorEndSize;
  tailEndWidth?: PresentationConnectorEndSize;
  tailEndLength?: PresentationConnectorEndSize;
  /**
   * Persisted `wp:anchor/@relativeHeight`. Omission is distinct from the Word
   * default so legacy page atoms cannot rewrite authenticated source custody.
   */
  relativeHeight?: number;
  /**
   * Source-derived receipt: every body `wp:anchor` mapped exactly once to an
   * editable native shape. Omission/false keeps global z-order fail-closed.
   */
  zOrderStackComplete?: boolean;
  /**
   * Persisted `a:spLocks/@noChangeAspect` intent. Omission is distinct from an
   * explicit false so legacy page atoms cannot accidentally clear a true lock
   * from their authenticated original package during an unrelated edit.
   */
  lockAspectRatio?: boolean;
  /**
   * Word's Selection Pane name (`wp:docPr/@name`). Absent: the writer names
   * the drawing "Shape N" as Word does for an unnamed one.
   */
  name?: string;
  /**
   * Word's Selection Pane eye (`wp:docPr/@hidden`): the drawing stays in the
   * document and the pane, and is not drawn on the page.
   */
  hidden?: boolean;
  /**
   * Word's Format Shape > Text Box > "Resize shape to fit text" (`a:spAutoFit`
   * in `wps:bodyPr`): the shape's height follows its text. Word offers no
   * shrink-on-overflow (that is PowerPoint's `a:normAutofit`). `"none"` is an
   * explicit off, so an in-place save removes the file's element; absent
   * states nothing and the file keeps what it has.
   */
  textAutofit?: "none" | "shape";
  /**
   * The words' WordArt transform (`wps:bodyPr/a:prstTxWarp`; a VML WordArt's
   * shapetype): one `ST_TextShapeType` token and its stated guides, the ONE
   * warp value the three editors share (`@/lib/office-text-warp`). Absent
   * bends nothing. Drawn by `@/lib/office-text-warp-render`.
   */
  textWarp?: OfficeTextWarp;
  /**
   * Format Shape > 3-D Rotation > Keep text flat (`wps:bodyPr/a:flatTx`): the
   * words stay flat on the front of a shape its 3-D rotation turns. `true`
   * checked; `false` unchecked, so an in-place save takes the file's element
   * off; absent states nothing and the file keeps what it has.
   */
  keepTextFlat?: boolean;
  isTextBox: boolean;
  text: string;
  /**
   * The source `bodyPr` + `w:txbxContent` paragraph reading of an imported
   * shape's text. See {@link DocumentShapeTextBody}.
   */
  textBody?: DocumentShapeTextBody | null;
  /**
   * The words' own character formatting, run by run, where part of them is
   * formatted (lib/document-shape-text-runs). A run's field overrides the
   * whole-shape `text*` field; the runs always spell `text` exactly.
   */
  textRuns?: readonly TextRun[] | null;
  fill: string;
  outline: string;
  outlineWidthPt: number;
  outlineStyle: StrokeStyle;
  /**
   * `a:ln@algn`. `"inset"` is the `in` pen — the whole stroke inside the
   * geometry path, Office's `LineFormat.InsetPen`. `null` is a shape whose
   * source line states no alignment, which is every shape authored here and
   * every page stored before audit §64; it is drawn centred and exports no
   * `algn` at all. See `@/lib/office-shape-line-align`.
   */
  outlineAlign: OfficeStrokeAlign | null;
  /**
   * The exact `a:prstDash` preset when the file states more than the style's
   * family (Long Dash, Round Dot, Dash Dot...; `@/lib/office-stroke-dash`).
   * Absent on a shape authored here and on every earlier page.
   */
  outlineDash?: OfficeStrokeDash | null;
  /**
   * The line's compound type, cap, join and miter limit, each only when the
   * source line states it (`@/lib/office-stroke-detail`).
   */
  outlineCompound?: OfficeStrokeDetail["outlineCompound"] | null;
  outlineCap?: OfficeStrokeDetail["outlineCap"] | null;
  outlineJoin?: OfficeStrokeDetail["outlineJoin"] | null;
  outlineMiterLimit?: number | null;
  /**
   * Draw-tab ink. A stroke is an ordinary shape -- its own open outline
   * (`customPath`), a line, no fill, floating In Front Of Text -- and this
   * names the tool that drew it: a highlighter paints translucent and
   * multiplies over the words, as in Word; a pen paints opaque. Absent on
   * every shape that is not ink, so nothing stored before ink existed changes.
   * The same two tools the deck and the worksheet ink with.
   */
  inkTool?: DocumentInkTool;
  /** Draw-tab ink drawn with one of Word's pen effects: the stroke paints
   * the effect's gradient; `outline` holds its solid stand-in, which is what
   * every file writer states. Absent on every other shape. */
  inkEffect?: DocumentInkEffect;
  /** Draw-tab ink a point eraser cut: which of the stroke's ends stop flat
   * where the eraser was ({@link InkCutEnds}). Absent on every other shape. */
  inkCutEnds?: InkCutEnds;
  /** Draw-tab ink drawn with a pressure pen: one integer 0..100 per sample
   * of the outline, space-separated; the painter widens and thins the
   * stroke with it. Absent on a stroke drawn at one width and on every
   * other shape; no file writer reads it (a file line has one width). */
  inkPressure?: string;
  /** Draw-tab ink written while the Draw tab recorded audio: the recording
   * and the moment of it, `"<recording id> <ms>"` ({@link normalizeInkAudio}).
   * Playing the recording shows the writing as it was written; tapping the
   * stroke plays from there. Absent on every other stroke; no writer reads it. */
  inkAudio?: string;
  /** Draw ▸ Tape: the stroke is a strip of tape over the page, `"covered"`
   * (painted nearly opaque, hiding the words under it) or `"revealed"`
   * (peeled back to a faint strip, so they show); a tap flips it, for
   * learning by heart. Absent on every other stroke; the file writers draw
   * it as the pen line it is ({@link normalizeInkTape}). */
  inkTape?: InkTape;
  effect: DocumentShapeEffectPreset;
  textFill: string;
  textOutline: string;
  textOutlineWidthPt: number;
  textEffect: DocumentShapeEffectPreset;
  textFontSizePt: number;
  textFontFamily: string;
  textBold: boolean;
  /** `w:i` on the shape's text runs (Home ▸ Italic with the shape selected). */
  textItalic: boolean;
  /** `w:u` on the shape's text runs; "none" is no underline. */
  textUnderline: DocumentShapeTextUnderline;
  /** `w:strike` on the shape's text runs. */
  textStrike: boolean;
  textDirection: DocumentShapeTextDirection;
  textVerticalAlign: DocumentShapeTextVerticalAlign;
  





  textHorizontalAlign: DocumentShapeTextHorizontalAlign;
  /**
   * The alignment of the PARAGRAPH an "In Line with Text" shape stands in
   * alone (`w:jc` of that paragraph): Word treats an in-line drawing like a
   * character, so centring the paragraph centres the graphic. Absent is the
   * paragraph's default; only an in-line shape carries it. The same key a
   * chart node keeps for the same thing.
   */
  textAlign?: DocumentShapeParagraphAlign;
  textWrap: boolean;
  altText: string;
  /**
   * Office's Alt Text "Mark as decorative"
   * (`wp:docPr/a:extLst/a:ext[@uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"]
   * /adec:decorative@val`): a screen reader skips the drawing, and it carries
   * no description.
   */
  decorative?: boolean;
  wrapMode: DocumentShapeWrapMode;
  linkedTextBoxStoryId?: number;
  linkedTextBoxSequence?: number;
  source: Record<string, unknown> | null;
  /**
   * A strict native WordprocessingGroup payload. Members use the group's
   * `chOff/chExt` coordinate space; the outer shape owns the floating anchor.
   */
  groupMembers?: readonly DocumentShapeGroupMember[];
  groupCoordinateWidthEmu?: number;
  groupCoordinateHeightEmu?: number;
  /**
   * The group's OWN fill (`wpg:grpSpPr`): a linear gradient, a stretched
   * picture or a pattern laid over the group's child space, worn by the
   * members marked {@link DocumentShapeGroupMember.fillFromGroup}
   * (`a:grpFill`), each drawing the part under its own box when it is drawn
   * (`office-shape-group-fill`). The group itself draws no fill.
   */
  groupFill?: DocumentShapeGroupFill;
  /**
   * The catalog SmartArt this group IS (the engine kind, e.g. `basic-list`).
   * It is the identity the Docs SmartArt contextual tab recognises, and it is
   * what the DOCX writer stamps into the group's `wpg:grpSpPr/a:extLst` so an
   * exported document reimports as the same recognised diagram. Absent on an
   * ordinary shape group and on a preserved non-catalog Office diagram.
   */
  smartArtKind?: string;
  /** One id shared by every member of one diagram (the engine's `groupId`). */
  smartArtGroupId?: string;
  /**
   * Which of the fourteen PowerPoint SmartArt Styles this diagram wears
   * (`docs/office-graphics/SMARTART-STYLES.md` §1). It rides WITH the identity
   * — a group that is not a recognised diagram carries no style — and it is
   * stamped into the DOCX identity extension, so reopening the document knows
   * which of the fourteen it is instead of inferring one from how the members
   * happen to be painted.
   */
  smartArtStyle?: SmartArtStyleId;
  /**
   * SmartArt Design > Right to Left: the diagram is drawn mirrored (Office's
   * `dgm:dir val="rev"`). A property of the DIAGRAM, so it rides with the
   * identity exactly like {@link smartArtStyle}; the members only carry the
   * mirrored geometry and `flipH` it produces. Left to right is stated by
   * omission.
   */
  smartArtDirection?: "rtl";
  /**
   * SmartArt Design > Change Colors: the colour variation the diagram wears
   * (Office's id, `accent1_1` …; the default `accent1_2` is stated by
   * omission). A property of the DIAGRAM like {@link smartArtStyle}: the
   * Office-defined layouts are painted from that variation's definition,
   * which states every shape's AND every label's colours, so a relayout
   * paints them again as Office does (Colored Outline's dark words on its
   * white shapes, not the default variation's white).
   */
  smartArtColors?: string;
  /**
   * The Office ALIAS this diagram is laid out from (`smartArtLayoutUrn`,
   * `lib/office-smartart/native-layouts`): Office's urn for a layout that
   * stands for {@link smartArtKind}'s catalog layout, such as hProcess9 for a
   * Basic Process. Rides with the identity like {@link smartArtDirection};
   * absent on every diagram drawn as its catalog layout. On a diagram laid
   * out from its file's OWN layout definition (`office-definition`,
   * SMARTART-GAP.md §8.6) it is that definition's urn.
   */
  smartArtLayoutUrn?: string;
  /**
   * The sha256 of the layout part an `office-definition` diagram was laid out
   * from (`smartArtLayoutUrn`'s part in the document's own package): every
   * rebuild borrows that very part again, and the save writes the diagram only
   * over it. Present on such a diagram only.
   */
  smartArtLayoutSha256?: string;
}

export const DOCUMENT_INK_TOOLS = ["pen", "highlighter", "pencil"] as const;
export type DocumentInkTool = (typeof DOCUMENT_INK_TOOLS)[number];

export function normalizeDocumentInkTool(
  value: unknown,
): DocumentInkTool | undefined {
  return (DOCUMENT_INK_TOOLS as readonly unknown[]).includes(value)
    ? (value as DocumentInkTool)
    : undefined;
}

/** Word's eight pen effects (Draw tab, a pen's menu: Rainbow ... Bronze). */
export const DOCUMENT_INK_EFFECTS = OFFICE_INK_EFFECTS;
/**
 * A pen effect: one of Word's eight, one of MRAI's presets, or a custom
 * effect (`custom-<kind>-<rrggbb>…`); see `lib/office-ink-effects`.
 */
export type DocumentInkEffect = InkEffectId;
export { INK_EFFECT_MAX_CHARS, isInkEffect };

export function normalizeDocumentInkEffect(
  value: unknown,
): DocumentInkEffect | undefined {
  return normalizeInkEffect(value);
}

/**
 * Which ends of a stroke a point eraser cut: `"start"` is its first stored
 * point, `"end"` its last. A cut end stops FLAT, exactly where the eraser's
 * edge was; an end the pen made keeps the pen's round tip. Absent on every
 * stroke no point eraser has cut (and on every stroke stored before cut ends
 * existed): both ends round. Shared by the page, the worksheet and the deck.
 */
export const INK_CUT_ENDS = ["start", "end", "both"] as const;
export type InkCutEnds = (typeof INK_CUT_ENDS)[number];

/**
 * A stroke's pressures as the shape stores them: one integer 0..100 per
 * sample of its outline, space-separated (`inkPressure`). Absent on a stroke
 * drawn at one width (a mouse, a finger, a pen reporting none), so nothing
 * stored before pressure existed changes; the writers never read it (a file
 * line has one width), only the painter.
 */
export const MAX_INK_PRESSURE_CHARS = 4 * 1024;

export function normalizeInkPressure(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (trimmed === "" || trimmed.length > MAX_INK_PRESSURE_CHARS) return undefined;
  const tokens = trimmed.split(/\s+/);
  for (const token of tokens) {
    if (!/^(100|[1-9]?\d)$/.test(token)) return undefined;
  }
  return tokens.join(" ");
}

/** Draw ▸ Tape: whether a strip of tape covers the words under it or is peeled back. */
export const INK_TAPE_STATES = ["covered", "revealed"] as const;
export type InkTape = (typeof INK_TAPE_STATES)[number];

export function normalizeInkTape(value: unknown): InkTape | undefined {
  return (INK_TAPE_STATES as readonly unknown[]).includes(value) ? (value as InkTape) : undefined;
}

/** A stroke's place in a Draw-tab recording: the recording's id and the milliseconds into it. */
const INK_AUDIO = /^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}) (\d{1,9})$/;

export function normalizeInkAudio(value: unknown): string | undefined {
  return typeof value === "string" && INK_AUDIO.test(value) ? value : undefined;
}

/** The recording and the moment a stroke's `inkAudio` names. */
export function inkAudioPlace(value: unknown): { recordingId: string; atMs: number } | null {
  const match = typeof value === "string" ? INK_AUDIO.exec(value) : null;
  return match ? { recordingId: match[1]!, atMs: Number(match[2]) } : null;
}

export function normalizeInkCutEnds(value: unknown): InkCutEnds | undefined {
  return (INK_CUT_ENDS as readonly unknown[]).includes(value)
    ? (value as InkCutEnds)
    : undefined;
}

export type DocumentShapeAttrsInput =
  | Readonly<Record<string, unknown>>
  | DocumentShapeAttrs
  | null
  | undefined;

export interface DocumentShapeGroupMember {
  readonly memberId: string;
  /** Stable DrawingML cNvPr identity. Required by native connectors. */
  readonly drawingId?: number;
  /**
   * A shape, a connector, or a PICTURE.
   *
   * `CT_WordprocessingGroup` holds `pic:pic` in the same unbounded
   * `xsd:choice` as `wps:wsp` ([MS-ODRAWXML] 2.16.3.2), so a picture inside a
   * group is a MEMBER of the group: it sits in the same child coordinate
   * space, at its own place in document order, and moves, scales, rotates,
   * flips and z-orders with the group -- which is exactly what Microsoft
   * Support promises for "Group or ungroup shapes, pictures, or other
   * objects". Drawing such a picture beside its group instead
   * (`document_shape_group_pictures_apart`) said where it belonged and then
   * put it somewhere else (SHAPE-IMPORT-AUDIT.md §58).
   *
   * A picture member draws {@link src} clipped to its own `preset` frame, with
   * {@link pictureCropTop} and friends as its `a:srcRect`. It is not a shape
   * with a picture pasted on: it carries no label and no paint.
   */
  readonly kind?: "shape" | "connector" | "picture" | "chart";
  /**
   * A chart member's chart, as the page's chart node states one: a chart Word
   * grouped with other objects (a `wpg:graphicFrame`, corpus R48-D09). Drawn
   * from these attrs; the package keeps the chart's parts.
   */
  readonly chart?: Readonly<Record<string, unknown>>;
  /**
   * A picture member's `a:srcRect`, in the SAME four keys and the same whole
   * percent an image node already spells its crop with. Absent means identity,
   * which is what nearly every picture states.
   */
  readonly pictureCropTop?: number;
  readonly pictureCropRight?: number;
  readonly pictureCropBottom?: number;
  readonly pictureCropLeft?: number;
  readonly connectorPreset?: "straight" | "elbow";
  readonly startMemberId?: string;
  readonly startConnectionSite?: number;
  readonly endMemberId?: string;
  readonly endConnectionSite?: number;
  readonly preset: DocumentShapePreset;
  /**
   * The member's own outline, exactly as {@link DocumentShapeAttrs.customPath}
   * — the SmartArt engine draws several layouts as real paths, and a member
   * that only named its nearest preset drew a different figure than the same
   * diagram does on a slide. A connector never carries one: its geometry is
   * its route.
   */
  readonly customPath?: PresentationCustomPath;
  /** Its paths' own paints, as {@link DocumentShapeAttrs.customPathPaints}. */
  readonly customPathPaints?: readonly OfficeCustomPathPaint[];
  /**
   * A stroke the pen drew, grouped with others: still ink (its tool, effect,
   * cut ends and pressures, as {@link DocumentShapeAttrs.inkTool} and the
   * rest), so Ungroup gives back a stroke the lasso, eraser and replay know
   * as ink (live Word audit WUA 092).
   */
  readonly inkTool?: DocumentInkTool;
  readonly inkEffect?: DocumentInkEffect;
  readonly inkCutEnds?: InkCutEnds;
  readonly inkPressure?: string;
  readonly inkAudio?: string;
  readonly inkTape?: InkTape;
  /**
   * A connector ROUTE member's dragged handles, as `a:avLst` holds them — the
   * shape's own {@link DocumentShapeAttrs.adjustments}. An elbow is an open
   * route, which no outline can carry, so it keeps its preset and its handle:
   * Word's VML elbow brackets turn back past an end with the bend hundreds of
   * widths away. The package's, never edited on the page (the server's
   * `UniversalDocumentShapeGroupMember.adjustments`).
   */
  readonly adjustments?: Readonly<Record<string, number>>;
  /**
   * The member's own line ends (`a:ln/a:headEnd|a:tailEnd`), as on a lone
   * shape ({@link DocumentShapeAttrs.headEnd}): a line or connector inside a
   * Word group keeps its arrowheads through the page.
   */
  readonly headEnd?: PresentationConnectorEnd;
  readonly tailEnd?: PresentationConnectorEnd;
  readonly headEndWidth?: PresentationConnectorEndSize;
  readonly headEndLength?: PresentationConnectorEndSize;
  readonly tailEndWidth?: PresentationConnectorEndSize;
  readonly tailEndLength?: PresentationConnectorEndSize;
  readonly xEmu: number;
  readonly yEmu: number;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly relativeHeight?: number;
  readonly lockAspectRatio?: boolean;
  /**
   * `a:spLocks@noEditPoints`: Edit Points is unavailable on this member (a
   * Word-drawn freeform). Display/custody only; the writer states it back.
   */
  readonly lockEditPoints?: true;
  /**
   * Which of the file's groups this object is a child of, read flat: "" the
   * group itself, "2" the nested group that is its third child, "2/0" that
   * one's first. Word moves an object within its own group only (the
   * Selection Pane); absent on a group drawn here or not read so.
   */
  readonly groupPath?: string;
  /**
   * The member's own story when it holds a table or a picture its words
   * model cannot: Word's text box story as the header band's rich blocks
   * (`apps/api/.../projection._member_text_story`). Drawn by the band's
   * renderer; the server validates and canonicalizes it, and the member's
   * words are not edited while it is there (the writer writes the story).
   */
  readonly textStory?: readonly Readonly<Record<string, unknown>>[];
  readonly isTextBox?: boolean;
  readonly text: string;
  /** See {@link DocumentShapeTextBody}; a member's own text body. */
  readonly textBody?: DocumentShapeTextBody | null;
  





  readonly textRuns?: readonly TextRun[];
  /**
   * The words' text effects (an outline, a shadow, a reflection, a glow),
   * stated as a body run's `textEffects` mark states them (the gallery axes
   * plus Word's own statement) and written on every run of the member. Read
   * off a Word group member's runs when they agree; the editor draws it with
   * the body's own `textEffectsCss`.
   */
  readonly textEffects?: Readonly<Record<string, unknown>>;
  /**
   * Format Shape > Fill > Pattern fill on the member: the shared pattern
   * (`office-chart-shape-paint`, a preset and its two colours); `fill` is its
   * foreground and the member's `opacity` the transparency of both inks.
   */
  readonly fillPattern?: OfficeChartFillPattern;
  readonly fill: string;
  /**
   * The quick style's paint, in exactly the vocabulary the Presentation block
   * schema and both writers already share: `OfficeShapeGradient` (CSS degrees,
   * 180 = top→bottom, what `officeGradientXml` turns into `a:lin
   * ang="5400000"`) and `OfficeShapeShadow` (points, straight down). A member
   * with a gradient keeps `fill` as the ramp's first stop so a host that only
   * understands a solid fill still draws the right colour.
   */
  readonly fillGradient?: OfficeShapeGradient;
  readonly shapeShadow?: OfficeShapeShadow;
  /**
   * The member's own Shape Effects (glow, reflection, soft edges, an inner
   * shadow), in the model a leaf shape's `shapeEffects` keeps: drawn when the
   * group is expanded into its shapes, written back by the shared writer
   * (corpus R114-D11: a glowing star in a drawing canvas).
   */
  readonly shapeEffects?: OfficePictureEffects;
  /**
   * How opaque the member is drawn, as a PERCENTAGE (100 = opaque, absent =
   * opaque): the shared shape model's `opacity` (`normalizeShapeOpacity`), the
   * value a slide block and a sheet shape already carry. A SmartArt layout
   * draws its washes this way -- a Venn's overlapping discs, a Grouped List's
   * plate behind its items, a Funnel's vessel -- and Office draws them
   * translucent (`a:alpha` on the fill); a Docs member without it painted them
   * as solid blocks that hid what they hold. Only a value below 100 is kept.
   */
  readonly opacity?: number;
  










  readonly shape3d?: SmartArtShape3d;
  readonly scene3d?: SmartArtScene3d;
  /**
   * The box {@link scene3d} turns this member about, in the group's child
   * space, when the camera is a NESTED group's: Word turns that nested group
   * as one object about its own centre (group-scene3d-probe2 W2). Absent is
   * the whole group's box, as for a camera on the group itself.
   */
  readonly scene3dFrame?: DocumentShapeSceneFrame;
  /**
   * `a:grpFill`: the member wears the part of its group's fill
   * ({@link DocumentShapeAttrs.groupFill}) under its own box, worked out when
   * it is drawn ({@link documentShapeMemberGroupPaint}); its `fill` is only
   * the colour a host drawing a solid falls back to.
   */
  readonly fillFromGroup?: true;
  /**
   * The NESTED group's fill the member wears, when it is a nested group's
   * (read flat, its fill travels with its leaves), and that nested group's
   * box in the group's child space ({@link groupFillFrame}; alone, the box
   * the group's own fill is laid over). Absent: the group's fill over the
   * group's child space.
   */
  readonly groupFill?: DocumentShapeGroupFill;
  readonly groupFillFrame?: DocumentShapeSceneFrame;
  readonly textShape3d?: SmartArtTextShape3d;
  /** The member text's WordArt transform; see {@link DocumentShapeAttrs.textWarp}. */
  readonly textWarp?: OfficeTextWarp;
  readonly outline: string;
  readonly outlineWidthPt: number;
  readonly outlineStyle: StrokeStyle;
  /** The member line's exact dash preset (System Dash, Long Dash...); see {@link DocumentShapeAttrs.outlineDash}. */
  readonly outlineDash?: OfficeStrokeDash;
  /**
   * `a:ln@algn` on the MEMBER's own line. A member is a shape and owns its own
   * `a:ln` ([MS-ODRAWXML] 2.16.3.2), so an inset pen inside a group is the
   * member's statement, not the group's. Absent (the common case, and every
   * page stored before audit 68) draws centred and exports no attribute.
   */
  readonly outlineAlign?: OfficeStrokeAlign;
  /** The member line's own compound type, cap, join and miter limit. */
  readonly outlineCompound?: OfficeStrokeDetail["outlineCompound"];
  readonly outlineCap?: OfficeStrokeDetail["outlineCap"];
  readonly outlineJoin?: OfficeStrokeDetail["outlineJoin"];
  readonly outlineMiterLimit?: number;
  /**
   * Shape Outline > Sketched on the member's own line (a canvas of sketched
   * boxes, corpus R460-D01), drawn as a leaf shape's is. Absent on every
   * group stored before members carried it.
   */
  readonly outlineSketch?: OfficeLineSketch;
  readonly effect?: DocumentShapeEffectPreset;
  /**
   * WordArt on the member's words (Shape Format > WordArt Styles, Text Fill,
   * Text Outline, Text Effects), the same paint a leaf shape carries
   * ({@link DocumentShapeAttrs.textPaint}). Word puts it on every member's
   * runs when the style is chosen with the whole SmartArt graphic selected.
   */
  readonly textPaint?: OfficeTextPaint;
  readonly textFill?: string;
  readonly textOutline?: string;
  readonly textOutlineWidthPt?: number;
  readonly textEffect?: DocumentShapeEffectPreset;
  readonly textFontSizePt?: number;
  readonly textFontFamily?: string;
  readonly textBold?: boolean;
  readonly textItalic?: boolean;
  readonly textUnderline?: DocumentShapeTextUnderline;
  readonly textStrike?: boolean;
  readonly textDirection?: DocumentShapeTextDirection;
  readonly textVerticalAlign?: DocumentShapeTextVerticalAlign;
  readonly textHorizontalAlign?: DocumentShapeTextHorizontalAlign;
  readonly textWrap?: boolean;
  /**
   * Word's "Resize shape to fit text" on the member (`a:spAutoFit` in its own
   * `wps:bodyPr`): Word grows the member to its words, and the group and the
   * group's line with it (docx-look probe grp1). Drawn so; never written.
   */
  readonly textAutofit?: "shape";
  readonly altText?: string;
  /** The member's own hyperlink (Word's `cNvPr/a:hlinkClick`) and its
   * ScreenTip, carried as the server read them. */
  readonly linkHref?: string;
  readonly linkTitle?: string;
  readonly wrapMode?: DocumentShapeWrapMode;
  readonly linkedTextBoxStoryId?: number;
  readonly linkedTextBoxSequence?: number;
  /**
   * SmartArt identity, mirrored verbatim from the engine node this member
   * draws ({@link DocumentShapeAttrs.smartArtKind}). Present only on a group
   * that IS a catalog SmartArt; a hand-built shape group never carries it.
   */
  readonly smartArtRole?: DocumentShapeSmartArtRole;
  readonly smartArtItemId?: string;
  readonly smartArtItemIndex?: number;
  readonly smartArtLevel?: number;
  /** The item a title box belongs to (Name and Title), on a `connector` member. */
  readonly smartArtTitleIndex?: number;
  /**
   * The item a shape drawn for an item without its words belongs to
   * (Vertical Circle List's circles), on a `connector` member: Delete on it
   * takes that item out (`lib/office-smartart/model`).
   */
  readonly smartArtOwnerIndex?: number;
  /**
   * An organization chart box's Layout (Standard / Both / Left / Right
   * Hanging: Office's `hierBranch`), on its `node` member
   * (`lib/office-smartart/model` `SmartArtHierBranch`).
   */
  readonly smartArtHierBranch?: "std" | "hang" | "l" | "r";
  /**
   * This member is the diagram's PICTURE SLOT: a picture layout's `image`
   * node, projected as the neutral shape the worksheet host projects it as
   * (`SpreadsheetSmartArtMember.smartArtPicturePlaceholder`). A Word group
   * member holds no picture, so the slot is a shape — and saying so is what
   * lets the four picture layouts import as themselves instead of being
   * refused, and what a later picture-custody path would find them by.
   * Only a `node` of an identified diagram may declare it.
   */
  readonly smartArtPicturePlaceholder?: true;
  /**
   * The picture IN that slot, as the inline data URL every host spells a
   * SmartArt picture with (`office-smartart/picture`).
   *
   * A Word group member is a `wps:wsp` and holds no picture of its own, which
   * is why the slot was projected as a neutral shape in the first place. It is
   * still a shape — this field is not a second kind of member — but the slot
   * now REMEMBERS what the author (or the imported Office diagram) put in it,
   * so the canvas can draw it and the DiagramML writer can put it back on the
   * data point as an `a:blipFill` (SMARTART-IMPORT.md §5.3, measured: Word
   * keeps the blip through its own re-save and redraws it in the cache IT
   * regenerates).
   *
   * Two members may carry one, and no others. A
   * {@link smartArtPicturePlaceholder} slot, as above; and a
   * {@link kind} `"picture"` member, whose `src` is not a slot's contents but
   * the member's own `pic:blipFill` bytes. Anywhere else it would be a picture
   * no exporter has anywhere to write it to.
   */
  readonly src?: string;
  /**
   * The colours on this member are the AUTHOR's, not the catalog palette's.
   *
   * The export writes a per-point `dgm:spPr` only for colours it can tell
   * apart from the palette, and the only tell it had was disagreement: node
   * colours that DIFFER inside one diagram are somebody's choice, colours that
   * AGREE are the catalog's (SMARTART-IMPORT.md §5.3). So a diagram whose every
   * node was painted ONE colour exported without its colours — and since the
   * reader has had no such doubt (§5.12), it reopened in that colour and
   * re-exported without it.
   *
   * This is the page saying which it was. Optional and only ever `true`, only
   * ever on a `node` of an identified diagram: a page stored before it simply
   * does not carry it, which is the §5.3 heuristic exactly as it was.
   */
  readonly smartArtAuthorPaint?: true;
  /**
   * This member's text size was set by hand (Home's size box, Grow / Shrink
   * Font): it draws at `textFontSizePt` and leaves the graphic's shared fit
   * (Microsoft Support: "if you have previously customized the text size in
   * any shape, then the text in that shape will not resize"). Without it a
   * hand-set size equal to the default 14 pt read as never set and kept
   * shrinking with the others. Optional, only ever `true`, only on an item's
   * node or label.
   */
  readonly smartArtAuthorTextSize?: true;
  /**
   * Home > Font on this item's words: the face in `textFontFamily` is the
   * AUTHOR's, at any value. Without it the shape model's default face
   * (Arial, which also means "no face of its own") read as never set, and a
   * label the author set in Arial was drawn and exported in the document
   * theme's body font instead. Optional, only ever `true`, only on an item's
   * node or label.
   */
  readonly smartArtAuthorTextFont?: true;
  /**
   * SmartArt Format > Change Shape on this one shape: the preset the AUTHOR
   * gave it, which IS the member's own {@link preset} (Word keeps it as
   * `a:prstGeom` on the presentation point that draws the shape, where the
   * DOCX writer puts it and the reader finds it). A rebuild of the graphic
   * (Add Shape, Promote, Change Colors) draws the item in it again; Reset
   * Shape, Reset Graphic and a layout switch go back to the layout's own
   * shape. Optional, only on an item's node, dropped when it is not the
   * member's preset.
   */
  readonly smartArtAuthorShape?: DocumentShapePreset;
  /**
   * Picture Format on the picture in this SmartArt slot
   * (`lib/office-smartart/slot-picture-format`): Crop as the fill's
   * `a:srcRect` (the editors' signed `CropRect`), and the Picture Styles
   * gallery's pick. Only on a picture slot holding a picture; dropped
   * anywhere else.
   */
  readonly smartArtPictureCrop?: CropRect;
  readonly smartArtPictureStyle?: OfficePictureStyleId;
  /**
   * The author moved or resized this shape: its scale and the move of its
   * centre against its layout box, in Office's `dgm:prSet` unit
   * (`lib/office-smartart/custom-geometry.ts`). Only on an item's shape.
   */
  readonly smartArtCustom?: SmartArtCustomFactors;
  /**
   * Where this shape's words are laid out when it is not its own box: the
   * diagram's text box for the shape (Office's `dsp:txXfrm`), relative to the
   * member's box in Office's percentage unit (`lib/office-smartart/text-box`).
   * A Segmented Cycle wedge's words sit in the wedge, not at the centre of
   * the disc its box spans. Only on an item's node or label.
   */
  readonly smartArtTextBox?: SmartArtTextBox;
  /**
   * The line spacing the author gave this item's words (a converted point's
   * `a:lnSpc/a:spcPct`, Home > Line Spacing), as a multiple of a single line;
   * absent, the layout's own 90%. On an item's node or label only.
   */
  readonly smartArtLineSpacing?: number;
  /**
   * The item's words are set in capitals (a converted point's `cap="all"`,
   * Font > All Caps): drawn so, the words themselves untouched, and stated on
   * the point again by the DOCX writer (`authorTextCapitals`).
   */
  readonly smartArtTextCapitals?: true;
  /**
   * The size, in points, Office's autofit gave this item's words when the
   * diagram was laid out from Office's own definition (a converted Word
   * diagram): the graphic's one size in that frame, which the document draws
   * the words at, shrinking them with a smaller frame and never past it. Its
   * `smartArtTextBox` is then the words' area with Office's insets already
   * off. Drawing only: the DOCX writer never states it (no `sz`), and a
   * diagram MRAI inserted has none. On an item's node or label only.
   */
  readonly smartArtLabelSizePt?: number;
  /**
   * The item is an organization chart's ASSISTANT (`dgm:pt type="asst"`),
   * on its node and label only, only ever `true`.
   */
  readonly smartArtAssistant?: true;
  /**
   * The item's layout does not draw it (Office's "unused text does not
   * appear, but remains available if you switch layouts"): the member keeps
   * the item in the outline, is never painted or clicked, and the export
   * writes its data point with no drawing shape. Only on an item's node or
   * label, only ever `true`.
   */
  readonly smartArtHidden?: true;
}

/**
 * The role a SmartArt group member plays, exactly the engine's own vocabulary
 * (`lib/office-smartart/engine`). A Docs member merges an item's node and its
 * label into ONE shape, so `label` appears only for a label with no shape of
 * its own.
 */
export type DocumentShapeSmartArtRole = "node" | "label" | "connector";

/** Office's own bound, and the PPT identity writer's (`smartart_identity.py`). */
export const DOCUMENT_SMARTART_MAX_IDENTITY_ITEMS = 256;
export const DOCUMENT_SMARTART_MAX_LEVEL = 3;

/**
 * A catalog layout id / engine kind, in the one spelling every host persists.
 * Deliberately a token test rather than a catalog lookup: this module is the
 * document's shape model and must not depend on the SmartArt catalog.
 */
function smartArtKindToken(value: unknown): string | null {
  return typeof value === "string" && /^[a-z][a-z0-9-]{0,79}$/.test(value)
    ? value
    : null;
}

/** A group id or an outline item id (an Office `modelId` GUID, or a uuid). */
function smartArtIdentityToken(value: unknown): string | null {
  return typeof value === "string" &&
    /^[A-Za-z0-9][A-Za-z0-9._:{}-]{0,127}$/.test(value)
    ? value
    : null;
}

function smartArtOrdinal(value: unknown, max: number): number | null {
  return typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 0 &&
    value <= max
    ? value
    : null;
}

/**
 * The identity a member declares, or `null` when it declares none. Fail
 * closed: a partial identity (a role that owns an item but no item id, an
 * out-of-range index) is NOT an identity, so a malformed page atom can never
 * make a shape group look like a SmartArt it is not. A picture slot claimed by
 * anything but an item-owning node is refused the same way.
 */
export function documentShapeMemberSmartArtIdentity(
  member: Readonly<Record<string, unknown>>,
): Pick<
  DocumentShapeGroupMember,
  | "smartArtRole"
  | "smartArtItemId"
  | "smartArtItemIndex"
  | "smartArtLevel"
  | "smartArtPicturePlaceholder"
  | "smartArtAuthorPaint"
  | "smartArtAuthorTextSize"
  | "smartArtAuthorTextFont"
  | "smartArtAuthorShape"
  | "smartArtPictureCrop"
  | "smartArtPictureStyle"
  | "smartArtCustom"
  | "smartArtTextBox"
  | "smartArtAssistant"
  | "smartArtHidden"
  | "smartArtTitleIndex"
  | "smartArtOwnerIndex"
  | "smartArtHierBranch"
  | "src"
> | null {
  const role = member.smartArtRole;
  if (role !== "node" && role !== "label" && role !== "connector") return null;
  // A `kind: "picture"` member's `src` is its OWN bytes, read by
  // `documentShapeMemberPicture`. This function still owns every other `src`,
  // which is a diagram slot's contents and needs the identity to go with it.
  if (member.kind === "picture") return null;
  const picture = member.smartArtPicturePlaceholder;
  if (picture !== undefined && (picture !== true || role !== "node")) return null;
  // The picture IN the slot. Refused — not dropped — anywhere but on a slot,
  // and refused when it is not the one inline spelling every host writes: a
  // member stating a picture no writer could export is a malformed page atom,
  // exactly like a partial identity, and the whole group falls back to an
  // ordinary shape group rather than drawing a picture the file will not hold.
  const src = member.src;
  if (src !== undefined && (picture !== true || smartArtPictureSrc(src) === null)) {
    return null;
  }
  // A member's paint is the author's only where the member has paint of its
  // own to claim, which is a node; a label and a connector are MRAI's own
  // drawing and have no data point for the writer to put colours on.
  const authorPaint = member.smartArtAuthorPaint;
  if (authorPaint !== undefined && (authorPaint !== true || role !== "node")) {
    return null;
  }
  // A size the author set on this item's words. Dropped, not refused, when
  // it is not the one spelling: the words then fit with the others.
  const authorTextSize = member.smartArtAuthorTextSize === true && role !== "connector";
  // …and the face the author set on them, the same way.
  const authorTextFont = member.smartArtAuthorTextFont === true && role !== "connector";
  // Picture Format on the slot's picture. Dropped, not refused, anywhere but
  // on a slot holding a picture, or when it is not one.
  const pictureFormat =
    picture === true && typeof src === "string" ? smartArtSlotPictureFormatOf(member) : null;
  // The shape the author gave this item (Change Shape). Dropped, not refused,
  // when it is not the preset the member draws: the shape is then the
  // layout's, which is what a relayout would draw anyway.
  const authorShape =
    role === "node" &&
    picture !== true &&
    typeof member.smartArtAuthorShape === "string" &&
    member.smartArtAuthorShape === member.preset &&
    isDocumentShapePreset(member.smartArtAuthorShape)
      ? member.smartArtAuthorShape
      : null;
  // An item's shape the author moved or resized. A record that is not one
  // is dropped rather than refused: the shape is still the item's, only not
  // customized, which is what Office makes of it too.
  const custom = role === "connector" ? null : readSmartArtCustom(member.smartArtCustom);
  // The shape's own text box. Dropped, not refused, when it is not one, as
  // a customization is: the words then fill the shape's box.
  const textBox = role === "connector" ? null : readSmartArtTextBox(member.smartArtTextBox);
  // The line spacing the author gave this item's words. Dropped, not
  // refused, when it is not one: the words then stand at the layout's.
  const lineSpacing =
    role !== "connector" &&
    typeof member.smartArtLineSpacing === "number" &&
    Number.isFinite(member.smartArtLineSpacing) &&
    member.smartArtLineSpacing >= 0.05 &&
    member.smartArtLineSpacing <= 9.6
      ? member.smartArtLineSpacing
      : null;
  // Office's fitted size for the item's words; dropped, not refused, when it
  // is not a size: the words then take the document's own fit.
  const labelSize =
    role !== "connector" &&
    typeof member.smartArtLabelSizePt === "number" &&
    Number.isFinite(member.smartArtLabelSizePt) &&
    member.smartArtLabelSizePt >= 1 &&
    member.smartArtLabelSizePt <= 4000
      ? member.smartArtLabelSizePt
      : null;
  if (role === "connector") {
    // A title box names the item it titles (Name and Title); dropped, not
    // refused, when it is not an item index.
    const titled = smartArtOrdinal(member.smartArtTitleIndex, DOCUMENT_SMARTART_MAX_IDENTITY_ITEMS - 1);
    // The item a shape drawn for it without its words belongs to; dropped,
    // not refused, when it is not an item index (Delete then refuses it).
    const owner = titled === null
      ? smartArtOrdinal(member.smartArtOwnerIndex, DOCUMENT_SMARTART_MAX_IDENTITY_ITEMS - 1)
      : null;
    return member.smartArtItemId === undefined &&
      member.smartArtItemIndex === undefined &&
      member.smartArtLevel === undefined
      ? {
          smartArtRole: role,
          ...(titled !== null ? { smartArtTitleIndex: titled } : {}),
          ...(owner !== null ? { smartArtOwnerIndex: owner } : {}),
          // A title the layout does not draw (an item's transition words).
          ...(titled !== null && member.smartArtHidden === true
            ? { smartArtHidden: true as const }
            : {}),
        }
      : null;
  }
  const itemId = smartArtIdentityToken(member.smartArtItemId);
  const itemIndex = smartArtOrdinal(
    member.smartArtItemIndex,
    DOCUMENT_SMARTART_MAX_IDENTITY_ITEMS - 1,
  );
  const level = smartArtOrdinal(member.smartArtLevel, DOCUMENT_SMARTART_MAX_LEVEL);
  return itemId !== null && itemIndex !== null && level !== null
    ? {
        smartArtRole: role,
        smartArtItemId: itemId,
        smartArtItemIndex: itemIndex,
        smartArtLevel: level,
        ...(picture === true ? { smartArtPicturePlaceholder: true as const } : {}),
        ...(typeof src === "string" ? { src } : {}),
        ...(authorPaint === true ? { smartArtAuthorPaint: true as const } : {}),
        ...(authorTextSize ? { smartArtAuthorTextSize: true as const } : {}),
        ...(authorTextFont ? { smartArtAuthorTextFont: true as const } : {}),
        ...(authorShape ? { smartArtAuthorShape: authorShape } : {}),
        ...(pictureFormat ? smartArtSlotPictureFields(pictureFormat) : {}),
        ...(custom ? { smartArtCustom: custom } : {}),
        ...(textBox ? { smartArtTextBox: textBox } : {}),
        ...(lineSpacing !== null ? { smartArtLineSpacing: lineSpacing } : {}),
        ...(member.smartArtTextCapitals === true ? { smartArtTextCapitals: true as const } : {}),
        ...(labelSize !== null ? { smartArtLabelSizePt: labelSize } : {}),
        ...(member.smartArtAssistant === true && level > 0
          ? { smartArtAssistant: true as const }
          : {}),
        ...(member.smartArtHidden === true ? { smartArtHidden: true as const } : {}),
        // An org chart box's Layout; dropped, not refused, when it is not one.
        ...(role === "node" && isSmartArtHierBranch(member.smartArtHierBranch)
          ? { smartArtHierBranch: member.smartArtHierBranch }
          : {}),
      }
    : null;
}

/** A picture member's `a:srcRect`, in the page's own four keys. */
export const DOCUMENT_SHAPE_MEMBER_PICTURE_CROP_FIELDS = [
  "pictureCropTop",
  "pictureCropRight",
  "pictureCropBottom",
  "pictureCropLeft",
] as const;

/**
 * What a `kind: "picture"` member declares, or `null` when it declares a value
 * this model cannot prove.
 *
 * Fail closed exactly the way the identity and the quick style do: a malformed
 * crop or a `src` that is not the one inline spelling every host writes is a
 * malformed page atom, so the whole group falls back to an ordinary shape
 * group rather than drawing a picture no file will hold. A picture member with
 * NO `src` is legal and draws its frame -- that is a picture whose bytes did
 * not fit the page's budget, and the import receipt names it.
 */
function documentShapeMemberPicture(
  member: Readonly<Record<string, unknown>>,
): Pick<
  DocumentShapeGroupMember,
  | "src"
  | (typeof DOCUMENT_SHAPE_MEMBER_PICTURE_CROP_FIELDS)[number]
> | null {
  const src = member.src;
  if (
    src !== undefined &&
    (typeof src !== "string" ||
      // The inline spelling, or the stored file's own picture named by
      // digest (an import in picture-reference mode names a group's picture
      // as it names one beside it).
      !(src.startsWith("data:image/") || pictureObjectTokenSha(src) !== null) ||
      src.length > 4_000_000)
  ) {
    return null;
  }
  const crop: Record<string, number> = {};
  for (const key of DOCUMENT_SHAPE_MEMBER_PICTURE_CROP_FIELDS) {
    const value = member[key];
    if (value === undefined) continue;
    if (
      typeof value !== "number" ||
      !Number.isFinite(value) ||
      // Negative is an OUTSET (the source reaches past the picture: it is
      // drawn smaller than its frame), bounded as an image's crop is.
      value < -400 ||
      value >= 100 ||
      // `a:srcRect`'s own unit: a thousandth of a percent.
      Math.abs(value * 1_000 - Math.round(value * 1_000)) > 1e-6
    ) {
      return null;
    }
    if (value !== 0) crop[key] = value;
  }
  // A picture's `a:srcRect` keeps a window on the source: what the two
  // opposite edges crop leaves something.
  if (
    (crop.pictureCropLeft ?? 0) + (crop.pictureCropRight ?? 0) >= 100 ||
    (crop.pictureCropTop ?? 0) + (crop.pictureCropBottom ?? 0) >= 100
  ) {
    return null;
  }
  return { ...(typeof src === "string" ? { src } : {}), ...crop };
}

/** Every field a SmartArt quick style (or its layout's wash) writes onto ONE group member. */
export const DOCUMENT_SHAPE_MEMBER_STYLE_FIELDS = [
  "fillGradient",
  "shapeShadow",
  "shape3d",
  "scene3d",
  "scene3dFrame",
  "textShape3d",
  "opacity",
] as const;

/** A box in a group's child space, in EMU (see {@link DocumentShapeGroupMember.scene3dFrame}). */
export interface DocumentShapeSceneFrame {
  readonly xEmu: number;
  readonly yEmu: number;
  readonly widthEmu: number;
  readonly heightEmu: number;
}

function documentShapeSceneFrame(value: unknown): DocumentShapeSceneFrame | null | undefined {
  if (value === undefined) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const frame = value as Record<string, unknown>;
  const keys = Object.keys(frame).sort().join(",");
  if (keys !== "heightEmu,widthEmu,xEmu,yEmu") return null;
  const within = (key: string, min: number) =>
    Number.isSafeInteger(frame[key]) && (frame[key] as number) >= min && (frame[key] as number) <= DOCUMENT_SHAPE_MAX_EDGE_EMU;
  if (!within("xEmu", -DOCUMENT_SHAPE_MAX_EDGE_EMU) || !within("yEmu", -DOCUMENT_SHAPE_MAX_EDGE_EMU)) return null;
  if (!within("widthEmu", 1) || !within("heightEmu", 1)) return null;
  return {
    xEmu: frame.xEmu as number,
    yEmu: frame.yEmu as number,
    widthEmu: frame.widthEmu as number,
    heightEmu: frame.heightEmu as number,
  };
}

/**
 * A group's own fill, worn by its `a:grpFill` members: one paint in the
 * shared model (`office-shape-group-fill`, the API's
 * `office_shape_group.group_fill`). A Docs picture is a shape's picture fill
 * ({@link DocumentShapePictureFill}), stretched.
 */
export type DocumentShapeGroupFill =
  | { readonly gradient: OfficeShapeGradient }
  | { readonly picture: DocumentShapePictureFill }
  | { readonly pattern: OfficeChartFillPattern };

/** `value` as a {@link DocumentShapeGroupFill}, `undefined` when absent, null when it is not one. */
export function normalizeDocumentShapeGroupFill(value: unknown): DocumentShapeGroupFill | null | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) return null;
  const keys = Object.keys(value);
  if (keys.length !== 1) return null;
  const record = value as Record<string, unknown>;
  try {
    if (keys[0] === "gradient") {
      const gradient = normalizeOfficeShapeGradient(record.gradient);
      if (!gradient || gradient.path || gradient.tileRect || (gradient.flip !== undefined && gradient.flip !== "none")) return null;
      return { gradient };
    }
    if (keys[0] === "pattern") {
      const pattern = normalizeOfficeChartFillPattern(record.pattern);
      return pattern ? { pattern } : null;
    }
  } catch {
    return null;
  }
  if (keys[0] === "picture") {
    const picture = normalizeDocumentShapePictureFill(record.picture);
    return picture && picture.stretch === true && !picture.tile ? { picture } : null;
  }
  return null;
}

/**
 * A member's group-fill fields ({@link DocumentShapeGroupMember.fillFromGroup}
 * and the nested group's fill and frame beside it), `{}` for a member wearing
 * none, or null when they do not hold together (as the API's validator says).
 */
function documentShapeMemberGroupFill(
  member: Readonly<Record<string, unknown>>,
  kind: string,
): Pick<DocumentShapeGroupMember, "fillFromGroup" | "groupFill" | "groupFillFrame"> | null {
  const wears = member.fillFromGroup === true;
  if (member.fillFromGroup !== undefined && typeof member.fillFromGroup !== "boolean") return null;
  const nested = normalizeDocumentShapeGroupFill(member.groupFill);
  const frame = documentShapeSceneFrame(member.groupFillFrame);
  if (nested === null || frame === null) return null;
  if (!wears) return nested || frame ? null : {};
  if (kind !== "shape" || (nested && !frame)) return null;
  return {
    fillFromGroup: true,
    ...(nested ? { groupFill: nested } : {}),
    ...(frame ? { groupFillFrame: frame } : {}),
  };
}

/**
 * What a member wearing its group's fill draws of it, in the page's fill keys
 * (`fillGradient` / `fillPicture` / `fillPattern`): the group's `groupFill`
 * -- or the nested group's the member carries -- laid over that group's box
 * (`groupFillFrame`, else the group's child space), the member showing the
 * part under its own box AS IT IS NOW. Worked out in the group's page EMU,
 * where a gradient's aspect is the one drawn. Null when the member wears none
 * or its part cannot be said. `document_shape_member_group_paint`'s twin.
 */
export function documentShapeMemberGroupPaint(
  group: Pick<DocumentShapeAttrs, "groupFill" | "widthEmu" | "heightEmu" | "groupCoordinateWidthEmu" | "groupCoordinateHeightEmu" | "groupMembers">,
  member: Pick<DocumentShapeGroupMember, "fillFromGroup" | "groupFill" | "groupFillFrame" | "xEmu" | "yEmu" | "widthEmu" | "heightEmu"> &
    Partial<Pick<DocumentShapeGroupMember, "memberId">>,
): { fillGradient: OfficeShapeGradient } | { fillPicture: DocumentShapePictureFill } | { fillPattern: OfficeChartFillPattern } | null {
  if (member.fillFromGroup !== true) return null;
  const nested = member.groupFill;
  const fill = nested ?? group.groupFill;
  if (!fill) return null;
  const coordinateWidth = group.groupCoordinateWidthEmu ?? group.widthEmu;
  const coordinateHeight = group.groupCoordinateHeightEmu ?? group.heightEmu;
  if (!(coordinateWidth > 0) || !(coordinateHeight > 0)) return null;
  const scaleX = group.widthEmu / coordinateWidth;
  const scaleY = group.heightEmu / coordinateHeight;
  // The fill is laid over the UNION of the group's members' boxes as they
  // are now, not the box the group states (Word and PowerPoint alike,
  // group-fill-probe3 W1-W3 / Q1-Q4). A nested group's members are those
  // wearing the same nested fill; `member` stands for its own entry.
  type Box = Pick<DocumentShapeGroupMember, "fillFromGroup" | "groupFill" | "groupFillFrame" | "xEmu" | "yEmu" | "widthEmu" | "heightEmu">;
  let members: Box[] = (group.groupMembers ?? []).map((other) =>
    member.memberId !== undefined && other.memberId === member.memberId ? member : other,
  );
  if (!members.includes(member)) members.push(member);
  if (nested) {
    const key = JSON.stringify([nested, member.groupFillFrame ?? null]);
    members = members.filter(
      (other) => other.fillFromGroup === true && JSON.stringify([other.groupFill ?? null, other.groupFillFrame ?? null]) === key,
    );
  }
  const left = Math.min(...members.map((box) => box.xEmu));
  const top = Math.min(...members.map((box) => box.yEmu));
  const right = Math.max(...members.map((box) => box.xEmu + box.widthEmu));
  const bottom = Math.max(...members.map((box) => box.yEmu + box.heightEmu));
  const page = (x: number, y: number, w: number, h: number): OfficeGroupFillBox => [x * scaleX, y * scaleY, w * scaleX, h * scaleY];
  return officeGroupFillMemberPaint(
    fill,
    page(left, top, right - left, bottom - top),
    page(member.xEmu, member.yEmu, member.widthEmu, member.heightEmu),
  ) as { fillGradient: OfficeShapeGradient } | { fillPicture: DocumentShapePictureFill } | { fillPattern: OfficeChartFillPattern } | null;
}

/**
 * A member as it is DRAWN: wearing its part of its group's fill (worked out
 * now, from the boxes as they are) in place of its own fallback paint. Any
 * other member is returned as it is.
 */
export function documentShapeMemberWithGroupPaint(
  group: Parameters<typeof documentShapeMemberGroupPaint>[0],
  member: DocumentShapeGroupMember,
): DocumentShapeGroupMember {
  const paint = documentShapeMemberGroupPaint(group, member);
  if (!paint) return member;
  const { fillGradient: _gradient, fillPattern: _pattern, opacity: _opacity, ...rest } = member;
  return { ...rest, ...paint } as DocumentShapeGroupMember;
}

/**
 * A member's own transparency, as {@link DocumentShapeGroupMember.opacity}
 * keeps it: a whole percentage in `1..99`, or `undefined` for an opaque one.
 * Fully transparent is refused rather than kept, because a member drawn at
 * zero is a member nobody can see or select.
 */
export function documentShapeMemberOpacity(value: unknown): number | undefined | null {
  if (value === undefined) return undefined;
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  const rounded = Math.round(value);
  if (rounded !== value || rounded < 1 || rounded > 100) return null;
  return rounded === 100 ? undefined : rounded;
}

/**
 * The quick-style paint a member declares, or `null` when it declares a value
 * this model cannot prove.
 *
 * Fail closed exactly the way the identity does: a 3-D payload outside the
 * shared bounds (`isSmartArtShape3d` and friends, the browser mirror of the
 * API's `office_shape_paint` validators) is not "most of a style", it is a
 * malformed page atom, so the whole group is refused rather than drawn and
 * exported with half a style. `fillGradient`/`shapeShadow` arrive already
 * normalized by the leaf pass, which is the same bounds table the standalone
 * shape uses.
 */
function documentShapeMemberStyle(
  member: Readonly<Record<string, unknown>>,
  leaf: Pick<DocumentShapeAttrs, "fillGradient" | "shapeShadow">,
): Pick<
  DocumentShapeGroupMember,
  (typeof DOCUMENT_SHAPE_MEMBER_STYLE_FIELDS)[number]
> | null {
  // A member's solid and its group's camera are bounded by SHAPE, as the
  // API's readers are by the literal ECMA names (`office_shape_paint`): Word
  // writes any camera on a group (`wpg:grpSpPr/a:scene3d`, the whole group's
  // 3-D Rotation) and any bevel on a member, and the SmartArt Styles' six
  // cameras and four bevels refused every other group whole.
  for (const [key, guard] of [
    ["shape3d", isNativeShape3d],
    ["scene3d", isNativeScene3d],
    ["textShape3d", isSmartArtTextShape3d],
  ] as const) {
    if (member[key] !== undefined && !guard(member[key])) return null;
  }
  const opacity = documentShapeMemberOpacity(member.opacity);
  if (opacity === null) return null;
  // The frame only beside the camera it frames, as the API's validator says.
  const sceneFrame = documentShapeSceneFrame(member.scene3dFrame);
  if (sceneFrame === null || (sceneFrame && !isNativeScene3d(member.scene3d))) return null;
  return {
    ...(opacity !== undefined ? { opacity } : {}),
    ...(leaf.fillGradient ? { fillGradient: leaf.fillGradient } : {}),
    ...(leaf.shapeShadow ? { shapeShadow: leaf.shapeShadow } : {}),
    ...(isNativeShape3d(member.shape3d) ? { shape3d: member.shape3d } : {}),
    ...(isNativeScene3d(member.scene3d) ? { scene3d: member.scene3d } : {}),
    ...(sceneFrame ? { scene3dFrame: sceneFrame } : {}),
    ...(isSmartArtTextShape3d(member.textShape3d)
      ? { textShape3d: member.textShape3d }
      : {}),
  };
}

export interface DocumentConnectorRoute {
  readonly xEmu: number;
  readonly yEmu: number;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

function documentShapeConnectionPoint(
  member: DocumentShapeGroupMember,
  site: number,
): Readonly<{ x: number; y: number }> | null {
  if (!Number.isInteger(site) || site < 0 || site > 3) return null;
  const left = member.xEmu;
  const top = member.yEmu;
  const right = left + member.widthEmu;
  const bottom = top + member.heightEmu;
  if (site === 0) return { x: Math.floor((left + right + 1) / 2), y: top };
  if (site === 1) {
    return { x: right, y: Math.floor((top + bottom + 1) / 2) };
  }
  if (site === 2) {
    return { x: Math.floor((left + right + 1) / 2), y: bottom };
  }
  return { x: left, y: Math.floor((top + bottom + 1) / 2) };
}

export function documentShapeConnectorRoute(
  start: DocumentShapeGroupMember,
  startSite: number,
  end: DocumentShapeGroupMember,
  endSite: number,
): DocumentConnectorRoute | null {
  const from = documentShapeConnectionPoint(start, startSite);
  const to = documentShapeConnectionPoint(end, endSite);
  if (!from || !to) return null;
  return {
    xEmu: Math.round(Math.min(from.x, to.x)),
    yEmu: Math.round(Math.min(from.y, to.y)),
    // A connector is a route corner to corner in its box: its extent is the
    // box, zero across a level or upright one (a 16 px floor drew it slanted).
    widthEmu: Math.round(Math.abs(to.x - from.x)),
    heightEmu: Math.round(Math.abs(to.y - from.y)),
    flipH: from.x > to.x,
    flipV: from.y > to.y,
  };
}

export interface DocumentShapeGroupInput {
  readonly pos: number;
  readonly attrs: Readonly<Record<string, unknown>>;
  /**
   * A floating CHART grouped with the shapes (Word's Group takes a chart
   * that floats; Microsoft Support, "Group or ungroup shapes, pictures, or
   * other objects"): the chart as its member states one
   * (`DocumentShapeGroupMember.chart`). `attrs` then holds only its frame --
   * `xEmu` / `yEmu` / `widthEmu` / `heightEmu` and its `relativeHeight`.
   */
  readonly chart?: Readonly<Record<string, unknown>>;
}

/**
 * A chart's frame as the group plan reads a member: a clear rectangle with
 * no words and no paint, in the floating stack at the chart's own height.
 * The member it becomes draws the chart (`kind: "chart"`).
 */
function documentChartGroupFrameAttrs(
  frame: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  return {
    preset: "rect",
    xEmu: frame.xEmu,
    yEmu: frame.yEmu,
    widthEmu: frame.widthEmu,
    heightEmu: frame.heightEmu,
    rotation: 0,
    flipH: false,
    flipV: false,
    lockAspectRatio: false,
    relativeHeight: frame.relativeHeight,
    zOrderStackComplete: true,
    text: "",
    fill: "none",
    outline: "none",
    outlineWidthPt: 0,
    outlineStyle: "solid",
    source: null,
    ...(frame.wrapMode !== undefined ? { wrapMode: frame.wrapMode } : {}),
  };
}

type DocumentShapeLeafAttrs = Omit<
  DocumentShapeAttrs,
  "groupMembers" | "groupCoordinateWidthEmu" | "groupCoordinateHeightEmu"
>;

/**
 * The six optional line-end fields, held to the shared vocabularies. A value
 * outside them is DROPPED rather than failing the shape, exactly as a bad
 * custom outline is: the preset still names a real auto-shape to draw.
 */
function normalizeDocumentShapeLineEnds(
  attrs: Readonly<Record<string, unknown>>,
): Partial<DocumentShapeAttrs> {
  const fields: Record<string, string> = {};
  for (const [key, vocabulary] of [
    ["headEnd", PRESENTATION_CONNECTOR_ENDS],
    ["tailEnd", PRESENTATION_CONNECTOR_ENDS],
    ["headEndWidth", PRESENTATION_CONNECTOR_END_SIZES],
    ["headEndLength", PRESENTATION_CONNECTOR_END_SIZES],
    ["tailEndWidth", PRESENTATION_CONNECTOR_END_SIZES],
    ["tailEndLength", PRESENTATION_CONNECTOR_END_SIZES],
  ] as const) {
    const value = attrs[key];
    if (typeof value !== "string") continue;
    if (!(vocabulary as readonly string[]).includes(value)) continue;
    fields[key] = value;
  }
  return fields as Partial<DocumentShapeAttrs>;
}

/**
 * A GROUP MEMBER is held to no minimum edge, and that is deliberate rather
 * than an oversight. The floor buys a shape on the page the user can see and
 * grab; a member is neither — the page draws and grabs the GROUP, and the
 * member lives in the group's own child coordinate space, which is not the
 * page's. Word writes a diagram's separator rules as 2-EMU and 12_700-EMU
 * rectangles and no refinement of that child space can lift a 2-EMU edge to
 * 152_400 inside a box that must itself stay a legal shape edge. The canvas
 * guarantees visibility instead — a member's box is at least one stroke thick
 * — exactly as §21.3 decided for a line. The server twin is
 * `DOCUMENT_SHAPE_GROUP_MEMBER_MIN_EDGE_EMU`
 * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §22).
 */
export const DOCUMENT_SHAPE_GROUP_MEMBER_MIN_EDGE_EMU = 0;

/**
 * A group's CHILD coordinate space (`a:chExt`, VML `coordsize`) is a unit
 * grid, not a size on the page: it says how many units the group's box is
 * divided into. Word keeps the VML grid on every group it converted from VML
 * (a cover group of `chExt 11015 x 6336` in a 6.99 M EMU box), so holding it
 * to a shape's sixteen-pixel edge clamped the grid and drew every member at a
 * fourteenth of its size. It only has to exist. The server twin is
 * `office_shape_group.OFFICE_SHAPE_GROUP_COORDINATE_MIN_EXTENT`.
 */
export const DOCUMENT_SHAPE_GROUP_COORDINATE_MIN_EMU = 1;

/**
 * A group of ONE is a real drawing: Word's drawing canvas holding one line is
 * a box with one member in it, and `CT_WordprocessingGroup` bounds its members
 * at none. Grouping a SELECTION still needs two (`groupDocumentShapes`). The
 * server twin is `office_smartart_document_members.DOCUMENT_SHAPE_MIN_GROUP_MEMBERS`.
 */
export const DOCUMENT_SHAPE_MIN_GROUP_MEMBERS = 1;

/**
 * A custom path, validated once per path object: a group member's path is
 * read by every paint of its group (corpus R240-D26: 376 freeform members,
 * the path validator the page's main cost). The entry keeps the commands
 * array it was made from, so a path given new commands is read afresh.
 */
const normalizedCustomPathCache = new WeakMap<
  object,
  { readonly commands: unknown; readonly result: ReturnType<typeof normalizePresentationCustomPath> }
>();

function cachedCustomPath(value: unknown): ReturnType<typeof normalizePresentationCustomPath> {
  if (!value || typeof value !== "object") return normalizePresentationCustomPath(value, DOCUMENT_CUSTOM_PATH_MAX_COMMANDS);
  const commands = (value as { commands?: unknown }).commands;
  const cached = normalizedCustomPathCache.get(value);
  if (cached && cached.commands === commands) return cached.result;
  const result = normalizePresentationCustomPath(value, DOCUMENT_CUSTOM_PATH_MAX_COMMANDS);
  normalizedCustomPathCache.set(value, { commands, result });
  return result;
}

/**
 * One group member as a leaf shape (`normalizeDocumentShapeAttrs` of the
 * member with no source), once per member object: the members of a node's
 * normalized group are the same objects paint after paint.
 */
const memberLeafAttrsCache = new WeakMap<object, DocumentShapeAttrs>();

export function documentShapeMemberLeafAttrs(member: DocumentShapeGroupMember): DocumentShapeAttrs {
  const cached = memberLeafAttrsCache.get(member);
  if (cached) return cached;
  const result = normalizeDocumentShapeAttrs({ ...member, source: null });
  memberLeafAttrsCache.set(member, result);
  return result;
}

function normalizeDocumentShapeLeafAttrs(
  attrs: Readonly<Record<string, unknown>>,
  minEdgeEmu?: number,
): DocumentShapeLeafAttrs {
  const preset = isDocumentShapePreset(attrs.preset) ? attrs.preset : "rect";
  const customPath = cachedCustomPath(attrs.customPath);
  const rawRotation = finiteNumber(attrs.rotation, 0);
  const rotation = Math.round((((rawRotation % 360) + 360) % 360) * 10) / 10;
  const outlineStyle: StrokeStyle =
    attrs.outlineStyle === "dashed" || attrs.outlineStyle === "dotted"
      ? attrs.outlineStyle
      : "solid";
  const source =
    attrs.source &&
    typeof attrs.source === "object" &&
    !Array.isArray(attrs.source)
      ? (attrs.source as Record<string, unknown>)
      : null;
  const textPaint = normalizeOfficeTextPaint(attrs.textPaint);
  const isTextBox = attrs.isTextBox === true;
  const linkedTextBoxStoryId = isTextBox
    ? documentShapeLinkedTextBoxCoordinate(attrs.linkedTextBoxStoryId, 1)
    : null;
  const linkedTextBoxSequence =
    linkedTextBoxStoryId === null
      ? null
      : documentShapeLinkedTextBoxCoordinate(attrs.linkedTextBoxSequence, 0);
  return {
    schemaVersion: DOCUMENT_SHAPE_SCHEMA_VERSION,
    preset,
    // An outline that does not normalize is DROPPED rather than failing the
    // whole shape: `preset` is always a real auto-shape, so the shape keeps
    // drawing at its exact rect instead of disappearing.
    ...(customPath ? { customPath } : {}),
    ...(customPath && normalizeOfficeCustomPathPaints(attrs.customPathPaints, customPath)
      ? { customPathPaints: normalizeOfficeCustomPathPaints(attrs.customPathPaints, customPath)! }
      : {}),
    xEmu: boundedInteger(
      attrs.xEmu,
      0,
      -DOCUMENT_SHAPE_MAX_EDGE_EMU,
      DOCUMENT_SHAPE_MAX_EDGE_EMU,
    ),
    yEmu: boundedInteger(
      attrs.yEmu,
      0,
      -DOCUMENT_SHAPE_MAX_EDGE_EMU,
      DOCUMENT_SHAPE_MAX_EDGE_EMU,
    ),
    widthEmu: boundedInteger(
      attrs.widthEmu,
      DOCUMENT_SHAPE_DEFAULT_WIDTH_EMU,
      minEdgeEmu ?? documentShapeMinEdgeEmu(preset),
      DOCUMENT_SHAPE_MAX_EDGE_EMU,
    ),
    heightEmu: boundedInteger(
      attrs.heightEmu,
      DOCUMENT_SHAPE_DEFAULT_HEIGHT_EMU,
      minEdgeEmu ?? documentShapeMinEdgeEmu(preset),
      DOCUMENT_SHAPE_MAX_EDGE_EMU,
    ),
    rotation,
    flipH: attrs.flipH === true,
    flipV: attrs.flipV === true,
    ...(normalizeDocumentShapeAdjustments(attrs.adjustments)
      ? { adjustments: normalizeDocumentShapeAdjustments(attrs.adjustments)! }
      : {}),
    ...normalizeDocumentShapeLineEnds(attrs),
    ...(documentShapeRelativeHeight(attrs.relativeHeight) !== null
      ? { relativeHeight: documentShapeRelativeHeight(attrs.relativeHeight)! }
      : {}),
    ...(typeof attrs.zOrderStackComplete === "boolean"
      ? { zOrderStackComplete: attrs.zOrderStackComplete }
      : {}),
    ...(typeof attrs.lockAspectRatio === "boolean"
      ? { lockAspectRatio: attrs.lockAspectRatio }
      : {}),
    isTextBox,
    text: normalizeDocumentShapeText(attrs.text),
    ...(normalizeDocumentShapeTextBody(attrs.textBody)
      ? { textBody: normalizeDocumentShapeTextBody(attrs.textBody)! }
      : {}),
    ...(() => {
      const textRuns = normalizeDocumentShapeTextRuns(
        attrs.textRuns,
        normalizeDocumentShapeText(attrs.text),
      );
      return textRuns ? { textRuns } : {};
    })(),
    fill: normalizeDocumentShapeColor(attrs.fill, DOCUMENT_SHAPE_DEFAULT_FILL),
    ...(attrs.fillGradient !== undefined ? { fillGradient: normalizeOfficeShapeGradient(attrs.fillGradient) ?? null } : {}),
    ...(attrs.shapeShadow !== undefined ? { shapeShadow: normalizeOfficeShapeShadow(attrs.shapeShadow) ?? null } : {}),
    ...(attrs.shapeEffects !== undefined ? { shapeEffects: normalizeDocumentShapeEffects(attrs.shapeEffects) } : {}),
    ...(attrs.outlineSketch !== undefined ? { outlineSketch: normalizeOfficeLineSketch(attrs.outlineSketch) ?? null } : {}),
    ...(attrs.fillPattern !== undefined ? { fillPattern: officeFillPatternOf(attrs.fillPattern) } : {}),
    ...(attrs.fillPicture !== undefined ? { fillPicture: normalizeDocumentShapePictureFill(attrs.fillPicture) } : {}),
    ...(attrs.fillTransparency !== undefined ? { fillTransparency: documentShapeWholePercent(attrs.fillTransparency) } : {}),
    ...(attrs.outlineTransparency !== undefined ? { outlineTransparency: documentShapeWholePercent(attrs.outlineTransparency) } : {}),
    ...(attrs.outlineGradient !== undefined ? { outlineGradient: normalizeOfficeShapeGradient(attrs.outlineGradient) ?? null } : {}),
    outline: normalizeDocumentShapeColor(
      attrs.outline,
      DOCUMENT_SHAPE_DEFAULT_OUTLINE,
    ),
    outlineWidthPt: normalizeDocumentShapeOutlineWidthPt(
      attrs.outlineWidthPt,
      DOCUMENT_SHAPE_DEFAULT_OUTLINE_WIDTH_PT,
    ),
    outlineStyle,
    outlineAlign: officeStrokeAlign(attrs.outlineAlign),
    // Only a line whose source states such a preset carries the key, and
    // only while its style is still of the preset's family.
    ...(officeStrokeDash(outlineStyle, attrs.outlineDash)
      ? { outlineDash: officeStrokeDash(outlineStyle, attrs.outlineDash) }
      : {}),
    ...officeStrokeDetail(attrs),
    // Only a shape ink drew carries the key: an ordinary shape normalizes to
    // exactly the object it did before ink existed.
    ...(normalizeDocumentInkTool(attrs.inkTool) && normalizeDocumentInkEffect(attrs.inkEffect)
      ? { inkEffect: normalizeDocumentInkEffect(attrs.inkEffect) }
      : {}),
    ...(normalizeDocumentInkTool(attrs.inkTool)
      ? { inkTool: normalizeDocumentInkTool(attrs.inkTool) }
      : {}),
    ...(normalizeDocumentInkTool(attrs.inkTool) && normalizeInkCutEnds(attrs.inkCutEnds)
      ? { inkCutEnds: normalizeInkCutEnds(attrs.inkCutEnds) }
      : {}),
    ...(normalizeDocumentInkTool(attrs.inkTool) && normalizeInkPressure(attrs.inkPressure)
      ? { inkPressure: normalizeInkPressure(attrs.inkPressure) }
      : {}),
    ...(normalizeDocumentInkTool(attrs.inkTool) && normalizeInkAudio(attrs.inkAudio)
      ? { inkAudio: normalizeInkAudio(attrs.inkAudio) }
      : {}),
    ...(normalizeDocumentInkTool(attrs.inkTool) && normalizeInkTape(attrs.inkTape)
      ? { inkTape: normalizeInkTape(attrs.inkTape) }
      : {}),
    effect: normalizeDocumentShapeEffect(attrs.effect),
    ...(attrs.textPaint !== undefined ? { textPaint: textPaint ?? null } : {}),
    textFill: textPaint?.fill ?? normalizeDocumentShapeColor(
      attrs.textFill,
      DOCUMENT_SHAPE_DEFAULT_TEXT_FILL,
    ),
    textOutline: textPaint?.outline ?? normalizeDocumentShapeColor(attrs.textOutline, "none"),
    textOutlineWidthPt: textPaint?.outlineWidthPt ??
      Math.round(
        Math.min(
          12,
          Math.max(0, finiteNumber(attrs.textOutlineWidthPt, 0)),
        ) * 10,
      ) / 10,
    textEffect: textPaint ? "none" : normalizeDocumentShapeEffect(attrs.textEffect),
    textFontSizePt:
      Math.round(
        Math.min(
          // Word's ceiling (1638 pt): 120 typed for a shape's words became 96.
          DOCUMENT_SHAPE_MAX_TEXT_FONT_SIZE_PT,
          Math.max(
            // Word's own floor (1 pt): the size box takes 6 for body text, and
            
            1,
            finiteNumber(
              attrs.textFontSizePt,
              DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_SIZE_PT,
            ),
          ),
        ) * 10,
      ) / 10,
    textFontFamily: normalizeDocumentShapeFontFamily(attrs.textFontFamily),
    textBold: attrs.textBold === true,
    textItalic: attrs.textItalic === true,
    textUnderline: normalizeDocumentShapeTextUnderline(attrs.textUnderline),
    textStrike: attrs.textStrike === true,
    textDirection: normalizeDocumentShapeTextDirection(attrs.textDirection),
    textVerticalAlign: normalizeDocumentShapeTextVerticalAlign(
      attrs.textVerticalAlign,
    ),
    textHorizontalAlign: normalizeDocumentShapeTextHorizontalAlign(
      attrs.textHorizontalAlign,
    ),
    // The paragraph's alignment belongs to an in-line shape only.
    ...(normalizeDocumentShapeWrapMode(attrs.wrapMode) === "inline" &&
    normalizeDocumentShapeParagraphAlign(attrs.textAlign)
      ? { textAlign: normalizeDocumentShapeParagraphAlign(attrs.textAlign)! }
      : {}),
    textWrap: attrs.textWrap !== false,
    // A decorative drawing exports no description (the writer leaves
    // `descr` out); the model keeps the text so clearing the flag brings
    
    altText: normalizeDocumentShapeAltText(attrs.altText),
    ...(attrs.decorative === true ? { decorative: true } : {}),
    ...(normalizeDocumentShapeName(attrs.name) ? { name: normalizeDocumentShapeName(attrs.name)! } : {}),
    ...(attrs.hidden === true ? { hidden: true } : {}),
    ...(attrs.textAutofit === "shape" || attrs.textAutofit === "none"
      ? { textAutofit: attrs.textAutofit }
      : {}),
    ...(normalizeOfficeTextWarp(attrs.textWarp)
      ? { textWarp: normalizeOfficeTextWarp(attrs.textWarp)! }
      : {}),
    ...(typeof attrs.keepTextFlat === "boolean" ? { keepTextFlat: attrs.keepTextFlat } : {}),
    wrapMode: normalizeDocumentShapeWrapMode(attrs.wrapMode),
    ...(linkedTextBoxStoryId !== null && linkedTextBoxSequence !== null
      ? { linkedTextBoxStoryId, linkedTextBoxSequence }
      : {}),
    source,
  };
}

/** The ink a grouped stroke keeps (see {@link DocumentShapeGroupMember.inkTool}). */
function groupMemberInkAttrs(source: {
  inkTool?: unknown;
  inkEffect?: unknown;
  inkCutEnds?: unknown;
  inkPressure?: unknown;
  inkAudio?: unknown;
  inkTape?: unknown;
}): Pick<DocumentShapeGroupMember, "inkTool" | "inkEffect" | "inkCutEnds" | "inkPressure" | "inkAudio" | "inkTape"> {
  const inkTool = normalizeDocumentInkTool(source.inkTool);
  if (!inkTool) return {};
  const inkEffect = normalizeDocumentInkEffect(source.inkEffect);
  const inkCutEnds = normalizeInkCutEnds(source.inkCutEnds);
  const inkPressure = normalizeInkPressure(source.inkPressure);
  const inkAudio = normalizeInkAudio(source.inkAudio);
  const inkTape = normalizeInkTape(source.inkTape);
  return {
    inkTool,
    ...(inkEffect ? { inkEffect } : {}),
    ...(inkCutEnds ? { inkCutEnds } : {}),
    ...(inkPressure ? { inkPressure } : {}),
    ...(inkAudio ? { inkAudio } : {}),
    ...(inkTape ? { inkTape } : {}),
  };
}

/**
 * A group's members, normalized once per members array: a node's attributes
 * are immutable, and the same array is read by every paint, layout pass and
 * command. Normalizing all of them again each time was the page's main cost
 * for a large group (corpus R240-D26, 376 freeform members: 190 s of a 570 s
 * open in `normalizePresentationCustomPath` and the leaf normalizer). The
 * entry also keeps the members it was made from, so an array whose members
 * were swapped is read afresh.
 */
const normalizedGroupMembersCache = new WeakMap<
  object,
  {
    readonly members: readonly unknown[];
    readonly result: readonly DocumentShapeGroupMember[] | null;
  }
>();

function normalizeDocumentShapeGroupMembers(
  value: unknown,
): readonly DocumentShapeGroupMember[] | null {
  if (!Array.isArray(value)) return null;
  const cached = normalizedGroupMembersCache.get(value);
  if (
    cached &&
    cached.members.length === value.length &&
    cached.members.every((member, index) => member === value[index])
  ) {
    return cached.result;
  }
  const result = normalizeDocumentShapeGroupMembersUncached(value);
  normalizedGroupMembersCache.set(value, { members: [...value], result });
  return result;
}

function normalizeDocumentShapeGroupMembersUncached(
  value: unknown,
): readonly DocumentShapeGroupMember[] | null {
  if (
    !Array.isArray(value) ||
    value.length < DOCUMENT_SHAPE_MIN_GROUP_MEMBERS ||
    value.length > DOCUMENT_SHAPE_MAX_GROUP_MEMBERS
  ) {
    return null;
  }
  const members: DocumentShapeGroupMember[] = [];
  const ids = new Set<string>();
  const drawingIds = new Set<number>();
  for (const raw of value) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const input = raw as Record<string, unknown>;
    const memberId =
      typeof input.memberId === "string" &&
      /^[A-Za-z0-9._-]{1,64}$/.test(input.memberId)
        ? input.memberId
        : null;
    if (
      !memberId ||
      ids.has(memberId) ||
      !isDocumentShapePreset(input.preset)
    ) {
      return null;
    }
    const attrs = normalizeDocumentShapeLeafAttrs(
      { ...input, source: null },
      DOCUMENT_SHAPE_GROUP_MEMBER_MIN_EDGE_EMU,
    );
    const kind =
      input.kind === undefined || input.kind === "shape"
        ? "shape"
        : input.kind === "connector"
          ? "connector"
          : input.kind === "picture"
            ? "picture"
            : input.kind === "chart"
              ? "chart"
              : null;
    const drawingId =
      typeof input.drawingId === "number" &&
      Number.isSafeInteger(input.drawingId) &&
      input.drawingId >= 1 &&
      input.drawingId <= DOCUMENT_SHAPE_MAX_DRAWING_ID
        ? input.drawingId
        : null;
    if (
      kind === null ||
      !Number.isSafeInteger(input.xEmu) ||
      !Number.isSafeInteger(input.yEmu) ||
      !Number.isSafeInteger(input.widthEmu) ||
      !Number.isSafeInteger(input.heightEmu) ||
      // A member may stand outside the group's child space (corpus
      // R98-D11: a folded corner left of the child origin, turned back
      // inside): Word draws it where it is, so its place is bounded as a
      // shape's own, not by the space.
      Math.abs(attrs.xEmu) > DOCUMENT_SHAPE_MAX_EDGE_EMU ||
      Math.abs(attrs.yEmu) > DOCUMENT_SHAPE_MAX_EDGE_EMU
    ) {
      return null;
    }
    if (
      input.drawingId !== undefined &&
      (drawingId === null || drawingIds.has(drawingId))
    ) {
      return null;
    }
    if (drawingId !== null) drawingIds.add(drawingId);
    const connectorPreset =
      input.connectorPreset === "straight" || input.connectorPreset === "elbow"
        ? input.connectorPreset
        : null;
    const startMemberId =
      typeof input.startMemberId === "string" &&
      /^[A-Za-z0-9._-]{1,64}$/.test(input.startMemberId)
        ? input.startMemberId
        : null;
    const endMemberId =
      typeof input.endMemberId === "string" &&
      /^[A-Za-z0-9._-]{1,64}$/.test(input.endMemberId)
        ? input.endMemberId
        : null;
    const startConnectionSite =
      typeof input.startConnectionSite === "number" &&
      Number.isInteger(input.startConnectionSite) &&
      input.startConnectionSite >= 0 &&
      input.startConnectionSite <= 3
        ? input.startConnectionSite
        : null;
    const endConnectionSite =
      typeof input.endConnectionSite === "number" &&
      Number.isInteger(input.endConnectionSite) &&
      input.endConnectionSite >= 0 &&
      input.endConnectionSite <= 3
        ? input.endConnectionSite
        : null;
    if (
      kind === "connector" &&
      (drawingId === null ||
        connectorPreset === null ||
        startMemberId === null ||
        endMemberId === null ||
        startMemberId === endMemberId ||
        startConnectionSite === null ||
        endConnectionSite === null ||
        attrs.preset !== "line" ||
        attrs.fill !== "none" ||
        attrs.outline === "none" ||
        attrs.rotation !== 0 ||
        attrs.lockAspectRatio === true)
    ) {
      return null;
    }
    if (
      kind !== "connector" &&
      (input.connectorPreset !== undefined ||
        input.startMemberId !== undefined ||
        input.startConnectionSite !== undefined ||
        input.endMemberId !== undefined ||
        input.endConnectionSite !== undefined)
    ) {
      return null;
    }
    // A picture member's own fields, and only a picture member's. A picture
    // draws its bytes, so a label or paint on one would be two claims about
    // the same pixels -- refused, not silently ignored by whichever host
    // draws last (the server model and the page validator say the same).
    const picture = kind === "picture" ? documentShapeMemberPicture(input) : null;
    if (kind === "picture") {
      if (
        !picture ||
        attrs.text !== "" ||
        attrs.fill !== "none" ||
        // A picture's own line (Picture Border, `a:ln`) is drawn over its
        // pixels: a colour needs a width, and no line has none.
        (attrs.outline === "none") !== (attrs.outlineWidthPt === 0) ||
        input.customPath !== undefined
      ) {
        return null;
      }
    } else if (
      DOCUMENT_SHAPE_MEMBER_PICTURE_CROP_FIELDS.some(
        (key) => input[key] !== undefined,
      )
    ) {
      return null;
    }
    // A chart member draws its chart: no label, no paint (as the server says).
    const chart =
      kind === "chart" && input.chart && typeof input.chart === "object" && !Array.isArray(input.chart)
        ? (input.chart as Readonly<Record<string, unknown>>)
        : null;
    if (kind === "chart" && (!chart || attrs.text !== "" || attrs.fill !== "none" || attrs.outline !== "none")) {
      return null;
    }
    if (kind !== "chart" && input.chart !== undefined) return null;
    // A connector's geometry IS its route between two endpoints, recomputed
    // below and compared exactly. An outline on one would be a second,
    // unverifiable claim about where it goes, so it is refused the same way a
    // shape carrying connector fields is.
    if (
      kind === "connector" &&
      (input.customPath !== undefined || input.adjustments !== undefined)
    ) {
      return null;
    }
    const style = documentShapeMemberStyle(input, attrs);
    if (!style) return null;
    const groupFill = documentShapeMemberGroupFill(input, kind);
    if (!groupFill) return null;
    ids.add(memberId);
    members.push({
      memberId,
      ...(drawingId !== null ? { drawingId } : {}),
      ...(kind === "connector"
        ? {
            kind,
            connectorPreset: connectorPreset!,
            startMemberId: startMemberId!,
            startConnectionSite: startConnectionSite!,
            endMemberId: endMemberId!,
            endConnectionSite: endConnectionSite!,
          }
        : kind === "picture"
          ? { kind }
          : kind === "chart"
            ? { kind, chart: chart! }
            : {}),
      preset: attrs.preset,
      ...(attrs.customPath ? { customPath: attrs.customPath } : {}),
      ...(attrs.customPathPaints ? { customPathPaints: attrs.customPathPaints } : {}),
      ...(attrs.adjustments ? { adjustments: attrs.adjustments } : {}),
      // A line's arrowheads, read by the leaf normalizer as a lone shape's
      // (`normalizeDocumentShapeLineEnds`): dropped here, a Word group's
      // arrows were gone from the page, and a save that wrote the group anew
      // wrote it without them.
      ...(normalizeDocumentShapeLineEnds(attrs) as Pick<
        DocumentShapeGroupMember,
        "headEnd" | "tailEnd" | "headEndWidth" | "headEndLength" | "tailEndWidth" | "tailEndLength"
      >),
      ...(kind === "shape" ? groupMemberInkAttrs(attrs) : {}),
      xEmu: attrs.xEmu,
      yEmu: attrs.yEmu,
      widthEmu: attrs.widthEmu,
      heightEmu: attrs.heightEmu,
      rotation: attrs.rotation,
      flipH: attrs.flipH,
      flipV: attrs.flipV,
      ...(typeof attrs.relativeHeight === "number"
        ? { relativeHeight: attrs.relativeHeight }
        : {}),
      ...(typeof attrs.lockAspectRatio === "boolean"
        ? { lockAspectRatio: attrs.lockAspectRatio }
        : {}),
      ...(input.lockEditPoints === true ? { lockEditPoints: true as const } : {}),
      ...(typeof input.groupPath === "string" && /^(?:\d{1,4}(?:\/\d{1,4}){0,15})?$/.test(input.groupPath)
        ? { groupPath: input.groupPath }
        : {}),
      ...(documentShapeMemberTextStory(input.textStory)
        ? { textStory: documentShapeMemberTextStory(input.textStory)! }
        : {}),
      isTextBox: attrs.isTextBox,
      text: attrs.text,
      ...(attrs.textBody ? { textBody: attrs.textBody } : {}),
      ...(attrs.textRuns ? { textRuns: attrs.textRuns } : {}),
      ...(attrs.textWarp ? { textWarp: attrs.textWarp } : {}),
      ...(input.textEffects && typeof input.textEffects === "object" && !Array.isArray(input.textEffects)
        ? { textEffects: input.textEffects as Readonly<Record<string, unknown>> }
        : {}),
      ...(() => {
        try {
          const fillPattern = normalizeOfficeChartFillPattern(input.fillPattern);
          return fillPattern ? { fillPattern } : {};
        } catch {
          return {};
        }
      })(),
      fill: attrs.fill,
      ...style,
      ...groupFill,
      ...(kind === "shape" && attrs.shapeEffects ? { shapeEffects: attrs.shapeEffects } : {}),
      outline: attrs.outline,
      outlineWidthPt: attrs.outlineWidthPt,
      outlineStyle: attrs.outlineStyle,
      ...(attrs.outlineDash ? { outlineDash: attrs.outlineDash } : {}),
      // Only a member whose source line STATED an alignment carries the key,
      // so an ordinary group normalizes to exactly the object it did before.
      ...(attrs.outlineAlign ? { outlineAlign: attrs.outlineAlign } : {}),
      ...officeStrokeDetail(attrs),
      ...(kind === "shape" && attrs.outlineSketch ? { outlineSketch: attrs.outlineSketch } : {}),
      effect: attrs.effect,
      // The member's WordArt, read by the leaf normalizer as a shape's; a
      // connector or a picture draws no words to paint.
      ...(attrs.textPaint && kind === "shape" ? { textPaint: attrs.textPaint } : {}),
      textFill: attrs.textFill,
      textOutline: attrs.textOutline,
      textOutlineWidthPt: attrs.textOutlineWidthPt,
      textEffect: attrs.textEffect,
      textFontSizePt: attrs.textFontSizePt,
      textFontFamily: attrs.textFontFamily,
      textBold: attrs.textBold,
      ...(attrs.textItalic ? { textItalic: true } : {}),
      ...(attrs.textUnderline !== "none" ? { textUnderline: attrs.textUnderline } : {}),
      ...(attrs.textStrike ? { textStrike: true } : {}),
      textDirection: attrs.textDirection,
      textVerticalAlign: attrs.textVerticalAlign,
      ...(attrs.textHorizontalAlign !== "center"
        ? { textHorizontalAlign: attrs.textHorizontalAlign }
        : {}),
      textWrap: attrs.textWrap,
      ...(attrs.textAutofit === "shape" && kind === "shape" ? { textAutofit: "shape" as const } : {}),
      altText: attrs.altText,
      ...(typeof input.linkHref === "string" &&
      input.linkHref.length > 0 &&
      input.linkHref.length <= 2_000
        ? {
            linkHref: input.linkHref,
            ...(typeof input.linkTitle === "string" &&
            input.linkTitle.length > 0 &&
            input.linkTitle.length <= 260
              ? { linkTitle: input.linkTitle }
              : {}),
          }
        : {}),
      wrapMode: attrs.wrapMode,
      ...(typeof attrs.linkedTextBoxStoryId === "number" &&
      typeof attrs.linkedTextBoxSequence === "number"
        ? {
            linkedTextBoxStoryId: attrs.linkedTextBoxStoryId,
            linkedTextBoxSequence: attrs.linkedTextBoxSequence,
          }
        : {}),
      ...(documentShapeMemberSmartArtIdentity(input) ?? {}),
      ...(picture ?? {}),
    });
  }
  const connectorMembers = members.filter(
    (member) => member.kind === "connector",
  );
  if (connectorMembers.length > 0) {
    if (members.some((member) => typeof member.drawingId !== "number")) {
      return null;
    }
    const byId = new Map(members.map((member) => [member.memberId, member]));
    for (const connector of connectorMembers) {
      const start = byId.get(connector.startMemberId!);
      const end = byId.get(connector.endMemberId!);
      if (
        !start ||
        !end ||
        start.kind === "connector" ||
        end.kind === "connector"
      ) {
        return null;
      }
      const route = documentShapeConnectorRoute(
        start,
        connector.startConnectionSite!,
        end,
        connector.endConnectionSite!,
      );
      if (
        !route ||
        connector.xEmu !== route.xEmu ||
        connector.yEmu !== route.yEmu ||
        connector.widthEmu !== route.widthEmu ||
        connector.heightEmu !== route.heightEmu ||
        connector.flipH !== route.flipH ||
        connector.flipV !== route.flipV
      ) {
        return null;
      }
    }
  }
  return members;
}

/**
 * A shape's `shapeEffects`: Office's seven kinds (Preset, Shadow outer or
 * inner, Reflection, Glow, Soft Edges, Bevel, 3-D Rotation), each but the
 * Preset possibly "none". The server twin is
 * `document_docx_shape_effects.normalize_shape_effects`.
 */
export function normalizeDocumentShapeEffects(value: unknown): OfficePictureEffects | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  if (Object.keys(value).some((key) => !(OFFICE_PICTURE_EFFECT_KINDS as readonly string[]).includes(key))) {
    return null;
  }
  const effects = normalizeOfficePictureEffects(value);
  if (!effects || (effects as { preset?: unknown }).preset === "none") return null;
  return effects;
}

/**
 * A shape's effects as Office states them, for the full Shape Effects menu:
 * its `shapeEffects`, the outer shadow an older page keeps in `shapeShadow`
 * (when `shapeEffects` states no shadow), and the older one-of-seven `effect`
 * read as the effect it drew (`accent1`: the document theme's Accent 1).
 */
export function documentShapeOfficeEffects(
  attrs: Pick<DocumentShapeAttrs, "shapeEffects" | "shapeShadow" | "effect">,
  accent1?: string,
): OfficePictureEffects | undefined {
  const own = normalizeDocumentShapeEffects(attrs.shapeEffects) ?? {};
  const legacy: Partial<Record<OfficePictureEffectKind, unknown>> = {
    ...officeEffectsOfLegacyPreset(attrs.effect, accent1),
  };
  if (own.shadow === undefined) {
    if (attrs.shapeShadow) legacy.shadow = officeShadowFromShapeShadow(attrs.shapeShadow);
    else if (attrs.effect === "shadow") legacy.shadow = OFFICE_SWITCH_SHADOW;
  }
  return normalizeOfficePictureEffects({ ...legacy, ...own }) ?? undefined;
}

/**
 * An anonymous group's members without the diagram's own fields, once per
 * normalized members array: made afresh on every read, the members were new
 * objects each paint, and nothing kept per member could be kept (corpus
 * R240-D26: 6 s of the open in this map alone).
 */
const anonymousGroupMembersCache = new WeakMap<
  readonly DocumentShapeGroupMember[],
  readonly DocumentShapeGroupMember[]
>();

function anonymousGroupMembers(
  groupMembers: readonly DocumentShapeGroupMember[],
): readonly DocumentShapeGroupMember[] {
  const cached = anonymousGroupMembersCache.get(groupMembers);
  if (cached) return cached;
  const result = groupMembers.map(({
          smartArtRole: _role,
          smartArtItemId: _itemId,
          smartArtItemIndex: _itemIndex,
          smartArtLevel: _level,
          smartArtPicturePlaceholder: _picture,
          src,
          smartArtAuthorPaint: _authorPaint,
          smartArtAuthorTextSize: _authorTextSize,
          smartArtAuthorTextFont: _authorTextFont,
          smartArtAuthorShape: _authorShape,
          smartArtPictureCrop: _pictureCrop,
          smartArtPictureStyle: _pictureStyle,
          smartArtCustom: _custom,
          smartArtTextBox: _textBox,
          smartArtAssistant: _assistant,
          ...member
        }) =>
          // Everything above is a DIAGRAM's, and an anonymous shape group
          // cannot have authored any of it -- including a picture SLOT's
          // contents, which is what `src` used to mean and nothing else.
          // A member's own bevel (`shape3d`) and the camera of the group
          // that turns it (`scene3d`, Format Shape > 3-D Rotation on a Word
          // group; `scene3dFrame` for a nested one's) are the shapes' own,
          // read by the API for any group and drawn as Word draws them; so is
          // a text box's own text 3-D (`textShape3d`, `wps:bodyPr/a:sp3d`,
          // corpus R439-D15): dropped here, a save that wrote the group anew
          // wrote its words flat.
          //
          // A `kind: "picture"` member's `src` is not a slot's contents: it is
          // the member's own `pic:blipFill` bytes, and an ordinary Word group
          // full of pictures is the commonest place to find one. Dropping it
          // here left every picture member on a non-SmartArt group drawing an
          // empty frame -- read by both ends, projected by the server, and
          // then quietly emptied by the page normalizer
          // (SHAPE-IMPORT-AUDIT.md §58).
          member.kind === "picture" && typeof src === "string"
            ? { ...member, src }
            : member);
  anonymousGroupMembersCache.set(groupMembers, result);
  return result;
}

export function normalizeDocumentShapeAttrs(
  value: DocumentShapeAttrsInput,
): DocumentShapeAttrs {
  const attrs = (value ?? {}) as Readonly<Record<string, unknown>>;
  const normalized = normalizeDocumentShapeLeafAttrs(attrs);
  if (!Object.hasOwn(attrs, "groupMembers")) return normalized;
  const coordinateWidthEmu = boundedInteger(
    attrs.groupCoordinateWidthEmu,
    normalized.widthEmu,
    DOCUMENT_SHAPE_GROUP_COORDINATE_MIN_EMU,
    DOCUMENT_SHAPE_MAX_EDGE_EMU,
  );
  const coordinateHeightEmu = boundedInteger(
    attrs.groupCoordinateHeightEmu,
    normalized.heightEmu,
    DOCUMENT_SHAPE_GROUP_COORDINATE_MIN_EMU,
    DOCUMENT_SHAPE_MAX_EDGE_EMU,
  );
  const groupMembers = normalizeDocumentShapeGroupMembers(attrs.groupMembers);
  if (!groupMembers) return normalized;
  // The group's own fill, worn by its `fillFromGroup` members; one that does
  // not normalize is dropped (its members then draw their fallback colour).
  const groupFill = normalizeDocumentShapeGroupFill(attrs.groupFill) ?? undefined;
  // A SmartArt identity is only an identity when the WHOLE group carries it:
  // one kind, one group id, and an identity on every member. A half-stamped
  // group is an ordinary shape group, never a diagram the SmartArt tools own.
  const smartArtKind = smartArtKindToken(attrs.smartArtKind);
  const smartArtGroupId = smartArtIdentityToken(attrs.smartArtGroupId);
  const identified =
    smartArtKind !== null &&
    smartArtGroupId !== null &&
    groupMembers.every((member) => member.smartArtRole !== undefined);
  // A quick style belongs to a DIAGRAM, so it goes exactly where the identity
  // goes: an anonymous shape group keeps its members' ordinary fill and line,
  // their own bevels and its camera, and loses a style it could not have
  // authored.
  const smartArtStyle = isSmartArtStyleId(attrs.smartArtStyle)
    ? attrs.smartArtStyle
    : null;
  // …and so does its direction.
  const smartArtRightToLeft =
    (attrs as { smartArtDirection?: unknown }).smartArtDirection === "rtl";
  // …and its colour variation.
  const smartArtColorsRaw = (attrs as { smartArtColors?: unknown }).smartArtColors;
  const smartArtColors =
    typeof smartArtColorsRaw === "string" &&
    smartArtColorVariant(smartArtColorsRaw, { all: true }) !== null
      ? smartArtColorsRaw
      : null;
  // …and so does the alias it is drawn as, when that is one of its kind's;
  // a diagram laid out from its file's own definition keeps that definition's
  // urn and the digest of its part (stamped by the engine host on the way
  // out; the server holds the page to having both).
  const statedUrn = (attrs as { smartArtLayoutUrn?: unknown }).smartArtLayoutUrn;
  const statedDigest = (attrs as { smartArtLayoutSha256?: unknown }).smartArtLayoutSha256;
  const ownDefinition = smartArtKind === SMARTART_OFFICE_DEFINITION_KIND;
  const smartArtLayoutUrn = ownDefinition
    ? typeof statedUrn === "string" && statedUrn.length >= 1 && statedUrn.length <= 512
      ? statedUrn
      : null
    : smartArtKind !== null
      ? nativeSmartArtAliasOf(smartArtKind, statedUrn)
      : null;
  const smartArtLayoutSha256 =
    ownDefinition && typeof statedDigest === "string" && /^[0-9a-f]{64}$/.test(statedDigest)
      ? statedDigest
      : null;
  return {
    ...normalized,
    groupMembers: identified ? groupMembers : anonymousGroupMembers(groupMembers),
    groupCoordinateWidthEmu: coordinateWidthEmu,
    groupCoordinateHeightEmu: coordinateHeightEmu,
    ...(groupFill ? { groupFill } : {}),
    ...(identified ? { smartArtKind, smartArtGroupId } : {}),
    ...(identified && smartArtStyle !== null ? { smartArtStyle } : {}),
    ...(identified && smartArtRightToLeft
      ? { smartArtDirection: "rtl" as const }
      : {}),
    ...(identified && smartArtColors !== null ? { smartArtColors } : {}),
    ...(identified && smartArtLayoutUrn !== null ? { smartArtLayoutUrn } : {}),
    ...(identified && smartArtLayoutSha256 !== null ? { smartArtLayoutSha256 } : {}),
  };
}

/**
 * The catalog SmartArt this shape group IS, or `null` when it is an ordinary
 * group (or a preserved non-catalog Office diagram, which carries no catalog
 * identity by definition). This is what the Docs SmartArt contextual tab
 * recognises, on an inserted diagram and on an imported one alike.
 */
export function documentShapeSmartArtIdentity(
  value: DocumentShapeAttrsInput,
): { readonly kind: string; readonly groupId: string } | null {
  const attrs = normalizeDocumentShapeAttrs(value);
  return attrs.smartArtKind && attrs.smartArtGroupId
    ? { kind: attrs.smartArtKind, groupId: attrs.smartArtGroupId }
    : null;
}

/**
 * A non-catalog Office diagram kept exactly as Office wrote it: the preserve-
 * only group the docx projection emits. Its geometry and paint belong to the
 * original drawing part; only member TEXT may be written back.
 */
export function documentShapeIsPreservedNativeSmartArt(
  value: DocumentShapeAttrsInput,
): boolean {
  const attrs = normalizeDocumentShapeAttrs(value);
  const source = attrs.source;
  return Boolean(
    attrs.groupMembers?.length &&
      source &&
      source.adapter === "docx" &&
      source.nodeType === "nativeSmartArtGroup" &&
      typeof source.nativeSmartArtView === "object" &&
      source.nativeSmartArtView !== null,
  );
}

/**
 * Whether a member's LABEL may be edited in place. A preserved Office diagram
 * is geometry-locked but text-editable — the backend writes a text change into
 * the diagram's own data part and refuses every other delta — so it needs its
 * own answer, distinct from {@link documentShapeGroupIsEditable}.
 *
 * A preserved diagram whose own projection says its text is NOT editable is
 * excluded: without a text receipt the save has nothing to write the label
 * into, so offering the input would let an author type into nothing.
 * {@link documentShapeNativeSmartArtTextBlocker} says why, in their words.
 */
export function documentShapeGroupTextIsEditable(
  value: DocumentShapeAttrsInput,
): boolean {
  return (
    documentShapeGroupIsEditable(value) ||
    (documentShapeIsPreservedNativeSmartArt(value) &&
      documentShapeNativeSmartArtTextBlocker(value) === null)
  );
}

/**
 * Why the selected preserved diagram's text cannot be edited, or `null` when
 * it can. Reads the diagram's own projection — the `textEditable` its import
 * recorded and the labels Office actually drew — so the hosts and the shared
 * SmartArt tab state one answer.
 */
export function documentShapeNativeSmartArtTextBlocker(
  value: DocumentShapeAttrsInput,
): NativeSmartArtTextBlocker | null {
  if (!documentShapeIsPreservedNativeSmartArt(value)) return null;
  const attrs = normalizeDocumentShapeAttrs(value);
  const view = attrs.source?.nativeSmartArtView as
    | Record<string, unknown>
    | undefined;
  const members = Array.isArray(view?.members) ? view.members : [];
  return nativeSmartArtTextBlocker({
    textEditable: view?.textEditable === true,
    texts: [
      ...members.map((member) =>
        member && typeof member === "object"
          ? (member as Record<string, unknown>).text
          : undefined,
      ),
      ...(attrs.groupMembers ?? []).map((member) => member.text),
    ],
  });
}

export function documentShapeDefaultAttrs(
  preset: DocumentShapePreset,
): DocumentShapeAttrs {
  return normalizeDocumentShapeAttrs({
    preset,
    relativeHeight: DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
    zOrderStackComplete: true,
  });
}

export function documentShapeIsGroup(
  value: DocumentShapeAttrsInput,
): boolean {
  const attrs = normalizeDocumentShapeAttrs(value);
  return Boolean(attrs.groupMembers?.length);
}

/** A group may move as one only when authored here or source-authenticated. */
/**
 * The receipt a CONVERTED Office diagram keeps: the anchor of the diagram it
 * replaced, plus the parts and fingerprints that prove which one.
 *
 * A converted diagram is MRAI's own SmartArt — it is edited, restyled and
 * exported exactly like an inserted one — but the source `.docx` still holds
 * the DiagramML it replaced, and the original-package export has to replace
 * that drawing on EVERY save, not only the first. The page is the only place
 * that association can live (the patchback compares one saved version with the
 * next and rebuilds the package from the immutable original), so the converted
 * group carries it, the same way a converted worksheet diagram's members carry
 * `source.nodeType === "importedSmartArtMember"`.
 *
 * It is a receipt, never an authority: the server re-reads the diagram's parts
 * from the original package and refuses a receipt that does not match them.
 */
export function documentShapeConvertedSmartArtSource(
  value: DocumentShapeAttrsInput,
): {
  readonly sourceAnchor: string;
  readonly parts: Record<string, unknown>;
  readonly fingerprints: Record<string, unknown>;
} | null {
  const source = normalizeDocumentShapeAttrs(value).source;
  if (
    !source ||
    source.adapter !== "docx" ||
    source.nodeType !== "convertedSmartArtGroup" ||
    typeof source.sourceAnchor !== "string" ||
    !source.sourceAnchor
  ) {
    return null;
  }
  const view = source.nativeSmartArtView;
  if (!view || typeof view !== "object" || Array.isArray(view)) return null;
  const { parts, fingerprints } = view as Record<string, unknown>;
  return parts &&
    typeof parts === "object" &&
    !Array.isArray(parts) &&
    fingerprints &&
    typeof fingerprints === "object" &&
    !Array.isArray(fingerprints)
    ? {
        sourceAnchor: source.sourceAnchor,
        parts: parts as Record<string, unknown>,
        fingerprints: fingerprints as Record<string, unknown>,
      }
    : null;
}

export function documentShapeGroupIsEditable(
  value: DocumentShapeAttrsInput,
): boolean {
  const attrs = normalizeDocumentShapeAttrs(value);
  if (!attrs.groupMembers?.length) return false;
  if (attrs.source === null) return true;
  // A converted diagram is MRAI's own group: it keeps only the receipt that
  // says which drawing it replaces, and edits exactly like an inserted one.
  if (documentShapeConvertedSmartArtSource(attrs)) return true;
  return (
    attrs.source.adapter === "docx" &&
    attrs.source.nodeType === "shapeGroup" &&
    typeof attrs.source.sourceAnchor === "string" &&
    /^word\/document\.xml#p\[\d+\]\/group\[\d+\]$/.test(
      attrs.source.sourceAnchor,
    ) &&
    attrs.source.preserveOnly !== true
  );
}

export function documentShapeHasNativeConnectors(
  value: DocumentShapeAttrsInput,
): boolean {
  return Boolean(
    normalizeDocumentShapeAttrs(value).groupMembers?.some(
      (member) => member.kind === "connector",
    ),
  );
}

function documentConnectorCanvasMembers(
  members: readonly DocumentShapeGroupMember[],
): readonly DocumentShapeGroupMember[] | null {
  const byId = new Map(members.map((member) => [member.memberId, member]));
  const next = members.map((member) => {
    if (member.kind !== "connector") return member;
    const start = byId.get(member.startMemberId!);
    const end = byId.get(member.endMemberId!);
    const route =
      start && end
        ? documentShapeConnectorRoute(
            start,
            member.startConnectionSite!,
            end,
            member.endConnectionSite!,
          )
        : null;
    return route ? { ...member, ...route } : null;
  });
  return next.every(
    (member): member is DocumentShapeGroupMember => member !== null,
  )
    ? next
    : null;
}

export interface DocumentConnectorCanvasLabels {
  readonly start: string;
  readonly step: string;
  readonly finish: string;
}

const DEFAULT_DOCUMENT_CONNECTOR_CANVAS_LABELS: DocumentConnectorCanvasLabels =
  Object.freeze({ start: "Start", step: "Step", finish: "Finish" });

/** A bounded authored canvas: three WPS shapes joined by two native connectors. */
export function documentShapeConnectorCanvasAttrs(
  labels: DocumentConnectorCanvasLabels = DEFAULT_DOCUMENT_CONNECTOR_CANVAS_LABELS,
): DocumentShapeAttrs {
  const px = DOCUMENT_SHAPE_EMU_PER_PX;
  const shapes: DocumentShapeGroupMember[] = [
    {
      memberId: "shape-a",
      drawingId: 1,
      preset: "roundRect",
      xEmu: 80 * px,
      yEmu: 80 * px,
      widthEmu: 160 * px,
      heightEmu: 90 * px,
      rotation: 0,
      flipH: false,
      flipV: false,
      lockAspectRatio: false,
      text: labels.start,
      fill: "#dbeafe",
      outline: "#2563eb",
      outlineWidthPt: 1.5,
      outlineStyle: "solid",
    },
    {
      memberId: "shape-b",
      drawingId: 2,
      preset: "rect",
      xEmu: 360 * px,
      yEmu: 180 * px,
      widthEmu: 160 * px,
      heightEmu: 90 * px,
      rotation: 0,
      flipH: false,
      flipV: false,
      lockAspectRatio: false,
      text: labels.step,
      fill: "#fef3c7",
      outline: "#d97706",
      outlineWidthPt: 1.5,
      outlineStyle: "solid",
    },
    {
      memberId: "shape-c",
      drawingId: 3,
      preset: "roundRect",
      xEmu: 640 * px,
      yEmu: 80 * px,
      widthEmu: 160 * px,
      heightEmu: 90 * px,
      rotation: 0,
      flipH: false,
      flipV: false,
      lockAspectRatio: false,
      text: labels.finish,
      fill: "#dcfce7",
      outline: "#16a34a",
      outlineWidthPt: 1.5,
      outlineStyle: "solid",
    },
  ];
  const connectors: DocumentShapeGroupMember[] = [
    {
      memberId: "connector-ab",
      drawingId: 4,
      kind: "connector",
      connectorPreset: "elbow",
      startMemberId: "shape-a",
      startConnectionSite: 1,
      endMemberId: "shape-b",
      endConnectionSite: 3,
      preset: "line",
      xEmu: 0,
      yEmu: 0,
      widthEmu: DOCUMENT_SHAPE_MIN_EDGE_EMU,
      heightEmu: DOCUMENT_SHAPE_MIN_EDGE_EMU,
      rotation: 0,
      flipH: false,
      flipV: false,
      lockAspectRatio: false,
      text: "",
      fill: "none",
      outline: "#475569",
      outlineWidthPt: 1.5,
      outlineStyle: "solid",
    },
    {
      memberId: "connector-bc",
      drawingId: 5,
      kind: "connector",
      connectorPreset: "elbow",
      startMemberId: "shape-b",
      startConnectionSite: 1,
      endMemberId: "shape-c",
      endConnectionSite: 3,
      preset: "line",
      xEmu: 0,
      yEmu: 0,
      widthEmu: DOCUMENT_SHAPE_MIN_EDGE_EMU,
      heightEmu: DOCUMENT_SHAPE_MIN_EDGE_EMU,
      rotation: 0,
      flipH: false,
      flipV: false,
      lockAspectRatio: false,
      text: "",
      fill: "none",
      outline: "#475569",
      outlineWidthPt: 1.5,
      outlineStyle: "solid",
    },
  ];
  const groupMembers = documentConnectorCanvasMembers([
    ...connectors,
    ...shapes,
  ])!;
  return normalizeDocumentShapeAttrs({
    preset: "rect",
    xEmu: 0,
    yEmu: 0,
    widthEmu: 880 * px,
    heightEmu: 360 * px,
    rotation: 0,
    flipH: false,
    flipV: false,
    relativeHeight: DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
    zOrderStackComplete: true,
    lockAspectRatio: false,
    text: "",
    fill: "none",
    outline: "none",
    outlineWidthPt: 0,
    outlineStyle: "solid",
    source: null,
    groupCoordinateWidthEmu: 880 * px,
    groupCoordinateHeightEmu: 360 * px,
    groupMembers,
  });
}

export interface DocumentMindMapCanvasLabels {
  readonly center: string;
  readonly branches: readonly [string, string, string, string];
}

/** A native radial mind map: one center, four branches, four connectors. */
export function documentShapeMindMapCanvasAttrs(
  labels: DocumentMindMapCanvasLabels,
): DocumentShapeAttrs {
  const px = DOCUMENT_SHAPE_EMU_PER_PX;
  const shapePositions = [
    { memberId: "branch-a", x: 60, y: 55 },
    { memberId: "branch-b", x: 650, y: 55 },
    { memberId: "branch-c", x: 60, y: 375 },
    { memberId: "branch-d", x: 650, y: 375 },
  ] as const;
  const center: DocumentShapeGroupMember = {
    memberId: "center",
    drawingId: 1,
    preset: "ellipse",
    xEmu: 360 * px,
    yEmu: 215 * px,
    widthEmu: 160 * px,
    heightEmu: 90 * px,
    rotation: 0,
    flipH: false,
    flipV: false,
    lockAspectRatio: false,
    text: labels.center,
    fill: "#ede9fe",
    outline: "#7c3aed",
    outlineWidthPt: 1.5,
    outlineStyle: "solid",
  };
  const branches: DocumentShapeGroupMember[] = shapePositions.map(
    (position, index) => ({
      memberId: position.memberId,
      drawingId: index + 2,
      preset: "roundRect",
      xEmu: position.x * px,
      yEmu: position.y * px,
      widthEmu: 170 * px,
      heightEmu: 80 * px,
      rotation: 0,
      flipH: false,
      flipV: false,
      lockAspectRatio: false,
      text: labels.branches[index],
      fill: "#f8fafc",
      outline: "#64748b",
      outlineWidthPt: 1.25,
      outlineStyle: "solid",
    }),
  );
  const connectors: DocumentShapeGroupMember[] = branches.map(
    (branch, index) => ({
      memberId: `connector-${index + 1}`,
      drawingId: index + 6,
      kind: "connector",
      connectorPreset: "straight",
      startMemberId: center.memberId,
      startConnectionSite: index < 2 ? 0 : 2,
      endMemberId: branch.memberId,
      endConnectionSite: index % 2 === 0 ? 1 : 3,
      preset: "line",
      xEmu: 0,
      yEmu: 0,
      widthEmu: DOCUMENT_SHAPE_MIN_EDGE_EMU,
      heightEmu: DOCUMENT_SHAPE_MIN_EDGE_EMU,
      rotation: 0,
      flipH: false,
      flipV: false,
      lockAspectRatio: false,
      text: "",
      fill: "none",
      outline: "#7c3aed",
      outlineWidthPt: 1.25,
      outlineStyle: "solid",
    }),
  );
  const groupMembers = documentConnectorCanvasMembers([
    ...connectors,
    center,
    ...branches,
  ])!;
  return normalizeDocumentShapeAttrs({
    preset: "rect",
    xEmu: 0,
    yEmu: 0,
    widthEmu: 880 * px,
    heightEmu: 520 * px,
    rotation: 0,
    flipH: false,
    flipV: false,
    relativeHeight: DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
    zOrderStackComplete: true,
    lockAspectRatio: false,
    text: "",
    fill: "none",
    outline: "none",
    outlineWidthPt: 0,
    outlineStyle: "solid",
    source: null,
    groupCoordinateWidthEmu: 880 * px,
    groupCoordinateHeightEmu: 520 * px,
    groupMembers,
  });
}

/**
 * Move ONE member inside its group, in the group's own child space.
 *
 * Microsoft Support, "Group or ungroup shapes, pictures, or other objects":
 * after grouping you may still "select any single item within the group". This
 * used to answer `null` for every group that held no connector, so a drag on a
 * member of an ordinary group did nothing at all and said nothing either --
 * which across the owner's corpus was every group there is
 * (SHAPE-IMPORT-AUDIT.md §34). The server's two gates were widened with it, so
 * the move the canvas previews is the move the package receives.
 *
 * The clamp is the stored-page validator's own rule: a member stays inside the
 * group's child box. A connector is never dragged directly -- it is a route
 * between two members, recomputed from them.
 */
export function documentShapeGroupMoveMember(
  value: DocumentShapeAttrsInput,
  memberId: string,
  xEmu: number,
  yEmu: number,
): DocumentShapeAttrs | null {
  const group = normalizeDocumentShapeAttrs(value);
  const members = group.groupMembers;
  const coordinateWidth = group.groupCoordinateWidthEmu;
  const coordinateHeight = group.groupCoordinateHeightEmu;
  if (
    !members ||
    typeof coordinateWidth !== "number" ||
    typeof coordinateHeight !== "number" ||
    !Number.isSafeInteger(xEmu) ||
    !Number.isSafeInteger(yEmu)
  ) {
    return null;
  }
  const selected = members.find((member) => member.memberId === memberId);
  if (!selected || selected.kind === "connector") return null;
  const moved = members.map((member) =>
    member.memberId === memberId
      ? {
          ...member,
          xEmu: Math.min(coordinateWidth - member.widthEmu, Math.max(0, xEmu)),
          yEmu: Math.min(
            coordinateHeight - member.heightEmu,
            Math.max(0, yEmu),
          ),
        }
      : member,
  );
  const groupMembers = documentConnectorCanvasMembers(moved);
  if (!groupMembers) return null;
  const next = normalizeDocumentShapeAttrs({ ...group, groupMembers });
  return next.groupMembers?.length === groupMembers.length ? next : null;
}

/** The keys a pane writes when it picks a member's own fill. */
const DOCUMENT_SHAPE_MEMBER_FILL_KEYS = ["fill", "fillGradient", "fillPattern", "fillPicture", "fillTransparency"] as const;

/**
 * Paint or format ONE member of an editable group -- a SmartArt node, a
 * flowchart box -- as the leaf shape it is, the way a slide or a worksheet
 * writes the selected shape inside a graphic. The member's own fill, outline
 * and text paint are laid over a leaf of its preset, the leaf writer decides
 * the change, and only that member takes it; the group's other members and
 * its frame stay as they were.
 */
export function documentShapeGroupMemberPatch(
  value: DocumentShapeAttrsInput,
  memberId: string,
  patchLeaf: (
    leaf: DocumentShapeAttrs,
  ) => Readonly<Partial<DocumentShapeAttrs>> | null,
): Readonly<Pick<DocumentShapeAttrs, "groupMembers">> | null {
  const group = normalizeDocumentShapeAttrs(value);
  const members = group.groupMembers;
  if (!members || !documentShapeGroupIsEditable(group)) return null;
  const member = members.find((candidate) => candidate.memberId === memberId);
  if (!member) return null;
  // The leaf the pane edits is the member as it is DRAWN: wearing its part
  // of its group's fill, when it wears one.
  const {
    memberId: _memberId,
    drawingId: _drawingId,
    kind: _kind,
    connectorPreset: _connectorPreset,
    startMemberId: _startMemberId,
    startConnectionSite: _startConnectionSite,
    endMemberId: _endMemberId,
    endConnectionSite: _endConnectionSite,
    fillFromGroup: _fillFromGroup,
    groupFill: _groupFill,
    groupFillFrame: _groupFillFrame,
    ...leafFields
  } = documentShapeMemberWithGroupPaint(group, member);
  const leaf = normalizeDocumentShapeAttrs({
    ...documentShapeDefaultAttrs(member.preset),
    ...leafFields,
  });
  const patch = patchLeaf(leaf);
  if (!patch) return null;
  const updated: Record<string, unknown> = { ...member };
  if (DOCUMENT_SHAPE_MEMBER_FILL_KEYS.some((key) => Object.hasOwn(patch, key))) {
    // A fill picked for THIS member is its own: it no longer wears its
    // group's (`a:grpFill` gives way to the fill written for it).
    delete updated.fillFromGroup;
    delete updated.groupFill;
    delete updated.groupFillFrame;
  }
  for (const [key, next] of Object.entries(patch)) {
    if (next === null || next === undefined) delete updated[key];
    else updated[key] = next;
  }
  const groupMembers = members.map((candidate) =>
    candidate.memberId === memberId
      ? (updated as unknown as DocumentShapeGroupMember)
      : candidate,
  );
  const next = normalizeDocumentShapeAttrs({ ...group, groupMembers });
  return next.groupMembers?.length === members.length
    ? { groupMembers: next.groupMembers }
    : null;
}

function groupMemberTextPatch(
  value: DocumentShapeAttrsInput,
  memberId: string,
  text: string,
  connectorsOnly: boolean,
  runs?: readonly TextRun[] | null,
): DocumentShapeAttrs | null {
  const group = normalizeDocumentShapeAttrs(value);
  const members = group.groupMembers;
  if (!members) return null;
  const label = normalizeDocumentShapeText(text);
  let found = false;
  const groupMembers = members.map((member) => {
    if (
      member.memberId !== memberId ||
      (connectorsOnly && member.kind !== "connector") ||
      // A picture draws its bytes; it has no `w:txbxContent` for a label to
      // be written into, and the page model refuses a picture member carrying
      // text at all. Leaving it out here is what keeps a mistyped label from
      // being planned as an edit the writer would then decline.
      member.kind === "picture"
    ) {
      return member;
    }
    found = true;
    // The imported `bodyPr` + paragraph reading described the words it was read
    // WITH; retyping them makes it stale, so it goes with the old text.
    // The runs the field typed go with the words; retyped words without
    // runs carry none (the normalizer drops runs that no longer spell them).
    const { textBody: _replaced, textRuns: _formatted, ...rest } = member;
    return runs?.length ? { ...rest, text: label, textRuns: runs } : { ...rest, text: label };
  });
  if (!found) return null;
  const next = normalizeDocumentShapeAttrs({ ...group, groupMembers });
  return next.groupMembers?.length === groupMembers.length ? next : null;
}

export function documentShapeConnectorCanvasLabelPatch(
  value: DocumentShapeAttrsInput,
  memberId: string,
  text: string,
): DocumentShapeAttrs | null {
  return groupMemberTextPatch(value, memberId, text, true);
}

/**
 * One group member's own label. A member's text IS its shape text — the same
 * attribute a standalone shape carries and the same one the DOCX writer emits
 * as that member's `<wps:txbx>` — so editing a flowchart box, a mind-map branch
 * or a SmartArt node all commit through this single patch.
 */
export function documentShapeGroupMemberTextPatch(
  value: DocumentShapeAttrsInput,
  memberId: string,
  text: string,
  runs?: readonly TextRun[] | null,
): DocumentShapeAttrs | null {
  return groupMemberTextPatch(value, memberId, text, false, runs);
}

/** Plan one native group from independently-authored, fully-owned shapes. */
export function documentShapeGroupAttrs(
  inputs: readonly DocumentShapeGroupInput[],
): DocumentShapeAttrs | null {
  if (inputs.length < 2 || inputs.length > DOCUMENT_SHAPE_MAX_GROUP_MEMBERS) {
    return null;
  }
  const normalized = inputs.map((input) => {
    const chart =
      input.chart && typeof input.chart === "object" && !Array.isArray(input.chart)
        ? input.chart
        : null;
    const raw = chart ? documentChartGroupFrameAttrs(input.attrs) : input.attrs;
    return { pos: input.pos, raw, chart, attrs: normalizeDocumentShapeAttrs(raw) };
  });
  // A group of charts alone has no shape to stand in: Group takes a chart
  // WITH the shapes and text boxes it is grouped with.
  if (normalized.every(({ chart }) => chart !== null)) return null;
  if (
    normalized.some(
      ({ pos, raw, attrs }) =>
        !Number.isSafeInteger(pos) ||
        pos < 0 ||
        (raw.groupMembers !== undefined && raw.groupMembers !== null) ||
        attrs.source !== null ||
        // A member carries its WordArt, gradient and shadow (below): a shape
        // a Shape Styles preset painted groups as Word's does (it was
        // refused, and Group did nothing; live audit WUA 020).
        attrs.zOrderStackComplete !== true ||
        typeof attrs.relativeHeight !== "number" ||
        typeof attrs.linkedTextBoxStoryId === "number",
    )
  ) {
    return null;
  }
  const positions = new Set(normalized.map(({ pos }) => pos));
  if (positions.size !== normalized.length) return null;
  const host = normalized.find(({ chart }) => chart === null)!;
  const minX = Math.min(...normalized.map(({ attrs }) => attrs.xEmu));
  const minY = Math.min(...normalized.map(({ attrs }) => attrs.yEmu));
  const maxX = Math.max(
    ...normalized.map(({ attrs }) => attrs.xEmu + attrs.widthEmu),
  );
  const maxY = Math.max(
    ...normalized.map(({ attrs }) => attrs.yEmu + attrs.heightEmu),
  );
  const coordinateWidthEmu = maxX - minX;
  const coordinateHeightEmu = maxY - minY;
  if (
    coordinateWidthEmu < DOCUMENT_SHAPE_MIN_EDGE_EMU ||
    coordinateWidthEmu > DOCUMENT_SHAPE_MAX_EDGE_EMU ||
    coordinateHeightEmu < DOCUMENT_SHAPE_MIN_EDGE_EMU ||
    coordinateHeightEmu > DOCUMENT_SHAPE_MAX_EDGE_EMU
  ) {
    return null;
  }
  const ordered = [...normalized].sort(
    (left, right) =>
      left.attrs.relativeHeight! - right.attrs.relativeHeight! ||
      left.pos - right.pos,
  );
  // A member is the shape it was, less what the group now owns (its place in
  // the floating stack, its wrap, its source): its words keep their size,
  // weight, face and paint, so Ungroup gives back the shape that was grouped
  
  
  const groupMembers: DocumentShapeGroupMember[] = ordered.map(
    ({ attrs, chart }, index) => ({
      memberId: `member-${index + 1}`,
      // A chart member draws its chart at its box (no label, no paint).
      ...(chart ? { kind: "chart" as const, chart } : {}),
      preset: attrs.preset,
      ...(attrs.customPath ? { customPath: attrs.customPath } : {}),
      ...groupMemberInkAttrs(attrs),
      xEmu: attrs.xEmu - minX,
      yEmu: attrs.yEmu - minY,
      widthEmu: attrs.widthEmu,
      heightEmu: attrs.heightEmu,
      rotation: attrs.rotation,
      flipH: attrs.flipH,
      flipV: attrs.flipV,
      ...(typeof attrs.lockAspectRatio === "boolean"
        ? { lockAspectRatio: attrs.lockAspectRatio }
        : {}),
      isTextBox: attrs.isTextBox,
      text: attrs.text,
      ...(attrs.textBody ? { textBody: attrs.textBody } : {}),
      fill: attrs.fill,
      ...(attrs.fillGradient ? { fillGradient: attrs.fillGradient } : {}),
      ...(attrs.shapeShadow ? { shapeShadow: attrs.shapeShadow } : {}),
      ...(attrs.shapeEffects ? { shapeEffects: attrs.shapeEffects } : {}),
      outline: attrs.outline,
      outlineWidthPt: attrs.outlineWidthPt,
      outlineStyle: attrs.outlineStyle,
      ...(attrs.outlineAlign ? { outlineAlign: attrs.outlineAlign } : {}),
      ...officeStrokeDetail(attrs),
      effect: attrs.effect,
      // ...and its WordArt, which Ungroup gives back with the rest.
      ...(attrs.textPaint ? { textPaint: attrs.textPaint } : {}),
      textFill: attrs.textFill,
      textOutline: attrs.textOutline,
      textOutlineWidthPt: attrs.textOutlineWidthPt,
      textEffect: attrs.textEffect,
      textFontSizePt: attrs.textFontSizePt,
      textFontFamily: attrs.textFontFamily,
      textBold: attrs.textBold,
      ...(attrs.textItalic ? { textItalic: true } : {}),
      ...(attrs.textUnderline !== "none" ? { textUnderline: attrs.textUnderline } : {}),
      ...(attrs.textStrike ? { textStrike: true } : {}),
      textDirection: attrs.textDirection,
      textVerticalAlign: attrs.textVerticalAlign,
      ...(attrs.textHorizontalAlign !== "center"
        ? { textHorizontalAlign: attrs.textHorizontalAlign }
        : {}),
      textWrap: attrs.textWrap,
      altText: attrs.altText,
    }),
  );
  const group = normalizeDocumentShapeAttrs({
    preset: "rect",
    xEmu: minX,
    yEmu: minY,
    widthEmu: coordinateWidthEmu,
    heightEmu: coordinateHeightEmu,
    rotation: 0,
    flipH: false,
    flipV: false,
    relativeHeight: Math.max(
      ...ordered.map(({ attrs }) => attrs.relativeHeight!),
    ),
    zOrderStackComplete: true,
    lockAspectRatio: false,
    text: "",
    fill: "none",
    outline: "none",
    outlineWidthPt: 0,
    outlineStyle: "solid",
    source: null,
    groupCoordinateWidthEmu: coordinateWidthEmu,
    groupCoordinateHeightEmu: coordinateHeightEmu,
    // The group wraps the text as its shapes did: Word's group keeps its
    // shapes' layout. Without it two In Front of Text rectangles became a
    // square-wrapped group, and the body below moved out of their way onto
    // a second page (WUA 167). Shapes that differ take the first one's,
    // whose place the group takes; an In Line shape has no floating layout
    // to give.
    // A chart takes the group's layout from the shapes it joins.
    ...(host.attrs.wrapMode !== "inline" ? { wrapMode: host.attrs.wrapMode } : {}),
    groupMembers,
  });
  return group.groupMembers?.length === groupMembers.length ? group : null;
}

/**
 * Expand one authored group back into independent native shapes. The group's
 * outer move/resize/rotate/flip is composed into each child deterministically.
 * A group holding a chart answers through {@link documentShapeUngroupMembers}
 * (its chart is no shape); here each member's frame is what Ungroup places.
 */
export function documentShapeUngroupAttrs(
  value: DocumentShapeAttrsInput,
): readonly DocumentShapeAttrs[] | null {
  return documentShapeUngroupMembers(value)?.map(({ attrs }) => attrs) ?? null;
}

/** One object Ungroup gives back: a shape, or a chart at `attrs`' frame. */
export interface DocumentShapeUngroupedMember {
  /** The shape -- or, for a chart, its frame: offsets, size, layer, wrap. */
  readonly attrs: DocumentShapeAttrs;
  /** A chart member's chart (`DocumentShapeGroupMember.chart`), else null. */
  readonly chart: Readonly<Record<string, unknown>> | null;
}

/**
 * Word's Ungroup "leaves the individual objects in their current location":
 * every member as the object it was, a chart as a chart. A chart cannot turn
 * or mirror, so a chart member of a turned group keeps its centre and stands
 * upright.
 */
export function documentShapeUngroupMembers(
  value: DocumentShapeAttrsInput,
): readonly DocumentShapeUngroupedMember[] | null {
  const group = normalizeDocumentShapeAttrs(value);
  const members = group.groupMembers;
  const coordinateWidthEmu = group.groupCoordinateWidthEmu;
  const coordinateHeightEmu = group.groupCoordinateHeightEmu;
  if (
    !members?.length ||
    members.some((member) => member.kind === "connector") ||
    // A PICTURE member has no standalone `documentShape` to become: a
    // `documentShape` holds no picture, and the per-member record this
    // function builds below carries neither `src` nor a crop. Ungrouping such
    // a group would therefore drop the picture silently. Word's own Ungroup
    // "leaves the individual objects in their current location", picture
    // included, so the honest answer here is to refuse until a released
    // picture has somewhere to land -- not to release the shapes and lose the
    // picture (SHAPE-IMPORT-AUDIT.md §58).
    //
    // This is belt and braces today: every group holding a picture comes from
    // a package and so already fails the `source !== null` test below. It is
    // here so the picture is safe by its OWN rule rather than by a rule about
    // where the group came from.
    members.some((member) => member.kind === "picture") ||
    group.source !== null ||
    typeof coordinateWidthEmu !== "number" ||
    typeof coordinateHeightEmu !== "number"
  ) {
    return null;
  }
  const scaleX = group.widthEmu / coordinateWidthEmu;
  const scaleY = group.heightEmu / coordinateHeightEmu;
  if (!Number.isFinite(scaleX) || !Number.isFinite(scaleY)) return null;
  const groupFrame = {
    left: group.xEmu, top: group.yEmu, width: group.widthEmu, height: group.heightEmu,
    rotation: group.rotation, flipH: group.flipH, flipV: group.flipV,
  };
  const reflected = group.flipH !== group.flipV;
  const fallbackHeight =
    group.relativeHeight ?? DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT;
  const expanded = members.map((member, index) => {
    // The member's centre in the group's box, mirrored and turned with it.
    const localCenterX = (member.xEmu + member.widthEmu / 2) * scaleX;
    const localCenterY = (member.yEmu + member.heightEmu / 2) * scaleY;
    const { x: memberCenterX, y: memberCenterY } = officeFramePlace(
      groupFrame,
      group.widthEmu ? localCenterX / group.widthEmu : 0.5,
      group.heightEmu ? localCenterY / group.heightEmu : 0.5,
    );
    const memberWidthEmu = Math.round(member.widthEmu * scaleX);
    const memberHeightEmu = Math.round(member.heightEmu * scaleY);
    const requested = {
      preset: member.preset,
      // The member's own outline: a freeform (a stroke the pen drew among
      // them) is drawn by it. Left out, Ungroup gave back the preset's
      // rectangle in its place (WUA 092).
      ...(member.customPath ? { customPath: member.customPath } : {}),
      ...(member.adjustments ? { adjustments: member.adjustments } : {}),
      ...groupMemberInkAttrs(member),
      xEmu: Math.round(memberCenterX - memberWidthEmu / 2),
      yEmu: Math.round(memberCenterY - memberHeightEmu / 2),
      widthEmu: memberWidthEmu,
      heightEmu: memberHeightEmu,
      rotation:
        group.rotation + (reflected ? -member.rotation : member.rotation),
      flipH: group.flipH !== member.flipH,
      flipV: group.flipV !== member.flipV,
      relativeHeight: member.relativeHeight ?? fallbackHeight + index,
      zOrderStackComplete: true,
      ...(typeof member.lockAspectRatio === "boolean"
        ? { lockAspectRatio: member.lockAspectRatio }
        : {}),
      isTextBox: member.isTextBox,
      text: member.text,
      fill: member.fill,
      ...(member.fillGradient ? { fillGradient: member.fillGradient } : {}),
      ...(member.shapeShadow ? { shapeShadow: member.shapeShadow } : {}),
      ...(member.shapeEffects ? { shapeEffects: member.shapeEffects } : {}),
      outline: member.outline,
      outlineWidthPt: member.outlineWidthPt,
      outlineStyle: member.outlineStyle,
      ...(member.outlineDash ? { outlineDash: member.outlineDash } : {}),
      ...officeStrokeDetail(member),
      ...(member.outlineSketch ? { outlineSketch: member.outlineSketch } : {}),
      effect: member.effect,
      ...(member.textPaint ? { textPaint: member.textPaint } : {}),
      textFill: member.textFill,
      textOutline: member.textOutline,
      textOutlineWidthPt: member.textOutlineWidthPt,
      textEffect: member.textEffect,
      textFontSizePt: member.textFontSizePt,
      textFontFamily: member.textFontFamily,
      textBold: member.textBold,
      textItalic: member.textItalic,
      textUnderline: member.textUnderline,
      textStrike: member.textStrike,
      textDirection: member.textDirection,
      textVerticalAlign: member.textVerticalAlign,
      textHorizontalAlign: member.textHorizontalAlign,
      textWrap: member.textWrap,
      altText: member.altText,
      // A shape that left the group wraps as the group did (WUA 167): a
      // member reads the default (square) unless one was stated for it.
      wrapMode:
        group.wrapMode !== "inline" && (member.wrapMode ?? "square") === "square"
          ? group.wrapMode
          : member.wrapMode,
      ...(typeof member.linkedTextBoxStoryId === "number" &&
      typeof member.linkedTextBoxSequence === "number"
        ? {
            linkedTextBoxStoryId: member.linkedTextBoxStoryId,
            linkedTextBoxSequence: member.linkedTextBoxSequence,
          }
        : {}),
      source: null,
    };
    const attrs = normalizeDocumentShapeAttrs(requested);
    return attrs.xEmu === requested.xEmu &&
      attrs.yEmu === requested.yEmu &&
      attrs.widthEmu === requested.widthEmu &&
      attrs.heightEmu === requested.heightEmu
      ? attrs
      : null;
  });
  if (!expanded.every((attrs): attrs is DocumentShapeAttrs => attrs !== null)) return null;
  return expanded.map((attrs, index) => {
    const member = members[index]!;
    return member.kind === "chart" && member.chart
      ? { attrs: { ...attrs, rotation: 0, flipH: false, flipV: false }, chart: member.chart }
      : { attrs, chart: null };
  });
}

/**
 * Resolve the next persisted native-shape position without creating no-op
 * history. The same normalizer that guards imported/model attributes owns the
 * coordinate bounds, so menu nudges can never write a value that reload or
 * DOCX export would later coerce.
 */
export function documentShapePositionPatch(
  value: DocumentShapeAttrsInput,
  action: DocumentShapePositionAction,
): Readonly<Pick<DocumentShapeAttrs, "xEmu" | "yEmu">> | null {
  const current = normalizeDocumentShapeAttrs(value);
  // An inline shape has no position of its own: its paragraph places it.
  if (current.wrapMode === "inline") return null;
  const deltaX =
    action === "left"
      ? -DOCUMENT_SHAPE_POSITION_STEP_EMU
      : action === "right"
        ? DOCUMENT_SHAPE_POSITION_STEP_EMU
        : 0;
  const deltaY =
    action === "up"
      ? -DOCUMENT_SHAPE_POSITION_STEP_EMU
      : action === "down"
        ? DOCUMENT_SHAPE_POSITION_STEP_EMU
        : 0;
  const next = normalizeDocumentShapeAttrs({
    ...current,
    xEmu: action === "reset" ? 0 : current.xEmu + deltaX,
    yEmu: action === "reset" ? 0 : current.yEmu + deltaY,
  });
  if (next.xEmu === current.xEmu && next.yEmu === current.yEmu) {
    return null;
  }
  return { xEmu: next.xEmu, yEmu: next.yEmu };
}

/**
 * Word's arrows on a selected floating shape (the shared reading,
 * `lib/office-object-nudge`): move, resize or rotate it through the same
 * normalizer every other writer uses. A shape in line with text is Word's to
 * place, not OfficeArt's, so it is not nudged (null: the arrow moves the caret).
 */
export function documentShapeKeyPatch(
  value: DocumentShapeAttrsInput,
  action: import("./office-object-nudge").OfficeObjectKeyAction,
): Readonly<Partial<Pick<DocumentShapeAttrs, "xEmu" | "yEmu" | "widthEmu" | "heightEmu" | "rotation">>> | null {
  const current = normalizeDocumentShapeAttrs(value);
  if (current.wrapMode === "inline") return null;
  if (action.kind === "move") {
    const next = normalizeDocumentShapeAttrs({
      ...current,
      xEmu: current.xEmu + documentShapePxToEmu(action.dx),
      yEmu: current.yEmu + documentShapePxToEmu(action.dy),
    });
    return next.xEmu === current.xEmu && next.yEmu === current.yEmu
      ? null
      : { xEmu: next.xEmu, yEmu: next.yEmu };
  }
  if (action.kind === "resize") {
    return action.dw !== 0
      ? documentShapeSizeEmuPatch(current, "width", current.widthEmu + documentShapePxToEmu(action.dw))
      : documentShapeSizeEmuPatch(current, "height", current.heightEmu + documentShapePxToEmu(action.dh));
  }
  const next = normalizeDocumentShapeAttrs({
    ...current,
    rotation: current.rotation + action.degrees,
  });
  return next.rotation === current.rotation ? null : { rotation: next.rotation };
}

/** Align one floating shape inside the authenticated printable page width. */
export function documentShapeAlignPatch(
  value: DocumentShapeAttrsInput,
  action: DocumentShapeAlignAction,
  printableWidthPx: number,
): Readonly<Pick<DocumentShapeAttrs, "xEmu" | "yEmu">> | null {
  if (!Number.isFinite(printableWidthPx) || printableWidthPx <= 0) return null;
  const current = normalizeDocumentShapeAttrs(value);
  if (current.groupMembers?.length && !documentShapeGroupIsEditable(current)) {
    return null;
  }
  // An inline shape is aligned by its paragraph, never by an offset it lacks.
  if (current.wrapMode === "inline") return null;
  const printableWidthEmu = Math.round(
    Math.min(5_000, printableWidthPx) * DOCUMENT_SHAPE_EMU_PER_PX,
  );
  const maximumX = Math.max(0, printableWidthEmu - current.widthEmu);
  const xEmu =
    action === "left"
      ? 0
      : action === "right"
        ? maximumX
        : Math.round(maximumX / 2);
  return xEmu === current.xEmu
    ? null
    : { xEmu, yEmu: current.yEmu };
}

/** Resolve a clockwise rotation delta (or reset) without an empty Undo step. */
export function documentShapeRotationPatch(
  value: DocumentShapeAttrsInput,
  action: DocumentShapeRotationAction,
): Readonly<Pick<DocumentShapeAttrs, "rotation">> | null {
  const current = normalizeDocumentShapeAttrs(value);
  const next = normalizeDocumentShapeAttrs({
    ...current,
    rotation: action === 0 ? 0 : current.rotation + action,
  });
  return next.rotation === current.rotation
    ? null
    : { rotation: next.rotation };
}

/** Toggle one native DrawingML mirror flag through the shared flip authority. */
export function documentShapeFlipPatch(
  value: DocumentShapeAttrsInput,
  axis: DocumentShapeFlipAxis,
): Readonly<Pick<DocumentShapeAttrs, DocumentShapeFlipAxis>> {
  const current = normalizeDocumentShapeAttrs(value);
  return { [axis]: toggleFlip(current[axis]) } as Readonly<
    Pick<DocumentShapeAttrs, DocumentShapeFlipAxis>
  >;
}

/** Toggle the persisted DrawingML `a:spLocks/@noChangeAspect` receipt. */
export function documentShapeAspectRatioLockPatch(
  value: DocumentShapeAttrsInput,
): Readonly<Pick<DocumentShapeAttrs, "lockAspectRatio">> {
  const current = normalizeDocumentShapeAttrs(value);
  return { lockAspectRatio: !current.lockAspectRatio };
}

export type DocumentShapePaintPatch = Readonly<
  Partial<
    Pick<
      DocumentShapeAttrs,
      "fill" | "outline" | "outlineWidthPt" | "outlineStyle" | "fillGradient"
    >
  >
>;

export interface DocumentShapePaintRequest {
  readonly fill?: unknown;
  readonly outline?: unknown;
  readonly outlineWidthPt?: unknown;
  readonly outlineStyle?: unknown;
  /** Shape Outline > Dashes, one of Office's eight (`@/lib/office-stroke-dash`):
   * the family, the exact dash and Round Dot's round cap together. */
  readonly outlineMenuDash?: unknown;
  /** Shape Outline > Sketched: "none" or one of Office's three. */
  readonly outlineSketch?: unknown;
}

/** Change one leaf preset while preserving its paint, text and geometry. */
export function documentShapePresetPatch(
  value: DocumentShapeAttrsInput,
  preset: unknown,
): Readonly<Pick<DocumentShapeAttrs, "preset">> | null {
  const current = normalizeDocumentShapeAttrs(value);
  if (
    current.groupMembers?.length ||
    // A line carries no words and no WordArt (the save refuses both on one:
    // "line shapes cannot carry text"), so a shape that holds either keeps
    // its outline: Word's Change Shape does not turn a shape's words into a
    // line's.
    ((current.textPaint || current.text) &&
      isDocumentShapePreset(preset) &&
      officeShapeDefinition(preset).lineLike) ||
    !isDocumentShapePreset(preset) ||
    current.preset === preset
  ) {
    return null;
  }
  return { preset };
}

/**
 * A shape's two `a:ln` decorations, as the arrow picker reads and writes them.
 *
 * A document shape carries its ends itself (Docs has no preserve-only outline:
 * the writer rebuilds `a:ln` from the page), so the picker reaches every shape
 * whose outline HAS ends to decorate — a connector, an `arc`, a bracket, a
 * callout's leader, an open freeform — and none of the closed ones, which is
 * where Office greys its own arrow settings out (SHAPE-IMPORT-AUDIT.md §7.10).
 */
export function documentShapeLineEnds(
  value: DocumentShapeAttrsInput,
): OfficeLineEndsState {
  const current = normalizeDocumentShapeAttrs(value);
  const implied = normalizePresentationConnectorSpec({
    shape: officeShapeDefinition(current.preset).previewKind,
  });
  return readOfficeLineEnds(
    current as unknown as Record<string, unknown>,
    implied ?? undefined,
  );
}

/** Whether this shape's own geometry has ends for an arrowhead to sit on. */
export function documentShapeTakesLineEnds(
  value: DocumentShapeAttrsInput,
): boolean {
  const current = normalizeDocumentShapeAttrs(value);
  if (current.groupMembers?.length) return false;
  if (current.customPath) return current.customPath.closed === false;
  const kind = officeShapeDefinition(current.preset).previewKind;
  if (normalizePresentationConnectorSpec({ shape: kind })) return true;
  const geometry = presentationPresetGeometryForKind(
    kind,
    documentShapeEmuToPx(current.widthEmu),
    documentShapeEmuToPx(current.heightEmu),
    current.adjustments,
  );
  return geometry ? officeGeometryTakesLineEnds(geometry.paths) : false;
}

/**
 * One arrow pick → the attrs to store.
 *
 * Both types are written EXPLICITLY, `none` included: an absent end field
 * inherits the authenticated original in `patchback_docx`, which is how a Word
 * shape keeps an arrowhead nobody touched — and would also be how one the user
 * just removed came back.
 */
export function documentShapeLineEndsPatch(
  value: DocumentShapeAttrsInput,
  ends: OfficeLineEndsState,
): Readonly<Partial<DocumentShapeAttrs>> | null {
  const current = normalizeDocumentShapeAttrs(value);
  if (!documentShapeTakesLineEnds(current)) return null;
  const patch = officeShapeLineEndsPatch(current.preset, ends);
  if (!patch) return null;
  const next: Partial<DocumentShapeAttrs> = {
    // A connector preset names WHICH ends are decorated, so the pick may move
    // the preset — to the one this shape's own exported file imports back as.
    ...(patch.preset === current.preset ? {} : { preset: patch.preset }),
    headEnd: patch.headEnd,
    tailEnd: patch.tailEnd,
    // An absent size stays absent: `@w`/`@len` have no schema default.
    headEndWidth: patch.headEndWidth,
    headEndLength: patch.headEndLength,
    tailEndWidth: patch.tailEndWidth,
    tailEndLength: patch.tailEndLength,
  };
  const changed = (Object.keys(next) as (keyof DocumentShapeAttrs)[]).some(
    (key) => next[key] !== current[key],
  );
  return changed ? next : null;
}









function documentShapeOutlineShown(
  current: Pick<DocumentShapeAttrs, "outline" | "outlineWidthPt">,
  request: {
    readonly outline?: unknown;
    readonly outlineWidthPt?: unknown;
    readonly outlineStyle?: unknown;
    readonly outlineSketch?: unknown;
  },
): { outline?: string; outlineWidthPt?: number } {
  const implied: { outline?: string; outlineWidthPt?: number } = {};
  const colour = Object.hasOwn(request, "outline")
    ? normalizeDocumentShapeColor(request.outline, current.outline)
    : current.outline;
  const width = Object.hasOwn(request, "outlineWidthPt")
    ? finiteNumber(request.outlineWidthPt, current.outlineWidthPt)
    : current.outlineWidthPt;
  const dash =
    request.outlineStyle !== undefined ||
    (request.outlineSketch !== undefined && normalizeOfficeLineSketch(request.outlineSketch) !== undefined);
  if (
    colour === "none" &&
    !Object.hasOwn(request, "outline") &&
    ((width > 0 && Object.hasOwn(request, "outlineWidthPt")) || dash)
  ) {
    implied.outline = DOCUMENT_SHAPE_DEFAULT_OUTLINE;
  }
  if (
    !Object.hasOwn(request, "outlineWidthPt") &&
    (implied.outline !== undefined || colour !== "none") &&
    (Object.hasOwn(request, "outline") || dash) &&
    !(width > 0)
  ) {
    implied.outlineWidthPt = officeOutlineWeightWhenShown(width, OFFICE_OUTLINE_SHOWN_WEIGHT_PT);
  }
  return implied;
}

/**
 * Resolve leaf fill/outline changes through the persisted DrawingML model.
 * Invalid values collapse to the current value, group paint remains member-
 * owned, and an unchanged request cannot create an empty history item.
 */
export function documentShapePaintPatch(
  value: DocumentShapeAttrsInput,
  picked: DocumentShapePaintRequest,
): DocumentShapePaintPatch | null {
  const current = normalizeDocumentShapeAttrs(value);
  if (current.groupMembers?.length) return null;
  const menuDash = (OFFICE_MENU_DASHES as readonly unknown[]).includes(picked.outlineMenuDash)
    ? officeMenuDashFields(picked.outlineMenuDash as OfficeMenuDash)
    : null;
  const stated: DocumentShapePaintRequest = menuDash ? { ...picked, outlineStyle: menuDash.style } : picked;
  const request: DocumentShapePaintRequest = {
    ...stated,
    ...documentShapeOutlineShown(current, stated),
  };

  const next = {
    fill: Object.hasOwn(request, "fill")
      ? normalizeDocumentShapeColor(request.fill, current.fill)
      : current.fill,
    outline: Object.hasOwn(request, "outline")
      ? normalizeDocumentShapeColor(request.outline, current.outline)
      : current.outline,
    outlineWidthPt: Object.hasOwn(request, "outlineWidthPt")
      ? normalizeDocumentShapeOutlineWidthPt(request.outlineWidthPt, current.outlineWidthPt)
      : current.outlineWidthPt,
    outlineStyle:
      request.outlineStyle === "dashed" || request.outlineStyle === "dotted"
        ? request.outlineStyle
        : request.outlineStyle === "solid"
          ? "solid"
          : current.outlineStyle,
  } satisfies Pick<
    DocumentShapeAttrs,
    "fill" | "outline" | "outlineWidthPt" | "outlineStyle"
  >;

  const patch: Partial<
    Pick<
      DocumentShapeAttrs,
      | "fill"
      | "outline"
      | "outlineWidthPt"
      | "outlineStyle"
      | "outlineDash"
      | "outlineCap"
      | "outlineSketch"
      | "fillGradient"
      | "fillPicture"
    >
  > = {};
  if (next.fill !== current.fill) patch.fill = next.fill;
  // Shape Fill's colour (or No Fill) replaces a picture fill.
  if (Object.hasOwn(request, "fill") && normalizeDocumentShapeColor(request.fill, "") && current.fillPicture) {
    patch.fillPicture = null;
    patch.fill = next.fill;
  }
  if (Object.hasOwn(request, "fill") && normalizeDocumentShapeColor(request.fill, "") && (current.fillGradient ||
    (current.source?.adapter === "docx" && value?.fillGradient === undefined))) {
    patch.fillGradient = null;
  }
  if (next.outline !== current.outline) patch.outline = next.outline;
  if (next.outlineWidthPt !== current.outlineWidthPt) {
    patch.outlineWidthPt = next.outlineWidthPt;
  }
  if (next.outlineStyle !== current.outlineStyle) {
    patch.outlineStyle = next.outlineStyle;
  }
  if (menuDash) {
    // One of Office's eight: its exact dash, and Round Dot's round cap (a cap
    // the earlier Round Dot set goes with it).
    if ((current.outlineDash ?? null) !== (menuDash.dash ?? null)) patch.outlineDash = menuDash.dash ?? null;
    if (menuDash.cap && current.outlineCap !== menuDash.cap) patch.outlineCap = menuDash.cap;
    if (!menuDash.cap && current.outlineCap === "round" && current.outlineDash === "sysDot") patch.outlineCap = null;
  } else if (
    // A Dashes pick states the style's own dash: the imported preset it
    // refined (Long Dash, Round Dot...) goes with the old pick.
    current.outlineDash &&
    (request.outlineStyle === "solid" ||
      request.outlineStyle === "dashed" ||
      request.outlineStyle === "dotted")
  ) {
    patch.outlineDash = null;
  }
  if (Object.hasOwn(request, "outlineSketch")) {
    const sketch = normalizeOfficeLineSketch(request.outlineSketch) ?? null;
    if ((current.outlineSketch ?? null) !== sketch) patch.outlineSketch = sketch;
  }
  return Object.keys(patch).length ? patch : null;
}

export type DocumentShapeFormatPatch = Readonly<
  Partial<
    Pick<
      DocumentShapeAttrs,
      | "fill"
      | "fillGradient"
      | "shapeShadow"
      | "shapeEffects"
      | "outline"
      | "outlineWidthPt"
      | "outlineStyle"
      | "isTextBox"
      | "effect"
      | "textFill"
      | "textPaint"
      | "textOutline"
      | "textOutlineWidthPt"
      | "textEffect"
      | "textFontSizePt"
      | "textFontFamily"
      | "textBold"
      | "textItalic"
      | "textUnderline"
      | "textStrike"
      | "textDirection"
      | "textVerticalAlign"
      | "textHorizontalAlign"
      | "textBody"
      | "textRuns"
      | "textWrap"
      | "textAutofit"
      | "keepTextFlat"
      | "altText"
      | "decorative"
      | "wrapMode"
    >
  >
>;

export interface DocumentShapeFormatRequest {
  /** Text Effects > Transform: a preset warp, or null for none. */
  readonly textWarp?: unknown;
  /** Format Shape > Line: compound, cap and join (null: the default). */
  readonly outlineCompound?: unknown;
  readonly outlineCap?: unknown;
  readonly outlineJoin?: unknown;
  readonly outlineMiterLimit?: unknown;
  readonly textPaint?: unknown;
  readonly fill?: unknown;
  readonly fillGradient?: unknown;
  /** Format Shape > Fill & Line: a pattern fill, the fill's and line's
   * Transparency, a gradient line (null: none). */
  readonly fillPattern?: unknown;
  /** Picture or texture fill: `{src, stretch: true}`, or null for none. */
  readonly fillPicture?: unknown;
  readonly fillTransparency?: unknown;
  readonly outlineTransparency?: unknown;
  readonly outlineGradient?: unknown;
  readonly shapeShadow?: unknown;
  readonly shapeEffects?: unknown;
  readonly outline?: unknown;
  readonly outlineWidthPt?: unknown;
  readonly outlineStyle?: unknown;
  readonly isTextBox?: unknown;
  readonly effect?: unknown;
  readonly textFill?: unknown;
  readonly textOutline?: unknown;
  readonly textOutlineWidthPt?: unknown;
  readonly textEffect?: unknown;
  readonly textFontSizePt?: unknown;
  readonly textFontFamily?: unknown;
  readonly textBold?: unknown;
  readonly textItalic?: unknown;
  readonly textUnderline?: unknown;
  readonly textStrike?: unknown;
  readonly textDirection?: unknown;
  readonly textVerticalAlign?: unknown;
  readonly textHorizontalAlign?: unknown;
  readonly textWrap?: unknown;
  readonly textAutofit?: unknown;
  /** Format Shape > 3-D Rotation > Keep text flat: true or false. */
  readonly keepTextFlat?: unknown;
  readonly altText?: unknown;
  readonly decorative?: unknown;
  readonly wrapMode?: unknown;
}

const DOCUMENT_SHAPE_FORMAT_KEYS = [
  "fill",
  "fillGradient",
  "fillPattern",
  "fillPicture",
  "fillTransparency",
  "outlineTransparency",
  "outlineGradient",
  "shapeShadow",
  "shapeEffects",
  "outline",
  "outlineWidthPt",
  "outlineStyle",
  "isTextBox",
  "effect",
  "textFill",
  "textPaint",
  "textOutline",
  "textOutlineWidthPt",
  "textEffect",
  "textFontSizePt",
  "textFontFamily",
  "textBold",
  "textItalic",
  "textUnderline",
  "textStrike",
  "textDirection",
  "textVerticalAlign",
  "textHorizontalAlign",
  "textWrap",
  "textAutofit",
  "keepTextFlat",
  "altText",
  "decorative",
  "wrapMode",
  "textWarp",
  "outlineCompound",
  "outlineCap",
  "outlineJoin",
  "outlineMiterLimit",
] as const;

/** Resolve every Shape Format writer through the same strict persisted model. */
export function documentShapeFormatPatch(
  value: DocumentShapeAttrsInput,
  request: DocumentShapeFormatRequest,
): DocumentShapeFormatPatch | null {
  const current = normalizeDocumentShapeAttrs(value);
  if (current.groupMembers?.length) return null;
  const validRequest = { ...request };
  if (current.textPaint && validRequest.textEffect === "softEdge") return null;
  // A line carries no WordArt (the save refuses it: "line shapes cannot
  // carry WordArt paint"); the WordArt gallery is greyed on one already.
  if (validRequest.textPaint && officeShapeDefinition(current.preset).lineLike) return null;
  for (const key of ["fill", "textFill"] as const) {
    if (Object.hasOwn(validRequest, key) && !normalizeDocumentShapeColor(validRequest[key], "")) {
      delete validRequest[key];
    }
  }
  if (Object.hasOwn(validRequest, "effect") && normalizeDocumentShapeEffect(validRequest.effect) !== validRequest.effect) {
    delete validRequest.effect;
  }
  if (Object.hasOwn(validRequest, "fillPicture") && validRequest.fillPicture !== null) {
    // Picture or texture fill: only a picture the model holds, and never on
    // a line (it encloses nothing to fill).
    if (!normalizeDocumentShapePictureFill(validRequest.fillPicture) || officeShapeDefinition(current.preset).lineLike) {
      return null;
    }
  }
  const fillPictureWritten = Object.hasOwn(validRequest, "fillPicture") && validRequest.fillPicture !== null;
  const effectiveRequest = {
    ...validRequest,
    // Any other fill replaces a picture fill (Office's one Fill choice); a
    // picture leaves the gradient and pattern it covers behind.
    ...((Object.hasOwn(validRequest, "fill") || Object.hasOwn(validRequest, "fillGradient") || Object.hasOwn(validRequest, "fillPattern")) &&
    !Object.hasOwn(validRequest, "fillPicture") &&
    current.fillPicture
      ? { fillPicture: null }
      : {}),
    ...(fillPictureWritten ? { fillGradient: null, fillPattern: null, fillTransparency: null } : {}),
    ...documentShapeOutlineShown(current, validRequest),
    ...(Object.hasOwn(validRequest, "fill") && !Object.hasOwn(validRequest, "fillGradient")
      ? { fillGradient: null }
      : {}),
    // A new fill replaces a pattern it does not restate and is opaque unless
    // it states its Transparency; a new line colour likewise drops a gradient
    // line and its Transparency (as the server writes them).
    ...((Object.hasOwn(validRequest, "fill") || Object.hasOwn(validRequest, "fillGradient")) &&
    !Object.hasOwn(validRequest, "fillPattern")
      ? { fillPattern: null }
      : {}),
    ...((Object.hasOwn(validRequest, "fill") || Object.hasOwn(validRequest, "fillGradient") || Object.hasOwn(validRequest, "fillPattern")) &&
    !Object.hasOwn(validRequest, "fillTransparency")
      ? { fillTransparency: null }
      : {}),
    ...(Object.hasOwn(validRequest, "outline") && !Object.hasOwn(validRequest, "outlineGradient")
      ? { outlineGradient: null }
      : {}),
    ...((Object.hasOwn(validRequest, "outline") || Object.hasOwn(validRequest, "outlineGradient")) &&
    !Object.hasOwn(validRequest, "outlineTransparency")
      ? { outlineTransparency: null }
      : {}),
    ...(Object.hasOwn(validRequest, "effect") && !Object.hasOwn(validRequest, "shapeShadow")
      ? { shapeShadow: null }
      : {}),
    // A Shape Effects preset replaces the effects the shape came with.
    ...(Object.hasOwn(validRequest, "effect") && !Object.hasOwn(validRequest, "shapeEffects")
      ? { shapeEffects: null }
      : {}),
  };
  // Text Fill ▸ No Fill, Text Outline and Text Effects are WordArt in Word
  // (`w14:textFill` / `w14:textOutline` / `w14:shadow|reflection|glow` on the
  // runs), and WordArt paint is what the page's save keeps on a shape: the
  // plain text keys take only a colour (`textFill`) and refused the whole page
  // for anything else ("non-default shape formatting is not supported by the
  // native save writer"). So a plain shape's first such pick starts its
  // WordArt paint from the text's own colour, and a line, which carries no
  // text and no WordArt, takes none.
  const requestedOutlineWidth = Number(validRequest.textOutlineWidthPt);
  const startsWordArt =
    !current.textPaint &&
    !Object.hasOwn(validRequest, "textPaint") &&
    ((Object.hasOwn(validRequest, "textOutline") && validRequest.textOutline !== "none") ||
      (Object.hasOwn(validRequest, "textEffect") && validRequest.textEffect !== "none") ||
      (Object.hasOwn(validRequest, "textOutlineWidthPt") &&
        Number.isFinite(requestedOutlineWidth) &&
        requestedOutlineWidth > 0) ||
      validRequest.textFill === "none");
  if (startsWordArt) {
    if (officeShapeDefinition(current.preset).lineLike || validRequest.textEffect === "softEdge") {
      return null;
    }
    effectiveRequest.textPaint = normalizeOfficeTextPaint({
      fill: current.textFill,
      outline: "none",
      outlineWidthPt: 0,
    });
  }
  const wordArtBase = current.textPaint ?? (effectiveRequest.textPaint as OfficeTextPaint | undefined);
  if (wordArtBase && !Object.hasOwn(validRequest, "textPaint") &&
      ["textFill", "textOutline", "textOutlineWidthPt", "textEffect"].some(key => Object.hasOwn(validRequest, key))) {
    // The Format pane's number field commits a string; the WordArt paint
    // only took a number, so "6" was skipped and the width normalized back
    
    const rawOutlineWidth = validRequest.textOutlineWidthPt;
    const outlineWidthPt =
      typeof rawOutlineWidth === "number"
        ? rawOutlineWidth
        : typeof rawOutlineWidth === "string" && rawOutlineWidth.trim() !== ""
          ? Number(rawOutlineWidth)
          : NaN;
    effectiveRequest.textPaint = patchOfficeTextPaint(wordArtBase, {
      ...(typeof validRequest.textFill === "string" ? { fill: validRequest.textFill } : {}),
      ...(typeof validRequest.textOutline === "string" ? { outline: validRequest.textOutline } : {}),
      ...(Number.isFinite(outlineWidthPt) ? { outlineWidthPt } : {}),
      ...(typeof validRequest.textEffect === "string" ? { effect: validRequest.textEffect } : {}),
    });
  }
  const requested: Record<string, unknown> = { ...current };
  for (const key of DOCUMENT_SHAPE_FORMAT_KEYS) {
    if (Object.hasOwn(effectiveRequest, key)) requested[key] = effectiveRequest[key];
  }
  const next = normalizeDocumentShapeAttrs(requested);
  const patch: Partial<
    Pick<DocumentShapeAttrs, (typeof DOCUMENT_SHAPE_FORMAT_KEYS)[number] | "textBody">
  > = {};
  for (const key of DOCUMENT_SHAPE_FORMAT_KEYS) {
    if (
      Object.hasOwn(effectiveRequest, key) &&
      (JSON.stringify(current[key] ?? null) !== JSON.stringify(next[key] ?? null) ||
        (["fillGradient", "shapeShadow", "textFill"].includes(key) &&
          current.source?.adapter === "docx" &&
          value?.[key] === undefined && effectiveRequest[key] !== undefined))
    ) {
      (patch as Record<string, unknown>)[key] = next[key] ?? null;
    }
  }
  // Aligning the whole shape's words aligns every paragraph the imported
  // body reading keeps, so the page draws what the file will say — also
  // when the shape's own scalar already says so but a paragraph does not
  
  const requestedAlign = effectiveRequest.textHorizontalAlign;
  if (
    !patch.textHorizontalAlign &&
    requestedAlign &&
    requestedAlign === next.textHorizontalAlign &&
    current.textBody?.paragraphs.some((paragraph) => (paragraph.align ?? "left") !== requestedAlign)
  ) {
    patch.textHorizontalAlign = next.textHorizontalAlign;
  }
  if (patch.textHorizontalAlign && current.textBody) {
    patch.textBody = {
      ...current.textBody,
      paragraphs: current.textBody.paragraphs.map((paragraph) => ({
        ...paragraph,
        align: patch.textHorizontalAlign!,
      })),
    };
  }
  return Object.keys(patch).length ? patch : null;
}

/**
 * The text box a plain click with Insert > Text Box's crosshair places, in
 * page px at 100 % (96 per inch): 3 in wide, so a line of words fits before
 * it wraps, and one line tall (its height then follows its words). A
 * vertical box is the same turned on its side.
 */
export function documentTextBoxClickSizePx(
  direction: DocumentShapeTextDirection = "horizontal",
): { readonly width: number; readonly height: number } {
  return direction === "horizontal" ? { width: 288, height: 40 } : { width: 40, height: 288 };
}

/**
 * A new text box (Insert > Text Box), drawn as the presentation's text box
 * is (`textBoxBlock`, lib/presentation-shape-buttons): no fill and no line,
 * its words from the top-left corner inside Office's default insets (0.1 in
 * left and right, 0.05 in top and bottom), and its height following them
 * (Resize shape to fit text, `a:spAutoFit`) while its width stays as drawn.
 * A vertical box keeps the height it was drawn at: its lines stack across
 * it, so its words never ask for more height.
 *
 * Every value is stated, so a text box made with the earlier defaults (a
 * slate outline, its words centred) keeps drawing as it was stored.
 */
export function documentShapeTextBoxDefaultAttrs(
  direction: DocumentShapeTextDirection = "horizontal",
): DocumentShapeAttrs {
  return normalizeDocumentShapeAttrs({
    ...documentShapeDefaultAttrs("rect"),
    isTextBox: true,
    textDirection: direction,
    fill: "none",
    outline: "none",
    textVerticalAlign: "top",
    textHorizontalAlign: "left",
    ...(direction === "horizontal" ? { textAutofit: "shape" } : {}),
  });
}

/**
 * The height Word's "Resize shape to fit text" (`a:spAutoFit`, Format Shape >
 * Text Box) gives a horizontal text box: its words wrapped at the box's inner
 * width (Office's default insets, 0.1 in left and right, 0.05 in top and
 * bottom), every paragraph at least one line, at the box's own font size --
 * the deck's AutoFit estimate (`presentation-autofit`), a paragraph with CJK
 * characters at a full em each. The editor refits the box to its drawn words
 * once the page is open (the node's `fitShapeToText`); this is the height an
 * export, a print and a page no editor has opened yet show, so no line is
 * cut off there. Never below the shape floor nor past its ceiling.
 */
export function documentTextBoxFittedHeightEmu(value: DocumentShapeAttrsInput): number {
  const attrs = normalizeDocumentShapeAttrs(value);
  const fontPx = (attrs.textFontSizePt * 96) / 72;
  const innerWidthPx = Math.max(1, documentShapeEmuToPx(attrs.widthEmu - 2 * 91_440));
  let textPx = 0;
  for (const paragraph of attrs.text.split("\n")) {
    const cjk = /[\u3000-\u303f\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/u.test(paragraph) ? 1 : undefined;
    textPx += estimateTextHeight(paragraph || " ", innerWidthPx, fontPx, DEFAULT_LINE_HEIGHT, cjk);
  }
  const needed = documentShapePxToEmu(Math.ceil(textPx)) + 2 * 45_720;
  return Math.min(DOCUMENT_SHAPE_MAX_EDGE_EMU, Math.max(DOCUMENT_SHAPE_MIN_EDGE_EMU, needed));
}

/**
 * WordArt — the shared primitive.
 *
 * Word's Insert → WordArt inserts a text box whose text carries WordArt STYLES
 * (text fill, text outline, text effects); the same styles are what the Shape
 * Format contextual tab's "WordArt Styles" group edits afterwards. MRAI's shape
 * node already models exactly that vocabulary (`textFill`, `textOutline`,
 * `textOutlineWidthPt`, `textEffect`), the Shape Format pane already edits it,
 * and the DOCX writer/importer already round-trip it as `w14:textFill` /
 * `w14:textOutline` / `w14:glow|shadow|reflection`.
 *
 * The Insert ribbon's WordArt gallery used to be a SECOND, thinner design —
 * run marks (font size / colour / highlight) with no outline, no effects and no
 * DOCX representation. This definition is the one gallery both entry points
 * now share: a style is expressed in the shape's own WordArt vocabulary, so
 * what the gallery inserts is what the Shape Format tab then edits.
 */
export interface DocumentWordArtStyle {
  readonly id: string;
  readonly textFill: string;
  readonly textOutline: string;
  readonly textOutlineWidthPt: number;
  readonly textEffect: DocumentShapeEffectPreset;
  readonly textFontFamily: string;
}

/** Word's WordArt gallery families: solid fill, fill + outline, hollow outline,
 * and one cell each for the shadow / reflection / glow effect rows. */
export const DOCUMENT_WORD_ART_STYLES: readonly DocumentWordArtStyle[] = [
  {
    id: "ink",
    textFill: "#0f172a",
    textOutline: "none",
    textOutlineWidthPt: 0,
    textEffect: "none",
    textFontFamily: "Arial",
  },
  {
    id: "teal",
    textFill: "#0a7d72",
    textOutline: "#0f172a",
    textOutlineWidthPt: 0.75,
    textEffect: "none",
    textFontFamily: "Arial",
  },
  {
    id: "crimson",
    textFill: "none",
    textOutline: "#dc2626",
    textOutlineWidthPt: 1,
    textEffect: "none",
    textFontFamily: "Georgia",
  },
  {
    id: "royal",
    textFill: "#2563eb",
    textOutline: "none",
    textOutlineWidthPt: 0,
    textEffect: "shadow",
    textFontFamily: "Arial",
  },
  {
    id: "gold",
    textFill: "#b45309",
    textOutline: "#78350f",
    textOutlineWidthPt: 0.5,
    textEffect: "reflection",
    textFontFamily: "Georgia",
  },
  {
    id: "marker",
    textFill: "#1e293b",
    textOutline: "none",
    textOutlineWidthPt: 0,
    textEffect: "glow",
    textFontFamily: "Arial",
  },
];

export const DOCUMENT_WORD_ART_DEFAULT_FONT_SIZE_PT = 36;
export const DOCUMENT_WORD_ART_DEFAULT_WIDTH_EMU = 360 * DOCUMENT_SHAPE_EMU_PER_PX;
export const DOCUMENT_WORD_ART_DEFAULT_HEIGHT_EMU = 72 * DOCUMENT_SHAPE_EMU_PER_PX;

/**
 * The shape a WordArt gallery cell inserts: a borderless, fill-less text box
 * (Word's own WordArt host) carrying the style's text paint. Built on
 * {@link documentShapeTextBoxDefaultAttrs} so it inherits every text-box
 * behaviour (linking, direction, wrap) and the Shape Format tab.
 */
export function documentShapeWordArtDefaultAttrs(
  style: DocumentWordArtStyle,
  text: string,
): DocumentShapeAttrs {
  return normalizeDocumentShapeAttrs({
    ...documentShapeTextBoxDefaultAttrs("horizontal"),
    widthEmu: DOCUMENT_WORD_ART_DEFAULT_WIDTH_EMU,
    heightEmu: DOCUMENT_WORD_ART_DEFAULT_HEIGHT_EMU,
    fill: "none",
    outline: "none",
    text,
    textFill: style.textFill,
    textOutline: style.textOutline,
    textOutlineWidthPt: style.textOutlineWidthPt,
    textEffect: style.textEffect,
    textFontFamily: style.textFontFamily,
    textFontSizePt: DOCUMENT_WORD_ART_DEFAULT_FONT_SIZE_PT,
    textBold: true,
    // WordArt keeps its own centred words and its drawn frame.
    textVerticalAlign: "middle",
    textHorizontalAlign: "center",
    textAutofit: undefined,
  });
}

/**
 * The CSS a gallery cell paints its sample letter with — the same three
 * declarations the shape node view derives from these attrs (colour,
 * `-webkit-text-stroke`, `text-shadow`), so the preview is the inserted look.
 */
export function documentWordArtPreviewStyle(
  style: DocumentWordArtStyle,
): Record<string, string> {
  const css: Record<string, string> = {
    color: style.textFill === "none" ? "transparent" : style.textFill,
    fontFamily: style.textFontFamily,
    fontWeight: "700",
  };
  if (style.textOutline !== "none" && style.textOutlineWidthPt > 0) {
    css.WebkitTextStroke = `${style.textOutlineWidthPt}px ${style.textOutline}`;
  }
  const shadow = documentShapeTextEffectShadowCss(style.textEffect);
  if (shadow) css.textShadow = shadow;
  return css;
}

/** One `text-shadow` per text-effect preset — the node view's rendering rule,
 * exported so the WordArt gallery previews with the identical declaration. */
export function documentShapeTextEffectShadowCss(
  effect: DocumentShapeEffectPreset,
): string {
  switch (effect) {
    case "shadow":
      return "2px 2px 2px rgb(15 23 42 / 45%)";
    case "reflection":
      return "0 6px 4px rgb(15 23 42 / 28%)";
    case "glow":
      return "0 0 6px rgb(59 130 246 / 80%)";
    case "softEdge":
      return "0 0 3px currentColor";
    case "bevel":
      return "1px 1px 0 #fff, -1px -1px 0 rgb(15 23 42 / 35%)";
    case "threeD":
      return "3px 3px 0 rgb(15 23 42 / 35%)";
    default:
      return "";
  }
}

export interface DocumentShapeLinkTarget {
  readonly pos: number;
  readonly attrs: Readonly<Record<string, unknown>>;
}

export interface DocumentShapeLinkPatch {
  readonly pos: number;
  readonly patch: Readonly<
    Pick<
      DocumentShapeAttrs,
      "linkedTextBoxStoryId" | "linkedTextBoxSequence"
    >
  >;
}

/** Link one text box to one empty successor in a single document transaction. */
export function documentShapeLinkedTextBoxPatches(
  source: DocumentShapeLinkTarget,
  target: DocumentShapeLinkTarget,
  requestedStoryId: number,
): readonly DocumentShapeLinkPatch[] | null {
  if (
    !Number.isSafeInteger(source.pos) ||
    !Number.isSafeInteger(target.pos) ||
    source.pos === target.pos
  ) {
    return null;
  }
  const sourceAttrs = normalizeDocumentShapeAttrs(source.attrs);
  const targetAttrs = normalizeDocumentShapeAttrs(target.attrs);
  if (
    sourceAttrs.groupMembers?.length ||
    targetAttrs.groupMembers?.length ||
    !sourceAttrs.isTextBox ||
    !targetAttrs.isTextBox ||
    targetAttrs.text.length > 0 ||
    typeof targetAttrs.linkedTextBoxStoryId === "number"
  ) {
    return null;
  }
  const existingStory = sourceAttrs.linkedTextBoxStoryId;
  const existingSequence = sourceAttrs.linkedTextBoxSequence;
  const storyId =
    typeof existingStory === "number"
      ? existingStory
      : documentShapeLinkedTextBoxCoordinate(requestedStoryId, 1);
  const sourceSequence =
    typeof existingSequence === "number" ? existingSequence : 0;
  if (storyId === null || sourceSequence >= 0xffff) return null;
  return [
    {
      pos: source.pos,
      patch: {
        linkedTextBoxStoryId: storyId,
        linkedTextBoxSequence: sourceSequence,
      },
    },
    {
      pos: target.pos,
      patch: {
        linkedTextBoxStoryId: storyId,
        linkedTextBoxSequence: sourceSequence + 1,
      },
    },
  ];
}

/** Resolve an exact degree entry through the persisted 0.1-degree authority. */
export function documentShapeExactRotationPatch(
  value: DocumentShapeAttrsInput,
  rawDegrees: string | number,
): Readonly<Pick<DocumentShapeAttrs, "rotation">> | null {
  const source =
    typeof rawDegrees === "string" ? rawDegrees.trim() : rawDegrees;
  if (source === "") return null;
  const requestedDegrees = Number(source);
  if (!Number.isFinite(requestedDegrees)) return null;

  const current = normalizeDocumentShapeAttrs(value);
  const next = normalizeDocumentShapeAttrs({
    ...current,
    rotation: requestedDegrees,
  });
  return next.rotation === current.rotation
    ? null
    : { rotation: next.rotation };
}

/**
 * Resolve one exact edge edit from the Layout ribbon without creating empty
 * history. The user-facing control is expressed in CSS px because the editor
 * canvas uses px, while this helper immediately converts to the persisted
 * DrawingML EMU authority used by import, both exporters, and the NodeView.
 */
function halfUpRatio(
  value: number,
  numerator: number,
  denominator: number,
): number | null {
  if (
    !Number.isSafeInteger(value) ||
    !Number.isSafeInteger(numerator) ||
    !Number.isSafeInteger(denominator) ||
    value <= 0 ||
    numerator <= 0 ||
    denominator <= 0
  ) {
    return null;
  }
  const product = value * numerator;
  if (!Number.isSafeInteger(product)) return null;
  return Math.floor((product + Math.floor(denominator / 2)) / denominator);
}

function resolveDocumentShapeSizeEmuPatch(
  current: DocumentShapeAttrs,
  axis: DocumentShapeSizeAxis,
  requestedEmu: number,
  preserveAspectRatio: boolean,
): Readonly<Pick<DocumentShapeAttrs, "widthEmu" | "heightEmu">> | null {
  if (!Number.isSafeInteger(requestedEmu)) return null;

  let widthEmu = axis === "width" ? requestedEmu : current.widthEmu;
  let heightEmu = axis === "height" ? requestedEmu : current.heightEmu;
  if (preserveAspectRatio) {
    if (axis === "width") {
      const derived = halfUpRatio(
        requestedEmu,
        current.heightEmu,
        current.widthEmu,
      );
      if (derived === null) return null;
      heightEmu = derived;
    } else {
      const derived = halfUpRatio(
        requestedEmu,
        current.widthEmu,
        current.heightEmu,
      );
      if (derived === null) return null;
      widthEmu = derived;
    }
    // A locked edit is atomic. If either requested or derived edge is outside
    // the native shape bounds, refuse the whole transaction rather than clamp
    // one edge and silently distort the ratio.
    if (
      widthEmu < documentShapeMinEdgeEmu(current.preset) ||
      widthEmu > DOCUMENT_SHAPE_MAX_EDGE_EMU ||
      heightEmu < documentShapeMinEdgeEmu(current.preset) ||
      heightEmu > DOCUMENT_SHAPE_MAX_EDGE_EMU
    ) {
      return null;
    }
  } else {
    const normalized = normalizeDocumentShapeAttrs({
      ...current,
      widthEmu,
      heightEmu,
    });
    widthEmu = normalized.widthEmu;
    heightEmu = normalized.heightEmu;
  }

  if (widthEmu === current.widthEmu && heightEmu === current.heightEmu) {
    return null;
  }
  return { widthEmu, heightEmu };
}

/**
 * Resolve one persisted-EMU edge edit. NodeView resize uses this directly so
 * pointer and ribbon sizing share the exact half-up aspect authority.
 */
export function documentShapeSizeEmuPatch(
  value: DocumentShapeAttrsInput,
  axis: DocumentShapeSizeAxis,
  requestedEmu: number,
  preserveAspectRatio?: boolean,
): Readonly<Pick<DocumentShapeAttrs, "widthEmu" | "heightEmu">> | null {
  const current = normalizeDocumentShapeAttrs(value);
  return resolveDocumentShapeSizeEmuPatch(
    current,
    axis,
    requestedEmu,
    preserveAspectRatio ?? (current.lockAspectRatio === true),
  );
}

export function documentShapeSizePatch(
  value: DocumentShapeAttrsInput,
  axis: DocumentShapeSizeAxis,
  rawPx: string | number,
): Readonly<Pick<DocumentShapeAttrs, "widthEmu" | "heightEmu">> | null {
  const source = typeof rawPx === "string" ? rawPx.trim() : rawPx;
  if (source === "") return null;
  const requestedPx = Number(source);
  if (!Number.isFinite(requestedPx)) return null;

  return documentShapeSizeEmuPatch(
    value,
    axis,
    documentShapePxToEmu(requestedPx),
  );
}

export function documentShapeEmuToPx(value: number): number {
  return value / DOCUMENT_SHAPE_EMU_PER_PX;
}

export function documentShapePxToEmu(value: number): number {
  return Math.round(value * DOCUMENT_SHAPE_EMU_PER_PX);
}
