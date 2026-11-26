"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One block of the slide's editing canvas (TRIAGE 155).
 *
 * The canvas drew its blocks in one inline map of the window, each with a
 * bag of window values rebuilt on every render, so any render of the window
 * -- a change of selection included -- drew every block again: on a slide of
 * 2,351 objects a Tab, a Select All, a marquee or a Selection Pane row was a
 * 0.8-1.2 s task. A block now takes
 *
 * - its own `block` (the same object while the slide does not change it),
 * - `view`: what the selection, the gestures and the previews make of THIS
 *   block (`presentationCanvasBlockView`), plain values compared one by one,
 * - `shared`: the window's values that no selection changes, one object for
 *   every block (its functions are stable: `useStableHandlers`),
 * - `chrome`: the selection itself, for the block that wears the selection's
 *   frame and grips (null for every other block),
 * - `kindCtl`: a chart's, a table's or a media object's own window bindings,
 *
 * and is drawn again only when one of them changes.
 */

import { memo, useMemo } from "react";
import type {
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  PointerEvent as ReactPointerEvent,
  RefObject,
} from "react";
import { shallowEqualRecords } from "@mrnote/core/web/lib/shallow-equal.ts";
import type { PresentationBlock, PresentationImageBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { PresentationCanvasShapeBlockCtl } from "./PresentationCanvasShapeBlock";
import type { PresentationCanvasTextBlockCtl } from "./PresentationCanvasTextBlock";
import type { PresentationCanvasImageBlockCtl } from "./PresentationCanvasImageBlock";
import type { PresentationCanvasChartBlockCtl } from "./PresentationCanvasChartBlock";
import type { PresentationCanvasTableBlockCtl } from "./PresentationCanvasTableBlock";
import type { PresentationCanvasEmbedBlockCtl } from "./PresentationCanvasEmbedBlock";
import type { PresentationBlockHandlesCtl } from "./PresentationBlockHandles";
import type { PresentationBlockAutofitSmartTagCtl } from "./PresentationBlockAutofitSmartTag";
import type { PresentationSmartArtGraphicHitLayerCtl } from "./PresentationSmartArtGraphicHitArea";
import type { PresentationFormatPainterState } from "@mrnote/core/web/lib/presentation-format-painter.ts";
import type { TimelineStep } from "@mrnote/core/web/lib/presentation-anim-timeline.ts";
import type { PresentationPictureSizeCommandAdapter } from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import { canChangePresentationChartFrame } from "@mrnote/core/web/lib/presentation-chart-frame-operations.ts";
import { isNativeSmartArtMember, readNativeSmartArtView } from "@mrnote/core/web/lib/presentation-native-smartart.ts";
import { officeDrawingSelectionStyles, officeDrawingVisibleResizeHandles } from "@mrnote/core/web/components/office-ribbon/OfficeDrawingSelection.tsx";
import type { CSSProperties } from "react";
import { createPortal } from "react-dom";
import { ChevronDown, ChevronUp } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { presentationPressDragEnabled } from "../../../../lib/presentation-flags";
import { pictureObjectTokenSha, pictureObjectUrlSha } from "@mrnote/core/web/lib/page-picture-objects.ts";
import { applyPresentationTableStyleSelection } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationTableStyleGallery.tsx";
import { applyPresentationPictureFrameStyle } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationPictureStyleGallery.tsx";
import { isPresentationConnectorShapeKind, normalizeShapeKind, presentationShapeBlendCss } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { logicalCellPositions, tableLogicalWidth } from "@mrnote/core/web/lib/presentation-table-merge.ts";
import { setAxisSize } from "@mrnote/core/web/lib/presentation-table-size.ts";
import { readSmartArtBlockMarker, smartArtPictureReplaceTargetId } from "@mrnote/core/web/lib/presentation-smartart.ts";
import { type SmartArtScene3d } from "@mrnote/core/web/lib/office-smartart/styles-model.ts";
import { smartArtCameraCss, smartArtCameraWithOwnTransform } from "@mrnote/core/web/lib/office-smartart/canvas-3d.ts";
import { smartArtEditIn2dKey, smartArtEditingIn2d, smartArtMemberHasScene } from "@mrnote/core/web/lib/office-smartart/edit-in-2d.ts";
import { blockGeometry } from "@mrnote/core/web/lib/presentation-conveniences.ts";
import { legacyLineSlopeFlipsHorizontally } from "@mrnote/core/web/lib/presentation-connectors.ts";
import { exactNativeGroupMembers, geometryNumber, hasLayoutGeometry, isImportedPptxBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { presentationTextUsesPoints } from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import { resolvePresentationTextStyle } from "../../../../lib/presentation-text-style-resolve";
import { normalizeRotationDegrees, type BlockGeometry } from "@mrnote/core/web/lib/presentation-drag.ts";
import { isTextEditTarget } from "../../../../lib/presentation-a11y";
import { withPresentationBlackWhiteStyle } from "@mrnote/core/web/lib/presentation/presentation-black-white-mode.ts";
import { isLayoutGraphicOnSlide } from "@mrnote/core/web/lib/presentation-master.ts";
import { presentationBlockOverFooterFieldsZIndex, withFooterFieldStacking } from "@mrnote/core/web/lib/presentation-paint-order.ts";
import { isCompatMouseFromTouch, isTouchLikePointer } from "@mrnote/core/web/lib/pointer-capability.ts";
import { officePressIntoTextHost } from "@mrnote/core/web/lib/office-off-line-text-press.ts";
import { INK_LASSO_MEMBER_ATTRIBUTE } from "@mrnote/core/web/lib/office-ink-lasso-selection.ts";
import { canManipulateBlock, isBlockKeptReadOnly, isBlockLocked, isFrameLockedByFile, isTextEditLocked } from "@mrnote/core/web/lib/presentation-text-case.ts";
import { blockGroupId } from "@mrnote/core/web/lib/presentation-crop-group.ts";
import { presentationBlockFrameOverflow } from "../../../../lib/presentation/block-frame-overflow";
import type { PresentationTableBlock, PresentationTextBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { AUTOFIT_EMU_PER_PX, autofitSmartTagVisibleForBlock, importedPptxTextHasPreview, textBlockDrawsInsets, textBlockInsetsCss, textBlockStyleCss, type PresentationFontSizeUnit } from "../../../../lib/presentation/text-layout-css";
import { presentationDoubleClickWord, presentationPreviewOffsetAtPoint, presentationPreviewPlainText, setPresentationPreviewSelection } from "@mrnote/core/web/lib/presentation/text-selection.ts";
import { blockOrderButtonStyle } from "../../../../lib/presentation/window-chrome-styles";
import { blockClaimsImportedGroup, blockGeometryStyle, nativeGroupMoveMembers, presentationInkBlockPlacement, presentationSmartArtGroupBox, presentationSmartArtLabelShape, presentationSmartArtShapesLockedIn3d } from "../../../../lib/presentation/selection-geometry";
import { tableWithFrameAlignedGrid, tableWithScaledGestureGeometry } from "../../../../lib/presentation/table-canvas-layout";
import { PresentationCanvasTextBlock } from "./PresentationCanvasTextBlock";
import { PresentationCanvasEmbedBlock } from "./PresentationCanvasEmbedBlock";
import { PresentationCanvasShapeBlock } from "./PresentationCanvasShapeBlock";
import { PresentationCanvasTableBlock } from "./PresentationCanvasTableBlock";
import { PresentationCanvasChartBlock } from "./PresentationCanvasChartBlock";
import { PresentationCanvasImageBlock } from "./PresentationCanvasImageBlock";
import { PresentationDegradedObjectOverlay } from "./PresentationDegradedObject";
import { degradedObjectOf } from "@mrnote/core/web/lib/presentation-object-degraded.ts";
import { PresentationBlockHandles } from "./PresentationBlockHandles";
import { PresentationBlockAutofitSmartTag } from "./PresentationBlockAutofitSmartTag";
import { PresentationSmartArtGraphicHitLayer, presentationShapeGeometryHitClassName, presentationShapeIsNativeTextBox } from "./PresentationSmartArtGraphicHitArea";
import { PresentationLiveBlockFrame, presentationLiveFrameOf } from "./PresentationLiveBlockFrame";

type ShapeCtl = PresentationCanvasShapeBlockCtl;
type TextCtl = PresentationCanvasTextBlockCtl;
type ImageCtl = PresentationCanvasImageBlockCtl;

/** The window's functions a block calls, each the same on every render. */
export type PresentationCanvasBlockHandlers = Pick<
      ShapeCtl,
      | "setSelectedBlockId"
      | "setEditingShapeTextId"
      | "setShapeTextBlankDraft"
      | "editSmartArtShapeText"
      | "handleShapeTextChange"
      | "nudgeCropHandle"
      | "beginCropGesture"
      | "recordShapeTextSelection"
      | "commitShapeRichText"
      | "recordShapeTextRange"
      | "takeShapeRichSelection"
      | "placeShapeTextCaret"
      | "toggleShapeWordsRunStyle"
      | "copyShapeText"
      | "pasteShapeTextRuns"
      | "undo"
      | "redo"
    > &
    Pick<
      TextCtl,
      | "deleteBlock"
      | "captureTextEditSelection"
      | "consumeTextEditSelection"
      | "updateTextBlock"
      | "handleTextBlockBodyChange"
      | "captureImportedRichTextSelection"
      | "followImportedRichTextSelection"
      | "restoreImportedPreviewSelectionAfterCommit"
      | "commitImportedPreviewText"
      | "replaceImportedPreviewSelection"
      | "clearLiveRibbonList"
      | "pasteIntoImportedPreview"
      | "copyFromImportedPreview"
      | "adjustParagraphCaretLevel"
      | "applyTextBlockRunToggle"
      | "handleFindRevealedKey"
      | "fillContentPlaceholder"
    > &
    Pick<ImageCtl, "replaceBlock" | "selectBlockFromClick" | "beginReplaceImage"> &
    Pick<PresentationBlockHandlesCtl, "beginBlockGesture" | "beginRotateGesture"> &
    Pick<PresentationSmartArtGraphicHitLayerCtl, "setSmartArtGraphicOnly" | "setBlockContextMenu"> &
    Pick<PresentationBlockAutofitSmartTagCtl, "setAutofitSmartMenuBlockId" | "setTextBlockAutofit"> & {
  keepTextSelectionSnapshot: TextCtl["setTextSelectionSnapshot"];
  noteBlockTouchPress: (event: ReactPointerEvent<HTMLElement>, blockId: string) => void;
  setGroupMemberFocusId: (blockId: string | null) => void;
  setSelectedBlockIds: (ids: ReadonlySet<string>) => void;
  clearTextSelectionAuthority: () => void;
  paintBlockFormat: (blockId: string) => void;
  applyAnimationToBlock: (blockId: string) => void;
  openBlockContextMenu: (event: ReactMouseEvent<HTMLElement>, blockId: string) => void;
  handleBlockKeyDown: (event: ReactKeyboardEvent<HTMLDivElement>, blockId: string) => void;
  moveBlock: (blockId: string, direction: "up" | "down") => void;
};

/**
 * The window's values no selection changes: one object for every block of
 * the canvas, rebuilt only when one of them changes (a slide edited, the
 * zoom, the theme, the fonts...). Its functions read refs, or are the
 * window's stable handlers: none of them stands for a value that changes.
 */
export interface PresentationCanvasBlockShared {
  t: ShapeCtl["t"];
  activeSlide: SlideModel;
  activeSlideBounds: { width: number; height: number };
  /** The slide's drawn blocks (the SmartArt hit areas are found in them). */
  activeSlideVisibleBlocks: readonly PresentationBlock[];
  theme: TextCtl["theme"];
  masters: TextCtl["masters"];
  fontDisplay: ShapeCtl["fontDisplay"];
  previewMode: boolean;
  /** No SmartArt hit area: a preview slide, or a drawing tool on the slide. */
  hitLayerInert: boolean;
  viewColorMode: Parameters<typeof withPresentationBlackWhiteStyle>[0];
  liveTransformSurface: string;
  selectionChromeHost: HTMLDivElement | null;
  slideScreenScale: number;
  deckThemeGridColors: ImageCtl["deckThemeGridColors"];
  showProofingErrors: boolean;
  editingHostHandover: ShapeCtl["editingHostHandover"];
  readOnly: boolean;
  readOnlyReason: string | null | undefined;
  notebookId: string | null | undefined;
  pictureReplaceBusy: boolean;
  /**
   * The Edit in 2-D state's version (lib/office-smartart/edit-in-2d): a
   * graphic going flat or back draws the canvas again.
   */
  smartArtEditIn2dVersion: number;
  // Read while drawing: each reads refs, so it is the same on every render.
  heldBlockWords: ShapeCtl["heldBlockWords"];
  pendingImportedPreviewCaret: TextCtl["pendingImportedPreviewCaret"];
  runStyleAtBodySelection: TextCtl["runStyleAtBodySelection"];
  blockAriaName: TextCtl["blockAriaName"];
  refs: {
    animationPainterRef: RefObject<TimelineStep[] | null>;
    autofitSmartMenuTriggerRef: PresentationBlockAutofitSmartTagCtl["autofitSmartMenuTriggerRef"];
    composingImportedPreviewRef: TextCtl["composingImportedPreviewRef"];
    composingTextRef: TextCtl["composingTextRef"];
    deckContentGenerationRef: TextCtl["deckContentGenerationRef"];
    findRevealRef: TextCtl["findRevealRef"];
    formatPainterRef: RefObject<PresentationFormatPainterState | null>;
    gestureEndClickRef: RefObject<{ blockId: string; at: number } | null>;
    groupMemberPressRef: RefObject<string | null>;
    pictureSizeCommandRef: RefObject<PresentationPictureSizeCommandAdapter<PresentationImageBlock, SlideModel>>;
    selectedBlockIdRef: RefObject<string | null>;
    selectedBlockIdsRef: RefObject<ReadonlySet<string>>;
    shapeLastPressRef: RefObject<{ blockId: string; at: number; x: number; y: number } | null>;
    shapeWordPressRef: RefObject<{ blockId: string; at: number } | null>;
    tabKeyAtRef: RefObject<number>;
    textSelectionIntentRef: RefObject<number>;
    textSelectionRef: TextCtl["textSelectionRef"];
    textTapClickRef: RefObject<string | null>;
    touchPressRef: RefObject<{ blockId: string; wasSelected: boolean } | null>;
    wordsPressBlockRef: RefObject<string | null>;
  };
  handlers: PresentationCanvasBlockHandlers;
}

/**
 * The selection as the block wearing its frame and grips reads it: the
 * primary block's (or, while a SmartArt label is typed in, its shape's).
 * Rebuilt on every render of the window and given to that block alone.
 */
export interface PresentationCanvasBlockChrome {
  smartArtGraphicOnly: boolean;
  selectedBlockIds: ReadonlySet<string>;
  selectedGeometryBlocks: readonly PresentationBlock[];
  selectedGroupView: {
    ids: ReadonlySet<string>;
    primaryId: string;
    box: { left: number; top: number; width: number; height: number };
    memberSelected: boolean;
    rotation: number;
  } | null;
  /** The selected picture's size may be changed (Picture Format ▸ Size). */
  pictureCanResize: () => boolean;
  selectedChartElementArea: unknown;
  smartArtFrameRotateGrip: PresentationBlockHandlesCtl["smartArtFrameRotateGrip"];
  presentationShapeDrawingOverlay: PresentationBlockHandlesCtl["presentationShapeDrawingOverlay"];
}

/** A text body that draws no insets of its own. */
const NO_INSETS: CSSProperties = {};

const NO_CHROME: PresentationCanvasBlockChrome = {
  smartArtGraphicOnly: false,
  selectedBlockIds: new Set<string>(),
  selectedGeometryBlocks: [],
  selectedGroupView: null,
  pictureCanResize: () => false,
  selectedChartElementArea: null,
  smartArtFrameRotateGrip: null,
  presentationShapeDrawingOverlay: () => null,
};

/** A chart's, a table's or a media object's own bindings (rebuilt every render). */
export type PresentationCanvasBlockKindCtl =
  | {
      kind: "chart";
      ctl: Omit<PresentationCanvasChartBlockCtl, "block" | "isSelected" | "importedPptxBlock" | "previewGeom">;
      selectedGroupView: PresentationCanvasBlockChrome["selectedGroupView"];
    }
  | {
      kind: "table";
      ctl: Omit<
        PresentationCanvasTableBlockCtl,
        "block" | "showHandles" | "displayedTableBlock" | "tableCellPositions" | "tableColumnCount"
      >;
      tableDrawBoundaryPreview: PresentationTableBlock | null;
      tableStylePreview: {
        blockId: string;
        selection: Parameters<typeof applyPresentationTableStyleSelection>[1];
      } | null;
    }
  | { kind: "embed"; ctl: Omit<PresentationCanvasEmbedBlockCtl, "block"> };

/** What the selection, the gestures and the previews make of one block. */
export interface PresentationCanvasBlockView {
  /** The selection's primary block. */
  isSelected: boolean;
  /** In the selected set. */
  inSelection: boolean;
  /** The window's single-selection pointer (`selectedBlockId`). */
  legacySelected: boolean;
  /**
   * The object chrome (frame, handles) of the selection: the primary
   * block's own, except a SmartArt label being typed in, whose chrome is
   * drawn on its shape (`smartArtLabelShapeId`).
   */
  ownsSelectionChrome: boolean;
  /**
   * A SmartArt diagram's other members sit inside the diagram border
   * instead of each wearing a frame, as PowerPoint frames only the member
   * that was clicked.
   */
  framedBySmartArtBorder: boolean;
  /**
   * A member of a whole selected group wears no frame of its own: the
   * group's one frame stands for it (`selectedGroupView`), and the
   * primary's chrome is drawn on the group's box -- except the shape being
   * typed in, which keeps its own text-edit frame.
   */
  groupChromeOnBox: boolean;
  framedByGroup: boolean;
  /** A press on it selects it inside the group selected whole. */
  groupMemberPressable: boolean;
  /** A member selected inside its group: it wears its own grips. */
  memberAloneChrome: boolean;
  /** Its words are being typed in (a text box's, or a shape's label). */
  textEditingHere: boolean;
  /** A shape whose label editor is open (`editingShapeTextId`). */
  editingShapeText: boolean;
  /** It holds the canvas's one roving tab stop. */
  isRoving: boolean;
  /** A stroke Lasso Select took. */
  inkLassoMember: boolean;
  /** An armed Format Painter: its source, or a target. */
  formatPainterRole: "source" | "target" | undefined;
  /** The selection's primary, for the SmartArt graphic hit area it carries (else null). */
  smartArtPrimaryId: string | null;
  /** A picture of a selected SmartArt graphic: the member Replace targets. */
  pictureSmartArtMemberId: string | null;
  /** Selected, and a metafile the server could not draw. */
  pictureNotDrawn: boolean;
  /** A Copilot preview shows other objects only. */
  copilotHidden: boolean;
  /** A prepared SmartArt conversion: the file's drawing, or the conversion under it. */
  preparedRole: "original" | "converted" | undefined;
  cropEditing: boolean;
  cropPreview: ShapeCtl["cropPreview"];
  /** This desk's own gesture: the live rect, and its mirror while a handle passes the pinned edge. */
  dragGeom: BlockGeometry | undefined;
  dragFlip: { flipH: boolean; flipV: boolean } | undefined;
  /** A turn's live frame of this member. */
  rotateMember: { left: number; top: number; width: number; height: number; rotation: number } | undefined;
  /** The rotate handle's live angle on the block it was pressed on. */
  rotateRotation: number | undefined;
  shapeTextBlankDraft: ShapeCtl["shapeTextBlankDraft"];
  shapeAdjustPreview: ShapeCtl["shapeAdjustPreview"];
  /** A Picture Styles gallery item hovered for it. */
  pictureStyle: Parameters<typeof applyPresentationPictureFrameStyle>[1] | undefined;
  /** The equation being written in this box, drawn in its line. */
  equationPreviewBlock: TextCtl["previewBlock"];
  /** Its AutoFit smart tag's menu is open. */
  autofitMenuOpen: boolean;
}

/** The window's selection and previews, read once per render for every block. */
export interface PresentationCanvasSelectionState {
  slideBlocks: readonly PresentationBlock[];
  primaryId: string | null;
  selectedBlockId: string | null;
  selectedBlockIds: ReadonlySet<string>;
  smartArtLabelShapeId: string | null;
  /** The members of the selected SmartArt diagram while its frame is drawn. */
  smartArtBorderMemberIds: ReadonlySet<string> | null;
  selectedGroupView: PresentationCanvasBlockChrome["selectedGroupView"];
  groupMemberAloneId: string | null;
  textEditingBlockId: string | null;
  editingShapeTextId: string | null;
  rovingBlockId: string | null;
  inkLassoIds: ReadonlySet<string> | null;
  /** Null: no painter armed. */
  formatPainterSourceId: string | null;
  formatPainterArmed: boolean;
  /** The SmartArt graphic hit areas of the drawn blocks (empty when inert). */
  smartArtHitAreas: ReadonlyMap<string, unknown>;
  picturesNotDrawn: ReadonlySet<string>;
  copilotPreviewBlockIds: readonly string[] | null | undefined;
  copilotPreviewBlockId: string | null | undefined;
  preparedOriginals: readonly PresentationBlock[];
  preparedHidden: ReadonlySet<string> | null | undefined;
  cropEditingId: string | null;
  cropPreview: ShapeCtl["cropPreview"];
  dragPreview: {
    blockId: string;
    geom: BlockGeometry;
    members?: ReadonlyMap<string, BlockGeometry>;
    flip?: { flipH: boolean; flipV: boolean };
  } | null;
  rotatePreview: {
    blockId: string;
    rotation: number;
    members?: ReadonlyMap<string, { left: number; top: number; width: number; height: number; rotation: number }>;
  } | null;
  shapeTextBlankDraft: ShapeCtl["shapeTextBlankDraft"];
  shapeAdjustPreview: ShapeCtl["shapeAdjustPreview"];
  pictureStylePreview: { blockId: string; style: Parameters<typeof applyPresentationPictureFrameStyle>[1] } | null;
  equationSlot: { blockId: string; block: NonNullable<TextCtl["previewBlock"]> } | null | undefined;
  autofitSmartMenuBlockId: string | null;
}

/** One block's `view` (allocated per block; compared field by field). */
export function presentationCanvasBlockView(
  block: PresentationBlock,
  selection: PresentationCanvasSelectionState,
): PresentationCanvasBlockView {
  const id = block.id;
  const isSelected = selection.primaryId === id;
  const ownsSelectionChrome =
    selection.smartArtLabelShapeId !== null ? selection.smartArtLabelShapeId === id : isSelected;
  const group = selection.selectedGroupView;
  const drag = selection.dragPreview;
  const rotate = selection.rotatePreview;
  const crop = selection.cropPreview;
  const blank = selection.shapeTextBlankDraft;
  const adjust = selection.shapeAdjustPreview;
  const inSelection = selection.selectedBlockIds.has(id);
  return {
    isSelected,
    inSelection,
    legacySelected: selection.selectedBlockId === id,
    ownsSelectionChrome,
    framedBySmartArtBorder:
      !ownsSelectionChrome && Boolean(selection.smartArtBorderMemberIds?.has(id)),
    groupChromeOnBox:
      group !== null && ownsSelectionChrome && group.primaryId === id && !group.memberSelected,
    framedByGroup:
      group !== null && group.ids.has(id) && !(group.memberSelected && group.primaryId === id),
    groupMemberPressable: group !== null && !group.memberSelected && group.ids.has(id),
    memberAloneChrome: selection.groupMemberAloneId === id,
    textEditingHere: selection.textEditingBlockId === id || selection.editingShapeTextId === id,
    editingShapeText: selection.editingShapeTextId === id,
    isRoving: selection.rovingBlockId === id,
    inkLassoMember: Boolean(selection.inkLassoIds?.has(id)),
    formatPainterRole: selection.formatPainterArmed
      ? selection.formatPainterSourceId === id
        ? "source"
        : "target"
      : undefined,
    smartArtPrimaryId: selection.smartArtHitAreas.has(id) ? selection.primaryId : null,
    pictureSmartArtMemberId:
      block.type === "image"
        ? smartArtPictureReplaceTargetId(selection.slideBlocks, selection.selectedBlockIds, id)
        : null,
    pictureNotDrawn:
      block.type === "image" &&
      inSelection &&
      selection.picturesNotDrawn.has(pictureObjectTokenSha(block.src) ?? pictureObjectUrlSha(block.src) ?? ""),
    copilotHidden: selection.copilotPreviewBlockIds
      ? !selection.copilotPreviewBlockIds.includes(id)
      : Boolean(selection.copilotPreviewBlockId && selection.copilotPreviewBlockId !== id),
    preparedRole: selection.preparedOriginals.includes(block)
      ? "original"
      : selection.preparedHidden?.has(id)
        ? "converted"
        : undefined,
    cropEditing: selection.cropEditingId === id,
    cropPreview: crop && crop.blockId === id ? crop : null,
    dragGeom: drag ? (drag.blockId === id ? drag.geom : drag.members?.get(id)) : undefined,
    dragFlip: drag && drag.blockId === id ? drag.flip : undefined,
    rotateMember: rotate?.members?.get(id),
    rotateRotation: rotate && rotate.blockId === id ? rotate.rotation : undefined,
    shapeTextBlankDraft: blank && blank.blockId === id ? blank : null,
    shapeAdjustPreview: adjust && adjust.blockId === id ? adjust : null,
    pictureStyle:
      selection.pictureStylePreview?.blockId === id ? selection.pictureStylePreview.style : undefined,
    equationPreviewBlock:
      selection.equationSlot?.blockId === id ? selection.equationSlot.block : undefined,
    autofitMenuOpen: selection.autofitSmartMenuBlockId === id,
  };
}

/** Whether the block also takes the selection itself (`chrome`). */
export function presentationCanvasBlockTakesChrome(view: PresentationCanvasBlockView): boolean {
  return view.isSelected || view.ownsSelectionChrome;
}

export interface PresentationCanvasBlockSlotProps {
  block: PresentationBlock;
  blockIndex: number;
  view: PresentationCanvasBlockView;
  shared: PresentationCanvasBlockShared;
  chrome: PresentationCanvasBlockChrome | null;
  kindCtl: PresentationCanvasBlockKindCtl | null;
}

/**
 * The slot is drawn again when any prop changes: `view` is compared field by
 * field, every other prop by identity. Every prop and every field counts, so
 * one added later counts too.
 */
export function presentationCanvasBlockSlotPropsEqual(
  previous: Readonly<PresentationCanvasBlockSlotProps>,
  next: Readonly<PresentationCanvasBlockSlotProps>,
): boolean {
  const prev = previous as unknown as Record<string, unknown>;
  const nxt = next as unknown as Record<string, unknown>;
  const keys = Object.keys(prev);
  if (keys.length !== Object.keys(nxt).length) return false;
  for (const key of keys) {
    if (!Object.prototype.hasOwnProperty.call(nxt, key)) return false;
    if (key === "view") {
      if (
        !shallowEqualRecords(
          prev.view as Record<string, unknown>,
          nxt.view as Record<string, unknown>,
        )
      ) {
        return false;
      }
    } else if (!Object.is(prev[key], nxt[key])) {
      return false;
    }
  }
  return true;
}

function PresentationCanvasBlockSlotBody({
  block,
  blockIndex,
  view,
  shared,
  chrome,
  kindCtl,
}: PresentationCanvasBlockSlotProps) {
  const {
    t,
    activeSlide,
    activeSlideBounds,
    activeSlideVisibleBlocks,
    theme,
    masters,
    fontDisplay,
    previewMode,
    viewColorMode,
    liveTransformSurface,
    selectionChromeHost,
    slideScreenScale,
    deckThemeGridColors,
    showProofingErrors,
    editingHostHandover,
    heldBlockWords,
    pendingImportedPreviewCaret,
    runStyleAtBodySelection,
    blockAriaName,
  } = shared;
  const {
    animationPainterRef,
    autofitSmartMenuTriggerRef,
    composingImportedPreviewRef,
    composingTextRef,
    deckContentGenerationRef,
    findRevealRef,
    formatPainterRef,
    gestureEndClickRef,
    groupMemberPressRef,
    pictureSizeCommandRef,
    selectedBlockIdRef,
    selectedBlockIdsRef,
    shapeLastPressRef,
    shapeWordPressRef,
    tabKeyAtRef,
    textSelectionIntentRef,
    textSelectionRef,
    textTapClickRef,
    touchPressRef,
    wordsPressBlockRef,
  } = shared.refs;
  const {
    adjustParagraphCaretLevel,
    applyAnimationToBlock,
    applyTextBlockRunToggle,
    beginBlockGesture,
    beginCropGesture,
    beginReplaceImage,
    beginRotateGesture,
    captureImportedRichTextSelection,
    captureTextEditSelection,
    clearLiveRibbonList,
    clearTextSelectionAuthority,
    commitImportedPreviewText,
    commitShapeRichText,
    consumeTextEditSelection,
    copyFromImportedPreview,
    copyShapeText,
    deleteBlock,
    editSmartArtShapeText,
    fillContentPlaceholder,
    followImportedRichTextSelection,
    handleBlockKeyDown,
    handleFindRevealedKey,
    handleShapeTextChange,
    handleTextBlockBodyChange,
    keepTextSelectionSnapshot,
    moveBlock,
    noteBlockTouchPress,
    nudgeCropHandle,
    openBlockContextMenu,
    paintBlockFormat,
    pasteIntoImportedPreview,
    pasteShapeTextRuns,
    placeShapeTextCaret,
    recordShapeTextRange,
    recordShapeTextSelection,
    redo,
    replaceBlock,
    replaceImportedPreviewSelection,
    restoreImportedPreviewSelectionAfterCommit,
    selectBlockFromClick,
    setAutofitSmartMenuBlockId,
    setBlockContextMenu,
    setEditingShapeTextId,
    setGroupMemberFocusId,
    setSelectedBlockId,
    setSelectedBlockIds,
    setShapeTextBlankDraft,
    setSmartArtGraphicOnly,
    setTextBlockAutofit,
    takeShapeRichSelection,
    toggleShapeWordsRunStyle,
    undo,
    updateTextBlock,
  } = shared.handlers;
  const sel = chrome ?? NO_CHROME;
  const {
    selectedGroupView,
    selectedGeometryBlocks,
    selectedChartElementArea,
    smartArtFrameRotateGrip,
    presentationShapeDrawingOverlay,
  } = sel;
  const chartKind = kindCtl?.kind === "chart" ? kindCtl : null;
  const tableKind = kindCtl?.kind === "table" ? kindCtl : null;
  const embedKind = kindCtl?.kind === "embed" ? kindCtl : null;
  const tableAxisPreview = tableKind?.ctl.tableAxisPreview ?? null;
  const isGeometryBlock = hasLayoutGeometry(block);
  const {
    isSelected,
    ownsSelectionChrome,
    framedBySmartArtBorder,
    groupChromeOnBox,
    framedByGroup,
    memberAloneChrome,
  } = view;
  const isMultiSelected = isGeometryBlock && view.inSelection;
  // A SmartArt label lets a press through to its shape until its
  // words are being typed in: one click selects the shape, a
  // double click (or Enter / F2 / a key typed on the shape) puts
  // the caret in its words.
  const smartArtLabelYieldsToShape =
    !isSelected &&
    block.type === "text" &&
    presentationSmartArtLabelShape(activeSlide.blocks, block) !== null;
  const importedPptxBlock = isImportedPptxBlock(block);
  // A SmartArt member is drawn by its diagram, never by the
  // authoring plate the canvas puts under a free text box: the
  // members overlap (rings, bands, labels over nodes), so the
  // translucent plate of each block would band across the
  // others. After a reopen the members carry a native source and
  // take the transparent path anyway; a freshly inserted diagram
  // must look the same.
  const smartArtMember =
    typeof (block as { smartArtRole?: unknown }).smartArtRole === "string";
  const wearsDrawingSelection =
    !framedBySmartArtBorder &&
    (isGeometryBlock ||
      importedPptxBlock ||
      block.type === "image" ||
      smartArtMember ||
      block.type === "shape" ||
      block.type === "chart");
  // PowerPoint's text-edit mode on a text box, a shape or a SmartArt
  // shape: the same outline and handles, the outline dashed
  // (lib/office-text-edit-mode). A table keeps its cell selection.
  const textEditing =
    (block.type === "text" || block.type === "shape") && view.textEditingHere;
  const fontSizeUnit: PresentationFontSizeUnit = presentationTextUsesPoints(block)
    ? "point"
    : "design-px";
  const importedPptxText =
    importedPptxBlock && block.type === "text";
  // Positioned rich text uses one visible body regardless of origin;
  // the legacy textarea remains only as a selection/input bridge.
  // Every text box on the slide is typed in the body that draws it
  // at rest: a plain placeholder's rows (the master's bullet, its
  // 90% line spacing, the hanging indent) were drawn at rest and
  // swapped for the bare textarea on the first click, so its words
  // changed size, spacing and wrap on entering editing and again
  // on leaving it. PowerPoint types in the words as they are drawn.
  const importedTextHasPreview =
    block.type === "text" &&
    (isGeometryBlock || (importedPptxText && importedPptxTextHasPreview(block)));
  // The same object while the block is (the text body is memoized on its ctl).
  const importedTextInsets = useMemo(
    () =>
      importedPptxText || (block.type === "text" && textBlockDrawsInsets(block))
        ? textBlockInsetsCss(block)
        : NO_INSETS,
    [block, importedPptxText],
  );
  const nativeGroupExactHandleMembers = ownsSelectionChrome
    ? exactNativeGroupMembers(activeSlide.blocks, block)
    : null;
  // A group that carries one move authority moves whole and only
  // moves: it wears the move handle, never resize or rotate.
  const nativeGroupMoveOnlyMembers =
    ownsSelectionChrome && !nativeGroupExactHandleMembers
      ? nativeGroupMoveMembers(activeSlide.blocks, block)
      : null;
  const nativeGroupMoveOnly = nativeGroupMoveOnlyMembers !== null;
  const nativeGroupHandleMembers =
    nativeGroupExactHandleMembers ?? nativeGroupMoveOnlyMembers;
  const nativeGroupHandleBlocked =
    ownsSelectionChrome &&
    blockClaimsImportedGroup(block) &&
    (!nativeGroupHandleMembers ||
      nativeGroupHandleMembers.some((member) =>
        isBlockLocked(member),
      ));
  // While this block is being dragged/resized, render the live
  // (uncommitted) rect; otherwise use the persisted geometry.
  // Every member of a turn swings about the selection's centre,
  // the block under the hand too, as the commit lands them
  // (commitBlockRotation): it turned in place and jumped there.
  const rotateMemberPreview = view.rotateMember;
  const rawPreviewGeom =
    view.dragGeom ??
        (rotateMemberPreview
          ? {
              left: rotateMemberPreview.left,
              top: rotateMemberPreview.top,
              width: rotateMemberPreview.width,
              height: rotateMemberPreview.height,
            }
          : undefined);
  const tableGesturePreview =
    block.type === "table" && rawPreviewGeom
      ? tableWithScaledGestureGeometry(block, rawPreviewGeom)
      : null;
  const previewGeom = tableGesturePreview
    ? {
        left:
          geometryNumber(tableGesturePreview, "left") ??
          rawPreviewGeom!.left,
        top:
          geometryNumber(tableGesturePreview, "top") ??
          rawPreviewGeom!.top,
        width:
          geometryNumber(tableGesturePreview, "width") ??
          rawPreviewGeom!.width,
        height:
          geometryNumber(tableGesturePreview, "height") ??
          rawPreviewGeom!.height,
      }
    : rawPreviewGeom;
  // While the rotate handle is dragged, render the live rotation so
  // the block visibly turns without mutating slide state.
  const previewRotation = rotateMemberPreview
    ? Math.round(normalizeRotationDegrees(rotateMemberPreview.rotation))
    : view.rotateRotation;
  // A handle dragged through the pinned edge previews the mirror.
  const previewFlip = view.dragFlip;
  // The diagram camera of styles 3d5-3d9: ONE transform, identical
  // on every member, about the bounding box of the whole group, so
  // the diagram tilts as a single object instead of each node
  // spinning about its own centre. Approximate (§5): the selection
  // overlay and the hit areas stay axis-aligned.
  // Edit in 2-D: the graphic is drawn flat (no camera, no solid),
  // and until then its shapes take no move, size or turn of
  // their own (lib/office-smartart/edit-in-2d).
  const smartArtMarkerOfBlock = readSmartArtBlockMarker(block);
  const smartArtFlat2d =
    smartArtMarkerOfBlock !== null &&
    smartArtMemberHasScene(block) &&
    smartArtEditingIn2d(smartArtEditIn2dKey("presentation", smartArtMarkerOfBlock.groupId));
  const smartArtSceneLocked =
    ownsSelectionChrome &&
    smartArtMarkerOfBlock !== null &&
    !sel.smartArtGraphicOnly &&
    presentationSmartArtShapesLockedIn3d(activeSlide.blocks, block);
  const smartArtScene3d = smartArtFlat2d
    ? undefined
    : (block as { scene3d?: SmartArtScene3d }).scene3d;
  const smartArtMemberBox = smartArtScene3d
    ? (() => {
        const geom = blockGeometry(block);
        return typeof geom.left === "number" &&
          typeof geom.top === "number" &&
          typeof geom.width === "number" &&
          typeof geom.height === "number"
          ? {
              left: geom.left,
              top: geom.top,
              width: geom.width,
              height: geom.height,
            }
          : null;
      })()
    : null;
  const smartArtCamera =
    smartArtScene3d && smartArtMemberBox
      ? smartArtCameraCss(
          smartArtScene3d,
          presentationSmartArtGroupBox(
            activeSlideVisibleBlocks,
            blockGroupId(block) ?? null,
          ) ?? smartArtMemberBox,
          smartArtMemberBox,
        )
      : null;
  // The member's own flip/rotate must survive, so the camera
  // multiplies IN FRONT of it rather than replacing it. The origin
  // is the diagram's centre for the tilt to be shared; the member's
  // own rotation still turns about the member's centre
  // (smartArtCameraWithOwnTransform: Basic Cycle's arrows).
  const smartArtCameraStyle = smartArtCamera
    ? (() => {
        const own = blockGeometryStyle(
          block,
          activeSlideBounds,
          previewGeom,
          previewRotation,
          previewFlip,
        )?.transform;
        return {
          transform: smartArtCameraWithOwnTransform(
            smartArtCamera,
            typeof own === "string" ? own : null,
          ),
          transformOrigin: smartArtCamera.transformOrigin,
          transformStyle: smartArtCamera.transformStyle,
        };
      })()
    : null;
  // Direct-manipulation handles are shown only for the selected
  // positioned block in the editor (never present mode). They give a
  // non-text affordance to start a gesture so text editing inside a
  // text block is never disturbed. A LOCKED block shows NO handles
  // (it can't be moved/resized/rotated) — the gesture starters are
  // also guarded, this just removes the affordance.
  const showHandles =
    isGeometryBlock &&
    ownsSelectionChrome &&
    !(sel.smartArtGraphicOnly && readSmartArtBlockMarker(block)) &&
    !isBlockLocked(block) &&
    !nativeGroupHandleBlocked &&
    canChangePresentationChartFrame(block, "move");
  // A selection that holds a picture is not resized by hand
  // (beginBlockGesture: a picture's size goes through its own Size
  // transaction): its grips are not offered, rather than drawn
  
  // shape's grips, which did nothing).
  // A member selected inside its group wears its own grips: it is
  // sized alone, a picture by its own frame in the group's space.
  // The file's frame locks: the group's on its frame, the
  // object's own otherwise.
  const frameLockScope = Boolean(blockGroupId(block)) && !memberAloneChrome;
  const showResizeHandles =
    !isFrameLockedByFile(block, "resize", frameLockScope) &&
    showHandles && (!nativeGroupMoveOnly || groupChromeOnBox || memberAloneChrome) && (groupChromeOnBox && nativeGroupHandleMembers ? true : memberAloneChrome ? true : block.type !== "image" ? (
      !activeSlide.blocks.some((candidate) => candidate.type === "image" && sel.selectedBlockIds.has(candidate.id))
    ) : (
      !view.cropEditing && sel.pictureCanResize()
    ));
  // The grips are painted in design pixels inside the scaled slide
  // surface, so the frame is measured in the same unit: a side too
  // short for its middle grip to clear the corners drops it, as
  // PowerPoint does (a thin bar's middles sat on its corners).
  // A line or connector shows its two end points, as PowerPoint
  
  // of the frame, which the wrapper mirrors with the flips, so
  // the ends are the frame's nw/se corners (ne/sw for a legacy
  // "trbl" line, mirrored inside the box).
  const connectorLine =
    block.type === "shape" &&
    isPresentationConnectorShapeKind(normalizeShapeKind(block.shape));
  // On the group's box the grips scale every member. A group the
  // file holds is sized by its own frame, pictures and all; a group
  // made here has each picture sized through its own Size
  // transaction, so it offers none (as `showResizeHandles` does for
  // a selection holding a picture).
  const visibleResizeHandles = groupChromeOnBox && selectedGroupView
    ? showResizeHandles &&
      (nativeGroupHandleMembers !== null ||
        !selectedGeometryBlocks.some((member) => member.type === "image"))
      ? officeDrawingVisibleResizeHandles(
          selectedGroupView.box.width / AUTOFIT_EMU_PER_PX,
          selectedGroupView.box.height / AUTOFIT_EMU_PER_PX,
        )
      : []
    : showResizeHandles && connectorLine
    ? legacyLineSlopeFlipsHorizontally(
        (block as Record<string, unknown>).lineSlope,
      )
      ? (["ne", "sw"] as const)
      : (["nw", "se"] as const)
    : showResizeHandles
    ? officeDrawingVisibleResizeHandles(
        (previewGeom?.width ?? geometryNumber(block, "width") ?? Number.NaN) / AUTOFIT_EMU_PER_PX,
        (previewGeom?.height ?? geometryNumber(block, "height") ?? Number.NaN) / AUTOFIT_EMU_PER_PX,
      )
    : [];
  const tableFrameAlignedBlock =
    block.type === "table"
      ? tableWithFrameAlignedGrid(block)
      : null;
  const tableInteractionBlock =
    (tableKind?.tableDrawBoundaryPreview?.id === block.id ? tableKind.tableDrawBoundaryPreview : null) ?? tableGesturePreview ?? tableFrameAlignedBlock;
  const tableBlockWithStylePreview =
    tableInteractionBlock &&
    tableKind?.tableStylePreview?.blockId === block.id
      ? applyPresentationTableStyleSelection(
          tableInteractionBlock,
          tableKind.tableStylePreview.selection,
        )
      : tableInteractionBlock;
  const displayedTableBlock =
    tableBlockWithStylePreview &&
    tableAxisPreview && tableAxisPreview.blockId === block.id
      ? ({
          ...tableBlockWithStylePreview,
          [tableAxisPreview.axis === "col"
            ? "colWidths"
            : "rowHeights"]: setAxisSize(
              tableAxisPreview.axis === "col"
              ? tableBlockWithStylePreview.colWidths
              : tableBlockWithStylePreview.rowHeights,
            tableAxisPreview.index,
            tableAxisPreview.size,
          ),
        } as PresentationTableBlock)
      : tableBlockWithStylePreview;
  const pictureFramePreviewBlock =
    block.type === "image" &&
    view.pictureStyle !== undefined
      ? applyPresentationPictureFrameStyle(
          block,
          view.pictureStyle,
        )
      : block.type === "image"
        ? block
        : null;
  const displayedImageBlock = pictureFramePreviewBlock && previewGeom
    ? { ...pictureFramePreviewBlock, width: previewGeom.width, height: previewGeom.height }
    : pictureFramePreviewBlock;
  // An object the validator refused, kept read-only (TRIAGE 152).
  const degradedObject = degradedObjectOf(block);
  const pictureSmartArtMemberId = view.pictureSmartArtMemberId;
  const pictureSmartArtReplaceDisabledReason = (() => {
    if (!pictureSmartArtMemberId) return null;
    if (shared.readOnly) {
      return shared.readOnlyReason ?? t("ai.readOnlyLock");
    }
    if (isBlockKeptReadOnly(block)) {
      return t("presentation.contextMenu.lock");
    }
    if (
      block.source?.pictureSourceDisposition === "preserve-only"
    ) {
      return t("presentation.image.sourcePreserveOnly");
    }
    if (!shared.notebookId) return t("common.noNotebook");
    return shared.pictureReplaceBusy ? t("common.loading") : null;
  })();
  const pictureSmartArtReplaceLabel = `${t("presentation.image.replace")}${
    pictureSmartArtReplaceDisabledReason
      ? `: ${pictureSmartArtReplaceDisabledReason}`
      : ""
  }`;
  const autofitSmartTagVisible =
    isSelected &&
    block.type === "text" &&
    autofitSmartTagVisibleForBlock(
      block,
      activeSlideBounds,
      textBlockStyleCss(
        block.style,
        resolvePresentationTextStyle(activeSlide, block, masters, theme),
        fontSizeUnit, fontDisplay.aliases,
      ).fontSize as number | undefined,
    );
  const tableCellPositions =
    block.type === "table"
      ? logicalCellPositions(block.rows)
      : null;
  const tableColumnCount =
    block.type === "table" ? tableLogicalWidth(block.rows) : 0;
  // The block's frame style at a geometry: the persisted one, this
  // desk's own gesture preview, or another member's live frame
  // (PresentationLiveBlockFrame re-renders only this block for it).
  const blockFrameStyleAt = (
    frameGeom: BlockGeometry | undefined,
    frameRotation: number | undefined,
  ): CSSProperties | undefined =>
    ((placed) => {
      // A stroke of ink draws from the slide's corner, as the
      // wet layer does (presentationInkBlockPlacement).
      const style = (block as { shape?: unknown }).shape === "ink"
        ? presentationInkBlockPlacement(placed)
        : placed;
      const blend = presentationShapeBlendCss(block as { type?: unknown; blendMode?: unknown });
      const blended = withFooterFieldStacking(
        blend && style ? { ...style, ...blend } : style,
        activeSlide,
        block,
      );
      return smartArtLabelYieldsToShape && blended
        ? { ...blended, pointerEvents: "none" as const }
        : blended;
    })(withPresentationBlackWhiteStyle(
      viewColorMode,
      block as { blackWhiteMode?: unknown },
      importedPptxBlock || block.type === "image" || smartArtMember
        ? {
            ...(blockGeometryStyle(
              block,
              activeSlideBounds,
              frameGeom,
              frameRotation,
              previewFlip,
            ) ?? {}),
            ...(smartArtCameraStyle ?? {}),
            background: "transparent",
            border: "0 solid transparent",
            borderRadius: 0,
            boxShadow: "none",
            // At rest nothing may paint a ring around the
            // object; once selected the shared drawing frame
            // class paints it, like every other block.
            outline: wearsDrawingSelection && (isSelected || isMultiSelected) ? undefined : "none",
            outlineOffset: 0,
            // At rest the frame cuts what the slide show cuts
            // (`presentationBlockFrameOverflow`, one rule for
            // both); the slide canvas remains the clip edge.
            // Explicit clip/ellipsis is enforced again on the
            // inner text surface so selected handles stay visible.
            overflow:
              presentationBlockFrameOverflow(block) === "visible"
                ? "visible"
                : isSelected
                  ? "visible"
                  : "hidden",
          }
        : // A text box's words past its frame show at rest as
          // they do in the show and the file (Do not Autofit;
          // only a stated clip cuts them), as imported text's
          
          presentationBlockFrameOverflow(block) === "visible"
          ? {
              // The outline is centred on the geometry path, so
              // half of `a:ln@w` is drawn OUTSIDE this frame. The
              // frame is the geometry box the file states and is
              // untouched; it just stops cutting the shape's own
              // ink, exactly as the imported branch above and the
              // Docs shape stage already do.
              ...(blockGeometryStyle(
                block,
                activeSlideBounds,
                frameGeom,
                frameRotation,
                previewFlip,
              ) ?? {}),
              overflow: "visible",
            }
          : blockGeometryStyle(
              block,
              activeSlideBounds,
              frameGeom,
              frameRotation,
              previewFlip,
            ),
    ));
  return (
    <PresentationSmartArtGraphicHitLayer ctl={{ blocks: activeSlideVisibleBlocks, block, activeSlideBounds, primaryId: view.smartArtPrimaryId, inert: shared.hitLayerInert, selectBlockFromClick, setSmartArtGraphicOnly, beginBlockGesture, setBlockContextMenu, editSmartArtShapeText, zIndex: presentationBlockOverFooterFieldsZIndex(activeSlide, block) }}>
    <PresentationLiveBlockFrame
      className={
        (isGeometryBlock
          ? "presentation-window__block presentation-window__block--geometry"
          : "presentation-window__block") + (block.type === "table" ? ` ${presentationChrome.tableGeometryFrame}` : "") +
          (wearsDrawingSelection ? ` ${officeDrawingSelectionStyles.selectedHost}` : "")
          + presentationShapeGeometryHitClassName(block, previewMode)
      }
      style={blockFrameStyleAt(previewGeom, previewRotation)}
      styleAt={blockFrameStyleAt}
      liveSurface={previewMode ? null : liveTransformSurface}
      liveBlockId={block.id}
      committedFrame={presentationLiveFrameOf(block, activeSlideBounds)}
      ownPreview={previewGeom !== undefined || previewRotation !== undefined}
      data-block-id={block.id}
      // A layout's graphic is drawn here and taken in Slide Master
      // view: no press on it selects it (native-page-window.css).
      data-layout-graphic={isLayoutGraphicOnSlide(block) ? "true" : undefined}
      data-smartart-prepared={view.preparedRole}
      data-connector-line={connectorLine ? "true" : undefined}
      data-copilot-object-hidden={view.copilotHidden ? "true" : undefined}
      data-table-block={
        block.type === "table" ? "true" : undefined
      }
      data-placeholder-role={
        typeof (block as Record<string, unknown>).role === "string"
          ? String((block as Record<string, unknown>).role)
          : undefined
      }
      data-placeholder-empty={
        // An empty SmartArt shape draws its `[Text]` prompt inside
        // the shape, with no placeholder frame (PowerPoint).
        block.type === "text" &&
        block.text.trim() === "" &&
        (block as Record<string, unknown>).smartArtRole !== "label"
          ? "true"
          : undefined
      }
      data-smartart-role={
        typeof (block as Record<string, unknown>).smartArtRole === "string"
          ? String((block as Record<string, unknown>).smartArtRole)
          : undefined
      }
      // A SmartArt member is pressed where it is drawn: drawn turned
      // in 3-D, by its turned layer, not by its flat box, which
      // stood over the next item's words (native-page-window.css).
      // Not a label letting its presses through to its shape, nor
      // a shape already hit by its painted geometry.
      data-hit-where-drawn={
        !previewMode &&
        (isNativeSmartArtMember(block) || (smartArtMember && block.type === "text" && !smartArtLabelYieldsToShape))
          ? ""
          : undefined
      }
      // A11y: each block is a labelled group item with an accessible
      // name (text → its text, shape → its kind, image → its alt). It
      // owns the single roving tab stop when it is the roving block
      // (tabIndex 0), else it is removed from the tab order (-1), so
      // exactly one block is Tab-reachable at a time — Tab/arrows in
      // `handleBlockKeyDown` move between blocks / nudge geometry.
      role="group"
      aria-label={blockAriaName(block)}
      aria-roledescription={t(
        `presentation.block.${block.type}` as "presentation.block.text",
      )}
      tabIndex={view.isRoving ? 0 : -1}
      onKeyDown={(event) => handleBlockKeyDown(event, block.id)}
      data-roving={view.isRoving ? "true" : undefined}
      data-selected={isSelected ? "true" : "false"}
      data-text-editing={textEditing ? "true" : undefined}
      data-multiselected={
        isMultiSelected || (ownsSelectionChrome && !isSelected)
          ? "true"
          : undefined
      }
      {...(view.inkLassoMember
        ? { [INK_LASSO_MEMBER_ATTRIBUTE]: "true" }
        : {})}
      data-group-frame-member={framedByGroup ? "true" : undefined}
      data-dragging={previewGeom ? "true" : undefined}
      // Armed-painter affordance: a paint cursor + the source-block cue.
      data-format-painter={view.formatPainterRole}
      // Seen in capture: the rich body stops its press from
      // bubbling here.
      onPointerDownCapture={(event) => {
        wordsPressBlockRef.current =
          block.type === "text" &&
          event.target instanceof Element &&
          event.target.closest(
            "textarea,[data-selection-owner='rich-preview']",
          )
            ? block.id
            : null;
      }}
      onPointerDown={(event) => {
        noteBlockTouchPress(event, block.id);
        if (block.type === "shape" && event.button === 0) {
          // Windows' double click: the second press within 500 ms
          // and 4 px of the first (SM_CXDOUBLECLK).
          const now = performance.now();
          const last = shapeLastPressRef.current;
          const second =
            event.detail >= 2 ||
            (last !== null &&
              last.blockId === block.id &&
              now - last.at <= 500 &&
              Math.abs(event.clientX - last.x) <= 4 &&
              Math.abs(event.clientY - last.y) <= 4);
          shapeWordPressRef.current = second ? { blockId: block.id, at: now } : null;
          shapeLastPressRef.current = second ? null : { blockId: block.id, at: now, x: event.clientX, y: event.clientY };
        }
        // A press on a member of the group selected whole: its
        // click selects that member inside the group.
        groupMemberPressRef.current =
          event.button === 0 &&
          !event.shiftKey && !event.ctrlKey && !event.metaKey &&
          view.groupMemberPressable
            ? block.id
            : null;
        // A finger on a text box nobody is typing in holds the
        
        // selects it and the next tap types (onMouseDown below).
        // A press on an EMPTY text box or placeholder nobody is
        // typing in is a press on the object too (PowerPoint drags
        // "Click to add text" by its middle); its click types
        // there (`editsTextOnTap`).
        const restingEmptyText =
          block.type === "text" &&
          event.target instanceof HTMLTextAreaElement &&
          document.activeElement !== event.target &&
          !(block as PresentationTextBlock).text;
        const touchOnRestingText =
          (block.type === "text" &&
            isTouchLikePointer(event) &&
            event.target instanceof HTMLTextAreaElement &&
            document.activeElement !== event.target) ||
          restingEmptyText;
        if (!touchOnRestingText && isTextEditTarget(event.target)) {
          textSelectionIntentRef.current += 1;
          return;
        }
        
        // NEXT_PUBLIC_PRESENTATION_PRESS_DRAG_V1): the guard below
        // admits an unselected manipulable block when the flag is
        // on. The four earlier failures were verification, not the
        // change — jsdom dispatch to the exact wrapper (the
        // multi-transform harness) sidesteps every trap this note
        // records; the notes stay for whoever drives it in a real
        // browser:
        //
        // PowerPoint and WPS move a shape from a press on it, selected
        // or not. Here the block must already be selected, so moving
        // something takes two presses, and the second has to land in
        // the middle because the eight resize grips cover the edges
        // and corners.
        //
        // The change is two lines: drop `showHandles` from this guard
        // and call `selectBlockFromClick` before `beginBlockGesture`.
        // Written and reverted four times for want of a way to VERIFY
        // it. What that verification actually needs, established since
        // and worth having written down:
        //
        //   - Drive it from `/dev/presentation-solo`, not the gallery.
        //     The gallery frames the editor in a fixed box, which makes
        //     screen coordinates unusable. That was the first three
        //     failures.
        //   - Judge "did it move" from the block's own inline left/top,
        //     NEVER from a screen boundingBox: the canvas shifts under
        //     the blocks whenever the toolbar changes height, and a
        //     `.last()` locator silently retargets. Measuring screen
        //     positions produced a convincing 335px "move" that never
        //     happened.
        //   - Judge "is it selected" from `data-selected`, not from a
        //     handle count.
        //   - Selection must come from the typed selection authority;
        //     Esc clears it completely, so a one-shape slide can still
        //     represent a genuine no-selection state.
        //   - With two, the second covers the whole grabbable area of
        //     the first — the insert cascade steps 4% of the slide
        //     while a default shape is 34% wide. A press aimed at the
        //     lower one lands on the upper one. Not obviously wrong
        //     (PowerPoint overlaps too), but it means the test has to
        //     hit-test for a point the lower shape actually owns
        //     instead of computing one.
        // PRESS-DRAG v1 (flag-gated): PowerPoint/WPS move a
        // shape from a press on it, selected or not — the press
        // both selects (beginBlockGesture sets the id) and starts
        // the move. Text editing survives untouched because a text
        // block's body is a textarea and the interactive-target
        // guard below bails before any gesture starts.
        // A finger always may: it lifts an unselected object only
        // after a hold, so a swipe across it still scrolls.
        const pressDragUnselected =
          !showHandles &&
          (presentationPressDragEnabled() || isTouchLikePointer(event)) &&
          hasLayoutGeometry(block) &&
          canManipulateBlock(block);
        if (
          (!showHandles && !pressDragUnselected) ||
          event.button !== 0
        ) {
          return;
        }
        const target = event.target;
        if (
          !touchOnRestingText &&
          target instanceof Element &&
          target.closest(
            "textarea,input,select,button,a,[role='button'],[contenteditable='true']",
          )
        ) {
          return;
        }
        beginBlockGesture(block.id, "move", event);
      }}
      onMouseDown={(event) => {
        // A double click on the words of a SmartArt shape being
        // typed in selects the word under it, as PowerPoint does,
        // where the label's move band straddles a frame as tight
        // as its words: the band took the second press, and the
        
        if (
          event.button === 0 &&
          event.detail === 2 &&
          block.type === "text" &&
          event.target instanceof Element &&
          event.target.closest(".presentation-window__text-move-band") &&
          // (A converted diagram's label, or the words of a file's own
          // drawing opened before it converts: the band straddles both.)
          (presentationSmartArtLabelShape(activeSlide.blocks, block) !== null ||
            readNativeSmartArtView(activeSlide.blocks, block) !== null)
        ) {
          const rich = event.currentTarget.querySelector<HTMLElement>(
            '[data-selection-owner="rich-preview"][contenteditable="true"]',
          );
          const words = rich ? presentationPreviewPlainText(rich) : "";
          const hit = rich && words
            ? presentationPreviewOffsetAtPoint(
                rich,
                event.clientX,
                event.clientY,
                words.length,
                undefined,
                [...event.currentTarget.querySelectorAll<HTMLElement>(
                  "textarea, .presentation-window__text-move-band",
                )],
              )
            : null;
          if (rich && hit !== null) {
            const [start, end] = presentationDoubleClickWord(words, hit);
            event.preventDefault();
            event.stopPropagation();
            if (rich.ownerDocument.activeElement !== rich) rich.focus({ preventScroll: true });
            setPresentationPreviewSelection(rich, start, end);
            return;
          }
        }
        // A finger's first tap on a text box selects the box, not
        // its text: no caret, no keyboard (the tap's compatibility
        // mousedown would focus the field). The next tap types.
        const press = touchPressRef.current;
        // The press on an empty box moves it or (a click) opens
        // it: the field takes no caret from the press itself.
        if (
          block.type === "text" &&
          !isCompatMouseFromTouch(event.nativeEvent) &&
          event.button === 0 &&
          event.target instanceof HTMLTextAreaElement &&
          document.activeElement !== event.target &&
          !(block as PresentationTextBlock).text &&
          canManipulateBlock(block)
        ) {
          event.preventDefault();
        }
        if (
          block.type === "text" &&
          press?.blockId === block.id &&
          !press.wasSelected &&
          isCompatMouseFromTouch(event.nativeEvent) &&
          event.target instanceof HTMLTextAreaElement &&
          document.activeElement !== event.target
        ) {
          event.preventDefault();
        }
        // PowerPoint for iPad: "tap anywhere on the table to
        // select the table". The first tap selects the table
        // object; a tap on a cell of the selected table puts the
        // caret there (and the cell's selection handles).
        const pressedCell =
          block.type === "table" && event.target instanceof Element
            ? event.target.closest('[data-testid="presentation-table-cell"]')
            : null;
        if (
          pressedCell &&
          press?.blockId === block.id &&
          !press.wasSelected &&
          isCompatMouseFromTouch(event.nativeEvent) &&
          !pressedCell.contains(document.activeElement)
        ) {
          event.preventDefault();
        }
      }}
      onClick={(event) => {
        event.stopPropagation();
        const pressBeganInWords =
          wordsPressBlockRef.current === block.id;
        wordsPressBlockRef.current = null;
        // The click that ends a press which moved or sized this
        // object: the selection it moved stays as it is.
        const gestureEnd = gestureEndClickRef.current;
        gestureEndClickRef.current = null;
        const pressedGroupMember = groupMemberPressRef.current === block.id;
        groupMemberPressRef.current = null;
        if (
          gestureEnd?.blockId === block.id &&
          performance.now() - gestureEnd.at < 1000
        ) {
          event.currentTarget.focus({ preventScroll: true });
          return;
        }
        // The second click on a selected group selects the object
        // clicked in it (its own frame and grips; the group's box
        // dashed): the group stays the selection its commands take.
        // A text box or a shape being typed in takes the caret
        // instead, as PowerPoint types in a grouped box's words.
        if (
          pressedGroupMember &&
          block.type !== "text" &&
          !view.editingShapeText
        ) {
          if (selectedBlockIdRef.current !== block.id) {
            selectedBlockIdRef.current = block.id;
            setSelectedBlockId(block.id);
          }
          setGroupMemberFocusId(block.id);
          event.currentTarget.focus({ preventScroll: true });
          return;
        }
        // The click that ends the press which just opened this
        // shape's text editor (the move gesture captured the
        // pointer, so it lands on the block): the caret stays in
        // the words -- focusing the block here blurred the editor
        // and closed it the moment it opened.
        if (block.type === "shape" && view.editingShapeText) {
          return;
        }
        // The same for a text box whose words took the caret --
        // the click ending the tap that opened them, or a click
        // in the words themselves -- unless a painter is armed:
        // its click paints (below). A click on the box's border
        // selects the box as an object.
        const endsTextTap = textTapClickRef.current === block.id;
        if (endsTextTap) textTapClickRef.current = null;
        if (
          block.type === "text" &&
          !formatPainterRef.current &&
          !animationPainterRef.current &&
          document.activeElement instanceof HTMLTextAreaElement &&
          event.currentTarget.contains(document.activeElement) &&
          (endsTextTap ||
            (event.target instanceof Element &&
              Boolean(
                // the words' own box (a press in the field
                // and its release on the drawn row click there)
                event.target.closest(".presentation-window__text-block"),
              )))
        ) {
          // The box being typed in is the selected one, as in
          // PowerPoint (its tools, Delete of the box after Esc).
          if (
            selectedBlockIdRef.current !== block.id ||
            !selectedBlockIdsRef.current.has(block.id)
          ) {
            selectBlockFromClick(block.id, false, true);
          }
          return;
        }
        const eventTargetOwnsTextSelection =
          (block.type === "text" || block.type === "table") &&
          (pressBeganInWords ||
            (event.target instanceof Element &&
              Boolean(
                event.target.closest(
                  "textarea,[data-selection-owner='rich-preview'],[data-selection-owner='cell-rich-editor']",
                ),
              )));
        // Canvas/object focus supersedes any cached run range,
        // including a range from this SAME text block. A click in
        // the real textarea/rich editor retains its live range so
        // an armed painter can apply to the highlighted target.
        if (!eventTargetOwnsTextSelection) {
          clearTextSelectionAuthority();
        }
        // FORMAT PAINTER: when armed, a click PAINTS the captured style
        // onto this block (one undo step) instead of selecting. Single
        // mode disarms after; sticky stays armed (handled inside).
        if (formatPainterRef.current) {
          paintBlockFormat(block.id);
          return;
        }
        // ANIMATION PAINTER: when armed, a click applies the captured
        // animation to this block (one undo step) instead of selecting.
        if (animationPainterRef.current) {
          applyAnimationToBlock(block.id);
          return;
        }
        // A click in a table's cell beside or under its words puts
        // the caret in that cell, on its nearest line, as in
        // PowerPoint ("click in a cell to type in it"): the cell's
        // words fill only their own lines, and a click on the rest
        // of the cell selected the table as an object and kept the
        // keys on the table, so typed letters went nowhere and Tab
        
        // first tap on a table still selects the table (iPad).
        // (The press that began on a selected table captured the
        // pointer for a move, so its click lands on the block: the
        // cell is the one under the pointer.)
        const cellSelector = "td[data-table-cell-row], th[data-table-cell-row]";
        const blockBox = event.currentTarget;
        const pressedCellBox =
          block.type === "table" &&
          !eventTargetOwnsTextSelection &&
          !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey &&
          event.target instanceof Element &&
          !event.target.closest("button,a,[role='button'],input,select,textarea")
            ? event.target.closest<HTMLElement>(cellSelector) ??
              ((blockBox.ownerDocument.elementsFromPoint?.(event.clientX, event.clientY) ?? [])
                .find((element): element is HTMLElement =>
                  element instanceof HTMLElement && element.matches(cellSelector) &&
                  blockBox.contains(element)) ?? null)
            : null;
        const cellWords = pressedCellBox?.querySelector<HTMLElement>(
          '[data-testid="presentation-table-cell"]',
        );
        if (
          cellWords &&
          !isTextEditLocked(block) &&
          (!isCompatMouseFromTouch(event.nativeEvent) || showHandles) &&
          officePressIntoTextHost(cellWords, event.clientX, event.clientY)
        ) {
          selectBlockFromClick(block.id, false, true);
          return;
        }
        if (
          eventTargetOwnsTextSelection ||
          isTextEditTarget(event.target)
        ) {
          // Cell, shape and other nested editors own their native
          // caret/range. Selecting their object must not focus its
          // border (or turn Shift text/cell selection into an
          // object multi-select). Only text blocks retain run authority.
          selectBlockFromClick(
            block.id,
            false,
            eventTargetOwnsTextSelection,
          );
          return;
        }
        // Safari can retain toolbar/body focus because the drag
        // pointerdown is prevented. Keep keys on the clicked object.
        event.currentTarget.focus({ preventScroll: true });
        const wantsMultiSelect =
          event.shiftKey || event.ctrlKey || event.metaKey;
        // The geometry multi-select set (align/distribute) only accepts
        // positioned blocks. A shift/ctrl/cmd-click on a NON-geometry
        // block must NOT touch that set — adding the id corrupts it with
        // a block that has no rect, and clearing it would silently drop
        // an in-progress multi-selection. So such a click only updates
        // the legacy single-selection pointer (the inspector follows it)
        // and leaves the geometry multi-select set untouched.
        if (wantsMultiSelect && !isGeometryBlock) {
          setSelectedBlockId(block.id);
          return;
        }
        // shift / ctrl / cmd click on a geometry block toggles
        // multi-select; a plain click selects only this block.
        selectBlockFromClick(
          block.id,
          isGeometryBlock && wantsMultiSelect,
        );
      }}
      onFocus={(event) => {
        // The editor gives its text back the focus after every
        // keystroke; re-selecting the selected object then would
        // re-render the whole editor for nothing.
        if (
          !view.legacySelected ||
          selectedBlockIdRef.current !== block.id
        ) {
          // Focus going into a box's words selects the box as
          // the click ending that press does (below), so the
          // press draws the window once, not once per event.
          if (
            block.type === "text" &&
            event.target instanceof HTMLTextAreaElement &&
            !formatPainterRef.current &&
            !animationPainterRef.current
          ) {
            selectBlockFromClick(block.id, false, true);
          } else {
            setSelectedBlockId(block.id);
          }
        }
        // Tab bringing focus onto an object from outside the
        // selection selects it whole: a grouped object as its
        // group, as a click and the next Tab do.
        const focusGroupId = (block as { groupId?: unknown }).groupId;
        if (
          event.target === event.currentTarget &&
          performance.now() - tabKeyAtRef.current < 400 &&
          typeof focusGroupId === "string" &&
          focusGroupId !== "" &&
          !selectedBlockIdsRef.current.has(block.id)
        ) {
          const members = new Set(
            activeSlide.blocks
              .filter(
                (candidate) =>
                  (candidate as { groupId?: unknown }).groupId === focusGroupId,
              )
              .map((candidate) => candidate.id),
          );
          selectedBlockIdRef.current = block.id;
          selectedBlockIdsRef.current = members;
          setSelectedBlockIds(members);
        }
        // Focus coming back to the words being typed -- a text
        // box's, a shape's label, a table cell's -- keeps the
        // range they hold: a ribbon tab, a closing menu or a
        // dialog gives it back, and a shape's selected word then
        // lost its range and Bold made the whole shape bold
        
        if (
          !(
            event.target instanceof Element &&
            event.target !== event.currentTarget &&
            event.target.closest(
              "textarea,[data-selection-owner='rich-preview'],[data-selection-owner='cell-rich-editor'],[data-shape-rich-editor='true']",
            )
          )
        ) {
          clearTextSelectionAuthority();
        }
      }}
      // Right-click a positioned block → its context menu (Copy / Paste
      // / z-order / Group / Ungroup / Delete). Only geometry blocks get
      // it; a plain body block falls through to the native menu.
      onContextMenu={
        isGeometryBlock
          ? (event) => openBlockContextMenu(event, block.id)
          : undefined
      }
      // The move gesture captures the pointer on this block, so
      // the browser delivers a double click here, not to the
      // chart body whose handler edits the title under the
      // pointer (PowerPoint: double-click a chart title to type
      // in it). The chart body gets the double click where it
      // happened; its own handler stops the bubble.
      onDoubleClick={
        block.type === "chart"
          ? (event) => {
              if (event.target !== event.currentTarget) return;
              const body = event.currentTarget.querySelector<HTMLElement>(
                ".presentation-window__chart-block",
              );
              if (!body) return;
              const hit =
                typeof document.elementFromPoint === "function"
                  ? document.elementFromPoint(event.clientX, event.clientY)
                  : null;
              (hit && body.contains(hit) ? hit : body).dispatchEvent(
                new MouseEvent("dblclick", {
                  bubbles: true,
                  cancelable: true,
                  composed: true,
                  clientX: event.clientX,
                  clientY: event.clientY,
                  screenX: event.screenX,
                  screenY: event.screenY,
                  detail: 2,
                }),
              );
            }
          : undefined
      }
    >
      {autofitSmartTagVisible ? (
        <PresentationBlockAutofitSmartTag
          ctl={{
            t,
            autofitSmartMenuBlockId: view.autofitMenuOpen ? block.id : null,
            setAutofitSmartMenuBlockId,
            autofitSmartMenuTriggerRef,
            setTextBlockAutofit,
            block,
          }}
        />
      ) : null}
      {








}
      {(block.type === "text" || presentationShapeIsNativeTextBox(block)) &&
      isGeometryBlock &&
      !smartArtLabelYieldsToShape && // its shape takes the press
      !isBlockLocked(block) &&
      !nativeGroupHandleBlocked &&
      canManipulateBlock(block)
        ? (["n", "e", "s", "w"] as const).map((edge) => (
            <div
              key={edge}
              className="presentation-window__text-move-band"
              data-edge={edge}
              data-testid={`presentation-text-move-band-${edge}`}
              aria-hidden="true"
            />
          ))
        : null}
      {showHandles
        ? ((handles) =>
            selectionChromeHost
              ? createPortal(
                  // The block's own frame and transform, mirrored,
                  // so the chrome sits where the block does (a
                  // drag preview included) while painting above
                  // every block.
                  <div
                    key={`chrome:${block.id}`}
                    className="presentation-window__selection-chrome"
                    data-testid="presentation-selection-chrome"
                    data-chrome-for={block.id}
                    // A stroke the lasso took: the one shared
                    // frame stands for its own frame and grips.
                    {...(view.inkLassoMember
                      ? { [INK_LASSO_MEMBER_ATTRIBUTE]: "true" }
                      : {})}
                    // One of the chart's elements selected: the
                    // chart's own frame and grips step back.
                    data-chart-element-selected={
                      block.type === "chart" && view.legacySelected && selectedChartElementArea
                        ? "true"
                        : undefined
                    }
                    data-group-frame={groupChromeOnBox ? "true" : undefined}
                    style={{
                      ...(groupChromeOnBox && selectedGroupView
                        ? {
                            position: "absolute",
                            left: `${(selectedGroupView.box.left / activeSlideBounds.width) * 100}%`,
                            top: `${(selectedGroupView.box.top / activeSlideBounds.height) * 100}%`,
                            width: `${(selectedGroupView.box.width / activeSlideBounds.width) * 100}%`,
                            height: `${(selectedGroupView.box.height / activeSlideBounds.height) * 100}%`,
                            transform: selectedGroupView.rotation
                              ? `rotate(${selectedGroupView.rotation}deg)`
                              : undefined,
                            transformOrigin: "center center",
                          }
                        : {
                            ...(blockGeometryStyle(
                              block,
                              activeSlideBounds,
                              previewGeom,
                              previewRotation,
                              previewFlip,
                            ) ?? {}),
                            ...(smartArtCameraStyle ?? {}),
                          }),
                      zIndex: undefined,
                      pointerEvents: "none",
                      overflow: "visible",
                    }}
                  >
                    {wearsDrawingSelection || groupChromeOnBox ? (
                      <div
                        className={officeDrawingSelectionStyles.frame}
                        data-text-editing={textEditing ? "true" : undefined}
                        aria-hidden="true"
                        // The object's own frame is its outline (the
                        // block's `selectedHost`); a group's is this
                        // line, drawn the size of the group's box.
                        style={groupChromeOnBox ? { inset: 0 } : undefined}
                      />
                    ) : null}
                    {handles}
                  </div>,
                  selectionChromeHost,
                )
              : handles)(
            <PresentationBlockHandles
              ctl={{
                t,
                // the grip's glyph keeps its size on screen
                slideDesignScale: slideScreenScale,
                smartArtFrameRotateGrip,
                activeSlideBounds,
                pictureSizeCommandRef,
                beginBlockGesture,
                beginRotateGesture,
                presentationShapeDrawingOverlay,
                block,
                // More than this one object selected: a lone selected shape shows its
                
                // selection too, so Edit Points drew nothing).
                isMultiSelected: isMultiSelected && sel.selectedBlockIds.size > 1,
                // On the group's own frame a group the file
                // nests turns whole (its `rot`); a member of it
                // does not turn alone.
                nativeGroupMoveOnly: nativeGroupMoveOnly && !groupChromeOnBox && !memberAloneChrome,
                showResizeHandles,
                visibleResizeHandles,
                sceneLocked: smartArtSceneLocked,
                // On the group's frame the grips are the group's.
                lineEnds: connectorLine && !groupChromeOnBox,
                rotateLocked: isFrameLockedByFile(block, "rotate", frameLockScope),
              }}
            />,
          )
        : null}
      <div
        className="presentation-window__block-order"
        style={{
          display: importedPptxBlock ? "none" : "flex",
          gap: 2,
          justifyContent: "flex-end",
          marginBottom: 2,
        }}
      >
        <button
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            moveBlock(block.id, "up");
          }}
          disabled={blockIndex === 0}
          title={t("presentation.moveBlockUp")}
          aria-label={t("presentation.moveBlockUp")}
          data-testid="presentation-move-block-up"
          style={blockOrderButtonStyle}
        >
          <ChevronUp size={12} />
        </button>
        <button
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            moveBlock(block.id, "down");
          }}
          disabled={blockIndex === activeSlide.blocks.length - 1}
          title={t("presentation.moveBlockDown")}
          aria-label={t("presentation.moveBlockDown")}
          data-testid="presentation-move-block-down"
          style={blockOrderButtonStyle}
        >
          <ChevronDown size={12} />
        </button>
      </div>
      {block.type === "image" ? (
        <PresentationCanvasImageBlock
          ctl={{
            t,
            cropEditingId: view.cropEditing ? block.id : null,
            cropPreview: view.cropPreview,
            replaceBlock,
            deleteBlock,
            selectBlockFromClick,
            beginReplaceImage,
            nudgeCropHandle,
            beginCropGesture,
            block,
            displayedImageBlock,
            pictureSmartArtMemberId,
            pictureSmartArtReplaceDisabledReason,
            pictureSmartArtReplaceLabel,
            deckThemeGridColors,
            pictureNotDrawn: view.pictureNotDrawn,
            model3dFailureNotice: view.inSelection
              ? t("presentation.preservedObject.model3d")
              : null,
          }}
        />
      ) : block.type === "chart" ? (
        <PresentationCanvasChartBlock
          ctl={{
            ...chartKind!.ctl,
            block,
            // A chart of a group selected whole is not selected
            // itself: its elements are picked once the chart is
            // (the next click), as PowerPoint goes group, chart,
            // element.
            isSelected:
              isSelected &&
              !(chartKind!.selectedGroupView?.ids.has(block.id) && !chartKind!.selectedGroupView.memberSelected),
            importedPptxBlock,
            previewGeom,
          }}
        />
      ) : block.type === "table" ? (
        <PresentationCanvasTableBlock
          ctl={{
            ...tableKind!.ctl,
            block,
            showHandles,
            displayedTableBlock,
            tableCellPositions,
            tableColumnCount,
          }}
        />
      ) : block.type === "shape" ? (
        <PresentationCanvasShapeBlock
          ctl={{
            t,
            setSelectedBlockId,
            editingShapeTextId: view.editingShapeText ? block.id : null,
            setEditingShapeTextId,
            shapeTextBlankDraft: view.shapeTextBlankDraft,
            setShapeTextBlankDraft,
            shapeAdjustPreview: view.shapeAdjustPreview,
            textSelectionIntentRef,
            cropEditingId: view.cropEditing ? block.id : null,
            cropPreview: view.cropPreview,
            fontDisplay,
            activeSlide,
            editSmartArtShapeText,
            handleShapeTextChange,
            nudgeCropHandle,
            beginCropGesture,
            recordShapeTextSelection,
            commitShapeRichText,
            recordShapeTextRange,
            takeShapeRichSelection,
            placeShapeTextCaret,
            heldBlockWords,
            editingHostHandover,
            toggleShapeWordsRunStyle,
            copyShapeText,
            pasteShapeTextRuns,
            undo,
            redo,
            block,
            isSelected,
            fontSizeUnit,
            previewGeom,
            smartArtFlat2d,
          }}
        />
      ) : block.type === "embed" ? (
        <PresentationCanvasEmbedBlock
          ctl={{
            ...embedKind!.ctl,
            block,
          }}
        />
      ) : (
        <PresentationCanvasTextBlock
          ctl={{
            t,
            theme,
            masters,
            findRevealRef,
            textSelectionRef,
            setTextSelectionSnapshot: keepTextSelectionSnapshot,
            textSelectionIntentRef,
            deckContentGenerationRef,
            fontDisplay,
            activeSlide,
            activeSlideBounds,
            undo,
            redo,
            showProofingErrors,
            deleteBlock,
            captureTextEditSelection,
            consumeTextEditSelection,
            updateTextBlock,
            composingTextRef,
            composingImportedPreviewRef,
            handleTextBlockBodyChange,
            captureImportedRichTextSelection,
            followImportedRichTextSelection,
            restoreImportedPreviewSelectionAfterCommit,
            pendingImportedPreviewCaret,
            heldBlockWords,
            editingHostHandover,
            commitImportedPreviewText,
            replaceImportedPreviewSelection,
            clearLiveRibbonList,
            pasteIntoImportedPreview,
            copyFromImportedPreview,
            adjustParagraphCaretLevel,
            applyTextBlockRunToggle,
            runStyleAtBodySelection,
            blockAriaName,
            handleFindRevealedKey,
            fillContentPlaceholder,
            block,
            // The equation being written in this box, drawn in
            // its line (only drawn: the box's words are its own).
            previewBlock: view.equationPreviewBlock,
            isSelected,
            fontSizeUnit,
            importedPptxText,
            importedTextHasPreview,
            importedTextInsets,
          }}
        />
      )}
      {degradedObject ? (
        <PresentationDegradedObjectOverlay t={t} record={degradedObject} isSelected={isSelected} />
      ) : null}
    </PresentationLiveBlockFrame>
    </PresentationSmartArtGraphicHitLayer>
  );
}

/** One block of the editing canvas, drawn again only when its own props change. */
export const PresentationCanvasBlockSlot = memo(
  PresentationCanvasBlockSlotBody,
  presentationCanvasBlockSlotPropsEqual,
);
