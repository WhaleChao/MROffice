"use client";
/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The ONE Shape Format tab of the three native editors.
 *
 * PowerPoint, Word and Excel all show the same contextual tab for a selected
 * shape — Insert Shapes | Shape Styles | WordArt Styles | Accessibility |
 * Arrange | Size (Word adds a Text group, because a Word shape is a text
 * container with direction, alignment and linked text boxes) — and MRAI used
 * to draw it three times: the document toolbar had the full tab, the slide
 * editor most of it, the worksheet a third of it, each with its own markup,
 * its own ids and its own gaps. This component is the tab; a host supplies an
 * {@link OfficeShapeFormatAdapter} — what the selection is, what it may
 * change, and how to change it — and nothing else.
 *
 * The rule a host adapter follows: a capability the host cannot honour is
 * still SHOWN, disabled, with the host's reason in its accessible label and
 * `data-disabled-reason`. Hiding the control would make the three tabs
 * differ again; a greyed control with a reason is the honest state.
 *
 * Every control's test id is `${testIdPrefix}-<control>`; hosts that shipped
 * other ids for the same control pass them in `ids` so their galleries,
 * walkthroughs and tests keep working unchanged.
 */
import { Fragment, type ReactNode, useState } from "react";
import { Maximize2, Minimize2, Square } from "./icons/houseIcons";

import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "../ui/dropdown-menu";
import type { OfficeShapePreset } from "../../lib/office-shape-catalog";
import type { PresentationCustomGeometryTool } from "../../lib/presentation-freeform";
import type { OfficeShapeStyleRecipe } from "../../lib/office-shape-style-catalog";
import type { OfficeWordArtRecipe } from "../../lib/office-wordart-style-catalog";

import { OfficeColorPaletteActionRow, OfficeColorPalettePanel } from "./OfficeColorPalettePanel";
import { OfficeColorPalettePopover } from "./OfficeColorPalettePopover";
import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
  OfficeHomeSplitLayout,
} from "./OfficeHomeRibbon";
import { OfficeLineEndsPicker } from "./OfficeLineEndsPicker";
import {
  OfficeSectionedAlignMenuItems,
  type OfficeAlignMenuSectionLabels,
  type OfficeAlignTargetChoice,
} from "./OfficeObjectArrangeMenuItems";
import { OfficeOutlineMenuContent } from "./OfficeOutlineMenuContent";
import { OfficeShapeCatalogPanel } from "./OfficeShapeCatalogPanel";
import officeShapeCatalogStyles from "./OfficeShapeCatalogPanel.module.css";
import { OfficeShapeFormatStyleGroup } from "./OfficeShapeFormatStyleGroup";
import { OfficeShapeQuickGallery } from "./OfficeShapeQuickGallery";
import { OfficeShapeSizeFields } from "./OfficeShapeSizeFields";
import { OfficeWordArtStyleGallery } from "./OfficeWordArtStyleGallery";
import { OfficeSemanticIcon } from "./icons/OfficeSemanticIcon";
import { OFFICE_MENU_DASHES, type OfficeMenuDash } from "../../lib/office-stroke-dash";
import { OfficePictureMenuIcon } from "./icons/OfficePictureMenuIcon";
import type { OfficeOutlineSketch } from "./OfficeOutlineMenuContent";
import { OfficeTextTransformSubmenu, type OfficeTextTransformLabels } from "./OfficeTextTransformMenu";
import styles from "./OfficeShapeFormatTab.module.css";

export type OfficeShapeEffect =
  | "none"
  | "shadow"
  | "reflection"
  | "glow"
  | "softEdge"
  | "bevel"
  | "threeD";
export const OFFICE_SHAPE_EFFECTS: readonly OfficeShapeEffect[] = [
  "none",
  "shadow",
  "reflection",
  "glow",
  "softEdge",
  "bevel",
  "threeD",
];
export type OfficeShapeTextDirection =
  | "horizontal"
  | "rotate90"
  | "rotate270"
  | "stacked";
export type OfficeShapeTextVerticalAlign = "top" | "middle" | "bottom";
export type OfficeShapeOutlineStyle = "solid" | "dashed" | "dotted";
export type OfficeShapeZOrderAction =
  | "bringForward"
  | "bringToFront"
  | "sendBackward"
  | "sendToBack";
export type OfficeShapeAlignAction =
  | "left"
  | "center"
  | "right"
  | "top"
  | "middle"
  | "bottom"
  | "distributeHorizontal"
  | "distributeVertical";
export type OfficeShapeRotateAction =
  | "rotateRight90"
  | "rotateLeft90"
  | "flipVertical"
  | "flipHorizontal";

/** A control the host may not be able to honour: disabled with its reason. */
export interface OfficeShapeCapability {
  readonly disabled?: boolean;
  /** Shown in the accessible label and `data-disabled-reason`. Required when
   * `disabled` is true: a greyed control that cannot say why is the exact
   * silent gap this tab exists to end. */
  readonly disabledReason?: string;
  /** The host's command-authority id, carried onto the control. */
  readonly controlId?: string;
}

export interface OfficeShapeColorAdapter extends OfficeShapeCapability {
  /** Current colour, `"none"` for no fill/outline, `undefined` when mixed or unknown. */
  readonly value: string | "none" | undefined;
  /** The colour the split's primary face re-applies (Office keeps the last
   * pick under the bucket); defaults to the current value. */
  readonly lastColor?: string | "none";
  readonly paletteId?: string;
  /** A host that derives its theme columns from the document (the deck's
   * theme recipes) instead of a palette id. */
  readonly themeColumns?: Parameters<typeof OfficeColorPalettePanel>[0]["themeColumns"];
  readonly recents?: readonly string[];
  /** A multi-selection whose colours differ: the primary face cannot
   * re-apply anything and says so. */
  readonly mixed?: boolean;
  readonly mixedReason?: string;
  readonly onPick: (color: string | null) => void;
  /** Called when the palette opens or closes (Docs pins its write target). */
  readonly onOpenChange?: (open: boolean) => void;
  /** Rows under the palette, in Office's order (Shape Fill: Picture…, More
   * Gradients…); each closes the palette when picked. */
  readonly extras?: readonly (OfficeShapeCapability & {
    readonly id: string;
    readonly label: string;
    readonly onSelect: () => void;
  })[];
}

export interface OfficeShapeOutlineAdapter extends OfficeShapeColorAdapter {
  readonly weightPt: number | null;
  readonly onWeight: (pt: number) => void;
  /** An exact weight box under the ladder (PowerPoint's "More Lines" without
   * leaving the menu): the host renders its own field. */
  readonly weightField?: ReactNode;
  readonly weight?: OfficeShapeCapability;
  readonly lineStyle: OfficeShapeOutlineStyle | null;
  readonly onLineStyle: (style: OfficeShapeOutlineStyle) => void;
  /**
   * Office's eight Dashes (Solid, Round Dot, Square Dot, Dash, Dash Dot, Long
   * Dash, Long Dash Dot, Long Dash Dot Dot) for a host that keeps the exact
   * dash; it replaces the three above.
   */
  readonly dashes?: {
    readonly value: OfficeMenuDash | null;
    readonly onPick: (dash: OfficeMenuDash) => void;
    readonly optionLabel: (dash: OfficeMenuDash) => string;
  };
  /** Shape Outline > Sketched (None / Curved / Freehand / Scribble), listed
   * between Weight and Dashes as Office lists it. */
  readonly sketched?: OfficeShapeCapability & {
    readonly label: string;
    readonly value: OfficeOutlineSketch;
    readonly onPick: (sketch: OfficeOutlineSketch) => void;
    readonly optionLabel: (sketch: OfficeOutlineSketch) => string;
  };
  readonly lineStyleCapability?: OfficeShapeCapability;
  readonly lineEnds: Parameters<typeof OfficeLineEndsPicker>[0]["value"];
  readonly onLineEnds: NonNullable<
    Parameters<typeof OfficeLineEndsPicker>[0]["onChange"]
  >;
  readonly lineEndsCapability?: OfficeShapeCapability;
  readonly onMoreArrows?: () => void;
}

export interface OfficeShapeEffectAdapter extends OfficeShapeCapability {
  readonly value: OfficeShapeEffect;
  /** Which effects the host can write. The others stay listed, disabled. */
  readonly available: readonly OfficeShapeEffect[];
  readonly unavailableReason?: string;
  readonly onPick: (effect: OfficeShapeEffect) => void;
  /**
   * PowerPoint's Shape Effects: Preset, Shadow, Reflection, Glow, Soft
   * Edges, Bevel and 3-D Rotation, each a submenu that offers "No <effect>"
   * and the variations the host can write. Without it the menu stays the
   * flat list Docs and the worksheet show.
   */
  readonly cascade?: {
    readonly presets: string;
    readonly category: (effect: Exclude<OfficeShapeEffect, "none">) => string;
    readonly none: (effect: Exclude<OfficeShapeEffect, "none">) => string;
  };
  /**
   * A host that keeps Office's combinable effects (`shapeEffects`, the model
   * pictures share) draws Office's whole Shape Effects menu itself: Preset,
   * Shadow (Outer / Inner / Perspective), Reflection, Glow, Soft Edges, Bevel
   * and 3-D Rotation with every variation. It replaces the list above.
   */
  readonly render?: (trigger: {
    readonly label: string;
    /** The face's own short word (Office's "Effects" beside "Shape Effects"). */
    readonly displayLabel: string;
    readonly testId: string;
    readonly controlId?: string;
    /** Rows after the categories (Text Effects > Transform). */
    readonly extra?: ReactNode;
  }) => ReactNode;
}

export interface OfficeShapeMenuAction extends OfficeShapeCapability {
  readonly onSelect: () => void;
}

export interface OfficeShapeFormatLabels {
  readonly insertShapes: string;
  readonly shapes: string;
  readonly editShape: string;
  readonly changeShape: string;
  readonly editPoints: string;
  readonly editPointsReason: string;
  /** Edit Shape ▸ Reroute Connectors (when the host offers it). */
  readonly rerouteConnectors?: string;
  readonly textBox: string;
  readonly textBoxHorizontal: string;
  readonly textBoxVertical: string;
  readonly shapeStyles: string;
  readonly moreShapeStyles: string;
  readonly previousStyles: string;
  readonly nextStyles: string;
  readonly shapeFill: string;
  readonly fill: string;
  readonly shapeOutline: string;
  readonly outline: string;
  readonly shapeEffects: string;
  readonly effects: string;
  readonly effect: (effect: OfficeShapeEffect) => string;
  readonly outlineLine: string;
  readonly arrows: string;
  readonly weight: string;
  readonly lineStyle: string;
  readonly lineStyleOption: (style: OfficeShapeOutlineStyle) => string;
  readonly wordArtStyles: string;
  readonly moreWordArtStyles: string;
  readonly textFill: string;
  readonly textOutline: string;
  readonly textEffects: string;
  readonly text: string;
  readonly textDirection: string;
  readonly textDirectionOption: (direction: OfficeShapeTextDirection) => string;
  readonly alignText: string;
  readonly alignTextOption: (align: OfficeShapeTextVerticalAlign) => string;
  readonly createLink: string;
  readonly accessibility: string;
  readonly altText: string;
  readonly arrange: string;
  readonly position: string;
  readonly wrapText: string;
  readonly bringForward: string;
  readonly bringToFront: string;
  readonly sendBackward: string;
  readonly sendToBack: string;
  readonly selectionPane: string;
  readonly align: string;
  readonly alignOption: (action: OfficeShapeAlignAction) => string;
  readonly group: string;
  readonly ungroup: string;
  readonly rotate: string;
  readonly rotateOption: (action: OfficeShapeRotateAction) => string;
  readonly moreRotationOptions: string;
  readonly resetRotation?: string;
  /** Why the forward (backward) z-order menu is greyed: already at that end.
   * Used when the host names no per-direction reason of its own. */
  readonly alreadyAtFront?: string;
  readonly alreadyAtBack?: string;
  readonly size: string;
  readonly width: string;
  readonly height: string;
  readonly lockAspectRatio: string;
  /** The size fields' unit as the host's locale writes it (PowerPoint zh: 厘米). Default "cm". */
  readonly sizeUnit?: string;
  readonly mixed: string;
  readonly formatPane: string;
  readonly themeColors: string;
  readonly standardColors: string;
  readonly recentColors: string;
  readonly moreFillColors: string;
  readonly moreOutlineColors: string;
  readonly moreTextColors: string;
  readonly noFill: string;
  readonly noOutline: string;
}

export interface OfficeShapeFormatAdapter {
  readonly application: "docs" | "spreadsheet" | "presentation";
  readonly labels: OfficeShapeFormatLabels;
  /** The editor's own quick-insert affordance (a pointer-down that remembers
   * the caret, for instance). */
  readonly onRememberSelection?: () => void;
  /** Insert Shapes: absent when the host offers insertion elsewhere only. */
  readonly insert: {
    readonly onInsert: (preset: OfficeShapePreset) => boolean | void;
    readonly recentStorageKey: string;
    /** A host with its own catalog (the deck's ShapeKind gallery) renders it
     * here instead of the shared preset catalog; the quick strip is its own
     * too when `quick` is given. */
    readonly renderCatalog?: () => ReactNode;
    readonly quick?: ReactNode;
    /** Lines > Curve / Freeform: Shape / Freeform: Scribble: arm the host's drawing surface. */
    readonly onInsertTool?: (tool: PresentationCustomGeometryTool) => void;
  } & OfficeShapeCapability;
  readonly changeShape: {
    readonly onChange: (preset: OfficeShapePreset) => boolean | void;
    readonly selectedPreset?: OfficeShapePreset;
    readonly recentStorageKey: string;
    readonly tileIdentity?: (preset: OfficeShapePreset) => string;
    readonly renderCatalog?: () => ReactNode;
    /** Mixed or unknown selection: the trigger shows the mixed state. */
    readonly mixed?: boolean;
  } & OfficeShapeCapability;
  readonly editPoints?: OfficeShapeMenuAction;
  /** PowerPoint's Edit Shape ▸ Reroute Connectors; absent, no item. */
  readonly rerouteConnectors?: OfficeShapeMenuAction;
  /** Change Shape as a submenu of Edit Shape, beside Edit Points and Reroute
   * Connectors (PowerPoint, Word and Excel alike; every MRAI editor sets it).
   * Without it the gallery fills the menu itself. */
  readonly changeShapeSubmenu?: boolean;
  readonly textBox: {
    readonly onInsert: (direction: "horizontal" | "rotate90") => void;
    /** Word offers horizontal and vertical; PowerPoint and Excel draw one
     * text box (a single direction renders a plain button). */
    readonly directions?: readonly ("horizontal" | "rotate90")[];
  } & OfficeShapeCapability;
  /**
   * PowerPoint's Insert Shapes ▸ Merge Shapes ▾ (Union, Combine, Fragment,
   * Intersect, Subtract -- "Merge shapes", Microsoft Support). Optional: a
   * host without a merge engine omits it and the stack is unchanged.
   */
  readonly mergeShapes?: {
    readonly label: string;
    readonly items: readonly {
      readonly key: string;
      readonly label: string;
      readonly onSelect: () => void;
      readonly testId?: string;
    }[];
  } & OfficeShapeCapability;
  readonly styles: {
    readonly recipes: readonly OfficeShapeStyleRecipe[];
    readonly value: string | null;
    /** A host's own names; by default Office's, from the catalogue. */
    readonly labelFor?: (recipe: OfficeShapeStyleRecipe) => string;
    readonly onCommit: (recipe: OfficeShapeStyleRecipe) => void;
    readonly getDisabledReason?: (
      recipe: OfficeShapeStyleRecipe,
    ) => string | undefined;
    /** Quick samples shown before "more": Word shows 3, Excel 7. */
    readonly quickCount?: 3 | 7;
    /** The style row the inline strip starts with (default Intense Effect;
     * PowerPoint's strip opens on Subtle Effect). */
    readonly quickRow?: OfficeShapeStyleRecipe["row"];
  } & OfficeShapeCapability;
  readonly fill: OfficeShapeColorAdapter;
  readonly outline: OfficeShapeOutlineAdapter;
  readonly effects: OfficeShapeEffectAdapter;
  readonly wordArt: {
    /** A host's own WordArt gallery and its names (the deck's twenty under
     * its theme, Word's fifteen); by default the site's twenty. */
    readonly recipes?: readonly OfficeWordArtRecipe[];
    readonly labelFor?: (recipe: OfficeWordArtRecipe) => string;
    readonly sample?: string;
    readonly value: string | null;
    readonly onCommit: (recipe: OfficeWordArtRecipe) => void;
  } & OfficeShapeCapability;
  readonly textFill: OfficeShapeColorAdapter;
  readonly textOutline: OfficeShapeColorAdapter;
  readonly textEffects: OfficeShapeEffectAdapter;
  /** Text Effects > Transform: the words' preset warp (a token), or null. */
  readonly textTransform?: OfficeShapeCapability & {
    readonly value: string | null;
    readonly onPick: (token: string | null) => void;
    readonly labels: OfficeTextTransformLabels;
  };
  /** Word only: the Text group (Text Direction, Align Text, Create Link).
   * Excel's and PowerPoint's Shape Format tabs keep those on Home, so a host
   * without it draws no group (a documented host difference, not a gap). */
  readonly text?: {
    readonly direction: {
      readonly value: OfficeShapeTextDirection;
      readonly onPick: (direction: OfficeShapeTextDirection) => void;
    } & OfficeShapeCapability;
    readonly align: {
      readonly value: OfficeShapeTextVerticalAlign;
      readonly onPick: (align: OfficeShapeTextVerticalAlign) => void;
    } & OfficeShapeCapability;
    readonly createLink?: {
      readonly active: boolean;
      readonly onOpen: () => void;
    } & OfficeShapeCapability;
  };
  readonly altText: {
    readonly active: boolean;
    readonly onOpen: () => void;
  } & OfficeShapeCapability;
  /** Word only: position on the page and text wrapping (a documented host
   * difference — a slide and a sheet have no text to wrap). */
  readonly position?: {
    readonly icon: ReactNode;
    readonly menu: ReactNode;
  } & OfficeShapeCapability;
  readonly wrapText?: {
    readonly icon: ReactNode;
    readonly menu: ReactNode;
  } & OfficeShapeCapability;
  readonly zOrder: {
    readonly onAction: (action: OfficeShapeZOrderAction) => void;
    readonly atBoundary?: (action: OfficeShapeZOrderAction) => boolean;
    /** Hosts whose command authority names each move separately. */
    readonly controlIdFor?: (action: OfficeShapeZOrderAction) => string | undefined;
    /** A host that explains a closed direction separately (the deck's
     * "already at the front" versus "read-only"). */
    readonly disabledReasonFor?: (action: OfficeShapeZOrderAction) => string | undefined;
    /**
     * Word only (MsoZOrderCmd msoBringInFrontOfText / msoSendBehindText):
     * Bring Forward's menu ends with Bring in Front of Text and Send
     * Backward's with Send Behind Text, the text-layer moves a single
     * floating object has too. A slide and a sheet have no text layer.
     */
    readonly textLayer?: {
      readonly inFrontOfTextLabel: string;
      readonly behindTextLabel: string;
      readonly onInFrontOfText: () => void;
      readonly onBehindText: () => void;
      readonly disabled?: boolean;
    };
  } & OfficeShapeCapability;
  readonly selectionPane: OfficeShapeMenuAction;
  /** PowerPoint's Reorder Objects ▾ (Reorder Objects / Reorder Overlapping
   * Objects), between Selection Pane and Align. */
  readonly reorder?: {
    readonly label: string;
    readonly overlappingLabel: string;
    readonly onReorder: () => void;
    readonly onReorderOverlapping: () => void;
    readonly overlappingDisabled?: boolean;
  } & OfficeShapeCapability;
  readonly alignObjects: {
    readonly onAction: (action: OfficeShapeAlignAction) => void;
    readonly actions?: readonly OfficeShapeAlignAction[];
    readonly isActionDisabled?: (action: OfficeShapeAlignAction) => boolean;
    /** Office's "Align to Slide" / "Align to Page" toggle above the edges. */
    readonly alignToTarget?: {
      readonly checked: boolean;
      readonly label: string;
      readonly onChange: (checked: boolean) => void;
    };
    /** PowerPoint's sectioned menu: Horizontal / Vertical / Distribute heads
     * over the actions, and `target` (Align to Slide / Align Selected
     * Objects) as an exclusive pair at the bottom. Replaces `alignToTarget`. */
    readonly sections?: OfficeAlignMenuSectionLabels;
    readonly target?: OfficeAlignTargetChoice;
    /** Checked options closing the menu (Excel's Snap to Grid / Snap to
     * Shape), after the distribute commands. */
    readonly toggles?: readonly {
      readonly key: string;
      readonly checked: boolean;
      readonly label: string;
      readonly onChange: (checked: boolean) => void;
      readonly testId?: string;
    }[];
  } & OfficeShapeCapability;
  readonly group: {
    readonly grouped: boolean;
    readonly onGroup: () => void;
    readonly onUngroup: () => void;
    readonly canGroup?: boolean;
    readonly canUngroup?: boolean;
    /** PowerPoint's Regroup: re-forms the group the selection was last
     * ungrouped from. Shown under Ungroup when the host offers it. */
    readonly regroup?: OfficeShapeMenuAction & { readonly label: string };
    /** Extra actions under Group, drawn as a menu (the deck's Connect). */
    readonly extra?: readonly (OfficeShapeMenuAction & {
      readonly key: string;
      readonly label: string;
      readonly icon?: ReactNode;
      readonly testId?: string;
    })[];
  } & OfficeShapeCapability;
  /** A badge beside Arrange saying how many objects are selected. */
  readonly selectionCount?: { readonly count: number; readonly label: string };
  readonly rotate: {
    readonly onAction: (action: OfficeShapeRotateAction) => void;
    /** "More Rotation Options": opens the host's dialog or pane. */
    readonly exact?: OfficeShapeMenuAction;
    /** Docs' strength: an exact rotation box inside the menu, host-rendered. */
    readonly exactField?: ReactNode;
    /** Reset to 0 degrees (Docs offers it; PowerPoint does not). */
    readonly reset?: OfficeShapeMenuAction;
  } & OfficeShapeCapability;
  readonly size: {
    readonly width: string;
    readonly height: string;
    readonly mixed?: { width: boolean; height: boolean };
    readonly onCommit: (axis: "width" | "height", value: string) => void;
    readonly max?: number;
    /** Word's Size group lists Height above Width. */
    readonly heightFirst?: boolean;
    /** Arrow step in the unit the host shows (default 0.1). */
    readonly step?: number;
    readonly lockAspect: {
      readonly value: boolean;
      readonly onToggle: () => void;
    } & OfficeShapeCapability;
  } & OfficeShapeCapability;
  readonly formatPane?: OfficeShapeMenuAction & { readonly icon?: ReactNode };
}

/** Host ids for controls that predate the shared tab. */
export type OfficeShapeFormatIds = Partial<
  Record<
    | "insertGroup"
    | "insert"
    | "insertMore"
    | "insertCatalog"
    | "changeShape"
    | "changeShapeCatalog"
    | "changeShapeTile"
    | "editPoints"
    | "textBox"
    | "stylesGroup"
    | "styles"
    | "fill"
    | "fillSplit"
    | "fillPrimary"
    | "fillTrigger"
    | "outline"
    | "outlineSplit"
    | "outlinePrimary"
    | "outlineTrigger"
    | "outlineWeight"
    | "outlineDashes"
    | "arrows"
    | "outlineLineStyle"
    | "lineEnds"
    | "effects"
    | "effect"
    | "wordArtGroup"
    | "wordArt"
    | "textFill"
    | "textOutline"
    | "textEffects"
    | "textEffect"
    | "textGroup"
    | "textDirection"
    | "textAlign"
    | "createLink"
    | "altTextGroup"
    | "altText"
    | "arrangeGroup"
    | "position"
    | "wrapText"
    | "layer"
    | "layerSplit"
    | "layerMenu"
    | "send"
    | "sendSplit"
    | "sendMenu"
    | "selectionPane"
    | "reorder"
    | "reorderItem"
    | "reorderOverlappingItem"
    | "regroupItem"
    | "bringForwardItem"
    | "bringToFrontItem"
    | "sendBackwardItem"
    | "sendToBackItem"
    | "align"
    | "alignToTarget"
    | "alignLeft"
    | "alignCenter"
    | "alignRight"
    | "alignTop"
    | "alignMiddle"
    | "alignBottom"
    | "alignDistributeHorizontal"
    | "alignDistributeVertical"
    | "groupItem"
    | "ungroupItem"
    | "group"
    | "rotate"
    | "rotateExact"
    | "rotateReset"
    | "rotateRight90"
    | "rotateLeft90"
    | "flipHorizontal"
    | "flipVertical"
    | "sizeGroup"
    | "sizeRow"
    | "size"
    | "alignMenu"
    | "selectionCount"
    | "lockAspect"
    | "formatPane",
    string
  >
>;

export interface OfficeShapeFormatTabProps {
  readonly adapter: OfficeShapeFormatAdapter;
  readonly testIdPrefix: string;
  readonly ids?: OfficeShapeFormatIds;
  /** Docs tags its groups so its strip knows the tab they belong to. */
  readonly ribbon?: string;
  readonly menuClassName?: string;
  readonly onMenuCloseAutoFocus?: (event: Event) => void;
  /**
   * A host's own first group in place of Insert Shapes: PowerPoint's
   * SmartArt Format tab opens on Shapes (Change Shape, Larger, Smaller)
   * because a shape inside a graphic is not inserted from there. Absent,
   * the tab is unchanged.
   */
  readonly shapesGroup?: ReactNode;
  /**
   * The selection is a shape INSIDE a SmartArt graphic: the tab opens on
   * Office's SmartArt Format > Shapes group -- Change Shape (the adapter's
   * own catalog), Larger and Smaller -- instead of Insert Shapes, because a
   * shape inside a graphic is not inserted from there (Excel, PowerPoint and
   * Word's SmartArt Format tab). Ignored when `shapesGroup` is given.
   */
  readonly smartArtShapes?: {
    readonly groupLabel: string;
    readonly larger: string;
    readonly smaller: string;
    /** Larger / Smaller can act on the selected shape. */
    readonly canResize: boolean;
    readonly resizeDisabledReason?: string;
    readonly onResize: (direction: "larger" | "smaller") => void;
    /** Edit in 2-D, the first command of the stack (see {@link OfficeSmartArtEditIn2dProps}). */
    readonly editIn2d?: OfficeSmartArtEditIn2dProps;
  };
}

/**
 * SmartArt Format ▸ Shapes ▸ Edit in 2-D (Excel, PowerPoint and Word): a
 * toggle, pressed while the graphic is shown flat. Enabled only for a graphic
 * wearing a 3-D style with a perspective view (`lib/office-smartart/
 * edit-in-2d`); the same button in every editor.
 */
export interface OfficeSmartArtEditIn2dProps {
  readonly label: string;
  readonly pressed: boolean;
  readonly disabled: boolean;
  readonly disabledReason?: string;
  readonly onToggle: () => void;
}

export function OfficeSmartArtEditIn2dCommand({
  editIn2d,
  testId,
}: {
  editIn2d: OfficeSmartArtEditIn2dProps;
  testId: string;
}) {
  return (
    <OfficeHomeCommand
      variant="compact"
      controlId="SmartArtEditIn2D"
      icon={<Square size={16} />}
      label={editIn2d.label}
      active={editIn2d.pressed}
      disabled={editIn2d.disabled}
      disabledReason={editIn2d.disabled ? editIn2d.disabledReason : undefined}
      onClick={editIn2d.onToggle}
      testId={testId}
    />
  );
}

const reasonProps = (capability: OfficeShapeCapability | undefined) => ({
  disabled: Boolean(capability?.disabled),
  disabledReason: capability?.disabled ? capability.disabledReason : undefined,
  controlId: capability?.controlId,
});

function colorGlyph(icon: ReactNode, color: string | "none" | undefined) {
  return (
    <span className="note-editor-office-toolbar__color-glyph">
      {icon}
      <span
        className="note-editor-office-toolbar__color-underline"
        style={{ backgroundColor: !color || color === "none" ? "transparent" : color }}
      />
    </span>
  );
}

/**
 * Office's colour bucket: a split whose primary face re-applies the last
 * colour and whose trigger opens the palette. Fill opens the palette as a
 * popover; Outline opens a menu that also holds Weight, Dashes and Arrows,
 * the way PowerPoint's Shape Outline does.
 */
function ColorSplit({
  adapter,
  icon,
  label,
  displayLabel,
  ids,
  labels,
  moreColors,
  noFill,
  open,
  onOpenChange,
  outline,
  menuClassName,
}: {
  adapter: OfficeShapeColorAdapter;
  icon: ReactNode;
  label: string;
  displayLabel: string;
  ids: { split: string; primary: string; trigger: string; palette: string };
  labels: OfficeShapeFormatLabels;
  moreColors: string;
  noFill: string;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  outline?: {
    adapter: OfficeShapeOutlineAdapter;
    ids: { weight: string; dashes: string; arrows: string; lineEnds: string };
  };
  menuClassName: string;
}) {
  const { disabled, disabledReason, controlId } = reasonProps(adapter);
  const last = adapter.lastColor ?? adapter.value;
  const setOpen = (next: boolean) => {
    adapter.onOpenChange?.(next);
    onOpenChange(next);
  };
  const primaryDisabled = disabled || adapter.mixed === true || last === undefined;
  const primaryReason = disabled
    ? disabledReason
    : adapter.mixed || last === undefined
      ? adapter.mixedReason
      : undefined;
  const paletteProps = {
    paletteId: adapter.paletteId,
    themeColumns: adapter.themeColumns,
    recents: adapter.recents,
    value: adapter.value === "none" ? undefined : adapter.value,
    allowNoFill: true,
    noFillSelected: adapter.value === "none",
    labels: {
      themeColors: labels.themeColors,
      standardColors: labels.standardColors,
      recentColors: labels.recentColors,
      moreColors,
      noFill,
    },
    onPick: (color: string | null) => {
      adapter.onPick(color);
      setOpen(false);
    },
    testIdPrefix: ids.palette,
    ...(adapter.extras?.length
      ? {
          footer: (
            <>
              {adapter.extras.map((extra) => (
                <OfficeColorPaletteActionRow
                  key={extra.id}
                  label={extra.label}
                  disabled={extra.disabled}
                  disabledReason={extra.disabledReason}
                  testId={`${ids.palette}-${extra.id}`}
                  onSelect={() => {
                    setOpen(false);
                    extra.onSelect();
                  }}
                />
              ))}
            </>
          ),
        }
      : {}),
  };
  const primary = (
    <OfficeHomeCommand
      variant="compact"
      controlId={controlId}
      disabled={primaryDisabled}
      disabledReason={primaryReason}
      icon={colorGlyph(icon, adapter.mixed ? undefined : last)}
      displayLabel={displayLabel}
      label={label}
      onClick={() => {
        if (last === undefined) return;
        adapter.onPick(last === "none" ? null : last);
      }}
      testId={ids.primary}
    />
  );
  if (!outline) {
    return (
      <div className={styles.colorHost}>
        <OfficeHomeSplitLayout
          className={styles.colorSplit}
          testId={ids.split}
          primary={primary}
          trigger={
            <OfficeHomeMenuTrigger
              variant="icon"
              icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
              active={open}
              controlId={controlId}
              disabled={disabled}
              disabledReason={disabledReason}
              label={label}
              aria-expanded={open}
              aria-haspopup="dialog"
              onClick={() => setOpen(!open)}
              testId={ids.trigger}
            />
          }
        />
        {open ? (
          <OfficeColorPalettePopover
            {...paletteProps}
            onDismiss={() => setOpen(false)}
          />
        ) : null}
      </div>
    );
  }
  const weight = reasonProps(outline.adapter.weight ?? outline.adapter);
  const dashes = reasonProps(outline.adapter.lineStyleCapability ?? outline.adapter);
  const ends = reasonProps(outline.adapter.lineEndsCapability ?? outline.adapter);
  const sketch = outline.adapter.sketched;
  const outlineMenu = {
    palette: paletteProps,
    menuClassName,
    ...(sketch
      ? {
          sketched: {
            ...reasonProps(sketch),
            label: sketch.label,
            testId: `${outline.ids.dashes.replace(/-dashes$/, "")}-sketched`,
            value: sketch.value,
            optionLabel: sketch.optionLabel,
            onPick: sketch.onPick,
          },
        }
      : {}),
    weight: {
      ...weight,
      label: labels.weight,
      testId: outline.ids.weight,
      valuePt: outline.adapter.weightPt,
      onPick: outline.adapter.onWeight,
      field: outline.adapter.weightField,
    },
    arrows: {
      ...ends,
      label: labels.arrows,
      testId: outline.ids.arrows,
      content: (
        <OfficeLineEndsPicker
          value={outline.adapter.lineEnds}
          disabled={ends.disabled}
          disabledReason={ends.disabledReason}
          onChange={outline.adapter.onLineEnds}
          onMoreArrows={outline.adapter.onMoreArrows}
          testIdPrefix={outline.ids.lineEnds}
        />
      ),
    },
  };
  return (
    <div className={styles.colorHost}>
      <OfficeHomeSplitLayout
        className={styles.colorSplit}
        testId={ids.split}
        primary={primary}
        trigger={
          <DropdownMenu modal={false} open={open} onOpenChange={setOpen}>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                variant="icon"
                icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                active={open}
                controlId={controlId}
                disabled={disabled}
                disabledReason={disabledReason}
                label={label}
                testId={ids.trigger}
              />
            </DropdownMenuTrigger>
            <DropdownMenuContent align="start" className={menuClassName}>
              {outline.adapter.dashes ? (
                <OfficeOutlineMenuContent<OfficeMenuDash>
                  {...outlineMenu}
                  dashes={{
                    ...dashes,
                    label: labels.lineStyle,
                    testId: outline.ids.dashes,
                    value: outline.adapter.dashes.value,
                    options: OFFICE_MENU_DASHES,
                    optionLabel: outline.adapter.dashes.optionLabel,
                    optionIcon: (dash) => <OfficePictureMenuIcon icon={{ kind: "line-dash", dash }} />,
                    onPick: outline.adapter.dashes.onPick,
                  }}
                />
              ) : (
                <OfficeOutlineMenuContent
                  {...outlineMenu}
                  dashes={{
                    ...dashes,
                    label: labels.lineStyle,
                    testId: outline.ids.dashes,
                    value: outline.adapter.lineStyle ?? "solid",
                    optionLabel: labels.lineStyleOption,
                    onPick: outline.adapter.onLineStyle,
                  }}
                />
              )}
            </DropdownMenuContent>
          </DropdownMenu>
        }
      />
    </div>
  );
}

function EffectsMenu({
  adapter,
  label,
  displayLabel,
  labels,
  testId,
  itemTestIdPrefix,
  menuClassName,
  extra,
}: {
  /** Host rows after the effects (Text Effects > Transform). */
  extra?: ReactNode;
  adapter: OfficeShapeEffectAdapter;
  label: string;
  displayLabel: string;
  labels: OfficeShapeFormatLabels;
  testId: string;
  itemTestIdPrefix: string;
  menuClassName?: string;
}) {
  const { disabled, disabledReason, controlId } = reasonProps(adapter);
  if (adapter.render) return <>{adapter.render({ label, displayLabel, testId, controlId, extra })}</>;
  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          active={adapter.value !== "none"}
          controlId={controlId}
          disabled={disabled}
          disabledReason={disabledReason}
          hasMenuIndicator
          icon={<OfficeSemanticIcon semanticId="shape.effects" size={18} />}
          label={label}
          displayLabel={displayLabel}
          testId={testId}
          variant="compact"
        />
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className={menuClassName}>
        {adapter.cascade ? (
          <EffectsCascade
            adapter={adapter}
            cascade={adapter.cascade}
            labels={labels}
            itemTestIdPrefix={itemTestIdPrefix}
            menuClassName={menuClassName}
          />
        ) : (
        <DropdownMenuRadioGroup value={adapter.value}>
          {OFFICE_SHAPE_EFFECTS.map((effect) => {
            const available = adapter.available.includes(effect);
            return (
              <DropdownMenuRadioItem
                key={effect}
                data-testid={`${itemTestIdPrefix}-${effect}`}
                data-disabled-reason={
                  available ? undefined : adapter.unavailableReason
                }
                disabled={!available}
                title={available ? undefined : adapter.unavailableReason}
                onSelect={() => adapter.onPick(effect)}
                value={effect}
              >
                {labels.effect(effect)}
              </DropdownMenuRadioItem>
            );
          })}
        </DropdownMenuRadioGroup>
        )}
        {extra ? (
          <>
            <DropdownMenuSeparator />
            {extra}
          </>
        ) : null}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

/** The cascading Shape Effects menu (see `OfficeShapeEffectAdapter.cascade`). */
function EffectsCascade({
  adapter,
  cascade,
  labels,
  itemTestIdPrefix,
  menuClassName,
}: {
  adapter: OfficeShapeEffectAdapter;
  cascade: NonNullable<OfficeShapeEffectAdapter["cascade"]>;
  labels: OfficeShapeFormatLabels;
  itemTestIdPrefix: string;
  menuClassName?: string;
}) {
  const kinds = OFFICE_SHAPE_EFFECTS.filter(
    (effect): effect is Exclude<OfficeShapeEffect, "none"> => effect !== "none",
  );
  return (
    <>
      <DropdownMenuSub>
        <DropdownMenuSubTrigger
          disabled
          data-testid={`${itemTestIdPrefix}-presets`}
          data-disabled-reason={adapter.unavailableReason}
          title={adapter.unavailableReason}
        >
          {cascade.presets}
        </DropdownMenuSubTrigger>
      </DropdownMenuSub>
      {kinds.map((effect) => {
        const available = adapter.available.includes(effect);
        const current = adapter.value === effect;
        return (
          <DropdownMenuSub key={effect}>
            <DropdownMenuSubTrigger data-testid={`${itemTestIdPrefix}-${effect}-menu`}>
              {cascade.category(effect)}
            </DropdownMenuSubTrigger>
            <DropdownMenuSubContent className={menuClassName}>
              <DropdownMenuRadioGroup value={current ? effect : "none"}>
                <DropdownMenuRadioItem
                  value="none"
                  data-testid={`${itemTestIdPrefix}-${effect}-none`}
                  // "No Shadow" takes the shadow off; it leaves another
                  // effect the shape wears as it is.
                  onSelect={() => {
                    if (current) adapter.onPick("none");
                  }}
                >
                  {cascade.none(effect)}
                </DropdownMenuRadioItem>
                <DropdownMenuRadioItem
                  value={effect}
                  data-testid={`${itemTestIdPrefix}-${effect}`}
                  data-disabled-reason={available ? undefined : adapter.unavailableReason}
                  disabled={!available}
                  title={available ? undefined : adapter.unavailableReason}
                  onSelect={() => adapter.onPick(effect)}
                >
                  {labels.effect(effect)}
                </DropdownMenuRadioItem>
              </DropdownMenuRadioGroup>
            </DropdownMenuSubContent>
          </DropdownMenuSub>
        );
      })}
    </>
  );
}

export function OfficeShapeFormatTab({
  adapter,
  testIdPrefix,
  ids = {},
  ribbon,
  menuClassName = "note-editor-toolbar-menu",
  onMenuCloseAutoFocus,
  shapesGroup: hostShapesGroup,
  smartArtShapes,
}: OfficeShapeFormatTabProps) {
  const { labels } = adapter;
  const id = (key: keyof OfficeShapeFormatIds, fallback: string) =>
    ids[key] ?? `${testIdPrefix}-${fallback}`;
  const [openPalette, setOpenPalette] = useState<
    "fill" | "outline" | "textFill" | "textOutline" | null
  >(null);
  const paletteProps = (key: NonNullable<typeof openPalette>) => ({
    open: openPalette === key,
    onOpenChange: (open: boolean) => setOpenPalette(open ? key : null),
  });
  const menuContentProps = {
    align: "start" as const,
    className: menuClassName,
    onCloseAutoFocus: onMenuCloseAutoFocus,
  };
  const insert = reasonProps(adapter.insert);
  const change = reasonProps(adapter.changeShape);
  const shapesGroup =
    hostShapesGroup ??
    (smartArtShapes ? (
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={smartArtShapes.groupLabel}
        testId={`${testIdPrefix}-smartart-shapes-group`}
      >
        {/* Change Shape beside its stack (Edit in 2-D, Larger, Smaller), as
            Office lays the group out; the group itself stacks its children. */}
        <OfficeHomeCommandRow>
        <DropdownMenu modal={false}>
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              controlId={change.controlId}
              disabled={change.disabled}
              disabledReason={change.disabledReason}
              hasMenuIndicator
              icon={<OfficeSemanticIcon semanticId="shape.change" size={24} />}
              label={labels.changeShape}
              onPointerDown={adapter.onRememberSelection}
              testId={`${testIdPrefix}-smartart-change-shape`}
              variant="large"
            />
          </DropdownMenuTrigger>
          <DropdownMenuContent
            align="start"
            className={`${menuClassName} ${officeShapeCatalogStyles.menu}`}
            onCloseAutoFocus={onMenuCloseAutoFocus}
          >
            <DropdownMenuLabel>{labels.changeShape}</DropdownMenuLabel>
            {adapter.changeShape.renderCatalog?.() ?? (
              <OfficeShapeCatalogPanel
                application={adapter.application === "spreadsheet" ? "spreadsheet" : "document"}
                disabled={change.disabled}
                onSelect={adapter.changeShape.onChange}
                recentStorageKey={adapter.changeShape.recentStorageKey}
                selectedPreset={adapter.changeShape.selectedPreset}
                testIdPrefix={`${testIdPrefix}-smartart-change-shape`}
                tileTestIdPrefix={`${testIdPrefix}-smartart-change-shape`}
                tileIdentity={adapter.changeShape.tileIdentity}
              />
            )}
          </DropdownMenuContent>
        </DropdownMenu>
        <OfficeHomeCommandStack>
          {smartArtShapes.editIn2d ? (
            <OfficeSmartArtEditIn2dCommand
              editIn2d={smartArtShapes.editIn2d}
              testId={`${testIdPrefix}-smartart-edit-in-2d`}
            />
          ) : null}
          {(["larger", "smaller"] as const).map((direction) => (
            <OfficeHomeCommand
              key={direction}
              variant="compact"
              icon={direction === "larger" ? <Maximize2 size={16} /> : <Minimize2 size={16} />}
              label={smartArtShapes[direction]}
              disabled={!smartArtShapes.canResize}
              disabledReason={smartArtShapes.canResize ? undefined : smartArtShapes.resizeDisabledReason}
              onClick={() => smartArtShapes.onResize(direction)}
              testId={`${testIdPrefix}-smartart-${direction}`}
            />
          ))}
        </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
    ) : undefined);
  // PowerPoint's Edit Shape is a menu of three commands (Change Shape ▸,
  // Edit Points, Reroute Connectors): the button stays live while any of
  // them is, so a selected connector -- which has no shape to change --
  // still reaches Reroute Connectors; only the Change Shape submenu greys.
  const editShapeMenuLive = Boolean(
    adapter.changeShapeSubmenu &&
      ((adapter.editPoints && !adapter.editPoints.disabled) ||
        (adapter.rerouteConnectors && !adapter.rerouteConnectors.disabled)),
  );
  const editShapeTrigger =
    change.disabled && editShapeMenuLive
      ? { ...change, disabled: false, disabledReason: undefined }
      : change;
  const textBox = reasonProps(adapter.textBox);
  const stylesCap = reasonProps(adapter.styles);
  const outlineWeight = reasonProps(
    adapter.outline.weight ?? adapter.outline,
  );
  const outlineLineStyle = reasonProps(
    adapter.outline.lineStyleCapability ?? adapter.outline,
  );
  const lineEnds = reasonProps(adapter.outline.lineEndsCapability ?? adapter.outline);
  const wordArt = reasonProps(adapter.wordArt);
  const text = adapter.text;
  const textDirection = reasonProps(adapter.text?.direction);
  const textAlign = reasonProps(adapter.text?.align);
  const altText = reasonProps(adapter.altText);
  const zOrder = reasonProps(adapter.zOrder);
  const selectionPane = reasonProps(adapter.selectionPane);
  const alignObjects = reasonProps(adapter.alignObjects);
  const group = reasonProps(adapter.group);
  // The Group menu opens while anything in it can run: Regroup re-forms the
  // group a lone former member was ungrouped from even when Group and
  
  // greyed).
  const groupMenuDisabled = group.disabled && (adapter.group.regroup?.disabled ?? true);
  const rotate = reasonProps(adapter.rotate);
  const size = reasonProps(adapter.size);
  const lockAspect = reasonProps(adapter.size.lockAspect);
  const alignActions: readonly OfficeShapeAlignAction[] =
    adapter.alignObjects.actions ?? [
      "left",
      "center",
      "right",
      "top",
      "middle",
      "bottom",
      "distributeHorizontal",
      "distributeVertical",
    ];

  return (
    <>
      {/* ---- Insert Shapes -------------------------------------------- */}
      {shapesGroup ?? (
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.insertShapes}
        className={styles.insertGroup}
        testId={id("insertGroup", "insert-group")}
      >
        <OfficeHomeCommandRow className={styles.insertBody}>
          {adapter.insert.quick ?? (
          <OfficeShapeQuickGallery
            onSelect={adapter.insert.onInsert}
            onRememberSelection={adapter.onRememberSelection}
            disabled={insert.disabled}
            testIdPrefix={id("insert", "insert")}
            more={
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    controlId={insert.controlId}
                    disabled={insert.disabled}
                    disabledReason={insert.disabledReason}
                    geometry="gallery"
                    icon={
                      <OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />
                    }
                    label={labels.shapes}
                    onPointerDown={adapter.onRememberSelection}
                    testId={id("insertMore", "insert-more")}
                    variant="icon"
                  />
                </DropdownMenuTrigger>
                <DropdownMenuContent
                  {...menuContentProps}
                  className={`${menuClassName} ${officeShapeCatalogStyles.menu}`}
                >
                  {adapter.insert.renderCatalog?.() ?? (
                    <OfficeShapeCatalogPanel
                      application={
                        adapter.application === "spreadsheet" ? "spreadsheet" : "document"
                      }
                      disabled={insert.disabled}
                      onSelect={adapter.insert.onInsert}
                      onSelectTool={adapter.insert.onInsertTool}
                      recentStorageKey={adapter.insert.recentStorageKey}
                      testIdPrefix={ids.insertCatalog ?? id("insertMore", "insert-more")}
                    />
                  )}
                </DropdownMenuContent>
              </DropdownMenu>
            }
          />
          )}
          <OfficeHomeCommandStack>
            <DropdownMenu modal={false}>
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  active={adapter.changeShape.mixed ? "mixed" : false}
                  controlId={editShapeTrigger.controlId}
                  disabled={editShapeTrigger.disabled}
                  disabledReason={editShapeTrigger.disabledReason}
                  hasMenuIndicator
                  icon={<OfficeSemanticIcon semanticId="shape.change" size={24} />}
                  label={labels.editShape}
                  onPointerDown={adapter.onRememberSelection}
                  testId={id("changeShape", "change")}
                  variant="compact"
                />
              </DropdownMenuTrigger>
              <DropdownMenuContent
                {...menuContentProps}
                className={`${menuClassName} ${officeShapeCatalogStyles.menu}`}
              >
                {(() => {
                  const catalog = adapter.changeShape.renderCatalog?.() ?? (
                    <OfficeShapeCatalogPanel
                      application={
                        adapter.application === "spreadsheet" ? "spreadsheet" : "document"
                      }
                      disabled={change.disabled}
                      onSelect={adapter.changeShape.onChange}
                      recentStorageKey={adapter.changeShape.recentStorageKey}
                      selectedPreset={adapter.changeShape.selectedPreset}
                      testIdPrefix={ids.changeShapeCatalog ?? id("changeShape", "change")}
                      tileTestIdPrefix={id("changeShapeTile", "change")}
                      tileIdentity={adapter.changeShape.tileIdentity}
                    />
                  );
                  return adapter.changeShapeSubmenu ? (
                    <DropdownMenuSub>
                      <DropdownMenuSubTrigger
                        data-testid={`${id("changeShape", "change")}-submenu`}
                        disabled={change.disabled}
                        title={change.disabledReason}
                        data-disabled-reason={change.disabledReason}
                      >
                        <OfficeSemanticIcon semanticId="shape.change" size={16} />
                        {labels.changeShape}
                      </DropdownMenuSubTrigger>
                      <DropdownMenuSubContent
                        className={`${menuClassName} ${officeShapeCatalogStyles.menu} ${officeShapeCatalogStyles.submenu}`}
                      >
                        {catalog}
                      </DropdownMenuSubContent>
                    </DropdownMenuSub>
                  ) : (
                    <>
                      <DropdownMenuLabel>{labels.changeShape}</DropdownMenuLabel>
                      {catalog}
                      <DropdownMenuSeparator />
                    </>
                  );
                })()}
                {/* Edit Points: every host enters the shared vertex editor
                    (`lib/office-drawing-overlays`); an object with no single
                    outline keeps the item, disabled, saying why. */}
                {adapter.editPoints && !adapter.editPoints.disabled ? (
                  <DropdownMenuItem
                    data-testid={id("editPoints", "edit-points")}
                    onSelect={() => adapter.editPoints?.onSelect()}
                  >
                    <OfficeSemanticIcon semanticId="shape.editPoints" size={16} />
                    {labels.editPoints}
                  </DropdownMenuItem>
                ) : (
                  <DropdownMenuItem
                    disabled
                    data-testid={id("editPoints", "edit-points")}
                    data-disabled-reason={adapter.editPoints?.disabledReason ?? labels.editPointsReason}
                    aria-label={`${labels.editPoints}: ${
                      adapter.editPoints?.disabledReason ?? labels.editPointsReason
                    }`}
                  >
                    <OfficeSemanticIcon semanticId="shape.editPoints" size={16} />
                    <span className={styles.unavailableItem}>
                      <span>{labels.editPoints}</span>
                      <span
                        className={styles.unavailableReason}
                        data-testid={`${id("editPoints", "edit-points")}-reason`}
                      >
                        {adapter.editPoints?.disabledReason ?? labels.editPointsReason}
                      </span>
                    </span>
                  </DropdownMenuItem>
                )}
                {adapter.rerouteConnectors && labels.rerouteConnectors ? (
                  <DropdownMenuItem
                    disabled={adapter.rerouteConnectors.disabled}
                    title={adapter.rerouteConnectors.disabled ? adapter.rerouteConnectors.disabledReason : undefined}
                    data-disabled-reason={
                      adapter.rerouteConnectors.disabled ? adapter.rerouteConnectors.disabledReason : undefined
                    }
                    data-testid={`${id("changeShape", "change")}-reroute-connectors`}
                    onSelect={() => adapter.rerouteConnectors?.onSelect()}
                  >
                    <OfficeSemanticIcon semanticId="connector.reroute" size={16} />
                    {labels.rerouteConnectors}
                  </DropdownMenuItem>
                ) : null}
              </DropdownMenuContent>
            </DropdownMenu>
            {(adapter.textBox.directions ?? ["horizontal", "rotate90"]).length === 1 ? (
              <OfficeHomeCommand
                controlId={textBox.controlId}
                disabled={textBox.disabled}
                disabledReason={textBox.disabledReason}
                icon={<OfficeSemanticIcon semanticId="textBox.insert" size={24} />}
                label={labels.textBox}
                onClick={() =>
                  adapter.textBox.onInsert(
                    (adapter.textBox.directions ?? ["horizontal"])[0] ?? "horizontal",
                  )
                }
                testId={id("textBox", "text-box")}
                variant="compact"
              />
            ) : (
            <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    controlId={textBox.controlId}
                    disabled={textBox.disabled}
                    disabledReason={textBox.disabledReason}
                    hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="textBox.insert" size={24} />}
                    label={labels.textBox}
                    testId={id("textBox", "text-box")}
                    variant="compact"
                  />
                </DropdownMenuTrigger>
                <DropdownMenuContent {...menuContentProps}>
                  <DropdownMenuItem
                    data-testid={`${id("textBox", "text-box")}-horizontal`}
                    onSelect={() => adapter.textBox.onInsert("horizontal")}
                  >
                    <OfficeSemanticIcon semanticId="textBox.insert" size={16} />
                    {labels.textBoxHorizontal}
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    data-testid={`${id("textBox", "text-box")}-vertical`}
                    onSelect={() => adapter.textBox.onInsert("rotate90")}
                  >
                    <OfficeSemanticIcon semanticId="textBox.insert" size={16} />
                    {labels.textBoxVertical}
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            )}
            {adapter.mergeShapes ? (
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    controlId={adapter.mergeShapes.controlId}
                    disabled={Boolean(adapter.mergeShapes.disabled)}
                    disabledReason={
                      adapter.mergeShapes.disabled
                        ? adapter.mergeShapes.disabledReason
                        : undefined
                    }
                    hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="shape.change" size={24} />}
                    label={adapter.mergeShapes.label}
                    onPointerDown={adapter.onRememberSelection}
                    testId={`${testIdPrefix}-merge`}
                    variant="compact"
                  />
                </DropdownMenuTrigger>
                <DropdownMenuContent {...menuContentProps}>
                  {adapter.mergeShapes.items.map((item) => (
                    <DropdownMenuItem
                      key={item.key}
                      data-testid={item.testId ?? `${testIdPrefix}-merge-${item.key}`}
                      onSelect={item.onSelect}
                    >
                      {item.label}
                    </DropdownMenuItem>
                  ))}
                </DropdownMenuContent>
              </DropdownMenu>
            ) : null}
          </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      )}

      {/* ---- Shape Styles --------------------------------------------- */}
      <OfficeShapeFormatStyleGroup
        testId={id("stylesGroup", "styles-group")}
        application={adapter.application}
        recipes={adapter.styles.recipes}
        quickCount={adapter.styles.quickCount ?? 3}
        quickRow={adapter.styles.quickRow}
        pagingLabels={{ previous: labels.previousStyles, next: labels.nextStyles }}
        label={labels.shapeStyles}
        moreLabel={labels.moreShapeStyles}
        labelFor={adapter.styles.labelFor}
        value={adapter.styles.value}
        disabled={stylesCap.disabled}
        getDisabledReason={adapter.styles.getDisabledReason}
        onCommit={adapter.styles.onCommit}
        testIdPrefix={id("styles", "styles")}
        commands={
          <OfficeHomeCommandStack className={styles.paintCommands} rows={3}>
            <OfficeHomeCommandRow>
              <ColorSplit
                adapter={adapter.fill}
                icon={<OfficeSemanticIcon semanticId="shape.fill" size={20} />}
                label={labels.shapeFill}
                displayLabel={labels.fill}
                ids={{
                  split: id("fillSplit", "fill-split"),
                  primary: id("fillPrimary", "fill"),
                  trigger: id("fillTrigger", "fill-options"),
                  palette: id("fill", "fill"),
                }}
                labels={labels}
                moreColors={labels.moreFillColors}
                noFill={labels.noFill}
                menuClassName={menuClassName}
                {...paletteProps("fill")}
              />
            </OfficeHomeCommandRow>
            <OfficeHomeCommandRow>
              <ColorSplit
                adapter={adapter.outline}
                icon={<OfficeSemanticIcon semanticId="shape.outline" size={20} />}
                label={labels.shapeOutline}
                displayLabel={labels.outline}
                ids={{
                  split: id("outlineSplit", "outline-split"),
                  primary: id("outlinePrimary", "outline"),
                  trigger: id("outlineTrigger", "outline-options"),
                  palette: id("outline", "outline"),
                }}
                labels={labels}
                moreColors={labels.moreOutlineColors}
                noFill={labels.noOutline}
                menuClassName={menuClassName}
                outline={{
                  adapter: adapter.outline,
                  ids: {
                    weight: id("outlineWeight", "outline-weight"),
                    dashes: id("outlineDashes", "outline-dashes"),
                    arrows: id("arrows", "arrows"),
                    lineEnds: id("lineEnds", "line-ends"),
                  },
                }}
                {...paletteProps("outline")}
              />
            </OfficeHomeCommandRow>
            <OfficeHomeCommandRow>
              <EffectsMenu
                adapter={adapter.effects}
                label={labels.shapeEffects}
                displayLabel={labels.effects}
                labels={labels}
                testId={id("effects", "effects")}
                itemTestIdPrefix={id("effect", "effect")}
                menuClassName={menuClassName}
              />
            </OfficeHomeCommandRow>
          </OfficeHomeCommandStack>
        }
      />

      {/* ---- WordArt Styles ------------------------------------------- */}
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.wordArtStyles}
        className={styles.wordArtGroup}
        testId={id("wordArtGroup", "text-styles-group")}
      >
        <div className={styles.wordArtBody}>
          <OfficeWordArtStyleGallery
            application={adapter.application}
            recipes={adapter.wordArt.recipes}
            labelFor={adapter.wordArt.labelFor}
            sample={adapter.wordArt.sample}
            value={adapter.wordArt.value}
            disabled={wordArt.disabled}
            disabledReason={wordArt.disabledReason}
            labels={{ title: labels.wordArtStyles, more: labels.moreWordArtStyles }}
            onCommit={adapter.wordArt.onCommit}
            testIdPrefix={id("wordArt", "text-styles")}
          />
          <OfficeHomeCommandStack className={styles.paintCommands} rows={3}>
            <OfficeHomeCommandRow>
              <ColorSplit
                adapter={adapter.textFill}
                icon={<OfficeSemanticIcon semanticId="text.fontColor" size={18} />}
                label={labels.textFill}
                displayLabel={labels.textFill}
                ids={{
                  split: `${id("textFill", "text-fill")}-split`,
                  primary: id("textFill", "text-fill"),
                  trigger: `${id("textFill", "text-fill")}-options`,
                  palette: id("textFill", "text-fill"),
                }}
                labels={labels}
                moreColors={labels.moreTextColors}
                noFill={labels.noFill}
                menuClassName={menuClassName}
                {...paletteProps("textFill")}
              />
            </OfficeHomeCommandRow>
            <OfficeHomeCommandRow>
              <ColorSplit
                adapter={adapter.textOutline}
                icon={<OfficeSemanticIcon semanticId="format.lineColor" size={18} />}
                label={labels.textOutline}
                displayLabel={labels.textOutline}
                ids={{
                  split: `${id("textOutline", "text-outline")}-split`,
                  primary: id("textOutline", "text-outline"),
                  trigger: `${id("textOutline", "text-outline")}-options`,
                  palette: id("textOutline", "text-outline"),
                }}
                labels={labels}
                moreColors={labels.moreOutlineColors}
                noFill={labels.noOutline}
                menuClassName={menuClassName}
                {...paletteProps("textOutline")}
              />
            </OfficeHomeCommandRow>
            <OfficeHomeCommandRow>
              <EffectsMenu
                adapter={adapter.textEffects}
                label={labels.textEffects}
                displayLabel={labels.textEffects}
                labels={labels}
                testId={id("textEffects", "text-effects")}
                itemTestIdPrefix={id("textEffect", "text-effect")}
                menuClassName={menuClassName}
                extra={
                  adapter.textTransform ? (
                    <OfficeTextTransformSubmenu
                      value={adapter.textTransform.value}
                      onPick={adapter.textTransform.onPick}
                      labels={adapter.textTransform.labels}
                      disabled={adapter.textTransform.disabled}
                      disabledReason={adapter.textTransform.disabledReason}
                      testId={`${id("textEffects", "text-effects")}-transform`}
                      menuClassName={menuClassName}
                    />
                  ) : null
                }
              />
            </OfficeHomeCommandRow>
          </OfficeHomeCommandStack>
        </div>
      </OfficeHomeRibbonGroup>

      {/* ---- Text (Word only) ----------------------------------------- */}
      {text ? (
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.text}
        className={styles.textGroup}
        testId={id("textGroup", "text-group")}
      >
        <OfficeHomeCommandStack rows={text.createLink ? 3 : undefined}>
          <DropdownMenu modal={false}>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                active={text.direction.value !== "horizontal"}
                controlId={textDirection.controlId}
                disabled={textDirection.disabled}
                disabledReason={textDirection.disabledReason}
                hasMenuIndicator
                icon={
                  <OfficeSemanticIcon semanticId="table.cell.textDirection" size={18} />
                }
                label={labels.textDirection}
                testId={id("textDirection", "text-direction")}
                variant="compact"
              />
            </DropdownMenuTrigger>
            <DropdownMenuContent {...menuContentProps}>
              <DropdownMenuRadioGroup value={text.direction.value}>
                {(["horizontal", "rotate90", "rotate270", "stacked"] as const).map(
                  (direction) => (
                    <DropdownMenuRadioItem
                      key={direction}
                      data-testid={`${id("textDirection", "text-direction")}-${direction}`}
                      onSelect={() => text.direction.onPick(direction)}
                      value={direction}
                    >
                      {labels.textDirectionOption(direction)}
                    </DropdownMenuRadioItem>
                  ),
                )}
              </DropdownMenuRadioGroup>
            </DropdownMenuContent>
          </DropdownMenu>
          <DropdownMenu modal={false}>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                active={text.align.value !== "middle"}
                controlId={textAlign.controlId}
                disabled={textAlign.disabled}
                disabledReason={textAlign.disabledReason}
                hasMenuIndicator
                icon={<OfficeSemanticIcon semanticId="object.align" size={18} />}
                label={labels.alignText}
                testId={id("textAlign", "text-align")}
                variant="compact"
              />
            </DropdownMenuTrigger>
            <DropdownMenuContent {...menuContentProps}>
              <DropdownMenuRadioGroup value={text.align.value}>
                {(["top", "middle", "bottom"] as const).map((align) => (
                  <DropdownMenuRadioItem
                    key={align}
                    data-testid={`${id("textAlign", "text-align")}-${align}`}
                    onSelect={() => text.align.onPick(align)}
                    value={align}
                  >
                    {labels.alignTextOption(align)}
                  </DropdownMenuRadioItem>
                ))}
              </DropdownMenuRadioGroup>
            </DropdownMenuContent>
          </DropdownMenu>
          {text.createLink ? (
            <OfficeHomeCommand
              active={text.createLink.active}
              controlId={text.createLink.controlId}
              disabled={Boolean(text.createLink.disabled)}
              disabledReason={
                text.createLink.disabled
                  ? text.createLink.disabledReason
                  : undefined
              }
              displayLabel={labels.createLink}
              icon={<OfficeSemanticIcon semanticId="link.insert" size={18} />}
              label={labels.createLink}
              onClick={text.createLink.onOpen}
              testId={id("createLink", "create-link")}
              variant="compact"
            />
          ) : null}
        </OfficeHomeCommandStack>
      </OfficeHomeRibbonGroup>

      ) : null}

      {/* ---- Accessibility -------------------------------------------- */}
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.accessibility}
        className={styles.altTextGroup}
        testId={id("altTextGroup", "alt-text-group")}
      >
        <OfficeHomeCommand
          active={adapter.altText.active}
          controlId={altText.controlId}
          disabled={altText.disabled}
          disabledReason={altText.disabledReason}
          icon={<OfficeSemanticIcon semanticId="object.altText" size={24} />}
          label={labels.altText}
          onClick={adapter.altText.onOpen}
          testId={id("altText", "alt-text")}
          variant="large"
        />
      </OfficeHomeRibbonGroup>

      {/* ---- Arrange -------------------------------------------------- */}
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.arrange}
        className={styles.arrangeGroup}
        testId={id("arrangeGroup", "arrange-group")}
      >
        <OfficeHomeCommandRow>
          {adapter.selectionCount ? (
            <span
              className={styles.selectionCount}
              data-testid={id("selectionCount", "selection-count")}
              data-count={adapter.selectionCount.count}
              aria-label={adapter.selectionCount.label}
            >
              {adapter.selectionCount.label}
            </span>
          ) : null}
          {adapter.position ? (
            <DropdownMenu modal={false}>
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  {...reasonProps(adapter.position)}
                  hasMenuIndicator
                  icon={adapter.position.icon}
                  label={labels.position}
                  testId={id("position", "position")}
                  variant="large"
                />
              </DropdownMenuTrigger>
              <DropdownMenuContent {...menuContentProps}>{adapter.position.menu}</DropdownMenuContent>
            </DropdownMenu>
          ) : null}
          {adapter.wrapText ? (
            <DropdownMenu modal={false}>
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  {...reasonProps(adapter.wrapText)}
                  hasMenuIndicator
                  icon={adapter.wrapText.icon}
                  label={labels.wrapText}
                  testId={id("wrapText", "wrap-text")}
                  variant="large"
                />
              </DropdownMenuTrigger>
              <DropdownMenuContent {...menuContentProps}>{adapter.wrapText.menu}</DropdownMenuContent>
            </DropdownMenu>
          ) : null}
          <OfficeHomeCommandStack rows={3}>
            {(
              [
                ["layer", "bring-forward", "bringForward", "bringToFront", "object.bringForward"],
                ["send", "send-backward", "sendBackward", "sendToBack", "object.sendBackward"],
              ] as const
            ).map(([key, fallback, primaryAction, menuAction, iconId]) => {
              const primaryLabel =
                primaryAction === "bringForward" ? labels.bringForward : labels.sendBackward;
              const menuLabel =
                menuAction === "bringToFront" ? labels.bringToFront : labels.sendToBack;
              const splitId = id(`${key}Split` as "layerSplit", `${fallback}-split`);
              const primaryControlId =
                adapter.zOrder.controlIdFor?.(primaryAction) ?? zOrder.controlId;
              const primaryDisabled =
                zOrder.disabled || Boolean(adapter.zOrder.atBoundary?.(primaryAction));
              const atEndReason =
                key === "layer" ? labels.alreadyAtFront : labels.alreadyAtBack;
              const primaryReason = primaryDisabled
                ? zOrder.disabled
                  ? zOrder.disabledReason
                  : (adapter.zOrder.disabledReasonFor?.(primaryAction) ?? atEndReason)
                : undefined;
              const menuId = id(`${key}Menu` as "layerMenu", `${fallback}-menu`);
              // The menu holds this direction's one-step and all-the-way
              // moves. When both are refused (the object is already at that
              // end of the stack) the chevron greys with the face and says
              // why; a live chevron over a greyed face opened a menu of two
              // greyed items.
              const textLayer = adapter.zOrder.textLayer;
              const menuDisabled =
                zOrder.disabled ||
                (Boolean(
                  adapter.zOrder.atBoundary?.(primaryAction) &&
                    adapter.zOrder.atBoundary?.(menuAction),
                ) &&
                  (!textLayer || textLayer.disabled === true));
              const menuReason = zOrder.disabled
                ? zOrder.disabledReason
                : menuDisabled
                  ? (adapter.zOrder.disabledReasonFor?.(primaryAction) ?? atEndReason)
                  : undefined;
              return (
                <OfficeHomeSplitLayout
                  key={key}
                  className={styles.arrangeSplit}
                  testId={splitId}
                  primary={
                    <OfficeHomeCommand
                      variant="compact"
                      controlId={primaryControlId}
                      disabled={primaryDisabled}
                      disabledReason={primaryReason}
                      icon={<OfficeSemanticIcon semanticId={iconId} size={18} />}
                      displayLabel={primaryLabel}
                      label={primaryLabel}
                      onClick={() => adapter.zOrder.onAction(primaryAction)}
                      testId={id(key, fallback)}
                    />
                  }
                  trigger={
                    <DropdownMenu modal={false}>
                      <DropdownMenuTrigger asChild>
                        <OfficeHomeMenuTrigger
                          variant="icon"
                          icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                          controlId={primaryControlId}
                          disabled={menuDisabled}
                          disabledReason={menuReason}
                          label={primaryLabel}
                          testId={menuId}
                        />
                      </DropdownMenuTrigger>
                      <DropdownMenuContent {...menuContentProps}>
                        <DropdownMenuItem
                          data-testid={
                            ids[`${primaryAction}Item` as "bringForwardItem"] ??
                            `${menuId}-${primaryAction}`
                          }
                          disabled={adapter.zOrder.atBoundary?.(primaryAction)}
                          onSelect={() => adapter.zOrder.onAction(primaryAction)}
                        >
                          {primaryLabel}
                        </DropdownMenuItem>
                        <DropdownMenuItem
                          data-testid={
                            ids[`${menuAction}Item` as "bringToFrontItem"] ??
                            `${menuId}-${menuAction}`
                          }
                          disabled={adapter.zOrder.atBoundary?.(menuAction)}
                          onSelect={() => adapter.zOrder.onAction(menuAction)}
                        >
                          {menuLabel}
                        </DropdownMenuItem>
                        {textLayer ? (
                          <DropdownMenuItem
                            data-testid={`${menuId}-${key === "layer" ? "inFrontOfText" : "behindText"}`}
                            disabled={textLayer.disabled}
                            onSelect={
                              key === "layer" ? textLayer.onInFrontOfText : textLayer.onBehindText
                            }
                          >
                            {key === "layer"
                              ? textLayer.inFrontOfTextLabel
                              : textLayer.behindTextLabel}
                          </DropdownMenuItem>
                        ) : null}
                      </DropdownMenuContent>
                    </DropdownMenu>
                  }
                />
              );
            })}
            <OfficeHomeCommand
              controlId={selectionPane.controlId}
              disabled={selectionPane.disabled}
              disabledReason={selectionPane.disabledReason}
              displayLabel={labels.selectionPane}
              icon={<OfficeSemanticIcon semanticId="object.selectionPane" size={18} />}
              label={labels.selectionPane}
              onClick={adapter.selectionPane.onSelect}
              testId={id("selectionPane", "selection-pane")}
              variant="compact"
            />
          </OfficeHomeCommandStack>
          {adapter.reorder ? (
            <OfficeHomeCommandStack rows={3}>
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    {...reasonProps(adapter.reorder)}
                    hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="object.selectionPane" size={18} />}
                    label={adapter.reorder.label}
                    testId={id("reorder", "reorder")}
                    variant="compact"
                  />
                </DropdownMenuTrigger>
                <DropdownMenuContent {...menuContentProps}>
                  <DropdownMenuItem
                    data-testid={id("reorderItem", "reorder-objects")}
                    onSelect={adapter.reorder.onReorder}
                  >
                    {adapter.reorder.label}
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    data-testid={id("reorderOverlappingItem", "reorder-overlapping")}
                    disabled={adapter.reorder.overlappingDisabled}
                    onSelect={adapter.reorder.onReorderOverlapping}
                  >
                    {adapter.reorder.overlappingLabel}
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </OfficeHomeCommandStack>
          ) : null}
          <OfficeHomeCommandStack rows={3}>
            <DropdownMenu modal={false}>
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  controlId={alignObjects.controlId}
                  disabled={alignObjects.disabled}
                  disabledReason={alignObjects.disabledReason}
                  hasMenuIndicator
                  icon={<OfficeSemanticIcon semanticId="object.align" size={18} />}
                  label={labels.align}
                  testId={id("align", "align")}
                  variant="compact"
                />
              </DropdownMenuTrigger>
              <DropdownMenuContent {...menuContentProps} data-testid={id("alignMenu", "align-menu")}>
                {adapter.alignObjects.sections ? (
                  <OfficeSectionedAlignMenuItems
                    sections={adapter.alignObjects.sections}
                    label={labels.alignOption}
                    isDisabled={adapter.alignObjects.isActionDisabled}
                    onAction={adapter.alignObjects.onAction}
                    target={adapter.alignObjects.target}
                    disabled={alignObjects.disabled}
                    testIdFor={(action) =>
                      ids[
                        `align${action[0]!.toUpperCase()}${action.slice(1)}` as "alignLeft"
                      ] ?? `${id("align", "align")}-${action}`
                    }
                  />
                ) : null}
                {!adapter.alignObjects.sections && adapter.alignObjects.alignToTarget ? (
                  <>
                    <DropdownMenuCheckboxItem
                      checked={adapter.alignObjects.alignToTarget.checked}
                      data-testid={id("alignToTarget", "align-to-target")}
                      disabled={alignObjects.disabled}
                      onCheckedChange={adapter.alignObjects.alignToTarget.onChange}
                    >
                      {adapter.alignObjects.alignToTarget.label}
                    </DropdownMenuCheckboxItem>
                    <DropdownMenuSeparator />
                  </>
                ) : null}
                {(adapter.alignObjects.sections ? [] : alignActions).map((action, index) => (
                  <Fragment key={action}>
                    {index > 0 &&
                    action.startsWith("distribute") &&
                    !alignActions[index - 1]!.startsWith("distribute") ? (
                      <DropdownMenuSeparator />
                    ) : null}
                    <DropdownMenuItem
                      data-testid={
                        ids[
                          `align${action[0]!.toUpperCase()}${action.slice(1)}` as "alignLeft"
                        ] ?? `${id("align", "align")}-${action}`
                      }
                      disabled={adapter.alignObjects.isActionDisabled?.(action)}
                      onSelect={() => adapter.alignObjects.onAction(action)}
                    >
                      {labels.alignOption(action)}
                    </DropdownMenuItem>
                  </Fragment>
                ))}
                {adapter.alignObjects.toggles?.length ? <DropdownMenuSeparator /> : null}
                {adapter.alignObjects.toggles?.map((toggle) => (
                  <DropdownMenuCheckboxItem
                    key={toggle.key}
                    checked={toggle.checked}
                    data-testid={toggle.testId ?? `${id("align", "align")}-${toggle.key}`}
                    onCheckedChange={toggle.onChange}
                  >
                    {toggle.label}
                  </DropdownMenuCheckboxItem>
                ))}
              </DropdownMenuContent>
            </DropdownMenu>
            {adapter.group.extra?.length || adapter.group.regroup ? (
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    active={adapter.group.grouped}
                    controlId={group.controlId}
                    disabled={groupMenuDisabled}
                    disabledReason={groupMenuDisabled ? group.disabledReason : undefined}
                    hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="object.group" size={18} />}
                    displayLabel={labels.group}
                    label={labels.group}
                    testId={id("group", "group")}
                    variant="compact"
                  />
                </DropdownMenuTrigger>
                <DropdownMenuContent {...menuContentProps}>
                  <DropdownMenuItem
                    data-testid={id("groupItem", "group-group")}
                    disabled={adapter.group.canGroup === false}
                    onSelect={adapter.group.onGroup}
                  >
                    <OfficeSemanticIcon semanticId="object.group" size={16} />
                    {labels.group}
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    data-testid={id("ungroupItem", "group-ungroup")}
                    disabled={adapter.group.canUngroup === false}
                    onSelect={adapter.group.onUngroup}
                  >
                    <OfficeSemanticIcon semanticId="object.ungroup" size={16} />
                    {labels.ungroup}
                  </DropdownMenuItem>
                  {adapter.group.regroup ? (
                    <DropdownMenuItem
                      data-testid={id("regroupItem", "group-regroup")}
                      disabled={adapter.group.regroup.disabled}
                      data-disabled-reason={
                        adapter.group.regroup.disabled
                          ? adapter.group.regroup.disabledReason
                          : undefined
                      }
                      onSelect={adapter.group.regroup.onSelect}
                    >
                      <OfficeSemanticIcon semanticId="object.group" size={16} />
                      {adapter.group.regroup.label}
                    </DropdownMenuItem>
                  ) : null}
                  {adapter.group.extra?.length ? <DropdownMenuSeparator /> : null}
                  {(adapter.group.extra ?? []).map((item) => (
                    <DropdownMenuItem
                      key={item.key}
                      data-testid={item.testId ?? `${id("group", "group")}-${item.key}`}
                      disabled={item.disabled}
                      data-disabled-reason={item.disabled ? item.disabledReason : undefined}
                      onSelect={item.onSelect}
                    >
                      {item.icon}
                      {item.label}
                    </DropdownMenuItem>
                  ))}
                </DropdownMenuContent>
              </DropdownMenu>
            ) : (
              <OfficeHomeCommand
                active={adapter.group.grouped}
                controlId={group.controlId}
                disabled={group.disabled}
                disabledReason={group.disabledReason}
                displayLabel={adapter.group.grouped ? labels.ungroup : labels.group}
                icon={<OfficeSemanticIcon semanticId="object.group" size={18} />}
                label={adapter.group.grouped ? labels.ungroup : labels.group}
                onClick={adapter.group.grouped ? adapter.group.onUngroup : adapter.group.onGroup}
                testId={id("group", "group")}
                variant="compact"
              />
            )}
            <DropdownMenu modal={false}>
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  controlId={rotate.controlId}
                  disabled={rotate.disabled}
                  disabledReason={rotate.disabledReason}
                  hasMenuIndicator
                  icon={<OfficeSemanticIcon semanticId="object.rotate" size={18} />}
                  label={labels.rotate}
                  testId={id("rotate", "rotate")}
                  variant="compact"
                />
              </DropdownMenuTrigger>
              <DropdownMenuContent {...menuContentProps}>
                {(
                  ["rotateRight90", "rotateLeft90", "flipVertical", "flipHorizontal"] as const
                ).map((action) => (
                  <DropdownMenuItem
                    key={action}
                    data-testid={ids[action] ?? `${id("rotate", "rotate")}-${action}`}
                    onSelect={() => adapter.rotate.onAction(action)}
                  >
                    {labels.rotateOption(action)}
                  </DropdownMenuItem>
                ))}
                {adapter.rotate.exactField ? (
                  <>
                    <DropdownMenuSeparator />
                    {adapter.rotate.exactField}
                  </>
                ) : null}
                {adapter.rotate.reset ? (
                  <>
                    <DropdownMenuSeparator />
                    <DropdownMenuItem
                      data-testid={id("rotateReset", "rotate-reset")}
                      disabled={adapter.rotate.reset.disabled}
                      onSelect={adapter.rotate.reset.onSelect}
                    >
                      {labels.resetRotation ?? labels.moreRotationOptions}
                    </DropdownMenuItem>
                  </>
                ) : null}
                {adapter.rotate.exact ? (
                  <>
                    <DropdownMenuSeparator />
                    <DropdownMenuItem
                      data-testid={id("rotateExact", "rotate-exact")}
                      disabled={adapter.rotate.exact.disabled}
                      onSelect={adapter.rotate.exact.onSelect}
                    >
                      {labels.moreRotationOptions}
                    </DropdownMenuItem>
                  </>
                ) : null}
              </DropdownMenuContent>
            </DropdownMenu>
          </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>

      {/* ---- Size ----------------------------------------------------- */}
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={labels.size}
        className={styles.sizeGroup}
        size="standard"
        testId={id("sizeGroup", "size-group")}
      >
        <OfficeHomeCommandRow className={styles.sizeRow} testId={id("sizeRow", "size-row")}>
          <OfficeShapeSizeFields
            heightFirst={adapter.size.heightFirst}
            values={{ width: adapter.size.width, height: adapter.size.height }}
            mixed={adapter.size.mixed}
            mixedLabel={labels.mixed}
            labels={{ width: labels.width, height: labels.height }}
            unitLabel={labels.sizeUnit}
            step={adapter.size.step}
            disabled={size.disabled}
            disabledReason={size.disabledReason}
            max={adapter.size.max}
            onCommit={adapter.size.onCommit}
            onRemember={adapter.onRememberSelection}
            testIdPrefix={id("size", "size")}
          />
          {/* Beside the fields, not under them: the group is a column, so a
              stack left outside this row drew a third (Docs: a fourth) row
              that pushed the fields over the band's top edge and the caption
              out of the band. */}
          <OfficeHomeCommandStack>
            <OfficeHomeCommand
              active={adapter.size.lockAspect.value}
              controlId={lockAspect.controlId}
              disabled={lockAspect.disabled}
              disabledReason={lockAspect.disabledReason}
              displayLabel={labels.lockAspectRatio}
              icon={<OfficeSemanticIcon semanticId="object.lockAspectRatio" size={18} />}
              label={labels.lockAspectRatio}
              onClick={adapter.size.lockAspect.onToggle}
              testId={id("lockAspect", "lock-aspect")}
              variant="compact"
            />
            {adapter.formatPane ? (
              <OfficeHomeCommand
                controlId={adapter.formatPane.controlId}
                disabled={Boolean(adapter.formatPane.disabled)}
                disabledReason={
                  adapter.formatPane.disabled ? adapter.formatPane.disabledReason : undefined
                }
                displayLabel={labels.formatPane}
                icon={
                  adapter.formatPane.icon ?? (
                    <OfficeSemanticIcon semanticId="object.selectionPane" size={18} />
                  )
                }
                label={labels.formatPane}
                onClick={adapter.formatPane.onSelect}
                testId={id("formatPane", "format-pane")}
                variant="compact"
              />
            ) : null}
          </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
    </>
  );
}
