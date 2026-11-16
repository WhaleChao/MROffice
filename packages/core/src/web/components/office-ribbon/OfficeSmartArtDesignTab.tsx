"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The SmartArt Design tab — ONE contextual Ribbon tab, mounted by every host.
 *
 * PowerPoint shows this tab whenever a SmartArt graphic is selected, and MRAI's
 * three editors show the SAME tab: the layouts gallery (the 39 catalog
 * layouts), the fourteen-style gallery, Change Colors, the structural commands
 * (Text Pane, Add Shape, Promote, Demote, Right to Left), Reset Graphic and Alt
 * Text. Nothing here knows about slides, documents or sheets.
 *
 * ## Mounting it
 *
 * A host renders `<OfficeSmartArtDesignTab …/>` where its contextual tab's
 * groups go, and supplies:
 *
 * - `testIdPrefix` — `"presentation"`, `"document"` or `"spreadsheet"`. Every
 *   test id is `<prefix>-ribbon-smartart-…` / `<prefix>-smartart-…`, so the
 *   three hosts' suites address the same controls under their own names.
 * - `labels` — every visible string, already translated. This component calls
 *   no `t()`, so a host keeps its own i18n namespace.
 * - `diagram` — the selected AUTHORED diagram's state (layout `kind`, quick
 *   style `styleId`, its `members` so picture layouts can be blocked, whether
 *   it is writable and what its outline currently allows) plus one callback per
 *   command. Each callback performs the host's own transaction, normally by
 *   handing the diagram to `lib/office-smartart/design-ops` and writing the
 *   returned members back in ONE undo step. Pass `null` when the selection is
 *   not an authored diagram.
 * - `preserved` — the native (not-yet-converted) diagram's state: whether to
 *   offer "convert to editable" (the shared layout gallery picks the catalog
 *   layout it becomes) and/or the in-place native text commands. Pass `null`
 *   when the selection is not a native diagram.
 * - `placeholders` / `previewBounds` — the words and frame the host's insert
 *   seeds, so a Change Layout thumbnail is the scene that layout would build,
 *   drawn by the ONE shared seed preview every editor's gallery uses.
 * - `renderStyleShape` — an optional host renderer for the style tiles; it
 *   falls back to the shared SVG drawing, so a host needs none.
 *
 * ## What it never does
 *
 * It holds no state, runs no effect and owns no transaction: every enable /
 * disable decision arrives as a boolean and every command leaves through a
 * callback. A host that wires the callbacks to `design-ops` gets exactly the
 * behaviour Presentation ships.
 */

import {
  ArrowDown,
  ArrowUp,
  List,
  Maximize2,
  Minimize2,
  Plus,
  Shapes,
} from "./icons/houseIcons";
import type { ReactNode, RefObject } from "react";

import { RibbonDropdownMenuContent } from "./CupertinoRibbonMenu";
import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
  OfficeHomeSplitLayout,
} from "./OfficeHomeRibbon";
import {
  OfficeGalleryMenuContent,
  OfficeGalleryMenuGrid,
  OfficeGalleryMenuSection,
  OfficeGalleryMenuTileBody,
  officeGalleryMenuStyles,
} from "./OfficeGalleryMenu";
import {
  OfficeStyleGallery,
  type OfficeStyleGalleryMenuGrid,
} from "./OfficeStyleGallery";
import { OfficeSmartArtStylePreview } from "./OfficeSmartArtStylePreview";
import { OfficeSmartArtSeedPreview } from "./OfficeSmartArtSeedPreview";
import {
  parseSmartArtGalleryId,
  SMARTART_ALIAS_GALLERY,
  smartArtGalleryId,
} from "../../lib/office-smartart/native-layouts";
import { OfficeSemanticIcon } from "./icons/OfficeSemanticIcon";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "../ui/dropdown-menu";
import {
  OFFICE_SMARTART_CONVERT_ONLY_LAYOUTS,
  OFFICE_SMARTART_CATEGORY_IDS,
  OFFICE_SMARTART_LAYOUTS,
  officeSmartArtGalleryTabs,
  type OfficeSmartArtCategoryId,
  type OfficeSmartArtLayout,
  type OfficeSmartArtPlaceholders,
} from "../../lib/office-smartart/catalog";
import {
  OFFICE_SMARTART_COLOR_GROUPS,
  resolveSmartArtColorVariant,
  smartArtColorVariantSwatch,
  type SmartArtColorGroup,
  type SmartArtColorGroupId,
  type SmartArtColorTheme,
  type SmartArtColorVariant,
} from "../../lib/office-smartart/colors";
import {
  OFFICE_SMARTART_STYLE_SECTIONS,
  readSmartArtGroupLayoutUrn,
  SMARTART_NO_OUTLINE_COMMANDS,
  smartArtDiagramItemCount,
  smartArtLayoutHasAssistant,
  type OfficeSmartArtStyleSectionId,
  type SmartArtOutlineCommandState,
  type SmartArtOutlineInsertPosition,
} from "../../lib/office-smartart/design-model";
import {
  smartArtLayoutTakesHierBranch,
  type OfficeSmartArtBlockLike,
  type OfficeSmartArtShapeNode,
  type SmartArtHierBranch,
  type SmartArtKind,
  type SmartArtNodePaint,
  type SmartArtStyle,
} from "../../lib/office-smartart/model";
import { type SmartArtPictureSlot } from "../../lib/office-smartart/picture-command";
import {
  SMARTART_STYLES,
  type SmartArtStyleId,
} from "../../lib/office-smartart/styles-model";

import styles from "./OfficeSmartArtDesignTab.module.css";

/**
 * The More sheets of Layouts and SmartArt Styles: five captioned drawings a
 * row, each drawn into this box (px), as PowerPoint's sheets are -- the sheet
 * is as wide as those five tiles, not a fixed panel.
 */
const SMARTART_MENU_THUMB = [76, 56] as const;
const SMARTART_MENU_GRID: OfficeStyleGalleryMenuGrid = {
  columns: 5,
  thumbnail: SMARTART_MENU_THUMB,
};
/** Change Colors: five variation swatches a row (a swatch is up to five bars). */
const SMARTART_COLOR_COLUMNS = 5;
const SMARTART_COLOR_TILE_PX = 92;

/**
 * The tab's style slots. A host that already owns this geometry passes its own
 * class per slot (Presentation does — its copy is pinned by
 * `tests/unit/presentation-ribbon-more-css.test.ts`); anything it omits falls
 * back to {@link ./OfficeSmartArtDesignTab.module.css}, so a new host needs no
 * CSS at all.
 */
export interface OfficeSmartArtDesignTabClasses {
  createGroup: string;
  createRow: string;
  createHierarchyLane: string;
  layoutsGroup: string;
  layoutGallery: string;
  stylesGroup: string;
  styleGallery: string;
  resetGroup: string;
  propertiesGroup: string;
}

/** Every visible string, already translated by the host. */
export interface OfficeSmartArtDesignTabLabels {
  /** Group label over the structural commands ("Create Graphic"). */
  createGraphicGroup: string;
  /** The preserved diagram's one command, and its gallery's "more" label. */
  convertToEditable: string;
  /** Group label over the layouts gallery ("Layouts"). */
  layoutsGroupLabel: string;
  /** Accessible name of the layouts gallery itself. */
  layoutsGroup: string;
  /** Takes the catalog entry, not its kind: two layouts may share a kind. */
  layoutName: (layout: OfficeSmartArtLayout) => string;
  /**
   * An aligned Office ALIAS's own name (`hProcess9` -> "Continuous Block
   * Process"), listed beside the catalog layout it stands for. Absent, the
   * gallery lists the catalog layouts only.
   */
  aliasLayoutName?: (suffix: string) => string;
  /** Section caption of the Layouts sheet ("List", "Process", ...). Absent,
   * the sheet is one ungrouped grid. */
  layoutCategoryName?: (category: OfficeSmartArtCategoryId) => string;
  /** Why a Picture layout refuses to become a non-picture one. */
  pictureLayoutPreserveReason: string;
  /**
   * Why the diagram's CURRENT layout is shown in the gallery but cannot be
   * picked: it is a convert-only layout (SMARTART-IMPORT.md §5.34, §5.39) —
   * Office's own definition, kept from the file, which MRAI can draw and keep
   * but cannot create from nothing.
   *
   * Optional, so a host that has not translated it yet still gets the marker;
   * without it the tile carries its layout name and no explanation.
   */
  currentLayoutUnavailable?: string;
  /** A gallery's overflow-menu label. */
  quickOverflow: string;
  /**
   * The picture-slot command, shown only for a diagram that HAS slots. Absent
   * in a host that does not offer it (the presentation's SmartArt member is an
   * image block of its own and has "Replace picture" already).
   */
  insertPicture?: string;
  /** One slot in that command's menu — its 1-based ordinal and whether it is filled. */
  pictureSlotName?: (slot: SmartArtPictureSlot) => string;
  /** Why a slot Office's own layout draws no picture for cannot be filled. */
  pictureSlotUnexportable?: string;
  textPane: string;
  addShape: string;
  promote: string;
  demote: string;
  rightToLeft: string;
  stylesGroup: string;
  changeColors: string;
  /** One Change Colors variation's name (Office's tooltip for it). */
  colorName: (variant: SmartArtColorVariant) => string;
  /** A Change Colors heading: Primary Theme Colors, Colorful, Accent N. */
  colorGroupName: (group: {
    readonly id: SmartArtColorGroupId;
    readonly accent?: number;
  }) => string;
  styleName: (styleId: SmartArtStyleId) => string;
  styleSection: (sectionId: OfficeSmartArtStyleSectionId) => string;
  /** Group label over Reset Graphic ("Reset"). */
  resetGroupLabel: string;
  resetGraphic: string;
  propertiesGroup: string;
  altText: string;
  /**
   * Add Shape's positions (Microsoft `MsoSmartArtNodePosition`), Move Up / Move
   * Down, and Reset > Convert. Optional so an older caller keeps compiling; the
   * shared `useOfficeSmartArtDesignLabels` supplies every one.
   */
  addShapeAfter?: string;
  addShapeBefore?: string;
  addShapeAbove?: string;
  addShapeBelow?: string;
  addAssistant?: string;
  /**
   * Create Graphic > Add Bullet: a Level 2 line under the selected shape. In
   * the data model it is the same child point Add Shape Below makes, so it
   * runs that position and is offered exactly when the layout nests.
   */
  addBullet?: string;
  /** Why Add Assistant is refused on an organization chart layout. */
  addAssistantUnavailable?: string;
  /** Why Add Assistant is refused on any other layout. */
  addAssistantNotOrgChart?: string;
  /** Layout on an org chart box, and its four entries (Office's words). */
  orgChartLayout?: string;
  orgChartLayoutStandard?: string;
  orgChartLayoutBoth?: string;
  orgChartLayoutLeftHanging?: string;
  orgChartLayoutRightHanging?: string;
  /** Why Layout is refused on a layout that is no organization chart. */
  orgChartLayoutNotOrgChart?: string;
  moveUp?: string;
  moveDown?: string;
  convert?: string;
  convertToShapes?: string;
  convertToText?: string;
  /**
   * Why a greyed command is greyed. Office greys a command it cannot run and
   * MRAI says why (`data-disabled-reason`), so every disabled command here
   * carries one of these. Optional so an older caller keeps compiling; the
   * shared `useOfficeSmartArtDesignLabels` supplies every one.
   */
  readOnlyReason?: string;
  /** No shape of the diagram is selected, so an outline verb has no target. */
  noShapeSelectedReason?: string;
  /** The selected shape does not take this verb in this layout. */
  commandUnavailableReason?: string;
  /** A preserved (native) diagram: its structure belongs to the package. */
  nativeStructureUnavailable?: string;
  /** SmartArt Format > Shapes: the group caption and its two commands. */
  shapesGroup?: string;
  larger?: string;
  smaller?: string;
}

/** The selected AUTHORED diagram: what the tab shows, and where each command goes. */
export interface OfficeSmartArtDesignTabDiagram {
  /** Its catalog layout, checked in the layouts gallery. */
  kind: SmartArtKind;
  /** Its quick style, checked in the style gallery; null when members disagree. */
  styleId: SmartArtStyleId | null;
  /**
   * Whether the diagram runs Right to Left. Right to Left is a toggle (click
   * it again to go back), so the command shows as pressed while it is on.
   * Omitted by a host that cannot say, which then draws it unpressed.
   */
  rightToLeft?: boolean;
  /** Its members, read only to block a Picture → non-Picture layout change. */
  members: readonly OfficeSmartArtBlockLike[];
  /**
   * The Office alias it is drawn as, when the host keeps that on the diagram
   * rather than on every member (the Docs group); else read off the members.
   */
  layoutUrn?: string | null;
  /** False disables every command that would write (read-only page, lock, …). */
  writable: boolean;
  /** Whether the diagram has an outline at all (gates the Text Pane toggle). */
  hasOutline: boolean;
  textPaneOpen: boolean;
  canPromote: boolean;
  canDemote: boolean;
  /**
   * The diagram's picture slots ({@link smartArtPictureSlots}), in member
   * order. Empty or omitted and no picture command is drawn at all, which is
   * every non-picture layout and every host that does not offer one.
   */
  pictureSlots?: readonly SmartArtPictureSlot[];
  /**
   * Fill or replace ONE slot. The host opens its own file picker — this tab
   * owns no state and no transaction — reads the file through
   * {@link readSmartArtPictureFile}, and writes the `src` onto that member in
   * one undo step.
   */
  onPickPicture?: (slot: SmartArtPictureSlot) => void;
  onToggleTextPane: () => void;
  onPromote: () => void;
  onDemote: () => void;
  onRightToLeft: () => void;
  /**
   * Change Layout: `layoutUrn` is the aligned Office alias picked (its own
   * gallery entry), `null` a catalog layout -- which gives that real layout.
   */
  onChangeLayout: (kind: SmartArtKind, layoutUrn?: string | null) => void;
  /**
   * Change Colors: the diagram's colour set, and each item's own fill and line
   * when the variation paints the shapes in turn (Colorful, Gradient Range).
   */
  onChangeColors: (
    colors: SmartArtStyle,
    nodePaint?: readonly SmartArtNodePaint[],
    /** The variation picked, by Office's id, for a host that records it. */
    variantId?: string,
  ) => void;
  /**
   * The colour variation the diagram wears (Office's id), checked in the
   * Change Colors gallery. Omitted by a host that does not record one.
   */
  colorsId?: string | null;
  /**
   * The colour set the diagram wears, resolved on the host's theme: the
   * Styles gallery's tiles are drawn in it, as PowerPoint's are (a tile
   * without it is drawn in the default blue; acceptance finding 8).
   */
  colors?: SmartArtStyle;
  onChangeStyle: (styleId: SmartArtStyleId) => void;
  onReset: () => void;
  /** `trigger` is the element focus returns to when the dialog closes. */
  onAltText: (trigger: HTMLElement | null) => void;
  /**
   * What the selected item accepts ({@link smartArtOutlineCommandState}); the
   * right-click menu reads the same object, so the two never disagree.
   */
  commands?: SmartArtOutlineCommandState;
  /** Add Shape After / Before / Above / Below; absent, no Add Shape is drawn. */
  onAddShape?: (position: SmartArtOutlineInsertPosition) => void;
  /**
   * Layout on an organization chart (Standard / Both / Left Hanging / Right
   * Hanging) for the selected box; absent, no Layout menu is drawn.
   */
  onOrgChartLayout?: (branch: SmartArtHierBranch) => void;
  /** The selected box's own Layout, checked in the menu; absent: Standard. */
  selectedHierBranch?: SmartArtHierBranch | null;
  onMoveUp?: () => void;
  onMoveDown?: () => void;
  /**
   * Reset > Convert. PowerPoint has both entries, Excel has Convert to Shapes
   * only and Word has neither (Microsoft Support, "Convert your SmartArt
   * graphic into shapes"), so a host passes exactly the ones its Office app
   * has and the menu is not drawn when it passes none.
   */
  onConvertToShapes?: () => void;
  onConvertToText?: () => void;
  /**
   * SmartArt Format > Shapes > Larger / Smaller on the selected item: its scale,
   * the graphic laid out again around it (`smartArtMembersWithItemScale`).
   * PowerPoint for the web has one SmartArt tab, so
   * the Shapes group sits in it here rather than on a second tab.
   */
  onResizeShape?: (direction: "larger" | "smaller") => void;
  canResizeShape?: boolean;
  /**
   * The selected item is an organization chart's assistant, which is what
   * makes Add Assistant say "select its manager" rather than a generic reason.
   */
  selectedIsAssistant?: boolean;
  /**
   * Whether one shape (outline item) of the diagram is the target of the
   * outline verbs. `false` makes a greyed verb say "select a shape"; omitted,
   * a greyed verb says it does not apply to the selected shape.
   */
  hasSelectedItem?: boolean;
}

/** What a diagram's tab says of it, without the commands that leave it. */
export type OfficeSmartArtDesignTabDiagramState = Omit<
  OfficeSmartArtDesignTabDiagram,
  | "onToggleTextPane"
  | "onPromote"
  | "onDemote"
  | "onRightToLeft"
  | "onChangeLayout"
  | "onChangeColors"
  | "onChangeStyle"
  | "onReset"
  | "onAltText"
>;

/** The selected NATIVE (preserved) diagram — the pre-conversion surface. */
export interface OfficeSmartArtDesignTabPreserved {
  /** Offer "convert to editable" (the author picks the catalog layout). */
  convertible: boolean;
  /** The diagram MRAI edits in place through its native text receipt. */
  nativeEditing: boolean;
  /**
   * Why this diagram's text cannot be edited, already translated. Present
   * only when `nativeEditing` is false AND the reason is known — an Office
   * diagram still holding its `[Text]` prompts has no text in the file to
   * change. Edit Text is then SHOWN and disabled carrying this reason, so an
   * author reads why instead of finding nothing to type into.
   */
  textBlockedReason?: string | null;
  /** False disables the convert gallery (read-only page). */
  writable: boolean;
  /**
   * Why converting this diagram into a PICTURE layout cannot work from here,
   * already translated — the diagram carries pictures and there is no saved
   * page to read them out of (the guest preview), so the request the command
   * would make could only fail.
   *
   * Present, the picture layouts of the 转换为可编辑 gallery are SHOWN and
   * disabled carrying this reason, the same move `textBlockedReason` makes for
   * Edit Text and Change Layout makes for a Picture layout a pictured diagram
   * may not become. Office greys a command out; it does not offer one that
   * must do nothing. Every non-picture layout stays offered, because it
   * converts with no request at all.
   */
  pictureConversionBlockedReason?: string | null;
  textPaneOpen: boolean;
  canAdd: boolean;
  canPromote: boolean;
  canDemote: boolean;
  /**
   * Move Up / Move Down, where the host can rewrite the package's point
   * order (a flat flow it lays out itself). Omitted: the pair is not drawn.
   */
  canMoveUp?: boolean;
  canMoveDown?: boolean;
  onConvert: (kind: SmartArtKind) => void;
  onToggleTextPane: () => void;
  onAddShape: () => void;
  onPromote: () => void;
  onDemote: () => void;
  onMoveUp?: () => void;
  onMoveDown?: () => void;
}

export interface OfficeSmartArtDesignTabProps {
  /** `"presentation"` → `presentation-ribbon-smartart-…`, and so on. */
  testIdPrefix: string;
  /**
   * The ribbon token every group of this tab carries (`data-ribbon`). A host
   * whose strip collects its contextual groups by that token (Docs' strip
   * keeps only the children whose `ribbon` names the active tab) mounts this
   * element DIRECTLY in its track with `ribbon="smartArt"` — never through a
   * `display: contents` wrapper, which takes the groups out of the track's
   * `.note-editor-ribbon-track > .note-editor-office-toolbar__group` anatomy
   * (column flex, caption under the commands) and leaves them as row flex
   * boxes whose caption lands beside or over the commands. Defaults to the
   * shared group's own `"home"`.
   */
  ribbon?: string;
  /** Which editor's gallery chrome to wear. */
  application: "docs" | "spreadsheet" | "presentation";
  labels: OfficeSmartArtDesignTabLabels;
  diagram: OfficeSmartArtDesignTabDiagram | null;
  preserved: OfficeSmartArtDesignTabPreserved | null;
  classes?: Partial<OfficeSmartArtDesignTabClasses>;
  /** The host's control-id vocabulary, where it has one. */
  controlIds?: { colors?: string };
  /** Focus returns here when the text pane closes. */
  textPaneTriggerRef?: RefObject<HTMLButtonElement | null>;
  /** The words the host's insert seeds (`useOfficeSmartArtPlaceholders`). */
  placeholders?: OfficeSmartArtPlaceholders;
  /** The frame the host's insert lays a diagram out in (a deck's slide). */
  previewBounds?: { width: number; height: number };
  /** Draw one style-tile shape the host's way; omitted, the shared preset. */
  renderStyleShape?: (node: OfficeSmartArtShapeNode) => ReactNode;
  /**
   * The document's theme colours, which Change Colors resolves its variations
   * against (a deck's or a workbook's accents); omitted, Office's own theme.
   */
  colorTheme?: Partial<SmartArtColorTheme> | null;
  /**
   * The Change Colors gallery's groups. Omitted, the list the Docs and
   * worksheet hosts ship (`OFFICE_SMARTART_COLOR_GROUPS`); the deck passes
   * PowerPoint's whole gallery (`OFFICE_SMARTART_POWERPOINT_COLOR_GROUPS`).
   */
  colorGroups?: readonly SmartArtColorGroup[];
}

/**
 * Add Shape's four positions in Office's menu order, each with its label and
 * whether the selected item accepts it. The Design tab's split button and the
 * right-click Add Shape submenu both render this list.
 */
export function officeSmartArtAddShapeEntries(
  labels: Pick<
    OfficeSmartArtDesignTabLabels,
    "addShapeAfter" | "addShapeBefore" | "addShapeAbove" | "addShapeBelow"
  >,
  commands: SmartArtOutlineCommandState,
): { position: SmartArtOutlineInsertPosition; label: string; enabled: boolean }[] {
  return [
    { position: "after", label: labels.addShapeAfter ?? "", enabled: commands.addAfter },
    { position: "before", label: labels.addShapeBefore ?? "", enabled: commands.addBefore },
    { position: "above", label: labels.addShapeAbove ?? "", enabled: commands.addAbove },
    { position: "below", label: labels.addShapeBelow ?? "", enabled: commands.addBelow },
  ];
}

/**
 * Why Add Assistant is greyed, or `undefined` when it is offered. Office has
 * it on organization charts only; on one, it is refused where the selected
 * shape is itself an assistant (an assistant has no assistant of its own) and
 * otherwise for the host's own reason (`fallback`: read-only, nothing
 * selected, the layout's capacity).
 */
export function officeSmartArtAssistantReason(
  labels: Pick<
    OfficeSmartArtDesignTabLabels,
    "addAssistantUnavailable" | "addAssistantNotOrgChart"
  >,
  kind: SmartArtKind,
  enabled = false,
  selectedIsAssistant = false,
  fallback?: string,
): string | undefined {
  if (!smartArtLayoutHasAssistant(kind)) return labels.addAssistantNotOrgChart;
  if (enabled) return undefined;
  return selectedIsAssistant ? labels.addAssistantUnavailable : fallback;
}


export function OfficeSmartArtDesignTab({
  testIdPrefix,
  ribbon,
  application,
  labels,
  diagram,
  preserved,
  classes,
  controlIds,
  textPaneTriggerRef,
  placeholders,
  previewBounds,
  renderStyleShape,
  colorTheme,
  colorGroups = OFFICE_SMARTART_COLOR_GROUPS,
}: OfficeSmartArtDesignTabProps) {
  const cls = (slot: keyof OfficeSmartArtDesignTabClasses) =>
    classes?.[slot] ?? styles[slot];
  const ribbonId = (suffix: string) =>
    `${testIdPrefix}-ribbon-smartart-${suffix}`;
  const surfaceId = (suffix: string) => `${testIdPrefix}-smartart-${suffix}`;
  const pictureSlots = diagram?.pictureSlots ?? [];
  const commands = diagram?.commands ?? SMARTART_NO_OUTLINE_COMMANDS;
  const addShapeEntries = officeSmartArtAddShapeEntries(labels, commands);
  const assistantEnabled = Boolean(diagram?.writable && commands.addAssistant);
  // Why a greyed authored-diagram command is greyed (undefined when enabled).
  const unavailable = (enabled: boolean): string | undefined =>
    enabled || !diagram
      ? undefined
      : !diagram.writable
        ? labels.readOnlyReason
        : diagram.hasSelectedItem === false
          ? labels.noShapeSelectedReason
          : labels.commandUnavailableReason;
  const writableReason = (): string | undefined =>
    diagram?.writable === false ? labels.readOnlyReason : undefined;
  const addShapeMenuEnabled = Boolean(
    diagram?.writable &&
      (addShapeEntries.some((entry) => entry.enabled) || commands.addAssistant),
  );
  const assistantReason = diagram
    ? officeSmartArtAssistantReason(
        labels,
        diagram.kind,
        assistantEnabled,
        diagram.selectedIsAssistant === true,
        unavailable(assistantEnabled),
      )
    : undefined;
  // The diagram's own layout when it is one the gallery does not list — at
  // most one entry, spread into the gallery's items and its quick row below.
  const currentConvertOnlyLayout = OFFICE_SMARTART_CONVERT_ONLY_LAYOUTS.filter(
    (entry) => entry.kind === diagram?.kind,
  );
  // PowerPoint's Change Layout sheet is sectioned by category (List, Process,
  // Cycle, ...), as its Choose a SmartArt Graphic dialog is: a layout is in
  // EVERY tab its definition's `catLst` names, at its own place there
  // (`officeSmartArtGalleryTabs`), and so is each alias PowerPoint lists as a
  // layout of its own. Entries that are only "current" (convert-only) sit in
  // their own tabs like any other, so the marked tile is where the author
  // expects it.
  const aliasGallery = labels.aliasLayoutName ? SMARTART_ALIAS_GALLERY : [];
  const layoutGroups = (entries: readonly OfficeSmartArtLayout[], withAliases = false) => {
    if (!labels.layoutCategoryName) return undefined;
    const sources = entries.flatMap((entry) => [
      { itemId: entry.kind, cardId: entry.id },
      ...(withAliases
        ? aliasGallery
            .filter((alias) => alias.catalogId === entry.id)
            .map((alias) => ({ itemId: alias.galleryId, cardId: `${alias.catalogId}@${alias.suffix}` }))
        : []),
    ]);
    return OFFICE_SMARTART_CATEGORY_IDS.map((category) => ({
      id: category,
      label: labels.layoutCategoryName!(category),
      itemIds: sources
        .flatMap(({ itemId, cardId }, order) => {
          const tab = officeSmartArtGalleryTabs(cardId).find((candidate) => candidate.category === category);
          return tab ? [{ itemId, pri: tab.pri, order }] : [];
        })
        .sort((a, b) => a.pri - b.pri || a.order - b.order)
        .map(({ itemId }) => itemId as SmartArtKind),
    })).filter((group) => group.itemIds.length > 0);
  };
  // One thumbnail box per surface: the sheet's art tile is 116x84, the ribbon
  // strip keeps its 58x44 slot. The shared seed preview scales into either,
  // and draws the scene the Insert gallery card for the same layout draws.
  const layoutThumbnail = (galleryId: SmartArtKind, surface: "inline" | "menu") => {
    const { kind: id, layoutUrn } = parseSmartArtGalleryId(galleryId);
    return (
    <span
      aria-hidden="true"
      style={
        surface === "menu"
          ? { display: "block", height: SMARTART_MENU_THUMB[1], width: SMARTART_MENU_THUMB[0] }
          : { display: "block", height: 44, width: 58 }
      }
    >
      <OfficeSmartArtSeedPreview
        layoutId={id}
        {...(layoutUrn ? { layoutUrn } : {})}
        itemLabel={placeholders?.itemLabel}
        rootLabel={placeholders?.rootLabel}
        bounds={previewBounds}
        fit="content"
        labels="lines"
      />
    </span>
    );
  };

  return (
    <>
      {preserved ? (
        <OfficeHomeRibbonGroup
          ribbon={ribbon}
          className={cls("createGroup")}
          label={labels.createGraphicGroup}
          size="standard"
          testId={`${testIdPrefix}-ribbon-native-smartart-create-group`}
        >
          <OfficeHomeCommandRow>
            {/* A preserved diagram MRAI does not own stays preserved until the
                author picks the catalog layout it becomes, which it does inside
                its own frame. */}
            {preserved.convertible ? (
              <OfficeStyleGallery<SmartArtKind>
                application={application}
                ariaLabel={labels.layoutsGroup}
                disabled={!preserved.writable}
                items={OFFICE_SMARTART_LAYOUTS.map((entry) => {
                  // The picture reason is about the diagram and greys the whole
                  // picture family. A layout's capacity greys nothing: the
                  // items it does not draw stay as hidden members.
                  const reason =
                    (entry.picture
                      ? preserved.pictureConversionBlockedReason
                      : null) ?? null;
                  return {
                    id: entry.kind,
                    label: labels.layoutName(entry),
                    disabled: Boolean(reason),
                    disabledReason: reason ?? undefined,
                  };
                })}
                labels={{ more: labels.convertToEditable }}
                groups={layoutGroups(OFFICE_SMARTART_LAYOUTS)}
                repeatAcrossGroups
                menuTile="art"
                menuGrid={SMARTART_MENU_GRID}
                quickIds={[]}
                menuMode="labeled"
                menuTrigger={
                  <OfficeHomeMenuTrigger
                    variant="large"
                    icon={
                      <OfficeSemanticIcon
                        semanticId="diagram.smartArt.insert"
                        size={24}
                      />
                    }
                    label={labels.convertToEditable}
                    title={labels.convertToEditable}
                    disabled={!preserved.writable}
                    testId={surfaceId("convert-native")}
                  />
                }
                renderPreview={(item, surface) => layoutThumbnail(item.id, surface)}
                testIdPrefix={surfaceId("convert-gallery")}
                value={null}
                onPreviewChange={() => undefined}
                onCommit={preserved.onConvert}
              />
            ) : null}
            {preserved.nativeEditing ? (
              <>
                <OfficeHomeCommand
                  variant="large"
                  active={preserved.textPaneOpen}
                  buttonRef={textPaneTriggerRef}
                  icon={
                    <OfficeSemanticIcon
                      semanticId="view.navigationPane"
                      size={24}
                    />
                  }
                  label={labels.textPane}
                  onClick={preserved.onToggleTextPane}
                  testId={ribbonId("textPane")}
                />
                <OfficeHomeCommand
                  variant="large"
                  icon={<Plus size={24} aria-hidden="true" />}
                  label={labels.addShape}
                  disabled={!preserved.canAdd}
                  disabledReason={
                    preserved.canAdd ? undefined : labels.nativeStructureUnavailable
                  }
                  onClick={preserved.onAddShape}
                  testId={ribbonId("addShape")}
                />
                <OfficeHomeCommandStack>
                  <OfficeHomeCommand
                    icon={
                      <OfficeSemanticIcon
                        semanticId="paragraph.indent.decrease"
                        size={20}
                      />
                    }
                    label={labels.promote}
                    disabled={!preserved.canPromote}
                    disabledReason={
                      preserved.canPromote
                        ? undefined
                        : labels.nativeStructureUnavailable
                    }
                    onClick={preserved.onPromote}
                    testId={ribbonId("promote")}
                  />
                  <OfficeHomeCommand
                    icon={
                      <OfficeSemanticIcon
                        semanticId="paragraph.indent.increase"
                        size={20}
                      />
                    }
                    label={labels.demote}
                    disabled={!preserved.canDemote}
                    disabledReason={
                      preserved.canDemote
                        ? undefined
                        : labels.nativeStructureUnavailable
                    }
                    onClick={preserved.onDemote}
                    testId={ribbonId("demote")}
                  />
                </OfficeHomeCommandStack>
                {preserved.onMoveUp || preserved.onMoveDown ? (
                  <OfficeHomeCommandStack>
                    <OfficeHomeCommand
                      icon={<ArrowUp size={18} aria-hidden="true" />}
                      label={labels.moveUp ?? ""}
                      disabled={!preserved.canMoveUp}
                      disabledReason={
                        preserved.canMoveUp
                          ? undefined
                          : labels.nativeStructureUnavailable
                      }
                      onClick={() => preserved.onMoveUp?.()}
                      testId={ribbonId("move-up")}
                    />
                    <OfficeHomeCommand
                      icon={<ArrowDown size={18} aria-hidden="true" />}
                      label={labels.moveDown ?? ""}
                      disabled={!preserved.canMoveDown}
                      disabledReason={
                        preserved.canMoveDown
                          ? undefined
                          : labels.nativeStructureUnavailable
                      }
                      onClick={() => preserved.onMoveDown?.()}
                      testId={ribbonId("move-down")}
                    />
                  </OfficeHomeCommandStack>
                ) : null}
              </>
            ) : preserved.textBlockedReason ? (
              // Office's own command, in Office's own place, disabled with the
              // reason: the diagram carries no text this host could write to.
              <OfficeHomeCommand
                variant="large"
                icon={
                  <OfficeSemanticIcon semanticId="view.navigationPane" size={24} />
                }
                label={labels.textPane}
                disabled
                disabledReason={preserved.textBlockedReason}
                onClick={() => undefined}
                testId={ribbonId("textPane")}
              />
            ) : null}
          </OfficeHomeCommandRow>
        </OfficeHomeRibbonGroup>
      ) : null}
      {diagram ? (
        <>
          <OfficeHomeRibbonGroup
            ribbon={ribbon}
            className={cls("createGroup")}
            label={labels.createGraphicGroup}
            size="standard"
            testId={ribbonId("create-group")}
          >
            <OfficeHomeCommandRow
              className={cls("createRow")}
              testId={ribbonId("create-row")}
            >
              {diagram.onAddShape ? (
                <DropdownMenu>
                  <OfficeHomeSplitLayout
                    testId={ribbonId("add-shape-split")}
                    primary={
                      <OfficeHomeCommand
                        variant="large"
                        icon={<Plus size={24} aria-hidden="true" />}
                        label={labels.addShape}
                        disabled={!diagram.writable || !commands.addAfter}
                        disabledReason={unavailable(
                          diagram.writable && commands.addAfter,
                        )}
                        onClick={() => diagram.onAddShape?.("after")}
                        testId={ribbonId("add-shape")}
                      />
                    }
                    trigger={
                      <DropdownMenuTrigger asChild>
                        <OfficeHomeMenuTrigger
                          variant="icon"
                          icon={
                            <OfficeSemanticIcon
                              semanticId="menu.chevron.down"
                              size={12}
                            />
                          }
                          label={labels.addShape}
                          // A menu whose every position is refused is not
                          // offered: the chevron greys with the face and says
                          // why, instead of opening a list of greyed items.
                          disabled={!addShapeMenuEnabled}
                          disabledReason={unavailable(addShapeMenuEnabled)}
                          testId={ribbonId("add-shape-menu")}
                        />
                      </DropdownMenuTrigger>
                    }
                  />
                  <RibbonDropdownMenuContent
                    align="start"
                    data-testid={ribbonId("add-shape-menu-content")}
                  >
                    {addShapeEntries.map((entry) => (
                      <DropdownMenuItem
                        key={entry.position}
                        disabled={!entry.enabled}
                        onSelect={() => diagram.onAddShape?.(entry.position)}
                        data-testid={ribbonId(`add-shape-${entry.position}`)}
                      >
                        {entry.label}
                      </DropdownMenuItem>
                    ))}
                    <DropdownMenuItem
                      disabled={!assistantEnabled}
                      title={assistantReason}
                      data-disabled-reason={assistantReason}
                      onSelect={() => diagram.onAddShape?.("assistant")}
                      data-testid={ribbonId("add-assistant")}
                    >
                      {labels.addAssistant}
                    </DropdownMenuItem>
                  </RibbonDropdownMenuContent>
                </DropdownMenu>
              ) : null}
              <OfficeHomeCommand
                variant="large"
                active={diagram.textPaneOpen}
                buttonRef={textPaneTriggerRef}
                icon={
                  <OfficeSemanticIcon
                    semanticId="view.navigationPane"
                    size={24}
                  />
                }
                label={labels.textPane}
                disabled={!diagram.hasOutline}
                disabledReason={unavailable(diagram.hasOutline)}
                onClick={diagram.onToggleTextPane}
                testId={ribbonId("textPane")}
              />
              <OfficeHomeCommandStack className={cls("createHierarchyLane")} rows={3}>
                <OfficeHomeCommand
                  icon={
                    <OfficeSemanticIcon
                      semanticId="paragraph.indent.decrease"
                      size={20}
                    />
                  }
                  label={labels.promote}
                  disabled={!diagram.canPromote}
                  disabledReason={unavailable(diagram.canPromote)}
                  onClick={() => diagram.onPromote()}
                  testId={ribbonId("promote")}
                />
                <OfficeHomeCommand
                  icon={
                    <OfficeSemanticIcon
                      semanticId="paragraph.indent.increase"
                      size={20}
                    />
                  }
                  label={labels.demote}
                  disabled={!diagram.canDemote}
                  disabledReason={unavailable(diagram.canDemote)}
                  onClick={() => diagram.onDemote()}
                  testId={ribbonId("demote")}
                />
                <OfficeHomeCommand
                  icon={
                    <OfficeSemanticIcon
                      semanticId="object.flipHorizontal"
                      size={20}
                    />
                  }
                  label={labels.rightToLeft}
                  active={diagram.rightToLeft}
                  disabled={!diagram.writable}
                  disabledReason={writableReason()}
                  onClick={() => diagram.onRightToLeft()}
                  testId={ribbonId("right-to-left")}
                />
              </OfficeHomeCommandStack>
              {diagram.onOrgChartLayout && labels.orgChartLayout ? (
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <OfficeHomeMenuTrigger
                      variant="large"
                      hasMenuIndicator
                      icon={<OfficeSemanticIcon semanticId="list.tree" size={24} />}
                      label={labels.orgChartLayout}
                      displayLabel={labels.orgChartLayout}
                      disabled={!diagram.writable || !commands.orgChartLayout}
                      disabledReason={
                        diagram.writable && !commands.orgChartLayout && !smartArtLayoutTakesHierBranch(diagram.kind)
                          ? labels.orgChartLayoutNotOrgChart
                          : unavailable(diagram.writable && commands.orgChartLayout)
                      }
                      testId={ribbonId("org-chart-layout")}
                    />
                  </DropdownMenuTrigger>
                  <RibbonDropdownMenuContent
                    align="start"
                    data-testid={ribbonId("org-chart-layout-menu")}
                  >
                    {(
                      [
                        ["std", labels.orgChartLayoutStandard],
                        ["hang", labels.orgChartLayoutBoth],
                        ["l", labels.orgChartLayoutLeftHanging],
                        ["r", labels.orgChartLayoutRightHanging],
                      ] as const
                    ).map(([branch, label]) => (
                      <DropdownMenuItem
                        key={branch}
                        onSelect={() => diagram.onOrgChartLayout?.(branch)}
                        data-testid={ribbonId(`org-chart-layout-${branch}`)}
                        data-checked={(diagram.selectedHierBranch ?? "std") === branch ? "true" : undefined}
                        aria-checked={(diagram.selectedHierBranch ?? "std") === branch}
                        role="menuitemradio"
                      >
                        {label ?? branch}
                      </DropdownMenuItem>
                    ))}
                  </RibbonDropdownMenuContent>
                </DropdownMenu>
              ) : null}
              {diagram.onMoveUp || diagram.onMoveDown ? (
                <OfficeHomeCommandStack
                  rows={
                    diagram.onAddShape && labels.addBullet ? 3 : undefined
                  }
                >
              {diagram.onAddShape && labels.addBullet ? (
                <OfficeHomeCommand
                  icon={<List size={18} aria-hidden="true" />}
                  label={labels.addBullet}
                  disabled={!diagram.writable || !commands.addBullet}
                  disabledReason={unavailable(
                    diagram.writable && commands.addBullet,
                  )}
                  onClick={() => diagram.onAddShape?.("bullet")}
                  testId={ribbonId("add-bullet")}
                />
              ) : null}
                  <OfficeHomeCommand
                    icon={<ArrowUp size={18} aria-hidden="true" />}
                    label={labels.moveUp ?? ""}
                    disabled={!diagram.writable || !commands.moveUp}
                    disabledReason={unavailable(
                      diagram.writable && commands.moveUp,
                    )}
                    onClick={() => diagram.onMoveUp?.()}
                    testId={ribbonId("move-up")}
                  />
                  <OfficeHomeCommand
                    icon={<ArrowDown size={18} aria-hidden="true" />}
                    label={labels.moveDown ?? ""}
                    disabled={!diagram.writable || !commands.moveDown}
                    disabledReason={unavailable(
                      diagram.writable && commands.moveDown,
                    )}
                    onClick={() => diagram.onMoveDown?.()}
                    testId={ribbonId("move-down")}
                  />
                </OfficeHomeCommandStack>
              ) : null}
              {/* The picture command, for a PICTURE layout only. Office puts a
                  picture icon in each empty slot and opens the insert-picture
                  dialog for that one slot; MRAI names the slots in a menu
                  instead, because a Docs group member and a worksheet member
                  have no in-place picture affordance of their own and the
                  author must be able to say WHICH slot regardless of what the
                  canvas has selected. Filled slots say so, the way Office
                  shows the picture already in the frame. */}
              {pictureSlots.length > 0 &&
              labels.insertPicture &&
              labels.pictureSlotName &&
              diagram.onPickPicture ? (
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <OfficeHomeMenuTrigger
                      variant="large"
                      hasMenuIndicator
                      icon={
                        <OfficeSemanticIcon
                          semanticId="image.change"
                          size={24}
                        />
                      }
                      label={labels.insertPicture}
                      title={labels.insertPicture}
                      disabled={!diagram.writable}
                      disabledReason={writableReason()}
                      testId={ribbonId("picture")}
                    />
                  </DropdownMenuTrigger>
                  <RibbonDropdownMenuContent
                    align="start"
                    data-testid={ribbonId("picture-menu")}
                  >
                    {pictureSlots.map((slot) => (
                      <DropdownMenuItem
                        key={slot.memberIndex}
                        // A slot Office's own layout binds no picture to is
                        // shown and DISABLED, carrying why: accepting a picture
                        // MRAI knows Office discards on its first re-save is
                        // the silent loss this whole campaign is against, and
                        // hiding the slot would leave the author looking for a
                        // slot the canvas is drawing.
                        disabled={!slot.exportable}
                        title={
                          slot.exportable
                            ? undefined
                            : labels.pictureSlotUnexportable
                        }
                        onSelect={() =>
                          slot.exportable
                            ? diagram.onPickPicture?.(slot)
                            : undefined
                        }
                        data-slot-filled={slot.filled ? "true" : "false"}
                        data-slot-exportable={slot.exportable ? "true" : "false"}
                        data-testid={ribbonId(`picture-slot-${slot.ordinal}`)}
                      >
                        {labels.pictureSlotName?.(slot)}
                      </DropdownMenuItem>
                    ))}
                  </RibbonDropdownMenuContent>
                </DropdownMenu>
              ) : null}
            </OfficeHomeCommandRow>
          </OfficeHomeRibbonGroup>
          <OfficeHomeRibbonGroup
            ribbon={ribbon}
            className={cls("layoutsGroup")}
            label={labels.layoutsGroupLabel}
            size="gallery"
            testId={ribbonId("layouts-group")}
          >
            <OfficeStyleGallery<SmartArtKind>
              application={application}
              ariaLabel={labels.layoutsGroup}
              className={cls("layoutGallery")}
              disabled={!diagram.writable}
              items={[
                // The diagram's OWN layout, when the gallery cannot list it.
                //
                // A convert-only layout is drawn and owned by MRAI but never
                // OFFERED, because there is no preset to author one from — so
                // a converted Upward Arrow or Picture Grid sat in Change
                // Layout with nothing marked current and the author could not
                // tell what the diagram was (§5.38). It is shown here as a
                // DISABLED tile that the gallery's `value` selects, which is
                // the honest answer: this IS the layout, and it is not one you
                // can switch to.
                ...currentConvertOnlyLayout,
                ...OFFICE_SMARTART_LAYOUTS,
              ].map((entry) => {
                if (entry.convertOnly) {
                  return {
                    id: entry.kind,
                    label: labels.layoutName(entry),
                    disabled: true,
                    disabledReason: labels.currentLayoutUnavailable,
                  };
                }
                // Every offered layout is open, a picture layout's included:
                // Office keeps the pictures a layout without picture shapes
                // cannot show, and shows them again on a picture layout
                // (`./held-pictures`; Microsoft Support, "Anything above
                // marked with a red X…"). This tile used to be greyed.
                return {
                  id: entry.kind,
                  label: labels.layoutName(entry),
                };
              }).flatMap((item) => [
                item,
                ...aliasGallery
                  .filter((alias) => alias.kind === item.id)
                  .map((alias) => ({
                    id: alias.galleryId as SmartArtKind,
                    label: labels.aliasLayoutName!(alias.suffix),
                    disabled: item.disabled,
                    disabledReason: item.disabledReason,
                  })),
              ])}
              labels={{ more: labels.quickOverflow }}
              groups={layoutGroups([
                ...currentConvertOnlyLayout,
                ...OFFICE_SMARTART_LAYOUTS,
              ], true)}
              repeatAcrossGroups
              menuTile="art"
              menuGrid={SMARTART_MENU_GRID}
              quickIds={[
                // The current-layout marker leads the quick row, so it is
                // visible without opening the overflow.
                ...currentConvertOnlyLayout,
                ...OFFICE_SMARTART_LAYOUTS.slice(0, 6),
              ].map((entry) => entry.kind)}
              renderPreview={(item, surface) => layoutThumbnail(item.id, surface)}
              testIdPrefix={surfaceId("layout-gallery")}
              value={smartArtGalleryId(diagram.kind, diagram.layoutUrn ?? readSmartArtGroupLayoutUrn(diagram.members)) as SmartArtKind}
              onPreviewChange={() => undefined}
              onCommit={(id) => {
                const picked = parseSmartArtGalleryId(id);
                diagram.onChangeLayout(picked.kind as SmartArtKind, picked.layoutUrn);
              }}
            />
          </OfficeHomeRibbonGroup>
          <OfficeHomeRibbonGroup
            ribbon={ribbon}
            className={cls("stylesGroup")}
            label={labels.stylesGroup}
            size="gallery"
            testId={ribbonId("styles-group")}
          >
            <OfficeHomeCommandRow>
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    variant="large"
                    displayLabel={labels.changeColors}
                    hasMenuIndicator
                    icon={
                      <OfficeSemanticIcon semanticId="image.color" size={24} />
                    }
                    label={labels.changeColors}
                    disabled={!diagram.writable}
                    disabledReason={writableReason()}
                    controlId={controlIds?.colors}
                    testId={ribbonId("colors")}
                  />
                </DropdownMenuTrigger>
                <OfficeGalleryMenuContent data-testid={ribbonId("colors-menu")}
                >
                  {/* Office's Change Colors gallery: its groups, each a row of
                      variations drawn in the document's own theme colours. */}
                  {colorGroups.map((group) => {
                    const heading = labels.colorGroupName(group);
                    return (
                      <div
                        key={group.id}
                        role="group"
                        aria-label={heading}
                        data-testid={ribbonId(`color-group-${group.id}`)}
                      >
                        <OfficeGalleryMenuSection>{heading}</OfficeGalleryMenuSection>
                        <OfficeGalleryMenuGrid
                          columns={SMARTART_COLOR_COLUMNS}
                          tileWidth={SMARTART_COLOR_TILE_PX}
                        >
                          {group.variants.map((variant) => {
                            const name = labels.colorName(variant);
                            const swatch = smartArtColorVariantSwatch(
                              variant,
                              colorTheme,
                            );
                            const current =
                              diagram.colorsId !== undefined &&
                              (diagram.colorsId ?? "accent1_2") === variant.id;
                            return (
                              <DropdownMenuItem
                                key={variant.id}
                                aria-label={name}
                                title={name}
                                className={officeGalleryMenuStyles.tile}
                                data-nameless="true"
                                data-selected={current ? "true" : "false"}
                                data-current={current ? "true" : undefined}
                                onSelect={() => {
                                  const resolved = resolveSmartArtColorVariant(
                                    variant,
                                    smartArtDiagramItemCount(diagram.members),
                                    colorTheme,
                                  );
                                  diagram.onChangeColors(
                                    resolved.colors,
                                    resolved.nodePaint,
                                    variant.id,
                                  );
                                }}
                                data-testid={ribbonId(`color-${variant.id}`)}
                              >
                                <OfficeGalleryMenuTileBody
                                  thumbnail={
                                    <span className={styles.colorSwatch}>
                                      {swatch.fills.map((fill, index) => (
                                        <span
                                          key={index}
                                          style={{
                                            background: fill,
                                            borderColor: swatch.line,
                                            // Transparent Gradient Range: the
                                            // shapes are drawn translucent.
                                            ...(swatch.opacities
                                              ? { opacity: swatch.opacities[index] / 100 }
                                              : {}),
                                          }}
                                        />
                                      ))}
                                    </span>
                                  }
                                />
                              </DropdownMenuItem>
                            );
                          })}
                        </OfficeGalleryMenuGrid>
                      </div>
                    );
                  })}
                </OfficeGalleryMenuContent>
              </DropdownMenu>
              <OfficeStyleGallery<SmartArtStyleId>
                application={application}
                ariaLabel={labels.stylesGroup}
                className={cls("styleGallery")}
                disabled={!diagram.writable}
                items={SMARTART_STYLES.map((entry) => ({
                  id: entry.id,
                  label: labels.styleName(entry.id),
                  groupId: entry.section,
                }))}
                groups={OFFICE_SMARTART_STYLE_SECTIONS.map((section) => ({
                  id: section,
                  label: labels.styleSection(section),
                  itemIds: SMARTART_STYLES.filter(
                    (entry) => entry.section === section,
                  ).map((entry) => entry.id),
                }))}
                labels={{ more: labels.quickOverflow }}
                menuTile="art"
                menuGrid={SMARTART_MENU_GRID}
                quickIds={SMARTART_STYLES.map((entry) => entry.id)}
                renderPreview={(item, surface) => (
                  <span
                    aria-hidden="true"
                    // The overflow sheet is PowerPoint's SmartArt Styles
                    // gallery: a large unframed thumbnail per style; the
                    // inline strip keeps its 54×44 slot.
                    style={
                      surface === "menu"
                        ? { display: "block", height: SMARTART_MENU_THUMB[1], width: SMARTART_MENU_THUMB[0] }
                        : { display: "block", height: 44, width: 54 }
                    }
                    // The per-style hook. It rides the INLINE surface, which
                    // lists all fourteen and is always in the DOM; the overflow
                    // menu keeps the gallery's own `…-gallery-menu-<id>` ids, so
                    // the two never collide while the menu is open.
                    {...(surface === "inline"
                      ? { "data-testid": surfaceId(`style-${item.id}`) }
                      : {})}
                    data-style-node-fill={diagram.colors?.nodeFill}
                  >
                    <OfficeSmartArtStylePreview
                      styleId={item.id}
                      kind={diagram.kind}
                      style={diagram.colors}
                      renderShape={renderStyleShape}
                      // The thumbnail's own box (not a 16:9 slide scaled
                      // down): the preview lays the diagram out with this
                      // aspect and fits it, centred, to the box.
                      bounds={
                        surface === "menu"
                          ? { width: SMARTART_MENU_THUMB[0] * 9525, height: SMARTART_MENU_THUMB[1] * 9525 }
                          : { width: 54 * 9525, height: 44 * 9525 }
                      }
                    />
                  </span>
                )}
                testIdPrefix={surfaceId("style-gallery")}
                value={diagram.styleId}
                onPreviewChange={() => undefined}
                onCommit={(id) => diagram.onChangeStyle(id)}
              />
            </OfficeHomeCommandRow>
          </OfficeHomeRibbonGroup>
          {diagram.onResizeShape ? (
            <OfficeHomeRibbonGroup
              ribbon={ribbon}
              label={labels.shapesGroup ?? ""}
              size="standard"
              testId={ribbonId("shapes-group")}
            >
              <OfficeHomeCommandStack>
                <OfficeHomeCommand
                  icon={<Maximize2 size={18} aria-hidden="true" />}
                  label={labels.larger ?? ""}
                  disabled={!diagram.writable || !diagram.canResizeShape}
                  disabledReason={unavailable(
                    diagram.writable && Boolean(diagram.canResizeShape),
                  )}
                  onClick={() => diagram.onResizeShape?.("larger")}
                  testId={ribbonId("larger")}
                />
                <OfficeHomeCommand
                  icon={<Minimize2 size={18} aria-hidden="true" />}
                  label={labels.smaller ?? ""}
                  disabled={!diagram.writable || !diagram.canResizeShape}
                  disabledReason={unavailable(
                    diagram.writable && Boolean(diagram.canResizeShape),
                  )}
                  onClick={() => diagram.onResizeShape?.("smaller")}
                  testId={ribbonId("smaller")}
                />
              </OfficeHomeCommandStack>
            </OfficeHomeRibbonGroup>
          ) : null}
          <OfficeHomeRibbonGroup
            ribbon={ribbon}
            className={cls("resetGroup")}
            label={labels.resetGroupLabel}
            size="standard"
            testId={ribbonId("reset-group")}
          >
            {/* One row: the group is a column, so Reset Graphic and Convert
                left as bare children stacked on top of each other and Convert
                fell out of the band. */}
            <OfficeHomeCommandRow>
              <OfficeHomeCommand
                variant="large"
                icon={<OfficeSemanticIcon semanticId="image.reset" size={24} />}
                label={labels.resetGraphic}
                disabled={!diagram.writable}
                disabledReason={writableReason()}
                onClick={() => diagram.onReset()}
                testId={ribbonId("reset")}
              />
              {diagram.onConvertToShapes || diagram.onConvertToText ? (
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <OfficeHomeMenuTrigger
                      variant="large"
                      hasMenuIndicator
                      icon={<Shapes size={24} aria-hidden="true" />}
                      label={labels.convert ?? ""}
                      displayLabel={labels.convert}
                      disabled={!diagram.writable}
                      disabledReason={writableReason()}
                      testId={ribbonId("convert")}
                    />
                  </DropdownMenuTrigger>
                  <RibbonDropdownMenuContent
                    align="start"
                    data-testid={ribbonId("convert-menu")}
                  >
                    {diagram.onConvertToText ? (
                      <DropdownMenuItem
                        onSelect={() => diagram.onConvertToText?.()}
                        data-testid={ribbonId("convert-to-text")}
                      >
                        {labels.convertToText}
                      </DropdownMenuItem>
                    ) : null}
                    {diagram.onConvertToShapes ? (
                      <DropdownMenuItem
                        onSelect={() => diagram.onConvertToShapes?.()}
                        data-testid={ribbonId("convert-to-shapes")}
                      >
                        {labels.convertToShapes}
                      </DropdownMenuItem>
                    ) : null}
                  </RibbonDropdownMenuContent>
                </DropdownMenu>
              ) : null}
            </OfficeHomeCommandRow>
          </OfficeHomeRibbonGroup>
        </>
      ) : null}
      {diagram ? (
        <>
          <OfficeHomeRibbonGroup
            ribbon={ribbon}
            className={cls("propertiesGroup")}
            label={labels.propertiesGroup}
            size="standard"
            testId={ribbonId("properties-group")}
          >
            <OfficeHomeCommand
              variant="large"
              icon={<OfficeSemanticIcon semanticId="object.altText" size={24} />}
              label={labels.altText}
              disabled={!diagram.writable}
              disabledReason={writableReason()}
              onClick={() =>
                diagram.onAltText(
                  document.activeElement instanceof HTMLElement
                    ? document.activeElement
                    : null,
                )
              }
              testId={ribbonId("alt-text")}
            />
          </OfficeHomeRibbonGroup>
        </>
      ) : null}
    </>
  );
}
