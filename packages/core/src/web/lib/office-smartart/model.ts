/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The SmartArt MODEL: what a stored diagram member says about itself, and the
 * constants every host reads -- the layout kinds, the member identity and
 * marker, the picture-slot custody rules, the outline depth, the default
 * colours. No geometry is computed here.
 *
 * The layout engine (`./engine`, which re-exports all of this) runs on the
 * server (`engine-host/engines/smartart.ts`); the browser draws the members a
 * page stores and reads them through this module, and asks the server for
 * every new layout (`./remote`). Keep layout algorithms out of this file:
 * `scripts/check-engine-client-graph.mjs` proves the engine is not in the
 * browser's module graph, and this module is how the browser avoids needing it.
 */

import type {
  SmartArtFillGradient,
  SmartArtScene3d,
  SmartArtShape3d,
  SmartArtShapeShadow,
  SmartArtStyleId,
  SmartArtTextShape3d,
} from "./styles-model";














export const SMARTART_CONVERT_ONLY_PICTURE_KINDS = [
  "picture-lineup",
  "picture-grid",
  "bending-picture-blocks",
  "bending-picture-caption",
  "bending-picture-semi-transparent-text",
  "picture-strips",
  "bending-picture-caption-list",
  "titled-picture-blocks",
  "framed-text-picture",
  "vertical-picture-list",
  "vertical-picture-accent-list",
  "bending-picture-accent-list",
  // §5.46: the last nine, on the same three builders plus one for the only
  // layout whose slot is at Level 2.
  "captioned-pictures",
  "snapshot-picture-list",
  "title-picture-lineup",
  "picture-accent-blocks",
  "picture-accent-list",
  "alternating-picture-circles",
  "continuous-picture-list",
  "horizontal-picture-list",
  "titled-picture-accent-list",
  // §5.51: the last seven picture layouts in the corpus. Six of them are in
  // Office's own 图片 / Picture tab; `radial-list` is the exception and is
  // filed where Office files it (Relationship), because a picture SLOT and a
  // gallery TAB are two different facts and this list is about the slot.
  "bubble-picture-list",
  "hexagon-cluster",
  "spiral-picture",
  "ascending-picture-accent-process",
  "accented-picture",
  "circle-picture-hierarchy",
  "radial-list",
  // Office's ICON layouts (2018, category `icon`): each item's circle or
  // square is an `iconRect` picture placeholder (`blipPhldr`) one `ch` frame
  // deep, which PowerPoint fills with an SVG icon and its PNG fallback and
  // keeps on the item's `iconRect` presentation point. PowerPoint's Choose a
  // SmartArt Graphic dialog has no Icon tab -- these reach a deck through
  // Designer and Convert to SmartArt -- so they are convert-only like the
  // rest of this list.
  "icon-vertical-solid-list",
  "icon-circle-list",
  "icon-circle-label-list",
  
  // `iconRect` one `ch` frame deep, presenting nothing. Icon Label
  // Description List and its centred twin draw Level 2 text too (`desTx`,
  // `presOf axis="des"`); Icon Leaf Label List captions Level 1 only.
  "icon-label-description-list",
  "centered-icon-label-description-list",
  "icon-leaf-label-list",
] as const;

export const SMARTART_PICTURE_KINDS = [
  "picture-caption-list",
  "alternating-picture-blocks",
  "circular-picture-callout",
  "picture-accent-process",
  ...SMARTART_CONVERT_ONLY_PICTURE_KINDS,
] as const;

/**
 * Office's FIXED-CAPACITY processes: layouts whose Microsoft description names
 * a hard number of Level 1 items and states that the rest is not drawn
 * (SMARTART-IMPORT.md §5.34). They are convert-only — an imported diagram
 * becomes one, the insert gallery never offers one — and the capacity itself
 * is stated once in `./catalog` rather than here.
 */
export const SMARTART_CAPPED_PROCESS_KINDS = [
  "upward-arrow",
  "descending-process",
  "phased-process",
  "sub-step-process",
  "increasing-arrows-process",
  "increasing-circle-process",
  "pie-process",
] as const;









export const SMARTART_OFFICE_LIST_KINDS = [
  "vertical-bullet-list",
  "horizontal-bullet-list",
  "stacked-list",
  "lined-list",
  "vertical-curved-list",
  "trapezoid-list",
  "alternating-hexagons",
  "square-accent-list",
  "vertical-block-list",
  "vertical-accent-list",
  "table-list",
  "vertical-circle-list",
] as const;













export const SMARTART_OFFICE_DEFINITION_KINDS = [
  "arrow-ribbon",
  "block-cycle",
  "horizontal-labeled-hierarchy",
  "detailed-process",
  "plus-and-minus",
  "reverse-list",
  "circular-bending-process",
  "counterbalance-arrows",
  // Measured against PowerPoint's own drawings of three outlines each
  
  "basic-pie",
  "circle-arrow-process",
  "circle-relationship",
  "horizontal-multi-level-hierarchy",
  "opposing-ideas",
  "nested-target",
  "vertical-arrow-list",
  "basic-hierarchy",
  "process-arrows",
  "radial-venn",
  "closed-chevron-process",
  "hierarchy-list",
  "basic-equation",
  
  // parity on its own definition: at most 0.02 of the frame on every pure
  // PowerPoint drawing of it (truth-more, the rest decks), fonts included.
  "diverging-arrows",
  "converging-arrows",
  "circle-accent-timeline",
  "multidirectional-cycle",
  "vertical-equation",
  "table-hierarchy",
  "alternating-flow",
  "process-list",
  "chevron-list",
  "converging-radial",
  "diverging-radial",
  "radial-cluster",
  "random-to-result-process",
  "stacked-venn",
  "step-down-process",
  "target-list",
  
  // each carry a title on their sibling transition (`OutlineItem.title`).
  "name-and-title-org",
] as const;

/**
 * Office layouts drawn only from their own definition that PowerPoint's
 * Choose a SmartArt Graphic dialog files under Office.com (`officeonline`) --
 * a tab MRAI does not ship -- so they are convert-only: a diagram of one
 * converts on open and keeps the layout part its own package holds, and
 * nothing authors one from nothing. Interconnected Rings (`rings+Icon`),
 * Hexagon Radial, Tabbed Arc (a `diagrams.loki3.com` layout), Tab List and
 * Chevron Accent Process (`chevronAccent+Icon`).
 * Each definition presents descendants (`presOf axis="desOrSelf"` or
 * `"ch desOrSelf"`), so Level 2 text is drawn.
 */
export const SMARTART_OFFICE_ONLINE_KINDS = [
  "interconnected-rings",
  "hexagon-radial",
  "tabbed-arc",
  "tab-list",
  "chevron-accent-process",
] as const;

/**
 * Office's 2016-2017 layouts with no picture placeholder, which reach a deck
 * through PowerPoint Designer and Convert to SmartArt, never through the
 * Choose a SmartArt Graphic dialog (Basic Linear Process Numbered, Vertical
 * Solid Action List, Drop Pin Timeline). Convert-only like the icon layouts:
 * a diagram of one converts on open, drawn from its own definition, and
 * keeps the layout part its package holds. Each presents its descendants
 * (`presOf axis="desOrSelf"` / `"des"`), so Level 2 text is drawn.
 */
export const SMARTART_OFFICE_DESIGNER_KINDS = [
  "basic-linear-process-numbered",
  "vertical-solid-action-list",
  "drop-pin-timeline",
] as const;

/**
 * A diagram laid out from the layout definition its OWN file carries, for a
 * layout MRAI does not stock (SMARTART-GAP.md §8.6): every member records the
 * file's urn (`smartArtLayoutUrn`) and the part's digest
 * (`smartArtLayoutSha256`), and every rebuild is laid out from that very part
 * again. Converted only when its drawing matches PowerPoint's; never offered
 * from nothing, and with no MRAI builder of its own to fall back on.
 */
export const SMARTART_OFFICE_DEFINITION_KIND = "office-definition" as const;

/** The SmartArt layouts this generator produces. */
export type SmartArtKind =
  | (typeof SMARTART_PICTURE_KINDS)[number]
  | (typeof SMARTART_OFFICE_LIST_KINDS)[number]
  | (typeof SMARTART_OFFICE_DEFINITION_KINDS)[number]
  | (typeof SMARTART_OFFICE_ONLINE_KINDS)[number]
  | (typeof SMARTART_OFFICE_DESIGNER_KINDS)[number]
  | typeof SMARTART_OFFICE_DEFINITION_KIND
  | "list"
  | "vertical-box-list"
  | "grouped-list"
  | "card-grid"
  | "step-list"
  | "process"
  | "chevron-process"
  | "vertical-process"
  | "alternating-process"
  | "timeline"
  | "step-up-process"
  | (typeof SMARTART_CAPPED_PROCESS_KINDS)[number]
  | "hierarchy"
  | "horizontal-hierarchy"
  | "labeled-tree"
  | "half-circle-org"
  | "cycle"
  | "continuous-cycle"
  | "radial-cycle"
  | "segmented-cycle"
  | "pyramid"
  | "relationship-venn"
  | "linear-venn"
  | "opposing-arrows"
  | "balance"
  | "funnel"
  | "target"
  | "radial"
  | "gears"
  | "basic-matrix"
  | "titled-matrix"
  | "grid-matrix"
  | "cycle-matrix"
  | "inverted-pyramid"
  | "segmented-pyramid"
  | "pyramid-list";

export const SMARTART_KINDS: readonly SmartArtKind[] = [
  "list",
  "vertical-box-list",
  "grouped-list",
  "card-grid",
  "step-list",
  "process",
  "chevron-process",
  "vertical-process",
  "alternating-process",
  "timeline",
  "step-up-process",
  "hierarchy",
  "horizontal-hierarchy",
  "labeled-tree",
  "half-circle-org",
  "cycle",
  "continuous-cycle",
  "radial-cycle",
  "segmented-cycle",
  "pyramid",
  "relationship-venn",
  "linear-venn",
  "opposing-arrows",
  "balance",
  "funnel",
  "target",
  "radial",
  "gears",
  "basic-matrix",
  "titled-matrix",
  "grid-matrix",
  "cycle-matrix",
  "inverted-pyramid",
  "segmented-pyramid",
  "pyramid-list",
  ...SMARTART_CAPPED_PROCESS_KINDS,
  ...SMARTART_PICTURE_KINDS,
  ...SMARTART_OFFICE_LIST_KINDS,
  ...SMARTART_OFFICE_DEFINITION_KINDS,
  ...SMARTART_OFFICE_ONLINE_KINDS,
  ...SMARTART_OFFICE_DESIGNER_KINDS,
  SMARTART_OFFICE_DEFINITION_KIND,
];

const SMARTART_KIND_SET: ReadonlySet<string> = new Set(SMARTART_KINDS);

/** Host-neutral alias for {@link SmartArtKind} (the catalog's PPT legacy id). */
export type OfficeSmartArtKind = SmartArtKind;

/**
 * The role a scene node plays. `node` and `label` are the two halves of ONE
 * outline item (exactly one of each per item); `connector` is structural and
 * never acquires item identity.
 */
export type SmartArtBlockRole = "node" | "label" | "connector";
export type OfficeSmartArtNodeRole = SmartArtBlockRole;

/**
 * Durable identity every generated node carries. A host persists these fields
 * verbatim so Add / Delete / Promote / Demote and picture custody survive a
 * relayout; they are never inferred from visual similarity on import.
 */
export type OfficeSmartArtNodeIdentity = {
  /** One id shared by every node of one diagram. */
  groupId: string;
  /** 0 connectors · 1 nodes · 2 labels (paint order within the group). */
  zIndex: number;
  smartArtKind?: OfficeSmartArtKind;
  smartArtRole?: OfficeSmartArtNodeRole;
  smartArtItemId?: string;
  smartArtItemIndex?: number;
  smartArtLevel?: number;
  /**
   * The item whose TITLE this text member is (its index, like
   * `smartArtItemIndex`): Name and Title Organization Chart's box under each
   * name, which Office keeps as the words of the item's following sibling
   * transition (`dgm:pt type="sibTrans"`). Only on `connector`-role text
   * members; its text is the item's title ({@link readSmartArtOutline}).
   */
  smartArtTitleIndex?: number;
  /**
   * The item a shape Office draws FOR an item without carrying its words
   * belongs to (its index, like `smartArtItemIndex`): Vertical Circle List's
   * big and small circles, an accent beside an item. Its presentation point's
   * `presAssocID` is the item's data point, so Office deletes (and cuts) that
   * item when the shape is deleted. Only on `connector`-role members, which
   * every host draws as the engine's own; a true transition (`sibTrans`,
   * `parTrans`) carries none and cannot be deleted.
   */
  smartArtOwnerIndex?: number;
  /**
   * An organization chart box's Layout (SmartArt Design > Layout: Office's
   * `hierBranch`, {@link SmartArtHierBranch}), on the item's `node` member:
   * how the box's subordinates hang below it. Absent: the layout's own.
   */
  smartArtHierBranch?: SmartArtHierBranch;
  /**
   * The one quick style the whole diagram carries, on every member (including
   * images) exactly like {@link OfficeSmartArtNodeIdentity.smartArtKind}, so a
   * restyle never has to infer a style from how a member happens to look.
   */
  smartArtStyle?: SmartArtStyleId;
  /**
   * SmartArt Design > Right to Left: the diagram is laid out MIRRORED, on
   * every member exactly like {@link OfficeSmartArtNodeIdentity.smartArtStyle}.
   * Only `"rtl"` is ever stored; left to right is stated by omission.
   *
   * It is the graphic's spatial direction, not its order: Office keeps it as
   * the layout variable `dgm:dir val="rev"` (ISO/IEC 29500-1 §21.4.5.4
   * `presLayoutVars`, §21.4.7.26 `ST_Direction`) on the diagram's root
   * presentation point, and its own layout definitions answer it with a
   * mirrored arrangement (`process1`'s `linDir="fromR"`, `chevron1`'s
   * `rot="180"` chevrons) while the Text Pane keeps its order.
   */
  smartArtDirection?: "rtl";
  /**
   * SmartArt Design > Change Colors: the colour VARIATION the diagram wears,
   * as Office's own id (`accent1_3`, `colorful2`, … — the last segment of the
   * `dgm:colorsDef` urn, `./colors`), on every member exactly like
   * {@link OfficeSmartArtNodeIdentity.smartArtStyle}. Absent is Office's
   * default `accent1_2` (Colored Fill - Accent 1) and every page stored before
   * the field existed. The export names it in the data part's `csTypeId` and
   * writes PowerPoint's own definition as the diagram's colours part, so
   * Office paints the diagram from the deck's theme the way it would itself.
   */
  smartArtColors?: string;
  /**
   * The Office layout this diagram IS when that is an ALIAS drawn from its
   * own definition (`ALIGNED_NATIVE_SMARTART_ALIASES` in `./native-layouts`):
   * Office's full `uniqueId` urn (`…/layout/hProcess9`), on every member
   * exactly like {@link OfficeSmartArtNodeIdentity.smartArtKind}, which still
   * names the catalog layout it stands for. Every rebuild lays the diagram
   * out from this urn; a Change Layout to another catalog layout drops it.
   * Absent on every diagram of a stocked layout.
   */
  smartArtLayoutUrn?: string;
  /**
   * The fill and outline on this node are the AUTHOR's, not the catalog
   * palette's — set only on a `node`, and only ever `true`.
   *
   * The export can carry a member's colours into Office as a per-point
   * `dgm:spPr`, but only for colours it can tell apart from the palette, and
   * the only tell it had was disagreement between the diagram's own nodes
   * (`SMARTART-IMPORT.md` §5.3). A diagram whose every node wears ONE chosen
   * colour is indistinguishable from the palette by that rule, so it exported
   * without its colours. The conversion on open knows better — a data point is
   * painted only when somebody chose the paint — so it says so here, the host
   * projections persist it, and the next export writes the colours again.
   */
  smartArtAuthorPaint?: true;
  /**
   * SmartArt Format > Change Shape: the shape the AUTHOR gave this item, set
   * only on a `node` (the editor's shape kind, equal to its `shape`). A
   * rebuild draws the item in it instead of the layout's own; Office keeps it
   * as `a:prstGeom` on the presentation point that draws the shape, which is
   * where the export writes it and the import reads it back. Reset Shape and
   * Reset Graphic drop it.
   */
  smartArtAuthorShape?: string;
  /**
   * SmartArt Format on this one shape beyond its colours (PPT-A010): its
   * EFFECTS (`effect`, `shadow`, `shapeShadow`) and a Shape Styles gradient
   * are the AUTHOR's, set only on a `node`. Office keeps them on the item's
   * data point, where a stated `a:effectLst` (even an empty one) replaces the
   * style's effects; a rebuild draws the item with them again, Reset Shape
   * and Reset Graphic drop them.
   */
  smartArtAuthorFormat?: true;
  /**
   * This item is an ASSISTANT (`dgm:pt type="asst"`): Office's Add Assistant
   * on an organization chart. On the item's node and label only, only ever
   * `true`, set from the outline by `applySmartArtOutline`.
   */
  smartArtAssistant?: true;
  /**
   * The layout does not draw this item: Office's "unused text does not
   * appear, but remains available if you switch layouts". The item stays in
   * the outline (the Text Pane lists it, a layout switch draws it, the export
   * writes its data point) through this member, its node and label sized to
   * the drawn diagram's box, but no host paints it, hit-tests it or selects
   * it, and the cached drawing has no shape for it. Set only by the engine's
   * build (`buildSmartArt`), on an item's node and label -- and on a title
   * box (`smartArtTitleIndex`) whose words the layout does not draw -- only
   * ever `true`.
   */
  smartArtHidden?: true;
  /**
   * Where Office paints this member among the diagram's shapes, lowest first:
   * an Office scene's layout order (its nodes as the definition states them,
   * `zOrderOff` applied). Absent from MRAI's own layouts, which paint by
   * `zIndex`. An order to paint in, nothing else reads it.
   */
  smartArtPaint?: number;
};

/** Geometry in EMU (914,400 per inch), relative to the diagram frame origin. */
export type OfficeSmartArtNodeGeometry = {
  left: number;
  top: number;
  width: number;
  height: number;
  /** Degrees, clockwise; absent means unrotated. */
  rotation?: number;
};

type OfficeSmartArtNodeBase = OfficeSmartArtNodeIdentity &
  OfficeSmartArtNodeGeometry & { id: string };

/**
 * Label character formatting. Deliberately the subset every host and the PPTX
 * authored-group writer accept (`_STYLE` in
 * `apps/api/app/services/file_adapters/pptx/smartart_insertion.py`).
 */
export type OfficeSmartArtTextStyle = {
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** Shown in capitals, the words untouched (`a:defRPr@cap`, a layout's own list style). */
  caps?: "all" | "small";
  align?: "left" | "center" | "right" | "justify" | "distribute";
  verticalAlign?: "top" | "middle" | "bottom";
  /** Paragraph spacing as a multiple of a single line (PowerPoint's SmartArt 90%, or the author's). */
  lineSpacing?: number;
  /**
   * The unit `fontSize` is stated in: `"px"` for the Office-drawn scene's
   * labels (the slide's design px). Absent, points (the catalog generators),
   * unless the label carries the scene's 90% spacing, which is how a label
   * stored before this field said px.
   */
  sizeUnit?: "px";
};

/**
 * The line spacing every label of the Office-drawn scene (`ms-engine/scene`)
 * carries: PowerPoint's 90% of the single line (`a:lnSpc`), which the
 * engine's fit and every host's drawing of SmartArt words read. The catalog
 * generators set none.
 */
export const OFFICE_SCENE_LABEL_LINE_SPACING = 0.9;

/**
 * Whether a label's `fontSize` is design px rather than points. The
 * Office-drawn scene states its labels in the slide's design px (4/3 of a
 * point, Office's 19 pt as 25.5); the catalog generators state points (their
 * lines are `fontSize * EMU_PER_PT` tall). A host that speaks px converts only
 * the latter -- the worksheet multiplied both by 4/3 and drew Excel's 18 pt
 * Basic Cycle labels at 25.5 pt.
 */
export function smartArtLabelSizeIsPx(style: OfficeSmartArtTextStyle | undefined): boolean {
  if (style?.sizeUnit === "px") return true;
  // A label stored before `sizeUnit` said px by the scene's 90% spacing; the
  // author's own spacing (`smartArtAuthorLineSpacing`) comes with the unit.
  return style?.sizeUnit === undefined && style?.lineSpacing === OFFICE_SCENE_LABEL_LINE_SPACING;
}

/** Bounded custom geometry in the shape's own 0..100000 space. */
export type OfficeSmartArtPathCommand =
  | { type: "moveTo"; x: number; y: number }
  | { type: "lineTo"; x: number; y: number };

export type OfficeSmartArtCustomPath = {
  version: 1;
  closed: boolean;
  commands: readonly OfficeSmartArtPathCommand[];
};

/** A vector member: an auto-shape preset, or explicit `customPath` geometry. */
export type OfficeSmartArtShapeNode = OfficeSmartArtNodeBase & {
  type: "shape";
  /** Auto-shape preset id; `custom-geometry` defers to {@link OfficeSmartArtShapeNode.customPath}. */
  shape: string;
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  /** Percent, 0..100; absent is fully opaque. */
  opacity?: number;
  /**
   * Drawn mirrored about its own vertical axis (DrawingML `a:xfrm@flipH`,
   * applied before `rotation`). Only a Right to Left diagram's direction-
   * bearing shapes carry it — its arrows and chevrons point the other way —
   * and only ever as `true`.
   */
  flipH?: boolean;
  /**
   * The preset's adjust values as Office's layout states them (`a:avLst` by
   * handle name: `adj`, `adj1` …): a Process arrow's shaft, a Grouped List
   * box's corners. Absent is the preset's own defaults.
   */
  adjustments?: Readonly<Record<string, number>>;
  customPath?: OfficeSmartArtCustomPath;
  /**
   * Quick-style paint, exactly the writer's member contract. `fillGradient`'s
   * `angle` is in DrawingML degrees (90 = top→bottom, `a:lin ang="5400000"`);
   * the canvas derives its own CSS-side companions at render time rather than
   * storing them, because the authored-group writer rejects a member carrying
   * any field outside this set. `shadow` is the renderer's on/off gate for the
   * parameters in `shapeShadow`, and is part of that contract already.
   */
  fillGradient?: SmartArtFillGradient;
  shadow?: boolean;
  shapeShadow?: SmartArtShapeShadow;
  shape3d?: SmartArtShape3d;
  scene3d?: SmartArtScene3d;
};

/** One outline item's text, laid out as a free-floating centered box. */
export type OfficeSmartArtTextNode = OfficeSmartArtNodeBase & {
  type: "text";
  /** Required so a host positions the label at its own geometry. */
  kind: "textbox";
  text: string;
  style: OfficeSmartArtTextStyle;
  /** 3-D on the text body itself (style `3d9` only); the canvas may ignore it. */
  textShape3d?: SmartArtTextShape3d;
  /** The diagram camera, identical on every member (styles `3d5`–`3d9`). */
  scene3d?: SmartArtScene3d;
  /**
   * How far apart the label's Chinese lines stand, as a multiple of its font
   * (the Office scene's, PowerPoint's 90% spacing included): the line its
   * words were fitted at. Drawn only; absent on a label without Chinese words.
   */
  smartArtLinePitch?: number;
  /** The label's colour as the theme colour its point states (`a:schemeClr`), authenticated by `rgb`. */
  smartArtTextColorScheme?: { scheme: string; mods: readonly (readonly [string, number])[]; rgb: string };
  






  smartArtFirstLineIndent?: number;
  





  smartArtTextRtl?: true;
};

/** A picture member of a Picture layout; `src` starts as the placeholder. */
export type OfficeSmartArtImageNode = OfficeSmartArtNodeBase & {
  type: "image";
  src: string;
  alt: string;
  borderColor?: string;
  borderWidth?: number;
  cornerRadius?: number;
  /** An oval slot drawn oval at any size (Crop to Shape > Oval). */
  clipShape?: "ellipse";
  /**
   * The colour Office paints the slot while it holds no picture: the shape's
   * own fill (Circular Picture Callout's grey hub, its white caption on it).
   * Drawn only; the file's slot is the layout's.
   */
  smartArtSlotFill?: string;
  /**
   * The slot's own geometry where it is not a plain box or an oval: Hexagon
   * Cluster's slots are hexagons, Picture Caption List's rounded. Drawn only.
   */
  smartArtSlotShape?: "hexagon" | "roundRect";
  /**
   * The slot holds the DIAGRAM's own picture, not an item's: its layout node
   * (`presName`) is keyed to the document point -- Radial List's centre
   * `visible` (readings row 1005). The file keeps that picture, and its line,
   * on that presentation point.
   */
  smartArtDocSlot?: string;
  /**
   * The line that slot's presentation point states of its own (the author's
   * Shape Outline on it): `#rrggbb` or `"none"`, or a theme colour (`scheme`
   * and its modifiers, drawn as `rgb`), and its width. Written back on that
   * point as stated; absent, the point states none.
   */
  smartArtDocSlotLine?: {
    color?: string;
    scheme?: string;
    mods?: readonly (readonly [string, number])[];
    rgb?: string;
    widthPt?: number;
  };
};

/**
 * ONE scene node. The generator emits nothing else, and every field is chosen
 * so the node is assignable to the presentation block model without a cast and
 * survives the authored-group PPTX writer's member validation.
 */
export type OfficeSmartArtNode =
  | OfficeSmartArtShapeNode
  | OfficeSmartArtTextNode
  | OfficeSmartArtImageNode;

/**
 * A block the engine READS: either a node it generated, or a host's own richer
 * block (PPT `PresentationBlock`, a Docs drawing record, …). Functions that
 * return blocks are generic over this so a host gets its own type back.
 */
export type OfficeSmartArtHostNode = {
  id: string;
  type: string;
  [key: string]: unknown;
};

/** The loosest read surface: a host block or a plain persisted record. */
export type OfficeSmartArtBlockLike =
  | OfficeSmartArtHostNode
  | Record<string, unknown>;

export const SMARTART_PICTURE_PLACEHOLDER_SRC =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAwCAYAAAChS3wfAAAAp0lEQVR42u3aMQqDMBgGUHfP4u3cvYSzSy/RixVKERTBrTURGxGTN3yQKeF7ZPgDqfrHcyo5FYB18f6MRQUAAAAAsgJ4DXV5AEvprWQPECq/BwEAgBsD7CkfQ3ADAAAwB5gEvQUAAAgCNF37lRSH/9r37ABIBfAvwhXlAaQGOIpwVflTAO4WAAAAAAAAAAAAAAAAAAAAAAAAAAAAxD9L5wRw6Ld4CQAzW09A/zw6u8wAAAAASUVORK5CYII=";

/**
 * How a picture slot that holds no picture is DRAWN: one vector glyph every
 * host shows in place of the stored `SMARTART_PICTURE_PLACEHOLDER_SRC` (a
 * 64 x 48 bitmap, which stays the stored token the files and the server
 * recognise). Stretched over the slot, the bitmap came out blurred and
 * blocky; this draws sharp at any size, its picture icon centred at its own
 * proportions however the slot is stretched (Office's empty picture
 * placeholder: a pale slot and a line picture icon).
 */
const PLACEHOLDER_ICON =
  '<rect x="2.5" y="4.5" width="19" height="15" rx="2.2"/>' +
  '<circle cx="8.6" cy="9.6" r="1.8"/>' +
  '<path d="M3.2 17.6l5.4-5.2 3.6 3.4 3.3-3.1 5.6 5.1"/>';
const placeholderSvg = (background: boolean, fill = "#f1f5f9", stroke = "#94a3b8") =>
  "data:image/svg+xml;charset=utf-8," +
  encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%">' +
      (background ? `<rect width="100%" height="100%" fill="${fill}"/>` : "") +
      `<svg x="${background ? 34 : 0}%" y="${background ? 34 : 0}%" width="${background ? 32 : 100}%" height="${background ? 32 : 100}%" viewBox="0 0 24 24" preserveAspectRatio="xMidYMid meet" fill="none" stroke="${stroke}" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">` +
      PLACEHOLDER_ICON +
      "</svg></svg>",
  );
/** The empty slot as the slide draws it: the pale slot and its icon. */
export const SMARTART_PICTURE_PLACEHOLDER_DISPLAY_SRC = placeholderSvg(true);
/** The icon alone, for a host that paints the slot itself (Docs). */
export const SMARTART_PICTURE_PLACEHOLDER_ICON_SRC = placeholderSvg(false);

/**
 * What a picture's `src` is drawn with: the vector slot for the stored
 * placeholder token, in the colour Office paints the empty slot when the
 * member states one (`smartArtSlotFill`), its icon light on a dark slot.
 */
export function smartArtPictureDisplaySrc(src: string, slotFill?: unknown): string;
export function smartArtPictureDisplaySrc(src: string | undefined, slotFill?: unknown): string | undefined;
export function smartArtPictureDisplaySrc(src: string | undefined, slotFill?: unknown): string | undefined {
  if (src !== SMARTART_PICTURE_PLACEHOLDER_SRC) return src;
  if (typeof slotFill !== "string" || !/^#[0-9a-f]{6}$/i.test(slotFill)) return SMARTART_PICTURE_PLACEHOLDER_DISPLAY_SRC;
  const [r, g, b] = [1, 3, 5].map((at) => parseInt(slotFill.slice(at, at + 2), 16));
  const light = 0.299 * r + 0.587 * g + 0.114 * b > 150;
  return placeholderSvg(true, slotFill.toLowerCase(), light ? "#475569" : "#f8fafc");
}

/**
 * What a picture's `src` is drawn with OUTSIDE the editing canvas -- the show,
 * the slide rail and every printed or exported page: an empty slot is the
 * colour Office paints it (`smartArtSlotFill`) and nothing else. PowerPoint
 * shows its "insert picture" icon only while editing; its show, its pictures
 * and its PDF draw the bare slot, under every style (truth office-truth/probes/
 * sa36-s22slot: Picture Strips and Picture Caption List under simple1/2/4/5,
 * sa36-3dcam under 3d1-3d9; TRIAGE S22). A slot with no fill of its own is
 * drawn as on the canvas.
 */
export function smartArtPictureOutputSrc(src: string, slotFill?: unknown): string;
export function smartArtPictureOutputSrc(src: string | undefined, slotFill?: unknown): string | undefined;
export function smartArtPictureOutputSrc(src: string | undefined, slotFill?: unknown): string | undefined {
  if (src !== SMARTART_PICTURE_PLACEHOLDER_SRC) return src;
  if (typeof slotFill !== "string" || !/^#[0-9a-f]{6}$/i.test(slotFill)) return smartArtPictureDisplaySrc(src, slotFill);
  return (
    "data:image/svg+xml;charset=utf-8," +
    encodeURIComponent(
      `<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%"><rect width="100%" height="100%" fill="${slotFill.toLowerCase()}"/></svg>`,
    )
  );
}

/**
 * The neutral SLOT a host draws for a picture point that holds no picture, for
 * the hosts whose shape model carries no picture of its own: the worksheet
 * (`lib/spreadsheet/smart-art.ts`) and the Word shape group
 * (`lib/office-smartart/document-shapes.ts`) both project an `image` node as a
 * plain shape in these colours. One pair of constants so an imported picture
 * layout looks the same in both hosts (docs/office-graphics/SMARTART-IMPORT.md).
 */
export const SMARTART_PICTURE_PLACEHOLDER_FILL = "#f1f5f9";
export const SMARTART_PICTURE_PLACEHOLDER_STROKE = "#94a3b8";

/** Resolve one picture member while its whole authored SmartArt group is selected. */
export function smartArtPictureReplaceTargetId(
  blocks: readonly OfficeSmartArtBlockLike[],
  selectedIds: ReadonlySet<string>,
  primaryId: string | null,
): string | null {
  if (!primaryId) return null;
  const primaryMatches = blocks.filter((block) => block.id === primaryId);
  if (primaryMatches.length !== 1) return null;
  const primary = primaryMatches[0];
  const groupId = primary.groupId;
  const kind = primary.smartArtKind;
  if (
    primary.type !== "image" ||
    primary.smartArtRole !== "node" ||
    typeof groupId !== "string" ||
    !groupId ||
    typeof kind !== "string" ||
    !SMARTART_PICTURE_KINDS.some((candidate) => candidate === kind)
  ) {
    return null;
  }
  const members = blocks.filter((block) => block.groupId === groupId);
  if (
    members.length < 2 ||
    members.length > 256 ||
    selectedIds.size !== members.length ||
    members.some((block) => !selectedIds.has(String(block.id))) ||
    new Set(members.map((block) => block.id)).size !== members.length
  ) {
    return null;
  }
  const items = new Map<
    number,
    { itemId: string; level: number; roles: Set<string> }
  >();
  for (const member of members) {
    const role = member.smartArtRole;
    if (
      member.smartArtKind !== kind ||
      (role !== "node" && role !== "label" && role !== "connector")
    ) {
      return null;
    }
    if (role === "connector") {
      if (
        member.smartArtItemId !== undefined ||
        member.smartArtItemIndex !== undefined ||
        member.smartArtLevel !== undefined
      ) {
        return null;
      }
      continue;
    }
    const itemId = member.smartArtItemId;
    const itemIndex = member.smartArtItemIndex;
    const level = member.smartArtLevel;
    if (
      typeof itemId !== "string" ||
      !itemId ||
      !Number.isInteger(itemIndex) ||
      Number(itemIndex) < 0 ||
      !Number.isInteger(level) ||
      Number(level) < 0 ||
      Number(level) > MAX_SMARTART_OUTLINE_LEVEL
    ) {
      return null;
    }
    const index = Number(itemIndex);
    const pair = items.get(index) ?? {
      itemId,
      level: Number(level),
      roles: new Set<string>(),
    };
    if (
      pair.itemId !== itemId ||
      pair.level !== level ||
      pair.roles.has(role)
    ) {
      return null;
    }
    pair.roles.add(role);
    items.set(index, pair);
  }
  const indices = [...items.keys()].sort((left, right) => left - right);
  return indices.length > 0 &&
    indices.every((index, position) => index === position) &&
    [...items.values()].every(
      (item) =>
        item.roles.size === 2 &&
        item.roles.has("node") &&
        item.roles.has("label"),
    )
    ? primaryId
    : null;
}

export function isSmartArtPictureKind(value: unknown): boolean {
  return (
    typeof value === "string" &&
    SMARTART_PICTURE_KINDS.some((candidate) => candidate === value)
  );
}

export function smartArtPictureNodes(
  blocks: readonly OfficeSmartArtBlockLike[],
) {
  return blocks.filter(
    (block) =>
      block.type === "image" &&
      block.smartArtRole === "node" &&
      isSmartArtPictureKind(block.smartArtKind),
  );
}

export function hasAuthenticatedPictureCustody(
  block: OfficeSmartArtBlockLike,
): boolean {
  return (
    block.src !== SMARTART_PICTURE_PLACEHOLDER_SRC ||
    [
      "source",
      "sourceKind",
      "assetId",
      "mimeType",
      "pixelWidth",
      "pixelHeight",
      "sizeBytes",
      "hasAlpha",
      "exifOrientation",
      "officeGraphicKind",
      "graphicFallbackSrc",
    ].some((key) => block[key] !== undefined)
  );
}

export function smartArtPictureLayoutChangeBlocked(
  blocks: readonly OfficeSmartArtBlockLike[],
  targetKind: SmartArtKind,
): boolean {
  return (
    !isSmartArtPictureKind(targetKind) &&
    smartArtPictureNodes(blocks).some(hasAuthenticatedPictureCustody)
  );
}


export function isSmartArtKind(value: unknown): value is SmartArtKind {
  return typeof value === "string" && SMARTART_KIND_SET.has(value);
}

export interface SmartArtBlockMarker {
  groupId: string;
  kind: SmartArtKind;
  role: SmartArtBlockRole;
  itemId: string | null;
  itemIndex: number | null;
  level: number | null;
}

/**
 * The item Delete and Cut on a SmartArt member take out: the member's own
 * item, or -- for a shape Office draws for an item without its words
 * (`smartArtOwnerIndex`: Vertical Circle List's circles) -- the item it is
 * drawn for, found as that item's node among the graphic's members. Null for
 * a connecting arrow and anything else that belongs to no item, which Office
 * does not delete (Microsoft Support, "Delete a SmartArt graphic or a shape
 * in a SmartArt graphic").
 */
export function smartArtMemberDeletedItemId(
  members: readonly unknown[],
  member: unknown,
): string | null {
  const marker = readSmartArtBlockMarker(member);
  if (!marker) return null;
  if (marker.role !== "connector") return marker.itemId;
  const owner = (member as { smartArtOwnerIndex?: unknown }).smartArtOwnerIndex;
  if (typeof owner !== "number" || !Number.isInteger(owner) || owner < 0) return null;
  for (const candidate of members) {
    const own = readSmartArtBlockMarker(candidate);
    if (own?.groupId === marker.groupId && own.role === "node" && own.itemIndex === owner && own.itemId) {
      return own.itemId;
    }
  }
  return null;
}

/**
 * Whether a label's box lies within its host shape's box (a little slack for
 * the engine's rounding), i.e. the label is the shape's own centred text
 * rather than a callout placed beside it.
 *
 * A host that draws a shape's label INSIDE the shape (a Docs group member, a
 * sheet shape) merges only such a label into its item's shape. The engine
 * also authors labels OUTSIDE their shape -- a timeline's captions above and
 * below the axis, a target's callouts beside each ring -- with the same item
 * index, and merging one of those put its words inside a 12 px dot.
 */
export function smartArtLabelInsideHost(
  label: OfficeSmartArtNode,
  host: OfficeSmartArtNode,
): boolean {
  const values = [
    label.left, label.top, label.width, label.height,
    host.left, host.top, host.width, host.height,
  ];
  if (values.some((value) => !Number.isFinite(value))) return true;
  const slackX = Math.max(1, host.width * 0.02);
  const slackY = Math.max(1, host.height * 0.02);
  return (
    label.left >= host.left - slackX &&
    label.top >= host.top - slackY &&
    label.left + label.width <= host.left + host.width + slackX &&
    label.top + label.height <= host.top + host.height + slackY
  );
}

/**
 * Durable identity for MRAI-authored SmartArt clusters. Imported or ordinary
 * grouped shapes deliberately do not infer this marker from visual similarity.
 */
export function readSmartArtBlockMarker(
  block: unknown,
): SmartArtBlockMarker | null {
  if (!block || typeof block !== "object" || Array.isArray(block)) return null;
  const record = block as Record<string, unknown>;
  const groupId = record.groupId;
  const kind = record.smartArtKind;
  const role = record.smartArtRole;
  if (
    typeof groupId !== "string" ||
    groupId.length === 0 ||
    !isSmartArtKind(kind) ||
    (role !== "node" && role !== "label" && role !== "connector")
  ) {
    return null;
  }
  const itemId =
    typeof record.smartArtItemId === "string" &&
    record.smartArtItemId.length > 0
      ? record.smartArtItemId
      : null;
  const itemIndex = Number.isInteger(record.smartArtItemIndex)
    ? (record.smartArtItemIndex as number)
    : null;
  const level = Number.isInteger(record.smartArtLevel)
    ? (record.smartArtLevel as number)
    : null;
  return { groupId, kind, role, itemId, itemIndex, level };
}

/**
 * A member of an item the layout does not draw
 * ({@link OfficeSmartArtNodeIdentity.smartArtHidden}): every host skips it
 * when painting, hit-testing, selecting and listing objects, and keeps it
 * everywhere the outline is read.
 */
export function isSmartArtHiddenMember(block: unknown): boolean {
  return (
    !!block &&
    typeof block === "object" &&
    (block as Record<string, unknown>).smartArtHidden === true
  );
}

/** The items of these members the layout does not draw, by item id. */
export function smartArtUnusedItemIds(members: readonly unknown[]): Set<string> {
  const ids = new Set<string>();
  for (const member of members) {
    const itemId = isSmartArtHiddenMember(member)
      ? (member as Record<string, unknown>).smartArtItemId
      : undefined;
    if (typeof itemId === "string" && itemId) ids.add(itemId);
  }
  return ids;
}

/**
 * Which way a diagram runs: Office's default left to right (clockwise for a
 * circular layout), or Right to Left (SmartArt Design > Create Graphic).
 */
export type SmartArtDirection = "ltr" | "rtl";

/** The direction one member states ({@link OfficeSmartArtNodeIdentity.smartArtDirection}). */
export function readSmartArtDirection(block: unknown): SmartArtDirection {
  return block &&
    typeof block === "object" &&
    (block as Record<string, unknown>).smartArtDirection === "rtl"
    ? "rtl"
    : "ltr";
}

/**
 * The engine's direction for a value read from a file or a projection:
 * Office's `ST_Direction` (`norm` / `rev`) or the engine's own spelling.
 * Anything else is Office's default.
 */
export function smartArtDirectionFromOffice(value: unknown): SmartArtDirection {
  return value === "rev" || value === "rtl" ? "rtl" : "ltr";
}


/**
 * Deepest outline level the layouts nest. Lives with the generator because
 * `buildSmartArt` refuses a deeper level rather than silently flattening it;
 * `office-smartart/outline` reuses it for the editing caps.
 */
export const MAX_SMARTART_OUTLINE_LEVEL = 3;


/**
 * The frame an insert lays a diagram out in — the editor's own default slide
 * bounds. Every layout's proportions are tuned for this box, so an importer
 * that has to land the diagram in an arbitrary Office frame builds here first
 * and then scales the whole result with {@link fitSmartArtToFrame}.
 */
export const DEFAULT_SMARTART_INSERT_BOUNDS: SmartArtBounds = {
  width: 9_144_000,
  height: 5_143_500,
};

export interface SmartArtBounds {
  width: number;
  height: number;
}

/** Colors + type size the generated nodes / connectors / labels use. */
export interface SmartArtStyle {
  /** Node box fill (six-digit hex). */
  nodeFill: string;
  /** Node box outline (six-digit hex). */
  nodeStroke: string;
  /** Connector segment color (six-digit hex). */
  connector: string;
  /** Label text color (six-digit hex). */
  textColor: string;
  /** Label font size in points. */
  fontSize: number;
}



/** Sensible default styling, matching the editor's default shape palette. */
export const DEFAULT_SMARTART_STYLE: SmartArtStyle = {
  nodeFill: "#4472c4",
  nodeStroke: "#ffffff",
  connector: "#b4bbc3",
  textColor: "#ffffff",
  fontSize: 32,
};


/** One shape's own paint (see `buildSmartArt`'s `nodePaint`). */
export interface SmartArtNodePaint {
  readonly fill?: string;
  readonly stroke?: string;
  /**
   * The shape's opacity in percent, when its paint is translucent: Office's
   * Transparent Gradient Range variations (`accentN_5`) fill at 90 % to 50 %.
   * Absent is opaque, which every other paint is.
   */
  readonly opacity?: number;
}


export const HIERARCHY_SMARTART_KINDS: ReadonlySet<SmartArtKind> = new Set([
  "hierarchy",
  "horizontal-hierarchy",
  "labeled-tree",
  "half-circle-org",
  "grouped-list",
  // The five capped processes whose Microsoft description says Level 2 text is
  // DRAWN ("Level 2 can contain large amounts of text", "unlimited sub-steps",
  // "unlimited Level 2 items"). The other two — Upward Arrow and Descending
  // Process — name only Level 1, so Demote is refused there and the level cap
  // in `./catalog` is the same fact written where a conversion can read it.
  "phased-process",
  "sub-step-process",
  "increasing-arrows-process",
  "increasing-circle-process",
  "pie-process",
  // The twelve picture layouts of §5.39. Every one of their Office definitions
  // presents descendants (`presOf axis="desOrSelf"`, or `axis="des"` for
  // Titled Picture Blocks and Bending Picture Accent List), so Level 2 text is
  // DRAWN and Promote / Demote have somewhere to put it.
  ...SMARTART_CONVERT_ONLY_PICTURE_KINDS,
  // The List-tab layouts of §5.70: Microsoft names their Level 2 text in so
  // many words ("Level 2 text appears beside a small square shape", "the
  // corresponding rectangles contain Level 2 text") and every one of their
  // definitions presents descendants.
  ...SMARTART_OFFICE_LIST_KINDS,
  // The gallery layouts drawn from their Office definitions: every one
  // presents descendants, and Microsoft names Level 2 for Detailed Process
  // ("large amounts of Level 2 text"), Plus and Minus ("multiple
  // sub-levels") and Counterbalance Arrows ("works well with Level 2 text").
  ...SMARTART_OFFICE_DEFINITION_KINDS,
  // The Office.com layouts: every definition presents descendants, and Tab
  // List's description names its Level 2 text.
  ...SMARTART_OFFICE_ONLINE_KINDS,
  // Designer's layouts: each presents descendants, and Microsoft names the
  // Level 2 text of each ("Level 1 and Level 2 text appear in a rectangle",
  // "the description is immediately below").
  ...SMARTART_OFFICE_DESIGNER_KINDS,
  // A file's own definition draws whatever levels it presents.
  SMARTART_OFFICE_DEFINITION_KIND,
]);

/**
 * The organization chart layouts, the only ones with an ASSISTANT position
 * (Microsoft Support, "Create an organization chart": Add Assistant "is
 * available only for organization chart layouts"). Every other layout draws
 * an item as the ordinary child it also is, so the assistant marker is kept
 * only on these.
 */
export const SMARTART_ASSISTANT_KINDS: ReadonlySet<SmartArtKind> = new Set<SmartArtKind>([
  "hierarchy",
  "half-circle-org",
  // Microsoft: "The assistant shape and Org Chart hanging layouts are
  // available with this layout" (Name and Title Organization Chart).
  "name-and-title-org",
]);








export const SMARTART_NO_BULLET_KINDS: ReadonlySet<SmartArtKind> = new Set<SmartArtKind>([
  "hierarchy",
  "horizontal-hierarchy",
  "half-circle-org",
  "name-and-title-org",
  "circle-relationship",
  "horizontal-multi-level-hierarchy",
  "basic-hierarchy",
  "table-hierarchy",
]);

/** Whether Add Bullet is offered on `kind` ({@link SMARTART_NO_BULLET_KINDS}). */
export function smartArtLayoutTakesBullets(kind: SmartArtKind | undefined): boolean {
  return kind !== undefined && !SMARTART_NO_BULLET_KINDS.has(kind);
}

/**
 * An organization chart branch's layout (SmartArt Design > Create Graphic >
 * Layout), Office's `hierBranch` layout variable (ISO/IEC 29500-1 §21.4.7.33
 * `ST_HierBranchStyle`) on the box whose subordinates it arranges:
 * `std` Standard (side by side), `hang` Both (two per row under it), `l`
 * Left Hanging and `r` Right Hanging (stacked on that side of the line down
 * from it). Microsoft Support, "Create an organization chart"; Microsoft
 * Learn, `MsoOrgChartLayoutType`.
 */
export type SmartArtHierBranch = "std" | "hang" | "l" | "r";

export const SMARTART_HIER_BRANCHES: readonly SmartArtHierBranch[] = ["std", "hang", "l", "r"];

export function isSmartArtHierBranch(value: unknown): value is SmartArtHierBranch {
  return typeof value === "string" && (SMARTART_HIER_BRANCHES as readonly string[]).includes(value);
}

/**
 * Whether Layout (Standard / Both / Left Hanging / Right Hanging) is offered
 * on `kind`: the organization charts, whose Office definitions read the
 * variable ("hanging layouts work only with organization chart layouts").
 */
export function smartArtLayoutTakesHierBranch(kind: SmartArtKind | undefined): boolean {
  return kind !== undefined && SMARTART_ASSISTANT_KINDS.has(kind);
}

export function smartArtSupportsHierarchy(kind: SmartArtKind | undefined): boolean {
  return kind !== undefined && HIERARCHY_SMARTART_KINDS.has(kind);
}

/**
 * Whether Promote and Demote work on `kind`: on every layout, as in
 * PowerPoint, whose Text Pane takes Tab / Shift+Tab whatever the layout. A
 * layout that draws a Level 2 line shows it (bullet text in its item's shape,
 * a child box); one that does not (the matrices, Funnel, Balance) keeps it as
 * unused text (`smartArtHidden`), listed in the Text Pane and drawn again by a
 * layout that has room. Every layout that used to refuse is drawn from
 * Office's own definition, which decides. Add Shape Above / Below, conversion
 * and default nesting still follow {@link smartArtSupportsHierarchy}.
 */
export function smartArtAcceptsLevelTwo(kind: SmartArtKind | undefined): boolean {
  return kind !== undefined;
}

/**
 * Whether a SmartArt shape's own Shape Effects state a 3-D look: a Preset, a
 * Bevel, a 3-D Rotation off the front or a 3-D Format beyond its light.
 * Office's point-level `a:scene3d` / `a:sp3d` replace the SmartArt Style's on
 * that one shape (ECMA-376 §21.4.3.8), so the style's `shape3d` / `scene3d`
 * come off it and
 * the shape's own effects draw its solid -- one 3-D, never the style's and
 * the author's on top of each other.
 */
export function smartArtEffectsOwn3d(effects: unknown): boolean {
  if (!effects || typeof effects !== "object") return false;
  const record = effects as Record<string, unknown>;
  const stated = (value: unknown) => value !== undefined && value !== null && value !== "none";
  const rotation = record.rotation3d;
  // A point's own `a:scene3d` (its light) with no `a:sp3d` keeps the
  // style's solid: PowerPoint draws such a shape with the style's bevel
  // (nobevel-style-probe p8, probe2: its whole metal, deep solid). Its 3-D
  // Format then states lighting alone, which is no solid of its own.
  const format3d = record.format3d;
  const ownSolid =
    stated(format3d) &&
    (typeof format3d !== "object" || Object.keys(format3d as object).some((key) => key !== "lighting"));
  return (
    stated(record.preset) ||
    stated(record.bevel) ||
    ownSolid ||
    (!!rotation &&
      typeof rotation === "object" &&
      (rotation as { camera?: unknown }).camera !== "orthographicFront")
  );
}














export function smartArtShape3dUnderEffects<T extends { bevelTop?: unknown }>(
  shape3d: T | undefined,
  effects: unknown,
): T | undefined {
  if (!shape3d || shape3d.bevelTop === undefined) return shape3d;
  if (!smartArtEffectsNoBevel(effects)) return shape3d;
  const rest: Partial<T> = { ...shape3d };
  delete rest.bevelTop;
  return Object.keys(rest).length ? (rest as T) : undefined;
}

/** Whether a SmartArt shape's own Shape Effects state No Bevel (`bevel: "none"`). */
export function smartArtEffectsNoBevel(effects: unknown): boolean {
  return !!effects && typeof effects === "object" && (effects as { bevel?: unknown }).bevel === "none";
}

/**
 * Whether a SmartArt node wears its SmartArt Style's top bevel (the style's
 * `shape3d`, the node's own 3-D not yet stated): No Bevel then states the
 * bevel `"none"`, Office's point-level `a:sp3d` with no `a:bevelT`.
 */
export function smartArtNodeWearsStyleBevel(block: unknown): boolean {
  if (readSmartArtBlockMarker(block)?.role !== "node") return false;
  const shape3d = (block as { shape3d?: { bevelTop?: unknown } }).shape3d;
  return !!shape3d && shape3d.bevelTop !== undefined && shape3d.bevelTop !== null;
}
