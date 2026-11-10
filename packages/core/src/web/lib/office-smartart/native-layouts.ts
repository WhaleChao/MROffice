/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The ONE place Office's SmartArt layout identities are known.
 *
 * MRAI supports the layouts of `./catalog`. When an Office file
 * (pptx / docx / xlsx) carries a native SmartArt whose DiagramML layout part
 * declares one of the `uniqueId`s below, the import converts it into MRAI's own
 * SmartArt (see `./native-convert`); every other layout stays preserved and is
 * only converted when the author picks a catalog layout by hand.
 *
 * Source of the table: `docs/office-graphics/SMARTART-IMPORT.md` §1, verified
 * against the PowerPoint-made 39-diagram fixture
 * (`apps/api/tests/file_adapters/pptx/fixtures/powerpoint-smartart-39-native-golden.pptx`,
 * slide N = catalog layout N).
 *
 * Nothing here is host-shaped: the functions take and return plain strings so
 * the PPT receipt, the docx importer and the xlsx importer all share them.
 */

import { officeSmartArtLayout } from "./catalog";

/**
 * Office writes layout urns under three vocabularies, and which one a layout
 * uses is not derivable from its name: the 2007 gallery is `2005/8`, the 2010
 * additions are `2008`, and two of them (`BlockDescendingList`, `StepUpProcess`)
 * are `2009/3`. Every prefix below is the one the fixture deck actually carries.
 */
const LAYOUT_2005 = "urn:microsoft.com/office/officeart/2005/8/layout/";
const LAYOUT_2008 = "urn:microsoft.com/office/officeart/2008/layout/";
const LAYOUT_2009 = "urn:microsoft.com/office/officeart/2009/3/layout/";





const LAYOUT_2009_BARE = "urn:microsoft.com/office/officeart/2009/layout/";
/** The Office.com additions of 2011 (Hexagon Radial, Tab List). */
const LAYOUT_2011 = "urn:microsoft.com/office/officeart/2011/layout/";
/** The icon layouts of 2018, in two vocabularies (`2018/2`, `2018/5`). */
const LAYOUT_2018_2 = "urn:microsoft.com/office/officeart/2018/2/layout/";
const LAYOUT_2018_5 = "urn:microsoft.com/office/officeart/2018/5/layout/";
/** Designer's layouts of 2016 and 2017, two vocabularies more (`2016/7`, `2017/3`). */
const LAYOUT_2016_7 = "urn:microsoft.com/office/officeart/2016/7/layout/";
const LAYOUT_2017_3 = "urn:microsoft.com/office/officeart/2017/3/layout/";

/**
 * catalog id → the Office layout `uniqueId` it is imported from.
 *
 * `default` (Office's "Basic Block List") is the source of BOTH `basic-list`
 * and `card-grid`; it maps to `basic-list`, and `card-grid` is reachable only
 * through the insert gallery or the manual conversion (SMARTART-IMPORT.md §1).
 */
export const NATIVE_SMARTART_LAYOUT_IDENTITIES: Readonly<
  Record<string, string>
> = {
  "basic-list": `${LAYOUT_2005}default`,
  "vertical-box-list": `${LAYOUT_2005}list1`,
  "grouped-list": `${LAYOUT_2005}lProcess2`,
  "card-grid": `${LAYOUT_2005}default`,
  "step-list": `${LAYOUT_2009}BlockDescendingList`,
  "basic-process": `${LAYOUT_2005}process1`,
  "chevron-process": `${LAYOUT_2005}chevron1`,
  "vertical-process": `${LAYOUT_2005}process2`,
  "alternating-process": `${LAYOUT_2005}vProcess5`,
  timeline: `${LAYOUT_2005}hProcess11`,
  "step-up-process": `${LAYOUT_2009}StepUpProcess`,
  "basic-cycle": `${LAYOUT_2005}cycle2`,
  "continuous-cycle": `${LAYOUT_2005}cycle3`,
  "radial-cycle": `${LAYOUT_2005}radial6`,
  "segmented-cycle": `${LAYOUT_2005}cycle8`,
  "org-chart": `${LAYOUT_2005}orgChart1`,
  "horizontal-hierarchy": `${LAYOUT_2005}hierarchy2`,
  "labeled-tree": `${LAYOUT_2005}hierarchy6`,
  "half-circle-org": `${LAYOUT_2008}HalfCircleOrganizationChart`,
  "opposing-arrows": `${LAYOUT_2005}arrow4`,
  balance: `${LAYOUT_2005}balance1`,
  funnel: `${LAYOUT_2005}funnel1`,
  venn: `${LAYOUT_2005}venn1`,
  "linear-venn": `${LAYOUT_2005}venn3`,
  target: `${LAYOUT_2005}target1`,
  radial: `${LAYOUT_2005}radial1`,
  gears: `${LAYOUT_2005}gear1`,
  "basic-matrix": `${LAYOUT_2005}matrix3`,
  "titled-matrix": `${LAYOUT_2005}matrix1`,
  "grid-matrix": `${LAYOUT_2005}matrix2`,
  "cycle-matrix": `${LAYOUT_2005}cycle4`,
  "basic-pyramid": `${LAYOUT_2005}pyramid1`,
  "inverted-pyramid": `${LAYOUT_2005}pyramid3`,
  "segmented-pyramid": `${LAYOUT_2005}pyramid4`,
  "pyramid-list": `${LAYOUT_2005}pyramid2`,
  "picture-caption-list": `${LAYOUT_2005}pList1`,
  "alternating-picture-blocks": `${LAYOUT_2008}AlternatingPictureBlocks`,
  "circular-picture-callout": `${LAYOUT_2008}CircularPictureCallout`,
  "picture-accent-process": `${LAYOUT_2005}hProcess10`,
  // The seven FIXED-CAPACITY processes (SMARTART-IMPORT.md §5.34). They are
  // catalog identities like any other — an import reads them, an export claims
  // them — but MRAI stocks no preset for them, so the only way one reaches a
  // file is the patchback keeping Office's own layout part. `catalog.ts` marks
  // them `convertOnly`, which is what keeps them out of the insert gallery.
  // Every vocabulary below is READ off the corpus's own layout parts, never
  // inferred from the name: six of the seven are `2009/3`, which no naming rule
  // predicts (§5.27 made the same correction for `HorizontalOrganizationChart`).
  "upward-arrow": `${LAYOUT_2005}arrow2`,
  "descending-process": `${LAYOUT_2009}DescendingProcess`,
  "phased-process": `${LAYOUT_2009}PhasedProcess`,
  "sub-step-process": `${LAYOUT_2009}SubStepProcess`,
  "increasing-arrows-process": `${LAYOUT_2009}IncreasingArrowsProcess`,
  "increasing-circle-process": `${LAYOUT_2008}IncreasingCircleProcess`,
  "pie-process": `${LAYOUT_2009}PieProcess`,
  // The twelve PICTURE layouts of §5.39, also convert-only. Every vocabulary
  // and every spelling below was read off the corpus's own layout parts: eight
  // are `2008`, `FramedTextPicture` is `2009/3`, and the three list layouts
  // keep Office's own lower-camel ids (`bList2`, `vList3`, `vList4`).
  "picture-lineup": `${LAYOUT_2008}PictureLineup`,
  "picture-grid": `${LAYOUT_2008}PictureGrid`,
  "bending-picture-blocks": `${LAYOUT_2008}BendingPictureBlocks`,
  "bending-picture-caption": `${LAYOUT_2008}BendingPictureCaption`,
  "bending-picture-semi-transparent-text": `${LAYOUT_2008}BendingPictureSemiTransparentText`,
  "picture-strips": `${LAYOUT_2008}PictureStrips`,
  "bending-picture-caption-list": `${LAYOUT_2008}BendingPictureCaptionList`,
  "titled-picture-blocks": `${LAYOUT_2008}TitledPictureBlocks`,
  "framed-text-picture": `${LAYOUT_2009}FramedTextPicture`,
  "vertical-picture-list": `${LAYOUT_2005}vList4`,
  "vertical-picture-accent-list": `${LAYOUT_2005}vList3`,
  "bending-picture-accent-list": `${LAYOUT_2005}bList2`,
  // The last nine picture layouts on the ordinary node binding (§5.46). Again
  // every vocabulary is READ off the corpus's own layout parts: five are
  // `2008`, `SnapshotPictureList` is `2009/3`, and the three classic list
  // layouts keep Office's lower-camel ids. Two of Microsoft's names are close
  // enough to confuse — `hList2` is "Picture Accent List" (its blip is on the
  // Level 1 node) and `PictureAccentList` is "Titled Picture Accent List"
  // ("an accent picture for each Level 2 text", and its blip really is two
  // `ch` frames deep) — so the ids below follow Microsoft's names, not the
  // urns' spelling.
  "captioned-pictures": `${LAYOUT_2008}CaptionedPictures`,
  "snapshot-picture-list": `${LAYOUT_2009}SnapshotPictureList`,
  "title-picture-lineup": `${LAYOUT_2008}TitlePictureLineup`,
  "picture-accent-blocks": `${LAYOUT_2008}PictureAccentBlocks`,
  "picture-accent-list": `${LAYOUT_2005}hList2`,
  "alternating-picture-circles": `${LAYOUT_2008}AlternatingPictureCircles`,
  "continuous-picture-list": `${LAYOUT_2005}hList7`,
  "horizontal-picture-list": `${LAYOUT_2005}pList2`,
  "titled-picture-accent-list": `${LAYOUT_2008}PictureAccentList`,
  // The last seven picture layouts in the corpus (§5.51). `SpiralPicture` is
  // `2009/3`, `CirclePictureHierarchy` is the bare `2009` — a vocabulary this
  // table had never seen — `radial2` keeps Office's lower-camel id, and the
  // rest are `2008`. All seven were read off the corpus's own layout parts.
  "bubble-picture-list": `${LAYOUT_2008}BubblePictureList`,
  "hexagon-cluster": `${LAYOUT_2008}HexagonCluster`,
  "spiral-picture": `${LAYOUT_2009}SpiralPicture`,
  "ascending-picture-accent-process": `${LAYOUT_2008}AscendingPictureAccentProcess`,
  "accented-picture": `${LAYOUT_2008}AccentedPicture`,
  "circle-picture-hierarchy": `${LAYOUT_2009_BARE}CirclePictureHierarchy`,
  "radial-list": `${LAYOUT_2005}radial2`,
  
  
  // `2018/5`, that no naming rule predicts either.
  "icon-vertical-solid-list": `${LAYOUT_2018_2}IconVerticalSolidList`,
  "icon-circle-list": `${LAYOUT_2018_2}IconCircleList`,
  "icon-circle-label-list": `${LAYOUT_2018_5}IconCircleLabelList`,
  
  
  
  "icon-label-description-list": `${LAYOUT_2018_2}IconLabelDescriptionList`,
  "centered-icon-label-description-list": `${LAYOUT_2018_5}CenteredIconLabelDescriptionList`,
  "icon-leaf-label-list": `${LAYOUT_2018_5}IconLeafLabelList`,
  "basic-linear-process-numbered": `${LAYOUT_2016_7}BasicLinearProcessNumbered`,
  "vertical-solid-action-list": `${LAYOUT_2016_7}VerticalSolidActionList`,
  "drop-pin-timeline": `${LAYOUT_2017_3}DropPinTimeline`,
  
  // layout part. Two keep a `+Icon` suffix Office's ids carry although
  // neither holds a picture; Tabbed Arc is not Microsoft's vocabulary at all
  // (`diagrams.loki3.com`), and Hexagon Radial and Tab List are `2011`.
  "interconnected-rings": `${LAYOUT_2005}rings+Icon`,
  "hexagon-radial": `${LAYOUT_2011}HexagonRadial`,
  "tabbed-arc": "urn:diagrams.loki3.com/TabbedArc+Icon",
  "tab-list": `${LAYOUT_2011}TabList`,
  "chevron-accent-process": `${LAYOUT_2005}chevronAccent+Icon`,
  // PowerPoint's List tab beyond the first five (SMARTART-IMPORT.md §5.70).
  // STOCKED, unlike every block since the 39: Office's own definition of each
  // is in the preset store, so these are inserted and exported under their own
  // urn. Every vocabulary was READ off the layout parts Office wrote in the
  
  // ids, the 2010 additions (`LinedList`, `VerticalCurvedList`,
  // `AlternatingHexagons`, `SquareAccentList`, `VerticalAccentList`,
  // `VerticalCircleList`) are `2008`.
  "lined-list": `${LAYOUT_2008}LinedList`,
  "vertical-bullet-list": `${LAYOUT_2005}vList2`,
  "horizontal-bullet-list": `${LAYOUT_2005}hList1`,
  "stacked-list": `${LAYOUT_2005}hList9`,
  "trapezoid-list": `${LAYOUT_2005}hList6`,
  "vertical-curved-list": `${LAYOUT_2008}VerticalCurvedList`,
  "alternating-hexagons": `${LAYOUT_2008}AlternatingHexagons`,
  "square-accent-list": `${LAYOUT_2008}SquareAccentList`,
  "vertical-block-list": `${LAYOUT_2005}vList5`,
  "vertical-accent-list": `${LAYOUT_2008}VerticalAccentList`,
  "table-list": `${LAYOUT_2005}hList3`,
  "vertical-circle-list": `${LAYOUT_2008}VerticalCircleList`,
  // The six gallery layouts drawn only from their Office definitions
  
  // Plus and Minus's carries a trailing space -- Office's own spelling, in
  // its definition and in every file it writes -- so it is kept exactly.
  "arrow-ribbon": `${LAYOUT_2005}arrow6`,
  "block-cycle": `${LAYOUT_2005}cycle5`,
  "horizontal-labeled-hierarchy": `${LAYOUT_2005}hierarchy5`,
  "detailed-process": `${LAYOUT_2005}hProcess7`,
  "plus-and-minus": `${LAYOUT_2009}PlusandMinus `,
  "reverse-list": `${LAYOUT_2009_BARE}ReverseList`,
  "circular-bending-process": `${LAYOUT_2005}bProcess2`,
  "counterbalance-arrows": `${LAYOUT_2005}arrow3`,
  "basic-pie": `${LAYOUT_2005}chart3`,
  "circle-arrow-process": `${LAYOUT_2009_BARE}CircleArrowProcess`,
  "circle-relationship": `${LAYOUT_2009}CircleRelationship`,
  "horizontal-multi-level-hierarchy": `${LAYOUT_2008}HorizontalMultiLevelHierarchy`,
  "opposing-ideas": `${LAYOUT_2009}OpposingIdeas`,
  "nested-target": `${LAYOUT_2005}target2`,
  "vertical-arrow-list": `${LAYOUT_2005}vList6`,
  "basic-hierarchy": `${LAYOUT_2005}hierarchy1`,
  "process-arrows": `${LAYOUT_2005}hProcess6`,
  "radial-venn": `${LAYOUT_2005}radial3`,
  "closed-chevron-process": `${LAYOUT_2005}hChevron3`,
  "hierarchy-list": `${LAYOUT_2005}hierarchy3`,
  "basic-equation": `${LAYOUT_2005}equation1`,
  "diverging-arrows": `${LAYOUT_2005}arrow1`,
  "converging-arrows": `${LAYOUT_2005}arrow5`,
  "circle-accent-timeline": `${LAYOUT_2008}CircleAccentTimeline`,
  "multidirectional-cycle": `${LAYOUT_2005}cycle7`,
  "vertical-equation": `${LAYOUT_2005}equation2`,
  "table-hierarchy": `${LAYOUT_2005}hierarchy4`,
  "alternating-flow": `${LAYOUT_2005}hProcess4`,
  "process-list": `${LAYOUT_2005}lProcess1`,
  "chevron-list": `${LAYOUT_2005}lProcess3`,
  "converging-radial": `${LAYOUT_2005}radial4`,
  "diverging-radial": `${LAYOUT_2005}radial5`,
  "radial-cluster": `${LAYOUT_2008}RadialCluster`,
  "random-to-result-process": `${LAYOUT_2009}RandomtoResultProcess`,
  "stacked-venn": `${LAYOUT_2005}venn2`,
  "step-down-process": `${LAYOUT_2005}StepDownProcess`,
  "target-list": `${LAYOUT_2005}target3`,
  "name-and-title-org": `${LAYOUT_2008}NameandTitleOrganizationalChart`,
};

/**
 * Office layouts MRAI stocks no definition for, but whose MEANING is one of the
 * catalog engines: Office's `uniqueId` → the catalog layout it imports as.
 * Python mirror: `ALIAS_LAYOUT_IDENTITIES` in
 * `apps/api/app/services/office_smartart_diagram_write.py`.
 *
 * Every entry got Microsoft's own per-layout description read BEFORE it was
 * mapped (`docs/office-graphics/SMARTART-IMPORT.md` §5.23 and §5.27): the
 * purpose sentence has to be the catalog layout's, and the text LEVELS Office
 * displays have to match, or converting would change what the diagram means or
 * hide text Office shows. A difference that is only about how the shapes use
 * the space — bending, horizontal, which element is emphasised — is a drawing
 * difference, which §5.2 already accepts.
 *
 * An alias is ONE-DIRECTIONAL: it is what an import READS, never what an export
 * CLAIMS. {@link catalogIdToNativeLayout} still answers with the canonical urn,
 * so the insert galleries and every diagram MRAI authors from nothing stay on
 * the 39 stocked presets. An alias urn reaches a file only one way — the
 * patchback keeping the layout part Office itself wrote, which is what lets a
 * mapped diagram round-trip as its ORIGINAL layout.
 */
export const NATIVE_SMARTART_LAYOUT_ALIASES: Readonly<
  Record<string, string>
> = {
  [`${LAYOUT_2005}process5`]: "basic-process",
  [`${LAYOUT_2005}bProcess3`]: "basic-process",
  [`${LAYOUT_2005}bProcess4`]: "basic-process",
  [`${LAYOUT_2005}cycle1`]: "basic-cycle",
  
  // are catalog layouts of their own now (Vertical / Horizontal Bullet List,
  // §5.70), so an import reads them as themselves. A page converted while they
  // were aliases still says `basic-list` over a package holding Office's
  // `vList2` part; the export side keeps that part (`RETIRED_ALIAS_LAYOUT_
  // IDENTITIES` in `office_smartart_diagram_write.py`), so such a diagram
  // still reaches Office as the layout its author had.
  [`${LAYOUT_2009}HorizontalOrganizationChart`]: "org-chart",
  // §5.31. Microsoft's purpose sentence for each of these five contains Basic
  // Process's own — "sequential steps in a task, process, or workflow" — and
  // each states that BOTH Level 1 and Level 2 text is displayed, which is what
  // Basic Process displays. What is left over in every one of them is where
  // the text is drawn (inside/outside the arrow, in a separate shape, which
  // level is emphasised), and §5.2 already accepts a drawing difference.
  [`${LAYOUT_2005}hProcess3`]: "basic-process",
  [`${LAYOUT_2005}hProcess9`]: "basic-process",
  [`${LAYOUT_2005}process3`]: "basic-process",
  [`${LAYOUT_2005}process4`]: "basic-process",
  [`${LAYOUT_2005}chevron2`]: "basic-process",
  // Nondirectional Cycle's purpose sentence is Basic Cycle's word for word.
  // "Works well when direction does not need to be indicated" is about the
  // ARROWS, not about the order — which is why `cycle7` ("can occur in any
  // direction", a cycle with no sequence at all) stays rejected and this one
  // does not. Its levels match too: like Basic Cycle and like the already
  // mapped `cycle1`, every corpus diagram of it carries Level 1 points only.
  [`${LAYOUT_2005}cycle6`]: "basic-cycle",
  // Office ships a second definition named "Bending Picture Semi-Transparent
  // Text" under a hyphenated urn. PowerPoint's gallery does not list it (it is
  // not among the 134 layouts of the owner's every-layout Word file, and its
  // category is "Picture", which no gallery tab is), but a file can carry it;
  // its purpose and levels are the catalog layout's own.
  [`${LAYOUT_2008}BendingPictureSemi-TransparentText`]: "bending-picture-semi-transparent-text",
};



















export const ALIGNED_NATIVE_SMARTART_ALIASES: ReadonlySet<string> = new Set([
  
  `${LAYOUT_2005}hProcess9`,
  
  `${LAYOUT_2005}bProcess4`,
  `${LAYOUT_2005}cycle1`,
  `${LAYOUT_2009}HorizontalOrganizationChart`,
  
  
  `${LAYOUT_2005}process3`,
  `${LAYOUT_2005}process4`,
  `${LAYOUT_2005}process5`,
  `${LAYOUT_2005}hProcess3`,
  `${LAYOUT_2005}bProcess3`,
  `${LAYOUT_2005}chevron2`,
  `${LAYOUT_2005}cycle6`,
  
  `${LAYOUT_2008}BendingPictureSemi-TransparentText`,
]);

/**
 * Aligned aliases PowerPoint's own galleries never list (above): drawn and
 * kept as themselves, but given no Change Layout entry of their own, so the
 * gallery does not show a second card under the catalog layout's name.
 */
const UNLISTED_NATIVE_SMARTART_ALIASES: ReadonlySet<string> = new Set([
  `${LAYOUT_2008}BendingPictureSemi-TransparentText`,
]);

/**
 * The reverse table. `default` resolves to `basic-list` (the ambiguity the
 * decision resolves), so this is NOT a mechanical inversion of the map above.
 */
const CATALOG_BY_IDENTITY: ReadonlyMap<string, string> = new Map(
  Object.entries(NATIVE_SMARTART_LAYOUT_IDENTITIES)
    .filter(([layoutId]) => layoutId !== "card-grid")
    .map(([layoutId, uniqueId]) => [uniqueId, layoutId] as const),
);

/**
 * The catalog layout id an Office layout `uniqueId` imports as, or `null` when
 * MRAI does not own that layout (the diagram stays preserved).
 *
 * Accepts the full urn only: a bare suffix such as `process1` is not an Office
 * identity and must never be guessed at.
 */
export function nativeLayoutToCatalogId(uniqueId: unknown): string | null {
  if (typeof uniqueId !== "string") return null;
  return (
    CATALOG_BY_IDENTITY.get(uniqueId) ??
    NATIVE_SMARTART_LAYOUT_ALIASES[uniqueId] ??
    null
  );
}

/**
 * Is this urn an ALIAS — a layout MRAI converts but does not stock?
 *
 * The answer the export side needs: such a diagram keeps the layout part its
 * own package holds, and MRAI writes the other four.
 */
export function isAliasNativeSmartArtLayout(uniqueId: unknown): boolean {
  return (
    typeof uniqueId === "string" &&
    Object.hasOwn(NATIVE_SMARTART_LAYOUT_ALIASES, uniqueId)
  );
}














export function nativeLayoutConvertsOnOpen(uniqueId: unknown): boolean {
  return (
    nativeLayoutToCatalogId(uniqueId) !== null &&
    (!isAliasNativeSmartArtLayout(uniqueId) ||
      ALIGNED_NATIVE_SMARTART_ALIASES.has(uniqueId as string))
  );
}

/**
 * The alias urn a diagram of catalog layout `kind` is laid out from, or
 * `null`: `urn` when it is an ALIGNED alias
 * ({@link ALIGNED_NATIVE_SMARTART_ALIASES}) of exactly that catalog layout.
 * `kind` may be any id the catalog resolves (engine kind, canonical id).
 *
 * The one reading of a member's `smartArtLayoutUrn` and of a source
 * diagram's own urn: a record that disagrees with the kind (hProcess9 on a
 * cycle) names nothing, so the diagram is drawn as its catalog layout.
 */
export function smartArtAliasLayoutUrn(kind: unknown, urn: unknown): string | null {
  if (typeof urn !== "string" || !ALIGNED_NATIVE_SMARTART_ALIASES.has(urn)) {
    return null;
  }
  const target = officeSmartArtLayout(NATIVE_SMARTART_LAYOUT_ALIASES[urn]);
  const entry = officeSmartArtLayout(kind);
  return target && entry && target.id === entry.id ? urn : null;
}

/**
 * `urn` when it is an alias (aligned or not) of catalog layout `kind`, else
 * `null`: what a stored `smartArtLayoutUrn` is held to on the way in. Looser
 * than {@link smartArtAliasLayoutUrn} on purpose — a page stays valid as it
 * was written even if an alias ever leaves the aligned set; only the DRAWING
 * then falls back to the catalog layout.
 */
export function nativeSmartArtAliasOf(kind: unknown, urn: unknown): string | null {
  if (typeof urn !== "string" || !isAliasNativeSmartArtLayout(urn)) return null;
  const target = officeSmartArtLayout(NATIVE_SMARTART_LAYOUT_ALIASES[urn]);
  const entry = officeSmartArtLayout(kind);
  return target && entry && target.id === entry.id ? urn : null;
}

/**
 * The urn to convert a preserved diagram whose own urn is `sourceUrn` as, when
 * catalog layout `layoutId` is asked for (the open-time conversion, or the
 * author's 转换为可编辑 pick): the diagram's OWN urn when it is an aligned alias
 * of that layout — the author asked for the layout the diagram already is —
 * else the catalog layout's.
 */
export function nativeLayoutIdentityForCatalogPick(
  sourceUrn: unknown,
  layoutId: unknown,
): string | null {
  return smartArtAliasLayoutUrn(layoutId, sourceUrn) ?? catalogIdToNativeLayout(layoutId);
}

/**
 * The Office layout `uniqueId` a catalog layout came from (or would be written
 * as). Accepts any id the catalog resolves (canonical id, engine kind, alias).
 */
export function catalogIdToNativeLayout(idOrAlias: unknown): string | null {
  const entry = officeSmartArtLayout(idOrAlias);
  return entry ? (NATIVE_SMARTART_LAYOUT_IDENTITIES[entry.id] ?? null) : null;
}

/**
 * Every Office layout identity the import owns: the stocked and convert-only
 * ones in catalog order, then the aliases. Both are layouts a file opens as an editable MRAI
 * diagram, which is what every caller of this list asks about.
 */
export const IMPORTED_NATIVE_SMARTART_LAYOUTS: readonly string[] = [
  ...CATALOG_BY_IDENTITY.keys(),
  ...Object.keys(NATIVE_SMARTART_LAYOUT_ALIASES),
];

/**
 * The Change Layout gallery's entry for an aligned alias, beside the catalog
 * layout it stands for: PowerPoint lists Continuous Block Process as a layout
 * of its own, so picking it gives that layout and picking Basic Process gives
 * Basic Process. `galleryId` is `<engine kind>@<urn suffix>`.
 */
export interface SmartArtAliasGalleryEntry {
  readonly urn: string;
  /** The urn's last segment (`hProcess9`): the message key's suffix. */
  readonly suffix: string;
  readonly catalogId: string;
  readonly kind: string;
  readonly category: string;
  readonly galleryId: string;
}

export const SMARTART_ALIAS_GALLERY: readonly SmartArtAliasGalleryEntry[] = [
  ...ALIGNED_NATIVE_SMARTART_ALIASES,
].filter((urn) => !UNLISTED_NATIVE_SMARTART_ALIASES.has(urn)).flatMap((urn) => {
  const entry = officeSmartArtLayout(NATIVE_SMARTART_LAYOUT_ALIASES[urn]);
  const suffix = urn.slice(urn.lastIndexOf("/") + 1);
  return entry
    ? [{ urn, suffix, catalogId: entry.id, kind: entry.kind, category: entry.category, galleryId: `${entry.kind}@${suffix}` }]
    : [];
});

/** The gallery id a diagram of `kind` drawn as `layoutUrn` is listed under. */
export function smartArtGalleryId(kind: string, layoutUrn?: string | null): string {
  const alias = smartArtAliasLayoutUrn(kind, layoutUrn);
  return SMARTART_ALIAS_GALLERY.find((entry) => entry.urn === alias)?.galleryId ?? kind;
}

/** A gallery id back to the engine kind and the alias it names (or `null`). */
export function parseSmartArtGalleryId(id: string): { kind: string; layoutUrn: string | null } {
  const alias = SMARTART_ALIAS_GALLERY.find((entry) => entry.galleryId === id);
  return alias ? { kind: alias.kind, layoutUrn: alias.urn } : { kind: id, layoutUrn: null };
}
