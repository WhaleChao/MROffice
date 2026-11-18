/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import galleryFacts from "./gallery-facts.json";
import {
  SMARTART_PICTURE_KINDS,
  smartArtSupportsHierarchy,
  type OfficeSmartArtKind,
} from "./model";

/** Office's gallery tabs, minus the ones we do not ship (All / Office.com / Other). */
export const OFFICE_SMARTART_CATEGORY_IDS = [
  "list",
  "process",
  "cycle",
  "hierarchy",
  "relationship",
  "matrix",
  "pyramid",
  "picture",
] as const;

export type OfficeSmartArtCategoryId =
  (typeof OFFICE_SMARTART_CATEGORY_IDS)[number];

export interface OfficeSmartArtLayout {
  /** Canonical id (the Docs id); what a host should persist for new content. */
  id: string;
  /** Primary category. Office's membership is many-to-many; ours is one tab. */
  category: OfficeSmartArtCategoryId;
  /** Engine id accepted by `buildSmartArt`, and the legacy presentation id. */
  kind: OfficeSmartArtKind;
   
  aliases: readonly string[];
  /** i18n keys shared by every host's gallery. */
  nameKey: string;
  descriptionKey: string;
  categoryKey: string;
  /** The layout draws picture placeholders (custody rules apply on relayout). */
  picture: boolean;
  /** The layout reads outline levels (Promote / Demote change the geometry). */
  hierarchy: boolean;
  /** Number of seed items an insert/preview uses. */
  seedItems: number;
  /** The first seed item is the tree root and takes the host's root label. */
  seedRootLabel: boolean;
  /**
   * Seed outline levels, when they differ from the engine's default nesting
   * (level 0 for item 1, level 1 for the rest of a hierarchy layout, level 0
   * everywhere else). A host that ignores them still gets a valid diagram.
   */
  seedLevels?: readonly number[];
  /**
   * The seed items that are ASSISTANTS (`dgm:pt type="asst"`), by index:
   * Office's organization charts insert a manager, an assistant and three
   * subordinates (their definitions' `dgm:sampData`). The Text Pane lists an
   * assistant after its manager's other children, so it is the last of them.
   */
  seedAssistants?: readonly number[];
  /** Documented Office shape capacity; absent = documented as unlimited. */
  officeMaxItems?: number;
  /** Documented level/role note for that capacity, verbatim in spirit. */
  officeLevels?: string;
  /**
   * Deepest outline level Office DRAWS for this layout (0 = Level 1 only);
   * absent = Microsoft states no restriction. Read this round only for the
   * capped processes (SMARTART-IMPORT.md §5.34).
   */
  officeMaxLevel?: number;
  /**
   * Per Level 1 item, how many Level 2 items Office draws under it; a position
   * past the end of the array is unlimited. One layout states this — Phased
   * Process, "the first two Level 1 items can each contain four Level 2 items,
   * and the third… an unlimited number".
   */
  officeMaxChildren?: readonly number[];
  /**
   * Not offered from nothing: it is reached only by CONVERTING a diagram whose
   * package already carries Office's own layout part, as PowerPoint's own
   * dialog does not offer it either (the icon and Office.com layouts). It is
   * absent from the insert gallery and from Change Layout.
   */
  convertOnly?: boolean;
  

















  officePictureLevel?: number;
  



















  officePictureItems?: number;
}

const key = (id: string) => ({
  nameKey: `toolbar.smartArtLayoutName_${id}`,
  descriptionKey: `toolbar.smartArtLayoutDesc_${id}`,
});

/** What an entry states; everything else is derived from the engine. */
interface OfficeSmartArtLayoutInput {
  id: string;
  category: OfficeSmartArtCategoryId;
  /** Only when the engine id differs from the canonical id. */
  kind?: OfficeSmartArtKind;
  aliases?: readonly string[];
  seedItems: number;
  seedRootLabel?: boolean;
  seedLevels?: readonly number[];
  seedAssistants?: readonly number[];
  officeMaxItems?: number;
  officeLevels?: string;
  officeMaxLevel?: number;
  officeMaxChildren?: readonly number[];
  convertOnly?: boolean;
  officePictureLevel?: number;
  officePictureItems?: number;
}

const layout = (entry: OfficeSmartArtLayoutInput): OfficeSmartArtLayout => {
  const kind = (entry.kind ?? entry.id) as OfficeSmartArtKind;
  return {
    ...entry,
    ...key(entry.id),
    kind,
    aliases: entry.aliases ?? [],
    categoryKey: `toolbar.smartArtCategory_${entry.category}`,
    picture: SMARTART_PICTURE_KINDS.some((candidate) => candidate === kind),
    hierarchy: smartArtSupportsHierarchy(kind),
    seedRootLabel: entry.seedRootLabel ?? false,
  };
};

/** Seed levels for a root with three children (also the engine's default). */
const ROOT_WITH_THREE = [0, 1, 1, 1] as const;
/**
 * An organization chart's sample (orgChart1, Half Circle and Name and Title's
 * `dgm:sampData`): the manager, three subordinates and an assistant, listed
 * last as the Text Pane lists it.
 */
const MANAGER_WITH_ASSISTANT = { seedItems: 5, seedLevels: [0, 1, 1, 1, 1], seedAssistants: [4] } as const;
/** Office's sample of three headings with two second-level lines each. */
const THREE_WITH_TWO_EACH = [0, 1, 1, 0, 1, 1, 0, 1, 1] as const;
/** A hub and the four points around it (Office's radial layouts). */
const HUB_WITH_FOUR = [0, 1, 1, 1, 1] as const;

/**
 * Four SIBLINGS. The picture layouts of §5.39 are hierarchy-capable — Office
 * draws their Level 2 text — but their seed is flat, because a picture layout
 * seeds one picture per item, not a root with children.
 */
const FLAT_FOUR = [0, 0, 0, 0] as const;

// Seeds follow Office's own default data for the layout (the definition's
// `dgm:sampData`, the shapes PowerPoint, Excel and Word insert); a definition
// that names none (`useDef="1"`) inserts Office's default of three Level 1
// shapes, as Basic Process, Basic Venn and Gear do.
const LAYOUTS: readonly OfficeSmartArtLayout[] = [
  // List
  layout({ id: "basic-list", category: "list", kind: "list", seedItems: 5 }),
  layout({ id: "vertical-box-list", category: "list", seedItems: 3 }),
  layout({
    id: "grouped-list",
    category: "list",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
  layout({ id: "card-grid", category: "list", seedItems: 6 }),
  // Office's Descending Block List inserts three blocks with two lines each.
  layout({ id: "step-list", category: "list", seedItems: 9, seedLevels: THREE_WITH_TWO_EACH }),
  // Process
  layout({ id: "basic-process", category: "process", kind: "process", seedItems: 3 }),
  layout({ id: "chevron-process", category: "process", seedItems: 3 }),
  layout({ id: "vertical-process", category: "process", seedItems: 3 }),
  layout({ id: "alternating-process", category: "process", seedItems: 3 }),
  layout({ id: "timeline", category: "process", seedItems: 3 }),
  layout({ id: "step-up-process", category: "process", seedItems: 3 }),
  // Cycle
  layout({ id: "basic-cycle", category: "cycle", kind: "cycle", seedItems: 5 }),
  layout({ id: "continuous-cycle", category: "cycle", seedItems: 5 }),
  layout({
    id: "radial-cycle",
    category: "cycle",
    seedItems: 5,
    // Office's radial6 draws the level-1 point as the hub and ITS children
    
    // export-look comparison: a flat seed reached PowerPoint as a lone hub).
    seedRootLabel: true,
    seedLevels: HUB_WITH_FOUR,
  }),
  layout({
    id: "segmented-cycle",
    category: "cycle",
    seedItems: 3,
    officeMaxItems: 7,
  }),
  // Hierarchy
  layout({
    id: "org-chart",
    category: "hierarchy",
    kind: "hierarchy",
    ...MANAGER_WITH_ASSISTANT,
    seedRootLabel: true,
  }),
  layout({
    id: "horizontal-hierarchy",
    category: "hierarchy",
    seedItems: 6,
    seedRootLabel: true,
    seedLevels: [0, 1, 2, 2, 1, 2],
  }),
  layout({
    id: "labeled-tree",
    category: "hierarchy",
    seedItems: 9,
    seedRootLabel: true,
    seedLevels: [0, 1, 2, 2, 1, 2, 0, 0, 0],
  }),
  layout({
    id: "half-circle-org",
    category: "hierarchy",
    ...MANAGER_WITH_ASSISTANT,
    seedRootLabel: true,
  }),
  // Relationship
  layout({
    id: "opposing-arrows",
    category: "relationship",
    seedItems: 2,
    officeMaxItems: 2,
  }),
  layout({
    id: "balance",
    category: "relationship",
    // Office's own sample (balance1's `sampData`): two headings, the weights
    // under them as their children. Seeded flat, Office's definition draws
    // the first two items as the headings and nothing at all for the rest.
    seedItems: 7,
    seedLevels: [0, 1, 1, 0, 1, 1, 1],
  }),
  layout({
    id: "funnel",
    category: "relationship",
    seedItems: 4,
    officeMaxItems: 4,
  }),
  layout({ id: "venn", category: "relationship", kind: "relationship-venn", seedItems: 3 }),
  layout({ id: "linear-venn", category: "relationship", seedItems: 4 }),
  layout({
    id: "target",
    category: "relationship",
    seedItems: 3,
    officeMaxItems: 5,
  }),
  layout({
    id: "radial",
    category: "relationship",
    seedItems: 5,
    // radial1: the hub and its children, as radial-cycle above.
    seedRootLabel: true,
    seedLevels: HUB_WITH_FOUR,
  }),
  layout({
    id: "gears",
    category: "relationship",
    seedItems: 3,
    officeMaxItems: 3,
  }),
  // Matrix
  layout({
    id: "basic-matrix",
    category: "matrix",
    aliases: ["matrix"],
    seedItems: 4,
    officeMaxItems: 4,
  }),
  layout({
    id: "titled-matrix",
    category: "matrix",
    seedItems: 5,
    officeMaxItems: 5,
    // matrix1 puts the first Level 1 line in the centre and the first four
    // Level 2 lines in the quadrants: a flat seed left the quadrants empty.
    seedRootLabel: true,
    seedLevels: HUB_WITH_FOUR,
  }),
  layout({
    id: "grid-matrix",
    category: "matrix",
    seedItems: 4,
    officeMaxItems: 4,
  }),
  layout({
    id: "cycle-matrix",
    category: "matrix",
    // Office's sample: four ideas, each with one second-level description.
    seedItems: 8,
    seedLevels: [0, 1, 0, 1, 0, 1, 0, 1],
    officeMaxItems: 4,
  }),
  // Pyramid
  layout({ id: "basic-pyramid", category: "pyramid", kind: "pyramid", seedItems: 3 }),
  layout({ id: "inverted-pyramid", category: "pyramid", seedItems: 3 }),
  layout({ id: "segmented-pyramid", category: "pyramid", seedItems: 4 }),
  layout({ id: "pyramid-list", category: "pyramid", seedItems: 3 }),
  // Picture. The three whose blip is bound `<dgm:forEach axis="ch">` one frame
  // deep declare it, like every picture layout added since (§5.46): without it
  // a Level 2 item of theirs got a whole block with a picture slot MRAI drew
  // and Office draws nothing into -- measured on `hProcess10`, whose corpus
  // diagram carries six Level 2 points. `circular-picture-callout` is the one
  // exception: its blip hangs off a TRANSITION, which `picture-bindings.json`
  // states instead, so it declares no level here.
  layout({ id: "picture-caption-list", category: "picture", seedItems: 4, officePictureLevel: 0 }),
  layout({ id: "alternating-picture-blocks", category: "picture", seedItems: 3, officePictureLevel: 0 }),
  // Hub plus three accents: every picture member seeds a real placeholder
  // payload, so the picture layouts stay at the Office-documented 3-4 range.
  layout({ id: "circular-picture-callout", category: "picture", seedItems: 4 }),
  // Office's Picture Accent Process inserts three pictures with two lines each.
  layout({ id: "picture-accent-process", category: "picture", seedItems: 9, seedLevels: THREE_WITH_TWO_EACH, officePictureLevel: 0 }),
];

/**
 * PowerPoint's List tab beyond the first five (SMARTART-IMPORT.md §5.70), in
 * Office's own gallery order. Ordinary gallery layouts: MRAI stocks Office's
 * definition part for each, so they are inserted, exported as a real SmartArt
 * under their own urn and imported as themselves.
 *
 * Appended after the 39 rather than spliced into the List block above, and on
 * purpose: the owner's Office-made fixtures are indexed by the 39's position
 * (slide N = layout N), and a gallery is drawn CATEGORY first
 * ({@link smartArtLayoutsByCategory}), so the List tab shows these right after
 * the first five either way.
 *
 * Each `seedLevels` is the outline PowerPoint's own insert of the layout
 * carries (read off the data part Office wrote for it in the owner's
 * 134-layout Word file), and each `officeLevels` is Microsoft's own sentence
 * from *All SmartArt graphics, described*. Only Table List states a capacity.
 */
const OFFICE_LIST_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  layout({
    id: "alternating-hexagons",
    category: "list",
    seedItems: 6,
    seedLevels: [0, 1, 0, 1, 0, 1],
  }),
  layout({
    id: "lined-list",
    category: "list",
    seedItems: 4,
    seedLevels: [0, 1, 1, 1],
  }),
  layout({
    id: "vertical-bullet-list",
    category: "list",
    seedItems: 4,
    seedLevels: [0, 1, 0, 1],
  }),
  layout({
    id: "horizontal-bullet-list",
    category: "list",
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "square-accent-list",
    category: "list",
    seedItems: 8,
    seedLevels: [0, 1, 1, 1, 0, 1, 1, 1],
  }),
  layout({
    id: "stacked-list",
    category: "list",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "vertical-block-list",
    category: "list",
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "vertical-accent-list",
    category: "list",
    seedItems: 6,
    seedLevels: [0, 1, 0, 1, 0, 1],
  }),
  layout({
    id: "trapezoid-list",
    category: "list",
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
  }),
  layout({
    // The one with a capacity: its definition's root states `chMax val="1"`
    // and lays out only the first Level 1 item (`forEach axis="ch" cnt="1"`),
    // which is Microsoft's sentence too -- the first Level 1 line is the top
    // shape and ITS Level 2 text makes the lists.
    id: "table-list",
    category: "list",
    seedItems: 4,
    seedLevels: [0, 1, 1, 1],
    officeMaxItems: 1,
  }),
  layout({
    id: "vertical-curved-list",
    category: "list",
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
  layout({
    id: "vertical-circle-list",
    category: "list",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
];

/**
 * Office's fixed-capacity processes, converted but never offered.
 *
 * Every number below is Microsoft's own, quoted in
 * `docs/office-graphics/SMARTART-IMPORT.md` §5.34 and corroborated against the
 * count ladder in Office's own shipped `layoutDef`. `seedItems` is the
 * capacity, because these are never seeded from the gallery — it is the size a
 * reader should picture, and the catalog test holds it to the cap.
 */








const OFFICE_DEFINITION_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  layout({
    id: "arrow-ribbon",
    category: "relationship",
    seedItems: 2,
    seedLevels: [0, 0],
    officeMaxItems: 2,
  }),
  layout({
    id: "block-cycle",
    category: "cycle",
    seedItems: 5,
    seedLevels: [0, 0, 0, 0, 0],
  }),
  layout({
    id: "horizontal-labeled-hierarchy",
    category: "hierarchy",
    seedItems: 9,
    seedRootLabel: true,
    seedLevels: [0, 1, 2, 2, 1, 2, 0, 0, 0],
  }),
  layout({
    id: "detailed-process",
    category: "process",
    seedItems: 6,
    seedLevels: [0, 1, 0, 1, 0, 1],
  }),
  layout({
    id: "plus-and-minus",
    category: "relationship",
    seedItems: 2,
    seedLevels: [0, 0],
    officeMaxItems: 2,
  }),
  layout({
    id: "reverse-list",
    category: "relationship",
    seedItems: 2,
    seedLevels: [0, 0],
    officeMaxItems: 2,
  }),
  layout({
    id: "circular-bending-process",
    category: "process",
    seedItems: 9,
    seedLevels: [0, 0, 0, 0, 0, 0, 0, 0, 0],
  }),
  layout({
    id: "counterbalance-arrows",
    category: "relationship",
    seedItems: 2,
    seedLevels: [0, 0],
    officeMaxItems: 2,
  }),
  
  // flat, nested, CJK in a wide frame). Capacities are Microsoft's words
  // (support.microsoft.com, All SmartArt graphics, described); a definition
  // with no sample data of its own (`useDef`) seeds Office's three.
  layout({
    id: "basic-pie",
    category: "relationship",
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officeMaxItems: 7,
  }),
  layout({
    id: "circle-arrow-process",
    category: "process",
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
  layout({
    id: "circle-relationship",
    category: "relationship",
    seedItems: 3,
    seedRootLabel: true,
    seedLevels: [0, 1, 1],
    officeMaxItems: 1,
  }),
  layout({
    id: "horizontal-multi-level-hierarchy",
    category: "hierarchy",
    seedItems: 4,
    seedRootLabel: true,
    seedLevels: [0, 1, 1, 1],
  }),
  layout({
    id: "opposing-ideas",
    category: "relationship",
    seedItems: 4,
    seedLevels: [0, 1, 0, 1],
    officeMaxItems: 2,
  }),
  layout({
    id: "nested-target",
    category: "relationship",
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
    officeMaxItems: 3,
  }),
  layout({
    id: "vertical-arrow-list",
    category: "process",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "basic-hierarchy",
    category: "hierarchy",
    seedItems: 6,
    seedRootLabel: true,
    seedLevels: [0, 1, 2, 2, 1, 2],
  }),
  layout({
    id: "process-arrows",
    category: "process",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "radial-venn",
    category: "relationship",
    seedItems: 5,
    seedRootLabel: true,
    seedLevels: HUB_WITH_FOUR,
  }),
  layout({
    id: "closed-chevron-process",
    category: "process",
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
  layout({
    id: "hierarchy-list",
    category: "hierarchy",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "basic-equation",
    category: "relationship",
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
  
  // parity on its own definition. Seeds are their definitions' sample data
  // (Vertical Equation's `useDef` seeds three items); capacities are the
  // definitions' own `chMax` and Microsoft's words (support.microsoft.com,
  // SmartArt graphic descriptions).
  layout({
    id: "diverging-arrows",
    category: "relationship",
    seedItems: 2,
    seedLevels: [0, 0],
  }),
  layout({
    id: "converging-arrows",
    category: "relationship",
    seedItems: 2,
    seedLevels: [0, 0],
  }),
  layout({
    id: "circle-accent-timeline",
    category: "process",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "multidirectional-cycle",
    category: "cycle",
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
  layout({
    id: "vertical-equation",
    category: "relationship",
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
  layout({
    id: "table-hierarchy",
    category: "hierarchy",
    seedItems: 6,
    seedRootLabel: true,
    seedLevels: [0, 1, 2, 2, 1, 2],
  }),
  layout({
    id: "alternating-flow",
    category: "process",
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "process-list",
    category: "process",
    seedItems: 6,
    seedLevels: [0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "chevron-list",
    category: "process",
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "converging-radial",
    category: "relationship",
    seedItems: 4,
    seedRootLabel: true,
    seedLevels: [0, 1, 1, 1],
    officeMaxItems: 1,
  }),
  layout({
    id: "diverging-radial",
    category: "relationship",
    seedItems: 5,
    seedRootLabel: true,
    seedLevels: [0, 1, 1, 1, 1],
    officeMaxItems: 1,
  }),
  layout({
    id: "radial-cluster",
    category: "relationship",
    seedItems: 4,
    seedRootLabel: true,
    seedLevels: [0, 1, 1, 1],
    officeMaxItems: 1,
    officeMaxChildren: [7],
  }),
  layout({
    id: "random-to-result-process",
    category: "process",
    seedItems: 4,
    seedLevels: [0, 1, 0, 1],
  }),
  layout({
    id: "stacked-venn",
    category: "relationship",
    seedItems: 4,
    seedLevels: [0, 0, 0, 0],
    officeMaxItems: 7,
  }),
  layout({
    id: "step-down-process",
    category: "process",
    seedItems: 6,
    seedLevels: [0, 1, 0, 1, 0, 1],
  }),
  layout({
    id: "target-list",
    category: "relationship",
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
    officeMaxItems: 7,
  }),
  layout({
    id: "name-and-title-org",
    category: "hierarchy",
    ...MANAGER_WITH_ASSISTANT,
    seedRootLabel: true,
  }),
];











const CAPPED_AND_PICTURE_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  layout({
    id: "upward-arrow",
    category: "process",
    seedItems: 5,
    officeMaxItems: 5,
    officeMaxLevel: 0,
  }),
  layout({
    id: "descending-process",
    category: "process",
    seedItems: 7,
    officeMaxItems: 7,
    officeMaxLevel: 0,
  }),
  layout({
    id: "phased-process",
    category: "process",
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officeMaxItems: 3,
    officeMaxLevel: 1,
    officeMaxChildren: [4, 4],
  }),
  layout({
    id: "sub-step-process",
    category: "process",
    seedItems: 7,
    seedLevels: [0, 0, 0, 0, 0, 0, 0],
    officeMaxItems: 7,
    officeMaxLevel: 1,
  }),
  layout({
    id: "increasing-arrows-process",
    category: "process",
    seedItems: 5,
    seedLevels: [0, 0, 0, 0, 0],
    officeMaxItems: 5,
    officeMaxLevel: 1,
  }),
  layout({
    id: "increasing-circle-process",
    category: "process",
    seedItems: 7,
    seedLevels: [0, 0, 0, 0, 0, 0, 0],
    officeMaxItems: 7,
    officeMaxLevel: 1,
  }),
  layout({
    id: "pie-process",
    category: "process",
    seedItems: 7,
    seedLevels: [0, 0, 0, 0, 0, 0, 0],
    officeMaxItems: 7,
    officeMaxLevel: 1,
  }),
  // Office's other PICTURE layouts (SMARTART-IMPORT.md §5.39). Every one of
  // them binds its blip with `<dgm:forEach axis="ch" ptType="node">` — one
  // picture per Level 1 item, on the item's own node — so every one states
  // `officePictureLevel: 0`, and none needs a row in `picture-bindings.json`.
  // Each `officeLevels` is Microsoft's own sentence from *All SmartArt
  // graphics, described*, which is also what the engine draws. No capacity:
  // Microsoft states a number for none of the twelve and none of their layout
  // definitions carries a count ladder.
  layout({
    id: "picture-lineup",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "picture-grid",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "bending-picture-blocks",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "bending-picture-caption",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "bending-picture-semi-transparent-text",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "picture-strips",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "bending-picture-caption-list",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "titled-picture-blocks",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "framed-text-picture",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "vertical-picture-list",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "vertical-picture-accent-list",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "bending-picture-accent-list",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  // The last nine of the ordinary-node-binding picture layouts
  // (SMARTART-IMPORT.md §5.46). `officeLevels` is again Microsoft's own
  // sentence from *All SmartArt graphics, described*, and every
  // `officePictureLevel` is the MEASURED `ch`-frame depth of that layout's own
  // `blipPhldr` shape, not an assumption -- which is how Titled Picture Accent
  // List turned out to be a LEVEL 2 binding after §5.44 had listed it with the
  // rest. Their `gte 1` / `gt 1` / `gt 5` branches are drawing variants, not
  // capacities (the `cycle2` precedent in §5.34), and Microsoft states a
  // number for none of them, so none declares `officeMaxItems`.
  layout({
    id: "captioned-pictures",
    category: "picture",
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "snapshot-picture-list",
    category: "picture",
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "title-picture-lineup",
    category: "picture",
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "picture-accent-blocks",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "picture-accent-list",
    category: "picture",
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "alternating-picture-circles",
    category: "picture",
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "continuous-picture-list",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "horizontal-picture-list",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "titled-picture-accent-list",
    category: "picture",
    seedItems: 3,
    seedLevels: [0, 1, 1],
    // MEASURED: two `ch` frames enclose this layout's `blipPhldr`, and
    // Microsoft says the same thing in words -- "an accent picture for each
    // Level 2 text". A picture on its Level 1 item is one Office does not
    // draw, which only a single level can express.
    officePictureLevel: 1,
  }),
  // The last seven picture layouts the corpus holds (SMARTART-IMPORT.md
  // §5.51). Four of them bind their blip to a sibling TRANSITION -- the
  // mechanism `picture-bindings.json` states and §5.21 proved -- and three
  // are ordinary node bindings that needed measuring rather than assuming:
  // Accented Picture binds its SUBJECT to a transition and every other item
  // to its own node, Circle Picture Hierarchy carries FOUR `blipPhldr`
  // shapes (one per hierarchy level, so no single level is "the" one), and
  // Radial List carries exactly one, on the FIRST item and nowhere else.
  //
  // `officeMaxItems` is declared only where Microsoft states a number in
  // words: "up to eight Level 1 pictures" and "up to five pictures". The
  // count ladders in the other five are drawing variants, the `cycle2`
  // precedent of §5.34.
  layout({
    id: "bubble-picture-list",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
    officeMaxItems: 8,
  }),
  layout({
    id: "hexagon-cluster",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "spiral-picture",
    category: "picture",
    seedItems: 5,
    seedLevels: [0, 0, 0, 0, 0],
    officePictureLevel: 0,
    officeMaxItems: 5,
  }),
  layout({
    id: "ascending-picture-accent-process",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "accented-picture",
    category: "picture",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
  }),
  layout({
    id: "circle-picture-hierarchy",
    category: "picture",
    seedItems: 4,
    seedRootLabel: true,
    seedLevels: ROOT_WITH_THREE,
    // NO `officePictureLevel`: this is the one picture layout whose definition
    // gives EVERY level a slot -- four `blipPhldr` shapes, `image` … `image4`,
    // one per `hierRoot` -- so an item at any depth the tree draws has a
    // picture and none of them is the exception the field exists to name.
    officeMaxLevel: 3,
  }),
  layout({
    id: "radial-list",
    convertOnly: true,
    // Office files this one under RELATIONSHIP only -- it is not in the
    // Picture tab, and the corpus's own category line says 关系 -- while its
    // definition still carries a picture slot. The tab and the slot are two
    // facts, so each is stated where it belongs.
    category: "relationship",
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
    // NO item has a slot: the centre's picture is the DIAGRAM's, kept on the
    // `visible` presentation point keyed to the document point, and every
    
    
    // An item's own picture fills that item's circle in PowerPoint
    // (sa36-radial2b), which a page read under readings row 1005 draws so;
    // anywhere else such a diagram stays as the file drew it.
    officePictureItems: 0,
  }),
  // Office's ICON layouts (2018; each definition's only `dgm:cat` is `icon`,
  // a tab PowerPoint's Choose a SmartArt Graphic dialog does not have: they
  // reach a deck through Designer and Convert to SmartArt). Each item's
  // `iconRect` is a picture placeholder (`blipPhldr`) one `ch` frame deep,
  // which PowerPoint fills with an SVG icon and its PNG fallback on the
  // item's `iconRect` presentation point. Filed under List, which is what
  // each description says they are; the tab of a convert-only layout is
  // never shown.
  layout({
    id: "icon-vertical-solid-list",
    category: "list",
    convertOnly: true,
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "icon-circle-list",
    category: "list",
    convertOnly: true,
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
    // "corresponding text box shows Level 1 text": its words are presented
    // `presOf axis="self"` and nothing presents a descendant.
    officeMaxLevel: 0,
  }),
  layout({
    id: "icon-circle-label-list",
    category: "list",
    convertOnly: true,
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officePictureLevel: 0,
    // Its caption is presented `presOf axis="self"`: Level 1 only.
    officeMaxLevel: 0,
  }),
  
  // the first three: each definition's only `dgm:cat` is `icon`, its
  // `iconRect` the item's picture one `ch` frame deep, presenting nothing.
  layout({
    id: "icon-label-description-list",
    category: "list",
    convertOnly: true,
    // `sampData useDef="1"`: Office's default of three Level 1 shapes.
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "centered-icon-label-description-list",
    category: "list",
    convertOnly: true,
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
  }),
  layout({
    id: "icon-leaf-label-list",
    category: "list",
    convertOnly: true,
    seedItems: 3,
    seedLevels: [0, 0, 0],
    officePictureLevel: 0,
    // Its caption is presented `presOf axis="self"`: Level 1 only.
    officeMaxLevel: 0,
  }),
];

/**
 * A diagram laid out from its own file's definition
 * (`SMARTART_OFFICE_DEFINITION_KIND`): convert-only, its category never shown
 * (Change Layout lists it as the current layout, under List).
 */
const OFFICE_FILE_DEFINITION_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  layout({ id: "office-definition", category: "list", convertOnly: true, seedItems: 3, seedLevels: [0, 0, 0] }),
];

/**
 * Office's 2016-2017 layouts with no picture (`SMARTART_OFFICE_DESIGNER_KINDS`),
 * reached through PowerPoint Designer and Convert to SmartArt. Convert-only:
 * MRAI stocks no preset for them, so a diagram of one converts on open and
 * keeps the definition its package holds. Each one's tab is its definition's
 * first `dgm:cat`; seeds are each definition's `dgm:sampData`, and
 * `officeLevels` Microsoft's own `dgm:desc`.
 */
const OFFICE_DESIGNER_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  layout({
    id: "basic-linear-process-numbered",
    category: "process",
    convertOnly: true,
    seedItems: 6,
    seedLevels: [0, 1, 0, 1, 0, 1],
  }),
  layout({
    id: "vertical-solid-action-list",
    category: "list",
    convertOnly: true,
    seedItems: 10,
    seedLevels: [0, 1, 0, 1, 0, 1, 0, 1, 0, 1],
  }),
  layout({
    // Its first `dgm:cat` is `timeline`, a tab the dialog does not have; the
    // second is Process, which is where it is filed.
    id: "drop-pin-timeline",
    category: "process",
    convertOnly: true,
    seedItems: 6,
    seedLevels: [0, 1, 0, 1, 0, 1],
  }),
];

/**
 * Office.com layouts (`SMARTART_OFFICE_ONLINE_KINDS`): drawn only from the
 * definition each deck carries, under the tab its definition's FIRST
 * `dgm:cat` names (the second is `officeonline`, PowerPoint's Office.com tab,
 * which MRAI does not ship, so they are convert-only). Capacities are
 * Microsoft's words in each definition's `dgm:desc` and its root `chMax`.
 */
const OFFICE_ONLINE_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  layout({
    id: "interconnected-rings",
    category: "relationship",
    convertOnly: true,
    seedItems: 4,
    seedLevels: FLAT_FOUR,
    officeMaxItems: 7,
  }),
  layout({
    id: "hexagon-radial",
    category: "cycle",
    convertOnly: true,
    seedItems: 7,
    seedRootLabel: true,
    seedLevels: [0, 1, 1, 1, 1, 1, 1],
    // The root's `chMax val="1"`: one centre; "limited to six Level 2 shapes".
    officeMaxItems: 1,
    officeMaxChildren: [6],
  }),
  layout({
    id: "tabbed-arc",
    category: "relationship",
    convertOnly: true,
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
  layout({
    id: "tab-list",
    category: "list",
    convertOnly: true,
    seedItems: 9,
    seedLevels: [0, 1, 1, 0, 1, 1, 0, 1, 1],
  }),
  layout({
    id: "chevron-accent-process",
    category: "process",
    convertOnly: true,
    seedItems: 3,
    seedLevels: [0, 0, 0],
  }),
];

/**
 * Every layout an author can insert: the 39, then the List-tab layouts of
 * §5.70, then those drawn only from their Office definitions. A gallery reads it category first ({@link smartArtLayoutsByCategory},
 * {@link OFFICE_SMARTART_GALLERY_GROUPS}); the flat order keeps the 39 where
 * the Office-made fixtures index them.
 */
export const OFFICE_SMARTART_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  ...LAYOUTS,
  ...OFFICE_LIST_LAYOUTS,
  ...OFFICE_DEFINITION_LAYOUTS,
  ...CAPPED_AND_PICTURE_LAYOUTS.filter((entry) => !entry.convertOnly),
];

/**
 * The layouts MRAI draws and owns but does not OFFER: an imported diagram of
 * one of these converts and round-trips as itself, and nothing can author one
 * from nothing because there is no stocked Office preset to write.
 */
export const OFFICE_SMARTART_CONVERT_ONLY_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  ...CAPPED_AND_PICTURE_LAYOUTS.filter((entry) => entry.convertOnly),
  ...OFFICE_ONLINE_LAYOUTS,
  ...OFFICE_DESIGNER_LAYOUTS,
  ...OFFICE_FILE_DEFINITION_LAYOUTS,
];

/** Everything {@link officeSmartArtLayout} resolves: the gallery and the rest. */
const ALL_LAYOUTS: readonly OfficeSmartArtLayout[] = [
  ...OFFICE_SMARTART_LAYOUTS,
  ...OFFICE_SMARTART_CONVERT_ONLY_LAYOUTS,
];

const BY_ANY_ID = new Map<string, OfficeSmartArtLayout>(
  ALL_LAYOUTS.flatMap((entry) =>
    [entry.id, entry.kind, ...entry.aliases].map(
      (id) => [id, entry] as const,
    ),
  ),
);

/** Resolve a canonical id, an engine kind or a legacy alias to one layout. */
export function officeSmartArtLayout(
  idOrAlias: unknown,
): OfficeSmartArtLayout | null {
  return typeof idOrAlias === "string"
    ? (BY_ANY_ID.get(idOrAlias) ?? null)
    : null;
}

/**
 * The engine kind for any accepted id. This is the ONLY place that knows the
 * alias table, so a host never hard-codes `"venn" -> "relationship-venn"`.
 */
export function canonicalSmartArtKind(
  idOrAlias: unknown,
): OfficeSmartArtKind | null {
  return officeSmartArtLayout(idOrAlias)?.kind ?? null;
}

/** The canonical (Docs) layout id for any accepted id. */
export function canonicalSmartArtLayoutId(idOrAlias: unknown): string | null {
  return officeSmartArtLayout(idOrAlias)?.id ?? null;
}

/**
 * Which items of this outline a layout drawn by MRAI's own builder DRAWS, or
 * `null` when it draws them all.
 *
 * Office's capped layouts do not complain: the items past the cap simply do
 * not appear ("Unused text does not appear, but remains available if you
 * switch layouts"). An item is drawn when its Level 1 item is one of the
 * first `officeMaxItems`, and, under a layout that caps them, when it is one
 * of the first `officeMaxChildren` Level 2 items of that Level 1 item; a
 * deeper item follows its parent. The rest stay in the outline as hidden
 * members (`buildSmartArt`), never refused.
 *
 * A layout drawn from Office's own definition is not asked: the definition
 * itself says what it draws (`ms-engine/scene`).
 */
export function smartArtLayoutDrawnItems(
  idOrAlias: unknown,
  levels: readonly number[],
): boolean[] | null {
  const entry = officeSmartArtLayout(idOrAlias);
  if (!entry || (entry.officeMaxItems === undefined && entry.officeMaxChildren === undefined)) {
    return null;
  }
  const drawn: boolean[] = [];
  let topLevel = -1;
  let children = 0;
  // Whether the item one level up is drawn, by level.
  const parents: boolean[] = [];
  for (const level of levels) {
    let shown: boolean;
    if (level <= 0) {
      topLevel += 1;
      children = 0;
      shown = entry.officeMaxItems === undefined || topLevel < entry.officeMaxItems;
    } else if (level === 1) {
      children += 1;
      const cap = entry.officeMaxChildren?.[topLevel];
      shown = (parents[0] ?? false) && (cap === undefined || children <= cap);
    } else {
      shown = parents[level - 1] ?? false;
    }
    parents[Math.max(0, level)] = shown;
    parents.length = Math.max(0, level) + 1;
    drawn.push(shown);
  }
  return drawn.every(Boolean) ? null : drawn;
}

export interface OfficeSmartArtCategoryGroup {
  category: OfficeSmartArtCategoryId;
  categoryKey: string;
  layouts: readonly OfficeSmartArtLayout[];
}

/**
 * Where Office files each layout: the tabs of its definition's `dgm:catLst`
 * that MRAI ships, with each one's `pri` -- the place PowerPoint's Choose a
 * SmartArt Graphic dialog gives it in that tab (SMARTART-GAP.md §1.7) -- by
 * catalog id and by alias card id; and the aliases PowerPoint lists as layouts
 * of their own, each with its definition's own sample. Generated from
 * `ms-definitions.json` and pinned to it by
 * `tests/unit/office-smartart-gallery-facts.test.ts`.
 */
const GALLERY_FACTS = galleryFacts as unknown as {
  readonly tabs: Readonly<Record<string, readonly (readonly [string, number])[]>>;
  readonly aliasCards: readonly {
    readonly id: string;
    readonly layoutId: string;
    readonly layoutUrn: string;
    readonly suffix: string;
    readonly seedLevels?: readonly number[];
    readonly seedAssistants?: readonly number[];
  }[];
};

const CATEGORY_SET: ReadonlySet<string> = new Set(OFFICE_SMARTART_CATEGORY_IDS);

/**
 * The tabs a layout or alias card is listed in, with its place in each. A
 * layout whose definition files it under none MRAI ships is listed in its own
 * category, last.
 */
function galleryTabs(id: string, fallback: OfficeSmartArtCategoryId): readonly (readonly [OfficeSmartArtCategoryId, number])[] {
  const stated = (GALLERY_FACTS.tabs[id] ?? []).filter(
    (tab): tab is readonly [OfficeSmartArtCategoryId, number] => CATEGORY_SET.has(tab[0]),
  );
  return stated.length ? stated : [[fallback, Number.POSITIVE_INFINITY]];
}

/**
 * One card of the Insert ▸ SmartArt gallery. `id` is what the card hands its
 * host's insert: the catalog id, or `<catalog id>@<urn suffix>` for an Office
 * alias PowerPoint lists as a layout of its own (Continuous Block Process
 * beside Basic Process), which inserts as that alias, laid out from its own
 * definition and recorded on every member (`smartArtLayoutUrn`).
 */
export interface OfficeSmartArtGalleryCard {
  readonly id: string;
  /** The catalog layout the card inserts (the alias's own catalog layout). */
  readonly layoutId: string;
  /** The alias's urn; absent for a catalog layout. */
  readonly layoutUrn?: string;
  /** The alias's urn suffix, its message key (`smartArtAliasLayoutName_<suffix>`). */
  readonly aliasSuffix?: string;
}

const ALIAS_CARDS: ReadonlyMap<string, (typeof GALLERY_FACTS.aliasCards)[number]> = new Map(
  GALLERY_FACTS.aliasCards
    .filter((card) => OFFICE_SMARTART_LAYOUTS.some((entry) => entry.id === card.layoutId))
    .map((card) => [card.id, card] as const),
);

/** A gallery card by the id it hands the insert, or `null` for no card. */
export function officeSmartArtGalleryCard(id: unknown): OfficeSmartArtGalleryCard | null {
  if (typeof id !== "string") return null;
  const alias = ALIAS_CARDS.get(id);
  if (alias) return { id, layoutId: alias.layoutId, layoutUrn: alias.layoutUrn, aliasSuffix: alias.suffix };
  const entry = OFFICE_SMARTART_LAYOUTS.find((layout) => layout.id === id);
  return entry ? { id, layoutId: entry.id } : null;
}

/**
 * The alias card that inserts catalog layout `layoutId` as the Office alias
 * `layoutUrn`, or `null`: what a preview drawn "as that alias" seeds from, so
 * the card shows the alias's own sample, as its insert does.
 */
export function officeSmartArtAliasCardId(layoutId: unknown, layoutUrn: unknown): string | null {
  const entry = officeSmartArtLayout(layoutId);
  if (!entry || typeof layoutUrn !== "string") return null;
  for (const card of ALIAS_CARDS.values()) {
    if (card.layoutId === entry.id && card.layoutUrn === layoutUrn) return card.id;
  }
  return null;
}

/**
 * Each tab's cards in PowerPoint's order: a layout appears in EVERY tab its
 * definition names (Segmented Process in Process and in List), sorted by its
 * `pri` there; equal places keep catalog order, an alias after its layout.
 */
function galleryTabCards(withAliases: boolean): readonly { category: OfficeSmartArtCategoryId; cards: readonly OfficeSmartArtGalleryCard[] }[] {
  const sources: { card: OfficeSmartArtGalleryCard; fallback: OfficeSmartArtCategoryId }[] = OFFICE_SMARTART_LAYOUTS.flatMap((entry) => [
    { card: { id: entry.id, layoutId: entry.id } as OfficeSmartArtGalleryCard, fallback: entry.category },
    ...(withAliases
      ? [...ALIAS_CARDS.values()]
          .filter((alias) => alias.layoutId === entry.id)
          .map((alias) => ({
            card: { id: alias.id, layoutId: alias.layoutId, layoutUrn: alias.layoutUrn, aliasSuffix: alias.suffix },
            fallback: entry.category,
          }))
      : []),
  ]);
  return OFFICE_SMARTART_CATEGORY_IDS.map((category) => ({
    category,
    cards: sources
      .flatMap(({ card, fallback }, order) => {
        const tab = galleryTabs(card.id, fallback).find(([id]) => id === category);
        return tab ? [{ card, pri: tab[1], order }] : [];
      })
      .sort((a, b) => a.pri - b.pri || a.order - b.order)
      .map(({ card }) => card),
  }));
}

/**
 * The gallery grouping: every category, with every catalog layout its
 * definition files under it, in Office's order there (a layout of several
 * tabs is in each of them).
 */
export function smartArtLayoutsByCategory(): readonly OfficeSmartArtCategoryGroup[] {
  return galleryTabCards(false).map(({ category, cards }) => ({
    category,
    categoryKey: `toolbar.smartArtCategory_${category}`,
    layouts: cards.map((card) => officeSmartArtLayout(card.layoutId)!),
  }));
}

/** The tabs a catalog layout (or alias card) is listed in, and its place in each. */
export function officeSmartArtGalleryTabs(
  id: string,
): readonly { readonly category: OfficeSmartArtCategoryId; readonly pri: number }[] {
  const card = officeSmartArtGalleryCard(id);
  const entry = officeSmartArtLayout(card?.layoutId ?? id);
  if (!entry) return [];
  return galleryTabs(card?.id ?? entry.id, entry.category).map(([category, pri]) => ({ category, pri }));
}

/** One tab of the Insert ▸ SmartArt gallery: its category and its cards. */
export interface OfficeSmartArtGalleryGroup {
  readonly id: OfficeSmartArtCategoryId;
  readonly layouts: readonly OfficeSmartArtGalleryCard[];
}

const freezeGroups = (withAliases: boolean): readonly OfficeSmartArtGalleryGroup[] =>
  Object.freeze(
    galleryTabCards(withAliases)
      .map(({ category, cards }) =>
        Object.freeze({ id: category, layouts: Object.freeze(cards.map((card) => Object.freeze({ ...card }))) }),
      )
      .filter((group) => group.layouts.length > 0),
  );

/**
 * Office's one SmartArt gallery: every insertable layout and every alias
 * PowerPoint lists, under each tab its definition names, in PowerPoint's
 * order. Docs, the deck and the worksheet read this one, so a layout added to
 * the catalog appears in all three galleries at once.
 */
export const OFFICE_SMARTART_GALLERY_GROUPS: readonly OfficeSmartArtGalleryGroup[] = freezeGroups(true);

/**
 * The same tabs with the catalog layouts only: for a picker whose command
 * builds a catalog layout and nothing else (Convert to SmartArt).
 */
export const OFFICE_SMARTART_LAYOUT_GALLERY_GROUPS: readonly OfficeSmartArtGalleryGroup[] = freezeGroups(false);

/** Every insertable catalog layout once, in gallery order (first tab first). */
export const OFFICE_SMARTART_GALLERY_LAYOUT_IDS: readonly string[] = Object.freeze([
  ...new Set(OFFICE_SMARTART_LAYOUT_GALLERY_GROUPS.flatMap((group) => group.layouts.map((card) => card.layoutId))),
]);

/** The layout a SmartArt command inserts when no card was picked. */
export const OFFICE_SMARTART_DEFAULT_LAYOUT_ID: string =
  OFFICE_SMARTART_GALLERY_LAYOUT_IDS[0]!;

/**
 * The placeholder text one editor seeds a SmartArt item with. Office writes the
 * literal `[Text]`; MRAI numbers its items so a fresh diagram reads as an
 * outline. ONE rule for the three editors (`useOfficeSmartArtPlaceholders`):
 * the insert, the gallery card, the Change Layout thumbnail, Add Shape and a
 * converted diagram's empty point all say the same words.
 */
export interface OfficeSmartArtPlaceholders {
  readonly itemLabel: (index: number) => string;
  readonly rootLabel: string;
}

export interface OfficeSmartArtSeed {
  /** The engine kind to build. */
  kind: OfficeSmartArtKind;
  /** Localized placeholder labels, one per seeded item. */
  items: string[];
  /** Outline levels for those items, when the layout declares them. */
  levels?: number[];
  /** Which items are assistants (an organization chart's sample), by index. */
  assistants?: boolean[];
  /**
   * The Office alias an alias card inserts as (`smartArtLayoutUrn`): the
   * build lays the diagram out from that alias's own definition and records
   * it on every member. Absent for a catalog layout.
   */
  layoutUrn?: string;
}

/**
 * What a layout id or an alias card id seeds from: the catalog layout, and for
 * an alias card its own sample (Office's Vertical Bending Process inserts
 * nine shapes where Basic Process inserts three) and urn.
 */
function seedSource(idOrAlias: unknown): {
  entry: OfficeSmartArtLayout;
  seedItems: number;
  seedLevels?: readonly number[];
  seedAssistants?: readonly number[];
  layoutUrn?: string;
} | null {
  const alias = typeof idOrAlias === "string" ? ALIAS_CARDS.get(idOrAlias) : undefined;
  const entry = officeSmartArtLayout(alias ? alias.layoutId : idOrAlias);
  if (!entry) return null;
  const own = { entry, seedItems: entry.seedItems, seedLevels: entry.seedLevels, seedAssistants: entry.seedAssistants };
  if (!alias) return own;
  // A definition naming no sample (`useDef`) inserts the layout's own.
  return alias.seedLevels?.length
    ? { entry, seedItems: alias.seedLevels.length, seedLevels: alias.seedLevels, seedAssistants: alias.seedAssistants, layoutUrn: alias.layoutUrn }
    : { ...own, layoutUrn: alias.layoutUrn };
}

/** How many items a layout's (or an alias card's) own insert seeds. */
export function officeSmartArtSeedItemCount(idOrAlias: unknown): number | null {
  return seedSource(idOrAlias)?.seedItems ?? null;
}

/**
 * The seed an insert and both preview sizes share, so what the gallery card
 * shows is exactly what lands in the document. Labels stay localized: the host
 * supplies them (Office's own placeholder is the literal `[Text]`).
 *
 * `itemCount` seeds that many items instead of the layout's own sample size:
 * an insert that already has its items (the AI's `document.insert.smartart`)
 * draws every one of them, as Office's text pane adds a shape per line (H41
 * B05: five milestones in a 3-shape timeline kept the first three and dropped
 * the rest without a word). The layout's sample outline and assistants only
 * fit its own size; another count is flat, under the root when the layout has
 * one.
 */
export function officeSmartArtSeed(
  idOrAlias: unknown,
  itemLabel: (index: number) => string,
  rootLabel: string,
  itemCount?: number,
): OfficeSmartArtSeed | null {
  const source = seedSource(idOrAlias);
  if (!source) return null;
  const { entry } = source;
  const count =
    itemCount !== undefined && Number.isInteger(itemCount) && itemCount > 0
      ? itemCount
      : source.seedItems;
  const items = officeSmartArtSeedLabels(idOrAlias, count, itemLabel, rootLabel);
  if (!items) return null;
  const sampleShape = count === source.seedItems;
  const levels = sampleShape
    ? source.seedLevels
    : entry.seedRootLabel
      ? items.map((_, index) => (index === 0 ? 0 : 1))
      : undefined;
  return {
    kind: entry.kind,
    items,
    ...(levels ? { levels: [...levels] } : {}),
    ...(sampleShape && source.seedAssistants?.length
      ? { assistants: items.map((_, index) => source.seedAssistants!.includes(index)) }
      : {}),
    ...(source.layoutUrn ? { layoutUrn: source.layoutUrn } : {}),
  };
}

/**
 * The placeholder labels an insert seeds, for ANY item count.
 *
 * {@link officeSmartArtSeed} is this at the layout's own seed size. An importer
 * reuses it so an Office point whose text is the empty `[文本]` placeholder
 * becomes the same label an inserted diagram shows at that index, rather than a
 * blank the editor renders as an empty text box.
 */
export function officeSmartArtSeedLabels(
  idOrAlias: unknown,
  count: number,
  itemLabel: (index: number) => string,
  rootLabel: string,
): string[] | null {
  const entry = seedSource(idOrAlias)?.entry;
  if (!entry || !Number.isInteger(count) || count < 0) return null;
  // Children are numbered from 1 whether or not the first item is the root.
  let ordinal = 0;
  return Array.from({ length: count }, (_, index) =>
    entry.seedRootLabel && index === 0 ? rootLabel : itemLabel((ordinal += 1)),
  );
}
