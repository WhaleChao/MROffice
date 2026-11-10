/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isOpenBasicShapeKind, isPresentationConnectorShapeKind, normalizeShapeKind, type ShapeKind } from "./presentation-shapes";
import {
  PRESENTATION_CUSTOM_GEOMETRY_TOOLS,
  normalizePresentationCustomPath,
  type PresentationCustomGeometryTool,
} from "./presentation-freeform";
import {
  PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS,
  type PresentationActionButtonShapeKind,
  type PresentationRelativeClickAction,
} from "./presentation-block-fx";

/**
 * Versioned PowerPoint shape-authoring catalog.
 *
 * `drawingMlPreset` is the exact `a:prstGeom@prst` identity written by the
 * native PPTX adapters.  Gallery order, search, recents and import projection
 * all consume this catalog, so a visual tile can never silently point at a
 * different DrawingML geometry.
 */
export const PRESENTATION_SHAPE_CATALOG_VERSION =
  "drawingml-prstgeom-v14" as const;

export const PRESENTATION_SHAPE_CATEGORY_IDS = [
  "lines",
  "rectangles",
  "basic",
  "arrows",
  "equation",
  "flowchart",
  "stars-banners",
  "callouts",
  "action-buttons",
] as const;

export type PresentationShapeCategoryId =
  (typeof PRESENTATION_SHAPE_CATEGORY_IDS)[number];

export type PresentationShapeCatalogSelection =
  | ShapeKind
  | PresentationCustomGeometryTool;

export interface PresentationPresetShapeCatalogEntry {
  readonly id: `shape.${ShapeKind}`;
  readonly kind: ShapeKind;
  readonly selection: ShapeKind;
  readonly category: PresentationShapeCategoryId;
  readonly drawingMlPreset: string;
  readonly geometry: "preset" | "connector";
  readonly defaultClickAction?: PresentationRelativeClickAction;
}

export interface PresentationCustomToolCatalogEntry {
  readonly id: `shape.${PresentationCustomGeometryTool}`;
  readonly tool: PresentationCustomGeometryTool;
  readonly selection: PresentationCustomGeometryTool;
  readonly category: "lines";
  readonly geometry: "custom-tool";
}

export type PresentationShapeCatalogEntry =
  | PresentationPresetShapeCatalogEntry
  | PresentationCustomToolCatalogEntry;

/**
 * Narrow a catalog member to a durable DrawingML-backed shape.
 *
 * Custom geometry tools deliberately do not carry a fake `kind` or
 * `drawingMlPreset`, so consumers must make this distinction before using the
 * native preset contract.
 */
export function isPresentationPresetShapeCatalogEntry(
  catalogEntry: PresentationShapeCatalogEntry,
): catalogEntry is PresentationPresetShapeCatalogEntry {
  return catalogEntry.geometry !== "custom-tool";
}

function entry(
  category: PresentationShapeCategoryId,
  kind: ShapeKind,
  drawingMlPreset: string,
  geometry: PresentationPresetShapeCatalogEntry["geometry"] = "preset",
): PresentationPresetShapeCatalogEntry {
  return Object.freeze({
    id: `shape.${kind}`,
    kind,
    selection: kind,
    category,
    drawingMlPreset,
    geometry,
  });
}

function customToolEntry(
  tool: PresentationCustomGeometryTool,
): PresentationCustomToolCatalogEntry {
  return Object.freeze({
    id: `shape.${tool}`,
    tool,
    selection: tool,
    category: "lines",
    geometry: "custom-tool",
  });
}

function actionEntry(
  kind: PresentationActionButtonShapeKind,
  drawingMlPreset: string,
): PresentationPresetShapeCatalogEntry {
  return Object.freeze({
    id: `shape.${kind}`,
    kind,
    selection: kind,
    category: "action-buttons",
    drawingMlPreset,
    geometry: "preset",
    ...(kind in PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS
      ? {
          defaultClickAction:
            PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS[
              kind as keyof typeof PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS
            ],
        }
      : {}),
  });
}

/** The authenticated 196-item authoring set; durable custom geometry is hidden. */
export const PRESENTATION_SHAPE_CATALOG = Object.freeze([
  entry("lines", "line", "line", "connector"),
  entry("lines", "arrow", "line", "connector"),
  entry("lines", "double-arrow", "line", "connector"),
  entry("lines", "elbow-connector", "bentConnector3", "connector"),
  entry("lines", "elbow-arrow-connector", "bentConnector3", "connector"),
  entry("lines", "elbow-double-arrow-connector", "bentConnector3", "connector"),
  entry("lines", "curved-connector", "curvedConnector3", "connector"),
  entry("lines", "curved-arrow-connector", "curvedConnector3", "connector"),
  entry(
    "lines",
    "curved-double-arrow-connector",
    "curvedConnector3",
    "connector",
  ),
  // The rest of DrawingML's connector presets. Office re-routes a connector by
  // rewriting `@prst` to the variant with the right number of segments, so a
  // real document holds these even though no gallery item inserts one. They
  // were the last eight `prst` values with no kind at all (audit §2.6/§7).
  entry("lines", "elbow-connector-2", "bentConnector2", "connector"),
  entry("lines", "elbow-connector-4", "bentConnector4", "connector"),
  entry("lines", "elbow-connector-5", "bentConnector5", "connector"),
  entry("lines", "curved-connector-2", "curvedConnector2", "connector"),
  entry("lines", "curved-connector-4", "curvedConnector4", "connector"),
  entry("lines", "curved-connector-5", "curvedConnector5", "connector"),
  entry("lines", "straight-connector", "straightConnector1", "connector"),
  // An autoshape, not a connector: `lineInv`'s geometry is fillable and carries
  // its own connection sites, and Office's shape enums list it while no
  // connector enum does.
  entry("lines", "line-inverse", "lineInv"),
  ...PRESENTATION_CUSTOM_GEOMETRY_TOOLS.map(customToolEntry),

  entry("rectangles", "rectangle", "rect"),
  entry("rectangles", "rounded-rectangle", "roundRect"),
  entry("rectangles", "single-round-corner", "round1Rect"),
  entry("rectangles", "same-side-corners-rounded", "round2SameRect"),
  entry("rectangles", "diagonal-corners-rounded", "round2DiagRect"),
  entry("rectangles", "single-corner-snipped", "snip1Rect"),
  entry("rectangles", "same-side-corners-snipped", "snip2SameRect"),
  entry("rectangles", "diagonal-corners-snipped", "snip2DiagRect"),
  entry("rectangles", "single-corner-rounded-snipped", "snipRoundRect"),

  entry("basic", "ellipse", "ellipse"),
  entry("basic", "triangle", "triangle"),
  entry("basic", "right-triangle", "rtTriangle"),
  entry("basic", "parallelogram", "parallelogram"),
  entry("basic", "trapezoid", "trapezoid"),
  entry("basic", "non-isosceles-trapezoid", "nonIsoscelesTrapezoid"),
  entry("basic", "diamond", "diamond"),
  entry("basic", "pentagon", "pentagon"),
  entry("basic", "hexagon", "hexagon"),
  entry("basic", "heptagon", "heptagon"),
  entry("basic", "octagon", "octagon"),
  entry("basic", "decagon", "decagon"),
  entry("basic", "dodecagon", "dodecagon"),
  entry("basic", "pie", "pie"),
  entry("basic", "pie-wedge", "pieWedge"),
  entry("basic", "chord", "chord"),
  entry("basic", "teardrop", "teardrop"),
  entry("basic", "frame", "frame"),
  entry("basic", "half-frame", "halfFrame"),
  entry("basic", "corner", "corner"),
  entry("basic", "diagonal-stripe", "diagStripe"),
  entry("basic", "plus", "plus"),
  entry("basic", "plaque", "plaque"),
  entry("basic", "can", "can"),
  entry("basic", "cube", "cube"),
  entry("basic", "bevel", "bevel"),
  entry("basic", "donut", "donut"),
  entry("basic", "no-symbol", "noSmoking"),
  entry("basic", "block-arc", "blockArc"),
  entry("basic", "folded-corner", "foldedCorner"),
  entry("basic", "smiley", "smileyFace"),
  entry("basic", "heart", "heart"),
  entry("basic", "lightning-bolt", "lightningBolt"),
  entry("basic", "sun", "sun"),
  entry("basic", "moon", "moon"),
  entry("basic", "cloud", "cloud"),
  // Basic Shapes Word, Excel and PowerPoint all offer that this catalog had
  // never carried, so no host could import them (SHAPE-IMPORT-AUDIT.md §2.6).
  entry("basic", "gear-6", "gear6"),
  entry("basic", "gear-9", "gear9"),
  entry("basic", "funnel", "funnel"),
  entry("basic", "chart-plus", "chartPlus"),
  entry("basic", "chart-star", "chartStar"),
  entry("basic", "chart-x", "chartX"),
  entry("basic", "corner-tabs", "cornerTabs"),
  entry("basic", "square-tabs", "squareTabs"),
  entry("basic", "plaque-tabs", "plaqueTabs"),
  entry("basic", "arc", "arc"),
  entry("basic", "double-bracket", "bracketPair"),
  entry("basic", "double-brace", "bracePair"),
  entry("basic", "left-bracket", "leftBracket"),
  entry("basic", "right-bracket", "rightBracket"),
  entry("basic", "left-brace", "leftBrace"),
  entry("basic", "right-brace", "rightBrace"),
  entry("arrows", "right-arrow", "rightArrow"),
  entry("arrows", "left-arrow", "leftArrow"),
  entry("arrows", "up-arrow", "upArrow"),
  entry("arrows", "down-arrow", "downArrow"),
  entry("arrows", "left-right-arrow", "leftRightArrow"),
  entry("arrows", "up-down-arrow", "upDownArrow"),
  entry("arrows", "quad-arrow", "quadArrow"),
  entry("arrows", "left-right-up-arrow", "leftRightUpArrow"),
  entry("arrows", "bent-arrow", "bentArrow"),
  entry("arrows", "u-turn-arrow", "uturnArrow"),
  entry("arrows", "left-up-arrow", "leftUpArrow"),
  entry("arrows", "bent-up-arrow", "bentUpArrow"),
  entry("arrows", "curved-right-arrow", "curvedRightArrow"),
  entry("arrows", "curved-left-arrow", "curvedLeftArrow"),
  entry("arrows", "curved-up-arrow", "curvedUpArrow"),
  entry("arrows", "curved-down-arrow", "curvedDownArrow"),
  entry("arrows", "striped-right-arrow", "stripedRightArrow"),
  entry("arrows", "notched-right-arrow", "notchedRightArrow"),
  // python-pptx calls this MSO_SHAPE.PENTAGON; its exact DrawingML identity
  // is homePlate, distinct from REGULAR_PENTAGON -> pentagon.
  entry("arrows", "home-plate", "homePlate"),
  entry("arrows", "chevron", "chevron"),
  entry("arrows", "right-arrow-callout", "rightArrowCallout"),
  entry("arrows", "down-arrow-callout", "downArrowCallout"),
  entry("arrows", "left-arrow-callout", "leftArrowCallout"),
  entry("arrows", "up-arrow-callout", "upArrowCallout"),
  entry("arrows", "left-right-arrow-callout", "leftRightArrowCallout"),
  entry("arrows", "quad-arrow-callout", "quadArrowCallout"),
  entry("arrows", "circular-arrow", "circularArrow"),
  entry("arrows", "left-circular-arrow", "leftCircularArrow"),
  entry("arrows", "left-right-circular-arrow", "leftRightCircularArrow"),
  entry("arrows", "swoosh-arrow", "swooshArrow"),
  entry("arrows", "up-down-arrow-callout", "upDownArrowCallout"),

  entry("equation", "math-plus", "mathPlus"),
  entry("equation", "math-minus", "mathMinus"),
  entry("equation", "math-multiply", "mathMultiply"),
  entry("equation", "math-divide", "mathDivide"),
  entry("equation", "math-equal", "mathEqual"),
  entry("equation", "math-not-equal", "mathNotEqual"),

  entry("flowchart", "terminator", "flowChartTerminator"),
  entry("flowchart", "document", "flowChartDocument"),
  entry("flowchart", "alternate-process", "flowChartAlternateProcess"),
  entry("flowchart", "card", "flowChartPunchedCard"),
  entry("flowchart", "collate", "flowChartCollate"),
  entry("flowchart", "flowchart-connector", "flowChartConnector"),
  entry("flowchart", "data", "flowChartInputOutput"),
  entry("flowchart", "decision", "flowChartDecision"),
  entry("flowchart", "delay", "flowChartDelay"),
  entry("flowchart", "direct-access-storage", "flowChartMagneticDrum"),
  entry("flowchart", "display", "flowChartDisplay"),
  entry("flowchart", "extract", "flowChartExtract"),
  entry("flowchart", "internal-storage", "flowChartInternalStorage"),
  entry("flowchart", "magnetic-disk", "flowChartMagneticDisk"),
  entry("flowchart", "manual-input", "flowChartManualInput"),
  entry("flowchart", "manual-operation", "flowChartManualOperation"),
  entry("flowchart", "merge", "flowChartMerge"),
  entry("flowchart", "multiple-documents", "flowChartMultidocument"),
  entry("flowchart", "offline-storage", "flowChartOfflineStorage"),
  entry("flowchart", "off-page-connector", "flowChartOffpageConnector"),
  entry("flowchart", "flowchart-or", "flowChartOr"),
  // Its kind, geometry and validation have always existed; only the catalog
  // entry was missing, which kept it out of the gallery, out of Word and Excel
  // import, and off the exact preset-program renderer (audit §2.7).
  entry("flowchart", "flowchart-summing-junction", "flowChartSummingJunction"),
  entry("flowchart", "predefined-process", "flowChartPredefinedProcess"),
  entry("flowchart", "preparation", "flowChartPreparation"),
  entry("flowchart", "process", "flowChartProcess"),
  entry("flowchart", "punched-tape", "flowChartPunchedTape"),
  entry("flowchart", "sequential-access-storage", "flowChartMagneticTape"),
  entry("flowchart", "sort", "flowChartSort"),
  entry("flowchart", "stored-data", "flowChartOnlineStorage"),

  entry("stars-banners", "star4", "star4"),
  entry("stars-banners", "star", "star5"),
  entry("stars-banners", "star6", "star6"),
  entry("stars-banners", "star7", "star7"),
  entry("stars-banners", "star8", "star8"),
  entry("stars-banners", "star10", "star10"),
  entry("stars-banners", "star12", "star12"),
  entry("stars-banners", "star16", "star16"),
  entry("stars-banners", "star24", "star24"),
  entry("stars-banners", "star32", "star32"),
  entry("stars-banners", "explosion-1", "irregularSeal1"),
  entry("stars-banners", "explosion-2", "irregularSeal2"),
  // PowerPoint's visual labels and DrawingML names are intentionally crossed:
  // Up/Down Ribbon serialize as ribbon2/ribbon; curved variants follow the
  // same ellipseRibbon2/ellipseRibbon identity. Do not normalize the suffix.
  entry("stars-banners", "up-ribbon", "ribbon2"),
  entry("stars-banners", "down-ribbon", "ribbon"),
  entry("stars-banners", "curved-up-ribbon", "ellipseRibbon2"),
  entry("stars-banners", "curved-down-ribbon", "ellipseRibbon"),
  entry("stars-banners", "left-right-ribbon", "leftRightRibbon"),
  entry("stars-banners", "vertical-scroll", "verticalScroll"),
  entry("stars-banners", "horizontal-scroll", "horizontalScroll"),
  entry("stars-banners", "wave", "wave"),
  entry("stars-banners", "double-wave", "doubleWave"),

  entry("callouts", "rectangle-callout", "wedgeRectCallout"),
  entry("callouts", "callout", "wedgeRoundRectCallout"),
  entry("callouts", "oval-callout", "wedgeEllipseCallout"),
  entry("callouts", "cloud-callout", "cloudCallout"),
  entry("callouts", "border-callout-1", "borderCallout1"),
  entry("callouts", "border-callout-2", "borderCallout2"),
  entry("callouts", "border-callout-3", "borderCallout3"),
  entry("callouts", "accent-callout-1", "accentCallout1"),
  entry("callouts", "accent-callout-2", "accentCallout2"),
  entry("callouts", "accent-callout-3", "accentCallout3"),
  entry("callouts", "callout-1", "callout1"),
  entry("callouts", "callout-2", "callout2"),
  entry("callouts", "callout-3", "callout3"),
  entry("callouts", "accent-border-callout-1", "accentBorderCallout1"),
  entry("callouts", "accent-border-callout-2", "accentBorderCallout2"),
  entry("callouts", "accent-border-callout-3", "accentBorderCallout3"),

  actionEntry("action-button-back-previous", "actionButtonBackPrevious"),
  actionEntry("action-button-forward-next", "actionButtonForwardNext"),
  actionEntry("action-button-beginning", "actionButtonBeginning"),
  actionEntry("action-button-end", "actionButtonEnd"),
  actionEntry("action-button-home", "actionButtonHome"),
  actionEntry("action-button-information", "actionButtonInformation"),
  actionEntry("action-button-return", "actionButtonReturn"),
  actionEntry("action-button-movie", "actionButtonMovie"),
  actionEntry("action-button-document", "actionButtonDocument"),
  actionEntry("action-button-sound", "actionButtonSound"),
  actionEntry("action-button-help", "actionButtonHelp"),
  actionEntry("action-button-custom", "actionButtonBlank"),
] as const satisfies readonly PresentationShapeCatalogEntry[]);

const ENTRY_BY_SELECTION = new Map(
  PRESENTATION_SHAPE_CATALOG.map((catalogEntry) => [
    catalogEntry.selection,
    catalogEntry,
  ]),
);

const ENTRY_BY_PRESET = new Map(
  PRESENTATION_SHAPE_CATALOG.filter(
    (catalogEntry): catalogEntry is PresentationPresetShapeCatalogEntry =>
      catalogEntry.geometry === "preset",
  ).map((catalogEntry) => [catalogEntry.drawingMlPreset, catalogEntry]),
);

export const PRESENTATION_SHAPE_DEFAULT_KIND: ShapeKind = "rectangle";
export const PRESENTATION_SHAPE_RECENT_LIMIT = 18;
// The stored value is only an array of stable ShapeKind IDs. Keep the original
// key while the catalog grows so a version bump does not erase real recents;
// parsePresentationShapeRecents still drops any stale/unknown ID fail-closed.
export const PRESENTATION_SHAPE_RECENT_STORAGE_KEY =
  "mrnote.presentation.shape-recents.drawingml-prstgeom-v1";

export function presentationShapeCatalogEntry(
  selection: ShapeKind,
): PresentationPresetShapeCatalogEntry | null;
export function presentationShapeCatalogEntry(
  selection: PresentationShapeCatalogSelection,
): PresentationShapeCatalogEntry | null {
  return ENTRY_BY_SELECTION.get(selection) ?? null;
}













export function presentationShapeKindTakesFill(kind: ShapeKind): boolean {
  return (
    kind !== "ink" &&
    !isOpenBasicShapeKind(kind) &&
    presentationShapeCatalogEntry(kind)?.geometry !== "connector"
  );
}










export function presentationShapeBlockTakesFill(block: {
  shape?: unknown;
  customPath?: unknown;
  openPathFilled?: unknown;
}): boolean {
  const kind = normalizeShapeKind(block.shape);
  if (!presentationShapeKindTakesFill(kind)) return false;
  if (kind !== "custom-geometry") return true;
  return normalizePresentationCustomPath(block.customPath)?.closed !== false || block.openPathFilled === true;
}










export function presentationShapeBlockTakesWordArt(block: {
  type?: unknown;
  shape?: unknown;
  text?: unknown;
}): boolean {
  if (block.type !== "shape") return true;
  const kind = normalizeShapeKind(block.shape);
  // Ink is strokes drawn with a pen; PowerPoint holds no text in it either.
  if (kind !== "ink" && !isPresentationConnectorShapeKind(kind)) return true;
  return typeof block.text === "string" && block.text.trim().length > 0;
}

/**
 * Resolve only an exactly supported preset.  Unsupported/unknown DrawingML
 * geometry returns null so an imported original package remains authoritative;
 * it is never coerced into a rectangle.
 */
export function presentationShapeCatalogEntryForPreset(
  preset: string | null | undefined,
): PresentationPresetShapeCatalogEntry | null {
  if (typeof preset !== "string") return null;
  const normalized = preset.trim();
  return normalized ? (ENTRY_BY_PRESET.get(normalized) ?? null) : null;
}

export function presentationShapeCatalogByCategory(
  category: Exclude<PresentationShapeCategoryId, "lines">,
): readonly PresentationPresetShapeCatalogEntry[];
export function presentationShapeCatalogByCategory(
  category: "lines",
): readonly PresentationShapeCatalogEntry[];
export function presentationShapeCatalogByCategory(
  category: PresentationShapeCategoryId,
): readonly PresentationShapeCatalogEntry[];
export function presentationShapeCatalogByCategory(
  category: PresentationShapeCategoryId,
): readonly PresentationShapeCatalogEntry[] {
  return PRESENTATION_SHAPE_CATALOG.filter(
    (catalogEntry) => catalogEntry.category === category,
  );
}

/**
 * Insert > Shapes' "Recently Used Shapes" before any shape was drawn: the
 * row PowerPoint for Mac 16.86 shows (audit dump ppt-menu-ins-shapes.txt),
 * minus its two text boxes, which this catalog does not hold (PPT-A112).
 */
export const PRESENTATION_SHAPE_DEFAULT_RECENTS: readonly PresentationShapeCatalogSelection[] =
  Object.freeze([
    "line",
    "arrow",
    "rectangle",
    "ellipse",
    "rounded-rectangle",
    "triangle",
    "elbow-connector",
    "elbow-arrow-connector",
    "right-arrow",
    "down-arrow",
    "freeform",
    "scribble",
    "arc",
    "curve",
    "left-brace",
    "right-brace",
  ] as PresentationShapeCatalogSelection[]);

export function parsePresentationShapeRecents(
  value: unknown,
): PresentationShapeCatalogSelection[] {
  let candidate = value;
  if (typeof value === "string") {
    try {
      candidate = JSON.parse(value);
    } catch {
      return [];
    }
  }
  if (!Array.isArray(candidate)) return [];
  const seen = new Set<PresentationShapeCatalogSelection>();
  const result: PresentationShapeCatalogSelection[] = [];
  for (const rawSelection of candidate) {
    if (typeof rawSelection !== "string") continue;
    const catalogEntry = ENTRY_BY_SELECTION.get(
      rawSelection as PresentationShapeCatalogSelection,
    );
    if (!catalogEntry || seen.has(catalogEntry.selection)) continue;
    seen.add(catalogEntry.selection);
    result.push(catalogEntry.selection);
    if (result.length === PRESENTATION_SHAPE_RECENT_LIMIT) break;
  }
  return result;
}

export function pushPresentationShapeRecent(
  current: readonly PresentationShapeCatalogSelection[],
  selection: PresentationShapeCatalogSelection,
): PresentationShapeCatalogSelection[] {
  if (!ENTRY_BY_SELECTION.has(selection)) {
    return parsePresentationShapeRecents(current);
  }
  return [
    selection,
    ...current.filter((candidate) => candidate !== selection),
  ].slice(0, PRESENTATION_SHAPE_RECENT_LIMIT);
}

export function presentationShapeMatchesSearch(
  catalogEntry: PresentationShapeCatalogEntry,
  query: string,
  localizedLabel: string,
): boolean {
  const needle = query.trim().toLocaleLowerCase();
  if (!needle) return true;
  return [
    localizedLabel,
    catalogEntry.selection,
    "drawingMlPreset" in catalogEntry ? catalogEntry.drawingMlPreset : "",
    catalogEntry.category,
  ].some((value) => value.toLocaleLowerCase().includes(needle));
}
