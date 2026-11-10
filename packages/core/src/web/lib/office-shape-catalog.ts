/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shared Office-for-the-web preset-shape catalog.
 *
 * Word and Excel keep different object/anchor schemas, but DrawingML preset
 * geometry is the same vocabulary in DOCX, XLSX and PPTX. This module is the
 * one browser-side registry for that vocabulary; editor adapters keep their
 * own persistence and command authority.
 */

import {
  isPresentationPresetShapeCatalogEntry,
  PRESENTATION_SHAPE_CATALOG,
  type PresentationPresetShapeCatalogEntry,
  type PresentationShapeCategoryId,
} from "./presentation-shape-catalog";
import {
  normalizePresentationConnectorSpec,
  PRESENTATION_CONNECTOR_END_SIZES,
  PRESENTATION_CONNECTOR_ENDS,
  PRESENTATION_FREE_END_CONNECTOR_KINDS,
  type PresentationConnectorEnd,
  type ShapeKind,
} from "./presentation-shapes";

const LEGACY_OFFICE_SHAPE_DEFINITIONS = [
  ["line", "line", "line"],
  ["lineArrow", "line", "arrow"],
  ["rect", "rect", "rectangle"],
  ["roundRect", "roundRect", "rounded-rectangle"],
  ["ellipse", "ellipse", "ellipse"],
  ["triangle", "triangle", "triangle"],
  ["rtTriangle", "rtTriangle", "right-triangle"],
  ["diamond", "diamond", "diamond"],
  ["parallelogram", "parallelogram", "parallelogram"],
  ["trapezoid", "trapezoid", "trapezoid"],
  ["pentagon", "pentagon", "pentagon"],
  ["hexagon", "hexagon", "hexagon"],
  ["plus", "plus", "plus"],
  ["heart", "heart", "heart"],
  ["cloud", "cloud", "cloud"],
  ["sun", "sun", "sun"],
  ["moon", "moon", "moon"],
  ["pie", "pie", "pie"],
  ["lightningBolt", "lightningBolt", "lightning-bolt"],
  ["rightArrow", "rightArrow", "right-arrow"],
  ["leftArrow", "leftArrow", "left-arrow"],
  ["upArrow", "upArrow", "up-arrow"],
  ["downArrow", "downArrow", "down-arrow"],
  ["leftRightArrow", "leftRightArrow", "left-right-arrow"],
  ["upDownArrow", "upDownArrow", "up-down-arrow"],
  ["chevron", "chevron", "chevron"],
  ["bentArrow", "bentArrow", "bent-arrow"],
  ["uturnArrow", "uturnArrow", "u-turn-arrow"],
  ["flowChartTerminator", "flowChartTerminator", "terminator"],
  ["flowChartDocument", "flowChartDocument", "document"],
  ["star5", "star5", "star"],
  ["star4", "star4", "star4"],
  ["star6", "star6", "star6"],
  ["wedgeRectCallout", "wedgeRectCallout", "rectangle-callout"],
  ["wedgeEllipseCallout", "wedgeEllipseCallout", "oval-callout"],
  ["cloudCallout", "cloudCallout", "cloud-callout"],
] as const satisfies readonly (readonly [string, string, ShapeKind])[];

type LegacyOfficeShapePreset =
  (typeof LEGACY_OFFICE_SHAPE_DEFINITIONS)[number][0];

/** Stable persisted selection accepted by every Office editor adapter. */
export type OfficeShapePreset = LegacyOfficeShapePreset | ShapeKind;

export const OFFICE_SHAPE_CATEGORY_IDS = [
  "lines",
  "rectangles",
  "basic",
  "arrows",
  "equation",
  "flowchart",
  "stars-banners",
  "callouts",
] as const satisfies readonly PresentationShapeCategoryId[];

export type OfficeShapeCategoryId =
  (typeof OFFICE_SHAPE_CATEGORY_IDS)[number];

export interface OfficeShapeDefinition {
  readonly preset: OfficeShapePreset;
  /** DrawingML a:prstGeom/@prst. */
  readonly drawingPreset: string;
  /** Shared SVG geometry identity. */
  readonly previewKind: ShapeKind;
  /**
   * The gallery group, or `action-buttons` for a geometry Word and Excel can
   * READ but offer no picker for. {@link OFFICE_SHAPE_CATEGORIES} lists only
   * the gallery groups, so an action button never reaches a shape gallery
   * while still resolving on import.
   */
  readonly category: PresentationShapeCategoryId;
  readonly lineLike?: true;
  readonly lineStart?: "triangle";
  readonly lineEnd?: "triangle";
}

const legacyPresetByCatalogIdentity = new Map(
  LEGACY_OFFICE_SHAPE_DEFINITIONS.map(
    ([preset, drawingPreset, previewKind]) => [
      `${drawingPreset}:${previewKind}`,
      preset,
    ],
  ),
);

/**
 * Every DrawingML geometry a Word or Excel document may CONTAIN.
 *
 * The twelve action buttons used to be filtered out here, which is why S182-S193
 * of the Excel corpus were dropped on import (SHAPE-IMPORT-AUDIT.md §2.8). They
 * are ordinary preset geometries in the format — `a:prstGeom` is the same
 * element in WordprocessingML, SpreadsheetML and PresentationML — and Office's
 * own APIs create them in both hosts (`Word.GeometricShapeType` and
 * `Excel.GeometricShapeType` list all twelve; so does the shared
 * `MsoAutoShapeType` that Word's and Excel's `Shapes.AddShape` takes). Only the
 * "Action Buttons" GALLERY GROUP and the click/mouse-over action semantics are
 * PowerPoint-only, and those live in {@link OFFICE_SHAPE_CATEGORIES}, which
 * still lists the eight gallery groups and not this one.
 */
const officeShapeCatalogEntries = PRESENTATION_SHAPE_CATALOG.filter(
  isPresentationPresetShapeCatalogEntry,
);

export const OFFICE_SHAPE_DEFINITIONS: readonly OfficeShapeDefinition[] =
  Object.freeze(
    officeShapeCatalogEntries.map((entry) => {
      const connector = normalizePresentationConnectorSpec({
        shape: entry.kind,
      });
      const legacyPreset = legacyPresetByCatalogIdentity.get(
        `${entry.drawingMlPreset}:${entry.kind}`,
      );
      return Object.freeze({
        preset: legacyPreset ?? entry.kind,
        drawingPreset: entry.drawingMlPreset,
        previewKind: entry.kind,
        category: entry.category,
        ...(entry.geometry === "connector" ? { lineLike: true as const } : {}),
        ...(connector?.headEnd === "triangle"
          ? { lineStart: "triangle" as const }
          : {}),
        ...(connector?.tailEnd === "triangle"
          ? { lineEnd: "triangle" as const }
          : {}),
      });
    }),
  );

export const OFFICE_SHAPE_PRESETS: readonly OfficeShapePreset[] = Object.freeze(
  OFFICE_SHAPE_DEFINITIONS.map((definition) => definition.preset),
);

const definitionByPreset = new Map(
  OFFICE_SHAPE_DEFINITIONS.map((definition) => [
    definition.preset,
    definition,
  ]),
);

export interface OfficeShapeCategory {
  readonly id: OfficeShapeCategoryId;
  readonly label: Readonly<{ en: string; zh: string }>;
  readonly presets: readonly OfficeShapePreset[];
}

const OFFICE_SHAPE_CATEGORY_LABELS: Readonly<
  Record<OfficeShapeCategoryId, Readonly<{ en: string; zh: string }>>
> = Object.freeze({
  lines: Object.freeze({ en: "Lines", zh: "线条" }),
  rectangles: Object.freeze({ en: "Rectangles", zh: "矩形" }),
  basic: Object.freeze({ en: "Basic Shapes", zh: "基本形状" }),
  arrows: Object.freeze({ en: "Block Arrows", zh: "箭头总汇" }),
  equation: Object.freeze({ en: "Equation Shapes", zh: "公式形状" }),
  flowchart: Object.freeze({ en: "Flowchart", zh: "流程图" }),
  "stars-banners": Object.freeze({
    en: "Stars and Banners",
    zh: "星与旗帜",
  }),
  callouts: Object.freeze({ en: "Callouts", zh: "标注" }),
});

export const OFFICE_SHAPE_CATEGORIES: readonly OfficeShapeCategory[] =
  Object.freeze(
    OFFICE_SHAPE_CATEGORY_IDS.map((id) =>
      Object.freeze({
        id,
        label: OFFICE_SHAPE_CATEGORY_LABELS[id],
        presets: Object.freeze(
          OFFICE_SHAPE_DEFINITIONS.filter(
            (definition) => definition.category === id,
          ).map((definition) => definition.preset),
        ),
      }),
    ),
  );

const presetSet = new Set<string>(OFFICE_SHAPE_PRESETS);

/**
 * The presets a gallery can actually insert. Recents are a memory of what the
 * user picked, so they are drawn from here and not from the whole registry:
 * an action button is readable in a Word or Excel document but has no picker
 * in either, and a recents chip for a shape nothing can insert is dead UI.
 */
const galleryPresetSet = new Set<string>(
  OFFICE_SHAPE_CATEGORIES.flatMap((category) => category.presets),
);

export function isOfficeShapePreset(
  value: unknown,
): value is OfficeShapePreset {
  return typeof value === "string" && presetSet.has(value);
}

function isOfficeShapeGalleryPreset(
  value: unknown,
): value is OfficeShapePreset {
  return typeof value === "string" && galleryPresetSet.has(value);
}

export function officeShapeDefinition(
  preset: OfficeShapePreset,
): OfficeShapeDefinition {
  return definitionByPreset.get(preset) ?? definitionByPreset.get("rect")!;
}

/**
 * The gallery preset for one `a:prstGeom/@prst` plus its two `a:ln` ends.
 *
 * A gallery entry names the route and WHETHER each end carries a decoration;
 * WHICH decoration it is lives on `a:ln` and is none of `@prst`'s business, so
 * all six `a:ST_LineEndType` members resolve through the same entries
 * (SHAPE-IMPORT-AUDIT.md §7.4). The ladder mirrors the server's
 * `office_connector_identity` exactly: presence match, then the preset's own
 * free-end entry, then the generic arrow carrier a head-only straight
 * connector has always used. A NON-connector geometry wearing an arrowhead
 * still resolves to nothing — its outline belongs to the source package.
 *
 * `keepRoute` is for a host whose preset is ALL it knows of the route. A slide
 * block keeps `route`/`connectorPreset` beside its kind, so the generic arrow
 * there is still an elbow; a document shape and a sheet shape keep nothing
 * else, and the same answer turned a head-only `bentConnector3` into a
 * straight line. Such a host gets the preset's own undecorated entry and
 * states the ends itself. Mirrors `office_connector_identity(keep_route=True)`
 * (SHAPE-IMPORT-AUDIT.md §7.11).
 */
export function officeShapePresetFromDrawingMl(
  drawingPreset: unknown,
  lineEnd: unknown = null,
  lineStart: unknown = null,
  options: { readonly keepRoute?: boolean } = {},
): OfficeShapePreset | null {
  const normalizeLineEnd = (value: unknown): PresentationConnectorEnd | null =>
    value === null || value === undefined
      ? "none"
      : typeof value === "string" &&
          (PRESENTATION_CONNECTOR_ENDS as readonly string[]).includes(value)
        ? (value as PresentationConnectorEnd)
        : null;
  const normalizedLineEnd = normalizeLineEnd(lineEnd);
  const normalizedLineStart = normalizeLineEnd(lineStart);
  if (normalizedLineEnd === null || normalizedLineStart === null) return null;
  const candidates = OFFICE_SHAPE_DEFINITIONS.filter(
    (definition) => definition.drawingPreset === drawingPreset,
  );
  const byPresence = candidates.find(
    (definition) =>
      (definition.lineStart !== undefined) === (normalizedLineStart !== "none") &&
      (definition.lineEnd !== undefined) === (normalizedLineEnd !== "none"),
  );
  if (byPresence) return byPresence.preset;
  const connectors = candidates.filter((definition) => definition.lineLike);
  if (connectors.length === 0) return null;
  const freeEnd = connectors.find((definition) =>
    PRESENTATION_FREE_END_CONNECTOR_KINDS.has(definition.previewKind),
  );
  if (freeEnd) return freeEnd.preset;
  const genericArrow = OFFICE_SHAPE_DEFINITIONS.find(
    (definition) => definition.previewKind === "arrow",
  );
  if (options.keepRoute && genericArrow?.drawingPreset !== drawingPreset) {
    return (
      connectors.find(
        (definition) =>
          definition.lineStart === undefined && definition.lineEnd === undefined,
      )?.preset ?? null
    );
  }
  return genericArrow?.preset ?? null;
}

/** The page-level spelling of one `a:ln` end's four optional attributes. */
export interface OfficeShapeLineEndFields {
  headEnd?: PresentationConnectorEnd;
  tailEnd?: PresentationConnectorEnd;
  headEndWidth?: string;
  headEndLength?: string;
  tailEndWidth?: string;
  tailEndLength?: string;
}

/**
 * The end fields a shape must carry beyond what its gallery preset implies.
 *
 * A connector preset already names WHICH ends are decorated, so a value it
 * implies stays absent — every document that only ever wore `none`/`triangle`
 * therefore projects exactly the attrs it projected before. A size for an
 * undecorated end is dropped: `@w`/`@len` size a decoration that is not there.
 * Mirrors `_document_shape_end_fields` on the server.
 */
export function officeShapeLineEndFields(
  preset: OfficeShapePreset | null,
  headEnd: PresentationConnectorEnd,
  tailEnd: PresentationConnectorEnd,
  sizes: Readonly<Record<string, string | null | undefined>> = {},
): OfficeShapeLineEndFields {
  const fields: OfficeShapeLineEndFields = {};
  for (const key of [
    "headEndWidth",
    "headEndLength",
    "tailEndWidth",
    "tailEndLength",
  ] as const) {
    const size = sizes[key];
    if (typeof size !== "string") continue;
    if (!(PRESENTATION_CONNECTOR_END_SIZES as readonly string[]).includes(size)) {
      continue;
    }
    if ((key.startsWith("head") ? headEnd : tailEnd) === "none") continue;
    fields[key] = size;
  }
  const definition = preset ? definitionByPreset.get(preset) : undefined;
  const spec = definition
    ? normalizePresentationConnectorSpec({ shape: definition.previewKind })
    : null;
  if (headEnd !== (spec?.headEnd ?? "none")) fields.headEnd = headEnd;
  if (tailEnd !== (spec?.tailEnd ?? "none")) fields.tailEnd = tailEnd;
  return fields;
}

/**
 * The semantic connector KIND for one `a:prstGeom/@prst` plus its two ends —
 * the same ladder {@link officeShapePresetFromDrawingMl} climbs, in the kind
 * vocabulary the slide canvas speaks. One table for Word, Excel and PowerPoint;
 * the three used to keep a hand-written copy each (SHAPE-IMPORT-AUDIT.md §2.4).
 */
export function officeConnectorKindFromDrawingMl(
  drawingPreset: unknown,
  headEnd: PresentationConnectorEnd,
  tailEnd: PresentationConnectorEnd,
): ShapeKind | null {
  const preset = officeShapePresetFromDrawingMl(drawingPreset, tailEnd, headEnd);
  const definition = preset ? definitionByPreset.get(preset) : undefined;
  return definition?.lineLike ? definition.previewKind : null;
}

export const OFFICE_SHAPE_RECENT_LIMIT = 18;

export function parseOfficeShapeRecents(value: unknown): OfficeShapePreset[] {
  let candidate = value;
  if (typeof value === "string") {
    try {
      candidate = JSON.parse(value);
    } catch {
      return [];
    }
  }
  if (!Array.isArray(candidate)) return [];
  const seen = new Set<OfficeShapePreset>();
  const result: OfficeShapePreset[] = [];
  for (const rawPreset of candidate) {
    if (!isOfficeShapeGalleryPreset(rawPreset) || seen.has(rawPreset)) continue;
    seen.add(rawPreset);
    result.push(rawPreset);
    if (result.length === OFFICE_SHAPE_RECENT_LIMIT) break;
  }
  return result;
}

export function pushOfficeShapeRecent(
  current: readonly OfficeShapePreset[],
  preset: OfficeShapePreset,
): OfficeShapePreset[] {
  if (!isOfficeShapeGalleryPreset(preset)) {
    return parseOfficeShapeRecents(current);
  }
  return [preset, ...current.filter((candidate) => candidate !== preset)].slice(
    0,
    OFFICE_SHAPE_RECENT_LIMIT,
  );
}

export function officeShapeMatchesSearch(
  preset: OfficeShapePreset,
  query: string,
  localizedLabel: string,
): boolean {
  const normalizeSearchText = (value: string) =>
    value
      .normalize("NFKD")
      .toLocaleLowerCase()
      .replace(/[\s_-]+/g, "");
  const needle = normalizeSearchText(query.trim());
  if (!needle) return true;
  const definition = officeShapeDefinition(preset);
  return [
    localizedLabel,
    preset,
    definition.previewKind,
    definition.drawingPreset,
    definition.category,
  ].some((value) => normalizeSearchText(value).includes(needle));
}

/**
 * How far OUTSIDE the shape's own box a preset preview draws, as a fraction of
 * the rendered element on each side.
 *
 * The preview renders into a view box larger than the shape so a callout's
 * leader and a stroke's outer half are not clipped. The consequence is that
 * the shape's outline lands INSET inside the element, and anything a host
 * overlays on the shape -- its text, above all -- must use the same inset or
 * it is drawn against the element's edge while the outline sits somewhere
 * inside it. Measured in a real browser on the sheet's own QA route: element
 * 561x274, the drawn outline 501x245 at +30,+15, and the text box's words
 * accordingly above and to the LEFT of its own border
 * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §41).
 *
 * These are the numbers `PresentationPresetShapePreview` builds its view box
 * from; they live here so a host can ask for them without pulling the preview
 * component in.
 */
export const OFFICE_SHAPE_PREVIEW_VIEW_BOX = Object.freeze({
  callout: Object.freeze({ left: -0.52, top: -0.12, width: 1.64, height: 1.52 }),
  default: Object.freeze({ left: -0.06, top: -0.06, width: 1.12, height: 1.12 }),
});

export interface OfficeShapePreviewInset {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/**
 * Where the shared preset preview must be SIZED, as fractions of the box a
 * shape is drawn in, so that its 0..1 geometry fills that box exactly: the
 * preview's view box bleeds past the geometry (a leader's overshoot), so the
 * `<svg>` is placed that much outside the box. A drawing that stands in for
 * Excel's shape draws it this way -- Excel paints the geometry edge to edge
 * of the anchored box -- while a gallery tile keeps the bleed as air.
 */
export function officeShapePreviewBleedFrame(kind: string): {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
} {
  const box = kind.includes("callout")
    ? OFFICE_SHAPE_PREVIEW_VIEW_BOX.callout
    : OFFICE_SHAPE_PREVIEW_VIEW_BOX.default;
  return { left: box.left, top: box.top, width: box.width, height: box.height };
}

export function officeShapePreviewInset(kind: string): OfficeShapePreviewInset {
  const box = kind.includes("callout")
    ? OFFICE_SHAPE_PREVIEW_VIEW_BOX.callout
    : OFFICE_SHAPE_PREVIEW_VIEW_BOX.default;
  return {
    left: -box.left / box.width,
    top: -box.top / box.height,
    right: (box.width + box.left - 1) / box.width,
    bottom: (box.height + box.top - 1) / box.height,
  };
}
