/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation text-box and shape blocks (PowerPoint / WPS parity subset).
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor,
 * and the present-mode overlay so the allowed shape set, default geometry, and
 * fill/stroke resolution live in exactly one place. Keep this module
 * dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * imported from the zod schema.
 *
 * Geometry is expressed in EMUs (English Metric Units), matching the existing
 * presentation geometry (`left`/`top`/`width`/`height` on every block).
 */

import type { CSSProperties } from "react";

import {
  officeFillPatternCss,
  officeFillPatternOf,
  type OfficeChartFillPattern,
} from "./office-chart-shape-paint";

import { isDrawableImageSource, isValidInlineImageSource } from "./inline-image-source";
import {
  officePictureFillOptionsOf,
  officePictureFillOptionsStated,
  type OfficePictureFillOptions,
} from "./office-picture-fill-options";
import {
  PRESENTATION_INK_MAX_POINTS,
  normalizePresentationCustomPath,
  presentationCustomPathToSvgPath,
  type PresentationCustomPath,
} from "./presentation-freeform";
import {
  parsePolygonString,
  type Point as PolygonPoint,
} from "./presentation-polygon-points";

/** A 4 : 3-ish default slide canvas, mirroring PresentationWindow defaults. */
const DEFAULT_SLIDE_WIDTH_EMU = 9_144_000;
const DEFAULT_SLIDE_HEIGHT_EMU = 5_143_500;

/** The supported shape kinds. `rectangle` is the implicit default. */
export const SHAPE_KINDS = [
  "rectangle",
  "ellipse",
  "line",
  "arrow",
  // DrawingML connector lines. Each stable kind captures the route and the
  // independently authored line ends instead of folding every p:cxnSp onto
  // the historical generic line/arrow pair.
  "double-arrow",
  "elbow-connector",
  "elbow-arrow-connector",
  "elbow-double-arrow-connector",
  "curved-connector",
  "curved-arrow-connector",
  "curved-double-arrow-connector",
  // The rest of DrawingML's connector family. Word, Excel and PowerPoint write
  // these themselves whenever a connector needs a different number of bends
  // than the three-segment default, so a document can contain them even though
  // no gallery item inserts one. Without a kind of their own they were the last
  // eight `prst` values no host could read at all
  // (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §2.6/§7).
  "elbow-connector-2",
  "elbow-connector-4",
  "elbow-connector-5",
  "curved-connector-2",
  "curved-connector-4",
  "curved-connector-5",
  "straight-connector",
  // `lineInv` is the other diagonal of `line`. Office models it as an AUTOSHAPE,
  // not a connector: MsoAutoShapeType carries msoShapeLineInverse and
  // Word/Excel's GeometricShapeType carries lineInverse, while no connector
  // enum does; its preset geometry is fillable and declares connection sites of
  // its own, which is what a shape a connector attaches TO looks like.
  "line-inverse",
  "rounded-rectangle",
  // DrawingML's complete nine-item Rectangles family. These seven keep their
  // distinct preset identity instead of collapsing into rect / roundRect.
  "single-round-corner",
  "same-side-corners-rounded",
  "diagonal-corners-rounded",
  "single-corner-snipped",
  "same-side-corners-snipped",
  "diagonal-corners-snipped",
  "single-corner-rounded-snipped",
  // DrawingML's complete six-item Equation Shapes family. These are ordinary
  // preset geometries (not equation objects or action buttons), so their
  // stable identity is the exact math* prstGeom name carried by the catalog.
  "math-plus",
  "math-minus",
  "math-multiply",
  "math-divide",
  "math-equal",
  "math-not-equal",
  "triangle",
  "diamond",
  "pentagon",
  "hexagon",
  "star",
  "star4",
  "star6",
  // DrawingML's complete Stars and Banners family. The existing star/star4/
  // star6 identifiers stay stable; every additional kind names one exact
  // prstGeom identity in the versioned authoring catalog.
  "star7",
  "star8",
  "star10",
  "star12",
  "star16",
  "star24",
  "star32",
  "explosion-1",
  "explosion-2",
  "up-ribbon",
  "down-ribbon",
  "curved-up-ribbon",
  "curved-down-ribbon",
  // The fifth ribbon, same provenance as the Basic Shapes below.
  "left-right-ribbon",
  "vertical-scroll",
  "horizontal-scroll",
  "wave",
  "double-wave",
  "chevron",
  "callout",
  
  // highest-frequency PowerPoint gallery shapes our straight-edge polygon
  // renderer can carry natively.
  "parallelogram",
  "trapezoid",
  "right-triangle",
  "octagon",
  "plus",
  
  // than a polygon — shapePathD supplies the outline.
  "cloud",
  "heart",
  "can",
  "moon",
  
  "heptagon",
  "decagon",
  "dodecagon",
  "corner",
  "chord",
  "teardrop",
  "sun",
  "pie",
  "lightning-bolt",
  // DrawingML Basic Shapes catalog completion. The ten silhouettes below
  // retain ordinary fill semantics; the seven open arc/bracket/brace presets
  // are separately enumerated in OPEN_BASIC_SHAPE_KINDS so no renderer can
  // repaint them through the generic filled-polygon fallback.
  "non-isosceles-trapezoid",
  "pie-wedge",
  "half-frame",
  "diagonal-stripe",
  "plaque",
  "cube",
  "bevel",
  "no-symbol",
  "block-arc",
  "folded-corner",
  "arc",
  "double-bracket",
  "double-brace",
  "left-bracket",
  "right-bracket",
  "left-brace",
  "right-brace",
  // Hollow kinds (fill-rule evenodd: outer outline minus inner cutout).
  "frame",
  "donut",
  // Basic Shapes the shared catalog had never carried at all (SHAPE-IMPORT-AUDIT
  // §2.6). Their outlines are the same Apache POI default preset programs every
  // other preset already ships, transcribed into this 0..100 space so the Docs
  // canvas — which never consults those programs — can paint them too.
  "gear-6",
  "gear-9",
  "funnel",
  "chart-plus",
  "chart-star",
  "chart-x",
  "corner-tabs",
  "square-tabs",
  "plaque-tabs",
  // Multi-part kinds: several SVG parts with per-part paint roles.
  "smiley",
  // Block-arrow family (filled directional arrows — NOT line connectors).
  "right-arrow",
  "left-arrow",
  "up-arrow",
  "down-arrow",
  "left-right-arrow",
  "up-down-arrow",
  "bent-arrow",
  "u-turn-arrow",
  // DrawingML Block Arrows completion. Each kind below owns one exact
  // prstGeom identity; curved/circular members use real banded SVG paths and
  // callouts keep a broad text-bearing body instead of a rectangle fallback.
  "quad-arrow",
  "left-right-up-arrow",
  "left-up-arrow",
  "bent-up-arrow",
  "curved-right-arrow",
  "curved-left-arrow",
  "curved-up-arrow",
  "curved-down-arrow",
  "striped-right-arrow",
  "notched-right-arrow",
  "home-plate",
  "right-arrow-callout",
  "down-arrow-callout",
  "left-arrow-callout",
  "up-arrow-callout",
  "left-right-arrow-callout",
  "quad-arrow-callout",
  "circular-arrow",
  // The remaining Block Arrows, same provenance as the Basic Shapes above.
  "left-circular-arrow",
  "left-right-circular-arrow",
  "swoosh-arrow",
  "up-down-arrow-callout",
  // DrawingML's complete 28-item Flowchart family. Every kind below keeps
  // its exact prstGeom identity; storage, document and internal-line members
  // use authored paths/multipart geometry instead of a rectangle fallback.
  "terminator",
  "document",
  "alternate-process",
  "card",
  "collate",
  "flowchart-connector",
  "data",
  "decision",
  "delay",
  "direct-access-storage",
  "display",
  "extract",
  "internal-storage",
  "magnetic-disk",
  "manual-input",
  "manual-operation",
  "merge",
  "multiple-documents",
  "offline-storage",
  "off-page-connector",
  "flowchart-or",
  "flowchart-summing-junction",
  "predefined-process",
  "preparation",
  "process",
  "punched-tape",
  "sequential-access-storage",
  "sort",
  "stored-data",
  // DrawingML's complete Callouts family. The four wedge/cloud callouts are
  // filled silhouettes; the twelve line-callout presets below are open,
  // outline-only text frames/leaders rendered by shapeMultiParts.
  "rectangle-callout",
  "oval-callout",
  "cloud-callout",
  "border-callout-1",
  "border-callout-2",
  "border-callout-3",
  "accent-callout-1",
  "accent-callout-2",
  "accent-callout-3",
  "callout-1",
  "callout-2",
  "callout-3",
  "accent-border-callout-1",
  "accent-border-callout-2",
  "accent-border-callout-3",
  // Provider-free Action Buttons. Their exact DrawingML geometry and default
  // click action are coupled in the versioned shape catalog.
  "action-button-back-previous",
  "action-button-forward-next",
  "action-button-beginning",
  "action-button-end",
  "action-button-home",
  "action-button-information",
  "action-button-return",
  "action-button-movie",
  "action-button-document",
  "action-button-sound",
  "action-button-help",
  "action-button-custom",
  // Durable bounded DrawingML `custGeom`. Curve / Freeform / Scribble remain
  // transient tools and never become persisted shape identities.
  "custom-geometry",
  // Freehand INK (PPT Draw tab): an open polyline stroke. Points live on
  // the block (`points`, 0-100 space) like a merged shape custom outline;
  // the renderer draws a polyline (no fill, no closing edge), not a polygon.
  "ink",
] as const;

export type ShapeKind = (typeof SHAPE_KINDS)[number];

/** Native connector route, preset and endpoint vocabularies. */
export const PRESENTATION_CONNECTOR_ROUTES = [
  "straight",
  "elbow",
  "curved",
] as const;
export type PresentationConnectorRoute =
  (typeof PRESENTATION_CONNECTOR_ROUTES)[number];

export const PRESENTATION_CONNECTOR_PRESETS = [
  "line",
  "straightConnector1",
  "bentConnector2",
  "bentConnector3",
  "bentConnector4",
  "bentConnector5",
  "curvedConnector2",
  "curvedConnector3",
  "curvedConnector4",
  "curvedConnector5",
] as const;
export type PresentationConnectorPreset =
  (typeof PRESENTATION_CONNECTOR_PRESETS)[number];

/**
 * Every member of `a:ST_LineEndType` (ECMA-376 §20.1.10.33), in schema order.
 * PowerPoint's own Format Shape ▸ Line ▸ "Begin/End Arrow type" menu offers all
 * six, and so do Word's and Excel's. This used to stop at `triangle`, so a
 * connector wearing any of the other four was refused outright by all three
 * hosts (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §7.4).
 */
export const PRESENTATION_CONNECTOR_ENDS = [
  "none",
  "triangle",
  "stealth",
  "diamond",
  "oval",
  "arrow",
] as const;
export type PresentationConnectorEnd =
  (typeof PRESENTATION_CONNECTOR_ENDS)[number];

/**
 * `a:ST_LineEndWidth` / `a:ST_LineEndLength` (§20.1.10.34 / §20.1.10.32) — the
 * arrowhead's size relative to the line weight, PowerPoint's "Begin/End Arrow
 * size". Both attributes are OPTIONAL on `CT_LineEndProperties` and the schema
 * declares NO default for either, nor for `type`. Real packages prove it: Word
 * writes `<a:tailEnd type="triangle"/>` bare while a slide writes
 * `<a:tailEnd type="triangle" w="med" len="med"/>`. An absent size therefore
 * stays absent here and on the way out — a written-out `med` would change the
 * bytes of a shape nobody touched.
 */
export const PRESENTATION_CONNECTOR_END_SIZES = ["sm", "med", "lg"] as const;
export type PresentationConnectorEndSize =
  (typeof PRESENTATION_CONNECTOR_END_SIZES)[number];

export interface PresentationConnectorSpec {
  readonly route: PresentationConnectorRoute;
  readonly connectorPreset: PresentationConnectorPreset;
  readonly headEnd: PresentationConnectorEnd;
  readonly tailEnd: PresentationConnectorEnd;
  readonly headEndWidth?: PresentationConnectorEndSize;
  readonly headEndLength?: PresentationConnectorEndSize;
  readonly tailEndWidth?: PresentationConnectorEndSize;
  readonly tailEndLength?: PresentationConnectorEndSize;
}

const PRESENTATION_CONNECTOR_KIND_DEFAULTS: Readonly<
  Partial<Record<ShapeKind, PresentationConnectorSpec>>
> = Object.freeze({
  line: {
    route: "straight",
    connectorPreset: "line",
    headEnd: "none",
    tailEnd: "none",
  },
  arrow: {
    route: "straight",
    connectorPreset: "line",
    headEnd: "none",
    tailEnd: "triangle",
  },
  "double-arrow": {
    route: "straight",
    connectorPreset: "line",
    headEnd: "triangle",
    tailEnd: "triangle",
  },
  "elbow-connector": {
    route: "elbow",
    connectorPreset: "bentConnector3",
    headEnd: "none",
    tailEnd: "none",
  },
  "elbow-arrow-connector": {
    route: "elbow",
    connectorPreset: "bentConnector3",
    headEnd: "none",
    tailEnd: "triangle",
  },
  "elbow-double-arrow-connector": {
    route: "elbow",
    connectorPreset: "bentConnector3",
    headEnd: "triangle",
    tailEnd: "triangle",
  },
  "curved-connector": {
    route: "curved",
    connectorPreset: "curvedConnector3",
    headEnd: "none",
    tailEnd: "none",
  },
  "curved-arrow-connector": {
    route: "curved",
    connectorPreset: "curvedConnector3",
    headEnd: "none",
    tailEnd: "triangle",
  },
  "curved-double-arrow-connector": {
    route: "curved",
    connectorPreset: "curvedConnector3",
    headEnd: "triangle",
    tailEnd: "triangle",
  },
  // One kind per remaining DrawingML connector preset. Their ends stay
  // independent (`a:ln/a:headEnd|a:tailEnd` is separate from `@prst` in the
  // format), so unlike the three-segment trio above these do not need an
  // arrow/double-arrow kind each.
  "elbow-connector-2": {
    route: "elbow",
    connectorPreset: "bentConnector2",
    headEnd: "none",
    tailEnd: "none",
  },
  "elbow-connector-4": {
    route: "elbow",
    connectorPreset: "bentConnector4",
    headEnd: "none",
    tailEnd: "none",
  },
  "elbow-connector-5": {
    route: "elbow",
    connectorPreset: "bentConnector5",
    headEnd: "none",
    tailEnd: "none",
  },
  "curved-connector-2": {
    route: "curved",
    connectorPreset: "curvedConnector2",
    headEnd: "none",
    tailEnd: "none",
  },
  "curved-connector-4": {
    route: "curved",
    connectorPreset: "curvedConnector4",
    headEnd: "none",
    tailEnd: "none",
  },
  "curved-connector-5": {
    route: "curved",
    connectorPreset: "curvedConnector5",
    headEnd: "none",
    tailEnd: "none",
  },
  "straight-connector": {
    route: "straight",
    connectorPreset: "straightConnector1",
    headEnd: "none",
    tailEnd: "none",
  },
});

/**
 * The kinds whose ends are free, because their DrawingML preset says nothing
 * about arrowheads. The historical nine fold route AND ends into the kind, so
 * each of those stays an exact tuple.
 */
export const PRESENTATION_FREE_END_CONNECTOR_KINDS: ReadonlySet<ShapeKind> = new Set([
  "elbow-connector-2",
  "elbow-connector-4",
  "elbow-connector-5",
  "curved-connector-2",
  "curved-connector-4",
  "curved-connector-5",
  "straight-connector",
]);

const PRESENTATION_CONNECTOR_KIND_SET: ReadonlySet<string> = new Set(
  Object.keys(PRESENTATION_CONNECTOR_KIND_DEFAULTS),
);
const PRESENTATION_CONNECTOR_ROUTE_SET: ReadonlySet<string> = new Set(
  PRESENTATION_CONNECTOR_ROUTES,
);
const PRESENTATION_CONNECTOR_PRESET_SET: ReadonlySet<string> = new Set(
  PRESENTATION_CONNECTOR_PRESETS,
);
const PRESENTATION_CONNECTOR_END_SET: ReadonlySet<string> = new Set(
  PRESENTATION_CONNECTOR_ENDS,
);
const PRESENTATION_CONNECTOR_END_SIZE_SET: ReadonlySet<string> = new Set(
  PRESENTATION_CONNECTOR_END_SIZES,
);
const PRESENTATION_CONNECTOR_END_SIZE_FIELDS = [
  "headEndWidth",
  "headEndLength",
  "tailEndWidth",
  "tailEndLength",
] as const;

export function isPresentationConnectorShapeKind(
  // A plain string, because the callers that ask are holding a preset name
  // from the catalog (`lineArrow`) or a kind (`line`) and both have to be
  // answerable without a cast at every call site.
  kind: string,
): boolean {
  return PRESENTATION_CONNECTOR_KIND_SET.has(kind);
}

export function presentationConnectorDefaults(
  kind: ShapeKind,
): PresentationConnectorSpec | null {
  return PRESENTATION_CONNECTOR_KIND_DEFAULTS[kind] ?? null;
}

function routeForConnectorPreset(
  preset: PresentationConnectorPreset,
): PresentationConnectorRoute {
  if (preset.startsWith("bentConnector")) return "elbow";
  if (preset.startsWith("curvedConnector")) return "curved";
  return "straight";
}

/**
 * Shape Outline > Arrows on an OPEN shape: Office puts arrowheads on the two
 * free ends of an arc, a bracket or brace, an ink stroke and an open freeform
 * (the `a:ln` head/tail ends, on the connector's vocabulary, without a
 * connector's route). Closed shapes list Arrows disabled.
 */
export function presentationShapeTakesLineEnds(input: {
  shape?: unknown;
  customPath?: unknown;
}): boolean {
  const kind = normalizeShapeKind(input.shape);
  if (isOpenBasicShapeKind(kind) || kind === "ink") return true;
  if (kind !== "custom-geometry") return false;
  return normalizePresentationCustomPath(input.customPath)?.closed === false;
}

export interface PresentationOpenShapeLineEnds {
  readonly headEnd: PresentationConnectorEnd;
  readonly tailEnd: PresentationConnectorEnd;
  readonly headEndWidth?: PresentationConnectorEndSize;
  readonly headEndLength?: PresentationConnectorEndSize;
  readonly tailEndWidth?: PresentationConnectorEndSize;
  readonly tailEndLength?: PresentationConnectorEndSize;
}

/**
 * An open shape's ends (an absent end is none), or null when one is outside
 * `a:ST_LineEndType` or a size is stated for an undecorated end -- the rule
 * the page validator applies on the server.
 */
export function normalizePresentationOpenShapeLineEnds(input: {
  headEnd?: unknown;
  tailEnd?: unknown;
  headEndWidth?: unknown;
  headEndLength?: unknown;
  tailEndWidth?: unknown;
  tailEndLength?: unknown;
}): PresentationOpenShapeLineEnds | null {
  const ends: Record<string, string> = {};
  for (const side of ["head", "tail"] as const) {
    const end = input[`${side}End`] ?? "none";
    if (typeof end !== "string" || !PRESENTATION_CONNECTOR_END_SET.has(end)) return null;
    ends[`${side}End`] = end;
    for (const suffix of ["Width", "Length"] as const) {
      const size = input[`${side}End${suffix}`];
      if (size === undefined) continue;
      if (typeof size !== "string" || !PRESENTATION_CONNECTOR_END_SIZE_SET.has(size) || end === "none") {
        return null;
      }
      ends[`${side}End${suffix}`] = size;
    }
  }
  return ends as unknown as PresentationOpenShapeLineEnds;
}

const asEndRecord = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;

/**
 * The end fields an open shape draws: its own when it states either end,
 * else the ends its imported drawing carries (a page read before the server
 * wrote them as fields).
 */
export function presentationOpenShapeEndFields(
  block: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  if (block.headEnd !== undefined || block.tailEnd !== undefined) return block;
  const native = asEndRecord(asEndRecord(asEndRecord(block.source)?.nativeShapeDrawing)?.lineEnds);
  const fields: Record<string, unknown> = {};
  for (const side of ["head", "tail"] as const) {
    const end = asEndRecord(native?.[`${side}End`]);
    if (typeof end?.type !== "string") continue;
    fields[`${side}End`] = end.type;
    if (end.w !== undefined) fields[`${side}EndWidth`] = end.w;
    if (end.len !== undefined) fields[`${side}EndLength`] = end.len;
  }
  return fields;
}

/** An open shape's ends as the native drawing's `lineEnds` (`{type, w, len}`), or null for none. */
export function presentationOpenShapeNativeLineEnds(
  ends: PresentationOpenShapeLineEnds,
): { headEnd?: Record<string, string>; tailEnd?: Record<string, string> } | null {
  const side = (type: string, w?: string, len?: string) =>
    type === "none" ? undefined : { type, ...(w ? { w } : {}), ...(len ? { len } : {}) };
  const headEnd = side(ends.headEnd, ends.headEndWidth, ends.headEndLength);
  const tailEnd = side(ends.tailEnd, ends.tailEndWidth, ends.tailEndLength);
  return headEnd || tailEnd ? { ...(headEnd ? { headEnd } : {}), ...(tailEnd ? { tailEnd } : {}) } : null;
}

/**
 * Resolve a connector block into one strict native tuple. Explicit `straight`
 * is accepted for migration, but callers persist it as absent. Unsupported
 * presets/end types and contradictory route/preset/kind combinations fail
 * closed instead of being coerced into a visually similar connector.
 */
export function normalizePresentationConnectorSpec(input: {
  shape?: unknown;
  route?: unknown;
  connectorPreset?: unknown;
  headEnd?: unknown;
  tailEnd?: unknown;
  headEndWidth?: unknown;
  headEndLength?: unknown;
  tailEndWidth?: unknown;
  tailEndLength?: unknown;
}): PresentationConnectorSpec | null {
  const kind = normalizeShapeKind(input.shape);
  const defaults = presentationConnectorDefaults(kind);
  if (!defaults) return null;

  // `w` / `len` are independent optional attributes of `CT_LineEndProperties`.
  // They carry through verbatim and an absent one stays absent; an unknown
  // value fails closed rather than being rounded to `med`.
  const sizes: {
    -readonly [K in (typeof PRESENTATION_CONNECTOR_END_SIZE_FIELDS)[number]]?: PresentationConnectorEndSize;
  } = {};
  for (const field of PRESENTATION_CONNECTOR_END_SIZE_FIELDS) {
    const value = input[field];
    if (value === undefined) continue;
    if (
      typeof value !== "string" ||
      !PRESENTATION_CONNECTOR_END_SIZE_SET.has(value)
    ) {
      return null;
    }
    sizes[field] = value as PresentationConnectorEndSize;
  }

  if (
    input.route !== undefined &&
    (typeof input.route !== "string" ||
      !PRESENTATION_CONNECTOR_ROUTE_SET.has(input.route))
  ) {
    return null;
  }
  if (
    input.connectorPreset !== undefined &&
    (typeof input.connectorPreset !== "string" ||
      !PRESENTATION_CONNECTOR_PRESET_SET.has(input.connectorPreset))
  ) {
    return null;
  }
  if (
    input.headEnd !== undefined &&
    (typeof input.headEnd !== "string" ||
      !PRESENTATION_CONNECTOR_END_SET.has(input.headEnd))
  ) {
    return null;
  }
  if (
    input.tailEnd !== undefined &&
    (typeof input.tailEnd !== "string" ||
      !PRESENTATION_CONNECTOR_END_SET.has(input.tailEnd))
  ) {
    return null;
  }

  const explicitPreset = input.connectorPreset as
    | PresentationConnectorPreset
    | undefined;
  const explicitRoute = input.route as PresentationConnectorRoute | undefined;
  const connectorPreset =
    explicitPreset ??
    (explicitRoute === "elbow"
      ? "bentConnector3"
      : explicitRoute === "curved"
        ? "curvedConnector3"
        : defaults.connectorPreset);
  const route = explicitRoute ?? routeForConnectorPreset(connectorPreset);
  if (routeForConnectorPreset(connectorPreset) !== route) return null;

  const headEnd =
    (input.headEnd as PresentationConnectorEnd | undefined) ?? defaults.headEnd;
  const tailEnd =
    (input.tailEnd as PresentationConnectorEnd | undefined) ?? defaults.tailEnd;

  // `@w`/`@len` SIZE a decoration, so an end that carries none has nothing to
  // size. The server's save gate has always refused that pair; this half
  // accepted it, which is a 422 on a deck the editor was happy to build — the
  // §7.8 drift, one input class further on.
  for (const field of PRESENTATION_CONNECTOR_END_SIZE_FIELDS) {
    if (sizes[field] === undefined) continue;
    if ((field.startsWith("head") ? headEnd : tailEnd) === "none") return null;
  }

  // Stable catalog kinds pin WHETHER each end carries a decoration. The legacy
  // generic arrow remains a migration carrier for head-only imported
  // connectors; it must still carry at least one real arrow end. Generic line
  // remains strictly end-less.
  if (kind === "line" && (headEnd !== "none" || tailEnd !== "none")) return null;
  if (kind === "arrow" && headEnd === "none" && tailEnd === "none") return null;
  // A preset-identified connector owns its ROUTE; DrawingML keeps the ends on
  // `a:ln`, so any head/tail pair is legal on it. Its preset still may not be
  // swapped for another route's.
  if (PRESENTATION_FREE_END_CONNECTOR_KINDS.has(kind)) {
    if (
      route !== defaults.route ||
      connectorPreset !== defaults.connectorPreset
    ) {
      return null;
    }
  } else if (kind !== "line" && kind !== "arrow") {
    // The historical nine fold route and the PRESENCE of each end into the
    // kind. WHICH decoration sits there is `a:ln`'s business — `@prst` says
    // nothing about it, exactly as PowerPoint separates the line gallery from
    // Format Shape ▸ Line ▸ Arrow type — so a `stealth` head on an
    // `elbow-double-arrow-connector` is the same kind wearing a different
    // arrowhead. Turning an end ON or OFF still changes the kind and still
    // fails closed here (SHAPE-IMPORT-AUDIT.md §7.4/§7.8).
    if (
      route !== defaults.route ||
      connectorPreset !== defaults.connectorPreset ||
      (headEnd === "none") !== (defaults.headEnd === "none") ||
      (tailEnd === "none") !== (defaults.tailEnd === "none")
    ) {
      return null;
    }
  }

  return { route, connectorPreset, headEnd, tailEnd, ...sizes };
}

/** Byte-clean connector fields for a newly authored stable kind. */
export function presentationConnectorFieldsForKind(
  kind: ShapeKind,
): {
  route?: Exclude<PresentationConnectorRoute, "straight">;
  connectorPreset: PresentationConnectorPreset;
  headEnd: PresentationConnectorEnd;
  tailEnd: PresentationConnectorEnd;
} | null {
  const spec = presentationConnectorDefaults(kind);
  if (!spec) return null;
  return {
    ...(spec.route === "straight" ? {} : { route: spec.route }),
    connectorPreset: spec.connectorPreset,
    headEnd: spec.headEnd,
    tailEnd: spec.tailEnd,
  };
}

/**
 * DrawingML line-callout presets. These are text-bearing autoshapes, not line
 * connectors, but their default visual is stroke-only: an open leader plus an
 * optional text-frame border/accent bar. Keeping the set explicit prevents a
 * stale fill value from repainting them as solid polygons in any renderer.
 */
export const LINE_CALLOUT_SHAPE_KINDS: ReadonlySet<ShapeKind> = new Set([
  "border-callout-1",
  "border-callout-2",
  "border-callout-3",
  "accent-callout-1",
  "accent-callout-2",
  "accent-callout-3",
  "callout-1",
  "callout-2",
  "callout-3",
  "accent-border-callout-1",
  "accent-border-callout-2",
  "accent-border-callout-3",
]);

export function isLineCalloutShapeKind(kind: ShapeKind): boolean {
  return LINE_CALLOUT_SHAPE_KINDS.has(kind);
}

/** DrawingML Basic Shapes whose default geometry is an open stroked path. */
// `line-inverse` is deliberately NOT here: this set means "stroke-only
// multi-part outline", and `lineInv` is a single degenerate path that encloses
// no area, so a fill on it is already invisible.
export const OPEN_BASIC_SHAPE_KINDS: ReadonlySet<ShapeKind> = new Set([
  "arc",
  "double-bracket",
  "double-brace",
  "left-bracket",
  "right-bracket",
  "left-brace",
  "right-brace",
]);

export function isOpenBasicShapeKind(kind: ShapeKind): boolean {
  return OPEN_BASIC_SHAPE_KINDS.has(kind);
}

/** The shape kind used when a block has no explicit (or invalid) value. */
export const DEFAULT_SHAPE_KIND: ShapeKind = "rectangle";

const SHAPE_KIND_SET: ReadonlySet<string> = new Set(SHAPE_KINDS);

/** Fallback fill / stroke when a shape omits (or mis-specifies) colours. */
export const DEFAULT_SHAPE_FILL = "#dbeafe";
export const DEFAULT_SHAPE_STROKE = "#2563eb";
/** Default outline weight (px) and the inclusive bound a value is clamped to. */
export const DEFAULT_STROKE_WIDTH = 2;
export const MAX_STROKE_WIDTH = 24;
/** Default gradient sweep angle (degrees), top→bottom like PowerPoint. */
export const DEFAULT_GRADIENT_ANGLE = 90;

/** Outline dash patterns (PowerPoint parity subset). */
export const STROKE_STYLES = ["solid", "dashed", "dotted"] as const;
export type StrokeStyle = (typeof STROKE_STYLES)[number];
const STROKE_STYLE_SET: ReadonlySet<string> = new Set(STROKE_STYLES);

/** Six-digit hex colour, e.g. `#1a2b3c`. Matches the slide background regex. */
const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

/**
 * The supported object BLEND modes (PowerPoint/WPS "Transparency / blend" parity
 * subset → CSS `mix-blend-mode`). A whitelisted set of safe, widely-supported
 * compositing modes; `normal` is the historical default (ordinary stacking, no
 * blend). Shared by shapes and pictures so both composite identically.
 */
export const SHAPE_BLEND_MODES = [
  "normal",
  "multiply",
  "screen",
  "overlay",
  "darken",
  "lighten",
] as const;
export type ShapeBlendMode = (typeof SHAPE_BLEND_MODES)[number];
const SHAPE_BLEND_MODE_SET: ReadonlySet<string> = new Set(SHAPE_BLEND_MODES);
/** The blend mode used when a block has no explicit (or invalid) value. */
export const DEFAULT_SHAPE_BLEND_MODE: ShapeBlendMode = "normal";

/**
 * Normalize a persisted blend mode into an allowed CSS `mix-blend-mode`,
 * defaulting to {@link DEFAULT_SHAPE_BLEND_MODE} (`normal`) for an unknown /
 * non-string value. Whitelisted so a stale or hand-edited deck can never inject
 * an arbitrary / unsupported compositing keyword.
 */
export function normalizeShapeBlendMode(value: unknown): ShapeBlendMode {
  return typeof value === "string" && SHAPE_BLEND_MODE_SET.has(value)
    ? (value as ShapeBlendMode)
    : DEFAULT_SHAPE_BLEND_MODE;
}







export function presentationShapeBlendCss(block: { type?: unknown; blendMode?: unknown }): { mixBlendMode: ShapeBlendMode } | null {
  if (block.type !== "shape") return null;
  const mode = normalizeShapeBlendMode(block.blendMode);
  return mode === DEFAULT_SHAPE_BLEND_MODE ? null : { mixBlendMode: mode };
}

/**
 * Type guard: is `value` exactly one of the allowed shape identifiers?
 */
export function isShapeKind(value: unknown): value is ShapeKind {
  return typeof value === "string" && SHAPE_KIND_SET.has(value);
}

/**
 * Normalize an arbitrary persisted value into an allowed shape kind. Unknown,
 * empty, or non-string values collapse to {@link DEFAULT_SHAPE_KIND} so a stale
 * or hand-edited deck can never inject an unsupported shape. Common synonyms
 * (`circle` → `ellipse`, `rect`/`square` → `rectangle`) are accepted. The
 * comparison is case-insensitive and trims surrounding whitespace.
 */
export function normalizeShapeKind(value: unknown): ShapeKind {
  if (typeof value !== "string") return DEFAULT_SHAPE_KIND;
  const candidate = value.trim().toLowerCase();
  if (candidate === "circle" || candidate === "oval") return "ellipse";
  if (candidate === "rect" || candidate === "square") return "rectangle";
  if (candidate === "roundrect" || candidate === "rounded") return "rounded-rectangle";
  if (candidate === "star5") return "star";
  if (candidate === "speech" || candidate === "speech-bubble") return "callout";
  if (candidate === "righttriangle" || candidate === "right_triangle") {
    return "right-triangle";
  }
  if (candidate === "cross") return "plus";
  if (candidate === "cylinder") return "can";
  if (candidate === "crescent" || candidate === "crescentmoon") return "moon";
  if (candidate === "lshape" || candidate === "l-shape") return "corner";
  if (candidate === "tear" || candidate === "drop") return "teardrop";
  if (candidate === "stadium") return "terminator";
  if (candidate === "lightning" || candidate === "bolt") return "lightning-bolt";
  return isShapeKind(candidate) ? candidate : DEFAULT_SHAPE_KIND;
}

/** Deterministic 0..100 preview geometry for the higher-point star presets. */
function regularStarPolygonPoints(pointCount: number): string {
  const innerRadius = pointCount <= 8 ? 22 : pointCount <= 16 ? 28 : 34;
  return Array.from({ length: pointCount * 2 }, (_, index) => {
    const radius = index % 2 === 0 ? 48 : innerRadius;
    const angle = -Math.PI / 2 + (Math.PI * index) / pointCount;
    const x = Number((50 + Math.cos(angle) * radius).toFixed(2));
    const y = Number((50 + Math.sin(angle) * radius).toFixed(2));
    return `${x},${y}`;
  }).join(" ");
}

/**
 * SVG polygon points (0..100 viewBox, non-uniform scaled to the block box)
 * for the polygonal shape kinds, or `null` for kinds that render as a CSS box
 * (rectangle family, ellipse) or an SVG line (line/arrow). One source of
 * truth shared by the editor canvas and the present overlay.
 */
export function shapePolygonPoints(kind: ShapeKind): string | null {
  switch (kind) {
    case "triangle":
      return "50,4 96,96 4,96";
    case "diamond":
      return "50,2 98,50 50,98 2,50";
    case "pentagon":
      return "50,2 98,38 79,96 21,96 2,38";
    case "hexagon":
      return "25,4 75,4 98,50 75,96 25,96 2,50";
    case "heptagon":
      return "50,2 87.5,20.1 96.8,60.7 70.8,93.2 29.2,93.2 3.2,60.7 12.5,20.1";
    case "decagon":
      return "64.8,4.3 88.8,21.8 98,50 88.8,78.2 64.8,95.7 35.2,95.7 11.2,78.2 2,50 11.2,21.8 35.2,4.3";
    case "dodecagon":
      return "62.4,3.6 83.9,16.1 96.4,37.6 96.4,62.4 83.9,83.9 62.4,96.4 37.6,96.4 16.1,83.9 3.6,62.4 3.6,37.6 16.1,16.1 37.6,3.6";
    case "corner":
      return "2,2 40,2 40,60 98,60 98,98 2,98";
    case "star4":
      return "50,2 61,39 98,50 61,61 50,98 39,61 2,50 39,39";
    case "star6":
      return "50,2 62,29 92,26 74,50 92,74 62,71 50,98 38,71 8,74 26,50 8,26 38,29";
    case "star":
      return "50,2 61,35 98,35 68,57 79,92 50,70 21,92 32,57 2,35 39,35";
    case "star7":
      return regularStarPolygonPoints(7);
    case "star8":
      return regularStarPolygonPoints(8);
    case "star10":
      return regularStarPolygonPoints(10);
    case "star12":
      return regularStarPolygonPoints(12);
    case "star16":
      return regularStarPolygonPoints(16);
    case "star24":
      return regularStarPolygonPoints(24);
    case "star32":
      return regularStarPolygonPoints(32);
    case "explosion-1":
      return "50,2 58,25 75,10 73,34 96,26 78,48 98,60 73,64 84,92 58,74 50,98 42,74 16,92 27,64 2,60 22,48 4,26 27,34 25,10 42,25";
    case "explosion-2":
      return "50,2 62,28 82,8 78,36 98,42 76,55 92,82 64,72 54,98 45,73 18,92 28,63 2,54 30,45 12,18 41,31";
    case "chevron":
      return "2,4 75,4 98,50 75,96 2,96 25,50";
    case "rectangle-callout":
      // Rectangular speech bubble: sharp body + lower-left tail.
      return "4,4 96,4 96,68 40,68 22,94 26,68 4,68";
    case "parallelogram":
      return "22,4 98,4 78,96 2,96";
    case "trapezoid":
      return "22,4 78,4 98,96 2,96";
    case "non-isosceles-trapezoid":
      return "18,4 88,4 98,96 2,96";
    case "diagonal-stripe":
      return "18,4 48,4 82,96 52,96";
    case "half-frame":
      return "4,4 96,4 96,96 72,96 72,28 28,28 28,72 4,72";
    case "right-triangle":
      return "4,4 4,96 96,96";
    case "octagon":
      return "30,2 70,2 98,30 98,70 70,98 30,98 2,70 2,30";
    case "plus":
      return "35,2 65,2 65,35 98,35 98,65 65,65 65,98 35,98 35,65 2,65 2,35 35,35";
    case "lightning-bolt":
      return "38,2 78,2 58,36 82,36 30,98 44,52 20,52";
    case "right-arrow":
      return "2,30 60,30 60,8 98,50 60,92 60,70 2,70";
    case "left-arrow":
      return "98,30 40,30 40,8 2,50 40,92 40,70 98,70";
    case "up-arrow":
      return "30,98 30,40 8,40 50,2 92,40 70,40 70,98";
    case "down-arrow":
      return "30,2 70,2 70,60 92,60 50,98 8,60 30,60";
    case "left-right-arrow":
      return "2,50 25,15 25,35 75,35 75,15 98,50 75,85 75,65 25,65 25,85";
    case "quad-arrow":
      return "50,2 68,22 68,34 78,34 78,22 98,50 78,78 78,66 68,66 68,78 50,98 32,78 32,66 22,66 22,78 2,50 22,22 22,34 32,34 32,22";
    case "left-right-up-arrow":
      return "50,2 70,25 60,25 60,36 78,36 78,24 98,50 78,76 78,64 22,64 22,76 2,50 22,24 22,36 40,36 40,25 30,25";
    case "left-up-arrow":
      return "50,2 70,24 60,24 60,58 22,58 22,70 2,50 22,30 22,42 40,42 40,24 30,24";
    case "bent-up-arrow":
      return "4,70 58,70 58,38 46,38 70,4 94,38 82,38 82,94 58,94 58,84 4,84";
    case "notched-right-arrow":
      return "2,30 62,30 62,10 98,50 62,90 62,70 2,70 18,50";
    case "home-plate":
      return "4,6 62,6 96,50 62,94 4,94";
    case "right-arrow-callout":
      return "4,14 64,14 64,30 98,50 64,70 64,86 4,86";
    case "down-arrow-callout":
      return "14,4 86,4 86,64 70,64 50,98 30,64 14,64";
    case "left-arrow-callout":
      return "96,14 36,14 36,30 2,50 36,70 36,86 96,86";
    case "up-arrow-callout":
      return "14,96 86,96 86,36 70,36 50,2 30,36 14,36";
    case "left-right-arrow-callout":
      return "2,50 26,18 26,30 74,30 74,18 98,50 74,82 74,70 26,70 26,82";
    case "quad-arrow-callout":
      return "50,2 66,22 66,28 72,28 72,34 78,34 78,22 98,50 78,78 78,66 72,66 72,72 66,72 66,78 50,98 34,78 34,72 28,72 28,66 22,66 22,78 2,50 22,22 22,34 28,34 28,28 34,28 34,22";
    case "bent-arrow":
      // Thick L: up the left leg, right along the top arm, arrowhead right.
      return "2,98 2,40 60,40 60,20 98,50 60,80 60,60 22,60 22,98";
    case "up-down-arrow":
      return "50,2 85,25 65,25 65,75 85,75 50,98 15,75 35,75 35,25 15,25";
    case "card":
      return "18,4 96,4 96,96 4,96 4,18";
    case "data":
      return "18,6 96,6 82,94 4,94";
    case "decision":
      return "50,4 96,50 50,96 4,50";
    case "extract":
      return "50,4 96,94 4,94";
    case "manual-input":
      return "18,18 96,4 96,94 4,94";
    case "manual-operation":
      return "4,6 96,6 78,94 22,94";
    case "merge":
      return "4,6 96,6 50,94";
    case "off-page-connector":
      return "4,6 96,6 96,66 50,96 4,66";
    case "preparation":
      return "20,6 80,6 98,50 80,94 20,94 2,50";
    default:
      return null;
  }
}

/**
 * SVG path (`d`, 0..100 viewBox, non-uniform scaled) for the CURVED shape
 * kinds a polygon cannot carry. Same contract as {@link shapePolygonPoints}:
 * one source of truth for the editor canvas and the present overlay; `null`
 * for every straight-edge / box / linear kind.
 */
/**
 * Kinds whose multi-subpath outline encodes a CUTOUT (evenodd); every other
 * path kind unions overlapping subpaths (nonzero) — an oval callout's tail
 * must merge with its body, not punch a hole in it.
 */
export const HOLLOW_PATH_KINDS: ReadonlySet<string> = new Set([
  "frame",
  "donut",
  "no-symbol",
  "block-arc",
]);

export function shapePathD(kind: ShapeKind): string | null {
  switch (kind) {
    // Default DrawingML adjustment/orientation for the seven asymmetric
    // rectangle presets. Coordinates use the same 0..100 viewBox as every
    // other shared Presentation/Docs preview path.
    // `lineInv` is `line` on the other diagonal: bottom-left to top-right.
    case "line-inverse":
      return "M4 96 L96 4";
    case "single-round-corner":
      return "M4 4 H78 Q96 4 96 22 V96 H4 Z";
    case "same-side-corners-rounded":
      return "M22 4 H78 Q96 4 96 22 V96 H4 V22 Q4 4 22 4 Z";
    case "diagonal-corners-rounded":
      return "M22 4 H96 V78 Q96 96 78 96 H4 V22 Q4 4 22 4 Z";
    case "single-corner-snipped":
      return "M4 4 H78 L96 22 V96 H4 Z";
    case "same-side-corners-snipped":
      return "M22 4 H78 L96 22 V96 H4 V22 Z";
    case "diagonal-corners-snipped":
      return "M4 4 H78 L96 22 V96 H22 L4 78 Z";
    case "single-corner-rounded-snipped":
      return "M22 4 H78 L96 22 V96 H4 V22 Q4 4 22 4 Z";
    // Default, adjustment-free outlines for DrawingML's six Equation Shapes.
    // Disconnected bars/dots stay subpaths of one geometry so fill, outline,
    // selection and slideshow rendering continue through the shared path.
    case "math-plus":
      return "M40 8 H60 V40 H92 V60 H60 V92 H40 V60 H8 V40 H40 Z";
    case "math-minus":
      return "M8 40 H92 V60 H8 Z";
    case "math-multiply":
      return "M20 8 L50 38 L80 8 L92 20 L62 50 L92 80 L80 92 L50 62 L20 92 L8 80 L38 50 L8 20 Z";
    case "math-divide":
      return "M50 8 A10 10 0 1 0 50 28 A10 10 0 1 0 50 8 Z M8 42 H92 V58 H8 Z M50 72 A10 10 0 1 0 50 92 A10 10 0 1 0 50 72 Z";
    case "math-equal":
      return "M8 28 H92 V44 H8 Z M8 56 H92 V72 H8 Z";
    case "math-not-equal":
      return "M8 28 H92 V44 H8 Z M8 56 H92 V72 H8 Z M68 10 L84 20 L32 90 L16 80 Z";
    // Stars/Banners curved and folded presets keep distinct authored previews
    // while their persisted identity remains the exact catalog prstGeom.
    case "up-ribbon":
      return "M14 28 H86 V70 H64 L50 92 L36 70 H14 Z M14 38 H2 L10 52 L2 66 H28 V38 Z M86 38 H98 L90 52 L98 66 H72 V38 Z";
    case "down-ribbon":
      return "M14 30 H36 L50 8 L64 30 H86 V72 H14 Z M14 40 H2 L10 54 L2 68 H28 V40 Z M86 40 H98 L90 54 L98 68 H72 V40 Z";
    case "curved-up-ribbon":
      return "M10 42 C30 18 70 18 90 42 L90 76 C70 52 30 52 10 76 Z M10 48 H2 L10 62 L2 78 H28 L24 58 Z M90 48 H98 L90 62 L98 78 H72 L76 58 Z";
    case "curved-down-ribbon":
      return "M10 24 C30 48 70 48 90 24 L90 58 C70 82 30 82 10 58 Z M10 22 H2 L10 38 L2 52 H28 L24 42 Z M90 22 H98 L90 38 L98 52 H72 L76 42 Z";
    case "vertical-scroll":
      return "M28 8 H78 C88 8 94 14 94 24 C94 34 88 40 78 40 H72 V88 H28 C18 88 12 82 12 72 C12 62 18 56 28 56 H34 V24 C34 14 28 8 18 8 Z M18 8 C28 8 34 14 34 24 H18 C12 24 8 20 8 16 C8 12 12 8 18 8 Z M28 72 H72 C78 72 82 76 82 80 C82 84 78 88 72 88 H28 C34 88 38 84 38 80 C38 76 34 72 28 72 Z";
    case "horizontal-scroll":
      return "M8 28 C8 18 14 12 24 12 C34 12 40 18 40 28 V34 H88 V72 C88 82 82 88 72 88 C62 88 56 82 56 72 V66 H24 C14 66 8 60 8 50 Z M8 28 C8 34 12 38 16 38 C20 38 24 34 24 28 V12 C14 12 8 18 8 28 Z M72 66 V72 C72 78 68 82 64 82 C60 82 56 78 56 72 V66 Z";
    case "wave":
      return "M4 34 C24 12 38 12 54 30 C70 48 82 48 96 34 V68 C78 82 62 82 46 64 C30 46 18 46 4 68 Z";
    case "double-wave":
      return "M4 24 C20 10 34 10 50 24 C66 38 80 38 96 24 V48 C80 62 66 62 50 48 C34 34 20 34 4 48 Z M4 58 C20 44 34 44 50 58 C66 72 80 72 96 58 V82 C80 94 66 94 50 80 C34 66 20 66 4 80 Z";
    case "heart":
      return "M50 88 C20 64 2 46 2 28 C2 12 14 4 27 4 C37 4 46 10 50 19 C54 10 63 4 73 4 C86 4 98 12 98 28 C98 46 80 64 50 88 Z";
    case "cloud":
      return "M25 78 C10 78 2 68 2 58 C2 48 9 41 18 40 C18 26 29 16 42 16 C52 16 61 22 65 31 C68 29 72 28 76 28 C88 28 98 38 98 50 C98 64 88 78 74 78 Z";
    case "can":
      return "M2 16 C2 9 23 4 50 4 C77 4 98 9 98 16 L98 84 C98 91 77 96 50 96 C23 96 2 91 2 84 Z M2 16 C2 23 23 28 50 28 C77 28 98 23 98 16";
    case "chord":
      return "M14 86 A48 48 0 1 1 86 86 Z";
    case "moon":
      return "M64 4 C31 12 10 34 10 50 C10 66 31 88 64 96 C42 84 30 68 30 50 C30 32 42 16 64 4 Z";
    case "teardrop":
      return "M50 4 A46 46 0 1 0 96 50 L96 4 Z";
    case "sun":
      return (
        "M50 28 A22 22 0 1 1 50 72 A22 22 0 1 1 50 28 Z " +
        "M43 2 L57 2 L50 16 Z M84 9 L94 19 L77 23 Z " +
        "M98 43 L98 57 L84 50 Z M94 84 L84 94 L77 77 Z " +
        "M57 98 L43 98 L50 84 Z M16 94 L6 84 L23 77 Z " +
        "M2 57 L2 43 L16 50 Z M6 16 L16 6 L23 23 Z"
      );
    case "pie":
      return "M50 50 L50 4 A46 46 0 1 0 96 50 Z";
    case "pie-wedge":
      return "M50 50 L88 18 A46 46 0 0 1 88 82 Z";
    case "plaque":
      return "M18 4 H82 C82 14 86 18 96 18 V82 C86 82 82 86 82 96 H18 C18 86 14 82 4 82 V18 C14 18 18 14 18 4 Z";
    case "no-symbol":
      return "M50 4 A46 46 0 1 0 50 96 A46 46 0 1 0 50 4 Z M50 22 A28 28 0 1 1 50 78 A28 28 0 1 1 50 22 Z M22 30 L30 22 L78 70 L70 78 Z";
    case "block-arc":
      return "M10 78 A50 50 0 0 1 90 18 L72 36 A26 26 0 0 0 28 78 Z M31 72 A23 23 0 0 1 69 40 L58 51 A10 10 0 0 0 42 72 Z";
    case "terminator":
      // Stadium: straight top/bottom, semicircular ends.
      return "M25 10 H75 A23 40 0 0 1 75 90 H25 A23 40 0 0 1 25 10 Z";
    case "u-turn-arrow":
      // Thick U with an arrowhead on the returning arm.
      return "M6 98 L6 40 A34 34 0 0 1 74 40 L74 52 L94 52 L64 86 L34 52 L54 52 L54 40 A14 14 0 0 0 26 40 L26 98 Z";
    case "curved-right-arrow":
      return "M8 78 C24 36 50 22 76 38 L68 20 L96 32 L84 60 L78 46 C58 32 34 44 20 84 Z";
    case "curved-left-arrow":
      return "M92 78 C76 36 50 22 24 38 L32 20 L4 32 L16 60 L22 46 C42 32 66 44 80 84 Z";
    case "curved-up-arrow":
      return "M20 92 C18 58 34 32 58 22 L40 16 L62 2 L80 25 L66 23 C42 34 30 58 34 88 Z";
    case "curved-down-arrow":
      return "M80 8 C82 42 66 68 42 78 L60 84 L38 98 L20 75 L34 77 C58 66 70 42 66 12 Z";
    case "striped-right-arrow":
      return "M24 30 H62 V10 L98 50 L62 90 V70 H24 Z M4 30 H12 V70 H4 Z M15 30 H20 V70 H15 Z";
    case "circular-arrow":
      // A single connected ring band with an integrated arrowhead. The inner
      // arc remains genuinely hollow without relying on a filled-circle shell.
      return "M50 6 A44 44 0 1 0 88 72 L98 50 L76 36 L80 52 A30 30 0 1 1 50 20 Z";
    case "callout":
      // Rounded rectangular callout. This used to share the sharp polygon
      // preview with wedgeRectCallout; the rounded body makes the two exact
      // DrawingML identities visually distinct while retaining the tail.
      return "M18 4 H82 Q96 4 96 18 V58 Q96 72 82 72 H42 L20 96 L25 72 H18 Q4 72 4 58 V18 Q4 4 18 4 Z";
    case "oval-callout":
      // Ellipse body UNIONed with its tail wedge (overlapping subpaths under
      // fill-rule nonzero merge — which is why hollow kinds carry their own
      // evenodd list instead of a global rule).
      return "M50 4 A48 34 0 1 0 50 72 A48 34 0 1 0 50 4 Z M30 60 L18 96 L52 70 Z";
    case "cloud-callout":
      return (
        "M25 62 C12 62 4 54 4 45 C4 37 10 31 17 30 C17 19 26 10 37 10 C45 10 52 15 55 22 C58 20 61 19 65 19 C75 19 83 27 83 37 C83 48 75 62 63 62 Z " +
        "M74 74 A5 4 0 1 0 74 82 A5 4 0 1 0 74 74 Z M88 86 A4 3 0 1 0 88 92 A4 3 0 1 0 88 86 Z"
      );
    case "document":
      // Flowchart document: rect with the classic wavy bottom.
      return "M2 4 H98 V82 C82 70 66 94 50 88 C34 82 18 98 2 88 Z";
    case "alternate-process":
      return "M20 6 H80 Q96 6 96 22 V78 Q96 94 80 94 H20 Q4 94 4 78 V22 Q4 6 20 6 Z";
    case "collate":
      return "M6 6 H94 L50 50 L94 94 H6 L50 50 Z";
    case "flowchart-connector":
      // Filled circular Flowchart Connector AutoShape, not a line connector.
      return "M50 6 A44 44 0 1 0 50 94 A44 44 0 1 0 50 6 Z";
    case "delay":
      return "M4 6 H58 C84 6 96 24 96 50 C96 76 84 94 58 94 H4 Z";
    case "display":
      return "M18 6 H72 L96 50 L72 94 H18 Q2 50 18 6 Z";
    case "offline-storage":
      return "M14 6 H96 L82 94 H4 Z";
    case "process":
      return "M6 8 H94 V92 H6 Z";
    case "punched-tape":
      return "M4 18 C24 2 40 2 50 18 C60 34 76 34 96 18 V82 C76 98 60 98 50 82 C40 66 24 66 4 82 Z";
    case "sequential-access-storage":
      return "M50 6 A44 44 0 1 0 76 86 L96 94 L84 70 A34 34 0 1 1 50 16 Z";
    case "stored-data":
      return "M18 6 H92 C78 22 78 78 92 94 H18 C2 94 2 6 18 6 Z";
    case "frame":
      // Outer rect minus inner rect — rendered with fill-rule evenodd.
      return "M2 2 H98 V98 H2 Z M22 22 H78 V78 H22 Z";
    case "donut":
      return "M50 2 A48 48 0 1 0 50 98 A48 48 0 1 0 50 2 Z M50 26 A24 24 0 1 1 50 74 A24 24 0 1 1 50 26 Z";
    // The single-outline members of the presets the shared catalog had never
    // carried (SHAPE-IMPORT-AUDIT.md §2.6). Each `d` is the shape's own Apache
    // POI default preset program evaluated at 100 x 100 and rounded to 1/100 of
    // a unit, so the outline here, the slide canvas and the exported
    // `a:prstGeom` are the same geometry rather than three hand-drawn guesses.
    // `funnel` is deliberately two subpaths: DrawingML draws the cone and its
    // rim separately.
    case "gear-6":
      return "M74.82 25.33 L89.58 20.88 L95.01 30.28 L83.78 40.84 A35 35 0 0 1 83.78 59.16 L95.01 69.72 L89.58 79.12 L74.82 74.67 A35 35 0 0 1 58.95 83.84 L55.43 98.84 L44.57 98.84 L41.05 83.84 A35 35 0 0 1 25.18 74.67 L10.42 79.12 L4.99 69.72 L16.22 59.16 A35 35 0 0 1 16.22 40.84 L4.99 30.28 L10.42 20.88 L25.18 25.33 A35 35 0 0 1 41.05 16.16 L44.57 1.16 L55.43 1.16 L58.95 16.16 A35 35 0 0 1 74.82 25.33 Z";
    case "gear-9":
      return "M70.98 15.94 L78.76 9.42 L84.97 14.63 L79.9 23.42 A40 40 0 0 1 87.96 37.4 L98.12 37.4 L99.53 45.39 L89.98 48.86 A40 40 0 0 1 87.18 64.75 L94.96 71.27 L90.9 78.3 L81.36 74.83 A40 40 0 0 1 69 85.2 L70.77 95.2 L63.14 97.97 L58.07 89.18 A40 40 0 0 1 41.93 89.18 L36.86 97.97 L29.23 95.2 L31 85.2 A40 40 0 0 1 18.64 74.83 L9.1 78.3 L5.04 71.27 L12.82 64.75 A40 40 0 0 1 10.02 48.86 L0.47 45.39 L1.88 37.4 L12.04 37.4 A40 40 0 0 1 20.1 23.42 L15.03 14.63 L21.24 9.42 L29.02 15.94 A40 40 0 0 1 44.18 10.43 L45.94 0.43 L54.06 0.43 L55.82 10.43 A40 40 0 0 1 70.98 15.94 Z";
    case "funnel":
      return "M0.49 28.48 A50 25 0 0 1 50 0 A50 25 0 0 1 99.51 28.48 L62.38 94.62 A12.5 6.25 0 0 1 37.62 94.62 Z M5 25 A45 20 0 0 0 95 25 A45 20 0 0 0 5 25 Z";
    case "left-circular-arrow":
      return "M6.25 50 L18.75 50 A31.25 31.25 0 0 0 78.76 62.23 L72.95 62.23 L87.5 50 L97.95 62.23 L92.01 62.23 A43.75 43.75 0 0 1 6.25 50 Z";
    case "left-right-circular-arrow":
      return "M12.5 50 L2.05 37.77 L7.99 37.77 A43.75 43.75 0 0 1 92.01 37.77 L97.95 37.77 L87.5 50 L72.95 37.77 L78.76 37.77 A31.25 31.25 0 0 0 21.24 37.77 L27.05 37.77 Z";
    case "swoosh-arrow":
      return "M0 100 Q16.67 33.33 83.33 12.5 L81.92 0 L100 20 L87.56 50 L86.15 37.5 Q25 45.83 0 100 Z";
    case "up-down-arrow-callout":
      return "M0 25.94 L37.5 25.94 L37.5 25 L25 25 L50 0 L75 25 L62.5 25 L62.5 25.94 L100 25.94 L100 74.06 L62.5 74.06 L62.5 75 L75 75 L50 100 L25 75 L37.5 75 L37.5 74.06 L0 74.06 Z";
    // The two fold creases stay subpaths of the one outline, exactly as `frame`
    // and `donut` keep their cutouts: they enclose no area, so they stroke
    // without filling, and the ribbon remains one fill with one outline.
    //
    // Sampled at 2 : 1 rather than square, then expressed in this box. Both
    // arrowheads are sized from the SHORT side, so a square sample collapses
    // them into each other and the ribbon degenerates into a diamond — the
    // aspect distortion this whole 0..100 space carries (audit §2.9), at its
    // worst. 2 : 1 is the proportion PowerPoint itself draws the shape in.
    case "left-right-ribbon":
      return "M0 41.67 L25 0 L25 16.67 L50 16.67 A3.12 4.17 0 0 1 50 25 A3.12 4.17 0 0 0 50 33.33 L75 33.33 L75 16.67 L100 58.33 L75 100 L75 83.33 L50 83.33 A3.12 4.17 0 0 1 46.88 79.17 L46.88 66.67 L25 66.67 L25 83.33 Z M53.12 20.83 L53.12 33.33 M46.88 29.17 L46.88 66.67";
    default:
      return null;
  }
}

/** One drawable part of a multi-part shape kind. */
export interface ShapePart {
  d: string;
  /**
   * Paint role: "body" fills with the shape fill and strokes with the outline
   * (the silhouette); "ink" fills with the OUTLINE color (facial features);
   * "stroke" is an open stroked path, never filled.
   */
  role: "body" | "ink" | "stroke";
}

/**
 * Parts for the multi-part kinds — shapes whose features need different paint
 * per part, which one path with one fill cannot carry (the reason smiley was
 * shelved until this existed). Same 0-100 space and sharing rules as the
 * other outline sources; null for every single-outline kind.
 */
export function shapeMultiParts(kind: ShapeKind): readonly ShapePart[] | null {
  switch (kind) {
    case "smiley":
      return [
        { d: "M50 2 A48 48 0 1 0 50 98 A48 48 0 1 0 50 2 Z", role: "body" },
        { d: "M33 34 A6 7 0 1 0 33 48 A6 7 0 1 0 33 34 Z", role: "ink" },
        { d: "M67 34 A6 7 0 1 0 67 48 A6 7 0 1 0 67 34 Z", role: "ink" },
        { d: "M27 62 Q50 84 73 62", role: "stroke" },
      ];
    case "action-button-back-previous":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M62 25 L34 50 L62 75 Z", role: "ink" },
      ];
    case "action-button-forward-next":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M38 25 L66 50 L38 75 Z", role: "ink" },
      ];
    case "action-button-beginning":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M30 24 H38 V76 H30 Z M68 24 L40 50 L68 76 Z", role: "ink" },
      ];
    case "action-button-end":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M32 24 L60 50 L32 76 Z M62 24 H70 V76 H62 Z", role: "ink" },
      ];
    case "action-button-home":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M24 49 L50 25 L76 49 V76 H57 V58 H43 V76 H24 Z", role: "ink" },
      ];
    case "action-button-return":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M27 35 L16 48 L27 61 V53 H59 Q72 53 72 66 Q72 78 58 78 H42 V68 H58 Q62 68 62 64 Q62 59 57 59 H27 Z", role: "ink" },
      ];
    case "action-button-information":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M46 38 H54 V76 H46 Z M45 23 A5 5 0 1 0 55 23 A5 5 0 1 0 45 23 Z", role: "ink" },
      ];
    case "action-button-movie":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M25 25 H75 V75 H25 Z M25 38 H75 M25 62 H75 M38 25 V75 M62 25 V75", role: "stroke" },
        { d: "M43 37 L65 50 L43 63 Z", role: "ink" },
      ];
    case "action-button-document":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M30 20 H58 L72 34 V80 H30 Z M58 20 V35 H72 M38 48 H64 M38 58 H64 M38 68 H58", role: "stroke" },
      ];
    case "action-button-sound":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M24 42 H37 L53 29 V71 L37 58 H24 Z", role: "ink" },
        { d: "M61 38 Q72 50 61 62 M68 30 Q86 50 68 70", role: "stroke" },
      ];
    case "action-button-help":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
        { d: "M38 36 Q39 24 51 24 Q64 24 64 36 Q64 45 55 49 Q50 52 50 61 M46 72 H54", role: "stroke" },
      ];
    case "action-button-custom":
      return [
        { d: "M5 12 Q5 5 12 5 H88 Q95 5 95 12 V88 Q95 95 88 95 H12 Q5 95 5 88 Z", role: "body" },
      ];
    case "cube":
      return [
        {
          d: "M18 24 L50 6 L84 24 V74 L50 94 L18 74 Z",
          role: "body",
        },
        { d: "M18 24 L50 44 L84 24 M50 44 V94", role: "stroke" },
      ];
    case "bevel":
      return [
        { d: "M6 6 H94 V94 H6 Z", role: "body" },
        { d: "M6 6 L24 24 H76 L94 6 M24 24 V76 L6 94 M24 76 H76 L94 94 M76 24 V76", role: "stroke" },
      ];
    case "folded-corner":
      return [
        { d: "M8 6 H70 L92 28 V94 H8 Z", role: "body" },
        { d: "M70 6 V28 H92", role: "stroke" },
      ];
    case "predefined-process":
      return [
        { d: "M6 8 H94 V92 H6 Z", role: "body" },
        { d: "M20 8 V92 M80 8 V92", role: "stroke" },
      ];
    case "internal-storage":
      return [
        { d: "M6 8 H94 V92 H6 Z", role: "body" },
        { d: "M6 26 H94 M24 8 V92", role: "stroke" },
      ];
    case "multiple-documents":
      return [
        {
          d: "M12 22 H88 V88 C72 76 58 96 44 88 C30 80 20 94 12 88 Z",
          role: "body",
        },
        { d: "M20 14 H92 V78", role: "stroke" },
        { d: "M28 6 H96 V68", role: "stroke" },
      ];
    case "magnetic-disk":
      return [
        {
          d: "M10 18 C10 8 28 4 50 4 C72 4 90 8 90 18 V82 C90 92 72 96 50 96 C28 96 10 92 10 82 Z",
          role: "body",
        },
        {
          d: "M10 18 C10 28 28 32 50 32 C72 32 90 28 90 18 M10 82 C10 72 28 68 50 68 C72 68 90 72 90 82",
          role: "stroke",
        },
      ];
    case "direct-access-storage":
      return [
        {
          d: "M18 8 H78 C90 8 96 26 96 50 C96 74 90 92 78 92 H18 C6 92 4 74 4 50 C4 26 6 8 18 8 Z",
          role: "body",
        },
        {
          d: "M78 8 C66 8 64 26 64 50 C64 74 66 92 78 92 M18 8 C30 8 32 26 32 50 C32 74 30 92 18 92",
          role: "stroke",
        },
      ];
    case "flowchart-or":
      return [
        {
          d: "M50 6 A44 44 0 1 0 50 94 A44 44 0 1 0 50 6 Z",
          role: "body",
        },
        { d: "M50 16 V84 M16 50 H84", role: "stroke" },
      ];
    case "flowchart-summing-junction": {
      const lo = 50 - 50 / Math.sqrt(2), hi = 50 + 50 / Math.sqrt(2);
      return [
        { d: "M0 50 A50 50 0 1 0 100 50 A50 50 0 1 0 0 50 Z", role: "body" },
        { d: `M${lo} ${lo} L${hi} ${hi} M${hi} ${lo} L${lo} ${hi}`, role: "stroke" },
      ];
    }
    case "sort":
      return [
        { d: "M50 4 L96 50 L50 96 L4 50 Z", role: "body" },
        { d: "M18 50 H82", role: "stroke" },
      ];
    case "arc":
      return [{ d: "M8 82 C16 30 55 6 92 48", role: "stroke" }];
    case "double-bracket":
      return [
        { d: "M32 8 H16 V92 H32", role: "stroke" },
        { d: "M68 8 H84 V92 H68", role: "stroke" },
      ];
    case "double-brace":
      return [
        { d: "M34 8 C22 8 22 24 22 36 C22 46 16 50 8 50 C16 50 22 54 22 64 C22 76 22 92 34 92", role: "stroke" },
        { d: "M66 8 C78 8 78 24 78 36 C78 46 84 50 92 50 C84 50 78 54 78 64 C78 76 78 92 66 92", role: "stroke" },
      ];
    case "left-bracket":
      return [{ d: "M42 8 H20 V92 H42", role: "stroke" }];
    case "right-bracket":
      return [{ d: "M58 8 H80 V92 H58", role: "stroke" }];
    case "left-brace":
      return [
        { d: "M58 8 C38 8 38 24 38 36 C38 46 28 50 16 50 C28 50 38 54 38 64 C38 76 38 92 58 92", role: "stroke" },
      ];
    case "right-brace":
      return [
        { d: "M42 8 C62 8 62 24 62 36 C62 46 72 50 84 50 C72 50 62 54 62 64 C62 76 62 92 42 92", role: "stroke" },
      ];
    case "callout-1":
      return [{ d: "M30 58 L5 92", role: "stroke" }];
    case "callout-2":
      return [{ d: "M30 58 H17 L5 92", role: "stroke" }];
    case "callout-3":
      return [{ d: "M30 58 H17 V76 L5 92", role: "stroke" }];
    case "accent-callout-1":
      return [
        { d: "M34 8 V68", role: "stroke" },
        { d: "M30 58 L5 92", role: "stroke" },
      ];
    case "accent-callout-2":
      return [
        { d: "M34 8 V68", role: "stroke" },
        { d: "M30 58 H17 L5 92", role: "stroke" },
      ];
    case "accent-callout-3":
      return [
        { d: "M34 8 V68", role: "stroke" },
        { d: "M30 58 H17 V76 L5 92", role: "stroke" },
      ];
    case "border-callout-1":
      return [
        { d: "M28 8 H94 V68 H28 V8", role: "stroke" },
        { d: "M28 58 L5 92", role: "stroke" },
      ];
    case "border-callout-2":
      return [
        { d: "M28 8 H94 V68 H28 V8", role: "stroke" },
        { d: "M28 58 H17 L5 92", role: "stroke" },
      ];
    case "border-callout-3":
      return [
        { d: "M28 8 H94 V68 H28 V8", role: "stroke" },
        { d: "M28 58 H17 V76 L5 92", role: "stroke" },
      ];
    case "accent-border-callout-1":
      return [
        { d: "M28 8 H94 V68 H28 V8", role: "stroke" },
        { d: "M38 8 V68", role: "stroke" },
        { d: "M28 58 L5 92", role: "stroke" },
      ];
    case "accent-border-callout-2":
      return [
        { d: "M28 8 H94 V68 H28 V8", role: "stroke" },
        { d: "M38 8 V68", role: "stroke" },
        { d: "M28 58 H17 L5 92", role: "stroke" },
      ];
    case "accent-border-callout-3":
      return [
        { d: "M28 8 H94 V68 H28 V8", role: "stroke" },
        { d: "M38 8 V68", role: "stroke" },
        { d: "M28 58 H17 V76 L5 92", role: "stroke" },
      ];
    // The several-part members of the presets the shared catalog had never
    // carried (SHAPE-IMPORT-AUDIT.md §2.6), transcribed from the same Apache
    // POI default preset programs as the single-outline ones in `shapePathD`.
    // The three chart markers are a filled plate plus an unfilled marker, and
    // the three tab presets are four independent corner plates — neither can
    // be one path with one paint, which is what this function exists for.
    case "chart-plus":
      return [
        { d: "M0 0 L0 100 L100 100 L100 0 Z", role: "body" },
        { d: "M50 0 L50 100 M0 50 L100 50", role: "stroke" },
      ];
    case "chart-star":
      return [
        { d: "M0 0 L0 100 L100 100 L100 0 Z", role: "body" },
        { d: "M0 0 L100 100 M0 100 L100 0 M50 0 L50 100", role: "stroke" },
      ];
    case "chart-x":
      return [
        { d: "M0 0 L0 100 L100 100 L100 0 Z", role: "body" },
        { d: "M0 0 L100 100 M0 100 L100 0", role: "stroke" },
      ];
    case "corner-tabs":
      return [
        { d: "M0 0 L7.07 0 L0 7.07 Z", role: "body" },
        { d: "M0 92.93 L7.07 100 L0 100 Z", role: "body" },
        { d: "M92.93 0 L100 0 L100 7.07 Z", role: "body" },
        { d: "M100 92.93 L100 100 L92.93 100 Z", role: "body" },
      ];
    case "square-tabs":
      return [
        { d: "M0 0 L7.07 0 L7.07 7.07 L0 7.07 Z", role: "body" },
        { d: "M0 92.93 L7.07 92.93 L7.07 100 L0 100 Z", role: "body" },
        { d: "M92.93 0 L100 0 L100 7.07 L92.93 7.07 Z", role: "body" },
        { d: "M92.93 92.93 L100 92.93 L100 100 L92.93 100 Z", role: "body" },
      ];
    case "plaque-tabs":
      return [
        { d: "M0 0 L7.07 0 A7.07 7.07 0 0 1 0 7.07 Z", role: "body" },
        { d: "M0 92.93 A7.07 7.07 0 0 1 7.07 100 L0 100 Z", role: "body" },
        { d: "M100 0 L100 7.07 A7.07 7.07 0 0 1 92.93 0 Z", role: "body" },
        { d: "M92.93 100 A7.07 7.07 0 0 1 100 92.93 L100 100 Z", role: "body" },
      ];
    default:
      return null;
  }
}

/** Upper bound on a custom polygon's vertex count (guards persisted size). */
const MAX_CUSTOM_POLYGON_POINTS = 256;

/**
 * Validate + canonicalize a custom polygon-outline string ("x,y x,y …") carried
 * by a merged shape (Merge Shapes → Intersect). Coordinates live in the shape's
 * own 0–100 normalized space; a generous ±1000 clamp tolerates rounding without
 * letting a bad value blow up the render. Returns null for anything malformed,
 * out of range, under `min` vertices (3; an open ink line passes 2), or over
 * `max` (the polygon cap; ink passes its own) — callers then fall back to
 * the shape kind's preset outline so a corrupt field can never erase the shape.
 */
export function normalizeCustomPolygonPoints(
  value: unknown,
  { min = 3, max = MAX_CUSTOM_POLYGON_POINTS }: { min?: number; max?: number } = {},
): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  const pairs = trimmed.split(/\s+/);
  if (pairs.length < min || pairs.length > max) return null;
  const fmt = (v: number) =>
    Number.isInteger(v) ? String(v) : String(Number(v.toFixed(2)));
  const out: string[] = [];
  for (const pair of pairs) {
    const parts = pair.split(",");
    if (parts.length !== 2) return null;
    const x = Number(parts[0]);
    const y = Number(parts[1]);
    if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
    if (x < -1000 || x > 1000 || y < -1000 || y > 1000) return null;
    out.push(`${fmt(x)},${fmt(y)}`);
  }
  return out.join(" ");
}

/**
 * Normalize an arbitrary colour value into a six-digit hex string, or return
 * `fallback` (which may itself be a valid hex) when the value is missing /
 * malformed. Always lower-cases so persisted decks stay byte-stable. A `null`
 * fallback lets callers express "no colour" (e.g. a transparent line fill).
 */
export function normalizeShapeColor<F extends string | null>(
  value: unknown,
  fallback: F,
): string | F {
  if (typeof value === "string") {
    const trimmed = value.trim().toLowerCase();
    // PPT's 无填充/无轮廓: "none" is a first-class value (SVG fill/stroke
    // "none"; OOXML a:noFill on export).
    if (trimmed === "none") return "none";
    if (HEX_COLOR_RE.test(trimmed)) return trimmed;
  }
  return fallback;
}

/** Clamp a persisted outline weight to a sane px range, else the default. */
export function normalizeStrokeWidth(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_STROKE_WIDTH;
  }
  return Math.min(MAX_STROKE_WIDTH, Math.max(0, value));
}

const INK_STROKE_EMU_PER_PX = 9525;

/**
 * A slide ink stroke's `strokeWidth` in the design pixels every ink states.
 * Record Studio stored its width in EMU, 9525 times too wide (a pen 41 148
 * wide on a 16:9 slide); nothing else writes an ink wider than the thickest
 * outline, so such a width is read back as the pixels it meant. Anything else
 * comes back unchanged. The server's twin is `presentation_ink_stroke_width_px`.
 */
export function presentationInkStrokeWidthPx(value: unknown): unknown {
  return typeof value === "number" && Number.isFinite(value) && value > MAX_STROKE_WIDTH
    ? Math.round((value / INK_STROKE_EMU_PER_PX) * 100) / 100
    : value;
}

/** Normalize a persisted outline dash style, defaulting to solid. */
export function normalizeStrokeStyle(value: unknown): StrokeStyle {
  return typeof value === "string" && STROKE_STYLE_SET.has(value)
    ? (value as StrokeStyle)
    : "solid";
}

/** Normalize a gradient sweep angle into [0, 360); default top→bottom. */
export function normalizeGradientAngle(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_GRADIENT_ANGLE;
  }
  const wrapped = value % 360;
  return wrapped < 0 ? wrapped + 360 : wrapped;
}

/** How a shape's PICTURE fill sizes within the shape box (PowerPoint parity). */
export const SHAPE_FILL_IMAGE_FITS = ["cover", "contain", "tile"] as const;
export type ShapeFillImageFit = (typeof SHAPE_FILL_IMAGE_FITS)[number];
const SHAPE_FILL_IMAGE_FIT_SET: ReadonlySet<string> = new Set(
  SHAPE_FILL_IMAGE_FITS,
);
export const DEFAULT_SHAPE_FILL_IMAGE_FIT: ShapeFillImageFit = "cover";

/** A shape PICTURE fill (image url + sizing), as persisted on the block. */
export interface ShapeFillImage extends OfficePictureFillOptions {
  url: string;
  /**
   * `tile` repeats the picture (Office's "Tile picture as texture", with
   * `tile` stating its offsets, scale, alignment and mirror when present);
   * `cover` / `contain` size it in the box -- unless `fillRect` states the
   * stretch's offsets, when the picture is stretched into that rectangle as
   * Office draws `a:stretch/a:fillRect`.
   */
  fit: ShapeFillImageFit;
}

/** Normalize a persisted fit value, defaulting to `cover`. */
export function normalizeShapeFillImageFit(value: unknown): ShapeFillImageFit {
  return typeof value === "string" && SHAPE_FILL_IMAGE_FIT_SET.has(value)
    ? (value as ShapeFillImageFit)
    : DEFAULT_SHAPE_FILL_IMAGE_FIT;
}

/**
 * Normalize a shape's PICTURE-fill field into a render-ready descriptor, or
 * `null` when there is no usable picture fill. The url is validated through
 * {@link isDrawableImageSource}, so a stale or hand-edited deck can never persist
 * a dangerous `url()`, and a large imported picture still draws; an empty (not-yet-pasted) url also yields `null` (no fill yet). The
 * fit collapses to `cover` when missing / invalid. Side-effect free.
 */
export function normalizeShapeFillImage(value: unknown): ShapeFillImage | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const url = (value as { url?: unknown }).url;
  if (typeof url !== "string") return null;
  const trimmed = url.trim();
  if (trimmed === "" || !isDrawableImageSource(trimmed)) return null;
  const fit = normalizeShapeFillImageFit((value as { fit?: unknown }).fit);
  // Format Shape's picture options: a tile only on a tiled fill, the
  // stretch's offsets only on one that is not (an option outside the shared
  // model is dropped and the fill draws as before it).
  const options = officePictureFillOptionsOf(value);
  const { tile, fillRect, ...rest } = options;
  return {
    url: trimmed,
    fit,
    ...(fit === "tile" && tile ? { tile } : {}),
    ...(fit !== "tile" && fillRect ? { fillRect } : {}),
    ...rest,
  };
}

/** Whether a picture fill states Office's options the CSS box path cannot draw. */
export function shapeFillImageNeedsPattern(image: ShapeFillImage | null | undefined): boolean {
  return Boolean(image && officePictureFillOptionsStated(image));
}

export interface ShapeGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * A sensible default rect for a newly inserted block, centred-ish on the
 * canvas. Text boxes and most shapes get a wide-ish box; line/arrow get a thin
 * horizontal strip. Geometry is clamped to the provided slide bounds (falling
 * back to the standard 16:9-ish canvas) so an inserted block always lands fully
 * on-slide regardless of the deck's dimensions.
 */
/** How far each cascaded insertion steps, as a fraction of the slide. */
const INSERT_CASCADE_STEP = 0.04;
/** After this many steps the cascade restarts rather than marching off-slide. */
const INSERT_CASCADE_LIMIT = 8;

/**
 * Where a newly inserted object lands.
 *
 * Every insertion used to return the SAME spot, so a second shape landed exactly
 * on top of the first: indistinguishable on screen, and a press aimed at the one
 * underneath always hit the one on top. PowerPoint and WPS cascade instead —
 * each new object steps down and to the right of the last one placed there.
 *
 * `occupied` carries the positions already taken on the slide. Omit it and the
 * behaviour is exactly as before, so every existing caller is unchanged.
 */
export function defaultShapeGeometry(
  kind: "textbox" | ShapeKind,
  bounds?: { width?: number | null; height?: number | null },
  occupied?: readonly unknown[],
): ShapeGeometry {
  const slideWidth =
    typeof bounds?.width === "number" && bounds.width > 0
      ? bounds.width
      : DEFAULT_SLIDE_WIDTH_EMU;
  const slideHeight =
    typeof bounds?.height === "number" && bounds.height > 0
      ? bounds.height
      : DEFAULT_SLIDE_HEIGHT_EMU;
  const connector =
    kind === "textbox" ? null : presentationConnectorDefaults(kind);
  const isLinear = connector !== null;
  const isRoutedConnector =
    connector?.route === "elbow" || connector?.route === "curved";
  const width = Math.round(slideWidth * (isLinear ? 0.4 : 0.34));
  // Preserve the historical flat geometry for straight line/arrow insertions.
  // Routed connectors need a real vertical span or their elbow/curve collapses
  // into a visually false almost-straight line.
  const height = Math.round(
    slideHeight * (isRoutedConnector ? 0.22 : isLinear ? 0.04 : 0.22),
  );
  const baseLeft = Math.round(slideWidth * 0.33);
  const baseTop = Math.round(slideHeight * 0.39);
  if (!occupied?.length) {
    return { left: baseLeft, top: baseTop, width, height };
  }

  const stepX = Math.round(slideWidth * INSERT_CASCADE_STEP);
  const stepY = Math.round(slideHeight * INSERT_CASCADE_STEP);
  const numeric = (value: unknown): number | null =>
    typeof value === "number" && Number.isFinite(value) ? value : null;
  const taken = (left: number, top: number): boolean =>
    occupied.some((entry) => {
      // Deliberately loose: callers pass their own block unions, and all this
      // needs from one is whatever `left` / `top` it happens to carry.
      const block = (entry ?? {}) as { left?: unknown; top?: unknown };
      const blockLeft = numeric(block.left);
      const blockTop = numeric(block.top);
      if (blockLeft === null || blockTop === null) return false;
      // "Close enough to look stacked" rather than exact equality: a nudged
      // object should still count as holding that spot.
      return (
        Math.abs(blockLeft - left) < stepX / 2 &&
        Math.abs(blockTop - top) < stepY / 2
      );
    });

  let left = baseLeft;
  let top = baseTop;
  for (let step = 0; step < INSERT_CASCADE_LIMIT && taken(left, top); step += 1) {
    left += stepX;
    top += stepY;
  }
  // Never march an object off the slide; a full cascade restarts at the origin,
  // which is also what PowerPoint does once its own cascade runs out of room.
  if (left + width > slideWidth || top + height > slideHeight) {
    return { left: baseLeft, top: baseTop, width, height };
  }
  return { left, top, width, height };
}

const ELLIPSE_MERGE_SEGMENTS = 64;
const FULL_BOX_OUTLINE: PolygonPoint[] = [
  { x: 0, y: 0 },
  { x: 100, y: 0 },
  { x: 100, y: 100 },
  { x: 0, y: 100 },
];

/** Tessellate the unit ellipse (its 0–100 box) into a convex polygon for merging. */
function ellipseMergeOutline(): PolygonPoint[] {
  const out: PolygonPoint[] = [];
  for (let i = 0; i < ELLIPSE_MERGE_SEGMENTS; i += 1) {
    const angle = (i / ELLIPSE_MERGE_SEGMENTS) * Math.PI * 2;
    out.push({ x: 50 + 50 * Math.cos(angle), y: 50 + 50 * Math.sin(angle) });
  }
  return out;
}

/** A shape block reduced to the fields Merge Shapes needs (geometry + outline). */
export interface MergeableShape {
  shape?: unknown;
  points?: unknown;
  left?: unknown;
  top?: unknown;
  width?: unknown;
  height?: unknown;
  /** Degrees, clockwise on screen — the block's `rotation` contract. */
  rotation?: unknown;
  flipH?: unknown;
  flipV?: unknown;
}

/**
 * The shape's filled outline in its own 0–100 space: a custom polygon when
 * present, else the kind's preset polygon, else a tessellated ellipse, else the
 * full box (rectangle / rounded-rectangle). Lines / arrows have no fillable area
 * → null (they cannot participate in a merge).
 */
function mergeOutlinePoints(block: MergeableShape): PolygonPoint[] | null {
  const custom = normalizeCustomPolygonPoints(block.points);
  if (custom) return parsePolygonString(custom);
  const kind = normalizeShapeKind(block.shape);
  if (
    isPresentationConnectorShapeKind(kind) ||
    kind === "ink" ||
    kind === "custom-geometry"
  ) {
    return null;
  }
  if (kind === "ellipse") return ellipseMergeOutline();
  const preset = shapePolygonPoints(kind);
  if (preset) return parsePolygonString(preset);
  return FULL_BOX_OUTLINE;
}

/** A shape's unrotated canvas box plus the transform drawn on top of it. */
export interface MergeFrame {
  left: number;
  top: number;
  width: number;
  height: number;
  /** Degrees, clockwise on screen (0 when absent / non-finite). */
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

/** The shape's canvas frame (null when any geometry field is unusable). */
function mergeRectOf(block: MergeableShape): MergeFrame | null {
  const left = Number(block.left);
  const top = Number(block.top);
  const width = Number(block.width);
  const height = Number(block.height);
  if (
    !Number.isFinite(left) ||
    !Number.isFinite(top) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }
  const rotation = Number(block.rotation);
  return {
    left,
    top,
    width,
    height,
    rotation: Number.isFinite(rotation) ? rotation : 0,
    flipH: block.flipH === true,
    flipV: block.flipV === true,
  };
}

/** True when the block has a fillable outline and a usable frame, i.e. it can
 * take part in Merge Shapes at all (a line / arrow / ink stroke / curved custom
 * geometry cannot). Lets the menu refuse BEFORE anything changes. */
export function isMergeableShapeBlock(block: MergeableShape): boolean {
  return mergeOutlinePoints(block) !== null && mergeRectOf(block) !== null;
}

/**
 * A mergeable shape's outline in its own 0–100 box and its canvas frame (null
 * when it cannot take part in Merge Shapes). The boolean itself runs on the
 * server (`lib/presentation-shape-merge.ts`, engine `tools`, `shapes.merge`).
 */
export function mergeableShapeOutline(
  block: MergeableShape,
): { points: PolygonPoint[]; frame: MergeFrame } | null {
  const points = mergeOutlinePoints(block);
  const frame = mergeRectOf(block);
  return points && frame ? { points, frame } : null;
}

/**
 * The fields of a shape block Merge Shapes reads (the request to the server
 * carries these and nothing else of the block).
 */
export function mergeableShapeFields(block: MergeableShape): MergeableShape {
  const fields: MergeableShape = {};
  for (const key of ["shape", "points", "left", "top", "width", "height", "rotation", "flipH", "flipV"] as const) {
    if (block[key] !== undefined) fields[key] = block[key];
  }
  return fields;
}

/**
 * The result of Merge Shapes: the fields for a new custom shape. A result that
 * is ONE simple region carries a `points` polygon outline; a result with
 * several regions or a hole carries a multi-figure `customPath` (rendered and
 * exported with one closed sub-path per region / hole, even-odd filled).
 * Exactly one of the two is present.
 */
export interface MergedShapeResult {
  points?: string;
  customPath?: PresentationCustomPath;
  left: number;
  top: number;
  width: number;
  height: number;
}

import {normalizeOfficeShapeGradient,normalizeOfficeShapeShadow,officeGradientAngleForBox,officeGradientCss,officeGradientDrawnStops,officeShadowCss,officeShadowOffset,type OfficeShapeGradient,type OfficeShapeShadow} from "./office-shape-paint";
import { officeRectPathGradientBackground } from "./office-path-gradient";
import { officeStrokeAlign, type OfficeStrokeAlign } from "./office-shape-line-align";
import { officeLineFamilyDash, officeStrokeDash, officeStrokeDashArray, type OfficeStrokeDash } from "./office-stroke-dash";
import { officeStrokeDetail, type OfficeStrokeCap, type OfficeStrokeJoin } from "./office-stroke-detail";
// Re-exported so the canvas aims its SVG drop shadow with the SAME measured
// vector the CSS box shadow uses, instead of deriving a second one.
export { officeShadowOffset };
export interface ResolvedShapeStyle {
  fillGradient?: OfficeShapeGradient;
  shapeShadow?: OfficeShapeShadow;
  kind: ShapeKind;
  fill: string | null;
  /** Second gradient stop; non-null ⇒ the fill is a gradient from fill→fillTo. */
  fillTo: string | null;
  /** Gradient sweep angle in degrees (CSS convention: 90 = top→bottom). */
  gradientAngle: number;
  /**
   * A PICTURE fill (image url + sizing), or `null` for a colour / gradient
   * fill. When present it takes precedence over the solid / gradient fill
   * (PowerPoint: a picture fill replaces the colour). Linear shapes never carry
   * a picture fill.
   */
  fillImage: ShapeFillImage | null;
  /**
   * Format Shape > Fill > Pattern fill: one of Office's 48 presets in its
   * foreground over its background (a page's four-pattern form read as the
   * preset the server writes for it), or null.
   */
  fillPattern: OfficeChartFillPattern | null;
  stroke: string;
  /** Second outline stop; non-null means the outline is a linear gradient. */
  strokeTo: string | null;
  /**
   * The outline's WHOLE gradient (N stops, per-stop identity and opacity) when
   * the block carries one -- the same model `fillGradient` holds. `strokeTo`
   * is still the last stop, so a caller that only knows two colours keeps
   * working.
   */
  strokeGradient?: OfficeShapeGradient;
  /** CSS-convention gradient angle for the outline. */
  strokeGradientAngle: number;
  /** Outline weight in px (0 ⇒ no visible outline). */
  strokeWidth: number;
  strokeStyle: StrokeStyle;
  /**
   * `a:ln@algn`, in the shared vocabulary (`@/lib/office-shape-line-align`).
   * `"inset"` is the `in` pen: the whole stroke belongs INSIDE the geometry
   * path, which is what Office's `LineFormat.InsetPen` states. `null` is a
   * line that states no alignment -- every shape authored here, and every
   * imported one whose `a:ln` carries no `algn` -- and is drawn centred.
   */
  strokeAlign: OfficeStrokeAlign | null;
  /**
   * The outline's exact Office dash (`a:prstDash`) when it refines
   * `strokeStyle`'s family (Docs' `outlineDash`, `@/lib/office-stroke-dash`),
   * and its cap and join (`a:ln@cap`, `a:round|bevel|miter`,
   * `@/lib/office-stroke-detail`): null when the line states none.
   */
  outlineDash: OfficeStrokeDash | null;
  outlineCap: OfficeStrokeCap | null;
  outlineJoin: OfficeStrokeJoin | null;
  /** Whether a soft drop shadow is drawn under the shape. */
  shadow: boolean;
  /** Object transparency as a PERCENTAGE (100 = fully opaque, default). */
  opacity: number;
  /** CSS `mix-blend-mode` (`normal` = ordinary stacking, the default). */
  blendMode: ShapeBlendMode;
  /** Whether the shape renders as an SVG line/arrow rather than a CSS box. */
  isLinear: boolean;
  /** SVG polygon points for polygonal kinds; null for box/linear kinds. */
  polygonPoints: string | null;
  /** Curved outline (SVG path d) — exclusive with polygonPoints. */
  pathD: string | null;
  /** Strict durable custom path, when this block owns bounded `custGeom`. */
  customPath: PresentationCustomPath | null;
  /** Multi-part kinds (per-part paint roles) — exclusive with the others. */
  multiParts: readonly ShapePart[] | null;
}

/**
 * Resolve a shape block's persisted kind + colours + outline + effects into a
 * normalized, render-ready descriptor. Lines/arrows have no meaningful fill, so
 * `fill`/`fillTo` are forced to `null` for them; box shapes default to
 * {@link DEFAULT_SHAPE_FILL}. A gradient is active only when BOTH a base fill
 * and a distinct second stop resolve (linear shapes never gradient-fill).
 */
/** Opacity bounds (a PERCENTAGE; 100 = fully opaque). */
export const SHAPE_OPACITY_MIN = 0;
export const SHAPE_OPACITY_MAX = 100;
export const DEFAULT_SHAPE_OPACITY = 100;

/**
 * Clamp a persisted shape opacity percentage to [0, 100] (rounded), defaulting
 * to fully opaque (100) for a missing / non-finite value. Mirrors the image
 * opacity normalizer so shapes + pictures share one transparency model.
 */
export function normalizeShapeOpacity(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_SHAPE_OPACITY;
  }
  const rounded = Math.round(value);
  if (rounded < SHAPE_OPACITY_MIN) return SHAPE_OPACITY_MIN;
  return rounded > SHAPE_OPACITY_MAX ? SHAPE_OPACITY_MAX : rounded;
}

/** A shape's fill Transparency: a whole percentage 1..100, else null (opaque). */
export function shapeFillTransparency(value: unknown): number | null {
  return typeof value === "number" && Number.isInteger(value) && value > 0 && value <= 100
    ? value
    : null;
}

export function resolveShapeStyle(input: {
  shape?: unknown;
  fill?: unknown;
  fillTransparency?: unknown;
  fillTo?: unknown;
  fillGradient?: unknown;
  shapeShadow?: unknown;
  gradientAngle?: unknown;
  fillImage?: unknown;
  fillPattern?: unknown;
  stroke?: unknown;
  outlineTransparency?: unknown;
  strokeTo?: unknown;
  strokeGradient?: unknown;
  strokeGradientAngle?: unknown;
  strokeWidth?: unknown;
  strokeStyle?: unknown;
  strokeAlign?: unknown;
  outlineDash?: unknown;
  outlineCap?: unknown;
  outlineJoin?: unknown;
  shadow?: unknown;
  opacity?: unknown;
  blendMode?: unknown;
  points?: unknown;
  customPath?: unknown;
  /** An imported OPEN freeform that is painted inside (`openPathFilled`). */
  openPathFilled?: unknown;
}): ResolvedShapeStyle {
  const kind = normalizeShapeKind(input.shape);
  // A merged shape (Merge Shapes → Intersect) carries its own custom polygon
  // outline; when present + valid it overrides the kind's preset polygon.
  // An ink stroke is an open line: two points are a line (and a tap's dot),
  // and it may hold as many as the ink envelope (an imported path's points).
  const customPoints = normalizeCustomPolygonPoints(
    input.points,
    kind === "ink" ? { min: 2, max: PRESENTATION_INK_MAX_POINTS } : undefined,
  );
  const customPath = normalizePresentationCustomPath(input.customPath);
  const customPathD = customPath
    ? presentationCustomPathToSvgPath(customPath)
    : null;
  const isLinear =
    isPresentationConnectorShapeKind(kind) ||
    kind === "ink" ||
    (kind === "custom-geometry" && customPath?.closed === false);
  // DrawingML fills an OPEN path too, closing it for the fill; an imported
  
  const openFilled =
    kind === "custom-geometry" && customPath?.closed === false && input.openPathFilled === true;
  const hasFill =
    (!isLinear || openFilled) &&
    (!isLineCalloutShapeKind(kind) || input.fill !== undefined || input.fillImage !== undefined || input.fillPattern !== undefined) &&
    !isOpenBasicShapeKind(kind);
  const stroke = normalizeShapeColor(input.stroke, DEFAULT_SHAPE_STROKE);
  const strokeTo =
    stroke === "none" ? null : normalizeShapeColor(input.strokeTo, null);
  const opaqueFill = hasFill
    ? normalizeShapeColor(input.fill, DEFAULT_SHAPE_FILL)
    : null;
  const fillTo =
    !hasFill || opaqueFill === null ? null : normalizeShapeColor(input.fillTo, null);
  // A translucent SOLID fill draws as `#rrggbbaa`, which both the box path's
  // CSS background and the SVG path's `fill` paint as the colour at that
  // alpha; a gradient, picture or pattern keeps its own paint.
  const transparency = shapeFillTransparency(input.fillTransparency);
  const fill =
    opaqueFill && opaqueFill !== "none" && transparency && fillTo === null && !input.fillGradient
      ? `${opaqueFill}${Math.round(((100 - transparency) / 100) * 255)
          .toString(16)
          .padStart(2, "0")}`
      : opaqueFill;
  // A picture fill is meaningful only on a box / polygon shape (never a line).
  const fillImage = hasFill ? normalizeShapeFillImage(input.fillImage) : null;
  const fillPattern = hasFill ? officeFillPatternOf(input.fillPattern) : null;
  return {
    kind,
    fill,
    fillTo,
    ...(input.fillGradient&&hasFill?{fillGradient:normalizeOfficeShapeGradient(input.fillGradient)}:{}),
    ...(input.shapeShadow?{shapeShadow:normalizeOfficeShapeShadow(input.shapeShadow)}:{}),
    gradientAngle: normalizeGradientAngle(input.gradientAngle),
    fillImage,
    fillPattern,
    // A solid line's Transparency (Format Shape > Line > Transparency, the
    // line colour's `a:alpha`) draws as `#rrggbbaa`, as the fill's does; a
    // gradient line keeps its stops' own.
    stroke:
      stroke !== "none" && strokeTo === null && !input.strokeGradient && shapeFillTransparency(input.outlineTransparency)
        ? `${stroke}${Math.round(((100 - shapeFillTransparency(input.outlineTransparency)!) / 100) * 255)
            .toString(16)
            .padStart(2, "0")}`
        : stroke,
    strokeTo,
    ...(input.strokeGradient && stroke !== "none"
      ? { strokeGradient: normalizeOfficeShapeGradient(input.strokeGradient) }
      : {}),
    strokeGradientAngle: normalizeGradientAngle(input.strokeGradientAngle),
    strokeWidth: normalizeStrokeWidth(input.strokeWidth),
    strokeStyle: normalizeStrokeStyle(input.strokeStyle),
    strokeAlign: officeStrokeAlign(input.strokeAlign),
    outlineDash: officeStrokeDash(normalizeStrokeStyle(input.strokeStyle), input.outlineDash),
    outlineCap: officeStrokeDetail({ outlineCap: input.outlineCap }).outlineCap ?? null,
    outlineJoin: officeStrokeDetail({ outlineJoin: input.outlineJoin }).outlineJoin ?? null,
    shadow: input.shadow === true,
    opacity: normalizeShapeOpacity(input.opacity),
    blendMode: normalizeShapeBlendMode(input.blendMode),
    isLinear,
    polygonPoints: customPath ? null : (customPoints ?? shapePolygonPoints(kind)),
    // A custom merged outline overrides the preset curve too.
    pathD: customPathD ?? (customPoints ? null : shapePathD(kind)),
    customPath,
    multiParts: customPath || customPoints ? null : shapeMultiParts(kind),
  };
}

/**
 * The `background` value for a box shape: a solid colour, or a linear-gradient
 * `fill → fillTo` at the resolved angle when a gradient is active. Returns
 * `"transparent"` for a fill-less shape. Shared so the SVG path can derive the
 * same two-stop vector.
 *
 * `box` is the shape's drawn size. A DrawingML gradient with `a:lin@scaled="1"`
 * — which is what this product WRITES, and what the corpus's gradients carry —
 * states its angle in the shape's own normalised space, so it comes out
 * aspect-scaled: 45° on a 16:3 box is drawn at 79.380° by real PowerPoint
 * (§32.2). Without the box the literal angle is all there is.
 */
export function shapeFillBackground(
  style: ResolvedShapeStyle,
  box?: { width?: number; height?: number },
): string {
  if (style.fill === null) return "transparent";
  if (style.fillPattern) return officeFillPatternCss(style.fillPattern);
  // A rectangular path gradient: rectangular contours over the box, as
  // PowerPoint prints a shape's (`officeRectPathGradientBackground`; look-fit
  // truth-probes/lookfit-rect-shape-probe.pptx), not the ellipse CSS's radial one is.
  if (style.fillGradient?.path?.kind === "rect")
    return officeRectPathGradientBackground(
      officeGradientDrawnStops(style.fillGradient.stops),
      style.fillGradient.path.fillToRect,
      style.fillGradient.tileRect,
    );
  if (style.fillGradient)
    return officeGradientCss(style.fillGradient, box?.width, box?.height);
  if (style.fillTo === null) return style.fill;
  // Two colours end to end: drawn as PowerPoint draws them (`officeGradientDrawnStops`).
  const stops = officeGradientDrawnStops([
    { position: 0, color: style.fill },
    { position: 1, color: style.fillTo },
  ]);
  return `linear-gradient(${style.gradientAngle}deg, ${stops.map((stop) => `${stop.color} ${stop.position * 100}%`).join(", ")})`;
}

/**
 * The CSS angle the shape's gradient is actually drawn at inside `box`.
 * The SVG path needs the same number the box path gets from
 * `shapeFillBackground`, so both come from here.
 */
export function shapeFillGradientAngle(
  style: ResolvedShapeStyle,
  box?: { width?: number; height?: number },
): number {
  return style.fillGradient
    ? officeGradientAngleForBox(style.fillGradient, box?.width, box?.height)
    : style.gradientAngle;
}

/**
 * The same number for the OUTLINE. An `a:lin@scaled="1"` outline gradient is
 * aspect-scaled exactly as a fill's is (§32.2), so the two cannot be allowed to
 * ask different questions -- they ask this one.
 */
export function shapeStrokeGradientAngle(
  style: ResolvedShapeStyle,
  box?: { width?: number; height?: number },
): number {
  return style.strokeGradient
    ? officeGradientAngleForBox(style.strokeGradient, box?.width, box?.height)
    : style.strokeGradientAngle;
}

/**
 * The stops an outline gradient draws: the whole model when the block carries
 * one, else the two ends every older block stores.
 */
export function shapeStrokeGradientStops(
  style: ResolvedShapeStyle,
): readonly { position: number; color: string; alpha?: number }[] {
  // Drawn as PowerPoint draws a fill's (`officeGradientDrawnStops`).
  if (style.strokeGradient) return officeGradientDrawnStops(style.strokeGradient.stops);
  return officeGradientDrawnStops([
    { position: 0, color: style.stroke },
    { position: 1, color: style.strokeTo ?? style.stroke },
  ]);
}
















export function gradientVector(
  angleDeg: number,
  width?: number,
  height?: number,
): {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
} {
  // CSS: 0deg -> upward, clockwise; screen space has y down.
  const rad = (angleDeg * Math.PI) / 180;
  const ux = Math.sin(rad);
  const uy = -Math.cos(rad);
  const w = Number.isFinite(width) && width! > 0 ? width! : 1;
  const h = Number.isFinite(height) && height! > 0 ? height! : 1;
  // The pixel-space isolines (direction (-uy, ux)) map to (-uy / w, ux / h) in
  // box units; the box-unit gradient direction is perpendicular to that.
  let gx = ux / h;
  let gy = uy / w;
  const length = Math.hypot(gx, gy) || 1;
  gx /= length;
  gy /= length;
  // The unit square's extent along it: the far corners land on 0 and 1.
  const extent = Math.abs(gx) + Math.abs(gy);
  const x1 = 0.5 - (gx * extent) / 2;
  const y1 = 0.5 - (gy * extent) / 2;
  const x2 = 0.5 + (gx * extent) / 2;
  const y2 = 0.5 + (gy * extent) / 2;
  return { x1, y1, x2, y2 };
}








export function strokeDashArray(style: ResolvedShapeStyle): string | undefined {
  const width = Math.max(1, style.strokeWidth);
  // Round Dot is `sysDot` with a round cap: the round dots of the dotted
  // family. Every other exact dash is its preset in line widths.
  if (style.outlineDash && !(style.outlineDash === "sysDot" && style.outlineCap === "round")) {
    // A round or square cap the line states is counted inside each dash, as
    // Office counts it (`officeDashLengthsWithCaps`); SVG paints it outside.
    return officeStrokeDashArray(
      style.outlineDash,
      width,
      style.outlineCap === "round" || style.outlineCap === "square",
    );
  }
  if (style.strokeStyle === "dotted" && style.outlineCap && style.outlineCap !== "round") {
    // The family's own `dot` preset under a flat or square cap the line
    // states: dots one width long every four widths (ECMA-376 `dot` 1-3).
    // The family's round dot is a dash of no length, which such a cap
    
    // with `cap="flat"` was missing).
    const length = (units: number) => Math.round(units * width * 100) / 100;
    return `${length(1)} ${length(3)}`;
  }
  return officeLineFamilyDash(style.strokeStyle, width)?.dashArray;
}

/** The cap an outline's dashes are drawn with: the line's own cap when it
 * states one, else round dots and flat dashes, or `undefined` for a solid
 * line. */
export function strokeDashLinecap(style: ResolvedShapeStyle): "round" | "butt" | "square" | undefined {
  if (style.outlineCap) return style.outlineCap === "flat" ? "butt" : style.outlineCap;
  // An exact dash stating no cap is drawn with flat ends, `sysDot` too: Round
  // Dot is `sysDot` WITH `cap="rnd"`, as PowerPoint writes it. Taken as round
  
  // flat) drew capsules two widths long with hardly a gap, PowerPoint square
  // dots one width on, one off.
  if (style.outlineDash) return "butt";
  return officeLineFamilyDash(style.strokeStyle, Math.max(1, style.strokeWidth))?.linecap;
}

/** A soft drop-shadow `box-shadow` / SVG-equivalent used when `shadow` is set. */
export const SHAPE_SHADOW_CSS = "0 4px 10px rgba(15, 23, 42, 0.28)";

/**
 * The CSS `background-*` properties for a PICTURE fill at the given fit:
 * `cover` / `contain` map to `backgroundSize` with no-repeat; `tile` repeats the
 * image at its natural size. The url is wrapped in `url("…")` (the value is
 * already validated by {@link normalizeShapeFillImage}). Shared so the box path
 * and any other CSS consumer size a picture fill identically.
 */
export function shapeFillImageCss(image: ShapeFillImage): CSSProperties {
  const base: CSSProperties = {
    backgroundImage: `url("${image.url}")`,
    backgroundPosition: "center",
  };
  if (image.fit === "tile") {
    base.backgroundRepeat = "repeat";
    base.backgroundSize = "auto";
  } else {
    base.backgroundRepeat = "no-repeat";
    base.backgroundSize = image.fit === "contain" ? "contain" : "cover";
  }
  return base;
}

/**
 * CSS for a box-shape (rectangle / ellipse) container. Lines/arrows render via
 * SVG instead and should not call this. Returns background (a PICTURE fill when
 * set, else solid / gradient) + outline (weight / dash style / colour) + radius
 * (50% for ellipses) + optional drop shadow. Side-effect free; safe in both the
 * editor and present overlay.
 */
export function boxShapeCss(
  style: ResolvedShapeStyle,
  box?: { width?: number; height?: number },
): CSSProperties {
  const css: CSSProperties = {
    border:
      style.strokeWidth > 0
        ? `${style.strokeWidth}px ${style.strokeStyle} ${style.stroke}`
        : "none",
    borderRadius:
      style.kind === "ellipse" ? "50%" : style.kind === "rounded-rectangle" ? 16 : 4,
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
  };
  // A picture fill (PowerPoint "Picture or texture fill") takes precedence over
  // the solid / gradient colour; otherwise the historical colour background.
  if (style.fillImage) {
    Object.assign(css, shapeFillImageCss(style.fillImage));
  } else {
    css.background = shapeFillBackground(style, box);
  }
  if (style.shadow) css.boxShadow = SHAPE_SHADOW_CSS;
  if(style.shapeShadow)css.boxShadow=officeShadowCss(style.shapeShadow);
  if (style.opacity < DEFAULT_SHAPE_OPACITY) css.opacity = style.opacity / 100;
  // Compositing: only emit `mix-blend-mode` for a non-default mode so an
  // ordinary (normal) shape keeps no blend property and renders byte-for-byte.
  if (style.blendMode !== DEFAULT_SHAPE_BLEND_MODE) {
    css.mixBlendMode = style.blendMode;
  }
  return css;
}
