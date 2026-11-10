/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One arrow picker's model, shared by Docs, the spreadsheet and the deck.
 *
 * Office keeps this on the outline: Shape Format ▸ Shape Outline ▸ Arrows is a
 * gallery of begin/end combinations, and "More Arrows…" opens Format Shape ▸
 * Line, where Begin/End Arrow **type** offers the six `a:ST_LineEndType`
 * members and Begin/End Arrow **size** the nine width × length pairs
 * (`a:ST_LineEndWidth` × `a:ST_LineEndLength`). Microsoft's own reference for
 * the object model says the same: `MsoArrowheadStyle` has None, Triangle,
 * Open, Stealth, Diamond and Oval, and `Excel.Line` takes an arrowhead
 * `Short|Medium|Long` length and a `Narrow|Medium|Wide` width — on a LINE.
 *
 * Nothing here is a host's persistence: an adapter turns one change into the
 * fields its own editor stores (`officeShapeLineEndsPatch` for a document or a
 * sheet shape, `presentationConnectorEndsPatch` for a slide block).
 */

import {
  isOfficeShapePreset,
  officeShapeDefinition,
  officeShapePresetFromDrawingMl,
  type OfficeShapePreset,
} from "./office-shape-catalog";
import { nativeLineEndPlacements } from "./native-drawing-geometry";
import {
  normalizePresentationConnectorSpec,
  PRESENTATION_CONNECTOR_END_SIZES,
  PRESENTATION_CONNECTOR_ENDS,
  type PresentationConnectorEnd,
  type PresentationConnectorEndSize,
  type ShapeKind,
} from "./presentation-shapes";

export type OfficeLineEndType = PresentationConnectorEnd;
export type OfficeLineEndSize = PresentationConnectorEndSize;

/** What a shape's two ends currently are, as a picker reads them. */
export interface OfficeLineEndsState {
  readonly headEnd: OfficeLineEndType;
  readonly tailEnd: OfficeLineEndType;
  readonly headEndWidth?: OfficeLineEndSize;
  readonly headEndLength?: OfficeLineEndSize;
  readonly tailEndWidth?: OfficeLineEndSize;
  readonly tailEndLength?: OfficeLineEndSize;
}

/** One Arrow size cell: a width and a length, the pair Office writes together. */
export interface OfficeLineEndSizeChoice {
  readonly w: OfficeLineEndSize;
  readonly len: OfficeLineEndSize;
}

/** A picked change; an absent key leaves that half of the state alone. */
export interface OfficeLineEndsChange {
  readonly headEnd?: OfficeLineEndType;
  readonly tailEnd?: OfficeLineEndType;
  /** `null` clears `@w`/`@len` — back to the absent size Office draws medium for. */
  readonly headEndSize?: OfficeLineEndSizeChoice | null;
  readonly tailEndSize?: OfficeLineEndSizeChoice | null;
}

export const OFFICE_LINE_END_TYPES: readonly OfficeLineEndType[] =
  PRESENTATION_CONNECTOR_ENDS;
export const OFFICE_LINE_END_SIZES: readonly OfficeLineEndSize[] =
  PRESENTATION_CONNECTOR_END_SIZES;

/**
 * The nine Arrow size cells, in Office's own order: length runs across
 * (Short, Medium, Long) and width down (Narrow, Medium, Wide).
 */
export const OFFICE_LINE_END_SIZE_CHOICES: readonly OfficeLineEndSizeChoice[] =
  Object.freeze(
    OFFICE_LINE_END_SIZES.flatMap((w) =>
      OFFICE_LINE_END_SIZES.map((len) => Object.freeze({ w, len })),
    ),
  );

/** One gallery tile: the pair of decorations it sets, and nothing else. */
export interface OfficeLineEndGalleryItem {
  readonly id: string;
  readonly headEnd: OfficeLineEndType;
  readonly tailEnd: OfficeLineEndType;
}

/**
 * The Arrows gallery.
 *
 * Microsoft's documentation names only the first tile ("Arrow Style 1 (no
 * arrowheads)") and never lists the rest, so this is OUR set of eleven
 * combinations over the six documented decorations — not a copy of Office's
 * gallery, which we have no specification for. Every combination it cannot
 * show is still reachable through More Arrows.
 */
export const OFFICE_LINE_END_GALLERY: readonly OfficeLineEndGalleryItem[] =
  Object.freeze(
    (
      [
        ["none", "none"],
        ["none", "triangle"],
        ["triangle", "none"],
        ["triangle", "triangle"],
        ["none", "arrow"],
        ["arrow", "none"],
        ["arrow", "arrow"],
        ["none", "stealth"],
        ["stealth", "stealth"],
        ["oval", "triangle"],
        ["diamond", "triangle"],
      ] as const
    ).map(([headEnd, tailEnd]) =>
      Object.freeze({ id: `${headEnd}-${tailEnd}`, headEnd, tailEnd }),
    ),
  );

function isEndType(value: unknown): value is OfficeLineEndType {
  return (
    typeof value === "string" &&
    (OFFICE_LINE_END_TYPES as readonly string[]).includes(value)
  );
}

function isEndSize(value: unknown): value is OfficeLineEndSize {
  return (
    typeof value === "string" &&
    (OFFICE_LINE_END_SIZES as readonly string[]).includes(value)
  );
}

/** Read the six fields off any host's shape record, holding each to its vocabulary. */
export function readOfficeLineEnds(
  record: Readonly<Record<string, unknown>> | null | undefined,
  implied?: { readonly headEnd?: unknown; readonly tailEnd?: unknown },
): OfficeLineEndsState {
  const value = record ?? {};
  const type = (key: "headEnd" | "tailEnd"): OfficeLineEndType =>
    isEndType(value[key])
      ? value[key]
      : isEndType(implied?.[key])
        ? implied[key]
        : "none";
  const size = (key: keyof OfficeLineEndsState): OfficeLineEndSize | undefined =>
    isEndSize(value[key]) ? value[key] : undefined;
  const headEnd = type("headEnd");
  const tailEnd = type("tailEnd");
  return {
    headEnd,
    tailEnd,
    // `@w`/`@len` size a decoration; an end that carries none has nothing to
    // size, and the save gate refuses the pair outright.
    ...(headEnd === "none"
      ? {}
      : {
          ...(size("headEndWidth") ? { headEndWidth: size("headEndWidth") } : {}),
          ...(size("headEndLength") ? { headEndLength: size("headEndLength") } : {}),
        }),
    ...(tailEnd === "none"
      ? {}
      : {
          ...(size("tailEndWidth") ? { tailEndWidth: size("tailEndWidth") } : {}),
          ...(size("tailEndLength") ? { tailEndLength: size("tailEndLength") } : {}),
        }),
  };
}

/**
 * Apply one pick.
 *
 * An absent size STAYS absent — `CT_LineEndProperties` declares no default for
 * `@w` or `@len`, and Office itself writes `<a:tailEnd type="triangle"/>` bare
 * — so picking a type never invents a size, and picking a size writes the
 * width and the length together, as the Arrow size grid does.
 */
export function applyOfficeLineEnds(
  current: OfficeLineEndsState,
  change: OfficeLineEndsChange,
): OfficeLineEndsState {
  const next: {
    -readonly [K in keyof OfficeLineEndsState]: OfficeLineEndsState[K];
  } = { ...current };
  for (const side of ["head", "tail"] as const) {
    const typeKey = `${side}End` as const;
    const widthKey = `${side}EndWidth` as const;
    const lengthKey = `${side}EndLength` as const;
    const sizeKey = `${side}EndSize` as const;
    const type = change[typeKey];
    if (type !== undefined) next[typeKey] = type;
    if (Object.hasOwn(change, sizeKey)) {
      const size = change[sizeKey];
      if (size === null) {
        delete next[widthKey];
        delete next[lengthKey];
      } else if (size) {
        next[widthKey] = size.w;
        next[lengthKey] = size.len;
      }
    }
    if (next[typeKey] === "none") {
      delete next[widthKey];
      delete next[lengthKey];
    }
  }
  return next;
}

export function officeLineEndsEqual(
  a: OfficeLineEndsState,
  b: OfficeLineEndsState,
): boolean {
  return (
    a.headEnd === b.headEnd &&
    a.tailEnd === b.tailEnd &&
    a.headEndWidth === b.headEndWidth &&
    a.headEndLength === b.headEndLength &&
    a.tailEndWidth === b.tailEndWidth &&
    a.tailEndLength === b.tailEndLength
  );
}

/** The gallery tile a state matches, by its two decorations alone. */
export function officeLineEndGalleryId(
  state: OfficeLineEndsState,
): string | null {
  return (
    OFFICE_LINE_END_GALLERY.find(
      (item) => item.headEnd === state.headEnd && item.tailEnd === state.tailEnd,
    )?.id ?? null
  );
}

/**
 * Whether a geometry HAS ends to decorate, by the same rule the canvases draw
 * by: Office greys its arrow settings out on a shape whose outline is closed
 * (a Block Arc, a rectangle), because an arrowhead needs a beginning and an
 * end to sit on.
 */
export function officeGeometryTakesLineEnds(
  paths: readonly { readonly d: string; readonly stroke?: boolean }[],
): boolean {
  return nativeLineEndPlacements(paths).length > 0;
}

/**
 * The fields a DOCUMENT or SHEET shape stores for one pick, with the gallery
 * preset the result is: a connector's preset names WHICH ends are decorated,
 * so turning one on or off is a different preset — the very one this shape's
 * own exported file would import back as (`officeShapePresetFromDrawingMl`),
 * which keeps edit and re-import the same operation.
 *
 * Both types are written EXPLICITLY, `none` included. A document's patchback
 * inherits an absent end from the authenticated original and a workbook's
 * keeps whatever the package already holds, so silence is not how a user takes
 * an arrowhead off (SHAPE-IMPORT-AUDIT.md §7.10).
 */
export function officeShapeLineEndsPatch(
  preset: OfficeShapePreset,
  ends: OfficeLineEndsState,
): {
  readonly preset: OfficeShapePreset;
  readonly headEnd: OfficeLineEndType;
  readonly tailEnd: OfficeLineEndType;
  readonly headEndWidth?: OfficeLineEndSize;
  readonly headEndLength?: OfficeLineEndSize;
  readonly tailEndWidth?: OfficeLineEndSize;
  readonly tailEndLength?: OfficeLineEndSize;
} | null {
  if (!isOfficeShapePreset(preset)) return null;
  const definition = officeShapeDefinition(preset);
  const nextPreset = definition.lineLike
    ? officeShapePresetFromDrawingMl(
        definition.drawingPreset,
        ends.tailEnd,
        ends.headEnd,
        // A document or sheet shape's preset is all it knows of the route.
        { keepRoute: true },
      )
    : preset;
  if (!nextPreset) return null;
  return { preset: nextPreset, ...ends };
}

/** The same for a SLIDE block, whose connector identity is a shape kind. */
export function presentationConnectorEndsPatch(
  block: Readonly<Record<string, unknown>>,
  ends: OfficeLineEndsState,
): {
  readonly shape: ShapeKind;
  readonly route: string | undefined;
  readonly connectorPreset: string;
  readonly headEnd: OfficeLineEndType;
  readonly tailEnd: OfficeLineEndType;
  readonly headEndWidth?: OfficeLineEndSize;
  readonly headEndLength?: OfficeLineEndSize;
  readonly tailEndWidth?: OfficeLineEndSize;
  readonly tailEndLength?: OfficeLineEndSize;
} | null {
  const spec = normalizePresentationConnectorSpec(block);
  if (!spec) return null;
  const preset = officeShapePresetFromDrawingMl(
    spec.connectorPreset,
    ends.tailEnd,
    ends.headEnd,
  );
  const definition =
    preset && isOfficeShapePreset(preset) ? officeShapeDefinition(preset) : null;
  if (!definition?.lineLike) return null;
  const next = {
    shape: definition.previewKind,
    // A connector's route is its preset's; `straight` persists as absent.
    route: spec.route === "straight" ? undefined : spec.route,
    connectorPreset: spec.connectorPreset,
    ...ends,
  };
  // The pick has to survive the gate both sides of the save share. It always
  // does — the 17 × 6 × 6 matrix proves it — but a block that was already
  // contradictory must not become a plausible connector here.
  return normalizePresentationConnectorSpec({ ...next, shape: next.shape })
    ? next
    : null;
}

/** PowerPoint's Connector Type ▸, in its order. */
export const PRESENTATION_CONNECTOR_TYPES = ["straight", "elbow", "curved"] as const;
export type PresentationConnectorType = (typeof PRESENTATION_CONNECTOR_TYPES)[number];

const CONNECTOR_TYPE_PRESET: Record<PresentationConnectorType, string> = {
  straight: "straightConnector1",
  elbow: "bentConnector3",
  curved: "curvedConnector3",
};

/**
 * Connector Type ▸ Straight / Elbow / Curved on a slide connector: the same
 * ends and sizes on the new route's preset (the kind follows the preset, as
 * an Arrows pick re-kinds it). Null when the connector is already that type,
 * or the block is no connector.
 */
export function presentationConnectorTypePatch(
  block: Readonly<Record<string, unknown>>,
  type: PresentationConnectorType,
): ReturnType<typeof presentationConnectorEndsPatch> {
  const spec = normalizePresentationConnectorSpec(block);
  if (!spec || spec.route === type) return null;
  const connectorPreset = CONNECTOR_TYPE_PRESET[type];
  const preset = officeShapePresetFromDrawingMl(connectorPreset, spec.tailEnd, spec.headEnd);
  const definition =
    preset && isOfficeShapePreset(preset) ? officeShapeDefinition(preset) : null;
  if (!definition?.lineLike) return null;
  const ends = readOfficeLineEnds(block, spec);
  const next = {
    shape: definition.previewKind,
    route: type === "straight" ? undefined : type,
    connectorPreset,
    ...ends,
  };
  return normalizePresentationConnectorSpec(next) ? next : null;
}

/** The Connector Type ▸ entry a slide connector shows checked. */
export function presentationConnectorType(
  block: Readonly<Record<string, unknown>>,
): PresentationConnectorType | null {
  return normalizePresentationConnectorSpec(block)?.route ?? null;
}
