/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide size / aspect ratio model for the presentation editor (PowerPoint / WPS
 * "Slide Size" parity).
 *
 * A deck's slide dimensions live on each slide as optional `widthEmu` /
 * `heightEmu` (English Metric Units — the OOXML / PowerPoint native unit, with
 * 914400 EMU per inch). Historically those fields were only READ at render time
 * (`slideBounds`) and never set, so every deck rendered at the built-in default.
 * This module supplies the pure conversion + preset + clamp logic the editor's
 * "Slide Size" setter uses to WRITE coherent `widthEmu`/`heightEmu` onto every
 * slide (slide size is deck-wide in PowerPoint, never per-slide).
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit tested
 * in isolation and shared by the editor + any serialize-edge guard.
 */

import { PRESENTATION_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";
import { MAX_PRESENTATION_SLIDES } from "./presentation/slide-limits";

/** EMU per inch — the OOXML constant (1 inch = 914400 EMU). */
import { EMU_PER_INCH } from "./office-units";
export { EMU_PER_INCH };

/**
 * Inclusive bounds for a scratch-deck custom slide dimension, expressed in
 * inches. PowerPoint allows 1"..56". Native imported decks are additionally
 * gated by the stricter authenticated OOXML/model envelope below.
 */
export const MIN_SLIDE_INCHES = 1;
export const MAX_SLIDE_INCHES = 56;

/** Convert a length in inches to EMU, rounded to an integer EMU. */
export function inchesToEmu(inches: number): number {
  return Math.round(inches * EMU_PER_INCH);
}

/** Convert a length in EMU to inches (may be fractional). */
export function emuToInches(emu: number): number {
  return emu / EMU_PER_INCH;
}

/** A concrete slide size in EMU (the unit persisted on each slide). */
export interface SlideSizeEmu {
  widthEmu: number;
  heightEmu: number;
}

export const PPTX_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION =
  "pptx-slide-size-cascade-v1" as const;
export const PPTX_SCOPED_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION =
  "pptx-slide-size-cascade-v2" as const;

export type NativeSlideSizePartKind = "slide" | "layout" | "master";

export interface NativeSlideSizeObjectAuthority {
  readonly shapeId: string;
  readonly tag: "sp" | "pic" | "cxnSp" | "grpSp" | "graphicFrame";
  readonly xfrmPath: string;
  readonly xfrmFingerprint: string;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly tableGrid?: {
    readonly columns: readonly number[];
    readonly rows: readonly number[];
  };
}

export interface NativeSlideSizePartAuthority {
  readonly kind: NativeSlideSizePartKind;
  readonly ownerId: string;
  readonly part: string;
  readonly fingerprint: string;
  readonly objects: readonly NativeSlideSizeObjectAuthority[];
}

/** Immutable, package-authenticated custody for native PPTX Slide Size. */
export interface NativeSlideSizeAuthority {
  readonly schemaVersion:
    | typeof PPTX_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION
    | typeof PPTX_SCOPED_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION;
  readonly scopeSlideId?: string;
  readonly sourceSlideIds?: readonly string[];
  readonly disposition: "native";
  readonly packageFingerprint: string;
  readonly presentationPart: string;
  readonly presentationFingerprint: string;
  readonly slideSize: Readonly<
    SlideSizeEmu & {
      type: string | null;
      typePresent: boolean;
      fingerprint: string;
    }
  >;
  readonly parts: readonly NativeSlideSizePartAuthority[];
}

export interface PresentationDesignAuthority {
  binding: {
    schemaVersion: "pptx-design-authority-v1";
    pageId: string;
    contentVersion: number;
    contentDigest: string;
    resourceId: string;
    resourceVersionId: string;
    packageFingerprint: string;
  };
  authority: unknown;
  slides: {
    slideId: string;
    blocks: {
      id: string;
      shapeId: string | null;
      geometry: unknown;
      tableGrid?: unknown;
    }[];
  }[];
}

/** Runtime-only adoption: the caller commits this copy only for an explicit resize. */
export function adoptNativeSlideSizeAuthority<
  S extends SlideSizeAuthoritySlide,
>(slides: readonly S[], response: PresentationDesignAuthority): S[] | null {
  const authority = parseNativeSlideSizeAuthority(response.authority);
  if (
    !authority ||
    authority.packageFingerprint !== response.binding.packageFingerprint ||
    !Array.isArray(response.slides)
  )
    return null;
  const originals = new Map(
    response.slides.map((slide) => [slide.slideId, slide]),
  );
  if (originals.size !== response.slides.length) return null;
  const order = response.slides.map((slide) => slide.slideId);
  const liveOrder = slides.map((slide) =>
    order.indexOf(
      String(
        (slide.source as Record<string, unknown> | undefined)?.slideId ?? "",
      ),
    ),
  );
  if (
    liveOrder.some(
      (position, index) =>
        position < 0 || (index > 0 && position <= liveOrder[index - 1]),
    )
  )
    return null;
  const next = slides.map((slide) => {
    const source = slide.source as Record<string, unknown> | undefined;
    const original = originals.get(String(source?.slideId ?? ""));
    if (!source || !original || !Array.isArray(slide.blocks)) return null;
    const blocks = slide.blocks.map((block: Record<string, unknown>) => {
      const blockSource = (block.source ?? {}) as Record<string, unknown>;
      const matches = original.blocks.filter(
        (item) =>
          String(item.shapeId ?? "") === String(blockSource.shapeId ?? "") &&
          item.id === block.id,
      );
      if (!blockSource.shapeId && block.type === "chart") return block;
      if (matches.length !== 1 || !matches[0].geometry) return null;
      return {
        ...block,
        source: {
          ...blockSource,
          slideSizeOriginalGeometry: matches[0].geometry,
          ...(matches[0].tableGrid
            ? { slideSizeOriginalTableGrid: matches[0].tableGrid }
            : {}),
        },
      };
    });
    if (blocks.some((block: unknown) => block === null)) return null;
    return {
      ...slide,
      blocks,
      source: {
        ...source,
        slideSizeAuthority: compactNativeSlideSizeAuthority(
          authority,
          original.slideId,
        ),
        designAuthorityBinding: response.binding,
      },
    };
  });
  if (
    next.some((slide) => slide === null) ||
    !nativeSlideSizeCanApply(next as S[])
  )
    return null;
  return next as S[];
}

type SlideSizeAuthoritySlide = {
  readonly widthEmu?: unknown;
  readonly heightEmu?: unknown;
  readonly blocks?: unknown;
  readonly source?: unknown;
};

const SHA256_RE = /^[0-9a-f]{64}$/;
const SHAPE_ID_RE = /^shp[1-9][0-9]*$/;
const OPC_PART_RE =
  /^(?!\/)(?!.*(?:^|\/)\.\.?(?:\/|$))(?!.*\\)[^/]+(?:\/[^/]+)*\.xml$/;
const XFRM_PATH_RE = /^[A-Za-z0-9_:\-\[\]@='".\/]{1,512}$/;
const SLIDE_SIZE_PART_KINDS = new Set<NativeSlideSizePartKind>([
  "slide",
  "layout",
  "master",
]);
const SLIDE_SIZE_OBJECT_TAGS = new Set<NativeSlideSizeObjectAuthority["tag"]>([
  "sp",
  "pic",
  "cxnSp",
  "grpSp",
  "graphicFrame",
]);
const MAX_SLIDE_SIZE_GEOMETRY_EMU = PRESENTATION_GEOMETRY_MAX_EMU;






const MAX_SLIDE_SIZE_OBJECTS = 4_096;

function boundedOffset(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= -MAX_SLIDE_SIZE_GEOMETRY_EMU &&
    value <= MAX_SLIDE_SIZE_GEOMETRY_EMU
  );
}

function boundedExtent(value: unknown): value is number {
  return positiveInteger(value) && value <= MAX_SLIDE_SIZE_GEOMETRY_EMU;
}

function exactRecord(
  value: unknown,
  keys: readonly string[],
): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const actual = Object.keys(record).sort();
  const expected = [...keys].sort();
  return actual.length === expected.length &&
    actual.every((key, index) => key === expected[index])
    ? record
    : null;
}

function exactOptionalRecord(
  value: unknown,
  requiredKeys: readonly string[],
  optionalKeys: readonly string[],
): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const allowed = new Set([...requiredKeys, ...optionalKeys]);
  return requiredKeys.every((key) => key in record) &&
    Object.keys(record).every((key) => allowed.has(key))
    ? record
    : null;
}

function positiveInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value > 0;
}

/**
 * Each receipt object parsed once: the editor checks every slide's receipt on
 * every render (`nativeSlideSizeCanApply`), and a slide it did not change
 * holds the same receipt object. The parsed receipt is shared and must not be
 * changed.
 */
const parsedNativeSlideSizeAuthorities = new WeakMap<
  object,
  NativeSlideSizeAuthority | null
>();

/** The deck's slide ids a scoped receipt lists, checked once per list: every
 * slide's receipt holds the same list, so checking each one was quadratic. */
const nativeSourceSlideIdSets = new WeakMap<object, ReadonlySet<string> | null>();

function nativeSourceSlideIdSet(value: unknown): ReadonlySet<string> | null {
  if (!Array.isArray(value)) return null;
  if (nativeSourceSlideIdSets.has(value)) return nativeSourceSlideIdSets.get(value) ?? null;
  const ids = new Set<string>(value as string[]);
  const valid =
    value.length >= 2 &&
    value.length <= MAX_PRESENTATION_SLIDES &&
    value.every((id) => typeof id === "string" && /^sld[1-9][0-9]*$/.test(id)) &&
    ids.size === value.length;
  nativeSourceSlideIdSets.set(value, valid ? ids : null);
  return valid ? ids : null;
}

/** Parse the exact server-issued native Slide Size custody receipt. */
export function parseNativeSlideSizeAuthority(
  value: unknown,
): NativeSlideSizeAuthority | null {
  if (!value || typeof value !== "object") return null;
  if (parsedNativeSlideSizeAuthorities.has(value)) {
    return parsedNativeSlideSizeAuthorities.get(value) ?? null;
  }
  const parsed = parseNativeSlideSizeAuthorityOnce(value);
  parsedNativeSlideSizeAuthorities.set(value, parsed);
  return parsed;
}

function parseNativeSlideSizeAuthorityOnce(
  value: unknown,
): NativeSlideSizeAuthority | null {
  const scoped = Boolean(
    value &&
    typeof value === "object" &&
    (value as Record<string, unknown>).schemaVersion ===
      PPTX_SCOPED_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION,
  );
  const record = exactRecord(value, [
    "schemaVersion",
    "disposition",
    "packageFingerprint",
    "presentationPart",
    "presentationFingerprint",
    "slideSize",
    "parts",
    ...(scoped ? ["scopeSlideId", "sourceSlideIds"] : []),
  ]);
  if (
    !record ||
    (!scoped &&
      record.schemaVersion !== PPTX_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION) ||
    record.disposition !== "native" ||
    typeof record.packageFingerprint !== "string" ||
    !SHA256_RE.test(record.packageFingerprint) ||
    typeof record.presentationPart !== "string" ||
    record.presentationPart !== "ppt/presentation.xml" ||
    typeof record.presentationFingerprint !== "string" ||
    !SHA256_RE.test(record.presentationFingerprint)
  ) {
    return null;
  }
  if (
    scoped &&
    (typeof record.scopeSlideId !== "string" ||
      !nativeSourceSlideIdSet(record.sourceSlideIds)?.has(record.scopeSlideId))
  )
    return null;
  const slideSize = exactOptionalRecord(
    record.slideSize,
    ["widthEmu", "heightEmu", "type", "typePresent", "fingerprint"],
    [],
  );
  if (
    !slideSize ||
    !boundedExtent(slideSize.widthEmu) ||
    !boundedExtent(slideSize.heightEmu) ||
    (slideSize.type !== null &&
      (typeof slideSize.type !== "string" || slideSize.type.length > 80)) ||
    typeof slideSize.typePresent !== "boolean" ||
    slideSize.typePresent !== (slideSize.type !== null) ||
    typeof slideSize.fingerprint !== "string" ||
    !SHA256_RE.test(slideSize.fingerprint)
  ) {
    return null;
  }
  if (
    !Array.isArray(record.parts) ||
    record.parts.length < 1 ||
    record.parts.length > 512
  ) {
    return null;
  }
  const seenParts = new Set<string>();
  let objectCount = 0;
  const parts: NativeSlideSizePartAuthority[] = [];
  for (const rawPart of record.parts) {
    const part = exactRecord(rawPart, [
      "kind",
      "ownerId",
      "part",
      "fingerprint",
      "objects",
    ]);
    if (
      !part ||
      typeof part.kind !== "string" ||
      !SLIDE_SIZE_PART_KINDS.has(part.kind as NativeSlideSizePartKind) ||
      typeof part.part !== "string" ||
      !OPC_PART_RE.test(part.part) ||
      typeof part.ownerId !== "string" ||
      part.ownerId.length < 1 ||
      part.ownerId.length > 128 ||
      seenParts.has(part.part) ||
      typeof part.fingerprint !== "string" ||
      !SHA256_RE.test(part.fingerprint) ||
      !Array.isArray(part.objects) ||
      part.objects.length > MAX_SLIDE_SIZE_OBJECTS
    ) {
      return null;
    }
    seenParts.add(part.part);
    objectCount += part.objects.length;
    if (objectCount > MAX_SLIDE_SIZE_OBJECTS) return null;
    const seenShapes = new Set<string>();
    const objects: NativeSlideSizeObjectAuthority[] = [];
    for (const rawObject of part.objects) {
      const object = exactOptionalRecord(
        rawObject,
        [
          "shapeId",
          "tag",
          "xfrmPath",
          "xfrmFingerprint",
          "left",
          "top",
          "width",
          "height",
        ],
        ["tableGrid"],
      );
      if (
        !object ||
        typeof object.shapeId !== "string" ||
        !SHAPE_ID_RE.test(object.shapeId) ||
        seenShapes.has(object.shapeId) ||
        typeof object.tag !== "string" ||
        !SLIDE_SIZE_OBJECT_TAGS.has(
          object.tag as NativeSlideSizeObjectAuthority["tag"],
        ) ||
        typeof object.xfrmPath !== "string" ||
        !XFRM_PATH_RE.test(object.xfrmPath) ||
        object.xfrmPath !==
          (object.tag === "graphicFrame"
            ? "xfrm"
            : object.tag === "grpSp"
              ? "grpSpPr/xfrm"
              : "spPr/xfrm") ||
        typeof object.xfrmFingerprint !== "string" ||
        !SHA256_RE.test(object.xfrmFingerprint) ||
        !boundedOffset(object.left) ||
        !boundedOffset(object.top) ||
        !(
          boundedExtent(object.width) ||
          (object.tag === "cxnSp" && object.width === 0)
        ) ||
        !(
          boundedExtent(object.height) ||
          (object.tag === "cxnSp" && object.height === 0)
        ) ||
        (object.width === 0 && object.height === 0)
      ) {
        return null;
      }
      const tableGrid =
        object.tableGrid === undefined
          ? undefined
          : parseNativeSlideSizeTableGrid(object.tableGrid);
      if (
        object.tableGrid !== undefined &&
        (!tableGrid || object.tag !== "graphicFrame")
      )
        return null;
      seenShapes.add(object.shapeId);
      objects.push({
        shapeId: object.shapeId,
        tag: object.tag as NativeSlideSizeObjectAuthority["tag"],
        xfrmPath: object.xfrmPath,
        xfrmFingerprint: object.xfrmFingerprint,
        left: object.left,
        top: object.top,
        width: object.width,
        height: object.height,
        ...(tableGrid ? { tableGrid } : {}),
      });
    }
    parts.push({
      kind: part.kind as NativeSlideSizePartKind,
      ownerId: part.ownerId,
      part: part.part,
      fingerprint: part.fingerprint,
      objects,
    });
  }
  if (
    scoped &&
    (parts.filter((part) => part.kind === "slide").length !== 1 ||
      parts.find((part) => part.kind === "slide")?.ownerId !==
        record.scopeSlideId)
  )
    return null;
  return {
    schemaVersion: scoped
      ? PPTX_SCOPED_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION
      : PPTX_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION,
    ...(scoped
      ? {
          scopeSlideId: record.scopeSlideId as string,
          sourceSlideIds: record.sourceSlideIds as string[],
        }
      : {}),
    disposition: "native",
    packageFingerprint: record.packageFingerprint,
    presentationPart: record.presentationPart,
    presentationFingerprint: record.presentationFingerprint,
    slideSize: {
      widthEmu: slideSize.widthEmu,
      heightEmu: slideSize.heightEmu,
      type: slideSize.type as string | null,
      typePresent: slideSize.typePresent,
      fingerprint: slideSize.fingerprint,
    },
    parts,
  };
}

/** Persist only this slide's geometry plus shared layout/master custody. */
export function compactNativeSlideSizeAuthority(
  authority: NativeSlideSizeAuthority,
  slideId: string,
): NativeSlideSizeAuthority | null {
  if (
    authority.schemaVersion === PPTX_SCOPED_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION
  ) {
    return authority.scopeSlideId === slideId ? authority : null;
  }
  const owners = authority.parts
    .filter((part) => part.kind === "slide")
    .map((part) => part.ownerId);
  if (!owners.includes(slideId) || new Set(owners).size !== owners.length)
    return null;
  if (owners.length === 1) return authority;
  return {
    ...authority,
    schemaVersion: PPTX_SCOPED_SLIDE_SIZE_AUTHORITY_SCHEMA_VERSION,
    scopeSlideId: slideId,
    sourceSlideIds: owners,
    parts: authority.parts.filter(
      (part) => part.kind !== "slide" || part.ownerId === slideId,
    ),
  };
}

function parseNativeSlideSizeTableGrid(
  value: unknown,
): NativeSlideSizeObjectAuthority["tableGrid"] | null {
  const grid = exactRecord(value, ["columns", "rows"]);
  if (
    !grid ||
    ![grid.columns, grid.rows].every(
      (axis) =>
        Array.isArray(axis) &&
        axis.length >= 1 &&
        axis.length <= 200 &&
        axis.every(boundedExtent),
    )
  )
    return null;
  return {
    columns: [...(grid.columns as number[])],
    rows: [...(grid.rows as number[])],
  };
}

/** What every slide's receipt must agree on (its deck part), once per receipt. */
const deckSlideSizeCustodies = new WeakMap<NativeSlideSizeAuthority, string>();

function deckSlideSizeCustody(receipt: NativeSlideSizeAuthority): string {
  let encoded = deckSlideSizeCustodies.get(receipt);
  if (encoded === undefined) {
    encoded = JSON.stringify(
      receipt.scopeSlideId
        ? {
            ...receipt,
            scopeSlideId: undefined,
            parts: receipt.parts.filter((part) => part.kind !== "slide"),
          }
        : receipt,
    );
    deckSlideSizeCustodies.set(receipt, encoded);
  }
  return encoded;
}

/** Every imported slide must carry consistent full or per-slide native custody. */
export function nativeSlideSizeCanApply(
  slides: readonly SlideSizeAuthoritySlide[],
): boolean {
  const imported = slides.some((slide) => {
    const source = slide.source;
    return Boolean(
      source &&
      typeof source === "object" &&
      !Array.isArray(source) &&
      typeof (source as Record<string, unknown>).slideId === "string",
    );
  });
  if (!imported) return true;
  if (slides.length === 0) return false;
  let canonical = "";
  let deckWidth: number | null = null;
  let deckHeight: number | null = null;
  let deckRevision: number | null = null;
  let deckStrategy: SlideSizeFitStrategy | null = null;
  for (const slide of slides) {
    const source =
      slide.source &&
      typeof slide.source === "object" &&
      !Array.isArray(slide.source)
        ? (slide.source as Record<string, unknown>)
        : null;
    if (!source || typeof source.slideId !== "string") return false;
    const receipt = parseNativeSlideSizeAuthority(source.slideSizeAuthority);
    if (
      !receipt ||
      (receipt.scopeSlideId && receipt.scopeSlideId !== source.slideId)
    )
      return false;
    const encoded = deckSlideSizeCustody(receipt);
    if (canonical && canonical !== encoded) return false;
    canonical = encoded;
    const revision = source.slideSizeRevision ?? 0;
    const strategy = source.slideSizeFitStrategy ?? null;
    if (
      typeof revision !== "number" ||
      !Number.isInteger(revision) ||
      revision < 0 ||
      revision > 2_147_483_647 ||
      (revision === 0 && strategy !== null) ||
      (revision > 0 && strategy !== "maximize" && strategy !== "ensure-fit")
    ) {
      return false;
    }
    if (deckRevision === null) {
      deckRevision = revision;
      deckStrategy = strategy as SlideSizeFitStrategy | null;
    } else if (revision !== deckRevision || strategy !== deckStrategy) {
      return false;
    }
    if (!boundedExtent(slide.widthEmu) || !boundedExtent(slide.heightEmu))
      return false;
    if (deckWidth === null) {
      deckWidth = slide.widthEmu;
      deckHeight = slide.heightEmu;
    } else if (slide.widthEmu !== deckWidth || slide.heightEmu !== deckHeight) {
      return false;
    }
  }
  const receipt = parseNativeSlideSizeAuthority(
    (slides[0]?.source &&
    typeof slides[0].source === "object" &&
    !Array.isArray(slides[0].source)
      ? (slides[0].source as Record<string, unknown>)
      : {}
    ).slideSizeAuthority,
  );
  if (!receipt || deckWidth === null || deckHeight === null) return false;
  const receiptSlideOwners =
    receipt.sourceSlideIds ??
    receipt.parts
      .filter((part) => part.kind === "slide")
      .map((part) => part.ownerId);
  const liveSlideOwners = slides.map((slide) => {
    const source = slide.source as Record<string, unknown>;
    return source.slideId as string;
  });
  const receiptSlideOwnerSet = new Set(receiptSlideOwners);
  if (
    receiptSlideOwnerSet.size !== receiptSlideOwners.length ||
    new Set(liveSlideOwners).size !== liveSlideOwners.length ||
    liveSlideOwners.some((ownerId) => !receiptSlideOwnerSet.has(ownerId))
  ) {
    return false;
  }
  // Package parts are path-sorted (slide1, slide10, slide2), not slide order.
  // The server planner and the hydration response verify the true slide order.
  const originalSize = sameSlideSize(
    { widthEmu: deckWidth, heightEmu: deckHeight },
    receipt.slideSize,
  );
  if ((deckRevision === 0) !== originalSize) return false;
  if (deckRevision === 0 && deckStrategy !== null) return false;
  if (
    !slides.every((slide) => {
      const source = slide.source as Record<string, unknown>;
      const ownReceipt = parseNativeSlideSizeAuthority(
        source.slideSizeAuthority,
      );
      const owner = ownReceipt?.parts.find(
        (part) => part.kind === "slide" && part.ownerId === source.slideId,
      );
      if (!owner) return false;
      const blocks = Array.isArray((slide as Record<string, unknown>).blocks)
        ? ((slide as Record<string, unknown>).blocks as readonly Record<
            string,
            unknown
          >[])
        : [];
      return blocks.every((block) => {
        const original = nativeOriginalGeometry(block);
        const source = block.source as Record<string, unknown> | undefined;
        const nativeObject = owner.objects.find(
          (item) => item.shapeId === source?.shapeId,
        );
        if (
          original &&
          nativeObject &&
          SLIDE_SIZE_GEOMETRY_KEYS.some(
            (key) =>
              original[key] !== nativeObject[key] &&
              // A connector the package states zero thick is drawn one EMU
              // thick (the CDC deck's rules under its headings): the writer
              // scales the package's zero and the reopened block is one
              // again, so the deck's custody stands (it was refused whole).
              !(
                nativeObject.tag === "cxnSp" &&
                (key === "width" || key === "height") &&
                nativeObject[key] === 0 &&
                original[key] === 1
              ),
          )
        )
          return false;
        if (
          original &&
          (source?.inheritedFrom || source?.nativeGroupAuthority)
        ) {
          const expected =
            deckRevision === 0
              ? original
              : resizeSlideSizeBlock(
                  original,
                  receipt.slideSize,
                  { widthEmu: deckWidth!, heightEmu: deckHeight! },
                  deckStrategy!,
                );
          if (
            SLIDE_SIZE_GEOMETRY_KEYS.some((key) => block[key] !== expected[key])
          )
            return false;
        }
        const supportedAppend = block.type === "chart" && !source?.shapeId;
        if (!original && !supportedAppend) return false;
        if (
          (block.width === 0 || block.height === 0) &&
          !owner.objects.some(
            (item) => item.shapeId === source?.shapeId && item.tag === "cxnSp",
          )
        )
          return false;
        return (
          validSlideSizeGeometry(block) &&
          (!source?.slideSizeOriginalTableGrid ||
            Boolean(
              parseNativeSlideSizeTableGrid(source.slideSizeOriginalTableGrid),
            ))
        );
      });
    })
  )
    return false;
  return true;
}

/** Stamp one authenticated native size intent across a projected deck. */
export function stampNativeSlideSizeIntent<S extends SlideSizeAuthoritySlide>(
  slides: readonly S[],
  strategy: SlideSizeFitStrategy,
): S[] | null {
  if (!nativeSlideSizeCanApply(slides)) return null;
  const imported = slides.some((slide) => {
    const source = slide.source;
    return Boolean(
      source &&
      typeof source === "object" &&
      !Array.isArray(source) &&
      typeof (source as Record<string, unknown>).slideId === "string",
    );
  });
  if (!imported) return [...slides];
  let revision = 0;
  for (const slide of slides) {
    const source = slide.source as Record<string, unknown>;
    const value = source.slideSizeRevision;
    if (typeof value === "number" && value > revision) revision = value;
  }
  if (revision >= 2_147_483_647) return null;
  return slides.map((slide) => ({
    ...slide,
    source: {
      ...(slide.source as Record<string, unknown>),
      slideSizeRevision: revision + 1,
      slideSizeFitStrategy: strategy,
    },
  }));
}

/** A named slide-size preset (id + human label + its EMU dimensions). */
export interface SlideSizePreset {
  /** Stable id used as the radio/select value + React key. */
  readonly id: string;
  /** i18n message key suffix (e.g. `presentation.slideSize.widescreen`). */
  readonly labelKey: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
}

/**
 * The slide-size presets the setter offers, matching PowerPoint's built-ins:
 *
 *  - `widescreen` 16:9 — 13.333" × 7.5" (the modern PowerPoint default size).
 *  - `standard`   4:3  — 10"     × 7.5".
 *  - `legacy-16-9` 16:9 — 10"    × 5.625" (the legacy on-screen 16:9, == the
 *    historical built-in default 9_144_000 × 5_143_500 this codebase shipped).
 *
 * The "custom" option is NOT a preset (it is driven by the inch inputs); it is
 * handled separately by {@link customSlideSize}.
 */
export const SLIDE_SIZE_PRESETS: readonly SlideSizePreset[] = [
  {
    id: "widescreen",
    labelKey: "presentation.slideSize.widescreen",
    // 13.333" × 7.5" — PowerPoint's "Widescreen (16:9)".
    widthEmu: 12_192_000,
    heightEmu: 6_858_000,
  },
  {
    id: "standard",
    labelKey: "presentation.slideSize.standard",
    // 10" × 7.5" — PowerPoint's "Standard (4:3)".
    widthEmu: 9_144_000,
    heightEmu: 6_858_000,
  },
  {
    id: "legacy-16-9",
    labelKey: "presentation.slideSize.onscreen16x9",
    // 10" × 5.625" — the on-screen 16:9 this deck historically defaulted to
    // (5_143_500 was 5.625" rounded; we expose the exact 16:9 EMU here).
    widthEmu: 9_144_000,
    heightEmu: 5_143_500,
  },
] as const;

/**
 * Design > Slide Size's own menu, in PowerPoint's order: Standard (4:3),
 * Widescreen (16:9), then the Page Setup dialog (PPT-A075). The legacy
 * on-screen 16:9 is a dialog preset, not a menu entry.
 */
export const SLIDE_SIZE_MENU_PRESET_IDS: readonly string[] = ["standard", "widescreen"];

/**
 * The Page Setup dialog's "Slide sized for" list, in PowerPoint's order and
 * at its sizes (Microsoft Support, "Change the size of your slides"; the
 * EMU are the package values those inches/centimetres round from, PPT-B003).
 * Letter Paper and Overhead have the On-screen Show (4:3) size: the deck
 * keeps only a size, so reopening the dialog names them On-screen Show (4:3).
 */
export const PAGE_SETUP_SLIDE_SIZE_PRESETS: readonly SlideSizePreset[] = [
  { id: "standard", labelKey: "presentation.slideSize.page.standard", widthEmu: 9_144_000, heightEmu: 6_858_000 },
  { id: "letter", labelKey: "presentation.slideSize.page.letter", widthEmu: 9_144_000, heightEmu: 6_858_000 },
  { id: "ledger", labelKey: "presentation.slideSize.page.ledger", widthEmu: 12_179_300, heightEmu: 9_134_475 },
  { id: "a3", labelKey: "presentation.slideSize.page.a3", widthEmu: 12_801_600, heightEmu: 9_601_200 },
  { id: "a4", labelKey: "presentation.slideSize.page.a4", widthEmu: 9_906_000, heightEmu: 6_858_000 },
  { id: "b4-iso", labelKey: "presentation.slideSize.page.b4Iso", widthEmu: 10_826_750, heightEmu: 8_120_063 },
  { id: "b5-iso", labelKey: "presentation.slideSize.page.b5Iso", widthEmu: 7_169_150, heightEmu: 5_376_863 },
  { id: "35mm", labelKey: "presentation.slideSize.page.slides35mm", widthEmu: 10_287_000, heightEmu: 6_858_000 },
  { id: "overhead", labelKey: "presentation.slideSize.page.overhead", widthEmu: 9_144_000, heightEmu: 6_858_000 },
  { id: "banner", labelKey: "presentation.slideSize.page.banner", widthEmu: 7_315_200, heightEmu: 914_400 },
  { id: "legacy-16-9", labelKey: "presentation.slideSize.page.onscreen16x9", widthEmu: 9_144_000, heightEmu: 5_143_500 },
  { id: "onscreen-16-10", labelKey: "presentation.slideSize.page.onscreen16x10", widthEmu: 9_144_000, heightEmu: 5_715_000 },
  { id: "widescreen", labelKey: "presentation.slideSize.page.widescreen", widthEmu: 12_192_000, heightEmu: 6_858_000 },
];

const PRESET_BY_ID: ReadonlyMap<string, SlideSizePreset> = new Map(
  [...PAGE_SETUP_SLIDE_SIZE_PRESETS, ...SLIDE_SIZE_PRESETS].map((preset) => [
    preset.id,
    preset,
  ]),
);

/** Look up a preset's EMU size by id (or `undefined` for an unknown id). */
export function presetSize(id: string): SlideSizeEmu | undefined {
  const preset = PRESET_BY_ID.get(id);
  return preset
    ? { widthEmu: preset.widthEmu, heightEmu: preset.heightEmu }
    : undefined;
}

/** Clamp a single inch dimension into the supported range (non-finite → min). */
export function clampInches(inches: number): number {
  if (!Number.isFinite(inches)) return MIN_SLIDE_INCHES;
  return Math.min(MAX_SLIDE_INCHES, Math.max(MIN_SLIDE_INCHES, inches));
}

/**
 * Build a custom {@link SlideSizeEmu} from width/height in INCHES, clamping each
 * dimension to the supported range and converting to integer EMU. Pure — used by
 * the custom-size inputs.
 */
export function customSlideSize(
  widthInches: number,
  heightInches: number,
): SlideSizeEmu {
  return {
    widthEmu: inchesToEmu(clampInches(widthInches)),
    heightEmu: inchesToEmu(clampInches(heightInches)),
  };
}

/**
 * The aspect ratio (width / height) of a size, or `null` when the height is
 * non-positive / non-finite (so a caller never divides by zero). Pure.
 */
export function aspectRatio(size: SlideSizeEmu): number | null {
  if (!Number.isFinite(size.heightEmu) || size.heightEmu <= 0) return null;
  if (!Number.isFinite(size.widthEmu) || size.widthEmu <= 0) return null;
  return size.widthEmu / size.heightEmu;
}

/**
 * Resolve which preset id (if any) a concrete size corresponds to, for
 * reflecting the current deck size back into the selector. A size that matches a
 * preset's EMU dimensions EXACTLY returns that preset id; anything else (incl. a
 * custom size, or the implicit default when no slide carries a size) returns
 * `"custom"`. Pure.
 */
export function matchPresetId(size: SlideSizeEmu | null | undefined): string {
  if (!size) return "custom";
  const match = SLIDE_SIZE_PRESETS.find(
    (preset) =>
      preset.widthEmu === size.widthEmu && preset.heightEmu === size.heightEmu,
  );
  return match ? match.id : "custom";
}

/** {@link matchPresetId} over the Page Setup dialog's list (first exact match). */
export function matchPageSetupPresetId(
  size: SlideSizeEmu | null | undefined,
): string {
  if (!size) return "custom";
  const match = PAGE_SETUP_SLIDE_SIZE_PRESETS.find(
    (preset) =>
      preset.widthEmu === size.widthEmu && preset.heightEmu === size.heightEmu,
  );
  return match ? match.id : "custom";
}

/** PowerPoint's content treatment when a slide changes aspect ratio. */
export type SlideSizeFitStrategy = "maximize" | "ensure-fit";

/** Structural block shape accepted by the deck/master resize adapter. */
export interface SlideSizeGeometryBlock {
  readonly id?: unknown;
  readonly left?: unknown;
  readonly top?: unknown;
  readonly width?: unknown;
  readonly height?: unknown;
  readonly source?: unknown;
}

const SLIDE_SIZE_GEOMETRY_KEYS = ["left", "top", "width", "height"] as const;

function nativeOriginalGeometry(
  block: SlideSizeGeometryBlock,
): SlideSizeGeometryBlock | null {
  const source =
    block.source &&
    typeof block.source === "object" &&
    !Array.isArray(block.source)
      ? (block.source as Record<string, unknown>)
      : null;
  const geometry = exactRecord(source?.slideSizeOriginalGeometry, [
    ...SLIDE_SIZE_GEOMETRY_KEYS,
  ]);
  if (
    !geometry ||
    !boundedOffset(geometry.left) ||
    !boundedOffset(geometry.top) ||
    !validSlideSizeGeometry(geometry)
  ) {
    return null;
  }
  return geometry as SlideSizeGeometryBlock;
}

function validSlideSizeGeometry(value: SlideSizeGeometryBlock): boolean {
  return (
    boundedOffset(value.left) &&
    boundedOffset(value.top) &&
    (boundedExtent(value.width) || value.width === 0) &&
    (boundedExtent(value.height) || value.height === 0) &&
    !(value.width === 0 && value.height === 0)
  );
}

export function nativeSlideSizeOriginalGeometry(
  block: SlideSizeGeometryBlock,
): SlideSizeGeometryBlock | null {
  return nativeOriginalGeometry(block);
}

/**
 * Project an imported deck from its immutable native receipt, rather than from
 * the previous preview. This keeps A→B→C identical to a direct A→C native
 * transform and prevents cumulative Math.round drift.
 */
export function projectNativeSlideSizeDeck<
  S extends SlideSizeAuthoritySlide & {
    readonly blocks?: readonly SlideSizeGeometryBlock[];
  },
>(
  slides: readonly S[],
  next: SlideSizeEmu,
  strategy: SlideSizeFitStrategy,
): S[] | null {
  if (
    !boundedExtent(next.widthEmu) ||
    !boundedExtent(next.heightEmu) ||
    !nativeSlideSizeCanApply(slides)
  ) {
    return null;
  }
  const imported = slides.some((slide) => {
    const source = slide.source;
    return Boolean(
      source &&
      typeof source === "object" &&
      !Array.isArray(source) &&
      typeof (source as Record<string, unknown>).slideId === "string",
    );
  });
  if (!imported) {
    const current = {
      widthEmu: slides[0]?.widthEmu as number,
      heightEmu: slides[0]?.heightEmu as number,
    };
    return slides.map((slide) => ({
      ...slide,
      widthEmu: next.widthEmu,
      heightEmu: next.heightEmu,
      blocks: resizeSlideSizeBlocks(slide.blocks, current, next, strategy),
    })) as S[];
  }
  const firstSource = slides[0]!.source as Record<string, unknown>;
  const receipt = parseNativeSlideSizeAuthority(firstSource.slideSizeAuthority);
  if (!receipt) return null;
  const stamped = stampNativeSlideSizeIntent(slides, strategy);
  if (!stamped) return null;
  // The text scales with the slide, as the Slide Size writer scales every
  
  // words did not). From the size it shows to the one it will show.
  const shownSize = {
    widthEmu: slides[0]!.widthEmu as number,
    heightEmu: slides[0]!.heightEmu as number,
  };
  const shownContentScale = firstSource.slideSizeRevision
    ? slideSizeContentScale(
        receipt.slideSize,
        shownSize,
        firstSource.slideSizeFitStrategy as SlideSizeFitStrategy,
      )
    : 1;
  const nextContentScale = sameSlideSize(next, receipt.slideSize)
    ? 1
    : slideSizeContentScale(receipt.slideSize, next, strategy);
  const projected = stamped.map((slide) => {
    const blocks = slide.blocks;
    if (!blocks) return null;
    const nextBlocks: SlideSizeGeometryBlock[] = [];
    for (const block of blocks) {
      const original = nativeOriginalGeometry(block);
      const source = (block.source ?? {}) as Record<string, unknown>;
      const priorIntent = exactRecord(source.slideSizeAuthoringGeometry, [
        ...SLIDE_SIZE_GEOMETRY_KEYS,
      ]);
      let authoring =
        priorIntent && validSlideSizeGeometry(priorIntent)
          ? priorIntent
          : original;
      const currentSize = {
        widthEmu: slide.widthEmu as number,
        heightEmu: slide.heightEmu as number,
      };
      const priorStrategy = firstSource.slideSizeFitStrategy as
        SlideSizeFitStrategy | undefined;
      const priorProjected =
        authoring &&
        (firstSource.slideSizeRevision
          ? resizeSlideSizeBlock(
              authoring,
              receipt.slideSize,
              currentSize,
              priorStrategy!,
            )
          : authoring);
      if (
        !priorProjected ||
        SLIDE_SIZE_GEOMETRY_KEYS.some(
          (key) => block[key] !== priorProjected[key],
        )
      ) {
        const scale = firstSource.slideSizeRevision
          ? (priorStrategy === "maximize" ? Math.max : Math.min)(
              currentSize.widthEmu / receipt.slideSize.widthEmu,
              currentSize.heightEmu / receipt.slideSize.heightEmu,
            )
          : 1;
        const x =
          (currentSize.widthEmu - receipt.slideSize.widthEmu * scale) / 2;
        const y =
          (currentSize.heightEmu - receipt.slideSize.heightEmu * scale) / 2;
        authoring = {
          left: Math.round(((block.left as number) - x) / scale),
          top: Math.round(((block.top as number) - y) / scale),
          width: Math.round((block.width as number) / scale),
          height: Math.round((block.height as number) / scale),
        };
      }
      if (!authoring || !validSlideSizeGeometry(authoring)) return null;
      const geometry = resizeSlideSizeBlock(
        authoring,
        receipt.slideSize,
        next,
        strategy,
      );
      if (
        !boundedOffset(geometry.left) ||
        !boundedOffset(geometry.top) ||
        !validSlideSizeGeometry(geometry)
      ) {
        return null;
      }
      const grid = parseNativeSlideSizeTableGrid(
        source.slideSizeOriginalTableGrid,
      );
      const scale = (strategy === "maximize" ? Math.max : Math.min)(
        next.widthEmu / receipt.slideSize.widthEmu,
        next.heightEmu / receipt.slideSize.heightEmu,
      );
      let authoringGrid =
        parseNativeSlideSizeTableGrid(source.slideSizeAuthoringTableGrid) ??
        grid;
      const pixelAxis = (axis: readonly number[], factor: number) =>
        axis.map(
          (value) =>
            Math.round((Math.max(1, Math.round(value * factor)) / 9525) * 10) /
            10,
        );
      if (authoringGrid) {
        const priorScale = firstSource.slideSizeRevision
          ? (priorStrategy === "maximize" ? Math.max : Math.min)(
              currentSize.widthEmu / receipt.slideSize.widthEmu,
              currentSize.heightEmu / receipt.slideSize.heightEmu,
            )
          : 1;
        const copied = {
          columns: [...authoringGrid.columns],
          rows: [...authoringGrid.rows],
        };
        for (const [field, axis] of [
          ["colWidths", "columns"],
          ["rowHeights", "rows"],
        ] as const) {
          const live = (block as Record<string, unknown>)[field];
          const expected = pixelAxis(copied[axis], priorScale);
          if (
            Array.isArray(live) &&
            live.length === expected.length &&
            live.some((value, index) => value !== expected[index])
          ) {
            if (
              !live.every(
                (value) =>
                  typeof value === "number" &&
                  Number.isFinite(value) &&
                  value > 0,
              )
            )
              return null;
            copied[axis] = live.map((value) =>
              Math.round((value * 9525) / priorScale),
            );
          }
        }
        authoringGrid = parseNativeSlideSizeTableGrid(copied);
        if (!authoringGrid) return null;
      }
      const tableAxes = authoringGrid
        ? {
            colWidths: pixelAxis(authoringGrid.columns, scale),
            rowHeights: pixelAxis(authoringGrid.rows, scale),
          }
        : {};
      const textAtSize = nativeSlideSizeScalesText(block)
        ? nativeSlideSizeTextAt(block, shownContentScale, nextContentScale)
        : block;
      nextBlocks.push({
        ...textAtSize,
        ...geometry,
        ...tableAxes,
        source: {
          ...((textAtSize.source ?? source) as Record<string, unknown>),
          ...(original &&
          SLIDE_SIZE_GEOMETRY_KEYS.every(
            (key) => original[key] === authoring![key],
          )
            ? {}
            : { slideSizeAuthoringGeometry: authoring }),
          ...(authoringGrid
            ? { slideSizeAuthoringTableGrid: authoringGrid }
            : {}),
        },
      });
    }
    return {
      ...slide,
      widthEmu: next.widthEmu,
      heightEmu: next.heightEmu,
      blocks: nextBlocks,
      ...(sameSlideSize(next, receipt.slideSize)
        ? {
            source: Object.fromEntries(
              Object.entries(slide.source as Record<string, unknown>).filter(
                ([key]) =>
                  key !== "slideSizeRevision" && key !== "slideSizeFitStrategy",
              ),
            ),
          }
        : {}),
    };
  });
  if (projected.some((slide) => slide === null)) return null;
  return projected as S[];
}

/**
 * Whether an imported block's text is scaled with the slide: its words live
 * in the slide, layout or master part the Slide Size writer scales. A
 * SmartArt's live in its diagram parts and a chart's in its chart part, which
 * keep their sizes.
 */
function nativeSlideSizeScalesText(block: SlideSizeGeometryBlock): boolean {
  const record = block as Record<string, unknown>;
  if (!["text", "shape", "table"].includes(String(record.type ?? ""))) return false;
  const source =
    record.source && typeof record.source === "object" && !Array.isArray(record.source)
      ? (record.source as Record<string, unknown>)
      : {};
  return !Object.keys(source).some((key) => /smartart|diagram/i.test(key)) &&
    !Object.keys(record).some((key) => /^smartArt/.test(key));
}

/**
 * Every font size a block states, in one order -- its own, its runs', its
 * rows' runs', a table's cells' and their runs' -- each replaced by
 * `next(value, index)`. The sizes it read come back with the block.
 */
function mapSlideSizeBlockFonts<B extends object>(
  block: B,
  next: (value: number, index: number) => number,
): { block: B; sizes: number[] } {
  const sizes: number[] = [];
  const record = (value: unknown): Record<string, unknown> | undefined =>
    value && typeof value === "object" && !Array.isArray(value)
      ? (value as Record<string, unknown>)
      : undefined;
  const sized = (value: unknown): unknown => {
    const entry = record(value);
    if (!entry || typeof entry.fontSize !== "number" || !Number.isFinite(entry.fontSize)) return value;
    sizes.push(entry.fontSize);
    return { ...entry, fontSize: next(entry.fontSize, sizes.length - 1) };
  };
  const runs = (value: unknown): unknown =>
    Array.isArray(value)
      ? value.map((run) => {
          const entry = record(run);
          const style = record(entry?.style);
          return entry && style ? { ...entry, style: sized(style) } : run;
        })
      : value;
  const source = block as Record<string, unknown>;
  const result: Record<string, unknown> = { ...source };
  if (record(source.style)) result.style = sized(source.style);
  if (Array.isArray(source.runs)) result.runs = runs(source.runs);
  if (Array.isArray(source.paragraphs)) {
    result.paragraphs = source.paragraphs.map((row) => {
      const entry = record(row);
      if (!entry) return row;
      const ownsEnd = typeof entry.endFontSize === "number" && Number.isFinite(entry.endFontSize);
      if (!Array.isArray(entry.runs) && !ownsEnd) return row;
      const scaled: Record<string, unknown> = Array.isArray(entry.runs) ? { ...entry, runs: runs(entry.runs) } : { ...entry };
      // A paragraph's end mark (`a:endParaRPr`) scales with its words, as the
      // Slide Size writer scales its `@sz`: an empty line keeps its height
      
      // 10.5 at 4:3).
      if (ownsEnd) {
        sizes.push(entry.endFontSize as number);
        scaled.endFontSize = next(entry.endFontSize as number, sizes.length - 1);
      }
      return scaled;
    });
  }
  if (source.type === "table" && Array.isArray(source.rows)) {
    result.rows = source.rows.map((row) =>
      Array.isArray(row)
        ? row.map((cell) => {
            const entry = sized(cell);
            const scaled = record(entry);
            return scaled && Array.isArray(scaled.runs) ? { ...scaled, runs: runs(scaled.runs) } : entry;
          })
        : row,
    );
  }
  return { block: result as B, sizes };
}

/** A font size times a content scale, to 0.01 pt as the editor rounds sizes. */
function slideSizeScaledFont(value: number, scale: number): number {
  return Math.min(4000, Math.max(1, Math.round(value * scale * 100) / 100));
}

/**
 * Every font size a block states, times `factor` (`scaleSlideSizeBlockContent`
 * rounds them alike).
 */
export function scaleSlideSizeBlockFonts<B extends object>(block: B, factor: number): B {
  if (!Number.isFinite(factor) || factor <= 0 || Math.abs(factor - 1) < 1e-12) return block;
  return mapSlideSizeBlockFonts(block, (value) => slideSizeScaledFont(value, factor)).block;
}








function nativeSlideSizeTextAt<B extends SlideSizeGeometryBlock>(
  block: B,
  shownScale: number,
  nextScale: number,
): B {
  const source = (block.source ?? {}) as Record<string, unknown>;
  const shown = mapSlideSizeBlockFonts(block, (value) => value).sizes;
  if (!shown.length) return block;
  const kept = source.slideSizeOriginalFontSizes;
  const base =
    Array.isArray(kept) &&
    kept.length === shown.length &&
    kept.every((value) => typeof value === "number" && Number.isFinite(value) && value > 0) &&
    kept.every((value, index) => slideSizeScaledFont(value as number, shownScale) === shown[index])
      ? (kept as number[])
      : shown.map((value) => (shownScale === 1 ? value : Math.round((value / shownScale) * 100) / 100));
  const { block: scaled } = mapSlideSizeBlockFonts(block, (_value, index) =>
    nextScale === 1 ? base[index]! : slideSizeScaledFont(base[index]!, nextScale),
  );
  const nextSource = { ...((scaled.source ?? {}) as Record<string, unknown>) };
  if (nextScale === 1) delete nextSource.slideSizeOriginalFontSizes;
  else nextSource.slideSizeOriginalFontSizes = base;
  return { ...scaled, source: nextSource } as B;
}

/** Exact size equality, used to avoid a no-op history entry. */
export function sameSlideSize(a: SlideSizeEmu, b: SlideSizeEmu): boolean {
  return a.widthEmu === b.widthEmu && a.heightEmu === b.heightEmu;
}

/**
 * PowerPoint asks Maximize / Ensure Fit only when the ASPECT changes. A same-
 * aspect size change has one unambiguous uniform scale, so no choice is needed.
 * Integer cross-products avoid floating point ratio drift.
 */
export function slideSizeNeedsFitChoice(
  current: SlideSizeEmu,
  next: SlideSizeEmu,
): boolean {
  if (sameSlideSize(current, next)) return false;
  return (
    current.widthEmu * next.heightEmu !== next.widthEmu * current.heightEmu
  );
}

function finiteGeometry(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/**
 * Uniformly project one block from `current` slide coordinates into `next`.
 * Ensure Fit letterboxes the old slide inside the new one; Maximize fills the
 * new slide and deliberately allows edge clipping. Blocks without a complete
 * layout box are preserved byte-for-byte.
 */
export function resizeSlideSizeBlock<B extends SlideSizeGeometryBlock>(
  block: B,
  current: SlideSizeEmu,
  next: SlideSizeEmu,
  strategy: SlideSizeFitStrategy,
): B {
  const left = finiteGeometry(block.left);
  const top = finiteGeometry(block.top);
  const width = finiteGeometry(block.width);
  const height = finiteGeometry(block.height);
  if (left === null || top === null || width === null || height === null) {
    return block;
  }
  const scaleX = next.widthEmu / current.widthEmu;
  const scaleY = next.heightEmu / current.heightEmu;
  const scale =
    strategy === "maximize"
      ? Math.max(scaleX, scaleY)
      : Math.min(scaleX, scaleY);
  if (!Number.isFinite(scale) || scale <= 0) return block;
  const offsetX = (next.widthEmu - current.widthEmu * scale) / 2;
  const offsetY = (next.heightEmu - current.heightEmu * scale) / 2;
  return {
    ...block,
    left: Math.round(left * scale + offsetX),
    top: Math.round(top * scale + offsetY),
    width: Math.max(0, Math.round(width * scale)),
    height: Math.max(0, Math.round(height * scale)),
  };
}










export function nativeSlideSizeLayoutProjection(slide: {
  readonly widthEmu?: unknown;
  readonly heightEmu?: unknown;
  readonly source?: unknown;
}): {
  geometry: <G extends SlideSizeGeometryBlock>(geometry: G) => G;
  fontScale: number;
  scaleFont: (value: number) => number;
} | null {
  const source =
    slide.source && typeof slide.source === "object" && !Array.isArray(slide.source)
      ? (slide.source as Record<string, unknown>)
      : null;
  if (!source?.slideSizeRevision) return null;
  const strategy = source.slideSizeFitStrategy;
  if (strategy !== "maximize" && strategy !== "ensure-fit") return null;
  const receipt = parseNativeSlideSizeAuthority(source.slideSizeAuthority);
  if (!receipt || !boundedExtent(slide.widthEmu) || !boundedExtent(slide.heightEmu)) return null;
  const shown = { widthEmu: slide.widthEmu, heightEmu: slide.heightEmu };
  if (sameSlideSize(shown, receipt.slideSize)) return null;
  const fontScale = slideSizeContentScale(receipt.slideSize, shown, strategy);
  return {
    geometry: (geometry) => resizeSlideSizeBlock(geometry, receipt.slideSize, shown, strategy),
    fontScale,
    scaleFont: (value) => slideSizeScaledFont(value, fontScale),
  };
}

/** Resize an editor/master/layout block list without mutating its source. */
export function resizeSlideSizeBlocks<B extends SlideSizeGeometryBlock>(
  blocks: readonly B[] | undefined,
  current: SlideSizeEmu,
  next: SlideSizeEmu,
  strategy: SlideSizeFitStrategy,
): B[] | undefined {
  if (!blocks) return undefined;
  return blocks.map((block) =>
    resizeSlideSizeBlock(block, current, next, strategy),
  );
}

/**
 * The uniform factor a slide-size change scales content by: the new slide
 * over the old one on the axis the strategy keeps (Ensure Fit the smaller,
 * Maximize the larger; one factor when the aspect stays).
 */
export function slideSizeContentScale(
  current: SlideSizeEmu,
  next: SlideSizeEmu,
  strategy: SlideSizeFitStrategy,
): number {
  const scaleX = next.widthEmu / current.widthEmu;
  const scaleY = next.heightEmu / current.heightEmu;
  const scale =
    strategy === "maximize" ? Math.max(scaleX, scaleY) : Math.min(scaleX, scaleY);
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

/** The size a slide that states none is drawn at (the historical default, 10 x 5.625 in). */
export const UNSTATED_SLIDE_SIZE_EMU: SlideSizeEmu = {
  widthEmu: 9_144_000,
  heightEmu: 5_143_500,
};

/** A slide's own stated size, or null. */
export function statedSlideSize(slide: {
  widthEmu?: unknown;
  heightEmu?: unknown;
}): SlideSizeEmu | null {
  const width = finiteGeometry(slide.widthEmu);
  const height = finiteGeometry(slide.heightEmu);
  return width !== null && height !== null && width > 0 && height > 0
    ? { widthEmu: width, heightEmu: height }
    : null;
}

/**
 * A slide brought into a deck from another one (Reuse Slides) takes the
 * deck's size, as every slide of one presentation has one size (p:sldSz): its
 * frames and the sizes it states are scaled and centred the way Slide Size's
 * Ensure Fit does it, so nothing falls off the destination slide. The slide
 * itself when it already has the deck's size.
 */
export function slideFittedToDeckSize<
  S extends { widthEmu?: unknown; heightEmu?: unknown; blocks?: unknown },
>(slide: S, deckSize: SlideSizeEmu): S {
  const current = statedSlideSize(slide) ?? UNSTATED_SLIDE_SIZE_EMU;
  if (sameSlideSize(current, deckSize)) return slide;
  const factor = slideSizeContentScale(current, deckSize, "ensure-fit");
  const blocks = Array.isArray(slide.blocks)
    ? (slide.blocks as SlideSizeGeometryBlock[]).map((block) =>
        scaleSlideSizeBlockContent(
          resizeSlideSizeBlock(block, current, deckSize, "ensure-fit"),
          factor,
        ),
      )
    : slide.blocks;
  return {
    ...slide,
    widthEmu: deckSize.widthEmu,
    heightEmu: deckSize.heightEmu,
    blocks,
  };
}

/** A table cell's text draws at 17 CSS px when it states no size, in points. */
const TABLE_CELL_DRAWN_FONT_SIZE_PT = 12.75;

type Scaled = Record<string, unknown>;

/**
 * Scale the fixed lengths a block states along with its frame when the slide
 * it sits on is resized (PPT-B089).
 *
 * PowerPoint's Slide Size scales the CONTENT, not just the frames: Ensure Fit
 * "decrease[s] the size of your content" and a same-aspect change scales it
 * outright, so a 24 pt title on a slide halved in size is a 12 pt title and a
 * table that fitted still fits. `resizeSlideSizeBlock` moves and sizes the
 * frame; this scales what is drawn inside it at a fixed length: every stated
 * font size (block, runs, paragraph runs, table cells and their runs, chart
 * text), paragraph and character spacing, the row's own indentation and the
 * text-frame insets, outline weights, and a table's column widths and row
 * heights. Proportions, colours and every unitless value are untouched.
 *
 * `drawnFontSize` is the size, in the block's own unit, that its text draws at
 * when it states none (inherited from the canvas or the master). A text or a
 * labelled shape that states no size is given that size scaled, so it keeps
 * its proportion too; without it such a block keeps inheriting.
 */
export function scaleSlideSizeBlockContent<B extends object>(
  block: B,
  factor: number,
  drawnFontSize?: number,
): B {
  if (!Number.isFinite(factor) || factor <= 0 || factor === 1) return block;
  // Font sizes stay inside the editor's own 1..400 bound.
  const size = (value: number) =>
    Math.min(400, Math.max(1, Math.round(value * factor * 100) / 100));
  const length = (value: number) => Math.round(value * factor * 100) / 100;
  const emu = (value: number) => Math.round(value * factor);
  const isNumber = (value: unknown): value is number =>
    typeof value === "number" && Number.isFinite(value);
  const record = (value: unknown): Scaled | undefined =>
    value && typeof value === "object" && !Array.isArray(value)
      ? (value as Scaled)
      : undefined;
  const scaleFields = (
    source: Scaled,
    fields: Readonly<Record<string, (value: number) => number>>,
  ): Scaled => {
    let next = source;
    for (const [field, scale] of Object.entries(fields)) {
      const value = source[field];
      if (isNumber(value)) {
        if (next === source) next = { ...source };
        next[field] = scale(value);
      }
    }
    return next;
  };
  const scaleRuns = (runs: unknown): unknown =>
    Array.isArray(runs)
      ? runs.map((run) => {
          const entry = record(run);
          const style = record(entry?.style);
          return entry && style && isNumber(style.fontSize)
            ? { ...entry, style: { ...style, fontSize: size(style.fontSize) } }
            : run;
        })
      : runs;
  const source = block as Scaled;
  const next: Scaled = { ...source };
  const hasText =
    source.type === "text" ||
    (source.type === "shape" &&
      ((typeof source.text === "string" && source.text.length > 0) ||
        Array.isArray(source.paragraphs)));
  const style = record(source.style);
  if (style || (hasText && isNumber(drawnFontSize))) {
    let nextStyle = scaleFields(style ?? {}, {
      fontSize: size,
      spaceBefore: length,
      spaceAfter: length,
      letterSpacing: length,
    });
    if (!isNumber(nextStyle.fontSize) && hasText && isNumber(drawnFontSize)) {
      nextStyle = { ...nextStyle, fontSize: size(drawnFontSize) };
    }
    if (nextStyle !== style) next.style = nextStyle;
  }
  if (Array.isArray(source.runs)) next.runs = scaleRuns(source.runs);
  if (Array.isArray(source.paragraphs)) {
    next.paragraphs = source.paragraphs.map((paragraph) => {
      const row = record(paragraph);
      if (!row) return paragraph;
      const scaled = scaleFields(row, {
        spaceBefore: length,
        spaceAfter: length,
        marL: emu,
        indent: emu,
      });
      return Array.isArray(row.runs) ? { ...scaled, runs: scaleRuns(row.runs) } : scaled;
    });
  }
  const insets = record(source.textInsets);
  if (insets) {
    next.textInsets = scaleFields(insets, { left: emu, right: emu, top: emu, bottom: emu });
  }
  for (const field of ["strokeWidth", "outlineWidth"] as const) {
    const value = source[field];
    if (isNumber(value)) next[field] = length(value);
  }
  for (const axis of ["colWidths", "rowHeights"] as const) {
    const values = source[axis];
    if (Array.isArray(values) && values.every(isNumber)) {
      next[axis] = values.map((value) => Math.max(1, Math.round(value * factor * 10) / 10));
    }
  }
  if (source.type === "table" && Array.isArray(source.rows)) {
    next.rows = source.rows.map((row) =>
      Array.isArray(row)
        ? row.map((cell) => {
            const entry = record(cell);
            if (!entry) return cell;
            const scaled: Scaled = {
              ...scaleFields(entry, {
                fontSize: size,
                marginLeft: length,
                marginRight: length,
                marginTop: length,
                marginBottom: length,
              }),
            };
            if (!isNumber(scaled.fontSize)) {
              scaled.fontSize = size(TABLE_CELL_DRAWN_FONT_SIZE_PT);
            }
            if (Array.isArray(entry.runs)) scaled.runs = scaleRuns(entry.runs);
            return scaled;
          })
        : row,
    );
  }
  const chart = record(source.chartV2);
  const format = record(chart?.format);
  const text = record(format?.text);
  if (chart && format && text && isNumber(text.fontSizePt)) {
    next.chartV2 = {
      ...chart,
      format: { ...format, text: { ...text, fontSizePt: size(text.fontSizePt) } },
    };
  }
  return next as B;
}
