/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PdfSemanticAppearanceGeometry,
  PdfSemanticAppearanceOwner,
  PdfSemanticAppearanceOwnerInventory,
  PdfSemanticAppearancePreserveOnly,
  PdfSemanticAppearanceFormEvidence,
} from "./semantic-appearance-owner-inventory";
import type { PdfVectorPageGeometry } from "./vector-capability-reconciliation";
import type {
  PdfVectorContentItem,
  PdfVectorSimplePathCommand,
  PdfVectorSimpleSubpath,
} from "./vector-content-inventory";

export const PDF_SEMANTIC_APPEARANCE_RECONCILIATION_SCHEMA_VERSION =
  "mrai.pdf-semantic-appearance-reconciliation.v1" as const;

const MAX_INSTANCES_PER_PAGE = 4_096;
const MAX_VISIBLE_VECTORS_PER_PAGE = 100_000;
const MAX_OPERATORS_PER_PAGE = 1_000_000;
const MATRIX_TOLERANCE = 0.000_01;
const RECT_TOLERANCE = 0.000_1;
const OWNER_BOUNDS_TOLERANCE = 0.004;
const TEXT_ANNOTATION_ICON_SIZE = 22;
const ANNOTATION_REF = /^([1-9][0-9]{0,9})R([0-9]{1,5})?$/;

type Matrix = readonly [number, number, number, number, number, number];
type RawRect = readonly [number, number, number, number];

export interface PdfSemanticAppearanceOperatorIds {
  readonly beginAnnotation: number;
  readonly endAnnotation: number;
}

export type PdfSemanticAppearanceInstanceDisposition =
  "semantic-owner-eligible" | "semantic-owner-preserve-only" | "preserve-only";

export type PdfSemanticAppearanceInstanceReason =
  | "shared-appearance-stream"
  | "unlocatable-annotation-id"
  | "malformed-payload"
  | "nested-block"
  | "unbalanced-block"
  | "duplicate-instance"
  | "ambiguous-owner"
  | "unmatched-owner"
  | "rect-mismatch"
  | "transform-mismatch"
  | "matrix-mismatch"
  | "appearance-state-mismatch"
  | "unsupported-own-canvas"
  | "vector-reprojection-failed"
  | "vector-outside-owner";

export interface PdfSemanticAppearanceMatchedOwner {
  readonly recordId: string;
  readonly objectNumber: number;
  readonly generation: number;
  readonly ownerKind: "widget" | "annotation";
  readonly semanticKind:
    | "text"
    | "checkbox"
    | "radio"
    | "dropdown"
    | "listbox"
    | "Link"
    | "Text"
    | "Highlight"
    | "Underline"
    | "Square";
  readonly sharedAppearance: boolean;
}

export interface PdfSemanticAppearanceReprojectedVector {
  readonly visibleOperatorIndex: number;
  readonly pageBounds: Readonly<{
    x: number;
    y: number;
    width: number;
    height: number;
  }>;
  readonly pageSubpath?: PdfVectorSimpleSubpath;
}

export interface PdfSemanticAppearanceInstanceEvidence {
  readonly instanceIndex: number;
  readonly beginOperatorIndex: number;
  readonly endOperatorIndex: number;
  readonly annotationRef: Readonly<{
    objectNumber: number;
    generation: number;
  }> | null;
  readonly rect: RawRect | null;
  readonly transform: Matrix | null;
  readonly matrix: Matrix | null;
  readonly hasOwnCanvas: boolean | null;
  readonly disposition: PdfSemanticAppearanceInstanceDisposition;
  readonly reason: PdfSemanticAppearanceInstanceReason | null;
  readonly owner: PdfSemanticAppearanceMatchedOwner | null;
  readonly visibleOperatorIndexes: readonly number[];
  readonly vectors: readonly PdfSemanticAppearanceReprojectedVector[];
}

export interface PdfSemanticAppearanceVectorAssignment {
  readonly visibleOperatorIndex: number;
  readonly instanceIndex: number | null;
  readonly disposition: PdfSemanticAppearanceInstanceDisposition;
  readonly reason:
    PdfSemanticAppearanceInstanceReason | "outside-annotation-block" | null;
  readonly owner: PdfSemanticAppearanceMatchedOwner | null;
  readonly pageBounds:
    PdfSemanticAppearanceReprojectedVector["pageBounds"] | null;
  readonly pageSubpath?: PdfVectorSimpleSubpath;
}

export interface PdfSemanticAppearanceReconciliationResult {
  readonly schemaVersion: typeof PDF_SEMANTIC_APPEARANCE_RECONCILIATION_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly instances: readonly PdfSemanticAppearanceInstanceEvidence[];
  readonly vectorAssignments: readonly PdfSemanticAppearanceVectorAssignment[];
  readonly diagnostics: Readonly<{
    instanceCount: number;
    matchedInstanceCount: number;
    eligibleInstanceCount: number;
    sharedPreserveInstanceCount: number;
    synthesizedInstanceCount: number;
    invalidInstanceCount: number;
    unmatchedOwnerCount: number;
    eligibleVectorCount: number;
    preserveOnlyVectorCount: number;
    outsideBlockVectorCount: number;
    ambiguousInstanceCount: number;
  }>;
}

export interface PdfSemanticAppearanceReconciliationInput {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly operatorIds: PdfSemanticAppearanceOperatorIds;
  readonly operatorList: Readonly<{
    fnArray: ArrayLike<number>;
    argsArray: ArrayLike<unknown>;
  }>;
  readonly visibleVectors: readonly PdfVectorContentItem[];
  readonly inventory: PdfSemanticAppearanceOwnerInventory;
}

interface ParsedPayload {
  readonly annotationRef: Readonly<{
    objectNumber: number;
    generation: number;
  }> | null;
  readonly rect: RawRect | null;
  readonly transform: Matrix | null;
  readonly matrix: Matrix | null;
  readonly hasOwnCanvas: boolean | null;
  readonly malformed: boolean;
}

interface RawInstance {
  readonly beginOperatorIndex: number;
  readonly endOperatorIndex: number;
  readonly payload: ParsedPayload;
  readonly structuralReason: "nested-block" | "unbalanced-block" | null;
}

interface OwnerCandidate {
  readonly recordId: string;
  readonly objectNumber: number;
  readonly generation: number;
  readonly pageNumber: number;
  readonly geometry: PdfSemanticAppearanceGeometry;
  readonly ownerKind: "widget" | "annotation";
  readonly semanticKind: PdfSemanticAppearanceMatchedOwner["semanticKind"];
  readonly sharedAppearance: boolean;
  readonly exactOwner: PdfSemanticAppearanceOwner | null;
  readonly sharedRecord: PdfSemanticAppearancePreserveOnly | null;
}

function rounded(value: number): number {
  const result = Number(value.toFixed(12));
  return Object.is(result, -0) ? 0 : result;
}

function exactPageGeometry(input: PdfVectorPageGeometry) {
  const crop = input.cropBox;
  if (
    ![crop.xMin, crop.yMin, crop.xMax, crop.yMax].every(Number.isFinite) ||
    crop.xMax <= crop.xMin ||
    crop.yMax <= crop.yMin ||
    ![0, 90, 180, 270].includes(input.rotation)
  ) {
    throw new Error("PDF appearance reconciliation page geometry is invalid.");
  }
  return Object.freeze({
    cropBox: Object.freeze({ ...crop }),
    rotation: input.rotation,
  });
}

function projectedPageSize(geometry: PdfVectorPageGeometry) {
  const cropWidth = geometry.cropBox.xMax - geometry.cropBox.xMin;
  const cropHeight = geometry.cropBox.yMax - geometry.cropBox.yMin;
  return Object.freeze({
    width:
      geometry.rotation === 0 || geometry.rotation === 180
        ? cropWidth
        : cropHeight,
    height:
      geometry.rotation === 0 || geometry.rotation === 180
        ? cropHeight
        : cropWidth,
  });
}

function projectedPoint(x: number, y: number, geometry: PdfVectorPageGeometry) {
  const crop = geometry.cropBox;
  const page = projectedPageSize(geometry);
  let projectedX: number;
  let projectedY: number;
  if (geometry.rotation === 0) {
    projectedX = x - crop.xMin;
    projectedY = crop.yMax - y;
  } else if (geometry.rotation === 90) {
    projectedX = y - crop.yMin;
    projectedY = x - crop.xMin;
  } else if (geometry.rotation === 180) {
    projectedX = crop.xMax - x;
    projectedY = y - crop.yMin;
  } else {
    projectedX = crop.yMax - y;
    projectedY = crop.xMax - x;
  }
  return Object.freeze({
    x: rounded(projectedX / page.width),
    y: rounded(projectedY / page.height),
  });
}

function rawPointFromNormalized(
  x: number,
  y: number,
  geometry: PdfVectorPageGeometry,
) {
  const crop = geometry.cropBox;
  const page = projectedPageSize(geometry);
  const projectedX = x * page.width;
  const projectedY = y * page.height;
  if (geometry.rotation === 0) {
    return Object.freeze({
      x: crop.xMin + projectedX,
      y: crop.yMax - projectedY,
    });
  }
  if (geometry.rotation === 90) {
    return Object.freeze({
      x: crop.xMin + projectedY,
      y: crop.yMin + projectedX,
    });
  }
  if (geometry.rotation === 180) {
    return Object.freeze({
      x: crop.xMax - projectedX,
      y: crop.yMin + projectedY,
    });
  }
  return Object.freeze({
    x: crop.xMax - projectedY,
    y: crop.yMax - projectedX,
  });
}

function rawRectFromOwner(
  geometry: PdfSemanticAppearanceGeometry,
  pageGeometry: PdfVectorPageGeometry,
): RawRect {
  const rect = geometry.rect;
  const points = [
    rawPointFromNormalized(rect.x, rect.y, pageGeometry),
    rawPointFromNormalized(rect.x + rect.width, rect.y, pageGeometry),
    rawPointFromNormalized(
      rect.x + rect.width,
      rect.y + rect.height,
      pageGeometry,
    ),
    rawPointFromNormalized(rect.x, rect.y + rect.height, pageGeometry),
  ];
  return Object.freeze([
    Math.min(...points.map((point) => point.x)),
    Math.min(...points.map((point) => point.y)),
    Math.max(...points.map((point) => point.x)),
    Math.max(...points.map((point) => point.y)),
  ]);
}

function exactNumbers(
  value: unknown,
  length: number,
): readonly number[] | null {
  if (
    !Array.isArray(value) &&
    !(ArrayBuffer.isView(value) && !(value instanceof DataView))
  ) {
    return null;
  }
  const values = Array.from(value as ArrayLike<unknown>, Number);
  return values.length === length && values.every(Number.isFinite)
    ? Object.freeze(values)
    : null;
}

function matrix(value: unknown): Matrix | null {
  const values = exactNumbers(value, 6);
  return values
    ? Object.freeze([
        values[0],
        values[1],
        values[2],
        values[3],
        values[4],
        values[5],
      ])
    : null;
}

function rect(value: unknown): RawRect | null {
  const values = exactNumbers(value, 4);
  if (!values || values[2] <= values[0] || values[3] <= values[1]) return null;
  return Object.freeze([values[0], values[1], values[2], values[3]]);
}

function annotationRef(value: unknown) {
  if (typeof value !== "string") return null;
  const match = ANNOTATION_REF.exec(value);
  if (!match) return null;
  const objectNumber = Number(match[1]);
  const generation = match[2] ? Number(match[2]) : 0;
  return Number.isSafeInteger(objectNumber) && Number.isSafeInteger(generation)
    ? Object.freeze({ objectNumber, generation })
    : null;
}

function parsedPayload(value: unknown): ParsedPayload {
  if (!Array.isArray(value) || value.length !== 5) {
    return Object.freeze({
      annotationRef: Array.isArray(value) ? annotationRef(value[0]) : null,
      rect: null,
      transform: null,
      matrix: null,
      hasOwnCanvas: null,
      malformed: true,
    });
  }
  const parsedRef = annotationRef(value[0]);
  const parsedRect = rect(value[1]);
  const parsedTransform = matrix(value[2]);
  const parsedMatrix = matrix(value[3]);
  const ownCanvas = typeof value[4] === "boolean" ? value[4] : null;
  const matricesAreInvertible =
    parsedTransform !== null &&
    parsedMatrix !== null &&
    Math.abs(
      parsedTransform[0] * parsedTransform[3] -
        parsedTransform[1] * parsedTransform[2],
    ) > 1e-12 &&
    Math.abs(
      parsedMatrix[0] * parsedMatrix[3] - parsedMatrix[1] * parsedMatrix[2],
    ) > 1e-12;
  return Object.freeze({
    annotationRef: parsedRef,
    rect: parsedRect,
    transform: parsedTransform,
    matrix: parsedMatrix,
    hasOwnCanvas: ownCanvas,
    malformed:
      !parsedRef ||
      !parsedRect ||
      !parsedTransform ||
      !parsedMatrix ||
      !matricesAreInvertible ||
      ownCanvas === null,
  });
}

function parseInstances(
  operatorList: PdfSemanticAppearanceReconciliationInput["operatorList"],
  ids: PdfSemanticAppearanceOperatorIds,
): Readonly<{
  instances: readonly RawInstance[];
  structuralInvalid: boolean;
}> {
  const count = Math.min(
    operatorList.fnArray.length,
    operatorList.argsArray.length,
  );
  if (
    operatorList.fnArray.length !== operatorList.argsArray.length ||
    count > MAX_OPERATORS_PER_PAGE
  ) {
    throw new Error("PDF appearance reconciliation operator list is invalid.");
  }
  const instances: RawInstance[] = [];
  let active: Readonly<{
    beginOperatorIndex: number;
    payload: ParsedPayload;
    nested: boolean;
  }> | null = null;
  let structuralInvalid = false;
  for (let index = 0; index < count; index += 1) {
    const fn = operatorList.fnArray[index];
    if (fn === ids.beginAnnotation) {
      if (active) {
        structuralInvalid = true;
        active = Object.freeze({
          beginOperatorIndex: active.beginOperatorIndex,
          payload: active.payload,
          nested: true,
        });
      } else {
        active = Object.freeze({
          beginOperatorIndex: index,
          payload: parsedPayload(operatorList.argsArray[index]),
          nested: false,
        });
      }
      continue;
    }
    if (fn !== ids.endAnnotation) continue;
    if (!active) {
      structuralInvalid = true;
      continue;
    }
    instances.push(
      Object.freeze({
        beginOperatorIndex: active.beginOperatorIndex,
        endOperatorIndex: index,
        payload: active.payload,
        structuralReason: active.nested ? "nested-block" : null,
      }),
    );
    active = null;
    if (instances.length > MAX_INSTANCES_PER_PAGE) {
      throw new Error(
        "PDF appearance reconciliation instance capacity exceeded.",
      );
    }
  }
  if (active) {
    structuralInvalid = true;
    instances.push(
      Object.freeze({
        beginOperatorIndex: active.beginOperatorIndex,
        endOperatorIndex: count,
        payload: active.payload,
        structuralReason: "unbalanced-block",
      }),
    );
  }
  return Object.freeze({
    instances: Object.freeze(instances),
    structuralInvalid,
  });
}

function multiply(left: Matrix, right: Matrix): Matrix {
  return Object.freeze([
    left[0] * right[0] + left[2] * right[1],
    left[1] * right[0] + left[3] * right[1],
    left[0] * right[2] + left[2] * right[3],
    left[1] * right[2] + left[3] * right[3],
    left[0] * right[4] + left[2] * right[5] + left[4],
    left[1] * right[4] + left[3] * right[5] + left[5],
  ]);
}

function transformedPoint(
  value: Readonly<{ x: number; y: number }>,
  transform: Matrix,
) {
  return Object.freeze({
    x: transform[0] * value.x + transform[2] * value.y + transform[4],
    y: transform[1] * value.x + transform[3] * value.y + transform[5],
  });
}

function matrixCompatible(left: Matrix, right: Matrix): boolean {
  return left.every(
    (value, index) => Math.abs(value - right[index]) <= MATRIX_TOLERANCE,
  );
}

function rectCompatible(left: RawRect, right: RawRect): boolean {
  return left.every(
    (value, index) => Math.abs(value - right[index]) <= RECT_TOLERANCE,
  );
}

function axialBounds(rectangle: RawRect, transform: Matrix): RawRect {
  const points = [
    transformedPoint({ x: rectangle[0], y: rectangle[1] }, transform),
    transformedPoint({ x: rectangle[2], y: rectangle[1] }, transform),
    transformedPoint({ x: rectangle[2], y: rectangle[3] }, transform),
    transformedPoint({ x: rectangle[0], y: rectangle[3] }, transform),
  ];
  return Object.freeze([
    Math.min(...points.map((point) => point.x)),
    Math.min(...points.map((point) => point.y)),
    Math.max(...points.map((point) => point.x)),
    Math.max(...points.map((point) => point.y)),
  ]);
}

function expectedAnnotationTransform(
  ownerRect: RawRect,
  form: PdfSemanticAppearanceFormEvidence,
): Matrix {
  const transformedBBox = axialBounds(form.bbox, form.matrix);
  const width = transformedBBox[2] - transformedBBox[0];
  const height = transformedBBox[3] - transformedBBox[1];
  if (width === 0 || height === 0) {
    return Object.freeze([1, 0, 0, 1, ownerRect[0], ownerRect[1]]);
  }
  const xRatio = (ownerRect[2] - ownerRect[0]) / width;
  const yRatio = (ownerRect[3] - ownerRect[1]) / height;
  return Object.freeze([
    xRatio,
    0,
    0,
    yRatio,
    ownerRect[0] - transformedBBox[0] * xRatio,
    ownerRect[1] - transformedBBox[1] * yRatio,
  ]);
}

function ownerCandidate(owner: PdfSemanticAppearanceOwner): OwnerCandidate {
  return Object.freeze({
    recordId: owner.id,
    objectNumber: owner.annotation.objectNumber,
    generation: owner.annotation.generation,
    pageNumber: owner.pageNumber,
    geometry: owner.geometry,
    ownerKind: owner.ownerKind,
    semanticKind:
      owner.ownerKind === "widget" ? owner.widget.fieldKind : owner.subtype,
    sharedAppearance: false,
    exactOwner: owner,
    sharedRecord: null,
  });
}

function sharedCandidate(
  item: PdfSemanticAppearancePreserveOnly,
): OwnerCandidate | null {
  if (
    item.reason !== "shared-appearance-stream" ||
    !item.widget ||
    !item.geometry ||
    !("objectNumber" in item.locator)
  ) {
    return null;
  }
  return Object.freeze({
    recordId: item.id,
    objectNumber: item.locator.objectNumber,
    generation: item.locator.generation,
    pageNumber: item.pageNumber,
    geometry: item.geometry,
    ownerKind: "widget",
    semanticKind: item.widget.fieldKind,
    sharedAppearance: true,
    exactOwner: null,
    sharedRecord: item,
  });
}

function matchedOwner(
  candidate: OwnerCandidate,
): PdfSemanticAppearanceMatchedOwner {
  return Object.freeze({
    recordId: candidate.recordId,
    objectNumber: candidate.objectNumber,
    generation: candidate.generation,
    ownerKind: candidate.ownerKind,
    semanticKind: candidate.semanticKind,
    sharedAppearance: candidate.sharedAppearance,
  });
}

function widgetEvidence(candidate: OwnerCandidate) {
  if (candidate.exactOwner?.ownerKind === "widget") {
    return candidate.exactOwner.widget;
  }
  return candidate.sharedRecord?.widget ?? null;
}

function selectedForm(candidate: OwnerCandidate) {
  const widget = widgetEvidence(candidate);
  if (!widget) return null;
  if (widget.appearance.kind === "direct-form") {
    return widget.appearance.form;
  }
  const selectedEntries = widget.appearance.entries.filter(
    (entry) => entry.state === widget.appearance.selectedAppearanceState,
  );
  if (selectedEntries.length !== 1) return null;
  const selectedEntry = selectedEntries[0].form;
  const selectedForm = widget.appearance.selectedForm;
  return selectedEntry.objectNumber === selectedForm.objectNumber &&
    selectedEntry.generation === selectedForm.generation &&
    selectedEntry.objectSha256 === selectedForm.objectSha256 &&
    selectedEntry.streamSha256 === selectedForm.streamSha256
    ? selectedForm
    : null;
}

function validateInstance(
  instance: RawInstance,
  candidate: OwnerCandidate,
  pageGeometry: PdfVectorPageGeometry,
): PdfSemanticAppearanceInstanceReason | null {
  const payload = instance.payload;
  if (instance.structuralReason) return instance.structuralReason;
  if (
    payload.malformed ||
    !payload.rect ||
    !payload.transform ||
    !payload.matrix
  ) {
    return payload.annotationRef
      ? "malformed-payload"
      : "unlocatable-annotation-id";
  }
  const sourceRect = rawRectFromOwner(candidate.geometry, pageGeometry);
  if (candidate.ownerKind === "widget") {
    if (!rectCompatible(payload.rect, sourceRect)) return "rect-mismatch";
    if (payload.hasOwnCanvas !== false) return "unsupported-own-canvas";
    const form = selectedForm(candidate);
    if (!form) return "appearance-state-mismatch";
    if (!matrixCompatible(payload.matrix, form.matrix))
      return "matrix-mismatch";
    if (
      !matrixCompatible(
        payload.transform,
        expectedAnnotationTransform(sourceRect, form),
      )
    ) {
      return "transform-mismatch";
    }
    return candidate.sharedAppearance ? "shared-appearance-stream" : null;
  }
  if (
    !candidate.exactOwner ||
    candidate.exactOwner.ownerKind !== "annotation"
  ) {
    return "unmatched-owner";
  }
  const subtype = candidate.exactOwner.subtype;
  const identity: Matrix = [1, 0, 0, 1, 0, 0];
  if (subtype === "Text") {
    const expectedRect: RawRect = [
      sourceRect[0],
      sourceRect[3] - TEXT_ANNOTATION_ICON_SIZE,
      sourceRect[0] + TEXT_ANNOTATION_ICON_SIZE,
      sourceRect[3],
    ];
    const expectedTransform: Matrix = [
      TEXT_ANNOTATION_ICON_SIZE,
      0,
      0,
      TEXT_ANNOTATION_ICON_SIZE,
      expectedRect[0],
      expectedRect[1],
    ];
    if (!rectCompatible(payload.rect, expectedRect)) return "rect-mismatch";
    if (payload.hasOwnCanvas !== true) return "unsupported-own-canvas";
    if (!matrixCompatible(payload.matrix, identity)) return "matrix-mismatch";
    return matrixCompatible(payload.transform, expectedTransform)
      ? null
      : "transform-mismatch";
  }
  if (!rectCompatible(payload.rect, sourceRect)) return "rect-mismatch";
  if (payload.hasOwnCanvas !== false) return "unsupported-own-canvas";
  if (!matrixCompatible(payload.matrix, identity)) return "matrix-mismatch";
  return matrixCompatible(payload.transform, identity)
    ? null
    : "transform-mismatch";
}

function localPointFromNormalized(
  point: Readonly<{ x: number; y: number }>,
  pageGeometry: PdfVectorPageGeometry,
) {
  const page = projectedPageSize(pageGeometry);
  return Object.freeze({
    x: point.x * page.width,
    y: page.height - point.y * page.height,
  });
}

function reprojectPoint(
  point: Readonly<{ x: number; y: number }>,
  combined: Matrix,
  pageGeometry: PdfVectorPageGeometry,
) {
  const local = localPointFromNormalized(point, pageGeometry);
  const pagePoint = transformedPoint(local, combined);
  return projectedPoint(pagePoint.x, pagePoint.y, pageGeometry);
}

function reprojectSubpath(
  subpath: PdfVectorSimpleSubpath,
  combined: Matrix,
  pageGeometry: PdfVectorPageGeometry,
): PdfVectorSimpleSubpath | null {
  const commands: PdfVectorSimplePathCommand[] = [];
  for (const command of subpath.commands) {
    if (command.kind === "closePath") {
      commands.push(Object.freeze({ kind: "closePath" }));
      continue;
    }
    if ("point" in command) {
      commands.push(
        Object.freeze({
          kind: command.kind,
          point: reprojectPoint(command.point, combined, pageGeometry),
        }),
      );
      continue;
    }
    if (command.kind === "quadraticCurveTo") {
      commands.push(
        Object.freeze({
          kind: command.kind,
          control: reprojectPoint(command.control, combined, pageGeometry),
          end: reprojectPoint(command.end, combined, pageGeometry),
        }),
      );
      continue;
    }
    commands.push(
      Object.freeze({
        kind: "curveTo",
        control1: reprojectPoint(command.control1, combined, pageGeometry),
        control2: reprojectPoint(command.control2, combined, pageGeometry),
        end: reprojectPoint(command.end, combined, pageGeometry),
      }),
    );
  }
  return Object.freeze({
    commands: Object.freeze(commands),
    closed: subpath.closed,
  });
}

function reprojectVector(
  vector: PdfVectorContentItem,
  transform: Matrix,
  annotationMatrix: Matrix,
  pageGeometry: PdfVectorPageGeometry,
): PdfSemanticAppearanceReprojectedVector | null {
  const page = projectedPageSize(pageGeometry);
  const localLeft = vector.x * page.width;
  const localRight = (vector.x + vector.width) * page.width;
  const localTop = page.height - vector.y * page.height;
  const localBottom = localTop - vector.height * page.height;
  const combined = multiply(transform, annotationMatrix);
  const corners = [
    transformedPoint({ x: localLeft, y: localBottom }, combined),
    transformedPoint({ x: localRight, y: localBottom }, combined),
    transformedPoint({ x: localRight, y: localTop }, combined),
    transformedPoint({ x: localLeft, y: localTop }, combined),
  ].map((point) => projectedPoint(point.x, point.y, pageGeometry));
  const left = Math.min(...corners.map((point) => point.x));
  const top = Math.min(...corners.map((point) => point.y));
  const right = Math.max(...corners.map((point) => point.x));
  const bottom = Math.max(...corners.map((point) => point.y));
  const pageBounds = Object.freeze({
    x: rounded(left),
    y: rounded(top),
    width: rounded(right - left),
    height: rounded(bottom - top),
  });
  if (
    ![pageBounds.x, pageBounds.y, pageBounds.width, pageBounds.height].every(
      Number.isFinite,
    ) ||
    pageBounds.width <= 0 ||
    pageBounds.height <= 0
  ) {
    return null;
  }
  return Object.freeze({
    visibleOperatorIndex: vector.operatorIndex,
    pageBounds,
    ...(vector.simpleSubpath
      ? {
          pageSubpath: reprojectSubpath(
            vector.simpleSubpath,
            combined,
            pageGeometry,
          )!,
        }
      : {}),
  });
}

function boundsInsideOwner(
  bounds: PdfSemanticAppearanceReprojectedVector["pageBounds"],
  geometry: PdfSemanticAppearanceGeometry,
): boolean {
  const rect = geometry.rect;
  return (
    bounds.x >= rect.x - OWNER_BOUNDS_TOLERANCE &&
    bounds.x + bounds.width <= rect.x + rect.width + OWNER_BOUNDS_TOLERANCE &&
    bounds.y >= rect.y - OWNER_BOUNDS_TOLERANCE &&
    bounds.y + bounds.height <= rect.y + rect.height + OWNER_BOUNDS_TOLERANCE
  );
}

export async function reconcilePdfSemanticAppearanceInstances(
  input: PdfSemanticAppearanceReconciliationInput,
): Promise<PdfSemanticAppearanceReconciliationResult> {
  const sourceSha256 = input.sourceSha256.trim().toLowerCase();
  if (!/^[0-9a-f]{64}$/.test(sourceSha256)) {
    throw new Error("PDF appearance reconciliation requires source SHA-256.");
  }
  if (!Number.isSafeInteger(input.pageNumber) || input.pageNumber < 1) {
    throw new Error("PDF appearance reconciliation requires a page number.");
  }
  if (input.inventory.sourceSha256 !== sourceSha256) {
    throw new Error("PDF appearance reconciliation source authority mismatch.");
  }
  if (
    !Number.isSafeInteger(input.inventory.pageCount) ||
    input.pageNumber > input.inventory.pageCount ||
    [...input.inventory.owners, ...input.inventory.preserveOnly].some(
      (record) => record.sourceSha256 !== sourceSha256,
    )
  ) {
    throw new Error("PDF appearance reconciliation page authority mismatch.");
  }
  if (input.visibleVectors.length > MAX_VISIBLE_VECTORS_PER_PAGE) {
    throw new Error("PDF appearance reconciliation vector capacity exceeded.");
  }
  const visibleOperatorIndexes = new Set<number>();
  for (const vector of input.visibleVectors) {
    if (
      !Number.isSafeInteger(vector.operatorIndex) ||
      vector.operatorIndex < 0 ||
      visibleOperatorIndexes.has(vector.operatorIndex)
    ) {
      throw new Error(
        "PDF appearance reconciliation vector identity is invalid.",
      );
    }
    visibleOperatorIndexes.add(vector.operatorIndex);
  }
  if (
    !Number.isSafeInteger(input.operatorIds.beginAnnotation) ||
    !Number.isSafeInteger(input.operatorIds.endAnnotation) ||
    input.operatorIds.beginAnnotation === input.operatorIds.endAnnotation
  ) {
    throw new Error(
      "PDF appearance reconciliation operator identity is invalid.",
    );
  }
  const pageGeometry = exactPageGeometry(input.pageGeometry);
  const parsed = parseInstances(input.operatorList, input.operatorIds);
  const ownerCandidates: OwnerCandidate[] = [
    ...input.inventory.owners
      .filter((owner) => owner.pageNumber === input.pageNumber)
      .map(ownerCandidate),
    ...input.inventory.preserveOnly
      .filter((item) => item.pageNumber === input.pageNumber)
      .map(sharedCandidate)
      .filter((item): item is OwnerCandidate => item !== null),
  ];
  const ownerCandidatesByRef = new Map<string, OwnerCandidate[]>();
  for (const candidate of ownerCandidates) {
    const key = `${candidate.objectNumber}:${candidate.generation}`;
    ownerCandidatesByRef.set(key, [
      ...(ownerCandidatesByRef.get(key) ?? []),
      candidate,
    ]);
  }
  const instanceRefCounts = new Map<string, number>();
  for (const instance of parsed.instances) {
    const ref = instance.payload.annotationRef;
    if (!ref) continue;
    const key = `${ref.objectNumber}:${ref.generation}`;
    instanceRefCounts.set(key, (instanceRefCounts.get(key) ?? 0) + 1);
  }
  const coveredVisible = new Set<number>();
  const instances: PdfSemanticAppearanceInstanceEvidence[] = [];
  const vectorAssignments: PdfSemanticAppearanceVectorAssignment[] = [];
  for (const [instanceIndex, instance] of parsed.instances.entries()) {
    const visible = input.visibleVectors.filter(
      (vector) =>
        vector.operatorIndex > instance.beginOperatorIndex &&
        vector.operatorIndex < instance.endOperatorIndex,
    );
    for (const vector of visible) coveredVisible.add(vector.operatorIndex);
    const ref = instance.payload.annotationRef;
    const candidates = ref
      ? (ownerCandidatesByRef.get(`${ref.objectNumber}:${ref.generation}`) ??
        [])
      : [];
    let candidate: OwnerCandidate | null = null;
    let reason: PdfSemanticAppearanceInstanceReason | null = null;
    if (parsed.structuralInvalid || instance.structuralReason) {
      reason = instance.structuralReason ?? "unbalanced-block";
    } else if (instance.payload.malformed) {
      reason = ref ? "malformed-payload" : "unlocatable-annotation-id";
    } else if (
      ref &&
      (instanceRefCounts.get(`${ref.objectNumber}:${ref.generation}`) ?? 0) > 1
    ) {
      reason = "duplicate-instance";
    } else if (candidates.length === 0) {
      reason = "unmatched-owner";
    } else if (candidates.length > 1) {
      reason = "ambiguous-owner";
    } else {
      candidate = candidates[0];
      reason = validateInstance(instance, candidate, pageGeometry);
    }
    let vectors: readonly PdfSemanticAppearanceReprojectedVector[] = [];
    if (
      candidate &&
      instance.payload.transform &&
      instance.payload.matrix &&
      (reason === null || reason === "shared-appearance-stream")
    ) {
      const reprojected = visible.map((vector) =>
        reprojectVector(
          vector,
          instance.payload.transform!,
          instance.payload.matrix!,
          pageGeometry,
        ),
      );
      if (reprojected.some((vector) => vector === null)) {
        reason = "vector-reprojection-failed";
      } else {
        vectors = Object.freeze(
          reprojected as PdfSemanticAppearanceReprojectedVector[],
        );
        if (
          vectors.some(
            (vector) =>
              !boundsInsideOwner(vector.pageBounds, candidate!.geometry),
          )
        ) {
          reason = "vector-outside-owner";
        }
      }
    }
    const owner =
      candidate && (reason === null || reason === "shared-appearance-stream")
        ? matchedOwner(candidate)
        : null;
    const disposition: PdfSemanticAppearanceInstanceDisposition =
      reason === null
        ? "semantic-owner-eligible"
        : reason === "shared-appearance-stream" && owner
          ? "semantic-owner-preserve-only"
          : "preserve-only";
    const evidence = Object.freeze({
      instanceIndex,
      beginOperatorIndex: instance.beginOperatorIndex,
      endOperatorIndex: instance.endOperatorIndex,
      annotationRef: instance.payload.annotationRef,
      rect: instance.payload.rect,
      transform: instance.payload.transform,
      matrix: instance.payload.matrix,
      hasOwnCanvas: instance.payload.hasOwnCanvas,
      disposition,
      reason,
      owner,
      visibleOperatorIndexes: Object.freeze(
        visible.map((vector) => vector.operatorIndex),
      ),
      vectors,
    }) satisfies PdfSemanticAppearanceInstanceEvidence;
    instances.push(evidence);
    const vectorByIndex = new Map(
      vectors.map((vector) => [vector.visibleOperatorIndex, vector]),
    );
    for (const vector of visible) {
      const projected = vectorByIndex.get(vector.operatorIndex);
      vectorAssignments.push(
        Object.freeze({
          visibleOperatorIndex: vector.operatorIndex,
          instanceIndex,
          disposition,
          reason,
          owner,
          pageBounds: projected?.pageBounds ?? null,
          ...(projected?.pageSubpath
            ? { pageSubpath: projected.pageSubpath }
            : {}),
        }),
      );
    }
  }
  for (const vector of input.visibleVectors) {
    if (coveredVisible.has(vector.operatorIndex)) continue;
    vectorAssignments.push(
      Object.freeze({
        visibleOperatorIndex: vector.operatorIndex,
        instanceIndex: null,
        disposition: "preserve-only" as const,
        reason: "outside-annotation-block" as const,
        owner: null,
        pageBounds: null,
      }),
    );
  }
  instances.sort(
    (left, right) => left.beginOperatorIndex - right.beginOperatorIndex,
  );
  vectorAssignments.sort(
    (left, right) => left.visibleOperatorIndex - right.visibleOperatorIndex,
  );
  const matchedOwnerRecords = new Set(
    instances.map((instance) => instance.owner?.recordId).filter(Boolean),
  );
  return Object.freeze({
    schemaVersion: PDF_SEMANTIC_APPEARANCE_RECONCILIATION_SCHEMA_VERSION,
    sourceSha256,
    pageNumber: input.pageNumber,
    instances: Object.freeze(instances),
    vectorAssignments: Object.freeze(vectorAssignments),
    diagnostics: Object.freeze({
      instanceCount: instances.length,
      matchedInstanceCount: instances.filter((instance) => instance.owner)
        .length,
      eligibleInstanceCount: instances.filter(
        (instance) => instance.disposition === "semantic-owner-eligible",
      ).length,
      sharedPreserveInstanceCount: instances.filter(
        (instance) =>
          instance.disposition === "semantic-owner-preserve-only" &&
          instance.reason === "shared-appearance-stream",
      ).length,
      synthesizedInstanceCount: instances.filter(
        (instance) => instance.owner?.ownerKind === "annotation",
      ).length,
      invalidInstanceCount: instances.filter(
        (instance) => instance.disposition === "preserve-only",
      ).length,
      unmatchedOwnerCount: ownerCandidates.filter(
        (candidate) => !matchedOwnerRecords.has(candidate.recordId),
      ).length,
      eligibleVectorCount: vectorAssignments.filter(
        (assignment) => assignment.disposition === "semantic-owner-eligible",
      ).length,
      preserveOnlyVectorCount: vectorAssignments.filter(
        (assignment) => assignment.disposition !== "semantic-owner-eligible",
      ).length,
      outsideBlockVectorCount: vectorAssignments.filter(
        (assignment) => assignment.reason === "outside-annotation-block",
      ).length,
      ambiguousInstanceCount: instances.filter(
        (instance) =>
          instance.reason === "ambiguous-owner" ||
          instance.reason === "duplicate-instance",
      ).length,
    }),
  });
}
