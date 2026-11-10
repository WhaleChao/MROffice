/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { isPdfStandardBlendMode } from "./blend-modes";
import type { PdfStandardBlendMode } from "./blend-modes";
import { readPdfCommandOwnershipManifest } from "./engine/command-ownership";
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "./engine/contracts";
import type { PdfNativeVectorObjectStyle, PdfNativeVectorObjectStyleChanges, PdfNativeVectorRgbColor } from "./engine/native/vector-style-transaction";
import { MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION, MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION } from "./engine/native/vocabulary";
import type { PdfNativeVectorStyleSelectionEvidence } from "./semantic-appearance-inspection";
import type { PdfVectorPageGeometry } from "./vector-capability-reconciliation";
import type { PdfNativeVectorStyleSelectionEditPayload } from "./native-vector-style-selection-edit";

export const MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION =
  "pdf.vector.style.commit" as const;

export const PDF_NATIVE_VECTOR_STYLE_FEATURE_ID = "PDF-EDT-OBJECT-01" as const;

export const PDF_NATIVE_VECTOR_STYLE_CONTROL_ID =
  "PDFCTL.PAGE.VECTOR.STYLE" as const;

export const SHA256_PATTERN = /^[0-9a-f]{64}$/;

export const MAX_PROOF_ID_LENGTH = 2_048;

export const MAX_ABS_GEOMETRY = 10_000_000;

export function invalid(message: string): never {
  throw new Error(message);
}

export function plainRecord(value: unknown): Readonly<Record<string, unknown>> | null {
  try {
    if (!value || typeof value !== "object" || Array.isArray(value))
      return null;
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) return null;
    const descriptors = Object.getOwnPropertyDescriptors(value);
    const result: Record<string, unknown> = {};
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== "string") return null;
      const descriptor = descriptors[key];
      if (!descriptor?.enumerable || !("value" in descriptor)) return null;
      result[key] = descriptor.value;
    }
    return Object.freeze(result);
  } catch {
    return null;
  }
}

export function finite(value: unknown, label: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    Math.abs(value) > MAX_ABS_GEOMETRY
  ) {
    invalid(`Invalid PDF vector style ${label}.`);
  }
  return Object.is(value, -0) ? 0 : value;
}

export function integer(value: unknown, label: string, minimum = 0): number {
  if (!Number.isSafeInteger(value) || Number(value) < minimum) {
    invalid(`Invalid PDF vector style ${label}.`);
  }
  return Number(value);
}

export function sha(value: unknown, label: string): string {
  const normalized = typeof value === "string" ? value.toLowerCase() : "";
  if (!SHA256_PATTERN.test(normalized)) {
    invalid(`Invalid PDF vector style ${label}.`);
  }
  return normalized;
}

export function pageGeometry(value: unknown): PdfVectorPageGeometry {
  const record = plainRecord(value);
  const crop = plainRecord(record?.cropBox);
  if (
    !record ||
    !crop ||
    (record.rotation !== 0 &&
      record.rotation !== 90 &&
      record.rotation !== 180 &&
      record.rotation !== 270)
  ) {
    invalid("Invalid PDF vector style page geometry.");
  }
  const result = Object.freeze({
    cropBox: Object.freeze({
      xMin: finite(crop.xMin, "CropBox"),
      yMin: finite(crop.yMin, "CropBox"),
      xMax: finite(crop.xMax, "CropBox"),
      yMax: finite(crop.yMax, "CropBox"),
    }),
    rotation: record.rotation,
  });
  if (
    result.cropBox.xMax <= result.cropBox.xMin ||
    result.cropBox.yMax <= result.cropBox.yMin
  ) {
    invalid("Invalid PDF vector style CropBox.");
  }
  return result;
}

export function color(value: unknown, label: string): PdfNativeVectorRgbColor {
  const record = plainRecord(value);
  if (!record) invalid(`Invalid PDF vector style ${label} color.`);
  const result = Object.freeze({
    r: finite(record.r, `${label} color`),
    g: finite(record.g, `${label} color`),
    b: finite(record.b, `${label} color`),
  });
  if (
    [result.r, result.g, result.b].some(
      (component) => component < 0 || component > 1,
    )
  ) {
    invalid(`Invalid PDF vector style ${label} color.`);
  }
  return result;
}

export function changes(value: unknown): PdfNativeVectorObjectStyleChanges {
  const record = plainRecord(value);
  if (!record || Object.keys(record).length === 0) {
    invalid("Invalid PDF vector style changes.");
  }
  const allowed = new Set([
    "fillColor",
    "strokeColor",
    "lineWidth",
    "dashArray",
    "dashPhase",
    "lineCap",
    "lineJoin",
    "miterLimit",
    "opacity",
    "blendMode",
    "paintMode",
    "translate",
    "transform",
  ]);
  if (Object.keys(record).some((key) => !allowed.has(key))) {
    invalid("Invalid PDF vector style changes.");
  }
  const result: {
    fillColor?: PdfNativeVectorRgbColor;
    strokeColor?: PdfNativeVectorRgbColor;
    lineWidth?: number;
    dashArray?: readonly number[];
    dashPhase?: number;
    lineCap?: "butt" | "round" | "square";
    lineJoin?: "miter" | "round" | "bevel";
    miterLimit?: number;
    opacity?: number;
    blendMode?: PdfStandardBlendMode;
    paintMode?: "stroke" | "fill" | "fill-stroke";
    translate?: Readonly<{ x: number; y: number }>;
    transform?: Readonly<{ a: number; b: number; c: number; d: number; e: number; f: number }>;
  } = {};
  if (Object.hasOwn(record, "transform")) {
    // A resize of the whole object (its stroke scales with it): it comes
    // alone and must be invertible.
    const affine = plainRecord(record.transform);
    if (!affine || Object.keys(record).length !== 1) {
      invalid("Invalid PDF vector style transform.");
    }
    const transform = Object.freeze({
      a: finite(affine.a, "transform"),
      b: finite(affine.b, "transform"),
      c: finite(affine.c, "transform"),
      d: finite(affine.d, "transform"),
      e: finite(affine.e, "transform"),
      f: finite(affine.f, "transform"),
    });
    if (Math.abs(transform.a * transform.d - transform.b * transform.c) <= 1e-12) {
      invalid("Invalid PDF vector style transform.");
    }
    result.transform = transform;
  }
  if (Object.hasOwn(record, "translate")) {
    const shift = plainRecord(record.translate);
    if (!shift || !Object.hasOwn(record, "lineWidth")) {
      invalid("Invalid PDF vector style translation.");
    }
    result.translate = Object.freeze({
      x: finite(shift.x, "translation"),
      y: finite(shift.y, "translation"),
    });
  }
  if (Object.hasOwn(record, "paintMode")) {
    if (
      record.paintMode !== "stroke" &&
      record.paintMode !== "fill" &&
      record.paintMode !== "fill-stroke"
    ) {
      invalid("Invalid PDF vector style paint mode.");
    }
    result.paintMode = record.paintMode;
  }
  if (Object.hasOwn(record, "fillColor"))
    result.fillColor = color(record.fillColor, "fill");
  if (Object.hasOwn(record, "strokeColor")) {
    result.strokeColor = color(record.strokeColor, "stroke");
  }
  if (Object.hasOwn(record, "lineWidth")) {
    const value = finite(record.lineWidth, "line width");
    if (value <= 0 || value > 1_000)
      invalid("Invalid PDF vector style line width.");
    result.lineWidth = value;
  }
  if (Object.hasOwn(record, "dashArray")) {
    const value = record.dashArray;
    if (
      !Array.isArray(value) ||
      value.length > 64 ||
      value.some((entry) => finite(entry, "dash array") < 0) ||
      (value.length > 0 && value.every((entry) => entry === 0))
    ) {
      invalid("Invalid PDF vector style dash array.");
    }
    result.dashArray = Object.freeze(value.map((entry) => Number(entry)));
  }
  if (Object.hasOwn(record, "dashPhase")) {
    result.dashPhase = finite(record.dashPhase, "dash phase");
  }
  if (Object.hasOwn(record, "lineCap")) {
    if (
      record.lineCap !== "butt" &&
      record.lineCap !== "round" &&
      record.lineCap !== "square"
    ) {
      invalid("Invalid PDF vector style line cap.");
    }
    result.lineCap = record.lineCap;
  }
  if (Object.hasOwn(record, "lineJoin")) {
    if (
      record.lineJoin !== "miter" &&
      record.lineJoin !== "round" &&
      record.lineJoin !== "bevel"
    ) {
      invalid("Invalid PDF vector style line join.");
    }
    result.lineJoin = record.lineJoin;
  }
  if (Object.hasOwn(record, "miterLimit")) {
    const value = finite(record.miterLimit, "miter limit");
    if (value < 1) invalid("Invalid PDF vector style miter limit.");
    result.miterLimit = value;
  }
  if (Object.hasOwn(record, "opacity")) {
    const value = finite(record.opacity, "opacity");
    if (value < 0 || value > 1) {
      invalid("Invalid PDF vector style opacity.");
    }
    result.opacity = value;
  }
  if (Object.hasOwn(record, "blendMode")) {
    if (!isPdfStandardBlendMode(record.blendMode)) {
      invalid("Invalid PDF vector style blend mode.");
    }
    result.blendMode = record.blendMode;
  }
  return Object.freeze(result);
}

export function styleEvidence(value: unknown): PdfNativeVectorStyleSelectionEvidence {
  const record = plainRecord(value);
  const proofId = typeof record?.proofId === "string" ? record.proofId : "";
  if (
    !record ||
    record.schemaVersion !== MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION ||
    proofId.length === 0 ||
    proofId.length > MAX_PROOF_ID_LENGTH
  ) {
    invalid("Invalid PDF vector style selection evidence.");
  }
  const style = plainRecord(record.style);
  if (!style) invalid("Invalid PDF vector style selection evidence.");
  const dashArray = style.dashArray;
  if (
    dashArray !== null &&
    (!Array.isArray(dashArray) ||
      dashArray.length > 64 ||
      dashArray.some((entry) => finite(entry, "dash array") < 0) ||
      (dashArray.length > 0 && dashArray.every((entry) => entry === 0)))
  ) {
    invalid("Invalid PDF vector style selection evidence.");
  }
  const lineCap =
    style.lineCap === null ||
    style.lineCap === "butt" ||
    style.lineCap === "round" ||
    style.lineCap === "square"
      ? style.lineCap
      : invalid("Invalid PDF vector style line cap.");
  const lineJoin =
    style.lineJoin === null ||
    style.lineJoin === "miter" ||
    style.lineJoin === "round" ||
    style.lineJoin === "bevel"
      ? style.lineJoin
      : invalid("Invalid PDF vector style line join.");
  if (typeof style.hasExternalGraphicsState !== "boolean") {
    invalid("Invalid PDF vector style graphics-state evidence.");
  }
  const lineWidth =
    style.lineWidth === null ? null : finite(style.lineWidth, "line width");
  const miterLimit =
    style.miterLimit === null ? null : finite(style.miterLimit, "miter limit");
  const opacity =
    style.opacity === null ? null : finite(style.opacity, "opacity");
  const blendMode =
    style.blendMode === null || isPdfStandardBlendMode(style.blendMode)
      ? (style.blendMode as PdfStandardBlendMode | null)
      : invalid("Invalid PDF vector style blend mode.");
  if (
    (lineWidth !== null && (lineWidth < 0 || lineWidth > 1_000)) ||
    (miterLimit !== null && miterLimit < 1) ||
    (opacity !== null && (opacity < 0 || opacity > 1))
  ) {
    invalid("Invalid PDF vector style selection evidence.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
    proofId,
    proofDigest: sha(record.proofDigest, "proof digest"),
    pageProofIndex: integer(record.pageProofIndex, "proof index"),
    style: Object.freeze({
      fillColor:
        style.fillColor === null ? null : color(style.fillColor, "fill"),
      strokeColor:
        style.strokeColor === null ? null : color(style.strokeColor, "stroke"),
      lineWidth,
      dashArray:
        dashArray === null
          ? null
          : Object.freeze(
              dashArray.map((entry) => finite(entry, "dash array")),
            ),
      dashPhase:
        style.dashPhase === null ? null : finite(style.dashPhase, "dash phase"),
      lineCap,
      lineJoin,
      miterLimit,
      opacity,
      blendMode,
      hasExternalGraphicsState: style.hasExternalGraphicsState,
      ...(style.fillUnreadable === true ? { fillUnreadable: true } : {}),
    }),
  });
}

export function route(
  value: unknown,
): PdfNativeVectorStyleSelectionEditPayload["route"] {
  const record = plainRecord(value);
  const bounds = plainRecord(record?.bounds);
  if (
    !record ||
    !bounds ||
    record.selectionOnly !== true ||
    (record.proofFamily !== "shape" && record.proofFamily !== "subpath") ||
    (record.nativeKind !== "line" &&
      record.nativeKind !== "rectangle" &&
      record.nativeKind !== "ellipse" &&
      record.nativeKind !== "subpath") ||
    (record.paintMode !== "stroke" &&
      record.paintMode !== "fill" &&
      record.paintMode !== "fill-stroke")
  ) {
    invalid("Invalid PDF vector style route.");
  }
  return Object.freeze({
    selectionOnly: true,
    sourceSha256: sha(record.sourceSha256, "source SHA-256"),
    pageNumber: integer(record.pageNumber, "page number", 1),
    rendererOperatorIndex: integer(
      record.rendererOperatorIndex,
      "operator index",
    ),
    proofFamily: record.proofFamily,
    nativeKind: record.nativeKind,
    paintMode: record.paintMode,
    proofId:
      typeof record.proofId === "string" &&
      record.proofId.length <= MAX_PROOF_ID_LENGTH
        ? record.proofId
        : invalid("Invalid PDF vector style route proof."),
    proofDigest: sha(record.proofDigest, "route proof digest"),
    contentStreamObjectNumber: integer(
      record.contentStreamObjectNumber,
      "content stream",
      1,
    ),
    contentStreamGeneration: integer(
      record.contentStreamGeneration,
      "content stream generation",
    ),
    operatorStart: integer(record.operatorStart, "operator start"),
    operatorEnd: integer(record.operatorEnd, "operator end", 1),
    style: styleEvidence(record.style),
    bounds: Object.freeze({
      x: finite(bounds.x, "bounds"),
      y: finite(bounds.y, "bounds"),
      width: finite(bounds.width, "bounds"),
      height: finite(bounds.height, "bounds"),
    }),
  });
}

export function readPdfNativeVectorStyleSelectionEditPayload(
  value: unknown,
): PdfNativeVectorStyleSelectionEditPayload {
  const record = plainRecord(value);
  if (
    !record ||
    record.schemaVersion !==
      MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION
  ) {
    invalid("Unsupported PDF vector style edit payload.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION,
    route: route(record.route),
    pageGeometry: pageGeometry(record.pageGeometry),
    changes: changes(record.changes),
  });
}

export function samePdfNativeVectorObjectStyle(
  left: PdfNativeVectorObjectStyle,
  right: PdfNativeVectorObjectStyle,
): boolean {
  // (The external graphics state and an unreadable fill are not what the
  // reader sees; an absent `fillUnreadable` reads as false.)
  const visible = ({ hasExternalGraphicsState: _external, fillUnreadable: _unreadable, ...rest }:
    PdfNativeVectorObjectStyle) => JSON.stringify(rest);
  return visible(left) === visible(right);
}

export function createPdfNativeVectorStyleOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_VECTOR_STYLE_FEATURE_ID,
      action: MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function buildPdfNativeVectorStyleCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativeVectorStyleSelectionEditPayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativeVectorStyleSelectionEditPayload(input.payload);
  if (
    payload.route.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() ||
    !Number.isSafeInteger(input.sessionRevision) ||
    input.sessionRevision < 0 ||
    !Number.isSafeInteger(input.selectionRevision) ||
    input.selectionRevision < 1
  ) {
    invalid("PDF vector style command authority is stale.");
  }
  return Object.freeze({
    ...input.source,
    featureId: PDF_NATIVE_VECTOR_STYLE_FEATURE_ID,
    controlId: PDF_NATIVE_VECTOR_STYLE_CONTROL_ID,
    action: MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
}
