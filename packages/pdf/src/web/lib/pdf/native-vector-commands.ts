/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { readPdfCommandOwnershipManifest } from "./engine/command-ownership";
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "./engine/contracts";
import type { PdfNativeListTextReplayAuthority } from "./engine/native/list-text-command-lifecycle";
import { MRAI_NATIVE_SOURCE_CONSOLIDATION_ACTION, MRAI_NATIVE_SOURCE_CONSOLIDATION_KIND } from "./engine/native/source-consolidation";
import type { PdfNativeVectorAffineTransform, PdfNativeVectorBounds, PdfNativeVectorGeometry, PdfNativeVectorSubpathCommand, PdfNativeVectorSubpathProof } from "./engine/native/vector-content-transaction";
import { PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN, pdfNativeVectorProofId } from "./engine/native/vector-identity";
import { MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION } from "./engine/native/vocabulary";
import type { PdfNativeVectorSelectionRoute } from "./semantic-appearance-inspection";
import type { PdfVectorPageGeometry } from "./vector-capability-reconciliation";
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import type { PdfNativeRectangleGeometry, PdfNativeVectorAcceptedTarget, PdfNativeVectorDeletePayload, PdfNativeVectorDuplicatePayload, PdfNativeVectorHistoryEndpoint, PdfNativeVectorNormalizedBounds, PdfNativeVectorReorderPayload, PdfNativeVectorReorderRefusal, PdfNativeVectorRotationHistory, PdfNativeVectorRotationOrigin, PdfNativeVectorOutlineReplace, PdfNativeVectorSelectionEditPayload, PdfNativeVectorSelectionEditPlan, PdfNativeVectorVertexEdit, PdfNativeVectorWriteSelection } from "./native-vector-selection-edit";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION =
  "pdf.vector.geometry.commit" as const;

export const MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION =
  "mrai.pdf-native-vector-delete.v1" as const;

export const MRAI_NATIVE_VECTOR_DELETE_ACTION =
  "pdf.vector.delete.commit" as const;

export const PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID =
  "PDF-EDT-OBJECT-01" as const;

export const PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID =
  "PDFCTL.PAGE.VECTOR.GEOMETRY" as const;

export const PDF_NATIVE_VECTOR_DELETE_CONTROL_ID =
  PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID;

export const MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION =
  "mrai.pdf-native-vector-duplicate.v1" as const;

export const MRAI_NATIVE_VECTOR_DUPLICATE_ACTION =
  "pdf.vector.duplicate.commit" as const;

export const MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION =
  "mrai.pdf-native-vector-reorder.v1" as const;

export const MRAI_NATIVE_VECTOR_REORDER_ACTION =
  "pdf.vector.reorder.commit" as const;

export const SHA256_PATTERN = /^[0-9a-f]{64}$/;

export const MAX_ABS_COORDINATE = 10_000_000;

export const MAX_PROOF_ID_LENGTH = 2_048;

export const PAGE_EPSILON = 0.000_001;

export const GEOMETRY_EPSILON = 0.000_01;

export const AFFINE_DETERMINANT_EPSILON = 1e-12;

export const MAX_ACKNOWLEDGED_NATIVE_TRANSACTIONS = 31;

/** Display-space nudge (points) applied to a duplicated vector, like WPS;
 * a paste with the pointer on no page steps by it too. */
export const PDF_NATIVE_VECTOR_DUPLICATE_NUDGE_POINTS = 12;

/** The planner's refusal for a shape drawn wholly outside its page's CropBox
 * (an imposed sheet's other windows): not this page's to edit. */
export const PDF_NATIVE_VECTOR_OFF_PAGE_MESSAGE =
  "PDF vector is drawn outside its page.";

/** The planner's refusal for an edit that would take a shape wholly off its
 * page: past the edge is allowed (the page cuts it), wholly off is not. */
export const PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE =
  "PDF vector edit would move geometry outside its CropBox.";

export function pdfNativeVectorReorderRefusal(
  error: unknown,
): PdfNativeVectorReorderRefusal | null {
  if (!(error instanceof Error)) return null;
  switch (error.message) {
    case "PDF vector is already the topmost path on its page.":
      return "at-front";
    case "PDF vector is already the bottommost path on its page.":
      return "at-back";
    case "PDF vector neighbours are separated by unsupported content.":
      return "other-content";
    case "PDF vector cannot pass another of its page's content streams.":
      return "other-stream";
    case "PDF vector neighbour lives in a shared stream.":
      return "shared-neighbour";
    case PDF_NATIVE_VECTOR_OFF_PAGE_MESSAGE:
      return "off-page";
    default:
      return null;
  }
}

/** The smallest box around every member's bounds (page fractions). */
export function pdfNativeVectorGroupBounds(
  members: readonly Readonly<{ bounds: PdfNativeVectorNormalizedBounds }>[],
): PdfNativeVectorNormalizedBounds | null {
  if (members.length === 0) return null;
  let xMin = Infinity, yMin = Infinity, xMax = -Infinity, yMax = -Infinity;
  for (const { bounds } of members) {
    xMin = Math.min(xMin, bounds.x);
    yMin = Math.min(yMin, bounds.y);
    xMax = Math.max(xMax, bounds.x + bounds.width);
    yMax = Math.max(yMax, bounds.y + bounds.height);
  }
  return Object.freeze({ x: xMin, y: yMin, width: xMax - xMin, height: yMax - yMin });
}

export function pdfNativeVectorAcceptedTarget(
  selection: PdfNativeVectorWriteSelection,
): PdfNativeVectorAcceptedTarget {
  return Object.freeze({
    sourceSha256: selection.sourceSha256,
    pageNumber: selection.pageNumber,
    proofFamily: selection.proofFamily,
    nativeKind: selection.nativeKind,
    paintMode: selection.paintMode,
    proofId: selection.proofId,
    proofDigest: selection.proofDigest,
    contentStreamObjectNumber: selection.contentStreamObjectNumber,
    contentStreamGeneration: selection.contentStreamGeneration,
    operatorStart: selection.operatorStart,
    operatorEnd: selection.operatorEnd,
    ...(selection.formInstance === undefined
      ? {}
      : { formInstance: selection.formInstance }),
  });
}

/** Same proof identity the writer uses to resolve a route (see
 * `resolvePdfNativeVectorDisplayRoute`), plus the revision it was proved in. */
export function pdfNativeVectorRouteIsAcceptedTarget(
  route: PdfNativeVectorSelectionRoute,
  target: PdfNativeVectorAcceptedTarget,
): boolean {
  return (
    route.selectionOnly === true &&
    route.sourceSha256 === target.sourceSha256 &&
    route.pageNumber === target.pageNumber &&
    route.proofFamily === target.proofFamily &&
    route.nativeKind === target.nativeKind &&
    route.paintMode === target.paintMode &&
    route.proofId === target.proofId &&
    route.proofDigest === target.proofDigest &&
    route.contentStreamObjectNumber === target.contentStreamObjectNumber &&
    route.contentStreamGeneration === target.contentStreamGeneration &&
    route.operatorStart === target.operatorStart &&
    route.operatorEnd === target.operatorEnd &&
    (route.formInstance ?? null) === (target.formInstance ?? null)
  );
}

export function fail(message: string): never {
  throw new Error(message);
}

export function rounded(value: number): number {
  const result = Number(value.toFixed(6));
  return Object.is(result, -0) ? 0 : result;
}

export function exactFinite(value: unknown, label: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    Math.abs(value) > MAX_ABS_COORDINATE
  ) {
    fail(`Invalid PDF vector ${label}.`);
  }
  return value;
}

export function exactInteger(value: unknown, label: string, minimum = 0): number {
  if (!Number.isSafeInteger(value) || Number(value) < minimum) {
    fail(`Invalid PDF vector ${label}.`);
  }
  return Number(value);
}

export function exactSha256(value: unknown, label: string): string {
  if (typeof value !== "string" || !SHA256_PATTERN.test(value)) {
    fail(`Invalid PDF vector ${label} SHA-256.`);
  }
  return value;
}

export function exactPoint(value: unknown, label: string) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail(`Invalid PDF vector ${label}.`);
  }
  const point = value as { x?: unknown; y?: unknown };
  return Object.freeze({
    x: exactFinite(point.x, `${label} x`),
    y: exactFinite(point.y, `${label} y`),
  });
}

export function exactPageGeometry(value: unknown): PdfVectorPageGeometry {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid PDF vector page geometry.");
  }
  const geometry = value as {
    cropBox?: {
      xMin?: unknown;
      yMin?: unknown;
      xMax?: unknown;
      yMax?: unknown;
    };
    rotation?: unknown;
  };
  const crop = geometry.cropBox;
  if (!crop || typeof crop !== "object" || Array.isArray(crop)) {
    fail("Invalid PDF vector CropBox.");
  }
  const xMin = exactFinite(crop.xMin, "CropBox xMin");
  const yMin = exactFinite(crop.yMin, "CropBox yMin");
  const xMax = exactFinite(crop.xMax, "CropBox xMax");
  const yMax = exactFinite(crop.yMax, "CropBox yMax");
  const rotation = exactInteger(geometry.rotation, "page rotation");
  if (xMax <= xMin || yMax <= yMin || ![0, 90, 180, 270].includes(rotation)) {
    fail("Invalid PDF vector page geometry.");
  }
  return Object.freeze({
    cropBox: Object.freeze({ xMin, yMin, xMax, yMax }),
    rotation: rotation as PdfVectorPageGeometry["rotation"],
  });
}

/** `allowIdentity`: Merge Shapes' record, whose outline -- not an affine --
 * is the change (the merged shape may stand in the first one's very box). */
export function exactAffine(value: unknown, allowIdentity = false): PdfNativeVectorAffineTransform {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid PDF vector affine transform.");
  }
  const candidate = value as Record<
    keyof PdfNativeVectorAffineTransform,
    unknown
  >;
  const affine = Object.freeze({
    a: exactFinite(candidate.a, "affine a"),
    b: exactFinite(candidate.b, "affine b"),
    c: exactFinite(candidate.c, "affine c"),
    d: exactFinite(candidate.d, "affine d"),
    e: exactFinite(candidate.e, "affine e"),
    f: exactFinite(candidate.f, "affine f"),
  });
  const determinant = affine.a * affine.d - affine.b * affine.c;
  if (
    !Number.isFinite(determinant) ||
    Math.abs(determinant) <= AFFINE_DETERMINANT_EPSILON ||
    (!allowIdentity && [affine.a - 1, affine.b, affine.c, affine.d - 1, affine.e, affine.f].every(
      (entry) => Math.abs(entry) <= GEOMETRY_EPSILON,
    ))
  ) {
    fail("PDF vector affine transform is singular or unchanged.");
  }
  return affine;
}

export function exactNativeBounds(value: unknown): PdfNativeVectorBounds {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid PDF vector native bounds.");
  }
  const bounds = value as Record<keyof PdfNativeVectorBounds, unknown>;
  const result = Object.freeze({
    x: exactFinite(bounds.x, "native bounds x"),
    y: exactFinite(bounds.y, "native bounds y"),
    width: exactFinite(bounds.width, "native bounds width"),
    height: exactFinite(bounds.height, "native bounds height"),
  });
  if (
    result.width < 0 ||
    result.height < 0 ||
    Math.max(result.width, result.height) <= GEOMETRY_EPSILON
  ) {
    fail("PDF vector native bounds are degenerate.");
  }
  return result;
}

export function exactGeometry(
  value: unknown,
  expectedKind?: "line" | "rectangle" | "ellipse",
): PdfNativeVectorGeometry {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid PDF vector geometry.");
  }
  const geometry = value as Record<string, unknown>;
  const kind = geometry.kind;
  if (expectedKind && kind !== expectedKind) {
    fail("PDF vector geometry kind changed.");
  }
  if (kind === "line") {
    const start = exactPoint(geometry.start, "line start");
    const end = exactPoint(geometry.end, "line end");
    if (Math.hypot(end.x - start.x, end.y - start.y) <= GEOMETRY_EPSILON) {
      fail("PDF vector line is empty.");
    }
    return Object.freeze({ kind, start, end });
  }
  if (kind === "rectangle") {
    if (!Array.isArray(geometry.corners) || geometry.corners.length !== 4) {
      fail("Invalid PDF vector rectangle.");
    }
    const corners = geometry.corners.map((entry, index) =>
      exactPoint(entry, `rectangle corner ${index + 1}`),
    ) as unknown as PdfNativeRectangleGeometry["corners"];
    return Object.freeze({ kind, corners: Object.freeze(corners) });
  }
  if (kind !== "ellipse") fail("Invalid PDF vector geometry kind.");
  return Object.freeze({
    kind,
    center: exactPoint(geometry.center, "ellipse center"),
    axisX: exactPoint(geometry.axisX, "ellipse x-axis"),
    axisY: exactPoint(geometry.axisY, "ellipse y-axis"),
  });
}

export function exactCommands(
  value: unknown,
): readonly PdfNativeVectorSubpathCommand[] {
  if (!Array.isArray(value) || value.length < 2 || value.length > 4096) {
    fail("Invalid PDF vector subpath commands.");
  }
  // One path of one or more subpaths: each opens with a move after the
  // subpath before it drew, closes with `h` at most once, and a segment
  // straight after a close opens the next subpath at the closed one's start
  // (mirror of the engine's scanner).
  let closed = false;
  let segments = 0;
  const commands = value.map((entry, index): PdfNativeVectorSubpathCommand => {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
      fail("Invalid PDF vector subpath command.");
    }
    const command = entry as Record<string, unknown>;
    if (command.operator === "m") {
      if ((index === 0) !== (segments === 0 && !closed) || (index !== 0 && segments === 0 && !closed)) {
        fail("Invalid PDF vector subpath move topology.");
      }
      closed = false;
      segments = 0;
      return Object.freeze({
        operator: command.operator,
        point: exactPoint(command.point, `${command.operator} point`),
      });
    }
    if (command.operator === "l") {
      if (index === 0) {
        fail("Invalid PDF vector subpath move topology.");
      }
      closed = false;
      segments += 1;
      return Object.freeze({
        operator: command.operator,
        point: exactPoint(command.point, `${command.operator} point`),
      });
    }
    if (command.operator === "h") {
      if (closed) {
        fail("Invalid PDF vector subpath close topology.");
      }
      closed = true;
      return Object.freeze({ operator: "h" });
    }
    if (
      command.operator !== "c" &&
      command.operator !== "v" &&
      command.operator !== "y"
    ) {
      fail("Invalid PDF vector subpath operator.");
    }
    if (index === 0) fail("Invalid PDF vector curve topology.");
    closed = false;
    segments += 1;
    return Object.freeze({
      operator: command.operator,
      control1: exactPoint(command.control1, `${command.operator} control 1`),
      control2: exactPoint(command.control2, `${command.operator} control 2`),
      end: exactPoint(command.end, `${command.operator} end`),
    });
  });
  if (segments === 0 && !closed) fail("Invalid PDF vector subpath move topology.");
  return Object.freeze(commands);
}

export function canonicalGeometry(geometry: PdfNativeVectorGeometry): unknown {
  if (geometry.kind === "line") {
    return {
      kind: geometry.kind,
      start: { x: rounded(geometry.start.x), y: rounded(geometry.start.y) },
      end: { x: rounded(geometry.end.x), y: rounded(geometry.end.y) },
    };
  }
  if (geometry.kind === "rectangle") {
    return {
      kind: geometry.kind,
      corners: geometry.corners.map((point) => ({
        x: rounded(point.x),
        y: rounded(point.y),
      })),
    };
  }
  return {
    kind: geometry.kind,
    center: { x: rounded(geometry.center.x), y: rounded(geometry.center.y) },
    axisX: { x: rounded(geometry.axisX.x), y: rounded(geometry.axisX.y) },
    axisY: { x: rounded(geometry.axisY.x), y: rounded(geometry.axisY.y) },
  };
}

export function canonicalCommands(commands: readonly PdfNativeVectorSubpathCommand[]) {
  return commands.map((command) => {
    if (command.operator === "m" || command.operator === "l") {
      return {
        operator: command.operator,
        point: { x: rounded(command.point.x), y: rounded(command.point.y) },
      };
    }
    if (command.operator === "h") return { operator: command.operator };
    if (!("control1" in command)) {
      fail("Invalid PDF vector curve command.");
    }
    return {
      operator: command.operator,
      control1: {
        x: rounded(command.control1.x),
        y: rounded(command.control1.y),
      },
      control2: {
        x: rounded(command.control2.x),
        y: rounded(command.control2.y),
      },
      end: { x: rounded(command.end.x), y: rounded(command.end.y) },
    };
  });
}

export function geometryState(selection: PdfNativeVectorWriteSelection): string {
  return JSON.stringify(
    selection.proofFamily === "shape"
      ? {
          proofFamily: selection.proofFamily,
          nativeKind: selection.nativeKind,
          paintMode: selection.paintMode,
          geometry: canonicalGeometry(selection.geometry),
        }
      : {
          proofFamily: selection.proofFamily,
          nativeKind: selection.nativeKind,
          paintMode: selection.paintMode,
          paintOperator: selection.paintOperator,
          commands: canonicalCommands(selection.commands),
          closed: selection.closed,
        },
  );
}

export function exactSelection(value: unknown): PdfNativeVectorWriteSelection {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid PDF vector selection authority.");
  }
  const selection = value as Record<string, unknown>;
  const proofFamily = selection.proofFamily;
  const sourceSha256 = exactSha256(selection.sourceSha256, "selection source");
  const pageNumber = exactInteger(selection.pageNumber, "selection page", 1);
  const pageProofIndex = exactInteger(selection.pageProofIndex, "proof index");
  const proofId =
    typeof selection.proofId === "string" ? selection.proofId : "";
  if (
    proofId.length < 1 ||
    proofId.length > MAX_PROOF_ID_LENGTH ||
    proofId.trim() !== proofId
  ) {
    fail("Invalid PDF vector proof id.");
  }
  const proofDigest = exactSha256(selection.proofDigest, "proof digest");
  const contentStreamObjectNumber = exactInteger(
    selection.contentStreamObjectNumber,
    "content stream object",
    1,
  );
  const contentStreamGeneration = exactInteger(
    selection.contentStreamGeneration,
    "content stream generation",
  );
  const operatorStart = exactInteger(selection.operatorStart, "operator start");
  const operatorEnd = exactInteger(selection.operatorEnd, "operator end", 1);
  const operatorSha256 = exactSha256(selection.operatorSha256, "operator span");
  const bounds = exactNativeBounds(selection.bounds);
  if (operatorEnd <= operatorStart) fail("Invalid PDF vector operator span.");
  const formInstance = selection.formInstance;
  if (
    formInstance !== undefined &&
    (typeof formInstance !== "string" ||
      !PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN.test(formInstance))
  ) {
    fail("Invalid PDF vector Form drawing.");
  }
  const expectedId = pdfNativeVectorProofId({
    family: proofFamily === "subpath" ? "subpath" : "shape",
    sourceSha256,
    pageNumber,
    objectNumber: contentStreamObjectNumber,
    generation: contentStreamGeneration,
    start: operatorStart,
    end: operatorEnd,
    formInstance,
  });
  if (proofId !== expectedId) fail("PDF vector proof id is inconsistent.");
  const common = {
    sourceSha256,
    pageNumber,
    pageProofIndex,
    proofId,
    proofDigest,
    contentStreamObjectNumber,
    contentStreamGeneration,
    operatorStart,
    operatorEnd,
    operatorSha256,
    bounds,
    ...(formInstance !== undefined ? { formInstance } : {}),
  } as const;
  if (proofFamily === "shape") {
    if (
      selection.nativeKind !== "line" &&
      selection.nativeKind !== "rectangle" &&
      selection.nativeKind !== "ellipse"
    ) {
      fail("Invalid PDF vector shape kind.");
    }
    if (
      selection.paintMode !== "stroke" &&
      selection.paintMode !== "fill" &&
      selection.paintMode !== "fill-stroke"
    ) {
      fail("Invalid PDF vector paint mode.");
    }
    return Object.freeze({
      proofFamily,
      ...common,
      nativeKind: selection.nativeKind,
      paintMode: selection.paintMode,
      geometry: exactGeometry(selection.geometry, selection.nativeKind),
    });
  }
  if (
    proofFamily !== "subpath" ||
    selection.nativeKind !== "subpath" ||
    (selection.paintMode !== "stroke" &&
      selection.paintMode !== "fill" &&
      selection.paintMode !== "fill-stroke") ||
    !["S", "s", "f", "F", "f*", "B", "B*", "b", "b*"].includes(
      String(selection.paintOperator),
    ) ||
    typeof selection.closed !== "boolean"
  ) {
    fail("Invalid PDF vector subpath authority.");
  }
  return Object.freeze({
    proofFamily,
    ...common,
    nativeKind: "subpath",
    paintMode: selection.paintMode,
    paintOperator:
      selection.paintOperator as PdfNativeVectorSubpathProof["paintOperator"],
    commands: exactCommands(selection.commands),
    closed: selection.closed,
  });
}

export function nativeBoundsMatch(
  left: PdfNativeVectorBounds,
  right: PdfNativeVectorBounds,
  tolerance = PAGE_EPSILON,
): boolean {
  return (Object.keys(left) as Array<keyof PdfNativeVectorBounds>).every(
    (key) => Math.abs(left[key] - right[key]) <= tolerance,
  );
}

export function historyEndpoint(
  selection: PdfNativeVectorWriteSelection,
): PdfNativeVectorHistoryEndpoint {
  const {
    sourceSha256: _sourceSha256,
    proofId: _proofId,
    proofDigest: _proofDigest,
    operatorStart: _operatorStart,
    operatorEnd: _operatorEnd,
    operatorSha256: _operatorSha256,
    ...endpoint
  } = selection;
  return Object.freeze(endpoint);
}

export function createPdfNativeVectorHistoryEndpoint(
  selection: PdfNativeVectorWriteSelection,
): PdfNativeVectorHistoryEndpoint {
  return historyEndpoint(selection);
}

















export function pdfNativeVectorHistoryEndpointOnStreamOf(
  endpoint: PdfNativeVectorHistoryEndpoint,
  current: PdfNativeVectorHistoryEndpoint,
): PdfNativeVectorHistoryEndpoint {
  if (
    endpoint.contentStreamObjectNumber === current.contentStreamObjectNumber &&
    endpoint.contentStreamGeneration === current.contentStreamGeneration &&
    endpoint.formInstance === current.formInstance
  ) {
    return endpoint;
  }
  const rebased: Record<string, unknown> = {
    ...endpoint,
    contentStreamObjectNumber: current.contentStreamObjectNumber,
    contentStreamGeneration: current.contentStreamGeneration,
  };
  if (current.formInstance !== undefined) {
    rebased.formInstance = current.formInstance;
  } else {
    delete rebased.formInstance;
  }
  return Object.freeze(rebased) as unknown as PdfNativeVectorHistoryEndpoint;
}

export function samePdfNativeVectorHistoryEndpoint(
  left: PdfNativeVectorHistoryEndpoint,
  right: PdfNativeVectorHistoryEndpoint,
): boolean {
  return (
    left.proofFamily === right.proofFamily &&
    left.pageNumber === right.pageNumber &&
    left.pageProofIndex === right.pageProofIndex &&
    left.nativeKind === right.nativeKind &&
    left.paintMode === right.paintMode &&
    left.contentStreamObjectNumber === right.contentStreamObjectNumber &&
    left.contentStreamGeneration === right.contentStreamGeneration &&
    left.formInstance === right.formInstance &&
    nativeBoundsMatch(left.bounds, right.bounds, HISTORY_POINT_TOLERANCE) &&
    sameNumbersWithin(
      JSON.parse(geometryState(left as PdfNativeVectorWriteSelection)),
      JSON.parse(geometryState(right as PdfNativeVectorWriteSelection)),
      HISTORY_POINT_TOLERANCE,
    )
  );
}






const HISTORY_POINT_TOLERANCE = 0.000_1;

function sameNumbersWithin(left: unknown, right: unknown, tolerance: number): boolean {
  if (typeof left === "number" && typeof right === "number") return Math.abs(left - right) <= tolerance;
  if (Array.isArray(left) && Array.isArray(right)) {
    return left.length === right.length && left.every((value, index) => sameNumbersWithin(value, right[index], tolerance));
  }
  if (left && right && typeof left === "object" && typeof right === "object" && !Array.isArray(left) && !Array.isArray(right)) {
    const leftKeys = Object.keys(left);
    const rightRecord = right as Record<string, unknown>;
    return leftKeys.length === Object.keys(right).length && leftKeys.every((key) =>
      key in rightRecord && sameNumbersWithin((left as Record<string, unknown>)[key], rightRecord[key], tolerance));
  }
  return left === right;
}

/** Thickness (points) under which a selection's axis has no extent. */
export const LINE_AXIS_EPSILON_POINTS = 0.05;

/** How far (points) the line's own axis may wander in a drag across it. */
export const LINE_ALONG_TOLERANCE_POINTS = 0.25;

/**
 * A drag of the frame's edge across a straight horizontal or vertical line:
 * the line has no extent that way, and WPS reads the drag as a thicker (or
 * thinner) line whose other edge stays where it was. That is a new line
 * width plus a shift of the line by half the change, which one style write
 * makes (its `cm` moves the line; the geometry write cannot scale a zero
 * extent). Returns the width to add (points, negative to thin it) and the
 * shift in page space, or null for any other drag (one along the line, a
 * corner, a move).
 */
export function nativeVectorLineThicknessDrag(input: {
  readonly routeBounds: PdfNativeVectorNormalizedBounds;
  readonly targetBounds: PdfNativeVectorNormalizedBounds;
  readonly pageGeometry: PdfVectorPageGeometry;
  /** The stroke's width: a route's box may include it (a line's frame is
   * then as thick as the line looks, as in WPS). */
  readonly lineWidth?: number | null;
}): { readonly added: number; readonly translate: Readonly<{ x: number; y: number }> | null } | null {
  const crop = input.pageGeometry.cropBox;
  const turned = input.pageGeometry.rotation % 180 !== 0;
  const pageWidth = turned ? crop.yMax - crop.yMin : crop.xMax - crop.xMin;
  const pageHeight = turned ? crop.xMax - crop.xMin : crop.yMax - crop.yMin;
  const from = input.routeBounds;
  const to = input.targetBounds;
  const flat = LINE_AXIS_EPSILON_POINTS + Math.max(0, input.lineWidth ?? 0);
  const flatY = from.height * pageHeight <= flat;
  const flatX = from.width * pageWidth <= flat;
  if (flatX === flatY) return null;
  const [alongStart, alongExtent, alongTo, alongToExtent, alongPage, across, acrossTo, acrossPage] = flatY
    ? [from.x, from.width, to.x, to.width, pageWidth, from.height, to.height, pageHeight]
    : [from.y, from.height, to.y, to.height, pageHeight, from.width, to.width, pageWidth];
  if (
    Math.abs(alongTo - alongStart) * alongPage > LINE_ALONG_TOLERANCE_POINTS ||
    Math.abs(alongToExtent - alongExtent) * alongPage > LINE_ALONG_TOLERANCE_POINTS
  ) {
    return null;
  }
  const added = (acrossTo - across) * acrossPage;
  if (Math.abs(added) <= 0.1) return null;
  // The line moves with the middle of its frame: by half the change, toward
  // the dragged edge, so the other edge stays put. A display shift, turned
  // into page space as displayAffineToPdfUserSpace turns a translation.
  const tx = flatX ? to.x + to.width / 2 - (from.x + from.width / 2) : 0;
  const ty = flatY ? to.y + to.height / 2 - (from.y + from.height / 2) : 0;
  const cropWidth = crop.xMax - crop.xMin;
  const cropHeight = crop.yMax - crop.yMin;
  const rotation = input.pageGeometry.rotation;
  const [x, y] =
    rotation === 0
      ? [cropWidth * tx, -cropHeight * ty]
      : rotation === 90
        ? [cropWidth * ty, cropHeight * tx]
        : rotation === 180
          ? [-cropWidth * tx, cropHeight * ty]
          : [-cropWidth * ty, -cropHeight * tx];
  return Object.freeze({
    added,
    translate:
      Math.abs(x) < 1e-6 && Math.abs(y) < 1e-6
        ? null
        : Object.freeze({ x: x + 0, y: y + 0 }),
  });
}

/**
 * Any affine of the page as shown -- in display points, origin at the shown
 * page's top-left, y down (a rotated frame's resize: scale along the frame's
 * own axes about its held side) -- as the PDF user-space transform a write
 * applies: shown page -> transform -> back to the page's user space.
 */
export function pdfNativeVectorDisplayMatrixAffine(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  /** `[a, b, c, d, e, f]`: x' = a x + c y + e, y' = b x + d y + f. */
  readonly matrix: readonly [number, number, number, number, number, number];
}): PdfNativeVectorAffineTransform {
  const geometry = exactPageGeometry(input.pageGeometry);
  const [ma, mb, mc, md, me, mf] = input.matrix.map((value, index) =>
    exactFinite(value, `display matrix ${index}`),
  );
  if (Math.abs(ma * md - mb * mc) <= AFFINE_DETERMINANT_EPSILON) {
    fail("PDF vector display transform is singular.");
  }
  const { xMin, yMin, xMax, yMax } = geometry.cropBox;
  const toDisplay = (x: number, y: number): [number, number] =>
    geometry.rotation === 0
      ? [x - xMin, yMax - y]
      : geometry.rotation === 90
        ? [y - yMin, x - xMin]
        : geometry.rotation === 180
          ? [xMax - x, y - yMin]
          : [yMax - y, xMax - x];
  const fromDisplay = (x: number, y: number): [number, number] =>
    geometry.rotation === 0
      ? [x + xMin, yMax - y]
      : geometry.rotation === 90
        ? [y + xMin, x + yMin]
        : geometry.rotation === 180
          ? [xMax - x, y + yMin]
          : [xMax - y, yMax - x];
  const map = (x: number, y: number): [number, number] => {
    const [dx, dy] = toDisplay(x, y);
    return fromDisplay(ma * dx + mc * dy + me, mb * dx + md * dy + mf);
  };
  const origin = map(0, 0);
  const xAxis = map(1, 0);
  const yAxis = map(0, 1);
  return exactAffine({
    a: xAxis[0] - origin[0],
    b: xAxis[1] - origin[1],
    c: yAxis[0] - origin[0],
    d: yAxis[1] - origin[1],
    e: origin[0],
    f: origin[1],
  });
}

export function displayAffineToPdfUserSpace(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly scaleX: number;
  readonly scaleY: number;
  readonly translateX: number;
  readonly translateY: number;
}): PdfNativeVectorAffineTransform {
  const geometry = exactPageGeometry(input.pageGeometry);
  const sx = exactFinite(input.scaleX, "display scale x");
  const sy = exactFinite(input.scaleY, "display scale y");
  const tx = exactFinite(input.translateX, "display translate x");
  const ty = exactFinite(input.translateY, "display translate y");
  // Negative is a mirror (a handle pulled across the opposite edge).
  if (Math.abs(sx) <= AFFINE_DETERMINANT_EPSILON || Math.abs(sy) <= AFFINE_DETERMINANT_EPSILON) {
    fail("PDF vector display transform is singular.");
  }
  const crop = geometry.cropBox;
  const width = crop.xMax - crop.xMin;
  const height = crop.yMax - crop.yMin;
  if (geometry.rotation === 0) {
    return exactAffine({
      a: sx,
      b: 0,
      c: 0,
      d: sy,
      e: (1 - sx) * crop.xMin + width * tx,
      f: (1 - sy) * crop.yMax - height * ty,
    });
  }
  if (geometry.rotation === 90) {
    return exactAffine({
      a: sy,
      b: 0,
      c: 0,
      d: sx,
      e: (1 - sy) * crop.xMin + width * ty,
      f: (1 - sx) * crop.yMin + height * tx,
    });
  }
  if (geometry.rotation === 180) {
    return exactAffine({
      a: sx,
      b: 0,
      c: 0,
      d: sy,
      e: (1 - sx) * crop.xMax - width * tx,
      f: (1 - sy) * crop.yMin + height * ty,
    });
  }
  return exactAffine({
    a: sy,
    b: 0,
    c: 0,
    d: sx,
    e: (1 - sy) * crop.xMax - width * ty,
    f: (1 - sx) * crop.yMax - height * tx,
  });
}


/**
 * The page-space affine of a display-space one given in the page's own
 * units as shown (x to the right, y down, from the displayed top-left
 * corner; equal on both axes, so a turn can be expressed in it): what
 * carries a straight line's thickness along its own normal.
 */
export function displayUnitAffineToPdfUserSpace(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly matrix: Readonly<{ a: number; b: number; c: number; d: number; e: number; f: number }>;
}): PdfNativeVectorAffineTransform {
  const geometry = exactPageGeometry(input.pageGeometry);
  const { xMin, xMax, yMin, yMax } = geometry.cropBox;
  const m = input.matrix;
  for (const value of [m.a, m.b, m.c, m.d, m.e, m.f]) exactFinite(value, "display matrix");
  if (Math.abs(m.a * m.d - m.b * m.c) <= AFFINE_DETERMINANT_EPSILON) {
    fail("PDF vector display transform is singular.");
  }
  // The same page-to-display mapping displayAffineToPdfUserSpace inverts.
  const toDisplay = (x: number, y: number): [number, number] =>
    geometry.rotation === 0 ? [x - xMin, yMax - y]
      : geometry.rotation === 90 ? [y - yMin, x - xMin]
        : geometry.rotation === 180 ? [xMax - x, y - yMin]
          : [yMax - y, xMax - x];
  const toPage = (u: number, v: number): [number, number] =>
    geometry.rotation === 0 ? [xMin + u, yMax - v]
      : geometry.rotation === 90 ? [xMin + v, yMin + u]
        : geometry.rotation === 180 ? [xMax - u, yMin + v]
          : [xMax - v, yMax - u];
  const carry = (x: number, y: number) => {
    const [u, v] = toDisplay(x, y);
    return toPage(m.a * u + m.c * v + m.e, m.b * u + m.d * v + m.f);
  };
  const [ox, oy] = carry(0, 0);
  const [xx, xy] = carry(1, 0);
  const [yx, yy] = carry(0, 1);
  return exactAffine({ a: xx - ox, b: xy - oy, c: yx - ox, d: yy - oy, e: ox, f: oy });
}

/**
 * A display-space shift (points, to the right and down as the page is shown)
 * as the PDF user-space offset it is on a page turned by /Rotate: what a
 * duplicate pasted at the pointer is moved by.
 */
export function pdfNativeVectorDisplayShiftOffset(
  geometry: PdfVectorPageGeometry,
  shift: Readonly<{ x: number; y: number }>,
): Readonly<{ dx: number; dy: number }> {
  const x = exactFinite(shift.x, "duplicate shift x");
  const y = exactFinite(shift.y, "duplicate shift y");
  const rotation = exactPageGeometry(geometry).rotation;
  const [dx, dy] =
    rotation === 0 ? [x, -y] : rotation === 90 ? [y, x] : rotation === 180 ? [-x, y] : [-y, -x];
  return Object.freeze({ dx: dx + 0, dy: dy + 0 });
}

/**
 * The page-space affine that carries a frame `from` onto `to` (both page
 * fractions as displayed): a scale about the frame's own edges, mirrored on
 * an axis whose handle was pulled across the opposite edge. A resize written
 * as this transform scales the whole appearance -- stroke, dash, corners --
 * with the box, as WPS resizes a shape (its 线宽 still reads the width the
 * stream sets). Null when `from` has no extent to scale on an axis.
 */
export function pdfNativeVectorDisplayBoxAffine(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly from: PdfNativeVectorNormalizedBounds;
  readonly to: PdfNativeVectorNormalizedBounds;
  readonly mirror?: Readonly<{ x: boolean; y: boolean }> | null;
}): PdfNativeVectorAffineTransform | null {
  const { from, to } = input;
  if (!(from.width > GEOMETRY_EPSILON) || !(from.height > GEOMETRY_EPSILON)) return null;
  if (!(to.width > 0) || !(to.height > 0)) return null;
  const scaleX = ((input.mirror?.x ? -1 : 1) * to.width) / from.width;
  const scaleY = ((input.mirror?.y ? -1 : 1) * to.height) / from.height;
  // The frame's start edge lands on the target's start edge, or on its far
  // edge when that axis comes out mirrored.
  const translateX = (input.mirror?.x ? to.x + to.width : to.x) - scaleX * from.x;
  const translateY = (input.mirror?.y ? to.y + to.height : to.y) - scaleY * from.y;
  return displayAffineToPdfUserSpace({
    pageGeometry: input.pageGeometry,
    scaleX,
    scaleY,
    translateX,
    translateY,
  });
}

export function readPdfNativeVectorDeletePayload(
  value: unknown,
): PdfNativeVectorDeletePayload {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid native PDF vector delete payload.");
  }
  const payload = value as Record<string, unknown>;
  if (payload.schemaVersion !== MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION) {
    fail("Unsupported native PDF vector delete schema.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
    selection: exactSelection(payload.selection),
    pageGeometry: exactPageGeometry(payload.pageGeometry),
  });
}

export function exactRotationHistory(value: unknown): PdfNativeVectorRotationHistory {
  const history = plainRecord(value);
  const origin = plainRecord(history?.origin);
  const page = plainRecord(origin?.page);
  if (!history || !origin || !page ||
    (history.direction !== "undo" && history.direction !== "redo") ||
    origin.schemaVersion !== "mrai.pdf-native-vector-rotation-origin.v1") {
    fail("PDF vector rotation history binding is invalid.");
  }
  const sourceLength = exactInteger(origin.sourceLength, "rotation source length", 1);
  const candidateLength = exactInteger(origin.candidateLength, "rotation candidate length", 1);
  if (sourceLength >= candidateLength || candidateLength > PDF_EDITABLE_MAX_BYTES) {
    fail("PDF vector rotation history lengths are invalid.");
  }
  return Object.freeze({
    direction: history.direction,
    origin: Object.freeze({
      schemaVersion: "mrai.pdf-native-vector-rotation-origin.v1",
      sourceLength,
      candidateLength,
      page: Object.freeze({
        objectNumber: exactInteger(page.objectNumber, "rotation page object", 1),
        generation: exactInteger(page.generation, "rotation page generation"),
      }),
      pageGeometry: exactPageGeometry(origin.pageGeometry),
      affine: exactAffine(origin.affine),
      before: exactSelection(origin.before),
      after: exactSelection(origin.after),
    }),
  });
}

/** Capture only after the parent has verified the successful save receipt. */
export function createPdfNativeVectorRotationOrigin(
  plan: PdfNativeVectorSelectionEditPlan,
): PdfNativeVectorRotationOrigin {
  return exactRotationHistory({
    direction: "undo",
    origin: {
      schemaVersion: "mrai.pdf-native-vector-rotation-origin.v1",
      sourceLength: plan.prepared.sourceLength,
      candidateLength: plan.prepared.candidateBytes.length,
      page: plan.evidence.page,
      pageGeometry: plan.evidence.pageGeometry,
      affine: plan.payload.affine,
      before: plan.beforeSelection,
      after: plan.afterSelection,
    },
  }).origin;
}

export function readPdfNativeVectorSelectionEditPayload(
  value: unknown,
): PdfNativeVectorSelectionEditPayload {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid native PDF vector command payload.");
  }
  const payload = value as Record<string, unknown>;
  if (
    payload.schemaVersion !== MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION
  ) {
    fail("Unsupported native PDF vector command schema.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
    selection: exactSelection(payload.selection),
    pageGeometry: exactPageGeometry(payload.pageGeometry),
    affine: exactAffine(payload.affine, payload.outlineReplace !== undefined),
    ...(payload.rotationHistory === undefined ? {} : {
      rotationHistory: exactRotationHistory(payload.rotationHistory),
    }),
    ...(payload.vertexEdit === undefined ? {} : {
      vertexEdit: exactVertexEdit(payload.vertexEdit),
    }),
    ...(payload.outlineReplace === undefined ? {} : {
      outlineReplace: exactOutlineReplace(payload.outlineReplace),
    }),
  });
}

const MERGE_OPERATIONS = new Set(["union", "combine", "fragment", "intersect", "subtract"]);

/** Merge Shapes' outline: an operation and the path it makes (exactCommands). */
export function exactOutlineReplace(value: unknown): PdfNativeVectorOutlineReplace {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid PDF vector outline replacement.");
  }
  const record = value as { operation?: unknown; commands?: unknown };
  if (typeof record.operation !== "string" || !MERGE_OPERATIONS.has(record.operation)) {
    fail("Invalid PDF vector merge operation.");
  }
  return Object.freeze({
    operation: record.operation as PdfNativeVectorOutlineReplace["operation"],
    commands: exactCommands(record.commands),
  });
}

export function exactVertexEdit(value: unknown): PdfNativeVectorVertexEdit {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail("Invalid PDF vector vertex edit.");
  }
  const anchorIndex = exactInteger(
    (value as { anchorIndex?: unknown }).anchorIndex,
    "vertex anchor index",
  );
  if (anchorIndex > 4096) fail("Invalid PDF vector vertex anchor index.");
  return Object.freeze({ anchorIndex });
}

export function createPdfNativeVectorGeometryOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
      action: MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
      action: MRAI_NATIVE_VECTOR_DELETE_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
      action: MRAI_NATIVE_VECTOR_DUPLICATE_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
    {
      featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
      action: MRAI_NATIVE_VECTOR_REORDER_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function buildPdfNativeVectorGeometryCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativeVectorSelectionEditPayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativeVectorSelectionEditPayload(input.payload);
  if (
    payload.selection.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() ||
    !Number.isSafeInteger(input.sessionRevision) ||
    input.sessionRevision < 0 ||
    !Number.isSafeInteger(input.selectionRevision) ||
    input.selectionRevision < 1
  ) {
    fail("PDF vector command authority is stale.");
  }
  return Object.freeze({
    ...input.source,
    featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
    controlId: PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID,
    action: MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
}

export function buildPdfNativeVectorDeleteCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativeVectorDeletePayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativeVectorDeletePayload(input.payload);
  if (
    payload.selection.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() ||
    !Number.isSafeInteger(input.sessionRevision) ||
    input.sessionRevision < 0 ||
    !Number.isSafeInteger(input.selectionRevision) ||
    input.selectionRevision < 1
  ) {
    fail("PDF vector delete command authority is stale.");
  }
  return Object.freeze({
    ...input.source,
    featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
    controlId: PDF_NATIVE_VECTOR_DELETE_CONTROL_ID,
    action: MRAI_NATIVE_VECTOR_DELETE_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
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

/**
 * The native receipt journal is shared by text, images, outlines and vectors.
 * Keep the older reader untouched, but recognize the dedicated vector record
 * while validating the complete mixed source-SHA chain.
 */
export function readPdfNativeVectorAwareReplayAuthority(input: {
  readonly collection: unknown;
  readonly patchback: Readonly<{
    resourceId: string;
    resourceVersionId: string;
    outputSha256: string;
  }> | null;
  readonly renderedSourceSha256: string | null;
  readonly legacy: PdfNativeListTextReplayAuthority;
}): PdfNativeListTextReplayAuthority {
  if (
    input.legacy.state === "empty" ||
    input.legacy.state === "current" ||
    input.legacy.state === "exhausted"
  ) {
    return input.legacy;
  }
  if (!Array.isArray(input.collection)) return input.legacy;
  const records = input.collection.map(plainRecord);
  if (
    records.length < 1 ||
    records.length > MAX_ACKNOWLEDGED_NATIVE_TRANSACTIONS ||
    records.some((record) => !record) ||
    !records.some(
      (record) =>
        (record?.action === MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION &&
          record.kind === "vector-geometry-edit" &&
          record.schemaVersion ===
            MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION) ||
        (record?.action === MRAI_NATIVE_VECTOR_DUPLICATE_ACTION &&
          record.kind === "vector-duplicate" &&
          record.schemaVersion === MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION) ||
        (record?.action === MRAI_NATIVE_VECTOR_REORDER_ACTION &&
          record.kind === "vector-reorder" &&
          record.schemaVersion === MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION) ||
        (record?.action === MRAI_NATIVE_VECTOR_DELETE_ACTION &&
          record.kind === "vector-delete" &&
          record.schemaVersion === MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION) ||
        (record?.action === "pdf.vector.style.commit" &&
          record.kind === "vector-style-edit" &&
          record.schemaVersion === "mrai.pdf-native-vector-style-edit.v1"),
    )
  ) {
    return input.legacy;
  }
  const ids = new Set<string>();
  let resourceId: string | null = null;
  let priorOutputSha256: string | null = null;
  let priorVersionId: string | null = null;
  const validPairs = new Set([
    "native.list-text.replace\u0000list-text-replacement",
    "native.list-text.replace\u0000text-replacement",
    // A native text move rides the text edit command (text-move.ts).
    "native.list-text.replace\u0000text-move",
    // So does a group move of a whole box and its restore (text-group-move.ts).
    "native.list-text.replace\u0000text-group-move",
    // And editor history: the byte restore of an earlier text write (undo
    // and redo, byte-restore-history.ts). The legacy reader always accepted
    // it; leaving it out here refused every undo in a journal that also held
    // a vector or picture write, after the server had already persisted it.
    "native.list-text.replace\u0000text-window-history",
    "native.embedded-image.remove\u0000embedded-image-removal",
    "native.embedded-image.replace\u0000embedded-image-replacement",
    "native.embedded-image.transform\u0000embedded-image-transform",
    "native.embedded-image.reorder\u0000embedded-image-reorder",
    "native.text-object.delete\u0000text-object-delete",
    "native.text-decoration.commit\u0000text-decoration",
    "native.outline.mutate\u0000outline-mutation",
    "native.page.split\u0000page-split",
    "native.page.background\u0000page-background",
    `${MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION}\u0000vector-geometry-edit`,
    `${MRAI_NATIVE_VECTOR_DELETE_ACTION}\u0000vector-delete`,
    `${MRAI_NATIVE_VECTOR_DUPLICATE_ACTION}\u0000vector-duplicate`,
    `${MRAI_NATIVE_VECTOR_REORDER_ACTION}\u0000vector-reorder`,
    "pdf.vector.style.commit\u0000vector-style-edit",
    "pdf.text.font-extension.commit\u0000font-extension-text-replacement",
    "pdf.text.font-reuse.commit\u0000font-reuse-text-replacement",
    "pdf.text.font-extension.history\u0000font-extension-text-history",
    "native.paragraph-properties.replace\u0000paragraph-properties",
    "pdf.content.append.commit\u0000content-append",
    // A revision the engine appended, published as its bytes (stage 2).
    "native.bytes.publish\u0000bytes-publish",
    // The server's consolidation of the chain (source-consolidation.ts).
    `${MRAI_NATIVE_SOURCE_CONSOLIDATION_ACTION}\u0000${MRAI_NATIVE_SOURCE_CONSOLIDATION_KIND}`,
  ]);
  const textSchemas: Readonly<Record<string, string>> = {
    "pdf.text.font-extension.commit": "mrai.pdf-native-text.font-extension.v1",
    "pdf.text.font-reuse.commit": "mrai.pdf-native-text.font-reuse.v1",
    "pdf.text.font-extension.history": "mrai.pdf-native-text.font-extension-history.v1",
    "native.paragraph-properties.replace": "mrai.pdf-native-paragraph-properties.v1",
    "pdf.content.append.commit": "mrai.pdf-native-content-append.v1",
    "native.page.split": "mrai.pdf-native-page-split.v1",
    "native.page.background": "mrai.pdf-native-page-background.v1",
    "native.bytes.publish": "mrai.pdf-native-bytes-publish.v1",
  };
  for (const record of records as Readonly<Record<string, unknown>>[]) {
    const source = plainRecord(record.source);
    const derived = plainRecord(record.derivedCandidate);
    const id = typeof record.id === "string" ? record.id : "";
    const nextResourceId =
      typeof source?.resourceId === "string" ? source.resourceId : "";
    const nextVersionId =
      typeof source?.resourceVersionId === "string"
        ? source.resourceVersionId
        : "";
    const sourceSha256 =
      typeof source?.sha256 === "string" ? source.sha256.toLowerCase() : "";
    const outputSha256 =
      typeof derived?.sha256 === "string" ? derived.sha256.toLowerCase() : "";
    if (
      !id ||
      id !== record.transactionId ||
      ids.has(id) ||
      typeof record.action !== "string" ||
      typeof record.kind !== "string" ||
      !validPairs.has(`${record.action}\u0000${record.kind}`) ||
      (textSchemas[record.action] !== undefined &&
        record.schemaVersion !== textSchemas[record.action]) ||
      (record.kind === "text-move" && record.schemaVersion !== "mrai.pdf-native-text-move.v1") ||
      (record.kind === "text-group-move" && record.schemaVersion !== "mrai.pdf-native-text-group-move.v1") ||
      (record.kind === "text-window-history" &&
        record.schemaVersion !== "mrai.pdf-native-text.window-history.v1") ||
      (record.action === MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION &&
        (record.schemaVersion !==
          MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION ||
          record.featureId !== PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID ||
          record.controlId !== PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID)) ||
      (record.action === MRAI_NATIVE_VECTOR_DELETE_ACTION &&
        (record.schemaVersion !== MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION ||
          record.featureId !== PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID ||
          record.controlId !== PDF_NATIVE_VECTOR_DELETE_CONTROL_ID)) ||
      (record.action === "pdf.vector.style.commit" &&
        (record.schemaVersion !== "mrai.pdf-native-vector-style-edit.v1" ||
          record.featureId !== PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID ||
          record.controlId !== "PDFCTL.PAGE.VECTOR.STYLE")) ||
      !nextResourceId ||
      !nextVersionId ||
      !SHA256_PATTERN.test(sourceSha256) ||
      !SHA256_PATTERN.test(outputSha256) ||
      sourceSha256 === outputSha256 ||
      (resourceId !== null && resourceId !== nextResourceId) ||
      (priorOutputSha256 !== null && sourceSha256 !== priorOutputSha256) ||
      (priorVersionId !== null && priorVersionId === nextVersionId)
    ) {
      return input.legacy;
    }
    ids.add(id);
    resourceId = nextResourceId;
    priorVersionId = nextVersionId;
    priorOutputSha256 = outputSha256;
  }
  const patchback = input.patchback;
  if (
    !patchback ||
    resourceId === null ||
    priorVersionId === null ||
    priorOutputSha256 === null ||
    patchback.resourceId !== resourceId ||
    patchback.resourceVersionId === priorVersionId ||
    patchback.outputSha256.toLowerCase() !== priorOutputSha256 ||
    input.renderedSourceSha256?.toLowerCase() !== priorOutputSha256
  ) {
    return Object.freeze({
      state: "pending" as const,
      transactionCount: records.length,
      patchback,
    });
  }
  return records.length === MAX_ACKNOWLEDGED_NATIVE_TRANSACTIONS
    ? Object.freeze({
        state: "exhausted" as const,
        transactionCount: MAX_ACKNOWLEDGED_NATIVE_TRANSACTIONS,
        patchback,
        latestDerivedSha256: priorOutputSha256,
      })
    : Object.freeze({
        state: "current" as const,
        transactionCount: records.length,
        patchback,
        latestDerivedSha256: priorOutputSha256,
      });
}

export function exactByteList(value: unknown, label: string): readonly number[] {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    value.length > PDF_EDITABLE_MAX_BYTES ||
    value.some(
      (entry) =>
        typeof entry !== "number" ||
        !Number.isInteger(entry) ||
        entry < 0 ||
        entry > 255,
    )
  ) {
    fail(`PDF vector ${label} bytes are invalid.`);
  }
  return Object.freeze([...(value as number[])]);
}

export function exactOffset(value: unknown): Readonly<{ dx: number; dy: number }> {
  const record = plainRecord(value);
  if (!record) fail("PDF vector duplicate offset is invalid.");
  return Object.freeze({
    dx: exactFinite(record.dx, "duplicate offset x"),
    dy: exactFinite(record.dy, "duplicate offset y"),
  });
}

export function readPdfNativeVectorDuplicatePayload(
  value: unknown,
): PdfNativeVectorDuplicatePayload {
  const payload = plainRecord(value);
  if (!payload) fail("Invalid native PDF vector duplicate payload.");
  if (payload.schemaVersion !== MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION) {
    fail("Unsupported native PDF vector duplicate schema.");
  }
  if (payload.direction !== "insert" && payload.direction !== "remove") {
    fail("Invalid native PDF vector duplicate direction.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
    selection: exactSelection(payload.selection),
    pageGeometry: exactPageGeometry(payload.pageGeometry),
    direction: payload.direction,
    offset: exactOffset(payload.offset),
    ...(payload.direction === "remove"
      ? { insertedBytes: exactByteList(payload.insertedBytes, "duplicate") }
      : {}),
  });
}

export function readPdfNativeVectorReorderPayload(
  value: unknown,
): PdfNativeVectorReorderPayload {
  const payload = plainRecord(value);
  if (!payload) fail("Invalid native PDF vector reorder payload.");
  if (payload.schemaVersion !== MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION) {
    fail("Unsupported native PDF vector reorder schema.");
  }
  if (
    payload.direction !== "forward" &&
    payload.direction !== "backward" &&
    payload.direction !== "front" &&
    payload.direction !== "back" &&
    payload.direction !== "restore"
  ) {
    fail("Invalid native PDF vector reorder direction.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
    selection: exactSelection(payload.selection),
    pageGeometry: exactPageGeometry(payload.pageGeometry),
    direction: payload.direction,
    ...(payload.direction === "restore"
      ? {
          ...(payload.neighbour === undefined || payload.neighbour === null
            ? {}
            : { neighbour: exactSelection(payload.neighbour) }),
          reorderedBytes: exactByteList(payload.reorderedBytes, "reorder"),
          originalBytes: exactByteList(payload.originalBytes, "reorder"),
        }
      : {}),
  });
}

export function vectorCommandAuthority(
  selection: PdfNativeVectorWriteSelection,
  input: {
    readonly source: PdfEngineSourceIdentity;
    readonly transactionId: string;
    readonly sessionRevision: number;
    readonly selectionRevision: number;
  },
  label: string,
): void {
  if (
    selection.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() ||
    !Number.isSafeInteger(input.sessionRevision) ||
    input.sessionRevision < 0 ||
    !Number.isSafeInteger(input.selectionRevision) ||
    input.selectionRevision < 1
  ) {
    fail(`PDF vector ${label} command authority is stale.`);
  }
}

export function buildPdfNativeVectorDuplicateCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativeVectorDuplicatePayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativeVectorDuplicatePayload(input.payload);
  vectorCommandAuthority(payload.selection, input, "duplicate");
  return Object.freeze({
    ...input.source,
    featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
    controlId: PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID,
    action: MRAI_NATIVE_VECTOR_DUPLICATE_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
}

export function buildPdfNativeVectorReorderCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativeVectorReorderPayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativeVectorReorderPayload(input.payload);
  vectorCommandAuthority(payload.selection, input, "reorder");
  return Object.freeze({
    ...input.source,
    featureId: PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID,
    controlId: PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID,
    action: MRAI_NATIVE_VECTOR_REORDER_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
}
