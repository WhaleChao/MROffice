/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION } from "./semantic-appearance-owner-schema";
export { PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION };
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRawStream,
  PDFRef,
  type PDFObject,
  type PDFPage,
} from "pdf-lib";

import { ACROFORM_MAX_FIELDS } from "@mrnote/core/web/lib/pdf/acroform.ts";
import { loadPdfForInspection } from "./encrypted-guard";
import {
  countMraiNativeIndirectObjectReferences,
  inspectMraiNativeObject,
  MAX_MRAI_NATIVE_SOURCE_BYTES,
  sha256PdfBytes,
  type PdfNativeObjectInspection,
} from "./engine/native/content-stream-transaction";


/** A page's annotations read at most (a form's fields, a page of links). */
const MAX_PAGE_ANNOTATIONS = Math.min(ACROFORM_MAX_FIELDS, 4_096);




const MAX_ANNOTATIONS = 16_384;
const MAX_PARENT_DEPTH = 32;
const MAX_APPEARANCE_STATES = 64;
const MAX_GEOMETRY_POINTS = 256;
const NAME_SUBTYPE = PDFName.of("Subtype");
const NAME_ANNOTS = PDFName.of("Annots");
const NAME_RECT = PDFName.of("Rect");
const NAME_QUAD_POINTS = PDFName.of("QuadPoints");
const NAME_PARENT = PDFName.of("Parent");
const NAME_FT = PDFName.of("FT");
const NAME_FF = PDFName.of("Ff");
const NAME_AP = PDFName.of("AP");
const NAME_N = PDFName.of("N");
const NAME_AS = PDFName.of("AS");
const NAME_BBOX = PDFName.of("BBox");
const NAME_MATRIX = PDFName.of("Matrix");
const NAME_RESOURCES = PDFName.of("Resources");
const FIELD_FLAG_RADIO = 1 << 15;
const FIELD_FLAG_PUSH_BUTTON = 1 << 16;
const FIELD_FLAG_COMBO = 1 << 17;

export type PdfSemanticAppearanceInventoryErrorCode =
  "source-too-large" | "encrypted" | "inspection-failure" | "capacity-exceeded";

export class PdfSemanticAppearanceInventoryError extends Error {
  constructor(readonly code: PdfSemanticAppearanceInventoryErrorCode) {
    super(`PDF semantic appearance inventory failed: ${code}.`);
    this.name = "PdfSemanticAppearanceInventoryError";
  }
}

export interface PdfSemanticAppearanceObjectRef {
  readonly objectNumber: number;
  readonly generation: number;
  readonly objectSha256: string;
}

export interface PdfSemanticAppearanceNormalizedRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface PdfSemanticAppearanceNormalizedPoint {
  readonly x: number;
  readonly y: number;
}

export interface PdfSemanticAppearanceGeometry {
  readonly rect: PdfSemanticAppearanceNormalizedRect;
  readonly quadPoints?: readonly PdfSemanticAppearanceNormalizedPoint[];
}

export type PdfSemanticAppearanceResourceIdentity = Readonly<
  | { kind: "missing" }
  | {
      kind: "indirect";
      objectNumber: number;
      generation: number;
      objectSha256: string;
    }
  | { kind: "embedded"; dictionarySha256: string }
>;

export interface PdfSemanticAppearanceFormEvidence {
  readonly objectNumber: number;
  readonly generation: number;
  readonly objectSha256: string;
  readonly streamSha256: string;
  readonly referenceCount: number;
  readonly bbox: readonly [number, number, number, number];
  readonly matrix: readonly [number, number, number, number, number, number];
  readonly resources: PdfSemanticAppearanceResourceIdentity;
}

export type PdfSemanticWidgetFieldKind =
  "text" | "checkbox" | "radio" | "dropdown" | "listbox";

export type PdfSemanticWidgetAppearanceEvidence = Readonly<
  | {
      kind: "direct-form";
      selectedAppearanceState: string | null;
      form: PdfSemanticAppearanceFormEvidence;
    }
  | {
      kind: "state-dictionary";
      selectedAppearanceState: string;
      stateDictionaryRef?: PdfSemanticAppearanceObjectRef;
      entries: readonly Readonly<{
        state: string;
        form: PdfSemanticAppearanceFormEvidence;
      }>[];
      selectedForm: PdfSemanticAppearanceFormEvidence;
    }
>;

export interface PdfSemanticAppearanceOwnerBase {
  readonly schemaVersion: typeof PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION;
  readonly id: string;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly pageRef: Readonly<{ objectNumber: number; generation: number }>;
  readonly annotationIndex: number;
  readonly annotation: PdfSemanticAppearanceObjectRef;
  readonly subtype:
    "Widget" | "Link" | "Text" | "Highlight" | "Underline" | "Square";
  readonly geometry: PdfSemanticAppearanceGeometry;
}

export interface PdfSemanticWidgetAppearanceOwner extends PdfSemanticAppearanceOwnerBase {
  readonly ownerKind: "widget";
  readonly subtype: "Widget";
  readonly widget: Readonly<{
    readonly fieldType: "Tx" | "Btn" | "Ch";
    readonly fieldKind: PdfSemanticWidgetFieldKind;
    readonly terminalField: PdfSemanticAppearanceObjectRef;
    readonly parentChain: readonly PdfSemanticAppearanceObjectRef[];
    readonly fieldTypeSource: PdfSemanticAppearanceObjectRef;
    readonly fieldFlagsSource: PdfSemanticAppearanceObjectRef | null;
    readonly appearance: PdfSemanticWidgetAppearanceEvidence;
  }>;
}

export interface PdfSemanticSynthesizedAnnotationOwner extends PdfSemanticAppearanceOwnerBase {
  readonly ownerKind: "annotation";
  readonly subtype: "Link" | "Text" | "Highlight" | "Underline" | "Square";
  readonly annotationAppearance: Readonly<{ kind: "synthesized-no-ap" }>;
}

export type PdfSemanticAppearanceOwner =
  PdfSemanticWidgetAppearanceOwner | PdfSemanticSynthesizedAnnotationOwner;

export type PdfSemanticAppearancePreserveReason =
  | "direct-annotation"
  | "duplicate-annotation"
  | "unsupported-subtype"
  | "unsupported-field"
  | "missing-appearance"
  | "malformed-appearance"
  | "unlocatable-appearance"
  | "missing-appearance-state"
  | "unknown-appearance-state"
  | "shared-appearance-stream"
  | "stored-annotation-appearance"
  | "malformed-geometry"
  | "malformed-parent-chain"
  | "cyclic-parent-chain";

export interface PdfSemanticAppearancePreserveOnly {
  readonly id: string;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly annotationIndex: number;
  readonly subtype: string | null;
  readonly locator: Readonly<
    | PdfSemanticAppearanceObjectRef
    | {
        kind: "page-annotation-array";
        pageNumber: number;
        annotationIndex: number;
        dictionarySha256: string;
      }
  >;
  readonly geometry: PdfSemanticAppearanceGeometry | null;
  readonly reason: PdfSemanticAppearancePreserveReason;
  /** Sanitized semantic field/AP evidence retained when only stream sharing blocks ownership. */
  readonly widget?: PdfSemanticWidgetAppearanceOwner["widget"];
}

export interface PdfSemanticAppearanceOwnerInventory {
  readonly schemaVersion: typeof PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageCount: number;
  readonly owners: readonly PdfSemanticAppearanceOwner[];
  readonly preserveOnly: readonly PdfSemanticAppearancePreserveOnly[];
  readonly diagnostics: Readonly<{
    annotationCount: number;
    exactOwnerCount: number;
    widgetOwnerCount: number;
    synthesizedOwnerCount: number;
    directFormCount: number;
    stateDictionaryCount: number;
    preserveOnlyCount: number;
    sharedAppearanceCount: number;
    duplicateAnnotationCount: number;
    unsupportedCount: number;
    malformedCount: number;
  }>;
}

interface AnnotationScan {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly pageRef: Readonly<{ objectNumber: number; generation: number }>;
  readonly annotationIndex: number;
  readonly annotationRef: PDFRef | null;
  readonly annotationDict: PDFDict | null;
  readonly annotationObject: PdfSemanticAppearanceObjectRef | null;
  readonly fallbackDictionarySha256: string;
  readonly subtype: string | null;
  readonly geometry: PdfSemanticAppearanceGeometry | null;
}

interface OwnerDraft {
  readonly scan: AnnotationScan;
  readonly owner:
    | Omit<PdfSemanticWidgetAppearanceOwner, "id">
    | Omit<PdfSemanticSynthesizedAnnotationOwner, "id">;
  readonly appearanceRefs: readonly string[];
}

class PreserveProblem extends Error {
  constructor(readonly reason: PdfSemanticAppearancePreserveReason) {
    super(reason);
  }
}

function refKey(ref: PDFRef): string {
  return `${ref.objectNumber}:${ref.generationNumber}`;
}

function refCoordinates(ref: PDFRef) {
  return Object.freeze({
    objectNumber: ref.objectNumber,
    generation: ref.generationNumber,
  });
}

function rounded(value: number): number {
  const result = Number(value.toFixed(12));
  return Object.is(result, -0) ? 0 : result;
}

function exactName(value: PDFObject | undefined): string | null {
  if (!(value instanceof PDFName)) return null;
  const name = value.asString();
  return name.startsWith("/") ? name.slice(1) : name;
}

function resolvedObject(document: PDFDocument, value: PDFObject | undefined) {
  return value instanceof PDFRef ? document.context.lookup(value) : value;
}

function resolvedDict(
  document: PDFDocument,
  value: PDFObject | undefined,
): PDFDict | null {
  const resolved = resolvedObject(document, value);
  return resolved instanceof PDFDict ? resolved : null;
}

function resolvedArray(
  document: PDFDocument,
  value: PDFObject | undefined,
): PDFArray | null {
  const resolved = resolvedObject(document, value);
  return resolved instanceof PDFArray ? resolved : null;
}

function exactNumber(value: PDFObject | undefined): number | null {
  return value instanceof PDFNumber && Number.isFinite(value.asNumber())
    ? value.asNumber()
    : null;
}

function exactNumberArray(
  document: PDFDocument,
  value: PDFObject | undefined,
  exactLength?: number,
): readonly number[] | null {
  const array = resolvedArray(document, value);
  if (!array || (exactLength !== undefined && array.size() !== exactLength)) {
    return null;
  }
  const result: number[] = [];
  for (let index = 0; index < array.size(); index += 1) {
    const number = exactNumber(resolvedObject(document, array.get(index)));
    if (number === null) return null;
    result.push(number);
  }
  return Object.freeze(result);
}

function projectedPoint(
  x: number,
  y: number,
  page: PDFPage,
): PdfSemanticAppearanceNormalizedPoint {
  const crop = page.getCropBox();
  const rotation = ((page.getRotation().angle % 360) + 360) % 360;
  const cropWidth = crop.width;
  const cropHeight = crop.height;
  let projectedX: number;
  let projectedY: number;
  let width: number;
  let height: number;
  if (rotation === 0) {
    projectedX = x - crop.x;
    projectedY = crop.y + crop.height - y;
    width = cropWidth;
    height = cropHeight;
  } else if (rotation === 90) {
    projectedX = y - crop.y;
    projectedY = x - crop.x;
    width = cropHeight;
    height = cropWidth;
  } else if (rotation === 180) {
    projectedX = crop.x + crop.width - x;
    projectedY = y - crop.y;
    width = cropWidth;
    height = cropHeight;
  } else if (rotation === 270) {
    projectedX = crop.y + crop.height - y;
    projectedY = crop.x + crop.width - x;
    width = cropHeight;
    height = cropWidth;
  } else {
    throw new PreserveProblem("malformed-geometry");
  }
  const normalized = Object.freeze({
    x: rounded(projectedX / width),
    y: rounded(projectedY / height),
  });
  if (!Number.isFinite(normalized.x) || !Number.isFinite(normalized.y)) {
    throw new PreserveProblem("malformed-geometry");
  }
  return normalized;
}

function normalizedGeometry(
  document: PDFDocument,
  page: PDFPage,
  annotation: PDFDict,
  subtype: string | null,
): PdfSemanticAppearanceGeometry {
  const rect = exactNumberArray(document, annotation.get(NAME_RECT), 4);
  if (!rect) throw new PreserveProblem("malformed-geometry");
  const xMin = Math.min(rect[0], rect[2]);
  const yMin = Math.min(rect[1], rect[3]);
  const xMax = Math.max(rect[0], rect[2]);
  const yMax = Math.max(rect[1], rect[3]);
  if (xMax <= xMin || yMax <= yMin) {
    throw new PreserveProblem("malformed-geometry");
  }
  const corners = [
    projectedPoint(xMin, yMin, page),
    projectedPoint(xMax, yMin, page),
    projectedPoint(xMax, yMax, page),
    projectedPoint(xMin, yMax, page),
  ];
  const left = Math.min(...corners.map((point) => point.x));
  const top = Math.min(...corners.map((point) => point.y));
  const right = Math.max(...corners.map((point) => point.x));
  const bottom = Math.max(...corners.map((point) => point.y));
  const geometry: {
    rect: PdfSemanticAppearanceNormalizedRect;
    quadPoints?: readonly PdfSemanticAppearanceNormalizedPoint[];
  } = {
    rect: Object.freeze({
      x: rounded(left),
      y: rounded(top),
      width: rounded(right - left),
      height: rounded(bottom - top),
    }),
  };
  const quadEntry = annotation.get(NAME_QUAD_POINTS);
  const needsQuad = subtype === "Highlight" || subtype === "Underline";
  if (quadEntry || needsQuad) {
    const values = exactNumberArray(document, quadEntry);
    if (
      !values ||
      values.length < 8 ||
      values.length % 8 !== 0 ||
      values.length / 2 > MAX_GEOMETRY_POINTS
    ) {
      throw new PreserveProblem("malformed-geometry");
    }
    const points: PdfSemanticAppearanceNormalizedPoint[] = [];
    for (let index = 0; index < values.length; index += 2) {
      points.push(projectedPoint(values[index], values[index + 1], page));
    }
    geometry.quadPoints = Object.freeze(points);
  }
  return Object.freeze(geometry);
}

async function sha256Text(value: string): Promise<string> {
  return sha256PdfBytes(new TextEncoder().encode(value));
}

function objectCacheFor(sourceBytes: Uint8Array) {
  const cache = new Map<string, Promise<PdfNativeObjectInspection>>();
  return (ref: PDFRef) => {
    const key = refKey(ref);
    let inspection = cache.get(key);
    if (!inspection) {
      inspection = inspectMraiNativeObject(sourceBytes, {
        objectNumber: ref.objectNumber,
        generation: ref.generationNumber,
      });
      cache.set(key, inspection);
    }
    return inspection;
  };
}

async function refEvidence(
  ref: PDFRef,
  inspectObject: (ref: PDFRef) => Promise<PdfNativeObjectInspection>,
): Promise<PdfSemanticAppearanceObjectRef> {
  const inspection = await inspectObject(ref);
  return Object.freeze({
    objectNumber: ref.objectNumber,
    generation: ref.generationNumber,
    objectSha256: inspection.objectSha256,
  });
}

async function resourcesIdentity(
  document: PDFDocument,
  stream: PDFRawStream,
  inspectObject: (ref: PDFRef) => Promise<PdfNativeObjectInspection>,
): Promise<PdfSemanticAppearanceResourceIdentity> {
  const entry = stream.dict.get(NAME_RESOURCES);
  if (!entry) return Object.freeze({ kind: "missing" });
  if (entry instanceof PDFRef) {
    const evidence = await refEvidence(entry, inspectObject);
    return Object.freeze({ kind: "indirect", ...evidence });
  }
  const dictionary = resolvedDict(document, entry);
  if (!dictionary) throw new PreserveProblem("malformed-appearance");
  return Object.freeze({
    kind: "embedded",
    dictionarySha256: await sha256Text(dictionary.toString()),
  });
}

async function appearanceFormEvidence(
  sourceBytes: Uint8Array,
  document: PDFDocument,
  ref: PDFRef,
  inspectObject: (ref: PDFRef) => Promise<PdfNativeObjectInspection>,
): Promise<PdfSemanticAppearanceFormEvidence> {
  const object = document.context.lookup(ref);
  if (!(object instanceof PDFRawStream)) {
    throw new PreserveProblem("malformed-appearance");
  }
  if (exactName(object.dict.get(NAME_SUBTYPE)) !== "Form") {
    throw new PreserveProblem("malformed-appearance");
  }
  const bbox = exactNumberArray(document, object.dict.get(NAME_BBOX), 4);
  const matrix = object.dict.get(NAME_MATRIX)
    ? exactNumberArray(document, object.dict.get(NAME_MATRIX), 6)
    : Object.freeze([1, 0, 0, 1, 0, 0]);
  if (!bbox || !matrix || bbox[2] <= bbox[0] || bbox[3] <= bbox[1]) {
    throw new PreserveProblem("malformed-appearance");
  }
  const inspection = await inspectObject(ref);
  if (!inspection.streamSha256) {
    throw new PreserveProblem("malformed-appearance");
  }
  return Object.freeze({
    objectNumber: ref.objectNumber,
    generation: ref.generationNumber,
    objectSha256: inspection.objectSha256,
    streamSha256: inspection.streamSha256,
    referenceCount: countMraiNativeIndirectObjectReferences(sourceBytes, {
      objectNumber: ref.objectNumber,
      generation: ref.generationNumber,
    }),
    bbox: Object.freeze([bbox[0], bbox[1], bbox[2], bbox[3]]) as readonly [
      number,
      number,
      number,
      number,
    ],
    matrix: Object.freeze([
      matrix[0],
      matrix[1],
      matrix[2],
      matrix[3],
      matrix[4],
      matrix[5],
    ]) as readonly [number, number, number, number, number, number],
    resources: await resourcesIdentity(document, object, inspectObject),
  });
}

async function fieldAuthority(
  document: PDFDocument,
  widgetRef: PDFRef,
  widget: PDFDict,
  inspectObject: (ref: PDFRef) => Promise<PdfNativeObjectInspection>,
) {
  const parentRefs: PDFRef[] = [];
  const seen = new Set<string>([refKey(widgetRef)]);
  let current = widget;
  let currentRef = widgetRef;
  let type: string | null = null;
  let typeSource: PDFRef | null = null;
  let flags = 0;
  let flagsSource: PDFRef | null = null;
  let terminalRef = widgetRef;
  for (let depth = 0; depth <= MAX_PARENT_DEPTH; depth += 1) {
    if (type === null) {
      const candidate = exactName(current.get(NAME_FT));
      if (candidate) {
        type = candidate;
        typeSource = currentRef;
      }
    }
    if (flagsSource === null) {
      const candidate = exactNumber(
        resolvedObject(document, current.get(NAME_FF)),
      );
      if (
        candidate !== null &&
        Number.isSafeInteger(candidate) &&
        candidate >= 0
      ) {
        flags = candidate;
        flagsSource = currentRef;
      }
    }
    const parentEntry = current.get(NAME_PARENT);
    if (!parentEntry) break;
    if (!(parentEntry instanceof PDFRef)) {
      throw new PreserveProblem("malformed-parent-chain");
    }
    const key = refKey(parentEntry);
    if (seen.has(key)) throw new PreserveProblem("cyclic-parent-chain");
    seen.add(key);
    const parent = document.context.lookup(parentEntry);
    if (!(parent instanceof PDFDict)) {
      throw new PreserveProblem("malformed-parent-chain");
    }
    parentRefs.push(parentEntry);
    if (parentRefs.length === 1) terminalRef = parentEntry;
    current = parent;
    currentRef = parentEntry;
  }
  if (parentRefs.length > MAX_PARENT_DEPTH) {
    throw new PreserveProblem("malformed-parent-chain");
  }
  if (!type || !typeSource) throw new PreserveProblem("unsupported-field");
  let fieldKind: PdfSemanticWidgetFieldKind;
  if (type === "Tx") fieldKind = "text";
  else if (type === "Btn") {
    if ((flags & FIELD_FLAG_PUSH_BUTTON) !== 0) {
      throw new PreserveProblem("unsupported-field");
    }
    fieldKind = (flags & FIELD_FLAG_RADIO) !== 0 ? "radio" : "checkbox";
  } else if (type === "Ch") {
    fieldKind = (flags & FIELD_FLAG_COMBO) !== 0 ? "dropdown" : "listbox";
  } else {
    throw new PreserveProblem("unsupported-field");
  }
  return Object.freeze({
    fieldType: type as "Tx" | "Btn" | "Ch",
    fieldKind,
    terminalField: await refEvidence(terminalRef, inspectObject),
    parentChain: Object.freeze(
      await Promise.all(
        parentRefs.map((ref) => refEvidence(ref, inspectObject)),
      ),
    ),
    fieldTypeSource: await refEvidence(typeSource, inspectObject),
    fieldFlagsSource: flagsSource
      ? await refEvidence(flagsSource, inspectObject)
      : null,
  });
}

async function widgetAppearance(
  sourceBytes: Uint8Array,
  document: PDFDocument,
  widget: PDFDict,
  inspectObject: (ref: PDFRef) => Promise<PdfNativeObjectInspection>,
): Promise<
  Readonly<{
    evidence: PdfSemanticWidgetAppearanceEvidence;
    refs: readonly string[];
  }>
> {
  const ap = resolvedDict(document, widget.get(NAME_AP));
  if (!ap) {
    throw new PreserveProblem(
      widget.get(NAME_AP) ? "malformed-appearance" : "missing-appearance",
    );
  }
  const normalEntry = ap.get(NAME_N);
  if (!normalEntry) throw new PreserveProblem("missing-appearance");
  if (normalEntry instanceof PDFRef) {
    const resolved = document.context.lookup(normalEntry);
    if (resolved instanceof PDFRawStream) {
      return Object.freeze({
        evidence: Object.freeze({
          kind: "direct-form",
          selectedAppearanceState: exactName(widget.get(NAME_AS)),
          form: await appearanceFormEvidence(
            sourceBytes,
            document,
            normalEntry,
            inspectObject,
          ),
        }),
        refs: Object.freeze([refKey(normalEntry)]),
      });
    }
  }
  const stateRef = normalEntry instanceof PDFRef ? normalEntry : null;
  const states = resolvedDict(document, normalEntry);
  if (!states) {
    throw new PreserveProblem(
      normalEntry instanceof PDFRawStream
        ? "unlocatable-appearance"
        : "malformed-appearance",
    );
  }
  const entries = states.entries();
  if (entries.length === 0 || entries.length > MAX_APPEARANCE_STATES) {
    throw new PreserveProblem("malformed-appearance");
  }
  const selected = exactName(widget.get(NAME_AS));
  if (!selected) throw new PreserveProblem("missing-appearance-state");
  const evidenceEntries: Array<{
    state: string;
    form: PdfSemanticAppearanceFormEvidence;
  }> = [];
  for (const [stateName, value] of entries) {
    if (!(value instanceof PDFRef)) {
      throw new PreserveProblem("unlocatable-appearance");
    }
    evidenceEntries.push({
      state: exactName(stateName)!,
      form: await appearanceFormEvidence(
        sourceBytes,
        document,
        value,
        inspectObject,
      ),
    });
  }
  evidenceEntries.sort((left, right) => left.state.localeCompare(right.state));
  const selectedEntry = evidenceEntries.find(
    (entry) => entry.state === selected,
  );
  if (!selectedEntry) throw new PreserveProblem("unknown-appearance-state");
  return Object.freeze({
    evidence: Object.freeze({
      kind: "state-dictionary",
      selectedAppearanceState: selected,
      ...(stateRef
        ? { stateDictionaryRef: await refEvidence(stateRef, inspectObject) }
        : {}),
      entries: Object.freeze(
        evidenceEntries.map((entry) => Object.freeze(entry)),
      ),
      selectedForm: selectedEntry.form,
    }),
    refs: Object.freeze(
      evidenceEntries.map(
        (entry) => `${entry.form.objectNumber}:${entry.form.generation}`,
      ),
    ),
  });
}

function supportedAnnotationSubtype(
  subtype: string | null,
): subtype is "Link" | "Text" | "Highlight" | "Underline" | "Square" {
  return ["Link", "Text", "Highlight", "Underline", "Square"].includes(
    subtype ?? "",
  );
}

async function ownerId(value: unknown): Promise<string> {
  return `pdf-semantic-appearance:${await sha256Text(JSON.stringify(value))}`;
}

async function preserveRecord(
  scan: AnnotationScan,
  reason: PdfSemanticAppearancePreserveReason,
): Promise<PdfSemanticAppearancePreserveOnly> {
  const locator = scan.annotationObject
    ? scan.annotationObject
    : Object.freeze({
        kind: "page-annotation-array" as const,
        pageNumber: scan.pageNumber,
        annotationIndex: scan.annotationIndex,
        dictionarySha256: scan.fallbackDictionarySha256,
      });
  const identity = {
    sourceSha256: scan.sourceSha256,
    pageNumber: scan.pageNumber,
    annotationIndex: scan.annotationIndex,
    locator,
    reason,
  };
  return Object.freeze({
    id: `pdf-semantic-appearance-preserve:${await sha256Text(
      JSON.stringify(identity),
    )}`,
    sourceSha256: scan.sourceSha256,
    pageNumber: scan.pageNumber,
    annotationIndex: scan.annotationIndex,
    subtype: scan.subtype,
    locator,
    geometry: scan.geometry,
    reason,
  });
}

function preserveReasonCategory(reason: PdfSemanticAppearancePreserveReason) {
  if (reason === "duplicate-annotation") return "duplicate" as const;
  if (reason === "shared-appearance-stream") return "shared" as const;
  if (reason === "unsupported-subtype" || reason === "unsupported-field") {
    return "unsupported" as const;
  }
  return "malformed" as const;
}

export async function inspectPdfSemanticAppearanceOwners(
  sourceBytes: Uint8Array,
): Promise<PdfSemanticAppearanceOwnerInventory> {
  if (
    sourceBytes.length === 0 ||
    sourceBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES
  ) {
    throw new PdfSemanticAppearanceInventoryError("source-too-large");
  }
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  let document: PDFDocument | null;
  try {
    document = await loadPdfForInspection(sourceBytes, {
      updateMetadata: false,
    });
  } catch {
    throw new PdfSemanticAppearanceInventoryError("inspection-failure");
  }
  if (!document) throw new PdfSemanticAppearanceInventoryError("encrypted");
  const inspectObject = objectCacheFor(sourceBytes);
  const scans: AnnotationScan[] = [];
  let annotationCount = 0;
  for (const [pageIndex, page] of document.getPages().entries()) {
    const annots = resolvedArray(document, page.node.get(NAME_ANNOTS));
    if (!annots) continue;
    annotationCount += annots.size();
    if (annots.size() > MAX_PAGE_ANNOTATIONS || annotationCount > MAX_ANNOTATIONS) {
      throw new PdfSemanticAppearanceInventoryError("capacity-exceeded");
    }
    // A page's annotations are read together: each waits on its hashes,
    // and read one after another a page of a thousand links waited out
    
    const pageScans = await Promise.all(Array.from({ length: annots.size() }, async (_, annotationIndex): Promise<AnnotationScan> => {
      const entry = annots.get(annotationIndex);
      const ref = entry instanceof PDFRef ? entry : null;
      const dictionary = resolvedDict(document, entry);
      let objectEvidence: PdfSemanticAppearanceObjectRef | null = null;
      let dictionarySha256: string;
      if (ref) {
        try {
          objectEvidence = await refEvidence(ref, inspectObject);
          dictionarySha256 = objectEvidence.objectSha256;
        } catch {
          dictionarySha256 = await sha256Text(dictionary?.toString() ?? "null");
        }
      } else {
        dictionarySha256 = await sha256Text(
          dictionary?.toString() ?? String(entry),
        );
      }
      const subtype = dictionary
        ? exactName(dictionary.get(NAME_SUBTYPE))
        : null;
      let geometry: PdfSemanticAppearanceGeometry | null = null;
      if (dictionary) {
        try {
          geometry = normalizedGeometry(document, page, dictionary, subtype);
        } catch {
          geometry = null;
        }
      }
      return {
        sourceSha256,
        pageNumber: pageIndex + 1,
        pageRef: refCoordinates(page.ref),
        annotationIndex,
        annotationRef: ref,
        annotationDict: dictionary,
        annotationObject: objectEvidence,
        fallbackDictionarySha256: dictionarySha256,
        subtype,
        geometry,
      };
    }));
    scans.push(...pageScans);
  }

  const annotationRefCounts = new Map<string, number>();
  for (const scan of scans) {
    if (!scan.annotationRef) continue;
    const key = refKey(scan.annotationRef);
    annotationRefCounts.set(key, (annotationRefCounts.get(key) ?? 0) + 1);
  }
  const drafts: OwnerDraft[] = [];
  const preserveOnly: PdfSemanticAppearancePreserveOnly[] = [];
  // Every annotation read together, in order (the same reason as above).
  const classified = await Promise.all(scans.map(async (scan): Promise<
    Readonly<{ draft: OwnerDraft }> | Readonly<{ preserved: PdfSemanticAppearancePreserveOnly }>
  > => {
    try {
      if (
        !scan.annotationRef ||
        !scan.annotationObject ||
        !scan.annotationDict
      ) {
        throw new PreserveProblem("direct-annotation");
      }
      if ((annotationRefCounts.get(refKey(scan.annotationRef)) ?? 0) !== 1) {
        throw new PreserveProblem("duplicate-annotation");
      }
      if (!scan.geometry) throw new PreserveProblem("malformed-geometry");
      if (scan.subtype === "Widget") {
        const field = await fieldAuthority(
          document,
          scan.annotationRef,
          scan.annotationDict,
          inspectObject,
        );
        const appearance = await widgetAppearance(
          sourceBytes,
          document,
          scan.annotationDict,
          inspectObject,
        );
        const base: Omit<PdfSemanticWidgetAppearanceOwner, "id"> = {
          schemaVersion: PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION,
          sourceSha256,
          pageNumber: scan.pageNumber,
          pageRef: scan.pageRef,
          annotationIndex: scan.annotationIndex,
          annotation: scan.annotationObject,
          subtype: "Widget",
          geometry: scan.geometry,
          ownerKind: "widget",
          widget: Object.freeze({ ...field, appearance: appearance.evidence }),
        };
        return { draft: {
          scan,
          owner: Object.freeze(base),
          appearanceRefs: appearance.refs,
        } };
      }
      if (!supportedAnnotationSubtype(scan.subtype)) {
        throw new PreserveProblem("unsupported-subtype");
      }
      if (scan.annotationDict.get(NAME_AP)) {
        throw new PreserveProblem("stored-annotation-appearance");
      }
      const base: Omit<PdfSemanticSynthesizedAnnotationOwner, "id"> = {
        schemaVersion: PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION,
        sourceSha256,
        pageNumber: scan.pageNumber,
        pageRef: scan.pageRef,
        annotationIndex: scan.annotationIndex,
        annotation: scan.annotationObject,
        subtype: scan.subtype,
        geometry: scan.geometry,
        ownerKind: "annotation",
        annotationAppearance: Object.freeze({ kind: "synthesized-no-ap" }),
      };
      return { draft: {
        scan,
        owner: Object.freeze(base),
        appearanceRefs: Object.freeze([]),
      } };
    } catch (error) {
      const reason =
        error instanceof PreserveProblem
          ? error.reason
          : "malformed-appearance";
      return { preserved: await preserveRecord(scan, reason) };
    }
  }));
  for (const item of classified) {
    if ("draft" in item) drafts.push(item.draft);
    else preserveOnly.push(item.preserved);
  }

  const appearanceOwnerCounts = new Map<string, number>();
  for (const draft of drafts) {
    for (const ref of new Set(draft.appearanceRefs)) {
      appearanceOwnerCounts.set(ref, (appearanceOwnerCounts.get(ref) ?? 0) + 1);
    }
  }
  const owners: PdfSemanticAppearanceOwner[] = [];
  const exact: OwnerDraft[] = [];
  for (const draft of drafts) {
    const appearanceForms =
      draft.owner.ownerKind === "widget"
        ? draft.owner.widget.appearance.kind === "direct-form"
          ? [draft.owner.widget.appearance.form]
          : draft.owner.widget.appearance.entries.map((entry) => entry.form)
        : [];
    const shared = appearanceForms.some(
      (form) =>
        form.referenceCount !== 1 ||
        (appearanceOwnerCounts.get(`${form.objectNumber}:${form.generation}`) ??
          0) !== 1,
    );
    if (shared) {
      const preserved = await preserveRecord(
        draft.scan,
        "shared-appearance-stream",
      );
      preserveOnly.push(
        Object.freeze({
          ...preserved,
          ...(draft.owner.ownerKind === "widget"
            ? { widget: draft.owner.widget }
            : {}),
        }),
      );
      continue;
    }
    exact.push(draft);
  }
  const ids = await Promise.all(exact.map((draft) => ownerId(draft.owner)));
  exact.forEach((draft, index) => {
    owners.push(
      Object.freeze({ ...draft.owner, id: ids[index] }) as PdfSemanticAppearanceOwner,
    );
  });
  owners.sort(
    (left, right) =>
      left.pageNumber - right.pageNumber ||
      left.annotation.objectNumber - right.annotation.objectNumber ||
      left.annotationIndex - right.annotationIndex,
  );
  preserveOnly.sort(
    (left, right) =>
      left.pageNumber - right.pageNumber ||
      left.annotationIndex - right.annotationIndex ||
      left.id.localeCompare(right.id),
  );
  let sharedAppearanceCount = 0;
  let duplicateAnnotationCount = 0;
  let unsupportedCount = 0;
  let malformedCount = 0;
  for (const item of preserveOnly) {
    const category = preserveReasonCategory(item.reason);
    if (category === "shared") sharedAppearanceCount += 1;
    else if (category === "duplicate") duplicateAnnotationCount += 1;
    else if (category === "unsupported") unsupportedCount += 1;
    else malformedCount += 1;
  }
  return Object.freeze({
    schemaVersion: PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION,
    sourceSha256,
    pageCount: document.getPageCount(),
    owners: Object.freeze(owners),
    preserveOnly: Object.freeze(preserveOnly),
    diagnostics: Object.freeze({
      annotationCount,
      exactOwnerCount: owners.length,
      widgetOwnerCount: owners.filter((owner) => owner.ownerKind === "widget")
        .length,
      synthesizedOwnerCount: owners.filter(
        (owner) => owner.ownerKind === "annotation",
      ).length,
      directFormCount: owners.filter(
        (owner) =>
          owner.ownerKind === "widget" &&
          owner.widget.appearance.kind === "direct-form",
      ).length,
      stateDictionaryCount: owners.filter(
        (owner) =>
          owner.ownerKind === "widget" &&
          owner.widget.appearance.kind === "state-dictionary",
      ).length,
      preserveOnlyCount: preserveOnly.length,
      sharedAppearanceCount,
      duplicateAnnotationCount,
      unsupportedCount,
      malformedCount,
    }),
  });
}
