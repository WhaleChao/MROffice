/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The field a page-made object's provenance lives in. */
export const OFFICE_PROVENANCE_FIELD = "provenance";

/**
 * The keys that name where an object sits in its imported package (the
 * workbook drawing reader's and the slide reader's). A `source` holding none
 * of them names no package part.
 */
export const OFFICE_IMPORT_IDENTITY_KEYS: ReadonlySet<string> = new Set([
  "part",
  "drawingPart",
  "relationshipId",
  "nonVisualId",
  "controlShapeId",
  "sheetPart",
  "pictureFillPart",
  "linkedPicture",
  "sourceAnchor",
  "slideId",
  "slidePart",
  "shapeId",
  "inheritedFrom",
]);

export const DATA_ITEM_PROVENANCE = "dataItem";

export type OfficeObjectProvenance = Readonly<
  { kind: string } & Record<string, string | number | boolean | null>
>;

const MAX_PROVENANCE_KEYS = 16;
const MAX_PROVENANCE_TEXT = 500;
/** What marks a `source` as a provenance record: its `kind`, or (an older
 *  record without one) the workspace file it names. */
const PROVENANCE_MARKERS = ["kind", "dataItemId", "data_item_id"] as const;

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function plainValue(value: unknown): value is string | number | boolean | null {
  return (
    value === null ||
    typeof value === "boolean" ||
    (typeof value === "number" && Number.isFinite(value)) ||
    (typeof value === "string" && value.length <= MAX_PROVENANCE_TEXT)
  );
}

function namesPackagePart(record: Record<string, unknown>): boolean {
  return Object.keys(record).some((key) => OFFICE_IMPORT_IDENTITY_KEYS.has(key));
}

/** The provenance of content taken from a workspace file (DataItem). */
export function dataItemProvenance(dataItemId: string): OfficeObjectProvenance {
  return { kind: DATA_ITEM_PROVENANCE, dataItemId };
}

/** A well-formed provenance record: a `kind` and plain values, no key that
 *  names a package part (it can never pass for an identity). */
export function isOfficeProvenanceRecord(value: unknown): value is OfficeObjectProvenance {
  if (!isRecord(value)) return false;
  const keys = Object.keys(value);
  if (keys.length === 0 || keys.length > MAX_PROVENANCE_KEYS) return false;
  const kind = value.kind;
  if (typeof kind !== "string" || !kind.trim() || kind.length > 64) return false;
  return !namesPackagePart(value) && keys.every((key) => plainValue(value[key]));
}

/** A `source` that only records where the page took the object from (any
 *  kind): no package identity. */
export function sourceIsProvenanceRecord(source: unknown): boolean {
  return (
    isRecord(source) &&
    PROVENANCE_MARKERS.some((marker) => marker in source) &&
    !namesPackagePart(source)
  );
}

/** No package identity: no `source`, an empty one, or only a provenance record. */
export function sourceIsPageProvenance(source: unknown): boolean {
  if (source === undefined || source === null) return true;
  if (!isRecord(source)) return false;
  return Object.keys(source).length === 0 || sourceIsProvenanceRecord(source);
}

/** Whether an object is one an import placed: its `source` names its
 *  package part. Everything else is the page's own. */
export function officeObjectHasImportIdentity(object: unknown): boolean {
  return isRecord(object) && !sourceIsPageProvenance(object.source);
}

/** What a writer may put in `source`: an import identity, naming at least
 *  one package key and recording no provenance. */
export function sourceIsValidImportIdentity(source: unknown): boolean {
  if (!isRecord(source) || "kind" in source) return false;
  return [...OFFICE_IMPORT_IDENTITY_KEYS].some((key) => {
    const value = source[key];
    return value !== undefined && value !== null && value !== false && value !== "";
  });
}

function provenanceOfSource(source: Record<string, unknown>): OfficeObjectProvenance | null {
  const record: Record<string, string | number | boolean | null> = {};
  for (const [key, value] of Object.entries(source)) {
    if (plainValue(value)) record[key] = value;
  }
  const named = record.dataItemId ?? record.data_item_id;
  const withKind =
    !("kind" in record) && typeof named === "string" && named
      ? { kind: DATA_ITEM_PROVENANCE, ...record }
      : record;
  return isOfficeProvenanceRecord(withKind) ? withKind : null;
}






export function officeObjectProvenance(object: unknown): OfficeObjectProvenance | null {
  if (!isRecord(object)) return null;
  const recorded = object[OFFICE_PROVENANCE_FIELD];
  if (isOfficeProvenanceRecord(recorded)) return { ...recorded };
  const source = object.source;
  if (isRecord(source) && sourceIsProvenanceRecord(source)) return provenanceOfSource(source);
  const assetId = object.assetId;
  if (object.sourceKind === "asset" && typeof assetId === "string" && assetId.trim()) {
    return dataItemProvenance(assetId.trim());
  }
  return null;
}

/**
 * `object` with a provenance record kept in `source` moved into
 * `provenance` (a provenance already recorded wins), or `object` itself
 * when it holds none. A real identity is never touched.
 */
export function withProvenanceOutOfSource<T extends object>(object: T): T {
  const source = (object as { source?: unknown }).source;
  if (!isRecord(source) || !sourceIsProvenanceRecord(source)) return object;
  const next: Record<string, unknown> = { ...(object as Record<string, unknown>) };
  delete next.source;
  if (!isOfficeProvenanceRecord(next[OFFICE_PROVENANCE_FIELD])) {
    const record = provenanceOfSource(source);
    if (record) next[OFFICE_PROVENANCE_FIELD] = record;
  }
  return next as T;
}

/** `wanted`, or the sheet's own form of it when another object holds it
 *  (the API's `free_workbook_object_id`; the AI workbook studio's rule). */
export function freeWorkbookObjectId(wanted: string, sheetId: string, taken: ReadonlySet<string>): string {
  if (!taken.has(wanted)) return wanted;
  const base = `${sheetId}-${wanted}`.slice(0, 190);
  let candidate = base;
  for (let suffix = 2; taken.has(candidate); suffix += 1) candidate = `${base}-${suffix}`;
  return candidate;
}
