/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type PdfSemanticAppearanceOwner, type PdfSemanticAppearanceOwnerInventory, type PdfSemanticAppearancePreserveOnly } from "./semantic-appearance-owner-inventory";
import { PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION } from "./semantic-appearance-owner-schema";
import {
  PDF_SEMANTIC_APPEARANCE_RECONCILIATION_SCHEMA_VERSION,
  type PdfSemanticAppearanceInstanceDisposition,
  type PdfSemanticAppearanceInstanceEvidence,
  type PdfSemanticAppearanceInstanceReason,
  type PdfSemanticAppearanceMatchedOwner,
  type PdfSemanticAppearanceReconciliationResult,
  type PdfSemanticAppearanceVectorAssignment,
} from "./semantic-appearance-reconciliation";

export const PDF_SEMANTIC_APPEARANCE_COMMAND_ROUTING_SCHEMA_VERSION =
  "mrai.pdf-semantic-appearance-command-routing.v1" as const;

const SHA256_PATTERN = /^[0-9a-f]{64}$/i;
const MAX_ROUTING_RECORDS = 16_384;
const MAX_ROUTING_TARGETS = 16_384;
const MAX_RECONCILIATIONS = 4_096;
const MAX_INSTANCES_PER_RECONCILIATION = 4_096;
const MAX_VECTOR_EXCLUSIONS = 1_000_000;
const PRESERVE_ONLY_INSTANCE_REASONS =
  new Set<PdfSemanticAppearanceInstanceReason>([
    "unlocatable-annotation-id",
    "malformed-payload",
    "nested-block",
    "unbalanced-block",
    "duplicate-instance",
    "ambiguous-owner",
    "unmatched-owner",
    "rect-mismatch",
    "transform-mismatch",
    "matrix-mismatch",
    "appearance-state-mismatch",
    "unsupported-own-canvas",
    "vector-reprojection-failed",
    "vector-outside-owner",
  ]);

export type PdfSemanticAppearanceCommandKind =
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

export type PdfSemanticAppearanceCommandSurface =
  | "form-field"
  | "native-link"
  | "native-note"
  | "native-text-markup"
  | "native-shape";

/**
 * Sanitized identity of one editor surface that already exists in PDFViewer.
 * It deliberately carries no field name/value, annotation text, destination,
 * writer callback, or appearance-stream locator.
 */
export interface PdfSemanticAppearanceCommandTarget {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly objectNumber: number;
  readonly generation: number;
  readonly semanticKind: PdfSemanticAppearanceCommandKind;
  readonly surface: PdfSemanticAppearanceCommandSurface;
}

export type PdfSemanticAppearanceOwnerDisposition =
  "exact-owner" | "shared-appearance-preserve-only";

export type PdfSemanticAppearanceRouteEvidenceSource =
  "annotation-instance" | "source-owner-only";

export interface PdfSemanticAppearanceSelectionRoute {
  readonly selectionOnly: true;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly objectNumber: number;
  readonly generation: number;
  readonly semanticKind: PdfSemanticAppearanceCommandKind;
  readonly surface: PdfSemanticAppearanceCommandSurface;
  readonly ownerDisposition: PdfSemanticAppearanceOwnerDisposition;
  readonly evidenceSource: PdfSemanticAppearanceRouteEvidenceSource;
  /** Verbatim B2 disposition when a renderer instance exists; never invented. */
  readonly instanceDisposition: PdfSemanticAppearanceInstanceDisposition | null;
}

export type PdfSemanticAppearanceRouteFailureReason =
  | "duplicate-owner"
  | "target-missing"
  | "malformed-target"
  | "stale-target"
  | "target-kind-mismatch"
  | "duplicate-target"
  | "appearance-evidence-missing"
  | "instance-evidence-mismatch"
  | "ambiguous-instance"
  /**
   * The reader targeted an annotation the owner inventory never listed as a
   * routable record (a Square carrying a stored appearance the inventory
   * keeps preserve-only, an unsupported subtype, ...). The target is answered
   * so the caller's one-to-one contract holds; it simply has no route.
   */
  | "owner-missing";

export interface PdfSemanticAppearanceUnresolvedRoute {
  readonly pageNumber: number;
  readonly objectNumber: number;
  readonly generation: number;
  readonly semanticKind: PdfSemanticAppearanceCommandKind;
  readonly reason: PdfSemanticAppearanceRouteFailureReason;
}

/**
 * A renderer-local exclusion key only. This is not a native content-stream
 * operator span and can never be presented to a writer as mutation authority.
 */
export interface PdfSemanticAppearanceVectorExclusion {
  readonly pageNumber: number;
  readonly rendererOperatorIndex: number;
  readonly excludedFromGenericVectorPool: true;
  readonly disposition: PdfSemanticAppearanceInstanceDisposition;
}

export interface PdfSemanticAppearanceCommandRoutingResult {
  readonly schemaVersion: typeof PDF_SEMANTIC_APPEARANCE_COMMAND_ROUTING_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly routes: readonly PdfSemanticAppearanceSelectionRoute[];
  readonly unresolved: readonly PdfSemanticAppearanceUnresolvedRoute[];
  readonly vectorExclusions: readonly PdfSemanticAppearanceVectorExclusion[];
  readonly diagnostics: Readonly<{
    semanticRecordCount: number;
    routeCount: number;
    exactOwnerRouteCount: number;
    sharedPreserveRouteCount: number;
    instanceRouteCount: number;
    sourceOwnerOnlyRouteCount: number;
    unresolvedCount: number;
    ambiguousRouteCount: number;
    invalidTargetCount: number;
    excludedVectorCount: number;
    eligibleExcludedVectorCount: number;
    sharedPreserveExcludedVectorCount: number;
    preserveOnlyExcludedVectorCount: number;
  }>;
}

export interface PdfSemanticAppearanceCommandRoutingInput {
  readonly sourceSha256: string;
  readonly inventory: PdfSemanticAppearanceOwnerInventory;
  readonly reconciliations: readonly PdfSemanticAppearanceReconciliationResult[];
  readonly targets: readonly PdfSemanticAppearanceCommandTarget[];
}

export type PdfSemanticAppearanceCommandRoutingErrorCode =
  "source-mismatch" | "invalid-evidence" | "capacity-exceeded";

export class PdfSemanticAppearanceCommandRoutingError extends Error {
  constructor(readonly code: PdfSemanticAppearanceCommandRoutingErrorCode) {
    super(`PDF semantic appearance command routing failed: ${code}.`);
    this.name = "PdfSemanticAppearanceCommandRoutingError";
  }
}

interface RoutingRecord {
  readonly recordId: string;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly objectNumber: number;
  readonly generation: number;
  readonly semanticKind: PdfSemanticAppearanceCommandKind;
  readonly ownerKind: "widget" | "annotation";
  readonly ownerDisposition: PdfSemanticAppearanceOwnerDisposition;
}

type SanitizedTarget = PdfSemanticAppearanceCommandTarget;

function fail(code: PdfSemanticAppearanceCommandRoutingErrorCode): never {
  throw new PdfSemanticAppearanceCommandRoutingError(code);
}

function exactSha256(value: unknown): string {
  if (typeof value !== "string" || !SHA256_PATTERN.test(value.trim())) {
    return fail("invalid-evidence");
  }
  return value.trim().toLowerCase();
}

function isExactInteger(value: unknown, minimum: number): value is number {
  return Number.isSafeInteger(value) && (value as number) >= minimum;
}

function isSemanticKind(
  value: unknown,
): value is PdfSemanticAppearanceCommandKind {
  return (
    value === "text" ||
    value === "checkbox" ||
    value === "radio" ||
    value === "dropdown" ||
    value === "listbox" ||
    value === "Link" ||
    value === "Text" ||
    value === "Highlight" ||
    value === "Underline" ||
    value === "Square"
  );
}

function isSurface(
  value: unknown,
): value is PdfSemanticAppearanceCommandSurface {
  return (
    value === "form-field" ||
    value === "native-link" ||
    value === "native-note" ||
    value === "native-text-markup" ||
    value === "native-shape"
  );
}

function isInstanceDisposition(
  value: unknown,
): value is PdfSemanticAppearanceInstanceDisposition {
  return (
    value === "semantic-owner-eligible" ||
    value === "semantic-owner-preserve-only" ||
    value === "preserve-only"
  );
}

function isPreserveOnlyInstanceReason(
  value: unknown,
): value is PdfSemanticAppearanceInstanceReason {
  return (
    typeof value === "string" &&
    PRESERVE_ONLY_INSTANCE_REASONS.has(
      value as PdfSemanticAppearanceInstanceReason,
    )
  );
}

function surfaceForKind(
  kind: PdfSemanticAppearanceCommandKind,
): PdfSemanticAppearanceCommandSurface {
  if (
    kind === "text" ||
    kind === "checkbox" ||
    kind === "radio" ||
    kind === "dropdown" ||
    kind === "listbox"
  ) {
    return "form-field";
  }
  if (kind === "Link") return "native-link";
  if (kind === "Text") return "native-note";
  if (kind === "Highlight" || kind === "Underline") {
    return "native-text-markup";
  }
  return "native-shape";
}

function semanticKindForOwner(
  owner: PdfSemanticAppearanceOwner,
): PdfSemanticAppearanceCommandKind {
  const kind =
    owner.ownerKind === "widget" ? owner.widget.fieldKind : owner.subtype;
  if (!isSemanticKind(kind)) return fail("invalid-evidence");
  return kind;
}

function exactRecord(
  owner: PdfSemanticAppearanceOwner,
  sourceSha256: string,
  pageCount: number,
): RoutingRecord {
  if (
    owner.schemaVersion !== PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION ||
    exactSha256(owner.sourceSha256) !== sourceSha256 ||
    !owner.id.trim() ||
    !isExactInteger(owner.pageNumber, 1) ||
    owner.pageNumber > pageCount ||
    !isExactInteger(owner.annotation.objectNumber, 1) ||
    !isExactInteger(owner.annotation.generation, 0)
  ) {
    return fail("invalid-evidence");
  }
  return Object.freeze({
    recordId: owner.id,
    sourceSha256,
    pageNumber: owner.pageNumber,
    objectNumber: owner.annotation.objectNumber,
    generation: owner.annotation.generation,
    semanticKind: semanticKindForOwner(owner),
    ownerKind: owner.ownerKind,
    ownerDisposition: "exact-owner",
  });
}

function sharedRecord(
  item: PdfSemanticAppearancePreserveOnly,
  sourceSha256: string,
  pageCount: number,
): RoutingRecord | null {
  if (item.reason !== "shared-appearance-stream") return null;
  if (
    exactSha256(item.sourceSha256) !== sourceSha256 ||
    !item.id.trim() ||
    !item.widget ||
    !("objectNumber" in item.locator) ||
    !isExactInteger(item.pageNumber, 1) ||
    item.pageNumber > pageCount ||
    !isExactInteger(item.locator.objectNumber, 1) ||
    !isExactInteger(item.locator.generation, 0) ||
    !isSemanticKind(item.widget.fieldKind)
  ) {
    return fail("invalid-evidence");
  }
  return Object.freeze({
    recordId: item.id,
    sourceSha256,
    pageNumber: item.pageNumber,
    objectNumber: item.locator.objectNumber,
    generation: item.locator.generation,
    semanticKind: item.widget.fieldKind,
    ownerKind: "widget",
    ownerDisposition: "shared-appearance-preserve-only",
  });
}

function recordRefKey(
  record: Pick<RoutingRecord, "pageNumber" | "objectNumber" | "generation">,
): string {
  return `${record.pageNumber}:${record.objectNumber}:${record.generation}`;
}

function sanitizeTargets(
  targets: readonly PdfSemanticAppearanceCommandTarget[],
  pageCount: number,
): Readonly<{
  valid: readonly SanitizedTarget[];
  invalidRefKeys: ReadonlySet<string>;
  invalidCount: number;
}> {
  if (targets.length > MAX_ROUTING_TARGETS) return fail("capacity-exceeded");
  const valid: SanitizedTarget[] = [];
  const invalidRefKeys = new Set<string>();
  let invalidCount = 0;
  for (const target of targets) {
    const hasRefIdentity =
      target &&
      typeof target === "object" &&
      isExactInteger(target.pageNumber, 1) &&
      target.pageNumber <= pageCount &&
      isExactInteger(target.objectNumber, 1) &&
      isExactInteger(target.generation, 0);
    if (
      !target ||
      typeof target !== "object" ||
      typeof target.sourceSha256 !== "string" ||
      !SHA256_PATTERN.test(target.sourceSha256.trim()) ||
      !isExactInteger(target.pageNumber, 1) ||
      target.pageNumber > pageCount ||
      !isExactInteger(target.objectNumber, 1) ||
      !isExactInteger(target.generation, 0) ||
      !isSemanticKind(target.semanticKind) ||
      !isSurface(target.surface)
    ) {
      invalidCount += 1;
      if (hasRefIdentity) {
        invalidRefKeys.add(
          `${target.pageNumber}:${target.objectNumber}:${target.generation}`,
        );
      }
      continue;
    }
    valid.push(
      Object.freeze({
        sourceSha256: target.sourceSha256.trim().toLowerCase(),
        pageNumber: target.pageNumber,
        objectNumber: target.objectNumber,
        generation: target.generation,
        semanticKind: target.semanticKind,
        surface: target.surface,
      }),
    );
  }
  return Object.freeze({
    valid: Object.freeze(valid),
    invalidRefKeys,
    invalidCount,
  });
}

function targetFailureReason(
  record: RoutingRecord,
  targets: Readonly<{
    valid: readonly SanitizedTarget[];
    invalidRefKeys: ReadonlySet<string>;
  }>,
): PdfSemanticAppearanceRouteFailureReason | null {
  if (targets.invalidRefKeys.has(recordRefKey(record))) {
    return "malformed-target";
  }
  const related = targets.valid.filter(
    (target) =>
      target.pageNumber === record.pageNumber &&
      target.objectNumber === record.objectNumber &&
      target.generation === record.generation,
  );
  if (related.length === 0) return "target-missing";
  if (related.length > 1) return "duplicate-target";
  const target = related[0];
  if (target.sourceSha256 !== record.sourceSha256) return "stale-target";
  return target.semanticKind === record.semanticKind &&
    target.surface === surfaceForKind(record.semanticKind)
    ? null
    : "target-kind-mismatch";
}

function exactTarget(
  record: RoutingRecord,
  targets: readonly SanitizedTarget[],
): SanitizedTarget {
  const target = targets.find(
    (candidate) =>
      candidate.sourceSha256 === record.sourceSha256 &&
      candidate.pageNumber === record.pageNumber &&
      candidate.objectNumber === record.objectNumber &&
      candidate.generation === record.generation &&
      candidate.semanticKind === record.semanticKind &&
      candidate.surface === surfaceForKind(record.semanticKind),
  );
  if (!target) return fail("invalid-evidence");
  return target;
}

function instanceRefMatches(
  instance: PdfSemanticAppearanceInstanceEvidence,
  record: RoutingRecord,
): boolean {
  return (
    instance.annotationRef?.objectNumber === record.objectNumber &&
    instance.annotationRef.generation === record.generation
  );
}

function instanceEvidenceMatches(
  instance: PdfSemanticAppearanceInstanceEvidence,
  record: RoutingRecord,
): boolean {
  const owner = instance.owner;
  if (!owner || !instanceRefMatches(instance, record)) return false;
  const shared = record.ownerDisposition === "shared-appearance-preserve-only";
  const expectedDisposition: PdfSemanticAppearanceInstanceDisposition = shared
    ? "semantic-owner-preserve-only"
    : "semantic-owner-eligible";
  return (
    owner.recordId === record.recordId &&
    owner.objectNumber === record.objectNumber &&
    owner.generation === record.generation &&
    owner.ownerKind === record.ownerKind &&
    owner.semanticKind === record.semanticKind &&
    owner.sharedAppearance === shared &&
    instance.disposition === expectedDisposition &&
    (shared
      ? instance.reason === "shared-appearance-stream"
      : instance.reason === null)
  );
}

function matchedOwnerIsValid(
  owner: PdfSemanticAppearanceMatchedOwner | null,
): owner is PdfSemanticAppearanceMatchedOwner {
  return Boolean(
    owner &&
    typeof owner.recordId === "string" &&
    owner.recordId.trim() &&
    isExactInteger(owner.objectNumber, 1) &&
    isExactInteger(owner.generation, 0) &&
    (owner.ownerKind === "widget" || owner.ownerKind === "annotation") &&
    isSemanticKind(owner.semanticKind) &&
    typeof owner.sharedAppearance === "boolean",
  );
}

function sameMatchedOwner(
  left: PdfSemanticAppearanceMatchedOwner | null,
  right: PdfSemanticAppearanceMatchedOwner | null,
): boolean {
  if (left === null || right === null) return left === right;
  return (
    left.recordId === right.recordId &&
    left.objectNumber === right.objectNumber &&
    left.generation === right.generation &&
    left.ownerKind === right.ownerKind &&
    left.semanticKind === right.semanticKind &&
    left.sharedAppearance === right.sharedAppearance
  );
}

function instancePrimitiveIsValid(
  instance: PdfSemanticAppearanceInstanceEvidence,
): boolean {
  if (
    !isExactInteger(instance.instanceIndex, 0) ||
    !isInstanceDisposition(instance.disposition)
  ) {
    return false;
  }
  if (instance.disposition === "semantic-owner-eligible") {
    return (
      instance.reason === null &&
      matchedOwnerIsValid(instance.owner) &&
      !instance.owner.sharedAppearance
    );
  }
  if (instance.disposition === "semantic-owner-preserve-only") {
    return (
      instance.reason === "shared-appearance-stream" &&
      matchedOwnerIsValid(instance.owner) &&
      instance.owner.sharedAppearance
    );
  }
  return (
    isPreserveOnlyInstanceReason(instance.reason) && instance.owner === null
  );
}

function assignmentMatchesInstance(
  assignment: PdfSemanticAppearanceVectorAssignment,
  instance: PdfSemanticAppearanceInstanceEvidence,
): boolean {
  return (
    assignment.instanceIndex === instance.instanceIndex &&
    assignment.disposition === instance.disposition &&
    assignment.reason === instance.reason &&
    sameMatchedOwner(assignment.owner, instance.owner)
  );
}

function validatedVectorExclusions(
  result: PdfSemanticAppearanceReconciliationResult,
): readonly PdfSemanticAppearanceVectorExclusion[] {
  if (result.vectorAssignments.length > MAX_VECTOR_EXCLUSIONS) {
    return fail("capacity-exceeded");
  }
  const instancesByIndex = new Map<
    number,
    Readonly<{
      instance: PdfSemanticAppearanceInstanceEvidence;
      visibleOperatorIndexes: ReadonlySet<number>;
    }>
  >();
  const instanceByOperator = new Map<number, number>();
  const requiredAssignments = new Set<string>();
  for (const instance of result.instances) {
    if (
      !instancePrimitiveIsValid(instance) ||
      instancesByIndex.has(instance.instanceIndex)
    ) {
      return fail("invalid-evidence");
    }
    const visibleOperatorIndexes = new Set<number>();
    for (const operatorIndex of instance.visibleOperatorIndexes) {
      if (
        !isExactInteger(operatorIndex, 0) ||
        visibleOperatorIndexes.has(operatorIndex) ||
        instanceByOperator.has(operatorIndex) ||
        requiredAssignments.size >= MAX_VECTOR_EXCLUSIONS
      ) {
        return fail("invalid-evidence");
      }
      visibleOperatorIndexes.add(operatorIndex);
      instanceByOperator.set(operatorIndex, instance.instanceIndex);
      requiredAssignments.add(`${instance.instanceIndex}:${operatorIndex}`);
    }
    instancesByIndex.set(
      instance.instanceIndex,
      Object.freeze({ instance, visibleOperatorIndexes }),
    );
  }

  const assignmentOperators = new Set<number>();
  const observedAssignments = new Set<string>();
  const exclusions: PdfSemanticAppearanceVectorExclusion[] = [];
  for (const assignment of result.vectorAssignments) {
    if (
      !isExactInteger(assignment.visibleOperatorIndex, 0) ||
      !isInstanceDisposition(assignment.disposition) ||
      assignmentOperators.has(assignment.visibleOperatorIndex) ||
      (assignment.instanceIndex !== null &&
        !isExactInteger(assignment.instanceIndex, 0))
    ) {
      return fail("invalid-evidence");
    }
    assignmentOperators.add(assignment.visibleOperatorIndex);
    if (assignment.instanceIndex === null) {
      if (
        assignment.disposition !== "preserve-only" ||
        assignment.reason !== "outside-annotation-block" ||
        assignment.owner !== null ||
        instanceByOperator.has(assignment.visibleOperatorIndex)
      ) {
        return fail("invalid-evidence");
      }
      continue;
    }

    const bound = instancesByIndex.get(assignment.instanceIndex);
    const assignmentKey = `${assignment.instanceIndex}:${assignment.visibleOperatorIndex}`;
    if (
      !bound ||
      !bound.visibleOperatorIndexes.has(assignment.visibleOperatorIndex) ||
      !assignmentMatchesInstance(assignment, bound.instance) ||
      observedAssignments.has(assignmentKey)
    ) {
      return fail("invalid-evidence");
    }
    observedAssignments.add(assignmentKey);
    exclusions.push(
      Object.freeze({
        pageNumber: result.pageNumber,
        rendererOperatorIndex: assignment.visibleOperatorIndex,
        excludedFromGenericVectorPool: true,
        disposition: assignment.disposition,
      }),
    );
  }
  for (const required of requiredAssignments) {
    if (!observedAssignments.has(required)) return fail("invalid-evidence");
  }
  return Object.freeze(exclusions);
}

function unresolved(
  record: RoutingRecord,
  reason: PdfSemanticAppearanceRouteFailureReason,
): PdfSemanticAppearanceUnresolvedRoute {
  return Object.freeze({
    pageNumber: record.pageNumber,
    objectNumber: record.objectNumber,
    generation: record.generation,
    semanticKind: record.semanticKind,
    reason,
  });
}

function ambiguousFailure(reason: PdfSemanticAppearanceRouteFailureReason) {
  return (
    reason === "duplicate-owner" ||
    reason === "duplicate-target" ||
    reason === "ambiguous-instance"
  );
}

/**
 * Reconcile source-owned semantic appearances with selection surfaces that are
 * already mounted by the PDF reader. A successful route is still only a focus
 * destination: every later mutation must independently re-check its existing
 * field/annotation writer and current source authority.
 */
export function routePdfSemanticAppearanceCommands(
  input: PdfSemanticAppearanceCommandRoutingInput,
): PdfSemanticAppearanceCommandRoutingResult {
  const sourceSha256 = exactSha256(input.sourceSha256);
  const { inventory, reconciliations, targets } = input;
  if (
    inventory.schemaVersion !== PDF_SEMANTIC_APPEARANCE_OWNER_SCHEMA_VERSION ||
    exactSha256(inventory.sourceSha256) !== sourceSha256 ||
    !isExactInteger(inventory.pageCount, 1)
  ) {
    return fail("source-mismatch");
  }
  if (
    inventory.owners.length + inventory.preserveOnly.length >
      MAX_ROUTING_RECORDS ||
    reconciliations.length > MAX_RECONCILIATIONS
  ) {
    return fail("capacity-exceeded");
  }

  const records = [
    ...inventory.owners.map((owner) =>
      exactRecord(owner, sourceSha256, inventory.pageCount),
    ),
    ...inventory.preserveOnly.flatMap((item) => {
      const record = sharedRecord(item, sourceSha256, inventory.pageCount);
      return record ? [record] : [];
    }),
  ];
  const recordsByRef = new Map<string, RoutingRecord[]>();
  for (const record of records) {
    const key = recordRefKey(record);
    recordsByRef.set(key, [...(recordsByRef.get(key) ?? []), record]);
  }

  const reconciliationByPage = new Map<
    number,
    PdfSemanticAppearanceReconciliationResult
  >();
  const vectorExclusions: PdfSemanticAppearanceVectorExclusion[] = [];
  const exclusionKeys = new Set<string>();
  for (const result of reconciliations) {
    if (
      result.schemaVersion !==
        PDF_SEMANTIC_APPEARANCE_RECONCILIATION_SCHEMA_VERSION ||
      !isExactInteger(result.pageNumber, 1) ||
      result.pageNumber > inventory.pageCount ||
      reconciliationByPage.has(result.pageNumber) ||
      result.instances.length > MAX_INSTANCES_PER_RECONCILIATION
    ) {
      return fail("invalid-evidence");
    }
    if (exactSha256(result.sourceSha256) !== sourceSha256) {
      return fail("source-mismatch");
    }
    reconciliationByPage.set(result.pageNumber, result);
    for (const exclusion of validatedVectorExclusions(result)) {
      const key = `${result.pageNumber}:${exclusion.rendererOperatorIndex}`;
      if (exclusionKeys.has(key)) return fail("invalid-evidence");
      exclusionKeys.add(key);
      if (vectorExclusions.length >= MAX_VECTOR_EXCLUSIONS) {
        return fail("capacity-exceeded");
      }
      vectorExclusions.push(exclusion);
    }
  }

  const sanitizedTargets = sanitizeTargets(targets, inventory.pageCount);
  const routes: PdfSemanticAppearanceSelectionRoute[] = [];
  const unresolvedRoutes: PdfSemanticAppearanceUnresolvedRoute[] = [];
  for (const record of records) {
    if ((recordsByRef.get(recordRefKey(record)) ?? []).length !== 1) {
      unresolvedRoutes.push(unresolved(record, "duplicate-owner"));
      continue;
    }
    const targetReason = targetFailureReason(record, sanitizedTargets);
    if (targetReason) {
      unresolvedRoutes.push(unresolved(record, targetReason));
      continue;
    }

    const pageEvidence = reconciliationByPage.get(record.pageNumber);
    const instances = pageEvidence
      ? pageEvidence.instances.filter((instance) =>
          instanceRefMatches(instance, record),
        )
      : [];
    let evidenceSource: PdfSemanticAppearanceRouteEvidenceSource =
      "source-owner-only";
    let instanceDisposition: PdfSemanticAppearanceInstanceDisposition | null =
      null;
    if (instances.length > 1) {
      unresolvedRoutes.push(unresolved(record, "ambiguous-instance"));
      continue;
    }
    if (instances.length === 1) {
      const instance = instances[0];
      if (!instanceEvidenceMatches(instance, record)) {
        unresolvedRoutes.push(unresolved(record, "instance-evidence-mismatch"));
        continue;
      }
      evidenceSource = "annotation-instance";
      instanceDisposition = instance.disposition;
    } else if (pageEvidence && record.semanticKind !== "Link") {
      unresolvedRoutes.push(unresolved(record, "appearance-evidence-missing"));
      continue;
    }

    const target = exactTarget(record, sanitizedTargets.valid);
    routes.push(
      Object.freeze({
        selectionOnly: true,
        sourceSha256,
        pageNumber: target.pageNumber,
        objectNumber: target.objectNumber,
        generation: target.generation,
        semanticKind: target.semanticKind,
        surface: target.surface,
        ownerDisposition: record.ownerDisposition,
        evidenceSource,
        instanceDisposition,
      }),
    );
  }

  // Every target the caller sent is answered exactly once, as a route or as
  // an unresolved entry. A target whose annotation is not a routing record
  // at all (preserve-only for a reason other than a shared appearance stream,
  // or not in the inventory) would otherwise vanish from both lists, and the
  // session that asked treats a silently dropped target as a broken worker.
  const recordRefKeys = new Set(records.map(recordRefKey));
  for (const target of sanitizedTargets.valid) {
    if (recordRefKeys.has(recordRefKey(target))) continue;
    unresolvedRoutes.push(
      Object.freeze({
        pageNumber: target.pageNumber,
        objectNumber: target.objectNumber,
        generation: target.generation,
        semanticKind: target.semanticKind,
        reason: "owner-missing" as const,
      }),
    );
  }

  routes.sort(
    (left, right) =>
      left.pageNumber - right.pageNumber ||
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation ||
      left.semanticKind.localeCompare(right.semanticKind),
  );
  unresolvedRoutes.sort(
    (left, right) =>
      left.pageNumber - right.pageNumber ||
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation ||
      left.semanticKind.localeCompare(right.semanticKind) ||
      left.reason.localeCompare(right.reason),
  );
  vectorExclusions.sort(
    (left, right) =>
      left.pageNumber - right.pageNumber ||
      left.rendererOperatorIndex - right.rendererOperatorIndex,
  );

  return Object.freeze({
    schemaVersion: PDF_SEMANTIC_APPEARANCE_COMMAND_ROUTING_SCHEMA_VERSION,
    sourceSha256,
    routes: Object.freeze(routes),
    unresolved: Object.freeze(unresolvedRoutes),
    vectorExclusions: Object.freeze(vectorExclusions),
    diagnostics: Object.freeze({
      semanticRecordCount: records.length,
      routeCount: routes.length,
      exactOwnerRouteCount: routes.filter(
        (route) => route.ownerDisposition === "exact-owner",
      ).length,
      sharedPreserveRouteCount: routes.filter(
        (route) => route.ownerDisposition === "shared-appearance-preserve-only",
      ).length,
      instanceRouteCount: routes.filter(
        (route) => route.evidenceSource === "annotation-instance",
      ).length,
      sourceOwnerOnlyRouteCount: routes.filter(
        (route) => route.evidenceSource === "source-owner-only",
      ).length,
      unresolvedCount: unresolvedRoutes.length,
      ambiguousRouteCount: unresolvedRoutes.filter((route) =>
        ambiguousFailure(route.reason),
      ).length,
      invalidTargetCount: sanitizedTargets.invalidCount,
      excludedVectorCount: vectorExclusions.length,
      eligibleExcludedVectorCount: vectorExclusions.filter(
        (item) => item.disposition === "semantic-owner-eligible",
      ).length,
      sharedPreserveExcludedVectorCount: vectorExclusions.filter(
        (item) => item.disposition === "semantic-owner-preserve-only",
      ).length,
      preserveOnlyExcludedVectorCount: vectorExclusions.filter(
        (item) => item.disposition === "preserve-only",
      ).length,
    }),
  });
}
