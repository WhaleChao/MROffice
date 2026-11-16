/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Several layer moves of a page's source vectors written as ONE revision,
 * one Undo step (Office: one Arrange command is one Undo step whatever it
 * moves; a name dragged in the Selection pane is one Undo step too).
 *
 * The engine's own reorder (planPdfNativeVectorSelectionReorder) moves one
 * object one step or to an end, each a proven write. A chain plans those
 * moves one after another on the bytes the one before left and hands the
 * result back as one revision (planPdfNativeRevisionChain), so an object goes
 * to the k-th place, or a multiple selection moves together, in one write.
 * Each step is the same proven transform the single command writes; the
 * recipe the working copy publishes is the list of those steps, so the
 * revision is planned again step by step and checked as any group edit is.
 *
 * Office's rules for a multiple selection (PowerPoint / Word, Arrange):
 * Bring to Front and Send to Back take every selected object to that end
 * keeping their order among themselves; Bring Forward and Send Backward move
 * each selected object one place, the selected objects never passing one
 * another (an object already at the end, or behind another selected object
 * that cannot move, stays).
 */
import type { PdfNativeChainEngine, PdfNativeChainStep } from "./native-revision-chain";
import { planPdfNativeRevisionChain } from "./native-revision-chain";
import {
  MRAI_NATIVE_CONTENT_UNIT_ALREADY_BACK,
  MRAI_NATIVE_CONTENT_UNIT_ALREADY_FRONT,
  MRAI_NATIVE_CONTENT_UNIT_ROWS_DIFFER,
  type PdfNativeContentUnitIdentity,
} from "./engine/native/content-unit-model";
import {
  MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK,
  MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT,
} from "./engine/native/embedded-image-arrange-model";
import { pdfNativeEmbeddedImagePaintMatrixMatches } from "./engine/native/embedded-image-model";
import { MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION } from "./engine/native/vocabulary";
import type { PdfNativePreparedContentStreamTransaction } from "./engine/native/content-stream-transaction";
import type {
  PdfNativeContentUnitReorderPlan,
} from "./engine/native/vector-content-transaction";
import type {
  PdfNativeEmbeddedImageReorderPayload,
  PdfNativeEmbeddedImageReorderPlanResult,
} from "./native-embedded-image-reorder";
import {
  MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
  createPdfNativeVectorHistoryEndpoint,
  pdfNativeVectorReorderRefusal,
} from "./native-vector-commands";
import type {
  PdfNativeVectorHistoryEndpoint,
  PdfNativeVectorReorderDirection,
  PdfNativeVectorReorderPayload,
  PdfNativeVectorSelectionReorderPlan,
  PdfNativeVectorWriteSelection,
} from "./native-vector-selection-edit";
import type { PdfVectorPageGeometry } from "./vector-capability-reconciliation";
import type { PdfNativeVectorSelectionRoute } from "./semantic-appearance-inspection";

/** A reorder of several moves the viewer asks the window for: the members
 * as the reader routes them, back to front (the order the page paints
 * them), and the moves by place in that list. */
export interface PdfNativeVectorReorderGroupCommitRequest {
  readonly routes: readonly PdfNativeVectorSelectionRoute[];
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly moves: readonly PdfNativeVectorReorderChainMove[];
}

/** A layer move of a page's objects of any kind the reader asks the window
 * for (a Selection pane drop, a picture's Arrange command): the members as
 * the reader names them, back to front, and the moves by place in that list.
 * `sourceSha256` is the bytes the reader read them from. */
export interface PdfNativeContentReorderCommitRequest {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  /** The page geometry the vectors' planner reads (null without vectors). */
  readonly pageGeometry: PdfVectorPageGeometry | null;
  readonly members: readonly PdfNativeContentReorderRequestMember[];
  readonly moves: readonly PdfNativeVectorReorderChainMove[];
}

export type PdfNativeContentReorderRequestMember =
  | Readonly<{ kind: "vector"; route: PdfNativeVectorSelectionRoute }>
  | Readonly<{ kind: "image"; pageNumber: number; imageOrder: number; matrix: readonly number[] | null }>
  | Readonly<{
      kind: "unit";
      pageNumber: number;
      unit: Readonly<{ kind: PdfNativeContentUnitIdentity["kind"]; order: number }>;
      /** The page's rows as the reader read them (pdfPageObjectRows). */
      kinds: string;
    }>;

/** What a reorder came to: the members where the result has them, or why
 * nothing moved (`end`: every member already stands at that end). */
export type PdfNativeContentReorderOutcome =
  | Readonly<{ ok: true; members: readonly PdfNativeContentReorderMember[]; sourceSha256: string }>
  | Readonly<{ ok: false; end?: "front" | "back"; reason?: string }>;

/** Which end a planner said an object already stands at, of any kind. */
export function pdfNativeContentReorderArrangedEnd(error: unknown): "front" | "back" | null {
  return arrangedEnd(error);
}

/** One single-object move of the chain: `member` is a place in the chain's
 * member list. With `yieldToMembers`, the move is left out when it would
 * pass another member or the object is already at that end (Office's
 * multi-object Bring Forward / Send Backward); otherwise such a refusal
 * refuses the whole chain, nothing half made. */
export interface PdfNativeVectorReorderChainMove {
  readonly member: number;
  readonly direction: PdfNativeVectorReorderDirection;
  readonly yieldToMembers?: boolean;
}

/** The planners a chain reads (the engine's own, or the browser's remote
 * callers of them). */
export interface PdfNativeVectorReorderChainPlanners {
  readonly resolvePdfNativeVectorHistorySelection: (input: {
    readonly sourceBytes: Uint8Array;
    readonly endpoint: PdfNativeVectorHistoryEndpoint;
  }) => Promise<PdfNativeVectorWriteSelection>;
  readonly planPdfNativeVectorSelectionReorder: (input: {
    readonly sourceBytes: Uint8Array;
    readonly payload: PdfNativeVectorReorderPayload;
  }) => Promise<PdfNativeVectorSelectionReorderPlan>;
}

/** The planners a chain of any page objects reads besides the vectors': the
 * picture's own reorder (the one its Arrange command writes), and the layer
 * move of a text object, a shading, a Form or a group. */
export interface PdfNativeContentReorderChainPlanners extends PdfNativeVectorReorderChainPlanners {
  readonly planPdfNativeEmbeddedImageReorder?: (input: {
    readonly sourceBytes: Uint8Array;
    readonly payload: PdfNativeEmbeddedImageReorderPayload;
  }) => Promise<PdfNativeEmbeddedImageReorderPlanResult>;
  readonly listMraiNativeEmbeddedImagePaintMatrices?: (input: {
    readonly sourceBytes: Uint8Array;
    readonly pageNumber: number;
  }) => Promise<readonly (readonly number[] | null)[]>;
  readonly planMraiNativeContentUnitReorder?: (input: {
    readonly sourceBytes: Uint8Array;
    readonly pageNumber: number;
    readonly unit: PdfNativeContentUnitIdentity;
    readonly direction: PdfNativeVectorReorderDirection;
  }) => Promise<PdfNativeContentUnitReorderPlan>;
  readonly listMraiNativePageContentUnits?: (input: {
    readonly sourceBytes: Uint8Array;
    readonly pageNumber: number;
  }) => Promise<Readonly<{ kinds: string; rows: readonly Readonly<{ kind: string; digest: string }>[] }>>;
}

/**
 * One object a chain moves: a source vector (by its history endpoint), a
 * picture (by its paint order, and the matrix it is drawn at, which finds it
 * again once other moves shifted its order), or a text object, shading, Form
 * or group (by its kind, place and digest; `kinds` is the page's row
 * sequence as the reader showed it, checked before the chain moves anything).
 */
export type PdfNativeContentReorderMember =
  | Readonly<{ kind: "vector"; endpoint: PdfNativeVectorHistoryEndpoint }>
  | Readonly<{ kind: "image"; pageNumber: number; imageOrder: number; matrix: readonly number[] | null }>
  | Readonly<{
      kind: "unit";
      pageNumber: number;
      /** Without its digest only with `kinds`: the row at its place then. */
      unit: Omit<PdfNativeContentUnitIdentity, "digest"> & Readonly<{ digest?: string }>;
      kinds?: string;
    }>;

/** One step of a chain: the planner's own plan of one single-object move. */
export type PdfNativeContentReorderStep = Readonly<
  | { kind: "vector"; plan: PdfNativeVectorSelectionReorderPlan; prepared: PdfNativeVectorSelectionReorderPlan["prepared"] }
  | { kind: "image"; plan: PdfNativeEmbeddedImageReorderPlanResult; prepared: PdfNativePreparedContentStreamTransaction }
  | { kind: "unit"; plan: PdfNativeContentUnitReorderPlan; prepared: PdfNativePreparedContentStreamTransaction }
>;

export interface PdfNativeContentReorderChainResult {
  readonly steps: readonly PdfNativeChainStep<PdfNativeContentReorderStep>[];
  readonly candidateBytes: Uint8Array;
  readonly candidateSha256: string;
  /** Every member as the result names it, in the order given. */
  readonly members: readonly PdfNativeContentReorderMember[];
  /** Every member as the head names it (each with its digest): the chain's
   * recipe, planned again on other bytes. */
  readonly startMembers: readonly PdfNativeContentReorderMember[];
}

export interface PdfNativeVectorReorderChainResult {
  readonly steps: readonly PdfNativeChainStep<PdfNativeVectorSelectionReorderPlan>[];
  readonly candidateBytes: Uint8Array;
  readonly candidateSha256: string;
  /** Every member as the result names it, in the order given. */
  readonly memberEndpoints: readonly PdfNativeVectorHistoryEndpoint[];
}

/**
 * Office's multi-object Arrange command as single-object moves, for members
 * listed back to front (the order the page paints them): to the front from
 * the back-most up, so each lands on the one before it; to the back from the
 * front-most down; one step forward from the front-most down and backward
 * from the back-most up, each yielding to the members around it.
 */
export function pdfNativeVectorArrangeMoves(
  memberCount: number,
  command: PdfNativeVectorReorderDirection,
): readonly PdfNativeVectorReorderChainMove[] {
  if (!Number.isSafeInteger(memberCount) || memberCount < 1) return [];
  const backToFront = Array.from({ length: memberCount }, (_unused, index) => index);
  const frontToBack = [...backToFront].reverse();
  switch (command) {
    case "front":
      return backToFront.map((member) => ({ member, direction: "front" as const }));
    case "back":
      return frontToBack.map((member) => ({ member, direction: "back" as const }));
    case "forward":
      return frontToBack.map((member) => ({ member, direction: "forward" as const, yieldToMembers: true }));
    case "backward":
      return backToFront.map((member) => ({ member, direction: "backward" as const, yieldToMembers: true }));
  }
}

/** The member's place in its family once `moved` went from `before` to
 * `after` (the family's list shifts by one between the two places). */
function shiftedProofIndex(
  member: PdfNativeVectorHistoryEndpoint,
  before: PdfNativeVectorHistoryEndpoint,
  after: PdfNativeVectorHistoryEndpoint,
): PdfNativeVectorHistoryEndpoint {
  if (member.proofFamily !== before.proofFamily || member.pageNumber !== before.pageNumber) return member;
  const from = before.pageProofIndex;
  const to = after.pageProofIndex;
  const index = member.pageProofIndex;
  let next = index;
  if (from < to && index > from && index <= to) next = index - 1;
  else if (to < from && index >= to && index < from) next = index + 1;
  return next === index ? member : Object.freeze({ ...member, pageProofIndex: next });
}

/** A place `index` in a list once the entry at `from` moved to `to`. */
function shiftedPlace(index: number, from: number, to: number): number {
  if (from < to && index > from && index <= to) return index - 1;
  if (to < from && index >= to && index < from) return index + 1;
  return index;
}

const sameObject = (left: PdfNativeVectorWriteSelection, right: PdfNativeVectorWriteSelection) =>
  left.proofFamily === right.proofFamily &&
  left.contentStreamObjectNumber === right.contentStreamObjectNumber &&
  left.contentStreamGeneration === right.contentStreamGeneration &&
  left.operatorStart === right.operatorStart &&
  left.operatorEnd === right.operatorEnd &&
  (left.formInstance ?? null) === (right.formInstance ?? null);

/** The farthest a member is looked for from its predicted place (an object of
 * the other family moving shifts a family's places the prediction does not
 * see: a whole path passing a member's subpath). */
const MEMBER_SEARCH_RADIUS = 64;

/**
 * The members on `sourceBytes`: each at its predicted place, else the
 * nearest place of its family that names it. Two members found on the same
 * object, or one found nowhere, refuse the chain.
 */
async function resolveMembers(
  planners: PdfNativeVectorReorderChainPlanners,
  sourceBytes: Uint8Array,
  endpoints: readonly PdfNativeVectorHistoryEndpoint[],
): Promise<PdfNativeVectorWriteSelection[]> {
  const found: PdfNativeVectorWriteSelection[] = [];
  for (const endpoint of endpoints) {
    let selection: PdfNativeVectorWriteSelection | null = null;
    for (let distance = 0; distance <= MEMBER_SEARCH_RADIUS && !selection; distance += 1) {
      for (const offset of distance === 0 ? [0] : [distance, -distance]) {
        const pageProofIndex = endpoint.pageProofIndex + offset;
        if (pageProofIndex < 0) continue;
        try {
          const candidate = await planners.resolvePdfNativeVectorHistorySelection({
            sourceBytes,
            endpoint: offset === 0 ? endpoint : Object.freeze({ ...endpoint, pageProofIndex }),
          });
          if (!found.some((other) => sameObject(other, candidate))) {
            selection = candidate;
            break;
          }
        } catch {
          // Not at this place.
        }
      }
    }
    if (!selection) throw new Error("PDF vector reorder member is stale or ambiguous.");
    found.push(selection);
  }
  return found;
}

/** Whether a planned step passes (or swaps with) another member. */
function passesMember(
  plan: PdfNativeVectorSelectionReorderPlan,
  moved: PdfNativeVectorWriteSelection,
  others: readonly PdfNativeVectorWriteSelection[],
): boolean {
  const neighbour = plan.evidence.neighbour;
  if (neighbour && others.some((other) => sameObject(other, neighbour))) return true;
  const passed = plan.evidence.passed;
  if (!passed) return false;
  return others.some(
    (other) =>
      other.contentStreamObjectNumber === moved.contentStreamObjectNumber &&
      other.contentStreamGeneration === moved.contentStreamGeneration &&
      (other.formInstance ?? null) === (moved.formInstance ?? null) &&
      other.operatorStart >= passed.start &&
      other.operatorEnd <= passed.end,
  );
}

/** Which end a planner said an object already stands at, of any kind. */
function arrangedEnd(error: unknown): "front" | "back" | null {
  const refusal = pdfNativeVectorReorderRefusal(error);
  if (refusal === "at-front") return "front";
  if (refusal === "at-back") return "back";
  const message = error instanceof Error ? error.message : "";
  if (message === MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT || message === MRAI_NATIVE_CONTENT_UNIT_ALREADY_FRONT) {
    return "front";
  }
  if (message === MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK || message === MRAI_NATIVE_CONTENT_UNIT_ALREADY_BACK) {
    return "back";
  }
  return null;
}

/** The pictures among the members on `sourceBytes`: each at the paint drawn
 * at its matrix nearest its predicted order (no matrix: its order). */
async function resolveImageMembers(
  planners: PdfNativeContentReorderChainPlanners,
  sourceBytes: Uint8Array,
  members: readonly PdfNativeContentReorderMember[],
): Promise<PdfNativeContentReorderMember[]> {
  const listed = new Map<number, Promise<readonly (readonly number[] | null)[] | null>>();
  const taken = new Set<string>();
  const out: PdfNativeContentReorderMember[] = [];
  for (const member of members) {
    if (member.kind !== "image" || !member.matrix || !planners.listMraiNativeEmbeddedImagePaintMatrices) {
      out.push(member);
      continue;
    }
    let paints = listed.get(member.pageNumber);
    if (!paints) {
      paints = planners.listMraiNativeEmbeddedImagePaintMatrices({ sourceBytes, pageNumber: member.pageNumber })
        .catch(() => null);
      listed.set(member.pageNumber, paints);
    }
    const matrices = await paints;
    let found = -1;
    for (const [order, matrix] of (matrices ?? []).entries()) {
      if (taken.has(`${member.pageNumber}:${order}`) || !pdfNativeEmbeddedImagePaintMatrixMatches(matrix, member.matrix)) {
        continue;
      }
      if (found < 0 || Math.abs(order - member.imageOrder) < Math.abs(found - member.imageOrder)) found = order;
    }
    if (matrices && found < 0) throw new Error("PDF picture reorder member is stale or ambiguous.");
    const imageOrder = found < 0 ? member.imageOrder : found;
    taken.add(`${member.pageNumber}:${imageOrder}`);
    out.push(imageOrder === member.imageOrder ? member : Object.freeze({ ...member, imageOrder }));
  }
  return out;
}

/**
 * Plan `moves` of `members` -- a page's objects of any kind -- on
 * `sourceBytes` as one revision. Every move is the planner's own proven
 * reorder of one object (a vector's, a picture's, or the layer move of a
 * text object, shading, Form or group) on the bytes the move before left.
 * Refuses (throws the planner's refusal) when a move that may not yield is
 * refused, or when nothing moves at all (every member already at that end:
 * the refusal says which end).
 */
export async function planPdfNativeContentReorderChain(input: {
  readonly engine: PdfNativeChainEngine;
  readonly planners: PdfNativeContentReorderChainPlanners;
  readonly sourceBytes: Uint8Array;
  readonly members: readonly PdfNativeContentReorderMember[];
  /** The page geometry the vectors' planner reads (needed with vectors). */
  readonly pageGeometry: PdfVectorPageGeometry | null;
  readonly moves: readonly PdfNativeVectorReorderChainMove[];
}): Promise<PdfNativeContentReorderChainResult> {
  if (input.members.length === 0 || input.moves.length === 0) {
    throw new Error("A PDF reorder chain needs a member and a move.");
  }
  for (const move of input.moves) {
    if (!Number.isSafeInteger(move.member) || move.member < 0 || move.member >= input.members.length) {
      throw new Error("A PDF reorder chain move names no member.");
    }
  }
  const planners = input.planners;
  // The page reads as the rows the reader showed, or nothing moves; each
  // object the reader named by its place is held to its bytes from here on.
  const listed = new Map<number, Readonly<{ kinds: string; rows: readonly Readonly<{ kind: string; digest: string }>[] }>>();
  let members: PdfNativeContentReorderMember[] = [];
  for (const member of input.members) {
    if (member.kind !== "unit") {
      members.push(member);
      continue;
    }
    if (!planners.planMraiNativeContentUnitReorder) throw new Error("PDF page object reorder is not available here.");
    if (member.kinds === undefined) {
      if (member.unit.digest === undefined) throw new Error("A PDF page object reorder member names no object.");
      members.push(member);
      continue;
    }
    if (!planners.listMraiNativePageContentUnits) throw new Error("PDF page object reorder is not available here.");
    let page = listed.get(member.pageNumber);
    if (!page) {
      page = await planners.listMraiNativePageContentUnits({ sourceBytes: input.sourceBytes,
        pageNumber: member.pageNumber });
      listed.set(member.pageNumber, page);
    }
    if (page.kinds !== member.kinds) throw new Error(MRAI_NATIVE_CONTENT_UNIT_ROWS_DIFFER);
    const row = page.rows.filter((entry) => entry.kind === member.unit.kind)[member.unit.order];
    if (!row || (member.unit.digest !== undefined && member.unit.digest !== row.digest)) {
      throw new Error(MRAI_NATIVE_CONTENT_UNIT_ROWS_DIFFER);
    }
    members.push(Object.freeze({ kind: "unit" as const, pageNumber: member.pageNumber,
      unit: Object.freeze({ kind: member.unit.kind, order: member.unit.order, digest: row.digest }) }));
  }
  members = await resolveImageMembers(planners, input.sourceBytes, members);
  const startMembers = Object.freeze([...members]);
  const vectorEndpoints = () =>
    members.flatMap((member) => (member.kind === "vector" ? [member.endpoint] : []));
  let lastRefusal: unknown = null;
  const chain = await planPdfNativeRevisionChain<PdfNativeContentReorderStep>(
    input.engine,
    input.sourceBytes,
    input.moves.length,
    async (index, sourceBytes) => {
      const move = input.moves[index];
      if (!move) return null;
      members = await resolveImageMembers(planners, sourceBytes, members);
      const member = members[move.member];
      if (!member) throw new Error("A PDF reorder chain move names no member.");
      const yieldAtEnd = (error: unknown): null => {
        const end = arrangedEnd(error);
        if (
          (move.yieldToMembers && end !== null) ||
          (move.direction === "front" && end === "front") ||
          (move.direction === "back" && end === "back")
        ) {
          // To an end: an object already there stays (Office brings a
          // selection whose front-most object is already in front forward
          // around it).
          lastRefusal = error;
          return null;
        }
        throw error;
      };
      if (member.kind === "vector") {
        if (!input.pageGeometry) throw new Error("A PDF vector reorder needs its page geometry.");
        const endpoints = vectorEndpoints();
        const vectorAt = members.slice(0, move.member).filter((entry) => entry.kind === "vector").length;
        const resolved = await resolveMembers(planners, sourceBytes, endpoints);
        const selection = resolved[vectorAt];
        if (!selection) throw new Error("A PDF reorder chain move names no member.");
        let plan: PdfNativeVectorSelectionReorderPlan;
        try {
          plan = await planners.planPdfNativeVectorSelectionReorder({
            sourceBytes,
            payload: Object.freeze({
              schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
              selection,
              pageGeometry: input.pageGeometry,
              direction: move.direction,
            }),
          });
        } catch (error) {
          return yieldAtEnd(error);
        }
        if (move.yieldToMembers && passesMember(plan, selection, resolved.filter((_entry, at) => at !== vectorAt))) {
          plan.prepared.candidateBytes.fill(0);
          return null;
        }
        // The moved member is the plan's own reading of it; the others keep
        // their geometry and shift one place in their family.
        members = members.map((entry, at) =>
          at === move.member
            ? Object.freeze({ kind: "vector" as const, endpoint: plan.afterEndpoint })
            : entry.kind === "vector"
              ? Object.freeze({ kind: "vector" as const,
                  endpoint: shiftedProofIndex(entry.endpoint, plan.beforeEndpoint, plan.afterEndpoint) })
              : entry);
        return Object.freeze({ kind: "vector" as const, plan, prepared: plan.prepared });
      }
      if (member.kind === "image") {
        if (!planners.planPdfNativeEmbeddedImageReorder) throw new Error("PDF picture reorder is not available here.");
        let plan: PdfNativeEmbeddedImageReorderPlanResult;
        try {
          plan = await planners.planPdfNativeEmbeddedImageReorder({
            sourceBytes,
            payload: Object.freeze({
              schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
              selection: Object.freeze({ pageNumber: member.pageNumber, imageOrder: member.imageOrder }),
              direction: move.direction,
            }),
          });
        } catch (error) {
          return yieldAtEnd(error);
        }
        const from = plan.evidence.imageOrder;
        const to = plan.evidence.imageOrderAfter;
        members = members.map((entry, at) =>
          at === move.member
            ? Object.freeze({ ...member, imageOrder: to })
            : entry.kind === "image" && entry.pageNumber === member.pageNumber
              ? Object.freeze({ ...entry, imageOrder: shiftedPlace(entry.imageOrder, from, to) })
              : entry);
        return Object.freeze({ kind: "image" as const, plan, prepared: plan.prepared });
      }
      if (!planners.planMraiNativeContentUnitReorder) throw new Error("PDF page object reorder is not available here.");
      let plan: PdfNativeContentUnitReorderPlan;
      try {
        const digest = member.unit.digest;
        if (digest === undefined) throw new Error("A PDF page object reorder member names no object.");
        plan = await planners.planMraiNativeContentUnitReorder({
          sourceBytes,
          pageNumber: member.pageNumber,
          unit: Object.freeze({ kind: member.unit.kind, order: member.unit.order, digest }),
          direction: move.direction,
        });
      } catch (error) {
        return yieldAtEnd(error);
      }
      members = members.map((entry, at) =>
        at === move.member
          ? Object.freeze({ kind: "unit" as const, pageNumber: member.pageNumber, unit: plan.after })
          : entry.kind === "unit" && entry.pageNumber === member.pageNumber && entry.unit.kind === member.unit.kind
            ? Object.freeze({ ...entry, unit: Object.freeze({ ...entry.unit,
                order: shiftedPlace(entry.unit.order, plan.before.order, plan.after.order) }) })
            : entry);
      return Object.freeze({ kind: "unit" as const, plan, prepared: plan.prepared });
    },
  ).catch((error: unknown) => {
    if (error instanceof Error && error.message === "A PDF edit chain wrote nothing." && lastRefusal) {
      throw lastRefusal;
    }
    throw error;
  });
  members = await resolveImageMembers(planners, chain.candidateBytes, members);
  const vectors = vectorEndpoints();
  const resolvedVectors = vectors.length
    ? (await resolveMembers(planners, chain.candidateBytes, vectors)).map(createPdfNativeVectorHistoryEndpoint)
    : [];
  let vectorAt = 0;
  return Object.freeze({
    steps: chain.steps,
    candidateBytes: chain.candidateBytes,
    candidateSha256: chain.candidateSha256,
    members: Object.freeze(members.map((member) =>
      member.kind === "vector"
        ? Object.freeze({ kind: "vector" as const, endpoint: resolvedVectors[vectorAt++] ?? member.endpoint })
        : member)),
    startMembers,
  });
}

/**
 * Plan `moves` on `sourceBytes` as one revision. `endpoints` name the
 * members; every move is the engine's own proven reorder of one member on
 * the bytes the move before left (planPdfNativeContentReorderChain with
 * vectors only).
 */
export async function planPdfNativeVectorReorderChain(input: {
  readonly engine: PdfNativeChainEngine;
  readonly planners: PdfNativeVectorReorderChainPlanners;
  readonly sourceBytes: Uint8Array;
  readonly endpoints: readonly PdfNativeVectorHistoryEndpoint[];
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly moves: readonly PdfNativeVectorReorderChainMove[];
}): Promise<PdfNativeVectorReorderChainResult> {
  if (input.endpoints.length === 0 || input.moves.length === 0) {
    throw new Error("A PDF vector reorder chain needs a member and a move.");
  }
  const chain = await planPdfNativeContentReorderChain({
    engine: input.engine,
    planners: input.planners,
    sourceBytes: input.sourceBytes,
    members: input.endpoints.map((endpoint) => Object.freeze({ kind: "vector" as const, endpoint })),
    pageGeometry: input.pageGeometry,
    moves: input.moves,
  });
  return Object.freeze({
    steps: Object.freeze(chain.steps.flatMap((step) =>
      step.plan.kind === "vector" ? [Object.freeze({ plan: step.plan.plan, sourceSha256: step.sourceSha256 })] : [])),
    candidateBytes: chain.candidateBytes,
    candidateSha256: chain.candidateSha256,
    memberEndpoints: Object.freeze(chain.members.flatMap((member) => (member.kind === "vector" ? [member.endpoint] : []))),
  });
}
