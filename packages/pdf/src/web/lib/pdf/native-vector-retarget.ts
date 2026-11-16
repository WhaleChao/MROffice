/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A queued edit of one of the page's own shapes, stood on other bytes than
 * it was made on: the version another window saved first, or the same
 * content the server rewrote as another file.
 *
 * A shape's command names its object by where its operators lie in the file
 * it was made on (a proof id, a digest, a span of the content stream), and
 * the planner holds it to every one of them: on any other bytes the command
 * is refused, and the queue let the edit go. Here the object is found again
 * on those bytes by what the command already says of it -- its kind, its
 * paint, its place in the page's drawing and the geometry it had -- and the
 * command is written for it there. Nothing is read from history: the object
 * is taken from the shapes the planner itself reads off the new bytes
 * (`readPdfNativeVectorPageShapes`), so the planner proves it as it proves
 * any selection.
 *
 * The order of trust: the object in its own slot of the stream, as it was;
 * else the one object of the page that is what it was, when the page never
 * drew a twin of it (the other window changed something drawn before it);
 * else, when the other window changed this very object, the later save
 * stands -- and only when it is certain which object that is
 * (`resolvePdfNativeVectorLaterSave`). Anything less is left to the planner,
 * which turns the edit away: an edit made on one object is never made on
 * another, however alike the two are drawn.
 */
import {
  MRAI_NATIVE_VECTOR_DELETE_ACTION,
  MRAI_NATIVE_VECTOR_DUPLICATE_ACTION,
  MRAI_NATIVE_VECTOR_REORDER_ACTION,
  MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION,
  createPdfNativeVectorHistoryEndpoint,
  exactSelection,
  samePdfNativeVectorHistoryEndpoint,
} from "./native-vector-commands";
import { MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION } from "./native-vector-style-commands";
import type { PdfNativeVectorAffineTransform } from "./engine/native/vector-content-transaction";
import type { PdfNativeVectorHistoryEndpoint, PdfNativeVectorWriteSelection } from "./native-vector-selection-edit";

export type PdfNativeVectorCommandKind = "geometry" | "delete" | "duplicate" | "reorder" | "style";

/** What a queued shape command changes, as the command (or its recipe) says
 * it stood when the edit was made. */
export type PdfNativeVectorCommandTarget = Readonly<{
  kind: PdfNativeVectorCommandKind;
  pageNumber: number;
  /** The bytes the command's own proof names. */
  sourceSha256: string;
  /** The object as the edit was made on it. */
  made: PdfNativeVectorWriteSelection;
  /** The edit leaves the object where and as it was drawn (a colour, a
   * place in the stack, a copy made of it): it reads the same after it. */
  keepsGeometry: boolean;
  /** The edit says where the object ends, whatever it was before (a delete,
   * a merged outline, a style that moves nothing), or can be said so (a
   * plain move, turn or resize): the later save can stand on an object the
   * other window changed. */
  laterSave: "as-it-is" | "onto-its-result" | null;
}>;

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null;
}

function selectionOf(value: unknown): PdfNativeVectorWriteSelection | null {
  try { return exactSelection(value); } catch { return null; }
}

/**
 * The target of a shape command. A style command's route holds no geometry,
 * so its object comes from the recipe (`recipeTarget`, kept when the edit
 * was staged); the others carry their selection themselves.
 */
export function readPdfNativeVectorCommandTarget(
  command: unknown,
  recipeTarget?: unknown,
): PdfNativeVectorCommandTarget | null {
  const action = record(command)?.action;
  const payload = record(record(command)?.payload);
  if (!payload) return null;
  if (action === MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION) {
    const route = record(payload.route);
    const changes = record(payload.changes);
    const made = selectionOf(record(record(recipeTarget)?.vector)?.selection);
    if (!route || !changes || !made || typeof route.sourceSha256 !== "string" ||
      // The recipe's object is the one the route names, in the same bytes.
      made.sourceSha256 !== route.sourceSha256 || made.proofId !== route.proofId ||
      made.proofDigest !== route.proofDigest || made.pageNumber !== route.pageNumber) return null;
    const moves = changes.translate !== undefined || changes.transform !== undefined;
    return Object.freeze({ kind: "style" as const, pageNumber: made.pageNumber, sourceSha256: made.sourceSha256, made,
      keepsGeometry: !moves, laterSave: moves ? null : "as-it-is" as const });
  }
  const kind = action === MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION ? "geometry" as const
    : action === MRAI_NATIVE_VECTOR_DELETE_ACTION ? "delete" as const
    : action === MRAI_NATIVE_VECTOR_DUPLICATE_ACTION ? "duplicate" as const
    : action === MRAI_NATIVE_VECTOR_REORDER_ACTION ? "reorder" as const : null;
  if (!kind) return null;
  const made = selectionOf(payload.selection);
  if (!made) return null;
  // A step of the shape's own history (a turn taken back, a copy taken away,
  // a place in the stack put back) names bytes of the file it was made on.
  const history = payload.rotationHistory !== undefined || payload.direction === "remove" || payload.direction === "restore";
  const laterSave = kind === "delete" ? "as-it-is" as const
    : kind !== "geometry" || history || payload.vertexEdit !== undefined ? null
    : payload.outlineReplace !== undefined ? "as-it-is" as const : "onto-its-result" as const;
  return Object.freeze({ kind, pageNumber: made.pageNumber, sourceSha256: made.sourceSha256, made,
    keepsGeometry: kind === "duplicate" || kind === "reorder", laterSave });
}

/** The object `made` names, read as `candidate` on other bytes: every field
 * of a history endpoint (kind, paint, page, bounds, geometry), and -- `there`
 * -- its place in the page's drawing and the stream that draws it. */
function standsAs(candidate: PdfNativeVectorWriteSelection, made: PdfNativeVectorHistoryEndpoint,
  where: "there" | "anywhere"): boolean {
  const read = createPdfNativeVectorHistoryEndpoint(candidate);
  if (where === "there") return samePdfNativeVectorHistoryEndpoint(read, made);
  const { formInstance: _formInstance, ...rest } = made as PdfNativeVectorHistoryEndpoint & { formInstance?: string };
  return samePdfNativeVectorHistoryEndpoint(read, { ...rest, pageProofIndex: read.pageProofIndex,
    contentStreamObjectNumber: read.contentStreamObjectNumber, contentStreamGeneration: read.contentStreamGeneration,
    ...(read.formInstance !== undefined ? { formInstance: read.formInstance } : {}) } as PdfNativeVectorHistoryEndpoint);
}

/**
 * The object an edit was made on, among the shapes of the same page read off
 * other bytes: the one shape of the page that is what it was (its kind, its
 * paint, its geometry), when it is certain that shape is the object and not
 * its twin.
 *
 * In its own slot -- the same stream, the same offset into it -- it is the
 * object: nothing before it in the stream changed, and it reads as it did.
 * Elsewhere (the other window changed something drawn before it, so its
 * operators moved along the stream or down the drawing order) it is the
 * object only if the page never drew a second shape like it at that place:
 * a page that draws two (an outline over its own fill's outline, a shape
 * pasted onto itself) and whose other window moved or deleted the edited
 * one still shows one shape there -- the other one. Whether it ever drew a
 * twin is read from the page as this edit left it when it was first planned
 * (`planned.shapes`; `planned.keepsGeometry`: the edit left the object where
 * it was, so that page shows it there too). Without that page nothing proves
 * it, and only the slot is taken.
 *
 * Null when no single shape is the object for certain.
 */
export function resolvePdfNativeVectorSelectionOn(
  shapes: readonly PdfNativeVectorWriteSelection[],
  made: PdfNativeVectorWriteSelection,
  planned?: Readonly<{ shapes: readonly PdfNativeVectorWriteSelection[]; keepsGeometry: boolean }> | null,
): Readonly<{ selection: PdfNativeVectorWriteSelection; relocated: boolean }> | null {
  const wanted = createPdfNativeVectorHistoryEndpoint(made);
  const like = shapes.filter((shape) => standsAs(shape, wanted, "anywhere"));
  const inSlot = like.filter((shape) => inSlotOf(shape, made));
  if (inSlot.length === 1) return Object.freeze({ selection: inSlot[0]!, relocated: false });
  if (like.length !== 1 || !planned) return null;
  const twins = planned.shapes.filter((shape) => standsAs(shape, wanted, "anywhere")).length - (planned.keepsGeometry ? 1 : 0);
  return twins === 0 ? Object.freeze({ selection: like[0]!, relocated: true }) : null;
}

/** The slot `made` was drawn in: the same stream, the same offset into it,
 * the same family and kind of object. An object rewritten where it stood
 * keeps it; one drawn anew (after a delete, by an append) never has it. */
function inSlotOf(shape: PdfNativeVectorWriteSelection, made: PdfNativeVectorWriteSelection): boolean {
  return shape.pageNumber === made.pageNumber && shape.proofFamily === made.proofFamily &&
    shape.nativeKind === made.nativeKind &&
    shape.contentStreamObjectNumber === made.contentStreamObjectNumber &&
    shape.contentStreamGeneration === made.contentStreamGeneration &&
    shape.formInstance === made.formInstance && shape.operatorStart === made.operatorStart;
}

const same = (left: PdfNativeVectorWriteSelection, right: PdfNativeVectorWriteSelection, where: "there" | "anywhere") =>
  standsAs(left, createPdfNativeVectorHistoryEndpoint(right), where);

/**
 * The object an edit was made on, when the other window changed that very
 * object: `onHead` are the page's shapes on the version it saved, `onPlanned`
 * the same page as this edit left it when it was first planned. It is the
 * object in the edit's own slot when every other shape of the page is where
 * this edit left it -- then nothing but that object differs between the two
 * windows' pages, and the later save stands on it. `already-there`: the page
 * holds this edit's result already (the other window made the same change).
 * Null whenever that cannot be told for certain.
 */
export function resolvePdfNativeVectorLaterSave(input: Readonly<{
  target: PdfNativeVectorCommandTarget;
  onHead: readonly PdfNativeVectorWriteSelection[];
  onPlanned: readonly PdfNativeVectorWriteSelection[];
}>): Readonly<{ selection: PdfNativeVectorWriteSelection; result: PdfNativeVectorWriteSelection | null }> |
  "already-there" | null {
  const { target, onHead, onPlanned } = input;
  if (!target.laterSave) return null;
  if (target.kind === "delete") {
    // This edit took the object away: the other shapes follow one another
    // as they did, one family index down past it.
    if (onHead.length === onPlanned.length) {
      return onHead.every((shape, index) => same(shape, onPlanned[index]!, "anywhere")) &&
        !onHead.some((shape) => inSlotOf(shape, target.made)) ? "already-there" : null;
    }
    if (onHead.length !== onPlanned.length + 1) return null;
    const slot = onHead.findIndex((shape) => inSlotOf(shape, target.made));
    if (slot < 0 || onHead.some((shape, index) => index !== slot && inSlotOf(shape, target.made))) return null;
    const others = onHead.filter((_, index) => index !== slot);
    return others.every((shape, index) => same(shape, onPlanned[index]!, "anywhere"))
      ? Object.freeze({ selection: onHead[slot]!, result: null }) : null;
  }
  if (onHead.length !== onPlanned.length) return null;
  const differing = onHead.flatMap((shape, index) => same(shape, onPlanned[index]!, "there") ? [] : [index]);
  if (differing.length === 0) return target.keepsGeometry ? null : "already-there";
  if (differing.length !== 1) return null;
  const selection = onHead[differing[0]!]!, result = onPlanned[differing[0]!]!;
  if (!inSlotOf(selection, target.made)) return null;
  return Object.freeze({ selection, result });
}

type Point = Readonly<{ x: number; y: number }>;

/** The points a shape's outline is drawn through, in drawing order. */
function outlinePoints(selection: PdfNativeVectorWriteSelection): readonly Point[] {
  if (selection.proofFamily === "subpath") {
    return selection.commands.flatMap((command): Point[] => command.operator === "h" ? []
      : "point" in command ? [command.point] : [command.control1, command.control2, command.end]);
  }
  const geometry = selection.geometry;
  if (geometry.kind === "line") return [geometry.start, geometry.end];
  if (geometry.kind === "rectangle") return [...geometry.corners];
  return [geometry.center,
    { x: geometry.center.x + geometry.axisX.x, y: geometry.center.y + geometry.axisX.y },
    { x: geometry.center.x + geometry.axisY.x, y: geometry.center.y + geometry.axisY.y }];
}

/** How far (points) a mapped outline point may land from where it is asked. */
const ONTO_TOLERANCE = 0.001;

/**
 * The page transform that takes `from` onto `onto`: the two are one object
 * drawn twice over (the same family, kind and outline operators), so the
 * transform is read from their points and holds for every one of them. A
 * straight outline gives no third point; it is taken along itself (a turn
 * and a stretch of its own length). Null when the outlines are not one
 * shape under one transform, or `from` already lies on `onto`.
 */
export function pdfNativeVectorAffineOnto(from: PdfNativeVectorWriteSelection,
  onto: PdfNativeVectorWriteSelection): PdfNativeVectorAffineTransform | null {
  if (from.proofFamily !== onto.proofFamily || from.nativeKind !== onto.nativeKind) return null;
  if (from.proofFamily === "subpath" && onto.proofFamily === "subpath" &&
    (from.closed !== onto.closed || from.commands.length !== onto.commands.length ||
      from.commands.some((command, index) => command.operator !== onto.commands[index]!.operator))) return null;
  const source = outlinePoints(from), wanted = outlinePoints(onto);
  if (source.length !== wanted.length || source.length < 2) return null;
  const origin = source[0]!, target = wanted[0]!;
  // The point farthest from the first, then the one farthest off their line.
  const far = source.reduce((best, point, index) =>
    Math.hypot(point.x - origin.x, point.y - origin.y) > Math.hypot(source[best]!.x - origin.x, source[best]!.y - origin.y)
      ? index : best, 0);
  const ux = source[far]!.x - origin.x, uy = source[far]!.y - origin.y;
  const length = Math.hypot(ux, uy);
  if (!(length > ONTO_TOLERANCE)) return null;
  const off = (point: Point) => Math.abs((point.x - origin.x) * uy - (point.y - origin.y) * ux) / length;
  const third = source.reduce((best, point, index) => off(point) > off(source[best]!) ? index : best, 0);
  const px = wanted[far]!.x - target.x, py = wanted[far]!.y - target.y;
  let a: number, b: number, c: number, d: number;
  if (off(source[third]!) > ONTO_TOLERANCE) {
    const vx = source[third]!.x - origin.x, vy = source[third]!.y - origin.y;
    const qx = wanted[third]!.x - target.x, qy = wanted[third]!.y - target.y;
    const determinant = ux * vy - uy * vx;
    a = (px * vy - qx * uy) / determinant; c = (qx * ux - px * vx) / determinant;
    b = (py * vy - qy * uy) / determinant; d = (qy * ux - py * vx) / determinant;
  } else {
    // Along itself: u -> p as a turn and a stretch.
    const scale = length * length;
    a = (px * ux + py * uy) / scale; b = (py * ux - px * uy) / scale;
    c = -b; d = a;
  }
  const e = target.x - a * origin.x - c * origin.y, f = target.y - b * origin.x - d * origin.y;
  if (![a, b, c, d, e, f].every(Number.isFinite) || Math.abs(a * d - b * c) <= 1e-9) return null;
  const lands = source.every((point, index) =>
    Math.abs(a * point.x + c * point.y + e - wanted[index]!.x) <= ONTO_TOLERANCE &&
    Math.abs(b * point.x + d * point.y + f - wanted[index]!.y) <= ONTO_TOLERANCE);
  if (!lands) return null;
  if ([a - 1, b, c, d - 1, e, f].every((value) => Math.abs(value) <= 0.000_01)) return null;
  return Object.freeze({ a, b, c, d, e, f });
}

export type PdfNativeVectorStanding =
  | Readonly<{ status: "stood"; how: "in-place" | "relocated" | "later-save";
      /** The object on the new bytes (a style command's route is resolved for it). */
      selection: PdfNativeVectorWriteSelection;
      /** The command's payload written for it; a style command's still
       * carries its old route (`selection` names the object to resolve). */
      payload: Record<string, unknown> }>
  | Readonly<{ status: "already-there" }>
  | Readonly<{ status: "not-found" }>;

/**
 * A shape command's payload written for other bytes: its object found again
 * among `onHead` (the page's shapes on those bytes), and -- when the other
 * window changed that very object and `onPlanned` (the same page as this
 * edit first left it) proves which one it is -- the edit made on the object
 * as that window left it, so that the later save stands.
 *
 * `plannedByGroup`: `onPlanned` is the page as the whole group this edit is
 * a member of left it (its members are planned one on another's bytes, and
 * only the group's last page is kept). It still shows whether the page drew
 * a twin of the object; it is not this edit's own page, so no later save is
 * read from it.
 */
export function standPdfNativeVectorCommandOn(input: Readonly<{
  target: PdfNativeVectorCommandTarget;
  payload: Record<string, unknown>;
  onHead: readonly PdfNativeVectorWriteSelection[];
  onPlanned?: readonly PdfNativeVectorWriteSelection[] | null;
  plannedByGroup?: boolean;
}>): PdfNativeVectorStanding {
  const { target, payload, onHead } = input;
  const planned = input.onPlanned ? { shapes: input.onPlanned, keepsGeometry: target.keepsGeometry } : null;
  const written = (selection: PdfNativeVectorWriteSelection, more: Record<string, unknown> = {}) => {
    if (target.kind === "style") return { ...payload, ...more };
    // A place in the stack put back names the neighbour it was swapped with.
    const neighbour = payload.neighbour === undefined || payload.neighbour === null ? null : selectionOf(payload.neighbour);
    // (It stands where it stood on the page this edit left.)
    const neighbourOn = neighbour ? resolvePdfNativeVectorSelectionOn(onHead, neighbour,
      input.onPlanned ? { shapes: input.onPlanned, keepsGeometry: true } : null) : null;
    if (neighbour && !neighbourOn) return null;
    return { ...payload, selection, ...(neighbourOn ? { neighbour: neighbourOn.selection } : {}), ...more };
  };
  const found = resolvePdfNativeVectorSelectionOn(onHead, target.made, planned);
  if (found) {
    const stood = written(found.selection);
    return stood ? Object.freeze({ status: "stood" as const, how: found.relocated ? "relocated" as const : "in-place" as const,
      selection: found.selection, payload: stood }) : Object.freeze({ status: "not-found" as const });
  }
  const later = input.onPlanned && !input.plannedByGroup
    ? resolvePdfNativeVectorLaterSave({ target, onHead, onPlanned: input.onPlanned }) : null;
  if (later === "already-there") return Object.freeze({ status: "already-there" as const });
  if (!later) return Object.freeze({ status: "not-found" as const });
  if (target.laterSave === "onto-its-result") {
    // Where this edit put the object, from where the other window left it.
    const affine = later.result ? pdfNativeVectorAffineOnto(later.selection, later.result) : null;
    if (!affine) return Object.freeze({ status: "not-found" as const });
    const stood = written(later.selection, { affine });
    return stood ? Object.freeze({ status: "stood" as const, how: "later-save" as const, selection: later.selection,
      payload: stood }) : Object.freeze({ status: "not-found" as const });
  }
  const stood = written(later.selection);
  return stood ? Object.freeze({ status: "stood" as const, how: "later-save" as const, selection: later.selection,
    payload: stood }) : Object.freeze({ status: "not-found" as const });
}
