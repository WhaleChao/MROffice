/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  planSelectionPaneReorder,
  type OfficeStackCommand,
} from "@mrnote/core/web/components/office-ribbon/officeSelectionPaneReorder.ts";
import {
  pdfNativeContentUnitKinds,
  type PdfNativeContentUnitKind,
  type PdfNativeMovableContentUnitKind,
} from "./engine/native/content-unit-model";
import type { PdfNativeVectorReorderChainMove } from "./native-vector-reorder-chain";

export type PdfPageObjectUnitKind = PdfNativeContentUnitKind;

/** One object a layer step passes (the engine reads the same rows off the
 * file: vector-content-transaction contentUnitRows), back first. */
export interface PdfPageObjectRow {
  readonly kind: PdfPageObjectUnitKind;
  /** The operator that starts it in the page's operator list. */
  readonly operatorIndex: number;
  /** Its place among the page's rows of its kind, back first. */
  readonly order: number;
  /** How many objects a group holds (1 for any other row). */
  readonly count: number;
  /** A group's objects, back first. */
  readonly members?: readonly PdfPageObjectRow[];
}

export interface PdfPageObjectRows {
  readonly rows: readonly PdfPageObjectRow[];
  /** The rows as one letter each, which the engine compares before a move. */
  readonly kinds: string;
}

const IMAGE_OPERATORS = [
  "paintImageXObject",
  "paintImageXObjectRepeat",
  "paintImageMaskXObject",
  "paintImageMaskXObjectGroup",
  "paintImageMaskXObjectRepeat",
  "paintJpegXObject",
  "paintSolidColorImageMask",
] as const;

const INLINE_IMAGE_OPERATORS = ["paintInlineImageXObject", "paintInlineImageXObjectGroup"] as const;

const TEXT_SHOW_OPERATORS = [
  "showText",
  "showSpacedText",
  "nextLineShowText",
  "nextLineSetSpacingShowText",
] as const;

const PAINT_OPERATORS = [
  "stroke",
  "closeStroke",
  "fill",
  "eoFill",
  "fillStroke",
  "eoFillStroke",
  "closeFillStroke",
  "closeEOFillStroke",
] as const;

const OPEN_OPERATORS = ["save", "beginMarkedContent", "beginMarkedContentProps"] as const;
const CLOSE_OPERATORS = ["restore", "endMarkedContent"] as const;

type Leaf = Readonly<{ leaf: true; kind: PdfPageObjectUnitKind; operatorIndex: number }>;
type Container = Readonly<{ leaf: false; operatorIndex: number; children: readonly (Leaf | Container)[] }>;
type Unordered = Readonly<{ kind: PdfPageObjectUnitKind; operatorIndex: number; count: number;
  members?: readonly Unordered[] }>;

function childRows(children: readonly (Leaf | Container)[]): Unordered[] {
  const rows: Unordered[] = [];
  for (const child of children) {
    if (child.leaf) {
      rows.push({ kind: child.kind, operatorIndex: child.operatorIndex, count: 1 });
      continue;
    }
    const inner = childRows(child.children);
    const count = inner.reduce((sum, row) => sum + row.count, 0);
    // A group or section holding one object is that object's own wrapper;
    // one holding two or more is passed whole by a layer step.
    if (count < 2) rows.push(...inner);
    else rows.push({ kind: "group", operatorIndex: child.operatorIndex, count, members: inner });
  }
  return rows;
}

function ordered(rows: readonly Unordered[]): PdfPageObjectRow[] {
  const seen = new Map<PdfPageObjectUnitKind, number>();
  return rows.map((row) => {
    const order = seen.get(row.kind) ?? 0;
    seen.set(row.kind, order + 1);
    return Object.freeze({
      kind: row.kind,
      operatorIndex: row.operatorIndex,
      order,
      count: row.count,
      ...(row.members ? { members: Object.freeze(ordered(row.members)) } : {}),
    });
  });
}

/**
 * The page content's rows, back to front (the order they are painted), read
 * from the renderer's operator list: a path, a text object (`BT`…`ET` that
 * shows text), a picture, an inline picture, a shading, a Form drawn whole,
 * or a group -- the outermost `q`…`Q` or marked-content section holding two
 * or more objects, which one layer step passes whole. Annotation appearances
 * (`beginAnnotation`…`endAnnotation`) are not page content.
 */
export function pdfPageObjectRows(
  fnArray: readonly number[],
  argsArray: readonly unknown[] | undefined,
  ops: Readonly<Record<string, number | undefined>>,
): PdfPageObjectRows {
  const idsOf = (names: readonly string[]) =>
    new Set(names.map((name) => ops[name]).filter((id): id is number => typeof id === "number"));
  const images = idsOf(IMAGE_OPERATORS);
  const inlineImages = idsOf(INLINE_IMAGE_OPERATORS);
  const shows = idsOf(TEXT_SHOW_OPERATORS);
  const paints = idsOf(PAINT_OPERATORS);
  const opens = idsOf(OPEN_OPERATORS);
  const closes = idsOf(CLOSE_OPERATORS);
  const stack: { operatorIndex: number; children: (Leaf | Container)[] }[] = [{ operatorIndex: -1, children: [] }];
  const top = () => stack[stack.length - 1]!;
  const leaf = (kind: PdfPageObjectUnitKind, operatorIndex: number) =>
    top().children.push(Object.freeze({ leaf: true as const, kind, operatorIndex }));
  let annotationDepth = 0;
  let formDepth = 0;
  let text: { start: number; shows: boolean } | null = null;
  for (let index = 0; index < fnArray.length; index += 1) {
    const fn = fnArray[index];
    if (fn === undefined) continue;
    if (fn === ops.beginAnnotation) { annotationDepth += 1; continue; }
    if (fn === ops.endAnnotation) { annotationDepth = Math.max(0, annotationDepth - 1); continue; }
    if (annotationDepth > 0) continue;
    if (fn === ops.paintFormXObjectBegin) {
      if (formDepth === 0 && !text) leaf("form", index);
      formDepth += 1;
      continue;
    }
    if (fn === ops.paintFormXObjectEnd) { formDepth = Math.max(0, formDepth - 1); continue; }
    if (formDepth > 0) continue;
    if (text) {
      if (fn === ops.endText) {
        if (text.shows) leaf("text", text.start);
        text = null;
      } else if (shows.has(fn)) {
        text.shows = true;
      }
      continue;
    }
    if (fn === ops.beginText) { text = { start: index, shows: false }; continue; }
    if (opens.has(fn)) { stack.push({ operatorIndex: index, children: [] }); continue; }
    if (closes.has(fn)) {
      if (stack.length < 2) continue;
      const frame = stack.pop()!;
      top().children.push(Object.freeze({ leaf: false as const, operatorIndex: frame.operatorIndex,
        children: frame.children }));
      continue;
    }
    if (fn === ops.constructPath) {
      const args = argsArray?.[index];
      const paint = Array.isArray(args) ? Number(args[0]) : Number.NaN;
      if (paints.has(paint)) leaf("path", index);
      continue;
    }
    if (paints.has(fn)) leaf("path", index);
    else if (images.has(fn)) leaf("image", index);
    else if (inlineImages.has(fn)) leaf("inline-image", index);
    else if (fn === ops.shadingFill) leaf("shading", index);
  }
  while (stack.length > 1) {
    const frame = stack.pop()!;
    top().children.push(...frame.children);
  }
  const rows = Object.freeze(ordered(childRows(stack[0]!.children)));
  return Object.freeze({ rows, kinds: pdfNativeContentUnitKinds(rows) });
}

/** What moves when a pane row moves: a source vector the reader routes, a
 * picture the reader targets (both by their renderer operator index), or a
 * text object, shading, Form or group by its place among the page's rows of
 * its kind. */
export type PdfSelectionPaneMember =
  | Readonly<{ kind: "vector"; operatorIndex: number }>
  | Readonly<{ kind: "image"; operatorIndex: number }>
  | Readonly<{ kind: "unit"; unitKind: PdfNativeMovableContentUnitKind; order: number }>;

export interface PdfSelectionPaneContentRow {
  readonly key: string;
  readonly kind: PdfPageObjectUnitKind;
  readonly operatorIndex: number;
  /** How many layer steps passing the row takes (text objects listed as
   * one row are passed one by one; a group is passed whole). */
  readonly weight: number;
  /** What moves when the row moves, back first; empty: it does not move
   * from here (an inline picture, a group's member, a path the reader does
   * not route). */
  readonly members: readonly PdfSelectionPaneMember[];
  /** A group's objects, front first (shown when it is expanded). */
  readonly children?: readonly PdfSelectionPaneContentRow[];
  /** A source vector the reader routes: its renderer operator index. */
  readonly vectorOperatorIndex?: number;
}

/**
 * The pane's page-content rows, front first. A path moves when the reader
 * routes it, a picture when the reader targets it; text objects, shadings,
 * Forms and groups always (the engine holds each to the rows it reads).
 * Text objects painted one after another are one row.
 */
export function pdfSelectionPaneContentRows(
  rows: readonly PdfPageObjectRow[],
  movable: Readonly<{ vectors: ReadonlySet<number>; images: ReadonlySet<number> }>,
  /** Rows inside a group: listed, not moved on their own. */
  inGroup = false,
): readonly PdfSelectionPaneContentRow[] {
  const out: PdfSelectionPaneContentRow[] = [];
  for (const row of rows) {
    const previous = out[out.length - 1];
    if (row.kind === "text" && previous?.kind === "text") {
      out[out.length - 1] = Object.freeze({
        ...previous,
        weight: previous.weight + 1,
        members: inGroup ? previous.members
          : Object.freeze([...previous.members, Object.freeze({ kind: "unit" as const, unitKind: "text" as const,
            order: row.order })]),
      });
      continue;
    }
    const members: PdfSelectionPaneMember[] = [];
    if (!inGroup) {
      if (row.kind === "path" && movable.vectors.has(row.operatorIndex)) {
        members.push({ kind: "vector", operatorIndex: row.operatorIndex });
      } else if (row.kind === "image" && movable.images.has(row.operatorIndex)) {
        members.push({ kind: "image", operatorIndex: row.operatorIndex });
      } else if (row.kind === "text" || row.kind === "shading" || row.kind === "form" || row.kind === "group") {
        members.push({ kind: "unit", unitKind: row.kind, order: row.order });
      }
    }
    out.push(Object.freeze({
      key: `${row.kind}:${row.operatorIndex}`,
      kind: row.kind,
      operatorIndex: row.operatorIndex,
      weight: 1,
      members: Object.freeze(members),
      ...(row.members ? { children: pdfSelectionPaneContentRows(row.members, movable, true) } : {}),
      ...(row.kind === "path" && movable.vectors.has(row.operatorIndex)
        ? { vectorOperatorIndex: row.operatorIndex }
        : {}),
    }));
  }
  return Object.freeze(out.reverse());
}

interface PaneModel {
  readonly order: readonly string[];
  readonly moves: readonly Readonly<{ key: string; direction: OfficeStackCommand }>[];
}

/**
 * One Arrange command on the pane's model: the moving rows as Office moves a
 * selection (all to an end keeping their order; or each one place, never
 * past one another), and the single-object layer moves the writer makes for
 * it -- one per object a step passes. Null when nothing moves.
 */
function paneStep(
  model: PaneModel,
  moving: ReadonlySet<string>,
  weightOf: (key: string) => number,
  command: OfficeStackCommand,
): PaneModel | null {
  const order = [...model.order];
  const moves = [...model.moves];
  const picked = order.filter((key) => moving.has(key));
  if (picked.length === 0) return null;
  if (command === "front" || command === "back") {
    const rest = order.filter((key) => !moving.has(key));
    const next = command === "front" ? [...picked, ...rest] : [...rest, ...picked];
    if (next.every((key, index) => key === order[index])) return null;
    // To the front from the back-most selected up; to the back from the
    // front-most down: each lands next to the one moved before it.
    const sequence = command === "front" ? [...picked].reverse() : picked;
    for (const key of sequence) moves.push({ key, direction: command });
    return { order: next, moves };
  }
  let moved = false;
  if (command === "forward") {
    for (let index = 1; index < order.length; index += 1) {
      const key = order[index]!;
      const ahead = order[index - 1]!;
      if (!moving.has(key) || moving.has(ahead)) continue;
      order[index - 1] = key;
      order[index] = ahead;
      for (let step = 0; step < weightOf(ahead); step += 1) moves.push({ key, direction: "forward" });
      moved = true;
    }
  } else {
    for (let index = order.length - 2; index >= 0; index -= 1) {
      const key = order[index]!;
      const behind = order[index + 1]!;
      if (!moving.has(key) || moving.has(behind)) continue;
      order[index + 1] = key;
      order[index] = behind;
      for (let step = 0; step < weightOf(behind); step += 1) moves.push({ key, direction: "backward" });
      moved = true;
    }
  }
  return moved ? { order, moves } : null;
}

export interface PdfSelectionPaneMoves {
  /** What moves, back to front as the page paints it now: the chain's
   * members. */
  readonly members: readonly PdfSelectionPaneMember[];
  readonly moves: readonly PdfNativeVectorReorderChainMove[];
}

/**
 * What a Selection pane drop (the moved rows to `slot` of the page content,
 * counting the rows that stay in front) or a Ctrl+Shift+F / Ctrl+Shift+B
 * step asks of the writer: every object of the moved rows, back to front,
 * and their single-object layer moves in order -- each row's objects moved
 * one after another so they never pass one another. Only rows with members
 * move; null when nothing would.
 */
export function pdfSelectionPaneMoves(input: {
  readonly rows: readonly PdfSelectionPaneContentRow[];
  readonly keys: readonly string[];
  readonly slot?: number;
  readonly step?: "forward" | "backward";
}): PdfSelectionPaneMoves | null {
  const byKey = new Map(input.rows.map((row) => [row.key, row]));
  const moving = new Set(input.keys.filter((key) => (byKey.get(key)?.members.length ?? 0) > 0));
  if (moving.size === 0) return null;
  const weightOf = (key: string) => byKey.get(key)?.weight ?? 1;
  const start: PaneModel = { order: input.rows.map((row) => row.key), moves: [] };
  const result = input.step
    ? paneStep(start, moving, weightOf, input.step)
    : planSelectionPaneReorder<PaneModel>({
        model: start,
        order: (model) => model.order,
        moving: [...moving],
        slot: input.slot ?? 0,
        step: (model, command) => paneStep(model, moving, weightOf, command),
      });
  if (!result || result.moves.length === 0) return null;
  // Members back to front, as the page paints them now.
  const keysBackToFront = start.order.filter((key) => moving.has(key)).reverse();
  const members: PdfSelectionPaneMember[] = [];
  const memberAt = new Map<string, number[]>();
  for (const key of keysBackToFront) {
    const places: number[] = [];
    for (const member of byKey.get(key)!.members) {
      places.push(members.length);
      members.push(member);
    }
    memberAt.set(key, places);
  }
  const moves: PdfNativeVectorReorderChainMove[] = [];
  for (const move of result.moves) {
    const places = memberAt.get(move.key);
    if (!places || places.length === 0) return null;
    // A row of several objects: the one the move leads with first, so they
    // never pass one another (forward and back: the front-most; backward
    // and front: the back-most).
    const sequence = move.direction === "forward" || move.direction === "back" ? [...places].reverse() : places;
    for (const member of sequence) moves.push(Object.freeze({ member, direction: move.direction }));
  }
  return Object.freeze({ members: Object.freeze(members), moves: Object.freeze(moves) });
}
