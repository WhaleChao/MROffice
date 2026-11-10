/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
interface PaintedBlock {
  readonly zIndex?: unknown;
  readonly sourceOrder?: unknown;
  readonly smartArtRole?: unknown;
  readonly groupId?: unknown;
  readonly groupPath?: unknown;
  readonly source?: unknown;
}

export interface PresentationPaintOrderOptions {
  /** The slide records the author's reorders of its imported objects
   * (`presentationSlideReordersNativeObjects`). */
  readonly reordered?: boolean;
}

/** Whether a slide records reorders of its imported objects: the list, not
 * the places they state, then orders them. */
export function presentationSlideReordersNativeObjects(slide: unknown): boolean {
  return (
    !!slide &&
    typeof slide === "object" &&
    (slide as { nativeShapeOrderTarget?: unknown }).nativeShapeOrderTarget !== undefined
  );
}

/** What a slide says of its Header & Footer fields' places among its shapes
 * (`slide.headerFooterPaintOrders`, read on import), lowest first. Empty when
 * it says nothing or its objects have been reordered since. */
function footerFieldOrders(slide: unknown): ReadonlyArray<readonly [string, number]> {
  if (!slide || typeof slide !== "object" || presentationSlideReordersNativeObjects(slide)) return [];
  const orders = (slide as { headerFooterPaintOrders?: unknown }).headerFooterPaintOrders;
  if (!orders || typeof orders !== "object" || Array.isArray(orders)) return [];
  return Object.entries(orders as Record<string, unknown>)
    .filter((entry): entry is [string, number] => typeof entry[1] === "number" && Number.isFinite(entry[1]))
    .sort((a, b) => a[1] - b[1]);
}

/**
 * The stacking the Header & Footer fields' places come to. The fields are drawn
 * over the slide (`PresentationHeaderFooterOverlay`), and PowerPoint paints a
 * footer placeholder where it sits in the shape tree, so a shape after it covers
 * it. Every block stands at `z-index: 0`, the fields at 2: a field that has
 * something after it stands at 2 + 2n (n: the fields below it), a block over n
 * fields at 1 + 2n, and a field with nothing after it over all of them. Undefined
 * where a slide says nothing, so such a slide is drawn as ever.
 */
export function presentationFooterFieldZIndex(slide: unknown, kind: string): number | undefined {
  const orders = footerFieldOrders(slide);
  if (orders.length === 0) return undefined;
  const rank = orders.findIndex(([field]) => field === kind);
  return 2 + 2 * (rank < 0 ? orders.length : rank);
}

/** See `presentationFooterFieldZIndex`. Undefined for what is under every field
 * and for what is not the slide's own (its master's and layout's shapes). A block
 * with no place of its own was added here, after what the file holds. */
export function presentationBlockOverFooterFieldsZIndex(
  slide: unknown,
  block: { readonly id?: unknown; readonly sourceOrder?: unknown },
): number | undefined {
  const orders = footerFieldOrders(slide);
  if (orders.length === 0) return undefined;
  const own = (slide as { blocks?: unknown }).blocks;
  if (!Array.isArray(own) || !own.some((candidate) => (candidate as { id?: unknown } | null)?.id === block.id)) {
    return undefined;
  }
  const place = finite(block.sourceOrder);
  const below = place === null ? orders.length : orders.filter(([, order]) => order < place).length;
  return below === 0 ? undefined : 1 + 2 * below;
}

/** A block's frame at its place among the slide's Header & Footer fields. */
export function withFooterFieldStacking<S extends { zIndex?: unknown }>(
  style: S | undefined,
  slide: unknown,
  block: { readonly id?: unknown; readonly sourceOrder?: unknown },
): S | undefined {
  if (!style) return style;
  const zIndex = presentationBlockOverFooterFieldsZIndex(slide, block);
  return zIndex === undefined ? style : { ...style, zIndex };
}

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/** The diagram an authored SmartArt member belongs to (its innermost group:
 * a diagram the author grouped with other objects sits inside that group),
 * or null for any other object. An imported member states its file's place
 * and ranks as any imported object does. */
function authoredDiagramKey(block: PaintedBlock): string | null {
  if (typeof block.smartArtRole !== "string" || block.smartArtRole === "") return null;
  if (finite(block.sourceOrder) !== null) return null;
  const path = Array.isArray(block.groupPath)
    ? block.groupPath.filter((entry): entry is string => typeof entry === "string" && entry.trim() !== "")
    : [];
  const inner = path.at(-1) ?? (typeof block.groupId === "string" ? block.groupId.trim() : "");
  return inner === "" ? null : inner;
}

/** The place an object states for itself (a whole place, never below the
 * slide's first), or null when its place in the list is all it has. */
function statedPlace(block: PaintedBlock): number | null {
  const place = finite(block.zIndex) ?? finite(block.sourceOrder);
  return place === null ? null : Math.max(0, Math.round(place));
}

/** An object of the slide's own file tree: it names its slide, and it is not
 * a layout's or master's graphic the slide shows. */
function ownNativeObject(block: PaintedBlock): boolean {
  const source = block.source;
  if (!source || typeof source !== "object") return false;
  const record = source as { slideId?: unknown; inheritedFrom?: unknown };
  return (
    typeof record.slideId === "string" &&
    record.slideId.trim() !== "" &&
    record.inheritedFrom === undefined
  );
}

/** An object an imported file placed: one of the slide's own shapes (by its
 * slide, or its slide part for an XML-only chart) or a layout's or master's
 * graphic the slide shows. */
function fileCustody(block: PaintedBlock): boolean {
  const source = block.source;
  if (!source || typeof source !== "object") return false;
  const record = source as { slideId?: unknown; slidePart?: unknown; inheritedFrom?: unknown };
  return (
    (typeof record.slideId === "string" && record.slideId.trim() !== "") ||
    typeof record.slidePart === "string" ||
    record.inheritedFrom === "layout" ||
    record.inheritedFrom === "master"
  );
}











export function presentationRestackedPlaces<T extends object>(blocks: readonly T[]): T[] {
  const list = blocks as readonly PaintedBlock[];
  const dealt = list
    .map((block, index) => index)
    .filter(
      (index) =>
        statedPlace(list[index]) !== null &&
        !fileCustody(list[index]) &&
        typeof list[index].smartArtRole !== "string",
    );
  const places = dealt.map((index) => statedPlace(list[index]) as number).sort((a, b) => a - b);
  const next = [...blocks];
  dealt.forEach((index, rank) => {
    const block = list[index];
    const place = places[rank];
    if (statedPlace(block) === place) return;
    const restated: Record<string, unknown> = { ...(block as Record<string, unknown>) };
    if (finite(block.zIndex) !== null) restated.zIndex = place;
    if (finite(block.sourceOrder) !== null) restated.sourceOrder = place;
    next[index] = restated as T;
  });
  return next;
}

/** The places of the list's objects; on a reordered slide the slide's own
 * imported objects take their places back in the list's order. */
function statedPlaces(blocks: readonly PaintedBlock[], reordered: boolean): Array<number | null> {
  const places = blocks.map(statedPlace);
  if (!reordered) return places;
  const dealt = blocks
    .map((block, index) => index)
    .filter((index) => places[index] !== null && ownNativeObject(blocks[index]));
  const sorted = dealt.map((index) => places[index] as number).sort((a, b) => a - b);
  dealt.forEach((index, rank) => {
    places[index] = sorted[rank];
  });
  return places;
}

/**
 * The list's indices in paint order, back to front. An object that states no
 * place of its own paints in front of every object before it in the list (so
 * a new object, or one brought to the front, is in front of an imported
 * slide's objects too); ties keep the list's order.
 */
export function presentationPaintOrder(
  list: readonly object[],
  options: PresentationPaintOrderOptions = {},
): number[] {
  const blocks = list as readonly PaintedBlock[];
  const places = statedPlaces(blocks, options.reordered === true);
  const units: Array<{ place: number; members: number[] }> = [];
  let front = -Infinity;
  let index = 0;
  while (index < blocks.length) {
    const key = authoredDiagramKey(blocks[index]);
    let end = index + 1;
    if (key !== null) {
      while (end < blocks.length && authoredDiagramKey(blocks[end]) === key) end += 1;
    }
    const members = Array.from({ length: end - index }, (_, offset) => index + offset);
    if (key !== null) {
      // Inside the diagram, its own layers; the sort is stable.
      members.sort((a, b) => (finite(blocks[a].zIndex) ?? 0) - (finite(blocks[b].zIndex) ?? 0));
    }
    const place = key === null ? (places[index] ?? front) : front;
    units.push({ place, members });
    front = Math.max(front, place);
    index = end;
  }
  // Array.prototype.sort is stable, so equal places keep the list's order.
  return units
    .map((unit, order) => ({ ...unit, order }))
    .sort((a, b) => (a.place === b.place ? a.order - b.order : a.place < b.place ? -1 : 1))
    .flatMap((unit) => unit.members);
}

/**
 * The place the first of `blocks` to state one paints at, or null when none
 * does (an imported SmartArt diagram's members all state their frame's).
 */
export function presentationFirstStatedPlace(blocks: readonly object[]): number | null {
  for (const block of blocks as readonly PaintedBlock[]) {
    const place = statedPlace(block);
    if (place !== null) return place;
  }
  return null;
}










export function presentationListIndexForPlace(
  list: readonly object[],
  place: number,
  options: PresentationPaintOrderOptions = {},
): number {
  const places = statedPlaces(list as readonly PaintedBlock[], options.reordered === true);
  const index = places.findIndex((own) => own !== null && own > place);
  return index < 0 ? list.length : index;
}

/** The list in paint order, each object with its index in the list. */
export function presentationInPaintOrder<T extends object>(
  blocks: readonly T[],
  options: PresentationPaintOrderOptions = {},
): Array<{ block: T; index: number }> {
  return presentationPaintOrder(blocks, options).map((index) => ({ block: blocks[index], index }));
}
