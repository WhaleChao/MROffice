/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PageItem } from "./notebook-sdk";
import {
  slideFittedToDeckSize,
  statedSlideSize,
  UNSTATED_SLIDE_SIZE_EMU,
} from "./presentation-slide-size";
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import {
  DEFAULT_LAYOUT_ID,
  nativeLayoutIdFromType,
  parseNativeLayoutCascadeReceipt,
  type NativeLayoutCascadeReceipt,
  type NativeLayoutCatalogEntryReceipt,
} from "./presentation-layouts";
import type { SlideMaster } from "./presentation-master";
import { expandPresentationMedia } from "./presentation-page-media";
import { MAX_PRESENTATION_SLIDES } from "./presentation/slide-limits";

export const PPTX_REUSE_SLIDE_RECEIPT_VERSION = "pptx-reuse-slide-v1" as const;
export type PresentationReuseThemePolicy = "destination" | "source";

export interface ReusablePresentationSlide {
  id: string;
  blocks: PresentationBlock[];
  source?: Record<string, unknown>;
  layoutId?: string;
  masterId?: string;
  /** The built-in layout of a slide made here (no layout part of its own). */
  layout?: string;
  /** The slide's own size, when it states one (a deck states it per slide). */
  widthEmu?: number;
  heightEmu?: number;
  commentAuthority?: unknown;
  [key: string]: unknown;
}

export interface PresentationReuseSelectionMap {
  sourceModelSlideId: string;
  sourceSlideId: string;
  modelSlideId: string;
  destinationLayoutPart: string | null;
}

export interface NativeSlideReuseReceipt {
  schemaVersion: typeof PPTX_REUSE_SLIDE_RECEIPT_VERSION;
  planId: string;
  modelSlideId: string;
  targetPackageFingerprint: string;
  targetAnchorSlideId: string;
  destinationLayoutPart: string | null;
  sourcePageId: string;
  sourceContentVersion: number;
  sourceContentSha256: string;
  sourceResourceId: string | null;
  sourceResourceVersionId: string | null;
  sourcePackageFingerprint: string;
  sourceModelSlideId: string;
  sourceSlideId: string;
  themePolicy: PresentationReuseThemePolicy;
  selectionIndex: number;
  selectionCount: number;
  selectionMap: PresentationReuseSelectionMap[];
}

export interface PreparePresentationReuseResult<T extends ReusablePresentationSlide> {
  slides: T[];
  masters: SlideMaster[];
  clones: T[];
  insertedIndex: number;
  planId: string;
}

const SHA256 = /^[0-9a-f]{64}$/;
const MAX_REUSED_SLIDES = 20;
const MAX_DECK_SLIDES = MAX_PRESENTATION_SLIDES;
const MAX_COMMENT_RECORDS_PER_SLIDE = 20_200;

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function exactString(value: unknown, maxLength = 512): string | null {
  return typeof value === "string" && value.length > 0 && value.length <= maxLength
    ? value
    : null;
}

function deepJsonClone<T>(value: T): T | null {
  try {
    return JSON.parse(JSON.stringify(value)) as T;
  } catch {
    return null;
  }
}

function flattenCommentRecords(
  value: unknown,
): Record<string, unknown>[] | null {
  if (value === undefined) return [];
  if (!Array.isArray(value) || value.length > 200) return null;
  const records: Record<string, unknown>[] = [];
  const pending: unknown[] = [...value];
  for (let cursor = 0; cursor < pending.length; cursor += 1) {
    const comment = record(pending[cursor]);
    if (!comment) return null;
    records.push(comment);
    if (records.length > MAX_COMMENT_RECORDS_PER_SLIDE) return null;
    const replies = comment.replies;
    if (replies === undefined) continue;
    if (!Array.isArray(replies) || replies.length > 100) return null;
    if (pending.length + replies.length > MAX_COMMENT_RECORDS_PER_SLIDE) {
      return null;
    }
    pending.push(...replies);
  }
  return records;
}

function nativeSlideReceipt(slide: ReusablePresentationSlide) {
  return parseNativeLayoutCascadeReceipt(record(slide.source)?.layoutCascade);
}

function placeholderSignature(
  value: { placeholderType: string; placeholderIndex: string }[],
): string {
  return value
    .map((item) => `${item.placeholderType}:${item.placeholderIndex}`)
    .sort()
    .join("|");
}

function destinationLayoutForSource(
  source: NativeLayoutCascadeReceipt,
  target: NativeLayoutCascadeReceipt,
): NativeLayoutCatalogEntryReceipt | null {
  const sourceSignature = placeholderSignature(source.placeholders);
  return (
    target.layouts.find(
      (layout) =>
        layout.layoutType === source.layoutType &&
        placeholderSignature(layout.placeholders) === sourceSignature,
    ) ??
    target.layouts.find((layout) => layout.layoutType === source.layoutType) ??
    target.layouts.find((layout) => layout.layoutPart === target.layoutPart) ??
    null
  );
}

function remapSlideLinks(
  value: unknown,
  slideIdMap: ReadonlyMap<string, string>,
): boolean {
  if (Array.isArray(value)) {
    return value.every((item) => remapSlideLinks(item, slideIdMap));
  }
  if (!value || typeof value !== "object") return true;
  const next = value as Record<string, unknown>;
  for (const [key, item] of Object.entries(next)) {
    if (key === "targetSlideId" && typeof item === "string") {
      const targetId = slideIdMap.get(item);
      if (!targetId) return false;
      next[key] = targetId;
    } else if (typeof item === "string" && item.startsWith("slide://")) {
      const sourceId = item.slice("slide://".length);
      const targetId = slideIdMap.get(sourceId);
      if (!targetId) return false;
      next[key] = `slide://${targetId}`;
    } else if (!remapSlideLinks(item, slideIdMap)) {
      return false;
    }
  }
  return true;
}

function sourcePackageIdentity(page: PageItem, sourceSlides: ReusablePresentationSlide[]) {
  const marker = record(page.editable_page_patchback);
  if (marker?.status === "succeeded") {
    const resourceId = exactString(marker.resource_id, 128);
    const versionId = exactString(marker.resource_version_id, 128);
    const outputSha256 = exactString(marker.output_sha256, 64);
    if (!resourceId || !versionId || !outputSha256 || !SHA256.test(outputSha256)) {
      return null;
    }
    return {
      resourceId,
      versionId,
      packageFingerprint: outputSha256,
    };
  }
  if ((page.content_version ?? 1) !== 1) return null;
  const fingerprints = new Set(
    sourceSlides
      .map((slide) => nativeSlideReceipt(slide)?.packageFingerprint)
      .filter((value): value is string => Boolean(value)),
  );
  const packageFingerprint = fingerprints.size === 1 ? [...fingerprints][0] : null;
  if (!packageFingerprint || !SHA256.test(packageFingerprint)) return null;
  return { resourceId: null, versionId: null, packageFingerprint };
}

type ReserveId = () => string | null;

/**
 * A copy of `source` for another deck: its own id, fresh ids for its blocks
 * and comments with every reference among them (connector ends, animation
 * steps, comment replies and anchors, links to other reused slides)
 * following, and the destination's comment authority. `null` when an id
 * cannot be reserved or a link points at a slide that is not coming along.
 */
function cloneReusedSlide<T extends ReusablePresentationSlide>(
  source: ReusablePresentationSlide,
  slideId: string,
  reserveId: ReserveId,
  targetCommentAuthority: Record<string, unknown> | null | undefined,
  slideIdMap: ReadonlyMap<string, string>,
): T | null {
  const clone = deepJsonClone(source) as T | null;
  if (!clone || !Array.isArray(clone.blocks)) return null;
  clone.id = slideId;
  const blockIdMap = new Map<string, string>();
  for (const block of source.blocks) {
    const id = reserveId();
    if (!id) return null;
    blockIdMap.set(block.id, id);
  }
  for (let blockIndex = 0; blockIndex < clone.blocks.length; blockIndex += 1) {
    const block = clone.blocks[blockIndex] as PresentationBlock &
      Record<string, unknown>;
    const sourceBlock = source.blocks[blockIndex];
    const id = blockIdMap.get(sourceBlock.id);
    if (!id) return null;
    block.id = id;
    for (const key of ["startShapeId", "endShapeId"] as const) {
      const current = block[key];
      if (typeof current === "string" && blockIdMap.has(current)) {
        block[key] = blockIdMap.get(current);
      }
    }
  }
  if (Array.isArray(clone.animationSequence)) {
    for (const step of clone.animationSequence) {
      for (const key of ["blockId", "triggerBlockId"] as const) {
        const current = step[key];
        if (typeof current === "string" && blockIdMap.has(current)) {
          step[key] = blockIdMap.get(current);
        }
      }
    }
  }
  const commentRecords = flattenCommentRecords(clone.comments);
  if (commentRecords === null) return null;
  const commentIdMap = new Map<string, string>();
  for (const comment of commentRecords) {
    const current = exactString(comment.id, 256);
    const id = current ? reserveId() : null;
    if (!current || !id || commentIdMap.has(current)) return null;
    commentIdMap.set(current, id);
  }
  for (const comment of commentRecords) {
    const current = String(comment.id);
    comment.id = commentIdMap.get(current)!;
    if (typeof comment.parentId === "string") {
      const parentId = commentIdMap.get(comment.parentId);
      if (!parentId) return null;
      comment.parentId = parentId;
    }
    if (
      typeof comment.blockId === "string" &&
      blockIdMap.has(comment.blockId)
    ) {
      comment.blockId = blockIdMap.get(comment.blockId);
    }
  }
  if (!targetCommentAuthority) {
    delete clone.commentAuthority;
  } else {
    const inheritedAuthority = deepJsonClone(targetCommentAuthority);
    if (!inheritedAuthority) return null;
    clone.commentAuthority = inheritedAuthority;
  }
  if (!remapSlideLinks(clone, slideIdMap)) return null;
  return clone;
}

/**
 * Keep Source Formatting: the source masters the reused slides use, copied
 * into the destination under fresh ids (their layouts too), and a master
 * carrying the source theme for a slide that names none. Use Destination
 * Theme copies nothing.
 */
function copyReusedSourceMasters(
  targetMasters: readonly SlideMaster[],
  sourceContent: Record<string, unknown> | null,
  selected: readonly ReusablePresentationSlide[],
  themePolicy: PresentationReuseThemePolicy,
  reserveId: ReserveId,
): {
  nextMasters: SlideMaster[];
  masterIdMap: Map<string, string>;
  layoutIdMap: Map<string, string>;
  fallbackSourceMasterId: string | null;
} | null {
  const sourceMasters = Array.isArray(sourceContent?.masters)
    ? (deepJsonClone(sourceContent.masters) as SlideMaster[] | null) ?? []
    : [];
  const nextMasters = deepJsonClone([...targetMasters]) ?? [];
  const masterIdMap = new Map<string, string>();
  const layoutIdMap = new Map<string, string>();
  let fallbackSourceMasterId: string | null = null;
  if (themePolicy === "source") {
    const selectedMasterIds = new Set(
      selected
        .map((slide) => exactString(slide.masterId, 512))
        .filter((value): value is string => Boolean(value)),
    );
    for (const sourceMaster of sourceMasters) {
      if (!selectedMasterIds.has(sourceMaster.id)) continue;
      const masterId = reserveId();
      if (!masterId) return null;
      masterIdMap.set(sourceMaster.id, masterId);
      const master = deepJsonClone(sourceMaster);
      if (!master) return null;
      master.id = masterId;
      master.name = sourceMaster.name || "Source master";
      if (Array.isArray(master.layouts)) {
        for (const layout of master.layouts) {
          const layoutId = reserveId();
          if (!layoutId) return null;
          layoutIdMap.set(layout.id, layoutId);
          layout.id = layoutId;
        }
      }
      nextMasters.push(master);
    }
    const sourceTheme = record(sourceContent?.theme);
    for (const sourceSlide of selected) {
      const sourceMasterId = exactString(sourceSlide.masterId, 512);
      if (sourceMasterId && masterIdMap.has(sourceMasterId)) continue;
      if (!sourceMasterId && fallbackSourceMasterId) continue;
      const syntheticId = reserveId();
      if (!syntheticId) return null;
      if (sourceMasterId) {
        masterIdMap.set(sourceMasterId, syntheticId);
      } else {
        fallbackSourceMasterId = syntheticId;
      }
      nextMasters.push({
        id: syntheticId,
        name: "Source formatting",
        ...(sourceTheme ? { theme: deepJsonClone(sourceTheme) ?? undefined } : {}),
      });
    }
  }

  return { nextMasters, masterIdMap, layoutIdMap, fallbackSourceMasterId };
}

/**
 * Fresh ids for the reused slides, their blocks and comments: none may be
 * one either deck already uses. `null` when a comment carries no id.
 */
function createReuseIdReserver(
  targetSlides: readonly ReusablePresentationSlide[],
  sourceSlides: readonly ReusablePresentationSlide[],
  nextId: () => string,
): ReserveId | null {
  const liveIds = new Set<string>();
  for (const slide of [...targetSlides, ...sourceSlides]) {
    liveIds.add(slide.id);
    for (const block of slide.blocks) liveIds.add(block.id);
    const comments = flattenCommentRecords(slide.comments);
    if (comments === null) return null;
    for (const comment of comments) {
      const commentId = exactString(comment.id, 256);
      if (!commentId) return null;
      liveIds.add(commentId);
    }
  }
  return () => {
    const candidate = nextId();
    if (!exactString(candidate, 256) || liveIds.has(candidate)) return null;
    liveIds.add(candidate);
    return candidate;
  };
}

/** The presentation slides a source page carries, its media expanded. */
function reuseSourceSlides(sourcePage: PageItem): {
  content: Record<string, unknown> | null;
  slides: ReusablePresentationSlide[];
} {
  // The editor works on the bytes of a picture the source stores once
  // (`presentation-page-media`); the target deck stores it again itself.
  const content = record(expandPresentationMedia(sourcePage.content_json));
  const slides = Array.isArray(content?.slides)
    ? (content.slides.filter(
        (slide): slide is ReusablePresentationSlide =>
          Boolean(record(slide)) &&
          exactString(record(slide)?.id, 256) !== null &&
          Array.isArray(record(slide)?.blocks),
      ) as ReusablePresentationSlide[])
    : [];
  return { content, slides };
}

/**
 * The ids of the slides Reuse Slides offers from `sourcePage`, in the source
 * deck's order: the pane's thumbnails, slide 1 first.
 */
export function presentationReuseSourceSlideIds(sourcePage: Pick<PageItem, "content_json">): string[] {
  return reuseSourceSlides(sourcePage as PageItem).slides.map((slide) => slide.id);
}

export function preparePresentationReuseSlides<T extends ReusablePresentationSlide>(
  targetSlides: readonly T[],
  targetMasters: readonly SlideMaster[],
  targetIndex: number,
  sourcePage: PageItem,
  selectedSourceModelIds: ReadonlySet<string>,
  themePolicy: PresentationReuseThemePolicy,
  nextId: () => string,
): PreparePresentationReuseResult<T> | null {
  const targetAnchor = targetSlides[targetIndex];
  const targetReceipt = targetAnchor ? nativeSlideReceipt(targetAnchor) : null;
  const { content: sourceContent, slides: sourceSlides } = reuseSourceSlides(sourcePage);
  const selected = sourceSlides.filter((slide) => selectedSourceModelIds.has(slide.id));
  if (
    !targetAnchor ||
    !targetReceipt ||
    sourcePage.page_type !== "presentation" ||
    !exactString(sourcePage.id, 128) ||
    !Number.isInteger(sourcePage.content_version) ||
    (sourcePage.content_version ?? 0) < 1 ||
    !SHA256.test(sourcePage.content_sha256 ?? "") ||
    selected.length === 0 ||
    selected.length > MAX_REUSED_SLIDES ||
    selected.length !== selectedSourceModelIds.size ||
    targetSlides.length + selected.length > MAX_DECK_SLIDES
  ) {
    return null;
  }
  const targetCommentAuthorities = targetSlides.map((slide) =>
    record(slide.commentAuthority),
  );
  const targetCommentAuthorityCount = targetCommentAuthorities.filter(
    (value): value is Record<string, unknown> => value !== null,
  ).length;
  if (
    targetCommentAuthorityCount !== 0 &&
    targetCommentAuthorityCount !== targetSlides.length
  ) {
    return null;
  }
  const targetCommentAuthority = targetCommentAuthorities[0];
  if (targetCommentAuthority !== null) {
    const canonical = JSON.stringify(targetCommentAuthority);
    if (
      targetCommentAuthorities.some(
        (authority) =>
          authority === null || JSON.stringify(authority) !== canonical,
      )
    ) {
      return null;
    }
  }
  const sourcePackage = sourcePackageIdentity(sourcePage, selected);
  if (!sourcePackage) return null;
  const sourceReceipts = selected.map((slide) => nativeSlideReceipt(slide));
  if (
    sourceReceipts.some((receipt) => !receipt) ||
    sourceReceipts.some(
      (receipt) =>
        receipt?.packageFingerprint !== sourceReceipts[0]?.packageFingerprint,
    )
  ) {
    return null;
  }

  const planId = nextId();
  if (!exactString(planId, 256)) return null;
  const reserveId = createReuseIdReserver(targetSlides, sourceSlides, nextId);
  if (!reserveId) return null;
  const slideIdMap = new Map<string, string>();
  for (const slide of selected) {
    const id = reserveId();
    if (!id) return null;
    slideIdMap.set(slide.id, id);
  }
  const selectionMap: PresentationReuseSelectionMap[] = selected.map(
    (slide, index) => ({
      sourceModelSlideId: slide.id,
      sourceSlideId: sourceReceipts[index]!.slideId,
      modelSlideId: slideIdMap.get(slide.id)!,
      destinationLayoutPart:
        themePolicy === "destination"
          ? destinationLayoutForSource(sourceReceipts[index]!, targetReceipt)
              ?.layoutPart ?? null
          : null,
    }),
  );
  if (
    themePolicy === "destination" &&
    selectionMap.some((item) => item.destinationLayoutPart === null)
  ) {
    return null;
  }

  const copied = copyReusedSourceMasters(targetMasters, sourceContent, selected, themePolicy, reserveId);
  if (!copied) return null;
  const { nextMasters, masterIdMap, layoutIdMap, fallbackSourceMasterId } = copied;

  const clones: T[] = [];
  for (let index = 0; index < selected.length; index += 1) {
    const source = selected[index];
    const sourceReceipt = sourceReceipts[index]!;
    const clone = cloneReusedSlide<T>(source, selectionMap[index].modelSlideId, reserveId, targetCommentAuthority, slideIdMap);
    if (!clone) return null;
    if (themePolicy === "destination") {
      const destinationLayout = targetReceipt.layouts.find(
        (layout) =>
          layout.layoutPart === selectionMap[index].destinationLayoutPart,
      );
      if (!destinationLayout) return null;
      clone.layoutId = destinationLayout.layoutPart;
      clone.masterId = destinationLayout.masterPart;
    } else {
      const sourceMasterId = exactString(source.masterId, 512);
      const sourceLayoutId = exactString(source.layoutId, 512);
      if (sourceMasterId && masterIdMap.has(sourceMasterId)) {
        clone.masterId = masterIdMap.get(sourceMasterId);
      } else if (!sourceMasterId && fallbackSourceMasterId) {
        clone.masterId = fallbackSourceMasterId;
      }
      if (sourceLayoutId && layoutIdMap.has(sourceLayoutId)) {
        clone.layoutId = layoutIdMap.get(sourceLayoutId);
      }
    }
    const receipt: NativeSlideReuseReceipt = {
      schemaVersion: PPTX_REUSE_SLIDE_RECEIPT_VERSION,
      planId,
      modelSlideId: clone.id,
      targetPackageFingerprint: targetReceipt.packageFingerprint,
      targetAnchorSlideId: targetReceipt.slideId,
      destinationLayoutPart: selectionMap[index].destinationLayoutPart,
      sourcePageId: sourcePage.id,
      sourceContentVersion: sourcePage.content_version!,
      sourceContentSha256: sourcePage.content_sha256!,
      sourceResourceId: sourcePackage.resourceId,
      sourceResourceVersionId: sourcePackage.versionId,
      sourcePackageFingerprint: sourcePackage.packageFingerprint,
      sourceModelSlideId: source.id,
      sourceSlideId: sourceReceipt.slideId,
      themePolicy,
      selectionIndex: index,
      selectionCount: selected.length,
      selectionMap,
    };
    clone.source = { nativeSlideReuse: receipt };
    clones.push(clone);
  }

  const insertedIndex = targetIndex + 1;
  return {
    slides: [
      ...targetSlides.slice(0, insertedIndex),
      ...clones,
      ...targetSlides.slice(insertedIndex),
    ],
    masters: nextMasters,
    clones,
    insertedIndex,
    planId,
  };
}

/**
 * Reuse Slides into a deck made here, which writes its own file rather than
 * patching a PowerPoint package: Home ▸ New Slide ▸ Reuse Slides inserts the
 * chosen slides after the current one (Microsoft Support, "Reuse (import)
 * slides from another presentation"). Use Destination Theme puts each slide
 * on the destination's layout of its kind; Keep Source Formatting brings the
 * source master along. Any presentation is a source, one made here or an
 * imported one; an imported slide's tie to its file stays behind, since the
 * destination's file is written from its own model.
 */
export function prepareAuthoredPresentationReuseSlides<T extends ReusablePresentationSlide>(
  targetSlides: readonly T[],
  targetMasters: readonly SlideMaster[],
  targetIndex: number,
  sourcePage: PageItem,
  selectedSourceModelIds: ReadonlySet<string>,
  themePolicy: PresentationReuseThemePolicy,
  nextId: () => string,
): PreparePresentationReuseResult<T> | null {
  const targetAnchor = targetSlides[targetIndex];
  const { content: sourceContent, slides: sourceSlides } = reuseSourceSlides(sourcePage);
  const selected = sourceSlides.filter((slide) => selectedSourceModelIds.has(slide.id));
  if (
    !targetAnchor ||
    sourcePage.page_type !== "presentation" ||
    selected.length === 0 ||
    selected.length > MAX_REUSED_SLIDES ||
    selected.length !== selectedSourceModelIds.size ||
    targetSlides.length + selected.length > MAX_DECK_SLIDES
  ) {
    return null;
  }
  const targetCommentAuthority = record(targetAnchor.commentAuthority);
  const planId = nextId();
  if (!exactString(planId, 256)) return null;
  const reserveId = createReuseIdReserver(targetSlides, sourceSlides, nextId);
  if (!reserveId) return null;
  const slideIdMap = new Map<string, string>();
  for (const slide of selected) {
    const id = reserveId();
    if (!id) return null;
    slideIdMap.set(slide.id, id);
  }
  const copied = copyReusedSourceMasters(targetMasters, sourceContent, selected, themePolicy, reserveId);
  if (!copied) return null;
  const { nextMasters, masterIdMap, layoutIdMap, fallbackSourceMasterId } = copied;
  const clones: T[] = [];
  // One presentation, one slide size: a slide from a deck of another size is
  // fitted to this one (as Slide Size's Ensure Fit fits it), not kept at its
  // own size beside slides of another.
  const deckSize = statedSlideSize(targetAnchor) ?? UNSTATED_SLIDE_SIZE_EMU;
  for (const source of selected) {
    const copy = cloneReusedSlide<T>(source, slideIdMap.get(source.id)!, reserveId, targetCommentAuthority, slideIdMap);
    if (!copy) return null;
    const clone = slideFittedToDeckSize(copy, deckSize);
    const receipt = nativeSlideReceipt(source);
    // The slide leaves its file behind: no package identity on the slide
    // or on its blocks (which would make the destination an imported deck).
    delete clone.source;
    for (const block of clone.blocks as (PresentationBlock & Record<string, unknown>)[]) {
      const blockSource = record(block.source);
      if (blockSource && "slideId" in blockSource) {
        const rest = { ...blockSource };
        delete rest.slideId;
        if (Object.keys(rest).length) block.source = rest;
        else delete block.source;
      }
    }
    if (themePolicy === "destination") {
      // The destination's layout of the slide's kind, on its own master.
      const layout = typeof source.layout === "string"
        ? source.layout
        : receipt
          ? nativeLayoutIdFromType(receipt.layoutType)
          : null;
      delete clone.layoutId;
      delete clone.masterId;
      if (layout && layout !== DEFAULT_LAYOUT_ID) clone.layout = layout;
      else delete clone.layout;
    } else {
      const sourceMasterId = exactString(source.masterId, 512);
      const sourceLayoutId = exactString(source.layoutId, 512);
      if (sourceMasterId && masterIdMap.has(sourceMasterId)) {
        clone.masterId = masterIdMap.get(sourceMasterId);
      } else if (!sourceMasterId && fallbackSourceMasterId) {
        clone.masterId = fallbackSourceMasterId;
      } else {
        delete clone.masterId;
      }
      if (sourceLayoutId && layoutIdMap.has(sourceLayoutId)) {
        clone.layoutId = layoutIdMap.get(sourceLayoutId);
      } else {
        delete clone.layoutId;
      }
    }
    clones.push(clone);
  }
  const insertedIndex = targetIndex + 1;
  return {
    slides: [
      ...targetSlides.slice(0, insertedIndex),
      ...clones,
      ...targetSlides.slice(insertedIndex),
    ],
    masters: nextMasters,
    clones,
    insertedIndex,
    planId,
  };
}

export function parseNativeSlideReuseReceipt(
  value: unknown,
): NativeSlideReuseReceipt | null {
  const candidate = record(value);
  if (!candidate || candidate.schemaVersion !== PPTX_REUSE_SLIDE_RECEIPT_VERSION) {
    return null;
  }
  const expectedKeys = new Set([
    "schemaVersion",
    "planId",
    "modelSlideId",
    "targetPackageFingerprint",
    "targetAnchorSlideId",
    "destinationLayoutPart",
    "sourcePageId",
    "sourceContentVersion",
    "sourceContentSha256",
    "sourceResourceId",
    "sourceResourceVersionId",
    "sourcePackageFingerprint",
    "sourceModelSlideId",
    "sourceSlideId",
    "themePolicy",
    "selectionIndex",
    "selectionCount",
    "selectionMap",
  ]);
  if (
    Object.keys(candidate).length !== expectedKeys.size ||
    Object.keys(candidate).some((key) => !expectedKeys.has(key))
  ) {
    return null;
  }
  const strings = [
    "planId",
    "modelSlideId",
    "targetPackageFingerprint",
    "targetAnchorSlideId",
    "sourcePageId",
    "sourceContentSha256",
    "sourcePackageFingerprint",
    "sourceModelSlideId",
    "sourceSlideId",
  ] as const;
  if (strings.some((key) => exactString(candidate[key]) === null)) return null;
  if (
    !SHA256.test(String(candidate.targetPackageFingerprint)) ||
    !SHA256.test(String(candidate.sourceContentSha256)) ||
    !SHA256.test(String(candidate.sourcePackageFingerprint)) ||
    !/^sld[1-9][0-9]*$/.test(String(candidate.targetAnchorSlideId)) ||
    !/^sld[1-9][0-9]*$/.test(String(candidate.sourceSlideId)) ||
    (candidate.themePolicy !== "destination" && candidate.themePolicy !== "source") ||
    !Number.isInteger(candidate.sourceContentVersion) ||
    Number(candidate.sourceContentVersion) < 1 ||
    !Number.isInteger(candidate.selectionIndex) ||
    !Number.isInteger(candidate.selectionCount) ||
    Number(candidate.selectionCount) < 1 ||
    Number(candidate.selectionCount) > MAX_REUSED_SLIDES ||
    Number(candidate.selectionIndex) < 0 ||
    Number(candidate.selectionIndex) >= Number(candidate.selectionCount)
  ) {
    return null;
  }
  const destinationLayout = candidate.destinationLayoutPart;
  if (
    (candidate.themePolicy === "destination" &&
      (typeof destinationLayout !== "string" ||
        !/^ppt\/slideLayouts\/[^/]+\.xml$/.test(destinationLayout))) ||
    (candidate.themePolicy === "source" && destinationLayout !== null)
  ) {
    return null;
  }
  const resourceId = candidate.sourceResourceId;
  const versionId = candidate.sourceResourceVersionId;
  if (
    (resourceId === null) !== (versionId === null) ||
    (resourceId !== null && exactString(resourceId, 128) === null) ||
    (versionId !== null && exactString(versionId, 128) === null)
  ) {
    return null;
  }
  const selectionMap = candidate.selectionMap;
  if (
    !Array.isArray(selectionMap) ||
    selectionMap.length !== Number(candidate.selectionCount) ||
    selectionMap.some((item) => {
      const entry = record(item);
      return (
        !entry ||
        Object.keys(entry).length !== 4 ||
        exactString(entry.sourceModelSlideId, 256) === null ||
        exactString(entry.modelSlideId, 256) === null ||
        exactString(entry.sourceSlideId, 128) === null ||
        !/^sld[1-9][0-9]*$/.test(String(entry.sourceSlideId)) ||
        (candidate.themePolicy === "destination"
          ? typeof entry.destinationLayoutPart !== "string" ||
            !/^ppt\/slideLayouts\/[^/]+\.xml$/.test(
              entry.destinationLayoutPart,
            )
          : entry.destinationLayoutPart !== null)
      );
    })
  ) {
    return null;
  }
  const selected = record(selectionMap[Number(candidate.selectionIndex)]);
  if (
    selected?.sourceModelSlideId !== candidate.sourceModelSlideId ||
    selected?.sourceSlideId !== candidate.sourceSlideId ||
    selected?.modelSlideId !== candidate.modelSlideId ||
    selected?.destinationLayoutPart !== candidate.destinationLayoutPart
  ) {
    return null;
  }
  const cloned = deepJsonClone(candidate);
  return cloned as NativeSlideReuseReceipt | null;
}
