/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the slide-editor convenience actions (PowerPoint / WPS
 * parity): duplicate slide, text → SmartArt conversion, "reset picture" (clear
 * effects + crop), and "replace image" (swap src, keep geometry). Each is a
 * TOTAL, side-effect free transform over presentation state / block fields so it
 * can be unit tested in isolation and wired into the editor with confidence that
 * ONLY the intended fields change.
 *
 * Keep this module dependency-free of React / the DOM. It depends only on the
 * sibling PURE presentation modules (paragraphs, smartart, image-effects,
 * crop-group), mirroring their conventions: no mutation of inputs, deterministic
 * output (id generation is injected), and a byte-stable result for the no-op
 * case (e.g. resetting an already-clean image returns an equal block).
 */

import { MAX_PRESENTATION_SLIDES } from "./presentation/slide-limits";
import type {
  PresentationBlock,
  PresentationImageBlock,
} from "../components/console/editor/native-document/documentCommands";
import {
  IMAGE_EFFECT_KEYS,
} from "./presentation-image-effects";
import {
  createNativeSlideInsertionReceipt,
  nativeSlideCanDuplicate,
  parseNativeSlideInsertionReceipt,
} from "./presentation-layouts";
import {
  normalizeParagraphs,
  paragraphsFromText,
} from "./presentation-paragraphs";
import { keptObjectCopySource } from "./presentation-object-degraded";

/**
 * Structural subset of the editor's richer `SlideModel` needed to duplicate a
 * slide. Unknown JSON-safe slide fields are deep-detached; only the slide id,
 * block ids, connector endpoints, animation targets, and comment block anchors
 * are rewritten.
 */
export interface DuplicablePresentationSlide {
  id: string;
  blocks: PresentationBlock[];
  animationSequence?: unknown[];
  [key: string]: unknown;
}

export interface DuplicatePresentationSlideResult<
  T extends DuplicablePresentationSlide,
> {
  slides: T[];
  insertedIndex: number;
  clone: T;
}

export interface DuplicatePresentationSlidesResult<
  T extends DuplicablePresentationSlide,
> {
  slides: T[];
  clones: T[];
}

export interface PastePresentationSlideResult<
  T extends DuplicablePresentationSlide,
> {
  slides: T[];
  insertedIndex: number;
  clone: T;
}

/**
 * The persisted presentation schema and load edge both cap a deck at
 * `MAX_PRESENTATION_SLIDES` slides. Structural commands must enforce the same
 * limit before allocating identity so an in-memory slide past it can never be
 * saved and then disappear on reload.
 */
export { MAX_PRESENTATION_SLIDES };

type CloneResult<T> = { ok: true; value: T } | { ok: false };

const INVALID_CLONE: CloneResult<never> = { ok: false };

/**
 * Deep-copy one persistence-safe presentation value. The slide model is
 * JSON-backed, so functions, symbols, bigint, non-finite numbers, accessors,
 * exotic/custom prototypes, non-enumerable fields, and cycles are rejected
 * instead of being silently dropped or changed at autosave. Plain objects keep
 * their Object/null prototype; cloned data properties are deliberately mutable
 * even when a valid source snapshot was frozen.
 *
 * Repeated (but acyclic) references are materialized independently, matching a
 * JSON save/reload boundary and guaranteeing that no nested mutable object is
 * shared with the source or another branch of the clone.
 */
function clonePresentationValue<T>(value: T): CloneResult<T> {
  const ancestors = new WeakSet<object>();

  const visit = (current: unknown): CloneResult<unknown> => {
    if (
      current === null ||
      current === undefined ||
      typeof current === "string" ||
      typeof current === "boolean"
    ) {
      return { ok: true, value: current };
    }
    if (typeof current === "number") {
      return Number.isFinite(current)
        ? { ok: true, value: current }
        : INVALID_CLONE;
    }
    if (typeof current !== "object") return INVALID_CLONE;
    if (ancestors.has(current)) return INVALID_CLONE;
    const isArray = Array.isArray(current);
    const prototype = Object.getPrototypeOf(current);
    if (
      (isArray && prototype !== Array.prototype) ||
      (!isArray && prototype !== Object.prototype && prototype !== null)
    ) {
      return INVALID_CLONE;
    }

    const clone: object = isArray
      ? new Array(current.length)
      : Object.create(prototype);
    ancestors.add(current);
    try {
      for (const key of Reflect.ownKeys(current)) {
        if (key === "length" && isArray) continue;
        if (typeof key === "symbol") return INVALID_CLONE;
        const descriptor = Object.getOwnPropertyDescriptor(current, key);
        if (!descriptor || !("value" in descriptor) || !descriptor.enumerable) {
          return INVALID_CLONE;
        }
        if (isArray && (!/^\d+$/.test(key) || Number(key) >= current.length)) {
          return INVALID_CLONE;
        }
        const child = visit(descriptor.value);
        if (!child.ok) return INVALID_CLONE;
        Object.defineProperty(clone, key, {
          value: child.value,
          enumerable: true,
          configurable: true,
          writable: true,
        });
      }
    } finally {
      ancestors.delete(current);
    }
    return { ok: true, value: clone };
  };

  return visit(value) as CloneResult<T>;
}

function nonEmptyId(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function hasNativePptxSlideSource(slide: DuplicablePresentationSlide): boolean {
  const source =
    slide.source &&
    typeof slide.source === "object" &&
    !Array.isArray(slide.source)
      ? (slide.source as Record<string, unknown>)
      : null;
  return nonEmptyId(source?.slideId);
}

function hasNativePptxPackageAuthority(
  slide: DuplicablePresentationSlide,
): boolean {
  if (hasNativePptxSlideSource(slide)) return true;
  const source =
    slide.source &&
    typeof slide.source === "object" &&
    !Array.isArray(slide.source)
      ? (slide.source as Record<string, unknown>)
      : null;
  if (
    source?.nativeSlideInsertion &&
    typeof source.nativeSlideInsertion === "object" &&
    !Array.isArray(source.nativeSlideInsertion)
  ) {
    return true;
  }
  return slide.blocks.some((block) => {
    const blockSource =
      block.source &&
      typeof block.source === "object" &&
      !Array.isArray(block.source)
        ? (block.source as Record<string, unknown>)
        : null;
    return nonEmptyId(blockSource?.slideId);
  });
}

/**
 * Collect every live slide/block identity and reject ambiguous deck slide ids.
 * Block ids are intentionally allowed to repeat across DIFFERENT slides (Morph
 * correspondence is slide-relative), but generated ids may collide with none
 * of them.
 */
function livePresentationIdentities<T extends DuplicablePresentationSlide>(
  slides: readonly T[],
): Set<string> | null {
  const slideIds = new Set<string>();
  const liveIds = new Set<string>();
  for (const slide of slides) {
    if (!slide || typeof slide !== "object") return null;
    if (!nonEmptyId(slide.id) || slideIds.has(slide.id)) return null;
    slideIds.add(slide.id);
    liveIds.add(slide.id);
    if (!Array.isArray(slide.blocks)) return null;
    for (const block of slide.blocks) {
      if (!nonEmptyId(block?.id)) return null;
      liveIds.add(block.id);
    }
    const comments = (slide as Record<string, unknown>).comments;
    if (comments !== undefined) {
      if (!Array.isArray(comments)) return null;
      for (const comment of comments) {
        if (!comment || typeof comment !== "object") return null;
        const record = comment as Record<string, unknown>;
        if (!nonEmptyId(record.id) || liveIds.has(record.id)) return null;
        liveIds.add(record.id);
        const replies = record.replies;
        if (replies === undefined) continue;
        if (!Array.isArray(replies)) return null;
        for (const reply of replies) {
          if (!reply || typeof reply !== "object") return null;
          const replyId = (reply as Record<string, unknown>).id;
          if (!nonEmptyId(replyId) || liveIds.has(replyId)) return null;
          liveIds.add(replyId);
        }
      }
    }
  }
  return liveIds;
}

function setCloneField(
  record: Record<string, unknown>,
  key: string,
  value: unknown,
): boolean {
  return Reflect.set(record, key, value);
}

/**
 * A copied slide's comment threads as new threads of the copy: every thread
 * and reply takes a fresh id (ids are unique across a deck), keeps its words,
 * author, time and resolved state, and a thread pinned to an object is pinned
 * to that object's copy (one pinned to an object not on the slide is pinned to
 * the slide). Undefined when there are none; null when an id cannot be had.
 */
function remintCopiedComments(
  value: unknown,
  idMap: ReadonlyMap<string, string>,
  reserveFreshId: () => string | null,
): Record<string, unknown>[] | null | undefined {
  if (!Array.isArray(value) || value.length === 0) return undefined;
  const threads: Record<string, unknown>[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;
    const thread = entry as Record<string, unknown>;
    const threadId = reserveFreshId();
    if (!threadId || !setCloneField(thread, "id", threadId)) return null;
    if (typeof thread.blockId === "string") {
      const anchor = idMap.get(thread.blockId);
      if (anchor) {
        if (!setCloneField(thread, "blockId", anchor)) return null;
      } else {
        delete thread.blockId;
      }
    }
    if (Array.isArray(thread.replies)) {
      const replies: Record<string, unknown>[] = [];
      for (const reply of thread.replies) {
        if (!reply || typeof reply !== "object" || Array.isArray(reply)) continue;
        const replyId = reserveFreshId();
        const copied = reply as Record<string, unknown>;
        if (!replyId || !setCloneField(copied, "id", replyId)) return null;
        replies.push(copied);
      }
      if (!setCloneField(thread, "replies", replies)) return null;
    }
    threads.push(thread);
  }
  return threads;
}

/** Validate one selected source and produce its detached graph without ids. */
function preparePresentationSlideClone<T extends DuplicablePresentationSlide>(
  source: T,
): CloneResult<T> {
  if (!source || typeof source !== "object" || !Array.isArray(source.blocks)) {
    return INVALID_CLONE;
  }
  const sourceBlockIds = new Set<string>();
  for (const block of source.blocks) {
    if (!block || !nonEmptyId(block.id) || sourceBlockIds.has(block.id)) {
      return INVALID_CLONE;
    }
    sourceBlockIds.add(block.id);
  }
  return clonePresentationValue(source);
}

/**
 * Freeze the exact JSON-safe slide graph copied from the thumbnail rail.
 * Clipboard state must not retain live nested references: later author edits
 * cannot rewrite what a prior Copy/Cut action will paste.
 */
export function snapshotPresentationSlideForClipboard<
  T extends DuplicablePresentationSlide,
>(source: T): T | null {
  const prepared = preparePresentationSlideClone(source);
  return prepared.ok ? prepared.value : null;
}

/**
 * Duplicate one slide immediately after itself, deep-copying the full slide
 * graph and giving the slide and every block a fresh id. References that point
 * within the source slide (connector endpoints, animation/trigger targets, and
 * comment block anchors) are remapped onto the clones.
 *
 * This is deliberately a pure domain transform: it neither selects the clone
 * nor commits history. The command adapter owns those host effects so Ribbon,
 * shortcut, and context-menu entry points can share one transaction boundary.
 * Invalid indexes, cap overflow, malformed/non-persistable input, and identity
 * collisions fail closed with `null`; unexpected id-factory errors propagate.
 */
export function duplicatePresentationSlideAt<
  T extends DuplicablePresentationSlide,
>(
  slides: readonly T[],
  sourceIndex: number,
  nextId: () => string,
): DuplicatePresentationSlideResult<T> | null {
  const source = slides[sourceIndex];
  if (!source) return null;
  const nativeSource = hasNativePptxSlideSource(source);
  if (nativeSource && !nativeSlideCanDuplicate(source)) return null;

  // Match the persisted/load-edge cap before calling nextId. This keeps the
  // rejected action byte-stable and prevents a transient 201st slide that the
  // next reload would silently truncate.
  if (slides.length >= MAX_PRESENTATION_SLIDES) return null;

  const liveIds = livePresentationIdentities(slides);
  if (!liveIds) return null;

  // Clone the complete slide graph (including unknown index-signature extras)
  // before rewriting identity. A source-local duplicate block id makes
  // connector/comment/timing remapping ambiguous. Invalid/non-persistable
  // graphs fail closed; an unexpected clone exception remains visible.
  const clonedSource = preparePresentationSlideClone(source);
  if (!clonedSource.ok) return null;

  const reserveFreshId = (): string | null => {
    // Do not catch nextId: generator failures must remain observable. Empty or
    // colliding values are expected validation failures and return null.
    const candidate = nextId();
    if (!nonEmptyId(candidate) || liveIds.has(candidate)) return null;
    liveIds.add(candidate);
    return candidate;
  };

  const idMap = new Map<string, string>();
  for (const block of source.blocks) {
    const freshId = reserveFreshId();
    if (!freshId) return null;
    idMap.set(block.id, freshId);
  }
  const slideId = reserveFreshId();
  if (!slideId) return null;
  // A slide New Slide or Duplicate made and not saved yet carries no package
  // slide, only its insertion receipt: its copy is another slide of that
  // kind, on the same layout, and carries the same receipt under its own id
  
  // package slide, 422 presentation_native_slide_copy_source_invalid). Only
  // a receipt the slide holds for itself is carried; a slide with none gets
  // none, and the server's check stays the one gate.
  const heldInsertionReceipt = nativeSource
    ? null
    : parseNativeSlideInsertionReceipt(
        source.source && typeof source.source === "object" && !Array.isArray(source.source)
          ? (source.source as Record<string, unknown>).nativeSlideInsertion
          : undefined,
      );
  const nativeInsertionReceipt = nativeSource
    ? createNativeSlideInsertionReceipt(source, slideId, "duplicate")
    : heldInsertionReceipt && heldInsertionReceipt.modelSlideId === source.id
      ? { ...heldInsertionReceipt, modelSlideId: slideId }
      : null;
  if (nativeSource && !nativeInsertionReceipt) return null;

  const clone = clonedSource.value;
  const cloneRecord = clone as Record<string, unknown>;
  if (!setCloneField(cloneRecord, "id", slideId)) return null;
  const clonedBlocks = cloneRecord.blocks;
  if (
    !Array.isArray(clonedBlocks) ||
    clonedBlocks.length !== source.blocks.length
  ) {
    return null;
  }
  for (let index = 0; index < clonedBlocks.length; index += 1) {
    const block = clonedBlocks[index];
    if (!block || typeof block !== "object") return null;
    const next = block as PresentationBlock & Record<string, unknown>;
    const sourceBlock = source.blocks[index];
    const remappedId = idMap.get(sourceBlock.id);
    if (!remappedId || !setCloneField(next, "id", remappedId)) return null;
    delete next.source;
    // An object the import kept read-only (TRIAGE 152) stays kept in the copy:
    // its record and the package object it copies (`keptFrom`), whose own XML
    // the save takes for it -- never the page's drawing of it.
    const kept = keptObjectCopySource(sourceBlock, source as unknown as { source?: unknown });
    if (kept) {
      if (!setCloneField(next, "source", kept) || !setCloneField(next, "locked", true)) return null;
    }
    // The copy is the same object for Morph (PowerPoint matches the objects of
    // a duplicated slide): name the original it came from. A native package
    // slide is copied by the package writer, which keeps its own identity.
    if (!nativeSource) {
      const origin =
        typeof sourceBlock.morphSourceId === "string" && sourceBlock.morphSourceId
          ? sourceBlock.morphSourceId
          : sourceBlock.id;
      if (!setCloneField(next, "morphSourceId", origin)) return null;
    }
    for (const key of ["startShapeId", "endShapeId"] as const) {
      const ref = next[key];
      if (
        typeof ref === "string" &&
        idMap.has(ref) &&
        !setCloneField(next, key, idMap.get(ref))
      ) {
        return null;
      }
    }
    // A Sound / Movie action button plays a player on its own slide: the
    // copy's button plays the copy's player (PPT-B077: it kept the original's
    // id, found nothing on the copy and did nothing).
    const action = next.clickAction;
    if (
      action &&
      typeof action === "object" &&
      typeof (action as { targetBlockId?: unknown }).targetBlockId === "string" &&
      idMap.has((action as { targetBlockId: string }).targetBlockId) &&
      !setCloneField(next, "clickAction", {
        ...(action as Record<string, unknown>),
        targetBlockId: idMap.get((action as { targetBlockId: string }).targetBlockId),
      })
    ) {
      return null;
    }
  }

  if (Array.isArray(cloneRecord.animationSequence)) {
    for (const step of cloneRecord.animationSequence) {
      if (!step || typeof step !== "object") continue;
      const remapped = step as Record<string, unknown>;
      for (const key of ["blockId", "triggerBlockId"] as const) {
        const ref = remapped[key];
        if (
          typeof ref === "string" &&
          idMap.has(ref) &&
          !setCloneField(remapped, key, idMap.get(ref))
        ) {
          return null;
        }
      }
    }
  }

  // A slide copied whole keeps its comments, as PowerPoint's does ("any
  // comments or animation effects on the slide are included", Copy and paste
  // in PowerPoint for the web). Each thread and reply is a new one with its
  
  // A slide under an imported package's comment custody is copied by the
  // server's whole-slide copy, which re-mints the threads in the package
  // itself; the local fallback keeps the copy comment-free there and the host
  // commit injects the target deck's authority onto it.
  const carriedComments =
    !nativeSource && cloneRecord.commentAuthority === undefined
      ? remintCopiedComments(cloneRecord.comments, idMap, reserveFreshId)
      : undefined;
  if (carriedComments === null) return null;
  delete cloneRecord.comments;
  if (carriedComments && carriedComments.length > 0) {
    cloneRecord.comments = carriedComments;
  }
  delete cloneRecord.commentAuthority;
  delete cloneRecord.source;
  if (nativeInsertionReceipt) {
    cloneRecord.source = { nativeSlideInsertion: nativeInsertionReceipt };
  }

  const insertedIndex = sourceIndex + 1;
  return {
    slides: [
      ...slides.slice(0, insertedIndex),
      clone,
      ...slides.slice(insertedIndex),
    ],
    insertedIndex,
    clone,
  };
}

/**
 * Paste one detached thumbnail-rail slide immediately after `afterIndex`.
 *
 * From-scratch decks can place the snapshot at any gap. Native PPTX patchback
 * can only prove a duplicate directly after the unchanged authenticated source
 * slide, so every other native target fails closed instead of minting an
 * insertion receipt the package writer would reject.
 */
export function pastePresentationSlideAfter<
  T extends DuplicablePresentationSlide,
>(
  slides: readonly T[],
  clipboardSlide: T,
  afterIndex: number,
  nextId: () => string,
): PastePresentationSlideResult<T> | null {
  if (
    afterIndex < 0 ||
    afterIndex >= slides.length ||
    slides.length >= MAX_PRESENTATION_SLIDES
  ) {
    return null;
  }
  const preparedClipboard = preparePresentationSlideClone(clipboardSlide);
  if (!preparedClipboard.ok) return null;

  const deckHasNativeAuthority = slides.some(hasNativePptxPackageAuthority);
  const clipboardHasNativeAuthority = hasNativePptxPackageAuthority(
    preparedClipboard.value,
  );
  if (deckHasNativeAuthority || clipboardHasNativeAuthority) {
    const liveSource = slides[afterIndex];
    const liveSourceRecord = liveSource.source as
      | Record<string, unknown>
      | undefined;
    const clipboardSourceRecord = preparedClipboard.value.source as
      | Record<string, unknown>
      | undefined;
    if (
      !hasNativePptxSlideSource(liveSource) ||
      !hasNativePptxSlideSource(preparedClipboard.value) ||
      liveSourceRecord?.slideId !== clipboardSourceRecord?.slideId ||
      JSON.stringify(liveSource) !== JSON.stringify(preparedClipboard.value)
    ) {
      return null;
    }
    return duplicatePresentationSlideAt(slides, afterIndex, nextId);
  }

  const destinationIds = livePresentationIdentities(slides);
  if (!destinationIds) return null;
  const detached = duplicatePresentationSlideAt(
    [preparedClipboard.value],
    0,
    () => {
      const candidate = nextId();
      if (!nonEmptyId(candidate) || destinationIds.has(candidate)) return "";
      destinationIds.add(candidate);
      return candidate;
    },
  );
  if (!detached) return null;
  const insertedIndex = afterIndex + 1;
  return {
    slides: [
      ...slides.slice(0, insertedIndex),
      detached.clone,
      ...slides.slice(insertedIndex),
    ],
    insertedIndex,
    clone: detached.clone,
  };
}

/**
 * Duplicate every addressed source slide in deck order as one pure batch.
 *
 * This deliberately delegates each clone to {@link duplicatePresentationSlideAt}
 * so single-slide and Slide Sorter duplication cannot drift on block,
 * connector, or animation identity. Callers still own the single transaction
 * boundary for the whole returned batch. Empty or stale selections fail closed
 * with `null` and consume no identity.
 */
export function duplicatePresentationSlidesById<
  T extends DuplicablePresentationSlide,
>(
  slides: readonly T[],
  sourceIds: ReadonlySet<string>,
  nextId: () => string,
): DuplicatePresentationSlidesResult<T> | null {
  if (sourceIds.size === 0) return null;
  if (slides.length + sourceIds.size > MAX_PRESENTATION_SLIDES) return null;
  if (!livePresentationIdentities(slides)) return null;
  const availableIds = new Set(slides.map((slide) => slide.id));
  for (const sourceId of sourceIds) {
    if (!availableIds.has(sourceId)) return null;
  }

  // Preflight every addressed source before the first allocation. A malformed
  // later slide must never consume identities for an earlier provisional clone.
  for (const slide of slides) {
    if (sourceIds.has(slide.id) && !preparePresentationSlideClone(slide).ok) {
      return null;
    }
  }

  let nextSlides = [...slides];
  const clones: T[] = [];
  let insertedCount = 0;
  for (let sourceIndex = 0; sourceIndex < slides.length; sourceIndex += 1) {
    if (!sourceIds.has(slides[sourceIndex].id)) continue;
    const result = duplicatePresentationSlideAt(
      nextSlides,
      sourceIndex + insertedCount,
      nextId,
    );
    if (!result) return null;
    nextSlides = result.slides;
    clones.push(result.clone);
    insertedCount += 1;
  }

  return clones.length > 0 ? { slides: nextSlides, clones } : null;
}

/**
 * Derive SmartArt node labels from a text block's plain `text`. Splits on
 * newlines via the SAME `paragraphsFromText` mapping the editor uses (so one
 * line = one label), trims each line, and drops blank lines — an empty / all
 * whitespace block yields an empty array (the caller then degrades to a no-op).
 * Deterministic and side-effect-free.
 */
export function smartArtLabelsFromText(text: string): string[] {
  return paragraphsFromText(text)
    .map((paragraph) => paragraph.text.trim())
    .filter((line) => line.length > 0);
}

/**
 * A text block's lines as SmartArt outline items: each non-blank paragraph is
 * one node and its indent level is the node's level, as Convert to SmartArt
 * reads a bulleted list (PPT-B042: a three-level list became 0/1/1). Levels
 * are made a valid outline — the first item is the top and no item is more
 * than one level below the one before it.
 */
export function smartArtOutlineFromTextBlock(block: PresentationBlock): {
  labels: string[];
  levels: number[];
} {
  const record = block as Record<string, unknown>;
  const paragraphs =
    normalizeParagraphs(record.paragraphs) ??
    paragraphsFromText(typeof record.text === "string" ? record.text : "");
  const rows = paragraphs
    .map((paragraph) => ({ text: paragraph.text.trim(), level: paragraph.level }))
    .filter((row) => row.text.length > 0);
  const top = rows.length > 0 ? Math.min(...rows.map((row) => row.level)) : 0;
  const levels: number[] = [];
  for (const row of rows) {
    const previous = levels.length > 0 ? levels[levels.length - 1] : -1;
    levels.push(Math.max(0, Math.min(row.level - top, previous + 1)));
  }
  return { labels: rows.map((row) => row.text), levels };
}

/**
 * The geometry an inserted block carries (EMU). Optional because a flat-text
 * block (no canvas geometry) has none — in that case the SmartArt falls back to
 * the slide bounds.
 */
export interface BlockGeometry {
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

/** Read a finite positive number off a record, else `undefined`. */
function positiveNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : undefined;
}

/** Read a finite number off a record, else `undefined` (0 / negatives allowed). */
function finiteNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

/**
 * The geometry (left/top/width/height) a block carries, with only the fields
 * that are present + valid. Used to place a converted SmartArt where the source
 * text block sat.
 */
export function blockGeometry(block: PresentationBlock): BlockGeometry {
  const record = block as Record<string, unknown>;
  const out: BlockGeometry = {};
  const left = finiteNumber(record.left);
  const top = finiteNumber(record.top);
  const width = positiveNumber(record.width);
  const height = positiveNumber(record.height);
  if (left !== undefined) out.left = left;
  if (top !== undefined) out.top = top;
  if (width !== undefined) out.width = width;
  if (height !== undefined) out.height = height;
  return out;
}

/**
 * Return a NEW image block with every visual effect (shadow / frame / corner
 * radius) AND any crop cleared back to the defaults, while KEEPING the image
 * itself: `src`, geometry (left/top/width/height/rotation), `alt`, `caption`,
 * `groupId`, animations, hyperlinks, and any other passthrough field stay
 * untouched. The effect / crop keys are DELETED (not set to a default value), so
 * a reset image serializes byte-for-byte like an image that never had effects —
 * and resetting an already-clean image returns an equal block.
 *
 * Mirrors PowerPoint's "Reset Picture": it strips formatting, NOT the picture.
 */
export function resetImagePicture(
  block: PresentationImageBlock,
): PresentationImageBlock {
  const next = { ...block } as Record<string, unknown>;
  for (const key of IMAGE_EFFECT_KEYS) delete next[key];
  delete next.pictureStylePreset;
  // Glow, Reflection, Soft Edges, Bevel and 3-D Rotation are formatting
  // changes too: Reset Picture discards them with the rest (they stayed).
  delete next.pictureEffects;
  delete next.cropRect;
  // Reset Picture discards every formatting change, the artistic effect
  // included (PPT-B039). One a file brought in is already drawn into the
  // picture itself, so there is no original to go back to: it stays.
  if (!artisticEffectIsRendered(block)) delete next.artisticEffect;
  return next as unknown as PresentationImageBlock;
}

function artisticEffectIsRendered(block: PresentationImageBlock): boolean {
  return (block as Record<string, unknown>).artisticEffectRendered === true;
}

/**
 * Return a NEW image block whose `src` is swapped to `nextSrc`, KEEPING
 * everything else — geometry (left/top/width/height/rotation), effects (shadow /
 * frame / corner radius), crop, `alt`, `caption`, `groupId`, animation,
 * hyperlink, and any passthrough fields. Mirrors PowerPoint's "Change Picture":
 * the frame + formatting persist, only the image content changes.
 *
 * A blank / whitespace-only `nextSrc` is rejected (returns the block unchanged)
 * so a cancelled / empty file picker can never blank out the image. `src` is
 * stored verbatim; the render edge already validates the source before display.
 */
export function replaceImageSource(
  block: PresentationImageBlock,
  nextSrc: string,
): PresentationImageBlock {
  if (typeof nextSrc !== "string" || nextSrc.trim().length === 0) return block;
  // Remove Background's session (its marks and the original it starts from)
  // describes the old picture's pixels, as Erase's result no longer has one:
  // kept, the next Remove Background previewed the old picture and Keep
  
  // Change Picture starts the new picture with no removal.
  const {
    backgroundRemoval: _removal,
    backgroundRemovalOriginalSrc: _original,
    ...rest
  } = block as PresentationImageBlock & { backgroundRemoval?: unknown; backgroundRemovalOriginalSrc?: unknown };
  void _removal;
  void _original;
  return { ...rest, src: nextSrc } as PresentationImageBlock;
}

/**
 * True when an image block carries ANY non-default effect or a real crop — i.e.
 * "Reset" would actually change something. Lets the caller disable the Reset
 * affordance when there is nothing to reset (mirrors the crop-reset button's
 * disabled state). Defensive against stale values: any present effect key or a
 * `cropRect` object counts.
 */
export function imageHasResettableFormatting(
  block: PresentationImageBlock,
): boolean {
  const record = block as Record<string, unknown>;
  if (record.cropRect && typeof record.cropRect === "object") return true;
  if (typeof record.pictureStylePreset === "string" && record.pictureStylePreset) return true;
  if (
    typeof record.artisticEffect === "string" &&
    record.artisticEffect !== "" &&
    record.artisticEffect !== "none" &&
    !artisticEffectIsRendered(block)
  ) {
    return true;
  }
  return IMAGE_EFFECT_KEYS.some((key) => {
    const value = record[key];
    if (value === undefined) return false;
    // Opacity's default is 1 (opaque), so 0 is the furthest from it: a fully
    // transparent picture has the most to reset, and Reset was greyed out
    // there (PPT-B134).
    if (key === "opacity") return typeof value === "number" && value !== 1;
    // A falsy default (false / 0 / "") is not a real effect.
    if (value === false || value === 0 || value === "") return false;
    return true;
  });
}
