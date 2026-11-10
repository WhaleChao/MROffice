/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ChartDrawingFrame } from "./chart-groups";

export interface DrawingPreviewFrame {
  id: string;
  before: ChartDrawingFrame;
  after: ChartDrawingFrame;
  rotation?: number;
}

/** Gesture-only compositor projection. The model and React tree commit once
 * at pointer-up. Restore exactly the inline properties this projection owns. */
export function createDrawingPreview(getRoot: () => HTMLElement | null) {
  type Saved = { transform: string; origin: string; willChange: string; width: string; height: string; minWidth: string; minHeight: string; applied?: string; sized?: boolean };
  const saved = new Map<HTMLElement, Saved>();
  let pending: number | null = null;
  let latest: readonly DrawingPreviewFrame[] = [];
  const restoreOne = (element: HTMLElement, original: Saved) => {
    // React may already have published a remote replacement. Restore only
    // properties still carrying this gesture's own projection.
    if (element.style.transform === original.applied) element.style.transform = original.transform;
    if (element.style.transformOrigin === "0px 0px" || element.style.transformOrigin === "0 0") element.style.transformOrigin = original.origin;
    if (element.style.willChange === "transform") element.style.willChange = original.willChange;
    if (original.sized) {
      element.style.width = original.width;
      element.style.height = original.height;
      element.style.minWidth = original.minWidth;
      element.style.minHeight = original.minHeight;
    }
  };
  const restore = () => {
    for (const [element, original] of saved) restoreOne(element, original);
    saved.clear();
  };
  const paint = () => {
    pending = null;
    const root = getRoot();
    if (!root) return;
    for (const { id, before, after, rotation } of latest) {
      if (before.width <= 0 || before.height <= 0) continue;
      for (const element of root.querySelectorAll<HTMLElement>(`[data-drawing-preview-id="${CSS.escape(id)}"]`)) {
        if (!saved.has(element)) saved.set(element, { transform: element.style.transform, origin: element.style.transformOrigin, willChange: element.style.willChange, width: element.style.width, height: element.style.height, minWidth: element.style.minWidth, minHeight: element.style.minHeight });
        const original = saved.get(element)!;
        element.style.transformOrigin = "0 0";
        element.style.willChange = "transform";
        const turn = rotation === undefined ? (original.transform === "none" ? "" : original.transform) : `rotate(${rotation}deg)`;
        if (element.dataset.drawingPreviewFit === "position") {
          // Content that lays itself out again at the gesture's live size (a
          // chart, re-planned by its engine as Excel re-draws one while it is
          // dragged): the host draws it at `after`'s size, so the projection
          // only moves it. Scaling it stretched the last drawing, text and
          // lines included, until the model committed at pointer-up.
          element.style.transform = `translate(${after.left - before.left}px, ${after.top - before.top}px) translate(${after.width / 2}px, ${after.height / 2}px) ${turn} translate(${-after.width / 2}px, ${-after.height / 2}px)`;
        } else if (element.dataset.drawingPreviewFit === "box") {
          // A selection frame is chrome, not content: its border weight, its
          // resize grips and its rotate handle must stay the size the pointer
          // can still grab. Move and RESIZE the box instead of scaling it.
          original.sized = true;
          element.style.width = `${after.width}px`;
          element.style.height = `${after.height}px`;
          // The box carries min-width/min-height equal to its own size (the
          // object layer sets both, so a parent can never squeeze the chrome).
          // Width alone therefore only ever GROWS the box: a shrink is clamped
          // by the pre-gesture minimum and the frame stays around the old
          // picture until the model commits.
          element.style.minWidth = `${after.width}px`;
          element.style.minHeight = `${after.height}px`;
          element.style.transform = `translate(${after.left - before.left}px, ${after.top - before.top}px) translate(${after.width / 2}px, ${after.height / 2}px) ${turn} translate(${-after.width / 2}px, ${-after.height / 2}px)`;
        } else {
          element.style.transform = `translate(${after.left - before.left}px, ${after.top - before.top}px) translate(${after.width / 2}px, ${after.height / 2}px) ${turn} scale(${after.width / before.width}, ${after.height / before.height}) translate(${-before.width / 2}px, ${-before.height / 2}px)`;
        }
        original.applied = element.style.transform;

      }
    }
  };
  return {
    schedule(frames: readonly DrawingPreviewFrame[]) {
      latest = frames;
      if (pending === null) pending = requestAnimationFrame(paint);
    },
    clear() {
      if (pending !== null) cancelAnimationFrame(pending);
      pending = null;
      latest = [];
      restore();
    },
    /** Paint the scheduled frames now (the committed frames under them just changed). */
    flush() {
      if (pending === null) return;
      cancelAnimationFrame(pending);
      paint();
    },
    /** One object leaves the projection (the others keep theirs, unrepainted). */
    release(id: string) {
      latest = latest.filter((frame) => frame.id !== id);
      for (const [element, original] of saved) {
        if (element.dataset.drawingPreviewId !== id) continue;
        restoreOne(element, original);
        saved.delete(element);
      }
    },
  };
}

export interface DrawingLiveSize {
  width: number;
  height: number;
}

/**
 * The size each object is being dragged to, for content that lays itself
 * out again at it (a chart's engine re-plans at the live size). A store
 * rather than host state, so one gesture frame re-renders only the objects
 * that read it, not the window around them.
 */
export interface DrawingLiveSizeStore {
  get(id: string): DrawingLiveSize | undefined;
  set(sizes: ReadonlyMap<string, DrawingLiveSize>): void;
  subscribe(listener: () => void): () => void;
}

export function createDrawingLiveSizeStore(): DrawingLiveSizeStore {
  let sizes: ReadonlyMap<string, DrawingLiveSize> = new Map();
  const listeners = new Set<() => void>();
  return {
    get: (id) => sizes.get(id),
    set(next) {
      sizes = next;
      for (const listener of listeners) listener();
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  };
}
