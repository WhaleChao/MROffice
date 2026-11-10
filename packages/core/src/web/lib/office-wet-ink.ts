/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Wet ink: finished strokes the document has not taken yet.
 *
 * Taking a stroke into the document (the deck's blocks, the worksheet's
 * objects, the Docs node tree) is the expensive part of inking: the window
 * renders, the save serializes, history closes. Done on every lift it runs
 * in the gap between two strokes, where the next stroke's first samples
 * queue behind it and the ink starts late. Office keeps the strokes "wet"
 * on the surface and takes them in a batch when the hand pauses.
 *
 * This is the batching, host-neutral: a surface adds each finished stroke,
 * draws the wet ones itself, and the batch lands through `onFlush` after a
 * pause of {@link WET_INK_SETTLE_MS} -- never while the pen is down (the
 * next lift schedules it again) -- or at once after {@link WET_INK_MAX_STROKES},
 * or when the host asks (`flush`: a mode change, an eraser about to measure
 * the document, the surface going away).
 */

import { inkPointerArbiter, watchPenHover } from "./ink-pointer-policy";

/** Dispatched on `window` to have every surface's wet strokes land at once
 * (a save about to read the document, a test). */
export const OFFICE_WET_INK_FLUSH_EVENT = "mrai:office-ink-flush";












export const WET_INK_SETTLE_MS = 900;
/** Kept for the callers that named the iPad's pause; the same pause now. */
export const WET_INK_SETTLE_MS_IPAD = WET_INK_SETTLE_MS;

/** The hand's pause before a batch lands. */
export function wetInkSettleMs(): number {
  return WET_INK_SETTLE_MS;
}
/**
 * The most wet strokes kept before they land regardless. The batch lands
 * once the pen has left the page's air or the hand pauses (see
 * {@link inkLandsAtLift}); a character is ten strokes or more, and a line of
 * fast writing a few dozen.
 */
export const WET_INK_MAX_STROKES = 64;
/** Kept for the callers that named the iPad's count; the same count now. */
export const WET_INK_MAX_STROKES_IPAD = WET_INK_MAX_STROKES;

/** The most wet strokes kept before they land regardless. */
export function wetInkMaxStrokes(): number {
  return WET_INK_MAX_STROKES;
}
/** How often a batch held for a hovering pen looks again (ms). */
export const WET_INK_HOVER_RETRY_MS = 90;












export function inkLandsAtLift(conversionMode?: string | null): boolean {
  void conversionMode;
  return false;
}

/**
 * Whether wet ink should wait a moment more: the pen hovers over the page
 * between strokes, `heldMs` after the last lift (lib/ink-pointer-policy).
 */
export function wetInkWaitsForPen(heldMs: number): boolean {
  if (!inkPointerArbiter.penHovering()) return false;
  // A batch landed under a hovering pen meets the next contact
  // (inkLandsAtLift): it waits for the pen to leave, however long, bounded by
  // the stroke count alone (wetInkMaxStrokes). `heldMs` is kept for callers.
  void heldMs;
  return true;
}

export interface WetInkBatcherOptions<Stroke> {
  /** True while a stroke is being drawn: the batch waits for the lift. */
  readonly penDown: () => boolean;
  /** Takes the batch into the document. */
  readonly onFlush: (strokes: readonly Stroke[]) => void;
  /**
   * The host's view of the wet strokes (to draw them), after each change,
   * and what changed: a stroke `add`ed, the batch `land`ing (`onFlush` runs
   * next, so the host's render that takes the landed strokes is the one to
   * let the wet ones go: lib/office-wet-ink-layer), strokes scratched out
   * (`remove`) or dropped (`discard`).
   */
  readonly onChange?: (strokes: readonly Stroke[], change: WetInkChange) => void;
  readonly settleMs?: number;
  readonly maxStrokes?: number;
  /** True when a stroke lands at its lift instead of with the pause ({@link inkLandsAtLift}). */
  readonly landAtLift?: () => boolean;
}

/** What changed in a batch of wet ink ({@link WetInkBatcherOptions.onChange}). */
export type WetInkChange = "add" | "land" | "remove" | "discard";

export class WetInkBatcher<Stroke> {
  private strokes: Stroke[] = [];
  private timer: ReturnType<typeof setTimeout> | null = null;
  private options: WetInkBatcherOptions<Stroke>;
  private releaseHoverWatch: () => void;

  private readonly onFlushEvent = () => this.flush();

  constructor(options: WetInkBatcherOptions<Stroke>) {
    this.options = options;
    this.releaseHoverWatch = watchPenHover();
    if (typeof window !== "undefined") window.addEventListener(OFFICE_WET_INK_FLUSH_EVENT, this.onFlushEvent);
  }

  /** The host went away: land what is wet and stop listening. */
  dispose(): void {
    this.flush();
    this.releaseHoverWatch();
    if (typeof window !== "undefined") window.removeEventListener(OFFICE_WET_INK_FLUSH_EVENT, this.onFlushEvent);
  }

  /** The host's handlers may change between renders; the newest are used. */
  update(options: Partial<WetInkBatcherOptions<Stroke>>): void {
    this.options = { ...this.options, ...options };
  }

  get pending(): readonly Stroke[] {
    return this.strokes;
  }

  /**
   * A stroke the pen just finished: wet until the batch lands. One that
   * `waits` (a circle offering to select, Pen Gestures ▸ Circle to Lasso)
   * holds the batch until `resume`; a landing for any other reason still
   * takes it.
   */
  add(stroke: Stroke, options?: { readonly waits?: boolean }): void {
    this.strokes = [...this.strokes, stroke];
    this.options.onChange?.(this.strokes, "add");
    if (options?.waits) {
      if (this.timer !== null) clearTimeout(this.timer);
      this.timer = null;
      return;
    }
    this.resume();
  }

  /** The batch goes on to land, at the lift's or the pause's moment. */
  resume(): void {
    if (this.strokes.length === 0) return;
    if (this.options.landAtLift?.()) {
      // In the task after the lift, not inside it (see DocumentInkSurface:
      // WebKit drops a Pencil contact that lands while the page still
      // answers the lift). A pen down again by then waits for its own lift.
      if (this.timer !== null) clearTimeout(this.timer);
      this.timer = setTimeout(() => {
        this.timer = null;
        if (this.options.penDown()) return;
        this.flush();
      }, 0);
      return;
    }
    this.schedule();
  }

  private schedule(): void {
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
    if (this.strokes.length >= (this.options.maxStrokes ?? wetInkMaxStrokes())) {
      this.flush();
      return;
    }
    const settleMs = this.options.settleMs ?? wetInkSettleMs();
    const land = (heldMs: number) => {
      this.timer = null;
      // The pen is down again: the document's work for the batch would hold
      // this stroke's samples back. The lift of this stroke schedules again.
      if (this.options.penDown()) return;
      // The pen hovers just above the page: the next stroke is coming.
      if (wetInkWaitsForPen(heldMs)) {
        this.timer = setTimeout(() => land(heldMs + WET_INK_HOVER_RETRY_MS), WET_INK_HOVER_RETRY_MS);
        return;
      }
      this.flush();
    };
    this.timer = setTimeout(() => land(settleMs), settleMs);
  }

  /** Land the wet strokes now, in one batch. */
  flush(): void {
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
    const strokes = this.strokes;
    if (strokes.length === 0) return;
    this.strokes = [];
    this.options.onChange?.(this.strokes, "land");
    this.options.onFlush(strokes);
  }

  /**
   * Drop the wet strokes `drop` picks, without landing them (a scribble
   * scratched them out before they landed). How many went.
   */
  remove(drop: (stroke: Stroke) => boolean): number {
    const kept = this.strokes.filter((stroke) => !drop(stroke));
    const removed = this.strokes.length - kept.length;
    if (removed === 0) return 0;
    this.strokes = kept;
    if (kept.length === 0 && this.timer !== null) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    this.options.onChange?.(this.strokes, "remove");
    return removed;
  }

  /** Drop the wet strokes without landing them (the gesture was cancelled by the host). */
  discard(): void {
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
    if (this.strokes.length === 0) return;
    this.strokes = [];
    this.options.onChange?.(this.strokes, "discard");
  }
}
