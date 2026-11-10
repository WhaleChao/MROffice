/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One stroke, one pointer: the pointer policy every ink surface shares
 * (Docs, the sheet, the deck and the PDF viewer).
 *
 * On a tablet a stroke is seldom the only contact. The hand holding the pen
 * rests on the glass (a palm touch lands before, during or after the pen),
 * a second finger brushes the surface, the browser takes a touch for a
 * gesture and reports `pointercancel`. Each surface used to answer these on
 * its own — a second touch dropped the stroke, a palm's pointer moves were
 * mixed into the pen's points, a palm lifting committed the stroke early, a
 * cancel threw away everything drawn — and on a phone or an iPad that read
 * as the pen "cutting out" every few strokes.
 *
 * The rules, as Office on a tablet behaves:
 * - A stroke belongs to the pointer that started it. Every other pointer's
 *   down, move and up are ignored while it runs (palm rejection); only the
 *   owner's move adds points, only the owner's up ends it.
 * - A pen outranks a finger. A pen landing while a finger inks takes the
 *   stroke over (the finger was the palm, down first); a finger landing
 *   while a pen inks is ignored; a finger landing shortly after a pen was
 *   seen (in contact or hovering) is ignored too, the hand is still there.
 * - Two fingers landing together are a pinch, not a palm: the first finger's
 *   stroke (a few points old at most) is dropped and the editor's pinch zoom
 *   takes both. A second finger landing later does not interrupt.
 * - A `pointercancel` for the owner is the end of the stroke, not its loss:
 *   what was drawn is kept (the pen lifted as far as the user is concerned).
 * - The pointer's coalesced samples are all used, so a fast stroke keeps
 *   its curve instead of the one sample per frame the event carries.
 */

/** A finger landing this soon after a pen was seen is the hand, not a stroke. */
export const INK_PALM_REJECTION_MS = 700;
/** Two fingers landing this close together are a pinch. */
export const INK_PINCH_PAIR_MS = 150;






export const INK_PINCH_LATE_MS = 350;
export const INK_PINCH_LATE_TRAVEL_PX = 24;
/** A touch this wide AND tall (CSS px) is a resting hand, not a fingertip —
 * GoodNotes' palm rejection grades contacts by size the same way. One long
 * side alone is a fingertip with its nail, unless it is as long as
 * {@link INK_PALM_EDGE_PX} (the side of a hand). */
export const INK_PALM_CONTACT_PX = 48;
export const INK_PALM_EDGE_PX = 100;
/**
 * A running stroke whose pointer has not been heard from for this long has
 * lifted where no surface saw it (iPadOS drops the release when the page
 * changes under the pen). The stroke itself ends with the next press
 * ("replace"), but meanwhile nothing waits on it: a conversion that waits
 * for the hand, a save held while writing.
 */
export const INK_STROKE_SILENCE_MS = 1500;
/**
 * A pen heard in the air (moving, no tip down) this lately is still over the
 * page: the hand is between two strokes of its writing, not done. An Apple
 * Pencil reports its hover a centimetre above the glass.
 */
export const INK_PEN_HOVER_MS = 110;

function performanceClock(): number {
  return typeof performance === "undefined" ? Date.now() : performance.now();
}

export interface InkPointerLike {
  pointerId: number;
  pointerType: string;
  timeStamp: number;
  /** The contact's size (PointerEvent.width/height), when the device reports it. */
  width?: number;
  height?: number;
  /** Buttons held (0: a pen in the air, hovering), when known. */
  buttons?: number;
  /** Where it went down (client pixels), when known. */
  clientX?: number;
  clientY?: number;
}

/** A touch whose contact is the size of a hand, whatever else is going on. */
export function isPalmSizedTouch(event: InkPointerLike): boolean {
  const width = event.width ?? 0;
  const height = event.height ?? 0;
  return (
    event.pointerType === "touch" &&
    ((width >= INK_PALM_CONTACT_PX && height >= INK_PALM_CONTACT_PX) ||
      Math.max(width, height) >= INK_PALM_EDGE_PX)
  );
}

/** What a surface does with a pointer that just went down. */
export type InkPointerVerdict =
  /** No stroke runs; this pointer starts one. */
  | "start"
  /** Not this pointer's stroke (a palm, a second finger, a finger near the pen): nothing happens. */
  | "ignore"
  /** The pen landed while a finger (the palm) inks: drop that stroke and start the pen's. */
  | "replace"
  /** A second finger landed with the first: drop the stroke and let the pinch run. */
  | "pinch";

export class InkPointerArbiter {
  private penSeenAt = Number.NEGATIVE_INFINITY;
  /** When a pen was last heard hovering (performance time). */
  private penHoverAt = Number.NEGATIVE_INFINITY;
  private active: { id: number; type: string; downAt: number } | null = null;
  private beganAt: number | null = null;
  private endedAt: number | null = null;
  private startPoint: { readonly x: number; readonly y: number } | null = null;
  /** The latest strokes' starts (client pixels, event time), newest last. */
  private starts: { readonly x: number; readonly y: number; readonly at: number }[] = [];
  /** When the running stroke's pointer was last heard from (performance time). */
  private activeHeardAt: number | null = null;
  /** How far the running stroke's pointer has travelled from its start (client px). */
  private activeTravel = 0;

  /** Every pointer event a surface sees (down, move, hover): a pen in range leaves its time. */
  observe(event: InkPointerLike): void {
    if (event.pointerType === "pen") {
      this.penSeenAt = Math.max(this.penSeenAt, event.timeStamp);
      if (event.buttons === 0 && !(this.active && event.pointerId === this.active.id)) {
        this.penHoverAt = performanceClock();
      }
    }
    if (this.active && event.pointerId === this.active.id) {
      this.activeHeardAt = performanceClock();
      if (
        this.startPoint &&
        typeof event.clientX === "number" &&
        typeof event.clientY === "number" &&
        Number.isFinite(event.clientX) &&
        Number.isFinite(event.clientY)
      ) {
        this.activeTravel = Math.max(
          this.activeTravel,
          Math.hypot(event.clientX - this.startPoint.x, event.clientY - this.startPoint.y),
        );
      }
    }
  }

  /**
   * A stroke is being drawn: one began and has not ended, and its pointer
   * was heard from within {@link INK_STROKE_SILENCE_MS}. The hand that waits
   * on a stroke (a conversion, a save) asks this, not `activePointerId`.
   */
  strokeRunning(now: number = performanceClock()): boolean {
    if (!this.active) return false;
    return this.activeHeardAt === null || now - this.activeHeardAt < INK_STROKE_SILENCE_MS;
  }

  /**
   * The pen is in the air over the page between strokes (see
   * {@link INK_PEN_HOVER_MS}): wet ink waits rather than landing in the gap
   * the next stroke is about to start in.
   */
  penHovering(now: number = performanceClock()): boolean {
    return !this.active && now - this.penHoverAt < INK_PEN_HOVER_MS;
  }

  verdict(event: InkPointerLike): InkPointerVerdict {
    this.observe(event);
    if (isPalmSizedTouch(event)) return "ignore";
    const active = this.active;
    if (!active) {
      if (event.pointerType === "touch" && event.timeStamp - this.penSeenAt < INK_PALM_REJECTION_MS) {
        return "ignore";
      }
      return "start";
    }
    // A pointer goes down only after it went up, and one pen or mouse is
    // never down twice: a press from the running stroke's own pointer, or
    // from another of its kind, means its lift never reached the surface
    // (the page changed under it). That stroke is over; ignoring the new
    // press left the pen dead until the tool was picked again.
    if (event.pointerId === active.id) return "replace";
    if (event.pointerType !== "touch" && event.pointerType === active.type) return "replace";
    if (event.pointerType === "pen" && active.type === "touch") return "replace";
    if (event.pointerType === "touch" && active.type === "touch") {
      const age = event.timeStamp - active.downAt;
      if (age < INK_PINCH_PAIR_MS) return "pinch";
      if (age < INK_PINCH_LATE_MS && this.activeTravel < INK_PINCH_LATE_TRAVEL_PX) return "pinch";
    }
    return "ignore";
  }

  begin(event: InkPointerLike): void {
    beginInkStreamline(event.pointerId, event.clientX as number, event.clientY as number);
    this.active = { id: event.pointerId, type: event.pointerType, downAt: event.timeStamp };
    this.activeHeardAt = performanceClock();
    this.activeTravel = 0;
    this.beganAt = event.timeStamp;
    this.startPoint =
      typeof event.clientX === "number" && typeof event.clientY === "number" &&
      Number.isFinite(event.clientX) && Number.isFinite(event.clientY)
        ? { x: event.clientX, y: event.clientY }
        : null;
    if (this.startPoint) {
      this.starts.push({ ...this.startPoint, at: event.timeStamp });
      if (this.starts.length > 64) this.starts.splice(0, this.starts.length - 64);
    }
  }

  /**
   * Where the strokes begun after `at` (event time) began: strokes a surface
   * may still hold as wet ink, not yet in its document.
   */
  strokeStartsSince(at: number): { readonly x: number; readonly y: number }[] {
    return this.starts.filter((start) => start.at > at).map(({ x, y }) => ({ x, y }));
  }

  /**
   * Where the latest stroke began (client pixels at its pointerdown), running
   * or ended: where the hand is writing (Ink to Text reads a finished line
   * while the pen writes the next one). Null when unknown.
   */
  get lastStrokeStartPoint(): { readonly x: number; readonly y: number } | null {
    return this.startPoint;
  }

  /** When the latest stroke began (its pointerdown's time), running or ended. */
  get lastStrokeBeganAt(): number | null {
    return this.beganAt;
  }

  /** When the latest stroke ended (lifted, cancelled or dropped); null before any did. */
  get lastStrokeEndedAt(): number | null {
    return this.endedAt;
  }

  /** True when the event is the running stroke's own pointer. */
  owns(event: { pointerId: number }): boolean {
    return this.active !== null && this.active.id === event.pointerId;
  }

  get activePointerId(): number | null {
    return this.active?.id ?? null;
  }

  /** Forget the pen and the stroke (tests). */
  reset(): void {
    this.penSeenAt = Number.NEGATIVE_INFINITY;
    this.penHoverAt = Number.NEGATIVE_INFINITY;
    this.active = null;
    this.beganAt = null;
    this.endedAt = null;
    this.startPoint = null;
    this.starts = [];
    this.activeHeardAt = null;
  }

  /** The stroke ended (lifted, cancelled or dropped). */
  end(at: number = typeof performance === "undefined" ? 0 : performance.now()): void {
    const active = this.active;
    if (active) endInkStreamline(active.id);
    this.active = null;
    this.activeHeardAt = null;
    if (active) this.endedAt = at;
    // The hand stays on the glass a moment after the pen lifts.
    if (active?.type === "pen") this.penSeenAt = Math.max(this.penSeenAt, at);
  }
}

/** The one arbiter: one stroke runs in the product at a time, and a pen
 * seen on one surface is the same hand over the next. */
export const inkPointerArbiter = new InkPointerArbiter();

let penHoverWatchers = 0;
const onWindowPenMove = (event: PointerEvent) => {
  if (event.pointerType === "pen" && event.buttons === 0) inkPointerArbiter.observe(event);
};

/**
 * Hear a pen hovering anywhere in the window while a surface keeps wet ink
 * (the surfaces themselves listen only while a stroke runs). Returns the
 * release; the listener goes when the last surface releases it.
 */
export function watchPenHover(): () => void {
  if (typeof window === "undefined") return () => {};
  if (penHoverWatchers++ === 0) {
    window.addEventListener("pointermove", onWindowPenMove, { capture: true, passive: true });
  }
  let released = false;
  return () => {
    if (released) return;
    released = true;
    if (--penHoverWatchers === 0) window.removeEventListener("pointermove", onWindowPenMove, { capture: true });
  };
}

/** One sample of a pointer: where, and (for a pen) how hard, when. */
export interface InkPointerSample {
  readonly clientX: number;
  readonly clientY: number;
  /** PointerEvent.pressure: 0..1, 0.5 for a mouse button; undefined when unknown. */
  readonly pressure?: number;
  readonly pointerType?: string;
  readonly timeStamp?: number;
}

/** Each pointer's latest samples taken (time and place), newest last. */
const samplesTaken = new Map<number, string[]>();

/**
 * Streamline (Excalidraw's word for it): each sample is moved this share of
 * the way back toward the one before it, which takes the pointer's jitter
 * out of the line while keeping its shape. Excalidraw uses 0.2 for a pen
 * ("precise") and 0.5 for a mouse; a mouse here draws its samples as they
 * come, a finger (the jerkiest input) a little more smoothed than a pen.
 */
const INK_STREAMLINE: Readonly<Record<string, number>> = { pen: 0.2, touch: 0.35 };
/** The last streamlined point of each pointer's running stroke. */
const streamlinedLast = new Map<number, { x: number; y: number }>();

/** A stroke begins at `x`,`y`: its samples streamline from here. */
export function beginInkStreamline(pointerId: number, x: number, y: number): void {
  if (Number.isFinite(x) && Number.isFinite(y)) streamlinedLast.set(pointerId, { x, y });
  else streamlinedLast.delete(pointerId);
}

/** The stroke ended: nothing to streamline from any more. */
export function endInkStreamline(pointerId: number): void {
  streamlinedLast.delete(pointerId);
}

function streamlined(
  pointerId: number,
  samples: ReadonlyArray<InkPointerSample>,
  commit: boolean,
): ReadonlyArray<InkPointerSample> {
  const from = streamlinedLast.get(pointerId);
  if (!from || samples.length === 0) return samples;
  const share = INK_STREAMLINE[samples[0].pointerType ?? ""];
  if (!share) return samples;
  let last: { x: number; y: number } = from;
  const out: InkPointerSample[] = [];
  for (const sample of samples) {
    const next: { x: number; y: number } = {
      x: last.x + (sample.clientX - last.x) * (1 - share),
      y: last.y + (sample.clientY - last.y) * (1 - share),
    };
    out.push({
      clientX: next.x,
      clientY: next.y,
      pressure: sample.pressure,
      pointerType: sample.pointerType,
      timeStamp: sample.timeStamp,
    });
    last = next;
  }
  if (commit) streamlinedLast.set(pointerId, last);
  return out;
}












export function inkPointerSamples(
  event: PointerEvent | { nativeEvent?: PointerEvent; clientX: number; clientY: number },
): ReadonlyArray<InkPointerSample> {
  const native = "nativeEvent" in event && event.nativeEvent ? event.nativeEvent : (event as PointerEvent);
  const coalesced =
    typeof (native as PointerEvent).getCoalescedEvents === "function"
      ? (native as PointerEvent).getCoalescedEvents()
      : [];
  const samples: ReadonlyArray<InkPointerSample> = coalesced.length > 0 ? coalesced : [native];
  const pointerId = (native as Partial<PointerEvent>).pointerId;
  if (typeof pointerId !== "number") return samples;
  if (samples.length < 2) return streamlined(pointerId, samples, true);
  const taken = samplesTaken.get(pointerId) ?? [];
  const fresh = samples.filter((sample) => {
    if (!(typeof sample.timeStamp === "number" && sample.timeStamp > 0)) return true;
    const key = `${sample.timeStamp}:${sample.clientX}:${sample.clientY}`;
    if (taken.includes(key)) return false;
    taken.push(key);
    return true;
  });
  if (taken.length > 32) taken.splice(0, taken.length - 32);
  samplesTaken.set(pointerId, taken);
  return streamlined(pointerId, fresh, true);
}

/**
 * The samples the browser expects next (PointerEvent.getPredictedEvents):
 * where the pen will be a few milliseconds on, from its speed and curve.
 * Drawn ahead of the tip they hide the frame between the pen and its ink;
 * none where the platform has no guess.
 */
export function inkPredictedSamples(
  event: PointerEvent | { nativeEvent?: PointerEvent; clientX: number; clientY: number },
): ReadonlyArray<InkPointerSample> {
  const native = "nativeEvent" in event && event.nativeEvent ? event.nativeEvent : (event as PointerEvent);
  if (typeof (native as PointerEvent).getPredictedEvents !== "function") return [];
  try {
    const predicted = (native as PointerEvent).getPredictedEvents();
    const pointerId = (native as Partial<PointerEvent>).pointerId;
    // The prediction continues from the last streamlined sample, without
    // moving it: the next real samples take over from there.
    return typeof pointerId === "number" ? streamlined(pointerId, predicted, false) : predicted;
  } catch {
    return [];
  }
}

/**
 * The pressure a sample's ink is drawn at: a pen's own reading (0..1), or
 * none. A mouse reports 0.5 whenever its button is down and a finger on most
 * glass reports nothing meaningful (0 on WebKit), so only a pen varies the
 * width; the width factor of an unknown pressure is 1 (lib/office-ink-smooth).
 */
export function inkSamplePressure(sample: InkPointerSample): number | undefined {
  if (sample.pointerType !== "pen") return undefined;
  const pressure = sample.pressure;
  if (pressure === undefined || !Number.isFinite(pressure) || pressure <= 0) return undefined;
  return Math.min(1, pressure);
}
