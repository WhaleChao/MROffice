/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pxToPt } from "./office-units";
import { cssFontFamilyList } from "./css-font-family";
/**
 * Draw ▸ Convert, the modes: Ink to Shape, Ink to Math and Ink to Text as
 * the Docs and deck Draw tabs run them (Microsoft Support, "Change
 * handwritten ink to shapes, text, or math in PowerPoint for Microsoft 365";
 * "Explore the Ink to Text Pen"). One conversion mode is on at a time. With
 * Ink to Math or Ink to Text on, the strokes are drawn as ink as usual; when
 * the writer pauses, the strokes written since the last conversion are read
 * as ONE piece of writing and replaced by what they say, where they were.
 *
 * This module is the editors' shared, DOM-free part: the mode toggle, the
 * pause grouping, what counts as a usable reading, and the placement rules.
 * The recogniser is `lib/office-ink-recognition`; each editor places the
 * result itself (`documentInk.ts`, the deck window).
 */

export type OfficeInkConversionMode = "shape" | "math" | "text";
export type OfficeInkRecognitionMode = Extract<OfficeInkConversionMode, "math" | "text">;

/** The pressed mode turns on (and every other one off), or off when it was on. */
export function toggledOfficeInkConversionMode(
  current: OfficeInkConversionMode | null,
  pressed: OfficeInkConversionMode,
): OfficeInkConversionMode | null {
  return current === pressed ? null : pressed;
}

export function isOfficeInkRecognitionMode(
  mode: OfficeInkConversionMode | null | undefined,
): mode is OfficeInkRecognitionMode {
  return mode === "math" || mode === "text";
}

/**
 * What one editor offers a Convert button over the ink selected in it
 * (`T` is how the editor names a stroke).
 */
export interface OfficeInkSelectionConverter<T> {
  /**
   * The ink strokes the selection holds now: those Lasso Select took, or the
   * ink picked. Never a shape, picture or text: Office's lasso cannot take
   * one, and one picked with the ink stays as it is.
   */
  readonly selectedInk: () => readonly T[];
  /** Ink to Shape: each stroke that plainly is a shape becomes it, the rest stay ink; one Undo step. */
  readonly toShapes: (strokes: readonly T[]) => void;
  /** Ink to Math / Ink to Text: the strokes are read as ONE piece of writing and replaced by it; one Undo step. */
  readonly read: (strokes: readonly T[], mode: OfficeInkRecognitionMode) => void;
  /** No ink selected: the mode in which new writing converts turns on or off. */
  readonly toggleMode: (mode: OfficeInkConversionMode) => void;
}











export function pressOfficeInkConvert<T>(
  mode: OfficeInkConversionMode,
  converter: OfficeInkSelectionConverter<T>,
): "selection" | "mode" {
  const strokes = converter.selectedInk();
  if (strokes.length === 0) {
    converter.toggleMode(mode);
    return "mode";
  }
  if (mode === "shape") converter.toShapes(strokes);
  else converter.read(strokes, mode);
  return "selection";
}

/**
 * How long the pen rests before the strokes written so far are read. A
 * stroke that starts inside the pause joins the same piece of writing.
 */
export const OFFICE_INK_CONVERSION_PAUSE_MS = 800;


















export const OFFICE_INK_MATH_PAUSE_MS = 1200;
/**
 * The pause follows the writer. Someone who stopped for a while inside this
 * piece of writing and went on stops like that again: the pause grows to
 * this many times the longest stop so far, up to the mode's ceiling. A fixed
 * pause read a long equation in the middle of a thought.
 */
export const OFFICE_INK_PAUSE_GAP_FACTOR = 1.6;
export const OFFICE_INK_TEXT_PAUSE_MAX_MS = 2000;
export const OFFICE_INK_MATH_PAUSE_MAX_MS = 3000;
/**
 * A reading that turned out early (the writer went on next to the equation
 * it made) lengthens the pause for the rest of the session by this much.
 */
export const OFFICE_INK_PAUSE_LEARNING_FACTOR = 1.5;
/**
 * Writing that BEGINS within this long after an equation was made, on its
 * line and right after it, goes on with that equation: the whole writing is
 * read again. However long the continuation then takes to write does not
 * count (the window used to close on the writer mid-equation).
 */
export const OFFICE_INK_MATH_CONTINUE_MS = 20_000;

/**
 * The most other readings a converted piece of writing offers to correct
 * it with (OneNote's alternates; PowerPoint's More Suggestions).
 */
export const OFFICE_INK_MAX_ALTERNATIVES = 5;

/** Most strokes one reading of text takes (a paragraph of handwriting is far fewer). */
export const OFFICE_INK_CONVERSION_MAX_STROKES = 400;
/**
 * An equation is never read because it grew long: a long equation is many
 * strokes, and reading it at a count cut it in two. This is only the bound
 * past which a piece of writing is plainly not one equation.
 */
export const OFFICE_INK_MATH_MAX_STROKES = 1500;

interface OfficeInkTimers {
  readonly setTimeout: (callback: () => void, ms: number) => unknown;
  readonly clearTimeout: (handle: unknown) => void;
  /** The clock the strokes' stops are measured on (ms). */
  readonly now?: () => number;
}

export interface OfficeInkStrokeGrouperOptions<T> {
  /** The strokes written since the last pause, oldest first, and the pause that ended them. */
  readonly onPause: (group: readonly T[], pauseMs: number) => void;
  /** The pause, or a function of the moment (the mode may change). */
  readonly pauseMs?: number | (() => number);
  /**
   * How far the pause may grow after the writer's own stops
   * ({@link OFFICE_INK_PAUSE_GAP_FACTOR}). Without it the pause is `pauseMs`.
   */
  readonly maxPauseMs?: number | (() => number);
  /**
   * When the stroke being committed began, on the timers' clock: a stop is
   * measured from one stroke's end to the next one's start. Without it, to
   * the next one's end.
   */
  readonly strokeStartedAt?: () => number | null;
  /** Most strokes one piece takes before it is read as it stands. */
  readonly maxStrokes?: number | (() => number);
  /**
   * A stroke is being drawn right now: the pause is not over (the stroke
   * will join the group when it lands). Read when the pause would end.
   */
  readonly isBusy?: () => boolean;
  readonly timers?: OfficeInkTimers;
  /**
   * Writing in several places at once (Ink to Text, 「边写边转」): each place
   * is a piece of its own with its own pause, and a piece whose pause ran out
   * is read while the pen writes somewhere else. Off (or a function of the
   * moment that says no), every stroke joins the one piece, as before.
   */
  readonly regional?: boolean | (() => boolean);
  /** Where strokes are now (any one space); null when they cannot be measured. Needed for `regional`. */
  readonly boxOf?: (strokes: readonly T[]) => OfficeInkBounds | null;
  /** Whether writing at `next` goes on with the piece at `piece` (default: {@link officeInkSameLine}). */
  readonly together?: (piece: OfficeInkBounds, next: OfficeInkBounds) => boolean;
  /**
   * Where the strokes begun after `since` (on the timers' clock) began, in
   * `boxOf`'s space: the pen's latest writing, some of it perhaps not landed
   * yet. A piece is handed over while the pen is busy only when none of them
   * goes on with it.
   */
  readonly writingSince?: (since: number) => readonly OfficeInkBounds[];
}

export interface OfficeInkStrokeGrouper<T> {
  /** A stroke landed: it joins the pending group and the pause restarts. */
  strokeCommitted(stroke: T): void;
  /**
   * Strokes handed over at a pause come back: the writer went on before
   * their reading landed. They lead the pending group, as they were written
   * first.
   */
  adopt(strokes: readonly T[]): void;
  /** Drop the pending group without reading it (mode off, slide changed). */
  cancel(): void;
  /** The strokes waiting for the pause, oldest first. */
  pending(): readonly T[];
  /** The pause as it stands for the pending group. */
  pause(): number;
  /** A pending stroke was rewritten (the editor moved it back under the pen): `to` stands for `from`. */
  replace(from: T, to: T): void;
}

const DEFAULT_TIMERS: OfficeInkTimers = {
  setTimeout: (callback, ms) => globalThis.setTimeout(callback, ms),
  clearTimeout: (handle) => globalThis.clearTimeout(handle as ReturnType<typeof setTimeout>),
};

function resolved(option: number | (() => number) | undefined): number | undefined {
  return typeof option === "function" ? option() : option;
}

/**
 * Writing at `next` goes on with the line of writing at `line` (Ink to
 * Text's pieces). It is another line when it is clearly apart: far to either
 * side, well above or below, or below the line and back to its left, where
 * the pen goes to begin the next line. Below the line's END it goes on with
 * it: the lower strokes of the character being written (六's two dots under
 * its bar, a descender) are there. Sizes are measured in the larger of the
 * two heights, at least 24 px, so a first small stroke (a dot) does not make
 * the next one a stranger.
 */
export function officeInkSameLine(line: OfficeInkBounds, next: OfficeInkBounds): boolean {
  const height = Math.max(24, line.bottom - line.top, next.bottom - next.top);
  if (next.left > line.right + height * 3 || next.right < line.left - height * 3) return false;
  if (next.top > line.bottom + height * 1.2 || next.bottom < line.top - height * 1.2) return false;
  const middle = (next.top + next.bottom) / 2;
  const backToTheLeft = next.right < line.right - height * 1.5;
  return !(middle > line.bottom && backToTheLeft);
}

interface InkPiece<T> {
  strokes: T[];
  timer: unknown;
  /** When its last stroke landed, and its writer's longest stop. */
  landedAt: number | null;
  longestStop: number;
}

/**
 * Groups strokes into pieces of writing by the pen's pauses: each landed
 * stroke restarts the pause; when it runs out with no stroke in progress the
 * group is handed over and a new one begins. A stroke still being drawn when
 * the pause ends holds the group open, so it extends it. The pause is
 * counted from the last stroke, never from the first: a piece of writing has
 * no length at which it is read.
 *
 * Regional (`regional`): a stroke joins the piece it is written with (the
 * same line); a stroke elsewhere begins its own. A piece whose pause ran out
 * is handed over even while the pen is busy, as long as the pen is busy
 * somewhere else: that line is finished.
 */
export function createOfficeInkStrokeGrouper<T>(
  options: OfficeInkStrokeGrouperOptions<T>,
): OfficeInkStrokeGrouper<T> {
  const timers = options.timers ?? DEFAULT_TIMERS;
  const clock = timers.now ?? (() => (typeof performance === "undefined" ? Date.now() : performance.now()));
  const together = options.together ?? officeInkSameLine;
  let pieces: InkPiece<T>[] = [];

  const regional = () => {
    const option = options.regional;
    return Boolean(typeof option === "function" ? option() : option) && typeof options.boxOf === "function";
  };
  const boxOf = (strokes: readonly T[]): OfficeInkBounds | null => {
    if (!options.boxOf || strokes.length === 0) return null;
    try {
      return options.boxOf(strokes);
    } catch {
      return null;
    }
  };
  const pauseOf = (piece: InkPiece<T> | undefined) => {
    const base = resolved(options.pauseMs) ?? OFFICE_INK_CONVERSION_PAUSE_MS;
    const ceiling = resolved(options.maxPauseMs);
    if (ceiling === undefined || !(ceiling > base)) return base;
    return Math.min(ceiling, Math.max(base, (piece?.longestStop ?? 0) * OFFICE_INK_PAUSE_GAP_FACTOR));
  };
  const stop = (piece: InkPiece<T>) => {
    if (piece.timer !== null) timers.clearTimeout(piece.timer);
    piece.timer = null;
  };
  const arm = (piece: InkPiece<T>) => {
    stop(piece);
    piece.timer = timers.setTimeout(() => fire(piece), pauseOf(piece));
  };
  const drop = (piece: InkPiece<T>) => {
    stop(piece);
    pieces = pieces.filter((candidate) => candidate !== piece);
  };
  /**
   * The pen is writing elsewhere than `piece`: it is finished. Every stroke
   * begun since the piece's last one landed is elsewhere (one begun beside
   * it may still be on its way into the document); with none begun since,
   * the latest stroke is.
   */
  const writingElsewhere = (piece: InkPiece<T>): boolean => {
    if (!regional() || piece.landedAt === null) return false;
    let starts: readonly OfficeInkBounds[] = [];
    try {
      starts = options.writingSince?.(piece.landedAt) ?? [];
      if (starts.length === 0) starts = (options.writingSince?.(Number.NEGATIVE_INFINITY) ?? []).slice(-1);
    } catch {
      starts = [];
    }
    const box = boxOf(piece.strokes);
    return Boolean(box && starts.length > 0 && !starts.some((start) => together(box, start)));
  };
  function fire(piece: InkPiece<T>) {
    piece.timer = null;
    if (!pieces.includes(piece)) return;
    if (piece.strokes.length === 0) {
      drop(piece);
      return;
    }
    if (options.isBusy?.() && !writingElsewhere(piece)) {
      arm(piece);
      return;
    }
    const ready = piece.strokes;
    const waited = pauseOf(piece);
    drop(piece);
    options.onPause(ready, waited);
  }
  /** The piece `stroke` (at `box`) joins: every piece it goes on with, merged; null for a new one. */
  const pieceFor = (box: OfficeInkBounds | null): InkPiece<T> | null => {
    if (!regional()) return pieces[0] ?? null;
    if (!box) return pieces[pieces.length - 1] ?? null;
    const joined = pieces.filter((piece) => {
      const at = boxOf(piece.strokes);
      return !at || together(at, box);
    });
    if (joined.length === 0) return null;
    const [first, ...rest] = joined;
    for (const other of rest) {
      first!.strokes.push(...other.strokes);
      first!.longestStop = Math.max(first!.longestStop, other.longestStop);
      first!.landedAt = Math.max(first!.landedAt ?? 0, other.landedAt ?? 0) || first!.landedAt;
      drop(other);
    }
    return first!;
  };

  return {
    strokeCommitted(stroke) {
      const now = clock();
      let piece = pieceFor(regional() ? boxOf([stroke]) : null);
      if (!piece) {
        piece = { strokes: [], timer: null, landedAt: null, longestStop: 0 };
        pieces.push(piece);
      }
      if (piece.strokes.length > 0 && piece.landedAt !== null) {
        const began = options.strokeStartedAt?.();
        const resumed =
          typeof began === "number" && Number.isFinite(began) && began >= piece.landedAt && began <= now ? began : now;
        piece.longestStop = Math.max(piece.longestStop, resumed - piece.landedAt);
      }
      piece.landedAt = now;
      piece.strokes.push(stroke);
      if (piece.strokes.length >= (resolved(options.maxStrokes) ?? OFFICE_INK_CONVERSION_MAX_STROKES)) {
        // A very long piece of writing is read as it stands.
        stop(piece);
        fire(piece);
        return;
      }
      arm(piece);
    },
    adopt(strokes) {
      if (strokes.length === 0) return;
      let piece = pieceFor(regional() ? boxOf(strokes) : null);
      if (!piece) {
        piece = { strokes: [], timer: null, landedAt: null, longestStop: 0 };
        pieces.unshift(piece);
      }
      piece.strokes = [...strokes, ...piece.strokes];
      if (piece.landedAt === null) piece.landedAt = clock();
      arm(piece);
    },
    cancel() {
      for (const piece of pieces) stop(piece);
      pieces = [];
    },
    pending() {
      return pieces.flatMap((piece) => piece.strokes);
    },
    pause() {
      return pauseOf(pieces[pieces.length - 1]);
    },
    replace(from, to) {
      for (const piece of pieces) {
        const at = piece.strokes.indexOf(from);
        if (at >= 0) piece.strokes[at] = to;
      }
    },
  };
}









export function usableOfficeInkReading(
  answer: { readonly value: string | null | undefined; readonly confidence?: number | null } | null,
): string | null {
  const value = answer?.value?.trim();
  return value ? value : null;
}

/** The recogniser's other readings worth offering: trimmed, none twice, never the reading itself. */
export function officeInkReadingAlternatives(
  value: string,
  alternatives: readonly (string | null | undefined)[] | null | undefined,
): string[] {
  const out: string[] = [];
  for (const raw of alternatives ?? []) {
    const candidate = typeof raw === "string" ? raw.trim() : "";
    if (candidate && candidate !== value && !out.includes(candidate)) out.push(candidate);
    if (out.length >= OFFICE_INK_MAX_ALTERNATIVES) break;
  }
  return out;
}

/** How far a pen may move and still tap (CSS px): a tap, not a dot of ink, where a tap means something. */
export const OFFICE_INK_TAP_SLOP_PX = 6;

/** Whether a stroke's samples (client pixels) are a tap: none strays from the first. */
export function isOfficeInkTap(points: readonly { readonly x: number; readonly y: number }[]): boolean {
  const first = points[0];
  if (!first) return false;
  return points.every((point) => Math.hypot(point.x - first.x, point.y - first.y) <= OFFICE_INK_TAP_SLOP_PX);
}

export interface OfficeInkBounds {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/** The box the strokes cover, in their own units; null with no samples. */
export function officeInkStrokeBounds(
  strokes: readonly (readonly { readonly x: number; readonly y: number }[])[],
): OfficeInkBounds | null {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const stroke of strokes) {
    for (const point of stroke) {
      if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
      left = Math.min(left, point.x);
      top = Math.min(top, point.y);
      right = Math.max(right, point.x);
      bottom = Math.max(bottom, point.y);
    }
  }
  return Number.isFinite(left) ? { left, top, right, bottom } : null;
}

/**
 * Writing that goes on with the equation just made: it sits on that
 * equation's line (its middle within half a line above or below it) and
 * starts no further left than the equation does and no further right than a
 * few symbols past its end. A line written under the equation is the next
 * equation, not this one's continuation: the looser "anywhere near it" rule
 * read a derivation's second line into its first.
 */
export function officeInkContinuesLine(equation: OfficeInkBounds, writing: OfficeInkBounds): boolean {
  const line = Math.max(24, equation.bottom - equation.top);
  const middle = (writing.top + writing.bottom) / 2;
  return (
    middle >= equation.top - line / 2 &&
    middle <= equation.bottom + line / 2 &&
    writing.left >= equation.left - line / 2 &&
    writing.left <= equation.right + line * 2.5
  );
}

/**
 * Where `box` is now, when the thing it sat beside moved from `then` to
 * `now` (the page scrolled or zoomed between two readings).
 */
export function officeInkBoundsMoved(
  box: OfficeInkBounds,
  then: OfficeInkBounds,
  now: OfficeInkBounds,
): OfficeInkBounds {
  const scaleX = then.right - then.left > 0 ? (now.right - now.left) / (then.right - then.left) : 1;
  const scaleY = then.bottom - then.top > 0 ? (now.bottom - now.top) / (then.bottom - then.top) : 1;
  const sx = Number.isFinite(scaleX) && scaleX > 0 ? scaleX : 1;
  const sy = Number.isFinite(scaleY) && scaleY > 0 ? scaleY : 1;
  return {
    left: now.left + (box.left - then.left) * sx,
    top: now.top + (box.top - then.top) * sy,
    right: now.left + (box.right - then.left) * sx,
    bottom: now.top + (box.bottom - then.top) * sy,
  };
}

/**
 * Where an equation's handwriting is taken to stand when the writer goes on
 * right after the typeset equation: at its own size (`scale` for a page that
 * zoomed since), ending where the equation ends, centred on its line. The
 * typeset equation is smaller than the hand that wrote it; squeezed into its
 * box, the earlier strokes read as small print beside the new ones.
 */
export function officeInkBoundsEndingAt(
  written: OfficeInkBounds,
  equation: OfficeInkBounds,
  scale = 1,
): OfficeInkBounds {
  const factor = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const width = (written.right - written.left) * factor;
  const height = (written.bottom - written.top) * factor;
  const middle = (equation.top + equation.bottom) / 2;
  return {
    left: equation.right - width,
    top: middle - height / 2,
    right: equation.right,
    bottom: middle + height / 2,
  };
}

/**
 * Strokes drawn in `from`, laid into `into`: the earlier part of an equation
 * is read again where the writer now sees it (the typeset equation, or the
 * place its ink was), so what is written after it lines up with it.
 */
export function officeInkStrokesFitted<P extends { readonly x: number; readonly y: number }>(
  strokes: readonly (readonly P[])[],
  from: OfficeInkBounds,
  into: OfficeInkBounds,
): { x: number; y: number }[][] {
  const width = Math.max(1e-6, from.right - from.left);
  const height = Math.max(1e-6, from.bottom - from.top);
  return strokes.map((stroke) =>
    stroke.map((point) => ({
      x: into.left + ((point.x - from.left) / width) * (into.right - into.left),
      y: into.top + ((point.y - from.top) / height) * (into.bottom - into.top),
    })),
  );
}

/**
 * Where a reading lands in a flowing document (Word):
 *  - `inline`: the strokes were written over one line of an existing
 *    paragraph's text, so the reading joins that line where the pen began
 *    (Word's Ink to Text Pen writes into the line; the Math Input Control
 *    inserts an inline equation in a paragraph that holds other text).
 *  - `fill`: over an empty paragraph, the reading becomes that paragraph
 *    (the Math Input Control's display equation in an empty paragraph).
 *  - `after`: anywhere else (taller than a line, or below the text), the
 *    reading is a paragraph of its own after the one the strokes sit over.
 */
export type OfficeInkFlowPlacement = "inline" | "fill" | "after";

/** Strokes up to this many line heights tall still read as one line of writing. */
export const OFFICE_INK_ONE_LINE_RATIO = 1.8;

export function officeInkFlowPlacement(input: {
  readonly strokes: Pick<OfficeInkBounds, "top" | "bottom">;
  /** The line under the strokes' start, on screen. */
  readonly line: Pick<OfficeInkBounds, "top" | "bottom"> | null;
  readonly paragraphEmpty: boolean;
}): OfficeInkFlowPlacement {
  if (input.paragraphEmpty) return "fill";
  const line = input.line;
  if (!line) return "after";
  const lineHeight = Math.max(1, line.bottom - line.top);
  const height = input.strokes.bottom - input.strokes.top;
  const middle = (input.strokes.top + input.strokes.bottom) / 2;
  // One line of writing over that line: as tall as about a line, its
  // middle within the line's own height around it.
  const overLine = middle >= line.top - lineHeight / 2 && middle <= line.bottom + lineHeight / 2;
  return height <= lineHeight * OFFICE_INK_ONE_LINE_RATIO && overLine ? "inline" : "after";
}

/**
 * How tall a line of `text` stands, set at 1 em: from the top of its tallest
 * letter to the bottom of its lowest, in em. A Chinese or Japanese character
 * fills most of its em box; a Latin capital, a digit or an ascender stands
 * about 0.72 em; a lowercase letter 0.5 em, and a descender reaches 0.22 em
 * below the line. Handwriting is measured the same way (the box of its
 * strokes), so the two compare: written "ace" is as tall as its x-height,
 * written "Apply" from its capital to its descender.
 */
export function officeInkTextEmExtent(text: string): number {
  let top = Number.NEGATIVE_INFINITY;
  let bottom = Number.POSITIVE_INFINITY;
  for (const character of text) {
    let extent: readonly [number, number] | null = null;
    if (/[\u2e80-\u2fff\u3040-\u30ff\u3100-\u31ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/.test(character)) {
      extent = [-0.1, 0.8];
    } else if (/[gjpqy]/.test(character)) {
      extent = [-0.22, 0.5];
    } else if (/[bdfhklt]/.test(character)) {
      extent = [0, 0.74];
    } else if (/[0-9]/.test(character) || /\p{Lu}/u.test(character)) {
      extent = [0, 0.72];
    } else if (/\p{Ll}/u.test(character)) {
      extent = [0, 0.5];
    } else if (/\p{L}/u.test(character)) {
      extent = [0, 0.72];
    }
    if (!extent) continue;
    bottom = Math.min(bottom, extent[0]);
    top = Math.max(top, extent[1]);
  }
  return Number.isFinite(top) && Number.isFinite(bottom) && top > bottom ? top - bottom : 0.72;
}

/**
 * The bands a piece of handwriting stands in, top to bottom: strokes whose
 * vertical spans overlap, directly or through others, are one band. A
 * line of writing is one band or a few (a character such as 是 or 三 has
 * gaps inside it, from one stroke to the next below); separate lines are
 * separate bands.
 */
export function officeInkWritingLines(
  strokes: readonly (readonly { readonly x: number; readonly y: number }[])[],
): OfficeInkBounds[] {
  const boxes = strokes
    .map((stroke) => officeInkStrokeBounds([stroke]))
    .filter((box): box is OfficeInkBounds => box !== null)
    .sort((a, b) => a.top - b.top);
  const bands: OfficeInkBounds[] = [];
  for (const box of boxes) {
    const last = bands[bands.length - 1];
    if (last && box.top <= last.bottom) {
      bands[bands.length - 1] = {
        left: Math.min(last.left, box.left),
        top: last.top,
        right: Math.max(last.right, box.right),
        bottom: Math.max(last.bottom, box.bottom),
      };
      continue;
    }
    bands.push({ ...box });
  }
  return bands;
}

/**
 * The handwriting's `count` lines, top to bottom: its bands
 * ({@link officeInkWritingLines}) joined across the narrowest gaps until
 * `count` are left -- the gaps inside a line's characters are narrower than
 * the space between two lines. Fewer bands than `count` stay as they are.
 */
export function officeInkWritingLinesOf(
  strokes: readonly (readonly { readonly x: number; readonly y: number }[])[],
  count: number,
): OfficeInkBounds[] {
  const lines = officeInkWritingLines(strokes);
  while (lines.length > Math.max(1, count)) {
    let at = 0;
    let narrowest = Number.POSITIVE_INFINITY;
    for (let index = 0; index + 1 < lines.length; index += 1) {
      const gap = lines[index + 1]!.top - lines[index]!.bottom;
      if (gap < narrowest) {
        narrowest = gap;
        at = index;
      }
    }
    const upper = lines[at]!;
    const lower = lines[at + 1]!;
    lines.splice(at, 2, {
      left: Math.min(upper.left, lower.left),
      top: Math.min(upper.top, lower.top),
      right: Math.max(upper.right, lower.right),
      bottom: Math.max(upper.bottom, lower.bottom),
    });
  }
  return lines;
}

/** Two bands are one equation's while the gap between them is at most this
 * many times the taller band's height (a fraction's bar and its terms). */
export const OFFICE_INK_EQUATION_BAND_GAP = 0.6;
/** Writing on one line further apart than this many times its height is two equations. */
export const OFFICE_INK_EQUATION_SIDE_GAP = 3;

/**
 * The separate equations in a selection of ink, in reading order (top to
 * bottom, then left to right), each as the indexes of its strokes. Ink
 * selected across the page (an equation on one line, another further down or
 * after the text) was read as one equation and placed as one; Office's Ink
 * Equation converts what is on one line as one equation, so each line here is
 * read and placed on its own, where it was written. Bands of writing
 * ({@link officeInkWritingLines}) closer than {@link OFFICE_INK_EQUATION_BAND_GAP}
 * of a band's height stay one equation (a fraction's terms and its bar, a
 * limit under a sum); a line's writing parted by more than
 * {@link OFFICE_INK_EQUATION_SIDE_GAP} of its height is two.
 */
export function officeInkSeparateEquations(
  strokes: readonly (readonly { readonly x: number; readonly y: number }[])[],
): number[][] {
  const boxes = strokes.map((stroke) => officeInkStrokeBounds([stroke]));
  const order = boxes
    .map((box, index) => ({ box, index }))
    .filter((entry): entry is { box: OfficeInkBounds; index: number } => entry.box !== null)
    .sort((a, b) => a.box.top - b.box.top);
  // Bands: strokes whose vertical spans overlap, directly or through others.
  const bands: Array<{ top: number; bottom: number; members: number[] }> = [];
  for (const { box, index } of order) {
    const last = bands[bands.length - 1];
    if (last && box.top <= last.bottom) {
      last.bottom = Math.max(last.bottom, box.bottom);
      last.members.push(index);
    } else {
      bands.push({ top: box.top, bottom: box.bottom, members: [index] });
    }
  }
  // Lines: bands joined across the gaps an equation has inside it.
  const lines: Array<{ bottom: number; height: number; members: number[] }> = [];
  for (const band of bands) {
    const height = band.bottom - band.top;
    const last = lines[lines.length - 1];
    if (last && band.top - last.bottom <= OFFICE_INK_EQUATION_BAND_GAP * Math.max(last.height, height)) {
      last.bottom = Math.max(last.bottom, band.bottom);
      last.height = Math.max(last.height, height);
      last.members.push(...band.members);
    } else {
      lines.push({ bottom: band.bottom, height, members: [...band.members] });
    }
  }
  const equations: number[][] = [];
  for (const line of lines) {
    const across = [...line.members].sort((a, b) => boxes[a]!.left - boxes[b]!.left);
    let right = Number.NEGATIVE_INFINITY;
    let current: number[] = [];
    for (const index of across) {
      const box = boxes[index]!;
      if (current.length > 0 && box.left - right > OFFICE_INK_EQUATION_SIDE_GAP * Math.max(1, line.height)) {
        equations.push(current);
        current = [];
      }
      current.push(index);
      right = Math.max(right, box.right);
    }
    if (current.length > 0) equations.push(current.sort((a, b) => a - b));
  }
  return equations.map((members) => members.sort((a, b) => a - b));
}

/** A glyph that reads one em wide: CJK, kana, hangul and full-width forms. */
const OFFICE_INK_WIDE_READING = /[\u2e80-\u2fff\u3000-\u30ff\u3100-\u31ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/u;
/** Glyphs with no height to measure by: a rule, a dash, a dot, an ellipsis. */
const OFFICE_INK_FLAT_READING = /^[\s\u4e00\u2014\u2013\u2015\u2010\-=_~\u2026\u00b7\u30fb.,\u3001\u3002\uff0c\uff0e\u30fc]+$/u;

/**
 * How wide a line stands in em once typed, ink to ink: a wide glyph advances
 * 1 em, a space 0.3, anything else about 0.55, and the last wide glyph's ink
 * stops short of its advance (about 0.9 em of ink).
 */
function officeInkReadingAdvanceEm(text: string): number {
  let ems = 0;
  for (const char of text) ems += OFFICE_INK_WIDE_READING.test(char) ? 1 : /\s/.test(char) ? 0.3 : 0.55;
  const last = [...text].pop();
  return last !== undefined && OFFICE_INK_WIDE_READING.test(last) ? ems - 0.1 : ems;
}















function officeInkLineEmPx(line: OfficeInkBounds, words: string): number {
  const byHeight = Math.max(1, line.bottom - line.top) / officeInkTextEmExtent(words);
  const glyphs = [...words].filter((char) => !/\s/.test(char));
  const advance = officeInkReadingAdvanceEm(words.trim());
  const width = line.right - line.left;
  if (glyphs.length === 0 || !(advance > 0) || !(width > 0)) return byHeight;
  const byWidth = width / advance;
  if (OFFICE_INK_FLAT_READING.test(words)) return byWidth;
  const wide = glyphs.filter((char) => OFFICE_INK_WIDE_READING.test(char)).length;
  if (wide * 2 < glyphs.length) return byHeight;
  return Math.max(Math.min(byHeight, byWidth), byHeight * 0.55);
}

/** The size in `sizes` (points) nearest `points`; `points` itself when the list is empty. */
export function officeInkNearestSize(points: number, sizes: readonly number[]): number {
  let best: number | null = null;
  for (const size of sizes) {
    if (!Number.isFinite(size) || size <= 0) continue;
    if (best === null || Math.abs(size - points) < Math.abs(best - points)) best = size;
  }
  return best ?? points;
}

/** The font size an equation is measured at to size it to its writing (px). */
const OFFICE_INK_EQUATION_PROBE_PX = 20;










export function officeInkEquationFontSizePt(
  written: { readonly width: number; readonly height: number },
  measureAt: (fontSizePx: number) => { readonly width: number; readonly height: number } | null,
  sizes: readonly number[],
): number | null {
  if (!(written.width > 0) || !(written.height > 0)) return null;
  const measured = measureAt(OFFICE_INK_EQUATION_PROBE_PX);
  if (!measured || !(measured.width > 0) || !(measured.height > 0)) return null;
  const px = OFFICE_INK_EQUATION_PROBE_PX * Math.min(written.height / measured.height, written.width / measured.width);
  if (!Number.isFinite(px) || px <= 0) return null;
  return officeInkNearestSize(pxToPt(px), sizes);
}











export function officeInkWrittenTextSizes(
  strokes: readonly (readonly { readonly x: number; readonly y: number }[])[],
  text: string,
  sizes: readonly number[],
): number[] {
  const textLines = text.split("\n");
  const said = textLines.map((line, index) => ({ line, index })).filter(({ line }) => line.trim() !== "");
  // The handwriting in as many lines as the reading has: a line of Chinese
  // stood in several bands (a character's strokes leave gaps between them),
  // each measured as a line of its own -- the middle of them, a bar or a
  
  const written = officeInkWritingLinesOf(strokes, said.length);
  const fallback = sizes.length ? officeInkNearestSize(11, sizes) : 11;
  if (written.length === 0) return textLines.map(() => fallback);
  const sizeOf = (line: OfficeInkBounds, words: string, paired: boolean) => {
    // A written line paired with its reading is measured by height and
    // width; unpaired, its width belongs to no one line of the reading.
    const emPx = paired
      ? officeInkLineEmPx(line, words)
      : Math.max(1, line.bottom - line.top) / officeInkTextEmExtent(words);
    return officeInkNearestSize(pxToPt(emPx), sizes);
  };
  const out = textLines.map(() => fallback);
  if (said.length === written.length) {
    said.forEach(({ line, index }, at) => {
      out[index] = sizeOf(written[at]!, line, true);
    });
    return out;
  }
  // Fewer written lines than read ones: all take the middle of the written lines' sizes.
  const each = written.map((line) => sizeOf(line, text, false)).sort((a, b) => a - b);
  const middle = each[Math.floor((each.length - 1) / 2)] ?? fallback;
  return textLines.map(() => middle);
}

/**
 * One size for a text box that holds the whole reading (a slide's, a
 * sheet's): the middle of its lines' written sizes ({@link officeInkWrittenTextSizes}).
 */
export function officeInkWrittenTextSize(
  strokes: readonly (readonly { readonly x: number; readonly y: number }[])[],
  text: string,
  sizes: readonly number[],
): number {
  const lines = text.split("\n");
  const each = officeInkWrittenTextSizes(strokes, text, sizes)
    .filter((_, index) => (lines[index] ?? "").trim() !== "")
    .sort((a, b) => a - b);
  if (each.length === 0) return officeInkWrittenTextSizes(strokes, text, sizes)[0] ?? 11;
  return each[Math.floor((each.length - 1) / 2)]!;
}

/**
 * How a text box draws the words a reading became, in the host's own CSS px
 * (a slide's design px, a sheet's px at 100 %): the font, its size and line
 * box, and the frame's four insets.
 */
export interface OfficeInkTextFrame {
  readonly fontSizePx: number;
  /** The CSS font stack the box draws its words in (what the host renders). */
  readonly fontFamily: string;
  readonly fontWeight?: string | number;
  /** Line box as a multiple of the font size. */
  readonly lineHeight: number;
  readonly insets: { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number };
}

/** A CJK or full-width glyph: about one em wide. */
const OFFICE_INK_WIDE_GLYPH = /[ᄀ-ᅟ⺀-〾ぁ-㏿㐀-䶿一-鿿ꥠ-꥿가-힣豈-﫿︰-﹏＀-｠￠-￦]/u;

let officeInkMeasureContext: CanvasRenderingContext2D | null | undefined;

/**
 * One line's advance in px, drawn in `frame`'s font: measured with the
 * document's own fonts (a canvas resolves the stack and its fallbacks as the
 * page does); without a canvas, estimated (a wide glyph 1 em, others 0.6 em).
 */
export function officeInkTextLineWidthPx(line: string, frame: OfficeInkTextFrame): number {
  if (officeInkMeasureContext === undefined) {
    try {
      officeInkMeasureContext =
        typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
    } catch {
      officeInkMeasureContext = null;
    }
  }
  const context = officeInkMeasureContext;
  if (context) {
    context.font = `${frame.fontWeight ?? 400} ${frame.fontSizePx}px ${cssFontFamilyList(frame.fontFamily)}`;
    const width = context.measureText(line).width;
    if (Number.isFinite(width) && (width > 0 || line === "")) return width;
  }
  let ems = 0;
  for (const char of line) ems += OFFICE_INK_WIDE_GLYPH.test(char) ? 1 : 0.6;
  return ems * frame.fontSizePx;
}

/**
 * How many lines `line` is drawn on at `available` px: broken at its spaces,
 * and between wide glyphs (CJK lines break between any two), as the text box
 * wraps it. A word wider than the room still takes one line.
 */
function officeInkWrappedLineCount(line: string, available: number, frame: OfficeInkTextFrame): number {
  if (!(available > 0)) return 1;
  const pieces = line.match(/[^\s\u1100-\u115f\u2e80-\u303e\u3041-\u33ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7a3\uf900-\ufaff\uff00-\uff60]+\s*|[\u1100-\u115f\u2e80-\u303e\u3041-\u33ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7a3\uf900-\ufaff\uff00-\uff60]\s*|\s+/gu) ?? [];
  let count = 1;
  let used = 0;
  for (const piece of pieces) {
    const width = officeInkTextLineWidthPx(piece, frame);
    const ink = officeInkTextLineWidthPx(piece.trimEnd(), frame);
    if (used > 0 && used + ink > available) {
      count += 1;
      used = width;
    } else {
      used += width;
    }
  }
  return count;
}















export function officeInkTextBoxSize(
  text: string,
  frame: OfficeInkTextFrame,
  atLeast: { readonly width: number; readonly height: number },
  maxWidth?: number,
): { width: number; height: number } {
  const lines = text.split("\n");
  const widest = Math.max(0, ...lines.map((line) => officeInkTextLineWidthPx(line, frame)));
  const room = Math.max(2, frame.fontSizePx * 0.15);
  const extra = room + frame.insets.left + frame.insets.right;
  let width = Math.ceil(widest + extra);
  let drawn = lines.length;
  if (maxWidth !== undefined && Number.isFinite(maxWidth) && width > maxWidth) {
    width = Math.max(Math.floor(maxWidth), Math.ceil(atLeast.width));
    drawn = lines.reduce((sum, line) => sum + officeInkWrappedLineCount(line, width - extra, frame), 0);
  }
  const height = Math.ceil(drawn * frame.fontSizePx * frame.lineHeight + 2 + frame.insets.top + frame.insets.bottom);
  return { width: Math.max(atLeast.width, width), height: Math.max(atLeast.height, height) };
}
