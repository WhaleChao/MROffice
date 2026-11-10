"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useEffect, useState, useSyncExternalStore } from "react";

import { penPresent, subscribeDrawWithTouch, useDeviceHasTouch } from "./draw-with-touch";
import { useDrawTabScribbleGuard } from "./ink-scribble-guard";
import { installPenPaletteMotion } from "./pen-palette-motion";
import { waitForInkHoldRest } from "./office-ink-hold-shape";














function serverNoPen(): boolean {
  return false;
}

/** True while a pen has touched or hovered the page lately (lib/draw-with-touch), live. */
export function usePenPresent(): boolean {
  return useSyncExternalStore(subscribeDrawWithTouch, () => penPresent(), serverNoPen);
}

export interface PenFirstDraw {
  /** The ribbon is folded and the floating palette stands in for it. */
  readonly compact: boolean;
  /** The reader asked for the ribbon: it stays until the Draw tab is left. */
  readonly showRibbon: () => void;
  /** Back to the palette from the ribbon (while drawing with a pen). */
  readonly showPalette: () => void;
  /** A pen is about on a touch device: the palette is available. */
  readonly available: boolean;
}

/**
 * The pen-first state of one editor's Draw tab. `drawActive`: the Draw tab
 * is the ribbon's tab and the editor can be drawn on.
 */
/**
 * `onDrawTab`: the Draw tab is open, whether or not the page may be drawn on
 * yet (a page still connecting, a read-only one): Scribble stays off there
 * too. Defaults to `drawActive`.
 */
/**
 * When a pen last did anything in the window: touched, moved, hovered (an
 * Apple Pencil over an M2 iPad reports its hover), lifted. The ribbon folds
 * into the palette, or unfolds, only once the pen has been quiet for
 * {@link PEN_LAYOUT_QUIET_MS}: the page moves when it folds, and a page that
 * moves under a running stroke is paper pulled from under the pen -- the
 * first stroke of a session, a 30px line, landed 210px long. A Pencil
 * announces itself by its hover before it touches, and the fold's glide ran
 * on into the stroke; folding at a pause in the writing, nothing is drawn
 * while the page moves.
 */
export const PEN_LAYOUT_QUIET_MS = 900;
let penActiveAt = Number.NEGATIVE_INFINITY;
let penOnGlass = false;

function installPenActivityWatch(): void {
  if (typeof window === "undefined" || (installPenActivityWatch as { done?: boolean }).done) return;
  (installPenActivityWatch as { done?: boolean }).done = true;
  const note = (event: PointerEvent) => {
    if (event.pointerType !== "pen") return;
    penActiveAt = performance.now();
    if (event.type === "pointerdown") penOnGlass = true;
    else if (event.type === "pointerup" || event.type === "pointercancel") penOnGlass = false;
  };
  const options = { capture: true, passive: true } as const;
  for (const type of ["pointerdown", "pointermove", "pointerup", "pointercancel"]) {
    window.addEventListener(type, note as EventListener, options);
  }
}

installPenActivityWatch();









export const DRAW_ACTIVE_DIP_MS = 3000;

/** Milliseconds until the pen will have been quiet long enough (0: it has). */
function penQuietIn(now: number = performance.now()): number {
  // On the glass, unless its lift was lost (no word from it for seconds).
  if (penOnGlass && now - penActiveAt < 5000) return PEN_LAYOUT_QUIET_MS;
  return Math.max(0, penActiveAt + PEN_LAYOUT_QUIET_MS - now);
}

/**
 * `drawingArea` is where the editor's pages are (read when a finger or the
 * keyboard takes a field): on the Draw tab nothing in it is given back for
 * typing, so the pencil never starts Scribble there.
 */
export function usePenFirstDraw(
  drawActive: boolean,
  onDrawTab: boolean = drawActive,
  drawingArea?: () => readonly (Element | null | undefined)[],
): PenFirstDraw {
  const penNow = usePenPresent();
  // The pen as the layout knows it: it changes only while the pen rests.
  const [pen, setPen] = useState(penNow);
  useEffect(() => {
    if (pen === penNow) return;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const settle = () => {
      const wait = penQuietIn();
      if (wait > 0) {
        timer = setTimeout(settle, wait + 16);
        return;
      }
      timer = null;
      setPen(penNow);
    };
    settle();
    return () => {
      if (timer !== null) clearTimeout(timer);
    };
  }, [pen, penNow]);
  const touch = useDeviceHasTouch();
  const [ribbonAsked, setRibbonAsked] = useState(false);
  // Drawing as the layout knows it. It comes on at once, and goes at once
  // when the Draw tab is left; with the tab still open it goes only after
  // {@link DRAW_ACTIVE_DIP_MS}, and with the pen at rest.
  const [drawShown, setDrawShown] = useState(drawActive);
  if (drawShown !== drawActive && (drawActive || !onDrawTab)) setDrawShown(drawActive);
  useEffect(() => {
    if (drawShown === drawActive || drawActive || !onDrawTab) return;
    const since = performance.now();
    let timer: ReturnType<typeof setTimeout> | null = null;
    const settle = () => {
      const wait = Math.max(penQuietIn(), DRAW_ACTIVE_DIP_MS - (performance.now() - since));
      if (wait > 0) {
        timer = setTimeout(settle, wait + 16);
        return;
      }
      timer = null;
      setDrawShown(false);
    };
    settle();
    return () => {
      if (timer !== null) clearTimeout(timer);
    };
  }, [drawActive, drawShown, onDrawTab]);
  // Watching the ribbons before any folds: the fold and the unfold glide.
  useEffect(() => installPenPaletteMotion(), []);
  // On the Draw tab the Apple Pencil never starts Scribble, whatever tool is
  // picked (lib/ink-scribble-guard).
  useDrawTabScribbleGuard(onDrawTab, drawingArea);
  // Leaving the Draw tab forgets the ask: coming back with the pen folds again.
  const [wasActive, setWasActive] = useState(drawShown);
  if (wasActive !== drawShown) {
    setWasActive(drawShown);
    if (!drawShown) setRibbonAsked(false);
  }
  const showRibbon = useCallback(() => setRibbonAsked(true), []);
  const showPalette = useCallback(() => setRibbonAsked(false), []);
  const available = pen && touch;
  return { compact: drawShown && available && !ribbonAsked, showRibbon, showPalette, available };
}

/** How long the pen rests on the page, unmoved, before the stroke turns into the eraser. */
export const INK_HOLD_TO_ERASE_MS = 350;






export const INK_HOLD_STILL_PX = 10;

export interface InkHoldToErase {
  /** Each move of the pen: one that strays past the still reach ends the wait. */
  move(event: { readonly pointerId: number; readonly clientX: number; readonly clientY: number }): void;
  /** The stroke ended (lifted, cancelled, replaced) before the hold. */
  cancel(): void;
}












export function startInkHoldToErase(
  event: { readonly pointerType: string; readonly pointerId: number; readonly clientX: number; readonly clientY: number },
  onHold: () => void,
): InkHoldToErase | null {
  if (event.pointerType !== "pen") return null;
  const { pointerId, clientX: startX, clientY: startY } = event;
  // A rest the page could hear (waitForInkHoldRest): a first stroke written
  // while the page was busy is not a pen held still.
  let timer: { cancel(): void } | null = waitForInkHoldRest(() => {
    timer = null;
    onHold();
  }, INK_HOLD_TO_ERASE_MS);
  const cancel = () => {
    timer?.cancel();
    timer = null;
  };
  return {
    move(next) {
      if (timer === null || next.pointerId !== pointerId) return;
      if (Math.hypot(next.clientX - startX, next.clientY - startY) > INK_HOLD_STILL_PX) cancel();
    },
    cancel,
  };
}
