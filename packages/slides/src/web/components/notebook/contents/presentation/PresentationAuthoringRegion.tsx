"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The authoring region wrapper that freezes the Ribbon and workspace under a
 * running slideshow: it stops re-rendering them and refuses the events through
 * which a covered control could still run a command.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { memo, startTransition, useCallback, useDeferredValue, useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { flushSync } from "react-dom";

// Keep the covered authoring DOM mounted during a slideshow without rebuilding
// its Ribbon and thumbnails on every animation step. The latest render closure
// is used immediately on exit; document state and the live show remain outside.
const PresentationAuthoringRegionContent = memo(
  function PresentationAuthoringRegionContent({ render }: { frozen: boolean; render: () => ReactNode }) {
    return render();
  },
  // The region renders ONCE more as it freezes, so what the show hides
  // (the Mini Toolbar over a selection, portaled to the body) leaves with
  // the show's first frame instead of staying over it (PPT-B238). A
  // deferred region (below) handed the closure it already drew is not drawn
  // again.
  (previous, next) => (next.frozen && previous.frozen) || previous.render === next.render,
);

/** The events, beside the refused ones, that make the input a deferred
 * region's own, and how long after the last one it is still drawn at once. */
const PRESENTATION_REGION_OWN_INPUT_EVENT_PROPS = [
  "onKeyDownCapture",
  "onFocusCapture",
] as const;
const PRESENTATION_REGION_OWN_INPUT_MS = 1000;
/** Where a deferred region's portaled controls are drawn. */
const PRESENTATION_REGION_POPUP_SELECTOR =
  '[data-radix-popper-content-wrapper], [data-radix-portal], [role="menu"], [role="listbox"], [role="dialog"], [role="tooltip"]';
/** The inputs before which a deferred region's pending drawing is done. */
const PRESENTATION_REGION_INPUT_EVENTS = [
  "pointerdown",
  "mousedown",
  "click",
  "auxclick",
  "dblclick",
  "contextmenu",
  "keydown",
  "beforeinput",
  "paste",
  "drop",
] as const;

/** Whether this document is painted (not jsdom, which renders nothing). */
function regionPaints(): boolean {
  return typeof window !== "undefined" && typeof window.requestAnimationFrame === "function" &&
    !(typeof navigator !== "undefined" && /jsdom/i.test(navigator.userAgent));
}

/** The order the region was handed its render closures in (each parent
 * render makes a new one): a deferred region compares them by it. */
const renderGenerations = new WeakMap<() => ReactNode, number>();
let renderGeneration = 0;

/**
 * The events through which a covered control could run a command. Keyboard
 * activation of a button or a menu item reaches a handler as a `click` (the
 * browser's for a button, Radix's own `.click()` for an item), so it is refused
 * here too. `keydown` itself is not on the list: the show's keyboard listener
 * sits on the document, and a region that still held focus when the show began
 * must not swallow the keys that drive it. `inert` below takes focus and
 * keyboard reach away from the region in the browser.
 */
const PRESENTATION_FROZEN_REGION_EVENT_PROPS = [
  "onClickCapture",
  "onDoubleClickCapture",
  "onAuxClickCapture",
  "onContextMenuCapture",
  "onPointerDownCapture",
  "onPointerUpCapture",
  "onMouseDownCapture",
  "onMouseUpCapture",
  "onChangeCapture",
  "onInputCapture",
  "onBeforeInputCapture",
  "onSubmitCapture",
  "onPasteCapture",
  "onCutCapture",
  "onDropCapture",
] as const;














export function PresentationAuthoringRegion({
  frozen,
  render,
  deferred = false,
  held = false,
  heldKey,
}: {
  frozen: boolean;
  render: () => ReactNode;
  













  deferred?: boolean;
  








  held?: boolean;
  heldKey?: string;
}) {
  if (!renderGenerations.has(render)) renderGenerations.set(render, ++renderGeneration);
  // A deferred region takes a new closure only once the frame that answers
  // the input has been painted (a frame, then a task after it), and then at
  // React's deferred priority. React starts deferred work at once, so the
  // Ribbon's new contextual tab -- its style, layout and layers -- was still
  
  // table, 104 ms to the next paint with 12 ms of it in the handler). At most
  // one hand-over a frame, always the latest closure: renders that come
  // faster than frames never starve it. jsdom paints nothing: there the
  // closure is handed over as React defers it, as before.
  // (A function held as state: set and initialised through a function that
  // returns it, so React neither calls it nor takes it for an updater.)
  const [paintedRender, setPaintedRender] = useState<() => ReactNode>(() => render);
  const latestForPaintRef = useRef(render);
  const handOverRef = useRef<{ frame: number; timer: number } | null>(null);
  useLayoutEffect(() => {
    latestForPaintRef.current = render;
  });
  const waitsForPaint = deferred && regionPaints();
  useEffect(() => {
    if (!waitsForPaint || paintedRender === render || handOverRef.current) return;
    const handOver = { frame: 0, timer: 0 };
    handOverRef.current = handOver;
    handOver.frame = requestAnimationFrame(() => {
      handOver.timer = window.setTimeout(() => {
        handOverRef.current = null;
        const latest = latestForPaintRef.current;
        startTransition(() => setPaintedRender(() => latest));
      }, 0);
    });
  }, [waitsForPaint, paintedRender, render]);
  useEffect(() => () => {
    const handOver = handOverRef.current;
    if (!handOver) return;
    cancelAnimationFrame(handOver.frame);
    window.clearTimeout(handOver.timer);
    handOverRef.current = null;
  }, []);
  const deferredRender = useDeferredValue(waitsForPaint ? paintedRender : render);
  // Freezing and thawing draw the latest closure at once (the show's first
  // frame, its exit); only a live region's updates wait. What this region
  // last committed, and whether it was frozen then:
  const committedRef = useRef<{ render: () => ReactNode; frozen: boolean; held: boolean; heldKey: string | undefined }>(
    { render, frozen, held: held && regionPaints(), heldKey },
  );
  // Held from the render after the one that hid it, for the same object. In
  // jsdom, which paints nothing, a held region is drawn at every render as
  // before (tests drive the closed pane's controls without opening it).
  const holdsWhenHidden = held && regionPaints();
  const holds = holdsWhenHidden && committedRef.current.held && committedRef.current.heldKey === heldKey;
  // Either way the region may show an older closure than the window's latest.
  const lags = deferred || holdsWhenHidden;
  const thawing = committedRef.current.frozen && !frozen;
  const hostRef = useRef<HTMLDivElement | null>(null);
  const ownInputAtRef = useRef(Number.NEGATIVE_INFINITY);
  const ownInput =
    lags &&
    (performance.now() - ownInputAtRef.current < PRESENTATION_REGION_OWN_INPUT_MS ||
      (typeof document !== "undefined" && Boolean(hostRef.current?.contains(document.activeElement))));
  // An input that arrives while the deferred drawing is still to come draws
  // it first (below): no press, key or click ever reaches a control of the
  // region that runs an older closure.
  const [, setDrawLatest] = useState(0);
  const drawLatestRef = useRef(false);
  const latestRef = useRef(render);
  let shown = !frozen && !thawing && !ownInput && !drawLatestRef.current
    ? holds
      ? committedRef.current.render
      : deferred
        ? deferredRender
        : render
    : render;
  // Never back to a closure older than the one drawn: a second input before
  // the deferred pass lands hands the older deferred value back for a moment.
  if ((renderGenerations.get(shown) ?? 0) < (renderGenerations.get(committedRef.current.render) ?? 0)) {
    shown = committedRef.current.render;
  }
  useLayoutEffect(() => {
    committedRef.current = { render: shown, frozen, held: holdsWhenHidden, heldKey };
    latestRef.current = render;
    drawLatestRef.current = false;
  });
  useEffect(() => {
    if (!lags || typeof window === "undefined") return;
    // The window's capture phase runs before React reads the target's
    // handlers for the event: drawn now, the region hands it its latest ones
    // (a menu it opened in a portal included).
    const drawLatest = (event: Event) => {
      if (committedRef.current.render === latestRef.current || committedRef.current.frozen) return;
      // Only an input the region's controls can receive: inside it, or in a
      // popup (its menus and galleries are portaled to the body). A press on
      // the slide leaves the drawing deferred.
      const target = event.type === "keydown" || event.type === "beforeinput" || event.type === "paste"
        ? (document.activeElement ?? event.target)
        : event.target;
      if (!(target instanceof Element)) return;
      // A held region has no popup open (it is not on screen): a popup is
      // another region's.
      const inPopup = !committedRef.current.held && target.closest(PRESENTATION_REGION_POPUP_SELECTOR) !== null;
      if (!hostRef.current?.contains(target) && !inPopup) return;
      drawLatestRef.current = true;
      flushSync(() => setDrawLatest((count) => count + 1));
    };
    for (const type of PRESENTATION_REGION_INPUT_EVENTS) window.addEventListener(type, drawLatest, true);
    return () => {
      for (const type of PRESENTATION_REGION_INPUT_EVENTS) window.removeEventListener(type, drawLatest, true);
    };
  }, [lags]);
  const refuse = useCallback(
    (event: { cancelable: boolean; preventDefault(): void; stopPropagation(): void }) => {
      if (!frozen) return;
      event.stopPropagation();
      if (event.cancelable) event.preventDefault();
    },
    [frozen],
  );
  const noteOwnInput = useCallback(() => {
    ownInputAtRef.current = performance.now();
  }, []);
  const refuseOrNote = useCallback(
    (event: { cancelable: boolean; preventDefault(): void; stopPropagation(): void }) => {
      if (frozen) refuse(event);
      else noteOwnInput();
    },
    [frozen, noteOwnInput, refuse],
  );
  const handlers = {
    ...(lags ? Object.fromEntries(PRESENTATION_REGION_OWN_INPUT_EVENT_PROPS.map((name) => [name, noteOwnInput])) : {}),
    ...Object.fromEntries(
      PRESENTATION_FROZEN_REGION_EVENT_PROPS.map((name) => [name, lags ? refuseOrNote : refuse]),
    ),
  };
  return (
    <div
      ref={hostRef}
      {...handlers}
      style={{ display: "contents" }}
      inert={frozen || undefined}
      data-presentation-authoring-region={frozen ? "frozen" : "live"}
    >
      <PresentationAuthoringRegionContent frozen={frozen} render={shown} />
    </div>
  );
}
