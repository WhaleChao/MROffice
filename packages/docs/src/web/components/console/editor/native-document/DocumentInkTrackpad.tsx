"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office for Mac's "Draw with Trackpad": while the Draw tab's switch is on,
 * the trackpad is a writing tablet. The document dims, a frame the shape of
 * the trackpad stands over it, and
 *
 *   - one finger draws inside the frame, onto the document under it;
 *   - two fingers move, the pen inside the frame and, pushed against an
 *     edge, the frame over the document;
 *   - any key ends the mode (Undo and Redo act on the strokes instead).
 *
 * Office reads the finger's absolute place on the trackpad. A web page is
 * given no such thing, so the pointer is locked instead and the pen follows
 * the finger's movement: a stroke ends when the finger rests or lifts, and
 * two fingers carry the pen to where the next one starts. Where the pointer
 * cannot be locked the frame is written in with the button held.
 *
 * Entering, the frame gathers from the four edges of the document area to
 * its place over the same time, on the same curve, as the switch's knob
 * crosses its track (OFFICE_INK_TRACKPAD_TRANSITION_MS).
 */

import { ptToPx } from "@mrnote/core/web/lib/office-units.ts";
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type PointerEvent as ReactPointerEvent,
} from "react";
import { createPortal } from "react-dom";

import {
  OFFICE_INK_TRACKPAD_EASING,
  OFFICE_INK_TRACKPAD_TRANSITION_MS,
} from "@mrnote/core/web/components/office-ribbon/officeInkTrackpadMotion.ts";
import {
  MAX_DOCUMENT_INK_POINTS,
  documentInkEffectPaint,
  documentInkOpacity,
  documentInkPencilFilter,
  documentInkPencilRegionFor,
  type DocumentInkPoint,
  type DocumentInkPreset,
} from "@mrnote/core/web/lib/document-ink.ts";
import { claimInkTrackpadLock, inkTrackpadLockHeld } from "@mrnote/core/web/lib/ink-trackpad-lock.ts";
import { inkStrokePathData } from "@mrnote/core/web/lib/office-ink-smooth.ts";
import { DocumentInkPencilFilter } from "./DocumentInkPencilFilter";

/** The frame's size at rest: a trackpad's proportions. */
export const DOCUMENT_INK_TRACKPAD_WIDTH_PX = 430;
export const DOCUMENT_INK_TRACKPAD_HEIGHT_PX = 280;
/** A finger at rest this long has finished its stroke. */
const STROKE_REST_MS = 170;
const FRAME_RADIUS_PX = 14;
const EXIT_MS = 220;

interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface DocumentInkTrackpadProps {
  /** The mode is on (the Draw tab's switch). */
  open: boolean;
  preset: DocumentInkPreset;
  /** The mode's name, for assistive technology. */
  title: string;
  /** "Draw with one finger. Move with two fingers. Press any key to exit." */
  hint: string;
  /**
   * The element the document scrolls in; the mode dims and frames its
   * nearest scrolling ancestor (the whole viewport without one).
   */
  anchor?: () => Element | null;
  /**
   * The ribbon, when it stands inside that scrolling element: the mode
   * dims the document under it, never the ribbon itself.
   */
  below?: () => Element | null;
  /**
   * The page the frame stands on, in client pixels (the paper, the grid, the
   * slide): no edge of the frame leaves it, so nothing is drawn off the page.
   * Given the dimmed area, answers the page to draw on.
   */
  page?: (area: {
    left: number;
    top: number;
    right: number;
    bottom: number;
  }) => { left: number; top: number; right: number; bottom: number } | null;
  /** One finished stroke, in client pixels. */
  onStroke: (points: readonly DocumentInkPoint[]) => void;
  onClose: () => void;
}

const clamp = (value: number, low: number, high: number) =>
  Math.min(Math.max(value, low), Math.max(low, high));

function areaOf(anchor: Element | null, below: Element | null): Box {
  const area = scrollAreaOf(anchor);
  const floor = below?.getBoundingClientRect().bottom ?? area.top;
  const top = clamp(floor, area.top, area.top + area.height - 1);
  return { ...area, top, height: area.top + area.height - top };
}

function scrollAreaOf(anchor: Element | null): Box {
  const viewport = { left: 0, top: 0, width: window.innerWidth, height: window.innerHeight };
  for (let element = anchor; element; element = element.parentElement) {
    if (!(element instanceof HTMLElement)) continue;
    const overflow = window.getComputedStyle(element).overflowY;
    if (overflow !== "auto" && overflow !== "scroll") continue;
    const rect = element.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) {
      return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
    }
  }
  return viewport;
}

/**
 * Where the frame may stand, in the area's own pixels: the part of the page
 * the area shows (the whole area, less a margin, without a page).
 */
function frameLimits(
  area: Box,
  page: { left: number; top: number; right: number; bottom: number } | null,
): Box {
  if (!page) {
    return { left: 24, top: 48, width: Math.max(1, area.width - 48), height: Math.max(1, area.height - 96) };
  }
  const left = clamp(page.left - area.left, 0, area.width);
  const top = clamp(page.top - area.top, 0, area.height);
  const right = clamp(page.right - area.left, left, area.width);
  const bottom = clamp(page.bottom - area.top, top, area.height);
  return { left, top, width: right - left, height: bottom - top };
}

/** The frame at rest: centred in its limits, never larger than them. */
function restingFrame(limits: Box): Box {
  const ratio = DOCUMENT_INK_TRACKPAD_HEIGHT_PX / DOCUMENT_INK_TRACKPAD_WIDTH_PX;
  let width = Math.min(DOCUMENT_INK_TRACKPAD_WIDTH_PX, limits.width);
  let height = width * ratio;
  if (height > limits.height) {
    height = limits.height;
    width = Math.min(width, height / ratio);
  }
  return {
    left: limits.left + (limits.width - width) / 2,
    top: limits.top + (limits.height - height) / 2,
    width,
    height,
  };
}


export default function DocumentInkTrackpad({
  open,
  preset,
  title,
  hint,
  anchor,
  below,
  page,
  onStroke,
  onClose,
}: DocumentInkTrackpadProps) {
  // Mounted through the closing animation, past `open`.
  const [session, setSession] = useState<{ area: Box; limits: Box } | null>(null);
  const [closing, setClosing] = useState(false);
  const [frame, setFrame] = useState<Box | null>(null);
  const [pen, setPen] = useState<DocumentInkPoint | null>(null);
  const [preview, setPreview] = useState<DocumentInkPoint[] | null>(null);
  const [locked, setLocked] = useState(false);

  const rootRef = useRef<HTMLDivElement | null>(null);
  const frameElementRef = useRef<HTMLDivElement | null>(null);
  const frameRef = useRef<Box | null>(null);
  const penRef = useRef<DocumentInkPoint | null>(null);
  const strokeRef = useRef<DocumentInkPoint[] | null>(null);
  const restTimerRef = useRef<number | null>(null);
  const lockedRef = useRef(false);
  const everLockedRef = useRef(false);
  const dragRef = useRef<number | null>(null);
  const onStrokeRef = useRef(onStroke);
  const onCloseRef = useRef(onClose);
  onStrokeRef.current = onStroke;
  onCloseRef.current = onClose;
  const anchorRef = useRef(anchor);
  anchorRef.current = anchor;
  const belowRef = useRef(below);
  belowRef.current = below;
  const pageRef = useRef(page);
  pageRef.current = page;

  const area = session?.area ?? null;
  const sessionAreaRef = useRef<Box | null>(null);
  sessionAreaRef.current = area;

  /** Hand the stroke in progress to the document (points are area pixels). */
  const commitStroke = useCallback(() => {
    if (restTimerRef.current !== null) {
      window.clearTimeout(restTimerRef.current);
      restTimerRef.current = null;
    }
    const points = strokeRef.current;
    strokeRef.current = null;
    setPreview(null);
    const origin = sessionAreaRef.current;
    if (!points || points.length === 0 || !origin) return;
    onStrokeRef.current(
      points.map((point) => ({ x: point.x + origin.left, y: point.y + origin.top })),
    );
  }, []);
  const extendStroke = useCallback((point: DocumentInkPoint) => {
    const stroke = strokeRef.current;
    if (!stroke) {
      strokeRef.current = [point];
      setPreview([point]);
      return;
    }
    if (stroke.length >= MAX_DOCUMENT_INK_POINTS) return;
    const previous = stroke[stroke.length - 1];
    if (Math.hypot(point.x - previous.x, point.y - previous.y) < 1) return;
    stroke.push(point);
    setPreview([...stroke]);
  }, []);

  // Open: measure, and take the session. Close: play out, then leave.
  useEffect(() => {
    if (open) {
      const measured = areaOf(anchorRef.current?.() ?? null, belowRef.current?.() ?? null);
      const limits = frameLimits(
        measured,
        pageRef.current?.({
          left: measured.left,
          top: measured.top,
          right: measured.left + measured.width,
          bottom: measured.top + measured.height,
        }) ?? null,
      );
      const resting = restingFrame(limits);
      const centre = {
        x: resting.left + resting.width / 2,
        y: resting.top + resting.height / 2,
      };
      frameRef.current = resting;
      penRef.current = centre;
      everLockedRef.current = false;
      setFrame(resting);
      setPen(centre);
      setClosing(false);
      setSession({ area: measured, limits });
      return;
    }
    if (!sessionAreaRef.current) return;
    commitStroke();
    const root = rootRef.current;
    if (inkTrackpadLockHeld(root)) document.exitPointerLock?.();
    const reduced =
      typeof window.matchMedia === "function" &&
      window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (!root || typeof root.animate !== "function" || reduced) {
      setSession(null);
      return;
    }
    setClosing(true);
    const leave = window.setTimeout(() => {
      setClosing(false);
      setSession(null);
    }, EXIT_MS);
    return () => window.clearTimeout(leave);
  }, [commitStroke, open]);

  // The frame gathers from the area's edges as the switch's knob slides.
  useLayoutEffect(() => {
    const element = frameElementRef.current;
    const root = rootRef.current;
    const resting = frameRef.current;
    if (!session || !element || !root || !resting) return;
    if (typeof element.animate !== "function") return;
    const reduced =
      typeof window.matchMedia === "function" &&
      window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduced) return;
    const timing = {
      duration: OFFICE_INK_TRACKPAD_TRANSITION_MS,
      easing: OFFICE_INK_TRACKPAD_EASING,
    };
    const gather = element.animate(
      [
        {
          left: "0px",
          top: "0px",
          width: `${session.area.width}px`,
          height: `${session.area.height}px`,
          borderRadius: "0px",
          backgroundColor: "rgb(255 255 255 / 0%)",
        },
        {
          left: `${resting.left}px`,
          top: `${resting.top}px`,
          width: `${resting.width}px`,
          height: `${resting.height}px`,
          borderRadius: `${FRAME_RADIUS_PX}px`,
          backgroundColor: "rgb(255 255 255 / 46%)",
        },
      ],
      timing,
    );
    const dim = root.animate([{ backgroundColor: "rgb(15 23 42 / 0%)" }, {}], timing);
    return () => {
      gather.cancel();
      dim.cancel();
    };
  }, [session]);

  // Closing: the frame lets go the way it came.
  useLayoutEffect(() => {
    const element = frameElementRef.current;
    const root = rootRef.current;
    const from = frameRef.current;
    if (!closing || !element || !root || !from || !session) return;
    if (typeof element.animate !== "function") return;
    const timing = { duration: EXIT_MS, easing: "cubic-bezier(0.4, 0, 1, 1)", fill: "forwards" as const };
    element.animate(
      [
        {},
        {
          left: "0px",
          top: "0px",
          width: `${session.area.width}px`,
          height: `${session.area.height}px`,
          borderRadius: "0px",
          backgroundColor: "rgb(255 255 255 / 0%)",
          opacity: 0,
        },
      ],
      timing,
    );
    root.animate([{}, { backgroundColor: "rgb(15 23 42 / 0%)" }], timing);
  }, [closing, session]);

  // The pointer is the pen's while the mode is on.
  useEffect(() => {
    const root = rootRef.current;
    if (!session || closing || !root) return;
    const onLockChange = () => {
      const held = inkTrackpadLockHeld(root);
      lockedRef.current = held;
      setLocked(held);
      if (held) {
        everLockedRef.current = true;
        claimInkTrackpadLock();
        return;
      }
      // Esc (the browser's own way out of a locked pointer) is a key too.
      if (everLockedRef.current) onCloseRef.current();
    };
    document.addEventListener("pointerlockchange", onLockChange);
    // The switch locked the pointer in its own click (lib/ink-trackpad-lock):
    // a request made here, renders later, is refused by Safari.
    if (inkTrackpadLockHeld(root)) {
      onLockChange();
    } else if (typeof root.requestPointerLock === "function") {
      try {
        const request = root.requestPointerLock() as unknown;
        if (request instanceof Promise) request.catch(() => {});
      } catch {
        // No lock: the frame is written in with the button held.
      }
    }
    return () => {
      document.removeEventListener("pointerlockchange", onLockChange);
      lockedRef.current = false;
      setLocked(false);
      if (inkTrackpadLockHeld(root)) document.exitPointerLock?.();
    };
  }, [closing, session]);

  // One finger draws; two fingers move; any key ends the mode.
  useEffect(() => {
    const root = rootRef.current;
    if (!session || closing || !root) return;
    const bounds = session.limits;
    const within = (box: Box, left: number, top: number): Box => ({
      ...box,
      left: clamp(left, bounds.left, bounds.left + bounds.width - box.width),
      top: clamp(top, bounds.top, bounds.top + bounds.height - box.height),
    });

    const penInside = (point: DocumentInkPoint, box: Box): DocumentInkPoint => ({
      x: clamp(point.x, box.left + 1, box.left + box.width - 1),
      y: clamp(point.y, box.top + 1, box.top + box.height - 1),
    });

    const onMouseMove = (event: MouseEvent) => {
      if (!lockedRef.current) return;
      const box = frameRef.current;
      const from = penRef.current;
      if (!box || !from) return;
      if (event.movementX === 0 && event.movementY === 0) return;
      const next = penInside(
        { x: from.x + event.movementX, y: from.y + event.movementY },
        box,
      );
      if (!strokeRef.current) extendStroke(from);
      extendStroke(next);
      penRef.current = next;
      setPen(next);
      if (restTimerRef.current !== null) window.clearTimeout(restTimerRef.current);
      restTimerRef.current = window.setTimeout(commitStroke, STROKE_REST_MS);
    };

    const onWheel = (event: WheelEvent) => {
      event.preventDefault();
      const box = frameRef.current;
      const from = penRef.current;
      if (!box || !from) return;
      commitStroke();
      // The page follows the fingers (natural scrolling), and so does this.
      const dx = -event.deltaX;
      const dy = -event.deltaY;
      let nextBox = box;
      let nextPen = from;
      if (lockedRef.current) {
        const wanted = { x: from.x + dx, y: from.y + dy };
        nextPen = penInside(wanted, box);
        // What the frame's edge held back pushes the frame along.
        const pushX = wanted.x - nextPen.x;
        const pushY = wanted.y - nextPen.y;
        if (pushX !== 0 || pushY !== 0) {
          nextBox = within(box, box.left + pushX, box.top + pushY);
          nextPen = penInside(
            { x: nextPen.x + nextBox.left - box.left, y: nextPen.y + nextBox.top - box.top },
            nextBox,
          );
        }
      } else {
        nextBox = within(box, box.left + dx, box.top + dy);
        nextPen = {
          x: from.x + nextBox.left - box.left,
          y: from.y + nextBox.top - box.top,
        };
      }
      if (nextBox !== box) {
        frameRef.current = nextBox;
        setFrame(nextBox);
      }
      penRef.current = nextPen;
      setPen(nextPen);
    };

    const onKeyDown = (event: KeyboardEvent) => {
      // A modifier on its way down is not yet a key press: Cmd alone must not
      // end the mode before the Z that follows it arrives.
      if (["Meta", "Control", "Shift", "Alt"].includes(event.key)) return;
      // Undo and Redo take back and restore strokes without leaving the
      // mode: the stroke in hand lands first, so it is the one undone, and
      // the shortcut goes on to the editor.
      if (
        (event.metaKey || event.ctrlKey) &&
        !event.altKey &&
        ["z", "y"].includes(event.key.toLowerCase())
      ) {
        commitStroke();
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      onCloseRef.current();
    };
    const swallow = (event: Event) => {
      if (lockedRef.current) event.preventDefault();
    };
    // A locked pointer sends its wheel and buttons to the element it is
    // locked to (the frame, or the switch's lock target), so they are heard
    // on the document; unlocked, only over the frame.
    const onAnyWheel = (event: WheelEvent) => {
      if (!lockedRef.current && !(event.target instanceof Node && root.contains(event.target))) return;
      onWheel(event);
    };
    const onAnyButton = (event: Event) => {
      if (!lockedRef.current && !(event.target instanceof Node && root.contains(event.target))) return;
      swallow(event);
    };

    document.addEventListener("mousemove", onMouseMove, true);
    document.addEventListener("wheel", onAnyWheel, { capture: true, passive: false });
    window.addEventListener("keydown", onKeyDown, true);
    document.addEventListener("mousedown", onAnyButton, true);
    document.addEventListener("contextmenu", onAnyButton, true);
    return () => {
      document.removeEventListener("mousemove", onMouseMove, true);
      document.removeEventListener("wheel", onAnyWheel, { capture: true });
      window.removeEventListener("keydown", onKeyDown, true);
      document.removeEventListener("mousedown", onAnyButton, true);
      document.removeEventListener("contextmenu", onAnyButton, true);
    };
  }, [closing, commitStroke, extendStroke, session]);

  useEffect(
    () => () => {
      if (restTimerRef.current !== null) window.clearTimeout(restTimerRef.current);
    },
    [],
  );

  // Without a locked pointer: press in the frame and write.
  const framePoint = (event: ReactPointerEvent<HTMLDivElement>): DocumentInkPoint | null => {
    const box = frameRef.current;
    const origin = sessionAreaRef.current;
    if (!box || !origin) return null;
    return {
      x: clamp(event.clientX - origin.left, box.left, box.left + box.width),
      y: clamp(event.clientY - origin.top, box.top, box.top + box.height),
    };
  };
  const onPointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (lockedRef.current || event.button !== 0 || dragRef.current !== null) return;
    const start = framePoint(event);
    if (!start) return;
    event.preventDefault();
    dragRef.current = event.pointerId;
    try {
      event.currentTarget.setPointerCapture?.(event.pointerId);
    } catch {
      // A synthetic pointer has no capture.
    }
    strokeRef.current = null;
    extendStroke(start);
  };
  const onPointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (dragRef.current !== event.pointerId) return;
    const point = framePoint(event);
    if (point) extendStroke(point);
  };
  const onPointerUp = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (dragRef.current !== event.pointerId) return;
    dragRef.current = null;
    commitStroke();
  };
  const onPointerCancel = () => {
    dragRef.current = null;
    strokeRef.current = null;
    setPreview(null);
  };

  if (typeof document === "undefined" || !session || !area || !frame) return null;
  const effect =
    preset.tool === "pen" && preset.effect ? documentInkEffectPaint(preset.effect) : null;
  const strokeWidthPx = Math.max(1, ptToPx(preset.widthPt));
  const pencil =
    preset.tool === "pencil" && preview && preview.length > 0
      ? documentInkPencilFilter(strokeWidthPx, documentInkPencilRegionFor(preview, strokeWidthPx), 7)
      : null;

  return createPortal(
    <div
      ref={rootRef}
      role="application"
      aria-label={title}
      data-testid="document-draw-trackpad"
      data-locked={locked || undefined}
      data-closing={closing || undefined}
      style={{
        position: "fixed",
        left: area.left,
        top: area.top,
        width: area.width,
        height: area.height,
        zIndex: 60,
        overflow: "hidden",
        backgroundColor: "rgb(15 23 42 / 30%)",
        cursor: locked ? "none" : "default",
        userSelect: "none",
        touchAction: "none",
      }}
    >
      <p
        role="status"
        data-testid="document-draw-trackpad-hint"
        style={{
          position: "absolute",
          left: "50%",
          top: 20,
          transform: "translateX(-50%)",
          margin: 0,
          maxWidth: "min(460px, calc(100% - 32px))",
          padding: "10px 18px",
          borderRadius: 10,
          background: "rgb(38 38 42 / 86%)",
          color: "#ffffff",
          fontSize: 15,
          lineHeight: 1.45,
          textAlign: "center",
          textWrap: "balance",
          pointerEvents: "none",
          opacity: closing ? 0 : 1,
          transition: `opacity ${closing ? EXIT_MS : OFFICE_INK_TRACKPAD_TRANSITION_MS}ms ease`,
        }}
      >
        {hint}
      </p>
      <div
        ref={frameElementRef}
        data-testid="document-draw-trackpad-pad"
        style={{
          position: "absolute",
          left: frame.left,
          top: frame.top,
          width: frame.width,
          height: frame.height,
          boxSizing: "border-box",
          border: "1.5px solid rgb(28 28 30 / 88%)",
          borderRadius: FRAME_RADIUS_PX,
          backgroundColor: "rgb(255 255 255 / 46%)",
          boxShadow: "0 12px 40px rgb(15 23 42 / 18%)",
          cursor: locked ? "none" : "crosshair",
        }}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerCancel}
      />
      <svg
        aria-hidden="true"
        width={area.width}
        height={area.height}
        style={{ position: "absolute", inset: 0, pointerEvents: "none", overflow: "visible" }}
      >
        {effect ? (
          <defs>
            <linearGradient id="docs-ink-trackpad-effect" x1="0" y1="0" x2="1" y2="1">
              {effect.stops.map((stop, index) => (
                <stop
                  key={`${stop}-${index}`}
                  offset={`${(index / Math.max(1, effect.stops.length - 1)) * 100}%`}
                  stopColor={stop}
                />
              ))}
            </linearGradient>
          </defs>
        ) : null}
        {pencil ? (
          <defs>
            <DocumentInkPencilFilter id="docs-ink-trackpad-pencil" filter={pencil} />
          </defs>
        ) : null}
        {preview && preview.length > 0 ? (
          <g filter={pencil ? "url(#docs-ink-trackpad-pencil)" : undefined}>
            <path
              d={inkStrokePathData(preview)}
              fill="none"
              stroke={effect ? "url(#docs-ink-trackpad-effect)" : preset.color}
              strokeWidth={strokeWidthPx}
              strokeOpacity={documentInkOpacity(preset.tool)}
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </g>
        ) : null}
        {locked && pen && !closing ? (
          // The pen's tip: where the next stroke starts.
          <circle
            data-testid="document-draw-trackpad-pen"
            cx={pen.x}
            cy={pen.y}
            r={Math.max(3, strokeWidthPx / 2 + 1.5)}
            fill={preset.color}
            stroke="#ffffff"
            strokeWidth={1.5}
          />
        ) : null}
      </svg>
    </div>,
    document.body,
  );
}
