"use client";















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";

import { inkFloatLayer } from "./office-ink-live";

export interface InkPenHoverLook {
  /** The disc's diameter on screen (CSS px): the stroke's width at the page's zoom, or the eraser's reach. */
  readonly diameterPx: number;
  /** The pen's colour; ignored for the eraser. */
  readonly color: string;
  readonly eraser: boolean;
}

export interface InkPenHoverOptions {
  /** A pen or eraser is picked: the pen's hover shows where it would land. */
  readonly active: boolean;
  /** Read at each hover move, so zoom and tool changes show at once. */
  readonly look: () => InkPenHoverLook | null;
}

/** The smallest disc shown (CSS px): a hairline pen still has a visible spot. */
const MIN_DIAMETER_PX = 5;

/**
 * While `options.active`, a pen hovering over `root` shows where it will
 * touch down. Hidden as soon as it touches, leaves the page or the tool is
 * put away.
 */
export function useInkPenHover(
  target: HTMLElement | null | (() => HTMLElement | null),
  options: InkPenHoverOptions,
): void {
  const latest = useRef(options);
  latest.current = options;
  // A ref's element is read when the hover starts (the page is mounted by then).
  const element = typeof target === "function" ? null : target;
  const readTarget = useRef(target);
  readTarget.current = target;
  useEffect(() => {
    const current = readTarget.current;
    const root = typeof current === "function" ? current() : current;
    if (!root || !options.active) return;
    const doc = root.ownerDocument;
    const disc = doc.createElement("div");
    disc.setAttribute("data-testid", "ink-pen-hover");
    Object.assign(disc.style, {
      position: "absolute",
      left: "0px",
      top: "0px",
      borderRadius: "50%",
      pointerEvents: "none",
      display: "none",
      boxSizing: "border-box",
    } satisfies Partial<CSSStyleDeclaration>);
    inkFloatLayer(doc).append(disc);
    const hide = () => {
      disc.style.display = "none";
    };
    const onMove = (event: PointerEvent) => {
      if (event.pointerType !== "pen" || event.buttons !== 0) {
        hide();
        return;
      }
      const look = latest.current.look();
      if (!look || !look.eraser) {
        hide();
        return;
      }
      const size = Math.max(MIN_DIAMETER_PX, look.diameterPx);
      disc.style.display = "block";
      disc.style.width = `${size}px`;
      disc.style.height = `${size}px`;
      disc.style.transform = `translate(${event.clientX - size / 2}px, ${event.clientY - size / 2}px)`;
      if (look.eraser) {
        disc.style.border = "1px solid rgba(71, 85, 105, 0.9)";
        disc.style.background = "rgba(148, 163, 184, 0.22)";
      } else {
        disc.style.border = `1px solid ${look.color}`;
        disc.style.background = `color-mix(in srgb, ${look.color} 30%, transparent)`;
      }
    };
    const onOut = (event: PointerEvent) => {
      if (event.pointerType === "pen") hide();
    };
    root.addEventListener("pointermove", onMove, { passive: true });
    root.addEventListener("pointerdown", hide, { passive: true });
    root.addEventListener("pointerleave", onOut, { passive: true });
    root.addEventListener("pointercancel", hide, { passive: true });
    doc.addEventListener("scroll", hide, { capture: true, passive: true });
    return () => {
      root.removeEventListener("pointermove", onMove);
      root.removeEventListener("pointerdown", hide);
      root.removeEventListener("pointerleave", onOut);
      root.removeEventListener("pointercancel", hide);
      doc.removeEventListener("scroll", hide, { capture: true });
      disc.remove();
    };
  }, [element, options.active]);
}
