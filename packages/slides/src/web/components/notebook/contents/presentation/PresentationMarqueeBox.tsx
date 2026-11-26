"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useState } from "react";

/** The marquee's box in slide EMU (clamped into the slide), or none. */
export type PresentationMarqueeRect = { left: number; top: number; width: number; height: number };









export function PresentationMarqueeBox({
  register,
  slideWidth,
  slideHeight,
}: {
  register: (setter: ((rect: PresentationMarqueeRect | null) => void) | null) => void;
  slideWidth: number;
  slideHeight: number;
}) {
  const [rect, setRect] = useState<PresentationMarqueeRect | null>(null);
  useEffect(() => {
    register(setRect);
    return () => register(null);
  }, [register]);
  if (!rect) return null;
  return (
    <div
      className="presentation-window__marquee"
      data-testid="presentation-marquee"
      aria-hidden="true"
      style={{
        position: "absolute",
        left: `${(rect.left / slideWidth) * 100}%`,
        top: `${(rect.top / slideHeight) * 100}%`,
        width: `${(rect.width / slideWidth) * 100}%`,
        height: `${(rect.height / slideHeight) * 100}%`,
        border: "1px solid rgba(45, 127, 116, 0.9)",
        background: "color-mix(in srgb, var(--console-accent, var(--console-accent, #2d7f74)) 12%, transparent)",
        pointerEvents: "none",
        zIndex: 9998,
      }}
    />
  );
}
