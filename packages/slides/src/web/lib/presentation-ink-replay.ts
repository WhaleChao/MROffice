/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink Replay on a slide (PPT-A066): the slide's ink strokes as the
 * shared replay timeline (`lib/office-ink-replay.ts`, the one Excel's Draw tab
 * plays) takes them.
 *
 * Microsoft Support, "Replay your ink strokes in Office": the ink is removed
 * and redrawn in the order it was drawn. A slide stroke keeps no per-point
 * times, so each is drawn at the timeline's constant pace, in the order the
 * slide stacks its strokes (the order they were drawn).
 */

import { isInkEffect, type InkEffectId } from "@mrnote/core/web/lib/office-ink-effects.ts";

/** Slide EMU per CSS pixel: the timeline paces strokes in pixels. */
const EMU_PER_PX = 9_525;

export interface PresentationInkReplayStroke {
  readonly id: string;
  /** Slide EMU, in drawing order. */
  readonly points: readonly { readonly x: number; readonly y: number }[];
  readonly color: string;
  /** The stroke's width in slide EMU. */
  readonly widthEmu: number;
  /** 0-1. */
  readonly opacity: number;
  /** A pen's effect, drawn in its layers as the slide draws it. */
  readonly effect?: InkEffectId;
}

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/** The slide's ink strokes, back to front, as replay strokes. */
export function presentationInkReplayStrokes(
  blocks: readonly Record<string, unknown>[],
): PresentationInkReplayStroke[] {
  return blocks.flatMap((block) => {
    if (block.type !== "shape" || block.shape !== "ink" || typeof block.points !== "string") return [];
    if (block.hidden === true) return [];
    const left = finite(block.left);
    const top = finite(block.top);
    const width = finite(block.width);
    const height = finite(block.height);
    if (left === null || top === null || width === null || height === null) return [];
    const points = block.points
      .trim()
      .split(/\s+/)
      .map((pair) => pair.split(",").map(Number))
      .filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y))
      .map(([x, y]) => ({ x: left + (x! / 100) * width, y: top + (y! / 100) * height }));
    if (points.length < 2 || typeof block.id !== "string") return [];
    const opacity = finite(block.opacity);
    return [
      {
        id: block.id,
        points,
        color: typeof block.stroke === "string" ? block.stroke : "#000000",
        widthEmu: (finite(block.strokeWidth) ?? 1) * EMU_PER_PX,
        opacity: opacity === null ? 1 : Math.min(1, Math.max(0, opacity / 100)),
        ...((block.inkTool === undefined || block.inkTool === "pen") && isInkEffect(block.inkEffect)
          ? { effect: block.inkEffect }
          : {}),
      },
    ];
  });
}

/** The strokes in the timeline's pixel space. */
export function presentationInkReplayTimelineInput(
  strokes: readonly PresentationInkReplayStroke[],
): { id: string; points: { x: number; y: number }[] }[] {
  return strokes.map((stroke) => ({
    id: stroke.id,
    points: stroke.points.map((point) => ({ x: point.x / EMU_PER_PX, y: point.y / EMU_PER_PX })),
  }));
}
