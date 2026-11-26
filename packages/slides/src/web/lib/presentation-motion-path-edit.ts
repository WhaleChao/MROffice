/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Edit Points on a motion path (GAP #10, step 2): the adapter between a
 * path as PowerPoint writes it (shares of the slide, measured from the
 * object's centre on the slide; `lib/presentation-motion-path-native`) and
 * the editors' shared Edit Points overlay (`mountOfficeEditPoints`, figures
 * of lines and cubics in its host's pixels; `lib/office-custom-path-edit`).
 *
 * The host is the slide: a path is drawn and edited in the slide's frame,
 * whatever the turn of its object (PowerPoint does not turn a path with its
 * object, truth motion-path-roundtrip p7); `rAng` is kept as it is.
 */

import {
  officeEditPathFromCustomPath,
  type OfficeEditFigure,
  type OfficeEditPath,
  type OfficeEditSegment,
} from "@mrnote/core/web/lib/office-custom-path-edit.ts";
import { formatNativeMotionPath, parseNativeMotionPath } from "@mrnote/core/web/lib/presentation-motion-path-native.ts";

/** The object's centre on the slide and the slide's size, EMU; the host's size, px. */
export interface MotionPathEditFrame {
  readonly centreXEmu: number;
  readonly centreYEmu: number;
  readonly slideWidthEmu: number;
  readonly slideHeightEmu: number;
  readonly hostWidthPx: number;
  readonly hostHeightPx: number;
}

function frameUsable(frame: MotionPathEditFrame): boolean {
  return frame.slideWidthEmu > 0 && frame.slideHeightEmu > 0 && frame.hostWidthPx > 0 && frame.hostHeightPx > 0;
}

/** The path as the overlay edits it, in the slide host's pixels; null when it is no path. */
export function motionPathEditPath(path: string, frame: MotionPathEditFrame): OfficeEditPath | null {
  const parsed = parseNativeMotionPath(path);
  if (!parsed || !frameUsable(frame)) return null;
  const ox = frame.centreXEmu / frame.slideWidthEmu;
  const oy = frame.centreYEmu / frame.slideHeightEmu;
  const px = ([x, y]: readonly [number, number]) => ({
    x: (ox + x) * frame.hostWidthPx,
    y: (oy + y) * frame.hostHeightPx,
  });
  return {
    figures: parsed.subpaths
      .filter((subpath) => subpath.segments.length > 0)
      .map(
        (subpath): OfficeEditFigure => ({
          closed: subpath.closed,
          start: px(subpath.start),
          segments: subpath.segments.map(
            (segment): OfficeEditSegment =>
              segment.kind === "L"
                ? { kind: "line", to: px(segment.to) }
                : { kind: "cubic", c1: px(segment.c1), c2: px(segment.c2), to: px(segment.to) },
          ),
        }),
      ),
  };
}

/**
 * The path an edit made, as PowerPoint writes it, with its point types: the
 * old ones when the edit kept the number of points, else one Auto ("A") a
 * point (what PowerPoint writes for each type waits on truth). Null when
 * nothing is left of it.
 */
export function motionPathFromEdit(
  edit: OfficeEditPath,
  frame: MotionPathEditFrame,
  previous: { path: string; ptsTypes: string },
): { path: string; ptsTypes: string } | null {
  if (!frameUsable(frame)) return null;
  const figures = edit.figures.filter((figure) => figure.segments.length > 0);
  if (figures.length === 0) return null;
  const ox = frame.centreXEmu / frame.slideWidthEmu;
  const oy = frame.centreYEmu / frame.slideHeightEmu;
  const share = (point: { x: number; y: number }): [number, number] => [
    point.x / frame.hostWidthPx - ox,
    point.y / frame.hostHeightPx - oy,
  ];
  const ended = parseNativeMotionPath(previous.path)?.ended ?? true;
  const path = formatNativeMotionPath({
    ended,
    subpaths: figures.map((figure) => ({
      start: share(figure.start),
      closed: figure.closed,
      segments: figure.segments.map((segment) =>
        segment.kind === "line"
          ? { kind: "L" as const, to: share(segment.to) }
          : { kind: "C" as const, c1: share(segment.c1), c2: share(segment.c2), to: share(segment.to) },
      ),
    })),
  });
  const points = figures.reduce((count, figure) => count + 1 + figure.segments.length, 0);
  return { path, ptsTypes: previous.ptsTypes.length === points ? previous.ptsTypes : "A".repeat(points) };
}

/**
 * Motion Paths > Custom Path (GAP #10, step 3): a path drawn with the
 * editors' shared Curve / Freeform / Scribble tools (`mountOfficeFreeformDraw`,
 * its outline in a box on the slide, EMU) as PowerPoint writes it, measured
 * from the object's centre; a drawn path ends where the drawing ended (`E`).
 * Null when the drawing holds no path.
 */
export function motionPathFromDrawing(
  drawing: { left: number; top: number; width: number; height: number; customPath: unknown },
  frame: MotionPathEditFrame,
): { path: string; ptsTypes: string } | null {
  if (!frameUsable(frame)) return null;
  const pxPerEmuX = frame.hostWidthPx / frame.slideWidthEmu;
  const pxPerEmuY = frame.hostHeightPx / frame.slideHeightEmu;
  const box = officeEditPathFromCustomPath(
    drawing.customPath,
    Math.max(drawing.width, 1) * pxPerEmuX,
    Math.max(drawing.height, 1) * pxPerEmuY,
  );
  if (!box) return null;
  const dx = drawing.left * pxPerEmuX;
  const dy = drawing.top * pxPerEmuY;
  const move = (point: { x: number; y: number }) => ({ x: point.x + dx, y: point.y + dy });
  const onSlide: OfficeEditPath = {
    figures: box.figures.map((figure) => ({
      closed: figure.closed,
      start: move(figure.start),
      segments: figure.segments.map((segment) =>
        segment.kind === "line"
          ? { kind: "line" as const, to: move(segment.to) }
          : { kind: "cubic" as const, c1: move(segment.c1), c2: move(segment.c2), to: move(segment.to) },
      ),
    })),
  };
  return motionPathFromEdit(onSlide, frame, { path: "M 0 0 E", ptsTypes: "" });
}
