"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The selected effect's motion path drawn over the slide, as PowerPoint
 * shows it while the effect is selected (a dotted line from a green start
 * mark to a red end mark; a path played backwards, Reverse Path Direction,
 * starts at its last point, so the marks change ends), and Edit Points on it through the editors' shared
 * overlay, and Path Animation > Custom Path drawn with the editors' shared
 * Freeform overlay (GAP #10). The path is laid out in the slide's frame from
 * its object's centre (`lib/presentation-motion-path-edit`, the adapter).
 */

import { useRef } from "react";

import { OfficeDrawingEditPoints, OfficeFreeformDrawLayer } from "@mrnote/core/web/components/office-ribbon/OfficeDrawingSelection.tsx";
import type { OfficeEditPointsLabels } from "@mrnote/core/web/lib/office-drawing-overlays.ts";
import {
  motionPathEditPath,
  motionPathFromDrawing,
  motionPathFromEdit,
  type MotionPathEditFrame,
} from "../../../../lib/presentation-motion-path-edit";
import { nativeMotionPathPoints } from "@mrnote/core/web/lib/presentation-motion-path-native.ts";

/** EMU per CSS pixel at 100 %: the unit the deck's overlays lay out in. */
const EMU_PER_PX = 9525;

export interface PresentationMotionPathLayerProps {
  readonly slideWidthEmu: number;
  readonly slideHeightEmu: number;
  /** The object's centre on the slide, EMU. */
  readonly centreXEmu: number;
  readonly centreYEmu: number;
  readonly path: string;
  readonly ptsTypes: string;
  /** Reverse Path Direction: the object travels the path from its end. */
  readonly reversed?: boolean;
  readonly editing: boolean;
  readonly labels: OfficeEditPointsLabels;
  readonly onCommit: (next: { path: string; ptsTypes: string }) => void;
  readonly onExit: () => void;
}

export function PresentationMotionPathLayer(props: PresentationMotionPathLayerProps) {
  const { slideWidthEmu: W, slideHeightEmu: H, centreXEmu, centreYEmu } = props;
  const points = nativeMotionPathPoints(props.path);
  if (!points || points.length === 0 || !(W > 0) || !(H > 0)) return null;
  const frame: MotionPathEditFrame = {
    centreXEmu,
    centreYEmu,
    slideWidthEmu: W,
    slideHeightEmu: H,
    hostWidthPx: W / EMU_PER_PX,
    hostHeightPx: H / EMU_PER_PX,
  };
  const at = ([x, y]: readonly [number, number]) => [centreXEmu + x * W, centreYEmu + y * H] as const;
  const drawn = points.map(at);
  const [sx, sy] = props.reversed ? drawn[drawn.length - 1] : drawn[0];
  const [ex, ey] = props.reversed ? drawn[0] : drawn[drawn.length - 1];
  const mark = Math.min(W, H) / 60;
  if (props.editing) {
    const editPath = motionPathEditPath(props.path, frame);
    if (editPath) {
      return (
        <div data-testid="presentation-motion-path-edit-host" style={{ position: "absolute", inset: 0, zIndex: 41 }}>
          <OfficeDrawingEditPoints
            width={frame.hostWidthPx}
            height={frame.hostHeightPx}
            path={editPath}
            labels={props.labels}
            testIdPrefix="presentation-motion-path-edit-points"
            onCommit={(next) => {
              const written = motionPathFromEdit(next, frame, { path: props.path, ptsTypes: props.ptsTypes });
              if (written) props.onCommit(written);
            }}
            onExit={props.onExit}
          />
        </div>
      );
    }
  }
  return (
    <svg
      data-testid="presentation-motion-path-layer"
      viewBox={`0 0 ${W} ${H}`}
      preserveAspectRatio="none"
      aria-hidden="true"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%", overflow: "visible", pointerEvents: "none", zIndex: 38 }}
    >
      <polyline
        points={drawn.map(([x, y]) => `${x},${y}`).join(" ")}
        fill="none"
        stroke="#7f7f7f"
        strokeWidth={1.25}
        strokeDasharray="4 3"
        vectorEffect="non-scaling-stroke"
      />
      <polygon
        data-testid="presentation-motion-path-start"
        points={`${sx - mark},${sy - mark} ${sx + mark},${sy} ${sx - mark},${sy + mark}`}
        fill="#00b050"
      />
      <polygon
        data-testid="presentation-motion-path-end"
        points={`${ex - mark},${ey - mark} ${ex + mark},${ey} ${ex - mark},${ey + mark}`}
        fill="#c00000"
      />
    </svg>
  );
}

export interface PresentationMotionPathDrawLayerProps {
  readonly slideWidthEmu: number;
  readonly slideHeightEmu: number;
  readonly centreXEmu: number;
  readonly centreYEmu: number;
  readonly onComplete: (drawn: { path: string; ptsTypes: string }) => void;
  readonly onCancel: () => void;
}

/**
 * Path Animation > Custom Path: the shared Freeform overlay over the slide
 * (click for a straight piece, drag for a drawn one, double-click, Enter or
 * Esc to end -- PowerPoint's Esc keeps the path), the drawing turned into a
 * path measured from the object's centre.
 */
export function PresentationMotionPathDrawLayer(props: PresentationMotionPathDrawLayerProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const { slideWidthEmu: W, slideHeightEmu: H } = props;
  return (
    <div
      ref={hostRef}
      data-testid="presentation-motion-path-draw-host"
      style={{ position: "absolute", inset: 0, zIndex: 41 }}
    >
      <OfficeFreeformDrawLayer
        tool="freeform"
        escapeFinishes
        testId="presentation-motion-path-draw"
        onComplete={(result) => {
          // The overlay reports the drawing in its own on-screen pixels (no
          // scale given); the slide's on-screen size turns them into EMU.
          const rect = hostRef.current?.getBoundingClientRect();
          const widthPx = rect && rect.width > 0 ? rect.width : W / 9525;
          const heightPx = rect && rect.height > 0 ? rect.height : H / 9525;
          const drawn = motionPathFromDrawing(
            {
              left: (result.left * W) / widthPx,
              top: (result.top * H) / heightPx,
              width: (result.width * W) / widthPx,
              height: (result.height * H) / heightPx,
              customPath: result.customPath,
            },
            {
              centreXEmu: props.centreXEmu,
              centreYEmu: props.centreYEmu,
              slideWidthEmu: W,
              slideHeightEmu: H,
              hostWidthPx: W / 9525,
              hostHeightPx: H / 9525,
            },
          );
          if (drawn) props.onComplete(drawn);
          else props.onCancel();
        }}
        onCancel={props.onCancel}
      />
    </div>
  );
}
