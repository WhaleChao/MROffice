"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useId, type ReactNode } from "react";

import { normalizePictureClipShape, pictureClipShapeOutline } from "../../../../lib/presentation-picture-crop-shape";
import { PRESENTATION_CUSTOM_PATH_GRID, presentationCustomPathToSvgPath } from "@mrnote/core/web/lib/presentation-freeform.ts";
import { officeLineFamilyDash } from "@mrnote/core/web/lib/office-stroke-dash.ts";

/** A border drawn along the shape: design px, its colour and its dash. */
export interface ClipShapeStroke {
  width: number;
  color: string;
  style: "solid" | "dashed" | "dotted";
}













export function PresentationClipShapeFrame({
  clipShape,
  width,
  height,
  fillParent,
  includeEllipse = false,
  stroke,
  adjustments = null,
  customPath = null,
  children,
}: {
  clipShape: unknown;
  /** The crop shape's dragged handles (`clipShapeAdjustments`). */
  adjustments?: unknown;
  /** A freeform outline (`clipCustomPath`, the shapes' path model). */
  customPath?: unknown;
  width: unknown;
  height: unknown;
  fillParent?: boolean;
  includeEllipse?: boolean;
  stroke?: ClipShapeStroke | null;
  children: ReactNode;
}) {
  const reactId = useId();
  const preset = normalizePictureClipShape(clipShape);
  const widthEmu = Number(width) || 0;
  const heightEmu = Number(height) || 0;
  const customD = !preset && customPath ? presentationCustomPathToSvgPath(customPath) : null;
  // The path is on its own 100000 grid; the outline is that grid stretched
  // to the frame, in design px (the stroke's unit).
  const outline =
    customD && widthEmu > 0 && heightEmu > 0
      ? {
          // `presentationCustomPathToSvgPath` writes the grid in thousandths.
          d: scaleSvgPath(
            customD,
            (widthEmu / 9525) / (PRESENTATION_CUSTOM_PATH_GRID / 1000),
            (heightEmu / 9525) / (PRESENTATION_CUSTOM_PATH_GRID / 1000),
          ),
          width: widthEmu / 9525,
          height: heightEmu / 9525,
        }
      : preset && (includeEllipse || stroke || preset !== "ellipse")
      ? pictureClipShapeOutline(preset, widthEmu, heightEmu, adjustments)
      : null;
  if (!outline) return <>{children}</>;
  const domId = `presentation-clip-shape-${reactId.replace(/[^A-Za-z0-9_-]/g, "")}`;
  const clipDefs = (
    <svg aria-hidden="true" width="0" height="0" style={{ position: "absolute" }}>
      <defs>
        <clipPath id={domId} clipPathUnits="objectBoundingBox">
          <path d={outline.d} transform={`scale(${1 / outline.width} ${1 / outline.height})`} />
        </clipPath>
      </defs>
    </svg>
  );
  const frameStyle = {
    position: "relative" as const,
    display: fillParent ? "block" : "inline-block",
    ...(fillParent ? { width: "100%", height: "100%" } : {}),
  };
  // The outline is in design px, the stroke's unit.
  const line = stroke && stroke.width > 0 ? outline : null;
  if (!line || !stroke) {
    return (
      <span
        data-testid="presentation-picture-clip-shape"
        data-clip-shape={preset ?? undefined}
        style={{ ...frameStyle, clipPath: `url(#${domId})` }}
      >
        {clipDefs}
        {children}
      </span>
    );
  }
  return (
    <span data-testid="presentation-picture-clip-shape-frame" style={frameStyle}>
      <span
        data-testid="presentation-picture-clip-shape"
        data-clip-shape={preset ?? undefined}
        style={{ ...frameStyle, display: "block", width: "100%", height: "100%", clipPath: `url(#${domId})` }}
      >
        {clipDefs}
        {children}
      </span>
      <svg
        aria-hidden="true"
        data-testid="presentation-picture-clip-shape-line"
        viewBox={`0 0 ${line.width} ${line.height}`}
        preserveAspectRatio="none"
        overflow="visible"
        style={{ position: "absolute", inset: 0, width: "100%", height: "100%", pointerEvents: "none" }}
      >
        <path
          d={line.d}
          fill="none"
          stroke={stroke.color}
          strokeWidth={stroke.width}
          strokeLinejoin="round"
          
          strokeDasharray={officeLineFamilyDash(stroke.style, stroke.width)?.dashArray}
          strokeLinecap={officeLineFamilyDash(stroke.style, stroke.width)?.linecap}
        />
      </svg>
    </span>
  );
}

/**
 * An SVG path stretched by `sx`, `sy`: every coordinate pair, and an arc's
 * radii (its x-axis rotation is 0 in the paths `presentationCustomPathToSvgPath`
 * writes, so a stretch keeps it an axis-aligned arc).
 */
function scaleSvgPath(d: string, sx: number, sy: number): string {
  const tokens = d.match(/[MLQCAZ]|-?\d*\.?\d+(?:e[-+]?\d+)?/gi) ?? [];
  const out: string[] = [];
  let command = "";
  let index = 0;
  for (const token of tokens) {
    if (/^[A-Za-z]$/.test(token)) {
      command = token.toUpperCase();
      index = 0;
      out.push(token);
      continue;
    }
    const value = Number(token);
    let scaled = value;
    if (command === "A") {
      // rx ry rotation large-arc sweep x y
      const slot = index % 7;
      scaled = slot === 0 || slot === 5 ? value * sx : slot === 1 || slot === 6 ? value * sy : value;
    } else {
      scaled = index % 2 === 0 ? value * sx : value * sy;
    }
    index += 1;
    out.push(String(Math.round(scaled * 1000) / 1000));
  }
  return out.join(" ");
}
