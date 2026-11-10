"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type ReactElement } from "react";
import { sparklinePoints, sparklineStrokeWidthPx, type Sparkline } from "@mrnote/core/web/lib/spreadsheet/sparkline.ts";

// Conditional-format icon glyphs: components/notebook/contents/spreadsheet/
// ConditionalIconGlyph (Excel's icon look, XLS-A044/A045/A046).

// Default sparkline colors when the model carries none. The positive/primary
// color is the accent; negatives/losses use a muted red.
export const SPARKLINE_DEFAULT_COLOR = "#0d9488";
export const SPARKLINE_DEFAULT_NEGATIVE = "#dc2626";
/** A line sparkline's data marker, in the glyph's 80 x 18 viewBox. */
const MARKER_RADIUS = 1.7;

/**
 * Render an in-cell sparkline (Excel/WPS parity) as a small inline SVG from the
 * sparkline model + the numeric values it is bound to. Pure presentation: it
 * calls the shared `sparklinePoints` geometry helper and draws the result.
 * A line draws a polyline; a column/win-loss draws mini bars (negatives/losses in
 * the negative color). The SVG scales to the cell via 100% width/height.
 */
export function SparklineGlyph({
  sparkline,
  values,
  axisBounds,
  scale = 1,
}: {
  sparkline: Sparkline;
  values: Array<number | null>;
  axisBounds?: { minimum?: number; maximum?: number };
  /** The surface's own magnification (Page Layout's page scale). */
  scale?: number;
}): ReactElement {
  // Fixed internal coordinate space; the SVG scales to the cell box. A wider box
  // gives the line/bars room without re-deriving geometry per cell size.
  const W = 80;
  const H = 18;
  // A line's points sit inside the box by a marker's radius, so a marker on
  // the first, last, highest or lowest point shows whole instead of a
  // quarter of it at the box's edge (live audit S-048).
  const INSET = sparkline.type === "line" ? MARKER_RADIUS + 0.3 : 0;
  const color = sparkline.color || SPARKLINE_DEFAULT_COLOR;
  const negColor = sparkline.negativeColor || SPARKLINE_DEFAULT_NEGATIVE;
  const model = sparklinePoints(values, sparkline.type, W - 2 * INSET, H - 2 * INSET, {
    displayEmptyCellsAs: sparkline.displayEmptyCellsAs,
    includeZero: sparkline.displayXAxis,
    minimum: axisBounds?.minimum,
    maximum: axisBounds?.maximum,
    customMinimum: sparkline.minAxisType === "custom",
    customMaximum: sparkline.maxAxisType === "custom",
  });
  const plottedValues =
    model.type === "line" ? model.pointValues : model.barValues;
  const minimum = plottedValues.length ? Math.min(...plottedValues) : 0;
  const maximum = plottedValues.length ? Math.max(...plottedValues) : 0;
  // Excel's precedence when a point is several of these (Excel for Mac
  // measured, X-198 SP_priority): high, then low, first, last, negative,
  // then the plain marker -- a first point that is also the lowest wears the
  // low colour, a negative first or last point its first/last colour.
  const pointColor = (value: number, ordinal: number): string | null => {
    if (sparkline.high && value === maximum)
      return sparkline.highColor || color;
    if (sparkline.low && value === minimum)
      return sparkline.lowColor || color;
    if (sparkline.first && ordinal === 0)
      return sparkline.firstColor || color;
    if (sparkline.last && ordinal === plottedValues.length - 1)
      return sparkline.lastColor || color;
    if (sparkline.negative && value < 0) return negColor;
    if (sparkline.markers) return sparkline.markerColor || color;
    return null;
  };
  return (
    <svg
      className="spreadsheet-window__sparkline"
      width="100%"
      height="100%"
      viewBox={`0 0 ${W} ${H}`}
      preserveAspectRatio="none"
      aria-hidden="true"
      focusable="false"
      data-testid="spreadsheet-sparkline-svg"
      data-sparkline-type={sparkline.type}
    >
      {sparkline.displayXAxis && model.zeroY !== null ? (
        <line
          x1="0"
          y1={model.zeroY + INSET}
          x2={W}
          y2={model.zeroY + INSET}
          stroke={sparkline.axisColor || "#94a3b8"}
          strokeWidth="0.8"
          vectorEffect="non-scaling-stroke"
          data-testid="spreadsheet-sparkline-axis"
        />
      ) : null}
      {model.type === "line" ? (
        <g transform={INSET ? `translate(${INSET} ${INSET})` : undefined}>
          {model.segments.map((segment, index) =>
            segment.length > 1 ? (
              <polyline
                key={`segment-${index}`}
                points={segment.map((p) => `${p.x},${p.y}`).join(" ")}
                fill="none"
                stroke={color}
                // The group's lineWeight (points), drawn at its size in CSS
                // px whatever the cell's box stretches the viewBox to.
                strokeWidth={sparklineStrokeWidthPx(sparkline, scale)}
                strokeLinejoin="round"
                strokeLinecap="round"
                vectorEffect="non-scaling-stroke"
              />
            ) : null,
          )}
          {model.points.map((point, index) => {
            const markerColor = pointColor(model.pointValues[index], index);
            if (!markerColor && model.points.length !== 1) return null;
            // A marker is a dot of its own CSS size, about three line
            // weights across as Excel draws it (S21_Sparklines, X-198: a
            // 1 pt line takes ~3.5 pt markers): a zero-length round-capped
            // stroke that keeps its width and stays round however the cell
            // stretches the view box (a circle in view-box units became a
            // tall oval as large as the cell was high).
            return (
              <line
                key={`point-${model.pointIndexes[index]}`}
                x1={point.x}
                y1={point.y}
                x2={point.x}
                y2={point.y}
                stroke={markerColor || color}
                strokeWidth={sparklineStrokeWidthPx(sparkline, scale) * (markerColor ? 3 : 2)}
                strokeLinecap="round"
                vectorEffect="non-scaling-stroke"
                data-testid={
                  markerColor ? "spreadsheet-sparkline-marker" : undefined
                }
              />
            );
          })}
        </g>
      ) : (
        model.bars.map((bar, i) => (
          <rect
            key={i}
            x={bar.x}
            y={bar.y}
            width={bar.width}
            height={bar.height}
            fill={
              pointColor(model.barValues[i], i) ||
              (sparkline.negative && bar.positive === false ? negColor : color)
            }
          />
        ))
      )}
    </svg>
  );
}
