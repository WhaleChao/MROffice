"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { SparklineType } from "@mrnote/core/web/lib/spreadsheet/sparkline.ts";
import { sparklineStylePatch, type SparklineStyle } from "../../../../lib/spreadsheet/sparkline-styles";

/** A fixed little series with a clear high, low, first, last and a dip below
 *  zero, so every colour of a style shows. */
const SAMPLE: readonly number[] = [3, 5, 2, 6, -2, 4, 7, 1, 5];

/**
 * One tile of the Sparkline Style gallery: the active group's kind (line,
 * column, win/loss) drawn in the style's colours, resolved in the workbook's
 * theme, as Excel's gallery draws its tiles.
 */
export function SpreadsheetSparklineStyleSwatch({
  style,
  type,
  paletteId,
  width = 44,
  height = 28,
}: {
  readonly style: SparklineStyle;
  readonly type: SparklineType;
  readonly paletteId: string;
  readonly width?: number;
  readonly height?: number;
}) {
  const colors = sparklineStylePatch(style, paletteId);
  const pad = 3;
  const n = SAMPLE.length;
  const max = Math.max(...SAMPLE);
  const min = Math.min(...SAMPLE);
  const high = SAMPLE.indexOf(max);
  const low = SAMPLE.indexOf(min);
  const pointColor = (index: number, value: number): string =>
    index === high
      ? colors.highColor
      : index === low
        ? colors.lowColor
        : index === 0
          ? colors.firstColor
          : index === n - 1
            ? colors.lastColor
            : value < 0
              ? colors.negativeColor
              : colors.color;
  const x = (index: number) => pad + ((width - pad * 2) * index) / (n - 1);
  const y = (value: number) => pad + ((height - pad * 2) * (max - value)) / (max - min);
  if (type === "line") {
    const points = SAMPLE.map((value, index) => `${x(index).toFixed(1)},${y(value).toFixed(1)}`).join(" ");
    return (
      <svg aria-hidden="true" width={width} height={height} viewBox={`0 0 ${width} ${height}`}>
        <polyline points={points} fill="none" stroke={colors.color} strokeWidth={1.25} strokeLinejoin="round" />
        {SAMPLE.map((value, index) => (
          <circle
            key={index}
            cx={x(index)}
            cy={y(value)}
            r={1.6}
            fill={
              index === high || index === low || index === 0 || index === n - 1 || value < 0
                ? pointColor(index, value)
                : colors.markerColor
            }
          />
        ))}
      </svg>
    );
  }
  const slot = (width - pad * 2) / n;
  const bar = Math.max(1.5, slot * 0.7);
  const zero = type === "winloss" ? height / 2 : y(0);
  return (
    <svg aria-hidden="true" width={width} height={height} viewBox={`0 0 ${width} ${height}`}>
      {SAMPLE.map((value, index) => {
        const top =
          type === "winloss" ? (value >= 0 ? pad : zero) : Math.min(y(value), zero);
        const tall =
          type === "winloss" ? height / 2 - pad : Math.max(1, Math.abs(y(value) - zero));
        return (
          <rect
            key={index}
            x={pad + slot * index + (slot - bar) / 2}
            y={top}
            width={bar}
            height={tall}
            fill={pointColor(index, value)}
          />
        );
      })}
    </svg>
  );
}
