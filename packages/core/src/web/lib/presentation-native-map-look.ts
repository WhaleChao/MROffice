/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A deck chart's native map view (the server's `nativeMap` profile on the
 * block's source), as a look: model reading of stored metadata, no drawing.
 */
import { z } from "zod";
import type { PresentationChartBlock } from "../components/console/editor/native-document/documentCommands";
import { normalizeChartNativeLook, type ChartNativeLook } from "./office-chart-native-look";

const color = z.string().regex(/^#[a-fA-F0-9]{6}$/);
export const NATIVE_MAP_VIEW_SCHEMA = z.object({
  version: z.literal(1),
  projection: z.literal("native-africa-albers-auto-v1"),
  parentBounds: z.tuple([z.number(), z.number(), z.number(), z.number()]),
  featuresBase64: z.string().max(2000000),
  featureBytes: z.number().int().min(1).max(4000000),
  bindings: z
    .array(
      z.object({
        id: z.string().regex(/^\d{3}$/),
        entityId: z.string().max(20),
      }),
    )
    .min(1)
    .max(50),
  cacheFingerprint: z.string().regex(/^[a-f0-9]{64}$/),
  mapName: z.string().regex(/^mrai-native-map-[a-f0-9]{64}$/),
  fontFamily: z.string().max(100),
  fontSize: z.number().positive().max(100),
  legendFontSize: z.number().positive().max(100),
  fontColor: color,
  borderColor: color,
  borderWidthPt: z.number().min(0).max(24),
  lowColor: color,
  highColor: color,
  attribution: z.string().max(300),
  copyrights: z.array(z.string().max(200)).max(100),
  seriesName: z.string().max(200).nullable(),
});
const record = (value: unknown): Record<string, unknown> =>
  value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};

/**
 * The profile's view as a look, for the shared planner: a map whose title or
 * legend was edited leaves this profile, and keeps the continent, the colour
 * scale and the best-fit names this profile drew. The profile only takes a
 * part that names its regions best-fit.
 */
export function presentationNativeMapLook(
  block: PresentationChartBlock,
): ChartNativeLook | undefined {
  const parsed = NATIVE_MAP_VIEW_SCHEMA.safeParse(
    record(block.source?.nativeChartView).nativeMap,
  );
  if (!parsed.success) return undefined;
  const native = parsed.data;
  return normalizeChartNativeLook({
    map: {
      labelLayout: "bestFitOnly",
      // The profile only takes a right-hand scale at the top of its side.
      legendAlign: "min",
      projection: native.projection,
      parentBounds: native.parentBounds,
      featuresBase64: native.featuresBase64,
      featureBytes: native.featureBytes,
      bindings: native.bindings,
      mapName: native.mapName,
      attribution: native.attribution,
      copyrights: native.copyrights,
      fontFamily: native.fontFamily,
      seriesName: native.seriesName ?? undefined,
      scale: { low: native.lowColor, high: native.highColor },
    },
    dataLabel: { sizePt: native.fontSize, color: native.fontColor },
    legend: { sizePt: native.legendFontSize, color: native.fontColor },
    points: [
      { outline: native.borderColor, outlineWidthPt: native.borderWidthPt },
    ],
  });
}

