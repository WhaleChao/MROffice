/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { z } from "zod";
import {
  presentationChartIsImported,
  presentationChartIsImportedV2,
} from "./presentation-chart-v2";
import { chartAreaStatesFill, excelBuiltInChartTextLook, withOfficeAutomaticChartLine } from "./spreadsheet/chart-builtin-text";
import { presentationNativeMapLook } from "./presentation-native-map-look";
import {
  normalizeChartNativeLook,
  normalizeChartNativeLookBuiltInStyle,
  type ChartNativeLook,
} from "./office-chart-native-look";
import type { ChartBlockAttrsV2 } from "../components/console/editor/extensions/chartBlockShared";
import type { PresentationChartBlock } from "../components/console/editor/native-document/documentCommands";

const paint = z.string().regex(/^(#[a-fA-F0-9]{6}|none)$/);
const axis = z.object({
  min: z.number().finite().optional(),
  max: z.number().finite().optional(),
  majorUnit: z.number().finite().positive().optional(),
  numberFormat: z.string().max(100).optional(),
  fontSize: z.number().positive().max(400).optional(),
  fontFamily: z.string().max(100).optional(),
  // Format Axis > "Values/Categories in reverse order" (`c:orientation
  // val="maxMin"`): the axis runs from its maximum.
  reversed: z.literal(true).optional(),
});
const displaySchema = z.object({
  version: z.literal(1),
  background: paint.optional(),
  plotBackground: paint.optional(),
  valueAxis: axis.optional(),
  categoryAxis: axis.optional(),
  blanks: z.enum(["gap", "span", "zero"]).optional(),
  series: z
    .array(
      z.object({
        name: z.string().max(500),
        color: paint.optional(),
        lineColor: paint.optional(),
        lineWidth: z.number().nonnegative().max(100).optional(),
        marker: z
          .enum(["circle", "square", "diamond", "triangle", "none"])
          .optional(),
        markerSize: z.number().positive().max(100).optional(),
      }),
    )
    .max(12),
});
export type PresentationChartDisplay = z.infer<typeof displaySchema>;

/**
 * A stored native appearance (`source.chartDisplay`) as the schema admits it.
 * The appearance is read from the chart part on the server (the import, and
 * `GET /pages/{id}/presentation-chart-displays` for a deck imported before:
 * `app/services/file_adapters/pptx/chart_display.py`); the browser only
 * validates it.
 */
export function parsePresentationChartDisplay(
  value: unknown,
): PresentationChartDisplay | undefined {
  const parsed = displaySchema.safeParse(value);
  return parsed.success ? parsed.data : undefined;
}

// An imported Office 2016 chart's look rides on its block (`nativeChartLook`,
// read from its style and colour parts). A page imported before that reading
// still says Office's default: a treemap's branch names overlap its tiles. A
// chart this editor authored keeps the banner its writer states.
const IMPORTED_CHART_EX_LOOK: ChartNativeLook = { parentLabelLayout: "overlapping" };

/** The look an imported chart's parts give it, for a host drawing it outside
 * the slide (a ribbon's style tiles draw the same picture the slide draws). */
export function presentationImportedChartLook(
  block: PresentationChartBlock,
  attrs: ChartBlockAttrsV2,
): ChartNativeLook | undefined {
  return block.chartV2 === undefined || presentationChartIsImportedV2(block)
    ? importedChartExLook(block, attrs)
    : undefined;
}

export function importedChartExLook(block: PresentationChartBlock, attrs: ChartBlockAttrsV2): ChartNativeLook {
  const stored = normalizeChartNativeLook(block.source?.nativeChartLook);
  const look: ChartNativeLook = { ...IMPORTED_CHART_EX_LOOK, ...stored };
  if (attrs.chartType !== "map" || look.map) return look;
  // A map the profile drew keeps its view in the profile's own metadata.
  const profile = presentationNativeMapLook(block);
  return profile ? { ...profile, ...look, map: profile.map } : look;
}




const BUILT_IN_STYLE_FRAME = /* own data: BUILT_IN_STYLE_FRAME */ "#7b838e";





















export function presentationBuiltInStyleLook(
  block: PresentationChartBlock,
  attrs: ChartBlockAttrsV2,
): ChartNativeLook | undefined {
  if (!presentationChartIsImported(block)) return undefined;
  if (attrs.stylePreset !== undefined && attrs.stylePreset !== "classic") return undefined;
  const source = block.source as { builtInStyleLook?: unknown; legacyChartText?: unknown } | undefined;
  const hex = (value: unknown) => (typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : undefined);
  const recordOf = (value: unknown) =>
    typeof value === "object" && value !== null && !Array.isArray(value) ? (value as Record<string, unknown>) : undefined;
  const record = recordOf(source?.builtInStyleLook);
  const style = typeof record?.style === "number" && Number.isInteger(record.style) && record.style >= 1 && record.style <= 48
    ? record.style
    : 0;
  const ink = hex(recordOf(source?.legacyChartText)?.ink);
  if (!ink && !style) return undefined;
  const format = attrs.format as Record<string, unknown> | undefined;
  let look: ChartNativeLook = {};
  if (ink) {
    const { chartArea: _excelFrame, ...base } = excelBuiltInChartTextLook(format, style ? record : undefined) ?? {};
    look = base;
    if (ink !== "#000000") {
      // The built-in look's black is the part's tx1 (a dark slide maps it light).
      for (const key of ["title", "axisTitle", "axis", "legend", "dataLabel"] as const) {
        const entry = look[key];
        if (entry && entry.color === "#000000") look = { ...look, [key]: { ...entry, color: ink } };
      }
    }
    // The automatic axis lines, gridlines and tick marks PowerPoint draws
    // otherwise than #7b838e (`legacyChartText.line`: black in a deck of
    // AppVersion 14+ for a part stating no c14:style, #898989 under Office
    // 2013's theme; `withOfficeAutomaticChartLine`).
    look = withOfficeAutomaticChartLine(look, hex(recordOf(source?.legacyChartText)?.line));
  }
  const framed = style >= 33 && style <= 40;
  const builtInStyle = style ? normalizeChartNativeLookBuiltInStyle(record) : undefined;
  if (builtInStyle) {
    look.builtInStyle = builtInStyle;
    look.builtInStylePaint = true;
  }
  if (style && !chartAreaStatesFill(format?.chartAreaStyle)) {
    const area = hex(record?.chartAreaFill);
    if (area) look.chartArea = { fill: area };
    else if (framed) look.chartArea = { fill: "#ffffff", outline: BUILT_IN_STYLE_FRAME, outlineWidthPt: 0.75 };
  }
  const plot = hex(record?.plotAreaFill);
  if (plot && !chartAreaStatesFill(format?.plotAreaStyle)) {
    look.plotArea = framed ? { fill: plot, outline: BUILT_IN_STYLE_FRAME, outlineWidthPt: 0.75 } : { fill: plot };
  }
  const styleInk = hex(record?.textColor);
  if (styleInk && !ink) {
    look.title = { ...look.title, color: styleInk };
    look.axisTitle = { ...look.axisTitle, color: styleInk };
    look.axis = { ...look.axis, color: styleInk };
    look.legend = { ...look.legend, color: styleInk };
    look.dataLabel = { ...look.dataLabel, color: styleInk };
  }
  return Object.keys(look).length ? look : undefined;
}

/** The schemas the deck plan reads a stored display with (server side). */
export { displaySchema as PRESENTATION_CHART_DISPLAY_SCHEMA, axis as PRESENTATION_CHART_AXIS_SCHEMA };
