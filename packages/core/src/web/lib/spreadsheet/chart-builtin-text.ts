/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeChartNativeLookBuiltInStyle, type ChartNativeLook } from "../office-chart-native-look";

type Text = { fontSizePt?: number; bold?: boolean; color?: string } | undefined;

const EXCEL_BUILT_IN_TEXT_PT = 10;
const EXCEL_BUILT_IN_TITLE_PT = 18;
const EXCEL_BUILT_IN_INK = "#000000";
const EXCEL_BUILT_IN_LINE = /* own data: EXCEL_BUILT_IN_LINE */ "#7b838e";
const EXCEL_BUILT_IN_MINOR_LINE = /* own data: EXCEL_BUILT_IN_MINOR_LINE */ "#bcc1c8";
const EXCEL_AUTOMATIC_MARKERS = ["diamond", "square", "triangle", "x", "star", "circle", "plus", "dot", "dash"] as const;

/** Is this record a classic chart part imported from a workbook? */
export function isImportedClassicChartPart(source: unknown): boolean {
  if (typeof source !== "object" || source === null) return false;
  const part = (source as { part?: unknown }).part;
  return typeof part === "string" && /^xl\/charts\/chart\d+\.xml$/i.test(part);
}

const textOf = (owner: unknown): Text =>
  typeof owner === "object" && owner !== null ? ((owner as { text?: Text }).text ?? undefined) : undefined;

/**
 * The built-in text look for an imported classic chart's format. An element
 * that states its own text (an axis', the legend's, the data labels' txPr)
 * draws what it states over the look; the other elements keep the built-in
 * one (H23_32readwriteBarChart4: yellow value labels, and still the 18pt bold
 * black title; X-072/X-105). Never null; the type keeps the callers' guard.
 */
/** Whether a stored area style states a fill of its own (a colour, none,
 * a gradient or a pattern); one stating a line alone does not. */
export function chartAreaStatesFill(style: unknown): boolean {
  if (typeof style !== "object" || style === null) return false;
  const record = style as { fill?: unknown; shapePaint?: { fillGradient?: unknown; fillPattern?: unknown } };
  return record.fill !== undefined || record.shapePaint?.fillGradient !== undefined || record.shapePaint?.fillPattern !== undefined;
}












export function withOfficeAutomaticChartLine(
  look: ChartNativeLook,
  line: string | undefined,
  frame?: string,
): ChartNativeLook {
  let out = look;
  if (line) {
    if (out.axis?.line) out = { ...out, axis: { ...out.axis, line: { ...out.axis.line, outline: line } } };
    for (const key of ["gridline", "minorGridline"] as const) {
      const entry = out[key];
      if (entry) out = { ...out, [key]: { ...entry, outline: line } };
    }
  }
  if (frame && out.chartArea?.outline) out = { ...out, chartArea: { ...out.chartArea, outline: frame } };
  return out;
}

export function excelBuiltInChartTextLook(
  format: Record<string, unknown> | undefined,
  styleLook?: unknown,
): ChartNativeLook | null {
  const own = (key: string) => textOf(format?.[key]);
  // Office's built-in chart style 9-48 over the workbook theme (the server's
  // `builtInStyleLook`): the paint of unpainted filled series, and for styles
  // 33-48 the plot area, and for 41-48 the dark chart area and light text.
  const style = typeof styleLook === "object" && styleLook !== null ? (styleLook as Record<string, unknown>) : undefined;
  const builtInStyle = normalizeChartNativeLookBuiltInStyle(style);
  const hex = (value: unknown) => (typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : undefined);
  const styleInk = hex(style?.textColor);
  const styleChartArea = hex(style?.chartAreaFill);
  const stylePlotArea = hex(style?.plotAreaFill);
  const chart = (format?.text ?? undefined) as Text;
  const title = own("chartTitleStyle");
  const base = chart?.fontSizePt ?? EXCEL_BUILT_IN_TEXT_PT;
  const ink = chart?.color ?? styleInk ?? EXCEL_BUILT_IN_INK;
  const strong = chart?.bold ?? true;
  const plain = { sizePt: base, color: ink, ...(chart?.bold ? { bold: true } : {}) };
  return {
    builtInText: true,
    title: {
      // A title whose text states properties but no size (a `defRPr` /
      // `rPr` without `sz`, Google Sheets' grey titles) is drawn at the
      // chart's text size, not the built-in title's (Excel for Mac
      
      // alone 10pt, neither 18pt; X-263).
      sizePt:
        title?.fontSizePt ??
        (title && Object.keys(title).length > 0
          ? base
          : chart?.fontSizePt !== undefined
            ? // 1.2 times the chart's text, to the hundredth of a point Excel
              // keeps sizes in (Windows Excel's PDFs, lookwin p159b: 9.6 pt
              // over 8 pt text, 16.8 pt over 14 pt; it stood 9 / 16).
              Math.round(base * 120) / 100
            : EXCEL_BUILT_IN_TITLE_PT),
      bold: title?.bold ?? strong,
      color: title?.color ?? ink,
    },
    // Excel's default for an axis title that states no text of its own;
    // a title that states its own keeps it, axis by axis (the plan's
    // native look draws each title's own text over this).
    axisTitle: {
      sizePt: base,
      bold: strong,
      color: ink,
    },
    axis: { ...plain, line: { outline: EXCEL_BUILT_IN_LINE, outlineWidthPt: 0.75 } },
    gridline: { outline: EXCEL_BUILT_IN_LINE, outlineWidthPt: 0.75 },
    // Minor gridlines that state no paint: #bcc1c8 (Excel for Mac's print
    // of H26_32readwriteLineChart4, its category axis' unpainted minors).
    minorGridline: { outline: EXCEL_BUILT_IN_MINOR_LINE, outlineWidthPt: 0.75 },
    // The frame only round a chart area that states no paint of its own;
    // a dark style (41-48) fills it and draws no frame.
    ...(format?.chartAreaStyle === undefined
      ? {
          chartArea: styleChartArea
            ? { fill: styleChartArea }
            : { outline: EXCEL_BUILT_IN_LINE, outlineWidthPt: 0.75 },
        }
      : styleChartArea && !chartAreaStatesFill(format.chartAreaStyle)
        ? // An area stating a line alone (LibreOffice's) keeps the style's
          // fill under it (matrix L_P20 style 42: black round dark grey).
          { chartArea: { fill: styleChartArea } }
        : !styleChartArea && (format.chartAreaStyle as { outline?: unknown } | null)?.outline === undefined
          ? // An area stating a fill but no line keeps the frame (Excel for
            // Mac, XW414_chart_pattern09's dotted yellow area and
            // XW303_chart_gradient08's gradient, each in the grey frame;
            // X-345).
            { chartArea: { outline: EXCEL_BUILT_IN_LINE, outlineWidthPt: 0.75 } }
          : {}),
    ...(stylePlotArea && !chartAreaStatesFill(format?.plotAreaStyle)
      ? { plotArea: { fill: stylePlotArea } }
      : // A plot area stating nothing inside a filled chart area is white
        // (bg1), not the area's fill showing through (the same two files).
        !stylePlotArea && chartAreaStatesFill(format?.chartAreaStyle) && format?.plotAreaStyle === undefined
        ? { plotArea: { fill: "#ffffff" } }
        : {}),
    ...(builtInStyle ? { builtInStyle } : {}),
    legend: plain,
    dataLabel: plain,
    // A trendline whose part states no line: a solid black 0.75pt line
    // (the same file, measured from Excel's own PDF: the legend key's stroke
    // beside the 0.75pt gridlines).
    trendline: { outline: EXCEL_BUILT_IN_INK, outlineWidthPt: 0.75 },
    // Drop lines, high-low lines, series lines and error bars that state no
    // line: the same black 0.75pt (Excel for Mac, XW014_chart_drop_lines01,
    // XW245's high-low lines, XW017's error bars; X-292).
    helperLine: { outline: EXCEL_BUILT_IN_INK, outlineWidthPt: 0.75 },
    // A marked line or XY series that states no marker: Excel's automatic
    // cycle by series, 7pt (markers.xlsx, nine series, Excel's own PDF).
    markers: { symbols: EXCEL_AUTOMATIC_MARKERS, sizePt: 7 },
  };
}
