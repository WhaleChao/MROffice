/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's Format Axis > Axis Options, ONE record for the three editors.
 *
 * Microsoft Support, "Change the scale of the vertical (value) axis in a
 * chart" and "Change the scale of the horizontal (category) axis in a chart":
 * Bounds (Minimum / Maximum, automatic unless typed), Units (Major / Minor),
 * Horizontal axis crosses (Automatic / Axis value / Maximum axis value),
 * Display units (None, Hundreds .. Trillions, with an optional label),
 * Logarithmic scale (base 10; "cannot be used for negative values or zero"),
 * Values in reverse order, Tick Marks (Major / Minor type: None, Inside,
 * Outside, Cross), Labels (Label position: Next to Axis, High, Low, None),
 * Number (a format code, or linked to source). A category axis has no bounds
 * or units; its reverse order flips the categories.
 *
 * ECMA-376 / ISO 29500 21.2.2.226 `c:valAx` and 21.2.2.25 `c:catAx`:
 * `scaling` (`logBase`, `orientation` minMax|maxMin, `max`, `min`),
 * `numFmt`, `majorTickMark` / `minorTickMark` (cross|in|none|out),
 * `tickLblPos` (high|low|nextTo|none), `crosses` (autoZero|max|min) or
 * `crossesAt`, `majorUnit`, `minorUnit`, `dispUnits` (`builtInUnit`,
 * `dispUnitsLbl`). The Python twin is `app/services/office_chart_axis_options.py`.
 *
 * A category axis adds (Microsoft Support, "Change the scale of the horizontal
 * (category) axis in a chart"): Axis Type (Automatically select based on data,
 * the default; Text axis; Date axis, which "displays dates in chronological
 * order at set intervals or base units ... even if the dates on the worksheet
 * are not in order"), Interval between tick marks ("how many categories are
 * displayed between the tick marks"), Interval between labels (Automatic, or
 * Specify interval unit: 1 labels every category, 2 every other one) and the
 * labels' Distance from axis. In the part (ECMA-376 21.2.2.25 `c:catAx`,
 * 21.2.2.39 `c:dateAx`): `c:auto` (1 = automatic), the element name itself
 * (`c:dateAx` for a date axis), `c:lblOffset` (0..1000, default 100),
 * `c:tickLblSkip` / `c:tickMarkSkip` (at least 1; a text axis only -- a date
 * axis has none) and `c:baseTimeUnit` (days / months / years; date axis only).
 *
 * A text axis over multi-level categories (`c:multiLvlStrRef`, a PivotChart
 * with two or more row fields) adds Labels > "Multi-level Category Labels",
 * checked by default: cleared, Excel draws each category as flat text, its
 * own label over the outer labels that start at it (`c:noMultiLvlLbl`, a
 * `CT_Boolean` whose `val` defaults to true; `c:catAx` only).
 *
 * The Depth (Series) Axis of a 3-D chart that has one (3-D Column, 3-D Line,
 * 3-D Area, 3-D Surface; ECMA-376 21.2.2.175 `c:serAx`) is the record's
 * `series` entry. Microsoft Support, "Change the display of a 3-D chart":
 * Format Axis on the depth axis offers Series in reverse order, Interval
 * between tick marks, Interval between labels, the tick mark types and the
 * label position, and Chart Elements > Axes > Depth Axis shows or hides it.
 * In the part: `c:scaling/c:orientation` (maxMin = reverse), `c:delete`,
 * `c:majorTickMark` / `c:minorTickMark`, `c:tickLblPos`, `c:tickLblSkip` /
 * `c:tickMarkSkip`. The depth axis alone states `hidden` (the other axes are
 * shown or hidden by `format.axisX` / `axisY` / `secondaryValueAxis`).
 */

import DISPLAY_UNIT_LABELS from "./display-unit-labels.json";
import { officeChartLayoutWithAxisTicks } from "./chart-layout";
import {
  EXCEL_1900_FICTITIOUS_LEAP_SERIAL,
  excelSerialToUtcMs,
} from "../spreadsheet/excel-1900-calendar";


export const OFFICE_CHART_TICK_MARKS = ["none", "in", "out", "cross"] as const;
export type OfficeChartTickMark = (typeof OFFICE_CHART_TICK_MARKS)[number];

export const OFFICE_CHART_TICK_LABEL_POSITIONS = ["nextTo", "high", "low", "none"] as const;
export type OfficeChartTickLabelPosition = (typeof OFFICE_CHART_TICK_LABEL_POSITIONS)[number];

export const OFFICE_CHART_DISPLAY_UNITS = [
  "hundreds",
  "thousands",
  "tenThousands",
  "hundredThousands",
  "millions",
  "tenMillions",
  "hundredMillions",
  "billions",
  "trillions",
] as const;
/** An axis's label as it would read without its display units: the room kept. */
export const OFFICE_LABEL_BEFORE_UNITS = "officeLabelBeforeUnits";

export type OfficeChartDisplayUnit = (typeof OFFICE_CHART_DISPLAY_UNITS)[number];

export const OFFICE_CHART_DISPLAY_UNIT_FACTORS: Record<OfficeChartDisplayUnit, number> = {
  hundreds: 1e2,
  thousands: 1e3,
  tenThousands: 1e4,
  hundredThousands: 1e5,
  millions: 1e6,
  tenMillions: 1e7,
  hundredMillions: 1e8,
  billions: 1e9,
  trillions: 1e12,
};

export const OFFICE_CHART_AXIS_CROSSES = ["autoZero", "max", "min"] as const;
export type OfficeChartAxisCrosses = (typeof OFFICE_CHART_AXIS_CROSSES)[number];

export const OFFICE_CHART_AXIS_TYPES = ["text", "date"] as const;
export type OfficeChartAxisType = (typeof OFFICE_CHART_AXIS_TYPES)[number];

export const OFFICE_CHART_BASE_TIME_UNITS = ["days", "months", "years"] as const;
export type OfficeChartBaseTimeUnit = (typeof OFFICE_CHART_BASE_TIME_UNITS)[number];

/** Office's bound on an interval between labels / tick marks (CT_Skip). */
export const OFFICE_CHART_MAX_SKIP = 31999;

/** On a category axis's `axisLabel`: the names are level by a stated Custom
 * angle of 0, so the host wraps them and never turns them. */
export const OFFICE_STATED_LEVEL_LABELS = "officeStatedLevelLabels";

/** On a day Date axis whose major unit is automatic (`true`) or a stated
 * whole number of days (that number): the host lays its labels out as
 * Office lays out a date axis's, a label a unit while they fit level, else
 * turned 45 or 90 degrees before any is skipped (X-340 / X-128: XW221's
 * eight days and XW008's ten days by a stated 1 day, at 45 degrees). */
export const OFFICE_AUTOMATIC_DAY_LABELS = "officeAutomaticDayLabels";





export const OFFICE_DATE_AXIS_FIRST = "officeDateAxisFirst";

export interface OfficeChartAxisOptions {
  /** Bounds; absent = automatic. */
  readonly min?: number;
  readonly max?: number;
  /** Units; absent = automatic. */
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  /** Logarithmic scale: the base (Office offers 10 by default, 2..1000). */
  readonly logBase?: number;
  /** Values (or categories) in reverse order: `c:orientation val="maxMin"`. */
  readonly reverse?: boolean;
  readonly majorTickMark?: OfficeChartTickMark;
  readonly minorTickMark?: OfficeChartTickMark;
  readonly labelPosition?: OfficeChartTickLabelPosition;
  /** A number format code (`c:numFmt@formatCode`); absent = linked to source. */
  readonly numberFormat?: string;
  readonly displayUnits?: OfficeChartDisplayUnit;
  readonly showDisplayUnitsLabel?: boolean;
  /** Where the OTHER axis crosses this one: a rule, or a value on this axis. */
  readonly crosses?: OfficeChartAxisCrosses | number;
  /** Category axis: Axis Type; absent = Automatically select based on data. */
  readonly axisType?: OfficeChartAxisType;
  /** Date axis: Base unit; absent = automatic. */
  readonly baseTimeUnit?: OfficeChartBaseTimeUnit;
  /** Date axis: the time unit `majorUnit` / `minorUnit` count (Format Axis >
   * Units: Major 3 Months; `c:majorTimeUnit` / `c:minorTimeUnit`); absent,
   * the base unit. */
  readonly majorTimeUnit?: OfficeChartBaseTimeUnit;
  readonly minorTimeUnit?: OfficeChartBaseTimeUnit;
  /** Category (text) axis: Interval between labels; absent = Automatic. */
  readonly tickLabelSkip?: number;
  /** Category (text) axis: Interval between tick marks; absent = 1. */
  readonly tickMarkSkip?: number;
  /** Category axis: the labels' Distance from axis (0..1000; absent = 100). */
  readonly labelOffset?: number;
  /** Label position in its band (`c:lblAlgn`): left / right; absent centred. */
  readonly labelAlign?: "l" | "r";
  /**
   * Category (text) axis: Labels > Multi-level Category Labels. Absent =
   * checked, Office's default (`c:noMultiLvlLbl val="0"`); only `false` (flat
   * labels, `c:noMultiLvlLbl val="1"`) is stated. Moot on a single-level axis.
   */
  readonly multiLevelLabels?: false;
  /**
   * Category axis: Axis position. "onTickMarks" draws each category ON its
   * tick mark, the plot running from the first category to the last (the
   * value axis crossing it states `c:crossBetween val="midCat"`); absent is
   * the chart's own default (between tick marks; an area chart's on them, as
   * Office inserts one), and "between" says so of an area chart
   * (X30_chart_axis32's XlsxWriter area chart, half a band at each end).
   */
  readonly axisPosition?: "onTickMarks" | "between";
  /**
   * Format Axis > Size & Properties > Alignment > Custom angle of the tick
   * labels, in whole degrees -90..90, clockwise (`c:txPr/a:bodyPr@rot` /
   * 60000; `vert="vert"` reads 90, `vert270` -90). Absent is automatic:
   * Office's own `rot="-60000000"`, or none stated.
   */
  readonly labelAngle?: number;
  /**
   * Format Axis Title > Alignment > Custom angle of the axis's title, as
   * above (`c:title/c:tx/c:rich/a:bodyPr@rot`, else `c:title/c:txPr`).
   * Absent is Office's default: a vertical axis's title reads bottom to top
   * (-90), a horizontal one's level (0).
   */
  readonly titleAngle?: number;
  /**
   * Format Axis > Alignment > Text direction > Stacked of the tick labels
   * (`c:txPr/a:bodyPr@vert="wordArtVert"`): each character under the one
   * before, level ("10" reads 1 over 0). It replaces any `labelAngle`.
   */
  readonly labelStacked?: true;
  /** Text direction Stacked of the axis's title, as above; replaces `titleAngle`. */
  readonly titleStacked?: true;
  /**
   * The depth axis only: Chart Elements > Axes > Depth Axis cleared (its
   * `c:delete val="1"`); absent, the axis is shown.
   */
  readonly hidden?: true;
}

/** The options a chart carries per physical axis. */
export interface OfficeChartAxesOptions {
  readonly category?: OfficeChartAxisOptions;
  readonly value?: OfficeChartAxisOptions;
  readonly secondaryValue?: OfficeChartAxisOptions;
  /** A 3-D chart's Depth (Series) Axis (`c:serAx`). */
  readonly series?: OfficeChartAxisOptions;
}

export const OFFICE_CHART_AXIS_KEYS = ["category", "value", "secondaryValue", "series"] as const;
export type OfficeChartAxisKey = (typeof OFFICE_CHART_AXIS_KEYS)[number];

const MAX_ABS = 1e15;
const MAX_FORMAT_CODE = 255;

function finite(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && Math.abs(value) <= MAX_ABS
    ? value
    : undefined;
}

function oneOf<T extends string>(value: unknown, allowed: readonly T[]): T | undefined {
  return typeof value === "string" && (allowed as readonly string[]).includes(value)
    ? (value as T)
    : undefined;
}

/**
 * The record with every field that holds; null when nothing does. A minimum
 * above the maximum, a unit that is not positive, a log base outside 2..1000
 * and a format code that is empty or too long are dropped, field by field.
 */
export function normalizeOfficeChartAxisOptions(value: unknown): OfficeChartAxisOptions | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const next: {
    -readonly [K in keyof OfficeChartAxisOptions]?: OfficeChartAxisOptions[K];
  } = {};
  let min = finite(record.min);
  let max = finite(record.max);
  if (min !== undefined && max !== undefined && min >= max) {
    min = undefined;
    max = undefined;
  }
  if (min !== undefined) next.min = min;
  if (max !== undefined) next.max = max;
  const majorUnit = finite(record.majorUnit);
  if (majorUnit !== undefined && majorUnit > 0) next.majorUnit = majorUnit;
  const minorUnit = finite(record.minorUnit);
  if (minorUnit !== undefined && minorUnit > 0) next.minorUnit = minorUnit;
  const logBase = finite(record.logBase);
  if (logBase !== undefined && logBase >= 2 && logBase <= 1000) {
    next.logBase = logBase;
    // A logarithmic axis has no zero or negative bound: such a bound is
    // automatic again, as the chart drew it anyway (PPT-B252).
    if (next.min !== undefined && next.min <= 0) delete next.min;
    if (next.max !== undefined && next.max <= 0) delete next.max;
  }
  if (record.reverse === true) next.reverse = true;
  const majorTickMark = oneOf(record.majorTickMark, OFFICE_CHART_TICK_MARKS);
  if (majorTickMark) next.majorTickMark = majorTickMark;
  const minorTickMark = oneOf(record.minorTickMark, OFFICE_CHART_TICK_MARKS);
  if (minorTickMark) next.minorTickMark = minorTickMark;
  const labelPosition = oneOf(record.labelPosition, OFFICE_CHART_TICK_LABEL_POSITIONS);
  if (labelPosition) next.labelPosition = labelPosition;
  if (
    typeof record.numberFormat === "string" &&
    record.numberFormat.trim() &&
    record.numberFormat.length <= MAX_FORMAT_CODE &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(record.numberFormat)
  ) {
    next.numberFormat = record.numberFormat;
  }
  const displayUnits = oneOf(record.displayUnits, OFFICE_CHART_DISPLAY_UNITS);
  if (displayUnits) {
    next.displayUnits = displayUnits;
    if (record.showDisplayUnitsLabel === true) next.showDisplayUnitsLabel = true;
  }
  const crossesRule = oneOf(record.crosses, OFFICE_CHART_AXIS_CROSSES);
  const crossesAt = finite(record.crosses);
  if (crossesRule) next.crosses = crossesRule;
  else if (crossesAt !== undefined) next.crosses = crossesAt;
  const axisType = oneOf(record.axisType, OFFICE_CHART_AXIS_TYPES);
  if (axisType) next.axisType = axisType;
  // A base unit belongs to a date axis; the skips to a text axis (a date
  // axis has no `c:tickLblSkip` / `c:tickMarkSkip`).
  const baseTimeUnit = oneOf(record.baseTimeUnit, OFFICE_CHART_BASE_TIME_UNITS);
  if (baseTimeUnit && axisType === "date") next.baseTimeUnit = baseTimeUnit;
  if (axisType === "date") {
    for (const key of ["majorTimeUnit", "minorTimeUnit"] as const) {
      const unit = oneOf(record[key], OFFICE_CHART_BASE_TIME_UNITS);
      if (unit) next[key] = unit;
    }
  }
  if (axisType !== "date") {
    for (const key of ["tickLabelSkip", "tickMarkSkip"] as const) {
      const skip = record[key];
      if (
        typeof skip === "number" &&
        Number.isInteger(skip) &&
        skip >= 1 &&
        skip <= OFFICE_CHART_MAX_SKIP
      ) {
        next[key] = skip;
      }
    }
  }
  if (record.axisPosition === "onTickMarks" || record.axisPosition === "between") next.axisPosition = record.axisPosition;
  for (const key of ["labelAngle", "titleAngle"] as const) {
    const angle = record[key];
    if (typeof angle === "number" && Number.isInteger(angle) && angle >= -90 && angle <= 90) {
      next[key] = angle;
    }
  }
  // Stacked text has no angle (Office greys Custom angle out for it).
  for (const [key, angleKey] of [["labelStacked", "labelAngle"], ["titleStacked", "titleAngle"]] as const) {
    if (record[key] === true) {
      next[key] = true;
      delete next[angleKey];
    }
  }
  const labelOffset = record.labelOffset;
  if (
    typeof labelOffset === "number" &&
    Number.isInteger(labelOffset) &&
    labelOffset >= 0 &&
    labelOffset <= 1000 &&
    labelOffset !== 100
  ) {
    next.labelOffset = labelOffset;
  }
  if ((record.labelAlign === "l" || record.labelAlign === "r") && axisType !== "date") next.labelAlign = record.labelAlign;
  // A date axis has no `c:noMultiLvlLbl` (CT_DateAx): its labels are dates.
  if (record.multiLevelLabels === false && axisType !== "date") next.multiLevelLabels = false;
  if (record.hidden === true) next.hidden = true;
  return Object.keys(next).length ? next : null;
}

/** One axis's record as that axis keeps it: `hidden` is the depth axis's alone. */
function officeChartAxisRecord(
  axis: OfficeChartAxisKey,
  value: unknown,
): OfficeChartAxisOptions | null {
  const normalized = normalizeOfficeChartAxisOptions(value);
  if (!normalized || axis === "series" || normalized.hidden === undefined) return normalized;
  const { hidden: _hidden, ...rest } = normalized;
  return Object.keys(rest).length ? rest : null;
}

/**
 * Text direction Stacked: each character of `text` on a line of its own
 * (a line break of the text stays one line break; spaces stay as lines).
 */
export function officeStackedText(text: string): string {
  return text
    .split("\n")
    .map((line) => Array.from(line).join("\n"))
    .join("\n");
}

export function normalizeOfficeChartAxesOptions(value: unknown): OfficeChartAxesOptions | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const next: { -readonly [K in OfficeChartAxisKey]?: OfficeChartAxisOptions } = {};
  for (const key of OFFICE_CHART_AXIS_KEYS) {
    const options = officeChartAxisRecord(key, record[key]);
    if (options) next[key] = options;
  }
  return Object.keys(next).length ? next : null;
}

/**
 * A chart (any host's attrs carrying `format.axisOptions`) with one axis's
 * record replaced; a record that states nothing removes the axis's entry,
 * and an empty map removes `axisOptions`. The one write every host's Format
 * pane makes.
 */
export function withOfficeChartAxisOptions<
  T extends { readonly format?: { readonly axisOptions?: OfficeChartAxesOptions } },
>(attrs: T, axis: OfficeChartAxisKey, options: OfficeChartAxisOptions | null): T {
  const axes: { -readonly [K in OfficeChartAxisKey]?: OfficeChartAxisOptions } = {
    ...(attrs.format?.axisOptions ?? {}),
  };
  const normalized = officeChartAxisRecord(axis, options);
  if (normalized) axes[axis] = normalized;
  else delete axes[axis];
  const format: Record<string, unknown> = { ...(attrs.format ?? {}) };
  if (Object.keys(axes).length) format.axisOptions = axes;
  else delete format.axisOptions;
  // A tick mark set here is the one Quick Layout's layout states too: one
  // record of one `c:majorTickMark` (`officeChartLayoutWithAxisTicks`).
  return { ...attrs, format: officeChartLayoutWithAxisTicks((attrs as { readonly chartType?: unknown }).chartType, format) };
}

export function sameOfficeChartAxisOptions(
  a: OfficeChartAxisOptions | null | undefined,
  b: OfficeChartAxisOptions | null | undefined,
): boolean {
  return JSON.stringify(normalizeOfficeChartAxisOptions(a) ?? {}) ===
    JSON.stringify(normalizeOfficeChartAxisOptions(b) ?? {});
}

const UNIT_LABELS = DISPLAY_UNIT_LABELS as Readonly<
  Record<string, Readonly<Record<OfficeChartDisplayUnit, string>>>
>;

/** The display-units table for a UI locale: exact tag, then its language, else English. */
function displayUnitLabelsFor(locale: string): Readonly<Record<OfficeChartDisplayUnit, string>> {
  const tag = locale.toLowerCase();
  const exact = Object.keys(UNIT_LABELS).find((key) => key.toLowerCase() === tag);
  if (exact) return UNIT_LABELS[exact];
  if (/^zh-(hant|hk|mo|tw)/.test(tag) && UNIT_LABELS["zh-TW"]) return UNIT_LABELS["zh-TW"];
  return UNIT_LABELS[tag.split(/[-_]/)[0]] ?? UNIT_LABELS.en;
}

/**
 * Office's own label for a display unit (what "Show display units label"
 * writes), in the reader's language: `language` is a UI locale, or the older
 * boolean for Chinese.
 */
export function officeChartDisplayUnitLabel(
  unit: OfficeChartDisplayUnit,
  language: boolean | string,
): string {
  const locale = typeof language === "string" ? language : language ? "zh" : "en";
  return displayUnitLabelsFor(locale)[unit] ?? UNIT_LABELS.en[unit];
}

/**
 * The ECharts axis fields one options record decides, merged over the axis
 * the plan already built. `formatNumber` is the host's number-format engine
 * (the worksheet's); a category axis takes no bounds, units or scale.
 */
export function officeChartAxisOptionsEChartsFields(
  options: OfficeChartAxisOptions | null | undefined,
  kind: "category" | "value",
  formatNumber: (value: number, format: string) => string,
  /** The reader's UI locale for the display-units label (or true for Chinese). */
  chinese: boolean | string = false,
  // The axis's PHYSICAL direction: a bar chart's category axis is the
  // vertical one. Defaults to a column chart's layout.
  vertical = kind === "value",
  // The colour a number format's section states ("[Red]0.0"), or null.
  formatColor?: (value: number, format: string) => string | null,
  // Plotted units per value: a 100% stacked chart plots 0..100 while its
  // axis values are shares (1 = 100%), so "0%" reads 100 as 100%, not
  
  formatUnitsPerValue = 1,
): Record<string, unknown> {
  const normalized = normalizeOfficeChartAxisOptions(options);
  if (!normalized) return {};
  const fields: Record<string, unknown> = {};
  if (normalized.reverse) fields.inverse = true;
  // Axis position ▸ On tick marks: no half-category gap at the plot's ends.
  if (kind === "category" && normalized.axisPosition === "onTickMarks") fields.boundaryGap = false;
  // Between tick marks: half a band at each end, the ticks between the bands.
  if (kind === "category" && normalized.axisPosition === "between") fields.boundaryGap = true;
  const axisTick: Record<string, unknown> =
    kind === "category" && normalized.axisPosition === "between" ? { alignWithLabel: false } : {};
  if (normalized.majorTickMark) {
    axisTick.show = normalized.majorTickMark !== "none";
    if (normalized.majorTickMark === "in") axisTick.inside = true;
    // Cross: ECharts draws a tick on one side only. The host draws the inner
    // half on a twin of the axis (PPT-B114); this axis keeps the outer half.
    if (normalized.majorTickMark === "cross") fields.crossTicks = true;
  }
  if (Object.keys(axisTick).length) fields.axisTick = axisTick;
  if (normalized.minorTickMark) {
    fields.minorTick = {
      show: normalized.minorTickMark !== "none",
      ...(normalized.minorTickMark === "in" ? { inside: true } : {}),
    };
  }
  const axisLabel: Record<string, unknown> = {};
  if (normalized.labelPosition === "none") axisLabel.show = false;
  // Low = the labels at the plot's low edge whatever the axis crosses; the
  // default Next to axis follows the crossing (the renderer's, by zero).
  else if (normalized.labelPosition === "low") fields.lowLabels = true;
  else if (normalized.labelPosition === "high") {
    // High = the labels on the far side of the plot, on the side the axis's
    // physical direction has (PPT-B087). Only the labels: the axis line and
    // its tick marks stay where the axis is (c:tickLblPos positions "the tick
    // labels on the axis"), so the host draws them on a twin of the axis
    
    fields.highLabels = vertical ? "right" : "top";
  }
  if (kind === "value") {
    // The bounds and units are stated in the axis's values, as Format Axis
    // shows them: a 100% stacked chart's 0..1 by 0.2 is 0..100 by 20 plotted.
    // Taken as plotted points, a Major unit of 0.2 drew 500 labels up the
    
    if (normalized.min !== undefined) fields.min = normalized.min * formatUnitsPerValue;
    if (normalized.max !== undefined) fields.max = normalized.max * formatUnitsPerValue;
    if (normalized.majorUnit !== undefined && normalized.logBase === undefined) {
      fields.interval = normalized.majorUnit * formatUnitsPerValue;
    }
    if (normalized.minorUnit !== undefined && normalized.majorUnit !== undefined) {
      const count = Math.round(normalized.majorUnit / normalized.minorUnit);
      if (count >= 1 && count <= 100) fields.minorTick = { ...(fields.minorTick as object), splitNumber: count };
    }
    if (normalized.logBase !== undefined) {
      fields.type = "log";
      fields.logBase = normalized.logBase;
    }
    const format = normalized.numberFormat;
    const factor =
      (normalized.displayUnits
        ? OFFICE_CHART_DISPLAY_UNIT_FACTORS[normalized.displayUnits]
        : 1) * (format ? formatUnitsPerValue : 1);
    if (normalized.displayUnits) {
      // PowerPoint keeps room for the labels as they would read before the units.
      fields[OFFICE_LABEL_BEFORE_UNITS] = (value: number) =>
        formatNumber(value * (format ? formatUnitsPerValue : 1), format || "General");
    }
    if (factor !== 1 || format) {
      axisLabel.formatter = (value: number) => {
        const scaled = value / factor;
        // Without a code the units read in General, through the host's
        // engine too (its separators are the reader's).
        return formatNumber(format ? scaled : Number(scaled.toPrecision(12)), format || "General");
      };
    }
    // A format's colour code colours the labels of the section it is in; the
    
    // axis's own label colour where the section states none.
    if (format && formatColor && format.includes("[")) {
      fields.labelColor = (value: number) => formatColor(value / factor, format);
    }
    // Not the axis `name`: that is the axis title, and Office draws the units
    // label as a text of its own beside it (PPT-B086). The host places it.
    if (normalized.displayUnits && normalized.showDisplayUnitsLabel) {
      fields.displayUnitsLabel = officeChartDisplayUnitLabel(normalized.displayUnits, chinese);
    }
  } else {
    if (normalized.numberFormat) {
      const format = normalized.numberFormat;
      // ECharts hands a category label as text, but a time axis hands the
      // same callback a number; neither may throw (PPT-B088).
      axisLabel.formatter = (value: unknown) => {
        const text = String(value ?? "");
        const numeric = Number(text);
        return Number.isFinite(numeric) && text.trim() !== "" ? formatNumber(numeric, format) : text;
      };
    }
    // Interval between labels / tick marks: ECharts skips `interval`
    // categories between two it draws, Office counts the step itself.
    if (normalized.tickLabelSkip !== undefined) axisLabel.interval = normalized.tickLabelSkip - 1;
    if (normalized.tickMarkSkip !== undefined) {
      fields.axisTick = { ...(fields.axisTick as object), interval: normalized.tickMarkSkip - 1 };
    }
  }
  // Label position: Excel sets a text axis's label at its band's left or
  // right edge (X11_chart_axis43's lblAlgn="l"); the picture renderer lays it
  // there once the band's width is known (`alignCategoryLabelsInBands`).
  if (normalized.labelAlign !== undefined) axisLabel.officeLabelAlign = normalized.labelAlign;
  // Distance from axis: Office's 100 is the usual gap (ECharts' 8 px).
  if (normalized.labelOffset !== undefined) {
    axisLabel.margin = Math.round((8 * normalized.labelOffset) / 100);
  }
  // Custom angles: DrawingML turns text clockwise, ECharts counter-clockwise.
  if (normalized.labelAngle !== undefined) axisLabel.rotate = 0 - normalized.labelAngle;
  // A Custom angle of 0 is level by the author's word: the names wrap into
  
  // names under their columns in PowerPoint, turned 45 degrees here).
  if (normalized.labelAngle === 0) axisLabel[OFFICE_STATED_LEVEL_LABELS] = true;
  if (normalized.titleAngle !== undefined) fields.nameRotate = 0 - normalized.titleAngle;
  // Stacked: level, a character per line. The host stacks the text over
  // whatever formats the labels (`officeStackedText`), and the title's,
  // whose words it owns.
  if (normalized.labelStacked) {
    axisLabel.rotate = 0;
    fields.labelStacked = true;
  }
  if (normalized.titleStacked) {
    fields.nameRotate = 0;
    fields.titleStacked = true;
  }
  if (Object.keys(axisLabel).length) fields.axisLabel = axisLabel;
  return fields;
}

/**
 * The GL `axis3D` fields a 3-D chart's Depth (Series) Axis takes from its
 * record (`format.axisOptions.series`): Series in reverse order, the tick
 * mark type, the label position None, the intervals between tick marks and
 * labels (Office counts the step, the engine the categories it skips), and
 * Axes > Depth Axis cleared, which keeps the axis's coordinates for the
 * camera but paints no line, tick or label.
 */
export function officeChartDepthAxisFields(
  options: OfficeChartAxisOptions | null | undefined,
): Record<string, unknown> {
  const normalized = normalizeOfficeChartAxisOptions(options);
  if (!normalized) return {};
  if (normalized.hidden) {
    return {
      ...(normalized.reverse ? { inverse: true } : {}),
      name: "",
      axisLine: { show: true, lineStyle: { color: "rgba(0,0,0,0)" } },
      axisTick: { show: false },
      axisLabel: { show: false },
    };
  }
  const fields: Record<string, unknown> = {};
  if (normalized.reverse) fields.inverse = true;
  const axisTick: Record<string, unknown> = {};
  if (normalized.majorTickMark) {
    axisTick.show = normalized.majorTickMark !== "none";
    if (normalized.majorTickMark === "in") axisTick.inside = true;
  }
  if (normalized.tickMarkSkip !== undefined) axisTick.interval = normalized.tickMarkSkip - 1;
  if (Object.keys(axisTick).length) fields.axisTick = axisTick;
  const axisLabel: Record<string, unknown> = {};
  if (normalized.labelPosition === "none") axisLabel.show = false;
  if (normalized.tickLabelSkip !== undefined) axisLabel.interval = normalized.tickLabelSkip - 1;
  if (Object.keys(axisLabel).length) fields.axisLabel = axisLabel;
  return fields;
}

// ---------------------------------------------------------------------------
// A date axis over the chart's categories.

const DAY_MS = 86_400_000;

function parseCategoryDate(raw: unknown): number | null {
  const text = typeof raw === "number" ? String(raw) : typeof raw === "string" ? raw.trim() : "";
  // A workbook date is its Excel serial (DATE(2026,1,1) is 46023): on a
  // Date axis Excel reads a number as that date, so the chart does too
  // (PPT-B281). Excel's fictitious 1900-02-29 is no date.
  if (/^\d{1,7}$/.test(text)) {
    const serial = Number(text);
    if (serial >= 1 && serial <= 2958465 && serial !== EXCEL_1900_FICTITIOUS_LEAP_SERIAL) {
      return excelSerialToUtcMs(serial);
    }
    return null;
  }
  let match = /^(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?$/.exec(text);
  let year: number, month: number, day: number;
  if (match) {
    [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3] ?? 1)];
  } else if ((match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text))) {
    [month, day, year] = [Number(match[1]), Number(match[2]), Number(match[3])];
  } else if ((match = /^(\d{4})年(\d{1,2})月(?:(\d{1,2})日)?$/.exec(text))) {
    [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3] ?? 1)];
  } else {
    return null;
  }
  const time = Date.UTC(year, month - 1, day);
  const date = new Date(time);
  return date.getUTCMonth() === month - 1 && date.getUTCDate() === day ? time : null;
}

/**
 * The categories of a Date axis as times (UTC ms), each put at the start of
 * its base unit (Office plots every date of one month at that month on a
 * Months axis), or null when a category is not a date: a date axis over text
 * categories is drawn as the text axis it reads as.
 */
export function officeChartDateAxisTimes(
  categories: readonly unknown[],
  baseTimeUnit?: OfficeChartBaseTimeUnit,
): number[] | null {
  if (!categories.length) return null;
  const times: number[] = [];
  for (const category of categories) {
    const time = parseCategoryDate(category);
    if (time === null) return null;
    const date = new Date(time);
    times.push(
      baseTimeUnit === "months"
        ? Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1)
        : baseTimeUnit === "years"
          ? Date.UTC(date.getUTCFullYear(), 0, 1)
          : time,
    );
  }
  return times;
}

/** The smallest step a date axis draws for its base unit, in ms. */
export function officeChartBaseTimeUnitMs(unit: OfficeChartBaseTimeUnit | undefined): number {
  return unit === "years" ? 365 * DAY_MS : unit === "months" ? 28 * DAY_MS : DAY_MS;
}
