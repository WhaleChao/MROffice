/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A plan drawn with the chart's hand-placed elements where they were put
 * (`format.elementLayout`, Office's `c:manualLayout`): the title, the axis
 * titles, the legend and the plot area at their fractions of the chart, and
 * each dragged data label at its offset (`dataLabelOverrides[].offset`).
 * Elements without a place keep the plan's automatic layout -- and a moved
 * legend or title leaves the plot area where it is, as Office's mouse moves do.
 *
 * Applied by the chart engine to every host's plan, at the picture's size.
 */
import { ptToPx } from "../office-units";
import { chartSupportsDraggedLabelLeaderLines, type ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";

import {
  normalizeChartElementBox,
  normalizeChartElementLayout,
  chartDataLabelOffsetSigns,
  chartElementPlaceIsBox,
  chartPercent,
  type ChartElementBox,
} from "./chart-element-layout";
import { officeChartAxisRole, officeChartHitSeries, officeChartPieItemPoint, officeChartPieSliceNames } from "./chart-hit";
import { officeChartTextEChartsStyle } from "./chart-text";
import {
  OFFICE_LEGEND_KEY_SIZED,
  OFFICE_LEGEND_KEY_TO_TEXT_PX,
  OFFICE_LEGEND_NAME_SHIFTED,
  officeLegendKeyPx,
  officeLegendKeyToNamePx,
} from "./legend-spacing";

type Option = Record<string, unknown>;

const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);

/** The graphic ids of axis titles drawn where they were placed. */
export const CHART_PLACED_AXIS_TITLE_IDS = {
  axisX: "mrai-axis-title-x",
  axisY: "mrai-axis-title-y",
  axisSeries: "mrai-axis-title-series",
} as const;
/** On an axis whose title was placed by hand over an automatic plot area:
 * the title's words, whose room at the axis the plot keeps. */
export const OFFICE_PLACED_AXIS_TITLE_ROOM = "officePlacedAxisTitleRoom";

export interface ChartPlacement {
  option: Option;
  /** The words of the axis titles drawn where they were placed (by physical axis). */
  placedAxisTitles: Partial<Record<"axisX" | "axisY", string>>;
  /** The legend's entries shown, by their index in its whole list (a deleted
   * entry is not shown); absent: every entry, in order. */
  legendEntries?: number[];
}

/** The series whose legend lists their points (a pie's slices), not themselves. */
const LEGEND_BY_POINT = new Set(["pie", "funnel", "radar"]);

/** The legend's whole list as ECharts makes it: each series' name, or its points' names. */
function legendNames(option: Option): string[] {
  const legend = first(option.legend);
  if (legend && Array.isArray(legend.data))
    return legend.data.map((entry) => (record(entry) ? String(entry.name ?? "") : String(entry ?? "")));
  const names: string[] = [];
  for (const series of Array.isArray(option.series) ? option.series : []) {
    if (!record(series)) continue;
    if (LEGEND_BY_POINT.has(String(series.type))) {
      for (const item of Array.isArray(series.data) ? series.data : [])
        if (record(item) && typeof item.name === "string" && !names.includes(item.name)) names.push(item.name);
    } else if (typeof series.name === "string" && !names.includes(series.name)) {
      names.push(series.name);
    }
  }
  return names;
}

/** Office's chart text where the chart states none (10 pt; the ink the
 * plan's elements take, `withOfficeInk`). */
const CHART_TEXT_DEFAULTS: Readonly<Record<string, unknown>> = {
  fontSize: (10 * 4) / 3,
  fontWeight: "normal",
  fontStyle: "normal",
  color: "#595959",
};

/**
 * The legend with each entry formatted apart drawn in its own text
 * (`format.legendEntryFonts`): the chart's text (`format.text`) with the
 * entry's over it -- what the legend states for itself (`legendStyle.text`)
 * does not reach such an entry, as Windows PowerPoint draws it (truth
 * office-truth/probes/c30legentry: in a 14 pt 404040 legend of a 12 pt
 * chart an entry stating bold alone drew 12 pt black bold, one stating red
 * alone 12 pt red; one stating italic off in an italic legend stood upright
 * and the others leaned). An entry is named by its place in the legend's
 * whole list (`chartLegendEntryOfPosition`, as the deleted entries are).
 */
function legendWithEntryFonts(
  option: Option,
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "format">,
  entryOf: (position: number) => number,
): unknown {
  const rows = new Map((attrs.format?.legendEntryFonts ?? []).map((row) => [row.entry, row.font]));
  const names = legendNames(option);
  const listed = first(option.legend)?.data;
  const chart = officeChartTextEChartsStyle(attrs.format?.text);
  const chartFamily = record(option.textStyle) ? option.textStyle.fontFamily : undefined;
  // What the legend states of its own, put back to the chart's.
  const reset: Record<string, unknown> = {};
  for (const key of Object.keys(officeChartTextEChartsStyle(attrs.format?.legendStyle?.text))) {
    const value = key in chart ? chart[key] : key === "fontFamily" ? chartFamily : CHART_TEXT_DEFAULTS[key];
    if (value !== undefined) reset[key] = value;
  }
  // A horizontal clustered bar chart's legend runs from its last series
  // (`officeBarSeriesOrder`, drawn after this): Office counts its entries as
  // they stand, from the top (c30legentry p8: idx 0 is the third series).
  const fromLast = legendListedFromLastSeries(attrs, option);
  let changed = false;
  const data = names.map((name, position) => {
    const own = Array.isArray(listed) ? listed[position] : undefined;
    const font = rows.get(fromLast ? names.length - 1 - position : entryOf(position));
    if (!font) return own !== undefined ? own : name;
    changed = true;
    const base = record(own) ? own : { name };
    const style = record(base.textStyle) ? base.textStyle : {};
    return { ...base, textStyle: { ...style, ...reset, ...officeChartTextEChartsStyle(font) } };
  });
  if (!changed) return option.legend;
  return withFirst(option.legend, (legend) => ({ ...legend, data, ...legendKeysForEntries(legend, data, option) }));
}










function legendKeysForEntries(legend: Record<string, unknown>, data: unknown[], option: Option): Record<string, unknown> {
  // Filled keys only: the probe measured squares; a legend of lines keeps its keys.
  if (legend[OFFICE_LEGEND_KEY_SIZED] !== "filled" || typeof legend.itemHeight !== "number") return {};
  const style = record(legend.textStyle) ? legend.textStyle : {};
  const chartText = record(option.textStyle) ? option.textStyle : {};
  const own = typeof style.fontSize === "number" ? style.fontSize : typeof chartText.fontSize === "number" ? chartText.fontSize : 12;
  const sizes = data.map((entry) =>
    record(entry) && record(entry.textStyle) && typeof entry.textStyle.fontSize === "number" ? entry.textStyle.fontSize : own,
  );
  const largest = Math.max(...sizes);
  if (!(largest > 0) || largest === own) return {};
  const family = typeof style.fontFamily === "string" && style.fontFamily ? style.fontFamily : chartText.fontFamily;
  const key = officeLegendKeyPx(largest, family);
  return {
    itemHeight: key,
    itemWidth: key,
    ...(legend[OFFICE_LEGEND_NAME_SHIFTED] === true
      ? { textStyle: { ...style, padding: [0, 0, 0, Math.round((officeLegendKeyToNamePx(key, false) - OFFICE_LEGEND_KEY_TO_TEXT_PX) * 100) / 100] } }
      : {}),
  };
}

/**
 * Whether the drawing lists the legend from the last series up
 * (`chart-render-server` `officeBarSeriesOrder`, its reversed legend: a
 * horizontal clustered bar chart of two or more series, its categories in
 * order, every series a bar on the primary axis, none stacked).
 */
export function legendListedFromLastSeries(
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "format">,
  option: Option,
): boolean {
  if (attrs.chartType !== "bar") return false;
  const reversed = (attrs.format?.axisOptions as { category?: { reverse?: boolean } } | undefined)?.category?.reverse === true;
  if (reversed) return false;
  const series = Array.isArray(option.series) ? option.series : [];
  return (
    series.length >= 2 &&
    series.every((entry) => record(entry) && entry.type === "bar" && !entry.stack) &&
    !series.some((entry) => record(entry) && entry.xAxisIndex === 1)
  );
}

/** The first of a component that may be one or several. */
function withFirst(value: unknown, change: (entry: Record<string, unknown>) => Record<string, unknown>): unknown {
  if (Array.isArray(value)) return value.map((entry, index) => (index === 0 && record(entry) ? change(entry) : entry));
  return record(value) ? change(value) : value;
}

function first(value: unknown): Record<string, unknown> | null {
  if (Array.isArray(value)) return record(value[0]) ? value[0] : null;
  return record(value) ? value : null;
}

function placedBox(entry: Record<string, unknown>, box: { x: number; y: number; w?: number; h?: number }) {
  const { right: _right, bottom: _bottom, ...rest } = entry;
  return {
    ...rest,
    left: chartPercent(box.x),
    top: chartPercent(box.y),
    ...(box.w !== undefined ? { width: chartPercent(box.w) } : {}),
    ...(box.h !== undefined ? { height: chartPercent(box.h) } : {}),
  };
}

/** A radius as ECharts reads it: px, or a percentage of half the smaller side. */
function radiusPx(value: unknown, half: number): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && /^-?\d+(\.\d+)?%$/.test(value.trim())) return (Number.parseFloat(value) / 100) * half;
  return null;
}

/** A round plot (a pie, a doughnut, a sunburst, a radar) filling the plot area's box. */
function placedRound(
  entry: Record<string, unknown>,
  box: ChartElementBox,
  width: number,
  height: number,
  widestPx?: number | null,
  /** Point Explosion, as a share of the radius (the largest asked for). */
  explosion = 0,
) {
  const half = Math.min(width, height) / 2;
  const radius = entry.radius;
  const [inner, current] = Array.isArray(radius) ? [radius[0], radius[1]] : [0, radius ?? "75%"];
  const innerPx = radiusPx(inner, half) ?? 0;
  const currentPx = radiusPx(current, half);
  // A doughnut's rings share the box: the outermost ring meets its edge and
  // every other ring keeps its share of that radius (H21's three rings were
  // each grown to the box's edge, one on another; X-182).
  const boxOuter = Math.min(box.w * width, box.h * height) / 2;
  const outer =
    widestPx && widestPx > 0 && currentPx && currentPx > 0 && currentPx < widestPx ? boxOuter * (currentPx / widestPx) : boxOuter;
  // A pie of no radius only keeps a legend (a 3-D pie's slices are its own
  // drawing, `lib/office-charts/pie-3d`): grown into the placed box it drew
  // the flat pie over the 3-D one (H11_32readwritePieChartExploded3D1; X-111).
  if (currentPx === 0) return entry;
  const ratio = currentPx && currentPx > 0 ? Math.max(0, Math.min(0.95, innerPx / currentPx)) : 0;
  // The centre is a fraction of the whole chart: a box the plan gave the
  // series (a side legend's room, `right: 71`) would make "50%" the middle
  // of what is left (Word, STRESS2 page 5: a pie placed at 0.18 + 0.64 / 2
  // was drawn 36px left of its box's middle).
  const { left: _left, right: _right, top: _top, bottom: _bottom, width: _width, height: _height, ...rest } = entry;
  // Exploded slices stay inside the box: the pie is smaller by the share
  // they are pulled out, and pulled out by that share of its radius
  
  // reaches 195 pt from its middle in PowerPoint; drawn the box's full
  // 195 pt with an 8 px pull, its slices only parted by their white edges).
  const exploded = explosion > 0 && entry.selectedOffset !== undefined ? explosion : 0;
  const ring = outer / (1 + exploded);
  return {
    ...rest,
    center: [chartPercent(box.x + box.w / 2), chartPercent(box.y + box.h / 2)],
    radius: ratio > 0 ? [ring * ratio, ring] : ring,
    ...(exploded > 0 ? { selectedOffset: ring * exploded } : {}),
  };
}

const ROUND_SERIES = new Set(["pie", "sunburst"]);
const BOXED_SERIES = new Set(["funnel", "treemap"]);

/**
 * The plan's option with the chart's placed elements put where they were
 * placed. `size` is the picture's (px): without it (a WebGL plan the browser
 * sizes) only what a percentage places is moved -- the title, the legend,
 * a 2-D plot area.
 */
export function optionWithChartPlacement(
  option: Option,
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "format" | "data"> | null | undefined,
  size: { width: number; height: number } | null,
  /** The width the drawing gives a text (the engine's measure). */
  measure: (text: string, font: string) => number = (text, font) =>
    text.length * (Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1]) || 12) * 0.6,
): ChartPlacement {
  const placedAxisTitles: ChartPlacement["placedAxisTitles"] = {};
  if (!attrs) return { option, placedAxisTitles };
  const layout = normalizeChartElementLayout(attrs.format?.elementLayout);
  const offsets = (attrs.format?.dataLabelOverrides ?? []).filter(
    (entry) => entry.offset && entry.pointIndex !== undefined,
  );
  const hidden = attrs.format?.legendHiddenEntries ?? [];
  const entryFonts = attrs.format?.legendEntryFonts ?? [];
  const entryOf = chartLegendEntryOfPosition(attrs);
  // A worksheet chart's plot area stated by its outer box (`plotAreaOuter`,
  // read only off a workbook's part).
  const outerPlot = normalizeChartElementBox(attrs.format?.plotAreaOuter);
  if (!layout && !outerPlot && !offsets.length && !hidden.length && !entryFonts.length) return { option, placedAxisTitles };
  const next: Option = { ...option };
  // The legend's entries formatted one at a time (`legendEntryFonts`, Office's
  // `c:legendEntry/c:txPr`), before the deleted ones go: each by its place in
  // the legend's whole list.
  if (entryFonts.length && next.legend !== undefined) next.legend = legendWithEntryFonts(next, attrs, entryOf);
  // The legend's deleted entries (Office's `c:legendEntry/c:delete`): its
  // list without them; the series and points stay drawn.
  let legendEntries: number[] | undefined;
  if (hidden.length && next.legend !== undefined) {
    const names = legendNames(next);
    legendEntries = names.map((_, index) => index).filter((index) => !hidden.includes(entryOf(index)));
    // An entry keeps its own key: the plan's listed entries carry their icon
    // and key paint, which a bare name loses (X-295: a line of no markers
    // keyed by the engine's default line-and-dot once another entry was
    // deleted).
    const listed = first(next.legend)?.data;
    const shown = legendEntries.map((index) => (Array.isArray(listed) && listed[index] !== undefined ? listed[index] : names[index]));
    next.legend = withFirst(next.legend, (legend) => ({ ...legend, data: shown }));
  }

  if (layout?.title && next.title !== undefined) next.title = withFirst(next.title, (title) => placedBox(title, layout.title!));
  // The legend's drawn box is its place (no padding of its own around it).
  // A side legend's entries run down the box: ECharts wraps the ones a
  // shorter box cannot hold into another column, which Office never draws
  // (and a box read from Excel's metrics, or a drag that lands a pixel short,
  // turned a two-entry column into a row), so its height wraps nothing; a
  // horizontal legend's entries still wrap into rows at its width, as
  // Office's do.
  if (layout?.legend && next.legend !== undefined)
    next.legend = withFirst(next.legend, (legend) => {
      const placed: Record<string, unknown> = { ...placedBox(legend, layout.legend!), padding: 0 };
      // Placed by hand, a Top Right legend stays where it was put, not at the
      // plot's top (`officeLegendAtPlotTop`; CDC p31's legend above its plot).
      delete placed.officeLegendAtPlotTop;
      // Placed by its corner alone, the legend keeps the size Office gives
      // it: its entries laid out as they are where the chart places it
      
      if (!chartElementPlaceIsBox(layout.legend)) return placed;
      // Office lays a hand-sized legend's entries out to fit its box: in a
      // box at least twice as wide as it is tall they stand in a row (corpus
      // R29-D05: a right-hand legend dragged across the chart's top reads
      // as one row in Word), wrapping at the box's width.
      const box = layout.legend;
      const wide = size !== null && box.w * size.width >= 2 * box.h * size.height;
      // A column of entries in a box sized by hand is spread down it, each
      // entry in the middle of an equal share of the box's height
      // (PowerPoint, CDC TB deck p36: four 18pt entries 27.8pt apart in a
      // 111pt box; p45: three 31.6pt apart in 95pt) -- entries of one line
      // each. PowerPoint sets wrapped entries a share apart from top to top
      
      // has one gap between entries, so entries of unlike lines take the gap
      // that brings the last to its place (CDC p24's two-line names).
      const spreadColumn = (placed: Record<string, unknown>) => {
        const names = legendNames(next);
        if (size === null || names.length < 1) return;
        const textStyle = record(placed.textStyle) ? placed.textStyle : record(legend.textStyle) ? legend.textStyle : {};
        const fontSize = typeof textStyle.fontSize === "number" ? textStyle.fontSize : 12;
        const lineHeight = fontSize * 1.2;
        const keyHeight = typeof legend.itemHeight === "number" ? legend.itemHeight : 14;
        const pitch = (box.h * size.height) / names.length;
        const font = `${String(textStyle.fontWeight ?? "normal")} ${fontSize}px ${String(textStyle.fontFamily ?? "sans-serif")}`;
        const wrapAt = typeof textStyle.width === "number" ? textStyle.width : null;
        // As the entries read: a name wrapped in its box is on its lines.
        const shown = (name: string) =>
          typeof placed.formatter === "function" ? String((placed.formatter as (name: string) => unknown)(name)) : name;
        // Keys with no words (`legendBlankNames`) are as tall as the key alone: taken
        
        // off the slide's own captions beside them.
        const entry = names.every((name) => shown(name) === "") ? keyHeight : Math.max(lineHeight, keyHeight);
        const oneLine = names.every((name) => !shown(name).includes("\n") && (wrapAt === null || measure(name, font) <= wrapAt));
        if (pitch > entry && oneLine) {
          placed.textStyle = { ...textStyle, lineHeight };
          placed.itemGap = pitch - entry;
          placed.padding = [(pitch - entry) / 2, 0, 0, 0];
        } else if (!oneLine && names.length > 1) {
          // Names broken into their lines before the engine sees them (the
          // side legend's own breaking, `withSideLegendRoom`): the engine
          // stands each entry as tall as its lines at the line height it is
          // given, so it is given the one its lines are drawn at. Names the
          // engine wraps itself (`textStyle.width`): it stands a wrapped entry
          // about 0.7 of its size taller than its lines, and already that much
          // down its box.
          const preBroken = names.some((name) => shown(name).includes("\n"));
          if (preBroken) placed.textStyle = { ...textStyle, lineHeight };
          const heights = names.map((name) => {
            const lines = shown(name).split("\n").length;
            return Math.max(entry, lines * lineHeight + (!preBroken && lines > 1 ? fontSize * 0.7 : 0));
          });
          const gap = pitch - heights.slice(0, -1).reduce((sum, value) => sum + value, 0) / (names.length - 1);
          if (gap > 0) {
            placed.itemGap = gap;
            placed.padding = 0;
          }
        }
      };
      // A column's name wider than the box wraps in it, between its words,
      
      // "Brennholz und feste biogene Treibstoffe" ran out of a 0.38-wide box
      
      // here, in the face Office sets them in: the engine would break them as
      // it measures its own face, sans-serif where none is stated -- wider
      
      const wrapInBox = (placed: Record<string, unknown>) => {
        if (size === null) return;
        const textStyle = record(placed.textStyle) ? placed.textStyle : record(legend.textStyle) ? legend.textStyle : {};
        if (typeof textStyle.width === "number") return;
        const stated = placed.formatter ?? legend.formatter;
        if (stated !== undefined && typeof stated !== "function") return;
        const keyWidth = typeof legend.itemWidth === "number" ? legend.itemWidth : 25;
        const room = box.w * size.width - keyWidth - 5;
        if (!(room > 0)) return;
        const fontSize = typeof textStyle.fontSize === "number" ? textStyle.fontSize : 12;
        // In the face the words are drawn in: their own, else the chart's;
        
        // Air Calibration" stands on one line in its 0.24-wide box).
        const chartText = record(next.textStyle) ? next.textStyle : {};
        const family = [textStyle.fontFamily, chartText.fontFamily].find((face) => typeof face === "string" && face) ?? "sans-serif";
        const judged = family === "sans-serif" ? "Calibri" : String(family);
        const font = `${String(textStyle.fontWeight ?? "normal")} ${fontSize}px ${judged}`;
        const wrapLine = (line: string) => {
          // Right-to-left words the measure has no widths of stay as they
          
          // each in PowerPoint; measured as the generic face they broke).
          if (/[֐-ࣿיִ-﷿ﹰ-﻿]/u.test(line)) return line;
          const lines: string[] = [];
          let current = "";
          for (const word of line.split(" ")) {
            const candidate = current ? `${current} ${word}` : word;
            if (current && measure(candidate, font) > room) {
              lines.push(current);
              current = word;
            } else current = candidate;
          }
          lines.push(current);
          return lines.join("\n");
        };
        const shown = (name: string) =>
          typeof stated === "function" ? String((stated as (name: string) => unknown)(name)) : name;
        const wrap = (name: string) => shown(name).split("\n").map(wrapLine).join("\n");
        if (legendNames(next).some((name) => wrap(name) !== shown(name))) placed.formatter = wrap;
      };
      if (legend.orient === "vertical" && wide) {
        placed.orient = "horizontal";
        delete placed.height;
      } else if (legend.orient === "vertical") {
        delete placed.height;
        // The entries' words wrap at the box's width, not the automatic
        // legend's (PH082_issue.2506 chart 4: a 0.52-wide box held
        // "surface secondaire couvert" on one line in Excel, and the
        // automatic fifth of the chart wrapped it and pushed the fifth entry
        // into a second column, X-252).
        const textStyle = record(legend.textStyle) ? legend.textStyle : {};
        if (size !== null && typeof textStyle.width === "number") {
          const keyWidth = typeof legend.itemWidth === "number" ? legend.itemWidth : 25;
          const room = box.w * size.width - keyWidth - 5;
          if (room > textStyle.width) placed.textStyle = { ...textStyle, width: room };
        } else {
          wrapInBox(placed);
        }
        spreadColumn(placed);
      }
      if (placed.orient !== "vertical" && size !== null) {
        const names = legendNames(next);
        const textStyle = record(legend.textStyle) ? legend.textStyle : {};
        const fontSize = typeof textStyle.fontSize === "number" ? textStyle.fontSize : 12;
        const font = `${String(textStyle.fontWeight ?? "normal")} ${fontSize}px ${String(textStyle.fontFamily ?? "sans-serif")}`;
        const keyWidth = typeof legend.itemWidth === "number" ? legend.itemWidth : 25;
        const widths = names.map((name) => keyWidth + 5 + measure(name, font));
        const entries = widths.reduce((sum, width) => sum + width, 0);
        const room = box.w * size.width;
        const stated = typeof legend.itemGap === "number" ? legend.itemGap : 10;
        // Two entries side by side: the two widest, not the widest twice (CDC
        // p57: "Recent transmission" beside "Not recent transmission" fit
        // their box in one row in PowerPoint, and were stacked).
        const [widest = 0, next2 = 0] = [...widths].sort((a, b) => b - a);
        if (names.length >= 2 && widest + next2 + stated > room) {
          // A box too narrow for two entries side by side holds them one
          // under another, a column, as a side legend does -- and is listed
          
          // above a stacked column chart, its seven long names one per row,
          // last series first in PowerPoint).
          placed.orient = "vertical";
          delete placed.height;
          wrapInBox(placed);
          spreadColumn(placed);
          return placed;
        }
        // A row that fits its box is spread across it, the same space before,
        // between and after the entries (Word, corpus R29-D05: three entries
        // in a box the chart's width stand 132px apart, the first 82px in).
        // (Two px an entry of slack: the engine measures its row with its own
        // rounding, and a row spread to the pixel wrapped its last entry.)
        const gap = names.length ? (room - entries - 2 * names.length) / (names.length + 1) : 0;
        // (A row the plan leaves unoriented is ECharts' horizontal one: CDC
        // p38's four entries spread across their foot box in PowerPoint.)
        if ((placed.orient ?? "horizontal") === "horizontal" && gap > stated) {
          placed.itemGap = gap;
          placed.padding = [0, 0, 0, gap];
        }
        // A row that fits its box stands in the middle of the box's height,
        
        // entries in a 0.097-high foot box, centred in it in PowerPoint and
        // drawn at its top, over the category names above it).
        if (
          (placed.orient ?? "horizontal") === "horizontal" &&
          names.length > 0 &&
          entries + stated * (names.length - 1) <= room
        ) {
          const keyHeight = typeof legend.itemHeight === "number" ? legend.itemHeight : 14;
          const free = box.h * size.height - Math.max(fontSize * 1.2, keyHeight);
          if (free > 2) {
            const padding = Array.isArray(placed.padding) ? [...(placed.padding as number[])] : [0, 0, 0, 0];
            padding[0] = free / 2;
            placed.padding = padding;
          }
        }
      }
      return placed;
    });

  // A WebGL plan (no size) keeps its 3-D plot where its view puts it.
  // A Pie of Pie / Bar of Pie lays its two plots out itself; its grid is the
  // second plot's bar, which the placed box stretched over the whole plot
  
  const ofPie = attrs.chartType === "pie" && ["pieOfPie", "barOfPie"].includes(String((attrs as { chartSubtype?: unknown }).chartSubtype));
  const plot = size && !ofPie ? layout?.plotArea : undefined;
  if (plot) {
    // The placed box is the INNER plot area (`c:layoutTarget val="inner"`,
    // the only target the model states): the rectangle the axes stand on,
    // the tick labels outside it. The plans' auto grids hold their labels
    // inside their box (ECharts 6 `outerBoundsMode: "same"`, and its default
    // `auto` holds them inside the picture); kept on a placed box, that
    // shrank the plot by the labels' room (Word, STRESS2 page 5: an inner
    // box 0.12/0.04/0.84/0.81 is drawn exactly there, our axes stood 31px
    // right of it and 19px above its foot).
    const innerBox = (grid: Record<string, unknown>) => {
      const { outerBoundsContain: _contain, outerBounds: _bounds, ...rest } = grid;
      return { ...placedBox(rest, plot), containLabel: false, outerBoundsMode: "none" };
    };
    if (next.grid !== undefined) {
      next.grid = withFirst(next.grid, innerBox);
    } else if (next.xAxis !== undefined || next.yAxis !== undefined) {
      // A plan that leaves its axes on ECharts' default grid (a scatter, a
      // bubble chart) states none: the placed plot area is that grid. Without
      // it a plot area dragged by hand was saved and drawn where it had been.
      next.grid = innerBox({});
    } else if (Array.isArray(next.series)) {
      const half = size ? Math.min(size.width, size.height) / 2 : 0;
      const widest = size
        ? Math.max(
            0,
            ...next.series.map((entry) => {
              if (!record(entry) || !ROUND_SERIES.has(String(entry.type))) return 0;
              const radius = entry.radius;
              return radiusPx(Array.isArray(radius) ? radius[1] : radius ?? "75%", half) ?? 0;
            }),
          )
        : 0;
      const format = (attrs as { format?: { pieExplosion?: number; sliceStyles?: readonly { explosion?: number }[] } } | null | undefined)?.format;
      const explosion =
        Math.max(0, format?.pieExplosion ?? 0, ...(format?.sliceStyles ?? []).map((style) => style?.explosion ?? 0)) / 100;
      next.series = next.series.map((entry) => {
        if (!record(entry)) return entry;
        if (BOXED_SERIES.has(String(entry.type))) return placedBox(entry, plot);
        if (ROUND_SERIES.has(String(entry.type)) && size) return placedRound(entry, plot, size.width, size.height, widest, explosion);
        return entry;
      });
    }
    if (record(next.radar) && size) {
      const half = Math.min(size.width, size.height) / 2;
      const outer = Math.min(plot.w * size.width, plot.h * size.height) / 2;
      next.radar = {
        ...next.radar,
        center: [chartPercent(plot.x + plot.w / 2), chartPercent(plot.y + plot.h / 2)],
        radius: outer,
      };
      void half;
    }
  }

  // A worksheet chart's plot area stated by its OUTER box (Office's manual
  // layout naming no `c:layoutTarget`, which is `outer`): the box holds the
  // plot with its tick labels, as Windows Excel lays it out (reading
  // xlsxChartOuterPlotArea, R456-X21: x 0.215, w 0.344 of an 885pt chart,
  // the value axis to 1,400; laid out automatically, the plot took the whole
  // chart and ran off the page). The grid holds its axis labels inside the
  // box (ECharts 6 `outerBoundsMode: "same"`); a deck's or a document's
  // chart never carries one.
  if (outerPlot && size && !ofPie && !plot && next.grid !== undefined) {
    next.grid = withFirst(next.grid, (grid: Record<string, unknown>) => {
      const { outerBounds: _bounds, containLabel: _contain, ...rest } = grid;
      return { ...placedBox(rest, outerPlot), outerBoundsMode: "same", outerBoundsContain: "axisLabel" };
    });
  }

  // An axis title placed by hand: drawn where it was put (ECharts places an
  // axis's name only along its axis), its axis left without a name.
  const graphics: Record<string, unknown>[] = [];
  // A title drawn here is a graphic, which takes none of the chart's text
  
  // depth axis title in Calibri, not the system face).
  const chartText = record(next.textStyle) ? next.textStyle : {};
  const chartFamily = typeof chartText.fontFamily === "string" && chartText.fontFamily ? chartText.fontFamily : "sans-serif";
  if (size) {
    for (const [key, role] of [
      ["axisTitleCategory", "category"],
      ["axisTitleValue", "value"],
    ] as const) {
      const place = layout?.[key];
      if (!place) continue;
      const physical = officeChartAxisRole(attrs, "X") === role ? "X" : "Y";
      const axisKey = physical === "X" ? "xAxis" : "yAxis";
      const axis = first(next[axisKey]);
      const name = axis?.name;
      if (!axis || typeof name !== "string" || !name) continue;
      const style = record(axis.nameTextStyle) ? axis.nameTextStyle : {};
      const fontSize = typeof style.fontSize === "number" ? style.fontSize : 12;
      const family = typeof style.fontFamily === "string" ? style.fontFamily : chartFamily;
      const weight = style.fontWeight === undefined ? "normal" : String(style.fontWeight);
      const lines = name.split("\n");
      const along = Math.max(...lines.map((line) => measure(line, `${weight} ${fontSize}px ${family}`)));
      // Office's chart text runs at 1.2 of its size; the drawing's default
      // (the size itself) put a wrapped title's two lines over each other
      // (CDC p4: "Cases per" / "100,000 persons" in 20 pt bold Calibri).
      const lineHeight = Math.round(fontSize * 1.2 * 100) / 100;
      const across = lines.length * lineHeight;
      // The title's angle (Format Axis Title > Custom angle, counter-
      // clockwise as ECharts states it); a vertical axis's reads bottom to top.
      const degrees =
        typeof axis.nameRotate === "number"
          ? axis.nameRotate
          : physical === "Y" && axis.nameLocation === "middle"
            ? 90
            : 0;
      const turn = (degrees * Math.PI) / 180;
      // The box the turned text takes, its top-left corner where it was placed.
      const boxWidth = Math.abs(along * Math.cos(turn)) + Math.abs(across * Math.sin(turn));
      const boxHeight = Math.abs(along * Math.sin(turn)) + Math.abs(across * Math.cos(turn));
      const left = place.x * size.width, top = place.y * size.height;
      graphics.push({
        id: CHART_PLACED_AXIS_TITLE_IDS[physical === "X" ? "axisX" : "axisY"],
        type: "text",
        silent: true,
        z: 50,
        x: left + boxWidth / 2,
        y: top + boxHeight / 2,
        rotation: turn,
        style: {
          text: name,
          fontSize,
          fontFamily: family,
          fontWeight: weight,
          lineHeight,
          ...(typeof style.fontStyle === "string" ? { fontStyle: style.fontStyle } : {}),
          fill: typeof style.color === "string" ? style.color : "#595959",
          align: "center",
          verticalAlign: "middle",
        },
      });
      // An automatic plot area keeps the room the title took at its own
      // place: PowerPoint does not lay the plot out again when the title is
      
      // a chart with no title has it). The host measures that room.
      next[axisKey] = withFirst(next[axisKey], (entry) => ({
        ...entry,
        name: "",
        ...(layout?.plotArea ? {} : { [OFFICE_PLACED_AXIS_TITLE_ROOM]: name }),
      }));
      placedAxisTitles[physical === "X" ? "axisX" : "axisY"] = name;
    }
  }
  // A 3-D chart's depth axis title placed by hand (`elementLayout.
  
  // level where it was put, in the axis titles' text (`cartesian-3d` draws
  // an unplaced one beside the depth labels).
  const seriesTitle = attrs.format?.seriesAxisTitle;
  const seriesPlace = layout?.axisTitleSeries;
  // (A depth axis title is valid on a chart with a depth axis alone:
  // `validateChartBlockAttrsV2`.)
  if (size && seriesPlace && typeof seriesTitle === "string" && seriesTitle) {
    const axes = [first(next.xAxis), first(next.yAxis)];
    const named = axes.find((axis) => record(axis?.nameTextStyle) && Object.keys(axis!.nameTextStyle as object).length);
    const style = (named?.nameTextStyle ?? { fontWeight: "bold" }) as Record<string, unknown>;
    const fontSize = typeof style.fontSize === "number" ? style.fontSize : 12;
    const family = typeof style.fontFamily === "string" ? style.fontFamily : chartFamily;
    const weight = style.fontWeight === undefined ? "normal" : String(style.fontWeight);
    const lines = seriesTitle.split("\n");
    const along = Math.max(...lines.map((line) => measure(line, `${weight} ${fontSize}px ${family}`)));
    const lineHeight = Math.round(fontSize * 1.2 * 100) / 100;
    graphics.push({
      id: CHART_PLACED_AXIS_TITLE_IDS.axisSeries,
      type: "text",
      silent: true,
      z: 50,
      x: seriesPlace.x * size.width + along / 2,
      y: seriesPlace.y * size.height + (lines.length * lineHeight) / 2,
      style: {
        text: seriesTitle,
        fontSize,
        fontFamily: family,
        fontWeight: weight,
        lineHeight,
        ...(typeof style.fontStyle === "string" ? { fontStyle: style.fontStyle } : {}),
        fill: typeof style.color === "string" ? style.color : "#595959",
        align: "center",
        verticalAlign: "middle",
      },
    });
  }
  if (graphics.length) next.graphic = [...(Array.isArray(next.graphic) ? next.graphic : []), ...graphics];

  // A data label dragged by hand: its offset from where the chart puts it,
  // a fraction of the chart (`c:dLbl/c:layout/c:manualLayout`, factor mode).
  if (offsets.length && size && Array.isArray(next.series)) {
    const drawn = next.series.slice();
    // The label's drawn place: its series by name (a hidden series is not
    // drawn, a namesake carries the planner's mark), its point among the
    // plotted ones (Chart Filters leave hidden categories out), and a pie's
    // slice by its name (a pie of pie draws its points sorted by size).
    const model = record(attrs.data) ? (attrs.data as { kind?: unknown; categories?: unknown }) : null;
    const categories =
      model?.kind === "category" && Array.isArray(model.categories) ? (model.categories as string[]) : null;
    const hiddenPoints = new Set(attrs.format?.hiddenCategories ?? []);
    const plotted = categories?.map((_, index) => index).filter((index) => !hiddenPoints.has(index)) ?? null;
    const sliceNames =
      attrs.chartType === "pie" && categories && plotted
        ? officeChartPieSliceNames(plotted.map((index) => categories[index]))
        : null;
    // Leader lines to labels dragged off their points (Excel 2013+ on a
    // column, bar, line, area or XY chart, the labels' c15 extension;
    // C16_testDataPointLabelCustomPos, X-086), in the lines' own paint
    // (Office's default tx1 35 %).
    const leaderStyle = record(attrs.format?.dataLabelLeaderLineStyle) ? attrs.format.dataLabelLeaderLineStyle : undefined;
    const leaderOutline = typeof leaderStyle?.outline === "string" ? leaderStyle.outline : "#a6a6a6";
    const leaderLine =
      attrs.format?.dataLabelLeaderLines === true &&
      chartSupportsDraggedLabelLeaderLines(attrs.chartType, (attrs as { chartSubtype?: unknown }).chartSubtype) &&
      leaderOutline !== "none"
        ? {
            show: true,
            // Its own width where the lines state one (C29: 7pt), else Office's 0.75pt.
            lineStyle: {
              color: leaderOutline,
              width: ptToPx(
                record(leaderStyle?.shapePaint) && typeof leaderStyle.shapePaint.outlineWidthPt === "number"
                  ? leaderStyle.shapePaint.outlineWidthPt
                  : 0.75,
              ),
            },
          }
        : null;
    const leaderShifts = new Map<number, Map<number, [number, number]>>();
    const offsetSigns = chartDataLabelOffsetSigns(attrs);
    for (const entry of offsets) {
      const index = drawn.findIndex(
        (series) =>
          model !== null &&
          record(series) &&
          typeof series.name === "string" &&
          officeChartHitSeries(attrs as ChartBlockAttrsV2, series.name) === entry.seriesIndex,
      );
      const at = index >= 0 ? index : entry.seriesIndex;
      const series = drawn[at];
      if (!record(series) || !Array.isArray(series.data)) continue;
      const place = plotted ? plotted.indexOf(entry.pointIndex!) : entry.pointIndex!;
      const drawnIndex =
        place >= 0 && sliceNames
          ? series.data.findIndex((item) => officeChartPieItemPoint(item, sliceNames) === place)
          : place;
      if (drawnIndex < 0) continue;
      const data = series.data.slice();
      const item = data[drawnIndex];
      if (item === undefined) continue;
      const own: Record<string, unknown> = record(item) ? item : { value: item };
      // Along the value axis' own direction (`chartDataLabelOffsetSigns`).
      const shift: [number, number] = [
        entry.offset!.x * offsetSigns.x * size.width,
        entry.offset!.y * offsetSigns.y * size.height,
      ];
      // A pie's label is moved by the label layout too: the engine lays an
      // outside pie label out itself and leaves its `offset` unused (CDC
      // p51/p55: labels dragged from outside their slices stayed there).
      if (leaderLine || series.type === "pie") {
        // Moved by the label layout, so the engine draws the label's leader
        // line from its point (an `offset` moves the words alone).
        if (leaderLine) data[drawnIndex] = { ...own, labelLine: leaderLine };
        const moved = new Map(leaderShifts.get(at) ?? []);
        moved.set(drawnIndex, shift);
        leaderShifts.set(at, moved);
      } else {
        // The engine moves a label by its offset in the label's own turned
        // frame: a turned label's shift is turned back into it, so the words
        
        // turned a quarter and dragged down its column landed a column to
        // the right).
        const ownLabel = record(own.label) ? own.label : {};
        const seriesLabel = record(series.label) ? series.label : {};
        const rotate = typeof ownLabel.rotate === "number" ? ownLabel.rotate : typeof seriesLabel.rotate === "number" ? seriesLabel.rotate : 0;
        const turn = (rotate * Math.PI) / 180;
        const offset: [number, number] = rotate
          ? [
              shift[0] * Math.cos(turn) - shift[1] * Math.sin(turn),
              shift[0] * Math.sin(turn) + shift[1] * Math.cos(turn),
            ]
          : shift;
        data[drawnIndex] = {
          ...own,
          label: {
            ...ownLabel,
            offset,
          },
        };
      }
      drawn[at] = { ...series, data };
    }
    for (const [at, moved] of leaderShifts) {
      const series = drawn[at];
      if (!record(series)) continue;
      const base = series.labelLayout;
      drawn[at] = {
        ...series,
        labelLayout: (params: { dataIndex: number; labelRect?: { width: number; height: number } }) => {
          const own = typeof base === "function" ? (base as (p: unknown) => unknown)(params) : base;
          const layout = record(own) ? own : {};
          const shift = moved.get(params.dataIndex);
          if (!shift) return layout;
          const box = params.labelRect;
          if (
            series.type === "pie" &&
            box &&
            typeof layout.x === "number" &&
            typeof layout.y === "number" &&
            layout.align === "center" &&
            layout.verticalAlign === "top"
          ) {
            // A pie label dragged from its Outside End place: moved there by
            // the drag, then kept inside the chart, as PowerPoint keeps it
            // (pie lab 2: labels dragged past the top stand on the chart's
            // top edge).
            const left = Math.min(Math.max(layout.x + shift[0] - box.width / 2, 0), size.width - box.width);
            const top = Math.min(Math.max(layout.y + shift[1], 0), size.height - box.height);
            // Its leader line from the slice's arc to the box's nearest point
            // (drawn where the chart shows leader lines), none when the box
            // stands on the arc.
            const arc = Array.isArray(layout.officeArc) ? (layout.officeArc as [number, number]) : null;
            const near: [number, number] | null = arc
              ? [Math.min(Math.max(arc[0], left), left + box.width), Math.min(Math.max(arc[1], top), top + box.height)]
              : null;
            const { officeArc: _arc, ...rest } = layout;
            return {
              ...rest,
              x: left + box.width / 2,
              y: top,
              ...(arc && near ? { labelLinePoints: Math.hypot(near[0] - arc[0], near[1] - arc[1]) > 12 ? [arc, near, near] : [arc, arc, arc] } : {}),
            };
          }
          return {
            ...layout,
            dx: (typeof layout.dx === "number" ? layout.dx : 0) + shift[0],
            dy: (typeof layout.dy === "number" ? layout.dy : 0) + shift[1],
          };
        },
      };
    }
    next.series = drawn;
  }
  return { option: next, placedAxisTitles, ...(legendEntries ? { legendEntries } : {}) };
}

/**
 * A legend placed by hand is the box it was given (`c:manualLayout` w/h):
 * its entries are laid out inside it, and its selection frame and sizing
 * handles stand on that box, not on the entries' own extent (a frame on the
 * entries shrank the legend by its padding at every move). Px of a picture
 * of `width` x `height`; null for a legend the chart places itself.
 */
export function placedChartLegendBox(
  attrs: Pick<ChartBlockAttrsV2, "format"> | null | undefined,
  width: number,
  height: number,
): { x: number; y: number; width: number; height: number } | null {
  const legend = normalizeChartElementLayout(attrs?.format?.elementLayout)?.legend;
  // A legend placed by its corner alone has no box of its own: its frame
  // stands on its entries, as an automatic legend's does.
  return chartElementPlaceIsBox(legend)
    ? { x: legend.x * width, y: legend.y * height, width: legend.w * width, height: legend.h * height }
    : null;
}








export function chartLegendEntryOfPosition(
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "format">,
): (position: number) => number {
  const byCategory = attrs.chartType === "pie";
  const filtered = new Set((byCategory ? attrs.format?.hiddenCategories : attrs.format?.hiddenSeries) ?? []);
  if (filtered.size === 0) return (position) => position;
  return (position) => {
    let seen = -1;
    for (let index = 0; ; index += 1) {
      if (filtered.has(index)) continue;
      seen += 1;
      if (seen === position) return index;
    }
  };
}
