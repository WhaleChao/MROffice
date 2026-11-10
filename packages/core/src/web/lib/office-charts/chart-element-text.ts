/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The text of one chart element, as Office formats it.
 *
 * Office formats what is selected: "To change the formatting for all of the
 * text in the chart at the same time, click the blank area between the border
 * of the chart and the plot area to select the chart area"; to format one
 * element, "click the chart text" (Microsoft Support, "Change the font and
 * size of text in a chart"). The chart area's text is `format.text`
 * (`c:chartSpace/c:txPr`), which every element inherits; an element's own
 * text sits on its style (`chartTitleStyle.text`, `c:title/c:txPr`) and wins
 * for that element alone. Shared by Docs, the deck and the worksheet.
 */
import type { ChartBlockAttrsV2, ChartBlockFormat } from "../../components/console/editor/extensions/chartBlockShared";
import type { OfficeHomeTextCommand } from "../office-home-object-target";
import { officeChartTextCommandOf } from "../office-home-object-target";

import {
  applyOfficeChartTextCommand,
  normalizeOfficeChartText,
  type OfficeChartText,
  type OfficeChartTextCommand,
} from "./chart-text";
import { MAX_CHART_TEXT_CHARS } from "./chart-limits.generated";
import { chartTextSelection, readChartTextHost, type OfficeChartTextValue } from "./chart-text-editing";
import { chartTitleRunStyleAt, chartTitleRunsFromPieces, chartTitleRunsWithCommand } from "./chart-title-runs";

/** What Home's Font group formats on a chart: the whole chart, or one
 * element -- the title, the legend, the data labels, an axis (its labels)
 * or an axis title, by the model axis. */
export type OfficeChartTextElement =
  | "chart"
  | "title"
  | "legend"
  | "dataLabels"
  | "axisCategory"
  | "axisValue"
  | "axisTitleCategory"
  | "axisTitleValue"
  /** One series' data labels (the series selected), one label (its point). */
  | `seriesLabels:${number}`
  | `pointLabel:${number}:${number}`;

/** The series and point a labels element names, or null. */
function labelsRule(element: OfficeChartTextElement): { seriesIndex: number; pointIndex?: number } | null {
  const series = /^seriesLabels:(\d+)$/.exec(element);
  if (series) return { seriesIndex: Number(series[1]) };
  const point = /^pointLabel:(\d+):(\d+)$/.exec(element);
  if (point) return { seriesIndex: Number(point[1]), pointIndex: Number(point[2]) };
  return null;
}

const ELEMENT_STYLE_KEY = {
  title: "chartTitleStyle",
  legend: "legendStyle",
  dataLabels: "dataLabelStyle",
} as const satisfies Partial<Record<OfficeChartTextElement, keyof ChartBlockFormat>>;

/** The element's style row among an axis-keyed list (`axisStyles`,
 * `axisTitleStyles`), or null for an element with a style of its own. */
function axisRule(
  element: OfficeChartTextElement,
): { key: "axisStyles" | "axisTitleStyles"; axis: "category" | "value" } | null {
  switch (element) {
    case "axisCategory":
      return { key: "axisStyles", axis: "category" };
    case "axisValue":
      return { key: "axisStyles", axis: "value" };
    case "axisTitleCategory":
      return { key: "axisTitleStyles", axis: "category" };
    case "axisTitleValue":
      return { key: "axisTitleStyles", axis: "value" };
    default:
      return null;
  }
}

/**
 * The element Home formats for the Format pane's Current Selection (the
 * element a click on the chart selected, or the pane's own pick): the chart
 * for the chart area, the plot area and anything that carries no text of its
 * own (a series, a point, the gridlines).
 */
export function officeChartTextElementOfArea(area: string | null | undefined): OfficeChartTextElement {
  switch (area) {
    case "chartTitleStyle":
      return "title";
    case "legendStyle":
      return "legend";
    case "dataLabelStyle":
      return "dataLabels";
    case "axis:category":
      return "axisCategory";
    case "axis:value":
      return "axisValue";
    case "axisTitle:category":
      return "axisTitleCategory";
    case "axisTitle:value":
      return "axisTitleValue";
    default: {
      // A series selected: its data labels (Office's Font group with a
      // series selected formats the series' labels); a point: its label; a
      // pie's slice: that point's label.
      const series = /^series:(\d+)$/.exec(area ?? "");
      if (series) return `seriesLabels:${Number(series[1])}`;
      const point = /^point:(\d+):(\d+)$/.exec(area ?? "");
      if (point) return `pointLabel:${Number(point[1])}:${Number(point[2])}`;
      const slice = /^slice:(\d+)$/.exec(area ?? "");
      if (slice) return `pointLabel:0:${Number(slice[1])}`;
      return "chart";
    }
  }
}

function labelOverride(
  attrs: Pick<ChartBlockAttrsV2, "format">,
  rule: { seriesIndex: number; pointIndex?: number },
): { font?: OfficeChartText } | undefined {
  return (attrs.format?.dataLabelOverrides ?? []).find(
    (entry) => entry.seriesIndex === rule.seriesIndex && entry.pointIndex === rule.pointIndex,
  );
}

/** The chart title being typed in a host's chart, read from the one title
 * input the shared `OfficeChartTitleInput` mounts (`data-office-chart-title`);
 * the input stays mounted while the ribbon has the focus, so the command a
 * ribbon click runs sees the element the user is in. Axis titles are formatted
 * as the chart's text (their part has no style of its own in the model). */
export function officeChartTitleEditElement(root: ParentNode | null | undefined): OfficeChartTextElement {
  const input = root?.querySelector<HTMLElement>("[data-office-chart-title]");
  return input?.getAttribute("data-office-chart-title") === "title" ? "title" : "chart";
}

/** The characters selected inside the chart title being typed (the title
 * input's own range), which Home formats alone ("select the individual
 * characters you want to format"); null for no range. */
export function officeChartTitleEditRange(
  root: ParentNode | null | undefined,
): { start: number; end: number; typed?: OfficeChartTextValue } | null {
  const input = root?.querySelector<HTMLInputElement>('input[data-office-chart-title="title"]');
  if (input) {
    const start = input.selectionStart ?? 0;
    const end = input.selectionEnd ?? 0;
    return end > start ? { start, end } : null;
  }
  // The title typed in place (`OfficeChartTextLayer`): its selected
  // characters, and the words as typed so far, which the command acts on.
  const host = root?.querySelector<HTMLElement>('[data-office-chart-title="title"][contenteditable="true"]');
  if (!host) return null;
  const range = chartTextSelection(host);
  return range && range.end > range.start ? { ...range, typed: readChartTextHost(host) } : null;
}

/**
 * The chart with a text typed in place: the title as its runs (each piece
 * keeps the format of the run it was typed in), an axis title as its words
 * (gone when they are cleared, as an axis title is). Null when nothing changed.
 */
export function officeChartWithTypedText(
  attrs: ChartBlockAttrsV2,
  part: "title" | "axisX" | "axisY",
  value: OfficeChartTextValue,
): ChartBlockAttrsV2 | null {
  if (part === "title") {
    const typed = chartTitleRunsFromPieces(attrs.title ?? "", attrs.titleRuns, value.pieces, MAX_CHART_TEXT_CHARS);
    if (typed.title === (attrs.title ?? "") && JSON.stringify(typed.titleRuns ?? null) === JSON.stringify(attrs.titleRuns ?? null))
      return null;
    const { titleRuns: _previous, ...rest } = attrs;
    return { ...rest, title: typed.title, ...(typed.titleRuns ? { titleRuns: typed.titleRuns } : {}) };
  }
  const words = value.text.replace(/\r\n?/g, "\n").slice(0, MAX_CHART_TEXT_CHARS);
  const next = words.trim() ? words : undefined;
  const key = part === "axisX" ? "axisTitleX" : "axisTitleY";
  if ((attrs[key] ?? undefined) === next) return null;
  const { [key]: _previous, ...rest } = attrs;
  return next === undefined ? (rest as ChartBlockAttrsV2) : { ...rest, [key]: next };
}

/** The text the characters at the range's start show (the run's over the title's). */
export function officeChartTitleRangeText(
  attrs: Pick<ChartBlockAttrsV2, "title" | "titleRuns" | "format">,
  range: { start: number; end: number },
): OfficeChartText | undefined {
  const inherited = officeChartElementText(attrs, "title");
  const own = chartTitleRunStyleAt(attrs.title ?? "", attrs.titleRuns, range.start);
  if (!inherited) return own;
  if (!own) return inherited;
  return normalizeOfficeChartText({ ...inherited, ...own }) ?? own;
}

/** The chart with a Home command on the title's characters in the range. */
export function officeChartWithTitleRangeCommand(
  base: ChartBlockAttrsV2,
  range: { start: number; end: number; typed?: OfficeChartTextValue },
  command: OfficeHomeTextCommand,
): ChartBlockAttrsV2 | null {
  // The words typed so far are the title the characters are counted in.
  const attrs = (range.typed && officeChartWithTypedText(base, "title", range.typed)) || base;
  const chartCommand = officeChartTextCommandOf(command, officeChartTitleRangeText(attrs, range));
  if (!chartCommand) return null;
  const runs = chartTitleRunsWithCommand(
    attrs.title ?? "",
    attrs.titleRuns,
    range.start,
    range.end,
    chartCommand,
    officeChartElementText(attrs, "title"),
  );
  if (runs === null) return null;
  const { titleRuns: _previous, ...rest } = attrs;
  return runs ? { ...rest, titleRuns: runs } : rest;
}

/** The element's own text (what its style states), or the chart's. */
export function officeChartElementOwnText(
  attrs: Pick<ChartBlockAttrsV2, "format">,
  element: OfficeChartTextElement,
): OfficeChartText | undefined {
  if (element === "chart") return attrs.format?.text;
  const labels = labelsRule(element);
  if (labels) return labelOverride(attrs, labels)?.font;
  const rule = axisRule(element);
  if (rule) {
    const rows = (attrs.format?.[rule.key] ?? []) as readonly { axis: string; text?: OfficeChartText }[];
    return rows.find((row) => row.axis === rule.axis)?.text;
  }
  const style = attrs.format?.[ELEMENT_STYLE_KEY[element as keyof typeof ELEMENT_STYLE_KEY]] as
    | { text?: OfficeChartText }
    | undefined;
  return style?.text;
}

/** The text the element shows: its own over the chart's (Office's inheritance). */
/** The size Office draws a chart title in when nothing states one: 14pt
 * (the chart's own 10pt text scaled as Excel's default title is). The
 * renderer and the Home Font Size box read the same number. */
export const OFFICE_CHART_TITLE_SIZE_PT = 14;

export function officeChartElementText(
  attrs: Pick<ChartBlockAttrsV2, "format">,
  element: OfficeChartTextElement,
): OfficeChartText | undefined {
  const own = officeChartElementOwnText(attrs, element);
  if (element === "chart") return own;
  const inherited = officeChartElementInheritedText(attrs, element);
  if (!inherited) return own;
  if (!own) return inherited;
  return normalizeOfficeChartText({ ...inherited, ...own }) ?? own;
}

/** What the element inherits under its own text: the chart's text; for a
 * label, then the data labels' own, then its series' (a point's label), as
 * Office draws the inheritance. */
export function officeChartElementInheritedText(
  attrs: Pick<ChartBlockAttrsV2, "format">,
  element: OfficeChartTextElement,
): OfficeChartText | undefined {
  if (element === "chart") return undefined;
  const labels = labelsRule(element);
  const layers: (OfficeChartText | undefined)[] = labels
    ? [
        attrs.format?.text,
        attrs.format?.dataLabelStyle?.text,
        labels.pointIndex === undefined ? undefined : labelOverride(attrs, { seriesIndex: labels.seriesIndex })?.font,
      ]
    : [attrs.format?.text];
  const stated = layers.filter((layer): layer is OfficeChartText => Boolean(layer));
  if (stated.length === 0) return undefined;
  if (stated.length === 1) return stated[0];
  return normalizeOfficeChartText(Object.assign({}, ...stated)) ?? stated[stated.length - 1];
}

/** The chart with the element's own text replaced (`undefined` removes it). */
export function officeChartWithElementText(
  attrs: ChartBlockAttrsV2,
  element: OfficeChartTextElement,
  text: OfficeChartText | undefined,
): ChartBlockAttrsV2 {
  const format: Record<string, unknown> = { ...(attrs.format ?? {}) };
  const rule = axisRule(element);
  const labels = labelsRule(element);
  if (element === "chart") {
    if (text) format.text = text;
    else delete format.text;
  } else if (labels) {
    const rows = ((format.dataLabelOverrides as readonly Record<string, unknown>[] | undefined) ?? []);
    const index = rows.findIndex(
      (row) => row.seriesIndex === labels.seriesIndex && row.pointIndex === labels.pointIndex,
    );
    const own: Record<string, unknown> = {
      ...(index >= 0 ? rows[index] : { seriesIndex: labels.seriesIndex, ...(labels.pointIndex === undefined ? {} : { pointIndex: labels.pointIndex }) }),
    };
    if (text) own.font = text;
    else delete own.font;
    const keeps = Object.keys(own).some((key) => key !== "seriesIndex" && key !== "pointIndex");
    const next = rows.filter((_, at) => at !== index);
    if (keeps) next.splice(index >= 0 ? index : next.length, 0, own);
    if (next.length) format.dataLabelOverrides = next;
    else delete format.dataLabelOverrides;
  } else if (rule) {
    const rows = ((format[rule.key] as readonly Record<string, unknown>[] | undefined) ?? []).filter(
      (row) => row.axis !== rule.axis,
    );
    const own: Record<string, unknown> = {
      ...(((format[rule.key] as readonly Record<string, unknown>[] | undefined) ?? []).find(
        (row) => row.axis === rule.axis,
      ) ?? { axis: rule.axis }),
    };
    if (text) own.text = text;
    else delete own.text;
    const next = Object.keys(own).length > 1 ? [...rows, own] : rows;
    // The model keeps an axis's style rows in model order (category, value).
    next.sort((a, b) => String(a.axis).localeCompare(String(b.axis)));
    if (next.length) format[rule.key] = next;
    else delete format[rule.key];
  } else {
    const key = ELEMENT_STYLE_KEY[element as keyof typeof ELEMENT_STYLE_KEY];
    const style: Record<string, unknown> = { ...((format[key] as Record<string, unknown> | undefined) ?? {}) };
    if (text) style.text = text;
    else delete style.text;
    if (Object.keys(style).length) format[key] = style;
    else delete format[key];
  }
  return { ...attrs, format: format as ChartBlockFormat };
}

/**
 * One Home command on one element. A toggle (Bold, Italic, Underline) turns
 * the element's SHOWN state, so a title that inherits bold from the chart
 * states bold off for itself; a value (size, face, colour) is written on the
 * element. On the chart itself this is `applyOfficeChartTextCommand`. Null
 * when nothing changes.
 */
export function officeChartWithElementTextCommand(
  attrs: ChartBlockAttrsV2,
  element: OfficeChartTextElement,
  command: OfficeHomeTextCommand,
): ChartBlockAttrsV2 | null {
  const shown = officeChartElementText(attrs, element);
  const chartCommand = officeChartTextCommandOf(command, shown);
  if (!chartCommand) return null;
  const own = officeChartElementOwnText(attrs, element);
  const inherited = officeChartElementInheritedText(attrs, element);
  const next = applyOfficeChartElementTextCommand(own, shown, inherited, chartCommand);
  // A colour is the text's one fill: Format ▸ Text Fill and Home ▸ Font
  // Color write the same run property in Office, so the colour chosen last
  // shows. An element's Text Fill (`textFill`, drawn over its text colour)
  // gives way to the Font Color chosen after it (audit
  // P-CHART-TEXT-COLOR-SCOPE: a title filled red stayed red after Home ▸
  // blue with the title selected as an object).
  const filled = command.kind === "color" ? withoutElementTextFill(attrs, element) : attrs;
  if (next === null) return filled === attrs ? null : filled;
  return officeChartWithElementText(filled, element, next);
}

/** The chart with the element's own Text Fill taken off, or the chart itself
 * when the element has none. */
function withoutElementTextFill(
  attrs: ChartBlockAttrsV2,
  element: OfficeChartTextElement,
): ChartBlockAttrsV2 {
  const format = attrs.format as Record<string, unknown> | undefined;
  if (!format) return attrs;
  const rule = axisRule(element);
  if (rule) {
    const rows = format[rule.key] as readonly Record<string, unknown>[] | undefined;
    if (!rows?.some((row) => row.axis === rule.axis && row.textFill !== undefined)) return attrs;
    const next = rows
      .map((row) => {
        if (row.axis !== rule.axis) return row;
        const { textFill: _fill, ...rest } = row;
        void _fill;
        return rest;
      })
      .filter((row) => Object.keys(row).length > 1);
    const nextFormat = { ...format };
    if (next.length) nextFormat[rule.key] = next;
    else delete nextFormat[rule.key];
    return { ...attrs, format: nextFormat as ChartBlockFormat };
  }
  const key = ELEMENT_STYLE_KEY[element as keyof typeof ELEMENT_STYLE_KEY];
  if (!key) return attrs;
  const style = format[key] as Record<string, unknown> | undefined;
  if (!style || style.textFill === undefined) return attrs;
  const { textFill: _fill, ...rest } = style;
  void _fill;
  const nextFormat = { ...format };
  if (Object.keys(rest).length) nextFormat[key] = rest;
  else delete nextFormat[key];
  return { ...attrs, format: nextFormat as ChartBlockFormat };
}

/** The element's own text after the command, against what it shows. */
export function applyOfficeChartElementTextCommand(
  own: OfficeChartText | undefined,
  shown: OfficeChartText | undefined,
  inherited: OfficeChartText | undefined,
  command: OfficeChartTextCommand,
): OfficeChartText | null | undefined {
  if (!inherited) return applyOfficeChartTextCommand(own, command);
  const draft: Record<string, unknown> = { ...(own ?? {}) };
  const settle = (key: string, value: unknown) => {
    // An own value equal to the inherited one says nothing: drop it (a
    // toggle the inheritance leaves unstated is off).
    const below = (inherited as Record<string, unknown>)[key] ?? (typeof value === "boolean" ? false : undefined);
    if (value === undefined || value === below) delete draft[key];
    else draft[key] = value;
  };
  switch (command.kind) {
    case "toggle":
      settle(command.field, shown?.[command.field] === true ? false : true);
      break;
    case "underline":
      // The model states no "no underline": an element back to the chart's
      // underline says nothing of its own.
      settle("underline", shown?.underline === command.variant ? undefined : command.variant);
      break;
    default: {
      const applied = applyOfficeChartTextCommand(own, command);
      if (applied === null) return null;
      return applied;
    }
  }
  const next = Object.keys(draft).length ? normalizeOfficeChartText(draft) : undefined;
  if (Object.keys(draft).length && !next) return null;
  return JSON.stringify(next ?? null) === JSON.stringify(own ?? null) ? null : next;
}
