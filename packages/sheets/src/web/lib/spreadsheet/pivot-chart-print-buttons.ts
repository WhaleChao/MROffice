/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PivotChart's field buttons as Windows Excel prints them (Save as PDF /
 * Print; lookwin lw121/fb, R480-X04's PivotChart with each switch of
 * `c14:pivotOptions` and `c16:showExpandCollapseFieldButtons`):
 *
 * - every button is a white box with a 0.75 pt #D9D9D9 line, 13.78 pt high,
 *   its field's name in black Calibri 10 pt from the box's left edge, the
 *   baseline 2.67 pt over the box's bottom; no dropdown arrow is printed (the
 *   arrow's room stays in the box);
 * - the filter and value buttons stand in one row from (8, 6) of the chart
 *   area, the axis (category) buttons in one row from 8 pt in and 6 pt over
 *   its bottom, side by side;
 * - a top row takes 19.78 pt off the top of what the chart lays its plot,
 *   legend and titles out in; an axis row, or the Expand/Collapse Entire Field
 *   buttons of a chart with two or more axis fields (which do not print), the
 *   same off its bottom. Hide All (`dropZonesVisible` 0, or no c14 part at
 *   all) takes nothing;
 * - a legend (series) button stands at the legend's top left.
 *
 * A box's width is the name measured in Tahoma 8 pt (the Windows UI face Excel
 * sizes its buttons with, whatever face prints the name), its design width
 * (its trailing spaces left out) rounded up to a whole pixel at 96 dpi, and a
 * pad: 8 px for a value button, which has no dropdown, 22 px for one with a
 * dropdown, 32 px for a field that hides some of its items (its filter mark
 * stands beside the arrow). Exact on all twelve of Excel's prints (lookwin
 * lw124/fb: R480-X04 and its variants, R476-X14): Avg 26 px, Sum of COUNT 79,
 * a 70-character value name 356, ADT 42 (52 filtered), Surface 58, MAX PAGE
 * 72, # OF PAGES 81, a 47-character axis name 236, "Applied Thickness " 106
 * (116 filtered). A name wider than its box is set narrower to fit it. The API's writer prints the same (notebooks.py
 * `pivot_chart_print_buttons`).
 */

/** The printed box, in points. */
export const PIVOT_PRINT_BUTTON_HEIGHT_PT = 13.78;
export const PIVOT_PRINT_BUTTON_LEFT_PT = 8;
export const PIVOT_PRINT_BUTTON_EDGE_PT = 6;
/** What a row of buttons takes off the chart's layout. */
export const PIVOT_PRINT_BUTTON_ROW_ROOM_PT = 19.78;
export const PIVOT_PRINT_BUTTON_BASELINE_PT = 2.67;
export const PIVOT_PRINT_BUTTON_LINE_PT = 0.75;
export const PIVOT_PRINT_BUTTON_LINE_COLOR = "#D9D9D9";
export const PIVOT_PRINT_BUTTON_FONT_PT = 10;
/** The legend's right edge stands this far in from the chart area's. */
export const PIVOT_PRINT_LEGEND_RIGHT_PT = 6.74;
/** A legend entry's room past its name: the key's inset from the legend's
 * left (3.97), the key (4.94), the gap to the name (2.19) and the room after
 * it (4.85), in points (lw124/fb u_longentries). */
export const PIVOT_PRINT_LEGEND_ENTRY_ROOM_PT = 15.95;
/** A legend entry's pitch and the gap under the legend's button (9 pt entries). */
export const PIVOT_PRINT_LEGEND_ENTRY_PT = 16.87;
export const PIVOT_PRINT_LEGEND_BUTTON_GAP_PT = 4.28;
/** The last entry's share of the centred block: its words' height, not a
 * full pitch (lw121/fb a_all and e_ser: the button at 101.29 and 91.4 pt, a
 * block of 46.78 pt for two 9 pt entries). */
export const PIVOT_PRINT_LEGEND_LAST_ENTRY_PT = 11.85;

// Calibri's design advances (thousandths of an em), printable ASCII from the space.
const CALIBRI_ADVANCES = [
  226, 326, 401, 498, 507, 715, 682, 221, 303, 303, 498, 498, 250, 306, 252, 386,
  507, 507, 507, 507, 507, 507, 507, 507, 507, 507, 268, 268, 498, 498, 498, 463,
  894, 579, 544, 533, 615, 488, 459, 631, 623, 252, 319, 520, 420, 855, 646, 662,
  517, 673, 543, 459, 487, 642, 567, 890, 519, 487, 468, 307, 386, 307, 498, 498,
  291, 479, 525, 423, 525, 498, 305, 471, 525, 230, 239, 455, 230, 799, 525, 527,
  525, 525, 349, 391, 335, 525, 452, 715, 433, 453, 395, 314, 460, 314, 498,
];

// Tahoma's design advances (units of 2048 an em), printable ASCII from the space.
const TAHOMA_ADVANCES = [
  640, 680, 822, 1490, 1118, 2000, 1380, 432, 784, 784, 1118, 1490, 620, 744, 620, 783,
  1118, 1118, 1118, 1118, 1118, 1118, 1118, 1118, 1118, 1118, 724, 724, 1490, 1490, 1490, 970,
  1862, 1228, 1207, 1230, 1389, 1149, 1068, 1367, 1383, 764, 853, 1204, 1019, 1578, 1367, 1449,
  1129, 1449, 1271, 1141, 1196, 1343, 1222, 1847, 1189, 1180, 1145, 784, 783, 784, 1490, 1118,
  1118, 1075, 1132, 945, 1132, 1078, 652, 1132, 1142, 468, 577, 1020, 468, 1720, 1142, 1112,
  1132, 1132, 738, 914, 685, 1142, 1020, 1520, 1014, 1020, 910, 984, 783, 984, 1490,
];

/** A button's pad past its name, in pixels at 96 dpi. */
export const PIVOT_PRINT_BUTTON_PAD_PX = { value: 8, dropdown: 22, filtered: 32 } as const;

/**
 * A printed button's width in points: `text` in Tahoma 8 pt, its trailing
 * spaces left out, rounded up to a whole pixel, and the pad of its kind. Tahoma 8 pt is 32/3 px an em, so a
 * design unit (2048 an em) is 1/192 px.
 */
export function pivotPrintButtonWidthPt(text: string, kind: keyof typeof PIVOT_PRINT_BUTTON_PAD_PX): number {
  let units = 0;
  for (const character of text.replace(/ +$/, "")) {
    const code = character.codePointAt(0) ?? 0;
    units += code >= 32 && code < 127 ? TAHOMA_ADVANCES[code - 32] ?? 1118 : code >= 0x2e80 ? 2048 : 1118;
  }
  return (Math.ceil(units / 192 - 1e-9) + PIVOT_PRINT_BUTTON_PAD_PX[kind]) * 0.75;
}

/** Calibri 10 pt's design width of `text`, in points (an ideograph an em). */
export function pivotPrintButtonTextWidthPt(text: string): number {
  let units = 0;
  for (const character of text) {
    const code = character.codePointAt(0) ?? 0;
    units += code >= 32 && code < 127 ? CALIBRI_ADVANCES[code - 32] ?? 500 : code >= 0x2e80 ? 1000 : 500;
  }
  return (units / 1000) * PIVOT_PRINT_BUTTON_FONT_PT;
}

/**
 * A button's name for a field the import keeps trimmed (a PivotTable kept
 * as the file wrote it): Excel shows the cache field's name as the source's
 * header cell writes it, a trailing space and all (R480-X04 "Applied
 * Thickness "); `header` is that cell's value.
 */
export function pivotPrintFieldNameAsWritten(name: string, header: unknown): string {
  return typeof header === "string" && header !== name && header.trim() === name ? header : name;
}

export interface PivotPrintButton {
  text: string;
  /** Its box's left and top in the chart area, and its width, in points. */
  leftPt: number;
  topPt: number;
  widthPt: number;
  /** The name's font size: 10 pt, less where the box is narrower than it. */
  fontPt: number;
}

export interface PivotPrintButtons {
  buttons: PivotPrintButton[];
  /** What the chart's picture gives up at its top and bottom, in points. */
  roomTopPt: number;
  roomBottomPt: number;
  /** With a legend button: its left (Excel widens the legend leftward to
   * it, its entries starting there) and where the entries' keys stand
   * centred, in points in the chart area. */
  legend?: { leftPt: number; entriesMiddlePt: number };
}

/** The fields a PivotChart's buttons name, and which switches show them. */
export interface PivotPrintButtonFields {
  filters: readonly string[];
  values: readonly string[];
  legend: readonly string[];
  axis: readonly string[];
  /** The Expand/Collapse Entire Field buttons show (two or more axis fields). */
  expandCollapse: boolean;
  /** The names of the fields among them that hide some of their items. */
  filtered?: readonly string[];
}

function button(text: string, kind: keyof typeof PIVOT_PRINT_BUTTON_PAD_PX, leftPt: number, topPt: number): PivotPrintButton {
  const textWidth = pivotPrintButtonTextWidthPt(text);
  const widthPt = pivotPrintButtonWidthPt(text, kind);
  const fontPt = textWidth > widthPt ? (PIVOT_PRINT_BUTTON_FONT_PT * widthPt) / textWidth : PIVOT_PRINT_BUTTON_FONT_PT;
  return { text, leftPt, topPt, widthPt, fontPt };
}

/**
 * The printed buttons of a chart `widthPt` x `heightPt` whose PivotTable
 * names `fields`, with `seriesCount` legend entries named `seriesNames`
 * (in `entryFontPt` Calibri); null when none prints and nothing is given up.
 * A legend button narrower than its entries is widened to them, as Excel
 * prints it (lw124/fb u_longentries: a 32-character entry, the button
 * 87 -> 131.6 pt, its right edge kept).
 */
export function pivotChartPrintButtons(
  fields: PivotPrintButtonFields,
  widthPt: number,
  heightPt: number,
  seriesCount: number,
  seriesNames: readonly string[] = [],
  entryFontPt = 9,
): PivotPrintButtons | null {
  const buttons: PivotPrintButton[] = [];
  const dropdown = (name: string): "filtered" | "dropdown" =>
    fields.filtered?.includes(name) ? "filtered" : "dropdown";
  let left = PIVOT_PRINT_BUTTON_LEFT_PT;
  for (const [names, valueRow] of [[fields.filters, false], [fields.values, true]] as const) {
    for (const name of names) {
      const next = button(name, valueRow ? "value" : dropdown(name), left, PIVOT_PRINT_BUTTON_EDGE_PT);
      buttons.push(next);
      left += next.widthPt;
    }
  }
  const roomTopPt = buttons.length ? PIVOT_PRINT_BUTTON_ROW_ROOM_PT : 0;
  const axisTop = heightPt - PIVOT_PRINT_BUTTON_EDGE_PT - PIVOT_PRINT_BUTTON_HEIGHT_PT;
  left = PIVOT_PRINT_BUTTON_LEFT_PT;
  for (const name of fields.axis) {
    const next = button(name, dropdown(name), left, axisTop);
    buttons.push(next);
    left += next.widthPt;
  }
  const roomBottomPt = fields.axis.length || fields.expandCollapse ? PIVOT_PRINT_BUTTON_ROW_ROOM_PT : 0;
  let legend: PivotPrintButtons["legend"];
  if (fields.legend.length) {
    // The legend (its buttons over its entries) stands centred between the
    // rows' rooms, its right edge where Excel keeps it.
    const entries = Math.max(1, seriesCount) * PIVOT_PRINT_LEGEND_ENTRY_PT;
    const block = fields.legend.length * PIVOT_PRINT_BUTTON_HEIGHT_PT + PIVOT_PRINT_LEGEND_BUTTON_GAP_PT
      + entries - PIVOT_PRINT_LEGEND_ENTRY_PT + PIVOT_PRINT_LEGEND_LAST_ENTRY_PT;
    let top = (roomTopPt + heightPt - roomBottomPt) / 2 - block / 2;
    let leftmost = widthPt;
    const entriesWidth = seriesNames.length
      ? Math.max(...seriesNames.map((name) => (pivotPrintButtonTextWidthPt(name) * entryFontPt) / PIVOT_PRINT_BUTTON_FONT_PT)) +
        PIVOT_PRINT_LEGEND_ENTRY_ROOM_PT
      : 0;
    for (const name of fields.legend) {
      const own = button(name, dropdown(name), 0, top);
      const sized = own.widthPt < entriesWidth ? { ...own, widthPt: entriesWidth } : own;
      buttons.push({ ...sized, leftPt: widthPt - PIVOT_PRINT_LEGEND_RIGHT_PT - sized.widthPt });
      leftmost = Math.min(leftmost, widthPt - PIVOT_PRINT_LEGEND_RIGHT_PT - sized.widthPt);
      top += PIVOT_PRINT_BUTTON_HEIGHT_PT;
    }
    // The entries' rows start at the buttons' bottom, each key on its row's
    // middle (R480-X04: keys 123.51 and 140.38 under 115.07).
    legend = { leftPt: leftmost, entriesMiddlePt: top + entries / 2 };
  }
  if (!buttons.length && !roomBottomPt) return null;
  return { buttons, roomTopPt, roomBottomPt, ...(legend ? { legend } : {}) };
}

/**
 * A PivotChart prints its legend only with its legend (series) field button
 * (Windows Excel, lookwin lw121/fb and lw125/fb9): R480-X04's chart with that
 * button switched off (`dropZoneSeries` 0, the other buttons on or off),
 * Hide All (`dropZonesVisible` 0) or no c14 part printed no legend, its plot
 * still laid out beside the legend's room (right edge 495.6 pt, as with the
 * legend and no button; 453.6 with both); R476-X14's one-series chart, its
 * PivotTable with no column field (so no legend button), printed none though
 * the part states a right legend, its plot short of it alike. The chart
 * engine draws it with its legend's room and no legend (`legendRoomOnly`).
 * A chart whose PivotTable is not found prints its legend.
 */
export function pivotChartPrintsLegend(fields: PivotPrintButtonFields | null): boolean {
  return Boolean(fields?.legend.length);
}

/** A PivotChart's automatic title as Windows Excel prints it where the chart
 * states no text of its own: 18 pt bold black (R476-X14, lookwin lw125/fb9:
 * Calibri Bold 18 over axes in 9 pt grey). */
export const PIVOT_CHART_AUTOMATIC_TITLE_TEXT = { fontSizePt: 18, bold: true, color: "#000000" } as const;

/**
 * A printed PivotChart record with its automatic title in its words and
 * look: Windows Excel prints a PivotChart's automatic title (no `c:title`,
 * `autoTitleDeleted` 0: R476-X14's "Total"), which its import reads as
 * `format.automaticTitle`; an ordinary chart of that shape, read with none,
 * stays untitled (R462-X23). The words are the one series' name, else
 * `chartTitleWord` ("Chart Title" in the reader's language); the API's print
 * alike (spreadsheet_chart_render `with_pivot_chart_automatic_title`).
 */
export function pivotChartRecordWithAutomaticTitle<T extends object>(record: T, chartTitleWord: string): T {
  const model = (record as { chartV2?: unknown }).chartV2;
  if (!model || typeof model !== "object") return record;
  const attrs = model as {
    title?: unknown;
    chartType?: unknown;
    chartSubtype?: unknown;
    data?: { series?: readonly { name?: unknown }[] };
    format?: { automaticTitle?: unknown; chartTitleStyle?: { text?: object } };
  };
  if ((typeof attrs.title === "string" && attrs.title.trim()) || attrs.format?.automaticTitle !== true) return record;
  const series = Array.isArray(attrs.data?.series) ? attrs.data.series : [];
  const firstOnly = attrs.chartType === "pie" && attrs.chartSubtype !== "doughnut";
  const name = series.length === 1 || (firstOnly && series.length > 1) ? series[0]?.name : undefined;
  const words = typeof name === "string" && name.trim() ? name : chartTitleWord;
  const titleStyle = attrs.format.chartTitleStyle ?? {};
  return {
    ...record,
    chartV2: {
      ...attrs,
      title: words,
      format: {
        ...attrs.format,
        chartTitleStyle: { ...titleStyle, text: { ...PIVOT_CHART_AUTOMATIC_TITLE_TEXT, ...(titleStyle.text ?? {}) } },
      },
    },
  };
}
