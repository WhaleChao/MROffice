/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The chart engine's wire: what the browser (and the API's exports) send to
 * `POST /api/v1/engine/charts/{method}` and what comes back. Types and plain
 * constants only -- the planner and the drawing run on the server
 * (`engine-host/engines/charts.ts` -> `lib/office-charts/chart-render-server.ts`).
 *
 * A chart is answered as a picture: inert SVG in the chart's own pixels that
 * keeps ECharts' marks of which series and point each shape draws
 * (`ecmeta_series_index` / `ecmeta_data_index`, read by a click through
 * `officeChartSvgHitAt`) and MRAI's mark of which chart part a shape belongs
 * to (`ecmeta_part`: title, axisX, axisY, legend), plus the hit metadata a
 * host needs without planning the chart again: the drawn series (name, kind,
 * own paint), the colour cycle, and the boxes of the parts.
 */

/** One chart to draw. Exactly the request the API's exports send (schema
 * `mrai.chart-render.v1`), plus the optional fields the editors add. */
export interface ChartRenderRequest {
  /** "document" (default): `attrs` is a chart node's attrs. "presentation":
   * `block` is a deck chart block. "spreadsheet": `sheet` is a worksheet
   * chart's reading of its record and live range (`SpreadsheetChartV2Input`). */
  host?: "document" | "presentation" | "spreadsheet";
  /**
   * A worksheet chart's `SpreadsheetChartV2Input` (rows, record, theme, ...),
   * or -- a print -- `{ record, sheetName, cells, ...input fields }`: the
   * chart's record with its model over the data it plots, read into the input
   * the sheet's canvas would hand (`spreadsheetChartPrintInput`); `attrs` then
   * names the model a record that reading refuses is printed as.
   */
  sheet?: Record<string, unknown>;
  /** The chart node's attrs: V2 attrs plus the node's `source` / `documentChart`. */
  attrs?: Record<string, unknown>;
  /** A deck's chart block (`chartV2`, `source`, `chartWorkbook`, ...). */
  block?: Record<string, unknown>;
  /** The host document's attrs (its theme is what a Change Colors pick resolves against). */
  documentAttrs?: Record<string, unknown> | null;
  /** A workbook's theme keys (`themeId`, `nativeThemeColors`, ...): the sheet's accents. */
  workbookTheme?: Record<string, unknown> | null;
  /** Explicit theme accents (a deck's theme), accent1..6. */
  themeAccents?: string[];
  /**
   * A deck's body face (its theme's minor Latin font, as a CSS font stack):
   * the face a deck chart's text that names none is drawn in, as PowerPoint
   * draws it (`+mn-lt`).
   */
  themeFontFamily?: string;
  /** An imported chart's look given beside V2 attrs that do not carry it in
   * `source.nativeLook` (a style gallery tile of such a chart). */
  nativeLook?: Record<string, unknown> | null;
  /** The host's catalogue word for a Pie of Pie's combined slice ("Other"). */
  otherSliceLabel?: string;
  /** The number format the chart's sheet gives its values (Linked to
   * source), when the host read it already (else from `attrs.documentChart`). */
  linkedValueNumberFormat?: string;
  /** Answer each point's tooltip words too (`ChartRenderPicture.tips`). */
  tips?: boolean;
  width: number;
  height: number;
  /** Draw the title inside the picture (the Docs host does). */
  includeTitle?: boolean;
  /**
   * A deck chart printed on its own (the PDF): its data table's band is left
   * out of the picture, which is drawn above it, and the table comes back as
   * `ChartRenderPicture.dataTable` for the printer to draw under it.
   * A workbook chart (any other host) is drawn on its whole area instead,
   * as Excel lays it out: the plot stands above the table's band and the
   * legend below it, and the table comes back with the band's place in the
   * picture (`ChartRenderDataTable.topPx`) for the sheet and its print.
   */
  dataTableBand?: boolean;
  /**
   * The editors draw the chart's texts themselves (the title and the axis
   * titles): the picture keeps their place and box but not their words, and
   * the answer says where and how each is drawn (`ChartRenderPicture.texts`).
   * The words then exist once, in the editor's own text, which is also what
   * is typed in -- Office edits a title in place, as the text box it is.
   * Exports and prints ask without it and get the words in the picture.
   */
  textLayer?: boolean;
  locale?: string;
}

/** How a run of a chart text is drawn (picture px; `fontWeight` as CSS). */
export interface ChartTextLook {
  fontFamily: string;
  fontSize: number;
  fontWeight: string;
  fontStyle: string;
  color: string;
  /** Character Spacing (px after every character; negative: condensed). */
  letterSpacing?: number;
  /** All Caps or Small Caps, the words kept as typed (`chart-letter-spacing`). */
  caps?: "all" | "small";
}

/** A chart text the editor draws (`textLayer`), in the picture's own px. */
export interface ChartTextFrame {
  /** The centre of the words. */
  cx: number;
  cy: number;
  /** The words' extent along the text (`width`) and across it (`height`). */
  width: number;
  height: number;
  /** Degrees clockwise: -90 reads bottom-to-top (a vertical axis title);
   * a title's Custom angle any other (C08_axis_title_rotated's 20). */
  rotation: number;
  /** The words as the chart's model spells them; `\n` starts a new line. */
  text: string;
  /** `text` in pieces, each with its look; `run` is the model run (a
   * title's `titleRuns`, 0 for a text of one format) the piece belongs to. */
  runs: { text: string; run: number; look: ChartTextLook }[];
  /** Drawn in capitals (an imported style's All Caps); the words keep their case. */
  caps?: boolean;
  /** The width (px) the drawing wraps the words at, at their spaces (a level
   * vertical-axis title): soft breaks, never a `\n` in the words. */
  wrapWidth?: number;
  /** Text direction Stacked: drawn level, a character under the one before
   * (`writing-mode: vertical-lr` upright); `text` stays the model's words. */
  stacked?: boolean;
  /** WordArt drawn over every run (`officeTextPaintCss`). */
  paint?: Record<string, unknown>;
}

export type ChartTextPart = "title" | "axisX" | "axisY";

/** A part of a drawn chart, in the picture's own pixels (origin top-left). */
export interface ChartPartBox {
  x: number;
  y: number;
  width: number;
  height: number;
  /** The text's look where the part is a text (title, axis title). */
  fontSize?: number;
  fontWeight?: string;
  fontFamily?: string;
  color?: string;
  /** A vertical axis title reads bottom-to-top. */
  vertical?: boolean;
}

export type ChartPartName =
  | "title"
  | "axisX"
  | "axisY"
  | "legend"
  | "plot"
  /** The labels along the horizontal / vertical axis. */
  | "axisLabelX"
  | "axisLabelY"
  /** The series' data labels, together. */
  | "dataLabels";

/** A series as the picture draws it (`ecmeta_series_index` is `index`). */
export interface ChartRenderSeries {
  index: number;
  /** The name the plan gives it (a repeated name carries the planner's mark). */
  name: string;
  /** ECharts' series type (bar, line, pie, custom, ...). */
  type: string;
  /** The series' own paint, when it states one; else `palette[index]`. */
  color?: string;
}

/** A deck chart's data table as the slide draws it under the plot (CSS px). */
export interface ChartRenderDataTable {
  categories: string[];
  /** Each row: its series' name, values and legend-key colour, and its
   * key's shape when it is not the filled square (a line series' line and
   * marker, as the legend keys it). */
  series: { name: string; values: string[]; color: string; key?: { line: true; marker?: string } }[];
  legendKeys: boolean;
  borders: { horizontal: boolean; vertical: boolean; outline: boolean };
  /** The band's height and one row's, and the text's size and colour. */
  heightPx: number;
  rowPx: number;
  fontSize: number;
  color: string;
  /** A workbook chart's band inside its picture: its top (px from the
   * picture's top), and the plot's span its category columns stand under
   * (the series names left of it, X-068). Absent: the band is under the
   * picture (a deck's). */
  topPx?: number;
  plotLeftPx?: number;
  plotWidthPx?: number;
  /** Format Data Table > Fill (behind the values and category names only)
   * and the rules' colour; the text's face; values centred (Excel). */
  fill?: string;
  rule?: string;
  fontFamily?: string;
  valueAlign?: "center" | "right";
}

export interface ChartRenderPicture {
  ok: true;
  svg: string;
  width: number;
  height: number;
  renderMode: "native" | "projection";
  /** A 3-D chart that needs WebGL is drawn as its flat projection. */
  flattened?: boolean;
  /** The plan's notice for the host ("threeDimensionalProjection"). */
  notice?: string;
  /** The colours the plan paints its series / points with, in order. */
  palette: string[];
  series: ChartRenderSeries[];
  /** Where the parts are drawn (a part that is not drawn is absent). */
  parts: Partial<Record<ChartPartName, ChartPartBox>>;
  /** A deck chart's native 3-D pie: the scene the browser rasters over the
   * picture (`renderNativePie3d` + `composeNativePie3dSvg`). */
  pie3d?: Record<string, unknown>;
  /** Asked for with `tips`: what a tooltip says for each drawn point. */
  tips?: ChartPointTips;
  /** Asked for with `dataTableBand`: the data table under the picture. */
  dataTable?: ChartRenderDataTable;
  /**
   * A worksheet chart's frame: the line the sheet's canvas draws round the
   * picture (the chart area's outline, else the style's hairline), which a
   * print draws too. Absent when no line is drawn.
   */
  frame?: {
    color: string;
    widthPt: number;
    style: "solid" | "dotted" | "dashed" | "double";
    radiusPx?: number;
  };
  /** Asked for with `textLayer`: the texts the editor draws itself. */
  texts?: Partial<Record<ChartTextPart, ChartTextFrame>>;
  /** Asked for with `textLayer`: each legend entry's box, by its index in the
   * legend's whole list (a deleted entry has none). */
  legendEntries?: { index: number; box: ChartPartBox }[];
  /** Asked for with `textLayer`: each data label's box and words, by the
   * chart's own series and point; its shapes carry `ecmeta_label` = `mark`
   * (the drawing's series and point). */
  labels?: { series: number; point: number; mark: string; text: string; box: ChartPartBox }[];
}

/**
 * Per drawn series (`ecmeta_series_index`), per drawn point
 * (`ecmeta_data_index`): the point's category (or X value, or slice name) and
 * its value as the chart's own data labels write it. The browser lays these
 * out; it never formats a value itself. A blank point is null.
 */
export type ChartPointTips = {
  name: string;
  points: ({ name: string; value: string } | null)[];
}[];

export type ChartRenderResult = ChartRenderPicture | { ok: false; code: string; detail?: string };

/** At most this many charts per `render` call (the exports' batch). */
export const CHARTS_ENGINE_MAX_CHARTS = 32;
/** At most this many tiles per `gallery` call (every style tile of one open). */
export const CHARTS_ENGINE_MAX_TILES = 64;
/** The style gallery's tile, in px (Word's own thumbnail proportions). */
export const CHART_GALLERY_TILE = { width: 576, height: 328 } as const;

export interface ChartGalleryTile {
  /** The tile's V2 attrs (a style applied to the chart). */
  attrs: Record<string, unknown>;
  /** Index into `looks` of the imported look the tile draws with. */
  look?: number;
  /** Index into `accents` of the theme accents a Change Colors pick resolves against. */
  accents?: number;
  /**
   * The size the chart is drawn at on its page (px): the tile is that chart,
   * laid out at that size and shown smaller, as Office's Chart Styles
   * gallery previews the chart itself -- drawn at the tile's own 576 x 328
   * its words stood a third larger than the chart's and crowded it.
   */
  width?: number;
  height?: number;
}

export interface ChartGalleryRequest {
  tiles: ChartGalleryTile[];
  /** Shared imported looks (one chart's tiles share one). */
  looks?: Record<string, unknown>[];
  /** Shared accent sets. */
  accents?: string[][];
  width?: number;
  height?: number;
  locale?: string;
}

/** A tile: the chart's picture, or the style's own colours for a chart no
 * flat picture exists of (a 3-D chart, a chart the planner refuses). */
export type ChartGalleryTileResult =
  /** The chart's picture and the size it was drawn at (its `viewBox`). */
  | { kind: "image"; svg: string; width?: number; height?: number }
  | { kind: "swatch"; background: string; colors: string[] };

/**
 * A chart that needs WebGL (3-D columns, bars, lines, surfaces): the plan's
 * ECharts option with no functions in it, for the browser's echarts-gl to
 * draw, and each series' data labels evaluated per point on the server.
 */
export interface ChartGlPlan {
  option: Record<string, unknown>;
  /** `labels[series][point]`: the label text the plan's formatter gives. */
  labels: (string | null)[][];
  /** Each point's tooltip words (as `ChartRenderPicture.tips`). */
  tips?: ChartPointTips;
  renderMode: "native" | "projection";
  notice?: string;
  /** Option paths of functions that were left out (diagnostics). */
  dropped?: string[];
}

export interface ChartsEngineMethods {
  render: {
    input: { charts: ChartRenderRequest[] };
    output: { results: ChartRenderResult[] };
  };
  gallery: {
    input: ChartGalleryRequest;
    output: { tiles: (ChartGalleryTileResult | null)[] };
  };
  /** The WebGL plan of one chart (null when the chart draws without WebGL). */
  gl: {
    input: { chart: Omit<ChartRenderRequest, "width" | "height"> & { width?: number; height?: number } };
    output: { gl: ChartGlPlan | null };
  };
}
