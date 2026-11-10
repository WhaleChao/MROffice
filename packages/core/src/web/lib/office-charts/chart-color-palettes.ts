/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart Design ▸ Change Colors, once for the three editors.
 *
 * Microsoft Support ("Change the color or style of a chart in Office", Excel,
 * Word and PowerPoint alike): a pick from the Color gallery is a colour
 * scheme made of the document's THEME colours; it does not change the theme.
 * In the file it is the chart's Chart Colors part ([MS-ODRAWXML] 2.1.1), and
 * an Office 2016 chart's series state no paint of their own, so that part is
 * all Office paints one with.
 *
 * The table and the bytes of the part a pick becomes are generated from the
 * server's one copy (`chart-color-palettes.generated.ts`). This module
 * resolves a pick against a host's theme accents the way Office does - each
 * token's `lumMod` / `lumOff` by the one colour rule (`office-color-transform`),
 * and once the palette is used up the cycle again under Excel's next variation
 * - and holds the edit the Docs and deck ribbons commit. The worksheet keeps
 * its object-level `chartColorPalette` and resolves it here too.
 *
 * Office's gallery has four Colorful and thirteen Monochromatic palettes, each
 * with its part's `meth`: a `cycle` walks its colours, and a within-linear
 * palette spreads one accent's shades over however many elements the chart
 * paints, darkest first (`withinLinear`) or lightest first
 * (`withinLinearReversed`) - so a colour is element `index` OF `count`.
 */
import {
  CHART_COLOR_PALETTE_FAMILIES,
  chartSeriesRenderAsOf,
  chartTakesColorPalette,
  isChartColorPaletteId,
  type ChartBlockAttrsV2,
} from "../../components/console/editor/extensions/chartBlockShared";
import { OFFICE_THEME_PALETTES } from "../office-color-palette";
import type {
  ChartNativeLook,
  ChartNativeLookArea,
} from "../office-chart-native-look";
import { officeTransformedHex } from "../office-color-transform";

import {
  CHART_COLOR_PALETTE_IDS,
  CHART_COLOR_PALETTES,
  CHART_COLOR_VARIATIONS,
  CHART_COLOR_WITHIN_LINEAR_STEPS,
  type ChartColorPaletteId,
  type ChartColorPaletteToken,
} from "./chart-color-palettes.generated";

export {
  CHART_COLOR_PALETTE_FAMILIES,
  CHART_COLOR_PALETTE_IDS,
  CHART_COLOR_PALETTES,
  chartTakesColorPalette,
  isChartColorPaletteId,
  type ChartColorPaletteId,
  type ChartColorPaletteToken,
};

/**
 * The accents a host draws a pick against when it knows no theme of its own:
 * Office's own theme (the `office` gallery entry every editor already draws its
 * default chart with).
 */
export const DEFAULT_CHART_THEME_ACCENTS: readonly string[] =
  OFFICE_THEME_PALETTES[0].accents;

/**
 * A document theme's six accents as a chart resolves them: the theme's own,
 * and Office's default theme's for a slot it leaves out (MRAI's deck themes
 * state five). The deck's PPTX export fills its theme part the same way, so a
 * chart's `accentN` reference names one colour on screen and in the file.
 */
export function chartThemeAccents(
  accents: readonly string[] | undefined,
): string[] {
  return DEFAULT_CHART_THEME_ACCENTS.map(
    (fallback, index) => accents?.[index] ?? fallback,
  );
}

/**
 * The theme's text and background slots a palette may name (Monochromatic
 * Palette 7 is Dark 1's grays), as Office's theme has them: no host theme
 * MRAI draws a chart against states other ones.
 */
const THEME_TEXT_SLOTS: Readonly<Record<string, string>> = {
  dk1: "#000000",
  tx1: "#000000",
  lt1: "#ffffff",
  bg1: "#ffffff",
};

function tokenColor(
  token: { scheme: string; lumMod: number | null; lumOff: number | null },
  accents: readonly string[],
): string {
  const text = THEME_TEXT_SLOTS[token.scheme];
  if (text !== undefined) {
    const transforms: [string, number][] = [];
    if (token.lumMod !== null) transforms.push(["lumMod", token.lumMod]);
    if (token.lumOff !== null) transforms.push(["lumOff", token.lumOff]);
    return (officeTransformedHex(text, transforms) ?? text).toLowerCase();
  }
  const slot = /^accent([1-6])$/.exec(token.scheme);
  // A slot the list leaves out is Office's own accent for that slot, as the
  // drawing fills it (`chartThemeAccents`): falling back to the FIRST colour
  // made the Change Colors menu show six blues over an empty or short list
  // while the chart it painted went on to orange (PPT-B336).
  const index = slot ? Number(slot[1]) - 1 : 0;
  const base = accents[index] ?? DEFAULT_CHART_THEME_ACCENTS[index];
  const transforms: [string, number][] = [];
  if (token.lumMod !== null) transforms.push(["lumMod", token.lumMod]);
  if (token.lumOff !== null) transforms.push(["lumOff", token.lumOff]);
  return (officeTransformedHex(base, transforms) ?? base).toLowerCase();
}

/** A luminance step as one signed shift: darker is negative, lighter positive. */
function stepShift(step: { lumMod: number | null; lumOff: number | null }): number {
  if (step.lumOff !== null) return step.lumOff;
  return step.lumMod !== null ? step.lumMod - 100_000 : 0;
}

/**
 * The shift element `position` of `count` takes along one accent's steps,
 * darkest first (lightest first, reversed): the ends are the steps' ends, the
 * elements between sit evenly along them. The twin of the server's
 * `office_chart_color_palette._within_linear_shift`, operation for operation,
 * so both land on the same value.
 */
function withinLinearShift(position: number, count: number, reversed: boolean): number {
  const shifts = CHART_COLOR_WITHIN_LINEAR_STEPS.map(stepShift);
  if (reversed) shifts.reverse();
  // A lone element is the first position: the darkest (lightest, reversed).
  if (count <= 1) return shifts[0];
  const span = shifts.length - 1;
  const numerator = position * span;
  const low = Math.floor(numerator / (count - 1));
  const remainder = numerator % (count - 1);
  const high = Math.min(low + 1, span);
  const exact = shifts[low] + ((shifts[high] - shifts[low]) * remainder) / (count - 1);
  return Math.floor(exact + 0.5);
}

/**
 * The theme token element `index` of `count` takes under a within-linear
 * palette: its one colour, darkened (`lumMod`) or lightened (`lumMod` +
 * `lumOff`) by the element's shift. A chart that does not say how many
 * elements it paints takes six, the steps themselves.
 */
function withinLinearToken(
  paletteId: ChartColorPaletteId,
  index: number,
  count: number | undefined,
): ChartColorPaletteToken {
  const palette = CHART_COLOR_PALETTES[paletteId];
  const total =
    count !== undefined && Number.isInteger(count) && count > 0
      ? count
      : CHART_COLOR_WITHIN_LINEAR_STEPS.length;
  const shift = withinLinearShift(
    index % total,
    total,
    palette.method === "withinLinearReversed",
  );
  const scheme = palette.tokens[0].scheme;
  if (shift < 0) return { scheme, lumMod: 100_000 + shift, lumOff: null };
  if (shift > 0) return { scheme, lumMod: 100_000 - shift, lumOff: shift };
  return { scheme, lumMod: null, lumOff: null };
}

/**
 * The colour a palette gives element `index` of `count` against these
 * accents. A cycle: the token `index mod n`, then the variation
 * `floor(index / n)` of Excel's ladder applied on top ("all contained colors
 * repeated each time with each variation applied", `CT_ColorStyle`). A
 * within-linear palette: its one colour at the element's place among `count`
 * (`withinLinearToken`).
 */
export function chartColorPaletteColor(
  paletteId: ChartColorPaletteId,
  accents: readonly string[],
  index: number,
  count?: number,
): string {
  if (CHART_COLOR_PALETTES[paletteId].method !== "cycle") {
    return tokenColor(withinLinearToken(paletteId, index, count), accents);
  }
  const tokens = CHART_COLOR_PALETTES[paletteId].tokens;
  const base = tokenColor(tokens[index % tokens.length], accents);
  const variation =
    CHART_COLOR_VARIATIONS[
      Math.floor(index / tokens.length) % CHART_COLOR_VARIATIONS.length
    ];
  const transforms: [string, number][] = [];
  if (variation.lumMod !== null) transforms.push(["lumMod", variation.lumMod]);
  if (variation.lumOff !== null) transforms.push(["lumOff", variation.lumOff]);
  return transforms.length
    ? (officeTransformedHex(base, transforms) ?? base).toLowerCase()
    : base;
}

/**
 * The colours Excel's built-in chart styles (c:style 1-48, the Office
 * 2007/2010 look) give `count` series that cycle the theme accents. Fewer
 * series than accents take the accents; from a full round on, every round is
 * shaded or tinted as a whole, darkest first, evenly: with n = floor(count /
 * accents), round `c` moves by f = (c - n/2) * 1.4 / (n + 2) -- a shade of
 * 1 + f below zero, a tint of 1 - f above (DrawingML `a:shade` / `a:tint`).
 * So exactly six series are already one 76.7 % shaded round, and a seventh
 * starts a 76.7 % tinted one. Measured on Excel for Mac prints of 5, 6, 7, 9,
 * 12, 18, 24 and 30 automatic columns and a six-point pie (X-046,
 * C02_auto_marker_excel10: 12 series read 4A7FB0 .. 5B9BD5 ..; X-069: style 2
 * on the Office 2007 theme, 5 series 4F81BD.., 6 series 4572A7.., the 7th
 * 93A9CF).
 */
export function excelBuiltInStyleSeriesColors(
  accents: readonly string[],
  count: number,
): string[] {
  if (!accents.length) return [];
  const rounds = Math.floor(Math.max(0, count) / accents.length);
  if (rounds < 1) return accents.map((accent) => accent.toLowerCase());
  return Array.from({ length: (rounds + 1) * accents.length }, (_, index) => {
    const base = accents[index % accents.length];
    const round = Math.floor(index / accents.length);
    const shift = ((round - rounds / 2) * 1.4) / (rounds + 2);
    const transforms: [string, number][] =
      Math.abs(shift) < 1e-9
        ? []
        : shift < 0
          ? [["shade", Math.round((1 + shift) * 100_000)]]
          : [["tint", Math.round((1 - shift) * 100_000)]];
    return (transforms.length ? officeTransformedHex(base, transforms) ?? base : base).toLowerCase();
  });
}

/** The tints of tx1 Excel's greyscale built-in styles cycle (X-069). */
const EXCEL_GREYSCALE_STYLE_TINTS = [88_000, 55_000, 78_000, 92_000, 70_000, 30_000] as const;

/**
 * The series colours of Excel's built-in chart style `cStyle` (c:style 1-48,
 * the Office 2007/2010 Chart Styles gallery: eight colour columns, six
 * effect rows) over these theme accents, or null for a style whose colours
 * this does not know. Column (cStyle - 1) mod 8: 1 cycles the accents
 * (`excelBuiltInStyleSeriesColors`); 2..7 is one accent (accent 1..6) spread
 * over the series, darkest first -- series i of n moves by
 * f = (i + 1) / (n + 1) * 1.4 - 0.7, a shade below zero and a tint above, so
 * two series read 76.7 % shade / 76.7 % tint (measured on Excel for Mac,
 * C24_tdf119138's styles 11, 12, 13: 4572A7 / 93A9CF, AA4643 / D19392,
 * 89A54E / B9CD96; X-071). Column 0 (greyscale, styles 1, 9, .. 41) cycles
 * six tints of the text colour (tx1) the same way as column 1 cycles the
 * accents: 88 %, 55 %, 78 %, 92 %, 70 % and 30 % -- on black 616161 B3B3B3
 * 818181 505050 959595 for up to five series, 555555 9E9E9E 727272 464646
 * 838383 C1C1C1 for six, then 9A9A9A C8C8C8 AAAAAA (Excel for Mac, X-069).
 */
export function excelBuiltInChartStyleColors(
  cStyle: number,
  accents: readonly string[],
  count: number,
  text: string = "#000000",
): string[] | null {
  if (!Number.isInteger(cStyle) || cStyle < 1 || cStyle > 48 || accents.length < 6) return null;
  const column = (cStyle - 1) % 8;
  if (column === 1) return excelBuiltInStyleSeriesColors(accents, count);
  if (column === 0) {
    const greys = EXCEL_GREYSCALE_STYLE_TINTS.map(
      (tint) => officeTransformedHex(text, [["tint", tint]]) ?? text,
    );
    return excelBuiltInStyleSeriesColors(greys, count);
  }
  const base = accents[column - 2];
  const n = Math.max(1, count);
  return Array.from({ length: n }, (_, index) => {
    const shift = ((index + 1) / (n + 1)) * 1.4 - 0.7;
    const transforms: [string, number][] =
      Math.abs(shift) < 1e-9
        ? []
        : shift < 0
          ? [["shade", Math.round((1 + shift) * 100_000)]]
          : [["tint", Math.round((1 - shift) * 100_000)]];
    return (transforms.length ? officeTransformedHex(base, transforms) ?? base : base).toLowerCase();
  });
}

/**
 * The colours a palette gives `count` elements, against these accents (six,
 * the gallery's row, when the caller does not say).
 */
export function chartColorPaletteColors(
  paletteId: ChartColorPaletteId,
  accents: readonly string[] = DEFAULT_CHART_THEME_ACCENTS,
  count: number = 6,
): string[] {
  return Array.from({ length: Math.max(0, count) }, (_, index) =>
    chartColorPaletteColor(paletteId, accents, index, count),
  );
}

/**
 * How many things a classic chart's Change Colors paints: a pie's points,
 * else its series (0 for a chart with neither).
 */
export function chartColorTargetCount(attrs: ChartBlockAttrsV2): number {
  if (attrs.chartType === "pie" && attrs.data.kind === "category") {
    return attrs.data.categories.length;
  }
  if (
    attrs.data.kind === "category" ||
    attrs.data.kind === "xy" ||
    attrs.data.kind === "stock"
  ) {
    return attrs.data.series.length;
  }
  return 0;
}

/**
 * Whether series `seriesIndex` of this chart is drawn as a line: a 2-D line,
 * a scatter that is not a bubble, a radar that draws lines, a combination's
 * line series. Its colour is its LINE (`outline`, its `c:spPr/a:ln`) - and its
 * markers' - so Change Colors repaints that, where a bar's, an area's or a
 * slice's colour is its fill and its outline is the chart style's border.
 */
export function chartSeriesColorIsItsLine(
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "chartSubtype" | "data">,
  seriesIndex: number,
): boolean {
  const drawn = chartSeriesRenderAsOf(attrs, seriesIndex);
  if (drawn === "line") return attrs.chartSubtype !== "line3d";
  if (drawn === "xy") return attrs.chartSubtype !== "bubble" && attrs.chartSubtype !== "bubble3d";
  if (drawn === "radar") return attrs.chartSubtype !== "filled";
  return false;
}

/**
 * A line series' style with the palette's colour on its line: the line
 * colour a chart style or an earlier pick wrote (and a line gradient, which
 * is a colour too) gives way to `color`, or to the palette's own colour when
 * `color` is undefined (the drawing paints it). A line the author turned off
 * (`outline: "none"`, a markers-only scatter) stays off; the width, dash and
 * shadow, which are not colour, stay.
 */
function withLineTakingPalette(
  entry: Record<string, unknown>,
  color: string | undefined,
): Record<string, unknown> {
  const next = { ...entry };
  if (next.outline !== "none") {
    if (color === undefined) delete next.outline;
    else next.outline = color;
  }
  const paint = next.shapePaint;
  if (paint && typeof paint === "object" && !Array.isArray(paint)) {
    const { lineGradient: _line, ...rest } = paint as Record<string, unknown>;
    if (Object.keys(rest).length) next.shapePaint = rest;
    else delete next.shapePaint;
  }
  return next;
}

/**
 * The chart after a pick, painted as literal fills: each series' (a pie's
 * each point's) `fill` is the palette's colour against these accents, and the
 * outline and effect each already carries stay - but a line series' line,
 * which IS its colour, takes the palette's colour too
 * (`chartSeriesColorIsItsLine`). For a chart whose writer in
 * the file paints series rather than a palette (a classic chart that is
 * already a part of an opened Word document). Null when nothing changes or
 * the chart has nothing to paint.
 */
export function chartWithResolvedColorPalette(
  attrs: ChartBlockAttrsV2,
  paletteId: ChartColorPaletteId,
  accents: readonly string[] = DEFAULT_CHART_THEME_ACCENTS,
): ChartBlockAttrsV2 | null {
  const count = chartColorTargetCount(attrs);
  if (count <= 0 || !chartTakesColorPalette(attrs.chartType)) return null;
  const colors = chartColorPaletteColors(paletteId, accents, count);
  const format: NonNullable<ChartBlockAttrsV2["format"]> = {
    ...(attrs.format ?? {}),
  };
  delete format.colorPalette;
  const points = attrs.chartType === "pie";
  const key = points ? "sliceStyles" : "seriesStyles";
  const index = points ? "pointIndex" : "seriesIndex";
  const existing = (format[key] ?? []) as readonly Record<string, unknown>[];
  (format as Record<string, unknown>)[key] = colors.map((fill, at) => {
    const painted = {
      ...(existing.find((entry) => entry[index] === at) ?? {}),
      [index]: at,
      fill,
    };
    return !points && chartSeriesColorIsItsLine(attrs, at)
      ? withLineTakingPalette(painted, fill)
      : painted;
  });
  if (JSON.stringify(format) === JSON.stringify(attrs.format ?? {})) return null;
  return { ...attrs, format };
}

/** The palette whose colours a chart's literal fills are, if they are one. */
export function chartResolvedColorPaletteId(
  attrs: ChartBlockAttrsV2 | null | undefined,
  accents: readonly string[] = DEFAULT_CHART_THEME_ACCENTS,
): ChartColorPaletteId | null {
  if (!attrs) return null;
  const count = chartColorTargetCount(attrs);
  if (count <= 0) return null;
  const points = attrs.chartType === "pie";
  const entries = (points ? attrs.format?.sliceStyles : attrs.format?.seriesStyles) ?? [];
  const fills = Array.from({ length: count }, (_, at) =>
    (entries as readonly Record<string, unknown>[])
      .find((entry) => entry[points ? "pointIndex" : "seriesIndex"] === at)
      ?.fill,
  );
  return (
    CHART_COLOR_PALETTE_IDS.find((id) => {
      const colors = chartColorPaletteColors(id, accents, count);
      return fills.every(
        (fill, at) => typeof fill === "string" && fill.toLowerCase() === colors[at],
      );
    }) ?? null
  );
}

/** The palette a chart states, or null. */
export function chartColorPaletteId(
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "format"> | null | undefined,
): ChartColorPaletteId | null {
  const value = attrs?.format?.colorPalette;
  return isChartColorPaletteId(value) ? value : null;
}

/**
 * The chart after a Change Colors pick. Office's pick repaints every series
 * from the palette, so an explicit series FILL goes with it, gradient and
 * pattern fills included; the series' outline and effect, which the palette
 * does not name, stay - except on a series drawn as a line, whose line IS its
 * colour: a line colour a chart style wrote (white, under a style that rims
 * its bars) went on drawing every line in it while only the legend took the
 * pick. A fill or line colour set afterwards (Format ▸ Fill / ▸ Line) wins
 * over the palette again for its series.
 * Null when this chart takes no palette or already shows this one.
 */
export function chartWithColorPalette(
  attrs: ChartBlockAttrsV2,
  paletteId: ChartColorPaletteId,
): ChartBlockAttrsV2 | null {
  if (!chartTakesColorPalette(attrs.chartType)) return null;
  const format: NonNullable<ChartBlockAttrsV2["format"]> = {
    ...(attrs.format ?? {}),
    colorPalette: paletteId,
  };
  for (const key of ["seriesStyles", "sliceStyles"] as const) {
    const entries = format[key] as
      | readonly Record<string, unknown>[]
      | undefined;
    if (!entries) continue;
    const kept = entries
      .map((entry) => {
        const { fill: _fill, ...unfilled } = entry;
        const rest =
          key === "seriesStyles" &&
          typeof entry.seriesIndex === "number" &&
          chartSeriesColorIsItsLine(attrs, entry.seriesIndex)
            ? withLineTakingPalette(unfilled, undefined)
            : unfilled;
        // The fill goes whole: a gradient or pattern fill hangs off it
        // (`fill` is the gradient's first stop, the pattern's foreground),
        // so left behind it made a record no chart takes and the pick was
        // refused (PPT-B335). The line's weight, dash, gradient and the
        // shadow, which the palette does not name, stay.
        const paint = rest.shapePaint;
        if (paint && typeof paint === "object" && !Array.isArray(paint)) {
          // Its transparency too: without the fill it hangs off, the model takes no
          
          const { fillGradient: _gradient, fillPattern: _pattern, fillAlpha: _alpha, ...line } = paint as Record<string, unknown>;
          if (Object.keys(line).length) rest.shapePaint = line;
          else delete rest.shapePaint;
        }
        return rest;
      })
      .filter((entry) =>
        Object.keys(entry).some(
          (name) => name !== "seriesIndex" && name !== "pointIndex",
        ),
      );
    if (kept.length) (format as Record<string, unknown>)[key] = kept;
    else delete format[key];
  }
  if (JSON.stringify(format) === JSON.stringify(attrs.format ?? {})) return null;
  return { ...attrs, format };
}

/**
 * The look a chart is drawn with once its Change Colors pick is applied.
 *
 * Office resolves an Office 2016 chart's data-point paint from its Chart
 * Colors part (the style part's `phClr`), so a pick recolours every point the
 * style paints and nothing else: a point the style fills takes the palette's
 * colour in its place (a gradient or pattern of `phClr` becomes that flat
 * colour), a point the style only outlines takes it as its outline, and the
 * outline width, dash and every non-point entry stay as the style said. The
 * colours are the palette's cycle against the look's own theme accents, or
 * Office's when the chart carries none.
 *
 * Unchanged when the chart takes no palette, states none, or states the one
 * its colour part already spells (`look.colorPalette`); a chart with no look
 * of its own gets one holding only the palette's points, which each family
 * maps onto its elements the way it maps an imported chart's.
 */
export function chartLookWithColorPalette(
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "format"> &
    Partial<Pick<ChartBlockAttrsV2, "data">>,
  look: ChartNativeLook | undefined,
  documentAccents?: readonly string[],
): ChartNativeLook | undefined {
  const paletteId = chartTakesColorPalette(attrs.chartType)
    ? chartColorPaletteId(attrs)
    : null;
  if (paletteId === null || look?.colorPalette === paletteId) return look;
  // The chart's own theme (an imported part's), else the document's it sits
  // in, else Office's default theme.
  const accents =
    look?.themeAccents?.length === 6
      ? look.themeAccents
      : documentAccents?.length === 6
        ? documentAccents
        : DEFAULT_CHART_THEME_ACCENTS;
  const stated = look?.points ?? [];
  const count = Math.max(stated.length, 12);
  // How many things the pick paints, which a within-linear palette spreads
  // its shades over: the chart's series (a pie's points) when it says.
  const painted = attrs.data
    ? chartColorTargetCount(attrs as ChartBlockAttrsV2)
    : 0;
  const points = Array.from({ length: count }, (_, index): ChartNativeLookArea => {
    const color = chartColorPaletteColor(
      paletteId,
      accents,
      index,
      painted > 0 ? painted : undefined,
    );
    const point = stated[index];
    if (!point) return { fill: color };
    const { fill, fillGradient, radialGradient, pattern, ...rest } = point;
    const filled =
      (fill !== undefined && fill !== "none") ||
      fillGradient !== undefined ||
      radialGradient !== undefined ||
      pattern !== undefined;
    if (filled) return { ...rest, fill: color };
    return rest.outline !== undefined && rest.outline !== "none"
      ? { ...rest, ...(fill === undefined ? {} : { fill }), outline: color }
      : { ...rest, fill: color };
  });
  return { ...(look ?? {}), points, colorPalette: paletteId };
}
