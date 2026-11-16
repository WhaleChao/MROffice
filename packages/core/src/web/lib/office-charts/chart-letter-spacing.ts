/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function officeSmallCapsSizePx(sizePx: number): number {
  return (Math.round((sizePx * 3) / 4 * 0.8) * 4) / 3;
}

export type ChartTextCaps = "all" | "small";

/** A size (px) rounded to whole points, half up: the size PowerPoint draws a run's glyphs at. */
export function officeWholePointPx(sizePx: number): number {
  return (Math.floor((sizePx * 3) / 4 + 0.5) * 4) / 3;
}

/** Whether a size (px) is no whole number of points (10.64 pt, 11.97 pt). */
export function chartSizeIsFractional(sizePx: number): boolean {
  const pt = (sizePx * 3) / 4;
  return sizePx > 0 && Math.abs(pt - Math.floor(pt + 0.5)) >= 0.01;
}

/**
 * A run of a fractional size drawn at whole points where its own size lays
 * it out, in a medium that has no `textLength` (the editor's HTML words): the
 * whole size, and the spacing (px) each character takes on top of its own so
 * the run is as wide as at its size -- the difference of the run's widths at
 * the two sizes, shared by its characters. Null for a whole size. `measure`
 * measures plain text in a font (`style weight sizepx family`); Small Caps'
 * smaller capitals are of the stated size and stand as they are.
 */
export function chartWholePointSpacing(
  words: string,
  look: { fontSize: number; fontFamily: string; fontWeight?: string; fontStyle?: string; caps?: ChartTextCaps },
  measure: (text: string, font: string) => number,
): { fontSize: number; spacing: number } | null {
  // (Lines' breaks take no room and no spacing.)
  const text = words.replace(/\n/g, "");
  if (!chartSizeIsFractional(look.fontSize) || !text) return null;
  const whole = officeWholePointPx(look.fontSize);
  const font = (size: number) => `${look.fontStyle ?? "normal"} ${look.fontWeight ?? "normal"} ${size}px ${look.fontFamily}`;
  const shown = look.caps === "all" ? chartTextAllCaps(text) : text;
  const full = look.caps === "small"
    ? chartTextSmallCapsPieces(text).filter((piece) => !piece.small).map((piece) => piece.text).join("")
    : shown;
  const count = [...text].length;
  if (!full || !count) return { fontSize: whole, spacing: 0 };
  return { fontSize: whole, spacing: (measure(full, font(look.fontSize)) - measure(full, font(whole))) / count };
}

export interface ChartTextMark {
  /** Px after every character (negative: condensed). */
  spacing: number;
  caps?: ChartTextCaps;
  /** Its glyphs drawn at its size rounded to whole points, where its own size lays them out (`officeWholePointPx`). */
  whole?: boolean;
}

const MARK = "mrai-chart-text";
const MARK_PATTERN = /,?\s*'mrai-chart-text (-?[\d.]+)(?: (all|small))?( whole)?'/;

/** The style keys `officeChartTextEChartsStyle` states them under. */
export const OFFICE_CHART_TEXT_SPACING = "officeTextSpacing";
export const OFFICE_CHART_TEXT_CAPS = "officeTextCaps";

/** The family stack with the mark at its end (no mark for no spacing and no caps). */
export function chartTextMarkedFamily(family: string | undefined, mark: ChartTextMark): string {
  const base = (family && family.trim()) || "sans-serif";
  const clean = base.replace(MARK_PATTERN, "");
  if (!mark.spacing && !mark.caps && !mark.whole) return clean;
  const spacing = Math.round(mark.spacing * 1000) / 1000;
  return `${clean}, '${MARK} ${spacing}${mark.caps ? ` ${mark.caps}` : ""}${mark.whole ? " whole" : ""}'`;
}

/** The mark a font or family string carries, and the string without it. */
export function readChartTextMark(font: string): { mark: ChartTextMark | null; font: string } {
  const match = MARK_PATTERN.exec(font);
  if (!match) return { mark: null, font };
  const spacing = Number(match[1]);
  return {
    mark: {
      spacing: Number.isFinite(spacing) ? spacing : 0,
      ...(match[2] ? { caps: match[2] as ChartTextCaps } : {}),
      ...(match[3] ? { whole: true } : {}),
    },
    font: font.replace(MARK_PATTERN, ""),
  };
}

/** The words in capitals (All Caps). */
export function chartTextAllCaps(text: string): string {
  return text.toLocaleUpperCase();
}

/** The words in Small Caps: pieces, `small` ones drawn as capitals at `officeSmallCapsSizePx`. */
export function chartTextSmallCapsPieces(text: string): { text: string; small: boolean }[] {
  const pieces: { text: string; small: boolean }[] = [];
  for (const ch of text) {
    const upper = ch.toLocaleUpperCase();
    const small = upper !== ch && ch.toLocaleLowerCase() === ch;
    const last = pieces[pieces.length - 1];
    if (last && last.small === small) last.text += small ? upper : ch;
    else pieces.push({ text: small ? upper : ch, small });
  }
  return pieces;
}

/**
 * A marked text's width: its capitals measured as drawn, and the spacing
 * after each of its characters. `measure` measures plain text in `font`.
 */
export function chartMarkedTextWidth(
  text: string,
  font: string,
  measure: (text: string, font: string) => number,
): number {
  const { mark, font: plain } = readChartTextMark(font);
  if (!mark) return measure(text, font);
  // The engine's probe of a line's height (zrender takes "国"'s width for
  
  // years stood 1 pt lower with 1 pt spacing counted in their line).
  if (text === "国") return measure(text, plain);
  const count = [...text].length;
  let width: number;
  if (mark.caps === "all") width = measure(chartTextAllCaps(text), plain);
  else if (mark.caps === "small") {
    const size = Number(/(\d+(?:\.\d+)?)px/.exec(plain)?.[1]) || 12;
    const smaller = plain.replace(/(\d+(?:\.\d+)?)px/, `${officeSmallCapsSizePx(size)}px`);
    width = chartTextSmallCapsPieces(text).reduce((sum, piece) => sum + measure(piece.text, piece.small ? smaller : plain), 0);
  } else width = measure(text, plain);
  return width + mark.spacing * count;
}

const unescapeXml = (value: string) =>
  value.replace(/&(amp|lt|gt|quot|apos|#39|#x27);/g, (_, name: string) =>
    ({ amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", "#39": "'", "#x27": "'" })[name] ?? "",
  );
const escapeXml = (value: string) =>
  value.replace(/[&<>]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[ch]!);

/**
 * The SVG with every marked text drawn as its mark says: the mark out of its
 * face, `letter-spacing` on it, its words in capitals (Small Caps: the
 * lower-case letters as smaller capitals).
 */
export function chartSvgTextMarksDrawn(svg: string): string {
  if (!svg.includes(MARK)) return svg;
  return svg.replace(/<text\b([^>]*)>([^<]*)<\/text>/g, (whole, attrs: string, content: string) => {
    if (!attrs.includes(MARK)) return whole;
    const { mark, font: cleaned } = readChartTextMark(attrs);
    if (!mark) return whole;
    let next = cleaned;
    if (mark.spacing) {
      const spacing = `letter-spacing:${Math.round(mark.spacing * 1000) / 1000}px;`;
      next = /\bstyle="/.test(next) ? next.replace(/\bstyle="/, `style="${spacing}`) : `${next} style="${spacing}"`;
    }
    const words = unescapeXml(content);
    let body = escapeXml(words);
    if (mark.caps === "all") body = escapeXml(chartTextAllCaps(words));
    else if (mark.caps === "small") {
      const size = Number(/font-size:\s*([\d.]+)px/.exec(next)?.[1] ?? /\b(\d+(?:\.\d+)?)px\b/.exec(next)?.[1]) || 12;
      body = chartTextSmallCapsPieces(words)
        .map((piece) =>
          piece.small
            ? `<tspan style="font-size:${Math.round(officeSmallCapsSizePx(size) * 1000) / 1000}px">${escapeXml(piece.text)}</tspan>`
            : escapeXml(piece.text),
        )
        .join("");
    }
    return `<text${next}>${body}</text>`;
  });
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);

/**
 * The option with every text style that states a spacing or capitals
 * (`OFFICE_CHART_TEXT_SPACING`, `OFFICE_CHART_TEXT_CAPS`) carrying them in its
 * face (`chartTextMarkedFamily`): its own face, else the nearest one it
 * inherits (the style it sits in, the chart's `textStyle`). Plain data and
 * functions pass through; a style already marked is marked afresh.
 */
export function withChartTextMarks<T>(option: T, options: { wholePoints?: boolean } = {}): T {
  const root = isRecord(option) && isRecord(option.textStyle) && typeof option.textStyle.fontFamily === "string"
    ? option.textStyle.fontFamily
    : undefined;
  const rootSize = isRecord(option) && isRecord(option.textStyle) && typeof option.textStyle.fontSize === "number"
    ? option.textStyle.fontSize
    : undefined;
  let changed = false;
  const visit = (value: unknown, family: string | undefined, depth: number, size?: number): unknown => {
    if (depth > 12) return value;
    if (Array.isArray(value)) {
      let any = false;
      const next = value.map((entry) => {
        const out = visit(entry, family, depth + 1, size);
        if (out !== entry) any = true;
        return out;
      });
      return any ? next : value;
    }
    if (!isRecord(value)) return value;
    const own = typeof value.fontFamily === "string" && value.fontFamily ? value.fontFamily : family;
    const ownSize = typeof value.fontSize === "number" ? value.fontSize : size;
    let next: Record<string, unknown> | null = null;
    for (const [key, entry] of Object.entries(value)) {
      if (key === "data" && Array.isArray(entry) && entry.length > 64 && !entry.some(isRecord)) continue;
      const out = visit(entry, own, depth + 1, ownSize);
      if (out !== entry) (next ??= { ...value })[key] = out;
    }
    const spacing = value[OFFICE_CHART_TEXT_SPACING];
    const caps = value[OFFICE_CHART_TEXT_CAPS];
    // (Whole points: a style stating a fractional size of its own, for the
    // canvas that draws it to round its glyphs.)
    const whole = options.wholePoints === true && typeof value.fontSize === "number" && chartSizeIsFractional(value.fontSize);
    if ((typeof spacing === "number" && spacing !== 0) || caps === "all" || caps === "small" || whole) {
      const marked = chartTextMarkedFamily(own, {
        spacing: typeof spacing === "number" ? spacing : 0,
        ...(caps === "all" || caps === "small" ? { caps } : {}),
        ...(whole ? { whole } : {}),
      });
      if (marked !== value.fontFamily) (next ??= { ...value }).fontFamily = marked;
    }
    if (next) changed = true;
    return next ?? value;
  };
  const out = visit(option, root, 0, rootSize) as T;
  return changed ? out : option;
}

let canvasMarksInstalled = false;

/**
 * A browser's 2-D canvas drawing a marked text (the WebGL chart's labels,
 * drawn by the engine on a canvas): measured with its spacing and capitals,
 * and drawn with them a character at a time -- a fractional size's glyphs at
 * whole points where its own size's advances stand them (`whole`). A text whose font carries no mark is the canvas's
 * own; installed once, in the browser only.
 */
export function installChartTextMarksOnCanvas(): void {
  if (canvasMarksInstalled || typeof CanvasRenderingContext2D === "undefined") return;
  canvasMarksInstalled = true;
  const proto = CanvasRenderingContext2D.prototype;
  const measure = proto.measureText;
  const fill = proto.fillText;
  const stroke = proto.strokeText;
  const plainWidth = (ctx: CanvasRenderingContext2D, text: string, font: string) => {
    const before = ctx.font;
    ctx.font = font;
    const width = measure.call(ctx, text).width;
    ctx.font = before;
    return width;
  };
  proto.measureText = function measureText(this: CanvasRenderingContext2D, text: string): TextMetrics {
    if (!this.font.includes(MARK)) return measure.call(this, text);
    const width = chartMarkedTextWidth(String(text), this.font, (piece, font) => plainWidth(this, piece, readChartTextMark(font).font));
    const metrics = measure.call(this, String(text));
    return new Proxy(metrics, { get: (target, key) => (key === "width" ? width : Reflect.get(target, key)) });
  };
  const draw = (paint: typeof fill) =>
    function drawText(this: CanvasRenderingContext2D, text: string, x: number, y: number, maxWidth?: number) {
      const font = this.font;
      if (!font.includes(MARK)) return maxWidth === undefined ? paint.call(this, text, x, y) : paint.call(this, text, x, y, maxWidth);
      const { mark, font: plain } = readChartTextMark(font);
      const words = String(text);
      const pieces = mark?.caps === "small"
        ? chartTextSmallCapsPieces(words)
        : [{ text: mark?.caps === "all" ? chartTextAllCaps(words) : words, small: false }];
      const size = Number(/(\d+(?:\.\d+)?)px/.exec(plain)?.[1]) || 12;
      const smaller = plain.replace(/(\d+(?:\.\d+)?)px/, `${officeSmallCapsSizePx(size)}px`);
      // A fractional size's glyphs at whole points, each standing where its
      // own size's advance puts it (`officeWholePointPx`).
      const glyphs = chartSizeIsFractional(size) ? plain.replace(/(\d+(?:\.\d+)?)px/, `${officeWholePointPx(size)}px`) : plain;
      const spacing = mark?.spacing ?? 0;
      const total = chartMarkedTextWidth(words, font, (piece, face) => plainWidth(this, piece, readChartTextMark(face).font));
      const align = this.textAlign;
      // From the run's start: the canvas's own alignment would place each piece.
      let at = align === "center" ? x - total / 2 : align === "right" || align === "end" ? x - total : x;
      this.save();
      this.textAlign = "left";
      for (const piece of pieces) {
        for (const ch of piece.text) {
          this.font = piece.small ? smaller : glyphs;
          paint.call(this, ch, at, y);
          this.font = piece.small ? smaller : plain;
          at += measure.call(this, ch).width + spacing;
        }
      }
      this.restore();
    };
  proto.fillText = draw(fill);
  proto.strokeText = draw(stroke);
}
