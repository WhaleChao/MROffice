/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeNormalLineHeightPx } from "@mrnote/core/web/lib/office-fonts/substitutes.ts";
import type { FontLoadReach } from "@mrnote/core/web/lib/spreadsheet/font-load-reach.ts";
import { fontMeasureContext } from "@mrnote/core/web/lib/spreadsheet/font-measure-context.ts";

/**
 * The grid's canvas text layer (on by default; `?gridCanvas=off`): the text of a cell at rest
 * is painted on canvas tiles instead of a display layer of spans in its cell.
 * The cell keeps its element, its fill, its borders and its resting editor
 * (focus, accessibility, hit testing); only the text moves to the canvas.
 *
 * A cell is painted here only when its text is one plain line the display
 * layer would draw unclipped inside the cell: everything else (wrap, rich
 * runs, rotation, fills, scripts, accounting rules, shrink, data
 * bars, icons, sparklines, pictures, controls, the active or edited cell)
 * keeps the DOM display layer. The layout below repeats the display layer's
 * own box model (spreadsheet-window__cell-display-layer: absolute over the
 * cell's padding box, a flex row with the cell's alignment and insets, one
 * line box of the editor's line height) so the text lands where the DOM
 * draws it. A plain line that overflows into empty neighbours is painted
 * too, in the spill layer's box (its `reach`).
 */

export type GridCanvasHorizontal = "start" | "center" | "end";
export type GridCanvasVertical = "top" | "middle" | "bottom";

/** One cell's text as the display layer would draw it. */
export type GridCanvasText = {
  /** The text runs and the letter spacing each is drawn with (native typography). */
  readonly runs: ReadonlyArray<{ readonly text: string; readonly spacing: number }>;
  /**
   * A number format's layout (SpreadsheetFormattedText) instead of `runs`:
   * pieces placed from the line's start, or, after a blank fill that takes
   * the room left, back from the end of the cell's text area.
   */
  readonly segments?: ReadonlyArray<GridCanvasSegment>;
  /** CSS font shorthand the text is laid out in. */
  readonly font: string;
  /**
   * The font its glyphs are painted in, when not `font`: a faux bold whose
   * display layer turns weight synthesis off draws the regular face (with
   * `stroke`) where `font` names bold.
   */
  readonly paintFont?: string;
  /** Faux bold: the outline drawn round each glyph, px (the layer's -webkit-text-stroke). */
  readonly stroke?: number;
  /** The display layer's own font (its line's strut), when it differs from `font`. */
  readonly strutFont?: string;
  readonly kerning: "none" | "normal";
  readonly color: string;
  readonly horizontal: GridCanvasHorizontal;
  readonly vertical: GridCanvasVertical;
  readonly direction: "ltr" | "rtl";
  /** The display layer's insets, px. */
  readonly padTop: number;
  readonly padRight: number;
  readonly padBottom: number;
  readonly padLeft: number;
  /** The line box height, px, or the font's own (`normal`). */
  readonly lineHeight: number | "normal";
  /** The width the text takes, px (measured with the same advances). */
  readonly width: number;
  /**
   * Text that overflows into empty neighbours (Excel's spill): how far its
   * box reaches past the cell's on each side, px, as the spill layer pulls
   * its edges out of the cell.
   */
  readonly reach?: { readonly left: number; readonly right: number };
  /**
   * Wrapped text (the multiline layer's native paragraph) instead of `runs`:
   * each line's runs and width, and its line box in the paragraph (`top`,
   * `height`, px) with its baseline from the paragraph's top, or null where
   * the line's font puts it (half-leading). `width` is then the widest line.
   */
  readonly lines?: ReadonlyArray<GridCanvasLine>;
  /** The paragraph's height, px (with `lines`). */
  readonly paragraphHeight?: number;
  /** `safe` vertical alignment (with `lines`): a paragraph taller than the
   * cell starts at its top, as the layer's `safe center` / `safe flex-end`. */
  readonly safeVertical?: boolean;
  /**
   * Angled text in the cell's own box (no band; angledCellTextBox "rect"):
   * the line box's top-left corner before turning, in the cell's padding box
   * (px), the turn about that corner (degrees, + clockwise) and the line
   * box's height. The text is clipped at the cell.
   */
  readonly turn?: {
    readonly x: number; readonly y: number; readonly rotate: number; readonly lineHeight: number;
    /** A band cell's band (its corners, in the same frame): the text is cut at it instead of at the cell. */
    readonly clip?: ReadonlyArray<readonly [number, number]>;
  };
};

/** One line of a wrapped cell's text (GridCanvasText `lines`). */
export type GridCanvasLine = {
  readonly runs: GridCanvasText["runs"];
  readonly width: number;
  readonly top: number;
  readonly height: number;
  readonly baseline: number | null;
};

/** One piece of a laid-out line: `offset` from the line's start, or (anchor `end`) from the text area's end back to the piece's start. */
export type GridCanvasSegment = {
  readonly runs: GridCanvasText["runs"];
  readonly anchor: "start" | "end";
  readonly offset: number;
};

let flag: boolean | undefined;



export function gridCanvasTextEnabled(): boolean {
  if (flag !== undefined) return flag;
  if (typeof window === "undefined") return false;
  try {
    // The tab keeps the choice: the workspace rewrites its address once a
    // page opens, and the flag is read when the first grid renders.
    const asked = new URLSearchParams(window.location.search).get("gridCanvas");
    if (asked === "on" || asked === "off") window.sessionStorage.setItem("mrai.gridCanvas", asked);
    flag = (asked ?? window.sessionStorage.getItem("mrai.gridCanvas")) !== "off";
  } catch {
    flag = true;
  }
  return flag;
}

/** For tests: set or clear the flag. */
export function setGridCanvasTextEnabledForTests(value: boolean | undefined) {
  flag = value;
}


const metricCache = new Map<string, { ascent: number; descent: number }>();
const normalLineCache = new Map<string, number>();
const widthCache = new Map<string, number>();
/** One run's advance in one font (`font\0kerning\0text`): a cell's runs are
 * often its single characters (each digit spaced to the device grid), and the
 * width of a whole cell text is kept only for that text, so every new cell
 * measured its "." and digits again (R116-X02's open: "." 206 times, 4,100
 * measures of some 450 strings, 1.3-1.8 s). */
const runWidthCache = new Map<string, number>();

/** The measures a font load can change, dropped: a font's metrics and line
 * when it can become the line's first available font, a width when it can
 * draw one of its texts. True when any went. */
export function dropGridCanvasTextMeasuresTouched(reach: FontLoadReach): boolean {
  let dropped = false;
  for (const cache of [metricCache, normalLineCache] as Array<Map<string, unknown>>) {
    for (const font of [...cache.keys()]) {
      if (reach.touchesFont(font, " ")) { cache.delete(font); dropped = true; }
    }
  }
  for (const key of [...widthCache.keys()]) {
    const [font, , runs = ""] = key.split("\u0000");
    const text = runs.split("\u0002").map((run) => run.slice(run.indexOf("\u0001") + 1)).join("");
    if (reach.touchesFont(font, text)) { widthCache.delete(key); dropped = true; }
  }
  for (const key of [...runWidthCache.keys()]) {
    const [font, , text = ""] = key.split("\u0000");
    if (reach.touchesFont(font, text)) { runWidthCache.delete(key); dropped = true; }
  }
  return dropped;
}

/** Clears the measurement caches (fonts finished loading). */
export function clearGridCanvasTextMeasures() {
  metricCache.clear();
  widthCache.clear();
  runWidthCache.clear();
  normalLineCache.clear();
}

/**
 * The font, kerning and letter spacing last set on a tile's context, so a
 * band of cells in one font sets it once: reading `ctx.font` back serialises
 * it, and a family it spells differently never compared equal, so every cell
 * parsed its font again. Forgotten whenever the context's state may have
 * changed behind it (a restore, a resize).
 */
const drawnState = new WeakMap<CanvasRenderingContext2D, { font?: string; kerning?: string; spacing?: string }>();

/** Forget what was set on `ctx`: after `restore()` or a change of its canvas's size, which reset its state. */
export function forgetGridCanvasTextState(ctx: CanvasRenderingContext2D) {
  drawnState.delete(ctx);
}

function applyFont(ctx: CanvasRenderingContext2D, font: string, kerning: "none" | "normal") {
  const state = drawnState.get(ctx);
  if (state) {
    if (state.font !== font) ctx.font = state.font = font;
    if (state.kerning !== kerning && typeof (ctx as { fontKerning?: unknown }).fontKerning === "string") ctx.fontKerning = state.kerning = kerning;
    return;
  }
  if (ctx.font !== font) ctx.font = font;
  if (typeof (ctx as { fontKerning?: unknown }).fontKerning === "string" && ctx.fontKerning !== kerning) ctx.fontKerning = kerning;
}

/** Track what is set on `ctx` from here on (a tile's own context; see forgetGridCanvasTextState). */
export function trackGridCanvasTextState(ctx: CanvasRenderingContext2D) {
  drawnState.set(ctx, {});
}

/** The characters CSS letter-spacing is added after, for a plain run. */
function spacedCharacters(text: string): number {
  return Array.from(text).length;
}

/**
 * The width of `runs` as the display layer draws them: each run's advance
 * plus its letter spacing after every character (CSS adds it after the last
 * one too). Null when nothing can measure (no canvas).
 */
export function gridCanvasTextWidth(runs: GridCanvasText["runs"], font: string, kerning: "none" | "normal"): number | null {
  const key = `${font}\u0000${kerning}\u0000${runs.map((run) => `${run.spacing}\u0001${run.text}`).join("\u0002")}`;
  const known = widthCache.get(key);
  if (known !== undefined) return known;
  // A context kept in this font (fontMeasureContext): cells of different
  // fonts interleave on every row, and setting one shared context's font made
  // each new text resolve the family list again -- about 0.5 ms a text on the
  
  const ctx = fontMeasureContext(font, kerning);
  if (!ctx) return null;
  let width = 0;
  for (const run of runs) {
    const runKey = `${font}\u0000${kerning}\u0000${run.text}`;
    let advance = runWidthCache.get(runKey);
    if (advance === undefined) {
      advance = ctx.measureText(run.text).width;
      if (runWidthCache.size > 50000) runWidthCache.clear();
      runWidthCache.set(runKey, advance);
    }
    width += advance + run.spacing * spacedCharacters(run.text);
  }
  if (widthCache.size > 50000) widthCache.clear();
  widthCache.set(key, width);
  return width;
}

/**
 * The sheet's zoom the text is laid out under (the layer sets it as it
 * paints): a zoomed sheet's text is laid out in the zoomed font, whose
 * ascent and descent layout rounds at that size, not at the font's own.
 * At 130% a line's baseline came out a pixel apart from its display layer's
 * on every other row (canvas fidelity pass 5: R88-X09, R87-X10).
 */
let metricZoom = 1;
export function setGridCanvasTextZoom(zoom: number) {
  metricZoom = zoom > 0 && Number.isFinite(zoom) ? zoom : 1;
}

/** `font` at the zoom's size. */
function zoomedFont(font: string, zoom: number): string {
  return zoom === 1 ? font : font.replace(/(\d*\.?\d+)px/, (_, size: string) => `${Number(size) * zoom}px`);
}

/** The font's ascent and descent as layout uses them (whole pixels of the zoomed font, as the line box is built from), in the font's own pixels. */
export function gridCanvasFontMetrics(font: string): { ascent: number; descent: number } | null {
  const zoom = metricZoom;
  const at = zoomedFont(font, zoom);
  const known = metricCache.get(at);
  if (known) return known;
  const ctx = fontMeasureContext(at, "normal");
  if (!ctx) return null;
  const metrics = ctx.measureText("Hg");
  if (!Number.isFinite(metrics.fontBoundingBoxAscent) || !Number.isFinite(metrics.fontBoundingBoxDescent)) return null;
  const answer = { ascent: Math.round(metrics.fontBoundingBoxAscent) / zoom, descent: Math.round(metrics.fontBoundingBoxDescent) / zoom };
  metricCache.set(at, answer);
  return answer;
}

/**
 * The leading layout puts above an inline box's ascent: half of `leading`
 * (the zoomed sheet's pixels), floored to a whole pixel. Chromium and WebKit
 * both floor it (the rest goes below): measured against their own baselines,
 * 1,440 of 1,440 lines in Chromium at 90%, 100% and 130%, where the exact
 * half put half the lines' text a pixel low.
 */
function leadingAbove(leading: number): number {
  return Math.floor(leading / 2);
}

/** The sheet's zoom the text is laid out under (``setGridCanvasTextZoom``). */
export function gridCanvasTextZoom(): number {
  return metricZoom;
}

/** The share of `leading` (the zoomed sheet's pixels) layout puts above an inline box's ascent. */
export function gridCanvasLeadingAbove(leading: number): number {
  return leadingAbove(leading);
}

/**
 * How far below a line box's top an inline box in `metrics` puts its
 * baseline: its leading's upper share above its ascent, as layout computes
 * it in the zoomed sheet's units (whole-pixel ascent and descent, the line
 * in 1/64 px).
 */
function baselineInLine(lineHeight: number, metrics: { ascent: number; descent: number }): number {
  const zoom = metricZoom;
  const ascent = Math.round(metrics.ascent * zoom), descent = Math.round(metrics.descent * zoom);
  const line = Math.round(lineHeight * zoom * 64) / 64;
  return (leadingAbove(line - ascent - descent) + ascent) / zoom;
}

/**
 * The height of a `line-height: normal` line in `font`, as layout builds it
 * (ascent, descent and the font's line gap): read once per font from a
 * probe line, since a canvas reports no line gap.
 */
export function gridCanvasNormalLineHeight(font: string): number | null {
  // WebKit sets a normal line in the stand-in's own metrics: the layer is
  // given the Office face's line there (officeNormalLineHeightPx), and so is
  // the canvas.
  const office = officeNormalLineOfFont(font);
  if (office !== null) return office;
  const zoom = metricZoom;
  const at = zoomedFont(font, zoom);
  const known = normalLineCache.get(at);
  if (known !== undefined) return known;
  if (typeof document === "undefined" || !document.body) return null;
  // A contained box of its own, so reading the probe lays out only it.
  const host = document.createElement("div");
  host.style.cssText = "position:absolute;left:-9999px;top:0;width:400px;height:400px;contain:strict;visibility:hidden;";
  const probe = document.createElement("span");
  probe.style.cssText = "display:block;white-space:pre;";
  probe.style.font = at;
  probe.style.lineHeight = "normal";
  probe.textContent = "Hg";
  host.appendChild(probe);
  document.body.appendChild(host);
  const height = probe.getBoundingClientRect().height;
  host.remove();
  if (!(height > 0)) return null;
  normalLineCache.set(at, height / zoom);
  return height / zoom;
}

/** The Office face's normal line for a canvas `font` ("[style] [weight] SIZEpx family, ..."), where the engine ignores its overrides. */
function officeNormalLineOfFont(font: string): number | null {
  const match = /(\d*\.?\d+)px\s+(.+)$/.exec(font);
  if (!match) return null;
  return officeNormalLineHeightPx(match[2], Number(match[1]));
}

/** A cell's padding box in the canvas tile's pixels. */
export type GridCanvasBox = { readonly left: number; readonly top: number; readonly width: number; readonly height: number };

/** The box a cell's text is laid out and clipped in: the cell's own, or, for
 * text that overflows into its neighbours, that box pulled out by its reach. */
export function gridCanvasTextBox(text: GridCanvasText, box: GridCanvasBox): GridCanvasBox {
  // A turned line is placed from the cell's own box; its reach only widens the tile.
  if (!text.reach || text.turn) return box;
  return { left: box.left - text.reach.left, top: box.top, width: box.width + text.reach.left + text.reach.right, height: box.height };
}

/**
 * Where the text's line starts and its baseline lies inside `box`, by the
 * display layer's box model: the insets, the flex alignment of one line box
 * of `lineHeight`, the font's ascent and descent centred in that box (CSS
 * half-leading). Null when the font cannot be measured.
 */
export function gridCanvasTextPlacement(text: GridCanvasText, box: GridCanvasBox): { x: number; baseline: number; innerRight: number } | null {
  const metrics = gridCanvasFontMetrics(text.font);
  if (!metrics) return null;
  const lineHeight = text.lineHeight === "normal" ? gridCanvasNormalLineHeight(text.font) : text.lineHeight;
  if (lineHeight === null) return null;
  const innerLeft = box.left + text.padLeft;
  const innerWidth = Math.max(0, box.width - text.padLeft - text.padRight);
  const innerTop = box.top + text.padTop;
  const innerHeight = Math.max(0, box.height - text.padTop - text.padBottom);
  const free = innerWidth - text.width;
  // justify-content: flex-start / center / flex-end in the line's direction.
  const startOffset = text.horizontal === "center" ? free / 2 : text.horizontal === "end" ? free : 0;
  const x = text.direction === "rtl" ? innerLeft + innerWidth - startOffset - text.width : innerLeft + startOffset;
  const lineTop = text.vertical === "top"
    ? innerTop
    : text.vertical === "middle"
      ? innerTop + (innerHeight - lineHeight) / 2
      : innerTop + innerHeight - lineHeight;
  // Each inline box of the line (the layer's strut, the text's run) puts
  // its baseline half its leading below the line's top; the line's baseline
  // is the lowest of them.
  const strut = text.strutFont ? gridCanvasFontMetrics(text.strutFont) : null;
  const baseline = lineTop + Math.max(baselineInLine(lineHeight, metrics), strut ? baselineInLine(lineHeight, strut) : -Infinity);
  return { x, baseline, innerRight: innerLeft + innerWidth };
}

/** Draws one run at `x` and returns where the pen stands after it. */
function paintRun(ctx: CanvasRenderingContext2D, run: GridCanvasText["runs"][number], x: number, baseline: number, text: GridCanvasText, spacing: boolean, last: boolean): number {
  if (spacing) {
    const wanted = `${run.spacing}px`;
    const state = drawnState.get(ctx);
    if (state ? state.spacing !== wanted : ctx.letterSpacing !== wanted) {
      ctx.letterSpacing = wanted;
      if (state) state.spacing = wanted;
    }
    ctx.fillText(run.text, x, baseline);
    if (text.stroke) ctx.strokeText(run.text, x, baseline);
    if (last) return x;
    // The pen moves by the run's width as layout measures it (advances
    // without spacing, plus the spacing after every character).
    return x + (gridCanvasTextWidth([run], text.paintFont ?? text.font, text.kerning) ?? 0);
  }
  // No canvas letter spacing: place each character.
  for (const character of Array.from(run.text)) {
    ctx.fillText(character, x, baseline);
    if (text.stroke) ctx.strokeText(character, x, baseline);
    x += ctx.measureText(character).width + run.spacing;
  }
  return x;
}

/** The faux bold outline, as -webkit-text-stroke draws it: the text's colour, centred on the glyph's edge. */
function applyStroke(ctx: CanvasRenderingContext2D, text: GridCanvasText) {
  if (!text.stroke) return;
  ctx.lineWidth = text.stroke;
  ctx.strokeStyle = text.color;
  ctx.lineJoin = "miter";
}

/** Where a wrapped text's paragraph starts, as the multiline layer's flex
 * column puts it in the cell's text area. */
function paragraphTop(text: GridCanvasText, box: GridCanvasBox): number {
  const innerTop = box.top + text.padTop;
  const innerHeight = Math.max(0, box.height - text.padTop - text.padBottom);
  const height = text.paragraphHeight ?? 0;
  if (text.vertical === "top" || (text.safeVertical && height > innerHeight)) return innerTop;
  return text.vertical === "middle" ? innerTop + (innerHeight - height) / 2 : innerTop + innerHeight - height;
}

function paintGridCanvasLines(ctx: CanvasRenderingContext2D, text: GridCanvasText, lines: ReadonlyArray<GridCanvasLine>, box: GridCanvasBox, devicePixels: number, originY: number): boolean {
  const metrics = gridCanvasFontMetrics(text.font);
  if (!metrics) return false;
  applyFont(ctx, text.paintFont ?? text.font, text.kerning);
  ctx.fillStyle = text.color;
  applyStroke(ctx, text);
  ctx.textBaseline = "alphabetic";
  ctx.textAlign = "left";
  const spacing = typeof (ctx as { letterSpacing?: unknown }).letterSpacing === "string";
  const innerLeft = box.left + text.padLeft;
  const innerWidth = Math.max(0, box.width - text.padLeft - text.padRight);
  const top = paragraphTop(text, box);
  for (const line of lines) {
    if (!line.runs.length) continue;
    const free = innerWidth - line.width;
    let x = innerLeft + (text.horizontal === "center" ? free / 2 : text.horizontal === "end" ? free : 0);
    const exact = top + (line.baseline ?? line.top + baselineInLine(line.height, metrics));
    const baseline = Math.round((originY + exact) * devicePixels) / devicePixels - originY;
    line.runs.forEach((run, index) => { x = paintRun(ctx, run, x, baseline, text, spacing, index === line.runs.length - 1); });
  }
  if (spacing && !drawnState.has(ctx) && ctx.letterSpacing !== "0px") ctx.letterSpacing = "0px";
  return true;
}

/**
 * Angled text as its display layer draws it: one line box of `lineHeight`
 * turned about its top-left corner, the baseline half the leading below its
 * top. Not snapped (a turned baseline lies on no pixel row).
 */
function paintTurnedText(ctx: CanvasRenderingContext2D, text: GridCanvasText, turn: NonNullable<GridCanvasText["turn"]>, box: GridCanvasBox): boolean {
  const metrics = gridCanvasFontMetrics(text.font);
  if (!metrics) return false;
  const strut = text.strutFont ? gridCanvasFontMetrics(text.strutFont) : null;
  const baseline = Math.max(baselineInLine(turn.lineHeight, metrics), strut ? baselineInLine(turn.lineHeight, strut) : -Infinity);
  applyFont(ctx, text.paintFont ?? text.font, text.kerning);
  ctx.fillStyle = text.color;
  applyStroke(ctx, text);
  ctx.textBaseline = "alphabetic";
  ctx.textAlign = "left";
  const spacing = typeof (ctx as { letterSpacing?: unknown }).letterSpacing === "string";
  // The frame put back by hand: a save and restore would drop the tracked font state.
  const frame = ctx.getTransform();
  ctx.translate(box.left + turn.x, box.top + turn.y);
  ctx.rotate((turn.rotate * Math.PI) / 180);
  let x = 0;
  text.runs.forEach((run, index) => { x = paintRun(ctx, run, x, baseline, text, spacing, index === text.runs.length - 1); });
  ctx.setTransform(frame);
  if (spacing && !drawnState.has(ctx) && ctx.letterSpacing !== "0px") ctx.letterSpacing = "0px";
  return true;
}

/**
 * Paints `text` into `ctx` (already scaled to CSS pixels) inside `box`. The
 * baseline is snapped to the device pixel grid (whose origin is `originY`
 * above the box's frame) as the browser snaps a text run's baseline; the pen
 * advances run by run with each run's spacing.
 */
export function paintGridCanvasText(ctx: CanvasRenderingContext2D, text: GridCanvasText, box: GridCanvasBox, devicePixels: number, originY = 0): boolean {
  if (text.lines) return paintGridCanvasLines(ctx, text, text.lines, box, devicePixels, originY);
  if (text.turn) return paintTurnedText(ctx, text, text.turn, box);
  const placed = gridCanvasTextPlacement(text, box);
  if (!placed) return false;
  applyFont(ctx, text.paintFont ?? text.font, text.kerning);
  ctx.fillStyle = text.color;
  applyStroke(ctx, text);
  ctx.textBaseline = "alphabetic";
  ctx.textAlign = "left";
  // Snapped where the screen's pixels are: `originY` is the box's frame on the screen.
  const baseline = Math.round((originY + placed.baseline) * devicePixels) / devicePixels - originY;
  // Canvas letter spacing is recent (Chrome 99, Safari 18); without it each character is placed.
  const spacing = typeof (ctx as { letterSpacing?: unknown }).letterSpacing === "string";
  const segments: ReadonlyArray<GridCanvasSegment> = text.segments ?? [{ runs: text.runs, anchor: "start", offset: 0 }];
  for (const segment of segments) {
    let x = segment.anchor === "end" ? placed.innerRight - segment.offset : placed.x + segment.offset;
    segment.runs.forEach((run, index) => { x = paintRun(ctx, run, x, baseline, text, spacing, index === segment.runs.length - 1); });
  }
  // A tracked context keeps its spacing for the next cell (most share it).
  if (spacing && !drawnState.has(ctx) && ctx.letterSpacing !== "0px") ctx.letterSpacing = "0px";
  return true;
}
