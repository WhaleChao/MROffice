/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";
import { clampNativeAutomaticRowHeight, nativeTextAdvances, nativeTextPrinterDpi, resolveSpreadsheetTextStyle } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { OFFICE_PRINT_DPI, officePrintedRowDots } from "./office-print-metrics";
import { layoutProbeParent } from "@mrnote/core/web/lib/layout-probe-root.ts";
import { spreadsheetEngineIdealAdvances, spreadsheetEngineLineBoxes, spreadsheetEngineLineMetrics, spreadsheetEngineParagraphHeight, spreadsheetTextEngineEpoch, spreadsheetTextEngineOn } from "@mrnote/core/web/lib/spreadsheet/spreadsheet-text-engine.ts";
import { breakLines, excelLineBreakOpportunities, lineBreakOpportunities } from "@mrnote/core/web/lib/text-engine/line-break.ts";
import { graphemesOf } from "@mrnote/core/web/lib/text-engine/measure.ts";
import { fontLoadReach, type FontLoadReach } from "@mrnote/core/web/lib/spreadsheet/font-load-reach.ts";
import { canvasLineBoxes } from "./canvas-line-boxes";
import { officeRowHeight, officeRowMetricsKnown } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

export interface NativeTextPart {
  text: string;
  style?: CSSProperties;
  /** Original source-run lead, retained when wrapping splits off a fragment. */
  fontLeadingText?: string;
}
export function nativeTextFontLead(part: NativeTextPart): string {
  const point = part.text.codePointAt(0);
  return part.fontLeadingText ?? (point === undefined ? "" : String.fromCodePoint(point));
}
/** A part whose size is relative to its paragraph ("75%": a superscript run
 * with no size of its own, cell-rich-text) is measured and drawn at the
 * resolved pixels, since every layout step merges the part over the
 * paragraph style and a percentage would no longer know what it is of. The
 * same array comes back when there is nothing to resolve. */
export function resolveRelativePartFontSizes(parts: readonly NativeTextPart[], style: CSSProperties): readonly NativeTextPart[] {
  const base = style.fontSize;
  if (typeof base !== "number" || !parts.some(part => typeof part.style?.fontSize === "string" && part.style.fontSize.endsWith("%"))) return parts;
  return parts.map(part => {
    const size = part.style?.fontSize;
    const percent = typeof size === "string" && size.endsWith("%") ? Number.parseFloat(size) : Number.NaN;
    return Number.isFinite(percent) ? { ...part, style: { ...part.style, fontSize: base * percent / 100 } } : part;
  });
}
/** A style key (`"1"`) that makes a cell's engine wraps keep room for the
 * space after a word (breakLines `keepSpace`): an imported Windows grid's
 * printed cell in a book that wrapped complex script (lookwin p144-lw). */
export const NATIVE_KEEP_SPACE_MARK = "--mrai-office-keep-space";
/** A style mark: the paragraph's room is the printer's re-wrap room (the
 * printed column less its margin, lookwin p138-lw), whose lines Windows
 * Excel's printer breaks at its own whole printer pixels per em, not the
 * screen's (excelWrapRoom). */
export const NATIVE_PRINTED_ROOM_MARK = "--mrai-office-printed-room";
export interface NativeTextLine { parts: NativeTextPart[]; hardBreak: string; emptyStyle?: CSSProperties; breakStyle?: CSSProperties }
/** Logical, unscaled cell geometry; independent of the painted text width. */
export interface NativeTextAlignmentBox { width: number; textLeft: number; textRight: number; wrapText: boolean }
const layouts = new Map<string, readonly NativeTextLine[]>();
let listening = false;
let lineMetricContext: CanvasRenderingContext2D | null | undefined;
const lineMetrics = new Map<string, number>();

/** Screen line boxes are resolved before view zoom, independently of browser
 * `normal` line-height rounding. The em floor and ascender clearance are
 * font-metric rules, shared by available and substituted faces. */
export function measureNativeParagraphLineHeight(parts: readonly NativeTextPart[], style: CSSProperties, dpi: number): number | undefined {
  if (dpi !== 72 || typeof document === "undefined" || typeof navigator === "undefined" || !/Mac/i.test(navigator.platform)) return undefined;
  if (parts.some(part => /[\u0590-\u0fff\u200e\u200f\u202a-\u202e\u2066-\u2069]/u.test(part.text))) return undefined;
  if (lineMetricContext === undefined) {
    lineMetricContext = document.createElement("canvas").getContext("2d");
    document.fonts?.addEventListener("loadingdone", (event) => dropTextCacheTouched(lineMetrics, event, (key, reach) => {
      const [font, , text] = JSON.parse(key) as [string, number, string];
      return reach.touchesFont(String(font), String(text));
    }));
  }
  if (!lineMetricContext) return undefined;
  const base: NativeTextPart = { text: " ", fontLeadingText: parts[0] ? nativeTextFontLead(parts[0]) : "" };
  let height = 0;
  for (const part of [base, ...parts]) {
    const resolved = resolveSpreadsheetTextStyle({ ...style, ...part.style }, nativeTextFontLead(part), dpi);
    if (typeof resolved.fontSize !== "number" || !Number.isFinite(resolved.fontSize) || resolved.fontSize <= 0) return undefined;
    const em = resolved.fontSize * dpi / 96;
    const font = fontCss({ ...resolved, fontSize: 2048 });
    const key = JSON.stringify([font, em, part.text]);
    let measured = lineMetrics.get(key);
    if (measured === undefined) {
      // The primary face's own ascent and descent from its file where the
      // text engine knows the face (the same in every browser), else the
      // browser's.
      const engine = spreadsheetEngineLineMetrics(font);
      let fontAscent: number, fontDescent: number;
      if (engine) {
        fontAscent = engine.ascent;
        fontDescent = engine.descent;
      } else {
        lineMetricContext.font = font;
        const metrics = lineMetricContext.measureText(part.text || " ");
        fontAscent = metrics.fontBoundingBoxAscent;
        fontDescent = metrics.fontBoundingBoxDescent;
      }
      const ascent = Math.round(fontAscent / 2048 * em);
      const descent = Math.round(fontDescent / 2048 * em);
      if (!Number.isFinite(ascent) || !Number.isFinite(descent) || ascent <= 0 || descent < 0) return undefined;
      measured = (Math.max(em, ascent) + descent + 1 + (ascent > em ? 1 : 0)) * 96 / dpi;
      // Keyed by text: a sheet's cells are thousands of keys.
      if (lineMetrics.size >= 32_768) lineMetrics.clear();
      lineMetrics.set(key, measured);
    }
    height = Math.max(height, measured);
  }
  return height || undefined;
}

/**
 * The pitch a Windows grid's paragraph is set at: Windows Excel lays each line
 * of a wrapped cell at its faces' automatic row height (officeRowHeight),
 * the height it fits the row to, so the lines fill the row (two lines of
 * Calibri 11: a 40 px row, the lines 20 px apart; Cambria 11: 19 px). The
 * browser's normal line box set them 18 and 17 px apart, the text standing
 * short of the row it had been fitted to. Undefined off a 96 dpi grid, or for
 * a face the table does not know.
 */
export function windowsNativeLinePitch(parts: readonly NativeTextPart[], style: CSSProperties, dpi: number): number | undefined {
  if (dpi !== 96) return undefined;
  let pitch = 0;
  for (const part of parts.length ? parts : [{ text: "" }]) {
    const merged = { ...style, ...part.style };
    const family = typeof merged.fontFamily === "string" ? merged.fontFamily.split(",")[0]?.trim().replace(/^["']|["']$/g, "").trim() : "";
    const size = typeof merged.fontSize === "number" ? merged.fontSize : Number.NaN;
    if (!family || !(size > 0) || !officeRowMetricsKnown(family, size * 0.75)) return undefined;
    // Printed, a line is the printer's line of its face, whatever the sheet's
    // Normal font (lookwin p107-lw: Calibri 11 lines 121 dots apart, 19.36 px,
    // in an Aptos Narrow 11 sheet; the API's office_printed_line_dots).
    const printed = nativeTextPrinterDpi() > 0 ? officePrintedRowDots({ family, sizePt: size * 0.75 }) : undefined;
    pitch = Math.max(pitch, printed !== undefined ? printed * 96 / OFFICE_PRINT_DPI : officeRowHeight({ family, sizePt: size * 0.75 }));
  }
  return pitch || undefined;
}

/** The fixed pitch a paragraph is painted at: a Mac grid's native line
 * height, a Windows grid's Excel line height, else none (the line boxes of
 * the faces). */
export function nativeParagraphLinePitch(parts: readonly NativeTextPart[], style: CSSProperties, dpi: number): number | undefined {
  return measureNativeParagraphLineHeight(parts, style, dpi) ?? windowsNativeLinePitch(parts, style, dpi);
}

/** Whether a font load can change `parts` drawn in `style`: what any of its
 * runs is drawn in, resolved as the grid draws it (font-load-reach). */
export function nativeTextTouchedBy(reach: FontLoadReach, parts: readonly NativeTextPart[], style: CSSProperties, dpi: number): boolean {
  const text = parts.map((part) => part.text).join("");
  for (const part of parts.length ? parts : [{ text: "" }]) {
    const resolved = resolveSpreadsheetTextStyle({ ...style, ...part.style }, nativeTextFontLead(part), dpi);
    if (reach.touches(typeof resolved.fontFamily === "string" ? resolved.fontFamily : undefined, text)) return true;
  }
  return false;
}

/** A text cache keyed `[parts, style, width, dpi, ...]`. */
function textKeyTouched(key: string, reach: FontLoadReach): boolean {
  const [parts, style, , dpi] = JSON.parse(key) as [NativeTextPart[], CSSProperties, number, number];
  if (!Array.isArray(parts) || !style || typeof style !== "object") return true;
  return nativeTextTouchedBy(reach, parts, style, typeof dpi === "number" ? dpi : 96);
}




function dropTextCacheTouched<V>(cache: Map<string, V>, event: Event, touched: (key: string, reach: FontLoadReach) => boolean): void {
  const reach = fontLoadReach(event);
  if (!reach) {
    cache.clear();
    return;
  }
  for (const key of [...cache.keys()]) {
    let drop = true;
    try {
      drop = touched(key, reach);
    } catch {
      drop = true;
    }
    if (drop) cache.delete(key);
  }
}

/**
 * Where the browser stands the baseline of a line set in `font` with a fixed
 * `lineHeightPx`, px below the line box's top: what its own layout does
 * (the strut's ascent and half-leading, rounded as the engine rounds them),
 * read once per font, size and line height from a probe. Undefined where
 * nothing lays out (no document, or a probe the engine does not lay out).
 * A Windows grid's printed line stands its baseline where Excel's does by it
 * (SpreadsheetNativeParagraph officeBaseline).
 */
export function nativeStrutBaselinePx(font: { fontFamily?: unknown; fontSize?: unknown; fontWeight?: unknown; fontStyle?: unknown }, lineHeightPx: number | "normal", generation = 0): number | undefined {
  if (typeof document === "undefined" || !document.body || (lineHeightPx !== "normal" && !(lineHeightPx > 0))) return undefined;
  const size = typeof font.fontSize === "number" ? `${font.fontSize}px` : typeof font.fontSize === "string" ? font.fontSize : "";
  const family = typeof font.fontFamily === "string" ? font.fontFamily : "";
  if (!size || !family) return undefined;
  const key = `${String(font.fontStyle ?? "normal")}\u0000${String(font.fontWeight ?? "normal")}\u0000${size}\u0000${family}\u0000${lineHeightPx}\u0000${generation}`;
  if (strutBaselines.has(key)) return strutBaselines.get(key) ?? undefined;
  const block = document.createElement("div");
  block.style.cssText = "position:absolute;visibility:hidden;left:-10000px;top:0;white-space:pre;margin:0;padding:0;border:0;";
  block.style.fontFamily = family;
  block.style.fontSize = size;
  block.style.fontWeight = String(font.fontWeight ?? "normal");
  block.style.fontStyle = String(font.fontStyle ?? "normal");
  block.style.lineHeight = lineHeightPx === "normal" ? "normal" : `${lineHeightPx}px`;
  // A box of no height on the baseline: its top is the baseline.
  const mark = document.createElement("span");
  mark.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline;";
  block.append(mark);
  document.body.append(block);
  const baseline = mark.getBoundingClientRect().top - block.getBoundingClientRect().top;
  const laidOut = block.getBoundingClientRect().height > 0;
  block.remove();
  const answer = laidOut && Number.isFinite(baseline) && baseline > 0 ? baseline : null;
  if (strutBaselines.size > 2000) strutBaselines.clear();
  strutBaselines.set(key, answer);
  return answer ?? undefined;
}
const strutBaselines = new Map<string, number | null>();

/** AutoFit allocates one device pixel of clearance per line. This is distinct
 * from both the default row and a single-line font's base height. */
export function measureNativeAutoRowHeight(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number, numericValue = false): number | undefined {
  const lineHeight = measureNativeParagraphLineHeight(parts, style, dpi);
  if (lineHeight === undefined) return undefined;
  const lines = layoutNativeText(parts, style, width, dpi);
  const devicePitch = Math.round(lineHeight * dpi / 96);
  // A displayed number (including a date or formula result) uses the font
  // base when it fits one line. The extra wrapped-text clearance is not a
  // property of the number-format string or of digits in an actual text value.
  if (numericValue && lines?.length === 1) return clampNativeAutomaticRowHeight(lineHeight);
  return lines ? clampNativeAutomaticRowHeight(lines.length * (devicePitch + 1) * 96 / dpi) : undefined;
}

/** Literal text can look numeric. Its stored type takes precedence over the
 * evaluator's arithmetic coercion; formulas use their actual result type. */
export function nativeAutoRowIsNumeric(cell: { kind?: string; formula?: string; value?: string }, computed: unknown): boolean {
  if (cell.formula || (cell.kind === undefined && cell.value?.startsWith("="))) return typeof computed === "number";
  return cell.kind === "number";
}

const unhungLines = new WeakMap<readonly NativeTextLine[], readonly NativeTextLine[]>();
/** A wrapped paragraph's lines without the spaces that end them: Excel sets
 * a wrapped line by its words, the spaces at its break (or after its last
 * word) hanging past the edge, so they move neither a centred nor a
 * right-aligned line (lookwin R223-X13: "Population" and 45 spaces, centred
 * in K3:M3, printed at the merge's left edge). A line of nothing but spaces
 * keeps them. The same array for the same lines. */
export function nativeLinesWithoutHangingSpaces(lines: readonly NativeTextLine[]): readonly NativeTextLine[] {
  const known = unhungLines.get(lines);
  if (known) return known;
  let changed = false;
  const out = lines.map((line) => {
    const parts = [...line.parts];
    while (parts.length) {
      const last = parts[parts.length - 1];
      if (!last) break;
      const text = last.text.replace(/ +$/u, "");
      if (text === last.text) break;
      if (text) { parts[parts.length - 1] = { ...last, text }; break; }
      parts.pop();
    }
    if (!parts.length || (parts.length === line.parts.length && parts.every((part, index) => part === line.parts[index]))) return line;
    changed = true;
    return { ...line, parts };
  });
  const result = changed ? out : lines;
  unhungLines.set(lines, result);
  return result;
}

export function measureNativeLineAdvance(parts: readonly NativeTextPart[], style: CSSProperties, dpi: number): number | undefined {
  let width = 0;
  for (const part of parts) {
    if (/[\r\n]/u.test(part.text)) return undefined;
    const runs = nativeTextAdvances(part.text, fontCss(resolveSpreadsheetTextStyle({ ...style, ...part.style }, nativeTextFontLead(part), dpi)), dpi, 1, "measure", width);
    if (runs.some(run => run.advance === undefined)) return undefined;
    width += runs.reduce((sum, run) => sum + (run.advance ?? 0), 0);
  }
  return width;
}

/** Mac Excel aligns screen text using device advances, including substitute
 * faces. Round the cell-relative origin before the page-layout scale.
 */
export function nativeTextAlignmentOffset(advance: number | undefined, align: CSSProperties["textAlign"], box: NativeTextAlignmentBox | undefined, dpi: number, scale = 1): number | undefined {
  if (dpi !== 72 || !box || advance === undefined || ![advance, box.width, box.textLeft, box.textRight, scale].every(Number.isFinite) || scale <= 0) return undefined;
  const unit = 96 / dpi;
  const width = Math.round(box.width / unit);
  const text = Math.round(advance / unit);
  // Unwrapped text centers within the asymmetric gridline/padding box.
  // Wrapped paragraphs center within the full cell. Native paired fixtures
  // differ only in wrapText and retain identical right-aligned origins.
  const centerInset = box.wrapText ? 0 : Math.round((box.textLeft - box.textRight) / unit);
  const start = align === "center" ? Math.round((width - text + centerInset) / 2)
    : align === "right" || align === "end" ? width - Math.round(box.textRight / unit) - text : undefined;
  return start === undefined ? undefined : (start * unit - box.textLeft) * scale;
}

export function nativeParagraphFontStyle(style: CSSProperties, scale = 1, text = "", dpi?: number): CSSProperties {
  const resolved = resolveSpreadsheetTextStyle(style, text, dpi);
  return {
    fontFamily: resolved.fontFamily,
    fontSize: typeof resolved.fontSize === "number" ? resolved.fontSize * scale : resolved.fontSize,
    fontWeight: resolved.fontWeight,
    fontStyle: resolved.fontStyle,
  };
}

export function nativeParagraphLineStyle(line: NativeTextLine, style: CSSProperties, scale = 1, dpi?: number, lineHeight?: number): CSSProperties {
  return {
    display: "block", whiteSpace: "pre", lineHeight: lineHeight === undefined ? "normal" : `${lineHeight * scale}px`,
    ...(lineHeight === undefined ? {} : { height: `${lineHeight * scale}px` }),
    ...(!line.parts.length ? { ...nativeParagraphFontStyle({ ...style, ...line.emptyStyle }, scale, "", dpi), height: "1lh" } : {}),
  };
}

/** Measure the same line boxes used for painting after native wrapping.
 * This is paragraph height, not the separate native default-row/AutoFit rule.
 * Unwrapped spans cannot change the already resolved line boundaries. */
export function measureNativeTextHeight(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number): number | undefined {
  const lines = layoutNativeText(parts, style, width, dpi);
  if (!lines) return undefined;
  const lineHeight = measureNativeParagraphLineHeight(parts, style, dpi);
  // The logical height is independent of CSS's fractional layout grid.
  // Painting places each baseline from its cumulative logical origin.
  if (lineHeight !== undefined) return lines.length * lineHeight;
  const engine = engineParagraphHeight(lines, parts, style, dpi);
  if (engine !== undefined) return engine;
  const key = JSON.stringify([parts, style, width, dpi]);
  const known = paragraphHeights.get(key);
  if (known !== undefined) return known;
  const root = buildNativeHeightProbe(lines, parts, style, width, dpi);
  try {
    layoutProbeParent(document).append(root);
    return rememberParagraphHeight(key, root.getBoundingClientRect().height);
  } finally { root.remove(); }
}

// The browser-measured height of a paragraph whose line box has no native
// rule here (any grid that is not a 72 dpi Mac one).
const paragraphHeights = new Map<string, number>();
let paragraphHeightsListening = false;
function rememberParagraphHeight(key: string, height: number): number {
  if (!paragraphHeightsListening) {
    document.fonts?.addEventListener("loadingdone", (event) => dropTextCacheTouched(paragraphHeights, event, textKeyTouched));
    paragraphHeightsListening = true;
  }
  if (paragraphHeights.size >= NATIVE_TEXT_LAYOUT_CACHE_LIMIT) paragraphHeights.clear();
  paragraphHeights.set(key, height);
  return height;
}

function buildNativeHeightProbe(lines: readonly NativeTextLine[], parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number): HTMLElement {
  const root = document.createElement("div");
  root.setAttribute("aria-hidden", "true");
  const apply = (node: HTMLElement, css: CSSProperties) => {
    for (const [key, value] of Object.entries(css)) {
      if (value === undefined || value === null) continue;
      node.style.setProperty(key.replace(/[A-Z]/g, char => `-${char.toLowerCase()}`),
        typeof value === "number" && key === "fontSize" ? `${value}px` : String(value));
    }
  };
  apply(root, nativeParagraphFontStyle(style, 1, parts[0] ? nativeTextFontLead(parts[0]) : "", dpi));
  Object.assign(root.style, {
    position: "fixed", left: "0", top: "0", visibility: "hidden", pointerEvents: "none",
    contain: "layout style paint", width: `${width}px`, lineHeight: "normal", border: "0", padding: "0", margin: "0",
  });
  for (const line of lines) {
    const box = document.createElement("span");
    apply(box, nativeParagraphLineStyle(line, style, 1, dpi));
    for (const part of line.parts) {
      const span = document.createElement("span");
      apply(span, { ...part.style, ...nativeParagraphFontStyle({ ...style, ...part.style }, 1, nativeTextFontLead(part), dpi) });
      span.textContent = part.text;
      box.appendChild(span);
    }
    if (line.hardBreak) {
      const terminal = document.createElement("span");
      apply(terminal, nativeParagraphFontStyle({ ...style, ...line.breakStyle }, 1, "", dpi));
      box.appendChild(terminal);
    }
    root.appendChild(box);
  }
  return root;
}

/**
 * {@link measureNativeTextHeight} for many paragraphs under two layouts of
 * the page (their lines, then their heights) instead of two per paragraph.
 */
export function measureNativeTextHeightBatch(requests: readonly NativeTextLayoutRequest[]): void {
  layoutNativeTextBatch(requests);
  const pending = new Map<string, HTMLElement>();
  let atoms = 0;
  const flush = () => {
    if (pending.size === 0) return;
    const host = document.createElement("div");
    host.setAttribute("aria-hidden", "true");
    for (const root of pending.values()) host.appendChild(root);
    try {
      layoutProbeParent(document).append(host);
      for (const [key, root] of pending) rememberParagraphHeight(key, root.getBoundingClientRect().height);
    } finally { host.remove(); pending.clear(); atoms = 0; }
  };
  for (const { parts, style, width, dpi } of requests) {
    if (measureNativeParagraphLineHeight(parts, style, dpi) !== undefined) continue;
    const key = JSON.stringify([parts, style, width, dpi]);
    if (paragraphHeights.has(key) || pending.has(key)) continue;
    const lines = layoutNativeText(parts, style, width, dpi);
    if (lines && engineParagraphHeight(lines, parts, style, dpi) !== undefined) continue;
    if (!lines) continue;
    pending.set(key, buildNativeHeightProbe(lines, parts, style, width, dpi));
    atoms += lines.reduce((sum, line) => sum + line.parts.length + 1, 0);
    if (atoms >= NATIVE_TEXT_PROBE_BATCH_ATOMS) flush();
  }
  flush();
}

function fontCss(style: CSSProperties): string {
  return `${style.fontStyle ?? "normal"} ${style.fontWeight ?? 400} ${typeof style.fontSize === "number" ? `${style.fontSize}px` : style.fontSize ?? "12px"} ${style.fontFamily ?? "sans-serif"}`;
}

function fontStyle(style: CSSProperties, scale: number): Partial<CSSStyleDeclaration> {
  return {
    fontFamily: style.fontFamily ?? "sans-serif",
    fontSize: typeof style.fontSize === "number" ? `${style.fontSize * scale}px` : style.fontSize ?? "12px",
    fontWeight: String(style.fontWeight ?? 400), fontStyle: style.fontStyle ?? "normal",
  };
}

/**
 * The room a wrapped paragraph breaks in on a grid of Windows pixels, in the
 * grid's pixels at the cell's own size. GDI sets a face at a whole-pixel em
 * (Calibri 11, 14.67 px, at 15 px), so Windows Excel's lines are those of the
 * text 15/14.67 wider than the size the grid draws: the same break, in a room
 * narrowed by that ratio. Against the line counts Windows Excel's own row
 * heights give for 3,900 wrapped cells (xlsx-rounds tools/wrapfit.py), lines
 * broken so agree with Excel for 76.5% of them, against 73.1% at the drawn
 * size (Calibri 11 64.6% to 68.9%, Arial 10 83.5% to 87.6%). A Mac grid and
 * unwrapped text (an unbounded room) are as they were.
 */
export function excelWrapRoom(width: number, style: CSSProperties, dpi: number): number {
  const size = style.fontSize;
  if (dpi !== 96 || !(width < 1e5) || typeof size !== "number" || !(size > 0)) return width;
  // The printer's re-wrap: its pixels per em are the 600-dpi ones (Calibri
  // 11: 92 for 91.67), a room 0.4 % short of the design widths' where the
  // screen's 15 px em cut 2.2 % (R35-X05's 356 pt note line, 360.5 pt of
  // room: Excel prints it whole, the screen's room sent "own" down). The
  // API prints with the printer's own advances (FaceMeasure.printed).
  if ((style as unknown as Record<string, unknown>)[NATIVE_PRINTED_ROOM_MARK] === "1") {
    const printed = size * 600 / 96;
    return width * printed / Math.max(1, Math.floor(printed + 0.5));
  }
  // A fixed-pitch face sets every glyph at round(0.6 em) of GDI's whole-pixel
  // em, the advance the screen's whole pixels already measure (Courier New
  // 10: 8 px at 13 px per em and at 13.33): its room is the column's. Cut to
  // the 14 px em's, R379-X18's 29-character "caretaker_hh_longitude_manual"
  // broke in two in a 247 px column and stood its row two lines tall, where
  // Windows Excel sets it on one.
  if (typeof style.fontFamily === "string" && FIXED_PITCH_OFFICE_FACE.test(style.fontFamily)) return width;
  return width * size / Math.ceil(size - 1e-6);
}
/** Office's fixed-pitch Latin faces, first in a font stack. */
const FIXED_PITCH_OFFICE_FACE = /^\s*["']?(courier new|consolas|lucida console)["']?\s*(,|$)/i;

type NativeTextAtom = {text:string;part:number;marker:HTMLElement;hardBreak?:boolean};

function nativeTextLayoutKey(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number): string | undefined {
  if (typeof document === "undefined" || !document.body || !(width > 0) || !Number.isFinite(width)) return undefined;
  // Preserve the existing whole-run shaping path for joining/bidirectional
  // scripts until a paragraph bidi/shaping contract is available.
  if (parts.some(part => /[\u0590-\u0fff\u200e\u200f\u202a-\u202e\u2066-\u2069]/u.test(part.text))) return undefined;
  if (!listening) {
    document.fonts?.addEventListener("loadingdone", (event) => dropTextCacheTouched(layouts, event, textKeyTouched));
    listening = true;
  }
  // With the text engine on, a layout made while a face loaded is not reused
  // once it has: the key carries the engine's epoch.
  return JSON.stringify([parts,style,width,dpi,spreadsheetTextEngineEpoch()]);
}

/** {@link measureNativeTextHeight}'s probe height, from the engine: each
 * line's runs in their drawn fonts, the paragraph font as the strut, the
 * hard break of an empty line in its own font. */
function engineParagraphHeight(lines: readonly NativeTextLine[], parts: readonly NativeTextPart[], style: CSSProperties, dpi: number): number | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  const strut = fontCss(nativeParagraphFontStyle(style, 1, parts[0] ? nativeTextFontLead(parts[0]) : "", dpi));
  return spreadsheetEngineParagraphHeight(lines.map((line) => line.parts.length
    ? line.parts.map((part) => ({ text: part.text, fontCss: fontCss(resolveSpreadsheetTextStyle({ ...style, ...part.style }, nativeTextFontLead(part), dpi)) }))
    : [{ text: "", fontCss: fontCss(resolveSpreadsheetTextStyle({ ...style, ...line.emptyStyle }, "", dpi)) }]), strut);
}









const probePool: { div: HTMLElement[]; span: HTMLElement[]; br: HTMLElement[]; text: Text[] } = { div: [], span: [], br: [], text: [] };
const PROBE_POOL_LIMIT = 40_000;

function probeElement(kind: "div" | "span" | "br"): HTMLElement {
  return probePool[kind].pop() ?? document.createElement(kind);
}

function probeText(text: string): Text {
  const node = probePool.text.pop();
  if (!node) return document.createTextNode(text);
  node.data = text;
  return node;
}

/** A read probe's nodes back into the pool, detached and stripped. */
function releaseNativeTextProbe(node: Node): void {
  node.parentNode?.removeChild(node);
  for (let child = node.firstChild; child; child = node.firstChild) releaseNativeTextProbe(child);
  if (node.nodeType === 3) {
    if (probePool.text.length < PROBE_POOL_LIMIT) probePool.text.push(node as Text);
    return;
  }
  const element = node as HTMLElement;
  const kind = element.tagName === "DIV" ? "div" : element.tagName === "SPAN" ? "span" : element.tagName === "BR" ? "br" : null;
  if (!kind) return;
  element.removeAttribute("style");
  element.removeAttribute("aria-hidden");
  if (probePool[kind].length < PROBE_POOL_LIMIT) probePool[kind].push(element);
}

/** A span's font as one declaration list (one parse per span, not a write
 * per property), the same properties {@link fontStyle} sets. */
const probeFontCss = new Map<string, string>();
function probeSpanCss(style: CSSProperties, scale: number): string {
  const font = fontStyle(style, scale);
  const key = `${font.fontFamily}|${font.fontSize}|${font.fontWeight}|${font.fontStyle}`;
  let css = probeFontCss.get(key);
  if (css === undefined) {
    const span = document.createElement("span");
    Object.assign(span.style, font);
    css = span.style.cssText;
    if (probeFontCss.size >= 4_096) probeFontCss.clear();
    probeFontCss.set(key, css);
  }
  return css;
}

/** The probe of one text, not yet in the document. */
function buildNativeTextProbe(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number): {root:HTMLElement;atoms:NativeTextAtom[]} {
  width = excelWrapRoom(width, style, dpi);
  const scale = dpi / 96;
  const root = probeElement("div");
  root.setAttribute("aria-hidden","true");
  Object.assign(root.style, fontStyle(resolveSpreadsheetTextStyle(style, parts[0] ? nativeTextFontLead(parts[0]) : "", dpi),scale), {
    position:"fixed", left:"0", top:"0", visibility:"hidden", pointerEvents:"none",
    contain:"layout style paint", width:`${width*scale}px`, whiteSpace:"pre-wrap",
    overflowWrap:"anywhere", lineHeight:"normal", textAlign:"left", border:"0", padding:"0", margin:"0",
  });
  const atoms: NativeTextAtom[] = [];
  const mark = () => {
    const marker = probeElement("span");
    marker.style.cssText = "font-size:0;line-height:0;letter-spacing:0;vertical-align:baseline";
    return marker;
  };
  parts.forEach((part,partIndex) => {
    const resolved = resolveSpreadsheetTextStyle({...style,...part.style}, nativeTextFontLead(part), dpi);
    for (const piece of part.text.split(/(\r\n|\r|\n)/)) {
      if (!piece) continue;
      if (/^(\r\n|\r|\n)$/.test(piece)) {
        const previous=atoms.at(-1);
        // A lone CR is drawn as nothing, Wrap Text on or off: Excel breaks
        // only at a line feed (RO050_x000D, X-247). One that ends a run
        // before the next run's LF joins that break.
        const nextPart = parts[partIndex + 1];
        const crBeforeLf = piece === "\r" && part.text.endsWith("\r") && nextPart?.text.startsWith("\n");
        if (piece === "\r" && !crBeforeLf) continue;
        if(piece==="\n" && previous?.hardBreak && previous.text==="\r") { previous.text+="\n"; continue; }
        const marker = mark(), holder=probeElement("span");
        holder.style.cssText = probeSpanCss(resolved, scale);
        holder.appendChild(marker);holder.appendChild(probeElement("br"));root.appendChild(holder);
        atoms.push({text:piece,part:partIndex,marker,hardBreak:true});
        continue;
      }
      const spanCss = probeSpanCss(resolved, scale);
      for (const advance of nativeTextAdvances(piece,fontCss(resolved),dpi,scale,"measure")) {
        const span=probeElement("span"),marker=mark();
        span.style.cssText = spanCss;
        span.style.fontKerning = "none";
        span.style.letterSpacing = `${advance.spacing*scale}px`;
        span.appendChild(probeText(advance.text));span.appendChild(marker);root.appendChild(span);
        atoms.push({text:advance.text,part:partIndex,marker});
      }
    }
  });
  glueProbeToExcelBreaks(atoms);
  return {root,atoms};
}











function glueProbeToExcelBreaks(atoms: readonly NativeTextAtom[]): void {
  let start = 0;
  while (start < atoms.length) {
    if (atoms[start].hardBreak) { start += 1; continue; }
    let end = start;
    while (end < atoms.length && !atoms[end].hardBreak) end += 1;
    const texts = atoms.slice(start, end).map((atom) => atom.text);
    const opportunities = excelLineBreakOpportunities(texts);
    for (let index = 0; index < texts.length - 1; index += 1) {
      const text = texts[index], next = texts[index + 1];
      const glue = next === " " ? "" : text === " " ? (opportunities[index + 1] ? "" : "\u2060") : text === "-" ? "\u200b" : "\u2060";
      if (glue) atoms[start + index].marker.appendChild(probeText(glue));
    }
    start = end;
  }
}










function probeLineStarts(atoms: readonly NativeTextAtom[], top: (index: number) => number): Set<number> {
  const starts = new Set<number>();
  let index = 0;
  while (index < atoms.length) {
    if (atoms[index].hardBreak) { index += 1; continue; }
    let end = index;
    while (end < atoms.length && !atoms[end].hardBreak) end += 1;
    let lineStart = index;
    while (lineStart < end) {
      const baseline = top(lineStart);
      let low = lineStart + 1, high = end;
      while (low < high) {
        const middle = (low + high) >> 1;
        if (Math.abs(top(middle) - baseline) > .1) high = middle; else low = middle + 1;
      }
      if (low < end) starts.add(low);
      lineStart = low;
    }
    index = end;
  }
  return starts;
}

/** `?textLayoutVerify=1`: every probe is also read marker by marker and the
 * two answers compared (window.__mraiNativeLayoutVerify), for checking
 * {@link probeLineStarts} against a corpus in a real browser. */
let verifyProbeLines: boolean | undefined;
function probeLinesVerified(): boolean {
  if (verifyProbeLines === undefined) {
    try { verifyProbeLines = new URLSearchParams(window.location.search).get("textLayoutVerify") === "1"; } catch { verifyProbeLines = false; }
  }
  return verifyProbeLines;
}

/** The lines of a probe that is in the document and laid out. */
function readNativeTextProbe(parts: readonly NativeTextPart[], style: CSSProperties, dpi: number, atoms: readonly NativeTextAtom[]): NativeTextLine[] {
  const tops = new Map<number, number>();
  const top = (index: number) => {
    let y = tops.get(index);
    if (y === undefined) { y = atoms[index].marker.getBoundingClientRect().top; tops.set(index, y); }
    return y;
  };
  const starts = probeLineStarts(atoms, top);
  if (probeLinesVerified()) {
    const every = new Set<number>();
    let previous: number | undefined;
    atoms.forEach((atom, index) => {
      if (atom.hardBreak) { previous = undefined; return; }
      const y = top(index);
      if (previous !== undefined && Math.abs(y - previous) > .1) every.add(index);
      previous = y;
    });
    const record = ((window as unknown as Record<string, unknown>).__mraiNativeLayoutVerify ??= { checked: 0, mismatches: [] as unknown[] }) as { checked: number; mismatches: unknown[] };
    record.checked += 1;
    if (every.size !== starts.size || [...every].some((index) => !starts.has(index))) {
      if (record.mismatches.length < 200) record.mismatches.push({ text: parts.map((part) => part.text).join("").slice(0, 120), every: [...every], searched: [...starts] });
    }
  }
  const lines: NativeTextLine[]=[];
  let previousPart=-1;
  for(let index=0;index<atoms.length;index+=1){
    const atom=atoms[index];
    if(atom.hardBreak){
      if(!lines.length)lines.push({parts:[],hardBreak:""});
      const line=lines[lines.length-1];
      const part = parts[atom.part];
      const resolved = resolveSpreadsheetTextStyle({...style,...part.style}, nativeTextFontLead(part), dpi);
      const breakStyle = {...part.style, fontFamily: resolved.fontFamily, fontStyle: resolved.fontStyle};
      line.hardBreak+=atom.text;line.breakStyle=breakStyle;
      if(!line.parts.length)line.emptyStyle=breakStyle;
      lines.push({parts:[],hardBreak:"",emptyStyle:breakStyle});previousPart=-1;
      continue;
    }
    if(!lines.length || starts.has(index)) {
      lines.push({parts:[],hardBreak:""});previousPart=-1;
    }
    const line=lines[lines.length-1];
    if(previousPart===atom.part && line.parts.length)line.parts[line.parts.length-1].text+=atom.text;
    else line.parts.push({text:atom.text,fontLeadingText:nativeTextFontLead(parts[atom.part]),...(parts[atom.part].style?{style:parts[atom.part].style}:{})});
    previousPart=atom.part;
  }
  return lines;
}

function rememberNativeTextLayout(key: string, lines: readonly NativeTextLine[]): void {
  // A sheet shows a few thousand wrapped cells; a cache that empties below
  // that size laid every one of them out again on each visit.
  if(layouts.size>=NATIVE_TEXT_LAYOUT_CACHE_LIMIT)layouts.clear();
  layouts.set(key,lines);
}
const NATIVE_TEXT_LAYOUT_CACHE_LIMIT = 32_768;





const NATIVE_TEXT_PROBE_BATCH_ATOMS = 20_000;

/** Whether a text's lines are already known (or can never be laid out here),
 * so reading them costs no layout. */
export function nativeTextLayoutIsKnown(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number): boolean {
  const key = nativeTextLayoutKey(parts, style, width, dpi);
  return key === undefined || layouts.has(key);
}

export interface NativeTextLayoutRequest { parts: readonly NativeTextPart[]; style: CSSProperties; width: number; dpi: number }

/**
 * Lay out many texts under ONE forced layout. A probe put in the document,
 * read and removed costs a layout pass of the whole page each time (about
 * 2.4 ms on a 21,000-element workbench): a sheet of a thousand wrapped cells
 * took seconds. All the probes go in together, the page is laid out once, and
 * every probe is read from that one layout (24x faster, measured).
 */
export function layoutNativeTextBatch(requests: readonly NativeTextLayoutRequest[]): void {
  const pending = new Map<string, {request:NativeTextLayoutRequest;root:HTMLElement;atoms:NativeTextAtom[]}>();
  let atoms = 0;
  const flush = () => {
    if (pending.size === 0) return;
    const host = document.createElement("div");
    host.setAttribute("aria-hidden","true");
    for (const probe of pending.values()) host.appendChild(probe.root);
    try {
      layoutProbeParent(document).append(host);
      for (const [key, probe] of pending) {
        const { parts, style, width, dpi } = probe.request;
        const probed = readNativeTextProbe(parts, style, dpi, probe.atoms);
        if (probeLinesVerified()) verifyCanvasAdvanceLayout(parts, style, width, dpi, probed);
        rememberNativeTextLayout(key, (probeLinesVerified() && canvasNativeTextLines(parts, style, width, dpi)) || probed);
      }
    } finally {
      host.remove();
      for (const probe of pending.values()) releaseNativeTextProbe(probe.root);
      pending.clear(); atoms = 0;
    }
  };
  for (const request of requests) {
    const key = nativeTextLayoutKey(request.parts, request.style, request.width, request.dpi);
    if (key === undefined || layouts.has(key) || pending.has(key)) continue;
    const engine = engineNativeTextLayout(request.parts, request.style, request.width, request.dpi);
    if (engine) {
      rememberNativeTextLayout(key, engine);
      continue;
    }
    const canvas = probeLinesVerified() ? undefined : canvasNativeTextLines(request.parts, request.style, request.width, request.dpi);
    if (canvas) {
      rememberNativeTextLayout(key, canvas);
      continue;
    }
    const probe = buildNativeTextProbe(request.parts, request.style, request.width, request.dpi);
    pending.set(key, {request, ...probe});
    atoms += probe.atoms.length;
    if (atoms >= NATIVE_TEXT_PROBE_BATCH_ATOMS) flush();
  }
  flush();
}

/** Browser line breaking on native device advances, independent of paint.
 * Zero-height inline markers identify baselines across mixed font sizes without
 * adding a break opportunity or changing any glyph's measured geometry. */
export function layoutNativeText(
  parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number,
): readonly NativeTextLine[] | undefined {
  const key = nativeTextLayoutKey(parts, style, width, dpi);
  if (key === undefined) return undefined;
  if (layouts.has(key)) return layouts.get(key);
  const engine = engineNativeTextLayout(parts, style, width, dpi);
  if (engine) {
    rememberNativeTextLayout(key, engine);
    return engine;
  }
  const verified = probeLinesVerified();
  const canvas = verified ? undefined : canvasNativeTextLines(parts, style, width, dpi);
  if (canvas) {
    rememberNativeTextLayout(key, canvas);
    return canvas;
  }
  const {root,atoms} = buildNativeTextProbe(parts, style, width, dpi);
  try {
    layoutProbeParent(document).append(root);
    const probed = readNativeTextProbe(parts, style, dpi, atoms);
    if (verified) verifyCanvasAdvanceLayout(parts, style, width, dpi, probed);
    // Verified with the canvas lines on, the page reads those (the probe is
    // only the comparison).
    const lines = (verified && canvasNativeTextLines(parts, style, width, dpi)) || probed;
    rememberNativeTextLayout(key,lines);
    return lines;
  } finally {releaseNativeTextProbe(root);}
}

/**
 * A paragraph's lines and line boxes when they cost no layout of the page:
 * lines the cache already holds or the text engine breaks, each line box from
 * the native line height (a 72 dpi Mac grid) or from the engine's faces (CSS
 * `line-height: normal`). `top` and `height` place a line box in the
 * paragraph, px; `baseline` is its baseline from the paragraph's top, or null
 * where the line's own font puts it (half-leading in a box of `height`).
 * Undefined when either needs the browser, so the DOM paragraph lays it out.
 */
export function nativeTextParagraphWithoutLayout(
  parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number,
): { lines: readonly NativeTextLine[]; boxes: readonly { top: number; height: number; baseline: number | null }[] } | undefined {
  const key = nativeTextLayoutKey(parts, style, width, dpi);
  if (key === undefined) return undefined;
  let lines = layouts.get(key);
  if (!lines) {
    const engine = engineNativeTextLayout(parts, style, width, dpi);
    if (!engine) return undefined;
    rememberNativeTextLayout(key, engine);
    lines = engine;
  }
  const fixed = nativeParagraphLinePitch(parts, style, dpi);
  if (fixed !== undefined) return { lines, boxes: lines.map((_, index) => ({ top: index * fixed, height: fixed, baseline: null })) };
  const strut = fontCss(nativeParagraphFontStyle(style, 1, parts[0] ? nativeTextFontLead(parts[0]) : "", dpi));
  const runs = lines.map((line) => line.parts.length
    ? line.parts.map((part) => ({ text: part.text, fontCss: fontCss(resolveSpreadsheetTextStyle({ ...style, ...part.style }, nativeTextFontLead(part), dpi)) }))
    : [{ text: "", fontCss: fontCss(resolveSpreadsheetTextStyle({ ...style, ...line.emptyStyle }, "", dpi)) }]);
  // A face the engine has no file of: the line boxes from the browser's own
  // faces, for lines the browser already broke (the cache holds them).
  const measured = spreadsheetEngineLineBoxes(runs, strut) ?? canvasLineBoxes(runs, strut);
  if (!measured) return undefined;
  let top = 0;
  const boxes = measured.map((box) => {
    const placed = { top, height: box.top + box.bottom, baseline: top + box.top };
    top += placed.height;
    return placed;
  });
  return { lines, boxes };
}

/**
 * The lines {@link layoutNativeText} reads off a browser probe, computed
 * instead from the text engine's advances (the same device-rounded advances
 * the probe's spans carry) and its line breaker: UAX #14 opportunities, the
 * Office kinsoku list, spaces hanging at a break, an over-long word broken
 * between graphemes -- what the probe's `pre-wrap` + `overflow-wrap:
 * anywhere` asks of the browser, without the browser. Undefined when the
 * engine is off or cannot measure every run, so the probe decides as before.
 */
/**
 * `?textLayoutVerify=1`, read-only: the lines the engine's breaker makes of
 * the browser's own device advances (the ones the probe spaces its spans
 * by), for a text the engine declined (a face it has no file of, such as an
 * installed Calibri), against the probe's (window.__mraiNativeLayoutVerify
 * .canvas). Nothing is laid out with them.
 */
function verifyCanvasAdvanceLayout(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number, probed: readonly NativeTextLine[]): void {
  const canvas = engineNativeTextLayout(parts, style, width, dpi, true);
  const record = ((window as unknown as Record<string, unknown>).__mraiNativeLayoutVerify ??= { checked: 0, mismatches: [] as unknown[] }) as {
    canvas?: { checked: number; unavailable: number; mismatches: unknown[] };
  };
  const tally = (record.canvas ??= { checked: 0, unavailable: 0, mismatches: [] }) as {
    checked: number; unavailable: number; mismatches: unknown[]; transient?: number;
  };
  if (!canvas) { tally.unavailable += 1; return; }
  tally.checked += 1;
  const spell = (lines: readonly NativeTextLine[]) => lines.map((line) => line.parts.map((part) => part.text).join("") + (line.hardBreak ? "\u21b5" : ""));
  const differ = (a: readonly string[], b: readonly string[]) => a.length !== b.length || a.some((line, index) => line !== b[index]);
  const a = spell(probed), b = spell(canvas);
  if (!differ(a, b)) return;
  // A face that has just loaded is in the canvas's measures at once, and in the
  // page's layout only once the load has restyled it: a probe laid out in
  // between broke R83-X03's Century Gothic paragraphs 13% wider apart than
  
  // text is compared again once the page has taken its fonts in; only what
  // still differs is a mismatch.
  tally.transient ??= 0;
  window.setTimeout(() => {
    const again = engineNativeTextLayout(parts, style, width, dpi, true);
    const { root, atoms } = buildNativeTextProbe(parts, style, width, dpi);
    try {
      layoutProbeParent(document).append(root);
      const probe = spell(readNativeTextProbe(parts, style, dpi, atoms));
      const lines = again ? spell(again) : b;
      if (!differ(probe, lines)) { tally.transient! += 1; return; }
      if (tally.mismatches.length < 200) tally.mismatches.push({ width, probe: probe.map((line) => line.slice(-24)), canvas: lines.map((line) => line.slice(-24)) });
    } finally { releaseNativeTextProbe(root); }
  }, VERIFY_AGAIN_MS);
}

/** How long after a mismatch it is compared again (``verifyCanvasAdvanceLayout``). */
const VERIFY_AGAIN_MS = 1500;

/** A grapheme that draws something (a zero-width joiner or space need not). */
const VISIBLE_GRAPHEME = /[\p{L}\p{N}\p{P}\p{S}]/u;









function browserFontReady(font: string, text: string): boolean {
  try {
    const fonts = typeof document !== "undefined" ? document.fonts : undefined;
    if (!fonts) return false;
    const families = fontStackFamilies(font);
    const points = [...new Set([...text].map((char) => char.codePointAt(0) ?? 0))].filter((point) => point > 0x20);
    const { weight, italic } = fontShorthandTraits(font);
    const index = registeredFacesByFamily(fonts);
    // The faces of the stack's families that can draw this text, by family.
    const byFamily = new Map<string, Array<{ status: FontFaceLoadStatus; ranges: Array<[number, number]>; same: boolean }>>();
    for (const family of families) {
      for (const { face, ranges } of index.get(family) ?? []) {
        if (!points.some((point) => ranges.some(([from, to]) => point >= from && point <= to))) continue;
        // The status is read now: only the index of faces is kept.
        if (face.status === "loading") return false;
        const same = weightOf(face.weight) === weight && (face.style === "italic" || face.style === "oblique") === italic;
        (byFamily.get(family) ?? byFamily.set(family, []).get(family)!).push({ status: face.status, ranges, same });
      }
    }
    // Each character is drawn by the first family in the stack with a face
    // for it (a family the page registers no face of is an installed font).
    // Until a face of that family in the text's own weight and style has
    // loaded, the page draws it in a later family while the browser's
    // advances may already be the first's, or the other way about: R83-X03's
    // Century Gothic, its metric face not yet loaded, laid out 13% narrower
    
    for (const point of points) {
      for (const family of families) {
        const faces = (byFamily.get(family) ?? []).filter((face) => face.ranges.some(([from, to]) => point >= from && point <= to) && face.status !== "error");
        if (!faces.length) continue;
        const own = faces.filter((face) => face.same);
        if (!(own.length ? own : faces).some((face) => face.status === "loaded")) return false;
        break;
      }
    }
    return true;
  } catch {
    return false;
  }
}

/**
 * The page's registered faces by family, each with its unicode ranges read:
 * the faces a text could be drawn in, without walking every face the page
 * registers (thousands: Noto's CJK slices) and reading its ranges again for
 * each text (R116's open: about half a second of 4,000 texts). Read again
 * when the set changes size (a face added or deleted); a load starting or
 * ending changes no face's family or ranges, and what a face's load status is
 * stays read from the face itself. Each face's family and ranges are read
 * once, so a set read again reads only the faces added since.
 */
type RegisteredFace = { face: FontFace; ranges: Array<[number, number]> };
let registeredFaces: { fonts: FontFaceSet; size: number; byFamily: Map<string, RegisteredFace[]> } | undefined;
const faceReadings = new WeakMap<FontFace, { family: string; entry: RegisteredFace }>();
function registeredFacesByFamily(fonts: FontFaceSet): Map<string, RegisteredFace[]> {
  const size = (fonts as unknown as { size?: number }).size ?? -1;
  const known = registeredFaces;
  if (known && known.fonts === fonts && known.size === size && size >= 0) return known.byFamily;
  const byFamily = new Map<string, RegisteredFace[]>();
  for (const face of fonts as unknown as Iterable<FontFace>) {
    let reading = faceReadings.get(face);
    if (!reading) {
      reading = { family: unquotedFamily(face.family), entry: { face, ranges: unicodeRanges(face.unicodeRange) } };
      faceReadings.set(face, reading);
    }
    (byFamily.get(reading.family) ?? byFamily.set(reading.family, []).get(reading.family)!).push(reading.entry);
  }
  registeredFaces = { fonts, size, byFamily };
  return byFamily;
}

function weightOf(value: string | undefined): number {
  if (!value || value === "normal") return 400;
  if (value === "bold") return 700;
  const number = Number.parseFloat(value);
  return Number.isFinite(number) ? number : 400;
}

/** The weight and italic of a CSS font shorthand ("italic 700 15px A, B"). */
function fontShorthandTraits(font: string): { weight: number; italic: boolean } {
  const head = font.replace(/\d+(?:\.\d+)?(?:px|pt|em|rem|%)(?:\/\S+)?\s+.*$/, "");
  const weight = /\b(bold|[1-9]00)\b/.exec(head)?.[1];
  return { weight: weightOf(weight), italic: /\b(italic|oblique)\b/.test(head) };
}

function unquotedFamily(family: string): string {
  return family.trim().replace(/^["']|["']$/g, "").toLowerCase();
}

/** The families a CSS font shorthand names, lower case and unquoted. */
function fontStackFamilies(font: string): Set<string> {
  const stack = font.replace(/^.*?\d+(?:\.\d+)?(?:px|pt|em|rem|%)(?:\/\S+)?\s+/, "");
  return new Set(stack.split(",").map(unquotedFamily).filter(Boolean));
}

/** A face's unicode-range ("U+0-FF, U+131, U+4??") as inclusive code point pairs. */
function unicodeRanges(range: string | undefined): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (const part of (range || "U+0-10FFFF").split(",")) {
    const token = part.trim().replace(/^u\+/i, "");
    if (!token) continue;
    if (token.includes("?")) {
      out.push([parseInt(token.replace(/\?/g, "0"), 16), parseInt(token.replace(/\?/g, "F"), 16)]);
      continue;
    }
    const [from, to] = token.split("-");
    const start = parseInt(from, 16);
    const end = to === undefined ? start : parseInt(to, 16);
    if (Number.isFinite(start) && Number.isFinite(end)) out.push([start, end]);
  }
  return out;
}





let canvasNativeTextLayout: boolean | undefined;
function canvasNativeTextLayoutOn(): boolean {
  if (canvasNativeTextLayout === undefined) {
    let asked: string | null = null;
    try { asked = new URLSearchParams(window.location.search).get("textLayoutCanvas"); } catch { asked = null; }
    canvasNativeTextLayout = asked === "1" ? true : asked === "0" ? false : browserLinesStandForProbe();
  }
  return canvasNativeTextLayout;
}














function browserLinesStandForProbe(): boolean {
  try {
    const brands = (navigator as Navigator & { userAgentData?: { brands?: ReadonlyArray<{ brand?: string }> } }).userAgentData?.brands;
    if (Array.isArray(brands) && brands.some((entry) => entry?.brand === "Chromium")) return true;
    return webKitLayoutEngine(navigator.userAgent ?? "");
  } catch {
    return false;
  }
}

/** Safari ("Version/18.1 Safari/605.1.15", no Chrome token) and the iOS
 * browsers, all WebKit (CriOS, FxiOS, EdgiOS). */
export function webKitLayoutEngine(userAgent: string): boolean {
  if (/\b(CriOS|FxiOS|EdgiOS)\//.test(userAgent)) return true;
  return /AppleWebKit\/\d/.test(userAgent) && /\bVersion\/\d[^ ]* (Mobile\/\S+ )?Safari\//.test(userAgent) && !/\b(Chrome|Chromium|Edg|OPR)\//.test(userAgent);
}

/** The lines on the browser's advances, when they may stand for the probe's. */
function canvasNativeTextLines(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number): NativeTextLine[] | undefined {
  if (!canvasNativeTextLayoutOn()) return undefined;
  canvasLayoutTally().attempts += 1;
  if (!spreadsheetTextEngineOn()) return canvasFallback("engineOff", "", "");
  const lines = engineNativeTextLayout(parts, style, width, dpi, true);
  if (lines) canvasLayoutTally().used += 1;
  return lines;
}




interface CanvasLayoutTally {
  attempts: number;
  used: number;
  fallbacks: Record<string, number>;
  samples: Array<{ reason: string; font: string; text: string; unloaded?: string[] }>;
}
function canvasLayoutTally(): CanvasLayoutTally {
  const host = window as unknown as Record<string, unknown>;
  return (host.__mraiCanvasTextLayout ??= { attempts: 0, used: 0, fallbacks: {}, samples: [] }) as CanvasLayoutTally;
}
function canvasFallback(reason: string, font: string, text: string): undefined {
  const tally = canvasLayoutTally();
  tally.fallbacks[reason] = (tally.fallbacks[reason] ?? 0) + 1;
  if (tally.samples.length < 40) {
    let unloaded: string[] | undefined;
    if (reason === "fontNotReady") {
      // The faces it waits for: those of the stack's families still loading.
      try {
        const families = fontStackFamilies(font);
        unloaded = [...(document.fonts as unknown as Iterable<FontFace>)]
          .filter((face) => face.status === "loading" && families.has(unquotedFamily(face.family)))
          .slice(0, 8)
          .map((face) => `${face.family} ${face.status} ${face.unicodeRange}`.slice(0, 160));
      } catch {
        unloaded = undefined;
      }
    }
    tally.samples.push({ reason, font: font.slice(0, 200), text: text.slice(0, 40), ...(unloaded ? { unloaded } : {}) });
  }
  return undefined;
}

function engineNativeTextLayout(parts: readonly NativeTextPart[], style: CSSProperties, width: number, dpi: number, browserAdvances = false): NativeTextLine[] | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  // The printer's re-wrap (NATIVE_PRINTED_ROOM_MARK on the print page) breaks
  // at the advances the line is painted at, the printer's whole dots at its
  // pixels per em, as Windows Excel's printer and the API's writer
  // (FaceMeasure.printed) break it: the screen's whole pixels ran Arial 11
  // 2.6 % narrow and R219-X03's notes took a word more a line than Excel's
  // (565 pt of printed advances in a 555 pt room measured 550).
  const printedRoom = (style as unknown as Record<string, unknown>)[NATIVE_PRINTED_ROOM_MARK] === "1" && nativeTextPrinterDpi() > 0;
  width = printedRoom ? width : excelWrapRoom(width, style, dpi);
  // Probes compare the grid's own decisions across browsers through this.
  (window as unknown as Record<string, unknown>).__mraiSpreadsheetTextLayout ??= { layoutNativeText, measureNativeParagraphLineHeight, measureNativeLineAdvance, measureNativeTextHeight, engineParagraphHeight };
  type Atom = { text: string; part: number; advance: number; hardBreak?: false } | { text: string; part: number; hardBreak: true };
  const atoms: Atom[] = [];
  for (let partIndex = 0; partIndex < parts.length; partIndex += 1) {
    const part = parts[partIndex];
    const resolved = resolveSpreadsheetTextStyle({ ...style, ...part.style }, nativeTextFontLead(part), dpi);
    const font = fontCss(resolved);
    for (const piece of part.text.split(/(\r\n|\r|\n)/)) {
      if (!piece) continue;
      if (/^(\r\n|\r|\n)$/.test(piece)) {
        const previous = atoms.at(-1);
        const nextPart = parts[partIndex + 1];
        const crBeforeLf = piece === "\r" && part.text.endsWith("\r") && nextPart?.text.startsWith("\n");
        if (piece === "\r" && !crBeforeLf) continue;
        if (piece === "\n" && previous?.hardBreak && previous.text === "\r") { previous.text += "\n"; continue; }
        atoms.push({ text: piece, part: partIndex, hardBreak: true });
        continue;
      }
      const segments = graphemesOf(piece);
      if (!browserAdvances && !spreadsheetEngineIdealAdvances(segments, font)) return undefined;
      // The browser's own advances are trusted only for a face it holds: a
      // web font still loading measures as its fallback (the probe then
      // lays the text out, and the load lays it out again).
      if (browserAdvances && !browserFontReady(font, piece)) return canvasFallback("fontNotReady", font, piece);
      const advances = nativeTextAdvances(piece, font, dpi, 1, printedRoom ? "paint" : "measure");
      if (advances.length !== segments.length || advances.some((advance) => advance.advance === undefined)) {
        return browserAdvances ? canvasFallback("noAdvances", font, piece) : undefined;
      }
      // A visible character measured as nothing is a face the browser could
      // not measure: the probe lays it out.
      if (browserAdvances && advances.some((advance) => !(advance.advance! > 0) && VISIBLE_GRAPHEME.test(advance.text))) {
        return canvasFallback("zeroAdvance", font, piece);
      }
      for (const advance of advances) atoms.push({ text: advance.text, part: partIndex, advance: advance.advance! });
    }
  }
  const lines: NativeTextLine[] = [];
  let previousPart = -1;
  const addText = (atom: Atom, newLine: boolean) => {
    if (!lines.length || newLine) {
      lines.push({ parts: [], hardBreak: "" });
      previousPart = -1;
    }
    const line = lines[lines.length - 1];
    if (previousPart === atom.part && line.parts.length) line.parts[line.parts.length - 1].text += atom.text;
    else line.parts.push({ text: atom.text, fontLeadingText: nativeTextFontLead(parts[atom.part]), ...(parts[atom.part].style ? { style: parts[atom.part].style } : {}) });
    previousPart = atom.part;
  };
  let run: Extract<Atom, { advance: number }>[] = [];
  const flush = () => {
    if (!run.length) return;
    // Where Windows Excel breaks a cell's line, on the engine's advances and
    
    // before, the engine broke by UAX #14 and the zh-CN kinsoku list, and the
    // browser's advances by its own ASCII table (after "?", "–", "…"; inside
    // a run of ideographs; never before "。").
    const opportunities = lineBreakOpportunities(run.map((atom) => atom.text), { kinsoku: null, excel: true });
    const starts = new Set(breakLines(run, width, opportunities, (style as unknown as Record<string, unknown>)[NATIVE_KEEP_SPACE_MARK] === "1"));
    run.forEach((atom, index) => addText(atom, index > 0 && starts.has(index)));
    run = [];
  };
  for (const atom of atoms) {
    if (!atom.hardBreak) {
      run.push(atom);
      continue;
    }
    flush();
    if (!lines.length) lines.push({ parts: [], hardBreak: "" });
    const line = lines[lines.length - 1];
    const part = parts[atom.part];
    const resolved = resolveSpreadsheetTextStyle({ ...style, ...part.style }, nativeTextFontLead(part), dpi);
    const breakStyle = { ...part.style, fontFamily: resolved.fontFamily, fontStyle: resolved.fontStyle };
    line.hardBreak += atom.text;
    line.breakStyle = breakStyle;
    if (!line.parts.length) line.emptyStyle = breakStyle;
    lines.push({ parts: [], hardBreak: "", emptyStyle: breakStyle });
    previousPart = -1;
  }
  flush();
  return lines;
}
