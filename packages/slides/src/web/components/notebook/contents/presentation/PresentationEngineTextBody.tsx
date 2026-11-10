"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A slide's text body drawn line by line where the text engine lays it out
 * (`lib/text-engine/office-text-body.ts`): every line ends where PowerPoint
 * ends it, stands on PowerPoint's baseline and spreads its words as
 * PowerPoint does, from the served font files' own advances. The browser
 * only paints the glyphs: each stretch of a run on a line is placed at the
 * engine's position, with no line breaking, kerning or ligatures of its own.
 *
 * The rows are read exactly as the browser-flowed drawing reads them
 * (`planPresentationParagraphRows`); a body the engine cannot lay out yet
 * (its faces still loading) or at all (an equation, a right-to-left row, a
 * face the page does not serve) is drawn the browser's way, so a body is
 * never drawn worse than before.
 *
 * Where it draws (`PresentationTextEngineSurfaceContext`): the deck's printed
 * pages and the slide thumbnails first (phase 2, step 1 of the text engine
 * plan, `lib/text-engine/PLAN-phase2.md`); the canvas follows with its caret.
 */

import { createContext, useContext, useEffect, useMemo, useState, type CSSProperties, type ReactElement } from "react";
import {
  layoutOfficeTextBody,
  type OfficeParagraphInput,
  type OfficeTextAlign,
  type OfficeTextBodyLayout,
  type OfficeTextFont,
  type OfficeTextRunInput,
} from "@mrnote/core/web/lib/text-engine/office-text-body.ts";
import { pageOfficeTextMeasurer } from "@mrnote/core/web/lib/text-engine/office-text-measurer.ts";
import { OFFICE_TEXT_HIT_ATTRIBUTE, registerOfficeTextHitMap } from "@mrnote/core/web/lib/text-engine/office-text-hit.ts";
import { OFFICE_TEXT_MARKER_ATTRIBUTE } from "@mrnote/core/web/lib/presentation/math-atom.ts";
import { onTextEngineChange, textEngineEpoch } from "@mrnote/core/web/lib/text-engine/measure.ts";
import { officeUnderline } from "@mrnote/core/web/lib/text-engine/office-underline.ts";
import { textEngineOn } from "@mrnote/core/web/lib/text-engine/page-adapter.ts";
import type { TextRun } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import { normalizePresentationTabStops } from "../../../../lib/presentation-tab-stops";
import { PRESENTATION_SCRIPT_SCALE, runStyleCss, type PresentationFontSizeUnit } from "../../../../lib/presentation/text-layout-css";
import { POWERPOINT_MISSING_SCRIPT_FAMILY, type PresentationFontAliases } from "@mrnote/core/web/lib/presentation-font-faces.ts";
import { presentationDisplayRuns, PresentationSlideNumberContext } from "./PresentationRunsText";
import type { PresentationParagraphRowPlan } from "./presentationParagraphRows";
import { PresentationFontAliasesContext } from "./usePresentationFontFaces";

/** Whether the bodies drawn under it may be drawn by the text engine. */
export const PresentationTextEngineSurfaceContext = createContext(false);

/** The attribute a body still waiting for its faces carries (the print waits on it). */
export const PRESENTATION_TEXT_ENGINE_ATTRIBUTE = "data-text-engine";

const PX_PER_PT = 4 / 3;
const PT_PER_PX = 0.75;

/** A CSS length in px: a number, `Npx`, or `Nem` of `emPx`; null for anything else. */
function cssPx(value: unknown, emPx: number): number | null {
  if (value === undefined || value === null || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string") return null;
  const match = /^\s*(-?[0-9]*\.?[0-9]+(?:e-?[0-9]+)?)(px|em)?\s*$/i.exec(value);
  if (!match) return null;
  const number = Number(match[1]);
  if (!Number.isFinite(number)) return null;
  return match[2]?.toLowerCase() === "em" ? number * emPx : number;
}

function weightOf(value: unknown): number | null {
  if (value === undefined || value === null || value === "normal") return 400;
  if (value === "bold") return 700;
  const number = typeof value === "number" ? value : Number(value);
  return Number.isFinite(number) ? number : null;
}

function alignOf(textAlign: unknown, textAlignLast: unknown): OfficeTextAlign | null {
  if (textAlignLast === "justify") return "dist";
  switch (textAlign) {
    case undefined:
    case null:
    case "left":
    case "start":
      return "l";
    case "center":
      return "ctr";
    case "right":
    case "end":
      return "r";
    case "justify":
      return "just";
    default:
      return null;
  }
}



















const scriptShiftPt = new WeakMap<object, number>();
function scriptRunCss(run: TextRun, css: CSSProperties, rowFontPx: number): CSSProperties {
  const style = run.style as { superscript?: boolean; subscript?: boolean } | undefined;
  if (!style?.superscript && !style?.subscript) return css;
  const drawn = cssPx(css.fontSize, rowFontPx);
  if (drawn === null || !(drawn > 0)) return css;
  const basePt = (drawn / PRESENTATION_SCRIPT_SCALE) * PT_PER_PX;
  const scriptPt = Math.max(1, Math.round((basePt * 2) / 3));
  scriptShiftPt.set(run, (style.superscript ? 0.3 : -0.25) * basePt);
  return { ...css, fontSize: scriptPt * PX_PER_PT, verticalAlign: undefined };
}

const APP_SYSTEM_STACK = /^\s*(?:system-ui|-apple-system|blinkmacsystemfont)\b/i;
const APP_SYSTEM_NAME = /^["']?\s*(?:system-ui|-apple-system|blinkmacsystemfont)\s*["']?$/i;
function officeFaceOr(stack: unknown, bodyFace: string | undefined): unknown {
  if (typeof stack !== "string" || !APP_SYSTEM_STACK.test(stack)) return stack;
  if (bodyFace && !APP_SYSTEM_STACK.test(bodyFace)) return bodyFace;
  // A run whose own face is a CSS system keyword (`a:latin
  // typeface="-apple-system"`): no font PowerPoint can have, so it draws the
  // run as it draws any face it lacks, in the stand-in the stack carries after
  
  // in PowerPoint's PDF). Only that stack: the app's own system default stays.
  const families = stack.split(",").map((family) => family.trim());
  if (!families.some((family) => family.replace(/^["']|["']$/g, "") === POWERPOINT_MISSING_SCRIPT_FAMILY)) return stack;
  while (families.length && APP_SYSTEM_NAME.test(families[0])) families.shift();
  return families.join(", ");
}

/** A body the engine leaves to the browser: thrown as a plain value (no stack taken), caught below. */
class Unsupported {
  constructor(readonly message: string) {}
}

/** A length the page wrote from points (rounded to hundredths of a px), read back as the hundredths of a point it was. */
function hundredthPt(px: number): number {
  return Math.round(px * PT_PER_PX * 100) / 100;
}

function need<T>(value: T | null | undefined, why: string): T {
  if (value === null || value === undefined) throw new Unsupported(why);
  return value;
}

interface EngineBodyInput {
  readonly paragraphs: OfficeParagraphInput[];
  /** Each row's runs as drawn (fields shown as the slide shows them). */
  readonly rowRuns: TextRun[][];
}

/** Everything the CSS cascade gives a run: its own style over its row's over the body's. */
function cascade(own: CSSProperties, row: CSSProperties, body: CSSProperties, key: keyof CSSProperties): unknown {
  return own[key] ?? row[key] ?? body[key];
}

/**
 * The engine's reading of a planned body, or why it cannot read it. Px are
 * the slide's design pixels (96 per inch); the engine works in points.
 */
export function presentationEngineBodyInput(args: {
  rows: readonly PresentationParagraphRowPlan[];
  bodyStyle: CSSProperties;
  fontSizeUnit: PresentationFontSizeUnit;
  fontAliases: PresentationFontAliases;
  autofitScale: number;
  slideNumber: number | undefined;
  tabStops: unknown;
  /** The default tab interval (`@defTabSz`), design px; unstated, an inch. */
  defaultTabPx?: unknown;
  /** `source.nativeKerning`: the size (pt) from which PowerPoint kerns each row's pairs. */
  kerning?: unknown;
  now?: Date;
}): EngineBodyInput | { unsupported: string } {
  const { rows, bodyStyle, fontSizeUnit, fontAliases, autofitScale } = args;
  try {
    const tabStops = normalizePresentationTabStops(args.tabStops);
    if (tabStops.some((stop) => stop.align !== "left")) throw new Unsupported("aligned tab stop");
    const paragraphs: OfficeParagraphInput[] = [];
    const rowRuns: TextRun[][] = [];
    const kerningOf = (index: number): number | null => {
      const value = Array.isArray(args.kerning) ? args.kerning[index] : args.kerning;
      return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : null;
    };
    // A row with no words of its own (an empty paragraph, a marker's face
    // nobody states) takes the face the body's words are in, as its end
    // mark does in PowerPoint (the body's first stated face).
    const bodyFace = (() => {
      for (const row of rows) {
        const stated = row.rowTextStyle.fontFamily ?? row.runs?.map((run) => runStyleCss(run.style, fontSizeUnit, fontAliases, autofitScale).fontFamily).find(Boolean);
        if (typeof stated === "string" && stated.trim()) return stated;
      }
      return typeof bodyStyle.fontFamily === "string" ? bodyStyle.fontFamily : undefined;
    })();
    for (const row of rows) {
      if (row.dir === "rtl") throw new Unsupported("right-to-left row");
      const kernFrom = kerningOf(row.index);
      const rowStyle = row.rowTextStyle;
      const bodyFontPx = need(cssPx(bodyStyle.fontSize ?? 24, 16), "body size");
      const rowFontPx = need(cssPx(rowStyle.fontSize ?? bodyStyle.fontSize ?? 24, bodyFontPx), "row size");
      const fontOf = (css: CSSProperties): OfficeTextFont => {
        const stack = officeFaceOr(cascade(css, rowStyle, bodyStyle, "fontFamily") ?? bodyFace, bodyFace);
        if (typeof stack !== "string" || !stack.trim()) throw new Unsupported("no face");
        if (cascade(css, rowStyle, bodyStyle, "fontVariant") === "small-caps") throw new Unsupported("small capitals");
        const size = need(cssPx(css.fontSize ?? rowStyle.fontSize ?? bodyStyle.fontSize ?? 24, rowFontPx), "run size");
        // DrawingML sizes are hundredths of a point; the CSS px they were
        // written as are rounded to hundredths of a pixel (10 pt -> 13.33 px
        // -> 9.9975 pt), and that shift moves glyph advances across the
        
        const sizePt = Math.round(size * PT_PER_PX * 100) / 100;
        return {
          stack,
          weight: need(weightOf(cascade(css, rowStyle, bodyStyle, "fontWeight")), "weight"),
          italic: cascade(css, rowStyle, bodyStyle, "fontStyle") === "italic",
          sizePt,
          // PowerPoint kerns a run's pairs from the row's `kern` size up.
          ...(kernFrom !== null && sizePt >= kernFrom - 1e-6 ? { kerning: true } : {}),
        };
      };
      const displayed = presentationDisplayRuns(row.runs ?? undefined, args.slideNumber, args.now ?? new Date()) ??
        (row.text ? [{ text: row.text }] : []);
      const runs: OfficeTextRunInput[] = displayed.map((run) => {
        if (run.math) throw new Unsupported("equation");
        // WordArt on a run (gradient fills, outlines, effects) and symbol-font
        // characters are the browser drawing's.
        if (run.style?.textPaint) throw new Unsupported("run WordArt");
        if (/[\uf000-\uf0ff]/.test(run.text)) throw new Unsupported("symbol-font characters");
        const css = scriptRunCss(run, row.runs ? runStyleCss(run.style, fontSizeUnit, fontAliases, autofitScale) : {}, rowFontPx);
        const font = fontOf(css);
        const spacing = cascade(css, rowStyle, bodyStyle, "letterSpacing");
        const spacingPx = spacing === "normal" ? 0 : need(cssPx(spacing, font.sizePt * PX_PER_PT), "letter spacing");
        return {
          text: run.text,
          font,
          ...(spacingPx ? { letterSpacingPt: hundredthPt(spacingPx) } : {}),
          ...(cascade(css, rowStyle, bodyStyle, "textTransform") === "uppercase" ? { allCaps: true } : {}),
        };
      });
      rowRuns.push(displayed);
      const endFont = fontOf({});
      const align = need(alignOf(rowStyle.textAlign ?? bodyStyle.textAlign, rowStyle.textAlignLast ?? bodyStyle.textAlignLast), "alignment");
      // Line spacing: a unitless CSS line height is the percentage times
      // Office's single line (its autofit reduction already taken off); a
      // length is Line Spacing > Exactly.
      const lineHeight = rowStyle.lineHeight ?? bodyStyle.lineHeight ?? 1.2;
      const lineSpacing = typeof lineHeight === "number"
        ? { kind: "percent" as const, value: lineHeight / 1.2 }
        : { kind: "points" as const, value: hundredthPt(need(cssPx(lineHeight, rowFontPx), "line height")) };
      // The row's own spacing over its style's, as the drawn row stacks them:
      // a preserved diagram's label states its drawing's (`paragraphMetrics`,
      // none before its first row or after its last).
      const before = need(cssPx(row.indentCss.paddingTop ?? rowStyle.paddingTop, rowFontPx), "space before");
      const after = need(cssPx(row.indentCss.paddingBottom ?? rowStyle.paddingBottom, rowFontPx), "space after");
      // The row's own indentation over its gutter, as the row's flex box reads it.
      const indent = { ...row.gutterCss, ...pickIndent(rowStyle), ...row.indentCss };
      const start = need(cssPx(indent.paddingInlineStart, rowFontPx), "indent");
      const end = need(cssPx(indent.paddingInlineEnd, rowFontPx), "right indent");
      const textIndent = need(cssPx(indent.textIndent, rowFontPx), "first line");
      let marker: OfficeParagraphInput["marker"];
      if (row.markerText !== undefined || row.markerPicture) {
        if (row.markerFollowsText) throw new Unsupported("marker that moves with its text");
        const markerSizePx = need(cssPx(row.markerCss.fontSize ?? rowFontPx, rowFontPx), "marker size");
        const markerFont: OfficeTextFont = {
          stack: String(officeFaceOr(row.markerCss.fontFamily ?? rowStyle.fontFamily ?? bodyStyle.fontFamily ?? bodyFace, bodyFace) ?? ""),
          weight: need(weightOf(row.markerCss.fontWeight ?? rowStyle.fontWeight ?? bodyStyle.fontWeight), "marker weight"),
          italic: (row.markerCss.fontStyle ?? rowStyle.fontStyle ?? bodyStyle.fontStyle) === "italic",
          sizePt: Math.round(markerSizePx * PT_PER_PX * 100) / 100,
        };
        if (!markerFont.stack) throw new Unsupported("marker face");
        const box = row.markerBoxCss;
        const hanging = box.width !== undefined;
        const offsetPx = need(cssPx(box.marginInlineStart, markerSizePx), "marker offset");
        const boxPx = hanging ? need(cssPx(box.width, markerSizePx), "marker width") : 1.4 * markerSizePx;
        const gapPx = need(cssPx(indent.gap, rowFontPx), "marker gap");
        marker = {
          text: row.markerText ?? "",
          font: markerFont,
          ...(row.markerSymbol ? { symbol: row.markerSymbol } : {}),
          // A picture marker is as wide as it is tall: the text's height.
          ...(row.markerPicture ? { widthPt: markerSizePx * PT_PER_PX } : {}),
          leftPt: Math.max(0, start + offsetPx) * PT_PER_PX,
          boxPt: boxPx * PT_PER_PX,
          gapPt: gapPx * PT_PER_PX,
          hanging,
        };
      }
      paragraphs.push({
        runs,
        endFont,
        align,
        marginLeftPt: start * PT_PER_PX,
        marginRightPt: end * PT_PER_PX,
        indentPt: textIndent * PT_PER_PX,
        lineSpacing,
        spaceBeforePt: hundredthPt(before),
        spaceAfterPt: hundredthPt(after),
        ...(marker ? { marker } : {}),
        tabStops: tabStops.map((stop) => ({ positionPt: stop.pos * PT_PER_PX, align: "l" as const })),
        defaultTabPt:
          typeof args.defaultTabPx === "number" && args.defaultTabPx > 0 ? args.defaultTabPx * PT_PER_PX : 72,
        kinsoku: "zh-CN",
      });
    }
    return { paragraphs, rowRuns };
  } catch (error) {
    if (error instanceof Unsupported) return { unsupported: error.message };
    throw error;
  }
}

function pickIndent(style: CSSProperties): CSSProperties {
  const out: CSSProperties = {};
  if (style.paddingInlineStart !== undefined) out.paddingInlineStart = style.paddingInlineStart;
  if (style.paddingInlineEnd !== undefined) out.paddingInlineEnd = style.paddingInlineEnd;
  if (style.textIndent !== undefined) out.textIndent = style.textIndent;
  if (style.gap !== undefined) out.gap = style.gap;
  return out;
}

/** Re-render when the engine's faces land; the epoch keys every layout. */
function useTextEngineEpoch(active: boolean): string {
  const [epoch, setEpoch] = useState(() => (active ? textEngineEpoch() : ""));
  useEffect(() => {
    if (!active) return;
    const update = () => setEpoch(textEngineEpoch());
    update();
    return onTextEngineChange(update);
  }, [active]);
  return epoch;
}

export type PresentationEngineBodyAnswer =
  | { readonly status: "ok"; readonly layout: Extract<OfficeTextBodyLayout, { status: "ok" }>; readonly input: EngineBodyInput }
  | { readonly status: "pending" }
  | { readonly status: "off" };

/**
 * The engine's layout of a body for this surface: `ok` to draw it line by
 * line, `pending` while its faces load (drawn the browser's way meanwhile,
 * the print waiting), `off` where the browser draws it.
 */
export function usePresentationEngineBody(args: {
  enabled: boolean;
  rows: readonly PresentationParagraphRowPlan[] | null;
  bodyStyle: CSSProperties;
  fontSizeUnit: PresentationFontSizeUnit;
  fontAliases: PresentationFontAliases;
  autofitScale: number;
  tabStops: unknown;
  defaultTabPx?: unknown;
  kerning?: unknown;
  widthPx: number | null;
  wrap: boolean;
  /** Lay out whatever the surface: the lines a body being typed in should end at. */
  force?: boolean;
}): PresentationEngineBodyAnswer {
  const surface = useContext(PresentationTextEngineSurfaceContext) || args.force === true;
  const slideNumber = useContext(PresentationSlideNumberContext);
  const active = args.enabled && surface && args.rows !== null && args.widthPx !== null && textEngineOn("slides");
  const epoch = useTextEngineEpoch(active);
  const { rows, bodyStyle, fontSizeUnit, fontAliases, autofitScale, tabStops, defaultTabPx, kerning, widthPx, wrap } = args;
  return useMemo((): PresentationEngineBodyAnswer => {
    void epoch;
    if (!surface) return { status: "off" };
    if (!active || !rows || widthPx === null) return noteEngine("off", !args.enabled ? "not a candidate" : !rows ? "no rows" : widthPx === null ? "no width" : "switched off");
    const input = presentationEngineBodyInput({ rows, bodyStyle, fontSizeUnit, fontAliases, autofitScale, slideNumber, tabStops, defaultTabPx, kerning });
    if ("unsupported" in input) return noteEngine("off", input.unsupported);
    const layout = layoutOfficeTextBody({ widthPt: widthPx * PT_PER_PX, wrap, paragraphs: input.paragraphs }, pageOfficeTextMeasurer);
    if (layout.status === "pending") return noteEngine("pending", "pending");
    if (layout.status === "unsupported") return noteEngine("off", layout.reason);
    return noteEngine("ok", "laid", { status: "ok", layout, input });
  }, [surface, active, args.enabled, epoch, rows, bodyStyle, fontSizeUnit, fontAliases, autofitScale, slideNumber, tabStops, defaultTabPx, kerning, widthPx, wrap]);
}

/**
 * Where each row's lines after its first start, as offsets in the row's text
 * (`ENGINE_LINE_STARTS_ATTRIBUTE`): the browser's rows of a body being typed
 * in are fitted to end their lines there (`powerpoint-line-fit-dom`).
 */
export function presentationEngineLineStarts(answer: Extract<PresentationEngineBodyAnswer, { status: "ok" }>): string[] {
  return answer.layout.paragraphs.map((paragraph, rowIndex) => {
    const runs = answer.input.rowRuns[rowIndex] ?? [];
    const runStarts: number[] = [];
    let total = 0;
    for (const run of runs) {
      runStarts.push(total);
      total += run.text.length;
    }
    return paragraph.lines
      .slice(1)
      .map((line) => line.segments[0])
      .filter((segment): segment is NonNullable<typeof segment> => Boolean(segment))
      .map((segment) => (runStarts[segment.run] ?? 0) + segment.start)
      .join(",");
  });
}

/**
 * What the engine answered for the bodies of this page, by reason
 * (`window.__mraiSlidesEngine`): the probes read it to see which bodies are
 * still drawn the browser's way, and why.
 */
function noteEngine(status: "ok" | "pending" | "off", reason: string, answer?: PresentationEngineBodyAnswer): PresentationEngineBodyAnswer {
  if (typeof window !== "undefined") {
    const tally = ((window as unknown as { __mraiSlidesEngine?: Record<string, number> }).__mraiSlidesEngine ??= {});
    const key = `${status}: ${reason.slice(0, 80)}`;
    tally[key] = (tally[key] ?? 0) + 1;
  }
  return answer ?? ({ status } as PresentationEngineBodyAnswer);
}

/**
 * A run's look as SVG text paints it, from the CSS the browser-flowed drawing
 * gives it: its colour as the fill, its face, size, weight, slant and lines.
 * Its spacing is not: every grapheme stands at the engine's own position.
 */
function svgTextPaint(css: CSSProperties, colorOverride: string | undefined): CSSProperties {
  const out: CSSProperties = {
    fontFamily: css.fontFamily,
    fontSize: css.fontSize,
    fontWeight: css.fontWeight,
    fontStyle: css.fontStyle,
    textDecoration: css.textDecoration,
    textDecorationStyle: css.textDecorationStyle,
  };
  const color = colorOverride ?? css.color;
  if (color !== undefined) out.fill = color as CSSProperties["fill"];
  if (css.WebkitTextStroke) out.WebkitTextStroke = css.WebkitTextStroke;
  return out;
}

const SVG_TEXT_STYLE: CSSProperties = {
  whiteSpace: "pre",
  fontKerning: "none",
  fontVariantLigatures: "none",
  letterSpacing: 0,
  wordSpacing: 0,
};

/** Positions in CSS px, at most 1/1000 px apart from the engine's. */
function px(points: number): number {
  return Math.round(points * PX_PER_PT * 1000) / 1000;
}

/**
 * The body drawn from an `ok` layout, as SVG text: each run's stretch of
 * each line a `tspan` whose every grapheme stands at the engine's position
 * on the line's baseline -- PowerPoint's 1/8 pt advances, its kerning and
 * its justification, whatever advances the platform's text stack would
 * have used (Chromium on Linux rounds each to a whole pixel: up to 0.37 pt
 * a glyph). The marker of a row stands on its first line.
 */
export function PresentationEngineTextBody({
  answer,
  rows,
  fontSizeUnit,
  fontAliases,
  autofitScale,
  colorOverride,
  overlay = false,
}: {
  answer: Extract<PresentationEngineBodyAnswer, { status: "ok" }>;
  rows: readonly PresentationParagraphRowPlan[];
  fontSizeUnit: PresentationFontSizeUnit;
  fontAliases: PresentationFontAliases;
  autofitScale: number;
  colorOverride?: string;
  /** Drawn over a body's editing rows (`PresentationEngineEditLayer`): the
   * rows are the text the readers read, so this drawing carries none of a
   * row's reading marks, takes no caret and is no part of the body's text. */
  overlay?: boolean;
}): ReactElement {
  const { layout, input } = answer;
  const nodes: ReactElement[] = [];
  layout.paragraphs.forEach((paragraph, rowIndex) => {
    const row = rows[rowIndex];
    const runs = input.rowRuns[rowIndex] ?? [];
    const runInputs = input.paragraphs[rowIndex]?.runs ?? [];
    const rowStyle = row?.rowTextStyle ?? {};
    const rowPaint = svgTextPaint(rowStyle, colorOverride ?? (rowStyle.color as string | undefined));
    const firstLine = paragraph.lines[0];
    // The row as the text readers read a drawn row (`presentationPreviewDomPoint`,
    // find highlights): its marker apart, its words in order in one group.
    const markerNodes: ReactElement[] = [];
    const lineNodes: ReactElement[] = [];
    if (paragraph.marker && row && firstLine) {
      const markerPaint = svgTextPaint({ ...rowStyle, ...row.markerCss }, colorOverride ?? (row.markerCss.color as string | undefined));
      if (row.markerPicture) {
        const size = runInputs[0]?.font.sizePt ?? paragraph.lines[0].sizePt;
        markerNodes.push(
          <image
            key={`m${rowIndex}`}
            href={row.markerPicture}
            x={px(paragraph.marker.x)}
            y={px(firstLine.baseline - size * 0.88)}
            width={px(size)}
            height={px(size)}
          />,
        );
      } else if (row.markerText) {
        markerNodes.push(
          <text key={`m${rowIndex}`} aria-hidden="true" data-presentation-selection-ignore="true" data-testid="presentation-paragraph-marker" x={px(paragraph.marker.x)} y={px(firstLine.baseline)} style={{ ...SVG_TEXT_STYLE, ...markerPaint }}>
            {row.markerText}
          </text>,
        );
      }
    }
    paragraph.lines.forEach((line, lineIndex) => {
      const spans: ReactElement[] = [];
      const underlines: { x0: number; x1: number; top: number; height: number; color: string }[] = [];
      line.segments.forEach((segment, segmentIndex) => {
        if (segment.tab) return;
        const run = runs[segment.run];
        if (!run) return;
        const caps = runInputs[segment.run]?.allCaps === true;
        const css = row?.runs ? runStyleCss(run.style, fontSizeUnit, fontAliases, autofitScale) : {};
        // Drawn in exactly the face, size, weight and slant the engine measured.
        const measured = runInputs[segment.run]?.font;
        const paint: CSSProperties = {
          ...svgTextPaint(css, colorOverride),
          ...(scriptShiftPt.has(run) ? { baselineShift: px(scriptShiftPt.get(run)!) } : {}),
          ...(measured
            ? { fontFamily: measured.stack, fontSize: px(measured.sizePt), fontWeight: measured.weight, fontStyle: measured.italic ? "italic" : "normal" }
            : {}),
        };
        // A single underline in whole points from the Office face's own
        // metrics, as PowerPoint draws it (lib/text-engine/office-underline.ts).
        const decoration = typeof paint.textDecoration === "string" ? paint.textDecoration : "";
        const underline =
          measured && /\bunderline\b/.test(decoration) && (paint.textDecorationStyle === undefined || paint.textDecorationStyle === "solid")
            ? officeUnderline(measured.stack, measured.weight, measured.italic, measured.sizePt)
            : null;
        if (underline) {
          paint.textDecoration = decoration.replace(/\bunderline\b/, "").trim() || "none";
          const color = String(paint.fill ?? rowPaint.fill ?? "currentColor");
          const top = line.baseline + underline.topPt - (scriptShiftPt.get(run) ?? 0);
          const last = underlines[underlines.length - 1];
          if (last && last.color === color && last.top === top && last.height === underline.thicknessPt && Math.abs(last.x1 - segment.x) < 0.5) {
            last.x1 = segment.x + segment.width;
          } else {
            underlines.push({ x0: segment.x, x1: segment.x + segment.width, top, height: underline.thicknessPt, color });
          }
        }
        const highlight = typeof run.style?.highlight === "string" ? run.style.highlight : undefined;
        if (highlight) {
          spans.push(
            <rect key={`h${segmentIndex}`} x={px(segment.x)} y={px(line.top)} width={px(segment.width)} height={px(line.bottom - line.top)} fill={highlight} />,
          );
        }
        const glyphs = segment.glyphs.filter((glyph) => glyph.text !== "\n");
        const shown = (text: string) => (caps ? text.toUpperCase() : text);
        if (glyphs.every((glyph) => shown(glyph.text).length === 1)) {
          spans.push(
            <tspan key={segmentIndex} data-run-index={segment.run} x={glyphs.map((glyph) => px(glyph.x)).join(" ")} style={paint}>
              {glyphs.map((glyph) => shown(glyph.text)).join("")}
            </tspan>,
          );
        } else {
          glyphs.forEach((glyph, glyphIndex) => {
            spans.push(
              <tspan key={`${segmentIndex}-${glyphIndex}`} data-run-index={segment.run} x={px(glyph.x)} style={paint}>
                {shown(glyph.text)}
              </tspan>,
            );
          });
        }
      });
      if (!spans.length) return;
      const highlights = spans.filter((span) => span.type === "rect");
      const texts = spans.filter((span) => span.type !== "rect");
      lineNodes.push(
        <g key={`${rowIndex}-${lineIndex}`} data-testid="presentation-engine-line" data-engine-line={lineIndex}>
          {highlights}
          <text y={px(line.baseline)} style={{ ...SVG_TEXT_STYLE, ...rowPaint }}>
            {texts}
          </text>
          {underlines.map((rule, index) => (
            // PowerPoint's underline is whole points long too.
            <rect key={`u${index}`} data-testid="presentation-engine-underline" x={px(rule.x0)} y={px(rule.top)} width={px(Math.max(1, Math.round(rule.x1 - rule.x0)))} height={px(rule.height)} fill={rule.color} />
          ))}
        </g>,
      );
    });
    nodes.push(
      overlay ? (
        <g key={rowIndex} data-testid="presentation-engine-paragraph">
          {markerNodes}
          <g>{lineNodes}</g>
        </g>
      ) : (
        <g key={rowIndex} data-testid="presentation-engine-paragraph" data-presentation-paragraph-index={rowIndex}>
          {markerNodes}
          <g data-presentation-paragraph-text="true">{lineNodes}</g>
        </g>
      ),
    );
  });
  // A press on the drawn words finds its character from the engine's lines
  // (`officeTextOffsetAtPoint`): SVG text is no caret target of the body's.
  const registerHits = (element: HTMLDivElement | null) => {
    if (element) registerOfficeTextHitMap(element, { layout, runLengths: input.rowRuns.map((runs) => runs.map((run) => run.text.length)) });
  };
  return (
    <div
      ref={registerHits}
      data-testid={overlay ? "presentation-engine-edit-text" : "presentation-engine-text"}
      {...(overlay
        ? { [OFFICE_TEXT_HIT_ATTRIBUTE]: "", [OFFICE_TEXT_MARKER_ATTRIBUTE]: "", "aria-hidden": true, contentEditable: false, "data-presentation-selection-ignore": "true" }
        : { [PRESENTATION_TEXT_ENGINE_ATTRIBUTE]: "laid", [OFFICE_TEXT_HIT_ATTRIBUTE]: "" })}
      style={{ position: "relative", width: "100%", height: px(layout.height), ...(overlay ? { userSelect: "none" as const } : {}) }}
    >
      <svg
        data-testid="presentation-engine-text-svg"
        style={{ position: "absolute", left: 0, top: 0, width: "100%", height: "100%", overflow: "visible" }}
      >
        {nodes}
      </svg>
    </div>
  );
}

/**
 * A body drawn by the engine where it can be, else `fallback` (the
 * browser-flowed drawing) -- for a host whose own render cannot call the
 * engine's hook unconditionally (a shape's label).
 */
export function PresentationEngineOrFallback({
  rows,
  bodyStyle,
  fontSizeUnit,
  autofitScale,
  tabStops,
  defaultTabPx,
  kerning,
  widthPx,
  wrap,
  colorOverride,
  fallback,
}: {
  rows: readonly PresentationParagraphRowPlan[] | null;
  bodyStyle: CSSProperties;
  fontSizeUnit: PresentationFontSizeUnit;
  autofitScale: number;
  tabStops?: unknown;
  defaultTabPx?: unknown;
  kerning?: unknown;
  widthPx: number | null;
  wrap: boolean;
  colorOverride?: string;
  fallback: ReactElement;
}): ReactElement {
  const fontAliases = useContext(PresentationFontAliasesContext);
  const engine = usePresentationEngineBody({
    enabled: true,
    rows,
    bodyStyle,
    fontSizeUnit,
    fontAliases,
    autofitScale,
    tabStops,
    defaultTabPx,
    kerning,
    widthPx,
    wrap,
  });
  if (engine.status === "ok" && rows) {
    return (
      <PresentationEngineTextBody
        answer={engine}
        rows={rows}
        fontSizeUnit={fontSizeUnit}
        fontAliases={fontAliases}
        autofitScale={autofitScale}
        colorOverride={colorOverride}
      />
    );
  }
  return engine.status === "pending" ? (
    <>
      <span hidden {...{ [PRESENTATION_TEXT_ENGINE_ATTRIBUTE]: "pending" }} />
      {fallback}
    </>
  ) : (
    fallback
  );
}
