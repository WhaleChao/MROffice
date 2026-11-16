/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The spreadsheet's adapter to the shared text engine (lib/text-engine):
 * cell text widths, glyph advances, wrap lines and line metrics computed
 * from the served font files, so the grid decides spill, wrap, "####",
 * AutoFit and Shrink to Fit the same way in every browser. Every function
 * answers undefined when the engine cannot (switched off, faces still
 * loading, a font only the visitor's machine has): the caller keeps its
 * browser measurement then.
 *
 * When faces the engine was waiting for land, the grid is told the way it
 * is told a web font landed (a `loadingdone` on document.fonts), so every
 * cache keyed on fonts empties and the cells measure again.
 */
import type { CSSProperties } from "react";
import { OFFICE_FAUX_BOLD_STROKE_WIDTH } from "../office-fonts/faux-bold";
import { textEngineEnabled } from "../text-engine/flags";
import OFFICE_DIGIT_ADVANCES from "./office-digit-advances.json";
import OFFICE_DIGIT_WHOLE_EM from "./office-digit-whole-em.json";
import { fontLoadReach } from "./font-load-reach";
import {
  engineFontFromCss,
  measureLineMetrics,
  measureLineBox,
  measureLineBoxHeight,
  measureText,
  onTextEngineChange,
  textEngineEpoch,
  explainText,
  type EngineFont,
  type EngineLineMetrics,
} from "../text-engine/measure";

let bridged = false;
function bridgeFontEvents(): void {
  if (bridged || typeof document === "undefined" || !document.fonts) return;
  bridged = true;
  // Probes and the phase-2 adapters read the engine through the page.
  (window as unknown as Record<string, unknown>).__mraiTextEngine = { measureText, measureLineMetrics, engineFontFromCss, textEngineEpoch, explainText };
  let queued = false;
  onTextEngineChange(() => {
    if (queued) return;
    queued = true;
    setTimeout(() => {
      queued = false;
      try {
        document.fonts.dispatchEvent(new Event("loadingdone"));
      } catch { /* an engine without EventTarget fonts */ }
    }, 0);
  });
}

export function spreadsheetTextEngineOn(): boolean {
  if (!textEngineEnabled("spreadsheet")) return false;
  bridgeFontEvents();
  return true;
}

/** A cache key part for anything built on cell measurements: changes when a
 * measurement may have moved from the browser's to the engine's. */
let epochOfThisTask: string | null = null;

/**
 * The text engine's epoch, read once a task as {@link fontMeasureStamp} is:
 * it is asked per cell drawn, and each read spelt out the face index's
 * signature (`document.styleSheets.length`, a style recalculation when styles
 * are dirty), 49 ms of a 1.6 MB workbook's open. Faces land and are verified
 * in tasks of their own, so the epoch holds within one.
 */
export function spreadsheetTextEngineEpoch(): string {
  if (!spreadsheetTextEngineOn()) return "";
  if (typeof queueMicrotask !== "function") return textEngineEpoch();
  if (epochOfThisTask === null) {
    epochOfThisTask = textEngineEpoch();
    queueMicrotask(() => { epochOfThisTask = null; });
  }
  return epochOfThisTask;
}

let fontLoads = 0;
let countingFontLoads = false;

let stampOfThisTask: string | null = null;

/** Changes whenever a font could measure differently: a key part for
 * anything kept from a measurement of text. */
export function fontMeasureStamp(): string {
  if (typeof document === "undefined") return "";
  if (!countingFontLoads) {
    countingFontLoads = true;
    // A load that names its faces is for each cache to judge (it drops what
    // the faces reach, font-load-reach); one that names none moves the stamp
    
    document.fonts?.addEventListener("loadingdone", (event) => { if (!fontLoadReach(event)) fontLoads += 1; });
  }
  // Asked per cell drawn, and fonts do not land inside a task: read once a task.
  if (stampOfThisTask === null) {
    stampOfThisTask = `${fontLoads}:${document.fonts?.size ?? 0}:${spreadsheetTextEngineEpoch()}`;
    queueMicrotask(() => { stampOfThisTask = null; });
  }
  return stampOfThisTask;
}











function cellWidthFont(fontCss: string): EngineFont | undefined {
  const font = engineFontFromCss(fontCss);
  return font ? { ...font, followLocalOnly: true } : undefined;
}

/**
 * Per-grapheme advances at 1000px (the design widths the native grid rounds
 * to its device pixels), in the font a canvas `font` string names. Segments
 * that are line breaks get 0. Undefined unless every segment is known.
 */
export function spreadsheetEngineIdealAdvances(segments: readonly string[], fontCss: string): number[] | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  const font = cellWidthFont(fontCss);
  if (!font) return undefined;
  const at1000: EngineFont = { ...font, sizePx: 1000 };
  const out: number[] = [];
  for (const segment of segments) {
    if (segment === "\n" || segment === "\r" || segment === "\r\n") {
      out.push(0);
      continue;
    }
    const measured = measureText(segment, at1000);
    if (measured.status !== "ok") return undefined;
    out.push(measured.width);
  }
  return out;
}

/** The width of a text in a canvas `font` string, as the engine measures it. */
export function spreadsheetEngineTextWidth(text: string, fontCss: string): number | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  const font = cellWidthFont(fontCss);
  if (!font) return undefined;
  const measured = measureText(text, font);
  return measured.status === "ok" ? measured.width : undefined;
}

/** The primary face's line metrics at a size, from the font file. */
export function spreadsheetEngineLineMetrics(fontCss: string): EngineLineMetrics | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  const font = engineFontFromCss(fontCss);
  if (!font) return undefined;
  const metrics = measureLineMetrics(font);
  return typeof metrics === "string" ? undefined : metrics;
}

/** Excel for Mac's AutoFit room beside the text, in 72-dpi pixels. */
const MAC_AUTOFIT_PADDING_PX = 12;

/** The Normal font's maximum digit width on Excel for Mac's 72-dpi grid
 * (the same table officeDigitWidth reads at 96 dpi). */
function macDigitWidth(font: { readonly family: string; readonly sizePt: number }): number {
  const key = font.family.trim().toLowerCase();
  if ((OFFICE_DIGIT_WHOLE_EM as readonly string[]).includes(key)) return Math.max(1, Math.floor(Math.round(font.sizePt) / 2 + 0.5));
  const advances = OFFICE_DIGIT_ADVANCES as Readonly<Record<string, number>>;
  return Math.max(1, Math.round((advances[key] ?? advances.calibri) * font.sizePt));
}












export function spreadsheetExcelAutoFitTextWidth(
  text: string,
  fontCss: string,
  grid: { readonly normalFont: { readonly family: string; readonly sizePt: number }; readonly digitWidth: number; readonly paddingPx: number },
): number | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  const font = cellWidthFont(fontCss);
  if (!font) return undefined;
  const at72 = measureText(text, { ...font, sizePx: font.sizePx * 0.75 });
  const drawn = measureText(text, font);
  if (at72.status !== "ok" || drawn.status !== "ok") return undefined;
  const devicePixels = at72.graphemes.reduce((sum, grapheme) => sum + Math.round(grapheme.advance), 0);
  const macDigit = macDigitWidth(grid.normalFont);
  const characters = (devicePixels + MAC_AUTOFIT_PADDING_PX - 5) / macDigit;
  const column = characters * grid.digitWidth + 5;
  return Math.max(drawn.width, column - 2 * grid.paddingPx);
}

/** The stroke an engine-measured faux bold draws with (faux-bold.ts's own
 * measured value, ~1px at 11pt). */
const ENGINE_FAUX_BOLD_STROKE = "0.06em";

/** The stroke a printed page's faux bold draws with: Windows Excel prints a
 * face without a bold cut as its regular outlines filled and stroked one
 * hundredth of an em wide (Microsoft Print to PDF, Tr 2 at w = size / 100 in
 * all 15 bench workbooks that print one: R121-X08's bold ＭＳ Ｐゴシック "1",
 * R361-X03's 245 runs, SimSun and MS Mincho alike) -- not the screen's one
 * pixel of emboldening, 0.06 em, six times as heavy on paper. */
const PRINTED_FAUX_BOLD_STROKE = "0.01em";

let fauxBoldPrinted = false;
/** Whether faux bold is drawn as Excel prints it (the print page, set with
 * the printer's resolution: native-typography setNativeTextPrinterDpi). */
export function setSpreadsheetFauxBoldPrinted(printed: boolean): void {
  fauxBoldPrinted = printed;
}

/**
 * Bold in an Office family with no bold face. Excel keeps the regular
 * advances (synthetic-bold.ts), and so does the engine; WebKit's own
 * synthesis widens every glyph and Chrome on macOS synthesizes none. With
 * the engine on, both browsers draw it one way: browser weight synthesis off
 * (`font-synthesis: style`, italic synthesis kept) and the stroke always
 * on, so the drawn width is the measured one in every browser. Off, the
 * stroke stands in only where the browser does not synthesize, as before.
 */
export function spreadsheetFauxBoldCss(): CSSProperties {
  if (fauxBoldPrinted) return { WebkitTextStrokeWidth: PRINTED_FAUX_BOLD_STROKE, fontSynthesis: "style" };
  return spreadsheetTextEngineOn()
    ? { WebkitTextStrokeWidth: ENGINE_FAUX_BOLD_STROKE, fontSynthesis: "style" }
    : { WebkitTextStrokeWidth: OFFICE_FAUX_BOLD_STROKE_WIDTH };
}

/** Whether a stroke width is one {@link spreadsheetFauxBoldCss} set. */
export function isSpreadsheetFauxBoldStroke(value: unknown): boolean {
  return value === ENGINE_FAUX_BOLD_STROKE || value === PRINTED_FAUX_BOLD_STROKE || value === OFFICE_FAUX_BOLD_STROKE_WIDTH;
}

/** The height of a laid-out paragraph's lines (`line-height: normal`), each
 * line a list of runs in canvas `font` strings, from the faces the engine
 * resolves; undefined unless every face is known. */
/** Each line box of a paragraph (as {@link spreadsheetEngineParagraphHeight}
 * adds them up): its top to the baseline and the baseline to its bottom, px. */
export function spreadsheetEngineLineBoxes(lines: readonly (readonly { readonly text: string; readonly fontCss: string }[])[], strutCss: string): { top: number; bottom: number }[] | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  const strut = engineFontFromCss(strutCss);
  if (!strut) return undefined;
  const boxes: { top: number; bottom: number }[] = [];
  for (const line of lines) {
    const runs: { text: string; font: EngineFont }[] = [];
    for (const run of line) {
      const font = engineFontFromCss(run.fontCss);
      if (!font) return undefined;
      runs.push({ text: run.text, font });
    }
    const box = measureLineBox(runs, strut);
    if (box === undefined) return undefined;
    boxes.push(box);
  }
  return boxes;
}

export function spreadsheetEngineParagraphHeight(lines: readonly (readonly { readonly text: string; readonly fontCss: string }[])[], strutCss: string): number | undefined {
  if (!spreadsheetTextEngineOn()) return undefined;
  const strut = engineFontFromCss(strutCss);
  if (!strut) return undefined;
  let height = 0;
  for (const line of lines) {
    const runs: { text: string; font: EngineFont }[] = [];
    for (const run of line) {
      const font = engineFontFromCss(run.fontCss);
      if (!font) return undefined;
      runs.push({ text: run.text, font });
    }
    const box = measureLineBoxHeight(runs, strut);
    if (box === undefined) return undefined;
    height += box;
  }
  return height;
}
