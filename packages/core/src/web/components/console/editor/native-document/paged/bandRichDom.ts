/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * DOM builder for RICH header/footer bands (Word letterhead parity):
 * paragraph runs (+ PAGE/NUMPAGES token substitution), embedded pictures
 * (data: URIs only — enforced by normalizeBandRichBlocks upstream) and mini
 * layout tables. Pure DOM construction (never innerHTML), shared by the
 * on-screen sheet overlay and both print pipelines so the three surfaces
 * render byte-identically.
 */

import { EMU_PER_PT, EMU_PER_PX, EMU_PER_TWIP, PX_PER_PT, ptToPx } from "../../../../../lib/office-units";
import { ruleAttrs } from "../wordHorizontalRule";
import { resolveFontScriptSlices } from "../../../../office-ribbon/officeFontScriptSlots";
import { textHasEastAsian, wordSingleLineRatio } from "../noteTextStyle";
import { documentStyleFontSizePx } from "../documentStyleFontSize";
import {
  documentParagraphBordersHtmlAttributes,
  documentTextBoxFillCss,
  documentTextBoxOutlineHtmlAttributes,
  paintPhysicalParagraphBorderEdges,
} from "../../../../../lib/document-borders";

import { paintShape } from "../documentShapeExtension";
import { drawDocumentPicture } from "../../../../../lib/document-picture-references";
import { floatWrapDistancesPx, type FloatAnchor as BandFloatAnchor } from "../../../../../lib/office-float-anchor";
import { normalizeChartBlockAttrs, type ChartBlockAttrsV2 } from "../../extensions/chartBlockShared";
import { documentChartPlanInputs } from "../../extensions/documentChartHost";
import { requestChartPicture } from "../../../../../lib/office-charts/chart-picture-client";
import { documentChartPictureContent } from "../../../../../lib/office-charts/document-chart-picture";
import { sanitizeChartPicture } from "../../../../../lib/office-charts/chart-picture-sanitize";
import { normalizeDocumentShapeAttrs } from "../../../../../lib/document-shapes";
import { bandListMarkers } from "./bandListMarkers";
import { numberingSuffixTabStopTwips } from "../documentNumberingLayout";

import type {
  BandRichBlock,
  BandRichBlockType,
  BandRichBorder,
  BandRichRun,
  BandRichFrame,
  BandRichParagraph,
  BandRichLine,
  BandRichShape,
  BandRichTable,
  BandRichChart,
  BandRichWordArt,
  BandFieldToken,
} from "../headerFooter";
import { bandFieldCachedText, formatNativeFieldNumber } from "../headerFooter";
import {
  normalizePositionalTab,
  normalizeTabStops,
  wordTabAdvances,
  wordTabLeader,
  type PositionalTab,
} from "../../../../../lib/document-tabstops";

export interface BandTokenContext {
  pageLabel: string;
  totalPages: number;
  /** The page's number as a value (9 for a page its section shows as ix):
   * what a page field's own number format switch formats. */
  pageNumberValue?: number;
  /** Cached values for the document-information and date fields. */
  fields?: Partial<Record<BandFieldToken, string>>;
  /** The document's Word styles (the `styleDefinitions` doc attr): a run
   * that states no size is set in its paragraph style's size, as Word sets
   * it. Absent on a page without Word styles, whose band keeps its own. */
  styleDefinitions?: unknown;
  /** The document's Word `compatibilityMode` (the `wordCompatibilityMode`
   * doc attr; null when it states none). Word 2010 layout keeps a wrapped
   * header or footer object on the paper (`keepBandObjectsOnPaper`); absent,
   * nothing is kept. */
  wordCompatibilityMode?: number | null;
}

const styleRecord = (value: unknown): Record<string, unknown> | null =>
  value !== null && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;

/**
 * The size, in px, Word gives a header or footer run that states none: its
 * paragraph style's (the document's default paragraph style when the line
 * names none) through the style chain, else the document defaults. Null
 * without Word styles -- a page MRAI made -- whose band keeps its own 12px.
 * The painted band and the story editor (`bandLineLayout.ts`) both set a
 * line in it.
 */
export function bandStyleRunSizePx(styleDefinitions: unknown, styleId: string | null | undefined): number | null {
  const defs = styleRecord(styleDefinitions);
  if (!defs || !Array.isArray(defs.styles)) return null;
  const id =
    typeof styleId === "string" && styleId
      ? styleId
      : defs.styles
          .map(styleRecord)
          .find((style) => style?.type === "paragraph" && style.default === true && typeof style.styleId === "string")
          ?.styleId;
  if (typeof id === "string") {
    const px = documentStyleFontSizePx({ styleDefinitions: defs }, { attrs: { docxStyleId: id } });
    if (px !== null) return Number.isFinite(px) && px > 0 ? px : null;
  }
  const halfPoints = styleRecord(styleRecord(defs.docDefaults)?.rPr)?.fontSizeHalfPoints;
  return typeof halfPoints === "number" && Number.isInteger(halfPoints) && halfPoints > 0 ? (halfPoints * 2) / 3 : null;
}


/**
 * The bold and italic a header or footer run that states neither takes from
 * its paragraph style: the nearest style of the chain (the document's
 * default paragraph style when the line names none) that states each, as
 * Word reads a style hierarchy (a child stating `w:b w:val="0"` turns its
 * base's bold off). Nothing without Word styles.
 */
export function bandStyleRunToggles(
  styleDefinitions: unknown,
  styleId: string | null | undefined,
): { bold: boolean; italic: boolean } {
  const defs = styleRecord(styleDefinitions);
  const out = { bold: false, italic: false };
  if (!defs || !Array.isArray(defs.styles)) return out;
  const styles = defs.styles.map(styleRecord).filter((style): style is Record<string, unknown> => style !== null);
  const byId = new Map(styles.map((style) => [style.styleId, style]));
  let id: unknown =
    typeof styleId === "string" && styleId
      ? styleId
      : styles.find((style) => style.type === "paragraph" && style.default === true)?.styleId;
  const settled: { bold?: boolean; italic?: boolean } = {};
  const seen = new Set<unknown>();
  while (typeof id === "string" && byId.has(id) && !seen.has(id) && seen.size < 32) {
    seen.add(id);
    const style = byId.get(id) as Record<string, unknown>;
    const rPr = styleRecord(style.rPr);
    for (const key of ["bold", "italic"] as const) {
      if (settled[key] === undefined && typeof rPr?.[key] === "boolean") settled[key] = rPr[key] as boolean;
    }
    id = style.basedOn;
  }
  const defaults = styleRecord(styleRecord(defs.docDefaults)?.rPr);
  for (const key of ["bold", "italic"] as const) {
    out[key] = settled[key] ?? (defaults?.[key] === true);
  }
  return out;
}


/** A Word border as a CSS border: its style (Word's single, double, dotted,
 * dashed and their thick kin; any other drawn as a single line), its width in
 * eighths of a point and its colour (`auto` is black). `nil` / `none` draw
 * nothing. */
function bandBorderCss(border: BandRichBorder | undefined): string | null {
  if (!border) return null;
  const value = border.val.toLowerCase();
  if (value === "nil" || value === "none") return "none";
  const style = value.startsWith("double") || value.startsWith("triple") ? "double"
    : value.startsWith("dotted") ? "dotted"
      : value.startsWith("dash") ? "dashed" : "solid";
  const widthPx = Math.max(1, ptToPx(((border.sz ?? 4) / 8)));
  const color = !border.color || border.color === "auto" ? "#000000" : `#${border.color}`;
  return `${style === "double" ? Math.max(3, widthPx) : widthPx}px ${style} ${color}`;
}

function runText(run: BandRichRun, context: BandTokenContext): string {
  // A page-number field of Word's part shows its own number format (its
  // `\\* roman` switch) over the page's.
  const formatted = (value: string) =>
    (run.nativeFieldFormat ?? run.fieldFormat) && (run.nativeFieldFormat ?? run.fieldFormat) !== "none" && /^[0-9]{1,6}$/.test(value)
      ? formatNativeFieldNumber(Number(value), run.nativeFieldFormat ?? run.fieldFormat)
      : value;
  if (run.token === "PAGE") {
    // An explicit switch formats the page's number, not the label its
    // section's format made of it: \* Arabic on a page shown as ix is 9
    // (OFL-013).
    const format = run.nativeFieldFormat ?? run.fieldFormat;
    const value = context.pageNumberValue;
    if (format && format !== "none" && typeof value === "number" && Number.isInteger(value) && value > 0) {
      return formatNativeFieldNumber(value, format);
    }
    return formatted(context.pageLabel);
  }
  if (run.token === "NUMPAGES") return formatted(String(context.totalPages));
  if (run.token) {
    // A document-information or date field shows its cached text; the
    // page's own metadata fills the ones it knows.
    if (run.datePicture) return bandFieldCachedText(run.token, new Date(), run.datePicture);
    const known = context.fields?.[run.token];
    return known !== undefined ? known : bandFieldCachedText(run.token);
  }
  return run.text ?? "";
}

function appendRuns(
  parent: HTMLElement,
  runs: readonly BandRichRun[],
  context: BandTokenContext,
): void {
  const entries = runs.map(run => {
    const font = run.font;
    // The font hint alone is not the run-level w:cs/rtl switch.
    const primary = font?.ascii ?? font?.eastAsia ?? font?.cs ?? null;
    // w:cs is the face every complex-script character of the run is drawn
    // from, whatever the Latin and East Asian slots name (ECMA-376 17.3.2.26).
    return {text: runText(run, context), primary, eastAsia: font?.eastAsia ?? null, complexScript: font?.cs ?? null, word: font ? {hint: font.hint === "cs" ? "default" as const : font.hint ?? "default", language: font.language, eastAsiaCharset: font.eastAsiaCharset, eastAsiaUsesPrimary: font.eastAsia?.toLowerCase() === "times new roman" && Boolean(font.ascii)} : undefined};
  });
  const slices = runs.some(run => run.font) ? resolveFontScriptSlices(entries)
    : entries.map((entry, index) => ({index, from: 0, to: entry.text.length, slot: "primary" as const, value: null}));
  for (const [index, run] of runs.entries()) {
    const parts = slices.filter(slice => slice.index === index);
    if (!parts.length) parts.push({index, from: 0, to: entries[index].text.length, slot: "primary", value: entries[index].primary});
    for (const part of parts) {
      // A tab is its own span, sized by `layoutBandTabs` to reach its stop the
      // way Word does (centre and right stops included); the text around it
      // stays in the run's own spans.
      const pieces = entries[index].text.slice(part.from, part.to).split("\t");
      for (const [pieceIndex, piece] of pieces.entries()) {
        if (pieceIndex > 0) {
          const tab = parent.ownerDocument.createElement("span");
          tab.className = "sheet-band-rich__tab";
          tab.textContent = "\t";
          // Word's absolute position tab goes to its own point on the line.
          if (run.ptab) tab.setAttribute("data-band-ptab", JSON.stringify(run.ptab));
          if (run.underline) tab.style.textDecoration = "underline";
          if (run.sizePt) tab.style.fontSize = `${run.sizePt}pt`;
          parent.appendChild(tab);
        }
        if (!piece && pieces.length > 1) continue;
        const span = parent.ownerDocument.createElement("span");
        span.setAttribute("data-band-run", "");
        span.textContent = piece;
        if (part.value) span.style.fontFamily = `"${part.value}"`;
        span.style.whiteSpace = "pre-wrap";
        // A run that states bold or italic, on or off, says so; one that
        // states neither takes its paragraph style's (`bandStyleRunToggles`).
        if (run.bold === true) span.style.fontWeight = "700";
        else if (run.bold === false) span.style.fontWeight = "400";
        if (run.italic === true) span.style.fontStyle = "italic";
        else if (run.italic === false) span.style.fontStyle = "normal";
        if (run.underline) span.style.textDecoration = "underline";
        if (run.sizePt) span.style.fontSize = `${run.sizePt}pt`;
        if (run.color) span.style.color = run.color;
        // Superscript / Subscript: raised or lowered at a smaller size, as a
        // body run's are.
        if (run.vertAlign) {
          span.style.verticalAlign = run.vertAlign === "superscript" ? "super" : "sub";
          span.style.fontSize = run.sizePt ? `${run.sizePt * 0.65}pt` : "0.65em";
        }
        if (run.revision) {
          // Word's All Markup: an insertion underlined, a deletion struck
          // through, in the revision colour the body's tracked text uses.
          const inserted = run.revision.kind === "ins";
          span.setAttribute("data-band-revision", run.revision.kind);
          if (run.revision.author) span.setAttribute("data-tc-author", run.revision.author);
          span.style.textDecoration = [run.underline ? "underline" : "", inserted ? "underline" : "line-through"]
            .filter(Boolean)
            .filter((value, at, all) => all.indexOf(value) === at)
            .join(" ");
          span.style.color = inserted ? "var(--console-accent, #2563eb)" : "var(--console-danger, #dc2626)";
        }
        parent.appendChild(span);
      }
    }
  }
}

/**
 * Word's line height for a band line, in CSS px, or null to keep the band's
 * default. `auto` spacing scales Word's single line for the tallest run's face
 * and size; `exact` is the stated height; `atLeast` never goes below Word's
 * single line.
 */
function bandWordLineHeightPx(doc: Document, line: BandRichLine, text: string): number | null {
  if (line.lineHeightRule === "exact" && line.lineHeightPt) return line.lineHeightPt * PX_PER_PT;
  if (!line.lineSizePt || !line.lineFont) {
    return line.lineHeightRule === "atLeast" && line.lineHeightPt ? line.lineHeightPt * PX_PER_PT : null;
  }
  const ratio = wordSingleLineRatio(doc, line.lineFont, textHasEastAsian(text));
  if (ratio === null) return null;
  const single = ratio * line.lineSizePt * PX_PER_PT;
  if (line.lineHeightRule === "atLeast" && line.lineHeightPt) return Math.max(single, line.lineHeightPt * PX_PER_PT);
  return single * (line.lineHeight ?? 1);
}

/**
 * The CSS line height of one band line, or null to keep the band's own single
 * line (`.sheet-band-rich`): Word's line height from the line's spacing and
 * the face and size of its tallest run, or its exact / at-least height. The
 * painted band and the header/footer story editor (`bandLineLayout.ts`) both
 * set a line with it, so opening a header or footer moves no line up or down.
 * `text` is the line's text without its fields (East Asian text takes the
 * face's taller line).
 */
export function bandLineHeightCss(doc: Document, line: BandRichLine, text: string): string | null {
  const height = bandWordLineHeightPx(doc, line, text);
  if (height !== null && Number.isFinite(height) && height > 0) return `${Math.round(height * 1000) / 1000}px`;
  // A line whose face and size the band does not know (a line typed in MRAI)
  // still keeps its spacing: that many of the band's own single lines.
  const multiple = line.lineHeight;
  if (line.lineHeightRule === undefined && typeof multiple === "number" && Number.isFinite(multiple) && multiple > 0 && multiple !== 1) {
    return String(Math.round(BAND_SINGLE_LINE_EM * multiple * 10000) / 10000);
  }
  return null;
}

/** The band's own single line, per em (`.sheet-band-rich`'s line height). */
export const BAND_SINGLE_LINE_EM = 1.35;

function applyBandLineHeight(element: HTMLElement, line: BandRichLine, runs: readonly BandRichRun[]): void {
  const text = runs.map((run) => run.text ?? "").join("");
  const height = bandLineHeightCss(element.ownerDocument, line, text);
  if (height !== null) element.style.lineHeight = height;
  // A Word line height in px is the line's whole height (note-editor.css).
  if (height?.endsWith("px")) element.setAttribute("data-band-word-line", "");
}


/**
 * A band paragraph's lines, each as tall as the characters on it. Word's line
 * is as tall as its tallest run and takes nothing from the paragraph mark while
 * it holds text (Word 16.86, g-l4mark: 10 pt text over a 20 pt mark, or under a
 * 24 pt style, stands at its own 11.5 pt line; corpus R283-D12's header, a 10 pt
 * line over a 12 pt line, in a paragraph marked 12 pt Consolas, drew both at the
 * mark's 14 pt and the body 1.7 pt low). The paragraph keeps its lowest run's
 * line, and a taller run raises the lines it stands on. White space raises
 * nothing here: at a line's end it hangs on the line it is drawn past.
 */
function applyBandRunLineHeights(paragraph: HTMLElement, line: BandRichLine): void {
  if (line.lineHeightRule !== undefined || !line.lineFont || !line.lineSizePt) return;
  const doc = paragraph.ownerDocument;
  const spans = (Array.from(paragraph.children) as HTMLElement[]).filter(
    (child) => child.hasAttribute("data-band-run") && /\S/.test(child.textContent ?? ""),
  );
  if (!spans.length) return;
  const paragraphPt = paragraph.style.fontSize.endsWith("px")
    ? Number.parseFloat(paragraph.style.fontSize) / PX_PER_PT
    : line.lineSizePt;
  const heights: number[] = [];
  for (const span of spans) {
    const sizePt = span.style.fontSize.endsWith("pt") ? Number.parseFloat(span.style.fontSize) : paragraphPt;
    const face = span.style.fontFamily.replace(/["']/g, "") || line.lineFont;
    const ratio = wordSingleLineRatio(doc, face, textHasEastAsian(span.textContent ?? ""));
    if (ratio === null || !(sizePt > 0)) return;
    heights.push(ratio * sizePt * PX_PER_PT * (line.lineHeight ?? 1));
  }
  const css = (px: number) => `${Math.round(px * 1000) / 1000}px`;
  const lowest = Math.min(...heights);
  paragraph.style.lineHeight = css(lowest);
  spans.forEach((span, index) => {
    if (heights[index] > lowest + 0.01) span.style.lineHeight = css(heights[index]);
  });
}

function frameVerticalFromParagraph(frame: BandRichFrame): boolean {
  return frame.vertical.relativeFrom === "paragraph" || frame.vertical.relativeFrom === "line";
}

/**
 * The frames this paragraph can place where Word places them. A frame whose
 * origin the band cannot resolve exactly is left out, and its runs stay in the
 * paragraph flow -- which is how the band drew every header text box before.
 */
function placeableBandFrames(block: BandRichParagraph): Map<number, BandRichFrame> {
  const out = new Map<number, BandRichFrame>();
  for (const frame of block.frames ?? []) {
    // An in-line box takes a place in the line (`appendLineWithInlineFrames`).
    if (frame.inline) continue;
    const h = frame.horizontal;
    const v = frame.vertical;
    // Centred or right on the page, and set at the page's or the margin
    // box's top, middle or bottom, as `placeBandBox` places a band's
    // anchored drawings: a classification label's "OFFICIAL" box centred at
    // the page foot was left in its right-aligned footer line (corpus
    // R282-D01, right where Word centres it).
    const horizontalOk =
      (h.offsetEmu !== undefined && ["column", "margin", "character", "page", "leftMargin"].includes(h.relativeFrom)) ||
      (h.align !== undefined && ["column", "margin"].includes(h.relativeFrom) && ["left", "center", "right"].includes(h.align)) ||
      (h.align !== undefined && h.relativeFrom === "page" && ["center", "right"].includes(h.align));
    const verticalOk =
      (v.offsetEmu !== undefined && ["paragraph", "line", "page", "margin", "topMargin"].includes(v.relativeFrom)) ||
      (v.offsetEmu === undefined && v.align !== undefined && ["top", "center", "bottom"].includes(v.align) &&
        ["page", "margin"].includes(v.relativeFrom) && frame.heightEmu > 0);
    if (horizontalOk && verticalOk) out.set(frame.index, frame);
  }
  return out;
}

function pt(emu: number): string {
  return `${emu / EMU_PER_PT}pt`;
}

/**
 * Puts one absolutely positioned band box where Word's `wp:anchor` axes put
 * it. Horizontal origins: the band content box IS the text column, so
 * `column`/`margin`/`character` offsets apply as they are and
 * `page`/`leftMargin` subtract the left page margin. Vertical: a
 * paragraph/line origin is the anchor paragraph's own top (the caller makes
 * the box that paragraph's child); a page/margin origin is measured from the
 * band box, which sits at the header distance from the page top. One rule for
 * a text-box frame and an anchored drawing, so the two can never disagree
 * about where the same `wp:anchor` lands.
 */
/**
 * A header or footer object the body text wraps around (Square, Tight,
 * Through, Top and Bottom): its wrap and wrap distances, for the page's
 * exclusion pass (`bandFloatExclusion`). Word lays the body out around such
 * an object on every page the band is on (corpus R22-D30's sidebar).
 */
function markBandWrap(element: HTMLElement, anchor: BandFloatAnchor): void {
  const kind = anchor.wrap?.kind;
  if (!kind || kind === "none" || anchor.behindDoc) return;
  const distances = floatWrapDistancesPx(anchor);
  element.dataset.bandWrap = kind;
  element.dataset.bandWrapText = String((anchor.wrap as { text?: string }).text ?? "bothSides");
  element.dataset.bandWrapDist = [distances.topPx, distances.rightPx, distances.bottomPx, distances.leftPx]
    .map((value) => Math.round(value * 100) / 100)
    .join(",");
}


/**
 * A header or footer table with Text Wrapping: Around placed from the page
 * or the margin (`w:tblpPr`): Word draws it there, out of the band's flow,
 * and wraps the body around it (corpus R98-D08: a letterhead's address
 * table 133pt down the first page; the letter starts below it). It was a
 * table in the header's flow at the header distance, over the letter.
 */
function placeBandFloatingTable(table: HTMLElement, block: BandRichTable): void {
  const position = block.floatPosition;
  if (!position) return;
  const widthTwips = block.nativeColumnWidthsTwips?.reduce((total, width) => total + width, 0) ?? 0;
  table.style.position = "absolute";
  table.style.margin = "0";
  placeBandBox(
    table,
    {
      relativeFrom: position.horzAnchor,
      ...(position.xTwips !== undefined ? { offsetEmu: position.xTwips * EMU_PER_TWIP } : { align: position.xAlign }),
    },
    { relativeFrom: position.vertAnchor, offsetEmu: position.yTwips * EMU_PER_TWIP },
    widthTwips * EMU_PER_TWIP,
  );
  const [top, right, bottom, left] = position.distTwips.map((twips) => Math.round((twips / 15) * 100) / 100);
  table.dataset.bandWrap = "square";
  table.dataset.bandWrapText = "bothSides";
  table.dataset.bandWrapDist = [top, right, bottom, left].join(",");
}

/** A header or footer object Word keeps on the paper in Word 2010 layout:
 * one the text wraps around (`keepBandObjectsOnPaper`). */
function markBandKeep(element: HTMLElement, anchor: BandFloatAnchor | null | undefined): void {
  const kind = anchor?.wrap?.kind;
  if (kind && kind !== "none" && !anchor?.behindDoc) element.setAttribute("data-band-keep", "");
}

/**
 * Word 2010 layout (compatibility mode 14 and earlier, or none stated) keeps
 * a header or footer object the text wraps around inside the paper, as it
 * keeps a body one (`office-float-anchor.keepFloatOnPage`; measured in Word
 * 16.86: a header picture 2" above its paragraph starts at the paper's top
 * edge, a footer one 2" below ends at its bottom edge). Word 2013 layout
 * keeps none of them. Run once the band is laid out, on screen and in the
 * print copy alike: the object's box is measured against its page (the
 * band box over the sheet, or the printed sheet) and moved back onto it.
 */
export function keepBandObjectsOnPaper(root: ParentNode): void {
  for (const band of root.querySelectorAll<HTMLElement>("[data-band-keep-on-paper]")) {
    const paper = band.closest<HTMLElement>(".sheet-band-box, article[data-page-sheet]");
    if (!paper || !(paper.offsetWidth > 0) || !(paper.offsetHeight > 0)) continue;
    const boxes = Array.from(band.querySelectorAll<HTMLElement>("[data-band-keep]"));
    const paperRect = paper.getBoundingClientRect();
    const scale = paperRect.width / paper.offsetWidth || 1;
    const within = (start: number, size: number, high: number) => Math.max(0, Math.min(start, high - size));
    // Measured where the anchor alone puts each box (its current place less
    // the move this pass wrote last), and written only when the move
    // changes: a pass run again on an unchanged band writes nothing.
    const moves = boxes.map((box) => {
      const [kept = "0", keptY = "0"] = box.style.translate.split(" ");
      const keptDx = Number.parseFloat(kept) || 0;
      const keptDy = Number.parseFloat(keptY) || 0;
      const rect = box.getBoundingClientRect();
      const left = (rect.left - paperRect.left) / scale - keptDx;
      const top = (rect.top - paperRect.top) / scale - keptDy;
      const dx = within(left, rect.width / scale, paper.offsetWidth) - left;
      const dy = within(top, rect.height / scale, paper.offsetHeight) - top;
      return { box, value: Math.abs(dx) > 0.5 || Math.abs(dy) > 0.5 ? `${Math.round(dx * 100) / 100}px ${Math.round(dy * 100) / 100}px` : "" };
    });
    for (const { box, value } of moves) {
      if (box.style.translate !== value) box.style.translate = value;
    }
  }
}

function placeBandBox(
  box: HTMLElement,
  h: { relativeFrom: string; offsetEmu?: number; align?: string },
  v: { relativeFrom: string; offsetEmu?: number; align?: string },
  widthEmu: number,
  heightEmu?: number,
): void {
  const marginLeft = "var(--page-margin-left, var(--page-margin-x, 96px))";
  if (h.offsetEmu !== undefined) {
    box.style.left =
      h.relativeFrom === "page" || h.relativeFrom === "leftMargin"
        ? `calc(${pt(h.offsetEmu)} - ${marginLeft})`
        : pt(h.offsetEmu);
  } else if (h.relativeFrom === "page" && (h.align === "center" || h.align === "right")) {
    // Aligned on the PAGE: the band box is the text column, so the page's
    // own edges are a left margin away on the left and a right margin away
    // on the right (Word's Layout > Position > Alignment "relative to Page").
    const pageWidth = `calc(100% + ${marginLeft} + var(--page-margin-right, var(--page-margin-x, 96px)))`;
    box.style.left =
      h.align === "center"
        ? `calc((${pageWidth} - ${pt(widthEmu)}) / 2 - ${marginLeft})`
        : `calc(${pageWidth} - ${pt(widthEmu)} - ${marginLeft})`;
  } else if (h.align === "center") {
    box.style.left = `calc((100% - ${pt(widthEmu)}) / 2)`;
  } else if (h.align === "right") {
    box.style.left = `calc(100% - ${pt(widthEmu)})`;
  } else {
    box.style.left = "0";
  }
  const offset = pt(v.offsetEmu ?? 0);
  const headerOffset = "var(--sheet-header-offset, 0px)";
  const marginTop = "var(--page-margin-top, var(--page-margin-y, 96px))";
  // A `wp:align` (top / center / bottom) instead of an offset places the box
  // inside its origin's own height -- the page, or the page's margin box.
  // Only a page/margin origin has a height a band can know; a paragraph or
  // line origin keeps its top.
  const aligned =
    v.offsetEmu === undefined &&
    heightEmu !== undefined &&
    (v.align === "center" || v.align === "bottom") &&
    v.relativeFrom !== "paragraph" &&
    v.relativeFrom !== "line";
  if (v.relativeFrom === "paragraph" || v.relativeFrom === "line") {
    box.style.top = offset;
  } else if (v.relativeFrom === "page") {
    if (aligned) {
      const room = `(var(--page-height-px, 1123px) - ${pt(heightEmu)})`;
      box.style.top = `calc(${v.align === "center" ? `${room} / 2` : room} - ${headerOffset})`;
    } else {
      box.style.top = `calc(${offset} - ${headerOffset})`;
    }
  } else if (aligned) {
    // margin / topMargin aligned: inside the text margin box.
    const room = `(var(--page-height-px, 1123px) - ${marginTop} - var(--page-margin-bottom, var(--page-margin-y, 96px)) - ${pt(heightEmu)})`;
    box.style.top = `calc(${marginTop} + ${v.align === "center" ? `${room} / 2` : room} - ${headerOffset})`;
  } else {
    // margin / topMargin: the top text margin edge.
    box.style.top = `calc(${offset} + ${marginTop} - ${headerOffset})`;
  }
}

/**
 * One anchored text box of a header/footer paragraph, positioned by Word's
 * `wp:anchor` rules against the band. Horizontal origins: the band content
 * box IS the text column, so `column`/`margin`/`character` offsets apply as
 * they are and `page`/`leftMargin` subtract the left page margin. Vertical:
 * see {@link frameVerticalFromParagraph}; a page/margin origin is measured
 * from the band box, which sits at the header distance from the page top.
 */
/**
 * A band text box's own fill and line, painted as the body's `textBox` frame
 * is (`lib/document-borders`): a header box used to be drawn as bare words.
 */
function paintBandFrame(box: HTMLElement, frame: BandRichFrame): void {
  const fill = frame.textBoxFill ? documentTextBoxFillCss(frame.textBoxFill) : null;
  if (fill) box.style.background = fill;
  const border = /(?:^|;)border:([^;]+)/.exec(documentTextBoxOutlineHtmlAttributes(frame.textBoxOutline).style ?? "")?.[1];
  if (border) box.style.border = border;
}

function buildBandFrame(
  doc: Document,
  frame: BandRichFrame,
  runs: readonly BandRichRun[],
  context: BandTokenContext,
): HTMLElement {
  const box = doc.createElement("div");
  box.className = "sheet-band-rich__frame";
  box.setAttribute("data-band-frame", String(frame.index));
  box.style.position = "absolute";
  box.style.boxSizing = "border-box";
  // A box Word fits to its words both ways is as wide as they are, on every
  // page (its PAGE number's digits); otherwise its width is the one read.
  box.style.width = frame.autoWidth ? "max-content" : pt(frame.widthEmu);
  // A box Word sizes to its words states only a floor for its height.
  if (frame.autoHeight) box.style.minHeight = pt(frame.heightEmu);
  else box.style.height = pt(frame.heightEmu);
  paintBandFrame(box, frame);
  const [left, top, right, bottom] = frame.insetsEmu;
  box.style.padding = `${pt(top)} ${pt(right)} ${pt(bottom)} ${pt(left)}`;
  box.style.textAlign = frame.align;
  box.style.overflow = "visible";
  // Word's behindDoc frames paint under the text; the others above it, all in
  // one document-wide relativeHeight order (no stacking context of their own).
  box.style.zIndex = frame.behindText ? "-1" : "auto";

  placeBandBox(box, frame.horizontal, frame.vertical, frame.widthEmu, frame.heightEmu);
  if (frame.autoWidth && frame.horizontal.offsetEmu === undefined) {
    // Aligned on its origin by its own width, which is its words': the
    // centre or the right edge stays where Word keeps it.
    if (frame.horizontal.align === "center") {
      box.style.left = frame.horizontal.relativeFrom === "page"
        ? `calc(${box.style.left} + ${pt(frame.widthEmu)} / 2)`
        : "50%";
      // A transform, not `translate`: the keep-on-paper pass owns that.
      box.style.transform = "translateX(-50%)";
    } else if (frame.horizontal.align === "right") {
      box.style.left = frame.horizontal.relativeFrom === "page"
        ? `calc(${box.style.left} + ${pt(frame.widthEmu)})`
        : "100%";
      box.style.transform = "translateX(-100%)";
    }
  }
  markBandKeep(box, frame.floatAnchor as BandFloatAnchor | null | undefined);
  // The frame's own turn (`a:xfrm/@rot`, 60 000ths of a degree clockwise,
  // about the box's centre), which the anchor it was read with states: the
  // box used to be drawn upright whatever the file said.
  const turn = (frame.floatAnchor?.rotation ?? 0) / 60_000;
  if (turn % 360 !== 0) {
    box.style.transformOrigin = "center";
    box.style.transform = `${box.style.transform} rotate(${turn}deg)`.trim();
  }

  // A box whose content is a STORY of its own (a letterhead grid) draws
  // through the band's own block renderer -- the same one the top level uses,
  // so the box's table is laid out by the rules every band table follows
  // (columns, native grid widths, cell margins, per-cell alignment). Before
  // this the box could only hold a run list and the grid was flattened, which
  // is what `header_footer_text_box_table_flattened` reported.
  if (frame.blocks?.length) {
    box.appendChild(buildBandRichDom(doc, frame.blocks, context, true));
    return box;
  }

  const line = doc.createElement("div");
  line.className = "sheet-band-rich__frame-text";
  if (frame.line) applyBandLineHeight(line, frame.line, runs);
  appendRuns(line, runs, context);
  box.appendChild(line);
  return box;
}

/**
 * One "In Line with Text" box of a header/footer paragraph: a box of the
 * frame's own width in the line's flow, its bottom on the baseline as Word
 * sits an in-line drawing, its height a floor (Word lets an overset box spill, and
 * sizes one that fits its text to its words), painted with the fill and the
 * line the body's in-line `textBox` frame is painted with
 * (`lib/document-borders`). It draws its own story (`blocks`) or the line's
 * runs tagged with it, as an anchored frame does.
 */
function buildBandInlineFrame(
  doc: Document,
  frame: BandRichFrame,
  runs: readonly BandRichRun[],
  context: BandTokenContext,
): HTMLElement {
  const box = doc.createElement("div");
  box.className = "sheet-band-rich__frame sheet-band-rich__frame--inline";
  box.setAttribute("data-band-frame", String(frame.index));
  box.setAttribute("data-band-frame-inline", "");
  box.style.display = "inline-block";
  // Word sits an in-line drawing with its bottom edge on the line's baseline.
  // An inline-block's baseline is its bottom edge only when it does not show
  // its overflow; the box grows with its words (`minHeight`), so nothing of
  // them is cut.
  box.style.verticalAlign = "baseline";
  box.style.overflow = "hidden";
  box.style.boxSizing = "border-box";
  box.style.width = pt(frame.widthEmu);
  box.style.maxWidth = "100%";
  box.style.minHeight = pt(frame.heightEmu);
  const [left, top, right, bottom] = frame.insetsEmu;
  box.style.padding = `${pt(top)} ${pt(right)} ${pt(bottom)} ${pt(left)}`;
  // The box is its own block of text: the host line's indent, alignment and
  // tab handling stop at its edge.
  box.style.textAlign = frame.align;
  box.style.textIndent = "0";
  box.style.whiteSpace = "normal";
  paintBandFrame(box, frame);
  if (frame.blocks?.length) {
    box.appendChild(buildBandRichDom(doc, frame.blocks, context, true));
    return box;
  }
  const line = doc.createElement("div");
  line.className = "sheet-band-rich__frame-text";
  if (frame.line) applyBandLineHeight(line, frame.line, runs);
  appendRuns(line, runs, context);
  box.appendChild(line);
  return box;
}

/**
 * A line's runs with its in-line boxes standing among them: a box that draws
 * the line's tagged runs stands where the first of them stood, one that draws
 * its own story after `runsBefore` of the line's own runs (at the end when it
 * says nothing). Runs an anchored frame draws are left to that frame.
 */
function appendLineWithInlineFrames(
  doc: Document,
  paragraph: HTMLElement,
  block: BandRichParagraph,
  floated: ReadonlyMap<number, BandRichFrame>,
  inlineFrames: readonly BandRichFrame[],
  context: BandTokenContext,
): void {
  const byIndex = new Map(inlineFrames.map((frame) => [frame.index, frame]));
  const placed = new Set<number>();
  const pending: BandRichRun[] = [];
  let own = 0;
  const flush = () => {
    if (pending.length) appendRuns(paragraph, pending.splice(0), context);
  };
  const place = (frame: BandRichFrame) => {
    if (placed.has(frame.index)) return;
    placed.add(frame.index);
    flush();
    paragraph.appendChild(
      buildBandInlineFrame(doc, frame, block.runs.filter((run) => run.frame === frame.index), context),
    );
  };
  const placeStories = (limit: number) => {
    for (const frame of inlineFrames) {
      if (frame.blocks?.length && (frame.runsBefore ?? Number.POSITIVE_INFINITY) <= limit) place(frame);
    }
  };
  for (const run of block.runs) {
    if (run.frame !== undefined && floated.has(run.frame)) continue;
    const frame = run.frame !== undefined ? byIndex.get(run.frame) : undefined;
    if (frame) {
      if (!frame.blocks?.length) place(frame);
      continue;
    }
    placeStories(own);
    pending.push(run);
    own += 1;
  }
  placeStories(Number.POSITIVE_INFINITY);
  flush();
}

/**
 * One WordArt shape of a header/footer paragraph — on this corpus always a
 * watermark: a `textPlain` warp, rotated 315°, painted behind the text.
 *
 * `textPlain` is the Plain Text warp: ISO/IEC 29500 gives every warp a TOP
 * path and a BASE path and maps the text rectangle onto them, and this one's
 * two paths are straight lines spanning the shape — so the drawing is the
 * text FITTED to the shape box, which is exactly what `transform: scale()`
 * on a measured line does. A warp the shared primitive does not draw never
 * reaches here (`normalizeBandRichWordArt` drops it), so the object stays
 * preserved and its named import receipt is what the reader is told.
 */
function buildBandWordArt(doc: Document, art: BandRichWordArt): HTMLElement {
  const box = doc.createElement("div");
  box.className = "sheet-band-rich__wordart";
  box.setAttribute("data-band-wordart", String(art.index));
  box.setAttribute("data-text-warp", art.warp);
  box.style.position = "absolute";
  box.style.boxSizing = "border-box";
  box.style.width = pt(art.widthEmu);
  box.style.height = pt(art.heightEmu);
  box.style.display = "flex";
  box.style.alignItems = "center";
  box.style.justifyContent = "center";
  box.style.overflow = "visible";
  box.style.pointerEvents = "none";
  box.style.whiteSpace = "pre";
  box.style.zIndex = art.behindText ? "-1" : "auto";
  if (art.rotation) {
    // 60000ths of a degree, as `a:xfrm/@rot` states it, about the box centre
    // — which is where Word rotates a watermark.
    box.style.transform = `rotate(${art.rotation / 60_000}deg)`;
  }

  const h = art.horizontal;
  if (h.offsetEmu !== undefined) {
    box.style.left =
      h.relativeFrom === "page" || h.relativeFrom === "leftMargin"
        ? `calc(${pt(h.offsetEmu)} - var(--page-margin-left, var(--page-margin-x, 96px)))`
        : pt(h.offsetEmu);
  } else if (h.align === "right") {
    box.style.left = `calc(100% - ${pt(art.widthEmu)})`;
  } else if (h.align === "left") {
    box.style.left = "0";
  } else {
    box.style.left = `calc((100% - ${pt(art.widthEmu)}) / 2)`;
  }

  const v = art.vertical;
  if (v.offsetEmu !== undefined) {
    box.style.top =
      v.relativeFrom === "page"
        ? `calc(${pt(v.offsetEmu)} - var(--sheet-header-offset, 0px))`
        : `calc(${pt(v.offsetEmu)} + var(--page-margin-top, var(--page-margin-y, 96px)) - var(--sheet-header-offset, 0px))`;
  } else {
    // A vertical ALIGNMENT against the text margin: the watermark's own case,
    // and the reason a watermark could not be placed before. The band box
    // sits at the header distance from the page top, so the margin box's
    // centre is half the text height below the top margin.
    const height = `calc(var(--page-height, 1056px) - var(--page-margin-top, var(--page-margin-y, 96px)) - var(--page-margin-bottom, var(--page-margin-y, 96px)))`;
    const top = `calc(var(--page-margin-top, var(--page-margin-y, 96px)) - var(--sheet-header-offset, 0px))`;
    box.style.top =
      v.align === "bottom"
        ? `calc(${top} + ${height} - ${pt(art.heightEmu)})`
        : v.align === "top"
          ? top
          : `calc(${top} + (${height} - ${pt(art.heightEmu)}) / 2)`;
  }

  const glyphs = doc.createElement("span");
  glyphs.className = "sheet-band-rich__wordart-text";
  glyphs.textContent = art.text;
  if (art.fontFamily) glyphs.style.fontFamily = `"${art.fontFamily}"`;
  if (art.bold) glyphs.style.fontWeight = "700";
  if (art.italic) glyphs.style.fontStyle = "italic";
  if (art.fill) glyphs.style.color = art.fill;
  if (art.fillOpacity !== undefined) glyphs.style.opacity = String(art.fillOpacity);
  if (art.outline) glyphs.style.webkitTextStroke = `1px ${art.outline}`;
  if (art.fontSizePt) {
    // The textpath states a size: Word draws the glyphs AT it and the box is
    // just where they sit.
    glyphs.style.fontSize = `${art.fontSizePt}pt`;
  } else {
    // No stated size: the Plain Text warp fills its box. The height is the
    // one measurement available without laying the text out, and Word's
    // WordArt cap height is what the box height means.
    glyphs.style.fontSize = pt(art.heightEmu);
    glyphs.style.lineHeight = "1";
  }
  box.appendChild(glyphs);
  return box;
}

const SVG_NS = "http://www.w3.org/2000/svg";

/**
 * One drawing a header/footer paragraph holds, painted by the SAME
 * `paintShape` the editor's own `documentShape` is painted by.
 *
 * It is drawn where a band PICTURE is drawn -- after the paragraph's runs, in
 * the flow -- because the band's run model has no place for a non-run and
 * because that is where the band has put its inline pictures since §43. The
 * box is the drawing's own `wp:extent` in points, so a shape keeps the size
 * Word gives it whatever the band's width turns out to be.
 */
/** Whether a group member's words would be drawn mirrored by its flips and
 * its group's (the body's `memberTextMirrored`), so their box undoes it. */
export function bandMemberTextMirrored(
  shape: Pick<BandRichShape, "flipH" | "flipV">,
  member: { flipH?: boolean; flipV?: boolean },
): boolean {
  return (member.flipH === true) !== (shape.flipH === true) && (member.flipV === true) === (shape.flipV === true);
}

function appendBandGroupMembers(doc: Document, box: HTMLElement, shape: BandRichShape): void {
  const width = shape.groupCoordinateWidthEmu ?? shape.widthEmu;
  const height = shape.groupCoordinateHeightEmu ?? shape.heightEmu;
  box.style.position = "relative";
  const transforms = [
    shape.rotation ? `rotate(${shape.rotation}deg)` : "",
    shape.flipH ? "scaleX(-1)" : "",
    shape.flipV ? "scaleY(-1)" : "",
  ].filter(Boolean);
  const layer = doc.createElement("span");
  layer.style.position = "absolute";
  layer.style.inset = "0";
  if (transforms.length) layer.style.transform = transforms.join(" ");
  for (const member of shape.groupMembers ?? []) {
    const cell = doc.createElement("span");
    cell.className = "sheet-band-rich__shape-member";
    cell.setAttribute("data-band-shape-member", member.memberId);
    cell.style.position = "absolute";
    cell.style.left = `${(member.xEmu / width) * 100}%`;
    cell.style.top = `${(member.yEmu / height) * 100}%`;
    cell.style.width = `${(member.widthEmu / width) * 100}%`;
    cell.style.height = `${(member.heightEmu / height) * 100}%`;
    // A member carries the extent the package wrote (a rule is 635 EMU
    // tall); its box is at least as thick as the stroke it has to show.
    const stroke = `${Math.max(1, ptToPx(member.outlineWidthPt))}px`;
    cell.style.minWidth = stroke;
    cell.style.minHeight = stroke;
    const own = [
      member.rotation ? `rotate(${member.rotation}deg)` : "",
      member.flipH ? "scaleX(-1)" : "",
      member.flipV ? "scaleY(-1)" : "",
    ].filter(Boolean);
    if (own.length) cell.style.transform = own.join(" ");
    const svg = doc.createElementNS(SVG_NS, "svg") as SVGSVGElement;
    svg.setAttribute("viewBox", "0 0 100 100");
    svg.setAttribute("preserveAspectRatio", "none");
    svg.setAttribute("aria-hidden", "true");
    svg.style.width = "100%";
    svg.style.height = "100%";
    svg.style.display = "block";
    svg.style.overflow = "visible";
    paintShape(svg, normalizeDocumentShapeAttrs({ ...member, source: null }));
    cell.appendChild(svg);
    if (member.text) {
      const label = doc.createElement("span");
      label.className = "sheet-band-rich__shape-text";
      label.textContent = member.text;
      // Words are not mirrored with a flipped shape (Word; the body's group
      // member label, `memberTextMirrored`): one mirror of the member's or
      // the group's is undone, both cancel.
      if (bandMemberTextMirrored(shape, member)) label.style.transform = "scaleX(-1)";
      cell.appendChild(label);
    }
    layer.appendChild(cell);
  }
  box.appendChild(layer);
}

/** A chart's box on the page, in points: the size the package gave it. */
function bandChartSizeEmu(chart: BandRichChart): { widthEmu: number; heightEmu: number } {
  const host = chart.documentChart as { size?: { widthEmu?: unknown; heightEmu?: unknown } } | undefined;
  const source = chart.source as { widthEmu?: unknown; heightEmu?: unknown } | undefined;
  const pick = (...values: unknown[]) =>
    values.find((value): value is number => typeof value === "number" && Number.isFinite(value) && value > 0);
  return {
    widthEmu: pick(host?.size?.widthEmu, source?.widthEmu) ?? 5_486_400,
    heightEmu: pick(host?.size?.heightEmu, source?.heightEmu) ?? 3_127_200,
  };
}

/**
 * One chart a header/footer line holds, drawn by the ONE chart planner the
 * Docs chart node is drawn by (`documentChartPictureContent` -> the engine's
 * picture), at the size the package gave it. The box is placed first, so the
 * band is laid out at its height while the picture is on its way; a chart the
 * engine cannot draw leaves the box empty rather than drawing something else.
 */
export function buildBandChart(doc: Document, chart: BandRichChart, index: number): HTMLElement {
  const box = doc.createElement("span");
  box.className = "sheet-band-rich__chart";
  box.setAttribute("data-band-chart", String(index));
  box.style.display = "inline-block";
  // On the line's baseline, as Word sits every in-line drawing: the line
  // keeps its descent below the chart (Word 16.86: a 0.6" chart in a footer
  // line ends 2.1pt above the line's bottom). Bottom-aligned, each such line
  // lost its descent and the footer's lines above it sat that much lower.
  box.style.verticalAlign = "baseline";
  const { widthEmu, heightEmu } = bandChartSizeEmu(chart);
  box.style.width = pt(widthEmu);
  box.style.height = pt(heightEmu);
  const width = Math.max(1, Math.round(ptToPx((widthEmu / EMU_PER_PT))));
  const height = Math.max(1, Math.round(ptToPx((heightEmu / EMU_PER_PT))));
  const inputs = documentChartPlanInputs(chart, null);
  const { content } = documentChartPictureContent({
    attrs: normalizeChartBlockAttrs(chart) as unknown as ChartBlockAttrsV2,
    nativeLook: inputs.nativeLook,
    themeAccents: inputs.themeAccents,
    linkedValueNumberFormat: inputs.linkedValueNumberFormat,
  });
  requestChartPicture({ ...content, width, height })
    .then((picture) => {
      box.innerHTML = sanitizeChartPicture(picture.svg, picture.width, picture.height);
      const svg = box.querySelector("svg");
      if (svg) {
        svg.style.width = "100%";
        svg.style.height = "100%";
        svg.style.display = "block";
      }
    })
    .catch(() => {
      // Offline, busy or refused: the box keeps the chart's place.
    });
  return box;
}

export function buildBandShape(doc: Document, shape: BandRichShape, index: number): HTMLElement {
  const box = doc.createElement("span");
  box.className = "sheet-band-rich__shape";
  box.setAttribute("data-band-shape", String(index));
  box.setAttribute("data-band-shape-preset", shape.preset);
  box.style.display = "inline-block";
  // On the baseline, as a chart is (buildBandChart).
  box.style.verticalAlign = "baseline";
  box.style.width = pt(shape.widthEmu);
  box.style.height = pt(shape.heightEmu);
  // A zero axis (a straight line's `a:ext cy="0"`) is drawn in a box as thick
  // as its stroke, centred on the line, as the node view draws it: an svg zero
  // tall draws nothing (corpus R283-D12, R279-D17: a footer's 6 pt rule).
  // Any axis thinner than the stroke is widened so, as the painter's canvas is.
  const strokePx = Math.max(1, ptToPx(shape.outlineWidthPt));
  const heightPx = shape.heightEmu / EMU_PER_PX, widthPx = shape.widthEmu / EMU_PER_PX;
  if (heightPx < strokePx) {
    box.style.height = `${strokePx}px`;
    box.style.marginTop = `${-(strokePx - heightPx) / 2}px`;
  }
  if (widthPx < strokePx) {
    box.style.width = `${strokePx}px`;
    box.style.marginLeft = `${-(strokePx - widthPx) / 2}px`;
  }
  const svg = doc.createElementNS(SVG_NS, "svg") as SVGSVGElement;
  svg.setAttribute("viewBox", "0 0 100 100");
  svg.setAttribute("preserveAspectRatio", "none");
  svg.setAttribute("aria-hidden", "true");
  svg.style.width = "100%";
  svg.style.height = "100%";
  svg.style.display = "block";
  if (shape.rotation) {
    // Degrees, as `DocumentShapeAttrs.rotation` states them, about the box
    // centre -- the same origin the node view rotates a body shape about.
    svg.style.transform = `rotate(${shape.rotation}deg)`;
  }
  if (shape.groupMembers?.length) {
    // A GROUP draws its members, each at its own place in the group's child
    // space, by the same painter -- a drawing canvas's logo pieces, its rule.
    appendBandGroupMembers(doc, box, shape);
    return box;
  }
  paintShape(svg, shape);
  box.appendChild(svg);
  if (shape.text) {
    const label = doc.createElement("span");
    label.className = "sheet-band-rich__shape-text";
    label.textContent = shape.text;
    box.appendChild(label);
  }
  return box;
}

/** A band line holding only empty typing slots: no words, field, list
 * marker, picture, drawing or box. */
function isEmptySlotLine(block: BandRichBlock, context: BandTokenContext): boolean {
  return (
    block.type === "paragraph" &&
    block.runs.length > 0 &&
    block.runs.every((run) => run.frame === undefined && !runText(run, context)) &&
    !block.images?.length &&
    !block.shapes?.length &&
    !block.charts?.length &&
    !block.frames?.length &&
    !block.wordArt?.length &&
    !block.list
  );
}

export function buildBandRichDom(
  doc: Document,
  blocks: BandRichBlock[],
  context: BandTokenContext,
  /** A cell's or a box's own flow, never a band story of its own. */
  innerFlow = false,
): HTMLElement {
  const root = doc.createElement("div");
  root.className = "sheet-band-rich";
  if (
    "wordCompatibilityMode" in context &&
    !(typeof context.wordCompatibilityMode === "number" && context.wordCompatibilityMode >= 15)
  ) {
    root.setAttribute("data-band-keep-on-paper", "");
  }
  // The list markers of this container's lines (a table interrupts a list,
  // and each cell counts its own), as Word counts them in a header.
  const markers = bandListMarkers(blocks.map((block) => (block.type === "paragraph" ? block : null)));
  // Each paragraph block's element, for a box anchored by an earlier one.
  const paragraphByBlock = new Map<number, HTMLElement>();
  for (const [blockIndex, block] of blocks.entries()) {
    // Every kind `BAND_RICH_BLOCK_TYPES` names has a branch below, and the
    // COMPILER is what says so: the closing `else` narrows `block` to `never`
    // only while the branches cover the whole union, and that union is held to
    // the one list. Adding a kind without drawing it does not build, instead
    // of a real document's header being drawn as something else in silence.
    const kind: BandRichBlockType = block.type;
    void kind;
    if (block.type === "paragraph") {
      const paragraph = doc.createElement("div");
      paragraph.className = "sheet-band-rich__p";
      paragraphByBlock.set(blockIndex, paragraph);
      for (const [key, value] of Object.entries(documentParagraphBordersHtmlAttributes(block.paragraphBorders))) {
        paragraph.setAttribute(key, value);
      }
      paragraph.style.textAlign = block.align ?? "left";
      if (block.textFrame && block.textFrame.wrap !== "notBeside") {
        // A framed line stands out of the band's flow, at the side its frame
        // names, and the lines after it run beside it (Word's text frame;
        // corpus R21-D04's framed page number takes no line of its own).
        const side = block.textFrame.xAlign === "right" || block.textFrame.xAlign === "outside" ? "right" : "left";
        paragraph.style.float = side;
        if (block.textFrame.w) paragraph.style.width = `${block.textFrame.w / 15}px`;
        if (block.textFrame.hSpace) paragraph.style.setProperty(`margin-${side === "left" ? "right" : "left"}`, `${block.textFrame.hSpace / 15}px`);
      }
      const runSize = bandStyleRunSizePx(context.styleDefinitions, block.docxStyleId);
      if (runSize !== null) paragraph.style.fontSize = `${Math.round(runSize * 1000) / 1000}px`;
      // A run that states no bold or italic takes its paragraph style's, as
      // Word draws it (corpus R158-D22: "OFFICIAL SENSITIVE " in a bold
      // Header style).
      const toggles = bandStyleRunToggles(context.styleDefinitions, block.docxStyleId);
      if (toggles.bold) paragraph.style.fontWeight = "700";
      if (toggles.italic) paragraph.style.fontStyle = "italic";
      // Tabs reach their stops as Word lays them out once the band is on
      // screen (`layoutBandTabs`); until then, and on paper that is never
      // measured, a uniform width from the first stop stands in.
      if (block.tabStops?.length) {
        paragraph.style.whiteSpace = "pre-wrap";
        paragraph.style.tabSize = `${Math.max(1, Math.round((block.tabStops[0].pos / 1440) * 96))}px`;
      }
      if (block.runs.some((run) => run.frame === undefined && run.text?.includes("\t"))) {
        paragraph.setAttribute("data-band-tabs", JSON.stringify(block.tabStops ?? []));
      }
      // A negative indent hangs the line into the margin (a letterhead title
      // at w:left="-360"): padding cannot be negative, so the box itself
      // reaches out by a negative margin (layoutBandTabs measures from the
      // text margin either way).
      if (block.indentLeftPt !== undefined) {
        if (block.indentLeftPt < 0) paragraph.style.marginLeft = `${block.indentLeftPt}pt`;
        else paragraph.style.paddingLeft = `${block.indentLeftPt}pt`;
      }
      if (block.indentRightPt !== undefined) {
        if (block.indentRightPt < 0) paragraph.style.marginRight = `${block.indentRightPt}pt`;
        else paragraph.style.paddingRight = `${block.indentRightPt}pt`;
      }
      if (block.firstLineIndentPt !== undefined) paragraph.style.textIndent = `${block.firstLineIndentPt}pt`;
      const marker = markers[blockIndex];
      if (marker && block.list) {
        // Word's list item: the text at the level's indent, the marker hung
        // into the first line's negative indent and followed by a tab to the
        // text (the level's `w:suff` default).
        paragraph.setAttribute("data-band-list", block.list.kind);
        paragraph.setAttribute("data-band-list-level", String(block.list.level));
        paragraph.style.paddingLeft = `${marker.indentLeftPt}pt`;
        paragraph.style.textIndent = `${marker.firstLineIndentPt}pt`;
        const label = doc.createElement("span");
        label.className = "sheet-band-rich__marker";
        label.textContent = marker.text;
        label.style.display = "inline-block";
        label.style.textIndent = "0";
        label.style.whiteSpace = "pre";
        if (marker.firstLineIndentPt < 0) label.style.minWidth = `${-marker.firstLineIndentPt}pt`;
        if (marker.font) label.style.fontFamily = `"${marker.font}"`;
        // The tab after the number: a number wider than the hanging indent
        // moves the text on to the next stop (`layoutBandListMarkers`).
        label.setAttribute("data-band-marker-tab", JSON.stringify({
          start: Math.round((marker.indentLeftPt + marker.firstLineIndentPt) * 20),
          body: Math.round(marker.indentLeftPt * 20),
          stops: (block.tabStops ?? []).map((stop) => stop.pos),
        }));
        paragraph.appendChild(label);
      }
      applyBandLineHeight(paragraph, block, block.runs.filter((run) => run.frame === undefined));
      const sameStyle = (other: BandRichBlock | undefined) => Boolean(block.contextualSpacing && other?.type === "paragraph" && (other.docxStyleId ?? "Normal") === (block.docxStyleId ?? "Normal"));
      if (sameStyle(blocks[blockIndex - 1])) paragraph.style.marginTop = "0pt";
      else if (block.spaceBeforePt !== undefined) paragraph.style.marginTop = `${block.spaceBeforePt}pt`;
      if (sameStyle(blocks[blockIndex + 1])) paragraph.style.marginBottom = "0pt";
      else if (block.spaceAfterPt !== undefined) paragraph.style.marginBottom = `${block.spaceAfterPt}pt`;
      // An anchored picture or shape is off the line: a paragraph holding only
      // those is still a blank line of its own height, as in Word.
      if (!block.runs.length && !block.images?.some((image) => !image.floatAnchor) && !block.shapes?.some((shape) => !shape.floatAnchor) && !block.charts?.length && !block.frames?.some((frame) => frame.inline)) paragraph.appendChild(doc.createElement("br"));
      const floated = placeableBandFrames(block);
      const inlineFrames = (block.frames ?? []).filter((frame) => frame.inline);
      if (inlineFrames.length) {
        appendLineWithInlineFrames(doc, paragraph, block, floated, inlineFrames, context);
      } else {
        appendRuns(
          paragraph,
          floated.size ? block.runs.filter((run) => run.frame === undefined || !floated.has(run.frame)) : block.runs,
          context,
        );
        // An empty line of a band story (its typing slots only) is a blank
        // line of its own height, as Word draws it: empty spans alone make no
        // line box. Word 16.86: R28-D13's header ends in an empty 10 pt
        // paragraph and the body starts 12.6 pt lower for it (124.4 pt);
        // R27-D13's footer opens with one, drawn as a line at 710 pt
        // (R54-D01: a blank header pushes the body down). R27-D13's page
        // count is kept by the footer's reach, not by dropping its line
        // (measureFooterBandOverflowBySheet). A cell's or a box's own flow
        // keeps its empty lines as it drew them.
        if (!innerFlow && isEmptySlotLine(block, context)) paragraph.appendChild(doc.createElement("br"));
        applyBandRunLineHeights(paragraph, block);
      }
      for (const frame of floated.values()) {
        const box = buildBandFrame(doc, frame, block.runs.filter((run) => run.frame === frame.index), context);
        if (frameVerticalFromParagraph(frame)) {
          // Paragraph/line origins are the anchor paragraph's own top: the
          // frame is its positioned child. The paragraph stays zero-effect
          // otherwise (no z-index), so frames still share one z-order. A box
          // a linked chain shows its story in is anchored by an EARLIER
          // paragraph (`anchorParagraph` blocks back): it is that one's child.
          const host =
            frame.anchorParagraph !== undefined
              ? paragraphByBlock.get(blockIndex + frame.anchorParagraph) ?? paragraph
              : paragraph;
          host.style.position = "relative";
          host.appendChild(box);
        } else {
          // Page/margin origins are measured from the band box, which is the
          // positioned ancestor of the (unpositioned) band root.
          root.appendChild(box);
        }
      }
      // A watermark is measured from the BAND box, never from the anchor
      // paragraph: it spans the whole text margin box and the paragraph it
      // happens to be written in is only where Word put the markup.
      for (const art of block.wordArt ?? []) {
        root.appendChild(buildBandWordArt(doc, art));
      }
      for (const [shapeIndex, shape] of (block.shapes ?? []).entries()) {
        const drawing = buildBandShape(doc, shape, shapeIndex);
        const anchor = shape.floatAnchor;
        if (!anchor) {
          paragraph.appendChild(drawing);
          continue;
        }
        // An anchored drawing is placed exactly as a frame is, and paints
        // under the text when Word's `behindDoc` says so.
        drawing.style.position = "absolute";
        drawing.style.display = "block";
        drawing.style.zIndex = anchor.behindDoc ? "-1" : "auto";
        placeBandBox(drawing, anchor.horizontal, anchor.vertical, shape.widthEmu, shape.heightEmu);
        markBandWrap(drawing, anchor);
        markBandKeep(drawing, anchor);
        if (anchor.vertical.relativeFrom === "paragraph" || anchor.vertical.relativeFrom === "line") {
          paragraph.style.position = "relative";
          paragraph.appendChild(drawing);
        } else {
          root.appendChild(drawing);
        }
      }
      for (const [chartIndex, chart] of (block.charts ?? []).entries()) {
        paragraph.appendChild(buildBandChart(doc, chart, chartIndex));
      }
      for (const image of block.images ?? []) {
        const img = doc.createElement("img");
        drawDocumentPicture(img, image.src);
        img.alt = "";
        // A header or footer picture is drawn, never dragged away natively
        // (lib/native-drag-guard).
        img.draggable = false;
        img.className = "sheet-band-rich__img";
        // At its exact extent: a whole-pixel picture stood a logo 0.3 pt
        // taller than Word's, and the body under the band with it.
        if (image.widthEmu) {
          img.style.width = `${Math.max(1, Math.round((image.widthEmu / EMU_PER_PX) * 1000) / 1000)}px`;
        }
        if (image.heightEmu) {
          img.style.height = `${Math.max(1, Math.round((image.heightEmu / EMU_PER_PX) * 1000) / 1000)}px`;
        }
        const anchor = image.floatAnchor;
        if (!anchor) {
          paragraph.appendChild(img);
          continue;
        }
        // An anchored picture is placed exactly as an anchored band shape is.
        img.style.position = "absolute";
        img.style.display = "block";
        img.style.maxWidth = "none";
        img.style.zIndex = anchor.behindDoc ? "-1" : "auto";
        // The band is drawn over the page's text: a picture Word puts behind
        // the text makes the overlay multiply onto the page (note-editor.css),
        // so the text shows through as if it were above (R21-D04's picture
        // watermark hid its first page).
        if (anchor.behindDoc) img.setAttribute("data-band-behind-text", "");
        // Word's Washout, as the document's own picture watermark draws it.
        img.style.opacity = image.washout ? "0.3" : "";
        placeBandBox(img, anchor.horizontal, anchor.vertical, image.widthEmu ?? anchor.widthEmu, image.heightEmu ?? anchor.heightEmu);
        markBandWrap(img, anchor);
        markBandKeep(img, anchor);
        if (anchor.vertical.relativeFrom === "paragraph" || anchor.vertical.relativeFrom === "line") {
          paragraph.style.position = "relative";
          paragraph.appendChild(img);
        } else {
          root.appendChild(img);
        }
      }
      // A line holding in-line pictures and no text is as tall as its
      // tallest picture: Word adds no descent under it (Word 16.86, corpus
      // R158-D19's header logo 75.45 pt: the next line starts 75.43 pt
      // down, where the band drew it 3 pt lower and every page's body
      // started late). With text beside the picture the descent stays.
      if (
        (block.images ?? []).some((image) => !image.floatAnchor) &&
        !block.runs.some((run) => run.frame === undefined && (run.text ?? "").trim() !== "")
      ) {
        paragraph.setAttribute("data-band-picture-line", "");
      }
      paintPhysicalParagraphBorderEdges(paragraph, block.paragraphBorders);
      if (block.wordHorizontalRule) {
        // Word's horizontal line the line hosts, drawn as a body paragraph's
        // is (`wordHorizontalRule`): at the bottom of its own line.
        const ruleAttributes = ruleAttrs(block.wordHorizontalRule, { mode: "bottom" }, block.runs.length > 0);
        paragraph.classList.add(ruleAttributes.class);
        for (const [key, value] of Object.entries(ruleAttributes)) {
          if (key === "class") continue;
          if (key === "style") paragraph.style.cssText += `; ${value}`;
          else paragraph.setAttribute(key, value);
        }
      }
      root.appendChild(paragraph);
    } else if (block.type === "table") {
      const table = doc.createElement("table");
      table.className = "sheet-band-rich__table";
      // A table Word sized as a PERCENTAGE fills the width available to it,
      // and its columns divide that width by their own percentages. The
      // absolute `w:tblGrid` such a table also carries is the grid cached at
      // the width it had when saved -- laying it out at those pixels made a
      // header table wider than the band's content width, pushing its cells
      // past the right margin and wrapping the last one.
      const relative = block.columnWidthsPct;
      const widths = block.nativeColumnWidthsTwips?.map(width => width / 15) ?? block.columnWidthsPx;
      if (relative && relative.length > 0) {
        const grid = doc.createElement("colgroup");
        for (const share of relative) {
          const col = doc.createElement("col");
          col.style.width = `${share}%`;
          grid.appendChild(col);
        }
        table.appendChild(grid);
        table.style.width = "100%";
        table.style.tableLayout = "fixed";
      } else if (widths) {
        const grid = doc.createElement("colgroup");
        for (const width of widths) { const col = doc.createElement("col"); col.style.width = `${width}px`; grid.appendChild(col); }
        table.appendChild(grid);
        table.style.width = `${widths.reduce((total, width) => total + width, 0)}px`;
      }
      if (block.tableAlign === "center") { table.style.marginLeft = "auto"; table.style.marginRight = "auto"; }
      else if (block.tableAlign === "right") table.style.marginLeft = "auto";
      const covered = new Set<string>();
      const byCoord = new Map<string, (typeof block.cells)[number]>();
      for (const cell of block.cells) {
        byCoord.set(`${cell.row}:${cell.col}`, cell);
        const colspan = cell.colspan ?? 1;
        const rowspan = cell.rowspan ?? 1;
        for (let r = 0; r < rowspan; r += 1) {
          for (let c = 0; c < colspan; c += 1) {
            if (r || c) covered.add(`${cell.row + r}:${cell.col + c}`);
          }
        }
      }
      if (block.borders || block.cells.some((cell) => cell.borders)) table.style.borderCollapse = "collapse";
      for (let row = 0; row < block.rows; row += 1) {
        const tr = doc.createElement("tr");
        const height = block.rowHeights?.[row];
        if (height) tr.style.height = `${height.twips / 15}px`;
        for (let col = 0; col < block.cols; col += 1) {
          const key = `${row}:${col}`;
          if (covered.has(key) && !byCoord.has(key)) continue;
          const cell = byCoord.get(key);
          const td = doc.createElement("td");
          td.style.whiteSpace = "pre-wrap";
          if (cell?.colspan) td.colSpan = cell.colspan;
          if (cell?.rowspan) td.rowSpan = cell.rowspan;
          if (cell?.align) td.style.textAlign = cell.align;
          if (cell?.verticalAlign) td.style.verticalAlign = cell.verticalAlign === "center" ? "middle" : cell.verticalAlign;
          if (cell?.shading) td.style.backgroundColor = cell.shading;
          if (block.borders || cell?.borders) {
            // Each side: the cell's own border, else the table's edge or
            // inside line the side lies on (Word's own precedence).
            const lastRow = row + (cell?.rowspan ?? 1) - 1 === block.rows - 1;
            const lastCol = col + (cell?.colspan ?? 1) - 1 === block.cols - 1;
            const sides = {
              top: cell?.borders?.top ?? (row === 0 ? block.borders?.top : block.borders?.insideH),
              bottom: cell?.borders?.bottom ?? (lastRow ? block.borders?.bottom : block.borders?.insideH),
              left: cell?.borders?.left ?? (col === 0 ? block.borders?.left : block.borders?.insideV),
              right: cell?.borders?.right ?? (lastCol ? block.borders?.right : block.borders?.insideV),
            };
            for (const [side, border] of Object.entries(sides)) {
              const css = bandBorderCss(border);
              if (css) td.style.setProperty(`border-${side}`, css);
            }
          }
          if (block.rowHeights?.[row]?.rule === "exact") td.style.overflow = "hidden";
          for (const side of ["top", "right", "bottom", "left"] as const) {
            const margin = cell?.cellMargin?.[side];
            if (margin === undefined) continue;
            // Word measures a cell's left and right margins from the grid
            // line, its border standing in the margin, never in the text; a
            // collapsed CSS border takes half its width before the padding
            // starts (the body tables' rule, tableBorders.ts). Charged twice,
            // R161-D07's 70.95 pt header cell lost 0.75 pt, "Regulation
            // 22(1)" (63.75 pt of 63.75) wrapped, and every page's body
            // started 11 pt low.
            const used = side === "left" || side === "right" ? Number.parseFloat(td.style.getPropertyValue(`border-${side}-width`)) || 0 : 0;
            const padding = used > 0 && margin >= used / 2 ? Number((margin - used / 2).toFixed(3)) : margin;
            td.style.setProperty(`padding-${side}`, `${padding}px`);
          }
          if (cell?.paragraphs?.length) {
            const paragraphs = buildBandRichDom(doc, cell.paragraphs, context, true);
            td.append(...paragraphs.childNodes);
          } else if (cell?.runs?.length) appendRuns(td, cell.runs, context);
          else if (!cell?.tables?.length) td.textContent = cell?.text ?? "";
          // A table nested in the cell stands after the first `at` of the
          // cell's paragraphs, a block of the cell like they are (R13-D31).
          for (const nested of cell?.tables ?? []) {
            const drawn = buildBandRichDom(doc, [nested.table], context);
            const before = td.querySelectorAll(":scope > .sheet-band-rich__p")[nested.at] ?? null;
            td.insertBefore(drawn.firstChild ?? doc.createTextNode(""), before);
          }
          tr.appendChild(td);
        }
        table.appendChild(tr);
      }
      if (block.floatPosition) placeBandFloatingTable(table, block);
      root.appendChild(table);
    } else {
      const undrawn: never = block;
      void undrawn;
    }
  }
  return root;
}


/**
 * The drawn width of inline text, in unscaled px, from its client rects: the
 * extent of each line it occupies, summed. A run that has wrapped still
 * measures its whole width -- `offsetWidth` would give the bounding box of
 * its fragments, and a tab sized from that keeps the text wrapped for good.
 * Rects of nested boxes on one line (a run's spans, a mark inside a mark)
 * count once. Null when there are no rects (nothing laid out).
 */
export function inlineTextWidthPx(rects: Iterable<DOMRectReadOnly>, scale = 1): number | null {
  const lines: { top: number; bottom: number; left: number; right: number }[] = [];
  for (const rect of rects) {
    if (rect.width <= 0 && rect.height <= 0) continue;
    const line = lines[lines.length - 1];
    if (line && rect.top < line.bottom && rect.bottom > line.top) {
      line.left = Math.min(line.left, rect.left);
      line.right = Math.max(line.right, rect.right);
      line.top = Math.min(line.top, rect.top);
      line.bottom = Math.max(line.bottom, rect.bottom);
    } else {
      lines.push({ top: rect.top, bottom: rect.bottom, left: rect.left, right: rect.right });
    }
  }
  if (!lines.length) return null;
  return lines.reduce((total, line) => total + (line.right - line.left), 0) / (scale || 1);
}

/** The positional tab (`w:ptab`) a band tab span draws, if it is one. */
function bandPositionalTab(tab: HTMLElement): PositionalTab | null {
  const raw = tab.getAttribute("data-band-ptab");
  if (!raw) return null;
  try {
    return normalizePositionalTab(JSON.parse(raw));
  } catch {
    return null;
  }
}

/** Null means the overlay still has no usable inline extent. A zero-width
 * mounting frame must not wrap every word and become authored body padding. */
/**
 * Size every tab of the painted bands under `root` so it reaches its stop as
 * Word lays it out (`wordTabAdvances`): the text after a centre stop centres
 * on it, after a right stop ends at it, and after a positional tab it aligns
 * at that tab's margin or indent point. Needs a laid-out DOM; a band that is
 * not on screen keeps its first-stop approximation. Idempotent -- call it
 * again when fonts load or the band resizes.
 */
export function layoutBandTabs(root: ParentNode): void {
  // The markers first: a text tab after a list number measures from where
  // the number's own tab ends.
  layoutBandListMarkers(root);
  for (const paragraph of root.querySelectorAll<HTMLElement>(".sheet-band-rich__p[data-band-tabs]")) {
    if (!paragraph.offsetWidth) continue;
    let stops: ReturnType<typeof normalizeTabStops> = null;
    try {
      stops = normalizeTabStops(JSON.parse(paragraph.getAttribute("data-band-tabs") ?? "[]"));
    } catch {
      stops = null;
    }
    // A stop past the right margin (a letterhead date right-aligned at
    // 6.75in on a 6.5in text column) is still reached: Word runs the line
    // into the margin up to the stop instead of wrapping it (R01-D04). The
    // box reaches out to the furthest stop.
    {
      const furthestPx = Math.max(0, ...(stops ?? []).map((stop) => stop.pos / 15));
      const own = paragraph.ownerDocument.defaultView?.getComputedStyle(paragraph);
      const reachLeft = Math.min(0, Number.parseFloat(own?.marginLeft ?? "0") || 0);
      const currentRight = Math.min(0, Number.parseFloat(own?.marginRight ?? "0") || 0);
      // The margin the line states itself, kept across passes.
      const stated = paragraph.dataset.bandStatedMarginRight ?? paragraph.style.marginRight;
      paragraph.dataset.bandStatedMarginRight = stated;
      const statedRight = Math.min(0, Number.parseFloat(stated) || 0);
      // Margin to margin, and where the box ends with the stated margin.
      const columnPx = paragraph.clientWidth + reachLeft + currentRight;
      // No tolerance: a stop a hair past the margin (corpus R158-D22's
      // right stop at 10206 twips on a 10204-twip column, 0.13px over) still
      // overflows the box by that hair, and the page number after it wrapped
      // onto a line of its own at the left.
      paragraph.style.marginRight = furthestPx > columnPx - statedRight
        ? `${columnPx - furthestPx}px`
        : stated;
    }
    const tabs: HTMLElement[] = [];
    const positional: (PositionalTab | null)[] = [];
    const runs: HTMLElement[][] = [[]];
    for (const child of paragraph.children) {
      const element = child as HTMLElement;
      if (element.classList.contains("sheet-band-rich__tab")) {
        tabs.push(element);
        positional.push(bandPositionalTab(element));
        runs.push([]);
      } else if (element.hasAttribute("data-band-run") || element.classList.contains("sheet-band-rich__marker")) {
        // A list item's marker is drawn on the line before its text, from
        // the line's first indent: the text after it starts where it ends.
        runs[runs.length - 1].push(element);
      }
    }
    if (!tabs.length) continue;
    // Each stretch between tabs as it is drawn -- its own fonts, fractional
    // widths, and all of it even where an earlier layout wrapped it.
    const scale = paragraph.getBoundingClientRect().width / paragraph.offsetWidth || 1;
    const segments = runs.map((spans) => ({
      widthPx: spans.reduce((total, span) => {
        const rects = typeof span.getClientRects === "function" ? Array.from(span.getClientRects()) : [];
        return total + (inlineTextWidthPx(rects, scale) ?? span.offsetWidth);
      }, 0),
    }));
    const style = paragraph.ownerDocument.defaultView?.getComputedStyle(paragraph);
    // The paragraph box spans the band's text column, margin to margin; its
    // padding is the paragraph's indents (what a positional tab set relative
    // to the indents measures from), and a negative indent is a negative
    // margin reaching into the page margin. Stops are measured from the text
    // margin, so the box's own reach is folded back in.
    const reachLeftPx = Math.min(0, Number.parseFloat(style?.marginLeft ?? "0") || 0);
    const reachRightPx = Math.min(0, Number.parseFloat(style?.marginRight ?? "0") || 0);
    const indentLeftPx = (Number.parseFloat(style?.paddingLeft ?? "0") || 0) + reachLeftPx;
    const startPx = indentLeftPx + (Number.parseFloat(style?.textIndent ?? "0") || 0);
    const line = {
      widthPx: paragraph.clientWidth + reachLeftPx + reachRightPx,
      indentLeftPx,
      indentRightPx: (Number.parseFloat(style?.paddingRight ?? "0") || 0) + reachRightPx,
    };
    const advances = wordTabAdvances({ segments, stops, startPx, positional, line });
    let pen = startPx + segments[0].widthPx;
    for (const [index, tab] of tabs.entries()) {
      const advance = advances[index] ?? 0;
      const leader = wordTabLeader(stops, pen, positional[index]);
      tab.style.display = "inline-block";
      tab.style.width = `${advance}px`;
      tab.style.tabSize = "0";
      tab.style.whiteSpace = "pre";
      if (leader !== "none") tab.setAttribute("data-leader", leader);
      else tab.removeAttribute("data-leader");
      pen += advance + segments[index + 1].widthPx;
    }
  }
}

/**
 * Size each painted list number's box to where Word's tab after it lands:
 * the hanging indent when the number fits in it, else the line's next own
 * stop or default stop (`numberingSuffixTabStopTwips`, as the body's list
 * markers do). Word 16 for Mac, wp-fix-a/share.docx: "<1.>" (wider than the
 * 0.25 in hanging) puts the Header-style line's text at its 3.25 in centre
 * stop and a plain line's at the 1 in default stop. Needs a laid-out DOM.
 */
export function layoutBandListMarkers(root: ParentNode): void {
  for (const label of root.querySelectorAll<HTMLElement>(".sheet-band-rich__marker[data-band-marker-tab]")) {
    const paragraph = label.parentElement;
    if (!paragraph?.offsetWidth) continue;
    let tab: { start: number; body: number; stops: number[] } | null = null;
    try {
      const raw = JSON.parse(label.getAttribute("data-band-marker-tab") ?? "null") as unknown;
      if (raw && typeof raw === "object") {
        const { start, body, stops } = raw as Record<string, unknown>;
        if (Number.isFinite(start) && Number.isFinite(body) && Array.isArray(stops) && stops.every((stop) => Number.isFinite(stop))) {
          tab = { start: start as number, body: body as number, stops: stops as number[] };
        }
      }
    } catch {
      tab = null;
    }
    if (!tab) continue;
    // The number's own drawn width, whatever an earlier pass sized its box to.
    const sized = { width: label.style.width, minWidth: label.style.minWidth };
    label.style.width = "";
    label.style.minWidth = "0";
    const scale = paragraph.getBoundingClientRect().width / paragraph.offsetWidth || 1;
    const widthPx = label.getBoundingClientRect().width / scale;
    const tabSize = Number.parseFloat(
      paragraph.ownerDocument.defaultView?.getComputedStyle(paragraph).getPropertyValue("--doc-default-tab-size") ?? "",
    );
    const next = numberingSuffixTabStopTwips(
      tab.start + widthPx * 15,
      tab.body,
      tab.stops,
      Number.isFinite(tabSize) && tabSize >= 0 ? tabSize * 15 : 720,
    );
    if (!Number.isFinite(widthPx) || widthPx <= 0 || next <= tab.start || next > 31_680) {
      label.style.width = sized.width;
      label.style.minWidth = sized.minWidth;
      continue;
    }
    label.style.minWidth = `${(next - tab.start) / 15}px`;
  }
}

/** Each sheet's own header overflow past its top margin (px, whole), by the
 * band box's sheet index; a sheet whose header fits is absent. Word starts
 * the body at the larger of the top margin and the header distance plus the
 * header's height (its last paragraph's space after included), with no gap
 * between: Word 16 for Mac, a one-line Cambria 11pt header at 24/30/54pt
 * under a 36pt margin starts the body 36.9/42.9/66.9pt down, and at 18pt
 * (30.9pt) at the margin. A 10px pad here put a landscape section's body
 * 7pt below Word's (header-footer parity #24). */
/** How far a band's distance from the page edge already reaches past its
 * margin (px). The band box spans the room between the two and is empty
 * when the distance is the larger: Word still starts the band at its
 * distance and the body below the band, so that reach is part of the push
 * (corpus R54-D01: header at 35.4 pt over a 14.2 pt top margin). */
function bandDistanceBeyondMargin(box: HTMLElement, edge: "top" | "bottom"): number {
  const style = box.ownerDocument.defaultView?.getComputedStyle(box);
  if (!style) return 0;
  const distance = Number.parseFloat(style[edge]);
  const margin = Number.parseFloat(style.getPropertyValue(edge === "top" ? "--page-margin-top" : "--page-margin-bottom"));
  if (!Number.isFinite(distance) || !Number.isFinite(margin)) return 0;
  return Math.max(0, distance - margin);
}

export function measureHeaderBandOverflowBySheet(overlay: Element): Record<number, number> {
  const out: Record<number, number> = {};
  for (const box of overlay.querySelectorAll<HTMLElement>(".sheet-band-box__header")) {
    const content = box.firstElementChild as HTMLElement | null;
    const sheet = Number(box.closest("[data-band-sheet-index]")?.getAttribute("data-band-sheet-index"));
    if (!content || !Number.isInteger(sheet) || box.clientWidth <= 0 || content.offsetWidth <= 0) continue;
    const extra = Math.min(Math.max(0, Math.round(content.offsetHeight - box.clientHeight + bandDistanceBeyondMargin(box, "top"))), 480);
    if (extra > 0) out[sheet] = extra;
  }
  return out;
}

/** Each sheet's own footer overflow past its bottom margin (px, whole), by
 * the band box's sheet index; a sheet whose footer fits is absent. Word ends
 * the body at the higher of the bottom margin and the footer's top (page
 * height less the footer distance and the footer's height), with no gap:
 * Word 16 for Mac, a four-line Cambria 11pt footer at 30/54pt over a 36pt
 * margin ends the body above 710/686pt on a 792pt page. Without it a footer
 * taller than its margin drew over the body's last lines (R36-D29: six
 * footer lines, three-column text under them). */
export function measureFooterBandOverflowBySheet(overlay: Element): Record<number, number> {
  const out: Record<number, number> = {};
  for (const box of overlay.querySelectorAll<HTMLElement>(".sheet-band-box__footer")) {
    const content = box.firstElementChild as HTMLElement | null;
    const sheet = Number(box.closest("[data-band-sheet-index]")?.getAttribute("data-band-sheet-index"));
    if (!content || !Number.isInteger(sheet) || box.clientWidth <= 0 || content.offsetWidth <= 0) continue;
    const extra = Math.min(Math.max(0, Math.round(footerReach(content) - box.clientHeight + bandDistanceBeyondMargin(box, "bottom"))), 480);
    if (extra > 0) out[sheet] = extra;
  }
  return out;
}

/** How far a footer's lines reach up from its end (px): the footer without
 * its first paragraph's space before. Word lifts the body only for the
 * footer's lines: R27-D13's footer (an empty first paragraph with a top
 * border, 10 pt before and after, then the page number) has its border at
 * 706.56 pt under a 702 pt margin, and the body runs on to the margin
 * (Word 16.86, 93 pages); counting the 10 pt before lifted every page's body
 * and added three to five pages. */
function footerReach(content: HTMLElement): number {
  const first = Array.from(content.children).find((child): child is HTMLElement =>
    child instanceof HTMLElement && getComputedStyle(child).position !== "absolute" && child.offsetHeight > 0);
  const contentRect = content.getBoundingClientRect();
  if (!first || !(content.offsetHeight > 0) || !(contentRect.height > 0)) return content.offsetHeight;
  const scale = contentRect.height / content.offsetHeight;
  const reach = (contentRect.bottom - first.getBoundingClientRect().top) / scale;
  return Number.isFinite(reach) && reach > 0 ? Math.min(reach, content.offsetHeight) : content.offsetHeight;
}

export function measureHeaderBandOverflow(overlay: Element): number | null {
  let extra = 0;
  for (const box of overlay.querySelectorAll<HTMLElement>(".sheet-band-box__header")) {
    const content = box.firstElementChild as HTMLElement | null;
    if (!content || box.clientWidth <= 0 || content.offsetWidth <= 0) return null;
    extra = Math.max(extra, content.offsetHeight - box.clientHeight + bandDistanceBeyondMargin(box, "top"));
  }
  return Math.min(Math.max(0, Math.round(extra)), 480);
}
