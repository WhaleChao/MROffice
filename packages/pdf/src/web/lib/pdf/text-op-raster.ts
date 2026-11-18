/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { layoutPdfTextBoxVerticalColumns, placePdfTextBoxVerticalColumns } from "./text-box-vertical";
import { pdfWordArtCanvasPainter } from "./text-box-wordart-canvas";
import type { FlattenImageBytes } from "./flatten-apply";
import type { TextOp } from "@mrnote/core/web/lib/pdf/flatten.ts";
import { CJK_SANS_STACK, CJK_SERIF_STACK, FONT_STACK, needsEastAsianFace, withEastAsianFallback } from "./font-resolve";
import { ensurePdfEastAsianFace } from "./east-asian-face";
import { pdfDraftItalic } from "./draft-italic";
import { PDF_NATIVE_SYNTHETIC_ITALIC_SKEW } from "./native-synthetic-style";
import {
  layoutBrowserPdfTextRuns,
  type PdfBrowserTextLayoutLine,
  type PdfBrowserTextSpacing,
} from "./browser-text-layout";
import { createPdfPortableTextLayout, type PdfPortableTextLayout } from "./portable-text-layout";
import { markPdfTextLatency } from "./text-edit-latency";
import { pdfTextFacesReady } from "./text-faces-ready";
import { justifyPdfTextLine, pdfTextLineParagraphEnds } from "./text-justify";
import { pdfTextBoxColumnRows, pdfTextBoxColumnWidth, pdfTextBoxVerticalOffset } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import {
  layoutPositionedPdfTextRuns,
  PDF_TEXTBOX_LINE_HEIGHT_RATIO,
  pdfTextRasterBackingSize,
} from "./text-run-layout";
import {
  pdfTextBoxRunCssFont,
  pdfTextBoxRunItalicPresentation,
  pdfTextBoxRunRise,
  pdfTextBoxRunRowGeometry,
  pdfTextBoxRunGlyphSize,
} from "./text-box-runs";
import { pageTextWidth } from "@mrnote/core/web/lib/text-engine/page-adapter.ts";

const MONO_CJK_EXPORT_STACK =
  '"Courier New", Courier, "Microsoft YaHei", "PingFang SC", SimHei, monospace';

function textExportFontFamily(font: TextOp["font"]): string {
  if (font === "serif") return CJK_SERIF_STACK;
  if (font === "mono") return MONO_CJK_EXPORT_STACK;
  return CJK_SANS_STACK;
}

function canvasToPngBytes(
  canvas: HTMLCanvasElement,
): Promise<Uint8Array | null> {
  return new Promise((resolve) => {
    canvas.toBlob(async (blob) => {
      if (!blob) {
        resolve(null);
        return;
      }
      resolve(new Uint8Array(await blob.arrayBuffer()));
    }, "image/png");
  });
}

export interface RenderPdfTextOpToImageOptions {
  /**
   * Optional opaque backing used by managed `/FreeText` appearances. Omit it
   * for the historical transparent flattened-export fallback.
   */
  backgroundColor?: string;
  /** Source-bound editable boxes share the native input's browser line layout. */
  browserLayout?: boolean;
  /** Source paragraph leading, shared by the input and portable row plan. */
  browserLineHeightRatio?: number;
  /** Live font alias is ephemeral; never stored in the document record. */
  displayFontFamily?: string;
  /** A source-bound box's character/word spacing and horizontal scale: the
   * row plan is measured and each row drawn with them. */
  browserSpacing?: PdfBrowserTextSpacing;
  /**
   * A text box set in runs (text-box-runs.ts): `op.runs` are its drawn runs
   * (sizes in points). Each is measured in the font program the writer
   * embeds, at its glyphs' own advances, and the rows are as tall as their
   * largest run (Word's single spacing). Implies `browserLayout`.
   */
  runRows?: boolean;
}

/**
 * Rasterize one PDF text draw op with the same mixed-run and line-layout
 * contract used by flattened export. The transparent PNG is also suitable for
 * a standard `/FreeText` appearance stream, so native writeback and the
 * historical flattened fallback cannot drift visually.
 */
export async function renderPdfTextOpToImage(
  op: TextOp,
  options: RenderPdfTextOpToImageOptions = {},
): Promise<(FlattenImageBytes & { portableLayout?: PdfPortableTextLayout }) | null> {
  if (typeof document === "undefined" || typeof window === "undefined") {
    return null;
  }
  if (!op.text.trim() || op.fontSize <= 0) return null;
  if (op.maxWidth <= 0 || op.maxHeight <= 0) return null;
  const boxWidth = Math.max(1, Math.ceil(op.maxWidth));
  const boxHeight = Math.max(1, Math.ceil(op.maxHeight));
  const desiredScale = Math.min(3, Math.max(2, window.devicePixelRatio || 1));
  const canvas = document.createElement("canvas");
  let backing = pdfTextRasterBackingSize(boxWidth, boxHeight, desiredScale);
  canvas.width = backing.width;
  canvas.height = backing.height;
  const initialContext = canvas.getContext("2d");
  if (!initialContext) return null;
  let ctx: CanvasRenderingContext2D = initialContext;
  ctx.scale(backing.scale, backing.scale);
  const richRuns =
    op.runs?.length && op.runs.map((run) => run.text).join("") === op.text
      ? op.runs
      : null;
  type RichRun = NonNullable<TextOp["runs"]>[number];
  const fallbackRun: RichRun = {
    text: op.text,
    fontSize: op.fontSize,
    color: op.color,
    font: op.font,
    bold: op.bold,
    italic: op.italic,
    fontFamily: op.fontFamily,
    sourceFontName: op.sourceFontName,
    displayFontFamily: options.displayFontFamily,
  };
  const layoutRuns: RichRun[] = richRuns ? [...richRuns] : [fallbackRun];
  const runRows = Boolean(options.runRows && richRuns);
  const browserLayout = Boolean(options.browserLayout || runRows);
  // A run's italic over a family that has no italic face of its own is drawn
  // upright and sheared as the writer slants it, not in another family
  // (lib/pdf/draft-italic.ts); the font asked for says "italic" only when
  // the family has one.
  const slantForRun = (run: RichRun): number => {
    if (runRows) return pdfTextBoxRunItalicPresentation(run, op.text).synthetic ? PDF_NATIVE_SYNTHETIC_ITALIC_SKEW : 0;
    const family = withEastAsianFallback(
      run.displayFontFamily || run.fontFamily ||
        (browserLayout ? FONT_STACK[run.font as keyof typeof FONT_STACK] ?? FONT_STACK.sans : textExportFontFamily(run.font)),
      op.text,
    );
    return pdfDraftItalic(family, Boolean(run.italic)).synthetic ? PDF_NATIVE_SYNTHETIC_ITALIC_SKEW : 0;
  };
  // WordArt: the words drawn with the box's Office paint (made once the
  // canvas is its final size: a pattern belongs to its context).
  let wordArt: ((text: string, x: number, baseline: number) => void) | null = null;
  const drawWords = (text: string, x: number, baseline: number) => {
    if (op.typography?.textPaint) {
      wordArt ??= pdfWordArtCanvasPainter(ctx, op.typography.textPaint, { width: op.maxWidth, height: op.maxHeight }, backing.scale);
      wordArt(text, x, baseline);
      return;
    }
    ctx.fillText(text, x, baseline);
  };
  const fillSlanted = (text: string, x: number, baseline: number, run: RichRun) => {
    const slant = slantForRun(run);
    if (!slant) { drawWords(text, x, baseline); return; }
    ctx.save();
    // x' = x + slant * (baseline - y): the glyphs lean right from their baseline.
    ctx.transform(1, 0, -slant, 1, slant * baseline, 0);
    drawWords(text, x, baseline);
    ctx.restore();
  };
  const fontForRun = (run: RichRun) => {
    if (runRows) return pdfTextBoxRunCssFont(run, run.fontSize, op.text);
    const style = run.italic && !slantForRun(run) ? "italic " : "";
    const weight = run.bold ? "bold " : "";
    // CJK falls back to the face the saved PDF sets it in, so a plan
    // measured here is the row the server draws (withEastAsianFallback).
    // A source-bound box's rows are the rows the server draws: its Latin in
    // the base-14 face of the box's class (Helvetica-Bold for a bold sans
    // line), anything else in the embedded Noto Sans SC. A box with no CSS
    // family of its own was measured in the CJK-first export stack instead,
    // its Latin in PingFang/YaHei, wider: a one-row box wrapped past its
    
    const family = withEastAsianFallback(
      run.displayFontFamily || run.fontFamily ||
        (browserLayout ? FONT_STACK[run.font as keyof typeof FONT_STACK] ?? FONT_STACK.sans : textExportFontFamily(run.font)),
      op.text,
    );
    return `${style}${weight}${run.fontSize}px ${family}`;
  };
  const measureRunCharacter = (character: string, run: RichRun) => {
    ctx.font = fontForRun(run);
    // The served face's own advance (lib/text-engine, pdf flag) where the
    // engine knows it: the same rows in every browser.
    return pageTextWidth("pdf", character, fontForRun(run)) ?? ctx.measureText(character).width;
  };
  // The rows measured here are the rows the server draws: the East Asian
  // face must be the PDF's own before anything is measured, not the system
  // face the stack falls back to while it loads.
  const stageAt = performance.now();
  const stages: Record<string, number> = {};
  const stage = (name: string) => { stages[name] = Math.round(performance.now() - stageAt); };
  if (needsEastAsianFace(op.text)) await ensurePdfEastAsianFace();
  stage("eastAsian");
  if (browserLayout) {
    // Every load answered, a face that cannot load included: one face of a
    // stack this machine lacks is no reason not to draw the box in the
    // faces the browser draws it with (text-faces-ready.ts).
    const refused = await pdfTextFacesReady(document.fonts, layoutRuns.map(run => [fontForRun(run), run.text] as const));
    if (refused) stages.facesRefused = refused;
  }
  stage("fonts");
  // Word's 竖排: the box in columns (text-box-vertical.ts), as the server
  // draws it -- no row plan, one style for the whole box.
  if (op.typography?.writingMode === "vertical-rl") {
    const run = fallbackRun;
    const pitchEm = op.typography.lineSpacing ? op.typography.lineSpacing / op.fontSize : PDF_TEXTBOX_LINE_HEIGHT_RATIO;
    const measureAt = (size: number) => (value: string) => {
      const font = fontForRun({ ...run, fontSize: size });
      ctx.font = font;
      return pageTextWidth("pdf", value, font) ?? ctx.measureText(value).width;
    };
    let size = op.fontSize;
    let columns = layoutPdfTextBoxVerticalColumns(op.text, { measure: measureAt(size), size, columnLength: op.maxHeight });
    for (let attempt = 0; attempt < 8; attempt += 1) {
      const required = Math.max(1, columns.length) * size * pitchEm;
      if (required <= op.maxWidth + 0.01) break;
      size = Math.max(0.01, size * Math.min(0.95, (op.maxWidth / required) * 0.98));
      columns = layoutPdfTextBoxVerticalColumns(op.text, { measure: measureAt(size), size, columnLength: op.maxHeight });
    }
    const measure = measureAt(size);
    const placed = placePdfTextBoxVerticalColumns(columns, {
      size, pitch: size * pitchEm, frameWidth: op.maxWidth, frameHeight: op.maxHeight,
      align: op.justify ? "justify" : op.align, verticalAlign: op.typography.verticalAlign, measure,
    });
    if (options.backgroundColor) {
      ctx.fillStyle = options.backgroundColor;
      ctx.fillRect(0, 0, boxWidth, boxHeight);
    }
    const highlight = op.typography.highlight;
    if (highlight) {
      ctx.fillStyle = highlight;
      for (const band of placed.bands) ctx.fillRect(band.centre - size * 0.6, band.top, size * 1.2, band.length);
    }
    const color = `rgb(${Math.round(run.color.r * 255)}, ${Math.round(run.color.g * 255)}, ${Math.round(run.color.b * 255)})`;
    ctx.fillStyle = color;
    ctx.font = fontForRun({ ...run, fontSize: size });
    ctx.textBaseline = "alphabetic";
    for (const glyph of placed.glyphs) {
      if (!glyph.rotated) { fillSlanted(glyph.text, glyph.x, glyph.y, run); continue; }
      ctx.save();
      ctx.translate(glyph.x, glyph.y);
      // A quarter clockwise (y runs down): the word reads down the column.
      ctx.rotate(Math.PI / 2);
      fillSlanted(glyph.text, 0, 0, run);
      ctx.restore();
    }
    const thickness = Math.max(0.5, size * 0.06);
    for (const band of placed.bands) {
      // Under a column is its right side.
      if (op.underline) ctx.fillRect(band.centre + size * 0.55 - thickness / 2, band.top, thickness, band.length);
      if (op.strikethrough) ctx.fillRect(band.centre - thickness / 2, band.top, thickness, band.length);
    }
    stage("drawn");
    const bytes = await canvasToPngBytes(canvas);
    if (!bytes) return null;
    return { bytes, mimeType: "image/png", displayWidth: boxWidth, displayHeight: boxHeight };
  }
  // Word's and PowerPoint's Columns: the rows are set at a column's width
  // and fill each column to the frame's foot before the next.
  const columnCount = op.typography?.columnCount && op.typography.columnCount > 1 ? op.typography.columnCount : 1;
  const columnGap = op.typography?.columnSpacing ?? 0;
  const rowWidth = pdfTextBoxColumnWidth(op.maxWidth, columnCount, columnGap);
  const browserLines = browserLayout && !op.lineLayout
    ? layoutBrowserPdfTextRuns(layoutRuns, rowWidth, fontForRun, options.browserLineHeightRatio,
      options.browserSpacing, { exactAdvances: runRows })
    : null;
  const laidOutLines = browserLines
    ? runRows ? pdfRunRowLines(op.text, layoutRuns, browserLines, options.browserLineHeightRatio) : browserLines
    : browserLayout && !op.lineLayout ? null : layoutPositionedPdfTextRuns({
    runs: layoutRuns,
    text: op.text,
    maxWidth: rowWidth,
    maxHeight: op.maxHeight,
    lineLayout: op.lineLayout,
    measure: measureRunCharacter,
  });
  if (!laidOutLines) return null;
  const columnedLines = columnCount > 1 && laidOutLines.length > 1 ? (() => {
    const pitch = laidOutLines[1]!.top - laidOutLines[0]!.top;
    const rows = pdfTextBoxColumnRows(op.maxHeight, pitch > 0 ? pitch : op.fontSize * PDF_TEXTBOX_LINE_HEIGHT_RATIO);
    return laidOutLines.map((positioned, index) => {
      const column = Math.floor(index / rows);
      if (column === 0) return positioned;
      const lift = laidOutLines[column * rows]!.top - laidOutLines[0]!.top;
      return {
        ...positioned,
        top: positioned.top - lift,
        left: positioned.left + column * (rowWidth + columnGap),
        ...("baseline" in positioned && typeof positioned.baseline === "number"
          ? { baseline: positioned.baseline - lift } : {}),
      };
    });
  })() : laidOutLines;
  // Word's Align Text: the rows move down into the frame's spare height
  // (and the row plan the writer draws carries them where they stand).
  const verticalOffset = pdfTextBoxVerticalOffset(
    op.typography?.verticalAlign,
    op.maxHeight,
    Math.max(0, ...columnedLines.map(({ top, line }) => top + line.height)),
  );
  const positionedLines = verticalOffset > 0
    ? columnedLines.map((positioned) => ({
      ...positioned,
      top: positioned.top + verticalOffset,
      ...("baseline" in positioned && typeof positioned.baseline === "number"
        ? { baseline: positioned.baseline + verticalOffset } : {}),
    }))
    : columnedLines;
  stage("lines");
  const portableLayout = browserLayout ? await createPdfPortableTextLayout({
    text:op.text,width:op.maxWidth,height:op.maxHeight,fontSize:op.fontSize,
    align:op.align,positionedLines,
    ...(runRows ? { minimumBaseline: Math.min(...layoutRuns.map((run) => run.fontSize)) } : {}),
    ...(columnCount > 1 ? { columnStep: rowWidth + columnGap } : {}),
  }) : null;
  if (browserLayout && !portableLayout && process.env.NODE_ENV === "development") {
    // Which row left the box, in which face (development only).
    try {
      console.warn("[pdf-text-box] portable rows refused", {
        width: op.maxWidth, height: op.maxHeight, size: op.fontSize, font: fontForRun(layoutRuns[0]!),
        rows: positionedLines.map(({ top, left, line }) => [Math.round(top * 10) / 10, Math.round(left * 10) / 10,
          Math.round(line.width * 10) / 10, line.segments.map((segment) => segment.text).join("").slice(0, 40)]),
      });
    } catch { /* never throws */ }
  }
  const displayWidth = Math.max(
    boxWidth,
    Math.ceil(
      Math.max(
        0,
        ...positionedLines.map(
          ({ left, line, horizontalScale }) =>
            left + line.width * horizontalScale,
        ),
      ),
    ),
  );
  const displayHeight = Math.max(
    boxHeight,
    Math.ceil(
      Math.max(0, ...positionedLines.map(({ top, line }) => top + line.height)),
    ),
  );
  backing = pdfTextRasterBackingSize(displayWidth, displayHeight, desiredScale);
  if (canvas.width !== backing.width || canvas.height !== backing.height) {
    canvas.width = backing.width;
    canvas.height = backing.height;
    const resizedContext = canvas.getContext("2d");
    if (!resizedContext) return null;
    ctx = resizedContext;
    ctx.scale(backing.scale, backing.scale);
    wordArt = null;
  }

  if (options.backgroundColor) {
    ctx.fillStyle = options.backgroundColor;
    ctx.fillRect(0, 0, displayWidth, displayHeight);
  }
  ctx.textBaseline = "alphabetic";
  // Canvas letter/word spacing (Chromium, Safari 17+): the row plan's own.
  const spacingContext = ctx as CanvasRenderingContext2D & { letterSpacing?: string; wordSpacing?: string };
  const applySpacing = (fontSize: number) => {
    if (!browserLayout || !options.browserSpacing) return;
    spacingContext.letterSpacing = `${(options.browserSpacing.characterEm ?? 0) * fontSize}px`;
    spacingContext.wordSpacing = `${(options.browserSpacing.wordEm ?? 0) * fontSize}px`;
  };
  const paragraphEnds = op.justify
    ? pdfTextLineParagraphEnds(
        op.text,
        positionedLines.map(({ line }) =>
          line.segments.map((segment) => segment.text).join(""),
        ),
      )
    : null;
  const decorate = (
    left: number,
    width: number,
    baselineY: number,
    fontSize: number,
    run?: RichRun,
  ) => {
    const thickness = Math.max(0.5, fontSize * 0.06);
    // A box set in runs underlines and strikes each run as it says.
    if (run && runRows ? run.underline : op.underline) {
      ctx.fillRect(left, baselineY + fontSize * 0.12 - thickness / 2, width, thickness);
    }
    if (run && runRows ? run.strikethrough : op.strikethrough) {
      ctx.fillRect(left, baselineY - fontSize * 0.3 - thickness / 2, width, thickness);
    }
  };
  for (const [lineIndex, positioned] of positionedLines.entries()) {
    const { line } = positioned;
    if (line.segments.length === 0) continue;
    const maxFontSize = Math.max(
      ...line.segments.map((segment) => segment.run.fontSize),
    );
    const renderedLineWidth = "contentAdvance" in positioned
      ? Number(positioned.contentAdvance) : line.width * positioned.horizontalScale;
    let x = positioned.left;
    if (op.align === "center") {
      x += Math.max(0, (positioned.availableWidth - renderedLineWidth) / 2);
    } else if (op.align === "right") {
      x += Math.max(0, positioned.availableWidth - renderedLineWidth);
    }
    const baselineY = "baseline" in positioned && typeof positioned.baseline === "number"
      ? positioned.baseline : positioned.top + maxFontSize;
    ctx.save();
    ctx.translate(x, 0);
    ctx.scale(positioned.horizontalScale, 1);
    // Word's Justify: a single-style line that does not end its paragraph is
    // stretched to the box width (text-justify.ts).
    const onlySegment = line.segments.length === 1 ? line.segments[0] : null;
    if (paragraphEnds && !paragraphEnds[lineIndex] && onlySegment) {
      const run = onlySegment.run;
      ctx.font = fontForRun(run);
      applySpacing(run.fontSize);
      ctx.fillStyle = `rgb(${Math.round(run.color.r * 255)}, ${Math.round(
        run.color.g * 255,
      )}, ${Math.round(run.color.b * 255)})`;
      const justified = justifyPdfTextLine(
        onlySegment.text,
        positioned.availableWidth / positioned.horizontalScale,
        (value) => pageTextWidth("pdf", value, fontForRun(run)) ?? ctx.measureText(value).width,
      );
      if (justified) {
        for (const piece of justified.pieces) {
          fillSlanted(piece.text, piece.x, baselineY, run);
        }
        decorate(0, justified.width, baselineY, run.fontSize);
        ctx.restore();
        continue;
      }
    }
    let segmentX = 0;
    for (const segment of line.segments) {
      const run = segment.run;
      ctx.font = fontForRun(run);
      applySpacing(run.fontSize);
      ctx.fillStyle = `rgb(${Math.round(run.color.r * 255)}, ${Math.round(
        run.color.g * 255,
      )}, ${Math.round(run.color.b * 255)})`;
      if (runRows) {
        // A run's glyphs at their own size, raised or lowered by its script.
        const glyphBaseline = baselineY - pdfTextBoxRunRise(run, run.fontSize);
        fillSlanted(segment.text, segmentX, glyphBaseline, run);
        decorate(segmentX, segment.width, glyphBaseline, pdfTextBoxRunGlyphSize(run, run.fontSize), run);
      } else {
        fillSlanted(segment.text, segmentX, baselineY, run);
        decorate(segmentX, segment.width, baselineY, run.fontSize);
      }
      segmentX += segment.width;
    }
    ctx.restore();
  }
  stage("drawn");
  const bytes = await canvasToPngBytes(canvas);
  stage("png");
  if (process.env.NODE_ENV === "development") markPdfTextLatency("raster.stages", stages);
  if (!bytes) return null;
  return {
    bytes,
    mimeType: "image/png",
    displayWidth,
    displayHeight,
    ...(portableLayout ? { portableLayout } : {}),
  };
}

export function rasterBytesToDataUrl(
  bytes: Uint8Array | ArrayBuffer,
  mimeType = "image/png",
): string {
  const normalized =
    bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let binary = "";
  for (let offset = 0; offset < normalized.length; offset += 0x8000) {
    binary += String.fromCharCode(
      ...normalized.subarray(offset, offset + 0x8000),
    );
  }
  return `data:${mimeType};base64,${globalThis.btoa(binary)}`;
}

/**
 * The browser's rows of a box set in runs, placed as the writer places them:
 * each row as tall as its largest run (text-box-runs.ts), its baseline one
 * row size below its top. The rows themselves (which characters) are the
 * browser's line breaks.
 */
function pdfRunRowLines<Run extends { text: string; fontSize: number }>(
  text: string,
  runs: readonly Run[],
  lines: readonly PdfBrowserTextLayoutLine<Run>[],
  lineHeightRatio?: number,
): (PdfBrowserTextLayoutLine<Run> & { baseline: number })[] {
  const starts: number[] = [];
  let runOffset = 0;
  for (const run of runs) {
    starts.push(runOffset);
    runOffset += run.text.length;
  }
  const sizeAt = (offset: number) => {
    let index = 0;
    while (index + 1 < starts.length && starts[index + 1] <= offset) index += 1;
    return runs[index]?.fontSize ?? 0;
  };
  const rows: { start: number; end: number; size: number }[] = [];
  let offset = 0;
  for (const positioned of lines) {
    const rowText = positioned.line.segments.map((segment) => segment.text).join("");
    // The rows' offsets as the row plan counts them (createPdfPortableTextLayout).
    while (offset < text.length && /[\r\n]/u.test(text[offset])) offset += 1;
    rows.push({
      start: offset,
      end: offset + rowText.length,
      size: Math.max(...positioned.line.segments.map((segment) => segment.run.fontSize)),
    });
    offset += rowText.length;
  }
  const geometry = pdfTextBoxRunRowGeometry(text, rows, sizeAt, lineHeightRatio);
  return lines.map((positioned, index) => ({
    ...positioned,
    top: geometry[index].top,
    baseline: geometry[index].baseline,
    line: { ...positioned.line, height: rows[index].size * (lineHeightRatio ?? 1.2) },
  }));
}
