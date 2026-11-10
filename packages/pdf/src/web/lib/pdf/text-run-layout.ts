/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizePdfTextLineLayout } from "@mrnote/core/web/lib/pdf/text-line-layout.ts";

/** Shared line-height used by PDF text editing, read overlays, and export. */
export const PDF_TEXTBOX_LINE_HEIGHT_RATIO = 1.2;
/** Browser-safe ceiling for one text-export canvas edge. */
export const PDF_TEXT_RASTER_MAX_DIMENSION = 8_192;
/** Browser-safe total backing-store budget (16 megapixels). */
export const PDF_TEXT_RASTER_MAX_PIXELS = 16_777_216;

export interface PdfTextRasterBackingSize {
  width: number;
  height: number;
  scale: number;
}

/**
 * Bound text-export backing resolution without changing its LOGICAL size.
 * The returned scale can fall below 1 for adversarially long/narrow edits;
 * callers still return the original logical displayWidth/displayHeight so the
 * PDF placement stays correct while the browser avoids an oversized canvas.
 */
export function pdfTextRasterBackingSize(
  logicalWidth: number,
  logicalHeight: number,
  desiredScale: number,
): PdfTextRasterBackingSize {
  const width = Math.max(1, Number.isFinite(logicalWidth) ? logicalWidth : 1);
  const height = Math.max(
    1,
    Number.isFinite(logicalHeight) ? logicalHeight : 1,
  );
  const requested = Math.max(
    Number.EPSILON,
    Number.isFinite(desiredScale) ? desiredScale : 1,
  );
  // Reserve headroom for ceil(width*scale) / ceil(height*scale).
  const safeDimension = PDF_TEXT_RASTER_MAX_DIMENSION - 1;
  const safePixels =
    PDF_TEXT_RASTER_MAX_PIXELS - PDF_TEXT_RASTER_MAX_DIMENSION * 2 - 1;
  const scale = Math.max(
    Number.EPSILON,
    Math.min(
      requested,
      safeDimension / width,
      safeDimension / height,
      Math.sqrt(safePixels / (width * height)),
    ),
  );
  return {
    width: Math.max(1, Math.ceil(width * scale)),
    height: Math.max(1, Math.ceil(height * scale)),
    scale,
  };
}

export interface PdfTextLayoutRun {
  text: string;
  fontSize: number;
}

export interface PdfTextLayoutSegment<Run extends PdfTextLayoutRun> {
  text: string;
  width: number;
  run: Run;
}

export interface PdfTextLayoutLine<Run extends PdfTextLayoutRun> {
  segments: Array<PdfTextLayoutSegment<Run>>;
  width: number;
  height: number;
}

export interface PdfPositionedTextLayoutLine<Run extends PdfTextLayoutRun> {
  line: PdfTextLayoutLine<Run>;
  left: number;
  top: number;
  availableWidth: number;
  /** Horizontal-only compression for source-line metric drift. */
  horizontalScale: number;
}

type Glyph<Run extends PdfTextLayoutRun> = {
  text: string;
  width: number;
  run: Run;
};

/**
 * Word-wrap styled PDF runs without dropping content.
 *
 * Normal words move as a unit; only a word wider than the box is broken by
 * glyph. Explicit newlines always start a line, and spaces remain present to
 * match `white-space: pre-wrap` in the editor/view overlays.
 */
export function layoutPdfTextRuns<Run extends PdfTextLayoutRun>(
  runs: readonly Run[],
  maxWidth: number,
  measure: (text: string, run: Run) => number,
): { lines: Array<PdfTextLayoutLine<Run>>; contentHeight: number } {
  const safeWidth = Math.max(0.01, maxWidth);
  const fallbackFontSize = Math.max(
    1,
    ...runs.map((run) => (run.fontSize > 0 ? run.fontSize : 0)),
  );
  const newLine = (): PdfTextLayoutLine<Run> => ({
    segments: [],
    width: 0,
    height: fallbackFontSize * PDF_TEXTBOX_LINE_HEIGHT_RATIO,
  });
  const lines: Array<PdfTextLayoutLine<Run>> = [newLine()];
  let token: Array<Glyph<Run>> = [];
  let tokenIsWhitespace = false;

  const currentLine = () => lines[lines.length - 1];
  const startLine = () => lines.push(newLine());
  const appendGlyph = (glyph: Glyph<Run>) => {
    const line = currentLine();
    const previous = line.segments[line.segments.length - 1];
    if (previous?.run === glyph.run) {
      previous.text += glyph.text;
      previous.width += glyph.width;
    } else {
      line.segments.push({ ...glyph });
    }
    line.width += glyph.width;
    line.height = Math.max(
      line.height,
      glyph.run.fontSize * PDF_TEXTBOX_LINE_HEIGHT_RATIO,
    );
  };
  const appendBroken = (glyphs: readonly Glyph<Run>[]) => {
    for (const glyph of glyphs) {
      if (
        currentLine().segments.length > 0 &&
        currentLine().width + glyph.width > safeWidth
      ) {
        startLine();
      }
      appendGlyph(glyph);
    }
  };
  const flushToken = () => {
    if (token.length === 0) return;
    const tokenWidth = token.reduce((sum, glyph) => sum + glyph.width, 0);
    const line = currentLine();
    if (line.segments.length > 0 && line.width + tokenWidth > safeWidth) {
      startLine();
    }
    if (tokenWidth > safeWidth) appendBroken(token);
    else for (const glyph of token) appendGlyph(glyph);
    token = [];
  };

  for (const run of runs) {
    for (const character of Array.from(run.text)) {
      if (character === "\n") {
        flushToken();
        startLine();
        continue;
      }
      const whitespace = /\s/u.test(character);
      if (token.length > 0 && whitespace !== tokenIsWhitespace) flushToken();
      if (token.length === 0) tokenIsWhitespace = whitespace;
      token.push({
        text: character,
        width: Math.max(0, measure(character, run)),
        run,
      });
    }
  }
  flushToken();

  return {
    lines,
    contentHeight: lines.reduce((sum, line) => sum + line.height, 0),
  };
}

/**
 * Position run lines inside source PDF line rectangles. With lineLayout, one
 * source line is ALWAYS one output line: browser font metrics can be slightly
 * wider than the embedded PDF face, but that must not create new lines and then
 * shrink the whole paragraph. Small width drift is handled by horizontal-only
 * compression. Invalid/legacy metadata keeps ordinary paragraph wrapping.
 */
export function layoutPositionedPdfTextRuns<Run extends PdfTextLayoutRun>({
  runs,
  text,
  maxWidth,
  maxHeight,
  lineLayout,
  measure,
}: {
  runs: readonly Run[];
  text: string;
  maxWidth: number;
  maxHeight: number;
  lineLayout: unknown;
  measure: (text: string, run: Run) => number;
}): PdfPositionedTextLayoutLine<Run>[] {
  const positioned: PdfPositionedTextLayoutLine<Run>[] = [];
  const sourceLines = normalizePdfTextLineLayout(lineLayout, text);
  if (!sourceLines) {
    const layout = layoutPdfTextRuns(runs, maxWidth, measure);
    let top = 0;
    for (const line of layout.lines) {
      positioned.push({
        line,
        left: 0,
        top,
        availableWidth: maxWidth,
        horizontalScale: 1,
      });
      top += line.height;
    }
    return positioned;
  }

  const runsForRange = (start: number, end: number): Run[] => {
    const sliced: Run[] = [];
    let offset = 0;
    for (const run of runs) {
      const runStart = offset;
      const runEnd = runStart + run.text.length;
      offset = runEnd;
      if (runEnd <= start || runStart >= end) continue;
      const localStart = Math.max(start, runStart) - runStart;
      const localEnd = Math.min(end, runEnd) - runStart;
      const runText = run.text.slice(localStart, localEnd);
      if (runText) sliced.push({ ...run, text: runText });
    }
    if (sliced.length > 0) return sliced;
    const fallback = runs.find((run) => run.text.length > 0) ?? runs[0];
    return fallback ? [{ ...fallback, text: "" }] : [];
  };

  for (const sourceLine of sourceLines) {
    const availableWidth = Math.max(0.01, sourceLine.width * maxWidth);
    const layout = layoutPdfTextRuns(
      runsForRange(sourceLine.start, sourceLine.end),
      Number.MAX_SAFE_INTEGER,
      measure,
    );
    const line = layout.lines[0];
    if (!line) continue;
    positioned.push({
      line,
      left: sourceLine.x * maxWidth,
      top: sourceLine.y * maxHeight,
      availableWidth,
      horizontalScale:
        line.width > availableWidth ? availableWidth / line.width : 1,
    });
  }
  return positioned;
}
