/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfRangeGlyphRect } from "./range-glyph-rect";
import type {
  PdfPositionedTextLayoutLine,
  PdfTextLayoutRun,
} from "./text-run-layout";
import { PDF_TEXTBOX_LINE_HEIGHT_RATIO } from "./text-run-layout";
import { applyPdfTextBoxExactAdvanceStyle } from "./text-box-runs";

/** The native input and saved appearance use the browser's SAME line breaker.
 * In particular, CJK punctuation, shaping, spaces and long tokens are not
 * approximated by adding individual canvas glyph advances. No soft newlines
 * are inserted into the user's text or persisted as source-line evidence. */
export const PDF_BROWSER_TEXT_WRAP = {
  whiteSpace: "pre-wrap",
  overflowWrap: "anywhere",
  wordBreak: "break-word",
} as const;
export type PdfBrowserTextLayoutLine<Run extends PdfTextLayoutRun> =
  PdfPositionedTextLayoutLine<Run> & { contentAdvance: number };

/** A run's character spacing and word spacing (ems of its font) and
 * horizontal scale (percent), as PDF Tc, Tw and Tz: the browser lays the
 * text out with letter-spacing and word-spacing in a frame 1/k as wide, and
 * each line is drawn scaled by k (the native dialog's wrapped draft does the
 * same, so the draft, the row plan and the saved rows break alike). */
export interface PdfBrowserTextSpacing {
  readonly characterEm?: number;
  readonly wordEm?: number;
  readonly horizontalScale?: number;
}

/** A very narrow frame may not fit a legal CJK punctuation pair or grapheme.
 * Measure the actual font/line breaker, not an average character width. Each
 * retry starts from the user's requested width; no layout-feedback growth. */
export function minimumPdfBrowserTextWidth<Run extends PdfTextLayoutRun>(
  runs: readonly Run[],
  width: number,
  fontForRun: (run: Run) => string,
): number | null {
  if (everyClusterFitsWidth(runs, width, fontForRun)) return width;
  let measuredWidth = width;
  for (let attempt = 0; attempt < 4; attempt++) {
    const lines = layoutBrowserPdfTextRuns(runs, measuredWidth, fontForRun);
    if (!lines) return null;
    const required = Math.max(
      measuredWidth,
      ...lines.map((line) => line.contentAdvance),
    );
    if (required <= measuredWidth + 0.01) return measuredWidth;
    measuredWidth = Math.ceil(required * 64) / 64;
  }
  return null;
}

let clusterContext: CanvasRenderingContext2D | null | undefined;
const clusterSegmenter = typeof Intl !== "undefined" && "Segmenter" in Intl
  ? new Intl.Segmenter(undefined, { granularity: "grapheme" }) : null;












function everyClusterFitsWidth<Run extends PdfTextLayoutRun>(
  runs: readonly Run[],
  width: number,
  fontForRun: (run: Run) => string,
): boolean {
  if (!clusterSegmenter || typeof document === "undefined" || !Number.isFinite(width) || width <= 2 || runs.length === 0) return false;
  if (runs.reduce((count, run) => count + run.text.length, 0) > 20_000) return false;
  if (clusterContext === undefined) {
    try {
      clusterContext = document.createElement("canvas").getContext("2d");
    } catch {
      clusterContext = null;
    }
  }
  const context = clusterContext;
  if (!context) return false;
  for (const run of runs) {
    if (!Number.isFinite(run.fontSize) || run.fontSize <= 0 || run.text.includes("\t")) return false;
    context.font = fontForRun(run);
    const measured = new Set<string>();
    for (const { segment } of clusterSegmenter.segment(run.text)) {
      if (segment === "\n" || segment === "\r" || segment === "\r\n" || measured.has(segment)) continue;
      measured.add(segment);
      const advance = context.measureText(segment).width;
      // Half the width: the canvas may resolve a generic family to another
      // face than the page does (fallback-glyph-advances measureBrowserSpacePx),
      // never one twice as wide.
      if (!(Number.isFinite(advance) && advance * 2 < width)) return false;
    }
  }
  return true;
}

export function layoutBrowserPdfTextRuns<Run extends PdfTextLayoutRun>(
  runs: readonly Run[],
  width: number,
  fontForRun: (run: Run) => string,
  lineHeightRatio = PDF_TEXTBOX_LINE_HEIGHT_RATIO,
  spacing: PdfBrowserTextSpacing = {},
  options: {
    /** A box set in runs: each glyph at its font's advance (no kerning,
     * ligatures or CJK punctuation trimming), as the writer sets it. */
    exactAdvances?: boolean;
  } = {},
): PdfBrowserTextLayoutLine<Run>[] | null {
  const characterEm = spacing.characterEm ?? 0;
  const wordEm = spacing.wordEm ?? 0;
  const scale = (spacing.horizontalScale ?? 100) / 100;
  if (
    !Number.isFinite(characterEm) || Math.abs(characterEm) > 5 ||
    !Number.isFinite(wordEm) || Math.abs(wordEm) > 5 ||
    !Number.isFinite(scale) || scale < 0.1 || scale > 5
  )
    return null;
  if (
    !Number.isFinite(width) ||
    width <= 0 ||
    !Number.isFinite(lineHeightRatio) || lineHeightRatio < 0.5 || lineHeightRatio > 10 ||
    runs.length === 0 ||
    runs.reduce((count, run) => count + run.text.length, 0) > 20_000 ||
    runs.some((run) => !Number.isFinite(run.fontSize) || run.fontSize <= 0)
  )
    return null;
  const host = document.createElement("div");
  const fontSize = Math.max(...runs.map((run) => run.fontSize));
  Object.assign(host.style, PDF_BROWSER_TEXT_WRAP, {
    position: "fixed",
    left: "0",
    top: "0",
    visibility: "hidden",
    pointerEvents: "none",
    contain: "layout style paint",
    margin: "0",
    padding: "0",
    border: "0",
    width: `${width / scale}px`,
    height: "auto",
    font: fontForRun(runs[0]),
    fontSize: `${fontSize}px`,
    lineHeight: String(lineHeightRatio),
    fontKerning: "auto",
    letterSpacing: characterEm ? `${characterEm}em` : "normal",
    wordSpacing: wordEm ? `${wordEm}em` : "normal",
    textAlign: "left",
    direction: "ltr",
    textIndent: "0",
  });
  if (options.exactAdvances) applyPdfTextBoxExactAdvanceStyle(host);
  const nodes = runs.map((run) => {
    const span = document.createElement("span");
    span.style.font = fontForRun(run);
    span.style.lineHeight = String(lineHeightRatio);
    // `font` resets kerning and ligatures to their initial values.
    if (options.exactAdvances) applyPdfTextBoxExactAdvanceStyle(span);
    // `font` resets neither; each run's ems are of its own size.
    if (characterEm) span.style.letterSpacing = `${characterEm}em`;
    if (wordEm) span.style.wordSpacing = `${wordEm}em`;
    const node = document.createTextNode(run.text);
    span.appendChild(node);
    host.appendChild(span);
    return { run, node };
  });
  document.body.appendChild(host);
  // A one-line copy supplies each font's first-line offset without inserting
  // markers into the real text (which would change shaping and line breaks).
  const reference = host.cloneNode(true) as HTMLDivElement;
  reference.style.width = "100000px";
  reference.style.whiteSpace = "pre";
  for (const span of reference.children) span.textContent = "M";
  document.body.appendChild(reference);
  try {
    const origin = host.getBoundingClientRect();
    if (!(origin.width > 0)) return null; // No layout engine: never fake evidence.
    const range = document.createRange();
    const lines: PdfBrowserTextLayoutLine<Run>[] = [];
    let previousTop = -Infinity;
    const referenceTop = reference.getBoundingClientRect().top;
    for (const [index, { run, node }] of nodes.entries()) {
      range.selectNodeContents(reference.children[index]);
      const firstLineOffset = range.getBoundingClientRect().top - referenceTop;
      let offset = 0;
      // Ranges observe actual shaped text. Iterating codepoints does not split
      // the DOM nodes or change kerning/ligatures/line-break opportunities.
      for (const character of run.text) {
        range.setStart(node, offset);
        offset += character.length;
        range.setEnd(node, offset);
        if (character === "\n" || character === "\r") continue;
        const rect = pdfRangeGlyphRect(range);
        if (
          ![rect.left, rect.top, rect.width, rect.height].every(
            Number.isFinite,
          ) ||
          rect.height <= 0
        )
          return null;
        const top = rect.top - origin.top - firstLineOffset;
        let positioned = lines.at(-1);
        if (!positioned || Math.abs(top - previousTop) > 0.5) {
          positioned = {
            left: 0,
            top,
            availableWidth: width,
            horizontalScale: scale,
            contentAdvance: 0,
            line: {
              segments: [],
              width: 0,
              height: run.fontSize * lineHeightRatio,
            },
          };
          lines.push(positioned);
          previousTop = top;
        }
        const line = positioned.line;
        const segment = line.segments.at(-1);
        const right = rect.right - origin.left;
        if (segment?.run === run) {
          segment.text += character;
          segment.width = right - (line.width - segment.width);
        } else {
          line.segments.push({
            run,
            text: character,
            width: right - line.width,
          });
        }
        line.width = right;
        // In page units: the line is drawn scaled by k.
        if (/\S/u.test(character)) positioned.contentAdvance = right * scale;
        line.height = Math.max(
          line.height,
          run.fontSize * lineHeightRatio,
        );
      }
    }
    return lines;
  } finally {
    host.remove();
    reference.remove();
  }
}
