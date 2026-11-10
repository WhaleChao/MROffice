/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfTextLineLayout } from "@mrnote/core/web/lib/pdf/text-line-layout.ts";
import {
  BOX_RECT_EPSILON,
  MIN_BOX_FRACTION,
  resizeBox,
  type BoxRect,
  type ResizeHandle,
} from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";
import { PDF_TEXTBOX_LINE_HEIGHT_RATIO } from "./text-run-layout";

/**
 * Content height (page-height fraction) the text needs inside a frame of the
 * given width (page-width fraction). `null` means "unknown": the caller cannot
 * measure (a formula, no text) and the resize falls back to plain geometry.
 */
export type TextBoxContentHeight = (widthFraction: number) => number | null;

function horizontalPart(handle: ResizeHandle): "e" | "w" | null {
  if (handle.includes("e")) return "e";
  if (handle.includes("w")) return "w";
  return null;
}

function verticalPart(handle: ResizeHandle): "n" | "s" | null {
  if (handle.includes("n")) return "n";
  if (handle.includes("s")) return "s";
  return null;
}

/**
 * Resize `box` by dragging `handle` with the pointer travel `dx`/`dy` (page
 * fractions, in the box's own frame), honouring the rules above. Fields other
 * than the rect are carried through untouched.
 */
export function resizeTextBoxToFitContent<T extends BoxRect>(
  box: T,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  contentHeight: TextBoxContentHeight,
): T {
  const horizontal = horizontalPart(handle);
  const vertical = verticalPart(handle);

  // Rule 1: the horizontal part always applies, on its own axis.
  const widened = horizontal ? resizeBox(box, horizontal, dx, 0) : box;
  const minHeight = contentHeight(widened.width);
  if (minHeight === null || !Number.isFinite(minHeight)) {
    // Nothing to fit against: plain geometry, exactly as before.
    return resizeBox(box, handle, dx, dy);
  }
  const required = Math.max(MIN_BOX_FRACTION, minHeight);

  // Rules 2 and 3: the vertical part applies down to the content height and
  // no further -- the dragged edge stops on the text instead of the frame
  // jumping back to the height it started with.
  let next: T = widened;
  if (vertical) {
    const candidate = resizeBox(widened, vertical, 0, dy);
    if (candidate.height + BOX_RECT_EPSILON >= required) {
      next = candidate;
    } else {
      const height = Math.min(required, Math.max(MIN_BOX_FRACTION, 1 - widened.y));
      // A top edge pulled down stops above the text: the bottom edge stays.
      const y = vertical === "n" ? widened.y + widened.height - height : widened.y;
      next = { ...widened, y, height };
    }
  }

  // Rules 1 and 4: never shorter than the content, never snapped back to it.
  if (next.height + BOX_RECT_EPSILON < required) {
    const height = Math.min(required, Math.max(MIN_BOX_FRACTION, 1 - next.y));
    next = { ...next, height };
  }
  return next;
}

/**
 * Keep a paragraph's lines where they are on the page when only the frame's
 * height changes. Line geometry is stored relative to the frame, so a taller
 * frame would otherwise spread the lines apart like a stretched picture. The
 * lines stay anchored to the frame's top (text is top-aligned), so dragging
 * the top edge moves the text with it and dragging the bottom edge moves
 * nothing.
 */
export function rebaseLineLayoutForHeight(
  lineLayout: readonly PdfTextLineLayout[] | undefined,
  fromHeight: number,
  toHeight: number,
): PdfTextLineLayout[] | undefined {
  if (!lineLayout) return undefined;
  if (
    !Number.isFinite(fromHeight) ||
    !Number.isFinite(toHeight) ||
    fromHeight <= 0 ||
    toHeight <= 0
  ) {
    return [...lineLayout];
  }
  const ratio = fromHeight / toHeight;
  if (Math.abs(ratio - 1) < 1e-9) return [...lineLayout];
  return lineLayout.map((line) => ({
    ...line,
    y: line.y * ratio,
    height: line.height * ratio,
  }));
}

/**
 * The content height (page-height fraction) a laid-out paragraph occupies:
 * the bottom of its lowest line, measured from the frame's top.
 */
export function lineLayoutContentHeight(
  lineLayout: readonly PdfTextLineLayout[],
  boxHeight: number,
): number {
  let bottom = 0;
  for (const line of lineLayout) {
    bottom = Math.max(bottom, line.y + line.height);
  }
  return bottom * boxHeight;
}

/** Whether a persisted edit's frame is one the user sized by hand. */
export interface TextBoxWrapLike {
  created?: boolean;
  fixedWidth?: boolean;
}

/**
 * A text box whose text wraps inside its frame and whose frame grows downward
 * to hold it: a box the user drew ("Add Text") or one whose width the user has
 * set by dragging a handle. An untouched extracted line keeps its single-line
 * caret-over-canvas behaviour and the typing auto-fit that widens it.
 */
export function textBoxWrapsToWidth(
  edit: TextBoxWrapLike | null | undefined,
): boolean {
  return Boolean(edit && (edit.created || edit.fixedWidth));
}

/** A measure of a string's advance, in the same unit as the width given. */
export type TextAdvanceMeasure = (text: string) => number;

const WRAP_UNIT = /[A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-‐]+|\s+|[^\sA-Za-zÀ-ÖØ-öø-ÿ0-9]/gu;

/**
 * Greedy line count for `text` wrapped at `width`: Latin words wrap as units,
 * each CJK character and symbol on its own, hard breaks always break, and a
 * unit wider than the frame is broken by character so no width yields a
 * runaway line. Used only where no layout engine can measure (tests, and as a
 * defensive fallback); the browser measures the real line breaker.
 */
export function estimateWrappedLineCount(
  text: string,
  width: number,
  measure: TextAdvanceMeasure,
): number {
  if (!text) return 1;
  const paragraphs = text.split(/\r\n|\r|\n/);
  if (!(width > 0)) return Math.max(1, paragraphs.length);
  const overflows = (value: string) => measure(value) > width + 1e-6;
  let lines = 0;
  for (const paragraph of paragraphs) {
    lines += 1;
    let line = "";
    for (const match of paragraph.matchAll(WRAP_UNIT)) {
      const unit = match[0];
      if (/^\s+$/.test(unit)) {
        // Trailing spaces hang past the edge, as `pre-wrap` lets them.
        line += unit;
        continue;
      }
      if (line.trim() && overflows(line + unit)) {
        lines += 1;
        line = "";
      }
      if (!overflows(unit)) {
        line += unit;
        continue;
      }
      // A unit wider than the frame breaks by character.
      let piece = line;
      for (const character of unit) {
        if (piece.trim() && overflows(piece + character)) {
          lines += 1;
          piece = "";
        }
        piece += character;
      }
      line = piece;
    }
  }
  return Math.max(1, lines);
}

/**
 * Height (same unit as `fontSize`) of `lineCount` lines of text at the shared
 * PDF text box leading. One line is never shorter than the extraction band of
 * an untouched line (1em above the baseline, 0.25em below), so an extracted
 * line's own frame already "equals its content height" and rule 2 holds for
 * it without a drag of a fraction of a pixel first shaving it.
 */
export function textBoxContentHeightForLines(
  lineCount: number,
  fontSize: number,
  lineHeight: number = fontSize * PDF_TEXTBOX_LINE_HEIGHT_RATIO,
): number {
  const lines = Math.max(1, Math.floor(lineCount));
  const leading = lineHeight > 0 ? lineHeight : fontSize * PDF_TEXTBOX_LINE_HEIGHT_RATIO;
  return Math.max(fontSize * 1.25, lines * leading);
}
