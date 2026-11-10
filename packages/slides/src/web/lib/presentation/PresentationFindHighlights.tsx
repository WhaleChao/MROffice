"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useLayoutEffect, useState, type RefObject } from "react";

import {
  officePaintTextControlMarks,
  type OfficeTextControlMark,
} from "@mrnote/core/web/lib/office-inactive-selection.ts";
import {
  OFFICE_FIND_CURRENT_PAINT,
  OFFICE_FIND_MATCH_PAINT,
  clearOfficeTextRangeHighlight,
  officeTextRangeHighlightsSupported,
  setOfficeTextRangeHighlight,
} from "@mrnote/core/web/lib/office-text-range-highlight.ts";

import styles from "./PresentationFindHighlights.module.css";

/**
 * Find's Reading Highlight on the slide: while the Find pane is open, every
 * match in the slide's text boxes is painted yellow and the one the pane is
 * on orange -- the document editor's `.document-find-match` rules
 * (styles/note-editor.css), drawn as a layer over the text instead of as
 * ProseMirror decorations. Nothing is written into the slide.
 *
 * A text box paints its words in one of two surfaces: a rich preview, or the
 * plain textarea. A hit in a preview is painted ON its words, as a named
 * highlight (lib/office-text-range-highlight): it turns with a turned box,
 * wraps with its lines and moves with the box, where a box measured from the
 * words' client rects stood upright around a turned line. A hit only a
 * textarea shows (whose text no range reaches), or any hit in a browser
 * without named highlights, keeps the layer: boxes measured on the
 * textarea's hidden mirror or the preview. A box with neither drawn (a
 * shape, a table cell, a chart title) gets no paint; the pane's selection
 * still puts the editor on it.
 */

export interface PresentationFindHighlight {
  readonly key: string;
  readonly blockId: string;
  /** Offsets into the box's flat text (paragraphs joined by "\n"). */
  readonly start: number;
  readonly end: number;
  readonly current: boolean;
}

/** A DOM point for a flat text offset inside a rich preview. */
export type PresentationPreviewPointResolver = (
  root: HTMLElement,
  offset: number,
) => { node: Node; offset: number };

interface HighlightRect {
  readonly key: string;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly current: boolean;
}

const PREVIEW_SELECTOR = [
  '[data-testid="presentation-paragraph-preview"]',
  '[data-testid="presentation-run-preview"]',
  '[data-testid="presentation-columns-preview"]',
  '[data-testid="presentation-imported-text-preview"]',
  // A plain box's words, drawn over its bridge textarea.
  '[data-testid="presentation-text-preview"]',
  '[data-testid="presentation-anchored-text-preview"]',
].join(",");

/** The textarea properties that decide where its glyphs fall. */
const MIRROR_PROPERTIES = [
  "box-sizing",
  "width",
  "height",
  "padding-top",
  "padding-right",
  "padding-bottom",
  "padding-left",
  "border-top-width",
  "border-right-width",
  "border-bottom-width",
  "border-left-width",
  "border-style",
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "font-stretch",
  "font-variant",
  "font-feature-settings",
  "line-height",
  "letter-spacing",
  "word-spacing",
  "text-align",
  "text-indent",
  "text-transform",
  "tab-size",
  "direction",
  "writing-mode",
] as const;

function cssEscape(value: string): string {
  return typeof CSS !== "undefined" && typeof CSS.escape === "function"
    ? CSS.escape(value)
    : value.replace(/["\\]/g, "\\$&");
}

function rangeRects(range: Range): DOMRect[] {
  try {
    return typeof range.getClientRects === "function"
      ? Array.from(range.getClientRects())
      : [];
  } catch {
    return [];
  }
}

/** The client rects of `[start, end)` in a textarea's text, via a mirror. */
function textareaRects(
  textarea: HTMLTextAreaElement,
  start: number,
  end: number,
): DOMRect[] {
  const parent = textarea.parentElement;
  if (!parent) return [];
  const computed = window.getComputedStyle(textarea);
  const mirror = document.createElement("div");
  for (const property of MIRROR_PROPERTIES) {
    mirror.style.setProperty(property, computed.getPropertyValue(property));
  }
  mirror.style.position = "absolute";
  mirror.style.left = `${textarea.offsetLeft}px`;
  mirror.style.top = `${textarea.offsetTop - textarea.scrollTop}px`;
  mirror.style.visibility = "hidden";
  mirror.style.pointerEvents = "none";
  mirror.style.whiteSpace = "pre-wrap";
  mirror.style.overflowWrap = "break-word";
  mirror.style.overflow = "hidden";
  mirror.style.borderColor = "transparent";
  const text = document.createTextNode(textarea.value);
  mirror.appendChild(text);
  parent.appendChild(mirror);
  try {
    const length = textarea.value.length;
    const range = document.createRange();
    range.setStart(text, Math.max(0, Math.min(length, start)));
    range.setEnd(text, Math.max(0, Math.min(length, end)));
    return rangeRects(range);
  } catch {
    return [];
  } finally {
    mirror.remove();
  }
}

/** The flat text range `[start, end)` of a rich preview, as a static range. */
function previewStaticRange(
  preview: HTMLElement,
  start: number,
  end: number,
  resolvePoint: PresentationPreviewPointResolver,
): StaticRange | null {
  try {
    const from = resolvePoint(preview, start);
    const to = resolvePoint(preview, end);
    return new StaticRange({
      startContainer: from.node,
      startOffset: from.offset,
      endContainer: to.node,
      endOffset: to.offset,
    });
  } catch {
    return null;
  }
}

function previewRects(
  preview: HTMLElement,
  start: number,
  end: number,
  resolvePoint: PresentationPreviewPointResolver,
): DOMRect[] {
  try {
    const from = resolvePoint(preview, start);
    const to = resolvePoint(preview, end);
    const range = document.createRange();
    range.setStart(from.node, from.offset);
    range.setEnd(to.node, to.offset);
    return rangeRects(range);
  } catch {
    return [];
  }
}

/**
 * The client rects of the flat text range `[start, end)` in one slide text
 * box (its `[data-block-id]` wrapper): measured in its rich preview when it
 * has one, else on a mirror of its textarea. Empty when neither is drawn.
 */
export function presentationTextRangeClientRects(
  wrapper: HTMLElement,
  start: number,
  end: number,
  resolvePoint: PresentationPreviewPointResolver,
): DOMRect[] {
  const preview = wrapper.querySelector<HTMLElement>(PREVIEW_SELECTOR);
  if (preview) return previewRects(preview, start, end, resolvePoint);
  const textarea = wrapper.querySelector<HTMLTextAreaElement>("textarea");
  return textarea ? textareaRects(textarea, start, end) : [];
}

interface MeasuredHighlights {
  /** Hits painted on their words: every match, and the current one. */
  readonly match: StaticRange[];
  readonly current: StaticRange[];
  /** Hits in a box being typed in (a textarea): painted on its own mirror. */
  readonly boxes: Map<HTMLTextAreaElement, OfficeTextControlMark[]>;
  /** Hits drawn as boxes on the layer (a browser without named highlights). */
  readonly rects: HighlightRect[];
}

// Word's Reading Highlight colours (OFFICE_FIND_*_PAINT), for a box's mirror.
const FIND_MATCH_BACKGROUND = "#fff200";
const FIND_CURRENT_BACKGROUND = "#ff9632";

function measureHighlights(
  canvas: HTMLElement,
  highlights: readonly PresentationFindHighlight[],
  resolvePoint: PresentationPreviewPointResolver,
  painted: boolean,
): MeasuredHighlights {
  const match: StaticRange[] = [];
  const current: StaticRange[] = [];
  const canvasRect = canvas.getBoundingClientRect();
  const scale =
    canvas.offsetWidth > 0 && canvasRect.width > 0
      ? canvasRect.width / canvas.offsetWidth
      : 1;
  const rects: HighlightRect[] = [];
  const boxes = new Map<HTMLTextAreaElement, OfficeTextControlMark[]>();
  for (const highlight of highlights) {
    const wrapper = canvas.querySelector<HTMLElement>(
      `[data-block-id="${cssEscape(highlight.blockId)}"]`,
    );
    if (!wrapper) continue;
    const preview = painted ? wrapper.querySelector<HTMLElement>(PREVIEW_SELECTOR) : null;
    const range = preview
      ? previewStaticRange(preview, highlight.start, highlight.end, resolvePoint)
      : null;
    if (range) {
      (highlight.current ? current : match).push(range);
      continue;
    }
    // A box being typed in shows its words in the textarea itself: the hit
    // is painted on a mirror of that box, beside it, in its own transform,
    // so a turned or scaled box keeps the band on its words (an upright
    // rectangle measured apart from the box sat off a turned one).
    const textarea = preview ? null : wrapper.querySelector<HTMLTextAreaElement>("textarea");
    if (textarea) {
      const marks = boxes.get(textarea) ?? [];
      marks.push({
        start: highlight.start,
        end: highlight.end,
        background: highlight.current ? FIND_CURRENT_BACKGROUND : FIND_MATCH_BACKGROUND,
        attributes: highlight.current ? { "data-find-current": "true" } : { "data-find-match": "true" },
      });
      boxes.set(textarea, marks);
      continue;
    }
    const clientRects = presentationTextRangeClientRects(
      wrapper,
      highlight.start,
      highlight.end,
      resolvePoint,
    );
    clientRects.forEach((rect, index) => {
      if (rect.width <= 0 || rect.height <= 0) return;
      rects.push({
        key: `${highlight.key}:${index}`,
        left: (rect.left - canvasRect.left) / scale,
        top: (rect.top - canvasRect.top) / scale,
        width: rect.width / scale,
        height: rect.height / scale,
        current: highlight.current,
      });
    });
  }
  return { match, current, boxes, rects };
}

export function PresentationFindHighlights({
  canvasRef,
  highlights,
  resolvePreviewPoint,
  layoutKey,
}: {
  readonly canvasRef: RefObject<HTMLElement | null>;
  readonly highlights: readonly PresentationFindHighlight[];
  readonly resolvePreviewPoint: PresentationPreviewPointResolver;
  /** Changes whenever the slide's layout may have (slide, zoom, text). */
  readonly layoutKey: string;
}): React.JSX.Element | null {
  // The rects belong to the highlights they were measured for, so a stale
  // measurement never paints over a newer list.
  const [measured, setMeasured] = useState<{
    readonly source: readonly PresentationFindHighlight[];
    readonly rects: readonly HighlightRect[];
  } | null>(null);
  useLayoutEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || highlights.length === 0) return;
    const owner = Symbol("presentation-find");
    const painted = officeTextRangeHighlightsSupported();
    let frame = 0;
    // The mirrors painted on boxes being typed in, by box; each follows its
    // box's layout itself and is redrawn here when the hits change.
    let mirrors: Array<() => void> = [];
    const clearMirrors = () => {
      mirrors.forEach((remove) => remove());
      mirrors = [];
    };
    const measure = () => {
      frame = 0;
      const measuredNow = measureHighlights(canvas, highlights, resolvePreviewPoint, painted);
      setOfficeTextRangeHighlight(OFFICE_FIND_MATCH_PAINT, owner, measuredNow.match);
      setOfficeTextRangeHighlight(OFFICE_FIND_CURRENT_PAINT, owner, measuredNow.current);
      clearMirrors();
      for (const [textarea, marks] of measuredNow.boxes) {
        const remove = officePaintTextControlMarks(textarea, marks);
        if (remove) mirrors.push(remove);
      }
      setMeasured({ source: highlights, rects: measuredNow.rects });
    };
    const schedule = () => {
      if (!frame) frame = window.requestAnimationFrame(measure);
    };
    // After the paint that placed the text (fonts and autofit included).
    frame = window.requestAnimationFrame(measure);
    const observer =
      typeof ResizeObserver === "undefined" ? null : new ResizeObserver(schedule);
    observer?.observe(canvas);
    // A painted hit follows its words through any move or turn; when the
    // words themselves are drawn again (new text nodes), it is found again.
    // The layer's own boxes are no news.
    const textObserver =
      painted && typeof MutationObserver !== "undefined"
        ? new MutationObserver((records) => {
            const outsideLayer = records.some((record) => {
              const element =
                record.target instanceof Element ? record.target : record.target.parentElement;
              // The layer's boxes and a box's own mirror are no news.
              const mirror = [...record.addedNodes, ...record.removedNodes].every(
                (node) => node instanceof Element && node.hasAttribute("data-office-inactive-selection"),
              ) && record.addedNodes.length + record.removedNodes.length > 0;
              return !mirror && !element?.closest('[data-testid="presentation-find-highlights"]');
            });
            if (outsideLayer) schedule();
          })
        : null;
    textObserver?.observe(canvas, { childList: true, characterData: true, subtree: true });
    return () => {
      if (frame) window.cancelAnimationFrame(frame);
      observer?.disconnect();
      textObserver?.disconnect();
      clearMirrors();
      clearOfficeTextRangeHighlight(OFFICE_FIND_MATCH_PAINT, owner);
      clearOfficeTextRangeHighlight(OFFICE_FIND_CURRENT_PAINT, owner);
    };
  }, [canvasRef, highlights, layoutKey, resolvePreviewPoint]);
  const rects = measured?.source === highlights ? measured.rects : [];
  if (rects.length === 0) return null;
  return (
    <div
      className={`office-find-highlights ${styles.layer}`}
      aria-hidden="true"
      data-testid="presentation-find-highlights"
    >
      {rects.map((rect) => (
        <span
          key={rect.key}
          className={
            rect.current
              ? `${styles.mark} document-find-match document-find-match--current`
              : `${styles.mark} document-find-match`
          }
          data-find-current={rect.current ? "true" : undefined}
          style={{
            left: rect.left,
            top: rect.top,
            width: rect.width,
            height: rect.height,
          }}
        />
      ))}
    </div>
  );
}
