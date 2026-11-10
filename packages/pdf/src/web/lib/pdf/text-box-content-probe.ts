/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * DOM-side measurement of how tall wrapped text is, for the WPS-style text box
 * resize (text-box-fit-resize.ts): a hidden probe laid out by the browser's
 * own line breaker, with a deterministic greedy estimate where no layout
 * engine exists (jsdom). Shared by the PDF text box overlay and the native
 * text edit dialog so both frames settle on the same height for the same text.
 */

import { FONT_STACK } from "./font-resolve";
import {
  estimateWrappedLineCount,
  textBoxContentHeightForLines,
} from "./text-box-fit-resize";
import { PDF_TEXTBOX_LINE_HEIGHT_RATIO } from "./text-run-layout";
import { pageTextWidth, pageWrappedLineCount } from "@mrnote/core/web/lib/text-engine/page-adapter.ts";

/** Canvas advance of `text` at `fontPx`; a half-em per character where no 2d
 * context exists (jsdom), so the estimate stays deterministic there. */
export function measureTextAdvancePx(
  text: string,
  fontPx: number,
  fontFamily: string,
  fontWeight: string,
  fontStyle: string,
): number {
  // A test double for the 2d context may lack measureText altogether; the
  // ruler is a convenience, never a reason for a drag to throw.
  // The served faces' own advances (lib/text-engine) where the engine knows
  // them: the same width in every browser.
  const engine = pageTextWidth("pdf", text, `${fontStyle} ${fontWeight} ${fontPx}px ${fontFamily || FONT_STACK.sans}`);
  if (engine !== undefined) return engine;
  try {
    const ctx = document.createElement("canvas").getContext("2d");
    if (ctx && typeof ctx.measureText === "function") {
      ctx.font = `${fontStyle} ${fontWeight} ${fontPx}px ${fontFamily || FONT_STACK.sans}`;
      const width = ctx.measureText(text)?.width;
      if (Number.isFinite(width)) return width;
    }
  } catch {
    // Fall through to the half-em estimate.
  }
  return text.length * fontPx * 0.5;
}

export interface WrappedTextHeightProbe {
  text: string;
  /** Frame width in CSS px. */
  widthPx: number;
  fontSizePx: number;
  /** Line pitch in CSS px; defaults to the shared PDF text box leading. */
  lineHeightPx?: number;
  fontFamily?: string;
  fontWeight?: string;
  fontStyle?: string;
  letterSpacing?: string;
}

/**
 * The height in CSS px that `text` needs when wrapped inside `widthPx` with
 * `pre-wrap` + `break-word` (the editor's own wrapping). Never shorter than an
 * extracted line's 1.25em band, so a one-line frame already equals its
 * content height.
 */
export function measureWrappedTextHeightPx(probe: WrappedTextHeightProbe): number {
  const fontSizePx = Math.max(1, probe.fontSizePx);
  const lineHeightPx =
    probe.lineHeightPx && probe.lineHeightPx > 0
      ? probe.lineHeightPx
      : fontSizePx * PDF_TEXTBOX_LINE_HEIGHT_RATIO;
  const fontFamily = probe.fontFamily || FONT_STACK.sans;
  const fontWeight = probe.fontWeight || "normal";
  const fontStyle = probe.fontStyle || "normal";
  // The line count from the engine's own line breaker where it can measure
  // the text (no letter spacing): every browser wraps the box alike.
  if (!probe.letterSpacing || probe.letterSpacing === "normal") {
    const engineLines = pageWrappedLineCount("pdf", probe.text || "", `${fontStyle} ${fontWeight} ${fontSizePx}px ${fontFamily}`, Math.max(1, probe.widthPx));
    if (engineLines !== undefined) return Math.max(Math.max(1, engineLines) * lineHeightPx, fontSizePx * 1.25);
  }
  const element = document.createElement("div");
  Object.assign(element.style, {
    position: "fixed",
    left: "0",
    top: "0",
    visibility: "hidden",
    pointerEvents: "none",
    margin: "0",
    padding: "0",
    border: "0",
    height: "auto",
    boxSizing: "border-box",
    whiteSpace: "pre-wrap",
    overflowWrap: "break-word",
    width: `${Math.max(1, probe.widthPx)}px`,
    fontSize: `${fontSizePx}px`,
    lineHeight: `${lineHeightPx}px`,
    fontFamily,
    fontWeight,
    fontStyle,
    letterSpacing: probe.letterSpacing || "normal",
  });
  element.textContent = probe.text || "​";
  document.body.appendChild(element);
  try {
    const measured = element.getBoundingClientRect().height;
    if (measured > 0) return Math.max(measured, fontSizePx * 1.25);
  } finally {
    element.remove();
  }
  const lines = estimateWrappedLineCount(probe.text, probe.widthPx, (value) =>
    measureTextAdvancePx(value, fontSizePx, fontFamily, fontWeight, fontStyle),
  );
  return textBoxContentHeightForLines(lines, fontSizePx, lineHeightPx);
}
