/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A chart's texts the editor draws itself (`ChartRenderPicture.texts`), as
 * SVG for the paths that take the picture as one image -- a print snapshot,
 * Copy as Picture: the words go back into the picture where the editor draws
 * them, in their runs' looks (a WordArt in its fill and outline).
 */
import { ptToPx } from "../office-units";
import type { ChartTextFrame, ChartTextPart } from "../engine-host/charts-types";
import { normalizeOfficeTextPaint } from "../office-text-paint";
import { cssFontFamilyList } from "../css-font-family";

const escape = (value: string) =>
  value.replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[ch]!);

function frameSvg(frame: ChartTextFrame): string {
  let paint: ReturnType<typeof normalizeOfficeTextPaint> | undefined;
  try {
    paint = frame.paint ? normalizeOfficeTextPaint(frame.paint) : undefined;
  } catch {
    paint = undefined;
  }
  // The runs in lines (a `\n` starts the next one).
  const lines: { text: string; look: ChartTextFrame["runs"][number]["look"] }[][] = [[]];
  for (const run of frame.runs) {
    const pieces = run.text.split("\n");
    pieces.forEach((piece, index) => {
      if (index > 0) lines.push([]);
      if (piece) lines[lines.length - 1].push({ text: frame.caps ? piece.toUpperCase() : piece, look: run.look });
    });
  }
  const fallback = frame.runs[0]?.look;
  const heights = lines.map((line) => Math.max(...line.map((run) => run.look.fontSize), fallback?.fontSize ?? 12));
  const total = heights.reduce((sum, value) => sum + value, 0);
  let top = -total / 2;
  const fill = paint ? (paint.fill === "none" ? "none" : paint.fillGradient?.stops[0]?.color ?? paint.fill) : null;
  const stroke = paint && paint.outline !== "none"
    ? ` stroke="${escape(paint.outline)}" stroke-width="${ptToPx(paint.outlineWidthPt)}"`
    : "";
  const texts = lines.map((line, index) => {
    const middle = top + heights[index] / 2;
    top += heights[index];
    const spans = line
      .map(
        (run) =>
          `<tspan font-family="${escape(cssFontFamilyList(run.look.fontFamily))}" font-size="${run.look.fontSize}" font-weight="${escape(run.look.fontWeight)}" font-style="${escape(run.look.fontStyle)}" fill="${escape(fill ?? run.look.color)}">${escape(run.text)}</tspan>`,
      )
      .join("");
    return `<text x="0" y="${middle}" text-anchor="middle" dominant-baseline="central" xml:space="preserve"${stroke}>${spans}</text>`;
  });
  const turn = frame.rotation ? ` rotate(${frame.rotation})` : "";
  return `<g transform="translate(${frame.cx} ${frame.cy})${turn}">${texts.join("")}</g>`;
}

/** The texts as SVG in the picture's own pixels. */
export function chartTextFramesSvg(texts: Partial<Record<ChartTextPart, ChartTextFrame>> | undefined): string {
  if (!texts) return "";
  return (["title", "axisX", "axisY"] as const)
    .map((part) => (texts[part] ? frameSvg(texts[part]!) : ""))
    .join("");
}

/** A shown picture with its texts put back, serialized (its own copy). */
export function chartPictureSvgWithTexts(
  svg: SVGSVGElement,
  texts: Partial<Record<ChartTextPart, ChartTextFrame>> | undefined,
): string {
  const copy = svg.cloneNode(true) as SVGSVGElement;
  const markup = chartTextFramesSvg(texts);
  if (markup) {
    const parsed = new DOMParser().parseFromString(
      `<svg xmlns="http://www.w3.org/2000/svg">${markup}</svg>`,
      "image/svg+xml",
    );
    for (const node of Array.from(parsed.documentElement.childNodes)) copy.appendChild(copy.ownerDocument.importNode(node, true));
  }
  return new XMLSerializer().serializeToString(copy);
}
