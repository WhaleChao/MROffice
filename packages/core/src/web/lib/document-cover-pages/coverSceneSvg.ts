/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A cover scene as one static SVG document, for surfaces that cannot run the
 * editor's node view (offline HTML export). Text is laid out by a plain
 * width estimate; the live editor lays it out with the browser instead.
 */

import { PX_PER_PT } from "../office-units";
import {
  formatCoverDate,
  type CoverFieldKey,
  type CoverScene,
  type CoverScenePaint,
  type CoverSceneRun,
} from "./coverPages";
import { charBefore, splitBehind } from "../regexp-behind";

const EMU_PER_PX = 9_525;

function escapeSvg(value: string): string {
  return value
    .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function paintAttrs(
  paint: CoverScenePaint | null | undefined,
  gradientId: string,
  role: "fill" | "stroke",
): string {
  if (!paint) return `${role}="none"`;
  if ("color" in paint) {
    return `${role}="${paint.color}"${paint.alpha !== undefined ? ` ${role}-opacity="${paint.alpha}"` : ""}`;
  }
  return `${role}="url(#${gradientId})"`;
}

function gradientDefinition(
  paint: CoverScenePaint | null | undefined,
  id: string,
): string {
  if (!paint || "color" in paint) return "";
  const radians = (paint.gradient.angle * Math.PI) / 180;
  const dx = Math.cos(radians) / 2;
  const dy = Math.sin(radians) / 2;
  const stops = paint.gradient.stops
    .map(
      (stop) =>
        `<stop offset="${stop.pos}" stop-color="${stop.color}" stop-opacity="${stop.alpha ?? 1}"/>`,
    )
    .join("");
  return `<linearGradient id="${id}" x1="${0.5 - dx}" y1="${0.5 - dy}" x2="${0.5 + dx}" y2="${0.5 + dy}">${stops}</linearGradient>`;
}

function runText(
  run: CoverSceneRun,
  values: Partial<Record<CoverFieldKey, string>>,
): string {
  if (!run.field) return run.text ?? "";
  const value = (values[run.field] ?? "").trim();
  if (!value) return run.placeholder ?? "";
  return run.dateFormat ? formatCoverDate(value, run.dateFormat) : value;
}

function estimatedWidth(text: string, sizePx: number): number {
  let width = 0;
  for (const char of text) {
    width += /[⺀-￯]/.test(char) ? sizePx : sizePx * 0.55;
  }
  return width;
}

function wrapWords(text: string, sizePx: number, maxWidth: number): string[] {
  if (maxWidth <= 0) return [text];
  const lines: string[] = [];
  let line = "";
  // `split(/(\s+|(?<=[⺀-￯]))/u)`: white space runs, and an empty cut after
    // CJK, tested in code (Safari before 16.4 cannot compile the lookbehind).
    for (const token of splitBehind(text, /(\s+|)/u, (t, i) => /^\s/u.test(t.slice(i, i + 1)) || /[⺀-￯]/u.test(charBefore(t, i, true)))) {
    if (!token) continue;
    const next = line + token;
    if (line && estimatedWidth(next, sizePx) > maxWidth) {
      lines.push(line.trimEnd());
      line = token.trimStart();
    } else {
      line = next;
    }
  }
  if (line) lines.push(line);
  return lines.length ? lines : [""];
}

/** The scene at 1 px per 9525 EMU, as a standalone SVG document string. */
export function coverSceneSvgMarkup(
  scene: CoverScene,
  values: Partial<Record<CoverFieldKey, string>>,
  fontFamily = "sans-serif",
): string {
  const width = scene.widthEmu / EMU_PER_PX;
  const height = scene.heightEmu / EMU_PER_PX;
  const defs: string[] = [];
  const body: string[] = [];
  scene.items.forEach((item, index) => {
    const x = item.x / EMU_PER_PX;
    const y = item.y / EMU_PER_PX;
    const w = Math.max(item.w / EMU_PER_PX, 0.01);
    const h = Math.max(item.h / EMU_PER_PX, 0.01);
    if (item.kind === "picture" && item.src) {
      const [cl, ct, cr, cb] = item.crop ?? [0, 0, 0, 0];
      const vw = Math.max(0.01, 1 - cl - cr);
      const vh = Math.max(0.01, 1 - ct - cb);
      const clip = item.paths?.length
        ? `<clipPath id="cover-clip-${index}"><path d="${escapeSvg(item.paths.map((path) => path.d).join(" "))}" transform="translate(${x} ${y}) scale(${w / 100} ${h / 100})"/></clipPath>`
        : "";
      if (clip) defs.push(clip);
      body.push(
        `<image href="${item.src}" x="${x - (cl / vw) * w}" y="${y - (ct / vh) * h}" width="${w / vw}" height="${h / vh}" preserveAspectRatio="none"${clip ? ` clip-path="url(#cover-clip-${index})"` : ""}/>`,
      );
    } else if (item.kind === "picture") {
      body.push(
        `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="#f2f2f2" stroke="#a6a6a6" stroke-dasharray="4 3"/>`,
      );
    } else if (item.kind === "shape") {
      const fillId = `cover-fill-${index}`;
      const lineId = `cover-line-${index}`;
      defs.push(gradientDefinition(item.fill, fillId));
      if (item.line) defs.push(gradientDefinition(item.line.paint, lineId));
      const stroke = item.line
        ? `${paintAttrs(item.line.paint, lineId, "stroke")} stroke-width="${item.line.widthPt * PX_PER_PT}" vector-effect="non-scaling-stroke"`
        : 'stroke="none"';
      const paths = (item.paths ?? [])
        .map(
          (path) =>
            `<path d="${escapeSvg(path.d)}" ${path.fill === "none" ? 'fill="none"' : paintAttrs(item.fill, fillId, "fill")} ${path.stroke === false ? 'stroke="none"' : stroke}/>`,
        )
        .join("");
      const flip = `${item.flipH ? " scale(-1 1) translate(-100 0)" : ""}${item.flipV ? " scale(1 -1) translate(0 -100)" : ""}`;
      const rotate = item.rotation
        ? `rotate(${item.rotation} ${x + w / 2} ${y + h / 2}) `
        : "";
      body.push(
        `<g transform="${rotate}translate(${x} ${y}) scale(${w / 100} ${h / 100})${flip}">${paths}</g>`,
      );
    }
    const text = item.text;
    if (!text) return;
    const [left, top, right, bottom] = text.inset.map((value) => value / EMU_PER_PX);
    const innerWidth = Math.max(0, w - left - right);
    const lines: Array<{ text: string; sizePx: number; color: string; weight: string; style: string; anchor: string; x: number; gap: number }> = [];
    for (const paragraph of text.paragraphs) {
      const first = paragraph.runs[0];
      const sizePx = (first?.sizePt ?? 11) * PX_PER_PT;
      const joined = paragraph.runs
        .map((run) => {
          const value = runText(run, values);
          return run.caps ? value.toUpperCase() : value;
        })
        .join("");
      const anchor =
        paragraph.align === "center" ? "middle" : paragraph.align === "right" ? "end" : "start";
      const lineX =
        anchor === "middle" ? x + left + innerWidth / 2 : anchor === "end" ? x + w - right : x + left;
      const wrapped = wrapWords(joined, sizePx, innerWidth);
      wrapped.forEach((line, lineIndex) => {
        lines.push({
          text: line,
          sizePx,
          color: first?.color ?? "#000000",
          weight: first?.bold ? "700" : "400",
          style: first?.italic ? "italic" : "normal",
          anchor,
          x: lineX,
          gap:
            (lineIndex === 0 ? (paragraph.beforePt ?? 0) * PX_PER_PT : 0) +
            (lineIndex === wrapped.length - 1 ? (paragraph.afterPt ?? 0) * PX_PER_PT : 0),
        });
      });
    }
    const totalHeight = lines.reduce((sum, line) => sum + line.sizePx * 1.2 + line.gap, 0);
    let cursor =
      text.anchor === "b"
        ? y + h - bottom - totalHeight
        : text.anchor === "ctr"
          ? y + top + (h - top - bottom - totalHeight) / 2
          : y + top;
    for (const line of lines) {
      cursor += line.sizePx;
      body.push(
        `<text x="${line.x}" y="${cursor}" font-size="${line.sizePx}" fill="${line.color}" font-weight="${line.weight}" font-style="${line.style}" text-anchor="${line.anchor}">${escapeSvg(line.text)}</text>`,
      );
      cursor += line.sizePx * 0.2 + line.gap;
    }
  });
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" font-family="${escapeSvg(fontFamily)}">` +
    `<rect width="${width}" height="${height}" fill="#ffffff"/>` +
    `<defs>${defs.join("")}</defs>` +
    body.join("") +
    "</svg>"
  );
}
