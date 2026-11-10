/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink to Shape in a document: the Office shape a pen stroke plainly
 * is, as a document shape.
 *
 * Microsoft Support, "Convert your ink to shapes and math equations": with
 * Ink to Shape on, a drawn rectangle, ellipse, triangle, diamond, other
 * polygon, line or arrow becomes that shape, outlined in the pen's colour
 * and width; anything else stays ink. The fit is the one Excel's and the
 * deck's Draw tabs use (`lib/office-ink-shape-fit.ts`); this states the
 * result the way a document stores a shape: a preset in the stroke's box, a
 * line or arrow whose direction is its flips, and any other polygon as the
 * closed custom geometry a rectangle carries (the DOCX writers keep a line
 * preset's own `prstGeom`, so a polygon is never a line kind). A highlighter
 * stroke that fits a shape becomes an ordinary outlined shape, as in Office.
 */

import {
  DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
  DOCUMENT_SHAPE_EMU_PER_PX,
  DOCUMENT_SHAPE_MIN_EDGE_EMU,
  normalizeDocumentShapeAttrs,
  type DocumentShapeAttrs,
  type DocumentShapePreset,
} from "./document-shapes";
import type { DocumentInkPoint, DocumentInkStrokeStyle } from "./document-ink";
import { inkHoldShapeFit } from "./office-ink-hold-shape";
import { fitInkStrokeToShape, inkSamplesForFitting, type InkShapeFit } from "./office-ink-shape-fit";
import { drawnShapeOutlineFractions } from "./office-drawn-shape-outline";
import {
  normalizePresentationCustomPath,
  PRESENTATION_CUSTOM_PATH_GRID,
  type PresentationCustomPathCommand,
} from "./presentation-freeform";

const PRESET_OF: Record<Exclude<InkShapeFit["kind"], "line" | "arrow" | "polygon">, DocumentShapePreset> = {
  rectangle: "rect",
  ellipse: "ellipse",
  diamond: "diamond",
  triangle: "triangle",
  "right-triangle": "rtTriangle",
};

/**
 * The document shape a stroke (EMU from the stroke's anchor origins, as
 * `documentInkStrokeAttrs` takes them) fits, outlined in its pen, or null
 * when the stroke stays ink. A highlighter's shape is drawn as the
 * highlighter draws: translucent, which the shape states as its opacity.
 */
export function documentShapeFromInkStroke(
  pointsEmu: readonly DocumentInkPoint[],
  pen: DocumentInkStrokeStyle,
  relativeHeight = DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
  /**
   * `stored`: a landed stroke's stored samples (read back at a pen's density
   * first). `held`: drawn and held, a shape the writer asked for -- where the
   * recognizer is unsure, the held stroke's own shape (a line, a polygon).
   */
  options: { readonly stored?: boolean; readonly held?: boolean } = {},
): DocumentShapeAttrs | null {
  const px = pointsEmu.map((point) => ({ x: point.x / DOCUMENT_SHAPE_EMU_PER_PX, y: point.y / DOCUMENT_SHAPE_EMU_PER_PX }));
  const fit =
    fitInkStrokeToShape(options.stored ? inkSamplesForFitting(px) : px) ?? (options.held ? inkHoldShapeFit(px) : null);
  if (!fit) return null;
  const emu = (value: number) => Math.round(value * DOCUMENT_SHAPE_EMU_PER_PX);
  // Outlined in the pen's colour and width, floating In Front of Text where
  // the stroke floated (a highlighter's wash is the stroke's, not a shape's).
  const outline = {
    fill: "none" as const,
    outline: pen.color,
    outlineWidthPt: Math.max(0.25, pen.widthPt),
    outlineStyle: "solid" as const,
    relativeHeight,
    wrapMode: "inFrontOfText" as const,
    zOrderStackComplete: true,
  };
  if (fit.kind === "line" || fit.kind === "arrow") {
    const left = Math.min(fit.start.x, fit.end.x);
    const top = Math.min(fit.start.y, fit.end.y);
    // A line is drawn corner to corner in its box: its box is exactly its
    // extent, zero across a level or upright line (Word's cy="0"). Raised to
    // the 16 px a new shape is drawn at, a level stroke became a line 16 px
    // out of level (WUA 091).
    return normalizeDocumentShapeAttrs({
      preset: fit.kind === "arrow" ? "lineArrow" : "line",
      ...outline,
      xEmu: emu(left),
      yEmu: emu(top),
      widthEmu: emu(Math.abs(fit.end.x - fit.start.x)),
      heightEmu: emu(Math.abs(fit.end.y - fit.start.y)),
      flipH: fit.start.x > fit.end.x,
      flipV: fit.start.y > fit.end.y,
    });
  }
  if (fit.kind === "polygon") {
    const box = fit.box;
    const grid = (value: number, origin: number, extent: number) =>
      Math.max(0, Math.min(PRESENTATION_CUSTOM_PATH_GRID, Math.round(((value - origin) / Math.max(extent, 1e-6)) * PRESENTATION_CUSTOM_PATH_GRID)));
    const commands: PresentationCustomPathCommand[] = fit.points.map((point, index) => ({
      type: index === 0 ? "moveTo" : "lineTo",
      x: grid(point.x, box.left, box.width),
      y: grid(point.y, box.top, box.height),
    }));
    const customPath = normalizePresentationCustomPath({ version: 1, closed: true, commands });
    if (!customPath) return null;
    return normalizeDocumentShapeAttrs({
      preset: "rect",
      customPath,
      ...outline,
      xEmu: emu(box.left),
      yEmu: emu(box.top),
      widthEmu: Math.max(DOCUMENT_SHAPE_MIN_EDGE_EMU, emu(box.width)),
      heightEmu: Math.max(DOCUMENT_SHAPE_MIN_EDGE_EMU, emu(box.height)),
    });
  }
  return normalizeDocumentShapeAttrs({
    preset: PRESET_OF[fit.kind],
    ...outline,
    xEmu: emu(fit.box.left),
    yEmu: emu(fit.box.top),
    widthEmu: Math.max(DOCUMENT_SHAPE_MIN_EDGE_EMU, emu(fit.box.width)),
    heightEmu: Math.max(DOCUMENT_SHAPE_MIN_EDGE_EMU, emu(fit.box.height)),
    flipH: fit.flipH === true,
  });
}

/** The presets Ink to Shape lands a stroke as (a polygon is a `rect` with custom geometry). */
const DRAWN_SHAPE_PRESETS: ReadonlySet<string> = new Set([...Object.values(PRESET_OF), "line", "lineArrow"]);











export function documentDrawnShapeOutline(
  value: unknown,
  hasStory = false,
): Array<{ x: number; y: number }> | null {
  if (hasStory || !value || typeof value !== "object") return null;
  const attrs = value as Readonly<Record<string, unknown>>;
  if (attrs.inkTool) return null;
  if (typeof attrs.preset !== "string" || !DRAWN_SHAPE_PRESETS.has(attrs.preset)) return null;
  if (attrs.fill !== "none" || attrs.fillPicture || attrs.fillGradient || attrs.fillPattern) return null;
  if (attrs.outline === "none" || attrs.isTextBox === true || attrs.wrapMode !== "inFrontOfText") return null;
  if (typeof attrs.text === "string" && attrs.text.trim() !== "") return null;
  if (attrs.textBody || (Array.isArray(attrs.textRuns) && attrs.textRuns.length > 0)) return null;
  if (Array.isArray(attrs.groupMembers) && attrs.groupMembers.length > 0) return null;
  const flips = { flipH: attrs.flipH === true, flipV: attrs.flipV === true };
  const custom = normalizePresentationCustomPath(attrs.customPath);
  if (custom) {
    if (attrs.preset !== "rect") return null;
    const polygon = custom.commands.flatMap((command) =>
      command.type === "moveTo" || command.type === "lineTo"
        ? [{ x: command.x / PRESENTATION_CUSTOM_PATH_GRID, y: command.y / PRESENTATION_CUSTOM_PATH_GRID }]
        : [],
    );
    return drawnShapeOutlineFractions("polygon", { polygon, closed: custom.closed, ...flips });
  }
  const kind = (Object.keys(PRESET_OF) as Array<keyof typeof PRESET_OF>).find((fit) => PRESET_OF[fit] === attrs.preset)
    ?? (attrs.preset === "lineArrow" ? "arrow" : "line");
  return drawnShapeOutlineFractions(kind, flips);
}
