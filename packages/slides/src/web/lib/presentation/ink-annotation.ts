/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Ink and annotation constants of the presentation editor: persisted ink
 * paths, the slideshow pen/highlighter colours, and the Record studio ink
 * colours and gesture authority types.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { normalizeInkCutEnds } from "@mrnote/core/web/lib/document-shapes.ts";
import { documentInkPressures } from "@mrnote/core/web/lib/document-ink.ts";
import type { AnnotationPoint, InkStroke, PersistedInkPath } from "@mrnote/core/web/lib/presentation-annotation.ts";
import { normalizeStrokeWidth } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { normalizePresentationCustomPath, PRESENTATION_CUSTOM_PATH_GRID, PRESENTATION_INK_MAX_POINTS } from "@mrnote/core/web/lib/presentation-freeform.ts";
import { drawnShapeOutlineFractions, type DrawnShapeKind } from "@mrnote/core/web/lib/office-drawn-shape-outline.ts";
import { parsePolygonString } from "@mrnote/core/web/lib/presentation-polygon-points.ts";
import type {
  PresentationBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import { AUTOFIT_EMU_PER_PX } from "./text-layout-css";
import { officeFramePlace } from "@mrnote/core/web/lib/office-object-frame.ts";

/** Authenticate persisted ink geometry and expand its normalized points to EMU. */
export function persistedInkPaths(blocks: readonly PresentationBlock[]): PersistedInkPath[] {
  const paths: PersistedInkPath[] = [];
  for (const block of blocks) {
    const value = block as unknown as Record<string, unknown>;
    if (value.shape !== "ink" || typeof value.points !== "string") continue;
    const left = Number(value.left);
    const top = Number(value.top);
    const width = Number(value.width);
    const height = Number(value.height);
    if (
      ![left, top, width, height].every(Number.isFinite) ||
      width < 0 ||
      height < 0
    ) {
      continue;
    }
    const normalized = parsePolygonString(value.points).slice(0, PRESENTATION_INK_MAX_POINTS);
    if (normalized.length === 0) continue;
    // The stroke as it is drawn: a rotated (or flipped) ink's points turn
    // about its centre with it, so the eraser meets the visible line, not
    // the trail the frame had before the turn (PPT-B235).
    const frame = { left, top, width, height, rotation: value.rotation, flipH: value.flipH, flipV: value.flipV };
    // The painted width, so an eraser meets the visible line (B410): the
    // slide draws a stroke's width in its design pixels.
    const strokeWidth = normalizeStrokeWidth(value.strokeWidth) * AUTOFIT_EMU_PER_PX;
    const cutEnds = normalizeInkCutEnds(value.inkCutEnds);
    // A pressure pen's reading rides with each point, so a piece the eraser
    // leaves keeps the width the pen drew it with.
    const pressures = documentInkPressures(value.inkPressure, normalized.length);
    paths.push({
      id: block.id,
      ...(strokeWidth > 0 ? { width: strokeWidth } : {}),
      ...(cutEnds ? { cutEnds } : {}),
      points: normalized.map((point, index) => ({
        ...officeFramePlace(frame, point.x / 100, point.y / 100),
        ...(pressures ? { p: pressures[index] } : {}),
      })),
    });
  }
  return paths;
}

/** The slide shape names Ink to Shape lands a stroke as (a polygon is custom geometry). */
const DRAWN_SHAPE_KINDS: Readonly<Record<string, DrawnShapeKind>> = {
  rectangle: "rectangle",
  ellipse: "ellipse",
  triangle: "triangle",
  "right-triangle": "right-triangle",
  diamond: "diamond",
  line: "line",
  arrow: "arrow",
  "custom-geometry": "polygon",
};








export function persistedDrawnShapePath(block: PresentationBlock): PersistedInkPath | null {
  const value = block as unknown as Record<string, unknown>;
  if (value.type !== "shape" || typeof value.shape !== "string") return null;
  const kind = DRAWN_SHAPE_KINDS[value.shape];
  if (!kind) return null;
  const lineLike = kind === "line" || kind === "arrow";
  if (!lineLike && value.fill !== "none") return null;
  if (value.stroke === "none" || value.fillPicture || value.fillGradient) return null;
  if ((typeof value.text === "string" && value.text.trim() !== "") || (Array.isArray(value.textRuns) && value.textRuns.length > 0)) return null;
  const left = Number(value.left);
  const top = Number(value.top);
  const width = Number(value.width);
  const height = Number(value.height);
  if (![left, top, width, height].every(Number.isFinite) || width < 0 || height < 0) return null;
  let polygon: { x: number; y: number }[] | undefined;
  let closed = true;
  if (kind === "polygon") {
    const custom = normalizePresentationCustomPath(value.customPath);
    if (!custom) return null;
    closed = custom.closed;
    polygon = custom.commands.flatMap((command) =>
      command.type === "moveTo" || command.type === "lineTo"
        ? [{ x: command.x / PRESENTATION_CUSTOM_PATH_GRID, y: command.y / PRESENTATION_CUSTOM_PATH_GRID }]
        : [],
    );
  }
  const outline = drawnShapeOutlineFractions(kind, {
    polygon,
    closed,
    flipH: value.flipH === true,
    flipV: value.flipV === true,
  });
  if (!outline) return null;
  // The outline already holds the mirror.
  const frame = { left, top, width, height, rotation: value.rotation };
  const strokeWidth = normalizeStrokeWidth(value.strokeWidth) * AUTOFIT_EMU_PER_PX;
  return {
    id: block.id,
    ...(strokeWidth > 0 ? { width: strokeWidth } : {}),
    points: outline.map((point) => officeFramePlace(frame, point.x, point.y, { mirrored: false })),
  };
}

/** The signature an eraser or lasso entry was taken with, read again from `block`. */
export function presentationInkEntrySignature(block: PresentationBlock, entry: { readonly whole?: boolean }): string | null {
  return entry.whole ? persistedDrawnShapeSignature(block) : persistedInkPathSignature(block);
}

/** A drawn shape's geometry as the gesture saw it: a write re-checks it. */
export function persistedDrawnShapeSignature(block: PresentationBlock): string | null {
  if (!persistedDrawnShapePath(block)) return null;
  const value = block as unknown as Record<string, unknown>;
  return JSON.stringify([
    block.type,
    value.shape,
    value.customPath ?? null,
    ...[value.left, value.top, value.width, value.height, value.rotation, value.flipH, value.flipV].map((v) => v ?? null),
  ]);
}

export function persistedInkPathSignature(block: PresentationBlock): string | null {
  const value = block as unknown as Record<string, unknown>;
  if (value.shape !== "ink" || typeof value.points !== "string") return null;
  // An open line the author drew or edited as a shape (Freeform, Curve, Edit
  // Points) is stored as "ink" too, with its outline in `customPath`. It is a
  // shape, not a pen stroke: Hide Ink and Delete All Ink leave it alone
  
  if (value.customPath !== undefined && value.customPath !== null) return null;
  const geometry = [value.left, value.top, value.width, value.height].map(Number);
  if (!geometry.every(Number.isFinite)) return null;
  return JSON.stringify([block.type, value.shape, value.points, ...geometry]);
}

/**
 * Present-mode ANNOTATION palette. The default ink/laser colour is PowerPoint's
 * laser red; the swatches give the presenter a small, high-contrast set that
 * reads on both light and dark slides. Session-only — never persisted.
 */
export const ANNOTATION_DEFAULT_COLOR = "#ff2d2d";

/** Stable empty stroke list so the draw layer's `strokes` prop is referentially
 * stable for a slide with no ink (avoids a fresh `[]` per render). */
export const EMPTY_INK_STROKES: readonly InkStroke[] = [];

export const ANNOTATION_COLORS: readonly string[] = [
  "#ff2d2d", // red (laser default)
  "#ffd400", // yellow (highlighter-friendly)
  "#22c55e", // green
  "#3b82f6", // blue
  "#111827", // near-black (light slides)
  "#ffffff", // white (dark slides)
];



export const ANNOTATION_COLOR_NAME_KEYS: Readonly<Record<string, string>> = {
  "#ff2d2d": "red",
  "#ffd400": "yellow",
  "#22c55e": "green",
  "#3b82f6": "blue",
  "#111827": "black",
  "#ffffff": "white",
};

/** PowerPoint Recording Studio's durable pen/highlighter palette. */
export const RECORD_STUDIO_INK_COLORS = [
  { id: "white", value: "#ffffff" },
  { id: "black", value: "#111827" },
  { id: "dark-red", value: "#991b1b" },
  { id: "red", value: "#ff2d2d" },
  { id: "orange", value: "#f97316" },
  { id: "yellow", value: "#ffd400" },
  { id: "light-green", value: "#86efac" },
  { id: "dark-green", value: "#15803d" },
  { id: "light-blue", value: "#7dd3fc" },
  { id: "medium-blue", value: "#3b82f6" },
  { id: "dark-blue", value: "#1d4ed8" },
  { id: "purple", value: "#7e22ce" },
] as const;

type RecordStudioInkAuthorityEntry = {
  readonly block: PresentationBlock;
  readonly id: string;
  readonly index: number;
  readonly path: PersistedInkPath;
  readonly pathSignature: string;
};

export type RecordStudioInkGesture = {
  readonly mode: "draw" | "eraser";
  readonly deckGeneration: number;
  readonly slideIndex: number;
  readonly slideId: string;
  readonly slide: SlideModel;
  readonly blocks: PresentationBlock[];
  readonly tool: "pen" | "highlighter" | "eraser";
  readonly color: string;
  points: AnnotationPoint[];
  readonly inkAuthority?: readonly RecordStudioInkAuthorityEntry[];
};
