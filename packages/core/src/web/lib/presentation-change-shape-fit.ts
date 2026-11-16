/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeDocumentShapeAdjustments } from "./document-shapes";
import { effectiveAutofitMode } from "./presentation-autofit";
import { presentationShapeCatalogEntry } from "./presentation-shape-catalog";
import { authoredShapeTextRectFractions, shapeTextRectFractions } from "./presentation-shape-text-rect";
import { normalizeShapeKind, type ShapeKind } from "./presentation-shapes";
import { presentationTextFitFrame, type TextFitAnchor } from "./presentation-text-fit";
import { shapeLabelVerticalAlign } from "./presentation-vertical-anchor";

type FitBlock = {
  type?: unknown;
  shape?: unknown;
  left?: unknown;
  top?: unknown;
  width?: unknown;
  height?: unknown;
  rotation?: unknown;
  adjustments?: unknown;
  source?: unknown;
  style?: unknown;
};

/**
 * Change Shape that PowerPoint does not change: a TEXT FRAME to a Rectangle.
 * A text frame already draws the frame's rectangle (`prstGeom rect`, the
 * gallery shows Rectangle as its current preset), so the box, its text-box
 * identity and its file stay exactly as they are.
 */
export function presentationTextFrameChangeShapeIsNoop(kind: ShapeKind): boolean {
  return kind === "rectangle";
}

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/** The text rectangle's share of the frame's height for this geometry. */
function textRectHeightShare(
  block: FitBlock,
  kind: ShapeKind | null,
  width: number,
  height: number,
  imported: boolean,
): number {
  if (!kind) return 1;
  const stated = imported ? shapeTextRectFractions(block.source, kind) : null;
  const entry = presentationShapeCatalogEntry(kind);
  const rect =
    stated ??
    authoredShapeTextRectFractions(
      entry?.geometry === "preset" ? (entry as { drawingMlPreset: string }).drawingMlPreset : null,
      width,
      height,
      normalizeDocumentShapeAdjustments(block.adjustments),
    );
  const share = rect ? rect.bottom - rect.top : 1;
  return share > 0 && Number.isFinite(share) ? share : 1;
}

/**
 * The frame an autofit box takes when Change Shape gives it `kind`, or null
 * when it keeps its frame (not an autofit box, unplaced, or no change worth a
 * design px). `block` is the box as it stands -- a text frame (whose
 * rectangle is the frame) or a shape of its current kind.
 */
export function presentationChangeShapeAutofitFrame(
  block: FitBlock,
  kind: ShapeKind,
): { left: number; top: number; width: number; height: number } | null {
  if (effectiveAutofitMode(block) !== "resize-shape") return null;
  const left = finite(block.left);
  const top = finite(block.top);
  const width = finite(block.width);
  const height = finite(block.height);
  if (left === null || top === null || width === null || height === null || !(width > 0) || !(height > 0)) {
    return null;
  }
  const fromKind = block.type === "shape" ? normalizeShapeKind(block.shape) : null;
  const words = height * textRectHeightShare(block, fromKind, width, height, true);
  let next = words;
  for (let step = 0; step < 64; step += 1) {
    const solved = words / textRectHeightShare(block, kind, width, next, false);
    if (Math.abs(solved - next) < 0.5) {
      next = solved;
      break;
    }
    next = solved;
  }
  if (!(next > 0) || !Number.isFinite(next)) return null;
  const style = (block.style && typeof block.style === "object" ? block.style : {}) as {
    verticalAlign?: unknown;
  };
  const vertical =
    block.type === "shape"
      ? (shapeLabelVerticalAlign(style, block.source) ?? "middle")
      : style.verticalAlign === "middle" || style.verticalAlign === "bottom"
        ? style.verticalAlign
        : "top";
  return presentationTextFitFrame(
    { left, top, width, height, rotation: finite(block.rotation) ?? 0 },
    { height: next },
    { anchor: vertical as TextFitAnchor, align: "left", wrap: true },
  );
}
