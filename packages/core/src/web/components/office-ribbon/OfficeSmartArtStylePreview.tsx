"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * ONE tile of the fourteen-style gallery: a real three-node diagram, built by
 * the ENGINE with that quick style and drawn by the same shape renderer the
 * host's canvas uses, so a tile cannot drift from what applying the style
 * actually does.
 *
 * Pure SVG for the reason {@link OfficeSmartArtSeedPreview} documents: a
 * `foreignObject` scene is scaled by the viewBox in Chromium but drawn at 1:1
 * CSS pixels (and clipped) by WebKit. Labels are empty because a 58px tile
 * cannot show text, and the diagram camera becomes the affine approximation
 * (SMARTART-STYLES §5) rather than the canvas's CSS 3-D transform.
 *
 * `renderShape` is the host's own richer renderer (the Presentation canvas
 * resolves native theme paint, gradients and pictures). Without it the preset
 * is drawn from the shared DrawingML default geometry PLUS the style's own
 * `shape3d` (extruded body, contour) through the same helpers the canvases use,
 * so a host needs no component of its own and still gets a 3-D tile.
 */

import { useId, type ReactNode } from "react";

import { OfficeSmartArtPresetGraphic } from "./OfficeSmartArtSeedPreview";
import {
  smartArtCameraSvgPoint,
  smartArtCameraSvgTransform,
  smartArtContourStroke,
  smartArtBodyDepthEmu,
  smartArtExtrusionOffset,
  smartArtShape3dFilter,
} from "../../lib/office-smartart/canvas-3d";
import { gradientVector } from "../../lib/presentation-shapes";
import {
  DEFAULT_SMARTART_STYLE,
  type OfficeSmartArtShapeNode,
  type SmartArtKind,
  type SmartArtStyle,
} from "../../lib/office-smartart/model";
import { useSmartArtPreview } from "../../lib/office-smartart/remote-previews";
import {
  smartArtGradientOfficePaint,
  type SmartArtScene3d,
  type SmartArtStyleId,
} from "../../lib/office-smartart/styles-model";

/** EMU per CSS pixel (96 dpi), the unit the preview viewBox is drawn in. */
const EMU_PER_CSS_PIXEL = 9525;

/** The frame one style tile lays its three-node diagram out in. */
export const OFFICE_SMARTART_STYLE_PREVIEW_BOUNDS = Object.freeze({
  width: 5_800_000,
  height: 4_400_000,
});

/**
 * The width a thumbnail's diagram is LAID OUT at, in EMU, before it is fitted
 * to the tile. Laid out at the tile's own pixel size (58 x 44) a cycle's
 * nodes came out a few pixels wide and kept to one corner; laid out at a
 * normal diagram size with the tile's aspect, every layout has its real
 * proportions, and the fit below fills the tile with it.
 */
const OFFICE_SMARTART_STYLE_PREVIEW_LAYOUT_WIDTH = 5_800_000;
/** The share of the tile left free on each side of the fitted diagram. */
const OFFICE_SMARTART_STYLE_PREVIEW_MARGIN = 0.1;

/** The layout a style thumbnail previews when nothing is selected to mirror. */
export const OFFICE_SMARTART_STYLE_PREVIEW_KIND: SmartArtKind = "process";

export interface OfficeSmartArtStylePreviewProps {
  styleId: SmartArtStyleId;
  /** The layout to draw — normally the selected diagram's own. */
  kind: SmartArtKind;
  bounds?: { width: number; height: number };
  style?: SmartArtStyle;
  /** Draw one shape node the host's own way instead of the shared preset. */
  renderShape?: (node: OfficeSmartArtShapeNode) => ReactNode;
}

export function OfficeSmartArtStylePreview({
  styleId,
  kind,
  bounds = OFFICE_SMARTART_STYLE_PREVIEW_BOUNDS,
  style = DEFAULT_SMARTART_STYLE,
  renderShape,
}: OfficeSmartArtStylePreviewProps) {
  const prefix = useId().replace(/[^A-Za-z0-9_-]/g, "");
  // `bounds` is the TILE; the diagram is laid out at a normal size with the
  // tile's aspect and fitted to it below. The engine is the server's: a
  // gallery's tiles are asked for together, and a tile draws its empty frame
  // until its scene arrives.
  const aspect = bounds.width > 0 && bounds.height > 0 ? bounds.height / bounds.width : 0.75;
  const layout = {
    width: OFFICE_SMARTART_STYLE_PREVIEW_LAYOUT_WIDTH,
    height: Math.round(OFFICE_SMARTART_STYLE_PREVIEW_LAYOUT_WIDTH * aspect),
  };
  const blocks =
    useSmartArtPreview({ kind, items: ["", "", ""], bounds: layout, style, styleId }, prefix) ?? [];
  const width = layout.width / EMU_PER_CSS_PIXEL;
  const height = layout.height / EMU_PER_CSS_PIXEL;
  const scene3d = blocks.reduce<SmartArtScene3d | undefined>(
    (found, block) =>
      found ?? (block.type === "image" ? undefined : block.scene3d),
    undefined,
  );
  const camera = scene3d
    ? smartArtCameraSvgTransform(scene3d, width / 2, height / 2)
    : null;
  // A strong camera (3d9's roll and tilt) carries a diagram laid out edge to
  // edge past the tile's frame; PowerPoint's own layout leaves a margin that
  // ours does not. Measure what the camera does to every member — its extruded
  // body included, since that is drawn too — and scale the result back inside
  // the frame about its centre. A tile that already fits is left untouched.
  const project = smartArtCameraSvgPoint(scene3d, width / 2, height / 2);
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const block of blocks) {
    if (block.type !== "shape") continue;
    const left = block.left / EMU_PER_CSS_PIXEL;
    const top = block.top / EMU_PER_CSS_PIXEL;
    const right = left + block.width / EMU_PER_CSS_PIXEL;
    const bottom = top + block.height / EMU_PER_CSS_PIXEL;
    const body = block.shape3d
      ? smartArtExtrusionOffset(
          block.scene3d,
          smartArtBodyDepthEmu(block.shape3d) / EMU_PER_CSS_PIXEL,
        )
      : null;
    const corners: Array<[number, number]> = [
      [left, top],
      [right, top],
      [left, bottom],
      [right, bottom],
    ];
    if (body) {
      for (const [x, y] of [...corners]) corners.push([x + body.dx, y + body.dy]);
    }
    for (const [x, y] of corners) {
      const point = project(x, y);
      minX = Math.min(minX, point.x);
      minY = Math.min(minY, point.y);
      maxX = Math.max(maxX, point.x);
      maxY = Math.max(maxY, point.y);
    }
  }
  // Fill the tile with the drawing -- enlarged as well as reduced -- and
  // centre it, as Office's SmartArt Styles gallery draws each thumbnail.
  // Before, only a drawing that overflowed was scaled: a small one kept its
  // layout position and size, a few pixels in the tile's top-left corner.
  const margin = OFFICE_SMARTART_STYLE_PREVIEW_MARGIN;
  const fit =
    Number.isFinite(minX) && Number.isFinite(maxX)
      ? Math.min(
          (width * (1 - 2 * margin)) / Math.max(maxX - minX, 1e-6),
          (height * (1 - 2 * margin)) / Math.max(maxY - minY, 1e-6),
        )
      : 1;
  const framed =
    Number.isFinite(minX) && Number.isFinite(maxX)
      ? `translate(${(width / 2).toFixed(3)} ${(height / 2).toFixed(3)}) scale(${fit.toFixed(5)}) translate(${(
          -(minX + maxX) / 2
        ).toFixed(3)} ${(-(minY + maxY) / 2).toFixed(3)})`
      : null;
  // Screen pixels per layout pixel, for effects that must stay visible at
  // thumbnail size (a style's shadow is a few pixels on a full diagram).
  const screenScale = (bounds.width / EMU_PER_CSS_PIXEL / width) * fit;
  const groupTransform = [framed, camera].filter(Boolean).join(" ") || null;
  return (
    <svg
      viewBox={`0 0 ${width} ${height}`}
      width="100%"
      height="100%"
      // Inline, so no host's ribbon rule (`button svg { width: 16px }` and
      // the like) can resize the thumbnail or the members inside it.
      style={{ display: "block", height: "100%", maxHeight: "none", maxWidth: "none", width: "100%" }}
      aria-hidden="true"
      data-smartart-style={styleId}
    >
      <defs>
        {blocks.map((block) => {
          // The style's gradient (Subtle, Moderate, Intense and the 3-D
          // fills): the tile used to paint the flat base colour only, so
          // styles that differ by their ramp looked identical.
          if (block.type !== "shape" || !block.fillGradient || renderShape) return null;
          const paint = smartArtGradientOfficePaint(block.fillGradient);
          return (
            <linearGradient
              key={block.id}
              id={`${prefix}-fill-${block.id}`}
              {...gradientVector(paint.angle)}
            >
              {paint.stops.map((stop) => (
                <stop key={stop.position} offset={stop.position} stopColor={stop.color} />
              ))}
            </linearGradient>
          );
        })}
      </defs>
      <g {...(groupTransform ? { transform: groupTransform } : {})}>
        {[...blocks]
          .sort((a, b) => Number(a.zIndex ?? 0) - Number(b.zIndex ?? 0))
          .map((block) => {
            if (block.type !== "shape") return null;
            // A tile has to carry the style's 3-D too, not only its camera: a
            // host that draws its own shapes (Presentation) brings the extruded
            // body and the contour with it, but the shared preset drawing has
            // neither, so on every other host a "3-D" tile was a flat rectangle
            // under a tilt. Same helpers the canvases use, so a tile cannot
            // paint something applying the style would not.
            // The style's shadow is drawn by the tile (below) for every host,
            // so a host renderer gets the member without it: at tile scale the
            // host's own few-pixel shadow shrinks to nothing.
            const shadow = block.shadow && block.shapeShadow ? block.shapeShadow : null;
            const own = renderShape?.(
              shadow ? { ...block, shadow: false, shapeShadow: undefined } : block,
            );
            const shape3d = block.shape3d;
            // Only the extruded body is a CSS filter here. The bevel rim is an
            // SVG filter carried by the shared preset drawing itself, so it
            // clips to the real silhouette (see OfficeSmartArtBevelFilter).
            const bodyFilter = own
              ? null
              : shape3d
                ? smartArtShape3dFilter(shape3d, block.fill, block.scene3d)
                : null;
            // The style's own shadow (Intense Effect), kept visible at tile
            // size: at least 1px offset and 1px blur on screen, and a little
            // lighter than the canvas's, or a small arrow turns into a smudge.
            const shadowFilter = shadow
              ? `drop-shadow(0 ${Math.max(shadow.distance, 1 / screenScale).toFixed(2)}px ${Math.max(shadow.blur / 2, 1 / screenScale).toFixed(2)}px rgba(0, 0, 0, ${Math.min(0.45, shadow.opacity).toFixed(2)}))`
              : null;
            const filter = [bodyFilter, shadowFilter].filter(Boolean).join(" ") || null;
            const contour = own || !shape3d ? null : smartArtContourStroke(shape3d);
            // Never overwrite a line the style itself draws (the canvases hold
            // to the same rule).
            const outlined0 =
              contour && (!block.stroke || block.stroke === "none")
                ? { ...block, stroke: contour.stroke, strokeWidth: contour.strokeWidth }
                : block;
            const outlined =
              !own && block.fillGradient
                ? { ...outlined0, fill: `url(#${prefix}-fill-${block.id})` }
                : outlined0;
            return (
              <svg
                key={block.id}
                x={block.left / EMU_PER_CSS_PIXEL}
                y={block.top / EMU_PER_CSS_PIXEL}
                width={block.width / EMU_PER_CSS_PIXEL}
                height={block.height / EMU_PER_CSS_PIXEL}
                viewBox={`0 0 ${block.width / EMU_PER_CSS_PIXEL} ${block.height / EMU_PER_CSS_PIXEL}`}
                overflow="visible"
                style={{
                  height: block.height / EMU_PER_CSS_PIXEL,
                  maxHeight: "none",
                  maxWidth: "none",
                  width: block.width / EMU_PER_CSS_PIXEL,
                  ...(filter ? { filter } : {}),
                }}
              >
                {own ?? (
                  <OfficeSmartArtPresetGraphic
                    node={outlined}
                    width={Math.max(block.width / EMU_PER_CSS_PIXEL, 0.01)}
                    height={Math.max(block.height / EMU_PER_CSS_PIXEL, 0.01)}
                  />
                )}
              </svg>
            );
          })}
      </g>
    </svg>
  );
}
