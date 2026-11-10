/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * WordArt (艺术字) on a 2-D canvas: a text box's raster appearance drawn
 * with its Office paint, as the server's vector appearance draws it
 * (text_box_wordart.py) -- glow and shadow under the words, an outline twice
 * its width under the fill, the fill flat, see-through, a gradient over the
 * box or Office's pattern. The canvas is in points (the paint's 96-dpi pixel
 * lengths x 0.75); `deviceScale` is the backing's pixels per point, which
 * canvas shadows are measured in.
 */
import type { OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import { officeFillPatternTile } from "@mrnote/core/web/lib/office-chart-shape-paint.ts";
import { officeGradientAngleForBox, officePathGradientFocus } from "@mrnote/core/web/lib/office-shape-paint.ts";

const PX = 0.75;

function rgba(hex: string, alpha = 1): string {
  const value = (offset: number) => Number.parseInt(hex.slice(offset, offset + 2), 16);
  return `rgba(${value(1)}, ${value(3)}, ${value(5)}, ${alpha})`;
}

export function pdfWordArtCanvasPainter(
  ctx: CanvasRenderingContext2D,
  paint: OfficeTextPaint,
  frame: Readonly<{ width: number; height: number }>,
  deviceScale: number,
): (text: string, x: number, baseline: number) => void {
  const fillStyle = (): string | CanvasGradient | CanvasPattern => {
    const gradient = paint.fillGradient;
    if (gradient) {
      let ramp: CanvasGradient;
      if (gradient.path) {
        const focus = officePathGradientFocus(gradient.path.fillToRect);
        const cx = focus.x * frame.width;
        const cy = focus.y * frame.height;
        const radius = Math.max(...[[0, 0], [frame.width, 0], [0, frame.height], [frame.width, frame.height]]
          .map(([x, y]) => Math.hypot(cx - x!, cy - y!)));
        ramp = ctx.createRadialGradient(cx, cy, 0, cx, cy, Math.max(radius, 0.01));
      } else {
        // CSS angle: 0 points up, 90 right; the line spans the box.
        const angle = (officeGradientAngleForBox(gradient, frame.width, frame.height) * Math.PI) / 180;
        const dx = Math.sin(angle);
        const dy = -Math.cos(angle);
        const half = (Math.abs(frame.width * dx) + Math.abs(frame.height * dy)) / 2;
        const cx = frame.width / 2;
        const cy = frame.height / 2;
        ramp = ctx.createLinearGradient(cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half);
      }
      for (const stop of gradient.stops) ramp.addColorStop(stop.position, rgba(stop.color, stop.alpha ?? 1));
      return ramp;
    }
    if (paint.fillPattern && typeof document !== "undefined") {
      const tile = officeFillPatternTile(paint.fillPattern);
      const canvas = document.createElement("canvas");
      canvas.width = tile.size;
      canvas.height = tile.size;
      const tileContext = canvas.getContext("2d");
      if (tileContext) {
        tileContext.fillStyle = paint.fillPattern.background;
        tileContext.fillRect(0, 0, tile.size, tile.size);
        tileContext.fillStyle = paint.fillPattern.foreground;
        for (const mark of tile.marks) {
          if (mark.tag === "rect") tileContext.fillRect(mark.x, mark.y, mark.width, mark.height);
        }
        const pattern = ctx.createPattern(canvas, "repeat");
        if (pattern) {
          pattern.setTransform?.(new DOMMatrix().scale(PX));
          return pattern;
        }
      }
    }
    return rgba(paint.fill === "none" ? "#000000" : paint.fill, paint.fillOpacity ?? 1);
  };
  const fill = paint.fill === "none" ? null : fillStyle();
  return (text, x, baseline) => {
    if (!text.trim()) return;
    const soft = (color: string, opacity: number, blur: number, offsetX: number, offsetY: number) => {
      ctx.save();
      ctx.fillStyle = rgba(color, 1);
      ctx.shadowColor = rgba(color, opacity);
      ctx.shadowBlur = blur * PX * deviceScale;
      // Drawn off the canvas, its shadow alone lands on the words' place.
      const away = 100_000;
      ctx.shadowOffsetX = (offsetX * PX + away) * deviceScale;
      ctx.shadowOffsetY = offsetY * PX * deviceScale;
      ctx.fillText(text, x - away, baseline);
      ctx.restore();
    };
    if (paint.glow) soft(paint.glow.color, paint.glow.opacity, paint.glow.radius * 2, 0, 0);
    if (paint.shadow) soft(paint.shadow.color, paint.shadow.opacity, paint.shadow.blur, paint.shadow.offsetX, paint.shadow.offsetY);
    if (paint.bevel) {
      soft("#ffffff", 0.5, 0, -0.6, -0.6);
      soft("#000000", 0.31, 0, 0.6, 0.6);
      if (paint.bevel.depth) soft("#000000", 0.22, 0, paint.bevel.depth / 2, paint.bevel.depth / 2);
    }
    if (paint.outline !== "none" && paint.outlineWidthPt > 0) {
      ctx.save();
      ctx.strokeStyle = paint.outline;
      ctx.lineWidth = fill ? paint.outlineWidthPt * 2 : paint.outlineWidthPt;
      ctx.lineJoin = "round";
      ctx.strokeText(text, x, baseline);
      ctx.restore();
    }
    if (fill) {
      ctx.save();
      ctx.fillStyle = fill;
      ctx.fillText(text, x, baseline);
      ctx.restore();
    }
  };
}
