/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PresentationShapeBlock,
  PresentationTextBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  presentationShapeObjectTextPatch,
  type PresentationShapeObjectPaint,
  type PresentationTextFramePaintPatch,
} from "@mrnote/core/web/lib/presentation-shape-style-adapter.ts";
import { officeMenuDashFields, type OfficeMenuDash } from "@mrnote/core/web/lib/office-stroke-dash.ts";

type Stops = { readonly stops?: ReadonlyArray<{ readonly color?: string }>; readonly angle?: number };

const firstStop = (gradient: Stops | undefined) => gradient?.stops?.[0]?.color;
const lastStop = (gradient: Stops | undefined) => gradient?.stops?.[gradient.stops.length - 1]?.color;

/** The text frame read as an autoshape: its fill, line and shadow in the autoshape's names. */
export function presentationTextFrameAsShapeView(block: PresentationTextBlock): PresentationShapeBlock {
  const record = block as unknown as Record<string, unknown>;
  return {
    ...record,
    type: "shape",
    shape: "rect",
    fill: typeof block.fill === "string" ? block.fill : "none",
    ...(typeof block.fillTo === "string" ? { fillTo: block.fillTo } : {}),
    ...(typeof block.gradientAngle === "number" ? { gradientAngle: block.gradientAngle } : {}),
    stroke:
      typeof block.outlineColor === "string" && (block.outlineWidth ?? 0) > 0 ? block.outlineColor : "none",
    ...(typeof record.outlineColorTo === "string" ? { strokeTo: record.outlineColorTo } : {}),
    ...(typeof block.outlineWidth === "number" ? { strokeWidth: block.outlineWidth } : {}),
    ...(block.outlineStyle ? { strokeStyle: block.outlineStyle } : {}),
    ...(block.shadow ? { shadow: true } : {}),
  } as unknown as PresentationShapeBlock;
}

/**
 * An autoshape patch (what a Format Shape change wrote on the frame's view)
 * in the text frame's own field names; empty when none of it is anything a
 * frame holds.
 */
export function presentationTextFramePatchFromShapePatch(
  block: PresentationTextBlock,
  patch: Partial<PresentationShapeBlock>,
): PresentationTextFramePaintPatch {
  const record = patch as Record<string, unknown>;
  const paint: PresentationShapeObjectPaint = {};
  const fillGradient = record.fillGradient as Stops | undefined;
  const fillPattern = record.fillPattern as { foreground?: string } | undefined;
  const fillChanged = "fill" in record || "fillGradient" in record || "fillPattern" in record;
  if (fillChanged) {
    const color = firstStop(fillGradient) ?? fillPattern?.foreground ?? (record.fill as string | undefined);
    paint.fill = !color || color === "none" ? null : color;
  }
  const strokeGradient = record.strokeGradient as Stops | undefined;
  if ("stroke" in record || "strokeGradient" in record) {
    const color = firstStop(strokeGradient) ?? (record.stroke as string | undefined);
    paint.outlineColor = !color || color === "none" ? null : color;
  }
  if (typeof record.strokeWidth === "number") paint.outlineWidth = record.strokeWidth;
  if (record.outlineDash !== undefined) {
    paint.outlineDash = record.outlineDash as OfficeMenuDash;
  } else if (typeof record.strokeStyle === "string") {
    paint.outlineStyle = record.strokeStyle as PresentationShapeObjectPaint["outlineStyle"];
  }
  // A shadow: the legacy switch, or the Effects pane's shadow.
  if ("shapeEffects" in record) {
    paint.shadow = Boolean((record.shapeEffects as { shadow?: unknown } | undefined)?.shadow);
  } else if ("shadow" in record) {
    paint.shadow = record.shadow === true;
  }
  if (Object.keys(paint).length === 0) return {};
  const next: Record<string, unknown> = { ...presentationShapeObjectTextPatch(block, paint) };
  // A gradient keeps its two ends and its angle, the frame's own gradient.
  if (fillChanged && paint.fill) {
    const to = lastStop(fillGradient) ?? (fillGradient ? undefined : (record.fillTo as string | undefined));
    if (to && to !== paint.fill) {
      next.fillTo = to;
      const angle = fillGradient?.angle ?? (record.gradientAngle as number | undefined);
      if (typeof angle === "number") next.gradientAngle = angle;
    }
  }
  if (paint.outlineColor && strokeGradient) {
    const to = lastStop(strokeGradient);
    if (to && to !== paint.outlineColor) {
      next.outlineColorTo = to;
      if (typeof strokeGradient.angle === "number") next.outlineGradientAngle = strokeGradient.angle;
    }
  }
  if (paint.outlineDash !== undefined && !next.outlineStyle) {
    next.outlineStyle = officeMenuDashFields(paint.outlineDash).style;
  }
  return next as PresentationTextFramePaintPatch;
}
