/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PdfNativeVectorObjectStyle,
  PdfNativeVectorObjectStyleChanges,
  PdfNativeVectorObjectStyleLineCap,
  PdfNativeVectorObjectStyleLineJoin,
  PdfNativeVectorRgbColor,
} from "./engine/native/vector-style-transaction";
import type { PdfStandardBlendMode } from "./blend-modes";









export interface PdfNativeVectorStyleIntent {
  readonly strokeColor?: PdfNativeVectorRgbColor;
  readonly lineWidth?: number;
  /** A colour turns the fill on; null turns it off. */
  readonly fillColor?: PdfNativeVectorRgbColor | null;
  readonly opacity?: number;
  readonly blendMode?: PdfStandardBlendMode;
  readonly lineCap?: PdfNativeVectorObjectStyleLineCap;
  readonly lineJoin?: PdfNativeVectorObjectStyleLineJoin;
  readonly lineStyle?: "solid" | "dotted" | "dashed";
}

export interface PdfNativeVectorStyleIntentTarget {
  readonly paintMode: "fill" | "stroke" | "fill-stroke" | string;
  /** How much a line width of the drawing's own space grows on the page:
   * the intent's width is in points on paper, written divided by it (user
   * audit PDF 0058: 6 pt asked of a drawing scaled to 0.65 drew 3.9 pt). */
  readonly lineWidthScale?: number;
  readonly style: Readonly<{
    style: Pick<PdfNativeVectorObjectStyle, "fillColor" | "lineWidth"> &
      Partial<PdfNativeVectorObjectStyle>;
  }>;
}

const sameColor = (
  a: PdfNativeVectorRgbColor | null | undefined,
  b: PdfNativeVectorRgbColor | null | undefined,
) =>
  Boolean(a && b) &&
  Math.abs(a!.r - b!.r) < 1e-6 &&
  Math.abs(a!.g - b!.g) < 1e-6 &&
  Math.abs(a!.b - b!.b) < 1e-6;

/**
 * Office does nothing when a command asks for what the object already has
 * (Solid on a solid line, No Fill on an outline, Round on a round cap). The
 * writer refuses an unchanged edit, so such channels are dropped here and an
 * empty result means there is nothing to write.
 */
function withoutUnchanged(
  changes: PdfNativeVectorObjectStyleChanges,
  target: PdfNativeVectorStyleIntentTarget,
): PdfNativeVectorObjectStyleChanges {
  const own = target.style.style;
  const out: Record<string, unknown> = { ...changes };
  if (out.paintMode === target.paintMode) delete out.paintMode;
  if (sameColor(changes.fillColor, own.fillColor)) delete out.fillColor;
  if (sameColor(changes.strokeColor, own.strokeColor)) delete out.strokeColor;
  for (const key of ["lineWidth", "lineCap", "lineJoin", "opacity", "blendMode"] as const) {
    if (changes[key] !== undefined && changes[key] === own[key]) delete out[key];
  }
  const ownDash = own.dashArray;
  if (
    changes.dashArray !== undefined &&
    ownDash &&
    changes.dashArray.length === ownDash.length &&
    changes.dashArray.every((value, index) => Math.abs(value - ownDash[index]!) < 1e-6) &&
    (changes.dashPhase ?? 0) === (own.dashPhase ?? 0)
  ) {
    delete out.dashArray;
    delete out.dashPhase;
  }
  return out as PdfNativeVectorObjectStyleChanges;
}

export function nativeVectorStyleChangesForTarget(
  intent: PdfNativeVectorStyleIntent,
  target: PdfNativeVectorStyleIntentTarget,
): PdfNativeVectorObjectStyleChanges {
  const own = target.style.style;
  const scale =
    typeof target.lineWidthScale === "number" &&
    Number.isFinite(target.lineWidthScale) &&
    target.lineWidthScale > 0
      ? target.lineWidthScale
      : 1;
  // A width on paper in the drawing's own space, to the writer's precision.
  const ownWidth = (points: number) =>
    scale === 1 ? points : Math.round((points / scale) * 1_000_000) / 1_000_000;
  let changes: {
    -readonly [K in keyof PdfNativeVectorObjectStyleChanges]: PdfNativeVectorObjectStyleChanges[K];
  } = {};
  if (intent.strokeColor) changes.strokeColor = intent.strokeColor;
  if (typeof intent.lineWidth === "number") changes.lineWidth = ownWidth(intent.lineWidth);
  // An outline colour, weight or dash asked of a fill-only object turns its
  // outline on, as Shape Outline does on a shape with No Outline in Office:
  // it repaints as fill+stroke, black and 0.75 pt unless the request names
  // them (a width the page left at 0 would draw a hairline).
  if (
    target.paintMode === "fill" &&
    intent.fillColor !== null &&
    (intent.strokeColor || typeof intent.lineWidth === "number" || intent.lineStyle)
  ) {
    changes.paintMode = "fill-stroke";
    changes.strokeColor ??= { r: 0, g: 0, b: 0 };
    if (changes.lineWidth === undefined && !((own.lineWidth ?? 0) > 0)) {
      changes.lineWidth = ownWidth(0.75);
    }
  }
  if (intent.fillColor) {
    changes.fillColor = intent.fillColor;
    // Fill switched on: an outline-only object repaints as fill+stroke so
    // its outline survives with the new fill.
    if (target.paintMode === "stroke") changes.paintMode = "fill-stroke";
  } else if (intent.fillColor === null) {
    changes.paintMode = "stroke";
    // A fill-only object has no outline of its own: it keeps its fill colour
    // as the outline it becomes. An object that already strokes keeps its.
    if (target.paintMode === "fill" && own.fillColor && !intent.strokeColor) {
      changes.strokeColor = own.fillColor;
    }
  }
  if (typeof intent.opacity === "number") changes.opacity = intent.opacity;
  if (intent.blendMode) changes.blendMode = intent.blendMode;
  if (intent.lineCap) changes.lineCap = intent.lineCap;
  if (intent.lineJoin) changes.lineJoin = intent.lineJoin;
  if (intent.lineStyle) {
    const width = Math.max(0.1, changes.lineWidth ?? own.lineWidth ?? 1);
    if (intent.lineStyle === "solid") {
      changes = { ...changes, dashArray: [], dashPhase: 0 };
    } else if (intent.lineStyle === "dotted") {
      changes = {
        ...changes,
        dashArray: [width, width * 1.5],
        dashPhase: 0,
        lineCap: "round",
      };
    } else {
      changes = {
        ...changes,
        dashArray: [width * 3, width * 2],
        dashPhase: 0,
        lineCap: "butt",
      };
    }
  }
  return withoutUnchanged(changes, target);
}
