/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PaintedTargetBounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

function overlap(a: PaintedTargetBounds, b: PaintedTargetBounds): boolean {
  return (
    Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > 1e-9 &&
    Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > 1e-9
  );
}

/** The filled vectors painted over a picture they overlap, in paint order. */
export function filledVectorsOverImages<
  V extends {
    readonly rendererOperatorIndex: number;
    readonly paintMode: string;
    readonly bounds: PaintedTargetBounds;
  },
>(
  vectors: readonly V[],
  images: ReadonlyArray<{
    readonly operatorIndex?: number;
    readonly bounds: PaintedTargetBounds;
  }>,
): V[] {
  const placed = images.filter(
    (image): image is { operatorIndex: number; bounds: PaintedTargetBounds } =>
      typeof image.operatorIndex === "number",
  );
  if (placed.length === 0) return [];
  return vectors
    .filter(
      (vector) =>
        (vector.paintMode === "fill" || vector.paintMode === "fill-stroke") &&
        placed.some(
          (image) =>
            image.operatorIndex < vector.rendererOperatorIndex &&
            overlap(image.bounds, vector.bounds),
        ),
    )
    .sort((a, b) => a.rendererOperatorIndex - b.rendererOperatorIndex);
}

/** At most this many click shields on one page: a chart page draws
 * thousands of paths, and each shield is one element. */
const MAX_UNROUTED_SHIELDS = 500;












export function unroutedVectorsInFrontOfTargets<
  I extends {
    readonly operatorIndex: number;
    readonly paintMode: string;
    readonly kind: string;
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
    readonly clippedBounds?: PaintedTargetBounds;
  },
>(
  items: readonly I[] | undefined,
  targets: ReadonlyArray<{
    readonly rendererOperatorIndex: number;
    readonly bounds: PaintedTargetBounds;
  }>,
): Array<Readonly<{ operatorIndex: number; bounds: PaintedTargetBounds }>> {
  if (!items || items.length === 0 || targets.length === 0) return [];
  const routed = new Set(targets.map((target) => target.rendererOperatorIndex));
  const shields: Array<Readonly<{ operatorIndex: number; bounds: PaintedTargetBounds }>> = [];
  for (const item of items) {
    if (routed.has(item.operatorIndex)) continue;
    if (item.paintMode !== "fill" && item.paintMode !== "fill-stroke") continue;
    if (!item.clippedBounds && item.kind !== "rect-like") continue;
    const bounds: PaintedTargetBounds = item.clippedBounds ?? {
      x: item.x,
      y: item.y,
      width: item.width,
      height: item.height,
    };
    if (!(bounds.width > 0) || !(bounds.height > 0)) continue;
    if (
      !targets.some(
        (target) =>
          target.rendererOperatorIndex < item.operatorIndex &&
          overlap(target.bounds, bounds),
      )
    ) {
      continue;
    }
    shields.push(Object.freeze({ operatorIndex: item.operatorIndex, bounds }));
    if (shields.length >= MAX_UNROUTED_SHIELDS) break;
  }
  return shields;
}
