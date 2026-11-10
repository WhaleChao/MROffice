/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How far a converted SmartArt lies from the drawing Office saved for it.
 *
 * Opening a file converts its diagrams into MRAI's editable SmartArt, laid out
 * again from the layout's own definition. Where the author dragged or resized
 * shapes by hand in ways the engine does not reproduce, the converted diagram
 * would no longer look like the one the author saw in PowerPoint, Word or
 * Excel. The drawing part is Office's own picture of the diagram, so the open
 * compares the two and keeps the diagram as Office drew it when they differ
 * (the author can still choose 转换为可编辑).
 *
 * The two sides name their shapes differently (Office's presentation points
 * against MRAI's items), so the comparison is order-free: each drawn box is
 * paired with the nearest converted box, closest pairs first, every box used
 * once, and the error of a pair is the sum of its four edges' offsets over the
 * frame's width — the measure the PowerPoint parity sets use.
 */

export interface DrawingBox {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  /** Where the shape actually paints inside the box, when that is known to differ. */
  readonly painted?: { readonly x: number; readonly y: number; readonly w: number; readonly h: number };
}

/** The mean edge error per paired shape, as a fraction of the frame width, and the shapes left unpaired. */
export interface DrawingLikeness {
  readonly mean: number;
  readonly worst: number;
  readonly unpaired: number;
}








export const SMARTART_DRAWING_LIKENESS_LIMIT = 0.06;










export const SMARTART_OWN_DEFINITION_LIKENESS_LIMIT = 0.02;

function boxCost(a: Omit<DrawingBox, "painted">, b: Omit<DrawingBox, "painted">, width: number): number {
  return (Math.abs(a.x - b.x) + Math.abs(a.y - b.y) + Math.abs(a.w - b.w) + Math.abs(a.h - b.h)) / width;
}

// A shape matches by its box or by where it paints, whichever is closer: a
// conversion may state either for the same arc (cycle5's segments tight,
// VerticalCurvedList's curve in its whole circle's box, as Office does).
function pairCost(a: DrawingBox, b: DrawingBox, width: number): number {
  return Math.min(
    boxCost(a, b, width),
    a.painted ? boxCost(a.painted, b, width) : Infinity,
    b.painted ? boxCost(a, b.painted, width) : Infinity,
  );
}

/** Compare two sets of boxes in one coordinate space. `null` when either side has nothing to compare. */
export function smartArtDrawingLikeness(
  drawn: readonly DrawingBox[],
  converted: readonly DrawingBox[],
  frameWidth: number,
): DrawingLikeness | null {
  if (!(frameWidth > 0) || drawn.length === 0 || converted.length === 0) return null;
  const pairs: { cost: number; a: number; b: number }[] = [];
  drawn.forEach((a, i) => converted.forEach((b, j) => pairs.push({ cost: pairCost(a, b, frameWidth), a: i, b: j })));
  pairs.sort((p, q) => p.cost - q.cost);
  const usedA = new Set<number>();
  const usedB = new Set<number>();
  const costs: number[] = [];
  for (const pair of pairs) {
    if (usedA.has(pair.a) || usedB.has(pair.b)) continue;
    usedA.add(pair.a);
    usedB.add(pair.b);
    costs.push(pair.cost);
  }
  const mean = costs.reduce((sum, cost) => sum + cost, 0) / costs.length;
  return { mean, worst: Math.max(...costs), unpaired: drawn.length + converted.length - 2 * costs.length };
}

/** Whether a converted diagram still looks like the drawing Office saved (or there is nothing to compare). */
export function smartArtConversionLooksLikeDrawing(
  drawn: readonly DrawingBox[],
  converted: readonly DrawingBox[],
  frameWidth: number,
): boolean {
  const likeness = smartArtDrawingLikeness(drawn, converted, frameWidth);
  return likeness === null || likeness.mean <= SMARTART_DRAWING_LIKENESS_LIMIT;
}

const unpainted = (value: unknown): boolean =>
  value === undefined || value === null || value === "none" || value === "transparent";

/**
 * The drawn shapes of a diagram's blocks (a slide's or a sheet's): every
 * `shape` with a size that paints something. Words, pictures and the
 * invisible member that states a diagram's frame are no shape of its
 * drawing.
 */
export function smartArtBlockDrawingBoxes(blocks: readonly unknown[]): DrawingBox[] {
  const boxes: DrawingBox[] = [];
  for (const value of blocks) {
    if (!value || typeof value !== "object") continue;
    const block = value as Record<string, unknown>;
    // A slide or sheet block says `type`; a native member a Word host holds
    // says `role`.
    if (block.type !== "shape" && block.role !== "shape") continue;
    const x = Number(block.left);
    const y = Number(block.top);
    const w = Number(block.width);
    const h = Number(block.height);
    if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) continue;
    if (unpainted(block.fill) && unpainted(block.stroke) && !block.fillGradient) continue;
    const painted = paintedBox(block, { x, y, w, h });
    boxes.push(painted === null ? { x, y, w, h } : { x, y, w, h, painted });
  }
  return boxes;
}









function paintedBox(block: Record<string, unknown>, box: DrawingBox): DrawingBox | null {
  const source = block.source as Record<string, unknown> | undefined;
  const view = source?.nativeSmartArtView as Record<string, unknown> | undefined;
  const geometry = view?.geometry as { paths?: unknown; viewBox?: unknown } | undefined;
  const viewBox = geometry?.viewBox;
  if (!Array.isArray(geometry?.paths) || !Array.isArray(viewBox) || !(Number(viewBox[0]) > 0) || !(Number(viewBox[1]) > 0)) {
    return null;
  }
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const path of geometry.paths as { d?: unknown }[]) {
    if (typeof path?.d !== "string") continue;
    for (const [x, y] of pathPoints(path.d)) {
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
  }
  if (!Number.isFinite(minX) || maxX - minX <= 0 || maxY - minY <= 0) return null;
  const sx = box.w / Number(viewBox[0]);
  const sy = box.h / Number(viewBox[1]);
  return { x: box.x + minX * sx, y: box.y + minY * sy, w: (maxX - minX) * sx, h: (maxY - minY) * sy };
}

/** The points an SVG path's commands pass through (absolute M/L/H/V/C/S/Q/T/A). */
function pathPoints(d: string): [number, number][] {
  const tokens = d.match(/[MLHVCSQTAZmlhvcsqtaz]|-?\d*\.?\d+(?:e[-+]?\d+)?/g) ?? [];
  const arity: Record<string, number> = { M: 2, L: 2, H: 1, V: 1, C: 6, S: 4, Q: 4, T: 2, A: 7, Z: 0 };
  const points: [number, number][] = [];
  let command = "";
  let x = 0;
  let y = 0;
  for (let i = 0; i < tokens.length; ) {
    if (/[A-Za-z]/.test(tokens[i])) {
      command = tokens[i].toUpperCase();
      i += 1;
      if (command === "Z") continue;
    }
    const n = arity[command];
    if (!n) return points;
    const args = tokens.slice(i, i + n).map(Number);
    if (args.length < n || args.some((v) => !Number.isFinite(v))) return points;
    i += n;
    if (command === "H") x = args[0];
    else if (command === "V") y = args[0];
    else if (command === "A") {
      points.push(...arcPoints(x, y, args));
      x = args[5];
      y = args[6];
      continue;
    } else {
      for (let k = 0; k + 1 < n; k += 2) points.push([args[k], args[k + 1]]);
      x = args[n - 2];
      y = args[n - 1];
      continue;
    }
    points.push([x, y]);
  }
  return points;
}

/**
 * The drawn shapes of a Word group's members, back in the frame's EMU: the
 * members stand in the group's child coordinate box, `scale` times the frame.
 */
export function smartArtGroupMemberDrawingBoxes(members: readonly unknown[], scale: number): DrawingBox[] {
  if (!(scale > 0)) return [];
  const boxes: DrawingBox[] = [];
  for (const value of members) {
    if (!value || typeof value !== "object") continue;
    const member = value as Record<string, unknown>;
    if (member.kind === "picture" || member.kind === "chart") continue;
    const x = Number(member.xEmu) / scale;
    const y = Number(member.yEmu) / scale;
    const w = Number(member.widthEmu) / scale;
    const h = Number(member.heightEmu) / scale;
    if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) continue;
    if (unpainted(member.fill) && unpainted(member.outline) && !member.fillGradient) continue;
    boxes.push({ x, y, w, h });
  }
  return boxes;
}

const SHAPE = /<dsp:sp\b[\s\S]*?<\/dsp:sp>/g;
const SHAPE_PROPERTIES = /<dsp:spPr\b[\s\S]*?<\/dsp:spPr>|<dsp:spPr\s*\/>/;
const OFFSET = /<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/;
const EXTENT = /<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/;

/**
 * The drawn shapes of a diagram's drawing part (`dsp:drawing`, the picture
 * Office saved), in the frame's EMU. A shape paints when its properties name
 * a fill or a line other than none, or leave them to the style's `fillRef` /
 * `lnRef` with a non-zero index; the words-only boxes Office draws for text
 * paint neither.
 */
export function smartArtDrawingPartBoxes(drawingXml: string): DrawingBox[] {
  const boxes: DrawingBox[] = [];
  for (const shape of drawingXml.match(SHAPE) ?? []) {
    const properties = shape.match(SHAPE_PROPERTIES)?.[0] ?? "";
    const offset = properties.match(OFFSET);
    const extent = properties.match(EXTENT);
    if (!offset || !extent) continue;
    const [x, y, w, h] = [offset[1], offset[2], extent[1], extent[2]].map(Number);
    if (w <= 0 || h <= 0) continue;
    const line = properties.match(/<a:ln\b[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>/)?.[0] ?? "";
    const outside = properties.replace(line, "");
    const fill = /<a:(solidFill|gradFill|blipFill|pattFill)\b/.test(outside)
      ? true
      : /<a:noFill\s*\/>/.test(outside)
        ? false
        : !/<a:fillRef\s+idx="0"/.test(shape);
    const stroke = /<a:(solidFill|gradFill|pattFill)\b/.test(line)
      ? true
      : /<a:noFill\s*\/>/.test(line)
        ? false
        : !/<a:lnRef\s+idx="0"/.test(shape);
    if (!fill && !stroke) continue;
    boxes.push({ x, y, w, h });
  }
  return boxes;
}

/**
 * Points along an SVG elliptical arc from (x0, y0) (`rx ry rotation large
 * sweep x y`), so a bulging arc's extent counts (VerticalCurvedList's long
 * curve, cycle3's segments): the endpoint-to-centre conversion of SVG 1.1
 * appendix F.6.5, sampled.
 */
function arcPoints(x0: number, y0: number, [rxIn, ryIn, rotation, large, sweep, x1, y1]: number[]): [number, number][] {
  let rx = Math.abs(rxIn);
  let ry = Math.abs(ryIn);
  if (!rx || !ry || (x0 === x1 && y0 === y1)) return [[x1, y1]];
  const phi = (rotation * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx = (x0 - x1) / 2;
  const dy = (y0 - y1) / 2;
  const px = cos * dx + sin * dy;
  const py = -sin * dx + cos * dy;
  const lambda = (px * px) / (rx * rx) + (py * py) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const sign = (large ? 1 : 0) === (sweep ? 1 : 0) ? -1 : 1;
  const numerator = rx * rx * ry * ry - rx * rx * py * py - ry * ry * px * px;
  const factor = sign * Math.sqrt(Math.max(0, numerator / (rx * rx * py * py + ry * ry * px * px)));
  const cxp = (factor * rx * py) / ry;
  const cyp = (-factor * ry * px) / rx;
  const cx = cos * cxp - sin * cyp + (x0 + x1) / 2;
  const cy = sin * cxp + cos * cyp + (y0 + y1) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const start = angle(1, 0, (px - cxp) / rx, (py - cyp) / ry);
  let delta = angle((px - cxp) / rx, (py - cyp) / ry, (-px - cxp) / rx, (-py - cyp) / ry);
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  const points: [number, number][] = [];
  const steps = 24;
  for (let i = 1; i <= steps; i += 1) {
    const t = start + (delta * i) / steps;
    const ex = rx * Math.cos(t);
    const ey = ry * Math.sin(t);
    points.push([cos * ex - sin * ey + cx, sin * ex + cos * ey + cy]);
  }
  return points;
}
