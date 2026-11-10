/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { readNativeDrawingGeometry, type NativeDrawingGeometry } from "@mrnote/core/web/lib/native-drawing-geometry.ts";
import { presentationCornerGeometry } from "./presentation-corner-geometry";

const number = (value: number) => String(Number(value.toFixed(7)));
const tokens = (path: string) => path.match(/[MLAZ]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?/g) ?? [];

function framePath(w: number, h: number, inset: number): string {
  const a = number(inset), right = number(w - inset), bottom = number(h - inset);
  return `M0 0 L${number(w)} 0 L${number(w)} ${number(h)} L0 ${number(h)} Z M${a} ${a} L${a} ${bottom} L${right} ${bottom} L${right} ${a} Z`;
}

type InsetKind = "frame" | "bevel" | "donut";

function donutPath(w: number, h: number, inset: number): string {
  const rx = w / 2, ry = h / 2, ix = rx - inset, iy = ry - inset;
  const arc = (x: number, y: number, sweep: number, ex: number, ey: number) => `A${number(x)} ${number(y)} 0 0 ${sweep} ${number(ex)} ${number(ey)}`;
  const outer = `M0 ${number(ry)} ${arc(rx, ry, 1, rx, 0)} ${arc(rx, ry, 1, w, ry)} ${arc(rx, ry, 1, rx, h)} ${arc(rx, ry, 1, 0, ry)} Z`;
  // DrawingML omits zero-radius arcs at the half-short-side boundary.
  const inner = ix > 0 && iy > 0 ? ` ${arc(ix, iy, 0, rx, h - inset)} ${arc(ix, iy, 0, w - inset, ry)} ${arc(ix, iy, 0, rx, inset)} ${arc(ix, iy, 0, inset, ry)}` : "";
  return `${outer} M${number(inset)} ${number(ry)}${inner} Z`;
}

function insetPaths(kind: InsetKind, w: number, h: number, inset: number): NativeDrawingGeometry["paths"] {
  if (kind === "frame") return [{ d: framePath(w, h, inset), fill: "norm", stroke: true }];
  if (kind === "donut") return [{ d: donutPath(w, h, inset), fill: "norm", stroke: true }];
  const a = number(inset), right = number(w - inset), bottom = number(h - inset);
  const width = number(w), height = number(h);
  const outer = `M0 0 L${width} 0 L${width} ${height} L0 ${height} Z`;
  const inner = `M${a} ${a} L${right} ${a} L${right} ${bottom} L${a} ${bottom} Z`;
  return [
    { d: inner, fill: "norm", stroke: false },
    { d: `M0 0 L${width} 0 L${right} ${a} L${a} ${a} Z`, fill: "lightenLess", stroke: false },
    { d: `M0 ${height} L${a} ${bottom} L${right} ${bottom} L${width} ${height} Z`, fill: "darkenLess", stroke: false },
    { d: `M0 0 L${a} ${a} L${a} ${bottom} L0 ${height} Z`, fill: "lighten", stroke: false },
    { d: `M${width} 0 L${width} ${height} L${right} ${bottom} L${right} ${a} Z`, fill: "darken", stroke: false },
    { d: `${outer} ${inner} M0 0 L${a} ${a} M0 ${height} L${a} ${bottom} M${width} 0 L${right} ${a} M${width} ${height} L${right} ${bottom}`, fill: "none", stroke: true },
  ];
}

function samePath(source: NativeDrawingGeometry["paths"][number], expected: NativeDrawingGeometry["paths"][number]): boolean {
  if (source.fill !== expected.fill || source.stroke !== expected.stroke) return false;
  const a = tokens(source.d), b = tokens(expected.d);
  return a.join("") === source.d.replace(/[\s,]/g, "") && a.length === b.length && a.every((value, i) =>
    Number.isNaN(Number(value)) || Number.isNaN(Number(b[i])) ? value === b[i] :
      Math.abs(Number(value) - Number(b[i])) <= 0.00001);
}

/** Frame, bevel and donut share the public DrawingML inset program:
 * min(w,h) * pin(0,adjustment,50000) / 100000. Recover the effective authored
 * ratio only if every original command AND paint role matches its preset.
 * Bevel keeps its five shaded faces and separate stroked edge path. */
function resizeNativeInset(kind: InsetKind, width: number, height: number, drawing?: Record<string, unknown>): NativeDrawingGeometry | null {
  if (drawing?.schemaVersion !== "pptx-shape-drawing-v1" || drawing.shapeKind !== kind ||
    drawing.customPathIdentity || drawing.connectorIdentity ||
    [width, height].some(value => !Number.isFinite(value) || value <= 0 || value > 100_000)) return null;
  const geometry = readNativeDrawingGeometry(drawing.geometry);
  if (!geometry?.viewBox) return null;
  const [w, h] = geometry.viewBox;
  const inset = Number(tokens(geometry.paths[0].d)[kind === "frame" ? 14 : kind === "donut" ? 37 : 1]);
  if (!Number.isFinite(inset) || inset < 0 || inset > Math.min(w, h) / 2) return null;
  const expected = insetPaths(kind, w, h, inset);
  if (geometry.paths.length !== expected.length || geometry.paths.some((path, index) => !samePath(path, expected[index]))) return null;
  if (width === w && height === h) return geometry;
  return { ...geometry, viewBox: [width, height], paths:
    insetPaths(kind, width, height, inset / Math.min(w, h) * Math.min(width, height)),
  };
}

const ARROW_ORDERS = {
  "right-arrow": [0, 1, 2, 3, 4, 5, 6],
  "left-arrow": [3, 2, 1, 0, 6, 5, 4],
  "up-arrow": [2, 3, 4, 5, 6, 0, 1],
  "down-arrow": [2, 1, 0, 6, 5, 4, 3],
} as const;
type SingleArrowKind = keyof typeof ARROW_ORDERS;

type DoubleArrowKind = "left-right-arrow" | "up-down-arrow";

function isSingleArrowKind(kind: string): kind is SingleArrowKind {
  return Object.prototype.hasOwnProperty.call(ARROW_ORDERS, kind);
}

export function isShortSideArrowKind(kind: string): kind is SingleArrowKind | DoubleArrowKind {
  return isSingleArrowKind(kind) || kind === "left-right-arrow" || kind === "up-down-arrow";
}

/** Orthogonal transforms from a rightward arrow's along/cross coordinates. */
function arrowTransform(kind: SingleArrowKind, w: number, h: number): readonly [number, number, number, number, number, number] {
  if (kind === "left-arrow") return [-1, 0, 0, 1, w, 0];
  if (kind === "up-arrow") return [0, -1, 1, 0, 0, h];
  if (kind === "down-arrow") return [0, 1, 1, 0, 0, 0];
  return [1, 0, 0, 1, 0, 0];
}

function singleArrowPath(kind: SingleArrowKind, w: number, h: number, shaftRatio: number, headRatio: number): string {
  const vertical = kind === "up-arrow" || kind === "down-arrow";
  const along = vertical ? h : w, cross = vertical ? w : h;
  const shoulder = along - Math.min(along, Math.min(w, h) * headRatio);
  const low = cross * (1 - shaftRatio) / 2, high = cross * (1 + shaftRatio) / 2;
  const points = [[0, low], [shoulder, low], [shoulder, 0], [along, cross / 2], [shoulder, cross], [shoulder, high], [0, high]];
  const [a, b, c, d, tx, ty] = arrowTransform(kind, w, h);
  return ARROW_ORDERS[kind].map((point, index) => {
    const [u, v] = points[point];
    return `${index === 0 ? "M" : "L"}${number(a * u + c * v + tx)} ${number(b * u + d * v + ty)}`;
  }).join(" ") + " Z";
}

function arrowAdjustmentValues(kind: SingleArrowKind | DoubleArrowKind, geometry: NativeDrawingGeometry, drawing: Record<string, unknown>): { adj1: number; adj2: number } | null | undefined {
  if (!Object.prototype.hasOwnProperty.call(drawing, "presetAdjustmentValues")) return undefined;
  const raw = drawing.presetAdjustmentValues;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const receipt = raw as Record<string, unknown>;
  const preset = { "right-arrow": "rightArrow", "left-arrow": "leftArrow", "up-arrow": "upArrow", "down-arrow": "downArrow", "left-right-arrow": "leftRightArrow", "up-down-arrow": "upDownArrow" }[kind];
  if (receipt.schemaVersion !== "pptx-preset-adjustment-values-v1" || receipt.preset !== preset ||
    Object.keys(receipt).some(key => !["schemaVersion", "preset", "sourceViewBox", "values"].includes(key)) ||
    !Array.isArray(receipt.sourceViewBox) || receipt.sourceViewBox.length !== 2 ||
    receipt.sourceViewBox.some((value, index) => value !== geometry.viewBox?.[index])) return null;
  const values = receipt.values;
  if (!values || typeof values !== "object" || Array.isArray(values)) return null;
  const record = values as Record<string, unknown>;
  if (Object.keys(record).length !== 2 || !["adj1", "adj2"].every(key => typeof record[key] === "number" && Number.isFinite(record[key]) && Math.abs(record[key] as number) <= 1_000_000_000)) return null;
  return { adj1: record.adj1 as number, adj2: record.adj2 as number };
}

/** Shaft thickness uses the cross dimension; head length uses the short side.
 * Recover both in the same canonical coordinates for all four single arrows.
 * An optional source-derived constant-value receipt preserves unclamped adj2.
 * Without it, a fully clamped source keeps the existing safe fallback. */
function resizeNativeSingleArrow(kind: SingleArrowKind, width: number, height: number, drawing?: Record<string, unknown>): NativeDrawingGeometry | null {
  if (drawing?.schemaVersion !== "pptx-shape-drawing-v1" || drawing.shapeKind !== kind ||
    drawing.customPathIdentity || drawing.connectorIdentity ||
    [width, height].some(value => !Number.isFinite(value) || value <= 0 || value > 100_000)) return null;
  const geometry = readNativeDrawingGeometry(drawing.geometry);
  if (!geometry?.viewBox || geometry.paths.length !== 1) return null;
  const [w, h] = geometry.viewBox, source = tokens(geometry.paths[0].d);
  if (source.length !== 22 || source[21] !== "Z" || ARROW_ORDERS[kind].some((_, index) => source[index * 3] !== (index === 0 ? "M" : "L"))) return null;
  const [a, b, c, d, tx, ty] = arrowTransform(kind, w, h);
  const canonical: Array<[number, number]> = [];
  ARROW_ORDERS[kind].forEach((point, index) => {
    const x = Number(source[index * 3 + 1]) - tx, y = Number(source[index * 3 + 2]) - ty;
    canonical[point] = [a * x + b * y, c * x + d * y];
  });
  const vertical = kind === "up-arrow" || kind === "down-arrow";
  const along = vertical ? h : w, cross = vertical ? w : h;
  const top = canonical[0][1], head = along - canonical[1][0];
  const values = arrowAdjustmentValues(kind, geometry, drawing);
  if (values === null || !Number.isFinite(top) || top < 0 || top > cross / 2 || !Number.isFinite(head) || head < 0 || head > along ||
    (values === undefined && head >= along - 0.00001)) return null;
  const shaftRatio = values ? Math.max(0, Math.min(100000, values.adj1)) / 100000 : 1 - top * 2 / cross;
  const headRatio = values ? Math.max(0, values.adj2) / 100000 : head / Math.min(w, h);
  const path = { d: singleArrowPath(kind, w, h, shaftRatio, headRatio), fill: "norm" as const, stroke: true };
  if (!samePath(geometry.paths[0], path)) return null;
  if (width === w && height === h) return geometry;
  return { ...geometry, viewBox: [width, height], paths: [{ ...path, d: singleArrowPath(kind, width, height, shaftRatio, headRatio) }] };
}

function doubleArrowPath(kind: DoubleArrowKind, w: number, h: number, shaftRatio: number, headRatio: number): string {
  const vertical = kind === "up-down-arrow", along = vertical ? h : w, cross = vertical ? w : h;
  const head = Math.min(along / 2, Math.min(w, h) * headRatio), other = along - head;
  const low = cross * (1 - shaftRatio) / 2, high = cross * (1 + shaftRatio) / 2;
  const points = [[0, cross / 2], [head, 0], [head, low], [other, low], [other, 0], [along, cross / 2], [other, cross], [other, high], [head, high], [head, cross]];
  const order = vertical ? [1, 0, 9, 8, 7, 6, 5, 4, 3, 2] : [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
  const [a, b, c, d, tx, ty] = arrowTransform(vertical ? "down-arrow" : "right-arrow", w, h);
  return order.map((point, index) => {
    const [u, v] = points[point];
    return `${index === 0 ? "M" : "L"}${number(a * u + c * v + tx)} ${number(b * u + d * v + ty)}`;
  }).join(" ") + " Z";
}

/** Two headed block arrows share a ten-point program. Each short-side head
 * can consume at most half the along dimension. Saturated source heads cannot
 * reveal the unclamped adjustment without the optional source-value receipt. */
function resizeNativeDoubleArrow(kind: DoubleArrowKind, width: number, height: number, drawing?: Record<string, unknown>): NativeDrawingGeometry | null {
  if (drawing?.schemaVersion !== "pptx-shape-drawing-v1" || drawing.shapeKind !== kind ||
    drawing.customPathIdentity || drawing.connectorIdentity ||
    [width, height].some(value => !Number.isFinite(value) || value <= 0 || value > 100_000)) return null;
  const geometry = readNativeDrawingGeometry(drawing.geometry);
  if (!geometry?.viewBox || geometry.paths.length !== 1) return null;
  const [w, h] = geometry.viewBox, source = tokens(geometry.paths[0].d);
  const vertical = kind === "up-down-arrow", along = vertical ? h : w, cross = vertical ? w : h;
  const head = Number(source[vertical ? 2 : 4]), low = Number(source[vertical ? 28 : 8]);
  const values = arrowAdjustmentValues(kind, geometry, drawing);
  if (values === null || !Number.isFinite(head) || head < 0 || head > along / 2 + 0.00001 || !Number.isFinite(low) || low < 0 || low > cross / 2 + 0.00001 ||
    (values === undefined && head >= along / 2 - 0.00001)) return null;
  const shaftRatio = values ? Math.max(0, Math.min(100000, values.adj1)) / 100000 : Math.max(0, 1 - low * 2 / cross);
  const headRatio = values ? Math.max(0, values.adj2) / 100000 : head / Math.min(w, h);
  const path = { d: doubleArrowPath(kind, w, h, shaftRatio, headRatio), fill: "norm" as const, stroke: true };
  if (!samePath(geometry.paths[0], path)) return null;
  if (width === w && height === h) return geometry;
  return { ...geometry, viewBox: [width, height], paths: [{ ...path, d: doubleArrowPath(kind, width, height, shaftRatio, headRatio) }] };
}

/** Narrow adapter for preset dimensions driven by the shorter side. Unknown
 * native/custom programs retain their imported drawing; existing corner
 * behavior remains in the shared corner implementation. */
export function presentationShortSideGeometry(
  kind: string, width: number, height: number, nativeDrawing?: Record<string, unknown>,
): NativeDrawingGeometry | null {
  if (kind === "frame" || kind === "bevel" || kind === "donut") return resizeNativeInset(kind, width, height, nativeDrawing);
  if (kind === "left-right-arrow" || kind === "up-down-arrow") return resizeNativeDoubleArrow(kind, width, height, nativeDrawing);
  if (isSingleArrowKind(kind)) return resizeNativeSingleArrow(kind, width, height, nativeDrawing);
  return presentationCornerGeometry(kind, width, height,
    nativeDrawing?.shapeKind === kind ? nativeDrawing.geometry : undefined);
}
