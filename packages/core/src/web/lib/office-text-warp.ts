/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The ONE preset text-warp model the three editors share (browser half).
 *
 * A text warp is DrawingML's `a:bodyPr/a:prstTxWarp` (ECMA-376 §20.1.9.19,
 * `ST_TextShapeType` §20.1.10.76) and VML's `v:textpath` on a WordArt
 * shapetype: one enumeration, spelled twice. Every member is a program of
 * guide formulas and paths (`presetTextWarpDefinitions.xml`), transcribed once
 * into `office-text-warp.shared.json` — byte-identical to the server's copy —
 * and evaluated with the same DrawingML guide semantics the preset shapes use.
 *
 * Three drawing kinds, measured against Word (docs/office-graphics/TEXT-WARP.md):
 *
 * - `none` (`textNoShape`): no warp;
 * - `fit` / `envelope`: the paths come in (top, bottom) pairs — one band per
 *   pair. The tight ink rectangle of a band's lines maps into it: a point at
 *   `(u, v)` lands at `T(u) + v·(B(u) − T(u))`, `T`/`B` by ARC LENGTH. `fit` is
 *   `textPlain`, whose paths are straight (a Word watermark);
 * - `follow` (`textArchUp/Down`, `textCircle`, `textButton`): each path is a
 *   baseline; glyphs keep their shapes, rotated onto it, the whole line scaled
 *   about the box centre by `R / (R + h)`.
 *
 * The server twin is `apps/api/app/services/office_text_warp.py`; both suites
 * pin the table and the Word measurements in `office-text-warp.word-measured.json`.
 */

import table from "./office-text-warp.shared.json";

export type TextWarpDrawing = "none" | "fit" | "envelope" | "follow";
export type TextWarpPoint = readonly [number, number];
export type TextWarpPath = readonly TextWarpPoint[];

type Atom = number | { readonly ref: string };
interface Guide { readonly name: string; readonly op: string; readonly args: readonly Atom[] }
interface Command { readonly op: string; readonly args: readonly Atom[] }
interface AdjustSpec { readonly default: number; readonly min: number; readonly max: number }

export interface TextWarpPreset {
  readonly token: string;
  readonly drawing: TextWarpDrawing;
  /** The MSOSPT ordinal VML names the same warp by, when it has one. */
  readonly spt: number | null;
  /** The `a:avLst` guide names with their defaults and handle ranges. */
  readonly adjustments: Readonly<Record<string, AdjustSpec>>;
}

interface Row {
  token: string;
  drawing: TextWarpDrawing;
  spt?: number | null;
  adjustments?: Record<string, AdjustSpec>;
  guides?: Guide[];
  paths?: Command[][];
}

const ROWS = (table.presets as unknown as Row[]);

export const TEXT_WARP_PRESETS: readonly TextWarpPreset[] = ROWS.map((row) => ({
  token: row.token,
  drawing: row.drawing,
  spt: row.spt ?? null,
  adjustments: row.adjustments ?? {},
}));

const BY_TOKEN = new Map(ROWS.map((row) => [row.token, row]));
const BY_SPT = new Map(
  ROWS.flatMap((row) => (row.spt == null ? [] : [[row.spt, row] as const])),
);

/** Before the follow-path scale, a glyph's baseline sits this many ems inside
 * the path (Word, Arial; TEXT-WARP.md §4.2.5). */
export const TEXT_WARP_FOLLOW_BASELINE_EM = 0.18;
const ANGLE_UNIT = Math.PI / 10_800_000;
const MAX_ADJUST = 2_147_483_647;

/** One `ST_TextShapeType` token, or null for anything else. */
export function textWarpToken(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const token = value.trim();
  return BY_TOKEN.has(token) ? token : null;
}

/** How to draw this warp, or null for a token outside the enumeration. */
export function textWarpDrawing(value: unknown): TextWarpDrawing | null {
  const token = textWarpToken(value);
  return token === null ? null : (BY_TOKEN.get(token)?.drawing ?? null);
}

/** The warp a VML WordArt shapetype ordinal (`o:spt`) names. */
export function vmlTextWarpToken(spt: unknown): string | null {
  if (typeof spt !== "number" || !Number.isInteger(spt)) return null;
  return BY_SPT.get(spt)?.token ?? null;
}

/**
 * The `a:avLst` guides a warp states, validated against its preset: only the
 * preset's own names, Int32 values; a guide equal to the default is dropped.
 * Null when nothing differs or when the value is not a valid set.
 */
export function textWarpAdjustments(
  token: unknown,
  value: unknown,
): Readonly<Record<string, number>> | null {
  const row = BY_TOKEN.get(textWarpToken(token) ?? "");
  if (!row || !value || typeof value !== "object" || Array.isArray(value)) return null;
  const declared = row.adjustments ?? {};
  const result: Record<string, number> = {};
  for (const [name, amount] of Object.entries(value as Record<string, unknown>)) {
    const spec = Object.prototype.hasOwnProperty.call(declared, name) ? declared[name] : null;
    if (!spec || typeof amount !== "number" || !Number.isSafeInteger(amount) || Math.abs(amount) > MAX_ADJUST) {
      return null;
    }
    if (amount !== spec.default) result[name] = amount;
  }
  return Object.keys(result).length ? result : null;
}

/** The persisted warp: one `ST_TextShapeType` token and its stated guides. */
export interface OfficeTextWarp {
  readonly preset: string;
  readonly adjustments?: Readonly<Record<string, number>>;
}

/**
 * The ONE page shape of a warp (`{preset, adjustments?}`), or null for no
 * warp — `textNoShape`, an unknown token, or guides the preset does not
 * declare (the server refuses the same things, `office_text_warp.text_warp_value`).
 */
export function normalizeOfficeTextWarp(value: unknown): OfficeTextWarp | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const token = textWarpToken(raw.preset);
  if (!token || token !== raw.preset || token === "textNoShape") return null;
  if (Object.keys(raw).some((key) => key !== "preset" && key !== "adjustments")) return null;
  if (raw.adjustments == null) return { preset: token };
  const adjustments = textWarpAdjustments(token, raw.adjustments);
  if (adjustments === null) {
    // Every stated guide equal to its default is the default warp; anything
    // the preset does not declare is no warp at all.
    const declared = BY_TOKEN.get(token)?.adjustments ?? {};
    const entries = typeof raw.adjustments === "object" ? Object.entries(raw.adjustments as object) : [];
    const allDefault = entries.every(([name, amount]) =>
      Object.prototype.hasOwnProperty.call(declared, name) && declared[name].default === amount);
    return allDefault ? { preset: token } : null;
  }
  return { preset: token, adjustments };
}

/** Two warps say the same thing. */
export function sameOfficeTextWarp(a: unknown, b: unknown): boolean {
  return JSON.stringify(normalizeOfficeTextWarp(a)) === JSON.stringify(normalizeOfficeTextWarp(b));
}

// ------------------------------------------------------------------ guides

function guideVariables(width: number, height: number): Record<string, number> {
  const ss = Math.min(width, height);
  const values: Record<string, number> = Object.assign(Object.create(null), {
    l: 0, t: 0, r: width, b: height, w: width, h: height, hc: width / 2, vc: height / 2,
    ss, ls: Math.max(width, height), cd2: 10_800_000, cd4: 5_400_000, cd8: 2_700_000,
    "3cd4": 16_200_000, "3cd8": 8_100_000, "5cd8": 13_500_000, "7cd8": 18_900_000,
  });
  for (const divisor of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    values[`wd${divisor}`] = width / divisor;
    values[`hd${divisor}`] = height / divisor;
    values[`ssd${divisor}`] = ss / divisor;
  }
  return values;
}

function atom(value: Atom, vars: Record<string, number>): number {
  if (typeof value === "number") return value;
  const resolved = vars[value.ref];
  if (typeof resolved !== "number") throw new Error(`unbound text warp guide ${value.ref}`);
  return resolved;
}

function evaluate(op: string, [x = 0, y = 0, z = 0]: number[]): number {
  switch (op) {
    case "val": return x;
    case "abs": return Math.abs(x);
    case "sqrt": return Math.sqrt(Math.max(0, x));
    case "max": return Math.max(x, y);
    case "min": return Math.min(x, y);
    case "at2": return Math.atan2(y, x) / ANGLE_UNIT;
    case "cos": return x * Math.cos(y * ANGLE_UNIT);
    case "sin": return x * Math.sin(y * ANGLE_UNIT);
    case "tan": return x * Math.tan(y * ANGLE_UNIT);
    case "*/": return z ? (x * y) / z : 0;
    case "+-": return x + y - z;
    case "+/": return z ? (x + y) / z : 0;
    case "?:": return x > 0 ? y : z;
    case "pin": return Math.min(Math.max(y, x), z);
    case "mod": return Math.hypot(x, y, z);
    case "cat2": return x * Math.cos(Math.atan2(z, y));
    case "sat2": return x * Math.sin(Math.atan2(z, y));
    default: throw new Error(`unsupported text warp formula ${op}`);
  }
}

function parameter(rx: number, ry: number, angle: number): number {
  // DrawingML's arc angle is the polar angle of a ray through the ellipse,
  // not its parameter angle (the distinction shows on a stretched box).
  return Math.atan2(rx * Math.sin(angle), ry * Math.cos(angle));
}

/**
 * The warp's paths as polylines in the box `(0, 0)–(width, height)`, curves
 * flattened at `segments` points per Bézier and per quarter turn of an arc.
 * Null for `textNoShape`, an unknown token or a degenerate box.
 */
export function textWarpPaths(
  token: unknown,
  width: number,
  height: number,
  adjustments?: Readonly<Record<string, number>> | null,
  segments = 32,
): TextWarpPath[] | null {
  const row = BY_TOKEN.get(textWarpToken(token) ?? "");
  if (!row?.paths?.length || !(width > 0) || !(height > 0)) return null;
  const vars = guideVariables(width, height);
  for (const guide of row.guides ?? []) {
    const override = adjustments && Object.prototype.hasOwnProperty.call(adjustments, guide.name)
      ? adjustments[guide.name]
      : undefined;
    vars[guide.name] = override !== undefined && guide.op === "val"
      ? override
      : evaluate(guide.op, guide.args.map((arg) => atom(arg, vars)));
  }
  const paths: TextWarpPath[] = [];
  for (const commands of row.paths) {
    const points: TextWarpPoint[] = [];
    let x = 0;
    let y = 0;
    for (const command of commands) {
      const a = command.args.map((arg) => atom(arg, vars));
      if (command.op === "moveTo" || command.op === "lnTo") {
        [x, y] = a;
        points.push([x, y]);
      } else if (command.op === "cubicBezTo") {
        for (let i = 1; i <= segments; i += 1) {
          const t = i / segments;
          const m = 1 - t;
          points.push([
            m ** 3 * x + 3 * m * m * t * a[0] + 3 * m * t * t * a[2] + t ** 3 * a[4],
            m ** 3 * y + 3 * m * m * t * a[1] + 3 * m * t * t * a[3] + t ** 3 * a[5],
          ]);
        }
        [x, y] = [a[4], a[5]];
      } else if (command.op === "quadBezTo") {
        for (let i = 1; i <= segments; i += 1) {
          const t = i / segments;
          const m = 1 - t;
          points.push([
            m * m * x + 2 * m * t * a[0] + t * t * a[2],
            m * m * y + 2 * m * t * a[1] + t * t * a[3],
          ]);
        }
        [x, y] = [a[2], a[3]];
      } else if (command.op === "arcTo") {
        const [rx, ry] = a;
        const start = a[2] * ANGLE_UNIT;
        const sweep = a[3] * ANGLE_UNIT;
        if (rx <= 0 || ry <= 0 || sweep === 0) continue;
        const t0 = parameter(rx, ry, start);
        const cx = x - rx * Math.cos(t0);
        const cy = y - ry * Math.sin(t0);
        let delta = parameter(rx, ry, start + sweep) - t0;
        // Keep the sweep's own turn and direction: the end point's parameter
        // is only known modulo a full turn.
        while (delta - sweep > Math.PI) delta -= 2 * Math.PI;
        while (sweep - delta > Math.PI) delta += 2 * Math.PI;
        const steps = Math.max(2, Math.ceil((Math.abs(delta) / (Math.PI / 2)) * segments));
        for (let i = 1; i <= steps; i += 1) {
          const t = t0 + (delta * i) / steps;
          points.push([cx + rx * Math.cos(t), cy + ry * Math.sin(t)]);
        }
        [x, y] = points[points.length - 1];
      } else {
        return null;
      }
    }
    paths.push(points);
  }
  return paths;
}

export function polylineLength(points: TextWarpPath): number {
  let total = 0;
  for (let i = 0; i + 1 < points.length; i += 1) {
    total += Math.hypot(points[i + 1][0] - points[i][0], points[i + 1][1] - points[i][1]);
  }
  return total;
}

/**
 * The point at `fraction` of a polyline's ARC LENGTH and its heading (radians).
 * Fractions outside 0..1 continue along the end segments.
 */
export function polylinePoint(points: TextWarpPath, fraction: number): [number, number, number] {
  if (points.length < 2) {
    const [x, y] = points[0] ?? [0, 0];
    return [x, y, 0];
  }
  const segments: [TextWarpPoint, TextWarpPoint, number][] = [];
  for (let i = 0; i + 1 < points.length; i += 1) {
    const length = Math.hypot(points[i + 1][0] - points[i][0], points[i + 1][1] - points[i][1]);
    if (length > 0) segments.push([points[i], points[i + 1], length]);
  }
  if (!segments.length) segments.push([points[0], points[1], 0]);
  const total = segments.reduce((sum, [, , length]) => sum + length, 0) || 1;
  const target = fraction * total;
  let walked = 0;
  for (let index = 0; index < segments.length; index += 1) {
    const [[x0, y0], [x1, y1], length] = segments[index];
    if (walked + length >= target || index === segments.length - 1) {
      let t = length > 0 ? (target - walked) / length : 0;
      if (index > 0) t = Math.max(t, 0);
      return [x0 + (x1 - x0) * t, y0 + (y1 - y0) * t, Math.atan2(y1 - y0, x1 - x0)];
    }
    walked += length;
  }
  /* c8 ignore next */
  return [points[points.length - 1][0], points[points.length - 1][1], 0];
}

/** An envelope warp's (top, bottom) path pairs; null for any other kind. */
export function textWarpBands(
  token: unknown,
  width: number,
  height: number,
  adjustments?: Readonly<Record<string, number>> | null,
  segments = 32,
): [TextWarpPath, TextWarpPath][] | null {
  const drawing = textWarpDrawing(token);
  if (drawing !== "fit" && drawing !== "envelope") return null;
  const paths = textWarpPaths(token, width, height, adjustments, segments);
  if (!paths?.length || paths.length % 2) return null;
  const bands: [TextWarpPath, TextWarpPath][] = [];
  for (let i = 0; i < paths.length; i += 2) bands.push([paths[i], paths[i + 1]]);
  return bands;
}

/** Where the point at `(u, v)` of a band's source rectangle lands. */
export function textWarpEnvelopePoint(
  top: TextWarpPath,
  bottom: TextWarpPath,
  u: number,
  v: number,
): [number, number] {
  const [tx, ty] = polylinePoint(top, u);
  const [bx, by] = polylinePoint(bottom, u);
  return [tx + v * (bx - tx), ty + v * (by - ty)];
}

/**
 * How many lines each band (or follow path) holds, in order. Word fills the
 * earlier groups first: 1 line over 2 bands is [1, 0], 3 over 2 is [2, 1].
 */
export function textWarpGroupSizes(lineCount: number, groupCount: number): number[] {
  if (groupCount <= 0) return [];
  const lines = Math.max(0, Math.floor(lineCount));
  const base = Math.floor(lines / groupCount);
  const extra = lines % groupCount;
  return Array.from({ length: groupCount }, (_, index) => base + (index < extra ? 1 : 0));
}

/** Word's follow-path scale: `R / (R + h)`, `R` half the box's LONGER side,
 * `h` the line's tight ink height at its natural size (TEXT-WARP.md §4.2.4). */
export function textWarpFollowScale(width: number, height: number, inkHeight: number): number {
  const radius = Math.max(width, height) / 2;
  return radius > 0 ? radius / (radius + Math.max(0, inkHeight)) : 1;
}

export interface TextWarpFollowLine {
  /** Glyph advance widths in order, at natural size. */
  readonly advances: readonly number[];
  /** The font size the advances were measured at. */
  readonly size: number;
  /** The line's tight ink height at natural size. */
  readonly inkHeight: number;
  readonly lineHeight?: number;
  readonly align?: "left" | "center" | "right" | "justify";
}

export interface TextWarpGlyphPlacement {
  /** The glyph's BASELINE ORIGIN (left end of its advance), box coordinates. */
  readonly x: number;
  readonly y: number;
  /** Rotation in radians. */
  readonly angle: number;
  /** The uniform scale to draw the glyph at. */
  readonly scale: number;
}

function followPath(points: TextWarpPath): TextWarpPath {
  // A path walked right to left would set the words upside down; Word reads
  // it from its left end (textArchDown). TEXT-WARP.md §4.2.3.
  if (points.length >= 2 && points[points.length - 1][0] < points[0][0] - 1e-9) {
    return [...points].reverse();
  }
  return points;
}

/** Where every glyph of a follow-path warp is drawn, line by line. */
export function textWarpFollowGlyphs(
  token: unknown,
  width: number,
  height: number,
  lines: readonly TextWarpFollowLine[],
  adjustments?: Readonly<Record<string, number>> | null,
): TextWarpGlyphPlacement[][] | null {
  if (textWarpDrawing(token) !== "follow") return null;
  const paths = textWarpPaths(token, width, height, adjustments, 48);
  if (!paths?.length) return null;
  const cx = width / 2;
  const cy = height / 2;
  const sizes = textWarpGroupSizes(lines.length, paths.length);
  const result: TextWarpGlyphPlacement[][] = [];
  let cursor = 0;
  paths.forEach((raw, pathIndex) => {
    const group = lines.slice(cursor, cursor + sizes[pathIndex]);
    cursor += sizes[pathIndex];
    if (!group.length) return;
    const path = followPath(raw);
    const length = polylineLength(path) || 1;
    const scale = textWarpFollowScale(width, height, Math.max(...group.map((line) => line.inkHeight || 0)));
    let depth = 0;
    for (const line of group) {
      const advances = line.advances.map((value) => Math.max(0, value));
      const textWidth = advances.reduce((sum, value) => sum + value, 0);
      const fit = textWidth > 0 ? Math.min(1, length / textWidth) : 1;
      const span = textWidth * fit;
      const start = line.align === "left" ? 0 : line.align === "right" ? length - span : (length - span) / 2;
      const offset = TEXT_WARP_FOLLOW_BASELINE_EM * line.size + depth;
      const placed: TextWarpGlyphPlacement[] = [];
      let walked = 0;
      for (const advance of advances) {
        const middle = start + (walked + advance / 2) * fit;
        const [px, py, heading] = polylinePoint(path, middle / length);
        const nx = Math.sin(heading);
        const ny = -Math.cos(heading);
        const half = (advance * fit) / 2;
        const ox = px - nx * offset - Math.cos(heading) * half;
        const oy = py - ny * offset - Math.sin(heading) * half;
        placed.push({ x: cx + scale * (ox - cx), y: cy + scale * (oy - cy), angle: heading, scale: scale * fit });
        walked += advance;
      }
      result.push(placed);
      depth += line.lineHeight ?? line.size * 1.2;
    }
  });
  return result;
}
